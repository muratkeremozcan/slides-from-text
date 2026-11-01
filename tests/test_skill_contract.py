"""Check that SKILL.md stays loadable and agrees with the script it drives."""

import base64
import io
import json
import os
import re
import subprocess
import tomllib
import zipfile
from pathlib import Path

import pytest
import yaml

import generate_pptx

ROOT = Path(__file__).resolve().parents[1]
SKILL_DIR = ROOT / "skills" / "slides-from-text"
SKILL_MD = SKILL_DIR / "SKILL.md"
SCRIPT = SKILL_DIR / "scripts" / "generate_pptx.py"

BACKTICKED = re.compile(r"`([^`\n]+)`")
PATH_SUFFIXES = (".md", ".json", ".py", ".sh", ".txt", ".yaml", ".yml", ".toml")


def split_frontmatter() -> tuple[dict, str]:
    match = re.match(r"^---\n(.*?)\n---\n(.*)$", SKILL_MD.read_text(encoding="utf-8"), re.DOTALL)
    assert match, "SKILL.md has no YAML frontmatter block"
    meta = yaml.safe_load(match.group(1))
    assert isinstance(meta, dict)
    return meta, match.group(2)


def inline_script_metadata() -> dict:
    """The PEP 723 block that lets `uv run` resolve the script's dependencies."""
    text = SCRIPT.read_text(encoding="utf-8")
    match = re.search(r"^# /// script\n((?:#.*\n)+?)# ///$", text, re.MULTILINE)
    assert match, "generate_pptx.py has no inline script metadata block"
    body = "".join(line[2:] if line.startswith("# ") else line[1:] for line in match.group(1).splitlines(True))
    return tomllib.loads(body)


def test_frontmatter_name_and_description() -> None:
    meta, body = split_frontmatter()
    assert meta.get("name") == SKILL_DIR.name
    assert re.match(r"^[a-z0-9]+(-[a-z0-9]+)*$", meta["name"]) and len(meta["name"]) <= 64
    assert isinstance(meta.get("description"), str) and meta["description"].strip()
    assert len(meta["description"]) <= 1024
    assert body.strip()


def test_relative_paths_the_skill_names_exist() -> None:
    _, body = split_frontmatter()
    paths = [t for t in BACKTICKED.findall(body) if t.endswith(PATH_SUFFIXES) and " " not in t]
    assert "scripts/generate_pptx.py" in paths
    for path in paths:
        assert not path.startswith(("/", "~")), f"SKILL.md names an absolute path: {path}"
        assert (SKILL_DIR / path).is_file(), f"SKILL.md names missing path {path}"


def test_slide_types_in_skill_match_the_script(tmp_path: Path) -> None:
    _, body = split_frontmatter()
    match = re.search(r'`type: "([^`]+)"`', body)
    assert match, "SKILL.md no longer lists the slide types"
    documented = {t.strip().strip('"') for t in match.group(1).split("|")}
    assert documented == set(generate_pptx.REQUIRED_FIELDS)
    for kind, fields in generate_pptx.REQUIRED_FIELDS.items():
        assert f'"type": "{kind}"' in generate_pptx.__doc__, f"script docstring omits slide type {kind}"
        values = {"headline": "h", "headers": ["a"], "rows": [["x"]]}
        slide = {"type": kind, **{field: values[field] for field in fields}}
        out = tmp_path / f"{kind}.pptx"
        generate_pptx.build({"title": "t", "slides": [slide]}, str(out))
        assert out.is_file(), f"the script does not render documented slide type {kind}"


def upload_hook_command() -> str:
    meta, _ = split_frontmatter()
    (entry,) = meta["hooks"]["PreToolUse"]
    assert entry["matcher"] == "mcp__claude_ai_Google_Drive__create_file"
    (hook,) = entry["hooks"]
    assert hook["type"] == "command"
    return hook["command"]


PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def repackaged(data: bytes, drop: str = "", corrupt: str = "") -> bytes:
    """The deck rezipped without member `drop`, or with the first compressed byte of `corrupt` set to 0xff."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename != drop:
                dst.writestr(item, src.read(item.filename))
    if not corrupt:
        return out.getvalue()
    raw = bytearray(out.getvalue())
    with zipfile.ZipFile(io.BytesIO(bytes(raw))) as z:
        info = z.getinfo(corrupt)
    start = info.header_offset + 30 + len(info.filename.encode()) + len(info.extra)
    raw[start] = 0xFF
    return bytes(raw)


def run_upload_hook(payload: str, mime: str = PPTX) -> subprocess.CompletedProcess:
    event = {
        "tool_name": "mcp__claude_ai_Google_Drive__create_file",
        "tool_input": {"title": "t", "base64Content": payload, "contentMimeType": mime},
    }
    return subprocess.run(
        ["/bin/sh", "-c", upload_hook_command()], input=json.dumps(event), capture_output=True, text=True, check=False
    )


@pytest.mark.parametrize(
    ("damage", "exit_code"),
    [
        (lambda d: d, 0),
        (lambda d: d[: len(d) // 2], 2),
        (lambda d: d[:2000] + bytes([d[2000] ^ 1]) + d[2001:], 2),
        (lambda d: b"", 2),
        (lambda d: repackaged(d, corrupt="ppt/presentation.xml"), 2),
        (lambda d: repackaged(d, drop="_rels/.rels"), 2),
    ],
    ids=["generated", "cut", "one-byte-changed", "empty", "corrupt-deflate", "missing-rels"],
)
def test_upload_hook_passes_only_the_generated_deck(tmp_path: Path, damage, exit_code: int) -> None:
    deck = tmp_path / "deck.pptx"
    generate_pptx.build({"title": "t", "slides": [{"headline": "h", "body": ["x"]}]}, str(deck))
    result = run_upload_hook(base64.b64encode(damage(deck.read_bytes())).decode())
    assert result.returncode == exit_code, result.stderr
    if exit_code:
        assert "Nothing was uploaded" in result.stderr


def test_upload_hook_blocks_a_deck_the_script_did_not_record(tmp_path: Path, isolated_cache: Path) -> None:
    deck = tmp_path / "deck.pptx"
    generate_pptx.build({"title": "t", "slides": []}, str(deck))
    for marker in (isolated_cache / "slides-from-text" / "decks").iterdir():
        marker.unlink()
    assert run_upload_hook(base64.b64encode(deck.read_bytes()).decode()).returncode == 2


def test_upload_hook_ignores_other_uploads() -> None:
    assert run_upload_hook(base64.b64encode(b"plain notes").decode(), mime="text/plain").returncode == 0


def test_table_row_cap_in_skill_matches_the_script() -> None:
    _, body = split_frontmatter()
    assert f"max {generate_pptx.MAX_TABLE_ROWS} data rows per slide" in body


def test_inline_metadata_declares_python_pptx() -> None:
    meta = inline_script_metadata()
    assert any(dep.split()[0].startswith("python-pptx") for dep in meta["dependencies"])
    assert meta["requires-python"]


def test_script_is_executable_with_a_uv_shebang() -> None:
    assert SCRIPT.read_text(encoding="utf-8").startswith("#!/usr/bin/env -S uv run --script\n")
    assert os.access(SCRIPT, os.X_OK)
    mode = subprocess.run(
        ["git", "ls-files", "--stage", str(SCRIPT.relative_to(ROOT))],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.split()
    assert mode and mode[0] == "100755", "generate_pptx.py must be tracked as executable"


def test_no_ds_store_is_tracked() -> None:
    tracked = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True).stdout
    assert not [line for line in tracked.splitlines() if Path(line).name == ".DS_Store"]
