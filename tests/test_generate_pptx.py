"""Build a deck from the fixture spec and read it back with python-pptx."""

import hashlib
import json
import math
import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest
from pptx import Presentation

import generate_pptx

HERE = Path(__file__).resolve().parent
SCRIPT = HERE.parent / "skills" / "slides-from-text" / "scripts" / "generate_pptx.py"
FIXTURE = HERE / "fixtures" / "spec.json"


def load_fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def slide_titles(pptx_path: Path) -> list[str]:
    """First non-empty text on each slide: the deck title, then every headline."""
    titles = []
    for slide in Presentation(str(pptx_path)).slides:
        texts = [
            shape.text_frame.text for shape in slide.shapes if shape.has_text_frame and shape.text_frame.text.strip()
        ]
        titles.append(texts[0])
    return titles


def test_build_round_trips_slide_count_and_titles(tmp_path: Path) -> None:
    spec = load_fixture()
    out = tmp_path / "deck.pptx"
    generate_pptx.build(spec, str(out))
    expected = [spec["title"], *(s["headline"] for s in spec["slides"])]
    assert slide_titles(out) == expected


def test_table_slide_round_trips_headers_and_rows(tmp_path: Path) -> None:
    spec = load_fixture()
    out = tmp_path / "deck.pptx"
    generate_pptx.build(spec, str(out))
    position = next(i for i, s in enumerate(spec["slides"]) if s["type"] == "table")
    table_spec = spec["slides"][position]
    slide = Presentation(str(out)).slides[position + 1]
    table = next(shape.table for shape in slide.shapes if shape.has_table)
    cells = [[table.cell(r, c).text for c in range(len(table.columns))] for r in range(len(table.rows))]
    assert cells == [table_spec["headers"], *table_spec["rows"]]


def test_cli_writes_the_deck(tmp_path: Path) -> None:
    out = tmp_path / "deck.pptx"
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--input", str(FIXTURE), "--output", str(out)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == f"saved: {out}"
    assert len(slide_titles(out)) == 1 + len(load_fixture()["slides"])


def test_body_over_the_line_budget_warns_on_stderr(tmp_path: Path, capsys) -> None:
    spec = {"title": "t", "slides": [{"headline": "h", "body": ["x"] * 11}]}
    generate_pptx.build(spec, str(tmp_path / "deck.pptx"))
    assert "warning: slide 1: 11 body lines, 10 fit cleanly" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("spec", "message"),
    [
        ({"title": "t", "slides": [{"headline": "h", "body": ["x " * 60] * 10}]}, "body needs about"),
        (
            {"title": "t", "slides": [{"type": "two_col", "headline": "h", "left": ["x"] * 20}]},
            "left column needs about",
        ),
        (
            {
                "title": "t",
                "slides": [{"type": "table", "headline": "h", "headers": ["a", "b"], "rows": [["y " * 45, "z"]] * 10}],
            },
            "table needs about",
        ),
        ({"title": "t", "slides": [{"headline": "h" * 121}]}, "headline is 121 characters"),
    ],
)
def test_text_past_the_slide_bounds_warns_but_still_builds(tmp_path: Path, capsys, spec: dict, message: str) -> None:
    out = tmp_path / "deck.pptx"
    generate_pptx.build(spec, str(out))
    assert message in capsys.readouterr().err
    assert out.exists()


@pytest.mark.parametrize(
    "slide",
    [
        {"type": "two_col", "headline": "h", "left": ["The initial limit is still within the initial limit."] * 10},
        {"headline": "h", "body": ["The initial limit is still within the initial limit. " * 2] * 10},
        {"type": "two_col", "headline": "h", "left": [("release readiness " * 6).strip()] * 5},
    ],
    ids=["column-of-narrow-lines", "body-of-narrow-lines", "column-of-wrapped-paragraphs"],
)
def test_text_that_fits_does_not_warn(tmp_path: Path, capsys, slide: dict) -> None:
    generate_pptx.build({"title": "t", "slides": [slide]}, str(tmp_path / "deck.pptx"))
    assert capsys.readouterr().err == ""


def test_each_table_row_gets_the_height_its_cells_need(tmp_path: Path, capsys) -> None:
    rows = [["wide " * 40, "z"]] + [["x", "z"]] * 9
    slide = {"type": "table", "headline": "h", "headers": ["a", "b"], "rows": rows}
    out = tmp_path / "deck.pptx"
    generate_pptx.build({"title": "t", "slides": [slide]}, str(out))
    table = next(shape.table for shape in Presentation(str(out)).slides[1].shapes if shape.has_table)
    heights = [row.height.inches for row in table.rows]
    assert heights[1] >= 1.0 > heights[2] == heights[-1]
    assert capsys.readouterr().err == ""


def test_uneven_rows_that_pass_the_slide_bottom_warn(tmp_path: Path, capsys) -> None:
    rows = [["wide " * 40, "z"]] * 3 + [["x", "z"]] * 7
    slide = {"type": "table", "headline": "h", "headers": ["a", "b"], "rows": rows}
    generate_pptx.build({"title": "t", "slides": [slide]}, str(tmp_path / "deck.pptx"))
    assert "table needs about" in capsys.readouterr().err


def test_columns_too_narrow_for_any_text_warn(tmp_path: Path, capsys) -> None:
    slide = {"type": "table", "headline": "h", "headers": ["c"] * 70, "rows": [["x"] * 70]}
    generate_pptx.build({"title": "t", "slides": [slide]}, str(tmp_path / "deck.pptx"))
    assert "table is too narrow to hold text" in capsys.readouterr().err
    assert generate_pptx.text_height(["x"], 12.5 / 70, 16, "Arial") == math.inf


def test_wide_glyphs_that_overflow_warn(tmp_path: Path, capsys) -> None:
    slide = {"type": "two_col", "headline": "h", "left": [("W " * 20).strip()] * 10}
    generate_pptx.build({"title": "t", "slides": [slide]}, str(tmp_path / "deck.pptx"))
    assert "left column needs about" in capsys.readouterr().err


def test_build_records_the_deck_digest_for_the_upload_hook(tmp_path: Path, isolated_cache: Path) -> None:
    out = tmp_path / "deck.pptx"
    generate_pptx.build(load_fixture(), str(out))
    assert (isolated_cache / "slides-from-text" / "decks" / hashlib.sha256(out.read_bytes()).hexdigest()).is_file()


def test_concurrent_builds_all_record_their_digest(tmp_path: Path, isolated_cache: Path) -> None:
    decks = [tmp_path / f"deck-{n}.pptx" for n in range(12)]
    with ThreadPoolExecutor(max_workers=12) as pool:
        list(pool.map(lambda d: generate_pptx.build({"title": d.stem, "slides": []}, str(d)), decks))
    recorded = {m.name for m in (isolated_cache / "slides-from-text" / "decks").iterdir()}
    assert recorded == {hashlib.sha256(d.read_bytes()).hexdigest() for d in decks}


def test_registry_prunes_only_week_old_markers(tmp_path: Path, isolated_cache: Path) -> None:
    decks = isolated_cache / "slides-from-text" / "decks"
    decks.mkdir(parents=True)
    week_ago = time.time() - generate_pptx.MARKER_TTL_SECONDS - 60
    stale, recent = decks / ("a" * 64), decks / ("b" * 64)
    stale.touch()
    recent.touch()
    os.utime(stale, (week_ago, week_ago))
    generate_pptx.build({"title": "t", "slides": []}, str(tmp_path / "deck.pptx"))
    assert not stale.exists()
    assert recent.exists()


def test_a_marker_refreshed_during_pruning_survives(tmp_path: Path, isolated_cache: Path, monkeypatch) -> None:
    """Build A lists an old marker as stale; build B regenerates that deck before A deletes it."""
    decks = isolated_cache / "slides-from-text" / "decks"
    decks.mkdir(parents=True)
    shared = decks / ("c" * 64)
    shared.touch()
    week_ago = time.time() - generate_pptx.MARKER_TTL_SECONDS - 60
    os.utime(shared, (week_ago, week_ago))
    listing = Path.iterdir

    def list_then_let_b_refresh(self):
        entries = list(listing(self))
        if self == decks:
            shared.touch()
        return iter(entries)

    monkeypatch.setattr(Path, "iterdir", list_then_let_b_refresh)
    generate_pptx.build({"title": "t", "slides": []}, str(tmp_path / "deck.pptx"))
    assert shared.exists()


def test_fixture_builds_without_warnings(tmp_path: Path, capsys) -> None:
    generate_pptx.build(load_fixture(), str(tmp_path / "deck.pptx"))
    assert capsys.readouterr().err == ""


def test_cli_reports_a_bad_spec_without_a_traceback(tmp_path: Path) -> None:
    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps({"title": "t", "slides": [{"headline": "h", "body": "one string"}]}), encoding="utf-8")
    out = tmp_path / "deck.pptx"
    result = subprocess.run(
        [sys.executable, str(SCRIPT), "--input", str(spec), "--output", str(out)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 2
    assert result.stderr.startswith("error: slide 1 (content): ")
    assert "Traceback" not in result.stderr
    assert not out.exists()


@pytest.mark.parametrize(
    ("spec", "message"),
    [
        ([], "spec must be a JSON object"),
        ({"slides": []}, 'missing "title"'),
        ({"title": "t", "slides": {"headline": "h"}}, '"slides" must be a list'),
        ({"title": "t", "slides": ["h"]}, "slide 1: must be an object"),
        ({"title": "t", "slides": [{"headline": "h", "body": "one string"}]}, '"body" must be a list of strings'),
        ({"title": "t", "slides": [{"type": "two_col", "headline": "h", "right": "x"}]}, '"right" must be a list'),
        ({"title": "t", "slides": [{"headline": "h", "bullets": ["a"]}]}, "unknown field bullets"),
        ({"title": "t", "slides": [{"headline": 3}]}, '"headline" must be a string'),
        ({"title": "t", "theme": {"bg": "12"}, "slides": []}, 'theme "bg" must be a 6-digit hex color'),
        ({"title": "t", "theme": {"bg": 123456}, "slides": []}, 'theme "bg" must be a 6-digit hex color'),
        ({"title": "t", "theme": {"font": 5}, "slides": []}, 'theme "font" must be a font name'),
        ({"title": "t", "slides": [{"type": [], "headline": "h"}]}, r"unknown slide type: \[\]"),
        (
            {"title": "t", "slides": [{"type": "table", "headline": "h", "headers": [], "rows": []}]},
            '"headers" must name at least one column',
        ),
        (
            {"title": "t", "slides": [{"type": "table", "headline": "h", "headers": "a, b", "rows": [["x", "y"]]}]},
            '"headers" must be a list of strings',
        ),
        (
            {"title": "t", "slides": [{"type": "table", "headline": "h", "headers": ["a"], "rows": ["x"]}]},
            '"rows" must be a list of rows',
        ),
        ({"title": "t", "slides": [{"type": "pie", "headline": "h"}]}, "unknown slide type: pie"),
        ({"title": "t", "slides": [{"type": "table", "headline": "h"}]}, "missing headers, rows"),
        (
            {"title": "t", "slides": [{"type": "table", "headline": "h", "headers": ["a"], "rows": [["x"]] * 11}]},
            "11 rows, max is 10",
        ),
        (
            {"title": "t", "slides": [{"type": "table", "headline": "h", "headers": ["a", "b"], "rows": [["x"]]}]},
            "row 1 has 1 cells, headers has 2",
        ),
    ],
)
def test_bad_spec_is_rejected_before_saving(tmp_path: Path, spec: dict, message: str) -> None:
    out = tmp_path / "deck.pptx"
    with pytest.raises(ValueError, match=message):
        generate_pptx.build(spec, str(out))
    assert not out.exists()
