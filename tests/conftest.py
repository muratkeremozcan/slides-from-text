"""Put the skill's scripts directory on sys.path, and keep the deck registry out of the real home."""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "skills" / "slides-from-text" / "scripts"))


@pytest.fixture(autouse=True)
def isolated_cache(tmp_path, monkeypatch):
    """generate_pptx records every deck under XDG_CACHE_HOME; point it at a per-test folder."""
    monkeypatch.setenv("XDG_CACHE_HOME", str(tmp_path / "cache"))
    return tmp_path / "cache"
