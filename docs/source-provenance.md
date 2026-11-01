# Source provenance

This repository was extracted on 2026-10-31 from a personal assistant repository at
`~/personal-agent`, HEAD `841f83ae36cc8adbd69c6e79ffd4b39dfe356c9d`. The last commit that
touched these sources there is `28a519f16476c0a0ae16f61c46f7d075df47318c`. The source
checkout and its discovery links were not modified.

| Source path (under `skills/slides-from-text/`) | SHA-256 | Became |
|---|---|---|
| `SKILL.md` | `dd632d7b01f1fd2d951602c474466e62bf4280e56b729218245ecc03963612b6` | `skills/slides-from-text/SKILL.md` |
| `scripts/generate_pptx.py` | `458dfc383102d698dad287f941c81b5db35b881b3477d0bf209084466ae96108` | `skills/slides-from-text/scripts/generate_pptx.py` |
| `scripts/tests/test_generate_pptx.py` | `5f1538388e4bda11912d76afa14b8919c7b4820e50c7b3a2e5f2708773083290` | `tests/test_generate_pptx.py` |
| `scripts/tests/conftest.py` | `5fe16b6ee33295f0e4d53b8d3c2cba37e88f64b890a7750c4065156f8c48c897` | `tests/conftest.py` |
| `scripts/tests/fixtures/spec.json` | `3d3f86b824836479ff4f0cc3eff34adf40132944ef004c17dfb088a07446aec2` | `tests/fixtures/spec.json` |

## What changed

- `SKILL.md`: the Reuse section names this repository as the canonical home and says how to
  install it.
- `generate_pptx.py`: reformatted with `ruff format`, then the per-type required fields moved
  from a local inside `build()` to the module constant `REQUIRED_FIELDS` so tests can compare
  SKILL.md against it. Rendering and every error message are unchanged. The executable bit is
  kept.
- Tests moved from inside the skill to `tests/` at the repository root, so an installed skill
  carries only what an agent runs. `conftest.py` has a new path and a reworded docstring.
  `test_generate_pptx.py` has the new script path, imports re-sorted by ruff, and two
  comprehensions and two parametrized specs joined onto single lines by `ruff format`.
- `tests/fixtures/spec.json`: every slide's text, numbers, names and dates were replaced with
  synthetic values. The shape is kept: a content slide with a note and non-ASCII arrows, a
  `two_col` slide whose headline crosses the 60 character step-down, and a two-row table. No
  assertion depends on the old values.
- `README.md` was rewritten from the init commit's stub.

New files: `tests/test_skill_contract.py`, CI, `pyproject.toml` and `ruff.toml`.

## License

MIT. See `LICENSE`.
