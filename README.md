# slides-from-text

A `slides-from-text` skill that turns raw text (a brainstorm doc, meeting notes, a Slack
thread) into a native Google Slides deck. The agent distills the text into a JSON slide spec,
`generate_pptx.py` renders it as a styled PPTX, and the agent uploads it to Google Drive,
which converts it to Slides. A hook in the skill blocks any PPTX upload whose payload differs
from the deck the script saved.

- `skills/slides-from-text/SKILL.md`: the skill, with the slide-writing rules.
- `skills/slides-from-text/scripts/generate_pptx.py`: the PPTX generator. Its docstring is the
  JSON spec schema. Inline script metadata declares `python-pptx`, so `uv run` needs no venv.
- `tests/`: round-trip tests for the generator and contract tests for SKILL.md.
- `docs/source-provenance.md`: where each file came from, with source SHA-256 values.

## Requirements

- [uv](https://docs.astral.sh/uv/) to run the generator.
- The Google Drive MCP connector in the calling session for the upload step. Spec writing and
  PPTX generation work without it.

## Install

Clone this repository, then symlink `skills/slides-from-text` into each discovery directory
that should load the skill, such as `~/.claude/skills/`. Linking to the checkout keeps the
installed skill current with every `git pull`.

## Generate a deck by hand

```bash
uv run skills/slides-from-text/scripts/generate_pptx.py --input spec.json --output deck.pptx
```

Each run leaves an empty file named by the deck's SHA-256 in `~/.cache/slides-from-text/decks/`
(under `$XDG_CACHE_HOME` when set); files untouched for seven days are pruned. The skill's
upload hook only lets through a PPTX payload with a matching file there.

## Tests

```bash
uv run pytest
```

The generator tests build a deck from `tests/fixtures/spec.json`, read it back with python-pptx
and check slide titles, table cells, the CLI and its error exit, the line and overflow warnings
and every spec rejection. The contract tests parse the SKILL.md frontmatter as YAML and check
that the paths it names exist, that its slide types equal the script's `REQUIRED_FIELDS` keys
and each one renders, that its table row cap matches the script, that its upload hook passes
only a deck the script recorded, and that the script keeps its inline metadata, `uv` shebang
and executable bit. CI also runs the script through `uv run` the way the skill does.
