---
name: slides-from-text
description: Turn raw text (a brainstorm doc, meeting notes, a Slack thread) into a polished Google Slides deck. Use when the user says 'make slides from this', 'turn this into a presentation', 'generate a deck', or hands over unstructured text expecting a presentation back.
hooks:
  PreToolUse:
    - matcher: "mcp__claude_ai_Google_Drive__create_file"
      hooks:
        - type: command
          command: |
            python3 -I -c "
            import base64, hashlib, json, os, sys
            pptx = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            payload = ''
            try:
                args = json.load(sys.stdin).get('tool_input', {})
                if args.get('contentMimeType') != pptx:
                    sys.exit(0)
                payload = args.get('base64Content') or ''
                cache = os.environ.get('XDG_CACHE_HOME') or os.path.join(os.path.expanduser('~'), '.cache')
                digest = hashlib.sha256(base64.b64decode(payload, validate=True)).hexdigest()
                matches = os.path.isfile(os.path.join(cache, 'slides-from-text', 'decks', digest))
            except Exception:
                matches = False
            if not matches:
                print(f'Blocked by slides-from-text: base64Content ({len(payload)} characters) is not byte-identical to a deck generate_pptx.py saved, so the payload was cut short or altered. Nothing was uploaded.', file=sys.stderr)
                sys.exit(2)
            "
---

# Slides From Text

## Overview

Takes raw, unstructured text and produces a real Google Slides presentation: an actual file in the user's Drive they can open and edit in Slides.

The pipeline is text → structured slide JSON → PPTX (python-pptx, styled) → upload to Google Drive (auto-converts PPTX to native Google Slides).

## When the source text isn't already slide-ready

Raw brainstorm dumps (bullet soup, open questions, no clear per-slide grouping) make bad decks directly. If the user has both a raw version and a distilled/pre-structured version of the same content (e.g. one doc marked "brainstorm" and another "v2" or "filtered"), use the structured one. If only raw text exists, do the distillation yourself before writing the JSON spec: one slide = one idea, headline states the point, body is evidence/support for that point, 5-10 lines max per slide.

## Step 1: Structure the content

Write a JSON file (schema documented in `scripts/generate_pptx.py`'s docstring; read it) with:
- `title`, `subtitle`, `footer` for the title slide
- `slides`: array of slide objects, each `type: "content" | "two_col" | "table"`

Slide-writing rules:
- Headline is a statement ("We already ran this. It works."); a category label ("Background") fails.
- Keep headlines under ~90 characters. The script shrinks the font on long headlines, but past ~120 characters even that wraps out of the header bar.
- One slide, one point. If a section has two distinct claims, split it into two slides.
- Body lines: short. Use `""` as a spacer between grouped bullets. Max 10 body lines per slide; 8 if the slide also has a `note`.
- Tables: max 10 data rows per slide. More rows means split into two table slides; the script rejects oversized tables.
- `two_col` is for a real contrast/comparison (before/after, proved/next, option A/B). Don't force it.
- `table` is for genuinely tabular data (action/owner/date).
- Prefer the theme defaults (dark bg, teal accent) unless the user specifies brand colors; pass a `theme` object to override.

Save the JSON to a scratch path outside the project directory.

## Step 2: Generate the PPTX

```
uv run {skill-root}/scripts/generate_pptx.py --input <spec.json> --output <deck.pptx>
```

The script declares python-pptx in its inline script metadata, so `uv run` resolves it; no venv setup needed.

A spec the script cannot render cleanly stops here. It exits 2 with `error: <reason>` on a malformed spec, and prints a `warning:` line on stderr when a slide will run past its bounds. Fix what it names and regenerate until the run prints no error or warning; only then go to Step 3. If the spec was the user's own file, leave it unchanged, fix a copy, and tell the user what you changed.

## Step 3: Upload to Google Drive as native Slides

Encode the deck to a file, never to stdout: `base64 < <deck.pptx> | tr -d '\n' > <deck.b64>`. If `mcp__claude_ai_Google_Drive__create_file` is deferred, load it with ToolSearch query `select:mcp__claude_ai_Google_Drive__create_file`. Read `<deck.b64>` and call the tool from this thread with `title`, the file's contents as `base64Content`, `contentMimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation"` and `disableConversionToGoogleType: false` (so Drive converts it to native Google Slides).

A model copying a long payload into a tool call can cut it short, and a cut payload lands in Drive as a broken file. The `PreToolUse` hook in this skill's frontmatter blocks a PPTX `create_file` call unless `base64Content` decodes to a deck `scripts/generate_pptx.py` saved, byte for byte; the script leaves a file named by each deck's SHA-256 for that check. Skill hooks do not run inside subagents, so the upload stays in this thread, where the hook can check it; the cost is the 50-150K-character payload in this conversation. If the hook blocks the call, retry once from the file. If it blocks again or the tool returns an error, the upload failed.

## Step 4: Report back

If the upload failed, tell the user, quote the hook or tool message, and give the local PPTX path; give no link. Otherwise give the user the Drive file link and a one-line description of what's in the deck. Don't re-describe the content slide-by-slide unless asked.

## Reuse

The canonical home is `skills/slides-from-text/` in the `slides-from-text` repository. Symlink that directory into `~/.claude/skills/` and it loads in every Claude Code session. Steps 1-2 (spec and PPTX generation) work anywhere; step 3 needs the Google Drive MCP connected in the calling session. The JSON-spec/script split means re-running with new content is just: write a new spec, regenerate, upload.
