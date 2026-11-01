#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx"]
# ///
"""Generate a styled PPTX deck from a JSON slide spec.

Usage:
    uv run generate_pptx.py --input slides.json --output deck.pptx

The inline script metadata above declares python-pptx, so `uv run` resolves it
without a venv. The file is executable, so `./generate_pptx.py ...` works too.

JSON schema:
{
  "title": "Deck Title",
  "subtitle": "Optional subtitle line",
  "footer": "Optional small footer text on the title slide",
  "theme": {
    "bg": "1A1A2E", "accent": "00D4AA", "white": "FFFFFF",
    "light_gray": "CCCCCC", "dark_gray": "33334A", "font": "Arial"
  },
  "slides": [
    {"type": "content", "headline": "...", "body": ["line1", "line2", ...],
     "note": "optional italic footer"},
    {"type": "two_col", "headline": "...",
     "left_title": "...", "left": ["...", ...],
     "right_title": "...", "right": ["...", ...]},
    {"type": "table", "headline": "...", "headers": ["Col1","Col2","Col3"],
     "rows": [["a","b","c"], ...]}
  ]
}

Body/left/right lines: use "" for a blank spacer line. Keep each slide's body
to 10 lines max (8 if the slide has a "note") so it stays readable at 20pt in
a 13.33x7.5in deck; the script warns past that. Tables are capped at 10 data
rows; split larger tables across slides. Every table row must have one cell
per header. Headlines shrink automatically past 60 and 90 characters, but keep
them under ~90.

A spec with a wrong type, an unknown slide field or a bad theme color is
rejected before anything is saved: the CLI prints "error: <reason>" and exits 2.
The script also warns on stderr when a headline passes 120 characters or when
wrapped text in a body, column or table would run past the slide edge.

Each saved deck leaves an empty file named by its SHA-256 in
$XDG_CACHE_HOME/slides-from-text/decks (default ~/.cache/...); files untouched for
seven days are pruned. SKILL.md's upload hook refuses a PPTX payload with no
matching file.
"""

import argparse
import hashlib
import json
import math
import os
import re
import sys
import time
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

W = Inches(13.33)
H = Inches(7.5)
MAX_TABLE_ROWS = 10
MAX_HEADLINE_CHARS = 120
MARKER_TTL_SECONDS = 7 * 24 * 3600
# Fields each slide type must carry; the keys are the supported slide types.
REQUIRED_FIELDS = {"content": ["headline"], "two_col": ["headline"], "table": ["headline", "headers", "rows"]}
# Every field a slide type renders; anything else would vanish from the deck without a word.
KNOWN_FIELDS = {
    "content": {"type", "headline", "body", "note"},
    "two_col": {"type", "headline", "left_title", "left", "right_title", "right"},
    "table": {"type", "headline", "headers", "rows"},
}
TEXT_FIELDS = ("headline", "note", "left_title", "right_title")
LINE_FIELDS = ("body", "left", "right", "headers")
THEME_COLORS = ("bg", "accent", "white", "light_gray", "dark_gray")


def hexcolor(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def validate(spec):
    """Reject a spec the renderer would turn into a broken deck, naming the slide and field."""
    if not isinstance(spec, dict):
        raise ValueError("spec must be a JSON object")
    if "title" not in spec:
        raise ValueError('spec is missing "title"')
    for key in ("title", "subtitle", "footer"):
        if key in spec and not isinstance(spec[key], str):
            raise ValueError(f'"{key}" must be a string')
    theme = spec.get("theme", {})
    if not isinstance(theme, dict):
        raise ValueError('"theme" must be an object')
    for key in THEME_COLORS:
        if key in theme and not (isinstance(theme[key], str) and re.fullmatch(r"#?[0-9A-Fa-f]{6}", theme[key])):
            raise ValueError(f'theme "{key}" must be a 6-digit hex color string like "1A1A2E", got {theme[key]!r}')
    if "font" in theme and not (isinstance(theme["font"], str) and theme["font"].strip()):
        raise ValueError(f'theme "font" must be a font name, got {theme["font"]!r}')
    slides = spec.get("slides", [])
    if not isinstance(slides, list):
        raise ValueError('"slides" must be a list of slide objects')
    for i, s in enumerate(slides, start=1):
        if not isinstance(s, dict):
            raise ValueError(f"slide {i}: must be an object")
        kind = s.get("type", "content")
        if not isinstance(kind, str) or kind not in REQUIRED_FIELDS:
            raise ValueError(f"slide {i}: unknown slide type: {kind}")
        missing = [k for k in REQUIRED_FIELDS[kind] if k not in s]
        if missing:
            raise ValueError(f"slide {i} ({kind}): missing {', '.join(missing)}")
        unknown = sorted(set(s) - KNOWN_FIELDS[kind])
        if unknown:
            raise ValueError(
                f"slide {i} ({kind}): unknown field {', '.join(unknown)}; "
                f"{kind} slides take {', '.join(sorted(KNOWN_FIELDS[kind]))}"
            )
        for key in TEXT_FIELDS:
            if key in s and not isinstance(s[key], str):
                raise ValueError(f'slide {i} ({kind}): "{key}" must be a string')
        for key in LINE_FIELDS:
            if key in s and not (isinstance(s[key], list) and all(isinstance(x, str) for x in s[key])):
                raise ValueError(f'slide {i} ({kind}): "{key}" must be a list of strings, one per line')
        if kind == "table":
            if not s["headers"]:
                raise ValueError(f'slide {i} (table): "headers" must name at least one column')
            rows = s["rows"]
            if not isinstance(rows, list) or not all(isinstance(r, list) for r in rows):
                raise ValueError(f'slide {i} (table): "rows" must be a list of rows, each a list of cells')
            if len(rows) > MAX_TABLE_ROWS:
                raise ValueError(
                    f"slide {i}: table has {len(rows)} rows, max is {MAX_TABLE_ROWS}; split it across two table slides"
                )
            for ri, row in enumerate(rows, start=1):
                if len(row) != len(s["headers"]):
                    raise ValueError(f"slide {i}: table row {ri} has {len(row)} cells, headers has {len(s['headers'])}")
                if not all(isinstance(c, (str, int, float)) for c in row):
                    raise ValueError(f"slide {i}: table row {ri} cells must be text or numbers")


# Helvetica/Arial advance widths for ASCII 32-126, in 1/1000 em (Adobe AFM). Other characters use 556.
GLYPH_WIDTHS = dict(
    zip(
        map(chr, range(32, 127)),
        [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278]
        + [556] * 10
        + [278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833]
        + [722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556]
        + [556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556]
        + [500, 722, 500, 500, 500, 334, 260, 334, 584],
        strict=True,
    )
)
# Width relative to Arial for common theme fonts; an unknown font is assumed slightly wider.
FONT_WIDTH_SCALE = {"arial": 1.0, "helvetica": 1.0, "calibri": 0.9, "times new roman": 0.9, "georgia": 1.12}


def wrapped_lines(text, width, font_size, font, bold=False):
    """Lines a word-wrapped paragraph takes in a box `width` inches wide (0.1in inset each side).

    A box with no room left after the insets cannot hold text at all, so it counts as infinite.
    """
    scale = FONT_WIDTH_SCALE.get(font.lower(), 1.1) * (1.07 if bold else 1.0) * font_size / 1000 / 72
    room = width - 0.2
    if room <= 0:
        return math.inf
    lines, used = 1, 0.0
    for word in str(text).split(" "):
        w = sum(GLYPH_WIDTHS.get(ch, 556) for ch in word) * scale
        gap = GLYPH_WIDTHS[" "] * scale if used else 0.0
        if used + gap + w <= room:
            used += gap + w
            continue
        if used:
            lines += 1
        # A word longer than the line breaks across lines on its own.
        lines += int(w // room)
        used = w % room
    return lines


def text_height(lines, width, font_size, font, space_before=0, bold=False):
    """Rough rendered height in inches of word-wrapped paragraphs at line spacing 1.2."""
    wrapped = sum(wrapped_lines(line, width, font_size, font, bold) for line in lines)
    # Paragraph spacing applies once per paragraph, not once per wrapped line.
    return (wrapped * font_size * 1.2 + len(lines) * space_before) / 72


def deck_registry():
    """Folder holding one empty file per generated deck, named by its SHA-256; SKILL.md's upload hook reads it."""
    cache = os.environ.get("XDG_CACHE_HOME") or os.path.join(os.path.expanduser("~"), ".cache")
    return Path(cache) / "slides-from-text" / "decks"


def record_deck(path):
    """Remember the deck's digest so the upload hook can confirm a payload is exactly this file.

    One file per digest means concurrent builds never rewrite a shared file. Markers untouched
    for a week are pruned; the age is read again right before each delete, so a marker another
    build just refreshed survives.
    """
    registry = deck_registry()
    registry.mkdir(parents=True, exist_ok=True)
    (registry / hashlib.sha256(Path(path).read_bytes()).hexdigest()).touch()
    cutoff = time.time() - MARKER_TTL_SECONDS
    for marker in list(registry.iterdir()):
        try:
            if marker.stat().st_mtime < cutoff:
                marker.unlink()
        except FileNotFoundError:  # pruned by a concurrent build
            pass


def warn(idx, message):
    print(f"warning: slide {idx}: {message}", file=sys.stderr)


def warn_if_overflow(idx, where, need, room):
    if math.isinf(need):
        warn(idx, f"{where} is too narrow to hold text; use fewer columns")
    elif need > room:
        warn(idx, f"{where} needs about {need:.1f}in, {room:.1f}in fit; split the slide or shorten lines")


def build(spec, out_path):
    validate(spec)
    theme = spec.get("theme", {})
    DARK_BG = hexcolor(theme.get("bg", "1A1A2E"))
    ACCENT = hexcolor(theme.get("accent", "00D4AA"))
    WHITE = hexcolor(theme.get("white", "FFFFFF"))
    LIGHT_GRAY = hexcolor(theme.get("light_gray", "CCCCCC"))
    DARK_GRAY = hexcolor(theme.get("dark_gray", "33334A"))
    FONT = theme.get("font", "Arial")

    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    def blank_slide():
        return prs.slides.add_slide(prs.slide_layouts[6])

    def fill_bg(slide, color=DARK_BG):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_rect(slide, left, top, width, height, color):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.fill.background()
        return shape

    def add_text(
        slide, text, left, top, width, height, font_size=24, bold=False, color=WHITE, align=PP_ALIGN.LEFT, italic=False
    ):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name = FONT
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.italic = italic
        return box

    def add_bullet_text(slide, lines, left, top, width, height, font_size=20, color=WHITE):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_before = Pt(4)
            run = p.add_run()
            run.text = line
            run.font.name = FONT
            run.font.size = Pt(font_size)
            run.font.color.rgb = color
        return box

    def title_slide(title, subtitle="", footer=""):
        slide = blank_slide()
        fill_bg(slide)
        add_rect(slide, Inches(0), Inches(0), Inches(0.12), H, ACCENT)
        add_rect(slide, Inches(0.12), Inches(6.8), W - Inches(0.12), Inches(0.06), ACCENT)
        add_text(slide, title, Inches(0.5), Inches(2.2), Inches(12), Inches(2.5), font_size=48, bold=True, color=WHITE)
        if subtitle:
            add_text(slide, subtitle, Inches(0.5), Inches(4.9), Inches(11), Inches(0.8), font_size=22, color=ACCENT)
        if footer:
            add_text(slide, footer, Inches(0.5), Inches(6.9), Inches(8), Inches(0.5), font_size=14, color=LIGHT_GRAY)

    def headline_bar(slide, headline, idx):
        # 32pt fits one line in the 1.3in bar; long headlines wrap, so step
        # the font down to keep two wrapped lines inside the bar.
        size = 32 if len(headline) <= 60 else 28 if len(headline) <= 90 else 24
        if len(headline) > MAX_HEADLINE_CHARS:
            warn(
                idx,
                f"headline is {len(headline)} characters; keep it under {MAX_HEADLINE_CHARS} so it stays in the bar",
            )
        add_rect(slide, Inches(0), Inches(0), W, Inches(0.08), ACCENT)
        add_rect(slide, Inches(0), Inches(0.08), W, Inches(1.3), DARK_GRAY)
        add_text(
            slide,
            headline,
            Inches(0.4),
            Inches(0.15),
            Inches(12.5),
            Inches(1.1),
            font_size=size,
            bold=True,
            color=WHITE,
        )

    def content_slide(s, idx):
        slide = blank_slide()
        fill_bg(slide)
        headline_bar(slide, s["headline"], idx)
        body = s.get("body", [])
        has_note = bool(s.get("note"))
        max_lines = 8 if has_note else 10
        if len(body) > max_lines:
            warn(idx, f"{len(body)} body lines, {max_lines} fit cleanly{' with a note' if has_note else ''}")
        # Body starts at 1.6in and must clear the note at 6.9in or the slide edge at 7.5in.
        warn_if_overflow(idx, "body", text_height(body, 12.3, 20, FONT, 4), (6.9 if has_note else 7.5) - 1.6)
        # Note sits at 6.9in; stop the body above it so the two never overlap.
        body_height = Inches(4.9) if has_note else Inches(5.5)
        add_bullet_text(slide, body, Inches(0.5), Inches(1.6), Inches(12.3), body_height, font_size=20)
        if has_note:
            add_text(
                slide,
                s["note"],
                Inches(0.5),
                Inches(6.9),
                Inches(12),
                Inches(0.5),
                font_size=13,
                color=ACCENT,
                italic=True,
            )

    def two_col_slide(s, idx):
        slide = blank_slide()
        fill_bg(slide)
        headline_bar(slide, s["headline"], idx)
        for side, width in (("left", 5.8), ("right", 6.0)):
            # Columns start at 2.2in; past 7.5in the text leaves the slide.
            warn_if_overflow(idx, f"{side} column", text_height(s.get(side, []), width, 19, FONT, 4), 7.5 - 2.2)
        add_rect(slide, Inches(6.6), Inches(1.55), Inches(0.04), Inches(5.6), ACCENT)
        add_text(
            slide,
            s.get("left_title", ""),
            Inches(0.5),
            Inches(1.6),
            Inches(5.8),
            Inches(0.5),
            font_size=18,
            bold=True,
            color=ACCENT,
        )
        add_bullet_text(slide, s.get("left", []), Inches(0.5), Inches(2.2), Inches(5.8), Inches(5.0), font_size=19)
        add_text(
            slide,
            s.get("right_title", ""),
            Inches(6.9),
            Inches(1.6),
            Inches(6.0),
            Inches(0.5),
            font_size=18,
            bold=True,
            color=ACCENT,
        )
        add_bullet_text(slide, s.get("right", []), Inches(6.9), Inches(2.2), Inches(6.0), Inches(5.0), font_size=19)

    def table_slide(s, idx):
        slide = blank_slide()
        fill_bg(slide)
        headline_bar(slide, s["headline"], idx)
        headers = s["headers"]
        rows = s["rows"]
        # Each row gets the height its tallest wrapped cell needs, so the file states the layout
        # every renderer draws. The table starts at 1.55in and the slide ends at 7.5in; renderers
        # pad the header row a little past the estimate, so keep 0.15in clear of the edge.
        col = 12.5 / len(headers)
        heights = [max(text_height([h], col, 17, FONT, bold=True) for h in headers) + 0.1]
        heights += [max(text_height([c], col, 16, FONT) for c in row) + 0.1 for row in rows]
        warn_if_overflow(idx, "table", sum(heights), 7.5 - 1.55 - 0.15)
        if math.isinf(sum(heights)):
            heights = [5.6 / len(heights)] * len(heights)
        table = slide.shapes.add_table(
            len(rows) + 1, len(headers), Inches(0.4), Inches(1.55), Inches(12.5), Inches(sum(heights))
        ).table
        for row, height in zip(table.rows, heights, strict=True):
            row.height = Inches(height)
        for ci, htext in enumerate(headers):
            cell = table.cell(0, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = ACCENT
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = htext
            run.font.name = FONT
            run.font.bold = True
            run.font.size = Pt(17)
            run.font.color.rgb = DARK_BG
        alt = RGBColor(0x22, 0x22, 0x3A)
        for ri, row in enumerate(rows):
            bg = DARK_GRAY if ri % 2 == 0 else alt
            for ci, val in enumerate(row):
                cell = table.cell(ri + 1, ci)
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg
                p = cell.text_frame.paragraphs[0]
                run = p.add_run()
                run.text = str(val)
                run.font.name = FONT
                run.font.size = Pt(16)
                run.font.color.rgb = WHITE

    title_slide(spec["title"], spec.get("subtitle", ""), spec.get("footer", ""))

    dispatch = {"content": content_slide, "two_col": two_col_slide, "table": table_slide}
    for i, s in enumerate(spec.get("slides", []), start=1):
        dispatch[s.get("type", "content")](s, i)

    prs.save(out_path)
    record_deck(out_path)
    return out_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="path to slides JSON spec")
    ap.add_argument("--output", required=True, help="path to write .pptx")
    args = ap.parse_args()

    try:
        with open(args.input, encoding="utf-8") as f:
            spec = json.load(f)
        path = build(spec, args.output)
    except (OSError, ValueError) as e:
        print(f"error: {e}", file=sys.stderr)
        sys.exit(2)
    print(f"saved: {path}")
