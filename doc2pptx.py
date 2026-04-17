#!/usr/bin/env python3
"""
doc2pptx — Convert documents (txt, md, docx, pdf) into PowerPoint presentations.

Optionally accepts an existing .pptx as a style template to inherit backgrounds,
color themes, fonts, and slide layouts.

Usage:
    python doc2pptx.py input.md -o output.pptx
    python doc2pptx.py report.pdf -o output.pptx --template brand.pptx
    python doc2pptx.py notes.txt -o output.pptx --template brand.pptx --title "My Deck"
"""

from __future__ import annotations

import argparse
import copy
import logging
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
# 1. DATA MODEL — Intermediate representation of slide content
# ─────────────────────────────────────────────────────────────

@dataclass
class SlideContent:
    """One logical slide."""
    title: str = ""
    body_lines: list[str] = field(default_factory=list)
    level: int = 0  # heading depth (1 = H1/title, 2 = H2/section, etc.)
    slide_type: str = "content"  # "title", "section", "content", "table", "quote", "bignum", "stat_grid"
    # Table data (used when slide_type == "table")
    table_headers: list[str] = field(default_factory=list)
    table_rows: list[list[str]] = field(default_factory=list)
    # LLM-suggested layout hint from a {layout=... accent=...} attribute on
    # the heading line. Untrusted — the renderer validates and snaps to an
    # available layout, falling back silently if the hint is unknown.
    layout_hint: str | None = None
    accent_hint: int | None = None
    # Creative content blocks parsed from the LLM output.
    quote_text: str | None = None
    big_number: tuple[str, str] | None = None  # (value, label)
    stats: list[tuple[str, str]] = field(default_factory=list)


@dataclass
class DeckContent:
    """Full parsed document ready for slide generation."""
    title: str = ""
    subtitle: str = ""
    slides: list[SlideContent] = field(default_factory=list)


# ─────────────────────────────────────────────────────────────
# 2. DOCUMENT PARSERS — Extract structured content from files
# ─────────────────────────────────────────────────────────────

def read_txt(path: str) -> str:
    """Read plain text file."""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def read_markdown(path: str) -> str:
    """Read markdown file (returns raw markdown text)."""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def read_docx(path: str) -> str:
    """Extract text from a .docx file, preserving heading markers."""
    from docx import Document as DocxDocument
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

    doc = DocxDocument(path)
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            lines.append("")
            continue
        style_name = (para.style.name or "").lower()
        # Map Word heading styles to markdown-style markers
        if "heading 1" in style_name:
            lines.append(f"# {text}")
        elif "heading 2" in style_name:
            lines.append(f"## {text}")
        elif "heading 3" in style_name:
            lines.append(f"### {text}")
        elif "heading 4" in style_name:
            lines.append(f"#### {text}")
        elif "title" in style_name:
            lines.append(f"# {text}")
        elif "subtitle" in style_name:
            lines.append(f"_subtitle: {text}")
        elif "list" in style_name:
            lines.append(f"- {text}")
        else:
            lines.append(text)
    return "\n".join(lines)


def read_pdf(path: str) -> str:
    """Extract text from a PDF file."""
    try:
        import pdfplumber
        texts = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    texts.append(page_text)
        if texts:
            return "\n\n".join(texts)
    except Exception:
        pass

    # Fallback to pypdf
    from pypdf import PdfReader
    reader = PdfReader(path)
    texts = []
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            texts.append(page_text)
    return "\n\n".join(texts)


def read_html(path: str) -> str:
    """Extract text from an HTML file, mapping headings to markdown markers."""
    from bs4 import BeautifulSoup

    with open(path, "r", encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    lines = []
    for el in soup.find_all(["h1", "h2", "h3", "h4", "p", "li"]):
        text = el.get_text(strip=True)
        if not text:
            continue
        tag = el.name.lower()
        if tag == "h1":
            lines.append(f"# {text}")
        elif tag == "h2":
            lines.append(f"## {text}")
        elif tag == "h3":
            lines.append(f"### {text}")
        elif tag == "h4":
            lines.append(f"#### {text}")
        elif tag == "li":
            lines.append(f"- {text}")
        else:
            lines.append(text)
        lines.append("")  # blank line after each block element
    return "\n".join(lines)


# ─── Spreadsheet reader (returns DeckContent, not raw text) ──

MAX_TABLE_ROWS_PER_SLIDE = 15   # data rows per table slide (excl. header)
MAX_TABLE_COLS = 10             # columns before we truncate
MAX_COL_CHAR_WIDTH = 30        # truncate long cell values


def _sanitize_cell(val) -> str:
    """Convert a cell value to a clean display string."""
    if val is None:
        return ""
    import math
    if isinstance(val, float):
        if math.isnan(val):
            return ""
        # Show ints cleanly, floats with reasonable precision
        if val == int(val):
            return str(int(val))
        return f"{val:.2f}"
    s = str(val).strip()
    if len(s) > MAX_COL_CHAR_WIDTH:
        s = s[: MAX_COL_CHAR_WIDTH - 1] + "…"
    return s


def read_xlsx(path: str, deck_title: str = "", max_rows_per_slide: int = MAX_TABLE_ROWS_PER_SLIDE) -> DeckContent:
    """
    Read an xlsx/xls/csv/tsv file and produce a DeckContent with table slides.

    Each worksheet becomes one or more table slides. If a sheet has more rows
    than max_rows_per_slide, it is split across multiple slides.
    """
    import pandas as pd

    ext = Path(path).suffix.lower()

    # Read all sheets
    if ext in (".csv", ".tsv"):
        sep = "\t" if ext == ".tsv" else ","
        sheets = {"Sheet1": pd.read_csv(path, sep=sep)}
    else:
        sheets = pd.read_excel(path, sheet_name=None)

    file_stem = Path(path).stem
    deck = DeckContent(title=deck_title or file_stem)

    # Title slide
    deck.slides.append(SlideContent(
        title=deck.title,
        slide_type="title",
    ))

    for sheet_name, df in sheets.items():
        if df.empty:
            continue

        # Truncate columns if too many
        if len(df.columns) > MAX_TABLE_COLS:
            df = df.iloc[:, :MAX_TABLE_COLS]

        headers = [_sanitize_cell(c) for c in df.columns]

        # Convert all data to strings
        all_rows = []
        for _, row in df.iterrows():
            all_rows.append([_sanitize_cell(v) for v in row])

        # Chunk rows into slides
        for chunk_idx in range(0, max(1, len(all_rows)), max_rows_per_slide):
            chunk = all_rows[chunk_idx : chunk_idx + max_rows_per_slide]
            page_num = chunk_idx // max_rows_per_slide + 1
            total_pages = (len(all_rows) + max_rows_per_slide - 1) // max_rows_per_slide

            if total_pages == 1:
                slide_title = sheet_name
            else:
                slide_title = f"{sheet_name} ({page_num}/{total_pages})"

            deck.slides.append(SlideContent(
                title=slide_title,
                slide_type="table",
                table_headers=headers,
                table_rows=chunk,
            ))

    return deck


def read_document(path: str) -> str:
    """Route to the right reader based on file extension."""
    ext = Path(path).suffix.lower()
    readers = {
        ".txt": read_txt,
        ".md": read_markdown,
        ".markdown": read_markdown,
        ".docx": read_docx,
        ".pdf": read_pdf,
        ".html": read_html,
        ".htm": read_html,
    }
    reader = readers.get(ext)
    if reader is None:
        raise ValueError(
            f"Unsupported file type: '{ext}'. "
            f"Supported: {', '.join(readers.keys())}"
        )
    return reader(path)


# ─────────────────────────────────────────────────────────────
# 2b. LLM REWRITE (Ollama) — prose → PPT-friendly markdown
# ─────────────────────────────────────────────────────────────

DEFAULT_OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
DEFAULT_OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2")

DEFAULT_REWRITE_PROMPT = """You are rewriting a document so it converts into a POWERFUL, VARIED PowerPoint deck — not a wall of bullet points.

Output rules:
- Output ONLY markdown. No preamble, no commentary, no code fences.
- Use `#` for the deck title and `##` for each slide title.
- Tag every `##` heading with a layout hint in curly braces AT THE END OF THE LINE:
    ## My slide title {layout=content}
    ## Key metrics {layout=stat_grid accent=1}
  Only use these layout values: title, section, content, two_content, comparison, picture_caption, quote, stat_grid, callout.
  `accent` is an optional integer 1-6 selecting a theme accent color for decoration.
- Under each `##`, write EITHER bullet lines starting with `- `, OR one of the special creative blocks below.

Creative blocks (use them! variety is the goal):
- Pull quote:  a single line starting with `> ` — use `{layout=quote}` on the heading.
- Big number:  `[[BIG: 87% | retention after 30 days]]` on its own line — use `{layout=bignum}` on the heading.
- Stat grid:   several `[[STAT: value | label]]` entries separated by spaces on one line — use `{layout=stat_grid}`.
- Comparison:  two groups of bullets separated by a line `---` — use `{layout=comparison}` or `{layout=two_content}`.

Content rules:
- Preserve the original meaning, facts, numbers, and ordering. Do not invent content.
- Cover everything, but summarize instead of quoting verbatim.
- Drop filler, repetition, and boilerplate.
- Vary the layout across slides. Aim for at least 2 non-"content" layouts per 10 slides when the source material supports it (quotes → quote, statistics → stat_grid / bignum, side-by-side ideas → two_content / comparison).
- A slide that introduces a new major part of the document → `## Title {layout=section}` with no bullets.
"""


def rewrite_for_pptx(
    raw_text: str,
    host: str = DEFAULT_OLLAMA_HOST,
    model: str = DEFAULT_OLLAMA_MODEL,
    system_prompt: str | None = None,
    template_kit: str | None = None,
    timeout: float = 120.0,
) -> str:
    """Ask a local Ollama server to reshape raw_text into PPT-friendly markdown.

    When ``template_kit`` is supplied, it is injected into the system prompt
    so the LLM knows which layouts and accent colors are actually available
    in the uploaded template — steering it toward compliant, varied output.

    Returns the rewritten markdown. Raises on HTTP/connection errors — callers
    decide whether to fall back to the original text.
    """
    import httpx

    base_prompt = system_prompt or DEFAULT_REWRITE_PROMPT
    if template_kit:
        full_prompt = (
            f"{base_prompt}\n\n"
            f"THIS TEMPLATE ADVERTISES:\n{template_kit}\n\n"
            "Only use layout values that appear above. If the template does not "
            "advertise a layout you want, fall back to `content` or `section`."
        )
    else:
        full_prompt = base_prompt

    payload = {
        "model": model,
        "prompt": raw_text,
        "system": full_prompt,
        "stream": False,
        "options": {"temperature": 0.2},
    }
    url = host.rstrip("/") + "/api/generate"
    with httpx.Client(timeout=timeout) as client:
        resp = client.post(url, json=payload)
        resp.raise_for_status()
        data = resp.json()
    return (data.get("response") or "").strip()


# ─────────────────────────────────────────────────────────────
# 3. TEXT → STRUCTURED SLIDES PARSER
# ─────────────────────────────────────────────────────────────

def _is_heading(line: str) -> tuple[int, str] | None:
    """Check if a line is a markdown-style heading. Returns (level, text) or None."""
    m = re.match(r"^(#{1,4})\s+(.+)$", line)
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None


_HEADING_HINT_RE = re.compile(r"\s*\{([^{}]+)\}\s*$")
_HINT_KV_RE = re.compile(r"(\w+)\s*=\s*([^\s]+)")
_BIG_RE = re.compile(r"\[\[BIG:\s*([^|\]]+?)\s*\|\s*([^\]]+?)\s*\]\]", re.IGNORECASE)
_STAT_RE = re.compile(r"\[\[STAT:\s*([^|\]]+?)\s*\|\s*([^\]]+?)\s*\]\]", re.IGNORECASE)


def _extract_heading_hints(text: str) -> tuple[str, str | None, int | None]:
    """Pull a trailing ``{layout=… accent=…}`` attribute block off a heading.

    Returns ``(clean_text, layout_hint, accent_hint)``. Unknown keys are
    ignored silently so the LLM can include extras without breaking parsing.
    """
    m = _HEADING_HINT_RE.search(text)
    if not m:
        return text.strip(), None, None
    inside = m.group(1)
    clean = text[: m.start()].strip()
    layout_hint: str | None = None
    accent_hint: int | None = None
    for key, val in _HINT_KV_RE.findall(inside):
        key_l = key.lower()
        if key_l == "layout":
            layout_hint = val.lower()
        elif key_l == "accent":
            try:
                accent_hint = max(1, min(6, int(val))) - 1  # 1-indexed in prompt, 0-indexed internally
            except ValueError:
                pass
    return clean, layout_hint, accent_hint


def _parse_stats_line(line: str) -> list[tuple[str, str]]:
    """Return a list of (value, label) pairs from one or more [[STAT:]] tokens."""
    return [(m.group(1).strip(), m.group(2).strip()) for m in _STAT_RE.finditer(line)]


def _parse_big_line(line: str) -> tuple[str, str] | None:
    """Return (value, label) from a [[BIG:]] token, or None."""
    m = _BIG_RE.search(line)
    if not m:
        return None
    return m.group(1).strip(), m.group(2).strip()


def _is_bullet(line: str) -> str | None:
    """Check if a line is a bullet point. Returns the text or None."""
    m = re.match(r"^\s*[-*•]\s+(.+)$", line)
    if m:
        return m.group(1).strip()
    # Numbered list
    m = re.match(r"^\s*\d+[.)]\s+(.+)$", line)
    if m:
        return m.group(1).strip()
    return None


def _looks_like_heading(line: str, prev_blank: bool, next_blank: bool) -> bool:
    """
    Heuristic: detect "implicit headings" in plain text.
    A short standalone line surrounded by blank lines that doesn't end with
    a period is likely a heading/section title.
    """
    stripped = line.strip()
    if not stripped:
        return False
    # Must be relatively short
    if len(stripped) > 80:
        return False
    # Must be surrounded by at least one blank line
    if not (prev_blank or next_blank):
        return False
    # Should NOT end with sentence-ending punctuation
    if stripped[-1] in ".!?;:,":
        return False
    # Should not be a bullet
    if _is_bullet(stripped):
        return False
    # Bonus: mostly title-case or all-caps is a strong signal
    words = stripped.split()
    if len(words) <= 8:
        return True
    return False


def _chunk_body_lines(lines: list[str], max_lines: int = 7) -> list[list[str]]:
    """Split a long list of body lines into chunks that fit on a slide."""
    if len(lines) <= max_lines:
        return [lines]
    chunks = []
    for i in range(0, len(lines), max_lines):
        chunks.append(lines[i : i + max_lines])
    return chunks


def _split_long_paragraph(text: str, max_chars: int = 200) -> list[str]:
    """Split a long paragraph into sentence-level bullet points."""
    sentences = re.split(r"(?<=[.!?])\s+", text)
    result = []
    current = ""
    for s in sentences:
        if current and len(current) + len(s) + 1 > max_chars:
            result.append(current.strip())
            current = s
        else:
            current = (current + " " + s).strip() if current else s
    if current:
        result.append(current.strip())
    return result


def _has_markdown_headings(text: str) -> bool:
    """Check whether the text contains any markdown-style headings."""
    return bool(re.search(r"^#{1,4}\s+.+", text, re.MULTILINE))


def parse_text_to_deck(raw_text: str, deck_title: str = "", max_bullets: int = 7) -> DeckContent:
    """
    Parse raw text (with optional markdown headings) into a DeckContent.

    Strategy:
    1. If markdown headings are present → use them as slide boundaries.
    2. Otherwise, detect implicit headings (short standalone lines) and
       paragraph blocks as sections.
    3. Bullet points and body text under a heading become that slide's body.
    4. Long slides are auto-chunked.
    """
    lines = raw_text.split("\n")
    deck = DeckContent(title=deck_title or "Presentation")

    # Check for subtitle marker (from docx parsing)
    for i, line in enumerate(lines):
        if line.startswith("_subtitle: "):
            deck.subtitle = line.replace("_subtitle: ", "").strip()
            lines[i] = ""
            break

    has_md = _has_markdown_headings(raw_text)

    if has_md:
        deck = _parse_markdown_structured(lines, deck, deck_title, max_bullets)
    else:
        deck = _parse_plain_text(lines, deck, deck_title, max_bullets)

    return deck


def _parse_markdown_structured(
    lines: list[str], deck: DeckContent, deck_title: str, max_bullets: int
) -> DeckContent:
    """Parse text that contains markdown headings."""
    # Find the deck title from the first H1 if not provided
    if not deck_title:
        for line in lines:
            heading = _is_heading(line)
            if heading and heading[0] == 1:
                deck.title = heading[1]
                break

    current_slide: SlideContent | None = None
    body_buffer: list[str] = []

    # When the LLM uses `---` between two groups of bullets in a comparison
    # slide, we stash the left column here and merge it back in flush_slide.
    comparison_left: list[str] | None = None

    def promote_slide_type(sc: SlideContent) -> None:
        """Upgrade an assertion-level slide_type based on its body content
        — e.g. a bullet-less blockquote becomes a quote slide — and fold
        LLM creative blocks (BIG, STAT) into the right structured field."""
        # Gather any [[BIG]] or [[STAT]] tokens from body lines.
        big_found: tuple[str, str] | None = None
        stats_found: list[tuple[str, str]] = []
        surviving: list[str] = []
        for ln in sc.body_lines:
            stats = _parse_stats_line(ln)
            if stats:
                stats_found.extend(stats)
                continue
            big = _parse_big_line(ln)
            if big is not None:
                big_found = big
                continue
            surviving.append(ln)

        if stats_found:
            sc.stats = stats_found
            if sc.layout_hint in (None, "content"):
                sc.slide_type = "stat_grid"
                sc.layout_hint = sc.layout_hint or "stat_grid"
        if big_found is not None:
            sc.big_number = big_found
            if sc.layout_hint in (None, "content"):
                sc.slide_type = "bignum"
                sc.layout_hint = sc.layout_hint or "bignum"
        sc.body_lines = surviving

        # Honor explicit layout hints.
        if sc.layout_hint == "section" and not sc.body_lines:
            sc.slide_type = "section"
        if sc.layout_hint == "quote" and sc.quote_text is None and sc.body_lines:
            # Promote the first line as the quote if the LLM didn't use `> `.
            sc.quote_text = sc.body_lines[0].lstrip("> ").strip()
            sc.body_lines = sc.body_lines[1:]
            sc.slide_type = "quote"
        if sc.quote_text is not None:
            sc.slide_type = "quote"

    def flush_slide():
        nonlocal current_slide, body_buffer, comparison_left
        if current_slide is not None:
            current_slide.body_lines = body_buffer
            if comparison_left is not None:
                # Comparison: left column was already saved; the buffer is
                # the right column. Encode as "LEFT: …" / "RIGHT: …" prefix
                # so the renderer can split them back out.
                left = comparison_left
                right = body_buffer
                current_slide.body_lines = (
                    [f"__LEFT__{ln}" for ln in left]
                    + [f"__RIGHT__{ln}" for ln in right]
                )
                if current_slide.layout_hint in (None, "content"):
                    current_slide.layout_hint = "comparison"
                    current_slide.slide_type = "content"
                comparison_left = None

            promote_slide_type(current_slide)

            # Drop the empty H2 companion content slide when no bullets
            # followed the section divider (avoids duplicate titles).
            if (
                current_slide.level == 2
                and current_slide.slide_type == "content"
                and not current_slide.body_lines
                and current_slide.quote_text is None
                and current_slide.big_number is None
                and not current_slide.stats
            ):
                current_slide = None
                body_buffer = []
                return

            # Creative slide types shouldn't be chunked by bullet-count —
            # they carry structured content that must stay intact. Same
            # goes for two-column layouts: chunking would lose the
            # __LEFT__/__RIGHT__ markers and collapse to one column.
            creative_types = {"quote", "bignum", "stat_grid", "section"}
            creative_hints = {"two_content", "comparison", "stat_grid", "bignum", "quote"}
            if (
                current_slide.slide_type in creative_types
                or current_slide.layout_hint in creative_hints
            ):
                deck.slides.append(current_slide)
            else:
                chunks = _chunk_body_lines(current_slide.body_lines, max_bullets)
                if len(chunks) == 1:
                    deck.slides.append(current_slide)
                else:
                    for idx, chunk in enumerate(chunks):
                        s = SlideContent(
                            title=current_slide.title + (f" (cont.)" if idx > 0 else ""),
                            body_lines=chunk,
                            level=current_slide.level,
                            slide_type=current_slide.slide_type,
                            layout_hint=current_slide.layout_hint,
                            accent_hint=current_slide.accent_hint,
                        )
                        deck.slides.append(s)
        current_slide = None
        body_buffer = []
        comparison_left = None

    first_h1_seen = False
    for line in lines:
        stripped = line.strip()
        heading = _is_heading(stripped)

        if heading:
            level, raw_text = heading
            text, layout_hint, accent_hint = _extract_heading_hints(raw_text)
            flush_slide()
            if level == 1:
                if not first_h1_seen:
                    first_h1_seen = True
                    deck.title = deck.title or text
                    current_slide = SlideContent(
                        title=text, level=1, slide_type="title",
                        layout_hint=layout_hint, accent_hint=accent_hint,
                    )
                else:
                    current_slide = SlideContent(
                        title=text, level=1, slide_type="section",
                        layout_hint=layout_hint, accent_hint=accent_hint,
                    )
            elif level == 2:
                # H2 is a section divider by default. When the LLM tags the
                # H2 with a creative layout (stat_grid, bignum, quote,
                # two_content, …) the heading IS the creative slide — no
                # section-divider companion. A plain H2 with no hint keeps
                # the legacy behavior: emit a section slide, then a companion
                # content slide for any following bullets.
                if layout_hint is None:
                    deck.slides.append(
                        SlideContent(title=text, level=2, slide_type="section")
                    )
                    current_slide = SlideContent(
                        title=text, level=2, slide_type="content",
                        layout_hint=None, accent_hint=accent_hint,
                    )
                elif layout_hint == "section":
                    current_slide = SlideContent(
                        title=text, level=2, slide_type="section",
                        layout_hint=layout_hint, accent_hint=accent_hint,
                    )
                else:
                    current_slide = SlideContent(
                        title=text, level=2, slide_type="content",
                        layout_hint=layout_hint, accent_hint=accent_hint,
                    )
            else:
                current_slide = SlideContent(
                    title=text, level=level, slide_type="content",
                    layout_hint=layout_hint, accent_hint=accent_hint,
                )
            continue

        # Comparison column separator: `---` on its own line.
        if stripped == "---" and current_slide is not None:
            comparison_left = body_buffer
            body_buffer = []
            continue

        # Blockquote → pull-quote slide content.
        if stripped.startswith("> "):
            if current_slide is None:
                current_slide = SlideContent(title="", slide_type="quote")
            current_slide.quote_text = stripped[2:].strip()
            current_slide.slide_type = "quote"
            if current_slide.layout_hint is None:
                current_slide.layout_hint = "quote"
            continue

        bullet_text = _is_bullet(stripped)
        if bullet_text:
            if current_slide is None:
                current_slide = SlideContent(title="", slide_type="content")
            body_buffer.append(bullet_text)
            continue

        if stripped:
            if current_slide is None:
                current_slide = SlideContent(title="", slide_type="content")
            # If this is a long paragraph, split into sentences
            if len(stripped) > 200:
                body_buffer.extend(_split_long_paragraph(stripped))
            else:
                body_buffer.append(stripped)

    flush_slide()
    return deck


def _parse_plain_text(
    lines: list[str], deck: DeckContent, deck_title: str, max_bullets: int
) -> DeckContent:
    """
    Parse unstructured plain text by detecting implicit headings
    (short standalone lines between blank lines) and grouping
    body paragraphs under them.
    """
    # Step 1: Detect implicit headings using heuristics
    annotated: list[tuple[str, str]] = []  # (type, text): "heading", "bullet", "body", "blank"

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            annotated.append(("blank", ""))
            continue

        bullet_text = _is_bullet(stripped)
        if bullet_text:
            annotated.append(("bullet", bullet_text))
            continue

        prev_blank = (i == 0) or (lines[i - 1].strip() == "")
        next_blank = (i == len(lines) - 1) or (lines[i + 1].strip() == "")

        if _looks_like_heading(stripped, prev_blank, next_blank):
            annotated.append(("heading", stripped))
        else:
            annotated.append(("body", stripped))

    # Step 2: Group into sections (heading → body/bullets until next heading)
    sections: list[tuple[str, list[str]]] = []
    current_heading = ""
    current_body: list[str] = []

    for kind, text in annotated:
        if kind == "heading":
            if current_heading or current_body:
                sections.append((current_heading, current_body))
            current_heading = text
            current_body = []
        elif kind == "bullet":
            current_body.append(text)
        elif kind == "body":
            # Split long paragraphs into digestible lines
            if len(text) > 200:
                current_body.extend(_split_long_paragraph(text))
            else:
                current_body.append(text)
        # blank → skip

    if current_heading or current_body:
        sections.append((current_heading, current_body))

    # Step 3: Build slides from sections
    if not sections:
        return deck

    # Use the first section heading (or first body text) as deck title
    first_heading = sections[0][0]
    first_body = sections[0][1]

    if not deck_title:
        if first_heading:
            deck.title = first_heading
        elif first_body:
            deck.title = first_body[0][:60]
        else:
            deck.title = "Presentation"

    # Title slide
    deck.slides.append(SlideContent(
        title=deck.title,
        slide_type="title",
    ))

    # If the first section was used as the title, still include its body as a content slide
    start = 0
    if first_heading:
        # First section heading was consumed as the deck title — skip its heading
        # but keep its body content if any
        if first_body:
            chunks = _chunk_body_lines(first_body, max_bullets)
            for idx, chunk in enumerate(chunks):
                deck.slides.append(SlideContent(
                    title=first_heading + (" (cont.)" if idx > 0 else ""),
                    body_lines=chunk,
                    slide_type="content",
                ))
        start = 1
    elif first_body:
        # No heading on first section — body was already noted, use it
        chunks = _chunk_body_lines(first_body, max_bullets)
        for idx, chunk in enumerate(chunks):
            deck.slides.append(SlideContent(
                title="" if idx == 0 else "(cont.)",
                body_lines=chunk,
                slide_type="content",
            ))
        start = 1

    for heading, body in sections[start:]:
        if not heading and not body:
            continue

        chunks = _chunk_body_lines(body, max_bullets) if body else [[]]
        for idx, chunk in enumerate(chunks):
            slide_type = "section" if not body and heading else "content"
            deck.slides.append(SlideContent(
                title=heading + (" (cont.)" if idx > 0 else ""),
                body_lines=chunk,
                level=2,
                slide_type=slide_type,
            ))

    return deck


# ─────────────────────────────────────────────────────────────
# 4. TEMPLATE ANALYZER — Extract style info from a .pptx
# ─────────────────────────────────────────────────────────────

@dataclass
class LayoutInfo:
    """Everything we want to know about one slide layout in a template."""
    idx: int
    name: str
    # List of (placeholder_idx, placeholder_type_str) tuples.
    placeholders: list[tuple[int, str]] = field(default_factory=list)
    # Inferred purpose tags, e.g. ["content", "two_content"]. Used by
    # _choose_layout_by_hint to snap LLM hints to available layouts.
    purposes: list[str] = field(default_factory=list)


# Neutral fallback palette when theme1.xml is unreadable. Mid-tone enough
# to read on both light and dark masters.
_NEUTRAL_ACCENTS = [
    RGBColor(0x2E, 0x5A, 0x87),  # slate blue
    RGBColor(0xE8, 0x8B, 0x2C),  # amber
    RGBColor(0x4E, 0x9C, 0x6B),  # green
    RGBColor(0xB8, 0x3F, 0x3F),  # red
    RGBColor(0x6F, 0x54, 0xA3),  # purple
    RGBColor(0x4A, 0x75, 0x8A),  # teal
]


@dataclass
class TemplateStyle:
    """Style information extracted from a template presentation."""
    presentation: Presentation
    # Layout indices by purpose (kept for backwards compat with callers).
    title_layout_idx: int = 0
    section_layout_idx: int = 0
    content_layout_idx: int = 1
    # Whether a real Section Header layout was found (vs. a fallback).
    # When False, section slides are synthesized using the template's
    # accent palette rather than the hardcoded navy default.
    section_layout_found: bool = False
    # Fonts (theme fonts preferred; placeholder sniffing is the fallback).
    title_font: str = "Calibri"
    body_font: str = "Calibri"
    major_font: str | None = None
    minor_font: str | None = None
    title_size: Pt = field(default_factory=lambda: Pt(36))
    section_title_size: Pt | None = None
    content_title_size: Pt | None = None
    body_size: Pt = field(default_factory=lambda: Pt(16))
    # Placeholder-derived colors (may be None if not specified in layout).
    title_color: RGBColor | None = None
    body_color: RGBColor | None = None
    # Theme palette.
    theme_colors: dict[str, RGBColor] = field(default_factory=dict)
    accent_palette: list[RGBColor] = field(default_factory=list)
    # Full layout catalog for hint-based selection.
    layout_catalog: list[LayoutInfo] = field(default_factory=list)
    # Rolling counter for cycling accents across repeated section/table
    # slides, so a deck with 5 section dividers feels designed rather than
    # repeating the same accent color every time.
    _accent_cursor: int = 0

    def next_accent(self) -> RGBColor:
        """Return the next accent color in rotation (modulo palette length)."""
        palette = self.accent_palette or _NEUTRAL_ACCENTS
        color = palette[self._accent_cursor % len(palette)]
        self._accent_cursor += 1
        return color

    def accent_at(self, idx: int | None) -> RGBColor:
        """Return a specific accent by index, or the first accent if idx is None."""
        palette = self.accent_palette or _NEUTRAL_ACCENTS
        if idx is None:
            return palette[0]
        return palette[idx % len(palette)]


def _find_best_layout(prs: Presentation, target: str) -> tuple[int, bool]:
    """Find the layout index that best matches a target purpose.

    Returns ``(index, matched)`` where ``matched`` is True when a layout
    whose name actually matched a keyword was found (vs. falling back to a
    generic index). Callers use ``matched`` to decide whether to trust the
    layout for decorative purposes — e.g. section slides fall back to the
    default renderer when no real Section Header layout exists.
    """
    layouts = prs.slide_layouts
    target_lower = target.lower()

    # Keywords to match layout names. "blank" is deliberately excluded from
    # "section" — a blank layout has no decoration and is indistinguishable
    # from a content slide, which defeats the purpose of a section divider.
    keyword_map = {
        "title": ["title slide", "title", "cover"],
        "section": ["section header", "section", "divider"],
        "content": [
            "title and content", "two content", "content",
            "title, content", "title and body", "text",
        ],
    }

    keywords = keyword_map.get(target_lower, [target_lower])
    for priority_kw in keywords:
        for idx, layout in enumerate(layouts):
            if layout.name and priority_kw in layout.name.lower():
                return idx, True

    # Fallback: title=0, section=0, content=1 (or 0 if only one layout)
    fallback = {"title": 0, "section": 0, "content": min(1, len(layouts) - 1)}
    return fallback.get(target_lower, 0), False


# ─────────────────────────────────────────────────────────────
# Theme XML inspection
# ─────────────────────────────────────────────────────────────

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _hex_to_rgb(hex_str: str) -> RGBColor | None:
    """Parse a 6-char hex string into an RGBColor, or None on failure."""
    if not hex_str:
        return None
    s = hex_str.strip().lstrip("#")
    if len(s) != 6:
        return None
    try:
        return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        return None


def _resolve_scheme_color(elem: ET.Element) -> RGBColor | None:
    """Given a <a:dk1>, <a:accent1>, etc., return an RGBColor. Handles the
    common srgbClr and sysClr (with lastClr fallback) variants. Returns None
    if neither child is present or parses successfully."""
    srgb = elem.find(f"{{{_A_NS}}}srgbClr")
    if srgb is not None:
        return _hex_to_rgb(srgb.get("val"))
    sys_clr = elem.find(f"{{{_A_NS}}}sysClr")
    if sys_clr is not None:
        # lastClr is the RGB that was in effect when the file was saved —
        # good enough as a deterministic mapping.
        return _hex_to_rgb(sys_clr.get("lastClr") or sys_clr.get("val"))
    return None


def _read_theme_from_pptx(template_path: str) -> tuple[dict[str, RGBColor], str | None, str | None]:
    """Pull the color and font scheme out of ppt/theme/theme1.xml.

    Returns ``(color_map, major_font, minor_font)``. ``color_map`` maps
    scheme slot names (dk1, lt1, accent1…accent6, hlink, folHlink) to
    RGBColor. Missing or malformed themes come back as ``({}, None, None)``
    — callers decide on fallbacks.
    """
    colors: dict[str, RGBColor] = {}
    major = None
    minor = None
    try:
        with zipfile.ZipFile(template_path) as zf:
            # Pick the first ppt/theme/theme*.xml — almost always theme1.
            theme_members = sorted(
                n for n in zf.namelist() if n.startswith("ppt/theme/") and n.endswith(".xml")
            )
            if not theme_members:
                return colors, major, minor
            with zf.open(theme_members[0]) as fh:
                root = ET.parse(fh).getroot()

        theme_elem = root.find(f"{{{_A_NS}}}themeElements")
        if theme_elem is None:
            return colors, major, minor

        clr_scheme = theme_elem.find(f"{{{_A_NS}}}clrScheme")
        if clr_scheme is not None:
            for child in clr_scheme:
                tag = child.tag.split("}", 1)[-1]
                rgb = _resolve_scheme_color(child)
                if rgb is not None:
                    colors[tag] = rgb

        font_scheme = theme_elem.find(f"{{{_A_NS}}}fontScheme")
        if font_scheme is not None:
            for child, key in (
                (font_scheme.find(f"{{{_A_NS}}}majorFont"), "major"),
                (font_scheme.find(f"{{{_A_NS}}}minorFont"), "minor"),
            ):
                if child is None:
                    continue
                latin = child.find(f"{{{_A_NS}}}latin")
                if latin is not None and latin.get("typeface"):
                    if key == "major":
                        major = latin.get("typeface")
                    else:
                        minor = latin.get("typeface")
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError) as exc:
        logger.warning("Could not parse theme from %s: %s", template_path, exc)

    return colors, major, minor


def _read_master_clr_map(prs: Presentation) -> dict[str, str]:
    """Read the slide master's clrMap — it remaps theme scheme slots to
    the names used in shapes (e.g. `accent1` in a shape may actually
    resolve to theme's `accent3`). Returns a dict like
    ``{"accent1": "accent1", "bg1": "lt1", …}``.

    Returns an empty dict if no master or no clrMap is present; callers
    should then use the theme's native slots directly.
    """
    if not prs.slide_masters:
        return {}
    # Pick master 0 — multi-master templates are rare; if present, master 0
    # is typically the primary. A more sophisticated future change could
    # pick the master that the most layouts reference.
    master = prs.slide_masters[0]
    clr_map = master.element.find(f"{{{_P_NS}}}clrMap")
    if clr_map is None:
        return {}
    return dict(clr_map.attrib)


def _build_accent_palette(
    theme_colors: dict[str, RGBColor], clr_map: dict[str, str]
) -> list[RGBColor]:
    """Assemble accent1..accent6 from the theme, honoring the master's
    color map. Falls back to neutral accents when nothing resolvable."""
    palette: list[RGBColor] = []
    for i in range(1, 7):
        key = f"accent{i}"
        # clrMap may remap "accent1" → some other slot; if missing, use
        # the theme's own key.
        theme_key = clr_map.get(key, key)
        rgb = theme_colors.get(theme_key) or theme_colors.get(key)
        if rgb is not None:
            palette.append(rgb)
    if not palette:
        return list(_NEUTRAL_ACCENTS)
    return palette


# ─────────────────────────────────────────────────────────────
# Layout catalog
# ─────────────────────────────────────────────────────────────

# Ordered so that higher-score purposes are preferred when multiple tie.
_PURPOSE_NAME_KEYWORDS: dict[str, list[str]] = {
    "title": ["title slide", "title", "cover"],
    "section": ["section header", "section", "divider"],
    "two_content": ["two content", "two-content", "comparison", "side by side", "two column"],
    "comparison": ["comparison", "versus", " vs "],
    "picture_caption": ["picture with caption", "picture", "photo", "image"],
    "quote": ["quote", "pull quote", "testimonial"],
    "stat_grid": ["stat", "metrics", "numbers", "kpi"],
    "callout": ["callout", "highlight", "spotlight"],
    "content": [
        "title and content", "title, content", "title and body",
        "content", "text",
    ],
    "blank": ["blank"],
}


def _placeholder_type_name(ph) -> str:
    """Return a short string for a placeholder's type (BODY, PICTURE, etc.).
    Falls back to the numeric repr if the enum is unavailable."""
    try:
        t = ph.placeholder_format.type
    except Exception:
        return "UNKNOWN"
    if t is None:
        return "UNKNOWN"
    # t is an enum value; str() gives e.g. "BODY (2)"
    return str(t).split(" ")[0]


def _score_layout_purposes(name: str, ph_types: list[str]) -> list[str]:
    """Score each candidate purpose and return the ones that pass threshold.

    Scoring (heuristic but stable):
      + 3 per name keyword match
      + 4 if placeholder signature strongly implies the purpose
      + 2 for weaker signature matches
    Purposes that score > 0 are returned, highest-first.
    """
    scores: dict[str, int] = {}
    name_lower = (name or "").lower()

    # Name keyword signals.
    for purpose, kws in _PURPOSE_NAME_KEYWORDS.items():
        for kw in kws:
            if kw in name_lower:
                scores[purpose] = scores.get(purpose, 0) + 3
                break

    # Placeholder-signature signals.
    body_count = ph_types.count("BODY") + ph_types.count("OBJECT")
    has_picture = "PICTURE" in ph_types
    has_chart = "CHART" in ph_types
    has_table = "TABLE" in ph_types
    has_title = "TITLE" in ph_types or "CENTER_TITLE" in ph_types
    has_subtitle = "SUBTITLE" in ph_types
    total = len(ph_types)

    if has_picture:
        scores["picture_caption"] = scores.get("picture_caption", 0) + 4
    if body_count >= 2:
        scores["two_content"] = scores.get("two_content", 0) + 4
        scores["comparison"] = scores.get("comparison", 0) + 2
    if body_count >= 1 and has_title and not has_picture:
        scores["content"] = scores.get("content", 0) + 3
    if has_title and body_count == 0 and not has_subtitle and total <= 2:
        # A title-only or title+decoration layout — good section divider
        # material whether or not the name says "section".
        scores["section"] = scores.get("section", 0) + 2
        scores["quote"] = scores.get("quote", 0) + 1
    if has_subtitle and has_title and body_count == 0:
        scores["title"] = scores.get("title", 0) + 4
    if has_chart or has_table:
        scores["stat_grid"] = scores.get("stat_grid", 0) + 3
    if total == 0:
        scores["blank"] = scores.get("blank", 0) + 5

    # Return purposes with any score, ordered by descending score.
    return [p for p, _ in sorted(scores.items(), key=lambda kv: -kv[1]) if scores[p] > 0]


def _build_layout_catalog(prs: Presentation) -> list[LayoutInfo]:
    """Enumerate every slide layout and score its purpose tags."""
    catalog: list[LayoutInfo] = []
    for idx, layout in enumerate(prs.slide_layouts):
        ph_tuples: list[tuple[int, str]] = []
        ph_types: list[str] = []
        for ph in layout.placeholders:
            try:
                ph_idx = ph.placeholder_format.idx
            except Exception:
                continue
            type_name = _placeholder_type_name(ph)
            ph_tuples.append((ph_idx, type_name))
            ph_types.append(type_name)
        purposes = _score_layout_purposes(layout.name or "", ph_types)
        catalog.append(LayoutInfo(
            idx=idx,
            name=layout.name or f"Layout {idx}",
            placeholders=ph_tuples,
            purposes=purposes,
        ))
    return catalog


def _choose_layout_by_hint(
    catalog: list[LayoutInfo],
    hint: str | None,
    default_purpose: str = "content",
) -> LayoutInfo:
    """Pick the best layout for the requested hint.

    ``hint`` is untrusted (typically from the LLM). Unknown values are
    silently downgraded to ``default_purpose``. When no layout in the
    catalog advertises the desired purpose, the function logs a warning
    and returns the simplest-signature layout so the slide still lands
    somewhere sensible.
    """
    if not catalog:
        raise ValueError("Empty layout catalog — template has no layouts.")

    # Known hints include every purpose we score in the catalog, plus a
    # handful of synonyms emitted by the creative renderers themselves
    # (``bignum`` lands on any single-emphasis layout; ``callout`` on a
    # minimal body). These aren't scored from placeholder signatures but
    # they're legitimate requests from callers.
    known = set(_PURPOSE_NAME_KEYWORDS.keys()) | {"bignum"}
    requested = hint if hint in known else default_purpose
    if hint and hint not in known:
        logger.warning("Unknown layout hint %r; falling back to %r", hint, default_purpose)

    # First try the exact requested purpose.
    for purpose in (requested, default_purpose, "content", "blank"):
        for entry in catalog:
            if purpose in entry.purposes:
                return entry

    # Nothing matched even "content" — return the layout with the fewest
    # placeholders (most generic), which is the safest empty canvas.
    return min(catalog, key=lambda li: len(li.placeholders))


def render_template_kit(style: TemplateStyle) -> str:
    """Produce a compact, human-readable summary of the template suitable
    for injecting into an LLM system prompt. Kept tight so it doesn't
    blow the context budget of small local models."""
    palette = style.accent_palette or _NEUTRAL_ACCENTS
    hex_palette = ", ".join(
        "#{:02X}{:02X}{:02X}".format(*[int(c) for c in (rgb[0], rgb[1], rgb[2])])
        for rgb in palette
    )
    major = style.major_font or style.title_font
    minor = style.minor_font or style.body_font
    purposes_seen: dict[str, str] = {}
    for entry in style.layout_catalog:
        for purpose in entry.purposes:
            purposes_seen.setdefault(purpose, entry.name)
    lines = [
        f"Template fonts: major={major}, minor={minor}",
        f"Template accent palette: {hex_palette}",
        "Available slide layouts in this template (use these purpose tags):",
    ]
    # Preferred presentation order.
    order = [
        "title", "section", "content", "two_content", "comparison",
        "picture_caption", "quote", "stat_grid", "callout", "blank",
    ]
    for p in order:
        if p in purposes_seen:
            lines.append(f"  - {p:<16} (e.g. \"{purposes_seen[p]}\")")
    return "\n".join(lines)


def _layout_title_size(layout) -> Pt | None:
    """Return the size of the first sized run on the layout's title
    placeholder (idx=0), or None if the layout doesn't advertise one."""
    for ph in layout.placeholders:
        if ph.placeholder_format.idx != 0 or not ph.has_text_frame:
            continue
        for para in ph.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    return run.font.size
    return None


def _extract_font_info(prs: Presentation) -> dict:
    """Scan existing slides/layouts to detect the fonts and colors in use."""
    info = {
        "title_font": "Calibri",
        "body_font": "Calibri",
        "title_size": Pt(36),
        "body_size": Pt(16),
        "title_color": None,
        "body_color": None,
    }

    # Scan slide layouts for placeholder fonts
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            if ph.has_text_frame:
                for para in ph.text_frame.paragraphs:
                    for run in para.runs:
                        font = run.font
                        if ph.placeholder_format.idx == 0:  # Title
                            if font.name:
                                info["title_font"] = font.name
                            if font.size:
                                info["title_size"] = font.size
                            try:
                                if font.color and font.color.type is not None:
                                    info["title_color"] = font.color.rgb
                            except (AttributeError, TypeError):
                                pass
                        elif ph.placeholder_format.idx == 1:  # Body
                            if font.name:
                                info["body_font"] = font.name
                            if font.size:
                                info["body_size"] = font.size
                            try:
                                if font.color and font.color.type is not None:
                                    info["body_color"] = font.color.rgb
                            except (AttributeError, TypeError):
                                pass

    # Also scan actual slides
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        font = run.font
                        if font.name and info["title_font"] == "Calibri":
                            info["title_font"] = font.name
                        if font.size and info["title_size"] == Pt(36):
                            info["title_size"] = font.size

    return info


def analyze_template(template_path: str) -> TemplateStyle:
    """Load a .pptx template and extract its style information.

    Goes beyond placeholder sniffing: also reads the theme's color scheme
    and font scheme, honors the master's clrMap remapping, and builds a
    full layout catalog with purpose tags. The result is rich enough to
    drive hint-based creative slide generation.
    """
    prs = Presentation(template_path)
    font_info = _extract_font_info(prs)

    theme_colors, major_font, minor_font = _read_theme_from_pptx(template_path)
    clr_map = _read_master_clr_map(prs)
    accent_palette = _build_accent_palette(theme_colors, clr_map)

    catalog = _build_layout_catalog(prs)

    # Keep the legacy by-purpose indices so downstream callers that still
    # reach into style.title_layout_idx / section_layout_idx / content_layout_idx
    # keep working without modification.
    title_idx, _ = _find_best_layout(prs, "title")
    section_idx, section_found = _find_best_layout(prs, "section")
    content_idx, _ = _find_best_layout(prs, "content")

    layouts = prs.slide_layouts
    section_title_size = (
        _layout_title_size(layouts[section_idx]) if section_found else None
    )
    content_title_size = _layout_title_size(layouts[content_idx])

    # Prefer theme-declared fonts, but keep placeholder-sniffed names as
    # a fallback — Office themes sometimes declare generic fonts like
    # "+mn-lt" that aren't useful to python-pptx's run.font.name.
    title_font = major_font or font_info["title_font"]
    body_font = minor_font or font_info["body_font"]

    style = TemplateStyle(
        presentation=prs,
        title_layout_idx=title_idx,
        section_layout_idx=section_idx,
        content_layout_idx=content_idx,
        section_layout_found=section_found,
        title_font=title_font,
        body_font=body_font,
        major_font=major_font,
        minor_font=minor_font,
        title_size=font_info["title_size"],
        section_title_size=section_title_size,
        content_title_size=content_title_size,
        body_size=font_info["body_size"],
        title_color=font_info["title_color"],
        body_color=font_info["body_color"],
        theme_colors=theme_colors,
        accent_palette=accent_palette,
        layout_catalog=catalog,
    )
    return style


# ─────────────────────────────────────────────────────────────
# 5. SLIDE GENERATORS — Build the .pptx output
# ─────────────────────────────────────────────────────────────

# ----- Default (no template) style constants -----

DEFAULT_COLORS = {
    "bg_dark": RGBColor(0x1E, 0x27, 0x61),      # navy
    "bg_light": RGBColor(0xFF, 0xFF, 0xFF),       # white
    "title_on_dark": RGBColor(0xFF, 0xFF, 0xFF),  # white text
    "title_on_light": RGBColor(0x1E, 0x27, 0x61), # navy text
    "body": RGBColor(0x33, 0x33, 0x33),           # dark gray
    "accent": RGBColor(0xCA, 0xDC, 0xFC),         # ice blue
    "subtle": RGBColor(0x64, 0x74, 0x8B),         # muted gray
}
DEFAULT_TITLE_FONT = "Georgia"
DEFAULT_BODY_FONT = "Calibri"


def _apply_font(run, font_name: str, size: Pt, color: RGBColor | None = None, bold: bool = False):
    """Apply font properties to a text run."""
    run.font.name = font_name
    run.font.size = size
    if color:
        run.font.color.rgb = color
    run.font.bold = bold


def _add_title_slide_default(prs: Presentation, slide_content: SlideContent, subtitle: str = ""):
    """Create a title slide with the default built-in style."""
    slide_layout = prs.slide_layouts[0]  # Title Slide layout
    slide = prs.slides.add_slide(slide_layout)

    # Set dark background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = DEFAULT_COLORS["bg_dark"]

    # Title
    if slide.shapes.title:
        tf = slide.shapes.title.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = slide_content.title
        _apply_font(run, DEFAULT_TITLE_FONT, Pt(40), DEFAULT_COLORS["title_on_dark"], bold=True)

    # Subtitle
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = subtitle or ""
            _apply_font(run, DEFAULT_BODY_FONT, Pt(18), DEFAULT_COLORS["accent"])
            break


def _add_section_slide_default(prs: Presentation, slide_content: SlideContent):
    """Create a section divider slide with default style."""
    # Use a blank layout and build manually
    slide_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]  # Blank
    slide = prs.slides.add_slide(slide_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = DEFAULT_COLORS["bg_dark"]

    # Accent line
    from pptx.util import Inches
    from pptx.shapes.autoshape import Shape
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(0.8), Inches(2.2), Inches(1.5), Inches(0.06)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = DEFAULT_COLORS["accent"]
    shape.line.fill.background()

    # Section title
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(8.4), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = slide_content.title
    _apply_font(run, DEFAULT_TITLE_FONT, Pt(36), DEFAULT_COLORS["title_on_dark"], bold=True)


def _add_content_slide_default(prs: Presentation, slide_content: SlideContent):
    """Create a content slide with default style."""
    slide_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]  # Blank
    slide = prs.slides.add_slide(slide_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = DEFAULT_COLORS["bg_light"]

    y_cursor = Inches(0.5)

    # Slide title
    if slide_content.title:
        txBox = slide.shapes.add_textbox(Inches(0.7), y_cursor, Inches(8.6), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = slide_content.title
        _apply_font(run, DEFAULT_TITLE_FONT, Pt(28), DEFAULT_COLORS["title_on_light"], bold=True)
        y_cursor = Inches(1.4)

    # Body content
    if slide_content.body_lines:
        txBox = slide.shapes.add_textbox(
            Inches(0.7), y_cursor, Inches(8.6), Inches(4.0)
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        for i, line in enumerate(slide_content.body_lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.alignment = PP_ALIGN.LEFT
            p.space_after = Pt(8)

            # Add a bullet indicator
            p.level = 0
            pPr = p._pPr
            if pPr is None:
                from pptx.oxml.ns import qn
                pPr = p._p.get_or_add_pPr()

            run = p.add_run()
            run.text = f"  •  {line}"
            _apply_font(run, DEFAULT_BODY_FONT, Pt(15), DEFAULT_COLORS["body"])


# ─── Table slide helpers ──────────────────────────────────────

def _compute_col_widths(headers: list[str], rows: list[list[str]], total_width: float) -> list[float]:
    """
    Compute proportional column widths based on max content length per column.
    Returns widths in inches that sum to total_width.
    """
    n_cols = len(headers)
    if n_cols == 0:
        return []

    max_lens = []
    for c in range(n_cols):
        col_max = len(headers[c])
        for row in rows:
            if c < len(row):
                col_max = max(col_max, len(row[c]))
        max_lens.append(max(col_max, 3))  # minimum 3 chars

    total_chars = sum(max_lens)
    # Proportional, with a minimum width per column
    min_width = 0.8  # inches
    widths = []
    for ml in max_lens:
        w = max(min_width, (ml / total_chars) * total_width)
        widths.append(w)

    # Normalize so they sum to total_width
    scale = total_width / sum(widths)
    widths = [w * scale for w in widths]
    return widths


def _build_table_on_slide(
    slide,
    headers: list[str],
    rows: list[list[str]],
    left: float,
    top: float,
    width: float,
    height: float,
    header_font: str = "Calibri",
    body_font: str = "Calibri",
    header_size: Pt = None,
    body_size: Pt = None,
    header_color: RGBColor | None = None,
    body_color: RGBColor | None = None,
    header_bg: RGBColor | None = None,
    stripe_bg: RGBColor | None = None,
):
    """
    Add a formatted table shape to a slide.
    Handles dynamic column widths and alternating row stripes.
    """
    if header_size is None:
        header_size = Pt(11)
    if body_size is None:
        body_size = Pt(10)
    if header_bg is None:
        header_bg = RGBColor(0x1E, 0x27, 0x61)  # navy
    if header_color is None:
        header_color = RGBColor(0xFF, 0xFF, 0xFF)  # white
    if body_color is None:
        body_color = RGBColor(0x33, 0x33, 0x33)
    if stripe_bg is None:
        stripe_bg = RGBColor(0xF0, 0xF4, 0xF8)  # very light blue-gray

    n_cols = len(headers)
    n_rows = len(rows) + 1  # +1 for header row

    # Compute proportional column widths
    col_widths = _compute_col_widths(headers, rows, width)

    table_shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(left), Inches(top),
        Inches(width), Inches(height),
    )
    table = table_shape.table

    # Set column widths
    for c, w in enumerate(col_widths):
        table.columns[c].width = Inches(w)

    # Disable built-in banding (we'll do our own striping)
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is not None:
        tblPr.set("bandRow", "0")
        tblPr.set("bandCol", "0")
        tblPr.set("firstRow", "0")
        tblPr.set("lastRow", "0")

    # Helper to style a cell
    def _style_cell(cell, text, font_name, font_size, font_color, bg_color=None, bold=False):
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = True
        # Reduce internal margins for compact tables
        tf.margin_left = Inches(0.05)
        tf.margin_right = Inches(0.05)
        tf.margin_top = Inches(0.03)
        tf.margin_bottom = Inches(0.03)

        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = font_size
        run.font.color.rgb = font_color
        run.font.bold = bold

        if bg_color:
            from pptx.oxml.ns import qn
            tcPr = cell._tc.get_or_add_tcPr()
            solidFill = tcPr.makeelement(qn("a:solidFill"), {})
            srgbClr = solidFill.makeelement(qn("a:srgbClr"), {"val": str(bg_color)})
            solidFill.append(srgbClr)
            tcPr.append(solidFill)

    # Header row
    for c, h in enumerate(headers):
        _style_cell(
            table.cell(0, c), h,
            header_font, header_size, header_color,
            bg_color=header_bg, bold=True,
        )

    # Data rows with alternating stripes
    for r, row_data in enumerate(rows):
        bg = stripe_bg if r % 2 == 1 else None
        for c in range(n_cols):
            val = row_data[c] if c < len(row_data) else ""
            _style_cell(
                table.cell(r + 1, c), val,
                body_font, body_size, body_color,
                bg_color=bg,
            )

    return table_shape


def _add_table_slide_default(prs: Presentation, slide_content: SlideContent):
    """Create a table slide with the default built-in style."""
    slide_layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]  # Blank
    slide = prs.slides.add_slide(slide_layout)

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = DEFAULT_COLORS["bg_light"]

    y_cursor = 0.4

    # Slide title
    if slide_content.title:
        txBox = slide.shapes.add_textbox(
            Inches(0.5), Inches(y_cursor), Inches(9.0), Inches(0.6)
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = slide_content.title
        _apply_font(run, DEFAULT_TITLE_FONT, Pt(24), DEFAULT_COLORS["title_on_light"], bold=True)
        y_cursor = 1.1

    # Table
    n_rows = len(slide_content.table_rows)
    # Dynamically size: more rows → smaller row height
    available_height = 5.625 - y_cursor - 0.3  # slide height minus margins
    row_height_estimate = min(0.35, available_height / (n_rows + 1))
    table_height = row_height_estimate * (n_rows + 1)

    _build_table_on_slide(
        slide,
        headers=slide_content.table_headers,
        rows=slide_content.table_rows,
        left=0.5,
        top=y_cursor,
        width=9.0,
        height=table_height,
        header_font=DEFAULT_BODY_FONT,
        body_font=DEFAULT_BODY_FONT,
        header_bg=DEFAULT_COLORS["bg_dark"],
        header_color=DEFAULT_COLORS["title_on_dark"],
        body_color=DEFAULT_COLORS["body"],
        stripe_bg=RGBColor(0xEE, 0xF2, 0xF7),
    )


def _add_table_slide_from_template(prs: Presentation, style: TemplateStyle, slide_content: SlideContent):
    """Create a table slide using the template's layout and style."""
    layouts = prs.slide_layouts
    layout_idx = min(style.content_layout_idx, len(layouts) - 1)
    slide = prs.slides.add_slide(layouts[layout_idx])

    # Set the title placeholder if available
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.text = slide_content.title
            for para in ph.text_frame.paragraphs:
                for run in para.runs:
                    run.font.name = style.title_font
                    if style.title_color:
                        run.font.color.rgb = style.title_color
            break

    n_rows = len(slide_content.table_rows)
    available_height = 4.0
    row_height_estimate = min(0.35, available_height / (n_rows + 1))
    table_height = row_height_estimate * (n_rows + 1)

    # Header background comes from the template's accent palette (cycling
    # each time a table is drawn so repeated table slides in one deck don't
    # all wear the same color). Falls back through the rotation helper to
    # the neutral palette if theme parsing failed.
    header_bg = style.next_accent()

    _build_table_on_slide(
        slide,
        headers=slide_content.table_headers,
        rows=slide_content.table_rows,
        left=0.5,
        top=1.5,
        width=9.0,
        height=table_height,
        header_font=style.body_font,
        body_font=style.body_font,
        header_size=Pt(11),
        body_size=style.body_size if style.body_size else Pt(10),
        header_bg=header_bg,
        header_color=RGBColor(0xFF, 0xFF, 0xFF),
        body_color=style.body_color or RGBColor(0x33, 0x33, 0x33),
    )


def _set_placeholder_text(ph, text: str, fallback_font: str,
                          fallback_size: Pt | None, fallback_color: RGBColor | None):
    """Write ``text`` into a placeholder while preserving the layout's
    existing run formatting where possible.

    The layout placeholder already carries the template designer's font,
    size, weight, and color. Calling ``ph.text = ...`` blows all of that
    away. Instead, reuse the first paragraph's first run (which inherits
    layout styling via the run/paragraph/default chain) and only fill in
    ``font.name``/``font.size``/``font.color`` when the layout itself
    didn't provide one.
    """
    tf = ph.text_frame
    # Keep the first paragraph (inherits layout pPr); drop the rest.
    while len(tf.paragraphs) > 1:
        tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
    p = tf.paragraphs[0]
    # Preserve the first run if present (keeps its rPr); otherwise make one.
    if p.runs:
        run = p.runs[0]
        # Clear any additional runs so we end up with a single run.
        for extra in list(p.runs[1:]):
            extra._r.getparent().remove(extra._r)
    else:
        run = p.add_run()
    run.text = text
    # Only fill in properties the layout didn't already supply.
    if not run.font.name and fallback_font:
        run.font.name = fallback_font
    if run.font.size is None and fallback_size:
        run.font.size = fallback_size
    try:
        has_color = run.font.color and run.font.color.type is not None
    except (AttributeError, TypeError):
        has_color = False
    if not has_color and fallback_color:
        run.font.color.rgb = fallback_color


def _fill_body_placeholder(ph, lines: list[str], style: TemplateStyle):
    """Write ``lines`` into a body placeholder, reusing the layout's
    existing bullet/paragraph formatting for the first line so the
    template's indent and bullet glyph stay intact."""
    tf = ph.text_frame
    # Drop every paragraph past the first so we can rebuild from a clean slate.
    while len(tf.paragraphs) > 1:
        tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
    first_p = tf.paragraphs[0]
    # Clear runs from the first paragraph but keep its pPr (bullet/indent).
    for r in list(first_p.runs):
        r._r.getparent().remove(r._r)

    for i, line in enumerate(lines):
        if i == 0:
            p = first_p
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        if not run.font.name and style.body_font:
            run.font.name = style.body_font
        if run.font.size is None and style.body_size:
            run.font.size = style.body_size
        try:
            has_color = run.font.color and run.font.color.type is not None
        except (AttributeError, TypeError):
            has_color = False
        if not has_color and style.body_color:
            run.font.color.rgb = style.body_color


def _slide_dims(prs: Presentation) -> tuple[float, float]:
    """Return (width_in, height_in) of the presentation slides."""
    return prs.slide_width / 914400.0, prs.slide_height / 914400.0


def _pick_layout(style: TemplateStyle, hint: str | None, default: str):
    """Resolve a slide's layout_hint (+ default purpose) to an actual layout
    object. Returns (layout_info, layout_obj)."""
    info = _choose_layout_by_hint(style.layout_catalog, hint, default)
    layout_obj = style.presentation.slide_layouts[info.idx]
    return info, layout_obj


def _add_quote_slide_from_template(
    prs: Presentation, style: TemplateStyle, sc: SlideContent,
):
    """Render a pull-quote slide: large centered quotation in the theme's
    major font, colored with an accent, with a short accent bar on the left."""
    info, layout = _pick_layout(style, sc.layout_hint or "quote", "section")
    slide = prs.slides.add_slide(layout)
    width_in, height_in = _slide_dims(prs)
    accent = style.accent_at(sc.accent_hint)

    # Wipe the layout's placeholders' default text so we can lay our own out.
    for ph in list(slide.placeholders):
        if ph.has_text_frame:
            ph.text_frame.clear()

    # Accent bar on the left.
    bar = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        Inches(0.8), Inches(height_in * 0.38),
        Inches(0.12), Inches(height_in * 0.25),
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()

    # Quote text box.
    tb = slide.shapes.add_textbox(
        Inches(1.2), Inches(height_in * 0.30),
        Inches(width_in - 2.0), Inches(height_in * 0.40),
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = f"\u201c{sc.quote_text or sc.title}\u201d"
    run.font.name = style.major_font or style.title_font
    run.font.size = Pt(32)
    run.font.color.rgb = accent
    run.font.italic = True

    # Slide title (the heading, if distinct from the quote) as a small
    # caption above the quote rather than an attribution — the heading is
    # the slide's topic, not the quote's author.
    if sc.title and sc.quote_text and sc.title != sc.quote_text:
        cap_tb = slide.shapes.add_textbox(
            Inches(1.2), Inches(height_in * 0.18),
            Inches(width_in - 2.4), Inches(0.6),
        )
        cap_tf = cap_tb.text_frame
        cap_tf.word_wrap = True
        cp = cap_tf.paragraphs[0]
        cp.alignment = PP_ALIGN.LEFT
        crun = cp.add_run()
        crun.text = sc.title
        crun.font.name = style.body_font
        crun.font.size = Pt(16)
        crun.font.bold = True
        if style.body_color:
            crun.font.color.rgb = style.body_color


def _add_bignum_slide_from_template(
    prs: Presentation, style: TemplateStyle, sc: SlideContent,
):
    """Render a big-number slide: one huge centered figure over a small label."""
    info, layout = _pick_layout(style, sc.layout_hint or "callout", "section")
    slide = prs.slides.add_slide(layout)
    width_in, height_in = _slide_dims(prs)
    accent = style.accent_at(sc.accent_hint)

    for ph in list(slide.placeholders):
        if ph.has_text_frame:
            ph.text_frame.clear()

    # Optional slide title above.
    if sc.title:
        title_tb = slide.shapes.add_textbox(
            Inches(0.7), Inches(0.4), Inches(width_in - 1.4), Inches(0.8),
        )
        ttf = title_tb.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.LEFT
        trun = tp.add_run()
        trun.text = sc.title
        trun.font.name = style.title_font
        trun.font.size = Pt(20)
        if style.title_color:
            trun.font.color.rgb = style.title_color

    value, label = sc.big_number or ("", "")
    # Big number.
    num_tb = slide.shapes.add_textbox(
        Inches(0.7), Inches(height_in * 0.28),
        Inches(width_in - 1.4), Inches(height_in * 0.38),
    )
    ntf = num_tb.text_frame
    ntf.word_wrap = True
    ntf.vertical_anchor = MSO_ANCHOR.MIDDLE
    np_ = ntf.paragraphs[0]
    np_.alignment = PP_ALIGN.CENTER
    nrun = np_.add_run()
    nrun.text = value
    nrun.font.name = style.major_font or style.title_font
    nrun.font.size = Pt(108)
    nrun.font.bold = True
    nrun.font.color.rgb = accent

    # Label.
    lab_tb = slide.shapes.add_textbox(
        Inches(0.7), Inches(height_in * 0.70),
        Inches(width_in - 1.4), Inches(0.8),
    )
    ltf = lab_tb.text_frame
    ltf.word_wrap = True
    lp = ltf.paragraphs[0]
    lp.alignment = PP_ALIGN.CENTER
    lrun = lp.add_run()
    lrun.text = label
    lrun.font.name = style.body_font
    lrun.font.size = Pt(20)
    if style.body_color:
        lrun.font.color.rgb = style.body_color


def _add_stat_grid_slide_from_template(
    prs: Presentation, style: TemplateStyle, sc: SlideContent,
):
    """Render a grid of 2–6 big numbers, each in a rotating accent color."""
    info, layout = _pick_layout(style, sc.layout_hint or "stat_grid", "content")
    slide = prs.slides.add_slide(layout)
    width_in, height_in = _slide_dims(prs)

    for ph in list(slide.placeholders):
        if ph.has_text_frame:
            ph.text_frame.clear()

    # Title bar at top.
    if sc.title:
        tb = slide.shapes.add_textbox(
            Inches(0.7), Inches(0.4), Inches(width_in - 1.4), Inches(0.8),
        )
        ttf = tb.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.LEFT
        trun = tp.add_run()
        trun.text = sc.title
        trun.font.name = style.title_font
        trun.font.size = Pt(24)
        trun.font.bold = True
        if style.title_color:
            trun.font.color.rgb = style.title_color

    stats = sc.stats or []
    if not stats:
        return

    n = max(1, min(6, len(stats)))
    cols = n if n <= 3 else (n + 1) // 2
    rows = 1 if n <= 3 else 2
    palette = style.accent_palette or _NEUTRAL_ACCENTS

    grid_top = 1.6
    grid_height = height_in - grid_top - 0.4
    cell_w = (width_in - 1.4) / cols
    cell_h = grid_height / rows

    for i, (value, label) in enumerate(stats[:n]):
        r, c = divmod(i, cols)
        left = 0.7 + c * cell_w
        top = grid_top + r * cell_h
        # Value
        vtb = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(cell_w - 0.2), Inches(cell_h * 0.55),
        )
        vtf = vtb.text_frame
        vtf.word_wrap = True
        vtf.vertical_anchor = MSO_ANCHOR.BOTTOM
        vp = vtf.paragraphs[0]
        vp.alignment = PP_ALIGN.CENTER
        vrun = vp.add_run()
        vrun.text = value
        vrun.font.name = style.major_font or style.title_font
        vrun.font.size = Pt(56 if cols <= 2 else 40)
        vrun.font.bold = True
        vrun.font.color.rgb = palette[i % len(palette)]
        # Label
        ltb = slide.shapes.add_textbox(
            Inches(left), Inches(top + cell_h * 0.58),
            Inches(cell_w - 0.2), Inches(cell_h * 0.35),
        )
        ltf = ltb.text_frame
        ltf.word_wrap = True
        lp = ltf.paragraphs[0]
        lp.alignment = PP_ALIGN.CENTER
        lrun = lp.add_run()
        lrun.text = label
        lrun.font.name = style.body_font
        lrun.font.size = Pt(14)
        if style.body_color:
            lrun.font.color.rgb = style.body_color


def _split_comparison_lines(lines: list[str]) -> tuple[list[str], list[str]]:
    """Split the __LEFT__/__RIGHT__-prefixed lines produced by the parser
    back into two column lists."""
    left = [ln[len("__LEFT__"):] for ln in lines if ln.startswith("__LEFT__")]
    right = [ln[len("__RIGHT__"):] for ln in lines if ln.startswith("__RIGHT__")]
    return left, right


def _add_two_column_slide_from_template(
    prs: Presentation, style: TemplateStyle, sc: SlideContent, purpose: str,
):
    """Render a two-column layout (two_content or comparison). Prefers a
    template layout with two BODY placeholders; synthesizes text boxes on a
    simpler layout when none exists."""
    info, layout = _pick_layout(style, purpose, "content")
    slide = prs.slides.add_slide(layout)
    width_in, height_in = _slide_dims(prs)
    accent = style.accent_at(sc.accent_hint)

    left_lines, right_lines = _split_comparison_lines(sc.body_lines)
    if not left_lines and not right_lines:
        # LLM used `{layout=two_content}` without `---`; fall back to splitting
        # the bullets in half so the slide still renders with two columns.
        mid = len(sc.body_lines) // 2 or 1
        left_lines = sc.body_lines[:mid]
        right_lines = sc.body_lines[mid:]

    # Try to use two body placeholders if the layout advertises them.
    body_phs = [
        ph for ph in slide.placeholders
        if ph.placeholder_format.idx != 0 and ph.has_text_frame
    ]
    title_ph = next(
        (ph for ph in slide.placeholders if ph.placeholder_format.idx == 0),
        None,
    )

    if title_ph and sc.title:
        _set_placeholder_text(
            title_ph, sc.title,
            fallback_font=style.title_font,
            fallback_size=style.content_title_size or style.title_size,
            fallback_color=style.title_color,
        )

    if len(body_phs) >= 2:
        _fill_body_placeholder(body_phs[0], left_lines or [""], style)
        _fill_body_placeholder(body_phs[1], right_lines or [""], style)
    else:
        # Synthesize two text boxes side-by-side.
        if title_ph is None and sc.title:
            ttb = slide.shapes.add_textbox(
                Inches(0.7), Inches(0.4), Inches(width_in - 1.4), Inches(0.8),
            )
            ttf = ttb.text_frame
            ttf.word_wrap = True
            tp = ttf.paragraphs[0]
            tp.alignment = PP_ALIGN.LEFT
            trun = tp.add_run()
            trun.text = sc.title
            trun.font.name = style.title_font
            trun.font.size = Pt(24)
            trun.font.bold = True
            if style.title_color:
                trun.font.color.rgb = style.title_color

        col_top = 1.4
        col_w = (width_in - 1.8) / 2
        col_h = height_in - col_top - 0.4
        for col_idx, col_lines in enumerate((left_lines, right_lines)):
            left = 0.7 + col_idx * (col_w + 0.4)
            # Divider bar for comparison layouts — emphasizes the split.
            if purpose == "comparison" and col_idx == 0:
                div = slide.shapes.add_shape(
                    1,
                    Inches(left + col_w + 0.15),
                    Inches(col_top),
                    Inches(0.06),
                    Inches(col_h),
                )
                div.fill.solid()
                div.fill.fore_color.rgb = accent
                div.line.fill.background()
            tb = slide.shapes.add_textbox(
                Inches(left), Inches(col_top),
                Inches(col_w), Inches(col_h),
            )
            tf = tb.text_frame
            tf.word_wrap = True
            for i, line in enumerate(col_lines):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(6)
                run = p.add_run()
                run.text = f"•  {line}"
                run.font.name = style.body_font
                run.font.size = style.body_size
                if style.body_color:
                    run.font.color.rgb = style.body_color


def _add_section_slide_synthesized(
    prs: Presentation, style: TemplateStyle, sc: SlideContent,
):
    """When the template advertises no Section Header layout, synthesize one
    on the simplest available layout using the template's accent palette —
    this is the template-aware replacement for the hardcoded navy default
    path at _add_section_slide_default."""
    # Prefer any layout tagged section/blank/title in the catalog — we want
    # something with minimal placeholders so we can lay out our own.
    info = _choose_layout_by_hint(style.layout_catalog, "blank", "section")
    layout = prs.slide_layouts[info.idx]
    slide = prs.slides.add_slide(layout)
    width_in, height_in = _slide_dims(prs)
    accent = style.next_accent()

    # Clear any inherited placeholder text so only our shapes show.
    for ph in list(slide.placeholders):
        if ph.has_text_frame:
            ph.text_frame.clear()

    # Large accent bar.
    bar = slide.shapes.add_shape(
        1,
        Inches(0.8), Inches(height_in * 0.42),
        Inches(1.8), Inches(0.08),
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()

    # Section title, below the bar.
    tb = slide.shapes.add_textbox(
        Inches(0.8), Inches(height_in * 0.46),
        Inches(width_in - 1.6), Inches(height_in * 0.35),
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = sc.title
    run.font.name = style.major_font or style.title_font
    run.font.size = Pt(40)
    run.font.bold = True
    # Title color: prefer explicit title_color, else a dark theme color that
    # contrasts with likely backgrounds, else the accent itself.
    if style.title_color:
        run.font.color.rgb = style.title_color
    else:
        run.font.color.rgb = accent


def _add_standard_slide_from_template(
    prs: Presentation, style: TemplateStyle, sc: SlideContent, subtitle: str = "",
):
    """The original title/section/content renderer, adapted to use the
    layout catalog so it can pick any content-capable layout (not just the
    single content_layout_idx)."""
    if sc.slide_type == "title":
        info, layout = _pick_layout(style, sc.layout_hint or "title", "title")
        title_size = style.title_size
    elif sc.slide_type == "section":
        info, layout = _pick_layout(style, sc.layout_hint or "section", "section")
        title_size = style.section_title_size or style.title_size
    else:
        info, layout = _pick_layout(style, sc.layout_hint or "content", "content")
        title_size = style.content_title_size or style.title_size

    slide = prs.slides.add_slide(layout)

    # Fill placeholders.
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            _set_placeholder_text(
                ph, sc.title,
                fallback_font=style.title_font,
                fallback_size=title_size,
                fallback_color=style.title_color,
            )
        elif idx == 1:
            if sc.slide_type == "title" and subtitle:
                _set_placeholder_text(
                    ph, subtitle,
                    fallback_font=style.body_font,
                    fallback_size=style.body_size,
                    fallback_color=style.body_color,
                )
            elif sc.body_lines:
                _fill_body_placeholder(ph, sc.body_lines, style)

    # Fallbacks for minimal layouts that have no title/body placeholders.
    title_ph_found = any(ph.placeholder_format.idx == 0 for ph in slide.placeholders)
    body_ph_found = any(ph.placeholder_format.idx == 1 for ph in slide.placeholders)

    body_top = Inches(1.8)
    if not title_ph_found and sc.title and sc.slide_type != "section":
        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), Inches(8.6), Inches(1.0))
        ttf = title_box.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.LEFT
        trun = tp.add_run()
        trun.text = sc.title
        trun.font.name = style.title_font
        trun.font.size = title_size or Pt(28)
        trun.font.bold = True
        if style.title_color:
            trun.font.color.rgb = style.title_color
        body_top = Inches(1.6)

    if not body_ph_found and sc.body_lines and sc.slide_type == "content":
        txBox = slide.shapes.add_textbox(Inches(0.7), body_top, Inches(8.6), Inches(3.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        for i, line in enumerate(sc.body_lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = f"•  {line}"
            run.font.name = style.body_font
            run.font.size = style.body_size
            if style.body_color:
                run.font.color.rgb = style.body_color
            p.space_after = Pt(6)


def _add_slide_from_template(
    prs: Presentation, style: TemplateStyle, slide_content: SlideContent, subtitle: str = "",
):
    """Top-level dispatch for template-driven slides: reads the slide's
    layout_hint (from the LLM or parser heuristics), snaps it to an
    available layout, and routes to the appropriate creative renderer."""
    # Section slides: when the template advertises no real Section Header
    # layout, synthesize one using theme accent colors — this replaces the
    # old hardcoded navy default fallback.
    if slide_content.slide_type == "section":
        if style.section_layout_found:
            _add_standard_slide_from_template(prs, style, slide_content, subtitle)
        else:
            _add_section_slide_synthesized(prs, style, slide_content)
        return

    # Creative slide types produced by the parser / LLM.
    if slide_content.slide_type == "quote" or slide_content.quote_text:
        _add_quote_slide_from_template(prs, style, slide_content)
        return
    if slide_content.slide_type == "bignum" or slide_content.big_number is not None:
        _add_bignum_slide_from_template(prs, style, slide_content)
        return
    if slide_content.slide_type == "stat_grid" or slide_content.stats:
        _add_stat_grid_slide_from_template(prs, style, slide_content)
        return

    # Two-column layouts driven purely by hint.
    if slide_content.layout_hint in ("two_content", "comparison"):
        _add_two_column_slide_from_template(
            prs, style, slide_content, slide_content.layout_hint,
        )
        return

    # Everything else: standard title / content via the layout catalog.
    _add_standard_slide_from_template(prs, style, slide_content, subtitle)


# ─────────────────────────────────────────────────────────────
# 6. MAIN ORCHESTRATOR
# ─────────────────────────────────────────────────────────────

def generate_pptx(
    input_path: str,
    output_path: str,
    template_path: str | None = None,
    title: str | None = None,
    max_bullets: int = 7,
    max_table_rows: int = MAX_TABLE_ROWS_PER_SLIDE,
    use_llm: bool = True,
    ollama_host: str = DEFAULT_OLLAMA_HOST,
    ollama_model: str = DEFAULT_OLLAMA_MODEL,
    llm_prompt_file: str | None = None,
):
    """
    Main entry point: read a document file and produce a .pptx presentation.

    Args:
        input_path:    Path to the input document (.txt, .md, .docx, .pdf, .html, .xlsx, .csv)
        output_path:   Path for the generated .pptx file
        template_path: Optional path to a .pptx template for styling
        title:         Optional override for the presentation title
        max_bullets:   Maximum bullet points per slide before auto-splitting
        max_table_rows: Maximum data rows per table slide (xlsx/csv only)
        use_llm:       If True, rewrite extracted text with a local Ollama server
                       before parsing. Skipped automatically for spreadsheets and
                       when the server is unreachable.
        ollama_host:   Base URL of the Ollama server.
        ollama_model:  Model name to request from Ollama.
        llm_prompt_file: Optional path to a text file containing a custom system
                       prompt to override the default rewrite instructions.
    """
    ext = Path(input_path).suffix.lower()
    is_spreadsheet = ext in (".xlsx", ".xls", ".xlsm", ".csv", ".tsv")

    # Analyze the template up-front (when present) so its layout catalog
    # and accent palette are available to inject into the LLM prompt.
    template_style: TemplateStyle | None = None
    template_kit: str | None = None
    if template_path:
        print(f"🎨 Analyzing template {template_path}...")
        template_style = analyze_template(template_path)
        template_kit = render_template_kit(template_style)
        purposes = {p for li in template_style.layout_catalog for p in li.purposes}
        print(f"   Layouts: {len(template_style.layout_catalog)} | "
              f"Purposes: {', '.join(sorted(purposes)) or '(none inferred)'} | "
              f"Accents: {len(template_style.accent_palette)}")

    # 1. Read & parse
    print(f"📄 Reading {input_path}...")

    if is_spreadsheet:
        deck = read_xlsx(input_path, deck_title=title or "", max_rows_per_slide=max_table_rows)
    else:
        raw_text = read_document(input_path)
        if not raw_text.strip():
            raise ValueError(f"No text content could be extracted from '{input_path}'.")

        if use_llm:
            print(f"🧠 Rewriting with Ollama ({ollama_model}) at {ollama_host}...")
            try:
                system_prompt = None
                if llm_prompt_file:
                    system_prompt = Path(llm_prompt_file).read_text(encoding="utf-8")
                rewritten = rewrite_for_pptx(
                    raw_text,
                    host=ollama_host,
                    model=ollama_model,
                    system_prompt=system_prompt,
                    template_kit=template_kit,
                )
                print("raw_text:\n", raw_text)
                print("rewritten:\n", rewritten)
                if rewritten:
                    raw_text = rewritten
                    print("   LLM rewrite applied.")
                else:
                    print("   ⚠️  LLM returned empty output; using original text.")
            except Exception as exc:
                print(f"   ⚠️  LLM rewrite skipped ({exc}); using original text.")

        print("🔍 Parsing document structure...")
        deck = parse_text_to_deck(raw_text, deck_title=title or "", max_bullets=max_bullets)

    if not deck.slides:
        raise ValueError("Could not extract any meaningful content for slides.")

    table_count = sum(1 for s in deck.slides if s.slide_type == "table")
    other_count = len(deck.slides) - table_count
    parts = []
    if other_count:
        parts.append(f"{other_count} content")
    if table_count:
        parts.append(f"{table_count} table")
    print(f"   Found {len(deck.slides)} slides ({', '.join(parts)})")

    # 2. Build the presentation
    if template_style is not None:
        style = template_style
        prs = style.presentation

        # Remove any existing slides from the template
        for _ in range(len(prs.slides)):
            sldId = prs.slides._sldIdLst[0]
            rId = sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rId is None:
                for attr_key, attr_val in sldId.attrib.items():
                    if attr_key.endswith('}id') or attr_key == 'r:id':
                        rId = attr_val
                        break
            if rId:
                try:
                    prs.part.drop_rel(rId)
                except KeyError:
                    pass
            prs.slides._sldIdLst.remove(sldId)

        print("   Generating slides with template styling...")
        for sc in deck.slides:
            if sc.slide_type == "table":
                _add_table_slide_from_template(prs, style, sc)
            else:
                _add_slide_from_template(prs, style, sc, subtitle=deck.subtitle)
    else:
        print("🎨 Using default styling...")
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)

        for sc in deck.slides:
            if sc.slide_type == "title":
                _add_title_slide_default(prs, sc, subtitle=deck.subtitle)
            elif sc.slide_type == "section":
                _add_section_slide_default(prs, sc)
            elif sc.slide_type == "table":
                _add_table_slide_default(prs, sc)
            else:
                # Content slide — also handles quote/bignum/stat_grid in the
                # default (no-template) path. Without a template palette we
                # can't meaningfully render those as distinct creative slides,
                # so we fold their extracted content back into bullets so the
                # default renderer can still show it.
                if sc.quote_text:
                    sc.body_lines = [f"\u201c{sc.quote_text}\u201d"] + sc.body_lines
                if sc.big_number:
                    sc.body_lines = [f"{sc.big_number[0]} — {sc.big_number[1]}"] + sc.body_lines
                if sc.stats:
                    sc.body_lines = [f"{v} — {l}" for v, l in sc.stats] + sc.body_lines
                _add_content_slide_default(prs, sc)

    # 3. Save
    prs.save(output_path)
    print(f"✅ Saved to {output_path}")
    print(f"   {len(deck.slides)} slides generated")
    return output_path


# ─────────────────────────────────────────────────────────────
# 7. CLI
# ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert documents (txt, md, docx, pdf, html, xlsx, csv) to PowerPoint presentations.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python doc2pptx.py notes.md -o presentation.pptx
  python doc2pptx.py report.pdf -o deck.pptx --template brand.pptx
  python doc2pptx.py article.docx -o slides.pptx --title "Q4 Report"
  python doc2pptx.py page.html -o slides.pptx --max-bullets 5
  python doc2pptx.py data.xlsx -o tables.pptx
  python doc2pptx.py data.csv -o tables.pptx --max-table-rows 10
  python doc2pptx.py report.pdf -o deck.pptx --no-llm
  python doc2pptx.py report.pdf -o deck.pptx --ollama-model qwen2.5
        """,
    )
    parser.add_argument("input", help="Input document path (.txt, .md, .docx, .pdf, .html, .xlsx, .csv, .tsv)")
    parser.add_argument("-o", "--output", default="output.pptx", help="Output .pptx path (default: output.pptx)")
    parser.add_argument("-t", "--template", default=None, help="Template .pptx to use for styling")
    parser.add_argument("--title", default=None, help="Override the presentation title")
    parser.add_argument("--max-bullets", type=int, default=7, help="Max bullet points per slide before splitting (default: 7)")
    parser.add_argument("--max-table-rows", type=int, default=MAX_TABLE_ROWS_PER_SLIDE,
                        help=f"Max data rows per table slide (default: {MAX_TABLE_ROWS_PER_SLIDE})")
    parser.add_argument("--no-llm", action="store_true",
                        help="Skip the Ollama rewrite step (LLM rewrite is on by default).")
    parser.add_argument("--ollama-host", default=DEFAULT_OLLAMA_HOST,
                        help=f"Ollama server base URL (default: {DEFAULT_OLLAMA_HOST}, env: OLLAMA_HOST)")
    parser.add_argument("--ollama-model", default=DEFAULT_OLLAMA_MODEL,
                        help=f"Ollama model name (default: {DEFAULT_OLLAMA_MODEL}, env: OLLAMA_MODEL)")
    parser.add_argument("--prompt-file", default=None,
                        help="Path to a text file containing a custom system prompt for the rewrite step.")

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.template and not os.path.exists(args.template):
        print(f"Error: Template file not found: {args.template}", file=sys.stderr)
        sys.exit(1)

    if args.prompt_file and not os.path.exists(args.prompt_file):
        print(f"Error: Prompt file not found: {args.prompt_file}", file=sys.stderr)
        sys.exit(1)

    generate_pptx(
        input_path=args.input,
        output_path=args.output,
        template_path=args.template,
        title=args.title,
        max_bullets=args.max_bullets,
        max_table_rows=args.max_table_rows,
        use_llm=not args.no_llm,
        ollama_host=args.ollama_host,
        ollama_model=args.ollama_model,
        llm_prompt_file=args.prompt_file,
    )


if __name__ == "__main__":
    main()
