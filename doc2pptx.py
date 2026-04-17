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
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ─────────────────────────────────────────────────────────────
# 1. DATA MODEL — Intermediate representation of slide content
# ─────────────────────────────────────────────────────────────

@dataclass
class SlideContent:
    """One logical slide."""
    title: str = ""
    body_lines: list[str] = field(default_factory=list)
    level: int = 0  # heading depth (1 = H1/title, 2 = H2/section, etc.)
    slide_type: str = "content"  # "title", "section", "content", "table"
    # Table data (used when slide_type == "table")
    table_headers: list[str] = field(default_factory=list)
    table_rows: list[list[str]] = field(default_factory=list)


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

DEFAULT_REWRITE_PROMPT = """You are rewriting a document so it converts cleanly into a PowerPoint deck.

Rules:
- Output ONLY markdown. No preamble, no commentary, no code fences.
- Use `#` for the deck title, `##` for each slide title, `###` for sub-sections.
- Under each slide title, write 3-7 bullet points starting with `- `.
- Preserve the original meaning, facts, numbers, and ordering. Do not invent content.
- Do not leave any content out, but do summarize ideas instead of repeating verbatim.
- Drop filler, repetition, and boilerplate.
"""


def rewrite_for_pptx(
    raw_text: str,
    host: str = DEFAULT_OLLAMA_HOST,
    model: str = DEFAULT_OLLAMA_MODEL,
    system_prompt: str | None = None,
    timeout: float = 120.0,
) -> str:
    """Ask a local Ollama server to reshape raw_text into PPT-friendly markdown.

    Returns the rewritten markdown. Raises on HTTP/connection errors — callers
    decide whether to fall back to the original text.
    """
    import httpx

    payload = {
        "model": model,
        "prompt": raw_text,
        "system": system_prompt or DEFAULT_REWRITE_PROMPT,
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

    def flush_slide():
        nonlocal current_slide, body_buffer
        if current_slide is not None:
            current_slide.body_lines = body_buffer
            # Drop the empty H2 companion content slide when no bullets
            # followed the section divider (avoids duplicate titles).
            if (
                current_slide.level == 2
                and current_slide.slide_type == "content"
                and not body_buffer
            ):
                current_slide = None
                body_buffer = []
                return
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
                    )
                    deck.slides.append(s)
        current_slide = None
        body_buffer = []

    first_h1_seen = False
    for line in lines:
        stripped = line.strip()
        heading = _is_heading(stripped)

        if heading:
            level, text = heading
            flush_slide()
            if level == 1:
                if not first_h1_seen:
                    first_h1_seen = True
                    deck.title = deck.title or text
                    current_slide = SlideContent(title=text, level=1, slide_type="title")
                else:
                    current_slide = SlideContent(title=text, level=1, slide_type="section")
            elif level == 2:
                # H2 is always a section divider. Emit the section slide
                # immediately, then start a companion content slide with the
                # same title so any bullets before the next heading spill
                # onto it. If no bullets arrive, flush_slide drops the empty
                # companion to avoid a duplicate title slide.
                deck.slides.append(
                    SlideContent(title=text, level=2, slide_type="section")
                )
                current_slide = SlideContent(title=text, level=2, slide_type="content")
            else:
                current_slide = SlideContent(title=text, level=level, slide_type="content")
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
class TemplateStyle:
    """Style information extracted from a template presentation."""
    presentation: Presentation
    # Layout indices by purpose
    title_layout_idx: int = 0
    section_layout_idx: int = 0
    content_layout_idx: int = 1
    # Whether a real Section Header layout was found (vs. a fallback).
    # When False, section slides use the default-style renderer so they
    # still look distinct from plain content slides.
    section_layout_found: bool = False
    # Fonts detected from template
    title_font: str = "Calibri"
    body_font: str = "Calibri"
    title_size: Pt = field(default_factory=lambda: Pt(36))
    section_title_size: Pt | None = None
    content_title_size: Pt | None = None
    body_size: Pt = field(default_factory=lambda: Pt(16))
    # Colors
    title_color: RGBColor | None = None
    body_color: RGBColor | None = None
    background_fill: any = None


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
    """Load a .pptx template and extract its style information."""
    prs = Presentation(template_path)
    font_info = _extract_font_info(prs)

    title_idx, _ = _find_best_layout(prs, "title")
    section_idx, section_found = _find_best_layout(prs, "section")
    content_idx, _ = _find_best_layout(prs, "content")

    layouts = prs.slide_layouts
    section_title_size = (
        _layout_title_size(layouts[section_idx]) if section_found else None
    )
    content_title_size = _layout_title_size(layouts[content_idx])

    style = TemplateStyle(
        presentation=prs,
        title_layout_idx=title_idx,
        section_layout_idx=section_idx,
        content_layout_idx=content_idx,
        section_layout_found=section_found,
        title_font=font_info["title_font"],
        body_font=font_info["body_font"],
        title_size=font_info["title_size"],
        section_title_size=section_title_size,
        content_title_size=content_title_size,
        body_size=font_info["body_size"],
        title_color=font_info["title_color"],
        body_color=font_info["body_color"],
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

    # Derive header background from template title color or fall back to navy
    header_bg = style.title_color or RGBColor(0x1E, 0x27, 0x61)

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


def _add_slide_from_template(prs: Presentation, style: TemplateStyle, slide_content: SlideContent, subtitle: str = ""):
    """Create a slide using the template's layouts and styles."""
    # Section slides without a real Section Header layout would be
    # indistinguishable from content slides; fall back to the default
    # renderer so they still look like section dividers.
    if slide_content.slide_type == "section" and not style.section_layout_found:
        _add_section_slide_default(prs, slide_content)
        return

    layouts = prs.slide_layouts

    if slide_content.slide_type == "title":
        layout_idx = style.title_layout_idx
        title_size = style.title_size
    elif slide_content.slide_type == "section":
        layout_idx = style.section_layout_idx
        title_size = style.section_title_size or style.title_size
    else:
        layout_idx = style.content_layout_idx
        title_size = style.content_title_size or style.title_size

    layout_idx = min(layout_idx, len(layouts) - 1)
    slide = prs.slides.add_slide(layouts[layout_idx])

    # Fill placeholders
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:  # Title placeholder
            _set_placeholder_text(
                ph, slide_content.title,
                fallback_font=style.title_font,
                fallback_size=title_size,
                fallback_color=style.title_color,
            )
        elif idx == 1:  # Body/subtitle placeholder
            if slide_content.slide_type == "title" and subtitle:
                _set_placeholder_text(
                    ph, subtitle,
                    fallback_font=style.body_font,
                    fallback_size=style.body_size,
                    fallback_color=style.body_color,
                )
            elif slide_content.body_lines:
                _fill_body_placeholder(ph, slide_content.body_lines, style)

    # If no body placeholder was found but we have body lines, add a text box
    body_ph_found = any(ph.placeholder_format.idx == 1 for ph in slide.placeholders)
    if not body_ph_found and slide_content.body_lines and slide_content.slide_type == "content":
        txBox = slide.shapes.add_textbox(Inches(0.7), Inches(1.8), Inches(8.6), Inches(3.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        for i, line in enumerate(slide_content.body_lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            run = p.add_run()
            run.text = f"•  {line}"
            run.font.name = style.body_font
            run.font.size = style.body_size
            if style.body_color:
                run.font.color.rgb = style.body_color
            p.space_after = Pt(6)


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
    if template_path:
        print(f"🎨 Loading template from {template_path}...")
        style = analyze_template(template_path)
        prs = Presentation(template_path)

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
