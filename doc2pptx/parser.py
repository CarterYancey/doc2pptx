"""Parse raw/rewritten text into a structured DeckContent."""

from __future__ import annotations

import re

from .models import DeckContent, SlideContent


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
    # TODO: Make this function better
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
    # TODO: Sometimes, extracting text from pdfs causes line spacing to be compressed, so this may not apply
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
