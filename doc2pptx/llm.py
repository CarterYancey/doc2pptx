"""LLM rewrite (Ollama) — prose → PPT-friendly markdown, with chunking."""

from __future__ import annotations

import os
import re
import time

from .logging_config import logger


DEFAULT_OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
DEFAULT_OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "gemma4:26b")

DEFAULT_REWRITE_PROMPT = """You are rewriting a document so it converts cleanly into a PowerPoint deck.

Rules:
- Output ONLY markdown. No preamble, no commentary, no code fences.
- Use `#` for the deck title, `##` for each new section title, `###` for slide headings. There should always be at least 2 slides per section, often more; group as many slides together under the same section as reasonably possible.
- Under each slide title, write 5-7 bullet points starting with `- `.
- Preserve the original meaning, facts, numbers, and ordering. Do not invent content.
- Do not leave any content out, but do summarize ideas instead of repeating verbatim.
- Drop filler, repetition, and boilerplate.
"""

CONTINUATION_REWRITE_PROMPT = """You are rewriting part of a longer document so it converts cleanly into a PowerPoint deck. This chunk is a CONTINUATION of an earlier rewrite.

Rules:
- Output ONLY markdown. No preamble, no commentary, no code fences.
- Do NOT emit a `#` deck title — the deck title was already produced. Start directly with a `##` slide title.
- Use `##` for each new major section title and `###` for slide headers. There should always be at least 2 slides per section, often more; group as many slides together under the same section as reasonably possible.
- Under each slide title, write 5-7 bullet points starting with `- `.
- Preserve the original meaning, facts, numbers, and ordering. Do not invent content.
- Do not leave any important content out, but do not repeat verbatim unless a direct quote is needed for clarity or completeness.
- Drop filler, repetition, boilerplate, etc.
"""


def rewrite_for_pptx(
    raw_text: str,
    host: str = DEFAULT_OLLAMA_HOST,
    model: str = DEFAULT_OLLAMA_MODEL,
    system_prompt: str | None = None,
    timeout: float = 240.0,
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
    logger.debug("POST %s (model=%s, prompt=%d chars, system=%d chars)",
                 url, model, len(raw_text), len(payload["system"]))
    with httpx.Client(timeout=timeout) as client:
        resp = client.post(url, json=payload)
        resp.raise_for_status()
        data = resp.json()
    return (data.get("response") or "").strip()


_HEADING_LINE_RE = re.compile(r"^#{1,4}\s+\S")


def _split_into_blocks(text: str) -> list[str]:
    """Split text into atomic blocks for greedy chunk packing.

    A block is either:
      - a heading line (starts with 1-4 `#` followed by whitespace and text), or
      - a paragraph (run of non-blank lines separated by blank lines).

    Blocks preserve their leading/trailing newlines so re-joining with `\\n\\n`
    produces readable text.
    """
    blocks: list[str] = []
    current: list[str] = []

    def flush() -> None:
        if current:
            joined = "\n".join(current).strip()
            if joined:
                blocks.append(joined)
            current.clear()

    for line in text.splitlines():
        if _HEADING_LINE_RE.match(line):
            # Headings are their own block.
            flush()
            blocks.append(line.strip())
        elif not line.strip():
            flush()
        else:
            current.append(line)
    flush()
    return blocks


def _split_oversized_block(block: str, max_chunk_chars: int) -> list[str]:
    """Split a single block that is larger than max_chunk_chars on sentence
    boundaries. Falls back to hard character slicing if sentences are still
    too long.
    """
    if len(block) <= max_chunk_chars:
        return [block]
    sentences = re.split(r"(?<=[.!?])\s+", block)
    pieces: list[str] = []
    buf = ""
    for s in sentences:
        if not s:
            continue
        if len(s) > max_chunk_chars:
            # Emit whatever is buffered, then hard-slice the long sentence.
            if buf:
                pieces.append(buf)
                buf = ""
            for i in range(0, len(s), max_chunk_chars):
                pieces.append(s[i : i + max_chunk_chars])
            continue
        candidate = (buf + " " + s).strip() if buf else s
        if len(candidate) > max_chunk_chars:
            pieces.append(buf)
            buf = s
        else:
            buf = candidate
    if buf:
        pieces.append(buf)
    return pieces


def _tail_sentences(text: str, n: int) -> str:
    """Return the last `n` sentences of text for use as cross-chunk overlap."""
    if n <= 0 or not text.strip():
        return ""
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    sentences = [s for s in sentences if s]
    if not sentences:
        return ""
    return " ".join(sentences[-n:]).strip()


def chunk_document_for_llm(
    text: str,
    max_chunk_chars: int = 6000,
    overlap_sentences: int = 2,
) -> list[str]:
    """Split `text` into structure-aware chunks for LLM rewriting.

    Splitting priority: markdown headings (`#`..`####`) → blank-line paragraphs
    → sentences (only when a single paragraph exceeds `max_chunk_chars`).

    The last `overlap_sentences` sentences of chunk N are prepended to chunk
    N+1 as orientation context. The overlap is labelled so the model knows not
    to reproduce it. Always returns at least one chunk.
    """
    if max_chunk_chars <= 0:
        raise ValueError("max_chunk_chars must be positive")
    text = text or ""
    if not text.strip():
        return [text]

    blocks = _split_into_blocks(text)
    if not blocks:
        return [text]

    # Pre-split any block larger than the budget.
    expanded: list[str] = []
    for b in blocks:
        if len(b) > max_chunk_chars:
            expanded.extend(_split_oversized_block(b, max_chunk_chars))
        else:
            expanded.append(b)

    chunks: list[str] = []
    current_parts: list[str] = []
    current_len = 0
    for block in expanded:
        extra = len(block) + (2 if current_parts else 0)  # "\n\n" separator
        if current_parts and current_len + extra > max_chunk_chars:
            chunks.append("\n\n".join(current_parts))
            current_parts = [block]
            current_len = len(block)
        else:
            current_parts.append(block)
            current_len += extra
    if current_parts:
        chunks.append("\n\n".join(current_parts))

    if len(chunks) <= 1 or overlap_sentences <= 0:
        final_chunks = chunks
    else:
        # Prepend overlap to every chunk after the first.
        with_overlap: list[str] = [chunks[0]]
        for i in range(1, len(chunks)):
            tail = _tail_sentences(chunks[i - 1], overlap_sentences)
            if tail:
                header = f"[Context from previous section — do not repeat in output]\n{tail}"
                with_overlap.append(f"{header}\n\n{chunks[i]}")
            else:
                with_overlap.append(chunks[i])
        final_chunks = with_overlap

    # Log distribution so the user can decide whether to tune max_chunk_chars.
    sizes = [len(c) for c in final_chunks]
    if sizes:
        logger.info(
            "   Chunking: %d chunk(s) — size min=%d / avg=%d / max=%d (budget=%d, overlap=%d)",
            len(sizes), min(sizes), sum(sizes) // len(sizes), max(sizes),
            max_chunk_chars, overlap_sentences,
        )
        for i, c in enumerate(final_chunks, start=1):
            logger.debug(
                "--- CHUNK %d/%d (%d chars) ---\n%s\n--- END CHUNK %d ---",
                i, len(final_chunks), len(c), c, i,
            )
    return final_chunks


def _strip_deck_title_lines(markdown: str) -> str:
    """Remove any `# ...` deck-title lines. Used on continuation-chunk output
    as a safety net — the deck title must appear only once.
    """
    kept = [ln for ln in markdown.splitlines() if not re.match(r"^#\s+\S", ln)]
    return "\n".join(kept)


def _derive_continuation_prompt(base_prompt: str) -> str:
    """Given a (possibly user-customized) base prompt, produce a continuation
    variant that forbids emitting a new `#` deck title.
    """
    if base_prompt.strip() == DEFAULT_REWRITE_PROMPT.strip():
        return CONTINUATION_REWRITE_PROMPT
    addendum = (
        "\n\nIMPORTANT — this input is a CONTINUATION of an earlier rewrite: "
        "do NOT emit a `#` deck title. Start directly with a `##` slide title. "
        "A short excerpt from the previous section may appear at the top of the "
        "input as context — do not repeat it in the output."
    )
    return base_prompt.rstrip() + addendum


def _format_eta(seconds: float) -> str:
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"{seconds}s"
    if seconds < 3600:
        return f"{seconds // 60}m{seconds % 60:02d}s"
    h, rem = divmod(seconds, 3600)
    return f"{h}h{rem // 60:02d}m"


def rewrite_for_pptx_chunked(
    raw_text: str,
    host: str = DEFAULT_OLLAMA_HOST,
    model: str = DEFAULT_OLLAMA_MODEL,
    system_prompt: str | None = None,
    timeout: float = 240.0,
    max_chunk_chars: int = 6000,
    overlap_sentences: int = 2,
) -> str:
    """Chunk `raw_text`, rewrite each chunk via Ollama, and stitch results.

    The first chunk uses `system_prompt` (or `DEFAULT_REWRITE_PROMPT`).
    Subsequent chunks use a continuation variant that forbids a new `#` deck
    title. If any single chunk fails, its raw text is substituted so the
    pipeline degrades gracefully (mirroring the single-pass failure mode).
    """
    chunks = chunk_document_for_llm(
        raw_text,
        max_chunk_chars=max_chunk_chars,
        overlap_sentences=overlap_sentences,
    )

    base_prompt = system_prompt or DEFAULT_REWRITE_PROMPT
    continuation_prompt = _derive_continuation_prompt(base_prompt)
    total = len(chunks)

    logger.info("   Rewriting with %s at %s (%d chunk(s))", model, host, total)
    logger.debug("=== BASE SYSTEM PROMPT ===\n%s\n=== END BASE SYSTEM PROMPT ===", base_prompt)
    if total > 1:
        logger.debug(
            "=== CONTINUATION SYSTEM PROMPT ===\n%s\n=== END CONTINUATION SYSTEM PROMPT ===",
            continuation_prompt,
        )

    rewritten_parts: list[str] = []
    durations: list[float] = []
    wall_start = time.monotonic()
    for i, chunk in enumerate(chunks, start=1):
        role = "base" if i == 1 else "continuation"
        prompt = base_prompt if i == 1 else continuation_prompt
        logger.info("   [chunk %d/%d] sending %d chars (%s)…", i, total, len(chunk), role)
        logger.debug(
            "=== LLM REQUEST chunk %d/%d (%s) ===\n--- USER INPUT (%d chars) ---\n%s\n=== END REQUEST chunk %d ===",
            i, total, role, len(chunk), chunk, i,
        )

        t0 = time.monotonic()
        failed = False
        try:
            out = rewrite_for_pptx(
                chunk,
                host=host,
                model=model,
                system_prompt=prompt,
                timeout=timeout,
            )
        except Exception as exc:
            failed = True
            logger.warning(
                "   [chunk %d/%d] rewrite failed (%s); using raw chunk text.",
                i, total, exc,
            )
            logger.debug("=== LLM ERROR chunk %d/%d ===\n%r", i, total, exc)
            out = chunk
        dt = time.monotonic() - t0
        durations.append(dt)

        if not failed and not out.strip():
            logger.warning(
                "   [chunk %d/%d] model returned empty output; using raw chunk text.",
                i, total,
            )
            out = chunk

        if i > 1:
            before = out
            out = _strip_deck_title_lines(out)
            if out != before:
                logger.debug("   [chunk %d/%d] stripped stray `#` deck-title line(s) from output.", i, total)

        # Progress + ETA based on rolling average.
        avg = sum(durations) / len(durations)
        remaining = (total - i) * avg
        elapsed = time.monotonic() - wall_start
        logger.info(
            "   [chunk %d/%d] ← %d chars in %.1fs  (avg %.1fs/chunk, elapsed %s, ETA %s)",
            i, total, len(out), dt, avg,
            _format_eta(elapsed), _format_eta(remaining) if total > i else "done",
        )
        logger.debug(
            "=== LLM RESPONSE chunk %d/%d (%.2fs, %d chars) ===\n%s\n=== END RESPONSE chunk %d ===",
            i, total, dt, len(out), out, i,
        )

        rewritten_parts.append(out.strip())

    total_elapsed = time.monotonic() - wall_start
    logger.info("   Rewrite complete: %d chunk(s) in %s", total, _format_eta(total_elapsed))
    return "\n\n".join(p for p in rewritten_parts if p)
