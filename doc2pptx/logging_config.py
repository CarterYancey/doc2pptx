"""Logging configuration for doc2pptx."""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path


logger = logging.getLogger("doc2pptx")


def _default_log_path(input_path: str | None) -> Path:
    """Build a timestamped default log path under ./logs/."""
    stem = Path(input_path).stem if input_path else "run"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path("logs") / f"doc2pptx_{ts}_{stem}.log"


def configure_logging(
    log_file: str | Path | None = None,
    verbose: bool = False,
    quiet: bool = False,
) -> Path | None:
    """Configure terminal + optional file logging for the `doc2pptx` logger.

    Terminal: concise, INFO by default (WARNING if quiet, DEBUG if verbose).
    File (when `log_file` is set): detailed, DEBUG always — captures full
    document extraction previews and full LLM prompt/input/response exchanges.

    Returns the resolved log file path (or None if file logging disabled).
    """
    # Clear any handlers attached to our namespace from a previous run (Gradio
    # keeps the process alive and would otherwise duplicate output).
    for h in list(logger.handlers):
        logger.removeHandler(h)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    term_level = logging.WARNING if quiet else (logging.DEBUG if verbose else logging.INFO)
    term = logging.StreamHandler(stream=sys.stderr)
    term.setLevel(term_level)
    term.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(term)

    resolved: Path | None = None
    if log_file is not None:
        resolved = Path(log_file)
        resolved.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(resolved, mode="w", encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)-5s %(message)s",
            datefmt="%H:%M:%S",
        ))
        logger.addHandler(fh)
        logger.debug("Log file opened at %s", resolved)
    return resolved


def _preview(text: str, n: int = 240) -> str:
    """Single-line preview of a longer string, for INFO-level logs."""
    text = (text or "").strip().replace("\n", " ⏎ ")
    if len(text) <= n:
        return text
    return text[:n] + f"… [+{len(text) - n} chars]"
