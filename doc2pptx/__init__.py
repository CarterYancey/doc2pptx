"""doc2pptx — Convert documents (txt, md, docx, pdf) into PowerPoint presentations.

Optionally accepts an existing .pptx as a style template to inherit backgrounds,
color themes, fonts, and slide layouts.
"""

from .llm import (
    DEFAULT_OLLAMA_HOST,
    DEFAULT_OLLAMA_MODEL,
    DEFAULT_REWRITE_PROMPT,
    rewrite_for_pptx,
    rewrite_for_pptx_chunked,
)
from .logging_config import configure_logging, logger
from .models import DeckContent, SlideContent, TemplateStyle
from .parser import parse_text_to_deck
from .pipeline import generate_pptx
from .readers import read_document, read_xlsx
from .template import analyze_template


__all__ = [
    "DEFAULT_OLLAMA_HOST",
    "DEFAULT_OLLAMA_MODEL",
    "DEFAULT_REWRITE_PROMPT",
    "DeckContent",
    "SlideContent",
    "TemplateStyle",
    "analyze_template",
    "configure_logging",
    "generate_pptx",
    "logger",
    "parse_text_to_deck",
    "read_document",
    "read_xlsx",
    "rewrite_for_pptx",
    "rewrite_for_pptx_chunked",
]
