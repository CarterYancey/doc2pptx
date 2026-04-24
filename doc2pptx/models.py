"""Data classes used across the doc2pptx pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt


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
    background_fill: Any = None
