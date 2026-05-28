"""pptx-agent-toolkit — Agent-native PPTX engineering library.

22 layouts, 4 brand palettes, WCAG-AA color safety, Chinese font support.
Built for AI agent pipelines, not template filling.
"""

# Builder: core PPTX generation
from pptx_agent_toolkit.builder import (
    init_prs, new_slide, hex_to_rgb, resolve_palette,
    PALETTES,
    add_tb, set_p, add_p, add_footer,
    add_rect, add_rrect, add_oval, place_image,
)

# Layout enrichment: analyze & improve existing decks
from pptx_agent_toolkit.layout_enricher import (
    LAYOUT_CATALOG,
    SlideProfile, RichnessScore, ImageNeed, LayoutRecommendation,
    EnrichmentPlan, EnrichmentEngine,
)

# Color QA: WCAG-AA contrast auditing
from pptx_agent_toolkit.color_qa import ColorQA

# Color utilities: contrast math, palette validation
from pptx_agent_toolkit.color_utils import (
    contrast_ratio, wcag_level, is_aa_pass, is_aaa_pass,
    find_contrast_fix, palette_aa_check, rgb_distance,
    relative_luminance, hex_to_rgb as color_hex_to_rgb, rgb_to_hex,
)

# Image tools
from pptx_agent_toolkit.image_extractor import (
    extract_from_pptx, extract_from_docx, extract_from_pdf,
)
from pptx_agent_toolkit.image_insert import insert_image, process_batch

# Layout engine: alternative construction API
from pptx_agent_toolkit.layout_engine import LayoutEngine

# PPTX XML utilities (low-level, for advanced users)
from pptx_agent_toolkit.pptx_utils import (
    get_text_color, get_shape_fill_color, is_picture_shape,
    get_slide_background, rects_overlap, find_bg_for_text,
)

__version__ = "0.1.0"

__all__ = [
    # Builder
    "init_prs", "new_slide", "hex_to_rgb", "resolve_palette", "PALETTES",
    "add_tb", "set_p", "add_p", "add_footer",
    "add_rect", "add_rrect", "add_oval", "place_image",
    # Layout
    "LAYOUT_CATALOG", "EnrichmentEngine", "EnrichmentPlan",
    "SlideProfile", "RichnessScore", "ImageNeed", "LayoutRecommendation",
    # Color QA
    "ColorQA",
    # Color utils
    "contrast_ratio", "wcag_level", "is_aa_pass", "is_aaa_pass",
    "find_contrast_fix", "palette_aa_check", "rgb_distance",
    "relative_luminance", "color_hex_to_rgb", "rgb_to_hex",
    # Images
    "extract_from_pptx", "extract_from_docx", "extract_from_pdf",
    "insert_image", "process_batch",
    # Layout engine
    "LayoutEngine",
    # PPTX utils
    "get_text_color", "get_shape_fill_color", "is_picture_shape",
    "get_slide_background", "rects_overlap", "find_bg_for_text",
]
