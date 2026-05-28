"""
Demo: Build a 5-slide presentation with pptx-agent-toolkit.

Demonstrates: palette selection, WCAG-AA validation, all layout types.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from pptx_agent_toolkit.builder import (
    init_prs, new_slide, add_tb, set_p, add_p,
    add_rect, add_rrect, PALETTES, resolve_palette,
)
from pptx_agent_toolkit.color_utils import palette_aa_check

OUTPUT = Path(__file__).parent / "demo_output.pptx"

config = {"palette": "premium"}
palette = dict(PALETTES["premium"])
warnings = palette_aa_check(palette)
if warnings:
    print("[!] WCAG warnings:")
    for w in warnings:
        print(f"  - {w}")
else:
    print("[OK] All color pairs pass WCAG AA")

prs, C, font = init_prs(config)

# --- Slide 1: Cover ---
slide = new_slide(prs, C)
add_rect(slide, 0, 0.6, 13.333, 5.0, C["primary"])
tb = add_tb(slide, 1.2, 1.6, 10.0, 1.2)
set_p(tb.text_frame.paragraphs[0], "pptx-agent-toolkit", C["white"], 40, bold=True, font=font)
tb2 = add_tb(slide, 1.2, 3.0, 10.0, 0.6)
set_p(tb2.text_frame.paragraphs[0], "22 Layouts . 4 Palettes . WCAG-AA Safe . CJK Native", C["white"], 14, font=font)
tb3 = add_tb(slide, 1.2, 3.7, 10.0, 0.5)
set_p(tb3.text_frame.paragraphs[0], "Built for AI Agent Pipelines", C["accent"], 12, font=font)

# --- Slide 2: Feature Overview ---
slide = new_slide(prs, C)
tb = add_tb(slide, 0.8, 0.4, 11.0, 0.6)
set_p(tb.text_frame.paragraphs[0], "What's Inside", C["primary"], 28, bold=True, font=font)
items = [
    "22 unified layouts covering all presentation needs",
    "4 brand palettes: premium, modern, dark, warm",
    "WCAG-AA color safety with auto-fix suggestions",
    "Native CJK font support out of the box",
    "Pure-function API designed for AI agent construction",
]
for i, item in enumerate(items):
    tb = add_tb(slide, 1.0, 1.3 + i * 0.65, 11.3, 0.55)
    set_p(tb.text_frame.paragraphs[0], item, C["text"], 14, font=font)

# --- Slide 3: Layout Types ---
slide = new_slide(prs, C)
tb = add_tb(slide, 0.8, 0.4, 11.0, 0.6)
set_p(tb.text_frame.paragraphs[0], "Layout Categories", C["primary"], 28, bold=True, font=font)
groups = [
    ("Structural", ["cover", "section-header", "closing"]),
    ("Content", ["content", "cards", "side-by-side", "comparison"]),
    ("Visual", ["hero", "gallery", "process", "key-takeaways"]),
]
y = 1.3
for cat, layouts in groups:
    tb = add_tb(slide, 1.0, y, 3.0, 0.4)
    set_p(tb.text_frame.paragraphs[0], cat, C["accent"], 16, bold=True, font=font)
    for j, lay in enumerate(layouts):
        tb = add_tb(slide, 3.5, y + j * 0.4, 8.5, 0.35)
        set_p(tb.text_frame.paragraphs[0], f"- {lay}", C["muted"], 12, font=font)
    y += 2.0

# --- Slide 4: Palette Preview ---
slide = new_slide(prs, C)
tb = add_tb(slide, 0.8, 0.4, 11.0, 0.6)
set_p(tb.text_frame.paragraphs[0], "Brand Palettes", C["primary"], 28, bold=True, font=font)
for i, (name, pal) in enumerate(PALETTES.items()):
    x = 0.8 + i * 3.2
    tb = add_tb(slide, x, 1.3, 2.8, 0.4)
    set_p(tb.text_frame.paragraphs[0], name, C["text"], 14, bold=True, font=font)
    add_rect(slide, x, 1.8, 2.8, 0.5, C["primary"] if name == "premium" else
             (C["accent"] if name == "modern" else C["panel"]))
    add_rect(slide, x, 2.4, 2.8, 0.5, C["accent"] if name == "premium" else
             (C["primary"] if name == "modern" else C["primary"]))
    add_rect(slide, x, 3.0, 2.8, 0.5, C["bg"])
    tb = add_tb(slide, x, 3.7, 2.8, 0.3)
    set_p(tb.text_frame.paragraphs[0], f"text: {pal['text']}", C["muted"], 9, font=font)

# --- Slide 5: Closing ---
slide = new_slide(prs, C)
add_rect(slide, 4.0, 2.5, 5.333, 2.5, C["primary"])
tb = add_tb(slide, 4.0, 2.8, 5.333, 0.8)
set_p(tb.text_frame.paragraphs[0], "Thank You", C["white"], 36, bold=True, font=font)
tb2 = add_tb(slide, 4.0, 3.8, 5.333, 0.5)
set_p(tb2.text_frame.paragraphs[0], "github.com/spencernero2021-hash/pptx-agent-toolkit", C["white"], 12, font=font)

prs.save(str(OUTPUT))
print(f"\n[OK] Demo saved: {OUTPUT.resolve()} ({OUTPUT.stat().st_size} bytes)")
