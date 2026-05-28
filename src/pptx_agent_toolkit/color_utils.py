"""
Shared color utilities: WCAG 2.1 contrast, luminance, palette validation.
Zero external dependencies (math only).
"""

import math


def hex_to_rgb(hex_color: str) -> tuple:
    """Parse hex color string (with or without #) to (R, G, B) 0-255."""
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert (R, G, B) 0-255 to uppercase hex string without #."""
    return f"{max(0, min(255, r)):02X}{max(0, min(255, g)):02X}{max(0, min(255, b)):02X}"


def _linearize(c: float) -> float:
    """Linearize a single sRGB channel (0.0-1.0)."""
    if c <= 0.04045:
        return c / 12.92
    return ((c + 0.055) / 1.055) ** 2.4


def relative_luminance(hex_color: str) -> float:
    """WCAG 2.1 relative luminance from an sRGB hex color."""
    r, g, b = hex_to_rgb(hex_color)
    return 0.2126 * _linearize(r / 255.0) + \
           0.7152 * _linearize(g / 255.0) + \
           0.0722 * _linearize(b / 255.0)


def contrast_ratio(fg_hex: str, bg_hex: str) -> float:
    """WCAG 2.1 contrast ratio between two sRGB hex colors."""
    l1 = relative_luminance(fg_hex)
    l2 = relative_luminance(bg_hex)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def wcag_level(ratio: float, large_text: bool = False) -> str:
    """Return 'AAA', 'AA', 'AA_large', or 'FAIL'."""
    if ratio >= 7.0:
        return "AAA"
    if ratio >= 4.5:
        return "AAA" if large_text else "AA"
    if large_text and ratio >= 3.0:
        return "AA_large"
    return "FAIL"


def is_aa_pass(ratio: float, large_text: bool = False) -> bool:
    """Check if contrast ratio passes WCAG AA."""
    return ratio >= (3.0 if large_text else 4.5)


def is_aaa_pass(ratio: float, large_text: bool = False) -> bool:
    """Check if contrast ratio passes WCAG AAA."""
    return ratio >= (4.5 if large_text else 7.0)


def find_contrast_fix(fg_hex: str, bg_hex: str, target: float = 4.5) -> str:
    """Find the closest safe text color that meets target contrast against bg.

    Adjusts text luminance up or down while minimizing RGB distance from original.
    """
    fg_r, fg_g, fg_b = hex_to_rgb(fg_hex)
    bg_lum = relative_luminance(bg_hex)

    best_hex = fg_hex
    best_dist = float("inf")

    # Try darkening
    for factor in range(0, 101):
        t = factor / 100.0
        r = int(fg_r * (1 - t))
        g = int(fg_g * (1 - t))
        b = int(fg_b * (1 - t))
        trial = rgb_to_hex(r, g, b)
        if trial == fg_hex:
            continue
        lum = relative_luminance(trial)
        cr = (max(lum, bg_lum) + 0.05) / (min(lum, bg_lum) + 0.05)
        if cr >= target:
            dist = math.sqrt((fg_r - r) ** 2 + (fg_g - g) ** 2 + (fg_b - b) ** 2)
            if dist < best_dist:
                best_dist = dist
                best_hex = trial

    # Try lightening
    for factor in range(0, 101):
        t = factor / 100.0
        r = int(fg_r + (255 - fg_r) * t)
        g = int(fg_g + (255 - fg_g) * t)
        b = int(fg_b + (255 - fg_b) * t)
        trial = rgb_to_hex(r, g, b)
        if trial == fg_hex:
            continue
        lum = relative_luminance(trial)
        cr = (max(lum, bg_lum) + 0.05) / (min(lum, bg_lum) + 0.05)
        if cr >= target:
            dist = math.sqrt((fg_r - r) ** 2 + (fg_g - g) ** 2 + (fg_b - b) ** 2)
            if dist < best_dist:
                best_dist = dist
                best_hex = trial

    return best_hex


def palette_aa_check(palette: dict[str, str]) -> list[str]:
    """Validate that key color pairs in a palette meet WCAG AA.

    palette should have keys like 'text', 'bg', 'muted', 'accent', 'primary', 'panel'.
    Returns list of warning strings (empty = all pass).
    """
    warnings = []
    pairs = [
        ("text", "bg", False),
        ("muted", "bg", False),
        ("text_on_primary", "primary", False),
        ("text", "panel", False),
    ]
    # Resolve derived colors
    resolved = dict(palette)
    resolved.setdefault("text_on_primary", "#FFFFFF")

    for fg_key, bg_key, large in pairs:
        if fg_key not in resolved or bg_key not in resolved:
            continue
        fg = resolved[fg_key]
        bg = resolved[bg_key]
        cr = contrast_ratio(fg, bg)
        if not is_aa_pass(cr, large):
            warnings.append(
                f"Low contrast: {fg_key}({fg}) vs {bg_key}({bg}) = {cr:.2f}:1 "
                f"(needs {'≥3.0' if large else '≥4.5'})"
            )
    return warnings


def rgb_distance(c1: str, c2: str) -> float:
    """Euclidean distance in RGB space between two hex colors."""
    r1, g1, b1 = hex_to_rgb(c1)
    r2, g2, b2 = hex_to_rgb(c2)
    return math.sqrt((r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2)
