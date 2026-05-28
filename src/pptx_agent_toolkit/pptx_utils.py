"""
Shared PPTX XML helpers for color extraction, shape inspection, and z-order.
"""

import io
from zipfile import ZipFile
from pathlib import Path
from PIL import Image

from pptx import Presentation
from pptx.oxml.ns import qn

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ---- Text color ------------------------------------------------------------

def get_text_color(run) -> str | None:
    """Extract effective text color from a run (handles defRPr inheritance).

    Checks: run.rPr > paragraph defRPr > returns None if inherited further.
    """
    # Check run-level rPr
    rPr = run._r.find(f"{{{NS_A}}}rPr")
    if rPr is not None:
        srgb = rPr.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
        if srgb is not None:
            return srgb.get("val")

    # Check paragraph-level defRPr
    p = run._r.getparent()
    if p is None:
        return None
    pPr = p.find(f"{{{NS_A}}}pPr")
    if pPr is not None:
        defRPr = pPr.find(f"{{{NS_A}}}defRPr")
        if defRPr is not None:
            srgb = defRPr.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
            if srgb is not None:
                return srgb.get("val")

    return None


# ---- Shape fill ------------------------------------------------------------

def get_shape_fill_color(shape) -> str | None:
    """Get solid fill color of a shape. Returns hex string or None."""
    try:
        spPr = shape._element.find(qn("p:spPr"))
    except AttributeError:
        return None
    if spPr is None:
        return None

    # No fill
    nofill = spPr.find(f"{{{NS_A}}}noFill")
    if nofill is not None:
        return None

    # Solid fill
    srgb = spPr.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
    if srgb is not None:
        return srgb.get("val")

    # Gradient - return first stop color
    grad = spPr.find(f"{{{NS_A}}}gradFill")
    if grad is not None:
        first_stop = grad.find(f"{{{NS_A}}}gsLst/{{{NS_A}}}gs")
        if first_stop is not None:
            srgb = first_stop.find(f"{{{NS_A}}}srgbClr")
            if srgb is not None:
                return srgb.get("val")

    return None


def get_shape_line_color(shape) -> str | None:
    """Get outline color of a shape."""
    try:
        spPr = shape._element.find(qn("p:spPr"))
    except AttributeError:
        return None
    if spPr is None:
        return None
    srgb = spPr.find(f"{{{NS_A}}}ln/{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
    if srgb is not None:
        return srgb.get("val")
    return None


def is_picture_shape(shape) -> bool:
    """Check if a shape is a picture (PICTURE type)."""
    try:
        return shape._element.tag == qn("p:pic")
    except AttributeError:
        return False


# ---- Slide background ------------------------------------------------------

def get_slide_background(slide) -> str | None:
    """Get slide background color. Returns hex or None."""
    bg = slide._element.find(f'{qn("p:cSld")}/{qn("p:bg")}')
    if bg is not None:
        srgb = bg.find(f".//{{{NS_A}}}srgbClr")
        if srgb is not None:
            return srgb.get("val")
    return None


def get_default_bg(prs) -> str:
    """Try to resolve default slide background from theme/layout."""
    # Check first slide layout's background
    if prs.slide_layouts:
        layout = prs.slide_layouts[0]
        try:
            bg = layout._element.find(f'{qn("p:cSld")}/{qn("p:bg")}')
            if bg is not None:
                srgb = bg.find(f".//{{{NS_A}}}srgbClr")
                if srgb is not None:
                    return srgb.get("val")
        except Exception:
            pass
    return "FFFFFF"


# ---- Shape bounds ----------------------------------------------------------

def get_shape_bounds_emu(shape) -> tuple:
    """Return (left, top, right, bottom) in EMU."""
    l = shape.left or 0
    t = shape.top or 0
    w = shape.width or 0
    h = shape.height or 0
    return (l, t, l + w, t + h)


def rects_overlap(a: tuple, b: tuple) -> bool:
    """Check if two EMU rectangles overlap."""
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


# ---- Background detection for text -----------------------------------------

def find_bg_for_text(text_shape, slide) -> tuple[str | None, str]:
    """Find the effective background behind a text shape.

    Returns (color_hex, source) where source is one of:
      'shape_fill', 'picture', 'slide_bg', 'none'
    """
    text_bounds = get_shape_bounds_emu(text_shape)

    # 1) Try parent shape's own fill
    own_fill = get_shape_fill_color(text_shape)
    if own_fill is not None:
        return own_fill, "shape_fill"

    # 2) Walk z-order (reverse) looking for overlapping filled shape or picture
    sp_tree = slide.shapes._spTree
    children = list(sp_tree)[2:]  # skip nvGrpSpPr, grpSpPr
    text_el = text_shape._element

    # Find text shape's position in children
    try:
        text_idx = children.index(text_el)
    except ValueError:
        text_idx = len(children)

    # Check shapes behind text (lower indices in z-order → behind)
    for i in range(text_idx - 1, -1, -1):
        child = children[i]
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag not in ("sp", "pic"):
            continue

        # Get bounds
        xfrm = child.find(f"{{{NS_A}}}spPr/{{{NS_A}}}xfrm") or \
               child.find(qn("p:spPr")) if False else None
        if xfrm is None:
            xfrm = child.find(f".//{{{NS_A}}}xfrm")
        if xfrm is None:
            off = child.find(f".//{{{NS_A}}}off")
            ext = child.find(f".//{{{NS_A}}}ext")
            if off is not None and ext is not None:
                lb = int(off.get("x", 0))
                tb = int(off.get("y", 0))
                rb = lb + int(ext.get("cx", 0))
                bb = tb + int(ext.get("cy", 0))
                child_bounds = (lb, tb, rb, bb)
            else:
                continue
        else:
            off = xfrm.find(f"{{{NS_A}}}off")
            ext = xfrm.find(f"{{{NS_A}}}ext")
            if off is None or ext is None:
                continue
            lb = int(off.get("x", 0))
            tb = int(off.get("y", 0))
            rb = lb + int(ext.get("cx", 0))
            bb = tb + int(ext.get("cy", 0))
            child_bounds = (lb, tb, rb, bb)

        if not rects_overlap(text_bounds, child_bounds):
            continue

        if tag == "pic":
            return None, "picture"

        # Check fill
        spPr = child.find(f"{{{NS_A}}}spPr") or child.find(qn("p:spPr"))
        if spPr is None:
            continue
        if spPr.find(f"{{{NS_A}}}noFill") is not None:
            continue
        srgb = spPr.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
        if srgb is not None:
            return srgb.get("val"), "shape_fill"

    # 3) Fallback to slide background
    slide_bg = get_slide_background(slide)
    if slide_bg:
        return slide_bg, "slide_bg"

    return None, "none"


# ---- Picture extraction ----------------------------------------------------

def sample_image_behind_text(slide, slide_idx: int, pptx_path: str,
                              text_shape) -> str | None:
    """Sample the dominant color from a background image behind text.

    Opens the PPTX as a zip, finds the picture behind the text shape,
    extracts image pixels in the text's bounding box, and returns the
    median color as a hex string.
    """
    text_bounds = get_shape_bounds_emu(text_shape)
    sp_tree = slide.shapes._spTree
    children = list(sp_tree)[2:]
    text_el = text_shape._element
    try:
        text_idx = children.index(text_el)
    except ValueError:
        return None

    # Find the first picture behind this text
    pic_el = None
    pic_bounds = None
    for i in range(text_idx - 1, -1, -1):
        child = children[i]
        if child.tag == qn("p:pic"):
            spPr = child.find(qn("p:spPr"))
            if spPr is not None:
                xfrm = spPr.find(f"{{{NS_A}}}xfrm")
                if xfrm is not None:
                    off = xfrm.find(qn("a:off"))
                    ext = xfrm.find(qn("a:ext"))
                    if off is not None and ext is not None:
                        lb = int(off.get("x", 0))
                        tb = int(off.get("y", 0))
                        rb = lb + int(ext.get("cx", 0))
                        bb = tb + int(ext.get("cy", 0))
                        if rects_overlap(text_bounds, (lb, tb, rb, bb)):
                            pic_el = child
                            pic_bounds = (lb, tb, rb, bb)
                            break

    if pic_el is None:
        return None

    # Get blip embed rId
    blip = pic_el.find(f".//{{{NS_A}}}blip")
    if blip is None:
        return None
    r_embed = blip.get(f"{{{NS_A.replace('drawingml', 'officeDocument')}}}embed") or \
              blip.get(qn("r:embed"))
    if r_embed is None:
        return None

    # Open pptx as zip, find the image
    pptx_path = Path(pptx_path)
    with ZipFile(pptx_path, "r") as zf:
        # Read slide rels
        rels_path = f"ppt/slides/_rels/slide{slide_idx + 1}.xml.rels"
        try:
            with zf.open(rels_path) as f:
                rels_xml = f.read()
        except KeyError:
            return None

        # Find relationship target
        from xml.etree.ElementTree import fromstring
        rels_root = fromstring(rels_xml)
        target = None
        for rel in rels_root:
            if rel.get("Id") == r_embed:
                target = rel.get("Target")
                break
        if target is None:
            return None

        # Resolve image path
        img_path = f"ppt/slides/{target}" if not target.startswith("../") else \
                   f"ppt/{target.replace('../', '')}"
        try:
            with zf.open(img_path) as f:
                img_data = f.read()
        except KeyError:
            # Try media folder
            media_name = Path(target).name
            try:
                with zf.open(f"ppt/media/{media_name}") as f:
                    img_data = f.read()
            except KeyError:
                return None

    # Load image and sample text region
    img = Image.open(io.BytesIO(img_data)).convert("RGB")
    iw, ih = img.size
    pw = pic_bounds[2] - pic_bounds[0]
    ph = pic_bounds[3] - pic_bounds[1]
    if pw == 0 or ph == 0:
        return None

    # Map EMU to pixel coordinates
    tl = int((text_bounds[0] - pic_bounds[0]) / pw * iw)
    tt = int((text_bounds[1] - pic_bounds[1]) / ph * ih)
    tr = int((text_bounds[2] - pic_bounds[0]) / pw * iw)
    tb = int((text_bounds[3] - pic_bounds[1]) / ph * ih)
    tl = max(0, min(iw - 1, tl))
    tt = max(0, min(ih - 1, tt))
    tr = max(tl + 1, min(iw, tr))
    tb = max(tt + 1, min(ih, tb))

    region = img.crop((tl, tt, tr, tb))
    pixels = list(region.getdata())
    if not pixels:
        return None

    # Use median of each channel
    from statistics import median
    r = int(median(p[0] for p in pixels))
    g = int(median(p[1] for p in pixels))
    b = int(median(p[2] for p in pixels))
    return f"{r:02X}{g:02X}{b:02X}"


# ---- Font size -------------------------------------------------------------

def get_font_size(run) -> float | None:
    """Get font size in points from a run. Returns None if inherited."""
    rPr = run._r.find(f"{{{NS_A}}}rPr")
    if rPr is not None:
        sz = rPr.get("sz")
        if sz:
            return int(sz) / 100.0
    p = run._r.getparent()
    if p is None:
        return None
    pPr = p.find(f"{{{NS_A}}}pPr")
    if pPr is not None:
        defRPr = pPr.find(f"{{{NS_A}}}defRPr")
        if defRPr is not None:
            sz = defRPr.get("sz")
            if sz:
                return int(sz) / 100.0
    return None


def is_large_text(font_size_pt: float | None) -> bool:
    """WCAG definition: >= 18pt or >= 14pt bold."""
    if font_size_pt is None:
        return False
    return font_size_pt >= 18.0
