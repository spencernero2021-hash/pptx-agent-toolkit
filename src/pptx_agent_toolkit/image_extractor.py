"""
Extract images from PPTX, DOCX, and PDF files.

Usage:
  python pptx_image_extractor.py <file_or_folder> -o <output_dir>
  python pptx_image_extractor.py course/ -o course_images/
  python pptx_image_extractor.py slides.pptx -o images/
"""

import argparse
import json
import sys
from pathlib import Path
from zipfile import ZipFile
from io import BytesIO

from PIL import Image


def extract_from_pptx(pptx_path: str, output_dir: str) -> list[dict]:
    """Extract all images from a PPTX file. Returns metadata list."""
    results = []
    base = Path(pptx_path).stem
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    with ZipFile(pptx_path, "r") as zf:
        # Find slide-to-image mapping via relationships
        slide_images = {}  # rId -> image_path
        for name in zf.namelist():
            if name.startswith("ppt/slides/_rels/") and name.endswith(".xml.rels"):
                rels_xml = zf.read(name).decode("utf-8")
                import re
                for rid, target in re.findall(
                    r'Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"', rels_xml
                ):
                    if "image" in target.lower() or target.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff", ".emf", ".wmf")):
                        slide_num = re.search(r'slide(\d+)', name)
                        sn = int(slide_num.group(1)) if slide_num else 0
                        # Resolve path
                        if target.startswith("../"):
                            resolved = f"ppt/{target.replace('../', '')}"
                        else:
                            resolved = f"ppt/slides/{target}"
                        slide_images.setdefault(sn, []).append((rid, resolved, target))

        # Extract media files
        for name in zf.namelist():
            if not name.startswith("ppt/media/"):
                continue
            fname = Path(name).name
            stem = Path(fname).stem
            ext = Path(fname).suffix.lower()
            if ext not in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff", ".emf", ".wmf"):
                continue

            data = zf.read(name)
            try:
                img = Image.open(BytesIO(data))
                w, h = img.size
            except Exception:
                w, h = 0, 0

            out_path = out / f"{base}_{stem}{ext}"
            with open(out_path, "wb") as f:
                f.write(data)

            # Find which slide(s) this image belongs to
            slides = []
            for sn, refs in slide_images.items():
                for rid, resolved, target in refs:
                    if resolved == name:
                        slides.append(sn)

            results.append({
                "path": str(out_path),
                "source": pptx_path,
                "source_type": "pptx",
                "slides": slides,
                "width": w,
                "height": h,
                "format": ext.lstrip("."),
                "size_bytes": len(data),
            })

    return results


def extract_from_docx(docx_path: str, output_dir: str) -> list[dict]:
    """Extract all images from a DOCX file."""
    results = []
    base = Path(docx_path).stem
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    with ZipFile(docx_path, "r") as zf:
        for name in zf.namelist():
            if not name.startswith("word/media/"):
                continue
            fname = Path(name).name
            ext = Path(fname).suffix.lower()

            data = zf.read(name)
            try:
                img = Image.open(BytesIO(data))
                w, h = img.size
            except Exception:
                w, h = 0, 0

            out_path = out / f"{base}_{fname}"
            # Avoid overwrite
            if out_path.exists():
                stem = Path(fname).stem
                out_path = out / f"{base}_{stem}_{len(results)}{ext}"
            with open(out_path, "wb") as f:
                f.write(data)

            results.append({
                "path": str(out_path),
                "source": docx_path,
                "source_type": "docx",
                "slides": [],
                "width": w,
                "height": h,
                "format": ext.lstrip("."),
                "size_bytes": len(data),
            })

    return results


def extract_from_pdf(pdf_path: str, output_dir: str) -> list[dict]:
    """Extract images and/or page renders from a PDF file."""
    import fitz
    results = []
    base = Path(pdf_path).stem
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Extract embedded images
        image_list = page.get_images(full=True)
        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]
            ext = base_image["ext"]

            out_path = out / f"{base}_p{page_num+1}_{img_idx}.{ext}"
            with open(out_path, "wb") as f:
                f.write(img_bytes)

            try:
                img = Image.open(BytesIO(img_bytes))
                w, h = img.size
            except Exception:
                w, h = 0, 0

            results.append({
                "path": str(out_path),
                "source": pdf_path,
                "source_type": "pdf",
                "page": page_num + 1,
                "slides": [],
                "width": w,
                "height": h,
                "format": ext,
                "size_bytes": len(img_bytes),
            })

        # If no embedded images, render page as image
        if not image_list and page_num < 5:  # limit to 5 pages
            pix = page.get_pixmap(dpi=150)
            out_path = out / f"{base}_p{page_num+1}_page.png"
            pix.save(str(out_path))
            results.append({
                "path": str(out_path),
                "source": pdf_path,
                "source_type": "pdf_page",
                "page": page_num + 1,
                "slides": [],
                "width": pix.width,
                "height": pix.height,
                "format": "png",
                "size_bytes": out_path.stat().st_size,
            })

    doc.close()
    return results


EXTRACTORS = {
    ".pptx": extract_from_pptx,
    ".ppt": extract_from_pptx,
    ".docx": extract_from_docx,
    ".doc": extract_from_docx,
    ".pdf": extract_from_pdf,
}


def extract_all(input_path: str, output_dir: str) -> list[dict]:
    """Extract images from a file or all supported files in a folder."""
    p = Path(input_path)
    results = []

    if p.is_file():
        ext = p.suffix.lower()
        if ext in EXTRACTORS:
            results = EXTRACTORS[ext](str(p), output_dir)
        else:
            print(f"Unsupported format: {ext}", file=sys.stderr)
    elif p.is_dir():
        for ext, extractor in EXTRACTORS.items():
            for f in p.glob(f"*{ext}"):
                try:
                    r = extractor(str(f), output_dir)
                    results.extend(r)
                except Exception as e:
                    print(f"Error extracting {f}: {e}", file=sys.stderr)
    else:
        print(f"Not found: {input_path}", file=sys.stderr)

    return results


def main():
    parser = argparse.ArgumentParser(description="Extract images from documents")
    parser.add_argument("input", help="File or folder to extract from")
    parser.add_argument("-o", "--output", default="extracted_images",
                        help="Output directory (default: extracted_images)")
    parser.add_argument("--json", action="store_true",
                        help="Output JSON manifest to stdout")
    args = parser.parse_args()

    results = extract_all(args.input, args.output)

    if args.json:
        print(json.dumps({"status": "ok", "count": len(results), "images": results},
                         ensure_ascii=False, indent=2))
    else:
        print(f"Extracted {len(results)} images to {args.output}/")
        for r in results:
            src = Path(r["source"]).name
            print(f"  {Path(r['path']).name}  ({r['width']}x{r['height']})  <- {src}")


if __name__ == "__main__":
    main()
