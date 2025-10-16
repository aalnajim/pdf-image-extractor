#!/usr/bin/env python3
# pip install PyMuPDF Pillow

# Author: Abdullah A. Alnajim
# Email: alnajim@protonmail.com
# GitHub: https://github.com/aalnajim/pdf-image-extractor 
# License: MIT
# Extract embedded images from a PDF (optionally export each page as PNG).
# Usage: python main.py input.pdf output_folder --export-pages --pages 1,3-5 --dpi 300 --overwrite
# Notes:
# - Embedded images are saved in their original format (JPEG, PNG, etc).
# - Page PNGs are rendered at specified DPI (default 200).
# - Safe filenames and --overwrite option.
# - Lightweight, fast, and cross-platform.

import argparse
import os
import re
from pathlib import Path
import fitz  # PyMuPDF

SAFE_CHARS = re.compile(r'[^A-Za-z0-9._@-]')

def safe_name(s: str) -> str:
    return SAFE_CHARS.sub("-", s)

def parse_page_range(pages: str, max_page: int) -> list[int]:
    """
    Parse '1,3-5,10' style ranges (1-based) -> zero-based list.
    """
    if not pages:
        return list(range(max_page))
    wanted = set()
    for part in pages.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            start = max(1, int(a))
            end = min(max_page, int(b))
            wanted.update(range(start, end + 1))
        else:
            p = int(part)
            if 1 <= p <= max_page:
                wanted.add(p)
    # to zero-based
    return sorted([p - 1 for p in wanted])

def extract_images_from_pdf(
    pdf_path: Path,
    output_folder: Path,
    export_pages: bool = False,
    dpi: int = 200,
    pages: str | None = None,
    overwrite: bool = False,
) -> None:
    output_folder.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)

    page_indices = parse_page_range(pages, len(doc))
    print(f"Opened: {pdf_path}  pages={len(doc)}  processing={len(page_indices)} page(s)")

    total_images = 0
    seen_xrefs: set[int] = set()  # avoid saving same image multiple times

    for idx in page_indices:
        page = doc[idx]
        page_prefix = f"page-{idx+1:03d}"

        # --- Embedded images (original bytes) ---
        images = page.get_images(full=True)
        if images:
            count_here = 0
            for img_pos, img in enumerate(images, start=1):
                xref = img[0]
                if xref in seen_xrefs:
                    # image object reused on another page; skip duplicate
                    continue
                seen_xrefs.add(xref)

                base = doc.extract_image(xref)
                data = base["image"]
                ext = base.get("ext", "bin")
                filename = safe_name(f"{page_prefix}-img-{img_pos}.{ext}")
                out_path = output_folder / filename

                if not overwrite and out_path.exists():
                    # if name collides (rare), add a suffix
                    stem = out_path.stem
                    suf = 2
                    while (output_folder / f"{stem}-{suf}.{ext}").exists():
                        suf += 1
                    out_path = output_folder / f"{stem}-{suf}.{ext}"

                with open(out_path, "wb") as f:
                    f.write(data)
                count_here += 1
                total_images += 1

            print(f"✅ {page_prefix}: extracted {count_here} image(s)")
        else:
            print(f"⚠️ {page_prefix}: no embedded images")

        # --- Optional: full page PNG ---
        if export_pages:
            pix = page.get_pixmap(dpi=dpi)  # renders vector content too
            out_png = output_folder / f"{page_prefix}.png"
            if not overwrite and out_png.exists():
                out_png = output_folder / f"{page_prefix}-1.png"
                n = 2
                while out_png.exists():
                    out_png = output_folder / f"{page_prefix}-{n}.png"
                    n += 1
            pix.save(out_png)
            print(f"🖼️ {page_prefix}: exported PNG → {out_png.name}")

    doc.close()
    print(f"\n🎉 Done. Extracted {total_images} unique image object(s).")
    print(f"Output → {output_folder.resolve()}")

def main():
    ap = argparse.ArgumentParser(description="Extract embedded images from a PDF (optionally export each page as PNG).")
    ap.add_argument("pdf", type=Path, help="Path to PDF")
    ap.add_argument("out", type=Path, help="Output folder")
    ap.add_argument("--export-pages", action="store_true", help="Also export each page as PNG")
    ap.add_argument("--dpi", type=int, default=200, help="DPI for page PNG render (default: 200)")
    ap.add_argument("--pages", type=str, default=None, help="Pages to process, e.g. '1,3-5,10' (1-based)")
    ap.add_argument("--overwrite", action="store_true", help="Overwrite existing files")
    args = ap.parse_args()

    extract_images_from_pdf(
        pdf_path=args.pdf,
        output_folder=args.out,
        export_pages=args.export_pages,
        dpi=args.dpi,
        pages=args.pages,
        overwrite=args.overwrite,
    )

if __name__ == "__main__":
    main()