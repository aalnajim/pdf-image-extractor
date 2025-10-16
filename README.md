# 🧩 PDF Image Extractor (Python)

Extract **embedded images** from PDF files in their **original format** (JPEG, JPX, PNG, CCITT/TIFF, etc.), and optionally export each **page as a PNG** at a custom DPI.  
Works on **macOS, Windows, and Linux**.

> Built with **PyMuPDF** (`fitz`) for PDF parsing and **Pillow** for image handling.



## ✨ Features

- Extracts **original embedded images** — no recompression.
- **De-duplicates** reused image objects (xref-aware).
- Optional **page rendering** to PNG (`--export-pages` with custom `--dpi`).
- Flexible **page range selection** (`--pages 1,3-5,10`).
- Safe filenames and `--overwrite` option.
- Lightweight, fast, and cross-platform.


## 📦 Requirements

- **Python 3.9+** (tested on 3.10–3.12)
- Install dependencies with:
  ```bash
  pip install -r requirements.txt
  ```


## 🚀 Usage

  ```bash
  python pdf-image-extractor.py <PDF_PATH> <OUTPUT_DIR> [--export-pages] [--dpi 200] [--pages "1,3-5"] [--overwrite]
```


## Examples
```bash
# Extract all embedded images only
python pdf-image-extractor.py input.pdf out

# Extract images and export pages 2–5 and 10 as PNG at 300 DPI
python pdf-image-extractor.py input.pdf out --pages 2-5,10 --export-pages --dpi 300

# Overwrite existing files in output directory
python pdf-image-extractor.py input.pdf out --overwrite
```


## 🪟 Windows (PowerShell)
```bash
py .\pdf-image-extractor.py .\input.pdf .\out --export-pages --dpi 300
```
---

## 🧰 CLI Options
### Positional arguments

| Argument     | Type      | Required | Description                          | Example      |
|--------------|-----------|----------|--------------------------------------|--------------|
| `PDF_PATH`   | file path | yes      | Path to the input PDF file           | `input.pdf`  |
| `OUTPUT_DIR` | dir path  | yes      | Directory to write extracted outputs | `out`        |

### Options

| Flag               | Type     | Default | Description                                                                 | Example                  |
|--------------------|----------|---------|-----------------------------------------------------------------------------|--------------------------|
| `--export-pages`   | boolean  | false   | Also render each selected page to a PNG file                                | `--export-pages`         |
| `--dpi <value>`    | integer  | 200     | Render DPI for page PNGs (used with `--export-pages`). Common: 150–300      | `--dpi 300`              |
| `--pages <spec>`   | ranges   | all     | 1-based page selection; comma-separated pages/ranges. Inclusive (e.g., 2-5) | `--pages "1,3-5,10"`     |
| `--overwrite`      | boolean  | false   | Overwrite existing files in `OUTPUT_DIR` (otherwise existing files are skipped) | `--overwrite`          |

Notes
- Page spec examples: `5`, `2-4`, `1,3-5,10` (1-based, inclusive)
- Without `--export-pages`, only embedded images are extracted (original formats preserved)


## 🗂️ Output Layout
````markdown
output/
  page-001-img-1.jpg
  page-001-img-2.jp2
  page-001.png          # if --export-pages is used
  page-002-img-1.tiff
  page-002.png
  ...
````


## 🔍 Tips & Troubleshooting
- **No images found?**

    Some PDFs use vector graphics, not embedded images.
Use --export-pages to render pages visually to PNGs.
- **Color / transparency issues?**

    Embedded images are raw; page PNGs are fully rendered.
- **Large PNG sizes?**

    Reduce DPI (e.g., --dpi 150 or --dpi 96).
- **Performance:**

    Rendering cost grows with pixel area — keep DPI reasonable for batch jobs.


## ✅ Quick Test
Run on a sample PDF with known photos (like a scanned brochure) and verify output files:
````markdown
page-001-img-1.jpg
page-001.png
````
Then compare the number of images with your PDF viewer’s “Content” inspector.



## 🗺️ Roadmap
- **Tkinter GUI** (Select PDF / Output Folder / Progress Log)
- Drag-and-drop support
- Batch mode for multiple PDFs
- JSON export with per-page metadata
- Docker image for headless servers


## 🏷️ License
MIT — see [LICENSE](./LICENSE)


## 🙌 Acknowledgments
- [PyMuPDF](https://pymupdf.readthedocs.io/) — high-performance PDF parsing and rendering
- [Pillow](https://pillow.readthedocs.io/) — Python Imaging Library fork for image handling
