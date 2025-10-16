#!/usr/bin/env python3
"""
PDF Image Extractor GUI (Tkinter + PyMuPDF) — Option B (separate Input & Output folders for batch)

Author: Abdullah A. Alnajim <alnajim@protonmail.com>

Features:
- Single-file mode OR Batch folder mode (separate input/output selection)
- Remembers last-used paths & DPI (JSON in ~/.config or %APPDATA%)
- Drag-and-drop PDF (single-file mode)
- JSON report with per-file and global summary
- Error summary dialog on completion
- Threaded extraction (no UI freeze), live log, progress bar
"""

__author__ = "Abdullah A. Alnajim"
__email__ = "alnajim@protonmail.com"


import os
import re
import json
import threading
import queue
import platform
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional drag-and-drop
try:
    # tkinterdnd2 exposes a wrapper module; keep a reference to the module for robust construction
    TKDND_MODULE = __import__("tkinterdnd2")
    DND_FILES = getattr(TKDND_MODULE, "DND_FILES", None)
    TkinterDnD = getattr(TKDND_MODULE, "TkinterDnD", None) or getattr(TKDND_MODULE, "TkinterDnD2", None)
    TKDND_AVAILABLE = True
except Exception:
    TKDND_MODULE = None
    TkinterDnD = None
    DND_FILES = None
    TKDND_AVAILABLE = False

import fitz  # PyMuPDF

# ---------------- CONFIG HANDLING ----------------

def _config_path() -> Path:
    """Return platform-appropriate config file path."""
    if platform.system() == "Windows":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        return base / "PDFImageExtractor" / "config.json"
    else:
        return Path.home() / ".config" / "pdf_image_extractor" / "config.json"

def load_config() -> dict:
    try:
        p = _config_path()
        if p.exists():
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_config(cfg: dict):
    try:
        p = _config_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

# ---------------- CORE EXTRACTION ----------------

SAFE_CHARS = re.compile(r'[^A-Za-z0-9._@-]')

def safe_name(s: str) -> str:
    return SAFE_CHARS.sub("-", s)

def parse_page_range(pages: str | None, max_page: int) -> list[int]:
    """Parse '1,3-5,10' style ranges (1-based) -> zero-based sorted unique list."""
    if not pages:
        return list(range(max_page))
    wanted = set()
    for part in pages.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start = max(1, int(a))
            end = min(max_page, int(b))
            if start <= end:
                wanted.update(range(start, end + 1))
        else:
            p = int(part)
            if 1 <= p <= max_page:
                wanted.add(p)
    return sorted([p - 1 for p in wanted])

class ExtractionReport:
    """Accumulates per-run metadata and writes JSON report at the end."""
    def __init__(self, batch_mode: bool):
        self.batch_mode = batch_mode
        self.started_at = datetime.utcnow().isoformat() + "Z"
        self.files: list[dict] = []
        self.global_errors: list[str] = []

    def add_file_result(self, *, input_pdf: str, total_pages: int, pages_processed: list[int],
                        images_extracted: int, page_pngs: int, errors: list[dict]):
        self.files.append({
            "input": input_pdf,
            "total_pages": total_pages,
            "pages_processed": pages_processed,
            "images_extracted": images_extracted,
            "page_pngs": page_pngs,
            "errors": errors,
        })

    def add_global_error(self, msg: str):
        self.global_errors.append(msg)

    def to_dict(self):
        return {
            "batch_mode": self.batch_mode,
            "started_at": self.started_at,
            "finished_at": datetime.utcnow().isoformat() + "Z",
            "file_count": len(self.files),
            "files": self.files,
            "global_errors": self.global_errors,
            "summary": {
                "total_images": sum(f["images_extracted"] for f in self.files),
                "total_pages_exported": sum(f["page_pngs"] for f in self.files),
                "files_with_errors": sum(1 for f in self.files if f["errors"]),
            }
        }

    def write_json(self, out_folder: Path, filename: str = "extraction_report.json") -> Path:
        try:
            out_folder.mkdir(parents=True, exist_ok=True)
            p = out_folder / filename
            with open(p, "w", encoding="utf-8") as f:
                json.dump(self.to_dict(), f, indent=2)
            return p
        except Exception:
            return out_folder / filename

def extract_single_pdf(
    *,
    pdf_path: Path,
    output_folder: Path,
    export_pages: bool,
    dpi: int,
    pages: str | None,
    overwrite: bool,
    log_cb,
    progress_cb_pages,   # (cur_page_index, total_pages) for this PDF
    stop_event: threading.Event | None = None,
    file_report: ExtractionReport | None = None,
):
    """Extract one PDF, updating per-file progress and file report."""
    output_folder.mkdir(parents=True, exist_ok=True)
    errors: list[dict] = []
    images_total = 0
    png_total = 0
    pages_done: list[int] = []

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        msg = f"❌ Failed to open PDF: {pdf_path} — {e}"
        log_cb(msg)
        if file_report:
            file_report.add_global_error(msg)
        return

    page_indices = parse_page_range(pages, len(doc))
    total_pages = len(page_indices)
    progress_cb_pages(0, total_pages)
    log_cb(f"Opened: {pdf_path}  pages={len(doc)}  processing={total_pages} page(s)")

    seen_xrefs: set[int] = set()

    for idx_i, idx in enumerate(page_indices, start=1):
        if stop_event is not None and stop_event.is_set():
            log_cb(f"⏹️ Extraction cancelled: {pdf_path}")
            break
        try:
            page = doc[idx]
        except Exception as e:
            em = f"Failed to read page {idx+1}: {e}"
            log_cb(f"⚠️  {em}")
            errors.append({"page": idx+1, "error": em})
            progress_cb_pages(idx_i, total_pages)
            continue

        page_prefix = f"page-{idx+1:03d}"
        count_here = 0
        # Extract embedded images (original bytes)
        try:
            images = page.get_images(full=True)
        except Exception as e:
            images = []
            em = f"{page_prefix}: failed to list images: {e}"
            log_cb(f"⚠️  {em}")
            errors.append({"page": idx+1, "error": em})

        for img_pos, img in enumerate(images, start=1):
            try:
                xref = img[0]
                if xref in seen_xrefs:
                    continue
                seen_xrefs.add(xref)
                base = doc.extract_image(xref)
                data = base["image"]
                ext = base.get("ext", "bin")
                filename = safe_name(f"{page_prefix}-img-{img_pos}.{ext}")
                out_path = output_folder / filename
                if not overwrite and out_path.exists():
                    stem = out_path.stem
                    suf = 2
                    while (output_folder / f"{stem}-{suf}.{ext}").exists():
                        suf += 1
                    out_path = output_folder / f"{stem}-{suf}.{ext}"
                with open(out_path, "wb") as f:
                    f.write(data)
                count_here += 1
                images_total += 1
            except Exception as e:
                em = f"{page_prefix}: failed to save image {img_pos}: {e}"
                log_cb(f"⚠️  {em}")
                errors.append({"page": idx+1, "error": em})

        if count_here:
            log_cb(f"✅ {page_prefix}: extracted {count_here} image(s)")
        else:
            log_cb(f"ℹ️  {page_prefix}: no embedded images")

        # Optional: render page PNG
        if export_pages:
            try:
                pix = page.get_pixmap(dpi=dpi)
                out_png = output_folder / f"{page_prefix}.png"
                if not overwrite and out_png.exists():
                    n = 1
                    while (output_folder / f"{page_prefix}-{n}.png").exists():
                        n += 1
                    out_png = output_folder / f"{page_prefix}-{n}.png"
                pix.save(out_png)
                png_total += 1
                log_cb(f"🖼️  {page_prefix}: exported PNG → {out_png.name}")
            except Exception as e:
                em = f"{page_prefix}: failed to render PNG: {e}"
                log_cb(f"⚠️  {em}")
                errors.append({"page": idx+1, "error": em})
        # end page loop
        pages_done.append(idx+1)
        progress_cb_pages(idx_i, total_pages)

    try:
        doc.close()
    except Exception:
        pass

    # Add to report
    if isinstance(file_report, ExtractionReport):
        file_report.add_file_result(
            input_pdf=str(pdf_path),
            total_pages=len(doc),
            pages_processed=pages_done,
            images_extracted=images_total,
            page_pngs=png_total,
            errors=errors
        )

def extract_batch_folder(
    *,
    input_folder: Path,
    output_folder: Path,
    export_pages: bool,
    dpi: int,
    pages: str | None,
    overwrite: bool,
    log_cb,
    progress_cb_files,  # (cur_file_idx, total_files)
    progress_cb_pages,  # (cur_page, total_pages) for each file
    report: ExtractionReport,
    stop_event: threading.Event | None = None,
):
    pdfs = sorted([p for p in input_folder.glob("*.pdf") if p.is_file()])
    if not pdfs:
        log_cb("ℹ️  No PDF files found in the input folder.")
        return

    total = len(pdfs)
    progress_cb_files(0, total)
    log_cb(f"Batch mode: found {total} PDF(s) in {input_folder}")

    for i, pdf in enumerate(pdfs, start=1):
        if stop_event is not None and stop_event.is_set():
            log_cb("⏹️ Batch extraction cancelled")
            break
        # Create per-file subfolder under output for neatness:
        sub_out = output_folder / pdf.stem
        log_cb(f"—— Processing [{i}/{total}]: {pdf.name} ——")
        extract_single_pdf(
            pdf_path=pdf,
            output_folder=sub_out,
            export_pages=export_pages,
            dpi=dpi,
            pages=pages,
            overwrite=overwrite,
            log_cb=log_cb,
            progress_cb_pages=progress_cb_pages,
            stop_event=stop_event,
            file_report=report
        )
        progress_cb_files(i, total)

# ---------------- GUI ----------------

class PDFExtractorGUI(tk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master)
        self.master = master

        # Configure root window
        self.master.title("PDF Image Extractor")
        try:
            self.master.geometry("840x640")
            self.master.minsize(780, 560)
        except Exception:
            pass

        self._apply_style()
        self._load_last_config()
        self._build_widgets()
        # Bind drag-and-drop after widgets are created so binding can attach to the root
        self._bind_drag_drop()

        # Worker communications
        self.log_queue: queue.Queue = queue.Queue()
        self.master.after(50, self._drain_log_queue)

        self.running = False

    def _system_prefers_dark(self) -> bool:
        # Heuristic retained for possible future use, but dark-mode UI removed
        return platform.system() == "Darwin"

    # ---------- Config ----------
    def _load_last_config(self):
        cfg = load_config()
        self.pdf_path = tk.StringVar(value=cfg.get("pdf_path", ""))
        self.input_folder = tk.StringVar(value=cfg.get("input_folder", ""))
        self.out_path = tk.StringVar(value=cfg.get("output_folder", ""))
        self.dpi = tk.IntVar(value=cfg.get("dpi", 200))
        self.export_pages = tk.BooleanVar(value=False)
        self.overwrite = tk.BooleanVar(value=False)
        self.pages = tk.StringVar()
        self.batch_mode = tk.BooleanVar(value=False)

    def _save_last_config(self):
        save_config({
            "pdf_path": self.pdf_path.get(),
            "input_folder": self.input_folder.get(),
            "output_folder": self.out_path.get(),
            "dpi": self.dpi.get(),
        })

    # ---------- GUI Setup ----------
    def _apply_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=4)
        style.configure("Header.TLabel", font=("Helvetica", 16, "bold"))
        style.configure("Info.TLabel", foreground="#555")

    def _bind_drag_drop(self):
        # Only bind if the underlying tk widget supports the methods (checked at runtime)
        if TKDND_AVAILABLE:
            root = getattr(self, "master", None)
            if root is None:
                return
            drop_reg = getattr(root, "drop_target_register", None)
            dnd_b = getattr(root, "dnd_bind", None)
            if callable(drop_reg) and callable(dnd_b):
                try:
                    drop_reg(DND_FILES)
                    dnd_b("<<Drop>>", self._on_drop_file)
                    print("[INFO] Drag-and-drop bound to root")
                    log_fn = getattr(self, "_log", None)
                    if callable(log_fn):
                        try:
                            log_fn("🧷 Drag-and-drop enabled")
                        except Exception:
                            pass
                except Exception:
                    # If anything goes wrong, skip DnD binding
                    print("[WARN] Drag-and-drop binding failed")
                    log_fn = getattr(self, "_log", None)
                    if callable(log_fn):
                        try:
                            log_fn("⚠️ Drag-and-drop unavailable")
                        except Exception:
                            pass
                    pass

    def _on_drop_file(self, event):
        # Only applies to single-file mode
        if self.batch_mode.get():
            self._log("ℹ️  Drag-and-drop is for single-file mode. Uncheck Batch Mode.")
            return
        file = event.data.strip().strip("{}")
        if file.lower().endswith(".pdf"):
            self.pdf_path.set(file)
            self._log(f"📂 Dropped PDF: {file}")
        else:
            self._log("⚠️  Only PDF files are supported for drag-and-drop.")

    # ---------- Widgets ----------
    def _build_widgets(self):
        padx, pady = 12, 8

        # Header label
        self.panel_bg = "#e6e3e0"  # neutral light grey for panels
        self._panel_frames = []
        header = ttk.Label(self, text="PDF Image Extractor", style="Header.TLabel")
        header.grid(row=0, column=0, columnspan=4, sticky="w", padx=padx, pady=pady)

        # Batch toggle
        self.batch_chk = ttk.Checkbutton(self, text="Batch Folder Mode", variable=self.batch_mode, command=self._on_toggle_batch)
        self.batch_chk.grid(row=1, column=0, sticky="w", padx=padx, pady=pady)

        # Single-file row (hidden in batch mode)
        self.label_pdf = ttk.Label(self, text="PDF File:")
        self.label_pdf.grid(row=2, column=0, sticky="e", padx=padx, pady=pady)
        self.pdf_entry = ttk.Entry(self, textvariable=self.pdf_path)
        self.pdf_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=padx, pady=pady)
        self.choose_pdf_btn = ttk.Button(self, text="Select PDF…", command=self._select_pdf)
        self.choose_pdf_btn.grid(row=2, column=3, sticky="w", padx=padx, pady=pady)

        # Batch input folder row (hidden in single mode)
        self.label_input = ttk.Label(self, text="Input Folder (batch):")
        self.input_entry = ttk.Entry(self, textvariable=self.input_folder)
        self.choose_input_btn = ttk.Button(self, text="Select Input Folder…", command=self._select_input_folder)

        # Output folder row (always shown)
        ttk.Label(self, text="Output Folder:").grid(row=4, column=0, sticky="e", padx=padx, pady=pady)
        self.out_entry = ttk.Entry(self, textvariable=self.out_path)
        self.out_entry.grid(row=4, column=1, columnspan=2, sticky="ew", padx=padx, pady=pady)
        ttk.Button(self, text="Select Folder…", command=self._select_output).grid(row=4, column=3, sticky="w", padx=padx, pady=pady)

        # Options
        opts = tk.Frame(self, bg=self.panel_bg)
        opts.grid(row=5, column=0, columnspan=4, sticky="ew", padx=12)
        opts.grid_columnconfigure(5, weight=1)

        ttk.Checkbutton(opts, text="Export each page as PNG", variable=self.export_pages).grid(row=0, column=0, sticky="w", padx=(0, 18))
        ttk.Label(opts, text="DPI:").grid(row=0, column=1, sticky="e")
        ttk.Entry(opts, textvariable=self.dpi, width=7).grid(row=0, column=2, sticky="w", padx=(6, 18))
        ttk.Checkbutton(opts, text="Overwrite existing files", variable=self.overwrite).grid(row=0, column=3, sticky="w", padx=(0, 18))
        ttk.Label(opts, text="Pages (e.g. 1,3-5):").grid(row=1, column=0, sticky="w", pady=(4, 6))
        ttk.Entry(opts, textvariable=self.pages, width=28).grid(row=1, column=1, columnspan=3, sticky="w", padx=(6, 0))

        # Action + progress
        action = tk.Frame(self, bg=self.panel_bg)
        action.grid(row=6, column=0, columnspan=4, sticky="ew", padx=12, pady=4)
        action.grid_columnconfigure(1, weight=1)

        self.start_btn = ttk.Button(action, text="Start Extraction", command=self._start)
        self.start_btn.grid(row=0, column=0, sticky="w")

        self.cancel_btn = ttk.Button(action, text="Cancel", command=self._cancel)
        self.cancel_btn.grid(row=0, column=3, sticky="e")

        self.progress_files = ttk.Progressbar(action, orient="horizontal", mode="determinate", length=240, maximum=100)
        self.progress_files.grid(row=0, column=1, sticky="ew", padx=(12, 0))
        self.progress_files_label = ttk.Label(action, text="Files: 0/0", style="Info.TLabel")
        self.progress_files_label.grid(row=0, column=2, sticky="e", padx=(12, 0))

        self.progress_pages = ttk.Progressbar(action, orient="horizontal", mode="determinate", length=240, maximum=100)
        self.progress_pages.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=(6, 0))
        self.progress_pages_label = ttk.Label(action, text="Pages: 0/0", style="Info.TLabel")
        self.progress_pages_label.grid(row=1, column=2, sticky="e", padx=(12, 0), pady=(6, 0))

        # Log
        ttk.Label(self, text="Log Output:").grid(row=7, column=0, columnspan=4, sticky="w", padx=12)
        log_frame = ttk.Frame(self)
        log_frame.grid(row=8, column=0, columnspan=4, sticky="nsew", padx=12, pady=(0, 12))
        self.grid_rowconfigure(8, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.log_text = tk.Text(log_frame, wrap="word", height=16, state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        yscroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=yscroll.set)

        # Footer
        footer = tk.Frame(self, bg=self.panel_bg)
        footer.grid(row=9, column=0, columnspan=4, sticky="ew", padx=12, pady=(0, 12))
        ttk.Button(footer, text="About", command=self._show_about).pack(side="left")
        ttk.Button(footer, text="Reveal Output Folder", command=self._open_output_folder).pack(side="left", padx=(8, 0))
        ttk.Button(footer, text="Exit", command=self._exit).pack(side="right")

        # track panel frames so we can reapply bg after theme changes
        try:
            self._panel_frames.extend([opts, action, footer])
        except Exception:
            pass

        self._on_toggle_batch()  # initialize visibility

    # ----- Visibility toggle for batch mode -----
    def _on_toggle_batch(self):
        padx, pady = 12, 8
        if self.batch_mode.get():
            # Hide single PDF widgets
            self.label_pdf.grid_remove()
            self.pdf_entry.grid_remove()
            self.choose_pdf_btn.grid_remove()

            # Show input folder widgets
            self.label_input.grid(row=3, column=0, sticky="e", padx=padx, pady=pady)
            self.input_entry.grid(row=3, column=1, columnspan=2, sticky="ew", padx=padx, pady=pady)
            self.choose_input_btn.grid(row=3, column=3, sticky="w", padx=padx, pady=pady)
        else:
            # Show single PDF widgets
            self.label_pdf.grid(row=2, column=0, sticky="e", padx=padx, pady=pady)
            self.pdf_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=padx, pady=pady)
            self.choose_pdf_btn.grid(row=2, column=3, sticky="w", padx=padx, pady=pady)

            # Hide input folder widgets
            self.label_input.grid_remove()
            self.input_entry.grid_remove()
            self.choose_input_btn.grid_remove()

    # ---------- Handlers ----------
    def _select_pdf(self):
        path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.pdf_path.set(path)

    def _select_input_folder(self):
        path = filedialog.askdirectory(title="Select Input Folder (batch)")
        if path:
            self.input_folder.set(path)

    def _select_output(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.out_path.set(path)

    def _open_output_folder(self):
        folder = self.out_path.get()
        if not folder or not os.path.isdir(folder):
            return
        if platform.system() == "Darwin":
            subprocess.run(["open", folder])
        elif platform.system() == "Windows":
            subprocess.run(["explorer", folder])
        else:
            subprocess.run(["xdg-open", folder])

    def _show_about(self):
        """Show an About dialog with author information."""
        try:
            info = f"PDF Image Extractor\n\nAuthor: {__author__}\nEmail: {__email__}"
            messagebox.showinfo("About", info)
        except Exception:
            # Fallback if messagebox fails
            print(f"Author: {__author__} <{__email__}>")

    def _exit(self):
        # Signal worker to stop, wait briefly, then destroy root
        try:
            if hasattr(self, "_stop_event") and isinstance(self._stop_event, threading.Event):
                self._log("⏹️ Stopping background worker…")
                self._stop_event.set()
            # If there's a worker thread, wait up to 3 seconds for it to finish
            wt = getattr(self, "_worker_thread", None)
            if isinstance(wt, threading.Thread) and wt.is_alive():
                wt.join(timeout=3.0)
        except Exception:
            pass
        # Finally destroy the root
        try:
            root = getattr(self, "master", None)
            if root:
                root.destroy()
            else:
                self.destroy()
        except Exception:
            pass

    def _start(self):
        if self.running:
            return
        # Validate
        outdir = Path(self.out_path.get().strip())
        if not outdir:
            messagebox.showerror("Invalid Output Folder", "Please select an output folder.")
            return

        # Common options
        try:
            dpi_val = int(self.dpi.get())
            if not (50 <= dpi_val <= 1200):
                raise ValueError
        except Exception:
            messagebox.showerror("Invalid DPI", "Please enter a DPI between 50 and 1200.")
            return

        pages = self.pages.get().strip()
        export_pages = self.export_pages.get()
        overwrite = self.overwrite.get()

        # Mode-specific validation
        if self.batch_mode.get():
            in_folder = Path(self.input_folder.get().strip())
            if not in_folder.exists() or not in_folder.is_dir():
                messagebox.showerror("Invalid Input Folder", "Please select a valid input folder.")
                return
            worker_args = ("batch", in_folder, outdir, export_pages, dpi_val, pages, overwrite)
        else:
            pdf = Path(self.pdf_path.get().strip())
            if not pdf.exists() or pdf.suffix.lower() != ".pdf":
                messagebox.showerror("Invalid PDF", "Please select a valid PDF file.")
                return
            worker_args = ("single", pdf, outdir, export_pages, dpi_val, pages, overwrite)

        self._save_last_config()
        self._set_running_state(True)
        self._log("▶️ Starting extraction…")

        # Prepare stop event and worker thread
        self._stop_event = threading.Event()
        # Start non-daemon thread so we can join during exit for graceful shutdown
        th = threading.Thread(target=self._worker, args=(*worker_args, self._stop_event), daemon=False)
        self._worker_thread = th
        th.start()

    def _worker(self, mode, *args):
        # Expect optional stop_event as final arg
        stop_event = None
        if args and isinstance(args[-1], threading.Event):
            *args, stop_event = args
        error_count = 0

        def log_cb(msg: str):
            nonlocal error_count
            if msg.startswith("⚠️") or msg.startswith("❌"):
                error_count += 1
            self.log_queue.put(msg)

        # pages and files progress callbacks
        def progress_pages(cur: int, tot: int):
            self.log_queue.put(("__PAGES__", cur, tot))

        def progress_files(cur: int, tot: int):
            self.log_queue.put(("__FILES__", cur, tot))

        report = ExtractionReport(batch_mode=(mode == "batch"))

        try:
            if mode == "batch":
                in_folder, outdir, export_pages, dpi_val, pages, overwrite = args
                extract_batch_folder(
                    input_folder=in_folder,
                    output_folder=outdir,
                    export_pages=export_pages,
                    dpi=dpi_val,
                    pages=pages if pages else None,
                    overwrite=overwrite,
                    log_cb=log_cb,
                    progress_cb_files=progress_files,
                    progress_cb_pages=progress_pages,
                    report=report,
                    stop_event=stop_event,
                )
            else:
                pdf, outdir, export_pages, dpi_val, pages, overwrite = args
                extract_single_pdf(
                    pdf_path=pdf,
                    output_folder=outdir,
                    export_pages=export_pages,
                    dpi=dpi_val,
                    pages=pages if pages else None,
                    overwrite=overwrite,
                    log_cb=log_cb,
                    progress_cb_pages=progress_pages,
                    stop_event=stop_event,
                    file_report=report,
                )
        finally:
            # Write JSON report
            report_path = report.write_json(Path(self.out_path.get().strip()))
            self.log_queue.put(f"🧾 Wrote report → {report_path}")
            # Signal done + error count
            self.log_queue.put(("__DONE__", error_count, report))

    def _drain_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item:
                    tag = item[0]
                    if tag == "__PAGES__":
                        _, cur, tot = item
                        self._update_progress_pages(cur, tot)
                    elif tag == "__FILES__":
                        _, cur, tot = item
                        self._update_progress_files(cur, tot)
                    elif tag == "__DONE__":
                        _, err_count, report = item
                        self._set_running_state(False)
                        self._finish_dialog(err_count, report)
                    else:
                        # Unknown tuple
                        pass
                else:
                    self._log(str(item))
        except queue.Empty:
            pass
        finally:
            self.after(50, self._drain_log_queue)

    # ---------- UI Utilities ----------

    def _update_progress_files(self, cur: int, tot: int):
        pct = int(cur / tot * 100) if tot else 0
        self.progress_files.configure(value=pct, maximum=100)
        self.progress_files_label.configure(text=f"Files: {cur}/{tot}" if tot else "Files: 0/0")

    def _update_progress_pages(self, cur: int, tot: int):
        pct = int(cur / tot * 100) if tot else 0
        self.progress_pages.configure(value=pct, maximum=100)
        self.progress_pages_label.configure(text=f"Pages: {cur}/{tot}" if tot else "Pages: 0/0")

    def _log(self, text: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_running_state(self, running: bool):
        self.running = running
        state = "disabled" if running else "normal"
        for w in (
            self.pdf_entry, self.out_entry, self.input_entry if self.batch_mode.get() else self.pdf_entry,
            self.start_btn, self.choose_pdf_btn, self.choose_input_btn if self.batch_mode.get() else self.choose_pdf_btn,
            self.batch_chk
        ):
            try:
                w.configure(state=state)
            except Exception:
                pass
        if running:
            self.progress_files_label.configure(text="Working…")
            self.progress_pages_label.configure(text="Working…")

    def _cancel(self):
        # Request worker stop without exiting the app
        try:
            if hasattr(self, "_stop_event") and isinstance(self._stop_event, threading.Event):
                self._log("⏹️ Cancellation requested…")
                self._stop_event.set()
            # update UI state to reflect cancellation pending
            self._set_running_state(False)
        except Exception:
            pass

    def _finish_dialog(self, err_count: int, report: ExtractionReport):
        if err_count > 0 or report.to_dict()["summary"]["files_with_errors"] > 0:
            messagebox.showwarning(
                "Completed with Warnings",
                f"Extraction finished with {err_count} warning(s).\n"
                f"Check the log and extraction_report.json for details."
            )
        else:
            messagebox.showinfo("Completed", "Extraction completed successfully!")
        # Reset progress labels
        self.progress_files.configure(value=0)
        self.progress_pages.configure(value=0)
        self.progress_files_label.configure(text="Files: 0/0")
        self.progress_pages_label.configure(text="Pages: 0/0")

if __name__ == "__main__":
    # Create the appropriate root window: prefer TkinterDnD root when available
    if TKDND_AVAILABLE:
        # Try the module-level constructor first (tkinterdnd2.Tk()), then class-based
        try:
            if TKDND_MODULE and hasattr(TKDND_MODULE, "Tk"):
                root = TKDND_MODULE.Tk()
            elif TkinterDnD and hasattr(TkinterDnD, "Tk"):
                root = TkinterDnD.Tk()
            else:
                root = tk.Tk()
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()

    app = PDFExtractorGUI(root)
    app.pack(fill="both", expand=True)

    # good resizing behavior
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(8, weight=1)
    root.mainloop()
