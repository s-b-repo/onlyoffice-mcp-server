"""Document preview — render pages as PNG images for AI visual inspection.

Documents are first converted to PDF via LibreOffice headless, then individual
pages are rendered as PNG at the requested DPI. Rendering uses PyMuPDF (fitz)
when it is importable, otherwise it falls back to poppler's ``pdftoppm`` /
``pdfinfo`` command-line tools — so preview works whether or not the optional
``fitz`` package is installed.

Temp images are written to ~/.onlyoffice-mcp/preview/ and auto-cleaned
after 30 minutes.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from .validation import validate_path
from .storage import home as get_home

log = logging.getLogger(__name__)

PREVIEW_DIR_NAME = "preview"
PREVIEW_TTL_SECONDS = 1800  # 30 min

_SUPPORTED_EXTENSIONS = {"docx", "xlsx", "pptx", "pdf", "odt", "ods", "odp"}


def _preview_dir() -> Path:
    d = get_home() / PREVIEW_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _cleanup_stale(directory: Path) -> None:
    """Remove preview images older than PREVIEW_TTL_SECONDS."""
    now = time.time()
    for f in directory.iterdir():
        if f.is_file() and (now - f.stat().st_mtime) > PREVIEW_TTL_SECONDS:
            f.unlink(missing_ok=True)


def _convert_to_pdf(src: Path) -> Path:
    """Convert a document to PDF via LibreOffice headless. Returns PDF path."""
    soffice = shutil.which("soffice")
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. Install it for document preview:\n"
            "  sudo apt install libreoffice-common"
        )

    with tempfile.TemporaryDirectory(prefix="oomcp-preview-") as tmpdir:
        cmd = [
            soffice, "--headless", "--convert-to", "pdf",
            "--outdir", tmpdir, str(src),
        ]
        log.info("Converting to PDF: %s", " ".join(cmd))
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed (exit {result.returncode}):\n"
                f"{result.stderr[:500]}"
            )

        pdf_files = list(Path(tmpdir).glob("*.pdf"))
        if not pdf_files:
            raise RuntimeError(
                "LibreOffice produced no PDF output.\n"
                f"stdout: {result.stdout[:300]}\nstderr: {result.stderr[:300]}"
            )

        dest = _preview_dir() / f"{src.stem}_{int(time.time())}.pdf"
        shutil.move(str(pdf_files[0]), str(dest))
        return dest


def _parse_page_range(pages: str | None, total: int) -> list[int]:
    """Parse a page range string like '1-3,5,8-10' into zero-based indices."""
    if pages is None:
        return list(range(total))

    result: list[int] = []
    for part in pages.split(","):
        part = part.strip()
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = max(1, int(start_s))
            end = min(total, int(end_s))
            result.extend(range(start - 1, end))
        else:
            idx = int(part) - 1
            if 0 <= idx < total:
                result.append(idx)
    return sorted(set(result))


def _try_fitz():
    """Return the imported ``fitz`` module, or ``None`` if unavailable."""
    try:
        import fitz  # PyMuPDF
        return fitz
    except Exception:  # ImportError or a broken build
        return None


def _pdf_page_count(pdf_path: Path, fitz_mod) -> int:
    """Total page count via fitz if available, else poppler ``pdfinfo``."""
    if fitz_mod is not None:
        doc = fitz_mod.open(str(pdf_path))
        try:
            return doc.page_count
        finally:
            doc.close()
    pdfinfo = shutil.which("pdfinfo")
    if pdfinfo:
        result = subprocess.run(
            [pdfinfo, str(pdf_path)], capture_output=True, text=True, timeout=60,
        )
        for line in result.stdout.splitlines():
            if line.lower().startswith("pages:"):
                return int(line.split(":", 1)[1].strip())
    raise RuntimeError(
        "Cannot render preview: neither PyMuPDF ('fitz') nor poppler ('pdfinfo') "
        "is available. Install one:\n"
        "  pip install pymupdf      # or\n"
        "  sudo apt install poppler-utils"
    )


def _render_page_pdftoppm(pdf_path: Path, page_no: int, dpi: int, out_path: Path) -> tuple[int, int]:
    """Render a single 1-indexed page to ``out_path`` (PNG) via poppler.

    ``pdftoppm -singlefile`` writes ``<prefix>.png``, so the prefix is the
    output path with its ``.png`` suffix removed. Returns (width, height) px.
    """
    tool = shutil.which("pdftoppm") or shutil.which("pdftocairo")
    if not tool:
        raise RuntimeError(
            "poppler 'pdftoppm' not found (needed when PyMuPDF is absent).\n"
            "Install it: sudo apt install poppler-utils"
        )
    prefix = str(out_path)
    if prefix.lower().endswith(".png"):
        prefix = prefix[:-4]
    cmd = [
        tool, "-png", "-r", str(dpi),
        "-f", str(page_no), "-l", str(page_no),
        "-singlefile", str(pdf_path), prefix,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0 or not out_path.exists():
        raise RuntimeError(
            f"pdftoppm failed for page {page_no} (exit {result.returncode}): "
            f"{result.stderr[:300]}"
        )
    from PIL import Image
    with Image.open(out_path) as im:
        return im.width, im.height


def doc_preview(
    path: str,
    *,
    pages: str | None = None,
    dpi: int = 150,
    max_pages: int = 10,
) -> dict:
    """Render document pages as PNG images for AI visual inspection.

    Returns a dict with page image paths, total page count, and rendering
    metadata. The AI can view these images using its file-read capability.
    Uses PyMuPDF (fitz) when available, otherwise poppler's pdftoppm.
    """
    p = validate_path(path, must_exist=True, operation="doc_preview")
    ext = p.suffix.lstrip(".").lower()
    if ext not in _SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported format '.{ext}' for preview.\n"
            f"Supported: {', '.join(sorted(_SUPPORTED_EXTENSIONS))}"
        )

    if dpi < 36 or dpi > 600:
        raise ValueError(
            f"dpi={dpi} is out of range [36, 600].\n"
            f"Recommended: 72 (fast/small), 150 (balanced), 300 (high quality)."
        )

    fitz_mod = _try_fitz()
    engine = "fitz" if fitz_mod is not None else "pdftoppm"

    preview_dir = _preview_dir()
    _cleanup_stale(preview_dir)

    pdf_path: Path | None = None
    temp_pdf = False
    fitz_doc = None
    try:
        if ext == "pdf":
            pdf_path = p
        else:
            pdf_path = _convert_to_pdf(p)
            temp_pdf = True

        total_pages = _pdf_page_count(pdf_path, fitz_mod)

        page_indices = _parse_page_range(pages, total_pages)
        if len(page_indices) > max_pages:
            page_indices = page_indices[:max_pages]
            truncated = True
        else:
            truncated = False

        if fitz_mod is not None:
            fitz_doc = fitz_mod.open(str(pdf_path))
            zoom = dpi / 72.0
            mat = fitz_mod.Matrix(zoom, zoom)

        result_pages: list[dict] = []
        stem = p.stem
        ts = int(time.time())

        for idx in page_indices:
            out_path = preview_dir / f"{stem}_{ts}_p{idx + 1}.png"
            if fitz_doc is not None:
                pix = fitz_doc[idx].get_pixmap(matrix=mat)
                pix.save(str(out_path))
                w_px, h_px = pix.width, pix.height
            else:
                w_px, h_px = _render_page_pdftoppm(pdf_path, idx + 1, dpi, out_path)

            result_pages.append({
                "page": idx + 1,
                "path": str(out_path),
                "width_px": w_px,
                "height_px": h_px,
            })
            log.info("Rendered page %d → %s (%dx%d, %s)", idx + 1, out_path, w_px, h_px, engine)

        rendered_range = (
            f"{page_indices[0] + 1}-{page_indices[-1] + 1}"
            if page_indices else "none"
        )

        return {
            "page_images": result_pages,
            "total_pages": total_pages,
            "rendered": f"{len(result_pages)} of {total_pages} pages ({rendered_range})",
            "truncated": truncated,
            "dpi": dpi,
            "engine": engine,
            "source": str(p),
            "hint": (
                "Use your file-read tool to view each page image. "
                f"{'More pages available — call again with pages= parameter.' if truncated else ''}"
            ),
        }
    finally:
        if fitz_doc is not None:
            fitz_doc.close()
        if temp_pdf and pdf_path and pdf_path.exists():
            pdf_path.unlink(missing_ok=True)
