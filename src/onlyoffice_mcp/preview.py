"""Document preview — render pages as PNG images for AI visual inspection.

Uses PyMuPDF (fitz) to render PDF pages. Documents are first converted to
PDF via LibreOffice headless, then individual pages are rendered as PNG at
the requested DPI.

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
    """
    import fitz  # PyMuPDF

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

    preview_dir = _preview_dir()
    _cleanup_stale(preview_dir)

    pdf_path: Path | None = None
    temp_pdf = False
    try:
        if ext == "pdf":
            pdf_path = p
        else:
            pdf_path = _convert_to_pdf(p)
            temp_pdf = True

        doc = fitz.open(str(pdf_path))
        total_pages = len(doc)

        page_indices = _parse_page_range(pages, total_pages)
        if len(page_indices) > max_pages:
            page_indices = page_indices[:max_pages]
            truncated = True
        else:
            truncated = False

        result_pages: list[dict] = []
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)

        stem = p.stem
        ts = int(time.time())

        for idx in page_indices:
            page = doc[idx]
            pix = page.get_pixmap(matrix=mat)

            out_name = f"{stem}_{ts}_p{idx + 1}.png"
            out_path = preview_dir / out_name
            pix.save(str(out_path))

            result_pages.append({
                "page": idx + 1,
                "path": str(out_path),
                "width_px": pix.width,
                "height_px": pix.height,
            })
            log.info("Rendered page %d → %s (%dx%d)", idx + 1, out_path, pix.width, pix.height)

        doc.close()

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
            "source": str(p),
            "hint": (
                "Use your file-read tool to view each page image. "
                f"{'More pages available — call again with pages= parameter.' if truncated else ''}"
            ),
        }
    finally:
        if temp_pdf and pdf_path and pdf_path.exists():
            pdf_path.unlink(missing_ok=True)
