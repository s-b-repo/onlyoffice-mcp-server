"""LibreOffice (``soffice --headless``) integration.

Used as a conversion fallback in ``converter.py`` when ONLYOFFICE Document
Builder isn't installed, and as the engine for ``doc_page_count(precise=True)``
(via PDF + pypdf).

The ``-env:UserInstallation`` flag is mandatory: it makes soffice use a
fresh per-call profile so concurrent calls don't collide on the user's
default ``~/.config/libreoffice`` lock.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from .errors import EngineMissing


_CANDIDATES = ("soffice", "libreoffice")

# soffice uses an export filter token for some formats; common cases.
_SOFFICE_FILTER = {
    "txt": "txt:Text",
    "csv": "csv",
    "pdf": "pdf",
    "html": "html",
    # docx / xlsx / pptx / odt / ods / odp / rtf / doc / xls / ppt: token == ext
}


def find_soffice() -> str | None:
    for candidate in _CANDIDATES:
        path = shutil.which(candidate)
        if path:
            return path
    return None


def status() -> dict:
    binary = find_soffice()
    if not binary:
        return {
            "installed": False,
            "path": None,
            "version": None,
            "install_instructions": (
                "Install LibreOffice: 'sudo apt install libreoffice' "
                "or download from https://www.libreoffice.org/."
            ),
        }
    try:
        r = subprocess.run(
            [binary, "--version"], capture_output=True, text=True, timeout=10
        )
        version = (r.stdout + r.stderr).strip() or "(unknown)"
    except Exception as e:
        version = f"(error: {e})"
    return {"installed": True, "path": binary, "version": version}


def convert(input_path: str | Path, output_path: str | Path, *, timeout: int = 180) -> str:
    """Convert via ``soffice --headless --convert-to <fmt>``.

    Returns the resolved output path on success. Raises
    :class:`EngineMissing` if soffice isn't on PATH, or :class:`RuntimeError`
    if conversion fails.
    """
    binary = find_soffice()
    if not binary:
        raise EngineMissing("LibreOffice (soffice) not installed on PATH")

    in_ = Path(input_path).expanduser().resolve()
    out = Path(output_path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    out.parent.mkdir(parents=True, exist_ok=True)

    out_ext = out.suffix.lstrip(".").lower()
    filter_token = _SOFFICE_FILTER.get(out_ext, out_ext)

    with tempfile.TemporaryDirectory(prefix="oo_mcp_lo_") as tmp:
        profile = Path(tmp) / "profile"
        cmd = [
            binary,
            "--headless",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to",
            filter_token,
            "--outdir",
            str(out.parent),
            str(in_),
        ]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        if r.returncode != 0:
            raise RuntimeError(
                f"soffice exited {r.returncode}\n"
                f"--- stdout ---\n{r.stdout}\n"
                f"--- stderr ---\n{r.stderr}"
            )

    # soffice always writes <stem>.<out_ext> in outdir.
    produced = out.parent / f"{in_.stem}.{out_ext}"
    if produced.exists() and produced != out:
        produced.rename(out)
    if not out.exists():
        raise RuntimeError(
            f"soffice did not produce {out}. Available in outdir: "
            f"{list(out.parent.iterdir())[:10]}"
        )
    return str(out)


def page_count_via_pdf(input_path: str | Path, *, timeout: int = 60) -> int:
    """Convert the document to PDF in a tempdir and count pages with pypdf."""
    try:
        from pypdf import PdfReader
    except ImportError as e:
        raise EngineMissing(
            "pypdf is not installed. Run: pip install 'onlyoffice-mcp[pdf]'"
        ) from e
    in_ = Path(input_path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    with tempfile.TemporaryDirectory(prefix="oo_mcp_pdf_") as tmp:
        pdf_out = Path(tmp) / (in_.stem + ".pdf")
        convert(in_, pdf_out, timeout=timeout)
        return len(PdfReader(str(pdf_out)).pages)
