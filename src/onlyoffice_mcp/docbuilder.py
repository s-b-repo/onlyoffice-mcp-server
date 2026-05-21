"""ONLYOFFICE Document Builder CLI integration.

Document Builder is a standalone tool from ONLYOFFICE that executes
JavaScript-like scripts to generate and convert documents. It supports the
widest set of formats (docx/xlsx/pptx/pdf/odt/ods/odp/rtf/txt/csv/html/epub).

Install on Debian/Ubuntu/Kali:
    wget https://download.onlyoffice.com/install/documentbuilder/linux/onlyoffice-documentbuilder_amd64.deb
    sudo dpkg -i onlyoffice-documentbuilder_amd64.deb
    sudo apt -f install  # if any deps are missing

Reference: https://api.onlyoffice.com/docbuilder/basic
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

# Common install paths for the documentbuilder binary on Linux.
_CANDIDATE_BINARIES = [
    "documentbuilder",
    "docbuilder",
    "/usr/bin/documentbuilder",
    "/opt/onlyoffice/documentbuilder/documentbuilder",
]

# Extensions that Document Builder accepts in builder.SaveFile(format, path).
_SUPPORTED_FORMATS = {
    "docx", "doc", "odt", "rtf", "txt", "html", "epub",
    "xlsx", "xls", "ods", "csv",
    "pptx", "ppt", "odp",
    "pdf", "pdfa",
}


def find_binary() -> str | None:
    """Return the path to the documentbuilder binary, or None if not installed."""
    for candidate in _CANDIDATE_BINARIES:
        path = shutil.which(candidate)
        if path:
            return path
        if Path(candidate).is_file():
            return candidate
    return None


def status() -> dict:
    """Return a dict describing whether documentbuilder is installed."""
    binary = find_binary()
    if not binary:
        return {
            "installed": False,
            "path": None,
            "version": None,
            "supported_formats": sorted(_SUPPORTED_FORMATS),
            "install_instructions": (
                "Install ONLYOFFICE Document Builder for full conversion support:\n"
                "  wget https://download.onlyoffice.com/install/documentbuilder/linux/"
                "onlyoffice-documentbuilder_amd64.deb\n"
                "  sudo dpkg -i onlyoffice-documentbuilder_amd64.deb\n"
                "  sudo apt -f install"
            ),
        }
    try:
        result = subprocess.run(
            [binary, "--version"], capture_output=True, text=True, timeout=10
        )
        version = (result.stdout + result.stderr).strip() or "(unknown)"
    except Exception as e:  # noqa: BLE001
        version = f"(error: {e})"
    return {
        "installed": True,
        "path": binary,
        "version": version,
        "supported_formats": sorted(_SUPPORTED_FORMATS),
    }


def run(script: str, output_path: str | None = None, *, timeout: int = 120) -> str:
    """Execute a Document Builder script.

    If `output_path` is provided and the script does not call
    `builder.SaveFile(...)` itself, a save instruction is appended automatically
    based on the output_path extension.

    Returns the resolved output path on success, raises RuntimeError otherwise.
    """
    binary = find_binary()
    if not binary:
        raise RuntimeError(
            "ONLYOFFICE Document Builder is not installed. "
            "Call `docbuilder_status` for install instructions."
        )

    out_path = None
    if output_path:
        out_path = Path(output_path).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)

    # Stage the script. We prepend an OUTPUT_PATH global for convenience and
    # auto-append a save call if the user's script didn't include one.
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".docbuilder",
        prefix="oo_mcp_",
        delete=False,
        encoding="utf-8",
    ) as f:
        if out_path is not None:
            f.write(f'var OUTPUT_PATH = "{out_path.as_posix()}";\n')
        f.write(script)
        if out_path is not None and "SaveFile" not in script:
            ext = out_path.suffix.lstrip(".").lower()
            fmt = ext if ext in _SUPPORTED_FORMATS else "docx"
            f.write(f'\nbuilder.SaveFile("{fmt}", OUTPUT_PATH);\nbuilder.CloseFile();\n')
        script_path = f.name

    try:
        result = subprocess.run(
            [binary, script_path],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"documentbuilder exited {result.returncode}\n"
                f"--- stdout ---\n{result.stdout}\n"
                f"--- stderr ---\n{result.stderr}"
            )
    finally:
        try:
            Path(script_path).unlink()
        except OSError:
            pass

    if out_path is not None and not out_path.exists():
        raise RuntimeError(
            f"documentbuilder finished but output file was not written: {out_path}\n"
            "Make sure the script calls builder.SaveFile(format, path) or pass output_path."
        )
    return str(out_path) if out_path else ""


def build_conversion_script(input_path: str, output_path: str) -> str:
    """Return a Document Builder script that opens `input_path` and saves it
    as the format implied by `output_path`'s extension."""
    in_ = Path(input_path).expanduser().resolve()
    out = Path(output_path).expanduser().resolve()
    out_fmt = out.suffix.lstrip(".").lower()
    if out_fmt not in _SUPPORTED_FORMATS:
        raise ValueError(
            f"Output format '{out_fmt}' not supported. "
            f"Supported: {sorted(_SUPPORTED_FORMATS)}"
        )
    return (
        f'builder.OpenFile("{in_.as_posix()}");\n'
        f'builder.SaveFile("{out_fmt}", "{out.as_posix()}");\n'
        f'builder.CloseFile();\n'
    )
