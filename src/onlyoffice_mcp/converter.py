"""Format conversion engine chain.

Order of preference:
    1. ONLYOFFICE Document Builder (widest format set, if installed)
    2. LibreOffice headless (broad set, if soffice is on PATH)
    3. Pure-Python pairs (docx<->txt, xlsx<->csv, pptx->txt)
"""

from __future__ import annotations

import csv
from pathlib import Path

from . import docbuilder, docx_ops, libreoffice, pptx_ops, xlsx_ops


# Conversion pairs we can do purely in Python without any external engine.
_PYTHON_FALLBACK_PAIRS = {
    ("docx", "txt"),
    ("xlsx", "csv"),
    ("csv", "xlsx"),
    ("pptx", "txt"),
}


def convert(input_path: str, output_path: str) -> str:
    """Convert a document from one format to another.

    Output format is inferred from ``output_path``'s extension. Engines are
    tried in order: Document Builder, LibreOffice, then Python fallbacks.
    """
    in_ = Path(input_path).expanduser().resolve()
    out = Path(output_path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    out.parent.mkdir(parents=True, exist_ok=True)

    in_ext = in_.suffix.lstrip(".").lower()
    out_ext = out.suffix.lstrip(".").lower()

    # No-op short circuit: copying same format is just a file copy.
    if in_ext == out_ext:
        out.write_bytes(in_.read_bytes())
        return str(out)

    # 1. ONLYOFFICE Document Builder — widest format set.
    if docbuilder.find_binary():
        script = docbuilder.build_conversion_script(str(in_), str(out))
        return docbuilder.run(script, str(out))

    # 2. LibreOffice headless — broad format set, available on this host.
    if libreoffice.find_soffice():
        return libreoffice.convert(in_, out)

    # Python fallbacks for common pairs.
    pair = (in_ext, out_ext)

    if pair == ("docx", "txt"):
        text = docx_ops.read(str(in_), include_tables=True)
        out.write_text(text, encoding="utf-8")
        return str(out)

    if pair == ("xlsx", "csv"):
        data = xlsx_ops.read(str(in_))
        first_sheet = next(iter(data))
        rows = data[first_sheet]
        with out.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for row in rows:
                writer.writerow(["" if v is None else v for v in row])
        return str(out)

    if pair == ("csv", "xlsx"):
        with in_.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.reader(f))
        xlsx_ops.create(str(out), {"Sheet1": rows}, header_bold=False)
        return str(out)

    if pair == ("pptx", "txt"):
        info = pptx_ops.read(str(in_))
        lines: list[str] = []
        for slide in info["slides"]:
            lines.append(f"# Slide {slide['index']}: {slide['title']}")
            for text in slide["text"]:
                lines.append(f"  - {text}")
            if slide["notes"]:
                lines.append(f"  (notes) {slide['notes']}")
            lines.append("")
        out.write_text("\n".join(lines), encoding="utf-8")
        return str(out)

    raise RuntimeError(
        f"Cannot convert {in_ext} -> {out_ext}.\n"
        f"Python fallbacks available: {sorted(_PYTHON_FALLBACK_PAIRS)}\n"
        f"For full format support, install one of:\n"
        f"  - ONLYOFFICE Document Builder: wget https://download.onlyoffice.com/"
        f"install/documentbuilder/linux/onlyoffice-documentbuilder_amd64.deb "
        f"&& sudo dpkg -i onlyoffice-documentbuilder_amd64.deb\n"
        f"  - LibreOffice: sudo apt install libreoffice"
    )
