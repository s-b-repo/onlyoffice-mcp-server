"""Word document operations using python-docx.

OOXML output is natively compatible with ONLYOFFICE, Microsoft Word,
LibreOffice Writer, and Google Docs.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Map of friendly style names to python-docx style names. python-docx loads
# the default Office template which ships with these names.
STYLE_MAP: dict[str, str] = {
    "heading1": "Heading 1",
    "heading2": "Heading 2",
    "heading3": "Heading 3",
    "heading4": "Heading 4",
    "heading5": "Heading 5",
    "heading6": "Heading 6",
    "title": "Title",
    "subtitle": "Subtitle",
    "body": "Normal",
    "normal": "Normal",
    "quote": "Quote",
    "code": "Intense Quote",
    "caption": "Caption",
    "list": "List Bullet",
    "bullet": "List Bullet",
    "numbered": "List Number",
}

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _apply_style(paragraph, style_name: str, doc: Document) -> None:
    style_key = style_name.lower()
    docx_style = STYLE_MAP.get(style_key, style_name)
    try:
        paragraph.style = doc.styles[docx_style]
    except KeyError:
        # Style not in template — leave as Normal.
        pass


def _add_block(doc: Document, item: Any) -> None:
    if isinstance(item, str):
        doc.add_paragraph(item)
        return
    if not isinstance(item, dict):
        raise ValueError(f"Unsupported paragraph item type: {type(item).__name__}")

    item_type = item.get("type", "paragraph")

    if item_type in ("paragraph", "text") or ("text" in item and item_type == "paragraph"):
        text = item.get("text", "")
        style = item.get("style", "normal")
        para = doc.add_paragraph(text)
        _apply_style(para, style, doc)
        if item.get("align"):
            align = ALIGN_MAP.get(item["align"].lower())
            if align is not None:
                para.alignment = align
        if item.get("bold") or item.get("italic") or item.get("color") or item.get("size"):
            for run in para.runs:
                if item.get("bold"):
                    run.bold = True
                if item.get("italic"):
                    run.italic = True
                if item.get("size"):
                    run.font.size = Pt(item["size"])
                if item.get("color"):
                    color = item["color"].lstrip("#")
                    if len(color) == 6:
                        run.font.color.rgb = RGBColor(
                            int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                        )

    elif item_type == "heading":
        text = item.get("text", "")
        level = max(0, min(int(item.get("level", 1)), 9))
        doc.add_heading(text, level=level)

    elif item_type == "table":
        data = item.get("data", [])
        if not data:
            return
        rows = len(data)
        cols = max(len(row) for row in data)
        tbl = doc.add_table(rows=rows, cols=cols)
        tbl.style = item.get("style", "Light Grid Accent 1")
        for ri, row in enumerate(data):
            for ci, cell in enumerate(row):
                tbl.cell(ri, ci).text = "" if cell is None else str(cell)
        # Bold the header row if requested (default True).
        if item.get("header", True) and rows > 0:
            for cell in tbl.rows[0].cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

    elif item_type == "image":
        img_path = Path(item["path"]).expanduser().resolve()
        if not img_path.exists():
            raise ValueError(f"Image not found: {img_path}")
        width = item.get("width_inches")
        kwargs: dict[str, Any] = {}
        if width is not None:
            kwargs["width"] = Inches(width)
        doc.add_picture(str(img_path), **kwargs)

    elif item_type == "pagebreak":
        doc.add_page_break()

    elif item_type == "list":
        items = item.get("items", [])
        ordered = item.get("ordered", False)
        style = "List Number" if ordered else "List Bullet"
        for line in items:
            para = doc.add_paragraph(str(line))
            try:
                para.style = doc.styles[style]
            except KeyError:
                pass

    else:
        raise ValueError(f"Unknown block type: {item_type!r}")


def create(
    path: str,
    paragraphs: list[Any],
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
) -> str:
    """Create a .docx file at `path` with the given content blocks."""
    out = Path(path).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    cp = doc.core_properties
    if title:
        cp.title = title
    if author:
        cp.author = author
    if subject:
        cp.subject = subject
    for item in paragraphs:
        _add_block(doc, item)
    doc.save(str(out))
    return str(out)


def read(path: str, include_tables: bool = True) -> str:
    """Return the textual content of a .docx file."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    doc = Document(str(in_))
    lines: list[str] = [p.text for p in doc.paragraphs]
    if include_tables:
        for tbl in doc.tables:
            lines.append("")  # blank line before table
            for row in tbl.rows:
                cells = [c.text for c in row.cells]
                lines.append(" | ".join(cells))
    return "\n".join(lines)


def append(path: str, paragraphs: list[Any]) -> str:
    """Append content blocks to an existing .docx file."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    doc = Document(str(in_))
    for item in paragraphs:
        _add_block(doc, item)
    doc.save(str(in_))
    return str(in_)


def set_metadata(
    path: str,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
) -> str:
    """Set the core document properties on an existing .docx file."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    doc = Document(str(in_))
    cp = doc.core_properties
    if title is not None:
        cp.title = title
    if author is not None:
        cp.author = author
    if subject is not None:
        cp.subject = subject
    if keywords is not None:
        cp.keywords = keywords
    if comments is not None:
        cp.comments = comments
    doc.save(str(in_))
    return str(in_)
