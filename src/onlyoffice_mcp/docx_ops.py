"""Word document operations using python-docx.

OOXML output is natively compatible with ONLYOFFICE, Microsoft Word,
LibreOffice Writer, and Google Docs.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.shared import Inches, Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .validation import validate_path, validate_color, validate_align, sanitize_text

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


def _apply_style(paragraph, style_name: str, doc: Document) -> str | None:
    style_key = style_name.lower()
    docx_style = STYLE_MAP.get(style_key, style_name)
    try:
        paragraph.style = doc.styles[docx_style]
        return None
    except KeyError:
        available = [s.name for s in doc.styles if s.type == 1][:20]
        return (
            f"Style '{style_name}' not found in template (defaulting to Normal). "
            f"Available: {available}"
        )


def _add_block(doc: Document, item: Any) -> None:
    if isinstance(item, str):
        doc.add_paragraph(sanitize_text(item, "paragraph"))
        return
    if not isinstance(item, dict):
        raise ValueError(f"Unsupported paragraph item type: {type(item).__name__}")

    item_type = item.get("type", "paragraph")

    if item_type in ("paragraph", "text") or ("text" in item and item_type == "paragraph"):
        text = sanitize_text(item.get("text", ""), "paragraph text")
        style = item.get("style", "normal")
        para = doc.add_paragraph(text)
        _apply_style(para, style, doc)
        if item.get("align"):
            align = ALIGN_MAP.get(item["align"].lower())
            if align is not None:
                para.alignment = align
        has_fmt = any(item.get(k) for k in ("bold", "italic", "underline", "strikethrough", "color", "size", "font"))
        if has_fmt:
            for run in para.runs:
                if item.get("bold"):
                    run.bold = True
                if item.get("italic"):
                    run.italic = True
                if item.get("underline"):
                    run.underline = True
                if item.get("strikethrough"):
                    run.font.strike = True
                if item.get("font"):
                    run.font.name = item["font"]
                if item.get("size"):
                    run.font.size = Pt(item["size"])
                if item.get("color"):
                    color = validate_color(item["color"])
                    run.font.color.rgb = RGBColor(
                        int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
                    )

    elif item_type == "heading":
        text = sanitize_text(item.get("text", ""), "heading text")
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
                tbl.cell(ri, ci).text = "" if cell is None else sanitize_text(str(cell), "table cell")
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
            para = doc.add_paragraph(sanitize_text(str(line), "list item"))
            try:
                para.style = doc.styles[style]
            except KeyError:
                import logging
                logging.getLogger(__name__).debug("list style %r not found, using default", style)

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
    out = validate_path(path, expected_ext="docx", for_creation=True, operation="create")
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
    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="read")
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
    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="append")
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
    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="set_metadata")
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


def set_page_setup(
    path: str,
    *,
    size: str | None = None,
    width_mm: float | None = None,
    height_mm: float | None = None,
    top_mm: float | None = None,
    bottom_mm: float | None = None,
    left_mm: float | None = None,
    right_mm: float | None = None,
    orientation: str | None = None,
) -> str:
    """Set page size, margins, and orientation on a .docx file.

    ``size`` accepts named sizes: letter, a4, a3, a5, legal, tabloid.
    Or set ``width_mm`` / ``height_mm`` directly for custom sizes.
    Margins are in millimetres. ``orientation`` is 'portrait' or 'landscape'.
    """
    from .validation import validate_page_size

    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="set_page_setup")
    doc = Document(str(in_))

    for section in doc.sections:
        if size:
            w_mm, h_mm = validate_page_size(size)
            section.page_width = Mm(w_mm)
            section.page_height = Mm(h_mm)
        if width_mm is not None:
            section.page_width = Mm(width_mm)
        if height_mm is not None:
            section.page_height = Mm(height_mm)
        if top_mm is not None:
            section.top_margin = Mm(top_mm)
        if bottom_mm is not None:
            section.bottom_margin = Mm(bottom_mm)
        if left_mm is not None:
            section.left_margin = Mm(left_mm)
        if right_mm is not None:
            section.right_margin = Mm(right_mm)
        if orientation is not None:
            from docx.enum.section import WD_ORIENT
            o = orientation.lower().strip()
            if o not in ("portrait", "landscape"):
                raise ValueError(
                    f"Invalid orientation: '{orientation}'.\n"
                    f"Valid values: 'portrait', 'landscape'"
                )
            section.orientation = (
                WD_ORIENT.LANDSCAPE if o == "landscape" else WD_ORIENT.PORTRAIT
            )
            if o == "landscape":
                w, h = section.page_width, section.page_height
                if w < h:
                    section.page_width, section.page_height = h, w
            else:
                w, h = section.page_width, section.page_height
                if w > h:
                    section.page_width, section.page_height = h, w

    doc.save(str(in_))
    return str(in_)


def insert_paragraph(path: str, paragraph_index: int, content: Any) -> str:
    """Insert a content block before ``paragraph_index`` (0-based).

    ``content`` uses the same schema as items in ``docx_create``'s paragraphs
    list: a plain string, or a dict with type/text/style/etc.
    """
    from .validation import validate_paragraph_index

    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="insert_paragraph")
    doc = Document(str(in_))
    validate_paragraph_index(doc, paragraph_index)
    target = doc.paragraphs[paragraph_index]._element
    para_count_before = len(doc.paragraphs)
    _add_block(doc, content)
    if len(doc.paragraphs) > para_count_before:
        new_element = doc.paragraphs[-1]._element
        new_element.getparent().remove(new_element)
        target.addprevious(new_element)
    doc.save(str(in_))
    return str(in_)


def delete_paragraph(path: str, paragraph_index: int) -> str:
    """Delete a paragraph by index from a .docx file."""
    from .validation import validate_paragraph_index

    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="delete_paragraph")
    doc = Document(str(in_))
    validate_paragraph_index(doc, paragraph_index)
    p_element = doc.paragraphs[paragraph_index]._element
    p_element.getparent().remove(p_element)
    doc.save(str(in_))
    return str(in_)


def edit_paragraph(
    path: str,
    paragraph_index: int,
    *,
    text: str | None = None,
    style: str | None = None,
    align: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font: str | None = None,
    size: int | None = None,
    color: str | None = None,
) -> dict:
    """Edit an existing paragraph's text and formatting in place.

    Only the parameters you provide are changed — everything else is
    preserved. Returns the paragraph's state after editing.
    """
    from .validation import validate_paragraph_index

    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="edit_paragraph")
    doc = Document(str(in_))
    validate_paragraph_index(doc, paragraph_index)
    para = doc.paragraphs[paragraph_index]

    if text is not None:
        clean = sanitize_text(text, "paragraph text")
        for run in para.runs[1:]:
            run._element.getparent().remove(run._element)
        if para.runs:
            para.runs[0].text = clean
        else:
            para.add_run(clean)

    if style is not None:
        _apply_style(para, style, doc)

    if align is not None:
        validate_align(align)
        alignment = ALIGN_MAP.get(align.lower())
        if alignment is not None:
            para.alignment = alignment

    for run in para.runs:
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline
        if strikethrough is not None:
            run.font.strike = strikethrough
        if font is not None:
            run.font.name = sanitize_text(font, "font name")
        if size is not None:
            run.font.size = Pt(size)
        if color is not None:
            c = validate_color(color)
            run.font.color.rgb = RGBColor(
                int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            )

    doc.save(str(in_))

    result_runs = [_run_formatting(r) for r in para.runs]
    return {
        "path": str(in_),
        "paragraph_index": paragraph_index,
        "text": para.text,
        "style": para.style.name if para.style else "Normal",
        "alignment": _align_name(para.alignment),
        "runs": result_runs,
    }


def read_metadata(path: str) -> dict:
    """Read core document properties from a .docx file."""
    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="read_metadata")
    doc = Document(str(in_))
    cp = doc.core_properties
    return {
        "title": cp.title or "",
        "author": cp.author or "",
        "subject": cp.subject or "",
        "keywords": cp.keywords or "",
        "comments": cp.comments or "",
        "created": str(cp.created) if cp.created else "",
        "modified": str(cp.modified) if cp.modified else "",
        "last_modified_by": cp.last_modified_by or "",
        "revision": cp.revision,
        "category": cp.category or "",
    }


def _align_name(alignment) -> str:
    """Convert a WD_ALIGN_PARAGRAPH enum value to a friendly name."""
    _REVERSE_ALIGN = {v: k for k, v in ALIGN_MAP.items()}
    if alignment is None:
        return "left"
    return _REVERSE_ALIGN.get(alignment, str(alignment))


def _rgb_to_hex(rgb) -> str | None:
    """Convert an RGBColor to '#RRGGBB' string, or None."""
    if rgb is None:
        return None
    return f"#{str(rgb)}"


def _run_formatting(run) -> dict[str, Any]:
    """Extract formatting info from a single run."""
    fmt: dict[str, Any] = {"text": run.text}
    if run.bold:
        fmt["bold"] = True
    if run.italic:
        fmt["italic"] = True
    if run.underline:
        fmt["underline"] = True
    if run.font.strike:
        fmt["strikethrough"] = True
    if run.font.name:
        fmt["font"] = run.font.name
    if run.font.size:
        fmt["size_pt"] = run.font.size.pt
    color_rgb = run.font.color.rgb if run.font.color and run.font.color.rgb else None
    if color_rgb:
        fmt["color"] = _rgb_to_hex(color_rgb)
    return fmt


def get_formatting(path: str, paragraph_index: int) -> dict[str, Any]:
    """Inspect formatting of a specific paragraph and its runs."""
    from .validation import validate_paragraph_index

    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="get_formatting")
    doc = Document(str(in_))
    validate_paragraph_index(doc, paragraph_index)
    para = doc.paragraphs[paragraph_index]

    runs = [_run_formatting(r) for r in para.runs]
    style_name = para.style.name if para.style else "Normal"
    alignment = _align_name(para.alignment)
    validate_align(alignment)

    return {
        "paragraph_index": paragraph_index,
        "text": para.text,
        "style": style_name,
        "alignment": alignment,
        "runs": runs,
        "run_count": len(runs),
    }


def get_config(path: str) -> dict[str, Any]:
    """Detect full document configuration — page setup, styles, fonts, colors, sections."""
    from lxml import etree

    from .safety import safe_parse_xml
    from .validation import format_for_path

    in_ = validate_path(path, must_exist=True, expected_ext="docx", operation="get_config")
    doc = Document(str(in_))
    fmt = format_for_path(str(in_))

    sections_info = []
    for i, section in enumerate(doc.sections):
        sec = {
            "index": i,
            "page_width_mm": round(section.page_width / Mm(1), 1) if section.page_width else None,
            "page_height_mm": round(section.page_height / Mm(1), 1) if section.page_height else None,
            "top_margin_mm": round(section.top_margin / Mm(1), 1) if section.top_margin else None,
            "bottom_margin_mm": round(section.bottom_margin / Mm(1), 1) if section.bottom_margin else None,
            "left_margin_mm": round(section.left_margin / Mm(1), 1) if section.left_margin else None,
            "right_margin_mm": round(section.right_margin / Mm(1), 1) if section.right_margin else None,
        }
        orient = section.orientation
        sec["orientation"] = "landscape" if orient and orient == 1 else "portrait"
        sec["has_header"] = bool(section.header and section.header.paragraphs and
                                 any(p.text.strip() for p in section.header.paragraphs))
        sec["has_footer"] = bool(section.footer and section.footer.paragraphs and
                                 any(p.text.strip() for p in section.footer.paragraphs))
        if sec["has_header"]:
            sec["header_text"] = " ".join(p.text for p in section.header.paragraphs if p.text.strip())
        if sec["has_footer"]:
            sec["footer_text"] = " ".join(p.text for p in section.footer.paragraphs if p.text.strip())
        sections_info.append(sec)

    fonts_used: set[str] = set()
    colors_used: set[str] = set()
    styles_used: set[str] = set()

    for para in doc.paragraphs:
        if para.style:
            styles_used.add(para.style.name)
        for run in para.runs:
            if run.font.name:
                fonts_used.add(run.font.name)
            if run.font.color and run.font.color.rgb:
                hex_c = _rgb_to_hex(run.font.color.rgb)
                if hex_c:
                    colors_used.add(hex_c)

    styles_available = [
        {"name": s.name, "type": str(s.type)}
        for s in doc.styles
        if s.type == 1
    ][:30]

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = doc.element
    bg_el = root.find(f"{{{W_NS}}}background")
    background_color = None
    if bg_el is not None:
        background_color = bg_el.get(f"{{{W_NS}}}color")

    has_background_image = False
    for section in doc.sections:
        header_el = section.header._element
        for anchor in header_el.iter():
            if anchor.tag.endswith("}docPr") and anchor.get("name") == "PageBackground":
                has_background_image = True
                break

    has_watermark = False
    V_NS = "urn:schemas-microsoft-com:vml"
    for section in doc.sections:
        header_el = section.header._element
        for shape in header_el.iter(f"{{{V_NS}}}shape"):
            if shape.get("id") == "WatermarkShape":
                has_watermark = True
                break

    return {
        "format": fmt,
        "paragraph_count": len(doc.paragraphs),
        "section_count": len(doc.sections),
        "sections": sections_info,
        "fonts_used": sorted(fonts_used),
        "colors_used": sorted(colors_used),
        "styles_used": sorted(styles_used),
        "styles_available": styles_available,
        "background_color": background_color,
        "has_background_image": has_background_image,
        "has_watermark": has_watermark,
        "table_count": len(doc.tables),
        "metadata": read_metadata(path),
    }
