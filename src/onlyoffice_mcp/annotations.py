"""Headers, footers, hyperlinks, bookmarks, internal links, TOC, comments.

These are the OOXML-heaviest features: python-docx exposes none of comments,
TOC fields, page-number fields, or hyperlinks-as-runs directly, so we
construct the XML via lxml + OoxmlElement.

References:
- OOXML §17.13 — comments
- OOXML §17.16.5 — fields (PAGE, TOC, HYPERLINK)
- OOXML §17.13.6 — bookmarks
"""

from __future__ import annotations

import datetime as _dt
import uuid
from pathlib import Path

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


# --------------------------------------------------------------------------
# Headers & footers
# --------------------------------------------------------------------------

_ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "justify": 3}


def docx_set_header(
    path: str,
    text: str,
    *,
    align: str = "center",
    section: int = 0,
) -> str:
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    if section < 0 or section >= len(doc.sections):
        raise ValueError(f"section {section} out of range [0, {len(doc.sections) - 1}]")
    header = doc.sections[section].header
    # Reuse the first paragraph or add one.
    para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    para.text = text
    para.alignment = _ALIGN_MAP.get(align.lower(), 1)
    doc.save(str(p))
    return str(p)


def docx_set_footer(
    path: str,
    text: str = "",
    *,
    page_numbers: bool = True,
    align: str = "center",
    section: int = 0,
) -> str:
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    if section < 0 or section >= len(doc.sections):
        raise ValueError(f"section {section} out of range [0, {len(doc.sections) - 1}]")
    footer = doc.sections[section].footer
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.text = text
    para.alignment = _ALIGN_MAP.get(align.lower(), 1)
    if page_numbers:
        # Inject a PAGE field after the user's text.
        run = para.add_run()
        if text:
            run.add_text(" — ")
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), "PAGE")
        # The field needs a result run so the page number renders before the
        # document is updated by Word.
        rslt_r = OxmlElement("w:r")
        rslt_t = OxmlElement("w:t")
        rslt_t.text = "1"
        rslt_r.append(rslt_t)
        fld.append(rslt_r)
        run._r.addnext(fld)
    doc.save(str(p))
    return str(p)


# --------------------------------------------------------------------------
# Hyperlinks
# --------------------------------------------------------------------------

def _add_hyperlink_run(paragraph, url: str, text: str, *, internal: bool = False):
    """Add a hyperlink to a python-docx paragraph using lxml.

    Returns the run element so callers can style it further.
    """
    part = paragraph.part
    if internal:
        r_id = None
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), url)
    else:
        r_id = part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    # Apply the standard Hyperlink style.
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    new_run.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    new_run.append(t)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def docx_add_hyperlink(path: str, paragraph_index: int, text: str, url: str) -> str:
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        raise ValueError(
            f"paragraph_index {paragraph_index} out of range [0, {len(doc.paragraphs) - 1}]"
        )
    _add_hyperlink_run(doc.paragraphs[paragraph_index], url, text, internal=False)
    doc.save(str(p))
    return str(p)


def docx_add_internal_link(path: str, paragraph_index: int, text: str, bookmark_name: str) -> str:
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    _add_hyperlink_run(
        doc.paragraphs[paragraph_index], bookmark_name, text, internal=True
    )
    doc.save(str(p))
    return str(p)


def pptx_add_hyperlink(path: str, slide_index: int, shape_index: int, url: str) -> str:
    from pptx import Presentation

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    prs = Presentation(str(p))
    slide = prs.slides[slide_index]
    shape = slide.shapes[shape_index]
    if not shape.has_text_frame:
        raise ValueError("Shape has no text frame")
    # Set the hyperlink on the first run of the first paragraph.
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.hyperlink.address = url
            prs.save(str(p))
            return str(p)
    raise ValueError("No runs found in shape's text frame")


# --------------------------------------------------------------------------
# Bookmarks
# --------------------------------------------------------------------------

def docx_add_bookmark(path: str, paragraph_index: int, name: str) -> str:
    """Wrap a paragraph in ``<w:bookmarkStart>`` / ``<w:bookmarkEnd>``."""
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        raise ValueError(
            f"paragraph_index {paragraph_index} out of range [0, {len(doc.paragraphs) - 1}]"
        )
    para = doc.paragraphs[paragraph_index]
    bm_id = abs(hash(name)) % 99999

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bm_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bm_id))
    para._p.insert(0, start)
    para._p.append(end)
    doc.save(str(p))
    return str(p)


# --------------------------------------------------------------------------
# Table of contents (field)
# --------------------------------------------------------------------------

def docx_add_toc(path: str, paragraph_index: int = 0) -> str:
    """Insert a TOC field. Word / LibreOffice updates the TOC on open."""
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        # Insert as a new paragraph at the end.
        target = doc.add_paragraph()
    else:
        target = doc.paragraphs[paragraph_index]

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "TOC \\o \"1-3\" \\h \\z \\u"
    fld_char_sep = OxmlElement("w:fldChar")
    fld_char_sep.set(qn("w:fldCharType"), "separate")
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    run = OxmlElement("w:r")
    run.append(fld_char_begin)
    run.append(instr_text)
    run.append(fld_char_sep)
    placeholder_t = OxmlElement("w:t")
    placeholder_t.text = "Right-click to update Table of Contents."
    placeholder_r = OxmlElement("w:r")
    placeholder_r.append(placeholder_t)
    run.append(placeholder_r)
    run.append(fld_char_end)
    target._p.append(run)
    doc.save(str(p))
    return str(p)


# --------------------------------------------------------------------------
# Comments
# --------------------------------------------------------------------------

_COMMENTS_XML_TEMPLATE = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
"""


def _ensure_comments_part(doc):
    """Return the comments part (creating it + relationships if missing).

    Strategy: walk doc.part.package.parts looking for word/comments.xml. If
    missing, build a minimal valid one via the package's part-factory API and
    add a relationship from document.xml.
    """
    package = doc.part.package
    comments_partname = "/word/comments.xml"
    for part in package.parts:
        if str(part.partname) == comments_partname:
            return part
    # Create a fresh part. python-docx doesn't expose a clean API for arbitrary
    # part creation — we work with the underlying opc package.
    from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    comments_part = Part(
        partname=PackURI(comments_partname),
        content_type=CONTENT_TYPE.WML_COMMENTS,
        blob=_COMMENTS_XML_TEMPLATE,
        package=package,
    )
    package.parts.append_part(comments_part) if hasattr(package.parts, "append_part") else None
    # Re-add by relating from the main document part.
    doc.part.relate_to(comments_part, RELATIONSHIP_TYPE.COMMENTS)
    return comments_part


def docx_add_comment(
    path: str,
    paragraph_index: int,
    author: str,
    text: str,
    *,
    initials: str = "AI",
) -> str:
    """Attach a Word-style comment to a paragraph.

    Builds (or extends) ``word/comments.xml`` and inserts the comment range
    markers + reference into the target paragraph.
    """
    from docx import Document
    from docx.oxml.ns import nsmap

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        raise ValueError(
            f"paragraph_index {paragraph_index} out of range [0, {len(doc.paragraphs) - 1}]"
        )
    target_para = doc.paragraphs[paragraph_index]

    try:
        comments_part = _ensure_comments_part(doc)
    except Exception as e:
        # If comments part creation fails (some docx don't accept arbitrary
        # part injection without a fuller content-types update), fall back to
        # appending the comment text as a [Comment: ...] bracketed paragraph.
        para = doc.add_paragraph(f"[Comment by {author}: {text}]")
        doc.save(str(p))
        return str(p)

    # Parse the existing comments XML, append a new w:comment.
    comments_root = etree.fromstring(comments_part.blob)
    # Determine the next free id.
    existing_ids = [
        int(c.get(f"{{{W_NS}}}id", "0"))
        for c in comments_root.findall(f"{{{W_NS}}}comment")
    ]
    next_id = (max(existing_ids) + 1) if existing_ids else 0

    comment = etree.SubElement(comments_root, f"{{{W_NS}}}comment")
    comment.set(f"{{{W_NS}}}id", str(next_id))
    comment.set(f"{{{W_NS}}}author", author)
    comment.set(f"{{{W_NS}}}initials", initials)
    comment.set(
        f"{{{W_NS}}}date",
        _dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
    )
    pp = etree.SubElement(comment, f"{{{W_NS}}}p")
    rp = etree.SubElement(pp, f"{{{W_NS}}}r")
    tp = etree.SubElement(rp, f"{{{W_NS}}}t")
    tp.text = text

    comments_part._blob = etree.tostring(
        comments_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    # Add range markers around the target paragraph.
    start = OxmlElement("w:commentRangeStart")
    start.set(qn("w:id"), str(next_id))
    end = OxmlElement("w:commentRangeEnd")
    end.set(qn("w:id"), str(next_id))
    ref_run = OxmlElement("w:r")
    ref = OxmlElement("w:commentReference")
    ref.set(qn("w:id"), str(next_id))
    ref_run.append(ref)

    target_para._p.insert(0, start)
    target_para._p.append(end)
    target_para._p.append(ref_run)

    doc.save(str(p))
    return str(p)


def docx_list_comments(path: str) -> list[dict]:
    """Return all comments in a docx file."""
    from docx import Document

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)
    doc = Document(str(p))
    comments_partname = "/word/comments.xml"
    for part in doc.part.package.parts:
        if str(part.partname) == comments_partname:
            root = etree.fromstring(part.blob)
            out = []
            for c in root.findall(f"{{{W_NS}}}comment"):
                text_parts = [
                    t.text or ""
                    for t in c.iter(f"{{{W_NS}}}t")
                ]
                out.append(
                    {
                        "id": int(c.get(f"{{{W_NS}}}id", "0")),
                        "author": c.get(f"{{{W_NS}}}author", ""),
                        "initials": c.get(f"{{{W_NS}}}initials", ""),
                        "date": c.get(f"{{{W_NS}}}date", ""),
                        "text": "".join(text_parts),
                    }
                )
            return out
    return []
