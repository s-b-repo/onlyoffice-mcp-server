"""Page background, watermark, slide background and sheet tab colour.

DOCX page background and watermark require direct XML manipulation
(python-docx doesn't expose either). Slide background and sheet tab colour
have native APIs.

References:
- OOXML §17.2.1 — w:background element on w:document
- OOXML §17.15.1.20 — w:displayBackgroundShape in settings.xml
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"


def _normalise_color(color_hex: str) -> str:
    """Accept "#RRGGBB" or "RRGGBB"; return "RRGGBB" (no hash)."""
    h = color_hex.lstrip("#").strip().upper()
    if len(h) != 6:
        raise ValueError(f"Color must be 6-hex-digit RGB; got {color_hex!r}")
    int(h, 16)  # validate hex
    return h


# --------------------------------------------------------------------------
# DOCX page background
# --------------------------------------------------------------------------

def docx_set_background(path: str, color_hex: str) -> str:
    """Set the page background colour on a .docx file.

    Inserts ``<w:background w:color="RRGGBB"/>`` as the first child of
    ``<w:document>`` AND ``<w:displayBackgroundShape/>`` in settings.xml so
    LibreOffice and Word Online actually render the colour.
    """
    from docx import Document

    color = _normalise_color(color_hex)
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)

    doc = Document(str(p))
    root = doc.element  # <w:document>
    nsmap = {"w": W_NS}

    # Remove any existing w:background then prepend a new one.
    for existing in root.findall(f"{{{W_NS}}}background"):
        root.remove(existing)
    bg = etree.SubElement(root, f"{{{W_NS}}}background")
    bg.set(f"{{{W_NS}}}color", color)
    # SubElement appends; we need it as the first child.
    root.remove(bg)
    root.insert(0, bg)

    # Settings part: add w:displayBackgroundShape so the colour is shown.
    try:
        settings_part = doc.part.package.part_related_by_reltype(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
        )
    except Exception:
        settings_part = None
    if settings_part is not None:
        settings_root = etree.fromstring(settings_part.blob)
        if settings_root.find(f"{{{W_NS}}}displayBackgroundShape") is None:
            display = etree.SubElement(
                settings_root, f"{{{W_NS}}}displayBackgroundShape"
            )
            # Re-insert as an early child for canonical order.
            settings_root.remove(display)
            settings_root.insert(0, display)
            settings_part._blob = etree.tostring(
                settings_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

    doc.save(str(p))
    return str(p)


# --------------------------------------------------------------------------
# DOCX watermark (VML in header)
# --------------------------------------------------------------------------

_WATERMARK_VML_TEMPLATE = """
<w:p xmlns:w="{w_ns}" xmlns:v="{v_ns}" xmlns:o="{o_ns}">
  <w:r>
    <w:pict>
      <v:shape id="WatermarkShape" o:spid="_x0000_s1026" type="#_x0000_t136"
        style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;
               width:468pt;height:54pt;rotation:315;z-index:-251658240;
               mso-position-horizontal:center;mso-position-horizontal-relative:margin;
               mso-position-vertical:center;mso-position-vertical-relative:margin"
        fillcolor="#{color}" stroked="f">
        <v:textpath style="font-family:&quot;Calibri&quot;;font-size:{size}pt"
          string="{text}"/>
      </v:shape>
    </w:pict>
  </w:r>
</w:p>
"""


def docx_set_watermark(
    path: str,
    text: str,
    *,
    color_hex: str = "#BFBFBF",
    font_size: int = 144,
) -> str:
    """Add a diagonal text watermark to every section's header.

    Best-effort across renderers: Word renders fully; LibreOffice partially;
    Google Docs ignores VML.
    """
    from docx import Document

    color = _normalise_color(color_hex)
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)

    doc = Document(str(p))
    for section in doc.sections:
        header = section.header
        # Build the watermark paragraph and append.
        wm_xml = _WATERMARK_VML_TEMPLATE.format(
            w_ns=W_NS, v_ns=V_NS, o_ns=O_NS,
            color=color, size=font_size, text=text,
        ).strip()
        try:
            element = etree.fromstring(wm_xml)
            # Append to the header's body via the python-docx element API.
            header._element.append(element)
        except Exception:
            # Fall back to a simple text paragraph if the VML parse fails.
            para = header.add_paragraph(text)
            para.alignment = 1  # center
    doc.save(str(p))
    return str(p)


# --------------------------------------------------------------------------
# PPTX slide background
# --------------------------------------------------------------------------

def pptx_set_slide_background(
    path: str,
    slide_index: int | None = None,
    color_hex: str = "#FFFFFF",
) -> str:
    """Set a solid background colour on one slide (``slide_index``) or all
    slides (``slide_index=None``)."""
    from pptx import Presentation
    from pptx.dml.color import RGBColor

    color = _normalise_color(color_hex)
    rgb = RGBColor(int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))

    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)

    prs = Presentation(str(p))
    slides = prs.slides if slide_index is None else [prs.slides[slide_index]]
    for slide in slides:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb
    prs.save(str(p))
    return str(p)


# --------------------------------------------------------------------------
# XLSX sheet tab colour
# --------------------------------------------------------------------------

def xlsx_set_sheet_tab_color(path: str, sheet: str, color_hex: str) -> str:
    """Set the colour of a sheet's tab strip."""
    from openpyxl import load_workbook

    color = _normalise_color(color_hex)
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(p)

    wb = load_workbook(str(p))
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet}' not found. Available: {wb.sheetnames}")
    wb[sheet].sheet_properties.tabColor = color
    wb.save(str(p))
    return str(p)
