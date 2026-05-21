"""ONLYOFFICE MCP server v0.2.

Exposes tools for creating, reading, editing and converting OOXML documents
(docx, xlsx, pptx) compatible with ONLYOFFICE, Microsoft Office, LibreOffice,
and Google Docs. Includes:
- Edit history / version control with snapshots + revert
- AI virtual cursor (current paragraph / cell / slide)
- Spell-check and auto-correct
- Charts (matplotlib for docx, native for xlsx / pptx)
- Page background, watermarks, sheet tab colours
- Headers / footers / hyperlinks / bookmarks / TOC / comments
- Find & replace with run-formatting preservation
- Format conversion via ONLYOFFICE Document Builder OR LibreOffice headless

Transport: stdio (the standard for Claude Code / Claude Desktop).
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import (
    __version__,
    annotations,
    charts,
    converter,
    cursor,
    docbuilder,
    docx_ops,
    history,
    libreoffice,
    pptx_ops,
    search,
    spell,
    stats,
    storage,
    styling,
    xlsx_ops,
)

logger = logging.getLogger("onlyoffice-mcp")
mcp = FastMCP("onlyoffice")


# ---------------------------------------------------------------------------
# Word (docx)
# ---------------------------------------------------------------------------

@mcp.tool()
@history.record_operation("docx_create")
def docx_create(
    path: str,
    paragraphs: list[Any] | None = None,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
) -> str:
    """Create a Word document (.docx) at `path`.

    `paragraphs` is a list where each item is either a plain string (a body
    paragraph) or a dict describing a block:

      {"type": "paragraph", "text": "...", "style": "normal|heading1|...",
       "align": "left|center|right|justify", "bold": true, "italic": true,
       "size": 12, "color": "#ff0000"}
      {"type": "heading", "text": "...", "level": 1}
      {"type": "table", "data": [[row], [row]], "header": true,
       "style": "Light Grid Accent 1"}
      {"type": "image", "path": "...", "width_inches": 4}
      {"type": "list", "items": ["a", "b"], "ordered": false}
      {"type": "pagebreak"}

    Returns the absolute path of the created file.
    """
    return docx_ops.create(
        path, paragraphs or [], title=title, author=author, subject=subject
    )


@mcp.tool()
def docx_read(path: str, include_tables: bool = True) -> str:
    """Return the text content of a .docx file. Paragraphs are joined with
    newlines; tables are appended as pipe-separated rows when
    `include_tables` is true (default)."""
    return docx_ops.read(path, include_tables=include_tables)


@mcp.tool()
@history.record_operation("docx_append")
def docx_append(path: str, paragraphs: list[Any]) -> str:
    """Append content blocks to an existing .docx file. `paragraphs` uses the
    same schema as `docx_create`."""
    result = docx_ops.append(path, paragraphs)
    cursor.maybe_auto_advance(path, "docx_append")
    return result


@mcp.tool()
@history.record_operation("docx_set_metadata")
def docx_set_metadata(
    path: str,
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
) -> str:
    """Update the core document properties (Title / Author / Subject / Keywords
    / Comments) on an existing .docx file."""
    return docx_ops.set_metadata(
        path,
        title=title,
        author=author,
        subject=subject,
        keywords=keywords,
        comments=comments,
    )


# ---------------------------------------------------------------------------
# Excel (xlsx)
# ---------------------------------------------------------------------------

@mcp.tool()
@history.record_operation("xlsx_create")
def xlsx_create(
    path: str,
    sheets: dict[str, list[list[Any]]],
    header_bold: bool = True,
) -> str:
    """Create an Excel workbook (.xlsx) at `path`.

    `sheets` maps sheet name -> rows (list of row lists). Example:
        {"EMEA": [["Q1","Q2"],[100,200]], "AMER": [["Q1","Q2"],[300,400]]}

    If `header_bold` is true, the first row of each sheet is bolded.
    Returns the absolute path of the created file.
    """
    return xlsx_ops.create(path, sheets, header_bold=header_bold)


@mcp.tool()
def xlsx_read(path: str, sheet: str | None = None) -> Any:
    """Read a .xlsx workbook.

    If `sheet` is provided, returns `{"sheet": <name>, "rows": [[...]]}`.
    Otherwise returns `{sheet_name: [[...]]}` for every sheet."""
    return xlsx_ops.read(path, sheet=sheet)


@mcp.tool()
@history.record_operation("xlsx_append_rows")
def xlsx_append_rows(path: str, sheet: str, rows: list[list[Any]]) -> str:
    """Append rows to an existing sheet. Creates the sheet if it doesn't
    exist."""
    result = xlsx_ops.append_rows(path, sheet, rows)
    cursor.maybe_auto_advance(path, "xlsx_append_rows")
    return result


@mcp.tool()
@history.record_operation("xlsx_set_cell")
def xlsx_set_cell(path: str, sheet: str, cell: str, value: Any) -> str:
    """Set a single cell (e.g. "A1", "B2") in an existing workbook."""
    return xlsx_ops.set_cell(path, sheet, cell, value)


@mcp.tool()
def xlsx_list_sheets(path: str) -> list[str]:
    """Return the sheet names in a .xlsx workbook."""
    return xlsx_ops.list_sheets(path)


# ---------------------------------------------------------------------------
# PowerPoint (pptx)
# ---------------------------------------------------------------------------

@mcp.tool()
@history.record_operation("pptx_create")
def pptx_create(path: str, slides: list[dict]) -> str:
    """Create a PowerPoint presentation (.pptx) at `path`.

    Each slide dict supports:
      {"layout": "title", "title": "...", "subtitle": "..."}
      {"layout": "content", "title": "...", "body": ["bullet 1", "bullet 2"]}
      {"layout": "title_only", "title": "..."}
      {"layout": "image", "title": "...", "image_path": "...",
       "left_inches": 1, "top_inches": 1.5, "width_inches": 8}
      {"layout": "blank"}

    Any slide may include a "notes" key to add speaker notes.
    Returns the absolute path of the created file.
    """
    return pptx_ops.create(path, slides)


@mcp.tool()
def pptx_read(path: str) -> dict:
    """Extract slide titles, body text, and speaker notes from a .pptx file."""
    return pptx_ops.read(path)


@mcp.tool()
@history.record_operation("pptx_add_slide")
def pptx_add_slide(path: str, slide: dict) -> str:
    """Append a single slide to an existing .pptx file. Uses the same slide
    schema as `pptx_create`."""
    result = pptx_ops.add_slide(path, slide)
    cursor.maybe_auto_advance(path, "pptx_add_slide")
    return result


# ---------------------------------------------------------------------------
# ONLYOFFICE Document Builder bridge
# ---------------------------------------------------------------------------

@mcp.tool()
def docbuilder_status() -> dict:
    """Check whether ONLYOFFICE Document Builder is installed. Returns install
    state, binary path, version, and (if missing) install instructions."""
    return docbuilder.status()


@mcp.tool()
def docbuilder_run(script: str, output_path: str | None = None) -> str:
    """Execute an ONLYOFFICE Document Builder script (JavaScript-like .docbuilder
    syntax). Reference: https://api.onlyoffice.com/docbuilder/basic

    If `output_path` is provided and the script does not call
    `builder.SaveFile(...)` itself, a save call is appended automatically
    based on the output_path extension.

    Returns the resolved output path on success.
    """
    return docbuilder.run(script, output_path)


# ---------------------------------------------------------------------------
# Conversion & workspace utilities
# ---------------------------------------------------------------------------

@mcp.tool()
def convert(input_path: str, output_path: str) -> str:
    """Convert a document from one format to another. Output format is inferred
    from the output_path extension.

    With ONLYOFFICE Document Builder installed: full support for docx, xlsx,
    pptx, pdf, odt, ods, odp, rtf, txt, csv, html, epub.

    Without Document Builder: limited Python fallbacks — docx->txt, xlsx->csv,
    csv->xlsx, pptx->txt.
    """
    return converter.convert(input_path, output_path)


@mcp.tool()
def list_workspace(directory: str, recursive: bool = False) -> list[dict]:
    """List Office documents (docx, xlsx, pptx, odt, ods, odp, pdf, rtf, txt,
    csv) under `directory`. Returns each file's path, name, size, mtime, and
    extension."""
    base = Path(directory).expanduser().resolve()
    if not base.exists():
        raise ValueError(f"Directory does not exist: {base}")
    if not base.is_dir():
        raise ValueError(f"Not a directory: {base}")
    extensions = {
        ".docx", ".xlsx", ".pptx",
        ".odt", ".ods", ".odp",
        ".pdf", ".rtf", ".txt", ".csv", ".html", ".epub",
    }
    pattern = "**/*" if recursive else "*"
    out: list[dict] = []
    for p in base.glob(pattern):
        if not p.is_file():
            continue
        if p.suffix.lower() not in extensions:
            continue
        stat = p.stat()
        out.append(
            {
                "path": str(p),
                "name": p.name,
                "size": stat.st_size,
                "mtime": stat.st_mtime,
                "ext": p.suffix.lower(),
            }
        )
    out.sort(key=lambda x: x["name"])
    return out


# ---------------------------------------------------------------------------
# Version control / Edit history (Group A — 8 tools)
# ---------------------------------------------------------------------------

@mcp.tool()
def doc_history(path: str, limit: int = 20) -> list[dict]:
    """Return the most recent edits to a document (newest first). Each entry
    has revision, ts, tool, diff_summary, args_summary, snapshot_saved."""
    return history.list_history(path, limit=limit)


@mcp.tool()
def doc_history_show(path: str, revision: int) -> dict:
    """Return the full record of one revision, including its unified diff."""
    return history.show_revision(path, revision)


@mcp.tool()
def doc_last_edit(path: str) -> dict:
    """Return the most recent op record + the elapsed seconds since it ran."""
    return history.last_edit(path)


@mcp.tool()
def doc_what_was_removed(path: str, revision: int | None = None) -> dict:
    """Return ONLY the '-' lines from a revision's diff — the direct answer
    to 'what did I remove?'. Defaults to the most recent edit."""
    return history.what_was_removed(path, revision)


@mcp.tool()
def doc_diff(
    path: str,
    from_rev: int | None = None,
    to_rev: int | None = None,
) -> dict:
    """Unified text diff between two revisions (defaults: prev -> current)."""
    return history.diff_revisions(path, from_rev=from_rev, to_rev=to_rev)


@mcp.tool()
def doc_revert(path: str, revision: int) -> dict:
    """Restore the document from the snapshot at `revision`. The revert is
    itself recorded as a new edit."""
    return history.revert(path, revision)


@mcp.tool()
def doc_clear_history(path: str, keep_last: int = 0) -> dict:
    """Wipe edit history for a doc; returns bytes freed and kept revisions."""
    return history.clear_history(path, keep_last=keep_last)


@mcp.tool()
def doc_history_stats() -> dict:
    """Return total tracked docs, total disk usage, oldest entry, etc."""
    return history.history_stats()


# ---------------------------------------------------------------------------
# Page tracking & navigation (Group B — 3 tools)
# ---------------------------------------------------------------------------

@mcp.tool()
def doc_get_cursor(path: str) -> dict:
    """Return the current AI virtual cursor + per-format bounds. The cursor
    is flagged stale if the document was edited externally and the cursor
    is now out of bounds."""
    return cursor.get_cursor(path)


@mcp.tool()
def doc_set_cursor(
    path: str,
    paragraph_index: int | None = None,
    page: int | None = None,
    sheet: str | None = None,
    cell: str | None = None,
    row: int | None = None,
    col: int | None = None,
    slide_index: int | None = None,
    clamp: bool = True,
) -> dict:
    """Move the AI virtual cursor.

    docx: pass paragraph_index or page. xlsx: pass sheet + cell (or row+col).
    pptx: pass slide_index. Values are clamped to document bounds by default;
    a warning is returned if clamping was needed."""
    return cursor.set_cursor(
        path,
        paragraph_index=paragraph_index,
        page=page,
        sheet=sheet,
        cell=cell,
        row=row,
        col=col,
        slide_index=slide_index,
        clamp=clamp,
    )


@mcp.tool()
def doc_page_count(path: str, precise: bool = False) -> dict:
    """Page / sheet / slide count, per format.

    docx: approximate (count of <w:br w:type='page'/> + <w:lastRenderedPageBreak/>).
    When precise=True and LibreOffice is installed, converts to PDF in a
    tempdir and counts via pypdf (~5s, exact)."""
    return cursor.page_count(path, precise=precise)


# ---------------------------------------------------------------------------
# Spell-check / Auto-correct (Group C — 3 tools)
# ---------------------------------------------------------------------------

@mcp.tool()
def doc_spell_check(
    path: str,
    language: str = "en",
    max_words: int = 200,
) -> dict:
    """Check a document for spelling errors. Returns the engine used plus a
    list of {word, suggestions, location, context} entries."""
    return spell.check_document(path, language=language, max_words=max_words)


@mcp.tool()
@history.record_operation("doc_apply_corrections")
def doc_apply_corrections(
    path: str,
    corrections: dict[str, str],
    scope: str = "all",
) -> dict:
    """Apply {misspelled: correction} substitutions in a document, preserving
    run formatting. Returns {applied: N, skipped_multi_run: M}."""
    return spell.apply_corrections(path, corrections, scope=scope)


@mcp.tool()
def spell_suggest(word: str, language: str = "en", max: int = 5) -> dict:
    """Look up suggestions for a single word (no document needed)."""
    return spell.suggest_single(word, language=language, max=max)


# ---------------------------------------------------------------------------
# Page background, watermark, slide / sheet styling (Group D — 4 tools)
# ---------------------------------------------------------------------------

@mcp.tool()
@history.record_operation("docx_set_background")
def docx_set_background(path: str, color_hex: str) -> str:
    """Set the page background colour on a .docx file. Hex like '#FFE6E6' or
    'FFE6E6'. Also enables <w:displayBackgroundShape/> in settings.xml so
    LibreOffice / Word Online render the colour."""
    return styling.docx_set_background(path, color_hex)


@mcp.tool()
@history.record_operation("docx_set_watermark")
def docx_set_watermark(
    path: str,
    text: str,
    color_hex: str = "#BFBFBF",
    font_size: int = 144,
) -> str:
    """Add a diagonal text watermark to every section header. Best-effort
    across renderers; Word renders fully, LibreOffice partially."""
    return styling.docx_set_watermark(
        path, text, color_hex=color_hex, font_size=font_size
    )


@mcp.tool()
@history.record_operation("pptx_set_slide_background")
def pptx_set_slide_background(
    path: str,
    slide_index: int | None = None,
    color_hex: str = "#FFFFFF",
) -> str:
    """Set a solid background colour on one slide or all slides
    (slide_index=None)."""
    return styling.pptx_set_slide_background(path, slide_index, color_hex)


@mcp.tool()
@history.record_operation("xlsx_set_sheet_tab_color")
def xlsx_set_sheet_tab_color(path: str, sheet: str, color_hex: str) -> str:
    """Set the colour of a sheet's tab strip."""
    return styling.xlsx_set_sheet_tab_color(path, sheet, color_hex)


# ---------------------------------------------------------------------------
# Headers / footers / hyperlinks / bookmarks / TOC / comments (Group E)
# ---------------------------------------------------------------------------

@mcp.tool()
@history.record_operation("docx_set_header")
def docx_set_header(
    path: str,
    text: str,
    align: str = "center",
    section: int = 0,
) -> str:
    """Set the header text on a section. align in {left, center, right, justify}."""
    return annotations.docx_set_header(path, text, align=align, section=section)


@mcp.tool()
@history.record_operation("docx_set_footer")
def docx_set_footer(
    path: str,
    text: str = "",
    page_numbers: bool = True,
    align: str = "center",
    section: int = 0,
) -> str:
    """Set the footer text and optionally include a PAGE field for page numbers."""
    return annotations.docx_set_footer(
        path, text, page_numbers=page_numbers, align=align, section=section
    )


@mcp.tool()
@history.record_operation("docx_add_hyperlink")
def docx_add_hyperlink(
    path: str,
    paragraph_index: int,
    text: str,
    url: str,
) -> str:
    """Append a clickable external hyperlink to a paragraph."""
    return annotations.docx_add_hyperlink(path, paragraph_index, text, url)


@mcp.tool()
@history.record_operation("docx_add_internal_link")
def docx_add_internal_link(
    path: str,
    paragraph_index: int,
    text: str,
    bookmark_name: str,
) -> str:
    """Append a clickable internal hyperlink to a bookmark in the same document."""
    return annotations.docx_add_internal_link(
        path, paragraph_index, text, bookmark_name
    )


@mcp.tool()
@history.record_operation("docx_add_bookmark")
def docx_add_bookmark(path: str, paragraph_index: int, name: str) -> str:
    """Wrap a paragraph in a named bookmark for internal links to target."""
    return annotations.docx_add_bookmark(path, paragraph_index, name)


@mcp.tool()
@history.record_operation("docx_add_toc")
def docx_add_toc(path: str, paragraph_index: int = 0) -> str:
    """Insert a Table of Contents field. Word/LibreOffice updates it on open."""
    return annotations.docx_add_toc(path, paragraph_index)


@mcp.tool()
@history.record_operation("docx_add_comment")
def docx_add_comment(
    path: str,
    paragraph_index: int,
    author: str,
    text: str,
    initials: str = "AI",
) -> str:
    """Attach a Word-style comment to a paragraph."""
    return annotations.docx_add_comment(
        path, paragraph_index, author, text, initials=initials
    )


@mcp.tool()
def docx_list_comments(path: str) -> list[dict]:
    """List all comments in a .docx file."""
    return annotations.docx_list_comments(path)


@mcp.tool()
@history.record_operation("pptx_add_hyperlink")
def pptx_add_hyperlink(
    path: str,
    slide_index: int,
    shape_index: int,
    url: str,
) -> str:
    """Add a hyperlink to the first run of a shape's text frame."""
    return annotations.pptx_add_hyperlink(path, slide_index, shape_index, url)


# ---------------------------------------------------------------------------
# Charts (Group F — 4 tools)
# ---------------------------------------------------------------------------

@mcp.tool()
@history.record_operation("xlsx_add_chart")
def xlsx_add_chart(
    path: str,
    sheet: str,
    chart_type: str,
    data_range: str,
    categories_range: str | None = None,
    anchor_cell: str = "E2",
    title: str | None = None,
) -> str:
    """Add a native chart to an .xlsx sheet. chart_type in
    {bar, line, pie, scatter, area}. data_range and categories_range are
    A1-style refs."""
    return charts.xlsx_add_chart(
        path,
        sheet,
        chart_type,
        data_range,
        categories_range=categories_range,
        anchor_cell=anchor_cell,
        title=title,
    )


@mcp.tool()
@history.record_operation("pptx_add_chart")
def pptx_add_chart(
    path: str,
    slide_index: int,
    chart_type: str,
    categories: list[Any],
    series: list[dict],
    left_inches: float = 1.0,
    top_inches: float = 2.0,
    width_inches: float = 8.0,
    height_inches: float = 5.0,
    title: str | None = None,
) -> str:
    """Add a native chart to a slide. For bar/line/pie/area each series is
    {name, values}; for scatter each series is {name, x, y}."""
    return charts.pptx_add_chart(
        path,
        slide_index,
        chart_type,
        categories,
        series,
        left_inches=left_inches,
        top_inches=top_inches,
        width_inches=width_inches,
        height_inches=height_inches,
        title=title,
    )


@mcp.tool()
@history.record_operation("docx_add_chart")
def docx_add_chart(
    path: str,
    chart_type: str,
    categories: list[Any],
    series: list[dict],
    title: str | None = None,
    width_inches: float = 6.0,
    height_inches: float = 4.0,
    paragraph_index: int | None = None,
) -> str:
    """Render a chart with matplotlib and embed it as a static image in a
    .docx. Note: result is a static image — NOT an editable chart."""
    return charts.docx_add_chart(
        path,
        chart_type,
        categories,
        series,
        title=title,
        width_inches=width_inches,
        height_inches=height_inches,
        paragraph_index=paragraph_index,
    )


@mcp.tool()
def chart_kinds() -> dict:
    """Return supported chart kinds, synonyms, and per-format support."""
    return charts.chart_kinds_info()


# ---------------------------------------------------------------------------
# Find & replace (Group G — 2 tools, polymorphic)
# ---------------------------------------------------------------------------

@mcp.tool()
def doc_find(
    path: str,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = True,
    include_notes: bool = True,
) -> list[dict]:
    """Find every occurrence of a pattern in a document. Returns list of
    {location, match_text, context}. Location shape is format-specific."""
    return search.find_in_document(
        path,
        pattern,
        regex=regex,
        case_sensitive=case_sensitive,
        include_notes=include_notes,
    )


@mcp.tool()
@history.record_operation("doc_replace")
def doc_replace(
    path: str,
    find: str,
    replace: str,
    regex: bool = False,
    case_sensitive: bool = True,
    count: int | None = None,
    dry_run: bool = False,
    include_formulas: bool = False,
    include_notes: bool = True,
) -> dict:
    """Replace text in a document. Preserves run-level formatting. Use
    dry_run=True to see what would change without modifying the file."""
    return search.replace_in_document(
        path,
        find,
        replace,
        regex=regex,
        case_sensitive=case_sensitive,
        count=count,
        dry_run=dry_run,
        include_formulas=include_formulas,
        include_notes=include_notes,
    )


# ---------------------------------------------------------------------------
# Document stats (Group H — 1 tool, polymorphic)
# ---------------------------------------------------------------------------

@mcp.tool()
def doc_stats(path: str) -> dict:
    """Polymorphic document statistics: word/char/page counts for docx,
    sheet/cell counts for xlsx, slide/shape counts for pptx."""
    return stats.stats(path)


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

@mcp.tool()
def server_info() -> dict:
    """Return server version and capabilities (used for diagnostics)."""
    db = docbuilder.status()
    lo = libreoffice.status()
    try:
        hist_stats = history.history_stats()
    except Exception as e:
        hist_stats = {"error": str(e)}
    return {
        "name": "onlyoffice-mcp",
        "version": __version__,
        "pid": os.getpid(),
        "python": os.sys.version.split()[0],
        "docbuilder": db,
        "libreoffice": lo,
        "history": hist_stats,
        "workspace": str(storage.home()),
        "formats": {
            "create": ["docx", "xlsx", "pptx"],
            "read": ["docx", "xlsx", "pptx"],
            "convert_with_docbuilder": db.get("supported_formats", []) if db.get("installed") else [],
            "convert_with_libreoffice": (
                ["docx", "xlsx", "pptx", "pdf", "odt", "ods", "odp", "rtf", "txt", "csv", "html"]
                if lo.get("installed") else []
            ),
            "convert_python_fallback": ["docx->txt", "xlsx->csv", "csv->xlsx", "pptx->txt"],
        },
        "chart_kinds": charts.chart_kinds_info(),
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    """Run the MCP server on stdio."""
    log_level = os.environ.get("ONLYOFFICE_MCP_LOG_LEVEL", "INFO").upper()
    logging.basicConfig(
        level=getattr(logging, log_level, logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        # MCP uses stdout for the protocol; logs MUST go to stderr.
        stream=os.sys.stderr,
    )
    logger.info("Starting OnlyOffice MCP server v%s", __version__)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
