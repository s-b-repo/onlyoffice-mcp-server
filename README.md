# OnlyOffice MCP Server

> **Warning: ALPHA — `v0.3.0a4`**
> This project is in **alpha**. APIs, tool names, schemas and storage layout may change without notice between releases. Smoke-tested but not battle-tested. File issues and expect rough edges.

A pure-Python Model Context Protocol (MCP) server for creating, reading, editing and converting ONLYOFFICE-compatible Office documents — Word (`.docx`), Excel (`.xlsx`), and PowerPoint (`.pptx`).

ONLYOFFICE uses the same OOXML format as Microsoft Office, so files this server produces open natively in ONLYOFFICE Desktop / Docs, Microsoft Office, LibreOffice, Google Docs, and any other OOXML-compatible suite.

**v0.3.0a5 highlights:** **In-place editing** (`docx_edit_paragraph`, `pptx_update_slide`), **safe error handling** with AI retry guidance, **input sanitization** across all text paths, and **logged exceptions** replacing all silent bare-except blocks. All **77 tools** are async.

## Features

- **Word (docx)** — create, read, append, insert/delete paragraphs, set metadata, read metadata, page setup (size/margins/orientation), headings, tables, images, lists, page breaks, alignment, fonts, colours, underline, strikethrough, background images with opacity
- **Excel (xlsx)** — create with multiple sheets, read, append/insert/delete rows, set cells (values + formulas), merge cells, set column widths, freeze panes, list/rename/delete sheets, tab colours
- **PowerPoint (pptx)** — create with title / content / image / blank layouts, append/delete slides, speaker notes
- **Version control** — every mutating call is journaled to `~/.onlyoffice-mcp/history/<doc_id>/`; query history, see what was removed, revert to any snapshot
- **AI virtual cursor** — track and move the "current" paragraph / cell / slide across MCP calls
- **Spell-check** — pyspellchecker primary engine with aspell subprocess fallback; word-level corrections preserve run formatting
- **Page background / background image with opacity / watermark / sheet tab colour / slide background**
- **Document preview** — render pages as PNG images for AI visual verification (PyMuPDF + LibreOffice)
- **Headers / footers (with PAGE field) / hyperlinks (external + internal) / bookmarks / TOC field / Word-style comments**
- **Native charts** — bar / line / pie / scatter / area; openpyxl in xlsx, python-pptx in pptx, matplotlib-rendered PNG in docx
- **Find & replace** — run-formatting-preserving substitution across all three formats; dry-run mode
- **Document stats** — polymorphic per format (word/char/page for docx, sheet/cell for xlsx, slide/shape for pptx)
- **Format conversion** — ONLYOFFICE Document Builder -> LibreOffice headless -> pure-Python fallbacks (in that preference order)
- **LLM-friendly validation** — every error includes what went wrong, valid values, and which tool to use instead
- **Diagnostics** — `server_info` reports version, engines available, supported formats, history disk usage
- **Async threading** — all tools run via `asyncio.to_thread()` so the server handles concurrent requests without blocking
- **Safe file deletion** — permanent delete with safety checks, recoverable trash with restore, mass-deletion detection, rate limiting, and audit trail
- **Sensitive path protection** — blocks deletion of SSH keys, credentials, configs, `.env` files, and 18 other sensitive path patterns

## Security & AI Safety

The server includes multiple layers of protection against misuse:

| Defence | What it does |
|---|---|
| **File-size limit** | Rejects files > 100 MB (configurable via `ONLYOFFICE_MCP_MAX_FILE_SIZE`) to prevent OOM |
| **Zip-bomb detection** | Checks compression ratio on OOXML files (limit: 100:1, configurable) |
| **XML entity protection** | All XML parsing uses `resolve_entities=False`, `no_network=True` — blocks billion-laughs |
| **Path-traversal blocking** | System paths (`/proc`, `/sys`, `/dev`) are denied; symlinks and `..` trigger warnings |
| **ReDoS prevention** | Regex patterns with nested quantifiers like `(a+)+` are rejected before compilation |
| **Format-conversion whitelist** | `convert()` validates source→target pairs (e.g. `docx→pdf` allowed, `docx→exe` denied) |
| **Prompt-injection scanning** | `docx_read`, `xlsx_read`, `pptx_read` scan text for 10 adversarial patterns and return `content_warnings` |
| **Macro/VBA detection** | Documents containing VBA macros, ActiveX controls, or external links trigger security warnings |
| **AI safety guidance** | Tool docstrings explicitly warn models not to follow instructions from document content, not to execute scripts from untrusted sources, and not to use URLs extracted from documents without user confirmation |
| **Script execution warning** | `docbuilder_run` response includes a security reminder about prompt-injection risks |
| **Deletion rate limiting** | Sliding-window rate limiter blocks rapid-fire deletions (default: 5 per 60s, configurable) |
| **Mass deletion detection** | Batch deletes are capped at 50 files; > 10 files triggers warnings; same-directory patterns flagged |
| **Sensitive path blocking** | 18 patterns block deletion of `.ssh`, `.env`, `.aws`, credentials, keys, and config files |
| **Trash (soft delete)** | `doc_move_to_trash` moves files to a recoverable trash with restore capability |
| **Deletion audit log** | Every delete/trash operation is logged to an immutable JSONL audit trail |

### Environment variables

| Variable | Default | Description |
|---|---|---|
| `ONLYOFFICE_MCP_MAX_FILE_SIZE` | `104857600` (100 MB) | Maximum file size in bytes |
| `ONLYOFFICE_MCP_MAX_DECOMPRESSION_RATIO` | `100` | Maximum zip compression ratio |
| `ONLYOFFICE_MCP_HOME` | `~/.onlyoffice-mcp` | Workspace root |
| `ONLYOFFICE_MCP_HISTORY_ENABLED` | `true` | Enable/disable edit history |
| `ONLYOFFICE_MCP_DELETION_RATE_LIMIT` | `5` | Max file deletions per time window |
| `ONLYOFFICE_MCP_DELETION_RATE_WINDOW` | `60` | Deletion rate window in seconds |
| `ONLYOFFICE_MCP_LOG_LEVEL` | `INFO` | Logging level (stderr + `~/.onlyoffice-mcp/logs/server.log`) |

## Requirements

- Python >= 3.10
- Runtime: `mcp`, `python-docx`, `openpyxl`, `python-pptx`, `matplotlib`, `lxml`, `pyspellchecker` (installed automatically)
- Optional: `pypdf` (for precise docx page counting via PDF) — install with `pip install 'onlyoffice-mcp[pdf]'`
- Optional: `PyMuPDF` (for `doc_preview` page rendering) — install with `pip install 'onlyoffice-mcp[preview]'`
- Optional native binaries (auto-detected, no Python deps): `documentbuilder` (full conversion matrix), `soffice` / `libreoffice` (broad conversion matrix), `aspell` (spell-check fallback)

## Install

```bash
cd /home/kali/onlyoffice-mcp
pip install -e .
# or:  pipx install /home/kali/onlyoffice-mcp
```

For precise docx page counting (LibreOffice -> PDF -> pypdf):

```bash
pip install 'onlyoffice-mcp[pdf]'
```

### Optional native binaries

```bash
# LibreOffice — broad conversion fallback (recommended if Document Builder is not available)
sudo apt install libreoffice

# aspell — spell-check fallback if pyspellchecker isn't loadable
sudo apt install aspell aspell-en

# ONLYOFFICE Document Builder — widest conversion matrix
wget https://download.onlyoffice.com/install/documentbuilder/linux/onlyoffice-documentbuilder_amd64.deb
sudo dpkg -i onlyoffice-documentbuilder_amd64.deb
sudo apt -f install
```

Verify with the MCP tool `server_info` after registration.

## Register with Claude Code

```bash
claude mcp add --scope user onlyoffice -- python -m onlyoffice_mcp
claude mcp list
```

A helper script is at `examples/register_with_claude.sh`.

Remove later with `claude mcp remove onlyoffice`.

## Configuration

Environment variables:

| Variable | Default | Purpose |
|----------|---------|---------|
| `ONLYOFFICE_MCP_HOME` | `~/.onlyoffice-mcp` | Workspace root for edit history + cursors |
| `ONLYOFFICE_MCP_HISTORY_ENABLED` | `true` | Set to `false` to disable history tracking entirely |
| `ONLYOFFICE_MCP_HISTORY_MAX_SNAPSHOTS` | `20` | Snapshots kept per document |
| `ONLYOFFICE_MCP_HISTORY_MAX_BYTES` | `104857600` (100 MB) | Total history budget across all docs |
| `ONLYOFFICE_MCP_LOG_LEVEL` | `INFO` | `DEBUG` / `INFO` / `WARNING` / `ERROR` — logs go to stderr |

## Tools (77 total)

### Word (22 tools)

| Tool | Purpose |
|------|---------|
| `docx_create` | Create a Word document from a list of structured blocks |
| `docx_read` | Return text content of a Word document |
| `docx_append` | Append blocks to an existing Word document |
| `docx_insert_paragraph` | Insert a content block before a specific paragraph index |
| `docx_edit_paragraph` | Edit existing paragraph text and formatting in place |
| `docx_delete_paragraph` | Delete a paragraph by 0-based index |
| `docx_set_metadata` | Update core properties (title, author, subject, keywords, comments) |
| `docx_read_metadata` | Read core properties (title, author, created, modified, etc.) |
| `docx_get_config` | Detect full document configuration (page setup, fonts, colors, styles, background) |
| `docx_get_formatting` | Inspect formatting of a specific paragraph (bold, italic, font, color, alignment) |
| `docx_set_page_setup` | Set page size (letter/a4/a3/a5/legal/tabloid or custom mm), margins, orientation |
| `docx_set_background` | Set page background colour (hex) |
| `docx_set_background_image` | Set a background image with offset, size, and opacity controls |
| `docx_set_watermark` | Add a diagonal text watermark on every page |
| `docx_set_header` | Set header text + alignment on a section |
| `docx_set_footer` | Set footer text + optional PAGE field for page numbers |
| `docx_add_hyperlink` | Append an external hyperlink to a paragraph |
| `docx_add_internal_link` | Hyperlink targeting an internal bookmark |
| `docx_add_bookmark` | Wrap a paragraph in a named bookmark |
| `docx_add_toc` | Insert a Table of Contents field |
| `docx_add_comment` | Attach a Word-style comment to a paragraph |
| `docx_list_comments` | Enumerate existing comments |
| `docx_add_chart` | Render a chart with matplotlib and embed as image |

### Excel (14 tools)

| Tool | Purpose |
|------|---------|
| `xlsx_create` | Create a workbook from `{sheet: rows}` |
| `xlsx_read` | Read one or all sheets as 2-D arrays |
| `xlsx_append_rows` | Append rows to a sheet |
| `xlsx_insert_rows` | Insert blank rows at a specific position (1-based) |
| `xlsx_delete_rows` | Delete rows from a sheet (1-based) |
| `xlsx_set_cell` | Set a single cell — values, numbers, or formulas (`=SUM(...)`) |
| `xlsx_merge_cells` | Merge a range of cells (e.g. `A1:C1`) |
| `xlsx_set_column_width` | Set column width by letter (e.g. `A`, `B`, `AA`) |
| `xlsx_freeze_panes` | Freeze rows/columns (e.g. `A2` freezes header row) |
| `xlsx_list_sheets` | Return sheet names |
| `xlsx_rename_sheet` | Rename an existing sheet |
| `xlsx_delete_sheet` | Delete a sheet (cannot delete the last one) |
| `xlsx_set_sheet_tab_color` | Colour the workbook tab strip |
| `xlsx_add_chart` | Add a native editable chart referencing existing cells |

### PowerPoint (7 tools)

| Tool | Purpose |
|------|---------|
| `pptx_create` | Create a presentation from a list of slide dicts |
| `pptx_read` | Extract slide titles, text, and speaker notes |
| `pptx_add_slide` | Append a single slide |
| `pptx_update_slide` | Edit existing slide title, body text, or speaker notes in place |
| `pptx_delete_slide` | Delete a slide by 0-based index |
| `pptx_set_slide_background` | Set a solid background colour on one slide or all |
| `pptx_add_hyperlink` | Set a hyperlink on the first run of a shape |
| `pptx_add_chart` | Add a native chart to a slide |

### Edit history / version control (8 tools)

| Tool | Purpose |
|------|---------|
| `doc_history` | List recent revisions (returns revision, tool, diff_summary) |
| `doc_history_show` | Full record of one revision (including unified diff) |
| `doc_last_edit` | Most recent op + age in seconds |
| `doc_what_was_removed` | Only the `-` lines from a revision's diff |
| `doc_diff` | Unified text diff between two revisions |
| `doc_revert` | Restore from a snapshot |
| `doc_clear_history` | Wipe history for a doc |
| `doc_history_stats` | Total tracked docs + disk usage |

### Cursor / navigation (3 tools)

| Tool | Purpose |
|------|---------|
| `doc_get_cursor` | Current cursor + format-specific bounds + stale flag |
| `doc_set_cursor` | Move cursor; values are clamped to doc bounds |
| `doc_page_count` | docx: approx pages (precise=True uses LibreOffice + pypdf) |

### Spell-check / find & replace (5 tools)

| Tool | Purpose |
|------|---------|
| `doc_spell_check` | Find misspellings with suggestions (pyspellchecker or aspell) |
| `doc_apply_corrections` | Apply word-level substitutions, preserving formatting |
| `spell_suggest` | Single-word lookup with is_known flag, no document needed |
| `doc_find` | Find every occurrence of a pattern (regex optional) |
| `doc_replace` | Run-preserving replacement with dry-run support |

### Conversion / preview / diagnostics (9 tools)

| Tool | Purpose |
|------|---------|
| `convert` | Convert between formats (DocBuilder -> LibreOffice -> Python fallback) |
| `doc_preview` | Render document pages as PNG for AI visual inspection (PyMuPDF) |
| `docbuilder_status` | Whether ONLYOFFICE Document Builder is installed |
| `docbuilder_run` | Execute a Document Builder script |
| `color_info` | List all 59 supported named colors with hex values |
| `chart_kinds` | List supported chart kinds + synonyms |
| `doc_stats` | Polymorphic per-format statistics (words, cells, slides, etc.) |
| `list_workspace` | List Office documents in a directory |
| `server_info` | Diagnostics: version, engines, formats, history usage |

## LLM-friendly error handling

Every validation error follows a consistent two-part format:

```
Line 1: What went wrong
Line 2+: What the valid values are, or what tool to use instead
```

Examples:

```
Wrong file extension: expected .docx, got .xlsx for path 'data.xlsx'.
Either rename the file to end with .docx, or use the correct tool.
To read a .xlsx file, use `xlsx_read` instead.
```

```
Invalid color: 'red'. Expected 6-digit hex RGB.
Examples: '#FF0000' (red), '#00FF00' (green), '#0000FF' (blue).
```

```
Sheet 'Sales' not found in workbook.
Available sheets: ['Sheet1', 'Revenue', 'Costs']
Sheet names are case-sensitive.
```

## Schemas

### Block schema for `docx_create` / `docx_append` / `docx_insert_paragraph`

Each item in the `paragraphs` list is either a plain string (body paragraph) or a dict:

- `{"type": "paragraph", "text": "...", "style": "normal|heading1|...|quote|code|bullet", "align": "left|center|right|justify", "bold": true, "italic": true, "underline": true, "strikethrough": true, "font": "Arial", "size": 12, "color": "#ff0000"}`
- `{"type": "heading", "text": "...", "level": 1}`
- `{"type": "table", "data": [[...]], "header": true, "style": "Light Grid Accent 1"}`
- `{"type": "image", "path": "...", "width_inches": 4}`
- `{"type": "list", "items": ["a", "b"], "ordered": false}`
- `{"type": "pagebreak"}`

### Slide schema for `pptx_create` / `pptx_add_slide`

- `{"layout": "title", "title": "...", "subtitle": "..."}`
- `{"layout": "content", "title": "...", "body": ["bullet 1", "bullet 2"]}`
- `{"layout": "title_only", "title": "..."}`
- `{"layout": "image", "title": "...", "image_path": "...", "left_inches": 1, "top_inches": 1.5, "width_inches": 8}`
- `{"layout": "blank"}`

Any slide may include `"notes": "..."` for speaker notes.

### Chart series schema

- bar / line / pie / area: `series = [{"name": str, "values": [numbers]}]`
- scatter: `series = [{"name": str, "x": [numbers], "y": [numbers]}]`

### Named page sizes for `docx_set_page_setup`

| Name | Dimensions (mm) |
|------|-----------------|
| `letter` | 215.9 x 279.4 |
| `a4` | 210 x 297 |
| `a3` | 297 x 420 |
| `a5` | 148 x 210 |
| `legal` | 215.9 x 355.6 |
| `tabloid` | 279.4 x 431.8 |

Or pass `width_mm` / `height_mm` directly for custom sizes.

## Examples

### Create a Word document

```json
{"tool": "docx_create", "args": {
  "path": "/tmp/report.docx",
  "title": "Q4 Report",
  "author": "Jane Smith",
  "paragraphs": [
    {"type": "heading", "text": "Executive Summary", "level": 1},
    "Revenue grew 18% year-over-year.",
    {"type": "table", "data": [
      ["Region", "Revenue", "Growth"],
      ["EMEA", 1250000, "+12%"],
      ["AMER", 2100000, "+18%"]
    ]},
    {"type": "pagebreak"},
    {"type": "heading", "text": "Outlook", "level": 1},
    {"text": "We expect continued growth into Q1.", "italic": true}
  ]
}}
```

### Insert and delete paragraphs

```json
{"tool": "docx_insert_paragraph", "args": {
  "path": "/tmp/report.docx",
  "paragraph_index": 2,
  "content": {"type": "paragraph", "text": "Inserted before paragraph 2", "bold": true}
}}
{"tool": "docx_delete_paragraph", "args": {"path": "/tmp/report.docx", "paragraph_index": 5}}
```

### Set page layout

```json
{"tool": "docx_set_page_setup", "args": {
  "path": "/tmp/report.docx",
  "size": "a4",
  "orientation": "landscape",
  "top_mm": 20,
  "bottom_mm": 20,
  "left_mm": 25,
  "right_mm": 25
}}
```

### Excel: insert rows, freeze header, set column width

```json
{"tool": "xlsx_insert_rows", "args": {"path": "/tmp/data.xlsx", "sheet": "Sales", "row": 5, "count": 3}}
{"tool": "xlsx_freeze_panes", "args": {"path": "/tmp/data.xlsx", "sheet": "Sales", "cell": "A2"}}
{"tool": "xlsx_set_column_width", "args": {"path": "/tmp/data.xlsx", "sheet": "Sales", "column": "A", "width": 25}}
{"tool": "xlsx_merge_cells", "args": {"path": "/tmp/data.xlsx", "sheet": "Sales", "range_str": "A1:D1"}}
```

### Use edit history

After any mutating call the change is logged automatically. The AI can introspect:

```json
{"tool": "doc_history", "args": {"path": "/tmp/report.docx"}}
{"tool": "doc_last_edit", "args": {"path": "/tmp/report.docx"}}
{"tool": "doc_what_was_removed", "args": {"path": "/tmp/report.docx"}}
{"tool": "doc_revert", "args": {"path": "/tmp/report.docx", "revision": 2}}
```

### Add a chart to a spreadsheet

```json
{"tool": "xlsx_add_chart", "args": {
  "path": "/tmp/sales.xlsx",
  "sheet": "Sales",
  "chart_type": "bar",
  "data_range": "B1:C5",
  "categories_range": "A2:A5",
  "anchor_cell": "E2",
  "title": "Quarterly Revenue"
}}
```

### Spell-check + auto-correct

```json
{"tool": "doc_spell_check", "args": {"path": "/tmp/report.docx"}}
{"tool": "doc_apply_corrections", "args": {
  "path": "/tmp/report.docx",
  "corrections": {"recieve": "receive", "occured": "occurred"}
}}
```

### Page background + watermark + background image

```json
{"tool": "docx_set_background", "args": {"path": "/tmp/report.docx", "color_hex": "#FFFFE6"}}
{"tool": "docx_set_watermark", "args": {"path": "/tmp/report.docx", "text": "CONFIDENTIAL"}}
{"tool": "docx_set_background_image", "args": {
  "path": "/tmp/report.docx",
  "image_path": "/tmp/letterhead.png",
  "opacity": 20,
  "offset_y_mm": 10,
  "width_mm": 210,
  "height_mm": 40
}}
```

### Preview document pages

```json
{"tool": "doc_preview", "args": {
  "path": "/tmp/report.docx",
  "pages": "1-3",
  "dpi": 150
}}
```

Returns paths to PNG images the AI can view to verify formatting, backgrounds, and layout.

### Convert formats

```json
{"tool": "convert", "args": {"input_path": "/tmp/report.docx", "output_path": "/tmp/report.pdf"}}
```

With LibreOffice installed locally this works out of the box. With ONLYOFFICE Document Builder installed, the same call uses Document Builder's wider format set.

## Honesty about DOCX page counts

True page boundaries are NOT stored in `.docx` — Word/OnlyOffice compute them at layout time. `doc_page_count` returns:

- **Approximate** (default): count of explicit page breaks plus `<w:lastRenderedPageBreak/>` markers that real editors write when saving. Cheap; under-counts for files that have never been opened in a layout-aware editor.
- **Precise** (`precise=True`): converts to PDF via LibreOffice in a tempdir, counts pages via `pypdf`. ~5 seconds; exact.

For files in active editing only the approximate count is reliable.

## Roadmap

### v0.3.0 (stable alpha -> beta)

- [ ] Automated test suite with >80% coverage
- [ ] Read-back tools for headers/footers (`docx_read_header`, `docx_read_footer`)
- [ ] List existing bookmarks and hyperlinks (`docx_list_bookmarks`, `docx_list_hyperlinks`)
- [ ] Delete comments (`docx_delete_comment`)
- [ ] Edit existing slide text (`pptx_update_slide`)
- [ ] Reorder slides (`pptx_reorder_slides`)
- [ ] Cell styling (`xlsx_set_cell_style` — font, colour, borders, number format)
- [ ] Auto-filter on xlsx sheets (`xlsx_auto_filter`)
- [ ] Performance benchmarks for large documents

### v0.4.0

- [ ] Template system for common document types (invoices, reports, proposals)
- [ ] Batch operations (apply the same edit to multiple documents)
- [ ] Structured docx read (paragraphs as dicts with style/formatting metadata)
- [ ] ONLYOFFICE Docs API integration for live collaborative editing
- [ ] Streaming support for very large workbooks (>100k rows)

### v1.0.0

- [ ] Stable API — no breaking changes to tool names, parameters, or return schemas
- [ ] Full test suite with CI/CD pipeline
- [ ] Published to PyPI (`pip install onlyoffice-mcp`)
- [ ] Comprehensive documentation site
- [ ] Multi-language spell-check dictionaries

## Version history

| Version | Date | Tools | Highlights |
|---------|------|-------|------------|
| **0.3.0a5** | 2026-05-21 | 77 | In-place editing, error retry guidance, input sanitization, logged exceptions |
| **0.3.0a4** | 2026-05-21 | 75 | Doc preview, config detection, named colors, background image controls, text formatting, file logging |
| 0.3.0a3 | 2026-05-21 | 70 | File deletion safety, trash system, mass-deletion detection, audit logging |
| 0.3.0a2 | 2026-05-21 | 63 | Async threading, security hardening, prompt-injection defence |
| 0.3.0a1 | 2026-05-21 | 63 | LLM-friendly validation, 12 new tools, security fixes, bug fixes |
| 0.2.0a1 | 2026-05-21 | 51 | Edit history, cursor, spell-check, charts, styling, annotations, search |
| 0.1.0 | 2026-05-21 | 17 | Initial release: create/read/append for docx/xlsx/pptx + DocBuilder bridge |

See [CHANGELOG.md](CHANGELOG.md) for full details.

## Project layout

```
onlyoffice-mcp/
├── pyproject.toml
├── README.md
├── CHANGELOG.md
├── LICENSE
├── src/onlyoffice_mcp/
│   ├── __init__.py          # version
│   ├── __main__.py          # `python -m onlyoffice_mcp` entry point
│   ├── server.py            # FastMCP server + all @mcp.tool() definitions
│   ├── errors.py            # typed exceptions
│   ├── validation.py        # centralized LLM-friendly input validation
│   ├── storage.py           # workspace, doc_id, atomic writes, fcntl locks
│   ├── history.py           # edit log, snapshots, revert, record_operation decorator
│   ├── cursor.py            # AI virtual cursor + page-count estimation
│   ├── docx_ops.py          # Word operations (python-docx)
│   ├── xlsx_ops.py          # Excel operations (openpyxl)
│   ├── pptx_ops.py          # PowerPoint operations (python-pptx)
│   ├── docbuilder.py        # ONLYOFFICE Document Builder bridge
│   ├── libreoffice.py       # soffice --headless wrapper
│   ├── converter.py         # Format conversion engine chain
│   ├── charts.py            # matplotlib + openpyxl + python-pptx charts
│   ├── styling.py           # page bg, background image, watermark, slide bg, sheet tab
│   ├── preview.py           # document preview rendering (PyMuPDF)
│   ├── annotations.py       # headers/footers, hyperlinks, bookmarks, TOC, comments
│   ├── search.py            # find & replace with run preservation
│   ├── spell.py             # pyspellchecker + aspell fallback
│   └── stats.py             # polymorphic per-format stats
└── examples/
    ├── hello_world.docbuilder
    └── register_with_claude.sh
```

## License

MIT — see `LICENSE`.
