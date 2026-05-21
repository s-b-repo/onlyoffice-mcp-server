# OnlyOffice MCP Server

> **⚠️ ALPHA — `v0.2.0a1`**
> This project is in **alpha**. APIs, tool names, schemas and storage layout may change without notice between releases. Smoke-tested but not battle-tested. File issues and expect rough edges.

A pure-Python Model Context Protocol (MCP) server for creating, reading, editing and converting ONLYOFFICE-compatible Office documents — Word (`.docx`), Excel (`.xlsx`), and PowerPoint (`.pptx`).

ONLYOFFICE uses the same OOXML format as Microsoft Office, so files this server produces open natively in ONLYOFFICE Desktop / Docs, Microsoft Office, LibreOffice, Google Docs, and any other OOXML-compatible suite.

**v0.2 highlights:** built-in edit history with snapshots and revert, AI virtual cursor, spell-check + auto-correct, charts in all three formats, page backgrounds + watermarks, headers/footers/hyperlinks/bookmarks/TOC/comments, find & replace, plus a LibreOffice headless fallback for format conversion.

## Features

- **Word (docx)** — create, read, append, set metadata, headings, paragraphs, tables, images, lists, page breaks, alignment, fonts, colours
- **Excel (xlsx)** — create with multiple sheets, read, append rows, set individual cells, list sheets
- **PowerPoint (pptx)** — create with title / content / image / blank layouts, append slides, speaker notes
- **Version control** — every mutating call is journaled to `~/.onlyoffice-mcp/history/<doc_id>/`; query history, see what was removed, revert to any snapshot
- **AI virtual cursor** — track and move the "current" paragraph / cell / slide across MCP calls
- **Spell-check** — pyspellchecker primary engine with aspell subprocess fallback; word-level corrections preserve run formatting
- **Page background / watermark / sheet tab colour / slide background**
- **Headers / footers (with PAGE field) / hyperlinks (external + internal) / bookmarks / TOC field / Word-style comments**
- **Native charts** — bar / line / pie / scatter / area; openpyxl in xlsx, python-pptx in pptx, matplotlib-rendered PNG in docx
- **Find & replace** — run-formatting-preserving substitution across all three formats; dry-run mode
- **Document stats** — polymorphic per format (word/char/page for docx, sheet/cell for xlsx, slide/shape for pptx)
- **Format conversion** — ONLYOFFICE Document Builder → LibreOffice headless → pure-Python fallbacks (in that preference order)
- **Diagnostics** — `server_info` reports version, engines available, supported formats, history disk usage

## Requirements

- Python ≥ 3.10
- Runtime: `mcp`, `python-docx`, `openpyxl`, `python-pptx`, `matplotlib`, `lxml`, `pyspellchecker` (installed automatically)
- Optional: `pypdf` (for precise docx page counting via PDF) — install with `pip install 'onlyoffice-mcp[pdf]'`
- Optional native binaries (auto-detected, no Python deps): `documentbuilder` (full conversion matrix), `soffice` / `libreoffice` (broad conversion matrix), `aspell` (spell-check fallback)

## Install

```bash
cd /home/kali/onlyoffice-mcp
pip install -e .
# or:  pipx install /home/kali/onlyoffice-mcp
```

For precise docx page counting (LibreOffice → PDF → pypdf):

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

## Tools (51 total)

### Word

| Tool | Purpose |
|------|---------|
| `docx_create` | Create a Word document from a list of structured blocks |
| `docx_read` | Return text content of a Word document |
| `docx_append` | Append blocks to an existing Word document |
| `docx_set_metadata` | Update core properties (title, author, subject, keywords, comments) |
| `docx_set_background` | Set page background colour (hex) |
| `docx_set_watermark` | Add a diagonal text watermark on every page |
| `docx_set_header` | Set header text + alignment |
| `docx_set_footer` | Set footer text + optional PAGE field for page numbers |
| `docx_add_hyperlink` | Append an external hyperlink to a paragraph |
| `docx_add_internal_link` | Hyperlink targeting an internal bookmark |
| `docx_add_bookmark` | Wrap a paragraph in a named bookmark |
| `docx_add_toc` | Insert a Table of Contents field |
| `docx_add_comment` | Attach a Word-style comment to a paragraph |
| `docx_list_comments` | Enumerate existing comments |
| `docx_add_chart` | Render a chart with matplotlib and embed as image |

### Excel

| Tool | Purpose |
|------|---------|
| `xlsx_create` | Create a workbook from `{sheet: rows}` |
| `xlsx_read` | Read one or all sheets as 2-D arrays |
| `xlsx_append_rows` | Append rows to a sheet |
| `xlsx_set_cell` | Set a single cell (e.g. `B2`) |
| `xlsx_list_sheets` | Return sheet names |
| `xlsx_set_sheet_tab_color` | Colour the workbook tab strip |
| `xlsx_add_chart` | Add a native chart referencing existing cells |

### PowerPoint

| Tool | Purpose |
|------|---------|
| `pptx_create` | Create a presentation from a list of slide dicts |
| `pptx_read` | Extract slide titles, text, and speaker notes |
| `pptx_add_slide` | Append a single slide |
| `pptx_set_slide_background` | Set a solid background colour on one slide or all |
| `pptx_add_hyperlink` | Set a hyperlink on the first run of a shape |
| `pptx_add_chart` | Add a native chart to a slide |

### Edit history / version control

| Tool | Purpose |
|------|---------|
| `doc_history` | List recent revisions |
| `doc_history_show` | Full record of one revision (including diff) |
| `doc_last_edit` | Most recent op + age in seconds |
| `doc_what_was_removed` | **Only the `-` lines from a revision's diff** |
| `doc_diff` | Unified text diff between two revisions |
| `doc_revert` | Restore from a snapshot |
| `doc_clear_history` | Wipe history for a doc |
| `doc_history_stats` | Total tracked docs + disk usage |

### Cursor / navigation

| Tool | Purpose |
|------|---------|
| `doc_get_cursor` | Current cursor + format-specific bounds + stale flag |
| `doc_set_cursor` | Move cursor; values are clamped to doc bounds |
| `doc_page_count` | docx: approx pages (precise=True uses LibreOffice + pypdf) |

### Spell-check / find & replace

| Tool | Purpose |
|------|---------|
| `doc_spell_check` | Find misspellings with suggestions (engine: pyspellchecker → aspell fallback) |
| `doc_apply_corrections` | Apply word-level substitutions, preserving formatting |
| `spell_suggest` | Single-word lookup, no document needed |
| `doc_find` | Find every occurrence of a pattern (regex optional) |
| `doc_replace` | Run-preserving replacement with dry-run support |

### Conversion / metadata

| Tool | Purpose |
|------|---------|
| `convert` | Convert between formats (DocBuilder → LibreOffice → Python fallback) |
| `docbuilder_status` | Whether ONLYOFFICE Document Builder is installed |
| `docbuilder_run` | Execute a Document Builder script |
| `chart_kinds` | List supported chart kinds + synonyms |
| `doc_stats` | Polymorphic per-format statistics |
| `list_workspace` | List Office documents in a directory |
| `server_info` | Diagnostics: version, engines, formats, history usage |

## Schemas

### Block schema for `docx_create` / `docx_append`

Each item in the `paragraphs` list is either a plain string (body paragraph) or a dict:

- `{"type": "paragraph", "text": "...", "style": "normal|heading1|...|quote|code|bullet", "align": "left|center|right|justify", "bold": true, "italic": true, "size": 12, "color": "#ff0000"}`
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

### Page background + watermark

```json
{"tool": "docx_set_background", "args": {"path": "/tmp/report.docx", "color_hex": "#FFFFE6"}}
{"tool": "docx_set_watermark", "args": {"path": "/tmp/report.docx", "text": "CONFIDENTIAL"}}
```

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

## Project layout

```
onlyoffice-mcp/
├── pyproject.toml
├── README.md
├── LICENSE
├── src/onlyoffice_mcp/
│   ├── __init__.py          # version
│   ├── __main__.py          # `python -m onlyoffice_mcp` entry point
│   ├── server.py            # FastMCP server + all @mcp.tool() definitions
│   ├── errors.py            # typed exceptions
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
│   ├── styling.py           # page bg, watermark, slide bg, sheet tab
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
