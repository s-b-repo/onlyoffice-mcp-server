# Changelog

All notable changes to this project will be documented in this file.

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/),
and this project follows [PEP 440](https://peps.python.org/pep-0440/) for
version numbers (`MAJOR.MINOR.PATCH` plus pre-release tags like `a1`, `b1`,
`rc1`).

## [0.2.0a1] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** APIs, tool names, schemas, and the on-disk
history layout may change between alpha builds without backward-compatibility
guarantees.

### Added

- **Edit history & version control** (8 tools): every mutating tool is
  journaled to `~/.onlyoffice-mcp/history/<doc_id>/` with a JSONL op log
  and last-20 binary snapshots. New tools: `doc_history`,
  `doc_history_show`, `doc_last_edit`, `doc_what_was_removed`, `doc_diff`,
  `doc_revert`, `doc_clear_history`, `doc_history_stats`.
- **AI virtual cursor** (3 tools): `doc_get_cursor`, `doc_set_cursor`,
  `doc_page_count` (approx + optional precise via LibreOffice → pypdf).
- **Spell-check / auto-correct** (3 tools): primary engine is
  `pyspellchecker`; falls back to `aspell` subprocess when pyspellchecker
  is not loadable. New tools: `doc_spell_check`, `doc_apply_corrections`,
  `spell_suggest`.
- **Page background & styling** (4 tools): `docx_set_background`,
  `docx_set_watermark`, `pptx_set_slide_background`,
  `xlsx_set_sheet_tab_color`.
- **Headers / footers / hyperlinks / bookmarks / TOC / comments**
  (9 tools): `docx_set_header`, `docx_set_footer`, `docx_add_hyperlink`,
  `docx_add_internal_link`, `docx_add_bookmark`, `docx_add_toc`,
  `docx_add_comment`, `docx_list_comments`, `pptx_add_hyperlink`.
- **Charts** (4 tools): native `xlsx_add_chart` via openpyxl, native
  `pptx_add_chart` via python-pptx, matplotlib-rendered `docx_add_chart`
  (as a static image), and a `chart_kinds` meta tool.
- **Find & replace** (2 tools, polymorphic across formats): `doc_find`,
  `doc_replace` — both preserve run-level formatting where possible.
- **Document statistics** (1 polymorphic tool): `doc_stats`.
- **LibreOffice fallback** in `converter.py`: when ONLYOFFICE Document
  Builder is not installed, `soffice --headless` is tried before falling
  back to pure-Python. Uses `-env:UserInstallation` to avoid clashing
  with the user's open LibreOffice instance.
- **Diagnostics**: `server_info` now reports Document Builder + LibreOffice
  install state, supported formats per engine, history disk usage, and
  workspace path.

### Changed

- `__version__` bumped to `0.2.0a1`.
- `pyproject.toml` classifier changed to `Development Status :: 3 - Alpha`.
- `converter.py` now uses a three-tier engine chain:
  Document Builder → LibreOffice → Python fallbacks.
- Mutating tools are wrapped with the new `record_operation` decorator;
  existing v0.1 tool signatures are unchanged.

### Dependencies

- Added: `matplotlib>=3.7`, `lxml>=4.9`, `pyspellchecker>=0.7.2`.
- Optional: `pypdf>=4.0` (used only by `doc_page_count(precise=True)`).
- Optional native binaries (auto-detected): `documentbuilder`, `soffice`
  / `libreoffice`, `aspell`.

### Project layout

- Added 11 new modules: `errors.py`, `storage.py`, `history.py`,
  `cursor.py`, `spell.py`, `charts.py`, `styling.py`, `search.py`,
  `stats.py`, `annotations.py`, `libreoffice.py`.
- Total registered MCP tools: **51** (up from 17 in v0.1).

### Known limitations / sharp edges

- **DOCX page count is approximate** by default — actual page boundaries
  are computed at layout time by a real editor. Use `precise=True` (and
  install `pypdf`) for an exact count via PDF round-trip.
- **DOCX page background** sets `<w:background>` + `<w:displayBackgroundShape/>`
  but Google Docs ignores the colour entirely.
- **DOCX watermarks** use VML (the same approach Word uses). Word renders
  fully; LibreOffice partially; Google Docs ignores VML.
- **`docx_add_chart`** embeds a static matplotlib image — the chart is
  NOT editable inside Word/OnlyOffice. Use `xlsx_add_chart` for editable
  charts.
- **`docx_add_comment`** builds (or extends) `word/comments.xml` and adds
  range markers; on docs that don't already contain a comments part the
  package-injection path may degrade to a bracketed-text fallback. Test
  output by opening in LibreOffice before relying on it.
- **History disk usage** is capped at 20 snapshots × 100 MB total; older
  snapshots are pruned but the JSONL op log is uncapped. Call
  `doc_clear_history` to reclaim space manually.
- **Concurrent edits** to the same document from multiple MCP calls are
  serialised via `fcntl.flock`; contended calls retry briefly then raise
  `DocumentLocked`.

## [0.1.0] — 2026-05-21

Initial release. 17 tools covering docx/xlsx/pptx create, read, append
plus an ONLYOFFICE Document Builder bridge. See git history for the
exact tool set.
