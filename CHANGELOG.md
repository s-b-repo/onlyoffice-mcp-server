# Changelog

All notable changes to this project will be documented in this file.

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/),
and this project follows [PEP 440](https://peps.python.org/pep-0440/) for
version numbers (`MAJOR.MINOR.PATCH` plus pre-release tags like `a1`, `b1`,
`rc1`).

## [0.3.0a8] — 2026-05-28 — ⚠️ ALPHA

**This is an alpha release.** Rich per-cell / per-slide formatting, image
extraction, logo & branding overlays, colour-aware headers/footers, a document
preview that no longer requires PyMuPDF, and a fail-fast input-validation pass
that rejects malformed AI input with actionable errors.

### Added

- **7 new tools** (total now **92**):
  - `docx_format_cell` — format any table cell: optionally replace its text,
    then style the runs (bold/italic/underline, hex colour, size, font), set
    horizontal `align` + `vertical_align`, and cell fill `shading`. Enables
    white-on-dark cells directly — the create/append block API only bolds the
    header row.
  - `docx_extract_images` — extract every embedded image from a `.docx` to a
    directory (defaults to `<docname>_images`); returns `{count, directory, files}`.
  - `xlsx_format_cells` — apply `number_format`, font (bold/italic/colour/size/
    name), `fill_color`, alignment, `wrap_text` and borders over a cell or range
    (`"A1"` / `"A1:C5"`), preserving existing formatting where a parameter is omitted.
  - `pptx_add_textbox` — add a free-floating, positioned text box (inches) with
    font/size/colour/bold/alignment; newlines become separate paragraphs.
  - `pptx_set_speaker_notes` — set (replace) a slide's speaker notes.
  - `docx_place_image` — place a floating image (logo/crest) at a fixed page
    position on every page, anchored via section headers. Position is given in
    mm from the page top-left (`offset_x_mm`/`offset_y_mm`); `width_mm` is
    required and `height_mm` defaults to preserve the image's aspect ratio. Sits
    **in front of** text by default (`behind=true` to place it behind, like a
    background). Covers the default, first-page (`titlePg`) and even-page
    headers, so it appears on every page including a separate cover. Re-running
    with the same `name` replaces that overlay; a different `name` stacks.
    `opacity` 1–100.
  - `graphic_key_logo` — key a logo's flat background to transparency while
    **keeping the artwork's original colours and interior detail** (flood-fill
    inward from the image edges, so light areas enclosed by darker ink — e.g. a
    white roundel inside a coloured crest — survive rather than being punched
    out). Complements `graphic_recolor_image`, which flattens the mark to a
    single flat colour. Controls: `thresh` (flood tolerance), `feather` (soft
    edge), `crop`/`pad`, `scale` (1–6×).

### Changed

- **`docx_set_header` / `docx_set_footer` gained `color` + `size`** — the colour
  is applied to the text AND the PAGE-number field, so headers/footers stay
  visible over a dark page background (previously required a manual XML edit).
- **`doc_preview` no longer requires PyMuPDF** — it now renders via poppler's
  `pdftoppm` / `pdfinfo` when `fitz` (PyMuPDF) is unavailable, falling back
  automatically. The result dict gains an `engine` field (`"fitz"` |
  `"pdftoppm"`). This fixes `doc_preview` failing with `No module named 'fitz'`
  in environments without PyMuPDF installed.
- **Fail-fast input validation (no error swallowing)** — new `validate_choice()`
  and `validate_records()` helpers in `validation.py`. All new tools **and** the
  core `_add_block` content path (`docx_create` / `docx_append` /
  `docx_insert_paragraph`) now reject malformed/incorrect AI input with the
  standard two-line error instead of crashing or silently producing a broken
  document:
  - list-of-dict args (chart segments, bubble cards, infographic nodes) are
    validated element-by-element for required and numeric keys, pointing at the
    offending index;
  - enums (alignment, vertical-align, Excel border style, panel motif) are
    checked against an explicit set;
  - ranges/types are bounded (`cols`, image dimensions, `dpi`, font size, points);
  - a `table` block's `data` must be a list of row-lists and a `list` block's
    `items` must be a list (both previously yielded **silent garbage**); an
    `image` block now requires a `path` (previously a bare `KeyError`); and
    `cell_shading` / `cell_alignment` keys are validated as in-range `"row,col"`.

## [0.3.0a7] — 2026-05-28 — ⚠️ ALPHA

**This is an alpha release.** Slide-style graphic generation for dark-themed
"deck" reports — standalone PNG renderers (charts, cards, infographics,
backgrounds) to composite as a document background or embed as page images.

### Added

- **8 new graphic tools** (new `graphics.py` module) — all render on a
  transparent canvas with light text so they sit on a dark themed page, and
  supersample for smooth anti-aliased edges:
  - `graphic_tech_background` — dark gradient + hexagon grid + corner glow +
    dotted wave + corner vignette, with an optional branded header band
    (monogram/text on one side, logo on the other)
  - `graphic_recolor_image` — recolour a logo/mark to a single flat colour on
    transparency, tight-cropped and optionally upscaled
  - `graphic_donut_chart` / `graphic_bar_chart` — ring and bar charts on
    transparency with stroked light labels and value annotations
  - `graphic_bubble_cards` — rounded "bubble" card grid (circular ID badge +
    title + impact + severity pill, pill text auto-darkens on light colours)
  - `graphic_node_infographic` — glossy circular value nodes on a dashed zigzag
    connector
  - `graphic_numbered_cards` — auto-numbered recommendation/step cards in columns
  - `graphic_decorative_panel` — abstract hexagon-cluster panel with an optional
    centre motif (`lock` / `shield` / `none`)

## [0.3.0a6] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** Professional document formatting: paragraph
indentation/spacing, enhanced headings with formatting, multi-level nested
lists, advanced chart customization, and table cell styling.

### Added

- **Paragraph indentation & spacing** — `left_indent`, `right_indent`,
  `first_line_indent`, `space_before`, `space_after`, `line_spacing`,
  `keep_with_next` on paragraph, heading, and list blocks
- **Enhanced headings** — `alignment`, `bold`, `italic`, `underline`,
  `font_name`, `font_size`, `font_color`, `numbering_prefix` on heading blocks
- **Multi-level nested lists** — items as dicts with `text`, `level` (0-5),
  per-item formatting (`bold`, `italic`, `font_color`, `font_size`),
  custom `bullet_char` for bullet lists, styles up to "List Bullet 3" /
  "List Number 3"
- **Advanced docx charts** — `xlabel`, `ylabel`, `data_labels`, `colors`,
  `legend_position`, `stacked`, `horizontal`, `explode`, `donut`,
  `line_styles`, `grid`, `dpi` parameters on `docx_add_chart`
- **Enhanced xlsx charts** — `xlabel`, `ylabel`, `stacked`, `legend_position`,
  `data_labels` on `xlsx_add_chart`
- **Enhanced pptx charts** — `stacked`, `data_labels`, `legend_position`
  on `pptx_add_chart`
- **Table cell styling** — `col_widths`, `header_shading`, `cell_shading`
  (by "row,col" key), `cell_alignment` on table blocks
- **New validators** — `validate_line_spacing`, `validate_legend_position`,
  `validate_list_level`, `validate_indent` in validation.py
- **New list styles** in STYLE_MAP — `bullet2`, `bullet3`, `numbered2`,
  `numbered3`

## [0.3.0a5] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** Safe error handling with AI retry guidance,
input sanitization across all text paths, and in-place editing for documents
and presentations.

### Added

- **2 new tools** (total now **77**):
  - `docx_edit_paragraph` — edit existing paragraph text and formatting
    in place (text, style, alignment, bold, italic, underline,
    strikethrough, font, size, color). Returns paragraph state after edit.
  - `pptx_update_slide` — edit existing slide title, body text, or
    speaker notes in place. Returns slide state after edit.

### Changed

- **`_threaded` wrapper** now catches `ValueError`, `FileNotFoundError`,
  `PermissionError`, `RuntimeError`, and `OSError` — appends AI-friendly
  retry guidance to every error message so models know what to try next.
- **Input sanitization** via `sanitize_text()` wired into all text paths:
  - `docx_ops.py` — paragraph text, headings, table cells, list items
  - `pptx_ops.py` — slide titles, subtitles, body bullets, speaker notes
  - `xlsx_ops.py` — cell values, sheet names, appended row data
  - `annotations.py` — header/footer text, hyperlink text, bookmark names,
    comment author/initials/text
- **Bare except blocks** replaced with logged exceptions across 9 files:
  - `cursor.py` — 5 instances (docx/xlsx/pptx bounds, auto-advance)
  - `stats.py` — XML stats extraction
  - `styling.py` — settings part lookup
  - `safety.py` — zip-bomb check fallthrough
  - `history.py` — corrupt meta.json in history_stats

## [0.3.0a4] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** Document preview, enhanced background image
controls, extended text formatting, and persistent file-based logging.

### Added

- **5 new tools** (total now **75**):
  - `doc_preview` — render document pages as PNG images for AI visual
    inspection. Converts docx/xlsx/pptx/pdf/odt/ods/odp to PDF via
    LibreOffice, then renders pages at configurable DPI via PyMuPDF.
    Supports page ranges (`pages="1-3,5"`), DPI control (72/150/300),
    and batch limits (`max_pages`). Temp images auto-cleaned after 30 min.
  - `docx_get_config` — detect full document configuration: page setup
    (size, margins, orientation), sections with header/footer content,
    fonts and colors in use, paragraph styles, background/watermark
    status, table count, and metadata. Use before modifying a document.
  - `docx_get_formatting` — inspect per-paragraph formatting: style,
    alignment, and per-run details (bold, italic, underline,
    strikethrough, font name, font size, text color).
  - `color_info` — list all 59 supported CSS named colors with hex
    values and accepted input formats.
  - `docx_set_background_image` — **enhanced** with 5 new parameters:
    `offset_x_mm`, `offset_y_mm` (position from top-left), `width_mm`,
    `height_mm` (image dimensions), and `opacity` (1-100, with
    `a:alphaModFix` OOXML element). Replaces previous 2-parameter version.
- **Text formatting extensions** in `docx_create` / `docx_append` /
  `docx_insert_paragraph` paragraph blocks:
  - `underline: true` — underline text
  - `strikethrough: true` — strikethrough text
  - `font: "Arial"` — set font family name
  (Previously supported: `bold`, `italic`, `size`, `color`, `align`)
- **Named color support** in `validate_color()` — accepts 59 CSS color
  names (red, navy, steelblue, coral, gold, etc.) in addition to hex.
  All color parameters across all tools now accept named colors.
- **Persistent file-based logging** — server logs are written to
  `~/.onlyoffice-mcp/logs/server.log` via `RotatingFileHandler` (5 MB
  max, 3 backups). Errors, warnings, and info are persisted across
  sessions for debugging.
- **`preview.py` module** — document preview rendering pipeline:
  - `doc_preview()` — main entry point returning page image paths,
    dimensions, total page count, and rendering metadata.
  - `_convert_to_pdf()` — LibreOffice headless conversion.
  - `_parse_page_range()` — parses comma-separated page ranges like
    `"1-3,5,8-10"` to zero-based indices.
  - Auto-cleanup of stale preview images (30-minute TTL).

### Changed

- `docx_set_background_image` docstring now includes opacity guidance
  (15-30 for branded letterhead, 10-20 for decorative backgrounds) and
  recommends using `doc_preview` to verify results visually.
- `_build_bg_anchor_xml()` in `styling.py` — dynamic XML builder
  replaces static template; generates `a:alphaModFix` element when
  opacity < 100%.
- Background images are idempotent — `_remove_bg_images_from_header()`
  strips previous `PageBackground` anchors before inserting new ones.
- Watermarks are idempotent — `_remove_watermarks_from_header()` strips
  previous `WatermarkShape` VML before inserting new ones.
- Watermark `font_size` now validates as `int` with range [1, 999].
- Watermark text capped at 200 characters.
- `server.py` `main()` now configures `RotatingFileHandler` alongside
  stderr logging.

### Dependencies

- Optional: `PyMuPDF>=1.24` (for `doc_preview` rendering) — install
  with `pip install 'onlyoffice-mcp[preview]'`.
- Optional: LibreOffice (`soffice`) required for non-PDF preview
  conversion.

### Project layout

- Added `preview.py` — document preview rendering via PyMuPDF.
- Total registered MCP tools: **75** (up from 70 in v0.3.0a3).

---

## [0.3.0a3] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** File deletion safety, trash system, mass
deletion detection, and additional guard rails. Inspired by analysis of
the ONLYOFFICE/docspace-mcp project.

### Added

- **7 new tools** (total now **70**):
  - `doc_delete_file` — permanently delete a single document file with
    safety checks, rate limiting, and audit logging.
  - `doc_delete_files` — batch delete with mass-deletion detection and
    per-directory wipe warnings.
  - `doc_move_to_trash` — recoverable soft-delete (file moves to
    `~/.onlyoffice-mcp/trash/` with metadata for restore).
  - `doc_list_trash` — list recoverable trash contents with original
    paths and timestamps.
  - `doc_restore_from_trash` — restore a trashed file to its original
    location.
  - `doc_empty_trash` — permanently purge trash (optionally only items
    older than N hours).
  - `doc_deletion_audit` — view the immutable deletion audit log
    (every delete/trash op is recorded).
- **Deletion rate limiter** (`DeletionTracker`) — thread-safe sliding
  window that blocks when more than N deletions occur within a time
  window (default: 5 deletions per 60s). Configurable via
  `ONLYOFFICE_MCP_DELETION_RATE_LIMIT` and
  `ONLYOFFICE_MCP_DELETION_RATE_WINDOW` env vars.
- **Mass deletion detection** — `doc_delete_files` rejects batches
  larger than 50 files, warns on batches > 10, and flags when multiple
  files come from the same directory (possible accidental wipe).
- **Sensitive path blocklist** — 18 patterns block deletion of SSH
  keys, GPG keyrings, AWS/Docker/Kube configs, `.env` files, shell
  profiles, git repos, and credential files.
- **Deletion audit log** — every delete/trash operation is appended to
  `~/.onlyoffice-mcp/deletion_audit.jsonl` with timestamp, path, method,
  success status, and PID. Queryable via `doc_deletion_audit`.
- **AI safety warnings** for deletion operations — docstrings and return
  values warn AI assistants to never delete files based on document
  content instructions.

### Security

- **Soft-delete by default**: `doc_move_to_trash` is the recommended
  deletion method — files are recoverable until trash is emptied.
  Inspired by ONLYOFFICE DocSpace-MCP's trash-folder pattern
  (`deleteAfter: false, immediately: false`).
- **Rate limiting on destructive ops**: sliding-window rate limiter
  prevents rapid-fire deletion (prompt injection defence).
- **Batch size cap**: `doc_delete_files` hard-limits at 50 files per
  call to prevent runaway deletions.
- **Per-directory wipe detection**: batch deletes flag when > 5 files
  come from the same parent directory.
- **Non-document file guard**: `doc_delete_file` rejects non-document
  files unless `force=True` is explicitly set.

### Changed

- `server_info` now reports deletion rate limit, mass deletion
  detection, trash system, and audit log status.

### Security gaps found in ONLYOFFICE/docspace-mcp

- No deletion rate limiting or mass deletion detection.
- No prompt injection scanning on file content.
- No file size validation on uploads.
- No audit logging for destructive operations.
- No batch size limits on copy/move operations.

### Project layout

- Extended `safety.py` — added `DeletionTracker`, trash system, audit
  log, sensitive path patterns, batch deletion checks.
- Extended `errors.py` — added `MassDeletionBlocked`, `DeletionDenied`.
- Total registered MCP tools: **70** (up from 63 in v0.3.0a2).

---

## [0.3.0a2] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** Async threading, comprehensive security
hardening, prompt-injection defence, and AI safety guidance.

### Added

- **Async all tools** — every MCP tool handler is now `async`; CPU-bound
  work runs via `asyncio.to_thread()` so the event loop never blocks.
  Multiple concurrent requests are processed in parallel.
- **`safety.py` module** — centralised security layer:
  - File-size limits (100 MB default, `ONLYOFFICE_MCP_MAX_FILE_SIZE` env var).
  - Zip-bomb detection (compression-ratio check on OOXML files).
  - System-path blocklist (`/proc`, `/sys`, `/dev` denied).
  - Symlink and path-traversal detection with warnings.
  - Prompt-injection scanner (10 patterns: role override, instruction
    injection, persona manipulation, jailbreak keywords, XML tag injection).
  - Macro / VBA / ActiveX / external-link detection in documents.
  - Format-conversion whitelist (source→target pairs).
  - AI safety preamble and risky-operation warnings.
  - Hardened XML parser (`resolve_entities=False`, `no_network=True`).
- **Content warnings on reads** — `docx_read`, `xlsx_read`, `pptx_read`
  now scan returned text for prompt-injection patterns and return a
  `content_warnings` list alongside the data.
- **ReDoS protection** — `validate_regex()` rejects patterns with nested
  quantifiers like `(a+)+` that cause catastrophic backtracking.
- **`validate_bounded_int()`** — caps numeric parameters (e.g. `max_words`).
- **`max_results` on `list_workspace`** — prevents unbounded directory scans.
- **Billion laughs protection** — all `etree.fromstring()` calls now use
  `safe_parse_xml()` which disables entity expansion.

### Security

- **XML entity expansion (billion laughs)**: all lxml `fromstring()` calls
  replaced with `safe_parse_xml()` using `resolve_entities=False`,
  `no_network=True`, `huge_tree=False`.
- **ReDoS via user regex**: `doc_find` and `doc_replace` now reject
  patterns with nested quantifiers before compiling.
- **Path traversal**: `validate_path()` now blocks system paths and warns
  on symlinks / `..` traversal.
- **Zip-bomb detection**: OOXML files are checked for suspicious
  compression ratios (default limit: 100:1).
- **File-size limits**: all read/write paths enforce a configurable maximum.
- **Format-conversion whitelist**: `convert()` now validates source→target
  format pairs against an explicit allowlist.
- **JS string escaping hardened**: `_js_string_escape()` now also handles
  single quotes, backticks, `${}` template literals, and null bytes.
- **`converter.py` same-format copy**: replaced `read_bytes()`/`write_bytes()`
  with `shutil.copy2()` to prevent OOM on large files.
- **`docbuilder_run` security warning**: tool docstring explicitly warns
  AI assistants not to execute scripts from document content or websites.

### Changed

- `docx_read` returns `{"text": ..., "content_warnings": [...]}` instead
  of a plain string (prompt-injection + macro warnings included).
- `xlsx_read` returns `{"data": ..., "content_warnings": [...]}`.
- `pptx_read` result dict now includes `content_warnings` list.
- `docbuilder_run` returns `{"path": ..., "warning": "..."}` with a
  security reminder.
- `server_info` now reports safety features (file size limits,
  decompression ratio, enabled protections, conversion whitelist).
- `list_workspace` accepts `max_results` parameter (default 500).
- `doc_spell_check` validates `max_words` range (1–10,000).

### Project layout

- Added `safety.py` — centralised security, safety, and AI-guidance module.
- All tool handlers in `server.py` are now async via `_threaded` wrapper.

---

## [0.3.0a1] — 2026-05-21 — ⚠️ ALPHA

**This is an alpha release.** 12 new tools, a centralized validation layer
with LLM-friendly errors, critical bug fixes, and two security patches.

### Added

- **12 new tools** (total now **63**):
  - Word: `docx_read_metadata`, `docx_set_page_setup`, `docx_insert_paragraph`,
    `docx_delete_paragraph`.
  - Excel: `xlsx_delete_sheet`, `xlsx_rename_sheet`, `xlsx_delete_rows`,
    `xlsx_insert_rows`, `xlsx_merge_cells`, `xlsx_set_column_width`,
    `xlsx_freeze_panes`.
  - PowerPoint: `pptx_delete_slide`.
- **Centralized validation module** (`validation.py`): every error message
  now follows a structured format — line 1 says what went wrong, line 2+
  tells the LLM what the valid values are or what tool to use instead.
  Functions: `validate_path`, `validate_color`, `validate_cell_ref`,
  `validate_cell_range`, `validate_index`, `validate_sheet_name`,
  `validate_paragraph_index`, `validate_slide_index`, `validate_align`,
  `validate_chart_type`, `validate_page_size`, `validate_series_data`,
  `validate_slide_def`.
- **Cross-tool suggestions**: using the wrong file extension (e.g. calling
  `docx_read` on a `.xlsx`) now says which tool to use instead.
- **Named page sizes**: `docx_set_page_setup` accepts `letter`, `a4`, `a3`,
  `a5`, `legal`, `tabloid` plus custom `width_mm` / `height_mm`.

### Fixed

- **`docx_add_chart` crash** when `paragraph_index` was provided:
  `run.add_picture()` does not exist on `Run` objects — rewritten to use
  body element repositioning.
- **`doc_revert` before_hash bug**: the "before" hash was captured *after*
  the file was overwritten, making `before_hash == after_hash` always.
- **`doc_apply_corrections` accumulation bug**: `applied = count` on the
  multi-run fallback reset the running total instead of accumulating.
- **`spell_suggest` is_known logic**: was inverted — a known word showed
  `is_known: false`. Now uses `engine_obj.known()` directly.
- **`docx_add_internal_link`**: missing `paragraph_index` bounds check
  caused bare `IndexError`.
- **`pptx_add_hyperlink`**: missing bounds checks on `slide_index` and
  `shape_index`; also saved inside the nested loop instead of after.
- **`docx_add_bookmark` ID collisions**: `abs(hash(name)) % 99999` could
  collide; now uses incremental IDs from existing bookmarks.
- **`pptx_read` indexing**: used 1-based slide indexing while all other
  pptx tools use 0-based. Normalized to 0-based.
- **`doc_history_show`**: error on invalid revision now lists available
  revisions instead of bare "No such revision".
- **Search/stats unsupported format**: `find_in_document` silently returned
  `[]` and `stats` gave a bare "Unsupported format" error. Both now
  list supported formats.
- **Formula counting**: simplified the `elif` double-branch to a single
  `or` condition.
- **`datetime.utcnow()` deprecation** in `annotations.py`: replaced with
  `datetime.now(timezone.utc)`.

### Security

- **Script injection in Document Builder**: paths interpolated into
  JavaScript strings were not escaped. A path containing `"` could break
  out of the string literal and inject arbitrary script code. Fixed with
  `_js_string_escape()`.
- **XML injection in DOCX watermark**: watermark text was interpolated
  directly into VML XML via `.format()`. Fixed with
  `xml.sax.saxutils.escape()`.

### Changed

- All tools in `annotations.py` (10 functions) and `styling.py` (4
  functions) now use `validate_path()` — every `FileNotFoundError` includes
  the suggestion to create the file first with the correct tool.
- `storage.py`: renamed `ENV_DISABLE_HISTORY` constant to
  `ENV_HISTORY_ENABLED` to match the actual environment variable name.
- `spell_suggest`: renamed parameter `max` to `max_suggestions` to avoid
  shadowing the Python builtin.
- Improved docstrings on 15+ MCP tools: added return schemas, parameter
  descriptions, valid value lists, and usage guidance.
- `docx_create` paragraph blocks: color field now validated through
  `validate_color()` instead of silently skipping invalid hex.
- `charts.py`: chart type validation now delegates to `validate_chart_type`
  with synonym hints. Series data and xlsx data ranges are validated.
- `validation.py`: fixed stale reference to non-existent `docx_set_margins`
  tool (now correctly says `docx_set_page_setup`).
- Removed unused `PatternFill` import from `xlsx_ops.py`.

### Project layout

- Added `validation.py` — centralized input validation with LLM-friendly
  error messages.
- Total registered MCP tools: **63** (up from 51 in v0.2).

---

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
