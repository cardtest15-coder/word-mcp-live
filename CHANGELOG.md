# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2026-05-22

### Added
- **71 new tools** (124 → 195) across Phase 1 and Phase 2 feature expansions
- **Style tools** (6 file-based): `list_styles`, `get_style_details`, `apply_style`, `modify_style`, `delete_style`, `copy_styles_from_template`
- **Column tools** (4 file-based): `set_section_columns`, `insert_column_break`, `get_section_layout`, `set_column_widths`
- **Advanced table tools** (5 file-based): `delete_table`, `repeat_table_header`, `convert_table_to_text`, `convert_text_to_table`, `sort_table`
- **Object insertion tools** (5 file-based): `update_table_of_contents`, `insert_text_box`, `insert_chart`, `insert_index`, `insert_citation`
- **Export tools** (3 file-based + 3 live): `export_to_html`, `export_to_rtf`, `export_to_txt`, `word_live_export_to_html`, `word_live_export_to_rtf`, `word_live_export_to_txt`
- **Merge tools** (2 file-based + 2 live): `compare_documents`, `mail_merge`, `word_live_compare_documents`, `word_live_mail_merge`
- **Highlight tools** (2 file-based + 1 live): `highlight_text`, `remove_highlight`, `word_live_highlight_text`
- **Field tools** (3 file-based + 3 live): `insert_field`, `update_fields`, `insert_content_control`, `word_live_insert_field`, `word_live_update_fields`, `word_live_insert_content_control`
- **Property tools** (2 file-based + 2 live): `get_custom_properties`, `set_custom_property`, `word_live_get_custom_properties`, `word_live_set_custom_property`
- **Page design tools** (6 file-based + 4 live): `set_different_first_page`, `set_odd_even_headers`, `set_tab_stops`, `clear_tab_stops`, `insert_drop_cap`, `set_page_borders`, `word_live_set_different_first_page`, `word_live_set_odd_even_headers`, `word_live_insert_drop_cap`, `word_live_set_page_borders`
- **Extra live tools** (8): `word_live_spell_check`, `word_live_check_grammar`, `word_live_insert_table_of_figures`, `word_live_insert_shape`, `word_live_run_macro`, `word_live_print`, `word_live_get_doc_variable`, `word_live_set_doc_variable`
- **13 live style/column/table/object tools**: `word_live_list_styles`, `word_live_apply_style`, `word_live_modify_style`, `word_live_set_section_columns`, `word_live_insert_column_break`, `word_live_get_section_layout`, `word_live_delete_table`, `word_live_repeat_table_header`, `word_live_sort_table`, `word_live_update_table_of_contents`, `word_live_insert_text_box`, `word_live_insert_chart`, `word_live_insert_index`
- `comtypes_word_app` context manager in `word_com.py` — eliminates duplicate GetActiveObject/CreateObject/Quit boilerplate
- `_resolve_style()` with `_WD_STYLE_MAP` — locale-aware style name resolution for COM (English → NameLocal)
- `get_word_app()` now auto-starts Word if not running; `find_document()` auto-opens from disk
- 91 unit tests, 46/46 live tools verified via MCP integration test

### Fixed
- **8 COM bugs found and fixed through systematic live testing**:
  1. `get_word_app()` — now starts Word if not running (was: RuntimeError)
  2. `find_document()` — now opens file from disk if not open (was: "No documents open")
  3. `word_live_sort_table` — `tbl.Sort(Column=)` → `tbl.Range.Sort(SortFieldType=, SortOrder=, SortColumn=)`
  4. `word_live_insert_text_box` — `AddTextbox(Orientation=0,...)` → positional float args `AddTextbox(1, left, top, width, height)`
  5. `word_live_insert_content_control` — `PlaceholderText` readonly → `SetPlaceholderText()` method
  6. `word_live_set_page_borders` — `PageSetup.Borders` → `Section.Borders`
  7. `word_live_set_custom_property` — `CDP.Add()` broken in COM → zipfile fallback
  8. `word_live_spell_check` — `CheckSpelling()` blocks with modal dialog → `doc.SpellingErrors` read-only
- 198 parameter type mismatches (`str = None` → `Optional[str] = None`) for FastMCP Pydantic validation
- 5 JXA off-by-one errors, 1 copy-paste bug, 200+ error format fixes
- `WD_COLOR_INDEX` mapping — `CYAN`/`MAGENTA`/`NONE` don't exist; mapped to `TURQUOISE`/`PINK`/`AUTO`
- `WD_TAB_LEADER.LEADER_NONE` doesn't exist; replaced with `WD_TAB_LEADER.SPACES`
- `set_custom_property` — rewrote using zipfile instead of broken `docx.opc.part.Part` constructor
- Repository migrated to `cardtest15-coder/word-mcp-live` (author's own repo, not a fork)

### Changed
- Version bumped to 2.0.0
- All repo URLs updated from `ykarapazar` to `cardtest15-coder`
- `requirements.txt` aligned with `pyproject.toml` dependencies
- `smithery.yaml` command updated to `uv run --directory . word_mcp_server`
- `setup_mcp.py` server key: `word-document-server` → `word-mcp-live`
- `publish.yml` upgraded to trusted publisher (OIDC) via `pypa/gh-action-pypi-publish`
- `manifest.json` updated to v2.0.0, 195 tools, `uv run` command

## [1.6.0] - 2026-04-29

### Added
- **`word_live_set_core_properties`** — set Word document Title, Subject, Author, Keywords, Comments, Category, Manager, Company, Last Author via `Document.BuiltInDocumentProperties`. Wrapped in `undo_record` so a single Ctrl+Z reverts every property in one call. Equivalent to File > Info > Properties in the Word UI.
- New `word_document_server/utils/text_safety.py` — shared `reject_control_chars()` validator for Find/Replace/Insert text inputs.
- `scrub_orphans` parameter on `word_live_modify_table` `delete_table` operation (default `True`) — cleans stranded `\x07` cell-separator bytes the Word COM `Table.Delete()` occasionally leaves behind.

### Fixed
- **`word_live_replace_text` data-loss vector** — passing `\x07` (cell separator) as `find_text` previously matched across cell boundaries and could delete entire documents. Control bytes (U+0000–U+001F except `\t`, `\n`, `\r`) now rejected with a descriptive error before Find.Execute is reached.
- **`word_live_find_text`** — same control-byte protection applied to `search_text`.
- **`word_live_insert_text`** — same control-byte protection applied to `text` (prevents inserting orphan cell separators outside a real table).
- **`word_live_modify_table` `delete_table`** — leftover `\x07` separators after Word's native `Table.Delete()` now scrubbed by default (configurable via `scrub_orphans=False`).
- **`word_live_add_table`** — rejects `position` offset that falls inside an existing table's range or sits immediately after an orphan cell separator (would otherwise silently merge new content into existing/residual table structure).
- **`word_live_setup_heading_numbering`** — paragraphs that previously kept a custom template style (e.g. `Font Style30/31`) after a forced heading reassignment now (a) get explicit per-paragraph style assignment with try/except, (b) receive the same font/size/bold/color customizations as direct formatting so visual output matches even when the underlying style refuses to swap, (c) report any failed reassignments under a new `restyle_failures` field in the response.
- **`word_live_modify_table`** — re-reads `Tables.Count` per call and validates `table_index` with a helpful message ("table_index N out of range. Document has K table(s)…") instead of throwing "Document has no tables" when a stale index is passed after a prior delete.
- **`word_live_list_open`** — defensive per-document property access; one document in a broken COM proxy state no longer aborts the whole call. Each document entry now includes `index`, `track_revisions`, and per-property `errors` array.
- **`word_live_find_text`** — defensive `Range.Text` / `Range.Start` / `Range.End` / `Document.Name` access via internal `_safe_attr` helper. Transient COM marshalling failures after MCP reconnect now produce partial matches with `<unreadable>` placeholders and a `partial_errors` array, rather than aborting the call.

## [1.5.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — infinite loop when wildcard pattern matches zero-length strings (e.g., `*` alone); now skips forward on zero-length matches and enforces 50K replacement safety ceiling

## [1.5.0] - 2026-04-08

### Added
- **macOS live editing support** via JavaScript for Automation (JXA) — 33 of 41 `word_live_*` tools now work on macOS with Word for Mac
- New module `word_document_server/core/word_mac.py` — JXA bridge with 30+ functions for Word for Mac automation
- Platform auto-detection: same tool names and parameters on both Windows and macOS
- `pywin32` as conditional dependency (Windows only) in `pyproject.toml`

### Changed
- All `print()` calls in `main.py` redirected to stderr — fixes MCP stdio protocol corruption that prevented the server from loading in some clients
- All live tool functions now dispatch to macOS JXA implementations when `sys.platform == "darwin"`
- Updated tool count: 76 cross-platform + 41 Windows Live + 33 macOS Live

### Not Available on macOS
These 4 tools require Windows COM APIs with no AppleScript/JXA equivalent:
- `word_live_get_undo_history` — undo stack inspection not exposed in Word for Mac's scripting dictionary
- `word_live_reply_to_comment` — threaded comment replies not in AppleScript dictionary
- `word_live_resolve_comment` — comment Done property not in AppleScript dictionary
- `word_live_add_watermark` — requires VBA `Shapes.AddTextEffect` (VBA bridge killed by Apple sandboxing in Word 365)

## [1.4.1] - 2026-04-08

### Fixed
- `word_live_replace_text` — `^s` (non-breaking space) now converted to `\u00a0` in replacement text (#4)

## [1.4.0] - 2026-04-08

### Added
- `word_live_insert_paragraphs` — insert multiple paragraphs near a target (by text or index) in a single undo record
- `word_live_take_snapshot` — store paragraph baseline for efficient change detection
- `word_live_get_diff` — compare current document against snapshot, returns only changed paragraphs
- `word_live_snapshot_status` — check snapshot existence and age
- `word_live_modify_table` — new `set_row` and `set_range` operations for bulk cell updates

### Fixed
- `word_live_replace_text` — infinite loop when document has TrackRevisions enabled independently of `track_changes` parameter (#7)
- All destructive tools now unconditionally restore `doc.TrackRevisions` in `finally` block

### Credits
- Snapshot/diff tools, `insert_paragraphs`, and bulk table operations adapted from PR #5 by @FarhadGSRX

## [1.3.0] - 2026-02-28

### Added
- `word_live_modify_table` — table operations via COM: get info, set cell, add/delete rows/columns, merge cells, autofit, delete table
- `word_live_save` — save document in place or save-as (docx, pdf, rtf, txt)
- `word_live_toggle_track_changes` — toggle or explicitly set track changes mode on/off
- `word_live_insert_image` — insert image with sizing, alignment, wrapping, and optional border
- `word_live_insert_cross_reference` — insert live cross-references to headings, bookmarks, figures, tables, equations, footnotes, endnotes
- `word_live_list_cross_reference_items` — list available cross-reference targets with their indices
- `word_live_insert_equation` — insert mathematical equations using UnicodeMath syntax
- `word_live_reply_to_comment` — threaded comment replies (Word 2016+)
- `word_live_resolve_comment` — mark comments as resolved/unresolved (Word 2016+)
- `word_live_delete_comment` — permanently delete a comment
- Total tool count now **114** (75 cross-platform + 39 Windows Live)

### Changed
- `word_live_delete_text` — now table-aware: deletes table objects within range before text deletion
- `word_live_insert_text` — auto-chunks text >30K chars to avoid COM 32K limit
- `word_live_setup_heading_numbering` — handles inflated paragraph ranges from comment anchors
- `word_live_modify_table` set_cell operation now accepts tracked changes before writing to prevent layered content

## [1.2.0] - 2025-02-15

### Added
- `word_live_replace_text` — find & replace via COM that works across tracked change boundaries; supports wildcards (`^m`, `^t`, `^p`) and tracked changes mode
- `word_live_diagnose_layout` — read-only scan for layout problems: keep_with_next chains, heading styles on body text, PageBreakBefore misuse, manual breaks
- `word_live_get_paragraph_format` — inspect paragraph formatting (font, spacing, alignment, list info, style); `include_runs=True` for per-run detail
- `word_live_get_page_text` — read text from specific page(s) with char offsets for chaining into format/edit tools
- `word_live_get_undo_history` — list undo stack entries
- `word_live_apply_list` — apply bullet, numbered, or multilevel list formatting
- `word_live_setup_heading_numbering` — auto-numbered headings (1. / 1.1) via multilevel list linked to Heading styles; configurable style params (font, size, color, spacing)

### Changed
- `word_live_format_text` — added `paragraph_alignment`, `page_break_before`, paragraph-index addressing (`start_paragraph`/`end_paragraph`), `preserve_direct_formatting` for style changes
- `word_live_find_text` — added `use_wildcards` for `^m`/`^t`/`^p`/Word wildcard syntax; `context_chars` now configurable (default 60, was 30)
- `word_live_set_paragraph_spacing` — clarified that `line_spacing` is in points (1.15 lines = 13.8pt)

## [1.1.0] - 2025-01-10

### Added
- 27 Windows Live tools (`word_live_*`) using COM automation for editing documents open in Word
- Per-operation undo system — all destructive tools wrapped with `UndoRecord`; each tool call = one Ctrl+Z entry
- `word_live_undo` — programmatic undo of last N operations
- Live editing tools: `word_live_insert_text`, `word_live_delete_text`, `word_live_format_text`, `word_live_add_table`
- Live reading tools: `word_live_get_text`, `word_live_get_info`, `word_live_find_text`
- Live comment & revision tools: `word_live_add_comment`, `word_live_get_comments`, `word_live_list_revisions`, `word_live_accept_revisions`, `word_live_reject_revisions`
- Live layout tools: `word_live_set_page_layout`, `word_live_add_header_footer`, `word_live_add_page_numbers`, `word_live_add_section_break`, `word_live_set_paragraph_spacing`, `word_live_add_bookmark`, `word_live_add_watermark`
- `word_screen_capture` — screenshot of the Word window
- Cross-platform tracked changes: `track_replace`, `track_insert`, `track_delete`, `list_tracked_changes`, `accept_tracked_changes`, `reject_tracked_changes`
- Cross-platform comments: `add_comment` anchored to text
- Cross-platform hyperlinks: `manage_hyperlinks` (add, list, remove, update)
- Cross-platform layout tools: `set_page_layout`, `add_header_footer`, `add_page_numbers`, `add_section_break`, `set_paragraph_spacing`, `add_bookmark`, `add_watermark`
- Cross-platform footnote tools (10): add, delete, validate, customize footnotes and endnotes
- Cross-platform protection tools: `protect_document`, `unprotect_document`, `add_restricted_editing`, `add_digital_signature`, `verify_document`
- Multiple transport support: stdio (default), SSE, streamable-http
- `MCP_AUTHOR` / `MCP_AUTHOR_INITIALS` environment variables for author metadata
- PyPI packaging as `word-mcp-live`

## [1.0.0] - 2024-12-01

### Added
- Initial release based on [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server)
- 54 cross-platform tools using python-docx
- Document management, content editing, formatting, tables, extraction
- FastMCP server with stdio transport

[1.5.1]: https://github.com/cardtest15-coder/word-mcp-live/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/cardtest15-coder/word-mcp-live/compare/v1.4.1...v1.5.0
[1.4.1]: https://github.com/cardtest15-coder/word-mcp-live/compare/v1.4.0...v1.4.1
[1.3.0]: https://github.com/cardtest15-coder/word-mcp-live/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/cardtest15-coder/word-mcp-live/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/cardtest15-coder/word-mcp-live/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/cardtest15-coder/word-mcp-live/releases/tag/v1.0.0
