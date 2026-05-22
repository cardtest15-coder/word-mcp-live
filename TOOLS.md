# Tool Reference

Complete list of all **195 tools** provided by word-mcp-live.

---

## Cross-Platform Tools (115)

These work on Windows, macOS, and Linux using python-docx. The document file must be **closed** (not open in Word).

<details>
<summary><b>Document Management (7)</b></summary>

| Tool | Description |
|------|-------------|
| `create_document` | Create a new Word document with optional metadata |
| `copy_document` | Create a copy of a Word document |
| `get_document_info` | Get document properties and statistics |
| `get_document_text` | Extract all text from a document |
| `get_document_outline` | Get document heading structure |
| `list_available_documents` | List .docx files in a directory |
| `get_document_xml` | Get raw OOXML structure |

</details>

<details>
<summary><b>Content (14)</b></summary>

| Tool | Description |
|------|-------------|
| `add_paragraph` | Add a paragraph with optional formatting |
| `add_heading` | Add a heading (levels 1-9) with formatting |
| `add_table` | Add a table with custom data |
| `add_picture` | Add an image with proportional scaling |
| `add_page_break` | Insert a page break |
| `delete_paragraph` | Delete a paragraph by index |
| `search_and_replace` | Find and replace text |
| `add_table_of_contents` | Add a TOC based on heading styles |
| `insert_header_near_text` | Insert a heading before/after target text |
| `insert_line_or_paragraph_near_text` | Insert a paragraph before/after target text |
| `insert_numbered_list_near_text` | Insert a bulleted or numbered list |
| `replace_paragraph_block_below_header` | Replace content under a heading |
| `replace_block_between_manual_anchors` | Replace content between anchor texts |
| `merge_documents` | Merge multiple documents into one |

</details>

<details>
<summary><b>Formatting (27)</b></summary>

| Tool | Description |
|------|-------------|
| `format_text` | Format text (bold, italic, color, font, size) |
| `create_custom_style` | Create a custom document style |
| `modify_style` | Modify an existing style |
| `delete_style` | Delete a custom style |
| `apply_style` | Apply a named style to a paragraph |
| `format_table` | Format table borders and structure |
| `set_table_cell_shading` | Set cell background color |
| `apply_table_alternating_rows` | Alternating row colors |
| `highlight_table_header` | Highlight header row |
| `merge_table_cells` | Merge a rectangular cell area |
| `merge_table_cells_horizontal` | Merge cells in a row |
| `merge_table_cells_vertical` | Merge cells in a column |
| `set_table_cell_alignment` | Set cell text alignment |
| `set_table_alignment_all` | Set alignment for all cells |
| `set_table_column_width` | Set a column's width |
| `set_table_column_widths` | Set multiple column widths |
| `set_table_width` | Set overall table width |
| `auto_fit_table_columns` | Auto-fit columns to content |
| `format_table_cell_text` | Format text in a specific cell |
| `set_table_cell_padding` | Set cell padding |
| `highlight_text` | Apply highlighting to a text range |
| `remove_highlight` | Remove highlighting |
| `set_tab_stops` | Set tab stops for a paragraph |
| `clear_tab_stops` | Clear all tab stops |
| `insert_drop_cap` | Insert a drop cap effect |
| `copy_styles_from_template` | Copy styles from another document |
| `insert_text_box` | Insert a text box |

</details>

<details>
<summary><b>Comments (4)</b></summary>

| Tool | Description |
|------|-------------|
| `get_all_comments` | Extract all comments |
| `get_comments_by_author` | Filter comments by author |
| `get_comments_for_paragraph` | Get comments for a specific paragraph |
| `add_comment` | Add a comment anchored to text |

</details>

<details>
<summary><b>Tracked Changes (6)</b></summary>

| Tool | Description |
|------|-------------|
| `track_replace` | Replace text as a tracked change |
| `track_insert` | Insert text as a tracked change |
| `track_delete` | Delete text as a tracked change |
| `list_tracked_changes` | List all tracked changes |
| `accept_tracked_changes` | Accept all tracked changes |
| `reject_tracked_changes` | Reject all tracked changes |

</details>

<details>
<summary><b>Hyperlinks (1)</b></summary>

| Tool | Description |
|------|-------------|
| `manage_hyperlinks` | Add, insert, list, remove, and update hyperlinks |

</details>

<details>
<summary><b>Layout (17)</b></summary>

| Tool | Description |
|------|-------------|
| `set_page_layout` | Set orientation, size, and margins |
| `add_header_footer` | Add header/footer text |
| `add_page_numbers` | Add page numbers |
| `add_section_break` | Add section break (new page, continuous, etc.) |
| `set_paragraph_spacing` | Set paragraph spacing |
| `add_bookmark` | Add a named bookmark |
| `add_watermark` | Add a diagonal text watermark |
| `set_section_columns` | Set column layout for a section |
| `set_column_widths` | Set unequal column widths |
| `set_different_first_page` | Different first page header/footer |
| `set_odd_even_headers` | Odd/even page headers/footers |
| `set_page_borders` | Set page borders for a section |
| `insert_column_break` | Insert a column break |
| `get_section_layout` | Get section layout properties |
| `insert_chart` | Insert a chart |
| `insert_citation` | Insert a citation field |
| `insert_index` | Insert a subject index |

</details>

<details>
<summary><b>Footnotes (10)</b></summary>

| Tool | Description |
|------|-------------|
| `add_footnote_to_document` | Add a footnote to a paragraph |
| `add_footnote_after_text` | Add footnote after specific text |
| `add_footnote_before_text` | Add footnote before specific text |
| `add_footnote_enhanced` | Enhanced footnote with superscript |
| `add_footnote_robust` | Robust footnote with validation |
| `add_endnote_to_document` | Add an endnote |
| `customize_footnote_style` | Customize footnote numbering |
| `delete_footnote_from_document` | Delete a footnote |
| `delete_footnote_robust` | Delete with cleanup |
| `validate_document_footnotes` | Validate all footnotes |
| `convert_footnotes_to_endnotes` | Convert all footnotes to endnotes |

</details>

<details>
<summary><b>Protection (5)</b></summary>

| Tool | Description |
|------|-------------|
| `protect_document` | Add password protection |
| `unprotect_document` | Remove protection |
| `add_restricted_editing` | Restrict editing to specific sections |
| `add_digital_signature` | Add a digital signature |
| `verify_document` | Verify protection and signatures |

</details>

<details>
<summary><b>Export (3)</b></summary>

| Tool | Description |
|------|-------------|
| `convert_to_pdf` | Convert to PDF |
| `export_to_html` | Export to HTML |
| `export_to_rtf` | Export to RTF |
| `export_to_txt` | Export to plain text |

</details>

<details>
<summary><b>Extraction (6)</b></summary>

| Tool | Description |
|------|-------------|
| `get_paragraph_text_from_document` | Get text from a specific paragraph |
| `find_text_in_document` | Find text occurrences |
| `get_highlighted_text` | Extract highlighted/colored text |
| `get_style_details` | Get detailed style properties |
| `list_styles` | List all styles in a document |
| `get_custom_properties` | Get custom document properties |

</details>

<details>
<summary><b>Advanced Table (7)</b></summary>

| Tool | Description |
|------|-------------|
| `sort_table` | Sort a table by a column |
| `delete_table` | Delete a table |
| `convert_table_to_text` | Convert table to plain text |
| `convert_text_to_table` | Convert text into a table |
| `repeat_table_header` | Set header rows to repeat on each page |
| `compare_documents` | Compare two documents (legal blackline) |
| `mail_merge` | Perform mail merge with template |

</details>

<details>
<summary><b>Fields & Properties (7)</b></summary>

| Tool | Description |
|------|-------------|
| `insert_field` | Insert a field code (DATE, PAGE, etc.) |
| `insert_content_control` | Insert a content control (text, dropdown, date, checkbox) |
| `update_fields` | Update all fields (TOC, PAGE numbers, etc.) |
| `update_table_of_contents` | Update/refresh existing TOC |
| `set_custom_property` | Set a custom document property |
| `update_table_of_contents` | Update/refresh existing TOC |
| `word_screen_capture` | Screenshot of the Word window |

</details>

---

## Windows/macOS Live Tools (80)

These require Windows (COM) or macOS (JXA) with Microsoft Word installed. They operate on documents **currently open in Word**. Word is auto-started if not running; documents are auto-opened if not already open.

<details open>
<summary><b>Editing (17)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_insert_text` | Insert text at any position (auto-chunked for large text) |
| `word_live_delete_text` | Delete a character range |
| `word_live_replace_text` | Find & replace via COM — supports wildcards (`^s`, `^m`, `^p`, `^t`) |
| `word_live_insert_paragraphs` | Insert multiple paragraphs near a target (Find-based, fast) |
| `word_live_format_text` | Format text: font, color, highlight, style, alignment, page break |
| `word_live_add_table` | Insert a table with optional style and data |
| `word_live_format_table` | Format table: borders, cell shading, alignment, column widths |
| `word_live_modify_table` | Full table CRUD: get_info, set_cell, set_row, add/delete rows/columns, merge, autofit |
| `word_live_sort_table` | Sort a table by any column |
| `word_live_delete_table` | Delete a table with orphan cleanup |
| `word_live_save` | Save in place or save-as (docx, pdf, rtf, txt) |
| `word_live_toggle_track_changes` | Toggle or set track changes mode |
| `word_live_insert_image` | Insert image with sizing, alignment, wrapping, border |
| `word_live_insert_cross_reference` | Insert live cross-reference to headings, bookmarks, figures |
| `word_live_insert_equation` | Insert equation using UnicodeMath syntax |
| `word_live_apply_list` | Apply bullet/number/multilevel list formatting |
| `word_live_setup_heading_numbering` | Auto-numbered headings (1. / 1.1) with locale-aware style resolution |

</details>

<details open>
<summary><b>Reading (13)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_list_open` | List all open documents with name, path, pages, saved status |
| `word_live_get_text` | Get all text (capped at 200 paragraphs / 3 pages) |
| `word_live_get_page_text` | Get text from specific page(s) with char offsets |
| `word_live_get_paragraph_format` | Inspect formatting: font, spacing, alignment, list info, per-run detail |
| `word_live_get_info` | Document metadata: pages, words, sections, tables |
| `word_live_find_text` | Find text with context; supports wildcards |
| `word_live_get_section_layout` | Get section layout (orientation, margins, columns) |
| `word_live_list_styles` | List all document styles |
| `word_live_list_cross_reference_items` | List available cross-reference targets |
| `word_live_diagnose_layout` | Scan for layout problems (keep_with_next chains, style misuse) |
| `word_live_get_undo_history` | List undo stack entries |
| `word_live_take_snapshot` | Store paragraph baseline for diffing |
| `word_live_get_diff` | Compare against snapshot — returns only changes |
| `word_live_snapshot_status` | Check snapshot existence and age |
| `word_live_spell_check` | Get spelling error count and list |
| `word_live_check_grammar` | Run grammar check |

</details>

<details open>
<summary><b>Comments & Revisions (8)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_get_comments` | Get all comments |
| `word_live_add_comment` | Add a comment anchored to text |
| `word_live_delete_comment` | Permanently delete a comment |
| `word_live_reply_to_comment` | Add a threaded reply (Word 2016+) |
| `word_live_resolve_comment` | Resolve/unresolve a comment (Word 2016+) |
| `word_live_list_revisions` | List tracked changes |
| `word_live_accept_revisions` | Accept tracked changes (all or by author) |
| `word_live_reject_revisions` | Reject tracked changes (all or by author) |

</details>

<details open>
<summary><b>Layout & Design (20)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_set_page_layout` | Set orientation, size, margins |
| `word_live_add_header_footer` | Add header/footer text |
| `word_live_add_page_numbers` | Add page numbers with optional prefix/suffix |
| `word_live_add_section_break` | Add section break |
| `word_live_set_paragraph_spacing` | Set spacing, keep_with_next, keep_together, alignment |
| `word_live_add_bookmark` | Add a named bookmark |
| `word_live_add_watermark` | Add a diagonal text watermark |
| `word_live_set_section_columns` | Set column layout for a section |
| `word_live_set_different_first_page` | Different first page header/footer |
| `word_live_set_odd_even_headers` | Odd/even page headers/footers |
| `word_live_set_page_borders` | Set page borders (single, double, dashed, dotted, none) |
| `word_live_insert_column_break` | Insert a column break |
| `word_live_apply_style` | Apply a named style (locale-aware: English and local names) |
| `word_live_modify_style` | Modify style properties (font, size, color, spacing) |
| `word_live_highlight_text` | Apply highlighting to a text range |
| `word_live_set_custom_property` | Set custom doc property (with zipfile fallback) |
| `word_live_get_custom_properties` | Get all custom properties |
| `word_live_set_core_properties` | Set built-in properties (Title, Author, Keywords, etc.) |
| `word_live_insert_field` | Insert a field code (DATE, PAGE, AUTHOR, etc.) |
| `word_live_insert_content_control` | Insert content control (text, dropdown, date, checkbox) |

</details>

<details open>
<summary><b>Objects & Insertion (10)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_insert_shape` | Insert a shape (rectangle, oval, line, arrow, callout) |
| `word_live_insert_text_box` | Insert a text box at specified position |
| `word_live_insert_chart` | Insert a chart (bar, column, line, pie, area, scatter) |
| `word_live_insert_drop_cap` | Insert a drop cap on first character |
| `word_live_insert_table_of_figures` | Insert a table of figures |
| `word_live_insert_index` | Insert a subject index |
| `word_live_insert_image` | Insert image (see Editing above) |
| `word_live_repeat_table_header` | Set header rows to repeat on each page |
| `word_live_update_fields` | Update all fields (TOC, PAGE, etc.) |
| `word_live_update_table_of_contents` | Refresh existing TOC |

</details>

<details open>
<summary><b>Export & Utility (8)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_export_to_html` | Export to HTML (via COM) |
| `word_live_export_to_rtf` | Export to RTF (via COM) |
| `word_live_export_to_txt` | Export to plain text (via COM) |
| `word_live_mail_merge` | Perform mail merge with data source |
| `word_live_compare_documents` | Compare two documents (legal blackline) |
| `word_live_print` | Print document with optional page range |
| `word_live_run_macro` | Execute a VBA macro |
| `word_live_undo` | Undo last N operations (each MCP call = one entry) |

</details>

<details open>
<summary><b>Variables & Screen (4)</b></summary>

| Tool | Description |
|------|-------------|
| `word_live_get_doc_variable` | Get a document variable |
| `word_live_set_doc_variable` | Set a document variable |
| `word_live_list_revisions` | List tracked changes (see Comments & Revisions) |
| `word_screen_capture` | Screenshot of the Word window |

</details>
