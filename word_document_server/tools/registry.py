"""Declarative tool registry for Word Document MCP Server."""
from __future__ import annotations

from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    comment_tools,
    comment_write_tools,
    hyperlink_tools,
    tracked_changes_tools,
    live_tools,
    live_read_tools,
    live_layout_tools,
    screen_capture_tools,
    layout_tools,
    style_tools,
    column_tools,
    advanced_table_tools,
    object_tools,
    export_tools,
    merge_tools,
    highlight_tools,
    field_tools,
    property_tools,
    page_design_tools,
    live_doc_tools,
)
from word_document_server.tools.content_tools import (
    replace_paragraph_block_below_header_tool,
    replace_block_between_manual_anchors_tool,
)


ToolDef = dict  # {"name": str | None, "fn": callable, "annotations": dict}


TOOLS: list[ToolDef] = [
    # --- Document tools ---
    {
        "name": "create_document",
        "fn": document_tools.create_document,
        "annotations": {"title": "Create Word Document", "destructiveHint": True},
    },
    {
        "name": "copy_document",
        "fn": document_tools.copy_document,
        "annotations": {"title": "Copy Word Document", "destructiveHint": True},
    },
    {
        "name": "get_document_info",
        "fn": document_tools.get_document_info,
        "annotations": {"title": "Get Document Info", "readOnlyHint": True},
    },
    {
        "name": "get_document_text",
        "fn": document_tools.get_document_text,
        "annotations": {"title": "Get Document Text", "readOnlyHint": True},
    },
    {
        "name": "get_document_outline",
        "fn": document_tools.get_document_outline,
        "annotations": {"title": "Get Document Outline", "readOnlyHint": True},
    },
    {
        "name": "list_available_documents",
        "fn": document_tools.list_available_documents,
        "annotations": {"title": "List Available Documents", "readOnlyHint": True},
    },
    {
        "name": "get_document_xml",
        "fn": document_tools.get_document_xml_tool,
        "annotations": {"title": "Get Document XML", "readOnlyHint": True},
    },
    # --- Content: insert-near tools ---
    {
        "name": "insert_header_near_text",
        "fn": content_tools.insert_header_near_text_tool,
        "annotations": {"title": "Insert Header Near Text", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "insert_line_or_paragraph_near_text",
        "fn": content_tools.insert_line_or_paragraph_near_text_tool,
        "annotations": {"title": "Insert Line Near Text", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "insert_numbered_list_near_text",
        "fn": content_tools.insert_numbered_list_near_text_tool,
        "annotations": {"title": "Insert List Near Text", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Content tools ---
    {
        "name": "add_paragraph",
        "fn": content_tools.add_paragraph,
        "annotations": {"title": "Add Paragraph", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_heading",
        "fn": content_tools.add_heading,
        "annotations": {"title": "Add Heading", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_picture",
        "fn": content_tools.add_picture,
        "annotations": {"title": "Add Picture", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_table",
        "fn": content_tools.add_table,
        "annotations": {"title": "Add Table", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_page_break",
        "fn": content_tools.add_page_break,
        "annotations": {"title": "Add Page Break", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "delete_paragraph",
        "fn": content_tools.delete_paragraph,
        "annotations": {"title": "Delete Paragraph", "destructiveHint": True},
    },
    {
        "name": "search_and_replace",
        "fn": content_tools.search_and_replace,
        "annotations": {"title": "Search and Replace", "destructiveHint": True},
    },
    # --- Format tools ---
    {
        "name": "create_custom_style",
        "fn": format_tools.create_custom_style,
        "annotations": {"title": "Create Custom Style", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "format_text",
        "fn": format_tools.format_text,
        "annotations": {"title": "Format Text", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "format_table",
        "fn": format_tools.format_table,
        "annotations": {"title": "Format Table", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Table cell shading ---
    {
        "name": "set_table_cell_shading",
        "fn": format_tools.set_table_cell_shading,
        "annotations": {"title": "Set Table Cell Shading", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "apply_table_alternating_rows",
        "fn": format_tools.apply_table_alternating_rows,
        "annotations": {"title": "Apply Alternating Row Colors", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "highlight_table_header",
        "fn": format_tools.highlight_table_header,
        "annotations": {"title": "Highlight Table Header", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Cell merging ---
    {
        "name": "merge_table_cells",
        "fn": format_tools.merge_table_cells,
        "annotations": {"title": "Merge Table Cells", "readOnlyHint": False, "destructiveHint": True},
    },
    {
        "name": "merge_table_cells_horizontal",
        "fn": format_tools.merge_table_cells_horizontal,
        "annotations": {"title": "Merge Cells Horizontally", "readOnlyHint": False, "destructiveHint": True},
    },
    {
        "name": "merge_table_cells_vertical",
        "fn": format_tools.merge_table_cells_vertical,
        "annotations": {"title": "Merge Cells Vertically", "readOnlyHint": False, "destructiveHint": True},
    },
    # --- Cell alignment ---
    {
        "name": "set_table_cell_alignment",
        "fn": format_tools.set_table_cell_alignment,
        "annotations": {"title": "Set Cell Alignment", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "set_table_alignment_all",
        "fn": format_tools.set_table_alignment_all,
        "annotations": {"title": "Set Table Alignment", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Protection tools ---
    {
        "name": "protect_document",
        "fn": protection_tools.protect_document,
        "annotations": {"title": "Protect Document", "readOnlyHint": False, "destructiveHint": True},
    },
    {
        "name": "unprotect_document",
        "fn": protection_tools.unprotect_document,
        "annotations": {"title": "Unprotect Document", "readOnlyHint": False, "destructiveHint": True},
    },
    # --- Footnote tools ---
    {
        "name": "add_footnote_to_document",
        "fn": footnote_tools.add_footnote_to_document,
        "annotations": {"title": "Add Footnote", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_footnote_after_text",
        "fn": footnote_tools.add_footnote_after_text,
        "annotations": {"title": "Add Footnote After Text", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_footnote_before_text",
        "fn": footnote_tools.add_footnote_before_text,
        "annotations": {"title": "Add Footnote Before Text", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_footnote_enhanced",
        "fn": footnote_tools.add_footnote_enhanced,
        "annotations": {"title": "Add Footnote Enhanced", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "add_endnote_to_document",
        "fn": footnote_tools.add_endnote_to_document,
        "annotations": {"title": "Add Endnote", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "convert_footnotes_to_endnotes",
        "fn": footnote_tools.convert_footnotes_to_endnotes_in_document,
        "annotations": {"title": "Convert Footnotes to Endnotes", "readOnlyHint": False, "destructiveHint": True},
    },
    {
        "name": "customize_footnote_style",
        "fn": footnote_tools.customize_footnote_style,
        "annotations": {"title": "Customize Footnote Style", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "delete_footnote_from_document",
        "fn": footnote_tools.delete_footnote_from_document,
        "annotations": {"title": "Delete Footnote", "destructiveHint": True},
    },
    # --- Robust footnote tools ---
    {
        "name": "add_footnote_robust",
        "fn": footnote_tools.add_footnote_robust_tool,
        "annotations": {"title": "Add Footnote Robust", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "validate_document_footnotes",
        "fn": footnote_tools.validate_footnotes_tool,
        "annotations": {"title": "Validate Footnotes", "readOnlyHint": True},
    },
    {
        "name": "delete_footnote_robust",
        "fn": footnote_tools.delete_footnote_robust_tool,
        "annotations": {"title": "Delete Footnote Robust", "destructiveHint": True},
    },
    # --- Extended document tools ---
    {
        "name": "get_paragraph_text_from_document",
        "fn": extended_document_tools.get_paragraph_text_from_document,
        "annotations": {"title": "Get Paragraph Text", "readOnlyHint": True},
    },
    {
        "name": "find_text_in_document",
        "fn": extended_document_tools.find_text_in_document,
        "annotations": {"title": "Find Text", "readOnlyHint": True},
    },
    {
        "name": "get_highlighted_text",
        "fn": extended_document_tools.get_highlighted_text_from_document,
        "annotations": {"title": "Get Highlighted Text", "readOnlyHint": True},
    },
    {
        "name": "convert_to_pdf",
        "fn": extended_document_tools.convert_to_pdf,
        "annotations": {"title": "Convert to PDF", "destructiveHint": True},
    },
    {
        "name": "replace_paragraph_block_below_header",
        "fn": replace_paragraph_block_below_header_tool,
        "annotations": {"title": "Replace Block Below Header", "readOnlyHint": False, "destructiveHint": True},
    },
    {
        "name": "replace_block_between_manual_anchors",
        "fn": replace_block_between_manual_anchors_tool,
        "annotations": {"title": "Replace Block Between Anchors", "readOnlyHint": False, "destructiveHint": True},
    },
    # --- Comment read tools ---
    {
        "name": "get_all_comments",
        "fn": comment_tools.get_all_comments,
        "annotations": {"title": "Get All Comments", "readOnlyHint": True},
    },
    {
        "name": "get_comments_by_author",
        "fn": comment_tools.get_comments_by_author,
        "annotations": {"title": "Get Comments by Author", "readOnlyHint": True},
    },
    {
        "name": "get_comments_for_paragraph",
        "fn": comment_tools.get_comments_for_paragraph,
        "annotations": {"title": "Get Comments for Paragraph", "readOnlyHint": True},
    },
    # --- Comment write tools ---
    {
        "name": "add_comment",
        "fn": comment_write_tools.add_comment,
        "annotations": {"title": "Add Comment", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Hyperlink tools ---
    {
        "name": "manage_hyperlinks",
        "fn": hyperlink_tools.manage_hyperlinks,
        "annotations": {"title": "Manage Hyperlinks", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Table column width tools ---
    {
        "name": "set_table_column_width",
        "fn": format_tools.set_table_column_width,
        "annotations": {"title": "Set Column Width", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "set_table_column_widths",
        "fn": format_tools.set_table_column_widths,
        "annotations": {"title": "Set Column Widths", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "set_table_width",
        "fn": format_tools.set_table_width,
        "annotations": {"title": "Set Table Width", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "auto_fit_table_columns",
        "fn": format_tools.auto_fit_table_columns,
        "annotations": {"title": "Auto-Fit Table Columns", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Table cell text formatting and padding ---
    {
        "name": "format_table_cell_text",
        "fn": format_tools.format_table_cell_text,
        "annotations": {"title": "Format Cell Text", "readOnlyHint": False, "destructiveHint": False},
    },
    {
        "name": "set_table_cell_padding",
        "fn": format_tools.set_table_cell_padding,
        "annotations": {"title": "Set Cell Padding", "readOnlyHint": False, "destructiveHint": False},
    },
    # --- Tracked changes tools ---
    {
        "name": "track_replace",
        "fn": tracked_changes_tools.track_replace,
        "annotations": {"title": "Track Replace", "destructiveHint": True},
    },
    {
        "name": "track_insert",
        "fn": tracked_changes_tools.track_insert,
        "annotations": {"title": "Track Insert", "destructiveHint": True},
    },
    {
        "name": "track_delete",
        "fn": tracked_changes_tools.track_delete,
        "annotations": {"title": "Track Delete", "destructiveHint": True},
    },
    {
        "name": "list_tracked_changes",
        "fn": tracked_changes_tools.list_tracked_changes,
        "annotations": {"title": "List Tracked Changes", "readOnlyHint": True},
    },
    {
        "name": "accept_tracked_changes",
        "fn": tracked_changes_tools.accept_tracked_changes,
        "annotations": {"title": "Accept Tracked Changes", "destructiveHint": True},
    },
    {
        "name": "reject_tracked_changes",
        "fn": tracked_changes_tools.reject_tracked_changes,
        "annotations": {"title": "Reject Tracked Changes", "destructiveHint": True},
    },
    # --- Live editing tools (Windows only, requires Word running) ---
    {
        "name": "word_screen_capture",
        "fn": screen_capture_tools.word_screen_capture,
        "annotations": {"title": "Word Screen Capture", "readOnlyHint": True},
    },
    {
        "name": "word_live_insert_text",
        "fn": live_tools.word_live_insert_text,
        "annotations": {"title": "Word Live Insert Text", "destructiveHint": True},
    },
    {
        "name": "word_live_format_text",
        "fn": live_tools.word_live_format_text,
        "annotations": {"title": "Word Live Format Text", "destructiveHint": True},
    },
    {
        "name": "word_live_replace_text",
        "fn": live_tools.word_live_replace_text,
        "annotations": {"title": "Word Live Replace Text", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_paragraphs",
        "fn": live_tools.word_live_insert_paragraphs,
        "annotations": {"title": "Word Live Insert Paragraphs", "destructiveHint": True},
    },
    {
        "name": "word_live_add_table",
        "fn": live_tools.word_live_add_table,
        "annotations": {"title": "Word Live Add Table", "destructiveHint": True},
    },
    {
        "name": "word_live_format_table",
        "fn": live_tools.word_live_format_table,
        "annotations": {"title": "Word Live Format Table", "destructiveHint": True},
    },
    {
        "name": "word_live_modify_table",
        "fn": live_tools.word_live_modify_table,
        "annotations": {"title": "Word Live Modify Table", "destructiveHint": True},
    },
    {
        "name": "word_live_delete_text",
        "fn": live_tools.word_live_delete_text,
        "annotations": {"title": "Word Live Delete Text", "destructiveHint": True},
    },
    {
        "name": "word_live_apply_list",
        "fn": live_tools.word_live_apply_list,
        "annotations": {"title": "Word Live Apply List", "destructiveHint": True},
    },
    {
        "name": "word_live_setup_heading_numbering",
        "fn": live_tools.word_live_setup_heading_numbering,
        "annotations": {"title": "Word Live Setup Heading Numbering", "destructiveHint": True},
    },
    # --- Live read tools (Windows only, requires Word running) ---
    {
        "name": "word_live_get_text",
        "fn": live_read_tools.word_live_get_text,
        "annotations": {"title": "Word Live Get Text", "readOnlyHint": True},
    },
    {
        "name": "word_live_take_snapshot",
        "fn": live_read_tools.word_live_take_snapshot,
        "annotations": {"title": "Word Live Take Snapshot", "readOnlyHint": True},
    },
    {
        "name": "word_live_get_diff",
        "fn": live_read_tools.word_live_get_diff,
        "annotations": {"title": "Word Live Get Diff", "readOnlyHint": True},
    },
    {
        "name": "word_live_snapshot_status",
        "fn": live_read_tools.word_live_snapshot_status,
        "annotations": {"title": "Word Live Snapshot Status", "readOnlyHint": True},
    },
    {
        "name": "word_live_get_paragraph_format",
        "fn": live_read_tools.word_live_get_paragraph_format,
        "annotations": {"title": "Word Live Get Paragraph Format", "readOnlyHint": True},
    },
    {
        "name": "word_live_get_info",
        "fn": live_read_tools.word_live_get_info,
        "annotations": {"title": "Word Live Get Info", "readOnlyHint": True},
    },
    {
        "name": "word_live_set_core_properties",
        "fn": live_read_tools.word_live_set_core_properties,
        "annotations": {"title": "Word Live Set Core Properties", "destructiveHint": True},
    },
    {
        "name": "word_live_list_open",
        "fn": live_read_tools.word_live_list_open,
        "annotations": {"title": "Word Live List Open", "readOnlyHint": True},
    },
    {
        "name": "word_live_find_text",
        "fn": live_read_tools.word_live_find_text,
        "annotations": {"title": "Word Live Find Text", "readOnlyHint": True},
    },
    {
        "name": "word_live_get_comments",
        "fn": live_read_tools.word_live_get_comments,
        "annotations": {"title": "Word Live Get Comments", "readOnlyHint": True},
    },
    {
        "name": "word_live_add_comment",
        "fn": live_read_tools.word_live_add_comment,
        "annotations": {"title": "Word Live Add Comment", "destructiveHint": True},
    },
    {
        "name": "word_live_reply_to_comment",
        "fn": live_read_tools.word_live_reply_to_comment,
        "annotations": {"title": "Word Live Reply to Comment", "destructiveHint": True},
    },
    {
        "name": "word_live_resolve_comment",
        "fn": live_read_tools.word_live_resolve_comment,
        "annotations": {"title": "Word Live Resolve Comment", "destructiveHint": True},
    },
    {
        "name": "word_live_delete_comment",
        "fn": live_read_tools.word_live_delete_comment,
        "annotations": {"title": "Word Live Delete Comment", "destructiveHint": True},
    },
    {
        "name": "word_live_list_revisions",
        "fn": live_read_tools.word_live_list_revisions,
        "annotations": {"title": "Word Live List Revisions", "readOnlyHint": True},
    },
    {
        "name": "word_live_accept_revisions",
        "fn": live_read_tools.word_live_accept_revisions,
        "annotations": {"title": "Word Live Accept Revisions", "destructiveHint": True},
    },
    {
        "name": "word_live_reject_revisions",
        "fn": live_read_tools.word_live_reject_revisions,
        "annotations": {"title": "Word Live Reject Revisions", "destructiveHint": True},
    },
    {
        "name": "word_live_get_page_text",
        "fn": live_read_tools.word_live_get_page_text,
        "annotations": {"title": "Word Live Get Page Text", "readOnlyHint": True},
    },
    {
        "name": "word_live_get_undo_history",
        "fn": live_read_tools.word_live_get_undo_history,
        "annotations": {"title": "Word Live Get Undo History", "readOnlyHint": True},
    },
    {
        "name": "word_live_undo",
        "fn": live_tools.word_live_undo,
        "annotations": {"title": "Word Live Undo", "destructiveHint": True},
    },
    {
        "name": "word_live_save",
        "fn": live_tools.word_live_save,
        "annotations": {"title": "Word Live Save", "destructiveHint": True},
    },
    {
        "name": "word_live_toggle_track_changes",
        "fn": live_tools.word_live_toggle_track_changes,
        "annotations": {"title": "Word Live Toggle Track Changes", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_image",
        "fn": live_tools.word_live_insert_image,
        "annotations": {"title": "Word Live Insert Image", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_cross_reference",
        "fn": live_tools.word_live_insert_cross_reference,
        "annotations": {"title": "Word Live Insert Cross Reference", "destructiveHint": True},
    },
    {
        "name": "word_live_list_cross_reference_items",
        "fn": live_tools.word_live_list_cross_reference_items,
        "annotations": {"title": "Word Live List Cross Reference Items", "readOnlyHint": True},
    },
    {
        "name": "word_live_insert_equation",
        "fn": live_tools.word_live_insert_equation,
        "annotations": {"title": "Word Live Insert Equation", "destructiveHint": True},
    },
    {
        "name": "word_live_diagnose_layout",
        "fn": live_read_tools.word_live_diagnose_layout,
        "annotations": {"title": "Word Live Diagnose Layout", "readOnlyHint": True},
    },
    # --- Live layout tools (Windows only, requires Word running) ---
    {
        "name": "word_live_set_page_layout",
        "fn": live_layout_tools.word_live_set_page_layout,
        "annotations": {"title": "Word Live Set Page Layout", "destructiveHint": True},
    },
    {
        "name": "word_live_add_header_footer",
        "fn": live_layout_tools.word_live_add_header_footer,
        "annotations": {"title": "Word Live Add Header/Footer", "destructiveHint": True},
    },
    {
        "name": "word_live_add_page_numbers",
        "fn": live_layout_tools.word_live_add_page_numbers,
        "annotations": {"title": "Word Live Add Page Numbers", "destructiveHint": True},
    },
    {
        "name": "word_live_add_section_break",
        "fn": live_layout_tools.word_live_add_section_break,
        "annotations": {"title": "Word Live Add Section Break", "destructiveHint": True},
    },
    {
        "name": "word_live_set_paragraph_spacing",
        "fn": live_layout_tools.word_live_set_paragraph_spacing,
        "annotations": {"title": "Word Live Set Paragraph Spacing", "destructiveHint": True},
    },
    {
        "name": "word_live_add_bookmark",
        "fn": live_layout_tools.word_live_add_bookmark,
        "annotations": {"title": "Word Live Add Bookmark", "destructiveHint": True},
    },
    {
        "name": "word_live_add_watermark",
        "fn": live_layout_tools.word_live_add_watermark,
        "annotations": {"title": "Word Live Add Watermark", "destructiveHint": True},
    },
    # --- Layout, header/footer, spacing, bookmark, watermark tools ---
    {
        "name": "set_page_layout",
        "fn": layout_tools.set_page_layout,
        "annotations": {"title": "Set Page Layout", "destructiveHint": True},
    },
    {
        "name": "add_header_footer",
        "fn": layout_tools.add_header_footer,
        "annotations": {"title": "Add Header/Footer", "destructiveHint": True},
    },
    {
        "name": "add_page_numbers",
        "fn": layout_tools.add_page_numbers,
        "annotations": {"title": "Add Page Numbers", "destructiveHint": True},
    },
    {
        "name": "add_section_break",
        "fn": layout_tools.add_section_break,
        "annotations": {"title": "Add Section Break", "destructiveHint": True},
    },
    {
        "name": "set_paragraph_spacing",
        "fn": layout_tools.set_paragraph_spacing,
        "annotations": {"title": "Set Paragraph Spacing", "destructiveHint": True},
    },
    {
        "name": "add_bookmark",
        "fn": layout_tools.add_bookmark,
        "annotations": {"title": "Add Bookmark", "destructiveHint": True},
    },
    {
        "name": "add_watermark",
        "fn": layout_tools.add_watermark,
        "annotations": {"title": "Add Watermark", "destructiveHint": True},
    },
    # --- Previously unregistered existing tools ---
    {
        "name": "add_table_of_contents",
        "fn": content_tools.add_table_of_contents,
        "annotations": {"title": "Add Table of Contents", "destructiveHint": True},
    },
    {
        "name": "merge_documents",
        "fn": document_tools.merge_documents,
        "annotations": {"title": "Merge Documents", "destructiveHint": True},
    },
    {
        "name": "add_restricted_editing",
        "fn": protection_tools.add_restricted_editing,
        "annotations": {"title": "Add Restricted Editing", "destructiveHint": True},
    },
    {
        "name": "add_digital_signature",
        "fn": protection_tools.add_digital_signature,
        "annotations": {"title": "Add Digital Signature", "destructiveHint": True},
    },
    {
        "name": "verify_document",
        "fn": protection_tools.verify_document,
        "annotations": {"title": "Verify Document", "readOnlyHint": True},
    },
    # ── Style Tools ──────────────────────────────────────────────────────
    {
        "name": "list_styles",
        "fn": style_tools.list_styles,
        "annotations": {"title": "List Styles", "readOnlyHint": True},
    },
    {
        "name": "get_style_details",
        "fn": style_tools.get_style_details,
        "annotations": {"title": "Get Style Details", "readOnlyHint": True},
    },
    {
        "name": "apply_style",
        "fn": style_tools.apply_style,
        "annotations": {"title": "Apply Style", "destructiveHint": True},
    },
    {
        "name": "modify_style",
        "fn": style_tools.modify_style,
        "annotations": {"title": "Modify Style", "destructiveHint": True},
    },
    {
        "name": "delete_style",
        "fn": style_tools.delete_style,
        "annotations": {"title": "Delete Style", "destructiveHint": True},
    },
    {
        "name": "copy_styles_from_template",
        "fn": style_tools.copy_styles_from_template,
        "annotations": {"title": "Copy Styles from Template", "destructiveHint": True},
    },
    # ── Column / Section Layout Tools ─────────────────────────────────────
    {
        "name": "set_section_columns",
        "fn": column_tools.set_section_columns,
        "annotations": {"title": "Set Section Columns", "destructiveHint": True},
    },
    {
        "name": "insert_column_break",
        "fn": column_tools.insert_column_break,
        "annotations": {"title": "Insert Column Break", "destructiveHint": True},
    },
    {
        "name": "get_section_layout",
        "fn": column_tools.get_section_layout,
        "annotations": {"title": "Get Section Layout", "readOnlyHint": True},
    },
    {
        "name": "set_column_widths",
        "fn": column_tools.set_column_widths,
        "annotations": {"title": "Set Column Widths", "destructiveHint": True},
    },
    # ── Advanced Table Tools ──────────────────────────────────────────────
    {
        "name": "delete_table",
        "fn": advanced_table_tools.delete_table,
        "annotations": {"title": "Delete Table", "destructiveHint": True},
    },
    {
        "name": "repeat_table_header",
        "fn": advanced_table_tools.repeat_table_header,
        "annotations": {"title": "Repeat Table Header", "destructiveHint": True},
    },
    {
        "name": "convert_table_to_text",
        "fn": advanced_table_tools.convert_table_to_text,
        "annotations": {"title": "Convert Table to Text", "readOnlyHint": True},
    },
    {
        "name": "convert_text_to_table",
        "fn": advanced_table_tools.convert_text_to_table,
        "annotations": {"title": "Convert Text to Table", "destructiveHint": True},
    },
    {
        "name": "sort_table",
        "fn": advanced_table_tools.sort_table,
        "annotations": {"title": "Sort Table", "destructiveHint": True},
    },
    # ── Object Insertion Tools ────────────────────────────────────────────
    {
        "name": "update_table_of_contents",
        "fn": object_tools.update_table_of_contents,
        "annotations": {"title": "Update Table of Contents", "destructiveHint": True},
    },
    {
        "name": "insert_text_box",
        "fn": object_tools.insert_text_box,
        "annotations": {"title": "Insert Text Box", "destructiveHint": True},
    },
    {
        "name": "insert_chart",
        "fn": object_tools.insert_chart,
        "annotations": {"title": "Insert Chart", "destructiveHint": True},
    },
    {
        "name": "insert_index",
        "fn": object_tools.insert_index,
        "annotations": {"title": "Insert Index", "destructiveHint": True},
    },
    {
        "name": "insert_citation",
        "fn": object_tools.insert_citation,
        "annotations": {"title": "Insert Citation", "destructiveHint": True},
    },
    # ── Live Style Tools ─────────────────────────────────────────────────
    {
        "name": "word_live_list_styles",
        "fn": live_layout_tools.word_live_list_styles,
        "annotations": {"title": "Live List Styles", "readOnlyHint": True},
    },
    {
        "name": "word_live_apply_style",
        "fn": live_layout_tools.word_live_apply_style,
        "annotations": {"title": "Live Apply Style", "destructiveHint": True},
    },
    {
        "name": "word_live_modify_style",
        "fn": live_layout_tools.word_live_modify_style,
        "annotations": {"title": "Live Modify Style", "destructiveHint": True},
    },
    # ── Live Column Tools ─────────────────────────────────────────────────
    {
        "name": "word_live_set_section_columns",
        "fn": live_layout_tools.word_live_set_section_columns,
        "annotations": {"title": "Live Set Section Columns", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_column_break",
        "fn": live_layout_tools.word_live_insert_column_break,
        "annotations": {"title": "Live Insert Column Break", "destructiveHint": True},
    },
    {
        "name": "word_live_get_section_layout",
        "fn": live_layout_tools.word_live_get_section_layout,
        "annotations": {"title": "Live Get Section Layout", "readOnlyHint": True},
    },
    # ── Live Advanced Table Tools ─────────────────────────────────────────
    {
        "name": "word_live_delete_table",
        "fn": live_layout_tools.word_live_delete_table,
        "annotations": {"title": "Live Delete Table", "destructiveHint": True},
    },
    {
        "name": "word_live_repeat_table_header",
        "fn": live_layout_tools.word_live_repeat_table_header,
        "annotations": {"title": "Live Repeat Table Header", "destructiveHint": True},
    },
    {
        "name": "word_live_sort_table",
        "fn": live_layout_tools.word_live_sort_table,
        "annotations": {"title": "Live Sort Table", "destructiveHint": True},
    },
    # ── Live Object Insertion Tools ───────────────────────────────────────
    {
        "name": "word_live_update_table_of_contents",
        "fn": live_layout_tools.word_live_update_table_of_contents,
        "annotations": {"title": "Live Update TOC", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_text_box",
        "fn": live_layout_tools.word_live_insert_text_box,
        "annotations": {"title": "Live Insert Text Box", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_chart",
        "fn": live_layout_tools.word_live_insert_chart,
        "annotations": {"title": "Live Insert Chart", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_index",
        "fn": live_layout_tools.word_live_insert_index,
        "annotations": {"title": "Live Insert Index", "destructiveHint": True},
    },
    # ── Export Tools ──────────────────────────────────────────────────────
    {
        "name": "export_to_html",
        "fn": export_tools.export_to_html,
        "annotations": {"title": "Export to HTML", "destructiveHint": True},
    },
    {
        "name": "export_to_rtf",
        "fn": export_tools.export_to_rtf,
        "annotations": {"title": "Export to RTF", "destructiveHint": True},
    },
    {
        "name": "export_to_txt",
        "fn": export_tools.export_to_txt,
        "annotations": {"title": "Export to Text", "destructiveHint": True},
    },
    # ── Merge Tools ──────────────────────────────────────────────────────
    {
        "name": "compare_documents",
        "fn": merge_tools.compare_documents,
        "annotations": {"title": "Compare Documents", "destructiveHint": True},
    },
    {
        "name": "mail_merge",
        "fn": merge_tools.mail_merge,
        "annotations": {"title": "Mail Merge", "destructiveHint": True},
    },
    # ── Highlight Tools ──────────────────────────────────────────────────
    {
        "name": "highlight_text",
        "fn": highlight_tools.highlight_text,
        "annotations": {"title": "Highlight Text", "destructiveHint": True},
    },
    {
        "name": "remove_highlight",
        "fn": highlight_tools.remove_highlight,
        "annotations": {"title": "Remove Highlight", "destructiveHint": True},
    },
    # ── Field Tools ──────────────────────────────────────────────────────
    {
        "name": "insert_field",
        "fn": field_tools.insert_field,
        "annotations": {"title": "Insert Field", "destructiveHint": True},
    },
    {
        "name": "update_fields",
        "fn": field_tools.update_fields,
        "annotations": {"title": "Update Fields", "destructiveHint": True},
    },
    {
        "name": "insert_content_control",
        "fn": field_tools.insert_content_control,
        "annotations": {"title": "Insert Content Control", "destructiveHint": True},
    },
    # ── Property Tools ───────────────────────────────────────────────────
    {
        "name": "get_custom_properties",
        "fn": property_tools.get_custom_properties,
        "annotations": {"title": "Get Custom Properties", "readOnlyHint": True},
    },
    {
        "name": "set_custom_property",
        "fn": property_tools.set_custom_property,
        "annotations": {"title": "Set Custom Property", "destructiveHint": True},
    },
    # ── Page Design Tools ────────────────────────────────────────────────
    {
        "name": "set_different_first_page",
        "fn": page_design_tools.set_different_first_page,
        "annotations": {"title": "Set Different First Page", "destructiveHint": True},
    },
    {
        "name": "set_odd_even_headers",
        "fn": page_design_tools.set_odd_even_headers,
        "annotations": {"title": "Set Odd/Even Headers", "destructiveHint": True},
    },
    {
        "name": "set_tab_stops",
        "fn": page_design_tools.set_tab_stops,
        "annotations": {"title": "Set Tab Stops", "destructiveHint": True},
    },
    {
        "name": "clear_tab_stops",
        "fn": page_design_tools.clear_tab_stops,
        "annotations": {"title": "Clear Tab Stops", "destructiveHint": True},
    },
    {
        "name": "insert_drop_cap",
        "fn": page_design_tools.insert_drop_cap,
        "annotations": {"title": "Insert Drop Cap", "destructiveHint": True},
    },
    {
        "name": "set_page_borders",
        "fn": page_design_tools.set_page_borders,
        "annotations": {"title": "Set Page Borders", "destructiveHint": True},
    },
    # ── Live Export Tools ────────────────────────────────────────────────
    {
        "name": "word_live_export_to_html",
        "fn": live_doc_tools.word_live_export_to_html,
        "annotations": {"title": "Live Export to HTML", "destructiveHint": True},
    },
    {
        "name": "word_live_export_to_rtf",
        "fn": live_doc_tools.word_live_export_to_rtf,
        "annotations": {"title": "Live Export to RTF", "destructiveHint": True},
    },
    {
        "name": "word_live_export_to_txt",
        "fn": live_doc_tools.word_live_export_to_txt,
        "annotations": {"title": "Live Export to Text", "destructiveHint": True},
    },
    # ── Live Merge Tools ─────────────────────────────────────────────────
    {
        "name": "word_live_compare_documents",
        "fn": live_doc_tools.word_live_compare_documents,
        "annotations": {"title": "Live Compare Documents", "destructiveHint": True},
    },
    {
        "name": "word_live_mail_merge",
        "fn": live_doc_tools.word_live_mail_merge,
        "annotations": {"title": "Live Mail Merge", "destructiveHint": True},
    },
    # ── Live Highlight Tools ─────────────────────────────────────────────
    {
        "name": "word_live_highlight_text",
        "fn": live_doc_tools.word_live_highlight_text,
        "annotations": {"title": "Live Highlight Text", "destructiveHint": True},
    },
    # ── Live Spell/Grammar Check ─────────────────────────────────────────
    {
        "name": "word_live_spell_check",
        "fn": live_doc_tools.word_live_spell_check,
        "annotations": {"title": "Live Spell Check", "readOnlyHint": True},
    },
    {
        "name": "word_live_check_grammar",
        "fn": live_doc_tools.word_live_check_grammar,
        "annotations": {"title": "Live Grammar Check", "readOnlyHint": True},
    },
    # ── Live Field Tools ─────────────────────────────────────────────────
    {
        "name": "word_live_insert_field",
        "fn": live_doc_tools.word_live_insert_field,
        "annotations": {"title": "Live Insert Field", "destructiveHint": True},
    },
    {
        "name": "word_live_update_fields",
        "fn": live_doc_tools.word_live_update_fields,
        "annotations": {"title": "Live Update Fields", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_content_control",
        "fn": live_doc_tools.word_live_insert_content_control,
        "annotations": {"title": "Live Insert Content Control", "destructiveHint": True},
    },
    # ── Live Property Tools ──────────────────────────────────────────────
    {
        "name": "word_live_get_custom_properties",
        "fn": live_doc_tools.word_live_get_custom_properties,
        "annotations": {"title": "Live Get Custom Properties", "readOnlyHint": True},
    },
    {
        "name": "word_live_set_custom_property",
        "fn": live_doc_tools.word_live_set_custom_property,
        "annotations": {"title": "Live Set Custom Property", "destructiveHint": True},
    },
    # ── Live Page Design Tools ───────────────────────────────────────────
    {
        "name": "word_live_set_different_first_page",
        "fn": live_doc_tools.word_live_set_different_first_page,
        "annotations": {"title": "Live Set Different First Page", "destructiveHint": True},
    },
    {
        "name": "word_live_set_odd_even_headers",
        "fn": live_doc_tools.word_live_set_odd_even_headers,
        "annotations": {"title": "Live Set Odd/Even Headers", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_drop_cap",
        "fn": live_doc_tools.word_live_insert_drop_cap,
        "annotations": {"title": "Live Insert Drop Cap", "destructiveHint": True},
    },
    {
        "name": "word_live_set_page_borders",
        "fn": live_doc_tools.word_live_set_page_borders,
        "annotations": {"title": "Live Set Page Borders", "destructiveHint": True},
    },
    # ── Live Extra Tools ─────────────────────────────────────────────────
    {
        "name": "word_live_insert_table_of_figures",
        "fn": live_doc_tools.word_live_insert_table_of_figures,
        "annotations": {"title": "Live Insert Table of Figures", "destructiveHint": True},
    },
    {
        "name": "word_live_insert_shape",
        "fn": live_doc_tools.word_live_insert_shape,
        "annotations": {"title": "Live Insert Shape", "destructiveHint": True},
    },
    {
        "name": "word_live_run_macro",
        "fn": live_doc_tools.word_live_run_macro,
        "annotations": {"title": "Live Run Macro", "destructiveHint": True},
    },
    {
        "name": "word_live_print",
        "fn": live_doc_tools.word_live_print,
        "annotations": {"title": "Live Print Document", "destructiveHint": True},
    },
    {
        "name": "word_live_get_doc_variable",
        "fn": live_doc_tools.word_live_get_doc_variable,
        "annotations": {"title": "Live Get Doc Variable", "readOnlyHint": True},
    },
    {
        "name": "word_live_set_doc_variable",
        "fn": live_doc_tools.word_live_set_doc_variable,
        "annotations": {"title": "Live Set Doc Variable", "destructiveHint": True},
    },
]
