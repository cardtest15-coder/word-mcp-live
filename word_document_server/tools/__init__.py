"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

__all__ = [
    "create_document", "get_document_info", "get_document_text",
    "get_document_outline", "list_available_documents",
    "copy_document", "merge_documents",
    "add_heading", "add_paragraph", "add_table", "add_picture",
    "add_page_break", "add_table_of_contents", "delete_paragraph",
    "search_and_replace",
    "format_text", "create_custom_style", "format_table",
    "protect_document", "add_restricted_editing",
    "add_digital_signature", "verify_document",
    "add_footnote_to_document", "add_endnote_to_document",
    "convert_footnotes_to_endnotes_in_document", "customize_footnote_style",
    "get_all_comments", "get_comments_by_author", "get_comments_for_paragraph",
    "comment_write_tools", "hyperlink_tools", "tracked_changes_tools",
    "layout_tools", "live_tools", "live_read_tools",
    "live_layout_tools", "screen_capture_tools", "extended_document_tools",
    "style_tools", "column_tools", "advanced_table_tools", "object_tools",
    "export_tools", "merge_tools", "highlight_tools", "field_tools",
    "property_tools", "page_design_tools", "live_doc_tools",
]

from word_document_server.tools.document_tools import (
    create_document, get_document_info, get_document_text,
    get_document_outline, list_available_documents,
    copy_document, merge_documents,
)

from word_document_server.tools.content_tools import (
    add_heading, add_paragraph, add_table, add_picture,
    add_page_break, add_table_of_contents, delete_paragraph,
    search_and_replace,
)

from word_document_server.tools.format_tools import (
    format_text, create_custom_style, format_table,
)

from word_document_server.tools.protection_tools import (
    protect_document, add_restricted_editing,
    add_digital_signature, verify_document,
)

from word_document_server.tools.footnote_tools import (
    add_footnote_to_document, add_endnote_to_document,
    convert_footnotes_to_endnotes_in_document, customize_footnote_style,
)

from word_document_server.tools.comment_tools import (
    get_all_comments, get_comments_by_author, get_comments_for_paragraph,
)

from word_document_server.tools import comment_write_tools
from word_document_server.tools import hyperlink_tools
from word_document_server.tools import tracked_changes_tools
from word_document_server.tools import layout_tools
from word_document_server.tools import live_tools
from word_document_server.tools import live_read_tools
from word_document_server.tools import live_layout_tools
from word_document_server.tools import screen_capture_tools
from word_document_server.tools import extended_document_tools
from word_document_server.tools import style_tools
from word_document_server.tools import column_tools
from word_document_server.tools import advanced_table_tools
from word_document_server.tools import object_tools
from word_document_server.tools import export_tools
from word_document_server.tools import merge_tools
from word_document_server.tools import highlight_tools
from word_document_server.tools import field_tools
from word_document_server.tools import property_tools
from word_document_server.tools import page_design_tools
from word_document_server.tools import live_doc_tools
