"""
Hyperlink tools for Word Document Server.

These tools provide MCP interfaces for managing hyperlinks in Word documents.
"""

import json
import os
from typing import Optional

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock
from word_document_server.core.hyperlink_writer import add_hyperlink_to_doc, insert_hyperlink_to_doc


async def manage_hyperlinks(
    filename: str,
    action: str = "add",
    text: str = "",
    url: str = "",
    paragraph_index: Optional[int] = None,
) -> str:
    """Add or manage hyperlinks in a Word document.

    Args:
        filename: Path to Word document
        action: Action to perform:
            - "add": Find existing text and convert it to a hyperlink
            - "insert": Insert new text as a hyperlink at the specified paragraph (or end of document)
        text: Text to convert to a hyperlink (for "add"), or link display text (for "insert")
        url: URL the hyperlink should point to
        paragraph_index: If specified, only search in / insert at this paragraph (0-based)

    Returns:
        JSON string with result
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    if not url:
        return json.dumps({"success": False, "error": "url cannot be empty"})

    if action == "add":
        if not text:
            return json.dumps({"success": False, "error": "text cannot be empty for 'add' action"})

        try:
            async with get_file_lock(filename):
                result = add_hyperlink_to_doc(filename, text, url, paragraph_index)
            return json.dumps(result, ensure_ascii=False, indent=2)
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to add hyperlink: {str(e)}"})
    elif action == "insert":
        display_text = text if text else url

        try:
            async with get_file_lock(filename):
                result = insert_hyperlink_to_doc(filename, display_text, url, paragraph_index)
            return json.dumps(result, ensure_ascii=False, indent=2)
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to insert hyperlink: {str(e)}"})
    else:
        return json.dumps({"success": False, "error": f"Unknown action: {action}. Supported: 'add', 'insert'"})
