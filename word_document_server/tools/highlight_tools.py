"""Text highlighting tools for Word Document Server."""
import json
import os
from typing import Optional

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock

_HIGHLIGHT_MAP = {
    "yellow": WD_COLOR_INDEX.YELLOW,
    "green": WD_COLOR_INDEX.GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE,
    "magenta": WD_COLOR_INDEX.PINK,
    "blue": WD_COLOR_INDEX.BLUE,
    "red": WD_COLOR_INDEX.RED,
    "darkblue": WD_COLOR_INDEX.DARK_BLUE,
    "darkcyan": WD_COLOR_INDEX.TEAL,
    "darkgreen": WD_COLOR_INDEX.BRIGHT_GREEN,
    "darkmagenta": WD_COLOR_INDEX.VIOLET,
    "darkred": WD_COLOR_INDEX.DARK_RED,
    "darkyellow": WD_COLOR_INDEX.DARK_YELLOW,
    "gray": WD_COLOR_INDEX.GRAY_25,
    "gray50": WD_COLOR_INDEX.GRAY_50,
    "none": WD_COLOR_INDEX.AUTO,
}


async def highlight_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int, color: str = "yellow") -> str:
    """Apply highlighting to a text range within a paragraph.

    Args:
        filename: Path to the Word document.
        paragraph_index: Index of the paragraph (0-based).
        start_pos: Start character position.
        end_pos: End character position.
        color: Highlight color — yellow, green, cyan, magenta, blue, red, darkblue, darkcyan, darkgreen, darkmagenta, darkred, darkyellow, gray, gray50. Default "yellow".

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    hl_color = _HIGHLIGHT_MAP.get(color.lower())
    if hl_color is None:
        return json.dumps({"success": False, "error": f"Unknown color '{color}'. Use: {', '.join(_HIGHLIGHT_MAP.keys())}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return json.dumps({"success": False, "error": f"Paragraph index {paragraph_index} out of range"})
            para = doc.paragraphs[paragraph_index]
            full_text = para.text
            if start_pos < 0 or end_pos > len(full_text) or start_pos >= end_pos:
                return json.dumps({"success": False, "error": f"Invalid range {start_pos}-{end_pos} for paragraph length {len(full_text)}"})

            target_text = full_text[start_pos:end_pos]
            before_text = full_text[:start_pos]
            after_text = full_text[end_pos:]

            for run in para.runs:
                run.text = ""

            if before_text:
                para.add_run(before_text)
            hl_run = para.add_run(target_text)
            hl_run.font.highlight_color = hl_color
            if after_text:
                para.add_run(after_text)

            doc.save(filename)
            return json.dumps({"success": True, "highlighted": target_text[:50], "color": color, "paragraph": paragraph_index})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to highlight: {str(e)}"})


async def remove_highlight(filename: str, paragraph_index: Optional[int] = None) -> str:
    """Remove all highlighting from a paragraph or the entire document.

    Args:
        filename: Path to the Word document.
        paragraph_index: Paragraph index (0-based). None = remove from all paragraphs.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            count = 0
            paragraphs = [doc.paragraphs[paragraph_index]] if paragraph_index is not None else doc.paragraphs
            for para in paragraphs:
                for run in para.runs:
                    if run.font.highlight_color is not None and run.font.highlight_color != WD_COLOR_INDEX.AUTO:
                        run.font.highlight_color = WD_COLOR_INDEX.AUTO
                        count += 1

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if run.font.highlight_color is not None and run.font.highlight_color != WD_COLOR_INDEX.AUTO:
                                    run.font.highlight_color = WD_COLOR_INDEX.AUTO
                                    count += 1

            doc.save(filename)
            return json.dumps({"success": True, "removed_count": count, "scope": f"paragraph {paragraph_index}" if paragraph_index is not None else "entire document"})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to remove highlight: {str(e)}"})
