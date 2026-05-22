"""Column and section layout tools for Word Document Server."""
import json
import os
from typing import Optional

from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def set_section_columns(filename: str, section_index: int = 0, columns: int = 2, spacing_inches: Optional[float] = None, equal_columns: bool = True) -> str:
    """Set the number of columns for a document section.

    Args:
        filename: Path to the Word document.
        section_index: Section index (0-based). Default 0.
        columns: Number of columns (1-4). Default 2.
        spacing_inches: Space between columns in inches. Default 0.5.
        equal_columns: Whether columns have equal width. Default True.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})
    if columns < 1 or columns > 4:
        return json.dumps({"success": False, "error": "Columns must be 1-4"})
    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if section_index < 0 or section_index >= len(doc.sections):
                return json.dumps({"success": False, "error": f"Section {section_index} out of range"})
            section = doc.sections[section_index]

            from docx.oxml.ns import qn
            sect_pr = section._sectPr

            cols_elem = sect_pr.find(qn('w:cols'))
            if cols_elem is None:
                from lxml import etree
                cols_elem = etree.SubElement(sect_pr, qn('w:cols'))
                sect_pr.insert(list(sect_pr).index(sect_pr.find(qn('w:pgSz'))), cols_elem)

            spacing_emu = Inches(spacing_inches or 0.5)

            cols_elem.set(qn('w:num'), str(columns))
            cols_elem.set(qn('w:space'), str(int(spacing_emu)))
            if not equal_columns and columns > 1:
                cols_elem.set(qn('w:equalWidth'), '0')
            else:
                cols_elem.set(qn('w:equalWidth'), '1')

            doc.save(filename)
            return json.dumps({"success": True, "section": section_index, "columns": columns, "spacing_inches": spacing_inches or 0.5, "equal": equal_columns})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to set columns: {str(e)}"})


async def insert_column_break(filename: str) -> str:
    """Insert a column break at the end of the document.

    Args:
        filename: Path to the Word document.

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
            from docx.oxml.ns import qn
            from lxml import etree
            p = doc.add_paragraph()
            pPr = p._element.get_or_add_pPr()
            br_elem = etree.SubElement(pPr, qn('w:br'))
            br_elem.set(qn('w:type'), 'column')
            doc.save(filename)
            return json.dumps({"success": True, "action": "column_break"})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to insert column break: {str(e)}"})


async def get_section_layout(filename: str, section_index: int = 0) -> str:
    """Get layout properties of a document section (columns, margins, orientation, size).

    Args:
        filename: Path to the Word document.
        section_index: Section index (0-based). Default 0.

    Returns:
        JSON with section layout info.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    try:
        doc = Document(filename)
        if section_index < 0 or section_index >= len(doc.sections):
            return json.dumps({"success": False, "error": f"Section {section_index} out of range"})
        section = doc.sections[section_index]

        from docx.oxml.ns import qn
        cols_elem = section._sectPr.find(qn('w:cols'))
        col_count = int(cols_elem.get(qn('w:num'), '1')) if cols_elem is not None else 1
        col_space = int(cols_elem.get(qn('w:space'), '720')) if cols_elem is not None else 720
        col_equal = cols_elem.get(qn('w:equalWidth'), '1') != '0' if cols_elem is not None else True

        info = {
            "section_index": section_index,
            "orientation": "landscape" if section.orientation == WD_ORIENT.LANDSCAPE else "portrait",            "page_width_inches": round(section.page_width / 914400, 2),
            "page_height_inches": round(section.page_height / 914400, 2),
            "margin_top_inches": round(section.top_margin / 914400, 2),
            "margin_bottom_inches": round(section.bottom_margin / 914400, 2),
            "margin_left_inches": round(section.left_margin / 914400, 2),
            "margin_right_inches": round(section.right_margin / 914400, 2),
            "columns": col_count,
            "column_spacing_inches": round(col_space / 914400, 2),
            "equal_columns": col_equal,
            "gutter_inches": round(section.gutter / 914400, 2) if hasattr(section, 'gutter') else 0,
        }
        return json.dumps({"success": True, "layout": info}, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to get section layout: {str(e)}"})


async def set_column_widths(filename: str, section_index: int = 0, widths_inches: Optional[str] = None) -> str:
    """Set unequal column widths for a section. Automatically sets equal_columns=False.

    Args:
        filename: Path to the Word document.
        section_index: Section index (0-based). Default 0.
        widths_inches: Comma-separated column widths in inches (e.g. "2,4" for two columns).

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})
    if not widths_inches:
        return json.dumps({"success": False, "error": "widths_inches is required (e.g. '2,4')"})
    try:
        widths = [float(w.strip()) for w in widths_inches.split(",")]
    except ValueError:
        return json.dumps({"success": False, "error": "Invalid widths_inches format. Use comma-separated numbers (e.g. '2,4')"})
    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if section_index < 0 or section_index >= len(doc.sections):
                return json.dumps({"success": False, "error": f"Section {section_index} out of range"})
            section = doc.sections[section_index]

            from docx.oxml.ns import qn
            from lxml import etree
            sect_pr = section._sectPr

            cols_elem = sect_pr.find(qn('w:cols'))
            if cols_elem is None:
                cols_elem = etree.SubElement(sect_pr, qn('w:cols'))

            cols_elem.set(qn('w:num'), str(len(widths)))
            cols_elem.set(qn('w:equalWidth'), '0')
            cols_elem.set(qn('w:space'), str(int(Inches(0.5))))

            for child in list(cols_elem.findall(qn('w:col'))):
                cols_elem.remove(child)

            for i, w in enumerate(widths):
                col_elem = etree.SubElement(cols_elem, qn('w:col'))
                col_elem.set(qn('w:w'), str(int(Inches(w))))
                if i < len(widths) - 1:
                    col_elem.set(qn('w:space'), str(int(Inches(0.5))))

            doc.save(filename)
            return json.dumps({"success": True, "section": section_index, "column_widths": widths})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to set column widths: {str(e)}"})
