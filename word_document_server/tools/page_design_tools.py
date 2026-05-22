"""Page design tools for Word Document Server."""
import json
import os
from typing import Optional

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Inches, Pt

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def set_different_first_page(filename: str, section_index: int = 0, enabled: bool = True) -> str:
    """Set whether the first page of a section has different headers/footers.

    Args:
        filename: Path to the Word document.
        section_index: Section index (0-based). Default 0.
        enabled: True to enable different first page. Default True.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify: {err}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if section_index < 0 or section_index >= len(doc.sections):
                return json.dumps({"success": False, "error": f"Section index {section_index} out of range"})
            section = doc.sections[section_index]
            section.different_first_page_header_footer = enabled
            doc.save(filename)
            return json.dumps({"success": True, "section": section_index, "different_first_page": enabled})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed: {str(e)}"})


async def set_odd_even_headers(filename: str, section_index: int = 0, enabled: bool = True) -> str:
    """Set whether odd and even pages have different headers/footers.

    Args:
        filename: Path to the Word document.
        section_index: Section index (0-based). Default 0.
        enabled: True to enable odd/even headers. Default True.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify: {err}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if section_index < 0 or section_index >= len(doc.sections):
                return json.dumps({"success": False, "error": f"Section index {section_index} out of range"})
            from docx.oxml.ns import qn
            section = doc.sections[section_index]
            sect_pr = section._sectPr
            if enabled:
                sect_pr.set(qn('w:evenAndOddHeaders'), '1') if not sect_pr.get(qn('w:evenAndOddHeaders')) else None
            else:
                if sect_pr.get(qn('w:evenAndOddHeaders')):
                    del sect_pr.attrib[qn('w:evenAndOddHeaders')]
            doc.save(filename)
            return json.dumps({"success": True, "section": section_index, "odd_even_headers": enabled})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed: {str(e)}"})


async def set_tab_stops(filename: str, paragraph_index: int, positions: str, alignments: Optional[str] = None, leaders: Optional[str] = None) -> str:
    """Set tab stops for a paragraph.

    Args:
        filename: Path to the Word document.
        paragraph_index: Paragraph index (0-based).
        positions: Tab positions in inches, comma-separated (e.g. "0.5,1.5,3.0").
        alignments: Tab alignments, comma-separated — "left", "center", "right", "decimal", "bar". Default all "left".
        leaders: Tab leaders, comma-separated — "none", "dots", "dashes", "heavy", "middleDot". Default all "none".

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify: {err}"})

    align_map = {"left": WD_TAB_ALIGNMENT.LEFT, "center": WD_TAB_ALIGNMENT.CENTER, "right": WD_TAB_ALIGNMENT.RIGHT, "decimal": WD_TAB_ALIGNMENT.DECIMAL, "bar": WD_TAB_ALIGNMENT.BAR}
    leader_map = {"none": WD_TAB_LEADER.SPACES, "dots": WD_TAB_LEADER.DOTS, "dashes": WD_TAB_LEADER.DASHES, "heavy": WD_TAB_LEADER.HEAVY, "middledot": WD_TAB_LEADER.MIDDLE_DOT}

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return json.dumps({"success": False, "error": f"Paragraph index {paragraph_index} out of range"})
            para = doc.paragraphs[paragraph_index]
            pos_list = [float(p.strip()) for p in positions.split(",")]
            align_list = [a.strip().lower() for a in alignments.split(",")] if alignments else []
            leader_list = [ld.strip().lower() for ld in leaders.split(",")] if leaders else []

            pf = para.paragraph_format
            pf.tab_stops.clear_all()

            for i, pos in enumerate(pos_list):
                align = align_map.get(align_list[i], WD_TAB_ALIGNMENT.LEFT) if i < len(align_list) else WD_TAB_ALIGNMENT.LEFT
                lead = leader_map.get(leader_list[i], WD_TAB_LEADER.SPACES) if i < len(leader_list) else WD_TAB_LEADER.SPACES
                pf.tab_stops.add_tab_stop(Inches(pos), alignment=align, leader=lead)

            doc.save(filename)
            return json.dumps({"success": True, "paragraph": paragraph_index, "tabs_added": len(pos_list)})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed: {str(e)}"})


async def clear_tab_stops(filename: str, paragraph_index: int) -> str:
    """Clear all tab stops from a paragraph.

    Args:
        filename: Path to the Word document.
        paragraph_index: Paragraph index (0-based).

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify: {err}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return json.dumps({"success": False, "error": f"Paragraph index {paragraph_index} out of range"})
            doc.paragraphs[paragraph_index].paragraph_format.tab_stops.clear_all()
            doc.save(filename)
            return json.dumps({"success": True, "paragraph": paragraph_index, "tabs_cleared": True})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed: {str(e)}"})


async def insert_drop_cap(filename: str, paragraph_index: int, lines: int = 3, font_name: Optional[str] = None, font_size: Optional[int] = None) -> str:
    """Insert a drop cap effect on the first character of a paragraph.

    Args:
        filename: Path to the Word document.
        paragraph_index: Paragraph index (0-based).
        lines: Number of lines the drop cap spans. Default 3.
        font_name: Font name for the drop cap. None = inherit.
        font_size: Font size in points for the drop cap. None = auto based on lines.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify: {err}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return json.dumps({"success": False, "error": f"Paragraph index {paragraph_index} out of range"})
            para = doc.paragraphs[paragraph_index]
            text = para.text
            if not text:
                return json.dumps({"success": False, "error": "Paragraph is empty"})

            first_char = text[0]
            rest = text[1:]

            for run in para.runs:
                run.text = ""

            drop_run = para.add_run(first_char)
            size = font_size or (lines * 12)
            drop_run.font.size = Pt(size)
            if font_name:
                drop_run.font.name = font_name

            if rest:
                para.add_run(rest)

            doc.save(filename)
            return json.dumps({"success": True, "paragraph": paragraph_index, "drop_char": first_char, "lines": lines})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed: {str(e)}"})


async def set_page_borders(filename: str, section_index: int = 0, style: str = "single", color: str = "auto", size: int = 4, offset: float = 24.0, apply_to: str = "all") -> str:
    """Set page borders for a section.

    Args:
        filename: Path to the Word document.
        section_index: Section index (0-based). Default 0.
        style: Border style — "single", "double", "dashed", "dotted", "none". Default "single".
        color: Border color. Default "auto".
        size: Border width in 1/8 points. Default 4 (=0.5pt).
        offset: Border offset from page edge in points. Default 24.
        apply_to: Which pages — "all", "first", "notFirst". Default "all".

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify: {err}"})

    style_map = {"single": "single", "double": "double", "dashed": "dashed", "dotted": "dotted", "none": "none"}
    val = style_map.get(style.lower())
    if val is None:
        return json.dumps({"success": False, "error": f"Unknown style '{style}'. Use: single, double, dashed, dotted, none"})

    apply_map = {"all": "0", "first": "1", "notfirst": "2"}
    page_val = apply_map.get(apply_to.lower())
    if page_val is None:
        return json.dumps({"success": False, "error": f"Unknown apply_to '{apply_to}'. Use: all, first, notFirst"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            if section_index < 0 or section_index >= len(doc.sections):
                return json.dumps({"success": False, "error": f"Section index {section_index} out of range"})

            from docx.oxml.ns import qn
            from lxml import etree

            sect_pr = doc.sections[section_index]._sectPr

            pg_borders = sect_pr.find(qn('w:pgBorders'))
            if pg_borders is not None:
                sect_pr.remove(pg_borders)

            if val == "none":
                doc.save(filename)
                return json.dumps({"success": True, "section": section_index, "borders": "removed"})

            pg_borders = etree.SubElement(sect_pr, qn('w:pgBorders'))
            pg_borders.set(qn('w:offsetFrom'), 'page')
            if apply_to.lower() != "all":
                pg_borders.set(qn('w:display'), page_val)

            for side in ('top', 'left', 'bottom', 'right'):
                border = etree.SubElement(pg_borders, qn(f'w:{side}'))
                border.set(qn('w:val'), val)
                border.set(qn('w:sz'), str(size))
                border.set(qn('w:space'), str(int(offset)))
                border.set(qn('w:color'), color)

            doc.save(filename)
            return json.dumps({"success": True, "section": section_index, "style": style, "color": color, "apply_to": apply_to})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed: {str(e)}"})
