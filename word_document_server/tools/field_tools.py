"""Field and content control tools for Word Document Server."""
import json
import os
import sys
from typing import Optional

from docx import Document
from docx.oxml.ns import qn

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def insert_field(filename: str, field_type: str = "DATE", field_text: Optional[str] = None, paragraph_index: Optional[int] = None) -> str:
    """Insert a field code into a Word document (e.g. DATE, PAGE, NUMPAGES, TIME, AUTHOR).

    Args:
        filename: Path to the Word document.
        field_type: Field type — DATE, PAGE, NUMPAGES, TIME, AUTHOR, FILENAME, CREATEDATE, SAVEDATE, REVNUM. Default "DATE".
        field_text: Optional display text for the field.
        paragraph_index: Paragraph index to insert at (0-based). None = append at end.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    valid_fields = {"DATE", "PAGE", "NUMPAGES", "TIME", "AUTHOR", "FILENAME", "CREATEDATE", "SAVEDATE", "REVNUM", "HYPERLINK", "REF", "SEQ", "IF", "MERGEFIELD"}
    if field_type.upper() not in valid_fields:
        return json.dumps({"success": False, "error": f"Unknown field '{field_type}'. Use: {', '.join(sorted(valid_fields))}"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            from lxml import etree

            if paragraph_index is not None:
                if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                    return json.dumps({"success": False, "error": f"Paragraph index {paragraph_index} out of range"})
                para = doc.paragraphs[paragraph_index]
            else:
                para = doc.add_paragraph()

            run = para.add_run()
            fld_char_begin = etree.SubElement(run._element, qn('w:fldChar'))
            fld_char_begin.set(qn('w:fldCharType'), 'begin')

            instr_run = para.add_run()
            instr_text = etree.SubElement(instr_run._element, qn('w:instrText'))
            instr_text.set(qn('xml:space'), 'preserve')
            instr_text.text = f" {field_type.upper()} "

            if field_text:
                display_run = para.add_run()
                fld_char_sep = etree.SubElement(display_run._element, qn('w:fldChar'))
                fld_char_sep.set(qn('w:fldCharType'), 'separate')
                para.add_run(field_text)

            end_run = para.add_run()
            fld_char_end = etree.SubElement(end_run._element, qn('w:fldChar'))
            fld_char_end.set(qn('w:fldCharType'), 'end')

            doc.save(filename)
            return json.dumps({"success": True, "field": field_type, "paragraph": paragraph_index})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to insert field: {str(e)}"})


async def update_fields(filename: str) -> str:
    """Update all fields in a document (TOC, PAGE numbers, etc). Requires Microsoft Word (COM/JXA).

    Args:
        filename: Path to the Word document.

    Returns:
        JSON with result.
    """
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import comtypes_word_app
            with comtypes_word_app() as (word, _):
                doc = word.Documents.Open(os.path.abspath(filename))
                doc.Fields.Update()
                if doc.TablesOfContents.Count > 0:
                    doc.TablesOfContents(1).Update()
                doc.Save()
                doc.Close()
            return json.dumps({"success": True, "action": "update_all_fields"})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to update fields: {str(e)}"})

    return json.dumps({"success": False, "error": "Update fields requires Microsoft Word (Windows)"})


async def insert_content_control(filename: str, control_type: str = "text", title: Optional[str] = None, placeholder: Optional[str] = None, tag: Optional[str] = None, paragraph_index: Optional[int] = None) -> str:
    """Insert a content control into a Word document (text, dropdown, date, checkbox).

    Args:
        filename: Path to the Word document.
        control_type: Control type — "text", "dropdown", "date", "checkbox", "rich_text". Default "text".
        title: Optional control title.
        placeholder: Optional placeholder text.
        tag: Optional tag for identification.
        paragraph_index: Paragraph index (0-based). None = append.

    Returns:
        JSON with result.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})

    type_map = {"text": 0, "rich_text": 0, "dropdown": 2, "date": 6, "checkbox": 8}
    sdt_type = type_map.get(control_type.lower())
    if sdt_type is None:
        return json.dumps({"success": False, "error": f"Unknown control type '{control_type}'. Use: text, dropdown, date, checkbox, rich_text"})

    try:
        async with get_file_lock(filename):
            doc = Document(filename)
            from lxml import etree

            if paragraph_index is not None:
                if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                    return json.dumps({"success": False, "error": f"Paragraph index {paragraph_index} out of range"})

            sdt = etree.SubElement(doc.element.body, qn('w:sdt'))
            sdt_pr = etree.SubElement(sdt, qn('w:sdtPr'))

            doc_obj = etree.SubElement(sdt_pr, qn('w:docObj'))
            doc_obj.set(qn('w:val'), str(sdt_type))

            if title:
                alias = etree.SubElement(sdt_pr, qn('w:alias'))
                alias.set(qn('w:val'), title)
            if tag:
                tag_elem = etree.SubElement(sdt_pr, qn('w:tag'))
                tag_elem.set(qn('w:val'), tag)

            if placeholder:
                ph_elem = etree.SubElement(sdt_pr, qn('w:placeholder'))
                ph_doc = etree.SubElement(ph_elem, qn('w:docPartUnique'))
                ph_doc.set(qn('w:val'), placeholder)

            if control_type.lower() == "date":
                date_elem = etree.SubElement(sdt_pr, qn('w:date'))
                date_elem.set(qn('w:fullDate'), '')
                fmt = etree.SubElement(date_elem, qn('w:dateFormat'))
                fmt.set(qn('w:val'), 'yyyy-MM-dd')

            sdt_content = etree.SubElement(sdt, qn('w:sdtContent'))
            p = etree.SubElement(sdt_content, qn('w:p'))
            r = etree.SubElement(p, qn('w:r'))
            t = etree.SubElement(r, qn('w:t'))
            t.text = placeholder or title or control_type

            doc.save(filename)
            return json.dumps({"success": True, "control_type": control_type, "title": title, "tag": tag})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to insert content control: {str(e)}"})
