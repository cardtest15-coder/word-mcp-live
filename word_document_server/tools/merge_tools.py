"""Document merge and compare tools for Word Document Server."""
import json
import os
import sys
from typing import Optional

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def compare_documents(original_filename: str, revised_filename: str, destination_filename: Optional[str] = None) -> str:
    """Compare two documents and produce a legal blackline. Requires Microsoft Word (COM on Windows / JXA on macOS).

    Args:
        original_filename: Path to the original document.
        revised_filename: Path to the revised document.
        destination_filename: Path for the comparison result. Defaults to original + '_compared.docx'.

    Returns:
        JSON with result.
    """
    original_filename = ensure_docx_extension(original_filename)
    revised_filename = ensure_docx_extension(revised_filename)
    if not os.path.exists(original_filename):
        return json.dumps({"success": False, "error": f"Original document {original_filename} does not exist"})
    if not os.path.exists(revised_filename):
        return json.dumps({"success": False, "error": f"Revised document {revised_filename} does not exist"})

    if not destination_filename:
        destination_filename = os.path.splitext(original_filename)[0] + "_compared.docx"

    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import comtypes_word_app
            with comtypes_word_app() as (word, _):
                doc = word.Documents.Open(os.path.abspath(original_filename))
                doc.Compare(os.path.abspath(revised_filename))
                doc.SaveAs2(os.path.abspath(destination_filename))
                doc.Close()
            return json.dumps({"success": True, "output": destination_filename})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to compare: {str(e)}"})

    import platform
    if platform.system() == "Darwin":
        try:
            from word_document_server.core.word_mac import _run_applescript
            _run_applescript(f'''
tell application "Microsoft Word"
    open POSIX file "{os.path.abspath(original_filename)}"
    compare active document POSIX file "{os.path.abspath(revised_filename)}"
    save active document in POSIX file "{os.path.abspath(destination_filename)}"
    close active document
end tell
''')
            return json.dumps({"success": True, "output": destination_filename})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to compare: {str(e)}"})

    return json.dumps({"success": False, "error": "Compare requires Microsoft Word (Windows or macOS)"})


async def mail_merge(template_filename: str, data: str, output_filename: Optional[str] = None, separator: str = "comma") -> str:
    """Perform mail merge: replace merge fields in a template with data rows.

    Args:
        template_filename: Path to the template document with merge fields (<<Field>>).
        data: Data rows. First row = headers (field names), subsequent rows = values. Rows separated by \\n, fields by separator.
        output_filename: Path for the merged output. Defaults to template + '_merged.docx'.
        separator: Field separator — "comma", "tab", "semicolon". Default "comma".

    Returns:
        JSON with result.
    """
    template_filename = ensure_docx_extension(template_filename)
    if not os.path.exists(template_filename):
        return json.dumps({"success": False, "error": f"Template {template_filename} does not exist"})

    if not output_filename:
        output_filename = os.path.splitext(template_filename)[0] + "_merged.docx"

    sep_map = {"comma": ",", "tab": "\t", "semicolon": ";"}
    sep = sep_map.get(separator.lower(), ",")

    try:
        lines = [line.strip() for line in data.strip().split("\n") if line.strip()]
        if len(lines) < 2:
            return json.dumps({"success": False, "error": "Need at least a header row and one data row"})
        headers = [h.strip() for h in lines[0].split(sep)]
        rows = []
        for line in lines[1:]:
            vals = [v.strip() for v in line.split(sep)]
            rows.append(dict(zip(headers, vals)))

        from docx import Document as DocxDocument
        import shutil
        shutil.copy2(template_filename, output_filename)

        is_writeable, error_message = check_file_writeable(output_filename)
        if not is_writeable:
            return json.dumps({"success": False, "error": f"Cannot write output: {error_message}"})

        async with get_file_lock(output_filename):
            doc = DocxDocument(output_filename)
            merged_count = 0
            for row in rows:
                for para in doc.paragraphs:
                    for key, val in row.items():
                        placeholder = f"<<{key}>>"
                        if placeholder in para.text:
                            for run in para.runs:
                                if placeholder in run.text:
                                    run.text = run.text.replace(placeholder, val)
                                    merged_count += 1
                for table in doc.tables:
                    for table_row in table.rows:
                        for cell in table_row.cells:
                            for para in cell.paragraphs:
                                for key, val in row.items():
                                    placeholder = f"<<{key}>>"
                                    if placeholder in para.text:
                                        for run in para.runs:
                                            if placeholder in run.text:
                                                run.text = run.text.replace(placeholder, val)
                                                merged_count += 1

            if merged_count == 0:
                return json.dumps({"success": False, "error": "No merge fields found. Use <<FieldName>> syntax."})

            doc.save(output_filename)
            return json.dumps({"success": True, "output": output_filename, "fields": headers, "rows": len(rows), "replacements": merged_count})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to mail merge: {str(e)}"})
