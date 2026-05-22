"""Advanced table tools for Word Document Server."""
import json
import os
from typing import Optional

from docx import Document

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def delete_table(filename: str, table_index: int) -> str:
    """Delete a table from a Word document.

    Args:
        filename: Path to the Word document.
        table_index: Index of the table to delete (0-based).

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
            if table_index < 0 or table_index >= len(doc.tables):
                return json.dumps({"success": False, "error": f"Table index {table_index} out of range (0-{len(doc.tables)-1})"})
            table = doc.tables[table_index]
            table._element.getparent().remove(table._element)
            doc.save(filename)
            return json.dumps({"success": True, "deleted_table": table_index})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to delete table: {str(e)}"})


async def repeat_table_header(filename: str, table_index: int, header_rows: int = 1) -> str:
    """Set header rows of a table to repeat on each page.

    Args:
        filename: Path to the Word document.
        table_index: Index of the table (0-based).
        header_rows: Number of header rows to repeat. Default 1.

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
            if table_index < 0 or table_index >= len(doc.tables):
                return json.dumps({"success": False, "error": f"Table index {table_index} out of range"})
            table = doc.tables[table_index]
            if header_rows < 1 or header_rows > len(table.rows):
                return json.dumps({"success": False, "error": f"header_rows {header_rows} out of range"})

            from docx.oxml.ns import qn
            tbl_pr = table._tbl.find(qn('w:tblPr'))
            if tbl_pr is None:
                from lxml import etree
                tbl_pr = etree.SubElement(table._tbl, qn('w:tblPr'))

            from lxml import etree
            tbl_header = tbl_pr.find(qn('w:tblHeader'))
            if tbl_header is None:
                tbl_header = etree.SubElement(tbl_pr, qn('w:tblHeader'))

            for i, row in enumerate(table.rows[:header_rows]):
                tr_pr = row._tr.get_or_add_trPr()
                cntrl = tr_pr.find(qn('w:cntrl'))
                if cntrl is None:
                    cntrl = etree.SubElement(tr_pr, qn('w:cntrl'))

            doc.save(filename)
            return json.dumps({"success": True, "table": table_index, "header_rows": header_rows})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to set repeat header: {str(e)}"})


async def convert_table_to_text(filename: str, table_index: int, separator: str = "tab") -> str:
    """Convert a table to plain text with the specified separator.

    Args:
        filename: Path to the Word document.
        table_index: Index of the table (0-based).
        separator: Separator between cells — "tab", "comma", "pipe", or "paragraph". Default "tab".

    Returns:
        JSON with the converted text.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
    try:
        doc = Document(filename)
        if table_index < 0 or table_index >= len(doc.tables):
            return json.dumps({"success": False, "error": f"Table index {table_index} out of range"})

        sep_map = {"tab": "\t", "comma": ",", "pipe": "|", "paragraph": "\n"}
        sep = sep_map.get(separator.lower(), "\t")

        table = doc.tables[table_index]
        rows_text = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows_text.append(sep.join(cells))

        result = "\n".join(rows_text)
        return json.dumps({"success": True, "table_index": table_index, "separator": separator, "text": result, "rows": len(table.rows)})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to convert table to text: {str(e)}"})


async def convert_text_to_table(filename: str, text: str, columns: int = 2, separator: str = "tab", style: Optional[str] = None) -> str:
    """Convert text into a table. Each line becomes a row, each separator-delimited field a cell.

    Args:
        filename: Path to the Word document.
        text: Text to convert. Lines separated by \\n, cells by separator.
        columns: Number of columns. Default 2.
        separator: Field separator — "tab", "comma", "pipe". Default "tab".
        style: Optional table style name.

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

            sep_map = {"tab": "\t", "comma": ",", "pipe": "|"}
            sep = sep_map.get(separator.lower(), "\t")

            lines = [ln for ln in text.split("\n") if ln.strip()]
            rows_data = []
            for line in lines:
                cells = [c.strip() for c in line.split(sep)]
                while len(cells) < columns:
                    cells.append("")
                rows_data.append(cells[:columns])

            if not rows_data:
                return json.dumps({"success": False, "error": "No data rows found in text"})

            table = doc.add_table(rows=len(rows_data), cols=columns)
            if style:
                try:
                    table.style = doc.styles[style]
                except KeyError:
                    try:
                        table.style = doc.styles['Table Grid']
                    except KeyError:
                        pass
            else:
                try:
                    table.style = doc.styles['Table Grid']
                except KeyError:
                    pass

            for i, row_data in enumerate(rows_data):
                for j, cell_text in enumerate(row_data):
                    table.rows[i].cells[j].text = cell_text

            doc.save(filename)
            return json.dumps({"success": True, "rows": len(rows_data), "columns": columns, "separator": separator})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to convert text to table: {str(e)}"})


async def sort_table(filename: str, table_index: int = 0, column: int = 0, descending: bool = False, header_row: bool = True) -> str:
    """Sort a table by a column. Requires Microsoft Word (COM only, Windows/macOS).

    Args:
        filename: Path to the Word document.
        table_index: Index of the table (0-based). Default 0.
        column: Column index to sort by (0-based). Default 0.
        descending: Sort descending. Default False (ascending).
        header_row: First row is a header (excluded from sort). Default True.

    Returns:
        JSON with result.
    """
    import sys
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import get_word_app, find_document
            app = get_word_app()
            doc = find_document(app, filename)
            if table_index < 0 or table_index >= doc.Tables.Count:
                return json.dumps({"success": False, "error": f"Table index {table_index} out of range"})
            tbl = doc.Tables(table_index + 1)
            sort_order = 2 if descending else 1
            hdr = 1 if header_row else 0
            tbl.Sort(ExcludeHeader=hdr, SortFieldType=0, SortOrder=sort_order, Column=column + 1)
            return json.dumps({"success": True, "table": table_index, "sorted_by_column": column, "descending": descending})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to sort table: {str(e)}"})

    import platform
    if platform.system() == "Darwin":
        try:
            from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
            finder = _doc_finder_js(filename)
            sort_dir = "descending" if descending else "ascending"
            hdr_js = "true" if header_row else "false"
            return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
if ({table_index} >= d.tables.length) throw new Error("Table index out of range");
var tbl = d.tables[{table_index}];
app.sort(tbl, {{by: {column}, order: "{sort_dir}", excludeHeader: {hdr_js}}});
JSON.stringify({{success: true, table: {table_index}, sorted_by_column: {column}}});
""")
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to sort table: {str(e)}"})

    return json.dumps({"success": False, "error": "Sort requires Microsoft Word (Windows or macOS)"})
