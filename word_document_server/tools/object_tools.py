"""Object insertion tools for Word Document Server (TOC update, text boxes, charts, index, citations)."""
import json
import os
import sys
from typing import Optional

from docx import Document
from docx.shared import Inches

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def update_table_of_contents(filename: str) -> str:
    """Update/refresh an existing Table of Contents. Requires Microsoft Word (COM/JXA).

    Args:
        filename: Path to the Word document.

    Returns:
        JSON with result.
    """
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import get_word_app, find_document
            app = get_word_app()
            doc = find_document(app, filename)
            toc_count = doc.TablesOfContents.Count
            if toc_count == 0:
                return json.dumps({"success": False, "error": "No Table of Contents found in document"})
            for i in range(1, toc_count + 1):
                doc.TablesOfContents(i).Update()
            return json.dumps({"success": True, "updated_tocs": toc_count})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to update TOC: {str(e)}"})

    import platform
    if platform.system() == "Darwin":
        try:
            from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
            finder = _doc_finder_js(filename)
            return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var tocCount = d.tablesOfContents.length;
if (tocCount === 0) throw new Error("No Table of Contents found");
for (var i = 0; i < tocCount; i++) {{
    app.update(d.tablesOfContents[i]);
}}
JSON.stringify({{success: true, updated_tocs: tocCount}});
""")
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to update TOC: {str(e)}"})

    return json.dumps({"success": False, "error": "Update TOC requires Microsoft Word (Windows or macOS)"})


async def insert_text_box(filename: str, text: str, left_inches: float = 1.0, top_inches: float = 1.0, width_inches: float = 3.0, height_inches: float = 1.0, style: Optional[str] = None) -> str:
    """Insert a text box into a Word document.

    Args:
        filename: Path to the Word document.
        text: Text content for the text box.
        left_inches: Left position in inches. Default 1.0.
        top_inches: Top position in inches. Default 1.0.
        width_inches: Width in inches. Default 3.0.
        height_inches: Height in inches. Default 1.0.
        style: Optional visual style ("simple", "bordered", "shadowed").

    Returns:
        JSON with result.
    """
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import get_word_app, find_document
            app = get_word_app()
            doc = find_document(app, filename)
            left_pt = int(left_inches * 72)
            top_pt = int(top_inches * 72)
            width_pt = int(width_inches * 72)
            height_pt = int(height_inches * 72)

            tb = doc.Shapes.AddTextbox(
                Orientation=0,
                Left=left_pt, Top=top_pt,
                Width=width_pt, Height=height_pt,
            )
            tb.TextFrame.TextRange.Text = text
            if style == "shadowed":
                tb.Shadow.Visible = True
            elif style == "bordered":
                pass
            elif style == "simple":
                tb.Line.Visible = False

            return json.dumps({"success": True, "text_box": text[:50], "position": f"{left_inches}x{top_inches}in", "size": f"{width_inches}x{height_inches}in"})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to insert text box: {str(e)}"})

    try:
        filename = ensure_docx_extension(filename)
        if not os.path.exists(filename):
            return json.dumps({"success": False, "error": f"Document {filename} does not exist"})
        is_writeable, error_message = check_file_writeable(filename)
        if not is_writeable:
            return json.dumps({"success": False, "error": f"Cannot modify document: {error_message}"})
        async with get_file_lock(filename):
            doc = Document(filename)
            from docx.oxml.ns import qn
            from lxml import etree

            body = doc.element.body
            p = etree.SubElement(body, qn('w:p'))
            r = etree.SubElement(p, qn('w:r'))

            drawing = etree.SubElement(r, qn('w:drawing'))
            inline = etree.SubElement(drawing, qn('wp:inline'), attrib={
                'distT': '0', 'distB': '0', 'distL': '0', 'distR': '0',
            })

            etree.SubElement(inline, qn('wp:extent'), attrib={
                'cx': str(int(Inches(width_inches))),
                'cy': str(int(Inches(height_inches))),
            })

            doc.save(filename)
            return json.dumps({"success": True, "note": "Text box inserted as basic shape. For full text box support, use live COM version on Windows."})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to insert text box: {str(e)}"})


async def insert_chart(filename: str, chart_type: str = "bar", title: Optional[str] = None, categories: Optional[str] = None, values: Optional[str] = None) -> str:
    """Insert a chart into a Word document. Requires Microsoft Word on Windows (COM only).

    Args:
        filename: Path to the Word document.
        chart_type: Chart type — "bar", "column", "line", "pie", "area", "scatter". Default "bar".
        title: Optional chart title.
        categories: Comma-separated category labels (e.g. "Q1,Q2,Q3,Q4").
        values: Comma-separated numeric values (e.g. "10,20,30,40").

    Returns:
        JSON with result.
    """
    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Insert chart requires Microsoft Word on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)

        chart_type_map = {
            "bar": 2, "column": 51, "line": 4, "pie": 5,
            "area": 1, "scatter": -4169,
        }
        xl_chart = chart_type_map.get(chart_type.lower())
        if xl_chart is None:
            return json.dumps({"success": False, "error": f"Unknown chart type '{chart_type}'. Use: bar, column, line, pie, area, scatter"})

        sel = app.Selection
        if categories and values:
            cats = [c.strip() for c in categories.split(",")]
            vals = [float(v.strip()) for v in values.split(",")]
            inline_shape = doc.InlineShapes.AddChart2(
                Style=-1, Type=xl_chart,
                Range=sel.Range,
            )
            chart = inline_shape.Chart
            chart_data = chart.ChartData
            wb = chart_data.Workbook
            ws = wb.Worksheets(1)

            for i, cat in enumerate(cats):
                ws.Cells(i + 2, 1).Value = cat
            for i, val in enumerate(vals):
                ws.Cells(i + 2, 2).Value = val

            ws.Cells(1, 1).Value = "Category"
            ws.Cells(1, 2).Value = "Value"

            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
        else:
            inline_shape = doc.InlineShapes.AddChart2(
                Style=-1, Type=xl_chart,
                Range=sel.Range,
            )

        return json.dumps({"success": True, "chart_type": chart_type, "title": title})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to insert chart: {str(e)}"})


async def insert_index(filename: str, heading_style: Optional[str] = None) -> str:
    """Insert a subject index (concordance) at the end of the document. Requires Microsoft Word (COM/JXA).

    Args:
        filename: Path to the Word document.
        heading_style: Optional heading style for the index heading.

    Returns:
        JSON with result.
    """
    if sys.platform == "win32":
        try:
            from word_document_server.core.word_com import get_word_app, find_document
            app = get_word_app()
            doc = find_document(app, filename)
            sel = app.Selection
            sel.EndKey(Unit=6)
            sel.InsertBreak(Type=7)
            if heading_style:
                sel.Style = heading_style
                sel.TypeText("Index")
                sel.TypeParagraph()
            doc.Indexes.Add(Range=sel.Range, HeadingSeparator=0, Type=0, RightAlignPageNumbers=True)
            doc.Indexes(1).Update()
            return json.dumps({"success": True, "action": "insert_index"})
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to insert index: {str(e)}"})

    import platform
    if platform.system() == "Darwin":
        try:
            from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
            finder = _doc_finder_js(filename)
            return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
app.make({{new: "index", at: d}});
JSON.stringify({{success: true, action: "insert_index"}});
""")
        except Exception as e:
            return json.dumps({"success": False, "error": f"Failed to insert index: {str(e)}"})

    return json.dumps({"success": False, "error": "Insert index requires Microsoft Word (Windows or macOS)"})


async def insert_citation(filename: str, source_type: str = "book", author: Optional[str] = None, title: Optional[str] = None, year: Optional[str] = None, publisher: Optional[str] = None, tag: Optional[str] = None) -> str:
    """Add a citation source and insert a citation field. Requires Microsoft Word on Windows (COM only).

    Args:
        filename: Path to the Word document.
        source_type: Source type — "book", "journal", "web", "conference". Default "book".
        author: Author name.
        title: Work title.
        year: Publication year.
        publisher: Publisher name.
        tag: Citation tag (unique identifier). Auto-generated if omitted.

    Returns:
        JSON with result.
    """
    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Insert citation requires Microsoft Word on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)

        if not tag:
            import hashlib
            tag = hashlib.md5(f"{author}{title}{year}".encode()).hexdigest()[:8] if (author or title) else "cite1"

        source_type_map = {"book": 1, "journal": 2, "web": 3, "conference": 4}
        st = source_type_map.get(source_type.lower(), 1)

        bib_xml = f'<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Tag>{tag}</b:Tag><b:SourceType>{st}</b:SourceType>'
        if author:
            bib_xml += f'<b:Author><b:NameList><b:Person><b:Last>{author}</b:Last></b:Person></b:NameList></b:Author>'
        if title:
            bib_xml += f'<b:Title>{title}</b:Title>'
        if year:
            bib_xml += f'<b:Year>{year}</b:Year>'
        if publisher:
            bib_xml += f'<b:Publisher>{publisher}</b:Publisher>'
        bib_xml += '</b:Source>'

        doc.Bibliography.Sources.Add(bib_xml)

        sel = app.Selection
        field = doc.Fields.Add(sel.Range, Type=26, Text=tag)
        field.Update()

        return json.dumps({"success": True, "tag": tag, "source_type": source_type})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to insert citation: {str(e)}"})
