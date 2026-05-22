"""Live (COM/JXA) tools for document-level operations: export, merge, highlight, fields, properties, page design."""
import json
import os
import sys
from typing import Optional

_MAC_AVAILABLE = sys.platform == 'darwin'
_PTS_PER_INCH = 72.0


async def word_live_export_to_html(filename: Optional[str] = None, output_path: Optional[str] = None) -> str:
    """Export an open Word document to HTML via COM/JXA.

    Args:
        filename: Document name or path (None = active document).
        output_path: Output HTML path. Defaults to document name + .html.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_applescript, _doc_finder_js
        finder = _doc_finder_js(filename)
        out = output_path or "/tmp/export.html"
        return _run_applescript(f'''
tell application "Microsoft Word"
    {finder.split("var d =")[1].split(";")[0].replace("Application('Microsoft Word').documents.byName", "open POSIX file").replace("Application('Microsoft Word').activeDocument", "active document").strip() if "byName" in finder else ""}
    save active document in POSIX file "{out}"
end tell
''')

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if not output_path:
            output_path = os.path.splitext(doc.FullName)[0] + ".html"
        doc.SaveAs2(os.path.abspath(output_path), FileFormat=8)  # wdFormatHTML
        return json.dumps({"success": True, "output": output_path, "format": "html"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_export_to_rtf(filename: Optional[str] = None, output_path: Optional[str] = None) -> str:
    """Export an open Word document to RTF via COM.

    Args:
        filename: Document name or path (None = active document).
        output_path: Output RTF path. Defaults to document name + .rtf.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "RTF export via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if not output_path:
            output_path = os.path.splitext(doc.FullName)[0] + ".rtf"
        doc.SaveAs2(os.path.abspath(output_path), FileFormat=6)  # wdFormatRTF
        return json.dumps({"success": True, "output": output_path, "format": "rtf"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_export_to_txt(filename: Optional[str] = None, output_path: Optional[str] = None) -> str:
    """Export an open Word document to plain text via COM.

    Args:
        filename: Document name or path (None = active document).
        output_path: Output text path. Defaults to document name + .txt.

    Returns:
        JSON with result.
    """
    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if not output_path:
            output_path = os.path.splitext(doc.FullName)[0] + ".txt"
        doc.SaveAs2(os.path.abspath(output_path), FileFormat=2)  # wdFormatText
        return json.dumps({"success": True, "output": output_path, "format": "txt"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_compare_documents(original_filename: str, revised_filename: str, destination_filename: Optional[str] = None) -> str:
    """Compare two open documents and produce a legal blackline via COM.

    Args:
        original_filename: Path to the original document.
        revised_filename: Path to the revised document.
        destination_filename: Path for the comparison result.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_applescript
        if not destination_filename:
            destination_filename = os.path.splitext(original_filename)[0] + "_compared.docx"
        return _run_applescript(f'''
tell application "Microsoft Word"
    open POSIX file "{os.path.abspath(original_filename)}"
    compare active document POSIX file "{os.path.abspath(revised_filename)}"
    save active document in POSIX file "{os.path.abspath(destination_filename)}"
    close active document
end tell
''')

    try:
        from word_document_server.core.word_com import get_word_app
        app = get_word_app()
        doc = app.Documents.Open(os.path.abspath(original_filename))
        doc.Compare(os.path.abspath(revised_filename))
        if not destination_filename:
            destination_filename = os.path.splitext(original_filename)[0] + "_compared.docx"
        doc.SaveAs2(os.path.abspath(destination_filename))
        doc.Close()
        return json.dumps({"success": True, "output": destination_filename})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_mail_merge(template_filename: Optional[str] = None, data_source: Optional[str] = None, output_filename: Optional[str] = None) -> str:
    """Perform mail merge on an open document via COM.

    Args:
        template_filename: Document name or path (None = active document).
        data_source: Path to CSV/data source file for merge.
        output_filename: Path for merged output.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Mail merge via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, template_filename)
        if data_source:
            doc.MailMerge.OpenDataSource(Name=os.path.abspath(data_source))
        if doc.MailMerge.DataSource.RecordCount > 0:
            doc.MailMerge.Destination = 0  # wdSendToNewDocument
            doc.MailMerge.Execute()
            result_doc = app.ActiveDocument
            if output_filename:
                result_doc.SaveAs2(os.path.abspath(output_filename))
            else:
                output_filename = os.path.splitext(doc.FullName)[0] + "_merged.docx"
                result_doc.SaveAs2(os.path.abspath(output_filename))
            result_doc.Close()
            return json.dumps({"success": True, "output": output_filename, "records": doc.MailMerge.DataSource.RecordCount})
        return json.dumps({"success": False, "error": "No merge data source or records found"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_highlight_text(filename: Optional[str] = None, paragraph_index: int = 1, start_pos: int = 1, end_pos: int = 1, color: str = "yellow") -> str:
    """Apply highlighting to text in an open Word document via COM/JXA.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph number (1-based, COM style). Default 1.
        start_pos: Start character position (1-based). Default 1.
        end_pos: End character position (1-based). Default 1.
        color: Highlight color name — yellow, green, cyan, magenta, blue, red. Default "yellow".

    Returns:
        JSON with result.
    """
    _COLOR_MAP = {
        "yellow": 7, "brightgreen": 4, "cyan": 6, "magenta": 5, "blue": 2, "red": 3,
        "darkblue": 9, "darkcyan": 10, "darkgreen": 11, "darkmagenta": 12, "darkred": 13,
        "darkyellow": 14, "gray": 15, "gray50": 16, "none": 0,
    }
    hl_val = _COLOR_MAP.get(color.lower())
    if hl_val is None:
        return json.dumps({"success": False, "error": f"Unknown color '{color}'. Use: {', '.join(_COLOR_MAP.keys())}"})

    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
        var para = d.paragraphs[{paragraph_index - 1}];
var rng = para.characters[{start_pos - 1}].getRange();
rng.expand({{unit: "character", count: {end_pos - start_pos + 1}}});
rng.highlightColorIndex = {hl_val};
JSON.stringify({{success: true, color: "{color}"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Highlight Text"):
            para = doc.Paragraphs(paragraph_index)
            rng = para.Range
            rng.SetRange(start_pos - 1, end_pos)
            rng.HighlightColorIndex = hl_val
        return json.dumps({"success": True, "color": color, "paragraph": paragraph_index})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_spell_check(filename: Optional[str] = None) -> str:
    """Run spell check on an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Spell check via JXA not supported"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        errors = []
        for i in range(1, doc.SpellingErrors.Count + 1):
            err = doc.SpellingErrors(i)
            errors.append({"text": err.Text, "start": err.Start, "end": err.End})
        return json.dumps({"success": True, "error_count": len(errors), "errors": errors[:50]})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_check_grammar(filename: Optional[str] = None) -> str:
    """Run grammar check on an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Grammar check via JXA not supported"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        errors_before = doc.GrammaticalErrors.Count
        app.CheckGrammar()
        errors_after = doc.GrammaticalErrors.Count
        return json.dumps({"success": True, "errors_before": errors_before, "errors_after": errors_after})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_field(filename: Optional[str] = None, field_type: str = "DATE", field_text: Optional[str] = None, paragraph_index: Optional[int] = None) -> str:
    """Insert a field code in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        field_type: Field type — DATE, PAGE, NUMPAGES, TIME, AUTHOR, FILENAME. Default "DATE".
        field_text: Optional display text.
        paragraph_index: Paragraph number (1-based). None = at selection.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
        var rng = d.paragraphs[{paragraph_index - 1}].getRange() if ({paragraph_index}) else app.selection.paragraphs[0].getRange();
d.fields.add(rng, {{type: "{field_type}"}});
JSON.stringify({{success: true, field: "{field_type}"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Insert Field"):
            if paragraph_index is not None:
                rng = doc.Paragraphs(paragraph_index).Range
            else:
                rng = app.Selection.Range
            doc.Fields.Add(rng, -1, f" {field_type} ", True)
        return json.dumps({"success": True, "field": field_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_update_fields(filename: Optional[str] = None) -> str:
    """Update all fields in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
d.fields.update();
if (d.tablesOfContents.length > 0) d.tablesOfContents[0].update();
JSON.stringify({{success: true, action: "update_fields"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        doc.Fields.Update()
        if doc.TablesOfContents.Count > 0:
            doc.TablesOfContents(1).Update()
        return json.dumps({"success": True, "action": "update_fields"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_content_control(filename: Optional[str] = None, control_type: str = "text", title: Optional[str] = None, placeholder: Optional[str] = None, tag: Optional[str] = None) -> str:
    """Insert a content control in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        control_type: Control type — "text", "dropdown", "date", "checkbox". Default "text".
        title: Optional control title.
        placeholder: Optional placeholder text.
        tag: Optional tag.

    Returns:
        JSON with result.
    """
    _TYPE_MAP = {"text": 0, "rich_text": 0, "dropdown": 2, "combobox": 3, "date": 6, "checkbox": 8}
    wd_type = _TYPE_MAP.get(control_type.lower())
    if wd_type is None:
        return json.dumps({"success": False, "error": f"Unknown control type '{control_type}'. Use: text, dropdown, date, checkbox"})

    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Content controls via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Insert Content Control"):
            rng = app.Selection.Range
            cc = doc.ContentControls.Add(wd_type, rng)
            if title:
                cc.Title = title
            if placeholder:
                try:
                    cc.SetPlaceholderText(None, None, placeholder)
                except Exception:
                    pass
            if tag:
                cc.Tag = tag
            if control_type.lower() == "date":
                cc.DateDisplayFormat = "yyyy-MM-dd"
        return json.dumps({"success": True, "control_type": control_type, "title": title})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_get_custom_properties(filename: Optional[str] = None) -> str:
    """Get custom document properties from an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with properties.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Custom properties via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        props = {}
        for i in range(1, doc.BuiltInDocumentProperties.Count + 1):
            try:
                prop = doc.BuiltInDocumentProperties(i)
                props[prop.Name] = str(prop.Value)
            except Exception:
                pass
        for i in range(1, doc.CustomDocumentProperties.Count + 1):
            try:
                prop = doc.CustomDocumentProperties(i)
                props[prop.Name] = str(prop.Value)
            except Exception:
                pass
        return json.dumps({"success": True, "properties": props, "count": len(props)})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_set_custom_property(filename: Optional[str] = None, name: str = "", value: str = "", property_type: str = "text") -> str:
    """Set a custom document property in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        name: Property name.
        value: Property value.
        property_type: Value type — "text", "number", "date", "boolean". Default "text".

    Returns:
        JSON with result.
    """
    _TYPE_MAP = {"text": 2, "number": 1, "date": 3, "boolean": 4}
    _TYPE_MAP.get(property_type.lower(), 2)  # used by fallback

    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Custom properties via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        try:
            prop = doc.CustomDocumentProperties(name)
            prop.Value = value if property_type == "text" else (int(value) if property_type == "number" else (bool(value.lower() == "true") if property_type == "boolean" else value))
        except Exception:
            # CDP.Add is broken in COM automation (known Office bug)
            # Fallback: save, close, use zipfile manipulation, reopen
            fpath = doc.FullName
            doc.Save()
            doc.Close(0)
            import time
            time.sleep(0.5)
            from word_document_server.tools.property_tools import set_custom_property
            result = await set_custom_property(fpath, name, value, property_type)
            # Reopen
            app.Documents.Open(fpath)
            return result
        return json.dumps({"success": True, "property": name, "value": value, "type": property_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_set_different_first_page(filename: Optional[str] = None, section_index: int = 1, enabled: bool = True) -> str:
    """Set different first page header/footer in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-based). Default 1.
        enabled: True to enable different first page. Default True.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
d.sections[{section_index - 1}].differentFirstPageHeaderFooter = {str(enabled).lower()};
JSON.stringify({{success: true, section: {section_index}, different_first_page: {str(enabled).lower()}}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        doc.Sections(section_index).PageSetup.DifferentFirstPageHeaderFooter = enabled
        return json.dumps({"success": True, "section": section_index, "different_first_page": enabled})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_set_odd_even_headers(filename: Optional[str] = None, section_index: int = 1, enabled: bool = True) -> str:
    """Set odd/even page headers/footers in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-based). Default 1.
        enabled: True to enable odd/even headers. Default True.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
d.sections[{section_index - 1}].pageSetup.oddAndEvenPagesHeaderFooter = {str(enabled).lower()};
JSON.stringify({{success: true, section: {section_index}, odd_even: {str(enabled).lower()}}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        doc.Sections(section_index).PageSetup.OddAndEvenPagesHeaderFooter = enabled
        return json.dumps({"success": True, "section": section_index, "odd_even_headers": enabled})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_drop_cap(filename: Optional[str] = None, paragraph_index: int = 1, lines: int = 3, font_name: Optional[str] = None) -> str:
    """Insert a drop cap on the first character of a paragraph in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph number (1-based). Default 1.
        lines: Number of lines the drop cap spans. Default 3.
        font_name: Font name for the drop cap.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Drop cap via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Insert Drop Cap"):
            para = doc.Paragraphs(paragraph_index)
            rng = para.Range
            rng.SetRange(rng.Start, rng.Start + 1)
            rng.Font.Size = lines * 12
            if font_name:
                rng.Font.Name = font_name
        return json.dumps({"success": True, "paragraph": paragraph_index, "lines": lines})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_set_page_borders(filename: Optional[str] = None, section_index: int = 1, style: str = "single", color: str = "auto", size: int = 4, offset: float = 24.0, apply_to: str = "all") -> str:
    """Set page borders in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-based). Default 1.
        style: Border style — "single", "double", "dashed", "dotted", "none". Default "single".
        color: Border color. Default "auto".
        size: Border width in 1/8 points. Default 4.
        offset: Offset from page edge in points. Default 24.
        apply_to: Which pages — "all", "first", "notFirst". Default "all".

    Returns:
        JSON with result.
    """
    _STYLE_MAP = {"none": -1, "single": 1, "double": 3, "dashed": 5, "dotted": 7}
    border_style = _STYLE_MAP.get(style.lower())
    if border_style is None:
        return json.dumps({"success": False, "error": f"Unknown style '{style}'. Use: single, double, dashed, dotted, none"})

    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Page borders via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        sect = doc.Sections(section_index)
        borders = sect.Borders
        for side in range(1, 5):  # -1=wdBorderTop, -2=wdBorderLeft, -3=wdBorderBottom, -4=wdBorderRight
            border = borders(side)
            if border_style == -1:
                border.LineStyle = 0
            else:
                border.LineStyle = border_style
                border.LineWidth = size
                border.ColorIndex = 0  # wdAuto
        return json.dumps({"success": True, "section": section_index, "style": style})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_table_of_figures(filename: Optional[str] = None, caption_label: str = "Figure", title: Optional[str] = None) -> str:
    """Insert a table of figures in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        caption_label: Caption label — "Figure", "Table", "Equation". Default "Figure".
        title: Optional title for the table of figures.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Table of figures via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Insert Table of Figures"):
            sel = app.Selection
            sel.EndKey(Unit=6)  # wdStory
            sel.InsertBreak(Type=7)  # wdSectionBreakNextPage
            label_map = {"figure": 0, "table": 1, "equation": 2}
            label_id = label_map.get(caption_label.lower(), 0)
            doc.TablesOfFigures.Add(Range=sel.Range, Caption=label_id, IncludeLabel=True)
        return json.dumps({"success": True, "caption_label": caption_label})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_shape(filename: Optional[str] = None, shape_type: str = "rectangle", left: float = 72.0, top: float = 72.0, width: float = 144.0, height: float = 72.0, text: Optional[str] = None) -> str:
    """Insert a shape in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        shape_type: Shape type — "rectangle", "oval", "line", "arrow", "callout". Default "rectangle".
        left: Left position in points. Default 72.
        top: Top position in points. Default 72.
        width: Width in points. Default 144.
        height: Height in points. Default 72.
        text: Optional text inside the shape.

    Returns:
        JSON with result.
    """
    _SHAPE_MAP = {"rectangle": 1, "oval": 9, "line": 20, "arrow": 24, "callout": 33, "rounded_rectangle": 5, "diamond": 4, "pentagon": 51, "hexagon": 10, "star": 92}
    mso_type = _SHAPE_MAP.get(shape_type.lower())
    if mso_type is None:
        return json.dumps({"success": False, "error": f"Unknown shape '{shape_type}'. Use: {', '.join(_SHAPE_MAP.keys())}"})

    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Shape insertion via JXA not supported. Use file-based tool."})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Insert Shape"):
            sel = app.Selection
            shape = doc.Shapes.AddShape(mso_type, left, top, width, height, sel.Range)
            if text:
                shape.TextFrame.TextRange.Text = text
        return json.dumps({"success": True, "shape_type": shape_type})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_run_macro(filename: Optional[str] = None, macro_name: str = "", args: Optional[str] = None) -> str:
    """Run a VBA macro in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        macro_name: Name of the macro to run.
        args: Optional comma-separated arguments for the macro.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Macro execution via JXA not supported"})

    if not macro_name:
        return json.dumps({"success": False, "error": "macro_name is required"})

    try:
        from word_document_server.core.word_com import get_word_app
        app = get_word_app()
        if args:
            arg_list = [a.strip() for a in args.split(",")]
            app.Run(macro_name, *arg_list)
        else:
            app.Run(macro_name)
        return json.dumps({"success": True, "macro": macro_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_print(filename: Optional[str] = None, copies: int = 1, pages: Optional[str] = None) -> str:
    """Print an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        copies: Number of copies. Default 1.
        pages: Page range (e.g. "1-3,5"). None = all pages.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_applescript
        cmd = '''
tell application "Microsoft Word"
    print active document without print range
end tell
'''
        return _run_applescript(cmd)

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if pages:
            doc.PrintOut(Range=4, Pages=pages, Copies=copies)  # wdPrintRangeOfPages=4
        else:
            doc.PrintOut(Copies=copies)
        return json.dumps({"success": True, "copies": copies, "pages": pages or "all"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_get_doc_variable(filename: Optional[str] = None, name: str = "") -> str:
    """Get a document variable from an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        name: Variable name.

    Returns:
        JSON with variable value.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Doc variables via JXA not supported"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        val = doc.Variables(name).Value
        return json.dumps({"success": True, "name": name, "value": str(val)})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Variable '{name}' not found: {str(e)}"})


async def word_live_set_doc_variable(filename: Optional[str] = None, name: str = "", value: str = "") -> str:
    """Set a document variable in an open Word document via COM.

    Args:
        filename: Document name or path (None = active document).
        name: Variable name.
        value: Variable value.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "Doc variables via JXA not supported"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        try:
            doc.Variables(name).Value = value
        except Exception:
            doc.Variables.Add(Name=name, Value=value)
        return json.dumps({"success": True, "name": name, "value": value})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
