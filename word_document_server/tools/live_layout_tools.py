"""COM-based layout tools for Microsoft Word.

These tools operate on documents currently open in Word via COM automation.
They provide layout, header/footer, spacing, bookmark, watermark, and section
management for files that are open (and locked) in Word.
"""

import json
import sys
from typing import Optional

# macOS JXA dispatch
_MAC_AVAILABLE = sys.platform == 'darwin'

_PTS_PER_INCH = 72.0

_WD_STYLE_MAP = {
    "Normal": -1, "Heading 1": -2, "Heading 2": -3, "Heading 3": -4,
    "Heading 4": -5, "Heading 5": -6, "Heading 6": -7, "Heading 7": -8,
    "Heading 8": -9, "Heading 9": -10, "Title": -63, "Subtitle": -75,
    "Body Text": -67, "Body Text 2": -68, "Body Text 3": -69,
    "List Paragraph": -78, "Quote": -85, "Intense Quote": -86,
    "Caption": -40, "TOC Heading": -89, "TOC 1": -20, "TOC 2": -21,
    "Table Grid": -155, "Light Shading": -156, "Light List": -157,
    "Light Grid": -158, "Medium Shading 1": -159, "Medium Shading 2": -160,
    "Medium List 1": -161, "Medium List 2": -162, "Medium Grid 1": -163,
    "Medium Grid 2": -164, "Medium Grid 3": -165, "Dark List": -166,
    "Colorful Shading": -167, "Colorful List": -168, "Colorful Grid": -169,
    "Light Shading Accent 1": -170, "Light List Accent 1": -171,
    "Light Grid Accent 1": -172, "Medium Shading 1 Accent 1": -173,
    "Medium Shading 2 Accent 1": -174, "Medium List 1 Accent 1": -175,
}


def _resolve_style(doc, style_name: str):
    if style_name in _WD_STYLE_MAP:
        return _WD_STYLE_MAP[style_name]
    for i in range(1, doc.Styles.Count + 1):
        s = doc.Styles(i)
        if s.NameLocal == style_name:
            return s.NameLocal
    return None


async def word_live_set_page_layout(
    filename: Optional[str] = None,
    section_index: int = 1,
    orientation: Optional[str] = None,
    page_width_inches: Optional[float] = None,
    page_height_inches: Optional[float] = None,
    margin_top_inches: Optional[float] = None,
    margin_bottom_inches: Optional[float] = None,
    margin_left_inches: Optional[float] = None,
    margin_right_inches: Optional[float] = None,
) -> str:
    """Set page layout for a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed, COM style). Default 1.
        orientation: "portrait" or "landscape".
        page_width_inches: Page width in inches.
        page_height_inches: Page height in inches.
        margin_top_inches: Top margin in inches.
        margin_bottom_inches: Bottom margin in inches.
        margin_left_inches: Left margin in inches.
        margin_right_inches: Right margin in inches.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_set_page_layout
        return mac_set_page_layout(
            filename=filename, section_index=section_index, orientation=orientation,
            page_width=page_width_inches, page_height=page_height_inches,
            top_margin=margin_top_inches, bottom_margin=margin_bottom_inches,
            left_margin=margin_left_inches, right_margin=margin_right_inches,
        )

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"success": False,
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Set Page Layout"):
            ps = doc.Sections(section_index).PageSetup
            changes = []

            if orientation is not None:
                # wdOrientPortrait=0, wdOrientLandscape=1
                if orientation.lower() == "landscape":
                    ps.Orientation = 1
                    changes.append("orientation=landscape")
                elif orientation.lower() == "portrait":
                    ps.Orientation = 0
                    changes.append("orientation=portrait")

            if page_width_inches is not None:
                ps.PageWidth = page_width_inches * _PTS_PER_INCH
                changes.append(f"width={page_width_inches}in")
            if page_height_inches is not None:
                ps.PageHeight = page_height_inches * _PTS_PER_INCH
                changes.append(f"height={page_height_inches}in")
            if margin_top_inches is not None:
                ps.TopMargin = margin_top_inches * _PTS_PER_INCH
                changes.append(f"margin_top={margin_top_inches}in")
            if margin_bottom_inches is not None:
                ps.BottomMargin = margin_bottom_inches * _PTS_PER_INCH
                changes.append(f"margin_bottom={margin_bottom_inches}in")
            if margin_left_inches is not None:
                ps.LeftMargin = margin_left_inches * _PTS_PER_INCH
                changes.append(f"margin_left={margin_left_inches}in")
            if margin_right_inches is not None:
                ps.RightMargin = margin_right_inches * _PTS_PER_INCH
                changes.append(f"margin_right={margin_right_inches}in")

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "changes": changes,
        })

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_add_header_footer(
    filename: Optional[str] = None,
    section_index: int = 1,
    header_text: Optional[str] = None,
    footer_text: Optional[str] = None,
    header_alignment: str = "center",
    footer_alignment: str = "center",
) -> str:
    """Add header and/or footer text to a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed).
        header_text: Text for the header. None = don't change.
        footer_text: Text for the footer. None = don't change.
        header_alignment: "left", "center", "right".
        footer_alignment: "left", "center", "right".

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_header_footer
        return mac_add_header_footer(filename=filename, section_index=section_index, header_text=header_text, footer_text=footer_text, alignment=header_alignment or footer_alignment)

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"success": False,
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Header/Footer"):
            # Alignment map: 0=left, 1=center, 2=right
            align_map = {"left": 0, "center": 1, "right": 2}
            added = []

            # wdHeaderFooterPrimary = 1
            section = doc.Sections(section_index)

            if header_text is not None:
                hdr = section.Headers(1)  # Primary header
                hdr.Range.Text = header_text
                hdr.Range.ParagraphFormat.Alignment = align_map.get(
                    header_alignment.lower(), 1
                )
                added.append("header")

            if footer_text is not None:
                ftr = section.Footers(1)  # Primary footer
                ftr.Range.Text = footer_text
                ftr.Range.ParagraphFormat.Alignment = align_map.get(
                    footer_alignment.lower(), 1
                )
                added.append("footer")

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "added": added,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_add_page_numbers(
    filename: Optional[str] = None,
    section_index: int = 1,
    position: str = "footer",
    alignment: str = "center",
    prefix: str = "",
    suffix: str = "",
    include_total: bool = False,
) -> str:
    """Add page numbers to header or footer in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed).
        position: "header" or "footer".
        alignment: "left", "center", "right".
        prefix: Text before page number.
        suffix: Text after page number.
        include_total: If True, adds " / N" after page number.

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "word_live_add_page_numbers is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"success": False,
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Page Numbers"):
            # PageNumberAlignment: 0=left, 1=center, 2=right
            align_map = {"left": 0, "center": 1, "right": 2}
            pn_alignment = align_map.get(alignment.lower(), 1)

            section = doc.Sections(section_index)
            # wdHeaderFooterPrimary = 1
            target = section.Headers(1) if position == "header" else section.Footers(1)

            # Add page numbers via PageNumbers collection
            target.PageNumbers.Add(PageNumberAlignment=pn_alignment)

            # Add prefix/suffix/total by editing the range
            if prefix or suffix or include_total:
                rng = target.Range

                # Build the text with field codes
                # Clear and rebuild
                rng.Delete()

                if prefix:
                    rng.InsertAfter(prefix)

                # Insert PAGE field
                # wdFieldPage = 33
                rng.Collapse(0)  # wdCollapseEnd
                app.Selection.GoTo(What=1, Name=str(section_index))  # navigate to section
                field_range = target.Range
                field_range.Collapse(0)
                doc.Fields.Add(Range=field_range, Type=33)  # wdFieldPage

                if include_total:
                    end_range = target.Range
                    end_range.Collapse(0)
                    end_range.InsertAfter(" / ")
                    end_range = target.Range
                    end_range.Collapse(0)
                    doc.Fields.Add(Range=end_range, Type=26)  # wdFieldNumPages

                if suffix:
                    end_range = target.Range
                    end_range.Collapse(0)
                    end_range.InsertAfter(suffix)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "section": section_index,
            "position": position,
            "alignment": alignment,
            "include_total": include_total,
        })

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_add_section_break(
    filename: Optional[str] = None,
    break_type: str = "new_page",
) -> str:
    """Add a section break to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        break_type: "new_page", "continuous", "even_page", "odd_page".

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_section_break
        return mac_add_section_break(filename=filename, break_type=break_type)

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # wdSectionBreakNextPage=2, Continuous=3, EvenPage=4, OddPage=5
        type_map = {
            "new_page": 2,
            "continuous": 3,
            "even_page": 4,
            "odd_page": 5,
        }

        if break_type not in type_map:
            return json.dumps({"success": False,
                "error": f"Invalid break_type: {break_type}. Use: {list(type_map.keys())}"
            })

        with undo_record(app, "MCP: Add Section Break"):
            # Insert at end of document
            end_pos = doc.Content.End - 1
            rng = doc.Range(end_pos, end_pos)
            rng.InsertBreak(Type=type_map[break_type])

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "break_type": break_type,
            "total_sections": doc.Sections.Count,
        })

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_set_paragraph_spacing(
    filename: Optional[str] = None,
    paragraph_index: Optional[int] = None,
    start_paragraph: Optional[int] = None,
    end_paragraph: Optional[int] = None,
    space_before_pt: Optional[float] = None,
    space_after_pt: Optional[float] = None,
    line_spacing: Optional[float] = None,
    line_spacing_rule: Optional[str] = None,
    keep_with_next: Optional[bool] = None,
    keep_together: Optional[bool] = None,
    alignment: Optional[str] = None,
) -> str:
    """Set paragraph spacing and layout properties in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed). Ignored if start/end given.
        start_paragraph: Start of range (1-indexed, inclusive).
        end_paragraph: End of range (1-indexed, inclusive).
        space_before_pt: Space before paragraph in points.
        space_after_pt: Space after paragraph in points.
        line_spacing: Line spacing value IN POINTS (depends on rule).
            IMPORTANT: For "multiple" rule, value is in points, NOT a multiplier.
            Single spacing (1.0) = 12pt. So: 1.15 lines = 13.8pt, 1.5 lines = 18pt,
            2.0 lines = 24pt. Formula: desired_lines * 12 = points_value.
        line_spacing_rule: "single"(0), "1.5_lines"(1), "double"(2),
                           "at_least"(3), "exactly"(4), "multiple"(5).
        keep_with_next: Keep paragraph with next paragraph on same page (True/False).
        keep_together: Keep all lines of paragraph on same page (True/False).
        alignment: Paragraph alignment - "left"(0), "center"(1), "right"(2), "justify"(3).

    Returns:
        JSON with count of affected paragraphs.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_set_paragraph_spacing
        return mac_set_paragraph_spacing(filename=filename, paragraph_index=paragraph_index, start_paragraph=start_paragraph, end_paragraph=end_paragraph, space_before=space_before_pt, space_after=space_after_pt, line_spacing=line_spacing, keep_with_next=keep_with_next, keep_together=keep_together, alignment=alignment)

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.Paragraphs.Count

        # wdLineSpacing rules
        rule_map = {
            "single": 0,
            "1.5_lines": 1,
            "double": 2,
            "at_least": 3,
            "exactly": 4,
            "multiple": 5,
        }

        # Determine range of paragraphs (1-indexed)
        if start_paragraph is not None and end_paragraph is not None:
            indices = range(max(1, start_paragraph), min(end_paragraph + 1, total + 1))
        elif paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > total:
                return json.dumps({"success": False,
                    "error": f"paragraph_index {paragraph_index} out of range (1-{total})"
                })
            indices = [paragraph_index]
        else:
            indices = range(1, total + 1)

        with undo_record(app, "MCP: Set Paragraph Spacing"):
            count = 0
            for i in indices:
                pf = doc.Paragraphs(i).Format
                if space_before_pt is not None:
                    pf.SpaceBefore = space_before_pt
                if space_after_pt is not None:
                    pf.SpaceAfter = space_after_pt
                if line_spacing_rule is not None and line_spacing_rule in rule_map:
                    pf.LineSpacingRule = rule_map[line_spacing_rule]
                if line_spacing is not None:
                    pf.LineSpacing = line_spacing
                if keep_with_next is not None:
                    pf.KeepWithNext = keep_with_next
                if keep_together is not None:
                    pf.KeepTogether = keep_together
                if alignment is not None:
                    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
                    if alignment in align_map:
                        pf.Alignment = align_map[alignment]
                count += 1

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "paragraphs_affected": count,
        })

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_add_bookmark(
    filename: Optional[str] = None,
    paragraph_index: int = 1,
    bookmark_name: str = "",
) -> str:
    """Add a named bookmark at a paragraph in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph to bookmark (1-indexed).
        bookmark_name: Bookmark name (alphanumeric + underscore, no spaces).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import mac_add_bookmark
        return mac_add_bookmark(filename=filename, paragraph_index=paragraph_index, bookmark_name=bookmark_name)

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    if not bookmark_name:
        return json.dumps({"success": False, "error": "bookmark_name is required"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
            return json.dumps({"success": False,
                "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
            })

        with undo_record(app, "MCP: Add Bookmark"):
            rng = doc.Paragraphs(paragraph_index).Range
            doc.Bookmarks.Add(bookmark_name, rng)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "bookmark_name": bookmark_name,
            "paragraph_index": paragraph_index,
        })

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_add_watermark(
    filename: Optional[str] = None,
    text: str = "TASLAK",
    font_size: int = 72,
    font_color: str = "C0C0C0",
    rotation: int = -45,
    section_index: int = 1,
) -> str:
    """Add a diagonal text watermark to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Watermark text (e.g. "TASLAK", "DRAFT", "GİZLİ").
        font_size: Font size in points.
        font_color: Hex color without # (e.g. "C0C0C0").
        rotation: Rotation angle in degrees (e.g. -45).
        section_index: Section number (1-indexed).

    Returns:
        JSON with result info.
    """
    if _MAC_AVAILABLE:
        return json.dumps({"success": False, "error": "word_live_add_watermark is not yet implemented on macOS"})

    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Live layout tools are only available on Windows"})

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"success": False,
                "error": f"Section {section_index} out of range (1-{doc.Sections.Count})"
            })

        with undo_record(app, "MCP: Add Watermark"):
            section = doc.Sections(section_index)
            header = section.Headers(1)  # wdHeaderFooterPrimary

            # Parse color
            c = font_color.lstrip("#")
            r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
            rgb_color = r + (g << 8) + (b << 16)

            # AddTextEffect(PresetTextEffect, Text, FontName, FontSize,
            #               FontBold, FontItalic, Left, Top)
            # COM requires positional args
            shape = header.Shapes.AddTextEffect(
                0, text, "Calibri", font_size, False, False, 0, 0
            )

            # Configure shape
            shape.Fill.ForeColor.RGB = rgb_color
            shape.Fill.Transparency = 0.5
            shape.Line.Visible = False  # msoFalse
            shape.Rotation = rotation
            shape.LockAspectRatio = False

            # Position relative to page center
            # msoRelativeHorizontalPositionMargin = 0
            # msoRelativeVerticalPositionMargin = 0
            shape.RelativeHorizontalPosition = 0
            shape.RelativeVerticalPosition = 0
            shape.Left = -999995  # wdShapeCenter (magic value for centering)
            shape.Top = -999995  # wdShapeCenter

            # Send behind text
            shape.WrapFormat.Type = 3  # wdWrapBehind
            shape.WrapFormat.AllowOverlap = True

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "text": text,
            "font_size": font_size,
            "color": font_color,
            "rotation": rotation,
            "section": section_index,
        }, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Live Style Tools ────────────────────────────────────────────────────────


async def word_live_list_styles(filename: Optional[str] = None) -> str:
    """List all styles in an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with styles.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var ss = d.styles();
var results = [];
for (var i = 0; i < ss.length; i++) {{
    var s = ss[i];
    results.push({{name: s.name(), type: s.type().toString(), builtin: s.builtIn()}});
}}
JSON.stringify({{success: true, count: results.length, styles: results}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        styles = []
        for i in range(1, doc.Styles.Count + 1):
            s = doc.Styles(i)
            styles.append({"name": s.NameLocal, "type": s.Type, "builtin": s.BuiltIn})
        return json.dumps({"success": True, "count": len(styles), "styles": styles})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_apply_style(
    filename: Optional[str] = None,
    paragraph_index: Optional[int] = None,
    style_name: str = "Normal",
    apply_to: str = "paragraph",
) -> str:
    """Apply a style to a paragraph or selection in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: 1-based paragraph index. None = apply to selection.
        style_name: Style name to apply.
        apply_to: "paragraph" or "selection". Default "paragraph".

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        idx_js = f"var para = d.paragraphs[{paragraph_index - 1}];" if paragraph_index else "var para = app.selection.paragraphs[0];"
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
{idx_js}
para.style = "{style_name}";
JSON.stringify({{success: true, style: "{style_name}"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: Apply Style"):
            if paragraph_index:
                if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                    return json.dumps({"success": False, "error": f"paragraph_index {paragraph_index} out of range"})
                rng = doc.Paragraphs(paragraph_index).Range
            else:
                rng = app.Selection.Range
            try:
                rng.Style = doc.Styles(style_name)
            except Exception:
                resolved = _resolve_style(doc, style_name)
                if resolved is None:
                    return json.dumps({"success": False, "error": f"Style '{style_name}' not found"})
                rng.Style = resolved
        return json.dumps({"success": True, "style": style_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_modify_style(
    filename: Optional[str] = None,
    style_name: str = "Normal",
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
) -> str:
    """Modify properties of an existing style in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        style_name: Style to modify.
        font_name: New font name.
        font_size: New font size in points.
        bold: Bold on/off.
        italic: Italic on/off.
        color: Color name (e.g. "red", "blue").

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        props = []
        if font_name:
            props.append(f's.fontName = "{font_name}";')
        if font_size:
            props.append(f"s.fontSize = {font_size};")
        if bold is not None:
            props.append(f's.bold = {"true" if bold else "false"};')
        if italic is not None:
            props.append(f's.italic = {"true" if italic else "false"};')
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var s = d.styles().byName("{style_name}");
{chr(10).join(props)}
JSON.stringify({{success: true, style: "{style_name}"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        try:
            style = doc.Styles(style_name)
        except Exception:
            resolved = _resolve_style(doc, style_name)
            if resolved is None:
                return json.dumps({"success": False, "error": f"Style '{style_name}' not found"})
            style = doc.Styles(resolved)
        font = style.Font
        if font_name:
            font.Name = font_name
        if font_size:
            font.Size = font_size
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if color:
            from word_document_server.tools.live_doc_tools import _COLOR_MAP
            if color in _COLOR_MAP:
                font.ColorIndex = _COLOR_MAP[color]
            else:
                return json.dumps({"success": False, "error": f"Unknown color '{color}'"})
        return json.dumps({"success": True, "style": style_name})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_set_section_columns(
    filename: Optional[str] = None,
    section_index: int = 1,
    columns: int = 2,
    spacing: Optional[float] = None,
    equal_width: bool = True,
) -> str:
    """Set column layout for a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-based). Default 1.
        columns: Number of columns. Default 2.
        spacing: Spacing between columns in points. None = default.
        equal_width: Equal column widths. Default True.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        spacing_js = f", columnSpacing: {spacing}" if spacing else ""
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var ps = d.sections[{section_index - 1}].pageSetup;
ps.columnCount = {columns};
ps.equalColumnWidth = {"true" if equal_width else "false"}{spacing_js};
JSON.stringify({{success: true, section: {section_index}, columns: {columns}}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record
        app = get_word_app()
        doc = find_document(app, filename)
        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"success": False, "error": f"Section {section_index} out of range"})
        with undo_record(app, "MCP: Set Section Columns"):
            sec = doc.Sections(section_index)
            ps = sec.PageSetup
            ps.TextColumns.SetCount(columns)
            if spacing is not None:
                ps.TextColumns.Spacing = spacing
            if not equal_width and columns > 1:
                ps.TextColumns.EvenlySpaced = False
        return json.dumps({"success": True, "section": section_index, "columns": columns})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_column_break(filename: Optional[str] = None) -> str:
    """Insert a column break at the current selection in an open Word document.

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
app.selection.insertBreak({{breakType: "column break"}});
JSON.stringify({{success: true, action: "column_break"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, undo_record
        app = get_word_app()
        with undo_record(app, "MCP: Insert Column Break"):
            app.Selection.InsertBreak(Type=8)  # wdColumnBreak = 8
        return json.dumps({"success": True, "action": "column_break"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_get_section_layout(filename: Optional[str] = None, section_index: int = 1) -> str:
    """Get layout properties of a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed). Default 1.

    Returns:
        JSON with section layout info.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
var secIdx = Math.max(0, {section_index} - 1);
var ps = d.sections[secIdx].pageSetup;
JSON.stringify({{
    success: true,
    section: {section_index},
    orientation: ps.orientation().toString(),
    pageWidth: ps.pageWidth(),
    pageHeight: ps.pageHeight(),
    topMargin: ps.topMargin(),
    bottomMargin: ps.bottomMargin(),
    leftMargin: ps.leftMargin(),
    rightMargin: ps.rightMargin(),
    columns: ps.columnCount()
}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if section_index < 1 or section_index > doc.Sections.Count:
            return json.dumps({"success": False, "error": f"Section {section_index} out of range"})
        ps = doc.Sections(section_index).PageSetup
        info = {
            "success": True,
            "section": section_index,
            "orientation": "landscape" if ps.Orientation == 1 else "portrait",
            "page_width_pt": ps.PageWidth,
            "page_height_pt": ps.PageHeight,
            "margin_top_pt": ps.TopMargin,
            "margin_bottom_pt": ps.BottomMargin,
            "margin_left_pt": ps.LeftMargin,
            "margin_right_pt": ps.RightMargin,
            "columns": ps.TextColumns.Count,
        }
        return json.dumps(info, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Live Advanced Table Tools ───────────────────────────────────────────────


async def word_live_delete_table(filename: Optional[str] = None, table_index: int = 0) -> str:
    """Delete a table from an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: Index of the table (0-based). Default 0.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
if ({table_index} >= d.tables.length) throw new Error("Table index out of range");
app.delete(d.tables[{table_index}]);
JSON.stringify({{success: true, deleted_table: {table_index}}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return json.dumps({"success": False, "error": f"Table index {table_index} out of range"})
        doc.Tables(table_index + 1).Delete()
        return json.dumps({"success": True, "deleted_table": table_index})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_repeat_table_header(filename: Optional[str] = None, table_index: int = 0, header_rows: int = 1) -> str:
    """Set header rows of a table to repeat on each page in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: Index of the table (0-based). Default 0.
        header_rows: Number of header rows to repeat. Default 1.

    Returns:
        JSON with result.
    """
    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return json.dumps({"success": False, "error": f"Table index {table_index} out of range"})
        tbl = doc.Tables(table_index + 1)
        for i in range(1, min(header_rows + 1, tbl.Rows.Count + 1)):
            tbl.Rows(i).HeadingFormat = -1
        return json.dumps({"success": True, "table": table_index, "header_rows": header_rows})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_sort_table(filename: Optional[str] = None, table_index: int = 0, column: int = 0, descending: bool = False, header_row: bool = True) -> str:
    """Sort a table by a column in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        table_index: Index of the table (0-based). Default 0.
        column: Column to sort by (0-based). Default 0.
        descending: Sort descending. Default False.
        header_row: First row is header. Default True.

    Returns:
        JSON with result.
    """
    if _MAC_AVAILABLE:
        from word_document_server.core.word_mac import _run_jxa, _doc_finder_js
        finder = _doc_finder_js(filename)
        sort_dir = "descending" if descending else "ascending"
        hdr_js = "true" if header_row else "false"
        return _run_jxa(f"""
var app = Application("Microsoft Word");
{finder}
if ({table_index} >= d.tables.length) throw new Error("Table index out of range");
app.sort(d.tables[{table_index}], {{by: {column}, order: "{sort_dir}", excludeHeader: {hdr_js}}});
JSON.stringify({{success: true, table: {table_index}, sorted_by: {column}}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        if table_index < 0 or table_index >= doc.Tables.Count:
            return json.dumps({"success": False, "error": f"Table index {table_index} out of range"})
        tbl = doc.Tables(table_index + 1)
        sort_order = 2 if descending else 1
        hdr = 1 if header_row else 0
        col1_index = column + 1
        rng = tbl.Range
        rng.Sort(
            ExcludeHeader=hdr,
            SortFieldType=0,
            SortOrder=sort_order,
            SortColumn=col1_index,
        )
        return json.dumps({"success": True, "table": table_index, "sorted_by": column, "descending": descending})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


# ── Live Object Insertion Tools ─────────────────────────────────────────────


async def word_live_update_table_of_contents(filename: Optional[str] = None) -> str:
    """Update/refresh an existing Table of Contents in an open Word document.

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
var tocCount = d.tablesOfContents.length;
if (tocCount === 0) throw new Error("No TOC found");
for (var i = 0; i < tocCount; i++) app.update(d.tablesOfContents[i]);
JSON.stringify({{success: true, updated_tocs: tocCount}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        toc_count = doc.TablesOfContents.Count
        if toc_count == 0:
            return json.dumps({"success": False, "error": "No Table of Contents found"})
        for i in range(1, toc_count + 1):
            doc.TablesOfContents(i).Update()
        return json.dumps({"success": True, "updated_tocs": toc_count})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_text_box(filename: Optional[str] = None, text: str = "", left_inches: float = 1.0, top_inches: float = 1.0, width_inches: float = 3.0, height_inches: float = 1.0) -> str:
    """Insert a text box into an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Text content for the text box.
        left_inches: Left position in inches. Default 1.0.
        top_inches: Top position in inches. Default 1.0.
        width_inches: Width in inches. Default 3.0.
        height_inches: Height in inches. Default 1.0.

    Returns:
        JSON with result.
    """
    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        left_pt = float(left_inches * 72)
        top_pt = float(top_inches * 72)
        width_pt = float(width_inches * 72)
        height_pt = float(height_inches * 72)
        tb = doc.Shapes.AddTextbox(1, left_pt, top_pt, width_pt, height_pt)
        tb.TextFrame.TextRange.Text = text
        return json.dumps({"success": True, "text_box": text[:50], "position": f"{left_inches}x{top_inches}in", "size": f"{width_inches}x{height_inches}in"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_chart(filename: Optional[str] = None, chart_type: str = "bar", title: Optional[str] = None, categories: Optional[str] = None, values: Optional[str] = None) -> str:
    """Insert a chart into an open Word document (Windows only).

    Args:
        filename: Document name or path (None = active document).
        chart_type: Chart type — "bar", "column", "line", "pie", "area", "scatter". Default "bar".
        title: Optional chart title.
        categories: Comma-separated category labels.
        values: Comma-separated numeric values.

    Returns:
        JSON with result.
    """
    if sys.platform != "win32":
        return json.dumps({"success": False, "error": "Insert chart requires Microsoft Word on Windows"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        chart_type_map = {"bar": 2, "column": 51, "line": 4, "pie": 5, "area": 1, "scatter": -4169}
        xl_chart = chart_type_map.get(chart_type.lower())
        if xl_chart is None:
            return json.dumps({"success": False, "error": f"Unknown chart type '{chart_type}'"})
        sel = app.Selection
        inline_shape = doc.InlineShapes.AddChart2(Style=-1, Type=xl_chart, Range=sel.Range)
        if categories and values:
            chart = inline_shape.Chart
            chart_data = chart.ChartData
            wb = chart_data.Workbook
            ws = wb.Worksheets(1)
            cats = [c.strip() for c in categories.split(",")]
            vals_list = [float(v.strip()) for v in values.split(",")]
            ws.Cells(1, 1).Value = "Category"
            ws.Cells(1, 2).Value = "Value"
            for i, cat in enumerate(cats):
                ws.Cells(i + 2, 1).Value = cat
            for i, val in enumerate(vals_list):
                ws.Cells(i + 2, 2).Value = val
            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
        return json.dumps({"success": True, "chart_type": chart_type, "title": title})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})


async def word_live_insert_index(filename: Optional[str] = None) -> str:
    """Insert a subject index at the end of an open Word document.

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
app.make({{new: "index", at: d}});
JSON.stringify({{success: true, action: "insert_index"}});
""")

    try:
        from word_document_server.core.word_com import get_word_app, find_document
        app = get_word_app()
        doc = find_document(app, filename)
        sel = app.Selection
        sel.EndKey(Unit=6)
        sel.InsertBreak(Type=7)  # wdSectionBreakNextPage
        doc.Indexes.Add(Range=sel.Range, HeadingSeparator=0, Type=0, RightAlignPageNumbers=True)
        doc.Indexes(1).Update()
        return json.dumps({"success": True, "action": "insert_index"})
    except Exception as e:
        return json.dumps({"success": False, "error": str(e)})
