"""Custom document property tools for Word Document Server."""
import json
import os
import zipfile
import shutil
from typing import Optional, Union

from docx import Document

from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, get_file_lock


async def get_custom_properties(filename: str) -> str:
    """Get all custom document properties.

    Args:
        filename: Path to the Word document.

    Returns:
        JSON with custom properties.
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({"success": False, "error": f"Document {filename} does not exist"})

    try:
        doc = Document(filename)
        props = {}
        core = doc.core_properties
        for attr in ["author", "category", "comments", "content_status", "created", "identifier", "keywords", "language", "last_modified_by", "modified", "revision", "subject", "title", "version"]:
            val = getattr(core, attr, None)
            if val is not None:
                props[attr] = str(val)

        try:
            from lxml import etree
            with zipfile.ZipFile(filename, 'r') as zf:
                if 'docProps/custom.xml' in zf.namelist():
                    custom_xml = zf.read('docProps/custom.xml')
                    root = etree.fromstring(custom_xml)
                    ns_cp = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
                    ns_vt = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
                    for prop in root.findall(f'{{{ns_cp}}}property'):
                        pname = prop.get('name', '')
                        for tag in ['lpwstr', 'i4', 'r8', 'bool']:
                            val_elem = prop.find(f'{{{ns_vt}}}{tag}')
                            if val_elem is not None:
                                props[pname] = val_elem.text or ""
                                break
        except Exception:
            pass

        return json.dumps({"success": True, "properties": props, "count": len(props)})
    except Exception as e:
        return json.dumps({"success": False, "error": f"Failed to get properties: {str(e)}"})


async def set_custom_property(filename: str, name: str, value: Union[str, int, float, bool], property_type: Optional[str] = None) -> str:
    """Set a custom document property.

    Args:
        filename: Path to the Word document.
        name: Property name.
        value: Property value.
        property_type: Value type — "text", "number", "date", "boolean". Auto-detected if None.

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

            if property_type is None:
                if isinstance(value, bool):
                    property_type = "boolean"
                elif isinstance(value, int):
                    property_type = "number"
                elif isinstance(value, float):
                    property_type = "number"
                else:
                    property_type = "text"

            if name in ("author", "category", "comments", "keywords", "subject", "title"):
                core = doc.core_properties
                setattr(core, name, str(value))
                doc.save(filename)
                return json.dumps({"success": True, "property": name, "value": str(value), "type": "core"})

            from lxml import etree
            ns_cp = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
            ns_vt = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

            custom_root = None
            try:
                with zipfile.ZipFile(filename, 'r') as zf:
                    if 'docProps/custom.xml' in zf.namelist():
                        custom_root = etree.fromstring(zf.read('docProps/custom.xml'))
            except Exception:
                pass

            if custom_root is None:
                custom_root = etree.Element(f'{{{ns_cp}}}Properties', nsmap={
                    'cp': ns_cp,
                    'vt': ns_vt,
                })

            existing = custom_root.findall(f'{{{ns_cp}}}property')
            next_id = 2
            for p in existing:
                try:
                    pid = int(p.get('pid', '2'))
                    if pid >= next_id:
                        next_id = pid + 1
                except (ValueError, TypeError):
                    pass
                if p.get('name') == name:
                    custom_root.remove(p)

            prop_elem = etree.SubElement(custom_root, f'{{{ns_cp}}}property')
            prop_elem.set('name', name)
            prop_elem.set('fmtid', '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}')
            prop_elem.set('pid', str(next_id))

            if property_type == "text":
                vt_elem = etree.SubElement(prop_elem, f'{{{ns_vt}}}lpwstr')
            elif property_type == "number":
                if isinstance(value, int):
                    vt_elem = etree.SubElement(prop_elem, f'{{{ns_vt}}}i4')
                else:
                    vt_elem = etree.SubElement(prop_elem, f'{{{ns_vt}}}r8')
            elif property_type == "boolean":
                vt_elem = etree.SubElement(prop_elem, f'{{{ns_vt}}}bool')
                value = "true" if value else "false"
            else:
                vt_elem = etree.SubElement(prop_elem, f'{{{ns_vt}}}lpwstr')

            vt_elem.text = str(value)

            doc.save(filename)

            custom_xml_bytes = etree.tostring(custom_root, xml_declaration=True, encoding='UTF-8', standalone=True)

            tmp_docx = filename + ".tmp_prop"
            has_custom = False
            with zipfile.ZipFile(filename, 'r') as zin:
                with zipfile.ZipFile(tmp_docx, 'w', zipfile.ZIP_DEFLATED) as zout:
                    skip_names = set()
                    if not has_custom:
                        skip_names = {'[Content_Types].xml', '_rels/.rels'}

                    for item in zin.infolist():
                        if item.filename == 'docProps/custom.xml':
                            zout.writestr(item, custom_xml_bytes)
                            has_custom = True
                        elif item.filename in skip_names:
                            pass
                        else:
                            zout.writestr(item, zin.read(item.filename))
                    if not has_custom:
                        zout.writestr('docProps/custom.xml', custom_xml_bytes)

                        zout.writestr('docProps/_rels/custom.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')

                        ct_xml = zin.read('[Content_Types].xml')
                        ct_root = etree.fromstring(ct_xml)
                        ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
                        override = etree.SubElement(ct_root, f'{{{ct_ns}}}Override')
                        override.set('PartName', '/docProps/custom.xml')
                        override.set('ContentType', 'application/vnd.openxmlformats-officedocument.custom-properties+xml')
                        zout.writestr('[Content_Types].xml', etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True))

                        pkg_rels_xml = zin.read('_rels/.rels')
                        pr_root = etree.fromstring(pkg_rels_xml)
                        pr_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
                        max_rid = 0
                        for rel in pr_root.findall(f'{{{pr_ns}}}Relationship'):
                            rid = rel.get('Id', 'rId0')
                            try:
                                num = int(rid.replace('rId', ''))
                                if num > max_rid:
                                    max_rid = num
                            except (ValueError, TypeError):
                                pass
                        new_rel = etree.SubElement(pr_root, f'{{{pr_ns}}}Relationship')
                        new_rel.set('Id', f'rId{max_rid + 1}')
                        new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties')
                        new_rel.set('Target', 'docProps/custom.xml')
                        zout.writestr('_rels/.rels', etree.tostring(pr_root, xml_declaration=True, encoding='UTF-8', standalone=True))

            shutil.move(tmp_docx, filename)
            return json.dumps({"success": True, "property": name, "value": str(value), "type": property_type})
    except Exception as e:
        try:
            os.remove(filename + ".tmp_prop")
        except Exception:
            pass
        return json.dumps({"success": False, "error": f"Failed to set property: {str(e)}"})
