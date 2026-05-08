#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add error response snippet to unpacked_docx for review.
Generates a formatted error response section that can be reused as a template.
Usage:
    python scripts/add_error_response.py
"""
import os
import sys
import json
import uuid
from copy import deepcopy
from lxml import etree

sys.path.insert(0, os.path.dirname(__file__))
from fix_docx_format import (
    qn, set_fonts, set_size, set_bold, set_color, set_alignment, set_spacing, set_indent,
    set_cell_valign, set_cell_margins, reset_table_cell_paragraph_spacing,
    reset_table_cell_paragraph_indent, get_element_text, make_table_cell,
    apply_paragraph_style_layout, make_blank_paragraph_like,
    parse_template, parse_api_body_rules,
    fix_document, pack_docx
)


def generate_para_id():
    return ''.join(uuid.uuid4().hex.upper()[:8])


def make_run(text, font="Verdana", size=22, bold=False, color="000000"):
    run = etree.Element(qn("r"))
    rPr = etree.SubElement(run, qn("rPr"))
    set_fonts(rPr, font)
    set_size(rPr, size)
    set_bold(rPr, bold)
    set_color(rPr, color)
    t = etree.SubElement(run, qn("t"))
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
    if text.startswith(" ") or text.endswith(" "):
        t.set(XML_SPACE, "preserve")
    t.text = text
    return run


def make_label_para(text, config):
    """Create a label paragraph like 'Sample Response' or 'Attributes description'."""
    para = etree.Element(qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    pStyle = etree.SubElement(pPr, qn("pStyle"))
    pStyle.set(qn("val"), "Normal")
    apply_paragraph_style_layout(pPr, config.get("Normal", {}))
    set_indent(pPr, left_indent=0.25 * 72)
    set_spacing(pPr, 0, 3)
    spacing = pPr.find(qn("spacing"))
    if spacing is None:
        spacing = etree.SubElement(pPr, qn("spacing"))
    spacing.set(qn("line"), "276")
    spacing.set(qn("lineRule"), "auto")
    para.append(make_run(text, size=24, bold=True))
    return para


def make_heading_para(text, style_id, config):
    para = etree.Element(qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    pStyle = etree.SubElement(pPr, qn("pStyle"))
    pStyle.set(qn("val"), style_id)
    apply_paragraph_style_layout(pPr, config.get(style_id, {}))
    para.append(make_run(text, size=config.get(style_id, {}).get("sz", 24), bold=True,
                         color=config.get(style_id, {}).get("color", "0050A8")))
    return para


def find_clone_frame_para(root):
    """Find and clone an existing inline text-box frame.
    Prefer error frames (containing 'error' text) to preserve pink background."""
    body = root.find(".//" + qn("body"))
    if body is None:
        return None
    candidates = []
    for para in body.iter(qn("p")):
        drawing = para.find(".//" + qn("drawing"))
        if drawing is not None:
            txbx = drawing.find(".//" + qn("txbxContent"))
            if txbx is not None:
                text = get_element_text(txbx).lower()
                candidates.append((para, text))
    # Prefer frames that mention error/invalid
    for para, text in candidates:
        if "error" in text or "invalid" in text:
            return deepcopy(para)
    # Fallback to any frame
    if candidates:
        return deepcopy(candidates[0][0])
    return None


def clear_txbx_content(txbx):
    for p in list(txbx.findall(qn("p"))):
        txbx.remove(p)


def add_txbx_line(txbx, text):
    """Add a single line paragraph inside txbxContent.
    Standard: Verdana 10 pt, black, left-aligned (template Section 2.8.5)."""
    para = etree.SubElement(txbx, qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    rPr = etree.SubElement(pPr, qn("rPr"))
    set_fonts(rPr, "Verdana")
    set_size(rPr, 20)

    run = etree.SubElement(para, qn("r"))
    rPr2 = etree.SubElement(run, qn("rPr"))
    set_fonts(rPr2, "Verdana")
    set_size(rPr2, 20)
    set_bold(rPr2, False)
    set_color(rPr2, "000000")

    t = etree.SubElement(run, qn("t"))
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
    if text.startswith(" ") or text.endswith(" "):
        t.set(XML_SPACE, "preserve")
    t.text = text
    return para


def build_txbx_content(txbx, http_code, json_obj):
    """Populate a txbxContent with the error response lines."""
    clear_txbx_content(txbx)
    json_str = json.dumps(json_obj, indent=2, ensure_ascii=False)
    lines = json_str.splitlines()

    add_txbx_line(txbx, f"Http Code: {http_code}")
    add_txbx_line(txbx, "Content-type: application/json")
    add_txbx_line(txbx, "")  # blank line
    for line in lines:
        add_txbx_line(txbx, line)


def build_error_frame(clone_para, http_code, json_obj):
    """Build a new inline frame paragraph from a cloned frame."""
    new_para = deepcopy(clone_para)
    drawing = new_para.find(".//" + qn("drawing"))
    if drawing is None:
        return None

    txbx = drawing.find(".//" + qn("txbxContent"))
    if txbx is None:
        return None

    build_txbx_content(txbx, http_code, json_obj)

    # Center-align the outer paragraph containing the drawing
    pPr = new_para.find(qn("pPr"))
    if pPr is None:
        pPr = etree.SubElement(new_para, qn("pPr"))
    set_alignment(pPr, "center")
    set_indent(pPr, left_indent=0, hanging_indent=0)
    set_spacing(pPr, 0, 3)

    return new_para


def make_response_attributes_table(attributes):
    """Create a Response Attributes table."""
    tbl = etree.Element(qn("tbl"))
    tblPr = etree.SubElement(tbl, qn("tblPr"))
    tblW = etree.SubElement(tblPr, qn("tblW"))
    tblW.set(qn("w"), "5000")
    tblW.set(qn("type"), "pct")
    jc = etree.SubElement(tblPr, qn("jc"))
    jc.set(qn("val"), "center")
    tblInd = etree.SubElement(tblPr, qn("tblInd"))
    tblInd.set(qn("w"), "0")
    tblInd.set(qn("type"), "dxa")
    tblBorders = etree.SubElement(tblPr, qn("tblBorders"))
    for border in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = etree.SubElement(tblBorders, qn(border))
        b.set(qn("val"), "single")
        b.set(qn("sz"), "2")
        b.set(qn("space"), "0")
        b.set(qn("color"), "D9D9D9")
    tblGrid = etree.SubElement(tbl, qn("tblGrid"))
    for w in ("900", "2200", "1400", "1200", "3300"):
        gridCol = etree.SubElement(tblGrid, qn("gridCol"))
        gridCol.set(qn("w"), w)

    # Row 1: merged header "Response Attributes"
    tr1 = etree.SubElement(tbl, qn("tr"))
    trPr1 = etree.SubElement(tr1, qn("trPr"))
    trHeight1 = etree.SubElement(trPr1, qn("trHeight"))
    trHeight1.set(qn("val"), "420")
    trHeight1.set(qn("hRule"), "atLeast")
    tc1 = make_table_cell("Response Attributes", header=True)
    tcPr1 = tc1.find(qn("tcPr"))
    grid_span1 = etree.SubElement(tcPr1, qn("gridSpan"))
    grid_span1.set(qn("val"), "5")
    tcW1 = tcPr1.find(qn("tcW"))
    if tcW1 is not None:
        tcW1.set(qn("w"), "5000")
        tcW1.set(qn("type"), "pct")
    tr1.append(tc1)

    # Row 2: column headers
    tr2 = etree.SubElement(tbl, qn("tr"))
    trPr2 = etree.SubElement(tr2, qn("trPr"))
    trHeight2 = etree.SubElement(trPr2, qn("trHeight"))
    trHeight2.set(qn("val"), "420")
    trHeight2.set(qn("hRule"), "atLeast")
    for h in ("No", "Name", "Type", "Presence", "Description"):
        tr2.append(make_table_cell(h, header=True))

    # Data rows
    for no, attr in enumerate(attributes, start=1):
        tr = etree.SubElement(tbl, qn("tr"))
        trPr = etree.SubElement(tr, qn("trPr"))
        trHeight = etree.SubElement(trPr, qn("trHeight"))
        trHeight.set(qn("val"), "420")
        trHeight.set(qn("hRule"), "atLeast")
        values = [str(no), attr["name"], attr["type"], attr["presence"], attr["description"]]
        for val in values:
            tr.append(make_table_cell(val, header=False))

    return tbl


def insert_before_sectPr(body, elements):
    """Insert elements before the final sectPr in body."""
    sectPr = body.find(qn("sectPr"))
    if sectPr is not None:
        idx = list(body).index(sectPr)
        for offset, el in enumerate(elements):
            body.insert(idx + offset, el)
    else:
        for el in elements:
            body.append(el)


def add_error_response_section(root, config, heading_text, http_code, json_obj, attributes):
    """Insert a complete error response section at the end of the document."""
    body = root.find(".//" + qn("body"))
    if body is None:
        raise RuntimeError("No <w:body> found in document.xml")

    # Build elements
    section = []
    section.append(make_heading_para(heading_text, "Heading3", config))
    section.append(make_blank_paragraph_like(None, config, parse_api_body_rules(os.path.join("scripts", "GoPaperless_DOCX_Format_Template.md")).get("Before Sample Request", {})))
    section.append(make_label_para("Sample Response", config))

    # Frame
    clone_para = find_clone_frame_para(root)
    if clone_para is None:
        raise RuntimeError("No existing inline frame found in document to clone.")
    frame_para = build_error_frame(clone_para, http_code, json_obj)
    section.append(frame_para)

    section.append(make_blank_paragraph_like(None, config, parse_api_body_rules(os.path.join("scripts", "GoPaperless_DOCX_Format_Template.md")).get("Before Attributes description", {})))
    section.append(make_label_para("Attributes description", config))
    section.append(make_response_attributes_table(attributes))

    insert_before_sectPr(body, section)
    return True


def main():
    unpacked = "unpacked_docx"
    if not os.path.isdir(unpacked):
        print(f"Directory '{unpacked}' not found.")
        sys.exit(1)

    template_md = os.path.join("scripts", "GoPaperless_DOCX_Format_Template.md")
    config = parse_template(template_md)
    api_body_rules = parse_api_body_rules(template_md)

    document_xml = os.path.join(unpacked, "word", "document.xml")
    tree = etree.parse(document_xml)
    root = tree.getroot()

    # Define the error response to insert
    error_json = {
        "error_code": "526",
        "error_description": "User missing default enterprise to login",
        "message": "User missing default enterprise to login"
    }
    attributes = [
        {"name": "error_code", "type": "String", "presence": "Mandatory", "description": "Error code identifying the failure."},
        {"name": "error_description", "type": "String", "presence": "Mandatory", "description": "Detailed description of the error."},
        {"name": "message", "type": "String", "presence": "Mandatory", "description": "Human-readable error message."},
    ]

    add_error_response_section(
        root, config,
        heading_text="Error Response Template (Review)",
        http_code="400",
        json_obj=error_json,
        attributes=attributes
    )

    tree.write(document_xml, xml_declaration=True, encoding="utf-8", standalone=True)
    print("[OK] Inserted error response section into document.xml")

    # Run formatter to apply template rules
    from fix_docx_format import (
        parse_table_border_rules, parse_table_property_rules,
        parse_frame_method_rules, parse_heading_hierarchy_rules,
        fix_styles, fix_heading_numbering_layout, fix_document, pack_docx
    )
    table_border_rules = parse_table_border_rules(template_md)
    table_property_rules = parse_table_property_rules(template_md)
    method_rules = parse_frame_method_rules(template_md)
    heading_hierarchy_rules = parse_heading_hierarchy_rules(template_md)

    styles_xml = os.path.join(unpacked, "word", "styles.xml")
    numbering_xml = os.path.join(unpacked, "word", "numbering.xml")
    fix_styles(styles_xml, config)
    fix_heading_numbering_layout(numbering_xml)
    fix_document(document_xml, config, table_border_rules, api_body_rules,
                 table_property_rules, method_rules, heading_hierarchy_rules)

    output_dir = "formatted"
    os.makedirs(output_dir, exist_ok=True)
    output = os.path.join(output_dir, "GoPaperless_Error_Response_Review.docx")
    pack_docx(unpacked, output)
    print(f"\nDone! Review file: {output}")


if __name__ == "__main__":
    main()
