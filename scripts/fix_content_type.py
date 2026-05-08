#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Fix Content-Type consistency in Header Request Attributes tables.
- Adds missing Content-Type rows.
- Fixes Content-Type rows where Type is 'String' instead of the actual value.
- Adjusts Description to match the content type.
"""
import os
import re
from copy import deepcopy
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def qn(tag):
    return W + tag

def get_element_text(element):
    return "".join(t.text or "" for t in element.iter(qn("t"))).strip()

def get_para_style(para):
    pPr = para.find(qn("pPr"))
    if pPr is not None:
        ps = pPr.find(qn("pStyle"))
        if ps is not None:
            return ps.get(qn("val"))
    return None

def is_heading(para, levels=None):
    if levels is None:
        levels = {"Heading1", "Heading2", "Heading3", "Heading4"}
    return get_para_style(para) in levels

def parse_table(tbl):
    rows = []
    for tr in tbl.findall(qn("tr")):
        cells = [get_element_text(tc) for tc in tr.findall(qn("tc"))]
        rows.append(cells)
    return rows

def extract_text_from_txbx(txbx):
    paragraphs = []
    for para in txbx.iter(qn("p")):
        paragraphs.append(get_element_text(para))
    return "\n".join(paragraphs)

def extract_content_type_from_frame(section_nodes):
    in_request = False
    for node in section_nodes:
        if node.tag != qn("p"):
            continue
        text = get_element_text(node)
        if text == "Sample Request":
            in_request = True
            continue
        if in_request and text == "Sample Response":
            break
        if in_request and text == "Attributes description":
            break
        if in_request:
            for txbx in node.iter(qn("txbxContent")):
                txbx_text = extract_text_from_txbx(txbx)
                m = re.search(r'Content-Type:\s*(\S+)', txbx_text, re.IGNORECASE)
                if m:
                    return m.group(1)
            m = re.search(r'Content-Type:\s*(\S+)', text, re.IGNORECASE)
            if m:
                return m.group(1)
    return None

def find_section_tables(section_nodes):
    tables = []
    for node in section_nodes:
        if node.tag == qn("tbl"):
            rows = parse_table(node)
            if rows and rows[0] and rows[0][0].strip() in {
                "Header Attributes",
                "Header Request Attributes",
                "Header Response Attributes",
            }:
                tables.append((node, rows))
    return tables

def make_cell_with_text(text, width, align="left"):
    tc = etree.Element(qn("tc"))
    tcPr = etree.SubElement(tc, qn("tcPr"))
    tcW = etree.SubElement(tcPr, qn("tcW"))
    tcW.set(qn("w"), str(width))
    tcW.set(qn("type"), "dxa")
    shd = etree.SubElement(tcPr, qn("shd"))
    shd.set(qn("val"), "clear")
    shd.set(qn("color"), "auto")
    shd.set(qn("fill"), "auto")
    vAlign = etree.SubElement(tcPr, qn("vAlign"))
    vAlign.set(qn("val"), "center")

    p = etree.SubElement(tc, qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))
    pStyle = etree.SubElement(pPr, qn("pStyle"))
    pStyle.set(qn("val"), "Default")
    spacing = etree.SubElement(pPr, qn("spacing"))
    spacing.set(qn("line"), "276")
    spacing.set(qn("lineRule"), "auto")
    jc = etree.SubElement(pPr, qn("jc"))
    jc.set(qn("val"), align)
    rPr = etree.SubElement(pPr, qn("rPr"))
    sz = etree.SubElement(rPr, qn("sz"))
    sz.set(qn("val"), "20")
    szCs = etree.SubElement(rPr, qn("szCs"))
    szCs.set(qn("val"), "20")

    r = etree.SubElement(p, qn("r"))
    rPr2 = etree.SubElement(r, qn("rPr"))
    sz2 = etree.SubElement(rPr2, qn("sz"))
    sz2.set(qn("val"), "20")
    szCs2 = etree.SubElement(rPr2, qn("szCs"))
    szCs2.set(qn("val"), "20")
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return tc

def make_content_type_row(no, content_type, presence="M"):
    tr = etree.Element(qn("tr"))
    trPr = etree.SubElement(tr, qn("trPr"))
    trHeight = etree.SubElement(trPr, qn("trHeight"))
    trHeight.set(qn("val"), "325")

    if content_type == "application/octet-stream":
        description = "Request content type for binary file upload"
    elif content_type == "application/json":
        description = "Request content type (JSON)"
    else:
        description = f"Request content type ({content_type})"

    # Clone widths from a known table structure (Upload Document example)
    widths = [584, 2615, 1409, 1444, 4325]
    aligns = ["center", "left", "left", "center", "left"]
    values = [str(no), "Content-Type", content_type, presence, description]

    for w, a, v in zip(widths, aligns, values):
        tr.append(make_cell_with_text(v, w, a))
    return tr

def get_header_request_table(section_nodes):
    for node in section_nodes:
        if node.tag == qn("tbl"):
            rows = parse_table(node)
            if rows and rows[0] and rows[0][0].strip() in {"Header Attributes", "Header Request Attributes"}:
                return node, rows
    return None, None

def fix_content_type_in_document(document_path):
    tree = etree.parse(document_path)
    root = tree.getroot()
    body = root.find(".//" + qn("body"))
    if body is None:
        print("No body found")
        return

    children = list(body)

    # Identify API sections
    sections = []
    current_section = None
    for idx, child in enumerate(children):
        if child.tag == qn("p") and is_heading(child, {"Heading3", "Heading4"}):
            if current_section is not None:
                sections.append(current_section)
            current_section = {
                "heading_text": get_element_text(child),
                "start_idx": idx,
                "nodes": [],
            }
        elif current_section is not None:
            if child.tag == qn("p") and is_heading(child, {"Heading1", "Heading2", "Heading3", "Heading4"}):
                current_section["end_idx"] = idx
                sections.append(current_section)
                if is_heading(child, {"Heading3", "Heading4"}):
                    current_section = {
                        "heading_text": get_element_text(child),
                        "start_idx": idx,
                        "nodes": [],
                    }
                else:
                    current_section = None
            else:
                current_section["nodes"].append(child)

    if current_section is not None:
        sections.append(current_section)

    changes = []
    for sec in sections:
        heading = sec["heading_text"]
        nodes = sec["nodes"]

        tbl, rows = get_header_request_table(nodes)
        if tbl is None:
            continue

        ct_row_idx = None
        ct_row = None
        for idx, row in enumerate(rows):
            if len(row) >= 2 and row[1].strip().lower() == "content-type":
                ct_row_idx = idx
                ct_row = row
                break

        ct_frame = extract_content_type_from_frame(nodes)
        # Determine expected content type
        expected_ct = ct_frame or "application/json"
        # Fix known typo in frames
        if expected_ct.lower() == "application/octect-stream":
            expected_ct = "application/octet-stream"

        if ct_row_idx is None:
            # Missing Content-Type: add it
            # Find last row in the Header Request Attributes table
            tr_elements = tbl.findall(qn("tr"))
            # Count rows that belong to this table (skip header row 0 which is merged)
            # Row 0 = merged "Header Request Attributes", Row 1 = column headers
            # Data rows start from index 2
            last_data_row = tr_elements[-1]
            new_no = len(tr_elements) - 1  # after header + column headers
            new_row = make_content_type_row(new_no, expected_ct, "M")
            last_data_row.addnext(new_row)
            changes.append({
                "action": "ADD",
                "api": heading,
                "value": expected_ct,
            })
        else:
            # Content-Type exists; check if Type column is wrong
            # row structure: [No, Name, Type, Presence, Description]
            type_value = ct_row[2].strip() if len(ct_row) > 2 else ""
            if type_value.lower() in ("string", ""):
                tr_elements = tbl.findall(qn("tr"))
                if ct_row_idx < len(tr_elements):
                    target_tr = tr_elements[ct_row_idx]
                    # Update Type cell (index 2)
                    cells = target_tr.findall(qn("tc"))
                    if len(cells) >= 3:
                        type_cell = cells[2]
                        # Clear existing paragraphs
                        for para in list(type_cell.findall(qn("p"))):
                            type_cell.remove(para)
                        # Add new paragraph with correct type
                        p = etree.SubElement(type_cell, qn("p"))
                        pPr = etree.SubElement(p, qn("pPr"))
                        pStyle = etree.SubElement(pPr, qn("pStyle"))
                        pStyle.set(qn("val"), "Default")
                        spacing = etree.SubElement(pPr, qn("spacing"))
                        spacing.set(qn("line"), "276")
                        spacing.set(qn("lineRule"), "auto")
                        rPr = etree.SubElement(pPr, qn("rPr"))
                        sz = etree.SubElement(rPr, qn("sz"))
                        sz.set(qn("val"), "20")
                        szCs = etree.SubElement(rPr, qn("szCs"))
                        szCs.set(qn("val"), "20")
                        r = etree.SubElement(p, qn("r"))
                        rPr2 = etree.SubElement(r, qn("rPr"))
                        sz2 = etree.SubElement(rPr2, qn("sz"))
                        sz2.set(qn("val"), "20")
                        szCs2 = etree.SubElement(rPr2, qn("szCs"))
                        szCs2.set(qn("val"), "20")
                        t = etree.SubElement(r, qn("t"))
                        t.text = expected_ct

                        # Also update Description cell (index 4)
                        if len(cells) >= 5:
                            desc_cell = cells[4]
                            for para in list(desc_cell.findall(qn("p"))):
                                desc_cell.remove(para)
                            desc_text = (
                                "Request content type for binary file upload"
                                if expected_ct == "application/octet-stream"
                                else "Request content type (JSON)"
                            )
                            p2 = etree.SubElement(desc_cell, qn("p"))
                            pPr2 = etree.SubElement(p2, qn("pPr"))
                            pStyle2 = etree.SubElement(pPr2, qn("pStyle"))
                            pStyle2.set(qn("val"), "Default")
                            spacing2 = etree.SubElement(pPr2, qn("spacing"))
                            spacing2.set(qn("line"), "276")
                            spacing2.set(qn("lineRule"), "auto")
                            rPr3 = etree.SubElement(pPr2, qn("rPr"))
                            sz3 = etree.SubElement(rPr3, qn("sz"))
                            sz3.set(qn("val"), "20")
                            szCs3 = etree.SubElement(rPr3, qn("szCs"))
                            szCs3.set(qn("val"), "20")
                            r2 = etree.SubElement(p2, qn("r"))
                            rPr4 = etree.SubElement(r2, qn("rPr"))
                            sz4 = etree.SubElement(rPr4, qn("sz"))
                            sz4.set(qn("val"), "20")
                            szCs4 = etree.SubElement(rPr4, qn("szCs"))
                            szCs4.set(qn("val"), "20")
                            t2 = etree.SubElement(r2, qn("t"))
                            t2.text = desc_text

                        changes.append({
                            "action": "FIX",
                            "api": heading,
                            "old": type_value,
                            "new": expected_ct,
                        })

    tree.write(document_path, xml_declaration=True, encoding="utf-8", standalone=True)

    print("=" * 80)
    print(f"Total API sections analyzed: {len(sections)}")
    print(f"Changes made: {len(changes)}")
    print("=" * 80)

    adds = [c for c in changes if c["action"] == "ADD"]
    fixes = [c for c in changes if c["action"] == "FIX"]

    print(f"\n## 1. ADDED Content-Type rows ({len(adds)} APIs)")
    print("-" * 60)
    for c in adds:
        print(f"  + {c['api']} -> {c['value']}")

    print(f"\n## 2. FIXED Content-Type Type value ({len(fixes)} APIs)")
    print("-" * 60)
    for c in fixes:
        print(f"  ~ {c['api']}: '{c['old']}' -> '{c['new']}'")

    return changes

if __name__ == "__main__":
    doc_xml = os.path.join("unpacked_docx", "word", "document.xml")
    fix_content_type_in_document(doc_xml)
