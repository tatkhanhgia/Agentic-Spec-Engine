#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Analyze Content-Type consistency across API sections in a DOCX document.xml.
Reports APIs missing Content-Type in Header Attributes tables,
and APIs where Content-Type differs from application/json.
"""
import os
import re
import sys
from collections import defaultdict
from lxml import etree

NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
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
    style = get_para_style(para)
    if levels is None:
        levels = {"Heading1", "Heading2", "Heading3", "Heading4"}
    return style in levels

def parse_table(tbl):
    """Parse a table into list of list of strings."""
    rows = []
    for tr in tbl.findall(qn("tr")):
        cells = []
        for tc in tr.findall(qn("tc")):
            cells.append(get_element_text(tc))
        rows.append(cells)
    return rows

def find_header_in_table(tbl, name):
    """Find row index where first column matches name (case-insensitive)."""
    for idx, row in enumerate(parse_table(tbl)):
        if row and row[0].strip().lower() == name.strip().lower():
            return idx, row
    return None, None

def extract_text_from_txbx(txbx):
    """Extract text from txbxContent preserving paragraph boundaries with newlines."""
    paragraphs = []
    for para in txbx.iter(qn("p")):
        para_text = get_element_text(para)
        paragraphs.append(para_text)
    return "\n".join(paragraphs)

def extract_content_type_from_frame(section_nodes):
    """Look inside Sample Request frame paragraphs for Content-Type header value."""
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
            # Look for Content-Type line inside inline shapes (txbxContent)
            for txbx in node.iter(qn("txbxContent")):
                txbx_text = extract_text_from_txbx(txbx)
                m = re.search(r'Content-Type:\s*(\S+)', txbx_text, re.IGNORECASE)
                if m:
                    return m.group(1)
            # Also check direct paragraph text
            m = re.search(r'Content-Type:\s*(\S+)', text, re.IGNORECASE)
            if m:
                return m.group(1)
            # If we see a body start (empty line or { or [), stop looking
            if text.strip() == "" or text.strip().startswith(("{", "[")):
                continue
    return None

def main(document_xml):
    tree = etree.parse(document_xml)
    root = tree.getroot()
    body = root.find(".//" + qn("body"))
    if body is None:
        print("No body found")
        sys.exit(1)

    children = list(body)

    # Identify API sections: Heading3 or Heading4
    sections = []
    current_section = None
    for idx, child in enumerate(children):
        if child.tag == qn("p") and is_heading(child, {"Heading3", "Heading4"}):
            if current_section is not None:
                sections.append(current_section)
            current_section = {
                "heading_para": child,
                "heading_text": get_element_text(child),
                "heading_level": get_para_style(child),
                "start_idx": idx,
                "end_idx": None,
                "nodes": [],
            }
        elif current_section is not None:
            # Stop section if we hit a new Heading3/Heading4 or Heading1/Heading2
            if child.tag == qn("p") and is_heading(child, {"Heading1", "Heading2", "Heading3", "Heading4"}):
                # If it's Heading3/Heading4, it will start a new section in next iteration
                if is_heading(child, {"Heading1", "Heading2"}):
                    current_section["end_idx"] = idx
                    sections.append(current_section)
                    current_section = None
                else:
                    # Heading3/Heading4: end current and new section starts
                    current_section["end_idx"] = idx
                    sections.append(current_section)
                    current_section = {
                        "heading_para": child,
                        "heading_text": get_element_text(child),
                        "heading_level": get_para_style(child),
                        "start_idx": idx,
                        "end_idx": None,
                        "nodes": [],
                    }
            else:
                current_section["nodes"].append(child)

    if current_section is not None:
        sections.append(current_section)

    # For each section, find Header Request Attributes table and Content-Type in it
    results = []
    for sec in sections:
        heading = sec["heading_text"]
        level = sec["heading_level"]
        nodes = sec["nodes"]

        header_table = None
        header_table_rows = []
        for node in nodes:
            if node.tag == qn("tbl"):
                rows = parse_table(node)
                if rows and rows[0] and rows[0][0].strip() in {
                    "Header Attributes",
                    "Header Request Attributes",
                    "Header Response Attributes",
                }:
                    if rows[0][0].strip() in {"Header Attributes", "Header Request Attributes"}:
                        header_table = node
                        header_table_rows = rows
                        break

        ct_in_table = None
        ct_row = None
        if header_table_rows:
            # Look for Content-Type row (skip merged header rows)
            for row in header_table_rows[1:]:  # skip header row
                if len(row) >= 2 and row[1].strip().lower() == "content-type":
                    ct_in_table = row[2] if len(row) > 2 else row[1]
                    ct_row = row
                    break

        ct_in_frame = extract_content_type_from_frame(nodes)

        results.append({
            "heading": heading,
            "level": level,
            "has_header_table": header_table is not None,
            "ct_in_table": ct_in_table,
            "ct_row": ct_row,
            "ct_in_frame": ct_in_frame,
        })

    # Categorize
    missing_ct = []
    inconsistent_ct = []
    special_ct = []
    ok_ct = []

    for r in results:
        heading = r["heading"]
        ct_table = r["ct_in_table"]
        ct_frame = r["ct_in_frame"]

        if not r["has_header_table"]:
            continue  # skip if no header table at all

        if ct_table is None:
            missing_ct.append(r)
        elif ct_frame and ct_table.strip().lower() != ct_frame.strip().lower():
            inconsistent_ct.append(r)
        elif ct_table.strip().lower() != "application/json":
            special_ct.append(r)
        else:
            ok_ct.append(r)

    print("=" * 80)
    print(f"Total API sections analyzed: {len(results)}")
    print(f"Sections with Header table: {sum(1 for r in results if r['has_header_table'])}")
    print("=" * 80)

    print("\n## 1. APIs MISSING Content-Type in Header Attributes table")
    print("-" * 60)
    if missing_ct:
        for r in missing_ct:
            print(f"  - [{r['level']}] {r['heading']}")
            if r["ct_in_frame"]:
                print(f"      -> Frame has Content-Type: {r['ct_in_frame']}")
    else:
        print("  (None)")

    print("\n## 2. APIs with Content-Type MISMATCH (Table vs Frame)")
    print("-" * 60)
    if inconsistent_ct:
        for r in inconsistent_ct:
            print(f"  - [{r['level']}] {r['heading']}")
            print(f"      Table : {r['ct_in_table']}")
            print(f"      Frame : {r['ct_in_frame']}")
    else:
        print("  (None)")

    print("\n## 3. APIs with SPECIAL Content-Type (not application/json)")
    print("-" * 60)
    if special_ct:
        for r in special_ct:
            print(f"  - [{r['level']}] {r['heading']}")
            print(f"      Content-Type: {r['ct_in_table']}")
            if r["ct_in_frame"] and r["ct_in_frame"].lower() != r["ct_in_table"].lower():
                print(f"      Frame has   : {r['ct_in_frame']}")
    else:
        print("  (None)")

    print("\n## 4. APIs with OK Content-Type (application/json, consistent)")
    print("-" * 60)
    print(f"  Count: {len(ok_ct)}")

    print("\n" + "=" * 80)
    print("SUMMARY OF CHANGES NEEDED:")
    print("=" * 80)
    for r in missing_ct:
        ct = r["ct_in_frame"] or "application/json"
        print(f"ADD    Content-Type to table: [{r['level']}] {r['heading']} -> {ct}")
    for r in inconsistent_ct:
        print(f"FIX    Content-Type mismatch: [{r['level']}] {r['heading']} -> use {r['ct_in_frame']}")
    for r in special_ct:
        print(f"CHECK  Special Content-Type : [{r['level']}] {r['heading']} -> {r['ct_in_table']}")

if __name__ == "__main__":
    doc_xml = os.path.join("unpacked_docx", "word", "document.xml")
    if len(sys.argv) > 1:
        doc_xml = sys.argv[1]
    main(doc_xml)
