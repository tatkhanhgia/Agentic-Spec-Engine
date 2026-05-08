#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Normalize all existing error response frames in unpacked_docx.
- Parses each error frame to extract the original HTTP status code and JSON body.
- Rebuilds JSON into the standard shape: {error_code, error_description, message}.
- Rebuilds frame content with Verdana 10 pt and 2-space indentation.
Usage:
    python scripts/normalize_error_frames.py
"""
import os
import sys
import json
import re
from copy import deepcopy
from lxml import etree

sys.path.insert(0, os.path.dirname(__file__))
from fix_docx_format import (
    qn, get_element_text,
    parse_template, parse_api_body_rules,
    parse_table_border_rules, parse_table_property_rules,
    parse_frame_method_rules, parse_heading_hierarchy_rules,
    fix_styles, fix_heading_numbering_layout, fix_document, pack_docx
)
from add_error_response import build_txbx_content, clear_txbx_content


def is_error_frame(txbx):
    """Return True if this txbxContent is an error response frame."""
    text = get_element_text(txbx).lower()
    if "http code:" not in text:
        return False
    # Must contain explicit error keys (not just the word 'error' somewhere)
    return "error_description" in text or "error_code" in text


def extract_json_block(text):
    """Extract the first {...} or [...] block from raw frame text."""
    start = -1
    for i, ch in enumerate(text):
        if ch in "{[":
            start = i
            break
    if start == -1:
        return None

    open_char = text[start]
    close_char = "}" if open_char == "{" else "]"
    depth = 0
    end = -1
    for i in range(start, len(text)):
        if text[i] == open_char:
            depth += 1
        elif text[i] == close_char:
            depth -= 1
            if depth == 0:
                end = i
                break
    if end == -1:
        return None
    return text[start:end + 1]


def parse_existing_error(txbx):
    """Return (http_code, old_json_dict) from an error frame, or (None, None)."""
    paragraphs = txbx.findall(qn("p"))
    lines = [get_element_text(p) for p in paragraphs]

    http_code = None
    for line in lines:
        m = re.match(r"Http Code:\s*(\d+)", line.strip(), re.IGNORECASE)
        if m:
            http_code = m.group(1)
            break

    full_text = "\n".join(lines)
    json_block = extract_json_block(full_text)
    if json_block is None:
        return http_code, None

    try:
        data = json.loads(json_block)
    except json.JSONDecodeError:
        return http_code, None

    return http_code, data


AUTH_KEYWORDS = [
    "auth", "authenticate", "login",
    "client credentials", "otp",
    "single sign on", "sso"
]

AUTH_ERROR_TEMPLATE = {
    "error_description": "Invalid Request data",
    "error_code": "1101261",
    "message": "Your client_id is invalid. Please try again"
}

OTHER_ERROR_TEMPLATE = {
    "error_description": "Cannot verify token",
    "error_code": "-17",
    "message": "Không thể kiểm tra tính hợp lệ của token"
}


def get_nearest_heading(para, all_paras):
    """Find nearest preceding Heading2/Heading3 text."""
    try:
        idx = all_paras.index(para)
    except ValueError:
        return ""
    for prev in reversed(all_paras[:idx]):
        pPr = prev.find(qn("pPr"))
        if pPr is None:
            continue
        pStyle = pPr.find(qn("pStyle"))
        if pStyle is None:
            continue
        val = pStyle.get(qn("val"))
        if val in ("Heading2", "Heading3"):
            return get_element_text(prev)
    return ""


def is_auth_heading(text):
    lowered = text.lower()
    return any(kw in lowered for kw in AUTH_KEYWORDS)


def normalize_error_data(data):
    """Map old error JSON keys to the new standard shape."""
    error_code = data.get("error_code") or data.get("error", "")
    error_description = data.get("error_description", "")
    message = data.get("message", error_description)
    return {
        "error_code": str(error_code),
        "error_description": str(error_description),
        "message": str(message),
    }


def normalize_all_error_frames(root):
    """Scan the document and rebuild every error frame with standard templates."""
    body = root.find(".//" + qn("body"))
    all_paras = list(body.iter(qn("p"))) if body is not None else []

    count = 0
    for txbx in root.iter(qn("txbxContent")):
        if not is_error_frame(txbx):
            continue

        http_code, _ = parse_existing_error(txbx)
        if not http_code:
            http_code = "400"

        # Locate the paragraph that hosts this txbxContent
        para = txbx.getparent()
        while para is not None and para.tag != qn("p"):
            para = para.getparent()

        heading_text = get_nearest_heading(para, all_paras) if para is not None else ""
        if is_auth_heading(heading_text):
            new_data = deepcopy(AUTH_ERROR_TEMPLATE)
            # Authentication failures (invalid client_id, invalid token, etc.)
            # should return 401 Unauthorized, not 400/402/403.
            http_code = "401"
        else:
            new_data = deepcopy(OTHER_ERROR_TEMPLATE)

        build_txbx_content(txbx, http_code, new_data)
        count += 1
        print(
            f"[OK] Normalized error frame (HTTP {http_code}) under '{heading_text[:60]}'"
        )

    return count


# ---------------------------------------------------------------------------
# Response Attributes table normalization
# ---------------------------------------------------------------------------

def is_response_attributes_table(tbl):
    """Return True if the table header contains 'Response Attributes'."""
    rows = tbl.findall(qn("tr"))
    if not rows:
        return False
    return "Response Attributes" in get_element_text(rows[0])


def set_cell_text(cell, text):
    """Replace text inside a table cell, preserving the first paragraph and run formatting."""
    paras = cell.findall(qn("p"))
    if not paras:
        return
    para = paras[0]

    # Keep the rPr (run properties) of the first existing run so font/size/bold/color are preserved
    old_r = para.find(qn("r"))
    saved_rPr = None
    if old_r is not None:
        old_rPr = old_r.find(qn("rPr"))
        if old_rPr is not None:
            saved_rPr = deepcopy(old_rPr)

    for r in list(para.findall(qn("r"))):
        para.remove(r)

    run = etree.SubElement(para, qn("r"))
    if saved_rPr is not None:
        run.append(saved_rPr)

    t = etree.SubElement(run, qn("t"))
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
    if text.startswith(" ") or text.endswith(" "):
        t.set(XML_SPACE, "preserve")
    t.text = text


def get_row_cells(row):
    return row.findall(qn("tc"))


def get_row_name(row):
    cells = get_row_cells(row)
    if len(cells) >= 2:
        return get_element_text(cells[1]).strip()
    return ""


def remove_old_error_rows(tbl):
    """Remove rows whose Name column matches old/new error keys so we can re-insert them cleanly."""
    old_names = {"error", "error_code", "error_description", "message", "remark"}
    to_remove = []
    for row in tbl.findall(qn("tr"))[2:]:
        if get_row_name(row) in old_names:
            to_remove.append(row)
    for row in to_remove:
        tbl.remove(row)


def insert_error_rows(tbl, template_row=None):
    """Prepend standard error keys after the two header rows."""
    rows = tbl.findall(qn("tr"))
    if len(rows) < 2:
        return

    if template_row is None:
        for row in rows[2:]:
            if len(get_row_cells(row)) >= 5:
                template_row = row
                break

    if template_row is None:
        # Table has no data rows to clone; skip.
        return

    error_entries = [
        ("error_code", "String", "M", "Error code identifying the failure."),
        ("error_description", "String", "M", "Detailed description of the error."),
        ("message", "String", "M", "Human-readable error message."),
    ]

    anchor = rows[1]  # Insert after the column-header row (No/Name/Type...)
    for name, typ, presence, desc in error_entries:
        new_row = deepcopy(template_row)
        cells = get_row_cells(new_row)
        set_cell_text(cells[1], name)
        set_cell_text(cells[2], typ)
        set_cell_text(cells[3], presence)
        set_cell_text(cells[4], desc)
        anchor.addnext(new_row)
        anchor = new_row  # Next row inserts after this one


def renumber_table(tbl):
    """Re-assign sequential numbers in the No (first) column."""
    idx = 1
    for row in tbl.findall(qn("tr"))[2:]:
        cells = get_row_cells(row)
        if cells:
            set_cell_text(cells[0], str(idx))
            idx += 1


def normalize_table_body_font(tbl, size_val="22"):
    """Force every run in data rows to use Verdana and the given font size (half-points)."""
    for row in tbl.findall(qn("tr"))[2:]:
        for cell in row.findall(qn("tc")):
            for para in cell.iter(qn("p")):
                for run in para.findall(qn("r")):
                    rPr = run.find(qn("rPr"))
                    if rPr is None:
                        rPr = etree.SubElement(run, qn("rPr"))
                    # sz
                    sz = rPr.find(qn("sz"))
                    if sz is None:
                        sz = etree.SubElement(rPr, qn("sz"))
                    sz.set(qn("val"), size_val)
                    # szCs
                    szCs = rPr.find(qn("szCs"))
                    if szCs is None:
                        szCs = etree.SubElement(rPr, qn("szCs"))
                    szCs.set(qn("val"), size_val)
                    # rFonts -> Verdana
                    rFonts = rPr.find(qn("rFonts"))
                    if rFonts is None:
                        rFonts = etree.SubElement(rPr, qn("rFonts"))
                    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                        if rFonts.get(qn(attr)) is None:
                            rFonts.set(qn(attr), "Verdana")


def normalize_all_response_attributes_tables(root):
    """Update every Response Attributes table to list error keys first."""
    count = 0
    for tbl in root.iter(qn("tbl")):
        if not is_response_attributes_table(tbl):
            continue

        # Capture a template data row BEFORE we delete old error rows,
        # because some tables may consist only of error rows.
        rows = tbl.findall(qn("tr"))
        template_row = None
        for row in rows[2:]:
            if len(get_row_cells(row)) >= 5:
                template_row = row
                break

        remove_old_error_rows(tbl)
        insert_error_rows(tbl, template_row)
        renumber_table(tbl)
        normalize_table_body_font(tbl)
        count += 1
    print(f"[OK] Normalized {count} Response Attributes table(s).")
    return count


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

    count = normalize_all_error_frames(root)
    print(f"\nNormalized {count} error frame(s).")

    tbl_count = normalize_all_response_attributes_tables(root)
    print(f"Normalized {tbl_count} Response Attributes table(s).")

    tree.write(document_xml, xml_declaration=True, encoding="utf-8", standalone=True)

    # Apply full template formatting
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
    output = os.path.join(output_dir, "GoPaperless_Normalized_Errors.docx")
    pack_docx(unpacked, output)
    print(f"\nDone! Output: {output}")


if __name__ == "__main__":
    main()
