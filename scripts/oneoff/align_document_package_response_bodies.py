#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Align the four object-returning Document Package endpoints with the backend.

Everything below was read from the backend serialisers, read-only:

`Get Package Details` -> `GetPackageDetailsSerializer`
    documents[] carries `uploaded_on` / `uploaded_by`, not `created_at`, and
    also emits `modified_by`, which the sample omitted.

`Get Document Details` -> `GetDocumentDetailsSerializer`
    emits `form_field` (singular) and does NOT emit `document_source`,
    `certify` or `template`. It does emit `created_by` and `modified_by`,
    which the sample omitted.

`Get All Packages` -> `GetListPackageSerializer`
    emits `owner_email`, spells the object `access_duration`, and writes
    `decline` as an object. The old sample was not valid JSON (missing commas,
    `decline` shown as an array of key/value pairs) and carried an `unread`
    field the backend never writes, plus a misspelled `acccess_duration`.

`Apply Workflow Template` -> `TemplateService.applyTemplate`
    the success path is `return new InternalResponse<>()` - no body at all.
    The section documented a 16-field document object that is never returned.
    Its `Content-Type` description also claimed octet-stream while the endpoint
    takes a JSON body (`template_name`, `apply_to_all`).

Field names and structure are aligned to what the code emits. Descriptions keep
stating the intended meaning - several serialiser fields are filled with the
wrong source value in the backend, and those are reported separately rather
than written into the specification as if they were the contract.

Usage:
    python scripts/oneoff/align_document_package_response_bodies.py [--dry-run]
"""
import copy
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted"
    "_backup_pre_response_body_alignment.docx")

HEADINGS = {"Heading1", "Heading2", "Heading3"}
FRAME_HEADER = re.compile(r"^(Http Code:|Content-type:|Content-Type:|x-[a-z0-9-]+:)")

CONTENT_TYPE_TEXT = "The MIME type of the request body. Usually application/json"

PACKAGE_DETAILS_BODY = [
    "{",
    '  "name": "Package Name",',
    '  "owner": {',
    '    "name": "String",',
    '    "email": "email"',
    "  },",
    '  "document_status": "String",',
    '  "documents": [',
    "    {",
    '      "id": 0,',
    '      "name": "String",',
    '      "uploaded_on": "String",',
    '      "uploaded_by": "String",',
    '      "modified_at": "String",',
    '      "modified_by": "String",',
    '      "size": 0,',
    '      "type": "String",',
    '      "source": "String"',
    "    }",
    "  ]",
    "}",
]

DOCUMENT_DETAILS_BODY = [
    "{",
    '  "document_id": 0,',
    '  "document_name": "Name",',
    '  "document_order": 0,',
    '  "document_type": "PDF",',
    '  "document_width": 100,',
    '  "document_height": 100,',
    '  "document_pages": 1,',
    # The backend currently fills both of these from `document.getCreatedBy()`,
    # which is a defect the owner is having fixed. The specification documents
    # the intended meaning - a boolean flag - not the current wrong value.
    '  "form_field": true,',
    '  "lock_form_fields": true,',
    '  "created_at": "date",',
    '  "created_by": "String",',
    '  "modified_at": "date",',
    '  "modified_by": "String"',
    "}",
]

ALL_PACKAGES_BODY = [
    "{",
    '  "packages": [',
    "    {",
    '      "package_id": 123,',
    '      "package_name": "Package name",',
    '      "package_owner": "giatk@mobile-id.vn",',
    '      "owner_email": "giatk@mobile-id.vn",',
    '      "owner_name": "Tat Khanh Gia",',
    '      "package_status": "ACTIVE",',
    '      "next_signer_email": [',
    '        "giatk@mobile-id.vn"',
    "      ],",
    '      "created_at": "2023-07-21T11:12:13",',
    '      "created_by": "giatk@mobile-id.vn",',
    '      "modified_at": "2023-07-21T11:12:13",',
    '      "modified_by": "giatk@mobile-id.vn",',
    '      "access_duration": {',
    '        "date_from": "07212023",',
    '        "date_to": "07212024"',
    "      },",
    '      "decline": {',
    '        "reason": "For testing",',
    '        "declined_by": "giatk@mobile-id.vn"',
    "      }",
    "    }",
    "  ]",
    "}",
]

TEMPLATE_BODY = ["{}"]

# Response Attributes rows the backend never emits
DOCUMENT_DETAILS_DROP = ("document_source", "certify", "template")
DOCUMENT_DETAILS_RENAME = ("form_fields", "form_field")
# Apply Workflow Template returns no body, so only the error rows survive
TEMPLATE_KEEP = ("error_description", "error_code", "message")


def el_text(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def para_style(el):
    pPr = el.find(qn("pPr"))
    if pPr is None:
        return ""
    style = pPr.find(qn("pStyle"))
    return style.get(qn("val")) if style is not None else ""


def set_leaf_text(el, value):
    runs = [t for t in el.iter(qn("t"))]
    if not runs:
        return False
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""
    return True


def row_cells(tr):
    return tr.findall(qn("tc"))


def section_blocks(body, heading_text):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        sys.exit("section %r not found" % heading_text)
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if kids[j].tag == qn("p") and para_style(kids[j]) in HEADINGS:
            end = j
            break
    return kids[start:end]


def find_table(blocks, label):
    for el in blocks:
        if el.tag != qn("tbl"):
            continue
        rows = el.findall(qn("tr"))
        if rows and el_text(rows[0]).strip().startswith(label):
            return el
    return None


def data_rows(tbl):
    return [tr for tr in tbl.findall(qn("tr")) if len(row_cells(tr)) >= 5][1:]


def renumber(tbl):
    for n, tr in enumerate(data_rows(tbl), start=1):
        set_leaf_text(row_cells(tr)[0], str(n))


def success_frames(blocks):
    """Both copies of the 200 frame - Word keeps a DrawingML and a VML one."""
    out = []
    for el in blocks:
        if el.tag != qn("p"):
            continue
        for txbx in el.iter(qn("txbxContent")):
            paras = txbx.findall(qn("p"))
            if paras and el_text(paras[0]).strip().startswith("Http Code: 200"):
                out.append(txbx)
    return out


def rewrite_body(txbx, lines):
    """Replace the JSON body of a frame, keeping its header lines and styling."""
    paras = txbx.findall(qn("p"))
    body = [p for p in paras if not FRAME_HEADER.match(el_text(p).strip())]
    if not body:
        sys.exit("frame has no body lines to replace")
    template = copy.deepcopy(body[0])
    anchor = body[-1]
    for line in lines:
        new = copy.deepcopy(template)
        set_leaf_text(new, line)
        anchor.addprevious(new)
    for p in body:
        p.getparent().remove(p)


def fix_sample(body, heading, lines):
    frames = success_frames(section_blocks(body, heading))
    if len(frames) != 2:
        sys.exit("%r: expected 2 success frames, got %d" % (heading, len(frames)))
    for f in frames:
        rewrite_body(f, lines)
    return len(frames)


def main():
    dry = "--dry-run" in sys.argv
    tree = etree.parse(DOC_XML, etree.XMLParser(huge_tree=True))
    body = tree.getroot().find(qn("body"))

    # --- samples -----------------------------------------------------------
    for heading, lines in (
            ("Get Package Details", PACKAGE_DETAILS_BODY),
            ("Get Document Details", DOCUMENT_DETAILS_BODY),
            ("Get All Packages", ALL_PACKAGES_BODY),
            ("Apply Workflow Template", TEMPLATE_BODY)):
        n = fix_sample(body, heading, lines)
        print("sample rewritten: %-26s (%d frame copies, %d lines)"
              % (heading, n, len(lines)))

    # --- Get Document Details table ---------------------------------------
    tbl = find_table(section_blocks(body, "Get Document Details"), "Response Attributes")
    dropped = 0
    for tr in list(data_rows(tbl)):
        name = el_text(row_cells(tr)[1]).strip()
        if name in DOCUMENT_DETAILS_DROP:
            tbl.remove(tr)
            dropped += 1
        elif name == DOCUMENT_DETAILS_RENAME[0]:
            set_leaf_text(row_cells(tr)[1], DOCUMENT_DETAILS_RENAME[1])
    renumber(tbl)
    print("Get Document Details table: dropped %d row(s), %r -> %r, %d rows left"
          % (dropped, *DOCUMENT_DETAILS_RENAME, len(data_rows(tbl))))

    # --- Apply Workflow Template table + Content-Type ----------------------
    blocks = section_blocks(body, "Apply Workflow Template")
    tbl = find_table(blocks, "Response Attributes")
    dropped = 0
    for tr in list(data_rows(tbl)):
        if el_text(row_cells(tr)[1]).strip() not in TEMPLATE_KEEP:
            tbl.remove(tr)
            dropped += 1
    renumber(tbl)
    print("Apply Workflow Template table: dropped %d row(s), %d rows left"
          % (dropped, len(data_rows(tbl))))

    header_tbl = find_table(blocks, "Header Request Attributes")
    fixed = False
    for tr in data_rows(header_tbl):
        if el_text(row_cells(tr)[1]).strip() == "Content-Type":
            set_leaf_text(row_cells(tr)[4], CONTENT_TYPE_TEXT)
            fixed = True
    print("Apply Workflow Template Content-Type description: %s"
          % ("updated" if fixed else "ROW NOT FOUND"))
    if not fixed:
        sys.exit("Content-Type row missing")

    if dry:
        print("\n--dry-run: nothing written")
        return

    if os.path.exists(OUT_DOCX) and not os.path.exists(BACKUP_DOCX):
        shutil.copy2(OUT_DOCX, BACKUP_DOCX)
        print("\nbacked up %s" % BACKUP_DOCX)

    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


def pack(src_dir, out_path):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(src_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for base, _dirs, files in os.walk(src_dir):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                if rel != "[Content_Types].xml":
                    z.write(full, rel)
    shutil.move(tmp, out_path)


if __name__ == "__main__":
    main()
