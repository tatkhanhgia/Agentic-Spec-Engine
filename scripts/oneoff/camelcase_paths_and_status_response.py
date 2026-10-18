#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Two user decisions from the 2026-07-27 Register Certificate Processing session.

1. Path placeholders in this category used `{session_id}` while the document-wide
   convention (and template 2.8.9) is camelCase for URL parameters (packageId x111,
   documentId x69). Renamed to `{sessionId}` in the four request frames and their
   Path Attributes tables. JSON body keys named session_id stay snake_case - only
   the URL placeholder and the path-table row are renamed.

2. Get Credential Registration Status: owner decided the success response is
   documented in snake_case per 2.8.9 (request_type, citizen_pid, credential_id,
   serial_number, certificate_data) even though the backend currently emits
   camelCase through the default Jackson mapper. Backend fix tracked in
   TECHNICAL_DEBT.md section 7; doc-side protection in template 5.2.

Usage:
    python scripts/oneoff/camelcase_paths_and_status_response.py [--dry-run]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

qn = st.qn

SESSION_PATH_SECTIONS = [
    ("Send OTP for Certificate Registration", "POST"),
    ("Verify OTP for Certificate Registration", "POST"),
    ("Get List Tax Codes", "GET"),
    ("Finish Register Certificate", "POST"),
]


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    body = tree.getroot().find(qn("body"))
    kids = list(body)

    # ---- 1. {session_id} -> {sessionId} ----------------------------------
    for title, method in SESSION_PATH_SECTIONS:
        a, b = st.section_range(kids, st.find_heading(kids, 3, title))
        els = kids[a:b]
        n = 0
        for _txbx, paras in st.each_frame(els, method + " "):
            n += st.set_path_suffix(paras[0], "{session_id}", "{sessionId}")
        tbl = st.find_table(els, "Path Attributes")
        st.set_row(tbl, "session_id", new_name="sessionId")
        print("%-42s frame runs=%d, path row renamed" % (title, n))

    # ---- 2. Status response in snake_case --------------------------------
    a, b = st.section_range(
        kids, st.find_heading(kids, 3, "Get Credential Registration Status"))
    els = kids[a:b]
    tbl = st.find_table(els, "Response Attributes")
    st.set_row(tbl, "data", description=(
        "Details of the transaction. Contains request_type, citizen_pid, "
        "credential_id, serial_number and certificate_data."))
    for txbx, paras in st.each_frame(els, "Http Code: 200"):
        st.replace_frame_body(txbx, txbx.findall(qn("p")), [
            "{",
            '  "status": "",',
            '  "data": {',
            '    "request_type": "",',
            '    "citizen_pid": "",',
            '    "credential_id": "",',
            '    "serial_number": "",',
            '    "certificate_data": ""',
            "  }",
            "}",
        ])
    print("Get Credential Registration Status: data keys documented in snake_case")

    if dry:
        print("--dry-run: nothing written")
        return
    try:
        guarded_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("\nREFUSED TO WRITE\n%s" % ex)
    st.pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
