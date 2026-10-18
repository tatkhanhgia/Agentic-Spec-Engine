#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Restore the one-value-per-line layout of "Possible values are:" cells.

update_register_certificate_processing.py wrote three enum-list descriptions as a
single flattened paragraph, but the original level_of_assurance cell (and the
house style in this category) puts each `+ VALUE` on its own paragraph.

Usage:
    python scripts/oneoff/fix_multiline_value_cells.py [--dry-run]
"""
import os
import sys
from copy import deepcopy

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

qn = st.qn


def set_cell_lines(tc, lines):
    """One paragraph per line, all cloned from the cell's first paragraph."""
    st.set_cell_text(tc, lines[0])
    anchor = tc.findall(qn("p"))[0]
    for line in lines[1:]:
        new_p = deepcopy(anchor)
        st.set_para_text(new_p, line)
        anchor.addnext(new_p)
        anchor = new_p


def desc_cell(kids, section_title, table_title, row_name):
    a, b = st.section_range(kids, st.find_heading(kids, 3, section_title))
    tbl = st.find_table(kids[a:b], table_title)
    for tr in st.data_rows(tbl):
        tcs = tr.findall(qn("tc"))
        if st.cell_text(tcs[1]) == row_name:
            return tcs[4]
    raise LookupError("%s / %s" % (section_title, row_name))


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    body = tree.getroot().find(qn("body"))
    kids = list(body)

    set_cell_lines(
        desc_cell(kids, "Register Certificate via RSSP", "Request Attributes",
                  "level_of_assurance"),
        ["Level of Assurance that the user wants to create. Possible values are:",
         "+ ADVANCED_ELECTRONIC_SIGNATURE",
         "+ ELECTRONIC_SEAL",
         "+ QUALIFIED_ELECTRONIC_SEAL",
         "+ ADVANCED_ELECTRONIC",
         "+ QUALIFIED_ELECTRONIC_SIGNATURE"])
    set_cell_lines(
        desc_cell(kids, "Finish Register Certificate", "Request Attributes",
                  "auth_mode"),
        ["Authorization mode of the credential. Possible values are:",
         "+ EXPLICIT/PIN",
         "+ EXPLICIT/OTP-SMS",
         "+ EXPLICIT/OTP-EMAIL",
         "+ IMPLICIT/TSE",
         "+ EXPLICIT/JWT"])
    set_cell_lines(
        desc_cell(kids, "Finish Register Certificate", "Request Attributes",
                  "operation_mode"),
        ["Operation mode. Possible values are:",
         "+ A (asynchronous, default)",
         "+ S (synchronous)",
         "+ C"])
    print("3 cells rewritten as one value per line")

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
