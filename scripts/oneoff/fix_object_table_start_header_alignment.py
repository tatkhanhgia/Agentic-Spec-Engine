#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Left-align the merged start header row of every attribute table in the object
chapters (and any other table whose title row was still centred).

Template 2.4 requires left + middle for merged start headers such as
`Request Attributes` and object names like `AddUserToWorkflowRequest`.
`fix_docx_format.py` step 3 centres every `tblHeader` row; step 6 reset only
the six endpoint title strings, so 105 object-chapter tables stayed centred.

Formatting only: no cell text changes, no row is added or removed.

Usage:
    python scripts/oneoff/fix_object_table_start_header_alignment.py [--dry-run]
"""
import os
import sys

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

COLUMN_HEADER = ("No", "Name", "Type", "Presence", "Description")


def left_align(tc):
    changed = False
    for p in tc.findall(st.qn("p")):
        pPr = p.find(st.qn("pPr"))
        if pPr is None:
            pPr = etree.Element(st.qn("pPr"))
            p.insert(0, pPr)
        jc = pPr.find(st.qn("jc"))
        if jc is None:
            jc = etree.SubElement(pPr, st.qn("jc"))
        if jc.get(st.qn("val")) != "left":
            jc.set(st.qn("val"), "left")
            changed = True
    return changed


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()

    tables = 0
    for tbl in root.iter(st.qn("tbl")):
        rows = tbl.findall(st.qn("tr"))
        if len(rows) < 2:
            continue
        header = tuple(
            st.cell_text(tc).strip()
            for tc in rows[1].findall(st.qn("tc"))[:len(COLUMN_HEADER)]
        )
        if header != COLUMN_HEADER:
            continue
        cells = rows[0].findall(st.qn("tc"))
        if not cells:
            continue
        if left_align(cells[0]):
            tables += 1

    print("start header rows left-aligned: %d" % tables)
    if not tables:
        print("nothing to change")
        return
    if dry:
        print("\n--dry-run: nothing written")
        return
    try:
        guarded_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("\nREFUSED TO WRITE\n%s" % ex)
    st.pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
