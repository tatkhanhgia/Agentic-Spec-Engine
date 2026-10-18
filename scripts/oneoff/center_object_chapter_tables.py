#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Centre No and Presence across the object chapters.

Template 2.4 requires both columns centered in every attribute table, and an
attribute table is recognised by its `No | Name | Type | Presence | Description`
header row rather than by its title. `check_frame_format.py` used to match on the
title instead, so the tables in `REQUEST/RESPONSE OBJECT` and `RELATIVE OBJECT` -
which title that row with the object's own name - were never checked and drifted:
309 rows across 47 tables are left-aligned.

Scope is those two chapters only. The 50 remaining rows sit in endpoint sections
under `API Specification` and belong to work another agent is doing right now; the
project owner asked for them to be left alone.

Formatting only: no cell text changes, no row is added or removed.

Usage:
    python scripts/oneoff/center_object_chapter_tables.py [--dry-run]
"""
import collections
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

CHAPTERS = ("REQUEST/RESPONSE OBJECT", "RELATIVE OBJECT")
COLUMN_HEADER = ("No", "Name", "Type", "Presence", "Description")
CENTER_COLUMNS = ((0, "No"), (3, "Presence"))


def centre(tc):
    changed = False
    for p in tc.findall(st.qn("p")):
        pPr = p.find(st.qn("pPr"))
        if pPr is None:
            pPr = etree.Element(st.qn("pPr"))
            p.insert(0, pPr)
        jc = pPr.find(st.qn("jc"))
        if jc is None:
            jc = etree.SubElement(pPr, st.qn("jc"))
        if jc.get(st.qn("val")) != "center":
            jc.set(st.qn("val"), "center")
            changed = True
    return changed


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    kids = list(root.find(st.qn("body")))

    heads = [(i, st.text_of(e).strip()) for i, e in enumerate(kids)
             if e.tag == st.qn("p") and st.style_of(e) == "Heading1"]

    def chapter_of(idx):
        name = None
        for i, t in heads:
            if i <= idx:
                name = t
        return name

    per_chapter = collections.Counter()
    tables = 0
    rows = 0
    for i, el in enumerate(kids):
        if el.tag != st.qn("tbl"):
            continue
        if chapter_of(i) not in CHAPTERS:
            continue
        trs = el.findall(st.qn("tr"))
        if len(trs) < 3:
            continue
        header = tuple(st.cell_text(tc).strip()
                       for tc in trs[1].findall(st.qn("tc")))[:len(COLUMN_HEADER)]
        if header != COLUMN_HEADER:
            continue
        touched = False
        for tr in trs[2:]:
            tcs = tr.findall(st.qn("tc"))
            if len(tcs) < 5:
                continue
            hit = False
            for col, _label in CENTER_COLUMNS:
                if centre(tcs[col]):
                    hit = True
            if hit:
                rows += 1
                per_chapter[chapter_of(i)] += 1
                touched = True
        if touched:
            tables += 1

    for ch in CHAPTERS:
        print("  %-26s %d row" % (ch, per_chapter[ch]))
    print("tables touched: %d   rows centered: %d" % (tables, rows))
    if not rows:
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
