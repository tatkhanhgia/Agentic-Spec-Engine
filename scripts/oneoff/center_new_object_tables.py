#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Centre the No and Presence columns of the objects added on 2026-07-27.

Template 2.4 requires both columns centered in every attribute table. The eleven
field-attribute objects were cloned from `TextFieldAttribute`, whose own table is
left-aligned in all 21 rows because it predates the rule. The rows copied in from
the `Add X Field` sections were already normalised, so only the three inherited
`process_status` / `processed_on` / `processed_by` rows came across left-aligned -
33 rows in total.

`scripts/check_frame_format.py` does not catch them: it matches attribute tables by
title, and an object's table is titled with the object name rather than
"Response Attributes", so the whole object chapter is outside its coverage. That
gap is reported to the project owner rather than papered over here.

Scope is the objects written in this session. The five pre-existing
field-attribute objects (`TextFieldAttribute`, `CheckboxFieldAttribute`,
`QRFieldAttribute`, `InitialFieldAttribute`, `SignatureFieldAttribute`) are
left-aligned in every row and are left untouched - normalising them is a separate
decision.

Formatting only: no cell text changes, no row is added or removed.

Usage:
    python scripts/oneoff/center_new_object_tables.py [--dry-run]
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

OBJECTS = [
    "RedactPolicy", "RedactionRule", "FieldTypeGroup", "FieldTypeItem",
    "RadioboxFieldAttribute", "ComboBoxFieldAttribute", "DateTimeFieldAttribute",
    "ImageFieldAttribute", "CameraFieldAttribute", "AttachmentFieldAttribute",
    "HyperLinkFieldAttribute", "ToggleFieldAttribute", "NumericFieldAttribute",
    "CheckboxFieldAttributeV2", "RadioboxFieldAttributeV2",
]
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

    per_object = {}
    total = 0
    for name in OBJECTS:
        s, e = st.section_range(kids, st.find_heading(kids, 2, name))
        n = 0
        for el in kids[s:e]:
            if el.tag != st.qn("tbl"):
                continue
            for tr in st.data_rows(el):
                tcs = tr.findall(st.qn("tc"))
                if len(tcs) < 5:
                    continue
                touched = False
                for col, _label in CENTER_COLUMNS:
                    if centre(tcs[col]):
                        touched = True
                if touched:
                    n += 1
        if n:
            per_object[name] = n
            total += n

    for name, n in per_object.items():
        print("  %-26s %d row" % (name, n))
    print("rows centered: %d" % total)
    if not total:
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
