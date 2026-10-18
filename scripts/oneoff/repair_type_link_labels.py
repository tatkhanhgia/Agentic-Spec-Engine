#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Repair Type labels that were nested twice, and bring the rest to house style.

`section_tools.set_type_link` cloned a donor paragraph and then wrote the whole
label inside the `w:hyperlink`. That is right when the donor links its entire
label (`dimension`, `font`), but wrong when it does not: a `List<x>` Type cell is
a plain `List<` run, the hyperlink around `x`, then a plain `>`. Writing the full
label into the link produced `List<List<x>>` in twelve rows of
`Get Document Fields`.

The cause is fixed in `set_type_link`, which now splits the label against the
donor's surrounding runs. This script repairs the cells already written:

  * twelve rows of `Get Document Fields` whose label is doubled;
  * three cells written earlier from a donor with no wrapper
    (`Get Field Types` / `field_types`, `FieldTypeGroup` / `items`,
    `FieldTypeItem` / `children`), where the whole `List<...>` sits inside the
    hyperlink. They render as a fully underlined label instead of the house
    style, so they are re-done from a wrapped donor.

Every cell is rewritten from the intended full label, so the result is
independent of what the cell currently holds.

Usage:
    python scripts/oneoff/repair_type_link_labels.py [--dry-run]
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

# (heading level, section, table, field, intended label)
TARGETS = [
    (3, "Get Document Fields", "Response Attributes", "radiobox",   "List<radioboxFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "inperson",   "List<signatureFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "combobox",   "List<comboBoxFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "datetime",   "List<dateTimeFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "stamp",      "List<imageFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "camera",     "List<cameraFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "attachment", "List<attachmentFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "hyperlink",  "List<hyperLinkFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "toggle",     "List<toggleFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "stepper",    "List<numericFieldAttribute>"),
    (3, "Get Document Fields", "Response Attributes", "checkboxV2", "List<checkboxFieldAttributeV2>"),
    (3, "Get Document Fields", "Response Attributes", "radioboxV2", "List<radioboxFieldAttributeV2>"),
    (3, "Get Field Types", "Response Attributes", "field_types", "List<fieldTypeGroup>"),
    (2, "FieldTypeGroup", "Response Attributes", "items", "List<FieldTypeItem>"),
    (2, "FieldTypeItem", "Response Attributes", "children", "List<FieldTypeItem>"),
]


def wrapped_donor(kids):
    """A Type cell whose hyperlink covers only the inner name of a List<...>."""
    s, e = st.section_range(kids, st.find_heading(kids, 3, "Get Document Fields"))
    tbl = st.find_table(kids[s:e], "Response Attributes")
    for tr in st.data_rows(tbl):
        tcs = tr.findall(st.qn("tc"))
        if st.cell_text(tcs[1]) == "textbox":
            return tcs[2]
    raise LookupError("no wrapped donor cell found")


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))
    kids = list(body)

    donor = wrapped_donor(kids)
    d_txt = st.cell_text(donor)
    if not re.fullmatch(r"List<.+>", d_txt):
        sys.exit("donor is not a List<...> cell: %r" % d_txt)
    print("donor: %r" % d_txt)

    fixed = 0
    for level, section, table, field, label in TARGETS:
        s, e = st.section_range(kids, st.find_heading(kids, level, section))
        tbl = st.find_table(kids[s:e], table)
        for tr in st.data_rows(tbl):
            tcs = tr.findall(st.qn("tc"))
            if st.cell_text(tcs[1]) != field:
                continue
            anchors = [h.get(st.qn("anchor")) for h in tcs[2].iter(st.qn("hyperlink"))]
            if not anchors:
                sys.exit("%s / %s has no link to preserve" % (section, field))
            before = st.cell_text(tcs[2])
            st.set_type_link(tcs[2], label, anchors[0], donor)
            after = st.cell_text(tcs[2])
            if before != after:
                print("  %-20s %-12s %r -> %r" % (section[:20], field, before, after))
                fixed += 1
            break
        else:
            sys.exit("row %r not found in %r" % (field, section))

    print("cells changed: %d" % fixed)

    # nothing may be left doubled
    bad = []
    for tbl in root.iter(st.qn("tbl")):
        for tr in st.data_rows(tbl):
            tcs = tr.findall(st.qn("tc"))
            if len(tcs) < 5:
                continue
            t = st.cell_text(tcs[2])
            if "List<List<" in t:
                bad.append((st.cell_text(tcs[1]), t))
    print("labels still doubled: %d %s" % (len(bad), bad))

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
