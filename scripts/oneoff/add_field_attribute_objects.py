#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Write the eleven missing field-attribute objects and link `Get Document Fields` to them.

`Get Document Fields` returns one array per field type. Five of its rows link to an
object (`TextFieldAttribute`, `CheckboxFieldAttribute`, `QRFieldAttribute`,
`InitialFieldAttribute`, `SignatureFieldAttribute`); eleven name an object that was
never written, so those Type cells carry a dead label.

The content does not have to be invented, because the document already describes
every one of these shapes in the field type's own `Add X Field` section. The five
existing objects follow one recipe, verified against `TextFieldAttribute`:

    object rows = the `Add X Field` Request Attributes rows
                + process_status + processed_on + processed_by

`TextFieldAttribute` has 21 rows and `Add TextBox Field` has 18, with identical
names, and the three extra rows are exactly the server-set ones. Two editorial
normalisations are visible in that comparison and are reproduced here:

  * `field_name` leads, then `page` - the Add tables put `page` first;
  * a request-only presence of `RC(...)` becomes `O`, because in a response the
    attribute is simply present or absent.

Rows are deep-copied whole rather than rebuilt, so a Type cell that already links
to `dimension`, `font` or `alignmentOption` keeps its `w:hyperlink` intact - the
failure this whole exercise started from.

The shell (heading, bookmark, table styling and the three process rows) is cloned
from `TextFieldAttribute` itself, so the new objects are siblings of the existing
five in every respect, including the convention that the table's title row carries
the object name rather than the words "Response Attributes".

Usage:
    python scripts/oneoff/add_field_attribute_objects.py [--dry-run]
"""
import os
import re
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

SHELL_OBJECT = "TextFieldAttribute"
SHELL_ANCHOR = "_TextFieldAttribute"
LAST_OBJECT = "FieldTypeItem"          # new objects are appended after this one
PROCESS_ROWS = ("process_status", "processed_on", "processed_by")

# (Get Document Fields row, object heading, anchor, source Add section)
SPECS = [
    ("radiobox",   "RadioboxFieldAttribute",    "_RadioboxFieldAttribute",    "Add RadioBox Field"),
    ("combobox",   "ComboBoxFieldAttribute",    "_ComboBoxFieldAttribute",    "Add ComboBox Field"),
    ("datetime",   "DateTimeFieldAttribute",    "_DateTimeFieldAttribute",    "Add Datetime Field"),
    ("stamp",      "ImageFieldAttribute",       "_ImageFieldAttribute",       "Add Stamp Field"),
    ("camera",     "CameraFieldAttribute",      "_CameraFieldAttribute",      "Add Camera Field"),
    ("attachment", "AttachmentFieldAttribute",  "_AttachmentFieldAttribute",  "Add Attachment Field"),
    ("hyperlink",  "HyperLinkFieldAttribute",   "_HyperLinkFieldAttribute",   "Add Hyperlink Field"),
    ("toggle",     "ToggleFieldAttribute",      "_ToggleFieldAttribute",      "Add Toggle Field"),
    ("stepper",    "NumericFieldAttribute",     "_NumericFieldAttribute",     "Add Numeric Stepper Field"),
    ("checkboxV2", "CheckboxFieldAttributeV2",  "_CheckboxFieldAttributeV2",
     "Add CheckBox Field Version 2"),
    ("radioboxV2", "RadioboxFieldAttributeV2",  "_RadioboxFieldAttributeV2",
     "Add RadioBox Field Version 2"),
]


def set_table_title(tbl, text):
    trs = tbl.findall(st.qn("tr"))
    st.set_cell_text(trs[0].findall(st.qn("tc"))[0], text)


def source_rows(kids, section):
    s, e = st.section_range(kids, st.find_heading(kids, 3, section))
    tbl = st.find_table(kids[s:e], "Request Attributes")
    rows = [deepcopy(tr) for tr in st.data_rows(tbl)]

    # request-only presence has no meaning in a response
    for tr in rows:
        tcs = tr.findall(st.qn("tc"))
        if st.cell_text(tcs[3]).startswith("RC("):
            st.set_cell_text(tcs[3], "O")

    # the existing objects lead with field_name; the Add tables lead with page
    names = [st.cell_text(tr.findall(st.qn("tc"))[1]) for tr in rows]
    if len(names) >= 2 and names[0] == "page" and names[1] == "field_name":
        rows[0], rows[1] = rows[1], rows[0]
    return rows


def build(root, kids, heading, anchor, section, bm_id):
    clone = [deepcopy(e) for e in kids[slice(*st.section_range(
        kids, st.find_heading(kids, 2, SHELL_OBJECT)))]]

    drop = set()
    for el in clone:
        for bs in list(el.iter(st.qn("bookmarkStart"))):
            if bs.get(st.qn("name")) == SHELL_ANCHOR:
                old = bs.get(st.qn("id"))
                bs.set(st.qn("name"), anchor)
                bs.set(st.qn("id"), str(bm_id))
                for be in el.iter(st.qn("bookmarkEnd")):
                    if be.get(st.qn("id")) == old:
                        be.set(st.qn("id"), str(bm_id))
            else:
                drop.add(bs.get(st.qn("id")))
                bs.getparent().remove(bs)
    for el in clone:
        for be in list(el.iter(st.qn("bookmarkEnd"))):
            if be.get(st.qn("id")) in drop:
                be.getparent().remove(be)

    st.set_text_keep_bookmarks(clone[0], heading)

    tbl = next(e for e in clone if e.tag == st.qn("tbl"))
    set_table_title(tbl, heading)

    rows = st.data_rows(tbl)
    keep = [tr for tr in rows
            if st.cell_text(tr.findall(st.qn("tc"))[1]) in PROCESS_ROWS]
    if len(keep) != len(PROCESS_ROWS):
        raise RuntimeError("shell object lost its process_* rows")
    for tr in rows:
        if tr not in keep:
            tr.getparent().remove(tr)

    anchor_row = keep[0]
    for tr in reversed(source_rows(kids, section)):
        anchor_row.addprevious(tr)
    st.renumber(tbl)

    n = len(st.data_rows(tbl))
    print("  %-26s <- %-30s %2d rows (%d + 3)" % (heading, section, n, n - 3))
    return clone


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))
    kids = list(body)

    for _f, heading, _a, _s in SPECS:
        try:
            st.find_heading(kids, 2, heading)
        except LookupError:
            pass
        else:
            sys.exit("object %r already exists - nothing to do" % heading)

    bm_id = st.next_bookmark_id(root)
    built = []
    for _field, heading, anchor, section in SPECS:
        built.extend(build(root, kids, heading, anchor, section, bm_id))
        bm_id += 1

    st.insert_after_section(body, kids, LAST_OBJECT, built, level=2)
    print("inserted %d elements after object %r" % (len(built), LAST_OBJECT))

    # link the rows of Get Document Fields
    kids = list(body)
    s, e = st.section_range(kids, st.find_heading(kids, 3, "Get Document Fields"))
    tbl = st.find_table(kids[s:e], "Response Attributes")
    donor = None
    for tr in st.data_rows(tbl):
        tcs = tr.findall(st.qn("tc"))
        if st.cell_text(tcs[1]) == "textbox":
            donor = tcs[2]
            break
    if donor is None:
        sys.exit("no linked donor row in Get Document Fields")

    linked = 0
    by_field = {f: a for f, _h, a, _s in SPECS}
    for tr in st.data_rows(tbl):
        tcs = tr.findall(st.qn("tc"))
        field = st.cell_text(tcs[1])
        if field not in by_field:
            continue
        st.set_type_link(tcs[2], st.cell_text(tcs[2]), by_field[field], donor)
        linked += 1
    print("linked %d rows of Get Document Fields" % linked)

    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    missing = [a for _f, _h, a, _s in SPECS if a not in names]
    print("all new anchors resolve: %s" % ("yes" if not missing else "NO -> %s" % missing))

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
