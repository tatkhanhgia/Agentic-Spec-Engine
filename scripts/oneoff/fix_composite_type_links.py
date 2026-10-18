#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Give the composite Type cells of the new sections real object links.

CLAUDE.md ("Composite Types Link by Bookmark, Not by Name") records that a Type
cell holding a non-primitive is a `w:hyperlink` whose `w:anchor` points at a
bookmark on the object's Heading 2, and that an endpoint is not finished until
the objects behind its composite types exist. Three cells added on 2026-07-27
failed that:

  * `Get Field Types` / `field_types` - the donor cell (`List<Category>`) carried
    a link to the `Category` object; `set_cell_text` dropped every direct child of
    the paragraph, so the `w:hyperlink` went with it and the label was left as
    dead text.
  * `Add Redact Field` / `redact_policy` and `Update Redact Field` /
    `redact_policy` - written over a primitive donor cell (`is_transparent` /
    `Boolean`), so there was no link to inherit, and the objects they name did
    not exist at all.

This script creates the four missing objects at the end of the RELATIVE OBJECT
chapter and points all three cells - plus the composite cells inside the new
objects - at them.

Objects are built by cloning `Dimension`, the smallest existing object section
(Heading 2 + one table + a blank paragraph). Cloning brings the bookmark along:
the `_Dimension_1` bookmark is renamed to the new anchor and given a fresh id,
while the `_Toc...` bookmark is dropped because Word regenerates the table of
contents on open.

Field names and enum values come from the backend:
  * `RedactFieldAttribute.PolicyContainer` -> creator / others, each a
    `pre_process` / `post_process` of enum `RedactionPolicyType`
    (SHOW, REDACT, SHOW_ALL, REDACT_ALL); `others` also carries `remark`.
  * `GetFieldTypesSerializer` -> field_types[] -> {group, items[]} and
    items[] -> {id, type_name, type_id, parent_type, required, ordinary_id,
    remark, children[]}, with group values from the `FieldTypeGroup` enum
    (COMMON, SYSTEM, INDIVIDUAL).

Usage:
    python scripts/oneoff/fix_composite_type_links.py [--dry-run]
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

DONOR_OBJECT = "Dimension"
DONOR_ANCHOR = "_Dimension_1"
LAST_OBJECT = "IssuingProvider"          # new objects go after this one

RULE_VALUES = "Possible values are SHOW, REDACT, SHOW_ALL and REDACT_ALL."

# (heading, anchor, [(name, type, presence, description, link_anchor_or_None)])
OBJECTS = [
    ("RedactPolicy", "_RedactPolicy", [
        ("creator", "RedactionRule", "O",
         "Redaction behaviour applied to the owner of the workflow.", "_RedactionRule"),
        ("others", "RedactionRule", "O",
         "Redaction behaviour applied to every other participant.", "_RedactionRule"),
    ]),
    ("RedactionRule", "_RedactionRule", [
        ("pre_process", "String", "O",
         "Behaviour applied before the field is processed. " + RULE_VALUES, None),
        ("post_process", "String", "O",
         "Behaviour applied after the field is processed. " + RULE_VALUES, None),
        ("remark", "Object", "O", "Remark in your application.", None),
    ]),
    ("FieldTypeGroup", "_FieldTypeGroup", [
        ("group", "String", "M",
         "Name of the group. Possible values are COMMON, SYSTEM and INDIVIDUAL.", None),
        ("items", "List<FieldTypeItem>", "M",
         "Field types that belong to the group.", "_FieldTypeItem"),
    ]),
    ("FieldTypeItem", "_FieldTypeItem", [
        ("id", "String", "M", "Identifier of the field type.", None),
        ("type_name", "String", "M", "Name of the field type.", None),
        ("type_id", "Integer", "M", "Numeric identifier of the field type.", None),
        ("parent_type", "String", "M", "Name of the parent field type.", None),
        ("required", "Boolean", "M", "True if the field type is required.", None),
        ("ordinary_id", "Integer", "O",
         "Order used to sort the field types for display.", None),
        ("remark", "String", "O",
         "Remark of the field type, in the language requested by x-language-name.", None),
        ("children", "List<FieldTypeItem>", "O",
         "Field types nested under this one. Each element has the same shape.",
         "_FieldTypeItem"),
    ]),
]

# (endpoint section, table, field, label, anchor)
CELL_LINKS = [
    ("Add Redact Field", "Request Attributes", "redact_policy", "redactPolicy", "_RedactPolicy"),
    ("Update Redact Field", "Request Attributes", "redact_policy", "redactPolicy", "_RedactPolicy"),
    ("Get Field Types", "Response Attributes", "field_types",
     "List<fieldTypeGroup>", "_FieldTypeGroup"),
]


def find_donor_type_cell(kids):
    """A Type cell that already links correctly, used as the clone source."""
    s, e = st.section_range(kids, st.find_heading(kids, 3, "Add VNeID Field"))
    tbl = st.find_table(kids[s:e], "Request Attributes")
    for tr in st.data_rows(tbl):
        tcs = tr.findall(st.qn("tc"))
        if st.cell_text(tcs[1]) == "dimension":
            return tcs[2]
    raise LookupError("no linked donor Type cell found")


def build_object(root, kids, heading, anchor, rows, donor_cell, bm_id):
    clone = [deepcopy(e) for e in kids[slice(*st.section_range(
        kids, st.find_heading(kids, 2, DONOR_OBJECT)))]]

    # bookmarks: keep the object anchor, drop the TOC one (Word rebuilds it)
    drop_ids = set()
    for el in clone:
        for bs in list(el.iter(st.qn("bookmarkStart"))):
            if bs.get(st.qn("name")) == DONOR_ANCHOR:
                bs.set(st.qn("name"), anchor)
                old = bs.get(st.qn("id"))
                bs.set(st.qn("id"), str(bm_id))
                for be in el.iter(st.qn("bookmarkEnd")):
                    if be.get(st.qn("id")) == old:
                        be.set(st.qn("id"), str(bm_id))
            else:
                drop_ids.add(bs.get(st.qn("id")))
                bs.getparent().remove(bs)
    for el in clone:
        for be in list(el.iter(st.qn("bookmarkEnd"))):
            if be.get(st.qn("id")) in drop_ids:
                be.getparent().remove(be)

    # heading text - must not disturb the bookmark that lives on this paragraph
    st.set_text_keep_bookmarks(clone[0], heading)

    tbl = st.find_table(clone, "Response Attributes")
    have = len(st.data_rows(tbl))
    while len(st.data_rows(tbl)) < len(rows):
        last = st.data_rows(tbl)[-1]
        tcs = last.findall(st.qn("tc"))
        st.append_row(tbl, st.cell_text(tcs[1]), st.cell_text(tcs[2]),
                      st.cell_text(tcs[3]), st.cell_text(tcs[4]))
    while len(st.data_rows(tbl)) > len(rows):
        extra = st.data_rows(tbl)[-1]
        extra.getparent().remove(extra)
    st.renumber(tbl)

    for tr, (name, typ, presence, desc, link) in zip(st.data_rows(tbl), rows):
        tcs = tr.findall(st.qn("tc"))
        st.set_cell_text(tcs[1], name)
        st.set_cell_text(tcs[3], presence)
        st.set_cell_text(tcs[4], desc)
        if link:
            st.set_type_link(tcs[2], typ, link, donor_cell)
        else:
            st.set_cell_text(tcs[2], typ)

    print("  object %-16s anchor=%-18s rows %d -> %d" % (heading, anchor, have, len(rows)))
    return clone


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))
    kids = list(body)

    for heading, _a, _r in OBJECTS:
        try:
            st.find_heading(kids, 2, heading)
        except LookupError:
            pass
        else:
            sys.exit("object %r already exists - nothing to do" % heading)

    donor_cell = find_donor_type_cell(kids)
    bm_id = st.next_bookmark_id(root)

    built = []
    for heading, anchor, rows in OBJECTS:
        built.extend(build_object(root, kids, heading, anchor, rows, donor_cell, bm_id))
        bm_id += 1

    st.insert_after_section(body, kids, LAST_OBJECT, built, level=2)
    print("inserted %d elements after object %r" % (len(built), LAST_OBJECT))

    # now the three endpoint cells
    kids = list(body)
    for section, table, field, label, anchor in CELL_LINKS:
        s, e = st.section_range(kids, st.find_heading(kids, 3, section))
        tbl = st.find_table(kids[s:e], table)
        for tr in st.data_rows(tbl):
            tcs = tr.findall(st.qn("tc"))
            if st.cell_text(tcs[1]) == field:
                st.set_type_link(tcs[2], label, anchor, donor_cell)
                print("  linked %-22s %-14s -> %s" % (section, field, anchor))
                break
        else:
            sys.exit("row %r not found in %r of %r" % (field, table, section))

    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    missing = [a for _h, a, _r in OBJECTS if a not in names]
    print("new anchors resolvable: %s" % ("yes" if not missing else "NO -> %s" % missing))

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
