#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Remove the Qrypto field API from the specification.

The project owner decided on 2026-07-27 that Qrypto is not going to be
implemented and must leave the document, explicitly regardless of what the
Postman collection and the backend still carry. This is therefore a
documentation-only removal: the backend registry is untouched and no drift
report should later "restore" these sections - see the note added to section 5
of the format template.

Three places carry Qrypto:

  * `Add Qrypto Field` and `Update Qrypto Field` - the two endpoint sections;
  * `Get Document Fields` - one Response Attributes row,
    `qrypto | List<qryptoFieldAttribute>`, which describes a field type that will
    no longer exist;
  * the Table of Contents, which Word rebuilds on open because `w:updateFields`
    is set in settings.xml.

Nothing else links to the two sections: the only inbound hyperlinks are the
`_Toc...` entries. The shared `Items request` example under the `Example`
heading is deliberately left in place - `List<items>` is Qrypto-only, but the
example is reached from a separate `Items` section and another endpoint added in
this same period (`Import Excel to Clone Workflow`) also has a field named
`items`, so removing it is a separate decision for the owner.

Usage:
    python scripts/oneoff/remove_qrypto_sections.py [--dry-run]
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

SECTIONS = ("Add Qrypto Field", "Update Qrypto Field")
ROW_HOST = "Get Document Fields"
ROW_NAME = "qrypto"


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    removed_ids = set()
    total = 0
    for title in SECTIONS:
        kids = list(body)                      # re-read: indices shift after each removal
        try:
            idx = st.find_heading(kids, 3, title)
        except LookupError:
            print("  %-20s not found - already removed?" % title)
            continue
        start, end = st.section_range(kids, idx)
        block = kids[start:end]

        starts = {b.get(st.qn("id")) for e in block for b in e.iter(st.qn("bookmarkStart"))}
        ends = {b.get(st.qn("id")) for e in block for b in e.iter(st.qn("bookmarkEnd"))}
        removed_ids |= (starts - ends)         # ends that live outside the block

        for el in block:
            body.remove(el)
        total += len(block)
        print("  %-20s removed %d elements (%d p, %d tbl)"
              % (title, len(block),
                 sum(1 for e in block if e.tag == st.qn("p")),
                 sum(1 for e in block if e.tag == st.qn("tbl"))))

    # a bookmarkEnd whose start went away would dangle
    orphans = 0
    if removed_ids:
        for be in list(root.iter(st.qn("bookmarkEnd"))):
            if be.get(st.qn("id")) in removed_ids:
                be.getparent().remove(be)
                orphans += 1
    print("  orphaned bookmarkEnd removed: %d" % orphans)

    # the response row of Get Document Fields
    kids = list(body)
    s, e = st.section_range(kids, st.find_heading(kids, 3, ROW_HOST))
    tbl = st.find_table(kids[s:e], "Response Attributes")
    before = len(st.data_rows(tbl))
    try:
        st.drop_row(tbl, ROW_NAME)
        print("  %s: dropped row %r, rows %d -> %d"
              % (ROW_HOST, ROW_NAME, before, len(st.data_rows(tbl))))
    except LookupError:
        print("  %s: row %r already gone" % (ROW_HOST, ROW_NAME))

    left = sum(1 for t in root.iter(st.qn("t")) if t.text and "qrypto" in t.text.lower())
    print("  runs still mentioning qrypto: %d  (Items request example is kept on purpose)" % left)

    if not total:
        print("nothing removed")
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
