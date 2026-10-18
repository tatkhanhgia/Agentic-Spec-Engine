#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Repair the composite-type anchors and orphans belonging to Document Preparation.

Two kinds of defect, both scoped to this module - the other broken anchors live
in Document Processing (`Sign Document`) and Personal Settings (`Get Address
Book`) and are left for their owners.

**Dangling anchors.** The document decides each of these the same way: the same
field is linked correctly in most cells and wrongly in a few, so the majority
spelling is the answer, not a guess.

| Field | Correct anchor | Cells | Broken anchor | Cells |
|-------|----------------|-------|---------------|-------|
| `authentication_signing` | `_Authentication_Signing_(for_Workflo_1` | 4 | `_FieldAuthentication` | 4 |
| `dimension` | `_Dimension_1` | 38 | `_Dimension` | 2 |

Only the `w:anchor` attribute changes; the label and every run property stay as
they are, because per CLAUDE.md the label is only a display string and must not
be "corrected" to match the object heading.

**Orphans left by the Qrypto withdrawal.** Removing the two Qrypto endpoints on
2026-07-27 cut the only inbound link to the `Items` object (`field`, `type`,
`value`, `remark`, `mandatory_enabled` - the Qrypto item shape), and `Items` was
in turn the only thing linking to the `Items request` example. Both are now
unreachable and go with it.

The `Items request` section is the last one in the body and carries the body-level
`w:sectPr`. Removing the section wholesale would take the page setup with it, so
the `sectPr` is moved out first and re-attached at the end of the body.

Usage:
    python scripts/oneoff/fix_document_preparation_anchors.py [--dry-run]
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

RETARGET = {
    "_FieldAuthentication": "_Authentication_Signing_(for_Workflo_1",
    "_Dimension": "_Dimension_1",
}

DROP_OBJECTS = [("Items", 2), ("Items request", 2)]      # (heading text, heading level)


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    for old, new in RETARGET.items():
        if new not in names:
            sys.exit("target bookmark %r does not exist - refusing to retarget" % new)

    # ---- 1. retarget dangling anchors -------------------------------------
    counts = {}
    for link in root.iter(st.qn("hyperlink")):
        a = link.get(st.qn("anchor"))
        if a in RETARGET:
            link.set(st.qn("anchor"), RETARGET[a])
            counts[a] = counts.get(a, 0) + 1
    for old, new in RETARGET.items():
        print("  %-22s -> %-40s %d link" % (old, new, counts.get(old, 0)))

    # ---- 2. drop the objects orphaned by the Qrypto withdrawal ------------
    for title, level in DROP_OBJECTS:
        kids = list(body)
        try:
            idx = st.find_heading(kids, level, title)
        except LookupError:
            print("  object %-16s not found - already removed?" % title)
            continue
        start, end = st.section_range(kids, idx)
        block = kids[start:end]

        # the body-level sectPr must survive the removal
        rescued = None
        for el in block:
            if st.local(el) == "sectPr":
                rescued = el
                block = [b for b in block if b is not el]
                break

        starts = {b.get(st.qn("id")) for e in block for b in e.iter(st.qn("bookmarkStart"))}
        ends = {b.get(st.qn("id")) for e in block for b in e.iter(st.qn("bookmarkEnd"))}
        for el in block:
            body.remove(el)
        for be in list(root.iter(st.qn("bookmarkEnd"))):
            if be.get(st.qn("id")) in (starts - ends):
                be.getparent().remove(be)

        if rescued is not None:
            body.remove(rescued)
            body.append(rescued)
            print("  object %-16s removed %d elements (sectPr preserved)" % (title, len(block)))
        else:
            print("  object %-16s removed %d elements" % (title, len(block)))

    # ---- 3. report ---------------------------------------------------------
    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    dangling = sorted({h.get(st.qn("anchor")) for h in root.iter(st.qn("hyperlink"))
                       if h.get(st.qn("anchor")) and h.get(st.qn("anchor")) not in names})
    print("  anchors still dangling document-wide: %d %s" % (len(dangling), dangling))
    last = list(body)[-1]
    print("  last body child is sectPr: %s" % (st.local(last) == "sectPr"))

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
