#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Repair the duplicate drawing identifiers introduced by the 2026-07-27 section additions.

`section_tools.clone_section` calls `scan_id_space(root)` to find the next free id,
but a script that builds several sections inserts them only after the last one is
built. Every clone in such a run therefore received the same starting id and the
clones collided with each other: six `wp:docPr/@id` values and six VML
`v:shape/@id` values ended up duplicated across the seven new sections. Word
reports duplicated drawing ids as file corruption.

The cause is fixed in `section_tools` (ids handed out are remembered across calls
within a run). This script repairs the document that was already written, by
giving every drawing inside the seven new sections a fresh id above the current
document maximum in each id space.

Scope is deliberately limited to the sections added in that session. A
document-wide scan also finds 9 duplicate docPr ids, 9 duplicate shape ids, 2
duplicate spids, 24 duplicate bookmarkStart ids and 3 orphaned bookmarkEnd
elements in content that predates this work. Those are reported to the project
owner rather than silently rewritten.

Usage:
    python scripts/oneoff/repair_new_section_ids.py [--dry-run]
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

O = "urn:schemas-microsoft-com:office:office"

NEW_SECTIONS = [
    "Add VNeID Field", "Update VNeID Field",
    "Add Redact Field", "Update Redact Field",
    "Add Elabor Field", "Update Elabor Field",
    "Get Field Types",
]


def duplicates(root):
    import collections
    dp = collections.Counter(d.get("id") for d in root.iter() if st.local(d) == "docPr")
    sh = collections.Counter(d.get("id") for d in root.iter() if st.local(d) == "shape")
    return (sum(v - 1 for v in dp.values() if v > 1),
            sum(v - 1 for v in sh.values() if v > 1))


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))
    kids = list(body)

    before = duplicates(root)
    print("duplicate ids before: docPr=%d shape=%d" % before)

    # collect the body children that make up the new sections
    targets = []
    for title in NEW_SECTIONS:
        s, e = st.section_range(kids, st.find_heading(kids, 3, title))
        targets.extend(kids[s:e])

    # allocate above the current maximum in each space, counting the whole document
    next_docpr, next_spid = st.scan_id_space(root)
    n_doc = n_vml = 0
    for el in targets:
        for d in el.iter():
            if st.local(d) == "docPr":
                d.set("id", str(next_docpr))
                d.set("name", "Text Box %d" % next_docpr)
                next_docpr += 1
                n_doc += 1
        for d in el.iter():
            if st.local(d) == "shape":
                d.set("id", "_x0000_s%d" % next_spid)
                if d.get("{%s}spid" % O) is not None:
                    d.set("{%s}spid" % O, "_x0000_s%d" % next_spid)
                next_spid += 1
                n_vml += 1

    after = duplicates(root)
    print("re-ided in the 7 new sections: docPr=%d shape=%d" % (n_doc, n_vml))
    print("duplicate ids after : docPr=%d shape=%d   (remainder predates this work)"
          % after)

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
