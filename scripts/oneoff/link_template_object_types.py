#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Link the composite Type cell inside the new `Template(Get Templates)` object.

`link_template_composite_types.py` linked the endpoint tables. An object table
follows the same rule - `Workflow`.`details` already links to `WorkflowDetail` -
so the `documents` row of `Template(Get Templates)` has to link too, otherwise
the chain stops one hop short of the object it names.

Kept separate from the creating script because that one inserts its objects
unconditionally: re-running it would add a second copy of each.
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
ROOT = os.path.dirname(SCRIPTS)
sys.path.insert(0, SCRIPTS)

import section_tools as st          # noqa: E402
import docx_guard                   # noqa: E402
from link_template_composite_types import (  # noqa: E402
    find_type_cell, h2_in_chapter, only_table)

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_object_type_link.docx")

OBJECT = "Template(Get Templates)"
CHAPTER = "RELATIVE OBJECT"
ATTRIBUTE = "documents"
LABEL = "List<TemplateDocumentResponse>"
ANCHOR = "_TemplateDocumentResponse"


def main():
    tree, token = docx_guard.load(DOC)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    kids = list(body)
    donor_tc = find_type_cell(kids, "Get Template By ID", ATTRIBUTE)

    i = h2_in_chapter(kids, OBJECT, CHAPTER)
    s, e = st.section_range(kids, i)
    tbl = only_table(kids[s:e])
    target = None
    for tr in st.data_rows(tbl):
        tcs = tr.findall(st.qn("tc"))
        if st.cell_text(tcs[1]) == ATTRIBUTE:
            target = tcs[2]
            break
    if target is None:
        sys.exit("no %r row in %r" % (ATTRIBUTE, OBJECT))

    st.set_type_link(target, LABEL, ANCHOR, donor_tc)

    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    links = [h.get(st.qn("anchor")) for h in target.iter(st.qn("hyperlink"))]
    if not links or links[0] not in names or not st.cell_text(target).strip():
        sys.exit("link verification failed: %r" % links)
    print("-> %s.%s  %s  %s" % (OBJECT, ATTRIBUTE, LABEL, links[0]))
    print("verified: anchor resolves, label non-empty")

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
    st.pack(UNPACKED, OUT_DOCX)
    print("packed -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
