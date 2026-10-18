#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Undo the `Get Template Details`.`participants` link - it points at the wrong shape.

`link_template_composite_types.py` linked it to `TemplateWorkflowParticipant`,
whose attributes were taken from `TemplateWorkflowParticipantSerializer`. That
serializer is used by the six workflow-participant endpoints. **`Get Template
Details` does not use it.** `GetTemplateDetailsSerializer` writes

    jg.writeObjectField("participants", details.getParticipants() ...)

which is `List<GetWorkflowParticipantDTO>` handed straight to Jackson - a raw
persistence object: `id`, `packageId`, `workflowId`, `email`, `firstName`,
`lastName`, `mobileNumber`, `type`, `processAt`, `signingOrder`,
`metaInformation`, `status`, `uuid`, `guestUser`, `emailNotification`,
`securityDataEncrypted`. It shares almost nothing with the documented
participant shape and its keys are camelCase.

So the attribute goes back to a plain `Array`, and `TemplateWorkflowParticipant`
is removed because nothing links to it any more: the five root-array endpoints
have no attribute name to hang a composite type on, so they document their
fields inline. An orphan object invites exactly the mistake being undone here.
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
from link_template_composite_types import h2_in_chapter, find_type_cell  # noqa: E402

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_participants_revert.docx")

OBJECT = "TemplateWorkflowParticipant"
DESC = ("The participants configured in the template workflow, as stored. This "
        "attribute returns the internal participant record, which differs from "
        "the shape returned by Get Template Workflow Users - use that API when "
        "you need the participant list of a template.")


def main():
    tree, token = docx_guard.load(DOC)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    # set_cell_text drops every direct child of the cell paragraph, including the
    # w:hyperlink - which is exactly what turning a composite back into a plain
    # type requires.
    tc = find_type_cell(list(body), "Get Template Details", "participants")
    st.set_cell_text(tc, "Array")
    row = tc.getparent()
    st.set_cell_text(row.findall(st.qn("tc"))[4], DESC)
    print("~ Get Template Details.participants: link removed, type -> Array")

    kids = list(body)
    i = h2_in_chapter(kids, OBJECT, "RELATIVE OBJECT")
    s, e = st.section_range(kids, i)
    for el in kids[s:e]:
        body.remove(el)
    print("- removed orphan object %r (%d elements)" % (OBJECT, e - s))

    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    assert "_TemplateWorkflowParticipant" not in names, "bookmark still present"

    # nothing anywhere may still point at the deleted object
    dangling = [h.get(st.qn("anchor")) for h in root.iter(st.qn("hyperlink"))
                if h.get(st.qn("anchor")) == "_TemplateWorkflowParticipant"]
    if dangling:
        sys.exit("still %d hyperlink(s) to the removed object" % len(dangling))
    print("verified: no dangling reference to %r" % OBJECT)

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
    st.pack(UNPACKED, OUT_DOCX)
    print("packed -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
