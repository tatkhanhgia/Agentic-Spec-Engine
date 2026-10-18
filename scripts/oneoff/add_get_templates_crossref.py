#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Leave a cross-reference where `Get Templates` used to live.

The section was moved from `Personal Settings` into the `Template` chapter by
owner decision (2026-07-27). Moving it silently would strip the entry from the
table of contents of the chapter a reader is most likely to browse first, so a
short stub stays behind pointing at the new location.

The stub is a Heading 3 plus one paragraph and carries no Sample Request frame,
so `extract_doc_contract.py` does not read it as a second documented endpoint -
the endpoint count stays correct.

Placement restores the original reading order: `Get Templates` was the last
section of `Personal Settings`, directly after `Remove Contact in Address Book`.
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

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_crossref.docx")

DONOR = "Get Address Book"
ANCHOR = "Remove Contact in Address Book"
TITLE = "Get Templates"
NOTE = ("This service API is documented under Template > Get Templates. It "
        "returns the personal and/or enterprise templates available to the "
        "logged-in user, and is described there together with the other "
        "template operations.")


def is_plain_paragraph(el):
    """A paragraph that carries text, not a floating frame."""
    return el.tag == st.qn("p") and not list(el.iter(st.qn("txbxContent")))


def main():
    tree, token = docx_guard.load(DOC)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    # Reuse a real Personal Settings section so the heading and body paragraph
    # keep their styles instead of being invented.
    clone, _ = st.clone_section(root, list(body), DONOR, level=3)
    heading = clone[0]
    desc = next(e for e in clone[1:]
                if is_plain_paragraph(e) and st.text_of(e).strip())
    blank = next(e for e in clone[1:]
                 if is_plain_paragraph(e) and not st.text_of(e).strip())

    st.set_para_text(heading, TITLE)
    st.set_para_text(desc, NOTE)

    st.insert_after_section(body, list(body), ANCHOR, [heading, desc, blank],
                            level=3)
    print("+ cross-reference stub %r after %r" % (TITLE, ANCHOR))

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
    st.pack(UNPACKED, OUT_DOCX)
    print("packed -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
