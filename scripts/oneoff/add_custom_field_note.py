#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Record the custom-field behaviour of the text field in the TextBox sections.

Postman carries `Add Custom Field` and `Update Custom FIeld`, but they are not
separate endpoints: both POST/PUT to `/fields/text`. `DetectCustomFieldType`
reads `type` from the payload, matches it against the custom fields configured
for the caller's enterprise and, on a match, rewrites the URI to `/fields/text`,
injects `type_name` into the payload and selects the `custom` resolver instead of
the text resolver. So this is a payload behaviour of the existing text field, and
the project owner decided on 2026-07-27 to describe it in the two TextBox
sections rather than clone a section for it.

The description is appended to each section's existing prose paragraph. That
paragraph exists in every section already, so nothing new is introduced into the
document and no new rule is needed in the format template - which is what would
have been required had this been rendered as a callout or note block.

Facts from the backend:
  * the match is on the payload's `type` value against the enterprise's custom
    field list - `DetectCustomFieldType.detect`;
  * an account outside an enterprise is rejected with
    `SUBCODE_FIELD_CUSTOM_FIELD_REQUIRES_ENTERPRISE`;
  * the request keeps going to `/fields/text` - there is no `/fields/custom`
    route, and `custom` is only reachable through this detection path.

Usage:
    python scripts/oneoff/add_custom_field_note.py [--dry-run]
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

SECTIONS = ("Add TextBox Field", "Update TextBox Field")

NOTE = (" When the type parameter carries the name of a custom field configured for the "
        "enterprise, the request is processed as that custom field instead of a plain text "
        "field. The account must belong to an enterprise; otherwise the request is rejected.")

MARKER = "custom field configured for the enterprise"


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    kids = list(root.find(st.qn("body")))

    done = 0
    for title in SECTIONS:
        idx = st.find_heading(kids, 3, title)
        prose = kids[idx + 1]
        current = st.text_of(prose).strip()
        if MARKER in current:
            print("  %-22s already carries the note - skipped" % title)
            continue
        if not current.startswith("Business applications"):
            sys.exit("unexpected prose under %r: %r" % (title, current[:80]))
        st.set_para_text(prose, current + NOTE)
        print("  %-22s prose %d -> %d chars" % (title, len(current), len(current) + len(NOTE)))
        done += 1

    if not done:
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
