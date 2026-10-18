#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Link the `inperson` row of `Get Document Fields` to its object.

`Get Document Fields` lists one row per field type. Five of them link correctly;
the rest carry a composite label with no `w:hyperlink` at all. Only one of those
can be resolved from what the document already contains:

  `inperson` is typed `List<signatureFieldAttribute>` - the exact same label as
  the `signature` row two lines above it, which links to `_SignatureFieldAttribute_1`
  (object `SignatureFieldAttribute`). CLAUDE.md notes that one label legitimately
  serves several fields, so the shared label is not the evidence on its own; the
  evidence is that an in-person signature field and a signature field carry the
  same attribute shape, and the document defines exactly one object for it.

The remaining ten rows name objects that do not exist anywhere in the document
(`radioboxFieldAttribute`, `comboBoxFieldAttribute`, `dateTimeFieldAttribute`,
`imageFieldAttribute`, `cameraFieldAttribute`, `attachmentFieldAttribute`,
`hyperLinkFieldAttribute`, `toggleFieldAttribute`, `numericFieldAttribute`,
`checkboxFieldAttributeV2`, `radioboxFieldAttributeV2`). Writing them means
authoring eleven new object sections from eleven backend attribute classes, which
is a scope decision for the project owner, so they are reported rather than
guessed at.

Usage:
    python scripts/oneoff/fix_inperson_type_link.py [--dry-run]
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

SECTION = "Get Document Fields"
TABLE = "Response Attributes"
DONOR_FIELD = "signature"        # already linked, same label
TARGET_FIELD = "inperson"


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    kids = list(root.find(st.qn("body")))

    s, e = st.section_range(kids, st.find_heading(kids, 3, SECTION))
    tbl = st.find_table(kids[s:e], TABLE)

    donor = target = None
    for tr in st.data_rows(tbl):
        tcs = tr.findall(st.qn("tc"))
        name = st.cell_text(tcs[1])
        if name == DONOR_FIELD:
            donor = tcs[2]
        elif name == TARGET_FIELD:
            target = tcs[2]
    if donor is None or target is None:
        sys.exit("could not find both %r and %r rows" % (DONOR_FIELD, TARGET_FIELD))

    anchors = [h.get(st.qn("anchor")) for h in donor.iter(st.qn("hyperlink"))]
    if not anchors:
        sys.exit("donor row %r is not linked either - nothing to copy" % DONOR_FIELD)
    anchor = anchors[0]

    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    if anchor not in names:
        sys.exit("donor anchor %r does not resolve - refusing" % anchor)

    label = st.cell_text(target)
    if st.cell_text(donor) != label:
        sys.exit("labels differ (%r vs %r) - the shared-shape assumption fails"
                 % (st.cell_text(donor), label))

    st.set_type_link(target, label, anchor, donor)
    print("  %s / %s -> %s (label %r kept as-is)" % (SECTION, TARGET_FIELD, anchor, label))

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
