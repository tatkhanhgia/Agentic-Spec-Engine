#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Move the four personal-contact endpoint sections out of
`Application/Device Management` into `Personal Settings`.

Owner decision 2026-07-27: "di chuyen section di nhe, toi muon tuan thu
SigningHub". SigningHub 10.0 keeps Contacts inside its Personal Settings
chapter, ordered after Groups (our Address Book) and after Get Templates, with
the list call first (Get Contacts, Add Contact, Update Personal Contact,
Delete Personal Contact). The Postman collection groups the same four requests
under a "Personal Settings" folder as well.

This is a MOVE, not a clone: the very same elements are re-parented, so
bookmarks, hyperlink anchors, drawing ids and VML ids all stay unique and
valid. `Register Application User` remains in Application/Device Management,
so that section does not become empty.

Usage:
    python scripts/oneoff/move_contacts_into_personal_settings.py [--dry-run]
"""
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import docx_guard

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

SOURCE_H2 = "Application/Device Management"
TARGET_H2 = "Personal Settings"

# SigningHub order within the chapter: list call first.
MOVE_ORDER = [
    "Get Contacts",
    "Add Personal Contact",
    "Update Personal Contact",
    "Delete Personal Contact",
]


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def find_heading(kids, level, title):
    want = "Heading%d" % level
    for i, e in enumerate(kids):
        if e.tag == qn("p") and style_of(e) == want and text_of(e).strip() == title:
            return i
    raise LookupError("heading not found: %s" % title)


def section_range(kids, start):
    for j in range(start + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and re.match(r"Heading\d$", style_of(e)):
            return start, j
    return start, len(kids)


def pack(src_dir, out_path):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(src_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for base, _dirs, files in os.walk(src_dir):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                if rel != "[Content_Types].xml":
                    z.write(full, rel)
    shutil.move(tmp, out_path)


def main():
    dry = "--dry-run" in sys.argv
    tree, token = docx_guard.load(DOC_XML)
    body = tree.getroot().find(qn("body"))
    kids = list(body)

    src_h2 = find_heading(kids, 2, SOURCE_H2)
    src_end = len(kids)
    for j in range(src_h2 + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and style_of(e) in ("Heading1", "Heading2"):
            src_end = j
            break
    # collect the four blocks, verifying each lives inside the source section
    blocks = []
    for title in MOVE_ORDER:
        h3 = find_heading(kids, 3, title)
        if not (src_h2 < h3 < src_end):
            raise RuntimeError("'%s' (idx %d) is not inside '%s' [%d..%d) - "
                               "document changed, re-inspect before moving"
                               % (title, h3, SOURCE_H2, src_h2, src_end))
        s, e = section_range(kids, h3)
        blocks.append((title, kids[s:e]))
        print("block  : [%d..%d)  %s  (%d elements)" % (s, e, title, e - s))

    tgt_h2 = find_heading(kids, 2, TARGET_H2)
    # the target section must run right up to the source H2: no other Heading2
    # may sit between them, because we insert at the seam between the sections
    between = [i for i in range(tgt_h2 + 1, src_h2)
               if kids[i].tag == qn("p") and style_of(kids[i]) == "Heading2"]
    if between:
        raise RuntimeError("another Heading2 sits between '%s' and '%s' - "
                           "document changed, re-inspect before moving"
                           % (TARGET_H2, SOURCE_H2))
    anchor = kids[src_h2 - 1]
    print("anchor : %d  (last element of '%s' section)" % (src_h2 - 1, TARGET_H2))

    if dry:
        print("\n--dry-run: nothing written")
        return

    for _title, els in blocks:
        for el in els:
            anchor.addnext(el)     # lxml re-parents: this IS the move
            anchor = el

    docx_guard.save(tree, DOC_XML, token)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
