#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Rename the `Get All Package` section to `Get All Packages`.

The endpoint returns a list, and both the SigningHub manual (`Get Packages`) and
the Postman collection (`Get All Packages`) use the plural. Owner chose to add
the trailing "s" rather than adopt SigningHub's shorter title.

Renames the Heading 3 and its TOC row. Page numbers stay stale until the field
is refreshed in Word (Ctrl+A, then F9).

Usage:
    python scripts/oneoff/rename_get_all_packages.py [--dry-run]
"""
import os
import shutil
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

OLD = "Get All Package"
NEW = "Get All Packages"


def el_text(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def para_style(el):
    pPr = el.find(qn("pPr"))
    if pPr is None:
        return ""
    style = pPr.find(qn("pStyle"))
    return style.get(qn("val")) if style is not None else ""


def set_leaf_text(el, value):
    runs = [t for t in el.iter(qn("t"))]
    if not runs:
        return False
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""
    return True


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
    tree = etree.parse(DOC_XML, etree.XMLParser(huge_tree=True))
    root = tree.getroot()

    heading = 0
    toc = 0
    for para in root.iter(qn("p")):
        style = para_style(para)
        text = el_text(para).strip()
        if style == "Heading3" and text == OLD:
            set_leaf_text(para, NEW)
            heading += 1
        elif style == "TOC3" and OLD in text and NEW not in text:
            # the TOC row is "<number><title><page>"; rewrite only the title run
            for t in para.iter(qn("t")):
                if (t.text or "").strip() == OLD:
                    t.text = (t.text or "").replace(OLD, NEW)
                    toc += 1
                    break

    print("heading renamed: %d | TOC row renamed: %d" % (heading, toc))
    if heading != 1:
        sys.exit("expected exactly 1 heading named %r, found %d" % (OLD, heading))
    if toc != 1:
        sys.exit("expected exactly 1 TOC row for %r, found %d" % (OLD, toc))

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
