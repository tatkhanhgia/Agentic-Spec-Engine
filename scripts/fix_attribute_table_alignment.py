#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Centre the `No` and `Presence` columns of every attribute table.

Template 2.4 requires both columns centered. Part of the document predates the
rule, so the two styles were mixed - which is why a section cloned from an old
donor inherited left alignment. This normalises the whole document so
`scripts/check_frame_format.py` check D passes and future clones cannot inherit
the defect.

Formatting only: no cell text is touched, no row is added or removed.

Usage:
    python scripts/fix_attribute_table_alignment.py [--dry-run]
"""
import os
import shutil
import sys
import zipfile
import collections

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

ATTRIBUTE_TABLES = ("Header Request Attributes", "Path Attributes",
                    "Request Attributes", "Response Attributes",
                    "Header Response Attributes")

# Template 2.4 identifies an attribute table by its column header row, not by its
# title: endpoint tables use the titles above, but a table in the object chapters
# is titled with the object's own name. `check_frame_format.py` uses the same
# signature, so the fixer and the checker cannot disagree about what they cover.
COLUMN_HEADER = ("No", "Name", "Type", "Presence", "Description")
CENTER_COLUMNS = ((0, "No"), (3, "Presence"))


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def cell_text(tc):
    return " ".join(x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def centre(tc):
    """Force every paragraph in the cell to centered. Returns True if changed."""
    changed = False
    for p in tc.findall(qn("p")):
        pPr = p.find(qn("pPr"))
        if pPr is None:
            pPr = etree.Element(qn("pPr"))
            p.insert(0, pPr)
        jc = pPr.find(qn("jc"))
        if jc is None:
            # w:jc follows w:pStyle/w:numPr but Word tolerates append here;
            # keep it simple and let Word normalise on save.
            jc = etree.SubElement(pPr, qn("jc"))
        if jc.get(qn("val")) != "center":
            jc.set(qn("val"), "center")
            changed = True
    return changed


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
    # Another agent edits this repository too; the token lets save() refuse to
    # overwrite a file that changed while this script was running.
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()

    per_column = collections.Counter()
    per_table = collections.Counter()
    tables = 0

    for tbl in root.iter(qn("tbl")):
        trs = tbl.findall(qn("tr"))
        if len(trs) < 3:
            continue
        title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()
        header = tuple(cell_text(tc).strip()
                       for tc in trs[1].findall(qn("tc")))[:len(COLUMN_HEADER)]
        if header != COLUMN_HEADER:
            continue
        tables += 1
        touched = False
        for tr in trs[2:]:
            tcs = tr.findall(qn("tc"))
            if len(tcs) < 5:
                continue
            for col, label in CENTER_COLUMNS:
                if centre(tcs[col]):
                    per_column[label] += 1
                    touched = True
        if touched:
            per_table[title] += 1

    print("attribute tables scanned : %d" % tables)
    for label in ("No", "Presence"):
        print("cells centered (%-8s): %d" % (label, per_column[label]))
    print("tables changed by type   :")
    for k, v in per_table.most_common():
        print("   %-28s %d" % (k, v))

    if not sum(per_column.values()):
        print("\nnothing to change")
        return
    if dry:
        print("\n--dry-run: nothing written")
        return
    try:
        guarded_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("\nREFUSED TO WRITE\n%s" % ex)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
