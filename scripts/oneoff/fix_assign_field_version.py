#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Correct the API version of `Assign Document Field`: /v4/ -> /v1/.

The document declares `PUT {baseURL}/v4/packages/{packageId}/documents/{documentId}/fields/assign`.
The backend registry has `PUT /v1/packages/[0-9]*/documents/[0-9]*/fields/assign`, and
`contract/from_code.json` contains zero routes under /v4/ - the version simply does
not exist. The Postman collection also uses /v1/. So the document is the wrong one.

Only the version token is touched. The path lives in its own run
(`/v4/packages/`), so this never disturbs the badge run or the rest of the line.

Usage:
    python scripts/oneoff/fix_assign_field_version.py [--dry-run]
"""
import os
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

OLD = "/v4/packages/"
NEW = "/v1/packages/"


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
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()

    hits = 0
    for t in root.iter(qn("t")):
        if t.text and OLD in t.text:
            t.text = t.text.replace(OLD, NEW)
            hits += 1

    print("runs rewritten %s -> %s : %d" % (OLD, NEW, hits))
    if not hits:
        sys.exit("nothing matched - was this already fixed?")
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
