#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Correct the path of `Update Qrypto Field`: /fields/qryto -> /fields/qrypto.

The document is missing the `p`. `qryto` is not registered in
`FieldRegistryBootstrap`, so `ParseToField.parseToFieldV3` finds no resolver and
returns NotFound - the documented call cannot work. `Add Qrypto Field` in the same
section already uses the correct `/fields/qrypto`, and so does Postman.

`qryto` is not a substring of `qrypto`, so a plain substring match is exact here;
a document-wide scan finds the two runs of the one frame (Choice + Fallback) and
nothing else.

Usage:
    python scripts/oneoff/fix_qrypto_path_typo.py [--dry-run]
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

OLD = "/fields/qryto"
NEW = "/fields/qrypto"


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
