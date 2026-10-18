#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Repair the malformed Sample Request frame of `Delete Document Field`.

Two defects, both in that one frame, both visible in Word:

1. The URL line and the `Content-Type` header line are glued together inside a
   single run, so the frame renders as
   `DELETE {baseURL}/.../fieldsContent-Type: application/json`.
   This also made the contract extractor read the path as `.../fieldsContent-Type:`.
2. `Authorization: Bearer {accesstoken}"` is missing its opening quote. The other
   360 frames in the document all use `Authorization: "Bearer {accesstoken}"`.

The Content-Type line is created by cloning the frame's own `x-language-name`
line, so it inherits the bold-key / plain-value run split required by template
2.8.2 instead of being rebuilt from scratch.

Usage:
    python scripts/oneoff/fix_delete_field_frame.py [--dry-run]
"""
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

GLUED = "Content-Type: application/json"
BAD_AUTH = ": Bearer "
GOOD_AUTH = ': "Bearer '


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def runs_with_text(p):
    return [r for r in p.findall(qn("r")) if r.find(qn("t")) is not None]


def set_run_text(run, text):
    for extra in run.findall(qn("t"))[1:]:
        run.remove(extra)
    t = run.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")


def fix_frame(txbx):
    """Returns (split_done, quote_done) for one frame copy."""
    paras = txbx.findall(qn("p"))
    if not paras:
        return False, False
    first = paras[0]
    if not re.match(r"\s*(GET|POST|PUT|DELETE|PATCH)\s", text_of(first)):
        return False, False

    split_done = quote_done = False

    # 1. unglue the Content-Type header from the path run
    target = None
    for run in runs_with_text(first):
        if GLUED in text_of(run):
            target = run
            break
    if target is not None:
        whole = text_of(target)
        path_part = whole[:whole.index(GLUED)]
        set_run_text(target, path_part)

        donor = None
        for p in paras[1:]:
            if text_of(p).strip().startswith("x-language-name"):
                donor = p
                break
        if donor is None:
            raise LookupError("no x-language-name line to clone the header style from")
        new_p = deepcopy(donor)
        dr = runs_with_text(new_p)
        set_run_text(dr[0], "Content-Type")
        set_run_text(dr[1], ": application/json")
        for extra in dr[2:]:
            new_p.remove(extra)
        first.addnext(new_p)
        split_done = True

    # 2. restore the missing opening quote on the Authorization value
    for p in txbx.findall(qn("p")):
        if not text_of(p).strip().startswith("Authorization"):
            continue
        for run in runs_with_text(p):
            if text_of(run) == BAD_AUTH:
                set_run_text(run, GOOD_AUTH)
                quote_done = True
    return split_done, quote_done


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

    splits = quotes = 0
    for txbx in root.iter(qn("txbxContent")):
        whole = text_of(txbx)
        if GLUED not in whole and BAD_AUTH not in whole:
            continue
        if "/fields" not in whole:
            continue
        s, q = fix_frame(txbx)
        splits += int(s)
        quotes += int(q)

    print("URL line unglued from Content-Type : %d frame copies" % splits)
    print("Authorization opening quote restored: %d frame copies" % quotes)
    if not (splits or quotes):
        sys.exit("nothing matched")

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
