#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Fix JSON body indentation in Get Workflow Users success response frames (check E).

The nested group_member object was pasted at column 0 instead of the correct
depth (template 2.8.4: 2 spaces per nesting level). Both txbxContent copies
are updated.

Usage:
    python scripts/oneoff/fix_get_workflow_users_indent.py [--dry-run]
"""
import os
import shutil
import sys
import zipfile

from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from docx_guard import load, save, ConcurrentModification
from section_tools import qn, text_of, style_of, find_heading, section_range, pack

UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

REPLACEMENTS = {
    "{": "        {",
    '  "user_email": "email"': '          "user_email": "email"',
    "}": "        }",
}


def set_para_line(p, new_text):
    t_nodes = list(p.iter(qn("t")))
    if not t_nodes:
        return False
    if text_of(p) == new_text:
        return False
    t_nodes[0].text = new_text
    for extra in t_nodes[1:]:
        extra.text = ""
    return True


def fix_txbx(txbx):
    paras = txbx.findall(qn("p"))
    lines = [text_of(p) for p in paras]
    changed = 0
    for i, line in enumerate(lines):
        if line.strip() != "{":
            continue
        if i == 0 or not lines[i - 1].strip().endswith('"group_member":['):
            continue
        if i + 2 >= len(lines):
            continue
        if lines[i + 1].strip() != '"user_email": "email"' or lines[i + 2].strip() != "}":
            continue
        if set_para_line(paras[i], REPLACEMENTS["{"]):
            changed += 1
        if set_para_line(paras[i + 1], REPLACEMENTS['  "user_email": "email"']):
            changed += 1
        if set_para_line(paras[i + 2], REPLACEMENTS["}"]):
            changed += 1
    return changed


def main():
    dry = "--dry-run" in sys.argv
    tree, token = load(DOC_XML)
    root = tree.getroot()
    body = root.find(qn("body"))
    kids = list(body)

    try:
        idx = find_heading(kids, 3, "Get Workflow Users")
    except LookupError:
        sys.exit("Get Workflow Users heading not found")

    start, end = section_range(kids, idx)
    total = 0
    frames = 0
    for el in kids[start:end]:
        for txbx in el.iter(qn("txbxContent")):
            paras = txbx.findall(qn("p"))
            if not paras:
                continue
            if not text_of(paras[0]).strip().startswith("Http Code: 200"):
                continue
            n = fix_txbx(txbx)
            if n:
                frames += 1
                total += n

    print("Get Workflow Users: %d frame(s), %d line(s) re-indented" % (frames, total))
    if total == 0:
        return
    if dry:
        print("--dry-run: nothing written")
        return
    try:
        save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("REFUSED TO WRITE\n%s" % ex)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
