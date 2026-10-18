#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Correct the HTTP method of `Pre Login Authentication` from GET to POST.

The document showed `GET {baseURL}/v1/authenticate/pre` while:
  * the backend `AuthorizeController` handles that route in `doPost` only,
  * the SigningHub manual documents the call with a required `{username}` body,
  * the document's own Sample Request frame already carried a JSON body
    (`{"email": "String"}`), which a GET does not take.

Changing the verb also changes the badge colour: per template 2.8.1 GET is
green-700 and POST is orange-700. The colours are read from
`scripts/GoPaperless_DOCX_Format_Template.md` rather than hardcoded, per the
project rule that formatting values live in the template.

Usage:
    python scripts/oneoff/fix_prelogin_method.py [--dry-run]
"""
import os
import shutil
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fix_docx_format import parse_frame_method_rules

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
TEMPLATE = os.path.join(ROOT, "scripts", "GoPaperless_DOCX_Format_Template.md")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

TARGET_PATH = "/v1/authenticate/pre"
OLD_METHOD = "GET"
NEW_METHOD = "POST"


def frame_text(txbx):
    return "".join(t.text or "" for t in txbx.iter(qn("t")))


def retag_method(txbx, rules):
    """Rewrite the badge run of one request frame. Returns True if changed."""
    paras = txbx.findall(qn("p"))
    if not paras:
        return False
    runs = paras[0].findall(qn("r"))
    if not runs:
        return False

    first = runs[0]
    t = first.find(qn("t"))
    if t is None or (t.text or "").strip() != OLD_METHOD:
        return False

    style = rules[NEW_METHOD]
    t.text = NEW_METHOD

    rPr = first.find(qn("rPr"))
    if rPr is None:
        rPr = etree.SubElement(first, qn("rPr"))

    shd = rPr.find(qn("shd"))
    if shd is None:
        shd = etree.SubElement(rPr, qn("shd"))
        shd.set(qn("val"), "clear")
        shd.set(qn("color"), "auto")
    shd.set(qn("fill"), style["highlight"])

    color = rPr.find(qn("color"))
    if color is None:
        color = etree.SubElement(rPr, qn("color"))
    color.set(qn("val"), style["text_color"])
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
    rules = parse_frame_method_rules(TEMPLATE)
    for m in (OLD_METHOD, NEW_METHOD):
        if m not in rules:
            sys.exit("template 2.8.1 has no badge rule for %s" % m)
    print("template badge %s -> fill %s / text %s"
          % (NEW_METHOD, rules[NEW_METHOD]["highlight"], rules[NEW_METHOD]["text_color"]))

    tree = etree.parse(DOC_XML, etree.XMLParser(huge_tree=True))
    root = tree.getroot()

    changed = 0
    for txbx in root.iter(qn("txbxContent")):
        text = frame_text(txbx)
        if TARGET_PATH not in text or not text.lstrip().startswith(OLD_METHOD):
            continue
        if retag_method(txbx, rules):
            changed += 1

    print("request frames retagged %s -> %s : %d" % (OLD_METHOD, NEW_METHOD, changed))
    if changed == 0:
        sys.exit("nothing matched - the frame may already be POST")

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
