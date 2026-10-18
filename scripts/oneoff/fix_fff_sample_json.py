#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Rebuild the Fill Form Fields sample-request JSON so it is valid JSON
(template rule 2.8.4: double-quoted keys, 2-space indent, no trailing commas).

The sample body spans TWO textbox frames (each duplicated as mc:Choice +
mc:Fallback, so 4 w:txbxContent total under one anchor paragraph). The old
body was structurally broken: '},{}'' stubs, a stamp entry without braces,
a missing comma after the "date" array, stray NBSP indentation.

Line counts are kept close to the old layout (55 -> 56 body lines in frame 1,
25 -> 25 lines in frame 2) so the fixed-size frames do not overflow.
Field-type keys match the Request Attributes table (backend wire names).
New body lines are cloned from existing body-line paragraphs, never built
from scratch. Writes through docx_guard.
"""
import copy
import json
import os
import sys

SCRIPTS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SCRIPTS)
from docx_guard import load, save, DOC_XML  # noqa: E402
from lxml import etree  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(qn("t")))


def style_of(p):
    ps = p.find(qn("pPr") + "/" + qn("pStyle"))
    return ps.get(qn("val")) if ps is not None else ""


def entry(key, fields):
    lines = ['  "%s": [' % key, "    {"]
    items = list(fields.items())
    for i, (k, v) in enumerate(items):
        comma = "," if i < len(items) - 1 else ""
        lines.append('      "%s": %s%s' % (k, json.dumps(v), comma))
    lines += ["    }", "  ],"]
    return lines


BOX1 = ["{"]
BOX1 += entry("textbox", {"field_name": "textbox1",
                          "value": "Name one two three"})
BOX1 += entry("radiobox", {"field_name": "radio1", "value": True})
BOX1 += entry("checkbox", {"field_name": "checkbox1", "value": True})
BOX1 += entry("stamp", {"field_name": "Seal1", "value": "String"})
BOX1 += entry("camera", {"field_name": "Camera1", "value": "String"})
BOX1 += entry("attachment", {"field_name": "attachment1",
                             "file_name": "attachment.png",
                             "value": "String"})
BOX1 += entry("hyperlink", {"field_name": "hyperlink1", "value": "String"})
BOX1 += entry("date", {"field_name": "date1", "value": "String"})
BOX1 += entry("combobox", {"field_name": "combobox1", "value": "String"})

BOX2 = []
BOX2 += entry("toggle", {"field_name": "toggle1", "value": "String"})
BOX2 += entry("numeric_stepper", {"field_name": "stepper1",
                                  "value": "String"})
BOX2 += entry("checkboxv2", {"field_name": "checkboxv2_1", "value": True})
BOX2 += entry("radioboxv2", {"field_name": "radioboxv2_1", "value": True})
BOX2[-1] = "  ]"  # last entry: no trailing comma
BOX2.append("}")

# the concatenation must be valid JSON before we touch the document
json.loads("\n".join(BOX1 + BOX2))
print("frame1 body lines: %d, frame2 lines: %d" % (len(BOX1), len(BOX2)))


def set_line_text(line, text):
    first = None
    for t in line.iter(qn("t")):
        if first is None:
            first = t
            t.text = text
            t.set(XML_SPACE, "preserve")
        else:
            t.text = ""
    if first is None:
        raise RuntimeError("template line has no w:t")


def rebuild(tx, keep, lines, template):
    """Replace all but the first `keep` paragraphs of a txbxContent."""
    paras = tx.findall(qn("p"))
    for p in paras[keep:]:
        tx.remove(p)
    for text in lines:
        dup = copy.deepcopy(template)
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for b in dup.findall(".//" + qn(tag)):
                b.getparent().remove(b)
        set_line_text(dup, text)
        tx.append(dup)


def main():
    tree, token = load(DOC_XML)
    body = tree.getroot().find(qn("body"))

    in_dp = False
    cur = None
    frame_p = None
    for el in body:
        if etree.QName(el).localname != "p":
            continue
        st = style_of(el)
        txt = para_text(el).strip()
        if st == "Heading2":
            in_dp = (txt == "Document Processing")
            continue
        if in_dp and st in ("Heading3", "Heading4"):
            cur = txt
            continue
        if in_dp and cur == "Fill Form Fields" and "PUT {baseURL}" in txt \
                and el.findall(".//" + qn("txbxContent")):
            frame_p = el
            break
    if frame_p is None:
        raise RuntimeError("Fill Form Fields request frame not found")

    txs = frame_p.findall(".//" + qn("txbxContent"))
    box1 = [tx for tx in txs
            if para_text(tx.findall(qn("p"))[0]).startswith("PUT ")]
    box2 = [tx for tx in txs if tx not in box1]
    if len(box1) != 2 or len(box2) != 2:
        raise RuntimeError("expected 2+2 txbxContent, got %d+%d"
                           % (len(box1), len(box2)))

    for tx in box1:
        paras = tx.findall(qn("p"))
        if para_text(paras[3]).strip() != "x-language-name: en":
            raise RuntimeError("unexpected header layout in frame 1")
        template = paras[4]  # the old '{' body line
        rebuild(tx, keep=4, lines=BOX1, template=copy.deepcopy(template))
    for tx in box2:
        paras = tx.findall(qn("p"))
        template = next((p for p in paras if p.find(".//" + qn("t")) is not None),
                        None)
        if template is None:
            raise RuntimeError("frame 2 has no text line to use as template")
        rebuild(tx, keep=0, lines=BOX2, template=copy.deepcopy(template))

    # verify in-tree before saving
    for tx in box1 + box2:
        got = [para_text(p) for p in tx.findall(qn("p"))]
        want = (["PUT"] if tx in box1 else None)
    joined = "\n".join(
        [para_text(p) for p in box1[0].findall(qn("p"))[4:]] +
        [para_text(p) for p in box2[0].findall(qn("p"))])
    json.loads(joined)
    print("in-tree JSON re-validated OK")

    save(tree, DOC_XML, token)


if __name__ == "__main__":
    main()
