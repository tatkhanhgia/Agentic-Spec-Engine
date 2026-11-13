#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Cut a formatted specification DOCX down to a skeleton.

The renderer never writes WordprocessingML by hand: it clones a block that
already looks right and replaces its text. The skeleton is the smallest
document that still holds one of every block, so a render needs no earlier
specification:

  * the package: styles, numbering, headers, footers, logo, settings;
  * front matter: cover, History (one row), an emptied table of contents
    (Word rebuilds it on open), INTRODUCTION;
  * `API Specification`: its lead-in, one category heading, and a few
    endpoint sections chosen as templates (one per method and body shape,
    together holding every table kind and the shared error sample);
  * `REQUEST/RESPONSE OBJECT`, `RELATIVE OBJECT`, `ENUMERATION`: the chapter
    heading and one section each;
  * every other chapter (`HTTP ERROR CODE`, `Example`) unchanged.

The content of the template blocks is irrelevant: the renderer overwrites or
removes it (`render_from_openapi.py --skeleton`).

THIS SCRIPT ONLY READS the source document.

Usage:
    python scripts/build_skeleton.py <formatted.docx> [-o templates/spec-skeleton.docx]
"""
import argparse
import collections
import os
import sys
import zipfile

from lxml import etree

from extract_doc_contract import para_text, qn, style_of
from render_from_openapi import (
    TABLE_ORDER, drop_half_bookmarks, find_donors, plain_row, section_shape, unlink_dangling)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DEFAULT = os.path.join(ROOT, "templates", "spec-skeleton.docx")
API_CHAPTER = "API SPECIFICATION"
ONE_SECTION_CHAPTERS = ("REQUEST/RESPONSE OBJECT", "RELATIVE OBJECT", "ENUMERATION")
METHODS = ("GET", "POST", "PUT", "DELETE")


def chapters(body):
    """[(title, [elements])] split at Heading 1; the part before the first is titled ''."""
    out, title, current = [], "", []
    for el in body:
        if etree.QName(el).localname == "sectPr":
            continue
        if el.tag == qn("p") and style_of(el) == "Heading1":
            out.append((title, current))
            title, current = para_text(el).strip(), []
        current.append(el)
    out.append((title, current))
    return out


def sections_of(els, levels):
    """[(heading level, [elements])] of a chapter; elements before the first heading come first as level 0."""
    out, current, level = [], [], 0
    for el in els:
        lv = 0
        if el.tag == qn("p"):
            st = style_of(el)
            if st.startswith("Heading") and st[7:].isdigit() and int(st[7:]) in levels:
                lv = int(st[7:])
        if lv:
            out.append((level, current))
            current, level = [], lv
        current.append(el)
    out.append((level, current))
    return out


def choose_templates(sections):
    """Endpoint sections to keep: every method and body shape the document has,
    then whatever is needed to hold all table kinds and the shared error sample."""
    shaped = [(els, section_shape(els)) for lv, els in sections if lv == 3]
    shaped = [(els, sh) for els, sh in shaped if sh["method"] in METHODS]
    chosen = []
    for method in METHODS:
        for body in (True, False):
            pool = [(els, sh) for els, sh in shaped if sh["method"] == method and sh["body"] == body]
            if pool:
                chosen.append(max(pool, key=lambda x: (len(x[1]["kinds"]), x[1]["error"])))
    have = set().union(*[sh["kinds"] for _, sh in chosen]) if chosen else set()
    for kind in TABLE_ORDER:
        if kind not in have:
            extra = next(((els, sh) for els, sh in shaped if kind in sh["kinds"]), None)
            if extra is not None:
                chosen.append(extra)
                have |= extra[1]["kinds"]
    if not any(sh["error"] for _, sh in chosen):
        extra = next(((els, sh) for els, sh in shaped if sh["error"]), None)
        if extra is not None:
            chosen.append(extra)
    return [els for els, _ in chosen]


def has_plain_row(els):
    tbl = next((e for e in els if e.tag == qn("tbl")), None)
    return tbl is not None and any(plain_row(r) is not None for r in tbl.findall(qn("tr"))[2:])


def trim_toc(body):
    """Empty the table of contents down to its field.

    The entries name the source document's headings. The paragraph that opens
    the TOC field and the one that closes it stay, so Word rebuilds the table
    when it updates fields on open; every entry between them goes.
    """
    for sdt in body.findall(qn("sdt")):
        content = sdt.find(qn("sdtContent"))
        if content is None:
            continue
        paras = list(content)
        opens = next((i for i, p in enumerate(paras)
                      if " TOC " in "".join(t.text or "" for t in p.iter(qn("instrText")))), None)
        if opens is None:
            continue
        closes = max(i for i, p in enumerate(paras) if p.find(".//" + qn("fldChar")) is not None)
        for p in paras[opens + 1:closes]:
            content.remove(p)


def build(src, out):
    with zipfile.ZipFile(src) as z:
        root = etree.fromstring(z.read("word/document.xml"), etree.XMLParser(huge_tree=True))
    body = root.find(qn("body"))
    keep = []
    report = []
    for title, els in chapters(body):
        upper = title.upper()
        if upper == API_CHAPTER:
            parts = sections_of(els, (2, 3, 4))
            lead = parts[0][1]
            category = next((p for lv, p in parts if lv == 2), None)
            templates = choose_templates(parts)
            keep += lead
            if category is not None:
                keep.append(category[0])            # the Heading 2 paragraph alone
            for t in templates:
                keep += t
            report.append("%s: %d endpoint templates" % (title, len(templates)))
        elif upper in ONE_SECTION_CHAPTERS:
            parts = sections_of(els, (2,))
            keep += parts[0][1]
            sample = next((p for lv, p in parts[1:] if upper == "ENUMERATION" or has_plain_row(p)), None)
            if sample is not None:
                keep += sample
            report.append("%s: %d section kept of %d" % (title, int(sample is not None), len(parts) - 1))
        else:
            keep += els
            if title:
                report.append("%s: kept" % title)
    kept = {id(e) for e in keep}
    for el in list(body):
        if etree.QName(el).localname != "sectPr" and id(el) not in kept:
            body.remove(el)

    # History: the header row and one data row are enough to clone from.
    for i, el in enumerate(list(body)):
        if el.tag == qn("p") and style_of(el) == "Title" and para_text(el).strip() == "History":
            tbl = next((e for e in list(body)[i + 1:i + 4] if e.tag == qn("tbl")), None)
            if tbl is not None:
                for tr in tbl.findall(qn("tr"))[2:]:
                    tbl.remove(tr)

    trim_toc(body)

    stats = collections.Counter()
    drop_half_bookmarks(root, stats)
    unlink_dangling(root, stats)

    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, xml if item.filename == "word/document.xml" else zin.read(item.filename))
    return report, find_donors(body)


def main():
    parser = argparse.ArgumentParser(description="Cut a formatted specification DOCX down to a skeleton.")
    parser.add_argument("document")
    parser.add_argument("-o", "--out", default=OUT_DEFAULT)
    args = parser.parse_args()
    if os.path.exists(args.out) and os.path.samefile(args.document, args.out):
        sys.exit("Output must not be the source document")
    report, donors = build(args.document, args.out)
    for line in report:
        print(line)
    missing = [k for k in TABLE_ORDER if k not in donors["tables"]]
    print("table kinds available : %s" % ", ".join(k for k in TABLE_ORDER if k in donors["tables"]))
    if missing:
        print("table kinds MISSING   : %s" % ", ".join(missing))
    print("shared error sample   : %s" % ("yes" if donors["error"] is not None else "MISSING"))
    print("wrote %s (%d KB)" % (args.out, os.path.getsize(args.out) // 1024))


if __name__ == "__main__":
    main()
