#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extract a document's sections that no spec operation serves into the
planned-sections file (`doc-planned.yaml`, see `doc_planned.py`).

A section is extracted when `compare_openapi_doc.pair` finds no operation
for it: its method and path, its title and its route all miss the spec.
Everything the section shows is copied as written: prose, sample request
and response, and every table row except the shared error rows of the
response table, which the renderer adds to every section itself.

THIS SCRIPT ONLY READS the document and the contract.

Usage:
    python scripts/extract_planned.py <document.docx> <from_openapi.json> [-o config/doc-planned.yaml]
"""
import argparse
import json
import os
import re
import sys
import zipfile

import yaml
from lxml import etree

from compare_openapi_doc import ERROR_ROWS, doc_path, pair
from doc_planned import TABLES, validate
from extract_doc_contract import extract, para_text, qn, style_of
from render_from_openapi import frame_body, frame_kind, section_ranges

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DEFAULT = os.path.join(ROOT, "config", "doc-planned.yaml")

HEADER = """\
# Sections of the specification document for APIs that are designed but have
# no route in the code yet. openapi.yaml describes only what the code serves,
# so they live here; the renderer (scripts/render_from_openapi.py --planned)
# shows each one marked "(Not implemented yet)".
#
# Generated once from the approved document; edited by hand afterwards.
# When an API is implemented, describe it in openapi.yaml and delete its entry
# here: the renderer refuses an entry whose method and path the spec serves.
#
# `after` names the section the entry follows in its category; without it, or
# when no section has that title, the entry goes to the end of the category.
# `first: true` puts the entry at the start of its category instead.
# The response table's shared error rows are added by the renderer.
"""


def sample(txbx):
    """A frame's JSON body: the parsed value, or its text when it is not valid JSON."""
    found = frame_body(txbx)
    if found is None:
        return None
    paras, start = found
    text = "\n".join(para_text(p) for p in paras[start:]).strip()
    try:
        return json.loads(text)
    except ValueError:
        return text


def section_parts(els):
    """(prose lines, request sample, response sample) of a section's elements."""
    prose, collecting = [], True
    request = response = None
    for el in els[1:]:
        if el.tag == qn("tbl"):
            collecting = False
            continue
        if el.tag != qn("p"):
            continue
        frames = list(el.iter(qn("txbxContent")))
        text = para_text(el).strip()
        if frames or text == "Sample Request":
            collecting = False
        elif collecting:
            prose += [" ".join(x.split()) for x in text.split("\n") if x.strip()]
        seen = set()
        for txbx in frames:
            # A frame is stored twice (DrawingML and its VML fallback).
            kind = frame_kind(txbx)
            if kind in seen:
                continue
            seen.add(kind)
            if kind == "request" and request is None:
                request = sample(txbx)
            elif kind == "response" and response is None:
                response = sample(txbx)
    return prose, request, response


def main():
    parser = argparse.ArgumentParser(description="Extract the sections no spec operation serves.")
    parser.add_argument("document")
    parser.add_argument("contract")
    parser.add_argument("-o", "--out", default=OUT_DEFAULT)
    args = parser.parse_args()
    for path in (args.document, args.contract):
        if os.path.exists(args.out) and os.path.samefile(path, args.out):
            sys.exit("Output must not refer to an input")
    with open(args.contract, encoding="utf-8") as f:
        spec = json.load(f)["endpoints"]
    doc = extract(args.document)
    _, doc_only, _ = pair(spec, doc)
    wanted = {(d["category"], d["name"]): d for d in doc_only if d.get("method")}

    with zipfile.ZipFile(args.document) as z:
        root = etree.fromstring(z.read("word/document.xml"), etree.XMLParser(huge_tree=True))
    body = root.find(qn("body"))
    els = list(body)
    sections, previous = [], {}
    for category, _, title, start, end in section_ranges(body):
        d = wanted.get((category, title))
        if d is not None:
            # The section ends where the next heading of any level starts.
            stop = next((i for i in range(start + 1, end) if els[i].tag == qn("p")
                         and re.match(r"Heading\d$", style_of(els[i]))), end)
            prose, request, response = section_parts(els[start:stop])
            item = {"category": category, "title": title}
            if previous.get(category):
                item["after"] = previous[category]
            else:
                item["first"] = True
            item.update({"method": d["method"], "path": doc_path(d),
                         "description": prose or [title]})
            for table in TABLES:
                rows = [f for f in d[table] if table != "response_fields" or f["name"] not in ERROR_ROWS]
                if rows:
                    item[table] = [{k: f.get(k, "") for k in ("name", "type", "presence", "description")}
                                   for f in rows]
            if request is not None:
                item["request_example"] = request
            if response is not None:
                item["response_example"] = response
            sections.append(item)
        previous[category] = title
    for s in sections:
        if len(s["description"]) == 1:
            s["description"] = s["description"][0]
    validate({"sections": sections})
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(HEADER)
        yaml.safe_dump({"sections": sections}, f, allow_unicode=True, sort_keys=False, width=10000)
    print("planned sections : %d" % len(sections))
    print("with a request sample  : %d" % sum(1 for s in sections if "request_example" in s))
    print("with a response sample : %d" % sum(1 for s in sections if "response_example" in s))
    print("wrote %s" % args.out)


if __name__ == "__main__":
    main()
