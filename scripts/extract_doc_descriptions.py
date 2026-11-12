#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extract a document's field descriptions into the description file.

The spec (`openapi.yaml`) describes every field in words checked against the
code. The approved document words many of the same fields differently, and
the owner keeps that wording. `config/doc-descriptions.yaml` holds it, so
the renderer needs no earlier DOCX to know it:

    sections:
      <Heading 2>:
        <Heading 3>  or  <Heading 3> > <Heading 4>:
          headers | path_params | request_fields | response_fields | response_headers:
            <field>: one line, or a list of lines

An entry exists only where the document's wording differs from the spec's.
The renderer uses the entry when there is one and the spec otherwise, so
deleting an entry hands that field back to the spec.

THIS SCRIPT ONLY READS the document and the contract.

Usage:
    python scripts/extract_doc_descriptions.py <document.docx> <from_openapi.json> [-o out.yaml]
"""
import argparse
import json
import os
import sys

import yaml

from render_from_openapi import seed_descriptions

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DEFAULT = os.path.join(ROOT, "config", "doc-descriptions.yaml")

HEADER = """\
# Document wording of field descriptions, per section and table.
# Generated once by scripts/extract_doc_descriptions.py from the approved
# document; edited by hand afterwards.
#
# The renderer (scripts/render_from_openapi.py --descriptions) shows the entry
# when a field has one, and the description in openapi.yaml otherwise. Delete
# an entry to hand the field back to the spec. A value is one line, or a list
# of lines shown as separate paragraphs.
"""


def plain(node):
    """OrderedDicts to dicts, and a single line to a string."""
    if isinstance(node, dict):
        return {k: plain(v) for k, v in node.items()}
    if isinstance(node, list):
        return node[0] if len(node) == 1 else list(node)
    return node


def main():
    parser = argparse.ArgumentParser(description="Extract document wording into the description file.")
    parser.add_argument("document")
    parser.add_argument("contract")
    parser.add_argument("-o", "--out", default=OUT_DEFAULT)
    args = parser.parse_args()
    for path in (args.document, args.contract):
        if os.path.exists(args.out) and os.path.samefile(path, args.out):
            sys.exit("Output must not refer to an input")
    with open(args.contract, encoding="utf-8") as f:
        contract = json.load(f)
    sections = seed_descriptions(args.document, contract)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(HEADER)
        yaml.safe_dump({"sections": plain(sections)}, f, allow_unicode=True, sort_keys=False, width=10000)
    count = sum(len(fields) for tables in sections.values() for kinds in tables.values()
                for fields in kinds.values())
    print("sections with wording : %d" % sum(len(t) for t in sections.values()))
    print("descriptions          : %d" % count)
    print("wrote %s" % args.out)


if __name__ == "__main__":
    main()
