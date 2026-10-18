#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Correct four wrong endpoint descriptions in the Document Package module.

Wording follows the SigningHub 10.0 manual, which this specification is
designed against, adapted where GoPaperless deliberately differs.

1. `Get Package Details` carried the description of `Get Package Verification`
   ("obtain the verification results for all digital signature fields...").
   The endpoint actually returns name, owner, document_status and documents[].
   SigningHub: "Business applications can use this service API to get the
   details of an package." Extended here to name what the response contains.

2. `Download Document (base64)` said "download a document in binary format",
   copied from the non-base64 section, and used "Business application" while
   every other section in the module uses the plural.
   SigningHub: "...to download the document in base64 string format. The
   package and document ID is provided in the resource URL."

3. `Update Document` and `Update Document (base64)` both carried
   "Note a package must already exist before you can add a document using this
   call." - copied from `Upload Document`. These endpoints update a document
   that already exists. The other two notes apply to updates as well and are
   left alone. The base64 variant additionally described its body as "document
   bytes"; it takes a base64 string in the `document` field.

4. `Merge Document` said "seamlessly integrate documents into workflow
   processes", which does not describe the operation.
   SigningHub: "...to merge document at any location (TOP | BOTTOM) of the
   target document in a package." GoPaperless uses the `x-append-to-top`
   boolean header instead of SigningHub's `x-location` enum - a recorded
   divergence - so the sentence names our header.

Usage:
    python scripts/oneoff/fix_document_package_descriptions.py [--dry-run]
"""
import os
import shutil
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted"
    "_backup_pre_description_fix.docx")

HEADINGS = {"Heading1", "Heading2", "Heading3"}
MODULE = "Document Package"

# (section, prefix of the paragraph to replace, replacement text)
REWRITES = [
    ("Get Package Details",
     "Business applications can use this service API to obtain the verification",
     "Business applications can use this service API to get the details of a document package, "
     "including its name, owner, status and the list of documents it contains."),

    ("Download Document (base64)",
     "Business application can use this service API to download a document in binary",
     "Business applications can use this service API to download the document in base64 string "
     "format. The package ID and document ID are provided in the resource URL."),

    ("Update Document",
     "Note a package must already exist",
     "Note the document must already exist in the package before you can update it using this call."),

    ("Update Document (base64)",
     "Note a package must already exist",
     "Note the document must already exist in the package before you can update it using this call."),

    ("Update Document (base64)",
     "Business applications can use this service API to update a document of a document package.",
     "Business applications can use this service API to update a document of a document package. "
     "The document information is sent in the HTTP request header and the document content is sent "
     "as a base64 encoded string in the HTTP request body."),

    ("Merge Document",
     "Business applications can use this api to seamlessly integrate",
     "Business applications can use this service API to merge a document at the top or the bottom "
     "of the target document in a package. The x-append-to-top header controls the position."),
]


def el_text(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def para_style(el):
    pPr = el.find(qn("pPr"))
    if pPr is None:
        return ""
    style = pPr.find(qn("pStyle"))
    return style.get(qn("val")) if style is not None else ""


def set_leaf_text(el, value):
    runs = [t for t in el.iter(qn("t"))]
    if not runs:
        return False
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""
    return True


def description_paras(body, heading_text):
    """Paragraphs between the Heading 3 and the `Sample Request` label."""
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        sys.exit("section %r not found" % heading_text)
    out = []
    for j in range(start + 1, len(kids)):
        el = kids[j]
        if el.tag != qn("p"):
            break
        text = el_text(el).strip()
        if text == "Sample Request" or para_style(el) in HEADINGS:
            break
        out.append(el)
    return out


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
    tree = etree.parse(DOC_XML, etree.XMLParser(huge_tree=True))
    body = tree.getroot().find(qn("body"))

    applied = 0
    for heading, prefix, replacement in REWRITES:
        target = None
        for para in description_paras(body, heading):
            if el_text(para).strip().startswith(prefix):
                target = para
                break
        if target is None:
            print("  %-28s %-42s already applied / not found"
                  % (heading, prefix[:40] + "..."))
            continue
        set_leaf_text(target, replacement)
        applied += 1
        print("  %-28s rewritten (%d chars)" % (heading, len(replacement)))

    print("\n%d of %d rewrites applied" % (applied, len(REWRITES)))
    if applied == 0:
        sys.exit("nothing matched - descriptions may already be corrected")

    if dry:
        print("\n--dry-run: nothing written")
        return

    if os.path.exists(OUT_DOCX) and not os.path.exists(BACKUP_DOCX):
        shutil.copy2(OUT_DOCX, BACKUP_DOCX)
        print("\nbacked up %s" % BACKUP_DOCX)

    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
