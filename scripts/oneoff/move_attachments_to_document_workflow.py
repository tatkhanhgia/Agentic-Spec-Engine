#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Move the attachment endpoints from Document Package to Document Workflow.

The SigningHub manual's own section index lists `Get Attachments`,
`Upload Attachment`, `Download Attachment`, `Download Attachment (base64)` and
`Delete Attachment` under **Document Workflow**, not Document Package, where
this specification had them. Owner decided to follow SigningHub.

`Download Attachment (base64)` was removed earlier - it has no backend route -
so four sections move.

These are moves, not copies: the elements are relocated with lxml so each
heading keeps its `_Toc` bookmark and the TOC keeps resolving. Nothing is
deep-copied and no bookmark is stripped.

Separator hygiene works out on its own. Each section runs `[H3]...[tables]
[Normal]`, the cluster is contiguous, and the blank `Normal` in front of it
belongs to the preceding section - so removing the cluster leaves exactly one
separator before `Merge Document`, and appending it at the end of Document
Workflow puts one separator in front of the first moved heading.

Section numbering is automatic, so both modules renumber themselves. TOC rows
keep cached numbers and page numbers until the field is refreshed in Word
(Ctrl+A, then F9).

Usage:
    python scripts/oneoff/move_attachments_to_document_workflow.py [--dry-run]
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
    "_backup_pre_attachment_move.docx")

MOVING = ("Upload Attachment", "Get Attachments", "Delete Attachment", "Download Attachment")
FROM_H2 = "Document Package"
TO_H2 = "Document Workflow"
HEADINGS = {"Heading1", "Heading2", "Heading3"}


def el_text(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def para_style(el):
    pPr = el.find(qn("pPr"))
    if pPr is None:
        return ""
    style = pPr.find(qn("pStyle"))
    return style.get(qn("val")) if style is not None else ""


def is_heading(el):
    return el.tag == qn("p") and para_style(el) in HEADINGS


def section_span(kids, heading_text):
    start = None
    for i, el in enumerate(kids):
        if is_heading(el) and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        sys.exit("section %r not found" % heading_text)
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if is_heading(kids[j]):
            end = j
            break
    return start, end


def module_span(kids, h2_text):
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) == "Heading2" \
                and el_text(el).strip() == h2_text:
            start = i
            break
    if start is None:
        sys.exit("Heading2 %r not found" % h2_text)
    end = next((j for j in range(start + 1, len(kids))
                if kids[j].tag == qn("p") and para_style(kids[j]) == "Heading2"), len(kids))
    return start, end


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
    root = tree.getroot()
    body = root.find(qn("body"))
    kids = list(body)

    src_start, src_end = module_span(kids, FROM_H2)
    dst_start, dst_end = module_span(kids, TO_H2)

    spans = [section_span(kids, name) for name in MOVING]
    for name, (a, b) in zip(MOVING, spans):
        if not (src_start < a and b <= src_end):
            sys.exit("%r is not inside %r - refusing to move" % (name, FROM_H2))
        print("  %-22s blocks %d..%d" % (name, a, b - 1))

    lo = min(a for a, _ in spans)
    hi = max(b for _, b in spans)
    if hi - lo != sum(b - a for a, b in spans):
        sys.exit("the four sections are not contiguous - refusing to move a mixed range")
    moving = kids[lo:hi]
    print("cluster: %d..%d (%d blocks)" % (lo, hi - 1, len(moving)))

    # what the source is left with
    print("source after removal: %r -> %r"
          % (para_style(kids[lo - 1]) or kids[lo - 1].tag.split('}')[1],
             el_text(kids[hi]).strip()[:40]))

    anchor = kids[dst_end] if dst_end < len(kids) else None
    if anchor is None:
        sys.exit("no Heading2 after %r - refusing to guess the module end" % TO_H2)
    print("destination: appending before %r" % el_text(anchor).strip())

    # relocate in place; lxml detaches each element from its old parent
    for el in moving:
        anchor.addprevious(el)

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
