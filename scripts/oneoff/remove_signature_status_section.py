#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Remove the `Check status of Signature Field` section.

`GET {baseURL}/v1/packages/{packageId}/documents/{documentId}/fields/signature/status`
does not exist in the backend. The only GET registered under `/fields` is the exact
`.../fields` (Get Fields); there is no GET wildcard, and the string `signature/status`
appears nowhere in the backend source. The document was promising an endpoint that
returns 404. Removal was approved by the project owner on 2026-07-27.

Two references point into the section:

* the Table of Contents (`_Toc202862828`) - `w:updateFields` is set in settings.xml,
  so Word rebuilds the TOC when the file is opened; nothing to do here;
* the Version History table, row v1.9 / 20240917, where the API name is a hyperlink
  to the section. That row records what really happened in v1.9, so the text stays -
  only the hyperlink is unwrapped, and the `Hyperlink` character style is dropped from
  its runs so the text no longer renders as a clickable link.

Both `w:bookmarkStart` and their matching `w:bookmarkEnd` (ids 777, 778) sit inside
the removed block, so this leaves no orphaned bookmarkEnd - the defect hit when the
RemarkError section was removed.

The section is located by heading text rather than by index: another agent edits this
document concurrently, and an index captured at inspection time may no longer be valid.

Usage:
    python scripts/oneoff/remove_signature_status_section.py [--dry-run]
"""
import os
import re
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

HEADING = "Check status of Signature Field"
BOOKMARK = "_Check_status_of"
EXPECTED_NEXT = "Add In-person Field"


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def heading_level(el):
    if el.tag != qn("p"):
        return None
    m = re.match(r"Heading([1-6])$", style_of(el))
    return int(m.group(1)) if m else None


def find_block(body):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if heading_level(el) and HEADING in text_of(el):
            start = i
            break
    if start is None:
        sys.exit("heading %r not found - already removed?" % HEADING)

    lvl = heading_level(kids[start])
    for j in range(start + 1, len(kids)):
        nxt = heading_level(kids[j])
        if nxt and nxt <= lvl:
            return kids, start, j, text_of(kids[j]).strip()
    return kids, start, len(kids), "(end of document)"


def unwrap_hyperlink(root):
    """Replace the version-history hyperlink with its runs, keeping the text."""
    done = 0
    for link in list(root.iter(qn("hyperlink"))):
        if link.get(qn("anchor")) != BOOKMARK:
            continue
        parent = link.getparent()
        pos = list(parent).index(link)
        for offset, child in enumerate(list(link)):
            link.remove(child)
            parent.insert(pos + offset, child)
            # the run still carries rStyle="Hyperlink"; without a target it should
            # not keep rendering as a link
            rPr = child.find(qn("rPr")) if child.tag == qn("r") else None
            if rPr is not None:
                st = rPr.find(qn("rStyle"))
                if st is not None and st.get(qn("val")) == "Hyperlink":
                    rPr.remove(st)
        parent.remove(link)
        done += 1
    return done


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
    body = root.find(qn("body"))

    kids, start, end, next_title = find_block(body)
    block = kids[start:end]

    print("section  : %r" % text_of(kids[start]).strip())
    print("ends before: %r" % next_title)
    print("elements : %d (%d paragraphs, %d tables)" % (
        len(block),
        sum(1 for e in block if e.tag == qn("p")),
        sum(1 for e in block if e.tag == qn("tbl"))))

    if EXPECTED_NEXT not in next_title:
        sys.exit("unexpected next heading %r - the document changed, stopping" % next_title)

    starts = {b.get(qn("id")) for e in block for b in e.iter(qn("bookmarkStart"))}
    ends = {b.get(qn("id")) for e in block for b in e.iter(qn("bookmarkEnd"))}
    print("bookmarks: start %s / end %s" % (sorted(starts), sorted(ends)))
    orphan = starts - ends
    if orphan:
        # bookmarkEnd outside the block would be left dangling; remove those too
        print("bookmarkEnd outside the block to remove: %s" % sorted(orphan))
        for be in list(root.iter(qn("bookmarkEnd"))):
            if be.get(qn("id")) in orphan:
                be.getparent().remove(be)

    for el in block:
        body.remove(el)

    unwrapped = unwrap_hyperlink(root)
    print("version-history hyperlinks unwrapped: %d" % unwrapped)

    left = sum(1 for t in root.iter(qn("t")) if t.text and "signature/status" in t.text)
    print("runs still mentioning signature/status: %d" % left)

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
