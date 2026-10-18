#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Three corrections to the attachment endpoints of the `Document Package` section.

1. `Upload Attachment` verb PUT -> POST.
   The backend registers the route in `DocumentController.validURLs` and serves
   it from `doPost`; the Postman collection calls it with POST. The SigningHub
   manual documents a `file` (binary) request body for the same call but never
   quotes a verb, so it is not evidence either way. Owner decision: POST.
   Changing the verb also changes the badge colour - per template 2.8.1 PUT is
   blue-700 and POST is orange-700. Colours are read from
   `scripts/GoPaperless_DOCX_Format_Template.md`, never hardcoded.

2. `Download Attachment` header table cleanup.
   `x-file-name` was copied from `Upload Attachment`; a download does not name
   the file it uploads. Its `Content-Type` description ("...to upload your
   document with base64 format") was copied from the base64 variant and is
   wrong for a binary download, whose sample response is
   `application/octet-stream`. Restored to the wording the other download
   endpoints use.

3. `Download Attachment (base64)` section removed.
   It carried a path identical to `Download Attachment`, so one of the two was
   necessarily wrong. The backend has no `/attachments/{attachmentId}/base64`
   route and the Postman collection has no such request; SigningHub does
   document one. Owner decision: the spec describes what ships, so the section
   is removed rather than repointed at the SigningHub path.

The section's own heading number is automatic, so the body renumbers itself.
The matching TOC3 entry is removed here, but the remaining TOC rows keep their
cached numbers and page numbers until the field is refreshed in Word
(Ctrl+A, then F9).

Usage:
    python scripts/oneoff/fix_document_package_attachments.py [--dry-run]
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
BACKUP_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted"
    "_backup_pre_attachment_fix.docx")

UPLOAD_PATH = "/v1/packages/{packageId}/documents/{documentId}/attachments"
OLD_METHOD = "PUT"
NEW_METHOD = "POST"

# Every attachment section inherited `x-file-name` and the "...to upload your
# document with base64 format" Content-Type wording from `Upload Attachment`.
# Only the upload actually names a file, so the other three are cleaned.
HEADER_FIX_SECTIONS = ("Download Attachment", "Get Attachments", "Delete Attachment")
DROP_HEADING = "Download Attachment (base64)"
NEXT_HEADING = "Merge Document"

DROP_HEADER_ROW = "x-file-name"
CONTENT_TYPE_ROW = "Content-Type"
CONTENT_TYPE_TEXT = "The MIME type of the request body. Usually application/json"

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


def set_cell_text(cell, value):
    """Write `value` into a table cell, keeping the first run's formatting."""
    runs = list(cell.iter(qn("t")))
    if not runs:
        return False
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""
    return True


def row_cells(tr):
    return tr.findall(qn("tc"))


# --- 1) Upload Attachment: PUT -> POST -------------------------------------

def retag_upload_method(root, rules):
    changed = 0
    for txbx in root.iter(qn("txbxContent")):
        text = el_text(txbx)
        if UPLOAD_PATH not in text or not text.lstrip().startswith(OLD_METHOD):
            continue
        paras = txbx.findall(qn("p"))
        if not paras:
            continue
        runs = paras[0].findall(qn("r"))
        if not runs:
            continue
        first = runs[0]
        t = first.find(qn("t"))
        if t is None or (t.text or "").strip() != OLD_METHOD:
            continue

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
        changed += 1
    return changed


# --- 2) Download Attachment: header table cleanup --------------------------

def section_range(body, heading_text):
    """Body-child index range [start, end) of the section under `heading_text`."""
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if is_heading(el) and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        return None
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if is_heading(kids[j]):
            end = j
            break
    return start, end


def fix_header_table(body, heading):
    """Drop the x-file-name row, renumber, restore the Content-Type wording.

    Returns the surviving row count, or None when the section was already
    clean so the script stays re-runnable.
    """
    span = section_range(body, heading)
    if span is None:
        sys.exit("section %r not found" % heading)
    start, end = span
    kids = list(body)

    for el in kids[start:end]:
        if el.tag != qn("tbl"):
            continue
        rows = el.findall(qn("tr"))
        if not rows or not el_text(rows[0]).strip().startswith("Header Request Attributes"):
            continue

        dropped = False
        for tr in list(rows):
            cells = row_cells(tr)
            if len(cells) >= 2 and el_text(cells[1]).strip() == DROP_HEADER_ROW:
                el.remove(tr)
                dropped = True
                break
        if not dropped:
            return None

        # renumber the No column and repair the Content-Type description
        data_rows = [tr for tr in el.findall(qn("tr")) if len(row_cells(tr)) >= 5][1:]
        retyped = False
        for n, tr in enumerate(data_rows, start=1):
            cells = row_cells(tr)
            set_cell_text(cells[0], str(n))
            if el_text(cells[1]).strip() == CONTENT_TYPE_ROW:
                set_cell_text(cells[4], CONTENT_TYPE_TEXT)
                retyped = True
        if not retyped:
            sys.exit("%r row not found while repairing descriptions in %r"
                     % (CONTENT_TYPE_ROW, heading))
        return len(data_rows)

    sys.exit("no Header Request Attributes table inside %r" % heading)


# --- 3) Download Attachment (base64): remove the section -------------------

def drop_base64_section(body):
    span = section_range(body, DROP_HEADING)
    if span is None:
        return 0
    start, end = span
    kids = list(body)

    following = kids[end] if end < len(kids) else None
    if following is None or el_text(following).strip() != NEXT_HEADING:
        sys.exit("expected %r after %r, found %r"
                 % (NEXT_HEADING, DROP_HEADING,
                    el_text(following).strip() if following is not None else None))

    # The range already ends on the blank `Normal` separator that sits in front
    # of the next heading, so removing [start, end) leaves the separator that
    # preceded this section to serve the next one - exactly one, per template
    # 2.6. Extending backwards as well would leave none.
    for el in kids[start:end]:
        body.remove(el)
    return end - start


def drop_toc_entry(root):
    """Remove the TOC row. The TOC lives inside a w:sdt, not directly in body."""
    for el in root.iter(qn("p")):
        if para_style(el) == "TOC3" and DROP_HEADING in el_text(el):
            el.getparent().remove(el)
            return True
    return False


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
    body = root.find(qn("body"))

    # Word stores each frame twice - once under mc:Choice (DrawingML) and once
    # under mc:Fallback (VML) - so one visible frame is two txbxContent copies
    # and both have to carry the new badge.
    retagged = retag_upload_method(root, rules)
    if retagged == 0:
        print("1) Upload Attachment: already POST")
    elif retagged == 2:
        print("1) Upload Attachment frames retagged %s -> %s : 2 (Choice + Fallback)"
              % (OLD_METHOD, NEW_METHOD))
    else:
        sys.exit("expected the Choice/Fallback pair of 1 request frame, got %d" % retagged)

    for heading in HEADER_FIX_SECTIONS:
        kept = fix_header_table(body, heading)
        if kept is None:
            print("2) %s: header table already clean" % heading)
        else:
            print("2) %s: dropped %r, %d rows renumbered, Content-Type restored"
                  % (heading, DROP_HEADER_ROW, kept))

    removed = drop_base64_section(body)
    toc = drop_toc_entry(root)
    if removed == 0 and not toc:
        print("3) %r: already removed" % DROP_HEADING)
    else:
        print("3) %r removed: %d body blocks, TOC entry %s"
              % (DROP_HEADING, removed, "removed" if toc else "NOT FOUND"))
        if not toc:
            sys.exit("TOC entry for %r not found - aborting rather than "
                     "leaving a dangling link" % DROP_HEADING)

    if dry:
        print("\n--dry-run: nothing written")
        return

    if os.path.exists(OUT_DOCX) and not os.path.exists(BACKUP_DOCX):
        shutil.copy2(OUT_DOCX, BACKUP_DOCX)
        print("\nbacked up %s" % BACKUP_DOCX)

    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
