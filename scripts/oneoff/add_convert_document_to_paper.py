#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Document the `Convert Document to Paper` endpoint in the Document Package section.

The route is registered in `DocumentController.validURLs[31]` and dispatched to
`DocumentService.convertToPaper`, but no section described it. Facts taken from
the backend, read-only:

  * `POST /v1/packages/{packageId}/documents/{documentId}/paper`
  * `verifyDocumentAccess(..., true, false, ...)` -> Authorization is mandatory
  * body is the `ConvertToPaperRequest` record, one optional field
    `converter_display_name`; when absent the backend falls back to the
    authenticated user's name and then their email
  * on success the controller sets `ContentType.APPLICATION_OCTET_STREAM` and
    returns the stamped PDF bytes, so the success frame is binary, not JSON

SigningHub documents no equivalent call, so nothing here is derived from it.
Owner placed the endpoint in Document Package.

The section is cloned from `Merge Document` rather than synthesised: a request
frame is an mc:AlternateContent pair (DrawingML Choice + VML Fallback) whose
badge shading, bold URL segments and red placeholders would be easy to get
wrong by hand. `Merge Document` already carries the same POST verb and the same
`{packageId}`/`{documentId}` placeholders, so only leaf text changes. The
missing Request Attributes table is cloned from `Add Document Package`.

Heading numbering is automatic and TOC rows are a field result, so no TOC row is
written here - refresh the field in Word (Ctrl+A, then F9).

Usage:
    python scripts/oneoff/add_convert_document_to_paper.py [--dry-run]
"""
import copy
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
    "_backup_pre_convert_to_paper.docx")

DONOR = "Merge Document"
FIELD_DONOR = "Add Document Package"
NEW_HEADING = "Convert Document to Paper"
SECTION_H2 = "Document Package"

DESCRIPTION = ("Business applications can use this service API to produce a paper copy of a "
               "signed document. The returned PDF carries an appended conversion page that "
               "records who converted the document and a QR code linking back to the workflow.")

BODY_FIELD = "converter_display_name"
BODY_LINE = '  "%s": "String"' % BODY_FIELD
BODY_DESC = ("Name shown on the conversion page as the person who converted the document. "
             "When omitted, the name of the authenticated user is used, falling back to "
             "their email address.")

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


def set_leaf_text(el, value):
    """Write `value` into `el`, keeping the first run's formatting."""
    runs = [t for t in el.iter(qn("t"))]
    if not runs:
        return False
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""
    return True


def row_cells(tr):
    return tr.findall(qn("tc"))


def section_range(body, heading_text):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        return None
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if kids[j].tag == qn("p") and para_style(kids[j]) in HEADINGS:
            end = j
            break
    return start, end


def find_table(blocks, label):
    for el in blocks:
        if el.tag != qn("tbl"):
            continue
        rows = el.findall(qn("tr"))
        if rows and el_text(rows[0]).strip().startswith(label):
            return el
    return None


def strip_bookmarks(blocks):
    """A cloned heading would otherwise duplicate its _Toc bookmark identity."""
    n = 0
    for el in blocks:
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for bm in el.iter(qn(tag)):
                bm.getparent().remove(bm)
                n += 1
    return n


def choice_or_fallback_frames(block):
    """Every request/response frame exists twice; both copies need the edit."""
    return list(block.iter(qn("txbxContent")))


def rewrite_request_frame(txbx):
    paras = txbx.findall(qn("p"))
    # p0 ends with the literal path tail, p1 is Content-Type, p4 is the
    # merge-only x-append-to-top header, p5..p7 are the body block.
    tail = [t for t in paras[0].iter(qn("t")) if (t.text or "") == "/merge"]
    if not tail:
        sys.exit("donor request frame no longer ends in '/merge'")
    tail[0].text = "/paper"

    for t in paras[1].iter(qn("t")):
        if (t.text or "").startswith(": application/"):
            t.text = ": application/json"

    for para in list(paras):
        if el_text(para).startswith("x-append-to-top"):
            para.getparent().remove(para)

    for para in txbx.findall(qn("p")):
        if el_text(para).strip() == "File in binary format":
            set_leaf_text(para, BODY_LINE)


def rewrite_response_frame(txbx):
    """Only the success frame changes; the error frame is already correct."""
    texts = [el_text(p).strip() for p in txbx.findall(qn("p"))]
    if "{}" not in texts:
        return False
    for para in txbx.findall(qn("p")):
        body = el_text(para).strip()
        if body.startswith("Content-type:") or body.startswith("Content-Type:"):
            set_leaf_text(para, "Content-type: application/octet-stream")
        elif body == "{}":
            set_leaf_text(para, "{File in binary format}")
    return True


def build_section(body):
    span = section_range(body, DONOR)
    if span is None:
        sys.exit("donor section %r not found" % DONOR)
    start, end = span
    blocks = [copy.deepcopy(el) for el in list(body)[start:end]]
    strip_bookmarks(blocks)

    # heading + description
    set_leaf_text(blocks[0], NEW_HEADING)
    set_leaf_text(blocks[1], DESCRIPTION)

    # request frame lives in the block after the "Sample Request" label
    req = next(b for b in blocks if el_text(b).lstrip().startswith("POST"))
    frames = choice_or_fallback_frames(req)
    if len(frames) != 2:
        sys.exit("expected a Choice/Fallback request frame pair, got %d" % len(frames))
    for frame in frames:
        rewrite_request_frame(frame)

    # response frames: success turns binary, error frame is left alone
    resp = next(b for b in blocks
                if b.tag == qn("p") and el_text(b).lstrip().startswith("Http Code:"))
    rewritten = sum(1 for frame in choice_or_fallback_frames(resp)
                    if rewrite_response_frame(frame))
    if rewritten != 2:
        sys.exit("expected 2 success response frames to rewrite, got %d" % rewritten)

    # header table: drop the merge-only row, renumber, fix Content-Type wording
    header = find_table(blocks, "Header Request Attributes")
    if header is None:
        sys.exit("donor has no Header Request Attributes table")
    for tr in list(header.findall(qn("tr"))):
        cells = row_cells(tr)
        if len(cells) >= 2 and el_text(cells[1]).strip() == "x-append-to-top":
            header.remove(tr)
    data_rows = [tr for tr in header.findall(qn("tr")) if len(row_cells(tr)) >= 5][1:]
    for n, tr in enumerate(data_rows, start=1):
        cells = row_cells(tr)
        set_leaf_text(cells[0], str(n))
        if el_text(cells[1]).strip() == "Content-Type":
            set_leaf_text(cells[4], CONTENT_TYPE_TEXT)

    # Request Attributes table, cloned from a section that has one, inserted
    # between Path and Response per template 2.7 table order
    field_span = section_range(body, FIELD_DONOR)
    if field_span is None:
        sys.exit("field donor section %r not found" % FIELD_DONOR)
    field_blocks = list(body)[field_span[0]:field_span[1]]
    req_tbl = find_table(field_blocks, "Request Attributes")
    if req_tbl is None:
        sys.exit("%r has no Request Attributes table to clone" % FIELD_DONOR)
    req_tbl = copy.deepcopy(req_tbl)
    strip_bookmarks([req_tbl])
    rows = [tr for tr in req_tbl.findall(qn("tr")) if len(row_cells(tr)) >= 5]
    for tr in rows[2:]:
        req_tbl.remove(tr)
    cells = row_cells(rows[1])
    set_leaf_text(cells[0], "1")
    set_leaf_text(cells[1], BODY_FIELD)
    set_leaf_text(cells[2], "String")
    set_leaf_text(cells[3], "O")
    set_leaf_text(cells[4], BODY_DESC)

    path_tbl = find_table(blocks, "Path Attributes")
    if path_tbl is None:
        sys.exit("donor has no Path Attributes table")
    at = blocks.index(path_tbl)
    separator = copy.deepcopy(blocks[at + 1])   # the blank Normal after the table
    blocks[at + 1:at + 1] = [separator, req_tbl]
    return blocks


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

    if section_range(body, NEW_HEADING) is not None:
        sys.exit("%r is already documented - nothing to do" % NEW_HEADING)

    blocks = build_section(body)
    print("built %d blocks for %r from donor %r" % (len(blocks), NEW_HEADING, DONOR))

    # append at the end of the Document Package section, i.e. immediately
    # before the Heading2 that opens the next section
    kids = list(body)
    h2 = [i for i, el in enumerate(kids)
          if el.tag == qn("p") and para_style(el) == "Heading2"
          and el_text(el).strip() == SECTION_H2]
    if not h2:
        sys.exit("Heading2 %r not found" % SECTION_H2)
    nxt = next((i for i in range(h2[0] + 1, len(kids))
                if kids[i].tag == qn("p") and para_style(kids[i]) == "Heading2"), None)
    if nxt is None:
        sys.exit("no Heading2 after %r - refusing to guess the section end" % SECTION_H2)
    print("inserting before %r" % el_text(kids[nxt]).strip())

    anchor = kids[nxt]
    for el in blocks:
        anchor.addprevious(el)

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
