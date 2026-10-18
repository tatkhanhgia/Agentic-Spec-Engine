#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Document the two `Download Signature Summary` endpoints in Document Package.

Both routes exist in the backend and in the Postman collection but had no
section in the specification. Facts taken from the backend, read-only:

  * `GET /v1/packages/{packageId}/documents/{documentId}/signature-summary`
    `DocumentController.validURLs[29]` -> `DocumentService.downloadSignatureSummary`.
    Auth uses `VerifyStrategy.BLOCK_CLIENT_CREDENTIAL`, so a client-credentials
    token is rejected; the call needs a user token. The controller sets
    `APPLICATION_OCTET_STREAM` on success, so the body is the PDF itself.
  * `GET .../signature-summary/base64`
    `validURLs[30]` -> `downloadSignatureSummaryBase64`, which delegates to the
    call above and re-wraps the bytes as `{"file_data": "<base64>"}`. The
    controller leaves the content type at JSON.

Each section is cloned from the closest existing endpoint rather than
synthesised, because a request frame is an mc:AlternateContent pair whose badge
shading and red placeholders are easy to break by hand:

  * binary variant  <- `Download Attachment`        (GET, octet-stream success)
  * base64 variant  <- `Download Document (base64)` (GET, `file_data` response)

Naming follows the document's own `(base64)` convention rather than the backend
apiName or Postman's `Base64`; the project-wide naming question is still open.

Usage:
    python scripts/oneoff/add_signature_summary_sections.py [--dry-run]
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
    "_backup_pre_signature_summary.docx")

SECTION_H2 = "Document Package"
HEADINGS = {"Heading1", "Heading2", "Heading3"}

BINARY_NAME = "Download Signature Summary"
BINARY_DONOR = "Download Attachment"
BINARY_DESC = ("Business applications can use this service API to download the signature "
               "summary of a document as a PDF file. The summary lists the signatures "
               "applied to the document. This API requires a user access token; a token "
               "obtained through Client Credentials Authentication is rejected.")

BASE64_NAME = "Download Signature Summary (base64)"
BASE64_DONOR = "Download Document (base64)"
BASE64_DESC = ("Business applications can use this service API to download the signature "
               "summary of a document as a base64 encoded string. The content is the same "
               "PDF returned by Download Signature Summary. This API requires a user access "
               "token; a token obtained through Client Credentials Authentication is rejected.")


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
    for el in blocks:
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for bm in el.iter(qn(tag)):
                bm.getparent().remove(bm)


def clone_section(body, donor):
    span = section_range(body, donor)
    if span is None:
        sys.exit("donor section %r not found" % donor)
    start, end = span
    blocks = [copy.deepcopy(el) for el in list(body)[start:end]]
    strip_bookmarks(blocks)
    return blocks


def request_frames(blocks):
    req = next(b for b in blocks if el_text(b).lstrip().startswith("GET"))
    frames = list(req.iter(qn("txbxContent")))
    if len(frames) != 2:
        sys.exit("expected a Choice/Fallback request frame pair, got %d" % len(frames))
    return frames


def drop_frame_lines(txbx, prefixes):
    for para in list(txbx.findall(qn("p"))):
        if any(el_text(para).startswith(p) for p in prefixes):
            para.getparent().remove(para)


def drop_table_rows(tbl, names):
    """Remove rows whose Name cell matches, then renumber the No column."""
    for tr in list(tbl.findall(qn("tr"))):
        cells = row_cells(tr)
        if len(cells) >= 2 and el_text(cells[1]).strip() in names:
            tbl.remove(tr)
    data_rows = [tr for tr in tbl.findall(qn("tr")) if len(row_cells(tr)) >= 5][1:]
    for n, tr in enumerate(data_rows, start=1):
        set_leaf_text(row_cells(tr)[0], str(n))
    return len(data_rows)


def build_binary(body):
    blocks = clone_section(body, BINARY_DONOR)
    set_leaf_text(blocks[0], BINARY_NAME)
    set_leaf_text(blocks[1], BINARY_DESC)

    # `/attachments/{attachmentId}` -> `/signature-summary`
    for frame in request_frames(blocks):
        line = frame.findall(qn("p"))[0]
        hit = False
        for t in line.iter(qn("t")):
            if (t.text or "") == "/attachments/":
                t.text = "/signature-summary"
                hit = True
            elif (t.text or "") == "{attachmentId}":
                t.text = ""
        if not hit:
            sys.exit("donor request line no longer contains '/attachments/'")

    path_tbl = find_table(blocks, "Path Attributes")
    if path_tbl is None:
        sys.exit("%r has no Path Attributes table" % BINARY_DONOR)
    left = drop_table_rows(path_tbl, {"attachmentId"})
    if left != 2:
        sys.exit("expected packageId + documentId to remain, got %d rows" % left)
    return blocks


def build_base64(body):
    blocks = clone_section(body, BASE64_DONOR)
    set_leaf_text(blocks[0], BASE64_NAME)
    set_leaf_text(blocks[1], BASE64_DESC)

    # `/base64` -> `/signature-summary/base64`; this endpoint takes no document
    # password or OTP, and carries no request body.
    for frame in request_frames(blocks):
        line = frame.findall(qn("p"))[0]
        hit = False
        for t in line.iter(qn("t")):
            if (t.text or "") == "/base64":
                t.text = "/signature-summary/base64"
                hit = True
        if not hit:
            sys.exit("donor request line no longer ends in '/base64'")
        drop_frame_lines(frame, ("x-password", "x-otp"))
        for para in list(frame.findall(qn("p"))):
            if el_text(para).strip() == "{}":
                para.getparent().remove(para)

    header_tbl = find_table(blocks, "Header Request Attributes")
    if header_tbl is None:
        sys.exit("%r has no Header Request Attributes table" % BASE64_DONOR)
    left = drop_table_rows(header_tbl, {"x-password", "x-otp"})
    if left != 3:
        sys.exit("expected 3 header rows to remain, got %d" % left)
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

    for name in (BINARY_NAME, BASE64_NAME):
        if section_range(body, name) is not None:
            sys.exit("%r is already documented - nothing to do" % name)

    binary = build_binary(body)
    base64_ = build_base64(body)
    print("built %d + %d blocks" % (len(binary), len(base64_)))

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
    for el in binary + base64_:
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
