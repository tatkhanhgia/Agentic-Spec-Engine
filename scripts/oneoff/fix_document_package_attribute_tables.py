#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Four attribute-table corrections in the Document Package module.

1. `Share Document Package/ Share Workflow` documented its path variable as
   `package_name` (String, "Defines the name of the package."), but the path is
   `/v1/packages/{packageId}/workflow/share`. Replaced with `packageId`, reusing
   the description every other section already uses for that variable.

2. `Apply Workflow Template` declared the `x-append-to-top` header. The backend
   reads that header in exactly one place, inside the merge flow
   (`DocumentService` merge branch), so the row was copied from `Merge Document`
   and is removed.

3. `Upload Document` and `Upload Document (base64)` named the response field
   `documentid`. The backend serialises a `Document` object whose id carries
   `@JsonProperty("document_id")` (`Document.java:112`). Renamed in the table and
   in the sample response frames. The remaining fields that object emits are
   deliberately left undocumented - owner decision.

4. `Get Document Image` and `Get Document Image Base64` listed path variables as
   packageId, pageNo, documentId while the path orders them packageId,
   documentId, pageNo. Rows reordered to match the URL.

Deliberately NOT changed, per owner:
  * `Get Document Verification` / `Get Package Verification` return byte[], so
    their Response Attributes carry only the error rows.
  * `Pre Upload Document` keeps the `RC` presence value - it is a project
    specific requirement, not a typo.
  * `Download Package` / `Download Document` keep `x-password` and `x-otp`
    even though the backend does not read them yet.

Usage:
    python scripts/oneoff/fix_document_package_attribute_tables.py [--dry-run]
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
    "_backup_pre_attribute_table_fix.docx")

HEADINGS = {"Heading1", "Heading2", "Heading3"}

SHARE = "Share Document Package/ Share Workflow"
TEMPLATE = "Apply Workflow Template"
UPLOADS = ("Upload Document", "Upload Document (base64)")
IMAGES = ("Get Document Image", "Get Document Image Base64")
PACKAGE_ID_DONOR = "Download Package"      # any section with a correct packageId row

DROP_HEADER = "x-append-to-top"
OLD_FIELD = "documentid"
NEW_FIELD = "document_id"


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


def section_blocks(body, heading_text):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        sys.exit("section %r not found" % heading_text)
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if kids[j].tag == qn("p") and para_style(kids[j]) in HEADINGS:
            end = j
            break
    return kids[start:end]


def find_table(blocks, label):
    for el in blocks:
        if el.tag != qn("tbl"):
            continue
        rows = el.findall(qn("tr"))
        if rows and el_text(rows[0]).strip().startswith(label):
            return el
    return None


def data_rows(tbl):
    return [tr for tr in tbl.findall(qn("tr")) if len(row_cells(tr)) >= 5][1:]


def renumber(tbl):
    for n, tr in enumerate(data_rows(tbl), start=1):
        set_leaf_text(row_cells(tr)[0], str(n))


def named_row(tbl, name):
    for tr in data_rows(tbl):
        if el_text(row_cells(tr)[1]).strip() == name:
            return tr
    return None


# --- 1) Share Document Package: package_name -> packageId -------------------

def fix_share_path(body):
    donor_tbl = find_table(section_blocks(body, PACKAGE_ID_DONOR), "Path Attributes")
    donor = named_row(donor_tbl, "packageId")
    if donor is None:
        sys.exit("no packageId row in %r to copy the description from" % PACKAGE_ID_DONOR)
    desc = el_text(row_cells(donor)[4]).strip()

    tbl = find_table(section_blocks(body, SHARE), "Path Attributes")
    if tbl is None:
        sys.exit("%r has no Path Attributes table" % SHARE)
    row = named_row(tbl, "package_name")
    if row is None:
        return False
    cells = row_cells(row)
    set_leaf_text(cells[1], "packageId")
    set_leaf_text(cells[2], "Integer")
    set_leaf_text(cells[4], desc)
    return True


# --- 2) Apply Workflow Template: drop x-append-to-top -----------------------

def fix_template_header(body):
    tbl = find_table(section_blocks(body, TEMPLATE), "Header Request Attributes")
    if tbl is None:
        sys.exit("%r has no Header Request Attributes table" % TEMPLATE)
    row = named_row(tbl, DROP_HEADER)
    if row is None:
        return False
    tbl.remove(row)
    renumber(tbl)
    return True


# --- 3) Upload Document: documentid -> document_id --------------------------

def fix_upload_field(body, heading):
    blocks = section_blocks(body, heading)
    changed = 0

    tbl = find_table(blocks, "Response Attributes")
    if tbl is None:
        sys.exit("%r has no Response Attributes table" % heading)
    row = named_row(tbl, OLD_FIELD)
    if row is not None:
        set_leaf_text(row_cells(row)[1], NEW_FIELD)
        changed += 1

    # the sample response frames spell the field out too
    for el in blocks:
        for t in el.iter(qn("t")):
            if OLD_FIELD in (t.text or ""):
                t.text = t.text.replace(OLD_FIELD, NEW_FIELD)
                changed += 1
    return changed


# --- 4) Get Document Image: path variables in URL order ---------------------

def fix_image_path_order(body, heading):
    tbl = find_table(section_blocks(body, heading), "Path Attributes")
    if tbl is None:
        sys.exit("%r has no Path Attributes table" % heading)
    order = ["packageId", "documentId", "pageNo"]
    rows = data_rows(tbl)
    if [el_text(row_cells(r)[1]).strip() for r in rows] == order:
        return False
    by_name = {el_text(row_cells(r)[1]).strip(): r for r in rows}
    if set(by_name) != set(order):
        sys.exit("%r path rows are %s, expected %s" % (heading, sorted(by_name), sorted(order)))
    # Re-append in the wanted order. Anchoring on one of the moving rows would
    # cancel itself out, since a row cannot be inserted before itself.
    for name in order:
        tbl.append(by_name[name])
    if [el_text(row_cells(r)[1]).strip() for r in data_rows(tbl)] != order:
        sys.exit("%r reorder did not take effect" % heading)
    renumber(tbl)
    return True


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

    print("1) %s: package_name -> packageId : %s"
          % (SHARE, "done" if fix_share_path(body) else "already correct"))
    print("2) %s: drop %r : %s"
          % (TEMPLATE, DROP_HEADER, "done" if fix_template_header(body) else "already correct"))
    for h in UPLOADS:
        n = fix_upload_field(body, h)
        print("3) %s: %r -> %r in %d place(s)" % (h, OLD_FIELD, NEW_FIELD, n))
    for h in IMAGES:
        print("4) %s: path rows reordered : %s"
              % (h, "done" if fix_image_path_order(body, h) else "already correct"))

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
