#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Align the RELATIVE OBJECT definitions that Document Package endpoints point at.

The endpoint sections were corrected earlier in the session, but the objects
their composite types resolve to were not, so the sample bodies and the object
tables disagreed. Both sides now follow the backend serialisers.

1. `Package` (target of `Get All Packages` -> `packages`)
   vs `GetListPackageSerializer`:
     * `unread` is never written by the backend - removed.
     * `next_signer` is never written either; the line is commented out in the
       serialiser and only `next_signer_email` survives - removed.
     * `owner_email` IS written (right after `package_owner`) but was missing
       from the object - added.
     * the `decline` type pointed at a lowercase name; the object is `Decline`.

2. `Document(Package Details)` (target of `Get Package Details` -> `documents`)
   vs `GetPackageDetailsSerializer`:
     * the serialiser writes `uploaded_on` / `uploaded_by`, not
       `created_at` / `created_by`.
     * `id` is written with `writeNumberField`, so it is an Integer.

3. `Get All Packages` referenced `List<Packages>`; the object is `Package`.

4. `Get Package Details` referenced `owner` and `List<document>`; the objects
   are `Owner(Package Details)` and `Document(Package Details)`. The plain
   `Document` object stays as it is - three Document Workflow endpoints point
   at it and that module is owned by other work.

Note for the backend: `package_owner` and `owner_email` are both assigned the
same `ownerEmail` value, the same duplication already reported for
`owner.name` / `owner.email` in Get Package Details.

Usage:
    python scripts/oneoff/fix_document_package_relative_objects.py [--dry-run]
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guard_load, save as guard_save, ConcurrentModification

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
    "_backup_pre_relative_object_fix.docx")

CHAPTER = "RELATIVE OBJECT"
HEADINGS = {"Heading1", "Heading2", "Heading3"}

PACKAGE_DROP = ("unread", "next_signer")
OWNER_EMAIL_DESC = "Email of the user who owns this package."

DOC_DETAILS_RENAME = {
    "created_at": ("uploaded_on", "Date and time when the document was uploaded."),
    "created_by": ("uploaded_by", "Name of the user who uploaded this document."),
}

# The Type cells are NOT matched to objects by their text. Each carries a
# w:hyperlink whose w:anchor points at a bookmark on the object's heading, and
# all three already resolve correctly:
#   owner          -> _Owner(Package_Details)    -> "Owner(Package Details)"
#   List<document> -> _Document(Package_Details) -> "Document(Package Details)"
#   List<Packages> -> _Items                     -> "Package"
# The text is a display label, so there is nothing to rename here. Rewriting it
# would also strip the runs inside the hyperlink.


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


def object_table(body, title):
    """First table under a Heading 2 inside the RELATIVE OBJECT chapter."""
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) == "Heading1" \
                and el_text(el).strip() == CHAPTER:
            start = i
            break
    if start is None:
        sys.exit("chapter %r not found" % CHAPTER)
    at = None
    for i in range(start, len(kids)):
        el = kids[i]
        if el.tag == qn("p") and para_style(el) == "Heading2" \
                and el_text(el).strip() == title:
            at = i
            break
    if at is None:
        sys.exit("object %r not found in %s" % (title, CHAPTER))
    for j in range(at + 1, len(kids)):
        if kids[j].tag == qn("tbl"):
            return kids[j]
        if kids[j].tag == qn("p") and para_style(kids[j]) in HEADINGS:
            break
    sys.exit("object %r has no table" % title)


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


def endpoint_table(body, heading, label):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading:
            start = i
            break
    if start is None:
        sys.exit("section %r not found" % heading)
    for j in range(start + 1, len(kids)):
        el = kids[j]
        if el.tag == qn("p") and para_style(el) in HEADINGS:
            break
        if el.tag == qn("tbl"):
            rows = el.findall(qn("tr"))
            if rows and el_text(rows[0]).strip().startswith(label):
                return el
    sys.exit("%r has no %s table" % (heading, label))


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
    # Another agent edits this file. docx_guard remembers the digest read here
    # and refuses to write if anything touched the file in the meantime.
    tree, token = guard_load(DOC_XML)
    body = tree.getroot().find(qn("body"))

    # --- 1) Package -------------------------------------------------------
    tbl = object_table(body, "Package")
    dropped = []
    for name in PACKAGE_DROP:
        tr = named_row(tbl, name)
        if tr is not None:
            tbl.remove(tr)
            dropped.append(name)

    if named_row(tbl, "owner_email") is None:
        anchor = named_row(tbl, "package_owner")
        if anchor is None:
            sys.exit("Package has no package_owner row to anchor owner_email after")
        new = copy.deepcopy(anchor)
        cells = row_cells(new)
        set_leaf_text(cells[1], "owner_email")
        set_leaf_text(cells[2], "String")
        set_leaf_text(cells[3], "M")
        set_leaf_text(cells[4], OWNER_EMAIL_DESC)
        anchor.addnext(new)
        added = True
    else:
        added = False

    # The `decline` type cell is a hyperlink labelled "decline" whose anchor
    # `_WorkflowDetails` already resolves to the `Decline` heading, and
    # `access_duration` likewise resolves to `Access_duration_1 (Get Packages
    # request)`. The lowercase label is display text, so both are left alone -
    # rewriting either would empty the runs inside the hyperlink.
    renumber(tbl)
    print("1) Package: dropped %s | owner_email %s | %d rows"
          % (dropped or "none",
             "added" if added else "already present",
             len(data_rows(tbl))))

    # --- 2) Document(Package Details) -------------------------------------
    tbl = object_table(body, "Document(Package Details)")
    renamed = []
    for old, (new_name, new_desc) in DOC_DETAILS_RENAME.items():
        tr = named_row(tbl, old)
        if tr is None:
            continue
        cells = row_cells(tr)
        set_leaf_text(cells[1], new_name)
        set_leaf_text(cells[4], new_desc)
        renamed.append("%s->%s" % (old, new_name))
    id_row = named_row(tbl, "id")
    id_fixed = False
    if id_row is not None and el_text(row_cells(id_row)[2]).strip() != "Integer":
        set_leaf_text(row_cells(id_row)[2], "Integer")
        id_fixed = True
    print("2) Document(Package Details): %s | id type %s"
          % (", ".join(renamed) or "already renamed",
             "fixed" if id_fixed else "already Integer"))

    # The endpoint Type cells are left alone - see the note next to
    # ENDPOINT_TYPE_FIXES above. Verify their hyperlinks still resolve so a
    # future edit that breaks one is caught here.
    marks = {bm.get(qn("name")) for bm in tree.getroot().iter(qn("bookmarkStart"))}
    checked = broken = 0
    for heading in ("Get All Packages", "Get Package Details"):
        tbl = endpoint_table(body, heading, "Response Attributes")
        for tr in data_rows(tbl):
            for h in row_cells(tr)[2].iter(qn("hyperlink")):
                checked += 1
                anchor = h.get(qn("anchor"))
                if not anchor or anchor not in marks or not el_text(h).strip():
                    broken += 1
                    print("   BROKEN LINK in %s: label=%r anchor=%r"
                          % (heading, el_text(h), anchor))
    print("3) object hyperlinks checked: %d, broken: %d" % (checked, broken))
    if broken:
        sys.exit("refusing to write with broken object links")

    if dry:
        print("\n--dry-run: nothing written")
        return

    if os.path.exists(OUT_DOCX) and not os.path.exists(BACKUP_DOCX):
        shutil.copy2(OUT_DOCX, BACKUP_DOCX)
        print("\nbacked up %s" % BACKUP_DOCX)

    try:
        guard_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("%s\n\nNothing was written. Re-run to re-apply on top of the "
                 "current content." % ex)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
