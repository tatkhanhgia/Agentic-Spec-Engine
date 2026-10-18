#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convert every error response in the GoPaperless API specification DOCX to the
current standard defined in scripts/GoPaperless_DOCX_Format_Template.md 2.8.8.

Old shape (legacy):                     New shape:
    {                                       {
      "error": "...",                         "error_description": "...",
      "error_description": "...",             "error_code": "38",
      "remark": [                             "message": "..."
        {"language_name": "en", ...},       }
        {"language_name": "vi", ...}
      ]
    }

Two things are rewritten:
  1. Sample error response frames (textboxes).
  2. The three error rows of every "Response Attributes" table
     (error / error_description / remark  ->  error_description / error_code / message).

Usage:
    python scripts/oneoff/apply_new_error_format.py [--dry-run]
"""
import os
import re
import sys
import shutil
import zipfile
from copy import deepcopy

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx"
)

# Shared error code for every documented error (user decision, 2026-07-26).
ERROR_CODE = "38"

# error_description (technical, unchanged from the existing frames) -> end-user message.
# Per 2.8.8 the two must not be identical; the message is the plain-English wording
# that matches the Vietnamese user-facing text the legacy remark array carried.
MESSAGES = {
    "The token is not belongs to GoPaperless Workflow Service!":
        "The token does not belong to this system. Please check again.",
    "Your Refresh token is invalid.":
        "Your refresh token is invalid. Please sign in again.",
    "Your client ID is invalid. Please request another client ID":
        "The application credentials are not valid. Please request a new client ID.",
}

# Response Attributes rows, in the order mandated by 2.8.8.
TABLE_ROWS = [
    ("error_description", "String", "M/O", "Detailed description of the error code."),
    ("error_code", "String", "M/O", "Code identifying the error."),
    ("message", "String", "M/O", "Error message displayed to the end user."),
]
LEGACY_ROW_NAMES = ("error", "error_description", "remark")


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def set_para_text(p, text):
    """Force a paragraph down to a single run carrying exactly `text`.

    Text can live in nested elements (`w:hyperlink`, `w:smartTag`, ...) — the
    `List<remarkError>` type cells wrap `remarkError` in a cross-reference
    hyperlink — so every content child is dropped, not just direct `w:r`
    children, otherwise the nested text survives and is appended to the new one.
    """
    pPr = p.find(qn("pPr"))

    template = None
    for r in p.iter(qn("r")):
        if r.find(qn("t")) is not None and r.getparent() is p:
            template = r          # prefer a plain run: nested ones carry link styling
            break
    if template is None:
        template = next((r for r in p.iter(qn("r")) if r.find(qn("t")) is not None), None)
    if template is None:
        raise ValueError("paragraph has no text run")

    new_run = deepcopy(template)
    for child in list(new_run):
        if child.tag not in (qn("rPr"), qn("t")):
            new_run.remove(child)
    for extra in new_run.findall(qn("t"))[1:]:
        new_run.remove(extra)
    t = new_run.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")

    for child in list(p):
        if child is not pPr:
            p.remove(child)
    p.append(new_run)


# --------------------------------------------------------------------------
# 1. Error response frames
# --------------------------------------------------------------------------
def is_legacy_error_frame(lines):
    joined = "\n".join(lines)
    return '"error_description"' in joined and '"remark"' in joined


def rebuild_error_frame(txbx):
    """Rewrite one legacy error frame in place. Returns the error_description."""
    paras = txbx.findall(qn("p"))
    lines = [text_of(p) for p in paras]

    i_err = next(i for i, l in enumerate(lines) if '"error":' in l)
    i_desc = next(i for i, l in enumerate(lines) if '"error_description":' in l)
    i_remark = next(i for i, l in enumerate(lines) if '"remark":' in l)
    i_close = next(i for i in range(len(lines) - 1, -1, -1) if lines[i].strip() == "}")

    m = re.search(r'"error_description":\s*"(.*?)",?\s*$', lines[i_desc])
    if m is None:
        raise ValueError("cannot parse error_description: %r" % lines[i_desc])
    description = m.group(1).strip()
    message = MESSAGES.get(description)
    if message is None:
        raise KeyError("no end-user message defined for %r" % description)

    # Reuse the three surviving paragraphs so run formatting is preserved, then
    # drop everything between them and the closing brace.
    set_para_text(paras[i_err], '  "error_description": "%s",' % description)
    set_para_text(paras[i_desc], '  "error_code": "%s",' % ERROR_CODE)
    set_para_text(paras[i_remark], '  "message": "%s"' % message)
    for p in paras[i_remark + 1:i_close]:
        txbx.remove(p)
    return description


def rebuild_all_frames(root):
    done = []
    for txbx in root.iter(qn("txbxContent")):
        lines = [text_of(p) for p in txbx.findall(qn("p"))]
        if is_legacy_error_frame(lines):
            done.append(rebuild_error_frame(txbx))
    return done


# --------------------------------------------------------------------------
# 2. Response Attributes tables
# --------------------------------------------------------------------------
def row_cells(tr):
    return tr.findall(qn("tc"))


def cell_text(tc):
    return " ".join(
        x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x
    ).strip()


def set_cell_text(tc, text):
    paras = tc.findall(qn("p"))
    if not paras:
        raise ValueError("empty cell")
    set_para_text(paras[0], text)
    for extra in paras[1:]:
        tc.remove(extra)


def is_response_attributes_table(rows_text):
    return bool(rows_text) and "Response Attributes" in " ".join(rows_text[0])


def rebuild_table(tbl):
    """Rewrite the legacy error rows of one Response Attributes table.

    Returns True when the table was changed.
    """
    trs = tbl.findall(qn("tr"))
    grid = [[cell_text(tc) for tc in row_cells(tr)] for tr in trs]
    if not is_response_attributes_table(grid):
        return False

    idx = [
        i for i, row in enumerate(grid)
        if len(row) > 1 and row[1] in LEGACY_ROW_NAMES
    ]
    # Only touch the contiguous legacy triple error / error_description / remark.
    if len(idx) != 3 or idx[2] - idx[0] != 2:
        return False
    if [grid[i][1] for i in idx] != list(LEGACY_ROW_NAMES):
        return False

    for i, (name, typ, presence, desc) in zip(idx, TABLE_ROWS):
        cells = row_cells(trs[i])
        if len(cells) < 5:
            raise ValueError("unexpected column count: %d" % len(cells))
        set_cell_text(cells[1], name)
        set_cell_text(cells[2], typ)
        set_cell_text(cells[3], presence)
        set_cell_text(cells[4], desc)
    return True


def rebuild_all_tables(root):
    n = 0
    for tbl in root.iter(qn("tbl")):
        if rebuild_table(tbl):
            n += 1
    return n


# --------------------------------------------------------------------------
# 3. The now-orphaned RemarkError object definition
# --------------------------------------------------------------------------
def remove_remark_error_object(root):
    """Drop the `RemarkError` entry from the RELATIVE OBJECT section.

    Nothing references the type once `List<remarkError>` is gone, so the
    Heading 2 plus its definition table are dead weight.
    """
    body = root.find(qn("body"))
    kids = list(body)
    for i, el in enumerate(kids):
        if el.tag != qn("p"):
            continue
        pPr = el.find(qn("pPr"))
        pStyle = pPr.find(qn("pStyle")) if pPr is not None else None
        if pStyle is None or pStyle.get(qn("val")) != "Heading2":
            continue
        if text_of(el).strip() != "RemarkError":
            continue
        nxt = kids[i + 1] if i + 1 < len(kids) else None
        if nxt is None or nxt.tag != qn("tbl"):
            raise ValueError("RemarkError heading is not followed by its table")
        if "RemarkError" not in text_of(nxt):
            raise ValueError("table after RemarkError heading looks wrong")

        # A bookmarkStart inside the removed content can leave its bookmarkEnd
        # stranded elsewhere in the body; Word treats an unmatched bookmarkEnd
        # as a corrupt document and offers to repair it on open.
        dropped_ids = {
            b.get(qn("id"))
            for el2 in (el, nxt)
            for b in el2.iter(qn("bookmarkStart"))
        }
        body.remove(nxt)
        body.remove(el)
        for end in list(body.iter(qn("bookmarkEnd"))):
            if end.get(qn("id")) in dropped_ids:
                end.getparent().remove(end)
        return True
    return False


# --------------------------------------------------------------------------
# 4. Duplicate `message` row (Add Identity for a User)
# --------------------------------------------------------------------------
MERGED_MESSAGE_DESC = (
    "Status of the processing on success, or the error message shown to the "
    "end user on failure."
)


def merge_duplicate_message_row(root):
    """Collapse the two `message` rows of the one table that has both.

    `Add Identity for a User` returns `"message": "PROCESSING"` on success and
    the standard error `message` on failure, so its Response Attributes table
    ended up listing the same field name twice.
    """
    for tbl in root.iter(qn("tbl")):
        trs = tbl.findall(qn("tr"))
        grid = [[cell_text(tc) for tc in row_cells(tr)] for tr in trs]
        if not is_response_attributes_table(grid):
            continue
        msg_idx = [i for i, row in enumerate(grid) if len(row) > 1 and row[1] == "message"]
        if len(msg_idx) < 2:
            continue
        keep, drop = msg_idx[0], msg_idx[1:]
        set_cell_text(row_cells(trs[keep])[4], MERGED_MESSAGE_DESC)
        for i in reversed(drop):
            tbl.remove(trs[i])

        # Renumber the No column so it stays sequential after the removal.
        n = 1
        for tr in tbl.findall(qn("tr"))[2:]:
            cells = row_cells(tr)
            if cells:
                set_cell_text(cells[0], str(n))
                n += 1
        return True
    return False


# --------------------------------------------------------------------------
# 5. Make Word refresh the table of contents on open
# --------------------------------------------------------------------------
def enable_update_fields(settings_path):
    """Add <w:updateFields w:val="true"/> so the stale TOC refreshes on open.

    Removing the RemarkError heading left the cached TOC field out of date;
    Word only rebuilds it when the field is updated.
    """
    tree = etree.parse(settings_path)
    root = tree.getroot()
    if root.find(qn("updateFields")) is not None:
        return False, tree

    el = etree.Element(qn("updateFields"))
    el.set(qn("val"), "true")
    # CT_Settings is a sequence: updateFields sits just before hdrShapeDefaults.
    anchor = root.find(qn("hdrShapeDefaults"))
    if anchor is not None:
        anchor.addprevious(el)
    else:
        root.append(el)
    return True, tree


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

    frames = rebuild_all_frames(root)
    tables = rebuild_all_tables(root)
    removed = remove_remark_error_object(root)
    merged = merge_duplicate_message_row(root)

    settings_path = os.path.join(UNPACKED, "word", "settings.xml")
    fields_added, settings_tree = enable_update_fields(settings_path)

    print("error frames rebuilt        : %d" % len(frames))
    for desc in sorted(set(frames)):
        print("      * %-60s -> %s" % (desc[:60], MESSAGES[desc]))
    print("Response Attributes tables  : %d" % tables)
    print("RemarkError object removed  : %s" % removed)
    print("duplicate message row merged: %s" % merged)
    print("updateFields flag added     : %s" % fields_added)

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    if fields_added:
        settings_tree.write(
            settings_path, xml_declaration=True, encoding="UTF-8", standalone=True
        )
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
