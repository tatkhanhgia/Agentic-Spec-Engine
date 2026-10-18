#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add `Remove Account Request` (POST /v1/account/remove) to the Personal Settings
section of the specification.

Placement: SigningHub 10.0 lists "Remove Account Request" as the FIRST entry of
its Personal Settings chapter (NotebookLM query 2026-07-27), and the owner
decided "SigningHub ho dat o dau thi minh dat o do". So the new section is
inserted directly after the `Personal Settings` Heading 2, before
`Get Signature Settings`.

Donor: `Cancel Issue Certificate` (Register Certificate Processing). Picked from
contract/from_doc.json because it already matches the target shape exactly -
POST, the three standard headers (x-language-name O / Authorization M /
Content-Type M), one mandatory request-body field, an empty `{}` 200 frame, the
standard 400 error frame, and no Path Attributes table. The edit reduces to
text substitution only.

Facts documented here come from the backend, not from the model:
  * route `POST /v1/account/remove` - `UserController` validURLs entry 8,
    dispatched in `doPost` ("Remove Account Request" apiName).
  * body field `password` mandatory - `ValidateRequest.validateDeleteUser`
    rejects a missing password (SUBCODE_USER_MISSING_PASSWORD_OF_USER).
    The Postman collection request sends exactly {"password": "..."}.
  * `DeleteUserDTO.userEmail` is deliberately NOT documented: it is optional
    (defaults to the token's user) and its JSON key is camelCase, which
    template 2.8.9 forbids - recorded as a backend defect instead.
  * empty success body - `UserService.deleteUser` returns
    `new InternalResponse<Void>()`, the status-only pattern.

The clone keeps the donor's method badge run (POST, dark shading) and the red
`{baseURL}` placeholder run untouched; only the path portion after `{baseURL}`
is rewritten, so no formatter re-run is needed.

Identifier hygiene on cloning: bookmarks stripped, `wp:docPr/@id` and VML
`v:shape/@id`+`@o:spid` reallocated above the current maxima.

Usage:
    python scripts/oneoff/add_remove_account_request.py [--dry-run]
"""
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy

from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import docx_guard

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

DONOR_HEADING = "Cancel Issue Certificate"
SECTION_H2 = "Personal Settings"            # new section goes right after this H2

NEW_HEADING = "Remove Account Request"
NEW_PROSE = (
    "Business applications can use this service API to request the deletion of "
    "the account of the current user. The password of the account must be "
    "provided to confirm the removal request."
)
NEW_PATH_TAIL = "/v1/account/remove"        # replaces everything after {baseURL}
OLD_BODY_KEY = '"session_id"'
NEW_BODY_LINE = '  "password": ""'

NEW_REQUEST_ROW = (
    "password", "String", "M",
    "The password of the current user account. It is used to confirm the "
    "account removal request.")


# --------------------------------------------------------------------------
# small helpers (same shapes as add_switch_enterprise_context.py)
# --------------------------------------------------------------------------
def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def set_para_text(p, text):
    """Reduce a paragraph to one run carrying `text`, keeping its formatting."""
    pPr = p.find(qn("pPr"))
    template = None
    for r in p.iter(qn("r")):
        if r.find(qn("t")) is not None:
            template = r
            break
    if template is None:
        raise ValueError("paragraph has no text run: %r" % text_of(p))

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


def set_path_tail(p, tail):
    """Rewrite only the path portion after the {baseURL} run.

    Keeps the method badge run and the red {baseURL} placeholder run exactly
    as the donor had them, so the frame needs no formatter re-run.
    """
    seen_base = False
    replaced = False
    for r in p.findall(qn("r")):
        t = r.find(qn("t"))
        if t is None:
            continue
        if not seen_base:
            if (t.text or "") == "{baseURL}":
                seen_base = True
            continue
        if not replaced:
            t.text = tail
            t.set(XML_SPACE, "preserve")
            replaced = True
        else:
            t.text = ""
    if not (seen_base and replaced):
        raise RuntimeError("path line had no {baseURL} run: %r" % text_of(p))


def cell_text(tc):
    return " ".join(x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def set_cell_text(tc, text):
    paras = tc.findall(qn("p"))
    set_para_text(paras[0], text)
    for extra in paras[1:]:
        tc.remove(extra)


def table_title(tbl):
    trs = tbl.findall(qn("tr"))
    if not trs:
        return ""
    return " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()


def find_heading(kids, level, title):
    want = "Heading%d" % level
    for i, e in enumerate(kids):
        if e.tag == qn("p") and style_of(e) == want and text_of(e).strip() == title:
            return i
    raise LookupError("heading not found: %s" % title)


def section_range(kids, start):
    for j in range(start + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and re.match(r"Heading\d$", style_of(e)):
            return start, j
    return start, len(kids)


# --------------------------------------------------------------------------
# identifier hygiene
# --------------------------------------------------------------------------
def strip_bookmarks(elements):
    n = 0
    for el in elements:
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for b in list(el.iter(qn(tag))):
                b.getparent().remove(b)
                n += 1
    return n


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def reassign_ids(elements, next_docpr, next_spid):
    n_doc = n_vml = 0
    for el in elements:
        for d in el.iter():
            if local(d) == "docPr":
                d.set("id", str(next_docpr))
                d.set("name", "Text Box %d" % next_docpr)
                next_docpr += 1
                n_doc += 1
        for d in el.iter():
            if local(d) == "shape":
                d.set("id", "_x0000_s%d" % next_spid)
                if d.get("{%s}spid" % O) is not None:
                    d.set("{%s}spid" % O, "_x0000_s%d" % next_spid)
                next_spid += 1
                n_vml += 1
    return n_doc, n_vml


def scan_id_space(root):
    docpr = [int(d.get("id")) for d in root.iter()
             if local(d) == "docPr" and (d.get("id") or "").isdigit()]
    spids = []
    for d in root.iter():
        if local(d) in ("shape", "shapetype"):
            for v in (d.get("id"), d.get("{%s}spid" % O)):
                m = re.match(r"_x0000_s(\d+)$", v or "")
                if m:
                    spids.append(int(m.group(1)))
    return (max(docpr) + 1 if docpr else 1), (max(spids) + 1 if spids else 1000)


# --------------------------------------------------------------------------
# the edit
# --------------------------------------------------------------------------
def edit_clone(clone):
    set_para_text(clone[0], NEW_HEADING)
    set_para_text(clone[1], NEW_PROSE)

    paths_done = bodies_done = 0
    for el in clone:
        for txbx in el.iter(qn("txbxContent")):
            paras = txbx.findall(qn("p"))
            lines = [text_of(p) for p in paras]
            if not lines or not lines[0].startswith("POST "):
                continue
            set_path_tail(paras[0], NEW_PATH_TAIL)
            paths_done += 1
            for p, line in zip(paras, lines):
                if OLD_BODY_KEY in line:
                    set_para_text(p, NEW_BODY_LINE)
                    bodies_done += 1
    if paths_done == 0 or bodies_done == 0:
        raise RuntimeError("request frame rewrite incomplete: paths=%d bodies=%d"
                           % (paths_done, bodies_done))

    rows_done = 0
    for el in clone:
        if el.tag == qn("tbl") and table_title(el) == "Request Attributes":
            trs = el.findall(qn("tr"))[2:]
            if len(trs) != 1:
                raise RuntimeError("donor Request Attributes has %d data rows, need 1"
                                   % len(trs))
            tcs = trs[0].findall(qn("tc"))
            name, typ, presence, desc = NEW_REQUEST_ROW
            set_cell_text(tcs[1], name)
            set_cell_text(tcs[2], typ)
            set_cell_text(tcs[3], presence)
            set_cell_text(tcs[4], desc)
            rows_done += 1
    if rows_done != 1:
        raise RuntimeError("Request Attributes table not rewritten")
    return paths_done, bodies_done


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
    tree, token = docx_guard.load(DOC_XML)
    root = tree.getroot()
    body = root.find(qn("body"))
    kids = list(body)

    if any(e.tag == qn("p") and style_of(e) == "Heading3"
           and text_of(e).strip() == NEW_HEADING for e in kids):
        sys.exit("'%s' already exists - nothing to do" % NEW_HEADING)

    d_start, d_end = section_range(kids, find_heading(kids, 3, DONOR_HEADING))
    h2 = find_heading(kids, 2, SECTION_H2)
    print("donor  : [%d..%d)  %s" % (d_start, d_end, DONOR_HEADING))
    print("anchor : %d  (H2 %s)" % (h2, SECTION_H2))

    clone = [deepcopy(e) for e in kids[d_start:d_end]]
    n_bm = strip_bookmarks(clone)
    next_docpr, next_spid = scan_id_space(root)
    n_doc, n_vml = reassign_ids(clone, next_docpr, next_spid)
    n_paths, n_bodies = edit_clone(clone)

    print("cloned elements          : %d" % len(clone))
    print("bookmarks stripped       : %d" % n_bm)
    print("docPr re-ided (from %d)  : %d" % (next_docpr, n_doc))
    print("VML shapes re-ided (s%d) : %d" % (next_spid, n_vml))
    print("path lines rewritten     : %d" % n_paths)
    print("body lines rewritten     : %d" % n_bodies)

    anchor = kids[h2]
    for el in clone:
        anchor.addnext(el)
        anchor = el

    if dry:
        print("\n--dry-run: nothing written")
        return
    docx_guard.save(tree, DOC_XML, token)
    pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
