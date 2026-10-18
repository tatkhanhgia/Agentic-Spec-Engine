#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add `Switch Enterprise Context` (POST /v1/account/switch) to the Authentication
section of the specification.

The section is built by cloning an existing endpoint section rather than
assembling WordprocessingML by hand, so every style, table border, badge colour
and frame shape is inherited already-correct.

Donor: `Send OTP to Certificate Registration`. It was picked from
contract/from_doc.json because it already matches the target shape - POST, the
exact three request headers (x-language-name / Authorization / Content-Type),
two request-body fields, an empty `{}` success frame and the standard error
frame - which reduces the edit to text substitution plus dropping the one table
the new endpoint does not need (Path Attributes).

Cloning duplicates three kinds of identifier that must stay unique, or Word
reports the file as corrupt and offers to repair it:
  * `w:bookmarkStart` / `w:bookmarkEnd` - stripped from the copy entirely; both
    ends of every bookmark in the donor range live inside it, so nothing is left
    dangling.
  * `wp:docPr/@id` - reallocated above the current maximum.
  * VML `v:shape/@id` and `@o:spid` - a separate id space from docPr, so it is
    scanned and allocated separately.

Facts documented here come from the backend, not from the model:
  * route `POST /v1/account/switch` - `UserController` `validURLs` entry 15.
  * body fields `profile_id` / `enterprise_id` and the priority rule - the
    javadoc of `SwitchEnterpriseRequest`.
  * both fields optional - `ValidateRequest.validateSwitchEnterpriseRequest`
    only checks that the payload parses.
  * empty success body - `UserService.switchEnterprise` returns
    `new InternalResponse().setStatus(HTTP_CODE_SUCCESS)` with no data, the same
    status-only pattern used in 103 other places in that codebase.

Usage:
    python scripts/oneoff/add_switch_enterprise_context.py [--dry-run]
"""
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy

from lxml import etree

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

DONOR_HEADING = "Send OTP to Certificate Registration"
ANCHOR_HEADING = "Pre Login Authentication"     # new section goes after this one

NEW_HEADING = "Switch Enterprise Context"
NEW_PROSE = (
    "Business applications can use this service API to switch the working context of "
    "the current user between a personal account and an enterprise account that the "
    "user belongs to. All subsequent API calls are then performed in the selected context."
)
NEW_PATH_LINE = "POST {baseURL}/v1/account/switch"
NEW_BODY_LINES = ["{", '  "profile_id": 1', "}"]

NEW_REQUEST_ROWS = [
    ("profile_id", "Integer", "O",
     "The ID of the enterprise profile to switch to. This parameter takes priority "
     "over enterprise_id. A positive value resolves the enterprise from the profile "
     "list of the user. A value of 0 or less switches the context back to the "
     "personal account."),
    ("enterprise_id", "Integer", "O",
     "The ID of the enterprise to switch to. It is an alternative to profile_id and "
     "is used only when profile_id is not provided. A value of 0 or -1 switches the "
     "context back to the personal account."),
]


# --------------------------------------------------------------------------
# small helpers
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
    """From a Heading3 index to just before the next heading of any level."""
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
    """Give cloned drawings fresh identifiers in both id spaces."""
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
    """Turn the donor copy into the new section."""
    # heading + prose
    set_para_text(clone[0], NEW_HEADING)
    set_para_text(clone[1], NEW_PROSE)

    # request frame: path line, then rebuild the JSON body
    frames_done = 0
    for el in clone:
        for txbx in el.iter(qn("txbxContent")):
            paras = txbx.findall(qn("p"))
            lines = [text_of(p) for p in paras]
            if not lines or not lines[0].startswith("POST "):
                continue
            set_para_text(paras[0], NEW_PATH_LINE)
            body_start = next(i for i, l in enumerate(lines) if l.strip() == "{")
            keep = paras[:body_start]
            for p in paras[body_start:]:
                txbx.remove(p)
            anchor = keep[-1]
            src = paras[body_start]
            for line in NEW_BODY_LINES:
                new_p = deepcopy(src)
                set_para_text(new_p, line)
                anchor.addnext(new_p)
                anchor = new_p
            frames_done += 1
    if frames_done == 0:
        raise RuntimeError("request frame not found in the clone")

    # drop the Path Attributes table (and the blank paragraph after it)
    dropped = 0
    for i, el in enumerate(list(clone)):
        if el.tag == qn("tbl") and table_title(el) == "Path Attributes":
            nxt = clone[i + 1] if i + 1 < len(clone) else None
            clone.remove(el)
            if nxt is not None and nxt.tag == qn("p") and not text_of(nxt).strip():
                clone.remove(nxt)
            dropped += 1
            break

    # request attributes rows
    for el in clone:
        if el.tag == qn("tbl") and table_title(el) == "Request Attributes":
            trs = el.findall(qn("tr"))[2:]
            if len(trs) != len(NEW_REQUEST_ROWS):
                raise RuntimeError("donor Request Attributes has %d data rows, need %d"
                                   % (len(trs), len(NEW_REQUEST_ROWS)))
            for tr, (name, typ, presence, desc) in zip(trs, NEW_REQUEST_ROWS):
                tcs = tr.findall(qn("tc"))
                set_cell_text(tcs[1], name)
                set_cell_text(tcs[2], typ)
                set_cell_text(tcs[3], presence)
                set_cell_text(tcs[4], desc)
            break
    return frames_done, dropped


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

    if any(e.tag == qn("p") and style_of(e) == "Heading3"
           and text_of(e).strip() == NEW_HEADING for e in kids):
        sys.exit("'%s' already exists - nothing to do" % NEW_HEADING)

    d_start, d_end = section_range(kids, find_heading(kids, 3, DONOR_HEADING))
    a_start, a_end = section_range(kids, find_heading(kids, 3, ANCHOR_HEADING))
    print("donor  : [%d..%d)  %s" % (d_start, d_end, DONOR_HEADING))
    print("anchor : [%d..%d)  %s" % (a_start, a_end, ANCHOR_HEADING))

    clone = [deepcopy(e) for e in kids[d_start:d_end]]
    n_bm = strip_bookmarks(clone)
    next_docpr, next_spid = scan_id_space(root)
    n_doc, n_vml = reassign_ids(clone, next_docpr, next_spid)
    n_frames, n_dropped = edit_clone(clone)

    print("cloned elements          : %d" % len(clone))
    print("bookmarks stripped       : %d" % n_bm)
    print("docPr re-ided (from %d) : %d" % (next_docpr, n_doc))
    print("VML shapes re-ided (s%d): %d" % (next_spid, n_vml))
    print("request frames rewritten : %d" % n_frames)
    print("Path Attributes dropped  : %d" % n_dropped)

    anchor = kids[a_end - 1]
    for el in clone:
        anchor.addnext(el)
        anchor = el

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
