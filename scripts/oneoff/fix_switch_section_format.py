#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Repair the `Switch Enterprise Context` section added earlier.

Three defects, all introduced when the section was cloned:

1. **Badge shading covered the whole first line.** The clone helper rebuilt the
   paragraph from its first run, and that run is the method badge, so the bold
   white-on-orange formatting spread across `{baseURL}/v1/account/switch` too.
   Template 2.8.1 now spells out the required run structure; the fix restores it
   by editing only the path run and leaving badge / space / `{baseURL}` alone.

2. **`No` and `Presence` not centered** in Header Request Attributes. The donor
   section is one of the 57 endpoints that predate the alignment rule, so the
   defect was inherited. Template 2.4 now states the rule explicitly.

3. **`profile_id` documented as optional.** It is mandatory.

Usage:
    python scripts/oneoff/fix_switch_section_format.py [--dry-run]
"""
import os
import shutil
import sys
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
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

SECTION = "Switch Enterprise Context"
METHOD = "POST"
PATH = "/v1/account/switch"
CENTER_COLUMNS = (0, 3)          # No and Presence


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def cell_text(tc):
    return " ".join(x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def set_run_text(run, text):
    for extra in run.findall(qn("t"))[1:]:
        run.remove(extra)
    t = run.find(qn("t"))
    if t is None:
        t = etree.SubElement(run, qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")


def section_tables(body, title):
    """Tables belonging to the named Heading3/4 section, in order."""
    tables = []
    inside = False
    for el in body:
        if el.tag == qn("p"):
            st = style_of(el)
            if st.startswith("Heading"):
                inside = st in ("Heading3", "Heading4") and text_of(el).strip() == title
        elif el.tag == qn("tbl") and inside:
            tables.append(el)
    return tables


def section_frames(body, title):
    frames = []
    inside = False
    for el in body:
        if el.tag == qn("p"):
            st = style_of(el)
            if st.startswith("Heading"):
                inside = st in ("Heading3", "Heading4") and text_of(el).strip() == title
                continue
            if inside:
                frames.extend(el.iter(qn("txbxContent")))
    return frames


# --------------------------------------------------------------------------
# 1. badge run structure
# --------------------------------------------------------------------------
def rebuild_badge_line(txbx, donor_para):
    """Replace a collapsed first line with the donor's run structure."""
    paras = txbx.findall(qn("p"))
    first = paras[0]
    new_p = deepcopy(donor_para)

    runs = [r for r in new_p.findall(qn("r")) if r.find(qn("t")) is not None]
    texts = [text_of(r) for r in runs]
    # runs are: METHOD, space, {baseURL}, path, then empties
    set_run_text(runs[0], METHOD)
    idx_base = next(i for i, t in enumerate(texts) if t.strip() == "{baseURL}")
    set_run_text(runs[idx_base + 1], PATH)
    for r in runs[idx_base + 2:]:
        set_run_text(r, "")

    first.addprevious(new_p)
    txbx.remove(first)
    return new_p


def find_donor_first_para(body):
    """A correctly formatted POST frame first line to copy the structure from."""
    for txbx in body.iter(qn("txbxContent")):
        paras = txbx.findall(qn("p"))
        if not paras:
            continue
        runs = [r for r in paras[0].findall(qn("r")) if r.find(qn("t")) is not None]
        if len(runs) < 4:
            continue
        texts = [text_of(r) for r in runs]
        if texts[0].strip() == METHOD and "{baseURL}" in texts:
            rPr = runs[0].find(qn("rPr"))
            if rPr is not None and rPr.find(qn("shd")) is not None:
                return deepcopy(paras[0])
    raise LookupError("no correctly structured %s frame line found" % METHOD)


# --------------------------------------------------------------------------
# 2 + 3. table fixes
# --------------------------------------------------------------------------
def center_cell(tc):
    changed = False
    for p in tc.findall(qn("p")):
        pPr = p.find(qn("pPr"))
        if pPr is None:
            pPr = etree.Element(qn("pPr"))
            p.insert(0, pPr)
        jc = pPr.find(qn("jc"))
        if jc is None:
            jc = etree.SubElement(pPr, qn("jc"))
        if jc.get(qn("val")) != "center":
            jc.set(qn("val"), "center")
            changed = True
    return changed


def set_cell(tc, text):
    paras = tc.findall(qn("p"))
    p = paras[0]
    runs = [r for r in p.iter(qn("r")) if r.find(qn("t")) is not None]
    set_run_text(runs[0], text)
    for r in runs[1:]:
        r.getparent().remove(r)
    for extra in paras[1:]:
        tc.remove(extra)


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

    donor_para = find_donor_first_para(body)

    n_badge = 0
    for txbx in section_frames(body, SECTION):
        paras = txbx.findall(qn("p"))
        if not paras or not text_of(paras[0]).startswith(METHOD + " "):
            continue
        runs = [r for r in paras[0].findall(qn("r")) if r.find(qn("t")) is not None]
        if len(runs) > 1:
            continue                              # already correct
        rebuild_badge_line(txbx, donor_para)
        n_badge += 1

    n_center = n_presence = 0
    for tbl in section_tables(body, SECTION):
        for tr in tbl.findall(qn("tr"))[2:]:
            tcs = tr.findall(qn("tc"))
            if len(tcs) < 5:
                continue
            for col in CENTER_COLUMNS:
                if center_cell(tcs[col]):
                    n_center += 1
            if cell_text(tcs[1]) == "profile_id" and cell_text(tcs[3]) != "M":
                set_cell(tcs[3], "M")
                n_presence += 1

    print("frame lines rebuilt (badge scoped to method) : %d" % n_badge)
    print("cells centered (No / Presence)               : %d" % n_center)
    print("profile_id presence O -> M                   : %d" % n_presence)

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
