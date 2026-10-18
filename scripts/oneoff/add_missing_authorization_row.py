#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add the missing `Authorization` row to Header Request Attributes tables.

Two endpoints send `Authorization: "Bearer {accesstoken}"` in their Sample
Request frame but never document the header in the table beside it:

  * Pre Login Authentication   (Authenticate - Log in)
  * Delete Address Book        (Personal Settings)

The row is cloned from a table that already documents it, so wording, borders
and fonts stay identical to the other 183 tables. It is inserted directly after
`x-language-name`, matching the order used by the majority of the document
(x-language-name, Authorization, Content-Type), and the No column is renumbered.

Detection is behavioural, not a hardcoded list: a table is repaired only when
the endpoint's own request frame actually carries the header.

Usage:
    python scripts/oneoff/add_missing_authorization_row.py [--dry-run]
"""
import os
import re
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

HEADER_TABLE = "Header Request Attributes"
METHOD_RE = re.compile(r"^(GET|POST|PUT|DELETE|PATCH) ")


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


def set_para_text(p, text):
    pPr = p.find(qn("pPr"))
    template = None
    for r in p.iter(qn("r")):
        if r.find(qn("t")) is not None:
            template = r
            break
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


def set_cell_text(tc, text):
    paras = tc.findall(qn("p"))
    set_para_text(paras[0], text)
    for extra in paras[1:]:
        tc.remove(extra)


def is_header_table(tbl):
    trs = tbl.findall(qn("tr"))
    if len(trs) < 2:
        return False
    title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()
    return HEADER_TABLE in title


def row_names(tbl):
    out = []
    for tr in tbl.findall(qn("tr"))[2:]:
        tcs = tr.findall(qn("tc"))
        out.append(cell_text(tcs[1]) if len(tcs) > 1 else "")
    return out


def find_donor_row(body):
    """An existing Authorization row to clone, taken from any header table."""
    for tbl in body.iter(qn("tbl")):
        if not is_header_table(tbl):
            continue
        for tr in tbl.findall(qn("tr"))[2:]:
            tcs = tr.findall(qn("tc"))
            if len(tcs) >= 5 and cell_text(tcs[1]) == "Authorization":
                return deepcopy(tr)
    raise LookupError("no Authorization row found to clone")


def renumber(tbl):
    n = 1
    for tr in tbl.findall(qn("tr"))[2:]:
        tcs = tr.findall(qn("tc"))
        if tcs:
            set_cell_text(tcs[0], str(n))
            n += 1


def collect_targets(body):
    """Endpoints whose frame sends Authorization but whose table omits it."""
    targets = []
    heading2 = ""
    current = None
    for el in body:
        if el.tag == qn("p"):
            st = style_of(el)
            if st == "Heading2":
                heading2 = text_of(el).strip()
                current = None
            elif st in ("Heading3", "Heading4"):
                current = {"name": text_of(el).strip(), "category": heading2,
                           "frame_auth": False, "table": None}
                targets.append(current)
            if current is not None:
                for txbx in el.iter(qn("txbxContent")):
                    lines = [text_of(p) for p in txbx.findall(qn("p"))]
                    if lines and METHOD_RE.match(lines[0]):
                        if any(l.strip().lower().startswith("authorization") for l in lines):
                            current["frame_auth"] = True
        elif el.tag == qn("tbl") and current is not None:
            if is_header_table(el) and current["table"] is None:
                current["table"] = el
    return [t for t in targets
            if t["frame_auth"] and t["table"] is not None
            and "Authorization" not in row_names(t["table"])]


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

    donor = find_donor_row(body)
    targets = collect_targets(body)
    print("tables missing an Authorization row: %d" % len(targets))
    if not targets:
        sys.exit("nothing to do")

    for t in targets:
        tbl = t["table"]
        names = row_names(tbl)
        rows = tbl.findall(qn("tr"))
        # insert after x-language-name when present, otherwise after the header row
        if "x-language-name" in names:
            anchor = rows[2 + names.index("x-language-name")]
        else:
            anchor = rows[1]
        anchor.addnext(deepcopy(donor))
        renumber(tbl)
        print("   + %-38s | %-22s -> %s"
              % (t["name"][:38], t["category"][:22], ", ".join(row_names(tbl))))

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
