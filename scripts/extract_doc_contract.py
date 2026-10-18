#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extract the API contract from the formatted GoPaperless specification DOCX.

Reads the same structure the formatting rules define in
`scripts/GoPaperless_DOCX_Format_Template.md`:

  * Heading 2 = API category, Heading 3 = endpoint, Heading 4 = sub-feature
  * a Sample Request frame whose first line is `METHOD {baseURL}/path`
  * the four attribute tables (Header Request / Path / Request / Response)

and emits one JSON record per documented endpoint so it can be compared with
`from_code.json` produced by `extract_code_contract.py`.

THIS SCRIPT ONLY READS the DOCX. It does not modify the document, the unpacked
working copy, or any formatting rule.

Usage:
    python scripts/extract_doc_contract.py [path-to-docx] [-o out.json]
"""
import json
import os
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX_DEFAULT = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
OUT_DEFAULT = os.path.join(ROOT, "contract", "from_doc.json")

METHOD_RE = re.compile(r"\b(GET|POST|PUT|DELETE|PATCH)\s*(\{baseURL\}\S*)")

TABLE_KINDS = {
    "header request attributes": "headers",
    "path attributes": "path_params",
    "request attributes": "request_fields",
    "response attributes": "response_fields",
    "header response attributes": "response_headers",
}


def qn(tag):
    return "{%s}%s" % (W, tag)


def para_text(p):
    """Paragraph text; nested paragraphs (textbox lines) become newlines."""
    out = []
    for n in p.iter():
        tag = etree.QName(n).localname
        if tag == "p" and n is not p:
            out.append("\n")
        elif tag == "t":
            out.append(n.text or "")
        elif tag == "br":
            out.append("\n")
        elif tag == "tab":
            out.append("\t")
    return "".join(out)


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def cell_text(tc):
    return " ".join(
        x for x in (para_text(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def to_key(path):
    """Same normalisation as extract_code_contract.to_key."""
    p = path
    p = re.sub(r"^\{baseURL\}", "", p)
    p = p.split("?")[0]
    m = re.search(r"/v\d+/", p)
    if m:
        p = p[m.start():]
    p = re.sub(r"\{[^}]*\}", "{p}", p)
    segs = ["{p}" if s.isdigit() else s for s in p.split("/") if s != ""]
    return ("/" + "/".join(segs)).rstrip("/").lower()


def parse_table(tbl):
    """Return (kind, rows) for an attribute table, or (None, None)."""
    trs = tbl.findall(qn("tr"))
    if len(trs) < 2:
        return None, None
    title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()
    kind = TABLE_KINDS.get(title.lower())
    if kind is None:
        return None, None

    rows = []
    for tr in trs[2:]:
        cells = [cell_text(tc) for tc in tr.findall(qn("tc"))]
        if len(cells) < 2 or not cells[1]:
            continue
        rows.append({
            "name": cells[1],
            "type": cells[2] if len(cells) > 2 else "",
            "presence": cells[3] if len(cells) > 3 else "",
            "description": cells[4] if len(cells) > 4 else "",
        })
    return kind, rows


def extract(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(qn("body"))

    heading = {2: "", 3: "", 4: ""}
    endpoints = []
    current = None

    def flush():
        if current and current.get("method"):
            endpoints.append(current)

    for el in body:
        tag = etree.QName(el).localname

        if tag == "p":
            st = style_of(el)
            txt = para_text(el)
            m = re.match(r"Heading(\d)$", st)
            if m and int(m.group(1)) in (2, 3, 4):
                lv = int(m.group(1))
                flush()
                current = None
                heading[lv] = txt.strip()
                for k in range(lv + 1, 5):
                    heading[k] = ""
                if lv in (3, 4):
                    current = {
                        "name": heading[lv],
                        "category": heading[2],
                        "parent": heading[3] if lv == 4 else "",
                        "level": lv,
                        "method": "",
                        "path": "",
                        "key": "",
                        "headers": [],
                        "path_params": [],
                        "request_fields": [],
                        "response_fields": [],
                        "response_headers": [],
                    }
                continue

            if current is not None and not current["method"]:
                mm = METHOD_RE.search(txt)
                if mm:
                    current["method"] = mm.group(1).upper()
                    current["path"] = mm.group(2)
                    current["key"] = "%s %s" % (mm.group(1).upper(), to_key(mm.group(2)))

        elif tag == "tbl" and current is not None:
            kind, rows = parse_table(el)
            if kind and not current[kind]:
                current[kind] = rows

    flush()
    return endpoints


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    docx_path = args[0] if args else DOCX_DEFAULT
    out_path = OUT_DEFAULT
    if "-o" in sys.argv:
        out_path = sys.argv[sys.argv.index("-o") + 1]

    if not os.path.isfile(docx_path):
        sys.exit("docx not found: %s" % docx_path)

    endpoints = extract(docx_path)
    data = {
        "source": "doc",
        "docx": os.path.basename(docx_path),
        "count": len(endpoints),
        "endpoints": endpoints,
    }
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

    print("documented endpoints : %d" % len(endpoints))
    print("  with a header table: %d" % sum(1 for e in endpoints if e["headers"]))
    print("  with request fields: %d" % sum(1 for e in endpoints if e["request_fields"]))
    print("  with response table: %d" % sum(1 for e in endpoints if e["response_fields"]))
    print("wrote %s" % out_path)


if __name__ == "__main__":
    main()
