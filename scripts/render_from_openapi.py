#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Render the attribute tables of a seed specification DOCX from the OpenAPI contract.

The seed (normally the latest formatted V3) supplies everything the spec does
not hold: cover, history, introduction, object chapters, sample frames, and
the table styling. For every documented endpoint that pairs with a spec record
(see `compare_openapi_doc.pair`), the data rows of its Header Request, Path,
Request, Response and Header Response tables are rebuilt from
`contract/from_openapi.json`:

  * a field the seed already lists keeps its row (and so its Type hyperlink
    to the object chapter); name, presence and description are rewritten,
    and the type only when the seed cell is plain text;
  * a new field clones a plain-text row of the same table;
  * the shared error rows that end every Response table stay as they are.

Endpoints with no spec record, and sample frames, are left untouched. The
seed file is never modified; the result is written to a new DOCX.

Usage:
    python scripts/render_from_openapi.py <seed.docx> <from_openapi.json> <out.docx>
"""
import collections
import copy
import json
import os
import re
import sys
import zipfile

from lxml import etree

from compare_openapi_doc import ERROR_ROWS, pair
from extract_doc_contract import TABLE_KINDS, cell_text, extract, para_text, qn, style_of

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def set_cell_text(tc, text):
    """Replace a plain cell's text, keeping the first run's formatting."""
    paras = tc.findall(qn("p"))
    first = paras[0]
    for extra in paras[1:]:
        tc.remove(extra)
    runs = first.findall(qn("r"))
    for r in runs[1:]:
        first.remove(r)
    for child in list(first):
        if etree.QName(child).localname not in ("pPr", "r"):
            first.remove(child)
    if runs:
        run = runs[0]
        for t in run.findall(qn("t")):
            run.remove(t)
    else:
        run = etree.SubElement(first, qn("r"))
    t = etree.SubElement(run, qn("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def has_link(tc):
    return tc.find(".//" + qn("hyperlink")) is not None


def plain_row(tr):
    """A five-cell data row whose Type cell is plain text, or None."""
    cells = tr.findall(qn("tc"))
    if len(cells) == 5 and cell_text(cells[1]) and not has_link(cells[2]):
        return tr
    return None


def fallback_rows(body):
    """First plain data row of each table kind, for tables the seed left empty."""
    found = {}
    for tbl in body.iter(qn("tbl")):
        trs = tbl.findall(qn("tr"))
        if len(trs) < 3:
            continue
        title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip().lower()
        kind = TABLE_KINDS.get(title)
        if kind and kind not in found:
            row = next((r for r in trs[2:] if plain_row(r) is not None), None)
            if row is not None:
                found[kind] = row
    return found


def rebuild_table(tbl, kind, fields, stats, fallback):
    trs = tbl.findall(qn("tr"))
    data = trs[2:]
    by_name, error_rows, plain = {}, [], None
    for tr in data:
        cells = tr.findall(qn("tc"))
        if len(cells) < 5:
            return False
        name = cell_text(cells[1])
        if kind == "response_fields" and name in ERROR_ROWS:
            error_rows.append(tr)
            continue
        by_name.setdefault(name, tr)
        if plain is None and not has_link(cells[2]):
            plain = tr
    if not fields and not by_name:
        return False
    if plain is not None:
        template = plain
    elif kind in fallback:
        template = fallback[kind]
        stats["tables using a shared template row"] += 1
    elif by_name:
        template = next(iter(by_name.values()))
    else:
        stats["tables skipped (no template row)"] += 1
        return False

    for tr in data:
        tbl.remove(tr)
    new_rows = []
    for f in fields:
        old = by_name.get(f["name"])
        tr = copy.deepcopy(old if old is not None else template)
        cells = tr.findall(qn("tc"))
        if old is None:
            stats["rows added"] += 1
            set_cell_text(cells[2], f["type"])
        else:
            stats["rows kept"] += 1
            if not has_link(cells[2]) and cell_text(cells[2]) != f["type"]:
                set_cell_text(cells[2], f["type"])
                stats["types rewritten"] += 1
        set_cell_text(cells[1], f["name"])
        set_cell_text(cells[3], f["presence"])
        # The owner keeps the seed's wording for a field it already documents;
        # only a new field takes the spec's description. The spec writes
        # Markdown; the document shows plain text, so code markers are dropped.
        if old is None:
            set_cell_text(cells[4], f["description"].replace("`", ""))
        else:
            stats["descriptions kept from seed"] += 1
        new_rows.append(tr)
    stats["rows removed"] += len([n for n in by_name if n not in {f["name"] for f in fields}])
    for i, tr in enumerate(new_rows + error_rows, 1):
        tbl.append(tr)
        set_cell_text(tr.findall(qn("tc"))[0], str(i))
    return True


W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
HTTP_CODE_RE = re.compile(r"Http Code\s*:\s*(\d{3})", re.I)
METHOD_LINE_RE = re.compile(r"^(GET|POST|PUT|DELETE|PATCH)\b")
MISSING = object()


def frame_body(txbx):
    """(paragraphs, index of the first JSON body line) of a sample frame, or None."""
    paras = txbx.findall(qn("p"))
    if not paras:
        return None
    for i, p in enumerate(paras):
        if para_text(p).strip().startswith(("{", "[")):
            return paras, i
    return None


def frame_kind(txbx):
    """'request', 'response' (a 2xx frame) or None for any other frame."""
    paras = txbx.findall(qn("p"))
    first = para_text(paras[0]).strip() if paras else ""
    if METHOD_LINE_RE.match(first):
        return "request"
    m = HTTP_CODE_RE.match(first)
    if m and m.group(1).startswith("2"):
        return "response"
    return None


def merged_body(seed_body, fields):
    """Top-level keys in spec order; seed values where the seed has the key,
    the spec property example otherwise. None when the result cannot be built."""
    if not isinstance(seed_body, dict) or not fields:
        return None
    out = {}
    for f in fields:
        value = seed_body.get(f["name"], MISSING)
        if value is MISSING:
            value = f.get("example", MISSING)
        if value is MISSING:
            # No value to show: leaving the key out is truer than inventing one.
            continue
        out[f["name"]] = value
    return out


def rewrite_frame(txbx, fields, stats):
    found = frame_body(txbx)
    if found is None:
        return
    paras, start = found
    body_paras = paras[start:]
    text = "\n".join(para_text(p) for p in body_paras)
    try:
        seed_body = json.loads(text)
    except ValueError:
        stats["frames skipped (seed body not JSON)"] += 1
        return
    new_body = merged_body(seed_body, fields)
    if new_body is None:
        stats["frames left as is (no object body or no fields)"] += 1
        return
    if new_body == seed_body:
        stats["frames unchanged"] += 1
        return
    # The template's first run must carry text; a blank line can be a run with
    # only a break, and cloning it would leave nowhere to write the line.
    template = next((p for p in body_paras
                     if p.findall(qn("r")) and p.find(qn("r")).find(qn("t")) is not None), None)
    if template is None:
        stats["frames skipped (no template run)"] += 1
        return
    for p in body_paras:
        txbx.remove(p)
    for line in json.dumps(new_body, ensure_ascii=False, indent=2).split("\n"):
        p = copy.deepcopy(template)
        for attr in ("paraId", "textId"):
            p.attrib.pop("{%s}%s" % (W14, attr), None)
        runs = p.findall(qn("r"))
        for r in runs[1:]:
            p.remove(r)
        t = runs[0].find(qn("t"))
        t.text = line
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        txbx.append(p)
    stats["frames rewritten"] += 1


def render(seed, contract_path, out):
    with open(contract_path, encoding="utf-8") as f:
        spec = json.load(f)["endpoints"]
    doc = extract(seed)
    pairs, _, _ = pair(spec, doc)
    # A pair whose titles differ came from the broad x-route fallback (e.g. a
    # documented route the backend does not have); rendering it would write
    # another operation's fields into that section, so it is left as is.
    by_section = {(d["category"], d["parent"], d["name"]): o for o, d in pairs if o["name"] == d["name"]}

    with zipfile.ZipFile(seed) as z:
        root = etree.fromstring(z.read("word/document.xml"), etree.XMLParser(huge_tree=True))
    body = root.find(qn("body"))
    fallback = fallback_rows(body)

    stats = collections.Counter({k: 0 for k in ("sections paired", "sections rendered", "tables rebuilt", "rows kept",
                            "rows added", "rows removed", "types rewritten", "descriptions kept from seed",
                            "tables using a shared template row", "tables skipped (no template row)")})
    stats["sections paired"] = len(pairs)
    heading = {2: "", 3: "", 4: ""}
    current, seen = None, set()
    for el in body:
        tag = etree.QName(el).localname
        if tag == "p":
            m = re.match(r"Heading(\d)$", style_of(el))
            if m and int(m.group(1)) in (2, 3, 4):
                lv = int(m.group(1))
                heading[lv] = para_text(el).strip()
                for k in range(lv + 1, 5):
                    heading[k] = ""
                key = (heading[2], heading[3] if lv == 4 else "", heading[lv])
                current = by_section.get(key) if lv in (3, 4) else None
                seen = set()
                if current is not None:
                    stats["sections rendered"] += 1
                continue
            if current is not None:
                for txbx in el.iter(qn("txbxContent")):
                    kind = frame_kind(txbx)
                    if kind == "request":
                        rewrite_frame(txbx, current["request_fields"], stats)
                    elif kind == "response":
                        rewrite_frame(txbx, current["response_fields"], stats)
        elif tag == "tbl" and current is not None:
            trs = el.findall(qn("tr"))
            if len(trs) < 2:
                continue
            title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip().lower()
            kind = TABLE_KINDS.get(title)
            if kind and kind not in seen:
                seen.add(kind)
                if rebuild_table(el, kind, current[kind], stats, fallback):
                    stats["tables rebuilt"] += 1

    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(seed) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    return stats


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        sys.exit(2)
    seed, contract_path, out = sys.argv[1:]
    if os.path.exists(out) and os.path.samefile(seed, out):
        sys.exit("Output must not be the seed document")
    stats = render(seed, contract_path, out)
    for k, v in stats.items():
        print("%-34s: %d" % (k, v))
    print("wrote %s" % out)


if __name__ == "__main__":
    main()
