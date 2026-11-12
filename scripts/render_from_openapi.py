#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Render the attribute tables of a seed specification DOCX from the OpenAPI contract.

The seed (normally the latest formatted V3) supplies everything the spec does
not hold: cover, earlier history, introduction, object chapters, sample
frames, and the table styling. For every documented endpoint that pairs with a spec record
(see `compare_openapi_doc.pair`), the data rows of its Header Request, Path,
Request, Response and Header Response tables are rebuilt from
`contract/from_openapi.json`:

  * a field the seed already lists keeps its row (and so its Type hyperlink
    to the object chapter); name, presence and description are rewritten,
    and the type only when the seed cell is plain text;
  * a new field clones a plain-text row of the same table;
  * the shared error rows that end every Response table stay as they are.

The seed file is never modified; the result is written to a new DOCX.

With `--meta`, the cover, introduction, abbreviations, connection table and
History come from the backend's `doc-meta.yaml` (see `doc_meta.py`); without
it they stay as the seed has them.

Usage:
    python scripts/render_from_openapi.py <seed.docx> <from_openapi.json> <out.docx> [--meta doc-meta.yaml]
        [--descriptions doc-descriptions.yaml] [--objects doc-objects.yaml]
        [--allow-stale-history]
"""
import argparse
import collections
import copy
import json
import os
import re
import sys
import zipfile

import yaml
from lxml import etree

from compare_openapi_doc import ERROR_ROWS, pair
from doc_meta import UniqueKeyLoader, cover_version, load as load_meta
from propose_history import read_snapshot, signatures
from extract_doc_contract import TABLE_KINDS, cell_text, extract, para_text, qn, style_of, to_key

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


def rebuild_table(tbl, kind, fields, stats, fallback, fresh=False, linker=None, wording=None):
    """Rebuild a table's data rows from `fields`.

    `fresh` is for a section cloned from another endpoint: its rows only lend
    their formatting, so no row keeps the template's name, type or wording.
    `linker` links a Type cell naming a component schema to its object section.
    `wording` is this table's entry of the description file ({field: lines});
    when given (even empty) the Description cells come from it, falling back
    to the spec, and never from the seed.
    """
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
    if fresh:
        by_name = {}
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
        # The Type cell follows the spec: a field that points at a component
        # schema links to that schema's object section. Any other link the seed
        # had goes, except one into the ENUMERATION chapter, which the spec
        # cannot express yet.
        if linker is not None and f.get("ref") and linker.can_link(f["ref"]):
            linker.link(cells[2], f, stats)
        elif old is not None and has_link(cells[2]):
            if linker is not None and linker.keeps_link(cells[2]):
                stats["enumeration links kept"] += 1
            else:
                set_cell_text(cells[2], f["type"])
                stats["object links dropped (no schema in the spec)"] += 1
        elif old is not None and cell_text(cells[2]) != f["type"]:
            set_cell_text(cells[2], f["type"])
            stats["types rewritten"] += 1
        set_cell_text(cells[1], f["name"])
        set_cell_text(cells[3], f["presence"])
        spec_text = f["description"].replace("`", "")
        if wording is not None:
            # Description file first, the spec otherwise. An unchanged cell is
            # left alone, so it keeps its runs.
            lines = wording.get(f["name"]) or ([spec_text] if spec_text else [])
            if not lines:
                put(cells[4], "")
            elif put_lines(cells[4], lines):
                stats["descriptions written"] += 1
            stats["descriptions from the description file" if f["name"] in wording
                  else "descriptions from the spec"] += 1
        elif old is None:
            # The spec writes Markdown; the document shows plain text.
            set_cell_text(cells[4], spec_text)
        else:
            # No description file: a field the seed documents keeps its wording.
            stats["descriptions kept from seed"] += 1
        new_rows.append(tr)
    stats["rows removed"] += len([n for n in by_name if n not in {f["name"] for f in fields}])
    grid = [gc.get(qn("w")) for gc in tbl.findall(qn("tblGrid") + "/" + qn("gridCol"))]
    for i, tr in enumerate(new_rows + error_rows, 1):
        tbl.append(tr)
        set_cell_text(tr.findall(qn("tc"))[0], str(i))
        if template is not plain:
            # A row borrowed from another table carries that table's column
            # widths; give it this table's grid.
            for tc, width in zip(tr.findall(qn("tc")), grid):
                tcw = tc.find(qn("tcPr") + "/" + qn("tcW"))
                if tcw is not None and width:
                    tcw.set(qn("w"), width)
                    tcw.set(qn("type"), "dxa")
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


def rewrite_frame(txbx, fields, stats, fresh=False):
    """Rewrite a frame's JSON body. `fresh` ignores the seed's values (cloned section)."""
    found = frame_body(txbx)
    if found is None:
        return
    paras, start = found
    body_paras = paras[start:]
    text = "\n".join(para_text(p) for p in body_paras)
    try:
        seed_body = json.loads(text)
    except ValueError:
        if not fresh:
            stats["frames skipped (seed body not JSON)"] += 1
            return
        seed_body = {}
    if fresh:
        if not fields:
            # The new endpoint has no JSON body: drop the template's.
            for p in body_paras:
                txbx.remove(p)
            stats["frame bodies removed (no body in spec)"] += 1
            return
        seed_body = {} if isinstance(seed_body, dict) else seed_body
    new_body = merged_body(seed_body if not fresh else {}, fields)
    if new_body is None:
        stats["frames left as is (no object body or no fields)"] += 1
        return
    if new_body == seed_body and not fresh:
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


NOT_IMPLEMENTED = " (Not implemented yet)"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
VML = "urn:schemas-microsoft-com:vml"
VML_O = "urn:schemas-microsoft-com:office:office"


def has_content_type(record):
    return any(h["name"].lower() == "content-type" for h in record["headers"])


def fix_request_frame(txbx, record, stats):
    """Align a request frame's first line and header lines with the spec record.

    The first line is `METHOD {baseURL}/path` split into runs (method badge,
    space, red {baseURL}, path pieces). When the spec path differs, the path
    runs are replaced by one run holding the spec path, keeping the first path
    run's formatting. A `Content-Type:` header line is dropped when the
    endpoint has no request body.
    """
    paras = txbx.findall(qn("p"))
    if not paras:
        return
    runs = [r for r in paras[0].findall(qn("r")) if r.find(qn("t")) is not None]
    texts = [r.find(qn("t")).text or "" for r in runs]
    base = next((i for i, t in enumerate(texts) if "{baseURL}" in t), None)
    spec_path = re.sub(r"^\{baseURL\}", "", record["path"])
    if base is not None:
        line = "".join(texts[base:])
        frame_path = line[line.index("{baseURL}") + len("{baseURL}"):]
        if to_key(frame_path) != to_key(spec_path):
            space = "{http://www.w3.org/XML/1998/namespace}space"
            head, _, _ = texts[base].partition("{baseURL}")
            runs[base].find(qn("t")).text = head + "{baseURL}"
            for r in runs[base + 1:]:
                r.getparent().remove(r)
            # The path takes the formatting of the run before {baseURL} (black,
            # bold), never the red {baseURL} run's.
            path_run = copy.deepcopy(runs[base - 1] if base > 0 else runs[base])
            for t in path_run.findall(qn("t"))[1:]:
                path_run.remove(t)
            for sh in path_run.iter(qn("shd")):
                sh.getparent().remove(sh)
            path_run.find(qn("t")).text = spec_path
            path_run.find(qn("t")).set(space, "preserve")
            runs[base].addnext(path_run)
            stats["request URLs rewritten"] += 1
    if not has_content_type(record):
        found = frame_body(txbx)
        end = found[1] if found else len(paras)
        for p in paras[1:end]:
            if para_text(p).strip().lower().startswith("content-type"):
                txbx.remove(p)
                stats["Content-Type lines removed"] += 1
    elif record.get("media_types"):
        # A cloned template, or the seed itself, may show another media type.
        want = "Content-Type: %s" % record["media_types"][0]
        found = frame_body(txbx)
        end = found[1] if found else len(paras)
        for p in paras[1:end]:
            text = para_text(p).strip()
            if text.lower().startswith("content-type") and text != want:
                key_run = next((r for r in p.findall(qn("r")) if r.find(qn("t")) is not None), None)
                if key_run is None:
                    continue
                # Keep the bold "Content-Type" run; rewrite only the value runs.
                runs = [r for r in p.findall(qn("r")) if r.find(qn("t")) is not None]
                key_run.find(qn("t")).text = "Content-Type"
                if len(runs) > 1:
                    runs[1].find(qn("t")).text = ": %s" % record["media_types"][0]
                    runs[1].find(qn("t")).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for r in runs[2:]:
                        p.remove(r)
                else:
                    key_run.find(qn("t")).text = want
                stats["Content-Type lines rewritten"] += 1


def fix_response_media(txbx, record, stats):
    """Make a cloned response frame's `Content-type` line state this endpoint's media type."""
    media = record.get("response_media_types")
    if media is None:
        return
    paras = txbx.findall(qn("p"))
    found = frame_body(txbx)
    for p in paras[1:found[1] if found else len(paras)]:
        text = para_text(p).strip()
        if not text.lower().startswith("content-type"):
            continue
        if not media:
            txbx.remove(p)
            stats["response Content-type lines removed"] += 1
            continue
        value = text.split(":", 1)[1].strip() if ":" in text else ""
        runs = [r for r in p.findall(qn("r")) if r.find(qn("t")) is not None]
        if value == media[0] or len(runs) < 2:
            continue
        # Keep the bold key run; rewrite only the value.
        key = runs[0].find(qn("t")).text or ""
        runs[0].find(qn("t")).text = key.split(":", 1)[0]
        runs[1].find(qn("t")).text = ": %s" % media[0]
        runs[1].find(qn("t")).set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for r in runs[2:]:
            p.remove(r)
        stats["response Content-type lines rewritten"] += 1


def set_par_text(p, text):
    """Replace a paragraph's text, keeping its first text run's formatting."""
    runs = [r for r in p.findall(qn("r")) if r.find(qn("t")) is not None]
    if not runs:
        return
    for r in p.findall(qn("r"))[1:]:
        if r is not runs[0]:
            p.remove(r)
    t = runs[0].find(qn("t"))
    for extra in runs[0].findall(qn("t"))[1:]:
        runs[0].remove(extra)
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def mark_not_implemented(heading, stats):
    text = para_text(heading)
    if text.rstrip().endswith(NOT_IMPLEMENTED.strip()):
        return
    runs = [r for r in heading.findall(qn("r")) if r.find(qn("t")) is not None]
    if not runs:
        return
    note = copy.deepcopy(runs[-1])
    for t in note.findall(qn("t"))[1:]:
        note.remove(t)
    t = note.find(qn("t"))
    t.text = NOT_IMPLEMENTED.lstrip() if text.endswith((" ", "\t")) else NOT_IMPLEMENTED
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    runs[-1].addnext(note)
    stats["sections marked not implemented"] += 1


class IdAllocator:
    """Unique drawing ids for cloned content (Word reports duplicates as corrupt)."""

    def __init__(self, root):
        ids = [int(e.get("id")) for e in root.iter("{%s}docPr" % WP) if (e.get("id") or "").isdigit()]
        self.next_doc_pr = max(ids or [0]) + 1
        nums = [int(m.group(1)) for e in root.iter("{%s}shape" % VML)
                for m in [re.search(r"_x0000_s(\d+)", (e.get("{%s}spid" % VML_O) or "") + (e.get("id") or ""))] if m]
        self.next_shape = max(nums or [1024]) + 1

    def renumber(self, el):
        for b in list(el.iter(qn("bookmarkStart"))) + list(el.iter(qn("bookmarkEnd"))):
            if b.getparent() is not None:
                b.getparent().remove(b)
        for p in el.iter(qn("p")):
            for attr in ("paraId", "textId"):
                p.attrib.pop("{%s}%s" % (W14, attr), None)
        if el.tag == qn("p"):
            for attr in ("paraId", "textId"):
                el.attrib.pop("{%s}%s" % (W14, attr), None)
        for d in el.iter("{%s}docPr" % WP):
            d.set("id", str(self.next_doc_pr))
            self.next_doc_pr += 1
        for s in el.iter("{%s}shape" % VML):
            sid = "_x0000_s%d" % self.next_shape
            self.next_shape += 1
            s.set("id", sid)
            s.set("{%s}spid" % VML_O, sid)


def section_ranges(body):
    """[(category, heading level, title, start, end)] for every Heading 3/4 in body order."""
    els = list(body)
    out, category, open_idx = [], "", None
    for i, el in enumerate(els):
        if etree.QName(el).localname != "p":
            continue
        m = re.match(r"Heading(\d)$", style_of(el))
        if not m:
            continue
        lv = int(m.group(1))
        if open_idx is not None:
            out[-1][4] = i
            open_idx = None
        if lv == 2:
            category = para_text(el).strip()
        if lv in (3, 4):
            out.append([category, lv, para_text(el).strip(), i, len(els)])
            open_idx = i
    return out


def category_end(body, category):
    """Index just after the last element of a Heading 2 block, or None."""
    els = list(body)
    inside = False
    for i, el in enumerate(els):
        if etree.QName(el).localname != "p":
            continue
        m = re.match(r"Heading(\d)$", style_of(el))
        if not m:
            continue
        lv = int(m.group(1))
        if inside and lv <= 2:
            return i
        if lv == 2 and para_text(el).strip() == category:
            inside = True
    if not inside:
        return None
    # The body's last child must stay the final section properties.
    if els and etree.QName(els[-1]).localname == "sectPr":
        return len(els) - 1
    return len(els)


def snap_to_grid(tbl):
    """Give every unmerged row the table's grid widths.

    Some seed tables carry cell widths off their grid (TECHNICAL_DEBT); a new
    section must not copy the defect.
    """
    grid = [gc.get(qn("w")) for gc in tbl.findall(qn("tblGrid") + "/" + qn("gridCol"))]
    for tr in tbl.findall(qn("tr")):
        cells = tr.findall(qn("tc"))
        if len(cells) != len(grid):
            continue
        for tc, width in zip(cells, grid):
            tcw = tc.find(qn("tcPr") + "/" + qn("tcW"))
            if tcw is not None and width:
                tcw.set(qn("w"), width)
                tcw.set(qn("type"), "dxa")


def table_kind(tbl):
    trs = tbl.findall(qn("tr"))
    title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip().lower() if trs else ""
    return TABLE_KINDS.get(title)


def is_error_frame(txbx):
    lines = txbx.findall(qn("p"))
    first = para_text(lines[0]).strip() if lines else ""
    return frame_kind(txbx) is None and bool(HTTP_CODE_RE.match(first))


def frame_text(txbx):
    return "\n".join(para_text(p).strip() for p in txbx.findall(qn("p")))


# The order attribute tables take inside a section.
TABLE_ORDER = ("headers", "path_params", "request_fields", "response_headers", "response_fields")


def find_donors(body):
    """Parts a new section may lack in its template, taken from the seed.

    `error`: a paragraph holding only the seed's shared error sample (the
    frame most sections repeat). `tables`: per table kind, a table with a
    plain data row and the spacer paragraph that precedes it.
    """
    texts = collections.Counter(frame_text(t) for t in body.iter(qn("txbxContent")) if is_error_frame(t))
    shared = texts.most_common(1)[0][0] if texts else None
    donors = {"error": None, "tables": {}}
    prev = None
    for el in body:
        tag = etree.QName(el).localname
        if tag == "p" and donors["error"] is None and shared is not None:
            frames = list(el.iter(qn("txbxContent")))
            if frames and all(is_error_frame(t) and frame_text(t) == shared for t in frames):
                donors["error"] = el
        elif tag == "tbl":
            kind = table_kind(el)
            rows = el.findall(qn("tr"))
            if kind and kind not in donors["tables"] and any(plain_row(r) is not None for r in rows[2:]):
                spacer = prev if (prev is not None and prev.tag == qn("p")
                                  and not para_text(prev).strip()
                                  and not list(prev.iter(qn("txbxContent")))) else None
                donors["tables"][kind] = (el, spacer)
        prev = el
    return donors


def body_kind(record):
    return bool(record["request_fields"]) or has_content_type(record)


def build_section(template_els, record, ids, fallback, stats, linker=None, donors=None, wording=None):
    """Clone a rendered section and fill it from `record`."""
    # Body-level bookmarks name the template's heading; a clone must not carry them.
    els = [copy.deepcopy(e) for e in template_els
           if etree.QName(e).localname not in ("bookmarkStart", "bookmarkEnd")]
    for e in els:
        ids.renumber(e)
    heading = els[0]
    set_par_text(heading, record["name"])
    if record.get("status") == "not-implemented":
        mark_not_implemented(heading, stats)

    # Description: the text between the heading and "Sample Request".
    described, drop = False, []
    for e in els[1:]:
        if etree.QName(e).localname != "p":
            continue
        text = para_text(e).strip()
        if text == "Sample Request" or list(e.iter(qn("txbxContent"))):
            break
        if not text:
            continue
        if not described:
            set_par_text(e, record["description"].replace("`", "") or record["name"])
            described = True
        else:
            drop.append(e)

    seen, tables = set(), {}
    for e in els[1:]:
        tag = etree.QName(e).localname
        if tag == "p":
            frames = list(e.iter(qn("txbxContent")))
            if not frames:
                continue
            # The template's error frames describe another endpoint's failures.
            # A paragraph can anchor a success frame and an error frame
            # together, so each error frame's own drawing run is removed.
            kept = []
            for txbx in frames:
                if is_error_frame(txbx):
                    run = next((a for a in txbx.iterancestors() if a.tag == qn("r")), None)
                    if run is not None and run.getparent() is not None:
                        run.getparent().remove(run)
                        stats["template error frames removed"] += 1
                else:
                    kept.append(txbx)
            if not kept:
                drop.append(e)
                continue
            for txbx in kept:
                k = frame_kind(txbx)
                if k == "request":
                    fix_request_frame(txbx, record, stats)
                    # Header lines come from the template endpoint; keep only
                    # the headers this endpoint declares.
                    names = {h["name"].lower() for h in record["headers"]}
                    lines = txbx.findall(qn("p"))
                    found = frame_body(txbx)
                    for p in lines[1:found[1] if found else len(lines)]:
                        head = para_text(p).split(":", 1)[0].strip().lower()
                        if head and head not in names:
                            txbx.remove(p)
                            stats["template header lines removed"] += 1
                    rewrite_frame(txbx, record["request_fields"], stats, fresh=True)
                elif k == "response":
                    code = record.get("success_status") or "200"
                    lines = txbx.findall(qn("p"))
                    if lines:
                        set_par_text(lines[0], "Http Code: %s" % code)
                    fix_response_media(txbx, record, stats)
                    rewrite_frame(txbx, record["response_fields"], stats, fresh=True)
        elif tag == "tbl":
            kind = table_kind(e)
            if not kind or kind in seen:
                continue
            seen.add(kind)
            if not record[kind]:
                drop.append(e)
                continue
            rebuild_table(e, kind, record[kind], stats, fallback, fresh=True, linker=linker,
                          wording=None if wording is None else wording.get(kind, {}))
            snap_to_grid(e)
            tables[kind] = e
    out = [e for e in els if not any(e is d for d in drop)]
    donors = donors or {"error": None, "tables": {}}

    # The seed ends almost every Sample Response with one shared error sample
    # (an unusable access token). An endpoint that takes a token gets it too.
    if donors["error"] is not None and any(h["name"] == "Authorization" for h in record["headers"]):
        anchor = None
        for e in out:
            if e.tag == qn("p") and any(frame_kind(t) == "response" for t in e.iter(qn("txbxContent"))):
                anchor = e
        if anchor is not None:
            frame = copy.deepcopy(donors["error"])
            ids.renumber(frame)
            out.insert(next(i for i, e in enumerate(out) if e is anchor) + 1, frame)
            stats["shared error frames added"] += 1

    # A table kind the template section does not have comes from the seed.
    for pos, kind in enumerate(TABLE_ORDER):
        if not record[kind] or kind in tables:
            continue
        donor = donors["tables"].get(kind)
        before = [tables[k] for k in TABLE_ORDER[:pos] if k in tables]
        after = [tables[k] for k in TABLE_ORDER[pos + 1:] if k in tables]
        if donor is None or not (before or after):
            stats["new sections missing a %s table" % kind] += 1
            continue
        table = copy.deepcopy(donor[0])
        ids.renumber(table)
        rebuild_table(table, kind, record[kind], stats, fallback, fresh=True, linker=linker,
                      wording=None if wording is None else wording.get(kind, {}))
        snap_to_grid(table)
        new = [table]
        if donor[1] is not None:
            spacer = copy.deepcopy(donor[1])
            ids.renumber(spacer)
            new = [spacer, table] if before else [table, spacer]
        at = (next(i for i, e in enumerate(out) if e is before[-1]) + 1 if before
              else next(i for i, e in enumerate(out) if e is after[0]))
        out[at:at] = new
        tables[kind] = table
        stats["tables added from the seed"] += 1
    return out


def pick_template(sections, record):
    """Best rendered section to clone for a new endpoint.

    The same method keeps the badge colour right, and the same body/no-body
    shape keeps the sample frames right; when no section has both (e.g. a GET
    with a body), the body shape is dropped from the match first.
    """
    for strict in (True, False):
        best, best_score = None, -1
        for sec in sections:
            rec = sec["record"]
            if rec["method"] != record["method"]:
                continue
            if strict and body_kind(rec) != body_kind(record):
                continue
            score = 3 if sec["category"] == record["category"] else 0
            score += sum(1 for k in TABLE_KINDS.values() if bool(rec[k]) == bool(record[k]))
            if score > best_score:
                best, best_score = sec, score
        if best is not None:
            return best
    return None


class ObjectLinker:
    """Builds the object chapters from the spec and links Type cells to them.

    The seed's REQUEST/RESPONSE OBJECT and RELATIVE OBJECT sections are
    removed; one of them lends its heading and table as the template. Every
    component schema a rendered field points at (directly or through another
    object) gets a section: Heading 2 with a bookmark, and an attribute table.
    A Type cell naming the schema becomes a `w:hyperlink` to that bookmark
    (the document's convention).

    `naming` ({schema: {title, chapter}}) gives a section its document name
    and chapter; a schema without an entry keeps its own name and goes to
    RELATIVE OBJECT. `wording` ({schema: {field: lines}}) is the object part
    of the description file.
    """

    CHAPTERS = ("REQUEST/RESPONSE OBJECT", "RELATIVE OBJECT")
    DEFAULT = "RELATIVE OBJECT"
    ENUMERATION = "ENUMERATION"

    def __init__(self, root, body, objects, stats, naming=None, wording=None):
        self.body = body
        self.objects = objects
        self.naming = naming or {}
        self.wording = wording
        self.created = collections.OrderedDict()   # schema name -> anchor
        self.queue = []
        ids = [int(b.get(qn("id"))) for b in root.iter(qn("bookmarkStart"))
               if (b.get(qn("id")) or "").isdigit()]
        self.next_id = max(ids or [0]) + 1
        self.enum_anchors = self.chapter_anchors(self.ENUMERATION)
        self.parts = self.template()
        self.can_create = self.parts is not None and self.chapter_end(self.DEFAULT) is not None
        if self.can_create:
            self.clear_chapters(stats)
        self.marks = {b.get(qn("name")) for b in root.iter(qn("bookmarkStart"))}

    def chapter_anchors(self, chapter):
        """Bookmark names on (or just before) the Heading 2 paragraphs of a chapter."""
        names, inside, pending = set(), False, []
        for el in self.body:
            tag = etree.QName(el).localname
            if tag == "bookmarkStart":
                pending.append(el.get(qn("name")))
                continue
            if tag == "p":
                st = style_of(el)
                if st == "Heading1":
                    inside = para_text(el).strip().upper() == chapter
                elif inside and st == "Heading2":
                    names.update(pending)
                    names.update(b.get(qn("name")) for b in el.iter(qn("bookmarkStart")))
            if tag != "bookmarkEnd":
                pending = []
        return names

    def clear_chapters(self, stats):
        """Remove every section of the object chapters; the Heading 1 and its lead-in stay."""
        chapter, started = None, False
        for el in list(self.body):
            tag = etree.QName(el).localname
            if tag == "p" and style_of(el) == "Heading1":
                name = para_text(el).strip().upper()
                chapter, started = (name if name in self.CHAPTERS else None), False
                continue
            if chapter is None or tag == "sectPr":
                continue
            if tag == "p" and style_of(el) == "Heading2":
                started = True
                stats["seed object sections removed"] += 1
            if started:
                self.body.remove(el)

    def title(self, ref):
        return (self.naming.get(ref) or {}).get("title") or self.objects[ref]["label"]

    def can_link(self, ref):
        return ref in self.objects and self.can_create

    def keeps_link(self, tc):
        link = tc.find(".//" + qn("hyperlink"))
        return link is not None and link.get(qn("anchor")) in self.enum_anchors

    @staticmethod
    def inner(label):
        m = re.match(r"\s*List<(.*)>\s*$", label)
        return (m.group(1) if m else label).strip()

    def anchor_for(self, ref, stats):
        if ref not in self.created:
            anchor = "_OBJ_%s" % re.sub(r"[^A-Za-z0-9_]", "_", ref)
            while anchor in self.marks:
                anchor += "_"
            self.marks.add(anchor)
            self.created[ref] = anchor
            self.queue.append(ref)
        stats["type links to objects"] += 1
        return self.created[ref]

    def link(self, tc, field, stats):
        """Write a Type cell as `List<` + hyperlink + `>` (or a bare hyperlink)."""
        anchor = self.anchor_for(field["ref"], stats)
        title = self.title(field["ref"])
        is_list = field["type"].strip().startswith("List<")
        set_cell_text(tc, title)
        p = tc.find(qn("p"))
        run = p.find(qn("r"))

        def piece(text):
            r = copy.deepcopy(run)
            r.find(qn("t")).text = text
            return r

        link = etree.Element(qn("hyperlink"))
        link.set(qn("anchor"), anchor)
        link.set(qn("history"), "1")
        linked = piece(title)
        rpr = linked.find(qn("rPr"))
        if rpr is None:
            rpr = etree.Element(qn("rPr"))
            linked.insert(0, rpr)
        # The document's Type links use the Hyperlink character style in black.
        for style in rpr.findall(qn("rStyle")):
            rpr.remove(style)
        style = etree.Element(qn("rStyle"))
        style.set(qn("val"), "Hyperlink")
        rpr.insert(0, style)
        color = rpr.find(qn("color"))
        if color is None:
            color = etree.Element(qn("color"))
            anchor_el = rpr.find(qn("rFonts"))
            if anchor_el is not None:
                anchor_el.addnext(color)
            else:
                style.addnext(color)
        color.set(qn("val"), "000000")
        link.append(linked)
        p.remove(run)
        if is_list:
            p.append(piece("List<"))
        p.append(link)
        if is_list:
            p.append(piece(">"))

    def template(self):
        """Copies of (heading, table, spacer) of a seed object section with a plain data row."""
        els = list(self.body)
        inside = False
        for i, el in enumerate(els):
            if etree.QName(el).localname != "p":
                continue
            st = style_of(el)
            if st == "Heading1":
                inside = para_text(el).strip().upper() == self.DEFAULT
            elif inside and st == "Heading2" and i + 1 < len(els) and els[i + 1].tag == qn("tbl"):
                rows = els[i + 1].findall(qn("tr"))
                if len(rows) > 2 and any(plain_row(r) is not None for r in rows[2:]):
                    spacer = els[i + 2] if i + 2 < len(els) and els[i + 2].tag == qn("p") else None
                    if spacer is not None and (style_of(spacer) or para_text(spacer).strip()):
                        spacer = None
                    return (copy.deepcopy(el), copy.deepcopy(els[i + 1]),
                            copy.deepcopy(spacer) if spacer is not None else None)
        return None

    def chapter_end(self, chapter):
        els = list(self.body)
        inside = False
        for i, el in enumerate(els):
            if etree.QName(el).localname == "p" and style_of(el) == "Heading1":
                if inside:
                    return i
                inside = para_text(el).strip().upper() == chapter
        if not inside:
            return None
        # The chapter is the last one: insert before the body's final sectPr.
        if els and etree.QName(els[-1]).localname == "sectPr":
            return len(els) - 1
        return len(els)

    @staticmethod
    def strip_ids(el):
        for node in [el] + list(el.iter(qn("p"), qn("tr"))):
            for attr in ("paraId", "textId"):
                node.attrib.pop("{%s}%s" % (W14, attr), None)

    def flush(self, stats, fallback):
        """Create the queued object sections; their own fields may queue more."""
        if not self.can_create:
            stats["object sections skipped (no template or chapter)"] += len(self.queue)
            return
        t_heading, t_table, t_spacer = self.parts
        while self.queue:
            ref = self.queue.pop(0)
            obj = self.objects[ref]
            title = self.title(ref)
            chapter = ((self.naming.get(ref) or {}).get("chapter") or self.DEFAULT).upper()
            at = self.chapter_end(chapter)
            if at is None:
                at = self.chapter_end(self.DEFAULT)
            heading = copy.deepcopy(t_heading)
            for b in list(heading.iter(qn("bookmarkStart"))) + list(heading.iter(qn("bookmarkEnd"))):
                b.getparent().remove(b)
            self.strip_ids(heading)
            set_par_text(heading, title)
            start = etree.Element(qn("bookmarkStart"))
            start.set(qn("id"), str(self.next_id))
            start.set(qn("name"), self.created[ref])
            end = etree.Element(qn("bookmarkEnd"))
            end.set(qn("id"), str(self.next_id))
            self.next_id += 1
            ppr = heading.find(qn("pPr"))
            if ppr is not None:
                ppr.addnext(start)
            else:
                heading.insert(0, start)
            heading.append(end)

            table = copy.deepcopy(t_table)
            self.strip_ids(table)
            set_cell_text(table.find(qn("tr")).find(qn("tc")), title)
            rebuild_table(table, "object", obj["fields"], stats, fallback, fresh=True, linker=self,
                          wording=None if self.wording is None else self.wording.get(ref, {}))
            snap_to_grid(table)
            new = [heading, table]
            if t_spacer is not None:
                spacer = copy.deepcopy(t_spacer)
                self.strip_ids(spacer)
                new.append(spacer)
            for el in new:
                self.body.insert(at, el)
                at += 1
            stats["object sections added"] += 1


def drop_half_bookmarks(root, stats):
    """Remove a bookmark start or end whose other half went with rewritten content."""
    starts = {b.get(qn("id")) for b in root.iter(qn("bookmarkStart"))}
    ends = {b.get(qn("id")) for b in root.iter(qn("bookmarkEnd"))}
    for b in list(root.iter(qn("bookmarkStart"), qn("bookmarkEnd"))):
        if b.get(qn("id")) not in (ends if b.tag == qn("bookmarkStart") else starts):
            b.getparent().remove(b)
            stats["half bookmarks removed"] += 1


def unlink_dangling(root, stats):
    """Turn a table-cell link whose bookmark no longer exists into plain text.

    Sections the render leaves alone (documented but not implemented) can
    still point at an object section that was removed; the seed also carries
    a few links that were broken before.
    """
    marks = {b.get(qn("name")) for b in root.iter(qn("bookmarkStart"))}
    for link in list(root.iter(qn("hyperlink"))):
        anchor = link.get(qn("anchor"))
        if not anchor or anchor in marks or not any(a.tag == qn("tc") for a in link.iterancestors()):
            continue
        parent = link.getparent()
        at = parent.index(link)
        for run in list(link):
            for style in run.findall(qn("rPr") + "/" + qn("rStyle")):
                style.getparent().remove(style)
            parent.insert(at, run)
            at += 1
        parent.remove(link)
        stats["dangling links unlinked"] += 1


def add_new_sections(body, spec, by_section, fallback, stats, ids, linker=None, donors=None,
                     descriptions=None):
    """Append a section for every spec record the seed does not document.

    Each new section clones the closest rendered section (same method and
    body/no-body shape, same category preferred) and goes at the end of its
    category. Records sharing a title in a category (one operation served on
    several paths) become one section that lists the other paths.
    """
    used = {id(r) for r in by_section.values()}
    rendered = []
    els = list(body)
    for category, lv, title, start, end in section_ranges(body):
        rec = by_section.get((category, "", title)) if lv == 3 else None
        if rec is not None:
            rendered.append({"category": category, "record": rec, "els": els[start:end]})

    groups = collections.OrderedDict()
    for rec in spec:
        if id(rec) in used:
            continue
        groups.setdefault((rec["category"], rec["name"]), []).append(rec)

    for (category, name), recs in groups.items():
        record = dict(recs[0])
        if len(recs) > 1:
            others = ", ".join("%s %s" % (r["method"], re.sub(r"^\{baseURL\}", "", r["path"])) for r in recs[1:])
            record["description"] = "%s Also served at: %s." % (record["description"], others)
        template = pick_template(rendered, record)
        at = category_end(body, category)
        if template is None or at is None:
            stats["new sections skipped (no template or category)"] += 1
            continue
        for offset, el in enumerate(build_section(
                template["els"], record, ids, fallback, stats, linker, donors,
                None if descriptions is None else descriptions.get((category, name), {}))):
            body.insert(at + offset, el)
        stats["sections added"] += 1


def history_table(body):
    """The table under the `History` title, or None."""
    els = list(body)
    for i, el in enumerate(els):
        if el.tag == qn("p") and style_of(el) == "Title" and para_text(el).strip() == "History":
            return next((e for e in els[i + 1:i + 4] if e.tag == qn("tbl")), None)
    return None


# Links into the ENUMERATION chapter are written `[label](#Section title)` in
# description text, so a description file can carry them without a bookmark name.
ENUM_LINK_RE = re.compile(r"\[([^\]]+)\]\(#([^)]+)\)")
ENUMS = {"title_of": {}, "anchor_of": {}}


def index_enumerations(body):
    """Record the ENUMERATION chapter's sections: bookmark name <-> title."""
    ENUMS["title_of"].clear()
    ENUMS["anchor_of"].clear()
    inside, pending = False, []
    for el in body:
        tag = etree.QName(el).localname
        if tag == "bookmarkStart":
            pending.append(el.get(qn("name")))
            continue
        if tag == "p":
            st = style_of(el)
            if st == "Heading1":
                inside = para_text(el).strip().upper() == "ENUMERATION"
            elif inside and st == "Heading2":
                title = para_text(el).strip()
                for name in pending + [b.get(qn("name")) for b in el.iter(qn("bookmarkStart"))]:
                    ENUMS["title_of"][name] = title
                    ENUMS["anchor_of"].setdefault(title, name)
        if tag != "bookmarkEnd":
            pending = []


def cell_lines(tc):
    """A cell's text as lines: one per paragraph, split again at line breaks.

    A link into the ENUMERATION chapter reads `[label](#Section title)`.
    """
    def text_of(node):
        out = ""
        for child in node:
            tag = etree.QName(child).localname
            if tag == "t":
                out += child.text or ""
            elif tag == "br":
                out += "\n"
            elif tag == "hyperlink":
                label = text_of(child)
                title = ENUMS["title_of"].get(child.get(qn("anchor")))
                out += "[%s](#%s)" % (label, title) if title and label.strip() else label
            elif tag not in ("pPr", "rPr"):
                out += text_of(child)
        return out

    lines = []
    for p in tc.findall(qn("p")):
        lines += [" ".join(part.split()) for part in text_of(p).split("\n") if part.strip()]
    return lines


def put(tc, text):
    """Write a cell only when its text differs, so untouched cells keep every run."""
    if " ".join(cell_text(tc).split()) != text:
        set_cell_text(tc, text)
        return True
    return False


def put_lines(tc, lines):
    """Write a cell as one paragraph per line; unchanged cells are left alone."""
    if cell_lines(tc) == lines:
        return False
    # The template paragraph keeps its properties and one text run; links,
    # bookmarks and breaks of the old content must not be carried over.
    template = copy.deepcopy(tc.find(qn("p")))
    runs = template.findall(".//" + qn("r"))
    run = next((r for r in runs if r.find(qn("t")) is not None), runs[0] if runs else None)
    for child in list(template):
        if etree.QName(child).localname != "pPr":
            template.remove(child)
    if run is None:
        run = etree.Element(qn("r"))
    for child in list(run):
        if etree.QName(child).localname != "rPr":
            run.remove(child)
    rpr = run.find(qn("rPr"))
    if rpr is not None:
        for style in rpr.findall(qn("rStyle")):
            rpr.remove(style)        # e.g. the Hyperlink character style
    etree.SubElement(run, qn("t"))
    template.append(run)
    ObjectLinker.strip_ids(template)
    for p in tc.findall(qn("p")):
        tc.remove(p)
    space = "{http://www.w3.org/XML/1998/namespace}space"
    for line in lines:
        p = copy.deepcopy(template)
        base = p.find(qn("r"))
        p.remove(base)

        def piece(text):
            r = copy.deepcopy(base)
            r.find(qn("t")).text = text
            r.find(qn("t")).set(space, "preserve")
            return r

        at = 0
        for m in ENUM_LINK_RE.finditer(line):
            anchor = ENUMS["anchor_of"].get(m.group(2))
            if anchor is None:
                continue                      # unknown section: left as written
            if m.start() > at:
                p.append(piece(line[at:m.start()]))
            link = etree.SubElement(p, qn("hyperlink"))
            link.set(qn("anchor"), anchor)
            link.set(qn("history"), "1")
            run = piece(m.group(1))
            rpr = run.find(qn("rPr"))
            if rpr is None:
                rpr = etree.Element(qn("rPr"))
                run.insert(0, rpr)
            style = etree.Element(qn("rStyle"))
            style.set(qn("val"), "Hyperlink")
            rpr.insert(0, style)
            if rpr.find(qn("color")) is None:
                color = etree.Element(qn("color"))
                color.set(qn("val"), "000000")
                fonts = rpr.find(qn("rFonts"))
                (fonts if fonts is not None else style).addnext(color)
            link.append(run)
            at = m.end()
        if at < len(line) or at == 0:
            p.append(piece(line[at:]))
        tc.append(p)
    return True


def fill_rows(tbl, rows, first, stats, label):
    """Make the table's data rows (from index `first`) show `rows` (lists of cell texts).

    Rows are reused in place; a missing row clones the last data row and a
    surplus row is removed. A cell value of None keeps the cell (row number).
    """
    trs = tbl.findall(qn("tr"))
    data = trs[first:]
    if not data:
        stats["%s skipped (no data row to clone)" % label] += 1
        return
    for extra in data[len(rows):]:
        tbl.remove(extra)
        stats["%s rows removed" % label] += 1
    data = data[:len(rows)]
    while len(data) < len(rows):
        row = copy.deepcopy(data[-1])
        ObjectLinker.strip_ids(row)
        data[-1].addnext(row)
        data.append(row)
        stats["%s rows added" % label] += 1
    for tr, values in zip(data, rows):
        cells = tr.findall(qn("tc"))
        if len(cells) != len(values):
            stats["%s rows skipped (unexpected cell count)" % label] += 1
            continue
        changed = False
        for tc, value in zip(cells, values):
            if isinstance(value, list):
                changed |= put_lines(tc, value)
            elif value is not None:
                changed |= put(tc, value)
        if changed:
            stats["%s rows written" % label] += 1


def fill_paragraphs(els, texts, stats, label):
    """Make the text paragraphs among `els` read `texts`, in order."""
    paras = [e for e in els if e.tag == qn("p") and para_text(e).strip()
             and not list(e.iter(qn("txbxContent")))]
    if not paras:
        stats["%s skipped (no paragraph to clone)" % label] += 1
        return
    for extra in paras[len(texts):]:
        extra.getparent().remove(extra)
    paras = paras[:len(texts)]
    while len(paras) < len(texts):
        p = copy.deepcopy(paras[-1])
        ObjectLinker.strip_ids(p)
        paras[-1].addnext(p)
        paras.append(p)
    for p, text in zip(paras, texts):
        if " ".join(para_text(p).split()) != text:
            set_par_text(p, text)
            stats["%s paragraphs written" % label] += 1


def block_after(body, style, title):
    """Elements between the heading `title` and the next heading."""
    els = list(body)
    for i, el in enumerate(els):
        if el.tag == qn("p") and style_of(el) == style and para_text(el).strip() == title:
            out = []
            for e in els[i + 1:]:
                if e.tag == qn("p") and style_of(e).startswith("Heading"):
                    break
                out.append(e)
            return out
    return None


def data_table(els):
    """First table among `els` that has rows of its own."""
    return next((e for e in els or [] if e.tag == qn("tbl") and e.findall(qn("tr"))), None)


def apply_meta(body, meta, stats):
    """Write the front matter from `doc-meta.yaml` (see `doc_meta.py`).

    Cover, introduction, abbreviations, connection table and History come
    from the file; the seed only lends their formatting. Returns the date of
    the last history entry.
    """
    last = meta["history"]["entries"][-1]
    doc = meta["document"]
    for p in body.iter(qn("p")):
        style = style_of(p)
        text = " ".join(para_text(p).split())
        want = None
        if style == "MID-PageOverview-TitleBig" and text:
            want = doc["title"]
        elif style == "MID-PageOverView-TitleSmall" and text:
            want = doc["subtitle"]
        elif style == "MID-PageOverview-Small" and text.startswith("Version "):
            want = "Version %s" % cover_version(meta)
        elif style == "MID-PageOverview-Small" and re.search(r"- \d{4}$", text):
            want = "%s - %s" % (doc["place"], last["date"][:4])
        if want is not None and want != text:
            set_par_text(p, want)
            stats["cover lines written"] += 1

    tbl = history_table(body)
    if tbl is None:
        stats["history skipped (no History table)"] += 1
    else:
        fill_rows(tbl, [[e["version"], e["date"], e["changes"], e["author"]]
                        for e in meta["history"]["entries"]], 1, stats, "history")

    intro = block_after(body, "Heading1", "INTRODUCTION")
    if intro is not None:
        fill_paragraphs(intro, meta["introduction"]["overview"], stats, "introduction")
    target = block_after(body, "Heading2", "Target")
    if target is not None:
        fill_paragraphs(target, meta["introduction"]["target"], stats, "target")
    abbr = data_table(block_after(body, "Heading2", "Abbreviation"))
    if abbr is not None:
        fill_rows(abbr, [[a["notation"], a["description"]] for a in meta["abbreviations"]], 1, stats,
                  "abbreviation")
    conn = block_after(body, "Heading2", "System Connection Information")
    if conn is not None:
        fill_paragraphs(conn, [meta["connection"]["intro"]], stats, "connection")
        table = data_table(conn)
        if table is not None:
            fill_rows(table, [[str(i), c["name"], c["value"], c["description"]]
                              for i, c in enumerate(meta["connection"]["parameters"], 1)], 1, stats, "connection")
    return last["date"]


def stamp_modified(core_xml, date):
    """Set the package's last-modified date to the render date."""
    root = etree.fromstring(core_xml)
    el = root.find("{http://purl.org/dc/terms/}modified")
    if el is not None:
        el.text = "%s-%s-%sT00:00:00Z" % (date[:4], date[4:6], date[6:8])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def section_title(parent, name):
    """A section's key in the description file; a Heading 4 names its parent."""
    return "%s > %s" % (parent, name) if parent else name


def seed_descriptions(seed, contract):
    """{category: {section: {table kind: {field: [lines]}}}} of a document's wording.

    Only wording that differs from the spec is returned: the description of a
    field the document shares with the spec record its section pairs with.
    """
    spec = contract["endpoints"]
    pairs, _, _ = pair(spec, extract(seed))
    by_section = {(d["category"], d["parent"], d["name"]): o for o, d in pairs if o["name"] == d["name"]}
    with zipfile.ZipFile(seed) as z:
        root = etree.fromstring(z.read("word/document.xml"), etree.XMLParser(huge_tree=True))
    out = collections.OrderedDict()
    heading = {2: "", 3: "", 4: ""}
    current, key, seen = None, None, set()
    index_enumerations(root.find(qn("body")))
    for el in root.find(qn("body")):
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
        elif tag == "tbl" and current is not None:
            kind = table_kind(el)
            if not kind or kind in seen:
                continue
            seen.add(kind)
            spec_fields = {f["name"]: f for f in current[kind]}
            for tr in el.findall(qn("tr"))[2:]:
                cells = tr.findall(qn("tc"))
                if len(cells) != 5:
                    continue
                name = cell_text(cells[1])
                if name not in spec_fields or (kind == "response_fields" and name in ERROR_ROWS):
                    continue
                lines = cell_lines(cells[4])
                spec_text = " ".join(spec_fields[name]["description"].replace("`", "").split())
                if lines and lines != [spec_text]:
                    (out.setdefault(key[0], collections.OrderedDict())
                        .setdefault(section_title(key[1], key[2]), collections.OrderedDict())
                        .setdefault(kind, collections.OrderedDict()))[name] = lines
    return out


def text_lines(value, where):
    lines = [value] if isinstance(value, str) else value
    if not isinstance(lines, list) or not all(isinstance(x, str) and x.strip() for x in lines):
        raise ValueError("descriptions: %s must be text or a list of lines" % where)
    return [" ".join(x.split()) for x in lines]


def load_descriptions(path):
    """Read the description file: (section wording, object wording)."""
    with open(path, encoding="utf-8") as f:
        data = yaml.load(f, Loader=UniqueKeyLoader) or {}
    sections = {}
    for category, titles in (data.get("sections") or {}).items():
        for title, tables in (titles or {}).items():
            for kind, fields in (tables or {}).items():
                if kind not in TABLE_KINDS.values():
                    raise ValueError("descriptions: unknown table %r in %r" % (kind, title))
                for name, text in (fields or {}).items():
                    sections.setdefault((category, title), {}).setdefault(kind, {})[name] = text_lines(
                        text, "%s / %s / %s" % (title, kind, name))
    objects = {}
    for schema, fields in (data.get("objects") or {}).items():
        for name, text in (fields or {}).items():
            objects.setdefault(schema, {})[name] = text_lines(text, "object %s / %s" % (schema, name))
    return sections, objects


def load_naming(path):
    """Read the object naming file: {schema: {title, chapter}}."""
    with open(path, encoding="utf-8") as f:
        data = yaml.load(f, Loader=UniqueKeyLoader) or {}
    out = {}
    for schema, entry in (data.get("objects") or {}).items():
        entry = entry or {}
        title = str(entry.get("title") or "").strip()
        chapter = str(entry.get("chapter") or ObjectLinker.DEFAULT).strip().upper()
        if not title:
            raise ValueError("objects: %s needs a title" % schema)
        if chapter not in ObjectLinker.CHAPTERS:
            raise ValueError("objects: %s has unknown chapter %r" % (schema, entry.get("chapter")))
        out[schema] = {"title": title, "chapter": chapter}
    return out


def seed_objects(seed, contract):
    """Map a document's object sections to the spec's schemas.

    A Type cell links to an object section; the same field in the spec points
    at a schema. Following those pairs from every section that pairs with a
    spec record, and then through the object tables themselves, gives
    {schema: {title, chapter, wording}}: the section's heading and chapter,
    and its descriptions where they differ from the spec's. An object section
    no such field reaches has no schema and is not returned.
    """
    objects = contract.get("objects") or {}
    pairs, _, _ = pair(contract["endpoints"], extract(seed))
    by_section = {(d["category"], d["parent"], d["name"]): o for o, d in pairs if o["name"] == d["name"]}
    with zipfile.ZipFile(seed) as z:
        root = etree.fromstring(z.read("word/document.xml"), etree.XMLParser(huge_tree=True))
    els = list(root.find(qn("body")))
    index_enumerations(root.find(qn("body")))

    target, table_of, order = {}, {}, []     # anchor -> section; section -> its table
    chapter, pending = None, []
    for i, el in enumerate(els):
        tag = etree.QName(el).localname
        if tag == "bookmarkStart":
            pending.append(el.get(qn("name")))
            continue
        if tag == "p":
            st = style_of(el)
            if st == "Heading1":
                chapter = para_text(el).strip().upper()
            elif st == "Heading2" and chapter in ObjectLinker.CHAPTERS:
                key = (chapter, para_text(el).strip())
                order.append(key)
                for name in pending + [b.get(qn("name")) for b in el.iter(qn("bookmarkStart"))]:
                    target[name] = key
                if i + 1 < len(els) and els[i + 1].tag == qn("tbl"):
                    table_of.setdefault(key, els[i + 1])
        if tag != "bookmarkEnd":
            pending = []

    votes = collections.defaultdict(collections.Counter)

    def scan(tbl, fields):
        spec_fields = {f["name"]: f for f in fields}
        for tr in tbl.findall(qn("tr"))[2:]:
            cells = tr.findall(qn("tc"))
            if len(cells) != 5:
                continue
            f = spec_fields.get(cell_text(cells[1]))
            link = cells[2].find(".//" + qn("hyperlink"))
            if f is None or link is None or not f.get("ref"):
                continue
            key = target.get(link.get(qn("anchor")))
            if key is not None:
                votes[key][f["ref"]] += 1

    heading = {2: "", 3: "", 4: ""}
    current, seen = None, set()
    for el in els:
        tag = etree.QName(el).localname
        if tag == "p":
            m = re.match(r"Heading(\d)$", style_of(el))
            if m and int(m.group(1)) in (2, 3, 4):
                lv = int(m.group(1))
                heading[lv] = para_text(el).strip()
                for k in range(lv + 1, 5):
                    heading[k] = ""
                current = by_section.get((heading[2], heading[3] if lv == 4 else "", heading[lv])) \
                    if lv in (3, 4) else None
                seen = set()
        elif tag == "tbl" and current is not None:
            kind = table_kind(el)
            if kind and kind not in seen:
                seen.add(kind)
                scan(el, current[kind])
    done = set()
    while True:
        todo = [k for k in votes if k not in done and k in table_of]
        if not todo:
            break
        for key in todo:
            done.add(key)
            schema = votes[key].most_common(1)[0][0]
            if schema in objects:
                scan(table_of[key], objects[schema]["fields"])

    out = collections.OrderedDict()
    for key in order:                 # document order; the first section wins a schema
        if key not in votes:
            continue
        schema = votes[key].most_common(1)[0][0]
        if schema in out or schema not in objects:
            continue
        wording = collections.OrderedDict()
        spec_fields = {f["name"]: f for f in objects[schema]["fields"]}
        for tr in (table_of[key].findall(qn("tr"))[2:] if key in table_of else []):
            cells = tr.findall(qn("tc"))
            if len(cells) != 5 or cell_text(cells[1]) not in spec_fields:
                continue
            lines = cell_lines(cells[4])
            spec_text = " ".join(spec_fields[cell_text(cells[1])]["description"].replace("`", "").split())
            if lines and lines != [spec_text]:
                wording[cell_text(cells[1])] = lines
        out[schema] = {"title": key[1], "chapter": key[0], "wording": wording}
    return out


def request_field_update(settings_xml):
    """Ask Word to refresh fields on open, so the table of contents lists new sections."""
    root = etree.fromstring(settings_xml)
    if root.find(qn("updateFields")) is None:
        el = etree.Element(qn("updateFields"))
        el.set(qn("val"), "true")
        root.insert(0, el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def render(seed, contract_path, out, meta=None, descriptions=None, naming=None):
    with open(contract_path, encoding="utf-8") as f:
        contract = json.load(f)
    spec = contract["endpoints"]
    doc = extract(seed)
    pairs, _, _ = pair(spec, doc)
    # A pair whose titles differ came from the broad x-route fallback (e.g. a
    # documented route the backend does not have); rendering it would write
    # another operation's fields into that section, so it is left as is.
    by_section = {(d["category"], d["parent"], d["name"]): o for o, d in pairs if o["name"] == d["name"]}
    # A documented section is "not implemented" only when no route serves its
    # method and path; a second section for a served route is a duplicate, not
    # a missing feature.
    served = {o["key"] for o in spec}
    documented = {(d["category"], d["parent"], d["name"]) for d in doc
                  if d.get("method") and d["key"] not in served}

    with zipfile.ZipFile(seed) as z:
        root = etree.fromstring(z.read("word/document.xml"), etree.XMLParser(huge_tree=True))
    body = root.find(qn("body"))
    index_enumerations(body)
    fallback = fallback_rows(body)

    stats = collections.Counter({k: 0 for k in ("sections paired", "sections rendered", "tables rebuilt", "rows kept",
                            "rows added", "rows removed", "types rewritten", "descriptions kept from seed",
                            "tables using a shared template row", "tables skipped (no template row)")})
    stats["sections paired"] = len(pairs)
    section_wording, object_wording = descriptions if descriptions is not None else (None, None)
    linker = ObjectLinker(root, body, contract.get("objects") or {}, stats, naming, object_wording)
    heading = {2: "", 3: "", 4: ""}
    current, seen = None, set()
    lv_of = [3]

    def wording_for(heading, lv, kind):
        if section_wording is None:
            return None
        title = section_title(heading[3] if lv == 4 else "", heading[lv])
        return section_wording.get((heading[2], title), {}).get(kind, {})

    for el in body:
        tag = etree.QName(el).localname
        if tag == "p":
            m = re.match(r"Heading(\d)$", style_of(el))
            if m and int(m.group(1)) in (2, 3, 4):
                lv = int(m.group(1))
                lv_of[0] = lv
                heading[lv] = para_text(el).strip()
                for k in range(lv + 1, 5):
                    heading[k] = ""
                key = (heading[2], heading[3] if lv == 4 else "", heading[lv])
                current = by_section.get(key) if lv in (3, 4) else None
                seen = set()
                if current is not None:
                    stats["sections rendered"] += 1
                    if current.get("status") == "not-implemented":
                        mark_not_implemented(el, stats)
                elif key in documented:
                    # Documented in the seed but served by no route in the code.
                    mark_not_implemented(el, stats)
                continue
            if current is not None:
                for txbx in el.iter(qn("txbxContent")):
                    kind = frame_kind(txbx)
                    if kind == "request":
                        fix_request_frame(txbx, current, stats)
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
                if rebuild_table(el, kind, current[kind], stats, fallback, linker=linker,
                                 wording=wording_for(heading, lv_of[0], kind)):
                    stats["tables rebuilt"] += 1

    add_new_sections(body, spec, by_section, fallback, stats, IdAllocator(root), linker, find_donors(body),
                     section_wording)
    linker.flush(stats, fallback)
    drop_half_bookmarks(root, stats)
    unlink_dangling(root, stats)
    stamp = apply_meta(body, meta, stats) if meta else None

    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(seed) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                data = xml
            elif item.filename == "docProps/core.xml" and stamp:
                data = stamp_modified(zin.read(item.filename), stamp)
            elif item.filename == "word/settings.xml":
                data = request_field_update(zin.read(item.filename))
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    return stats


def main():
    parser = argparse.ArgumentParser(description="Render the specification DOCX from the OpenAPI contract.")
    parser.add_argument("seed")
    parser.add_argument("contract")
    parser.add_argument("out")
    parser.add_argument("--meta", help="doc-meta.yaml: cover, introduction, abbreviations, connection, history")
    parser.add_argument("--descriptions",
                        help="description file: document wording that overrides the spec's, per section and field")
    parser.add_argument("--objects",
                        help="object naming file: document title and chapter of a schema's object section")
    parser.add_argument("--allow-stale-history", action="store_true",
                        help="render although the API changed since the last history entry")
    args = parser.parse_args()
    if os.path.exists(args.out) and os.path.samefile(args.seed, args.out):
        sys.exit("Output must not be the seed document")
    meta = None
    if args.meta:
        try:
            meta = load_meta(args.meta)
        except (ValueError, yaml.YAMLError) as exc:
            sys.exit("invalid doc-meta: %s" % exc)
        # History must describe the API being rendered: when the spec moved
        # past the snapshot of the last entry, the entry for it is missing.
        last = meta["history"]["entries"][-1]["version"]
        snapshot = read_snapshot(last)
        with open(args.contract, encoding="utf-8") as f:
            current = signatures(json.load(f))
        if snapshot is None and not args.allow_stale_history:
            # Without a snapshot there is nothing to check the API against.
            sys.exit("history %s has no API snapshot; store it with scripts/propose_history.py "
                     "--snapshot, or pass --allow-stale-history" % last)
        elif snapshot is not None and snapshot != current and not args.allow_stale_history:
            sys.exit("the API changed since history %s; add an entry with "
                     "scripts/propose_history.py --write, or pass --allow-stale-history" % last)
    descriptions = None
    if args.descriptions:
        try:
            descriptions = load_descriptions(args.descriptions)
        except (ValueError, yaml.YAMLError) as exc:
            sys.exit("invalid description file: %s" % exc)
    naming = None
    if args.objects:
        try:
            naming = load_naming(args.objects)
        except (ValueError, yaml.YAMLError) as exc:
            sys.exit("invalid object naming file: %s" % exc)
    stats = render(args.seed, args.contract, args.out, meta, descriptions, naming)
    for k, v in stats.items():
        print("%-34s: %d" % (k, v))
    print("wrote %s" % args.out)


if __name__ == "__main__":
    main()
