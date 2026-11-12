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


def rebuild_table(tbl, kind, fields, stats, fallback, fresh=False, linker=None):
    """Rebuild a table's data rows from `fields`.

    `fresh` is for a section cloned from another endpoint: its rows only lend
    their formatting, so no row keeps the template's name, type or wording.
    `linker` links a Type cell naming a component schema to its object section.
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
            if not has_link(cells[2]) and cell_text(cells[2]) != f["type"]:
                set_cell_text(cells[2], f["type"])
                stats["types rewritten"] += 1
        if linker is not None and f.get("ref") and not has_link(cells[2]):
            linker.link(cells[2], f, stats)
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


def body_kind(record):
    return bool(record["request_fields"]) or has_content_type(record)


def build_section(template_els, record, ids, fallback, stats, linker=None):
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

    seen = set()
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
                lines = txbx.findall(qn("p"))
                first = para_text(lines[0]).strip() if lines else ""
                if frame_kind(txbx) is None and HTTP_CODE_RE.match(first):
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
                    rewrite_frame(txbx, record["response_fields"], stats, fresh=True)
        elif tag == "tbl":
            trs = e.findall(qn("tr"))
            title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip().lower() if trs else ""
            kind = TABLE_KINDS.get(title)
            if not kind or kind in seen:
                continue
            seen.add(kind)
            if not record[kind]:
                drop.append(e)
                continue
            rebuild_table(e, kind, record[kind], stats, fallback, fresh=True, linker=linker)
            # Some seed tables carry cell widths off their grid (TECHNICAL_DEBT);
            # a new section must not copy the defect, so every unmerged row
            # takes the grid widths.
            grid = [gc.get(qn("w")) for gc in e.findall(qn("tblGrid") + "/" + qn("gridCol"))]
            for tr in e.findall(qn("tr")):
                cells = tr.findall(qn("tc"))
                if len(cells) != len(grid):
                    continue
                for tc, width in zip(cells, grid):
                    tcw = tc.find(qn("tcPr") + "/" + qn("tcW"))
                    if tcw is not None and width:
                        tcw.set(qn("w"), width)
                        tcw.set(qn("type"), "dxa")
    for kind in TABLE_KINDS.values():
        if record[kind] and kind not in seen:
            stats["new sections missing a %s table" % kind] += 1
    return [e for e in els if e not in drop]


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
    """Links composite Type cells to object sections in the RELATIVE OBJECT chapter.

    A Type cell that names a component schema becomes a `w:hyperlink` whose
    anchor is a bookmark on the object's Heading 2 (the document's convention).
    When the seed already links the same label to exactly one existing object,
    that anchor is reused; otherwise an object section is created from the
    spec at the end of the chapter, cloned from an existing object section.
    """

    CHAPTER = "RELATIVE OBJECT"

    def __init__(self, root, body, objects):
        self.body = body
        self.objects = objects
        self.created = collections.OrderedDict()   # schema name -> anchor
        self.queue = []
        self.marks = {b.get(qn("name")) for b in root.iter(qn("bookmarkStart"))}
        ids = [int(b.get(qn("id"))) for b in root.iter(qn("bookmarkStart"))
               if (b.get(qn("id")) or "").isdigit()]
        self.next_id = max(ids or [0]) + 1
        # Labels the seed already links, by the name inside List<...>.
        seen = collections.defaultdict(set)
        for tr in root.iter(qn("tr")):
            cells = tr.findall(qn("tc"))
            if len(cells) != 5:
                continue
            link = cells[2].find(".//" + qn("hyperlink"))
            if link is not None and link.get(qn("anchor")) in self.marks:
                seen[self.inner(cell_text(cells[2]))].add(link.get(qn("anchor")))
        self.seed_anchor = {label: next(iter(a)) for label, a in seen.items() if len(a) == 1}
        self.can_create = self.template() is not None and self.chapter_end() is not None

    @staticmethod
    def inner(label):
        m = re.match(r"\s*List<(.*)>\s*$", label)
        return (m.group(1) if m else label).strip()

    def anchor_for(self, ref, label, stats):
        inner = self.inner(label)
        if inner in self.seed_anchor:
            stats["type links to existing objects"] += 1
            return self.seed_anchor[inner]
        # No link without a destination: a new object section needs a template
        # to clone and a chapter to go into.
        if ref not in self.objects or not self.can_create:
            return None
        if ref not in self.created:
            anchor = "_OBJ_%s" % re.sub(r"[^A-Za-z0-9_]", "_", ref)
            while anchor in self.marks:
                anchor += "_"
            self.marks.add(anchor)
            self.created[ref] = anchor
            self.queue.append(ref)
        stats["type links to new objects"] += 1
        return self.created[ref]

    def link(self, tc, field, stats):
        """Turn a plain Type cell into `List<` + hyperlink + `>` (or a bare hyperlink)."""
        label = field["type"]
        anchor = self.anchor_for(field["ref"], label, stats)
        if anchor is None:
            return
        set_cell_text(tc, label)
        p = tc.find(qn("p"))
        run = p.find(qn("r"))
        is_list = label.strip().startswith("List<")

        def piece(text):
            r = copy.deepcopy(run)
            r.find(qn("t")).text = text
            return r

        link = etree.Element(qn("hyperlink"))
        link.set(qn("anchor"), anchor)
        link.set(qn("history"), "1")
        linked = piece(self.inner(label))
        rpr = linked.find(qn("rPr"))
        if rpr is None:
            rpr = etree.Element(qn("rPr"))
            linked.insert(0, rpr)
        # The seed's Type links use the Hyperlink character style in black.
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
        """(heading, table, spacer) of an existing object section with a plain data row."""
        els = list(self.body)
        inside = False
        for i, el in enumerate(els):
            if etree.QName(el).localname != "p":
                continue
            st = style_of(el)
            if st == "Heading1":
                inside = para_text(el).strip().upper() == self.CHAPTER
            elif inside and st == "Heading2" and i + 1 < len(els) and els[i + 1].tag == qn("tbl"):
                rows = els[i + 1].findall(qn("tr"))
                if len(rows) > 2 and any(plain_row(r) is not None for r in rows[2:]):
                    spacer = els[i + 2] if i + 2 < len(els) and els[i + 2].tag == qn("p") else None
                    if spacer is not None and (style_of(spacer) or para_text(spacer).strip()):
                        spacer = None
                    return el, els[i + 1], spacer
        return None

    def chapter_end(self):
        els = list(self.body)
        inside = False
        for i, el in enumerate(els):
            if etree.QName(el).localname == "p" and style_of(el) == "Heading1":
                if inside:
                    return i
                inside = para_text(el).strip().upper() == self.CHAPTER
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
        found = self.template()
        at = self.chapter_end()
        if found is None or at is None:
            stats["object sections skipped (no template or chapter)"] += len(self.queue)
            return
        t_heading, t_table, t_spacer = found
        while self.queue:
            ref = self.queue.pop(0)
            obj = self.objects[ref]
            heading = copy.deepcopy(t_heading)
            for b in list(heading.iter(qn("bookmarkStart"))) + list(heading.iter(qn("bookmarkEnd"))):
                b.getparent().remove(b)
            self.strip_ids(heading)
            set_par_text(heading, obj["label"])
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
            set_cell_text(table.find(qn("tr")).find(qn("tc")), obj["label"])
            rebuild_table(table, "object", obj["fields"], stats, fallback, fresh=True, linker=self)
            grid = [gc.get(qn("w")) for gc in table.findall(qn("tblGrid") + "/" + qn("gridCol"))]
            for tr in table.findall(qn("tr")):
                cells = tr.findall(qn("tc"))
                if len(cells) != len(grid):
                    continue
                for tc, width in zip(cells, grid):
                    tcw = tc.find(qn("tcPr") + "/" + qn("tcW"))
                    if tcw is not None and width:
                        tcw.set(qn("w"), width)
                        tcw.set(qn("type"), "dxa")
            new = [heading, table]
            if t_spacer is not None:
                spacer = copy.deepcopy(t_spacer)
                self.strip_ids(spacer)
                new.append(spacer)
            for el in new:
                self.body.insert(at, el)
                at += 1
            stats["object sections added"] += 1


def add_new_sections(body, spec, by_section, fallback, stats, ids, linker=None):
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
        for offset, el in enumerate(build_section(template["els"], record, ids, fallback, stats, linker)):
            body.insert(at + offset, el)
        stats["sections added"] += 1


def request_field_update(settings_xml):
    """Ask Word to refresh fields on open, so the table of contents lists new sections."""
    root = etree.fromstring(settings_xml)
    if root.find(qn("updateFields")) is None:
        el = etree.Element(qn("updateFields"))
        el.set(qn("val"), "true")
        root.insert(0, el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def render(seed, contract_path, out):
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
    fallback = fallback_rows(body)
    linker = ObjectLinker(root, body, contract.get("objects") or {})

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
                if rebuild_table(el, kind, current[kind], stats, fallback, linker=linker):
                    stats["tables rebuilt"] += 1

    add_new_sections(body, spec, by_section, fallback, stats, IdAllocator(root), linker)
    linker.flush(stats, fallback)

    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(seed) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                data = xml
            elif item.filename == "word/settings.xml":
                data = request_field_update(zin.read(item.filename))
            else:
                data = zin.read(item.filename)
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
