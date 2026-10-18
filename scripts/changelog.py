#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Generate CHANGELOG.md entries automatically by comparing document snapshots.

The specification is a binary DOCX, so `git diff` says nothing useful and a
hand-written changelog records what an agent *believes* it changed. That belief
can be wrong: on 2026-07-27 a verification step reported "no change" while a
whole section had in fact been removed by another agent. Everything here is
therefore derived from the files themselves - nothing is typed from memory.

A snapshot captures three layers, because each is blind to the others:

  * **endpoints** - method, path, headers, request/response fields, taken with
    the same parser as `extract_doc_contract.py`;
  * **section text digests** - one hash per endpoint section, so prose and table
    edits that leave the endpoint shape intact are still visible;
  * **document stats** - paragraph / table / frame counts, which catch bulk
    structural change such as a deleted section.

Read-only with respect to the DOCX. It writes only snapshots and CHANGELOG.md,
so it is safe to run while another agent is editing.

Usage:
    python scripts/changelog.py snapshot [docx] [-o out.json]
    python scripts/changelog.py diff <old> <new>          # old/new: .json or .docx
    python scripts/changelog.py session <old> [new] [--title "..."] [--write]

`session --write` prepends a dated entry to CHANGELOG.md; without it the entry
is printed so it can be reviewed first.
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract_doc_contract import extract as extract_endpoints

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX_DEFAULT = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
CHANGELOG = os.path.join(ROOT, "CHANGELOG.md")
HISTORY = os.path.join(ROOT, "contract", "history")


def qn(tag):
    return "{%s}%s" % (W, tag)


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def section_digests(root):
    """sha1 of the full text of each Heading2/3/4 section."""
    body = root.find(qn("body"))
    digests = {}
    current = None
    buf = []

    def flush():
        if current:
            digests[current] = hashlib.sha1("\n".join(buf).encode("utf-8")).hexdigest()[:12]

    heading2 = ""
    for el in body:
        if el.tag == qn("p") and re.match(r"Heading[234]$", style_of(el)):
            flush()
            title = text_of(el).strip()
            lvl = int(style_of(el)[-1])
            if lvl == 2:
                heading2 = title
                current = title
            else:
                current = "%s / %s" % (heading2, title)
            buf = []
        else:
            buf.append(text_of(el))
    flush()
    return digests


def snapshot(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    eps = extract_endpoints(docx_path)
    return {
        "docx": os.path.basename(docx_path),
        "stats": {
            "paragraphs": len(list(root.iter(qn("p")))),
            "tables": len(list(root.iter(qn("tbl")))),
            "frames": len(list(root.iter(qn("txbxContent")))),
            "endpoints": len(eps),
        },
        "sections": section_digests(root),
        "endpoints": {
            "%s %s" % (e["method"], re.sub(r"^\{baseURL\}", "", e["path"])): {
                "name": e["name"],
                "category": e["category"],
                "headers": [h["name"] for h in e["headers"]],
                "request_fields": [(f["name"], f["presence"]) for f in e["request_fields"]],
                "response_fields": [f["name"] for f in e["response_fields"]],
            }
            for e in eps if e["method"]
        },
    }


def load_snapshot(path):
    if path.lower().endswith(".docx"):
        return snapshot(path)
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


# --------------------------------------------------------------------------
# diffing
# --------------------------------------------------------------------------
def diff(old, new):
    o, n = old["endpoints"], new["endpoints"]
    o_by_name = {v["name"]: k for k, v in o.items()}
    n_by_name = {v["name"]: k for k, v in n.items()}

    added, removed, moved, changed = [], [], [], []

    for key in sorted(set(n) - set(o)):
        name = n[key]["name"]
        if name in o_by_name:                       # same endpoint, new method/path
            moved.append((name, o_by_name[name], key))
        else:
            added.append((key, n[key]))
    for key in sorted(set(o) - set(n)):
        name = o[key]["name"]
        if name not in n_by_name:
            removed.append((key, o[key]))

    for key in sorted(set(o) & set(n)):
        deltas = []
        for field, label in (("headers", "header"), ("response_fields", "response field")):
            plus = [x for x in n[key][field] if x not in o[key][field]]
            minus = [x for x in o[key][field] if x not in n[key][field]]
            for x in plus:
                deltas.append("+%s %s" % (label, x))
            for x in minus:
                deltas.append("-%s %s" % (label, x))
        ofr = dict(o[key]["request_fields"])
        nfr = dict(n[key]["request_fields"])
        for name in nfr:
            if name not in ofr:
                deltas.append("+request field %s" % name)
            elif ofr[name] != nfr[name]:
                deltas.append("request field %s: %s -> %s" % (name, ofr[name], nfr[name]))
        for name in ofr:
            if name not in nfr:
                deltas.append("-request field %s" % name)
        if deltas:
            changed.append((key, n[key]["name"], deltas))

    os_, ns_ = old["sections"], new["sections"]
    text_changed = sorted(k for k in set(os_) & set(ns_) if os_[k] != ns_[k])
    return added, removed, moved, changed, text_changed


def render(old, new, title=None):
    added, removed, moved, changed, text_changed = diff(old, new)
    today = datetime.date.today().isoformat()
    out = ["## %s%s" % (today, " — %s" % title if title else ""), ""]

    so, sn = old["stats"], new["stats"]
    out.append("| | trước | sau |")
    out.append("|---|---|---|")
    for k, label in (("endpoints", "Endpoint"), ("paragraphs", "Đoạn văn"),
                     ("tables", "Bảng"), ("frames", "Khung mẫu")):
        mark = "" if so[k] == sn[k] else "  **%+d**" % (sn[k] - so[k])
        out.append("| %s | %d | %d%s |" % (label, so[k], sn[k], mark))
    out.append("")

    if added:
        out.append("### Đã thêm")
        for key, v in added:
            out.append("- `%s` — %s *(%s)*" % (key, v["name"], v["category"]))
        out.append("")
    if removed:
        out.append("### Đã gỡ")
        for key, v in removed:
            out.append("- `%s` — %s *(%s)*" % (key, v["name"], v["category"]))
        out.append("")
    if moved:
        out.append("### Đổi method / path")
        for name, before, after in moved:
            out.append("- %s: `%s` → `%s`" % (name, before, after))
        out.append("")
    if changed:
        # A document-wide change (the error response reshape, say) produces the
        # same delta on ~185 endpoints. Listing each one buries the one-off
        # changes that actually need reading, so identical deltas collapse.
        groups = {}
        for key, name, deltas in changed:
            groups.setdefault(tuple(sorted(deltas)), []).append(name)
        out.append("### Đổi thuộc tính")
        for deltas, names in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            what = ", ".join(deltas)
            if len(names) == 1:
                out.append("- %s: %s" % (names[0], what))
            else:
                sample = ", ".join(names[:3])
                more = "" if len(names) <= 3 else ", … (+%d)" % (len(names) - 3)
                out.append("- **%d endpoint**: %s" % (len(names), what))
                out.append("  - %s%s" % (sample, more))
        out.append("")
    if text_changed:
        out.append("### Đổi nội dung chữ (%d mục)" % len(text_changed))
        shown = text_changed[:25]
        for k in shown:
            out.append("- %s" % k)
        if len(text_changed) > len(shown):
            out.append("- … và %d mục khác" % (len(text_changed) - len(shown)))
        out.append("")

    if not any((added, removed, moved, changed, text_changed)):
        out.append("*Không có thay đổi nào.*")
        out.append("")
    return "\n".join(out)


def prepend_entry(entry):
    header = "# CHANGELOG\n\nSinh tự động bởi `scripts/changelog.py` bằng cách so sánh\nhai bản chụp tài liệu. Không viết tay.\n\n"
    body = ""
    if os.path.isfile(CHANGELOG):
        existing = open(CHANGELOG, encoding="utf-8").read()
        body = existing.split("\n", 4)[-1] if existing.startswith("# CHANGELOG") else existing
    with open(CHANGELOG, "w", encoding="utf-8") as fh:
        fh.write(header + entry + "\n" + body.lstrip("\n"))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("snapshot")
    s.add_argument("docx", nargs="?", default=DOCX_DEFAULT)
    s.add_argument("-o")

    d = sub.add_parser("diff")
    d.add_argument("old")
    d.add_argument("new", nargs="?", default=DOCX_DEFAULT)

    e = sub.add_parser("session")
    e.add_argument("old")
    e.add_argument("new", nargs="?", default=DOCX_DEFAULT)
    e.add_argument("--title")
    e.add_argument("--write", action="store_true")

    a = ap.parse_args()

    if a.cmd == "snapshot":
        snap = snapshot(a.docx)
        out = a.o
        if not out:
            os.makedirs(HISTORY, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")
            out = os.path.join(HISTORY, "%s.json" % stamp)
        with open(out, "w", encoding="utf-8") as fh:
            json.dump(snap, fh, ensure_ascii=False, indent=2)
        print("wrote %s" % out)
        print("  " + ", ".join("%s=%d" % kv for kv in snap["stats"].items()))
        return

    old = load_snapshot(a.old)
    new = load_snapshot(a.new)
    entry = render(old, new, getattr(a, "title", None))
    if a.cmd == "session" and a.write:
        prepend_entry(entry)
        print("prepended entry to %s" % CHANGELOG)
    else:
        print(entry)


if __name__ == "__main__":
    main()
