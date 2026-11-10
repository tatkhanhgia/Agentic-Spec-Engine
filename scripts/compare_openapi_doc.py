#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Compare the OpenAPI-derived contract with the DOCX-derived contract.

Inputs are the outputs of `openapi_to_contract.py` and `extract_doc_contract.py`.
Endpoints are paired by method plus the operation's `x-route` regex (the same
pattern the backend dispatches on), falling back to the normalised key. Pairing
by regex matters: the spec declares one templated operation for
`/fields/{fieldType}` while the document gives each field type its own section.

The shared error rows every document response table ends with
(`error_description`, `error_code`, `message`, presence `M/O`) are a rendering
convention, not per-endpoint content, so they are left out of the field diff.

THIS SCRIPT ONLY READS its two inputs and writes a Markdown report.

Usage:
    python scripts/compare_openapi_doc.py <from_openapi.json> <from_doc.json> [-o report.md]
"""
import collections
import json
import re
import sys

TABLES = ("headers", "path_params", "request_fields", "response_fields")
ERROR_ROWS = {"error", "error_description", "error_code", "message", "remark"}


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)["endpoints"]


def doc_path(e):
    p = re.sub(r"^\{baseURL\}", "", e["path"]).split("?")[0]
    return p


def pair(openapi, doc):
    by_key = {}
    for o in openapi:
        by_key.setdefault(o["key"], o)
    pairs, doc_only = [], []
    for d in doc:
        match = by_key.get(d["key"])
        if match is None:
            concrete = re.sub(r"\{[^}]*\}", "1", doc_path(d))
            for o in openapi:
                if o["method"] == d["method"] and o.get("route") and \
                        re.fullmatch(o["route"], concrete):
                    match = o
                    break
        if match is None:
            doc_only.append(d)
        else:
            pairs.append((match, d))
    matched = {id(o) for o, _ in pairs}
    openapi_only = [o for o in openapi if id(o) not in matched]
    return pairs, doc_only, openapi_only


def rows(e, table):
    return {f["name"]: f for f in e[table] if not (table == "response_fields" and f["name"] in ERROR_ROWS)}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    out = None
    if "-o" in sys.argv:
        out = sys.argv[sys.argv.index("-o") + 1]
        args = [a for a in args if a != out]
    if len(args) != 2:
        print(__doc__)
        sys.exit(2)
    openapi, doc = load(args[0]), load(args[1])
    pairs, doc_only, openapi_only = pair(openapi, doc)

    stats = collections.Counter()
    per_endpoint = []
    for o, d in pairs:
        issues = []
        if o["name"] != d["name"]:
            stats["name differs"] += 1
            issues.append("name: doc `%s` / spec `%s`" % (d["name"], o["name"]))
        if o["category"] != d["category"]:
            stats["category differs"] += 1
            issues.append("category: doc `%s` / spec `%s`" % (d["category"], o["category"]))
        for t in TABLES:
            a, b = rows(o, t), rows(d, t)
            stats[t + ": rows in doc"] += len(b)
            stats[t + ": rows in spec"] += len(a)
            only_doc, only_spec = sorted(b.keys() - a.keys()), sorted(a.keys() - b.keys())
            stats[t + ": only in doc"] += len(only_doc)
            stats[t + ": only in spec"] += len(only_spec)
            if only_doc:
                issues.append("%s only in doc: %s" % (t, ", ".join(only_doc)))
            if only_spec:
                issues.append("%s only in spec: %s" % (t, ", ".join(only_spec)))
            for n in sorted(a.keys() & b.keys()):
                stats[t + ": common"] += 1
                if a[n]["type"] != b[n]["type"]:
                    stats[t + ": type differs"] += 1
                if a[n]["presence"] != b[n]["presence"]:
                    stats[t + ": presence differs"] += 1
                    if b[n]["presence"].startswith("RC"):
                        stats[t + ": doc presence is conditional (RC)"] += 1
                if not a[n]["description"]:
                    stats[t + ": spec description empty"] += 1
        per_endpoint.append((d, o, issues))

    lines = ["# OpenAPI vs DOCX contract comparison", ""]
    lines.append("| Metric | Value |")
    lines.append("|---|---|")
    lines.append("| Spec operations | %d |" % len(openapi))
    lines.append("| Doc endpoints | %d |" % len(doc))
    lines.append("| Paired | %d |" % len(pairs))
    lines.append("| Doc only (no spec operation) | %d |" % len(doc_only))
    lines.append("| Spec only (not in doc) | %d |" % len(openapi_only))
    lines.append("| Paired with no field difference | %d |" % sum(1 for _, _, i in per_endpoint if not i))
    for k in sorted(stats):
        lines.append("| %s | %d |" % (k, stats[k]))
    lines += ["", "## Doc only", ""]
    lines += ["- `%s %s` — %s / %s" % (d["method"], doc_path(d), d["category"], d["name"]) for d in doc_only]
    lines += ["", "## Spec only", ""]
    lines += ["- `%s %s` — %s / %s" % (o["method"], doc_path(o), o["category"], o["name"]) for o in openapi_only]
    lines += ["", "## Paired endpoints with differences", ""]
    for d, o, issues in per_endpoint:
        if issues:
            lines.append("### %s — `%s %s`" % (d["name"], d["method"], doc_path(d)))
            lines += ["- " + i for i in issues]
            lines.append("")
    text = "\n".join(lines) + "\n"
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(text)
        print("wrote %s" % out)
    else:
        print(text)
    print("paired %d, doc only %d, spec only %d" % (len(pairs), len(doc_only), len(openapi_only)))


if __name__ == "__main__":
    main()
