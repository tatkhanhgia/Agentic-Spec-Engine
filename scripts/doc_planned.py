#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Load and validate the planned-sections file (`doc-planned.yaml`).

`openapi.yaml` describes only what the code serves. The specification
document also carries sections for APIs that are designed but have no route
in the code yet. They live in `config/doc-planned.yaml`, and the renderer
shows each one marked `(Not implemented yet)`:

    sections:
      - category: <Heading 2 the section belongs to>
        title: <section title>
        after: <title of the section it follows in that category>   # optional
        first: true                # optional: the section opens its category
        method: GET | POST | PUT | DELETE | PATCH
        path: /v1/...
        description: one line, or a list of lines
        headers | path_params | request_fields | response_headers | response_fields:
          - {name, type, presence, description}
        request_example | response_example: a JSON value, or the sample as text

`merge()` adds the sections to the contract's endpoints. It refuses a section
whose method and path the spec now serves, or whose title the spec now uses
in the same category: that API exists, so its entry here must be deleted.

THIS MODULE ONLY READS the file.

Usage:
    python scripts/doc_planned.py <doc-planned.yaml> [<from_openapi.json>]   # validate and summarise
"""
import json
import re
import sys

import yaml

from doc_meta import UniqueKeyLoader

METHODS = ("GET", "POST", "PUT", "DELETE", "PATCH")
TABLES = ("headers", "path_params", "request_fields", "response_headers", "response_fields")
KEYS = {"category", "title", "after", "first", "method", "path", "description",
        "request_example", "response_example"} | set(TABLES)


def to_key(method, path):
    """`METHOD /path` with every path parameter written `{}`."""
    return "%s %s" % (method, re.sub(r"\{[^}]*\}", "{}", path.split("?")[0]))


def lines(value, where):
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, list) or not value or not all(isinstance(v, str) and v.strip() for v in value):
        raise ValueError("%s must be a line or a non-empty list of lines" % where)
    return [" ".join(v.split()) for v in value]


def validate(data):
    if not isinstance(data, dict) or not isinstance(data.get("sections"), list):
        raise ValueError("doc-planned must be a mapping with a `sections` list")
    out, seen = [], set()
    for i, raw in enumerate(data["sections"], 1):
        if not isinstance(raw, dict):
            raise ValueError("sections[%d] must be a mapping" % i)
        where = "sections[%d] (%s)" % (i, raw.get("title", "?"))
        unknown = set(raw) - KEYS
        if unknown:
            raise ValueError("%s: unknown key %s" % (where, ", ".join(sorted(unknown))))
        item = {}
        for key in ("category", "title", "method", "path"):
            value = raw.get(key)
            if not isinstance(value, str) or not value.strip():
                raise ValueError("%s: %s is required" % (where, key))
            item[key] = value.strip()
        if item["method"] not in METHODS:
            raise ValueError("%s: method must be one of %s" % (where, ", ".join(METHODS)))
        if not item["path"].startswith("/"):
            raise ValueError("%s: path must start with /" % where)
        ident = (item["category"], item["title"])
        if ident in seen:
            raise ValueError("%s: the title is used twice in %s" % (where, item["category"]))
        seen.add(ident)
        item["after"] = str(raw.get("after") or "").strip()
        if raw.get("first") not in (None, True, False) or (raw.get("first") and item["after"]):
            raise ValueError("%s: first is true or false, and never set together with after" % where)
        item["first"] = bool(raw.get("first"))
        item["description"] = lines(raw.get("description"), "%s: description" % where)
        for table in TABLES:
            rows = raw.get(table) or []
            if not isinstance(rows, list):
                raise ValueError("%s: %s must be a list" % (where, table))
            item[table] = []
            for n, row in enumerate(rows, 1):
                if not isinstance(row, dict):
                    raise ValueError("%s: %s[%d] must be a mapping" % (where, table, n))
                field = {k: str(row.get(k) if row.get(k) is not None else "").strip()
                         for k in ("name", "type", "presence", "description")}
                if not field["name"]:
                    raise ValueError("%s: %s[%d] needs a name" % (where, table, n))
                item[table].append(field)
        for key in ("request_example", "response_example"):
            value = raw.get(key)
            if value is not None and not isinstance(value, (dict, list, str)):
                raise ValueError("%s: %s must be a JSON value or text" % (where, key))
            item[key] = value
        out.append(item)
    return out


def load(path):
    with open(path, encoding="utf-8") as f:
        return validate(yaml.load(f, Loader=UniqueKeyLoader))


def as_record(section, order):
    """The contract record a planned section renders from."""
    has_body = any(h["name"].lower() == "content-type" for h in section["headers"])
    record = {
        "name": section["title"],
        "category": section["category"],
        "category_order": order,
        "description": section["description"][0],
        "prose": section["description"],
        "method": section["method"],
        "path": "{baseURL}" + section["path"],
        "key": "%s %s" % (section["method"], section["path"].split("?")[0]),
        "status": "not-implemented",
        "planned": True,
        "success_status": "200",
        "request_example": section["request_example"],
        "response_example": section["response_example"],
        "media_types": ["application/json"] if has_body else [],
        "response_media_types": None,
    }
    for table in TABLES:
        record[table] = [dict(f) for f in section[table]]
    return record


def merge(endpoints, sections):
    """`endpoints` with the planned sections placed in their categories."""
    served = {to_key(e["method"], re.sub(r"^\{baseURL\}", "", e["path"])) for e in endpoints}
    titles = {(e["category"], e["name"]) for e in endpoints}
    for s in sections:
        if to_key(s["method"], s["path"]) in served:
            raise ValueError("planned section %r: the spec now serves %s %s; delete the entry"
                             % (s["title"], s["method"], s["path"]))
        if (s["category"], s["title"]) in titles:
            raise ValueError("planned section %r: the spec now has a section of that title in %s; "
                             "delete the entry" % (s["title"], s["category"]))
    order = {}
    for e in endpoints:
        order.setdefault(e["category"], e.get("category_order", 0))
    out = list(endpoints)
    for s in sections:
        if s["category"] not in order:
            order[s["category"]] = max(order.values(), default=-1) + 1
        record = as_record(s, order[s["category"]])
        same = [i for i, e in enumerate(out) if e["category"] == s["category"]]
        anchor = [i for i in same if out[i]["name"] == s["after"]] if s["after"] else []
        if s["first"] and same:
            at = same[0]
        else:
            at = (anchor[-1] if anchor else same[-1] if same else len(out) - 1) + 1
        out.insert(at, record)
    return out


def main():
    if len(sys.argv) not in (2, 3):
        print(__doc__)
        sys.exit(2)
    try:
        sections = load(sys.argv[1])
        if len(sys.argv) == 3:
            with open(sys.argv[2], encoding="utf-8") as f:
                merge(json.load(f)["endpoints"], sections)
    except (ValueError, yaml.YAMLError) as exc:
        sys.exit("invalid doc-planned: %s" % exc)
    print("planned sections : %d" % len(sections))
    for category in dict.fromkeys(s["category"] for s in sections):
        print("  %-32s %d" % (category, sum(1 for s in sections if s["category"] == category)))


if __name__ == "__main__":
    main()
