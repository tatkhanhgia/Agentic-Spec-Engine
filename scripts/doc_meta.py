#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Load and validate the document front-matter file (`doc-meta.yaml`).

The backend keeps `docs/api/doc-meta.yaml` beside `openapi.yaml`. It holds
what the specification document shows around the API and what OpenAPI has no
place for: cover, introduction, abbreviations, connection table and history.

`load()` returns the parsed file or raises `ValueError` naming the first rule
it breaks, so a renderer can refuse before it writes anything:

  * history versions are `major.minor` and strictly increasing;
  * history dates are real `YYYYMMDD` dates and never go back;
  * every entry has an author and at least one change line;
  * an entry newer than `history.limits.applies_after` has at most
    `max_lines` lines of at most `max_line_chars` characters.

THIS MODULE ONLY READS the file.

Usage:
    python scripts/doc_meta.py <doc-meta.yaml>      # validate and summarise
"""
import datetime
import re
import sys

import yaml

VERSION_RE = re.compile(r"(\d+)\.(\d+)$")
DEFAULT_LIMITS = {"max_lines": 6, "max_line_chars": 200, "applies_after": None}


def version_key(version):
    # An unquoted 3.10 is the YAML float 3.1: only a string can be trusted.
    if not isinstance(version, str):
        raise ValueError("a version must be a quoted string such as \"3.10\", got %r" % (version,))
    m = VERSION_RE.match(version)
    if not m:
        raise ValueError("history version must be major.minor, got %r" % (version,))
    return int(m.group(1)), int(m.group(2))


def check_date(date):
    try:
        if not re.match(r"\d{8}$", str(date)):
            raise ValueError
        datetime.datetime.strptime(str(date), "%Y%m%d")
    except ValueError:
        raise ValueError("history date must be a real YYYYMMDD date, got %r" % (date,)) from None
    return str(date)


def text_list(value, where):
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, list) or not value or not all(isinstance(v, str) and v.strip() for v in value):
        raise ValueError("%s must be a non-empty list of non-empty strings" % where)
    return [" ".join(v.split()) for v in value]


def validate_entry(entry, limits, where="history entry"):
    """Check one history entry on its own; returns it normalised."""
    if not isinstance(entry, dict):
        raise ValueError("%s must be a mapping" % where)
    version = entry.get("version", "")
    key = version_key(version)
    where = "history %s" % version
    date = check_date(entry.get("date", ""))
    author = str(entry.get("author") or "").strip()
    if not author:
        raise ValueError("%s: author is required" % where)
    changes = text_list(entry.get("changes"), "%s: changes" % where)
    after = limits.get("applies_after")
    if after is None or key > version_key(after):
        if len(changes) > limits["max_lines"]:
            raise ValueError("%s: %d change lines, the limit is %d" % (where, len(changes), limits["max_lines"]))
        for line in changes:
            if len(line) > limits["max_line_chars"]:
                raise ValueError("%s: a change line has %d characters, the limit is %d"
                                 % (where, len(line), limits["max_line_chars"]))
    return {"version": version, "date": date, "author": author, "changes": changes}


def validate(meta):
    if not isinstance(meta, dict):
        raise ValueError("doc-meta must be a mapping")
    out = {}

    doc = meta.get("document") or {}
    out["document"] = {}
    for key in ("title", "subtitle", "place"):
        value = str(doc.get(key) or "").strip()
        if not value:
            raise ValueError("document.%s is required" % key)
        out["document"][key] = value

    intro = meta.get("introduction") or {}
    out["introduction"] = {
        "overview": text_list(intro.get("overview"), "introduction.overview"),
        "target": text_list(intro.get("target"), "introduction.target"),
    }

    out["abbreviations"] = []
    for i, row in enumerate(meta.get("abbreviations") or [], 1):
        notation = str((row or {}).get("notation") or "").strip()
        description = str((row or {}).get("description") or "").strip()
        if not notation or not description:
            raise ValueError("abbreviations[%d] needs notation and description" % i)
        out["abbreviations"].append({"notation": notation, "description": description})
    if not out["abbreviations"]:
        raise ValueError("abbreviations must not be empty")

    conn = meta.get("connection") or {}
    out["connection"] = {"intro": str(conn.get("intro") or "").strip(), "parameters": []}
    if not out["connection"]["intro"]:
        raise ValueError("connection.intro is required")
    for i, row in enumerate(conn.get("parameters") or [], 1):
        item = {k: str((row or {}).get(k) or "").strip() for k in ("name", "value", "description")}
        if not all(item.values()):
            raise ValueError("connection.parameters[%d] needs name, value and description" % i)
        out["connection"]["parameters"].append(item)
    if not out["connection"]["parameters"]:
        raise ValueError("connection.parameters must not be empty")

    history = meta.get("history") or {}
    limits = dict(DEFAULT_LIMITS)
    limits.update(history.get("limits") or {})
    for key in ("max_lines", "max_line_chars"):
        if not isinstance(limits[key], int) or limits[key] < 1:
            raise ValueError("history.limits.%s must be a positive integer" % key)
    if limits["applies_after"] is not None:
        version_key(limits["applies_after"])
    entries = history.get("entries")
    if not isinstance(entries, list) or not entries:
        raise ValueError("history.entries must not be empty")
    out["history"] = {"limits": limits, "entries": []}
    previous = None
    for entry in entries:
        item = validate_entry(entry, limits)
        if previous is not None:
            if version_key(item["version"]) <= version_key(previous["version"]):
                raise ValueError("history %s must be greater than %s" % (item["version"], previous["version"]))
            if item["date"] < previous["date"]:
                raise ValueError("history %s: date %s is before %s of %s"
                                 % (item["version"], item["date"], previous["date"], previous["version"]))
        out["history"]["entries"].append(item)
        previous = item
    return out


class UniqueKeyLoader(yaml.SafeLoader):
    """SafeLoader that refuses a repeated mapping key instead of keeping the last one."""

    def construct_mapping(self, node, deep=False):
        seen = set()
        for key_node, _ in node.value:
            key = self.construct_object(key_node, deep=deep)
            if key in seen:
                raise ValueError("duplicate key %r at line %d" % (key, key_node.start_mark.line + 1))
            seen.add(key)
        return super().construct_mapping(node, deep=deep)


def load(path):
    with open(path, encoding="utf-8") as f:
        return validate(yaml.load(f, Loader=UniqueKeyLoader))


def cover_version(meta):
    last = meta["history"]["entries"][-1]
    return "%s.%s" % (last["version"], last["date"][2:])


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    try:
        meta = load(sys.argv[1])
    except (ValueError, yaml.YAMLError) as exc:
        sys.exit("invalid doc-meta: %s" % exc)
    last = meta["history"]["entries"][-1]
    print("history entries : %d" % len(meta["history"]["entries"]))
    print("last entry      : %s (%s, %s)" % (last["version"], last["date"], last["author"]))
    print("cover version   : %s" % cover_version(meta))


if __name__ == "__main__":
    main()
