#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Propose the next History entry of `doc-meta.yaml` from what changed in the API.

Each history version has a snapshot of the API it documents, kept in
`contract/openapi_history/<version>.json` (one signature per operation: its
name, category, and the name, type and presence of every field). This script compares the current
`contract/from_openapi.json` with the snapshot of the last history entry:

  * no difference  -> nothing to add;
  * a difference   -> the next entry: last version + 0.1, today's date, the
                      last entry's author, and change lines naming the APIs
                      added, removed and updated, cut to the limits of
                      `history.limits`.

Without `--write` it only prints the proposal. With `--write` it appends the
entry to `doc-meta.yaml` (the file's text is otherwise untouched) and stores
the new snapshot. Edit the wording in `doc-meta.yaml` afterwards if needed.
`--snapshot` stores the current API as the snapshot of the last entry, for a
version that has none yet.

Usage:
    python scripts/propose_history.py <from_openapi.json> <doc-meta.yaml>
        [--write] [--snapshot] [--date YYYYMMDD] [--author NAME]
"""
import argparse
import datetime
import json
import os
import sys

import yaml

from doc_meta import load as load_meta, validate_entry, version_key

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SNAPSHOT_DIR = os.path.join(ROOT, "contract", "openapi_history")
TABLES = ("headers", "path_params", "request_fields", "response_fields", "response_headers")


def signatures(contract):
    """{operation id: signature}; the id survives a renamed title."""
    out = {}
    for e in contract["endpoints"]:
        ident = "%s %s" % (e["key"], e.get("variant") or e["name"])
        out[ident] = {
            "name": e["name"],
            "category": e["category"],
            # Lists, not tuples: a snapshot read back from JSON holds lists.
            "fields": {t: sorted([f["name"], f["type"], f["presence"]] for f in e[t]) for t in TABLES},
            "status": e.get("status", ""),
        }
    return out


def snapshot_path(version):
    return os.path.join(SNAPSHOT_DIR, "%s.json" % version)


def read_snapshot(version):
    path = snapshot_path(version)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def write_snapshot(version, sigs):
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    with open(snapshot_path(version), "w", encoding="utf-8") as f:
        json.dump(sigs, f, ensure_ascii=False, indent=1, sort_keys=True)
        f.write("\n")


def diff(old, new):
    added = [new[k]["name"] for k in new if k not in old]
    removed = [old[k]["name"] for k in old if k not in new]
    updated = [new[k]["name"] for k in new if k in old and new[k] != old[k]]
    dedupe = lambda names: sorted(dict.fromkeys(names))
    return dedupe(added), dedupe(removed), dedupe(updated)


def name_line(verb, names, max_chars):
    """`+ Add 3 APIs: A, B, C`, cut to `max_chars` with `and N more`."""
    head = "+ %s %d API%s: " % (verb, len(names), "" if len(names) == 1 else "s")
    shown = []
    for i, name in enumerate(names):
        rest = len(names) - i - 1
        tail = " and %d more" % rest if rest else ""
        candidate = head + ", ".join(shown + [name]) + tail
        if len(candidate) > max_chars:
            break
        shown.append(name)
    rest = len(names) - len(shown)
    if not shown:
        return ("+ %s %d API%s" % (verb, len(names), "" if len(names) == 1 else "s"))[:max_chars]
    return head + ", ".join(shown) + (" and %d more" % rest if rest else "")


def propose(meta, old, new, date=None, author=None):
    """The next history entry, or None when the API did not change."""
    added, removed, updated = diff(old, new)
    if not (added or removed or updated):
        return None
    limits = meta["history"]["limits"]
    last = meta["history"]["entries"][-1]
    major, minor = version_key(last["version"])
    lines = [name_line(verb, names, limits["max_line_chars"])
             for verb, names in (("Add", added), ("Remove", removed), ("Update", updated)) if names]
    if len(lines) > limits["max_lines"]:
        # Dropping a line would record a snapshot the entry does not describe.
        raise ValueError("the change needs %d lines but history.limits.max_lines is %d"
                         % (len(lines), limits["max_lines"]))
    entry = {
        "version": "%d.%d" % (major, minor + 1),
        "date": date or datetime.date.today().strftime("%Y%m%d"),
        "author": author or last["author"],
        "changes": lines,
    }
    if entry["date"] < last["date"]:
        raise ValueError("date %s is before the last entry's %s" % (entry["date"], last["date"]))
    return validate_entry(entry, limits)


def entry_text(entry):
    """The entry as YAML in the layout `history.entries` uses."""
    q = lambda v: json.dumps(v, ensure_ascii=False)
    lines = ["    - version: %s" % q(entry["version"]),
             "      date: %s" % q(entry["date"]),
             "      author: %s" % q(entry["author"]),
             "      changes:"]
    lines += ["        - %s" % q(c) for c in entry["changes"]]
    return "\n".join(lines) + "\n"


def append_entry(meta_path, entry):
    """Append to the file's text and prove the result still loads; else restore."""
    with open(meta_path, encoding="utf-8") as f:
        original = f.read()
    text = original if original.endswith("\n") else original + "\n"
    with open(meta_path, "w", encoding="utf-8") as f:
        f.write(text + entry_text(entry))
    try:
        after = load_meta(meta_path)
        if after["history"]["entries"][-1] != entry:
            raise ValueError("the appended entry did not become the last history entry; "
                             "keep history.entries the last key of the file")
    except (ValueError, yaml.YAMLError):
        with open(meta_path, "w", encoding="utf-8") as f:
            f.write(original)
        raise


def main():
    parser = argparse.ArgumentParser(description="Propose the next History entry from the API change.")
    parser.add_argument("contract")
    parser.add_argument("meta")
    parser.add_argument("--write", action="store_true", help="append the entry and store its snapshot")
    parser.add_argument("--snapshot", action="store_true",
                        help="store the current API as the snapshot of the last entry")
    parser.add_argument("--date", help="entry date as YYYYMMDD (default: today)")
    parser.add_argument("--author", help="entry author (default: the last entry's)")
    args = parser.parse_args()
    try:
        meta = load_meta(args.meta)
        with open(args.contract, encoding="utf-8") as f:
            new = signatures(json.load(f))
        last = meta["history"]["entries"][-1]["version"]
        if args.snapshot:
            if read_snapshot(last) is not None:
                sys.exit("version %s already has a snapshot; it is not replaced" % last)
            write_snapshot(last, new)
            print("stored the snapshot of %s (%d operations)" % (last, len(new)))
            return
        old = read_snapshot(last)
        if old is None:
            sys.exit("no snapshot for version %s; run with --snapshot once the document of that "
                     "version is rendered" % last)
        entry = propose(meta, old, new, args.date, args.author)
        if entry is None:
            print("the API matches version %s: nothing to add" % last)
            return
        print(entry_text(entry), end="")
        if args.write:
            # Snapshot first: an entry without its snapshot would block the
            # next proposal, a snapshot without its entry is simply removed.
            write_snapshot(entry["version"], new)
            try:
                append_entry(args.meta, entry)
            except (ValueError, yaml.YAMLError, OSError):
                os.remove(snapshot_path(entry["version"]))
                raise
            print("appended to %s and stored the snapshot of %s" % (args.meta, entry["version"]))
    except (ValueError, yaml.YAMLError, OSError) as exc:
        sys.exit(str(exc))


if __name__ == "__main__":
    main()
