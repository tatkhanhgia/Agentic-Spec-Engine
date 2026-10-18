#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Concurrency guard for any script that rewrites the DOCX working copy.

More than one agent edits this repository. On 2026-07-27 a section
(`Download Attachment (base64)`) was deliberately removed by another agent in
between one script's backup and its verification step. The verification diff
compared against the stale backup, reported "no change", and the removal was
then undone by restoring that backup - destroying the other agent's work.

The failure was not in any single edit. It was the assumption that the file on
disk is still the file that was read. This module removes that assumption:

    from docx_guard import load, save

    tree, token = load(DOC_XML)      # remembers the file's digest
    ...  mutate tree ...
    save(tree, DOC_XML, token)       # refuses to write if disk changed

`save` re-hashes the file immediately before writing. If anything else touched
it in the meantime the write is refused and the script exits non-zero, so the
edit can be re-applied on top of the newer content instead of clobbering it.

Never bypass a refusal by re-running with a fresh token without first looking at
what changed - that is exactly the clobber this guard exists to prevent.

CLI:
    python scripts/docx_guard.py            # print the current digest
    python scripts/docx_guard.py <path>
"""
import hashlib
import os
import sys

from lxml import etree

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOC_XML = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")


class ConcurrentModification(RuntimeError):
    """Raised when the file changed between load() and save()."""


def digest(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load(path, parser=None):
    """Parse `path` and return (tree, token). Keep the token for save()."""
    token = digest(path)
    tree = etree.parse(path, parser or etree.XMLParser(huge_tree=True))
    return tree, token


def save(tree, path, token, quiet=False):
    """Write `tree` to `path`, but only if the file still matches `token`."""
    current = digest(path)
    if current != token:
        raise ConcurrentModification(
            "%s changed since it was read.\n"
            "  expected sha256 %s\n"
            "  found    sha256 %s\n"
            "Another agent almost certainly edited it. Do NOT overwrite and do\n"
            "NOT restore an old backup - re-read the file and re-apply this\n"
            "change on top of the current content, then ask the user if the\n"
            "difference looks deliberate." % (path, token[:16], current[:16]))
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    if not quiet:
        print("wrote %s (guard ok)" % path)
    return digest(path)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DOC_XML
    print("%s\n%s" % (path, digest(path)))


if __name__ == "__main__":
    main()
