#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Restore per-segment colouring on the request URLs written this session.

Template 2.8.6 colours a request URL by token: every `{placeholder}` is red
`FF0000` and every literal path segment is black `000000`, which is why the
donor's first line is split one run per token:

    run[0] 'POST' (badge)  run[1] ' '  run[2] '{baseURL}' red
    run[3] '/v1/packages/' black  run[4] '{packageId}' red  run[5] '/workflow/users' black

The section builder collapsed the whole URL into run[2] and blanked the rest.
Run[2] is the red one, so the entire path came out red.

Check C in check_frame_format.py did not catch it: it asserts that the run
carrying `{baseURL}` is red, which a fully red URL satisfies. The defect is only
visible by reading the colour of every run, which is what this script fixes and
what `verify` below re-checks.
"""
import os
import re
import shutil
import sys
from copy import deepcopy

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
ROOT = os.path.dirname(SCRIPTS)
sys.path.insert(0, SCRIPTS)

import section_tools as st          # noqa: E402
import docx_guard                   # noqa: E402

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_url_colour_fix.docx")

RED = "FF0000"
BLACK = "000000"

TITLES = [
    "Get Template Workflow Users", "Add Template Workflow Users",
    "Add Template Workflow Placeholder", "Update Template Workflow User",
    "Update Template Workflow Placeholder", "Reorder Template Workflow Participant",
    "Delete Template Workflow Participant", "Clone Template", "Get Template Details",
    "Add Template Field", "Update Template Field",
    "Get Groups", "Get Group", "Get Templates",
]

TOKEN_RE = re.compile(r"\{[^}]*\}|[^{]+")


def colour_of(run):
    rPr = run.find(st.qn("rPr"))
    if rPr is None:
        return None
    c = rPr.find(st.qn("color"))
    return c.get(st.qn("val")) if c is not None else None


def set_colour(run, val):
    rPr = run.find(st.qn("rPr"))
    if rPr is None:
        rPr = run.makeelement(st.qn("rPr"), {})
        run.insert(0, rPr)
    c = rPr.find(st.qn("color"))
    if c is None:
        c = rPr.makeelement(st.qn("color"), {})
        rPr.append(c)
    c.set(st.qn("val"), val)


def resplit(para):
    """Re-split a collapsed URL into one run per token, recolouring each."""
    runs = [r for r in para.findall(st.qn("r")) if r.find(st.qn("t")) is not None]
    if len(runs) < 3:
        return 0
    url = "".join(st.text_of(r) for r in runs[2:])
    if not url.startswith("{baseURL}"):
        return 0
    tokens = TOKEN_RE.findall(url)
    if len(tokens) < 2:
        return 0                      # nothing to split

    donor = runs[2]
    anchor = runs[1]
    for r in runs[2:]:
        para.remove(r)

    for tok in tokens:
        new_r = deepcopy(donor)
        for child in list(new_r):
            if child.tag not in (st.qn("rPr"), st.qn("t")):
                new_r.remove(child)
        for extra in new_r.findall(st.qn("t"))[1:]:
            new_r.remove(extra)
        t = new_r.find(st.qn("t"))
        t.text = tok
        t.set(st.XML_SPACE, "preserve")
        set_colour(new_r, RED if tok.startswith("{") else BLACK)
        anchor.addnext(new_r)
        anchor = new_r
    return 1


def each_request_line(body, title):
    kids = list(body)
    lo = st.find_heading(kids, 3, title)
    s, e = st.section_range(kids, lo)
    for el in kids[s:e]:
        for txbx in el.iter(st.qn("txbxContent")):
            paras = txbx.findall(st.qn("p"))
            if paras and "{baseURL}" in st.text_of(paras[0]):
                yield paras[0]


def verify(body, titles):
    bad = []
    for title in titles:
        for para in each_request_line(body, title):
            for r in para.findall(st.qn("r")):
                t = r.find(st.qn("t"))
                if t is None or not (t.text or ""):
                    continue
                txt = t.text
                if txt.startswith("{") and colour_of(r) != RED:
                    bad.append("%s: %r should be red" % (title, txt))
                elif not txt.startswith("{") and txt.strip() and \
                        colour_of(r) not in (BLACK, "FFFFFF"):
                    bad.append("%s: %r should be black" % (title, txt))
    return bad


def main():
    tree, token = docx_guard.load(DOC)
    body = tree.getroot().find(st.qn("body"))

    total = 0
    for title in TITLES:
        n = sum(resplit(p) for p in each_request_line(body, title))
        total += n
        print("~ %-38s %d frame copies re-split" % (title, n))

    bad = verify(body, TITLES)
    if bad:
        for b in bad:
            print("  FAIL", b)
        sys.exit("colour verification failed")
    print("verified: every {placeholder} red, every literal segment black")

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
    st.pack(UNPACKED, OUT_DOCX)
    print("packed -> %s  (%d paragraphs fixed)" % (os.path.basename(OUT_DOCX), total))


if __name__ == "__main__":
    main()
