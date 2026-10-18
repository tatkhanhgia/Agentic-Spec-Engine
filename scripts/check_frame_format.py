#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Validate the formatting rules that are easy to break when editing the DOCX
programmatically, and that a text-level diff cannot see.

Checks, all read from `scripts/GoPaperless_DOCX_Format_Template.md`:

  A. **Method badge scope (2.8.1).** On the first line of every Sample Request
     frame the shading must cover the method token only. A helper that rebuilds
     a paragraph from its first run silently spreads the badge across the whole
     URL - that is what happened when `Switch Enterprise Context` was first
     added, and it is invisible in a text diff because the characters do not
     change.
  B. **Badge colour (2.8.1).** The shading must be the colour the template
     assigns to that HTTP method.
  C. **`{baseURL}` colour (2.8.6).** Red `FF0000`.
  D. **Attribute table alignment (2.4).** `No` and `Presence` columns centered.
  E. **JSON body indentation (2.8.4).** Every Sample Request/Response frame body
     must indent 2 spaces per brace/bracket nesting level. Hand edits and
     clone-based scripts routinely paste bodies with mixed indentation; a text
     diff shows only "whitespace noise" so it slips through review. The width is
     read from the template's 2.8.4 table.
  F. **Presence completeness (2.4).** Every data row of every attribute table
     must display a non-empty Presence value. Eleven rows reached the document
     on 2026-07-27 with blank Presence cells (all of `Create Workflow From
     Template` and `Get Status of Import Excel`) because nothing guarded the
     column: the cell keeps its run and formatting, only the text is missing,
     so it is invisible to both the format checks and a casual read of the
     diff. This check guards completeness only - the value itself must come
     from the backend code or the owner.
  G. **Frame auto-fit (2.8.10).** Every sample frame textbox must auto-grow with
     its text: `<a:spAutoFit/>` in the DrawingML `wps:bodyPr` AND
     `mso-fit-shape-to-text:t` in the VML `v:textbox` style. A fixed-height
     frame clips JSON lines when the text is taller than the shape (silent data
     loss on view/print/PDF export) and holds dead space when it is shorter.
     On 2026-07-27, 206 of 714 frames were fixed-height and dozens clipped
     content; no character changes, so a text diff can never catch it.
  H. **Table width/grid consistency (2.4).** Regular cells in five-column
     attribute tables must use `dxa`, never percentage widths. The generated
     System Connection Information table must additionally have one grid
     column per cell and exact `tcW`/`gridCol` agreement. This catches generic
     cell builders that reuse a percentage intended for another column count.
  I. **Attribute table borders/headers (2.4).** Every attribute table must use
     black (`000000`) borders at size `4` and carry `tblHeader` on the column
     header row (`No | Name | Type | Presence | Description`). Generated Path
     Attributes tables historically used light gray (`D9D9D9`) borders and
     omitted `tblHeader`, so the column header row rendered as plain body text.
  J. **Attribute table start header alignment (2.4).** The merged title row
     (`Request Attributes`, `AddUserToWorkflowRequest`, `Package`, …) must be
     left-aligned. Step 3 of `fix_docx_format.py` centres every `tblHeader`
     row; step 6 must reset row 0 on every attribute table, not only endpoint
     titles. Object-chapter tables titled with the object's own name were left
     centred on 2026-07-27.

Exit code is 0 when clean, 1 when any check fails, so this can gate a commit.

Usage:
    python scripts/check_frame_format.py [path-to-docx] [--quiet]
"""
import os
import re
import sys
import zipfile
import collections

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fix_docx_format import parse_frame_method_rules

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
VML = "urn:schemas-microsoft-com:vml"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX_DEFAULT = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
TEMPLATE = os.path.join(ROOT, "scripts", "GoPaperless_DOCX_Format_Template.md")

METHOD_RE = re.compile(r"^(GET|POST|PUT|DELETE|PATCH)\b")
PLACEHOLDER_RED = "FF0000"
STRING_RE = re.compile(r'"(?:[^"\\]|\\.)*"')


def parse_indent_width(template_path):
    """Read the JSON indent width from template 2.8.4 (`| **Indentation** | N spaces ...`)."""
    with open(template_path, encoding="utf-8") as fh:
        m = re.search(r"\*\*Indentation\*\*\s*\|\s*(\d+)\s+spaces", fh.read())
    return int(m.group(1)) if m else 2


def reindent(lines, width):
    """Canonical indentation for frame body lines, driven by brace/bracket depth."""
    out = []
    depth = 0
    for line in lines:
        s = line.strip()
        if not s:
            out.append("")
            continue
        stripped = STRING_RE.sub('""', s)
        cur = max(0, depth - (1 if stripped[0] in "}]" else 0))
        out.append(" " * (width * cur) + s)
        depth = max(0, depth + stripped.count("{") + stripped.count("[")
                    - stripped.count("}") - stripped.count("]"))
    return out


def body_start_index(lines):
    """Index of the first JSON body line of a frame, or None."""
    first = lines[0].strip() if lines else ""
    if METHOD_RE.match(first):
        # A request body may be a root-level array as well as an object (2.8.4).
        # Matching only "{" made the scan skip the opening "[" and start at the
        # first object *inside* the array, so a correctly indented array body was
        # reported as wrongly indented.
        for i, l in enumerate(lines[1:], 1):
            if l.strip().startswith(("{", "[")):
                return i
        return None
    if first.startswith("Http Code") or first.startswith(("{", "[")):
        i = 0
        while i < len(lines) and not lines[i].strip().startswith(("{", "[")):
            i += 1
        return i if i < len(lines) else None
    return None
def is_sample_frame_line(first):
    """A txbxContent is a sample frame when its first line is a request line,
    an `Http Code` status line, or starts a bare JSON body."""
    return bool(METHOD_RE.match(first)) or first.startswith(("Http Code", "{", "["))


def frame_first_line(txbx):
    paras = txbx.findall(qn("p"))
    if not paras:
        return ""
    return "".join(t.text or "" for t in paras[0].iter(qn("t"))).strip()


ATTRIBUTE_TABLES = ("Header Request Attributes", "Path Attributes",
                    "Request Attributes", "Response Attributes",
                    "Header Response Attributes")

# Per template 2.4 an attribute table is recognised by its column header row, not
# by its title. Endpoint tables use the titles above, but a table in the
# REQUEST/RESPONSE OBJECT and RELATIVE OBJECT chapters is titled with the object's
# own name. Matching titles alone silently excluded every object table from check D.
COLUMN_HEADER = ("No", "Name", "Type", "Presence", "Description")


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def cell_text(tc):
    return " ".join(x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def run_shading(run):
    rPr = run.find(qn("rPr"))
    if rPr is None:
        return None
    shd = rPr.find(qn("shd"))
    return shd.get(qn("fill")) if shd is not None else None


def run_color(run):
    rPr = run.find(qn("rPr"))
    if rPr is None:
        return None
    c = rPr.find(qn("color"))
    return c.get(qn("val")) if c is not None else None


def cell_alignment(tc):
    for p in tc.findall(qn("p")):
        pPr = p.find(qn("pPr"))
        jc = pPr.find(qn("jc")) if pPr is not None else None
        return jc.get(qn("val")) if jc is not None else "left"
    return "left"


def parse_attribute_border_rules(template_path):
    """Read attribute-table border size and color from template 2.4."""
    with open(template_path, encoding="utf-8") as fh:
        text = fh.read()
    size_m = re.search(
        r"\*\*Attribute table border size\*\*\s*\|\s*`?(\d+)`?", text)
    color_m = re.search(
        r"\*\*Attribute table border color\*\*\s*\|\s*`?([0-9A-Fa-f]{6})`?", text)
    return (
        size_m.group(1) if size_m else "4",
        (color_m.group(1) if color_m else "000000").upper(),
    )


def attribute_table_border_issue(tbl, want_size, want_color):
    """Return a reason when tblBorders diverge from the template standard."""
    tbl_pr = tbl.find(qn("tblPr"))
    borders = tbl_pr.find(qn("tblBorders")) if tbl_pr is not None else None
    if borders is None:
        return "missing tblBorders"
    for tag in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = borders.find(qn(tag))
        if edge is None:
            return "%s border missing" % tag
        sz = edge.get(qn("sz"))
        color = (edge.get(qn("color")) or "").upper()
        if sz != want_size or color != want_color:
            return "%s is sz=%s color=%s (want sz=%s color=%s)" % (
                tag, sz, color, want_size, want_color)
    return None


def table_grid_widths(tbl):
    widths = []
    for grid_col in tbl.findall(qn("tblGrid") + "/" + qn("gridCol")):
        value = grid_col.get(qn("w"))
        if value is None or not value.isdigit():
            return []
        widths.append(int(value))
    return widths


def row_width_issue(tr, grid_widths, exact=False):
    """Return a concise invalid-width reason for a regular table row."""
    cells = tr.findall(qn("tc"))
    if len(cells) != len(grid_widths):
        return "%d cells for %d grid columns" % (len(cells), len(grid_widths))
    for col, (tc, expected) in enumerate(zip(cells, grid_widths), 1):
        tcW = tc.find(qn("tcPr") + "/" + qn("tcW"))
        if tcW is None:
            return "column %d has no tcW (want %d dxa)" % (col, expected)
        actual = tcW.get(qn("w"))
        width_type = tcW.get(qn("type"))
        if width_type != "dxa":
            return "column %d is %s %s (percentage widths are forbidden)" % (
                col, actual, width_type)
        if exact and actual != str(expected):
            return "column %d is %s %s (want %d dxa)" % (
                col, actual, width_type, expected)
    return None


def walk(body):
    """Yield (heading, element) with the nearest enclosing endpoint heading."""
    heading = ""
    for el in body:
        if el.tag == qn("p"):
            st = style_of(el)
            if st in ("Heading2", "Heading3", "Heading4"):
                heading = text_of(el).strip()
        yield heading, el


def check(docx_path, quiet=False):
    rules = parse_frame_method_rules(TEMPLATE)
    with zipfile.ZipFile(docx_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(qn("body"))

    problems = collections.defaultdict(list)
    frames = 0
    bodies = 0
    indent_w = parse_indent_width(TEMPLATE)
    border_size, border_color = parse_attribute_border_rules(TEMPLATE)

    for heading, el in walk(body):
        if el.tag == qn("p"):
            # G - frame auto-fit (2.8.10): both stored representations of every
            # sample frame must auto-size to their text. A fixed-height frame
            # clips or pads the JSON and no text diff can see it.
            for wtxbx in el.iter("{%s}txbx" % WPS):
                txbx = wtxbx.find(qn("txbxContent"))
                if txbx is None or not is_sample_frame_line(frame_first_line(txbx)):
                    continue
                bodyPr = wtxbx.getparent().find("{%s}bodyPr" % WPS)
                if bodyPr is None or bodyPr.find("{%s}spAutoFit" % A) is None:
                    problems["G. frame not auto-fit: DrawingML bodyPr lacks spAutoFit (2.8.10)"].append(
                        "%s | %s" % (heading, frame_first_line(txbx)[:50]))
            for vtb in el.iter("{%s}textbox" % VML):
                txbx = vtb.find(qn("txbxContent"))
                if txbx is None or not is_sample_frame_line(frame_first_line(txbx)):
                    continue
                if "mso-fit-shape-to-text:t" not in (vtb.get("style") or ""):
                    problems["G. frame not auto-fit: VML textbox lacks mso-fit-shape-to-text (2.8.10)"].append(
                        "%s | %s" % (heading, frame_first_line(txbx)[:50]))

            for txbx in el.iter(qn("txbxContent")):
                paras = txbx.findall(qn("p"))
                if not paras:
                    continue
                line = text_of(paras[0])

                # E - JSON body indentation, for request AND response frames
                lines = [text_of(p) for p in paras]
                bi = body_start_index(lines)
                if bi is not None:
                    bodies += 1
                    if lines[bi:] != reindent(lines[bi:], indent_w):
                        problems["E. JSON body indentation differs from template 2.8.4"].append(
                            "%s | first line %r" % (heading, lines[0].strip()[:50]))

                m = METHOD_RE.match(line.strip())
                if not m or "{baseURL}" not in line:
                    continue
                frames += 1
                method = m.group(1)
                runs = [r for r in paras[0].findall(qn("r")) if r.find(qn("t")) is not None]
                if not runs:
                    continue

                # A - shading must not extend past the method token
                shaded = [r for r in runs if run_shading(r)]
                shaded_text = "".join(text_of(r) for r in shaded).strip()
                if shaded_text != method:
                    problems["A. badge shading covers more than the method"].append(
                        "%s | shaded=%r" % (heading, shaded_text[:60]))

                # B - badge colour
                want = rules.get(method, {}).get("highlight")
                got = run_shading(runs[0])
                if want and got and got.upper() != want.upper():
                    problems["B. badge colour differs from template"].append(
                        "%s | %s: %s (want %s)" % (heading, method, got, want))

                # C - {baseURL} must be red
                for r in runs:
                    if text_of(r).strip() == "{baseURL}":
                        col = run_color(r)
                        if (col or "").upper() != PLACEHOLDER_RED:
                            problems["C. {baseURL} not red"].append(
                                "%s | color=%s" % (heading, col))
                        break

        elif el.tag == qn("tbl"):
            trs = el.findall(qn("tr"))
            if not trs:
                continue
            title = " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()
            first_row = tuple(cell_text(tc).strip() for tc in trs[0].findall(qn("tc")))
            header = tuple(cell_text(tc).strip()
                           for tc in trs[1].findall(qn("tc")))[:len(COLUMN_HEADER)] \
                if len(trs) > 1 else ()
            is_attribute = header == COLUMN_HEADER
            is_connection = first_row == ("No", "Name", "Value", "Description")
            if not is_attribute and not is_connection:
                continue

            grid_widths = table_grid_widths(el)
            width_rows = trs if is_connection else trs[1:]
            for row_no, tr in enumerate(width_rows, 1):
                issue = row_width_issue(
                    tr, grid_widths, exact=is_connection or title == "Path Attributes")
                if issue:
                    problems["H. table cell widths disagree with tblGrid (2.4)"].append(
                        "%s | %s | row %d: %s" % (
                            heading, title or "System Connection Information", row_no, issue))

            if not is_attribute:
                continue

            border_issue = attribute_table_border_issue(el, border_size, border_color)
            if border_issue:
                problems["I. attribute table border/header (2.4)"].append(
                    "%s | %s | %s" % (heading, title, border_issue))
            if len(trs) > 1:
                tr_pr = trs[1].find(qn("trPr"))
                if tr_pr is None or tr_pr.find(qn("tblHeader")) is None:
                    problems["I. attribute table border/header (2.4)"].append(
                        "%s | %s | column header row missing tblHeader" % (heading, title))

            if trs[0].findall(qn("tc")):
                start_jc = cell_alignment(trs[0].findall(qn("tc"))[0])
                if start_jc != "left":
                    problems["J. attribute table start header not left-aligned (2.4)"].append(
                        "%s | %s | jc=%s" % (heading, title, start_jc))

            for row_no, tr in enumerate(trs[2:], 1):
                tcs = tr.findall(qn("tc"))
                if len(tcs) < 5:
                    continue
                for col, label in ((0, "No"), (3, "Presence")):
                    if cell_alignment(tcs[col]) != "center":
                        problems["D. %s column not centered" % label].append(
                            "%s | %s" % (heading, title))

                # F - Presence cell must not be blank (template 2.4)
                if not cell_text(tcs[3]).strip():
                    name = cell_text(tcs[1]).strip() or ("row %d" % row_no)
                    problems["F. Presence cell empty in attribute table (2.4)"].append(
                        "%s | %s | %s" % (heading, title, name))

    print("request frames checked : %d" % frames)
    print("frame bodies checked   : %d (indent width %d)" % (bodies, indent_w))
    total = sum(len(v) for v in problems.values())
    if not total:
        print("all checks passed")
        return 0

    for kind in sorted(problems):
        items = problems[kind]
        print("\n%s : %d" % (kind, len(items)))
        shown = collections.Counter(items)
        for item, n in shown.most_common(None if not quiet else 5):
            print("   %s%s" % (item, "" if n == 1 else "  (x%d)" % n))
    print("\n%d problem(s)" % total)
    return 1


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    path = args[0] if args else DOCX_DEFAULT
    sys.exit(check(path, quiet="--quiet" in sys.argv))


if __name__ == "__main__":
    main()
