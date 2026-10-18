#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Shared helpers for building a new endpoint section by cloning an existing one.

CLAUDE.md requires sections to be written by cloning a donor that already has the
target shape, never by assembling WordprocessingML by hand. Cloning duplicates
identifiers that must be made unique again or Word reports the file as corrupt,
and it is easy to destroy the badge formatting of the first frame line. Both
hazards are handled here once instead of in every one-off script.

Nothing in this module decides content. It only moves and relabels existing XML.
"""
import os
import re
import shutil
import zipfile
from copy import deepcopy

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


# --------------------------------------------------------------------------
# reading
# --------------------------------------------------------------------------
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


def table_title(tbl):
    trs = tbl.findall(qn("tr"))
    if not trs:
        return ""
    return " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()


def find_heading(kids, level, title):
    want = "Heading%d" % level
    for i, e in enumerate(kids):
        if e.tag == qn("p") and style_of(e) == want and text_of(e).strip() == title:
            return i
    raise LookupError("heading not found: %s" % title)


def section_range(kids, start):
    """From a heading index to just before the next heading of any level."""
    for j in range(start + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and re.match(r"Heading\d$", style_of(e)):
            return start, j
    return start, len(kids)


# --------------------------------------------------------------------------
# writing text
# --------------------------------------------------------------------------
def set_para_text(p, text):
    """Reduce a paragraph to one run carrying `text`, keeping its formatting.

    Never use this on the first line of a request frame: that line's first run is
    the method badge, so keeping run[0] and dropping the rest spreads the badge
    shading and white text across the whole URL. Use `set_path_suffix` instead.
    """
    pPr = p.find(qn("pPr"))
    template = None
    for r in p.iter(qn("r")):
        if r.find(qn("t")) is not None:
            template = r
            break
    if template is None:
        raise ValueError("paragraph has no text run: %r" % text_of(p))

    new_run = deepcopy(template)
    for child in list(new_run):
        if child.tag not in (qn("rPr"), qn("t")):
            new_run.remove(child)
    for extra in new_run.findall(qn("t"))[1:]:
        new_run.remove(extra)
    t = new_run.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")

    for child in list(p):
        if child is not pPr:
            p.remove(child)
    p.append(new_run)


def set_cell_text(tc, text):
    """Plain-text cell content.

    This DESTROYS any `w:hyperlink` in the cell, because it drops every direct
    child of the paragraph. That is correct for the No / Name / Presence /
    Description columns and for a primitive `Type`, but never use it on a
    composite `Type` cell - those carry a hyperlink to the object's bookmark and
    the link is the real content. Use `set_type_link` there instead.
    """
    paras = tc.findall(qn("p"))
    set_para_text(paras[0], text)
    for extra in paras[1:]:
        tc.remove(extra)


def set_text_keep_bookmarks(p, text):
    """Set a paragraph's text without disturbing bookmarks anchored on it.

    Heading paragraphs carry `w:bookmarkStart` / `w:bookmarkEnd` children. The
    plain `set_para_text` removes every child except `w:pPr`, which silently
    deletes those bookmarks and orphans every hyperlink that pointed at them.
    """
    runs = [r for r in p.findall(qn("r")) if r.find(qn("t")) is not None]
    if not runs:
        raise ValueError("paragraph has no direct text run: %r" % text_of(p))
    for extra in runs[1:]:
        p.remove(extra)
    keep = runs[0]
    for extra in keep.findall(qn("t"))[1:]:
        keep.remove(extra)
    t = keep.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")


def set_type_link(tc, label, anchor, donor_tc):
    """Make a composite `Type` cell a live link to an object's bookmark.

    The paragraph is cloned from a cell that already links correctly, so the
    style, indent and run properties come along instead of being invented. Only
    the visible label and the anchor are substituted.
    """
    donor_p = None
    for p in donor_tc.findall(qn("p")):
        if p.find(qn("hyperlink")) is not None:
            donor_p = p
            break
    if donor_p is None:
        raise LookupError("donor cell carries no hyperlink: %r" % text_of(donor_tc))

    new_p = deepcopy(donor_p)
    link = new_p.find(qn("hyperlink"))
    link.set(qn("anchor"), anchor)

    # The house style links only the inner type name: a `List<x>` cell is a plain
    # `List<` run, then the hyperlink around `x`, then a plain `>`. Putting the
    # whole label inside the hyperlink instead produces `List<List<x>>`, so the
    # donor's surrounding runs decide how much of the label the link carries.
    children = list(new_p)
    at = children.index(link)
    prefix = "".join(text_of(c) for c in children[:at] if local(c) == "r")
    suffix = "".join(text_of(c) for c in children[at + 1:] if local(c) == "r")

    inner = label
    fits = (not prefix or label.startswith(prefix)) and (not suffix or label.endswith(suffix))
    if fits:
        inner = label[len(prefix):len(label) - len(suffix) if suffix else None]
    else:
        # the label has no wrapper - drop the donor's and link the whole thing
        for c in children:
            if local(c) == "r":
                new_p.remove(c)

    runs = [r for r in link.findall(qn("r")) if r.find(qn("t")) is not None]
    for extra in runs[1:]:
        link.remove(extra)
    t = runs[0].find(qn("t"))
    t.text = inner
    t.set(XML_SPACE, "preserve")

    for p in tc.findall(qn("p")):
        tc.remove(p)
    tc.append(new_p)


def next_bookmark_id(root):
    ids = [int(b.get(qn("id"))) for b in root.iter(qn("bookmarkStart"))
           if (b.get(qn("id")) or "").isdigit()]
    return (max(ids) + 1) if ids else 1


def set_path_suffix(para, old, new):
    """Rewrite only the run that carries `old` on a frame's first line.

    The method badge lives in its own run and the path is split across several
    more; touching a single run leaves every other run - and therefore all the
    badge formatting required by template 2.8.1 - exactly as the donor had it.
    """
    hits = 0
    for r in para.findall(qn("r")):
        t = r.find(qn("t"))
        if t is not None and t.text and old in t.text:
            t.text = t.text.replace(old, new)
            t.set(XML_SPACE, "preserve")
            hits += 1
    if not hits:
        raise LookupError("no run carries %r in %r" % (old, text_of(para)))
    return hits


# --------------------------------------------------------------------------
# frames
# --------------------------------------------------------------------------
def each_frame(elements, method_prefix):
    """Yield (txbxContent, paragraphs) for frames whose first line starts with `method_prefix`.

    A frame is stored twice - once under mc:Choice and once under mc:Fallback -
    so every edit must be applied to each copy the iterator returns.
    """
    for el in elements:
        for txbx in el.iter(qn("txbxContent")):
            paras = txbx.findall(qn("p"))
            if paras and text_of(paras[0]).startswith(method_prefix):
                yield txbx, paras


def replace_frame_body(txbx, paras, lines):
    """Replace the JSON body of a frame, keeping the header lines above it."""
    texts = [text_of(p) for p in paras]
    try:
        body_start = next(i for i, l in enumerate(texts) if l.strip() == "{")
    except StopIteration:
        raise LookupError("frame has no JSON body")
    template = paras[body_start]
    anchor = paras[body_start - 1]
    for p in paras[body_start:]:
        txbx.remove(p)
    for line in lines:
        new_p = deepcopy(template)
        set_para_text(new_p, line)
        anchor.addnext(new_p)
        anchor = new_p


# --------------------------------------------------------------------------
# tables
# --------------------------------------------------------------------------
def data_rows(tbl):
    """Attribute tables carry a title row and a column-header row before the data."""
    return tbl.findall(qn("tr"))[2:]


def append_row(tbl, name, typ, presence, description):
    """Clone the last data row and fill it in, then renumber the `No` column."""
    rows = data_rows(tbl)
    if not rows:
        raise LookupError("table %r has no data row to clone" % table_title(tbl))
    new_tr = deepcopy(rows[-1])
    tcs = new_tr.findall(qn("tc"))
    set_cell_text(tcs[1], name)
    set_cell_text(tcs[2], typ)
    set_cell_text(tcs[3], presence)
    set_cell_text(tcs[4], description)
    rows[-1].addnext(new_tr)
    renumber(tbl)
    return new_tr


def set_row(tbl, name, typ=None, presence=None, description=None, new_name=None):
    """Overwrite selected cells of the row whose Name column equals `name`."""
    for tr in data_rows(tbl):
        tcs = tr.findall(qn("tc"))
        if cell_text(tcs[1]) != name:
            continue
        if new_name is not None:
            set_cell_text(tcs[1], new_name)
        if typ is not None:
            set_cell_text(tcs[2], typ)
        if presence is not None:
            set_cell_text(tcs[3], presence)
        if description is not None:
            set_cell_text(tcs[4], description)
        return tr
    raise LookupError("no row named %r in %r" % (name, table_title(tbl)))


def drop_row(tbl, name):
    """Remove the data row whose Name column equals `name`, then renumber."""
    for tr in data_rows(tbl):
        tcs = tr.findall(qn("tc"))
        if cell_text(tcs[1]) == name:
            tr.getparent().remove(tr)
            renumber(tbl)
            return True
    raise LookupError("no row named %r in %r" % (name, table_title(tbl)))


def renumber(tbl):
    for n, tr in enumerate(data_rows(tbl), start=1):
        set_cell_text(tr.findall(qn("tc"))[0], str(n))


def find_table(elements, title):
    for el in elements:
        if el.tag == qn("tbl") and table_title(el) == title:
            return el
    raise LookupError("table not found: %s" % title)


def drop_table(elements, title):
    """Remove a table and the blank paragraph that follows it. `elements` is a list."""
    for i, el in enumerate(elements):
        if el.tag == qn("tbl") and table_title(el) == title:
            nxt = elements[i + 1] if i + 1 < len(elements) else None
            elements.pop(i)
            if nxt is not None and nxt.tag == qn("p") and not text_of(nxt).strip():
                elements.remove(nxt)
            return True
    return False


# --------------------------------------------------------------------------
# identifier hygiene
# --------------------------------------------------------------------------
def strip_bookmarks(elements):
    """Drop every bookmark from a clone.

    Both ends of a section's bookmarks live inside the section, so removing the
    pair leaves nothing dangling - and a duplicated bookmark id is one of the
    things Word reports as corruption.

    A bookmark can also sit directly between two body children rather than
    inside a paragraph. Such an element has no parent once it has been cloned,
    so it cannot be removed here - drop it from the list first with
    `without_loose_bookmarks`.
    """
    n = 0
    for el in elements:
        if local(el) in ("bookmarkStart", "bookmarkEnd"):
            continue
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for b in list(el.iter(qn(tag))):
                b.getparent().remove(b)
                n += 1
    return n


def without_loose_bookmarks(elements):
    """Drop bookmarks that are body children in their own right.

    `strip_bookmarks` can only unlink a bookmark from a parent, and a bookmark
    sitting directly in the body has none once cloned. Leaving it in would
    duplicate its id, which Word reports as corruption.
    """
    kept = [e for e in elements if local(e) not in ("bookmarkStart", "bookmarkEnd")]
    return kept, len(elements) - len(kept)


# Ids handed out during this process but not yet attached to the tree.
#
# A script that builds several sections calls clone_section once per section and
# only inserts them at the end. Scanning the tree alone would therefore return the
# same starting id for every clone and they would collide with each other - which
# is exactly how six duplicate docPr ids reached the document on 2026-07-27.
_HANDED_OUT = {"docpr": set(), "spid": set()}


def scan_id_space(root):
    """Next free id in each of the two independent id spaces."""
    docpr = [int(d.get("id")) for d in root.iter()
             if local(d) == "docPr" and (d.get("id") or "").isdigit()]
    docpr.extend(_HANDED_OUT["docpr"])
    spids = list(_HANDED_OUT["spid"])
    for d in root.iter():
        if local(d) in ("shape", "shapetype"):
            for v in (d.get("id"), d.get("{%s}spid" % O)):
                m = re.match(r"_x0000_s(\d+)$", v or "")
                if m:
                    spids.append(int(m.group(1)))
    return (max(docpr) + 1 if docpr else 1), (max(spids) + 1 if spids else 1000)


def reassign_ids(elements, next_docpr, next_spid):
    """Give cloned drawings fresh identifiers in both id spaces."""
    n_doc = n_vml = 0
    for el in elements:
        for d in el.iter():
            if local(d) == "docPr":
                d.set("id", str(next_docpr))
                d.set("name", "Text Box %d" % next_docpr)
                _HANDED_OUT["docpr"].add(next_docpr)
                next_docpr += 1
                n_doc += 1
        for d in el.iter():
            if local(d) == "shape":
                d.set("id", "_x0000_s%d" % next_spid)
                if d.get("{%s}spid" % O) is not None:
                    d.set("{%s}spid" % O, "_x0000_s%d" % next_spid)
                _HANDED_OUT["spid"].add(next_spid)
                next_spid += 1
                n_vml += 1
    return n_doc, n_vml


def clone_section(root, kids, donor_title, level=3):
    """Deep-copy a section and make every duplicated identifier unique again."""
    start, end = section_range(kids, find_heading(kids, level, donor_title))
    clone = [deepcopy(e) for e in kids[start:end]]
    clone, n_loose = without_loose_bookmarks(clone)
    n_bm = strip_bookmarks(clone)
    next_docpr, next_spid = scan_id_space(root)
    n_doc, n_vml = reassign_ids(clone, next_docpr, next_spid)
    return clone, {
        "range": (start, end),
        "elements": len(clone),
        "bookmarks_stripped": n_bm + n_loose,
        "docpr_reassigned": n_doc,
        "vml_reassigned": n_vml,
    }


def insert_after_section(body, kids, anchor_title, elements, level=3):
    """Place a built section directly after the anchor section."""
    _, a_end = section_range(kids, find_heading(kids, level, anchor_title))
    anchor = kids[a_end - 1]
    for el in elements:
        anchor.addnext(el)
        anchor = el


# --------------------------------------------------------------------------
# packaging
# --------------------------------------------------------------------------
def pack(src_dir, out_path):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(src_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for base, _dirs, files in os.walk(src_dir):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                if rel != "[Content_Types].xml":
                    z.write(full, rel)
    shutil.move(tmp, out_path)
