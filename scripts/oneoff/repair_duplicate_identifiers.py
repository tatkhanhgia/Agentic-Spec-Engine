#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Make every drawing and bookmark identifier unique again.

Word treats a duplicated identifier as file corruption: it may offer to "repair"
the document and drop content while doing so. Cloning a section to build a new one
- the method CLAUDE.md requires - copies the identifiers along with the drawing,
which is how these arose.

Three independent id spaces are repaired, plus the malformed bookmarks:

1. `wp:docPr/@id` - 9 duplicates, all in the Document Package sections
   (Download Document (base64), Download Signature Summary (base64), Merge
   Document, Convert Document to Paper, Download Attachment).
2. VML `v:shape/@id` and `@o:spid` - 9 and 2 duplicates, same sections. `@o:spid`
   is kept equal to `@id` on the same shape, as Word writes them.
3. `w:bookmarkStart/@id` - 18 ids used by more than one start/end pair.

Safety established by measurement before writing this:

* **No shape uses `mso-next-textbox`.** Nothing chains text from one box into
  another, so reallocating shape ids cannot break a text flow.
* **Bookmark names are never touched.** Hyperlinks resolve by name; only the
  numeric id changes, and a start and its end are always renumbered together.
* The 12 unclosed `w:bookmarkStart` and 8 orphaned `w:bookmarkEnd` are all
  Word's own `_Hlk...` editing markers and **not one of them is the target of a
  hyperlink** (checked: 0 references each). An unclosed start and a stray end are
  malformed, so they are removed rather than repaired.

Pairing is resolved by walking the document in order and closing the most recent
open start of that id, which is what Word does; a duplicated id would otherwise be
ambiguous about which end belongs to which start.

Usage:
    python scripts/oneoff/repair_duplicate_identifiers.py [--dry-run]
"""
import collections
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

O = "urn:schemas-microsoft-com:office:office"


def pair_bookmarks(root):
    """Walk in document order; an end closes the most recent open start of its id."""
    open_ = collections.defaultdict(list)
    pairs, orphan_end = [], []
    for e in root.iter():
        tag = st.local(e)
        if tag == "bookmarkStart":
            open_[e.get(st.qn("id"))].append(e)
        elif tag == "bookmarkEnd":
            i = e.get(st.qn("id"))
            if open_[i]:
                pairs.append((i, open_[i].pop(), e))
            else:
                orphan_end.append(e)
    unclosed = [x for v in open_.values() for x in v]
    return pairs, orphan_end, unclosed


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()

    # ---- safety gate: no linked text boxes -------------------------------
    chained = [d for d in root.iter()
               if st.local(d) == "shape" and "next-textbox" in (d.get("style") or "")]
    if chained:
        sys.exit("%d shapes are chained with mso-next-textbox - reallocating ids "
                 "could break the text flow; stopping" % len(chained))

    anchors_before = {h.get(st.qn("anchor")) for h in root.iter(st.qn("hyperlink"))
                      if h.get(st.qn("anchor"))}
    names_before = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}

    # ---- 1. drawing ids ---------------------------------------------------
    docpr = [d for d in root.iter() if st.local(d) == "docPr"]
    nums = [int(d.get("id")) for d in docpr if (d.get("id") or "").isdigit()]
    next_docpr = max(nums) + 1 if nums else 1
    seen, n_doc = set(), 0
    for d in docpr:
        i = d.get("id")
        if i in seen:
            d.set("id", str(next_docpr))
            d.set("name", "Text Box %d" % next_docpr)
            next_docpr += 1
            n_doc += 1
        else:
            seen.add(i)

    shapes = [d for d in root.iter() if st.local(d) == "shape"]
    spids = []
    for d in root.iter():
        if st.local(d) in ("shape", "shapetype"):
            for v in (d.get("id"), d.get("{%s}spid" % O)):
                m = re.match(r"_x0000_s(\d+)$", v or "")
                if m:
                    spids.append(int(m.group(1)))
    next_spid = max(spids) + 1 if spids else 1000
    seen, n_vml = set(), 0
    for d in shapes:
        i = d.get("id")
        if i in seen:
            new = "_x0000_s%d" % next_spid
            d.set("id", new)
            if d.get("{%s}spid" % O) is not None:
                d.set("{%s}spid" % O, new)
            next_spid += 1
            n_vml += 1
        else:
            seen.add(i)

    # ---- 2. malformed bookmarks ------------------------------------------
    pairs, orphan_end, unclosed = pair_bookmarks(root)
    kept_names = {b.get(st.qn("name")) for _i, b, _e in pairs}
    bad = [b for b in unclosed if b.get(st.qn("name")) in anchors_before]
    if bad:
        sys.exit("refusing: %d unclosed bookmarks are hyperlink targets: %s"
                 % (len(bad), [b.get(st.qn("name")) for b in bad]))
    for b in unclosed:
        b.getparent().remove(b)
    for e in orphan_end:
        e.getparent().remove(e)

    # ---- 3. duplicate bookmark ids ---------------------------------------
    used = {int(i) for i, _s, _e in pairs if i.isdigit()}
    next_bm = max(used) + 1 if used else 1
    counts = collections.Counter(i for i, _s, _e in pairs)
    taken, n_bm = set(), 0
    for i, s, e in pairs:
        if counts[i] > 1 and i in taken:
            s.set(st.qn("id"), str(next_bm))
            e.set(st.qn("id"), str(next_bm))
            next_bm += 1
            n_bm += 1
        else:
            taken.add(i)

    print("docPr re-ided            : %d" % n_doc)
    print("VML shape re-ided        : %d" % n_vml)
    print("bookmark pairs re-ided   : %d" % n_bm)
    print("unclosed bookmarks removed: %d" % len(unclosed))
    print("orphaned bookmarkEnd removed: %d" % len(orphan_end))

    # ---- verification -----------------------------------------------------
    dp = [d.get("id") for d in root.iter() if st.local(d) == "docPr"]
    sh = [d.get("id") for d in root.iter() if st.local(d) == "shape"]
    sp = [d.get("{%s}spid" % O) for d in root.iter()
          if st.local(d) == "shape" and d.get("{%s}spid" % O)]
    pairs2, orphan2, unclosed2 = pair_bookmarks(root)
    bmids = [i for i, _s, _e in pairs2]
    names_after = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    lost = (names_before - names_after) & anchors_before

    print()
    print("duplicate docPr   : %d" % (len(dp) - len(set(dp))))
    print("duplicate shape   : %d" % (len(sh) - len(set(sh))))
    print("duplicate spid    : %d" % (len(sp) - len(set(sp))))
    print("duplicate bookmark: %d" % (len(bmids) - len(set(bmids))))
    print("unclosed / orphan : %d / %d" % (len(unclosed2), len(orphan2)))
    print("linked names lost : %d %s" % (len(lost), sorted(lost)))

    if lost:
        sys.exit("a hyperlink target disappeared - refusing to write")

    if dry:
        print("\n--dry-run: nothing written")
        return
    try:
        guarded_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("\nREFUSED TO WRITE\n%s" % ex)
    st.pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
