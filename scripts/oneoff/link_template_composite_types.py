#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Give the Template module's composite Type cells real bookmark links.

`TECHNICAL_DEBT.md` §11: the sections added on 2026-07-27 typed their composite
attributes as plain text, so nothing linked to the object it actually refers to.
This resolves the cases where the target object is known; the rest are left
typed `Object`/`Array` on purpose and the reason is recorded, because a link to
the wrong object is worse than no link - the convention teaches readers to trust
it.

Two new objects, both named after the `Owner(Package Details)` /
`Document(Package Details)` precedent for "same concept, different shape":

* `TemplateDocumentResponse(Get Template Details)` — `Get Template Details`
  returns `form_fields` inside each document, the other two template calls do
  not (`GetTemplateDetailsSerializer.writeDocuments` vs
  `GetTemplateSerializer.writeDocuments`). One object cannot describe both.
* `Template(Get Templates)` — the list item returned by `Get Templates`. The
  existing RELATIVE OBJECT `Template` is a different shape entirely
  (`template_name`, `read_only`: the Apply Workflow Template request).

Deliberately NOT linked:

* `workflow`, `workflow_details` — the RELATIVE OBJECT `WorkflowDetail` holds a
  single `filter_level` row and `Workflow` was never checked against
  `GetWorkflowDTO`. Linking on the strength of the name would assert a match
  nobody verified.
* `workflow_categories`, `tasks`, `smart_form` — no object exists and their
  shapes are not established from the code.

ORDERING MATTERS. `set_cell_text` drops every direct child of the cell's
paragraph, including `w:hyperlink`, so all table content edits happen before any
link is written. For the same reason, re-running
`add_template_workflow_sections.py`, `add_template_rest_sections.py` or
`move_get_templates_section.py` after this script would silently strip every
link it just made.
"""
import os
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
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_type_links.docx")

DOC_OBJECT_ROWS = [
    ("id", "Integer", "M", "The ID of the document inside the template."),
    ("document_id", "Integer", "M", "The ID of the document. Carries the same value as id."),
    ("name", "String", "M", "File name of the document."),
]

DETAILS_DOC_ROWS = DOC_OBJECT_ROWS + [
    ("form_fields", "Array", "M",
     "The form fields placed on this document in the template. An empty array when none are placed."),
]

TEMPLATE_ITEM_ROWS = [
    ("template_id", "Integer", "M", "The ID of the template."),
    ("template_name", "String", "M", "Name of the template."),
    ("template_public", "Boolean", "M", "True if the template is marked as public."),
    ("read_only", "Boolean", "M", "True if the template is marked as read-only."),
    ("created_at", "String", "M", "The date and time the template was created, in UTC."),
    ("description", "String", "M", "Description of the template."),
    ("enterprise", "Boolean", "M", "True if the template belongs to the enterprise rather than to the user."),
    ("documents", "List<TemplateDocumentResponse>", "M", "The documents contained in the template."),
    ("smart_form", "Object", "M", "Summary of the smart form fields, as active_count and total_count."),
]

PARTICIPANT_ROWS = [
    ("order", "Integer", "M", "The position of the participant in the template workflow."),
    ("user_email", "String", "O", "Email address of the participant. Returned only for a real user, not for a placeholder."),
    ("user_name", "String", "O", "Full name of the participant. Returned only for a real user, not for a placeholder."),
    ("placeholder", "String", "O", "Name of the placeholder. Returned only when the participant is a placeholder, in which case user_email and user_name are absent."),
    ("role", "String", "M", "The role of the participant in the workflow. Possible values are: + SIGNER + REVIEWER + EDITOR + INPERSON_HOST + CARBON_COPY"),
    ("signing_order", "Integer", "O", "The order in which this participant signs. Applies when the workflow type is CUSTOM."),
    ("guest_user", "Boolean", "M", "True if the participant is not a registered user of the system."),
    ("email_language_code", "String", "M", "The language used for notification emails sent to this participant. Possible values are: + en + vi"),
    ("email_notification", "Boolean", "O", "If true, the participant is notified by email."),
    ("mobile_number", "String", "O", "Mobile number of the participant."),
]

NEW_OBJECTS = [
    # (heading, bookmark, rows, chapter, insert-after object heading)
    ("TemplateDocumentResponse(Get Template Details)",
     "_TemplateDocumentResponse(Get_Template_Details)",
     DETAILS_DOC_ROWS, "REQUEST/RESPONSE OBJECT", "TemplateDocumentResponse"),
    ("Template(Get Templates)", "_Template(Get_Templates)",
     TEMPLATE_ITEM_ROWS, "RELATIVE OBJECT", "Template"),
    ("TemplateWorkflowParticipant", "_TemplateWorkflowParticipant",
     PARTICIPANT_ROWS, "RELATIVE OBJECT", "Template(Get Templates)"),
]

# (endpoint section, attribute, displayed label, anchor)
LINKS = [
    ("Get Templates", "templates", "List<template>", "_Template(Get_Templates)"),
    ("Get Templates", "documents", "List<TemplateDocumentResponse>", "_TemplateDocumentResponse"),
    ("Get Template Details", "documents", "List<TemplateDocumentResponse>",
     "_TemplateDocumentResponse(Get_Template_Details)"),
    ("Get Template Details", "participants", "List<TemplateWorkflowParticipant>",
     "_TemplateWorkflowParticipant"),
    ("Add Template Field", "dimension", "dimension", "_Dimension_1"),
    ("Update Template Field", "dimension", "dimension", "_Dimension_1"),
]


def h2_indices(kids, title):
    return [i for i, e in enumerate(kids)
            if e.tag == st.qn("p") and st.style_of(e) == "Heading2"
            and st.text_of(e).strip() == title]


def h2_in_chapter(kids, title, chapter):
    """Resolve an H2 by title *and* owning H1.

    Two H2s are called `Template`: the API category and the RELATIVE OBJECT.
    `find_heading` returns the first, which would drop a new object into the
    middle of the API chapter.
    """
    cur1 = None
    for i, e in enumerate(kids):
        if e.tag != st.qn("p"):
            continue
        s = st.style_of(e)
        if s == "Heading1":
            cur1 = st.text_of(e).strip()
        elif s == "Heading2" and cur1 == chapter and st.text_of(e).strip() == title:
            return i
    raise LookupError("no H2 %r under %r" % (title, chapter))


def add_bookmark(root, para, name):
    """Anchor a bookmark on a heading. clone_section strips them, so a cloned
    object heading arrives with nothing for a hyperlink to resolve to."""
    bid = str(st.next_bookmark_id(root))
    bs = para.makeelement(st.qn("bookmarkStart"),
                          {st.qn("id"): bid, st.qn("name"): name})
    be = para.makeelement(st.qn("bookmarkEnd"), {st.qn("id"): bid})
    para.insert(1 if para.find(st.qn("pPr")) is not None else 0, bs)
    para.append(be)


def fill_object_table(tbl, rows):
    existing = st.data_rows(tbl)
    for i, (name, typ, presence, desc) in enumerate(rows):
        if i < len(existing):
            tcs = existing[i].findall(st.qn("tc"))
            for col, val in ((1, name), (2, typ), (3, presence), (4, desc)):
                st.set_cell_text(tcs[col], val)
        else:
            st.append_row(tbl, name, typ, presence, desc)
    for extra in st.data_rows(tbl)[len(rows):]:
        extra.getparent().remove(extra)
    st.renumber(tbl)


def only_table(elements):
    """The single table of an object section.

    Object tables cannot be found by title: `TemplateDocumentResponse` carries a
    title row reading `GetCertificatesProfileResponse`, left behind when it was
    first cloned. Selecting by position and rewriting the title fixes that.
    """
    tbls = [el for el in elements if el.tag == st.qn("tbl")]
    if len(tbls) != 1:
        raise LookupError("expected exactly one table, found %d" % len(tbls))
    return tbls[0]


def set_table_title(tbl, title):
    st.set_cell_text(tbl.findall(st.qn("tr"))[0].findall(st.qn("tc"))[0], title)


def h3_in_category(kids, title, category):
    """Resolve an H3 by title *and* owning H2.

    `Get Templates` exists twice at Heading 3: the real section in the Template
    chapter, and the cross-reference stub left behind in Personal Settings.
    `find_heading` returns the stub, which carries no tables.
    """
    cur2 = None
    for i, e in enumerate(kids):
        if e.tag != st.qn("p"):
            continue
        s = st.style_of(e)
        if s == "Heading2":
            cur2 = st.text_of(e).strip()
        elif s == "Heading3" and cur2 == category and st.text_of(e).strip() == title:
            return i
    raise LookupError("no H3 %r under %r" % (title, category))


def find_type_cell(kids, section_title, attribute):
    lo = h3_in_category(kids, section_title, "Template")
    s, e = st.section_range(kids, lo)
    for el in kids[s:e]:
        if el.tag != st.qn("tbl"):
            continue
        for tr in st.data_rows(el):
            tcs = tr.findall(st.qn("tc"))
            if len(tcs) >= 5 and st.cell_text(tcs[1]) == attribute:
                return tcs[2]
    raise LookupError("no %r row under %r" % (attribute, section_title))


def main():
    tree, token = docx_guard.load(DOC)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    # ---- 1. content edits first: they would destroy any hyperlink present ----
    kids = list(body)
    tdr = h2_in_chapter(kids, "TemplateDocumentResponse", "REQUEST/RESPONSE OBJECT")
    s, e = st.section_range(kids, tdr)
    tbl = only_table(kids[s:e])
    old_title = st.table_title(tbl)
    set_table_title(tbl, "TemplateDocumentResponse")
    fill_object_table(tbl, DOC_OBJECT_ROWS)
    print("~ TemplateDocumentResponse: document_id added (was id/name only)")
    print("~ TemplateDocumentResponse: stale title row %r -> %r"
          % (old_title, "TemplateDocumentResponse"))

    # ---- 2. new objects, cloned from a real one so styles come along ----
    for heading, bookmark, rows, chapter, after in NEW_OBJECTS:
        kids = list(body)
        donor_i = h2_in_chapter(kids, "TemplateDocumentResponse",
                                "REQUEST/RESPONSE OBJECT")
        ds, de = st.section_range(kids, donor_i)
        clone = [deepcopy(el) for el in kids[ds:de]]
        clone, _ = st.without_loose_bookmarks(clone)
        st.strip_bookmarks(clone)
        nd, ns = st.scan_id_space(root)
        st.reassign_ids(clone, nd, ns)

        st.set_para_text(clone[0], heading)
        add_bookmark(root, clone[0], bookmark)
        tbl = only_table(clone)
        set_table_title(tbl, heading)
        fill_object_table(tbl, rows)

        anchor_i = h2_in_chapter(list(body), after, chapter)
        _, a_end = st.section_range(list(body), anchor_i)
        node = list(body)[a_end - 1]
        for el in clone:
            node.addnext(el)
            node = el
        print("+ %-46s %-22s (%d rows)" % (heading, chapter, len(rows)))

    # ---- 3. links last ----
    kids = list(body)
    donor_tc = find_type_cell(kids, "Get Template By ID", "documents")
    for section, attribute, label, anchor in LINKS:
        tc = find_type_cell(list(body), section, attribute)
        st.set_type_link(tc, label, anchor, donor_tc)
        print("-> %-24s %-14s %-34s %s" % (section, attribute, label, anchor))

    # ---- 4. verify every link resolves ----
    names = {b.get(st.qn("name")) for b in root.iter(st.qn("bookmarkStart"))}
    bad = []
    for section, attribute, label, anchor in LINKS:
        tc = find_type_cell(list(body), section, attribute)
        links = list(tc.iter(st.qn("hyperlink")))
        if not links:
            bad.append("%s.%s lost its hyperlink" % (section, attribute))
            continue
        a = links[0].get(st.qn("anchor"))
        if a not in names:
            bad.append("%s.%s anchor %r resolves to nothing" % (section, attribute, a))
        if not st.cell_text(tc).strip():
            bad.append("%s.%s label is empty" % (section, attribute))
    if bad:
        for b in bad:
            print("  FAIL", b)
        sys.exit("link verification failed")
    print("verified: %d links, every anchor resolves, no empty label" % len(LINKS))

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
    st.pack(UNPACKED, OUT_DOCX)
    print("packed -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
