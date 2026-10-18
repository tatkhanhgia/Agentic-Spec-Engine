#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Second Template tranche: clone/details/field sections, plus the pagination fixes.

Adds four sections by cloning same-method donors, then corrects three paginated
paths that the document had in the wrong order.

The donor must share the HTTP method: the method badge carries a per-method
shading colour (format template 2.8.1, checks A and B in check_frame_format.py),
so cloning a PUT to build a POST would silently produce a wrongly coloured badge.

Pagination: `ValidateRequestRoot.getPagination()` reads the URL as
`.../{pageNo}/{recordPerPage}` - it takes `data[len-2]` as the page number and
`data[len-1]` as the row count. Three documented paths had the two the wrong way
round. `Get Templates` was wrong twice over: it also omitted the template-type
segment entirely, so the documented URL could not match the route at all.
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
ROOT = os.path.dirname(SCRIPTS)
sys.path.insert(0, SCRIPTS)

import section_tools as st          # noqa: E402
import docx_guard                   # noqa: E402
from add_template_workflow_sections import (  # noqa: E402
    ERROR_ROWS, TEMPLATE_ID_DESC, build, fill_table)

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_template_rest.docx")

DOC_ID_DESC = ("This is the ID of the document inside the template, as returned "
               "in the documents attribute of Get Template Details.")
FIELD_TYPE_DESC = ("The type of field to place on the document. Possible values "
                   "are: + signature + text + checkbox")

# Field payload as exported in the Postman collection for these two calls. The
# field schema is shared with the package-level field endpoints under Document
# Package, which document each attribute in full.
FIELD_ROWS = [
    ("field_name", "String", "M", "Unique name identifying the field on the document."),
    ("renamed_as", "String", "O", "New name for the field. Used to rename an existing field."),
    ("page", "Integer", "M", "The page of the document on which the field is placed."),
    ("dimension", "dimension", "M", "Position and size of the field on the page."),
    ("required", "Boolean", "O", "If true, the field must be filled before the document can be signed."),
    ("visible_enabled", "Boolean", "O", "If true, the field is visible on the document."),
    ("type_name", "String", "M", "The concrete field type. Possible values are: + SIGNATURE + TEXT + CHECKBOX"),
]

FIELD_SAMPLE = [
    "{",
    '  "field_name": "Signature01",',
    '  "renamed_as": "Signature01",',
    '  "page": 1,',
    '  "dimension": {',
    '    "x": 100,',
    '    "y": 500,',
    '    "width": 180,',
    '    "height": 60',
    "  },",
    '  "required": true,',
    '  "visible_enabled": true,',
    '  "type_name": "SIGNATURE"',
    "}",
]

DETAILS_SAMPLE = [
    "{",
    '  "id": 1,',
    '  "name": "Employment Contract Template",',
    '  "public": false,',
    '  "read_only": false,',
    '  "description": "Standard employment contract",',
    '  "created_on": "2026-07-27T09:15:00Z",',
    '  "enterprise": true,',
    '  "documents": [',
    "    {",
    '      "id": 12,',
    '      "document_id": 12,',
    '      "name": "contract.pdf",',
    '      "form_fields": []',
    "    }",
    "  ],",
    '  "workflow": {},',
    '  "workflow_details": [],',
    '  "workflow_categories": [],',
    '  "participants": [],',
    '  "tasks": {}',
    "}",
]

DETAILS_ROWS = [
    ("id", "Integer", "M", "The ID of the template."),
    ("name", "String", "M", "Name of the template."),
    ("public", "Boolean", "M", "True if the template is marked as public."),
    ("read_only", "Boolean", "M", "True if the template is marked as read-only."),
    ("description", "String", "M", "Description of the template."),
    ("created_on", "String", "M", "The date and time the template was created, in UTC."),
    ("enterprise", "Boolean", "M", "True if the template belongs to the enterprise rather than to the user."),
    ("documents", "List<TemplateDocumentResponse>", "M", "The documents contained in the template, each with the form fields placed on it."),
    ("workflow", "Object", "M", "The workflow configuration stored in the template. An empty object when the template has no workflow yet."),
    ("workflow_details", "Array", "M", "The workflow detail entries stored in the template. An empty array when there are none."),
    ("workflow_categories", "Array", "M", "The categories assigned to the workflow. An empty array when there are none."),
    ("participants", "Array", "M", "The participants configured in the template workflow. See Get Template Workflow Users for the shape of each entry."),
    ("tasks", "Object", "M", "The tasks stored in the template. An empty object when there are none."),
]

SECTIONS = [
    dict(
        title="Clone Template",
        donor="Save Workflow as Template",
        method="POST",
        path="{baseURL}/v1/templates/{id}/clone",
        desc=("Business applications can use this service API to create a copy of an "
              "existing personal/enterprise template under a new name. The documents, "
              "fields and workflow participants of the source template are copied to "
              "the new one."),
        req_body=["{", '  "name": "Employment Contract Template (copy)",',
                  '  "description": "Cloned from the standard contract template"', "}"],
        resp_body=["{}"],
        path_rows=[("id", "Integer", "M",
                    "This is the ID of the template to be cloned.")],
        req_rows=[("name", "String", "M", "Name of the new template. Between 1 and 255 characters."),
                  ("description", "String", "O", "Description of the new template. Up to 200 characters.")],
        resp_rows=ERROR_ROWS,
        status="200",
    ),
    dict(
        title="Get Template Details",
        donor="Get Template By ID",
        method="GET",
        path="{baseURL}/v1/templates/{id}/details",
        desc=("Business applications can use this service API to get the full "
              "content of a personal/enterprise template, including its documents "
              "and form fields, its workflow configuration and its participants. "
              "Get Template By ID returns only the template metadata."),
        req_body=None,
        resp_body=DETAILS_SAMPLE,
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC)],
        req_rows=None,
        resp_rows=ERROR_ROWS + DETAILS_ROWS,
        status="200",
    ),
    dict(
        title="Add Template Field",
        donor="Add Digital Signature Field",
        method="POST",
        path="{baseURL}/v1/templates/{id}/documents/{documentId}/fields/{fieldType}",
        desc=("Business applications can use this service API to place a new field "
              "on a document inside a personal/enterprise template. The field is "
              "stored in the template and is recreated on every workflow created "
              "from it."),
        req_body=FIELD_SAMPLE,
        resp_body=["{}"],
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC),
                   ("documentId", "Integer", "M", DOC_ID_DESC),
                   ("fieldType", "String", "M", FIELD_TYPE_DESC)],
        req_rows=FIELD_ROWS,
        resp_rows=ERROR_ROWS,
        status="200",
    ),
    dict(
        title="Update Template Field",
        donor="Update Digital Signature Field",
        method="PUT",
        path="{baseURL}/v1/templates/{id}/documents/{documentId}/fields/{fieldType}",
        desc=("Business applications can use this service API to update a field "
              "already placed on a document inside a personal/enterprise template. "
              "Use renamed_as to change the name of the field."),
        req_body=FIELD_SAMPLE,
        resp_body=["{}"],
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC),
                   ("documentId", "Integer", "M", DOC_ID_DESC),
                   ("fieldType", "String", "M", FIELD_TYPE_DESC)],
        req_rows=FIELD_ROWS,
        resp_rows=ERROR_ROWS,
        status="200",
    ),
]

# (section title, old path fragment, new path fragment). The backend reads
# `.../{pageNo}/{recordPerPage}`; these three documented it the other way round.
PATH_FIXES = [
    ("Get Groups", "{recordPerPage}/{pageNo}", "{pageNo}/{recordPerPage}"),
    ("Get Group", "{recordPerPage}/{pageNo}", "{pageNo}/{recordPerPage}"),
    ("Get Templates",
     "/v1/settings/templates/{recordPerPage}/{pageNo}",
     "/v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}"),
]


def fix_path(body, title, old, new):
    """Rewrite a documented URL in place, in every copy of its request frame."""
    kids = list(body)
    lo = st.find_heading(kids, 3, title)
    s, e = st.section_range(kids, lo)
    hits = 0
    for el in kids[s:e]:
        for txbx in el.iter(st.qn("txbxContent")):
            paras = txbx.findall(st.qn("p"))
            if not paras:
                continue
            line = st.text_of(paras[0])
            if "{baseURL}" not in line or old not in line:
                continue
            runs = [r for r in paras[0].findall(st.qn("r"))
                    if r.find(st.qn("t")) is not None]
            url = "".join(st.text_of(r) for r in runs[2:]).replace(old, new)
            t = runs[2].find(st.qn("t"))
            t.text = url
            t.set(st.XML_SPACE, "preserve")
            for r in runs[3:]:
                rt = r.find(st.qn("t"))
                if rt is not None and rt.text:
                    rt.text = ""
            hits += 1
    if not hits:
        raise LookupError("no frame carrying %r under %r" % (old, title))
    return hits


def main():
    tree, token = docx_guard.load(DOC)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    anchor = "Delete Template Workflow Participant"
    for spec in SECTIONS:
        clone, info = build(root, body, spec)
        st.insert_after_section(body, list(body), anchor, clone, level=3)
        anchor = spec["title"]
        print("+ %-24s from %-32s (%d elements)"
              % (spec["title"], spec["donor"], info["elements"]))

    for title, old, new in PATH_FIXES:
        n = fix_path(body, title, old, new)
        print("~ %-24s %s -> %s  (%d frame copies)" % (title, old, new, n))

    # Get Templates gains a path parameter it never documented.
    kids = list(body)
    lo = st.find_heading(kids, 3, "Get Templates")
    s, e = st.section_range(kids, lo)
    tbl = st.find_table(kids[s:e], "Path Attributes")
    fill_table(tbl, [
        ("templateType", "String", "M",
         "Which templates to return. Possible values are: + PERSONAL + ENTERPRISE + ALL"),
        ("pageNo", "Integer", "M", "The page of results to return, starting at 1."),
        ("recordPerPage", "Integer", "M", "The number of records to return per page."),
    ])
    print("~ Get Templates           Path Attributes rebuilt (templateType/pageNo/recordPerPage)")

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
        print("backed up -> %s" % os.path.basename(BACKUP))
    st.pack(UNPACKED, OUT_DOCX)
    print("packed    -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
