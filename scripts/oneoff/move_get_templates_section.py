#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Move `Get Templates` into the Template chapter and correct its content.

Owner decision (2026-07-27): the call is a Template operation, so it belongs in
the Template chapter rather than under Personal Settings, where the module was
split in two.

Three content corrections go with the move, all read from the backend:

* `x-search-text` and `x-total-records` request headers. Both are read by the
  code path (`SearchTextHeaderUtils.decodeNullable`, `getPagination`) and both
  are confirmed by the SigningHub 10.0 manual. `x-total-records` has a
  behavioural side effect worth stating: when it is present `getPagination`
  zeroes the page number and row count, so the two path parameters are ignored.
* the `x-total-records` response header, which `TemplateService` sets to the
  total number of matching templates.
* the success body, which is `GetListTemplateSerializer`'s
  `{"templates": [ ... ]}` - not the root array SigningHub returns.
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
from add_template_workflow_sections import ERROR_ROWS, fill_table, set_frame_body  # noqa: E402

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_get_templates_move.docx")

HEADER_ROWS = [
    ("x-language-name", "String", "O",
     "Display the error message in the specified language. Possible values are: + en + vi"),
    ("Authorization", "String", "M",
     "The authentication token (accessToken) required to access the API. Must be provided in the "
     "format “Bearer <accessToken>” Note: the accessToken is obtained after a successful authorization."),
    ("Content-Type", "String", "M",
     "The MIME type of the request body. Usually application/json"),
    ("x-search-text", "String", "O",
     "Filter the list by template name. The value must be Base64 encoded. Without it the entire list is returned."),
    ("x-total-records", "Integer", "O",
     "The number of records to return. When this header is present it overrides the pageNo and "
     "recordPerPage path parameters, which are then ignored."),
]

RESPONSE_ROWS = ERROR_ROWS + [
    ("templates", "List<TemplateListItem>", "M", "The templates matching the request."),
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

BODY = [
    "{",
    '  "templates": [',
    "    {",
    '      "template_id": 1,',
    '      "template_name": "Employment Contract Template",',
    '      "template_public": false,',
    '      "read_only": false,',
    '      "created_at": "2026-07-27T09:15:00Z",',
    '      "description": "Standard employment contract",',
    '      "enterprise": true,',
    '      "documents": [',
    "        {",
    '          "id": 12,',
    '          "document_id": 12,',
    '          "name": "contract.pdf"',
    "        }",
    "      ],",
    '      "smart_form": {',
    '        "active_count": 0,',
    '        "total_count": 0',
    "      }",
    "    }",
    "  ]",
    "}",
]

DESC = ("Business applications can use this service API to get the list of "
        "personal and/or enterprise templates available to the logged-in user. "
        "Use the templateType path parameter to choose which of the two, or ALL "
        "for both. Requesting ENTERPRISE with a client credentials token is "
        "rejected: the caller must be logged in as an enterprise user.")


def main():
    tree, token = docx_guard.load(DOC)
    body = tree.getroot().find(st.qn("body"))

    kids = list(body)
    lo = st.find_heading(kids, 3, "Get Templates")
    s, e = st.section_range(kids, lo)
    section = kids[s:e]

    # description
    for el in section[1:]:
        if el.tag == st.qn("p") and st.text_of(el).strip():
            st.set_para_text(el, DESC)
            break

    fill_table(st.find_table(section, "Header Request Attributes"), HEADER_ROWS)
    fill_table(st.find_table(section, "Response Attributes"), RESPONSE_ROWS)

    n = 0
    for txbx, paras in st.each_frame(section, "Http Code: 200"):
        set_frame_body(txbx, paras, BODY)
        n += 1
    print("~ success frame rewritten in %d copies" % n)

    # lxml re-parents on addnext, so this moves the section rather than copying it
    st.insert_after_section(body, list(body), "Save Workflow as Template",
                            section, level=3)
    print("~ moved Get Templates: Personal Settings -> Template")

    docx_guard.save(tree, DOC, token)
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
    st.pack(UNPACKED, OUT_DOCX)
    print("packed -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
