#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add `Get Field Types` (GET /v1/documents/fields) to Document Preparation.

The route exists in the backend - `FieldController` `validURLs` entry 3,
`.*documents/fields`, dispatched to `FieldService.getFieldType` with
`loggingObject.apiName = "Get Field Types"` - and Postman calls it, but the
document never described it. Several sections now refer to it as the place that
lists the valid `type` values, because those values are loaded from the database
by `Resources.reloadFieldTypes()` and therefore cannot be enumerated in prose.

Donor: `Get list of Category`. It was picked from contract/from_doc.json as the
GET section that already has the target shape - no path parameters, so no Path
Attributes table, the same three request headers, and a success frame that
returns a list under a single top-level key.

Response shape comes from `GetFieldTypesSerializer`, which writes the JSON by
hand rather than serialising the DTO, so the DTO's own getters are not the
contract. Its `writeStringField` / `writeFieldName` calls give exactly:
`field_types[] -> {group, items[] -> {id, type_name, type_id, parent_type,
required, ordinary_id, remark, children[]}}`. The serializer's own comment states
it "outputs all field names in snake_case format", and the group values are the
`FieldTypeGroup` constants COMMON, SYSTEM and INDIVIDUAL - the class javadoc
still says USER, which the enum contradicts.

Following the convention of `Get list of Category` and `Get Document Fields`,
only the top-level response key gets a table row; the nested shape is shown in
the sample frame.

Usage:
    python scripts/oneoff/add_get_field_types_section.py [--dry-run]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

DONOR = "Get list of Category"
ANCHOR = "Get Document Fields"
HEADING = "Get Field Types"

PROSE = ("Using this API to get every field type available in the system, grouped by "
         "SYSTEM, INDIVIDUAL and COMMON. The API requires a user access token; a client "
         "credential token is rejected.")

OLD_PATH = "/v1/category"
NEW_PATH = "/v1/documents/fields"

RESPONSE_BODY = [
    "{",
    '"field_types":[',
    "{",
    '"group": "SYSTEM",',
    '"items":[',
    "{",
    '"id": "TEXTBOX",',
    '"type_name": "TEXTBOX",',
    '"type_id": 1,',
    '"parent_type": "TEXTBOX",',
    '"required": true,',
    '"ordinary_id": 1,',
    '"remark": "String",',
    '"children":[…]',
    "},…",
    "]",
    "},…",
    "]",
    "}",
]

FIELD_TYPES_DESC = ("Array of field type groups. Each element carries the group name and the "
                    "items that belong to it, and every item may carry a children array of "
                    "the same shape.")


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))
    kids = list(body)

    try:
        st.find_heading(kids, 3, HEADING)
    except LookupError:
        pass
    else:
        sys.exit("%r already exists - nothing to do" % HEADING)

    clone, info = st.clone_section(root, kids, DONOR)
    st.set_para_text(clone[0], HEADING)
    st.set_para_text(clone[1], PROSE)

    # request frame: only the run carrying the path is touched, so the GET badge keeps
    # its own run and its shading stays on the method token
    req = 0
    for txbx, paras in st.each_frame(clone, "GET "):
        st.set_path_suffix(paras[0], OLD_PATH, NEW_PATH)
        req += 1

    # success frame only - the 400 frame carries the shared error body
    resp = 0
    for txbx, paras in st.each_frame(clone, "Http Code: 200"):
        st.replace_frame_body(txbx, paras, RESPONSE_BODY)
        resp += 1

    if not req or not resp:
        sys.exit("expected both frames in the clone, got req=%d resp=%d" % (req, resp))

    tbl = st.find_table(clone, "Response Attributes")
    st.set_row(tbl, "categories", new_name="field_types", typ="List<fieldTypeGroup>",
               presence="M", description=FIELD_TYPES_DESC)

    print("%-18s <- %-22s elements=%d req-frames=%d resp-frames=%d bm-%d docPr+%d vml+%d"
          % (HEADING, DONOR, info["elements"], req, resp,
             info["bookmarks_stripped"], info["docpr_reassigned"], info["vml_reassigned"]))

    st.insert_after_section(body, kids, ANCHOR, clone)
    print("inserted after %r" % ANCHOR)

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
