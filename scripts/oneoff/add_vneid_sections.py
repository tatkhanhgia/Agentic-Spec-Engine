#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add `Add VNeID Field` and `Update VNeID Field` to Document Preparation.

Both endpoints exist in the backend but were never documented. `vneid` is handled
by `ParseToField.parseToFieldV3` as a special case before the resolver registry is
consulted, and the routes themselves are the wildcards
`POST|PUT /v1/packages/[0-9]*/documents/[0-9]*/fields/.*`, so Add and Update are
both reachable. The project owner approved documenting the full pair on 2026-07-27.

Donors: `Add TextBox Field` and `Update TextBox Field`. They were chosen from
contract/from_doc.json because `VNeIDFieldAttribute` extends the same
`BasicFieldAttribute` and declares the same properties as the text field
(`value`, `read_only`, `multiline`, `max_length`, `format_type`, `font`,
`place_holder`, `tool_tip_text`, `align_option`, `color`), so the donor tables
already carry the right 18 rows and only need one extra row each.

Facts come from the backend, not from inference:
  * `type` is mandatory on Add and optional on Update - `VNeIDFieldResolver`
    validates it only when `!isUpdate`, raising SUBCODE_VNEID_MISSING_TYPE or
    SUBCODE_VNEID_INVALID_TYPE, and otherwise applies it merely if non-null.
  * `loaded_from_vneid` - `@JsonProperty("loaded_from_vneid")` on
    `VNeIDFieldAttribute`. Postman spells the key with a trailing space; that is a
    Postman defect, the property name has none.
  * the valid `type` values are not listed. They are loaded from the database by
    `Resources.reloadFieldTypes()`, so no static list exists in the source - the
    document points at the `Get Field Types` API, which is what returns them.
  * `max_length`, `alignment`, `color`, `read_only`, `multiline` and `font` all get
    defaults in the resolver's initialisation lambda, so they stay optional.

Sample values are taken from the Postman collection, except `x-language-name`,
which Postman sends as `vn` - not a valid language code under template 2.8.7. The
donor frame already carries `en`.

Usage:
    python scripts/oneoff/add_vneid_sections.py [--dry-run]
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

ANCHOR = "Update RadioBox Field Version 2"        # last section of Document Preparation

TYPE_DESC_ADD = ("The VNeID attribute that the field is populated with. The value must be "
                 "a field type that belongs to the VNeID group. The available types are "
                 "returned by the Get Field Types API.")
TYPE_DESC_UPDATE = ("The VNeID attribute that the field is populated with. The type is only "
                    "changed when this parameter is provided. The available types are "
                    "returned by the Get Field Types API.")
LOADED_DESC = ("True if the value of the field is taken from the VNeID application instead "
               "of being entered manually.")

FONT_AND_DIMENSION = [
    '  "font": {',
    '    "name": "Arial",',
    '    "size": 12,',
    '    "embedded_size": "String"',
    "  },",
    '  "dimension": {',
    '    "x": 0,',
    '    "y": 0,',
    '    "width": 100,',
    '    "height": 100',
    "  }",
    "}",
]

SPECS = [
    {
        "donor": "Add TextBox Field",
        "heading": "Add VNeID Field",
        "prose": ("Business applications can use this service API to add a new VNeID field "
                  "into the workflow."),
        "method": "POST ",
        "body": [
            "{",
            '  "field_name": "VNeID field 1",',
            '  "page": 1,',
            '  "type": "DATE_OF_EXPIRY",',
            '  "loaded_from_vneid": true,',
            '  "value": "String",',
            '  "place_holder": "DATE_OF_EXPIRY",',
            '  "max_length": 100,',
            '  "format_type": "DATE",',
            '  "multiline": false,',
        ] + FONT_AND_DIMENSION,
        "retype": ("type", None, None, TYPE_DESC_ADD),
        "append": [("loaded_from_vneid", "Boolean", "O", LOADED_DESC)],
    },
    {
        "donor": "Update TextBox Field",
        "heading": "Update VNeID Field",
        "prose": ("Business applications can use this service API to update a VNeID field "
                  "in the workflow."),
        "method": "PUT ",
        "body": [
            "{",
            '  "field_name": "VNeID field 1",',
            '  "renamed_as": "VNeID field 2",',
            '  "page": 1,',
            '  "type": "GENDER",',
            '  "loaded_from_vneid": true,',
            '  "value": "String",',
            '  "place_holder": "Gender",',
            '  "max_length": 100,',
            '  "format_type": "TEXT",',
            '  "multiline": false,',
        ] + FONT_AND_DIMENSION,
        "retype": None,
        "append": [("type", "String", "O", TYPE_DESC_UPDATE),
                   ("loaded_from_vneid", "Boolean", "O", LOADED_DESC)],
    },
]


def build(root, kids, spec):
    clone, info = st.clone_section(root, kids, spec["donor"])

    st.set_para_text(clone[0], spec["heading"])
    st.set_para_text(clone[1], spec["prose"])

    frames = 0
    for txbx, paras in st.each_frame(clone, spec["method"]):
        # only the run holding the path suffix is touched, so the method badge and
        # the {baseURL} colouring survive untouched
        st.set_path_suffix(paras[0], "/fields/text", "/fields/vneid")
        st.replace_frame_body(txbx, paras, spec["body"])
        frames += 1
    if frames == 0:
        raise RuntimeError("no %s frame found in the clone of %r"
                           % (spec["method"].strip(), spec["donor"]))

    tbl = st.find_table(clone, "Request Attributes")
    before = len(st.data_rows(tbl))
    if spec["retype"]:
        name, typ, presence, desc = spec["retype"]
        st.set_row(tbl, name, typ, presence, desc)
    for row in spec["append"]:
        st.append_row(tbl, *row)
    after = len(st.data_rows(tbl))

    print("  %-22s <- %-22s  elements=%d frames=%d rows %d->%d  bm-%d docPr+%d vml+%d"
          % (spec["heading"], spec["donor"], info["elements"], frames, before, after,
             info["bookmarks_stripped"], info["docpr_reassigned"], info["vml_reassigned"]))
    return clone


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(st.qn("body"))
    kids = list(body)

    for spec in SPECS:
        try:
            st.find_heading(kids, 3, spec["heading"])
        except LookupError:
            pass
        else:
            sys.exit("%r already exists - nothing to do" % spec["heading"])

    built = []
    for spec in SPECS:
        built.extend(build(root, kids, spec))

    st.insert_after_section(body, kids, ANCHOR, built)
    print("inserted %d elements after %r" % (len(built), ANCHOR))

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
