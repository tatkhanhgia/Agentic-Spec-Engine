#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add the `Redact Field` and `Elabor Field` pairs to Document Preparation.

Both field types are reachable but undocumented. `redaction` is registered in
`FieldRegistryBootstrap`; `elabor` is a special case in
`ParseToField.parseToFieldV3`. The routes are the wildcards
`POST|PUT /v1/packages/[0-9]*/documents/[0-9]*/fields/.*`, so Add and Update work
for both even though Postman only stores the Add request. Documenting the full
pair was approved by the project owner on 2026-07-27.

Donors were picked from contract/from_doc.json by matching attribute shape:

* Redact -> `Add/Update Qrypto Field`. `RedactFieldAttribute` adds a single
  object property to `BasicFieldAttribute`, exactly like the Qrypto donor
  (`items`), so the donor already carries the eight basic rows and needs one row
  renamed and one dropped.
* Elabor -> `Add/Update TextBox Field`. `ElaborFieldAttribute` declares
  maxLength / color / alignment / readOnly / multiline / font, which is the
  TextBox row set minus value, format_type, place_holder and tool_tip_text.

Facts come from the backend:

* Redact - `redact_policy` (`@JsonProperty` on `RedactFieldAttribute`) holds
  `creator` and `others`; both take `pre_process` / `post_process` of enum
  `RedactionPolicyType` = SHOW | REDACT | SHOW_ALL | REDACT_ALL, and `others`
  also takes `remark`. `RedactFieldResolver.check()` returns an empty response,
  so nothing in it is mandatory.
* Redact `type` is not documented, matching its sibling sections: for a registry
  field `FieldParseTemplate.parseAndValidate` falls back to the resolver's own
  `parentType()` when the payload omits it.
* Elabor - `type` is mandatory on create and optional on update.
  `ElaborFieldResolver` validates it only when `!isUpdate`, raising
  SUBCODE_ELABOR_MISSING_TYPE or SUBCODE_ELABOR_INVALID_TYPE, and requires the
  type to be a descendant of the `ELABOR_FIELD` sub-group.
* Elabor is documented with **snake_case** keys - `max_length`, `read_only`,
  `multiline` - per the project naming standard. `ElaborFieldAttribute` currently
  declares those properties with no `@JsonProperty`, so Jackson falls back to the
  bean names `maxLength` / `readOnly`, and `@JsonIgnoreProperties(ignoreUnknown =
  true)` makes the mismatch silent: a correctly-named key is dropped and the
  default applied, with no error. **The backend needs `@JsonProperty` added to
  those three properties.** The document states the standard; the code is the
  side that is wrong, the same call made for `DELETE /v1/tokens`.
* `alignment` is the bean name of the `AbstractAlignment` property. Sibling
  sections call the same object `align_option` because their attribute classes
  annotate it. That inconsistency is left as the backend has it and reported
  rather than silently harmonised.
* the alignment object itself uses `horizontal_alignment` /
  `vertical_alignment` (`@JsonProperty` on `AbstractAlignment`). Postman sends
  `horizontal` / `vertical`, which the parser ignores - a Postman defect, so the
  sample here follows the code.

Sample values come from the Postman collection where it has them
(`FULL_NAME_ELABOR`, the redaction policy shape), never invented.

Usage:
    python scripts/oneoff/add_redact_elabor_sections.py [--dry-run]
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

ANCHOR = "Update VNeID Field"      # the pair added immediately before these

REDACT_POLICY_DESC = (
    "Redaction policy of the field. It carries a creator object and an others object. "
    "Each of them takes pre_process and post_process, whose possible values are SHOW, "
    "REDACT, SHOW_ALL and REDACT_ALL. The others object also takes remark.")

ELABOR_TYPE_ADD = ("The Elabor attribute that the field is populated with. The value must be "
                   "a field type that belongs to the ELABOR_FIELD sub-group. The available "
                   "types are returned by the Get Field Types API.")
ELABOR_TYPE_UPDATE = ("The Elabor attribute that the field is populated with. The type is only "
                      "changed when this parameter is provided. The available types are "
                      "returned by the Get Field Types API.")

DIMENSION_BLOCK = [
    '  "dimension": {',
    '    "x": 0,',
    '    "y": 0,',
    '    "width": 100,',
    '    "height": 100',
    "  }",
]

REDACT_BODY_TAIL = [
    '  "redact_policy": {',
    '    "creator": {',
    '      "pre_process": "SHOW",',
    '      "post_process": "SHOW_ALL"',
    "    },",
    '    "others": {',
    '      "pre_process": "REDACT",',
    '      "post_process": "REDACT_ALL",',
    '      "remark": [1]',
    "    }",
    "  }",
    "}",
]

ELABOR_BODY_TAIL = [
    '  "max_length": 100,',
    '  "color": "BLACK",',
    '  "read_only": false,',
    '  "multiline": false,',
    '  "alignment": {',
    '    "horizontal_alignment": "LEFT",',
    '    "vertical_alignment": "MIDDLE"',
    "  },",
    '  "font": {',
    '    "name": "Arial",',
    '    "size": 12',
    "  },",
] + DIMENSION_BLOCK + ["}"]

SPECS = [
    {
        "donor": "Add Qrypto Field", "heading": "Add Redact Field", "method": "POST ",
        "old_path": "/fields/qrypto", "new_path": "/fields/redaction",
        "prose": ("Business applications can use this service API to add a new redact field "
                  "into the workflow."),
        "body": ["{", '  "field_name": "Redact field 1",', '  "page": 1,',
                 '  "visible_enabled": true,'] + DIMENSION_BLOCK[:-1] + ["  },"] + REDACT_BODY_TAIL,
        "rename": [("is_transparent", "redact_policy", "redactPolicy", "O", REDACT_POLICY_DESC)],
        "drop": ["items"], "append": [],
    },
    {
        "donor": "Update Qrypto Field", "heading": "Update Redact Field", "method": "PUT ",
        "old_path": "/fields/qrypto", "new_path": "/fields/redaction",
        "prose": ("Business applications can use this service API to update a redact field "
                  "in the workflow."),
        "body": ["{", '  "field_name": "Redact field 1",', '  "renamed_as": "Redact field 2",',
                 '  "page": 1,', '  "visible_enabled": true,'] + DIMENSION_BLOCK[:-1] + ["  },"]
                + REDACT_BODY_TAIL,
        "rename": [("is_transparent", "redact_policy", "redactPolicy", "O", REDACT_POLICY_DESC)],
        "drop": ["items"], "append": [],
    },
    {
        "donor": "Add TextBox Field", "heading": "Add Elabor Field", "method": "POST ",
        "old_path": "/fields/text", "new_path": "/fields/elabor",
        "prose": ("Business applications can use this service API to add a new Elabor field "
                  "into the workflow."),
        "body": ["{", '  "field_name": "Elabor field 1",', '  "page": 1,',
                 '  "type": "FULL_NAME_ELABOR",', '  "visible_enabled": true,'] + ELABOR_BODY_TAIL,
        "rename": [("align_option", "alignment", None, None, None),
                   ("type", None, None, None, ELABOR_TYPE_ADD)],
        "drop": ["value", "format_type", "place_holder", "tool_tip_text"], "append": [],
    },
    {
        "donor": "Update TextBox Field", "heading": "Update Elabor Field", "method": "PUT ",
        "old_path": "/fields/text", "new_path": "/fields/elabor",
        "prose": ("Business applications can use this service API to update an Elabor field "
                  "in the workflow."),
        "body": ["{", '  "field_name": "Elabor field 1",', '  "renamed_as": "Elabor field 2",',
                 '  "page": 1,', '  "type": "FULL_NAME_ELABOR",',
                 '  "visible_enabled": true,'] + ELABOR_BODY_TAIL,
        "rename": [("align_option", "alignment", None, None, None)],
        "drop": ["value", "format_type", "place_holder", "tool_tip_text"],
        "append": [("type", "String", "O", ELABOR_TYPE_UPDATE)],
    },
]


def build(root, kids, spec):
    clone, info = st.clone_section(root, kids, spec["donor"])

    st.set_para_text(clone[0], spec["heading"])
    st.set_para_text(clone[1], spec["prose"])

    frames = 0
    for txbx, paras in st.each_frame(clone, spec["method"]):
        st.set_path_suffix(paras[0], spec["old_path"], spec["new_path"])
        st.replace_frame_body(txbx, paras, spec["body"])
        frames += 1
    if frames == 0:
        raise RuntimeError("no %s frame in clone of %r" % (spec["method"].strip(), spec["donor"]))

    tbl = st.find_table(clone, "Request Attributes")
    before = len(st.data_rows(tbl))
    for name, new_name, typ, presence, desc in spec["rename"]:
        st.set_row(tbl, name, typ=typ, presence=presence, description=desc, new_name=new_name)
    for name in spec["drop"]:
        st.drop_row(tbl, name)
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
