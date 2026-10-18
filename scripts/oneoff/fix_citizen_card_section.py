#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Field-level audit fixes for `Get Information of Citizen card` (2026-07-27).

Backend evidence: RSSPService.getInformationWithCitizenCard reads ONLY `value`
from the payload (Utils.getFromJson_("value", ...)) and hard-codes
IdentityDocumentType.CITIZEN_CARD - the `document_type` key the document (and
Postman) sends is silently ignored. Response = GetSubjectData mapper wrapping
eid.restful.objects.PersonalInformations, serialized with an explicit
SNAKE_CASE mapper.

1. Description was a copy-paste of the Register section ("register new
   Certificate via RSSP") - rewritten to describe what the API does.
2. `document_type` removed from the request table and the sample body -
   documenting it ships a promise the API does not keep.
3. `value` description updated now that document_type no longer names the
   document kind.
4. Sample 200 key `"portrait "` (trailing space inside the key) -> `"portrait"`.
5. PersonalInformation object row 10 Type `Stringh` -> `String`.

NOT changed: the document spells `personal_identification` while the backend
field is the typo `personalIdentificatetion` (no @JsonProperty), so the wire key
is `personal_identificatetion`. Owner rule: camel/typo keys in the backend are
defects to report - recorded in template 5.2 and TECHNICAL_DEBT.md section 7.

Usage:
    python scripts/oneoff/fix_citizen_card_section.py [--dry-run]
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

qn = st.qn


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guarded_load(DOC_XML)
    body = tree.getroot().find(qn("body"))
    kids = list(body)

    a, b = st.section_range(
        kids, st.find_heading(kids, 3, "Get Information of Citizen card"))
    els = kids[a:b]

    # 1. description
    st.set_para_text(els[1], (
        "Business applications can use this service API to retrieve the citizen "
        "card information of a user from the identity provider."))

    # 2-3. request table
    tbl = st.find_table(els, "Request Attributes")
    st.drop_row(tbl, "document_type")
    st.set_row(tbl, "value", description="Citizen card number of the user to look up.")

    # 2. sample request body
    for txbx, _paras in st.each_frame(els, "POST "):
        st.replace_frame_body(txbx, txbx.findall(qn("p")), [
            "{",
            '  "value": ""',
            "}",
        ])

    # 4. sample 200: fix the "portrait " key (trailing space inside the key)
    n = 0
    for _txbx, paras in st.each_frame(els, "Http Code: 200"):
        for p in paras:
            if '"portrait "' in st.text_of(p):
                n += st.set_path_suffix(p, '"portrait "', '"portrait"')
    print("citizen card section: desc, request table, samples fixed (portrait runs=%d)" % n)

    # 5. PersonalInformation object: Stringh -> String
    i2 = st.find_heading(kids, 2, "PersonalInformation")
    a2, b2 = st.section_range(kids, i2)
    tbl = st.find_table(kids[a2:b2], "PersonalInformation")
    st.set_row(tbl, "issuance_date", typ="String")
    print("PersonalInformation object: issuance_date type Stringh -> String")

    if dry:
        print("--dry-run: nothing written")
        return
    try:
        guarded_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("\nREFUSED TO WRITE\n%s" % ex)
    st.pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
