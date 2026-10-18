#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add the Validation Management module (5 endpoints) to the specification.

All five endpoints exist in ValidationController.validURLs, are wired to real
service code (no stubs) and match the Postman folder "Validation Management"
5-for-5; the user approved adding them on 2026-07-27 with:
  - full RELATIVE OBJECT depth for the validation report,
  - module placed after Template, before ADMIN APIs,
  - actual backend behaviour documented (quirks recorded in TECHNICAL_DEBT.md),
  - GET /v1/account/dashboard/validations left for Account Management.

Facts from the backend, read-only:

1. Add Validation - POST /v1/validations
   Async: uploads the base64 document to FMS, inserts the validation row and
   returns {"validation_id": N} immediately; a background task performs the
   validation. Body: name (max length limited by config), validation_policy
   (ADVANCED_QUALIFIED_SEALS default | QUALIFIED_SEALS), liability_level
   (BASIC default | ADVANCED, alias "liability"), file (base64).
   Client-credential tokens rejected (BLOCK_CLIENT_CREDENTIAL).

2. Get List Validation - GET /v1/validations[/{page_no}/{records_per_page}]
   x-search-text (base64) filters by name; x-total-records header caps rows.
   Response {"validations": [...]} + X-Total-Records header
   (GetListValidationSerialize).

3. Get Report of Validation - GET /v1/validations/{validationId}/reports
   202 (code 28 subcode 8) while status PROCESSING; 400 subcode 7 when the
   validation belongs to another user; otherwise the DSS report serialized
   SNAKE_CASE/NON_NULL: overview, signatures[], seals[], details, file.
   basic/detail file data are nulled before serialization; the original
   document is embedded base64 in file.original_file_data.

4. Download Reports - GET /v1/validations/reports/{reportId}/{reportType}
   reportType in detailed|diagnostic|simple. detailed -> PDF
   (Content-Type application/pdf), diagnostic -> XML, simple -> PDF file but
   the controller sets Content-Type application/xml and the service names the
   attachment DiagnosticReport.pdf (backend quirks, see TECHNICAL_DEBT.md).
   The service performs no token verification (quirk, TECHNICAL_DEBT.md);
   Authorization is still documented M as every client sends it.

5. Delete Validation - DELETE /v1/validations/{validationId}
   Empty {} success. No ownership check in the service (TECHNICAL_DEBT.md).

Error domain: WorkflowConstant.CODE_VALIDATION = 28, subcodes 1-11 in
ValidationErrorSubCode; error_code is rendered "28.<subcode>"
(ResponseMessageFactory codeKey).

annotation_dimension / annotation_page are NOT documented: EnrichDSSResponse
looks for a "signature_scope" array but the serializer emits "scope", so the
fields are never added (TECHNICAL_DEBT.md).

Sections are cloned from same-shape donors; bookmarks are stripped and
DrawingML/VML ids re-issued (section_tools). Eleven relative objects are
appended to the RELATIVE OBJECT chapter, each with a fresh named bookmark so
composite Type cells can link. The TOC is refreshed by the user with F9.

Usage:
    python scripts/oneoff/add_validation_management_sections.py [--dry-run]
"""
import copy
import os
import re
import shutil
import sys

from lxml import etree  # noqa: F401  (parsed via docx_guard)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guard_load, save as guard_save, ConcurrentModification
import section_tools as st
from section_tools import qn

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted"
    "_backup_pre_validation_management.docx")

HEADINGS_RE = re.compile(r"Heading\d$")
HEADER_LINE = re.compile(
    r"^(GET |POST |PUT |DELETE |Http Code:|Content-[Tt]ype\s*:|Authorization\s*:|x-[a-z0-9-]+\s*:)")

BLOCK_CC = ("This API requires a user access token; a token obtained through "
            "Client Credentials Authentication is rejected.")

# anchors of the new relative objects (all verified free in the document)
A_VALIDATION = "_Validation_Object"
A_OVERVIEW = "_ValidationOverview"
A_SIGNATURE = "_Signature_Validation"
A_CERT = "_Certificate_Validation"
A_SCOPE = "_Scope_Validation"
A_DIGEST = "_Digest_Validation"
A_META = "_MetaData_Validation"
A_MESSAGE = "_Message_Validation"
A_POLICY = "_ValidationPolicy_Validation"
A_DETAILS = "_Details_Validation"
A_FILE = "_File_Validation"
A_DIMENSION = "_Dimension_1"  # existing object, reused


# ---------------------------------------------------------------------------
# frame helpers
#
# A response paragraph can hold SEVERAL frames (200 and 400 side by side).
# Each frame is one w:r holding an mc:AlternateContent whose Choice and
# Fallback each carry a txbxContent copy - so a frame run has exactly two
# txbxContent and every edit must hit both, but never a neighbouring frame.
# ---------------------------------------------------------------------------
def frame_runs(blocks, first_line_prefix):
    hits = []
    for el in blocks:
        if el.tag != qn("p"):
            continue
        for r in el.findall(qn("r")):
            txs = r.findall(".//" + qn("txbxContent"))
            if not txs:
                continue
            first = st.text_of(txs[0].findall(qn("p"))[0]).strip()
            if first.startswith(first_line_prefix):
                hits.append(r)
    return hits


def frame_copies(run):
    txs = run.findall(".//" + qn("txbxContent"))
    if len(txs) != 2:
        sys.exit("frame run has %d txbxContent copies, expected 2" % len(txs))
    return txs


def one(lst, what):
    if len(lst) != 1:
        sys.exit("expected exactly one %s, got %d" % (what, len(lst)))
    return lst[0]


def set_leaf_text(el, value):
    ts = [t for t in el.iter(qn("t"))]
    if not ts:
        sys.exit("element has no text run to rewrite")
    ts[0].text = value
    ts[0].set(st.XML_SPACE, "preserve")
    for extra in ts[1:]:
        extra.text = ""


def rewrite_body(run, body_lines):
    for tx in frame_copies(run):
        paras = tx.findall(qn("p"))
        body = [p for p in paras if not HEADER_LINE.match(st.text_of(p))]
        if not body:
            sys.exit("frame has no body paragraph to use as a style template")
        template = body[0]
        anchor = body[0]
        for extra in body[1:]:
            tx.remove(extra)
        for line in body_lines:
            new = copy.deepcopy(template)
            set_leaf_text(new, line)
            anchor.addnext(new)
            anchor = new
        tx.remove(template)


def drop_frame_lines_exact(run, texts):
    for tx in frame_copies(run):
        for p in list(tx.findall(qn("p"))):
            if st.text_of(p).strip() in texts:
                tx.remove(p)


def set_frame_line(run, prefix, new_text):
    hits = 0
    for tx in frame_copies(run):
        for p in tx.findall(qn("p")):
            if st.text_of(p).strip().startswith(prefix):
                set_leaf_text(p, new_text)
                hits += 1
                break
    if hits != 2:
        sys.exit("set_frame_line %r hit %d frame copies, expected 2" % (prefix, hits))


def set_url(run, new_url):
    """Rewrite the bold URL on the request frame's first line, keep the badge."""
    for tx in frame_copies(run):
        first = tx.findall(qn("p"))[0]
        done = False
        for r in first.findall(qn("r")):
            txt = "".join(t.text or "" for t in r.iter(qn("t"))).strip()
            if txt in ("GET", "POST", "PUT", "DELETE", ""):
                continue
            ts = list(r.iter(qn("t")))
            if not done:
                ts[0].text = new_url
                ts[0].set(st.XML_SPACE, "preserve")
                for extra in ts[1:]:
                    extra.text = ""
                done = True
            else:
                for t in ts:
                    t.text = ""
        if not done:
            sys.exit("request frame first line has no URL run to rewrite")


def set_description(blocks, text):
    seen = False
    for el in blocks[1:]:
        if el.tag != qn("p"):
            break
        t = st.text_of(el).strip()
        if t == "Sample Request":
            break
        if not t:
            continue
        if not seen:
            set_leaf_text(el, text)
            seen = True
        else:
            set_leaf_text(el, "")
    if not seen:
        sys.exit("no description paragraph found to overwrite")


# ---------------------------------------------------------------------------
# table helpers on top of section_tools
# ---------------------------------------------------------------------------
def rebuild_rows(tbl, rows):
    """Drop every data row, rebuild from the first donor row's formatting."""
    donors = st.data_rows(tbl)
    if not donors:
        sys.exit("table %r has no data row to use as a donor" % st.table_title(tbl))
    template = copy.deepcopy(donors[0])
    parent = donors[0].getparent()
    for tr in donors:
        parent.remove(tr)
    anchor = tbl.findall(qn("tr"))[-1]  # column-header row is last now
    for name, typ, presence, desc in rows:
        new = copy.deepcopy(template)
        tcs = new.findall(qn("tc"))
        st.set_cell_text(tcs[1], name)
        st.set_cell_text(tcs[2], typ)
        st.set_cell_text(tcs[3], presence)
        st.set_cell_text(tcs[4], desc)
        anchor.addnext(new)
        anchor = new
    st.renumber(tbl)


def row_type_cell(tbl, name):
    for tr in st.data_rows(tbl):
        tcs = tr.findall(qn("tc"))
        if st.cell_text(tcs[1]) == name:
            return tcs[2]
    sys.exit("no row named %r in %r" % (name, st.table_title(tbl)))


# ---------------------------------------------------------------------------
# donor lookups on the live body
# ---------------------------------------------------------------------------
def live_section(kids, title, level=3):
    start, end = st.section_range(kids, st.find_heading(kids, level, title))
    return kids[start:end]


def find_link_donor_cells(body):
    """A plain-label link cell (dimension) and a List<> wrapper cell (workflows)."""
    plain = wrapped = None
    for h in body.iter(qn("hyperlink")):
        label = st.text_of(h).strip()
        anchor = h.get(qn("anchor")) or ""
        p = h.getparent()
        tc = p.getparent()
        if tc is None or tc.tag != qn("tc"):
            continue
        if plain is None and anchor == A_DIMENSION and label == "dimension":
            plain = tc
        if wrapped is None and label and st.cell_text(tc).startswith("List<"):
            wrapped = tc
        if plain is not None and wrapped is not None:
            return plain, wrapped
    sys.exit("link donor cells not found (plain=%r wrapped=%r)" % (plain, wrapped))


# ---------------------------------------------------------------------------
# endpoint builders
# ---------------------------------------------------------------------------
def build_add_validation(root, kids):
    blocks, info = st.clone_section(root, kids, "Create Workflow From Template")
    st.set_text_keep_bookmarks(blocks[0], "Add Validation")
    set_description(blocks, (
        "Business applications can use this service API to submit a document for "
        "signature and seal validation. The document is sent in base64 format. The "
        "validation runs asynchronously: the call returns the validation ID "
        "immediately and the result is retrieved with Get Report of Validation. "
        + BLOCK_CC))

    req_frame = one(frame_runs(blocks, "POST "), "request frame")
    set_url(req_frame, "{baseURL}/v1/validations")
    rewrite_body(req_frame, [
        "{",
        '  "name": "contract.pdf",',
        '  "validation_policy": "ADVANCED_QUALIFIED_SEALS",',
        '  "liability_level": "BASIC",',
        '  "file": "Base64 encoded document"',
        "}",
    ])

    resp200 = one(frame_runs(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "validation_id": 123',
        "}",
    ])

    req = st.find_table(blocks, "Request Attributes")
    rebuild_rows(req, [
        ("name", "String", "M",
         "Display name of the validation, usually the file name of the document. "
         "The maximum length is limited by the server configuration; a longer "
         "name is rejected."),
        ("validation_policy", "String", "O",
         "Validation policy applied to the document. Possible values are: "
         "+ ADVANCED_QUALIFIED_SEALS (default) + QUALIFIED_SEALS"),
        ("liability_level", "String", "O",
         "Liability level applied to the validation. Possible values are: "
         "+ BASIC (default) + ADVANCED. The alias key liability is also accepted."),
        ("file", "String", "M",
         "Base64 encoded content of the document to validate."),
    ])

    resp = st.find_table(blocks, "Response Attributes")
    rebuild_rows(resp, [
        ("validation_id", "Integer", "M",
         "ID of the validation that was created. Use this ID with Get Report of "
         "Validation and Delete Validation."),
    ])
    return blocks


def build_get_list_validation(root, kids, wrapped_donor_tc):
    blocks, info = st.clone_section(root, kids, "Get All Workflow")
    st.set_text_keep_bookmarks(blocks[0], "Get List Validation")
    set_description(blocks, (
        "Business applications can use this service API to retrieve the validations "
        "of the current user. The x-search-text header can be used to filter the "
        "list by validation name. " + BLOCK_CC))

    req_frame = one(frame_runs(blocks, "GET "), "request frame")
    # template 2.8.9: URL path parameters are camelCase (the donor's
    # {page_no}/{records_per_page} spelling is existing debt, item 5)
    set_url(req_frame, "{baseURL}/v1/validations/{pageNo}/{recordsPerPage}")
    drop_frame_lines_exact(req_frame, {'x-search-text: "Hello"', "x-category-id: 1"})

    resp200 = one(frame_runs(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "validations": [',
        "    {",
        '      "validation_id": 123,',
        '      "validation_name": "contract.pdf",',
        '      "validation_status": "VALID",',
        '      "validation_date": "2026-07-27T10:15:30",',
        '      "total_valid_signatures": 1,',
        '      "total_valid_seals": 1,',
        '      "total_signatures": 1,',
        '      "total_seals": 1,',
        '      "detailed_validation_report": '
        '"/workflow/api/v1/validations/reports/{reportId}/detailed"',
        "    }",
        "  ]",
        "}",
    ])

    hdr = st.find_table(blocks, "Header Request Attributes")
    st.drop_row(hdr, "x-category-id")
    st.set_row(hdr, "x-search-text", description=(
        "The string name (after being encoded to base64) used to filter the "
        "validations by name."))
    st.set_row(hdr, "x-total-records", description=(
        "The number of records that the client wants to get. Note 1: when this "
        "parameter is defined, the parameters pageNo and recordsPerPage will "
        "be ignored. Note 2: Default value of this parameter is 100."))

    path = st.find_table(blocks, "Path Attributes")
    st.drop_row(path, "criteria")
    st.set_row(path, "page_no", new_name="pageNo")
    st.set_row(path, "records_per_page", new_name="recordsPerPage")

    resp = st.find_table(blocks, "Response Attributes")
    # keep the three standard error rows, replace the data row
    for tr in list(st.data_rows(resp)):
        name = st.cell_text(tr.findall(qn("tc"))[1])
        if name not in ("error_description", "error_code", "message"):
            tr.getparent().remove(tr)
    st.append_row(resp, "validations", "List<Validation>", "M",
                  "The list of validations of the current user.")
    st.set_type_link(row_type_cell(resp, "validations"), "List<Validation>",
                     A_VALIDATION, wrapped_donor_tc)
    return blocks


def build_get_report(root, kids, plain_donor_tc, wrapped_donor_tc):
    blocks, info = st.clone_section(root, kids, "Get Status of Import Excel")
    st.set_text_keep_bookmarks(blocks[0], "Get Report of Validation")
    set_description(blocks, (
        "Business applications can use this service API to retrieve the validation "
        "report of a document. The validation ID is provided in the resource URL. "
        "While the validation is still being processed the call returns HTTP code "
        "202; once finished it returns the full report with the overview, the "
        "validated signatures and seals, the report download links and the "
        "original document in base64 format. Only the user who created the "
        "validation can read its report. " + BLOCK_CC))

    req_frame = one(frame_runs(blocks, "GET "), "request frame")
    set_url(req_frame, "{baseURL}/v1/validations/{validationId}/reports")

    resp200 = one(frame_runs(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "overview": {',
        '    "total_signatures": 1,',
        '    "total_valid_signatures": 1,',
        '    "total_seals": 0,',
        '    "total_valid_seals": 0,',
        '    "validation_time": "2026-07-27T10:15:30",',
        '    "validation_policy": "ADVANCED_QUALIFIED_SEALS",',
        '    "liability_level": "BASIC",',
        '    "final_accepted": true',
        "  },",
        '  "signatures": [',
        "    {",
        '      "id": "String",',
        '      "field_name": "Signature1",',
        '      "type": "QUALIFIED_ELECTRONIC_SIGNATURE",',
        '      "format": "String",',
        '      "indication": "String",',
        '      "signature_valid": true,',
        '      "dss_valid": true,',
        '      "qualification": "String",',
        '      "qualification_accepted": true,',
        '      "final_accepted": true,',
        '      "signing_time": "String",',
        '      "meta_data": {"contact": "String", "reason": "String", '
        '"location": "String"},',
        '      "data_to_be_signed": {"digest_algorithm": "SHA256", '
        '"digest_value": "Base64 encoded digest"},',
        '      "signature_value": "Base64 encoded signature",',
        '      "signing_certificate": {"common_name": "String", '
        '"serial_number": "String", "value": "PEM encoded certificate"},',
        '      "scope": {"scope_name": "String", "name": "Signature1", '
        '"value": "String"},',
        '      "errors": [],',
        '      "warnings": [{"key": "String", "value": "String"}],',
        '      "page": 1,',
        '      "dimension": {"x": 0, "y": 0, "width": 0, "height": 0}',
        "    }",
        "  ],",
        '  "seals": [],',
        '  "details": {',
        '    "validation_detail_report_id": "String",',
        '    "validated_detail_document_hash": "String",',
        '    "validation_basic_report_id": "String",',
        '    "validated_basic_document_hash": "String",',
        '    "detailed_validation_report": '
        '"/workflow/api/v1/validations/reports/{reportId}/detailed",',
        '    "diagnostic_data": '
        '"/workflow/api/v1/validations/reports/{reportId}/diagnostic",',
        '    "basic_validation_report": '
        '"/workflow/api/v1/validations/reports/{reportId}/simple"',
        "  },",
        '  "file": {',
        '    "original_file_data": "Base64 encoded document",',
        '    "original_file_name": "contract.pdf",',
        '    "xml_diagnostic": "String"',
        "  }",
        "}",
    ])

    # 202 accepted frame while the validation is still processing.
    # The frame is one run inside the response paragraph, so the copy is
    # inserted next to the 200 run and its drawing ids re-issued.
    resp202 = copy.deepcopy(resp200)
    resp200.addnext(resp202)
    st.reassign_ids([resp202], *st.scan_id_space(root))
    set_frame_line(resp202, "Http Code:", "Http Code: 202")
    rewrite_body(resp202, [
        "{",
        '  "error_description": "Validation is still being processed",',
        '  "error_code": "28.8",',
        '  "message": "The validation is still being processed. '
        'Please try again later."',
        "}",
    ])

    path = st.find_table(blocks, "Path Attributes")
    rebuild_rows(path, [
        ("validationId", "Integer", "M",
         "The ID of the validation returned by Add Validation."),
    ])

    resp = st.find_table(blocks, "Response Attributes")
    rebuild_rows(resp, [
        ("overview", "Overview", "M", "Summary of the validation result."),
        ("signatures", "List<Signature>", "M",
         "Electronic signatures found in the document."),
        ("seals", "List<Seal>", "M",
         "Electronic seals found in the document. Each item has the same "
         "structure as a signature."),
        ("details", "Details", "M",
         "Report identifiers, document hashes and download links of the "
         "validation reports."),
        ("file", "File", "M",
         "The validated document and its diagnostic data."),
    ])
    st.set_type_link(row_type_cell(resp, "overview"), "Overview",
                     A_OVERVIEW, plain_donor_tc)
    st.set_type_link(row_type_cell(resp, "signatures"), "List<Signature>",
                     A_SIGNATURE, wrapped_donor_tc)
    st.set_type_link(row_type_cell(resp, "seals"), "List<Seal>",
                     A_SIGNATURE, wrapped_donor_tc)
    st.set_type_link(row_type_cell(resp, "details"), "Details",
                     A_DETAILS, plain_donor_tc)
    st.set_type_link(row_type_cell(resp, "file"), "File",
                     A_FILE, plain_donor_tc)
    return blocks


def build_download_reports(root, kids):
    blocks, info = st.clone_section(root, kids, "Download Attachment")
    st.set_text_keep_bookmarks(blocks[0], "Download Reports")
    set_description(blocks, (
        "Business applications can use this service API to download a validation "
        "report file. The report identifier is taken from the report links "
        "returned by Get Report of Validation. The detailed report is a PDF file "
        "returned with Content-Type application/pdf; the diagnostic report is an "
        "XML file returned with Content-Type application/xml; the simple report "
        "is a PDF file that is currently returned with Content-Type "
        "application/xml and file name DiagnosticReport.pdf."))

    req_frame = one(frame_runs(blocks, "GET "), "request frame")
    set_url(req_frame, "{baseURL}/v1/validations/reports/{reportId}/{reportType}")

    resp200 = one(frame_runs(blocks, "Http Code: 200"), "200 frame")
    set_frame_line(resp200, "Content-Type",
                   "Content-Type: application/pdf (detailed) or application/xml "
                   "(diagnostic, simple)")

    path = st.find_table(blocks, "Path Attributes")
    rebuild_rows(path, [
        ("reportId", "String", "M",
         "Identifier of the report, returned in the details links of Get Report "
         "of Validation: validation_detail_report_id for detailed and "
         "diagnostic, validation_basic_report_id for simple."),
        ("reportType", "String", "M",
         "Type of the report to download. Possible values are: + detailed "
         "(detailed validation report, PDF) + diagnostic (diagnostic data, XML) "
         "+ simple (simple validation report, PDF)"),
    ])
    return blocks


def build_delete_validation(root, kids):
    blocks, info = st.clone_section(root, kids, "Delete Workflow")
    st.set_text_keep_bookmarks(blocks[0], "Delete Validation")
    set_description(blocks, (
        "Business applications can use this service API to delete a validation. "
        "The validation ID is provided in the resource URL. " + BLOCK_CC))

    req_frame = one(frame_runs(blocks, "DELETE "), "request frame")
    set_url(req_frame, "{baseURL}/v1/validations/{validationId}")

    path = st.find_table(blocks, "Path Attributes")
    rebuild_rows(path, [
        ("validationId", "Integer", "M",
         "The ID of the validation returned by Add Validation."),
    ])
    return blocks


# ---------------------------------------------------------------------------
# relative objects
# ---------------------------------------------------------------------------
def add_heading_bookmark(heading_p, name, bm_id):
    pPr = heading_p.find(qn("pPr"))
    start = heading_p.makeelement(qn("bookmarkStart"),
                                  {qn("id"): str(bm_id), qn("name"): name})
    end = heading_p.makeelement(qn("bookmarkEnd"), {qn("id"): str(bm_id)})
    if pPr is not None:
        pPr.addnext(start)
    else:
        heading_p.insert(0, start)
    heading_p.append(end)


def build_object(root, kids, title, anchor, bm_id, rows,
                 links=None, plain_donor_tc=None, wrapped_donor_tc=None):
    blocks, info = st.clone_section(root, kids, "Dimension", level=2)
    st.set_text_keep_bookmarks(blocks[0], title)
    add_heading_bookmark(blocks[0], anchor, bm_id)
    tbl = st.find_table(blocks, "Response Attributes")
    rebuild_rows(tbl, rows)
    for name, label, target in (links or []):
        donor = wrapped_donor_tc if label.startswith("List<") else plain_donor_tc
        st.set_type_link(row_type_cell(tbl, name), label, target, donor)
    return blocks


OBJ_VALIDATION = [
    ("validation_id", "Integer", "M", "ID of the validation."),
    ("validation_name", "String", "M",
     "Name of the validation provided when it was created."),
    ("validation_status", "String", "M",
     "Status of the validation. Possible values are: + NOT_PROCESS "
     "+ NO_SIGNATURES_OR_SEALS_FOUND + WARNING + VALID + INVALID + PROCESSING"),
    ("validation_date", "String", "M",
     "Date and time of the validation in UTC."),
    ("total_valid_signatures", "Integer", "M",
     "Number of valid signatures found in the document."),
    ("total_valid_seals", "Integer", "M",
     "Number of valid seals found in the document."),
    ("total_signatures", "Integer", "M",
     "Total number of signatures found in the document."),
    ("total_seals", "Integer", "M",
     "Total number of seals found in the document."),
    ("detailed_validation_report", "String", "M",
     "Relative link to download the detailed validation report. Null until the "
     "validation has finished."),
]

OBJ_OVERVIEW = [
    ("total_signatures", "Integer", "M",
     "Total number of signatures found in the document."),
    ("total_valid_signatures", "Integer", "M",
     "Number of valid signatures found in the document."),
    ("total_seals", "Integer", "M",
     "Total number of seals found in the document."),
    ("total_valid_seals", "Integer", "M",
     "Number of valid seals found in the document."),
    ("validation_time", "String", "O",
     "Time of the validation reported by the validation service."),
    ("validation_policy", "String", "O",
     "Validation policy applied to the document. Possible values are: "
     "+ ADVANCED_QUALIFIED_SEALS + QUALIFIED_SEALS"),
    ("final_accepted", "Boolean", "M",
     "True when the overall validation result is accepted."),
    ("validation_mode", "String", "O",
     "Validation mode reported by the validation service."),
    ("multi_signature_policy", "String", "O",
     "Multi-signature policy applied by the validation service."),
    ("target_signature_id", "String", "O",
     "Identifier of the signature targeted by the validation, when a specific "
     "signature is validated."),
    ("liability_level", "String", "O",
     "Liability level applied to the validation. Possible values are: "
     "+ BASIC + ADVANCED"),
]

OBJ_SIGNATURE = [
    ("id", "String", "O", "Identifier of the signature in the validation report."),
    ("field_name", "String", "O", "Name of the signature field in the PDF."),
    ("type", "String", "O",
     "Type of the signature. Possible values are: "
     "+ ADVANCED_ELECTRONIC_SIGNATURE + ELECTRONIC_SEAL "
     "+ QUALIFIED_ELECTRONIC_SEAL + QUALIFIED_ELECTRONIC_SIGNATURE"),
    ("format", "String", "O",
     "Signature format reported by the validation service."),
    ("indication", "String", "O",
     "Main validation indication of the signature."),
    ("sub_indication", "String", "O",
     "Sub indication detailing the validation indication."),
    ("algorithm", "String", "O", "Signing algorithm of the signature."),
    ("signing_time", "String", "O", "Signing time recorded in the signature."),
    ("valid_from", "String", "O",
     "Start of the signing certificate validity period."),
    ("valid_to", "String", "O",
     "End of the signing certificate validity period."),
    ("meta_data", "MetaData", "O",
     "Contact, reason and location recorded when the document was signed."),
    ("validation_policy", "ValidationPolicy", "O",
     "Validation policy used to validate this signature."),
    ("data_to_be_signed", "Digest", "O", "Digest of the signed data."),
    ("signature_value", "String", "O", "Base64 encoded signature value."),
    ("signing_certificate", "Certificate", "O",
     "Certificate used to create the signature."),
    ("certificate_chain", "List<Object>", "O",
     "Raw certificate chain of the signing certificate as reported by the "
     "validation service (diagnostic data)."),
    ("filter", "String", "O", "PDF signature filter."),
    ("sub_filter", "String", "O", "PDF signature sub filter."),
    ("signature_valid", "Boolean", "M",
     "True when the cryptographic signature is valid."),
    ("dss_valid", "Boolean", "M",
     "True when the validation service accepted the signature."),
    ("qualification", "String", "O", "Qualification level of the signature."),
    ("qualification_accepted", "Boolean", "M",
     "True when the qualification level is accepted by the applied policy."),
    ("final_accepted", "Boolean", "M",
     "True when the signature is accepted overall."),
    ("qualification_reject_reason", "String", "O",
     "Reason the qualification was rejected, when applicable."),
    ("scope", "Scope", "O",
     "Scope of the document covered by the signature."),
    ("errors", "List<Message>", "O", "Validation errors of the signature."),
    ("warnings", "List<Message>", "O", "Validation warnings of the signature."),
    ("infos", "List<Message>", "O",
     "Validation information messages of the signature."),
    ("page", "Integer", "M", "Page number of the signature field."),
    ("dimension", "Dimension", "O",
     "Position and size of the signature field on the page."),
]

OBJ_CERTIFICATE = [
    ("id", "String", "O", "Identifier of the certificate in the report."),
    ("name", "String", "O", "Subject distinguished name of the certificate."),
    ("common_name", "String", "O", "Common name of the certificate subject."),
    ("given_name", "String", "O", "Given name of the certificate subject."),
    ("country_name", "String", "O", "Country of the certificate subject."),
    ("email", "String", "O", "Email address of the certificate subject."),
    ("entity_key", "String", "O", "Entity key of the certificate."),
    ("locality", "String", "O", "Locality of the certificate subject."),
    ("organization_identifier", "String", "O",
     "Organization identifier of the certificate subject."),
    ("organization_name", "String", "O",
     "Organization name of the certificate subject."),
    ("organizational_unit", "String", "O",
     "Organizational unit of the certificate subject."),
    ("state", "String", "O", "State or province of the certificate subject."),
    ("valid_from", "String", "O", "Start of the certificate validity period."),
    ("valid_to", "String", "O", "End of the certificate validity period."),
    ("serial_number", "String", "O", "Serial number of the certificate."),
    ("subject_dn", "Object", "O",
     "Subject distinguished name as a map of attribute names to values."),
    ("issuer_dn", "Object", "O",
     "Issuer distinguished name as a map of attribute names to values."),
    ("issuer", "String", "O", "Issuer distinguished name of the certificate."),
    ("value", "String", "O", "PEM encoded certificate."),
]

OBJ_SCOPE = [
    ("scope_name", "String", "O",
     "Type of the scope reported by the validation service."),
    ("name", "String", "O", "Name of the signature field."),
    ("value", "String", "O", "Description of the scope."),
]

OBJ_DIGEST = [
    ("digest_algorithm", "String", "O", "Digest algorithm."),
    ("digest_value", "String", "O", "Base64 encoded digest value."),
]

OBJ_META = [
    ("contact", "String", "O", "Contact recorded when the document was signed."),
    ("reason", "String", "O", "Reason recorded when the document was signed."),
    ("location", "String", "O",
     "Location recorded when the document was signed."),
]

OBJ_MESSAGE = [
    ("key", "String", "O", "Identifier of the validation message."),
    ("value", "String", "O", "Text of the validation message."),
]

OBJ_POLICY = [
    ("policy_name", "String", "O", "Name of the validation policy."),
    ("policy_description", "String", "O",
     "Description of the validation policy."),
]

OBJ_DETAILS = [
    ("validation_detail_report_id", "String", "O",
     "Identifier of the detailed report. Used as reportId to download the "
     "detailed and diagnostic reports."),
    ("validated_detail_document_hash", "String", "O",
     "Hash of the validated document recorded with the detailed report."),
    ("validation_basic_report_id", "String", "O",
     "Identifier of the simple report. Used as reportId to download the "
     "simple report."),
    ("validated_basic_document_hash", "String", "O",
     "Hash of the validated document recorded with the simple report."),
    ("detailed_validation_report", "String", "O",
     "Relative link to download the detailed validation report (PDF)."),
    ("diagnostic_data", "String", "O",
     "Relative link to download the diagnostic data (XML)."),
    ("basic_validation_report", "String", "O",
     "Relative link to download the simple validation report (PDF)."),
]

OBJ_FILE = [
    ("original_file_data", "String", "M",
     "Base64 encoded content of the validated document."),
    ("original_file_name", "String", "O",
     "File name of the validated document."),
    ("xml_diagnostic", "String", "O",
     "Raw diagnostic data of the validation in XML format."),
]


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guard_load(DOC_XML)
    root = tree.getroot()
    body = root.find(qn("body"))
    kids = list(body)

    for name in ("Add Validation", "Get List Validation", "Get Report of Validation",
                 "Download Reports", "Delete Validation"):
        try:
            st.find_heading(kids, 3, name)
            sys.exit("%r is already documented - nothing to do" % name)
        except LookupError:
            pass

    plain_tc, wrapped_tc = find_link_donor_cells(body)

    # ---- module heading (clone the RSSP Processing H2 paragraph) ----
    h2_donor = kids[st.find_heading(kids, 2, "RSSP Processing")]
    module_h2 = copy.deepcopy(h2_donor)
    for tag in ("bookmarkStart", "bookmarkEnd"):
        for bm in list(module_h2.iter(qn(tag))):
            bm.getparent().remove(bm)
    st.set_text_keep_bookmarks(module_h2, "Validation Management")

    # ---- endpoint sections ----
    endpoint_blocks = [module_h2]
    for builder, label in (
            (lambda: build_add_validation(root, kids), "Add Validation"),
            (lambda: build_get_list_validation(root, kids, wrapped_tc),
             "Get List Validation"),
            (lambda: build_get_report(root, kids, plain_tc, wrapped_tc),
             "Get Report of Validation"),
            (lambda: build_download_reports(root, kids), "Download Reports"),
            (lambda: build_delete_validation(root, kids), "Delete Validation")):
        blocks = builder()
        print("built %-26s %d blocks" % (label, len(blocks)))
        endpoint_blocks.extend(blocks)

    # ---- relative objects ----
    bm_id = st.next_bookmark_id(root)
    objects = []
    for title, anchor, rows, links in (
            ("Validation", A_VALIDATION, OBJ_VALIDATION, None),
            ("Overview(Validation)", A_OVERVIEW, OBJ_OVERVIEW, None),
            ("Signature(Validation)", A_SIGNATURE, OBJ_SIGNATURE, [
                ("meta_data", "MetaData", A_META),
                ("validation_policy", "ValidationPolicy", A_POLICY),
                ("data_to_be_signed", "Digest", A_DIGEST),
                ("signing_certificate", "Certificate", A_CERT),
                ("scope", "Scope", A_SCOPE),
                ("errors", "List<Message>", A_MESSAGE),
                ("warnings", "List<Message>", A_MESSAGE),
                ("infos", "List<Message>", A_MESSAGE),
                ("dimension", "Dimension", A_DIMENSION),
            ]),
            ("Certificate(Validation)", A_CERT, OBJ_CERTIFICATE, None),
            ("Scope(Validation)", A_SCOPE, OBJ_SCOPE, None),
            ("Digest(Validation)", A_DIGEST, OBJ_DIGEST, None),
            ("MetaData(Validation)", A_META, OBJ_META, None),
            ("Message(Validation)", A_MESSAGE, OBJ_MESSAGE, None),
            ("ValidationPolicy(Validation)", A_POLICY, OBJ_POLICY, None),
            ("Details(Validation)", A_DETAILS, OBJ_DETAILS, None),
            ("File(Validation)", A_FILE, OBJ_FILE, None)):
        blocks = build_object(root, kids, title, anchor, bm_id, rows,
                              links=links, plain_donor_tc=plain_tc,
                              wrapped_donor_tc=wrapped_tc)
        bm_id += 1
        print("built object %-28s %d rows" % (title, len(rows)))
        objects.append(blocks)
    object_blocks = [el for blocks in objects for el in blocks]

    # ---- insertion points ----
    admin_idx = st.find_heading(kids, 2, "ADMIN APIs")
    enum_idx = st.find_heading(kids, 1, "ENUMERATION")
    print("module goes before %r, objects before %r"
          % (st.text_of(kids[admin_idx]).strip(), st.text_of(kids[enum_idx]).strip()))

    if dry:
        print("\n--dry-run: nothing written")
        return

    anchor = kids[admin_idx]
    for el in endpoint_blocks:
        anchor.addprevious(el)
    anchor = kids[enum_idx]
    for el in object_blocks:
        anchor.addprevious(el)

    if os.path.exists(OUT_DOCX) and not os.path.exists(BACKUP_DOCX):
        shutil.copy2(OUT_DOCX, BACKUP_DOCX)
        print("backed up %s" % BACKUP_DOCX)

    try:
        guard_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("%s\n\nNothing was written. Re-run to re-apply on top of the "
                 "current content." % ex)

    st.pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
