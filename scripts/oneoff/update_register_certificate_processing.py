#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Bring the Register Certificate Processing category in line with the backend.

Backend evidence: RSSPController validURLs [6..12,19], RegisterCertificateRequest,
SendOTPToRegisterCertificateRequest, VerifyOTPRequest, FinishRegisterCertificateRequest,
CancelIssueCertificateRequest, ValidateRequest (rssp), LevelOfAssurance enum,
Types.AuthMode enum, OperationMode enum, TaxInformationResponse.DocumentData.

In-place fixes
  1. Register Certificate via RSSP - request table used `connector_name` (O) but the
     backend key is `enrollment_connector` and it is mandatory; `face_image` is not
     validated as mandatory; `mobile_number` / `document_extraction_connector` /
     `metadata` were missing; the level_of_assurance value list was incomplete.
  2. Send OTP - request table carried `otp`/`user_email`, which belong to Verify.
     The real body is `email` / `mobile` (at least one). Heading renamed to the
     backend apiName "Send OTP for Certificate Registration".
  3. Verify OTP - `user_email` does not exist in VerifyOTPRequest; row dropped.
  4. Get List Tax Codes - a GET with no body carried a copy-pasted request table;
     dropped. Success response documented (`tax_informations`).
  5. Finish Register Certificate - `auth_mode` / `webhook_url` / `operation_mode` /
     `metadata` were missing; async behaviour (5-minute held connection, cancelable)
     noted in the description.

New sections (cloned, never hand-assembled)
  6. Get Credential Registration Status  GET /v1/csc/credentials/{transactionCode}/status
     donor: Get List Tax Codes (after fix 4 it has the target shape: GET, path
     param, no request table).
  7. Cancel Issue Certificate  POST /v1/ras/issue/cancel
     donor: Verify OTP (after fix 3: POST, small body, empty 200 frame); its
     Path Attributes table is dropped because Cancel has no path parameter.

Deliberately NOT copied from the backend: the camelCase keys inside the Status
response `data` object (requestType, citizenPid, ...) - per the project naming
rule a camelCase body key is a backend defect to report, so the table describes
the object in prose instead of enumerating defective keys.

Usage:
    python scripts/oneoff/update_register_certificate_processing.py [--dry-run]
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guarded_load, save as guarded_save, ConcurrentModification
import section_tools as st

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace("-Formatted.docx", "-Formatted_backup_pre_register_cert_module.docx")

qn = st.qn

METADATA_DESC = "Process metadata from the client (app info, device fingerprint, network)."


def section(kids, title):
    a, b = st.section_range(kids, st.find_heading(kids, 3, title))
    return a, b, kids[a:b]


def request_frame(elements, method):
    """The single request frame of a section, as [(txbx, paras), ...] copies."""
    return list(st.each_frame(elements, method + " "))


def response_frame(elements, code):
    return [(t, p) for t, p in st.each_frame(elements, "Http Code: %d" % code)]


def set_frame_body(copies, lines):
    for txbx, paras in copies:
        st.replace_frame_body(txbx, txbx.findall(qn("p")), lines)


def first_line_replace(copies, pairs):
    for _txbx, paras in copies:
        for old, new in pairs:
            st.set_path_suffix(paras[0], old, new)


def main():
    dry = "--dry-run" in sys.argv

    tree, token = guarded_load(DOC_XML)
    root = tree.getroot()
    body = root.find(qn("body"))
    kids = list(body)

    # ---- 1. Register Certificate via RSSP --------------------------------
    a, b, els = section(kids, "Register Certificate via RSSP")
    tbl = st.find_table(els, "Request Attributes")
    st.set_row(tbl, "connector_name", new_name="enrollment_connector", presence="M",
               description="Connector used to enroll the certificate.")
    st.set_row(tbl, "signing_server",
               presence="RC(When the enrollment_connector is have the same name in DB)")
    st.set_row(tbl, "level_of_assurance", description=(
        "Level of Assurance that the user wants to create. Possible values are: "
        "+ ADVANCED_ELECTRONIC_SIGNATURE + ELECTRONIC_SEAL + QUALIFIED_ELECTRONIC_SEAL "
        "+ ADVANCED_ELECTRONIC + QUALIFIED_ELECTRONIC_SIGNATURE"))
    st.set_row(tbl, "face_image", presence="O")
    st.append_row(tbl, "mobile_number", "String", "O", "Mobile number of the user.")
    st.append_row(tbl, "document_extraction_connector", "String", "O",
                  "Connector used to extract the identity document information.")
    st.append_row(tbl, "metadata", "Object", "O", METADATA_DESC)
    set_frame_body(request_frame(els, "POST"), [
        "{",
        '  "enrollment_connector": "",',
        '  "level_of_assurance": "",',
        '  "document_type": "",',
        '  "value": "",',
        '  "face_image": ""',
        "}",
    ])
    print("1. Register Certificate via RSSP: request table + sample updated")

    # ---- 2. Send OTP ------------------------------------------------------
    i = st.find_heading(kids, 3, "Send OTP to Certificate Registration")
    st.set_text_keep_bookmarks(kids[i], "Send OTP for Certificate Registration")
    _a, _b, els = section(kids, "Send OTP for Certificate Registration")
    tbl = st.find_table(els, "Request Attributes")
    st.set_row(tbl, "otp", new_name="email", presence="RC", description=(
        "Email address to receive the OTP. Required when the user email has not "
        "been verified yet. At least one of email or mobile must be provided."))
    st.set_row(tbl, "user_email", new_name="mobile", presence="RC", description=(
        "Mobile number to receive the OTP. Required when the user mobile number has "
        "not been verified yet. At least one of email or mobile must be provided."))
    print("2. Send OTP: heading renamed, request table fixed")

    # ---- 3. Verify OTP ----------------------------------------------------
    _a, _b, els = section(kids, "Verify OTP for Certificate Registration")
    tbl = st.find_table(els, "Request Attributes")
    st.drop_row(tbl, "user_email")
    print("3. Verify OTP: user_email dropped")

    # ---- 4. Get List Tax Codes -------------------------------------------
    a, b, els = section(kids, "Get List Tax Codes")
    tbl = st.find_table(els, "Request Attributes")
    nxt = tbl.getnext()
    tbl.getparent().remove(tbl)
    if nxt is not None and nxt.tag == qn("p") and not st.text_of(nxt).strip():
        nxt.getparent().remove(nxt)
    kids = list(body)                       # tree changed shape
    _a, _b, els = section(kids, "Get List Tax Codes")
    tbl = st.find_table(els, "Response Attributes")
    st.append_row(tbl, "tax_informations", "List<Object>", "M", (
        "List of tax registration records associated with the certificate holder. "
        "Each record contains owner, director, chief accountant, company and "
        "source information."))
    set_frame_body(response_frame(els, 200), [
        "{",
        '  "tax_informations": []',
        "}",
    ])
    print("4. Get List Tax Codes: request table dropped, success response documented")

    # ---- 5. Finish Register Certificate ----------------------------------
    _a, _b, els = section(kids, "Finish Register Certificate")
    desc = els[1]
    st.set_para_text(desc, (
        "Business applications can use this service API to finish your flow to "
        "register new Certificate. This API is processed asynchronously: the "
        "connection is held open (up to 5 minutes) until the certificate issuance "
        "completes, or until the flow is canceled via Cancel Issue Certificate."))
    tbl = st.find_table(els, "Request Attributes")
    st.append_row(tbl, "auth_mode", "String", "O", (
        "Authorization mode of the credential. Possible values are: + EXPLICIT/PIN "
        "+ EXPLICIT/OTP-SMS + EXPLICIT/OTP-EMAIL + IMPLICIT/TSE + EXPLICIT/JWT"))
    st.append_row(tbl, "webhook_url", "String", "O",
                  "Webhook URL that receives the certificate issuance result callback.")
    st.append_row(tbl, "operation_mode", "String", "O",
                  "Operation mode. Possible values are: + A (asynchronous, default) "
                  "+ S (synchronous) + C")
    st.append_row(tbl, "metadata", "Object", "O", METADATA_DESC)
    set_frame_body(request_frame(els, "POST"), [
        "{",
        '  "certificate_profile": "",',
        '  "tax_code": "",',
        '  "issuing_provider": ""',
        "}",
    ])
    print("5. Finish Register Certificate: 4 fields added, async behaviour noted")

    # ---- 6. Get Credential Registration Status (clone of Tax Codes) ------
    kids = list(body)
    clone, info = st.clone_section(root, kids, "Get List Tax Codes")
    st.set_para_text(clone[0], "Get Credential Registration Status")
    st.set_para_text(clone[1], (
        "Business applications can use this service API to get the status of a "
        "certificate registration transaction."))
    first_line_replace(request_frame(clone, "GET"), [
        ("{session_id}", "{transactionCode}"),
        ("/taxcode", "/status"),
    ])
    tbl = st.find_table(clone, "Header Request Attributes")
    st.append_row(tbl, "x-connector-name", "String", "O", (
        "Connector used to query the registration status. Default value is "
        "SMART_ID_RSS_VNeID."))
    tbl = st.find_table(clone, "Path Attributes")
    st.set_row(tbl, "session_id", new_name="transactionCode",
               description="Transaction code of the certificate registration transaction.")
    tbl = st.find_table(clone, "Response Attributes")
    st.drop_row(tbl, "tax_informations")
    st.append_row(tbl, "status", "String", "M",
                  "Status of the certificate registration transaction.")
    st.append_row(tbl, "data", "Object", "O", (
        "Details of the transaction: request type, citizen PID, credential ID, "
        "serial number and certificate data."))
    set_frame_body(response_frame(clone, 200), [
        "{",
        '  "status": "",',
        '  "data": {}',
        "}",
    ])
    st.insert_after_section(body, kids, "Finish Register Certificate", clone)
    print("6. Get Credential Registration Status: cloned (%s)" % info)

    # ---- 7. Cancel Issue Certificate (clone of Verify OTP) ----------------
    kids = list(body)
    clone, info = st.clone_section(root, kids, "Verify OTP for Certificate Registration")
    st.set_para_text(clone[0], "Cancel Issue Certificate")
    st.set_para_text(clone[1], (
        "Business applications can use this service API to cancel an in-progress "
        "certificate registration transaction. Canceling releases the pending "
        "Finish Register Certificate request immediately."))
    first_line_replace(request_frame(clone, "POST"), [
        ("/v1/csc/credentials/", "/v1/ras/issue/"),
        ("{session_id}", ""),
        ("/verify", "cancel"),
    ])
    set_frame_body(request_frame(clone, "POST"), [
        "{",
        '  "session_id": ""',
        "}",
    ])
    st.drop_table(clone, "Path Attributes")
    tbl = st.find_table(clone, "Request Attributes")
    st.set_row(tbl, "otp", new_name="session_id", presence="M",
               description="Session id of the certificate registration flow to cancel.")
    st.insert_after_section(body, kids, "Get Credential Registration Status", clone)
    print("7. Cancel Issue Certificate: cloned (%s)" % info)

    # ---- verification: resulting category shape ---------------------------
    kids = list(body)
    i2 = next(i for i, e in enumerate(kids)
              if e.tag == qn("p") and st.style_of(e) == "Heading2"
              and st.text_of(e).strip() == "Register Certificate Processing")
    print("\ncategory after edit:")
    for j in range(i2 + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and st.style_of(e) in ("Heading1", "Heading2"):
            break
        if e.tag == qn("p") and st.style_of(e) == "Heading3":
            print("  H3:", st.text_of(e))
        if e.tag == qn("tbl"):
            print("      -", st.table_title(e), "rows=%d" % len(e.findall(qn("tr"))))

    # duplicate identifier sweep (the corruption class cloning creates)
    dp = [d.get("id") for d in root.iter() if st.local(d) == "docPr"]
    sh = [d.get("id") for d in root.iter() if st.local(d) == "shape"]
    bm = [x.get(qn("id")) for x in root.iter(qn("bookmarkStart"))]
    print("\nduplicate docPr=%d shape=%d bookmark=%d"
          % (len(dp) - len(set(dp)), len(sh) - len(set(sh)), len(bm) - len(set(bm))))
    if len(dp) != len(set(dp)) or len(sh) != len(set(sh)) or len(bm) != len(set(bm)):
        sys.exit("duplicate identifiers - refusing to write")

    if dry:
        print("\n--dry-run: nothing written")
        return

    shutil.copy2(OUT_DOCX, BACKUP)
    print("\nbackup: %s" % os.path.basename(BACKUP))
    try:
        guarded_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("\nREFUSED TO WRITE\n%s" % ex)
    st.pack(UNPACKED, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
