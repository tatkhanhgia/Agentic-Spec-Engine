#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add four undocumented RSSP Processing endpoints to the specification.

All four exist in RSSPController.validURLs, are wired to real service code
(no stubs) and live in the Postman collection; the user approved adding them
on 2026-07-27. Facts from the backend, read-only:

1. Get Signing Transaction Status - GET /v1/csc/signing/{transactionCode}/status
   RSSPService.getSigningTransactionStatus: tries the RSS signing status
   first (type SIGNING, data = transaction/bill cache), falls back to the
   data-sharing status (type SHARING, GetSharingStatusSerializer). Optional
   x-connector-name header, default SMART_ID_RSS_VNeID.

2. Get Pending Signing Sessions - GET /v1/csc/signing/pending-sessions
   RssService.getPendingSigningSessions: sessions of the current user still
   waiting in RAM. Client-credentials tokens are rejected. NOTE: the session
   keys are serialized in camelCase by the backend (PendingSigningSessionResponse
   record has no @JsonProperty) - reported as a backend naming bug; the sample
   shows what the API actually returns today.

3. Get Sign Hash From Service - POST /v1/csc/signatures/sign-hash
   RssService.signHash: standalone hash signing via RSS. Required:
   connector_name, credential_id, documents[]. operation_mode S (sync,
   default, waits up to 300s) or A (async, waiting-approval response).
   The `hashing_algo` key of the sync response currently carries the signed
   hash value (CreateAuthorizeRSSTask passes getSignedHash()) - reported as a
   backend bug; the docs describe the intended meaning.

4. Check OCSP Certificate Status - POST /v1/csc/credentials/ocsp-check
   RSSPService.checkOcspStatus: direct `certificate` path, or
   connector_name + credential_id path. Returns valid/message/cert_status
   (GOOD | REVOKED | UNKNOWN). Disabled OCSP -> valid true + UNKNOWN.

Sections are cloned from same-module donors (badge pair, table styles), then
bookmarks are stripped and DrawingML/VML ids are re-issued so Word does not
see duplicates. The TOC is refreshed by the user with Ctrl+A -> F9.

Usage:
    python scripts/oneoff/add_rssp_processing_sections.py [--dry-run]
"""
import copy
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from docx_guard import load as guard_load, save as guard_save, ConcurrentModification

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted"
    "_backup_pre_rssp_add.docx")

MODULE_H2 = "RSSP Processing"
HEADINGS = {"Heading1", "Heading2", "Heading3"}
GET_FILL = "047857"

HEADER_LINE = re.compile(
    r"^(GET |POST |PUT |DELETE |Http Code:|Content-[Tt]ype:|Authorization:|x-[a-z0-9-]+:)")

BLOCK_CC = ("This API requires a user access token; a token obtained through "
            "Client Credentials Authentication is rejected.")


def el_text(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def para_style(el):
    pPr = el.find(qn("pPr"))
    if pPr is None:
        return ""
    style = pPr.find(qn("pStyle"))
    return style.get(qn("val")) if style is not None else ""


def set_leaf_text(el, value):
    runs = [t for t in el.iter(qn("t"))]
    if not runs:
        return False
    runs[0].text = value
    for extra in runs[1:]:
        extra.text = ""
    return True


def row_cells(tr):
    return tr.findall(qn("tc"))


def section_range(body, heading_text):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        return None
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if kids[j].tag == qn("p") and para_style(kids[j]) in HEADINGS:
            end = j
            break
    return start, end


def strip_bookmarks(blocks):
    for el in blocks:
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for bm in el.iter(qn(tag)):
                bm.getparent().remove(bm)


def clone_section(body, donor):
    span = section_range(body, donor)
    if span is None:
        sys.exit("donor section %r not found" % donor)
    start, end = span
    blocks = [copy.deepcopy(el) for el in list(body)[start:end]]
    strip_bookmarks(blocks)
    return blocks


def find_table(blocks, label):
    for el in blocks:
        if el.tag != qn("tbl"):
            continue
        rows = el.findall(qn("tr"))
        if rows and el_text(rows[0]).strip().startswith(label):
            return el
    sys.exit("table %r not found in cloned section" % label)


def data_rows(tbl):
    return [tr for tr in tbl.findall(qn("tr")) if len(row_cells(tr)) >= 5][1:]


def renumber(tbl):
    for n, tr in enumerate(data_rows(tbl), start=1):
        set_leaf_text(row_cells(tr)[0], str(n))


def named_row(tbl, name):
    for tr in data_rows(tbl):
        if el_text(row_cells(tr)[1]).strip() == name:
            return tr
    return None


def require_row(tbl, name):
    tr = named_row(tbl, name)
    if tr is None:
        sys.exit("row %r not found in cloned table" % name)
    return tr


def unwrap_links(el):
    for h in list(el.iter(qn("hyperlink"))):
        parent = h.getparent()
        idx = list(parent).index(h)
        for child in list(h):
            parent.insert(idx, child)
            idx += 1
        parent.remove(h)


def set_row(tr, name, typ, presence, desc):
    unwrap_links(tr)
    cells = row_cells(tr)
    set_leaf_text(cells[1], name)
    set_leaf_text(cells[2], typ)
    set_leaf_text(cells[3], presence)
    set_leaf_text(cells[4], desc)


def clone_row(tbl, donor_name, name, typ, presence, desc, after=None):
    donor = require_row(tbl, donor_name)
    new = copy.deepcopy(donor)
    set_row(new, name, typ, presence, desc)
    anchor = named_row(tbl, after) if after else data_rows(tbl)[-1]
    anchor.addnext(new)
    return new


def drop_row(tbl, name):
    tr = named_row(tbl, name)
    if tr is not None:
        tbl.remove(tr)


def rebuild_rows(tbl, rows):
    """Drop every data row, then rebuild from the first donor row's style."""
    donors = data_rows(tbl)
    if not donors:
        sys.exit("table has no data row to use as a style donor")
    template = copy.deepcopy(donors[0])
    for tr in donors:
        tbl.remove(tr)
    anchor = tbl.findall(qn("tr"))[-1]  # column-header row is now last
    for name, typ, presence, desc in rows:
        new = copy.deepcopy(template)
        set_row(new, name, typ, presence, desc)
        anchor.addnext(new)
        anchor = new
    renumber(tbl)


def frames_in(blocks, first_line_prefix):
    hits = []
    for el in blocks:
        if el.tag != qn("p"):
            continue
        txs = el.findall(".//" + qn("txbxContent"))
        if not txs:
            continue
        first = el_text(txs[0].findall(qn("p"))[0]).strip()
        if first.startswith(first_line_prefix):
            hits.append(el)
    return hits


def one(lst, what):
    if len(lst) != 1:
        sys.exit("expected exactly one %s, got %d" % (what, len(lst)))
    return lst[0]


def rewrite_body(paragraph, body_lines):
    for tx in paragraph.findall(".//" + qn("txbxContent")):
        paras = tx.findall(qn("p"))
        body = [p for p in paras if not HEADER_LINE.match(el_text(p))]
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


def drop_frame_lines(paragraph, prefixes):
    for tx in paragraph.findall(".//" + qn("txbxContent")):
        for p in list(tx.findall(qn("p"))):
            if any(el_text(p).strip().startswith(x) for x in prefixes):
                tx.remove(p)


def retag_method(paragraph, old, new, fill):
    """Turn the shaded method badge of a request frame into another verb."""
    changed = 0
    for tx in paragraph.findall(".//" + qn("txbxContent")):
        first = tx.findall(qn("p"))[0]
        for r in first.iter(qn("r")):
            ts = list(r.iter(qn("t")))
            if not ts or "".join(t.text or "" for t in ts).strip() != old:
                continue
            rPr = r.find(qn("rPr"))
            shd = rPr.find(qn("shd")) if rPr is not None else None
            if shd is None:
                continue
            shd.set(qn("fill"), fill)
            ts[0].text = new
            for extra in ts[1:]:
                extra.text = ""
            changed += 1
    if changed != 2:
        sys.exit("method badge retag hit %d runs, expected 2 (Choice+Fallback)" % changed)


def set_url(paragraph, new_url):
    """Rewrite the bold URL of the request frame's first line."""
    for tx in paragraph.findall(".//" + qn("txbxContent")):
        first = tx.findall(qn("p"))[0]
        done = False
        for r in first.iter(qn("r")):
            txt = "".join(t.text or "" for t in r.iter(qn("t")))
            if txt.strip() in ("GET", "POST", "PUT", "DELETE") or txt.strip() == "":
                continue
            ts = list(r.iter(qn("t")))
            if not done:
                ts[0].text = new_url
                for extra in ts[1:]:
                    extra.text = ""
                done = True
            else:
                for t in ts:
                    t.text = ""
        if not done:
            sys.exit("request frame first line has no URL run to rewrite")


def set_description(blocks, text):
    """Overwrite the paragraphs between the H3 and the Sample Request label."""
    seen_desc = False
    for el in blocks[1:]:
        if el.tag != qn("p"):
            break
        t = el_text(el).strip()
        if t == "Sample Request":
            break
        if not t:
            continue
        if not seen_desc:
            set_leaf_text(el, text)
            seen_desc = True
        else:
            set_leaf_text(el, "")
    if not seen_desc:
        sys.exit("no description paragraph found to overwrite")


def reissue_ids(tree, new_blocks):
    """Re-issue wp:docPr ids and VML shape ids duplicated by cloning."""
    root = tree.getroot()
    max_docpr = 0
    for d in root.iter("{%s}docPr" % WP):
        try:
            max_docpr = max(max_docpr, int(d.get("id")))
        except (TypeError, ValueError):
            pass
    max_spid = 9000
    for s in root.iter("{%s}shape" % V):
        spid = s.get("{%s}spid" % O) or ""
        m = re.match(r"_x0000_s(\d+)$", spid)
        if m:
            max_spid = max(max_spid, int(m.group(1)))
    n_doc = n_shape = 0
    for el in new_blocks:
        for d in el.iter("{%s}docPr" % WP):
            max_docpr += 1
            d.set("id", str(max_docpr))
            n_doc += 1
        for s in el.iter("{%s}shape" % V):
            max_spid += 1
            s.set("id", "_x0000_s%d" % max_spid)
            s.set("{%s}spid" % O, "_x0000_s%d" % max_spid)
            n_shape += 1
    print("re-issued %d docPr ids, %d VML shape ids" % (n_doc, n_shape))


# ---------------------------------------------------------------------------

def build_signing_transaction_status(body):
    blocks = clone_section(body, "Get Signing Status (TSE MID)")
    set_leaf_text(blocks[0], "Get Signing Transaction Status")
    set_description(blocks, (
        "Business applications can use this service API to get the status of a "
        "standalone RSS signing or data sharing transaction. The transaction code "
        "returned when the transaction was initiated is provided in the resource URL. "
        "The response type field tells whether the result describes a SIGNING or a "
        "SHARING transaction."))

    req_frame = one(frames_in(blocks, "POST "), "request frame")
    retag_method(req_frame, "POST", "GET", GET_FILL)
    set_url(req_frame, "{baseURL}/v1/csc/signing/{transactionCode}/status")
    rewrite_body(req_frame, ["{}"])
    # optional connector header shown in the sample
    resp200 = one(frames_in(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "status": "String",',
        '  "type": "SIGNING",',
        '  "data": {',
        '    "transactionCode": "String",',
        '    "billCodeMap": "String",',
        '    "requestTimestamp": 0,',
        '    "webhook": null,',
        '    "metadata": "String"',
        "  },",
        '  "metadata": "String"',
        "}",
    ])
    sharing = copy.deepcopy(resp200)
    blocks.insert(blocks.index(resp200) + 1, sharing)
    rewrite_body(sharing, [
        "{",
        '  "status": "INPROGRESS",',
        '  "type": "SHARING",',
        '  "webhook_url": "String",',
        '  "metadata": "String",',
        '  "user_info": {}',
        "}",
    ])

    hdr = find_table(blocks, "Header Request Attributes")
    if named_row(hdr, "x-connector-name") is None:
        clone_row(hdr, "x-language-name", "x-connector-name", "String", "O",
                  "Name of the RSS connector to query. Default is SMART_ID_RSS_VNeID "
                  "when omitted.", after="x-language-name")
    renumber(hdr)

    # the GET call has a path parameter instead of a request body
    donor_path = find_table(section_blocks_live(body, "Sign Document via RSSP Directly"),
                            "Path Attributes")
    path_tbl = copy.deepcopy(donor_path)
    strip_bookmarks([path_tbl])
    drop_row(path_tbl, "documentId")
    set_row(require_row(path_tbl, "packageId"), "transactionCode", "String", "M",
            "Transaction code of the signing or sharing transaction, returned when "
            "the transaction was initiated.")
    renumber(path_tbl)
    req_tbl = find_table(blocks, "Request Attributes")
    blocks.insert(blocks.index(req_tbl), path_tbl)
    blocks.remove(req_tbl)

    resp = find_table(blocks, "Response Attributes")
    set_row(require_row(resp, "status"), "status", "String", "M",
            "Status of the transaction reported by the RSS provider.")
    for name, typ, presence, desc in (
            ("type", "String", "M",
             "Type of the transaction. Possible values are: + SIGNING + SHARING"),
            ("data", "Object", "RC (When type is SIGNING)",
             "Transaction and bill mapping details of the signing transaction."),
            ("webhook_url", "String", "RC (When type is SHARING)",
             "Webhook URL registered to receive the sharing result."),
            ("user_info", "Object", "RC (When type is SHARING)",
             "Citizen information shared by the provider. Empty until the user "
             "completes the sharing."),
            ("metadata", "String", "O",
             "Metadata supplied by the client when the transaction was created.")):
        if named_row(resp, name) is None:
            clone_row(resp, "status", name, typ, presence, desc)
    renumber(resp)
    return blocks


def build_pending_sessions(body):
    blocks = clone_section(body, "Get RSSP Information")
    set_leaf_text(blocks[0], "Get Pending Signing Sessions")
    set_description(blocks, (
        "Business applications can use this service API to list the RSS signing "
        "sessions of the current user that are still waiting for authorization. "
        + BLOCK_CC))

    req_frame = one(frames_in(blocks, "GET "), "request frame")
    set_url(req_frame, "{baseURL}/v1/csc/signing/pending-sessions")
    drop_frame_lines(req_frame, ("x-search-server",))

    hdr = find_table(blocks, "Header Request Attributes")
    drop_row(hdr, "x-search-server")
    renumber(hdr)

    resp = find_table(blocks, "Response Attributes")
    drop_row(resp, "signing_servers")
    if named_row(resp, "total") is None:
        clone_row(resp, "error_code", "total", "Integer", "M",
                  "Number of pending signing sessions of the current user.")
        clone_row(resp, "total", "sessions", "List<Object>", "M",
                  "Pending sessions. Each item contains requestId, credentialId, "
                  "connectorName, packageId, documentId, createdBy, createdAt "
                  "(epoch milliseconds), transactionId, status and isAgent.",
                  after="total")
    renumber(resp)

    resp200 = one(frames_in(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "total": 1,',
        '  "sessions": [',
        "    {",
        '      "requestId": "String",',
        '      "credentialId": "String",',
        '      "connectorName": "String",',
        '      "packageId": 0,',
        '      "documentId": 0,',
        '      "createdBy": "user@example.com",',
        '      "createdAt": 0,',
        '      "transactionId": "String",',
        '      "status": "WAITING_USER",',
        '      "isAgent": false',
        "    }",
        "  ]",
        "}",
    ])
    return blocks


def build_sign_hash(body):
    blocks = clone_section(body, "Get Credentials Info")
    set_leaf_text(blocks[0], "Get Sign Hash From Service")
    set_description(blocks, (
        "Business applications can use this service API to sign one or more document "
        "hashes directly via the RSS service, without a document package. Each digest "
        "is sent in the request and the signature value is returned per document. In "
        "synchronous mode (operation_mode S, default) the call waits for the user to "
        "authorize on the provider application; in asynchronous mode (A) the call "
        "returns a waiting-approval response immediately and the result is delivered "
        "later. " + BLOCK_CC))

    req_frame = one(frames_in(blocks, "POST "), "request frame")
    set_url(req_frame, "{baseURL}/v1/csc/signatures/sign-hash")
    rewrite_body(req_frame, [
        "{",
        '  "connector_name": "String",',
        '  "credential_id": "String",',
        '  "citizen_id": "String",',
        '  "documents": [',
        "    {",
        '      "document_name": "document1",',
        '      "digest_value": "Base64 encoded digest"',
        "    }",
        "  ],",
        '  "signing_algo": "SHA256",',
        '  "hashing_algo": "RSA",',
        '  "operation_mode": "S"',
        "}",
    ])

    req = find_table(blocks, "Request Attributes")
    rebuild_rows(req, [
        ("connector_name", "String", "M", "Name of the RSS connector used for signing."),
        ("credential_id", "String", "M", "Credential id used to sign the hashes."),
        ("documents", "List<Object>", "M",
         "Documents to sign. Each item contains document_name and digest_value "
         "(base64 encoded digest of the document)."),
        ("signing_algo", "String", "O", "Signing algorithm requested from the provider."),
        ("hashing_algo", "String", "O", "Hashing algorithm used to compute the digests."),
        ("citizen_id", "String", "O", "Citizen identity number of the signer."),
        ("operation_mode", "String", "O",
         "Operation mode. Possible values are: + S (synchronous, default) "
         "+ A (asynchronous)"),
        ("is_redirect_app_enabled", "Boolean", "O",
         "If true, the signing session is treated as running on a mobile channel and "
         "the provider app redirect is enabled."),
        ("metadata", "Object", "O",
         "Process metadata from the client (app info, device fingerprint, network)."),
    ])

    resp = find_table(blocks, "Response Attributes")
    keep_err = [(el_text(row_cells(tr)[1]).strip(),) for tr in data_rows(resp)]
    for tr in list(data_rows(resp)):
        name = el_text(row_cells(tr)[1]).strip()
        if name not in ("error_description", "error_code", "message"):
            resp.remove(tr)
    for name, typ, presence, desc in (
            ("status", "String", "M", "Status of the signing request."),
            ("signing_algo", "String", "RC (When operation_mode is S)",
             "Signing algorithm used by the provider."),
            ("hashing_algo", "String", "RC (When operation_mode is S)",
             "Hashing algorithm used for the digests."),
            ("documents", "List<Object>", "RC (When operation_mode is S)",
             "Signed documents. Each item contains document_name, digest_value and "
             "signature_value (base64 encoded signature)."),
            ("validity_period", "Integer", "RC (When operation_mode is A)",
             "Validity period of the authorization request in seconds."),
            ("request_id", "String", "RC (When operation_mode is A)",
             "Identifier of the authorization request, also returned as status_id "
             "and responseID for polling."),
            ("transaction_code", "String", "RC (When operation_mode is A)",
             "Transaction code usable with Get Signing Transaction Status.")):
        clone_row(resp, "message", name, typ, presence, desc)
    renumber(resp)

    resp200 = one(frames_in(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "status": "SUCCESS",',
        '  "signing_algo": "String",',
        '  "hashing_algo": "String",',
        '  "documents": [',
        "    {",
        '      "document_name": "document1",',
        '      "digest_value": "Base64 encoded digest",',
        '      "signature_value": "Base64 encoded signature"',
        "    }",
        "  ]",
        "}",
    ])
    async_frame = copy.deepcopy(resp200)
    blocks.insert(blocks.index(resp200) + 1, async_frame)
    rewrite_body(async_frame, [
        "{",
        '  "status": "WAITING APPROVAL",',
        '  "validity_period": 300,',
        '  "status_id": "String",',
        '  "responseID": "String",',
        '  "request_id": "String",',
        '  "transaction_code": "String",',
        '  "transaction_id": "String"',
        "}",
    ])
    return blocks


def build_ocsp_check(body):
    blocks = clone_section(body, "Get Credentials Info")
    set_leaf_text(blocks[0], "Check OCSP Certificate Status")
    set_description(blocks, (
        "Business applications can use this service API to check the revocation "
        "status of a signing certificate via OCSP. The certificate can be sent "
        "directly in base64 format, or connector_name and credential_id can be "
        "provided so the service loads the certificate from the signing provider. "
        "When OCSP checking is disabled for the connector, the call returns "
        "cert_status UNKNOWN with valid true. " + BLOCK_CC))

    req_frame = one(frames_in(blocks, "POST "), "request frame")
    set_url(req_frame, "{baseURL}/v1/csc/credentials/ocsp-check")
    rewrite_body(req_frame, [
        "{",
        '  "connector_name": "String",',
        '  "credential_id": "String",',
        '  "certificate": "Base64 encoded certificate"',
        "}",
    ])

    req = find_table(blocks, "Request Attributes")
    rebuild_rows(req, [
        ("connector_name", "String", "RC (When certificate is not provided)",
         "Name of the connector used to load the credential's certificate."),
        ("credential_id", "String", "RC (When certificate is not provided)",
         "Credential id whose certificate is checked."),
        ("certificate", "String", "O",
         "Base64 encoded certificate to check directly. When provided, "
         "connector_name and credential_id are not needed."),
    ])

    resp = find_table(blocks, "Response Attributes")
    for tr in list(data_rows(resp)):
        name = el_text(row_cells(tr)[1]).strip()
        if name not in ("error_description", "error_code", "message"):
            resp.remove(tr)
    for name, typ, presence, desc in (
            ("valid", "Boolean", "M",
             "True when the certificate passed the OCSP validation."),
            ("message", "String", "M",
             "Human readable description of the OCSP result (success response)."),
            ("cert_status", "String", "M",
             "OCSP certificate status. Possible values are: + GOOD + REVOKED "
             "+ UNKNOWN")):
        clone_row(resp, "error_code", name, typ, presence, desc)
    renumber(resp)

    resp200 = one(frames_in(blocks, "Http Code: 200"), "200 frame")
    rewrite_body(resp200, [
        "{",
        '  "valid": true,',
        '  "message": "Certificate is valid",',
        '  "cert_status": "GOOD"',
        "}",
    ])
    return blocks


def section_blocks_live(body, heading_text):
    span = section_range(body, heading_text)
    if span is None:
        sys.exit("section %r not found" % heading_text)
    start, end = span
    return list(body)[start:end]


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guard_load(DOC_XML)
    body = tree.getroot().find(qn("body"))

    names = ("Get Signing Transaction Status", "Get Pending Signing Sessions",
             "Get Sign Hash From Service", "Check OCSP Certificate Status")
    for name in names:
        if section_range(body, name) is not None:
            sys.exit("%r is already documented - nothing to do" % name)

    built = [
        build_signing_transaction_status(body),
        build_pending_sessions(body),
        build_sign_hash(body),
        build_ocsp_check(body),
    ]
    for name, blocks in zip(names, built):
        print("built %-34s %d blocks" % (name, len(blocks)))

    all_blocks = [el for blocks in built for el in blocks]
    reissue_ids(tree, all_blocks)

    kids = list(body)
    h2 = [i for i, el in enumerate(kids)
          if el.tag == qn("p") and para_style(el) == "Heading2"
          and el_text(el).strip() == MODULE_H2]
    if not h2:
        sys.exit("Heading2 %r not found" % MODULE_H2)
    nxt = next((i for i in range(h2[0] + 1, len(kids))
                if kids[i].tag == qn("p") and para_style(kids[i]) == "Heading2"), None)
    if nxt is None:
        sys.exit("no Heading2 after %r - refusing to guess the module end" % MODULE_H2)
    print("inserting before %r" % el_text(kids[nxt]).strip())

    if dry:
        print("\n--dry-run: nothing written")
        return

    anchor = kids[nxt]
    for el in all_blocks:
        anchor.addprevious(el)

    if os.path.exists(OUT_DOCX) and not os.path.exists(BACKUP_DOCX):
        shutil.copy2(OUT_DOCX, BACKUP_DOCX)
        print("\nbacked up %s" % BACKUP_DOCX)

    try:
        guard_save(tree, DOC_XML, token)
    except ConcurrentModification as ex:
        sys.exit("%s\n\nNothing was written. Re-run to re-apply on top of the "
                 "current content." % ex)

    tmp = OUT_DOCX + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(UNPACKED, "[Content_Types].xml"), "[Content_Types].xml")
        for base, _dirs, files in os.walk(UNPACKED):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, UNPACKED).replace(os.sep, "/")
                if rel != "[Content_Types].xml":
                    z.write(full, rel)
    shutil.move(tmp, OUT_DOCX)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
