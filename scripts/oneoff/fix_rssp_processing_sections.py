#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Align the seven existing RSSP Processing sections with the backend.

Facts come from the backend (read-only): RSSPController, RSSPService,
rssp/service/ValidateRequest, the request DTOs and the response serializers.
Decisions confirmed by the user on 2026-07-27:

1. Get RSSP Information
   * success sample was invalid JSON (trailing comma) and `client_id` is
     always written as null by `GetListRSSPConnectorsSerialize`.
   * the sample keeps both `authentication_required` and
     `authorization_required`: the serializer writes `authentication_required`
     twice (backend bug, reported), the second was meant to be
     `authorization_required` - the docs record the intended meaning.

2. Get Filtered Credential List
   * `x-search-server` header is never read by this endpoint (the signing
     server comes from the `signing_server` body field) - row removed.
   * `x-credential-id` optional filter header was missing - row added
     (RSSPService.getListCredentials filters by Headers.X_CREDENTIAL_ID).
   * success sample now shows the full item written by
     `GetCredentialsSerialize`: cert, issuer_dn, subject_dn, valid_from,
     valid_to.

3. Sign Document via RSSP Directly
   * table said `signer_contact`; the JSON key is `contact_information`
     (the sample was already correct).
   * `time_zone` does not exist on SignDocumentRSSPRequest - row and sample
     line removed.
   * `signing_algo` / `hashing_algo` are not enforced by validateSignWithRSSP
     - presence M -> O.
   * missing optional fields added: signing_server (RC), citizen_id,
     webhook_url, rssp_operation_mode (S|A), is_redirect_app_enabled,
     notification_template, metadata.
   * success sample aligned to the authorize response
     (validity_period / vc_code) instead of a free-text message.

4. RSSP Credential Authorization
   * backend accepts one of THREE authenticators: otp | pin | face_image
     (SUBCODE_RSSP_MISSING_ONE_OF_THREE_TYPE_AUTHENTICATION) - face_image
     row added, otp/pin descriptions updated.
   * `responseID` routes the call to the sign-hash polling flow - documented.

5. Get Signing Status (TSE MID)
   * canonical request key is `status_id` (`transaction_id` is a JsonAlias);
     value format is credentialId.packageId.documentId.fieldName - the
     description now says so.

6. Get Credentials Info
   * `connector_name` is required only when `credential_id` is used
     (certificate_encoded path needs no connector) - presence M -> RC.
   * response `cert` object per GetCredentialInfoSerialize: adds mode,
     status_desc; issuer_dn/subject_dn are objects; `certificates[]` and
     `serial_number` do not exist. Top-level `fields` object was missing.
   * sample had invalid JSON (`"multisign": ""1`).

7. Send OTP via RSSP
   * request accepts optional `notification_template`.
   * success response is {"validity_period": 300, "user_email": ...},
     not {}.

Objects (targets of the composite-type hyperlinks - see CLAUDE.md):
   * GetFilteredCredentialListResponse: issuer_dn/subject_dn typed
     List<String> but serialized as DN-component objects -> Object.
   * CredentialInfo: same DN fix, plus the missing `status_desc` row.

Usage:
    python scripts/oneoff/fix_rssp_processing_sections.py [--dry-run]
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
    "_backup_pre_rssp_fix.docx")

HEADINGS = {"Heading1", "Heading2", "Heading3"}

# Lines that are frame headers, not body: kept verbatim by rewrite_body().
HEADER_LINE = re.compile(
    r"^(GET |POST |PUT |DELETE |Http Code:|Content-[Tt]ype:|Authorization:|x-[a-z0-9-]+:)")


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


def section_blocks(body, heading_text):
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn("p") and para_style(el) in HEADINGS \
                and el_text(el).strip() == heading_text:
            start = i
            break
    if start is None:
        sys.exit("section %r not found" % heading_text)
    end = len(kids)
    for j in range(start + 1, len(kids)):
        if kids[j].tag == qn("p") and para_style(kids[j]) in HEADINGS:
            end = j
            break
    return kids[start:end]


def find_table(blocks, label):
    for el in blocks:
        if el.tag != qn("tbl"):
            continue
        rows = el.findall(qn("tr"))
        if rows and el_text(rows[0]).strip().startswith(label):
            return el
    sys.exit("table %r not found in section" % label)


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


def require_row(tbl, name, where):
    tr = named_row(tbl, name)
    if tr is None:
        sys.exit("row %r not found in %s" % (name, where))
    return tr


def unwrap_links(el):
    """Replace every hyperlink with its child runs (clone-only helper).

    Safe only on a copy whose text is fully overwritten afterwards - on a
    live cell this would discard the link the reader still needs."""
    for h in list(el.iter(qn("hyperlink"))):
        parent = h.getparent()
        idx = list(parent).index(h)
        for child in list(h):
            parent.insert(idx, child)
            idx += 1
        parent.remove(h)


def clone_row(tbl, donor_name, name, typ, presence, desc, after=None):
    """Copy a donor row, drop its hyperlinks, overwrite its cells, insert it."""
    donor = require_row(tbl, donor_name, "clone donor")
    new = copy.deepcopy(donor)
    unwrap_links(new)
    cells = row_cells(new)
    set_leaf_text(cells[1], name)
    set_leaf_text(cells[2], typ)
    set_leaf_text(cells[3], presence)
    set_leaf_text(cells[4], desc)
    anchor = named_row(tbl, after) if after else data_rows(tbl)[-1]
    anchor.addnext(new)
    return new


def drop_row(tbl, name):
    tr = named_row(tbl, name)
    if tr is not None:
        tbl.remove(tr)
        return True
    return False


def frames_in(blocks, first_line_prefix):
    """All frame paragraphs whose first frame line starts with the prefix."""
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


def rewrite_body(paragraph, body_lines):
    """Replace the JSON body of BOTH txbxContent copies, keep header lines."""
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


def one(lst, what):
    if len(lst) != 1:
        sys.exit("expected exactly one %s, got %d" % (what, len(lst)))
    return lst[0]


TX_BEFORE = "Transaction identifier returned by "
TX_AFTER = (" when the signing was initiated, in the format "
            "credentialId.packageId.documentId.fieldName. The canonical field name is "
            "status_id; transaction_id is accepted as an alias.")
XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def set_text_around_link(cell, before, after):
    """Rewrite a description cell while keeping its single hyperlink intact."""
    paras = cell.findall(qn("p"))
    para = next((p for p in paras if p.findall(qn("hyperlink"))), None)
    if para is None:
        sys.exit("description cell has no hyperlink paragraph to keep")
    links = para.findall(qn("hyperlink"))
    if len(links) != 1:
        sys.exit("expected one hyperlink in the description, got %d" % len(links))
    link = links[0]
    seen = False
    before_runs, after_runs = [], []
    for child in para:
        if child is link:
            seen = True
            continue
        if child.tag == qn("r"):
            (after_runs if seen else before_runs).append(child)

    def fill(runs, text, insert_after):
        ts = [t for r in runs for t in r.iter(qn("t"))]
        if ts:
            ts[0].text = text
            ts[0].set(XMLSPACE, "preserve")
            for extra in ts[1:]:
                extra.text = ""
            return
        # no run on this side of the link: clone one from inside the link
        donor = next(iter(link.findall(qn("r"))), None)
        if donor is None:
            sys.exit("hyperlink has no run to clone for the surrounding text")
        run = copy.deepcopy(donor)
        for t in run.iter(qn("t")):
            t.text = ""
        one_t = next(iter(run.iter(qn("t"))), None)
        if one_t is None:
            sys.exit("cloned run has no text node")
        one_t.text = text
        one_t.set(XMLSPACE, "preserve")
        insert_after.addnext(run) if insert_after is not None else para.append(run)

    fill(before_runs, before, None) if before_runs else fill([], before, None)
    if before_runs:
        pass
    else:
        # the cloned run was appended at the end - move it before the link
        last = list(para)[-1]
        if last.tag == qn("r"):
            link.addprevious(last)
    fill(after_runs, after, link)
    # wipe any other paragraphs in the cell so stale text does not linger
    for p in paras:
        if p is para:
            continue
        for t in p.iter(qn("t")):
            t.text = ""

ONE_OF_THREE = "At least one of otp, pin or face_image must be provided."


def main():
    dry = "--dry-run" in sys.argv
    tree, token = guard_load(DOC_XML)
    body = tree.getroot().find(qn("body"))

    # --- 1) Get RSSP Information -------------------------------------------
    blocks = section_blocks(body, "Get RSSP Information")
    frame = one(frames_in(blocks, "Http Code: 200"), "200 frame in Get RSSP Information")
    rewrite_body(frame, [
        "{",
        '  "signing_servers": [',
        "    {",
        '      "signing_server": "String",',
        '      "connector_name": "RSSP",',
        '      "host_url": "String",',
        '      "client_id": null,',
        '      "authentication_required": true,',
        '      "authorization_required": true,',
        '      "auth_type": "String",',
        '      "rssp_rp_name": "String",',
        '      "rssp_rp_user": "String"',
        "    },",
        "    {}",
        "  ]",
        "}",
    ])
    print("1) Get RSSP Information: 200 sample rewritten (valid JSON, client_id null)")

    # --- 2) Get Filtered Credential List -----------------------------------
    blocks = section_blocks(body, "Get Filtered Credential List")
    hdr = find_table(blocks, "Header Request Attributes")
    dropped = drop_row(hdr, "x-search-server")
    if named_row(hdr, "x-credential-id") is None:
        clone_row(hdr, "x-language-name", "x-credential-id", "String", "O",
                  "Optional filter. When provided, only the credential whose credential "
                  "ID matches this value (case-insensitive) is returned.",
                  after="x-language-name")
    renumber(hdr)
    frame = one(frames_in(blocks, "Http Code: 200"), "200 frame in credential list")
    rewrite_body(frame, [
        "{",
        '  "credentialIDs": [',
        "    {",
        '      "credential_id": "Id",',
        '      "mode": "TSE",',
        '      "status": "OPERATED",',
        '      "cert": "Base64 encoded certificate",',
        '      "issuer_dn": {},',
        '      "subject_dn": {},',
        '      "valid_from": "String",',
        '      "valid_to": "String"',
        "    },",
        "    {}",
        "  ]",
        "}",
    ])
    print("2) Get Filtered Credential List: x-search-server %s, x-credential-id added, "
          "200 sample completed" % ("dropped" if dropped else "already absent"))

    # --- 3) Sign Document via RSSP Directly --------------------------------
    blocks = section_blocks(body, "Sign Document via RSSP Directly")
    req = find_table(blocks, "Request Attributes")
    tr = named_row(req, "signer_contact")
    if tr is not None:
        cells = row_cells(tr)
        set_leaf_text(cells[1], "contact_information")
        set_leaf_text(cells[4],
                      "Contact information of the signer, for example an email address. "
                      "If omitted, the email of the current workflow participant is used.")
    drop_row(req, "time_zone")
    for algo in ("signing_algo", "hashing_algo"):
        row = named_row(req, algo)
        if row is not None:
            set_leaf_text(row_cells(row)[3], "O")
    if named_row(req, "signing_server") is None:
        clone_row(req, "signing_reason", "signing_server", "String",
                  "RC (When more than one connector shares the same connector name)",
                  "Name of the signing server used to disambiguate the connector.",
                  after="connector_name")
    new_fields = [
        ("citizen_id", "String", "O",
         "Citizen identity number of the signer, forwarded to the signing provider."),
        ("webhook_url", "String", "O",
         "Webhook URL that receives the signing result callback."),
        ("rssp_operation_mode", "String", "O",
         "Operation mode of the signing request. Possible values are: + S (synchronous, "
         "default) + A (asynchronous)"),
        ("is_redirect_app_enabled", "Boolean", "O",
         "If true, the signing session is treated as running on a mobile channel and the "
         "provider app redirect is enabled."),
        ("notification_template", "String", "O",
         "Custom notification template used for the authorization message. The default "
         "template is used when omitted."),
        ("metadata", "String", "O",
         "Free-form metadata from the client, stored with the signing request and "
         "returned in status/webhook payloads."),
    ]
    for name, typ, presence, desc in new_fields:
        if named_row(req, name) is None:
            clone_row(req, "signing_reason", name, typ, presence, desc)
    renumber(req)

    frame = one(frames_in(blocks, "POST "), "request frame in sign_rssp")
    rewrite_body(frame, [
        "{",
        '  "connector_name": "NameOfConnector",',
        '  "signing_server": "NameOfSigningServer",',
        '  "field_name": "SignatureField1",',
        '  "credential_id": "String",',
        '  "skip_verification": false,',
        '  "hand_signature_image": "Base64 encoded image",',
        '  "signing_reason": "String",',
        '  "signing_location": "String",',
        '  "contact_information": "signer@example.com",',
        '  "signing_algo": "String",',
        '  "hashing_algo": "String",',
        '  "appearance_design": "QUALIFIED_ELECTRONIC_SIGNATURE",',
        '  "citizen_id": "String",',
        '  "webhook_url": "String",',
        '  "rssp_operation_mode": "S",',
        '  "is_redirect_app_enabled": false,',
        '  "notification_template": "String",',
        '  "metadata": "String"',
        "}",
    ])
    frame = one(frames_in(blocks, "Http Code: 200"), "200 frame in sign_rssp")
    rewrite_body(frame, [
        "{",
        '  "validity_period": 300,',
        '  "vc_code": "12"',
        "}",
    ])
    print("3) Sign Document via RSSP: contact_information renamed, time_zone dropped, "
          "algos M->O, %d optional fields ensured, samples rewritten"
          % (len(new_fields) + 1))

    # --- 4) RSSP Credential Authorization ----------------------------------
    blocks = section_blocks(body, "RSSP Credential Authorization")
    req = find_table(blocks, "Request Attributes")
    for name, desc in (
            ("otp", "One-time password used to authorize the signing. " + ONE_OF_THREE),
            ("pin", "PIN code used to authorize the signing. " + ONE_OF_THREE)):
        row = named_row(req, name)
        if row is not None:
            set_leaf_text(row_cells(row)[4], desc)
    if named_row(req, "face_image") is None:
        clone_row(req, "otp", "face_image", "String", "O/M",
                  "Base64 encoded face image used to authorize the signing via face "
                  "matching. " + ONE_OF_THREE, after="pin")
    if named_row(req, "responseID") is None:
        clone_row(req, "otp", "responseID", "String", "O",
                  "Smart polling. When provided, the call is routed to the sign-hash "
                  "polling flow and waits for the signature result of that request; the "
                  "authentication fields above are ignored.")
    renumber(req)
    frame = one(frames_in(blocks, "POST "), "request frame in authorization")
    rewrite_body(frame, [
        "{",
        '  "credential_id": "String",',
        '  "package_id": 0,',
        '  "document_id": 0,',
        '  "field_name": "String",',
        '  "otp": "String",',
        '  "pin": "String",',
        '  "face_image": "Base64 encoded image"',
        "}",
    ])
    print("4) RSSP Credential Authorization: face_image + responseID documented")

    # --- 5) Get Signing Status (TSE MID) ------------------------------------
    blocks = section_blocks(body, "Get Signing Status (TSE MID)")
    req = find_table(blocks, "Request Attributes")
    row = require_row(req, "transaction_id", "Get Signing Status (TSE MID)")
    set_text_around_link(row_cells(row)[4], TX_BEFORE, TX_AFTER)
    print("5) Get Signing Status (TSE MID): transaction_id format documented")

    # --- 6) Get Credentials Info --------------------------------------------
    blocks = section_blocks(body, "Get Credentials Info")
    req = find_table(blocks, "Request Attributes")
    row = require_row(req, "connector_name", "Get Credentials Info")
    set_leaf_text(row_cells(row)[3], "RC (When credential_id is provided)")
    resp = find_table(blocks, "Response Attributes")
    if named_row(resp, "fields") is None:
        clone_row(resp, "auth_mode", "fields", "Object", "M",
                  "All fields parsed from the X.509 certificate, keyed by field name.",
                  after="multisign")
    renumber(resp)
    frame = one(frames_in(blocks, "Http Code: 200"), "200 frame in credentials info")
    rewrite_body(frame, [
        "{",
        '  "cert": {',
        '    "cert": "Base64 encoded certificate",',
        '    "mode": "TSE",',
        '    "status": "OPERATED",',
        '    "status_desc": "String",',
        '    "issuer_dn": {},',
        '    "subject_dn": {},',
        '    "valid_from": "String",',
        '    "valid_to": "String"',
        "  },",
        '  "auth_mode": "String",',
        '  "SCAL": 1,',
        '  "multisign": 1,',
        '  "fields": {}',
        "}",
    ])
    print("6) Get Credentials Info: connector_name RC, fields row, 200 sample rewritten")

    # --- 7) Send OTP via RSSP ------------------------------------------------
    blocks = section_blocks(body, "Send OTP via RSSP")
    req = find_table(blocks, "Request Attributes")
    row = require_row(req, "transaction_id", "Send OTP via RSSP")
    set_text_around_link(row_cells(row)[4], TX_BEFORE, TX_AFTER)
    if named_row(req, "notification_template") is None:
        clone_row(req, "transaction_id", "notification_template", "String", "O",
                  "Custom notification template used for the OTP message. The default "
                  "template is used when omitted.")
    renumber(req)
    resp = find_table(blocks, "Response Attributes")
    if named_row(resp, "validity_period") is None:
        clone_row(resp, "error_code", "validity_period", "Integer", "M",
                  "Validity period of the OTP in seconds.")
        clone_row(resp, "validity_period", "user_email", "String", "M",
                  "Email of the user the OTP was sent to.", after="validity_period")
    renumber(resp)
    frame = one(frames_in(blocks, "Http Code: 200"), "200 frame in send OTP")
    rewrite_body(frame, [
        "{",
        '  "validity_period": 300,',
        '  "user_email": "user@example.com"',
        "}",
    ])
    print("7) Send OTP via RSSP: notification_template + response fields + sample")

    # --- Objects --------------------------------------------------------------
    for obj, add_status_desc in (("GetFilteredCredentialListResponse", False),
                                 ("CredentialInfo", True)):
        blocks = section_blocks(body, obj)
        tbl = next(b for b in blocks if b.tag == qn("tbl"))
        for dn in ("issuer_dn", "subject_dn"):
            row = require_row(tbl, dn, obj)
            cell = row_cells(row)[2]
            if list(cell.iter(qn("hyperlink"))):
                print("   WARN: %s.%s type cell is a hyperlink - left alone" % (obj, dn))
                continue
            set_leaf_text(cell, "Object")
        if add_status_desc and named_row(tbl, "status_desc") is None:
            clone_row(tbl, "status", "status_desc", "String", "M",
                      "Human readable description of the credential status.",
                      after="status")
        renumber(tbl)
        print("O) %s: DN fields -> Object%s" % (
            obj, ", status_desc added" if add_status_desc else ""))

    if dry:
        print("\n--dry-run: nothing written")
        return

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
