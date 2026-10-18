#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Enterprise Management doc sync: add the 24 endpoints present in the backend
`validURLs` registry (EnterpriseController) but missing from the specification
DOCX, fix the Delete Enterprise User section to the real backend route, and
remove the Change User Service Plan section from Enterprise Management (kept in
ADMIN APIs) per owner decisions of 2026-07-27.

Write-by-cloning: every new section clones a donor of the same table shape;
the request-frame first line keeps its badge/red/black run structure
(template 2.8.1) and rewrite_url_line also recolors the badge when the method
changes. Identifier hygiene: bookmarks stripped, docPr/VML ids reallocated.

Facts come from the backend (EnterpriseController/EnterpriseService, DTOs,
serializers) as verified on 2026-07-27 — not from the model. Client-credentials
tokens are blocked module-wide; owner-only endpoints are marked in prose.

Usage:
    python scripts/oneoff/add_enterprise_management_sections.py [--dry-run]
"""
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from docx_guard import load, save  # noqa: E402

UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

BADGE = {"GET": "047857", "POST": "C2410C", "PUT": "1D4ED8", "DELETE": "B91C1C"}
METHODS = ("GET", "POST", "PUT", "DELETE", "PATCH")

DONOR_LIST = "Get User Activity Logs"       # Header + Path + Request + Response
DONOR_FLAT = "Account Usage Statistics"     # Header + Response
DONOR_POST = "Add Identity for a User"      # Header + Request + Response

PAGE_ROWS_M = [
    ("pageNo", "Integer", "M", "Page number to be retrieved (1-based)."),
    ("recordsPerPage", "Integer", "M", "Number of records to be retrieved in one page."),
]
PAGE_ROWS_O = [
    ("pageNo", "Integer", "O",
     "Page number to be retrieved (1-based). When the pagination segments are "
     "omitted the full result is exported."),
    ("recordsPerPage", "Integer", "O", "Number of records to be retrieved in one page."),
]
NO_CONTENT = ["Http Code: 204", "Content-Type: application/json", "{}"]
EMPTY_OK = ["Http Code: 200", "Content-Type: application/json", "{}"]
BINARY_OK = ["Http Code: 200", "Content-Type: application/octet-stream",
             "{", "  File(binary)", "}"]
INVITE_BODY = ["{", '  "user_email": "user@example.com",',
               '  "user_first_name": "Nguyen",',
               '  "user_last_name": "Van A",',
               '  "enterprise_role": "Enterprise User"', "}"]
INVITE_ROWS = [
    ("user_email", "String", "M",
     "Email of the user to be invited. Must be a valid email address."),
    ("user_first_name", "String", "O", "First name of the invited user."),
    ("user_last_name", "String", "O", "Last name of the invited user."),
    ("user_name", "String", "O",
     "Full name of the invited user. Derived from the first and last name "
     "when omitted."),
    ("enterprise_role", "String", "O",
     "Name of the enterprise role to assign. The system default role is "
     "assigned when omitted or not matched."),
    ("expired_time", "String", "O",
     "Expiry time of the invitation, ISO 8601. Defaults to 5 minutes from "
     "the invitation time."),
]

SPECS = [
    # ---- invitations cluster (after Register Enterprise User) ---------------
    dict(
        name="Agent Invite Enterprise User",
        donor=DONOR_POST, anchor="Register Enterprise User", method="POST",
        url="{baseURL}/v1/enterprise/users/agent-invite",
        prose=("Business applications can use this service API to invite a user into "
               "the enterprise on behalf of an agent. When the invited email does not "
               "exist in the system yet, a new user account is provisioned "
               "automatically and no registration email is sent. A successful call "
               "returns HTTP 204 with no content."),
        body_lines=INVITE_BODY, request_rows=INVITE_ROWS,
        response_rows=[], response_sample=NO_CONTENT,
    ),
    dict(
        name="Invite Enterprise User",
        donor=DONOR_POST, anchor="Agent Invite Enterprise User", method="POST",
        url="{baseURL}/v1/enterprise/invitations",
        prose=("Business applications can use this service API to invite an existing "
               "user of the system into the enterprise. The invited email must "
               "already belong to a registered user. A successful call returns HTTP "
               "204 with no content and an invitation email is sent to the user."),
        body_lines=INVITE_BODY, request_rows=INVITE_ROWS,
        response_rows=[], response_sample=NO_CONTENT,
    ),
    dict(
        name="Get Enterprise Invitations",
        donor=DONOR_LIST, anchor="Invite Enterprise User", method="GET",
        url="{baseURL}/v1/enterprise/invitations/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the paged list "
               "of enterprise invitations. Pending invitations past their validity "
               "period are returned as EXPIRED. The x-total-records response header "
               "contains the total number of records."),
        path_rows=PAGE_ROWS_M, drop_tables=["Request Attributes"],
        header_add=[("x-search-text", "String", "O",
                     "Optional search text, Base64-encoded UTF-8. Invalid Base64 is "
                     "rejected.")],
        response_rows=[
            ("id", "Integer", "M", "The ID of the invitation."),
            ("created_at", "String", "M", "Creation time of the invitation, UTC."),
            ("user_email", "String", "M", "Email of the invited user."),
            ("user_name", "String", "M", "Full name of the invited user."),
            ("user_name_alias", "String", "O", "Alias name of the invited user."),
            ("enterprise_role", "String", "M", "Assigned enterprise role name."),
            ("invited_by", "String", "M", "Email of the user who sent the invitation."),
            ("status", "String", "M",
             "Status of the invitation: PENDING, ACCEPTED, REJECTED, EXPIRED, "
             "CANCELLED or DIRECT_ASSIGNED."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "id": 12,',
                         '    "created_at": "2026-07-20T04:00:00Z",',
                         '    "user_email": "user@example.com",',
                         '    "user_name": "Nguyen Van A",',
                         '    "enterprise_role": "Enterprise User",',
                         '    "invited_by": "admin@example.com",',
                         '    "status": "PENDING"', "  }", "]"],
    ),
    dict(
        name="Resend Enterprise User Invitation",
        donor=DONOR_LIST, anchor="Get Enterprise Invitations", method="POST",
        url="{baseURL}/v1/enterprise/invitations/{invitationId}/resend",
        prose=("Business applications can use this service API to resend a pending or "
               "expired enterprise invitation. The invitation expiry is renewed and "
               "the invitation email is sent again. This API is rate limited to 1 "
               "request per 5 minutes per invitation (HTTP 429 with a Retry-After "
               "header when exceeded) and each invitation can be resent at most 3 "
               "times. A successful call returns HTTP 204 with no content."),
        path_rows=[("invitationId", "Integer", "M", "The ID of the invitation to resend.")],
        drop_tables=["Request Attributes"],
        response_rows=[], response_sample=NO_CONTENT,
    ),
    dict(
        name="Cancel Enterprise User Invitation",
        donor=DONOR_LIST, anchor="Resend Enterprise User Invitation", method="DELETE",
        url="{baseURL}/v1/enterprise/invitations/{invitationId}",
        prose=("Business applications can use this service API to cancel a pending "
               "enterprise invitation. Only the enterprise owner can call this API. "
               "The invitation must be pending and not expired; the cancellation is "
               "a status update, the invitation record is kept."),
        path_rows=[("invitationId", "Integer", "M", "The ID of the invitation to cancel.")],
        request_rows=[("user_email", "String", "O",
                       "When provided, must match the email of the invited user; "
                       "the request is rejected on mismatch.")],
        response_rows=[], response_sample=EMPTY_OK,
    ),
    dict(
        name="Delete Enterprise User Invitation",
        donor=DONOR_LIST, anchor="Cancel Enterprise User Invitation", method="DELETE",
        url="{baseURL}/v1/enterprise/invitations/{invitationId}/delete",
        prose=("Business applications can use this service API to delete a pending "
               "enterprise invitation from the invitation list. The operation is a "
               "status update to CANCELLED; the invitation record is not physically "
               "removed. Unlike Cancel Enterprise User Invitation, this API is not "
               "restricted to the enterprise owner."),
        path_rows=[("invitationId", "Integer", "M", "The ID of the invitation to delete.")],
        drop_tables=["Request Attributes"],
        response_rows=[], response_sample=EMPTY_OK,
    ),
    dict(
        name="Export Enterprise Invitations",
        donor=DONOR_LIST, anchor="Delete Enterprise User Invitation", method="GET",
        url="{baseURL}/v1/enterprise/invitations/export/{status}/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to export the "
               "enterprise invitation list as a file. The status and pagination "
               "segments are optional; without pagination the full list is exported. "
               "The response is a binary file with the x-total-records, x-file-name "
               "and Content-Disposition headers; the file name follows "
               "EnterpriseInvitations_yyyyMMdd_HHmmss with the chosen extension."),
        path_rows=[
            ("status", "String", "O",
             "Optional status filter: ALL, PENDING, ACCEPTED, REJECTED, EXPIRED or "
             "CANCELLED. Defaults to ALL."),
        ] + PAGE_ROWS_O,
        drop_tables=["Request Attributes"],
        header_add=[
            ("x-search-text", "String", "O", "Optional search text, Base64-encoded UTF-8."),
            ("x-role-id", "String", "O", "Filter by enterprise role ID."),
            ("x-from-date", "String", "O", "Lower bound of the created date, ISO 8601."),
            ("x-to-date", "String", "O", "Upper bound of the created date, ISO 8601."),
            ("x-sort-by", "String", "O",
             "Sort column: CREATEDON, NAME, ROLE or STATUS."),
            ("x-asc", "String", "O", "true for ascending order, false for descending."),
            ("x-export-format", "String", "O",
             "Export file format: xlsx (default) or csv."),
        ],
        response_rows=[], response_sample=BINARY_OK,
    ),
    # ---- users cluster (after Get Enterprise Users) -------------------------
    dict(
        name="Export Enterprise Users",
        donor=DONOR_LIST, anchor="Get Enterprise Users", method="GET",
        url="{baseURL}/v1/enterprise/users/export/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to export the "
               "enterprise user list as a file. The pagination segments are "
               "optional; without them the full list is exported. The response is a "
               "binary file with the x-total-records, x-file-name and "
               "Content-Disposition headers; the file name follows "
               "EnterpriseUsers_yyyyMMdd_HHmmss with the chosen extension."),
        path_rows=PAGE_ROWS_O, drop_tables=["Request Attributes"],
        header_add=[
            ("x-search-text", "String", "O", "Optional search text, Base64-encoded UTF-8."),
            ("x-role-id", "String", "O", "Filter by enterprise role ID."),
            ("x-from-date", "String", "O", "Lower bound of the created date, ISO 8601."),
            ("x-to-date", "String", "O", "Upper bound of the created date, ISO 8601."),
            ("x-export-format", "String", "O",
             "Export file format: xlsx (default) or csv."),
        ],
        response_rows=[], response_sample=BINARY_OK,
    ),
    dict(
        name="Get Enterprise User Quota",
        donor=DONOR_LIST, anchor="Export Enterprise Users", method="GET",
        url="{baseURL}/v1/enterprise/users/{enterpriseUserId}/quota",
        prose=("Business applications can use this service API to get the quota of "
               "one enterprise user. All quota values are serialized as strings; "
               "unlimited quotas are returned as Unlimited."),
        path_rows=[("enterpriseUserId", "Integer", "M",
                    "The ID of the enterprise user whose quota is retrieved.")],
        drop_tables=["Request Attributes"],
        response_rows=[
            ("user_id", "Integer", "M", "The ID of the enterprise user."),
            ("enterprise_id", "Integer", "M", "The ID of the enterprise."),
            ("quota", "Object", "M",
             "Quota of the user. Contains signature, workflow, template, storage "
             "and upload_size, each with allowed, consumed and remaining values."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "user_id": 452,',
                         '  "enterprise_id": 118,',
                         '  "quota": {',
                         '    "signature": {',
                         '      "allowed": "100",',
                         '      "consumed": "45",',
                         '      "remaining": "55"',
                         "    },",
                         '    "workflow": {},',
                         '    "template": {},',
                         '    "storage": {},',
                         '    "upload_size": {}', "  }", "}"],
    ),
    dict(
        name="Get Enterprise User Management Statistics",
        donor=DONOR_FLAT, anchor="Get Enterprise User Quota", method="GET",
        url="{baseURL}/v1/enterprise/users/statistics",
        prose=("Business applications can use this service API to get the user "
               "management statistics of the enterprise. All values are serialized "
               "as strings."),
        response_rows=[
            ("enterprise_users", "String", "M", "Total number of enterprise users."),
            ("pending_invitations", "String", "M", "Number of pending invitations."),
            ("signature_used", "String", "M", "Number of consumed signatures."),
            ("storage_used", "String", "M", "Consumed storage."),
            ("document_storage", "String", "M", "Consumed document storage."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "enterprise_users": "24",',
                         '  "pending_invitations": "3",',
                         '  "signature_used": "150",',
                         '  "storage_used": "1048576",',
                         '  "document_storage": "524288"', "}"],
    ),
    # ---- settings clusters (after Add Role, end of section) -----------------
    dict(
        name="Get Enterprise Custom Fields",
        donor=DONOR_LIST, anchor="Add Role", method="GET",
        url="{baseURL}/v1/enterprise/custom_fields/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the custom "
               "field types of the enterprise, sorted by creation time descending. "
               "The pagination segments are optional. The x-total-records response "
               "header contains the total number of records after filtering."),
        path_rows=PAGE_ROWS_O, drop_tables=["Request Attributes"],
        header_add=[
            ("x-parent-type", "String", "O",
             "Filter by parent field type, exact case-insensitive match."),
            ("x-search-type-name", "String", "O",
             "Search on the type name, Base64-encoded UTF-8."),
        ],
        response_rows=[
            ("id", "String", "M", "Server-generated ID of the custom field."),
            ("field_name", "String", "O", "Legacy identifier of the custom field."),
            ("type_name", "String", "M", "Display name of the custom field type."),
            ("parent_type", "String", "M", "Parent field type. Currently always TEXTBOX."),
            ("format_type", "String", "O", "DATE or DATETIME for date-like fields."),
            ("place_holder", "String", "O", "Placeholder text of the field."),
            ("suffix", "String", "O", "Suffix of the field."),
            ("description", "String", "O", "Description of the field."),
            ("created_at", "String", "M", "Creation time, UTC."),
            ("modified_at", "String", "O", "Last modification time, UTC."),
            ("date", "Object", "O",
             "Date constraints for date-like fields: minimum_date, maximum_date, "
             "default_date_enabled, default_date and date_format."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "id": "8f3k2a1b-1c2d-4e5f-8a9b-0c1d2e3f4a5b",',
                         '    "type_name": "Contract Number",',
                         '    "parent_type": "TEXTBOX",',
                         '    "place_holder": "Enter contract number",',
                         '    "description": "Internal contract number",',
                         '    "created_at": "2026-07-27T04:00:00.000Z"', "  }", "]"],
    ),
    dict(
        name="Create Enterprise Custom Field",
        donor=DONOR_POST, anchor="Get Enterprise Custom Fields", method="POST",
        url="{baseURL}/v1/enterprise/custom_fields",
        prose=("Business applications can use this service API to create a custom "
               "field type for the enterprise. Only the enterprise owner can call "
               "this API. A successful call returns HTTP 201 with the created "
               "custom field."),
        body_lines=["{", '  "type_name": "Contract Number",',
                    '  "description": "Internal contract number",',
                    '  "place_holder": "Enter contract number"', "}"],
        request_rows=[
            ("type_name", "String", "M",
             "Name of the custom field type. Letters, digits, underscore and "
             "space only; must be unique in the enterprise; maximum 50 characters."),
            ("description", "String", "O", "Description, maximum 200 characters."),
            ("place_holder", "String", "O", "Placeholder text, maximum 200 characters."),
            ("format_type", "String", "O",
             "DATE or DATETIME to create a date-like field; other values are ignored."),
            ("suffix", "String", "O", "Suffix of the field."),
            ("date", "Object", "O",
             "Date constraints, honored only for date-like fields: minimum_date, "
             "maximum_date, default_date_enabled, default_date and date_format."),
        ],
        response_rows=[
            ("id", "String", "M", "Server-generated ID of the created custom field."),
            ("type_name", "String", "M", "Name of the custom field type."),
            ("parent_type", "String", "M", "Parent field type, always TEXTBOX."),
            ("created_at", "String", "M", "Creation time, UTC."),
        ],
        response_sample=["Http Code: 201", "Content-Type: application/json", "{",
                         '  "id": "8f3k2a1b-1c2d-4e5f-8a9b-0c1d2e3f4a5b",',
                         '  "type_name": "Contract Number",',
                         '  "parent_type": "TEXTBOX",',
                         '  "place_holder": "Enter contract number",',
                         '  "description": "Internal contract number",',
                         '  "created_at": "2026-07-27T04:00:00.000Z"', "}"],
    ),
    dict(
        name="Update Enterprise Custom Field",
        donor=DONOR_LIST, anchor="Create Enterprise Custom Field", method="PUT",
        url="{baseURL}/v1/enterprise/custom_fields/{fieldId}",
        prose=("Business applications can use this service API to update a custom "
               "field type of the enterprise. Only the enterprise owner can call "
               "this API. Only the fields present in the request body are "
               "overwritten (merge behavior); the field name identity is immutable."),
        path_rows=[("fieldId", "String", "M",
                    "The ID of the custom field to update (the server-generated id; "
                    "the legacy field_name is accepted as a fallback).")],
        body_lines=["{", '  "description": "Updated description"', "}"],
        request_rows=[
            ("type_name", "String", "O", "New name, same constraints as create."),
            ("description", "String", "O", "Description, maximum 200 characters."),
            ("place_holder", "String", "O", "Placeholder text, maximum 200 characters."),
            ("format_type", "String", "O", "DATE or DATETIME; other values are ignored."),
            ("suffix", "String", "O", "Suffix of the field."),
            ("date", "Object", "O", "Date constraints for date-like fields."),
        ],
        response_rows=[
            ("id", "String", "M", "The ID of the custom field."),
            ("type_name", "String", "M", "Name of the custom field type."),
            ("modified_at", "String", "M", "Last modification time, UTC."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "id": "8f3k2a1b-1c2d-4e5f-8a9b-0c1d2e3f4a5b",',
                         '  "type_name": "Contract Number",',
                         '  "description": "Updated description",',
                         '  "modified_at": "2026-07-27T05:00:00.000Z"', "}"],
    ),
    dict(
        name="Delete Enterprise Custom Field",
        donor=DONOR_LIST, anchor="Update Enterprise Custom Field", method="DELETE",
        url="{baseURL}/v1/enterprise/custom_fields/{fieldId}",
        prose=("Business applications can use this service API to delete a custom "
               "field type of the enterprise. Only the enterprise owner can call "
               "this API. A successful call returns HTTP 204 with no content."),
        path_rows=[("fieldId", "String", "M",
                    "The ID of the custom field to delete (the server-generated id; "
                    "the legacy field_name is accepted as a fallback).")],
        drop_tables=["Request Attributes"],
        response_rows=[], response_sample=NO_CONTENT,
    ),
    dict(
        name="Get Enterprise Electronic Seals",
        donor=DONOR_FLAT, anchor="Delete Enterprise Custom Field", method="GET",
        url="{baseURL}/v1/enterprise/electronic-seals",
        prose=("Business applications can use this service API to get the electronic "
               "seal configuration of the enterprise. The response is an array "
               "containing the single configuration of the enterprise."),
        response_rows=[
            ("id", "Integer", "M", "The ID of the enterprise."),
            ("basic_information", "Object", "M",
             "Basic seal information: name, role, assurance_level, signing_server, "
             "certificate_id, certificate_pin and connector_name."),
            ("appearance", "Object", "M",
             "Seal appearance: design, hand_signature and logo."),
            ("detail", "Object", "M",
             "Seal details: signer_name, reason, location and contact_info."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "id": 118,',
                         '    "basic_information": {',
                         '      "name": "ABC Corporation Seal",',
                         '      "assurance_level": "AES",',
                         '      "connector_name": "MOBILE_ID_ESEAL"',
                         "    },",
                         '    "appearance": {},',
                         '    "detail": {}', "  }", "]"],
    ),
    dict(
        name="Update Enterprise Electronic Seals",
        donor=DONOR_POST, anchor="Get Enterprise Electronic Seals", method="PUT",
        url="{baseURL}/v1/enterprise/electronic-seals",
        prose=("Business applications can use this service API to update the "
               "electronic seal configuration of the enterprise. Only the "
               "enterprise owner can call this API. Only the fields present in the "
               "request are merged into the existing configuration; base64 logo and "
               "hand signature images are uploaded to the file management service."),
        body_lines=["{", '  "id": 118,',
                    '  "detail": {',
                    '    "signer_name": "ABC Corporation",',
                    '    "location": "Ho Chi Minh City"', "  }", "}"],
        request_rows=[
            ("id", "Integer", "M", "The ID of the enterprise. Must be greater than 0."),
            ("basic_information", "Object", "O",
             "Basic seal information to merge: name, role, assurance_level, "
             "signing_server, certificate_id, certificate_pin and connector_name."),
            ("appearance", "Object", "O",
             "Seal appearance to merge: design, hand_signature and logo (base64)."),
            ("detail", "Object", "O",
             "Seal details to merge: signer_name, reason, location and contact_info."),
        ],
        response_rows=[], response_sample=EMPTY_OK,
    ),
    dict(
        name="Get Enterprise TSA Configuration",
        donor=DONOR_FLAT, anchor="Update Enterprise Electronic Seals", method="GET",
        url="{baseURL}/v1/enterprise/settings/tsa",
        prose=("Business applications can use this service API to get the timestamp "
               "authority (TSA) configuration of the enterprise. When no "
               "configuration exists yet, an empty configuration object is "
               "returned. The authentication password is returned RSA-encrypted, "
               "never in plaintext."),
        response_rows=[
            ("tsa_uri", "String", "O", "URI of the timestamp authority."),
            ("policy_oid", "String", "O", "TSA policy OID."),
            ("hash_algorithm", "String", "O", "Hashing algorithm: SHA256, SHA384 or SHA512."),
            ("timeout", "Integer", "O", "Request timeout in milliseconds."),
            ("retry_count", "Integer", "O", "Number of retries on failure."),
            ("authentication", "Object", "M",
             "Authentication settings: enabled, username and password "
             "(RSA-encrypted)."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "tsa_uri": "https://tsa.example.com/tsa",',
                         '  "policy_oid": "1.2.3.4.5",',
                         '  "hash_algorithm": "SHA256",',
                         '  "timeout": 5000,',
                         '  "retry_count": 3,',
                         '  "authentication": {',
                         '    "enabled": true,',
                         '    "username": "tsa-user",',
                         '    "password": "{encrypted}"', "  }", "}"],
    ),
    dict(
        name="Add Enterprise TSA Setting",
        donor=DONOR_POST, anchor="Get Enterprise TSA Configuration", method="POST",
        url="{baseURL}/v1/enterprise/settings/tsa",
        prose=("Business applications can use this service API to add a timestamp "
               "authority (TSA) setting for the enterprise. Only the enterprise "
               "owner can call this API. A successful call returns HTTP 201 with "
               "the created setting."),
        body_lines=["{", '  "name": "Primary TSA",',
                    '  "url": "https://tsa.example.com/tsa",',
                    '  "is_default": true,',
                    '  "authentication": {',
                    '    "type": "BASIC",',
                    '    "username": "tsa-user",',
                    '    "password": "{password}"', "  }", "}"],
        request_rows=[
            ("name", "String", "M", "Display name of the TSA setting."),
            ("url", "String", "M", "TSA endpoint URL, http or https."),
            ("is_default", "Boolean", "M", "True to use this setting as the default."),
            ("authentication", "Object", "O",
             "Authentication settings. type: NONE, BASIC or CERTIFICATE; username "
             "and password are required for BASIC; client_certificate (base64) is "
             "required for CERTIFICATE."),
            ("policy_oid", "String", "O", "TSA policy OID."),
            ("hash_algorithm", "String", "O",
             "SHA256 (default), SHA384 or SHA512."),
            ("timeout", "Integer", "O",
             "Request timeout in milliseconds, greater than 0. Defaults to 5000."),
            ("tsa_users", "Array", "O", "Optional TSA user entries of the setting."),
        ],
        response_rows=[
            ("id", "Integer", "M", "The ID of the created TSA setting."),
            ("name", "String", "M", "Display name of the setting."),
            ("url", "String", "M", "TSA endpoint URL."),
            ("authentication_type", "String", "M", "NONE, BASIC or CERTIFICATE."),
            ("hash_algorithm", "String", "M", "Hashing algorithm of the setting."),
            ("is_default", "Boolean", "M", "True when the setting is the default."),
            ("created_on", "String", "M", "Creation time of the setting."),
            ("status", "String", "M", "Status of the setting."),
        ],
        response_sample=["Http Code: 201", "Content-Type: application/json", "{",
                         '  "id": 7,',
                         '  "name": "Primary TSA",',
                         '  "url": "https://tsa.example.com/tsa",',
                         '  "authentication_type": "BASIC",',
                         '  "hash_algorithm": "SHA256",',
                         '  "is_default": true,',
                         '  "created_on": "2026-07-27T04:00:00Z",',
                         '  "status": "ACTIVE"', "}"],
    ),
    dict(
        name="Update Enterprise TSA Configuration",
        donor=DONOR_POST, anchor="Add Enterprise TSA Setting", method="PUT",
        url="{baseURL}/v1/enterprise/settings/tsa",
        prose=("Business applications can use this service API to update the "
               "timestamp authority (TSA) configuration of the enterprise. Only the "
               "enterprise owner can call this API. Only the fields present in the "
               "request are merged into the stored configuration. The TSA endpoint "
               "is probed before saving; changed credentials are verified against "
               "the endpoint."),
        body_lines=["{", '  "tsa_uri": "https://tsa.example.com/tsa",',
                    '  "hash_algorithm": "SHA256",',
                    '  "timeout": 5000', "}"],
        request_rows=[
            ("tsa_uri", "String", "O", "TSA endpoint URL, http or https."),
            ("authentication", "Object", "O",
             "Authentication settings to merge: enabled, username and password."),
            ("policy_oid", "String", "O", "TSA policy OID."),
            ("hash_algorithm", "String", "O", "SHA256, SHA384 or SHA512."),
            ("timeout", "Integer", "O", "Request timeout in milliseconds, greater than 0."),
            ("retry_count", "Integer", "O", "Number of retries, greater than 0."),
            ("tsa_users", "Array", "O", "Optional TSA user entries to update."),
        ],
        response_rows=[
            ("tsa_uri", "String", "M", "URI of the timestamp authority."),
            ("policy_oid", "String", "O", "TSA policy OID."),
            ("hash_algorithm", "String", "M", "Hashing algorithm."),
            ("timeout", "Integer", "M", "Request timeout in milliseconds."),
            ("retry_count", "Integer", "O", "Number of retries."),
            ("authentication", "Object", "M",
             "Authentication settings: enabled and username. The password is "
             "never returned."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "tsa_uri": "https://tsa.example.com/tsa",',
                         '  "policy_oid": "1.2.3.4.5",',
                         '  "hash_algorithm": "SHA256",',
                         '  "timeout": 5000,',
                         '  "retry_count": 3,',
                         '  "authentication": {',
                         '    "enabled": true,',
                         '    "username": "tsa-user"', "  }", "}"],
    ),
    dict(
        name="Test Connection Enterprise TSA Configuration",
        donor=DONOR_POST, anchor="Update Enterprise TSA Configuration", method="POST",
        url="{baseURL}/v1/enterprise/settings/tsa/test-connection",
        prose=("Business applications can use this service API to test the "
               "connection of a timestamp authority (TSA) configuration without "
               "saving anything. Only the enterprise owner can call this API. The "
               "request body is optional; when omitted, the stored configuration is "
               "probed. The HTTP status is always 200; the probe outcome is "
               "reported in the response body."),
        body_lines=["{", '  "tsa_uri": "https://tsa.example.com/tsa",',
                    '  "authentication": {',
                    '    "enabled": true,',
                    '    "username": "tsa-user",',
                    '    "password": "{password}"', "  }", "}"],
        request_rows=[
            ("tsa_uri", "String", "O",
             "TSA endpoint URL to probe. Falls back to the stored configuration "
             "when omitted."),
            ("authentication", "Object", "O",
             "Credentials to verify: enabled, username and password. The "
             "credential probe runs only when enabled is true and both values are "
             "provided."),
            ("timeout", "Integer", "O", "Probe timeout in milliseconds. Defaults to 5000."),
        ],
        response_rows=[
            ("status", "String", "M", "Overall probe result: success or failed."),
            ("endpoint", "Object", "M",
             "Endpoint probe result: reachable, http_status, latency_ms and "
             "error_subcode (26 when the endpoint is not accessible)."),
            ("credentials", "Object", "M",
             "Credential probe result: tested, valid and error_subcode (27 when "
             "the credentials are invalid)."),
            ("error_subcode", "String", "O",
             "Subcode of the failing probe, when any."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "status": "success",',
                         '  "endpoint": {',
                         '    "reachable": true,',
                         '    "http_status": 200,',
                         '    "latency_ms": 120',
                         "  },",
                         '  "credentials": {',
                         '    "tested": true,',
                         '    "valid": true', "  }", "}"],
    ),
    dict(
        name="Get Enterprise Report List",
        donor=DONOR_LIST, anchor="Test Connection Enterprise TSA Configuration",
        method="GET",
        url="{baseURL}/v1/enterprise/reports/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the paged "
               "list of generated enterprise reports. The x-total-records response "
               "header contains the total number of records."),
        path_rows=PAGE_ROWS_M, drop_tables=["Request Attributes"],
        header_add=[
            ("x-from-date", "String", "O",
             "Lower bound of the export date, ISO 8601 offset date-time."),
            ("x-to-date", "String", "O",
             "Upper bound of the export date, ISO 8601 offset date-time."),
        ],
        response_rows=[
            ("id", "Integer", "M", "The ID of the report."),
            ("enterprise_id", "Integer", "M", "The ID of the enterprise."),
            ("name", "String", "M", "Name of the report."),
            ("document_uuid", "String", "M", "UUID of the report document."),
            ("dms_property", "String", "O", "Storage property of the report document."),
            ("export_date", "String", "M", "Export time of the report."),
            ("last_modified_by", "String", "O", "Last modifier of the report."),
            ("last_modified_at", "String", "O", "Last modification time."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "id": 31,',
                         '    "enterprise_id": 118,',
                         '    "name": "Workflow statistics July",',
                         '    "document_uuid": "3f2b9c80-2f4e-4d55-9d21-6a1b2c3d4e5f",',
                         '    "export_date": "2026-07-27T04:00:00Z"', "  }", "]"],
    ),
    dict(
        name="Get Enterprise Quota Statistics",
        donor=DONOR_LIST, anchor="Get Enterprise Report List", method="GET",
        url="{baseURL}/v1/enterprise/reports/quota/statistics/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the quota "
               "statistics snapshot of the enterprise. The pagination segments are "
               "required by the route but the response is always a single snapshot "
               "object (x-total-records is 1). Quota values are serialized as "
               "strings; unlimited quotas are returned as Unlimited."),
        path_rows=PAGE_ROWS_M, drop_tables=["Request Attributes"],
        header_add=[("x-search-text", "String", "O",
                     "Optional search text, Base64-encoded UTF-8. Invalid Base64 is "
                     "rejected.")],
        response_rows=[
            ("enterprise_name", "String", "M", "Name of the enterprise."),
            ("enterprise_code", "String", "O", "Code of the enterprise."),
            ("quota", "Object", "M",
             "Quota snapshot. Contains signatures, workflows, templates and "
             "storage, each with allowed, consumed and serviceplan (allowed)."),
            ("total_count", "Object", "M",
             "Consumed counts: signatures, workflows, templates and storage."),
            ("utilization", "Object", "M",
             "Utilization percentages: signatures, workflows, templates and "
             "storage, e.g. 12.50%."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "enterprise_name": "ABC Corporation",',
                         '  "quota": {',
                         '    "signatures": {',
                         '      "allowed": "100",',
                         '      "consumed": "45",',
                         '      "serviceplan": {',
                         '        "allowed": "100"',
                         "      }", "    },",
                         '    "workflows": {},',
                         '    "templates": {},',
                         '    "storage": {}', "  },",
                         '  "total_count": {',
                         '    "signatures": "45"', "  },",
                         '  "utilization": {',
                         '    "signatures": "45.00%"', "  }", "}"],
    ),
    dict(
        name="Export Enterprise Workflow Statistics",
        donor=DONOR_LIST, anchor="Get Enterprise Quota Statistics", method="GET",
        url="{baseURL}/v1/enterprise/reports/workflows/statistics/export/{status}/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to export the "
               "workflow statistics of the enterprise as a file. The status and "
               "pagination segments are optional; without pagination the full "
               "result is exported. When no date range is provided, the last 2 "
               "months are exported. The response is a binary file with the "
               "x-total-records, x-file-name and Content-Disposition headers; the "
               "file name is prefixed Report_ (or BaoCao_ for Vietnamese) followed "
               "by yyyyMMdd_HHmmss and the chosen extension."),
        path_rows=[
            ("status", "String", "O",
             "Optional workflow status filter: ALL (default), DRAFT, SIGNED, "
             "COMPLETED, DECLINED, EXPIRED, EXPIRING, INPROGRESS, PENDING, "
             "UPDATED, APPROVED, SEAL_ONLY or SIGNATURE_ONLY. Unrecognized values "
             "fall back to ALL."),
        ] + PAGE_ROWS_O,
        drop_tables=["Request Attributes"],
        header_add=[
            ("x-search-text", "String", "O",
             "Search on the document package name, Base64-encoded UTF-8."),
            ("x-workflow-ids", "String", "O",
             "Comma-separated workflow IDs to include, e.g. 1,2,3."),
            ("x-workflow-type", "String", "O", "Filter by workflow type, e.g. ECONTRACT."),
            ("x-from-date", "String", "O", "Lower bound of the created date, ISO 8601."),
            ("x-to-date", "String", "O", "Upper bound of the created date, ISO 8601."),
            ("x-sort-by", "String", "O", "Sort column: NAME, CREATED_AT or STATUS."),
            ("x-asc", "String", "O",
             "true for ascending order, false (default) for descending."),
            ("x-export-format", "String", "O",
             "Export file format: xlsx (default), csv, pdf, doc or docx."),
            ("x-vn-moha-compliance", "String", "O",
             "true to generate the Vietnamese MOHA labor compliance report; this "
             "forces the ECONTRACT workflow type and COMPLETED status."),
        ],
        response_rows=[], response_sample=BINARY_OK,
    ),
    dict(
        name="Get Enterprise Workflows",
        donor=DONOR_LIST, anchor="Export Enterprise Workflow Statistics", method="GET",
        url="{baseURL}/v1/enterprise/workflows/{status}/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the paged "
               "list of workflows across the enterprise. The x-total-records "
               "response header contains the total number of records. Status "
               "values without a direct workflow state mapping (SIGNED, EXPIRED, "
               "EXPIRING, UPDATED, APPROVED, SEAL_ONLY, SIGNATURE_ONLY) currently "
               "return the unfiltered list."),
        path_rows=[
            ("status", "String", "M",
             "Workflow status filter: ALL, DRAFT, SIGNED, COMPLETED, DECLINED, "
             "EXPIRED, EXPIRING, INPROGRESS, PENDING, UPDATED, APPROVED, "
             "SEAL_ONLY or SIGNATURE_ONLY (case-insensitive)."),
        ] + PAGE_ROWS_M,
        drop_tables=["Request Attributes"],
        header_add=[
            ("x-search-text", "String", "O",
             "Search on the document package name, Base64-encoded UTF-8."),
            ("x-workflow-ids", "String", "O",
             "Comma-separated workflow IDs to include, e.g. 1,2,3."),
            ("x-workflow-type", "String", "O", "Filter by workflow type, e.g. ECONTRACT."),
            ("x-from-date", "String", "O", "Lower bound of the created date, ISO 8601."),
            ("x-to-date", "String", "O", "Upper bound of the created date, ISO 8601."),
        ],
        response_rows=[
            ("workflows", "List<Object>", "M",
             "Array of enterprise workflows. Each item contains package_id, "
             "workflow_id, package_name, package_owner, owner_name, owner_email, "
             "is_signature_signing_enabled, is_eseal_signing_enabled, "
             "workflow_type, workflow_state, workflow_status, total_comments, "
             "is_shared, total_participants, total_participants_signed, "
             "categories, continue_on_decline, next_signer, deadline_at, "
             "created_by, created_at, modified_at, modified_by, access_duration "
             "(date_from, date_to) and decline (reason, declined_by, "
             "declined_by_order)."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "workflows": [', "    {",
                         '      "package_id": 118,',
                         '      "workflow_id": 92,',
                         '      "package_name": "Contract A",',
                         '      "owner_email": "user@example.com",',
                         '      "workflow_type": "ECONTRACT",',
                         '      "workflow_state": "INPROGRESS",',
                         '      "total_participants": 3,',
                         '      "total_participants_signed": 1,',
                         '      "next_signer": [',
                         '        "signer2@example.com"',
                         "      ],",
                         '      "created_at": "2026-07-27T04:00:00Z",',
                         '      "access_duration": {',
                         '        "date_from": "2026-07-27T04:00:00Z",',
                         '        "date_to": "unlimited"',
                         "      }", "    }", "  ]", "}"],
    ),
]


# ---------------------------------------------------------------------------
# helpers - identical approach to add_account_management_sections.py
# ---------------------------------------------------------------------------
def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def set_para_text(p, text):
    pPr = p.find(qn("pPr"))
    template = None
    for r in p.iter(qn("r")):
        if r.find(qn("t")) is not None:
            template = r
            break
    if template is None:
        raise ValueError("paragraph has no text run: %r" % text_of(p))
    new_run = deepcopy(template)
    for child in list(new_run):
        if child.tag not in (qn("rPr"), qn("t")):
            new_run.remove(child)
    for extra in new_run.findall(qn("t"))[1:]:
        new_run.remove(extra)
    t = new_run.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    for child in list(p):
        if child is not pPr:
            p.remove(child)
    p.append(new_run)


def cell_text(tc):
    return " ".join(x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def set_cell_text(tc, text):
    paras = tc.findall(qn("p"))
    set_para_text(paras[0], text)
    for extra in paras[1:]:
        tc.remove(extra)


def table_title(tbl):
    trs = tbl.findall(qn("tr"))
    if not trs:
        return ""
    return " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()


def find_heading(kids, level, title):
    want = "Heading%d" % level
    for i, e in enumerate(kids):
        if e.tag == qn("p") and style_of(e) == want and text_of(e).strip() == title:
            return i
    raise LookupError("heading not found: %s" % title)


def section_range(kids, start):
    for j in range(start + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and re.match(r"Heading\d$", style_of(e)):
            return start, j
    return start, len(kids)


def strip_bookmarks(elements):
    n = 0
    for el in elements:
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for b in list(el.iter(qn(tag))):
                b.getparent().remove(b)
                n += 1
    return n


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def reassign_ids(elements, next_docpr, next_spid):
    n_doc = n_vml = 0
    for el in elements:
        for d in el.iter():
            if local(d) == "docPr":
                d.set("id", str(next_docpr))
                d.set("name", "Text Box %d" % next_docpr)
                next_docpr += 1
                n_doc += 1
        for d in el.iter():
            if local(d) == "shape":
                d.set("id", "_x0000_s%d" % next_spid)
                if d.get("{%s}spid" % O) is not None:
                    d.set("{%s}spid" % O, "_x0000_s%d" % next_spid)
                next_spid += 1
                n_vml += 1
    return n_doc, n_vml, next_docpr, next_spid


def scan_id_space(root):
    docpr = [int(d.get("id")) for d in root.iter()
             if local(d) == "docPr" and (d.get("id") or "").isdigit()]
    spids = []
    for d in root.iter():
        if local(d) in ("shape", "shapetype"):
            for v in (d.get("id"), d.get("{%s}spid" % O)):
                m = re.match(r"_x0000_s(\d+)$", v or "")
                if m:
                    spids.append(int(m.group(1)))
    return (max(docpr) + 1 if docpr else 1), (max(spids) + 1 if spids else 1000)


def run_text(r):
    return "".join(t.text or "" for t in r.findall(qn("t")))


def run_color(r):
    pr = r.find(qn("rPr"))
    if pr is None:
        return ""
    c = pr.find(qn("color"))
    return (c.get(qn("val")) or "") if c is not None else ""


def make_run(template, text):
    r = deepcopy(template)
    ts = r.findall(qn("t"))
    ts[0].text = text
    ts[0].set(XML_SPACE, "preserve")
    for extra in ts[1:]:
        r.remove(extra)
    return r


def rewrite_url_line(p, method, url):
    runs = p.findall(qn("r"))
    method_run = None
    for r in runs:
        pr = r.find(qn("rPr"))
        if pr is not None and pr.find(qn("shd")) is not None and run_text(r) in METHODS:
            method_run = r
            break
    if method_run is None:
        raise RuntimeError("URL line without a method badge run: %r" % text_of(p))
    red_tpl = next(r for r in runs if run_color(r) == "FF0000")
    black_tpl = next(r for r in runs
                     if run_color(r) == "000000" and run_text(r) not in ("", " "))

    t = method_run.find(qn("t"))
    t.text = method
    method_run.find(qn("rPr")).find(qn("shd")).set(qn("fill"), BADGE[method])

    idx = runs.index(method_run)
    space_run = runs[idx + 1] if idx + 1 < len(runs) and run_text(runs[idx + 1]) == " " else None
    if space_run is None:
        space_run = make_run(black_tpl, " ")
        method_run.addnext(space_run)
    for r in runs:
        if r is method_run or r is space_run:
            continue
        if r.getparent() is p:
            p.remove(r)

    if not url.startswith("{baseURL}"):
        raise RuntimeError("url must start with {baseURL}: %s" % url)
    anchor = space_run
    for token in re.findall(r"\{[^}]+\}|[^{]+", url):
        tpl = red_tpl if token.startswith("{") else black_tpl
        new = make_run(tpl, token)
        anchor.addnext(new)
        anchor = new


def edit_request_frame(txbx, method, url, body_lines):
    url_p = None
    for p in txbx.findall(qn("p")):
        txt = text_of(p).strip()
        if any(txt.startswith(m + " ") for m in METHODS):
            url_p = p
            break
    if url_p is None:
        raise RuntimeError("request frame without a method line")
    rewrite_url_line(url_p, method, url)

    paras = txbx.findall(qn("p"))
    lines = [text_of(p) for p in paras]
    start_idx = paras.index(url_p)
    body_start = next((i for i, l in enumerate(lines) if i > start_idx
                       and l.strip().startswith("{")), None)
    if body_start is None:
        raise RuntimeError("request frame without a body line")
    src = paras[body_start]
    anchor = paras[body_start - 1]
    for p in paras[body_start:]:
        txbx.remove(p)
    for line in (body_lines or ["{}"]):
        new_p = deepcopy(src)
        set_para_text(new_p, line)
        anchor.addnext(new_p)
        anchor = new_p


def replace_frame_lines(txbx, lines):
    paras = txbx.findall(qn("p"))
    src = paras[0]
    anchor = None
    for p in paras:
        txbx.remove(p)
    for line in lines:
        new_p = deepcopy(src)
        set_para_text(new_p, line)
        if anchor is None:
            txbx.insert(0, new_p)
        else:
            anchor.addnext(new_p)
        anchor = new_p


def resize_fill(tbl, rows):
    trs = tbl.findall(qn("tr"))
    template = deepcopy(trs[2])
    for tr in trs[2:]:
        tbl.remove(tr)
    anchor = trs[1]
    for i, (name, typ, presence, desc) in enumerate(rows):
        tr = deepcopy(template)
        tcs = tr.findall(qn("tc"))
        for tc, val in zip(tcs, [str(i + 1), name, typ, presence, desc]):
            set_cell_text(tc, val)
        anchor.addnext(tr)
        anchor = tr


def fill_response_keep_errors(tbl, rows):
    trs = tbl.findall(qn("tr"))
    if len(trs) < 5:
        raise RuntimeError("response table too small: %s" % table_title(tbl))
    template = deepcopy(trs[-1])
    for tr in trs[5:]:
        tbl.remove(tr)
    names = [cell_text(tr.findall(qn("tc"))[1]) for tr in tbl.findall(qn("tr"))[2:5]]
    if names != ["error_description", "error_code", "message"]:
        raise RuntimeError("unexpected leading response rows: %s" % names)
    anchor = tbl.findall(qn("tr"))[-1]
    for name, typ, presence, desc in rows:
        tr = deepcopy(template)
        tcs = tr.findall(qn("tc"))
        for tc, val in zip(tcs, ["", name, typ, presence, desc]):
            set_cell_text(tc, val)
        anchor.addnext(tr)
        anchor = tr
    for i, tr in enumerate(tbl.findall(qn("tr"))[2:]):
        set_cell_text(tr.findall(qn("tc"))[0], str(i + 1))


def append_rows(tbl, rows):
    trs = tbl.findall(qn("tr"))
    template = deepcopy(trs[-1])
    anchor = trs[-1]
    for name, typ, presence, desc in rows:
        tr = deepcopy(template)
        tcs = tr.findall(qn("tc"))
        for tc, val in zip(tcs, ["", name, typ, presence, desc]):
            set_cell_text(tc, val)
        anchor.addnext(tr)
        anchor = tr
    for i, tr in enumerate(tbl.findall(qn("tr"))[2:]):
        set_cell_text(tr.findall(qn("tc"))[0], str(i + 1))


def first_line_of_first_frame(el):
    for txbx in el.iter(qn("txbxContent")):
        for p in txbx.findall(qn("p")):
            txt = text_of(p).strip()
            if txt:
                return txt
    return None


def edit_clone(clone, spec):
    set_para_text(clone[0], spec["name"])
    if clone[1].tag != qn("p") or not text_of(clone[1]).strip():
        raise RuntimeError("donor prose paragraph not where expected: %s" % spec["donor"])
    set_para_text(clone[1], spec["prose"])

    response_done = False
    req_frames = 0
    for el in clone:
        first = first_line_of_first_frame(el)
        if first is None:
            continue
        if any(first.startswith(m + " ") for m in METHODS):
            for txbx in el.iter(qn("txbxContent")):
                edit_request_frame(txbx, spec["method"], spec["url"],
                                   spec.get("body_lines"))
                req_frames += 1
        elif first.startswith("Http Code: 200") and spec.get("response_sample") \
                and not response_done:
            for txbx in el.iter(qn("txbxContent")):
                replace_frame_lines(txbx, spec["response_sample"])
            response_done = True
    if req_frames == 0:
        raise RuntimeError("no request frame found: %s" % spec["name"])
    if spec.get("response_sample") and not response_done:
        raise RuntimeError("no success frame found: %s" % spec["name"])

    for title in spec.get("drop_tables", []):
        for i, el in enumerate(list(clone)):
            if el.tag == qn("tbl") and table_title(el) == title:
                nxt = clone[i + 1] if i + 1 < len(clone) else None
                clone.pop(i)
                if nxt is not None and nxt.tag == qn("p") and not text_of(nxt).strip():
                    clone.remove(nxt)
                break

    for el in list(clone):
        if el.tag != qn("tbl"):
            continue
        title = table_title(el)
        if title == "Path Attributes" and spec.get("path_rows") is not None:
            resize_fill(el, spec["path_rows"])
        elif title == "Request Attributes" and spec.get("request_rows") is not None:
            resize_fill(el, spec["request_rows"])
        elif title == "Response Attributes" and spec.get("response_rows") is not None:
            fill_response_keep_errors(el, spec["response_rows"])
        elif title == "Header Request Attributes" and spec.get("header_add"):
            append_rows(el, spec["header_add"])
    return req_frames


# ---------------------------------------------------------------------------
# fixes for existing sections
# ---------------------------------------------------------------------------
def fix_delete_enterprise_user(kids, path_tbl_donor, blank_donor):
    start, end = section_range(kids, find_heading(kids, 3, "Delete Enterprise User"))
    set_para_text(kids[start + 1],
                  "Business applications can use this service API to remove a user "
                  "from the enterprise. The enterprise user ID is provided in the "
                  "resource URL. The enterprise owner cannot be removed.")
    n = 0
    for i in range(start, end):
        for txbx in kids[i].iter(qn("txbxContent")):
            first = None
            for p in txbx.findall(qn("p")):
                txt = text_of(p).strip()
                if txt:
                    first = txt
                    break
            if first and any(first.startswith(m + " ") for m in METHODS):
                edit_request_frame(
                    txbx, "DELETE",
                    "{baseURL}/v1/enterprise/users/{enterpriseUserId}", None)
                n += 1
    if n == 0:
        raise RuntimeError("delete enterprise user request frame not found")

    # drop the old body table, add a Path Attributes table
    hdr_idx = req_idx = None
    has_path = False
    for i in range(start, end):
        if kids[i].tag != qn("tbl"):
            continue
        title = table_title(kids[i])
        if title == "Header Request Attributes":
            hdr_idx = i
        elif title == "Request Attributes":
            req_idx = i
        elif title == "Path Attributes":
            has_path = True
    if req_idx is not None:
        body = kids[req_idx].getparent()
        nxt = kids[req_idx + 1] if req_idx + 1 < len(kids) else None
        body.remove(kids[req_idx])
        if nxt is not None and nxt.tag == qn("p") and not text_of(nxt).strip():
            body.remove(nxt)
    if not has_path:
        if hdr_idx is None:
            raise RuntimeError("delete enterprise user header table not found")
        new_tbl = deepcopy(path_tbl_donor)
        strip_bookmarks([new_tbl])
        resize_fill(new_tbl, [("enterpriseUserId", "Integer", "M",
                               "The ID of the enterprise user to remove.")])
        blank = deepcopy(blank_donor)
        strip_bookmarks([blank])
        after = kids[hdr_idx + 1]
        after.addnext(new_tbl)
        new_tbl.addnext(blank)
    print("FIXED Delete Enterprise User -> DELETE /users/{enterpriseUserId} (frames=%d)" % n)


def remove_cusp_from_enterprise(body):
    """Remove the Change User Service Plan section inside Enterprise Management."""
    kids = list(body)
    em_start = find_heading(kids, 2, "Enterprise Management")
    em_end = len(kids)
    for j in range(em_start + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and style_of(e) == "Heading2":
            em_end = j
            break
    idx = None
    for i in range(em_start, em_end):
        e = kids[i]
        if e.tag == qn("p") and style_of(e) == "Heading3" \
                and text_of(e).strip() == "Change User Service Plan":
            idx = i
            break
    if idx is None:
        print("SKIP remove: Change User Service Plan not found inside Enterprise Management")
        return
    start, end = section_range(kids, idx)
    for el in kids[start:end]:
        body.remove(el)
    print("REMOVED Change User Service Plan from Enterprise Management (els=%d)" % (end - start))


def pack(src_dir, out_path):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(src_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for base, _dirs, files in os.walk(src_dir):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                if rel != "[Content_Types].xml":
                    z.write(full, rel)
    shutil.move(tmp, out_path)


def main():
    dry = "--dry-run" in sys.argv
    tree, token = load(DOC_XML)
    root = tree.getroot()
    body = root.find(qn("body"))

    kids = list(body)
    al_start, al_end = section_range(kids, find_heading(kids, 3, DONOR_LIST))
    path_tbl_donor = next(kids[i] for i in range(al_start, al_end)
                          if kids[i].tag == qn("tbl")
                          and table_title(kids[i]) == "Path Attributes")
    blank_donor = next(kids[i] for i in range(al_start, al_end)
                       if kids[i].tag == qn("p") and not text_of(kids[i]).strip())

    # 1. fixes / removals first
    fix_delete_enterprise_user(kids, path_tbl_donor, blank_donor)
    remove_cusp_from_enterprise(body)

    # 2. add the 24 new sections
    next_docpr, next_spid = scan_id_space(root)
    for spec in SPECS:
        kids = list(body)
        if any(e.tag == qn("p") and style_of(e) == "Heading3"
               and text_of(e).strip() == spec["name"] for e in kids):
            print("SKIP (exists): %s" % spec["name"])
            continue
        d_start, d_end = section_range(kids, find_heading(kids, 3, spec["donor"]))
        a_start, a_end = section_range(kids, find_heading(kids, 3, spec["anchor"]))

        clone = [deepcopy(e) for e in kids[d_start:d_end]]
        n_bm = strip_bookmarks(clone)
        n_doc, n_vml, next_docpr, next_spid = reassign_ids(clone, next_docpr, next_spid)
        n_frames = edit_clone(clone, spec)

        anchor = kids[a_end - 1]
        for el in clone:
            anchor.addnext(el)
            anchor = el
        print("ADDED %-44s donor=%-26s after=%-40s els=%d frames=%d bm=%d docPr=%d vml=%d"
              % (spec["name"][:44], spec["donor"][:26], spec["anchor"][:40],
                 len(clone), n_frames, n_bm, n_doc, n_vml))

    if dry:
        print("\n--dry-run: nothing written")
        return
    save(tree, DOC_XML, token)
    backup = OUT_DOCX.replace(".docx", "_backup_pre_enterprise_management.docx")
    if os.path.exists(OUT_DOCX) and not os.path.exists(backup):
        shutil.copy2(OUT_DOCX, backup)
        print("backup %s" % backup)
    pack(UNPACKED, OUT_DOCX)
    print("packed %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
