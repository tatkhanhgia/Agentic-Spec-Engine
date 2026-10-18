#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Account Management doc sync: add the 17 endpoints that exist in the backend
`validURLs` registry (UserController) and in the Postman collection but were
missing from the specification DOCX, and align 4 existing sections with the
backend (invitation accept/reject redesign, pagination path, activity-logs path).

Every new section is built by cloning an existing donor section of the same
shape and substituting text, per the "write by cloning" rule. The first-line
run structure of the request frame is preserved (badge run + space + red
{baseURL} + alternating black/red path runs) per format template 2.8.1.

Facts documented here come from the backend, not from the model:
  * routes/methods  - UserController validURLs (byteMap[9..31]).
  * request fields  - InitializeBillingRequest, GetListCertificateEnrollRequest,
                      GetListAuthorizedRequestRequest, AcceptAccountInvitationRequest
                      (@JsonProperty names) and their validate methods.
  * response fields - Dashboard*Serializer, GetListAuthorizedRequestSerializer,
                      GetListCertificateEnrollSerializer, BillingPaymentHistoryResponse,
                      BillingPaymentDetailsResponse, InitializeBillingResponse,
                      AccountQuotaConsumeHistorySerializer, GetAccountEnterprisesResponse,
                      GetAccountProfilesResponse, GetAccountInvitationByIdResponse.
  * invitation accept/reject operate on ONE invitation by invitation_token and
    return 204 No Content (UserService.acceptAccountInvitation / rejectAccountInvitation).
    This is a deliberate divergence from SigningHub (template section 5).

Usage:
    python scripts/oneoff/add_account_management_sections.py [--dry-run]
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

DONOR_LIST = "Get User Activity Logs"       # GET + Path/Request/Response tables
DONOR_FLAT = "Account Usage Statistics"     # GET + Header/Response tables only
DONOR_POST = "Add Identity for a User"      # POST + Request/Response tables

PAGE_ROWS = [
    ("pageNo", "Integer", "M", "Page number to be retrieved (1-based)."),
    ("recordsPerPage", "Integer", "M", "Number of records to be retrieved in one page."),
]
HDR_DATE = [
    ("x-from-date", "String", "O",
     "Optional lower bound of the created date, ISO 8601 format, e.g. 2026-07-01T00:00:00Z."),
    ("x-to-date", "String", "O",
     "Optional upper bound of the created date, ISO 8601 format, e.g. 2026-07-31T23:59:59Z."),
]
HDR_SEARCH = [
    ("x-search-text", "String", "O",
     "Optional search text, Base64-encoded UTF-8. Invalid Base64 is rejected."),
]

# ---------------------------------------------------------------------------
# New section specifications (insertion order matters: anchors created earlier)
# ---------------------------------------------------------------------------
SPECS = [
    dict(
        name="Get Account Invitation By ID",
        donor=DONOR_LIST,
        anchor="Reject Account Invitation",
        method="GET",
        url="{baseURL}/v1/account/invitations/{invitationId}",
        prose=("Business applications can use this service API to get the details of one "
               "enterprise invitation of the logged in user. The invitation is returned "
               "only when the email of the authenticated user matches the invited user. "
               "A pending invitation past its validity period is automatically expired "
               "before it is returned."),
        path_rows=[("invitationId", "String", "M",
                    "ID of the invitation. Either the numeric invitation ID or the "
                    "invitation token reference.")],
        drop_tables=["Request Attributes"],
        response_rows=[
            ("invitation_id", "Integer", "M", "The ID of the invitation."),
            ("enterprise_id", "Integer", "M", "The ID of the inviting enterprise."),
            ("enterprise_name", "String", "M", "Name of the inviting enterprise."),
            ("invited_user_email", "String", "M", "Email of the invited user."),
            ("invited_user_name", "String", "M", "Full name of the invited user."),
            ("role_name", "String", "M", "Name of the enterprise role assigned by the invitation."),
            ("invited_by_email", "String", "M", "Email of the user who sent the invitation."),
            ("status", "String", "M", "Status of the invitation, e.g. PENDING."),
            ("created_at", "String", "M", "Creation time of the invitation."),
            ("expired_at", "String", "M", "Expiry time of the invitation."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "invitation_id": 12,',
                         '  "enterprise_id": 118,',
                         '  "enterprise_name": "ABC Corporation",',
                         '  "invited_user_email": "user@example.com",',
                         '  "invited_user_name": "Nguyen Van A",',
                         '  "role_name": "Enterprise User",',
                         '  "invited_by_email": "admin@example.com",',
                         '  "status": "PENDING",',
                         '  "created_at": "2026-07-20T04:00:00Z",',
                         '  "expired_at": "2026-07-27T04:00:00Z"', "}"],
    ),
    dict(
        name="Get Account Enterprises",
        donor=DONOR_FLAT,
        anchor="Get Account",
        method="GET",
        url="{baseURL}/v1/account/enterprises",
        prose=("Business applications can use this service API to get the list of "
               "enterprises that the logged in user belongs to. Fields with a null "
               "value are omitted from the response."),
        response_rows=[
            ("enterprise_id", "Integer", "M", "The ID of the enterprise."),
            ("enterprise_name", "String", "M", "Name of the enterprise."),
            ("role", "String", "O", "Role of the user in the enterprise."),
            ("is_active", "Boolean", "M", "True when the user is active in the enterprise."),
            ("logo_url", "String", "O", "URL of the enterprise logo. Currently not provided."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "enterprise_id": 118,',
                         '    "enterprise_name": "ABC Corporation",',
                         '    "role": "Enterprise Admin",',
                         '    "is_active": true', "  }", "]"],
    ),
    dict(
        name="Get Account Profiles",
        donor=DONOR_FLAT,
        anchor="Get Account Enterprises",
        method="GET",
        url="{baseURL}/v1/account/profiles",
        prose=("Business applications can use this service API to get the account "
               "profiles available to the logged in user for switching between the "
               "personal and enterprise contexts. A personal profile is always "
               "returned. This API requires a user access token; a client credentials "
               "access token is not allowed."),
        response_rows=[
            ("profile_id", "Integer", "M",
             "0 for the personal profile, otherwise the enterprise user ID."),
            ("enterprise_id", "Integer", "O",
             "The ID of the enterprise. Omitted for the personal profile."),
            ("enterprise_name", "String", "O",
             "Name of the enterprise. Omitted for the personal profile."),
            ("role", "String", "M",
             "Role of the profile. Admin for the personal profile, otherwise the "
             "enterprise role of the user."),
            ("user_email", "String", "M", "Email of the user."),
            ("is_active", "Boolean", "M",
             "True for the profile matching the current access token context."),
            ("logo_url", "String", "O", "URL of the enterprise logo. Currently not provided."),
            ("service_plan", "Object", "O",
             "Service plan of the profile. Returned for the personal profile and for "
             "enterprise profiles where the user is the owner. Contains id, name, "
             "display_name, type, type_remark, industry_type, industry_type_remark, "
             "is_public, payment_type, payment_type_remark, validity_period, "
             "start_date, end_date and billing (billing_cycle, next_billing_date, "
             "price, unit_price, currency)."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "profile_id": 0,',
                         '    "role": "Admin",',
                         '    "user_email": "user@example.com",',
                         '    "is_active": true,',
                         '    "service_plan": {}', "  },", "  {",
                         '    "profile_id": 452,',
                         '    "enterprise_id": 118,',
                         '    "enterprise_name": "ABC Corporation",',
                         '    "role": "Enterprise User",',
                         '    "user_email": "user@example.com",',
                         '    "is_active": false', "  }", "]"],
    ),
    dict(
        name="Get Signing History",
        donor=DONOR_LIST,
        anchor="Document Statistics",
        method="GET",
        url="{baseURL}/v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the signing "
               "transaction history of the logged in user. The optional status segment "
               "filters the transactions by their result. Filter headers override the "
               "same fields in the request body when both are provided. The "
               "x-total-records response header contains the total number of records."),
        path_rows=[
            ("status", "String", "O",
             "Optional status filter. Possible values are ALL, SUCCESS, FAILED, "
             "REJECTED and EXPIRED. Defaults to ALL when omitted."),
        ] + PAGE_ROWS,
        header_add=[
            ("x-criteria-status", "String", "O",
             "Status filter used when the status path segment is omitted."),
            ("x-app-name", "String", "O", "Filter by application name."),
            ("x-citizen-id", "String", "O", "Filter by citizen ID number."),
            ("x-connector-name", "String", "O", "Filter by RSSP connector name."),
            ("x-credential-id", "String", "O", "Filter by credential ID."),
            ("x-provider-name", "String", "O", "Filter by signing provider name."),
            ("x-workflow-id", "String", "O", "Filter by workflow ID."),
            ("x-ca-name", "String", "O", "Filter by certificate authority name."),
        ] + HDR_DATE,
        request_rows=[
            ("citizen_id", "String", "O", "Filter by citizen ID number."),
            ("credential_id", "String", "O", "Filter by credential ID."),
            ("connector_name", "String", "O", "Filter by RSSP connector name."),
            ("certificateAuthorityName", "String", "O",
             "Filter by certificate authority name."),
        ],
        response_rows=[
            ("signing_transactions", "List<Object>", "M",
             "Array of signing transactions. Each item contains id, transaction_id, "
             "package_id, document_id, application_id, credential_id, request_id, "
             "field_name, user_email, connector_name, connector_name_remark, "
             "citizen_id_number, auth_mode, status (id, name, enabled, "
             "is_continue_processing), expiry, created_at, application_name, "
             "certificate_authority_name, credential_info and metadata. Empty fields "
             "are omitted from the response."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "signing_transactions": [', "    {",
                         '      "id": 9021,',
                         '      "transaction_id": "wf-a1b2c3",',
                         '      "package_id": 118,',
                         '      "document_id": 240,',
                         '      "credential_id": "c81d4e2e",',
                         '      "user_email": "user@example.com",',
                         '      "connector_name": "SMART_ID_RSS_VNeID",',
                         '      "citizen_id_number": "079200011188",',
                         '      "auth_mode": "OTP",',
                         '      "status": {',
                         '        "id": 2,',
                         '        "name": "SUCCESS",',
                         '        "enabled": true,',
                         '        "is_continue_processing": false',
                         "      },",
                         '      "expiry": "2026-07-27T04:10:00Z",',
                         '      "created_at": "2026-07-27T04:00:00Z",',
                         '      "application_name": "GoPaperless Portal",',
                         '      "certificate_authority_name": "MobileID CA",',
                         '      "credential_info": {},',
                         '      "metadata": {}',
                         "    }", "  ]", "}"],
    ),
    dict(
        name="Get Certificate Enroll History",
        donor=DONOR_LIST,
        anchor="Get Signing History",
        method="GET",
        url="{baseURL}/v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the certificate "
               "enrollment history of the logged in user. The optional status segment "
               "filters the enrollments by their step. Filter headers override the "
               "same fields in the request body when both are provided. Enrollments "
               "waiting for verification for more than 5 minutes are automatically "
               "expired. The x-total-records response header contains the total "
               "number of records."),
        path_rows=[
            ("status", "String", "O",
             "Optional status filter. Possible values are ALL, DONE, FAILED, EXPIRED, "
             "REJECTED, WAITING_FOR_APPROVAL_IN_VNEID, GENERATE_SESSION_SUCCESS, "
             "WAITING_EMAIL_VERIFY, WAITING_MOBILE_VERIFY and NOT_VERIFIED. Defaults "
             "to DONE when only the page segments are provided."),
        ] + PAGE_ROWS,
        header_add=[
            ("x-citizen-id", "String", "O", "Filter by citizen ID number."),
            ("x-connector-name", "String", "O", "Filter by RSSP connector name."),
            ("x-ca-name", "String", "O", "Filter by certificate authority name."),
        ] + HDR_DATE,
        body_lines=["{", '  "citizen_id": "079200011188",',
                    '  "connector_name": "SMART_ID_RSS_VNeID"', "}"],
        request_rows=[
            ("citizen_id", "String", "O", "Filter by citizen ID number."),
            ("connector_name", "String", "O", "Filter by RSSP connector name."),
            ("user_email", "String", "O", "Filter by user email."),
            ("certificate_profile_id", "Integer", "O", "Filter by certificate profile ID."),
            ("certificate_authority_vneid_id", "Integer", "O",
             "Filter by VNeID certificate authority ID."),
            ("application_id", "Integer", "O", "Filter by application ID."),
            ("attribute_name", "String", "O", "Filter by an enrollment attribute name."),
            ("attribute_value", "String", "O", "Filter by an enrollment attribute value."),
        ],
        response_rows=[
            ("certificate_enrollments", "List<Object>", "M",
             "Array of certificate enrollments. Each item contains session_id, "
             "process_id, user_email, credential_id, application_id, "
             "citizen_id_number, connector_name, connector_name_remark, "
             "certificate_authority_name, certificate_profile, "
             "certificate_authority_vneid_id, step, status, is_email_verified, "
             "is_mobile_verified, created_at, request_type and credential_info."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "certificate_enrollments": [', "    {",
                         '      "session_id": "f2f9d8a0",',
                         '      "process_id": "p-4471",',
                         '      "user_email": "user@example.com",',
                         '      "credential_id": "c81d4e2e",',
                         '      "application_id": 12,',
                         '      "citizen_id_number": "079200011188",',
                         '      "connector_name": "SMART_ID_RSS_VNeID",',
                         '      "certificate_authority_name": "MobileID CA",',
                         '      "certificate_profile": {},',
                         '      "step": "DONE",',
                         '      "status": {',
                         '        "name": "DONE"',
                         "      },",
                         '      "is_email_verified": true,',
                         '      "is_mobile_verified": true,',
                         '      "created_at": "2026-07-27T04:00:00Z",',
                         '      "request_type": "NEW",',
                         '      "credential_info": {}',
                         "    }", "  ]", "}"],
    ),
    dict(
        name="Get Dashboard Workflows",
        donor=DONOR_FLAT,
        anchor="Get User Activity Logs",
        method="GET",
        url="{baseURL}/v1/account/dashboard/workflows",
        prose=("Business applications can use this service API to get the account "
               "level workflow status counts for dashboard display. All counters are "
               "serialized as strings."),
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("all", "String", "M", "Total number of workflows."),
            ("completed", "String", "M", "Number of completed workflows."),
            ("inprogress", "String", "M", "Number of workflows in progress."),
            ("declined", "String", "M", "Number of declined workflows."),
            ("expired", "String", "M", "Number of expired workflows."),
            ("draft", "String", "M", "Number of draft workflows."),
            ("pending", "String", "M", "Number of pending workflows."),
            ("signed", "String", "M", "Number of signed workflows."),
            ("approved", "String", "M", "Number of approved workflows."),
            ("updated", "String", "M", "Number of updated workflows."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "all": "124",', '  "completed": "80",',
                         '  "inprogress": "20",', '  "declined": "10",',
                         '  "expired": "14",', '  "draft": "0",',
                         '  "pending": "0",', '  "signed": "0",',
                         '  "approved": "0",', '  "updated": "0"', "}"],
    ),
    dict(
        name="Get Dashboard Quota",
        donor=DONOR_FLAT,
        anchor="Get Dashboard Workflows",
        method="GET",
        url="{baseURL}/v1/account/dashboard/quota",
        prose=("Business applications can use this service API to get the account "
               "level quota consumption for dashboard display. Each quota entry "
               "contains the consumed amount, the currently allowed amount and the "
               "service plan allowance. All values are serialized as strings; "
               "unlimited quotas are returned as Unlimited."),
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("template", "Object", "M",
             "Template quota: consumed, allowed and serviceplan (allowed)."),
            ("signature_vneid", "Object", "M",
             "VNeID signature quota: consumed, allowed and serviceplan (allowed)."),
            ("signature", "Object", "M",
             "Signature quota: consumed, allowed and serviceplan (allowed)."),
            ("zalo_sms", "Object", "M",
             "Zalo/SMS notification quota: consumed, allowed and serviceplan (allowed)."),
            ("users", "Object", "M",
             "User quota: consumed, allowed and serviceplan (allowed)."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "template": {',
                         '    "consumed": "5",',
                         '    "allowed": "Unlimited",',
                         '    "serviceplan": {',
                         '      "allowed": "Unlimited"',
                         "    }", "  },",
                         '  "signature": {',
                         '    "consumed": "45",',
                         '    "allowed": "100",',
                         '    "serviceplan": {',
                         '      "allowed": "100"',
                         "    }", "  },",
                         '  "signature_vneid": {},',
                         '  "zalo_sms": {},',
                         '  "users": {}', "}"],
    ),
    dict(
        name="Get Dashboard Workflow Types",
        donor=DONOR_FLAT,
        anchor="Get Dashboard Quota",
        method="GET",
        url="{baseURL}/v1/account/dashboard/workflow-types",
        prose=("Business applications can use this service API to get the workflow "
               "type distribution of the account for dashboard display. One entry is "
               "returned for every workflow type (ECONTRACT, BASIC), zero-filled when "
               "the type has no workflows. All counters are serialized as strings."),
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("types", "List<Object>", "M",
             "Array of workflow type counters. Each item contains type (ECONTRACT or "
             "BASIC) and count."),
            ("total_count", "Object", "M", "Contains total, the sum across all types."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "types": [', "    {",
                         '      "type": "ECONTRACT",',
                         '      "count": "45"',
                         "    },", "    {",
                         '      "type": "BASIC",',
                         '      "count": "55"',
                         "    }", "  ],",
                         '  "total_count": {',
                         '    "total": "100"',
                         "  }", "}"],
    ),
    dict(
        name="Get Dashboard Categories",
        donor=DONOR_FLAT,
        anchor="Get Dashboard Workflow Types",
        method="GET",
        url="{baseURL}/v1/account/dashboard/categories",
        prose=("Business applications can use this service API to get the workflow "
               "category distribution of the account for dashboard display. All "
               "counters are serialized as strings."),
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("categories", "List<Object>", "M",
             "Array of category counters. Each item contains id, name and count."),
            ("total_count", "Object", "M",
             "Contains total, the sum across all categories."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "categories": [', "    {",
                         '      "id": "1",',
                         '      "name": "Contract",',
                         '      "count": "40"',
                         "    }", "  ],",
                         '  "total_count": {',
                         '    "total": "55"',
                         "  }", "}"],
    ),
    dict(
        name="Get Dashboard Validations",
        donor=DONOR_FLAT,
        anchor="Get Dashboard Categories",
        method="GET",
        url="{baseURL}/v1/account/dashboard/validations",
        prose=("Business applications can use this service API to get the document "
               "validation status counts of the account for dashboard display. Every "
               "status key is always present, zero-filled when it has no documents. "
               "All counters are serialized as strings."),
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("not_process", "String", "M", "Number of documents not yet processed."),
            ("no_signatures_or_seals_found", "String", "M",
             "Number of documents with no signatures or seals found."),
            ("warning", "String", "M", "Number of documents validated with warnings."),
            ("valid", "String", "M", "Number of valid documents."),
            ("invalid", "String", "M", "Number of invalid documents."),
            ("processing", "String", "M", "Number of documents being processed."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "not_process": "0",',
                         '  "no_signatures_or_seals_found": "9",',
                         '  "warning": "0",',
                         '  "valid": "100",',
                         '  "invalid": "5",',
                         '  "processing": "10"', "}"],
    ),
    dict(
        name="Get Dashboard Overdue Workflows",
        donor=DONOR_LIST,
        anchor="Get Dashboard Validations",
        method="GET",
        url="{baseURL}/v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the paged list "
               "of overdue workflows of the account for dashboard display. The "
               "x-total-records response header contains the total number of records."),
        path_rows=PAGE_ROWS,
        drop_tables=["Request Attributes"],
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("items", "List<Object>", "M",
             "Array of overdue workflows. Each item contains workflow_id (Number), "
             "name, expired_at and status."),
            ("total_count", "Object", "M",
             "Contains total, the total number of overdue workflows."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "items": [', "    {",
                         '      "workflow_id": 123,',
                         '      "name": "Contract A",',
                         '      "expired_at": "2026-06-18T00:00:00Z",',
                         '      "status": "EXPIRED"',
                         "    }", "  ],",
                         '  "total_count": {',
                         '    "total": "14"',
                         "  }", "}"],
    ),
    dict(
        name="Get Dashboard Activities",
        donor=DONOR_LIST,
        anchor="Get Dashboard Overdue Workflows",
        method="GET",
        url="{baseURL}/v1/account/dashboard/activities/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the paged "
               "activity feed of the account for dashboard display. The "
               "x-total-records response header contains the total number of records."),
        path_rows=PAGE_ROWS,
        drop_tables=["Request Attributes"],
        header_add=HDR_DATE + HDR_SEARCH,
        response_rows=[
            ("items", "List<Object>", "M",
             "Array of account activities. Each item contains id, "
             "performed_by_email, performed_by_name, action, module, created_on, "
             "information_type and information_value."),
            ("total_count", "Object", "M",
             "Contains total, the total number of activities."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "items": [', "    {",
                         '      "id": "1",',
                         '      "performed_by_email": "user@example.com",',
                         '      "performed_by_name": "Nguyen Van A",',
                         '      "action": "SIGNED",',
                         '      "module": "WORKFLOW",',
                         '      "created_on": "2026-06-19T10:00:00Z",',
                         '      "information_type": "WORKFLOW_ID",',
                         '      "information_value": "123"',
                         "    }", "  ],",
                         '  "total_count": {',
                         '    "total": "150"',
                         "  }", "}"],
    ),
    dict(
        name="Get Dashboard Processing Trend",
        donor=DONOR_FLAT,
        anchor="Get Dashboard Activities",
        method="GET",
        url="{baseURL}/v1/account/dashboard/processing-trend",
        prose=("Business applications can use this service API to get the workflow "
               "processing trend of the account over time for dashboard display. All "
               "counters are serialized as strings."),
        header_add=HDR_DATE,
        response_rows=[
            ("trend", "List<Object>", "M",
             "Array of trend points. Each item contains label and count."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "trend": [', "    {",
                         '      "label": "2026-07-01",',
                         '      "count": "12"',
                         "    }", "  ]", "}"],
    ),
    dict(
        name="Initialize Billing",
        donor=DONOR_POST,
        anchor="Add Identity for a User",
        method="POST",
        url="{baseURL}/v1/account/billings/initialize",
        prose=("Business applications can use this service API to initialize a "
               "Trusted Billing payment for the selected service plan of the logged "
               "in personal or enterprise account. Quantity, billing cycle, price, "
               "currency and payment type are derived on the server from the selected "
               "service plan billing entry. Payment URL, QR code, deeplink and expiry "
               "fields are null when the payment provider does not supply them."),
        body_lines=["{", '  "service_plan_id": 2012,',
                    '  "service_plan_billing_id": 32,',
                    '  "webhook_url": "https://example.com/webhook"', "}"],
        request_rows=[
            ("service_plan_id", "Integer", "M", "The ID of the service plan to purchase."),
            ("service_plan_billing_id", "Integer", "M",
             "The ID of the service plan billing entry that defines the billing "
             "cycle and price."),
            ("webhook_url", "String", "O",
             "Callback URL notified about payment status updates. Must start with "
             "http:// or https://. A default URL is applied when omitted."),
            ("discount_code", "String", "O", "Reserved for future use."),
            ("referral_code", "String", "O", "Reserved for future use."),
        ],
        response_rows=[
            ("billing_id", "String", "M", "The ID of the local billing account."),
            ("correlation_id", "String", "M",
             "Correlation ID of the payment, format BILLING-yyyyMMdd-XXXXXXXX."),
            ("transaction_id", "String", "M", "The ID of the local payment transaction."),
            ("status", "String", "M", "Status of the payment, e.g. IN_PROGRESS."),
            ("amount", "Number", "M", "Amount of the payment."),
            ("currency", "String", "M", "ISO currency code of the payment."),
            ("payment_url", "String", "O", "Payment URL of the provider."),
            ("qr_code", "String", "O", "QR code of the payment."),
            ("va_number", "String", "O", "Virtual account number of the payment."),
            ("deeplink", "String", "O", "Deeplink of the payment."),
            ("expires_at", "String", "O", "Expiry time of the payment."),
            ("core_processing_available", "Boolean", "M",
             "True when the payment provider processed the request."),
            ("message_key", "String", "O", "Additional message key of the provider."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "billing_id": "1201",',
                         '  "correlation_id": "BILLING-20260727-8F3K2A1B",',
                         '  "transaction_id": "9110",',
                         '  "status": "IN_PROGRESS",',
                         '  "amount": 1200000,',
                         '  "currency": "VND",',
                         '  "payment_url": "https://pay.example.com/bill/9110",',
                         '  "qr_code": null,',
                         '  "va_number": null,',
                         '  "deeplink": null,',
                         '  "expires_at": "2026-07-28T04:00:00Z",',
                         '  "core_processing_available": true,',
                         '  "message_key": null', "}"],
    ),
    dict(
        name="Get Billing Payment History",
        donor=DONOR_LIST,
        anchor="Initialize Billing",
        method="GET",
        url="{baseURL}/v1/account/billings/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the payment "
               "level billing history of the logged in account. Optional query "
               "parameters filter the result. The x-total-records response header "
               "contains the total number of records."),
        path_rows=PAGE_ROWS,
        request_rows=[
            ("status", "String", "O",
             "Query parameter. Filter by payment status: PENDING, COMPLETED, "
             "PROVISIONING_PENDING, FAILED, CANCELLED or EXPIRED."),
            ("type", "String", "O",
             "Query parameter. Filter by payment type: FIRST_PAYMENT or RECURRENT."),
            ("connector_name", "String", "O",
             "Query parameter. Filter by exact connector name, e.g. TRUSTED_BILLING."),
            ("payment_number", "String", "O",
             "Query parameter. Filter by the provider bill or payment ID."),
        ],
        response_rows=[
            ("items", "List<Object>", "M",
             "Array of billing payments. Each item contains billing_payment_id, "
             "billing_id, transaction_id, currency_id, type, status, connector_name, "
             "payment_number, vat_number, amount (Number), remark, expiry_date, "
             "payment_date and created_at."),
            ("pagination", "Object", "M",
             "Contains page_no, records_per_page and total_records."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "items": [', "    {",
                         '      "billing_payment_id": 9110,',
                         '      "billing_id": 1201,',
                         '      "transaction_id": "9110",',
                         '      "currency_id": 1,',
                         '      "type": "FIRST_PAYMENT",',
                         '      "status": "COMPLETED",',
                         '      "connector_name": "TRUSTED_BILLING",',
                         '      "payment_number": "B-2026-000451",',
                         '      "amount": 1200000,',
                         '      "expiry_date": "2026-07-28T04:00:00Z",',
                         '      "payment_date": "2026-07-27T05:00:00Z",',
                         '      "created_at": "2026-07-27T04:00:00Z"',
                         "    }", "  ],",
                         '  "pagination": {',
                         '    "page_no": 1,',
                         '    "records_per_page": 10,',
                         '    "total_records": 1',
                         "  }", "}"],
    ),
    dict(
        name="Get Billing Payment Details",
        donor=DONOR_LIST,
        anchor="Get Billing Payment History",
        method="GET",
        url="{baseURL}/v1/account/billings/payments/{billingPaymentId}",
        prose=("Business applications can use this service API to get the details of "
               "a billing payment owned by the logged in account, enriched with the "
               "Trusted Billing bill details. Fields with a null value are omitted "
               "from the response."),
        path_rows=[("billingPaymentId", "Integer", "M",
                    "The ID of the billing payment to retrieve.")],
        drop_tables=["Request Attributes"],
        response_rows=[
            ("billing", "Object", "M",
             "Local billing account: billing_id, external_customer_id, "
             "reference_number, address, zipcode, mobile_number, fax, discount, "
             "description, duration, billing_date, created_at, modified_at, "
             "created_by and modified_by."),
            ("billing_payment", "Object", "M",
             "Local billing payment: billing_payment_id, billing_id, transaction_id, "
             "currency_id, type, status, connector_name, payment_number, vat_number, "
             "amount (Number), remark, expiry_date, payment_date, created_at, "
             "created_by and modified_by."),
            ("trusted_billing", "Object", "O",
             "Trusted Billing bill details: bill_id, bill_number, status, "
             "company_id, customer_id, customer_name, qr_code, va_number, "
             "item_properties (item_id, name, uom, quantity, unit_price, amount, "
             "tax, tax_value, amount_after_discount), created_at and updated_at. "
             "Omitted when the provider returns no details."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "{",
                         '  "billing": {',
                         '    "billing_id": 1201,',
                         '    "external_customer_id": "CUS-000118",',
                         '    "reference_number": "REF-2026-1201",',
                         '    "billing_date": "2026-07-27T04:00:00Z",',
                         '    "created_at": "2026-07-27T04:00:00Z"',
                         "  },",
                         '  "billing_payment": {',
                         '    "billing_payment_id": 9110,',
                         '    "billing_id": 1201,',
                         '    "status": "COMPLETED",',
                         '    "amount": 1200000,',
                         '    "created_at": "2026-07-27T04:00:00Z"',
                         "  },",
                         '  "trusted_billing": {',
                         '    "bill_id": "TB-000451",',
                         '    "bill_number": "B-2026-000451",',
                         '    "status": "PAID",',
                         '    "item_properties": [', "      {",
                         '        "item_id": "SP-2012",',
                         '        "name": "Enterprise Plan",',
                         '        "quantity": 1,',
                         '        "unit_price": 1200000,',
                         '        "amount": 1200000',
                         "      }", "    ]", "  }", "}"],
    ),
    dict(
        name="Get Account Quota Consume History",
        donor=DONOR_LIST,
        anchor="Get Billing Payment Details",
        method="GET",
        url="{baseURL}/v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}",
        prose=("Business applications can use this service API to get the quota "
               "consumption history of the logged in account. All numeric values are "
               "serialized as strings. The x-total-records response header contains "
               "the total number of records. The x-sort-by, x-asc, x-search-text and "
               "x-quota-dimension headers are not supported by this API and are "
               "rejected when present."),
        path_rows=PAGE_ROWS,
        drop_tables=["Request Attributes"],
        header_add=HDR_DATE + [
            ("x-billing-id", "String", "O", "Filter by billing account ID."),
            ("x-type", "String", "O",
             "Filter by quota type: SIGNATURES, WORKFLOWS, TEMPLATES, STORAGE, "
             "UPLOAD_SIZE, USERS, VNEID, ZALO_MESSAGES, TSA, ECONTRACT, EDELIVERY, "
             "VNEID_SHARE_INFO, ROLLOVER_OUT or ROLLOVER_IN."),
        ],
        response_rows=[
            ("id", "String", "O", "The ID of the quota log entry."),
            ("created_at", "String", "M", "Creation time of the entry, UTC ISO 8601."),
            ("type", "String", "M", "Quota type of the entry."),
            ("type_remark", "String", "O", "Localized label of the quota type."),
            ("source_system_type", "String", "O", "Source system of the consumption."),
            ("source_system_type_remark", "String", "O",
             "Localized label of the source system."),
            ("quota_dimension", "String", "M", "Quota dimension of the entry."),
            ("used_balance", "String", "M", "Consumed amount of the entry."),
            ("remaining_balance", "String", "M",
             "Negative used_balance for ROLLOVER_OUT entries, otherwise 0."),
            ("added_balance", "String", "O",
             "Equals used_balance for ROLLOVER_IN entries, otherwise null."),
            ("quantity", "String", "M", "Consumed quantity of the entry."),
            ("workflow_name", "String", "O", "Name of the related workflow."),
            ("sync_status", "String", "O", "Synchronization status of the entry."),
            ("sync_status_remark", "String", "O",
             "Localized label of the synchronization status."),
        ],
        response_sample=["Http Code: 200", "Content-Type: application/json", "[", "  {",
                         '    "id": "5012",',
                         '    "created_at": "2026-07-27T04:00:00Z",',
                         '    "type": "SIGNATURES",',
                         '    "type_remark": "Signatures",',
                         '    "quota_dimension": "SIGNATURES",',
                         '    "used_balance": "1",',
                         '    "remaining_balance": "0",',
                         '    "quantity": "1",',
                         '    "workflow_name": "Contract A",',
                         '    "sync_status": "SUCCESS"', "  }", "]"],
    ),
]


# ---------------------------------------------------------------------------
# helpers (same approach as add_document_workflow_missing_sections.py)
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


def set_text_keep_bookmarks(p, text):
    """Heading rename that must not delete the bookmarks anchored on it."""
    runs = [r for r in p.findall(qn("r")) if r.find(qn("t")) is not None]
    if not runs:
        raise ValueError("paragraph has no direct text run: %r" % text_of(p))
    for extra in runs[1:]:
        p.remove(extra)
    keep = runs[0]
    for extra in keep.findall(qn("t"))[1:]:
        keep.remove(extra)
    t = keep.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")


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


# ---------------------------------------------------------------------------
# frame editing - preserves the multi-run first line (template 2.8.1)
# ---------------------------------------------------------------------------
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


def first_method_line(txbx):
    paras = txbx.findall(qn("p"))
    for p in paras:
        txt = text_of(p).strip()
        if txt:
            return p, txt
    return None, ""


def edit_request_frame(txbx, method, url, body_lines):
    """Rewrite the URL line (wherever it is - some frames start with a blank
    line) and replace the JSON body."""
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


# ---------------------------------------------------------------------------
# table editing
# ---------------------------------------------------------------------------
def resize_fill(tbl, rows):
    """Replace every data row with `rows` [(name, type, presence, desc)]."""
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
    """Keep the standard error triple (rows 1-3), replace the payload rows."""
    trs = tbl.findall(qn("tr"))
    if len(trs) < 5:
        raise RuntimeError("response table too small: %s" % table_title(tbl))
    template = deepcopy(trs[-1])
    for tr in trs[5:]:
        tbl.remove(tr)
    anchor = tbl.findall(qn("tr"))[-1]
    if cell_text(anchor.findall(qn("tc"))[1]) not in ("message",):
        # trs[2:5] must be the error triple; verify before appending
        names = [cell_text(tr.findall(qn("tc"))[1]) for tr in tbl.findall(qn("tr"))[2:5]]
        if names != ["error_description", "error_code", "message"]:
            raise RuntimeError("unexpected leading response rows: %s" % names)
    for name, typ, presence, desc in rows:
        tr = deepcopy(template)
        tcs = tr.findall(qn("tc"))
        for tc, val in zip(tcs, ["", name, typ, presence, desc]):
            set_cell_text(tc, val)
        anchor.addnext(tr)
        anchor = tr
    # renumber all data rows
    for i, tr in enumerate(tbl.findall(qn("tr"))[2:]):
        set_cell_text(tr.findall(qn("tc"))[0], str(i + 1))


def append_rows(tbl, rows):
    """Append [(name, type, presence, desc)] after the existing data rows."""
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


# ---------------------------------------------------------------------------
# clone assembly
# ---------------------------------------------------------------------------
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

    # drop tables not applicable to this endpoint (plus the blank after each)
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
def section_slice(kids, title):
    start, end = section_range(kids, find_heading(kids, 3, title))
    return start, end


def frames_of(kids, start, end):
    out = []
    for i in range(start, end):
        for txbx in kids[i].iter(qn("txbxContent")):
            out.append(txbx)
    return out


def tables_of(kids, start, end):
    return [(i, kids[i]) for i in range(start, end) if kids[i].tag == qn("tbl")]


def fix_activity_logs(kids):
    start, end = section_slice(kids, "Get User Activity Logs")
    n = 0
    for txbx in frames_of(kids, start, end):
        first = first_line_of_first_frame_txbx(txbx)
        if first and first.startswith("GET "):
            edit_request_frame(
                txbx, "GET",
                "{baseURL}/v1/account/activity-logs/{pageNo}/{recordsPerPage}", None)
            n += 1
    if n == 0:
        raise RuntimeError("activity logs request frame not found")
    print("FIXED Get User Activity Logs path (frames=%d)" % n)


def first_line_of_first_frame_txbx(txbx):
    for p in txbx.findall(qn("p")):
        txt = text_of(p).strip()
        if txt:
            return txt
    return None


def fix_get_invitations(kids, body, path_tbl_donor, blank_donor):
    start, end = section_slice(kids, "Get Account Invitations")
    n = 0
    for txbx in frames_of(kids, start, end):
        first = first_line_of_first_frame_txbx(txbx)
        if first and first.startswith("GET "):
            url_p = None
            for p in txbx.findall(qn("p")):
                if text_of(p).strip().startswith("GET "):
                    url_p = p
                    break
            rewrite_url_line(
                url_p, "GET",
                "{baseURL}/v1/account/invitations/{pageNo}/{recordsPerPage}")
            n += 1
    if n == 0:
        raise RuntimeError("get invitations request frame not found")
    # insert a Path Attributes table (donor: activity logs) after the header table
    hdr_idx = None
    for i in range(start, end):
        if kids[i].tag == qn("tbl") and table_title(kids[i]) == "Header Request Attributes":
            hdr_idx = i
            break
    if hdr_idx is None:
        raise RuntimeError("get invitations header table not found")
    new_tbl = deepcopy(path_tbl_donor)
    strip_bookmarks([new_tbl])
    resize_fill(new_tbl, [
        ("pageNo", "Integer", "O",
         "Page number to be retrieved (1-based). The pagination segments are "
         "optional; all invitations are returned when omitted."),
        ("recordsPerPage", "Integer", "O",
         "Number of records to be retrieved in one page."),
    ])
    blank = deepcopy(blank_donor)
    strip_bookmarks([blank])
    after = kids[hdr_idx + 1] if hdr_idx + 1 < len(kids) else kids[hdr_idx]
    after.addnext(new_tbl)
    new_tbl.addnext(blank)
    print("FIXED Get Account Invitations pagination path (frames=%d, +Path table)" % n)


def fix_invitation_decision(kids, body, title, new_title, url, prose, req_tbl_donor,
                            blank_donor, need_req_tbl):
    start, end = section_slice(kids, title)
    set_text_keep_bookmarks(kids[start], new_title)
    set_para_text(kids[start + 1], prose)
    n = 0
    for txbx in frames_of(kids, start, end):
        first = first_line_of_first_frame_txbx(txbx)
        if first is None:
            continue
        if any(first.startswith(m + " ") for m in METHODS):
            edit_request_frame(txbx, "POST", url,
                               ["{", '  "invitation_token": "8f3k2a1b-...-9d1c"', "}"])
            n += 1
        elif first.startswith("Http Code: 200"):
            replace_frame_lines(txbx, ["Http Code: 204", "Content-Type: application/json",
                                       "{}"])
    if n == 0:
        raise RuntimeError("%s request frame not found" % title)

    token_row = [("invitation_token", "String", "M",
                  "Token of the invitation, as delivered in the invitation email link.")]
    has_req = False
    for i, tbl in tables_of(kids, start, end):
        if table_title(tbl) == "Request Attributes":
            resize_fill(tbl, token_row)
            has_req = True
    if not has_req and need_req_tbl:
        hdr_idx = None
        s2, e2 = section_slice(list(body), new_title)
        kids2 = list(body)
        for i in range(s2, e2):
            if kids2[i].tag == qn("tbl") and \
                    table_title(kids2[i]) == "Header Request Attributes":
                hdr_idx = i
                break
        if hdr_idx is None:
            raise RuntimeError("%s header table not found" % new_title)
        new_tbl = deepcopy(req_tbl_donor)
        strip_bookmarks([new_tbl])
        resize_fill(new_tbl, token_row)
        blank = deepcopy(blank_donor)
        strip_bookmarks([blank])
        after = kids2[hdr_idx + 1]
        after.addnext(new_tbl)
        new_tbl.addnext(blank)
    print("FIXED %s -> %s (frames=%d, req_tbl=%s)"
          % (title, new_title, n, "kept" if has_req else "added"))


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

    # --- donor artifacts for table insertion ---------------------------------
    kids = list(body)
    al_start, al_end = section_slice(kids, "Get User Activity Logs")
    path_tbl_donor = next(kids[i] for i in range(al_start, al_end)
                          if kids[i].tag == qn("tbl")
                          and table_title(kids[i]) == "Path Attributes")
    req_tbl_donor = next(kids[i] for i in range(al_start, al_end)
                         if kids[i].tag == qn("tbl")
                         and table_title(kids[i]) == "Request Attributes")
    blank_donor = next(kids[i] for i in range(al_start, al_end)
                       if kids[i].tag == qn("p") and not text_of(kids[i]).strip())

    # --- 1. fix existing sections (before insertion so anchors carry new names)
    fix_activity_logs(kids)
    fix_get_invitations(kids, body, path_tbl_donor, blank_donor)
    kids = list(body)
    fix_invitation_decision(
        kids, body, "Accept Account Invitations", "Accept Account Invitation",
        "{baseURL}/v1/account/invitations/accept",
        ("Business applications can use this service API to accept one enterprise "
         "invitation of the logged in user. The invitation is identified by its "
         "invitation token. Unlike SigningHub, other pending invitations are not "
         "affected. A successful call returns HTTP 204 with no content."),
        req_tbl_donor, blank_donor, need_req_tbl=False)
    kids = list(body)
    fix_invitation_decision(
        kids, body, "Reject All Account Invitations", "Reject Account Invitation",
        "{baseURL}/v1/account/invitations/reject",
        ("Business applications can use this service API to reject one pending "
         "enterprise invitation of the logged in user. The invitation is identified "
         "by its invitation token. Unlike SigningHub, only the referenced invitation "
         "is rejected. A successful call returns HTTP 204 with no content."),
        req_tbl_donor, blank_donor, need_req_tbl=True)

    # --- 2. add the 17 new sections -----------------------------------------
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
        print("ADDED %-38s donor=%-26s after=%-32s els=%d frames=%d bm=%d docPr=%d vml=%d"
              % (spec["name"][:38], spec["donor"][:26], spec["anchor"][:32],
                 len(clone), n_frames, n_bm, n_doc, n_vml))

    if dry:
        print("\n--dry-run: nothing written")
        return
    save(tree, DOC_XML, token)
    backup = OUT_DOCX.replace(".docx", "_backup_pre_account_management.docx")
    if os.path.exists(OUT_DOCX) and not os.path.exists(backup):
        shutil.copy2(OUT_DOCX, backup)
        print("backup %s" % backup)
    pack(UNPACKED, OUT_DOCX)
    print("packed %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
