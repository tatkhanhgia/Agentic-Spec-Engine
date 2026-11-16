# OpenAPI vs DOCX contract comparison

| Metric | Value |
|---|---|
| Spec records (operations expanded by x-doc-sections) | 329 |
| Doc endpoints | 353 |
| Paired | 320 |
| Doc only (no spec operation) | 33 |
| Spec only (not in doc) | 13 |
| Paired with no field difference | 313 |
| category differs | 1 |
| headers: common | 1108 |
| headers: only in doc | 7 |
| headers: only in spec | 10 |
| headers: rows in doc | 1115 |
| headers: rows in spec | 1118 |
| name differs | 6 |
| path_params: common | 335 |
| path_params: only in doc | 2 |
| path_params: only in spec | 6 |
| path_params: rows in doc | 337 |
| path_params: rows in spec | 341 |
| request_fields: common | 1576 |
| request_fields: only in doc | 0 |
| request_fields: only in spec | 0 |
| request_fields: rows in doc | 1576 |
| request_fields: rows in spec | 1576 |
| request_fields: type differs | 77 |
| response_fields: common | 710 |
| response_fields: only in doc | 5 |
| response_fields: only in spec | 1 |
| response_fields: rows in doc | 715 |
| response_fields: rows in spec | 711 |
| response_fields: type differs | 62 |

## Doc only

- `POST /v1/authenticate/otp` — Authenticate – Log in / OTP Login Authentication (Not implemented yet)
- `GET /v1/packages/{packageId}/verification` — Document Package / Get Package Verification (Not implemented yet)
- `POST /v1/packages/{packageId}/workflow/electronic-seals` — Document Workflow / Add Electronic Seal to Workflow (Not implemented yet)
- `PUT /v1/packages/{packageId}/workflow/reminders` — Document Workflow / Update Workflow Reminders (Not implemented yet)
- `GET /v1/packages/{packageId}/workflow/{order}/authentication` — Document Workflow / Get Workflow User Authentication (Document Opening) (Not implemented yet)
- `PUT /v1/packages/{packageId}/workflow/{order}/authentication` — Document Workflow / Update Workflow User Authentication (Document Opening) (Not implemented yet)
- `GET /v1/packages/{packageId}/workflow/{order}/permissions` — Document Workflow / Get Workflow User Permissions (Not implemented yet)
- `PUT /v1/packages/{packageId}/workflow/{order}/permissions` — Document Workflow / Update Workflow User Permissions (Not implemented yet)
- `GET /v1/enterprise/packages/{packageId}/workflow/{order}/permissions` — Enterprise Management / Get Workflow User Permissions of Enterprise Package (Not implemented yet)
- `PUT /v1/enterprise/packages/{packageId}/workflow/{order}/permissions` — Enterprise Management / Update Workflow User Permissions of Enterprise Package (Not implemented yet)
- `GET /v1/enterprise/packages/{packageId}/workflow/{order}/authentication` — Enterprise Management / Get Workflow User Authentication (Document Opening) of Enterprise Package (Not implemented yet)
- `PUT /v1/enterprise/packages/{packageId}/workflow/{order}/authentication` — Enterprise Management / Update Workflow User Authentication of Enterprise Package (Not implemented yet)
- `POST /v1/enterprise/users/{id}/reset-password` — Enterprise Management / Reset Enterprise User Password (Not implemented yet)
- `GET /v1/enterprise/users/{id}` — Enterprise Management / Get Enterprise User (Not implemented yet)
- `PUT /v1/enterprise/users` — Enterprise Management / Update Enterprise User (Not implemented yet)
- `GET /v1/settings/groups/{pageNo}/{recordPerPage}` — Enterprise Management / Get Groups (Not implemented yet)
- `DELETE /v1/settings/groups/{id}` — Enterprise Management / Delete Groups (Not implemented yet)
- `POST /v1/enterprise/groups/{id}` — Enterprise Management / Add Member into Group (Not implemented yet)
- `DELETE /v1/enterprise/groups/{id}` — Enterprise Management / Delete Member in Group (Not implemented yet)
- `GET /v1/enterprise/roles/{id}/users/{pageNo}/{recordPerPage}` — Enterprise Management / Get Role Users (Not implemented yet)
- `POST /v1/account` — Account Management / Register Enterprise User (Free Trial) (Not implemented yet)
- `GET /v1/account` — Account Management / Get Account (Not implemented yet)
- `POST /v1/account/password/reset` — Account Management / Send Forgot Password Request (Not implemented yet)
- `GET /v1/account/dashboard/processing-trend` — Account Management / Get Dashboard Processing Trend (Not implemented yet)
- `GET /v1/settings/groups/{pageNo}/{recordPerPage}` — Account Management / Get Group (Not implemented yet)
- `PUT /v1/settings/profile/password` — Account Management / Change Password (Not implemented yet)
- `PUT /v1/settings/profile/security` — Account Management / Update Security Settings (Not implemented yet)
- `PUT /v1/settings/addressbook` — Personal Settings / Add Contact into Address Book (Not implemented yet)
- `PUT /v1/settings/addressbook` — Personal Settings / Remove Contact in Address Book (Not implemented yet)
- `POST /v1/admin/authenticate/client_cert` — ADMIN APIs / Admin Authentication (Not implemented yet)
- `PUT /v1/admin/accounts` — ADMIN APIs / Change User Service Plan (Not implemented yet)
- `GET /v1/admin/accounts/{pageNo}/{recordsPerPage}` — ADMIN APIs / Get Accounts (Not implemented yet)
- `POST /v1/admin/accounts/statisics/reset` — ADMIN APIs / Service Plan Reset (Not implemented yet)

## Spec only

- `GET /v1/enterprise/invitations/export/{status}` — Enterprise Management / Export Enterprise Invitations
- `GET /v1/enterprise/invitations/export` — Enterprise Management / Export Enterprise Invitations
- `GET /v1/enterprise/reports/workflows/statistics/export/{status}` — Enterprise Management / Export Enterprise Workflow Statistics
- `GET /v1/enterprise/reports/workflows/statistics/export` — Enterprise Management / Export Enterprise Workflow Statistics
- `GET /v1/account/reports/quota/consume-history/export/{pageNo}/{recordsPerPage}` — Account Management / Export Account Quota Consume History
- `GET /v1/account/reports/quota/consume-history/export` — Account Management / Export Account Quota Consume History
- `POST /v1/identity/verification/email-otp` — Account Management / Send Email OTP
- `GET /v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}` — Template / Get List Template
- `GET /v1/terms` — Account Management / Get List Terms of Service
- `GET /v1/system/terms-of-service/{termsOfServiceUuid}` — Account Management / Download Terms of Service File
- `GET /v1/system/serviceplans/{pageNo}/{recordsPerPage}` — Enterprise Management / Get All Service Plans
- `DELETE /v1/application/queue` — Application/Device Management / Add API Request Queue
- `POST /v1/authenticate/tse/status` — Authenticate – Log in / Check TSE Status

## Paired endpoints with differences

### Login TSE (Not implemented yet) — `POST /v1/authenticate/tse`
- name: doc `Login TSE (Not implemented yet)` / spec `Login TSE`

### Check TSE Status (Not implemented yet) — `GET /v1/authenticate/tse/status`
- name: doc `Check TSE Status (Not implemented yet)` / spec `Check TSE Status`

### Share Document Package/ Share Workflow — `POST /v1/packages/{packageId}/workflow/share`
- name: doc `Share Document Package/ Share Workflow` / spec `Share Workflow`
- category: doc `Document Package` / spec `Document Workflow`
- headers only in doc: Content-Type

### Get Workflow Reminders (Not implemented yet) — `GET /v1/packages/{packageId}/workflow/reminders`
- name: doc `Get Workflow Reminders (Not implemented yet)` / spec `Get All Workflow`
- headers only in doc: Content-Type
- headers only in spec: x-category-id, x-category-mode, x-enterprise-mode, x-filter-field-type, x-from-date, x-search-text, x-to-date, x-total-records, x-workflow-type
- path_params only in doc: packageId
- path_params only in spec: criteria, pageNo, recordsPerPage
- response_fields only in doc: enabled, frequency, keep_reminding_after, remind_after, total_reminders
- response_fields only in spec: workflows

### Update Enterprise Electronic Seals — `PUT /v1/enterprise/electronic-seals/{sealId}`
- path_params only in spec: sealId

### Export Account Quota Consume History — `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`
- name: doc `Export Account Quota Consume History` / spec `Get Account Invitations`
- headers only in doc: x-billing-id, x-export-format, x-from-date, x-to-date, x-type

### Download Terms of Service File — `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`
- name: doc `Download Terms of Service File` / spec `Get Account Invitations`
- headers only in spec: x-total-records
- path_params only in doc: termsOfServiceUuid
- path_params only in spec: pageNo, recordsPerPage

