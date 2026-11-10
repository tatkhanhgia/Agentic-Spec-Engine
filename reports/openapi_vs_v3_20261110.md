# OpenAPI vs DOCX contract comparison

| Metric | Value |
|---|---|
| Spec records (operations expanded by x-doc-sections) | 329 |
| Doc endpoints | 265 |
| Paired | 229 |
| Doc only (no spec operation) | 36 |
| Spec only (not in doc) | 102 |
| Paired with no field difference | 28 |
| category differs | 1 |
| headers: common | 719 |
| headers: only in doc | 129 |
| headers: only in spec | 93 |
| headers: presence differs | 3 |
| headers: rows in doc | 848 |
| headers: rows in spec | 812 |
| headers: type differs | 9 |
| name differs | 2 |
| path_params: common | 254 |
| path_params: only in doc | 34 |
| path_params: only in spec | 36 |
| path_params: presence differs | 16 |
| path_params: rows in doc | 288 |
| path_params: rows in spec | 290 |
| path_params: type differs | 7 |
| request_fields: common | 559 |
| request_fields: doc presence is conditional (RC) | 15 |
| request_fields: only in doc | 502 |
| request_fields: only in spec | 285 |
| request_fields: presence differs | 182 |
| request_fields: rows in doc | 1061 |
| request_fields: rows in spec | 844 |
| request_fields: type differs | 25 |
| response_fields: common | 322 |
| response_fields: doc presence is conditional (RC) | 2 |
| response_fields: only in doc | 161 |
| response_fields: only in spec | 124 |
| response_fields: presence differs | 259 |
| response_fields: rows in doc | 483 |
| response_fields: rows in spec | 446 |
| response_fields: type differs | 94 |

## Doc only

- `POST /v1/authenticate/otp` — Authenticate – Log in / OTP Login Authentication
- `GET /v1/packages/{packageId}/verification` — Document Package / Get Package Verification
- `POST /v1/packages/{packageId}/workflow/electronic-seals` — Document Workflow / Add Electronic Seal to Workflow
- `PUT /v1/packages/{packageId}/workflow/reminders` — Document Workflow / Update Workflow Reminders
- `GET /v1/packages/{packageId}/workflow/{order}/authentication` — Document Workflow / Get Workflow User Authentication (Document Opening)
- `PUT /v1/packages/{packageId}/workflow/{order}/authentication` — Document Workflow / Update Workflow User Authentication (Document Opening)
- `GET /v1/packages/{packageId}/workflow/{order}/permissions` — Document Workflow / Get Workflow User Permissions
- `PUT /v1/packages/{packageId}/workflow/{order}/permissions` — Document Workflow / Update Workflow User Permissions
- `GET /v1/connectors/certificate-profiles` — Connector / Get Certificate Profile of Connector
- `GET /v1/enterprise/packages/{packageId}/workflow/{order}/permissions` — Enterprise Management / Get Workflow User Permissions of Enterprise Package
- `PUT /v1/enterprise/packages/{packageId}/workflow/{order}/permissions` — Enterprise Management / Update Workflow User Permissions of Enterprise Package
- `GET /v1/enterprise/packages/{packageId}/workflow/{order}/authentication` — Enterprise Management / Get Workflow User Authentication (Document Opening) of Enterprise Package
- `PUT /v1/enterprise/packages/{packageId}/workflow/{order}/authentication` — Enterprise Management / Update Workflow User Authentication of Enterprise Package
- `POST /v1/enterprise/users/{id}/reset-password` — Enterprise Management / Reset Enterprise User Password
- `GET /v1/enterprise/users/{id}` — Enterprise Management / Get Enterprise User
- `PUT /v1/enterprise/users` — Enterprise Management / Update Enterprise User
- `GET /v1/settings/groups/{pageNo}/{recordPerPage}` — Enterprise Management / Get Groups
- `DELETE /v1/settings/groups/{id}` — Enterprise Management / Delete Groups
- `POST /v1/enterprise/groups/{id}` — Enterprise Management / Add Member into Group
- `DELETE /v1/enterprise/groups/{id}` — Enterprise Management / Delete Member in Group
- `GET /v1/enterprise/roles/{id}/users/{pageNo}/{recordPerPage}` — Enterprise Management / Get Role Users
- `PUT /v1/enterprise/electronic-seals` — Enterprise Management / Update Enterprise Electronic Seals
- `POST /v1/account` — Account Management / Register Enterprise User (Free Trial)
- `GET /v1/account` — Account Management / Get Account
- `POST /v1/account/password/reset` — Account Management / Send Forgot Password Request
- `GET /v1/account/dashboard/processing-trend` — Account Management / Get Dashboard Processing Trend
- `GET /v1/settings/groups/{pageNo}/{recordPerPage}` — Account Management / Get Group
- `PUT /v1/settings/profile/password` — Account Management / Change Password
- `PUT /v1/settings/profile/security` — Account Management / Update Security Settings
- `PUT /v1/settings/addressbook/{id}` — Personal Settings / Update Address Book
- `PUT /v1/settings/addressbook` — Personal Settings / Add Contact into Address Book
- `PUT /v1/settings/addressbook` — Personal Settings / Remove Contact in Address Book
- `POST /v1/admin/authenticate/client_cert` — ADMIN APIs / Admin Authentication
- `PUT /v1/admin/accounts` — ADMIN APIs / Change User Service Plan
- `GET /v1/admin/accounts/{pageNo}/{recordsPerPage}` — ADMIN APIs / Get Accounts
- `POST /v1/admin/accounts/statisics/reset` — ADMIN APIs / Service Plan Reset

## Spec only

- `GET /v1/enterprise/electronic-seals/{pageNo}/{recordsPerPage}` — Enterprise Management / List Enterprise Electronic Seals
- `GET /v1/enterprise/electronic-seals/{sealId}` — Enterprise Management / Get Enterprise Electronic Seal By Id
- `PUT /v1/enterprise/electronic-seals/{sealId}` — Enterprise Management / Update Enterprise Electronic Seal By Id
- `DELETE /v1/enterprise/electronic-seals/{sealId}` — Enterprise Management / Delete Enterprise Electronic Seal
- `POST /v1/enterprise/electronic-seals` — Enterprise Management / Add Enterprise Electronic Seal
- `POST /v1/enterprise/electronic-seals/verify-pin` — Enterprise Management / Verify Enterprise Electronic Seal PIN (draft)
- `POST /v1/enterprise/electronic-seals/{sealId}/verify-pin` — Enterprise Management / Verify Saved Enterprise Electronic Seal PIN
- `GET /v1/enterprise/settings/download-link` — Enterprise Management / Get Enterprise Download Link Setting
- `PUT /v1/enterprise/settings/download-link` — Enterprise Management / Update Enterprise Download Link Setting
- `PUT /v1/enterprise/profile` — Enterprise Management / Update Enterprise Profile
- `GET /v1/enterprise/invitations/export/{pageNo}/{recordsPerPage}` — Enterprise Management / Export Enterprise Invitations
- `GET /v1/enterprise/invitations/export/{status}` — Enterprise Management / Export Enterprise Invitations
- `GET /v1/enterprise/invitations/export` — Enterprise Management / Export Enterprise Invitations
- `GET /v1/enterprise/users/export` — Enterprise Management / Export Enterprise Users
- `PUT /v1/enterprise/users/{enterpriseUserId}` — Enterprise Management / Update Enterprise User Role
- `GET /v1/enterprise/permissions` — Enterprise Management / Get Permission Catalog
- `GET /v1/enterprise/roles/{enterpriseRoleId}/permissions` — Enterprise Management / Get Role Permissions
- `PUT /v1/enterprise/roles/{enterpriseRoleId}/permissions` — Enterprise Management / Update Role Permissions
- `PUT /v1/enterprise/roles/{enterpriseRoleId}/basic` — Enterprise Management / Update Enterprise Role Basic
- `GET /v1/enterprise/reports/workflows/statistics/export/{pageNo}/{recordsPerPage}` — Enterprise Management / Export Enterprise Workflow Statistics
- `GET /v1/enterprise/reports/workflows/statistics/export/{status}` — Enterprise Management / Export Enterprise Workflow Statistics
- `GET /v1/enterprise/reports/workflows/statistics/export` — Enterprise Management / Export Enterprise Workflow Statistics
- `GET /v1/settings/templates/{pageNo}/{recordsPerPage}` — Template / Get List Template
- `POST /v1/verification/email-otp` — Account Management / Send Email OTP
- `GET /v1/account/dashboard/documents` — Account Management / Get Dashboard Documents
- `GET /v1/account/dashboard/recent-documents/{pageNo}/{recordsPerPage}` — Account Management / Get Dashboard Recent Documents
- `GET /v1/account/reports/quota/consume-history/export/{pageNo}/{recordsPerPage}` — Account Management / Export Account Quota Consume History
- `GET /v1/account/reports/quota/consume-history/export` — Account Management / Export Account Quota Consume History
- `POST /v1/packages/{packageId}/add_to` — Document Workflow / Add Package To Category
- `POST /v1/packages/{packageId}/remove_from` — Document Workflow / Remove Package From Category
- `POST /v1/packages/add-to/bulk` — Document Workflow / Add Packages To Category (Bulk)
- `POST /v1/packages/remove-from/bulk` — Document Workflow / Remove Packages From Category (Bulk)
- `POST /v1/workflow/from-paper` — Document Workflow / Create Workflow From Paper
- `POST /v1/workflow/from-paper/bulk` — Document Workflow / Create Bulk Workflow From Paper
- `GET /v1/workflow/from-paper/bulk/{batchId}` — Document Workflow / Get Bulk Workflow From Paper
- `GET /v1/workflow/from-paper/{status}/{pageNo}/{recordsPerPage}` — Document Workflow / List Paper Conversions
- `GET /v1/workflow/from-paper/{recordId}` — Document Workflow / Get Paper Conversion
- `PUT /v1/workflow/from-paper/{recordId}` — Document Workflow / Update Paper Conversion
- `DELETE /v1/workflow/from-paper/{recordId}` — Document Workflow / Delete Paper Conversion
- `POST /v1/workflow/from-paper/{recordId}/approve` — Document Workflow / Approve Paper Conversion
- `GET /v1/connectors` — Connector / Get List Login Connector
- `GET /v1/connectors/certificate-profile` — Connector / Get Certificate Profiles for Connector
- `GET /v1/packages/SIGN/status` — RSSP Processing / Get Bulk Signing Status
- `DELETE /v1/packages/SIGN/status` — RSSP Processing / Release Bulk Signing Management
- `POST /v1/ras/request/cancel` — RSSP Processing / Cancel Authorize Signing Request
- `POST /v1/csc/credentials/manage` — RSSP Processing / Open Trusted SIC Management View
- `POST /v1/packages/SIGN` — RSSP Processing / Bulk Sign Packages
- `POST /v1/identity` — Account Management / Facial Matching
- `POST /v1/identity/verification/email-otp` — Account Management / Send Email OTP
- `POST /v1/settings/templates` — Template / Create Personal Template
- `POST /v1/enterprise/templates` — Template / Create Enterprise Template
- `GET /v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}` — Template / Get List Template
- `POST /v1/templates/{templateId}` — Template / Save Template
- `GET /v1/templates/{templateId}/file` — Template / Get Template File
- `GET /v1/templates/{templateId}/documents/{documentId}/fields` — Template / Get Template Fields
- `DELETE /v1/templates/{templateId}/documents/{documentId}/fields` — Template / Delete Template Field
- `GET /v1/enterprise/category` — Document Workflow / Get List Enterprise Categories
- `POST /v1/enterprise/category` — Document Workflow / Create Enterprise Category
- `PUT /v1/enterprise/category/{categoryId}` — Document Workflow / Update Enterprise Category
- `DELETE /v1/enterprise/category/{categoryId}` — Document Workflow / Delete Enterprise Category
- `GET /v1/enterprise/category/{categoryId}/permission` — Document Workflow / List Category Grants
- `POST /v1/enterprise/category/{categoryId}/permission` — Document Workflow / Grant Category Permission
- `DELETE /v1/enterprise/category/{categoryId}/permission` — Document Workflow / Revoke Category Permission
- `POST /v1/system/reload` — ADMIN APIs / Reload System Resources
- `POST /v1/system/serviceplans/sync` — ADMIN APIs / Sync Service Plans
- `POST /v1/system/response-codes` — ADMIN APIs / Add Response Code
- `GET /v1/system/cache` — ADMIN APIs / Get Cache Data
- `DELETE /v1/system/cache` — ADMIN APIs / Delete Cache Data
- `GET /v1/system/file/{linkSegment}` — Document Processing / Download file system
- `GET /v1/system/terms-of-service` — Account Management / Get List Terms of Service
- `GET /v1/terms` — Account Management / Get List Terms of Service
- `GET /v1/system/terms-of-service/{termsOfServiceUuid}` — Account Management / Download Terms of Service File
- `GET /v1/providers` — Register Certificate Processing / Get List Providers
- `GET /v1/system/serviceplans` — Enterprise Management / Get All Service Plans
- `GET /v1/system/serviceplans/{pageNo}/{recordsPerPage}` — Enterprise Management / Get All Service Plans
- `GET /v1/system/serviceplans/{servicePlanId}/billing/catalog` — Enterprise Management / Get Service Plan Billing Catalog
- `GET /v1/system/billing/catalog` — Enterprise Management / Get Billing Catalog
- `GET /v1/system/tbs/bank-accounts` — Enterprise Management / Get Trusted Billing Bank Accounts
- `GET /v1/system/tax-info` — Enterprise Management / Get Tax Company Name
- `GET /v1/system/signing/apps` — Document Processing / Get App Name Mapping List
- `GET /v1/system/logs/{offset}/{rowcount}` — ADMIN APIs / List System Logs
- `GET /v1/system/logs/{systemLogId}/details` — ADMIN APIs / Get System Log Details
- `POST /v1/admin/accounts-personal` — ADMIN APIs / Create Personal Account
- `POST /v1/admin/enterprises-tools` — ADMIN APIs / Create Enterprise Account
- `POST /v1/admin/enterprises-tools/lookup` — ADMIN APIs / Lookup Enterprise By Tax Code For Tools
- `POST /v1/admin/billing-webhooks/redrive` — ADMIN APIs / Re-drive Trusted Billing Payment Webhook
- `GET /v1/admin/users` — ADMIN APIs / Get User by Email or Citizen ID
- `GET /v1/admin/system/refresh-packages` — ADMIN APIs / Refresh Certificate Profiles
- `POST /v1/trusted-billing/customers` — Enterprise Management / Create Trusted Billing Customer
- `POST /v1/trusted-billing/bills` — Enterprise Management / Create Trusted Billing Bill
- `POST /rss/webhook` — ADMIN APIs / RSS Webhook
- `POST /rss/webhook/{signingTransactionId}` — ADMIN APIs / RSS Webhook - Receive event start processing
- `POST /billing/webhook` — ADMIN APIs / Trusted Billing Webhook
- `POST /billing/webhook/{billingTransactionId}` — ADMIN APIs / Trusted Billing Webhook - Receive event processing
- `POST /v1/settings/addressbook/{addressBookId}` — Personal Settings / Update AddressBook
- `POST /v1/application/queue` — Application/Device Management / Add API Request Queue
- `DELETE /v1/application/queue` — Application/Device Management / Add API Request Queue
- `POST /v1/application/users-biometric` — Application/Device Management / Register Application User with Biometric
- `POST /v1/application/users/check` — Application/Device Management / Check User Existed in KeyCloak
- `POST /v1/authenticate/tse` — Authenticate – Log in / Login TSE
- `GET /v1/authenticate/tse/status` — Authenticate – Log in / Check TSE Status
- `POST /v1/authenticate/tse/status` — Authenticate – Log in / Check TSE Status

## Paired endpoints with differences

### Client Credentials Authentication — `POST /v1/authenticate`
- request_fields only in spec: password, refresh_token, scope, username
- response_fields only in doc: refresh_token_expires_in
- response_fields only in spec: refresh_expires_in

### Password Authentication — `POST /v1/authenticate`
- request_fields only in spec: refresh_token
- response_fields only in doc: refresh_token_expires_in
- response_fields only in spec: refresh_expires_in

### Single Sign On Authentication — `POST /v1/authenticate/sso`
- headers only in spec: x-client-id, x-client-secret
- request_fields only in doc: remember_me_enabled
- request_fields only in spec: redirect_uri, scope
- response_fields only in doc: refresh_token_expires_in
- response_fields only in spec: refresh_expires_in

### Get Information — `GET /v1/info`
- headers only in doc: Authorization, Content-Type
- response_fields only in spec: gopaperless.workflow.app.authType, gopaperless.workflow.app.dateformat, gopaperless.workflow.app.description, gopaperless.workflow.app.langs, gopaperless.workflow.app.logo, gopaperless.workflow.app.methods, gopaperless.workflow.app.name, gopaperless.workflow.app.region, gopaperless.workflow.app.version, gopaperless.workflow.login, gopaperless.workflow.url, server_timestamp

### Pre Login Authentication — `POST /v1/authenticate/pre`
- headers only in doc: Authorization

### Pre Upload Document — `POST /v1/packages/documents/pre`
- headers only in doc: Content-Type

### Add Document Package — `POST /v1/packages`
- request_fields only in spec: createdBy, hmac, package_id, type
- response_fields only in spec: createdBy, hmac, package_name, type

### Share Document Package/ Share Workflow — `POST /v1/packages/{packageId}/workflow/share`
- name: doc `Share Document Package/ Share Workflow` / spec `Share Workflow`
- category: doc `Document Package` / spec `Document Workflow`
- headers only in doc: Content-Type

### Upload Document — `POST /v1/packages/{packageId}/documents`
- headers only in doc: Content-Type
- headers only in spec: x-preserve-form-field
- response_fields only in spec: document_height, document_mimetype, document_name, document_pages, document_rotate, document_size, document_type, document_width, package_id

### Upload Document (base64) — `POST /v1/packages/{packageId}/documents/base64`
- headers only in spec: x-preserve-form-field
- response_fields only in spec: document_height, document_mimetype, document_name, document_pages, document_rotate, document_size, document_type, document_width, package_id

### Download Package — `GET /v1/packages/{packageId}`
- headers only in doc: Content-Type, x-otp, x-password

### Download Package (base64) — `GET /v1/packages/{packageId}/base64`
- headers only in doc: Content-Type, x-otp, x-password

### Download Document — `GET /v1/packages/{packageId}/documents/{documentId}`
- headers only in doc: Content-Type, x-otp, x-password
- headers only in spec: x-original-file

### Download Document (base64) — `GET /v1/packages/{packageId}/documents/{documentId}/base64`
- headers only in doc: Content-Type, x-otp, x-password
- headers only in spec: x-original-file

### Get All Packages — `GET /v1/packages/{document_status}/{page_no}/{records_per_page}`
- headers only in doc: Content-Type, x-search-text, x-total-records
- headers only in spec: x-number-records
- path_params only in doc: document_status, page_no, records_per_page
- path_params only in spec: packageStatus, pageNo, pageSize

### Get Document Image — `GET /v1/packages/{packageId}/documents/{documentId}/images/{pageNo}`
- headers only in doc: Content-Type, x-otp, x-password
- headers only in spec: x-original-file

### Get Document Image Base64 — `GET /v1/packages/{packageId}/documents/{documentId}/images/{pageNo}/base64`
- headers only in doc: Content-Type, x-otp, x-password

### Get Document Details — `GET /v1/packages/{packageId}/documents/{documentId}/details`
- headers only in doc: Content-Type

### Get Document Verification — `GET /v1/packages/{packageId}/documents/{documentId}/verification`
- headers only in doc: Content-Type

### Get Package Details — `GET /v1/packages/{packageId}/details`
- headers only in doc: Content-Type

### Rename Package — `PUT /v1/packages/{packageId}`
- request_fields only in spec: createdBy, hmac, package_id, type

### Rename Document — `PUT /v1/packages/{packageId}/documents/{documentId}`
- request_fields only in spec: document_height, document_id, document_mimetype, document_pages, document_rotate, document_size, document_type, document_width, package_id

### Delete Package — `DELETE /v1/packages/{packageId}`
- headers only in doc: Content-Type

### Delete Document — `DELETE /v1/packages/{packageId}/documents/{documentId}`
- headers only in doc: Content-Type

### Update Document — `PUT /v1/packages/{packageId}/documents/{documentId}/update`
- headers only in doc: Content-Type
- headers only in spec: x-preserve-form-field, x-sort-order

### Update Document (base64) — `PUT /v1/packages/{packageId}/documents/{documentId}/update/base64`
- headers only in spec: x-preserve-form-field
- request_fields only in spec: document_order

### Merge Document — `POST /v1/packages/{packageId}/documents/{documentId}/merge`
- headers only in doc: Content-Type

### Apply Workflow Template — `POST /v1/packages/{packageId}/documents/{documentId}/template`
- request_fields only in spec: template_id

### Download Signature Summary — `GET /v1/packages/{packageId}/documents/{documentId}/signature-summary`
- headers only in doc: Content-Type

### Download Signature Summary (base64) — `GET /v1/packages/{packageId}/documents/{documentId}/signature-summary/base64`
- headers only in doc: Content-Type

### Start new Workflow — `POST /v1/packages/{packageId}/workflow`
- request_fields only in spec: categories, category_id, is_econtract, message

### Create Workflow From Template — `POST /v1/workflow/from-template`
- response_fields only in doc: document_id
- response_fields only in spec: document_ids

### Share Workflow — `POST /v1/packages/{packageId}/workflow/share`
- headers only in doc: Content-Type

### Get All Workflow — `GET /v1/workflow/{criteria}/{page_no}/{records_per_page}`
- headers only in doc: Content-Type
- headers only in spec: x-category-mode, x-enterprise-mode, x-filter-field-type, x-from-date, x-to-date, x-workflow-type
- path_params only in doc: page_no, records_per_page
- path_params only in spec: pageNo, recordsPerPage

### Update Workflow Details — `PUT /v1/packages/{packageId}/workflow`
- request_fields only in doc: details
- request_fields only in spec: categories, is_econtract

### Delete Workflow — `DELETE /v1/packages/{packageId}/workflow`
- headers only in doc: Content-Type

### Get Workflow Details — `GET /v1/packages/{packageId}/workflow`
- headers only in doc: Content-Type
- headers only in spec: x-catch-event, x-enterprise-mode
- response_fields only in doc: next_signer_email, unread
- response_fields only in spec: owner_email

### Get Workflow Details by UUID — `GET /v1/workflow/{workflowUuid}`
- headers only in doc: Content-Type
- headers only in spec: x-catch-event
- response_fields only in doc: next_signer_email, unread
- response_fields only in spec: owner_email

### View Workflow — `POST /v1/workflow/{uuid}/view`
- headers only in spec: Authorization, x-coordinate-origin, x-dimension-unit
- path_params only in doc: uuid
- path_params only in spec: workflowUuid
- response_fields only in doc: next_signer, unread
- response_fields only in spec: owner_email

### Get Workflow Users — `GET /v1/packages/{packageId}/workflow/users`
- headers only in doc: Content-Type

### Update Workflow User — `PUT /v1/packages/{packageId}/workflow/{order}/user`
- request_fields only in spec: company_name, country, guest_user, reason, user_first_name, user_last_name, user_national_id
- response_fields only in doc: email_language_code, guest_user, signing_order, user_email

### Delete Workflow User — `DELETE /v1/packages/{packageId}/workflow/{order}`
- headers only in doc: Content-Type

### Get Workflow Reminders — `GET /v1/packages/{packageId}/workflow/reminders`
- name: doc `Get Workflow Reminders` / spec `Get All Workflow`
- headers only in doc: Content-Type
- headers only in spec: x-category-id, x-category-mode, x-enterprise-mode, x-filter-field-type, x-from-date, x-search-text, x-to-date, x-total-records, x-workflow-type
- path_params only in doc: packageId
- path_params only in spec: criteria, pageNo, recordsPerPage
- response_fields only in doc: enabled, frequency, keep_reminding_after, remind_after, total_reminders
- response_fields only in spec: workflows

### Send Reminder — `POST /v1/packages/{packageId}/workflow/{order}/remind`
- headers only in doc: Content-Type

### Get Workflow History — `GET /v1/packages/{packageId}/log/{page_no}/{records_per_page}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- path_params only in doc: page_no, records_per_page
- path_params only in spec: pageNo, recordsPerPage

### Get Workflow Comment/Message — `GET /v1/packages/{packageId}/comments`
- headers only in doc: Content-Type
- headers only in spec: x-total-records

### Delete Workflow Comment/Message — `DELETE /v1/packages/{packageId}/comments/{commentId}`
- headers only in doc: Content-Type

### Create new Category — `POST /v1/category`
- request_fields only in spec: scope

### Update Category — `PUT /v1/category/{categoryId}`
- request_fields only in spec: scope

### Get list of Category — `GET /v1/category`
- headers only in doc: Content-Type
- headers only in spec: x-enterprise-mode, x-total-records

### Delete Category — `DELETE /v1/category/{categoryId}`
- headers only in doc: Content-Type

### Add Placeholder to Workflow — `POST /v1/packages/{packageId}/workflow/placeholder`
- response_fields only in doc: placeholders
- response_fields only in spec: users

### Update Placeholder in Workflow — `PUT /v1/packages/{packageId}/workflow/{order}/placeholder`
- request_fields only in spec: email, email_language_code, guest_user, mobile_number

### Import Excel to Clone Workflow — `POST /v1/packages/{packageId}/workflow/{order}/import`
- headers only in spec: x-coordinate-origin, x-dimension-unit
- request_fields only in doc: items[].mobile_number, items[].personal_code, items[].user_email, items[].username, items[].vneid_birthday, items[].vneid_current_place, items[].vneid_date_expired, items[].vneid_date_issuer, items[].vneid_fullname, items[].vneid_gender, items[].vneid_level_account, items[].vneid_place, items[].vneid_type_account
- response_fields only in spec: status, transaction_id

### Get Status of Import Excel — `GET /v1/workflow/{transactionCode}/import`
- headers only in doc: Content-Type
- path_params only in doc: transactionCode
- path_params only in spec: importId

### Upload Attachment — `POST /v1/packages/{packageId}/documents/{documentId}/attachments`
- headers only in doc: Content-Type

### Get Attachments — `GET /v1/packages/{packageId}/documents/{documentId}/attachments`
- headers only in doc: Content-Type

### Delete Attachment — `DELETE /v1/packages/{packageId}/documents/{documentId}/attachments/{attachmentId}`
- headers only in doc: Content-Type
- response_fields only in doc: attachment_id

### Download Attachment — `GET /v1/packages/{packageId}/documents/{documentId}/attachments/{attachmentId}`
- headers only in doc: Content-Type

### Add Digital Signature Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/signature`
- request_fields only in doc: appearance_design, authentication_signing, level_of_assurance, remark, required, verification
- request_fields only in spec: renamed_as, rotate, type_name

### Update Digital Signature Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/signature`
- request_fields only in doc: appearance_design, authentication_signing, level_of_assurance, remark, required, verification
- request_fields only in spec: rotate, type_name
- response_fields only in doc: field_name

### Add TextBox Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/text`
- request_fields only in doc: align_option, color, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, type, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update TextBox Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/text`
- request_fields only in doc: align_option, color, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name
- response_fields only in doc: field_name

### Add Initial Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/initials`
- request_fields only in doc: apply_to_all, image, remark, replicate_all_pages, replicate_pages, required
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Initial Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/initials`
- request_fields only in doc: apply_to_all, image, remark, replicate_all_pages, replicate_pages, required
- request_fields only in spec: annotation_position, rotate, type_name

### Delete Document Field — `DELETE /v1/packages/{packageId}/documents/{documentId}/fields`
- headers only in doc: Content-Type

### Get Document Fields — `GET /v1/packages/{packageId}/documents/{documentId}/fields`
- headers only in doc: Content-Type
- headers only in spec: x-coordinate-origin, x-dimension-unit
- response_fields only in spec: qrypto, redaction, vneid_field

### Get Field Types — `GET /v1/documents/fields`
- headers only in doc: Content-Type

### Add In-person Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/in_person_signature`
- request_fields only in doc: appearance_design, authentication_signing, level_of_assurance, order, remark, required, verification
- request_fields only in spec: renamed_as, rotate, type_name

### Update In-person Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/in_person_signature`
- request_fields only in doc: appearance_design, authentication_signing, level_of_assurance, order, remark, required, verification
- request_fields only in spec: rotate, type_name
- response_fields only in doc: field_name

### Add Datetime Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/datetime`
- request_fields only in doc: align_option, color, date, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Datetime Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/datetime`
- request_fields only in doc: align_option, color, date, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add CheckBox Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/checkbox`
- request_fields only in doc: align_option, check_box_type, checked, color, font, group_name, multiple_checked, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update CheckBox Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/checkbox`
- request_fields only in doc: align_option, check_box_type, checked, color, font, group_name, multiple_checked, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add RadioBox Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/radiobox`
- request_fields only in doc: align_option, check_box_type, checked, color, font, multiple_checked, place_holder, radio_group_name, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update RadioBox Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/radiobox`
- request_fields only in doc: align_option, check_box_type, checked, color, font, multiple_checked, place_holder, radio_group_name, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add QR Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/qrcode`
- request_fields only in doc: allow_unauthenticated_workflow_view, custom_url, image_qr, is_transparent, read_only, remark, required, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update QR Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/qrcode`
- request_fields only in doc: image_qr, is_transparent, read_only, remark, required
- request_fields only in spec: annotation_position, rotate, type_name

### Add ComboBox Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/combobox`
- request_fields only in doc: align_option, color, combo_item, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update ComboBox Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/combobox`
- request_fields only in doc: align_option, color, combo_item, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add Toggle Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/toggle`
- request_fields only in doc: align_option, color, font, format_type, max_length, multiline, place_holder, read_only, remark, required, toggle_item, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Toggle Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/toggle`
- request_fields only in doc: align_option, color, font, format_type, max_length, multiline, place_holder, read_only, remark, required, toggle_item, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add Numeric Stepper Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/numeric_stepper`
- request_fields only in doc: align_option, color, font, format_type, max_length, multiline, numeric_stepper_items, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Numeric Stepper Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/numeric_stepper`
- request_fields only in doc: align_option, color, font, format_type, max_length, multiline, numeric_stepper_items, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add Hyperlink Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/hyperlink`
- request_fields only in doc: address, align_option, color, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Hyperlink Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/hyperlink`
- request_fields only in doc: address, align_option, color, font, format_type, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, value
- request_fields only in spec: annotation_position, rotate, type_name

### Add Stamp Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/stamp`
- request_fields only in doc: image, place_holder, read_only, remark, required, tool_tip_text
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Stamp Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/stamp`
- request_fields only in doc: image, place_holder, read_only, remark, required, tool_tip_text
- request_fields only in spec: annotation_position, rotate, type_name

### Add Camera Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/camera`
- request_fields only in doc: camera, place_holder, read_only, remark, required, tool_tip_text
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Camera Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/camera`
- request_fields only in doc: camera, place_holder, read_only, remark, required, tool_tip_text
- request_fields only in spec: annotation_position, rotate, type_name

### Add Attachment Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/attachment`
- request_fields only in doc: file, place_holder, read_only, remark, required, show_icon_enabled, tool_tip_text
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Attachment Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/attachment`
- request_fields only in doc: file, place_holder, read_only, remark, required, show_icon_enabled, tool_tip_text
- request_fields only in spec: annotation_position, rotate, type_name

### Add CheckBox Field Version 2 — `POST /v1/packages/{packageId}/documents/{documentId}/fields/checkboxV2`
- request_fields only in doc: align_option, checkbox_frame, checked, font, group_name, multiple_checked, place_holder, read_only, remark, required, text_next_to, tool_tip_text, uncheckbox_frame
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update CheckBox Field Version 2 — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/checkboxV2`
- request_fields only in doc: align_option, checkbox_frame, checked, font, group_name, multiple_checked, place_holder, read_only, remark, required, text_next_to, tool_tip_text, uncheckbox_frame
- request_fields only in spec: annotation_position, rotate, type_name

### Add RadioBox Field Version 2 — `POST /v1/packages/{packageId}/documents/{documentId}/fields/radioboxV2`
- request_fields only in doc: align_option, checkbox_frame, checked, font, group_name, multiple_checked, place_holder, read_only, remark, required, text_next_to, tool_tip_text, uncheckbox_frame
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update RadioBox Field Version 2 — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/radioboxV2`
- request_fields only in doc: align_option, checkbox_frame, checked, font, group_name, multiple_checked, place_holder, read_only, remark, required, text_next_to, tool_tip_text, uncheckbox_frame
- request_fields only in spec: annotation_position, rotate, type_name

### Add VNeID Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/vneid`
- request_fields only in doc: align_option, color, font, format_type, loaded_from_vneid, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, type, value
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update VNeID Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/vneid`
- request_fields only in doc: align_option, color, font, format_type, loaded_from_vneid, max_length, multiline, place_holder, read_only, remark, required, tool_tip_text, type, value
- request_fields only in spec: annotation_position, rotate, type_name
- response_fields only in doc: field_name

### Add Redact Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/redaction`
- request_fields only in doc: read_only, redact_policy, remark, required
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Redact Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/redaction`
- request_fields only in doc: read_only, redact_policy, remark, required
- request_fields only in spec: annotation_position, rotate, type_name

### Add Elabor Field — `POST /v1/packages/{packageId}/documents/{documentId}/fields/elabor`
- request_fields only in doc: alignment, color, font, max_length, multiline, read_only, remark, required, type
- request_fields only in spec: annotation_position, renamed_as, rotate, type_name

### Update Elabor Field — `PUT /v1/packages/{packageId}/documents/{documentId}/fields/elabor`
- request_fields only in doc: alignment, color, font, max_length, multiline, read_only, remark, required, type
- request_fields only in spec: annotation_position, rotate, type_name
- response_fields only in doc: field_name

### Sign Eseal — `POST /v1/packages/{packageId}/documents/{documentId}/sign_eseal`
- request_fields only in spec: signing_reason

### Fill Form Fields — `PUT /v1/packages/{packageId}/documents/{documentId}/fields`
- request_fields only in doc: attachment, camera, checkbox, checkboxv2, combobox, custom_field, date, datetime, hyperlink, numeric_stepper, qr, radiobox, radioboxv2, stamp, textbox, toggle, vneid_field

### Get List USB-Token Connector — `GET /v1/connectors/usbtoken`
- headers only in doc: Content-Type, x-search-server

### Get List Smart-ID Connector — `GET /v1/connectors/smartid`
- headers only in doc: Content-Type, x-search-server

### Get List Internal Signing Service Connector (ISS) — `GET /v1/connectors/iss`
- headers only in doc: Content-Type, x-search-server

### Get RSSP Information — `GET /v1/system/signing/servers`
- headers only in doc: Content-Type

### Get Filtered Credential List — `POST /v1/csc/credentials/list`
- request_fields only in spec: user_name, user_type

### Sign Document via RSSP Directly — `POST /v1/csc/packages/{packageId}/documents/{documentId}/sign_rssp`
- request_fields only in doc: webhook_url
- response_fields only in spec: mobile_number, signing_transaction_id, status, transaction_id, user_email

### RSSP Credential Authorization — `POST /v1/csc/credentials/authorization`
- request_fields only in doc: responseID

### Get Signing Status (TSE MID) — `POST /v1/csc/packages/sign/status`
- request_fields only in doc: transaction_id
- request_fields only in spec: notification_template, signing_transaction_id

### Get Credentials Info — `POST /v1/csc/credentials/info`
- headers only in spec: x-citizen-id

### Send OTP via RSSP — `POST /v1/csc/credentials/send-otp`
- request_fields only in doc: transaction_id
- request_fields only in spec: signing_transaction_id

### Get Signing Transaction Status — `GET /v1/csc/signing/{transactionCode}/status`
- headers only in doc: Content-Type
- response_fields only in doc: data, metadata, status, type, user_info, webhook_url

### Get Pending Signing Sessions — `GET /v1/csc/signing/pending-sessions`
- headers only in doc: Content-Type

### Get Sign Hash From Service — `POST /v1/csc/signatures/sign-hash`
- response_fields only in doc: documents, hashing_algo, request_id, signing_algo, status, transaction_code, validity_period

### Register Certificate via RSSP — `POST /v1/csc/credentials/register`
- request_fields only in spec: certificate_profile, data, issuing_provider
- response_fields only in doc: is_mobile_verified
- response_fields only in spec: is_mobile_veriried, taxcode_enabled

### Get List Tax Codes — `GET /v1/csc/credentials/{sessionId}/taxcode`
- headers only in doc: Content-Type
- headers only in spec: x-citizen-temp

### Finish Register Certificate — `POST /v1/csc/credentials/{sessionId}/finish`
- response_fields only in spec: SCAL, auth_mode, cert, fields, multisign

### Get Credential Registration Status — `GET /v1/csc/credentials/{transactionCode}/status`
- headers only in doc: Content-Type

### Cancel Issue Certificate — `POST /v1/ras/issue/cancel`
- headers only in doc: Content-Type

### Register Enterprise User — `POST /v1/enterprise/users`
- request_fields only in doc: certificate_id, user_csp_id, user_ra_id
- request_fields only in spec: city, expired_time, user_first_name, user_last_name
- response_fields only in spec: created_at, id

### Get Enterprise Invitations — `GET /v1/enterprise/invitations/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- response_fields only in doc: created_at, enterprise_role, id, invited_by, status, user_email, user_name, user_name_alias

### Resend Enterprise User Invitation — `POST /v1/enterprise/invitations/{invitationId}/resend`
- headers only in doc: Content-Type

### Delete Enterprise User Invitation — `DELETE /v1/enterprise/invitations/{invitationId}/delete`
- headers only in doc: Content-Type

### Export Enterprise Invitations — `GET /v1/enterprise/invitations/export/{status}/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records

### Get Enterprise Users — `GET /v1/enterprise/users/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-search-text, x-total-records
- request_fields only in doc: pageNo, recordsPerPage, x-search-text
- response_fields only in doc: users

### Export Enterprise Users — `GET /v1/enterprise/users/export/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records

### Get Enterprise User Quota — `GET /v1/enterprise/users/{enterpriseUserId}/quota`
- headers only in doc: Content-Type

### Get Enterprise User Management Statistics — `GET /v1/enterprise/users/statistics`
- headers only in doc: Content-Type

### Delete Enterprise User — `DELETE /v1/enterprise/users/{enterpriseUserId}`
- headers only in doc: Content-Type

### Get Roles — `GET /v1/enterprise/roles/{pageNo}/{recordPerPage}`
- headers only in doc: Content-Type, x-search-text
- path_params only in doc: recordPerPage
- path_params only in spec: recordsPerPage
- request_fields only in doc: pageNo, recordsPerPage
- response_fields only in doc: roles

### Delete Role — `DELETE /v1/enterprise/roles/{id}`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: enterpriseRoleId
- request_fields only in doc: id

### Add Role — `POST /v1/enterprise/roles`
- request_fields only in doc: default, description
- request_fields only in spec: feature, feature_definition_version, features, permissions, remark
- response_fields only in doc: default, description, id
- response_fields only in spec: assign_enabled, default_enabled, enterprise_role_id, permissions, permissions_granted, role_id

### Get Enterprise Custom Fields — `GET /v1/enterprise/custom_fields/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- response_fields only in doc: created_at, date, description, field_name, format_type, id, modified_at, parent_type, place_holder, suffix, type_name

### Create Enterprise Custom Field — `POST /v1/enterprise/custom_fields`
- request_fields only in spec: created_at, field_name, id, modified_at, parent_type
- response_fields only in spec: date, description, field_name, format_type, modified_at, place_holder, suffix

### Update Enterprise Custom Field — `PUT /v1/enterprise/custom_fields/{fieldId}`
- request_fields only in spec: created_at, field_name, id, modified_at, parent_type
- response_fields only in spec: created_at, date, description, field_name, format_type, parent_type, place_holder, suffix

### Delete Enterprise Custom Field — `DELETE /v1/enterprise/custom_fields/{fieldId}`
- headers only in doc: Content-Type

### Get Enterprise Electronic Seals — `GET /v1/enterprise/electronic-seals`
- headers only in doc: Content-Type
- headers only in spec: x-search-text
- response_fields only in doc: appearance, basic_information, detail, id
- response_fields only in spec: seals

### Get Enterprise TSA Configuration — `GET /v1/enterprise/settings/tsa`
- headers only in doc: Content-Type

### Test Connection Enterprise TSA Configuration — `POST /v1/enterprise/settings/tsa/test-connection`
- request_fields only in spec: hash_algorithm, policy_oid, retry_count, tsa_users

### Get Enterprise Report List — `GET /v1/enterprise/reports/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- response_fields only in doc: dms_property, document_uuid, enterprise_id, export_date, id, last_modified_at, last_modified_by, name

### Get Enterprise Quota Statistics — `GET /v1/enterprise/reports/quota/statistics/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-from-date, x-to-date

### Export Enterprise Workflow Statistics — `GET /v1/enterprise/reports/workflows/statistics/export/{status}/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: Accept-Language, x-total-records

### Get Enterprise Workflows — `GET /v1/enterprise/workflows/{status}/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records

### Get Account Invitations — `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- request_fields only in doc: user_email
- response_fields only in doc: invitations

### Get Account Invitation By ID — `GET /v1/account/invitations/{invitationId}`
- headers only in doc: Content-Type
- path_params only in doc: invitationId
- path_params only in spec: invitationReference

### Get Account Enterprises — `GET /v1/account/enterprises`
- headers only in doc: Content-Type
- response_fields only in doc: enterprise_id, enterprise_name, is_active, logo_url, role

### Get Account Profiles — `GET /v1/account/profiles`
- headers only in doc: Content-Type
- response_fields only in doc: enterprise_id, enterprise_name, is_active, logo_url, profile_id, role, service_plan, user_email

### Account Usage Statistics — `GET /v1/account/statistics/usage`
- headers only in doc: Content-Type
- response_fields only in spec: balance, elabor_contract, enterprise_quota, service_plan_duration, template, upload_size, validation

### Document Statistics — `GET /v1/account/statistics/documents`
- headers only in doc: Content-Type
- headers only in spec: x-enterprise-mode
- response_fields only in doc: archived, shared
- response_fields only in spec: all, expired, expiring_soon, inprogress

### Get Signing History — `GET /v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-fe-method-name, x-total-records
- request_fields only in doc: certificateAuthorityName, citizen_id, connector_name, credential_id

### Get Certificate Enroll History — `GET /v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-criteria-status, x-total-records
- request_fields only in doc: application_id, attribute_name, attribute_value, certificate_authority_vneid_id, certificate_profile_id, citizen_id, connector_name, user_email

### Get User Activity Logs — `GET /v1/account/activity-logs/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- request_fields only in doc: pageNo, recordsPerPage
- response_fields only in doc: user_activities
- response_fields only in spec: items, page, records_per_page, total_count

### Get Dashboard Workflows — `GET /v1/account/dashboard/workflows`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by
- response_fields only in doc: approved, signed, updated

### Get Dashboard Quota — `GET /v1/account/dashboard/quota`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by
- response_fields only in spec: validation

### Get Dashboard Workflow Types — `GET /v1/account/dashboard/workflow-types`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by, x-total-records

### Get Dashboard Categories — `GET /v1/account/dashboard/categories`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by, x-total-records

### Get Dashboard Validations — `GET /v1/account/dashboard/validations`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by

### Get Dashboard Overdue Workflows — `GET /v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by, x-total-records

### Get Dashboard Activities — `GET /v1/account/dashboard/activities/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-asc, x-enterprise-mode, x-sort-by, x-total-records

### Add Identity for a User — `POST /v1/account/identity`
- request_fields only in spec: async, citizen_id, metadata, package_id
- response_fields only in spec: billcode_id, transaction_code, transaction_id

### Initialize Billing — `POST /v1/account/billings/initialize`
- request_fields only in spec: bank_account_id
- response_fields only in spec: account_number, billing_payment_id

### Get Billing Payment History — `GET /v1/account/billings/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- request_fields only in doc: connector_name, payment_number, status, type

### Get Billing Payment Details — `GET /v1/account/billings/payments/{billingPaymentId}`
- headers only in doc: Content-Type

### Get Account Quota Consume History — `GET /v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type
- headers only in spec: x-total-records
- response_fields only in doc: added_balance, created_at, id, quantity, quota_dimension, remaining_balance, source_system_type, source_system_type_remark, sync_status, sync_status_remark, type, type_remark, used_balance, workflow_name

### Remove Account Request — `POST /v1/account/remove`
- request_fields only in spec: userEmail

### Get Signature Settings — `GET /v1/settings/signatures`
- headers only in doc: Content-Type

### Update Signature Appearance — `PUT /v1/settings/signatures/{appearance_id}`
- path_params only in doc: appearance_id
- path_params only in spec: appearanceId
- request_fields only in doc: appearance_id

### Delete Signature Appearance — `DELETE /v1/settings/signatures/{appearance_id}`
- headers only in doc: Content-Type
- path_params only in doc: appearance_id
- path_params only in spec: appearanceId
- request_fields only in doc: appearance_id

### Get General Profile Information — `GET /v1/settings/profile`
- headers only in doc: Content-Type
- response_fields only in spec: application, feature_permissions, permissions, service_plan, tbs

### Update General Profile Information — `PUT /v1/settings/profile/general`
- request_fields only in doc: country, language, time_zone, user_name
- request_fields only in spec: account_type, date_of_birth, date_of_expiry, date_of_issue, eid_account_level, full_name, gender, notifications, permanent_address, residential_address, security_answer, security_question, service_agreements, service_plan, user_email, user_first_name, user_last_name, user_national_id, user_password

### Get Address Book — `GET /v1/settings/addressbook`
- headers only in doc: Content-Type
- headers only in spec: x-total-records

### Delete Address Book — `DELETE /v1/settings/addressbook/{id}`
- headers only in doc: Content-Type
- path_params only in doc: Id
- path_params only in spec: addressBookId

### Get Contacts — `GET /v1/settings/contacts/{pageNo}/{recordPerPage}`
- headers only in doc: Content-Type, x-enterprise, x-search-text
- path_params only in doc: recordPerPage
- path_params only in spec: recordsPerPage
- request_fields only in doc: pageNo, recordsPerPage

### Add Personal Contact — `POST /v1/settings/contacts`
- request_fields only in spec: address_book_ids, contact_id, created_at, enterprise_id, national_id

### Update Personal Contact — `PUT /v1/settings/contacts/{id}`
- path_params only in doc: id
- path_params only in spec: contactId
- request_fields only in doc: id
- request_fields only in spec: address_book_ids, address_book_name, contact_id, created_at, enterprise_id, national_id

### Delete Personal Contact — `DELETE /v1/settings/contacts/{id}`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: contactId
- request_fields only in doc: id

### Register Application User — `POST /v1/application/users`
- headers only in doc: Authorization
- headers only in spec: x-client-id, x-client-secret, x-skip-login
- request_fields only in doc: certificate_id, common_name, country, email_notification, language, time_zone, user_csc_id, user_csp_id, user_ra_id
- request_fields only in spec: notifications, service_agreements, service_plan
- response_fields only in spec: access_token, expires_in, refresh_expires_in, refresh_token, token_type

### Save Workflow as Template — `POST /v1/packages/{packageId}/template`
- request_fields only in spec: allow_signing_after_deadline, categories, continue_on_decline, deadline_at, filter_level, is_econtract, is_enterprise, message, no_auth_required_participants, public, read_only, workflow_mode
- response_fields only in doc: Id
- response_fields only in spec: id

### Get Templates — `GET /v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}`
- headers only in doc: Content-Type
- path_params only in doc: recordPerPage
- path_params only in spec: recordsPerPage
- request_fields only in doc: description, name
- response_fields only in doc: created_at, description, documents, enterprise, read_only, smart_form, template_id, template_name, template_public

### Get Template By ID — `GET /v1/templates/{id}`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: templateId
- response_fields only in spec: lifecycle_state

### Update Template — `PUT /v1/templates/{id}`
- path_params only in doc: id
- path_params only in spec: templateId
- request_fields only in spec: allow_signing_after_deadline, categories, continue_on_decline, deadline_at, filter_level, is_econtract, message, no_auth_required_participants, workflow_mode

### Delete Template — `DELETE /v1/templates/{id}`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: templateId

### Get Template Workflow Users — `GET /v1/templates/{id}/workflow/users`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: templateId
- response_fields only in doc: email_language_code, email_notification, guest_user, mobile_number, order, placeholder, role, signing_order, user_email, user_name

### Add Template Workflow Users — `POST /v1/templates/{id}/workflow/users`
- path_params only in doc: id
- path_params only in spec: templateId
- request_fields only in doc: delivery_method, email_language_code, email_notification, mobile_number, role, signing_order, user_email, user_name
- response_fields only in doc: email_language_code, email_notification, guest_user, mobile_number, order, placeholder, role, signing_order, user_email, user_name

### Add Template Workflow Placeholder — `POST /v1/templates/{id}/workflow/placeholder`
- path_params only in doc: id
- path_params only in spec: templateId
- request_fields only in doc: email_notification, placeholder, role, signing_order
- response_fields only in doc: email_language_code, email_notification, guest_user, mobile_number, order, placeholder, role, signing_order, user_email, user_name

### Update Template Workflow User — `PUT /v1/templates/{id}/workflow/{order}/user`
- path_params only in doc: id
- path_params only in spec: templateId
- request_fields only in spec: guest_enabled
- response_fields only in spec: delivery_method, logged_in_user

### Update Template Workflow Placeholder — `PUT /v1/templates/{id}/workflow/{order}/placeholder`
- path_params only in doc: id
- path_params only in spec: templateId
- request_fields only in spec: guest_enabled
- response_fields only in spec: delivery_method, logged_in_user

### Reorder Template Workflow Participant — `PUT /v1/templates/{id}/workflow/{order}/reorder`
- path_params only in doc: id
- path_params only in spec: templateId

### Delete Template Workflow Participant — `DELETE /v1/templates/{id}/workflow/{order}`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: templateId

### Clone Template — `POST /v1/templates/{id}/clone`
- path_params only in doc: id
- path_params only in spec: templateId
- request_fields only in spec: public, read_only

### Get Template Details — `GET /v1/templates/{id}/details`
- headers only in doc: Content-Type
- path_params only in doc: id
- path_params only in spec: templateId
- response_fields only in doc: created_on, description, enterprise, id, name, participants, public, read_only, tasks, workflow_categories, workflow_details
- response_fields only in spec: template, template_id, users

### Add Template Field — `POST /v1/templates/{id}/documents/{documentId}/fields/{fieldType}`
- path_params only in doc: id
- path_params only in spec: templateId

### Update Template Field — `PUT /v1/templates/{id}/documents/{documentId}/fields/{fieldType}`
- path_params only in doc: id
- path_params only in spec: templateId

### Get List Validation — `GET /v1/validations/{pageNo}/{recordsPerPage}`
- headers only in doc: Content-Type

### Get Report of Validation — `GET /v1/validations/{validationId}/reports`
- headers only in doc: Content-Type

### Download Reports — `GET /v1/validations/reports/{reportId}/{reportType}`
- headers only in doc: Content-Type

### Delete Validation — `DELETE /v1/validations/{validationId}`
- headers only in doc: Content-Type

### Create Account — `POST /v1/admin/accounts`
- request_fields only in doc: service_plan
- request_fields only in spec: address, contact_address, contact_email, industry_type, invoice_address, invoice_email, service_plan_id, tax_code

