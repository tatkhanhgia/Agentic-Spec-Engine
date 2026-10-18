# CHANGELOG

Sinh tự động bởi `scripts/changelog.py` bằng cách so sánh
hai bản chụp tài liệu. Không viết tay.

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 265 | 265 |
| Đoạn văn | 50785 | 50785 |
| Bảng | 1033 | 1033 |
| Khung mẫu | 1598 | 1598 |

### Đổi nội dung chữ (1 mục)
- Document Workflow / Get Workflow Users

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 219 | 265  **+46** |
| Đoạn văn | 42823 | 50735  **+7912** |
| Bảng | 885 | 1033  **+148** |
| Khung mẫu | 1320 | 1598  **+278** |

### Đã thêm
- `DELETE /v1/enterprise/custom_fields/{fieldId}` — Delete Enterprise Custom Field *(Enterprise Management)*
- `DELETE /v1/enterprise/invitations/{invitationId}` — Cancel Enterprise User Invitation *(Enterprise Management)*
- `DELETE /v1/enterprise/invitations/{invitationId}/delete` — Delete Enterprise User Invitation *(Enterprise Management)*
- `DELETE /v1/validations/{validationId}` — Delete Validation *(Validation Management)*
- `GET /v1/account/billings/payments/{billingPaymentId}` — Get Billing Payment Details *(Account Management)*
- `GET /v1/account/billings/{pageNo}/{recordsPerPage}` — Get Billing Payment History *(Account Management)*
- `GET /v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}` — Get Certificate Enroll History *(Account Management)*
- `GET /v1/account/dashboard/activities/{pageNo}/{recordsPerPage}` — Get Dashboard Activities *(Account Management)*
- `GET /v1/account/dashboard/categories` — Get Dashboard Categories *(Account Management)*
- `GET /v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}` — Get Dashboard Overdue Workflows *(Account Management)*
- `GET /v1/account/dashboard/processing-trend` — Get Dashboard Processing Trend *(Account Management)*
- `GET /v1/account/dashboard/quota` — Get Dashboard Quota *(Account Management)*
- `GET /v1/account/dashboard/validations` — Get Dashboard Validations *(Account Management)*
- `GET /v1/account/dashboard/workflow-types` — Get Dashboard Workflow Types *(Account Management)*
- `GET /v1/account/dashboard/workflows` — Get Dashboard Workflows *(Account Management)*
- `GET /v1/account/enterprises` — Get Account Enterprises *(Account Management)*
- `GET /v1/account/invitations/{invitationId}` — Get Account Invitation By ID *(Account Management)*
- `GET /v1/account/profiles` — Get Account Profiles *(Account Management)*
- `GET /v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}` — Get Account Quota Consume History *(Account Management)*
- `GET /v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}` — Get Signing History *(Account Management)*
- `GET /v1/enterprise/custom_fields/{pageNo}/{recordsPerPage}` — Get Enterprise Custom Fields *(Enterprise Management)*
- `GET /v1/enterprise/electronic-seals` — Get Enterprise Electronic Seals *(Enterprise Management)*
- `GET /v1/enterprise/invitations/export/{status}/{pageNo}/{recordsPerPage}` — Export Enterprise Invitations *(Enterprise Management)*
- `GET /v1/enterprise/invitations/{pageNo}/{recordsPerPage}` — Get Enterprise Invitations *(Enterprise Management)*
- `GET /v1/enterprise/reports/quota/statistics/{pageNo}/{recordsPerPage}` — Get Enterprise Quota Statistics *(Enterprise Management)*
- `GET /v1/enterprise/reports/workflows/statistics/export/{status}/{pageNo}/{recordsPerPage}` — Export Enterprise Workflow Statistics *(Enterprise Management)*
- `GET /v1/enterprise/reports/{pageNo}/{recordsPerPage}` — Get Enterprise Report List *(Enterprise Management)*
- `GET /v1/enterprise/settings/tsa` — Get Enterprise TSA Configuration *(Enterprise Management)*
- `GET /v1/enterprise/users/export/{pageNo}/{recordsPerPage}` — Export Enterprise Users *(Enterprise Management)*
- `GET /v1/enterprise/users/statistics` — Get Enterprise User Management Statistics *(Enterprise Management)*
- `GET /v1/enterprise/users/{enterpriseUserId}/quota` — Get Enterprise User Quota *(Enterprise Management)*
- `GET /v1/enterprise/workflows/{status}/{pageNo}/{recordsPerPage}` — Get Enterprise Workflows *(Enterprise Management)*
- `GET /v1/validations/reports/{reportId}/{reportType}` — Download Reports *(Validation Management)*
- `GET /v1/validations/{pageNo}/{recordsPerPage}` — Get List Validation *(Validation Management)*
- `GET /v1/validations/{validationId}/reports` — Get Report of Validation *(Validation Management)*
- `POST /v1/account/billings/initialize` — Initialize Billing *(Account Management)*
- `POST /v1/account/invitations/accept` — Accept Account Invitation *(Account Management)*
- `POST /v1/account/invitations/reject` — Reject Account Invitation *(Account Management)*
- `POST /v1/account/remove` — Remove Account Request *(Personal Settings)*
- `POST /v1/enterprise/custom_fields` — Create Enterprise Custom Field *(Enterprise Management)*
- `POST /v1/enterprise/invitations` — Invite Enterprise User *(Enterprise Management)*
- `POST /v1/enterprise/invitations/{invitationId}/resend` — Resend Enterprise User Invitation *(Enterprise Management)*
- `POST /v1/enterprise/settings/tsa` — Add Enterprise TSA Setting *(Enterprise Management)*
- `POST /v1/enterprise/settings/tsa/test-connection` — Test Connection Enterprise TSA Configuration *(Enterprise Management)*
- `POST /v1/enterprise/users/agent-invite` — Agent Invite Enterprise User *(Enterprise Management)*
- `POST /v1/validations` — Add Validation *(Validation Management)*
- `PUT /v1/enterprise/custom_fields/{fieldId}` — Update Enterprise Custom Field *(Enterprise Management)*
- `PUT /v1/enterprise/electronic-seals` — Update Enterprise Electronic Seals *(Enterprise Management)*
- `PUT /v1/enterprise/settings/tsa` — Update Enterprise TSA Configuration *(Enterprise Management)*

### Đã gỡ
- `DELETE /v1/account/invitations` — Reject All Account Invitations *(Account Management)*
- `PUT /v1/account/invitations` — Accept Account Invitations *(Account Management)*

### Đổi method / path
- Delete Enterprise User: `DELETE /v1/enterprise/users` → `DELETE /v1/enterprise/users/{enterpriseUserId}`
- Get User Activity Logs: `GET /v1/account/log/{pageNo}/{recordsPerPage}` → `GET /v1/account/activity-logs/{pageNo}/{recordsPerPage}`
- Get Account Invitations: `GET /v1/account/invitations` → `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`

### Đổi thuộc tính
- Get Templates: +header x-search-text, +header x-total-records, +response field created_at, +response field description, +response field documents, +response field enterprise, +response field read_only, +response field smart_form, +response field template_id, +response field template_name, +response field template_public, +response field templates, -response field Id

### Đổi nội dung chữ (7 mục)
- Account Management / Get Account Invitations
- Account Management / Get User Activity Logs
- Enterprise Management / Delete Enterprise User
- Personal Settings / Get Templates
- RadioboxFieldAttributeV2
- Template / Get Template Details
- TemplateDocumentResponse

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 219 | 265  **+46** |
| Đoạn văn | 42823 | 50738  **+7915** |
| Bảng | 885 | 1033  **+148** |
| Khung mẫu | 1320 | 1598  **+278** |

### Đã thêm
- `DELETE /v1/enterprise/custom_fields/{fieldId}` — Delete Enterprise Custom Field *(Enterprise Management)*
- `DELETE /v1/enterprise/invitations/{invitationId}` — Cancel Enterprise User Invitation *(Enterprise Management)*
- `DELETE /v1/enterprise/invitations/{invitationId}/delete` — Delete Enterprise User Invitation *(Enterprise Management)*
- `DELETE /v1/validations/{validationId}` — Delete Validation *(Validation Management)*
- `GET /v1/account/billings/payments/{billingPaymentId}` — Get Billing Payment Details *(Account Management)*
- `GET /v1/account/billings/{pageNo}/{recordsPerPage}` — Get Billing Payment History *(Account Management)*
- `GET /v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}` — Get Certificate Enroll History *(Account Management)*
- `GET /v1/account/dashboard/activities/{pageNo}/{recordsPerPage}` — Get Dashboard Activities *(Account Management)*
- `GET /v1/account/dashboard/categories` — Get Dashboard Categories *(Account Management)*
- `GET /v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}` — Get Dashboard Overdue Workflows *(Account Management)*
- `GET /v1/account/dashboard/processing-trend` — Get Dashboard Processing Trend *(Account Management)*
- `GET /v1/account/dashboard/quota` — Get Dashboard Quota *(Account Management)*
- `GET /v1/account/dashboard/validations` — Get Dashboard Validations *(Account Management)*
- `GET /v1/account/dashboard/workflow-types` — Get Dashboard Workflow Types *(Account Management)*
- `GET /v1/account/dashboard/workflows` — Get Dashboard Workflows *(Account Management)*
- `GET /v1/account/enterprises` — Get Account Enterprises *(Account Management)*
- `GET /v1/account/invitations/{invitationId}` — Get Account Invitation By ID *(Account Management)*
- `GET /v1/account/profiles` — Get Account Profiles *(Account Management)*
- `GET /v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}` — Get Account Quota Consume History *(Account Management)*
- `GET /v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}` — Get Signing History *(Account Management)*
- `GET /v1/enterprise/custom_fields/{pageNo}/{recordsPerPage}` — Get Enterprise Custom Fields *(Enterprise Management)*
- `GET /v1/enterprise/electronic-seals` — Get Enterprise Electronic Seals *(Enterprise Management)*
- `GET /v1/enterprise/invitations/export/{status}/{pageNo}/{recordsPerPage}` — Export Enterprise Invitations *(Enterprise Management)*
- `GET /v1/enterprise/invitations/{pageNo}/{recordsPerPage}` — Get Enterprise Invitations *(Enterprise Management)*
- `GET /v1/enterprise/reports/quota/statistics/{pageNo}/{recordsPerPage}` — Get Enterprise Quota Statistics *(Enterprise Management)*
- `GET /v1/enterprise/reports/workflows/statistics/export/{status}/{pageNo}/{recordsPerPage}` — Export Enterprise Workflow Statistics *(Enterprise Management)*
- `GET /v1/enterprise/reports/{pageNo}/{recordsPerPage}` — Get Enterprise Report List *(Enterprise Management)*
- `GET /v1/enterprise/settings/tsa` — Get Enterprise TSA Configuration *(Enterprise Management)*
- `GET /v1/enterprise/users/export/{pageNo}/{recordsPerPage}` — Export Enterprise Users *(Enterprise Management)*
- `GET /v1/enterprise/users/statistics` — Get Enterprise User Management Statistics *(Enterprise Management)*
- `GET /v1/enterprise/users/{enterpriseUserId}/quota` — Get Enterprise User Quota *(Enterprise Management)*
- `GET /v1/enterprise/workflows/{status}/{pageNo}/{recordsPerPage}` — Get Enterprise Workflows *(Enterprise Management)*
- `GET /v1/validations/reports/{reportId}/{reportType}` — Download Reports *(Validation Management)*
- `GET /v1/validations/{page_no}/{records_per_page}` — Get List Validation *(Validation Management)*
- `GET /v1/validations/{validationId}/reports` — Get Report of Validation *(Validation Management)*
- `POST /v1/account/billings/initialize` — Initialize Billing *(Account Management)*
- `POST /v1/account/invitations/accept` — Accept Account Invitation *(Account Management)*
- `POST /v1/account/invitations/reject` — Reject Account Invitation *(Account Management)*
- `POST /v1/account/remove` — Remove Account Request *(Personal Settings)*
- `POST /v1/enterprise/custom_fields` — Create Enterprise Custom Field *(Enterprise Management)*
- `POST /v1/enterprise/invitations` — Invite Enterprise User *(Enterprise Management)*
- `POST /v1/enterprise/invitations/{invitationId}/resend` — Resend Enterprise User Invitation *(Enterprise Management)*
- `POST /v1/enterprise/settings/tsa` — Add Enterprise TSA Setting *(Enterprise Management)*
- `POST /v1/enterprise/settings/tsa/test-connection` — Test Connection Enterprise TSA Configuration *(Enterprise Management)*
- `POST /v1/enterprise/users/agent-invite` — Agent Invite Enterprise User *(Enterprise Management)*
- `POST /v1/validations` — Add Validation *(Validation Management)*
- `PUT /v1/enterprise/custom_fields/{fieldId}` — Update Enterprise Custom Field *(Enterprise Management)*
- `PUT /v1/enterprise/electronic-seals` — Update Enterprise Electronic Seals *(Enterprise Management)*
- `PUT /v1/enterprise/settings/tsa` — Update Enterprise TSA Configuration *(Enterprise Management)*

### Đã gỡ
- `DELETE /v1/account/invitations` — Reject All Account Invitations *(Account Management)*
- `PUT /v1/account/invitations` — Accept Account Invitations *(Account Management)*

### Đổi method / path
- Delete Enterprise User: `DELETE /v1/enterprise/users` → `DELETE /v1/enterprise/users/{enterpriseUserId}`
- Get User Activity Logs: `GET /v1/account/log/{pageNo}/{recordsPerPage}` → `GET /v1/account/activity-logs/{pageNo}/{recordsPerPage}`
- Get Account Invitations: `GET /v1/account/invitations` → `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`

### Đổi thuộc tính
- Get Templates: +header x-search-text, +header x-total-records, +response field created_at, +response field description, +response field documents, +response field enterprise, +response field read_only, +response field smart_form, +response field template_id, +response field template_name, +response field template_public, +response field templates, -response field Id

### Đổi nội dung chữ (7 mục)
- Account Management / Get Account Invitations
- Account Management / Get User Activity Logs
- Enterprise Management / Delete Enterprise User
- Personal Settings / Get Templates
- RadioboxFieldAttributeV2
- Template / Get Template Details
- TemplateDocumentResponse

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 219 | 265  **+46** |
| Đoạn văn | 42823 | 50738  **+7915** |
| Bảng | 885 | 1033  **+148** |
| Khung mẫu | 1320 | 1598  **+278** |

### Đã thêm
- `DELETE /v1/enterprise/custom_fields/{fieldId}` — Delete Enterprise Custom Field *(Enterprise Management)*
- `DELETE /v1/enterprise/invitations/{invitationId}` — Cancel Enterprise User Invitation *(Enterprise Management)*
- `DELETE /v1/enterprise/invitations/{invitationId}/delete` — Delete Enterprise User Invitation *(Enterprise Management)*
- `DELETE /v1/validations/{validationId}` — Delete Validation *(Validation Management)*
- `GET /v1/account/billings/payments/{billingPaymentId}` — Get Billing Payment Details *(Account Management)*
- `GET /v1/account/billings/{pageNo}/{recordsPerPage}` — Get Billing Payment History *(Account Management)*
- `GET /v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}` — Get Certificate Enroll History *(Account Management)*
- `GET /v1/account/dashboard/activities/{pageNo}/{recordsPerPage}` — Get Dashboard Activities *(Account Management)*
- `GET /v1/account/dashboard/categories` — Get Dashboard Categories *(Account Management)*
- `GET /v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}` — Get Dashboard Overdue Workflows *(Account Management)*
- `GET /v1/account/dashboard/processing-trend` — Get Dashboard Processing Trend *(Account Management)*
- `GET /v1/account/dashboard/quota` — Get Dashboard Quota *(Account Management)*
- `GET /v1/account/dashboard/validations` — Get Dashboard Validations *(Account Management)*
- `GET /v1/account/dashboard/workflow-types` — Get Dashboard Workflow Types *(Account Management)*
- `GET /v1/account/dashboard/workflows` — Get Dashboard Workflows *(Account Management)*
- `GET /v1/account/enterprises` — Get Account Enterprises *(Account Management)*
- `GET /v1/account/invitations/{invitationId}` — Get Account Invitation By ID *(Account Management)*
- `GET /v1/account/profiles` — Get Account Profiles *(Account Management)*
- `GET /v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}` — Get Account Quota Consume History *(Account Management)*
- `GET /v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}` — Get Signing History *(Account Management)*
- `GET /v1/enterprise/custom_fields/{pageNo}/{recordsPerPage}` — Get Enterprise Custom Fields *(Enterprise Management)*
- `GET /v1/enterprise/electronic-seals` — Get Enterprise Electronic Seals *(Enterprise Management)*
- `GET /v1/enterprise/invitations/export/{status}/{pageNo}/{recordsPerPage}` — Export Enterprise Invitations *(Enterprise Management)*
- `GET /v1/enterprise/invitations/{pageNo}/{recordsPerPage}` — Get Enterprise Invitations *(Enterprise Management)*
- `GET /v1/enterprise/reports/quota/statistics/{pageNo}/{recordsPerPage}` — Get Enterprise Quota Statistics *(Enterprise Management)*
- `GET /v1/enterprise/reports/workflows/statistics/export/{status}/{pageNo}/{recordsPerPage}` — Export Enterprise Workflow Statistics *(Enterprise Management)*
- `GET /v1/enterprise/reports/{pageNo}/{recordsPerPage}` — Get Enterprise Report List *(Enterprise Management)*
- `GET /v1/enterprise/settings/tsa` — Get Enterprise TSA Configuration *(Enterprise Management)*
- `GET /v1/enterprise/users/export/{pageNo}/{recordsPerPage}` — Export Enterprise Users *(Enterprise Management)*
- `GET /v1/enterprise/users/statistics` — Get Enterprise User Management Statistics *(Enterprise Management)*
- `GET /v1/enterprise/users/{enterpriseUserId}/quota` — Get Enterprise User Quota *(Enterprise Management)*
- `GET /v1/enterprise/workflows/{status}/{pageNo}/{recordsPerPage}` — Get Enterprise Workflows *(Enterprise Management)*
- `GET /v1/validations/reports/{reportId}/{reportType}` — Download Reports *(Validation Management)*
- `GET /v1/validations/{page_no}/{records_per_page}` — Get List Validation *(Validation Management)*
- `GET /v1/validations/{validationId}/reports` — Get Report of Validation *(Validation Management)*
- `POST /v1/account/billings/initialize` — Initialize Billing *(Account Management)*
- `POST /v1/account/invitations/accept` — Accept Account Invitation *(Account Management)*
- `POST /v1/account/invitations/reject` — Reject Account Invitation *(Account Management)*
- `POST /v1/account/remove` — Remove Account Request *(Personal Settings)*
- `POST /v1/enterprise/custom_fields` — Create Enterprise Custom Field *(Enterprise Management)*
- `POST /v1/enterprise/invitations` — Invite Enterprise User *(Enterprise Management)*
- `POST /v1/enterprise/invitations/{invitationId}/resend` — Resend Enterprise User Invitation *(Enterprise Management)*
- `POST /v1/enterprise/settings/tsa` — Add Enterprise TSA Setting *(Enterprise Management)*
- `POST /v1/enterprise/settings/tsa/test-connection` — Test Connection Enterprise TSA Configuration *(Enterprise Management)*
- `POST /v1/enterprise/users/agent-invite` — Agent Invite Enterprise User *(Enterprise Management)*
- `POST /v1/validations` — Add Validation *(Validation Management)*
- `PUT /v1/enterprise/custom_fields/{fieldId}` — Update Enterprise Custom Field *(Enterprise Management)*
- `PUT /v1/enterprise/electronic-seals` — Update Enterprise Electronic Seals *(Enterprise Management)*
- `PUT /v1/enterprise/settings/tsa` — Update Enterprise TSA Configuration *(Enterprise Management)*

### Đã gỡ
- `DELETE /v1/account/invitations` — Reject All Account Invitations *(Account Management)*
- `PUT /v1/account/invitations` — Accept Account Invitations *(Account Management)*

### Đổi method / path
- Delete Enterprise User: `DELETE /v1/enterprise/users` → `DELETE /v1/enterprise/users/{enterpriseUserId}`
- Get User Activity Logs: `GET /v1/account/log/{pageNo}/{recordsPerPage}` → `GET /v1/account/activity-logs/{pageNo}/{recordsPerPage}`
- Get Account Invitations: `GET /v1/account/invitations` → `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`

### Đổi thuộc tính
- Get Templates: +header x-search-text, +header x-total-records, +response field created_at, +response field description, +response field documents, +response field enterprise, +response field read_only, +response field smart_form, +response field template_id, +response field template_name, +response field template_public, +response field templates, -response field Id

### Đổi nội dung chữ (7 mục)
- Account Management / Get Account Invitations
- Account Management / Get User Activity Logs
- Enterprise Management / Delete Enterprise User
- Personal Settings / Get Templates
- RadioboxFieldAttributeV2
- Template / Get Template Details
- TemplateDocumentResponse

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 208 | 237  **+29** |
| Đoạn văn | 40798 | 45937  **+5139** |
| Bảng | 844 | 936  **+92** |
| Khung mẫu | 1254 | 1428  **+174** |

### Đã thêm
- `DELETE /v1/templates/{id}/workflow/{order}` — Delete Template Workflow Participant *(Template)*
- `GET /v1/account/billings/payments/{billingPaymentId}` — Get Billing Payment Details *(Account Management)*
- `GET /v1/account/billings/{pageNo}/{recordsPerPage}` — Get Billing Payment History *(Account Management)*
- `GET /v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}` — Get Certificate Enroll History *(Account Management)*
- `GET /v1/account/dashboard/activities/{pageNo}/{recordsPerPage}` — Get Dashboard Activities *(Account Management)*
- `GET /v1/account/dashboard/categories` — Get Dashboard Categories *(Account Management)*
- `GET /v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}` — Get Dashboard Overdue Workflows *(Account Management)*
- `GET /v1/account/dashboard/processing-trend` — Get Dashboard Processing Trend *(Account Management)*
- `GET /v1/account/dashboard/quota` — Get Dashboard Quota *(Account Management)*
- `GET /v1/account/dashboard/validations` — Get Dashboard Validations *(Account Management)*
- `GET /v1/account/dashboard/workflow-types` — Get Dashboard Workflow Types *(Account Management)*
- `GET /v1/account/dashboard/workflows` — Get Dashboard Workflows *(Account Management)*
- `GET /v1/account/enterprises` — Get Account Enterprises *(Account Management)*
- `GET /v1/account/invitations/{invitationId}` — Get Account Invitation By ID *(Account Management)*
- `GET /v1/account/profiles` — Get Account Profiles *(Account Management)*
- `GET /v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}` — Get Account Quota Consume History *(Account Management)*
- `GET /v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}` — Get Signing History *(Account Management)*
- `GET /v1/templates/{id}/details` — Get Template Details *(Template)*
- `GET /v1/templates/{id}/workflow/users` — Get Template Workflow Users *(Template)*
- `POST /v1/account/billings/initialize` — Initialize Billing *(Account Management)*
- `POST /v1/account/invitations/accept` — Accept Account Invitation *(Account Management)*
- `POST /v1/account/invitations/reject` — Reject Account Invitation *(Account Management)*
- `POST /v1/account/remove` — Remove Account Request *(Personal Settings)*
- `POST /v1/templates/{id}/clone` — Clone Template *(Template)*
- `POST /v1/templates/{id}/documents/{documentId}/fields/{fieldType}` — Add Template Field *(Template)*
- `POST /v1/templates/{id}/workflow/placeholder` — Add Template Workflow Placeholder *(Template)*
- `POST /v1/templates/{id}/workflow/users` — Add Template Workflow Users *(Template)*
- `PUT /v1/templates/{id}/documents/{documentId}/fields/{fieldType}` — Update Template Field *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/placeholder` — Update Template Workflow Placeholder *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/reorder` — Reorder Template Workflow Participant *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/user` — Update Template Workflow User *(Template)*

### Đã gỡ
- `DELETE /v1/account/invitations` — Reject All Account Invitations *(Account Management)*
- `PUT /v1/account/invitations` — Accept Account Invitations *(Account Management)*

### Đổi method / path
- Get User Activity Logs: `GET /v1/account/log/{pageNo}/{recordsPerPage}` → `GET /v1/account/activity-logs/{pageNo}/{recordsPerPage}`
- Get Account Invitations: `GET /v1/account/invitations` → `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`
- Get Group: `GET /v1/settings/groups/{recordPerPage}/{pageNo}` → `GET /v1/settings/groups/{pageNo}/{recordPerPage}`
- Get Templates: `GET /v1/settings/templates/{recordPerPage}/{pageNo}` → `GET /v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}`

### Đổi nội dung chữ (4 mục)
- Account Management / Get Account Invitations
- Account Management / Get Group
- Account Management / Get User Activity Logs
- Enterprise Management / Get Groups

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 208 | 237  **+29** |
| Đoạn văn | 40805 | 45884  **+5079** |
| Bảng | 844 | 936  **+92** |
| Khung mẫu | 1254 | 1428  **+174** |

### Đã thêm
- `DELETE /v1/templates/{id}/workflow/{order}` — Delete Template Workflow Participant *(Template)*
- `GET /v1/account/billings/payments/{billingPaymentId}` — Get Billing Payment Details *(Account Management)*
- `GET /v1/account/billings/{pageNo}/{recordsPerPage}` — Get Billing Payment History *(Account Management)*
- `GET /v1/account/certificate-enrollments/{status}/{pageNo}/{recordsPerPage}` — Get Certificate Enroll History *(Account Management)*
- `GET /v1/account/dashboard/activities/{pageNo}/{recordsPerPage}` — Get Dashboard Activities *(Account Management)*
- `GET /v1/account/dashboard/categories` — Get Dashboard Categories *(Account Management)*
- `GET /v1/account/dashboard/overdue-workflows/{pageNo}/{recordsPerPage}` — Get Dashboard Overdue Workflows *(Account Management)*
- `GET /v1/account/dashboard/processing-trend` — Get Dashboard Processing Trend *(Account Management)*
- `GET /v1/account/dashboard/quota` — Get Dashboard Quota *(Account Management)*
- `GET /v1/account/dashboard/validations` — Get Dashboard Validations *(Account Management)*
- `GET /v1/account/dashboard/workflow-types` — Get Dashboard Workflow Types *(Account Management)*
- `GET /v1/account/dashboard/workflows` — Get Dashboard Workflows *(Account Management)*
- `GET /v1/account/enterprises` — Get Account Enterprises *(Account Management)*
- `GET /v1/account/invitations/{invitationId}` — Get Account Invitation By ID *(Account Management)*
- `GET /v1/account/profiles` — Get Account Profiles *(Account Management)*
- `GET /v1/account/reports/quota/consume-history/{pageNo}/{recordsPerPage}` — Get Account Quota Consume History *(Account Management)*
- `GET /v1/account/signing-transactions/{status}/{pageNo}/{recordsPerPage}` — Get Signing History *(Account Management)*
- `GET /v1/templates/{id}/details` — Get Template Details *(Template)*
- `GET /v1/templates/{id}/workflow/users` — Get Template Workflow Users *(Template)*
- `POST /v1/account/billings/initialize` — Initialize Billing *(Account Management)*
- `POST /v1/account/invitations/accept` — Accept Account Invitation *(Account Management)*
- `POST /v1/account/invitations/reject` — Reject Account Invitation *(Account Management)*
- `POST /v1/account/remove` — Remove Account Request *(Personal Settings)*
- `POST /v1/templates/{id}/clone` — Clone Template *(Template)*
- `POST /v1/templates/{id}/documents/{documentId}/fields/{fieldType}` — Add Template Field *(Template)*
- `POST /v1/templates/{id}/workflow/placeholder` — Add Template Workflow Placeholder *(Template)*
- `POST /v1/templates/{id}/workflow/users` — Add Template Workflow Users *(Template)*
- `PUT /v1/templates/{id}/documents/{documentId}/fields/{fieldType}` — Update Template Field *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/placeholder` — Update Template Workflow Placeholder *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/reorder` — Reorder Template Workflow Participant *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/user` — Update Template Workflow User *(Template)*

### Đã gỡ
- `DELETE /v1/account/invitations` — Reject All Account Invitations *(Account Management)*
- `PUT /v1/account/invitations` — Accept Account Invitations *(Account Management)*

### Đổi method / path
- Get User Activity Logs: `GET /v1/account/log/{pageNo}/{recordsPerPage}` → `GET /v1/account/activity-logs/{pageNo}/{recordsPerPage}`
- Get Account Invitations: `GET /v1/account/invitations` → `GET /v1/account/invitations/{pageNo}/{recordsPerPage}`
- Get Group: `GET /v1/settings/groups/{recordPerPage}/{pageNo}` → `GET /v1/settings/groups/{pageNo}/{recordPerPage}`
- Get Templates: `GET /v1/settings/templates/{recordPerPage}/{pageNo}` → `GET /v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}`

### Đổi thuộc tính
- Get Information of Citizen card: -request field document_type

### Đổi nội dung chữ (7 mục)
- Account Management / Get Account Invitations
- Account Management / Get Group
- Account Management / Get User Activity Logs
- Enterprise Management / Get Groups
- Personal Settings / Get Templates
- PersonalInformation
- Register Certificate Processing / Get Information of Citizen card

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 208 | 219  **+11** |
| Đoạn văn | 40793 | 42823  **+2030** |
| Bảng | 844 | 885  **+41** |
| Khung mẫu | 1254 | 1320  **+66** |

### Đã thêm
- `DELETE /v1/templates/{id}/workflow/{order}` — Delete Template Workflow Participant *(Template)*
- `GET /v1/templates/{id}/details` — Get Template Details *(Template)*
- `GET /v1/templates/{id}/workflow/users` — Get Template Workflow Users *(Template)*
- `POST /v1/templates/{id}/clone` — Clone Template *(Template)*
- `POST /v1/templates/{id}/documents/{documentId}/fields/{fieldType}` — Add Template Field *(Template)*
- `POST /v1/templates/{id}/workflow/placeholder` — Add Template Workflow Placeholder *(Template)*
- `POST /v1/templates/{id}/workflow/users` — Add Template Workflow Users *(Template)*
- `PUT /v1/templates/{id}/documents/{documentId}/fields/{fieldType}` — Update Template Field *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/placeholder` — Update Template Workflow Placeholder *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/reorder` — Reorder Template Workflow Participant *(Template)*
- `PUT /v1/templates/{id}/workflow/{order}/user` — Update Template Workflow User *(Template)*

### Đổi method / path
- Get List Tax Codes: `GET /v1/csc/credentials/{session_id}/taxcode` → `GET /v1/csc/credentials/{sessionId}/taxcode`
- Get Group: `GET /v1/settings/groups/{recordPerPage}/{pageNo}` → `GET /v1/settings/groups/{pageNo}/{recordPerPage}`
- Get Templates: `GET /v1/settings/templates/{recordPerPage}/{pageNo}` → `GET /v1/settings/templates/{templateType}/{pageNo}/{recordPerPage}`
- Finish Register Certificate: `POST /v1/csc/credentials/{session_id}/finish` → `POST /v1/csc/credentials/{sessionId}/finish`
- Send OTP for Certificate Registration: `POST /v1/csc/credentials/{session_id}/send-otp` → `POST /v1/csc/credentials/{sessionId}/send-otp`
- Verify OTP for Certificate Registration: `POST /v1/csc/credentials/{session_id}/verify` → `POST /v1/csc/credentials/{sessionId}/verify`

### Đổi thuộc tính
- Get Status of Import Excel: +response field workflows, -response field results
- Get Information of Citizen card: -request field document_type

### Đổi nội dung chữ (12 mục)
- Account Management / Get Group
- Document Workflow / Create Workflow From Template
- Document Workflow / Get Status of Import Excel
- Enterprise Management / Get Groups
- Personal Settings / Get Templates
- PersonalInformation
- Register Certificate Processing / Finish Register Certificate
- Register Certificate Processing / Get Credential Registration Status
- Register Certificate Processing / Get Information of Citizen card
- Register Certificate Processing / Get List Tax Codes
- Register Certificate Processing / Send OTP for Certificate Registration
- Register Certificate Processing / Verify OTP for Certificate Registration

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 208 | 208 |
| Đoạn văn | 40805 | 40798  **-7** |
| Bảng | 844 | 844 |
| Khung mẫu | 1254 | 1254 |

### Đổi thuộc tính
- Get Information of Citizen card: -request field document_type

### Đổi nội dung chữ (2 mục)
- PersonalInformation
- Register Certificate Processing / Get Information of Citizen card

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 208 | 208 |
| Đoạn văn | 40793 | 40805  **+12** |
| Bảng | 844 | 844 |
| Khung mẫu | 1254 | 1254 |

### Đổi method / path
- Get List Tax Codes: `GET /v1/csc/credentials/{session_id}/taxcode` → `GET /v1/csc/credentials/{sessionId}/taxcode`
- Finish Register Certificate: `POST /v1/csc/credentials/{session_id}/finish` → `POST /v1/csc/credentials/{sessionId}/finish`
- Send OTP for Certificate Registration: `POST /v1/csc/credentials/{session_id}/send-otp` → `POST /v1/csc/credentials/{sessionId}/send-otp`
- Verify OTP for Certificate Registration: `POST /v1/csc/credentials/{session_id}/verify` → `POST /v1/csc/credentials/{sessionId}/verify`

### Đổi nội dung chữ (5 mục)
- Register Certificate Processing / Finish Register Certificate
- Register Certificate Processing / Get Credential Registration Status
- Register Certificate Processing / Get List Tax Codes
- Register Certificate Processing / Send OTP for Certificate Registration
- Register Certificate Processing / Verify OTP for Certificate Registration

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 206 | 208  **+2** |
| Đoạn văn | 40556 | 40793  **+237** |
| Bảng | 839 | 844  **+5** |
| Khung mẫu | 1242 | 1254  **+12** |

### Đã thêm
- `GET /v1/csc/credentials/{transactionCode}/status` — Get Credential Registration Status *(Register Certificate Processing)*
- `POST /v1/ras/issue/cancel` — Cancel Issue Certificate *(Register Certificate Processing)*

### Đổi thuộc tính
- Get List Tax Codes: +response field tax_informations, -request field otp, -request field user_email
- Register Certificate via RSSP: +request field document_extraction_connector, +request field enrollment_connector, +request field metadata, +request field mobile_number, -request field connector_name, request field face_image: M -> O, request field signing_server: RC(When the connector_name is have the same name in DB) -> RC(When the enrollment_connector is have the same name in DB)
- Finish Register Certificate: +request field auth_mode, +request field metadata, +request field operation_mode, +request field webhook_url
- Send OTP for Certificate Registration: +request field email, +request field mobile, -request field otp, -request field user_email
- Verify OTP for Certificate Registration: -request field user_email

### Đổi nội dung chữ (6 mục)
- Document Workflow / Create Workflow From Template
- Document Workflow / Get Status of Import Excel
- Register Certificate Processing / Finish Register Certificate
- Register Certificate Processing / Get List Tax Codes
- Register Certificate Processing / Register Certificate via RSSP
- Register Certificate Processing / Verify OTP for Certificate Registration

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 206 | 206 |
| Đoạn văn | 40545 | 40547  **+2** |
| Bảng | 839 | 839 |
| Khung mẫu | 1242 | 1242 |

### Đổi nội dung chữ (1 mục)
- Document Processing / Fill Form Fields

## 2026-07-27 — Document Preparation, Document Workflow, Document Package

| | trước | sau |
|---|---|---|
| Endpoint | 188 | 206  **+18** |
| Đoạn văn | 36237 | 40545  **+4308** |
| Bảng | 768 | 839  **+71** |
| Khung mẫu | 1126 | 1242  **+116** |

### Đã thêm
- `DELETE /v1/category/{categoryId}` — Delete Category *(Document Workflow)*
- `DELETE /v1/packages/{packageId}/workflow` — Delete Workflow *(Document Workflow)*
- `GET /v1/csc/signing/pending-sessions` — Get Pending Signing Sessions *(RSSP Processing)*
- `GET /v1/csc/signing/{transactionCode}/status` — Get Signing Transaction Status *(RSSP Processing)*
- `GET /v1/documents/fields` — Get Field Types *(Document Preparation)*
- `GET /v1/workflow/{transactionCode}/import` — Get Status of Import Excel *(Document Workflow)*
- `GET /v1/workflow/{workflowUuid}` — Get Workflow Details by UUID *(Document Workflow)*
- `POST /v1/csc/credentials/ocsp-check` — Check OCSP Certificate Status *(RSSP Processing)*
- `POST /v1/csc/signatures/sign-hash` — Get Sign Hash From Service *(RSSP Processing)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/elabor` — Add Elabor Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/redaction` — Add Redact Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/vneid` — Add VNeID Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/workflow/{order}/import` — Import Excel to Clone Workflow *(Document Workflow)*
- `POST /v1/packages/{packageId}/workflow/{order}/remind` — Send Reminder *(Document Workflow)*
- `POST /v1/workflow/from-template` — Create Workflow From Template *(Document Workflow)*
- `POST /v1/workflow/{uuid}/view` — View Workflow *(Document Workflow)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/elabor` — Update Elabor Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/redaction` — Update Redact Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/vneid` — Update VNeID Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/workflow/{participantId}/tos-acceptance` — Accept Terms of Service (ECONTRACT) *(Document Workflow)*

### Đã gỡ
- `GET /v1/packages/{packageId}/documents/{documentId}/fields/signature/status` — Check status of Signature Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/qrypto` — Add Qrypto Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/qryto` — Update Qrypto Field *(Document Preparation)*

### Đổi method / path
- Get Workflow History: `GET /v1/packages/{packageId}/log` → `GET /v1/packages/{packageId}/log/{page_no}/{records_per_page}`
- Assign Document Field: `PUT /v4/packages/{packageId}/documents/{documentId}/fields/assign` → `PUT /v1/packages/{packageId}/documents/{documentId}/fields/assign`
- Update Placeholder in Workflow: `POST /v1/packages/{packageId}/workflow/{order}/placeholder` → `PUT /v1/packages/{packageId}/workflow/{order}/placeholder`

### Đổi thuộc tính
- **2 endpoint**: +response field document_id, -response field documentid
  - Upload Document, Upload Document (base64)
- Get Document Details: +response field form_field, -response field certify, -response field document_source, -response field form_fields, -response field template
- Get Document Fields: -response field qrypto
- RSSP Credential Authorization: +request field face_image, +request field responseID
- Get Credentials Info: +response field fields, request field connector_name: M -> RC (When credential_id is provided)
- Get Filtered Credential List: +header x-credential-id, -header x-search-server
- Send OTP via RSSP: +request field notification_template, +response field user_email, +response field validity_period
- Sign Document via RSSP Directly: +request field citizen_id, +request field contact_information, +request field is_redirect_app_enabled, +request field metadata, +request field notification_template, +request field rssp_operation_mode, +request field signing_server, +request field webhook_url, -request field signer_contact, -request field time_zone, request field hashing_algo: M -> O, request field signing_algo: M -> O
- Approve Document: -header x-otp, -header x-password, request field reason: M -> O
- Get Document Hash: +request field appearance_design, +response field hash_value, -request field signing_order, -request field skip_verification, -request field time_zone, -response field hash, -response field hashing_algo, -response field state, request field certificate_chain: M -> O
- Fill Initials: +header x-handle-transaction, +request field initial_field, -request field initial_pages, request field field_name: M -> O, request field image: M -> RC (Required when the initial field has no stored image)
- Sign Document: -header x-otp, -request field appearance_design, -request field authentication, -request field connector_name, -request field signing_capacity, -request field signing_server
- Sign Eseal: +request field contact_information, +request field hand_signature_image, +request field signing_location, -header x-otp
- Apply Workflow Template: -header x-append-to-top, -response field certify, -response field created_at, -response field created_by, -response field document_height, -response field document_id, -response field document_name, -response field document_order, -response field document_pages, -response field document_source, -response field document_type, -response field document_width, -response field form_fields, -response field lock_form_fields, -response field modified_at, -response field modified_by, -response field template
- Submit Document: +request field reason, -header x-otp, -header x-password
- Start new Workflow: -request field filter_level
- Fill Form Fields: +header x-handle-transaction, +request field checkboxv2, +request field combobox, +request field custom_field, +request field datetime, +request field qr, +request field radiobox, +request field radioboxv2, +request field textbox, +request field vneid_field, -request field checkboxV2, -request field combo, -request field radio, -request field radioboxV2, -request field text
- Update Workflow Details: -request field workflow_type
- Update Workflow Reminders: +request field apply_to_all

### Đổi nội dung chữ (74 mục)
- Account Management / Get Account Invitations
- Application/Device Management / Register Application User
- Authenticate – Log in / Get Information
- Connector / Get Certificate Profile of Connector
- Connector / Get List Internal Signing Service Connector (ISS)
- Connector / Get List Smart-ID Connector
- Connector / Get List USB-Token Connector
- CredentialInfo
- Document Package / Apply Workflow Template
- Document Package / Download Document
- Document Package / Download Document (base64)
- Document Package / Download Package
- Document Package / Download Signature Summary
- Document Package / Get All Packages
- Document Package / Get Document Details
- Document Package / Get Document Image
- Document Package / Get Document Image Base64
- Document Package / Get Document Verification
- Document Package / Get Package Details
- Document Package / Get Package Verification
- Document Package / Merge Document
- Document Package / Pre Upload Document
- Document Package / Share Document Package/ Share Workflow
- Document Package / Update Document
- Document Package / Update Document (base64)
- … và 49 mục khác

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 202 | 202 |
| Đoạn văn | 38746 | 39660  **+914** |
| Bảng | 818 | 828  **+10** |
| Khung mẫu | 1210 | 1210 |

### Đổi thuộc tính
- Approve Document: -header x-otp, -header x-password, request field reason: M -> O
- Get Document Hash: +request field appearance_design, +response field hash_value, -request field signing_order, -request field skip_verification, -request field time_zone, -response field hash, -response field hashing_algo, -response field state, request field certificate_chain: M -> O
- Fill Initials: +header x-handle-transaction, +request field initial_field, -request field initial_pages, request field field_name: M -> O, request field image: M -> RC (Required when the initial field has no stored image)
- Sign Document: -header x-otp, -request field appearance_design, -request field authentication, -request field connector_name, -request field signing_capacity, -request field signing_server
- Sign Eseal: +request field contact_information, +request field hand_signature_image, +request field signing_location, -header x-otp
- Submit Document: +request field reason, -header x-otp, -header x-password
- Fill Form Fields: +header x-handle-transaction, +request field checkboxv2, +request field combobox, +request field custom_field, +request field datetime, +request field qr, +request field radiobox, +request field radioboxv2, +request field textbox, +request field vneid_field, -request field checkboxV2, -request field combo, -request field radio, -request field radioboxV2, -request field text

### Đổi nội dung chữ (13 mục)
- Document Preparation / Add Redact Field
- Document Preparation / Update Redact Field
- Document Processing / Approve Document
- Document Processing / Decline Document
- Document Processing / Fill Form Fields
- Document Processing / Fill Initials
- Document Processing / Get Document Hash
- Document Processing / Sign Document
- Document Processing / Sign Eseal
- Document Processing / Submit Document
- Document Workflow / Get Workflow Users
- FieldTypeItem
- FontName

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 188 | 202  **+14** |
| Đoạn văn | 36237 | 38652  **+2415** |
| Bảng | 768 | 814  **+46** |
| Khung mẫu | 1126 | 1210  **+84** |

### Đã thêm
- `DELETE /v1/category/{categoryId}` — Delete Category *(Document Workflow)*
- `DELETE /v1/packages/{packageId}/workflow` — Delete Workflow *(Document Workflow)*
- `GET /v1/documents/fields` — Get Field Types *(Document Preparation)*
- `GET /v1/workflow/{transactionCode}/import` — Get Status of Import Excel *(Document Workflow)*
- `GET /v1/workflow/{workflowUuid}` — Get Workflow Details by UUID *(Document Workflow)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/elabor` — Add Elabor Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/redaction` — Add Redact Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/vneid` — Add VNeID Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/workflow/{order}/import` — Import Excel to Clone Workflow *(Document Workflow)*
- `POST /v1/packages/{packageId}/workflow/{order}/remind` — Send Reminder *(Document Workflow)*
- `POST /v1/workflow/from-template` — Create Workflow From Template *(Document Workflow)*
- `POST /v1/workflow/{uuid}/view` — View Workflow *(Document Workflow)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/elabor` — Update Elabor Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/redaction` — Update Redact Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/vneid` — Update VNeID Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/workflow/{participantId}/tos-acceptance` — Accept Terms of Service (ECONTRACT) *(Document Workflow)*

### Đã gỡ
- `GET /v1/packages/{packageId}/documents/{documentId}/fields/signature/status` — Check status of Signature Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/qrypto` — Add Qrypto Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/qryto` — Update Qrypto Field *(Document Preparation)*

### Đổi method / path
- Get Workflow History: `GET /v1/packages/{packageId}/log` → `GET /v1/packages/{packageId}/log/{page_no}/{records_per_page}`
- Assign Document Field: `PUT /v4/packages/{packageId}/documents/{documentId}/fields/assign` → `PUT /v1/packages/{packageId}/documents/{documentId}/fields/assign`
- Update Placeholder in Workflow: `POST /v1/packages/{packageId}/workflow/{order}/placeholder` → `PUT /v1/packages/{packageId}/workflow/{order}/placeholder`

### Đổi thuộc tính
- **2 endpoint**: +response field document_id, -response field documentid
  - Upload Document, Upload Document (base64)
- Get Document Details: +response field form_field, -response field certify, -response field document_source, -response field form_fields, -response field template
- Get Document Fields: -response field qrypto
- Apply Workflow Template: -header x-append-to-top, -response field certify, -response field created_at, -response field created_by, -response field document_height, -response field document_id, -response field document_name, -response field document_order, -response field document_pages, -response field document_source, -response field document_type, -response field document_width, -response field form_fields, -response field lock_form_fields, -response field modified_at, -response field modified_by, -response field template
- Start new Workflow: -request field filter_level
- Update Workflow Details: -request field workflow_type
- Update Workflow Reminders: +request field apply_to_all

### Đổi nội dung chữ (27 mục)
- Document Package / Apply Workflow Template
- Document Package / Download Document (base64)
- Document Package / Get All Packages
- Document Package / Get Document Details
- Document Package / Get Document Image
- Document Package / Get Document Image Base64
- Document Package / Get Package Details
- Document Package / Merge Document
- Document Package / Share Document Package/ Share Workflow
- Document Package / Update Document
- Document Package / Update Document (base64)
- Document Package / Upload Document
- Document Package / Upload Document (base64)
- Document Preparation / Add TextBox Field
- Document Preparation / Assign Document Field
- Document Preparation / Get Document Fields
- Document Preparation / Update TextBox Field
- Document Workflow / Add User to Workflow
- Document Workflow / Get Workflow History
- Document Workflow / Get Workflow User Authentication (Document Opening)
- Document Workflow / Get Workflow User Permissions
- Document Workflow / Start new Workflow
- Document Workflow / Update Placeholder in Workflow
- Document Workflow / Update Workflow Details
- Document Workflow / Update Workflow Reminders
- … và 2 mục khác

## 2026-07-27

| | trước | sau |
|---|---|---|
| Endpoint | 188 | 203  **+15** |
| Đoạn văn | 36237 | 39116  **+2879** |
| Bảng | 768 | 820  **+52** |
| Khung mẫu | 1126 | 1216  **+90** |

### Đã thêm
- `DELETE /v1/category/{categoryId}` — Delete Category *(Document Workflow)*
- `DELETE /v1/packages/{packageId}/workflow` — Delete Workflow *(Document Workflow)*
- `GET /v1/workflow/{transactionCode}/import` — Get Status of Import Excel *(Document Workflow)*
- `GET /v1/workflow/{workflowUuid}` — Get Workflow Details by UUID *(Document Workflow)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/elabor` — Add Elabor Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/redaction` — Add Redact Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/documents/{documentId}/fields/vneid` — Add VNeID Field *(Document Preparation)*
- `POST /v1/packages/{packageId}/workflow/{order}/import` — Import Excel to Clone Workflow *(Document Workflow)*
- `POST /v1/packages/{packageId}/workflow/{order}/remind` — Send Reminder *(Document Workflow)*
- `POST /v1/workflow/from-template` — Create Workflow From Template *(Document Workflow)*
- `POST /v1/workflow/{uuid}/view` — View Workflow *(Document Workflow)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/elabor` — Update Elabor Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/redaction` — Update Redact Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/documents/{documentId}/fields/vneid` — Update VNeID Field *(Document Preparation)*
- `PUT /v1/packages/{packageId}/workflow/{participantId}/tos-acceptance` — Accept Terms of Service (ECONTRACT) *(Document Workflow)*

### Đã gỡ
- `GET /v1/packages/{packageId}/documents/{documentId}/fields/signature/status` — Check status of Signature Field *(Document Preparation)*

### Đổi method / path
- Get Workflow History: `GET /v1/packages/{packageId}/log` → `GET /v1/packages/{packageId}/log/{page_no}/{records_per_page}`
- Assign Document Field: `PUT /v4/packages/{packageId}/documents/{documentId}/fields/assign` → `PUT /v1/packages/{packageId}/documents/{documentId}/fields/assign`
- Update Qrypto Field: `PUT /v1/packages/{packageId}/documents/{documentId}/fields/qryto` → `PUT /v1/packages/{packageId}/documents/{documentId}/fields/qrypto`
- Update Placeholder in Workflow: `POST /v1/packages/{packageId}/workflow/{order}/placeholder` → `PUT /v1/packages/{packageId}/workflow/{order}/placeholder`

### Đổi thuộc tính
- **2 endpoint**: +response field document_id, -response field documentid
  - Upload Document, Upload Document (base64)
- Apply Workflow Template: -header x-append-to-top

### Đổi nội dung chữ (10 mục)
- Document Package / Apply Workflow Template
- Document Package / Get Document Image
- Document Package / Get Document Image Base64
- Document Package / Share Document Package/ Share Workflow
- Document Package / Upload Document
- Document Package / Upload Document (base64)
- Document Preparation / Assign Document Field
- Document Preparation / Update Qrypto Field
- Document Workflow / Get Workflow History
- Document Workflow / Update Placeholder in Workflow

## 2026-07-27 — Rà soát và đồng bộ tài liệu với backend

| | trước | sau |
|---|---|---|
| Endpoint | 185 | 186  **+1** |
| Đoạn văn | 39197 | 36015  **-3182** |
| Bảng | 759 | 762  **+3** |
| Khung mẫu | 1108 | 1114  **+6** |

### Đã thêm
- `POST /v1/account/switch` — Switch Enterprise Context *(Authenticate – Log in)*
- `POST /v1/packages/{packageId}/documents/{documentId}/paper` — Convert Document to Paper *(Document Package)*

### Đổi method / path
- Pre Login Authentication: `GET /v1/authenticate/pre` → `POST /v1/authenticate/pre`
- Upload Attachment: `PUT /v1/packages/{packageId}/documents/{documentId}/attachments` → `POST /v1/packages/{packageId}/documents/{documentId}/attachments`

### Đổi thuộc tính
- **173 endpoint**: +response field error_code, +response field message, -response field error, -response field remark
  - Reject All Account Invitations, Delete Member in Group, Delete Role, … (+170)
- **2 endpoint**: +response field error_code, +response field message, -header x-file-name, -response field error, -response field remark
  - Delete Attachment, Get Attachments
- Delete Address Book: +header Authorization, +response field error_code, +response field message, -response field error, -response field remark
- Download Attachment: +response field error_code, +response field message, -header x-file-name, -response field attachment_id, -response field error, -response field remark
- Add Identity for a User: +response field error_code, -response field error, -response field remark

### Đổi nội dung chữ (197 mục)
- ADMIN APIs / Admin Authentication
- ADMIN APIs / Change User Service Plan
- ADMIN APIs / Create Account
- ADMIN APIs / Get Accounts
- ADMIN APIs / Service Plan Reset
- Account Management / Accept Account Invitations
- Account Management / Add Identity for a User
- Account Management / Change Password
- Account Management / Get Account
- Account Management / Get Account Invitations
- Account Management / Get Group
- Account Management / Get User Activity Logs
- Account Management / Register Enterprise User (Free Trial)
- Account Management / Reject All Account Invitations
- Account Management / Send Forgot Password Request
- Account Management / Update Security Settings
- AddUserToWorkflowRequest
- Application/Device Management / Add Personal Contact
- Application/Device Management / Delete Personal Contact
- Application/Device Management / Get Contacts
- Application/Device Management / Register Application User
- Application/Device Management / Update Personal Contact
- AssignDocumentRequest
- Authenticate – Log in / Client Credentials Authentication
- Authenticate – Log in / Get Information
- … và 172 mục khác

