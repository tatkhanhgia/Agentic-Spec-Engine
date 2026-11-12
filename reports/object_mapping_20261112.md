# Đối chiếu object V3 với schema trong `openapi.yaml` (2026-11-12)

Dùng cho bước "sinh chương object và enumeration từ OpenAPI". Cách đối chiếu: một ô Type trong V3 link tới mục
object nào, và field đó trong spec trỏ tới schema nào. Chỉ đối chiếu được những mục có link từ một section API
đã ghép được với spec (trực tiếp hoặc qua object lồng nhau).

| | Số mục V3 | Đối chiếu được | Chưa đối chiếu được |
|---|---|---|---|
| REQUEST/RESPONSE OBJECT | 16 | 6 | 10 |
| RELATIVE OBJECT | 106 | 33 | 73 |
| ENUMERATION | 18 | 0 | 18 |

Spec hiện có 419 schema có thuộc tính; chưa có enum nào được khai báo thành schema có tên.

## REQUEST/RESPONSE OBJECT

### Đối chiếu được

| Mục trong V3 | Schema trong spec | Ghi chú |
|---|---|---|
| AddUserToWorkflowRequest | `AddUserToWorkflowRequestData` |  |
| AssignDocumentRequest | `AssignDocumentFieldRequestAssignElement` |  |
| GetListWorkflowResponse | `WorkflowListItem` |  |
| AddPlaceholderToWorkflowRequest | `AddPlaceholderRequestPlaceholder` |  |
| GetAttachmentsResponse | `AttachmentListItem` |  |
| TemplateDocumentResponse | `TemplateDocumentSummary` |  |

### Chưa đối chiếu được (cần owner chỉ schema, hoặc xác nhận bỏ)

- AddUserToWorkflowResponse
- GetWorkflowHistoryResponse
- VerificationSignatureRequest
- RegisterUserRequest
- UpdateWorkflowPermissionRequest
- GetFilteredCredentialListResponse
- BaseFillFormRequest
- AddPlaceholderToWorkflowResponse
- GetCertificateProfilesResponse
- TemplateDocumentResponse(Get Template Details)

## RELATIVE OBJECT

### Đối chiếu được

| Mục trong V3 | Schema trong spec | Ghi chú |
|---|---|---|
| Package | `PackageListItem` |  |
| Access_duration_1 (Get Packages request) | `PackageAccessDuration` |  |
| Decline | `PackageDecline` |  |
| User | `WorkflowParticipantItem` |  |
| Action | `WorkflowHistoryAction` |  |
| Comment/Message | `WorkflowCommentItem` |  |
| Dimension | `FieldDimension` | cũng được link từ field trỏ tới: `TemplateFieldDimension` |
| Font | `FieldFont` |  |
| AlignmentOption | `FieldAlignment` |  |
| Owner(Package Details) | `PackageDetailsOwner` |  |
| Document(Package Details) | `PackageDetailsDocument` |  |
| Enterprise | `ProfileEnterprise` |  |
| Consumer | `UsageMetric` |  |
| General | `ProfileGeneral` |  |
| Contact | `AddContactRequest` |  |
| Category | `GetListCategoriesResponse` |  |
| SignatureAppearance | `SignatureAppearanceItem` |  |
| MetadataSignature | `SignatureAppearanceMetadata` |  |
| CredentialConfiguration | `SignatureCredentialConfiguration` |  |
| ChooseItems | `FieldComboItems` |  |
| ChooseItem | `FieldComboItem` |  |
| Date | `FieldDateOptions` |  |
| File | `FieldAttachmentFile` |  |
| Frame | `FieldCheckboxFrame` |  |
| Image | `FieldCameraImage` |  |
| RedactPolicy | `FieldRedactPolicy` |  |
| RedactionRule | `FieldRedactPolicyRule` |  |
| FieldTypeGroup | `FieldTypeGroup` |  |
| Validation | `ValidationListItem` |  |
| Overview(Validation) | `ValidationOverview` |  |
| Signature(Validation) | `ValidationSignature` |  |
| Details(Validation) | `ValidationDetails` |  |
| File(Validation) | `ValidationFile` |  |

### Chưa đối chiếu được (cần owner chỉ schema, hoặc xác nhận bỏ)

- NextSigner
- Certify
- DocumentVerification
- PackageVerification
- Document
- Workflow
- Reminder
- Information
- Sender
- Authentication (for Workflow User Authentication)
- Password
- Sms
- Authentication_Signing (for Workflow User Authentication)
- Access_duration (for Workflow User Authentication)
- DurationByDate
- DurationByDays
- DurationDays
- DurationDate
- Permission
- FieldAuthentication
- Otp
- SignatureFieldAttribute
- VerificationData
- TextFieldAttribute
- CheckboxFieldAttribute
- QRFieldAttribute
- InitialFieldAttribute
- RSSPService
- Authentication(SigningRequest)
- Authentication (Workflow User Authentication)
- UserOfEnterprise
- Group
- Member
- Role
- Invitation
- ServicePlan
- Constraints
- Setting
- Billing
- UserActivity
- Information
- Security
- Locale
- Account
- Connector
- Address Book
- WorkflowDetail
- AssignedField
- CredentialInfo
- PersonalInformation
- NumericItems
- Template
- Template(Get Templates)
- CertificateProfile
- IssuingProvider
- FieldTypeItem
- RadioboxFieldAttribute
- ComboBoxFieldAttribute
- DateTimeFieldAttribute
- ImageFieldAttribute
- CameraFieldAttribute
- AttachmentFieldAttribute
- HyperLinkFieldAttribute
- ToggleFieldAttribute
- NumericFieldAttribute
- CheckboxFieldAttributeV2
- RadioboxFieldAttributeV2
- Certificate(Validation)
- Scope(Validation)
- Digest(Validation)
- MetaData(Validation)
- Message(Validation)
- ValidationPolicy(Validation)

## ENUMERATION

Cả 18 mục chưa đối chiếu được, vì spec khai báo enum ngay trong field:

- WorkflowMode
- WorkflowStatus
- LevelOfAssurance
- LockAction
- LockPermission
- AnnotationPosition
- DocumentType
- SigningAlgorithm
- HashingAlgorithm
- AppearanceDesign
- AppearanceAlignment
- WorkflowState
- FieldTypes
- AuthorizeMode
- Role
- DocumentTypeForEID
- Color
- FontName
