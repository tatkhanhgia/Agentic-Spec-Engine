# Technical Debt

Known defects in the specification that are **not** fixed yet and **should be**.

Do not confuse this file with section 5 of
`scripts/GoPaperless_DOCX_Format_Template.md`. That section lists *deliberate*
divergences which a later pass must **never** "correct". Everything here is the
opposite: it is unintended, and a later pass **should** clear it. Never move an
item between the two files without asking the project owner.

Every count below was measured from `unpacked_docx/word/document.xml`, not
recalled. Re-measure before acting - another agent edits this document too.

Last measured: 2026-07-27 (202 documented endpoints, 839 tables).

---

## 1. `issuerDNObject` is referenced but never written

**Priority: highest here** - it is the only broken link left in the object chain
reachable from Document Preparation. Owner asked on 2026-07-27 to record it and
add the object later.

```
Get Document Fields
  └─ .inperson ─────────► SignatureFieldAttribute        ok
       └─ .verification ─► VerificationData              ok
            └─ .issuer_dn_object ─► issuerDNObject       BROKEN
               anchor `_IssuerDNObject` resolves to nothing
```

`VerificationData` row 20 types `issuer_dn_object` as `issuerDNObject`, but no
such object exists in `REQUEST/RESPONSE OBJECT` or `RELATIVE OBJECT`, and no
bookmark of that name exists anywhere.

**Why it was not written:** the field list cannot be sourced. The string
`issuerDNObject` appears in exactly one backend file,
`ProjectRSSP_newest/FPS/src/java/vn/mobileid/id/fps/object/Signature.java`, and
that file is commented out in its entirety. `Gopaperless_Workflow_Maven` defines
no such class. CLAUDE.md forbids inventing field names, so nothing was guessed.

**To clear it,** one of:
- the owner supplies the attribute list, then build the object the same way the
  eleven field-attribute objects were built on 2026-07-27
  (`scripts/oneoff/add_field_attribute_objects.py`);
- or retype `issuer_dn_object` as `String`, matching `issuer_dn` on the row above it.

Verify afterwards with the transitive check: walk every composite Type link out of
the 44 Document Preparation endpoints and confirm each anchor resolves. That walk
reaches 35 objects at depth 2 and currently reports this one failure.

---

## 2. Duplicate drawing and bookmark identifiers

Word treats duplicated identifiers as corruption and may offer to "repair" the
file, silently dropping content.

| Kind | Duplicates | Where |
|------|-----------|-------|
| `wp:docPr/@id` | 9 | Document Package: Download Document (base64), Download Signature Summary (base64), Merge Document, Convert Document to Paper, Download Attachment |
| VML `v:shape/@id` | 9 | same sections |
| VML `@o:spid` | 2 | same sections |
| `w:bookmarkStart/@id` | 22 | document-wide |
| orphaned `w:bookmarkEnd` | 3 | ids 43, 69, 76 |

**Not blocked by ownership on the technical side, but note where they live:** the
docPr / shape / spid duplicates are all inside the **Document Package** section,
which the other agent owns per CLAUDE.md. Coordinate before touching it.

**Safe to fix:** a scan on 2026-07-27 found **zero** shapes using
`mso-next-textbox`, so no text-box chain can break when shape ids are reallocated.
Bookmark ids need more care: hyperlinks resolve by *name*, but a bookmark's start
and end pair by *id*, so both ends must be renumbered together.

Pattern to follow: `scripts/oneoff/repair_new_section_ids.py`.

---

## 3. Dangling cross-reference anchors

Ten `w:hyperlink` anchors resolve to no bookmark. `_Toc*` anchors are excluded -
Word rebuilds the table of contents on open because `w:updateFields` is set.

| Anchor | Uses | Note |
|--------|------|------|
| `_Sign Document via RSSP Directly` | 3 | |
| `_Add User to Workflow` | 2 | |
| `_Address Book` | 1 | **typo** - the real bookmark is `_Address_Book` (underscore, not space) |
| `_Get Worklow Users` | 1 | **typo** - "Worklow" is missing an `f` |
| `_Enterprise_Management` | 1 | |
| `_Add Contact` | 1 | |
| `_Get Address Book` | 1 | |
| `_Create Account` | 1 | |
| `_Get All Workflow` | 1 | |
| `_IssuerDNObject` | 1 | see item 1 |

The two typos are one-character fixes. The rest need their intended target
identified first - the method that worked on 2026-07-27 is to find other cells
linking the *same* field name and take the majority spelling.

---

## 4. Orphaned object definitions

Re-measured 2026-07-27 (later session): **ten** objects that nothing links to.

RELATIVE OBJECT: `NextSigner`, `Reminder`, `Address Book`, `Template`,
`DurationDate`, `FieldAuthentication`, `Authentication(SigningRequest)`.
REQUEST/RESPONSE OBJECT: `AddUserToWorkflowResponse`, `GetWorkflowHistoryResponse`,
`RegisterUserRequest`.

`Address Book` is orphaned only because its inbound link is the typo'd
`_Address Book` from item 3 - fixing that typo removes this entry too. The six
newly measured ones were not in the original count; likely causes are the
2026-07-27 section removals (Qrypto, signature status) de-linking them, or the
original count only covering the RELATIVE OBJECT chapter. Each needs a decision:
either something should link to it, or it is dead and can go. None of them is
referenced as plain text in a Type cell either (checked 2026-07-27).

---

## 5. Attribute naming violations of template 2.8.9

The rule: JSON body and response keys are `snake_case`; URL path and query
parameters are `camelCase`. Both halves are currently violated.

**Body / response spelled camelCase** (should be snake_case):

| Name | Occurrences |
|------|-------------|
| `pageNo` | 8 |
| `recordsPerPage` | 8 |
| `packageId` | 4 |
| `credentialIDs` | 1 |
| `responseID` | 1 |

Check the four `packageId` cases first - they may be path parameters mis-parsed
into the Request table rather than real body fields.

**URL parameters spelled snake_case** (should be camelCase):

| Parameter | Occurrences |
|-----------|-------------|
| `{session_id}` | 4 |
| `{page_no}` | 3 |
| `{records_per_page}` | 3 |
| `{appearance_id}` | 2 |
| `{document_status}` | 1 |

Note the document spells the same concept both ways: `{pageNo}` and
`{recordsPerPage}` already appear correctly elsewhere.

---

## 6. JSON body indentation

`check_frame_format.py` check E reports 4 failures, in `Get Workflow Users`
(Document Workflow) and `Fill Form Fields` (Document Processing) - two frame
copies each. Template 2.8.4 requires 2 spaces per nesting level.

---

## 7. Backend changes the document is already written against

The document states the standard and the backend is expected to follow. Recorded
in template section 5.2 so a drift pass does not reverse the document.

- **`ElaborFieldAttribute` needs `@JsonProperty`** on `maxLength` and `readOnly`.
  Until then a client sending the documented `max_length` / `read_only` gets **no
  error** - `@JsonIgnoreProperties(ignoreUnknown = true)` drops the key and the
  resolver default is applied.
- **`alignment` vs `align_option`** - the Elabor classes name the alignment object
  `alignment` while every sibling field type names it `align_option`. Both are
  snake_case so template 2.8.9 does not decide it; this needs a product decision.
- **`DELETE /v1/tokens`** - carried over from an earlier session; the backend still
  uses another verb where the design standard requires DELETE.
- **`GET /v1/csc/credentials/{transactionCode}/status` emits camelCase** (owner
  decision 2026-07-27: document snake_case, backend must follow). The success
  response serializes `GetCredentialStatusResponse` through the default Jackson
  mapper (`JacksonJsonService.writeValueAsString`), emitting `data.requestType`,
  `data.citizenPid`, `data.credentialId`, `data.serialNumber`,
  `data.certificateData`, plus an internal `header` object when non-null. The
  document lists the keys as `request_type`, `citizen_pid`, `credential_id`,
  `serial_number`, `certificate_data`. Fix: switch the endpoint to
  `writeValueAsStringWithSnakeCase` (or annotate the DTO) and stop serializing
  `header`. Same defect class, not yet decided: the `Get Signing Transaction
  Status` 200 sample shows `data.transactionCode` / `billCodeMap` /
  `requestTimestamp` - camelCase copied from the same mapper.
- **`PersonalInformations.personalIdentificatetion` is a typo** (owner decision
  2026-07-27: document the correct spelling). `Get Information of Citizen card`
  serializes this field as `personal_identificatetion`; the document spells it
  `personal_identification`. Fix: rename the Java field or add
  `@JsonProperty("personal_identification")` in
  `eid.restful.sdk/src/eid/restful/objects/PersonalInformations.java`.
- **`Update Address Book` uses the wrong verb** (owner decision 2026-07-27:
  follow SigningHub, backend to be fixed later). The document and SigningHub
  (`PUT /v4/settings/groups/{id}`) say **PUT** `/v1/settings/addressbook/{id}`;
  the backend currently routes the update through **POST**
  (`ContactController` `doPost`, validURLs entry 5, "Update AddressBook"), and
  the Postman collection still calls POST. Same class as the `DELETE /v1/tokens`
  item above: the document is correct, do NOT "fix" it to POST. Clear this when
  `ContactController` gains a `doPut` branch for
  `.*settings/addressbook/[0-9]*`.
- **`DeleteUserDTO` body keys are camelCase** (template 2.8.9). The
  `POST /v1/account/remove` request DTO declares `userEmail` (optional,
  defaults to the token's user) with no `@JsonProperty("user_email")`. The
  document deliberately documents only `password`; if the optional email
  override is ever documented, the backend must first accept `user_email`.

---

## 8. Blank attribute-table rows (check F, 2026-07-27)

`check_frame_format.py` check F (Presence completeness, template 2.4) reports
three rows whose Presence — and in two cases Type and Description — cells are
blank. The eleven blank cells in `Create Workflow From Template` and
`Get Status of Import Excel` from the same defect class were already filled with
`M` on 2026-07-27, sourced from `CreateWorkflowFromTemplateResponseDTO` and
`CloneWorkflowStatusSerializer`. These three remain because no source for their
values exists; each needs an owner decision:

| Where | Row | Blank cells | Evidence gathered |
|-------|-----|-------------|-------------------|
| `FieldAuthentication` (RELATIVE OBJECT 5.28) | 3 `totp` | Type, Presence, Description | No `totp` anywhere in `Gopaperless_Workflow_Maven` or `fps.core`; no `Totp` object exists in the document. SigningHub manual check still pending (NotebookLM auth expired). Likely dead stub — delete, or fill if the design standard defines it and the backend will follow. |
| `Document` (RELATIVE OBJECT) | 16 `template` | Type, Presence, Description | CHANGELOG 2026-07-27 removed `template` from `Get Document Details` and `Apply Workflow Template` responses — this row is probably the same cleanup left half-done. Confirm and delete, or fill. |
| `GetListWorkflowResponse` (REQUEST/RESPONSE OBJECT) | 15 | all five | Entirely blank row, no name to trace. Almost certainly an accidental extra row — confirm and delete. |

Do not delete without the owner's confirmation: blank rows can also be another
agent's work in progress.

---

## 9. CHANGELOG entries overlap under concurrency

`scripts/changelog.py` derives an entry by diffing two snapshots, so an entry
covers **every** change in that window, including another agent's. When two agents
each run it, both entries describe both agents' work from different snapshot
points, and `CHANGELOG.md` ends up with two same-day entries whose numbers differ
(202 vs 203 endpoints on 2026-07-27).

Nothing is wrong in the data. If it becomes confusing, record which baseline and
which agent produced each entry in its heading.

---

## 10. Template module: two multipart endpoints undocumented (2026-07-27)

`Create Personal Template` (`POST /v1/settings/templates`) and
`Create Enterprise Template` (`POST /v1/enterprise/templates`) are live in
`TemplateController` and present in the Postman collection, but are still
missing from the specification. Everything needed to write them is settled:

| | |
|---|---|
| Request | `multipart/form-data`: `file` (binary, required), `name` (required), `description`, `read_only`; the enterprise variant adds `public` |
| Response | **201**, `CreateTemplateResponseDTO` = `{ "id", "template_id", "documents": [ { "id", "document_id", "name" } ] }` (`id` and `template_id` both carry the template id) |
| SigningHub | same two paths and the same multipart schema — this is not a divergence |

**Why it was deferred, and what has to be decided first:** no endpoint anywhere
in the specification currently uses `multipart/form-data`. There is therefore no
donor section to clone and no rule in the format template for how a multipart
request frame is written — a JSON body frame cannot represent form parts. Adding
these two means inventing a presentation convention, which is an owner decision,
not a gap to be filled silently.

Proposal to put to the owner: add a rule to format template 2.8.4 for multipart
frames (`Content-Type: multipart/form-data`, one line per part as
`name: <value>` with `file: <binary>`), then clone from whichever of the two is
written first.

## 11. Template composite types: linked, with three deliberate exceptions

**Resolved 2026-07-27.** Every composite `Type` cell in the Template chapter now
links to the object it names, verified by resolving each anchor against the
document's `w:bookmarkStart` names:

| Where | Type | Resolves to |
|---|---|---|
| `Get Templates`.`templates` | `List<template>` | `Template(Get Templates)` |
| `Get Templates`.`documents` | `List<TemplateDocumentResponse>` | `TemplateDocumentResponse` |
| `Get Template By ID`.`documents` | `List<TemplateDocumentResponse>` | `TemplateDocumentResponse` |
| `Get Template Details`.`documents` | `List<TemplateDocumentResponse>` | `TemplateDocumentResponse(Get Template Details)` |
| `Get Template Details`.`participants` | `List<TemplateWorkflowParticipant>` | `TemplateWorkflowParticipant` |
| `Add` / `Update Template Field`.`dimension` | `dimension` | `Dimension` |
| `Template(Get Templates)`.`documents` | `List<TemplateDocumentResponse>` | `TemplateDocumentResponse` |

Three objects were created for this, and `TemplateDocumentResponse` was corrected
(it was missing `document_id`, and its title row still read
`GetCertificatesProfileResponse` from an old clone).

`TemplateDocumentResponse` had to be split: `Get Template Details` returns
`form_fields` inside each document and the other two template calls do not, so
one object could not describe both. Naming follows the existing
`Owner(Package Details)` / `Document(Package Details)` precedent.

### Still not linked — because `Get Template Details` leaks internal DB objects

Investigated 2026-07-27 on the owner's instruction to link the remainder. The
answer is that they **must not** be linked, and the reason is a backend defect,
not a documentation gap.

`GetTemplateDetailsSerializer` hands five attributes straight to Jackson with
`writeObjectField`, so the API returns the persistence objects verbatim —
camelCase keys and all, which contradicts the project's own naming rule that
response keys are snake_case:

| Attribute | Actual Java type written | Leaked keys |
|---|---|---|
| `workflow` | `GetWorkflowDTO` | `packageId`, `enterpriseId`, `workflowUuid` (+ `DatabaseDefaultObject`) |
| `workflow_details` | `List<GetWorkflowDetailDTO>` | `packageId`, `workflowAttributeId`, `value` |
| `workflow_categories` | `List<WorkflowCategoryDTO>` | `workflowId`, `categories`, `categories_`, `categoryId`, `totalCategories` |
| `participants` | `List<GetWorkflowParticipantDTO>` | `id`, `packageId`, `workflowId`, `email`, `firstName`, `lastName`, `mobileNumber`, `type`, `processAt`, `signingOrder`, `metaInformation`, `status`, `uuid`, `guestUser`, `emailNotification`, **`securityDataEncrypted`** |
| `form_fields` | `List<ExtendedFieldAttribute>` | raw field attribute object |
| `tasks` | `HashMap<Long, List<WorkflowTask>>` | a map keyed by document id — dynamic keys, so it cannot be an attribute table at all |

Consequences, all confirmed rather than assumed:

* **`workflow` is NOT the RELATIVE OBJECT `Workflow`.** That object
  (`workflow_type`, `continue_on_decline`, `workflow_status`, …) is what the
  dedicated serializer behind `Get Workflow Details` produces. `GetWorkflowDTO`
  is three database columns. Linking them would have asserted a match that does
  not exist — and the names made it look right.
* **`participants` is NOT the same shape as `Get Template Workflow Users`.**
  That endpoint runs `TemplateWorkflowParticipantSerializer`; this one does not.
  A link was briefly made here and then reverted, and the
  `TemplateWorkflowParticipant` object created for it was removed, since the
  five root-array endpoints have no attribute name to link from and an orphan
  object invites the same mistake again.
* `securityDataEncrypted` being reachable in an API response is worth raising
  with the backend owner on its own merits.

**Do not document these sub-objects and do not link them.** Documenting them
would enshrine the leak as the contract. The fix belongs in the backend: give
`Get Template Details` a serializer that emits a defined snake_case shape, the
way every other endpoint in this module already does. Until then the attributes
stay typed `Object` / `Array` with a description saying what they carry.

`smart_form` is left alone by owner decision (2026-07-27); both counts are
hardcoded `0` in `GetListTemplateSerializer` anyway.

### The five root-array participant tables are NOT debt

`Get / Add / Update Template Workflow Users` and the two placeholder calls return
a **root-level JSON array**, so there is no attribute name to hang a composite
`Type` on. Their Response Attributes tables therefore list the participant
attributes inline, which is correct rather than duplicated-by-oversight. The
shared definition now exists as `TemplateWorkflowParticipant` for the one place
that can reference it (`Get Template Details`.`participants`).

A change to a participant attribute must still be applied to all five tables plus
the object — six places. Inventing a wrapper field just to enable a link would
document a payload the backend does not send.

### Ordering hazard for whoever edits these next

`section_tools.set_cell_text` drops every direct child of a cell paragraph,
including `w:hyperlink`. Re-running `add_template_workflow_sections.py`,
`add_template_rest_sections.py` or `move_get_templates_section.py` would
silently strip every link above. Do content edits first, links last.

Two headings are ambiguous by title and must be resolved with their parent:
`Template` exists as both an API H2 and a RELATIVE OBJECT H2, and `Get Templates`
exists as both the real H3 and the Personal Settings cross-reference stub.
`section_tools.find_heading` returns the first match, which is the wrong one in
both cases — see `h2_in_chapter` / `h3_in_category` in
`scripts/oneoff/link_template_composite_types.py`.

---

## 14. Enterprise/Account group-and-role sections documenting routes absent from the backend registry (2026-07-27)

Owner decision 2026-07-27: keep the sections, record here, verify the RBAC/groups
roadmap with the backend team before deciding removal, PLANNED status, or
reimplementation.

| Method | Path | Doc section (module) |
|--------|------|----------------------|
| POST | `/v1/enterprise/roles` | Add Role (Enterprise Management) |
| DELETE | `/v1/enterprise/roles/{id}` | Delete Role (Enterprise Management) |
| GET | `/v1/enterprise/roles/{id}/users/{pageNo}/{recordPerPage}` | Get Role Users (Enterprise Management) |
| GET | `/v1/settings/groups/{pageNo}/{recordPerPage}` | Get Groups (Enterprise Management) |
| DELETE | `/v1/settings/groups/{id}` | Delete Groups (Enterprise Management) |
| POST | `/v1/enterprise/groups/{id}` | Add Member into Group (Enterprise Management) |
| DELETE | `/v1/enterprise/groups/{id}` | Delete Member in Group (Enterprise Management) |
| GET | `/v1/settings/groups/{pageNo}/{recordPerPage}` | Get Group (Account Management) |
| GET | `/v1/enterprise/users/{id}` | Get Enterprise User (Enterprise Management) |
| PUT | `/v1/enterprise/users` | Update Enterprise User (Enterprise Management) |
| POST | `/v1/enterprise/users/{id}/reset-password` | Reset Enterprise User Password (Enterprise Management) |

None of these routes exists in any controller `validURLs` registry or `web.xml`
mapping (verified by full enumeration of `EnterpriseController`/`UserController`
on 2026-07-27). Backend only implements `GET /v1/enterprise/roles/{p}/{r}`
(Get Roles) from this family.

---

## 15. Update Enterprise TSA Configuration response shape needs runtime verification (2026-07-27)

The documented PUT `/v1/enterprise/settings/tsa` response follows the
`UpdateEnterpriseTSAConfigurationResponse` DTO fields. The serializing
`IJsonService` instance is created with `SerializationFeature.WRAP_ROOT_VALUE`
enabled (`EnterpriseService.java` ~:607-610) and the DTO has no `@JsonRootName`,
so the actual wire body may be wrapped in a root name. Owner decision
2026-07-27: document per DTO now, verify with a live call later.

**To clear it:** call the endpoint on a test environment, compare the body with
the documented table, and either confirm or fix the section (and ideally the
backend serializer).

---

## 16. Validation Management backend quirks (module documented 2026-07-27)

The five Validation Management endpoints were added to the specification on
2026-07-27 (`scripts/oneoff/add_validation_management_sections.py`), documenting
**actual** backend behaviour per the owner's decision. Four backend defects were
found while reading `ValidationController` / `ValidationService`; the document
reflects what the API does today, so each fix below must also update the doc.

- **(a) `Download Reports` simple report: wrong file name and Content-Type.**
  `ValidationService.downloadReport` names the attachment with
  `isDetailed ? (isSimple ? "SimpleReport" : "DetailedReport") : "DiagnosticReport"`,
  so the `/simple` branch (isDetailed false) yields `DiagnosticReport.pdf`; and
  `ValidationController.doGet` sets `Content-Type: application/pdf` only when the
  URI contains `detailed`, so `/simple` returns a PDF payload labelled
  `application/xml`. The Download Reports section documents this behaviour
  explicitly. When the backend fixes the ternary and the content type, update the
  section's description and 200 frame.

- **(b) `Download Reports` performs no token verification.** Unlike every other
  validation call, `downloadReport` never calls `verifyToken`; the report is
  downloadable by anyone holding the report hash. The doc keeps `Authorization`
  as `M` (every client sends it, and the backend is expected to start verifying);
  security-relevant, raise with the backend team.

- **(c) `Delete Validation` has no ownership check.**
  `ValidationService.deleteValidation` verifies the token but never compares the
  validation's owner with the caller (unlike `getDataValidation`, which rejects
  with subcode 7). Any authenticated non-client-credential user can delete any
  validation by ID. Security-relevant, raise with the backend team.

- **(d) `annotation_dimension` / `annotation_page` are never emitted.**
  `EnrichDSSResponse` extracts the field name from a `signature_scope` array,
  but the report serializer (SNAKE_CASE on `ValidateCore.getScope`) emits a
  single `scope` object, so the match never fires and the two enrichment fields
  are never added to any signature. They are deliberately **not** documented in
  `Signature(Validation)`. If the backend fixes the key mismatch, add both
  fields to the object and to the Get Report of Validation sample.

Illustrative values note: the 202 frame of `Get Report of Validation` shows
`error_code "28.8"` (CODE_VALIDATION 28 + subcode 8, composed as
`code + "." + subcode` by `ResponseMessageFactory`); its `error_description` /
`message` wording is illustrative — the real strings live in the response-code
database and were not read.
