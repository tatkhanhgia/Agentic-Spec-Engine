# Project Context

## Directory Structure
- `original/` contains source/original DOCX files that should remain unchanged unless the user explicitly asks to modify the source document.
- `formatted/` contains generated or review-ready formatted DOCX outputs.
- `scripts/` contains automation used to analyze, format, or validate DOCX files.
- `unpacked_docx/` contains the extracted working copy of a DOCX package used by formatting scripts.
- `contract/` contains the machine-readable API contract used to detect drift between the document and the backend. It is generated, never hand-edited.
- **File sync between machines** lives in the sibling project `../GoPaperless-Sync-Pack/` (Syncthing + Tailscale). `scripts/sync-pack/` here is deprecated; do not extend it.

## Multiple Agents Edit This Repository

More than one agent works here at the same time. The DOCX under `unpacked_docx/`
and `formatted/` is shared mutable state.

- **Section ownership**: the other agent owns the **Document Package** section.
  Do not edit it. Coordinate through the user before touching a section you do
  not own.
- **Never assume the file on disk is the file you read.** Use
  `scripts/docx_guard.py` (`load()` / `save()`); `save()` refuses to write when
  the file changed in the meantime.
- **Never restore a backup to "repair" unexpected content.** Content you did not
  write is far more likely to be another agent's deliberate work than corruption.
  Restoring is destructive and silently discards it — investigate and ask first.
  This happened on 2026-07-27: `Download Attachment (base64)` was intentionally
  removed by the other agent, mistaken for data loss, and restored, undoing their
  change.
- A text-level diff cannot prove a formatting-only edit is correct. Run
  `scripts/check_frame_format.py` after any DOCX change; it exits non-zero on
  failure.

## Documentation Update Workflow

How to bring the specification in line with the system. Follow this whenever the
task is "what is missing / wrong in the docs" or "add endpoint X".

### Source hierarchy — what each source is allowed to decide

| Source | Authoritative for | Never authoritative for | How to read it |
|--------|-------------------|-------------------------|----------------|
| **Backend code** (`validURLs` + `web.xml`) | Which endpoints exist, their method and path, request/response fields, whether an endpoint actually works | Naming style, wording, whether a design is *good* | `scripts/extract_code_contract.py` |
| **Postman collection** | Sample values, request bodies, headers really sent | Which endpoints exist (it lags), response shapes (only ~11 of 248 requests store one) | `Downloads/GoPaperless Workflow Portal.postman_collection.json` — export cadence is not guaranteed, treat as supplementary |
| **SigningHub 10.0 manual** | Design standard: attribute naming, table structure, the wording style of descriptions, what a call *should* look like | What our backend does; it cannot overrule our code | NotebookLM — see the `gkg-signinghub-reference` skill |
| **Format template** | Every presentation decision | Content | `scripts/GoPaperless_DOCX_Format_Template.md` |

The backend registry is the primary source because it cannot go stale: a route
missing from `validURLs` is rejected at runtime. Postman and SigningHub are
cross-checks, not replacements.

### The loop

1. **Refresh both contracts** — run `extract_code_contract.py` and
   `extract_doc_contract.py`. Never compare against a stale `contract/*.json`.
2. **Compare and classify** every difference into exactly one of:
   *docs missing* · *backend missing* · *method or path mismatch* ·
   *deliberate divergence*. The classification is the deliverable, not a raw diff.
3. **Let the user decide** each item. Do not resolve a docs-vs-code disagreement
   on your own — either side can be the wrong one.
4. **Write by cloning.** To add a section, find an existing section that already
   matches the target shape (query `contract/from_doc.json` for method, header
   set, field counts, empty-vs-populated success frame) and copy it, then
   substitute text. Never assemble WordprocessingML by hand.
   Cloning duplicates identifiers that must be made unique again —
   `w:bookmarkStart`/`End`, `wp:docPr/@id`, and VML `v:shape/@id` + `@o:spid`
   (a separate id space) — or Word reports the file as corrupt.
5. **Verify on three axes**, because each is blind to the others:
   - text: leaf-paragraph diff — expect only the intended changes, zero
     unexplained deletions;
   - formatting: `scripts/check_frame_format.py` — a text diff cannot see a
     badge colour or a lost alignment;
   - structure: re-run `extract_doc_contract.py` and confirm the endpoint parses
     as intended and matches `from_code.json`.
6. **Deliver** — back up first, then write, then copy to `%USERPROFILE%\Documents\Sync` (GoPaperless Sync Pack — see `../GoPaperless-Sync-Pack/`) only
   when no other agent is mid-edit.
7. **Record the session** — `python scripts/changelog.py session <baseline> --write`
   prepends a dated entry to `CHANGELOG.md`. Take the baseline snapshot at the
   start of a session (`changelog.py snapshot`, stored in `contract/history/`)
   and compare against it at the end. The entry is derived entirely from the
   files: nothing is written from memory, so it stays correct even when another
   agent edited the document in the same period.

### Rules that are easy to get wrong

- **Do not document an endpoint that does not work.** Some routes exist but are
  stubs (the TSE login calls have their processing line commented out).
  Documenting them ships a promise the API cannot keep.
- **Do not invent values.** Error codes, field names and enum values come from
  the code or from the user — never from inference.
- **Record deliberate divergences in the template**, not in commit messages, so
  the next pass does not "fix" them back toward SigningHub.
- **Ask before widening scope.** Fixing one instance of a defect class usually
  reveals more; report the count and let the user choose.
- **Deferred defects go to `TECHNICAL_DEBT.md`.** Two files record "leave this
  alone for now" and they mean opposite things — never mix them up:

  | File | Meaning | A later pass must |
  |------|---------|-------------------|
  | Format template, section 5 | *Deliberate* divergence, decided by the owner | **never** "fix" it |
  | `TECHNICAL_DEBT.md` | *Unintended* defect, deferred | **fix** it when scope allows |

  Moving an item between the two changes what future agents do with it, so ask the
  owner first. Re-measure the counts in `TECHNICAL_DEBT.md` before acting on them —
  they were true when written, and another agent edits this document too.

## API Contract Extraction
- The backend lives at `C:\Users\Admin\Documents\NetBeansProjects\Library_RSSP_SDK\ProjectRSSP_newest\Gopaperless_Workflow_Maven` (Java servlet, JBoss — not Spring; its Swagger config is commented out and unusable).
- Routes are declared per controller in a static `validURLs` registry plus `web.xml` servlet mappings. That registry is the source of truth for what endpoints exist: a route missing from it is rejected at runtime.
- `scripts/extract_code_contract.py` reads the backend and writes `contract/from_code.json`. It is strictly read-only: never run Maven, never write to that project, never touch its git state.
- `scripts/extract_doc_contract.py` reads the formatted DOCX and writes `contract/from_doc.json`. Read-only as well.
- The template governs *presentation*; the contract governs *content*. Changing formatting rules never belongs in `contract/`.

## GoPaperless DOCX Formatting Workflow
- Use `scripts/GoPaperless_DOCX_Format_Template.md` as the single source of truth for DOCX style rules.
- Before changing any DOCX formatting script behavior, first add or update the corresponding rule/value in `scripts/GoPaperless_DOCX_Format_Template.md`; scripts must read formatting values from this template instead of hardcoding them.
- Use `scripts/fix_docx_format.py` to apply formatting rules to `unpacked_docx/` and generate a formatted DOCX.
- Save generated formatted DOCX files under `formatted/`, not the project root.
- Keep original source files under `original/` for comparison and recovery.

## Document Heading Conventions
- Heading 2 = API category sections, for example `Authenticate - Log in`.
- Heading 3 = API endpoints, for example `Client Credentials Authentication`, `OTP Login Authentication`, and `Single Sign On Authentication`.
- Heading 4 = sub-features or sub-steps under an endpoint/module, for example `Sign Document` under `Sign with GoPaperless Workflow`.
- All authentication endpoints under `Authenticate - Log in` must use Heading 3.

## Composite Types Link by Bookmark, Not by Name

When an attribute table's `Type` cell holds something that is not a primitive
(`owner`, `List<document>`, `List<Packages>`, `Certify`, `dimension`, ...), the
object it refers to is defined under one of two Heading 1 chapters:

- **REQUEST/RESPONSE OBJECT** — request and response payload objects.
- **RELATIVE OBJECT** — nested/shared objects (99 of them), e.g. `Package`,
  `Document`, `Owner(Package Details)`, `Decline`.

**The link is a `w:hyperlink` whose `w:anchor` points at a bookmark on the
object's Heading 2. The text in the cell is only a display label.** The label
and the heading routinely differ, and that is not a defect:

| Endpoint · field | Label shown | Anchor | Resolves to |
|---|---|---|---|
| `Get All Packages` · `packages` | `List<Packages>` | `_Items` | `Package` |
| `Get Package Details` · `owner` | `owner` | `_Owner(Package_Details)` | `Owner(Package Details)` |
| `Get Package Details` · `documents` | `List<document>` | `_Document(Package_Details)` | `Document(Package Details)` |
| `Package` · `decline` | `decline` | `_WorkflowDetails` | `Decline` |

Anchor names are Word-generated and often meaningless (`_Items`,
`_WorkflowDetails`). Judge a link by what the anchor **resolves to**, never by
how the anchor or the label is spelled.

Consequences when editing:

- **Never "fix" a Type label to match its object heading.** Singular/plural and
  casing differences are cosmetic. Resolve the anchor first; if it lands on the
  right heading there is no bug.
- **Never write a Type cell with a naive "set the first run, blank the rest"
  helper.** The runs live inside the `w:hyperlink`; that empties the visible
  label while leaving the anchor, producing a blank clickable cell. Primitive
  Type cells carry no hyperlink and are safe.
- **The same label can point at different objects.** `List<document>` is used by
  `Get Package Details` and by three Document Workflow endpoints, but the two
  document shapes differ, which is why both `Document` and
  `Document(Package Details)` exist. Follow the anchor to know which one applies.
- **Checking an endpoint is not finished until its objects are checked.** An
  endpoint table can be fully correct while the object behind a composite type
  is stale. On 2026-07-27 the `Get All Packages` sample was aligned to the
  backend while the `Package` object still carried `unread` and `next_signer`,
  which the serialiser never writes — the two then contradicted each other.

Verify links after any table edit: every `w:hyperlink` in a Type cell must have
a non-empty label and an anchor present in the document's `w:bookmarkStart`
names.
