# Scripts directory guide

This directory contains the runtime scripts used for the GoPaperless DOCX workflow.

## Source of Truth

- `GoPaperless_DOCX_Format_Template.md` is the single source of truth for formatting rules.
- If a formatting rule changes, update the template first.
- Update Python scripts only when they need to parse or apply a new template rule.

## Script roles

### Render from OpenAPI (primary workflow, see `CLAUDE.md`)

- `openapi_to_contract.py` flattens the backend's `openapi.yaml` into `contract/from_openapi.json`.
- `render_from_openapi.py` renders the specification DOCX from that contract; `--skeleton` builds everything from `templates/spec-skeleton.docx`.
- `build_skeleton.py` cuts a formatted document down to the skeleton. Run it only when the look of the document changes.
- `doc_meta.py` loads and validates `doc-meta.yaml` (cover, introduction, History).
- `propose_history.py` proposes or appends the next History entry when the API changed, and keeps the snapshots in `contract/openapi_history/`.
- `doc_planned.py` loads and validates `config/doc-planned.yaml` and merges the planned sections into the contract.
- `extract_doc_descriptions.py`, `extract_planned.py` generated the `config/` files from the approved V3 document. The files are edited by hand since; re-running overwrites them.
- `extract_doc_contract.py` reads a DOCX back into a contract; `compare_openapi_doc.py` compares it with `contract/from_openapi.json`.

### Formatter

- `fix_docx_format.py` applies the template-driven DOCX formatting workflow to `unpacked_docx/` and produces formatted output.

### Inspectors and analyzers

- `analyze_docx.py` summarizes page setup, styles, paragraph formatting, tables, and fonts from a DOCX file.
- `inspect_docx_layout.py` inspects paragraph spacing and heading-to-content transitions.
- `inspect_heading_xml.py` inspects heading paragraph XML details.
- `inspect_style_xml.py` inspects style XML details for key Word styles.
- `analyze_content_type.py` checks Content-Type consistency in request/header attribute sections.
- `check_frame_format.py` is the mandatory exit-code formatting gate. Check H rejects percentage widths in regular attribute cells and requires `Path Attributes` and generated connection-table cells to match `tblGrid`; `pack_docx()` enforces the same width invariant before writing any DOCX.

### Targeted fixers

- `fix_content_type.py` repairs Content-Type rows in request/header attribute tables.

### One-off or debug helpers

- `oneoff/check_upload.py` is an ad hoc checker for the `Upload Document` section in `unpacked_docx/word/document.xml`.
- Put future one-off investigation scripts under `scripts/oneoff/` so the main script list stays focused on reusable workflow tools.

## Path conventions

- `original/` holds source DOCX files and should remain unchanged unless explicitly requested.
- `formatted/` holds generated or review-ready DOCX outputs.
- `unpacked_docx/` is the working extracted DOCX package.

## Boundary with skill folders

- `.claude/skills/docx/` contains generic DOCX skill assets and helpers.
- `.claude/skills/gopaperless-docx/` contains project-specific workflow instructions.
- `scripts/` contains the project runtime logic that this repository actually uses.
