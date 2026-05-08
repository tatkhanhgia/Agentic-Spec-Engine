# Scripts directory guide

This directory contains the runtime scripts used for the GoPaperless DOCX workflow.

## Source of Truth

- `GoPaperless_DOCX_Format_Template.md` is the single source of truth for formatting rules.
- If a formatting rule changes, update the template first.
- Update Python scripts only when they need to parse or apply a new template rule.

## Script roles

### Formatter

- `fix_docx_format.py` applies the template-driven DOCX formatting workflow to `unpacked_docx/` and produces formatted output.

### Inspectors and analyzers

- `analyze_docx.py` summarizes page setup, styles, paragraph formatting, tables, and fonts from a DOCX file.
- `inspect_docx_layout.py` inspects paragraph spacing and heading-to-content transitions.
- `inspect_heading_xml.py` inspects heading paragraph XML details.
- `inspect_style_xml.py` inspects style XML details for key Word styles.
- `analyze_content_type.py` checks Content-Type consistency in request/header attribute sections.

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
