# Project Context

## Directory Structure
- `original/` contains source/original DOCX files that should remain unchanged unless the user explicitly asks to modify the source document.
- `formatted/` contains generated or review-ready formatted DOCX outputs.
- `scripts/` contains automation used to analyze, format, or validate DOCX files.
- `unpacked_docx/` contains the extracted working copy of a DOCX package used by formatting scripts.

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
