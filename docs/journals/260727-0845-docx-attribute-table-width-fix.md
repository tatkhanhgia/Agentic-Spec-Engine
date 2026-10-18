---
date: 2026-07-27
session: docx-attribute-table-width-fix
---

# DOCX Attribute Tables Were Generated with Impossible Widths

**Date**: 2026-07-27 08:45  
**Severity**: Medium  
**Component**: DOCX formatter and format checker  
**Status**: Resolved

## What Happened

`Get Enterprise Invitations` exposed a systemic formatter defect: regular cells in its five-column `Path Attributes` table were written as `1666 pct`, while `tblGrid` defined `900 / 2200 / 1400 / 1200 / 3300 dxa`. Five cells at roughly 33.32% each describe about 166.6% of the table, so Word had to guess how to render it. The defect had spread through generated and cloned tables.

## The Brutal Truth

We hard-coded a three-column-looking percentage width into a generic cell factory and never checked it against the table grid. That was careless. The maddening part is that most legacy tables looked correct, which hid the broken generation path until a specific Enterprise API made the damage obvious.

## Technical Details

`scripts/fix_docx_format.py::make_table_cell()` no longer defaults every cell to `1666 pct`; callers must provide an explicit positive `dxa` or `pct` width. Generated attribute widths now come from the format template, and the normalizer repairs legacy regular cells against each table's `tblGrid`. The merged title cell deliberately remains `5000 pct`.

The checker now enforces rule H for width/grid consistency. On the primary artifact it dropped from **132 violations to 0**: **187 Path Attributes tables / 0 width issues** and **1 connection table / 0 width issues**. Targeted width tests and Python compilation passed. The formatter preserved **60,816 text nodes** and the extracted **265-endpoint contract**, proving this was a structural repair rather than a content rewrite.

## Decision

We rejected a one-off patch for `Get Enterprise Invitations`; it would leave the faulty factory and cloned descendants waiting to fail again. We chose prevention at four layers: explicit-width construction, normalization of existing DOCX content, checker/test enforcement, and a `pack_docx()` hard gate that refuses to create an invalid archive.

## Known Unrelated Failures

The full checker still exits non-zero on five baseline content defects: two check-E JSON indentation failures in `Get Workflow Users` response frames, plus three check-F empty Presence cells (`GetListWorkflowResponse` row 15, `Document.template`, and `FieldAuthentication` response attribute `totp`). They are unrelated to table widths and were intentionally not hidden or widened into this fix.

## Lessons Learned

Never encode column geometry in a generic cell helper. Treat `tblGrid` as the source of truth, and make the checker fail before a malformed DOCX becomes a donor template.

## Next Steps

- Documentation owner: resolve the five baseline E/F failures in the next content-cleanup pass.
- Formatter owner: keep targeted width tests and rule H mandatory for every generated DOCX.
