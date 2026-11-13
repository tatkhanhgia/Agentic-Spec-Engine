# Render from the skeleton (no earlier specification as seed) — 2026-11-13

## What changed

Until now the renderer took the approved V3 document as its seed and rewrote
it. It now also renders from `templates/spec-skeleton.docx`, a 139 KB document
that holds only formatting: the package (styles, numbering, headers, footers,
logo), the front matter, five endpoint sections used as templates, and one
section each of the object and enumeration chapters.

| Input | Decides |
|---|---|
| `openapi.yaml` (through `contract/from_openapi.json`) | categories, sections, fields, samples, objects, enumerations |
| `doc-meta.yaml` | cover, introduction, abbreviations, connection table, history |
| `config/doc-descriptions.yaml` | document wording of field descriptions and of section prose |
| `config/doc-objects.yaml` | document title and chapter of object sections |
| `templates/spec-skeleton.docx` | every presentation decision |

Commands:

    python scripts/build_skeleton.py "<formatted V3>.docx"          # only when the look changes
    python scripts/render_from_openapi.py templates/spec-skeleton.docx contract/from_openapi.json <out.docx> \
        --skeleton --meta <doc-meta.yaml> --descriptions config/doc-descriptions.yaml --objects config/doc-objects.yaml

## Result

Output: `formatted/GoPaperless Workflow RESTful API Specification V3 - rendered from skeleton.docx`

| Check | Result |
|---|---|
| Categories / sections built | 15 / 315 |
| Sections pairing with the spec | 315 of 315, 313 with no difference; the 2 others differ only by the `(Not implemented yet)` suffix in the title |
| Field rows only in the document or only in the spec | 0 |
| Object sections / enumeration sections | 124 / 21 |
| `check_frame_format.py` | 4 problems, all column widths of the System Connection Information table (present in the seed too) |
| Duplicate `docPr` ids, duplicate bookmarks, links to a missing bookmark | 0 / 0 / 0 |

The 14 "spec only" lines of the comparison are extra paths of operations
documented once (path aliases), as in the seed-based render.

## New in this step

* `x-doc-sections[].example` (spec): the named request example a split section
  shows. `Client Credentials Authentication` and `Password Authentication` now
  each show their own request; before, both fell back to field examples.
* Section prose: the paragraphs under a section heading are extracted into
  `config/doc-descriptions.yaml` (`description`, 230 sections) and used instead
  of the spec description when present.

## What the skeleton render does not carry over from V3

* Sections that exist only in V3 and have no operation in the spec (most of
  the "Not implemented yet" ones). The seed-based render kept them untouched;
  the skeleton render has nothing to build them from.
* Heading 4 sub-structure: every operation is one Heading 3 section.
* Not yet opened in Microsoft Word (deferred by the owner).

The seed-based output is kept beside the new one until the owner accepts the
skeleton render.
