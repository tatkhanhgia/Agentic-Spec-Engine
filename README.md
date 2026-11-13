# Agentic-Spec-Engine
A specialized framework designed to bridge the gap between static company API templates and LLM-based agents. It automates the process of "teaching" agents (like Claude Code or Codex) to recognize, parse, and generate code based on unique internal API specifications.

## Prerequisites

- Python 3.8+
- pip

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/tatkhanhgia/Agentic-Spec-Engine.git
   cd Agentic-Spec-Engine
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Project Structure

- `templates/` - `spec-skeleton.docx`, the formatting skeleton the specification is rendered from.
- `config/` - Document-side content: description wording, object naming, planned sections.
- `contract/` - Generated machine-readable API contracts and History snapshots.
- `scripts/` - Rendering, checking and DOCX tooling.
- `formatted/` - Rendered and review-ready DOCX outputs.
- `reports/` - Dated verification reports.
- `original/` - Source DOCX files (keep unchanged).
- `unpacked_docx/` - Extracted working copy of a DOCX package used by the formatting scripts.

## Usage

The API specification DOCX is rendered from the backend's `openapi.yaml`. Check out
`GoPaperless-Workflow-API` (branch `dev`) beside this repository, then run from the
repository root:

```bash
python scripts/openapi_to_contract.py ../GoPaperless-Workflow-API/docs/api/openapi.yaml -o contract/from_openapi.json
python scripts/render_from_openapi.py templates/spec-skeleton.docx contract/from_openapi.json "formatted/out.docx" \
    --skeleton --planned config/doc-planned.yaml \
    --meta ../GoPaperless-Workflow-API/docs/api/doc-meta.yaml \
    --descriptions config/doc-descriptions.yaml --objects config/doc-objects.yaml
python scripts/check_frame_format.py "formatted/out.docx"
```

`CLAUDE.md` describes the whole workflow: what each input decides, where a change
goes, and how to verify a render. Refer to individual script headers for arguments
and options.
