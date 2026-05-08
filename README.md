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

- `original/` - Source DOCX files (keep unchanged).
- `formatted/` - Generated/review-ready formatted DOCX outputs.
- `scripts/` - Automation scripts for analyzing, formatting, and validating DOCX files.
- `unpacked_docx/` - Extracted working copy of a DOCX package used by formatting scripts.

## Usage

Run scripts from the repository root. For example:

```bash
python scripts/analyze_docx.py original/YourDocument.docx
python scripts/fix_docx_format.py
```

Refer to individual script headers for detailed arguments and options.
