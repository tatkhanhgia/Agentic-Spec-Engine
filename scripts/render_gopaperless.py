#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
One-command entrypoint that turns a backend OpenAPI spec into a release-ready
specification DOCX and gates it.

This is the single command GoPaperless CI should call. It runs, in order:

  1. openapi_to_contract.py   openapi.yaml            -> <workdir>/contract.json
  2. render_from_openapi.py   skeleton + contract      -> <out>  (--skeleton mode)
  3. check_frame_format.py    <out>                    -> exit code

and exits non-zero if any step fails, so a release is only published when the
document renders clean. Front matter (cover, introduction, History, connection
table) comes from doc-meta.yaml; every presentation decision comes from the
committed skeleton, styles and config files - not from any earlier document.

Usage:
    python scripts/render_gopaperless.py --openapi <openapi.yaml> --meta <doc-meta.yaml> --out <out.docx>

Optional overrides (defaults are the repo's committed inputs):
    --skeleton <docx>      default templates/spec-skeleton.docx
    --descriptions <yaml>  default config/doc-descriptions.yaml
    --objects <yaml>       default config/doc-objects.yaml
    --planned <yaml>       default config/doc-planned.yaml
    --workdir <dir>        where contract.json is written (default: next to <out>)
    --keep-contract        leave the intermediate contract.json in place
"""
import argparse
import os
import subprocess
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(REPO, "scripts")

DEFAULTS = {
    "skeleton": os.path.join(REPO, "templates", "spec-skeleton.docx"),
    "descriptions": os.path.join(REPO, "config", "doc-descriptions.yaml"),
    "objects": os.path.join(REPO, "config", "doc-objects.yaml"),
    "planned": os.path.join(REPO, "config", "doc-planned.yaml"),
}


def run(argv):
    """Run a step with the same interpreter; raise SystemExit on failure."""
    print("» " + " ".join(os.path.relpath(a, REPO) if os.path.isabs(a) else a for a in argv))
    proc = subprocess.run(argv)
    if proc.returncode != 0:
        sys.exit("step failed (exit %d): %s" % (proc.returncode, argv[1]))
    return proc


def main():
    p = argparse.ArgumentParser(
        description="Render a gated specification DOCX from a backend OpenAPI spec.")
    p.add_argument("--openapi", required=True, help="backend docs/api/openapi.yaml")
    p.add_argument("--meta", required=True, help="backend docs/api/doc-meta.yaml (front matter)")
    p.add_argument("--out", required=True, help="output DOCX path")
    for key in DEFAULTS:
        p.add_argument("--" + key, default=DEFAULTS[key])
    p.add_argument("--workdir", help="directory for the intermediate contract.json")
    p.add_argument("--keep-contract", action="store_true", help="do not delete the intermediate contract.json")
    args = p.parse_args()

    # Validate every input up front so CI fails fast with a clear message.
    missing = [label + " not found: " + path for label, path in [
        ("openapi", args.openapi), ("meta", args.meta),
        ("skeleton", args.skeleton), ("descriptions", args.descriptions),
        ("objects", args.objects), ("planned", args.planned)] if not os.path.isfile(path)]
    if missing:
        sys.exit("missing input(s):\n  " + "\n  ".join(missing))

    out = os.path.abspath(args.out)
    workdir = os.path.abspath(args.workdir) if args.workdir else os.path.dirname(out)
    os.makedirs(workdir, exist_ok=True)
    contract = os.path.join(workdir, "contract.json")
    py = sys.executable

    run([py, os.path.join(SCRIPTS, "openapi_to_contract.py"), args.openapi, "-o", contract])
    run([py, os.path.join(SCRIPTS, "render_from_openapi.py"),
         args.skeleton, contract, out,
         "--skeleton", "--meta", args.meta,
         "--descriptions", args.descriptions,
         "--objects", args.objects, "--planned", args.planned])
    run([py, os.path.join(SCRIPTS, "check_frame_format.py"), out])

    if not args.keep_contract and os.path.isfile(contract):
        os.remove(contract)
    print("\nOK: release-ready DOCX passed the gate: " + out)


if __name__ == "__main__":
    main()
