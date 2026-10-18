#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Normalize language-code data across the GoPaperless API specification DOCX.

Implements rule 2.8.7 of scripts/GoPaperless_DOCX_Format_Template.md:

  1. Error response bodies: "language_name": "vn"  ->  "vi"
  2. Header Request Attributes description: "+ VI"  ->  "+ vi"
  3. Request frames missing the "x-language-name: en" header line get it added
     (cloned from an existing frame so formatting is identical)
  4. Legacy error frames using "language_name": "EN" + "remark_language" are
     rebuilt into the standard shape: lowercase "en" plus a "vi" element,
     both using the "remark_message" key.

Operates on unpacked_docx/word/document.xml, then repacks the DOCX.

Usage:
    python scripts/oneoff/fix_language_consistency.py [--dry-run]
"""
import os
import re
import sys
import shutil
import zipfile
from copy import deepcopy

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx"
)

# Vietnamese counterparts for the legacy error messages (rule 2.8.7: every
# remark array carries both en and vi).
VI_MESSAGES = {
    "Your client ID is invalid. Please request another client ID":
        "Client ID không hợp lệ. Vui lòng yêu cầu một client ID khác",
    "Your Refresh token is invalid.":
        "Refresh token không hợp lệ.",
}


def para_lines(el):
    """Text of every direct child paragraph of el."""
    return ["".join(t.text or "" for t in p.iter(qn("t"))) for p in el.findall(qn("p"))]


def set_para_text(p, text):
    """Force a single-run paragraph to carry exactly `text`."""
    runs = p.findall(qn("r"))
    first = None
    for r in runs:
        t = r.find(qn("t"))
        if t is None:
            continue
        if first is None:
            first = t
        else:
            r.getparent().remove(r)
    if first is None:
        raise ValueError("paragraph has no run with text: %r" % para_lines(p))
    first.text = text
    first.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


# --------------------------------------------------------------------------
# Fix 1 - "vn" -> "vi" in error response bodies
# --------------------------------------------------------------------------
def fix_vn_to_vi(root):
    """Replace the vn language code wherever it is used as a language value.

    Note: `.vn` email/domain suffixes and the `"C": "VN"` country code in
    certificate Distinguished Names are NOT language codes and stay untouched.
    """
    n = 0
    for t in root.iter(qn("t")):
        if not t.text:
            continue
        before = t.text
        t.text = t.text.replace('"language_name": "vn"', '"language_name": "vi"')
        t.text = t.text.replace('"gopaperless.workflow.app.langs": "en;vn"',
                                '"gopaperless.workflow.app.langs": "en;vi"')
        if t.text != before:
            n += 1
    return n


# --------------------------------------------------------------------------
# Fix 2 - "+ VI" -> "+ vi" in Header Request Attributes descriptions
# --------------------------------------------------------------------------
def fix_uppercase_vi(root):
    """Only touch cells that describe x-language-name possible values."""
    n = 0
    for tc in root.iter(qn("tc")):
        text = "".join(t.text or "" for t in tc.iter(qn("t")))
        if "Possible values are" not in text or "VI" not in text:
            continue
        for t in tc.iter(qn("t")):
            if t.text and re.search(r"\bVI\b", t.text):
                t.text = re.sub(r"\bVI\b", "vi", t.text)
                n += 1
    return n


# --------------------------------------------------------------------------
# Fix 3 - add missing "x-language-name: en" line to request frames
# --------------------------------------------------------------------------
LANG_LINE_RE = re.compile(r"^\s*x-language-name\s*:", re.I)
METHOD_RE = re.compile(r"\b(GET|POST|PUT|DELETE|PATCH)\s*\{baseURL\}")


def find_lang_template_para(root):
    """A paragraph from an existing frame holding exactly 'x-language-name: en'."""
    for tx in root.iter(qn("txbxContent")):
        for p in tx.findall(qn("p")):
            line = "".join(t.text or "" for t in p.iter(qn("t")))
            if line.strip() == "x-language-name: en":
                return deepcopy(p)
    return None


def fix_missing_lang_header(root):
    template = find_lang_template_para(root)
    if template is None:
        raise RuntimeError("no x-language-name template paragraph found")

    fixed = []
    for tx in root.iter(qn("txbxContent")):
        paras = tx.findall(qn("p"))
        lines = ["".join(t.text or "" for t in p.iter(qn("t"))) for p in paras]
        if not any(METHOD_RE.search(l) for l in lines):
            continue                                   # not a request frame
        if any(LANG_LINE_RE.match(l) for l in lines):
            continue                                   # already compliant

        # Insert after the last header line (a line of the form "Key: value"),
        # which keeps it immediately before the JSON body.
        insert_at = 0
        for i, l in enumerate(lines):
            if METHOD_RE.search(l):
                insert_at = i + 1
            elif re.match(r"^[A-Za-z][A-Za-z0-9\-]*\s*:", l.strip()):
                insert_at = i + 1
        paras[insert_at - 1].addnext(deepcopy(template))
        fixed.append(lines[0].strip())
    return fixed


# --------------------------------------------------------------------------
# Fix 4 - rebuild legacy error frames into the standard remark shape
# --------------------------------------------------------------------------
def fix_legacy_error_frames(root):
    fixed = []
    for tx in root.iter(qn("txbxContent")):
        paras = tx.findall(qn("p"))
        lines = ["".join(t.text or "" for t in p.iter(qn("t"))) for p in paras]
        if not any("remark_language" in l for l in lines):
            continue

        i_lang = next(i for i, l in enumerate(lines) if '"language_name"' in l)
        i_msg = next(i for i, l in enumerate(lines) if '"remark_language"' in l)
        i_close = i_msg + 1                     # the "    }" closing the remark element

        m = re.search(r'"remark_language":\s*"(.*)"\s*$', lines[i_msg])
        if m is None:
            raise ValueError("cannot parse remark_language line: %r" % lines[i_msg])
        en_msg = m.group(1).strip()             # legacy data has stray leading spaces
        vi_msg = VI_MESSAGES.get(en_msg)
        if vi_msg is None:
            raise KeyError("no Vietnamese translation for %r" % en_msg)

        if lines[i_close].strip() != "}":
            raise ValueError("unexpected closing line: %r" % lines[i_close])

        set_para_text(paras[i_lang], '      "language_name": "en",')
        set_para_text(paras[i_msg], '      "remark_message": "%s"' % en_msg)
        set_para_text(paras[i_close], "    },")

        # Append the vi element, cloning existing paragraphs so the run
        # properties (Verdana 10 pt) are preserved.
        anchor = paras[i_close]
        for text, src in (
            ("    {", paras[i_lang - 1]),
            ('      "language_name": "vi",', paras[i_lang]),
            ('      "remark_message": "%s"' % vi_msg, paras[i_msg]),
            ("    }", paras[i_close]),
        ):
            new_p = deepcopy(src)
            set_para_text(new_p, text)
            anchor.addnext(new_p)
            anchor = new_p
        fixed.append(en_msg)
    return fixed


def pack(src_dir, out_path):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml must be the first entry in an OPC package.
        ct = os.path.join(src_dir, "[Content_Types].xml")
        z.write(ct, "[Content_Types].xml")
        for base, _dirs, files in os.walk(src_dir):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                if rel == "[Content_Types].xml":
                    continue
                z.write(full, rel)
    shutil.move(tmp, out_path)


def main():
    dry = "--dry-run" in sys.argv

    parser = etree.XMLParser(huge_tree=True)
    tree = etree.parse(DOC_XML, parser)
    root = tree.getroot()

    n_vn = fix_vn_to_vi(root)
    n_vi = fix_uppercase_vi(root)
    frames = fix_missing_lang_header(root)
    legacy = fix_legacy_error_frames(root)

    print("[1] language_name vn -> vi          : %d run(s)" % n_vn)
    print("[2] description VI -> vi            : %d run(s)" % n_vi)
    print("[3] frames given x-language-name    : %d" % len(frames))
    for f in frames:
        print("      +", f)
    print("[4] legacy error frames rebuilt     : %d" % len(legacy))
    for f in sorted(set(legacy)):
        print("      *", f)

    if dry:
        print("\n--dry-run: nothing written")
        return

    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
