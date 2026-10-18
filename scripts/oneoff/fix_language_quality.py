#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Proofreading pass over the GoPaperless API specification DOCX.

Fixes six classes of defect found in the source document (all pre-existing in
`original/`, none introduced by the formatting scripts):

  1. Grammar in the repeated error_description string.
  2. Spelling mistakes.
  3. Leaked product names from other services (SigningHub, SH Admin, FPS).
  4. Broken grammar in API descriptions.
  5. Punctuation and missing spaces.
  6. Wrong content (case, merged sentences, Date/Data mix-ups, heading wording).

Replacements are applied on the concatenated text of each leaf paragraph and
written back run by run, so runs that are not part of a match - hyperlinks such
as the "here" cross-references, red `{placeholder}` runs, bold header keys -
keep their formatting.

Usage:
    python scripts/oneoff/fix_language_quality.py [--dry-run]
"""
import os
import sys
import shutil
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx"
)

# Each rule is (group, old, new). Order matters: longer, more specific strings
# must come before the shorter ones they contain.
RULES = [
    # ---- 1. grammar in the repeated error_description -------------------
    (1, "The token is not belongs to GoPaperless Workflow Service!",
        "The token does not belong to GoPaperless Workflow Service!"),

    # ---- 3. leaked product names (before the grammar rewrites that
    #         touch the same sentences) -----------------------------------
    (3, "inside SigningHub Engine (ADSS Server)", "inside the GoPaperless Workflow Engine"),
    (3, "SigningHub will not allow the end users",
        "GoPaperless Workflow will not allow the end users"),
    (3, "Default timezone from SH Admin is followed the the time.",
        "The default timezone configured in GoPaperless Workflow Administration is applied to this time."),
    (3, "the FPS will generate QRCode", "the GoPaperless Workflow Service will generate QRCode"),
    (3, "your Paperless Workflow Administration client TLS certificate",
        "your GoPaperless Workflow Administration client TLS certificate"),

    # ---- 4. broken grammar (specific sentences first) -------------------
    (4, "Business Applications cans use this service API this rename to name of document which is inside the package.",
        "Business Applications can use this service API to rename a document inside the package."),
    (4, "cans use this service API", "can use this service API"),
    (4, "This document subscriber in detail the API", "This document describes in detail the API"),
    (4, "customers/companies whose would like to", "customers/companies who would like to"),
    (4, "Using this API to all the category belongs to the user.",
        "Using this API to get all the categories that belong to the user."),
    (4, "Using this API to create new category to filter that workflow.",
        "Using this API to create a new category to filter that workflow."),
    (4, "Only the owner is an enterprise account can sign Eseal.",
        "Only an owner with an enterprise account can sign an Eseal."),
    (4, "Define the name of to find in system.", "Define the name to find in the system."),
    (4, "The string name (after encode base64) uses to find the workflow label(in services)",
        "The string name (after being encoded to base64) used to find the workflow label (in services)"),
    (4, "This URL does not make sure that signer image is exist. If image exist it will be provided by hitting the defined url otherwise it will be an empty string.",
        "This URL does not guarantee that the signer image exists. If the image exists it will be provided by hitting the defined URL; otherwise it will be an empty string."),
    (4, "Is this comment is private", "True if this comment is private"),
    (4, "Is this comment can read", "True if this comment can be read"),
    (4, "if the field can be interact", "if the field can be interacted with"),
    (4, "The flag to define if the field can be have more than 2 check in one group.",
        "The flag to define if the field can have more than 2 checks in one group."),
    (4, "Email of the user that you wants to update.", "Email of the user that you want to update."),
    (4, "Level of Assurance that user wants to created.",
        "Level of Assurance that the user wants to create."),
    (4, "The OTP will be send to Application.", "The OTP will be sent to the Application."),
    (4, "If not provided, system will assign an unique auto generated name to the field.",
        "If not provided, the system will assign a unique auto-generated name to the field."),
    (4, "User may processed the field as", "User may have processed the field as"),
    (4, "If this workflow have an comment => True", "True if this workflow has a comment"),
    (4, "Group name that the user will be removed", "Group name that the user will be removed from"),
    (4, "Size if the document in bytes", "Size of the document in bytes"),
    (4, "to update the information of an user.", "to update the information of a user."),
    (4, "to update an checkbox in the workflow.", "to update a checkbox in the workflow."),
    (4, "to update an radiobox in the workflow.", "to update a radiobox in the workflow."),
    (4, "to add a initial field into the workflow.", "to add an initial field into the workflow."),
    (4, "to add an new address book", "to add a new address book"),
    (4, "to delete a member into the group.", "to delete a member from the group."),
    (4, "to an add in-person signature field", "to add an in-person signature field"),
    (4, "to approve the document before sign it.", "to approve the document before signing it."),
    (4, "to perform signing procession", "to perform the signing process"),
    (4, "If user wanna send comments to multiple recipients",
        "If the user wants to send comments to multiple recipients"),
    (4, "it contains authentication releated options", "It contains authentication related options"),
    (4, "provide group name for radio box", "Provide the group name for the radio box"),
    (4, "Certificate encoded. When provided this parameter. System will be analysied this certificate based on this parameter.",
        "Encoded certificate. When this parameter is provided, the system will analyze the certificate based on it."),
    (4, "to pre check a document before call subsequence APIs.",
        "to pre-check a document before calling subsequent APIs."),
    (4, "to get a list of user belong to this Enterprise.",
        "to get a list of users belonging to this Enterprise."),
    (4, "to get all document package in GoPaperless Workflow.",
        "to get all document packages in GoPaperless Workflow."),
    (4, "to check the status of signature field is signed.",
        "to check whether the signature field has been signed."),
    (4, "to get USB-Token that need to perform Signing USB Token (2 step signing).",
        "to get the USB-Token needed to perform USB Token signing (2-step signing)."),
    (4, "to get Smart-ID that need to perform Signing.",
        "to get the Smart-ID needed to perform signing."),
    (4, "to get ISS that need to perform Signing.", "to get the ISS needed to perform signing."),
    (4, "to get RSSP that need to perform Signing.", "to get the RSSP needed to perform signing."),
    (4, "Note that: the user must send the reason.", "Note: the user must send a reason."),
    (4, "The structure based on these example", "The structure is based on this example"),
    (4, "to get a list of user that belongs to the role",
        "to get a list of users that belong to the role"),
    (4, "to get the status of Signing when the credential mode “TSE”.",
        "to get the signing status when the credential mode is “TSE”."),

    # ---- 2. spelling ----------------------------------------------------
    (2, "Get Worklow Users", "Get Workflow Users"),
    (2, "Add Initital Field", "Add Initial Field"),
    (2, "administartor", "administrator"),
    (2, "records perpage", "records per page"),
    (2, "Extensio of the file/image", "Extension of the file/image"),
    (2, "Aligment of the text in appearance", "Alignment of the text in appearance"),
    (2, "Level of the certificate to fillter", "Level of the certificate to filter"),
    (2, "authoriation server", "authorization server"),
    (2, "to get the verificate result", "to get the verification result"),

    # ---- 6. wrong content ----------------------------------------------
    (6, "Top location of the field in pixel. Level Of Assurance to be added.",
        "Level Of Assurance to be added."),
    (6, "Data verification of the field True, if workflow", "True, if workflow"),
    (6, "+ enABLED", "+ ENABLED"),
    (6, "+DISABLED", "+ DISABLED"),
    (6, "Using “application/json” to upload your document with binary format.",
        "Using “application/json” to upload your document with base64 format."),
    (6, "Date of the camera field", "Data of the camera field"),
    (6, "Date of the attachment field", "Data of the attachment field"),
    (6, "ERROR CODE HTTP", "HTTP ERROR CODE"),
    (6, "Account Usage Statistic", "Account Usage Statistics"),
    (6, "Document Statistic", "Document Statistics"),

    # ---- 5. punctuation and spacing (last: earlier rewrites may already
    #         have removed some of these) ---------------------------------
    (5, "provided by the recipient..", "provided by the recipient."),
    (5, "are to be enabled..", "are to be enabled."),
    (5, "provided by user .", "provided by user."),
    (5, "workflow.Hence further work", "workflow. Hence further work"),
    (5, "in this way.The order number", "in this way. The order number"),
    (5, "the workflow.Therefore", "the workflow. Therefore"),
    (5, "workflow.They must either", "workflow. They must either"),
    (5, "same location.This type", "same location. This type"),
    (5, "can be changed.Where you wish", "can be changed. Where you wish"),
    (5, "unchanged.You can therefore", "unchanged. You can therefore"),
    (5, "Get Document Fields.The latter", "Get Document Fields. The latter"),
    (5, "returned.If you do not set", "returned. If you do not set"),
    (5, "In case of \"false\"only", "In case of \"false\" only"),
    (5, "registration,if password", "registration, if password"),
    (5, "Authentication(Document Opening)", "Authentication (Document Opening)"),
    (5, "List of email of  the participant", "List of email of the participant"),
    # "QR Code." / "Note:" only look joined in a whole-document text dump -
    # they are two separate paragraphs, so there is no missing space to fix.
]

GROUP_NAMES = {
    1: "grammar in error_description",
    2: "spelling",
    3: "other-service names",
    4: "grammar",
    5: "punctuation / spacing",
    6: "content",
}


def leaf_paragraphs(root):
    """Paragraphs that do not themselves contain other paragraphs.

    A textbox anchor paragraph contains the frame's paragraphs, so iterating
    every `w:p` would visit the same text twice.
    """
    txbx = ".//" + qn("txbxContent")
    for p in root.iter(qn("p")):
        # The only paragraph-in-paragraph nesting in this document is a textbox
        # anchor. A single C-level find beats walking each subtree in Python.
        if p.find(txbx) is not None:
            continue
        yield p


def apply_rules_to_paragraph(p, rules, counts):
    """Apply every rule to one paragraph, rewriting only the runs it touches."""
    ts = [t for t in p.iter(qn("t"))]
    if not ts:
        return
    texts = [t.text or "" for t in ts]
    changed = False
    # Joining is the expensive part on a 27 MB document, so keep one cached
    # copy and refresh it only after a replacement actually lands.
    joined = "".join(texts)

    for key, old, new in rules:
        # Resume the search after the text just written. Several replacements
        # contain their own search string ("Statistic" -> "Statistics",
        # "interact" -> "interacted with"); restarting at 0 would rematch the
        # result and loop forever.
        start = 0
        while True:
            pos = joined.find(old, start)
            if pos == -1:
                break
            end = pos + len(old)

            # Map the character span onto the runs that carry it.
            starts = []
            acc = 0
            for s in texts:
                starts.append(acc)
                acc += len(s)
            i0 = max(i for i in range(len(texts)) if starts[i] <= pos)
            i1 = max(i for i in range(len(texts)) if starts[i] < end)

            head = texts[i0][:pos - starts[i0]]
            tail = texts[i1][end - starts[i1]:]
            if i0 == i1:
                texts[i0] = head + new + tail
            else:
                texts[i0] = head + new
                for i in range(i0 + 1, i1):
                    texts[i] = ""
                texts[i1] = tail
            counts[key] = counts.get(key, 0) + 1
            changed = True
            joined = "".join(texts)
            start = pos + len(new)

    if changed:
        for t, s in zip(ts, texts):
            t.text = s
            t.set(XML_SPACE, "preserve")


def pack(src_dir, out_path):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(src_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for base, _dirs, files in os.walk(src_dir):
            for f in sorted(files):
                full = os.path.join(base, f)
                rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
                if rel != "[Content_Types].xml":
                    z.write(full, rel)
    shutil.move(tmp, out_path)


def main():
    dry = "--dry-run" in sys.argv
    tree = etree.parse(DOC_XML, etree.XMLParser(huge_tree=True))
    root = tree.getroot()

    indexed = [((g, old), old, new) for g, old, new in RULES]
    counts = {}
    n_paras = 0
    for p in leaf_paragraphs(root):
        apply_rules_to_paragraph(p, indexed, counts)
        n_paras += 1
    print("scanned %d leaf paragraphs" % n_paras)

    by_group = {}
    misses = []
    for g, old, new in RULES:
        n = counts.get((g, old), 0)
        by_group[g] = by_group.get(g, 0) + n
        if n == 0:
            misses.append((g, old))

    for g in sorted(GROUP_NAMES):
        print("group %d (%-28s): %4d replacement(s)" % (g, GROUP_NAMES[g], by_group.get(g, 0)))
    print("total: %d replacement(s) from %d rules" % (sum(counts.values()), len(RULES)))

    if misses:
        print("\nRULES THAT MATCHED NOTHING (%d):" % len(misses))
        for g, old in misses:
            print("   [group %d] %r" % (g, old[:70]))

    if dry:
        print("\n--dry-run: nothing written")
        return
    tree.write(DOC_XML, xml_declaration=True, encoding="UTF-8", standalone=True)
    pack(UNPACKED, OUT_DOCX)
    print("\nwrote %s" % DOC_XML)
    print("wrote %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
