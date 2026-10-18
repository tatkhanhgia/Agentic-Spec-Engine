#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Add the 10 Document Workflow endpoints that exist in the backend `validURLs`
registry (and in the Postman collection) but were missing from the
specification DOCX.

Every section is built by cloning an existing donor section of the same shape
(method style, table set, frame layout) and substituting text, per the
"write by cloning" rule. The first-line run structure of the request frame is
preserved (badge run + space + red {baseURL} + alternating black/red path
runs) - it is never collapsed into a single run (format template 2.8.1).

Facts documented here come from the backend, not from the model:
  * routes/methods  - WorkflowController / CategoryController validURLs.
  * request fields  - CreateWorkflowFromTemplateRequest, AcceptTosRequest,
                      CloneWorkflowFromExcelRequest (@JsonProperty names) and
                      their ValidateRequest / ValidateTosRequest checks.
  * response fields - CreateWorkflowFromTemplateResponseDTO,
                      CloneWorkflowStatusSerializer.
  * empty success   - services return status-only InternalResponse.

Identifier hygiene on each clone: bookmarks stripped, wp:docPr ids and VML
shape ids reallocated above the current maxima.

Usage:
    python scripts/oneoff/add_document_workflow_missing_sections.py [--dry-run]
"""
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
from docx_guard import load, save  # noqa: E402

UNPACKED = os.path.join(ROOT, "unpacked_docx")
DOC_XML = os.path.join(UNPACKED, "word", "document.xml")
OUT_DOCX = os.path.join(
    ROOT, "formatted", "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")

BADGE = {"GET": "047857", "POST": "C2410C", "PUT": "1D4ED8", "DELETE": "B91C1C"}
METHODS = ("GET", "POST", "PUT", "DELETE", "PATCH")

PKG_DESC = ("This is the ID of the package that was created. Using this ID to "
            "Interact with that document package.")

# ---------------------------------------------------------------------------
# Section specifications (insertion order matters: anchors created earlier)
# ---------------------------------------------------------------------------
SPECS = [
    dict(
        name="Create Workflow From Template",
        donor="Create new Category",
        anchor="Start new Workflow",
        method="POST",
        url="{baseURL}/v1/workflow/from-template",
        prose=("Business applications can use this service API to create a new draft "
               "document package, document and workflow in one call from a complete "
               "template snapshot. The template to apply is provided in the request body."),
        body_lines=["{", '  "template_id": 12,', '  "package_name": "Contract Package"', "}"],
        request_rows=[
            ("template_id", "Integer", "M",
             "The ID of the template used to create the new workflow."),
            ("package_name", "String", "O",
             "Name of the new document package. Defaults to the template name when omitted."),
        ],
        response_rows=[
            ("package_id", "Integer", "The ID of the document package that was created."),
            ("workflow_id", "Integer", "The ID of the workflow that was created."),
            ("document_id", "Integer", "The ID of the document that was created."),
            ("template_id", "Integer", "The ID of the template that was applied."),
        ],
        response_sample=["{", '  "package_id": 118,', '  "workflow_id": 92,',
                         '  "document_id": 240,', '  "template_id": 12', "}"],
    ),
    dict(
        name="Share Workflow",
        donor="Delete Workflow Comment/Message",
        anchor="Create Workflow From Template",
        method="POST",
        url="{baseURL}/v1/packages/{packageId}/workflow/share",
        prose=("Business applications can use this service API to share the document "
               "package with the signers and start the workflow. The package ID to be "
               "shared is provided in the resource URL."),
        path_rows=[("packageId", "Integer", "M", PKG_DESC)],
    ),
    dict(
        name="Delete Workflow",
        donor="Delete Workflow Comment/Message",
        anchor="Update Workflow Details",
        method="DELETE",
        url="{baseURL}/v1/packages/{packageId}/workflow",
        prose=("Business applications can use this service API to stop a workflow in "
               "progress and archive it. The package ID is provided in the resource URL."),
        path_rows=[("packageId", "Integer", "M", PKG_DESC)],
    ),
    dict(
        name="Get Workflow Details by UUID",
        donor="Get Workflow Details",
        anchor="Get Workflow Details",
        method="GET",
        url="{baseURL}/v1/workflow/{workflowUuid}",
        prose=("Business applications can use this service API to retrieve the details "
               "of a workflow by its UUID, as an alternative to the package ID based "
               "lookup. The response payload is identical to Get Workflow Details."),
        path_rows=[("workflowUuid", "String", "M", "The UUID of the workflow to retrieve.")],
    ),
    dict(
        name="View Workflow",
        donor="Get Workflow Details",
        anchor="Get Workflow Details by UUID",
        method="POST",
        url="{baseURL}/v1/workflow/{uuid}/view",
        prose=("Applications can call this service API when a shared document is opened "
               "by a recipient. It returns the workflow details for an unauthenticated "
               "recipient and records the view action. No access token is required."),
        path_rows=[("uuid", "String", "M", "The UUID of the workflow shared to the recipient.")],
        header_remove=["Authorization"],
    ),
    dict(
        name="Accept Terms of Service (ECONTRACT)",
        donor="Update Placeholder in Workflow",
        anchor="Delete Workflow User",
        method="PUT",
        url="{baseURL}/v1/packages/{packageId}/workflow/{participantId}/tos-acceptance",
        prose=("Business applications can use this service API to record the Terms of "
               "Service acceptance evidence of an ECONTRACT workflow participant. The "
               "package ID and the workflow participant ID are provided in the resource URL."),
        path_rows=[
            ("packageId", "Integer", "M", PKG_DESC),
            ("participantId", "Integer", "M",
             "The ID of the workflow participant accepting the Terms of Service."),
        ],
        body_lines=["{", '  "ip": "203.0.113.42",',
                    '  "ua": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",',
                    '  "ts": "1741012345",', '  "act": "accept"', "}"],
        request_rows=[
            ("ip", "String", "M",
             "Client IP address. The server always overrides this value with the "
             "server-verified IP address to prevent spoofing."),
            ("ua", "String", "M", "User-Agent string of the client that accepted the terms."),
            ("ts", "String", "M", "Unix timestamp in seconds of the acceptance."),
            ("act", "String", "M", "Action identifier of the consent, e.g. \"accept\"."),
        ],
    ),
    dict(
        name="Send Reminder",
        donor="Delete Workflow Comment/Message",
        anchor="Get Workflow Reminders",
        method="POST",
        url="{baseURL}/v1/packages/{packageId}/workflow/{order}/remind",
        prose=("Business applications can use this service API to send a reminder to a "
               "workflow user. The recipient is identified by the order in the workflow "
               "and must be a next signer, otherwise the request is rejected."),
        path_rows=[
            ("packageId", "Integer", "M", PKG_DESC),
            ("order", "Integer", "M", "Order of the recipient in the workflow."),
        ],
    ),
    dict(
        name="Import Excel to Clone Workflow",
        donor="Update Placeholder in Workflow",
        anchor="Update Placeholder in Workflow",
        method="POST",
        url="{baseURL}/v1/packages/{packageId}/workflow/{order}/import",
        prose=("Business applications can use this service API to bulk-clone the "
               "workflow for a list of recipients imported from an Excel file. The "
               "target placeholder is identified by the order at which it is added to "
               "the workflow. The operation runs asynchronously; use Get Status of "
               "Import Excel to poll the progress."),
        path_rows=[
            ("packageId", "Integer", "M", PKG_DESC),
            ("order", "Integer", "M", "Order of the placeholder in the workflow."),
        ],
        body_lines=["{", '  "items": [', "    {",
                    '      "username": "Participant1",',
                    '      "user_email": "participant1@example.com",',
                    '      "personal_code": "097200011188"', "    }", "  ],",
                    '  "webhook_url": "https://example.com/webhook",',
                    '  "share_enabled": true', "}"],
        request_rows=[
            ("items", "Array", "M", "List of recipients to clone the workflow for."),
            ("items[].username", "String", "M", "Display name of the recipient."),
            ("items[].user_email", "String", "M",
             "Email of the recipient. Required when the target placeholder has email "
             "notification enabled."),
            ("items[].mobile_number", "String", "M",
             "Mobile number of the recipient. Required when the target placeholder has "
             "email notification disabled."),
            ("items[].personal_code", "String", "O", "Personal identification code of the recipient."),
            ("items[].vneid_fullname", "String", "O", "Full name from the VNeID account."),
            ("items[].vneid_birthday", "String", "O", "Birthday from the VNeID account."),
            ("items[].vneid_gender", "String", "O", "Gender from the VNeID account."),
            ("items[].vneid_current_place", "String", "O", "Current residence from the VNeID account."),
            ("items[].vneid_place", "String", "O", "Registered residence from the VNeID account."),
            ("items[].vneid_level_account", "Integer", "O", "Account level from the VNeID account."),
            ("items[].vneid_type_account", "String", "O", "Account type from the VNeID account."),
            ("items[].vneid_date_issuer", "String", "O", "Issue date from the VNeID account."),
            ("items[].vneid_date_expired", "String", "O", "Expiry date from the VNeID account."),
            ("webhook_url", "String", "O", "Callback URL notified when the clone operation completes."),
            ("share_enabled", "Boolean", "O", "True to automatically share the cloned workflows."),
        ],
    ),
    dict(
        name="Get Status of Import Excel",
        donor="Get Workflow Comment/Message",
        anchor="Import Excel to Clone Workflow",
        method="GET",
        url="{baseURL}/v1/workflow/{transactionCode}/import",
        prose=("Business applications can use this service API to poll the progress of "
               "a bulk clone-workflow-from-Excel operation. The transaction code is "
               "provided in the resource URL."),
        path_rows=[("transactionCode", "String", "M",
                    "The transaction code returned when the Import Excel to Clone "
                    "Workflow request was submitted.")],
        response_rows=[
            ("status", "String",
             "Overall status of the import. Possible values are IN_PROGRESS, FAIL, "
             "NOT_FOUND and SUCCESS."),
            ("total_workflows", "Integer", "Total number of workflows to clone."),
            ("workflow_success", "Integer", "Number of workflows cloned successfully."),
            ("workflow_fail", "Integer", "Number of workflows that failed to clone."),
            ("workflow_inprogress", "Integer", "Number of workflows still processing."),
            ("percentage", "Integer", "Completion percentage of the operation."),
            ("results", "Array",
             "Per-workflow result: id (package ID), workflow_id, name, action "
             "(processing step), status (Boolean) and error (resolved error message)."),
        ],
        response_sample=["{", '  "status": "IN_PROGRESS",', '  "total_workflows": 10,',
                         '  "workflow_success": 4,', '  "workflow_fail": 1,',
                         '  "workflow_inprogress": 5,', '  "percentage": 50,',
                         '  "results": [', "    {", '      "id": 118,',
                         '      "workflow_id": 92,', '      "name": "Workflow 1",',
                         '      "status": true', "    }", "  ]", "}"],
    ),
    dict(
        name="Delete Category",
        donor="Delete Workflow Comment/Message",
        anchor="Get list of Category",
        method="DELETE",
        url="{baseURL}/v1/category/{categoryId}",
        prose=("Business applications can use this service API to delete a category "
               "used to organize workflows. The category ID is provided in the "
               "resource URL."),
        path_rows=[("categoryId", "Integer", "M", "The ID of the category to be deleted.")],
    ),
]


# ---------------------------------------------------------------------------
# helpers (same approach as add_switch_enterprise_context.py)
# ---------------------------------------------------------------------------
def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("t")))


def style_of(p):
    pr = p.find(qn("pPr"))
    if pr is None:
        return ""
    st = pr.find(qn("pStyle"))
    return (st.get(qn("val")) or "") if st is not None else ""


def set_para_text(p, text):
    """Reduce a paragraph to one run carrying `text`, keeping its formatting."""
    pPr = p.find(qn("pPr"))
    template = None
    for r in p.iter(qn("r")):
        if r.find(qn("t")) is not None:
            template = r
            break
    if template is None:
        raise ValueError("paragraph has no text run: %r" % text_of(p))
    new_run = deepcopy(template)
    for child in list(new_run):
        if child.tag not in (qn("rPr"), qn("t")):
            new_run.remove(child)
    for extra in new_run.findall(qn("t"))[1:]:
        new_run.remove(extra)
    t = new_run.find(qn("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    for child in list(p):
        if child is not pPr:
            p.remove(child)
    p.append(new_run)


def cell_text(tc):
    return " ".join(x for x in (text_of(p).strip() for p in tc.findall(qn("p"))) if x).strip()


def set_cell_text(tc, text):
    paras = tc.findall(qn("p"))
    set_para_text(paras[0], text)
    for extra in paras[1:]:
        tc.remove(extra)


def table_title(tbl):
    trs = tbl.findall(qn("tr"))
    if not trs:
        return ""
    return " ".join(cell_text(tc) for tc in trs[0].findall(qn("tc"))).strip()


def find_heading(kids, level, title):
    want = "Heading%d" % level
    for i, e in enumerate(kids):
        if e.tag == qn("p") and style_of(e) == want and text_of(e).strip() == title:
            return i
    raise LookupError("heading not found: %s" % title)


def section_range(kids, start):
    for j in range(start + 1, len(kids)):
        e = kids[j]
        if e.tag == qn("p") and re.match(r"Heading\d$", style_of(e)):
            return start, j
    return start, len(kids)


def strip_bookmarks(elements):
    n = 0
    for el in elements:
        for tag in ("bookmarkStart", "bookmarkEnd"):
            for b in list(el.iter(qn(tag))):
                b.getparent().remove(b)
                n += 1
    return n


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def reassign_ids(elements, next_docpr, next_spid):
    n_doc = n_vml = 0
    for el in elements:
        for d in el.iter():
            if local(d) == "docPr":
                d.set("id", str(next_docpr))
                d.set("name", "Text Box %d" % next_docpr)
                next_docpr += 1
                n_doc += 1
        for d in el.iter():
            if local(d) == "shape":
                d.set("id", "_x0000_s%d" % next_spid)
                if d.get("{%s}spid" % O) is not None:
                    d.set("{%s}spid" % O, "_x0000_s%d" % next_spid)
                next_spid += 1
                n_vml += 1
    return n_doc, n_vml, next_docpr, next_spid


def scan_id_space(root):
    docpr = [int(d.get("id")) for d in root.iter()
             if local(d) == "docPr" and (d.get("id") or "").isdigit()]
    spids = []
    for d in root.iter():
        if local(d) in ("shape", "shapetype"):
            for v in (d.get("id"), d.get("{%s}spid" % O)):
                m = re.match(r"_x0000_s(\d+)$", v or "")
                if m:
                    spids.append(int(m.group(1)))
    return (max(docpr) + 1 if docpr else 1), (max(spids) + 1 if spids else 1000)


# ---------------------------------------------------------------------------
# frame editing - preserves the multi-run first line (template 2.8.1)
# ---------------------------------------------------------------------------
def run_text(r):
    return "".join(t.text or "" for t in r.findall(qn("t")))


def run_color(r):
    pr = r.find(qn("rPr"))
    if pr is None:
        return ""
    c = pr.find(qn("color"))
    return (c.get(qn("val")) or "") if c is not None else ""


def make_run(template, text):
    r = deepcopy(template)
    ts = r.findall(qn("t"))
    ts[0].text = text
    ts[0].set(XML_SPACE, "preserve")
    for extra in ts[1:]:
        r.remove(extra)
    return r


def rewrite_url_line(p, method, url):
    """Rewrite `METHOD {baseURL}/path` keeping the badge/space/red/black runs."""
    runs = p.findall(qn("r"))
    method_run = None
    for r in runs:
        pr = r.find(qn("rPr"))
        if pr is not None and pr.find(qn("shd")) is not None and run_text(r) in METHODS:
            method_run = r
            break
    if method_run is None:
        raise RuntimeError("URL line without a method badge run: %r" % text_of(p))
    red_tpl = next(r for r in runs if run_color(r) == "FF0000")
    black_tpl = next(r for r in runs
                     if run_color(r) == "000000" and run_text(r) not in ("", " "))

    # badge text + fill
    t = method_run.find(qn("t"))
    t.text = method
    method_run.find(qn("rPr")).find(qn("shd")).set(qn("fill"), BADGE[method])

    idx = runs.index(method_run)
    space_run = runs[idx + 1] if idx + 1 < len(runs) and run_text(runs[idx + 1]) == " " else None
    if space_run is None:
        space_run = make_run(black_tpl, " ")
        method_run.addnext(space_run)
    # drop everything after the space run
    for r in runs:
        if r is method_run or r is space_run:
            continue
        if r.getparent() is p:
            p.remove(r)

    if not url.startswith("{baseURL}"):
        raise RuntimeError("url must start with {baseURL}: %s" % url)
    anchor = space_run
    for token in re.findall(r"\{[^}]+\}|[^{]+", url):
        tpl = red_tpl if token.startswith("{") else black_tpl
        new = make_run(tpl, token)
        anchor.addnext(new)
        anchor = new


def edit_request_frame(txbx, spec):
    paras = txbx.findall(qn("p"))
    rewrite_url_line(paras[0], spec["method"], spec["url"])
    # optional header line removal (e.g. Authorization for unauthenticated APIs)
    for hdr in spec.get("header_remove", []):
        for p in txbx.findall(qn("p")):
            if text_of(p).strip().startswith(hdr):
                txbx.remove(p)
    # body: replace from the first '{'-line to the end
    paras = txbx.findall(qn("p"))
    lines = [text_of(p) for p in paras]
    body_start = next((i for i, l in enumerate(lines[1:], 1) if l.strip().startswith("{")), None)
    if body_start is None:
        raise RuntimeError("request frame without a body line: %s" % spec["name"])
    src = paras[body_start]
    anchor = paras[body_start - 1]
    for p in paras[body_start:]:
        txbx.remove(p)
    for line in spec.get("body_lines", ["{}"]):
        new_p = deepcopy(src)
        set_para_text(new_p, line)
        anchor.addnext(new_p)
        anchor = new_p


def replace_frame_lines(txbx, lines):
    paras = txbx.findall(qn("p"))
    src = paras[0]
    anchor = None
    for p in paras:
        txbx.remove(p)
    for line in lines:
        new_p = deepcopy(src)
        set_para_text(new_p, line)
        if anchor is None:
            txbx.insert(0, new_p)
        else:
            anchor.addnext(new_p)
        anchor = new_p


# ---------------------------------------------------------------------------
# table editing
# ---------------------------------------------------------------------------
def resize_fill(tbl, rows, with_presence=True):
    """rows: list of tuples (name, type[, presence], desc); renumbers column 0."""
    trs = tbl.findall(qn("tr"))
    template = deepcopy(trs[2])
    for tr in trs[2:]:
        tbl.remove(tr)
    anchor = trs[1]
    for i, row in enumerate(rows):
        tr = deepcopy(template)
        tcs = tr.findall(qn("tc"))
        if with_presence:
            name, typ, presence, desc = row
            values = [str(i + 1), name, typ, presence, desc]
        else:
            name, typ, desc = row
            values = [str(i + 1), name, typ, desc]
            if len(tcs) == 5:
                values = [str(i + 1), name, typ, "", desc]
        for tc, val in zip(tcs, values):
            set_cell_text(tc, val)
        anchor.addnext(tr)
        anchor = tr


def remove_header_rows(tbl, names):
    trs = tbl.findall(qn("tr"))
    for tr in trs[2:]:
        tcs = tr.findall(qn("tc"))
        if len(tcs) > 1 and cell_text(tcs[1]) in names:
            tbl.remove(tr)
    # renumber
    for i, tr in enumerate(tbl.findall(qn("tr"))[2:]):
        set_cell_text(tr.findall(qn("tc"))[0], str(i + 1))


# ---------------------------------------------------------------------------
# clone assembly
# ---------------------------------------------------------------------------
def first_line_of_first_frame(el):
    for txbx in el.iter(qn("txbxContent")):
        paras = txbx.findall(qn("p"))
        if paras:
            return text_of(paras[0]).strip()
    return None


def edit_clone(clone, spec):
    set_para_text(clone[0], spec["name"])
    if clone[1].tag != qn("p") or not text_of(clone[1]).strip():
        raise RuntimeError("donor prose paragraph not where expected: %s" % spec["donor"])
    set_para_text(clone[1], spec["prose"])

    response_done = False
    req_frames = 0
    for el in clone:
        first = first_line_of_first_frame(el)
        if first is None:
            continue
        if any(first.startswith(m + " ") for m in METHODS):
            for txbx in el.iter(qn("txbxContent")):
                edit_request_frame(txbx, spec)
                req_frames += 1
        elif spec.get("response_sample") and not response_done:
            for txbx in el.iter(qn("txbxContent")):
                replace_frame_lines(txbx, spec["response_sample"])
            response_done = True
    if req_frames == 0:
        raise RuntimeError("no request frame found: %s" % spec["name"])

    for el in list(clone):
        if el.tag != qn("tbl"):
            continue
        title = table_title(el)
        if title == "Path Attributes" and spec.get("path_rows") is not None:
            resize_fill(el, spec["path_rows"])
        elif title == "Request Attributes" and spec.get("request_rows") is not None:
            resize_fill(el, spec["request_rows"])
        elif title == "Response Attributes" and spec.get("response_rows") is not None:
            resize_fill(el, spec["response_rows"], with_presence=False)
        elif title == "Header Request Attributes" and spec.get("header_remove"):
            remove_header_rows(el, spec["header_remove"])
    return req_frames


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
    tree, token = load(DOC_XML)
    root = tree.getroot()
    body = root.find(qn("body"))

    next_docpr, next_spid = scan_id_space(root)

    for spec in SPECS:
        kids = list(body)
        if any(e.tag == qn("p") and style_of(e) == "Heading3"
               and text_of(e).strip() == spec["name"] for e in kids):
            print("SKIP (exists): %s" % spec["name"])
            continue
        d_start, d_end = section_range(kids, find_heading(kids, 3, spec["donor"]))
        a_start, a_end = section_range(kids, find_heading(kids, 3, spec["anchor"]))

        clone = [deepcopy(e) for e in kids[d_start:d_end]]
        n_bm = strip_bookmarks(clone)
        n_doc, n_vml, next_docpr, next_spid = reassign_ids(clone, next_docpr, next_spid)
        n_frames = edit_clone(clone, spec)

        anchor = kids[a_end - 1]
        for el in clone:
            anchor.addnext(el)
            anchor = el
        print("ADDED %-40s donor=%-32s after=%-32s els=%d frames=%d bm=%d docPr=%d vml=%d"
              % (spec["name"][:40], spec["donor"][:32], spec["anchor"][:32],
                 len(clone), n_frames, n_bm, n_doc, n_vml))

    if dry:
        print("\n--dry-run: nothing written")
        return
    save(tree, DOC_XML, token)
    pack(UNPACKED, OUT_DOCX)
    print("packed %s" % OUT_DOCX)


if __name__ == "__main__":
    main()
