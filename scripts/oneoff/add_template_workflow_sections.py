#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Add the seven Template workflow-participant endpoints to the specification.

Every section is built by cloning an existing section that already has the
target shape (CLAUDE.md step 4) - none of the WordprocessingML is hand-written.
The donors all live under `Document Workflow`, which serves the same six
operations against a package instead of a template.

One thing the clone must NOT carry across: the package endpoints wrap their bulk
payload in a named object, `{"users": [...]}`. The template endpoints take a
root-level JSON array, which is what the backend parses and what the SigningHub
10.0 manual specifies (`Request Body schema: Array`). The request frame and the
request table are therefore rebuilt rather than edited.

Run from the project root. Writes unpacked_docx/word/document.xml through
docx_guard, which refuses to save if another agent changed the file meanwhile.
"""
import os
import shutil
import sys
from copy import deepcopy

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
ROOT = os.path.dirname(SCRIPTS)
sys.path.insert(0, SCRIPTS)

import section_tools as st          # noqa: E402
import docx_guard                   # noqa: E402

DOC = os.path.join(ROOT, "unpacked_docx", "word", "document.xml")
UNPACKED = os.path.join(ROOT, "unpacked_docx")
OUT_DOCX = os.path.join(
    ROOT, "formatted",
    "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
BACKUP = OUT_DOCX.replace(".docx", "_backup_pre_template_workflow.docx")

TEMPLATE_ID_DESC = ("This is the ID of the template. Use this ID to interact "
                    "with that template.")
ORDER_DESC = ("The position of the participant in the template workflow, as "
              "returned in the order attribute of Get Template Workflow Users.")

# The participant object every one of these endpoints returns. Only the
# attributes TemplateWorkflowParticipantSerializer actually writes are listed:
# the DTO also declares delivery_method and logged_in_user, but the serializer
# never sets them, so they never appear in a response.
PARTICIPANT_ROWS = [
    ("order", "Integer", "M", "The position of the participant in the template workflow."),
    ("user_email", "String", "O", "Email address of the participant. Returned only for a real user, not for a placeholder."),
    ("user_name", "String", "O", "Full name of the participant. Returned only for a real user, not for a placeholder."),
    ("placeholder", "String", "O", "Name of the placeholder. Returned only when the participant is a placeholder, in which case user_email and user_name are absent."),
    ("role", "String", "M", "The role of the participant in the workflow. Possible values are: + SIGNER + REVIEWER + EDITOR + INPERSON_HOST + CARBON_COPY"),
    ("signing_order", "Integer", "O", "The order in which this participant signs. Applies when the workflow type is CUSTOM."),
    ("guest_user", "Boolean", "M", "True if the participant is not a registered user of the system."),
    ("email_language_code", "String", "M", "The language used for notification emails sent to this participant. Possible values are: + en + vi"),
    ("email_notification", "Boolean", "O", "If true, the participant is notified by email."),
    ("mobile_number", "String", "O", "Mobile number of the participant."),
]

USER_REQUEST_ROWS = [
    ("user_email", "String", "M", "Email address of the user to be added to the template workflow."),
    ("user_name", "String", "M", "Name of the recipient to be added to the template workflow."),
    ("email_notification", "Boolean", "O", "If true, the user is notified by email when the workflow runs."),
    ("role", "String", "M", "The role of the user in the workflow. Possible values are: + SIGNER + REVIEWER + EDITOR + INPERSON_HOST + CARBON_COPY"),
    ("signing_order", "Integer", "O", "The order in which this user signs. Mandatory when the workflow type is CUSTOM."),
    ("email_language_code", "String", "O", "The language used for notification emails. Possible values are: + en + vi"),
    ("delivery_method", "String", "O", "How the document is delivered to the user. Possible values are: + EMAIL + SMS + EMAIL_AND_SMS"),
    ("mobile_number", "String", "O", "Mobile number of the user. Required when the delivery method includes SMS."),
]

PLACEHOLDER_REQUEST_ROWS = [
    ("placeholder", "String", "M", "Name of the placeholder to be added to the template workflow."),
    ("email_notification", "Boolean", "O", "If true, the participant who replaces this placeholder is notified by email."),
    ("role", "String", "M", "The role of the placeholder in the workflow. Possible values are: + SIGNER + REVIEWER + EDITOR + INPERSON_HOST + CARBON_COPY"),
    ("signing_order", "Integer", "O", "The order in which this placeholder signs. Mandatory when the workflow type is CUSTOM."),
]

ERROR_ROWS = [
    ("error_description", "String", "M/O", "Detailed description of the error code."),
    ("error_code", "String", "M/O", "Code identifying the error."),
    ("message", "String", "M/O", "Error message displayed to the end user."),
]

USER_SAMPLE = [
    '"user_email": "giatk@mobile-id.vn",',
    '"user_name": "TatKhanhGia",',
    '"email_notification": true,',
    '"role": "SIGNER",',
    '"signing_order": 1,',
    '"email_language_code": "en",',
    '"delivery_method": "EMAIL",',
    '"mobile_number": "0912345678"',
]

PLACEHOLDER_SAMPLE = [
    '"placeholder": "Legal Reviewer",',
    '"email_notification": true,',
    '"role": "SIGNER",',
    '"signing_order": 1',
]

USER_RESULT = [
    '"order": 1,',
    '"user_email": "giatk@mobile-id.vn",',
    '"user_name": "TatKhanhGia",',
    '"role": "SIGNER",',
    '"signing_order": 1,',
    '"guest_user": false,',
    '"email_language_code": "en",',
    '"email_notification": true,',
    '"mobile_number": "0912345678"',
]

PLACEHOLDER_RESULT = [
    '"order": 1,',
    '"placeholder": "Legal Reviewer",',
    '"role": "SIGNER",',
    '"signing_order": 1,',
    '"guest_user": false,',
    '"email_language_code": "en",',
    '"email_notification": true',
]


def wrap_array(inner):
    """Root-level array body. The house style indents two spaces per level."""
    return ["[", "  {"] + ["    " + l for l in inner] + ["  }", "]"]


def wrap_object(inner):
    return ["{"] + ["  " + l for l in inner] + ["}"]


SECTIONS = [
    dict(
        title="Get Template Workflow Users",
        donor="Get Workflow Users",
        method="GET",
        path="{baseURL}/v1/templates/{id}/workflow/users",
        desc=("Business applications can use this service API to get the list of "
              "participants configured in the workflow of a personal/enterprise "
              "template. This API can also be used by the business application to "
              "perform the action on behalf of the enterprise user."),
        req_body=None,
        resp_body=wrap_array(USER_RESULT),
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC)],
        req_rows=None,
        resp_rows=ERROR_ROWS + PARTICIPANT_ROWS,
        status="200",
    ),
    dict(
        title="Add Template Workflow Users",
        donor="Add User to Workflow",
        method="POST",
        path="{baseURL}/v1/templates/{id}/workflow/users",
        desc=("Business applications can use this service API to add one or more "
              "users to the workflow of a personal/enterprise template. The request "
              "body is a JSON array, so several users can be added in a single call. "
              "This API can also be used by the business application to perform the "
              "action on behalf of the enterprise user."),
        req_body=wrap_array(USER_SAMPLE),
        resp_body=wrap_array(USER_RESULT),
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC)],
        req_rows=USER_REQUEST_ROWS,
        resp_rows=ERROR_ROWS + PARTICIPANT_ROWS,
        status="201",
    ),
    dict(
        title="Add Template Workflow Placeholder",
        donor="Add Placeholder to Workflow",
        method="POST",
        path="{baseURL}/v1/templates/{id}/workflow/placeholder",
        desc=("Business applications can use this service API to add one or more "
              "placeholders to the workflow of a personal/enterprise template. A "
              "placeholder reserves a position in the workflow that is filled with a "
              "real user when a workflow is created from the template. The request "
              "body is a JSON array, so several placeholders can be added in a "
              "single call."),
        req_body=wrap_array(PLACEHOLDER_SAMPLE),
        resp_body=wrap_array(PLACEHOLDER_RESULT),
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC)],
        req_rows=PLACEHOLDER_REQUEST_ROWS,
        resp_rows=ERROR_ROWS + PARTICIPANT_ROWS,
        status="201",
    ),
    dict(
        title="Update Template Workflow User",
        donor="Update Workflow User",
        method="PUT",
        path="{baseURL}/v1/templates/{id}/workflow/{order}/user",
        desc=("Business applications can use this service API to update a user "
              "already present in the workflow of a personal/enterprise template. "
              "This API can also be used by the business application to perform the "
              "action on behalf of the enterprise user."),
        req_body=wrap_object(USER_SAMPLE),
        resp_body=wrap_object(USER_RESULT),
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC),
                   ("order", "Integer", "M", ORDER_DESC)],
        req_rows=USER_REQUEST_ROWS,
        resp_rows=ERROR_ROWS + PARTICIPANT_ROWS,
        status="200",
    ),
    dict(
        title="Update Template Workflow Placeholder",
        donor="Update Placeholder in Workflow",
        method="PUT",
        path="{baseURL}/v1/templates/{id}/workflow/{order}/placeholder",
        desc=("Business applications can use this service API to update a "
              "placeholder already present in the workflow of a personal/enterprise "
              "template. This API can also be used by the business application to "
              "perform the action on behalf of the enterprise user."),
        req_body=wrap_object(PLACEHOLDER_SAMPLE),
        resp_body=wrap_object(PLACEHOLDER_RESULT),
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC),
                   ("order", "Integer", "M", ORDER_DESC)],
        req_rows=PLACEHOLDER_REQUEST_ROWS,
        resp_rows=ERROR_ROWS + PARTICIPANT_ROWS,
        status="200",
    ),
    dict(
        title="Reorder Template Workflow Participant",
        donor="Update Placeholder in Workflow",
        method="PUT",
        path="{baseURL}/v1/templates/{id}/workflow/{order}/reorder",
        desc=("Business applications can use this service API to move a participant "
              "to a different position in the workflow of a personal/enterprise "
              "template. The participant currently at the target position and every "
              "participant in between are shifted accordingly."),
        req_body=["{", '  "order": 2', "}"],
        resp_body=["{}"],
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC),
                   ("order", "Integer", "M", ORDER_DESC)],
        req_rows=[("order", "Integer", "M",
                   "The new position of the participant in the template workflow.")],
        resp_rows=ERROR_ROWS,
        status="200",
    ),
    dict(
        title="Delete Template Workflow Participant",
        donor="Delete Workflow User",
        method="DELETE",
        path="{baseURL}/v1/templates/{id}/workflow/{order}",
        desc=("Business applications can use this service API to remove a "
              "participant - either a user or a placeholder - from the workflow of a "
              "personal/enterprise template. The remaining participants are "
              "renumbered."),
        req_body=None,
        resp_body=["{}"],
        path_rows=[("id", "Integer", "M", TEMPLATE_ID_DESC),
                   ("order", "Integer", "M", ORDER_DESC)],
        req_rows=None,
        resp_rows=ERROR_ROWS,
        status="200",
    ),
]


def set_frame_url(para, url):
    """Rewrite a frame's first line without disturbing the method badge.

    The line is always laid out as badge run, a space run, then the URL split
    across one run per path segment. Keeping the first two runs preserves the
    shading and white text template 2.8.1 requires on the badge; the URL is
    collapsed into the first segment run so the remaining ones can be emptied.
    """
    runs = para.findall(st.qn("r"))
    text_runs = [r for r in runs if r.find(st.qn("t")) is not None]
    if len(text_runs) < 3:
        raise ValueError("unexpected frame first line: %r" % st.text_of(para))
    target = text_runs[2]
    t = target.find(st.qn("t"))
    t.text = url
    t.set(st.XML_SPACE, "preserve")
    for r in text_runs[3:]:
        rt = r.find(st.qn("t"))
        if rt is not None and rt.text:
            rt.text = ""


def set_frame_body(txbx, paras, lines):
    """Replace a frame's payload, keeping the header lines above it.

    section_tools.replace_frame_body only recognises an object body. These
    endpoints also carry root-level arrays, so the first payload line may be
    '[' as well as '{' or '{}'.
    """
    texts = [st.text_of(p).strip() for p in paras]
    body_start = None
    for i, line in enumerate(texts):
        if i == 0:
            continue
        if line in ("{", "[", "{}"):
            body_start = i
            break
    if body_start is None:
        raise LookupError("frame has no payload: %r" % texts[:2])
    template = paras[body_start]
    anchor = paras[body_start - 1]
    for p in paras[body_start:]:
        txbx.remove(p)
    for line in lines:
        new_p = deepcopy(template)
        st.set_para_text(new_p, line)
        anchor.addnext(new_p)
        anchor = new_p


def fill_table(tbl, rows):
    """Make an attribute table hold exactly `rows`, reusing its existing rows."""
    existing = st.data_rows(tbl)
    for i, (name, typ, presence, desc) in enumerate(rows):
        if i < len(existing):
            tcs = existing[i].findall(st.qn("tc"))
            st.set_cell_text(tcs[1], name)
            st.set_cell_text(tcs[2], typ)
            st.set_cell_text(tcs[3], presence)
            st.set_cell_text(tcs[4], desc)
        else:
            st.append_row(tbl, name, typ, presence, desc)
    for extra in st.data_rows(tbl)[len(rows):]:
        extra.getparent().remove(extra)
    st.renumber(tbl)


def build(root, body, spec):
    kids = list(body)
    clone, info = st.clone_section(root, kids, spec["donor"], level=3)

    # heading + description
    st.set_para_text(clone[0], spec["title"])
    for el in clone[1:]:
        if el.tag == st.qn("p") and st.text_of(el).strip():
            st.set_para_text(el, spec["desc"])
            break

    # request frame (both the mc:Choice and mc:Fallback copies)
    n_req = 0
    for txbx, paras in st.each_frame(clone, spec["method"] + " "):
        set_frame_url(paras[0], spec["path"])
        set_frame_body(txbx, paras, spec["req_body"] or ["{}"])
        n_req += 1
    if not n_req:
        raise LookupError("no request frame starting with %r" % spec["method"])

    # success response frame - the donor's code may differ from ours
    n_resp = 0
    for prefix in ("Http Code: 200", "Http Code: 201"):
        for txbx, paras in st.each_frame(clone, prefix):
            st.set_para_text(paras[0], "Http Code: " + spec["status"])
            set_frame_body(txbx, paras, spec["resp_body"])
            n_resp += 1
    if not n_resp:
        raise LookupError("no success frame in donor %r" % spec["donor"])

    # tables
    fill_table(st.find_table(clone, "Path Attributes"), spec["path_rows"])
    fill_table(st.find_table(clone, "Response Attributes"), spec["resp_rows"])
    try:
        req_tbl = st.find_table(clone, "Request Attributes")
    except LookupError:
        req_tbl = None
    if spec["req_rows"]:
        if req_tbl is None:
            raise LookupError("donor %r has no Request Attributes table to fill"
                              % spec["donor"])
        fill_table(req_tbl, spec["req_rows"])
    elif req_tbl is not None:
        as_list = list(clone)
        st.drop_table(as_list, "Request Attributes")
        clone = as_list
    return clone, info


def main():
    tree, token = docx_guard.load(DOC)
    root = tree.getroot()
    body = root.find(st.qn("body"))

    anchor = "Delete Template"          # last section of the Template chapter
    for spec in SECTIONS:
        clone, info = build(root, body, spec)
        st.insert_after_section(body, list(body), anchor, clone, level=3)
        anchor = spec["title"]
        print("+ %-42s from %-32s (%d elements, %d docPr, %d vml)"
              % (spec["title"], spec["donor"], info["elements"],
                 info["docpr_reassigned"], info["vml_reassigned"]))

    docx_guard.save(tree, DOC, token)
    print("saved %s" % DOC)

    # Deliver. The validation scripts read the packed DOCX, not the unpacked
    # working copy, so nothing is actually verified until this has run.
    if not os.path.exists(BACKUP):
        shutil.copy2(OUT_DOCX, BACKUP)
        print("backed up -> %s" % os.path.basename(BACKUP))
    st.pack(UNPACKED, OUT_DOCX)
    print("packed    -> %s" % os.path.basename(OUT_DOCX))


if __name__ == "__main__":
    main()
