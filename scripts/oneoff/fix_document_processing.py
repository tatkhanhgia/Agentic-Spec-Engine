#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Align the Document Processing section with the backend contract (2026-07-27).

User-approved changes (groups 1-5 of the classification report):
  Approve Document : drop x-password/x-otp header rows, drop documentId path row,
                     reason M->O, drop x-password/x-otp lines from the sample frame.
  Decline Document : drop documentId path row, note 256-char limit on reason.
  Sign Document    : drop x-otp header row + sample line, drop request rows
                     appearance_design / signing_capacity / signing_server /
                     authentication / connector_name (backend DTO has none of
                     them), fix signing_location type Boolean->String, remove the
                     same keys from the sample request JSON.
  Sign Eseal       : drop x-otp header row + sample line, add optional request
                     rows hand_signature_image / signing_location /
                     contact_information (cloned from Sign Document).
  Get Document Hash: drop request rows skip_verification / signing_order /
                     time_zone, add appearance_design (O), certificate_chain
                     M->O, response rows hash/state/hashing_algo -> hash_value,
                     align both sample frames.
  Fill Form Fields : rename request keys text->textbox, radio->radiobox,
                     combo->combobox, checkboxV2->checkboxv2,
                     radioboxV2->radioboxv2 (tables + sample), add rows
                     datetime / qr / vneid_field / custom_field, add
                     x-handle-transaction header row.
  Fill Initials    : initial_pages -> initial_field (List<String>), field_name
                     M->O, image M->RC, add x-handle-transaction header row.
  Submit Document  : drop x-password/x-otp header rows + sample lines, add a
                     Request Attributes table with reason (O) cloned from
                     Decline, sample body {} -> { "reason": "String" }.

All table edits clone existing structures; nothing is assembled by hand.
Writes through docx_guard (load/save) so a concurrent edit aborts the write.
"""
import copy
import os
import sys

SCRIPTS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SCRIPTS)
from docx_guard import load, save, DOC_XML  # noqa: E402
from lxml import etree  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(tag):
    return "{%s}%s" % (W, tag)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(qn("t")))


def style_of(p):
    ps = p.find(qn("pPr") + "/" + qn("pStyle"))
    return ps.get(qn("val")) if ps is not None else ""


def cell_text(tc):
    return " ".join(para_text(p) for p in tc.iter(qn("p"))).strip()


def strip_bookmarks(el):
    for tag in ("bookmarkStart", "bookmarkEnd"):
        for b in el.findall(".//" + qn(tag)):
            b.getparent().remove(b)


def set_cell_text(tc, text):
    """Replace the visible text of a cell, respecting a hyperlink wrapper."""
    paras = tc.findall(qn("p"))
    first_t = None
    for i, p in enumerate(paras):
        container = p.find(qn("hyperlink"))
        scope = container if container is not None else p
        for r in scope.findall(qn("r")):
            t = r.find(qn("t"))
            if t is None:
                continue
            if first_t is None:
                first_t = t
                t.text = text
                t.set(XML_SPACE, "preserve")
            else:
                t.text = ""
    if first_t is None:
        raise RuntimeError("cell has no run to write into: %r" % cell_text(tc))


def data_rows(tbl):
    out = []
    for tr in tbl.findall(qn("tr")):
        cells = tr.findall(qn("tc"))
        if len(cells) >= 5 and cell_text(cells[0]) != "No":
            out.append((tr, cells))
    return out


def find_row(tbl, name):
    for tr, cells in data_rows(tbl):
        if cell_text(cells[1]) == name:
            return tr, cells
    raise RuntimeError("row %r not found" % name)


def delete_row(tbl, name):
    tr, _ = find_row(tbl, name)
    tr.getparent().remove(tr)


def renumber(tbl):
    for i, (tr, cells) in enumerate(data_rows(tbl), 1):
        set_cell_text(cells[0], str(i))


def clone_row(tbl, name):
    tr, _ = find_row(tbl, name)
    dup = copy.deepcopy(tr)
    strip_bookmarks(dup)
    return dup


def edit_row(tr, name=None, typ=None, presence=None, desc=None):
    cells = tr.findall(qn("tc"))
    if name is not None:
        set_cell_text(cells[1], name)
    if typ is not None:
        set_cell_text(cells[2], typ)
    if presence is not None:
        set_cell_text(cells[3], presence)
    if desc is not None:
        set_cell_text(cells[4], desc)


# ---------- frame (textbox) helpers -------------------------------------

def txbx_of(p):
    return p.findall(".//" + qn("txbxContent"))


def frame_remove_lines(p, pred):
    n = 0
    for tx in txbx_of(p):
        for line in list(tx.findall(qn("p"))):
            if pred(para_text(line).strip()):
                tx.remove(line)
                n += 1
    return n


def line_set_text(line, new_text):
    first = None
    for t in line.iter(qn("t")):
        if first is None:
            first = t
            t.text = new_text
            t.set(XML_SPACE, "preserve")
        else:
            t.text = ""


def frame_replace(p, old, new, count_holder):
    """Replace substring old->new in every frame line that contains it."""
    for tx in txbx_of(p):
        for line in tx.findall(qn("p")):
            full = para_text(line)
            if old not in full:
                continue
            done = False
            for t in line.iter(qn("t")):
                if t.text and old in t.text:
                    t.text = t.text.replace(old, new)
                    done = True
            if not done:  # split across runs: rebuild the line
                line_set_text(line, full.replace(old, new))
            count_holder[0] += 1


def frame_set_line(p, match_strip, new_text, only_containing=None):
    n = 0
    for tx in txbx_of(p):
        if only_containing is not None:
            joined = "".join(para_text(l) for l in tx.findall(qn("p")))
            if only_containing not in joined:
                continue
        for line in tx.findall(qn("p")):
            if para_text(line).strip() == match_strip:
                line_set_text(line, new_text)
                n += 1
    return n


# ---------- section walker ----------------------------------------------

ENDPOINTS = {
    "Approve Document", "Decline Document", "Sign Document", "Sign Eseal",
    "Get Document Hash", "Embed Signature", "Fill Form Fields",
    "Fill Initials", "Submit Document",
}


def collect_sections(body):
    """Map endpoint heading -> ordered list of following block elements."""
    sections = {}
    in_dp = False
    cur = None
    for el in body:
        tag = etree.QName(el).localname
        if tag == "p":
            st = style_of(el)
            txt = para_text(el).strip()
            if st == "Heading2":
                in_dp = (txt == "Document Processing")
                cur = None
                continue
            if not in_dp:
                continue
            if st in ("Heading3", "Heading4"):
                cur = txt if txt in ENDPOINTS else None
                continue
        if in_dp and cur:
            sections.setdefault(cur, []).append(el)
    missing = ENDPOINTS - set(sections)
    if missing:
        raise RuntimeError("endpoint sections not found: %s" % missing)
    return sections


def tables_of(section):
    out = {}
    for el in section:
        if etree.QName(el).localname != "tbl":
            continue
        rows = el.findall(qn("tr"))
        if not rows:
            continue
        label = cell_text(rows[0].findall(qn("tc"))[0])
        out[label] = el
    return out


def sample_paras(section):
    """Return (request_para, response_para) — paragraphs holding textboxes."""
    req = res = None
    mode = None
    for el in section:
        if etree.QName(el).localname != "p":
            continue
        txt = para_text(el).strip()
        if txt == "Sample Request":
            mode = "req"
            continue
        if txt == "Sample Response":
            mode = "res"
            continue
        if txbx_of(el):
            if mode == "req" and req is None:
                req = el
            elif mode == "res" and res is None:
                res = el
    if req is None or res is None:
        raise RuntimeError("sample frames not found")
    return req, res


def main():
    tree, token = load(DOC_XML)
    body = tree.getroot().find(qn("body"))
    sec = collect_sections(body)

    is_otp = lambda s: s.startswith("x-otp")
    is_pwd = lambda s: s.startswith("x-password")

    # ---------------- Sign Document (first: its rows are clone sources) ---
    s = sec["Sign Document"]
    t = tables_of(s)
    treq, thdr = t["Request Attributes"], t["Header Request Attributes"]
    edit_row(find_row(treq, "signing_location")[0], typ="String")
    eseal_clones = [clone_row(treq, n) for n in
                    ("hand_signature_image", "signing_location",
                     "contact_information")]
    appearance_clone = clone_row(treq, "appearance_design")
    for n in ("appearance_design", "signing_capacity", "signing_server",
              "authentication", "connector_name"):
        delete_row(treq, n)
    renumber(treq)
    delete_row(thdr, "x-otp")
    renumber(thdr)
    req_p, _ = sample_paras(s)
    drop = {'"appearance_design": "String",', '"signing_capacity": "String",',
            '"signing_server": "String",', '"connector_name": "String",',
            '"authentication": {', '"mechanism": "String",',
            '"password": "String",', '"token": "String",',
            '"user_identifier": "String"', "},"}
    n = frame_remove_lines(req_p, lambda x: x in drop or is_otp(x))
    print("Sign Document: sample lines removed:", n)

    # ---------------- Sign Eseal -----------------------------------------
    s = sec["Sign Eseal"]
    t = tables_of(s)
    treq, thdr = t["Request Attributes"], t["Header Request Attributes"]
    anchor = find_row(treq, "field_name")[0]
    for row in reversed(eseal_clones):
        anchor.addnext(copy.deepcopy(row))
    renumber(treq)
    delete_row(thdr, "x-otp")
    renumber(thdr)
    req_p, _ = sample_paras(s)
    print("Sign Eseal: sample lines removed:",
          frame_remove_lines(req_p, is_otp))

    # ---------------- Approve Document ------------------------------------
    s = sec["Approve Document"]
    t = tables_of(s)
    for n in ("x-password", "x-otp"):
        delete_row(t["Header Request Attributes"], n)
    renumber(t["Header Request Attributes"])
    delete_row(t["Path Attributes"], "documentId")
    renumber(t["Path Attributes"])
    edit_row(find_row(t["Request Attributes"], "reason")[0], presence="O")
    req_p, _ = sample_paras(s)
    print("Approve: sample lines removed:",
          frame_remove_lines(req_p, lambda x: is_otp(x) or is_pwd(x)))

    # ---------------- Decline Document ------------------------------------
    s = sec["Decline Document"]
    t = tables_of(s)
    delete_row(t["Path Attributes"], "documentId")
    renumber(t["Path Attributes"])
    edit_row(find_row(t["Request Attributes"], "reason")[0],
             desc="The reason for decline a package. "
                  "Maximum length: 256 characters.")
    decline_req_tbl = t["Request Attributes"]  # clone source for Submit
    decline_req_p, _ = sample_paras(s)

    # ---------------- Get Document Hash ------------------------------------
    s = sec["Get Document Hash"]
    t = tables_of(s)
    treq, tres = t["Request Attributes"], t["Response Attributes"]
    for n in ("skip_verification", "signing_order", "time_zone"):
        delete_row(treq, n)
    last = data_rows(treq)[-1][0]
    last.addnext(appearance_clone)
    edit_row(find_row(treq, "certificate_chain")[0], presence="O")
    renumber(treq)
    edit_row(find_row(tres, "hash")[0], name="hash_value")
    for n in ("state", "hashing_algo"):
        delete_row(tres, n)
    renumber(tres)
    req_p, res_p = sample_paras(s)
    n = frame_remove_lines(
        req_p, lambda x: x in ('"signing_order": 1,', '"time_zone": "String"'))
    holder = [0]
    frame_replace(req_p, "],", "]", holder)
    print("Hash request: lines removed %d, '],' fixed %d" % (n, holder[0]))
    n = frame_set_line(res_p, '"hash": "String",', '  "hash_value": "String"',
                       only_containing='"hash"')
    n += frame_remove_lines(
        res_p, lambda x: x in ('"state": "String",',
                               '"hashing_algo": "String",',
                               '"signing_algo": "String"'))
    print("Hash response: frame edits:", n)

    # ---------------- Fill Form Fields -------------------------------------
    s = sec["Fill Form Fields"]
    t = tables_of(s)
    treq, thdr = t["Request Attributes"], t["Header Request Attributes"]
    renames = [("text", "textbox", "Data of the textbox fields"),
               ("radio", "radiobox", "Data of the radiobox fields"),
               ("combo", "combobox", "Data of the combobox fields"),
               ("checkboxV2", "checkboxv2", "Data of the checkboxv2 fields"),
               ("radioboxV2", "radioboxv2", "Data of the radioboxv2 fields")]
    for old, new, desc in renames:
        edit_row(find_row(treq, old)[0], name=new, desc=desc)
    date_row = find_row(treq, "date")[0]
    dt = clone_row(treq, "date")
    edit_row(dt, name="datetime", desc="Data of the datetime fields")
    date_row.addnext(dt)
    tail = data_rows(treq)[-1][0]
    for nm, ds in [("qr", "Data of the QR fields"),
                   ("vneid_field", "Data of the VNeID fields"),
                   ("custom_field", "Data of the custom fields")]:
        row = clone_row(treq, "checkbox")
        edit_row(row, name=nm, desc=ds)
        tail.addnext(row)
        tail = row
    renumber(treq)
    handle_desc = ("Idempotency token of the request. If the same value is "
                   "sent again within 10 minutes, the server returns success "
                   "immediately without reprocessing. A new value is echoed "
                   "back in the response header.")
    hrow = clone_row(thdr, "x-idempotency-skip-error")
    edit_row(hrow, name="x-handle-transaction", typ="String", presence="O",
             desc=handle_desc)
    find_row(thdr, "x-idempotency-skip-error")[0].addnext(hrow)
    renumber(thdr)
    req_p, _ = sample_paras(s)
    holder = [0]
    for old, new in [('"text"', '"textbox"'), ('"radio"', '"radiobox"'),
                     ('"combo"', '"combobox"'),
                     ('"checkboxV2"', '"checkboxv2"'),
                     ('"radioboxV2"', '"radioboxv2"')]:
        frame_replace(req_p, old, new, holder)
    print("Fill Form Fields: sample key renames:", holder[0])

    # ---------------- Fill Initials ----------------------------------------
    s = sec["Fill Initials"]
    t = tables_of(s)
    treq, thdr = t["Request Attributes"], t["Header Request Attributes"]
    edit_row(find_row(treq, "field_name")[0], presence="O",
             desc="Name of the initial field. At least one of field_name or "
                  "initial_field must be provided.")
    edit_row(find_row(treq, "image")[0],
             presence="RC (Required when the initial field has no stored "
                      "image)")
    edit_row(find_row(treq, "initial_pages")[0], name="initial_field",
             typ="List<String>",
             desc="The list of initial field names to be filled.")
    hrow2 = copy.deepcopy(hrow)
    strip_bookmarks(hrow2)
    find_row(thdr, "Content-Type")[0].addprevious(hrow2)
    renumber(thdr)

    # ---------------- Submit Document --------------------------------------
    s = sec["Submit Document"]
    t = tables_of(s)
    for n in ("x-password", "x-otp"):
        delete_row(t["Header Request Attributes"], n)
    renumber(t["Header Request Attributes"])
    new_tbl = copy.deepcopy(decline_req_tbl)
    strip_bookmarks(new_tbl)
    edit_row(find_row(new_tbl, "reason")[0], presence="O",
             desc="The reason for submitting a package")
    tres = t["Response Attributes"]
    blank = None  # the empty paragraph right before the Response table
    prev = tres.getprevious()
    if prev is not None and etree.QName(prev).localname == "p" \
            and not para_text(prev).strip():
        blank = prev
    if blank is None:
        raise RuntimeError("Submit: blank paragraph before Response table "
                           "not found")
    new_blank = copy.deepcopy(blank)
    blank.addprevious(new_tbl)
    new_tbl.addprevious(new_blank)  # blank stays between the two tables
    req_p, _ = sample_paras(s)
    n = frame_remove_lines(req_p, lambda x: is_otp(x) or is_pwd(x))
    # body {} -> cloned { "reason": "String" } lines from the Decline frame
    donors = txbx_of(decline_req_p)
    body_lines = []
    for donor in donors:
        lines = [l for l in donor.findall(qn("p"))
                 if para_text(l).strip() in ("{", '"reason": "String"', "}")]
        if len(lines) == 3:
            body_lines = lines
            break
    if len(body_lines) != 3:
        raise RuntimeError("Submit: donor body lines not found in Decline "
                           "frame")
    replaced = 0
    for tx in txbx_of(req_p):
        for line in list(tx.findall(qn("p"))):
            if para_text(line).strip() == "{}":
                for repl in reversed(body_lines):
                    dup = copy.deepcopy(repl)
                    strip_bookmarks(dup)
                    line.addnext(dup)
                tx.remove(line)
                replaced += 1
    print("Submit: header lines removed %d, body frames replaced %d"
          % (n, replaced))

    # ---------------- hyperlink sanity check (edited section only) ---------
    anchors = {b.get(qn("name")) for b in
               tree.getroot().iter(qn("bookmarkStart"))}
    bad = []
    for name, blocks in sec.items():
        for el in blocks:
            if etree.QName(el).localname != "tbl":
                continue
            for h in el.findall(".//" + qn("hyperlink")):
                label = "".join(tt.text or "" for tt in h.iter(qn("t")))
                anchor = h.get(qn("anchor"))
                if anchor and (anchor not in anchors or not label.strip()):
                    bad.append((name, anchor, label))
    if bad:
        raise RuntimeError("hyperlink check failed: %s" % bad)
    print("hyperlink check: OK")

    save(tree, DOC_XML, token)


if __name__ == "__main__":
    main()
