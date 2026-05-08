#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Fix GoPaperless API Spec docx format inconsistencies.
Reads style configuration from GoPaperless_DOCX_Format_Template.md,
apply to unpacked_docx/ directory, then packs to a new .docx file.
"""
import os
import sys
import shutil
import zipfile
import re
from copy import deepcopy
from lxml import etree

# Namespaces used in WordprocessingML
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def qn(tag):
    """Return qualified name for w:* tag."""
    return W + tag


# ---------------------------------------------------------------------------
# Template parser
# ---------------------------------------------------------------------------

def parse_template(md_path):
    """Parse the style table from GoPaperless_DOCX_Format_Template.md."""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    header = None
    rows = []
    in_table = False

    for line in lines:
        if '| Style |' in line and 'Font' in line:
            in_table = True
        if in_table:
            if line.strip() == '' or line.startswith('##'):
                break
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c]
            if not cells:
                continue
            if set(cells[0]) <= set('-'):
                continue
            if header is None:
                header = cells
            else:
                rows.append(cells)

    def to_style_id(name):
        name = name.replace('**', '')
        mappings = {
            'List Paragraph': 'ListParagraph',
            'Heading 1': 'Heading1',
            'Heading 2': 'Heading2',
            'Heading 3': 'Heading3',
            'Heading 4': 'Heading4',
        }
        return mappings.get(name, name)

    def parse_size(size_str):
        m = re.search(r'(\d+(?:\.\d+)?)', size_str)
        return int(float(m.group(1)) * 2) if m else None

    def parse_bold(bold_str):
        s = bold_str.strip().lower()
        return s in ("yes", "true", "bold", "co", "có")

    def parse_color(color_str):
        s = color_str.strip().lower()
        if s in ('auto', 'none', ''):
            return None
        if re.match(r'^[0-9a-f]{6}$', s):
            return s.upper()
        return None

    def parse_alignment(align_str):
        s = align_str.lower()
        if 'left' in s:
            return 'left'
        if 'center' in s:
            return 'center'
        if 'right' in s:
            return 'right'
        return None

    def parse_pt(pt_str):
        m = re.search(r'(\d+(?:\.\d+)?)', pt_str)
        return float(m.group(1)) if m else None

    config = {}
    for row in rows:
        style_name = row[0]
        style_id = to_style_id(style_name)
        cfg = {
            'font': row[1] if len(row) > 1 else None,
            'sz': parse_size(row[2]) if len(row) > 2 else None,
            'bold': parse_bold(row[3]) if len(row) > 3 else None,
            'color': parse_color(row[4]) if len(row) > 4 else None,
            'alignment': parse_alignment(row[5]) if len(row) > 5 else None,
            'space_before': parse_pt(row[6]) if len(row) > 6 else None,
            'space_after': parse_pt(row[7]) if len(row) > 7 else None,
            'left_indent': parse_pt(row[8]) if len(row) > 8 else None,
            'hanging_indent': parse_pt(row[9]) if len(row) > 9 else None,
        }
        config[style_id] = cfg

    return config


def parse_heading_hierarchy_rules(md_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    rules = {}
    for line in content.splitlines():
        if "Authentication endpoint hierarchy" not in line:
            continue
        endpoints = re.findall(r'`([^`]+)`', line)
        for endpoint in endpoints:
            if endpoint != "Authenticate - Log in":
                rules[endpoint] = "Heading3"
    return rules


def read_markdown_table(md_path, header_marker):
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.splitlines()
    header = None
    rows = []
    in_table = False

    for line in lines:
        if header_marker in line:
            in_table = True
        if in_table:
            if line.strip() == '' or line.startswith('##'):
                break
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c]
            if not cells:
                continue
            if set(cells[0]) <= set('-'):
                continue
            if header is None:
                header = cells
            else:
                rows.append(cells)

    return header, rows


def parse_table_border_rules(md_path):
    _, rows = read_markdown_table(md_path, '| Section Heading |')

    rules = []
    for row in rows:
        if len(row) < 6:
            continue
        rules.append({
            'section_heading': row[0].replace('**', ''),
            'table_borders': [x.strip() for x in row[1].split(',') if x.strip()],
            'cell_borders': [x.strip() for x in row[2].split(',') if x.strip()],
            'border_size': row[3],
            'border_color': row[4].replace('#', '').upper(),
            'apply_cell_borders': row[5].strip().lower() in ('yes', 'true'),
        })

    return rules


def parse_table_property_rules(md_path):
    _, rows = read_markdown_table(md_path, '| Property |')
    return {row[0].replace('**', '').replace('`', '').strip(): row[1].replace('`', '').strip() for row in rows if len(row) >= 2}


def parse_int_property(table_property_rules, key):
    m = re.search(r'(\d+)', table_property_rules.get(key, ''))
    return int(m.group(1)) if m else None


def parse_api_body_rules(md_path):
    header, rows = read_markdown_table(md_path, '| Element | Text Match | Paragraph Type |')
    if not header:
        return {}

    def clean(value):
        return value.replace('**', '').replace('`', '').strip()

    def parse_bool(value):
        return clean(value).lower() in ('yes', 'true', 'bold', 'co', 'có')

    def parse_pt_value(value):
        m = re.search(r'(\d+(?:\.\d+)?)', value)
        return float(m.group(1)) if m else None

    rules = {}
    for row in rows:
        if len(row) < len(header):
            continue
        item = dict(zip(header, row))
        name = clean(item['Element'])
        rules[name] = {
            'text_match': clean(item.get('Text Match', '')),
            'paragraph_type': clean(item.get('Paragraph Type', '')),
            'font': clean(item.get('Font', '')),
            'size': parse_pt_value(item.get('Size', '')),
            'bold': parse_bool(item.get('Bold', '')),
            'space_before': parse_pt_value(item.get('Space Before', '')),
            'space_after': parse_pt_value(item.get('Space After', '')),
            'left_indent': parse_pt_value(item.get('Left Indent', '')),
            'blank_count': int(parse_pt_value(item.get('Blank Count', '')) or 0),
            'order_columns': clean(item.get('Order / Columns', '')),
        }

    return rules


def parse_frame_method_rules(md_path):
    """Parse HTTP method badge highlight/text colors from template."""
    header, rows = read_markdown_table(md_path, '| Method |')
    rules = {}
    for row in rows:
        if len(row) < 3:
            continue
        method = row[0].replace('**', '').strip()
        highlight = parse_hex_color(row[1])
        text_color = parse_hex_color(row[2])
        if method and highlight and text_color:
            rules[method.upper()] = {
                'highlight': highlight,
                'text_color': text_color,
            }
    return rules


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def get_or_create(parent, tag):
    """Get first child `tag` or create it."""
    child = parent.find(qn(tag))
    if child is None:
        child = etree.SubElement(parent, qn(tag))
    return child


def set_fonts(rPr, font_name="Verdana"):
    """Ensure rFonts inside rPr points to a single font."""
    rf = rPr.find(qn("rFonts"))
    if rf is None:
        rf = etree.SubElement(rPr, qn("rFonts"))
    for attr in list(rf.attrib.keys()):
        del rf.attrib[attr]
    rf.set(qn("ascii"), font_name)
    rf.set(qn("hAnsi"), font_name)
    rf.set(qn("cs"), font_name)
    rf.set(qn("eastAsia"), font_name)


def set_size(rPr, half_pts):
    """Set sz and szCs to half-points value."""
    for tag in ("sz", "szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set(qn("val"), str(half_pts))


def set_bold(rPr, bold=True):
    """Set or remove bold."""
    b = rPr.find(qn("b"))
    bCs = rPr.find(qn("bCs"))
    if bold:
        if b is None:
            etree.SubElement(rPr, qn("b"))
        if bCs is None:
            etree.SubElement(rPr, qn("bCs"))
    else:
        if b is not None:
            rPr.remove(b)
        if bCs is not None:
            rPr.remove(bCs)


def set_color(rPr, color_val=None):
    """Set or remove color. color_val like '0050A8' or None to remove."""
    c = rPr.find(qn("color"))
    if color_val:
        if c is None:
            c = etree.SubElement(rPr, qn("color"))
        for attr in list(c.attrib.keys()):
            del c.attrib[attr]
        c.set(qn("val"), color_val)
    else:
        if c is not None:
            rPr.remove(c)


def set_highlight(rPr, color_val=None):
    """Set or remove text highlight (w:highlight). Only supports named colors."""
    h = rPr.find(qn("highlight"))
    if color_val:
        if h is None:
            h = etree.SubElement(rPr, qn("highlight"))
        h.set(qn("val"), color_val)
    else:
        if h is not None:
            rPr.remove(h)


def set_shd(rPr, fill=None):
    """Set or remove shading (w:shd) with custom fill color."""
    s = rPr.find(qn("shd"))
    if fill:
        if s is None:
            s = etree.SubElement(rPr, qn("shd"))
        s.set(qn("val"), "clear")
        s.set(qn("color"), "auto")
        s.set(qn("fill"), fill)
    else:
        if s is not None:
            rPr.remove(s)


def set_alignment(pPr, alignment):
    """Set paragraph alignment. alignment: left, center, right."""
    align_map = {
        'left': 'left',
        'center': 'center',
        'right': 'right',
        'justify': 'both',
    }
    if alignment not in align_map:
        return
    jc = get_or_create(pPr, "jc")
    jc.set(qn("val"), align_map[alignment])


def set_cell_valign(tcPr, valign):
    """Set vertical alignment for a table cell. valign: top, center, bottom."""
    if valign not in {"top", "center", "bottom"}:
        return
    v = get_or_create(tcPr, "vAlign")
    v.set(qn("val"), valign)


def reset_table_cell_paragraph_spacing(pPr):
    spacing = get_or_create(pPr, "spacing")
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "240")
    spacing.set(qn("lineRule"), "auto")


def reset_table_cell_paragraph_indent(pPr):
    ind = get_or_create(pPr, "ind")
    ind.set(qn("left"), "0")
    ind.set(qn("right"), "0")
    ind.set(qn("firstLine"), "0")
    for attr in (qn("hanging"), qn("start"), qn("end"), qn("firstLineChars"), qn("hangingChars")):
        if attr in ind.attrib:
            del ind.attrib[attr]


def set_cell_margins(tcPr, margin="55"):
    tc_mar = get_or_create(tcPr, "tcMar")
    for side in ("top", "left", "bottom", "right"):
        node = get_or_create(tc_mar, side)
        node.set(qn("w"), margin)
        node.set(qn("type"), "dxa")


def make_run(text, font="Verdana", size=22, bold=False, color=None):
    run = etree.Element(qn("r"))
    rPr = etree.SubElement(run, qn("rPr"))
    set_fonts(rPr, font)
    set_size(rPr, size)
    set_bold(rPr, bold)
    set_color(rPr, color)
    t = etree.SubElement(run, qn("t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def make_paragraph(text, style="Normal", config=None, bold=False):
    para = etree.Element(qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    pStyle = etree.SubElement(pPr, qn("pStyle"))
    pStyle.set(qn("val"), style)
    if config and style in config:
        apply_paragraph_style_layout(pPr, config[style])
    para.append(make_run(text, bold=bold))
    return para


def make_table_cell(text, header=False):
    tc = etree.Element(qn("tc"))
    tcPr = etree.SubElement(tc, qn("tcPr"))
    tcW = etree.SubElement(tcPr, qn("tcW"))
    tcW.set(qn("w"), "1666")
    tcW.set(qn("type"), "pct")
    set_cell_valign(tcPr, "center")
    set_cell_margins(tcPr)
    para = etree.SubElement(tc, qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    reset_table_cell_paragraph_spacing(pPr)
    reset_table_cell_paragraph_indent(pPr)
    set_alignment(pPr, "center" if header else "left")
    run = make_run(text, size=24 if header else 22, bold=header, color="FFFFFF" if header else "000000")
    para.append(run)
    if header:
        shd = etree.SubElement(tcPr, qn("shd"))
        shd.set(qn("val"), "clear")
        shd.set(qn("color"), "auto")
        shd.set(qn("fill"), "0070C0")
    return tc


def make_connection_info_table():
    tbl = etree.Element(qn("tbl"))
    tblPr = etree.SubElement(tbl, qn("tblPr"))
    tblW = etree.SubElement(tblPr, qn("tblW"))
    tblW.set(qn("w"), "5000")
    tblW.set(qn("type"), "pct")
    jc = etree.SubElement(tblPr, qn("jc"))
    jc.set(qn("val"), "center")
    tblInd = etree.SubElement(tblPr, qn("tblInd"))
    tblInd.set(qn("w"), "0")
    tblInd.set(qn("type"), "dxa")
    tblBorders = etree.SubElement(tblPr, qn("tblBorders"))
    for border in ("top", "left", "bottom", "right", "insideH", "insideV"):
        set_border(tblBorders, border, size="2", color="D9D9D9")
    tblGrid = etree.SubElement(tbl, qn("tblGrid"))
    for width in ("1800", "2600", "3600"):
        gridCol = etree.SubElement(tblGrid, qn("gridCol"))
        gridCol.set(qn("w"), width)
    rows = [
        (True, ["No", "Name", "Value", "Description"]),
        (False, ["1", "domain", "prd-gopaperless.mobile-id.vn", "Current production domain"]),
        (False, ["2", "contextPath", "/workflow/api", "API context path"]),
        (False, ["3", "baseUrl", "https://{domain}/{contextPath}/XXX", "Template for API endpoint URLs"]),
    ]
    for is_header, values in rows:
        tr = etree.SubElement(tbl, qn("tr"))
        if is_header:
            trPr = etree.SubElement(tr, qn("trPr"))
            etree.SubElement(trPr, qn("tblHeader"))
        for value in values:
            tr.append(make_table_cell(value, header=is_header))
    return tbl


def set_spacing(pPr, space_before=None, space_after=None):
    """Set spaceBefore and spaceAfter in twips (1 pt = 20 twips)."""
    spacing = get_or_create(pPr, "spacing")
    if space_before is not None:
        spacing.set(qn("before"), str(int(space_before * 20)))
    if space_after is not None:
        spacing.set(qn("after"), str(int(space_after * 20)))


def set_indent(pPr, left_indent=None, hanging_indent=None):
    """Set left and/or hanging indent in twips."""
    ind = get_or_create(pPr, "ind")
    if left_indent is not None:
        ind.set(qn("left"), str(int(left_indent * 20)))
        for attr in (qn("firstLine"), qn("hanging")):
            if attr in ind.attrib:
                del ind.attrib[attr]
    if hanging_indent is not None:
        # When hanging is set, ensure left is also set (default to 0 if not provided)
        if left_indent is None and qn("left") not in ind.attrib:
            ind.set(qn("left"), "0")
        ind.set(qn("hanging"), str(int(hanging_indent * 20)))
        if qn("firstLine") in ind.attrib:
            del ind.attrib[qn("firstLine")]


def remove_numbering(pPr):
    numPr = pPr.find(qn("numPr"))
    if numPr is not None:
        pPr.remove(numPr)
    tabs = pPr.find(qn("tabs"))
    if tabs is not None:
        pPr.remove(tabs)


def ensure_rPr_first(run):
    """Ensure rPr is the first child of run. Returns rPr element."""
    rPr = run.find(qn("rPr"))
    if rPr is None:
        rPr = etree.Element(qn("rPr"))
        if len(run):
            run.insert(0, rPr)
        else:
            run.append(rPr)
    else:
        if run.index(rPr) != 0:
            run.remove(rPr)
            run.insert(0, rPr)
    return rPr


def rebuild_rPr(run, font_name, size, bold, color=None):
    """Remove old rPr and create a clean one at first position."""
    old = run.find(qn("rPr"))
    if old is not None:
        run.remove(old)
    rPr = etree.Element(qn("rPr"))
    set_fonts(rPr, font_name)
    set_size(rPr, size)
    set_bold(rPr, bold)
    set_color(rPr, color)
    if len(run):
        run.insert(0, rPr)
    else:
        run.append(rPr)


# ---------------------------------------------------------------------------
# Fixers
# ---------------------------------------------------------------------------

def fix_styles(styles_path, config):
    tree = etree.parse(styles_path)
    root = tree.getroot()

    for style in root.iter(qn("style")):
        sid = style.get(qn("styleId"))
        if sid not in config:
            continue
        cfg = config[sid]

        # Fix rPr (font, size, bold, color)
        rPr = style.find(qn("rPr"))
        if rPr is None:
            rPr = etree.SubElement(style, qn("rPr"))
        if cfg.get("font"):
            set_fonts(rPr, cfg["font"])
        if cfg.get("sz") is not None:
            set_size(rPr, cfg["sz"])
        if cfg.get("bold") is not None:
            set_bold(rPr, cfg["bold"])
        if cfg.get("color") is not None:
            set_color(rPr, cfg["color"])
        else:
            set_color(rPr, None)  # remove explicit color if template says Auto

        # Fix pPr (alignment, spacing, indent)
        pPr = style.find(qn("pPr"))
        if pPr is None:
            pPr = etree.SubElement(style, qn("pPr"))
        if cfg.get("alignment"):
            set_alignment(pPr, cfg["alignment"])
        if cfg.get("space_before") is not None or cfg.get("space_after") is not None:
            set_spacing(pPr, cfg.get("space_before"), cfg.get("space_after"))
        if cfg.get("left_indent") is not None or cfg.get("hanging_indent") is not None:
            set_indent(pPr, cfg.get("left_indent"), cfg.get("hanging_indent"))
        # Normal style gets Multiple 1.15 line spacing
        if sid == "Normal":
            spacing = get_or_create(pPr, "spacing")
            spacing.set(qn("line"), "276")
            spacing.set(qn("lineRule"), "auto")
        if sid in {"Heading1", "Heading2", "Heading3", "Heading4"}:
            numPr = get_or_create(pPr, "numPr")
            for child in list(numPr):
                numPr.remove(child)
            level_by_style = {"Heading1": "0", "Heading2": "1", "Heading3": "2", "Heading4": "3"}
            ilvl = etree.SubElement(numPr, qn("ilvl"))
            ilvl.set(qn("val"), level_by_style[sid])
            num_id = etree.SubElement(numPr, qn("numId"))
            num_id.set(qn("val"), "1")

    # Also fix linked character styles (e.g. Heading1Char) so they stay in sync
    char_configs = {}
    for style in root.iter(qn("style")):
        sid = style.get(qn("styleId"))
        link_el = style.find(qn("link"))
        if link_el is not None and sid.endswith("Char"):
            linked_val = link_el.get(qn("val"))
            if linked_val in config:
                char_configs[sid] = config[linked_val]

    for style in root.iter(qn("style")):
        sid = style.get(qn("styleId"))
        if sid not in char_configs:
            continue
        cfg = char_configs[sid]
        rPr = style.find(qn("rPr"))
        if rPr is None:
            rPr = etree.SubElement(style, qn("rPr"))
        if cfg.get("font"):
            set_fonts(rPr, cfg["font"])
        if cfg.get("sz") is not None:
            set_size(rPr, cfg["sz"])
        if cfg.get("bold") is not None:
            set_bold(rPr, cfg["bold"])
        if cfg.get("color") is not None:
            set_color(rPr, cfg["color"])
        else:
            set_color(rPr, None)

    tree.write(styles_path, xml_declaration=True, encoding="utf-8", standalone=True)
    print(f"[OK] Fixed styles.xml with {len(config)} style definitions from template")


def set_border(parent, tag, val="single", size="2", color="000000"):
    border = parent.find(qn(tag))
    if border is None:
        border = etree.SubElement(parent, qn(tag))
    border.set(qn("val"), val)
    if val != "nil":
        border.set(qn("sz"), size)
        border.set(qn("space"), "0")
        border.set(qn("color"), color)


def get_element_text(element):
    return "".join(t.text or "" for t in element.iter(qn("t"))).strip()


def apply_table_border_rules(root, table_border_rules):
    body = root.find(".//" + qn("body"))
    if body is None:
        return
    children = list(body)
    rules_by_heading = {rule['section_heading']: rule for rule in table_border_rules}
    for idx, child in enumerate(children):
        if child.tag != qn("p"):
            continue
        rule = rules_by_heading.get(get_element_text(child))
        if not rule:
            continue
        for candidate in children[idx + 1:]:
            if candidate.tag != qn("tbl"):
                continue
            tblPr = get_or_create(candidate, "tblPr")
            tblBorders = get_or_create(tblPr, "tblBorders")
            for tag in rule['table_borders']:
                set_border(tblBorders, tag, size=rule['border_size'], color=rule['border_color'])
            if rule['apply_cell_borders']:
                for tc in candidate.iter(qn("tc")):
                    tcPr = get_or_create(tc, "tcPr")
                    tcBorders = get_or_create(tcPr, "tcBorders")
                    for tag in rule['cell_borders']:
                        set_border(tcBorders, tag, size=rule['border_size'], color=rule['border_color'])
            break


def apply_paragraph_style_layout(pPr, cfg):
    if cfg.get("alignment"):
        set_alignment(pPr, cfg["alignment"])
    if cfg.get("space_before") is not None or cfg.get("space_after") is not None:
        set_spacing(pPr, cfg.get("space_before"), cfg.get("space_after"))
    if cfg.get("left_indent") is not None or cfg.get("hanging_indent") is not None:
        set_indent(pPr, cfg.get("left_indent"), cfg.get("hanging_indent"))


def get_child_index(parent, child):
    return list(parent).index(child)


def is_frame_paragraph(para):
    """Detect if a paragraph is a framed code block (shading/border/shape)."""
    pPr = para.find(qn("pPr"))
    if pPr is not None:
        if pPr.find(qn("framePr")) is not None:
            return True
        if pPr.find(qn("pBdr")) is not None:
            return True
        shd = pPr.find(qn("shd"))
        if shd is not None:
            fill = shd.get(qn("fill"))
            if fill and fill.lower() not in ("auto", "none", "000000"):
                return True
    for r in para.iter(qn("r")):
        if r.find(qn("pict")) is not None or r.find(".//" + qn("drawing")) is not None:
            return True
    return False


def split_label_from_frame_paragraphs(root):
    """Split 'Sample Request' / 'Sample Response' labels out of paragraphs
    that also contain an inline drawing/frame, so the label can be left-aligned
    while the frame stays centered."""
    LABELS = {"Sample Request", "Sample Response"}
    body = root.find(".//" + qn("body"))
    if body is None:
        return

    for para in list(body.iter(qn("p"))):
        para_text = get_element_text(para)
        if not any(para_text.startswith(label) for label in LABELS):
            continue

        runs = list(para.findall(qn("r")))
        frame_run_idx = None
        for idx, r in enumerate(runs):
            if r.find(qn("pict")) is not None or r.find(".//" + qn("drawing")) is not None:
                frame_run_idx = idx
                break
        if frame_run_idx is None or frame_run_idx == 0:
            continue

        label_runs = runs[:frame_run_idx]

        new_para = etree.Element(qn("p"))
        new_pPr = etree.SubElement(new_para, qn("pPr"))
        pStyle = etree.SubElement(new_pPr, qn("pStyle"))
        pStyle.set(qn("val"), "Normal")
        set_alignment(new_pPr, "left")
        set_indent(new_pPr, left_indent=360)
        set_spacing(new_pPr, space_before=0, space_after=3)
        spacing = get_or_create(new_pPr, "spacing")
        spacing.set(qn("line"), "276")
        spacing.set(qn("lineRule"), "auto")

        for r in label_runs:
            new_para.append(deepcopy(r))

        para.addprevious(new_para)

        for r in label_runs:
            para.remove(r)

        pPr = para.find(qn("pPr"))
        if pPr is None:
            pPr = etree.Element(qn("pPr"))
            para.insert(0, pPr)
        set_alignment(pPr, "center")
        set_indent(pPr, left_indent=0)
        set_spacing(pPr, space_before=0, space_after=3)

    print("[OK] Split mixed label+frame paragraphs")


def make_blank_paragraph_like(reference_para, config, rule):
    para = etree.Element(qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    pStyle = etree.SubElement(pPr, qn("pStyle"))
    pStyle.set(qn("val"), rule.get("paragraph_type") or "Normal")
    apply_paragraph_style_layout(pPr, config["Normal"])
    set_spacing(pPr, rule.get("space_before"), rule.get("space_after"))
    if rule.get("left_indent") is not None:
        set_indent(pPr, rule.get("left_indent") * 72, None)
    spacing = get_or_create(pPr, "spacing")
    spacing.set(qn("line"), "276")
    spacing.set(qn("lineRule"), "auto")
    if reference_para is not None:
        sectPr = reference_para.find(qn("pPr") + "/" + qn("sectPr"))
        if sectPr is not None:
            pPr.append(deepcopy(sectPr))
    return para


def normalize_blank_paragraph(para, config, rule):
    pPr = para.find(qn("pPr"))
    if pPr is None:
        pPr = etree.Element(qn("pPr"))
        para.insert(0, pPr)
    pStyle = pPr.find(qn("pStyle"))
    if pStyle is None:
        pStyle = etree.Element(qn("pStyle"))
        pPr.insert(0, pStyle)
    pStyle.set(qn("val"), rule.get("paragraph_type") or "Normal")
    apply_paragraph_style_layout(pPr, config["Normal"])
    set_spacing(pPr, rule.get("space_before"), rule.get("space_after"))
    if rule.get("left_indent") is not None:
        set_indent(pPr, rule.get("left_indent") * 72, None)
    spacing = get_or_create(pPr, "spacing")
    spacing.set(qn("line"), "276")
    spacing.set(qn("lineRule"), "auto")
    for run in list(para.findall(qn("r"))):
        para.remove(run)


def parse_hex_color(value):
    m = re.search(r'`?([0-9A-Fa-f]{6})`?', value or '')
    return m.group(1).upper() if m else None


def normalize_api_table_start_headers(root, table_property_rules):
    start_headers = {
        "Header Attributes",
        "Path Attributes",
        "Request Attributes",
        "Response Attributes",
        "Header Request Attributes",
        "Header Response Attributes",
    }
    color = parse_hex_color(table_property_rules.get("Attribute table start header font color")) or "FFFFFF"
    header_width = parse_int_property(table_property_rules, "Attribute table start header width")
    header_width_type = table_property_rules.get("Attribute table start header width type")
    header_grid_span = parse_int_property(table_property_rules, "Attribute table start header grid span")
    header_row_height = parse_int_property(table_property_rules, "Attribute table start header row height")
    header_row_height_rule = table_property_rules.get("Attribute table start header row height rule")
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        if not rows:
            continue
        first_row = rows[0]
        cells = first_row.findall(qn("tc"))
        if not cells:
            continue
        if get_element_text(cells[0]) not in start_headers:
            continue
        if header_row_height is not None and header_row_height_rule:
            trPr = get_or_create(first_row, "trPr")
            trHeight = get_or_create(trPr, "trHeight")
            trHeight.set(qn("val"), str(header_row_height))
            trHeight.set(qn("hRule"), header_row_height_rule)
        if header_width is not None and header_width_type:
            tblPr = get_or_create(tbl, "tblPr")
            tblW = get_or_create(tblPr, "tblW")
            tblW.set(qn("w"), str(header_width))
            tblW.set(qn("type"), header_width_type)
            tcPr = get_or_create(cells[0], "tcPr")
            if header_grid_span is not None:
                grid_span = get_or_create(tcPr, "gridSpan")
                grid_span.set(qn("val"), str(header_grid_span))
            tcW = get_or_create(tcPr, "tcW")
            tcW.set(qn("w"), str(header_width))
            tcW.set(qn("type"), header_width_type)
        for tc in cells:
            tcPr = get_or_create(tc, "tcPr")
            set_cell_valign(tcPr, "center")
            for para in tc.iter(qn("p")):
                pPr = para.find(qn("pPr"))
                if pPr is None:
                    pPr = etree.Element(qn("pPr"))
                    para.insert(0, pPr)
                set_alignment(pPr, "left")
                reset_table_cell_paragraph_spacing(pPr)
                reset_table_cell_paragraph_indent(pPr)
                for run in para.iter(qn("r")):
                    rPr = ensure_rPr_first(run)
                    set_bold(rPr, True)
                    set_color(rPr, color)


def rename_header_attribute_tables(root):
    """Rename generic 'Header Attributes' tables to 'Header Request Attributes'
    or 'Header Response Attributes' based on position within each API section."""
    body = root.find(".//" + qn("body"))
    if body is None:
        return

    def is_heading_break(node):
        if node.tag != qn("p"):
            return False
        pPr = node.find(qn("pPr"))
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None and ps.get(qn("val")) in {"Heading1", "Heading2", "Heading3", "Heading4"}:
                return True
        return False

    def get_table_header_text(tbl):
        rows = tbl.findall(qn("tr"))
        if not rows:
            return ""
        cells = rows[0].findall(qn("tc"))
        if not cells:
            return ""
        return get_element_text(cells[0]).strip()

    def set_table_header_text(tbl, new_text):
        rows = tbl.findall(qn("tr"))
        if not rows:
            return
        cells = rows[0].findall(qn("tc"))
        if not cells:
            return
        for para in cells[0].iter(qn("p")):
            for t in para.iter(qn("t")):
                if t.text and t.text.strip():
                    t.text = new_text
                    return

    def process_section_tables(tables):
        header_items = [(i, tbl) for i, tbl, htext in tables if htext == "Header Attributes"]
        if len(header_items) == 1:
            set_table_header_text(header_items[0][1], "Header Request Attributes")
        elif len(header_items) >= 2:
            set_table_header_text(header_items[0][1], "Header Request Attributes")
            set_table_header_text(header_items[-1][1], "Header Response Attributes")

    children = list(body)
    section_tables = []
    in_section = False

    for idx, child in enumerate(children):
        if is_heading_break(child):
            pPr = child.find(qn("pPr"))
            ps = pPr.find(qn("pStyle")) if pPr is not None else None
            style = ps.get(qn("val")) if ps is not None else ""
            if style in {"Heading3", "Heading4"}:
                # Process previous section if any
                if in_section and section_tables:
                    process_section_tables(section_tables)
                in_section = True
                section_tables = []
            elif style in {"Heading1", "Heading2"}:
                if in_section and section_tables:
                    process_section_tables(section_tables)
                in_section = False
                section_tables = []
            else:
                # Another Heading3/Heading4 within a section also starts a new sub-section
                if in_section and section_tables:
                    process_section_tables(section_tables)
                section_tables = []
        if in_section and child.tag == qn("tbl"):
            htext = get_table_header_text(child)
            if htext in {"Header Attributes", "Header Request Attributes", "Header Response Attributes",
                         "Path Attributes", "Request Attributes", "Response Attributes"}:
                section_tables.append((idx, child, htext))

    # Process final section
    if in_section and section_tables:
        process_section_tables(section_tables)


def make_path_attributes_table(path_vars):
    """Create a Path Attributes table for the given path variables."""
    tbl = etree.Element(qn("tbl"))
    tblPr = etree.SubElement(tbl, qn("tblPr"))
    tblW = etree.SubElement(tblPr, qn("tblW"))
    tblW.set(qn("w"), "5000")
    tblW.set(qn("type"), "pct")
    jc = etree.SubElement(tblPr, qn("jc"))
    jc.set(qn("val"), "center")
    tblInd = etree.SubElement(tblPr, qn("tblInd"))
    tblInd.set(qn("w"), "0")
    tblInd.set(qn("type"), "dxa")
    tblBorders = etree.SubElement(tblPr, qn("tblBorders"))
    for border in ("top", "left", "bottom", "right", "insideH", "insideV"):
        set_border(tblBorders, border, size="2", color="D9D9D9")
    tblGrid = etree.SubElement(tbl, qn("tblGrid"))
    for w in ("900", "2200", "1400", "1200", "3300"):
        gridCol = etree.SubElement(tblGrid, qn("gridCol"))
        gridCol.set(qn("w"), w)

    # Row 1: merged header "Path Attributes"
    tr1 = etree.SubElement(tbl, qn("tr"))
    trPr1 = etree.SubElement(tr1, qn("trPr"))
    trHeight1 = etree.SubElement(trPr1, qn("trHeight"))
    trHeight1.set(qn("val"), "420")
    trHeight1.set(qn("hRule"), "atLeast")
    tc1 = make_table_cell("Path Attributes", header=True)
    tcPr1 = tc1.find(qn("tcPr"))
    grid_span1 = get_or_create(tcPr1, "gridSpan")
    grid_span1.set(qn("val"), "5")
    tcW1 = tcPr1.find(qn("tcW"))
    if tcW1 is not None:
        tcW1.set(qn("w"), "5000")
        tcW1.set(qn("type"), "pct")
    tr1.append(tc1)

    # Row 2: column headers
    tr2 = etree.SubElement(tbl, qn("tr"))
    trPr2 = etree.SubElement(tr2, qn("trPr"))
    trHeight2 = etree.SubElement(trPr2, qn("trHeight"))
    trHeight2.set(qn("val"), "420")
    trHeight2.set(qn("hRule"), "atLeast")
    for h in ("No", "Name", "Type", "Presence", "Description"):
        tr2.append(make_table_cell(h, header=True))

    # Data rows
    for no, var in enumerate(path_vars, start=1):
        tr = etree.SubElement(tbl, qn("tr"))
        trPr = etree.SubElement(tr, qn("trPr"))
        trHeight = etree.SubElement(trPr, qn("trHeight"))
        trHeight.set(qn("val"), "420")
        trHeight.set(qn("hRule"), "atLeast")
        tr.append(make_table_cell(str(no), header=False))
        tr.append(make_table_cell(var, header=False))
        tr.append(make_table_cell("String", header=False))
        tr.append(make_table_cell("M", header=False))
        tr.append(make_table_cell("", header=False))

    return tbl


def create_path_attribute_tables(root, config, api_body_rules):
    """Auto-create Path Attributes tables for API sections that have URL path variables
    but are missing a Path Attributes table."""
    body = root.find(".//" + qn("body"))
    if body is None:
        return

    def is_heading_break(node):
        if node.tag != qn("p"):
            return False
        pPr = node.find(qn("pPr"))
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None and ps.get(qn("val")) in {"Heading1", "Heading2", "Heading3", "Heading4"}:
                return True
        return False

    def get_table_header_text(tbl):
        rows = tbl.findall(qn("tr"))
        if not rows:
            return ""
        cells = rows[0].findall(qn("tc"))
        if not cells:
            return ""
        return get_element_text(cells[0]).strip()

    def extract_path_vars(section_nodes):
        found_request = False
        for node in section_nodes:
            if node.tag == qn("p"):
                text = get_element_text(node)
                if text == "Sample Request":
                    found_request = True
                    continue
                if found_request and text.strip():
                    # Look for HTTP method + URL pattern
                    m = re.search(r'(https?://\{[^}]+\}/\{[^}]+\})(/[^?\s]*)', text)
                    if m:
                        path_part = m.group(2)
                        vars_found = re.findall(r'\{([^}]+)\}', path_part)
                        # Filter out domain/contextPath just in case
                        vars_found = [v for v in vars_found if v not in ("domain", "contextPath")]
                        return vars_found
                    # Fallback: any URL-looking segment with variables
                    m2 = re.search(r'/(v\d+[^?\s]*)', text)
                    if m2:
                        path_part = m2.group(1)
                        vars_found = re.findall(r'\{([^}]+)\}', "/" + path_part)
                        vars_found = [v for v in vars_found if v not in ("domain", "contextPath")]
                        return vars_found
                    # If we hit Sample Response or Attributes description, stop looking
                    if text in {"Sample Response", "Attributes description"}:
                        break
        return []

    children = list(body)
    idx = 0
    while idx < len(children):
        child = children[idx]
        if is_heading_break(child):
            pPr = child.find(qn("pPr"))
            ps = pPr.find(qn("pStyle")) if pPr is not None else None
            style = ps.get(qn("val")) if ps is not None else ""
            if style in {"Heading3", "Heading4"}:
                # Gather section nodes
                section_nodes = []
                j = idx + 1
                while j < len(children):
                    n = children[j]
                    if is_heading_break(n):
                        pPr2 = n.find(qn("pPr"))
                        ps2 = pPr2.find(qn("pStyle")) if pPr2 is not None else None
                        s2 = ps2.get(qn("val")) if ps2 is not None else ""
                        if s2 in {"Heading1", "Heading2", "Heading3", "Heading4"}:
                            break
                    section_nodes.append(n)
                    j += 1

                # Check if section already has Path Attributes
                has_path = False
                last_header_idx = None
                for k, node in enumerate(section_nodes):
                    if node.tag == qn("tbl"):
                        htext = get_table_header_text(node)
                        if htext == "Path Attributes":
                            has_path = True
                            break
                        if htext in {"Header Attributes", "Header Request Attributes"}:
                            last_header_idx = k

                if not has_path:
                    path_vars = extract_path_vars(section_nodes)
                    if path_vars:
                        # Insert after the last Header Attributes table if found
                        insert_offset = 0
                        if last_header_idx is not None:
                            insert_offset = last_header_idx + 1
                        real_insert = idx + 1 + insert_offset
                        # Adjust for any previously inserted tables in this loop
                        # (simplified: just insert)
                        path_tbl = make_path_attributes_table(path_vars)
                        blank = make_blank_paragraph_like(None, config, api_body_rules.get("Between attribute tables", {}))
                        body.insert(real_insert, blank)
                        body.insert(real_insert + 1, path_tbl)
                        children = list(body)  # refresh
                        idx = real_insert + 2
                        continue
        idx += 1


def apply_table_alignment(root, table_property_rules):
    """Set table justification and indent for ALL tables in the document."""
    alignment = table_property_rules.get("Table alignment", "center").lower()
    indent_str = table_property_rules.get("Table indent", "0 dxa")
    m = re.search(r'(\d+)', indent_str)
    indent_val = m.group(1) if m else "0"
    indent_type = "pct" if "pct" in indent_str.lower() else "dxa"

    for tbl in root.iter(qn("tbl")):
        tblPr = get_or_create(tbl, "tblPr")
        jc = get_or_create(tblPr, "jc")
        jc.set(qn("val"), alignment)
        tblInd = get_or_create(tblPr, "tblInd")
        tblInd.set(qn("w"), indent_val)
        tblInd.set(qn("type"), indent_type)


def normalize_inline_shapes(root, target_cx="6664325"):
    """Center-align paragraphs containing inline text-box shapes and unify their width."""
    txbx_ns = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}txbx"
    vml_txbx_ns = "{urn:schemas-microsoft-com:office:word}txbx"
    wp_extent_ns = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
    a_ext_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}ext"

    for para in root.iter(qn("p")):
        has_textbox = False
        for drawing in para.iter(qn("drawing")):
            if drawing.find(f".//{txbx_ns}") is not None or drawing.find(f".//{vml_txbx_ns}") is not None:
                has_textbox = True
                # Unify width in wp:extent
                for extent in drawing.iter(wp_extent_ns):
                    extent.set("cx", target_cx)
                # Unify width in a:xfrm -> a:ext
                for ext in drawing.iter(a_ext_ns):
                    ext.set("cx", target_cx)
                break

        if has_textbox:
            pPr = para.find(qn("pPr"))
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                para.insert(0, pPr)
            set_alignment(pPr, "center")
            reset_table_cell_paragraph_indent(pPr)


def normalize_api_body_spacing(root, config, api_body_rules):
    body = root.find(".//" + qn("body"))
    if body is None:
        return

    def children():
        return list(body)

    def is_text_para(node, text):
        return node.tag == qn("p") and get_element_text(node) == text

    def is_blank_para(node):
        return node.tag == qn("p") and get_element_text(node) == ""

    def is_bookmark(node):
        return node.tag in {qn("bookmarkStart"), qn("bookmarkEnd")}

    def is_heading_para(node):
        if node.tag != qn("p"):
            return False
        ps = node.find(qn("pPr") + "/" + qn("pStyle"))
        return ps is not None and ps.get(qn("val")) in {"Heading1", "Heading2", "Heading3", "Heading4"}

    def normalize_blank_at(index):
        current = children()
        rule = api_body_rules.get("Before Sample Response", {})
        if index < len(current):
            next_text = get_element_text(current[index])
            if next_text == api_body_rules.get("Attributes description label", {}).get("text_match"):
                rule = api_body_rules.get("Before Attributes description", rule)
            elif current[index].tag == qn("tbl"):
                rule = api_body_rules.get("Between attribute tables", rule)
        if index >= len(current):
            body.append(make_blank_paragraph_like(None, config, rule))
            return
        if is_blank_para(current[index]):
            normalize_blank_paragraph(current[index], config, rule)
            remove_at = index + 1
            while remove_at < len(children()) and is_blank_para(children()[remove_at]):
                body.remove(children()[remove_at])
            return
        body.insert(index, make_blank_paragraph_like(current[index - 1] if index else None, config, rule))

    label_rules = [
        api_body_rules.get("Sample Request label", {}),
        api_body_rules.get("Sample Response label", {}),
        api_body_rules.get("Attributes description label", {}),
    ]
    body_labels = {rule.get("text_match") for rule in label_rules if rule.get("text_match")}
    rules_by_text = {rule.get("text_match"): rule for rule in label_rules if rule.get("text_match")}
    for para in root.iter(qn("p")):
        para_text = get_element_text(para)
        if para_text not in body_labels:
            continue
        rule = rules_by_text[para_text]
        pPr = para.find(qn("pPr"))
        if pPr is None:
            pPr = etree.Element(qn("pPr"))
            para.insert(0, pPr)
        apply_paragraph_style_layout(pPr, config["Normal"])
        set_spacing(pPr, rule.get("space_before"), rule.get("space_after"))
        if rule.get("left_indent") is not None:
            set_indent(pPr, rule.get("left_indent") * 72, None)
        spacing = get_or_create(pPr, "spacing")
        spacing.set(qn("line"), "276")
        spacing.set(qn("lineRule"), "auto")

    request_label = api_body_rules.get("Sample Request label", {}).get("text_match", "Sample Request")
    response_label = api_body_rules.get("Sample Response label", {}).get("text_match", "Sample Response")
    attrs_label = api_body_rules.get("Attributes description label", {}).get("text_match", "Attributes description")
    before_request_rule = api_body_rules.get("Before Sample Request", api_body_rules.get("Before Sample Response", {}))
    request_frame_rule = api_body_rules.get("Sample Request frame", {})
    response_frame_rule = api_body_rules.get("Sample Response frame(s)", {})

    idx = 0
    while idx < len(children()):
        current = children()
        node = current[idx]
        if is_text_para(node, request_label):
            prev_idx = idx - 1
            while prev_idx >= 0 and is_bookmark(children()[prev_idx]):
                prev_idx -= 1
            if prev_idx >= 0:
                prev = children()[prev_idx]
                if is_blank_para(prev):
                    normalize_blank_paragraph(prev, config, before_request_rule)
                    remove_idx = prev_idx - 1
                    while remove_idx >= 0 and is_blank_para(children()[remove_idx]):
                        body.remove(children()[remove_idx])
                        remove_idx -= 1
                elif not is_heading_para(prev):
                    body.insert(idx, make_blank_paragraph_like(prev, config, before_request_rule))
                    idx += 1
            next_idx = idx + 1
            while next_idx < len(children()) and is_bookmark(children()[next_idx]):
                next_idx += 1
            if next_idx < len(children()) and children()[next_idx].tag == qn("p") and get_element_text(children()[next_idx]) not in {"", response_label}:
                request_frame = children()[next_idx]
                pPr = request_frame.find(qn("pPr"))
                if pPr is None:
                    pPr = etree.Element(qn("pPr"))
                    request_frame.insert(0, pPr)
                set_spacing(pPr, request_frame_rule.get("space_before"), request_frame_rule.get("space_after"))
                blank_idx = get_child_index(body, request_frame) + 1
                while blank_idx < len(children()) and is_bookmark(children()[blank_idx]):
                    blank_idx += 1
                normalize_blank_at(blank_idx)
        elif is_text_para(node, response_label):
            next_idx = idx + 1
            while next_idx < len(children()) and (is_bookmark(children()[next_idx]) or is_blank_para(children()[next_idx])):
                if is_blank_para(children()[next_idx]):
                    body.remove(children()[next_idx])
                else:
                    next_idx += 1
            while next_idx < len(children()):
                candidate = children()[next_idx]
                if is_bookmark(candidate):
                    next_idx += 1
                    continue
                if candidate.tag != qn("p") or get_element_text(candidate) == attrs_label or is_heading_para(candidate):
                    break
                if get_element_text(candidate) == "" and not is_frame_paragraph(candidate):
                    break
                pPr = candidate.find(qn("pPr"))
                if pPr is None:
                    pPr = etree.Element(qn("pPr"))
                    candidate.insert(0, pPr)
                set_spacing(pPr, response_frame_rule.get("space_before"), response_frame_rule.get("space_after"))
                next_idx += 1
            normalize_blank_at(next_idx)
        elif is_text_para(node, attrs_label):
            next_idx = idx + 1
            while next_idx < len(children()):
                candidate = children()[next_idx]
                if is_heading_para(candidate):
                    break
                if candidate.tag == qn("tbl"):
                    blank_idx = get_child_index(body, candidate) + 1
                    after_blank = blank_idx
                    while after_blank < len(children()) and is_bookmark(children()[after_blank]):
                        after_blank += 1
                    if after_blank < len(children()) and not is_heading_para(children()[after_blank]):
                        normalize_blank_at(after_blank)
                    next_idx = after_blank + 1
                else:
                    next_idx += 1
        idx += 1

    # Final pass: ensure exactly one blank paragraph before each
    # "Attributes description" label when preceded by non-heading content.
    for idx in range(len(children()) - 1, -1, -1):
        node = children()[idx]
        if not is_text_para(node, attrs_label):
            continue
        prev_idx = idx - 1
        while prev_idx >= 0 and is_bookmark(children()[prev_idx]):
            prev_idx -= 1
        if prev_idx < 0:
            continue
        prev = children()[prev_idx]
        if is_blank_para(prev):
            rule = api_body_rules.get("Before Attributes description", api_body_rules.get("Before Sample Response", {}))
            normalize_blank_paragraph(prev, config, rule)
            remove_idx = prev_idx - 1
            while remove_idx >= 0 and is_blank_para(children()[remove_idx]):
                body.remove(children()[remove_idx])
                remove_idx -= 1
            continue
        if is_heading_para(prev):
            continue
        rule = api_body_rules.get("Before Attributes description", api_body_rules.get("Before Sample Response", {}))
        body.insert(idx, make_blank_paragraph_like(prev, config, rule))


def normalize_heading3_spacing(root, config):
    body = root.find(".//" + qn("body"))
    if body is None:
        return

    def children():
        return list(body)

    def is_blank_para(node):
        return node.tag == qn("p") and get_element_text(node) == ""

    def is_bookmark(node):
        return node.tag in {qn("bookmarkStart"), qn("bookmarkEnd")}

    def is_heading2(node):
        if node.tag != qn("p"):
            return False
        ps = node.find(qn("pPr") + "/" + qn("pStyle"))
        return ps is not None and ps.get(qn("val")) == "Heading2"

    def is_heading3(node):
        if node.tag != qn("p"):
            return False
        ps = node.find(qn("pPr") + "/" + qn("pStyle"))
        return ps is not None and ps.get(qn("val")) == "Heading3"

    def make_standard_blank(ref):
        para = etree.Element(qn("p"))
        pPr = etree.SubElement(para, qn("pPr"))
        pStyle = etree.SubElement(pPr, qn("pStyle"))
        pStyle.set(qn("val"), "Normal")
        apply_paragraph_style_layout(pPr, config["Normal"])
        set_spacing(pPr, 0, 0)
        spacing = get_or_create(pPr, "spacing")
        spacing.set(qn("line"), "276")
        spacing.set(qn("lineRule"), "auto")
        if ref is not None:
            sectPr = ref.find(qn("pPr") + "/" + qn("sectPr"))
            if sectPr is not None:
                pPr.append(deepcopy(sectPr))
        return para

    for idx, node in enumerate(children()):
        if not is_heading3(node):
            continue
        # Skip the first Heading 3 that immediately follows a Heading 2
        prev_idx = idx - 1
        while prev_idx >= 0 and is_bookmark(children()[prev_idx]):
            prev_idx -= 1
        if prev_idx >= 0 and is_heading2(children()[prev_idx]):
            continue
        # Find the insertion point for the blank paragraph
        insert_idx = idx
        # Move back across any bookmarks that appear before this Heading 3
        while insert_idx > 0 and is_bookmark(children()[insert_idx - 1]):
            insert_idx -= 1
        if insert_idx > 0 and is_blank_para(children()[insert_idx - 1]):
            # Normalize the existing blank and remove extras
            blank = children()[insert_idx - 1]
            pPr = blank.find(qn("pPr"))
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                blank.insert(0, pPr)
            pStyle = pPr.find(qn("pStyle"))
            if pStyle is None:
                pStyle = etree.Element(qn("pStyle"))
                pPr.insert(0, pStyle)
            pStyle.set(qn("val"), "Normal")
            apply_paragraph_style_layout(pPr, config["Normal"])
            set_spacing(pPr, 0, 0)
            spacing = get_or_create(pPr, "spacing")
            spacing.set(qn("line"), "276")
            spacing.set(qn("lineRule"), "auto")
            for run in list(blank.findall(qn("r"))):
                blank.remove(run)
            # Remove any additional consecutive blanks before this one
            remove_idx = insert_idx - 2
            while remove_idx >= 0 and is_blank_para(children()[remove_idx]):
                body.remove(children()[remove_idx])
                remove_idx -= 1
        else:
            body.insert(insert_idx, make_standard_blank(children()[insert_idx - 1] if insert_idx > 0 else None))


def split_run_at(run, pos):
    """Split a w:r at a character position inside its w:t. Returns new run before split."""
    t = run.find(qn("t"))
    if t is None or pos <= 0 or pos >= len(t.text or ""):
        return None
    text = t.text
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

    new_run = etree.Element(qn("r"))
    rPr = run.find(qn("rPr"))
    if rPr is not None:
        new_run.append(deepcopy(rPr))
    new_t = etree.SubElement(new_run, qn("t"))
    if text[:pos].endswith(' ') or text[:pos].startswith(' '):
        new_t.set(XML_SPACE, "preserve")
    new_t.text = text[:pos]

    t.text = text[pos:]
    if t.text.endswith(' ') or t.text.startswith(' '):
        t.set(XML_SPACE, "preserve")
    return new_run


def _format_txbx_para(runs, method_rules):
    """Apply method badge, header bold and JSON formatting to a list of runs."""
    if not runs:
        return

    full_text = ""
    run_info = []
    for r in runs:
        # Skip runs that are the inline shape itself (drawing/pict)
        if r.find(qn("drawing")) is not None or r.find(qn("pict")) is not None:
            continue
        t = r.find(qn("t"))
        text = t.text or "" if t is not None else ""
        offset = len(full_text)
        full_text += text
        run_info.append({'offset': offset, 'text': text, 'run': r, 't': t})

    if not run_info:
        return

    # Detect HTTP method at the very start (handles missing space after method)
    method_re = re.match(r'^(\s*)(GET|POST|PUT|DELETE|PATCH)', full_text, re.IGNORECASE)
    if method_re:
        first_word = method_re.group(2).upper()
        method_offset = len(method_re.group(1))
        method_info = method_rules.get(first_word)
    else:
        first_word = ""
        method_offset = -1
        method_info = None

    # Detect header key: first ':' on a line that looks like an HTTP header
    header_end = -1
    if ':' in full_text:
        stripped = full_text.lstrip()
        colon_idx = full_text.index(':')
        key_text = full_text[:colon_idx].strip()
        # Only treat as header if key has no JSON characters
        starts_with_method = re.match(r'^(GET|POST|PUT|DELETE|PATCH)\b', stripped, re.IGNORECASE)
        if key_text and not starts_with_method and not any(c in key_text for c in ('{', '[', '"', "'")) and not stripped.startswith(('{', '[')):
            header_end = colon_idx

    # Detect URL/path region on first line (only when method badge is present)
    url_start = -1
    url_end = -1
    if method_info:
        url_start = method_offset + len(first_word)
        line1_end = full_text.find('\n')
        if line1_end == -1:
            line1_end = len(full_text)
        if url_start < line1_end:
            url_end = line1_end
        else:
            url_start = -1

    i = 0
    while i < len(run_info):
        info = run_info[i]
        r = info['run']
        t = info['t']
        if t is None:
            i += 1
            continue

        run_start = info['offset']
        run_end = run_start + len(info['text'])
        rPr = ensure_rPr_first(r)
        set_fonts(rPr, "Verdana")
        set_size(rPr, 20)
        set_highlight(rPr, None)
        set_shd(rPr, None)

        # Method badge: first token starts with method text
        if method_info and run_start == method_offset and info['text'].upper().startswith(first_word.upper()):
            method_len = len(first_word)
            if len(info['text']) > method_len:
                before_run = split_run_at(r, method_len)
                if before_run is not None:
                    parent = r.getparent()
                    idx = list(parent).index(r)
                    parent.insert(idx, before_run)
                    run_info.insert(i, {
                        'offset': run_start,
                        'text': info['text'][:method_len],
                        'run': before_run,
                        't': before_run.find(qn("t")),
                    })
                    run_info[i + 1]['offset'] = run_start + method_len
                    run_info[i + 1]['text'] = info['text'][method_len:]
                    # Format the method run (before_run)
                    before_rPr = ensure_rPr_first(before_run)
                    set_fonts(before_rPr, "Verdana")
                    set_size(before_rPr, 20)
                    set_bold(before_rPr, True)
                    set_color(before_rPr, method_info['text_color'])
                    set_shd(before_rPr, method_info['highlight'])
                    # Original run (r) becomes normal
                    set_bold(rPr, False)
                    set_color(rPr, "000000")
                    i += 1
                    continue
            set_bold(rPr, True)
            set_color(rPr, method_info['text_color'])
            set_shd(rPr, method_info['highlight'])
            i += 1
            continue

        # URL/path bold on first line of the frame
        if url_start >= 0 and url_end > url_start and run_start < url_end and run_end > url_start:
            if run_end > url_end:
                split_pos = url_end - run_start
                before_run = split_run_at(r, split_pos)
                if before_run is not None:
                    parent = r.getparent()
                    idx = list(parent).index(r)
                    parent.insert(idx, before_run)
                    run_info.insert(i, {
                        'offset': run_start,
                        'text': info['text'][:split_pos],
                        'run': before_run,
                        't': before_run.find(qn("t")),
                    })
                    run_info[i + 1]['offset'] = run_start + split_pos
                    run_info[i + 1]['text'] = info['text'][split_pos:]
                    before_rPr = ensure_rPr_first(before_run)
                    set_fonts(before_rPr, "Verdana")
                    set_size(before_rPr, 20)
                    set_bold(before_rPr, True)
                    set_color(before_rPr, "000000")
                    set_bold(rPr, False)
                    set_color(rPr, "000000")
                    i += 1
                    continue
            set_bold(rPr, True)
            set_color(rPr, "000000")
            i += 1
            continue

        # Header key bold (text before first ':')
        if header_end > 0 and run_start < header_end:
            if run_end > header_end:
                split_pos = header_end - run_start
                before_run = split_run_at(r, split_pos)
                if before_run is not None:
                    parent = r.getparent()
                    idx = list(parent).index(r)
                    parent.insert(idx, before_run)
                    run_info.insert(i, {
                        'offset': run_start,
                        'text': info['text'][:split_pos],
                        'run': before_run,
                        't': before_run.find(qn("t")),
                    })
                    run_info[i + 1]['offset'] = run_start + split_pos
                    run_info[i + 1]['text'] = info['text'][split_pos:]
                    before_rPr = ensure_rPr_first(before_run)
                    set_fonts(before_rPr, "Verdana")
                    set_size(before_rPr, 20)
                    set_bold(before_rPr, True)
                    set_color(before_rPr, "000000")
                    set_bold(rPr, False)
                    set_color(rPr, "000000")
                    i += 1
                    continue
            set_bold(rPr, True)
            set_color(rPr, "000000")
            i += 1
            continue

        set_bold(rPr, False)
        set_color(rPr, "000000")
        i += 1


def _highlight_variables_in_para(runs, variable_color="FF0000"):
    """Find variable placeholders like {domain}, {token} and color them red."""
    if not runs:
        return

    run_info = []
    for r in runs:
        if r.find(qn("drawing")) is not None or r.find(qn("pict")) is not None:
            continue
        t = r.find(qn("t"))
        text = t.text or "" if t is not None else ""
        run_info.append({'text': text, 'run': r, 't': t})

    i = 0
    while i < len(run_info):
        info = run_info[i]
        r = info['run']
        t = info['t']
        if t is None:
            i += 1
            continue

        text = t.text or ""
        match = re.search(r'\{[a-zA-Z0-9_-]+\}', text)
        if not match:
            i += 1
            continue

        var_start = match.start()
        var_end = match.end()

        # Split into [before | variable | after]
        # Step 1: split at var_end -> new_run=before+var, r=after
        if var_end < len(text):
            new_run = split_run_at(r, var_end)
            if new_run is not None:
                parent = r.getparent()
                idx = list(parent).index(r)
                parent.insert(idx, new_run)
                run_info[i] = {'text': text[:var_end], 'run': new_run, 't': new_run.find(qn("t"))}
                run_info.insert(i + 1, {'text': text[var_end:], 'run': r, 't': r.find(qn("t"))})
                r = new_run

        # Step 2: split at var_start -> before_run=before, r=variable
        if var_start > 0:
            before_run = split_run_at(r, var_start)
            if before_run is not None:
                parent = r.getparent()
                idx = list(parent).index(r)
                parent.insert(idx, before_run)
                run_info[i] = {'text': text[:var_start], 'run': before_run, 't': before_run.find(qn("t"))}
                run_info.insert(i + 1, {'text': text[var_start:var_end], 'run': r, 't': r.find(qn("t"))})
                i += 1

        # Apply red color to variable run
        rPr = ensure_rPr_first(r)
        set_color(rPr, variable_color)
        i += 1


def _try_repair_and_format_json(raw_text):
    """Try to parse JSON text; if it fails, attempt to repair missing commas
    and reformat with 2-space indentation."""
    import json

    # Direct parse
    try:
        obj = json.loads(raw_text)
        return json.dumps(obj, indent=2, ensure_ascii=False)
    except (json.JSONDecodeError, ValueError):
        pass

    # Repair: add missing commas between properties / array elements
    lines = raw_text.splitlines()
    repaired_lines = []
    for idx, line in enumerate(lines):
        stripped = line.rstrip()
        if not stripped:
            repaired_lines.append(line)
            continue

        # Structural bracket lines: keep as-is
        if stripped in ('{', '}', '[', ']', '},', '],'):
            repaired_lines.append(line)
            continue

        # Already ends with comma or opening bracket
        if stripped.endswith((',', '{', '[')):
            repaired_lines.append(line)
            continue

        # Find next non-empty line
        next_stripped = ''
        for j in range(idx + 1, len(lines)):
            n = lines[j].strip()
            if n:
                next_stripped = n
                break

        # If next line is a closing bracket, no comma needed
        if next_stripped and next_stripped.startswith(('}', ']')):
            repaired_lines.append(line)
            continue

        # Otherwise assume sibling element -> add comma
        repaired_lines.append(line + ',')

    repaired_text = '\n'.join(repaired_lines)
    try:
        obj = json.loads(repaired_text)
        return json.dumps(obj, indent=2, ensure_ascii=False)
    except (json.JSONDecodeError, ValueError):
        return None


def _make_json_para(text, template_pPr=None):
    """Create a paragraph for a single JSON line with Verdana 10pt black text."""
    para = etree.Element(qn("p"))
    if template_pPr is not None:
        para.append(deepcopy(template_pPr))
    else:
        pPr = etree.SubElement(para, qn("pPr"))
        set_spacing(pPr, 0, 0)

    run = etree.SubElement(para, qn("r"))
    rPr = etree.SubElement(run, qn("rPr"))
    set_fonts(rPr, "Verdana")
    set_size(rPr, 20)
    set_color(rPr, "000000")

    t = etree.SubElement(run, qn("t"))
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
    if text.startswith(" ") or text.endswith(" "):
        t.set(XML_SPACE, "preserve")
    t.text = text
    return para


def _prettify_json_in_txbx(txbx):
    """Detect JSON blocks inside a txbxContent and replace them with beautifully
    formatted paragraphs (2-space indentation, valid JSON)."""
    paragraphs = list(txbx.iter(qn("p")))
    i = 0
    while i < len(paragraphs):
        para = paragraphs[i]
        para_text = get_element_text(para).strip()

        # JSON block start: line begins with { or [
        if not para_text.startswith(('{', '[')):
            i += 1
            continue

        block_start = i
        block_end = i
        depth = 0
        for j in range(i, len(paragraphs)):
            txt = get_element_text(paragraphs[j]).strip()
            depth += txt.count('{') + txt.count('[')
            depth -= txt.count('}') + txt.count(']')
            block_end = j
            if depth <= 0 and txt.endswith(('}', ']')):
                break

        raw_text = '\n'.join(get_element_text(paragraphs[j]) for j in range(block_start, block_end + 1))
        formatted = _try_repair_and_format_json(raw_text)

        if formatted is not None:
            formatted_lines = formatted.splitlines()
            first_pPr = paragraphs[block_start].find(qn("pPr"))
            ref_node = paragraphs[block_end + 1] if block_end + 1 < len(paragraphs) else None

            # Remove old paragraphs in reverse
            for j in range(block_end, block_start - 1, -1):
                txbx.remove(paragraphs[j])

            # Insert new paragraphs in forward order before ref_node
            for line in formatted_lines:
                new_para = _make_json_para(line, first_pPr)
                if ref_node is not None:
                    ref_node.addprevious(new_para)
                else:
                    txbx.append(new_para)

            # Refresh paragraph list because DOM changed
            paragraphs = list(txbx.iter(qn("p")))
            if ref_node is not None:
                i = paragraphs.index(ref_node)
            else:
                i = len(paragraphs)
            continue

        i += 1


def normalize_x_language_name(root):
    """Normalize x-language-name header values to lowercase without quotes.
    Example: x-language-name: "EN" -> x-language-name: en
    """
    skip_styles = {
        "Heading1", "Heading2", "Heading3", "Heading4",
        "MID-PageOverview-TitleBig", "MID-PageOverView-TitleSmall", "MID-PageOverview-Small",
    }
    value_pattern = re.compile(r'^(: \s*)"?([A-Za-z]{2})"?$')
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

    for para in root.iter(qn("p")):
        pPr = para.find(qn("pPr"))
        pStyle = None
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None:
                pStyle = ps.get(qn("val"))
        if pStyle in skip_styles:
            continue

        para_text = get_element_text(para)
        if "x-language-name" not in para_text:
            continue

        for t in para.iter(qn("t")):
            if t.text is None:
                continue
            m = value_pattern.match(t.text)
            if m:
                new_text = m.group(1) + m.group(2).lower() + t.text[m.end():]
                if new_text != t.text:
                    t.text = new_text
                    if new_text.startswith(" ") or new_text.endswith(" "):
                        t.set(XML_SPACE, "preserve")


def highlight_variables_in_document(root):
    """Highlight variable placeholders in red across the entire document body."""
    skip_styles = {
        "Heading1", "Heading2", "Heading3", "Heading4",
        "MID-PageOverview-TitleBig", "MID-PageOverView-TitleSmall", "MID-PageOverview-Small",
    }
    for para in root.iter(qn("p")):
        pPr = para.find(qn("pPr"))
        pStyle = None
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None:
                pStyle = ps.get(qn("val"))
        if pStyle in skip_styles:
            continue
        runs = list(para.iter(qn("r")))
        _highlight_variables_in_para(runs)


def normalize_url_placeholders(root):
    old_base = "https://prd-gopaperless.mobile-id.vn/workflow/api"
    new_base = "https://{domain}/{contextPath}"
    for txbx in root.iter(qn("txbxContent")):
        text_nodes = [t for t in txbx.iter(qn("t"))]
        for idx, text_node in enumerate(text_nodes):
            if text_node.text is None or "prd-gopaperless.mobile-id.vn" not in text_node.text:
                continue
            previous_node = text_nodes[idx - 1] if idx > 0 else None
            if previous_node is None or previous_node.text is None:
                continue
            combined = previous_node.text + text_node.text
            if old_base not in combined:
                continue
            normalized = combined.replace(old_base, new_base)
            previous_node.text = ""
            text_node.text = normalized
            if text_node.text.startswith(" ") or text_node.text.endswith(" "):
                text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")



def normalize_txbx_content(root, method_rules):
    """Format text inside inline text-box frames (Sample Request / Response)."""
    # 1) Format text inside txbxContent
    for txbx in root.iter(qn("txbxContent")):
        for para in txbx.iter(qn("p")):
            runs = list(para.iter(qn("r")))
            _format_txbx_para(runs, method_rules)
            _highlight_variables_in_para(runs)
        _prettify_json_in_txbx(txbx)

    # 2) Format fallback text in body paragraphs containing inline shapes
    body = root.find(".//" + qn("body"))
    if body is not None:
        for para in body.iter(qn("p")):
            has_shape = para.find(".//" + qn("drawing")) is not None or para.find(".//" + qn("pict")) is not None
            if not has_shape:
                continue
            # Only process direct child runs (fallback text), not runs nested inside shapes/txbxContent
            direct_runs = [r for r in para if r.tag == qn("r")]
            if not direct_runs:
                continue
            para_text = "".join(t.text or "" for t in para.iter(qn("t")))
            if not para_text.strip():
                continue
            _format_txbx_para(direct_runs, method_rules)
            _highlight_variables_in_para(direct_runs)


def ensure_connection_info_section(root, config):
    body = root.find(".//" + qn("body"))
    if body is None:
        return
    children = list(body)
    sect_pr = body.find(qn("sectPr"))
    existing_idx = None
    for idx, child in enumerate(children):
        if child.tag == qn("p") and get_element_text(child).strip() == "System Connection Information":
            existing_idx = idx
            break
    if existing_idx is not None:
        remove_idx = existing_idx + 1
        while remove_idx < len(children):
            candidate = children[remove_idx]
            if candidate.tag == qn("p"):
                pPr = candidate.find(qn("pPr"))
                pStyle = pPr.find(qn("pStyle")) if pPr is not None else None
                if pStyle is not None and pStyle.get(qn("val")) in {"Heading1", "Heading2"}:
                    break
            body.remove(candidate)
            children = list(body)
        body.remove(children[existing_idx])
        children = list(body)
    insert_idx = None
    for idx, child in enumerate(children):
        if child.tag == qn("p") and get_element_text(child).strip() == "Abbreviation":
            insert_idx = idx
            break
    if insert_idx is None:
        insert_idx = children.index(sect_pr) if sect_pr is not None and sect_pr in children else len(children)
    else:
        insert_idx += 1
        while insert_idx < len(children):
            candidate = children[insert_idx]
            if candidate is sect_pr:
                break
            if candidate.tag == qn("p"):
                pPr = candidate.find(qn("pPr"))
                pStyle = pPr.find(qn("pStyle")) if pPr is not None else None
                if pStyle is not None and pStyle.get(qn("val")) in {"Heading1", "Heading2"}:
                    break
            insert_idx += 1
    section = [
        make_paragraph("System Connection Information", "Heading2", config),
        make_paragraph("The following table lists connection parameters used to build API endpoint URLs.", "Normal", config),
        make_connection_info_table(),
    ]
    for offset, element in enumerate(section):
        body.insert(insert_idx + offset, element)



def apply_heading_hierarchy_rules(root, heading_hierarchy_rules):
    if not heading_hierarchy_rules:
        return
    for para in root.iter(qn("p")):
        para_text = get_element_text(para).strip()
        target_style = heading_hierarchy_rules.get(para_text)
        if not target_style:
            continue
        pPr = para.find(qn("pPr"))
        if pPr is None:
            pPr = etree.Element(qn("pPr"))
            para.insert(0, pPr)
        pStyle = pPr.find(qn("pStyle"))
        if pStyle is None:
            pStyle = etree.Element(qn("pStyle"))
            pPr.insert(0, pStyle)
        pStyle.set(qn("val"), target_style)


def fix_heading_numbering_layout(numbering_path):
    tree = etree.parse(numbering_path)
    root = tree.getroot()
    # All heading levels use the same left/hanging indent (0.25" = 360 twips)
    # so that heading text starts at the same position regardless of level.
    heading_levels = {
        "0": {"left": "360", "hanging": "360"},
        "1": {"left": "360", "hanging": "360"},
        "2": {"left": "360", "hanging": "360"},
        "3": {"left": "360", "hanging": "360"},
    }
    for abstract_num in root.iter(qn("abstractNum")):
        for lvl in abstract_num.findall(qn("lvl")):
            ilvl = lvl.get(qn("ilvl"))
            p_style = lvl.find(qn("pStyle"))
            if ilvl not in heading_levels or p_style is None:
                continue
            if p_style.get(qn("val")) not in {"Heading1", "Heading2", "Heading3", "Heading4"}:
                continue
            pPr = get_or_create(lvl, "pPr")
            ind = get_or_create(pPr, "ind")
            ind.set(qn("left"), heading_levels[ilvl]["left"])
            ind.set(qn("hanging"), heading_levels[ilvl]["hanging"])
            rPr = get_or_create(lvl, "rPr")
            set_fonts(rPr, "Verdana")
            set_bold(rPr, True)
            set_color(rPr, "0050A8")
            heading_size = {"0": 32, "1": 28, "2": 24, "3": 22}[ilvl]
            set_size(rPr, heading_size)
    tree.write(numbering_path, xml_declaration=True, encoding="utf-8", standalone=True)
    print("[OK] Fixed heading numbering layout in numbering.xml")


def fix_document(document_path, config, table_border_rules, api_body_rules, table_property_rules, method_rules, heading_hierarchy_rules):
    tree = etree.parse(document_path)
    root = tree.getroot()

    # Keywords that should be uniform sub-headings
    SUBHEADING_KEYWORDS = [
        "Sample Request",
        "Sample Response",
        "Request Attributes",
        "Response Attributes",
    ]

    ensure_connection_info_section(root, config)
    apply_heading_hierarchy_rules(root, heading_hierarchy_rules)

    # 1) Normalize sub-heading paragraphs and "Attributes description"
    for para in root.iter(qn("p")):
        pPr = para.find(qn("pPr"))
        pStyle = None
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None:
                pStyle = ps.get(qn("val"))

        texts = []
        for child in para:
            if child.tag == qn("r"):
                for subchild in child:
                    if subchild.tag == qn("t") and subchild.text:
                        texts.append(subchild.text)
        para_text = "".join(texts).strip()

        matched = any(para_text == kw for kw in SUBHEADING_KEYWORDS)
        is_attrs = para_text == "Attributes description"

        if matched or is_attrs:
            for child in para:
                if child.tag == qn("r"):
                    rebuild_rPr(child, "Verdana", 22, True, None)
            if pPr is not None:
                p_rPr = pPr.find(qn("rPr"))
                if p_rPr is not None:
                    pPr.remove(p_rPr)

        if pStyle in (None, "Normal", "ListParagraph"):
            cfg_key = "ListParagraph" if pStyle == "ListParagraph" else "Normal"
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                if len(para):
                    para.insert(0, pPr)
                else:
                    para.append(pPr)
            apply_paragraph_style_layout(pPr, config[cfg_key])
            for child in para:
                if child.tag == qn("r"):
                    rPr = child.find(qn("rPr"))
                    if rPr is None:
                        continue
                    allowed_tags = {qn("rFonts"), qn("sz"), qn("szCs"), qn("lang"), qn("noProof")}
                    actual_tags = {sub.tag for sub in rPr}
                    if actual_tags.issubset(allowed_tags):
                        child.remove(rPr)

    # 2) Clean up Heading1-4 runs and paragraph-level rPr: remove redundant formatting
    #    so headings inherit font/size/bold/color cleanly from the style definition.
    heading_ids = {"Heading1", "Heading2", "Heading3", "Heading4"}
    for para in root.iter(qn("p")):
        pPr = para.find(qn("pPr"))
        pStyle = None
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None:
                pStyle = ps.get(qn("val"))
        if pStyle in heading_ids:
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                if len(para):
                    para.insert(0, pPr)
                else:
                    para.append(pPr)
            apply_paragraph_style_layout(pPr, config[pStyle])
            # Remove paragraph-level rPr (e.g. stray color overrides) so style wins
            p_rPr = pPr.find(qn("rPr"))
            if p_rPr is not None:
                pPr.remove(p_rPr)
            # Clean up run-level rPr
            for child in para:
                if child.tag == qn("r"):
                    rPr = child.find(qn("rPr"))
                    if rPr is not None:
                        lang = rPr.find(qn("lang"))
                        if lang is not None:
                            new_rPr = etree.Element(qn("rPr"))
                            new_rPr.append(deepcopy(lang))
                            child.remove(rPr)
                            if len(child):
                                child.insert(0, new_rPr)
                            else:
                                child.append(new_rPr)
                        else:
                            child.remove(rPr)

    # 2b) Clean up cover-page styles so they inherit alignment from the template
    #     and remove manual tabs used to fake centering.
    cover_styles = {"MID-PageOverview-TitleBig", "MID-PageOverView-TitleSmall", "MID-PageOverview-Small"}
    for para in root.iter(qn("p")):
        pPr = para.find(qn("pPr"))
        pStyle = None
        if pPr is not None:
            ps = pPr.find(qn("pStyle"))
            if ps is not None:
                pStyle = ps.get(qn("val"))
        if pStyle in cover_styles:
            para_text = get_element_text(para)
            if pPr is None:
                pPr = etree.Element(qn("pPr"))
                if len(para):
                    para.insert(0, pPr)
                else:
                    para.append(pPr)
            apply_paragraph_style_layout(pPr, config[pStyle])
            # Override alignment for specific cover lines
            if "Ho Chi Minh" in para_text:
                set_alignment(pPr, "center")
            elif "Version" in para_text:
                set_alignment(pPr, "right")
            # Remove paragraph-level rPr so style wins
            p_rPr = pPr.find(qn("rPr"))
            if p_rPr is not None:
                pPr.remove(p_rPr)
            # Clean up run-level rPr and remove manual <w:tab/> elements
            for child in list(para):
                if child.tag == qn("r"):
                    # Remove manual tabs that were used to fake center alignment
                    for tab_el in child.findall(qn("tab")):
                        child.remove(tab_el)
                    # Update year numbers on the cover page
                    for t_el in child.findall(qn("t")):
                        if t_el.text:
                            t_el.text = t_el.text.replace("2025", "2026")
                            t_el.text = t_el.text.replace("250609", "260505")
                    rPr = child.find(qn("rPr"))
                    if rPr is not None:
                        lang = rPr.find(qn("lang"))
                        if lang is not None:
                            new_rPr = etree.Element(qn("rPr"))
                            new_rPr.append(deepcopy(lang))
                            child.remove(rPr)
                            if len(child):
                                child.insert(0, new_rPr)
                            else:
                                child.append(new_rPr)
                        else:
                            child.remove(rPr)

    # 3) Bold, center-align and center-vertically all table header rows (rows with <w:tblHeader/>)
    for tr in root.iter(qn("tr")):
        trPr = tr.find(qn("trPr"))
        if trPr is None:
            continue
        if trPr.find(qn("tblHeader")) is None:
            continue
        for tc in tr.iter(qn("tc")):
            tcPr = get_or_create(tc, "tcPr")
            set_cell_valign(tcPr, "center")
            for para in tc.iter(qn("p")):
                pPr = para.find(qn("pPr"))
                if pPr is None:
                    pPr = etree.Element(qn("pPr"))
                    para.insert(0, pPr)
                set_alignment(pPr, "center")
                reset_table_cell_paragraph_spacing(pPr)
                reset_table_cell_paragraph_indent(pPr)
                for run in para.iter(qn("r")):
                    rPr = run.find(qn("rPr"))
                    if rPr is None:
                        rPr = etree.Element(qn("rPr"))
                        run.insert(0, rPr)
                    elif run.index(rPr) != 0:
                        run.remove(rPr)
                        run.insert(0, rPr)
                    set_fonts(rPr, "Verdana")
                    set_size(rPr, 24)
                    set_bold(rPr, True)
                    set_color(rPr, "FFFFFF")

    # 4) Unify table header background colors (0066CC -> 0070C0)
    for shd in root.iter(qn("shd")):
        fill = shd.get(qn("fill"))
        if fill and fill.upper() == "0066CC":
            shd.set(qn("fill"), "0070C0")

    # 4b) Center-align Version and Date columns in the History table
    for tbl in root.iter(qn("tbl")):
        header_row = None
        for tr in tbl.iter(qn("tr")):
            trPr = tr.find(qn("trPr"))
            if trPr is not None and trPr.find(qn("tblHeader")) is not None:
                header_row = tr
                break
            if header_row is None:
                header_row = tr
                break

        if header_row is None:
            continue

        cells = header_row.findall(qn("tc"))
        if len(cells) < 2:
            continue

        cell_texts = []
        for tc in cells:
            t = "".join(t.text or "" for t in tc.iter(qn("t"))).strip()
            cell_texts.append(t)

        if cell_texts[0].lower() != "version" or cell_texts[1].lower() != "date":
            continue

        tblPr = get_or_create(tbl, "tblPr")
        tbl_ind = tblPr.find(qn("tblInd"))
        if tbl_ind is not None:
            tbl_ind.set(qn("w"), "0")
            tbl_ind.set(qn("type"), "dxa")

        centered_history_columns = {0, 1, len(cells) - 1}
        for tr in tbl.iter(qn("tr")):
            row_cells = tr.findall(qn("tc"))
            is_header_row = tr is header_row
            for col_idx, tc in enumerate(row_cells):
                tcPr = get_or_create(tc, "tcPr")
                set_cell_valign(tcPr, "center")
                set_cell_margins(tcPr)
                for para in tc.iter(qn("p")):
                    pPr = para.find(qn("pPr"))
                    if pPr is None:
                        pPr = etree.Element(qn("pPr"))
                        para.insert(0, pPr)
                    if is_header_row or col_idx in centered_history_columns:
                        set_alignment(pPr, "center")
                    reset_table_cell_paragraph_spacing(pPr)
                    reset_table_cell_paragraph_indent(pPr)

    # 5) Apply table-specific border rules from the formatting template
    apply_table_border_rules(root, table_border_rules)

    # 5b) Rename generic 'Header Attributes' to Request/Response variants
    rename_header_attribute_tables(root)

    # 5c) Auto-create Path Attributes tables for APIs with URL path variables
    create_path_attribute_tables(root, config, api_body_rules)

    # 6) Enforce attribute table start-header alignment and font color from the template
    normalize_api_table_start_headers(root, table_property_rules)

    # 6b) Center-align all tables and unify table indent
    apply_table_alignment(root, table_property_rules)

    # 6c) Split label+frame paragraphs so labels stay left-aligned
    split_label_from_frame_paragraphs(root)

    # 6d) Center-align inline text-box frames (Sample Request/Response) and unify their width
    normalize_inline_shapes(root)

    # 7) Enforce spacing and blank-line rules inside each API body section
    normalize_api_body_spacing(root, config, api_body_rules)

    # 8) Ensure exactly one blank paragraph before each Heading 3 (except the first one after a Heading 2)
    normalize_heading3_spacing(root, config)

    # 9) Normalize request URLs and format inline text-box frame content
    normalize_url_placeholders(root)
    normalize_txbx_content(root, method_rules)

    # 10) Highlight variable placeholders like {domain}, {token} in red across the document
    highlight_variables_in_document(root)

    # 11) Normalize x-language-name header values to lowercase (en, vi, etc.)
    normalize_x_language_name(root)

    tree.write(document_path, xml_declaration=True, encoding="utf-8", standalone=True)
    print(f"[OK] Fixed document.xml")


def pack_docx(source_dir, output_path):
    """Pack unpacked directory into a .docx file."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(source_dir):
            for f in files:
                full = os.path.join(root, f)
                arc = os.path.relpath(full, source_dir)
                zf.write(full, arc)
    print(f"[OK] Packed -> {output_path}")


def main():
    unpacked = "unpacked_docx"
    if not os.path.isdir(unpacked):
        print(f"Directory '{unpacked}' not found.")
        sys.exit(1)

    template_md = os.path.join("scripts", "GoPaperless_DOCX_Format_Template.md")
    if not os.path.isfile(template_md):
        print(f"Template not found: {template_md}")
        sys.exit(1)

    config = parse_template(template_md)
    table_border_rules = parse_table_border_rules(template_md)
    table_property_rules = parse_table_property_rules(template_md)
    api_body_rules = parse_api_body_rules(template_md)
    method_rules = parse_frame_method_rules(template_md)
    heading_hierarchy_rules = parse_heading_hierarchy_rules(template_md)
    print(f"[OK] Loaded {len(config)} style definitions from {template_md}")
    print(f"[OK] Loaded {len(table_border_rules)} table border rules from {template_md}")
    print(f"[OK] Loaded {len(table_property_rules)} table property rules from {template_md}")
    print(f"[OK] Loaded {len(api_body_rules)} API body rules from {template_md}")
    print(f"[OK] Loaded {len(method_rules)} frame method rules from {template_md}")
    print(f"[OK] Loaded {len(heading_hierarchy_rules)} heading hierarchy rules from {template_md}")

    styles_xml = os.path.join(unpacked, "word", "styles.xml")
    document_xml = os.path.join(unpacked, "word", "document.xml")
    numbering_xml = os.path.join(unpacked, "word", "numbering.xml")

    fix_styles(styles_xml, config)
    fix_heading_numbering_layout(numbering_xml)
    fix_document(document_xml, config, table_border_rules, api_body_rules, table_property_rules, method_rules, heading_hierarchy_rules)

    output_dir = "formatted"
    os.makedirs(output_dir, exist_ok=True)
    output = os.path.join(output_dir, "GoPaperless Workflow RESTful API Specification V2-Formatted.docx")
    # If the file is locked (e.g., open in Word), write to a temporary name instead
    try:
        pack_docx(unpacked, output)
    except PermissionError:
        alt_output = os.path.join(output_dir, "GoPaperless Workflow RESTful API Specification V2-Formatted-New.docx")
        pack_docx(unpacked, alt_output)
        print(f"\n[Warning] Could not overwrite locked file: {output}")
        print(f"Done! New file: {alt_output}")
        return
    print(f"\nDone! New file: {output}")


if __name__ == "__main__":
    main()
