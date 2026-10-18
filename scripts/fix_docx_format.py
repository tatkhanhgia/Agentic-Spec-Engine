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


def parse_column_widths(table_property_rules, key, expected_count):
    """Read a required comma-separated dxa column-width list from the template."""
    value = table_property_rules.get(key, "")
    widths = [int(part) for part in re.findall(r"\d+", value)]
    if "dxa" not in value.lower() or len(widths) != expected_count or any(w <= 0 for w in widths):
        raise ValueError(
            "%s must define exactly %d positive dxa widths; got %r"
            % (key, expected_count, value))
    return widths


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


def parse_frame_url_rules(md_path):
    """Parse request-line URL rules from section 2.8 property table."""
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.read().splitlines()

    rules = {}
    in_section = False
    in_table = False
    for line in lines:
        if line.startswith('### 2.8. Request'):
            in_section = True
            continue
        if not in_section:
            continue
        if line.startswith('#### 2.8.1'):
            break
        if '| Property |' in line and 'Standard Value' in line:
            in_table = True
            continue
        if not in_table:
            continue
        if line.strip() == '' or not line.startswith('|'):
            break
        cells = [c.strip() for c in line.split('|')]
        cells = [c for c in cells if c]
        if len(cells) < 2 or set(cells[0]) <= set('-'):
            continue
        key = cells[0].replace('**', '').strip()
        val = cells[1].strip().strip('`')
        rules[key] = val
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


ATTRIBUTE_COLUMN_HEADER = ("No", "Name", "Type", "Presence", "Description")
ATTRIBUTE_CENTER_COLUMNS = (0, 3)  # No and Presence (template 2.4)


def make_table_cell(text, width, width_type="dxa", header=False, center=None):
    """Create a cell with an explicit width; callers cannot inherit a stale default."""
    if width_type not in {"dxa", "pct"} or int(width) <= 0:
        raise ValueError("invalid table-cell width: %r %r" % (width, width_type))
    tc = etree.Element(qn("tc"))
    tcPr = etree.SubElement(tc, qn("tcPr"))
    tcW = etree.SubElement(tcPr, qn("tcW"))
    tcW.set(qn("w"), str(width))
    tcW.set(qn("type"), width_type)
    set_cell_valign(tcPr, "center")
    set_cell_margins(tcPr)
    para = etree.SubElement(tc, qn("p"))
    pPr = etree.SubElement(para, qn("pPr"))
    reset_table_cell_paragraph_spacing(pPr)
    reset_table_cell_paragraph_indent(pPr)
    if center is None:
        center = header
    set_alignment(pPr, "center" if center else "left")
    run = make_run(text, size=24 if header else 22, bold=header, color="FFFFFF" if header else "000000")
    para.append(run)
    if header:
        shd = etree.SubElement(tcPr, qn("shd"))
        shd.set(qn("val"), "clear")
        shd.set(qn("color"), "auto")
        shd.set(qn("fill"), "0070C0")
    return tc


def make_connection_info_table(frame_url_rules, column_widths):
    frame_url_rules = frame_url_rules or {}
    if len(column_widths) != 4:
        raise ValueError("System Connection Information requires four column widths")
    domain = frame_url_rules.get('Current domain', 'prd-gopaperless.mobile-id.vn')
    context_path = frame_url_rules.get('Current contextPath', '/workflow/api')
    base_url_value = frame_url_rules.get('Connection table baseURL value', 'https://{domain}{contextPath}')
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
    for width in column_widths:
        gridCol = etree.SubElement(tblGrid, qn("gridCol"))
        gridCol.set(qn("w"), str(width))
    rows = [
        (True, ["No", "Name", "Value", "Description"]),
        (False, ["1", "domain", domain, "Current production domain"]),
        (False, ["2", "contextPath", context_path, "API context path"]),
        (False, ["3", "baseURL", base_url_value, "Base URL for all API endpoints"]),
    ]
    for is_header, values in rows:
        tr = etree.SubElement(tbl, qn("tr"))
        if is_header:
            trPr = etree.SubElement(tr, qn("trPr"))
            etree.SubElement(trPr, qn("tblHeader"))
        for value, width in zip(values, column_widths):
            tr.append(make_table_cell(value, width, header=is_header))
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


def ensure_tbl_header(tr):
    trPr = get_or_create(tr, "trPr")
    if trPr.find(qn("tblHeader")) is None:
        etree.SubElement(trPr, qn("tblHeader"))


def normalize_attribute_table_borders_and_headers(root, table_property_rules):
    """Apply black borders and tblHeader flags to every attribute table.

    Generated Path Attributes tables and sections cloned from a generated donor
    inherit D9D9D9 borders and omit tblHeader on the column-header row; legacy
    endpoint tables use 000000 / sz=4 with tblHeader on rows 0 and 1.
    """
    border_size = str(parse_int_property(table_property_rules, "Attribute table border size") or 4)
    border_color = parse_hex_color(table_property_rules.get("Attribute table border color")) or "000000"
    tables_changed = 0
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        if len(rows) < 2:
            continue
        header = tuple(
            get_element_text(tc).strip()
            for tc in rows[1].findall(qn("tc"))[:len(ATTRIBUTE_COLUMN_HEADER)]
        )
        if header != ATTRIBUTE_COLUMN_HEADER:
            continue
        tblPr = get_or_create(tbl, "tblPr")
        tblBorders = get_or_create(tblPr, "tblBorders")
        for tag in ("top", "left", "bottom", "right", "insideH", "insideV"):
            set_border(tblBorders, tag, size=border_size, color=border_color)
        ensure_tbl_header(rows[0])
        ensure_tbl_header(rows[1])
        tables_changed += 1
    print("[OK] Normalized attribute table borders/headers: %d tables" % tables_changed)


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
        if len(rows) < 2:
            continue
        column_header = tuple(
            get_element_text(tc).strip()
            for tc in rows[1].findall(qn("tc"))[:len(ATTRIBUTE_COLUMN_HEADER)]
        )
        if column_header != ATTRIBUTE_COLUMN_HEADER:
            continue
        first_row = rows[0]
        cells = first_row.findall(qn("tc"))
        if not cells:
            continue
        title = get_element_text(cells[0]).strip()
        if title not in start_headers:
            pass
        elif header_row_height is not None and header_row_height_rule:
            trPr = get_or_create(first_row, "trPr")
            trHeight = get_or_create(trPr, "trHeight")
            trHeight.set(qn("val"), str(header_row_height))
            trHeight.set(qn("hRule"), header_row_height_rule)
        if title in start_headers and header_width is not None and header_width_type:
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


def normalize_table_body_fonts(root):
    """Ensure all table body rows use Verdana 11 pt non-bold black text."""
    start_headers = {
        "Header Attributes",
        "Header Request Attributes",
        "Header Response Attributes",
        "Path Attributes",
        "Request Attributes",
        "Response Attributes",
    }
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        for tr in rows:
            trPr = tr.find(qn("trPr"))
            is_header_row = trPr is not None and trPr.find(qn("tblHeader")) is not None
            if is_header_row:
                continue

            cells = tr.findall(qn("tc"))
            if cells:
                first_cell_text = get_element_text(cells[0]).strip()
                if first_cell_text in start_headers:
                    continue

            for tc in tr.iter(qn("tc")):
                tcPr = get_or_create(tc, "tcPr")
                set_cell_valign(tcPr, "center")
                for para in tc.iter(qn("p")):
                    pPr = para.find(qn("pPr"))
                    if pPr is None:
                        pPr = etree.Element(qn("pPr"))
                        para.insert(0, pPr)
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
                        set_size(rPr, 22)
                        set_bold(rPr, False)
                        set_color(rPr, "000000")
                        set_highlight(rPr, None)
                        set_shd(rPr, None)
    print("[OK] Normalized table body fonts")


def normalize_attribute_table_column_alignment(root):
    """Centre the No and Presence columns of every attribute table (template 2.4).

    Step 1 applies Normal-style left alignment to all body paragraphs, including
    table cells, so this pass must run after that normalization.
    """
    cells_changed = 0
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        if len(rows) < 3:
            continue
        header = tuple(
            get_element_text(tc).strip()
            for tc in rows[1].findall(qn("tc"))[:len(ATTRIBUTE_COLUMN_HEADER)]
        )
        if header != ATTRIBUTE_COLUMN_HEADER:
            continue
        for tr in rows[2:]:
            cells = tr.findall(qn("tc"))
            if len(cells) < 5:
                continue
            for col in ATTRIBUTE_CENTER_COLUMNS:
                tc = cells[col]
                for para in tc.findall(qn("p")):
                    pPr = para.find(qn("pPr"))
                    if pPr is None:
                        pPr = etree.Element(qn("pPr"))
                        para.insert(0, pPr)
                    jc = pPr.find(qn("jc"))
                    if jc is None:
                        jc = etree.SubElement(pPr, qn("jc"))
                    if jc.get(qn("val")) != "center":
                        jc.set(qn("val"), "center")
                        cells_changed += 1
    print("[OK] Normalized attribute table column alignment: %d cells" % cells_changed)


def normalize_attribute_table_widths(root):
    """Replace invalid percentage widths in regular attribute-table rows.

    Existing legacy tables intentionally retain their current dxa proportions.
    Only non-dxa widths are synchronized to the table's own five-column grid.
    """
    tables_changed = 0
    cells_changed = 0
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        if len(rows) < 2:
            continue
        header = tuple(
            get_element_text(tc).strip()
            for tc in rows[1].findall(qn("tc"))[:len(ATTRIBUTE_COLUMN_HEADER)]
        )
        if header != ATTRIBUTE_COLUMN_HEADER:
            continue
        title = get_element_text(rows[0]).strip()
        exact_widths = title == "Path Attributes"
        grid_widths = [
            grid_col.get(qn("w"))
            for grid_col in tbl.findall(qn("tblGrid") + "/" + qn("gridCol"))
        ]
        if len(grid_widths) != len(ATTRIBUTE_COLUMN_HEADER) or any(
                width is None or not width.isdigit() for width in grid_widths):
            raise ValueError("attribute table has an invalid five-column tblGrid")

        table_changed = False
        for tr in rows[1:]:
            cells = tr.findall(qn("tc"))
            if len(cells) != len(grid_widths):
                raise ValueError("attribute table row does not match its five-column tblGrid")
            for tc, width in zip(cells, grid_widths):
                tcPr = get_or_create(tc, "tcPr")
                tcW = get_or_create(tcPr, "tcW")
                if tcW.get(qn("type")) == "dxa" and (
                        not exact_widths or tcW.get(qn("w")) == width):
                    continue
                tcW.set(qn("w"), width)
                tcW.set(qn("type"), "dxa")
                cells_changed += 1
                table_changed = True
        if table_changed:
            tables_changed += 1
    print("[OK] Normalized attribute table widths: %d tables, %d cells" % (
        tables_changed, cells_changed))


def normalize_connection_table_widths(root, table_property_rules):
    """Enforce the template grid on the generated connection-information table."""
    expected = parse_column_widths(
        table_property_rules, "Generated connection table column widths", 4)
    changed = False
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        if not rows:
            continue
        header = tuple(
            get_element_text(tc).strip() for tc in rows[0].findall(qn("tc")))
        if header != ("No", "Name", "Value", "Description"):
            continue

        tbl_grid = tbl.find(qn("tblGrid"))
        if tbl_grid is None:
            tbl_grid = etree.Element(qn("tblGrid"))
            tbl.insert(1, tbl_grid)
            changed = True
        current = [col.get(qn("w")) for col in tbl_grid.findall(qn("gridCol"))]
        if current != [str(width) for width in expected]:
            for col in tbl_grid.findall(qn("gridCol")):
                tbl_grid.remove(col)
            for width in expected:
                grid_col = etree.SubElement(tbl_grid, qn("gridCol"))
                grid_col.set(qn("w"), str(width))
            changed = True

        for tr in rows:
            cells = tr.findall(qn("tc"))
            if len(cells) != len(expected):
                raise ValueError(
                    "System Connection Information row must have four cells")
            for tc, width in zip(cells, expected):
                tcPr = get_or_create(tc, "tcPr")
                tcW = get_or_create(tcPr, "tcW")
                if (tcW.get(qn("w")), tcW.get(qn("type"))) == (
                        str(width), "dxa"):
                    continue
                tcW.set(qn("w"), str(width))
                tcW.set(qn("type"), "dxa")
                changed = True
    print("[OK] Normalized connection table widths%s" % (
        "" if changed else " (already clean)"))


def table_width_errors(root):
    """Return width invariant violations that must block DOCX packaging."""
    errors = []
    for tbl in root.iter(qn("tbl")):
        rows = tbl.findall(qn("tr"))
        if not rows:
            continue
        first_row = tuple(
            get_element_text(tc).strip() for tc in rows[0].findall(qn("tc")))
        title = get_element_text(rows[0]).strip()
        header = tuple(
            get_element_text(tc).strip()
            for tc in rows[1].findall(qn("tc"))[:len(ATTRIBUTE_COLUMN_HEADER)]
        ) if len(rows) > 1 else ()
        is_attribute = header == ATTRIBUTE_COLUMN_HEADER
        is_connection = first_row == ("No", "Name", "Value", "Description")
        if not is_attribute and not is_connection:
            continue

        grid_widths = [
            grid_col.get(qn("w"))
            for grid_col in tbl.findall(qn("tblGrid") + "/" + qn("gridCol"))
        ]
        expected_count = 5 if is_attribute else 4
        if len(grid_widths) != expected_count or any(
                width is None or not width.isdigit() for width in grid_widths):
            errors.append("%s: invalid %d-column tblGrid" % (
                title or "System Connection Information", expected_count))
            continue

        checked_rows = rows[1:] if is_attribute else rows
        exact_widths = is_connection or title == "Path Attributes"
        for row_no, tr in enumerate(checked_rows, 1):
            cells = tr.findall(qn("tc"))
            if len(cells) != expected_count:
                errors.append("%s row %d: %d cells for %d grid columns" % (
                    title, row_no, len(cells), expected_count))
                continue
            for col_no, (tc, grid_width) in enumerate(
                    zip(cells, grid_widths), 1):
                tcW = tc.find(qn("tcPr") + "/" + qn("tcW"))
                if tcW is None:
                    errors.append("%s row %d column %d: missing tcW" % (
                        title, row_no, col_no))
                    continue
                width = tcW.get(qn("w"))
                width_type = tcW.get(qn("type"))
                if width_type != "dxa":
                    errors.append(
                        "%s row %d column %d: %s %s; regular cells require dxa"
                        % (title, row_no, col_no, width, width_type))
                elif exact_widths and width != grid_width:
                    errors.append(
                        "%s row %d column %d: %s dxa; grid requires %s dxa"
                        % (title, row_no, col_no, width, grid_width))
    return errors


def validate_table_widths(root):
    """Raise before packaging when a table can reproduce the width defect."""
    errors = table_width_errors(root)
    if errors:
        shown = "\n".join("  - " + item for item in errors[:20])
        suffix = "\n  ... and %d more" % (len(errors) - 20) \
            if len(errors) > 20 else ""
        raise ValueError(
            "DOCX table-width validation failed (%d issue(s)):\n%s%s"
            % (len(errors), shown, suffix))


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


def make_path_attributes_table(path_vars, column_widths, table_property_rules=None):
    """Create a Path Attributes table for the given path variables."""
    if len(column_widths) != len(ATTRIBUTE_COLUMN_HEADER):
        raise ValueError("Path Attributes requires five column widths")
    rules = table_property_rules or {}
    border_size = str(parse_int_property(rules, "Attribute table border size") or 4)
    border_color = parse_hex_color(rules.get("Attribute table border color")) or "000000"
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
        set_border(tblBorders, border, size=border_size, color=border_color)
    tblGrid = etree.SubElement(tbl, qn("tblGrid"))
    for width in column_widths:
        gridCol = etree.SubElement(tblGrid, qn("gridCol"))
        gridCol.set(qn("w"), str(width))

    # Row 1: merged header "Path Attributes"
    tr1 = etree.SubElement(tbl, qn("tr"))
    trPr1 = etree.SubElement(tr1, qn("trPr"))
    etree.SubElement(trPr1, qn("tblHeader"))
    trHeight1 = etree.SubElement(trPr1, qn("trHeight"))
    trHeight1.set(qn("val"), "420")
    trHeight1.set(qn("hRule"), "atLeast")
    tc1 = make_table_cell("Path Attributes", 5000, width_type="pct", header=True)
    tcPr1 = tc1.find(qn("tcPr"))
    grid_span1 = get_or_create(tcPr1, "gridSpan")
    grid_span1.set(qn("val"), "5")
    tr1.append(tc1)

    # Row 2: column headers
    tr2 = etree.SubElement(tbl, qn("tr"))
    trPr2 = etree.SubElement(tr2, qn("trPr"))
    etree.SubElement(trPr2, qn("tblHeader"))
    trHeight2 = etree.SubElement(trPr2, qn("trHeight"))
    trHeight2.set(qn("val"), "420")
    trHeight2.set(qn("hRule"), "atLeast")
    for heading, width in zip(ATTRIBUTE_COLUMN_HEADER, column_widths):
        tr2.append(make_table_cell(heading, width, header=True))

    # Data rows
    for no, var in enumerate(path_vars, start=1):
        tr = etree.SubElement(tbl, qn("tr"))
        trPr = etree.SubElement(tr, qn("trPr"))
        trHeight = etree.SubElement(trPr, qn("trHeight"))
        trHeight.set(qn("val"), "420")
        trHeight.set(qn("hRule"), "atLeast")
        for col_idx, (value, width) in enumerate(
                zip((str(no), var, "String", "M", ""), column_widths)):
            tr.append(make_table_cell(
                value, width, header=False, center=(col_idx in ATTRIBUTE_CENTER_COLUMNS)))

    return tbl


def create_path_attribute_tables(root, config, api_body_rules, table_property_rules):
    """Auto-create Path Attributes tables for API sections that have URL path variables
    but are missing a Path Attributes table."""
    column_widths = parse_column_widths(
        table_property_rules,
        "Generated attribute table column widths",
        len(ATTRIBUTE_COLUMN_HEADER),
    )
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
                    m = re.search(r'\{baseURL\}(/[^?\s]*)', text)
                    if m:
                        path_part = m.group(1)
                        vars_found = re.findall(r'\{([^}]+)\}', path_part)
                        vars_found = [v for v in vars_found if v not in ("domain", "contextPath", "baseURL")]
                        return vars_found
                    # Legacy fallback while migrating older frame content
                    m = re.search(r'(https?://\{[^}]+\}/\{[^}]+\})(/[^?\s]*)', text)
                    if m:
                        path_part = m.group(2)
                        vars_found = re.findall(r'\{([^}]+)\}', path_part)
                        vars_found = [v for v in vars_found if v not in ("domain", "contextPath", "baseURL")]
                        return vars_found
                    # Fallback: any URL-looking segment with variables
                    m2 = re.search(r'/(v\d+[^?\s]*)', text)
                    if m2:
                        path_part = m2.group(1)
                        vars_found = re.findall(r'\{([^}]+)\}', "/" + path_part)
                        vars_found = [v for v in vars_found if v not in ("domain", "contextPath", "baseURL")]
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
                        path_tbl = make_path_attributes_table(
                            path_vars, column_widths, table_property_rules)
                        blank = make_blank_paragraph_like(None, config, api_body_rules.get("Between attribute tables", {}))
                        body.insert(real_insert, blank)
                        body.insert(real_insert + 1, path_tbl)
                        children = list(body)  # refresh
                        idx = real_insert + 2
                        continue
        idx += 1


def normalize_attribute_table_content(root):
    """Fix semantic content of API attribute tables:
    - Move accesstoken/access_token from Request Attributes to Header Request Attributes as Authorization
    - Ensure Content-Type is present in Header Request Attributes
    - Remove empty Request Attributes tables
    - Renumber rows after modifications
    """
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

    def get_row_name(tr):
        cells = tr.findall(qn("tc"))
        if len(cells) >= 2:
            return get_element_text(cells[1]).strip().lower()
        return ""

    def set_cell_text(tc, text):
        """Replace text in a table cell, preserving one paragraph with one run.
        Removes any extra paragraphs to avoid leftover text from cloned rows."""
        paras = tc.findall(qn("p"))
        if not paras:
            para = etree.SubElement(tc, qn("p"))
        else:
            para = paras[0]
            # Remove extra paragraphs to prevent leftover text
            for extra in paras[1:]:
                tc.remove(extra)
        for r in para.findall(qn("r")):
            para.remove(r)
        run = make_run(text, font="Verdana", size=22, bold=False, color="000000")
        para.append(run)

    def renumber_table(tbl, start_idx=2):
        rows = tbl.findall(qn("tr"))
        for i, tr in enumerate(rows[start_idx:], start=1):
            cells = tr.findall(qn("tc"))
            if cells:
                set_cell_text(cells[0], str(i))

    def clone_data_row(template_tr):
        """Clone a data row and return it for modification."""
        new_tr = deepcopy(template_tr)
        for attr in list(new_tr.attrib.keys()):
            if "rsid" in attr or "paraId" in attr or "textId" in attr:
                del new_tr.attrib[attr]
        for para in new_tr.iter(qn("p")):
            for attr in list(para.attrib.keys()):
                if "rsid" in attr or "paraId" in attr or "textId" in attr:
                    del para.attrib[attr]
        return new_tr

    children = list(body)
    idx = 0
    moved_count = 0
    added_count = 0
    removed_tables = []

    while idx < len(children):
        child = children[idx]
        if is_heading_break(child):
            pPr = child.find(qn("pPr"))
            ps = pPr.find(qn("pStyle")) if pPr is not None else None
            style = ps.get(qn("val")) if ps is not None else ""
            if style in {"Heading3", "Heading4"}:
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

                header_tbl = None
                request_tbl = None
                for node in section_nodes:
                    if node.tag == qn("tbl"):
                        htext = get_table_header_text(node)
                        if htext in {"Header Attributes", "Header Request Attributes"}:
                            header_tbl = node
                        elif htext == "Request Attributes":
                            request_tbl = node

                # Always add Content-Type to Header Attributes if missing
                if header_tbl is not None:
                    header_rows = header_tbl.findall(qn("tr"))
                    has_ct = False
                    for tr in header_rows[2:]:
                        if get_row_name(tr) == "content-type":
                            has_ct = True
                            break
                    if not has_ct:
                        template_row = header_rows[2] if len(header_rows) > 2 else None
                        if template_row is None and request_tbl is not None:
                            req_rows = request_tbl.findall(qn("tr"))
                            if len(req_rows) > 2:
                                template_row = req_rows[2]
                        if template_row is not None:
                            ct_row = clone_data_row(template_row)
                            cells = ct_row.findall(qn("tc"))
                            if len(cells) >= 5:
                                set_cell_text(cells[1], "Content-Type")
                                set_cell_text(cells[2], "String")
                                set_cell_text(cells[3], "M")
                                set_cell_text(cells[4], "The MIME type of the request body. Usually application/json")
                            header_tbl.append(ct_row)
                            added_count += 1
                            # Renumber if there won't be a second renumber below
                            if request_tbl is None:
                                renumber_table(header_tbl, start_idx=2)

                # Process accesstoken movement only when both tables exist
                if header_tbl is not None and request_tbl is not None:
                    req_rows = request_tbl.findall(qn("tr"))
                    rows_to_remove = []
                    moved_auth = False

                    for r_idx, tr in enumerate(req_rows[2:], start=2):
                        name = get_row_name(tr)
                        if name in ("accesstoken", "access_token"):
                            rows_to_remove.append(r_idx)
                            moved_auth = True

                    for r_idx in reversed(rows_to_remove):
                        request_tbl.remove(req_rows[r_idx])

                    header_rows = header_tbl.findall(qn("tr"))
                    has_auth = False
                    for tr in header_rows[2:]:
                        if get_row_name(tr) == "authorization":
                            has_auth = True
                            break

                    if moved_auth and not has_auth:
                        template_row = header_rows[2] if len(header_rows) > 2 else None
                        if template_row is None and len(req_rows) > 2:
                            template_row = req_rows[2]
                        if template_row is not None:
                            auth_row = clone_data_row(template_row)
                            cells = auth_row.findall(qn("tc"))
                            if len(cells) >= 5:
                                set_cell_text(cells[1], "Authorization")
                                set_cell_text(cells[2], "String")
                                set_cell_text(cells[3], "M")
                                set_cell_text(cells[4], "The authentication token (accessToken) required to access the API. Must be provided in the format: Bearer {accessToken}")
                            header_tbl.append(auth_row)
                            has_auth = True
                            moved_count += 1

                    renumber_table(header_tbl, start_idx=2)

                    # Check if Request Attributes table is now empty (only header rows)
                    req_rows_after = request_tbl.findall(qn("tr"))
                    if len(req_rows_after) <= 2:
                        removed_tables.append(request_tbl)
                    else:
                        renumber_table(request_tbl, start_idx=2)

        idx += 1

    # Remove empty Request Attributes tables
    for tbl in removed_tables:
        try:
            body.remove(tbl)
        except ValueError:
            pass

    if moved_count or added_count or removed_tables:
        print(f"[OK] Normalized attribute table content ({moved_count} auth moved, {added_count} Content-Type added, {len(removed_tables)} empty tables removed)")
    else:
        print("[OK] Attribute table content already correct")


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


_WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_VML_NS = "urn:schemas-microsoft-com:vml"
_FRAME_METHOD_RE = re.compile(r"^(GET|POST|PUT|DELETE|PATCH)\b")


def _is_sample_frame_txbx(txbx):
    """A txbxContent is a sample frame when its first line is a request line,
    an `Http Code` status line, or starts a bare JSON body (template 2.8.10)."""
    paras = txbx.findall(qn("p"))
    if not paras:
        return False
    first = "".join(t.text or "" for t in paras[0].iter(qn("t"))).strip()
    return bool(_FRAME_METHOD_RE.match(first)) or first.startswith(("Http Code", "{", "["))


def normalize_frame_autofit(root):
    """Template 2.8.10: every sample frame textbox must auto-fit its text.

    A fixed-height frame clips JSON lines when the text is taller than the
    shape (silent data loss) and reserves dead space when it is shorter. Each
    frame is stored twice inside mc:AlternateContent and both representations
    must agree:
      - DrawingML: <wps:bodyPr> carries <a:spAutoFit/> (replacing any
        <a:noAutofit/> / <a:normAutofit/> in place, to keep schema order);
      - VML fallback: the <v:textbox> style contains mso-fit-shape-to-text:t.
    The stored height is only the last rendered size; Word recomputes it on
    open once these flags are present, so it is left untouched.
    """
    fixed_drawing = fixed_vml = 0

    for wtxbx in root.iter("{%s}txbx" % _WPS_NS):
        txbx = wtxbx.find(qn("txbxContent"))
        if txbx is None or not _is_sample_frame_txbx(txbx):
            continue
        shape = wtxbx.getparent()
        bodyPr = shape.find("{%s}bodyPr" % _WPS_NS)
        if bodyPr is None:
            bodyPr = etree.SubElement(shape, "{%s}bodyPr" % _WPS_NS)
        if bodyPr.find("{%s}spAutoFit" % _A_NS) is None:
            replaced = False
            for tag in ("noAutofit", "normAutofit"):
                old = bodyPr.find("{%s}%s" % (_A_NS, tag))
                if old is not None:
                    bodyPr.replace(old, etree.Element("{%s}spAutoFit" % _A_NS))
                    replaced = True
                    break
            if not replaced:
                bodyPr.append(etree.Element("{%s}spAutoFit" % _A_NS))
            fixed_drawing += 1

    for vtb in root.iter("{%s}textbox" % _VML_NS):
        txbx = vtb.find(qn("txbxContent"))
        if txbx is None or not _is_sample_frame_txbx(txbx):
            continue
        style = vtb.get("style") or ""
        if "mso-fit-shape-to-text" not in style:
            vtb.set("style", (style.rstrip(";") + ";" if style else "") + "mso-fit-shape-to-text:t")
            fixed_vml += 1

    if fixed_drawing or fixed_vml:
        print(f"[OK] Frame auto-fit applied: {fixed_drawing} DrawingML bodyPr, {fixed_vml} VML textboxes")
    return fixed_drawing + fixed_vml


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

        # Whole block did not parse: this is a fragment (a body split across
        # two frames). Skip the entire scanned block. Descending into it would
        # re-dump nested objects at indent 0, which flattens the template
        # 2.8.4 "2 spaces per nesting level" indentation of the outer levels.
        i = block_end + 1


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


def _set_paragraph_text(para, new_text):
    """Replace paragraph text while preserving the first text node."""
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
    first_t = None
    for t in para.iter(qn("t")):
        if first_t is None:
            first_t = t
        else:
            t.text = ""
    if first_t is None:
        return
    first_t.text = new_text
    if new_text.startswith(" ") or new_text.endswith(" "):
        first_t.set(XML_SPACE, "preserve")


def _to_baseurl_endpoint(url_part, frame_url_rules):
    """Normalize a request-line URL segment to {baseURL}/path form."""
    url_part = url_part.strip()
    placeholder_name = frame_url_rules.get('Request line placeholder', 'baseURL')
    placeholder = '{' + placeholder_name + '}'

    if url_part.startswith(placeholder):
        return url_part

    domain = frame_url_rules.get('Current domain', 'prd-gopaperless.mobile-id.vn')
    context_path = frame_url_rules.get('Current contextPath', '/workflow/api')
    hardcoded_base = f'https://{domain}{context_path}'

    prefixes = [
        hardcoded_base,
        'https://{domain}/{contextPath}',
        'https://{domain}{contextPath}',
        f'https://{{domain}}{context_path}',
        f'https://{{domain}}/{{contextPath}}',
    ]
    for prefix in prefixes:
        if url_part.startswith(prefix):
            path = url_part[len(prefix):]
            if path and not path.startswith('/'):
                path = '/' + path
            return placeholder + path

    if url_part.startswith('/'):
        return placeholder + url_part

    return url_part


def normalize_url_placeholders(root, frame_url_rules):
    """Normalize request-line URLs in frames to {baseURL}/endpoint-path."""
    method_pattern = re.compile(
        r'^(\s*)(GET|POST|PUT|DELETE|PATCH)\s*(.+?)(?=Content-Type:|x-language-name:|Authorization:|\Z)',
        re.IGNORECASE,
    )
    for txbx in root.iter(qn("txbxContent")):
        for para in txbx.iter(qn("p")):
            full_text = get_element_text(para)
            if not full_text.strip():
                continue
            match = method_pattern.match(full_text)
            if not match:
                continue
            url_part = match.group(3).strip()
            new_url = _to_baseurl_endpoint(url_part, frame_url_rules)
            if new_url == url_part:
                continue
            new_text = f"{match.group(1)}{match.group(2)} {new_url}{full_text[match.end():]}"
            _set_paragraph_text(para, new_text)



def _remove_blank_between_headers_and_json(txbx):
    """Remove blank paragraphs between HTTP headers and JSON body inside text boxes."""
    paragraphs = list(txbx.iter(qn("p")))

    # Find JSON start index
    json_start = None
    for i, para in enumerate(paragraphs):
        text = get_element_text(para).strip()
        if text.startswith(('{', '[')):
            json_start = i
            break

    if json_start is None or json_start == 0:
        return

    # Find the last header before JSON (contains ':', not JSON)
    last_header = None
    for i in range(json_start - 1, -1, -1):
        text = get_element_text(paragraphs[i]).strip()
        if text == "":
            continue
        if ':' in text and not text.startswith(('{', '[')):
            last_header = i
            break
        # If we hit non-header, non-blank content, stop
        if text:
            break

    if last_header is None:
        return

    # Remove blank paragraphs between last_header and json_start
    for i in range(json_start - 1, last_header, -1):
        txbx.remove(paragraphs[i])


def normalize_txbx_content(root, method_rules):
    """Format text inside inline text-box frames (Sample Request / Response)."""
    # 1) Format text inside txbxContent
    for txbx in root.iter(qn("txbxContent")):
        for para in txbx.iter(qn("p")):
            runs = list(para.iter(qn("r")))
            _format_txbx_para(runs, method_rules)
            _highlight_variables_in_para(runs)
        _prettify_json_in_txbx(txbx)
        _remove_blank_between_headers_and_json(txbx)

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


def ensure_connection_info_section(root, config, frame_url_rules, table_property_rules):
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
    connection_widths = parse_column_widths(
        table_property_rules, "Generated connection table column widths", 4)
    section = [
        make_paragraph("System Connection Information", "Heading2", config),
        make_paragraph("The following table lists connection parameters used to build API endpoint URLs.", "Normal", config),
        make_connection_info_table(frame_url_rules, connection_widths),
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


def fix_document(document_path, config, table_border_rules, api_body_rules, table_property_rules, method_rules, heading_hierarchy_rules, frame_url_rules):
    tree = etree.parse(document_path)
    root = tree.getroot()

    # Keywords that should be uniform sub-headings
    SUBHEADING_KEYWORDS = [
        "Sample Request",
        "Sample Response",
        "Request Attributes",
        "Response Attributes",
    ]

    ensure_connection_info_section(root, config, frame_url_rules, table_property_rules)
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
                    rebuild_rPr(child, "Verdana", 24, True, None)
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

    # 2c) Attribute tables must carry black borders and tblHeader before step 3 styles them
    normalize_attribute_table_borders_and_headers(root, table_property_rules)

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
    create_path_attribute_tables(root, config, api_body_rules, table_property_rules)

    # 5d) Fix semantic content of attribute tables (move accesstoken to Authorization, add Content-Type)
    normalize_attribute_table_content(root)

    # 5e) Replace invalid regular-cell percentage widths with each table's dxa grid
    normalize_attribute_table_widths(root)
    normalize_connection_table_widths(root, table_property_rules)

    # 6) Enforce attribute table start-header alignment and font color from the template
    normalize_api_table_start_headers(root, table_property_rules)

    # 6a) Normalize table body rows to Verdana 11pt non-bold black
    normalize_table_body_fonts(root)

    # 6a2) Re-centre No / Presence after step 1 applied Normal left alignment to cells
    normalize_attribute_table_column_alignment(root)

    # 6b) Center-align all tables and unify table indent
    apply_table_alignment(root, table_property_rules)

    # 6c) Split label+frame paragraphs so labels stay left-aligned
    split_label_from_frame_paragraphs(root)

    # 6d) Center-align inline text-box frames (Sample Request/Response) and unify their width
    normalize_inline_shapes(root)

    # 6e) Make every sample frame auto-fit its text so no JSON line is clipped
    #     and no frame holds dead space (template 2.8.10)
    normalize_frame_autofit(root)

    # 7) Enforce spacing and blank-line rules inside each API body section
    normalize_api_body_spacing(root, config, api_body_rules)

    # 8) Ensure exactly one blank paragraph before each Heading 3 (except the first one after a Heading 2)
    normalize_heading3_spacing(root, config)

    # 9) Normalize request URLs and format inline text-box frame content
    normalize_url_placeholders(root, frame_url_rules)
    normalize_txbx_content(root, method_rules)

    # 10) Highlight variable placeholders like {domain}, {token} in red across the document
    highlight_variables_in_document(root)

    # 11) Normalize x-language-name header values to lowercase (en, vi, etc.)
    normalize_x_language_name(root)

    tree.write(document_path, xml_declaration=True, encoding="utf-8", standalone=True)
    print(f"[OK] Fixed document.xml")


def pack_docx(source_dir, output_path):
    """Pack unpacked directory into a .docx file."""
    document_xml = os.path.join(source_dir, "word", "document.xml")
    if os.path.isfile(document_xml):
        validate_table_widths(
            etree.parse(document_xml, etree.XMLParser(huge_tree=True)).getroot())
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
    frame_url_rules = parse_frame_url_rules(template_md)
    heading_hierarchy_rules = parse_heading_hierarchy_rules(template_md)
    print(f"[OK] Loaded {len(config)} style definitions from {template_md}")
    print(f"[OK] Loaded {len(table_border_rules)} table border rules from {template_md}")
    print(f"[OK] Loaded {len(table_property_rules)} table property rules from {template_md}")
    print(f"[OK] Loaded {len(api_body_rules)} API body rules from {template_md}")
    print(f"[OK] Loaded {len(method_rules)} frame method rules from {template_md}")
    print(f"[OK] Loaded {len(frame_url_rules)} frame URL rules from {template_md}")
    print(f"[OK] Loaded {len(heading_hierarchy_rules)} heading hierarchy rules from {template_md}")

    styles_xml = os.path.join(unpacked, "word", "styles.xml")
    document_xml = os.path.join(unpacked, "word", "document.xml")
    numbering_xml = os.path.join(unpacked, "word", "numbering.xml")

    fix_styles(styles_xml, config)
    fix_heading_numbering_layout(numbering_xml)
    fix_document(document_xml, config, table_border_rules, api_body_rules, table_property_rules, method_rules, heading_hierarchy_rules, frame_url_rules)

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
