#!/usr/bin/env python
# -*- coding: utf-8 -*-
import zipfile
from lxml import etree
import os, re

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def qn(tag): return W + tag

def get_text(element):
    return ''.join(t.text or '' for t in element.iter(qn('t'))).strip()

def get_para_style(para):
    pPr = para.find(qn('pPr'))
    if pPr is not None:
        pStyle = pPr.find(qn('pStyle'))
        if pStyle is not None:
            return pStyle.get(qn('val'))
    return 'Normal'

def is_bold_para(para):
    """Check if all runs in paragraph are bold (or paragraph has bold style)."""
    runs = para.findall(qn('r'))
    if not runs:
        return False
    for r in runs:
        rPr = r.find(qn('rPr'))
        if rPr is not None:
            if rPr.find(qn('b')) is None:
                return False
        else:
            return False
    return True

def get_para_size(para):
    """Get effective size from first run."""
    runs = para.findall(qn('r'))
    for r in runs:
        rPr = r.find(qn('rPr'))
        if rPr is not None:
            sz = rPr.find(qn('sz'))
            if sz is not None:
                return sz.get(qn('val'))
    return None

def get_para_indent_left(para):
    pPr = para.find(qn('pPr'))
    if pPr is not None:
        ind = pPr.find(qn('ind'))
        if ind is not None:
            return ind.get(qn('left'))
    return None

files = [
    'formatted/GoPaperless_Normalized_Errors.docx',
    'formatted/GoPaperless Workflow RESTful API Specification V2-Formatted.docx'
]

for f in files:
    if not os.path.exists(f):
        print(f'File not found: {f}')
        continue
    print(f'\n=== {f} ===')
    z = zipfile.ZipFile(f)
    doc_xml = z.read('word/document.xml')
    root = etree.fromstring(doc_xml)
    z.close()

    body = root.find('.//' + qn('body'))
    if body is None:
        continue

    # Detailed label check
    print('\n--- Label Detail Check (first 5 each) ---')
    labels_to_check = ['Sample Request', 'Sample Response', 'Attributes description']
    for para in list(body.iter(qn('p')))[:500]:
        text = get_text(para)
        for label in labels_to_check:
            if text.startswith(label):
                style = get_para_style(para)
                bold = is_bold_para(para)
                size = get_para_size(para)
                indent = get_para_indent_left(para)
                print(f'  "{text[:40]}" style={style} bold={bold} size={size} indent={indent}')
                break

    # Check method badge detail
    print('\n--- Method Badge Detail (first 10 issues) ---')
    method_colors = {'GET': '047857', 'POST': 'C2410C', 'PUT': '1D4ED8', 'DELETE': 'B91C1C', 'PATCH': '6D28D9'}
    issue_count = 0
    for para in body.iter(qn('p')):
        text = get_text(para)
        for method, expected_color in method_colors.items():
            if text.startswith(method + ' ') or text.startswith(method + '\t'):
                # Find the run containing the method text
                runs = para.findall(qn('r'))
                method_run = None
                for r in runs:
                    r_text = get_text(r)
                    if method in r_text:
                        method_run = r
                        break
                if method_run is None:
                    issue_count += 1
                    if issue_count <= 10:
                        print(f'  ISSUE: Could not find run for {method} in "{text[:60]}"')
                    continue
                rPr = method_run.find(qn('rPr'))
                if rPr is None:
                    issue_count += 1
                    if issue_count <= 10:
                        print(f'  ISSUE: {method} run has no rPr in "{text[:60]}"')
                    continue
                shd = rPr.find(qn('shd'))
                highlight = rPr.find(qn('highlight'))
                if shd is not None:
                    fill = shd.get(qn('fill'))
                    if fill and fill.upper() == expected_color:
                        pass
                    else:
                        issue_count += 1
                        if issue_count <= 10:
                            print(f'  ISSUE: {method} shd fill={fill} expected={expected_color}')
                elif highlight is not None:
                    hval = highlight.get(qn('val'))
                    issue_count += 1
                    if issue_count <= 10:
                        print(f'  ISSUE: {method} uses highlight={hval} instead of shd')
                else:
                    issue_count += 1
                    if issue_count <= 10:
                        print(f'  ISSUE: {method} missing both shd and highlight in "{text[:60]}"')
                break
    if issue_count > 10:
        print(f'  ... total {issue_count} badge issues')

    # Check blank paragraphs between sections
    print('\n--- Blank Paragraph Check (sample) ---')
    blank_count = 0
    for para in body.iter(qn('p')):
        text = get_text(para)
        if text == '':
            style = get_para_style(para)
            pPr = para.find(qn('pPr'))
            spacing_info = ''
            if pPr is not None:
                spacing = pPr.find(qn('spacing'))
                if spacing is not None:
                    before = spacing.get(qn('before'))
                    after = spacing.get(qn('after'))
                    spacing_info = f' before={before} after={after}'
            blank_count += 1
            if blank_count <= 5:
                print(f'  Blank para style={style}{spacing_info}')
    print(f'  Total blank paragraphs: {blank_count}')

    # Check variable placeholder color
    print('\n--- Variable Placeholder Color Check ---')
    placeholder_issues = []
    for r in body.iter(qn('r')):
        r_text = get_text(r)
        if '{' in r_text and '}' in r_text:
            rPr = r.find(qn('rPr'))
            if rPr is not None:
                color = rPr.find(qn('color'))
                if color is not None:
                    val = color.get(qn('val'))
                    if val and val.upper() != 'FF0000':
                        placeholder_issues.append(f'  Placeholder "{r_text}" color={val}')
                else:
                    placeholder_issues.append(f'  Placeholder "{r_text}" has no explicit color')
    if placeholder_issues:
        for pi in placeholder_issues[:10]:
            print(pi)
        if len(placeholder_issues) > 10:
            print(f'  ... and {len(placeholder_issues)-10} more')
    else:
        print('  No variable placeholders found or all are red')
