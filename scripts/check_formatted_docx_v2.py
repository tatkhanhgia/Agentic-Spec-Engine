#!/usr/bin/env python
# -*- coding: utf-8 -*-
import zipfile
from lxml import etree
import os, re, sys

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def qn(tag):
    return W + tag

def get_text(element):
    return ''.join(t.text or '' for t in element.iter(qn('t'))).strip()

def get_style_size(styles_root, style_id):
    """Get sz value from style definition."""
    for style in styles_root.iter(qn('style')):
        sid = style.get(qn('styleId'))
        if sid == style_id:
            rPr = style.find(qn('rPr'))
            if rPr is not None:
                sz = rPr.find(qn('sz'))
                if sz is not None:
                    return sz.get(qn('val'))
    return None

files = [
    'formatted/GoPaperless_Normalized_Errors.docx',
    'formatted/GoPaperless Workflow RESTful API Specification V2-Formatted.docx'
]

# Word-boundary misspellings using regex
misspelling_patterns = [
    (r'\bparamater\b', 'parameter'),
    (r'\bresponce\b', 'response'),
    (r'\brequst\b', 'request'),
    (r'\bauthetication\b', 'authentication'),
    (r'\bathentication\b', 'authentication'),
    (r'\bsucces\b(?!s)', 'success'),  # avoid matching inside success
    (r'\bfaiure\b', 'failure'),
    (r'\boccured\b', 'occurred'),
    (r'\bseperate\b', 'separate'),
    (r'\brecieve\b', 'receive'),
    (r'\bexmaple\b', 'example'),
    (r'\bmandetory\b', 'mandatory'),
    (r'\boptinal\b', 'optional'),
    (r'\bdesription\b', 'description'),
    (r'\battributs\b', 'attributes'),
    (r'\breponse\b', 'response'),
    (r'\breqest\b', 'request'),
    (r'\bauthenication\b', 'authentication'),
    (r'\bsignle\b', 'single'),
    (r'\bsing on\b', 'sign on'),
    (r'\bsingn\b', 'sign'),
    (r'\bfaild\b', 'failed'),
    (r'\beror\b', 'error'),
    (r'\berros\b', 'errors'),
    (r'\bmessge\b', 'message'),
    (r'\bmesage\b', 'message'),
    (r'\bfileds\b', 'fields'),
    (r'\bfeilds\b', 'fields'),
    (r'\bfeild\b', 'field'),
    (r'\bunauthorize\b', 'unauthorized'),
    (r'\bunautorized\b', 'unauthorized'),
    (r'\bunauthorised\b', 'unauthorized'),
    (r'\bacess token\b', 'access token'),
    (r'\brefesh token\b', 'refresh token'),
    (r'\brefesh\b', 'refresh'),
    (r'\bacess\b', 'access'),
    (r'\bcontex path\b', 'context path'),
    (r'\bdomian\b', 'domain'),
    (r'\bendpoit\b', 'endpoint'),
    (r'\bend point\b', 'endpoint'),
    (r'\bspecifcation\b', 'specification'),
    (r'\bspecifed\b', 'specified'),
    (r'\bretrive\b', 'retrieve'),
    (r'\bretun\b', 'return'),
    (r'\bretuns\b', 'returns'),
    (r'\bretured\b', 'returned'),
    (r'\breturing\b', 'returning'),
    (r'\bvalus\b', 'values'),
    (r'\bvaluse\b', 'values'),
    (r'\bdefalt\b', 'default'),
    (r'\bdefualt\b', 'default'),
    (r'\brequred\b', 'required'),
    (r'\brequried\b', 'required'),
    (r'\boptionl\b', 'optional'),
    (r'\boptioanl\b', 'optional'),
    (r'\baddtional\b', 'additional'),
    (r'\badditonal\b', 'additional'),
    (r'\binformtion\b', 'information'),
    (r'\binformaton\b', 'information'),
    (r'\bdocment\b', 'document'),
    (r'\bdocuemnt\b', 'document'),
    (r'\bverfication\b', 'verification'),
    (r'\bverificaton\b', 'verification'),
    (r'\bauthenfication\b', 'authentication'),
    (r'\bidentifer\b', 'identifier'),
    (r'\bidentfier\b', 'identifier'),
    (r'\busernme\b', 'username'),
    (r'\bpasword\b', 'password'),
    (r'\bpasssword\b', 'password'),
]

for f in files:
    if not os.path.exists(f):
        print(f'File not found: {f}')
        continue
    print(f'\n=== {f} ===')
    z = zipfile.ZipFile(f)
    doc_xml = z.read('word/document.xml')
    root = etree.fromstring(doc_xml)
    styles_xml = z.read('word/styles.xml')
    styles_root = etree.fromstring(styles_xml)
    z.close()

    body = root.find('.//' + qn('body'))
    if body is None:
        print('No body found')
        continue

    # Check style-level Normal size
    normal_style_size = get_style_size(styles_root, 'Normal')
    print(f'\n--- Style Check ---')
    print(f'  Normal style sz = {normal_style_size} (expected 22 for 11pt body)')

    # Check headings
    print('\n--- Heading Hierarchy Check ---')
    auth_endpoints = ['Client Credentials Authentication', 'OTP Login Authentication', 'Single Sign On Authentication']
    for para in body.iter(qn('p')):
        pPr = para.find(qn('pPr'))
        if pPr is not None:
            pStyle = pPr.find(qn('pStyle'))
            if pStyle is not None:
                val = pStyle.get(qn('val'))
                if val in ['Heading1', 'Heading2', 'Heading3', 'Heading4']:
                    text = get_text(para)
                    if any(ae in text for ae in auth_endpoints):
                        if val != 'Heading3':
                            print(f'  ISSUE: "{text}" should be Heading3 but is {val}')
                        else:
                            print(f'  OK: "{text}" is Heading3')

    # Check spelling with word boundaries
    print('\n--- Spelling Check ---')
    full_text = get_text(body)
    found_issues = []
    for pattern, correct in misspelling_patterns:
        matches = list(re.finditer(pattern, full_text, re.IGNORECASE))
        if matches:
            for m in matches[:5]:  # limit output
                snippet = full_text[max(0, m.start()-20):m.end()+20]
                found_issues.append(f'  ISSUE: Found "{m.group()}" at "...{snippet}..." (should be "{correct}")')
            if len(matches) > 5:
                found_issues.append(f'  ... and {len(matches)-5} more occurrences of "{matches[0].group()}"')
    if found_issues:
        for issue in found_issues:
            print(issue)
    else:
        print('  No common misspellings found.')

    # Check API body section labels formatting (consider style-level sz)
    print('\n--- API Body Section Labels ---')
    labels_to_check = ['Sample Request', 'Sample Response', 'Attributes description']
    for para in body.iter(qn('p')):
        text = get_text(para)
        for label in labels_to_check:
            if text.startswith(label):
                runs = para.findall(qn('r'))
                is_bold = False
                size_ok = False
                for r in runs:
                    rPr = r.find(qn('rPr'))
                    if rPr is not None:
                        if rPr.find(qn('b')) is not None:
                            is_bold = True
                        sz = rPr.find(qn('sz'))
                        if sz is not None and sz.get(qn('val')) == '24':
                            size_ok = True
                # Also accept if Normal style has sz=24
                if not size_ok and normal_style_size == '24':
                    size_ok = True
                if not is_bold:
                    print(f'  ISSUE: "{text[:50]}" is not bold')
                if not size_ok:
                    print(f'  ISSUE: "{text[:50]}" is not 12pt')
                if is_bold and size_ok:
                    pass  # OK, don't spam

    # Check frame content for HTTP method badges (check paragraph-level shading too)
    print('\n--- HTTP Method Badge Check ---')
    method_colors = {
        'GET': '047857',
        'POST': 'C2410C',
        'PUT': '1D4ED8',
        'DELETE': 'B91C1C',
        'PATCH': '6D28D9',
    }
    found_methods = set()
    badge_issues = []
    for para in body.iter(qn('p')):
        text = get_text(para)
        for method, expected_color in method_colors.items():
            if text.startswith(method + ' ') or text.startswith(method + '\t'):
                found_methods.add(method)
                # Check paragraph-level shading first
                pPr = para.find(qn('pPr'))
                para_shd_color = None
                if pPr is not None:
                    shd = pPr.find(qn('shd'))
                    if shd is not None:
                        para_shd_color = shd.get(qn('fill'))
                # Check run-level shading
                runs = para.findall(qn('r'))
                run_shd_color = None
                for r in runs:
                    r_text = get_text(r)
                    if method in r_text:
                        rPr = r.find(qn('rPr'))
                        if rPr is not None:
                            shd = rPr.find(qn('shd'))
                            if shd is not None:
                                run_shd_color = shd.get(qn('fill'))
                actual_color = run_shd_color or para_shd_color
                if actual_color and actual_color.upper() == expected_color:
                    pass  # OK
                elif actual_color:
                    badge_issues.append(f'  ISSUE: {method} badge color is {actual_color}, expected {expected_color}')
                else:
                    badge_issues.append(f'  ISSUE: {method} badge missing shading (no background color)')
    if badge_issues:
        # Deduplicate
        seen = set()
        for bi in badge_issues[:20]:
            if bi not in seen:
                seen.add(bi)
                print(bi)
        if len(badge_issues) > 20:
            print(f'  ... and {len(badge_issues)-20} more badge issues')
    elif not found_methods:
        print('  No HTTP method badges found (this may be OK for error-only doc)')
    else:
        print(f'  All {len(found_methods)} method types have correct badge colors.')

    # Check URL patterns
    print('\n--- URL Pattern Check ---')
    if 'prd-gopaperless.mobile-id.vn' not in full_text:
        print('  WARNING: Current domain prd-gopaperless.mobile-id.vn not found in document')
    else:
        print('  OK: Domain prd-gopaperless.mobile-id.vn found')
    if '/workflow/api' not in full_text:
        print('  WARNING: contextPath /workflow/api not found in document')
    else:
        print('  OK: contextPath /workflow/api found')

    # Check for stray fonts
    print('\n--- Font Check ---')
    stray_fonts = set()
    for r in body.iter(qn('r')):
        rPr = r.find(qn('rPr'))
        if rPr is not None:
            rf = rPr.find(qn('rFonts'))
            if rf is not None:
                ascii_font = rf.get(qn('ascii'))
                if ascii_font and ascii_font != 'Verdana':
                    stray_fonts.add(ascii_font)
    if stray_fonts:
        print(f'  ISSUE: Found non-Verdana fonts: {sorted(stray_fonts)}')
    else:
        print('  OK: All runs use Verdana')

    # Check table header colors
    print('\n--- Table Header Color Check ---')
    wrong_colors = set()
    for tbl in root.iter(qn('tbl')):
        rows = tbl.findall(qn('tr'))
        if rows:
            first_row = rows[0]
            trPr = first_row.find(qn('trPr'))
            is_header = trPr is not None and trPr.find(qn('tblHeader')) is not None
            if is_header:
                for tc in first_row.iter(qn('tc')):
                    tcPr = tc.find(qn('tcPr'))
                    if tcPr is not None:
                        shd = tcPr.find(qn('shd'))
                        if shd is not None:
                            fill = shd.get(qn('fill'))
                            if fill and fill.upper() not in ['0070C0', 'AUTO', '']:
                                wrong_colors.add(fill.upper())
    if wrong_colors:
        print(f'  ISSUE: Table headers with wrong colors: {sorted(wrong_colors)} (expected 0070C0)')
    else:
        print('  OK: All table headers use 0070C0')

    # Check page margins
    print('\n--- Page Setup Check ---')
    sectPr = body.find(qn('sectPr'))
    if sectPr is not None:
        pgMar = sectPr.find(qn('pgMar'))
        if pgMar is not None:
            left = pgMar.get(qn('left'))
            right = pgMar.get(qn('right'))
            top = pgMar.get(qn('top'))
            bottom = pgMar.get(qn('bottom'))
            print(f'  Margins: L={left}, R={right}, T={top}, B={bottom}')
            if left != '706' or right != '706':
                print('  ISSUE: Left/Right margins should be 706 (35.3 pt * 20 twips/pt)')
            if top != '1844':
                print('  ISSUE: Top margin should be 1844 (92.2 pt * 20 twips/pt)')
            if bottom != '1700':
                print('  ISSUE: Bottom margin should be 1700 (85.0 pt * 20 twips/pt)')
        else:
            print('  WARNING: No pgMar found')
    else:
        print('  WARNING: No sectPr found')
