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

files = [
    'formatted/GoPaperless_Normalized_Errors.docx',
    'formatted/GoPaperless Workflow RESTful API Specification V2-Formatted.docx'
]

common_misspellings = [
    ('paramater', 'parameter'),
    ('responce', 'response'),
    ('requst', 'request'),
    ('authetication', 'authentication'),
    ('athentication', 'authentication'),
    ('succes', 'success'),
    ('faiure', 'failure'),
    ('occured', 'occurred'),
    ('seperate', 'separate'),
    ('recieve', 'receive'),
    ('exmaple', 'example'),
    ('mandetory', 'mandatory'),
    ('optinal', 'optional'),
    ('desription', 'description'),
    ('attributs', 'attributes'),
    ('reponse', 'response'),
    ('reqest', 'request'),
    ('authenication', 'authentication'),
    ('signle', 'single'),
    ('sing on', 'sign on'),
    ('singn', 'sign'),
    ('faild', 'failed'),
    ('eror', 'error'),
    ('erros', 'errors'),
    ('messge', 'message'),
    ('mesage', 'message'),
    ('fileds', 'fields'),
    ('feilds', 'fields'),
    ('feild', 'field'),
    ('unauthorize', 'unauthorized'),
    ('unautorized', 'unauthorized'),
    ('unauthorised', 'unauthorized'),
    ('acess token', 'access token'),
    ('refesh token', 'refresh token'),
    ('refesh', 'refresh'),
    ('acess', 'access'),
    ('contex path', 'context path'),
    ('domian', 'domain'),
    ('endpoit', 'endpoint'),
    ('api specifcation', 'api specification'),
    ('specifcation', 'specification'),
    ('specifed', 'specified'),
    ('retrive', 'retrieve'),
    ('retun', 'return'),
    ('retuns', 'returns'),
    ('retured', 'returned'),
    ('returing', 'returning'),
    ('valus', 'values'),
    ('valuse', 'values'),
    ('defalt', 'default'),
    ('defualt', 'default'),
    ('requred', 'required'),
    ('requried', 'required'),
    ('optionl', 'optional'),
    ('optioanl', 'optional'),
    ('addtional', 'additional'),
    ('additonal', 'additional'),
    ('informtion', 'information'),
    ('informaton', 'information'),
    ('docment', 'document'),
    ('docuemnt', 'document'),
    ('verfication', 'verification'),
    ('verificaton', 'verification'),
    ('authenfication', 'authentication'),
    ('identifer', 'identifier'),
    ('identfier', 'identifier'),
    ('usernme', 'username'),
    ('pasword', 'password'),
    ('passsword', 'password'),
    ('Authetication', 'Authentication'),
    ('Athentication', 'Authentication'),
    ('Authenication', 'Authentication'),
    ('Authenfication', 'Authentication'),
    ('Paramater', 'Parameter'),
    ('Responce', 'Response'),
    ('Requst', 'Request'),
    ('Succes', 'Success'),
    ('Faiure', 'Failure'),
    ('Occured', 'Occurred'),
    ('Seperate', 'Separate'),
    ('Recieve', 'Receive'),
    ('Exmaple', 'Example'),
    ('Mandetory', 'Mandatory'),
    ('Optinal', 'Optional'),
    ('Desription', 'Description'),
    ('Attributs', 'Attributes'),
    ('Reponse', 'Response'),
    ('Reqest', 'Request'),
    ('Signle', 'Single'),
    ('Sing On', 'Sign On'),
    ('Singn', 'Sign'),
    ('Faild', 'Failed'),
    ('Eror', 'Error'),
    ('Erros', 'Errors'),
    ('Messge', 'Message'),
    ('Mesage', 'Message'),
    ('Fileds', 'Fields'),
    ('Feilds', 'Fields'),
    ('Feild', 'Field'),
    ('Unauthorize', 'Unauthorized'),
    ('Unautorized', 'Unauthorized'),
    ('Unauthorised', 'Unauthorized'),
    ('Acess Token', 'Access Token'),
    ('Refesh Token', 'Refresh Token'),
    ('Refesh', 'Refresh'),
    ('Acess', 'Access'),
    ('Contex Path', 'Context Path'),
    ('Domian', 'Domain'),
    ('Endpoit', 'Endpoint'),
    ('End Point', 'Endpoint'),
    ('Api Specifcation', 'API Specification'),
    ('Specifcation', 'Specification'),
    ('Specifed', 'Specified'),
    ('Retrive', 'Retrieve'),
    ('Retun', 'Return'),
    ('Retuns', 'Returns'),
    ('Retured', 'Returned'),
    ('Returing', 'Returning'),
    ('Valus', 'Values'),
    ('Valuse', 'Values'),
    ('Defalt', 'Default'),
    ('Defualt', 'Default'),
    ('Requred', 'Required'),
    ('Requried', 'Required'),
    ('Optionl', 'Optional'),
    ('Optioanl', 'Optional'),
    ('Addtional', 'Additional'),
    ('Additonal', 'Additional'),
    ('Informtion', 'Information'),
    ('Informaton', 'Information'),
    ('Docment', 'Document'),
    ('Docuemnt', 'Document'),
    ('Verfication', 'Verification'),
    ('Verificaton', 'Verification'),
    ('Identifer', 'Identifier'),
    ('Identfier', 'Identifier'),
    ('Usernme', 'Username'),
    ('Pasword', 'Password'),
    ('Passsword', 'Password'),
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
        print('No body found')
        continue

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

    # Check spelling
    print('\n--- Spelling Check ---')
    full_text = get_text(body)
    full_text_lower = full_text.lower()
    found_issues = []
    for wrong, correct in common_misspellings:
        if wrong in full_text:
            count = full_text.count(wrong)
            found_issues.append(f'  ISSUE: Found "{wrong}" (should be "{correct}") - {count} occurrence(s)')
    if found_issues:
        for issue in found_issues:
            print(issue)
    else:
        print('  No common misspellings found.')

    # Check API body section labels formatting
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
                if not is_bold:
                    print(f'  ISSUE: "{text[:50]}" is not bold')
                if not size_ok:
                    print(f'  ISSUE: "{text[:50]}" is not 12pt (val=24 half-points)')
                if is_bold and size_ok:
                    print(f'  OK: "{label}" label is bold 12pt')

    # Check frame content for HTTP method badges
    print('\n--- HTTP Method Badge Check ---')
    method_colors = {
        'GET': '047857',
        'POST': 'C2410C',
        'PUT': '1D4ED8',
        'DELETE': 'B91C1C',
        'PATCH': '6D28D9',
    }
    found_methods = set()
    for para in body.iter(qn('p')):
        text = get_text(para)
        for method, expected_color in method_colors.items():
            if text.startswith(method + ' ') or text.startswith(method + '\t'):
                found_methods.add(method)
                runs = para.findall(qn('r'))
                for r in runs:
                    r_text = get_text(r)
                    if method in r_text:
                        rPr = r.find(qn('rPr'))
                        if rPr is not None:
                            shd = rPr.find(qn('shd'))
                            if shd is not None:
                                fill = shd.get(qn('fill'))
                                if fill and fill.upper() == expected_color:
                                    print(f'  OK: {method} badge color {fill}')
                                elif fill:
                                    print(f'  ISSUE: {method} badge color is {fill}, expected {expected_color}')
                                else:
                                    print(f'  ISSUE: {method} badge has shd but no fill')
                            else:
                                print(f'  ISSUE: {method} badge missing shd (no background color)')
    if not found_methods:
        print('  No HTTP method badges found (this may be OK for error-only doc)')

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
