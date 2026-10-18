#!/usr/bin/env python
# -*- coding: utf-8 -*-
import zipfile
from lxml import etree
import os

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def qn(tag): return W + tag

def get_text(element):
    return ''.join(t.text or '' for t in element.iter(qn('t'))).strip()

files = [
    'formatted/GoPaperless_Normalized_Errors.docx',
    'formatted/GoPaperless Workflow RESTful API Specification V2-Formatted.docx'
]

for f in files:
    if not os.path.exists(f):
        continue
    print(f'\n=== {f} ===')
    z = zipfile.ZipFile(f)
    doc_xml = z.read('word/document.xml')
    root = etree.fromstring(doc_xml)
    z.close()

    body = root.find('.//' + qn('body'))
    if body is None:
        continue

    labels_to_check = ['Sample Request', 'Sample Response', 'Attributes description']
    for para in body.iter(qn('p')):
        text = get_text(para)
        for label in labels_to_check:
            if text.startswith(label):
                pPr = para.find(qn('pPr'))
                style = 'Normal'
                if pPr is not None:
                    pStyle = pPr.find(qn('pStyle'))
                    if pStyle is not None:
                        style = pStyle.get(qn('val'))
                runs = para.findall(qn('r'))
                bold = False
                size = None
                for r in runs:
                    rPr = r.find(qn('rPr'))
                    if rPr is not None:
                        if rPr.find(qn('b')) is not None:
                            bold = True
                        sz = rPr.find(qn('sz'))
                        if sz is not None:
                            size = sz.get(qn('val'))
                print(f'  Label: "{text[:45]}" style={style} bold={bold} run_size={size}')
                break
