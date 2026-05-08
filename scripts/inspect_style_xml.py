import sys
import zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def qn(tag):
    return W + tag


def attrs(el):
    return None if el is None else {k.split('}')[-1]: v for k, v in el.attrib.items()}

with zipfile.ZipFile(sys.argv[1]) as z:
    root = etree.fromstring(z.read('word/styles.xml'))

for sid in ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Normal', 'ListParagraph']:
    style = root.xpath(f'.//w:style[@w:styleId="{sid}"]', namespaces={'w': W[1:-1]})[0]
    print('=' * 80)
    print(sid)
    pPr = style.find(qn('pPr'))
    print('pPr children:', [] if pPr is None else [c.tag.split('}')[-1] for c in pPr])
    if pPr is not None:
        for name in ['numPr','tabs','ind','spacing','jc','outlineLvl']:
            child = pPr.find(qn(name))
            if child is not None:
                print(name, attrs(child))
                for sub in child:
                    print(' ', sub.tag.split('}')[-1], attrs(sub))
