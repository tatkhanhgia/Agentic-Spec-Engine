import sys
import zipfile
from lxml import etree

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def qn(tag):
    return W + tag


def text(el):
    return ''.join(t.text or '' for t in el.iter(qn('t'))).strip()


def attrs(el):
    if el is None:
        return None
    return {k.split('}')[-1]: v for k, v in el.attrib.items()}


def inspect_docx(path, limit=80):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    shown = 0
    for p in root.iter(qn('p')):
        pPr = p.find(qn('pPr'))
        if pPr is None:
            continue
        pStyle = pPr.find(qn('pStyle'))
        style = pStyle.get(qn('val')) if pStyle is not None else None
        if style not in {'Heading1', 'Heading2', 'Heading3', 'Heading4'}:
            continue
        print('=' * 90)
        print(style, repr(text(p)))
        print('pPr children:', [child.tag.split('}')[-1] for child in pPr])
        for name in ['numPr', 'tabs', 'ind', 'spacing', 'jc']:
            child = pPr.find(qn(name))
            if child is not None:
                print(name, attrs(child))
                for sub in child:
                    print('  ', sub.tag.split('}')[-1], attrs(sub))
        shown += 1
        if shown >= limit:
            break


if __name__ == '__main__':
    inspect_docx(sys.argv[1])
