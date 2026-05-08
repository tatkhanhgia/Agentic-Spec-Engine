from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def qn(t): return W+t

tree = etree.parse('unpacked_docx/word/document.xml')
root = tree.getroot()

def get_text(e):
    return ''.join(t.text or '' for t in e.iter(qn('t'))).strip()

def get_para_style(para):
    pPr = para.find(qn('pPr'))
    if pPr is not None:
        ps = pPr.find(qn('pStyle'))
        if ps is not None:
            return ps.get(qn('val'))
    return None

def is_heading(para, levels):
    return get_para_style(para) in levels

children = list(root.find('.//' + qn('body')))

# Find Upload Document section
in_upload = False
section_nodes = []
for child in children:
    if child.tag == qn('p') and is_heading(child, {'Heading3', 'Heading4'}):
        text = get_text(child)
        if text == 'Upload Document':
            in_upload = True
            section_nodes = []
            continue
        elif in_upload and is_heading(child, {'Heading1', 'Heading2', 'Heading3', 'Heading4'}):
            break
    if in_upload:
        section_nodes.append(child)

# Print tables in this section
for node in section_nodes:
    if node.tag == qn('tbl'):
        rows = []
        for tr in node.findall(qn('tr')):
            cells = [get_text(tc) for tc in tr.findall(qn('tc'))]
            rows.append(cells)
        print('Table header:', rows[0][0] if rows else 'NONE')
        for row in rows:
            print('  ', row)
        print()
