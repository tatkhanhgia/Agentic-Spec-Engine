#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Căn lại w:tcW của từng ô theo w:gridCol tương ứng (type dxa) cho bảng thuộc tính,
đúng theo chuẩn mà check_frame_format.py dùng để bắt lỗi độ rộng cột.

Chỉ sửa các dòng có số ô == số cột grid (dòng merge/span bị bỏ qua để không phá cấu trúc).
Không đụng tới bất kỳ thứ gì khác trong document.xml.
"""
import sys, zipfile, shutil, os
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def q(t): return "{%s}%s" % (W, t)

def grid_widths(tbl):
    ws = []
    for gc in tbl.findall(q("tblGrid") + "/" + q("gridCol")):
        v = gc.get(q("w"))
        if v is None or not v.isdigit():
            return []
        ws.append(int(v))
    return ws

def set_tcw(tc, width):
    tcPr = tc.find(q("tcPr"))
    if tcPr is None:
        tcPr = etree.Element(q("tcPr"))
        tc.insert(0, tcPr)
    tcW = tcPr.find(q("tcW"))
    if tcW is None:
        tcW = etree.SubElement(tcPr, q("tcW"))
        # tcW phải nằm đúng chỗ tcPr (thứ tự schema: tcW ngay sau tcPr mở đầu)
        tcPr.insert(0, tcW)
    tcW.set(q("w"), str(width))
    tcW.set(q("type"), "dxa")

def fix(src, dst):
    shutil.copyfile(src, dst)  # backup-safe: sửa trên bản copy
    shutil.copyfile(src, src + ".orig") if not os.path.exists(src + ".orig") else None
    zin = zipfile.ZipFile(src, "r")
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(q("body"))
    changed_rows = changed_cells = skipped_merge = no_grid = tables = 0
    for tbl in body.iter(q("tbl")):
        tables += 1
        gw = grid_widths(tbl)
        if not gw:
            no_grid += 1
            continue
        for tr in tbl.findall(q("tr")):
            cells = tr.findall(q("tc"))
            if len(cells) != len(gw):
                skipped_merge += 1
                continue
            row_changed = False
            for tc, expected in zip(cells, gw):
                tcPr = tc.find(q("tcPr"))
                tcW = tcPr.find(q("tcW")) if tcPr is not None else None
                cur = tcW.get(q("w")) if tcW is not None else None
                typ = tcW.get(q("type")) if tcW is not None else None
                if cur != str(expected) or typ != "dxa":
                    set_tcw(tc, expected)
                    changed_cells += 1
                    row_changed = True
            if row_changed:
                changed_rows += 1
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    zin.close()
    # ghi lại document.xml vào bản copy
    _rewrite(dst, new_xml)
    print("tables=%d  no_grid=%d  rows_changed=%d  cells_changed=%d  rows_skipped_merge=%d"
          % (tables, no_grid, changed_rows, changed_cells, skipped_merge))

def _rewrite(path, new_xml):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml
            zout.writestr(item, data)
    os.replace(tmp, path)

if __name__ == "__main__":
    fix(sys.argv[1], sys.argv[2])
