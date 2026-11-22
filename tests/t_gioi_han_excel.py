# -*- coding: utf-8 -*-
"""Giới hạn của Excel mà LibreOffice không báo: công thức tối đa 8192 ký tự, lồng hàm tối đa 64 tầng.
Vượt giới hạn thì Excel báo "Repaired" và xóa công thức khi mở (đã xảy ra ở bản 6.9 đầu tiên: LUU_NIEN, LUC_THAN)."""
import sys, zipfile
from lxml import etree
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def chay(xlsx):
    z = zipfile.ZipFile(xlsx); loi = []; mx = 0
    for n in z.namelist():
        if not n.startswith('xl/worksheets/sheet'): continue
        for c in etree.fromstring(z.read(n)).iter(NS + 'c'):
            f = c.find(NS + 'f')
            if f is None or not f.text: continue
            t = f.text; mx = max(mx, len(t)); d = m = 0; q = False
            for ch in t:
                if ch == '"': q = not q
                elif not q:
                    if ch == '(': d += 1; m = max(m, d)
                    elif ch == ')': d -= 1
            if len(t) > 8192: loi.append((n, c.get('r'), 'dài', len(t)))
            if m > 64: loi.append((n, c.get('r'), 'lồng', m))
    return loi, mx


if __name__ == '__main__':
    loi, mx = chay(sys.argv[1])
    print('công thức dài nhất %d ký tự; vượt giới hạn Excel: %d ô' % (mx, len(loi)))
    for x in loi[:20]: print('  ', x)
    sys.exit(1 if loi else 0)
