# -*- coding: utf-8 -*-
"""Dịch NGUYÊN VĂN các quy tắc Định dạng có điều kiện của workbook 6.3 → battu_core/sheets/_dinh_dang_dk.py.
Công thức điều kiện dịch bằng cùng bộ dịch với ô công thức (tools/dich_cong_thuc.Dich); địa chỉ tương đối tính
theo ô góc trên-trái của vùng đầu tiên trong sqref (như Excel/LibreOffice). Kiểu định dạng lấy từ styles.xml <dxfs>.
Dùng: python tools/dich_dinh_dang_dk.py <workbook.xlsx> <thư_mục_ra>"""
import sys, os, re, zipfile, hashlib, html
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl.utils.cell import range_boundaries
from doc_xlsx import doc
from dich_cong_thuc import Dich


def doc_dxf(st):
    m = re.search(r'<dxfs[^>]*>(.*?)</dxfs>', st, re.S)
    ra = []
    for dx in re.findall(r'<dxf>(.*?)</dxf>', m.group(1) if m else '', re.S):
        k = {}
        f = re.search(r'<font>(.*?)</font>', dx, re.S)
        if f:
            if re.search(r'<b( val="(1|true)")?/>', f.group(1)): k['dam'] = True
            c = re.search(r'<color rgb="([0-9A-Fa-f]{8})"', f.group(1))
            if c: k['chu'] = '#' + c.group(1)[2:].lower()
        b = re.search(r'<bgColor rgb="([0-9A-Fa-f]{8})"', dx)
        if b: k['nen'] = '#' + b.group(1)[2:].lower()
        ra.append(k)
    return ra


def main():
    wbp, outdir = sys.argv[1:3]
    sha = hashlib.sha256(open(wbp, 'rb').read()).hexdigest()
    sheets, cells, names, dts = doc(wbp)
    D = Dich(sheets, names)
    z = zipfile.ZipFile(wbp)
    wb = z.read('xl/workbook.xml').decode('utf-8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rm = {}
    for m in re.finditer(r'<Relationship [^>]*>', rels):
        t = m.group(0); i = re.search(r'Id="([^"]+)"', t).group(1); g = re.search(r'Target="([^"]+)"', t).group(1)
        rm[i] = g.lstrip('/').replace('xl/', '', 1) if g.startswith('/xl/') else g
    dxf = doc_dxf(z.read('xl/styles.xml').decode('utf-8'))
    L = ['# -*- coding: utf-8 -*-',
         '"""Định dạng có điều kiện – SINH TỰ ĐỘNG bởi tools/dich_dinh_dang_dk.py, KHÔNG SỬA TAY.',
         'Nguồn: %s (SHA-256 %s).' % (os.path.basename(wbp), sha),
         'QUY_TAC: (sheet, [(c1, r1, c2, r2)…], ưu tiên, hàm điều kiện g(E, r, c), kiểu {dam, chu, nen})."""',
         'from battu_core import xl as X', '']
    ds = []; n = 0
    for m in re.finditer(r'<sheet [^>]*>', wb):
        t = m.group(0)
        ten = html.unescape(re.search(r'name="([^"]+)"', t).group(1)); rid = re.search(r'r:id="([^"]+)"', t).group(1)
        x = z.read('xl/' + rm[rid]).decode('utf-8')
        for cf in re.finditer(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', x, re.S):
            vung = [range_boundaries(s) for s in cf.group(1).split()]
            c0, r0 = vung[0][0], vung[0][1]
            for ru in re.finditer(r'<cfRule ([^>]*)>(.*?)</cfRule>', cf.group(2), re.S):
                at = ru.group(1)
                kieu = re.search(r'type="([^"]+)"', at).group(1)
                if kieu != 'expression': raise SystemExit('Chưa hỗ trợ kiểu quy tắc %s (%s)' % (kieu, ten))
                ut = int(re.search(r'priority="(\d+)"', at).group(1)); di = int(re.search(r'dxfId="(\d+)"', at).group(1))
                f = html.unescape(re.search(r'<formula>(.*?)</formula>', ru.group(2), re.S).group(1))
                code = D.gen(D.parse(f), ten, r0, c0, 'v')
                fn = 'g%d' % n; n += 1
                L += ['def %s(E, r, c):' % fn, '    # %s %s: =%s' % (ten, cf.group(1), f), '    return %s' % code, '']
                ds.append((ten, vung, ut, fn, dxf[di]))
    L.append('QUY_TAC = [')
    for ten, vung, ut, fn, k in ds: L.append('    (%r, %r, %d, %s, %r),' % (ten, vung, ut, fn, k))
    L.append(']')
    open(os.path.join(outdir, '_dinh_dang_dk.py'), 'w').write('\n'.join(L) + '\n')
    print('Định dạng có điều kiện: %d quy tắc, %d kiểu' % (len(ds), len(dxf)))


if __name__ == '__main__':
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
    main()
