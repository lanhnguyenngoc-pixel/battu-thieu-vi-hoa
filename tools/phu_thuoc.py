# -*- coding: utf-8 -*-
"""Tính trước tập ô công thức PHỤ THUỘC (trực tiếp hoặc gián tiếp) vào ô vào của Bảng dữ liệu (KHONG_RO_GIO!C6).
Lõi tính dùng tập này để lõi phụ của vòng 12 giờ chỉ tính lại các ô đổi theo giờ; ô khác lấy từ lõi chính
(cùng ô nhập, trừ C6 → cùng giá trị). Chỉ là tối ưu tốc độ – không đổi kết quả (kiểm bằng T1b, T8).
Dùng: python phu_thuoc.py <workbook.xlsx> <battu_core/sheets>"""
import sys, os, re, bisect, collections, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from doc_xlsx import doc
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries, column_index_from_string, coordinate_from_string


def main():
    wbp, outdir = sys.argv[1:3]
    sheets, cells, names, dts = doc(wbp)
    def rc(ref):
        col, row = coordinate_from_string(ref); return row, column_index_from_string(col)
    idx = collections.defaultdict(lambda: collections.defaultdict(list))
    F = {}
    for s in sheets:
        for ref, v in cells[s].items():
            if v[0] == 'f':
                r, c = rc(ref); F[(s, r, c)] = v[1]; idx[s][c].append(r)
    for s in idx:
        for c in idx[s]: idx[s][c].sort()
    vao = set()
    for s, lst in dts.items():
        for ref, at in lst: vao.add((s,) + rc(at['r1']))
    # cạnh: ô công thức → (ô công thức trong vùng tham chiếu) và cờ "tham chiếu trực tiếp ô vào"
    dep = {}; truc_tiep = set()
    for (s, r, c), txt in F.items():
        out = []
        if txt is None: dep[(s, r, c)] = out; continue
        for t in Tokenizer('=' + txt).items:
            if t.type != 'OPERAND' or t.subtype != 'RANGE': continue
            x = names.get(t.value, t.value)
            if '!' in x: sh, a = x.rsplit('!', 1); sh = sh.strip("'").replace("''", "'")
            else: sh, a = s, x
            if sh not in idx and sh not in sheets: continue
            try: c1, r1, c2, r2 = range_boundaries(a.replace('$', ''))
            except Exception: continue
            if c1 is None: c1, c2 = 1, 16384
            if r1 is None: r1, r2 = 1, 1048576
            for (vs, vr, vc) in vao:
                if vs == sh and r1 <= vr <= r2 and c1 <= vc <= c2: truc_tiep.add((s, r, c))
            for cc, rows in idx.get(sh, {}).items():
                if c1 <= cc <= c2:
                    i = bisect.bisect_left(rows, r1); j = bisect.bisect_right(rows, r2)
                    out += [(sh, rr, cc) for rr in rows[i:j]]
        dep[(s, r, c)] = out
    # lan truyền ngược: ô phụ thuộc nếu tham chiếu trực tiếp ô vào hoặc tham chiếu ô phụ thuộc
    rev = collections.defaultdict(list)
    for a, lst in dep.items():
        for b in lst: rev[b].append(a)
    pt = set(truc_tiep); st = list(truc_tiep)
    while st:
        x = st.pop()
        for y in rev.get(x, ()):
            if y not in pt: pt.add(y); st.append(y)
    si = {s: i for i, s in enumerate(sheets)}
    L = ['# -*- coding: utf-8 -*-', '"""SINH TỰ ĐỘNG bởi tools/phu_thuoc.py – ô công thức phụ thuộc ô vào Bảng dữ liệu (KHONG_RO_GIO!C6)."""',
         'O_VAO = %r' % sorted((si[s], r, c) for s, r, c in vao),
         'TRUC_TIEP = frozenset(%r)' % sorted((si[s], r, c) for s, r, c in truc_tiep),
         'PHU_THUOC = frozenset(%r)' % sorted((si[s], r, c) for s, r, c in pt)]
    open(os.path.join(outdir, '_phu_thuoc.py'), 'w').write('\n'.join(L) + '\n')
    print('ô vào', sorted(vao), '| ô phụ thuộc:', len(pt), '/', len(F),
          dict(collections.Counter(s for s, r, c in pt).most_common(30)))


if __name__ == '__main__':
    main()
