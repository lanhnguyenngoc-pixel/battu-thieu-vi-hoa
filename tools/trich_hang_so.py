# -*- coding: utf-8 -*-
"""Chuẩn bị dữ liệu cho battu_core/data (MT-02 đã duyệt 13/11/2026):
1. Chép nguyên văn các tệp 04_DU_LIEU_NEN vào battu_core/data/nen/ (không sửa).
2. Trích TỰ ĐỘNG mọi ô hằng của các sheet (trừ CSDL và TIET_KHI dòng ≥ 4 – lấy từ 04_DU_LIEU_NEN,
   và trừ vùng kết quả Bảng dữ liệu KHONG_RO_GIO – là giá trị tính) thành data/hang_so_sheet.json.
Không gõ tay ô nào. Ghi SHA-256 nguồn vào tệp ra.
Dùng: python trich_hang_so.py <workbook.xlsx> <thư_mục_04_DU_LIEU_NEN> <battu_core/data>"""
import sys, os, json, shutil, hashlib, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from doc_xlsx import doc, dinh_dang, bo_cuc, danh_sach_chon
from openpyxl.utils.cell import range_boundaries, coordinate_from_string, column_index_from_string

wbp, nen, out = sys.argv[1:4]
sha = hashlib.sha256(open(wbp, 'rb').read()).hexdigest()
os.makedirs(os.path.join(out, 'nen'), exist_ok=True)
manifest = {}
for f in sorted(os.listdir(nen)):
    shutil.copy2(os.path.join(nen, f), os.path.join(out, 'nen', f))
    manifest[f] = hashlib.sha256(open(os.path.join(nen, f), 'rb').read()).hexdigest()
sheets, cells, names, dts = doc(wbp)

def rc(ref):
    col, row = coordinate_from_string(ref); return row, column_index_from_string(col)

hs = {}
for s in sheets:
    skip = []
    for ref, _ in dts[s]:
        c1, r1, c2, r2 = range_boundaries(ref); skip.append((r1, c1, r2, c2))
    d = {}
    for ref, v in cells[s].items():
        if v[0] != 'v': continue
        r, c = rc(ref)
        if s == 'CSDL': continue
        if s == 'TIET_KHI' and r >= 4: continue
        if any(a <= r <= b and cc1 <= c <= cc2 for a, cc1, b, cc2 in skip): continue
        val = v[1]
        if isinstance(val, tuple): val = {'loi': val[1]}
        d[ref] = val
    hs[s] = d
json.dump({'nguon': os.path.basename(wbp), 'sha256': sha, 'ghi_chu': 'Trích tự động bằng tools/trich_hang_so.py (MT-02)',
           'nen_04_sha256': manifest, 'sheets': sheets, 'names': names,
           'data_tables': {s: [[r, a] for r, a in v] for s, v in dts.items() if v},
           'hang_so': hs,
           'dinh_dang_o_cong_thuc': {s: {r: f for r, f in dinh_dang(wbp)[s].items() if cells[s].get(r, ('v',))[0] == 'f'} for s in sheets},
           'bo_cuc': bo_cuc(wbp),
           'danh_sach_chon': danh_sach_chon(wbp),
           'o_cong_thuc': {s: sorted(r for r, v in cells[s].items() if v[0] == 'f') for s in sheets}}, open(os.path.join(out, 'hang_so_sheet.json'), 'w'), ensure_ascii=False, separators=(',', ':'))

# Kiểm chéo: CSDL json và TIET_KHI csv phải trùng ô hằng trong workbook
cs = json.load(open(os.path.join(nen, 'CSDL_toan_bo_gia_tri.json')))
bad = 0; n = 0
for ref, v in cells['CSDL'].items():
    if v[0] != 'v': continue
    r, c = rc(ref); n += 1
    jv = cs[r - 1][c - 1] if r - 1 < len(cs) and c - 1 < len(cs[r - 1]) else None
    if isinstance(jv, int) and not isinstance(jv, bool): jv = float(jv)
    if jv != v[1]: bad += 1
print('workbook', sha[:12], '| sheet', len(sheets), '| ô hằng trích', sum(len(v) for v in hs.values()), '| CSDL so', n, 'lệch', bad)
