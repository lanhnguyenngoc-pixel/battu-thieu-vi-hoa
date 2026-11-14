# -*- coding: utf-8 -*-
"""Nạp dữ liệu nền (quy tắc 5: không gõ cứng).
  * CSDL      ← data/nen/CSDL_toan_bo_gia_tri.json  (bản chép nguyên văn 04_DU_LIEU_NEN)
  * TIET_KHI  ← data/nen/TIET_KHI.csv (dòng 4–2439); dòng tiêu đề 1–3 lấy từ hang_so_sheet.json
  * Ô hằng các sheet khác ← data/hang_so_sheet.json (trích tự động từ workbook 6.3, MT-02)
  * Defined Names ← data/nen/defined_names_5.9.json (đã đối chiếu: trùng workbook 6.3)
"""
import os, json, csv, re, functools

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')


def rc(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref)
    c = 0
    for ch in m.group(1): c = c * 26 + ord(ch) - 64
    return int(m.group(2)), c


def ref_of(r, c):
    s = ''
    while c: c, k = divmod(c - 1, 26); s = chr(65 + k) + s
    return '%s%d' % (s, r)


class Model:
    """Dữ liệu tĩnh của workbook: ô hằng, vị trí ô công thức, tên."""
    def __init__(self):
        hs = json.load(open(os.path.join(DATA, 'hang_so_sheet.json'), encoding='utf-8'))
        self.nguon_sha256 = hs['sha256']
        self.sheets = hs['sheets']
        self.si = {s: i for i, s in enumerate(self.sheets)}
        self.consts = [dict() for _ in self.sheets]
        for s, d in hs['hang_so'].items():
            k = self.consts[self.si[s]]
            for ref, v in d.items():
                if isinstance(v, dict): continue  # ô lỗi hằng (không có trong 6.3)
                k[rc(ref)] = v
        # CSDL từ 04_DU_LIEU_NEN
        grid = json.load(open(os.path.join(DATA, 'nen', 'CSDL_toan_bo_gia_tri.json'), encoding='utf-8'))
        k = self.consts[self.si['CSDL']]
        fcsdl = set(rc(x) for x in hs['o_cong_thuc']['CSDL'])
        for i, row in enumerate(grid):
            for j, v in enumerate(row):
                if v is None or (i + 1, j + 1) in fcsdl: continue
                if isinstance(v, bool): pass
                elif isinstance(v, (int, float)): v = float(v)
                k[(i + 1, j + 1)] = v
        # TIET_KHI từ 04_DU_LIEU_NEN
        k = self.consts[self.si['TIET_KHI']]
        with open(os.path.join(DATA, 'nen', 'TIET_KHI.csv'), encoding='utf-8-sig') as f:
            rd = csv.reader(f); next(rd)
            for i, row in enumerate(rd):
                r = 4 + i
                k[(r, 1)] = float(row[0]); k[(r, 2)] = row[1]; k[(r, 3)] = row[2]
                k[(r, 4)] = float(int(row[3])); k[(r, 5)] = float(int(row[4]))
        self.formula_cells = [set(rc(x) for x in hs['o_cong_thuc'].get(s, [])) for s in self.sheets]
        self.data_tables = hs.get('data_tables', {})
        self.dinh_dang = hs.get('dinh_dang_o_cong_thuc', {})
        self.danh_sach_chon = hs.get('danh_sach_chon', {})   # {sheet: {sqref: [lựa chọn]}} – Data validation của workbook
        self.bo_cuc = hs.get('bo_cuc', {})   # ô gộp, độ rộng cột (chỉ để hiển thị)
        self.names = json.load(open(os.path.join(DATA, 'nen', 'defined_names_5.9.json'), encoding='utf-8'))


@functools.lru_cache(maxsize=1)
def model():
    return Model()
