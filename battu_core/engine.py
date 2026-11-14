# -*- coding: utf-8 -*-
"""Lõi tính: nhận ô nhập → tính các ô công thức theo nhu cầu (có nhớ đệm), trả {sheet: {ô: giá trị}}.

Thứ tự tính theo đúng phụ thuộc của công thức (đồ thị ô không có vòng – kiểm ở GĐ0), nên kết quả
không phụ thuộc thứ tự gọi. Ô công thức CHƯA cài ở giai đoạn hiện tại:
  * khi kiểm thử: lấy giá trị "mượn" từ đáp án (dap_an_o) – xem tests/
  * khi chạy giao diện: coi như trống (None) và ghi vào self.chua_cai
"""
import sys, importlib
from battu_core import xl as X
from battu_core.du_lieu import model, rc, ref_of
from battu_core import sheets as _sheets

sys.setrecursionlimit(50000)

_CHUA_TINH = object()


def _cung(a, b):
    if type(a) is X.ErrVal or type(b) is X.ErrVal:
        return type(a) is type(b) and a.code == b.code
    return type(a) is type(b) and a == b


def _load_sheet_modules():
    M = model(); mods = {}
    for sh, mod in _sheets.MODULES.items():
        m = importlib.import_module('battu_core.sheets.' + mod)
        mods[M.si[sh]] = {rc(ref): (fn, r, c) for ref, (fn, r, c) in m.CELLS.items()}
    return mods

_MODS = None


def cell_modules():
    global _MODS
    if _MODS is None: _MODS = _load_sheet_modules()
    return _MODS


class Engine:
    """Bảng dữ liệu (Data Table) của Excel – 12 bảng ở KHONG_RO_GIO – được thay bằng VÒNG 12 GIỜ (MT-13):
    ô (r, c) trong vùng bảng = công thức đầu cột c (dòng trên vùng), tính trong một lõi tính phụ có ô vào
    `r1` (KHONG_RO_GIO!C6) = giá trị ở cột bên trái dòng r (B118… = Tí…Hợi). Mỗi giá trị vào một lõi tính phụ,
    có nhớ đệm; chỉ tính các ô công thức đầu bảng thực sự cần."""
    def __init__(self, nhap=None, muon=None, _la_phu=False, _cha=None):
        """nhap: {sheet: {ô: giá trị}} ghi đè ô hằng (ô nhập). Giá trị '' hoặc None = xóa ô.
        muon: {sheet: {ô: giá trị}} cho ô công thức chưa cài (chỉ dùng khi kiểm thử)."""
        self.M = model(); self.F = cell_modules()
        self.consts = [dict(d) for d in self.M.consts]
        for sh, d in (nhap or {}).items():
            k = self.consts[self.M.si[sh]]
            for ref, v in d.items():
                key = rc(ref)
                if v is None or v == '': k.pop(key, None)
                else: k[key] = float(v) if isinstance(v, int) and not isinstance(v, bool) else v
        self.muon = {}
        for sh, d in (muon or {}).items():
            si = self.M.si[sh]
            for ref, v in d.items(): self.muon[(si, rc(ref))] = v
        self.cache = {}
        self.chua_cai = set()
        self._nhap = nhap or {}; self._muon = muon or {}; self._phu = {}
        self.dt = {} if _la_phu else self._bang_du_lieu()
        self._cha = _cha; self._giong_cha = False
        if _cha is not None:
            # Tối ưu (không đổi kết quả): ô KHÔNG phụ thuộc ô vào lấy từ lõi chính; nếu mọi ô tham chiếu trực tiếp
            # ô vào cho cùng giá trị với lõi chính thì toàn bộ lõi phụ trùng lõi chính.
            from battu_core.sheets import _phu_thuoc as P
            self._pt = P.PHU_THUOC
            self._giong_cha = all(_cung(self.raw(*k), _cha.raw(*k)) for k in sorted(P.TRUC_TIEP) if k[0] in self.F)

    def _bang_du_lieu(self):
        from openpyxl.utils.cell import range_boundaries
        out = {}
        for sh, lst in self.M.data_tables.items():
            si = self.M.si[sh]
            for ref, at in lst:
                c1, r1, c2, r2 = range_boundaries(ref)
                if at.get('dt2D') == '1' or at.get('dtr') == '1':
                    raise NotImplementedError('Bảng dữ liệu 2 chiều / theo dòng chưa có trong workbook 6.3')
                vao = rc(at['r1'])
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        out[(si, r, c)] = (r1 - 1, c, (r, c1 - 1), vao)   # (dòng công thức, cột, ô giá trị vào, ô vào)
        return out

    def _tinh_bang(self, si, r, c):
        top, col, (rv, cv), (ri, ci) = self.dt[(si, r, c)]
        gia_tri_vao = self.raw(si, rv, cv)
        key = (si, ri, ci, gia_tri_vao if not isinstance(gia_tri_vao, float) else ('so', gia_tri_vao))
        E2 = self._phu.get(key)
        if E2 is None:
            nhap = {k: dict(v) for k, v in self._nhap.items()}
            nhap.setdefault(self.M.sheets[si], {})[ref_of(ri, ci)] = gia_tri_vao
            E2 = Engine(nhap, self._muon, _la_phu=True, _cha=self); self._phu[key] = E2
        v = E2.raw(si, top, col)
        self.chua_cai |= E2.chua_cai
        return v

    # ---- truy cập ô
    def raw(self, si, r, c):
        key = (si, r, c)
        v = self.cache.get(key, _CHUA_TINH)
        if v is not _CHUA_TINH: return v
        if self._cha is not None and (self._giong_cha or key not in self._pt) and (r, c) in self.M.formula_cells[si]:
            v = self._cha.raw(si, r, c); self.cache[key] = v; return v
        if key in self.dt:
            v = self._tinh_bang(si, r, c); self.cache[key] = v; return v
        fs = self.F.get(si)
        ent = fs.get((r, c)) if fs else None
        if ent is not None:
            fn, r0, c0 = ent
            try:
                v = X.ket_qua(fn(self, r0, c0))
            except X.XLError as e:
                v = X.ErrVal(e.code)
        elif (r, c) in self.M.formula_cells[si]:
            v = self.muon.get((si, (r, c)), _CHUA_TINH)
            if v is _CHUA_TINH:
                self.chua_cai.add((si, r, c)); v = None
        else:
            v = self.consts[si].get((r, c))
        self.cache[key] = v
        return v

    def v(self, si, r, c):
        v = self.raw(si, r, c)
        if type(v) is X.ErrVal: raise X.XLError(v.code)
        return v

    def rg(self, si, r1, c1, r2, c2):
        return X.Range(self, si, r1, c1, r2, c2)

    # ---- tiện ích
    def gia_tri(self, sheet, ref):
        r, c = rc(ref); return self.raw(self.M.si[sheet], r, c)

    def tinh_sheet(self, sheet):
        """Tính mọi ô công thức đã cài của một sheet → {ô: giá trị}."""
        si = self.M.si[sheet]; out = {}
        o = set(self.F.get(si, {})) | {(r, c) for (s_, r, c) in self.dt if s_ == si}
        for (r, c) in sorted(o):
            v = self.raw(si, r, c)
            if type(v) is bool: v = 1.0 if v else 0.0          # như getDataArray
            out[ref_of(r, c)] = v.code if isinstance(v, X.ErrVal) else v
        return out

    def tinh_tat_ca(self):
        return {self.M.sheets[si]: self.tinh_sheet(self.M.sheets[si]) for si in sorted(self.F)}
