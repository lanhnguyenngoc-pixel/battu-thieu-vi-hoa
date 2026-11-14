# -*- coding: utf-8 -*-
"""Bảng hiển thị một sheet như workbook: mọi ô hằng + ô công thức (chữ theo định dạng ô). Rê chuột: địa chỉ Excel.
dong: danh sách dòng Excel cần hiện (mặc định toàn sheet). Định dạng có điều kiện lấy từ battu_core.dinh_dang_dk
(dịch nguyên văn từ workbook)."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView
from battu_core.du_lieu import model, ref_of
from battu_core.hien_thi import hien_thi_o
from battu_core.dinh_dang_dk import kieu_sheet


def kich_thuoc(sheet):
    M = model(); si = M.si[sheet]
    keys = list(M.consts[si]) + list(M.formula_cells[si])
    return max(r for r, c in keys), max(c for r, c in keys)


class BangSheet(QTableWidget):
    def __init__(self, sheet, max_cot=None, dong=None):
        super().__init__(); self.sheet = sheet
        nr, nc = kich_thuoc(sheet)
        if max_cot: nc = min(nc, max_cot)
        self.dong = list(dong) if dong else list(range(1, nr + 1))
        self.vi_tri = {r: i for i, r in enumerate(self.dong)}
        self.nr, self.nc = len(self.dong), nc
        self.setRowCount(self.nr); self.setColumnCount(nc)
        self.setHorizontalHeaderLabels([ref_of(1, c)[:-1] for c in range(1, nc + 1)])
        self.setVerticalHeaderLabels([str(r) for r in self.dong])
        self.setEditTriggers(QAbstractItemView.NoEditTriggers); self.setWordWrap(True)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.horizontalHeader().setDefaultSectionSize(120)
        bc = model().bo_cuc.get(sheet, {})
        for k, w in bc.get('rong', {}).items():
            k = int(k)
            if k <= nc: self.setColumnWidth(k - 1, int(min(w, 80) * 7.2 + 10))
        self.an_cot = {c for c in bc.get('an_cot', []) if c <= nc}          # cột/dòng ẩn như workbook (ô tính phụ)
        self.an_dong = {r for r in bc.get('an_dong', []) if r in self.vi_tri}
        for c in self.an_cot: self.setColumnHidden(c - 1, True)
        for r in self.an_dong: self.setRowHidden(self.vi_tri[r], True)
        from openpyxl.utils.cell import range_boundaries
        for g in bc.get('gop', []):
            c1, r1, c2, r2 = range_boundaries(g)
            if r1 in self.vi_tri and c1 <= nc and (r2 > r1 or c2 > c1):
                rr = [r for r in range(r1, r2 + 1) if r in self.vi_tri]
                self.setSpan(self.vi_tri[r1], c1 - 1, len(rr), min(c2, nc) - c1 + 1)

    def hien(self, E):
        M = model(); si = M.si[self.sheet]; fc = M.formula_cells[si]
        dk = kieu_sheet(E, self.sheet)
        self.setUpdatesEnabled(False)
        for i, r in enumerate(self.dong):
            for c in range(1, self.nc + 1):
                ref = ref_of(r, c)
                if (r, c) in fc:
                    t = hien_thi_o(E, self.sheet, ref)
                else:
                    v = E.consts[si].get((r, c))   # gồm cả ô người dùng đã nhập
                    if v is None:
                        if self.item(i, c - 1): self.setItem(i, c - 1, None)
                        continue
                    t = hien_thi_o(E, self.sheet, ref) if isinstance(v, (int, float)) else str(v)
                it = QTableWidgetItem(t)
                it.setToolTip('%s!%s%s%s' % (self.sheet, ref, '  (ô công thức)' if (r, c) in fc else '', ('\n' + t[:1500]) if len(t) > 80 else ''))
                if (r, c) in fc: it.setForeground(QColor('#1e2140'))
                else: it.setForeground(QColor('#4b5070'))
                if (r, c) not in fc and isinstance(t, str) and len(t) > 6 and t[:1].isupper() and (t.isupper() or t[:2] in ('A.', 'B.', 'C.', 'D.', 'E.', 'F.', 'G.', 'H.', 'I.', 'J.', 'K.')):
                    f = QFont(); f.setBold(True); it.setFont(f); it.setForeground(QColor('#4338ca')); it.setBackground(QColor('#f3f0ff'))
                k = dk.get((r, c))
                if k:                                    # định dạng có điều kiện của workbook
                    if k.get('dam'): f = QFont(it.font()); f.setBold(True); it.setFont(f)
                    if k.get('chu'): it.setForeground(QColor(k['chu']))
                    if k.get('nen'): it.setBackground(QColor(k['nen']))
                self.setItem(i, c - 1, it)
        self.setUpdatesEnabled(True)
        self.resizeRowsToContents()
        for i in range(self.nr):                 # dòng ghi chú rất dài: giới hạn chiều cao (rê chuột xem đủ)
            if self.rowHeight(i) > 140: self.setRowHeight(i, 140)

    def o_hien(self):
        """Các ô (r, c) đang hiện (bỏ cột/dòng ẩn của workbook)."""
        return [(r, c) for r in self.dong if r not in self.an_dong for c in range(1, self.nc + 1) if c not in self.an_cot]

    def toi_o(self, r, c):
        """Cuộn tới và chọn ô Excel (r, c) nếu đang hiện."""
        if r in self.vi_tri:
            it = self.item(self.vi_tri[r], c - 1)
            self.setCurrentCell(self.vi_tri[r], c - 1)
            if it: self.scrollToItem(it, QAbstractItemView.PositionAtCenter)
