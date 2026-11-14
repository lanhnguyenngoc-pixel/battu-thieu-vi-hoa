# -*- coding: utf-8 -*-
"""Giao diện bản 1.4.0: hai bảng nhập còn thiếu so với workbook 6.3.
  * BangViecTimGio  – các việc lớn đã xảy ra để tìm giờ sinh (ô nhập KHONG_RO_GIO C11:E22, danh sách loại việc D11:D22)
  * TrangDoiChieu   – ghi thực tế từng năm và đánh giá luận giải (ô nhập THEO_DOI N/O/P dòng 12–111)
Giao diện chỉ đặt ô nhập và hiện nguyên văn kết quả của lõi tính; không tự tính quy tắc."""
import re
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QWidget, QHBoxLayout, QVBoxLayout, QGridLayout, QPushButton, QComboBox, QTableWidget,
                               QTableWidgetItem, QLabel, QFrame, QHeaderView)
from battu_core.du_lieu import model
from battu_core.nhap_lieu import SO_VIEC_TIM_GIO, SO_NAM_THEO_DOI
from export.lam_sach import hien_thi_sach


def lbl(text='', name=None, wrap=False):
    l = QLabel(text)
    if name: l.setObjectName(name)
    l.setWordWrap(wrap); l.setTextInteractionFlags(Qt.TextSelectableByMouse)
    return l


def o_chi_doc(t, nen='#fafaff'):
    it = QTableWidgetItem(t); it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    it.setForeground(QColor('#1e2140')); it.setBackground(QColor(nen))
    if len(t) > 40: it.setToolTip(t[:1500])
    return it


# ---------------------------------------------------------------- VIỆC LỚN ĐỂ TÌM GIỜ SINH
class BangViecTimGio(QWidget):
    doi = Signal(object)

    def __init__(self):
        super().__init__()
        M = model(); self.loai = [''] + M.danh_sach_chon.get('KHONG_RO_GIO', {}).get('D11:D22', [])
        v = QVBoxLayout(self); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(8)
        h = QHBoxLayout()
        h.addWidget(lbl('Nhập năm dương lịch và loại việc lớn đã xảy ra. Phần mềm tính lá số cho cả 12 giờ và xếp giờ '
                        'khớp nhiều việc nhất lên đầu (xem thẻ "Bảng 12 giờ").', 'mo', True), 1)
        self.n_xoa = QPushButton('Xóa hết'); self.n_xoa.setObjectName('nut_phu')
        self.n_ap = QPushButton('✔  Áp dụng && tính lại'); self.n_ap.setObjectName('nut_chinh')
        h.addWidget(self.n_xoa); h.addWidget(self.n_ap); v.addLayout(h)
        self.tom = lbl('', 'chip_info', True); v.addWidget(self.tom)
        self.bang = QTableWidget(SO_VIEC_TIM_GIO, 5); self.bang.setWordWrap(True)
        self.bang.setVerticalHeaderLabels([str(i + 1) for i in range(SO_VIEC_TIM_GIO)])
        for j, w in enumerate((80, 280, 200, 150, 520)): self.bang.setColumnWidth(j, w)
        self.bang.horizontalHeader().setStretchLastSection(True)
        self.cb = []
        for i in range(SO_VIEC_TIM_GIO):
            self.bang.setItem(i, 0, QTableWidgetItem('')); self.bang.setItem(i, 2, QTableWidgetItem(''))
            c = QComboBox(); c.addItems(self.loai); c.setFixedHeight(30); self.bang.setCellWidget(i, 1, c); self.cb.append(c)
            self.bang.setRowHeight(i, 36)
        v.addWidget(self.bang, 1)
        self.loi = lbl('', 'chip_err', True); self.loi.setVisible(False); v.addWidget(self.loi)
        self.n_ap.clicked.connect(self.ap_dung); self.n_xoa.clicked.connect(self.xoa)

    def nhap(self):
        ds = []; loi = []
        for i in range(SO_VIEC_TIM_GIO):
            nam = self.bang.item(i, 0).text().strip(); loai = self.cb[i].currentText(); ghi = self.bang.item(i, 2).text().strip()
            if not (nam or loai or ghi): continue
            if nam and not re.fullmatch(r'\d{4}', nam): loi.append('dòng %d: năm "%s" phải là 4 chữ số' % (i + 1, nam)); continue
            ds.append([int(nam) if nam else None, loai or None, ghi or None])
        return ds, loi

    def ap_dung(self):
        ds, loi = self.nhap()
        self.loi.setText('Chưa áp dụng – ' + '; '.join(loi)); self.loi.setVisible(bool(loi))
        if not loi: self.doi.emit(ds)

    def xoa(self):
        for i in range(SO_VIEC_TIM_GIO):
            self.bang.item(i, 0).setText(''); self.bang.item(i, 2).setText(''); self.cb[i].setCurrentIndex(0)
        self.doi.emit([])

    def hien(self, kq):
        E = kq.E; o = lambda ref: hien_thi_sach(E, 'KHONG_RO_GIO', ref)
        self.bang.setHorizontalHeaderLabels([o(c + '10') for c in 'CDEFG'])
        self.tom.setText('%s: %s / %s' % (o('B23'), o('E23'), o('F23')))
        ds = list(kq.la_so.viec_tim_gio or [])
        for i in range(SO_VIEC_TIM_GIO):
            nam, loai, ghi = (list(ds[i]) + [None] * 3)[:3] if i < len(ds) else (None, None, None)
            self.bang.item(i, 0).setText('' if nam in (None, '') else str(int(nam)))
            self.cb[i].setCurrentText(loai or ''); self.bang.item(i, 2).setText(ghi or '')
            r = 11 + i
            kq_f = o('F%d' % r)
            self.bang.setItem(i, 3, o_chi_doc(kq_f, '#ecfdf5' if kq_f == 'Khớp' else '#fafaff'))
            self.bang.setItem(i, 4, o_chi_doc(o('G%d' % r)))


# ---------------------------------------------------------------- ĐỐI CHIẾU THỰC TẾ (THEO_DOI)
class TrangDoiChieu(QWidget):
    doi = Signal(object)
    COT_XEM = [('A', 64), ('B', 86), ('C', 96), ('D', 96), ('G', 190)]      # năm, tuổi, lưu niên, đại vận, xu thế đang dùng
    COT_KQ = [('Q', 84), ('R', 84)]
    TEN_NHAP = ['Thực tế năm đó', 'Sự kiện thực tế', 'Đánh giá luận giải']

    def __init__(self):
        super().__init__()
        M = model(); ds = M.danh_sach_chon.get('THEO_DOI', {})
        self.ds_n = [''] + ds.get('N12:N111', []); self.ds_p = [''] + ds.get('P12:P111', [])
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); h.addWidget(lbl('Đối chiếu thực tế', 'tieude')); h.addStretch()
        self.n_xoa = QPushButton('Xóa hết'); self.n_xoa.setObjectName('nut_phu')
        self.n_ap = QPushButton('✔  Áp dụng && tính lại'); self.n_ap.setObjectName('nut_chinh')
        h.addWidget(self.n_xoa); h.addWidget(self.n_ap); v.addLayout(h)
        v.addWidget(lbl('Với những năm đã qua, ghi năm đó thực tế tốt hay xấu và luận giải đúng đến đâu. Khi đủ từ 5 năm, '
                        'phần mềm cho biết phương pháp vượng suy nào (PP1 hay PP2) khớp với đời thực của lá số này hơn.', 'mo', True))
        the = QFrame(); the.setObjectName('the'); g = QGridLayout(the); g.setContentsMargins(16, 12, 16, 12)
        g.setHorizontalSpacing(16); g.setVerticalSpacing(4); self.tom = []
        for i in range(5):
            a = lbl('', 'nhan'); b = lbl('', None, True); g.addWidget(a, i, 0); g.addWidget(b, i, 1); self.tom.append((a, b))
        g.setColumnStretch(1, 1); v.addWidget(the)
        n = SO_NAM_THEO_DOI
        self.bang = QTableWidget(n, len(self.COT_XEM) + 3 + len(self.COT_KQ)); self.bang.setWordWrap(True)
        self.bang.verticalHeader().setVisible(False)
        w = [x for _, x in self.COT_XEM] + [150, 240, 160] + [x for _, x in self.COT_KQ]
        for j, x in enumerate(w): self.bang.setColumnWidth(j, x)
        k = len(self.COT_XEM); self.cb_n = []; self.cb_p = []
        for i in range(n):
            a = QComboBox(); a.addItems(self.ds_n); b = QComboBox(); b.addItems(self.ds_p)
            a.setFixedHeight(30); b.setFixedHeight(30)
            self.bang.setCellWidget(i, k, a); self.bang.setItem(i, k + 1, QTableWidgetItem('')); self.bang.setCellWidget(i, k + 2, b)
            self.cb_n.append(a); self.cb_p.append(b); self.bang.setRowHeight(i, 36)
        self.bang.horizontalHeader().setSectionResizeMode(len(self.COT_XEM) - 1, QHeaderView.Stretch)
        self.bang.setToolTip('Thực tế: Tốt / Bình thường / Xấu. Đánh giá: Đúng / Một phần / Sai / Chưa rõ.')
        v.addWidget(self.bang, 1)
        self.n_ap.clicked.connect(self.ap_dung); self.n_xoa.clicked.connect(self.xoa)
        self.nam = [None] * n; self.kq = None

    def nhap(self):
        k = len(self.COT_XEM); d = {}
        for i, nam in enumerate(self.nam):
            if not isinstance(nam, int): continue
            tt = self.cb_n[i].currentText(); sk = self.bang.item(i, k + 1).text().strip(); dg = self.cb_p[i].currentText()
            if tt or sk or dg: d[nam] = [tt or None, sk or None, dg or None]
        return d

    def ap_dung(self):
        self.doi.emit(self.nhap())

    def xoa(self):
        k = len(self.COT_XEM)
        for i in range(SO_NAM_THEO_DOI):
            self.cb_n[i].setCurrentIndex(0); self.cb_p[i].setCurrentIndex(0); self.bang.item(i, k + 1).setText('')
        self.doi.emit({})

    def hien(self, kq):
        self.kq = kq; E = kq.E; o = lambda ref: hien_thi_sach(E, 'THEO_DOI', ref)
        for i, (a, b) in enumerate(self.tom):
            a.setText(o('A%d' % (4 + i))); b.setText(o('E%d' % (4 + i)))
        k = len(self.COT_XEM)
        self.bang.setHorizontalHeaderLabels([o(c + '11') for c, _ in self.COT_XEM] + self.TEN_NHAP +
                                            [o(c + '11') for c, _ in self.COT_KQ])
        td = kq.la_so.theo_doi or {}
        for i in range(SO_NAM_THEO_DOI):
            r = 12 + i; nam = E.gia_tri('THEO_DOI', 'A%d' % r)
            nam = int(nam) if isinstance(nam, float) else None; self.nam[i] = nam
            self.bang.setRowHidden(i, nam is None)
            for j, (c, _) in enumerate(self.COT_XEM): self.bang.setItem(i, j, o_chi_doc(o('%s%d' % (c, r))))
            g = (list(td.get(nam) or []) + [None] * 3)[:3] if nam is not None else [None] * 3
            self.cb_n[i].setCurrentText(g[0] or ''); self.bang.item(i, k + 1).setText(g[1] or ''); self.cb_p[i].setCurrentText(g[2] or '')
            for j, (c, _) in enumerate(self.COT_KQ):
                x = E.gia_tri('THEO_DOI', '%s%d' % (c, r))
                t = '✔' if x == 1 else ('✗' if x == 0 else '')
                it = o_chi_doc(t, '#ecfdf5' if x == 1 else ('#fef2f2' if x == 0 else '#fafaff')); it.setTextAlignment(Qt.AlignCenter)
                self.bang.setItem(i, k + 3 + j, it)
