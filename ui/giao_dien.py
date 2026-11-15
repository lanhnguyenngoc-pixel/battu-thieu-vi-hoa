# -*- coding: utf-8 -*-
"""Giao diện GĐ1 (bản thô): màn hình NHẬP LIỆU + thẻ TRUNG_GIAN.
Mọi giá trị hiển thị lấy từ lõi tính (battu_core) – giao diện không tự tính quy tắc."""
import os, sys, datetime, threading
from PySide6.QtCore import Qt, QTimer, Signal, QObject
from PySide6.QtGui import QFont, QFontDatabase, QColor, QIntValidator, QDoubleValidator
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton, QHBoxLayout,
                               QVBoxLayout, QGridLayout, QStackedWidget, QLineEdit, QSpinBox, QComboBox,
                               QCheckBox, QButtonGroup, QScrollArea, QTableWidget, QTableWidgetItem,
                               QHeaderView, QSizePolicy, QGraphicsDropShadowEffect, QAbstractItemView)

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.nhap_lieu import LaSo, CAN, CHI, VUONG_SUY, CHE_DO_NHT, LoiNhap
from battu_core.dich_vu import tinh, danh_sach_la_mau, KetQua
from battu_core import lich_am

PHIEN_BAN = 'Phiên bản 1.7.1 · Quy tắc 6.5'
# Màu ngũ hành (chỉ để hiển thị): 0 Mộc, 1 Hỏa, 2 Thổ, 3 Kim, 4 Thủy – theo TRUNG_GIAN G7
MAU_HANH = {0: '#16a34a', 1: '#e11d48', 2: '#a0522d', 3: '#b8860b', 4: '#2563eb'}
TEN_HANH = {0: 'Mộc', 1: 'Hỏa', 2: 'Thổ', 3: 'Kim', 4: 'Thủy'}

QSS = """
* { font-family: 'Inter'; color: #1e2140; }
QMainWindow, #nen { background: #f4f5fb; }
#sidebar { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4338ca, stop:1 #7c3aed); }
#logo { color: white; font-size: 22px; font-weight: 800; }
#logo_phu { color: rgba(255,255,255,0.75); font-size: 12px; }
#nav { background: transparent; color: rgba(255,255,255,0.85); text-align: left; padding: 11px 16px;
       border: none; border-radius: 12px; font-size: 14px; font-weight: 500; }
#nav:hover { background: rgba(255,255,255,0.12); }
#nav:checked { background: white; color: #4338ca; font-weight: 700; }
#nav:disabled { color: rgba(255,255,255,0.38); }
#cg { color: white; font-size: 12px; font-weight: 600; }
#cg::indicator:checked { background: #fbbf24; }
#phienban { color: rgba(255,255,255,0.7); font-size: 11px; }
#the { background: white; border: 1px solid #e7e8f3; border-radius: 20px; }
#tieude { font-size: 22px; font-weight: 800; }
#tieude_the { font-size: 16px; font-weight: 700; }
#mo { color: #7b809c; font-size: 12px; }
#nhan { color: #5b6080; font-size: 12px; font-weight: 600; }
#seg { background: #eef0f8; border-radius: 12px; }
#segbtn { background: transparent; border: none; border-radius: 10px; padding: 8px 14px; font-weight: 600; color: #5b6080; }
#segbtn:checked { background: white; color: #4f46e5; }
QLineEdit, QSpinBox, QComboBox { background: #f7f8fc; border: 1.5px solid #e3e5f0; border-radius: 10px;
       padding: 6px 10px; font-size: 14px; min-height: 26px; }
QLineEdit:focus, QSpinBox:focus, QComboBox:focus { border-color: #7c6cff; background: white; }
QComboBox::drop-down { border: none; width: 22px; }
QSpinBox::up-button, QSpinBox::down-button { width: 0; border: none; }
QCheckBox { font-size: 13px; font-weight: 600; spacing: 8px; }
QCheckBox::indicator { width: 38px; height: 22px; border-radius: 11px; background: #d9dbe9; }
QCheckBox::indicator:checked { background: #7c3aed; }
#nut_chinh { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #4f46e5, stop:1 #9333ea); color: white;
       border: none; border-radius: 12px; padding: 11px 20px; font-weight: 700; font-size: 14px; }
#nut_phu { background: #eef0f8; color: #4f46e5; border: none; border-radius: 10px; padding: 8px 14px; font-weight: 600; }
#tru { background: #fafaff; border: 1px solid #ecebfb; border-radius: 16px; }
#tru_nhat { background: #f3f0ff; border: 1.5px solid #c4b5fd; border-radius: 16px; }
#tru_ten { color: #7b809c; font-size: 12px; font-weight: 700; letter-spacing: 1px; }
#than { color: #6d5dfc; font-size: 12px; font-weight: 600; }
#chip_ok { background: #ecfdf3; color: #166534; border-radius: 12px; padding: 9px 12px; font-size: 13px; }
#chip_warn { background: #fff7e6; color: #92400e; border-radius: 12px; padding: 9px 12px; font-size: 13px; }
#chip_info { background: #f1f3fb; color: #3b4064; border-radius: 12px; padding: 9px 12px; font-size: 13px; }
#chip_err { background: #fdecef; color: #9f1239; border-radius: 12px; padding: 9px 12px; font-size: 13px; }
#badge { background: #eef0ff; color: #4f46e5; border-radius: 9px; padding: 3px 9px; font-size: 11px; font-weight: 700; }
QTableWidget { background: white; border: none; gridline-color: #eef0f6; font-size: 13px; }
QHeaderView::section { background: #f4f5fb; border: none; padding: 8px; font-weight: 700; color: #4b5070; }
QScrollArea { border: none; background: transparent; }
"""


def nap_font():
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')
    for f in os.listdir(d):
        if f.endswith('.otf'): QFontDatabase.addApplicationFont(os.path.join(d, f))


def bong(w, r=28, a=28):
    e = QGraphicsDropShadowEffect(w); e.setBlurRadius(r); e.setOffset(0, 6); e.setColor(QColor(60, 50, 160, a))
    w.setGraphicsEffect(e)


def lbl(text, name=None, wrap=False):
    l = QLabel(text)
    if name: l.setObjectName(name)
    if wrap: l.setWordWrap(True)
    return l


def so_hien(v):
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else ('%.6g' % v)
    return '' if v is None else str(v)


class Seg(QFrame):
    """Nút chọn dạng phân đoạn."""
    doi = Signal(int)
    def __init__(self, nhan):
        super().__init__(); self.setObjectName('seg')
        h = QHBoxLayout(self); h.setContentsMargins(4, 4, 4, 4); h.setSpacing(4)
        self.g = QButtonGroup(self); self.g.setExclusive(True)
        for i, t in enumerate(nhan):
            b = QPushButton(t); b.setObjectName('segbtn'); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor)
            self.g.addButton(b, i); h.addWidget(b)
        self.g.button(0).setChecked(True)
        self.g.idClicked.connect(self.doi.emit)
    def chon(self): return self.g.checkedId()
    def dat(self, i): self.g.button(i).setChecked(True)


def o_so(min_, max_, val=None, w=None):
    s = QSpinBox(); s.setRange(min_, max_)
    if val is not None: s.setValue(val)
    if w: s.setFixedWidth(w)
    s.setAlignment(Qt.AlignCenter)
    return s


def o_tuy_chon(ph='trống', so_thuc=False):
    e = QLineEdit(); e.setPlaceholderText(ph)
    e.setValidator(QDoubleValidator(-1e9, 1e9, 6) if so_thuc else QIntValidator(-100000, 100000))
    return e


def doc_tuy_chon(e, so_thuc=False):
    t = e.text().strip().replace(',', '.')
    if not t: return None
    try: return float(t) if so_thuc else int(t)
    except ValueError: return None


class TruWidget(QFrame):
    def __init__(self, ten, nhat=False):
        super().__init__(); self.setObjectName('tru_nhat' if nhat else 'tru')
        v = QVBoxLayout(self); v.setContentsMargins(14, 14, 14, 14); v.setSpacing(4)
        self.ten = lbl(ten.upper(), 'tru_ten'); self.ten.setAlignment(Qt.AlignCenter)
        self.than = lbl('', 'than'); self.than.setAlignment(Qt.AlignCenter)
        self.can = QLabel(); self.can.setAlignment(Qt.AlignCenter)
        self.chi = QLabel(); self.chi.setAlignment(Qt.AlignCenter)
        self.hanh = lbl('', 'mo'); self.hanh.setAlignment(Qt.AlignCenter)
        self.tang = QLabel(); self.tang.setAlignment(Qt.AlignCenter); self.tang.setWordWrap(True)
        self.ts = lbl('', 'mo'); self.ts.setAlignment(Qt.AlignCenter)
        for w in (self.ten, self.than, self.can, self.chi, self.hanh, self.tang, self.ts): v.addWidget(w)

    def dat(self, can, chi, hc, hchi, than, tang, ts):
        def big(t, h):
            col = MAU_HANH.get(h, '#1e2140')
            return "<span style='font-size:30px;font-weight:800;color:%s'>%s</span>" % (col, t or '–')
        self.can.setText(big(can, hc)); self.chi.setText(big(chi, hchi))
        self.than.setText(than or '')
        self.hanh.setText('%s · %s' % (TEN_HANH.get(hc, '?'), TEN_HANH.get(hchi, '?')) if can else '')
        self.tang.setText(' '.join("<span style='color:%s;font-weight:600;font-size:12px'>%s</span>" % (MAU_HANH.get(h, '#555'), t)
                                   for t, h in tang))
        self.ts.setText(ts or '')


class _BaoXong(QObject):
    xong = Signal(object, object)


class TrangNhap(QWidget):
    da_tinh = Signal(object)

    def __init__(self):
        super().__init__()
        self.la_mau = danh_sach_la_mau()
        root = QHBoxLayout(self); root.setContentsMargins(28, 24, 28, 24); root.setSpacing(22)
        # ---------------- cột nhập
        cot = QFrame(); cot.setObjectName('the'); bong(cot); cot.setFixedWidth(470)
        cv = QVBoxLayout(cot); cv.setContentsMargins(24, 22, 24, 22); cv.setSpacing(12)
        cv.addWidget(lbl('Lá số mới', 'tieude'))
        cv.addWidget(lbl('Nhập ngày giờ sinh hoặc tứ trụ rồi bấm Luận giải (hoặc phím Enter).', 'mo', True))
        self.ho_ten = QLineEdit(); self.ho_ten.setPlaceholderText('Họ tên (không bắt buộc)')
        cv.addWidget(self.ho_ten)
        self.cach = Seg(['Dương lịch', 'Âm lịch', 'Tự nhập can chi']); cv.addWidget(self.cach)
        self.trang_nhap = []
        # dương lịch
        p0 = QWidget(); g0 = QGridLayout(p0); g0.setContentsMargins(0, 0, 0, 0); g0.setHorizontalSpacing(8)
        self.d_ngay = o_so(1, 31, 15); self.d_thang = o_so(1, 12, 3); self.d_nam = o_so(1899, 2101, 1985)
        for i, (t, w) in enumerate([('Ngày', self.d_ngay), ('Tháng', self.d_thang), ('Năm', self.d_nam)]):
            g0.addWidget(lbl(t, 'nhan'), 0, i); g0.addWidget(w, 1, i)
        self.d_am = lbl('', 'mo'); g0.addWidget(self.d_am, 2, 0, 1, 2)
        self.nut_am = QPushButton('Điền theo âm lịch quy đổi'); self.nut_am.setObjectName('nut_phu')
        self.nut_am.setToolTip('Điền tháng/ngày âm lịch quy đổi – dùng cho phương hướng kết hôn'); g0.addWidget(self.nut_am, 2, 2)
        self.trang_nhap.append(p0); cv.addWidget(p0)
        # âm lịch
        p1 = QWidget(); g1 = QGridLayout(p1); g1.setContentsMargins(0, 0, 0, 0); g1.setHorizontalSpacing(8)
        self.a_ngay = o_so(1, 30, 1); self.a_thang = o_so(1, 12, 1); self.a_nam = o_so(1899, 2101, 1985)
        for i, (t, w) in enumerate([('Ngày âm', self.a_ngay), ('Tháng âm', self.a_thang), ('Năm âm', self.a_nam)]):
            g1.addWidget(lbl(t, 'nhan'), 0, i); g1.addWidget(w, 1, i)
        self.a_nhuan = QCheckBox('Tháng nhuận'); g1.addWidget(self.a_nhuan, 2, 0, 1, 2)
        self.a_duong = lbl('', 'mo'); g1.addWidget(self.a_duong, 3, 0, 1, 3)
        self.trang_nhap.append(p1); cv.addWidget(p1)
        # can chi
        p2 = QWidget(); g2 = QGridLayout(p2); g2.setContentsMargins(0, 0, 0, 0); g2.setHorizontalSpacing(8)
        self.cb_can = []; self.cb_chi = []
        for i, t in enumerate(['Năm', 'Tháng', 'Ngày', 'Giờ']):
            g2.addWidget(lbl(t, 'nhan'), 0, i)
            a = QComboBox(); a.addItems(CAN); b = QComboBox(); b.addItems(CHI)
            g2.addWidget(a, 1, i); g2.addWidget(b, 2, i); self.cb_can.append(a); self.cb_chi.append(b)
        g2.addWidget(lbl('Năm sinh dương lịch của trụ năm (để tự suy khởi vận)', 'nhan'), 3, 0, 1, 4)
        self.c30 = o_tuy_chon('vd 1963'); g2.addWidget(self.c30, 4, 0, 1, 2)
        self.trang_nhap.append(p2); cv.addWidget(p2)
        # giờ
        hg = QHBoxLayout(); hg.setSpacing(8)
        self.gio = o_so(0, 23, 10); self.phut = o_so(0, 59, 30)
        hg.addWidget(lbl('Giờ', 'nhan')); hg.addWidget(self.gio); hg.addWidget(lbl('Phút', 'nhan')); hg.addWidget(self.phut)
        hg.addWidget(lbl('UTC+', 'nhan')); self.mui = QComboBox(); self.mui.addItems([str(x) for x in range(-12, 15)])
        self.mui.setCurrentText('7'); hg.addWidget(self.mui)
        self.w_gio = QWidget(); self.w_gio.setLayout(hg); hg.setContentsMargins(0, 0, 0, 0); cv.addWidget(self.w_gio)
        hk = QHBoxLayout(); self.krg = QCheckBox('Không rõ giờ sinh'); hk.addWidget(self.krg); hk.addStretch()
        hk.addWidget(lbl('Giờ đang xem', 'nhan')); self.gio_xem = QComboBox(); self.gio_xem.addItems(CHI)
        self.gio_xem.setCurrentText('Ngọ'); hk.addWidget(self.gio_xem); cv.addLayout(hk)
        hgt = QHBoxLayout(); hgt.addWidget(lbl('Giới tính', 'nhan')); self.gt = Seg(['Nam', 'Nữ']); hgt.addWidget(self.gt)
        hgt.addSpacing(12); hgt.addWidget(lbl('Năm xem', 'nhan')); self.nam_xem = o_so(1900, 2200, datetime.date.today().year, 90)
        hgt.addWidget(self.nam_xem); cv.addLayout(hgt)
        # ngày sinh phối ngẫu (tùy chọn) – dùng cho phần hợp tuổi ở Hôn nhân
        cv.addWidget(lbl('Ngày sinh dương lịch của vợ/chồng (tùy chọn – có thể chỉ nhập năm)', 'nhan', True))
        hp = QHBoxLayout(); hp.setSpacing(6)
        self.pn = [o_tuy_chon(t) for t in ('ngày', 'tháng', 'năm', 'giờ', 'phút')]
        for w in self.pn: hp.addWidget(w)
        cv.addLayout(hp)
        # hoàn cảnh (tùy chọn) – chỉ dùng cho bản Luận giải tổng hợp (lọc mục không hợp), không đổi kết quả tính
        hhn = QHBoxLayout(); hhn.setSpacing(8)
        hhn.addWidget(lbl('Hôn nhân', 'nhan')); self.hon_nhan = QComboBox()
        self.hon_nhan.addItems(['(không nhập)', 'Độc thân', 'Đã kết hôn', 'Ly hôn', 'Góa']); hhn.addWidget(self.hon_nhan)
        hhn.addWidget(lbl('Số con', 'nhan')); self.so_con = o_tuy_chon('số'); hhn.addWidget(self.so_con)
        cv.addLayout(hhn)
        # nâng cao
        self.nut_nc = QPushButton('▸  Tùy chọn nâng cao'); self.nut_nc.setObjectName('nut_phu'); self.nut_nc.setCheckable(True)
        cv.addWidget(self.nut_nc)
        self.nc = QWidget(); gn = QGridLayout(self.nc); gn.setContentsMargins(0, 0, 0, 0); gn.setHorizontalSpacing(8)
        self.vs = QComboBox(); self.vs.addItems(VUONG_SUY); self.vs.setMinimumWidth(190)
        self.nht = QComboBox(); self.nht.addItems(CHE_DO_NHT)
        self.mau = QComboBox(); self.mau.addItems(self.la_mau)
        self.c11 = o_tuy_chon('tự tính', True); self.c28 = o_tuy_chon('trống', True); self.c29 = o_tuy_chon('trống', True)
        self.c32 = o_tuy_chon('1–12'); self.c33 = o_tuy_chon('1–30')
        rows = [('Vượng suy dùng cho kết luận', self.vs), ('Ngũ hành thiếu', self.nht), ('Lá số mẫu trong sách', self.mau),
                ('Số ngày tới tiết Lập', self.c11), ('Số ngày đến tiết lệnh', self.c28), ('Số giờ lẻ', self.c29),
                ('Tháng sinh âm', self.c32), ('Ngày sinh âm', self.c33)]
        for i, (t, w) in enumerate(rows):
            gn.addWidget(lbl(t, 'nhan'), i, 0); gn.addWidget(w, i, 1)
        self.nc.setVisible(False); cv.addWidget(self.nc)
        cv.addStretch()
        self.loi = lbl('', 'chip_err', True); self.loi.setVisible(False); cv.addWidget(self.loi)
        # nút Luận giải (cách làm phổ biến: sửa xong mới tính, tính ở nền – giao diện không đứng)
        self.trang_thai = lbl('', 'chip_warn', True); self.trang_thai.setVisible(False); cv.addWidget(self.trang_thai)
        self.nut_tinh = QPushButton('▶  Luận giải'); self.nut_tinh.setObjectName('nut_chinh'); self.nut_tinh.setCursor(Qt.PointingHandCursor)
        self.nut_tinh.setMinimumHeight(42); self.nut_tinh.setToolTip('Tính và luận giải lá số (phím tắt: Enter hoặc F5)')
        self.nut_tinh.clicked.connect(lambda: self.tinh()); cv.addWidget(self.nut_tinh)
        sc = QScrollArea(); sc.setWidgetResizable(True); sc.setWidget(cot); sc.setFixedWidth(490)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        root.addWidget(sc)
        # ---------------- cột kết quả
        kq = QFrame(); kq.setObjectName('the'); bong(kq)
        kv = QVBoxLayout(kq); kv.setContentsMargins(24, 22, 24, 22); kv.setSpacing(14)
        hh = QHBoxLayout(); hh.addWidget(lbl('Tứ trụ', 'tieude')); hh.addStretch(); self.thoi_gian = lbl('', 'badge'); self.thoi_gian.setVisible(False); hh.addWidget(self.thoi_gian)
        kv.addLayout(hh)
        self.dong_phu = lbl('', 'mo', True); kv.addWidget(self.dong_phu)
        ht = QHBoxLayout(); ht.setSpacing(12)
        self.tru = [TruWidget('Năm'), TruWidget('Tháng'), TruWidget('Ngày', True), TruWidget('Giờ')]
        for t in self.tru: ht.addWidget(t)
        kv.addLayout(ht)
        self.chip_box = QWidget(); self.chips = QVBoxLayout(self.chip_box); self.chips.setContentsMargins(0, 0, 0, 0)
        self.chips.setSpacing(8); kv.addWidget(self.chip_box)
        kv.addStretch()
        sc2 = QScrollArea(); sc2.setWidgetResizable(True); sc2.setWidget(kq)
        root.addWidget(sc2, 1)
        # tín hiệu
        # sửa ô nhập chỉ đánh dấu "đã thay đổi"; tính khi bấm Luận giải / Enter / F5
        self.hen = QTimer(self); self.hen.setSingleShot(True); self.hen.setInterval(120); self.hen.timeout.connect(self._da_doi)
        for w in [self.d_ngay, self.d_thang, self.d_nam, self.a_ngay, self.a_thang, self.a_nam, self.gio, self.phut, self.nam_xem]:
            w.valueChanged.connect(self.hen.start)
        for w in self.cb_can + self.cb_chi + [self.mui, self.gio_xem, self.vs, self.nht, self.mau]:
            w.currentIndexChanged.connect(self.hen.start)
        self.hon_nhan.currentIndexChanged.connect(self.hen.start)
        for w in [self.c30, self.c11, self.c28, self.c29, self.c32, self.c33, self.ho_ten, self.so_con] + self.pn:
            w.textChanged.connect(self.hen.start)
        for w in [self.krg, self.a_nhuan]: w.toggled.connect(self.hen.start)
        self.cach.doi.connect(self.doi_cach); self.gt.doi.connect(lambda _: self.hen.start())
        from PySide6.QtGui import QShortcut, QKeySequence
        for k in (Qt.Key_Return, Qt.Key_Enter):
            QShortcut(QKeySequence(k), self, lambda: self.tinh(), context=Qt.WidgetWithChildrenShortcut)
        self._ban = False; self._cho = False; self._nho = {}
        self._bao = _BaoXong(); self._bao.xong.connect(self._xong)
        self.nut_am.clicked.connect(self.dien_am)
        self.nut_nc.toggled.connect(lambda on: (self.nc.setVisible(on), self.nut_nc.setText(('▾' if on else '▸') + '  Tùy chọn nâng cao')))
        self.kq = None
        self.doi_cach(0)
        if os.environ.get('BATTU_DONG_BO'): self.tinh()
        else: QTimer.singleShot(0, lambda: self.tinh())        # mở cửa sổ trước, tính lá mặc định ở nền

    def dien_am(self):
        try:
            ad, am, ay, nh = lich_am.s2l(self.d_ngay.value(), self.d_thang.value(), self.d_nam.value())
        except Exception:
            return
        self.c32.setText(str(am)); self.c33.setText(str(ad)); self.nut_nc.setChecked(True)

    def doi_cach(self, i):
        for k, p in enumerate(self.trang_nhap): p.setVisible(k == i)
        self.w_gio.setVisible(i != 2); self.hen.start()

    def la_so(self):
        i = self.cach.chon()
        s = LaSo(ho_ten=self.ho_ten.text().strip(), cach_nhap=['duong', 'am', 'canchi'][i])
        s.ngay, s.thang, s.nam = self.d_ngay.value(), self.d_thang.value(), self.d_nam.value()
        s.am_ngay, s.am_thang, s.am_nam, s.am_nhuan = self.a_ngay.value(), self.a_thang.value(), self.a_nam.value(), self.a_nhuan.isChecked()
        s.gio, s.phut = self.gio.value(), self.phut.value()
        s.mui_gio = float(self.mui.currentText())
        if i == 2:
            s.can = [c.currentText() for c in self.cb_can]; s.chi = [c.currentText() for c in self.cb_chi]
            s.nam_sinh_tru_nam = doc_tuy_chon(self.c30)
        s.gioi_tinh = 'Nam' if self.gt.chon() == 0 else 'Nữ'
        s.khong_ro_gio = self.krg.isChecked(); s.gio_dang_xem = self.gio_xem.currentText()
        s.nam_xem = self.nam_xem.value(); s.vuong_suy = self.vs.currentText(); s.che_do_nht = self.nht.currentText()
        s.la_mau = self.mau.currentText()
        s.so_ngay_toi_tiet_lap = doc_tuy_chon(self.c11, True); s.so_ngay_den_tiet = doc_tuy_chon(self.c28, True)
        s.so_gio_le = doc_tuy_chon(self.c29, True); s.thang_am = doc_tuy_chon(self.c32); s.ngay_am = doc_tuy_chon(self.c33)
        s.phoi_ngau = self.doc_phoi_ngau()
        s.tinh_trang_hon_nhan = '' if self.hon_nhan.currentIndex() == 0 else self.hon_nhan.currentText()
        s.so_con = doc_tuy_chon(self.so_con)
        if getattr(self, '_ghi_chu', None): s.ghi_chu = self._ghi_chu
        s.su_kien = [list(x) for x in getattr(self, '_su_kien', [])]
        s.viec_tim_gio = [list(x) for x in getattr(self, '_viec_tim_gio', [])]
        s.theo_doi = {int(k): list(v) for k, v in getattr(self, '_theo_doi', {}).items()}
        return s

    def doc_phoi_ngau(self):
        """Chỉ có năm → NHAP_LIEU C34 = năm (như hướng dẫn của workbook); có ngày/tháng → C34:G34 = ngày|tháng|năm|giờ|phút."""
        d, m, y, h, mi = [doc_tuy_chon(w) for w in self.pn]
        if d is None and m is None and h is None and mi is None:
            return [y, None, None, None, None]
        return [d, m, y, h, mi]

    def dat_la_so(self, s):
        """Nạp lá số đã lưu vào ô nhập rồi tính lại (mở lại từ Danh sách lá số)."""
        ws = [self.d_ngay, self.d_thang, self.d_nam, self.a_ngay, self.a_thang, self.a_nam, self.gio, self.phut, self.nam_xem,
              self.mui, self.gio_xem, self.vs, self.nht, self.mau, self.krg, self.a_nhuan, self.ho_ten,
              self.c30, self.c11, self.c28, self.c29, self.c32, self.c33, self.hon_nhan, self.so_con] + self.cb_can + self.cb_chi + self.pn
        for w in ws: w.blockSignals(True)
        try:
            self.ho_ten.setText(s.ho_ten or '')
            i = ['duong', 'am', 'canchi'].index(s.cach_nhap); self.cach.dat(i)
            for w, x in ((self.d_ngay, s.ngay), (self.d_thang, s.thang), (self.d_nam, s.nam), (self.a_ngay, s.am_ngay),
                         (self.a_thang, s.am_thang), (self.a_nam, s.am_nam), (self.gio, s.gio), (self.phut, s.phut), (self.nam_xem, s.nam_xem)):
                if x is not None: w.setValue(int(x))
            self.a_nhuan.setChecked(bool(s.am_nhuan))
            for j in range(4):
                if s.can[j]: self.cb_can[j].setCurrentText(s.can[j])
                if s.chi[j]: self.cb_chi[j].setCurrentText(s.chi[j])
            self.mui.setCurrentText(str(int(s.mui_gio)) if float(s.mui_gio).is_integer() else str(s.mui_gio))
            self.krg.setChecked(bool(s.khong_ro_gio)); self.gio_xem.setCurrentText(s.gio_dang_xem)
            self.gt.dat(0 if s.gioi_tinh == 'Nam' else 1)
            self.vs.setCurrentText(s.vuong_suy); self.nht.setCurrentText(s.che_do_nht); self.mau.setCurrentText(s.la_mau)
            so = lambda x: '' if x is None else ('%g' % x)
            self.c30.setText(so(s.nam_sinh_tru_nam)); self.c11.setText(so(s.so_ngay_toi_tiet_lap)); self.c28.setText(so(s.so_ngay_den_tiet))
            self.c29.setText(so(s.so_gio_le)); self.c32.setText(so(s.thang_am)); self.c33.setText(so(s.ngay_am))
            pn = list(s.phoi_ngau or []) + [None] * 5
            if pn[0] is not None and all(x is None for x in pn[1:5]) and pn[0] >= 1900:   # chỉ nhập năm
                pn = [None, None, pn[0], None, None]
            for w, x in zip(self.pn, pn[:5]): w.setText(so(x))
            hn = getattr(s, 'tinh_trang_hon_nhan', '') or ''
            self.hon_nhan.setCurrentText(hn if hn else '(không nhập)'); self.so_con.setText(so(getattr(s, 'so_con', None)))
        finally:
            for w in ws: w.blockSignals(False)
        for k, p in enumerate(self.trang_nhap): p.setVisible(k == i)
        self.w_gio.setVisible(i != 2)
        self._ghi_chu = s.ghi_chu; self._su_kien = [list(x) for x in (s.su_kien or [])]
        self._viec_tim_gio = [list(x) for x in (s.viec_tim_gio or [])]; self._theo_doi = dict(s.theo_doi or {})
        self.hen.stop(); self.tinh()

    def chip(self, text, kind=None):
        text = lam_sach(text) if isinstance(text, str) else text
        if not text: return
        if kind is None:
            kind = 'chip_ok' if text.startswith('✔') else 'chip_warn' if (text.startswith('⚠') or 'CẢNH BÁO' in text or 'KHÔNG KHỚP' in text) \
                else 'chip_err' if text.startswith('LỖI') else 'chip_info'
        self.chips.addWidget(lbl(text, kind, True))

    def _da_doi(self):
        """Ô nhập vừa đổi: báo kết quả đang hiện là của thông tin cũ."""
        if self.kq is None and not self._ban: return
        self.trang_thai.setText('Thông tin đã thay đổi – bấm Luận giải (hoặc Enter) để cập nhật kết quả.'); self.trang_thai.setVisible(True)

    def tinh(self, dong_bo=None):
        """Tính lá số. Mặc định tính ở luồng nền (giao diện vẫn dùng được); lá đã tính được nhớ lại."""
        self.hen.stop()
        s = self.la_so()
        # hiển thị ngày quy đổi
        try:
            d, m, y = self.d_ngay.value(), self.d_thang.value(), self.d_nam.value()
            datetime.date(y, m, d)
            ad, am, ay, nh = lich_am.s2l(d, m, y)
            self.d_am.setText('Âm lịch: %d/%d%s/%d' % (ad, am, ' nhuận' if nh else '', ay))
        except ValueError:
            self.d_am.setText('Ngày dương lịch không hợp lệ')
        try:
            if s.cach_nhap == 'am':
                dd, mm, yy = s.ngay_duong(); self.a_duong.setText('Dương lịch: %02d/%02d/%d' % (dd, mm, yy))
            khoa = repr(sorted((k, repr(v)) for k, v in s.o_nhap().items()))
        except LoiNhap as e:
            self.loi.setText('⚠ ' + str(e)); self.loi.setVisible(True); return
        if khoa in self._nho:                                   # lá đã tính: dùng lại
            import copy
            c = self._nho[khoa]; ls = copy.copy(c.la_so); ls.ho_ten, ls.ghi_chu = s.ho_ten, s.ghi_chu
            ls.tinh_trang_hon_nhan, ls.so_con = s.tinh_trang_hon_nhan, s.so_con      # không phải ô workbook – lấy theo ô nhập
            self._xong(KetQua(ls, c.E, c.ms), None); return
        if self._ban: self._cho = True; return                  # đang tính: tính lại lần nữa khi xong
        dong_bo = bool(os.environ.get('BATTU_DONG_BO')) if dong_bo is None else dong_bo
        self._ban = True; self.nut_tinh.setEnabled(False); self.nut_tinh.setText('⏳  Đang tính…')
        self.trang_thai.setVisible(False); QApplication.setOverrideCursor(Qt.BusyCursor)

        def chay():
            try: self._bao.xong.emit((khoa, tinh(s)), None)
            except Exception as e: self._bao.xong.emit(None, e)
        if dong_bo: chay()
        else: threading.Thread(target=chay, daemon=True).start()

    def _xong(self, goi, loi):
        if self._ban:
            self._ban = False; QApplication.restoreOverrideCursor()
            self.nut_tinh.setEnabled(True); self.nut_tinh.setText('▶  Luận giải')
        if self._cho:
            self._cho = False; self.tinh(); return
        if loi is not None:
            self.loi.setText('⚠ ' + str(loi)); self.loi.setVisible(True); return
        if isinstance(goi, tuple):
            khoa, kq = goi; self._nho[khoa] = kq
            while len(self._nho) > 6: self._nho.pop(next(iter(self._nho)))
        else: kq = goi
        s = kq.la_so; self.trang_thai.setVisible(False)
        self.loi.setVisible(False)
        self.kq = kq
        o = kq.o
        self.thoi_gian.setText('%.0f ms' % kq.ms)
        ten = (s.ho_ten + ' · ') if s.ho_ten else ''
        self.dong_phu.setText(ten + s.gioi_tinh)
        cols = 'CDEF'
        for i, t in enumerate(self.tru):
            c = cols[i]
            def ti(r): v = o('TRUNG_GIAN', '%s%d' % (c, r)); return int(v) if isinstance(v, float) else None
            tang = []
            for rr, rh in ((15, 21), (17, 22), (19, 23)):
                idx = ti(rr)
                if idx is not None and idx >= 0: tang.append((CAN[idx], ti(rh)))
            t.dat(o('NHAP_LIEU', c + '16'), o('NHAP_LIEU', c + '17'), ti(7), ti(9), o('TRUNG_GIAN', c + '12'),
                  tang, o('TRUNG_GIAN', c + '29'))
        while self.chips.count():
            w = self.chips.takeAt(0).widget()
            if w: w.hide(); w.setParent(None)
        self.chip(o('NHAP_LIEU', 'G7'))
        if s.cach_nhap != 'canchi':
            self.chip('Tiết: %s → %s · năm can chi (theo Lập xuân): %s' % (o('DAI_VAN', 'C12'), o('DAI_VAN', 'C13'), so_hien(o('DAI_VAN', 'C14'))), 'chip_info')
        self.chip(o('DAI_VAN', 'C16'))
        kv = o('DAI_VAN', 'C21'); nd = o('DAI_VAN', 'C17')
        self.chip('Khởi vận %s tuổi (%s; số ngày đến tiết %s).' % (
            so_hien(kv), 'thuận' if o('DAI_VAN', 'C7') == 1 else 'nghịch', ('%.2f' % nd) if isinstance(nd, float) else 'không có'), 'chip_info')
        if s.cach_nhap == 'canchi': self.chip(o('NHAP_LIEU', 'D45'))
        self.da_tinh.emit(kq)


class TrangTrungGian(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        the = QFrame(); the.setObjectName('the'); bong(the); tv = QVBoxLayout(the); tv.setContentsMargins(20, 18, 20, 18)
        h = QHBoxLayout(); h.addWidget(lbl('Trung gian', 'tieude')); h.addStretch()
        tv.addLayout(h)
        tv.addWidget(lbl('Thập thần, tàng can, 12 cung, hội – hợp – xung.', 'mo'))
        self.bang = QTableWidget(); self.bang.setColumnCount(6)
        self.bang.setHorizontalHeaderLabels(['Hạng mục', 'Năm', 'Tháng', 'Ngày', 'Giờ', 'Căn cứ / diễn giải'])
        self.bang.verticalHeader().setVisible(False); self.bang.setEditTriggers(QAbstractItemView.NoEditTriggers)
        hd = self.bang.horizontalHeader(); hd.setSectionResizeMode(0, QHeaderView.Fixed); self.bang.setColumnWidth(0, 340)
        for i in range(1, 5): hd.setSectionResizeMode(i, QHeaderView.Fixed); self.bang.setColumnWidth(i, 96)
        hd.setSectionResizeMode(5, QHeaderView.Stretch); self.bang.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.bang.setWordWrap(True); self.bang.setShowGrid(True)
        tv.addWidget(self.bang)
        v.addWidget(the)

    def hien(self, kq):
        o = kq.o; rows = []
        for r in range(4, 79):
            vals = [o('TRUNG_GIAN', '%s%d' % (c, r)) for c in 'BCDEFG']
            if all(x in (None, '') for x in vals): continue
            rows.append((r, vals))
        self.bang.setRowCount(len(rows))
        for i, (r, vals) in enumerate(rows):
            sec = isinstance(vals[0], str) and vals[0].isupper() and len(vals[0]) > 20
            for j, x in enumerate(vals):
                it = QTableWidgetItem(so_hien(x)); it.setToolTip('TRUNG_GIAN!%s%d' % ('BCDEFG'[j], r))
                if j in (1, 2, 3, 4): it.setTextAlignment(Qt.AlignCenter)
                if sec:
                    f = it.font(); f.setBold(True); it.setFont(f); it.setBackground(QColor('#f3f0ff')); it.setForeground(QColor('#4338ca'))
                elif j == 5: it.setForeground(QColor('#7b809c'))
                self.bang.setItem(i, j, it)
        self.bang.resizeRowsToContents()


from ui.bang_sheet import BangSheet
from export.lam_sach import hien_thi_sach as hien_thi_o, lam_sach
from PySide6.QtWidgets import QTabWidget

SHEET_PHAN_TICH = [('PP1_DINH_TINH', 'PP1 – định tính'), ('PP2_DINH_LUONG', 'PP2 – định lượng'), ('THONG_CAN', 'Thông căn'),
                   ('DOI_CHIEU', 'Đối chiếu'), ('DUNG_HY_KY', 'Dụng – hỷ – kỵ')]
# ô kết luận chính (nhãn đọc từ cột B cùng dòng của workbook)
THE_CHINH = [('PP1_DINH_TINH', 91, ['C']), ('PP2_DINH_LUONG', 103, ['C']), ('DOI_CHIEU', 34, ['C']), ('DOI_CHIEU', 35, ['C']),
             ('DUNG_HY_KY', 106, ['C', 'D'])]


class TrangPhanTich(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(14)
        h = QHBoxLayout(); h.addWidget(lbl('Phân tích', 'tieude')); h.addStretch()
        v.addLayout(h)
        self.the = QHBoxLayout(); self.the.setSpacing(12); v.addLayout(self.the)
        self.the_w = []
        for sh, r, cols in THE_CHINH:
            f = QFrame(); f.setObjectName('the'); fl = QVBoxLayout(f); fl.setContentsMargins(16, 12, 16, 12); fl.setSpacing(4)
            a = lbl('', 'nhan', True); b = QLabel(); b.setWordWrap(True); fl.addWidget(a); fl.addWidget(b)
            self.the.addWidget(f); self.the_w.append((sh, r, cols, a, b))
        self.tab = QTabWidget(); self.tab.setDocumentMode(True)
        self.bang = {}
        for sh, ten in SHEET_PHAN_TICH:
            b = BangSheet(sh); self.bang[sh] = b; self.tab.addTab(b, ten)
        v.addWidget(self.tab, 1)
        self.kq = None
        self.tab.currentChanged.connect(lambda _: self._ve_tab())

    def hien(self, kq):
        self.kq = kq; E = kq.E
        for sh, r, cols, a, b in self.the_w:
            a.setText(str(E.gia_tri(sh, 'B%d' % r) or ''))
            vals = [hien_thi_o(E, sh, '%s%d' % (c, r)) for c in cols]
            if len(vals) == 2: txt = 'PP1: <b>%s</b><br>PP2: <b>%s</b>' % tuple(vals)
            else: txt = '<b>%s</b>' % vals[0]
            b.setText("<span style='font-size:15px;color:#1e2140'>%s</span>" % txt)
        self._da_ve = set(); self._ve_tab()

    def _ve_tab(self):
        if not self.kq: return
        sh = SHEET_PHAN_TICH[self.tab.currentIndex()][0]
        if sh in getattr(self, '_da_ve', set()): return
        self.bang[sh].hien(self.kq.E); self._da_ve.add(sh)


from ui.bieu_do import BieuDoNangLuc, BieuDoDaiVan


class TrangBangNhieuSheet(QWidget):
    """Trang gồm nhiều thẻ, mỗi thẻ một sheet (vẽ khi mở thẻ)."""
    def __init__(self, tieu_de, badge, ds):
        super().__init__(); self.ds = ds
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl(tieu_de, 'tieude')); h.addStretch(); v.addLayout(h)
        self.tab = QTabWidget(); self.tab.setDocumentMode(True); self.bang = {}
        from ui.bai_doc import KhungDocLuoi
        for sh, ten in ds:
            b = KhungDocLuoi(sh); self.bang[sh] = b; self.tab.addTab(b, ten)
        v.addWidget(self.tab, 1); self.kq = None; self._da_ve = set()
        self.tab.currentChanged.connect(lambda _: self._ve())

    def hien(self, kq):
        self.kq = kq; self._da_ve = set(); self._ve()

    def _ve(self):
        if not self.kq or not self.isVisible(): return
        sh = self.ds[self.tab.currentIndex()][0]
        if sh in self._da_ve: return
        self.bang[sh].hien(self.kq.E); self._da_ve.add(sh)

    def showEvent(self, e):
        super().showEvent(e); self._ve()


class TrangNguHanhThieu(QWidget):
    doi_che_do = Signal(str)

    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('Ngũ hành thiếu', 'tieude')); h.addSpacing(16)
        h.addWidget(lbl('Chế độ tính mục I, II', 'nhan')); self.che_do = QComboBox(); self.che_do.addItems(CHE_DO_NHT)
        self.che_do.setMinimumWidth(190); h.addWidget(self.che_do); h.addStretch()
        h.addWidget(lbl('Phương pháp riêng – chỉ tham khảo', 'badge')); v.addLayout(h)
        self.tom = lbl('', 'chip_info', True); v.addWidget(self.tom)
        hb = QHBoxLayout(); hb.setSpacing(14)
        c1 = QFrame(); c1.setObjectName('the'); l1 = QVBoxLayout(c1); l1.setContentsMargins(12, 12, 12, 12)
        self.bd1 = BieuDoNangLuc(); l1.addWidget(self.bd1); hb.addWidget(c1, 4)
        c2 = QFrame(); c2.setObjectName('the'); l2 = QVBoxLayout(c2); l2.setContentsMargins(12, 12, 12, 12)
        self.bd2 = BieuDoDaiVan(); l2.addWidget(self.bd2); hb.addWidget(c2, 6)
        v.addLayout(hb, 5)
        from ui.bai_doc import KhungDocLuoi
        self.bang = KhungDocLuoi('NGU_HANH_THIEU'); v.addWidget(self.bang, 4)
        self.che_do.currentTextChanged.connect(self.doi_che_do.emit)
        self.kq = None

    def hien(self, kq):
        self.kq = kq; E = kq.E
        self.che_do.blockSignals(True); self.che_do.setCurrentText(kq.la_so.che_do_nht); self.che_do.blockSignals(False)
        o = lambda r: hien_thi_o(E, 'NGU_HANH_THIEU', r)
        self.tom.setText('Mục I: %s → %s · Mục II: %s → %s' % (o('E15'), o('F15'), o('E27'), o('F27')))
        self.bd1.dat(E); self.bd2.dat(E)
        if self.isVisible(): self.bang.hien(E)

    def showEvent(self, e):
        super().showEvent(e)
        if self.kq: self.bang.hien(self.kq.E)


class TrangKhongRoGio(QWidget):
    doi_gio = Signal(str)

    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('Không rõ giờ', 'tieude')); h.addSpacing(16)
        h.addWidget(lbl('Giờ đang xem', 'nhan')); self.gio = QComboBox(); self.gio.addItems(CHI); h.addWidget(self.gio)
        h.addStretch(); v.addLayout(h)
        self.tom = lbl('', 'chip_info', True); v.addWidget(self.tom)
        from ui.trang_v14 import BangViecTimGio
        self.tab = QTabWidget(); self.viec = BangViecTimGio(); self.bang = BangSheet('KHONG_RO_GIO')
        self.tab.addTab(self.viec, 'Việc lớn đã xảy ra'); self.tab.addTab(self.bang, 'Bảng 12 giờ'); v.addWidget(self.tab, 1)
        self.gio.currentTextChanged.connect(self.doi_gio.emit); self.kq = None

    def hien(self, kq):
        self.kq = kq; E = kq.E
        self.gio.blockSignals(True); self.gio.setCurrentText(kq.la_so.gio_dang_xem); self.gio.blockSignals(False)
        o = lambda r: hien_thi_o(E, 'KHONG_RO_GIO', r)
        self.tom.setText('Chế độ: %s · giờ đang xem %s → trụ giờ %s · việc đã nhập / khớp: %s / %s' % (o('C5'), kq.la_so.gio_dang_xem, o('C27'), o('E23'), o('F23')))
        self.viec.hien(kq)
        if self.isVisible(): self.bang.hien(E)

    def showEvent(self, e):
        super().showEvent(e)
        if self.kq: self.bang.hien(self.kq.E)


SHEET_CHUYEN_DE = [('LUC_THAN', 'Lục thân'), ('HON_NHAN', 'Hôn nhân'), ('MENH_CUC', 'Mệnh cục'), ('TAI_LOC', 'Tài lộc'),
                   ('TO_CHAT', 'Tố chất'), ('SU_NGHIEP', 'Sự nghiệp'), ('SUC_KHOE', 'Sức khỏe (tham khảo)'),
                   ('THAN_SAT', 'Thần sát'), ('SAO', 'Sao'), ('BAO_PHU', 'Bảo phù')]


class CuaSoChinh(QMainWindow):
    """Thanh bên: tầng Đọc lá số (luôn hiện) + tầng Chuyên gia (bật bằng công tắc, ghi nhớ) + Tham khảo / Quản lý."""
    def __init__(self, kho=None):
        super().__init__(); self.setWindowTitle('Bát Tự Thiệu Vĩ Hoa'); self.resize(1400, 900)
        from PySide6.QtCore import QSettings
        from ui.trang_ket_qua import TrangTomTat, TrangLinhVuc, TrangThamKhao, TrangDanhSach, TrangSoSanh
        from ui.trang_v11 import TrangLaSo, TrangLuanGiai, TrangSuKien
        from ui import bang_sheet
        from battu_core.kho_la_so import Kho
        self.kho = kho or Kho(); self.cai_dat = QSettings('BatTuThieuViHoa', 'BatTuThieuViHoa')
        nen = QWidget(); nen.setObjectName('nen'); self.setCentralWidget(nen)
        h = QHBoxLayout(nen); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(0)
        sb = QFrame(); sb.setObjectName('sidebar'); sb.setFixedWidth(236)
        sv = QVBoxLayout(sb); sv.setContentsMargins(18, 24, 18, 18); sv.setSpacing(3)
        sv.addWidget(lbl('☯  Bát Tự', 'logo')); sv.addWidget(lbl('Theo sách Thiệu Vĩ Hoa', 'logo_phu')); sv.addSpacing(14)
        self.t_nhap = TrangNhap(); self.t_tt = TrangTomTat(); self.t_ls = TrangLaSo(); self.t_lv = TrangLinhVuc()
        self.t_lg = TrangLuanGiai(); self.t_sk = TrangSuKien(); self.t_tg = TrangTrungGian(); self.t_pt = TrangPhanTich()
        self.t_vt = TrangBangNhieuSheet('Vận trình', 'DAI_VAN 658 · LUU_NIEN 11.148 · THEO_DOI 5.006 ô công thức',
                                        [('DAI_VAN', 'Đại vận'), ('LUU_NIEN', 'Lưu niên 100 năm (A:DG)'), ('THEO_DOI', 'Theo dõi')])
        self.t_nht = TrangNguHanhThieu()
        self.t_cd = TrangBangNhieuSheet('Chuyên đề', '10 sheet · 7.469 ô công thức · sức khỏe chỉ tham khảo, không phải chẩn đoán', SHEET_CHUYEN_DE)
        self.t_krg = TrangKhongRoGio(); self.t_tk = TrangThamKhao()
        from ui.trang_v14 import TrangDoiChieu
        from ui.trang_tong_hop import TrangTongHop
        self.t_dc = TrangDoiChieu(); self.t_th = TrangTongHop()
        self.t_ds = TrangDanhSach(self.kho); self.t_ss = TrangSoSanh(self.kho)
        # (nhãn, trang, nhóm): 'A' đọc lá số – luôn hiện; 'CG' chỉ ở chế độ chuyên gia; 'KRG' khi bật Không rõ giờ hoặc chuyên gia
        self.muc = [('✎   Nhập liệu', self.t_nhap, 'A'), ('☷   Lá số', self.t_ls, 'A'), ('★   Luận giải chung', self.t_tt, 'A'),
                    ('✦   Luận giải tổng hợp', self.t_th, 'A'), ('❦   Luận giải chi tiết', self.t_lg, 'A'), ('↯   Vận trình', self.t_vt, 'A'), ('▤   Theo lĩnh vực', self.t_lv, 'A'),
                    ('☯   Ngũ hành thiếu', self.t_nht, 'A'), ('✓   Kiểm chứng sự việc', self.t_sk, 'A'),
                    ('⚑   Đối chiếu thực tế', self.t_dc, 'A'),
                    ('◷   Không rõ giờ', self.t_krg, 'KRG'), ('◫   Trung gian', self.t_tg, 'CG'), ('⚖   Phân tích', self.t_pt, 'CG'),
                    ('◎   Chuyên đề', self.t_cd, 'CG'), ('❐   Tham khảo', self.t_tk, 'A'), ('☰   Danh sách lá số', self.t_ds, 'A'),
                    ('⇄   So sánh hai lá', self.t_ss, 'A')]
        self.nav = QButtonGroup(self); self.trang = QStackedWidget()
        for i, (t, w, nhom) in enumerate(self.muc):
            b = QPushButton(t); b.setObjectName('nav'); b.setCheckable(True); b.setCursor(Qt.PointingHandCursor)
            self.nav.addButton(b, i); sv.addWidget(b); self.trang.addWidget(w)
            if i in (9, 13): sv.addSpacing(6)
        sv.addStretch()
        self.cg = QCheckBox('Chế độ chuyên gia'); self.cg.setObjectName('cg')
        self.cg.setToolTip('Hiện thêm Trung gian, Phân tích, Chuyên đề, Lưu niên 100 năm, Theo dõi và địa chỉ ô Excel khi rê chuột')
        sv.addWidget(self.cg); sv.addSpacing(6); sv.addWidget(lbl(PHIEN_BAN, 'phienban'))
        h.addWidget(sb); h.addWidget(self.trang, 1)
        self.nav.idClicked.connect(self.trang.setCurrentIndex)
        ds = [w for _, w, _ in self.muc if w is not self.t_nhap]
        # chỉ vẽ trang đang mở; trang khác vẽ khi được mở (danh sách, so sánh luôn nhận kết quả để lưu/so)
        self._kq = None; self._da_ve = set(); self._luon = {self.t_ds, self.t_ss}
        self.t_nhap.da_tinh.connect(self._nhan_kq)
        self.trang.currentChanged.connect(lambda _: self._ve_trang())
        from PySide6.QtGui import QShortcut, QKeySequence
        QShortcut(QKeySequence(Qt.Key_F5), self, lambda: self.t_nhap.tinh())
        self.t_nht.doi_che_do.connect(lambda x: self.t_nhap.nht.setCurrentText(x))
        self.t_krg.doi_gio.connect(lambda x: self.t_nhap.gio_xem.setCurrentText(x))
        self.t_tt.mo_muc.connect(lambda chu: (self.di(self.t_lv), self.t_lv.mo(chu)))
        self.t_tt.yeu_cau_nhap.connect(lambda: self.di(self.t_nhap)); self.t_tt.yeu_cau_luu.connect(self.t_ds.luu)
        self.t_tt.yeu_cau_xuat.connect(self.xuat)
        self.t_ds.mo_la.connect(lambda s, i: (self.t_nhap.dat_la_so(s), self.di(self.t_tt)))
        self.t_ds.so_sanh.connect(lambda a, b: (self.di(self.t_ss), self.t_ss.nap(a, b)))
        self.t_sk.doi_su_kien.connect(self._doi_su_kien)
        self.t_krg.viec.doi.connect(lambda ds: self._dat_nhap('_viec_tim_gio', ds))
        self.t_dc.doi.connect(lambda d: self._dat_nhap('_theo_doi', d))
        self.cg.setChecked(self.cai_dat.value('chuyen_gia', False, type=bool))
        self.cg.toggled.connect(self._doi_che_do)
        self._ap_che_do()
        if self.t_nhap.kq: self._nhan_kq(self.t_nhap.kq)
        self.di(self.t_nhap)                                 # màn hình đầu: Nhập liệu (duyệt 14/11/2026)

    def _dat_nhap(self, ten, gt):
        setattr(self.t_nhap, ten, gt); self.t_nhap.tinh()

    def _doi_su_kien(self, ds):
        self.t_nhap._su_kien = ds; self.t_nhap.tinh()

    def _nhan_kq(self, kq):
        self._kq = kq; self._da_ve = set()
        for w in self._luon: w.hien(kq); self._da_ve.add(w)
        self._ap_che_do(); self._ve_trang()

    def _ve_trang(self):
        w = self.trang.currentWidget()
        if self._kq is None or w is self.t_nhap or w in self._da_ve: return
        w.hien(self._kq); self._da_ve.add(w)

    def _doi_che_do(self, on):
        self.cai_dat.setValue('chuyen_gia', bool(on)); self._ap_che_do()
        if self._kq:                                          # vẽ lại để cập nhật chú giải / địa chỉ ô
            self._da_ve = set(self._luon); self._ve_trang()

    def _ap_che_do(self):
        from ui import bang_sheet
        cg = self.cg.isChecked(); krg = self.t_nhap.krg.isChecked(); bang_sheet.CHE_DO['chuyen_gia'] = cg
        for i, (_, w, nhom) in enumerate(self.muc):
            hien = nhom == 'A' or (nhom == 'CG' and cg) or (nhom == 'KRG' and (cg or krg))
            self.nav.button(i).setVisible(hien)
            if not hien and self.trang.currentWidget() is w: self.di(self.t_tt)
        for j in (1, 2): self.t_vt.tab.setTabVisible(j, cg)    # Lưu niên 100 năm, Theo dõi: chuyên gia
        if not cg and self.t_vt.tab.currentIndex() in (1, 2): self.t_vt.tab.setCurrentIndex(0)

    def xuat(self):
        if not self.t_nhap.kq: return
        from ui.hop_xuat import HopXuat
        HopXuat(self, self.t_nhap.kq).exec()

    def di(self, w):
        i = w if isinstance(w, int) else self.trang.indexOf(w)
        self.nav.button(i).setChecked(True); self.trang.setCurrentIndex(i)


def tao_app(argv=None):
    app = QApplication.instance() or QApplication(argv or sys.argv)
    nap_font(); app.setFont(QFont('Inter', 10)); app.setStyleSheet(QSS)
    from PySide6.QtGui import QIcon
    app.setWindowIcon(QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bieu_tuong.png')))
    app.setApplicationName('Bát Tự Thiệu Vĩ Hoa')
    return app


def main():
    app = tao_app(sys.argv); w = CuaSoChinh(); w.show(); sys.exit(app.exec())


if __name__ == '__main__':
    main()
