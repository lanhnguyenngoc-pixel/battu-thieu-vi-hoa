# -*- coding: utf-8 -*-
"""Giao diện GĐ5: TÓM TẮT (màn hình đầu), Theo lĩnh vực (KET_LUAN B–K, LUAN_GIAI, DOC_NHANH, SU_KIEN) + tìm kiếm,
Tham khảo (TRANG_CHU, HUONG_DAN, THAM_KHAO, TIET_KHI – chỉ đọc), Danh sách lá số (SQLite), So sánh hai lá.
MAU_THUAN, KIEM_THU, CSDL: chỉ tính ẩn trong lõi, không hiện (quyết định 14/11/2026).
Mọi chữ hiển thị lấy nguyên văn từ lõi tính / workbook – giao diện không tự tính quy tắc."""
import json, os, html
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QWidget, QFrame, QLabel, QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout, QLineEdit,
                               QComboBox, QScrollArea, QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem,
                               QStackedWidget, QTabWidget, QDialog, QDialogButtonBox, QPlainTextEdit, QMessageBox,
                               QAbstractItemView, QHeaderView, QSplitter, QFormLayout)
from export.lam_sach import hien_thi_sach as hien_thi_o
from battu_core.du_lieu import model, ref_of, rc
from battu_core.kho_la_so import Kho, anh_chup
from battu_core.dich_vu import tinh
from ui.bang_sheet import BangSheet, kich_thuoc

# Mục KET_LUAN theo lĩnh vực (YÊU CẦU 5.2): (chữ cái, dòng tiêu đề, dòng KẾT LUẬN). Tiêu đề đọc từ workbook.
MUC_KL = [('B', 14, None), ('C', 28, 29), ('D', 67, 68), ('E', 78, 79), ('F', 92, 93), ('G', 107, 108),
          ('H', 134, 135), ('I', 159, 160), ('J', 178, 179), ('K', 192, 193)]
DONG_AN_KL = (200, 251)            # vùng phụ trợ ẩn (dòng 200-250; dòng 251 chỉ có ô tính phụ AG:AI)
TOM_TAT = list(range(4, 13))       # B4:C12 ↔ mục C…K (MT-07)
TRANG_THEM = [('LUAN_GIAI', 'Luận giải chi tiết (bảng ô)'), ('DOC_NHANH', 'Đọc nhanh'), ('SU_KIEN', 'Đối chiếu sự việc')]


def lbl(text, name=None, wrap=False):
    l = QLabel(text)
    if name: l.setObjectName(name)
    l.setWordWrap(wrap); l.setTextInteractionFlags(Qt.TextSelectableByMouse)
    return l


def cuon(w):
    sc = QScrollArea(); sc.setWidgetResizable(True); sc.setWidget(w); sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    return sc


def chu_o(E, sheet, ref):
    """Chữ hiển thị của một ô (ô công thức hoặc ô hằng), đã bỏ tham chiếu nội bộ workbook."""
    M = model(); si = M.si[sheet]; r, c = rc(ref)
    if (r, c) in M.formula_cells[si]: return hien_thi_o(E, sheet, ref)
    v = E.consts[si].get((r, c))
    if v is None: return ''
    from export.lam_sach import lam_sach
    return hien_thi_o(E, sheet, ref) if isinstance(v, (int, float)) else lam_sach(str(v))


def dong_muc(i):
    _, r1, _ = MUC_KL[i]
    r2 = MUC_KL[i + 1][1] - 1 if i + 1 < len(MUC_KL) else DONG_AN_KL[0] - 1
    return list(range(r1, r2 + 1))


# ---------------------------------------------------------------- TÓM TẮT
class TrangTomTat(QWidget):
    mo_muc = Signal(str)          # chữ cái mục B…K
    yeu_cau_luu = Signal()
    yeu_cau_nhap = Signal()
    yeu_cau_xuat = Signal()

    def __init__(self):
        super().__init__()
        nen = QWidget(); v = QVBoxLayout(nen); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('Luận giải chung', 'tieude'), 1)
        b1 = QPushButton('✎  Nhập / sửa lá số'); b1.setObjectName('nut_phu'); b1.clicked.connect(self.yeu_cau_nhap.emit)
        b2 = QPushButton('💾  Lưu lá số'); b2.setObjectName('nut_chinh'); b2.clicked.connect(self.yeu_cau_luu.emit)
        b3 = QPushButton('⤓  Xuất / In'); b3.setObjectName('nut_phu'); b3.clicked.connect(self.yeu_cau_xuat.emit)
        h.addWidget(b1); h.addWidget(b3); h.addWidget(b2); v.addLayout(h)
        self.dau = lbl('', 'mo', True); v.addWidget(self.dau)
        self.canh_bao = lbl('', 'chip_warn', True); v.addWidget(self.canh_bao)
        g = QGridLayout(); g.setSpacing(12); self.the = []
        for i, r in enumerate(TOM_TAT):
            f = QFrame(); f.setObjectName('the'); fl = QVBoxLayout(f); fl.setContentsMargins(18, 14, 18, 14); fl.setSpacing(6)
            hh = QHBoxLayout(); a = lbl('', 'tieude_the', True); hh.addWidget(a, 1)
            chu = MUC_KL[i + 1][0]
            nut = QPushButton('Chi tiết mục %s →' % chu); nut.setObjectName('nut_phu'); nut.setCursor(Qt.PointingHandCursor)
            nut.clicked.connect(lambda _=False, x=chu: self.mo_muc.emit(x)); hh.addWidget(nut); fl.addLayout(hh)
            b = lbl('', None, True); b.setStyleSheet('font-size:13.5px;'); b.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            b.setTextFormat(Qt.RichText); b.setOpenExternalLinks(False)
            fl.addWidget(b); fl.addStretch(1)
            from ui.trich_dan import TrichDan
            td_ = TrichDan(); b.linkActivated.connect(lambda u, j=i: self._bam(j, u))
            g.addWidget(f, i // 2, i % 2); self.the.append((r, a, b, td_))
        v.addLayout(g); v.addStretch()
        lo = QVBoxLayout(self); lo.setContentsMargins(0, 0, 0, 0); lo.addWidget(cuon(nen))

    def hien(self, kq):
        E = kq.E; o = lambda r: chu_o(E, 'KET_LUAN', r)
        ten = kq.la_so.ho_ten
        tru = ' · '.join(o('%s15' % c) for c in 'CDEF')
        self.dau.setText(('<b>%s</b> — ' % html.escape(ten) if ten else '') + 'Tứ trụ (năm · tháng · ngày · giờ): <b>%s</b> · %s · năm xem %s'
                         % (tru, kq.la_so.gioi_tinh, chu_o(E, 'NHAP_LIEU', 'C31')))
        cb = o('B2'); self.canh_bao.setText(cb); self.canh_bao.setVisible(bool(cb.strip()))
        self.kq = kq
        for j, (r, a, b, td_) in enumerate(self.the):
            td_.mo = set(); a.setText(o('B%d' % r)); self._ve_the(j)

    def _ve_the(self, j):
        r, a, b, td_ = self.the[j]; td_.bat_dau()
        # kết luận của công cụ in đậm; căn cứ sách ở dòng riêng, bấm "Căn cứ sách" mới sổ ra
        b.setText('<div style="line-height:145%%">%s</div>' % td_.html(chu_o(self.kq.E, 'KET_LUAN', 'C%d' % r))); b.updateGeometry(); b.parentWidget().updateGeometry()

    def _bam(self, j, u):
        if self.the[j][3].bam(u): self._ve_the(j)


# ---------------------------------------------------------------- THEO LĨNH VỰC + TÌM KIẾM
class TrangLinhVuc(QWidget):
    """Mục B–K theo lĩnh vực. Chế độ thường: dạng văn bản đọc (ui.bai_doc); chế độ chuyên gia: lưới ô gốc."""
    def __init__(self):
        super().__init__()
        from ui.bai_doc import BaiDoc
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); h.addWidget(lbl('Theo lĩnh vực', 'tieude')); h.addStretch()
        self.o_tim = QLineEdit(); self.o_tim.setPlaceholderText('🔍  Tìm trong kết luận, luận giải… (Enter)'); self.o_tim.setMinimumWidth(360)
        h.addWidget(self.o_tim); v.addLayout(h)
        self.kq_tim = QListWidget(); self.kq_tim.setMaximumHeight(170); self.kq_tim.setVisible(False); v.addWidget(self.kq_tim)
        sp = QSplitter(Qt.Horizontal)
        self.ds = QListWidget(); self.ds.setFixedWidth(300); self.ds.setWordWrap(True)
        self.ds.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff); self.ds.setStyleSheet(
            'QListWidget{background:white;border:1px solid #e7e8f3;border-radius:14px;padding:6px;font-size:13px}'
            'QListWidget::item{padding:8px;border-radius:8px}QListWidget::item:selected{background:#eef0ff;color:#4338ca;font-weight:700}')
        sp.addWidget(self.ds)
        self.ngan = QStackedWidget(); sp.addWidget(self.ngan); sp.setStretchFactor(1, 1)
        v.addWidget(sp, 1)
        self.muc = []                     # (khóa, sheet, BangSheet, nhãn KẾT LUẬN, dòng KL, BaiDoc, ngăn đọc/lưới, dòng tiêu đề)
        from ui.trich_dan import TrichDan
        self.td_kl = {i: TrichDan() for i in range(len(MUC_KL))}
        for i, (chu, r0, rkl) in enumerate(MUC_KL):
            w = QWidget(); wl = QVBoxLayout(w); wl.setContentsMargins(0, 0, 0, 0); wl.setSpacing(8)
            kl = None
            if rkl:
                kl = lbl('', 'chip_info', True); kl.setStyleSheet('font-size:13.5px; padding:12px 14px;')
                kl.setTextFormat(Qt.RichText); kl.setOpenExternalLinks(False)
                kl.linkActivated.connect(lambda u, j=i: self.td_kl[j].bam(u) and self._ve_kl(j))
                sk = cuon(kl); sk.setMaximumHeight(210); wl.addWidget(sk)
            if chu == 'H':
                wl.addWidget(lbl('Nội dung sức khỏe chỉ để tham khảo, không phải chẩn đoán y khoa.', 'chip_warn', True))
            hai = QStackedWidget(); bd = BaiDoc(); b = BangSheet('KET_LUAN', dong=dong_muc(i))
            hai.addWidget(bd); hai.addWidget(b); wl.addWidget(hai, 1)
            self.ngan.addWidget(w); self.muc.append((chu, 'KET_LUAN', b, kl, rkl, bd, hai, r0))
            self.ds.addItem(QListWidgetItem('%s' % chu))
        for sh, ten in TRANG_THEM:
            hai = QStackedWidget(); bd = BaiDoc(); b = BangSheet(sh); hai.addWidget(bd); hai.addWidget(b)
            self.ngan.addWidget(hai); self.muc.append((sh, sh, b, None, None, bd, hai, None)); self.ds.addItem(QListWidgetItem(ten))
        self.ds.currentRowChanged.connect(self._chon)
        self.o_tim.returnPressed.connect(self.tim); self.kq_tim.itemClicked.connect(self._toi)
        self.kq = None; self._da_ve = set()
        self.ds.setCurrentRow(1)

    @staticmethod
    def _cg():
        from ui.bang_sheet import CHE_DO
        return CHE_DO['chuyen_gia']

    def hien(self, kq):
        self.kq = kq; self._da_ve = set(); E = kq.E; cg = self._cg()
        for i, (chu, r0, rkl) in enumerate(MUC_KL):
            self.ds.item(i).setText(chu_o(E, 'KET_LUAN', 'B%d' % r0) or chu)
            kl = self.muc[i][3]
            if kl: self.td_kl[i].mo = set(); self._ve_kl(i)
        for j, (sh, _) in enumerate(TRANG_THEM):            # Luận giải (bảng ô), Đối chiếu sự việc: đã có trang riêng → chỉ chuyên gia
            self.ds.setRowHidden(len(MUC_KL) + j, (not cg) and sh in ('LUAN_GIAI', 'SU_KIEN'))
        if self.ds.isRowHidden(self.ds.currentRow()): self.ds.setCurrentRow(1)
        self._ve()
        if self.kq_tim.isVisible(): self.tim()

    def _ve_kl(self, i):
        chu, r0, rkl = MUC_KL[i]; E = self.kq.E; t = self.td_kl[i]; t.bat_dau()
        them = ''
        if chu == 'G':      # mục 2.9: dòng tham khảo con cái theo lá số vợ/chồng (nếu đã nhập), ngay dưới kết luận Gia đạo
            from battu_core.tong_hop.viet_ban import dong_hon_phoi
            if not hasattr(self.kq, '_dong_hp'): self.kq._dong_hp = dong_hon_phoi(self.kq.la_so)
            if self.kq._dong_hp: them = '<br><br>' + html.escape(self.kq._dong_hp)
        self.muc[i][3].setText('<div style="line-height:145%%"><b>%s</b><br>%s%s</div>' % (html.escape(chu_o(E, 'KET_LUAN', 'B%d' % rkl)), t.html(chu_o(E, 'KET_LUAN', 'C%d' % rkl)), them))

    def _chon(self, i):
        self.ngan.setCurrentIndex(i); self._ve()

    def _ve(self):
        i = self.ngan.currentIndex()
        if not self.kq or i < 0 or i in self._da_ve or not self.isVisible(): return
        k, sh, b, kl, rkl, bd, hai, r0 = self.muc[i]
        if self._cg():
            b.hien(self.kq.E); hai.setCurrentIndex(1)
        else:
            bd.hien(self.kq.E, sh, dong_muc(i) if sh == 'KET_LUAN' else None, bo_dong={r0, rkl}); hai.setCurrentIndex(0)
        self._da_ve.add(i)

    def showEvent(self, e):
        super().showEvent(e); self._ve()

    def mo(self, chu):
        for i, m in enumerate(self.muc):
            if m[0] == chu: self.ds.setCurrentRow(i); return

    def tim(self):
        chu = self.o_tim.text().strip().lower(); self.kq_tim.clear()
        if not chu or not self.kq: self.kq_tim.setVisible(False); return
        E = self.kq.E; M = model(); n = 0; cg = self._cg()
        for i, m in enumerate(self.muc):
            k, sh, b = m[0], m[1], m[2]
            if self.ds.isRowHidden(i): continue
            si = M.si[sh]
            for r, c in b.o_hien():
                    if (r, c) not in M.formula_cells[si] and (r, c) not in E.consts[si]: continue
                    t = chu_o(E, sh, ref_of(r, c))
                    j = t.lower().find(chu)
                    if j < 0: continue
                    doan = t[max(0, j - 40): j + 80].replace('\n', ' ')
                    noi = ('%s!%s   ' % (sh, ref_of(r, c))) if cg else ''
                    it = QListWidgetItem('%s · %s…%s…' % (self.ds.item(i).text()[:38], noi, doan))
                    it.setData(Qt.UserRole, (i, r, c, t[j:j + 40])); self.kq_tim.addItem(it); n += 1
                    if n >= 300: break
            if n >= 300: break
        if n == 0: self.kq_tim.addItem('Không tìm thấy "%s".' % self.o_tim.text().strip())
        self.kq_tim.setVisible(True)

    def _toi(self, it):
        d = it.data(Qt.UserRole)
        if not d: return
        i, r, c, doan = d; self.ds.setCurrentRow(i)
        if self._cg(): self.muc[i][2].toi_o(r, c)
        else:
            bd = self.muc[i][5]; bd.moveCursor(bd.textCursor().MoveOperation.Start)
            bd.find(doan.split('\n')[0][:30])


# ---------------------------------------------------------------- THAM KHẢO (chỉ đọc)
THAM_KHAO = [('THAM_KHAO', 'Tham khảo ngoài sách'), ('TIET_KHI', 'Tiết khí 1899–2101')]   # Trang chủ, Hướng dẫn của workbook: chỉ hướng dẫn thao tác Excel → bỏ


class TrangThamKhao(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('Tham khảo', 'tieude')); h.addStretch(); v.addLayout(h)
        self.tab = QTabWidget(); self.tab.setDocumentMode(True); self.bang = []
        for sh, ten in THAM_KHAO:
            b = BangSheet(sh); self.bang.append(b); self.tab.addTab(b, ten)
        v.addWidget(self.tab, 1); self.kq = None; self._da_ve = set()
        self.tab.currentChanged.connect(lambda _: self._ve())

    def hien(self, kq):
        self.kq = kq; self._da_ve = set(); self._ve()

    def _ve(self):
        i = self.tab.currentIndex()
        if not self.kq or i in self._da_ve or not self.isVisible(): return
        self.bang[i].hien(self.kq.E); self._da_ve.add(i)

    def showEvent(self, e):
        super().showEvent(e); self._ve()


# ---------------------------------------------------------------- DANH SÁCH LÁ SỐ (SQLite)
class HopThongTin(QDialog):
    def __init__(self, cha, tieu_de, ten='', nhom='', ghi_chu='', ds_nhom=()):
        super().__init__(cha); self.setWindowTitle(tieu_de); self.setMinimumWidth(460)
        f = QFormLayout(self)
        self.ten = QLineEdit(ten); self.nhom = QComboBox(); self.nhom.setEditable(True); self.nhom.addItems([''] + list(ds_nhom))
        self.nhom.setCurrentText(nhom); self.ghi = QPlainTextEdit(ghi_chu); self.ghi.setFixedHeight(110)
        f.addRow('Tên lá số', self.ten); f.addRow('Nhóm', self.nhom); f.addRow('Ghi chú', self.ghi)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel); bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        f.addRow(bb)

    def gia_tri(self): return self.ten.text().strip(), self.nhom.currentText().strip(), self.ghi.toPlainText().strip()


class TrangDanhSach(QWidget):
    mo_la = Signal(object, int)          # LaSo, id
    so_sanh = Signal(int, int)

    COT = ['Tên', 'Nhóm', 'Tứ trụ', 'Ngày sinh (DL)', 'Ghi chú', 'Sửa lúc', 'Quy tắc']

    def __init__(self, kho=None):
        super().__init__(); self.kho = kho or Kho(); self.kq = None; self.id_dang_mo = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('Danh sách lá số', 'tieude')); h.addStretch()
        self.setToolTip('Lá số lưu tại: ' + self.kho.duong_dan); v.addLayout(h)
        ht = QHBoxLayout()
        self.o_tim = QLineEdit(); self.o_tim.setPlaceholderText('🔍  Tìm theo tên, nhóm, ghi chú, tứ trụ, ngày sinh'); ht.addWidget(self.o_tim, 1)
        self.cb_nhom = QComboBox(); self.cb_nhom.setMinimumWidth(180); ht.addWidget(self.cb_nhom); v.addLayout(ht)
        hb = QHBoxLayout()
        self.n_luu = QPushButton('💾  Lưu lá đang xem'); self.n_luu.setObjectName('nut_chinh')
        self.n_ghi = QPushButton('⟳  Ghi đè bằng lá đang xem'); self.n_mo = QPushButton('Mở'); self.n_sua = QPushButton('Sửa tên/nhóm/ghi chú')
        self.n_xoa = QPushButton('Xóa'); self.n_ss = QPushButton('⇄  So sánh 2 lá đã chọn')
        for b in (self.n_ghi, self.n_mo, self.n_sua, self.n_xoa, self.n_ss): b.setObjectName('nut_phu')
        for b in (self.n_luu, self.n_mo, self.n_ghi, self.n_sua, self.n_xoa, self.n_ss): hb.addWidget(b)
        hb.addStretch()
        self.n_tk = QPushButton('⚑  Thống kê đối chiếu')
        self.n_sl = QPushButton('⤓  Sao lưu kho…'); self.n_kp = QPushButton('⤒  Khôi phục từ bản sao lưu…')
        for b in (self.n_tk, self.n_sl, self.n_kp): b.setObjectName('nut_phu'); hb.addWidget(b)
        v.addLayout(hb)
        self.bang = QTableWidget(0, len(self.COT)); self.bang.setHorizontalHeaderLabels(self.COT)
        self.bang.setSelectionBehavior(QAbstractItemView.SelectRows); self.bang.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.bang.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive); self.bang.horizontalHeader().setStretchLastSection(True)
        for i, w in enumerate([200, 120, 260, 120, 260, 140]): self.bang.setColumnWidth(i, w)
        self.bang.verticalHeader().setVisible(False); v.addWidget(self.bang, 1)
        self.tt = lbl('', 'mo'); v.addWidget(self.tt)
        self.o_tim.textChanged.connect(self.nap); self.cb_nhom.currentIndexChanged.connect(self.nap)
        self.n_luu.clicked.connect(self.luu); self.n_ghi.clicked.connect(self.ghi_de); self.n_mo.clicked.connect(self.mo)
        self.n_sua.clicked.connect(self.sua); self.n_xoa.clicked.connect(self.xoa); self.n_ss.clicked.connect(self._ss)
        self.bang.doubleClicked.connect(lambda _: self.mo())
        self.n_sl.clicked.connect(self.sao_luu); self.n_kp.clicked.connect(self.khoi_phuc)
        self.n_tk.clicked.connect(lambda: HopThongKe(self, self.kho).exec())
        self.nap_nhom(); self.nap()

    def sao_luu(self):
        import datetime
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        ten = 'SaoLuu_LaSo_%s.sqlite' % datetime.date.today().strftime('%Y%m%d')
        p, _ = QFileDialog.getSaveFileName(self, 'Sao lưu kho lá số', os.path.join(os.path.expanduser('~'), ten), 'Kho lá số (*.sqlite)')
        if not p: return
        try:
            n = self.kho.sao_luu(p)
            QMessageBox.information(self, 'Sao lưu', 'Đã sao lưu %d lá số vào:\n%s' % (n, p))
        except Exception as e:
            QMessageBox.warning(self, 'Sao lưu', 'Không sao lưu được: %s' % e)

    def khoi_phuc(self):
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        p, _ = QFileDialog.getOpenFileName(self, 'Khôi phục từ bản sao lưu', os.path.expanduser('~'), 'Kho lá số (*.sqlite)')
        if not p: return
        try:
            them, bo = self.kho.khoi_phuc(p)
            QMessageBox.information(self, 'Khôi phục', 'Đã thêm %d lá số%s. Các lá đang có trong kho được giữ nguyên.' % (
                them, (' (bỏ qua %d lá đã có)' % bo) if bo else ''))
        except Exception as e:
            QMessageBox.warning(self, 'Khôi phục', 'Không khôi phục được: %s' % e)
        self.nap_nhom(); self.nap()

    def hien(self, kq): self.kq = kq

    def nap_nhom(self):
        cur = self.cb_nhom.currentText(); self.cb_nhom.blockSignals(True); self.cb_nhom.clear()
        self.cb_nhom.addItems(['Tất cả nhóm'] + self.kho.nhom()); self.cb_nhom.setCurrentText(cur or 'Tất cả nhóm')
        self.cb_nhom.blockSignals(False)

    def nap(self):
        nhom = self.cb_nhom.currentText(); nhom = None if nhom in ('', 'Tất cả nhóm') else nhom
        rows = self.kho.tim(self.o_tim.text(), nhom)
        self.bang.setRowCount(len(rows))
        for i, r in enumerate(rows):
            n = json.loads(r['nhap'])
            ngay = '%02d/%02d/%d' % (n['ngay'], n['thang'], n['nam']) if n.get('cach_nhap') == 'duong' and n.get('nam') else \
                ('ÂL %s/%s/%s' % (n.get('am_ngay'), n.get('am_thang'), n.get('am_nam')) if n.get('cach_nhap') == 'am' else 'Tự nhập can chi')
            for j, t in enumerate([r['ten'] or '(chưa đặt tên)', r['nhom'], r['tu_tru'], ngay, r['ghi_chu'].replace('\n', ' '),
                                   r['sua_luc'], r['phien_ban']]):
                it = QTableWidgetItem(t); it.setData(Qt.UserRole, r['id']); self.bang.setItem(i, j, it)
        self.tt.setText('%d lá số' % len(rows))

    def _chon(self):
        return sorted({self.bang.item(i.row(), 0).data(Qt.UserRole) for i in self.bang.selectionModel().selectedRows()})

    def luu(self):
        if not self.kq: return
        d = HopThongTin(self, 'Lưu lá số', self.kq.la_so.ho_ten, '', '', self.kho.nhom())
        if d.exec() != QDialog.Accepted: return
        ten, nhom, ghi = d.gia_tri()
        self.id_dang_mo = self.kho.luu(self.kq, ten, nhom, ghi); self.nap_nhom(); self.nap()

    def ghi_de(self):
        ids = self._chon()
        if len(ids) != 1 or not self.kq: return
        r = self.kho.lay(ids[0])
        if QMessageBox.question(self, 'Ghi đè', 'Thay dữ liệu nhập của "%s" bằng lá đang xem?' % (r['ten'] or r['id'])) != QMessageBox.Yes: return
        self.kho.luu(self.kq, r['ten'], r['nhom'], r['ghi_chu'], ids[0]); self.nap()

    def mo(self):
        ids = self._chon()
        if len(ids) != 1: return
        if not self.kho.cung_phien_ban(ids[0]):
            QMessageBox.information(self, 'Phiên bản quy tắc khác', 'Lá số được lưu với phiên bản quy tắc khác; kết quả sẽ tính lại theo quy tắc 6.3 hiện tại.')
        self.id_dang_mo = ids[0]; self.mo_la.emit(self.kho.la_so(ids[0]), ids[0])

    def sua(self):
        ids = self._chon()
        if len(ids) != 1: return
        r = self.kho.lay(ids[0]); d = HopThongTin(self, 'Sửa thông tin', r['ten'], r['nhom'], r['ghi_chu'], self.kho.nhom())
        if d.exec() == QDialog.Accepted:
            self.kho.sua_thong_tin(ids[0], *d.gia_tri()); self.nap_nhom(); self.nap()

    def xoa(self):
        ids = self._chon()
        if not ids: return
        if QMessageBox.question(self, 'Xóa lá số', 'Xóa %d lá số đã chọn? Không hoàn tác được.' % len(ids)) != QMessageBox.Yes: return
        for i in ids: self.kho.xoa(i)
        self.nap_nhom(); self.nap()

    def _ss(self):
        ids = self._chon()
        if len(ids) == 2: self.so_sanh.emit(ids[0], ids[1])
        else: QMessageBox.information(self, 'So sánh', 'Chọn đúng 2 lá số (giữ Ctrl khi bấm).')


# ---------------------------------------------------------------- SO SÁNH HAI LÁ
class TrangSoSanh(QWidget):
    def __init__(self, kho):
        super().__init__(); self.kho = kho; self.kq = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('So sánh hai lá số', 'tieude')); h.addStretch()
        v.addLayout(h)
        hc = QHBoxLayout(); self.cb = [QComboBox(), QComboBox()]
        for i, c in enumerate(self.cb):
            c.setMinimumWidth(300); hc.addWidget(lbl('Lá %d' % (i + 1), 'nhan')); hc.addWidget(c)
            c.currentIndexChanged.connect(self.ve)
        hc.addStretch(); v.addLayout(hc)
        self.bang = QTableWidget(0, 3); self.bang.setHorizontalHeaderLabels(['Mục', 'Lá 1', 'Lá 2'])
        self.bang.setEditTriggers(QAbstractItemView.NoEditTriggers); self.bang.setWordWrap(True)
        self.bang.verticalHeader().setVisible(False); self.bang.setColumnWidth(0, 200)
        self.bang.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch); self.bang.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        v.addWidget(self.bang, 1); self._cache = {}

    def hien(self, kq):
        self.kq = kq; self._cache.pop(0, None)
        if self.isVisible(): self.ve()

    def nap(self, a=None, b=None):
        ds = [(0, 'Lá đang xem')] + [(r['id'], (r['ten'] or '(chưa đặt tên)') + ' · ' + r['tu_tru']) for r in self.kho.tim()]
        for k, c in enumerate(self.cb):
            cur = c.currentData(); c.blockSignals(True); c.clear()
            for i, t in ds: c.addItem(t, i)
            muon = (a, b)[k]
            muon = muon if muon is not None else cur
            j = c.findData(muon) if muon is not None else -1
            c.setCurrentIndex(j if j >= 0 else min(k, c.count() - 1)); c.blockSignals(False)
        self.ve()

    def _ket_qua(self, i):
        if i == 0: return self.kq
        if i not in self._cache:
            s = self.kho.la_so(i)
            self._cache[i] = tinh(s) if s else None
        return self._cache[i]

    def showEvent(self, e):
        super().showEvent(e); self.nap()

    def ve(self):
        kqs = [self._ket_qua(c.currentData()) if c.currentData() is not None else None for c in self.cb]
        if not all(kqs): return
        dong = [('Tứ trụ', lambda E: ' · '.join(chu_o(E, 'KET_LUAN', '%s15' % c) for c in 'CDEF')),
                ('Cảnh báo', lambda E: chu_o(E, 'KET_LUAN', 'B2'))]
        dong += [(None, (lambda r: lambda E: chu_o(E, 'KET_LUAN', 'C%d' % r))(r), r) for r in TOM_TAT]
        self.bang.setRowCount(len(dong))
        for i, d in enumerate(dong):
            ten = d[0] if d[0] else chu_o(kqs[0].E, 'KET_LUAN', 'B%d' % d[2])
            it = QTableWidgetItem(ten); f = it.font(); f.setBold(True); it.setFont(f); self.bang.setItem(i, 0, it)
            t = [d[1](k.E) for k in kqs]
            for j in (0, 1):
                x = QTableWidgetItem(t[j])
                if t[0] != t[1]: x.setBackground(QColor('#fff7e6'))
                self.bang.setItem(i, j + 1, x)
        self.bang.resizeRowsToContents()
        for i in range(self.bang.rowCount()):
            if self.bang.rowHeight(i) > 220: self.bang.setRowHeight(i, 220)


class HopThongKe(QDialog):
    """Thống kê đối chiếu thực tế trên cả kho lá số (cộng các ô kết quả Theo dõi của từng lá; không thêm quy tắc)."""
    def __init__(self, cha, kho):
        super().__init__(cha); self.setWindowTitle('Thống kê đối chiếu thực tế'); self.resize(1000, 640); self.kho = kho
        from PySide6.QtWidgets import QProgressBar, QApplication
        v = QVBoxLayout(self); v.setContentsMargins(20, 18, 20, 18); v.setSpacing(10)
        v.addWidget(lbl('Thống kê đối chiếu thực tế trên cả kho', 'tieude'))
        v.addWidget(lbl('Cộng các năm đã ghi ở trang Đối chiếu thực tế của mọi lá số đã lưu: phương pháp vượng suy nào (PP1 hay PP2) '
                        'khớp đời thực hơn, và luận giải đúng đến đâu. Chỉ là thống kê, không đổi kết luận của lá nào.', 'mo', True))
        self.tom = lbl('', None, True); v.addWidget(self.tom)
        self.bang = QTableWidget(0, 8); self.bang.verticalHeader().setVisible(False)
        self.bang.setHorizontalHeaderLabels(['Tên', 'Tứ trụ', 'Số năm', 'Khớp PP1', 'Khớp PP2', 'Đúng / Một phần / Sai / Chưa rõ', 'Đang dùng', 'Gợi ý'])
        self.bang.setEditTriggers(QAbstractItemView.NoEditTriggers); self.bang.horizontalHeader().setStretchLastSection(True)
        for i, w in enumerate([160, 230, 70, 80, 80, 240, 140]): self.bang.setColumnWidth(i, w)
        v.addWidget(self.bang, 1)
        h = QHBoxLayout(); h.addStretch()
        self.n_x = QPushButton('Xuất Excel…'); self.n_x.setObjectName('nut_phu'); self.n_d = QPushButton('Đóng'); self.n_d.setObjectName('nut_chinh')
        h.addWidget(self.n_x); h.addWidget(self.n_d); v.addLayout(h)
        self.n_d.clicked.connect(self.accept); self.n_x.clicked.connect(self.xuat)
        self.show(); QApplication.processEvents()
        from battu_core.thong_ke import thong_ke
        self.tom.setText('Đang tính…'); QApplication.processEvents()
        self.kq = thong_ke(kho, lambda i, n: (self.tom.setText('Đang tính lá %d / %d…' % (i, n)), QApplication.processEvents()))
        self.hien()

    def hien(self):
        k = self.kq; pt = lambda x: '–' if x is None else '%.0f%%' % (100 * x)
        if not k['so_la']:
            self.tom.setText('Chưa có lá số nào được ghi thực tế. Mở một lá số → trang ⚑ Đối chiếu thực tế → ghi các năm đã qua → lưu lá số.')
            self.n_x.setEnabled(False); return
        d = k['danh_gia']
        self.tom.setText('<b>%d lá · %d năm đã ghi</b> — khớp xu thế: <b>PP1 %s</b>, <b>PP2 %s</b> · đánh giá luận giải: Đúng %d · Một phần %d · '
                         'Sai %d · Chưa rõ %d — tỉ lệ đúng (một phần tính ½): <b>%s</b>' % (
                             k['so_la'], k['so_nam'], pt(k['ty_le_pp1']), pt(k['ty_le_pp2']), d['Đúng'], d['Một phần'], d['Sai'], d['Chưa rõ'],
                             pt(k['ty_le_dung'])))
        self.bang.setRowCount(len(k['tung_la']))
        for i, x in enumerate(k['tung_la']):
            dg = ' / '.join(str(x['danh_gia'].get(t, 0)) for t in ('Đúng', 'Một phần', 'Sai', 'Chưa rõ'))
            for j, t in enumerate([x['ten'], x['tu_tru'], str(x['so_nam']), str(x['khop_pp1']), str(x['khop_pp2']), dg, x['dang_dung'], x['goi_y']]):
                it = QTableWidgetItem(t); it.setToolTip(t) if len(t) > 30 else None; self.bang.setItem(i, j, it)

    def xuat(self):
        from PySide6.QtWidgets import QFileDialog, QMessageBox
        from battu_core.thong_ke import xuat_xlsx
        p, _ = QFileDialog.getSaveFileName(self, 'Xuất thống kê', os.path.join(os.path.expanduser('~'), 'ThongKe_DoiChieu.xlsx'), 'Excel (*.xlsx)')
        if not p: return
        try: xuat_xlsx(self.kq, p); QMessageBox.information(self, 'Xuất thống kê', 'Đã lưu: ' + p)
        except Exception as e: QMessageBox.warning(self, 'Xuất thống kê', 'Không lưu được: %s' % e)
