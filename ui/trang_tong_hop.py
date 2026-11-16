# -*- coding: utf-8 -*-
"""Trang "Luận giải tổng hợp" (đề xuất hoàn thiện 15/11/2026, T4–T5): bản ~5 trang A4 không cần mạng, soạn từ hồ sơ
dữ kiện + lớp phán đoán (battu_core.tong_hop); tùy chọn "Nâng cao bằng AI" (Claude API, khóa của người dùng, không gửi
họ tên, bộ kiểm tra đối chiếu dữ kiện); xuất PDF hoặc Word (.docx) bản đang xem."""
import threading
from PySide6.QtCore import QObject, Signal, QSettings
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextBrowser, QFileDialog, QMessageBox,
                               QDialog, QLineEdit, QComboBox, QDialogButtonBox, QButtonGroup)


class _Bao(QObject):
    xong = Signal(object)


def _mo_hinh():
    from battu_core.tong_hop.ai_claude import MO_HINH, MO_HINH_MAC_DINH
    v = QSettings('BatTuThieuViHoa', 'BatTuThieuViHoa').value('ai_mo_hinh', MO_HINH_MAC_DINH)
    return v if v in [k for k, _ in MO_HINH] else MO_HINH_MAC_DINH


class HopKhoaAPI(QDialog):
    """Nhập / đổi / xóa khóa API và chọn mô hình."""
    def __init__(self, cha=None):
        super().__init__(cha); self.setWindowTitle('Khóa API Anthropic'); self.resize(520, 260)
        from ui import khoa_api
        from battu_core.tong_hop.ai_claude import MO_HINH
        v = QVBoxLayout(self)
        t = QLabel('Tạo khóa tại console.anthropic.com (mục API Keys). Khóa lưu trong kho khóa của hệ điều hành trên máy này; '
                   'phí dùng API tính vào tài khoản của bạn. Phần mềm KHÔNG gửi họ tên, chỉ gửi tứ trụ và các kết luận đã tính.')
        t.setWordWrap(True); v.addWidget(t)
        v.addWidget(QLabel('Khóa hiện tại: %s' % (khoa_api.che(khoa_api.doc()) or 'chưa có')))
        self.o = QLineEdit(); self.o.setEchoMode(QLineEdit.Password); self.o.setPlaceholderText('sk-ant-… (để trống = giữ khóa cũ)')
        v.addWidget(self.o)
        self.mh = QComboBox()
        for k, ten in MO_HINH: self.mh.addItem(ten, k)
        self.mh.setCurrentIndex([k for k, _ in MO_HINH].index(_mo_hinh())); v.addWidget(QLabel('Mô hình')); v.addWidget(self.mh)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.n_xoa = bb.addButton('Xóa khóa', QDialogButtonBox.DestructiveRole)
        bb.accepted.connect(self._luu); bb.rejected.connect(self.reject); self.n_xoa.clicked.connect(self._xoa); v.addWidget(bb)

    def _luu(self):
        from ui import khoa_api
        if self.o.text().strip():
            QMessageBox.information(self, 'Đã lưu', 'Đã lưu khóa vào %s.' % khoa_api.ghi(self.o.text().strip()))
        QSettings('BatTuThieuViHoa', 'BatTuThieuViHoa').setValue('ai_mo_hinh', self.mh.currentData()); self.accept()

    def _xoa(self):
        from ui import khoa_api
        khoa_api.xoa(); QMessageBox.information(self, 'Đã xóa', 'Đã xóa khóa API khỏi máy này.'); self.accept()


class TrangTongHop(QWidget):
    def __init__(self):
        super().__init__(); self.kq = None; self.B = None; self.BA = None; self.hs = self.pd = None; self._loi = ''
        self._bao = _Bao(); self._bao.xong.connect(self._ai_xong); self._dang = False
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); t = QLabel('Luận giải tổng hợp'); t.setObjectName('tieude'); h.addWidget(t); h.addStretch()
        self.n_ai = QPushButton('✨  Nâng cao bằng AI'); self.n_ai.setObjectName('nut_phu')
        self.n_ai.setToolTip('Tùy chọn, cần mạng và khóa API Anthropic của bạn. Gửi tứ trụ, năm sinh và kết luận đã tính; không gửi họ tên.')
        self.n_khoa = QPushButton('🔑'); self.n_khoa.setObjectName('nut_phu'); self.n_khoa.setToolTip('Khóa API và mô hình AI')
        self.n_pdf = QPushButton('📄  Xuất PDF'); self.n_pdf.setObjectName('nut_chinh')
        self.n_doc = QPushButton('📝  Xuất Word'); self.n_doc.setObjectName('nut_phu')
        for n in (self.n_ai, self.n_khoa, self.n_pdf, self.n_doc): h.addWidget(n)
        for n in (self.n_ai, self.n_pdf, self.n_doc): n.setEnabled(False)
        v.addLayout(h)
        hv = QHBoxLayout(); self.ban = QButtonGroup(self)
        self.b0 = QPushButton('Bản không mạng'); self.b1 = QPushButton('Bản AI')
        for i, b in enumerate((self.b0, self.b1)): b.setCheckable(True); b.setObjectName('nut_phu'); self.ban.addButton(b, i); hv.addWidget(b)
        self.b0.setChecked(True); self.b1.setEnabled(False); hv.addSpacing(10)
        self.tt = QLabel('Bản tổng hợp khoảng 5 trang A4, 10 lĩnh vực: mỗi việc một kết luận đã gỡ mâu thuẫn, kèm lý do. Chạy không cần mạng.')
        self.tt.setObjectName('mo'); self.tt.setWordWrap(True); hv.addWidget(self.tt, 1); v.addLayout(hv)
        self.xem = QTextBrowser(); self.xem.setOpenExternalLinks(False); v.addWidget(self.xem, 1)
        self.n_pdf.clicked.connect(self.pdf); self.n_doc.clicked.connect(self.docx)
        self.n_ai.clicked.connect(self.ai); self.n_khoa.clicked.connect(lambda: (HopKhoaAPI(self).exec(), self._thu_dem(), self._ve()))
        self.ban.idClicked.connect(lambda _: self._ve())

    # ------------------------------------------------------------ hiển thị
    def hien(self, kq):
        from battu_core.tong_hop.viet_ban import ban_luan_giai
        self.kq = kq; self.BA = None; self.b1.setEnabled(False); self.b0.setChecked(True)
        try:
            self.B, self.hs, self.pd = ban_luan_giai(kq); ok = True
        except Exception as e:                                   # lá số thiếu dữ kiện (vd chưa đủ tứ trụ)
            self.B = None; ok = False; self._loi = str(e)
        for n in (self.n_ai, self.n_pdf, self.n_doc): n.setEnabled(ok)
        if ok: self._thu_dem()
        self._ve()

    def _thu_dem(self):
        """Đã có bản AI lưu trên máy cho đúng lá số + mô hình này → bật sẵn, không gọi lại."""
        if not self.B: return
        from battu_core.tong_hop import ai_claude as A
        mh = _mo_hinh()
        try:
            chu = A.doc_dem(A.khoa_dem(self.hs, self.pd, mh))
            if chu:
                BA = A.doc_md(chu, self.B[0][1])
                if not A.kiem_tra(BA, self.hs, self.pd):
                    BA.insert(1, ('nho', 'Bản do AI (%s) viết lại – đã lưu trên máy, đã qua bộ kiểm tra.' % mh))
                    self.BA = BA; self.b1.setEnabled(True)
        except Exception:
            pass

    def _dang_xem(self):
        return self.BA if (self.ban.checkedId() == 1 and self.BA) else self.B

    def _ve(self):
        from export.xuat_luan_giai import html_ban
        B = self._dang_xem()
        self.xem.setHtml(html_ban(B) if B else '<p>Chưa soạn được bản tổng hợp cho lá số này: %s</p>' % self._loi)

    # ------------------------------------------------------------ AI
    def ai(self):
        if not self.B or self._dang: return
        from ui import khoa_api
        from battu_core.tong_hop import ai_claude as A
        khoa = khoa_api.doc()
        if not khoa:
            if HopKhoaAPI(self).exec() != QDialog.Accepted: return
            khoa = khoa_api.doc()
            if not khoa: return
        mh = _mo_hinh(); vao, ra = A.uoc_luong_token(self.hs, self.pd, self.B)
        if A.doc_dem(A.khoa_dem(self.hs, self.pd, mh)) is None:
            so = lambda n: '{:,}'.format(n).replace(',', '.')
            hoi = QMessageBox.question(self, 'Gửi tới AI?', (
                'Sẽ gửi tứ trụ, năm sinh, giới tính và các kết luận đã tính tới Claude API – mô hình %s. Không gửi họ tên.\n'
                'Lưu ý: từ tứ trụ và năm sinh có thể suy ra ngày giờ sinh gần đúng (trong khoảng 2 giờ).\n'
                'Bản AI chỉ diễn đạt lại; khi khác kết luận của phần mềm thì theo kết luận của phần mềm.\n\n'
                'Ước lượng: khoảng %s token vào, tối đa %s token ra. Giá theo bảng giá hiện hành trên Anthropic Console, '
                'tính vào tài khoản của bạn.\n\nTiếp tục?') % (mh, so(vao), so(ra)))
            if hoi != QMessageBox.Yes: return
        self._dang = True; self.n_ai.setEnabled(False); self.n_ai.setText('⏳  AI đang viết…')
        self.tt.setText('Đang gửi tới AI – thường mất 1–3 phút. Bản không mạng vẫn dùng được trong lúc chờ.')
        hs, pd, B, tieu_de = self.hs, self.pd, self.B, self.B[0][1]

        def chay():
            try: self._bao.xong.emit(('ok',) + A.nang_cao(khoa, hs, pd, B, tieu_de, mh))
            except A.LoiAI as e: self._bao.xong.emit(('loi', str(e)))
            except Exception as e: self._bao.xong.emit(('loi', 'Lỗi không mong đợi: %r' % e))
        threading.Thread(target=chay, daemon=True).start()

    def _ai_xong(self, kq):
        self._dang = False; self.n_ai.setEnabled(bool(self.B)); self.n_ai.setText('✨  Nâng cao bằng AI')
        if kq[0] == 'loi':
            self.tt.setText('AI không chạy được – đang hiện bản không mạng. ' + kq[1]); return
        _, BA, loi, tu_dem = kq
        if BA is None:
            self.tt.setText('Bản AI chưa qua bộ kiểm tra nên KHÔNG dùng (giữ bản không mạng): ' + ' '.join(loi[:4])); return
        self.BA = BA; self.b1.setEnabled(True); self.b1.setChecked(True); self._ve()
        self.tt.setText('Đang xem bản AI%s – đã qua bộ kiểm tra dữ kiện. Bấm "Bản không mạng" để so.' % (' (lưu sẵn trên máy)' if tu_dem else ''))

    # ------------------------------------------------------------ xuất
    def _ten(self, duoi):
        from ui.hop_xuat import _ten_tep
        ai = '_AI' if (self.BA is not None and self._dang_xem() is self.BA) else ''
        return _ten_tep(self.kq, duoi).replace('.' + duoi, '_luan_giai_tong_hop%s.%s' % (ai, duoi))

    def pdf(self):
        B = self._dang_xem()
        if not B: return
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu PDF', self._ten('pdf'), 'PDF (*.pdf)')
        if not p: return
        from export.xuat_luan_giai import xuat_pdf
        n = xuat_pdf(B, p, ten=self.kq.la_so.ho_ten or '')
        QMessageBox.information(self, 'Đã xuất', 'Đã lưu %s (%d trang).' % (p, n))

    def docx(self):
        B = self._dang_xem()
        if not B: return
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu Word', self._ten('docx'), 'Word (*.docx)')
        if not p: return
        from export.xuat_luan_giai import xuat_docx
        xuat_docx(B, p)
        QMessageBox.information(self, 'Đã xuất', 'Đã lưu %s.' % p)
