# -*- coding: utf-8 -*-
"""Trang "Luận giải tổng hợp" (đề xuất hoàn thiện 15/11/2026, giai đoạn T4): bản ~5 trang A4 không cần mạng,
soạn từ hồ sơ dữ kiện + lớp phán đoán (battu_core.tong_hop); xuất PDF hoặc Word (.docx)."""
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextBrowser, QFileDialog, QMessageBox


class TrangTongHop(QWidget):
    def __init__(self):
        super().__init__(); self.kq = None; self.B = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); t = QLabel('Luận giải tổng hợp'); t.setObjectName('tieude'); h.addWidget(t); h.addStretch()
        self.n_pdf = QPushButton('📄  Xuất PDF'); self.n_pdf.setObjectName('nut_chinh')
        self.n_doc = QPushButton('📝  Xuất Word'); self.n_doc.setObjectName('nut_phu')
        for n in (self.n_pdf, self.n_doc): n.setEnabled(False); h.addWidget(n)
        v.addLayout(h)
        mo = QLabel('Bản tổng hợp khoảng 5 trang A4, 10 lĩnh vực: mỗi việc một kết luận đã gỡ mâu thuẫn, kèm lý do. '
                    'Quy tắc tổng hợp: battu_core/data/quy_tac_tong_hop.json. Chạy không cần mạng.')
        mo.setObjectName('mo'); mo.setWordWrap(True); v.addWidget(mo)
        self.xem = QTextBrowser(); self.xem.setOpenExternalLinks(False); v.addWidget(self.xem, 1)
        self.n_pdf.clicked.connect(self.pdf); self.n_doc.clicked.connect(self.docx)

    def hien(self, kq):
        from battu_core.tong_hop.viet_ban import ban_luan_giai
        from export.xuat_luan_giai import html_ban
        self.kq = kq
        try:
            self.B, _, _ = ban_luan_giai(kq)
            self.xem.setHtml(html_ban(self.B))
            ok = True
        except Exception as e:                                   # lá số thiếu dữ kiện (vd chưa đủ tứ trụ)
            self.B = None; self.xem.setHtml('<p>Chưa soạn được bản tổng hợp cho lá số này: %s</p>' % e); ok = False
        self.n_pdf.setEnabled(ok); self.n_doc.setEnabled(ok)

    def _ten(self, duoi):
        from ui.hop_xuat import _ten_tep
        return _ten_tep(self.kq, duoi).replace('.' + duoi, '_luan_giai_tong_hop.' + duoi)

    def pdf(self):
        if not self.B: return
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu PDF', self._ten('pdf'), 'PDF (*.pdf)')
        if not p: return
        from export.xuat_luan_giai import xuat_pdf
        n = xuat_pdf(self.B, p, ten=self.kq.la_so.ho_ten or '')
        QMessageBox.information(self, 'Đã xuất', 'Đã lưu %s (%d trang).' % (p, n))

    def docx(self):
        if not self.B: return
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu Word', self._ten('docx'), 'Word (*.docx)')
        if not p: return
        from export.xuat_luan_giai import xuat_docx
        xuat_docx(self.B, p)
        QMessageBox.information(self, 'Đã xuất', 'Đã lưu %s.' % p)
