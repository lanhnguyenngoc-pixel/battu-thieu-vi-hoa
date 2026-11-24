# -*- coding: utf-8 -*-
"""Trang "Sức khỏe 10 năm" (bản 1.9.4): bảng sức khỏe 10 năm kể từ năm đặt xem và bốn kiểu lá số dễ gặp tai ương.
Chỉ trình bày lại dữ kiện đã tính (battu_core.tong_hop.tai_uong); không đổi kết luận nào của công cụ.
Mọi chỗ nói về bốn kiểu đều kèm nhãn: "cách luận tham khảo thêm thuộc dễ gặp tai ương nhất."."""
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextBrowser, QFileDialog, QMessageBox


class TrangSucKhoe(QWidget):
    def __init__(self):
        super().__init__(); self.kq = None; self.B = None
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); t = QLabel('Sức khỏe 10 năm'); t.setObjectName('tieude'); h.addWidget(t); h.addStretch()
        self.n_pdf = QPushButton('⤓  PDF'); self.n_pdf.setObjectName('nut_phu'); self.n_doc = QPushButton('⤓  Word')
        self.n_doc.setObjectName('nut_phu')
        for n in (self.n_pdf, self.n_doc): h.addWidget(n)
        b = QLabel('Chỉ tham khảo – không phải chẩn đoán y khoa'); b.setObjectName('badge'); h.addWidget(b); v.addLayout(h)
        self.tt = QLabel(''); self.tt.setObjectName('chip_info'); self.tt.setWordWrap(True); v.addWidget(self.tt)
        self.xem = QTextBrowser(); self.xem.setStyleSheet('QTextBrowser{background:white;border:1px solid #e7e8f3;border-radius:16px;padding:16px 24px;}')
        v.addWidget(self.xem, 1)
        self.n_pdf.clicked.connect(lambda: self._xuat('pdf')); self.n_doc.clicked.connect(lambda: self._xuat('docx'))

    def hien(self, kq):
        from battu_core.tong_hop.du_kien import ho_so
        from battu_core.tong_hop import tai_uong as T
        from export.xuat_luan_giai import html_ban
        self.kq = kq
        try:
            hs = ho_so(kq.E, kq.la_so, so_nam=10); self.B = T.khoi_bai_doc(hs, 10); p = T.phan_tich(hs)
            n = sum(k['co'] for k in p['kieu'])
            self.tt.setText('Năm đặt xem %s: bảng 10 năm liên tiếp. Bốn kiểu dễ gặp tai ương: lá số dính %d/4 (%s)' % (hs['la_so'].get('nam_xem'), n, T.NHAN.rstrip('.')))
            self.xem.setHtml(html_ban(self.B))
            for b in (self.n_pdf, self.n_doc): b.setEnabled(True)
        except Exception as e:                                   # lá số thiếu dữ kiện
            self.B = None
            for b in (self.n_pdf, self.n_doc): b.setEnabled(False)
            self.tt.setText(''); self.xem.setHtml('<p>Chưa dựng được bảng cho lá số này: %s</p>' % e)

    def _xuat(self, loai):
        if not self.B: return
        from export.xuat_luan_giai import xuat_pdf, xuat_docx
        ten = (getattr(self.kq.la_so, 'ho_ten', '') or 'la_so').strip().replace(' ', '_')
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu', 'suc_khoe_10_nam_%s.%s' % (ten, loai), 'PDF (*.pdf)' if loai == 'pdf' else 'Word (*.docx)')
        if not p: return
        try:
            if loai == 'pdf': xuat_pdf(self.B, p, 'Sức khỏe 10 năm', getattr(self.kq.la_so, 'ho_ten', '') or '')
            else: xuat_docx(self.B, p, 'Sức khỏe 10 năm')
            QMessageBox.information(self, 'Đã lưu', p)
        except Exception as e:
            QMessageBox.warning(self, 'Không lưu được', str(e))
