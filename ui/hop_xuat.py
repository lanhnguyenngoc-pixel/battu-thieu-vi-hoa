# -*- coding: utf-8 -*-
"""Hộp thoại Xuất / In (YÊU CẦU 5.4): chọn mục, kèm biểu đồ; PDF, XLSX, JSON, In."""
import os, re
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QCheckBox,
                               QPushButton, QFileDialog, QMessageBox, QApplication)
from export.noi_dung import MUC_XUAT, MUC_MAC_DINH


def _ten_tep(kq, duoi):
    ten = kq.la_so.ho_ten or 'la_so'
    ten = re.sub(r'[\\/:*?"<>|]+', '_', ten).strip() or 'la_so'
    return os.path.join(os.path.expanduser('~'), '%s.%s' % (ten, duoi))


class HopXuat(QDialog):
    def __init__(self, cha, kq):
        super().__init__(cha); self.kq = kq
        self.setWindowTitle('Xuất / In lá số'); self.resize(560, 640)
        v = QVBoxLayout(self); v.setSpacing(10)
        t = QLabel('Bản PDF / in luôn có: thông tin lá số, Luận giải chung, số trang, ngày in, phiên bản quy tắc.\n'
                   'Chọn thêm các mục dưới đây (mặc định: các mục theo lĩnh vực B–K và bảng Ngũ hành thiếu).')
        t.setWordWrap(True); v.addWidget(t)
        hh = QHBoxLayout(); a = QPushButton('Chọn tất cả'); b = QPushButton('Bỏ chọn'); c = QPushButton('Mặc định')
        for x in (a, b, c): x.setObjectName('nut_phu'); hh.addWidget(x)
        hh.addStretch(); v.addLayout(hh)
        self.ds = QListWidget(); v.addWidget(self.ds, 1)
        for k, ten, sh, _ in MUC_XUAT:
            it = QListWidgetItem(ten); it.setData(Qt.UserRole, k)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable); it.setCheckState(Qt.Checked if k in MUC_MAC_DINH else Qt.Unchecked)
            self.ds.addItem(it)
        a.clicked.connect(lambda: self._dat(lambda k: True)); b.clicked.connect(lambda: self._dat(lambda k: False))
        c.clicked.connect(lambda: self._dat(lambda k: k in MUC_MAC_DINH))
        self.bd = QCheckBox('Kèm 2 biểu đồ (năng lực ngũ hành, đại vận)'); self.bd.setChecked(True); v.addWidget(self.bd)
        h = QHBoxLayout()
        for ten, f in (('📄  PDF', self.pdf), ('🖨  In', self.in_), ('📊  XLSX (bố cục bảng)', self.xlsx), ('{ }  JSON', self.json_)):
            n = QPushButton(ten); n.setObjectName('nut_chinh' if 'PDF' in ten else 'nut_phu'); n.clicked.connect(f); h.addWidget(n)
        v.addLayout(h)
        self.tt = QLabel(''); self.tt.setWordWrap(True); v.addWidget(self.tt)

    def _dat(self, f):
        for i in range(self.ds.count()):
            it = self.ds.item(i); it.setCheckState(Qt.Checked if f(it.data(Qt.UserRole)) else Qt.Unchecked)

    def muc(self):
        return [self.ds.item(i).data(Qt.UserRole) for i in range(self.ds.count()) if self.ds.item(i).checkState() == Qt.Checked]

    def _lam(self, ham, duong_dan, ten):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            ra = ham()
        except Exception as e:
            QApplication.restoreOverrideCursor(); QMessageBox.warning(self, 'Lỗi', 'Không xuất được: %s' % e); return
        QApplication.restoreOverrideCursor()
        self.tt.setText('✔ Đã xuất %s: %s%s' % (ten, duong_dan, (' (%d trang)' % ra) if isinstance(ra, int) else ''))

    def pdf(self):
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu PDF', _ten_tep(self.kq, 'pdf'), 'PDF (*.pdf)')
        if not p: return
        from export.xuat_pdf import xuat_pdf
        self._lam(lambda: xuat_pdf(self.kq, p, self.muc(), self.bd.isChecked()), p, 'PDF')

    def xlsx(self):
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu XLSX', _ten_tep(self.kq, 'xlsx'), 'Excel (*.xlsx)')
        if not p: return
        from export.xuat_xlsx import xuat_xlsx
        self._lam(lambda: xuat_xlsx(self.kq, p), p, 'XLSX')

    def json_(self):
        p, _ = QFileDialog.getSaveFileName(self, 'Lưu JSON', _ten_tep(self.kq, 'json'), 'JSON (*.json)')
        if not p: return
        from export.xuat_json import xuat_json
        self._lam(lambda: xuat_json(self.kq, p), p, 'JSON')

    def in_(self):
        from PySide6.QtPrintSupport import QPrinter, QPrintDialog
        pr = QPrinter(QPrinter.HighResolution)
        if QPrintDialog(pr, self).exec() != QDialog.Accepted: return
        from export.xuat_pdf import in_ra
        self._lam(lambda: in_ra(self.kq, pr, self.muc(), self.bd.isChecked()), pr.printerName(), 'bản in')
