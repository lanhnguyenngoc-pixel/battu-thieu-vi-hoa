# -*- coding: utf-8 -*-
"""Hai biểu đồ của NGU_HANH_THIEU, vẽ ĐÚNG định nghĩa trong workbook 6.3 (xl/charts/chart1.xml, chart2.xml):
  * Biểu đồ năng lực: thanh NGANG, nhóm B48:B53, giá trị E48:E53 (định dạng 0\\%), trục 0–100 bước 10,
    nhóm đầu ở trên (orientation maxMin), nhãn giá trị ngoài đầu thanh, màu từng thanh như workbook, không chú giải.
  * Biểu đồ đại vận: ĐƯỜNG, nhóm A70:A77; 7 chuỗi E, G (Điểm đại vận, Thực tế – có điểm tròn) và I:M (5 đường ngang
    70 / 45 / 0 / −30 / −70); tên chuỗi = ô dòng 69; trục −100…100 bước 10; chú giải bên phải; không làm mượt.
Vẽ bằng QPainter (dùng lại được cho xuất PDF ở GĐ6). Dữ liệu chỉ đọc từ ô – giao diện không tự tính."""
from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import QPainter, QColor, QPen, QFont, QBrush, QFontMetrics
from PySide6.QtWidgets import QWidget

MAU_THANH = ['#ff6384', '#ff9f40', '#ffcd56', '#4bc0c0', '#36a2eb', '#9966ff']      # chart1 dPt 0..5
CHUOI_DV = [('E', '#ff6384', 2.2, True), ('G', '#ffcd56', 2.2, True), ('I', '#b9e0dd', 1.2, False),
            ('J', '#36a2eb', 1.7, False), ('K', '#d9c8f5', 1.2, False), ('L', '#c8c8c8', 1.2, False),
            ('M', '#7f7f7f', 1.7, False)]                                                # chart2: cột, màu, độ dày (pt), điểm


def _so(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0


class BieuDoNangLuc(QWidget):
    def __init__(self):
        super().__init__(); self.setMinimumSize(520, 330); self.nhan = []; self.gt = []

    def dat(self, E):
        self.nhan = [str(E.gia_tri('NGU_HANH_THIEU', 'B%d' % r) or '') for r in range(48, 54)]
        self.gt = [_so(E.gia_tri('NGU_HANH_THIEU', 'E%d' % r)) for r in range(48, 54)]
        self.update()

    def paintEvent(self, ev):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        W, H = self.width(), self.height()
        p.fillRect(self.rect(), QColor('white'))
        f = QFont('Inter', 10); p.setFont(f); fm = QFontMetrics(f)
        tf = QFont('Inter', 12); tf.setBold(True); p.setFont(tf)
        p.drawText(QRectF(0, 6, W, 24), Qt.AlignCenter, 'Biểu đồ năng lực')
        p.setFont(f)
        lw = max([fm.horizontalAdvance(t) for t in self.nhan] + [80]) + 14
        x0, x1, y0, y1 = lw, W - 46, 40, H - 30
        p.setPen(QPen(QColor('#e5e7eb'), 1))
        for k in range(0, 101, 10):
            x = x0 + (x1 - x0) * k / 100; p.drawLine(QPointF(x, y0), QPointF(x, y1))
            p.setPen(QColor('#6b7280')); p.drawText(QRectF(x - 20, y1 + 4, 40, 16), Qt.AlignCenter, '%d%%' % k)
            p.setPen(QPen(QColor('#e5e7eb'), 1))
        n = max(len(self.gt), 1); bh = (y1 - y0) / n
        for i, (t, v) in enumerate(zip(self.nhan, self.gt)):
            y = y0 + i * bh; h = bh / 1.4   # gapWidth 40
            yy = y + (bh - h) / 2
            p.setPen(QColor('#1f2937')); p.drawText(QRectF(0, y, lw - 8, bh), Qt.AlignRight | Qt.AlignVCenter, t)
            w = (x1 - x0) * max(0.0, min(v, 100.0)) / 100
            p.setPen(Qt.NoPen); p.setBrush(QColor(MAU_THANH[i % 6])); p.drawRect(QRectF(x0, yy, w, h))
            p.setPen(QColor('#000000')); p.drawText(QRectF(x0 + w + 4, y, 60, bh), Qt.AlignLeft | Qt.AlignVCenter, '%d%%' % round(v))
        p.setPen(QPen(QColor('#9ca3af'), 1)); p.drawLine(QPointF(x0, y0), QPointF(x0, y1))
        p.end()


class BieuDoDaiVan(QWidget):
    def __init__(self):
        super().__init__(); self.setMinimumSize(640, 360); self.nhom = []; self.chuoi = []

    def dat(self, E):
        self.nhom = [str(E.gia_tri('NGU_HANH_THIEU', 'A%d' % r) or '') for r in range(70, 78)]
        self.chuoi = [(str(E.gia_tri('NGU_HANH_THIEU', '%s69' % col) or col), [_so(E.gia_tri('NGU_HANH_THIEU', '%s%d' % (col, r))) for r in range(70, 78)],
                       mau, w, diem) for col, mau, w, diem in CHUOI_DV]
        self.update()

    def paintEvent(self, ev):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        W, H = self.width(), self.height()
        p.fillRect(self.rect(), QColor('white'))
        f = QFont('Inter', 9); p.setFont(f); fm = QFontMetrics(f)
        tf = QFont('Inter', 12); tf.setBold(True); p.setFont(tf)
        p.drawText(QRectF(0, 6, W - 150, 24), Qt.AlignCenter, 'Biểu đồ đại vận')
        p.setFont(f)
        x0, x1, y0, y1 = 46, W - 170, 38, H - 30
        def Y(v): return y1 - (y1 - y0) * (max(-100.0, min(100.0, v)) + 100) / 200
        for k in range(-100, 101, 10):
            y = Y(k); p.setPen(QPen(QColor('#d1d5db' if k == 0 else '#eef0f4'), 1)); p.drawLine(QPointF(x0, y), QPointF(x1, y))
            p.setPen(QColor('#6b7280')); p.drawText(QRectF(0, y - 8, x0 - 6, 16), Qt.AlignRight | Qt.AlignVCenter, str(k))
        n = max(len(self.nhom), 1); step = (x1 - x0) / n
        def X(i): return x0 + step * (i + 0.5)
        p.setPen(QColor('#374151'))
        for i, t in enumerate(self.nhom): p.drawText(QRectF(X(i) - 30, y1 + 6, 60, 16), Qt.AlignCenter, t)
        # đường ngang trước, hai chuỗi chính vẽ sau (nổi lên trên)
        for ten, vals, mau, w, diem in reversed(self.chuoi):
            pen = QPen(QColor(mau), w * 1.33); p.setPen(pen); p.setBrush(Qt.NoBrush)
            pts = [QPointF(X(i), Y(v)) for i, v in enumerate(vals)]
            for a, b in zip(pts, pts[1:]): p.drawLine(a, b)
            if diem:
                p.setBrush(QColor(mau))
                for q in pts: p.drawEllipse(q, 3.6, 3.6)
        # chú giải bên phải
        ly = y0 + 10
        for ten, vals, mau, w, diem in self.chuoi:
            p.setPen(QPen(QColor(mau), w * 1.33)); p.drawLine(QPointF(x1 + 18, ly), QPointF(x1 + 42, ly))
            if diem: p.setBrush(QColor(mau)); p.drawEllipse(QPointF(x1 + 30, ly), 3.2, 3.2)
            p.setPen(QColor('#1f2937')); p.drawText(QRectF(x1 + 48, ly - 9, 140, 18), Qt.AlignLeft | Qt.AlignVCenter, ten)
            ly += 22
        p.end()
