# -*- coding: utf-8 -*-
"""Trang NGŨ HÀNH THIẾU dạng trang đọc (bản 1.8.0), tách hai phần:
  * "Kết quả lá số": số liệu, bảng và kết luận mà công cụ tính cho lá số đang xem;
  * "Lý thuyết – cách tính": thang điểm, ngưỡng, cách tính từng tiêu chí (chung cho mọi lá số);
  * "Bảng gốc": lưới ô như workbook (để đối chiếu).
Chỉ ĐỌC ô của sheet NGU_HANH_THIEU (không tính lại gì) – vị trí ô theo workbook 6.7."""
import html, re
from PySide6.QtCore import Signal
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QComboBox, QTabWidget, QTextBrowser,
                               QScrollArea, QSizePolicy)
from export.lam_sach import hien_thi_sach

SH = 'NGU_HANH_THIEU'

CSS = """
body { font-family: 'Inter'; font-size: 10.5pt; color: #1e2140; }
h2 { font-size: 13pt; color: #3730a3; margin: 16px 0 6px 0; }
h3 { font-size: 11pt; color: #4338ca; margin: 12px 0 4px 0; }
p { margin: 4px 0 6px 0; }
table { margin: 4px 0 8px 0; border-collapse: collapse; }
td { padding: 4px 8px; border: 1px solid #e3e5f0; vertical-align: top; }
th { background: #f3f0ff; color: #4338ca; padding: 4px 8px; border: 1px solid #e3e5f0; text-align: left; }
.kl { background: #eef2ff; }
.so { text-align: right; white-space: nowrap; }
.canh { color: #9a3412; }
.nho { color: #6b7086; font-size: 9.5pt; }
"""


def _o(E, ref):
    t = hien_thi_sach(E, SH, ref)
    return '' if t is None else str(t).strip()


def _e(t):
    return html.escape(t).replace('\n', '<br>')


def _bo_tien_to(t):
    """Bỏ chữ 'Cách tính:' / 'Cách tính (…):' ở đầu để đọc như một đoạn lý thuyết."""
    t = re.sub(r'^\s*Cách tính[^:]{0,80}:\s*', '', t)
    return t[:1].upper() + t[1:]


def _pt(t):
    try: return '%g%%' % round(float(t) * 100, 1)
    except (TypeError, ValueError): return t


def _bang(E, dau, dong, cot, so=()):
    H = ['<table width="100%"><tr>' + ''.join('<th>%s</th>' % _e(x) for x in dau) + '</tr>']
    for r in dong:
        o = [_o(E, '%s%d' % (c, r)) for c in cot]
        if not any(o): continue
        H.append('<tr>' + ''.join('<td%s>%s</td>' % (' class="so"' if c in so else '', _e(x)) for c, x in zip(cot, o)) + '</tr>')
    return ''.join(H) + '</table>'


def _dong_kl(nhan, gia_tri):
    return '<tr class="kl"><td width="32%%"><b>%s</b></td><td>%s</td></tr>' % (_e(nhan), gia_tri)


def html_ket_qua(E):
    o = lambda r: _o(E, r)
    H = ['<html><body>']
    H.append('<h2>Kết luận chung</h2><table width="100%">')
    H.append(_dong_kl('Mức', '<b>%s</b>' % _e(o('C32'))))
    H.append(_dong_kl('Nội dung', _e(o('F32'))))
    H.append('</table><p class="nho">%s · Chế độ tính mục I, II: %s</p>' % (_e(o('B5')), _e(o('E3'))))

    H.append('<h2>I. Tìm ngũ hành thiếu – lực phe mình</h2>')
    H.append(_bang(E, ['Tiêu chí', 'Trạng thái', 'Kết quả', 'Điểm'], [7, 8, 9, 10, 11, 12, 13, 14], 'BCDE', so='E'))
    H.append('<table width="100%">')
    H.append(_dong_kl('Tổng điểm', '<b>%s</b> – %s (ngưỡng thiếu ≤ %s, dư ≥ %s)' % (_e(o('E15')), _e(o('F15')), _pt(o('H15')), _pt(o('H16')))))
    H.append(_dong_kl('Ngũ hành thiếu / dư', '%s · %s' % (_e(o('C16')), _e(o('F16')))))
    H.append('</table>')
    if o('B18'): H.append('<p class="canh">%s</p>' % _e(o('B18')))

    H.append('<h2>II. Thân vượng / thân yếu – chấm điểm cộng trừ</h2>')
    H.append(_bang(E, ['Tiêu chí', 'Trạng thái', 'Kết quả', 'Điểm'], range(21, 27), 'BCDE', so='E'))
    H.append('<table width="100%">')
    H.append(_dong_kl('Tổng điểm', '<b>%s</b> – %s' % (_e(o('E27')), _e(o('F27')))))
    H.append(_dong_kl('Hướng bổ sung', _e(o('F28'))))
    H.append(_dong_kl('Kết luận', _e(o('F29'))))
    H.append('</table>')
    if o('B31'): H.append('<p class="canh">%s</p>' % _e(o('B31')))

    H.append('<h2>III. Đại vận – vượng suy – thập thần</h2>')
    dau = [o('%s34' % c) or c for c in 'BCDEFGH']
    H.append(_bang(E, dau, range(35, 43), 'BCDEFGH'))

    H.append('<h2>IV. Năng lực gốc</h2><p class="nho">%s</p>' % _e(o('B46')))
    H.append(_bang(E, ['Năng lực', 'Can đại diện', 'Trạng thái tại chi tháng', 'Tỷ lệ', 'Diễn giải'], range(48, 54), 'BCDEF', so='E'))

    H.append('<h2>V. Điểm đại vận</h2><p class="nho">%s</p>' % _e(o('B68')))
    H.append(_bang(E, ['Đại vận', 'Thập thần can vận', 'Trạng thái tại chi tháng', 'Điểm vận', 'Thực tế', 'Đánh giá'],
                   range(70, 78), 'BCDEGH', so='EG'))
    nguong = ' · '.join('%s %s' % (_e(o('%s69' % c)), _e(o('%s70' % c))) for c in 'IJKLM' if o('%s69' % c))
    if nguong: H.append('<p class="nho">Mốc so sánh: %s</p>' % nguong)
    return ''.join(H) + '</body></html>'


def html_ly_thuyet(E):
    o = lambda r: _o(E, r)
    H = ['<html><body>']
    H.append('<h2>Giới thiệu</h2><p><b>%s</b></p><p>%s</p>' % (_e(o('A1')), _e(o('A2'))))
    H.append('<p>%s %s. %s</p>' % (_e(o('A3')), _e(o('E3')), _e(o('G3'))))
    H.append('<h2>I. Tìm ngũ hành thiếu</h2><p><b>%s</b></p>' % _e(o('A4')))
    H.append('<table width="100%"><tr><th width="28%">Tiêu chí</th><th>Cách tính</th></tr>')
    for r, ten in ((7, 'Được lệnh'), (8, 'Đắc địa'), (13, 'Cùng hành'), (14, 'Được sinh')):
        H.append('<tr><td>%s</td><td>%s</td></tr>' % (ten, _e(_bo_tien_to(o('F%d' % r)))))
    H.append('</table><p>%s</p>' % _e(o('B17')))
    H.append('<h2>II. Thân vượng / thân yếu</h2>')
    H.append('<table width="100%"><tr><th width="28%">Tiêu chí</th><th>Cách tính</th></tr>')
    for r, ten in ((21, 'Đắc lệnh'), (22, 'Được sinh'), (23, 'Bị tiết'), (24, 'Bị khắc'), (25, 'Thông căn'), (26, 'Thông gốc')):
        H.append('<tr><td>%s</td><td>%s</td></tr>' % (ten, _e(_bo_tien_to(o('F%d' % r)))))
    H.append('</table><p>%s</p>' % _e(o('B30')))
    H.append('<h2>III. Đại vận – vượng suy – thập thần</h2><p><b>%s</b></p><p>%s</p>' % (_e(o('A33')), _e(o('B43'))))
    H.append('<h2>IV. Biểu đồ năng lực</h2><p><b>%s</b></p><p>%s</p>' % (_e(o('A45')), _e(o('B54'))))
    H.append('<h2>V. Biểu đồ đại vận</h2><p><b>%s</b></p><p>%s</p>' % (_e(o('A67')), _e(o('B78'))))
    return ''.join(H) + '</body></html>'


def _trinh_doc():
    b = QTextBrowser(); b.setOpenLinks(False); b.document().setDefaultStyleSheet(CSS)
    b.setStyleSheet('QTextBrowser{background:white;border:1px solid #e7e8f3;border-radius:16px;padding:16px 24px;}')
    return b


class TrangNguHanhThieu(QWidget):
    doi_che_do = Signal(str)

    def __init__(self, CHE_DO_NHT, lbl):
        super().__init__()
        from ui.bieu_do import BieuDoNangLuc, BieuDoDaiVan
        from ui.bai_doc import KhungDocLuoi
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(12)
        h = QHBoxLayout(); h.addWidget(lbl('Ngũ hành thiếu', 'tieude')); h.addSpacing(16)
        h.addWidget(lbl('Chế độ tính mục I, II', 'nhan')); self.che_do = QComboBox(); self.che_do.addItems(CHE_DO_NHT)
        self.che_do.setMinimumWidth(190); h.addWidget(self.che_do); h.addStretch()
        h.addWidget(lbl('Phương pháp riêng – chỉ tham khảo', 'badge')); v.addLayout(h)
        self.tom = lbl('', 'chip_info', True); v.addWidget(self.tom)
        self.tab = QTabWidget(); v.addWidget(self.tab, 1)
        # Tab 1 – kết quả lá số (biểu đồ + bài đọc có bảng)
        t1 = QWidget(); l = QVBoxLayout(t1); l.setContentsMargins(0, 8, 0, 0); l.setSpacing(10)
        hb = QHBoxLayout(); hb.setSpacing(14)
        c1 = QFrame(); c1.setObjectName('the'); l1 = QVBoxLayout(c1); l1.setContentsMargins(12, 12, 12, 12)
        self.bd1 = BieuDoNangLuc(); l1.addWidget(self.bd1); hb.addWidget(c1, 4)
        c2 = QFrame(); c2.setObjectName('the'); l2 = QVBoxLayout(c2); l2.setContentsMargins(12, 12, 12, 12)
        self.bd2 = BieuDoDaiVan(); l2.addWidget(self.bd2); hb.addWidget(c2, 6)
        khung = QWidget(); khung.setLayout(hb); khung.setMinimumHeight(250); khung.setMaximumHeight(320)
        l.addWidget(khung); self.kq_doc = _trinh_doc(); l.addWidget(self.kq_doc, 1)
        self.tab.addTab(t1, 'Kết quả lá số')
        # Tab 2 – lý thuyết
        self.lt_doc = _trinh_doc(); self.tab.addTab(self.lt_doc, 'Lý thuyết – cách tính')
        # Tab 3 – bảng gốc
        self.bang = KhungDocLuoi(SH); self.tab.addTab(self.bang, 'Bảng gốc (như workbook)')
        self.tab.currentChanged.connect(lambda _: self._ve_bang())
        self.che_do.currentTextChanged.connect(self.doi_che_do.emit)
        self.kq = None

    def hien(self, kq):
        self.kq = kq; E = kq.E
        self.che_do.blockSignals(True); self.che_do.setCurrentText(kq.la_so.che_do_nht); self.che_do.blockSignals(False)
        o = lambda r: _o(E, r)
        self.tom.setText('Mục I: %s → %s · Mục II: %s → %s · Kết luận chung: %s' % (o('E15'), o('F15'), o('E27'), o('F27'), o('C32')))
        self.bd1.dat(E); self.bd2.dat(E)
        self.kq_doc.setHtml(html_ket_qua(E)); self.lt_doc.setHtml(html_ly_thuyet(E))
        self._ve_bang()

    def _ve_bang(self):
        if self.kq and self.isVisible() and self.tab.currentWidget() is self.bang: self.bang.hien(self.kq.E)

    def showEvent(self, e):
        super().showEvent(e); self._ve_bang()
