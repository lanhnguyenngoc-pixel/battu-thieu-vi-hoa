# -*- coding: utf-8 -*-
"""Xuất PDF và IN (YÊU CẦU 5.4): khối kết luận + các mục người dùng chọn (mặc định: KET_LUAN B–K và bảng Ngũ hành thiếu)
+ 2 biểu đồ (năng lực ngũ hành, đại vận); mỗi trang có số trang "Trang x/y", ngày in, phiên bản quy tắc.
Dựng bằng Qt (QTextDocument → QPdfWriter / QPrinter) để dùng đúng phông Inter kèm phần mềm, đủ dấu tiếng Việt.
Cần QApplication (giao diện đã có; dòng lệnh: QT_QPA_PLATFORM=offscreen)."""
import html, os
from PySide6.QtCore import Qt, QSizeF, QRectF, QMarginsF, QUrl, QPoint
from PySide6.QtGui import (QTextDocument, QPdfWriter, QPageSize, QPageLayout, QPainter, QImage, QFont, QColor, QPen,
                           QFontDatabase)
from export.noi_dung import (MUC_XUAT, MUC_MAC_DINH, TEN_PHAN_MEM, khoi, thong_tin_la, chu_ky, ngay_in, chu_o)

CSS = """
body { font-family: 'Inter'; font-size: 8.6pt; color: #1e2140; }
h1 { font-size: 16pt; color: #3730a3; margin: 0 0 4pt 0; }
h2 { font-size: 12pt; color: #3730a3; margin: 14pt 0 5pt 0; border-bottom: 1px solid #c7c2f5; }
h3 { font-size: 10pt; margin: 8pt 0 3pt 0; }
table.tt td { padding: 2pt 6pt; }
table.bang { border-collapse: collapse; }
table.bang td { border: 0.5px solid #d9dbe9; padding: 2.5pt 3.5pt; vertical-align: top; }
td.td { background: #f3f0ff; color: #4338ca; font-weight: bold; }
.kl { background: #f1f3fb; padding: 6pt; }
.cb { color: #92400e; }
.nho { color: #6b7090; font-size: 7.6pt; }
"""


def anh_bieu_do(E, rong=1500, cao=640):
    """Hai biểu đồ như giao diện (ui/bieu_do.py) → QImage."""
    from ui.bieu_do import BieuDoNangLuc, BieuDoDaiVan
    ra = []
    for L in (BieuDoNangLuc, BieuDoDaiVan):
        w = L(); w.dat(E); w.resize(rong // 2, cao // 2)
        img = QImage(rong, cao, QImage.Format_ARGB32); img.fill(QColor('white'))
        p = QPainter(img); p.setRenderHint(QPainter.Antialiasing); p.scale(2, 2); w.render(p, QPoint(0, 0)); p.end()
        ra.append(img)
    return ra


def _e(t): return html.escape(t).replace('\n', '<br>')


def _kl(t, dam=True):
    """Kết luận của công cụ (đậm) + căn cứ sách ở dòng riêng, chữ thường (bản in không có liên kết để bấm)."""
    from ui.trich_dan import tach, chi_la_trich
    noi, cits = tach(t)
    if chi_la_trich(noi, cits): return _e(t)
    h = ('<b>%s</b>' % _e(noi)) if dam else _e(noi)
    if cits: h += '<br><span style="font-weight:normal;color:#5b6078;font-size:8pt">Căn cứ sách: %s</span>' % _e('; '.join(cits))
    return h


def _bang_html(b):
    if not b['dong']: return '<p class="nho">(không có nội dung)</p>'
    tong = sum(b['rong'][c] for c in b['cot']) or 1
    h = ['<table class="bang" width="100%" cellspacing="0">', '<tr>']
    for c in b['cot']: h.append('<td width="%d%%" style="border:none;padding:0"></td>' % max(1, round(100 * b['rong'][c] / tong)))
    h.append('</tr>')
    for r in b['dong']:
        h.append('<tr>')
        for c in b['cot']:
            if (r, c) in b['che']: continue
            t, k = b['o'].get((r, c), ('', {}))
            rs, cs = b['gop'].get((r, c), (1, 1))
            st = []
            if k.get('chu'): st.append('color:%s' % k['chu'])
            if k.get('dam'): st.append('font-weight:bold')
            if k.get('nen'): st.append('background:%s' % k['nen'])
            cls = ' class="td"' if k.get('tieu_de') else ''
            h.append('<td%s%s%s%s>%s</td>' % (cls, ' rowspan="%d"' % rs if rs > 1 else '', ' colspan="%d"' % cs if cs > 1 else '',
                                            ' style="%s"' % ';'.join(st) if st else '', _kl(t, dam=False)))
        h.append('</tr>')
    h.append('</table>')
    return ''.join(h)


def dung_tai_lieu(kq, muc=None, bieu_do=True):
    """QTextDocument chứa toàn bộ nội dung xuất."""
    E = kq.E; muc = MUC_MAC_DINH if muc is None else muc
    doc = QTextDocument(); doc.setDefaultStyleSheet(CSS); doc.setDefaultFont(QFont('Inter', 9))
    o = lambda ref: chu_o(E, 'KET_LUAN', int(ref[1:]), ord(ref[0]) - 64)
    h = ['<html><body>', '<h1>%s – Lá số Tứ trụ</h1>' % TEN_PHAN_MEM, '<table class="tt">']
    for k, v in thong_tin_la(kq): h.append('<tr><td><b>%s</b></td><td>%s</td></tr>' % (_e(k), _e(v)))
    h.append('<tr><td><b>Phiên bản</b></td><td>%s · in ngày %s</td></tr></table>' % (_e(chu_ky()), ngay_in()))
    # khối kết luận (KET_LUAN B2:C12)
    h.append('<h2>Luận giải chung</h2>')
    if o('B2').strip(): h.append('<p class="cb">%s</p>' % _e(o('B2')))
    h.append('<table class="bang" width="100%" cellspacing="0">')
    for r in range(4, 13):
        h.append('<tr><td class="td" width="22%%">%s</td><td>%s</td></tr>' % (_e(o('B%d' % r)), _kl(o('C%d' % r))))
    h.append('</table><p class="nho">Nội dung sức khỏe chỉ để tham khảo, không phải chẩn đoán y khoa.</p>')
    if bieu_do:
        h.append('<h2>Biểu đồ năng lực và đại vận</h2>')
        for i, img in enumerate(anh_bieu_do(E)):
            url = QUrl('bieudo://%d' % i); doc.addResource(QTextDocument.ImageResource, url, img)
            h.append('<p><img src="bieudo://%d" width="640"></p>' % i)
    ten = {k: (t, sh, d) for k, t, sh, d in MUC_XUAT}
    for k in muc:
        if k not in ten: continue
        t, sh, d = ten[k]
        h.append('<h2>%s</h2>' % _e(t))
        if k == 'KL_H' or sh == 'SUC_KHOE': h.append('<p class="cb">Chỉ tham khảo, không phải chẩn đoán y khoa.</p>')
        h.append(_bang_html(khoi(E, sh, d)))
    h.append('</body></html>')
    doc.setHtml(''.join(h))
    return doc


def _ve(doc, thiet_bi, kq, trai=None, phai=None):
    """Phân trang doc lên thiết bị (QPdfWriter/QPrinter) kèm đầu trang, chân trang, số trang."""
    p = QPainter(thiet_bi)
    dpi = thiet_bi.logicalDpiY(); mm = dpi / 25.4
    vung = QRectF(0, 0, thiet_bi.width(), thiet_bi.height())
    dau, chan = 9 * mm, 9 * mm
    doc.documentLayout().setPaintDevice(thiet_bi)
    doc.setPageSize(QSizeF(vung.width(), vung.height() - dau - chan))
    n = doc.pageCount(); f = QFont('Inter', 7); ten = (kq.la_so.ho_ten or '') if kq else ''
    trai = trai or '%s%s' % (TEN_PHAN_MEM, (' · ' + ten) if ten else ''); phai = phai or chu_ky()
    for i in range(n):
        if i: thiet_bi.newPage()
        p.save(); p.setFont(f); p.setPen(QColor('#6b7090'))
        p.drawText(QRectF(0, 0, vung.width(), dau - 2 * mm), Qt.AlignLeft | Qt.AlignVCenter, trai)
        p.drawText(QRectF(0, 0, vung.width(), dau - 2 * mm), Qt.AlignRight | Qt.AlignVCenter, phai)
        p.setPen(QPen(QColor('#c7c2f5'), 0.3 * mm)); p.drawLine(0, dau - 1.5 * mm, vung.width(), dau - 1.5 * mm)
        y = vung.height() - chan
        p.setPen(QColor('#6b7090'))
        p.drawText(QRectF(0, y + 2 * mm, vung.width(), chan - 2 * mm), Qt.AlignLeft | Qt.AlignVCenter, 'In ngày %s' % ngay_in())
        p.drawText(QRectF(0, y + 2 * mm, vung.width(), chan - 2 * mm), Qt.AlignRight | Qt.AlignVCenter, 'Trang %d/%d' % (i + 1, n))
        p.restore()
        p.save(); p.translate(0, dau - i * (vung.height() - dau - chan))
        doc.drawContents(p, QRectF(0, i * (vung.height() - dau - chan), vung.width(), vung.height() - dau - chan))
        p.restore()
    p.end()
    return n


def xuat_pdf(kq, duong_dan, muc=None, bieu_do=True):
    w = QPdfWriter(duong_dan); w.setResolution(150)
    w.setPageSize(QPageSize(QPageSize.A4)); w.setPageMargins(QMarginsF(14, 10, 14, 10), QPageLayout.Millimeter)
    w.setTitle('%s – %s' % (TEN_PHAN_MEM, kq.la_so.ho_ten or 'lá số')); w.setCreator(TEN_PHAN_MEM + ' · ' + chu_ky())
    n = _ve(dung_tai_lieu(kq, muc, bieu_do), w, kq)
    return n


def in_ra(kq, may_in, muc=None, bieu_do=True):
    """In trực tiếp ra QPrinter (đã chọn qua QPrintDialog)."""
    return _ve(dung_tai_lieu(kq, muc, bieu_do), may_in, kq)


def html_ra_pdf(noi_dung_html, duong_dan, tieu_de, anh=None, css=CSS):
    """PDF từ HTML bất kỳ (dùng cho Hướng dẫn sử dụng); anh = {tên: QImage}."""
    doc = QTextDocument(); doc.setDefaultStyleSheet(css); doc.setDefaultFont(QFont('Inter', 9))
    for k, img in (anh or {}).items(): doc.addResource(QTextDocument.ImageResource, QUrl(k), img)
    doc.setHtml(noi_dung_html)
    w = QPdfWriter(duong_dan); w.setResolution(150)
    w.setPageSize(QPageSize(QPageSize.A4)); w.setPageMargins(QMarginsF(16, 12, 16, 12), QPageLayout.Millimeter)
    w.setTitle(tieu_de); w.setCreator(TEN_PHAN_MEM)
    return _ve(doc, w, None, tieu_de, chu_ky())
