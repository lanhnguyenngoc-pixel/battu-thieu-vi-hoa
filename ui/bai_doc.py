# -*- coding: utf-8 -*-
"""Trình bày một khối kết quả dạng VĂN BẢN ĐỌC (chế độ thường) thay cho lưới ô:
  * dòng chỉ có một ô chữ  → tiêu đề (nếu là tiêu đề của bảng) hoặc đoạn văn;
  * dòng có 2 ô            → đoạn "Hạng mục: nội dung";
  * các dòng liền nhau có ≥ 3 ô → bảng gọn (dòng đầu là đầu bảng nếu là chữ tiêu đề), không có số dòng / chữ cột;
  * giữ màu nhấn (chữ đỏ đậm khi tốt…) và chú giải thuật ngữ khi rê chuột lên chữ gạch chân.
Chữ lấy từ export.noi_dung.khoi (đã bỏ cột/dòng ẩn và tham chiếu nội bộ). Lưới ô gốc vẫn có ở chế độ chuyên gia."""
import html, re
from PySide6.QtCore import QUrl
from PySide6.QtGui import QCursor
from PySide6.QtWidgets import QTextBrowser, QToolTip
from battu_core.du_lieu import model
from battu_core.thuat_ngu import tu_dien
from export.noi_dung import khoi
from ui.trich_dan import TrichDan, tach

CSS = """
body { font-family: 'Inter'; font-size: 10.5pt; color: #1e2140; }
h2 { font-size: 13pt; color: #3730a3; margin: 14px 0 6px 0; }
h3 { font-size: 11pt; color: #4338ca; margin: 12px 0 4px 0; }
p { margin: 4px 0 6px 0; }
.nhan { font-weight: bold; color: #4f46e5; }
.kl { background: #f1f3fb; }
a { color: #5b21b6; text-decoration: underline; }
table { margin: 4px 0 10px 0; }
td { padding: 4px 7px; border: 1px solid #e3e5f0; vertical-align: top; }
th { background: #f3f0ff; color: #4338ca; padding: 4px 7px; border: 1px solid #e3e5f0; text-align: left; }
"""


def _kieu(k):
    st = []
    if k.get('chu'): st.append('color:%s' % k['chu'])
    if k.get('dam') and not k.get('tieu_de'): st.append('font-weight:bold')
    if k.get('nen'): st.append('background:%s' % k['nen'])
    return ';'.join(st)


class BaiDoc(QTextBrowser):
    def __init__(self):
        super().__init__(); self.setOpenLinks(False); self.document().setDefaultStyleSheet(CSS)
        self.setStyleSheet('QTextBrowser{background:white;border:1px solid #e7e8f3;border-radius:16px;padding:16px 24px;}')
        self.highlighted.connect(self._goi_y); self.anchorClicked.connect(self._bam); self.td = []
        self.tdan = TrichDan(); self._lai = None

    def _bam(self, u):
        if self.tdan.bam(u):
            if self._lai: self._lai()
            return
        self._goi_y(u)

    def _goi_y(self, u):
        s = u.toString() if isinstance(u, QUrl) else str(u)
        if s.startswith('cd:'): return
        if not s.startswith('tn:'): QToolTip.hideText(); return
        k, cau, o = self.td[int(s[3:])]
        QToolTip.showText(QCursor.pos(), '<b>%s</b><br>%s' % (html.escape(k), html.escape(cau)), self)

    def _lk(self, t):
        s = html.escape(t)
        for i, (k, cau, o) in enumerate(self.td):
            m = re.search(r'(?<![\wÀ-ỹ])(' + re.escape(html.escape(k)) + r')(?![\wÀ-ỹ])', s, re.I)
            if m and '<a ' not in s[max(0, m.start() - 60):m.start()]:
                s = s[:m.start()] + '<a href="tn:%d">%s</a>' % (i, m.group(1)) + s[m.end():]
        return s.replace('\n', '<br>')

    def hien(self, E, sheet, dong=None, bo_dong=(), mo_dau=''):
        """Vẽ khối (sheet, dòng) dạng bài đọc; bo_dong: các dòng không vẽ (đã hiện ở chỗ khác).
        Nội dung công cụ luận giải in đậm; trích dẫn sách tách ra dòng riêng, ẩn sau liên kết "Căn cứ sách"."""
        if self._lai is None or getattr(self, '_khoa', None) != (id(E), sheet):
            self.tdan.mo = set(); self._khoa = (id(E), sheet)
        self._lai = lambda d0=dong: self.hien(E, sheet, d0, bo_dong, mo_dau)
        self.td = tu_dien(E); b = khoi(E, sheet, dong); o = b['o']; self.tdan.bat_dau()
        fc = model().formula_cells[model().si[sheet]]
        H = ['<html><body>', mo_dau(self) if callable(mo_dau) else mo_dau]; bang = []
        tdan = self.tdan

        def o_bang(t):
            noi, cits = tach(t)
            return self._lk(noi) + (('<br>' + tdan.khoi(cits)) if cits else '')

        def o_dong(r):
            return [(c, o[(r, c)][0], o[(r, c)][1], b['gop'].get((r, c), (1, 1))[1]) for c in b['cot'] if (r, c) in o and o[(r, c)][0].strip()]

        def xa_bang():
            if not bang: return
            dau = bang[0]; la_dau = len(bang) > 1 and len(dau[1]) >= 3 and all(o[(dau[0], c)][1].get('tieu_de') or (dau[0], c) not in fc for c, *_ in dau[1]) \
                and all((dau[0], c) not in fc for c, *_ in dau[1])
            rong = max(sum(n for *_, n in cs) for r, cs in bang)
            H.append('<table cellspacing="0" width="100%">')
            for j, (r, cs) in enumerate(bang):
                the = 'th' if (j == 0 and la_dau) else 'td'
                cs = list(cs); thieu = rong - sum(n for *_, n in cs)
                if thieu > 0: c_, t_, k_, n_ = cs[-1]; cs[-1] = (c_, t_, k_, n_ + thieu)
                H.append('<tr>' + ''.join('<%s%s%s>%s</%s>' % (the, (' colspan="%d"' % n) if n > 1 else '',
                                                              (' style="%s"' % ((_kieu(k) + ';' if _kieu(k) else '') + ('font-weight:600;color:#3b4064' if (i == 0 and the == 'td' and len(cs) > 1) else ''))), o_bang(t), the)
                                          for i, (c, t, k, n) in enumerate(cs)) + '</tr>')
            H.append('</table>'); bang.clear()
        dong = [(r, o_dong(r)) for r in b['dong'] if r not in bo_dong]
        dong = [(r, cs) for r, cs in dong if cs]
        dai = lambda cs: len(cs) == 2 and len(cs[1][1]) > 220                 # hạng mục có nội dung dài → tiêu đề + đoạn văn
        for j, (r, cs) in enumerate(dong):
            ke = (j + 1 < len(dong) and len(dong[j + 1][1]) >= 2 and not dai(dong[j + 1][1])) or \
                 (j > 0 and len(dong[j - 1][1]) >= 2 and not dai(dong[j - 1][1]))
            if len(cs) >= 3 or (len(cs) == 2 and ke and not dai(cs)):   # các dòng 2–3+ ô liền nhau → một bảng liền
                bang.append((r, cs)); continue
            xa_bang()
            if len(cs) == 1:
                c, t, k, n = cs[0]
                if k.get('tieu_de'): H.append('<h3>%s</h3>' % html.escape(t))
                else: H.append('<p>%s</p>' % tdan.html(t, self._lk, kieu=_kieu(k)))
            else:
                (c1, t1, k1, _), (c2, t2, k2, _) = cs
                if dai(cs):
                    H.append('<h3>%s</h3><p>%s</p>' % (html.escape(t1.rstrip(':')), tdan.html(t2, self._lk, kieu=_kieu(k2)))); continue
                H.append('<p><span class="nhan">%s:</span> %s</p>' % (html.escape(t1.rstrip(':')), tdan.html(t2, self._lk, kieu=_kieu(k2))))
        xa_bang(); H.append('</body></html>')
        y = self.verticalScrollBar().value(); self.setHtml(''.join(H)); self.verticalScrollBar().setValue(y)


from PySide6.QtWidgets import QStackedWidget


class KhungDocLuoi(QStackedWidget):
    """Một sheet: chế độ thường → bài đọc; chế độ chuyên gia → lưới ô gốc (BangSheet)."""
    def __init__(self, sheet):
        super().__init__(); from ui.bang_sheet import BangSheet
        self.sheet = sheet; self.doc = BaiDoc(); self.luoi = BangSheet(sheet)
        self.addWidget(self.doc); self.addWidget(self.luoi)

    def hien(self, E):
        from ui.bang_sheet import CHE_DO
        if CHE_DO['chuyen_gia']: self.luoi.hien(E); self.setCurrentIndex(1)
        elif self.sheet == 'DAI_VAN':
            # chế độ thường: kết luận đại vận (mục F) ở đầu; ẩn bảng chấm điểm chi tiết (chỉ có ở chế độ chuyên gia)
            from ui.trich_dan import dai_van_html
            self.doc.hien(E, 'DAI_VAN', bo_dong=set(range(45, 75)) | set(range(95, 130)),
                          mo_dau=lambda bd: dai_van_html(E, bd.tdan, bd._lk)); self.setCurrentIndex(0)
        else: self.doc.hien(E, self.sheet); self.setCurrentIndex(0)
