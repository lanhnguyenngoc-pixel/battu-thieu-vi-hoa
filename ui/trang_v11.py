# -*- coding: utf-8 -*-
"""Giao diện bản 1.1.0 (duyệt 14/11/2026): Lá số trực quan, Luận giải dạng bài đọc + chú giải thuật ngữ,
Kiểm chứng sự việc (nhập SU_KIEN B6:D35). Mọi chữ lấy nguyên văn từ lõi tính / workbook; giao diện không tự tính quy tắc
(màu ngũ hành chỉ để hiển thị, đọc từ chữ hành trong ô)."""
import html, re
from PySide6.QtCore import Qt, Signal, QUrl
from PySide6.QtGui import QCursor, QColor
from PySide6.QtWidgets import (QWidget, QFrame, QLabel, QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout, QLineEdit,
                               QComboBox, QScrollArea, QTableWidget, QTableWidgetItem, QTextBrowser, QToolTip,
                               QAbstractItemView, QHeaderView)
from battu_core.du_lieu import model, ref_of
from battu_core.thuat_ngu import tu_dien, tim
from battu_core.dinh_dang_dk import kieu_sheet
from battu_core.nhap_lieu import SO_DONG_SU_KIEN
from export.noi_dung import chu_o, khoi, thong_tin_la

MAU = {'Mộc': '#16a34a', 'Hỏa': '#e11d48', 'Thổ': '#a0522d', 'Kim': '#b8860b', 'Thủy': '#2563eb'}


def _hanh(t):
    for k in MAU:
        if k in t: return k
    return None


def lbl(text='', name=None, wrap=False):
    l = QLabel(text)
    if name: l.setObjectName(name)
    l.setWordWrap(wrap); l.setTextInteractionFlags(Qt.TextSelectableByMouse)
    return l


def cuon(w):
    sc = QScrollArea(); sc.setWidgetResizable(True); sc.setWidget(w); sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    return sc


def goi_y(t, td):
    """Chú giải (nguyên văn workbook) cho các thuật ngữ có trong chữ t."""
    from ui.bang_sheet import CHE_DO
    h = tim(t, td)
    return '\n\n'.join(('• %s\n   (%s)' % (cau, o)) if CHE_DO['chuyen_gia'] else ('• ' + cau) for k, cau, o in h)


# ---------------------------------------------------------------- LÁ SỐ TRỰC QUAN
class TheTru(QFrame):
    def __init__(self):
        super().__init__(); self.setObjectName('tru')
        v = QVBoxLayout(self); v.setContentsMargins(14, 12, 14, 12); v.setSpacing(4)
        self.ten = lbl('', 'tru_ten'); self.ten.setAlignment(Qt.AlignCenter)
        self.tt_can = lbl('', 'than'); self.tt_can.setAlignment(Qt.AlignCenter)
        self.can = QLabel(); self.can.setAlignment(Qt.AlignCenter)
        self.chi = QLabel(); self.chi.setAlignment(Qt.AlignCenter)
        self.am_duong = lbl('', 'mo'); self.am_duong.setAlignment(Qt.AlignCenter)
        self.tang = lbl('', None, True); self.tang.setAlignment(Qt.AlignCenter)
        self.ts = lbl('', 'mo', True); self.ts.setAlignment(Qt.AlignCenter)
        for w in (self.ten, self.tt_can, self.can, self.am_duong, self.chi, self.tang, self.ts): v.addWidget(w)

    def dat(self, ten, can_tt, hanh_can, chi, tang, ts, nhat):
        self.setObjectName('tru_nhat' if nhat else 'tru'); self.style().unpolish(self); self.style().polish(self)
        self.ten.setText(ten.upper())
        can, _, tt = can_tt.partition(' – ')
        self.tt_can.setText(tt)
        hc = _hanh(hanh_can) or ''
        self.can.setText("<span style='font-size:40px;font-weight:800;color:%s'>%s</span>" % (MAU.get(hc, '#1e2140'), html.escape(can)))
        self.am_duong.setText(hanh_can)
        ten_chi, _, hchi = chi.partition(' (')
        hchi = _hanh(hchi) or ''
        self.chi.setText("<span style='font-size:40px;font-weight:800;color:%s'>%s</span><br><span style='color:#7b809c'>%s</span>"
                         % (MAU.get(hchi, '#1e2140'), html.escape(ten_chi), hchi))
        self.tang.setText('<br>'.join("<span style='color:%s'>%s</span>" % (MAU.get(_hanh(x) or '', '#3b4064'), html.escape(x))
                                      for x in tang if x and x != '—'))
        self.ts.setText(ts.split(':', 1)[-1].strip())


class TrangLaSo(QWidget):
    def __init__(self):
        super().__init__()
        nen = QWidget(); v = QVBoxLayout(nen); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(14)
        h = QHBoxLayout(); h.addWidget(lbl('Lá số', 'tieude')); h.addStretch()
        v.addLayout(h)
        self.tt = lbl('', 'mo', True); v.addWidget(self.tt)
        ht = QHBoxLayout(); ht.setSpacing(12); self.tru = [TheTru() for _ in range(4)]
        for t in self.tru: ht.addWidget(t)
        v.addLayout(ht)
        self.ghi = lbl('', 'chip_info', True); v.addWidget(self.ghi)
        v.addWidget(lbl('Đại vận', 'tieude_the'))
        self.dv_box = QHBoxLayout(); self.dv_box.setSpacing(8); self.dv = []
        for _ in range(10):
            f = QFrame(); f.setObjectName('the'); fl = QVBoxLayout(f); fl.setContentsMargins(8, 8, 8, 8); fl.setSpacing(2)
            a = QLabel(); a.setAlignment(Qt.AlignCenter); a.setWordWrap(True); fl.addWidget(a)
            self.dv_box.addWidget(f); self.dv.append((f, a))
        v.addLayout(self.dv_box)
        v.addStretch()
        lo = QVBoxLayout(self); lo.setContentsMargins(0, 0, 0, 0); lo.addWidget(cuon(nen))

    def hien(self, kq):
        E = kq.E; o = lambda r, c: chu_o(E, 'KET_LUAN', r, c); td = tu_dien(E)
        self.tt.setText(' · '.join('<b>%s</b>: %s' % (html.escape(k), html.escape(v)) for k, v in thong_tin_la(kq)))
        for i, t in enumerate(self.tru):
            c = 3 + i
            t.dat(o(16, c), o(17, c), o(18, c), o(19, c), [o(20, c), o(21, c), o(22, c)], o(23, c), i == 2)
            t.setToolTip(goi_y(' '.join([o(17, c), o(23, c), 'thập thần', 'nhật chủ' if i == 2 else '']), td))
        # nạp âm năm (LUC_THAN dòng 5) và tổng hợp thập thần (KET_LUAN C24)
        lt = [chu_o(E, 'LUC_THAN', 5, c) for c in range(3, 9)]
        nap = next((x for x in lt if x and not re.fullmatch(r'[\d.]+', x) and '(' in x), '')
        self.ghi.setText((('Nạp âm năm: <b>%s</b> · ' % html.escape(nap)) if nap else '') + html.escape(o(24, 3)))
        dk = kieu_sheet(E, 'DAI_VAN'); d = lambda r, c: chu_o(E, 'DAI_VAN', r, c)
        dong_xt = 69 if chu_o(E, 'DOI_CHIEU', 33, 3) == '2' else 54        # như DAI_VAN!C106: phương pháp khuyến nghị
        for i, (f, a) in enumerate(self.dv):
            c = 3 + i; dang = bool(d(44, c).strip())
            k = dk.get((35, c), {}); xt = d(dong_xt, c)
            mau_xt = '#c00000' if xt == 'Tốt' else '#6b7280' if xt == 'Xấu' else '#1e2140'
            a.setText("<span style='font-size:15px;font-weight:800;color:%s'>%s</span><br>"
                      "<span style='color:#7b809c'>từ %s tuổi<br>%s–%s</span><br>%s<br>"
                      "<span style='color:#7b809c;font-size:10px'>XU THẾ</span><br><b style='font-size:14px;color:%s'>%s</b><br>"
                      "<span style='color:#9094ad;font-size:10px'>Hóa giải toàn thư</span><br><span style='color:#9094ad'>%s</span>%s"
                      % (k.get('chu', '#1e2140'), html.escape(d(35, c)), d(36, c), d(37, c), d(38, c), html.escape(d(39, c)),
                         mau_xt, html.escape(xt), html.escape(d(102, c)), '<br><b style="color:#4f46e5">◀ ĐANG HÀNH</b>' if dang else ''))
            a.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
            f.setStyleSheet('#the{border:2px solid #7c3aed;}' if dang else '')
            f.setToolTip('%s – năm %s–%s\nNhật can tại chi vận: %s\nTàng can chi vận: %s\n\nXu thế: theo sách Thiệu Vĩ Hoa.\n'
                         'Hóa giải toàn thư (tr.83): chấm thêm theo vai trò can – chi vận, chỉ để tham khảo.\n\n%s'
                         % (d(35, c), d(37, c), d(38, c), d(41, c), d(40, c), goi_y('đại vận ' + d(39, c), td)))


# ---------------------------------------------------------------- LUẬN GIẢI – BÀI ĐỌC
CSS_BAI = """
body { font-family: 'Inter'; font-size: 10.5pt; color: #1e2140; }
h1 { font-size: 17pt; color: #3730a3; } h2 { font-size: 13.5pt; color: #3730a3; margin-top: 16px; }
h3 { font-size: 11.5pt; margin: 12px 0 4px 0; }
.nhan { font-weight: bold; color: #4f46e5; } .nhan_kl { font-weight: bold; color: #ffffff; background: #6d28d9; }
.nhan_ct { font-weight: bold; color: #166534; } .nhan_sach { color: #6b7090; } .nhan_kn { color: #b45309; }
.nho { color: #7b809c; font-size: 9pt; }
a { color: #5b21b6; text-decoration: underline; }
td { padding: 4px 6px; border: 1px solid #e3e5f0; vertical-align: top; } th { background: #f3f0ff; color: #4338ca; padding: 4px 6px; }
"""


class TrangLuanGiai(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); h.addWidget(lbl('Luận giải chi tiết', 'tieude')); h.addStretch()
        self.o_tim = QLineEdit(); self.o_tim.setPlaceholderText('🔍  Tìm trong bài (Enter = tiếp)'); self.o_tim.setMinimumWidth(280)
        h.addWidget(self.o_tim); v.addLayout(h)
        self.bai = QTextBrowser(); self.bai.setOpenLinks(False); self.bai.document().setDefaultStyleSheet(CSS_BAI)
        self.bai.setStyleSheet('QTextBrowser{background:white;border:1px solid #e7e8f3;border-radius:16px;padding:18px 26px;}')
        v.addWidget(self.bai, 1)
        self.bai.highlighted.connect(self._goi_y); self.bai.anchorClicked.connect(lambda u: self._goi_y(u))
        self.o_tim.returnPressed.connect(lambda: self.bai.find(self.o_tim.text()) or (self.bai.moveCursor(self.bai.textCursor().Start), self.bai.find(self.o_tim.text())))
        self.td = []; self.kq = None; self._da = False

    def _goi_y(self, u):
        s = u.toString() if isinstance(u, QUrl) else str(u)
        if not s.startswith('tn:'): QToolTip.hideText(); return
        k, cau, o = self.td[int(s[3:])]
        from ui.bang_sheet import CHE_DO
        nguon = ('<br><span style="color:#7b809c">%s</span>' % o) if CHE_DO['chuyen_gia'] else ''
        QToolTip.showText(QCursor.pos(), '<b>%s</b><br>%s%s' % (html.escape(k), html.escape(cau), nguon), self.bai)

    def _lk(self, t):
        """Escape + gắn liên kết chú giải cho lần xuất hiện đầu của mỗi thuật ngữ trong đoạn."""
        s = html.escape(t)
        for i, (k, cau, o) in enumerate(self.td):
            m = re.search(r'(?<![\wÀ-ỹ])(' + re.escape(html.escape(k)) + r')(?![\wÀ-ỹ])', s, re.I)
            if m and '<a ' not in s[max(0, m.start() - 60):m.start()]:
                s = s[:m.start()] + '<a href="tn:%d">%s</a>' % (i, m.group(1)) + s[m.end():]
        return s.replace('\n', '<br>')

    def hien(self, kq):
        self.kq = kq; self._da = False
        if self.isVisible(): self._ve()

    def showEvent(self, e):
        super().showEvent(e)
        if self.kq and not self._da: self._ve()

    def _ve(self):
        E = self.kq.E; self.td = tu_dien(E); b = khoi(E, 'LUAN_GIAI')
        o = b['o']; H = ['<html><body>']
        hdr105 = {c: o[(105, c)][0] for c in range(3, 10) if (105, c) in o}
        bang = []

        def xa_bang():
            if not bang: return
            H.append('<table cellspacing="0" width="100%">')
            for j, (r, cs) in enumerate(bang):
                H.append('<tr>' + ''.join(('<th>%s</th>' if j == 0 else '<td style="%s">%%s</td>' % (
                    ';'.join(x for x in (('color:%s' % k['chu']) if k.get('chu') else '', 'font-weight:bold' if k.get('dam') else '') if x)))
                    % self._lk(t) for c, t, k in cs) + '</tr>')
            H.append('</table>'); bang.clear()
        for r in b['dong']:
            cs = [(c, o[(r, c)][0], o[(r, c)][1]) for c in b['cot'] if (r, c) in o]
            if not cs: continue
            if r >= 106 and hdr105 and cs[0][0] == 1 and re.fullmatch(r'\d+(\.0)?', cs[0][1]):       # phần B: thẻ 6 ý
                xa_bang(); ten = next((t for c, t, k in cs if c == 2), '')
                H.append('<h3>%s. %s</h3>' % (cs[0][1].split('.')[0], self._lk(ten)))
                for c, t, k in cs:
                    if c >= 3 and t: H.append('<p><span class="nhan">%s:</span> %s</p>' % (html.escape(hdr105.get(c, '')), self._lk(t)))
                continue
            if r == 105: xa_bang(); continue
            if len(cs) >= 5: bang.append((r, cs)); continue
            xa_bang()
            if len(cs) == 1 and cs[0][0] == 1:
                t = cs[0][1]
                H.append(('<h1>%s</h1>' if r == 1 else '<h2>%s</h2>' if re.match(r'^(\d+\.|[A-Z]\.)', t) else '<p class="nho">%s</p>') % html.escape(t))
                continue
            nhan = next((t for c, t, k in cs if c == 2), ''); noi = ' '.join(t for c, t, k in cs if c >= 3)
            st = 'nhan_kl' if nhan.startswith('KẾT LUẬN') else 'nhan_ct' if nhan.startswith('CÁT TƯỜNG') else \
                'nhan_sach' if nhan.startswith('[Sách') else 'nhan_kn' if nhan.startswith('[Kinh') else 'nhan'
            if nhan or noi:
                H.append('<p><span class="%s">&nbsp;%s&nbsp;</span> %s</p>' % (st, html.escape(nhan), self._lk(noi)))
        xa_bang()
        H.append('<h2>Chú giải thuật ngữ</h2><table cellspacing="0" width="100%">')
        for k, cau, oo in self.td:
            H.append('<tr><td width="20%%"><b>%s</b></td><td>%s</td></tr>' % (html.escape(k), html.escape(cau)))
        H.append('</table></body></html>')
        y = self.bai.verticalScrollBar().value()
        self.bai.setHtml(''.join(H)); self.bai.verticalScrollBar().setValue(y); self._da = True


# ---------------------------------------------------------------- KIỂM CHỨNG SỰ VIỆC (SU_KIEN)
class TrangSuKien(QWidget):
    doi_su_kien = Signal(list)

    def __init__(self):
        super().__init__()
        M = model(); self.loai = [''] + M.danh_sach_chon.get('SU_KIEN', {}).get('C6:C35', [])
        v = QVBoxLayout(self); v.setContentsMargins(28, 24, 28, 24); v.setSpacing(10)
        h = QHBoxLayout(); h.addWidget(lbl('Kiểm chứng sự việc', 'tieude')); h.addStretch()
        self.n_ap = QPushButton('✔  Áp dụng & tính lại'); self.n_ap.setObjectName('nut_chinh')
        self.n_xoa = QPushButton('Xóa hết'); self.n_xoa.setObjectName('nut_phu')
        h.addWidget(self.n_xoa); h.addWidget(self.n_ap); v.addLayout(h)
        self.hd = lbl('', 'mo', True); v.addWidget(self.hd)
        self.kq_lbl = lbl('', 'chip_info', True); v.addWidget(self.kq_lbl)
        bc = M.bo_cuc.get('SU_KIEN', {}); an = set(bc.get('an_cot', []))
        self.cot_kq = [c for c in range(5, 13) if c not in an]
        self.bang = QTableWidget(SO_DONG_SU_KIEN, 3 + len(self.cot_kq)); self.bang.verticalHeader().setDefaultSectionSize(30)
        self.bang.setVerticalHeaderLabels([str(6 + i) for i in range(SO_DONG_SU_KIEN)]); self.bang.setWordWrap(True)
        self.bang.setColumnWidth(0, 80); self.bang.setColumnWidth(1, 300); self.bang.setColumnWidth(2, 220)
        for j in range(len(self.cot_kq)): self.bang.setColumnWidth(3 + j, 170)
        self.cb = []
        for i in range(SO_DONG_SU_KIEN):
            self.bang.setItem(i, 0, QTableWidgetItem('')); self.bang.setItem(i, 2, QTableWidgetItem(''))
            c = QComboBox(); c.addItems(self.loai); self.bang.setCellWidget(i, 1, c); self.cb.append(c)
        v.addWidget(self.bang, 1)
        self.loi = lbl('', 'chip_err', True); self.loi.setVisible(False); v.addWidget(self.loi)
        self.n_ap.clicked.connect(self.ap_dung); self.n_xoa.clicked.connect(self.xoa)
        self.kq = None

    def nhap(self):
        ds = []; loi = []
        for i in range(SO_DONG_SU_KIEN):
            nam = self.bang.item(i, 0).text().strip(); loai = self.cb[i].currentText(); ghi = self.bang.item(i, 2).text().strip()
            if not (nam or loai or ghi): continue
            if nam and not re.fullmatch(r'\d{4}', nam): loi.append('dòng %d: năm "%s" phải là 4 chữ số' % (6 + i, nam)); continue
            ds.append([int(nam) if nam else None, loai or None, ghi or None])
        return ds, loi

    def ap_dung(self):
        ds, loi = self.nhap()
        self.loi.setText('Chưa áp dụng – ' + '; '.join(loi)); self.loi.setVisible(bool(loi))
        if not loi: self.doi_su_kien.emit(ds)

    def xoa(self):
        for i in range(SO_DONG_SU_KIEN):
            self.bang.item(i, 0).setText(''); self.bang.item(i, 2).setText(''); self.cb[i].setCurrentIndex(0)
        self.doi_su_kien.emit([])

    def hien(self, kq):
        self.kq = kq; E = kq.E; M = model()
        self.hd.setText(chu_o(E, 'SU_KIEN', 2, 2)); self.kq_lbl.setText(chu_o(E, 'SU_KIEN', 3, 2))
        self.bang.setHorizontalHeaderLabels([chu_o(E, 'SU_KIEN', 5, c) for c in (2, 3, 4)] + [chu_o(E, 'SU_KIEN', 5, c) for c in self.cot_kq])
        sk = list(kq.la_so.su_kien or [])
        for i in range(SO_DONG_SU_KIEN):
            nam, loai, ghi = (list(sk[i]) + [None] * 3)[:3] if i < len(sk) else (None, None, None)
            self.bang.item(i, 0).setText('' if nam in (None, '') else str(int(nam)))
            self.cb[i].setCurrentText(loai or ''); self.bang.item(i, 2).setText(ghi or '')
            for j, c in enumerate(self.cot_kq):
                t = chu_o(E, 'SU_KIEN', 6 + i, c)
                it = QTableWidgetItem(t); it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                if len(t) > 40: it.setToolTip(t[:1500])
                it.setForeground(QColor('#1e2140')); it.setBackground(QColor('#fafaff')); self.bang.setItem(i, 3 + j, it)
        for i in range(SO_DONG_SU_KIEN): self.bang.setRowHeight(i, 34)      # chữ dài: rê chuột xem đủ
