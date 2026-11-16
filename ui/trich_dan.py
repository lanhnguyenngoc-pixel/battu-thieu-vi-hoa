# -*- coding: utf-8 -*-
"""Trình bày kết luận và trích dẫn sách (quyết định 15/11/2026):
  * nội dung công cụ luận giải, kết luận → CHỮ ĐẬM;
  * trích dẫn lý thuyết của sách (câu sách, số trang) → chữ thường, ở DÒNG RIÊNG, ẩn sau liên kết "📖 Căn cứ sách",
    bấm vào mới sổ ra.
tach(): tách phần trích dẫn khỏi câu luận giải (chỉ ở lớp hiển thị – không đổi giá trị gốc)."""
import html, re

_TR = r'(?:P1|P2|HG|UH|TTT|VD\d+\s+P\d)\s*tr\.\s*\d'
# (… P1 tr.70 …) – ngoặc đơn có số trang
_NGOAC = re.compile(r'\s*\((?=[^()]*' + _TR + r')[^()]*\)')
_NGOAC_Q = re.compile(r'\s*\((?=[^()]*[“"])(?=[^()]*' + _TR + r')[^()]*\)')
# "câu sách" – P1 tr.180   hoặc   “câu sách” (P1 tr.180)
_TRANG = r'(?:P1|P2|HG|UH)\s*tr\.\s*\d+(?:\s*[–\-]\s*\d+)?(?:\s*[,;]\s*(?:(?:P1|P2|HG|UH)\s*tr\.\s*)?\d+(?:\s*[–\-]\s*\d+)?)*'
_CAU = re.compile(r'[“"]([^”"]{2,400})[”"]\s*(?:[–—-]\s*)?\(?(' + _TRANG + r')\)?')


def _y_chinh(cau, toi_da=70):
    """Ý chính của câu sách dài để hiện tại chỗ (cắt ở ranh giới từ); câu đầy đủ nằm trong 'Căn cứ sách'."""
    if len(cau) <= toi_da: return cau
    cat = cau[:toi_da].rsplit(' ', 1)[0].rstrip(' ,;:–-')
    return cat + ' …'


def tach(t, giu_cau=False):
    """→ (nội dung luận giải, [các trích dẫn]). Không tách được thì trả nguyên văn.
    giu_cau=True (bản in, PDF): câu sách giữ NGUYÊN VĂN tại chỗ kèm số trang (bọc giữa \x01 và \x02 để lớp in
    trình bày chữ thường), không đưa vào danh sách trích dẫn.
    giu_cau=False (màn hình): câu ngắn giữ tại chỗ, câu dài hiện ý chính; câu đầy đủ ở liên kết "Căn cứ sách"."""
    if not t: return '', []
    cits = []

    def _cau(m):
        cau, trang = m.group(1).strip(), m.group(2).strip().rstrip(';,. ')
        if giu_cau: return '\x01“%s” – %s\x02' % (cau, trang)   # \x01…\x02: đánh dấu câu sách cho lớp in
        cits.append('“%s” (%s)' % (cau, trang))
        return '“%s”' % (cau if len(cau) <= 25 else _y_chinh(cau))

    def _ng(m):
        x = m.group(0).strip()[1:-1].strip()
        if x: cits.append(x)
        return ''
    s = _NGOAC_Q.sub(_ng, t)
    s = _CAU.sub(_cau, s)
    s = _NGOAC.sub(_ng, s)
    s = re.sub(r'[ \t]+([.,;:])', r'\1', s); s = re.sub(r'[ \t]{2,}', ' ', s); s = re.sub(r'\(\s*\)', '', s)
    s = re.sub(r'\n[ \t]+', '\n', s).strip()
    if not s: return t, []
    # gộp trùng, giữ thứ tự
    seen, out = set(), []
    for c in cits:
        if c not in seen: seen.add(c); out.append(c)
    return s, out


def chi_la_trich(noi, cits):
    """Nội dung chính chỉ là câu sách (tách xong không còn ý luận giải) → hiển thị nguyên văn, chữ thường."""
    con = re.sub(r'[^\wÀ-ỹ]+', '', re.sub(r'“[^”]* …”|\(xem căn cứ sách\)', '', re.sub(r'^[^:]{0,60}:', '', noi)))
    return bool(cits) and any(('“' in c or '"' in c) for c in cits) and len(con) < 12


class TrichDan:
    """Bộ đếm id liên kết + tập id đang mở (dùng chung cho một trang)."""
    def __init__(self):
        self.mo = set(); self.n = 0

    def bat_dau(self):
        self.n = 0

    def html(self, t, lk=None, dam=True, kieu='', the='span'):
        """Một đoạn: nội dung đậm + (dòng riêng) liên kết căn cứ sách; nội dung trích dẫn sổ ra khi đã bấm."""
        lk = lk or (lambda x: html.escape(x).replace('\n', '<br>'))
        noi, cits = tach(t)
        if chi_la_trich(noi, cits):
            # nội dung chính là câu sách (vd "Tính cách theo sách: • … • …") → giữ nguyên văn, chữ thường
            return '<%s style="%s">%s</%s>' % (the, kieu, lk(t), the)
        st = (kieu + ';' if kieu else '') + ('font-weight:bold' if dam else '')
        from export.ngat_doan import ngat
        H = '<%s style="%s">%s</%s>' % (the, st, lk(ngat(noi)), the)   # ngắt đoạn: mỗi ý một dòng
        if cits:
            i = self.n; self.n += 1
            mo = i in self.mo
            H += '<br><a href="cd:%d" style="font-weight:normal;font-size:9pt;color:#6d28d9">📖 Căn cứ sách %s</a>' % (i, '▴ (bấm để ẩn)' if mo else '▾')
            if mo:
                H += '<br><span style="font-weight:normal;color:#4b5068;font-size:9.5pt">%s</span>' % '<br>'.join(html.escape(c) for c in cits)
        return H

    def khoi(self, cits):
        """Liên kết căn cứ sách cho danh sách câu đã tách sẵn."""
        cits = [c for c in cits if c and c.strip()]
        if not cits: return ''
        i = self.n; self.n += 1; mo = i in self.mo
        H = '<a href="cd:%d" style="font-weight:normal;font-size:9pt;color:#6d28d9">📖 Căn cứ sách %s</a>' % (i, '▴ (bấm để ẩn)' if mo else '▾')
        if mo:
            H += '<br><span style="font-weight:normal;color:#4b5068;font-size:9.5pt">%s</span>' % '<br>'.join(html.escape(c) for c in cits)
        return H

    def bam(self, url):
        """Xử lý bấm liên kết cd:N → True nếu là liên kết căn cứ (cần vẽ lại)."""
        s = url.toString() if hasattr(url, 'toString') else str(url)
        if not s.startswith('cd:'): return False
        i = int(s[3:]); self.mo ^= {i}; return True


def dai_van_html(E, tdan, lk):
    """Khối KẾT LUẬN ĐẠI VẬN (đầu trang Vận trình – Đại vận, chế độ thường)."""
    from ui.trang_ket_qua import chu_o
    d = lambda ref: chu_o(E, 'DAI_VAN', ref)
    H = ['<h2>Kết luận đại vận</h2>']
    for c in 'CDEFGHIJ':
        ten = d(c + '35')
        if not ten.strip(): continue
        muc = d(c + '121'); mau = '#c00000' if 'TỐT' in muc else '#4b5563'
        try: t0 = int(float(d(c + '36')))
        except ValueError: t0 = None
        H.append('<h3 style="margin-bottom:2px">Vận %s (%s tuổi, năm %s–%s): <span style="color:%s">%s</span>%s</h3>'
                 % (html.escape(ten), ('%d–%d' % (t0, t0 + 9)) if t0 is not None else '', d(c + '37'), d(c + '38'), mau, html.escape(muc),
                    ' <span style="color:#4f46e5">◀ đang hành</span>' if d(c + '44').strip() else ''))
        tom = d(c + '122')
        H.append('<p style="margin:2px 0"><b>%s</b></p>' % html.escape(tom[:1].upper() + tom[1:]))
        H.append('<p style="margin:2px 0"><span class="nhan">5 năm đầu:</span> <b>%s</b></p>' % lk(d(c + '123')))
        H.append('<p style="margin:2px 0"><span class="nhan">5 năm sau:</span> <b>%s</b></p>' % lk(d(c + '124')))
        H.append('<p style="margin:2px 0"><span class="nhan">Tài lộc:</span> <b>%s – %s</b></p>'
                 % (html.escape(chu_o(E, 'TAI_LOC', c + '89')), lk(chu_o(E, 'TAI_LOC', c + '90'))))
        H.append('<p style="margin:2px 0"><span class="nhan">Công danh:</span> <b>%s</b></p>' % lk(d(c + '127')))
        H.append('<p style="margin:2px 0 10px 0">%s</p>' % tdan.khoi([d(c + '125'), chu_o(E, 'TAI_LOC', c + '91')]))
    return ''.join(H)
