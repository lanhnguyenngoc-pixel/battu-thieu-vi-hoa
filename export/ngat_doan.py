# -*- coding: utf-8 -*-
"""Ngắt đoạn cho các trang đọc (yêu cầu 16/11/2026: có ngắt đoạn, xuống dòng, tránh một đoạn dài, các ý không chồng lấn).
Chỉ ở LỚP HIỂN THỊ – không đổi giá trị ô; chỉ chèn xuống dòng ('\n'), không thêm bớt chữ.
Quy tắc (theo thứ tự):
  1. Luôn xuống dòng trước: gạch đầu dòng "•", ký hiệu "►", "⇒", mục "(a)…(h)", nhãn viết hoa "DỮ KIỆN:", "QUY TẮC:", "KẾT LUẬN:"…
  2. Đoạn còn dài: xuống dòng ở cuối câu khi câu sau mở bằng một nhãn ngắn kết thúc bằng dấu ":" (vd "Dấu hiệu mệnh yểu (3): …").
  3. Đoạn vẫn dài hơn ~DAI ký tự: gom câu thành đoạn ngắn, mỗi đoạn tối đa ~DAI ký tự.
Không ngắt trong ngoặc ( ), trong câu trích “ ”, trong chữ đậm ** **, trong câu sách đánh dấu \\x01…\\x02."""
import re

DAI = 230
NHAN_TU = 140          # dòng dài hơn mức này thì tách các câu mở bằng nhãn ("Dấu hiệu …: ") thành dòng riêng
_HOA = r'A-ZĐÀÁẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÈÉẺẼẸÊẾỀỂỄỆÌÍỈĨỊÒÓỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÙÚỦŨỤƯỨỪỬỮỰỲÝỶỸỴ'
_MANH = [re.compile(r' (?=[•►⇒] )'), re.compile(r' (?=\([a-h]\) )'),
         re.compile(r'(?<=[.;:]) (?=[%s]{2,}[%s ]{0,30}:)' % (_HOA, _HOA))]
_NHAN = re.compile(r'(?<=\.) (?=[%s][^.:;\n“”]{1,44}: )' % _HOA)
_CAU = re.compile(r'(?<=[.;!?]) (?=[%s0-9“"►⇒→•(])' % _HOA)


_BAO = [re.compile(r'“[^“”\n]{0,400}”'), re.compile(r'\*\*.+?\*\*'), re.compile(r'\x01[^\x02]*\x02')]


def _sau(s):
    """Vị trí an toàn để ngắt: ngoài ngoặc ( ) cân bằng, ngoài câu trích “…” có đóng, chữ đậm **…**, câu sách \x01…\x02."""
    cam = set()
    for mau in _BAO:
        for m in mau.finditer(s): cam.update(range(m.start(), m.end()))
    mo = []
    for i, c in enumerate(s):
        if c == '(': mo.append(i)
        elif c == ')' and mo: cam.update(range(mo.pop(), i + 1))
    return set(range(len(s))) - cam


def _cat(s, mau, an_toan):
    vt = [m.start() for m in mau.finditer(s) if m.start() in an_toan and m.start() > 0]
    if not vt: return [s]
    ra, dau = [], 0
    for v in vt:
        ra.append(s[dau:v]); dau = v + 1
    ra.append(s[dau:])
    return [x for x in ra if x.strip()]


def _gom(s, dai):
    """Gom câu thành đoạn ≤ dai ký tự (ngắt ở cuối câu an toàn)."""
    if len(s) <= dai: return [s]
    at = _sau(s); vt = [m.start() for m in _CAU.finditer(s) if m.start() in at]
    if not vt: return [s]
    ra, dau, cuoi = [], 0, 0
    for v in vt + [len(s)]:
        if v - dau > dai and cuoi > dau + 60:
            ra.append(s[dau:cuoi]); dau = cuoi + 1
        cuoi = v
    ra.append(s[dau:])
    return [x for x in ra if x.strip()]


def ngat(t, dai=DAI):
    """Chuỗi → chuỗi đã chèn xuống dòng. Dòng ngắn giữ nguyên; gọi lại nhiều lần cho cùng kết quả."""
    if not t or len(t) <= NHAN_TU and '•' not in t: return t
    ra = []
    for dong in t.split('\n'):
        thut = re.match(r'[ \t]*', dong).group(0); dong = dong[len(thut):]
        if len(dong) <= NHAN_TU and ' • ' not in dong: ra.append(thut + dong); continue
        dong = re.sub(r'(?<=[^\s])(?=[•►⇒] )', ' ', dong)            # "…[MT4]► Ý nghĩa" → xuống dòng trước ►
        phan = [dong]
        for mau in _MANH:
            phan = [y for x in phan for y in _cat(x, mau, _sau(x))]
        phan = [y for x in phan for y in (_cat(x, _NHAN, _sau(x)) if len(x) > NHAN_TU else [x])]
        phan = [y for x in phan for y in _gom(x, dai)]
        ra += [thut + x.strip() for x in phan]
    return '\n'.join(ra)


def doan(t, dai=DAI):
    """Chuỗi → danh sách đoạn (để dựng nhiều <p> / đoạn Word)."""
    return [x for x in ngat(t, dai).split('\n') if x.strip()]
