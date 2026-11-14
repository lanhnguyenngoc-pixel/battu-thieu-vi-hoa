# -*- coding: utf-8 -*-
"""Thư viện hàm Excel theo NGỮ NGHĨA LIBREOFFICE 24.2 (đáp án dap_an_o do LibreOffice tính – MT-01).

Quy ước giá trị:  số = float · chữ = str · ô trống = None · Đúng/Sai = bool · lỗi = ném XLError.
Các điểm đã đo trên LibreOffice (tools/lo_thu.py, 13/11/2026):
  * Cộng/trừ dùng approxAdd/approxSub; so sánh số dùng approxEqual (sai lệch < 2^-48 coi là bằng).
  * INT, MOD dùng approxFloor (làm tròn 15 chữ số có nghĩa trước khi lấy sàn).
  * ROUND và TEXT "0.00" làm tròn nửa xa số 0 sau khi chuẩn hóa 15 chữ số (ROUND(1.005;2)=1.01).
  * Số → chữ (phép &): 15 chữ số có nghĩa, mũ khi |x| ≥ 1E16 hoặc rất nhỏ (dạng 1E-020).
  * Đúng/Sai ghép chữ thành "1"/"0"; ô trống = 0 khi tính, = "" khi ghép chữ.
  * So sánh chữ không phân biệt hoa thường; số luôn nhỏ hơn chữ.
  * Định dạng giờ "hh:mm" cắt phần lẻ (không làm tròn).
  * SEARCH, MATCH(…;0), COUNTIF dùng ký tự đại diện * ? ~ (chế độ xlsx).
"""
import math, re, datetime as _dt

# ------------------------------------------------------------------ lỗi
class XLError(Exception):
    def __init__(self, code='#VALUE!'):
        super().__init__(code); self.code = code
    def __repr__(self): return self.code

NA = '#N/A'; E_VALUE = '#VALUE!'; DIV0 = '#DIV/0!'; REF = '#REF!'; NUM = '#NUM!'; NAME = '#NAME?'

class ErrVal:
    """Giá trị lỗi lưu trong ô (đọc ô này sẽ ném lại lỗi)."""
    __slots__ = ('code',)
    def __init__(self, code): self.code = code
    def __repr__(self): return self.code

class _Missing:
    def __repr__(self): return 'MISSING'
MISSING = _Missing()

# ------------------------------------------------------------------ số học gần đúng (rtl::math)
_E48 = 1.0 / (16777216.0 * 16777216.0)

def approx_value(x):
    if x == 0.0 or not math.isfinite(x): return x
    return float('%.15g' % x)

def approx_floor(x): return math.floor(approx_value(x))
def approx_ceil(x): return math.ceil(approx_value(x))

def approx_equal(a, b):
    if a == b: return True
    if a == 0.0 or b == 0.0: return False
    d = abs(a - b)
    if not math.isfinite(d): return False
    return d < abs(a) * _E48 and d < abs(b) * _E48

def approx_add(a, b):
    if ((a < 0.0 and b > 0.0) or (b < 0.0 and a > 0.0)) and approx_equal(a, -b): return 0.0
    return a + b

def approx_sub(a, b):
    if ((a < 0.0 and b < 0.0) or (a > 0.0 and b > 0.0)) and approx_equal(a, b): return 0.0
    return a - b

# ------------------------------------------------------------------ chuyển kiểu
_num_re = re.compile(r'^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$')

def num(x):
    """Giá trị → số (trong phép tính)."""
    t = type(x)
    if t is float: return x
    if x is None: return 0.0
    if t is bool: return 1.0 if x else 0.0
    if t is int: return float(x)
    if t is str:
        if _num_re.match(x): return float(x)
        raise XLError(E_VALUE)
    if isinstance(x, Range): return num(x.scalar())
    if x is MISSING: return 0.0
    raise XLError(E_VALUE)

def numstr(x):
    """Số → chữ kiểu 'General' của LibreOffice."""
    if x == 0: return '0'
    if not math.isfinite(x): raise XLError(NUM)
    s = '%.15g' % x
    if 'e' in s or 'E' in s:
        m, e = s.split('e'); e = int(e)
        if -16 < e < 16:
            # viết dạng thường
            from decimal import Decimal
            d = Decimal(s)
            t = format(d, 'f')
            if '.' in t: t = t.rstrip('0').rstrip('.')
            return t
        if '.' in m: m = m.rstrip('0').rstrip('.')
        return '%sE%s%03d' % (m, '-' if e < 0 else '+', abs(e))
    if '.' in s: s = s.rstrip('0').rstrip('.')
    return s

def text(x):
    """Giá trị → chữ (trong phép &, hàm chuỗi)."""
    t = type(x)
    if t is str: return x
    if x is None: return ''
    if t is bool: return '1' if x else '0'
    if t is float or t is int: return numstr(float(x))
    if isinstance(x, Range): return text(x.scalar())
    if x is MISSING: return ''
    raise XLError(E_VALUE)

def truthy(x):
    t = type(x)
    if t is bool: return x
    if t is float or t is int: return x != 0
    if x is None or x is MISSING: return False
    if t is str:
        u = x.upper()
        if u == 'TRUE': return True
        if u == 'FALSE': return False
        raise XLError(E_VALUE)
    if isinstance(x, Range): return truthy(x.scalar())
    raise XLError(E_VALUE)

def scal(x):
    return x.scalar() if isinstance(x, Range) else x

# ------------------------------------------------------------------ vùng ô
class Range:
    __slots__ = ('E', 'si', 'r1', 'c1', 'r2', 'c2')
    def __init__(self, E, si, r1, c1, r2, c2):
        self.E = E; self.si = si; self.r1 = r1; self.c1 = c1; self.r2 = r2; self.c2 = c2
    @property
    def nrows(self): return self.r2 - self.r1 + 1
    @property
    def ncols(self): return self.c2 - self.c1 + 1
    def raw(self, r, c):
        """Giá trị thô (không ném lỗi): có thể là ErrVal."""
        return self.E.raw(self.si, r, c)
    def cells(self):
        E = self.E; si = self.si
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                yield E.raw(si, r, c)
    def vector(self):
        return list(self.cells())
    def at(self, i, j):  # 0-based
        return self.E.v(self.si, self.r1 + i, self.c1 + j)
    def scalar(self):
        if self.r1 == self.r2 and self.c1 == self.c2:
            return self.E.v(self.si, self.r1, self.c1)
        # giao ngầm (implicit intersection) – chưa cần ở GĐ1
        raise XLError(E_VALUE)

class Arr:
    """Hằng mảng {..}."""
    __slots__ = ('rows',)
    def __init__(self, rows): self.rows = rows
    @property
    def nrows(self): return len(self.rows)
    @property
    def ncols(self): return len(self.rows[0])
    def cells(self):
        for row in self.rows:
            for v in row: yield v
    def vector(self): return list(self.cells())
    def at(self, i, j): return self.rows[i][j]
    def raw(self, r, c): return self.rows[r][c]
    def scalar(self): return self.rows[0][0]

def _vals(x):
    """Trải tham số thành danh sách giá trị thô (giữ ErrVal)."""
    if isinstance(x, (Range, Arr)): return x.vector()
    return [x]

# ------------------------------------------------------------------ toán tử
def add(a, b): return approx_add(num(a), num(b))
def sub(a, b): return approx_sub(num(a), num(b))
def mul(a, b): return num(a) * num(b)
def div(a, b):
    b = num(b); a = num(a)
    if b == 0.0: raise XLError(DIV0)
    return a / b
def pow_(a, b):
    a = num(a); b = num(b)
    try:
        r = a ** b
    except ZeroDivisionError:
        raise XLError(DIV0)
    if isinstance(r, complex): raise XLError(NUM)
    return float(r)
def neg(a): return -num(a)
def pct(a): return num(a) / 100.0
def cat(a, b): return text(a) + text(b)

def _cmp(a, b):
    """So sánh kiểu LibreOffice: trả -1/0/1."""
    a = scal(a); b = scal(b)
    if isinstance(a, bool): a = 1.0 if a else 0.0
    if isinstance(b, bool): b = 1.0 if b else 0.0
    if a is None or a is MISSING:
        a = '' if isinstance(b, str) else 0.0
    if b is None or b is MISSING:
        b = '' if isinstance(a, str) else 0.0
    ta = isinstance(a, str); tb = isinstance(b, str)
    if not ta and not tb:
        a = float(a); b = float(b)
        if approx_equal(a, b): return 0
        return -1 if a < b else 1
    if ta and tb:
        la = a.lower(); lb = b.lower()
        if la == lb: return 0
        return -1 if la < lb else 1
    return -1 if not ta else 1  # số < chữ

def eq(a, b): return _cmp(a, b) == 0
def ne(a, b): return _cmp(a, b) != 0
def lt(a, b): return _cmp(a, b) < 0
def le(a, b): return _cmp(a, b) <= 0
def gt(a, b): return _cmp(a, b) > 0
def ge(a, b): return _cmp(a, b) >= 0

# ------------------------------------------------------------------ logic
def IF(c, a=False, b=False):  # bản không lười (bộ dịch dùng biểu thức điều kiện Python)
    return a if truthy(c) else b

def AND(*args):
    res = True; seen = False
    for a in args:
        for v in _vals(a):
            if isinstance(v, ErrVal): raise XLError(v.code)
            if v is None or v is MISSING or isinstance(v, str) and isinstance(a, (Range, Arr)): continue
            seen = True
            if not truthy(v): res = False
    if not seen: raise XLError(E_VALUE)
    return res

def OR(*args):
    res = False; seen = False
    for a in args:
        for v in _vals(a):
            if isinstance(v, ErrVal): raise XLError(v.code)
            if v is None or v is MISSING or isinstance(v, str) and isinstance(a, (Range, Arr)): continue
            seen = True
            if truthy(v): res = True
    if not seen: raise XLError(E_VALUE)
    return res

def NOT(a): return not truthy(scal(a))
def TRUE(): return True
def FALSE(): return False

def IFERROR(fa, fb):
    try:
        v = fa()
        if isinstance(v, Range): v = v.scalar()
        return v
    except XLError:
        return fb()

def IFNA(fa, fb):
    try:
        return scal(fa())
    except XLError as e:
        if e.code == NA: return fb()
        raise

def CHOOSE(i, *fs):
    k = int(approx_floor(num(scal(i))))
    if k < 1 or k > len(fs): raise XLError(E_VALUE)
    return fs[k - 1]()

def _probe(f):
    try:
        v = f()
        if isinstance(v, Range):
            v = v.E.raw(v.si, v.r1, v.c1) if (v.r1 == v.r2 and v.c1 == v.c2) else v.scalar()
        return v, None
    except XLError as e:
        return None, e.code

def ISNUMBER(f):
    v, e = _probe(f)
    return e is None and isinstance(v, (int, float))   # LibreOffice: Đúng/Sai cũng là số

def ISTEXT(f):
    v, e = _probe(f); return e is None and isinstance(v, str)

def ISNA(f):
    v, e = _probe(f)
    if isinstance(v, ErrVal): return v.code == NA
    return e == NA

def ISERROR(f):
    v, e = _probe(f); return e is not None or isinstance(v, ErrVal)

def ISERR(f):
    v, e = _probe(f)
    if isinstance(v, ErrVal): return v.code != NA
    return e is not None and e != NA

def ISBLANK(f):
    v, e = _probe(f); return e is None and v is None

def ISLOGICAL(f):
    v, e = _probe(f); return e is None and isinstance(v, bool)

def NA_(): raise XLError(NA)

# ------------------------------------------------------------------ số
def INT(a): return float(approx_floor(num(scal(a))))

def MOD(a, b):
    a = num(scal(a)); b = num(scal(b))
    if b == 0.0: raise XLError(DIV0)
    q = approx_floor(a / b)
    r = approx_sub(a, q * b)
    # đảm bảo cùng dấu với số chia (như LibreOffice)
    if r != 0.0 and (r < 0) != (b < 0):
        r += b
    return r

_CORR = [0.0] + [9.0 * 10.0 ** -k for k in range(1, 16)]   # nCorrVal của rtl::math::round (LibreOffice)

def _round_half_away(x, n):
    """rtl::math::round(x, n, Corrected) của LibreOffice 24.2 (dùng cho ROUND, TEXT, FIXED, định dạng ô).
    n = 0: làm tròn nửa xa 0 đúng giá trị nhị phân (std::round) – đo: ROUND(55,49999999999999;0) = 55.
    n ≠ 0: nhân 10^n rồi floor(v + 0,5 + 9·10^-(15-e)) với e = số mũ thập phân – đo: ROUND(1,005;2) = 1,01."""
    if x == 0.0 or not math.isfinite(x): return x
    n = int(n); neg = x < 0; v = -x if neg else x
    if n == 0:
        r = math.floor(v + 0.5)
    else:
        fac = 10.0 ** abs(n)
        v = v * fac if n > 0 else v / fac
        e = int(math.floor(math.log10(v))) if v > 0 else 0
        idx = 15 if e < 0 else (0 if e >= 14 else 15 - e)
        r = math.floor(v + 0.5 + _CORR[idx])
        r = r / fac if n > 0 else r * fac
    return -r if neg else r

def ROUND(a, n=0.0):
    return _round_half_away(num(scal(a)), approx_floor(num(scal(n))) if n is not MISSING else 0)

def ROUNDUP(a, n=0.0):
    x = num(scal(a)); n = int(approx_floor(num(scal(n)))); f = 10.0 ** n
    y = approx_value(abs(x) * f); r = math.ceil(y) / f
    return -r if x < 0 else r

def ROUNDDOWN(a, n=0.0):
    x = num(scal(a)); n = int(approx_floor(num(scal(n)))); f = 10.0 ** n
    y = approx_value(abs(x) * f); r = math.floor(y) / f
    return -r if x < 0 else r

def ABS(a): return abs(num(scal(a)))
def SIGN(a):
    x = num(scal(a)); return 1.0 if x > 0 else (-1.0 if x < 0 else 0.0)

def _nums(args, strict_scalar=True):
    """Số trong các tham số theo quy tắc SUM/MAX: trong vùng bỏ chữ/trống/bool; tham số trực tiếp đổi kiểu."""
    out = []
    for a in args:
        if isinstance(a, (Range, Arr)):
            for v in a.cells():
                if isinstance(v, ErrVal): raise XLError(v.code)
                if isinstance(v, (int, float)): out.append(float(v))
        elif a is MISSING:
            continue
        else:
            if isinstance(a, str):
                out.append(num(a))
            elif a is None:
                continue
            else:
                out.append(num(a))
    return out

def SUM(*args):
    s = 0.0
    for x in _nums(args): s = approx_add(s, x)
    return s

def MAX(*args):
    v = _nums(args); return max(v) if v else 0.0
def MIN(*args):
    v = _nums(args); return min(v) if v else 0.0
def AVERAGE(*args):
    v = _nums(args)
    if not v: raise XLError(DIV0)
    s = 0.0
    for x in v: s = approx_add(s, x)
    return s / len(v)
def COUNT(*args):
    n = 0
    for a in args:
        if isinstance(a, (Range, Arr)):
            for v in a.cells():
                if isinstance(v, (int, float)): n += 1
        elif a is not MISSING and a is not None:
            try: num(a); n += 1
            except XLError: pass
    return float(n)
def COUNTA(*args):
    n = 0
    for a in args:
        if isinstance(a, (Range, Arr)):
            for v in a.cells():
                if v is not None: n += 1
        elif a is not MISSING: n += 1
    return float(n)

def LARGE(rng, k):
    v = sorted(_nums([rng]), reverse=True); k = int(approx_ceil(num(scal(k))))
    if k < 1 or k > len(v): raise XLError(NUM)
    return v[k - 1]
def SMALL(rng, k):
    v = sorted(_nums([rng])); k = int(approx_ceil(num(scal(k))))
    if k < 1 or k > len(v): raise XLError(NUM)
    return v[k - 1]

def RANK(x, rng, order=0.0):
    x = num(scal(x)); v = _nums([rng])
    asc = order is not MISSING and truthy(scal(order))
    if not any(approx_equal(x, y) for y in v): raise XLError(NA)
    # Đo trên LibreOffice: RANK đếm bằng so sánh CHÍNH XÁC (26,4 xếp sau 26,400000000000002),
    # khác phép =, >, COUNTIF (gần đúng 2^-48).
    if asc: return float(1 + sum(1 for y in v if y < x))
    return float(1 + sum(1 for y in v if y > x))

def SUMPRODUCT(*args):
    arrs = []
    for a in args:
        if isinstance(a, (Range, Arr)): arrs.append(a.vector())
        else: arrs.append([a])
    n = len(arrs[0])
    if any(len(x) != n for x in arrs): raise XLError(E_VALUE)
    s = 0.0
    for i in range(n):
        p = 1.0
        for x in arrs:
            v = x[i]
            if isinstance(v, ErrVal): raise XLError(v.code)
            p *= float(v) if (type(v) is float or type(v) is int or type(v) is bool) else 0.0
        s = approx_add(s, p)
    return s

def N(a):
    a = scal(a)
    if isinstance(a, bool): return 1.0 if a else 0.0
    if type(a) is float or type(a) is int: return float(a)
    return 0.0

def VALUE(a):
    a = scal(a)
    if isinstance(a, str):
        s = a.strip()
        if _num_re.match(s): return float(s)
        if s.endswith('%') and _num_re.match(s[:-1]): return float(s[:-1]) / 100
        t = _parse_time(s)
        if t is not None: return t
        raise XLError(E_VALUE)
    return num(a)

# ------------------------------------------------------------------ chuỗi
def CHAR(n): return chr(int(approx_floor(num(scal(n)))))
def LEN(a): return float(len(text(scal(a))))
def LOWER(a): return text(scal(a)).lower()
def UPPER(a): return text(scal(a)).upper()
def PROPER(a):
    s = text(scal(a)); out = []; prev_alpha = False
    for ch in s:
        out.append(ch.lower() if prev_alpha else ch.upper())
        prev_alpha = ch.isalpha()
    return ''.join(out)
def TRIM(a):
    s = text(scal(a))
    return re.sub(' +', ' ', s.strip(' '))
def LEFT(a, n=1.0):
    s = text(scal(a)); n = 1 if n is MISSING else int(approx_floor(num(scal(n))))
    if n < 0: raise XLError(E_VALUE)
    return s[:n]
def RIGHT(a, n=1.0):
    s = text(scal(a)); n = 1 if n is MISSING else int(approx_floor(num(scal(n))))
    if n < 0: raise XLError(E_VALUE)
    return s[-n:] if n else ''
def MID(a, st, n):
    s = text(scal(a)); st = int(approx_floor(num(scal(st)))); n = int(approx_floor(num(scal(n))))
    if st < 1 or n < 0: raise XLError(E_VALUE)
    return s[st - 1:st - 1 + n]
def REPT(a, n):
    n = int(approx_floor(num(scal(n))))
    if n < 0: raise XLError(E_VALUE)
    return text(scal(a)) * n
def SUBSTITUTE(a, old, new, inst=MISSING):
    s = text(scal(a)); old = text(scal(old)); new = text(scal(new))
    if old == '': return s
    if inst is MISSING: return s.replace(old, new)
    k = int(approx_floor(num(scal(inst))))
    if k < 1: raise XLError(E_VALUE)
    i = -1
    for _ in range(k):
        i = s.find(old, i + 1)
        if i < 0: return s
    return s[:i] + new + s[i + len(old):]
def FIND(f, s, st=1.0):
    f = text(scal(f)); s = text(scal(s)); st = 1 if st is MISSING else int(approx_floor(num(scal(st))))
    if st < 1 or st > len(s) + 1: raise XLError(E_VALUE)
    i = s.find(f, st - 1)
    if i < 0: raise XLError(E_VALUE)
    return float(i + 1)

def _wild_re(pat, whole=True):
    out = []; i = 0
    while i < len(pat):
        ch = pat[i]
        if ch == '~' and i + 1 < len(pat) and pat[i + 1] in '*?~':
            out.append(re.escape(pat[i + 1])); i += 2; continue
        if ch == '*': out.append('.*')
        elif ch == '?': out.append('.')
        else: out.append(re.escape(ch))
        i += 1
    body = ''.join(out)
    return re.compile(('^' + body + '$') if whole else body, re.S | re.I)

def _has_wild(s): return any(ch in s for ch in '*?~')

def SEARCH(f, s, st=1.0):
    f = text(scal(f)); s = text(scal(s)); st = 1 if st is MISSING else int(approx_floor(num(scal(st))))
    if st < 1 or st > len(s) + 1: raise XLError(E_VALUE)
    if _has_wild(f):
        m = _wild_re(f, whole=False).search(s, st - 1)
        if not m: raise XLError(E_VALUE)
        return float(m.start() + 1)
    i = s.lower().find(f.lower(), st - 1)
    if i < 0: raise XLError(E_VALUE)
    return float(i + 1)

# ------------------------------------------------------------------ ngày giờ (gốc 30/12/1899)
_EPOCH = _dt.datetime(1899, 12, 30)

def _ymd_to_serial(y, m, d):
    y = int(y); m = int(m); d = int(d)
    if 0 <= y < 100: y += 1900  # LibreOffice: năm 2 chữ số (cấu hình mặc định 1930-2029 không áp vào hàm DATE)
    y += (m - 1) // 12; m = (m - 1) % 12 + 1
    try:
        base = _dt.date(y, m, 1)
    except ValueError:
        raise XLError(NUM)
    return float((base - _EPOCH.date()).days + d - 1)

def DATE(y, m, d):
    return _ymd_to_serial(approx_floor(num(scal(y))), approx_floor(num(scal(m))), approx_floor(num(scal(d))))

def _serial_to_date(x):
    return _EPOCH.date() + _dt.timedelta(days=int(approx_floor(x)))

def DAY(a): return float(_serial_to_date(num(scal(a))).day)
def MONTH(a): return float(_serial_to_date(num(scal(a))).month)
def YEAR(a): return float(_serial_to_date(num(scal(a))).year)

def _parse_time(s):
    m = re.match(r'^\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d*)?))?\s*(AM|PM)?\s*$', s, re.I)
    if not m: return None
    h = int(m.group(1)); mi = int(m.group(2)); se = float(m.group(3) or 0)
    if m.group(4):
        if h == 12: h = 0
        if m.group(4).upper() == 'PM': h += 12
    return (h * 3600 + mi * 60 + se) / 86400.0

def TIMEVALUE(a):
    s = text(scal(a)); t = _parse_time(s)
    if t is None: raise XLError(E_VALUE)
    return t - math.floor(t)

# ------------------------------------------------------------------ TEXT / FIXED
def _fmt_fixed(x, dec, thousands):
    r = _round_half_away(x, dec)
    s = '%.*f' % (max(dec, 0), abs(r))
    if thousands:
        ip, _, fp = s.partition('.')
        ip = '{:,}'.format(int(ip))
        s = ip + ('.' + fp if fp else '')
    if r < 0 and float(s.replace(',', '')) != 0: s = '-' + s
    return s

def FIXED(x, dec=2.0, nocomma=False):
    x = num(scal(x)); dec = 2 if dec is MISSING else int(approx_floor(num(scal(dec))))
    nc = False if nocomma is MISSING else truthy(scal(nocomma))
    return _fmt_fixed(x, dec, not nc)

def _fmt_datetime(x, fmt):
    days = math.floor(x); frac = x - days
    secs = frac * 86400.0
    secs = math.floor(round(secs, 6))  # cắt phần lẻ giây (đo: 0,020833 → 00:29)
    if secs >= 86400: days += 1; secs -= 86400
    d = _EPOCH + _dt.timedelta(days=days, seconds=secs)
    out = fmt
    rep = [('yyyy', '%04d' % d.year), ('yy', '%02d' % (d.year % 100)), ('dd', '%02d' % d.day),
           ('hh', '%02d' % d.hour), ('ss', '%02d' % d.second)]
    # tháng/phút: "mm" sau hh là phút, còn lại là tháng
    res = ''; i = 0; last = ''
    toks = re.findall(r'yyyy|yy|dd|d|mm|m|hh|h|ss|s|[^ymdhs]+', fmt, re.I)
    for t in toks:
        tl = t.lower()
        if tl == 'yyyy': res += '%04d' % d.year; last = 'y'
        elif tl == 'yy': res += '%02d' % (d.year % 100); last = 'y'
        elif tl == 'dd': res += '%02d' % d.day; last = 'd'
        elif tl == 'd': res += '%d' % d.day; last = 'd'
        elif tl == 'hh': res += '%02d' % d.hour; last = 'h'
        elif tl == 'h': res += '%d' % d.hour; last = 'h'
        elif tl == 'ss': res += '%02d' % d.second; last = 's'
        elif tl == 's': res += '%d' % d.second; last = 's'
        elif tl in ('mm', 'm'):
            minute = last == 'h'
            v = d.minute if minute else d.month
            res += ('%02d' % v) if tl == 'mm' else ('%d' % v); last = 'm'
        else: res += re.sub(r'\\(.)', r'\1', t)
    return res

def _fmt_section(v, sec, keep_sign):
    m = re.fullmatch(r'([^0#]*)(#,##)?0(?:\.(0+))?(%)?([^0#]*)', sec)
    if not m: return None
    pre, th, dec, pc, suf = m.groups()
    pre = re.sub(r'\\(.)', r'\1', pre); suf = re.sub(r'\\(.)', r'\1', suf)   # ký tự thoát: 0\% = chữ "%"
    dec = len(dec or '')
    x = v * 100 if pc else v
    body = _fmt_fixed(x, dec, bool(th)) if keep_sign else _fmt_fixed(abs(x), dec, bool(th))
    return pre + body + ('%' if pc else '') + suf

def TEXT(x, fmt):
    fmt = text(scal(fmt)); x = scal(x)
    if isinstance(x, str):
        try: x = num(x)
        except XLError: return x
    xv = num(x)
    f = fmt
    if re.search(r'[ymdhs]', f, re.I) and not re.fullmatch(r'[0#.,%]+', f):
        return _fmt_datetime(xv, f)
    if ';' in f:
        # nhiều phần: chọn theo dấu của giá trị gốc (đo trên LibreOffice: TEXT(0,4;"+0;-0;0") = "+0", TEXT(-0,4;…) = "-0")
        secs = f.split(';')
        if xv > 0 or len(secs) == 1: sec, v = secs[0], xv
        elif xv < 0: sec, v = secs[1], -xv
        else: sec, v = (secs[2] if len(secs) > 2 else secs[0]), 0.0
        return _fmt_section(v, sec, keep_sign=(xv > 0 or len(secs) == 1))
    r = _fmt_section(xv, f, keep_sign=True)
    if r is not None: return r
    if f.lower() == 'general' or f == '':
        return numstr(xv)
    raise NotImplementedError('TEXT định dạng chưa hỗ trợ: %r' % fmt)

# ------------------------------------------------------------------ tra cứu
def ROW(ref=MISSING):
    if isinstance(ref, Range): return float(ref.r1)
    raise XLError(E_VALUE)
def COLUMN(ref=MISSING):
    if isinstance(ref, Range): return float(ref.c1)
    raise XLError(E_VALUE)

def INDEX(rng, row=MISSING, col=MISSING, area=MISSING):
    r = 0 if row is MISSING else int(approx_floor(num(scal(row))))
    c = 0 if col is MISSING else int(approx_floor(num(scal(col))))
    if isinstance(rng, Arr):
        nr, nc = rng.nrows, rng.ncols
        if col is MISSING and nr == 1: r, c = 1, r
        if r < 0 or c < 0 or r > nr or c > nc: raise XLError(REF)
        if r == 0 or c == 0:
            if nr == 1 and r == 0: r = 1
            elif nc == 1 and c == 0: c = 1
            else: raise XLError(E_VALUE)
        return rng.rows[r - 1][c - 1]
    if not isinstance(rng, Range):
        if r in (0, 1) and c in (0, 1): return rng
        raise XLError(REF)
    nr, nc = rng.nrows, rng.ncols
    if col is MISSING and nr == 1 and nc > 1:
        r, c = 1, r
    elif col is MISSING:
        c = 1 if nc == 1 else 0
    if r < 0 or c < 0 or r > nr or c > nc: raise XLError(REF)
    if r == 0 and c == 0: return rng
    if r == 0: return Range(rng.E, rng.si, rng.r1, rng.c1 + c - 1, rng.r2, rng.c1 + c - 1)
    if c == 0: return Range(rng.E, rng.si, rng.r1 + r - 1, rng.c1, rng.r1 + r - 1, rng.c2)
    return Range(rng.E, rng.si, rng.r1 + r - 1, rng.c1 + c - 1, rng.r1 + r - 1, rng.c1 + c - 1)

def _lookup_vec(a):
    if isinstance(a, (Range, Arr)):
        if a.nrows != 1 and a.ncols != 1: raise XLError(NA)
        return a.vector()
    return [a]

def MATCH(v, arr, typ=1.0):
    v = scal(v)
    t = 1 if typ is MISSING else int(approx_floor(num(scal(typ))))
    vec = _lookup_vec(arr)
    if v is None: v = 0.0  # ô trống tra như 0 (LibreOffice)
    if isinstance(v, bool): v = 1.0 if v else 0.0
    if t == 0:
        if isinstance(v, str):
            if _has_wild(v):
                rx = _wild_re(v)
                for i, x in enumerate(vec):
                    if isinstance(x, str) and rx.match(x): return float(i + 1)
            else:
                lv = v.lower()
                for i, x in enumerate(vec):
                    if isinstance(x, str) and x.lower() == lv: return float(i + 1)
        else:
            for i, x in enumerate(vec):
                if isinstance(x, (int, float)) and approx_equal(float(x), float(v)): return float(i + 1)
        raise XLError(NA)
    # tra gần đúng (dữ liệu đã sắp xếp): tìm nhị phân như LibreOffice, chỉ xét cùng kiểu
    isstr = isinstance(v, str)
    idx = [i for i, x in enumerate(vec) if (isinstance(x, str) if isstr else isinstance(x, (int, float)))]
    lo, hi = 0, len(idx) - 1; best = -1
    while lo <= hi:
        mid = (lo + hi) // 2; x = vec[idx[mid]]
        c = _cmp(x, v)
        if (c <= 0) if t > 0 else (c >= 0):
            best = mid; lo = mid + 1
        else:
            hi = mid - 1
    if best < 0: raise XLError(NA)
    # với giá trị bằng nhau liên tiếp, LibreOffice trả vị trí cuối của dãy bằng
    return float(idx[best] + 1)

# ------------------------------------------------------------------ đếm theo điều kiện
class _Crit:
    __slots__ = ('op', 'val', 'rx', 'isnum')
    def __init__(self, c):
        c = scal(c)
        if c is None or c is MISSING: c = ''
        if isinstance(c, bool): c = 1.0 if c else 0.0
        if not isinstance(c, str):
            self.op = '='; self.val = float(c); self.isnum = True; self.rx = None; return
        op = '='
        for p in ('<=', '>=', '<>', '<', '>', '='):
            if c.startswith(p): op = p; c = c[len(p):]; break
        self.op = op; self.rx = None
        if _num_re.match(c):
            self.val = float(c); self.isnum = True
        else:
            self.val = c; self.isnum = False
            if op in ('=', '<>') and _has_wild(c): self.rx = _wild_re(c)
    def ok(self, x):
        if isinstance(x, ErrVal): return False
        op = self.op
        if isinstance(x, bool): x = 1.0 if x else 0.0
        if self.isnum:
            if type(x) is float or type(x) is int:
                c = _cmp(float(x), self.val)
            elif isinstance(x, str) and _num_re.match(x) and op == '=':
                c = _cmp(float(x), self.val)
            else:
                return op == '<>'
        else:
            if self.val == '':
                empty = x is None or x == ''
                if op == '=': return empty
                if op == '<>': return not empty
                return False
            if x is None: return op == '<>'
            if not isinstance(x, str):
                return op == '<>'
            if self.rx is not None:
                m = bool(self.rx.match(x))
                return m if op == '=' else not m
            c = _cmp(x, self.val)
        return (c == 0 if op == '=' else c != 0 if op == '<>' else c < 0 if op == '<' else
                c <= 0 if op == '<=' else c > 0 if op == '>' else c >= 0)

def COUNTIF(rng, crit):
    cr = _Crit(crit)
    return float(sum(1 for x in _vals(rng) if cr.ok(x)))

def COUNTIFS(*args):
    if len(args) % 2: raise XLError(E_VALUE)
    vecs = []; crs = []
    for i in range(0, len(args), 2):
        vecs.append(_vals(args[i])); crs.append(_Crit(args[i + 1]))
    n = len(vecs[0])
    if any(len(v) != n for v in vecs): raise XLError(E_VALUE)
    k = 0
    for j in range(n):
        if all(crs[i].ok(vecs[i][j]) for i in range(len(vecs))): k += 1
    return float(k)

# ------------------------------------------------------------------ chuẩn hóa kết quả ô
def ket_qua(v):
    """Giá trị cuối của ô công thức (như getDataArray của LibreOffice)."""
    t = type(v)
    if t is float or t is str: return v
    if t is bool: return v          # giữ kiểu Đúng/Sai (hiển thị TRUE/FALSE); getDataArray = 1.0/0.0
    if t is int: return float(v)
    if v is None or v is MISSING: return 0.0
    if isinstance(v, Range): return ket_qua(v.scalar())
    return v

def loi(code):
    """Hằng lỗi trong công thức (#N/A…)."""
    raise XLError(code)


# ------------------------------------------------------------------ TÍNH TOÁN MẢNG (chỉ trong tham số SUMPRODUCT)
# Quét toàn workbook 6.3: trong SUMPRODUCT có toán tử = <> < <= > >= + - * và hàm MOD, MID, ISNUMBER, MATCH, SEARCH;
# ngoài SUMPRODUCT không có phép toán nào trên vùng nhiều ô (không cần giao ngầm).
def _mat(x):
    if isinstance(x, Range):
        E = x.E; si = x.si
        return [[E.raw(si, r, c) for c in range(x.c1, x.c2 + 1)] for r in range(x.r1, x.r2 + 1)]
    if isinstance(x, Arr): return x.rows
    return None

_OPS = {'add': add, 'sub': sub, 'mul': mul, 'div': div, 'pow_': pow_, 'cat': cat,
        'eq': eq, 'ne': ne, 'lt': lt, 'le': le, 'gt': gt, 'ge': ge}

def _elem(f, vals):
    for v in vals:
        if isinstance(v, ErrVal): return v
    try:
        r = f(*vals)
        return r
    except XLError as e:
        return ErrVal(e.code)

def _bcast(f, args):
    mats = [_mat(a) for a in args]
    if all(m is None for m in mats):
        return f(*args)
    nr = max(len(m) for m in mats if m is not None); nc = max(len(m[0]) for m in mats if m is not None)
    out = []
    for i in range(nr):
        row = []
        for j in range(nc):
            vals = []; bad = False
            for a, m in zip(args, mats):
                if m is None: vals.append(a); continue
                ii = 0 if len(m) == 1 else i; jj = 0 if len(m[0]) == 1 else j
                if ii >= len(m) or jj >= len(m[0]): bad = True; break
                vals.append(m[ii][jj])
            row.append(ErrVal(NA) if bad else _elem(f, vals))
        out.append(row)
    return Arr(out)

def A(op, a, b=None):
    """Toán tử trong ngữ cảnh mảng."""
    if op == 'neg': return _bcast(neg, [a])
    return _bcast(_OPS[op], [a, b])

def AF_MOD(a, b): return _bcast(MOD, [a, b])
def AF_MID(a, st, n): return _bcast(MID, [a, st, n])
def AF_SEARCH(f, s, st=MISSING):
    if st is MISSING: return _bcast(lambda x, y: SEARCH(x, y), [f, s])
    return _bcast(SEARCH, [f, s, st])
def AF_MATCH(v, arr, typ=MISSING):
    m = _mat(v)
    if m is None: return MATCH(v, arr, typ)
    return Arr([[_elem(lambda x: MATCH(x, arr, typ), [x]) for x in row] for row in m])
def AF_ISNUMBER(f):
    try:
        x = f()
    except XLError:
        return False
    m = _mat(x)
    if m is None:
        return isinstance(x, (int, float))
    return Arr([[isinstance(v, (int, float)) for v in row] for row in m])
