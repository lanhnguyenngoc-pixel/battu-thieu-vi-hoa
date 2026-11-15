# -*- coding: utf-8 -*-
"""Khôi phục chữ bị cắt cụt khi HIỂN THỊ / IN (đề xuất hoàn thiện 15/11/2026, mục C2).

Một số công thức workbook rút gọn chữ: IF(LEN(x)>N, LEFT(x,N)&"…", x) – LUAN_GIAI, THEO_DOI, KHONG_RO_GIO.
Giá trị ô vẫn đúng workbook (JSON, XLSX, kiểm thử không đổi). Chỉ khi hiện chữ cho người đọc (màn hình, PDF, in),
ô có chữ bị cắt được tính lại bằng một BẢN SAO của hàm công thức, trong đó mọi đoạn X.cat(X.LEFT(x, N), '…')
được thay bằng x. Bản sao dựng bằng ast từ mã sheet đã sinh (battu_core/sheets/*.py); các tệp đó không bị sửa."""
import ast, importlib, inspect, textwrap

_BAN_SAO = {}          # sheet → {(dòng, cột): (hàm không cắt, r, c)}


def _la_X(n, ten):
    return isinstance(n, ast.Attribute) and n.attr == ten and isinstance(n.value, ast.Name) and n.value.id == 'X'


class _BoCat(ast.NodeTransformer):
    """X.cat(X.LEFT(x, N), '…')  →  x"""
    so = 0

    def visit_Call(self, n):
        self.generic_visit(n)
        if (_la_X(n.func, 'cat') and len(n.args) == 2 and isinstance(n.args[1], ast.Constant) and n.args[1].value == '…'
                and isinstance(n.args[0], ast.Call) and _la_X(n.args[0].func, 'LEFT') and len(n.args[0].args) == 2):
            self.so += 1
            return n.args[0].args[0]
        return n


def _sinh_san():
    """Bảng sinh sẵn export/_khong_cat_sinh.py (tools/sinh_khong_cat.py) nếu khớp nguồn sheet hiện tại; không thì None.
    Bộ cài không kèm mã nguồn .py của sheet nên phải dùng bảng này."""
    from battu_core import sheets as S
    try:
        from export import _khong_cat_sinh as G
    except ImportError:
        return None
    return G.BAN_SAO if getattr(G, 'NGUON_SHA256', None) == S.NGUON_SHA256 else None


def _tu_ma_nguon(m, fn):
    """Dựng bản sao không cắt từ mã nguồn (khi chạy từ mã nguồn mà chưa sinh lại bảng)."""
    cay = ast.parse(textwrap.dedent(inspect.getsource(fn)))
    bc = _BoCat(); cay = ast.fix_missing_locations(bc.visit(cay))
    if not bc.so: return None
    ns = {}
    exec(compile(cay, '<khong_cat %s.%s>' % (m.__name__, fn.__name__), 'exec'), m.__dict__.copy(), ns)
    return ns[fn.__name__]


def _ban_sao(sheet):
    if sheet in _BAN_SAO: return _BAN_SAO[sheet]
    from battu_core import sheets as S
    ten = S.MODULES.get(sheet); kq = {}
    if ten:
        m = importlib.import_module('battu_core.sheets.' + ten)
        sinh = _sinh_san()
        sao = dict(sinh.get(sheet, {})) if sinh is not None else {}
        for ref, (fn, r, c) in m.CELLS.items():
            if fn.__name__ not in sao:
                if sinh is not None: sao[fn.__name__] = None
                else:
                    try: sao[fn.__name__] = _tu_ma_nguon(m, fn)
                    except (OSError, TypeError, SyntaxError): sao[fn.__name__] = None     # không có mã nguồn: hiện như workbook
            if sao[fn.__name__]: kq[(r, c)] = (sao[fn.__name__], r, c)
    _BAN_SAO[sheet] = kq
    return kq


class _EKhongCat:
    """Lõi tính "không cắt" bọc quanh Engine: ô có công thức rút gọn chữ được tính bằng bản sao không cắt; ô chép
    chữ từ ô đã bị cắt (LUAN_GIAI C18 = TRIM(J18)) được tính lại qua lõi này. Ô không chứa "…" đọc thẳng Engine gốc."""

    def __init__(self, E):
        self.E = E; self.cache = {}

    def __getattr__(self, k): return getattr(self.E, k)

    def raw(self, si, r, c):
        from battu_core import xl as X
        key = (si, r, c)
        if key in self.cache: return self.cache[key]
        v0 = self.E.raw(si, r, c)
        # Chỉ ô có chữ "…" mới có thể mang chữ bị cắt (của chính nó hoặc chép từ ô nguồn) → tính lại; ô khác dùng nguyên.
        if not (isinstance(v0, str) and '…' in v0) or key in self.E.dt: return v0
        x = _ban_sao(self.E.M.sheets[si]).get((r, c)) or (self.E.F.get(si) or {}).get((r, c))
        if not x: return v0
        try: v = X.ket_qua(x[0](self, x[1], x[2]))          # như Engine.raw
        except X.XLError as e: v = X.ErrVal(e.code)
        except Exception: v = v0
        self.cache[key] = v
        return v

    def v(self, si, r, c):
        from battu_core import xl as X
        v = self.raw(si, r, c)
        if type(v) is X.ErrVal: raise X.XLError(v.code)
        return v

    def rg(self, si, r1, c1, r2, c2):
        from battu_core import xl as X
        return X.Range(self, si, r1, c1, r2, c2)


def _ek(E):
    ek = getattr(E, '_khong_cat', None)
    if ek is None:
        ek = _EKhongCat(E)
        try: E._khong_cat = ek
        except Exception: pass
    return ek


def khoi_phuc(E, sheet, r, c, t):
    """t = chữ hiển thị của ô (sheet, r, c). Nếu ô có chữ bị cắt "…", trả chữ đầy đủ (tính lại không cắt)."""
    if not isinstance(t, str) or '…' not in t: return t
    v = _ek(E).raw(E.M.si[sheet], r, c)
    return v if isinstance(v, str) and v else t


def so_o_co_cat(sheet):
    """Số ô của sheet có công thức rút gọn chữ (dùng cho kiểm thử)."""
    return len(_ban_sao(sheet))
