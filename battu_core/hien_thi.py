# -*- coding: utf-8 -*-
"""Chữ hiển thị của một ô công thức theo định dạng số của ô (như getString() của LibreOffice).
Định dạng lấy từ workbook (hang_so_sheet.json → dinh_dang_o_cong_thuc), không gõ tay."""
from battu_core import xl as X
from battu_core.du_lieu import model


def hien_thi(v, ma='General'):
    if isinstance(v, X.ErrVal): return v.code
    if isinstance(v, str): return v
    if v is None: return ''
    if isinstance(v, bool) and ma == 'General': return 'TRUE' if v else 'FALSE'
    x = float(v)
    if ma == 'General': return X.numstr(x)
    return X.TEXT(x, ma)


def hien_thi_o(E, sheet, ref):
    M = model()
    ma = M.dinh_dang.get(sheet, {}).get(ref, 'General')
    return hien_thi(E.gia_tri(sheet, ref), ma)


def doc_nhu_uno(E, sheet, ref):
    """Như các script oracle: getValue() nếu ô là HẰNG SỐ (getType = VALUE), ngược lại getString()."""
    M = model(); from battu_core.du_lieu import rc
    if rc(ref) not in M.formula_cells[M.si[sheet]]:
        v = E.gia_tri(sheet, ref)
        if isinstance(v, bool): return 1.0 if v else 0.0
        if isinstance(v, (int, float)): return float(v)
        return '' if v is None else v
    return hien_thi_o(E, sheet, ref)
