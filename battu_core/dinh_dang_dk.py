# -*- coding: utf-8 -*-
"""Áp Định dạng có điều kiện của workbook (battu_core/sheets/_dinh_dang_dk.py – dịch nguyên văn).
Mọi quy tắc đúng đều được áp; thuộc tính trùng thì quy tắc ưu tiên cao hơn (số priority nhỏ hơn) thắng.
Điều kiện ra lỗi hoặc FALSE → không áp (như Excel/LibreOffice)."""
from battu_core import xl as X
from battu_core.sheets._dinh_dang_dk import QUY_TAC


def _dung(g, E, r, c):
    try:
        v = g(E, r, c)
        if isinstance(v, X.ErrVal) or isinstance(v, X.XLError): return False
        return bool(X.truthy(v))
    except Exception:
        return False


def kieu_sheet(E, sheet):
    """{(r, c): {'dam': True, 'chu': '#c00000', 'nen': '#ffe699'}} cho các ô thỏa điều kiện của sheet."""
    ra = {}
    for sh, vung, ut, g, kieu in sorted(QUY_TAC, key=lambda q: -q[2]):     # ưu tiên thấp áp trước, cao ghi đè
        if sh != sheet: continue
        for c1, r1, c2, r2 in vung:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if _dung(g, E, r, c): ra.setdefault((r, c), {}).update(kieu)
    return ra


def co_quy_tac(sheet):
    return any(q[0] == sheet for q in QUY_TAC)
