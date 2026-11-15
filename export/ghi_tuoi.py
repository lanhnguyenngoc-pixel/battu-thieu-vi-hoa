# -*- coding: utf-8 -*-
"""Ghi kèm năm dương lịch cho khoảng tuổi đại vận khi HIỂN THỊ / IN (đề xuất hoàn thiện 15/11/2026, mục C3).

Workbook ghi tuổi đại vận theo năm (năm − năm sinh, vd "Mậu Thìn (43–52 tuổi)") nhưng tuổi lưu niên theo tuổi mụ
("2026 … 49 tuổi mụ"), nên người đọc dễ lẫn. Lớp hiển thị thêm năm dương lịch của vận lấy từ DAI_VAN (dòng 35–38):
"Mậu Thìn (43–52 tuổi, 2021–2030)". Giá trị ô không đổi."""
import re

_DONG_TEN, _DONG_TUOI, _DONG_NAM_DAU, _DONG_NAM_CUOI = 35, 36, 37, 38
_COT = range(3, 13)          # C..L – 10 vận


def _dai_van(E):
    dv = getattr(E, '_ghi_tuoi_dv', None)
    if dv is not None: return dv
    si = E.M.si['DAI_VAN']; dv = []
    for c in _COT:
        try:
            ten, tuoi, y1, y2 = (E.raw(si, r, c) for r in (_DONG_TEN, _DONG_TUOI, _DONG_NAM_DAU, _DONG_NAM_CUOI))
        except Exception:
            continue
        if isinstance(ten, str) and ten and all(isinstance(x, float) for x in (tuoi, y1, y2)):
            dv.append((ten, int(tuoi), int(y1), int(y2)))
    try: E._ghi_tuoi_dv = dv
    except Exception: pass
    return dv


def them_nam(E, t):
    """'<Đại vận> (a–b tuổi)' hoặc '<Đại vận> (a–b)' → thêm ', năm đầu–năm cuối' trước ')' khi a là tuổi bắt đầu của vận."""
    if not isinstance(t, str) or '(' not in t: return t
    for ten, tuoi, y1, y2 in _dai_van(E):
        if ten not in t: continue
        t = re.sub(r'(%s \(%d\s*[–-]\s*\d+(?: tuổi)?)(?=[):])' % (re.escape(ten), tuoi), r'\g<1>, %d–%d' % (y1, y2), t)
    return t
