# -*- coding: utf-8 -*-
"""Chú giải thuật ngữ – lấy NGUYÊN VĂN từ workbook, không soạn thêm:
  * LUAN_GIAI phần B, cột E "(b) Thuật ngữ bằng lời thường" (dòng 106 trở đi): mọi câu dạng "X là …";
  * các ô ghi chú dạng "X = …" (ví dụ KET_LUAN G17 "Thập thần = quan hệ ngũ hành, âm dương của chữ với nhật can (P1 tr.47-52)").
Mỗi mục: (thuật ngữ, câu giải thích, ô nguồn). Dùng cho chú giải khi rê chuột."""
import re
from battu_core.du_lieu import model, ref_of
from battu_core.hien_thi import hien_thi_o

_CAU = re.compile(r'(?<=[.!?])\s+')
_LA = re.compile(r'^\s*(?:[•▸\-–]\s*)?([A-ZĐÂĂÊÔƠƯÁÀẢÃẠ-ỸÝ][^.;:,()“”"]{1,38}?)\s+là\s+\S')
_BANG = re.compile(r'^\s*([A-ZĐÂĂÊÔƠƯÁÀẢÃẠ-ỸÝ][^.;:,()“”"=]{1,30}?)\s*=\s*\S')
_BO = {'sách coi đây', 'đây', 'đó', 'này', 'mỗi câu', 'cách chấm của sách', 'kết luận', 'dữ kiện', 'khi', 'nếu', 'đó cũng'}


def tu_dien(E):
    """[(thuật ngữ, câu, 'SHEET!ô')] theo thứ tự xuất hiện; mỗi thuật ngữ giữ câu đầu tiên."""
    M = model(); ra = {}; si = M.si['LUAN_GIAI']; fc = M.formula_cells[si]
    nguon = []
    for r in range(106, max(r for r, c in fc) + 1):
        if (r, 5) in fc or (r, 5) in E.consts[si]: nguon.append(('LUAN_GIAI', r, 5))
    for sh, r, c in [('KET_LUAN', r, 7) for r in range(15, 27)]:
        nguon.append((sh, r, c))
    for sh, r, c in nguon:
        si_ = M.si[sh]
        t = hien_thi_o(E, sh, ref_of(r, c)) if (r, c) in M.formula_cells[si_] else str(E.consts[si_].get((r, c)) or '')
        for cau in _CAU.split(t.replace('\n', ' ')):
            m = _LA.match(cau) or _BANG.match(cau)
            if not m: continue
            k = m.group(1).strip(); kl = k.lower()
            if kl in _BO or len(kl) < 2 or kl in ra: continue
            from export.lam_sach import lam_sach
            ra[kl] = (k, lam_sach(cau.strip()), '%s!%s' % (sh, ref_of(r, c)))
    return list(ra.values())


def tim(chu, td, toi_da=4):
    """Các mục từ điển có thuật ngữ xuất hiện trong chữ (không phân biệt hoa thường, nguyên từ)."""
    if not chu: return []
    t = chu.lower(); ra = []
    for k, cau, o in sorted(td, key=lambda x: -len(x[0])):
        if re.search(r'(?<![\wÀ-ỹ])' + re.escape(k.lower()) + r'(?![\wÀ-ỹ])', t):
            if not any(k.lower() in x[0].lower() for x in ra): ra.append((k, cau, o))
        if len(ra) >= toi_da: break
    return ra
