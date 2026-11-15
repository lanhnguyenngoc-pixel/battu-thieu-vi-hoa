# -*- coding: utf-8 -*-
"""Quan hệ can chi – lớp tổng hợp (đề xuất hoàn thiện 15/11/2026, giai đoạn T2, mục B1–B2).

Workbook đã tính tam hợp, bán hợp (cặp kề), lục hợp, xung để chấm lực vượng suy (TRUNG_GIAN, PP2) nhưng không
luận ý nghĩa. Module này liệt kê ĐỦ quan hệ giữa các trụ của tứ trụ, và giữa đại vận / lưu niên với tứ trụ, để lớp
phán đoán và bản luận giải dùng. Chỉ thư viện chuẩn, không phụ thuộc Engine.

Bảng tra theo quy ước phổ biến của Tử Bình (Thiệu Vĩ Hoa P1 tr.53–58, P2 tr.252–256):
  thiên can ngũ hợp, thiên can tương xung; địa chi lục hợp, tam hợp (và bán hợp sinh–vượng, vượng–mộ; củng hợp
  sinh–mộ), tam hội, lục xung, tương hình (vô ân, trì thế, vô lễ, tự hình), lục hại, lục phá, ám hợp, mộ khố.
Hành đánh số như workbook: 0 Mộc, 1 Hỏa, 2 Thổ, 3 Kim, 4 Thủy."""
from itertools import combinations

CAN = ['Giáp', 'Ất', 'Bính', 'Đinh', 'Mậu', 'Kỷ', 'Canh', 'Tân', 'Nhâm', 'Quý']
CHI = ['Tí', 'Sửu', 'Dần', 'Mão', 'Thìn', 'Tị', 'Ngọ', 'Mùi', 'Thân', 'Dậu', 'Tuất', 'Hợi']
HANH = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']
HANH_CAN = [0, 0, 1, 1, 2, 2, 3, 3, 4, 4]
HANH_CHI = [4, 2, 0, 0, 2, 1, 1, 2, 3, 3, 2, 4]
TRU = ['năm', 'tháng', 'ngày', 'giờ']
# Cung của từng trụ (P2 tr.284; P1 tr.149) – can / chi
CUNG = {'năm': 'tổ nghiệp, cha mẹ, thời niên thiếu', 'tháng': 'cha mẹ, anh chị em, thời lập nghiệp',
        'ngày': 'bản thân (can) và vợ/chồng (chi)', 'giờ': 'con cái, hậu vận'}

NGU_HOP = {frozenset(('Giáp', 'Kỷ')): 2, frozenset(('Ất', 'Canh')): 3, frozenset(('Bính', 'Tân')): 4,
           frozenset(('Đinh', 'Nhâm')): 0, frozenset(('Mậu', 'Quý')): 1}
CAN_XUNG = {frozenset(p) for p in (('Giáp', 'Canh'), ('Ất', 'Tân'), ('Bính', 'Nhâm'), ('Đinh', 'Quý'))}
LUC_HOP = {frozenset(('Tí', 'Sửu')): 2, frozenset(('Dần', 'Hợi')): 0, frozenset(('Mão', 'Tuất')): 1,
           frozenset(('Thìn', 'Dậu')): 3, frozenset(('Tị', 'Thân')): 4, frozenset(('Ngọ', 'Mùi')): 2}   # Ngọ–Mùi hợp Thổ (như workbook TRUNG_GIAN dòng 64)
# Tam hợp: (sinh, vượng, mộ) → hành
TAM_HOP = [(('Thân', 'Tí', 'Thìn'), 4), (('Hợi', 'Mão', 'Mùi'), 0), (('Dần', 'Ngọ', 'Tuất'), 1), (('Tị', 'Dậu', 'Sửu'), 3)]
TAM_HOI = [(('Dần', 'Mão', 'Thìn'), 0), (('Tị', 'Ngọ', 'Mùi'), 1), (('Thân', 'Dậu', 'Tuất'), 3), (('Hợi', 'Tí', 'Sửu'), 4)]
LUC_XUNG = {frozenset(p) for p in (('Tí', 'Ngọ'), ('Sửu', 'Mùi'), ('Dần', 'Thân'), ('Mão', 'Dậu'), ('Thìn', 'Tuất'), ('Tị', 'Hợi'))}
LUC_HAI = {frozenset(p) for p in (('Tí', 'Mùi'), ('Sửu', 'Ngọ'), ('Dần', 'Tị'), ('Mão', 'Thìn'), ('Thân', 'Hợi'), ('Dậu', 'Tuất'))}
LUC_PHA = {frozenset(p) for p in (('Tí', 'Dậu'), ('Ngọ', 'Mão'), ('Thân', 'Tị'), ('Dần', 'Hợi'), ('Thìn', 'Sửu'), ('Tuất', 'Mùi'))}
HINH_BA = [(('Dần', 'Tị', 'Thân'), 'hình vô ân'), (('Sửu', 'Tuất', 'Mùi'), 'hình trì thế')]
HINH_HAI = {frozenset(('Tí', 'Mão')): 'hình vô lễ'}
TU_HINH = {'Thìn', 'Ngọ', 'Dậu', 'Hợi'}
# Ám hợp: bản khí tàng can của hai chi ngũ hợp (ba cặp thường dùng)
AM_HOP = {frozenset(('Dần', 'Sửu')): 'Giáp–Kỷ', frozenset(('Ngọ', 'Hợi')): 'Đinh–Nhâm', frozenset(('Mão', 'Thân')): 'Ất–Canh'}
MO_KHO = {'Thìn': 4, 'Tuất': 1, 'Sửu': 3, 'Mùi': 0}           # chi → hành được chứa (kho)
CAN_KHAC = lambda a, b: (HANH_CAN[CAN.index(b)] - HANH_CAN[CAN.index(a)]) % 5 == 2   # a khắc b


def _vi(i, ten=None):
    return ten[i] if ten else TRU[i]


def _q(loai, ten, vi_tri, chu, hanh=None, ghi_chu='', ke=None):
    d = {'loai': loai, 'ten': ten, 'vi_tri': list(vi_tri), 'chu': list(chu)}
    if hanh is not None: d['hanh'] = HANH[hanh]
    if ke is not None: d['ke_nhau'] = ke
    if ghi_chu: d['ghi_chu'] = ghi_chu
    return d


def _cap_can(a, b):
    """Quan hệ hai thiên can → [(loại, tên, hành|None, ghi chú)]."""
    ra = []
    k = frozenset((a, b))
    if a != b and k in NGU_HOP: ra.append(('can_hop', 'ngũ hợp %s–%s' % (a, b), NGU_HOP[k], 'hợp hóa %s' % HANH[NGU_HOP[k]]))
    if k in CAN_XUNG: ra.append(('can_xung', 'thiên can tương xung %s–%s' % (a, b), None, ''))
    elif a != b and k not in NGU_HOP and (CAN_KHAC(a, b) or CAN_KHAC(b, a)):     # cặp ngũ hợp: luận hợp, không luận khắc
        x, y = (a, b) if CAN_KHAC(a, b) else (b, a)
        ra.append(('can_khac', '%s khắc %s' % (x, y), None, ''))
    return ra


def _cap_chi(a, b):
    """Quan hệ hai địa chi → [(loại, tên, hành|None, ghi chú)]."""
    ra = []; k = frozenset((a, b))
    if a == b:
        if a in TU_HINH: ra.append(('hinh', 'tự hình %s–%s' % (a, b), None, 'hình'))
        return ra
    if k in LUC_HOP: ra.append(('luc_hop', 'lục hợp %s–%s' % (a, b), LUC_HOP[k], ''))
    for (s, v, m), h in TAM_HOP:
        if a in (s, v, m) and b in (s, v, m):
            if v in (a, b): ra.append(('ban_hop', 'bán hợp %s–%s (%s cục)' % (a, b, HANH[h]), h, 'sinh–vượng' if s in (a, b) else 'vượng–mộ'))
            else: ra.append(('cung_hop', 'củng hợp %s–%s (%s cục, thiếu %s)' % (a, b, HANH[h], v), h, 'sinh–mộ, lực yếu'))
    if k in LUC_XUNG: ra.append(('xung', 'lục xung %s–%s' % (a, b), None, ''))
    for ba, ten in HINH_BA:
        if a in ba and b in ba: ra.append(('hinh', '%s %s–%s' % (ten, a, b), None, ten))
    if k in HINH_HAI: ra.append(('hinh', '%s %s–%s' % (HINH_HAI[k], a, b), None, HINH_HAI[k]))
    if k in LUC_HAI: ra.append(('hai', 'lục hại %s–%s' % (a, b), None, ''))
    if k in LUC_PHA: ra.append(('pha', 'lục phá %s–%s' % (a, b), None, ''))
    if k in AM_HOP: ra.append(('am_hop', 'ám hợp %s–%s (%s)' % (a, b, AM_HOP[k]), None, 'tàng can bản khí ngũ hợp'))
    return ra


def trong_tu_tru(tru):
    """tru = [(can, chi)] × 4 (năm, tháng, ngày, giờ) → danh sách quan hệ trong tứ trụ."""
    ra = []
    for i, j in combinations(range(4), 2):
        ke = j == i + 1
        for loai, ten, h, gc in _cap_can(tru[i][0], tru[j][0]):
            ra.append(_q(loai, ten, ('can ' + TRU[i], 'can ' + TRU[j]), (tru[i][0], tru[j][0]), h, gc, ke))
        for loai, ten, h, gc in _cap_chi(tru[i][1], tru[j][1]):
            ra.append(_q(loai, ten, ('chi ' + TRU[i], 'chi ' + TRU[j]), (tru[i][1], tru[j][1]), h, gc, ke))
        if (frozenset((tru[i][1], tru[j][1])) in LUC_XUNG and (frozenset((tru[i][0], tru[j][0])) in CAN_XUNG
                                                             or CAN_KHAC(tru[i][0], tru[j][0]) or CAN_KHAC(tru[j][0], tru[i][0]))):
            ra.append(_q('thien_khac_dia_xung', 'thiên khắc địa xung trụ %s – trụ %s' % (TRU[i], TRU[j]),
                         ('trụ ' + TRU[i], 'trụ ' + TRU[j]), ('%s %s' % tru[i], '%s %s' % tru[j]), ke=ke))
    chi = [c for _, c in tru]
    for ba, h in TAM_HOP:
        if all(x in chi for x in ba):
            ra.append(_q('tam_hop', 'tam hợp %s (%s cục)' % ('–'.join(ba), HANH[h]), ['chi ' + TRU[chi.index(x)] for x in ba], ba, h))
    for ba, h in TAM_HOI:
        if all(x in chi for x in ba):
            ra.append(_q('tam_hoi', 'tam hội %s (%s phương)' % ('–'.join(ba), HANH[h]), ['chi ' + TRU[chi.index(x)] for x in ba], ba, h))
    for ba, ten in HINH_BA:
        if all(x in chi for x in ba):
            ra.append(_q('tam_hinh', 'tam hình %s (%s)' % ('–'.join(ba), ten), ['chi ' + TRU[chi.index(x)] for x in ba], ba))
    for i, c in enumerate(chi):
        if c in MO_KHO:
            ra.append(_q('kho', '%s là kho %s' % (c, HANH[MO_KHO[c]]), ('chi ' + TRU[i],), (c,), MO_KHO[c]))
    return ra


def voi_tu_tru(tru, can, chi, nhan):
    """Quan hệ của một trụ ngoài (đại vận / lưu niên / tiểu vận, tên = nhan) với từng trụ của tứ trụ."""
    ra = []
    for i in range(4):
        for loai, ten, h, gc in _cap_can(can, tru[i][0]):
            ra.append(_q(loai, ten, ('can ' + nhan, 'can ' + TRU[i]), (can, tru[i][0]), h, gc))
        for loai, ten, h, gc in _cap_chi(chi, tru[i][1]):
            ra.append(_q(loai, ten, ('chi ' + nhan, 'chi ' + TRU[i]), (chi, tru[i][1]), h, gc))
        if chi == tru[i][1] and can == tru[i][0]:
            ra.append(_q('phuc_ngam', 'phục ngâm trụ %s (%s %s)' % (TRU[i], can, chi), (nhan, 'trụ ' + TRU[i]), (can + ' ' + chi,)))
        elif chi == tru[i][1]:
            ra.append(_q('trung_chi', 'chi %s trùng chi trụ %s (phục ngâm chi)' % (chi, TRU[i]), ('chi ' + nhan, 'chi ' + TRU[i]), (chi,)))
        if frozenset((chi, tru[i][1])) in LUC_XUNG and (frozenset((can, tru[i][0])) in CAN_XUNG or CAN_KHAC(can, tru[i][0]) or CAN_KHAC(tru[i][0], can)):
            ra.append(_q('thien_khac_dia_xung', 'thiên khắc địa xung (phản ngâm) trụ %s' % TRU[i], (nhan, 'trụ ' + TRU[i]),
                         (can + ' ' + chi, '%s %s' % tru[i])))
    # tam hợp / tam hội / tam hình được trụ ngoài làm đủ
    chi_tt = [c for _, c in tru]
    for nhom, loai, nhan_nhom in ((TAM_HOP, 'tam_hop', 'tam hợp'), (TAM_HOI, 'tam_hoi', 'tam hội')):
        for ba, h in nhom:
            if chi in ba and all(x in chi_tt or x == chi for x in ba) and not all(x in chi_tt for x in ba):
                ra.append(_q(loai, '%s %s (%s) – %s làm đủ' % (nhan_nhom, '–'.join(ba), HANH[h], nhan), [nhan] + [
                    'chi ' + TRU[chi_tt.index(x)] for x in ba if x != chi], ba, h))
    for ba, ten in HINH_BA:
        if chi in ba and all(x in chi_tt or x == chi for x in ba) and not all(x in chi_tt for x in ba):
            ra.append(_q('tam_hinh', 'tam hình %s (%s) – %s làm đủ' % ('–'.join(ba), ten, nhan), [nhan], ba))
    return ra


def giua_hai(can1, chi1, nhan1, can2, chi2, nhan2):
    """Quan hệ giữa hai trụ ngoài (vd lưu niên với đại vận)."""
    ra = []
    for loai, ten, h, gc in _cap_can(can1, can2): ra.append(_q(loai, ten, ('can ' + nhan1, 'can ' + nhan2), (can1, can2), h, gc))
    for loai, ten, h, gc in _cap_chi(chi1, chi2): ra.append(_q(loai, ten, ('chi ' + nhan1, 'chi ' + nhan2), (chi1, chi2), h, gc))
    return ra


def danh_dau(qh, tru, dung, hy, ky):
    """Gắn hệ quả: chi/can nào mang hành dụng, hỷ, kỵ; cung nào bị chạm. dung/hy/ky = tên hành ('Hỏa', …).
    Chi mang hành theo bản khí (HANH_CHI); can theo HANH_CAN."""
    vai = {}
    for ten_vai, hh in (('dụng', dung), ('hỷ', hy), ('kỵ', ky)):
        if hh in HANH: vai.setdefault(HANH.index(hh), ten_vai)
    for q in qh:
        cham = []
        for vt in q['vi_tri']:
            p = vt.split(' ', 1)
            if len(p) == 2 and p[1] in TRU:
                i = TRU.index(p[1])
                if p[0] == 'chi':
                    v = vai.get(HANH_CHI[CHI.index(tru[i][1])])
                    cham.append({'vi_tri': vt, 'cung': CUNG[p[1]], 'vai_tro': v})
                elif p[0] == 'can':
                    v = vai.get(HANH_CAN[CAN.index(tru[i][0])])
                    cham.append({'vi_tri': vt, 'cung': CUNG[p[1]], 'vai_tro': 'nhật chủ' if p[1] == 'ngày' else v})
                elif p[0] == 'trụ':
                    cham.append({'vi_tri': vt, 'cung': CUNG[p[1]]})
        if cham: q['cham'] = cham
        q['cham_dung'] = any(x.get('vai_tro') == 'dụng' for x in cham)
        q['cham_cung_phu_the'] = any(x['vi_tri'] in ('chi ngày', 'trụ ngày') for x in cham)
    return qh
