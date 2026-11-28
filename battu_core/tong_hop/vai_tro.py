# -*- coding: utf-8 -*-
"""Hệ 5 vai của sách (Võ Quý Khanh) – mô-đun BỔ SUNG, không thay quy tắc Thiệu Vĩ Hoa (29/11/2026).

Vòng tương sinh Mộc→Hỏa→Thổ→Kim→Thủy xoay quanh M (đắc địa, trung lập): Y (cân bằng phụ, sinh X), X (cân bằng chính),
M, U (mất cân bằng phụ, sinh Z), Z (mất cân bằng chính). Sách chọn M "khó, cần kinh nghiệm" nên mô-đun KHÔNG tự chọn:
nó tính đủ 5 khả năng; người luận giải chọn một khả năng sau khi đối chiếu sự kiện đã biết (xem doi_chieu_su_kien).
Từ một vai đã chọn, mọi tính toán còn lại là quy tắc xác định:
- Trục tung / hoành theo tháng sinh (tr.47).
- Thịnh/suy từng đại vận: Can vận thuộc {Y, X} → thịnh; {U, Z} → suy; M → đắc địa (chuyển tiếp, xét theo chi vận).
- Năm nguy: vận suy + chi năm xung trục + Can năm thuộc phe mất cân bằng (U, Z); Z khắc X, U khắc Y.
Chỉ dùng thư viện chuẩn; không đọc lá số nào ngoài dữ liệu truyền vào."""

HANH = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']
CAN = ['Giáp', 'Ất', 'Bính', 'Đinh', 'Mậu', 'Kỷ', 'Canh', 'Tân', 'Nhâm', 'Quý']
CHI = ['Tý', 'Sửu', 'Dần', 'Mão', 'Thìn', 'Tị', 'Ngọ', 'Mùi', 'Thân', 'Dậu', 'Tuất', 'Hợi']
CAN_H = dict(zip(CAN, ['Mộc', 'Mộc', 'Hỏa', 'Hỏa', 'Thổ', 'Thổ', 'Kim', 'Kim', 'Thủy', 'Thủy']))
CHI_H = dict(zip(CHI, ['Thủy', 'Thổ', 'Mộc', 'Mộc', 'Thổ', 'Hỏa', 'Hỏa', 'Thổ', 'Kim', 'Kim', 'Thổ', 'Thủy']))
ALIAS = {'Tỵ': 'Tị'}
SINH = {h: HANH[(i + 1) % 5] for i, h in enumerate(HANH)}
CAN_AM_M = {'Mộc': 'Ất', 'Hỏa': 'Đinh', 'Thổ': 'Kỷ', 'Kim': 'Tân', 'Thủy': 'Quý'}
VAI = ('Y', 'X', 'M', 'U', 'Z')
TEN_VAI = {'Y': 'cân bằng phụ', 'X': 'cân bằng chính', 'M': 'đắc địa (trung lập)', 'U': 'mất cân bằng phụ', 'Z': 'mất cân bằng chính'}
# Công thức theo quan hệ của M với Can ngày (tr.124–126): M là Tỷ, Ấn, Thương, Quan, Tài
CT_THEO_M = {'Tỷ': 1, 'Ấn': 2, 'Thương': 3, 'Quan': 4, 'Tài': 5}
CT_TEN = {1: 'Quan + Ấn + Tỷ', 2: 'Tài + Quan + Ấn', 3: 'Ấn + Tỷ + Thương', 4: 'Thương + Tài + Quan', 5: 'Tỷ + Thương + Tài'}
XUNG = {'Tý': 'Ngọ', 'Ngọ': 'Tý', 'Sửu': 'Mùi', 'Mùi': 'Sửu', 'Dần': 'Thân', 'Thân': 'Dần',
        'Mão': 'Dậu', 'Dậu': 'Mão', 'Thìn': 'Tuất', 'Tuất': 'Thìn', 'Tị': 'Hợi', 'Hợi': 'Tị'}
# tháng sinh → (trục tung, trục hoành); bốn nhóm tháng cùng một cặp trục (đã kiểm với các lá số sách tr.47–216)
_NHOM_TRUC = {
    ('Dần', 'Thân', 'Tị', 'Hợi'): (('Dần', 'Thân'), ('Tị', 'Hợi')),
    ('Tý', 'Ngọ', 'Mão', 'Dậu'): (('Tý', 'Ngọ'), ('Mão', 'Dậu')),
    ('Thìn', 'Tuất', 'Sửu', 'Mùi'): (('Sửu', 'Mùi'), ('Thìn', 'Tuất')),
}


def _chuan(c): return ALIAS.get(c, c)


def quan_he_voi_ngay(h_ngay, h):
    if h == h_ngay: return 'Tỷ'
    if SINH[h] == h_ngay: return 'Ấn'
    if SINH[h_ngay] == h: return 'Thương'
    if HANH[(HANH.index(h_ngay) + 2) % 5] == h: return 'Tài'
    return 'Quan'


def truc(chi_thang):
    chi_thang = _chuan(chi_thang)
    for nhom, (tung, hoanh) in _NHOM_TRUC.items():
        if chi_thang in nhom: return tung, hoanh
    return None, None


def vai_theo_M(m):
    i = HANH.index(m)
    return {v: HANH[(i + k) % 5] for v, k in zip(VAI, (-2, -1, 0, 1, 2))}


def luc_hanh(can4, chi4):
    """Lực mỗi hành trong tứ trụ (Can 1 điểm, Chi 1 điểm; tháng lệnh thêm 1 điểm). Chỉ để so hai phe, không thay vượng suy TVH."""
    luc = {h: 0.0 for h in HANH}
    for i, c in enumerate(can4):
        if c: luc[CAN_H[c]] += 1
    for i, c in enumerate(chi4):
        if c:
            luc[CHI_H[_chuan(c)]] += 1 + (1 if i == 1 else 0)
    return luc


def ung_vien(can4, chi4):
    """Năm khả năng M. can4/chi4: [năm, tháng, ngày, giờ] (giờ có thể None). Trả list dict theo thứ tự hành M."""
    chi4 = [_chuan(c) if c else c for c in chi4]
    h_ngay = CAN_H[can4[2]]
    luc = luc_hanh(can4, chi4)
    out = []
    for m in HANH:
        v = vai_theo_M(m)
        pha_cb = luc[v['Y']] + luc[v['X']]; pha_mcb = luc[v['U']] + luc[v['Z']]
        # M bỏ phiếu cho phe yếu hơn? Sách: M bỏ phiếu cho bên nào thì bên đó nắm quyền → đo cả hai khả năng
        cb_manh = luc[m] + pha_cb > pha_mcb          # M về phe cân bằng
        mcb_manh = luc[m] + pha_mcb > pha_cb
        ct = CT_THEO_M[quan_he_voi_ngay(h_ngay, m)]
        out.append({'m': m, 'vai': v, 'cong_thuc': ct, 'ten_cong_thuc': CT_TEN[ct],
                    'luc': {k: luc[h] for k, h in v.items()}, 'cb_manh': cb_manh, 'mcb_manh': mcb_manh,
                    'khong_Y_X': (luc[v['Y']] + luc[v['X']]) == 0})
    return out


def thinh_suy_van(vai, can_van, chi_van):
    """'thinh' | 'suy' | 'dac_dia' cho một đại vận. Dùng Can vận; Can vận là M thì xét chi vận (nếu chi vận cũng M → đắc địa)."""
    inv = {h: k for k, h in vai.items()}
    k = inv[CAN_H[can_van]]
    if k in ('Y', 'X'): return 'thinh'
    if k in ('U', 'Z'): return 'suy'
    k2 = inv[CHI_H[_chuan(chi_van)]]
    return {'Y': 'thinh', 'X': 'thinh', 'U': 'suy', 'Z': 'suy'}.get(k2, 'dac_dia')


def nua_theo_vong(vai, van, can4=None, chi4=None):
    """Thịnh/suy từng đại vận theo HÌNH HỆ TRỤC của sách (H-7, H-88..H-96), thay cho quy tắc theo Can.
    Quy tắc đọc từ hình (29/11/2026, kiểm trên 9 ca sách):
    - Mười hai chi xếp theo thứ tự trên vòng; mỗi đại vận đặt tại CHI của nó.
    - Đường phân định là đường kính qua chi p của đại vận (hoặc trụ sinh) mang Can âm của M, và chi đối của p.
    - Đi ngược chiều tăng của chi từ p (5 chi trước p) là nửa chứa Y, X: thịnh. 5 chi sau p là nửa chứa U, Z: suy.
    - Vận đúng ở đầu đường (chi p, chi đối p) thuộc về nửa mà nó đang bước vào theo chiều đi của đại vận.
    van: list {ten} theo thứ tự thời gian. Trả list nhãn 'thinh'|'suy' kèm cờ 'bien' (vận ở đầu đường)."""
    can_am = CAN_AM_M[next(h for k, h in vai.items() if k == 'M')]
    chi_van = [_chuan(v['ten'].split()[1]) for v in van]
    can_van = [v['ten'].split()[0] for v in van]
    p = next((CHI.index(c) for c, k in zip(chi_van, can_van) if k == can_am), None)
    if p is None and can4 and chi4:
        for k, c in zip(can4, chi4):
            if k == can_am and c: p = CHI.index(_chuan(c)); break
    if p is None: return None
    tang = len(van) < 2 or ((CHI.index(chi_van[1]) - CHI.index(chi_van[0])) % 12) < 6   # chiều thuận: chi tăng dần
    out = []
    for c in chi_van:
        d = (CHI.index(c) - p) % 12
        if 1 <= d <= 5: out.append(('suy', False))
        elif 7 <= d <= 11: out.append(('thinh', False))
        elif d == 0: out.append(('suy' if tang else 'thinh', True))
        else: out.append(('thinh' if tang else 'suy', True))
    return out


def diem_nam(vai, tung_hoanh, can_nam, chi_nam, nua_van, co_mat=None):
    """Điểm nguy của một năm (0–6) và lý do. nua_van: 'thinh' | 'suy' | 'dac_dia'."""
    inv = {h: k for k, h in vai.items()}
    diem, ly = 0, []
    if nua_van == 'suy': diem += 2; ly.append('đang trong vận nửa suy')
    chi_nam = _chuan(chi_nam)
    cap = [x for p in tung_hoanh if p for x in p]
    if chi_nam in cap and XUNG.get(chi_nam) in cap:
        diem += 2; ly.append('chi năm %s đối xung trên trục %s–%s' % (chi_nam, chi_nam, XUNG[chi_nam]))
    kc = inv[CAN_H[can_nam]]
    if kc in ('U', 'Z'): diem += 1; ly.append('Can năm là hành %s (%s)' % (CAN_H[can_nam], kc))
    if co_mat:
        if kc == 'Z' and vai['X'] in co_mat: diem += 2; ly.append('Can năm %s (Z) khắc thẳng %s (X) có trong lá số' % (CAN_H[can_nam], vai['X']))
        if kc == 'U' and vai['Y'] in co_mat: diem += 1; ly.append('Can năm %s (U) khắc %s (Y) có trong lá số' % (CAN_H[can_nam], vai['Y']))
    kh = inv[CHI_H[chi_nam]]
    if kh in ('U', 'Z'): diem += 1; ly.append('Chi năm là hành %s (%s)' % (CHI_H[chi_nam], kh))
    return diem, ly


def can_chi_nam(n): return CAN[(n - 4) % 10], CHI[(n - 4) % 12]


def phan_tich(can4, chi4, dai_van, nam_tu, nam_den, ngong=4):
    """dai_van: list {ten, nam_bat_dau, nam_ket_thuc} (ten = 'Can Chi'). Trả 5 khả năng, mỗi khả năng kèm bảng vận và năm nguy."""
    tung, hoanh = truc(chi4[1])
    co_mat = {CAN_H[c] for c in can4 if c} | {CHI_H[_chuan(c)] for c in chi4 if c}
    kq = []
    for uv in ung_vien(can4, chi4):
        vai = uv['vai']
        van = []
        vong = nua_theo_vong(vai, dai_van, can4, chi4)      # theo hình hệ trục (H-52): vận đặt tại CHI, chia hai nửa bởi đường phân định
        for i, d in enumerate(dai_van):
            c, ch = d['ten'].split()
            if vong is not None:
                nua, bien = vong[i]
            else:
                nua, bien = thinh_suy_van(vai, c, _chuan(ch)), False
            van.append({'ten': d['ten'], 'nam_bat_dau': d['nam_bat_dau'], 'nam_ket_thuc': d['nam_ket_thuc'],
                        'nua': nua, 'bien': bien})
        nam = []
        for n in range(nam_tu, nam_den + 1):
            nua = next((v['nua'] for v in van if v['nam_bat_dau'] <= n <= v['nam_ket_thuc']), 'dac_dia')
            c, ch = can_chi_nam(n)
            diem, ly = diem_nam(vai, (tung, hoanh), c, ch, nua, co_mat)
            nam.append({'nam': n, 'can_chi': '%s %s' % (c, ch), 'diem': diem, 'ly_do': ly, 'nguy': diem >= ngong})
        kq.append(dict(uv, truc_tung=tung, truc_hoanh=hoanh, van=van, nam=nam,
                       phan_dinh='Can âm của M: %s (đường phân định đi qua vận mang Can này và vận đối diện)' % CAN_AM_M[uv['m']]))
    return kq


def doi_chieu_su_kien(kha_nang, su_kien):
    """Đối chiếu sự kiện đã biết với từng khả năng. su_kien: list {nam, loai} với loai 'tot' (thành đạt, thăng tiến, tin vui)
    hoặc 'xau' (bệnh nặng, mất mát, tai nạn, thất bại). Trả điểm khớp 0..n và chi tiết; cao nhất là khả năng đáng chọn."""
    ket = []
    for k in kha_nang:
        khop, ct = 0, []
        for s in su_kien:
            nua = next((v['nua'] for v in k['van'] if v['nam_bat_dau'] <= s['nam'] <= v['nam_ket_thuc']), None)
            nm = next((x for x in k['nam'] if x['nam'] == s['nam']), None)
            if s['loai'] == 'tot': ok = nua in ('thinh', 'dac_dia') and not (nm and nm['nguy'])
            else: ok = nua == 'suy' or bool(nm and nm['nguy'])
            khop += 1 if ok else 0
            ct.append((s['nam'], s['loai'], nua, ok))
        ket.append({'m': k['m'], 'khop': khop, 'tren': len(su_kien), 'chi_tiet': ct})
    return sorted(ket, key=lambda x: -x['khop'])
