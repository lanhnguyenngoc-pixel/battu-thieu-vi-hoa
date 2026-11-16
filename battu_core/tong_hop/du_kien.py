# -*- coding: utf-8 -*-
"""Hồ sơ dữ kiện – lớp tổng hợp (đề xuất hoàn thiện 15/11/2026, giai đoạn T2).

Rút từ kết quả Engine (đã khớp workbook 100%) những dữ kiện mà lớp phán đoán và bản luận giải cần, thành một dict
thuần (ghi được JSON). Chỉ ĐỌC giá trị ô; ô được tìm theo NHÃN của workbook (cột B của KET_LUAN, DAI_VAN; dòng tiêu đề
29 của LUU_NIEN) để không phụ thuộc số dòng khi workbook đổi nhỏ. Thêm quan hệ can chi (tong_hop.quan_he).

    ho_so(E, la_so=None, so_nam=11) → dict"""
import re
from battu_core.tong_hop import quan_he as Q

PHIEN_BAN_HO_SO = 1
_TRU = Q.TRU
# Mục KET_LUAN theo lĩnh vực (như export.noi_dung.MUC_KL; chép lại để battu_core chỉ dùng thư viện chuẩn)
_MUC_KL = [('B. Lá số', 14), ('C. Bản mệnh', 28), ('D. Tính cách – năng khiếu', 67), ('E. Tài lộc', 78),
           ('F. Sự nghiệp – công danh', 92), ('G. Gia đạo', 107), ('H. Sức khỏe', 134), ('I. Vận trình', 159),
           ('J. Hóa giải', 178), ('K. Độ tin cậy', 192)]


def _s(v):
    return v if isinstance(v, str) else ('' if v is None else v)


def _nhan_cot(E, sheet, cot_nhan, cot_gt, r1, r2):
    """{nhãn cột cot_nhan: giá trị cột cot_gt} cho các dòng r1..r2 (nhãn trùng: giữ lần đầu)."""
    si = E.M.si[sheet]; ra = {}
    for r in range(r1, r2 + 1):
        k = E.raw(si, r, cot_nhan)
        if isinstance(k, str) and k.strip() and k.strip() not in ra:
            ra[k.strip()] = (r, _s(E.raw(si, r, cot_gt)))
    return ra


def _hanh_dau(t):
    """'Hỏa (Thực thương)' → 'Hỏa'."""
    m = re.match(r'\s*(Mộc|Hỏa|Thổ|Kim|Thủy)', t or '')
    return m.group(1) if m else ''


def _tu_tru(E):
    si = E.M.si['KET_LUAN']; kl = _nhan_cot(E, 'KET_LUAN', 2, 3, 14, 27)
    dong = {k: r for k, (r, _) in kl.items()}
    def o(nhan, c):
        r = next((rr for k, rr in dong.items() if k.startswith(nhan)), None)
        return _s(E.raw(si, r, c)) if r else ''
    tru = []
    for i, c in enumerate(range(3, 7)):
        cc = o('Tứ trụ', c).split()
        tru.append({'tru': _TRU[i], 'can': cc[0] if cc else '', 'chi': cc[1] if len(cc) > 1 else '',
                    'thap_than_can': o('Thiên can – thập thần', c), 'hanh_can': o('Hành – âm dương', c),
                    'chi_hanh': o('Địa chi', c), 'tang_ban_khi': o('Tàng can bản khí', c),
                    'tang_trung_khi': o('Tàng can trung khí', c), 'tang_du_khi': o('Tàng can dư khí', c),
                    'truong_sinh': o('Nhật can tại các chi', c)})
    nhom = o('Thập thần theo 5 nhóm', 3)
    return tru, nhom


def _ban_menh(E):
    kl = {k: v for k, (_, v) in _nhan_cot(E, 'KET_LUAN', 2, 3, 28, 66).items()}
    g = lambda *ks: next((kl[k] for k in kl if any(k.startswith(x) for x in ks)), '')
    pp2 = g('Diễn giải PP2')
    m = re.search(r'=\s*([\d.]+)\s*độ\s*=\s*([\d.]+)%', pp2) or re.search(r'([\d.]+)%\s*tổng', pp2)
    phe_minh = float(m.group(m.lastindex)) if m else None
    do = {h: float(x) for h, x in re.findall(r'(Mộc|Hỏa|Thổ|Kim|Thủy)\s+([\d.]+)\s*\(×', pp2)}
    dung, hy, ky = g('DỤNG THẦN'), g('HỶ THẦN'), g('KỴ THẦN')
    return {'ket_luan': g('KẾT LUẬN BẢN MỆNH'), 'vuong_suy': g('VƯỢNG SUY – KHUYẾN NGHỊ'),
            'pp1': g('Trạng thái nhật chủ'), 'phe_minh_pt': phe_minh, 'do_theo_hanh': do,
            'cach_cuc': g('Cách cục'), 'dung_than': dung, 'hy_than': hy, 'ky_than': ky,
            'hanh_dung': _hanh_dau(dung), 'hanh_hy': _hanh_dau(hy), 'hanh_ky': _hanh_dau(ky),
            'vi_sao_dung': g('Vì sao chọn dụng'), 'dieu_hau': g('Điều hầu'), 'menh_cuc': g('Kết luận mệnh cục'),
            'ung_vien_dung': _s(E.gia_tri('DUNG_HY_KY', 'C226')), 'goi_y_su_kien': _s(E.gia_tri('SU_KIEN', 'B38')), 'tho_yeu_kl': _s(E.gia_tri('MENH_CUC', 'C126')), 'tho_yeu_ly_do': _s(E.gia_tri('MENH_CUC', 'C127')), 'tho_yeu_buoc1': _s(E.gia_tri('MENH_CUC', 'C114')), 'tho_yeu_bang173': _s(E.gia_tri('MENH_CUC', 'U21')), 'muc_chac_dung': _s(E.gia_tri('DUNG_HY_KY', 'C227')),
            'than_sat_hieu_luc': g('Có hiệu lực'), 'than_sat_vo_hieu': g('Có nhưng vô hiệu'),
            'quy_nhan': g('QUÝ NHÂN'), 'van_xuong': g('VĂN XƯƠNG'), 'thong_can': g('Kết luận thông căn')}


def _ket_luan_muc(E):
    """Mỗi lĩnh vực của KET_LUAN (C–K): {tên mục: {nhãn: chữ}}; và tóm tắt B4:C12."""
    si = E.M.si['KET_LUAN']
    tom = {_s(E.raw(si, r, 2)): _s(E.raw(si, r, 3)) for r in range(4, 13) if _s(E.raw(si, r, 2))}
    muc = {}
    for i, (ten, r1) in enumerate(_MUC_KL):
        r2 = _MUC_KL[i + 1][1] - 1 if i + 1 < len(_MUC_KL) else 199
        muc[ten] = {k: v for k, (_, v) in _nhan_cot(E, 'KET_LUAN', 2, 3, r1, r2).items() if v not in ('', None)}
    return tom, muc


_DV_NHAN = ['ĐẠI VẬN', 'Tuổi bắt đầu', 'Năm bắt đầu', 'Năm kết thúc', 'Thập thần của can vận', 'Tàng can của chi vận',
            'Nhật can tại chi vận', 'Đang hành']


def _dai_van(E):
    si = E.M.si['DAI_VAN']; dong = _nhan_cot(E, 'DAI_VAN', 2, 2, 30, 200)
    def r_(nhan, thu=0):
        ds = [r for k, (r, _) in sorted(dong.items(), key=lambda x: x[1][0]) if k.startswith(nhan)]
        return ds[thu] if len(ds) > thu else None
    # "XU THẾ ĐẠI VẬN" xuất hiện ở C1 (PP1) rồi C2 (PP2): lấy theo thứ tự dòng
    xt = sorted({r for r in range(30, 121) if isinstance(E.raw(si, r, 2), str) and E.raw(si, r, 2).startswith('XU THẾ ĐẠI VẬN')})
    ra = []
    for c in range(3, 13):
        ten = _s(E.raw(si, r_('ĐẠI VẬN'), c)) if r_('ĐẠI VẬN') else ''
        if not ten: continue
        d = {'ten': ten, 'can': ten.split()[0], 'chi': ten.split()[-1]}
        for k, nhan in (('tuoi_bat_dau', 'Tuổi bắt đầu'), ('nam_bat_dau', 'Năm bắt đầu'), ('nam_ket_thuc', 'Năm kết thúc'),
                        ('thap_than_can', 'Thập thần của can vận'), ('tang_can_chi', 'Tàng can của chi vận'),
                        ('truong_sinh', 'Nhật can tại chi vận')):
            rr = r_(nhan); v = E.raw(si, rr, c) if rr else None
            d[k] = int(v) if isinstance(v, float) else _s(v)
        rr = r_('Đang hành'); d['dang_hanh'] = bool(rr and _s(E.raw(si, rr, c)))
        # Kết luận đại vận theo cách của tác giả (P1 tr.124–125, đề xuất DV-1 đã duyệt) – DAI_VAN mục F
        for k, nhan in (('ket_luan', 'KẾT LUẬN ĐẠI VẬN'), ('diem_dau', 'Điểm 5 năm đầu (phương pháp khuyến nghị)'),
                        ('diem_sau', 'Điểm 5 năm sau (phương pháp khuyến nghị)'), ('tom_tat_10_nam', 'Tóm tắt 10 năm'),
                        ('vi_sao_dau', 'Vì sao – 5 năm đầu'), ('vi_sao_sau', 'Vì sao – 5 năm sau'),
                        ('tai_loc_van', 'Tài lộc trong vận'), ('cong_danh_van', 'Công danh trong vận')):
            rr = next((r for kk, (r, _) in dong.items() if kk == nhan or (len(nhan) > 20 and kk.startswith(nhan))), None)
            v = E.raw(si, rr, c) if rr else None
            d[k] = (int(v) if k.startswith('diem') and isinstance(v, float) else _s(v))
        for i, r in enumerate(xt[:2]): d['xu_the_pp%d' % (i + 1)] = _s(E.raw(si, r, c))
        ra.append(d)
    return ra


_LN_COT = {'Năm': 'nam', 'Tuổi thực': 'tuoi_thuc', 'Tuổi mụ': 'tuoi_mu', 'Đại vận': 'dai_van', 'Lưu niên': 'luu_nien',
           'Tiểu vận': 'tieu_van', 'Thập thần can LN': 'thap_than_can', 'PP1: XU THẾ LƯU NIÊN': 'xu_the_nam_pp1',
           'PP1: TỔNG HỢP (P1 tr.133)': 'tong_hop_pp1', 'PP2: XU THẾ LƯU NIÊN': 'xu_the_nam_pp2',
           'PP2: TỔNG HỢP (P1 tr.133)': 'tong_hop_pp2', 'Số dấu hiệu nặng (để sắp xếp)': 'so_dau_hieu_nang',
           'SỰ KIỆN / DẤU HIỆU CỦA NĂM': 'dau_hieu', 'TẠNG PHỦ CẦN CHÚ Ý TRONG NĂM (P2 tr.558-': 'tang_phu',
           'Việc ứng (P1 tr.98)': 'viec_ung', 'Mô tả năm (G90)': 'mo_ta', 'Điểm năm (G90)': 'diem_nam',
           'Lợi kết hôn': 'loi_ket_hon', 'Bất lợi hôn nhân': 'bat_loi_hon_nhan', 'Dấu hiệu bất lợi cho cha': 'bat_loi_cha',
           'Dấu hiệu bất lợi cho mẹ': 'bat_loi_me'}


def _luu_nien(E, nam_dau, so_nam):
    si = E.M.si['LUU_NIEN']
    cot = {}
    for c in range(1, 130):
        h = E.raw(si, 29, c)
        if isinstance(h, str):
            for k, ten in _LN_COT.items():
                # nhãn ngắn so TRÙNG KHỚP ('Đại vận' ≠ 'Đại vận số'); nhãn dài (có thể bị cắt trong workbook) so phần đầu
                if (h.strip() == k or (len(k) > 14 and h.strip().startswith(k))) and ten not in cot.values(): cot[c] = ten
    c_nam = next(c for c, t in cot.items() if t == 'nam')
    ra = []
    for r in range(30, 130):
        n = E.raw(si, r, c_nam)
        if isinstance(n, float) and nam_dau <= n < nam_dau + so_nam:
            d = {}
            for c, ten in cot.items():
                v = E.raw(si, r, c)
                d[ten] = int(v) if isinstance(v, float) and ten in ('nam', 'tuoi_thuc', 'tuoi_mu', 'so_dau_hieu_nang', 'loi_ket_hon',
                                                                    'bat_loi_hon_nhan', 'bat_loi_cha', 'bat_loi_me') else _s(v)
            ln = d.get('luu_nien', '').split(); d['can'], d['chi'] = (ln[0], ln[-1]) if ln else ('', '')
            ra.append(d)
    return ra


def ho_so(E, la_so=None, so_nam=11):
    """Hồ sơ dữ kiện của một lá số (dict). E = Engine đã tính; la_so = LaSo (tùy chọn: giới tính, tên…)."""
    tru, nhom = _tu_tru(E)
    bm = _ban_menh(E)
    tom, muc = _ket_luan_muc(E)
    nam_xem = E.raw(E.M.si['NHAP_LIEU'], 31, 3)
    nam_xem = int(nam_xem) if isinstance(nam_xem, float) else None
    tt = [(t['can'], t['chi']) for t in tru]
    dung, hy, ky = bm['hanh_dung'], bm['hanh_hy'], bm['hanh_ky']
    dv = _dai_van(E)
    for d in dv:
        d['quan_he'] = Q.danh_dau(Q.voi_tu_tru(tt, d['can'], d['chi'], 'đại vận'), tt, dung, hy, ky)
    ln = _luu_nien(E, nam_xem, so_nam) if nam_xem else []
    for y in ln:
        y['quan_he'] = Q.danh_dau(Q.voi_tu_tru(tt, y['can'], y['chi'], 'năm %d' % y['nam']), tt, dung, hy, ky)
        v = next((d for d in dv if d['ten'] == y.get('dai_van')), None)
        if v: y['quan_he_voi_dai_van'] = Q.giua_hai(y['can'], y['chi'], 'năm %d' % y['nam'], v['can'], v['chi'], 'đại vận')
    nam_sinh = (ln[0]['nam'] - ln[0]['tuoi_thuc']) if ln and isinstance(ln[0].get('tuoi_thuc'), int) else None
    return {
        'phien_ban_ho_so': PHIEN_BAN_HO_SO,
        'la_so': {'gioi_tinh': getattr(la_so, 'gioi_tinh', None) or _s(E.gia_tri('NHAP_LIEU', 'C22') or E.raw(E.M.si['NHAP_LIEU'], 10, 3)),
                  'nam_sinh': nam_sinh, 'nam_xem': nam_xem,
                  'tuoi_thuc': ln[0]['tuoi_thuc'] if ln else None, 'tuoi_mu': ln[0]['tuoi_mu'] if ln else None,
                  'tu_tru': ' · '.join('%s %s' % x for x in tt)},
        'tu_tru': tru, 'thap_than_nhom': nhom, 'ban_menh': bm,
        'quan_he_tu_tru': Q.danh_dau(Q.trong_tu_tru(tt), tt, dung, hy, ky),
        'dai_van': dv, 'luu_nien': ln, 'tom_tat': tom, 'ket_luan_muc': muc,
        # phương pháp khuyến nghị của workbook (DOI_CHIEU C33: 0 thống nhất, 1 PP1, 2 PP2) – dùng để lấy đúng cột kết luận
        'phuong_phap': int(E.gia_tri('DOI_CHIEU', 'C33') or 0),
    }
