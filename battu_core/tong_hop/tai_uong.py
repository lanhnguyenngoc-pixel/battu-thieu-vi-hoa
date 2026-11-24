# -*- coding: utf-8 -*-
"""Bốn kiểu lá số "dễ gặp tai ương" và bảng sức khỏe theo năm (kế hoạch KE_HOACH_4_KIEU_LA_SO_NGUY_HIEM, 24/11/2026).

NHÃN BẮT BUỘC (người dùng chốt 24/11/2026): mọi chỗ hiện đều kèm NHAN – đây là cách luận tham khảo thêm,
không phải quy tắc của sách Thiệu Vĩ Hoa. Chỉ ĐỌC hồ sơ dữ kiện (du_kien.ho_so); không đổi bất kỳ giá trị ô,
mức tài lộc, sự nghiệp, hôn nhân hay mức vận/năm nào. Chỉ thư viện chuẩn.

    phan_tich(hs)            → dict 4 kiểu của lá số + cờ theo đại vận + cờ theo năm
    bang_suc_khoe(hs, n=10)  → bảng sức khỏe n năm liên tiếp từ năm đặt xem
"""
import re
from battu_core.tong_hop import quan_he as Q

NHAN = 'cách luận tham khảo thêm thuộc dễ gặp tai ương nhất.'
CAN, CHI, HANH, HC = Q.CAN, Q.CHI, Q.HANH, Q.HANH_CAN

# ---- ngưỡng (xem KET_QUA_4_KIEU: chọn sau khi đo trên 73 lá sách + lá ngẫu nhiên; người dùng có thể đổi)
NGUONG = {
    'nhan_nang_tu': 2,          # tổng Dương Nhẫn (mệnh + vận/năm) từ ngưỡng này trở lên là "quá nặng"
    'lech_vuot_pct': 55.0,      # một hành chiếm từ ngưỡng % tổng độ trở lên
    'lech_thieu_pct': 5.0,      # hành "gần như không có" khi dưới ngưỡng %
    'lech_thieu_so': 2,         # số hành gần như không có tối thiểu
    'van_tuoi_toi_da': 80,      # kiểu 4: chỉ xét đại vận bắt đầu trước tuổi này
}

# Chi Kình dương (Dương Nhẫn) của từng can – đúng bảng CSDL B1–B6 của sổ tay
KINH_DUONG = {'Giáp': 'Mão', 'Ất': 'Dần', 'Bính': 'Ngọ', 'Đinh': 'Tị', 'Mậu': 'Ngọ', 'Kỷ': 'Tị',
              'Canh': 'Dậu', 'Tân': 'Thân', 'Nhâm': 'Tí', 'Quý': 'Hợi'}

TS_TEN = ['Trường sinh', 'Mộc dục', 'Quan đới', 'Lâm quan', 'Đế vượng', 'Suy', 'Bệnh', 'Tử', 'Mộ', 'Tuyệt', 'Thai', 'Dưỡng']
TS_KHOI = {'Giáp': 'Hợi', 'Bính': 'Dần', 'Mậu': 'Dần', 'Canh': 'Tị', 'Nhâm': 'Thân',
           'Ất': 'Ngọ', 'Đinh': 'Dậu', 'Kỷ': 'Dậu', 'Tân': 'Tí', 'Quý': 'Mão'}   # vòng chu kỳ Hỏa như phần còn lại của công cụ
XAU_TS = ('Tử', 'Mộ', 'Tuyệt')


def trang_thai(can, chi):
    """Trạng thái vòng trường sinh của can tại chi (dương thuận, âm nghịch)."""
    if can not in CAN or chi not in CHI: return ''
    i, s = CHI.index(chi), CHI.index(TS_KHOI[can])
    k = (i - s) % 12 if CAN.index(can) % 2 == 0 else (s - i) % 12
    return TS_TEN[k]


def thap_than(nc, c):
    """Thập thần của can c so với nhật chủ nc."""
    if nc not in CAN or c not in CAN: return ''
    hn, hc = HC[CAN.index(nc)], HC[CAN.index(c)]
    cung = (CAN.index(nc) % 2) == (CAN.index(c) % 2)
    if hn == hc: return 'Tỉ kiên' if cung else 'Kiếp tài'
    if (hc + 1) % 5 == hn: return 'Thiên ấn' if cung else 'Chính ấn'
    if (hn + 1) % 5 == hc: return 'Thực thần' if cung else 'Thương quan'
    if (hn + 2) % 5 == hc: return 'Thiên tài' if cung else 'Chính tài'
    return 'Thất sát' if cung else 'Chính quan'


def _tang(t):
    """Các can tàng trong một trụ: [(can, thập thần)] từ chuỗi 'Giáp – Thực thần'."""
    ra = []
    for k in ('tang_ban_khi', 'tang_trung_khi', 'tang_du_khi'):
        m = re.match(r'\s*(Giáp|Ất|Bính|Đinh|Mậu|Kỷ|Canh|Tân|Nhâm|Quý)\s*[–-]\s*(.+)', t.get(k) or '')
        if m: ra.append((m.group(1), m.group(2).strip()))
    return ra


def _chi_bi_xung_hinh(hs, chi):
    for q in hs.get('quan_he_tu_tru') or []:
        if q.get('loai') in ('xung', 'hinh') and chi in (q.get('chu') or []): return True
    return False


# ------------------------------------------------------------------ kiểu 1
def kieu1(hs):
    """Kiêu thần đoạt thực – mất "sao Thọ": Thiên ấn LỘ trên can áp chế Thực thần (lộ hoặc tàng) ngay trong mệnh."""
    tt = hs['tu_tru']; nc = tt[2]['can']; thang = tt[1]['chi']
    an, thuc = [], []                                         # (vị trí, can, lộ?)
    for t in tt:
        if t['tru'] != 'ngày':
            ptt = thap_than(nc, t['can'])
            if ptt == 'Thiên ấn': an.append(('can ' + t['tru'], t['can'], True))
            if ptt == 'Thực thần': thuc.append(('can ' + t['tru'], t['can'], True))
        for c, ten in _tang(t):
            if ten == 'Thiên ấn': an.append(('tàng ' + t['tru'], c, False))
            if ten == 'Thực thần': thuc.append(('tàng ' + t['tru'], c, False))
    r = {'ma': 1, 'ten': 'Kiêu thần đoạt thực – mất sao Thọ (Thiên ấn áp chế Thực thần)', 'co': False, 'muc': '', 'ly_do': []}
    lo_an = any(l for _, _, l in an); lo_thuc = any(l for _, _, l in thuc)
    if not (lo_an and thuc): return r
    r['co'] = True
    r['ly_do'].append('Thiên ấn lộ ở %s; Thực thần ở %s' % (', '.join(sorted({v for v, _, l in an if l})), ', '.join(sorted({v for v, _, _ in thuc}))))
    yeu = []
    cs = {c for _, c, _ in thuc}
    if all(trang_thai(c, thang) in XAU_TS for c in cs):
        yeu.append('Thực thần (%s) ở %s tại chi tháng %s' % ('/'.join(sorted(cs)), trang_thai(sorted(cs)[0], thang).lower(), thang))
    chi_thuc = {t['chi'] for t in tt if any(c in cs for c, _ in _tang(t))}
    if any(_chi_bi_xung_hinh(hs, c) for c in chi_thuc): yeu.append('chi chứa Thực thần bị xung hoặc hình')
    nang = lo_thuc or bool(yeu)
    if lo_thuc: r['ly_do'].append('Thiên ấn và Thực thần cùng lộ trên can')
    r['ly_do'] += ['Thực thần yếu: ' + y for y in yeu]
    r['muc'] = 'NẶNG' if nang else 'CÓ'
    return r


# ------------------------------------------------------------------ kiểu 2
def so_nhan_menh(hs):
    nc = hs['tu_tru'][2]['can']; kd = KINH_DUONG.get(nc)
    return kd, [t['tru'] for t in hs['tu_tru'] if t['chi'] == kd]


def kieu2(hs):
    """Dương Nhẫn (Kình dương) quá nặng: nhiều Dương Nhẫn (kể cả do đại vận/lưu niên đem tới)."""
    kd, vi = so_nhan_menh(hs)
    r = {'ma': 2, 'ten': 'Dương Nhẫn quá nặng', 'co': False, 'muc': '', 'ly_do': [], 'so_menh': len(vi)}
    if not kd: return r
    dv = next((d for d in hs['dai_van'] if d.get('dang_hanh')), None)
    n_dv = 1 if dv and dv['chi'] == kd else 0
    n_ln = [y['nam'] for y in hs['luu_nien'] if y.get('chi') == kd]
    tong = len(vi) + n_dv
    r['kinh_duong'] = kd
    if vi: r['ly_do'].append('chi Kình dương của nhật chủ (%s) có ở trụ %s' % (kd, ', '.join(vi)))
    if n_dv: r['ly_do'].append('đại vận đang đi (%s) đem thêm một chi %s' % (dv['ten'], kd))
    if n_ln: r['ly_do'].append('lưu niên đem chi %s: năm %s' % (kd, ', '.join(map(str, n_ln[:5]))))
    r['tong'] = tong
    if tong >= NGUONG['nhan_nang_tu']:
        r['co'] = True; r['muc'] = 'NẶNG'
        kem = _kem_nhan(hs)
        r['ly_do'] += kem
    elif vi: r['muc'] = 'bình thường (một Dương Nhẫn: nóng tính, quyết liệt)'
    return r


def _kem_nhan(hs):
    nc = hs['tu_tru'][2]['can']; ra = []
    sat = [t['tru'] for t in hs['tu_tru'] if t['tru'] != 'ngày' and thap_than(nc, t['can']) == 'Thất sát']
    if sat: ra.append('đi cùng Thất sát lộ ở can %s' % ', '.join(sat))
    if any(q.get('loai') == 'thien_khac_dia_xung' and 'ngày' in ' '.join(q.get('vi_tri') or []) for q in hs.get('quan_he_tu_tru') or []):
        ra.append('thiên khắc địa xung vào trụ Ngày')
    return ra


# ------------------------------------------------------------------ kiểu 3
def phan_tram(hs):
    do = hs['ban_menh'].get('do_theo_hanh') or {}
    tong = sum(do.values()) or 0.0
    return {h: (100.0 * do.get(h, 0.0) / tong if tong else 0.0) for h in HANH}


TANG_PHU = {'Hỏa': 'tim mạch, huyết áp', 'Kim': 'tai nạn, thương tích, hô hấp', 'Thủy': 'sức đề kháng, thận – bàng quang',
            'Mộc': 'tinh thần, gan – mật', 'Thổ': 'tiêu hóa, lá lách – dạ dày'}


def kieu3(hs):
    """Ngũ hành quá lệch: một hành áp đảo, hành khác gần như không có; thân quá yếu hoặc quá vượng."""
    pc = phan_tram(hs); bm = hs['ban_menh']
    r = {'ma': 3, 'ten': 'Ngũ hành quá lệch (thân quá yếu hoặc quá vượng)', 'co': False, 'muc': '', 'ly_do': [], 'pct': {h: round(v, 1) for h, v in pc.items()}}
    if not any(pc.values()): return r
    h_max = max(pc, key=pc.get); thieu = [h for h in HANH if pc[h] < NGUONG['lech_thieu_pct']]
    if pc[h_max] >= NGUONG['lech_vuot_pct'] and len(thieu) >= NGUONG['lech_thieu_so']:
        r['co'] = True; r['muc'] = 'NẶNG'
        r['ly_do'].append('%s chiếm %.0f%%; gần như không có: %s' % (h_max, pc[h_max], ', '.join(thieu)))
        vs = (bm.get('vuong_suy') or '').upper()
        if vs.startswith('NHƯỢC'): r['ly_do'].append('thân nhược (không rễ, không được sinh phù đủ)')
        elif vs.startswith('VƯỢNG'): r['ly_do'].append('thân vượng quá mức')
        r['tang_phu'] = '%s lệch cực đoan: chú ý %s' % (h_max, TANG_PHU[h_max]) + ('; ' + '; '.join('%s yếu: %s' % (h, TANG_PHU[h]) for h in thieu[:2]))
    return r


# ------------------------------------------------------------------ kiểu 4
def cac_van_tu_mo_tuyet(hs):
    """Các đại vận đẩy nhật chủ (hoặc can vận ở chi tháng, hoặc Thực thần) vào Tử, Mộ, Tuyệt – cờ giai đoạn.
    'nang' = trùng thiên khắc địa xung (bất kỳ trụ); 'xung_ngay' = trùng thiên khắc địa xung trụ Ngày;
    'chu_yeu' = nhật chủ hoặc Thực thần (không tính riêng can vận ở chi tháng)."""
    nc = hs['tu_tru'][2]['can']; thang = hs['tu_tru'][1]['chi']
    thuc = [c for c in CAN if thap_than(nc, c) == 'Thực thần']
    ra = []
    for d in hs['dai_van']:
        ly = []; chu_yeu = False
        if any(x in (d.get('truong_sinh') or '') for x in XAU_TS):
            ly.append('nhật chủ ở chi vận %s: %s' % (d['chi'], d['truong_sinh'].strip())); chu_yeu = True
        ts_can = trang_thai(d['can'], thang)
        if ts_can in XAU_TS: ly.append('can vận %s ở chi tháng %s: %s' % (d['can'], thang, ts_can))
        if thuc and all(trang_thai(c, d['chi']) in XAU_TS for c in thuc):
            ly.append('Thực thần (sao Thọ) ở chi vận: %s' % trang_thai(thuc[0], d['chi']).lower()); chu_yeu = True
        if not ly: continue
        nang = [q['ten'] for q in d.get('quan_he') or [] if q.get('loai') == 'thien_khac_dia_xung']
        ra.append({'van': d['ten'], 'tu': d['nam_bat_dau'], 'den': d['nam_ket_thuc'], 'tuoi': d.get('tuoi_bat_dau') or 0, 'ly_do': ly,
                   'nang': bool(nang), 'xung_ngay': any('trụ ngày' in x for x in nang), 'chu_yeu': chu_yeu, 'them': nang})
    return ra


def kieu4(hs):
    """Đại vận đi vào Tử – Mộ – Tuyệt: kiểu của lá khi có vận (bắt đầu trước 80 tuổi) đẩy nhật chủ/Thực thần vào
    Tử–Mộ–Tuyệt mà còn trùng thiên khắc địa xung trụ Ngày. Các vận Tử–Mộ–Tuyệt khác vẫn hiện như cờ giai đoạn."""
    cv = cac_van_tu_mo_tuyet(hs)
    nang = [v for v in cv if v['chu_yeu'] and v['xung_ngay'] and v['tuoi'] <= NGUONG['van_tuoi_toi_da']]
    r = {'ma': 4, 'ten': 'Đại vận đi vào Tử – Mộ – Tuyệt', 'co': bool(nang), 'muc': 'NẶNG' if nang else '', 'ly_do': [], 'cac_van': cv, 'van_nang': nang}
    for v in nang:
        r['ly_do'].append('%s (%s–%s): %s; trùng %s' % (v['van'], v['tu'], v['den'], '; '.join(v['ly_do']), ', '.join(v['them'])))
    return r


# ------------------------------------------------------------------ gộp
def kieu_theo_nam(hs, kd=None):
    """Cờ riêng của từng năm (trong hs['luu_nien']): kiêu đoạt thực do lưu niên/đại vận, Dương Nhẫn."""
    nc = hs['tu_tru'][2]['can']; kd = kd or KINH_DUONG.get(nc)
    tt = hs['tu_tru']
    co_thuc = any(thap_than(nc, t['can']) == 'Thực thần' for t in tt if t['tru'] != 'ngày') or \
        any(ten == 'Thực thần' for t in tt for _, ten in _tang(t))
    nhan_menh = sum(1 for t in tt if t['chi'] == kd)
    ra = {}
    for y in hs['luu_nien']:
        d = []
        if co_thuc and (y.get('thap_than_can') == 'Thiên ấn'):
            d.append('Kiêu đoạt thực: can năm là Thiên ấn gặp Thực thần trong mệnh')
        v = next((x for x in hs['dai_van'] if x['ten'] == y.get('dai_van')), None)
        if kd and y.get('chi') == kd:
            tong = nhan_menh + 1 + (1 if v and v['chi'] == kd else 0)
            d.append('Dương Nhẫn: chi năm là Kình dương (%s), tổng %d chữ%s' % (kd, tong, ' – quá nặng' if tong >= NGUONG['nhan_nang_tu'] else ''))
        ra[y['nam']] = d
    return ra


def co_dai_van(hs):
    """Cờ theo từng đại vận: {tên vận: [câu]} – Kiêu đoạt thực do vận Thiên ấn, Dương Nhẫn do chi vận, Tử/Mộ/Tuyệt."""
    nc = hs['tu_tru'][2]['can']; kd = KINH_DUONG.get(nc); tt = hs['tu_tru']
    co_thuc = any(thap_than(nc, t['can']) == 'Thực thần' for t in tt if t['tru'] != 'ngày') or \
        any(ten == 'Thực thần' for t in tt for _, ten in _tang(t))
    nhan_menh = sum(1 for t in tt if t['chi'] == kd)
    mo = {v['van']: v for v in cac_van_tu_mo_tuyet(hs)}
    ra = {}
    for d in hs['dai_van']:
        c = []
        if co_thuc and d.get('thap_than_can') == 'Thiên ấn':
            c.append('Kiêu đoạt thực: can vận là Thiên ấn gặp Thực thần trong mệnh')
        if kd and d['chi'] == kd:
            tong = nhan_menh + 1
            c.append('Dương Nhẫn: chi vận là Kình dương (%s), tổng %d chữ%s' % (kd, tong, ' – quá nặng' if tong >= NGUONG['nhan_nang_tu'] else ''))
        v = mo.get(d['ten'])
        if v: c.append(('' if v['chu_yeu'] else 'Giai đoạn ') + 'Tử/Mộ/Tuyệt: ' + '; '.join(v['ly_do']) + ('; trùng ' + ', '.join(v['them']) if v['them'] else ''))
        ra[d['ten']] = c
    return ra


def phan_tich(hs):
    k = [kieu1(hs), kieu2(hs), kieu3(hs), kieu4(hs)]
    return {'nhan': NHAN, 'kieu': k, 'co_it_nhat_mot': any(x['co'] for x in k), 'theo_nam': kieu_theo_nam(hs), 'theo_dai_van': co_dai_van(hs)}


def cau_tong_hop(hs):
    """Một câu tổng hợp cho luận giải chung (rỗng khi không kiểu nào dính)."""
    p = phan_tich(hs); co = [x for x in p['kieu'] if x['co']]
    if not co: return ''
    ten = '; '.join('%s (%s)' % (x['ten'].split(' (')[0], x['muc'].lower()) for x in co)
    return '**Lá số có %d trong 4 kiểu dễ gặp tai ương: %s** – %s' % (len(co), ten, NHAN)


# ------------------------------------------------------------------ bảng sức khỏe 10 năm
def _muc_nam(y, pp):
    return (y.get('tong_hop_pp%d' % pp) or y.get('tong_hop_pp1') or '').strip()


def bang_suc_khoe(hs, n=10):
    """n năm liên tiếp từ năm đặt xem (năm đặt xem là năm đầu). Mỗi dòng: năm, tuổi mụ, mức năm, tạng phủ, dấu hiệu."""
    pp = 2 if hs.get('phuong_phap') == 2 else 1
    theo_nam = kieu_theo_nam(hs)
    dong = []
    for y in hs['luu_nien'][:n]:
        dau = [s.strip(' •') for s in re.split(r'\s*•\s*', y.get('dau_hieu') or '') if s.strip(' •')]
        tp = (y.get('tang_phu') or '').strip().lstrip('• ').strip()
        m = re.match(r'(NẶNG|ĐỀ PHÒNG|CHÚ Ý)\s*[–-]\s*(.+?):', tp)
        dong.append({'nam': y['nam'], 'can_chi': y.get('luu_nien') or '', 'tuoi_mu': y.get('tuoi_mu'), 'muc': _muc_nam(y, pp),
                     'so_dau_hieu_nang': y.get('so_dau_hieu_nang') or 0, 'tang_phu': tp,
                     'tang_phu_muc': m.group(1) if m else '', 'tang_phu_ten': m.group(2).strip() if m else '',
                     'dau_hieu': dau, 'kieu': theo_nam.get(y['nam'], []), 'dai_van': y.get('dai_van') or ''})
    return dong


def nam_dang_chu_y(dong, k=3):
    """Các năm đáng chú ý nhất: nhiều dấu hiệu hạn nặng trước, rồi tạng phủ NẶNG, rồi cờ 4 kiểu."""
    def diem(d): return (d['so_dau_hieu_nang'], 1 if d['tang_phu_muc'] == 'NẶNG' else 0, len(d['kieu']))
    xs = [d for d in dong if diem(d) != (0, 0, 0)]
    return sorted(xs, key=lambda d: (-diem(d)[0], -diem(d)[1], -diem(d)[2], d['nam']))[:k]


# ------------------------------------------------------------------ khối bài đọc (màn hình, PDF, Word)
def _gon(t, n=150):
    """Rút gọn tại ranh giới mệnh đề (không dùng dấu chấm lửng, tránh bị coi là chữ cắt)."""
    t = re.sub(r'\s+', ' ', t or '').strip()
    if len(t) <= n: return t
    c = t[:n]; k = max(c.rfind('; '), c.rfind(', '), c.rfind(' – '))
    c = c[:k] if k > n // 3 else c.rsplit(' ', 1)[0]
    return c.rstrip(' ;,–:(')


def _dong_nam(d):
    ds = [re.sub(r'^\[[^\]]*\]\s*', '', x) for x in d['dau_hieu']]
    ds = [x.rstrip('.;') for x in ds if not x.startswith('…')]
    o = []
    if d['so_dau_hieu_nang']: o.append('**%d dấu hiệu hạn nặng**' % d['so_dau_hieu_nang'])
    o += [_gon(x) for x in ds[:2]]
    o += ['**%s**' % _gon(x, 110) for x in d['kieu']]
    return '; '.join(o) if o else '—'


def nam_ngan(hs, dong=None):
    """Chỉ liệt kê tên các năm đáng chú ý (cho luận giải tổng hợp, tiết kiệm chỗ)."""
    xs = nam_dang_chu_y(dong or bang_suc_khoe(hs), 3)
    return ', '.join('%d %s' % (d['nam'], d['can_chi']) for d in xs)


def khoi_nam_chu_y(hs, dong=None):
    """Hai ba năm đáng chú ý: một đoạn văn ngắn cho luận giải tổng hợp (rỗng khi không có năm nào)."""
    dong = dong or bang_suc_khoe(hs)
    xs = nam_dang_chu_y(dong, 3)[:3]
    if not xs: return ''
    return '; '.join('**%d %s** (%s)' % (d['nam'], d['can_chi'], _gon('; '.join(
        ([('%d dấu hiệu hạn nặng' % d['so_dau_hieu_nang'])] if d['so_dau_hieu_nang'] else []) +
        (['tạng phủ %s – %s' % (d['tang_phu_muc'].lower(), d['tang_phu_ten'])] if d['tang_phu_muc'] == 'NẶNG' else []) +
        [x.split(':')[0] for x in d['kieu']]), 70)) for d in xs)


def khoi_bai_doc(hs, n=10):
    """Khối bài đọc đầy đủ cho trang "Sức khỏe 10 năm" và phần xuất: 4 kiểu + bảng n năm + đại vận."""
    p = phan_tich(hs); dong = bang_suc_khoe(hs, n)
    la = hs['la_so']; nam0 = dong[0]['nam'] if dong else la.get('nam_xem')
    B = [('h2', 'Sức khỏe %d năm (từ năm đặt xem %s) và bốn kiểu lá số dễ gặp tai ương' % (len(dong), nam0)),
         ('nho', 'Chỉ là xu hướng mệnh lý để tham khảo, không phải chẩn đoán y khoa; không thay cho khám và tư vấn của bác sĩ. '
                 'Phần "bốn kiểu" là %s Mức từng năm lấy nguyên từ lưu niên của công cụ, bảng chỉ trình bày lại.' % NHAN)]
    B.append(('p', '**Bốn kiểu lá số – %s**' % NHAN))
    ul = []
    for k in p['kieu']:
        if k['co']:
            ul.append('**%s – có (%s).** %s.' % (k['ten'], k['muc'].lower(), '; '.join(k['ly_do'])))
        else:
            ghi = {1: 'Thiên ấn không lộ áp chế Thực thần trong mệnh', 2: 'số Dương Nhẫn (mệnh + vận hiện tại) dưới %d' % NGUONG['nhan_nang_tu'],
                   3: 'không có hành nào áp đảo kèm hành gần như không có', 4: 'không có đại vận nào vừa vào Tử – Mộ – Tuyệt vừa thiên khắc địa xung trụ Ngày'}[k['ma']]
            ul.append('%s – không dính (%s).' % (k['ten'], ghi))
    B.append(('ul', ul))
    if dong:
        rows = []
        for d in dong:
            tp = ('%s – %s' % (d['tang_phu_muc'].capitalize(), d['tang_phu_ten'])) if d['tang_phu_ten'] else '—'
            if d['tang_phu_muc'] == 'NẶNG': tp = '**%s**' % tp
            rows.append(['%d %s (%s)' % (d['nam'], d['can_chi'], d['tuoi_mu']), '**%s**' % (d['muc'] or '—'), tp, _dong_nam(d)])
        B.append(('bang', ['Năm (tuổi mụ)', 'Mức năm', 'Tạng phủ cần giữ', 'Dấu hiệu chính'], rows))
        cy = khoi_nam_chu_y(hs, dong)
        if cy: B.append(('p', '**Các năm đáng chú ý nhất:** %s.' % cy))
    th = p['theo_dai_van']; dv = hs['dai_van']
    dang = next((d for d in dv if d.get('dang_hanh')), None)
    if dang:
        B.append(('p', '**Đại vận đang đi: %s (%s–%s).** %s' % (dang['ten'], dang['nam_bat_dau'], dang['nam_ket_thuc'],
                                                              ('; '.join(th.get(dang['ten']) or []) or 'không có cờ riêng của bốn kiểu.'))))
    rows = [['%s (%s–%s)' % (d['ten'], d['nam_bat_dau'], d['nam_ket_thuc']), '**%s**' % (d.get('ket_luan') or '—').capitalize(), '; '.join(th.get(d['ten']) or []) or '—']
            for d in dv if th.get(d['ten']) and d is not dang]
    if rows: B.append(('bang', ['Đại vận có cờ', 'Mức vận', 'Cờ của bốn kiểu'], rows))
    B.append(('nho', 'Cách đọc: "hạn nặng" là dấu hiệu theo sách (P2 tr.282–289, 566), cần xét có cứu hay không; mức năm chỉ là một phần của vận 10 năm. '
                     'Các cờ Kiêu đoạt thực, Dương Nhẫn và Tử – Mộ – Tuyệt là ' + NHAN + ' Người dưới 18 tuổi đọc như người lớn, đúng theo lá số.'))
    return B
