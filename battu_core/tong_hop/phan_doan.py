# -*- coding: utf-8 -*-
"""Lớp phán đoán – gỡ mâu thuẫn (đề xuất hoàn thiện 15/11/2026, giai đoạn T3, mục A1–A6, B3, B4).

Đầu vào: hồ sơ dữ kiện (tong_hop.du_kien.ho_so) + quy tắc battu_core/data/quy_tac_tong_hop.json (chủ dự án duyệt).
Đầu ra: MỘT kết luận cho mỗi việc, kèm dòng lý do, để bản luận giải không còn hai kết luận trái nhau:
  dai_van   – thang 5 mức cho cả vận, 5 năm đầu (can) và 5 năm sau (chi)                       (A1)
  luu_nien  – một điểm năm duy nhất, trừ nặng khi xung chi mang dụng thần / cung phu thê, hạn nặng  (A2, A3)
  tho_yeu   – gộp nhiều tiêu chí; căn cứ ngược chiều thì không dùng chữ "yểu"                    (A4)
  tai_loc   – một thang                                                                        (A5)
  luc_than  – kết luận Thiệu Vĩ Hoa + nguồn khác nói ngược (có trọng số)                         (A6)
  vuong_suy – cờ sát ngưỡng + giả thuyết ngược + gợi ý kiểm chứng                                (B3)
  an_muc    – mục không hợp tuổi / hoàn cảnh                                                   (B4)
  doi_chieu – chỗ kết luận mới khác workbook, để người duyệt thấy
Không đổi giá trị ô nào. Chỉ thư viện chuẩn."""
import json, os, re
from battu_core.tong_hop import quan_he as Q

_QT = None
MUC = ['Rất xấu', 'Xấu', 'Trung bình', 'Tốt', 'Rất tốt']


def quy_tac():
    global _QT
    if _QT is None:
        p = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'quy_tac_tong_hop.json')
        _QT = json.load(open(p, encoding='utf-8'))
    return _QT


def _muc(diem, nguong):
    for ten in ('Rất tốt', 'Tốt', 'Trung bình', 'Xấu'):
        if diem >= nguong[ten]: return ten
    return 'Rất xấu'


# ---------------------------------------------------------------- vai trò hành
class VaiTro:
    """Điểm vai trò của từng hành với nhật chủ (quy_tac.vai_tro_hanh)."""

    def __init__(self, hs):
        bm = hs['ban_menh']; qt = quy_tac()['vai_tro_hanh']
        nc = hs['tu_tru'][2]['can']; self.N = Q.HANH_CAN[Q.CAN.index(nc)] if nc in Q.CAN else None
        # Kết luận khuyến nghị đứng ĐẦU chữ ("NHƯỢC (Theo PP1)\n⚠ VƯỢNG…" – phần sau là cảnh báo, không phải kết luận)
        m = re.match(r'\s*(VƯỢNG|NHƯỢC)', (bm.get('vuong_suy') or '').upper()) or re.search(r'(VƯỢNG|NHƯỢC)', (bm.get('pp1') or '').upper())
        self.vuong = (m.group(1) == 'VƯỢNG') if m else None
        self.chinh_cach = (bm.get('cach_cuc') or 'Chính cách').startswith('Chính')
        self.ten = {}
        for k in ('dung', 'hy', 'ky'):
            h = bm.get('hanh_' + k)
            if h in Q.HANH and Q.HANH.index(h) not in self.ten: self.ten[Q.HANH.index(h)] = k
        vs_qt = quy_tac()['vuong_suy']; pt = bm.get('phe_minh_pt')
        self.sat_nguong = isinstance(pt, (int, float)) and vs_qt['sat_nguong_tu'] <= pt <= vs_qt['sat_nguong_den']
        lech = qt['phe_gay_mat_can_bang_khi_sat_nguong'] if self.sat_nguong else qt['phe_gay_mat_can_bang']
        self.diem = {}
        for h in range(5):
            v = self.ten.get(h)
            if v: self.diem[h] = qt[v]; continue
            if self.N is None or self.vuong is None or not self.chinh_cach: self.diem[h] = qt['nhan']; continue
            phe_minh = h in (self.N, (self.N - 1) % 5)                       # tỉ kiếp, ấn
            if self.vuong: self.diem[h] = lech if phe_minh else qt['nhan_co_loi']
            else: self.diem[h] = qt['nhan_co_loi'] if phe_minh else lech

    def can(self, c): return self.diem.get(Q.HANH_CAN[Q.CAN.index(c)], 0.0) if c in Q.CAN else 0.0
    def chi(self, c): return self.diem.get(Q.HANH_CHI[Q.CHI.index(c)], 0.0) if c in Q.CHI else 0.0
    def nhan(self, h): return {'dung': 'dụng thần', 'hy': 'hỷ thần', 'ky': 'kỵ thần'}.get(self.ten.get(h), '')


def _dieu_chinh(qh, vt, qt, nhat_can):
    """Điều chỉnh điểm theo quan hệ của trụ ngoài với tứ trụ → (điểm can, điểm chi, [lý do])."""
    dc_can = dc_chi = 0.0; ly = []; xung = hinh = 0; hanh_hop = set()
    for q in qh:
        l = q['loai']; cham = q.get('cham') or []
        vai = cham[0].get('vai_tro') if cham else None
        if l == 'xung':
            if vai == 'dụng' and xung < qt['xung_toi_da_lan']:
                dc_chi += qt['xung_chi_dung']; xung += 1; ly.append('%s – chạm dụng thần (%+.1f)' % (q['ten'], qt['xung_chi_dung']))
            elif vai == 'kỵ':
                dc_chi += qt['xung_chi_ky']; ly.append('%s – xung mất kỵ thần (%+.1f)' % (q['ten'], qt['xung_chi_ky']))
            if q.get('cham_cung_phu_the'):
                dc_chi += qt['xung_cung_phu_the']; ly.append('%s – chạm cung phu thê (%+.1f)' % (q['ten'], qt['xung_cung_phu_the']))
        elif l == 'thien_khac_dia_xung':
            dc_chi += qt['thien_khac_dia_xung']; ly.append('%s (%+.1f)' % (q['ten'], qt['thien_khac_dia_xung']))
        elif l in ('ban_hop', 'tam_hop', 'tam_hoi') and q.get('hanh') in Q.HANH:
            h = Q.HANH.index(q['hanh'])
            if h in hanh_hop: continue
            hanh_hop.add(h); t = vt.ten.get(h)
            if t == 'dung': dc_chi += qt['hop_thanh_hanh_dung']; ly.append('%s – thành hành dụng thần (%+.1f)' % (q['ten'], qt['hop_thanh_hanh_dung']))
            elif t == 'ky': dc_chi += qt['hop_thanh_hanh_ky']; ly.append('%s – thành hành kỵ thần (%+.1f)' % (q['ten'], qt['hop_thanh_hanh_ky']))
        elif l == 'luc_hop' and q.get('cham_cung_phu_the') and 'hop_cung_phu_the' in qt:
            h = Q.HANH.index(q['hanh']) if q.get('hanh') in Q.HANH else None
            if vt.ten.get(h) != 'ky': dc_chi += qt['hop_cung_phu_the']; ly.append('%s – hợp cung phu thê (%+.1f)' % (q['ten'], qt['hop_cung_phu_the']))
        elif l == 'can_hop' and 'tai_hop_nhat_chu' in qt and q['vi_tri'][1] == 'can ngày':
            ngoai = q['chu'][0]
            if vt.N is not None and ngoai in Q.CAN and Q.HANH_CAN[Q.CAN.index(ngoai)] == (vt.N + 2) % 5:
                dc_can += qt['tai_hop_nhat_chu']; ly.append('%s – tài hợp nhật chủ (%+.1f)' % (q['ten'], qt['tai_hop_nhat_chu']))
        elif l in ('hinh', 'trung_chi'):
            if hinh * qt['hinh_moi_lan'] > qt['hinh_toi_da']:
                dc_chi += qt['hinh_moi_lan']; hinh += 1; ly.append('%s (%+.2f)' % (q['ten'], qt['hinh_moi_lan']))
        elif l == 'tam_hinh':
            dc_chi += qt['tam_hinh']; ly.append('%s (%+.1f)' % (q['ten'], qt['tam_hinh']))
    return dc_can, dc_chi, ly


def _vai_ly(vt, chu, la_can):
    d = vt.can(chu) if la_can else vt.chi(chu)
    h = Q.HANH_CAN[Q.CAN.index(chu)] if la_can else Q.HANH_CHI[Q.CHI.index(chu)]
    nhan = vt.nhan(h) or ('hành có lợi' if d > 0 else ('hành bất lợi' if d < 0 else 'nhàn thần'))
    return d, '%s %s (%s, %s) %+.1f' % ('can' if la_can else 'chi', chu, Q.HANH[h], nhan, d)


# ---------------------------------------------------------------- đại vận, lưu niên
def dai_van(hs, vt=None):
    vt = vt or VaiTro(hs); qt = quy_tac()['dai_van']; ra = []
    for d in hs['dai_van']:
        if d['can'] not in Q.CAN or d['chi'] not in Q.CHI: continue
        dcan, lcan = _vai_ly(vt, d['can'], True); dchi, lchi = _vai_ly(vt, d['chi'], False)
        a, b, ly = _dieu_chinh(d.get('quan_he', []), vt, qt, hs['tu_tru'][2]['can'])
        nua1, nua2 = dcan + a, dchi + b; tong = nua1 + nua2
        ra.append({'ten': d['ten'], 'nam_bat_dau': d.get('nam_bat_dau'), 'nam_ket_thuc': d.get('nam_ket_thuc'),
                   'tuoi_bat_dau': d.get('tuoi_bat_dau'), 'dang_hanh': d.get('dang_hanh'),
                   'diem': round(tong, 2), 'muc': _muc(tong, qt['nguong']),
                   'diem_5_nam_dau': round(nua1, 2), 'muc_5_nam_dau': _muc(nua1, qt['nguong_nua']),
                   'diem_5_nam_sau': round(nua2, 2), 'muc_5_nam_sau': _muc(nua2, qt['nguong_nua']),
                   'ly_do': [lcan, lchi] + ly, 'workbook': {'pp1': d.get('xu_the_pp1'), 'pp2': d.get('xu_the_pp2')}})
    return ra


def luu_nien(hs, dv=None, vt=None, nam=None):
    vt = vt or VaiTro(hs); qt = quy_tac()['luu_nien']; dv = dv if dv is not None else dai_van(hs, vt)
    ra = []
    for y in (nam if nam is not None else hs['luu_nien']):
        if y.get('can') not in Q.CAN or y.get('chi') not in Q.CHI: continue
        dcan, lcan = _vai_ly(vt, y['can'], True); dchi, lchi = _vai_ly(vt, y['chi'], False)
        a, b, ly = _dieu_chinh(y.get('quan_he', []), vt, qt, hs['tu_tru'][2]['can'])
        tong = dcan + dchi + a + b
        for q in y.get('quan_he_voi_dai_van', []):
            if q['loai'] == 'xung': tong += qt['xung_dai_van']; ly.append('%s – tuế vận xung (%+.1f)' % (q['ten'], qt['xung_dai_van']))
        v = next((x for x in dv if x['ten'] == y.get('dai_van')), None)
        if v and isinstance(v.get('nam_bat_dau'), int):
            nua = v['diem_5_nam_dau'] if y['nam'] - v['nam_bat_dau'] < 5 else v['diem_5_nam_sau']
            tong += qt['he_so_van'] * nua
            ly.append('đại vận %s, %s: %+.1f × %.1f' % (v['ten'], '5 năm đầu' if y['nam'] - v['nam_bat_dau'] < 5 else '5 năm sau', nua, qt['he_so_van']))
        han = y.get('so_dau_hieu_nang') if isinstance(y.get('so_dau_hieu_nang'), int) else 0
        if han:
            tong += qt['han_nang_moi_dau_hieu'] * han; ly.append('%d dấu hiệu hạn nặng của workbook (%+.1f)' % (han, qt['han_nang_moi_dau_hieu'] * han))
        muc = _muc(tong, qt['nguong'])
        if han and MUC.index(muc) > MUC.index(qt['tran_khi_han_nang']):
            muc = qt['tran_khi_han_nang']; ly.append('có hạn nặng → không xếp cao hơn "%s"' % muc)
        ra.append({'nam': y['nam'], 'can_chi': y.get('luu_nien'), 'tuoi_mu': y.get('tuoi_mu'), 'tuoi_thuc': y.get('tuoi_thuc'),
                   'diem': round(tong, 2), 'muc': muc, 'han_nang': han, 'ly_do': [lcan, lchi] + ly,
                   'workbook': {'pp1': y.get('tong_hop_pp1'), 'pp2': y.get('tong_hop_pp2')}})
    return ra


# ---------------------------------------------------------------- thọ yểu, tài lộc, lục thân
def _muc_kl(hs, muc, nhan):
    for k, v in hs['ket_luan_muc'].get(muc, {}).items():
        if k.startswith(nhan): return v if isinstance(v, str) else ''
    return ''


def tho_yeu(hs):
    qt = quy_tac()['tho_yeu']; bm = hs['ban_menh']
    goc = _muc_kl(hs, 'H. Sức khỏe', 'Thọ – yểu')
    tot = []
    if re.match(r'\s*VƯỢNG', (bm.get('vuong_suy') or '').upper()): tot.append('thân vượng')
    ts = (bm.get('than_sat_hieu_luc') or '').upper()
    if 'NGUYỆT ĐỨC' in ts: tot.append('nguyệt đức')
    if 'THIÊN ĐỨC' in ts: tot.append('thiên đức')
    if _muc_kl(hs, 'H. Sức khỏe', 'Sức khỏe chung').upper().startswith('THUẬN'): tot.append('sức khỏe chung thuận')
    if 'CÓ CỨU' in (bm.get('ket_luan') or '').upper(): tot.append('có cứu')
    if 'ít bệnh' in _muc_kl(hs, 'H. Sức khỏe', 'Sức khỏe (HG)'): tot.append('bình sinh ít bệnh')
    yeu = 'YỂU' in goc.upper()[:40]
    if yeu and len(tot) >= qt['so_dau_hieu_nguoc']:
        return {'muc': 'Căn cứ ngược chiều', 'goc_workbook': goc.split('.')[0],
                'dien_giai': 'Bảng P1 tr.173 nghiêng về yểu (chỉ so Thực thương với Ấn), nhưng có %d dấu hiệu ngược: %s. '
                             'Đọc là thể chất nền tốt, đời có giai đoạn thử thách sức khỏe – chú ý các năm hạn.' % (len(tot), ', '.join(tot)),
                'dau_hieu_tot': tot}
    return {'muc': 'Nghiêng về yểu' if yeu else ('Nghiêng về thọ' if 'THỌ' in goc.upper()[:40] else 'Trung bình'),
            'goc_workbook': goc.split('.')[0], 'dien_giai': '', 'dau_hieu_tot': tot}


def tai_loc(hs):
    kl = _muc_kl(hs, 'E. Tài lộc', 'KẾT LUẬN TÀI LỘC')
    m = re.search(r'MỨC PHÚ QUÝ TỔNG HỢP:\s*([A-ZÀ-Ỹ ]+?)\s*\(', kl)
    muc = m.group(1).strip().capitalize() if m else ''
    tp = {'Bậc giàu': _muc_kl(hs, 'E. Tài lộc', 'Bậc giàu').split('.')[0], 'Giàu – nghèo (bảng P1 tr.173)': _muc_kl(hs, 'E. Tài lộc', 'Giàu – nghèo').split('(')[0].strip(),
          'Tài vận': _muc_kl(hs, 'E. Tài lộc', 'Tài vận').split('(')[0].strip()}
    return {'muc': muc, 'thanh_phan': tp,
            'dien_giai': 'Một thang duy nhất là mức tổng hợp của workbook (%s); bậc giàu, bảng giàu – nghèo, tài vận là thành phần của thang này.' % (muc or '—')}


_LUC_THAN = [('Ông bà – tổ nghiệp', 'Ông bà', ()), ('Cha', 'Cha', ('Cha:',)), ('Mẹ', 'Mẹ', ('Mẹ:',)),
             ('Anh chị em', 'Anh chị em', ()), ('Con cái', 'Con cái', ('Con cái (UH)', 'Con cái – tử tức (HG)', '[Tham khảo ngoài sách] Con cái')),
             ('Vợ / chồng', 'Vợ / chồng', ('Hôn nhân (UH)', '[Tham khảo ngoài sách] Hôn nhân'))]
_CAP = {'TỐT': 1, 'KHÁ': 1, 'TRUNG BÌNH': 0, 'KÉM': -1, 'XẤU': -1}


def _nguon(nhan):
    if '(UH)' in nhan: return 'Uyên Hải Tử Bình'
    if '(HG)' in nhan: return 'Hóa giải toàn thư'
    if 'Tham khảo' in nhan: return 'Tham khảo'
    return 'Thiệu Vĩ Hoa'


def _cuc(t, qt):
    x = sum(w in t for w in qt['tu_xau']); y = sum(w in t for w in qt['tu_tot'])
    return -1 if x > y else (1 if y > x else 0)


def luc_than(hs):
    qt = quy_tac()['luc_than']; G = hs['ket_luan_muc'].get('G. Gia đạo', {}); ra = {}
    for nhan, ten, khac in _LUC_THAN:
        chinh = next((v for k, v in G.items() if k.startswith(nhan)), '') or ''
        cap = chinh.split(':')[0].strip().upper()
        d = {'muc': cap.capitalize(), 'chinh': chinh.split('Vì sao')[0].strip(), 'nguon_khac': []}
        for k, v in G.items():
            if not isinstance(v, str) or k.startswith(nhan): continue
            if not any(k.startswith(x) for x in khac) and not any(x in v for x in khac if x.endswith(':')): continue
            doan = v
            for x in khac:                                              # cắt đúng đoạn nói về người này ("Mẹ: …")
                if x.endswith(':') and x in v: doan = v[v.index(x):].split('[')[0]
            cuc = _cuc(doan, qt)
            if cuc and cuc != _CAP.get(cap, 0):
                ng = _nguon(k)
                d['nguon_khac'].append({'nguon': ng, 'trong_so': qt['trong_so_nguon'].get(ng, 1), 'noi': doan.strip()[:300],
                                        'chieu': 'tốt hơn' if cuc > _CAP.get(cap, 0) else 'xấu hơn'})
        if d['nguon_khac']:
            d['dien_giai'] = 'Giữ kết luận Thiệu Vĩ Hoa (%s). Nguồn có trọng số thấp hơn nói %s: %s.' % (
                d['muc'] or '—', ', '.join(sorted({x['chieu'] for x in d['nguon_khac']})), '; '.join(sorted({x['nguon'] for x in d['nguon_khac']})))
        ra[ten] = d
    return ra


# ---------------------------------------------------------------- vượng suy sát ngưỡng, lọc theo tuổi
def vuong_suy(hs, dv):
    qt = quy_tac()['vuong_suy']; bm = hs['ban_menh']; pt = bm.get('phe_minh_pt')
    sat = isinstance(pt, (int, float)) and qt['sat_nguong_tu'] <= pt <= qt['sat_nguong_den']
    d = {'ket_luan': bm.get('vuong_suy'), 'phe_minh_pt': pt, 'sat_nguong': sat}
    if sat:
        nx = hs['la_so'].get('nam_xem') or 0
        qua = [x for x in dv if isinstance(x.get('nam_ket_thuc'), int) and x['nam_ket_thuc'] < nx]
        tot = [x['ten'] for x in qua if x['muc'] in ('Tốt', 'Rất tốt')]; xau = [x['ten'] for x in qua if x['muc'] in ('Xấu', 'Rất xấu')]
        d['gia_thuyet_nguoc'] = 'thân nhược' if re.match(r'\s*VƯỢNG', (bm.get('vuong_suy') or '').upper()) else 'thân vượng'
        d['kiem_chung'] = ('Phe mình %.1f%% sát ngưỡng 40%%. Theo dụng thần %s, các vận đã qua %s phải thuận hơn hẳn %s. '
                           'Nếu thực tế ngược lại, nên xem lại theo giả thuyết %s.' % (
                               pt, bm.get('hanh_dung'), ', '.join(tot) or '—', ', '.join(xau) or '—', d['gia_thuyet_nguoc']))
    return d


def an_muc(hs, la_so=None):
    qt = quy_tac()['loc_theo_tuoi']; tuoi = hs['la_so'].get('tuoi_thuc'); ra = []
    hn = (getattr(la_so, 'tinh_trang_hon_nhan', '') or '') if la_so else ''
    so_con = getattr(la_so, 'so_con', None) if la_so else None
    if (isinstance(tuoi, int) and tuoi >= qt['an_thoi_diem_ket_hon_tu_tuoi']) or hn in ('Đã kết hôn', 'Ly hôn', 'Góa'):
        ra.append({'muc': 'Thời điểm kết hôn', 'ly_do': 'đã kết hôn' if hn == 'Đã kết hôn' else 'tuổi %s – đọc thành "năm thuận gia đạo"' % tuoi})
    if isinstance(tuoi, int) and tuoi >= qt['an_moc_hoc_tap_tu_tuoi']:
        ra.append({'muc': 'Mốc học tập', 'ly_do': 'tuổi %s' % tuoi})
    if isinstance(so_con, int) and so_con > 0:
        ra.append({'muc': 'Con đầu lòng', 'ly_do': 'đã có %d con' % so_con})
    return ra


def _chieu(muc): return {'Rất tốt': 1, 'Tốt': 1, 'Xấu': -1, 'Rất xấu': -1}.get(muc, 0)


def _chieu_wb(t):
    t = (t or '').lower().split(':')[-1]
    if 'tốt ít' in t or 'xấu nhiều' in t or 'xấu vừa' in t or t.strip() == 'xấu': return -1
    if 'tốt' in t: return 1
    return 0


def phan_doan(hs, la_so=None):
    """Toàn bộ kết luận đã gỡ mâu thuẫn cho một hồ sơ dữ kiện."""
    vt = VaiTro(hs); dv = dai_van(hs, vt); ln = luu_nien(hs, dv, vt)
    doi = []
    for ten, ds in (('Đại vận', dv), ('Năm', ln)):
        for x in ds:
            wb = x['workbook'].get('pp1') or ''
            if _chieu_wb(wb) * _chieu(x['muc']) < 0:               # chỉ ghi khi ĐỔI CHIỀU (tốt ↔ xấu)
                doi.append('%s %s: workbook "%s" → tổng hợp "%s"' % (ten, x.get('ten') or '%d %s' % (x['nam'], x['can_chi']), wb, x['muc']))
    ty = tho_yeu(hs)
    if ty['muc'] == 'Căn cứ ngược chiều': doi.append('Thọ – yểu: workbook "%s" → "căn cứ ngược chiều"' % ty['goc_workbook'])
    return {'phien_ban_quy_tac': quy_tac()['phien_ban'], 'dai_van': dv, 'luu_nien': ln, 'tho_yeu': ty, 'tai_loc': tai_loc(hs),
            'luc_than': luc_than(hs), 'vuong_suy': vuong_suy(hs, dv), 'an_muc': an_muc(hs, la_so), 'doi_chieu': doi}
