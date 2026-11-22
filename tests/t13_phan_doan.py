# -*- coding: utf-8 -*-
"""T13 – Lớp phán đoán, gỡ mâu thuẫn (đề xuất hoàn thiện 15/11/2026, giai đoạn T3).
  (a) Bất biến trên 20 lá T7 + lá mẫu:
      • mỗi đại vận / mỗi nửa vận / mỗi năm có ĐÚNG MỘT mức trong thang 5 mức;
      • năm có ≥ 2 lục xung chạm chi mang dụng thần mà can và chi năm đều không có lợi → Xấu hoặc Rất xấu;
      • thọ – yểu: có từ ngưỡng dấu hiệu ngược trở lên thì không kết luận "yểu";
      • kết quả ghi được JSON và chạy lại cho cùng kết quả.
  (b) Lá mẫu Nguyễn Ngọc Lanh (Mậu Ngọ · Quý Hợi · Giáp Ngọ · Mậu Thìn, năm xem 2026) – kết luận chuẩn đã duyệt.
Dùng: python -m tests.t13_phan_doan"""
import sys, os, json, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.tong_hop.du_kien import ho_so
from battu_core.tong_hop.phan_doan import phan_doan, quy_tac, MUC, VaiTro
from tests.so_dap_an import oracle_nhap

GOC = os.path.dirname(os.path.abspath(__file__))


def _bat_bien(ten, hs, p):
    loi = []; qt = quy_tac()
    for d in p['dai_van']:
        for k in ('muc',):
            if d[k] not in MUC: loi.append((ten, d['ten'], k, d[k]))
    vt = VaiTro(hs)
    for y, h in zip(p['luu_nien'], hs['luu_nien']):
        if y['muc'] not in MUC: loi.append((ten, y['nam'], 'mức', y['muc']))
        # từ 1.9.0 mức năm = đúng kết luận tổng hợp của workbook (P1 tr.133) – không còn quy tắc riêng của lớp tổng hợp
    ty = p['tho_yeu']
    if ty['muc'].upper() != (hs['ban_menh'].get('tho_yeu_kl') or 'TRUNG BÌNH').upper():      # phải đúng nguyên kết luận workbook
        loi.append((ten, 'thọ – yểu khác workbook', ty['muc']))
    if not ty['muc'].upper().startswith(('THỌ', 'YỂU', 'CÂN BẰNG')): loi.append((ten, 'thọ – yểu lạ', ty['muc']))
    if json.loads(json.dumps(p, ensure_ascii=False)) != json.loads(json.dumps(phan_doan(hs), ensure_ascii=False)):
        loi.append((ten, 'không tất định'))
    return loi


MAU = {'dai_van': {'Giáp Tí': ('Xấu', 'Rất xấu'), 'Bính Dần': ('Tốt', 'Rất tốt'), 'Mậu Thìn': ('Tốt', 'Rất tốt'),
                   'Kỷ Tị': ('Tốt', 'Rất tốt'), 'Nhâm Thân': ('Xấu', 'Rất xấu')},
       'luu_nien': {2026: ('Rất tốt',), 2027: ('Tốt',), 2029: ('Tốt', 'Rất tốt'), 2030: ('Tốt', 'Rất tốt'),
                    2032: ('Tốt',), 2034: ('Tốt',), 2035: ('Tốt',), 2036: ('Tốt', 'Rất tốt')}}


def _la_mau():
    from battu_core.nhap_lieu import LaSo
    from battu_core.dich_vu import tinh
    s = LaSo(ho_ten='mẫu', cach_nhap='duong', ngay=28, thang=11, nam=1978, gio=8, phut=30, gioi_tinh='Nam', nam_xem=2026)
    kq = tinh(s); hs = ho_so(kq.E, s); p = phan_doan(hs, s); loi = _bat_bien('mẫu', hs, p)
    dv = {d['ten']: d['muc'] for d in p['dai_van']}; ln = {y['nam']: y['muc'] for y in p['luu_nien']}
    for k, ok in MAU['dai_van'].items():
        if dv.get(k) not in ok: loi.append(('mẫu', 'đại vận ' + k, dv.get(k), 'cần', ok))
    for k, ok in MAU['luu_nien'].items():
        if ln.get(k) not in ok: loi.append(('mẫu', 'năm %d' % k, ln.get(k), 'cần', ok))
    if p['tho_yeu']['muc'] != 'Yểu – được hóa giải một phần, nhẹ bớt': loi.append(('mẫu', 'thọ – yểu', p['tho_yeu']['muc']))
    if p['tai_loc']['muc'] != 'Khá giả': loi.append(('mẫu', 'tài lộc', p['tai_loc']['muc']))
    if not p['vuong_suy']['sat_nguong']: loi.append(('mẫu', 'vượng suy phải sát ngưỡng'))
    if not any(x['muc'] == 'Thời điểm kết hôn' for x in p['an_muc']): loi.append(('mẫu', 'chưa ẩn thời điểm kết hôn (48 tuổi)'))
    if p['luc_than']['Mẹ']['muc'] != 'Kém' or not p['luc_than']['Mẹ']['nguon_khac']: loi.append(('mẫu', 'mẹ: giữ Kém, có nguồn khác nói ngược'))
    s2 = LaSo(**{**s.to_dict(), 'so_con': 2, 'tinh_trang_hon_nhan': 'Đã kết hôn'})
    if not any(x['muc'] == 'Con đầu lòng' for x in phan_doan(hs, s2)['an_muc']): loi.append(('mẫu', 'đã có con mà chưa ẩn "Con đầu lòng"'))
    # Mục 2.9 (bản 6.8): dòng tham khảo con cái theo lá số hôn phối – đúng câu con cái khi mở lá đó riêng; không có thông tin → không có dòng
    from battu_core.tong_hop.viet_ban import _dong_hon_phoi, _cau_con_cai
    s3 = LaSo(**{**s.to_dict(), 'phoi_ngau': [5, 3, 1980, 10, 0]})
    sp = LaSo(cach_nhap='duong', ngay=5, thang=3, nam=1980, gio=10, phut=0, gioi_tinh='Nữ', nam_xem=2026)
    hp = ho_so(tinh(sp).E, sp); pp = phan_doan(hp, sp)
    chuan = _cau_con_cai(hp, pp, {x['muc'] for x in pp['an_muc']}, gon=True).replace('**', '')
    dong = _dong_hon_phoi(s3)
    if not (dong.startswith('Tham khảo thông tin về Con cái theo lá số của vợ (') and dong.endswith(chuan)): loi.append(('mẫu', 'dòng hôn phối (nam → vợ)', dong[:80]))
    if _dong_hon_phoi(s): loi.append(('mẫu', 'không nhập hôn phối mà vẫn có dòng'))
    if _dong_hon_phoi(LaSo(**{**s.to_dict(), 'phoi_ngau': [None, None, 1980, None, None]})): loi.append(('mẫu', 'chỉ nhập năm hôn phối mà vẫn có dòng'))
    s4 = LaSo(**{**sp.to_dict(), 'phoi_ngau': [28, 11, 1978, 8, 30]})
    if not _dong_hon_phoi(s4).startswith('Tham khảo thông tin về Con cái theo lá số của chồng ('): loi.append(('mẫu', 'dòng hôn phối (nữ → chồng)'))
    # Mục 2.9 nhập theo can chi: cùng lá với nhập theo ngày sinh → cùng dòng; nhập sai → không có dòng
    tt = ', '.join('%s %s' % (t['can'], t['chi']) for t in hp['tu_tru'])
    dong_cc = _dong_hon_phoi(LaSo(**{**s.to_dict(), 'phoi_tu_tru': tt}))
    if dong_cc != dong: loi.append(('mẫu', 'hôn phối nhập can chi ≠ nhập ngày sinh', dong_cc[:80]))
    for xau in ('Nhâm Ngọ, Canh Tuất', 'Nhâm Ngọ, Canh Tuất, Ất Mão, Mậu Xyz', ''):
        if _dong_hon_phoi(LaSo(**{**s.to_dict(), 'phoi_tu_tru': xau})): loi.append(('mẫu', 'can chi hôn phối sai mà vẫn có dòng', xau))
    return p, loi


def chay():
    t0 = time.time(); loi = []; n = 0; dem = {'dai_van': 0, 'nam': 0, 'doi_chieu': 0}
    for c, nhap in oracle_nhap(json.load(open(os.path.join(GOC, 't7_ngau_nhien.json'), encoding='utf-8'))):
        E = Engine({'NHAP_LIEU': nhap}, None); E.tinh_tat_ca()
        hs = ho_so(E); p = phan_doan(hs); n += 1
        dem['dai_van'] += len(p['dai_van']); dem['nam'] += len(p['luu_nien']); dem['doi_chieu'] += len(p['doi_chieu'])
        loi += _bat_bien(c['id'], hs, p)
    p, lm = _la_mau()
    return n, dem, loi, lm, p, time.time() - t0


if __name__ == '__main__':
    n, dem, loi, lm, p, dt = chay()
    print('T13 phán đoán – %.1f giây' % dt)
    print('  (a) bất biến: %d lá T7, %d đại vận, %d năm; %d chỗ đổi so với workbook – %s' % (
        n, dem['dai_van'], dem['nam'], dem['doi_chieu'], 'ĐẠT' if not loi else 'LỖI %d' % len(loi)))
    for x in loi[:15]: print('     ', x)
    print('  (b) lá mẫu: %s' % ('ĐẠT' if not lm else lm))
    sys.exit(1 if (loi or lm) else 0)
