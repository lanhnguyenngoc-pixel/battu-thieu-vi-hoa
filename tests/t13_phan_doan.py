# -*- coding: utf-8 -*-
"""T13 – Lớp phán đoán, gỡ mâu thuẫn (đề xuất hoàn thiện 15/11/2026, giai đoạn T3).
  (a) Bất biến trên 20 lá T7 + lá mẫu:
      • mỗi đại vận / mỗi nửa vận / mỗi năm có ĐÚNG MỘT mức trong thang 5 mức;
      • năm có dấu hiệu hạn nặng không cao hơn trần (quy tắc luu_nien.tran_khi_han_nang);
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
        for k in ('muc', 'muc_5_nam_dau', 'muc_5_nam_sau'):
            if d[k] not in MUC: loi.append((ten, d['ten'], k, d[k]))
    vt = VaiTro(hs)
    for y, h in zip(p['luu_nien'], hs['luu_nien']):
        if y['muc'] not in MUC: loi.append((ten, y['nam'], 'mức', y['muc']))
        if y['han_nang'] and MUC.index(y['muc']) > MUC.index(qt['luu_nien']['tran_khi_han_nang']):
            loi.append((ten, y['nam'], 'hạn nặng mà mức', y['muc']))
        xd = sum(1 for q in h.get('quan_he', []) if q['loai'] == 'xung' and q.get('cham_dung'))
        if xd >= 2 and vt.can(h['can']) <= 0 and vt.chi(h['chi']) <= 0 and y['muc'] not in ('Xấu', 'Rất xấu'):
            loi.append((ten, y['nam'], 'xung 2 chi dụng thần mà mức', y['muc']))
    ty = p['tho_yeu']
    if len(ty['dau_hieu_tot']) >= qt['tho_yeu']['so_dau_hieu_nguoc'] and 'yểu' in ty['muc'].lower():
        loi.append((ten, 'thọ – yểu', ty['muc']))
    if json.loads(json.dumps(p, ensure_ascii=False)) != json.loads(json.dumps(phan_doan(hs), ensure_ascii=False)):
        loi.append((ten, 'không tất định'))
    return loi


MAU = {'dai_van': {'Giáp Tí': ('Xấu', 'Rất xấu'), 'Bính Dần': ('Tốt', 'Rất tốt'), 'Mậu Thìn': ('Tốt', 'Rất tốt'),
                   'Kỷ Tị': ('Tốt', 'Rất tốt'), 'Nhâm Thân': ('Xấu', 'Rất xấu')},
       'luu_nien': {2026: ('Rất tốt',), 2027: ('Tốt',), 2029: ('Tốt', 'Rất tốt'), 2030: ('Tốt', 'Rất tốt'),
                    2032: ('Rất xấu',), 2034: ('Trung bình',), 2035: ('Xấu', 'Rất xấu'), 2036: ('Tốt', 'Rất tốt')}}


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
    if p['tho_yeu']['muc'] != 'Căn cứ ngược chiều': loi.append(('mẫu', 'thọ – yểu', p['tho_yeu']['muc']))
    if p['tai_loc']['muc'] != 'Khá giả': loi.append(('mẫu', 'tài lộc', p['tai_loc']['muc']))
    if not p['vuong_suy']['sat_nguong']: loi.append(('mẫu', 'vượng suy phải sát ngưỡng'))
    if not any(x['muc'] == 'Thời điểm kết hôn' for x in p['an_muc']): loi.append(('mẫu', 'chưa ẩn thời điểm kết hôn (48 tuổi)'))
    if p['luc_than']['Mẹ']['muc'] != 'Kém' or not p['luc_than']['Mẹ']['nguon_khac']: loi.append(('mẫu', 'mẹ: giữ Kém, có nguồn khác nói ngược'))
    s2 = LaSo(**{**s.to_dict(), 'so_con': 2, 'tinh_trang_hon_nhan': 'Đã kết hôn'})
    if not any(x['muc'] == 'Con đầu lòng' for x in phan_doan(hs, s2)['an_muc']): loi.append(('mẫu', 'đã có con mà chưa ẩn "Con đầu lòng"'))
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
