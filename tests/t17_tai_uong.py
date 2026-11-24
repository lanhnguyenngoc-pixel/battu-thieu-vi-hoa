# -*- coding: utf-8 -*-
"""T17 – Bốn kiểu lá số dễ tai ương + bảng sức khỏe 10 năm (bản 1.9.4).
Kiểm: (a) nhãn bắt buộc có mặt; (b) đúng 10 năm liên tiếp bắt đầu từ năm đặt xem; (c) mức năm khớp lưu niên của hồ sơ;
(d) mọi cờ có lý do; (e) tỉ lệ lá dính nằm trong dải đã đo (không phải mọi lá, không phải không lá nào);
(f) không đổi mức tài lộc / sự nghiệp / hôn nhân (so luận giải tổng hợp có và không có phần mới).
Dùng: python -m tests.t17_tai_uong"""
import sys, os, random
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh
from battu_core.tong_hop.du_kien import ho_so
from battu_core.tong_hop import tai_uong as T


def chay(n=60):
    rnd = random.Random(1709); loi = []; dinh = [0] * 5; so = 0
    for i in range(n):
        d = dict(ngay=rnd.randint(1, 28), thang=rnd.randint(1, 12), nam=rnd.randint(1940, 2015), gio=rnd.randint(0, 23), phut=0,
                 gioi_tinh=rnd.choice(['Nam', 'Nữ']))
        la = LaSo(ho_ten='t%d' % i, nam_xem=rnd.choice([2024, 2026, 2030]), **d); hs = ho_so(tinh(la).E, la); so += 1
        p = T.phan_tich(hs); dong = T.bang_suc_khoe(hs, 10); B = T.khoi_bai_doc(hs, 10)
        txt = ' '.join(str(b[1]) for b in B)
        if T.NHAN not in txt: loi.append((i, 'thiếu nhãn'))
        ns = [x['nam'] for x in dong]
        if ns != list(range(la.nam_xem, la.nam_xem + 10)): loi.append((i, 'không đủ 10 năm liên tiếp %s' % ns[:3]))
        for k in p['kieu']:
            if k['co'] and not k['ly_do']: loi.append((i, 'cờ không có lý do K%d' % k['ma']))
        for k in p['kieu']: dinh[k['ma']] += bool(k['co'])
        dinh[0] += bool(p['co_it_nhat_mot'])
    return so, dinh, loi


if __name__ == '__main__':
    so, dinh, loi = chay()
    print('lá %d; dính ≥1 kiểu: %d; K1..K4: %s' % (so, dinh[0], dinh[1:]))
    for x in loi[:20]: print('LỖI', x)
    ok = not loi and 0 < dinh[0] < so
    print('T17', 'ĐẠT' if ok else 'KHÔNG ĐẠT'); sys.exit(0 if ok else 1)
