# -*- coding: utf-8 -*-
"""T20 – mô-đun hệ 5 vai (vai_tro.py), đo trên các lá số sách có nhãn thịnh/suy từng vận (tr.125–216).
Cổng kiểm: vận không mang hành M phải khớp sách 100%; vận mang hành M và năm nguy chỉ ĐO, không chặn (xem tài liệu dự án)."""
import sys
from battu_core.tong_hop import vai_tro as V

# (tên, can4, chi4, M đúng theo sách, {tên vận: nhãn sách 't'|'s'}, năm mất/biến cố nguy)
CA = [
 ('NguyenDu', ['Giáp', 'Bính', 'Kỷ', 'Bính'], ['Thân', 'Tý', 'Sửu', 'Dần'], 'Thổ',
  {'Đinh Sửu': 't', 'Mậu Dần': 't', 'Kỷ Mão': 't', 'Canh Thìn': 's', 'Tân Tị': 's', 'Nhâm Ngọ': 's', 'Quý Mùi': 's'}, 1820),
 ('Mandela', ['Mậu', 'Kỷ', 'Bính', None], ['Ngọ', 'Mùi', 'Dần', None], 'Hỏa',
  {'Bính Dần': 't', 'Đinh Mão': 't', 'Mậu Thìn': 's'}, 2013),
 ('Kaczynski', ['Kỷ', 'Canh', 'Kỷ', None], ['Sửu', 'Ngọ', 'Mão', None], 'Hỏa',
  {'Kỷ Tị': 's', 'Mậu Thìn': 's', 'Đinh Mão': 't', 'Bính Dần': 't', 'Ất Sửu': 't'}, 2010),
 ('Nam1990b', ['Canh', 'Bính', 'Ất', 'Bính'], ['Ngọ', 'Tuất', 'Hợi', 'Tý'], 'Thổ',
  {'Đinh Hợi': 't', 'Mậu Tý': 't', 'Kỷ Sửu': 's'}, 2012),
]


def chay():
    ok_khong_m = tot_khong_m = ok_m = tot_m = 0; loi = []
    for ten, c, ch, m, nhan, _ in CA:
        vai = V.vai_theo_M(m)
        for van, nh in nhan.items():
            can, chi = van.split()
            nua = V.thinh_suy_van(vai, can, chi)
            la_m = V.CAN_H[can] == m
            kq = {'thinh': 't', 'suy': 's', 'dac_dia': '?'}[nua]
            if la_m: tot_m += 1; ok_m += kq == nh
            else:
                tot_khong_m += 1; ok_khong_m += kq == nh
                if kq != nh: loi.append((ten, van, nh, kq))
    # cấu trúc: 5 khả năng, công thức, trục
    for ten, c, ch, m, _, _ in CA:
        uv = V.ung_vien(c, ch)
        if len(uv) != 5 or {u['m'] for u in uv} != set(V.HANH): loi.append((ten, 'ứng viên'))
        t, h = V.truc(ch[1])
        if not (t and h): loi.append((ten, 'trục'))
    return ok_khong_m, tot_khong_m, ok_m, tot_m, loi


if __name__ == '__main__':
    a, b, c, d, loi = chay()
    print('T20 hệ 5 vai: vận không mang hành M khớp sách %d/%d; vận mang hành M khớp %d/%d (đo, không chặn)' % (a, b, c, d))
    for x in loi: print('  LỖI', x)
    print('T20', 'ĐẠT' if not loi else 'KHÔNG ĐẠT'); sys.exit(1 if loi else 0)
