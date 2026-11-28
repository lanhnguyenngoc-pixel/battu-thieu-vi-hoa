# -*- coding: utf-8 -*-
"""T20 – mô-đun hệ 5 vai (vai_tro.py), đo trên các lá số sách có nhãn thịnh/suy từng vận (tr.125–216).
Cổng kiểm: vận không mang hành M phải khớp sách 100%; vận mang hành M và năm nguy chỉ ĐO, không chặn (xem tài liệu dự án)."""
import sys
from battu_core.tong_hop import vai_tro as V

# (tên, can4, chi4, M đúng theo sách, {tên vận: nhãn sách 't'|'s'}, năm mất/biến cố nguy)
CA = [
 ('NguyenDu', ['Giáp', 'Bính', 'Kỷ', 'Bính'], ['Thân', 'Tý', 'Sửu', 'Dần'], 'Thổ',
  {'Đinh Sửu': 't', 'Mậu Dần': 't', 'Kỷ Mão': 's', 'Canh Thìn': 's', 'Tân Tị': 's', 'Nhâm Ngọ': 's', 'Quý Mùi': 's'}, 1820),
 ('Mandela', ['Mậu', 'Kỷ', 'Bính', None], ['Ngọ', 'Mùi', 'Dần', None], 'Hỏa',
  {'Bính Dần': 't', 'Đinh Mão': 't', 'Mậu Thìn': 's'}, 2013),
 ('Kaczynski', ['Kỷ', 'Canh', 'Kỷ', None], ['Sửu', 'Ngọ', 'Mão', None], 'Hỏa',
  {'Kỷ Tị': 's', 'Mậu Thìn': 's', 'Đinh Mão': 't', 'Bính Dần': 't', 'Ất Sửu': 't'}, 2010),
 ('Nam1990b', ['Canh', 'Bính', 'Ất', 'Bính'], ['Ngọ', 'Tuất', 'Hợi', 'Tý'], 'Thổ',
  {'Đinh Hợi': 't', 'Mậu Tý': 't', 'Kỷ Sửu': 's'}, 2012),
]


NHAN_THU_TU = {  # dãy vận đầy đủ theo thứ tự thời gian (bảng vận trong sách)
 'NguyenDu': ['Đinh Sửu', 'Mậu Dần', 'Kỷ Mão', 'Canh Thìn', 'Tân Tị', 'Nhâm Ngọ', 'Quý Mùi', 'Giáp Thân'],
 'Mandela': ['Canh Thân', 'Tân Dậu', 'Nhâm Tuất', 'Quý Hợi', 'Giáp Tý', 'Ất Sửu', 'Bính Dần', 'Đinh Mão'],
 'Kaczynski': ['Kỷ Tị', 'Mậu Thìn', 'Đinh Mão', 'Bính Dần', 'Ất Sửu', 'Giáp Tý', 'Quý Hợi', 'Nhâm Tuất'],
 'Nam1990b': ['Đinh Hợi', 'Mậu Tý', 'Kỷ Sửu', 'Canh Dần', 'Tân Mão', 'Nhâm Thìn', 'Quý Tị'],
}


def chay():
    """Cổng kiểm: vận KHÔNG nằm trên đầu đường phân định phải khớp nhãn sách (nhãn đọc từ lời sách + hình hệ trục).
    Vận trên đầu đường (bien) chỉ đo: sách không nhất quán (xem tài liệu dự án)."""
    ok = tot = ok_b = tot_b = 0; loi = []
    for ten, c, ch, m, nhan, _ in CA:
        vai = V.vai_theo_M(m)
        van = [{'ten': t, 'nam_bat_dau': 0, 'nam_ket_thuc': 0} for t in NHAN_THU_TU[ten]]
        r = V.nua_theo_vong(vai, van, c, ch)
        if r is None: loi.append((ten, 'không tìm được đường phân định')); continue
        for v, (nua, bien) in zip(van, r):
            if v['ten'] not in nhan: continue
            dung = (nhan[v['ten']] == 't') == (nua == 'thinh')
            if bien: tot_b += 1; ok_b += dung
            else:
                tot += 1; ok += dung
                if not dung: loi.append((ten, v['ten'], nhan[v['ten']], nua))
        uv = V.ung_vien(c, ch)
        if len(uv) != 5 or {u['m'] for u in uv} != set(V.HANH): loi.append((ten, 'ứng viên'))
        t, h = V.truc(ch[1])
        if not (t and h): loi.append((ten, 'trục'))
    return ok, tot, ok_b, tot_b, loi


if __name__ == '__main__':
    a, b, c, d, loi = chay()
    print('T20 hệ 5 vai: vận ngoài đầu đường khớp sách %d/%d; vận trên đầu đường khớp %d/%d (đo, không chặn)' % (a, b, c, d))
    for x in loi: print('  LỖI', x)
    print('T20', 'ĐẠT' if not loi else 'KHÔNG ĐẠT'); sys.exit(1 if loi else 0)
