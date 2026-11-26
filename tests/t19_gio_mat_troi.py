# -*- coding: utf-8 -*-
"""T19 – Giờ mặt trời thật (tùy chọn, duyệt 26/11/2026).
(1) Mặc định tắt: kết quả giống hệt khi không có trường mới. (2) Bật: ô nhập C26:G26 = giờ đã quy đổi (đúng công thức),
sang ngày khi quá nửa đêm; tứ trụ đổi đúng khi qua ranh giờ. (3) Thiếu nơi sinh → báo lỗi. (4) Không rõ giờ / nhập can chi: bỏ qua.
(5) Phương trình thời gian khớp mốc đã biết (±1 phút): 11/2 ≈ −14, 3/11 ≈ +16."""
import datetime
from battu_core.nhap_lieu import LaSo, LoiNhap
from battu_core import gio_mat_troi as g
from battu_core.dich_vu import so_voi_gio_dong_ho


def chay():
    kq = {}
    a = LaSo(ngay=15, thang=8, nam=2012, gio=0, phut=45)
    b = LaSo(ngay=15, thang=8, nam=2012, gio=0, phut=45, gio_mat_troi=False, noi_sinh='Hà Nội')
    kq['mặc định tắt: ô nhập không đổi'] = a.o_nhap() == b.o_nhap() and a.gio_quy_doi() is None
    kq['phương trình thời gian 11/2'] = abs(g.phuong_trinh_thoi_gian(datetime.date(2020, 2, 11)) + 14.2) < 1
    kq['phương trình thời gian 3/11'] = abs(g.phuong_trinh_thoi_gian(datetime.date(2020, 11, 3)) - 16.4) < 1
    c = LaSo(ngay=11, thang=2, nam=2020, gio=7, phut=5, gio_mat_troi=True, noi_sinh='Hà Nội')
    nl = c.o_nhap()['NHAP_LIEU']
    kq['quy đổi 07:05 → 06:54'] = (nl['F26'], nl['G26']) == (6.0, 54.0)
    r = so_voi_gio_dong_ho(c)
    kq['trụ giờ đổi Thìn → Mão'] = r is not None and r[3] == ['giờ']
    d = LaSo(ngay=1, thang=3, nam=2021, gio=0, phut=3, gio_mat_troi=True, kinh_do=104.0)
    nl = d.o_nhap()['NHAP_LIEU']
    kq['sang ngày trước'] = (nl['C26'], nl['D26'], nl['F26']) == (28.0, 2.0, 23.0)
    try: LaSo(ngay=1, thang=1, nam=2000, gio=8, phut=0, gio_mat_troi=True).o_nhap(); ok = False
    except LoiNhap: ok = True
    kq['thiếu nơi sinh → lỗi'] = ok
    e = LaSo(ngay=1, thang=1, nam=2000, gio=8, phut=0, gio_mat_troi=True, noi_sinh='Hà Nội', khong_ro_gio=True)
    kq['không rõ giờ: bỏ qua'] = e.gio_quy_doi() is None
    f = LaSo(cach_nhap='canchi', can=['Giáp'] * 4, chi=['Tí'] * 4, gio_mat_troi=True, noi_sinh='Hà Nội')
    kq['nhập can chi: bỏ qua'] = f.gio_quy_doi() is None
    h = LaSo.from_dict(c.to_dict())
    kq['lưu/đọc lá số giữ tùy chọn'] = (h.gio_mat_troi, h.noi_sinh) == (True, 'Hà Nội')
    kq['kinh độ nhập tay ưu tiên'] = g.kinh_do_cua('Hà Nội', 104.0) == 104.0
    return kq


if __name__ == '__main__':
    k = chay(); [print(('ĐẠT ' if v else 'LỖI ') + n) for n, v in k.items()]; raise SystemExit(0 if all(k.values()) else 1)
