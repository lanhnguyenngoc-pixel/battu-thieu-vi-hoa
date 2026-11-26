# -*- coding: utf-8 -*-
"""Dịch vụ tính một lá số cho giao diện."""
import time
from battu_core.engine import Engine
from battu_core.nhap_lieu import LaSo
from battu_core.du_lieu import model


class KetQua:
    def __init__(self, la_so, E, ms):
        self.la_so = la_so; self.E = E; self.ms = ms

    def o(self, sheet, ref):
        v = self.E.gia_tri(sheet, ref)
        return v

    def chua_cai(self):
        M = self.E.M
        from battu_core.du_lieu import ref_of
        return sorted('%s!%s' % (M.sheets[si], ref_of(r, c)) for si, r, c in self.E.chua_cai)


def tinh(la_so: LaSo) -> KetQua:
    """Tính lá số. Với cách nhập theo ngày giờ: lần 1 tính tứ trụ tự động, rồi điền chính tứ trụ đó vào
    C7:F8 (như người dùng Excel nhập can chi khớp ngày sinh) để NHAP_LIEU G7 không báo "C7:F8 khác"."""
    t0 = time.time()
    E = Engine(la_so.o_nhap())
    if la_so.cach_nhap in ('duong', 'am') and (la_so.la_mau or 'Tự nhập') == 'Tự nhập':
        can = [E.gia_tri('NHAP_LIEU', c + '37') for c in 'CDEF']
        chi = [E.gia_tri('NHAP_LIEU', c + '38') for c in 'CDEF']
        if all(isinstance(x, str) and x for x in can + chi) and (can != la_so.can or chi != la_so.chi):
            la_so.can, la_so.chi = can, chi
            E = Engine(la_so.o_nhap())
    if la_so.theo_doi:
        # Đánh giá thực tế nhập theo NĂM; dòng của năm do lõi tính (THEO_DOI A12:A111 không phụ thuộc cột N/O/P).
        nam = [E.gia_tri('THEO_DOI', 'A%d' % r) for r in range(12, 112)]
        E = Engine(la_so.o_nhap(nam))
    E.tinh_tat_ca()   # mọi sheet đã cài ở giai đoạn hiện tại
    return KetQua(la_so, E, (time.time() - t0) * 1000)


def danh_sach_la_mau():
    """Tên các lá số mẫu trong KIEM_THU B6:B90 (để chọn ở NHAP_LIEU C13)."""
    M = model(); k = M.consts[M.si['KIEM_THU']]
    return ['Tự nhập'] + [k[(r, 2)] for r in range(6, 91) if isinstance(k.get((r, 2)), str) and k.get((r, 2))]


def so_voi_gio_dong_ho(la_so: LaSo):
    """Khi bật giờ mặt trời thật: so tứ trụ với cách tính theo giờ đồng hồ. → (chuỗi giờ, danh sách trụ khác) hoặc None.
    Chỉ tính phần tứ trụ (NHAP_LIEU hàng 37–38), không tính cả lá số."""
    import copy
    q = la_so.gio_quy_doi()
    if not q: return None
    d, m, y, gio, phut, lech = q
    dh = copy.copy(la_so); dh.gio_mat_troi = False
    def tru(ls):
        E = Engine(ls.o_nhap())
        return [(E.gia_tri('NHAP_LIEU', c + '37'), E.gia_tri('NHAP_LIEU', c + '38')) for c in 'CDEF']
    a, b = tru(dh), tru(la_so)
    khac = [n for n, x, z in zip(('năm', 'tháng', 'ngày', 'giờ'), a, b) if x != z]
    gio_dong_ho = '%02d:%02d' % (la_so.gio, la_so.phut or 0)
    mt = '%02d:%02d %02d/%02d/%d' % (gio, phut, d, m, y)
    return gio_dong_ho, mt, lech, khac, ['%s %s' % x for x in a], ['%s %s' % x for x in b]
