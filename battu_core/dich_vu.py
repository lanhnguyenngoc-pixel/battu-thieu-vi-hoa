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
    E.tinh_tat_ca()   # mọi sheet đã cài ở giai đoạn hiện tại
    return KetQua(la_so, E, (time.time() - t0) * 1000)


def danh_sach_la_mau():
    """Tên các lá số mẫu trong KIEM_THU B6:B90 (để chọn ở NHAP_LIEU C13)."""
    M = model(); k = M.consts[M.si['KIEM_THU']]
    return ['Tự nhập'] + [k[(r, 2)] for r in range(6, 91) if isinstance(k.get((r, 2)), str) and k.get((r, 2))]
