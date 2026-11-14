# -*- coding: utf-8 -*-
"""T9 – Nhập sự việc SU_KIEN (bản 1.1.0) trên 20 lá T7, mỗi lá 6 sự việc (tests/bien_the.json → T7_SU_KIEN):
(a) đáp án LibreOffice (tools/oracle_mo_rong.py) – so MỌI ô công thức của 33 sheet (t1b.chay_bien_the);
(b) đường nhập của phần mềm: LaSo.su_kien → o_nhap()['SU_KIEN'] phải sinh ĐÚNG các ô B/C/D mà oracle đã nhập;
(c) lưu SQLite → mở lại giữ nguyên danh sách sự việc.
Dùng: python -m tests.t9_su_kien"""
import sys, os, json, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from tests import t1b
from battu_core.nhap_lieu import LaSo

GOC = os.path.dirname(os.path.abspath(__file__))


def chay():
    n, tong, dung, lech = t1b.chay_bien_the(None, os.path.join(GOC, 'dap_an_mo_rong'), 'T7_SU_KIEN', os.path.join(GOC, 't7_ngau_nhien.json'))
    bt = json.load(open(os.path.join(GOC, 'bien_the.json'), encoding='utf-8'))['T7_SU_KIEN']['rieng']
    dung_b = 0; tong_b = 0
    for cid, d in bt.items():
        o = d['SU_KIEN']; ds = []
        for r in range(6, 36):
            if 'B%d' % r in o or 'C%d' % r in o:
                ds.append([int(o['B%d' % r]) if 'B%d' % r in o else None, o.get('C%d' % r), o.get('D%d' % r)])
        sinh = LaSo(su_kien=ds, cach_nhap='canchi').o_nhap().get('SU_KIEN', {})
        sinh = {k: v for k, v in sinh.items() if v is not None}
        tong_b += 1; dung_b += sinh == o
    from battu_core.kho_la_so import Kho
    from battu_core.dich_vu import tinh
    kho = Kho(os.path.join(tempfile.mkdtemp(), 'sk.sqlite'))
    s = LaSo(ngay=26, thang=11, nam=1957, gio=6, phut=0, nam_xem=2026, su_kien=[[1983, 'Hôn nhân – kết hôn', 'cưới'], [2008, 'Tài chính – khó khăn', None]])
    i = kho.luu(tinh(s), 'sk'); ok_c = kho.la_so(i).su_kien == s.su_kien
    return n, tong, dung, lech, (dung_b, tong_b), ok_c


if __name__ == '__main__':
    n, tong, dung, lech, (db, tb), okc = chay()
    print('T9 sự việc – %d lá: %d / %d ô' % (n, sum(dung.values()), sum(tong.values())), '| SU_KIEN %d / %d' % (dung['SU_KIEN'], tong['SU_KIEN']))
    print('  đường nhập LaSo.su_kien = ô oracle: %d / %d lá · SQLite giữ sự việc: %s' % (db, tb, 'ĐẠT' if okc else 'LỖI'))
    for x in lech[:20]: print('  LỆCH', x)
    sys.exit(0 if (not lech and db == tb and okc and n == 20) else 1)
