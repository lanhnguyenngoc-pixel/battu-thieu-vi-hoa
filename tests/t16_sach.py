# -*- coding: utf-8 -*-
"""T16 – Đối chiếu lớp tổng hợp với chính các ví dụ của Thiệu Vĩ Hoa (bộ 3 cuốn P1–P3).
  (a) Đại vận: 54 lá có nhãn tốt/xấu từ lời sách (tai_lieu/giai_doan_A/nhan_dai_van_sach.py), 171 vận.
      Hướng (tốt/xấu) của lớp tổng hợp phải khớp sách KHÔNG KÉM workbook (mục F, P1 tr.124–125).
  (b) Vượng – nhược: mọi lá ví dụ P1–P3 có kết luận vượng/nhược của sách (tests/cases501_*.json)
      và 50 lá rút ngẫu nhiên (hạt giống cố định) – tỉ lệ khớp của lớp tổng hợp ≥ workbook.
Dùng: python -m tests.t16_sach"""
import sys, os, json, time, random
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh
from battu_core.tong_hop.du_kien import ho_so
from battu_core.tong_hop.phan_doan import phan_doan, VaiTro

GOC = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
sys.path.insert(0, os.path.join(GOC, 'tai_lieu', 'giai_doan_A'))
from nhan_dai_van_sach import NHAN


def _h(x):
    x = (x or '').strip().upper()
    return 'T' if x in ('TỐT', 'RẤT TỐT') else ('X' if x in ('XẤU', 'RẤT XẤU') else 'B')


def _la(ten, can, chi, gioi='Nam'):
    s = LaSo(ho_ten=ten, cach_nhap='canchi', can=can, chi=chi, gioi_tinh=gioi, nam_xem=2026)
    return s, ho_so(tinh(s).E, s)


def dai_van():
    ds = {d['ten']: d for d in json.load(open(os.path.join(GOC, 'tai_lieu/giai_doan_A/dai_van_doi_chieu.json'), encoding='utf-8'))}
    n = moi = wb = nguoc = 0
    for ten, nhan in NHAN.items():
        d = ds.get(ten)
        if not d: continue
        tt = [x.strip().split() for x in d['tt'].split(',')]
        s, hs = _la(ten, [x[0] for x in tt], [x[1] for x in tt], d['g']); pd = phan_doan(hs, s)
        for vt, sach in nhan.items():
            if vt > len(hs['dai_van']): continue
            a, b = _h(pd['dai_van'][vt - 1]['muc']), _h(hs['dai_van'][vt - 1].get('ket_luan'))
            n += 1; moi += a == sach; wb += b == sach; nguoc += {a, sach} == {'T', 'X'}
    return n, moi, wb, nguoc


def vuong_suy():
    ca = []
    for f in ('cases501_a.json', 'cases501_b.json'):
        ca += json.load(open(os.path.join(GOC, 'tests', f), encoding='utf-8'))
    ca = [c for c in ca if c.get('st') in ('V', 'N') and c['name'][:2] in ('P1', 'P2', 'P3')]
    random.seed(2026); mau = random.sample(ca, min(50, len(ca)))
    kq = {}
    for ten, bo in (('tất cả', ca), ('50 ngẫu nhiên', mau)):
        moi = wb = 0
        for c in bo:
            s, hs = _la(c['name'], c['can'], c['chi'])
            moi += ('V' if VaiTro(hs).vuong else 'N') == c['st']
            wb += ('V' if 'vượng' in (c.get('wb_p1') or '').lower() else 'N') == c['st']
        kq[ten] = (len(bo), moi, wb)
    return kq


if __name__ == '__main__':
    t = time.time(); dat = True
    n, moi, wb, nguoc = dai_van()
    ok = moi >= wb; dat &= ok
    print('T16 đối chiếu sách Thiệu Vĩ Hoa')
    print('  (a) đại vận: %d vận – lớp tổng hợp %d (%d%%), workbook %d (%d%%), ngược chiều %d – %s'
          % (n, moi, round(100 * moi / n), wb, round(100 * wb / n), nguoc, 'ĐẠT' if ok else 'KHÔNG ĐẠT'))
    for ten, (k, m, w) in vuong_suy().items():
        ok = m >= w; dat &= ok
        print('  (b) vượng – nhược, %s: %d lá – lớp tổng hợp %d (%d%%), workbook %d (%d%%) – %s'
              % (ten, k, m, round(100 * m / k), w, round(100 * w / k), 'ĐẠT' if ok else 'KHÔNG ĐẠT'))
    print('  %.1f giây – %s' % (time.time() - t, 'ĐẠT' if dat else 'KHÔNG ĐẠT'))
    sys.exit(0 if dat else 1)
