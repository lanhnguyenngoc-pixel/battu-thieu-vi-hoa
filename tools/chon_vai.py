# -*- coding: utf-8 -*-
"""Chọn vai Y/X/M/U/Z bán tự động. Dùng: python3 -m tools.chon_vai "Can4" "Chi4" "vận1,vận2,..." "năm_bd" ["2010:xau,1995:tot"]
Ví dụ: python3 -m tools.chon_vai "Kỷ Canh Kỷ Mậu" "Sửu Ngọ Mão Dần" "Kỷ Tị,Mậu Thìn,Đinh Mão,Bính Dần,Ất Sửu,Giáp Tý" 1953 "2010:xau"
In 5 khả năng M, mỗi khả năng kèm vận thịnh/suy; có sự kiện thì xếp theo mức khớp."""
import sys
from battu_core.tong_hop import vai_tro as V

def main(a):
    can = a[0].split(); chi = a[1].split()
    van = a[2].split(','); n0 = int(a[3])
    dv = [{'ten': t, 'nam_bat_dau': n0 + 10 * i, 'nam_ket_thuc': n0 + 10 * i + 9} for i, t in enumerate(van)]
    sk = [{'nam': int(x.split(':')[0]), 'loai': x.split(':')[1]} for x in a[4].split(',')] if len(a) > 4 else None
    kq = V.phan_tich(can, chi, dv, n0, dv[-1]['nam_ket_thuc'])
    print(V.bang_kha_nang(kq, sk))

if __name__ == '__main__': main(sys.argv[1:])
