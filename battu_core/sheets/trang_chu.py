# -*- coding: utf-8 -*-
"""Module sheet TRANG_CHU – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: B69n.xlsx (SHA-256 9228d1ab613379f50ca7126e315efcbb940599c6e1784212c489756674ca8b67).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'TRANG_CHU'
SI = 0

def f0(E, r, c):
    # Ô C6
    # = KET_LUAN!$C$15&" · "&KET_LUAN!$D$15&" · "&KET_LUAN!$E$15&" · "&KET_LUAN!$F$15
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,15,3),' · '),E.v(3,15,4)),' · '),E.v(3,15,5)),' · '),E.v(3,15,6))

def f1(E, r, c):
    # Ô C7
    # = KET_LUAN!$C$202
    return E.v(3,202,3)

def f2(E, r, c):
    # Ô C8
    # = LEFT(KET_LUAN!$C$163,220)
    return X.LEFT(E.v(3,163,3), 220.0)

def f3(E, r, c):
    # Ô C9
    # = IF(THEO_DOI!$E$4=0,"Chưa ghi sự kiện thực tế – vào sheet THEO_DOI để ghi và so khớp PP1/PP2.",THEO_DOI!$E$4&" năm đã ghi · "&THEO_DOI!$E$7)
    return ('Chưa ghi sự kiện thực tế – vào sheet THEO_DOI để ghi và so khớp PP1/PP2.' if X.truthy(X.eq(E.v(5,4,5),0.0)) else X.cat(X.cat(E.v(5,4,5),' năm đã ghi · '),E.v(5,7,5)))

CELLS = {
    'C6': (f0, 6, 3),
    'C7': (f1, 7, 3),
    'C8': (f2, 8, 3),
    'C9': (f3, 9, 3),
}
