# -*- coding: utf-8 -*-
"""Module sheet KIEM_THU – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.5.xlsx (SHA-256 f19f3f74b720c1a067a640a31d58441793bf64bfdd12629f17ba4be6623c56f4).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'KIEM_THU'
SI = 26

def f0(E, r, c):
    # Ô R92, S92, T92, U92, W92, X92
    # = COUNTIF(R6:R90,"Khớp")&"/"&(COUNTIF(R6:R90,"Khớp")+COUNTIF(R6:R90,"LỆCH"))
    return X.cat(X.cat(X.COUNTIF(E.rg(26,r-86,c,r-2,c), 'Khớp'),'/'),X.add(X.COUNTIF(E.rg(26,r-86,c,r-2,c), 'Khớp'),X.COUNTIF(E.rg(26,r-86,c,r-2,c), 'LỆCH')))

def f1(E, r, c):
    # Ô C95
    # = NHAP_LIEU!C13
    return E.v(2,r-82,c)

def f2(E, r, c):
    # Ô C96
    # = NHAP_LIEU!C23&" | "&NHAP_LIEU!D23&" | "&NHAP_LIEU!E23&" | "&NHAP_LIEU!F23
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(2,r-73,c),' | '),E.v(2,r-73,c+1)),' | '),E.v(2,r-73,c+2)),' | '),E.v(2,r-73,c+3))

def f3(E, r, c):
    # Ô C97
    # = DUNG_HY_KY!$C$6
    return E.v(8,6,3)

def f4(E, r, c):
    # Ô C98
    # = DUNG_HY_KY!$D$6
    return E.v(8,6,4)

def f5(E, r, c):
    # Ô C99
    # = IF(DUNG_HY_KY!$C$72<>"","Cách "&DUNG_HY_KY!$C$72,DUNG_HY_KY!$C$106)
    return (X.cat('Cách ',E.v(8,72,3)) if X.truthy(X.ne(E.v(8,72,3),'')) else E.v(8,106,3))

def f6(E, r, c):
    # Ô C100
    # = IF(DUNG_HY_KY!$D$72<>"","Cách "&DUNG_HY_KY!$D$72,DUNG_HY_KY!$D$106)
    return (X.cat('Cách ',E.v(8,72,4)) if X.truthy(X.ne(E.v(8,72,4),'')) else E.v(8,106,4))

def f7(E, r, c):
    # Ô C101
    # = IFERROR(INDEX($K$6:$K$90,MATCH(NHAP_LIEU!$C$13,$B$6:$B$90,0)),"(tự nhập)")
    return X.IFERROR(lambda: X.INDEX(E.rg(26,6,11,90,11), X.MATCH(E.v(2,13,3), E.rg(26,6,2,90,2), 0.0)), lambda: '(tự nhập)')

def f8(E, r, c):
    # Ô C102
    # = IFERROR(INDEX($L$6:$L$90,MATCH(NHAP_LIEU!$C$13,$B$6:$B$90,0)),"(tự nhập)")
    return X.IFERROR(lambda: X.INDEX(E.rg(26,6,12,90,12), X.MATCH(E.v(2,13,3), E.rg(26,6,2,90,2), 0.0)), lambda: '(tự nhập)')

CELLS = {
    'R92': (f0, 92, 18),
    'S92': (f0, 92, 19),
    'T92': (f0, 92, 20),
    'U92': (f0, 92, 21),
    'W92': (f0, 92, 23),
    'X92': (f0, 92, 24),
    'C95': (f1, 95, 3),
    'C96': (f2, 96, 3),
    'C97': (f3, 97, 3),
    'C98': (f4, 98, 3),
    'C99': (f5, 99, 3),
    'C100': (f6, 100, 3),
    'C101': (f7, 101, 3),
    'C102': (f8, 102, 3),
}
