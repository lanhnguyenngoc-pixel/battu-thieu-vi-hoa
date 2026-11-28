# -*- coding: utf-8 -*-
"""Module sheet THONG_CAN – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: 6.14.xlsx (SHA-256 b6c1126b9b23cc29eb1d8e3cf5b84b07a74bfc629e1cb6a6cc9f6b4d74f758cd).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'THONG_CAN'
SI = 11

def f0(E, r, c):
    # Ô C6
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(0=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$C$21=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,3),E.v(23,7,3))) else 0.0))

def f1(E, r, c):
    # Ô D6
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(0=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$C$21=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,3),E.v(23,7,4))) else 0.0))

def f2(E, r, c):
    # Ô E6
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(0=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$C$21=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,3),E.v(23,7,5))) else 0.0))

def f3(E, r, c):
    # Ô F6
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(0=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$C$21=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,3),E.v(23,7,6))) else 0.0))

def f4(E, r, c):
    # Ô C7
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(1=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$C$22=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,3),E.v(23,7,3))) else 0.0))

def f5(E, r, c):
    # Ô D7
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(1=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$C$22=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,3),E.v(23,7,4))) else 0.0))

def f6(E, r, c):
    # Ô E7
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(1=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$C$22=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,3),E.v(23,7,5))) else 0.0))

def f7(E, r, c):
    # Ô F7
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(1=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$C$22=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,3),E.v(23,7,6))) else 0.0))

def f8(E, r, c):
    # Ô C8
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(2=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$C$23=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,3),E.v(23,7,3))) else 0.0))

def f9(E, r, c):
    # Ô D8
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(2=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$C$23=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,3),E.v(23,7,4))) else 0.0))

def f10(E, r, c):
    # Ô E8
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(2=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$C$23=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,3),E.v(23,7,5))) else 0.0))

def f11(E, r, c):
    # Ô F8
    # = IF(PP1_DINH_TINH!$C$27>=0,IF(AND(2=0,PP1_DINH_TINH!$C$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$C$23=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,3),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,3),E.v(23,7,6))) else 0.0))

def f12(E, r, c):
    # Ô C9
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(0=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$D$21=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,4),E.v(23,7,3))) else 0.0))

def f13(E, r, c):
    # Ô D9
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(0=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$D$21=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,4),E.v(23,7,4))) else 0.0))

def f14(E, r, c):
    # Ô E9
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(0=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$D$21=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,4),E.v(23,7,5))) else 0.0))

def f15(E, r, c):
    # Ô F9
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(0=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$D$21=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,4),E.v(23,7,6))) else 0.0))

def f16(E, r, c):
    # Ô C10
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(1=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$D$22=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,4),E.v(23,7,3))) else 0.0))

def f17(E, r, c):
    # Ô D10
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(1=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$D$22=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,4),E.v(23,7,4))) else 0.0))

def f18(E, r, c):
    # Ô E10
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(1=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$D$22=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,4),E.v(23,7,5))) else 0.0))

def f19(E, r, c):
    # Ô F10
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(1=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$D$22=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,4),E.v(23,7,6))) else 0.0))

def f20(E, r, c):
    # Ô C11
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(2=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$D$23=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,4),E.v(23,7,3))) else 0.0))

def f21(E, r, c):
    # Ô D11
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(2=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$D$23=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,4),E.v(23,7,4))) else 0.0))

def f22(E, r, c):
    # Ô E11
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(2=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$D$23=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,4),E.v(23,7,5))) else 0.0))

def f23(E, r, c):
    # Ô F11
    # = IF(PP1_DINH_TINH!$D$27>=0,IF(AND(2=0,PP1_DINH_TINH!$D$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$D$23=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,4),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,4),E.v(23,7,6))) else 0.0))

def f24(E, r, c):
    # Ô C12
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(0=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,3))) else 0.0))

def f25(E, r, c):
    # Ô D12
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(0=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,4))) else 0.0))

def f26(E, r, c):
    # Ô E12
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(0=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,5))) else 0.0))

def f27(E, r, c):
    # Ô F12
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(0=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,6))) else 0.0))

def f28(E, r, c):
    # Ô C13
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(1=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,3))) else 0.0))

def f29(E, r, c):
    # Ô D13
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(1=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,4))) else 0.0))

def f30(E, r, c):
    # Ô E13
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(1=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,5))) else 0.0))

def f31(E, r, c):
    # Ô F13
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(1=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,6))) else 0.0))

def f32(E, r, c):
    # Ô C14
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(2=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,3))) else 0.0))

def f33(E, r, c):
    # Ô D14
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(2=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,4))) else 0.0))

def f34(E, r, c):
    # Ô E14
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(2=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,5))) else 0.0))

def f35(E, r, c):
    # Ô F14
    # = IF(PP1_DINH_TINH!$E$27>=0,IF(AND(2=0,PP1_DINH_TINH!$E$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,5),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,5),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,6))) else 0.0))

def f36(E, r, c):
    # Ô C15
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(0=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$F$21=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,6),E.v(23,7,3))) else 0.0))

def f37(E, r, c):
    # Ô D15
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(0=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$F$21=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,6),E.v(23,7,4))) else 0.0))

def f38(E, r, c):
    # Ô E15
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(0=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$F$21=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,6),E.v(23,7,5))) else 0.0))

def f39(E, r, c):
    # Ô F15
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(0=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$F$21=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(0.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,21,6),E.v(23,7,6))) else 0.0))

def f40(E, r, c):
    # Ô C16
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(1=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$F$22=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,6),E.v(23,7,3))) else 0.0))

def f41(E, r, c):
    # Ô D16
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(1=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$F$22=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,6),E.v(23,7,4))) else 0.0))

def f42(E, r, c):
    # Ô E16
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(1=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$F$22=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,6),E.v(23,7,5))) else 0.0))

def f43(E, r, c):
    # Ô F16
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(1=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$F$22=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(1.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,22,6),E.v(23,7,6))) else 0.0))

def f44(E, r, c):
    # Ô C17
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(2=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$C$7),1,0),IF(TRUNG_GIAN!$F$23=TRUNG_GIAN!$C$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,3)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,6),E.v(23,7,3))) else 0.0))

def f45(E, r, c):
    # Ô D17
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(2=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$D$7),1,0),IF(TRUNG_GIAN!$F$23=TRUNG_GIAN!$D$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,4)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,6),E.v(23,7,4))) else 0.0))

def f46(E, r, c):
    # Ô E17
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(2=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$E$7),1,0),IF(TRUNG_GIAN!$F$23=TRUNG_GIAN!$E$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,5)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,6),E.v(23,7,5))) else 0.0))

def f47(E, r, c):
    # Ô F17
    # = IF(PP1_DINH_TINH!$F$27>=0,IF(AND(2=0,PP1_DINH_TINH!$F$27=TRUNG_GIAN!$F$7),1,0),IF(TRUNG_GIAN!$F$23=TRUNG_GIAN!$F$7,1,0))
    return ((1.0 if X.truthy(X.AND(X.eq(2.0,0.0), X.eq(E.v(9,27,6),E.v(23,7,6)))) else 0.0) if X.truthy(X.ge(E.v(9,27,6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,23,6),E.v(23,7,6))) else 0.0))

def f48(E, r, c):
    # Ô C20, D20, E20, F20, C21, D21, E21, F21, C22, D22, E22, F22
    # = IF(AND(C6=1,PP1_DINH_TINH!$C$29=0,PP1_DINH_TINH!$C$28=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(E.v(9,29,3),0.0), X.eq(E.v(9,28,3),0.0))) else 0.0)

def f49(E, r, c):
    # Ô C23, D23, E23, F23, C24, D24, E24, F24, C25, D25, E25, F25
    # = IF(AND(C9=1,PP1_DINH_TINH!$D$29=0,PP1_DINH_TINH!$D$28=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(E.v(9,29,4),0.0), X.eq(E.v(9,28,4),0.0))) else 0.0)

def f50(E, r, c):
    # Ô C26, D26, E26, F26, C27, D27, E27, F27, C28, D28, E28, F28
    # = IF(AND(C12=1,PP1_DINH_TINH!$E$29=0,PP1_DINH_TINH!$E$28=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(E.v(9,29,5),0.0), X.eq(E.v(9,28,5),0.0))) else 0.0)

def f51(E, r, c):
    # Ô C29, D29, E29, F29, C30, D30, E30, F30, C31, D31, E31, F31
    # = IF(AND(C15=1,PP1_DINH_TINH!$F$29=0,PP1_DINH_TINH!$F$28=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(E.v(9,29,6),0.0), X.eq(E.v(9,28,6),0.0))) else 0.0)

def f52(E, r, c):
    # Ô C34
    # = IF(AND(C20=1,0=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,3)))) else 0.0)

def f53(E, r, c):
    # Ô D34
    # = IF(AND(D20=1,0=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,4)))) else 0.0)

def f54(E, r, c):
    # Ô E34
    # = IF(AND(E20=1,0=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,5)))) else 0.0)

def f55(E, r, c):
    # Ô F34
    # = IF(AND(F20=1,0=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,6)))) else 0.0)

def f56(E, r, c):
    # Ô C35
    # = IF(AND(C21=1,1=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,3)))) else 0.0)

def f57(E, r, c):
    # Ô D35
    # = IF(AND(D21=1,1=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,4)))) else 0.0)

def f58(E, r, c):
    # Ô E35
    # = IF(AND(E21=1,1=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,5)))) else 0.0)

def f59(E, r, c):
    # Ô F35
    # = IF(AND(F21=1,1=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,6)))) else 0.0)

def f60(E, r, c):
    # Ô C36
    # = IF(AND(C22=1,2=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,3)))) else 0.0)

def f61(E, r, c):
    # Ô D36
    # = IF(AND(D22=1,2=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,4)))) else 0.0)

def f62(E, r, c):
    # Ô E36
    # = IF(AND(E22=1,2=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,5)))) else 0.0)

def f63(E, r, c):
    # Ô F36
    # = IF(AND(F22=1,2=0,PP1_DINH_TINH!$C$27<0,TRUNG_GIAN!$C$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,3),0.0), X.eq(E.v(23,15,3),E.v(23,5,6)))) else 0.0)

def f64(E, r, c):
    # Ô C37
    # = IF(AND(C23=1,0=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,3)))) else 0.0)

def f65(E, r, c):
    # Ô D37
    # = IF(AND(D23=1,0=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,4)))) else 0.0)

def f66(E, r, c):
    # Ô E37
    # = IF(AND(E23=1,0=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,5)))) else 0.0)

def f67(E, r, c):
    # Ô F37
    # = IF(AND(F23=1,0=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,6)))) else 0.0)

def f68(E, r, c):
    # Ô C38
    # = IF(AND(C24=1,1=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,3)))) else 0.0)

def f69(E, r, c):
    # Ô D38
    # = IF(AND(D24=1,1=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,4)))) else 0.0)

def f70(E, r, c):
    # Ô E38
    # = IF(AND(E24=1,1=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,5)))) else 0.0)

def f71(E, r, c):
    # Ô F38
    # = IF(AND(F24=1,1=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,6)))) else 0.0)

def f72(E, r, c):
    # Ô C39
    # = IF(AND(C25=1,2=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,3)))) else 0.0)

def f73(E, r, c):
    # Ô D39
    # = IF(AND(D25=1,2=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,4)))) else 0.0)

def f74(E, r, c):
    # Ô E39
    # = IF(AND(E25=1,2=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,5)))) else 0.0)

def f75(E, r, c):
    # Ô F39
    # = IF(AND(F25=1,2=0,PP1_DINH_TINH!$D$27<0,TRUNG_GIAN!$D$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,4),0.0), X.eq(E.v(23,15,4),E.v(23,5,6)))) else 0.0)

def f76(E, r, c):
    # Ô C40
    # = IF(AND(C26=1,0=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,3)))) else 0.0)

def f77(E, r, c):
    # Ô D40
    # = IF(AND(D26=1,0=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,4)))) else 0.0)

def f78(E, r, c):
    # Ô E40
    # = IF(AND(E26=1,0=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,5)))) else 0.0)

def f79(E, r, c):
    # Ô F40
    # = IF(AND(F26=1,0=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,6)))) else 0.0)

def f80(E, r, c):
    # Ô C41
    # = IF(AND(C27=1,1=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,3)))) else 0.0)

def f81(E, r, c):
    # Ô D41
    # = IF(AND(D27=1,1=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,4)))) else 0.0)

def f82(E, r, c):
    # Ô E41
    # = IF(AND(E27=1,1=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,5)))) else 0.0)

def f83(E, r, c):
    # Ô F41
    # = IF(AND(F27=1,1=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,6)))) else 0.0)

def f84(E, r, c):
    # Ô C42
    # = IF(AND(C28=1,2=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,3)))) else 0.0)

def f85(E, r, c):
    # Ô D42
    # = IF(AND(D28=1,2=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,4)))) else 0.0)

def f86(E, r, c):
    # Ô E42
    # = IF(AND(E28=1,2=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,5)))) else 0.0)

def f87(E, r, c):
    # Ô F42
    # = IF(AND(F28=1,2=0,PP1_DINH_TINH!$E$27<0,TRUNG_GIAN!$E$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,5),0.0), X.eq(E.v(23,15,5),E.v(23,5,6)))) else 0.0)

def f88(E, r, c):
    # Ô C43
    # = IF(AND(C29=1,0=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,3)))) else 0.0)

def f89(E, r, c):
    # Ô D43
    # = IF(AND(D29=1,0=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,4)))) else 0.0)

def f90(E, r, c):
    # Ô E43
    # = IF(AND(E29=1,0=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,5)))) else 0.0)

def f91(E, r, c):
    # Ô F43
    # = IF(AND(F29=1,0=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(0.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,6)))) else 0.0)

def f92(E, r, c):
    # Ô C44
    # = IF(AND(C30=1,1=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,3)))) else 0.0)

def f93(E, r, c):
    # Ô D44
    # = IF(AND(D30=1,1=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,4)))) else 0.0)

def f94(E, r, c):
    # Ô E44
    # = IF(AND(E30=1,1=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,5)))) else 0.0)

def f95(E, r, c):
    # Ô F44
    # = IF(AND(F30=1,1=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(1.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,6)))) else 0.0)

def f96(E, r, c):
    # Ô C45
    # = IF(AND(C31=1,2=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$C$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,3)))) else 0.0)

def f97(E, r, c):
    # Ô D45
    # = IF(AND(D31=1,2=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$D$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,4)))) else 0.0)

def f98(E, r, c):
    # Ô E45
    # = IF(AND(E31=1,2=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$E$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,5)))) else 0.0)

def f99(E, r, c):
    # Ô F45
    # = IF(AND(F31=1,2=0,PP1_DINH_TINH!$F$27<0,TRUNG_GIAN!$F$15=TRUNG_GIAN!$F$5),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-14,c),1.0), X.eq(2.0,0.0), X.lt(E.v(9,27,6),0.0), X.eq(E.v(23,15,6),E.v(23,5,6)))) else 0.0)

def f100(E, r, c):
    # Ô C48, D48, E48, F48, C49, D49, E49, F49, C50, D50, E50, F50
    # = IF(AND(C6=1,PP1_DINH_TINH!$C$28=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-42,c),1.0), X.eq(E.v(9,28,3),1.0))) else 0.0)

def f101(E, r, c):
    # Ô C51, D51, E51, F51, C52, D52, E52, F52, C53, D53, E53, F53
    # = IF(AND(C9=1,PP1_DINH_TINH!$D$28=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-42,c),1.0), X.eq(E.v(9,28,4),1.0))) else 0.0)

def f102(E, r, c):
    # Ô C54, D54, E54, F54, C55, D55, E55, F55, C56, D56, E56, F56
    # = IF(AND(C12=1,PP1_DINH_TINH!$E$28=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-42,c),1.0), X.eq(E.v(9,28,5),1.0))) else 0.0)

def f103(E, r, c):
    # Ô C57, D57, E57, F57, C58, D58, E58, F58, C59, D59, E59, F59
    # = IF(AND(C15=1,PP1_DINH_TINH!$F$28=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(11,r-42,c),1.0), X.eq(E.v(9,28,6),1.0))) else 0.0)

def f104(E, r, c):
    # Ô C64
    # = TRUNG_GIAN!$C$10
    return E.v(23,10,3)

def f105(E, r, c):
    # Ô D64
    # = TRUNG_GIAN!$D$10
    return E.v(23,10,4)

def f106(E, r, c):
    # Ô E64
    # = TRUNG_GIAN!$E$10
    return E.v(23,10,5)

def f107(E, r, c):
    # Ô F64
    # = TRUNG_GIAN!$F$10
    return E.v(23,10,6)

def f108(E, r, c):
    # Ô C65, D65, E65, F65
    # = SUM(C6:C17)
    return X.SUM(E.rg(11,r-59,c,r-48,c))

def f109(E, r, c):
    # Ô C66, D66, E66, F66
    # = SUM(C20:C31)
    return X.SUM(E.rg(11,r-46,c,r-35,c))

def f110(E, r, c):
    # Ô C67, D67, E67, F67
    # = SUM(C34:C45)
    return X.SUM(E.rg(11,r-33,c,r-22,c))

def f111(E, r, c):
    # Ô C68, D68, E68, F68
    # = SUM(C48:C59)
    return X.SUM(E.rg(11,r-20,c,r-9,c))

def f112(E, r, c):
    # Ô C69, D69, E69, F69
    # = IF(C65=0,"Không có gốc",IF(C66=0,IF(C68=C65,"Gốc bị hợp (đứng độc lập)","Gốc dễ bị nhổ"),IF(C67>0,"Gốc bền chặt","Có gốc (kém hơn)")))
    return ('Không có gốc' if X.truthy(X.eq(E.v(11,r-4,c),0.0)) else (('Gốc bị hợp (đứng độc lập)' if X.truthy(X.eq(E.v(11,r-1,c),E.v(11,r-4,c))) else 'Gốc dễ bị nhổ') if X.truthy(X.eq(E.v(11,r-3,c),0.0)) else ('Gốc bền chặt' if X.truthy(X.gt(E.v(11,r-2,c),0.0)) else 'Có gốc (kém hơn)')))

def f113(E, r, c):
    # Ô C70
    # = IF(C66=0,"",TRIM(IF((SUM(C20:C22))>0,"thông gốc tự vượng – cùng trụ: "&TRIM(IF(C20=1,TRUNG_GIAN!$C$11&" (năm, bản khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa
    #    "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1))&") ","")&IF(C21=1,TRUNG_GIAN!$C$11&" (năm, trung khí "&IF(PP1_DINH_TINH!$C
    #   $27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1))&") ","")&IF(C22=1,TRUNG_GIAN!$C$11&" (năm, dư khí "&IF(PP1_DINH
    #   _TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1))&") ",""))&"; ","")&IF((SUM(C23:C25))>0,"gốc ở trụ kề: "
    #   &TRIM(IF(C23=1,TRUNG_GIAN!$D$11&" (tháng, bản khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$D
    #   $15+1))&") ","")&IF(C24=1,TRUNG_GIAN!$D$11&" (tháng, trung khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,
    #   TRUNG_GIAN!$D$17+1))&") ","")&IF(C25=1,TRUNG_GIAN!$D$11&" (tháng, dư khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDE
    #   X(CAN_TEN,TRUNG_GIAN!$D$19+1))&") ",""))&"; ","")&IF((SUM(C26:C28)+SUM(C29:C31))>0,"gốc ở trụ cách xa (lực yếu): "&TRIM(IF(C26=1,TRUNG_GIAN!$E$11&" (n
    #   gày, bản khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1))&") ","")&IF(C27=1,TRUNG_GIAN!
    #   $E$11&" (ngày, trung khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1))&") ","")&IF(C28=1
    #   ,TRUNG_GIAN!$E$11&" (ngày, dư khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1))&") ","")
    #   &IF(C29=1,TRUNG_GIAN!$F$11&" (giờ, bản khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1))
    #   &") ","")&IF(C30=1,TRUNG_GIAN!$F$11&" (giờ, trung khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIA
    #   N!$F$17+1))&") ","")&IF(C31=1,TRUNG_GIAN!$F$11&" (giờ, dư khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,T
    #   RUNG_GIAN!$F$19+1))&") ",""))&"; ","")))
    return ('' if X.truthy(X.eq(E.v(11,r-4,c),0.0)) else X.TRIM(X.cat(X.cat((X.cat(X.cat('thông gốc tự vượng – cùng trụ: ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-50,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-49,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-48,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-50,c,r-48,c)),0.0)) else ''),(X.cat(X.cat('gốc ở trụ kề: ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-47,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-46,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-45,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-47,c,r-45,c)),0.0)) else '')),(X.cat(X.cat('gốc ở trụ cách xa (lực yếu): ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-44,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-43,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-42,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-41,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-40,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-39,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.add(X.SUM(E.rg(11,r-44,c,r-42,c)),X.SUM(E.rg(11,r-41,c,r-39,c))),0.0)) else ''))))

def f114(E, r, c):
    # Ô D70
    # = IF(D66=0,"",TRIM(IF((SUM(D23:D25))>0,"thông gốc tự vượng – cùng trụ: "&TRIM(IF(D23=1,TRUNG_GIAN!$D$11&" (tháng, bản khí "&IF(PP1_DINH_TINH!$D$27>=0,"h
    #   óa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1))&") ","")&IF(D24=1,TRUNG_GIAN!$D$11&" (tháng, trung khí "&IF(PP1_DINH_TIN
    #   H!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1))&") ","")&IF(D25=1,TRUNG_GIAN!$D$11&" (tháng, dư khí "&IF(PP
    #   1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1))&") ",""))&"; ","")&IF((SUM(D20:D22)+SUM(D26:D28))
    #   >0,"gốc ở trụ kề: "&TRIM(IF(D20=1,TRUNG_GIAN!$C$11&" (năm, bản khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_
    #   TEN,TRUNG_GIAN!$C$15+1))&") ","")&IF(D21=1,TRUNG_GIAN!$C$11&" (năm, trung khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1)
    #   ,INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1))&") ","")&IF(D22=1,TRUNG_GIAN!$C$11&" (năm, dư khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!
    #   $C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1))&") ","")&IF(D26=1,TRUNG_GIAN!$E$11&" (ngày, bản khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_
    #   DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1))&") ","")&IF(D27=1,TRUNG_GIAN!$E$11&" (ngày, trung khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(H
    #   ANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1))&") ","")&IF(D28=1,TRUNG_GIAN!$E$11&" (ngày, dư khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa 
    #   "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1))&") ",""))&"; ","")&IF((SUM(D29:D31))>0,"gốc ở trụ cách xa (lực yếu): "&TRIM
    #   (IF(D29=1,TRUNG_GIAN!$F$11&" (giờ, bản khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1))
    #   &") ","")&IF(D30=1,TRUNG_GIAN!$F$11&" (giờ, trung khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIA
    #   N!$F$17+1))&") ","")&IF(D31=1,TRUNG_GIAN!$F$11&" (giờ, dư khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,T
    #   RUNG_GIAN!$F$19+1))&") ",""))&"; ","")))
    return ('' if X.truthy(X.eq(E.v(11,r-4,c),0.0)) else X.TRIM(X.cat(X.cat((X.cat(X.cat('thông gốc tự vượng – cùng trụ: ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-47,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-46,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-45,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-47,c,r-45,c)),0.0)) else ''),(X.cat(X.cat('gốc ở trụ kề: ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-50,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-49,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-48,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-44,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-43,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-42,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.add(X.SUM(E.rg(11,r-50,c,r-48,c)),X.SUM(E.rg(11,r-44,c,r-42,c))),0.0)) else '')),(X.cat(X.cat('gốc ở trụ cách xa (lực yếu): ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-41,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-40,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-39,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-41,c,r-39,c)),0.0)) else ''))))

def f115(E, r, c):
    # Ô E70
    # = IF(E66=0,"",TRIM(IF((SUM(E26:E28))>0,"thông gốc tự vượng – cùng trụ: "&TRIM(IF(E26=1,TRUNG_GIAN!$E$11&" (ngày, bản khí "&IF(PP1_DINH_TINH!$E$27>=0,"hó
    #   a "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1))&") ","")&IF(E27=1,TRUNG_GIAN!$E$11&" (ngày, trung khí "&IF(PP1_DINH_TINH!
    #   $E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1))&") ","")&IF(E28=1,TRUNG_GIAN!$E$11&" (ngày, dư khí "&IF(PP1_D
    #   INH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1))&") ",""))&"; ","")&IF((SUM(E23:E25)+SUM(E29:E31))>0,
    #   "gốc ở trụ kề: "&TRIM(IF(E23=1,TRUNG_GIAN!$D$11&" (tháng, bản khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_T
    #   EN,TRUNG_GIAN!$D$15+1))&") ","")&IF(E24=1,TRUNG_GIAN!$D$11&" (tháng, trung khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1
    #   ),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1))&") ","")&IF(E25=1,TRUNG_GIAN!$D$11&" (tháng, dư khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TI
    #   NH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1))&") ","")&IF(E29=1,TRUNG_GIAN!$F$11&" (giờ, bản khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP
    #   1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1))&") ","")&IF(E30=1,TRUNG_GIAN!$F$11&" (giờ, trung khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa "&INDEX(
    #   HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1))&") ","")&IF(E31=1,TRUNG_GIAN!$F$11&" (giờ, dư khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa 
    #   "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1))&") ",""))&"; ","")&IF((SUM(E20:E22))>0,"gốc ở trụ cách xa (lực yếu): "&TRIM
    #   (IF(E20=1,TRUNG_GIAN!$C$11&" (năm, bản khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1))
    #   &") ","")&IF(E21=1,TRUNG_GIAN!$C$11&" (năm, trung khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIA
    #   N!$C$17+1))&") ","")&IF(E22=1,TRUNG_GIAN!$C$11&" (năm, dư khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,T
    #   RUNG_GIAN!$C$19+1))&") ",""))&"; ","")))
    return ('' if X.truthy(X.eq(E.v(11,r-4,c),0.0)) else X.TRIM(X.cat(X.cat((X.cat(X.cat('thông gốc tự vượng – cùng trụ: ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-44,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-43,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-42,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-44,c,r-42,c)),0.0)) else ''),(X.cat(X.cat('gốc ở trụ kề: ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-47,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-46,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-45,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-41,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-40,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-39,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.add(X.SUM(E.rg(11,r-47,c,r-45,c)),X.SUM(E.rg(11,r-41,c,r-39,c))),0.0)) else '')),(X.cat(X.cat('gốc ở trụ cách xa (lực yếu): ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-50,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-49,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-48,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-50,c,r-48,c)),0.0)) else ''))))

def f116(E, r, c):
    # Ô F70
    # = IF(F66=0,"",TRIM(IF((SUM(F29:F31))>0,"thông gốc tự vượng – cùng trụ: "&TRIM(IF(F29=1,TRUNG_GIAN!$F$11&" (giờ, bản khí "&IF(PP1_DINH_TINH!$F$27>=0,"hóa
    #    "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1))&") ","")&IF(F30=1,TRUNG_GIAN!$F$11&" (giờ, trung khí "&IF(PP1_DINH_TINH!$F
    #   $27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1))&") ","")&IF(F31=1,TRUNG_GIAN!$F$11&" (giờ, dư khí "&IF(PP1_DINH
    #   _TINH!$F$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1))&") ",""))&"; ","")&IF((SUM(F26:F28))>0,"gốc ở trụ kề: "
    #   &TRIM(IF(F26=1,TRUNG_GIAN!$E$11&" (ngày, bản khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$E$
    #   15+1))&") ","")&IF(F27=1,TRUNG_GIAN!$E$11&" (ngày, trung khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(CAN_TEN,TR
    #   UNG_GIAN!$E$17+1))&") ","")&IF(F28=1,TRUNG_GIAN!$E$11&" (ngày, dư khí "&IF(PP1_DINH_TINH!$E$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1),INDEX(C
    #   AN_TEN,TRUNG_GIAN!$E$19+1))&") ",""))&"; ","")&IF((SUM(F20:F22)+SUM(F23:F25))>0,"gốc ở trụ cách xa (lực yếu): "&TRIM(IF(F20=1,TRUNG_GIAN!$C$11&" (năm,
    #    bản khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1))&") ","")&IF(F21=1,TRUNG_GIAN!$C$1
    #   1&" (năm, trung khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1))&") ","")&IF(F22=1,TRUN
    #   G_GIAN!$C$11&" (năm, dư khí "&IF(PP1_DINH_TINH!$C$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1))&") ","")&IF(F2
    #   3=1,TRUNG_GIAN!$D$11&" (tháng, bản khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1))&") 
    #   ","")&IF(F24=1,TRUNG_GIAN!$D$11&" (tháng, trung khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,TRUNG_GIAN!
    #   $D$17+1))&") ","")&IF(F25=1,TRUNG_GIAN!$D$11&" (tháng, dư khí "&IF(PP1_DINH_TINH!$D$27>=0,"hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1),INDEX(CAN_TEN,T
    #   RUNG_GIAN!$D$19+1))&") ",""))&"; ","")))
    return ('' if X.truthy(X.eq(E.v(11,r-4,c),0.0)) else X.TRIM(X.cat(X.cat((X.cat(X.cat('thông gốc tự vượng – cùng trụ: ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-41,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-40,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,6),' (giờ, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))) if X.truthy(X.ge(E.v(9,27,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-39,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-41,c,r-39,c)),0.0)) else ''),(X.cat(X.cat('gốc ở trụ kề: ',X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-44,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-43,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,5),' (ngày, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))) if X.truthy(X.ge(E.v(9,27,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-42,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.SUM(E.rg(11,r-44,c,r-42,c)),0.0)) else '')),(X.cat(X.cat('gốc ở trụ cách xa (lực yếu): ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-50,c),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-49,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,3),' (năm, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))) if X.truthy(X.ge(E.v(9,27,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-48,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, bản khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-47,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, trung khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-46,c),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(23,11,4),' (tháng, dư khí '),(X.cat('hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))) if X.truthy(X.ge(E.v(9,27,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)))),') ') if X.truthy(X.eq(E.v(11,r-45,c),1.0)) else '')))),'; ') if X.truthy(X.gt(X.add(X.SUM(E.rg(11,r-50,c,r-48,c)),X.SUM(E.rg(11,r-47,c,r-45,c))),0.0)) else ''))))

def f117(E, r, c):
    # Ô C71, D71, E71, F71
    # = IF(OR(C69="Gốc bền chặt",C69="Có gốc (kém hơn)"),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(11,r-2,c),'Gốc bền chặt'), X.eq(E.v(11,r-2,c),'Có gốc (kém hơn)'))) else 0.0)

def f118(E, r, c):
    # Ô C72
    # = IF(C66=0,"",IF(AND(C67>0,(SUM(C34:C36)+SUM(C37:C39))=0),"Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn",IF((SUM(C20:C22)+SUM(C23:C25))=0,"Mọi gốc đ
    #   ều ở trụ cách xa → lực yếu","")))
    return ('' if X.truthy(X.eq(E.v(11,r-6,c),0.0)) else ('Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn' if X.truthy(X.AND(X.gt(E.v(11,r-5,c),0.0), X.eq(X.add(X.SUM(E.rg(11,r-38,c,r-36,c)),X.SUM(E.rg(11,r-35,c,r-33,c))),0.0))) else ('Mọi gốc đều ở trụ cách xa → lực yếu' if X.truthy(X.eq(X.add(X.SUM(E.rg(11,r-52,c,r-50,c)),X.SUM(E.rg(11,r-49,c,r-47,c))),0.0)) else '')))

def f119(E, r, c):
    # Ô D72
    # = IF(D66=0,"",IF(AND(D67>0,(SUM(D34:D36)+SUM(D37:D39)+SUM(D40:D42))=0),"Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn",IF((SUM(D20:D22)+SUM(D23:D25)+
    #   SUM(D26:D28))=0,"Mọi gốc đều ở trụ cách xa → lực yếu","")))
    return ('' if X.truthy(X.eq(E.v(11,r-6,c),0.0)) else ('Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn' if X.truthy(X.AND(X.gt(E.v(11,r-5,c),0.0), X.eq(X.add(X.add(X.SUM(E.rg(11,r-38,c,r-36,c)),X.SUM(E.rg(11,r-35,c,r-33,c))),X.SUM(E.rg(11,r-32,c,r-30,c))),0.0))) else ('Mọi gốc đều ở trụ cách xa → lực yếu' if X.truthy(X.eq(X.add(X.add(X.SUM(E.rg(11,r-52,c,r-50,c)),X.SUM(E.rg(11,r-49,c,r-47,c))),X.SUM(E.rg(11,r-46,c,r-44,c))),0.0)) else '')))

def f120(E, r, c):
    # Ô E72
    # = IF(E66=0,"",IF(AND(E67>0,(SUM(E37:E39)+SUM(E40:E42)+SUM(E43:E45))=0),"Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn",IF((SUM(E23:E25)+SUM(E26:E28)+
    #   SUM(E29:E31))=0,"Mọi gốc đều ở trụ cách xa → lực yếu","")))
    return ('' if X.truthy(X.eq(E.v(11,r-6,c),0.0)) else ('Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn' if X.truthy(X.AND(X.gt(E.v(11,r-5,c),0.0), X.eq(X.add(X.add(X.SUM(E.rg(11,r-35,c,r-33,c)),X.SUM(E.rg(11,r-32,c,r-30,c))),X.SUM(E.rg(11,r-29,c,r-27,c))),0.0))) else ('Mọi gốc đều ở trụ cách xa → lực yếu' if X.truthy(X.eq(X.add(X.add(X.SUM(E.rg(11,r-49,c,r-47,c)),X.SUM(E.rg(11,r-46,c,r-44,c))),X.SUM(E.rg(11,r-43,c,r-41,c))),0.0)) else '')))

def f121(E, r, c):
    # Ô F72
    # = IF(F66=0,"",IF(AND(F67>0,(SUM(F40:F42)+SUM(F43:F45))=0),"Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn",IF((SUM(F26:F28)+SUM(F29:F31))=0,"Mọi gốc đ
    #   ều ở trụ cách xa → lực yếu","")))
    return ('' if X.truthy(X.eq(E.v(11,r-6,c),0.0)) else ('Gốc bền chặt chỉ nằm ở trụ cách xa → lực yếu hơn' if X.truthy(X.AND(X.gt(E.v(11,r-5,c),0.0), X.eq(X.add(X.SUM(E.rg(11,r-32,c,r-30,c)),X.SUM(E.rg(11,r-29,c,r-27,c))),0.0))) else ('Mọi gốc đều ở trụ cách xa → lực yếu' if X.truthy(X.eq(X.add(X.SUM(E.rg(11,r-46,c,r-44,c)),X.SUM(E.rg(11,r-43,c,r-41,c))),0.0)) else '')))

def f122(E, r, c):
    # Ô B76
    # = TRUNG_GIAN!$C$11&" (Năm) – bản khí"
    return X.cat(E.v(23,11,3),' (Năm) – bản khí')

def f123(E, r, c):
    # Ô C76
    # = IF(TRUNG_GIAN!$C$15<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)))

def f124(E, r, c):
    # Ô D76
    # = IF(TRUNG_GIAN!$C$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,3),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,3),1.0)))

def f125(E, r, c):
    # Ô E76
    # = IF(TRUNG_GIAN!$C$15<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$15,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$15,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$C$15,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$15,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,15,3),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,15,3))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,15,3))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,3))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,15,3))) else '')))

def f126(E, r, c):
    # Ô F76
    # = IF(TRUNG_GIAN!$C$15<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$15)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$15)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$15))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$C$15,IF(E20=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$15)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$15)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$15)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$15)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$15)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$15)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$C$15,IF(E20=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,15,3),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,3))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,15,3)),X.eq(E.v(23,5,4),E.v(23,15,3))),X.eq(E.v(23,5,6),E.v(23,15,3))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,3)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,15,3)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,3)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,3)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,15,3)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,3)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,3))) else ''))))

def f127(E, r, c):
    # Ô B77
    # = TRUNG_GIAN!$C$11&" (Năm) – trung khí"
    return X.cat(E.v(23,11,3),' (Năm) – trung khí')

def f128(E, r, c):
    # Ô C77
    # = IF(TRUNG_GIAN!$C$17<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)))

def f129(E, r, c):
    # Ô D77
    # = IF(TRUNG_GIAN!$C$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,3),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,3),1.0)))

def f130(E, r, c):
    # Ô E77
    # = IF(TRUNG_GIAN!$C$17<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$17,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$17,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$C$17,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$17,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,17,3),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,17,3))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,17,3))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,3))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,17,3))) else '')))

def f131(E, r, c):
    # Ô F77
    # = IF(TRUNG_GIAN!$C$17<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$17)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$17)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$17))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$C$17,IF(E21=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$17)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$17)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$17)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$17)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$17)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$17)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$C$17,IF(E21=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,17,3),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,3))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,17,3)),X.eq(E.v(23,5,4),E.v(23,17,3))),X.eq(E.v(23,5,6),E.v(23,17,3))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,3)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,17,3)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,3)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,3)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,17,3)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,3)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,3))) else ''))))

def f132(E, r, c):
    # Ô B78
    # = TRUNG_GIAN!$C$11&" (Năm) – dư khí"
    return X.cat(E.v(23,11,3),' (Năm) – dư khí')

def f133(E, r, c):
    # Ô C78
    # = IF(TRUNG_GIAN!$C$19<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,3),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)))

def f134(E, r, c):
    # Ô D78
    # = IF(TRUNG_GIAN!$C$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,3),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,3),1.0)))

def f135(E, r, c):
    # Ô E78
    # = IF(TRUNG_GIAN!$C$19<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$19,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$19,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$C$19,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$19,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,19,3),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,19,3))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,19,3))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,3))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,19,3))) else '')))

def f136(E, r, c):
    # Ô F78
    # = IF(TRUNG_GIAN!$C$19<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$19)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$19)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$19))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$C$19,IF(E22=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$19)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$19)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$19)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$C$19)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$C$19)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$C$19)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$C$19,IF(E22=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,19,3),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,3))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,19,3)),X.eq(E.v(23,5,4),E.v(23,19,3))),X.eq(E.v(23,5,6),E.v(23,19,3))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,3)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,19,3)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,3)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,3)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,19,3)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,3)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,3))) else ''))))

def f137(E, r, c):
    # Ô B79
    # = TRUNG_GIAN!$D$11&" (Tháng) – bản khí"
    return X.cat(E.v(23,11,4),' (Tháng) – bản khí')

def f138(E, r, c):
    # Ô C79
    # = IF(TRUNG_GIAN!$D$15<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)))

def f139(E, r, c):
    # Ô D79
    # = IF(TRUNG_GIAN!$D$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,4),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,4),1.0)))

def f140(E, r, c):
    # Ô E79
    # = IF(TRUNG_GIAN!$D$15<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$15,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$15,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$D$15,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$15,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,15,4),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,15,4))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,15,4))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,4))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,15,4))) else '')))

def f141(E, r, c):
    # Ô F79
    # = IF(TRUNG_GIAN!$D$15<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$15)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$15)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$15))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$D$15,IF(E23=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$15)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$15)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$15)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$15)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$15)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$15)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$D$15,IF(E23=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,15,4),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,4))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,15,4)),X.eq(E.v(23,5,4),E.v(23,15,4))),X.eq(E.v(23,5,6),E.v(23,15,4))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,4)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,15,4)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,4)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,4)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,15,4)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,4)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,4))) else ''))))

def f142(E, r, c):
    # Ô B80
    # = TRUNG_GIAN!$D$11&" (Tháng) – trung khí"
    return X.cat(E.v(23,11,4),' (Tháng) – trung khí')

def f143(E, r, c):
    # Ô C80
    # = IF(TRUNG_GIAN!$D$17<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)))

def f144(E, r, c):
    # Ô D80
    # = IF(TRUNG_GIAN!$D$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,4),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,4),1.0)))

def f145(E, r, c):
    # Ô E80
    # = IF(TRUNG_GIAN!$D$17<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$17,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$17,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$D$17,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$17,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,17,4),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,17,4))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,17,4))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,4))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,17,4))) else '')))

def f146(E, r, c):
    # Ô F80
    # = IF(TRUNG_GIAN!$D$17<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$17)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$17)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$17))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$D$17,IF(E24=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$17)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$17)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$17)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$17)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$17)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$17)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$D$17,IF(E24=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,17,4),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,4))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,17,4)),X.eq(E.v(23,5,4),E.v(23,17,4))),X.eq(E.v(23,5,6),E.v(23,17,4))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,4)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,17,4)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,4)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,4)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,17,4)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,4)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,4))) else ''))))

def f147(E, r, c):
    # Ô B81
    # = TRUNG_GIAN!$D$11&" (Tháng) – dư khí"
    return X.cat(E.v(23,11,4),' (Tháng) – dư khí')

def f148(E, r, c):
    # Ô C81
    # = IF(TRUNG_GIAN!$D$19<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,4),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)))

def f149(E, r, c):
    # Ô D81
    # = IF(TRUNG_GIAN!$D$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,4),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,4),1.0)))

def f150(E, r, c):
    # Ô E81
    # = IF(TRUNG_GIAN!$D$19<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$19,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$19,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$D$19,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$19,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,19,4),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,19,4))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,19,4))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,4))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,19,4))) else '')))

def f151(E, r, c):
    # Ô F81
    # = IF(TRUNG_GIAN!$D$19<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$19)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$19)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$19))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$D$19,IF(E25=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$19)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$19)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$19)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$D$19)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$D$19)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$D$19)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$D$19,IF(E25=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,19,4),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,4))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,19,4)),X.eq(E.v(23,5,4),E.v(23,19,4))),X.eq(E.v(23,5,6),E.v(23,19,4))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,4)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,19,4)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,4)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,4)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,19,4)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,4)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,4))) else ''))))

def f152(E, r, c):
    # Ô B82
    # = TRUNG_GIAN!$E$11&" (Ngày) – bản khí"
    return X.cat(E.v(23,11,5),' (Ngày) – bản khí')

def f153(E, r, c):
    # Ô C82
    # = IF(TRUNG_GIAN!$E$15<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)))

def f154(E, r, c):
    # Ô D82
    # = IF(TRUNG_GIAN!$E$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,5),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,5),1.0)))

def f155(E, r, c):
    # Ô E82
    # = IF(TRUNG_GIAN!$E$15<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$15,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$15,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$E$15,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$15,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,15,5),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,15,5))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,15,5))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,5))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,15,5))) else '')))

def f156(E, r, c):
    # Ô F82
    # = IF(TRUNG_GIAN!$E$15<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$15)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$15)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$15))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$E$15,IF(E26=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$15)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$15)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$15)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$15)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$15)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$15)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$E$15,IF(E26=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,15,5),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,5))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,15,5)),X.eq(E.v(23,5,4),E.v(23,15,5))),X.eq(E.v(23,5,6),E.v(23,15,5))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,5)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,15,5)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,5)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,5)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,15,5)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,5)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,5))) else ''))))

def f157(E, r, c):
    # Ô B83
    # = TRUNG_GIAN!$E$11&" (Ngày) – trung khí"
    return X.cat(E.v(23,11,5),' (Ngày) – trung khí')

def f158(E, r, c):
    # Ô C83
    # = IF(TRUNG_GIAN!$E$17<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)))

def f159(E, r, c):
    # Ô D83
    # = IF(TRUNG_GIAN!$E$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,5),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,5),1.0)))

def f160(E, r, c):
    # Ô E83
    # = IF(TRUNG_GIAN!$E$17<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$17,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$17,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$E$17,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$17,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,17,5),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,17,5))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,17,5))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,5))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,17,5))) else '')))

def f161(E, r, c):
    # Ô F83
    # = IF(TRUNG_GIAN!$E$17<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$17)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$17)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$17))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$E$17,IF(E27=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$17)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$17)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$17)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$17)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$17)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$17)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$E$17,IF(E27=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,17,5),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,5))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,17,5)),X.eq(E.v(23,5,4),E.v(23,17,5))),X.eq(E.v(23,5,6),E.v(23,17,5))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,5)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,17,5)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,5)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,5)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,17,5)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,5)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,5))) else ''))))

def f162(E, r, c):
    # Ô B84
    # = TRUNG_GIAN!$E$11&" (Ngày) – dư khí"
    return X.cat(E.v(23,11,5),' (Ngày) – dư khí')

def f163(E, r, c):
    # Ô C84
    # = IF(TRUNG_GIAN!$E$19<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,5),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)))

def f164(E, r, c):
    # Ô D84
    # = IF(TRUNG_GIAN!$E$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,5),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,5),1.0)))

def f165(E, r, c):
    # Ô E84
    # = IF(TRUNG_GIAN!$E$19<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$19,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$19,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$E$19,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$19,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,19,5),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,19,5))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,19,5))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,5))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,19,5))) else '')))

def f166(E, r, c):
    # Ô F84
    # = IF(TRUNG_GIAN!$E$19<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$19)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$19)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$19))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$E$19,IF(E28=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$19)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$19)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$19)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$19)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$E$19)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$E$19)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$E$19,IF(E28=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,19,5),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,5))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,19,5)),X.eq(E.v(23,5,4),E.v(23,19,5))),X.eq(E.v(23,5,6),E.v(23,19,5))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,5)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,19,5)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,5)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,5)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,19,5)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,5)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,5))) else ''))))

def f167(E, r, c):
    # Ô B85
    # = TRUNG_GIAN!$F$11&" (Giờ) – bản khí"
    return X.cat(E.v(23,11,6),' (Giờ) – bản khí')

def f168(E, r, c):
    # Ô C85
    # = IF(TRUNG_GIAN!$F$15<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)))

def f169(E, r, c):
    # Ô D85
    # = IF(TRUNG_GIAN!$F$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$15+1))
    return ('' if X.truthy(X.lt(E.v(23,15,6),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,6),1.0)))

def f170(E, r, c):
    # Ô E85
    # = IF(TRUNG_GIAN!$F$15<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$15,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$15,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$F$15,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$15,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,15,6),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,15,6))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,15,6))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,6))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,15,6))) else '')))

def f171(E, r, c):
    # Ô F85
    # = IF(TRUNG_GIAN!$F$15<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$15)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$15)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$15))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$F$15,IF(E29=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$15)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$15)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$15)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$15)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$15)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$15)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$F$15,IF(E29=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,15,6),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,6))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,15,6)),X.eq(E.v(23,5,4),E.v(23,15,6))),X.eq(E.v(23,5,6),E.v(23,15,6))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,6)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,15,6)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,6)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,15,6)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,15,6)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,15,6)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,15,6))) else ''))))

def f172(E, r, c):
    # Ô B86
    # = TRUNG_GIAN!$F$11&" (Giờ) – trung khí"
    return X.cat(E.v(23,11,6),' (Giờ) – trung khí')

def f173(E, r, c):
    # Ô C86
    # = IF(TRUNG_GIAN!$F$17<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)))

def f174(E, r, c):
    # Ô D86
    # = IF(TRUNG_GIAN!$F$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$17+1))
    return ('' if X.truthy(X.lt(E.v(23,17,6),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,6),1.0)))

def f175(E, r, c):
    # Ô E86
    # = IF(TRUNG_GIAN!$F$17<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$17,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$17,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$F$17,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$17,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,17,6),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,17,6))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,17,6))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,6))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,17,6))) else '')))

def f176(E, r, c):
    # Ô F86
    # = IF(TRUNG_GIAN!$F$17<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$17)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$17)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$17))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$F$17,IF(E30=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$17)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$17)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$17)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$17)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$17)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$17)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$F$17,IF(E30=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,17,6),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,6))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,17,6)),X.eq(E.v(23,5,4),E.v(23,17,6))),X.eq(E.v(23,5,6),E.v(23,17,6))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,6)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,17,6)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,6)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,17,6)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,17,6)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,17,6)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,17,6))) else ''))))

def f177(E, r, c):
    # Ô B87
    # = TRUNG_GIAN!$F$11&" (Giờ) – dư khí"
    return X.cat(E.v(23,11,6),' (Giờ) – dư khí')

def f178(E, r, c):
    # Ô C87
    # = IF(TRUNG_GIAN!$F$19<0,"",INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,6),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)))

def f179(E, r, c):
    # Ô D87
    # = IF(TRUNG_GIAN!$F$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$19+1))
    return ('' if X.truthy(X.lt(E.v(23,19,6),0.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,6),1.0)))

def f180(E, r, c):
    # Ô E87
    # = IF(TRUNG_GIAN!$F$19<0,"",IF(TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$19,"Năm ","")&IF(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$19,"Tháng ","")&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN
    #   !$F$19,"Ngày (nhật can) ","")&IF(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$19,"Giờ ",""))
    return ('' if X.truthy(X.lt(E.v(23,19,6),0.0)) else X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(E.v(23,5,3),E.v(23,19,6))) else ''),('Tháng ' if X.truthy(X.eq(E.v(23,5,4),E.v(23,19,6))) else '')),('Ngày (nhật can) ' if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,6))) else '')),('Giờ ' if X.truthy(X.eq(E.v(23,5,6),E.v(23,19,6))) else '')))

def f181(E, r, c):
    # Ô F87
    # = IF(TRUNG_GIAN!$F$19<0,"",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$19)+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$19)+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$19))=0,IF(TRUNG_GIAN
    #   !$E$5=TRUNG_GIAN!$F$19,IF(E31=1,"Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)","Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí t
    #   an hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)"),"Tàng mà không thấu"),IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$19)
    #   *C$71+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$19)*D$71+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$19)*F$71)>0,"Thấu và có gốc",IF(((TRUNG_GIAN!$C$5=TRUNG_GIAN!$F$19)*(C$69
    #   ="Gốc dễ bị nhổ")+(TRUNG_GIAN!$D$5=TRUNG_GIAN!$F$19)*(D$69="Gốc dễ bị nhổ")+(TRUNG_GIAN!$F$5=TRUNG_GIAN!$F$19)*(F$69="Gốc dễ bị nhổ"))>0,"Lộ, gốc dễ b
    #   ị nhổ (chi bị xung)","Lộ mà không có gốc"))&IF(TRUNG_GIAN!$E$5=TRUNG_GIAN!$F$19,IF(E31=1,"; cũng là gốc của nhật chủ","; cùng can nhật chủ nhưng chi b
    #   ị xung / hợp nên không là gốc dùng được"),"")))
    return ('' if X.truthy(X.lt(E.v(23,19,6),0.0)) else ((('Cùng can với nhật chủ → là gốc của nhật chủ (xem dòng 70)' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else 'Cùng can với nhật chủ nhưng chi bị xung / hợp → tạp khí tan hoặc đứng độc lập, không còn là gốc dùng được (P1 tr.62; P2 tr.334-336; xem dòng 70)') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,6))) else 'Tàng mà không thấu') if X.truthy(X.eq(X.add(X.add(X.eq(E.v(23,5,3),E.v(23,19,6)),X.eq(E.v(23,5,4),E.v(23,19,6))),X.eq(E.v(23,5,6),E.v(23,19,6))),0.0)) else X.cat(('Thấu và có gốc' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,6)),E.v(11,71,c-3)),X.mul(X.eq(E.v(23,5,4),E.v(23,19,6)),E.v(11,71,c-2))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,6)),E.v(11,71,c))),0.0)) else ('Lộ, gốc dễ bị nhổ (chi bị xung)' if X.truthy(X.gt(X.add(X.add(X.mul(X.eq(E.v(23,5,3),E.v(23,19,6)),X.eq(E.v(11,69,c-3),'Gốc dễ bị nhổ')),X.mul(X.eq(E.v(23,5,4),E.v(23,19,6)),X.eq(E.v(11,69,c-2),'Gốc dễ bị nhổ'))),X.mul(X.eq(E.v(23,5,6),E.v(23,19,6)),X.eq(E.v(11,69,c),'Gốc dễ bị nhổ'))),0.0)) else 'Lộ mà không có gốc')),(('; cũng là gốc của nhật chủ' if X.truthy(X.eq(E.v(11,r-56,c-1),1.0)) else '; cùng can nhật chủ nhưng chi bị xung / hợp nên không là gốc dùng được') if X.truthy(X.eq(E.v(23,5,5),E.v(23,19,6))) else ''))))

def f182(E, r, c):
    # Ô C89
    # = KET_LUAN!$C$60
    return E.v(3,60,3)

CELLS = {
    'C6': (f0, 6, 3),
    'D6': (f1, 6, 4),
    'E6': (f2, 6, 5),
    'F6': (f3, 6, 6),
    'C7': (f4, 7, 3),
    'D7': (f5, 7, 4),
    'E7': (f6, 7, 5),
    'F7': (f7, 7, 6),
    'C8': (f8, 8, 3),
    'D8': (f9, 8, 4),
    'E8': (f10, 8, 5),
    'F8': (f11, 8, 6),
    'C9': (f12, 9, 3),
    'D9': (f13, 9, 4),
    'E9': (f14, 9, 5),
    'F9': (f15, 9, 6),
    'C10': (f16, 10, 3),
    'D10': (f17, 10, 4),
    'E10': (f18, 10, 5),
    'F10': (f19, 10, 6),
    'C11': (f20, 11, 3),
    'D11': (f21, 11, 4),
    'E11': (f22, 11, 5),
    'F11': (f23, 11, 6),
    'C12': (f24, 12, 3),
    'D12': (f25, 12, 4),
    'E12': (f26, 12, 5),
    'F12': (f27, 12, 6),
    'C13': (f28, 13, 3),
    'D13': (f29, 13, 4),
    'E13': (f30, 13, 5),
    'F13': (f31, 13, 6),
    'C14': (f32, 14, 3),
    'D14': (f33, 14, 4),
    'E14': (f34, 14, 5),
    'F14': (f35, 14, 6),
    'C15': (f36, 15, 3),
    'D15': (f37, 15, 4),
    'E15': (f38, 15, 5),
    'F15': (f39, 15, 6),
    'C16': (f40, 16, 3),
    'D16': (f41, 16, 4),
    'E16': (f42, 16, 5),
    'F16': (f43, 16, 6),
    'C17': (f44, 17, 3),
    'D17': (f45, 17, 4),
    'E17': (f46, 17, 5),
    'F17': (f47, 17, 6),
    'C20': (f48, 20, 3),
    'D20': (f48, 20, 4),
    'E20': (f48, 20, 5),
    'F20': (f48, 20, 6),
    'C21': (f48, 21, 3),
    'D21': (f48, 21, 4),
    'E21': (f48, 21, 5),
    'F21': (f48, 21, 6),
    'C22': (f48, 22, 3),
    'D22': (f48, 22, 4),
    'E22': (f48, 22, 5),
    'F22': (f48, 22, 6),
    'C23': (f49, 23, 3),
    'D23': (f49, 23, 4),
    'E23': (f49, 23, 5),
    'F23': (f49, 23, 6),
    'C24': (f49, 24, 3),
    'D24': (f49, 24, 4),
    'E24': (f49, 24, 5),
    'F24': (f49, 24, 6),
    'C25': (f49, 25, 3),
    'D25': (f49, 25, 4),
    'E25': (f49, 25, 5),
    'F25': (f49, 25, 6),
    'C26': (f50, 26, 3),
    'D26': (f50, 26, 4),
    'E26': (f50, 26, 5),
    'F26': (f50, 26, 6),
    'C27': (f50, 27, 3),
    'D27': (f50, 27, 4),
    'E27': (f50, 27, 5),
    'F27': (f50, 27, 6),
    'C28': (f50, 28, 3),
    'D28': (f50, 28, 4),
    'E28': (f50, 28, 5),
    'F28': (f50, 28, 6),
    'C29': (f51, 29, 3),
    'D29': (f51, 29, 4),
    'E29': (f51, 29, 5),
    'F29': (f51, 29, 6),
    'C30': (f51, 30, 3),
    'D30': (f51, 30, 4),
    'E30': (f51, 30, 5),
    'F30': (f51, 30, 6),
    'C31': (f51, 31, 3),
    'D31': (f51, 31, 4),
    'E31': (f51, 31, 5),
    'F31': (f51, 31, 6),
    'C34': (f52, 34, 3),
    'D34': (f53, 34, 4),
    'E34': (f54, 34, 5),
    'F34': (f55, 34, 6),
    'C35': (f56, 35, 3),
    'D35': (f57, 35, 4),
    'E35': (f58, 35, 5),
    'F35': (f59, 35, 6),
    'C36': (f60, 36, 3),
    'D36': (f61, 36, 4),
    'E36': (f62, 36, 5),
    'F36': (f63, 36, 6),
    'C37': (f64, 37, 3),
    'D37': (f65, 37, 4),
    'E37': (f66, 37, 5),
    'F37': (f67, 37, 6),
    'C38': (f68, 38, 3),
    'D38': (f69, 38, 4),
    'E38': (f70, 38, 5),
    'F38': (f71, 38, 6),
    'C39': (f72, 39, 3),
    'D39': (f73, 39, 4),
    'E39': (f74, 39, 5),
    'F39': (f75, 39, 6),
    'C40': (f76, 40, 3),
    'D40': (f77, 40, 4),
    'E40': (f78, 40, 5),
    'F40': (f79, 40, 6),
    'C41': (f80, 41, 3),
    'D41': (f81, 41, 4),
    'E41': (f82, 41, 5),
    'F41': (f83, 41, 6),
    'C42': (f84, 42, 3),
    'D42': (f85, 42, 4),
    'E42': (f86, 42, 5),
    'F42': (f87, 42, 6),
    'C43': (f88, 43, 3),
    'D43': (f89, 43, 4),
    'E43': (f90, 43, 5),
    'F43': (f91, 43, 6),
    'C44': (f92, 44, 3),
    'D44': (f93, 44, 4),
    'E44': (f94, 44, 5),
    'F44': (f95, 44, 6),
    'C45': (f96, 45, 3),
    'D45': (f97, 45, 4),
    'E45': (f98, 45, 5),
    'F45': (f99, 45, 6),
    'C48': (f100, 48, 3),
    'D48': (f100, 48, 4),
    'E48': (f100, 48, 5),
    'F48': (f100, 48, 6),
    'C49': (f100, 49, 3),
    'D49': (f100, 49, 4),
    'E49': (f100, 49, 5),
    'F49': (f100, 49, 6),
    'C50': (f100, 50, 3),
    'D50': (f100, 50, 4),
    'E50': (f100, 50, 5),
    'F50': (f100, 50, 6),
    'C51': (f101, 51, 3),
    'D51': (f101, 51, 4),
    'E51': (f101, 51, 5),
    'F51': (f101, 51, 6),
    'C52': (f101, 52, 3),
    'D52': (f101, 52, 4),
    'E52': (f101, 52, 5),
    'F52': (f101, 52, 6),
    'C53': (f101, 53, 3),
    'D53': (f101, 53, 4),
    'E53': (f101, 53, 5),
    'F53': (f101, 53, 6),
    'C54': (f102, 54, 3),
    'D54': (f102, 54, 4),
    'E54': (f102, 54, 5),
    'F54': (f102, 54, 6),
    'C55': (f102, 55, 3),
    'D55': (f102, 55, 4),
    'E55': (f102, 55, 5),
    'F55': (f102, 55, 6),
    'C56': (f102, 56, 3),
    'D56': (f102, 56, 4),
    'E56': (f102, 56, 5),
    'F56': (f102, 56, 6),
    'C57': (f103, 57, 3),
    'D57': (f103, 57, 4),
    'E57': (f103, 57, 5),
    'F57': (f103, 57, 6),
    'C58': (f103, 58, 3),
    'D58': (f103, 58, 4),
    'E58': (f103, 58, 5),
    'F58': (f103, 58, 6),
    'C59': (f103, 59, 3),
    'D59': (f103, 59, 4),
    'E59': (f103, 59, 5),
    'F59': (f103, 59, 6),
    'C64': (f104, 64, 3),
    'D64': (f105, 64, 4),
    'E64': (f106, 64, 5),
    'F64': (f107, 64, 6),
    'C65': (f108, 65, 3),
    'D65': (f108, 65, 4),
    'E65': (f108, 65, 5),
    'F65': (f108, 65, 6),
    'C66': (f109, 66, 3),
    'D66': (f109, 66, 4),
    'E66': (f109, 66, 5),
    'F66': (f109, 66, 6),
    'C67': (f110, 67, 3),
    'D67': (f110, 67, 4),
    'E67': (f110, 67, 5),
    'F67': (f110, 67, 6),
    'C68': (f111, 68, 3),
    'D68': (f111, 68, 4),
    'E68': (f111, 68, 5),
    'F68': (f111, 68, 6),
    'C69': (f112, 69, 3),
    'D69': (f112, 69, 4),
    'E69': (f112, 69, 5),
    'F69': (f112, 69, 6),
    'C70': (f113, 70, 3),
    'D70': (f114, 70, 4),
    'E70': (f115, 70, 5),
    'F70': (f116, 70, 6),
    'C71': (f117, 71, 3),
    'D71': (f117, 71, 4),
    'E71': (f117, 71, 5),
    'F71': (f117, 71, 6),
    'C72': (f118, 72, 3),
    'D72': (f119, 72, 4),
    'E72': (f120, 72, 5),
    'F72': (f121, 72, 6),
    'B76': (f122, 76, 2),
    'C76': (f123, 76, 3),
    'D76': (f124, 76, 4),
    'E76': (f125, 76, 5),
    'F76': (f126, 76, 6),
    'B77': (f127, 77, 2),
    'C77': (f128, 77, 3),
    'D77': (f129, 77, 4),
    'E77': (f130, 77, 5),
    'F77': (f131, 77, 6),
    'B78': (f132, 78, 2),
    'C78': (f133, 78, 3),
    'D78': (f134, 78, 4),
    'E78': (f135, 78, 5),
    'F78': (f136, 78, 6),
    'B79': (f137, 79, 2),
    'C79': (f138, 79, 3),
    'D79': (f139, 79, 4),
    'E79': (f140, 79, 5),
    'F79': (f141, 79, 6),
    'B80': (f142, 80, 2),
    'C80': (f143, 80, 3),
    'D80': (f144, 80, 4),
    'E80': (f145, 80, 5),
    'F80': (f146, 80, 6),
    'B81': (f147, 81, 2),
    'C81': (f148, 81, 3),
    'D81': (f149, 81, 4),
    'E81': (f150, 81, 5),
    'F81': (f151, 81, 6),
    'B82': (f152, 82, 2),
    'C82': (f153, 82, 3),
    'D82': (f154, 82, 4),
    'E82': (f155, 82, 5),
    'F82': (f156, 82, 6),
    'B83': (f157, 83, 2),
    'C83': (f158, 83, 3),
    'D83': (f159, 83, 4),
    'E83': (f160, 83, 5),
    'F83': (f161, 83, 6),
    'B84': (f162, 84, 2),
    'C84': (f163, 84, 3),
    'D84': (f164, 84, 4),
    'E84': (f165, 84, 5),
    'F84': (f166, 84, 6),
    'B85': (f167, 85, 2),
    'C85': (f168, 85, 3),
    'D85': (f169, 85, 4),
    'E85': (f170, 85, 5),
    'F85': (f171, 85, 6),
    'B86': (f172, 86, 2),
    'C86': (f173, 86, 3),
    'D86': (f174, 86, 4),
    'E86': (f175, 86, 5),
    'F86': (f176, 86, 6),
    'B87': (f177, 87, 2),
    'C87': (f178, 87, 3),
    'D87': (f179, 87, 4),
    'E87': (f180, 87, 5),
    'F87': (f181, 87, 6),
    'C89': (f182, 89, 3),
}
