# -*- coding: utf-8 -*-
"""Module sheet PP2_DINH_LUONG – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.5.xlsx (SHA-256 f19f3f74b720c1a067a640a31d58441793bf64bfdd12629f17ba4be6623c56f4).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'PP2_DINH_LUONG'
SI = 10

def f0(E, r, c):
    # Ô C6
    # = TRUNG_GIAN!$D$6
    return E.v(23,6,4)

def f1(E, r, c):
    # Ô C7
    # = IF(OR(NHAP_LIEU!$C$18="",NHAP_LIEU!$C$18<=P_THO_NGAY),INDEX(CSDL!$J$18:$J$29,$C$6+1),INDEX(CSDL!$K$18:$K$29,$C$6+1))
    return (X.INDEX(E.rg(24,18,10,29,10), X.add(E.v(10,6,3),1.0)) if X.truthy(X.OR(X.eq(E.v(2,18,3),''), X.le(E.v(2,18,3),E.v(24,98,2)))) else X.INDEX(E.rg(24,18,11,29,11), X.add(E.v(10,6,3),1.0)))

def f2(E, r, c):
    # Ô D7
    # = IF(AND(OR($C$6=1,$C$6=4,$C$6=7,$C$6=10),NHAP_LIEU!$C$18=""),"⚠ Thiếu số ngày tới tiết Lập – đang GIẢ ĐỊNH Thổ nắm lệnh (P2 tr.346); nhập NHAP_LIEU C11
    #    hoặc ngày giờ sinh","")
    return ('⚠ Thiếu số ngày tới tiết Lập – đang GIẢ ĐỊNH Thổ nắm lệnh (P2 tr.346); nhập NHAP_LIEU C11 hoặc ngày giờ sinh' if X.truthy(X.AND(X.OR(X.eq(E.v(10,6,3),1.0), X.eq(E.v(10,6,3),4.0), X.eq(E.v(10,6,3),7.0), X.eq(E.v(10,6,3),10.0)), X.eq(E.v(2,18,3),''))) else '')

def f3(E, r, c):
    # Ô C8
    # = TRUNG_GIAN!$D$21
    return E.v(23,21,4)

def f4(E, r, c):
    # Ô C11
    # = IF(TRUNG_GIAN!$C$44=0,0,IF((COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!$C$46)>0),1,0))
    return (0.0 if X.truthy(X.eq(E.v(23,44,3),0.0)) else (1.0 if X.truthy(X.gt(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,46,3)),0.0)) else 0.0))

def f5(E, r, c):
    # Ô C12, D12, E12
    # = IF(TRUNG_GIAN!C$57=1,IF((COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!C$55)>0),1,0),0)
    return ((1.0 if X.truthy(X.gt(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,55,c)),0.0)) else 0.0) if X.truthy(X.eq(E.v(23,57,c),1.0)) else 0.0)

def f6(E, r, c):
    # Ô C13, D13, E13
    # = IF(AND(TRUNG_GIAN!C$65=1,TRUNG_GIAN!C$68=0,OR($C$8=TRUNG_GIAN!C$64,MOD($C$8+1,5)=TRUNG_GIAN!C$64),(COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!C$64)>0)),1
    #   ,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,65,c),1.0), X.eq(E.v(23,68,c),0.0), X.OR(X.eq(E.v(10,8,3),E.v(23,64,c)), X.eq(X.MOD(X.add(E.v(10,8,3),1.0), 5.0),E.v(23,64,c))), X.gt(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,64,c)),0.0))) else 0.0)

def f7(E, r, c):
    # Ô C14
    # = IF(AND(TRUNG_GIAN!C$52=1,$C$11=1),TRUNG_GIAN!$C$46,IF(C12=1,TRUNG_GIAN!C$55,IF(C13=1,TRUNG_GIAN!C$64,-1)))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,11,3),1.0))) else (E.v(23,55,c) if X.truthy(X.eq(E.v(10,r-2,c),1.0)) else (E.v(23,64,c) if X.truthy(X.eq(E.v(10,r-1,c),1.0)) else X.neg(1.0))))

def f8(E, r, c):
    # Ô D14, E14
    # = IF(AND(TRUNG_GIAN!D$52=1,$C$11=1),TRUNG_GIAN!$C$46,IF(C12=1,TRUNG_GIAN!C$55,IF(D12=1,TRUNG_GIAN!D$55,IF(C13=1,TRUNG_GIAN!C$64,IF(D13=1,TRUNG_GIAN!D$64
    #   ,-1)))))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,11,3),1.0))) else (E.v(23,55,c-1) if X.truthy(X.eq(E.v(10,r-2,c-1),1.0)) else (E.v(23,55,c) if X.truthy(X.eq(E.v(10,r-2,c),1.0)) else (E.v(23,64,c-1) if X.truthy(X.eq(E.v(10,r-1,c-1),1.0)) else (E.v(23,64,c) if X.truthy(X.eq(E.v(10,r-1,c),1.0)) else X.neg(1.0))))))

def f9(E, r, c):
    # Ô F14
    # = IF(AND(TRUNG_GIAN!F$52=1,$C$11=1),TRUNG_GIAN!$C$46,IF(E12=1,TRUNG_GIAN!E$55,IF(E13=1,TRUNG_GIAN!E$64,-1)))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,11,3),1.0))) else (E.v(23,55,c-1) if X.truthy(X.eq(E.v(10,r-2,c-1),1.0)) else (E.v(23,64,c-1) if X.truthy(X.eq(E.v(10,r-1,c-1),1.0)) else X.neg(1.0))))

def f10(E, r, c):
    # Ô C15
    # = IF($D$14>=0,$D$14,$C$7)
    return (E.v(10,14,4) if X.truthy(X.ge(E.v(10,14,4),0.0)) else E.v(10,7,3))

def f11(E, r, c):
    # Ô C16, D16, E16
    # = IF(TRUNG_GIAN!C$76<0,0,IF(OR(TRUNG_GIAN!C$76=$C$15,TRUNG_GIAN!C$76=TRUNG_GIAN!$D$9,TRUNG_GIAN!C$76=TRUNG_GIAN!$D$21,TRUNG_GIAN!C$76=TRUNG_GIAN!$D$22,T
    #   RUNG_GIAN!C$76=TRUNG_GIAN!$D$23),1,0))
    return (0.0 if X.truthy(X.lt(E.v(23,76,c),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,76,c),E.v(10,15,3)), X.eq(E.v(23,76,c),E.v(23,9,4)), X.eq(E.v(23,76,c),E.v(23,21,4)), X.eq(E.v(23,76,c),E.v(23,22,4)), X.eq(E.v(23,76,c),E.v(23,23,4)))) else 0.0))

def f12(E, r, c):
    # Ô C17, D17, E17
    # = IF(OR(TRUNG_GIAN!C$76<0,TRUNG_GIAN!C$78=1),0,IF(AND(C16=1,AND(OR(TRUNG_GIAN!C$9=TRUNG_GIAN!C$76,MOD(TRUNG_GIAN!C$9+1,5)=TRUNG_GIAN!C$76,C14=TRUNG_GIAN
    #   !C$76),OR(TRUNG_GIAN!D$9=TRUNG_GIAN!C$76,MOD(TRUNG_GIAN!D$9+1,5)=TRUNG_GIAN!C$76,D14=TRUNG_GIAN!C$76))),1,IF(AND(IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TR
    #   UNG_GIAN!C$7)=$C$15,AND(OR(TRUNG_GIAN!C$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),MOD(TRUNG_GIAN!C$9+1,5)=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$
    #   7,TRUNG_GIAN!C$7),C14=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7)),OR(TRUNG_GIAN!D$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),MOD(TRU
    #   NG_GIAN!D$9+1,5)=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),D14=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7)))),2,IF(AND(IF(TRUNG_GIAN!C
    #   $8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7)=$C$15,AND(OR(TRUNG_GIAN!C$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7),MOD(TRUNG_GIAN!C$9+1,5)=IF(TRUNG_GI
    #   AN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7),C14=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7)),OR(TRUNG_GIAN!D$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,T
    #   RUNG_GIAN!D$7),MOD(TRUNG_GIAN!D$9+1,5)=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7),D14=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7)))),3,
    #   0))))
    return (0.0 if X.truthy(X.OR(X.lt(E.v(23,76,c),0.0), X.eq(E.v(23,78,c),1.0))) else (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-1,c),1.0), X.AND(X.OR(X.eq(E.v(23,9,c),E.v(23,76,c)), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),E.v(23,76,c)), X.eq(E.v(10,r-3,c),E.v(23,76,c))), X.OR(X.eq(E.v(23,9,c+1),E.v(23,76,c)), X.eq(X.MOD(X.add(E.v(23,9,c+1),1.0), 5.0),E.v(23,76,c)), X.eq(E.v(10,r-3,c+1),E.v(23,76,c)))))) else (2.0 if X.truthy(X.AND(X.eq((E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c)),E.v(10,15,3)), X.AND(X.OR(X.eq(E.v(23,9,c),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(E.v(10,r-3,c),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c)))), X.OR(X.eq(E.v(23,9,c+1),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(X.MOD(X.add(E.v(23,9,c+1),1.0), 5.0),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(E.v(10,r-3,c+1),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))))))) else (3.0 if X.truthy(X.AND(X.eq((E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1)),E.v(10,15,3)), X.AND(X.OR(X.eq(E.v(23,9,c),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(E.v(10,r-3,c),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1)))), X.OR(X.eq(E.v(23,9,c+1),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(X.MOD(X.add(E.v(23,9,c+1),1.0), 5.0),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(E.v(10,r-3,c+1),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))))))) else 0.0))))

def f13(E, r, c):
    # Ô C18, D18, E18, C28, D28, E28
    # = CHOOSE(C17+1,-1,TRUNG_GIAN!C$76,IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7))
    return X.CHOOSE(X.add(E.v(10,r-1,c),1.0), lambda: X.neg(1.0), lambda: E.v(23,76,c), lambda: (E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c)), lambda: (E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1)))

def f14(E, r, c):
    # Ô C21
    # = IF(TRUNG_GIAN!$C$44=0,0,IF((COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!$C$46)+COUNTIF($C$18:$E$18,TRUNG_GIAN!$C$46)>0),1,0))
    return (0.0 if X.truthy(X.eq(E.v(23,44,3),0.0)) else (1.0 if X.truthy(X.gt(X.add(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,46,3)),X.COUNTIF(E.rg(10,18,3,18,5), E.v(23,46,3))),0.0)) else 0.0))

def f15(E, r, c):
    # Ô C22, D22, E22
    # = IF(TRUNG_GIAN!C$57=1,IF((COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!C$55)+COUNTIF($C$18:$E$18,TRUNG_GIAN!C$55)>0),1,0),0)
    return ((1.0 if X.truthy(X.gt(X.add(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,55,c)),X.COUNTIF(E.rg(10,18,3,18,5), E.v(23,55,c))),0.0)) else 0.0) if X.truthy(X.eq(E.v(23,57,c),1.0)) else 0.0)

def f16(E, r, c):
    # Ô C23, D23, E23
    # = IF(AND(TRUNG_GIAN!C$65=1,TRUNG_GIAN!C$68=0,OR($C$8=TRUNG_GIAN!C$64,MOD($C$8+1,5)=TRUNG_GIAN!C$64),(COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!C$64)+COUNT
    #   IF($C$18:$E$18,TRUNG_GIAN!C$64)>0)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,65,c),1.0), X.eq(E.v(23,68,c),0.0), X.OR(X.eq(E.v(10,8,3),E.v(23,64,c)), X.eq(X.MOD(X.add(E.v(10,8,3),1.0), 5.0),E.v(23,64,c))), X.gt(X.add(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,64,c)),X.COUNTIF(E.rg(10,18,3,18,5), E.v(23,64,c))),0.0))) else 0.0)

def f17(E, r, c):
    # Ô C24
    # = IF(AND(TRUNG_GIAN!C$52=1,$C$21=1),TRUNG_GIAN!$C$46,IF(C22=1,TRUNG_GIAN!C$55,IF(C23=1,TRUNG_GIAN!C$64,-1)))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,21,3),1.0))) else (E.v(23,55,c) if X.truthy(X.eq(E.v(10,r-2,c),1.0)) else (E.v(23,64,c) if X.truthy(X.eq(E.v(10,r-1,c),1.0)) else X.neg(1.0))))

def f18(E, r, c):
    # Ô D24, E24
    # = IF(AND(TRUNG_GIAN!D$52=1,$C$21=1),TRUNG_GIAN!$C$46,IF(C22=1,TRUNG_GIAN!C$55,IF(D22=1,TRUNG_GIAN!D$55,IF(C23=1,TRUNG_GIAN!C$64,IF(D23=1,TRUNG_GIAN!D$64
    #   ,-1)))))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,21,3),1.0))) else (E.v(23,55,c-1) if X.truthy(X.eq(E.v(10,r-2,c-1),1.0)) else (E.v(23,55,c) if X.truthy(X.eq(E.v(10,r-2,c),1.0)) else (E.v(23,64,c-1) if X.truthy(X.eq(E.v(10,r-1,c-1),1.0)) else (E.v(23,64,c) if X.truthy(X.eq(E.v(10,r-1,c),1.0)) else X.neg(1.0))))))

def f19(E, r, c):
    # Ô F24
    # = IF(AND(TRUNG_GIAN!F$52=1,$C$21=1),TRUNG_GIAN!$C$46,IF(E22=1,TRUNG_GIAN!E$55,IF(E23=1,TRUNG_GIAN!E$64,-1)))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,21,3),1.0))) else (E.v(23,55,c-1) if X.truthy(X.eq(E.v(10,r-2,c-1),1.0)) else (E.v(23,64,c-1) if X.truthy(X.eq(E.v(10,r-1,c-1),1.0)) else X.neg(1.0))))

def f20(E, r, c):
    # Ô C25
    # = IF($D$24>=0,$D$24,$C$7)
    return (E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(10,7,3))

def f21(E, r, c):
    # Ô C26, D26, E26
    # = IF(TRUNG_GIAN!C$76<0,0,IF(OR(TRUNG_GIAN!C$76=$C$25,TRUNG_GIAN!C$76=TRUNG_GIAN!$D$9,TRUNG_GIAN!C$76=TRUNG_GIAN!$D$21,TRUNG_GIAN!C$76=TRUNG_GIAN!$D$22,T
    #   RUNG_GIAN!C$76=TRUNG_GIAN!$D$23),1,0))
    return (0.0 if X.truthy(X.lt(E.v(23,76,c),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,76,c),E.v(10,25,3)), X.eq(E.v(23,76,c),E.v(23,9,4)), X.eq(E.v(23,76,c),E.v(23,21,4)), X.eq(E.v(23,76,c),E.v(23,22,4)), X.eq(E.v(23,76,c),E.v(23,23,4)))) else 0.0))

def f22(E, r, c):
    # Ô C27, D27, E27
    # = IF(OR(TRUNG_GIAN!C$76<0,TRUNG_GIAN!C$78=1),0,IF(AND(C26=1,AND(OR(TRUNG_GIAN!C$9=TRUNG_GIAN!C$76,MOD(TRUNG_GIAN!C$9+1,5)=TRUNG_GIAN!C$76,C24=TRUNG_GIAN
    #   !C$76),OR(TRUNG_GIAN!D$9=TRUNG_GIAN!C$76,MOD(TRUNG_GIAN!D$9+1,5)=TRUNG_GIAN!C$76,D24=TRUNG_GIAN!C$76))),1,IF(AND(IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TR
    #   UNG_GIAN!C$7)=$C$25,AND(OR(TRUNG_GIAN!C$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),MOD(TRUNG_GIAN!C$9+1,5)=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$
    #   7,TRUNG_GIAN!C$7),C24=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7)),OR(TRUNG_GIAN!D$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),MOD(TRU
    #   NG_GIAN!D$9+1,5)=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7),D24=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!D$7,TRUNG_GIAN!C$7)))),2,IF(AND(IF(TRUNG_GIAN!C
    #   $8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7)=$C$25,AND(OR(TRUNG_GIAN!C$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7),MOD(TRUNG_GIAN!C$9+1,5)=IF(TRUNG_GI
    #   AN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7),C24=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7)),OR(TRUNG_GIAN!D$9=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,T
    #   RUNG_GIAN!D$7),MOD(TRUNG_GIAN!D$9+1,5)=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7),D24=IF(TRUNG_GIAN!C$8=0,TRUNG_GIAN!C$7,TRUNG_GIAN!D$7)))),3,
    #   0))))
    return (0.0 if X.truthy(X.OR(X.lt(E.v(23,76,c),0.0), X.eq(E.v(23,78,c),1.0))) else (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-1,c),1.0), X.AND(X.OR(X.eq(E.v(23,9,c),E.v(23,76,c)), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),E.v(23,76,c)), X.eq(E.v(10,r-3,c),E.v(23,76,c))), X.OR(X.eq(E.v(23,9,c+1),E.v(23,76,c)), X.eq(X.MOD(X.add(E.v(23,9,c+1),1.0), 5.0),E.v(23,76,c)), X.eq(E.v(10,r-3,c+1),E.v(23,76,c)))))) else (2.0 if X.truthy(X.AND(X.eq((E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c)),E.v(10,25,3)), X.AND(X.OR(X.eq(E.v(23,9,c),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(E.v(10,r-3,c),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c)))), X.OR(X.eq(E.v(23,9,c+1),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(X.MOD(X.add(E.v(23,9,c+1),1.0), 5.0),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))), X.eq(E.v(10,r-3,c+1),(E.v(23,7,c+1) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c))))))) else (3.0 if X.truthy(X.AND(X.eq((E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1)),E.v(10,25,3)), X.AND(X.OR(X.eq(E.v(23,9,c),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(E.v(10,r-3,c),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1)))), X.OR(X.eq(E.v(23,9,c+1),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(X.MOD(X.add(E.v(23,9,c+1),1.0), 5.0),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))), X.eq(E.v(10,r-3,c+1),(E.v(23,7,c) if X.truthy(X.eq(E.v(23,8,c),0.0)) else E.v(23,7,c+1))))))) else 0.0))))

def f23(E, r, c):
    # Ô C31, D31, E31
    # = IF(AND(TRUNG_GIAN!C$73=1,TRUNG_GIAN!C$60=0,TRUNG_GIAN!D$60=0,TRUNG_GIAN!C$69=0,TRUNG_GIAN!D$69=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,73,c),1.0), X.eq(E.v(23,60,c),0.0), X.eq(E.v(23,60,c+1),0.0), X.eq(E.v(23,69,c),0.0), X.eq(E.v(23,69,c+1),0.0))) else 0.0)

def f24(E, r, c):
    # Ô C32
    # = C31
    return E.v(10,r-1,c)

def f25(E, r, c):
    # Ô D32, E32
    # = C31+D31
    return X.add(E.v(10,r-1,c-1),E.v(10,r-1,c))

def f26(E, r, c):
    # Ô F32
    # = E31
    return E.v(10,r-1,c-1)

def f27(E, r, c):
    # Ô C33, D33, E33, F33
    # = IF(OR(AND(TRUNG_GIAN!C$60=1,C24<0),AND(TRUNG_GIAN!C$69=1,C24<0),C32>0),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,60,c),1.0), X.lt(E.v(10,r-9,c),0.0)), X.AND(X.eq(E.v(23,69,c),1.0), X.lt(E.v(10,r-9,c),0.0)), X.gt(E.v(10,r-1,c),0.0))) else 0.0)

def f28(E, r, c):
    # Ô C36, D36, E36, F36
    # = IF(C24>=0,C24,TRUNG_GIAN!C$21)
    return (E.v(10,r-12,c) if X.truthy(X.ge(E.v(10,r-12,c),0.0)) else E.v(23,21,c))

def f29(E, r, c):
    # Ô C37, D37, E37, F37
    # = IF(OR(C24>=0,C33=1),-1,TRUNG_GIAN!C$22)
    return (X.neg(1.0) if X.truthy(X.OR(X.ge(E.v(10,r-13,c),0.0), X.eq(E.v(10,r-4,c),1.0))) else E.v(23,22,c))

def f30(E, r, c):
    # Ô C38, D38, E38, F38
    # = IF(OR(C24>=0,C33=1),-1,TRUNG_GIAN!C$23)
    return (X.neg(1.0) if X.truthy(X.OR(X.ge(E.v(10,r-14,c),0.0), X.eq(E.v(10,r-5,c),1.0))) else E.v(23,23,c))

def f31(E, r, c):
    # Ô C40, D40, E40, F40
    # = IF(SUMPRODUCT(($C$36:$F$38>=0)*((($C$36:$F$38=TRUNG_GIAN!C$7)+(MOD($C$36:$F$38+1,5)=TRUNG_GIAN!C$7))>0))>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('ge',E.rg(10,36,3,38,6),0.0),X.A('gt',X.A('add',X.A('eq',E.rg(10,36,3,38,6),E.v(23,7,c)),X.A('eq',X.AF_MOD(X.A('add',E.rg(10,36,3,38,6),1.0), 5.0),E.v(23,7,c))),0.0))),0.0)) else 0.0)

def f32(E, r, c):
    # Ô C41, D41, E41, F41
    # = IF(C40=1,P_DUA,P_HUPHU)
    return (E.v(24,76,2) if X.truthy(X.eq(E.v(10,r-1,c),1.0)) else E.v(24,77,2))

def f33(E, r, c):
    # Ô C43, D43, E43
    # = IF(TRUNG_GIAN!C$76>=0,1,0)
    return (1.0 if X.truthy(X.ge(E.v(23,76,c),0.0)) else 0.0)

def f34(E, r, c):
    # Ô C44, D44, E44
    # = TRUNG_GIAN!C$78
    return E.v(23,78,c)

def f35(E, r, c):
    # Ô C45, D45, E45
    # = IF(AND(C43=1,C44=0,C40=0,D40=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-2,c),1.0), X.eq(E.v(10,r-1,c),0.0), X.eq(E.v(10,r-5,c),0.0), X.eq(E.v(10,r-5,c+1),0.0))) else 0.0)

def f36(E, r, c):
    # Ô C46, D46, E46
    # = IF(AND(C43=1,C44=0,C45=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-3,c),1.0), X.eq(E.v(10,r-2,c),0.0), X.eq(E.v(10,r-1,c),0.0))) else 0.0)

def f37(E, r, c):
    # Ô C47, D47, E47
    # = IF(AND(C46=1,C27>0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-1,c),1.0), X.gt(E.v(10,r-20,c),0.0))) else 0.0)

def f38(E, r, c):
    # Ô C49
    # = IF(OR(AND(C43=1,C44=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(10,r-6,c),1.0), X.eq(E.v(10,r-5,c),1.0)))) else 0.0)

def f39(E, r, c):
    # Ô D49, E49
    # = IF(OR(AND(C43=1,C44=1),AND(D43=1,D44=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(10,r-6,c-1),1.0), X.eq(E.v(10,r-5,c-1),1.0)), X.AND(X.eq(E.v(10,r-6,c),1.0), X.eq(E.v(10,r-5,c),1.0)))) else 0.0)

def f40(E, r, c):
    # Ô F49
    # = IF(OR(AND(E43=1,E44=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(10,r-6,c-1),1.0), X.eq(E.v(10,r-5,c-1),1.0)))) else 0.0)

def f41(E, r, c):
    # Ô C50, C51
    # = IF(OR(C46=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(10,r-4,c),1.0))) else 0.0)

def f42(E, r, c):
    # Ô D50, E50, D51, E51
    # = IF(OR(C46=1,D46=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(10,r-4,c-1),1.0), X.eq(E.v(10,r-4,c),1.0))) else 0.0)

def f43(E, r, c):
    # Ô F50, F51
    # = IF(OR(E46=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(10,r-4,c-1),1.0))) else 0.0)

def f44(E, r, c):
    # Ô C52
    # = IF(C47=1,C28,TRUNG_GIAN!C$7)
    return (E.v(10,r-24,c) if X.truthy(X.eq(E.v(10,r-5,c),1.0)) else E.v(23,7,c))

def f45(E, r, c):
    # Ô D52, E52
    # = IF(C47=1,C28,IF(D47=1,D28,TRUNG_GIAN!D$7))
    return (E.v(10,r-24,c-1) if X.truthy(X.eq(E.v(10,r-5,c-1),1.0)) else (E.v(10,r-24,c) if X.truthy(X.eq(E.v(10,r-5,c),1.0)) else E.v(23,7,c)))

def f46(E, r, c):
    # Ô F52
    # = IF(E47=1,E28,TRUNG_GIAN!F$7)
    return (E.v(10,r-24,c-1) if X.truthy(X.eq(E.v(10,r-5,c-1),1.0)) else E.v(23,7,c))

def f47(E, r, c):
    # Ô C53
    # = IF(OR(AND(C46=1,C47=0,MOD(TRUNG_GIAN!D$7+2,5)=TRUNG_GIAN!C$7)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(10,r-7,c),1.0), X.eq(E.v(10,r-6,c),0.0), X.eq(X.MOD(X.add(E.v(23,7,c+1),2.0), 5.0),E.v(23,7,c))))) else 0.0)

def f48(E, r, c):
    # Ô D53, E53
    # = IF(OR(AND(C46=1,C47=0,MOD(TRUNG_GIAN!C$7+2,5)=TRUNG_GIAN!D$7),AND(D46=1,D47=0,MOD(TRUNG_GIAN!E$7+2,5)=TRUNG_GIAN!D$7)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(10,r-7,c-1),1.0), X.eq(E.v(10,r-6,c-1),0.0), X.eq(X.MOD(X.add(E.v(23,7,c-1),2.0), 5.0),E.v(23,7,c))), X.AND(X.eq(E.v(10,r-7,c),1.0), X.eq(E.v(10,r-6,c),0.0), X.eq(X.MOD(X.add(E.v(23,7,c+1),2.0), 5.0),E.v(23,7,c))))) else 0.0)

def f49(E, r, c):
    # Ô F53
    # = IF(OR(AND(E46=1,E47=0,MOD(TRUNG_GIAN!E$7+2,5)=TRUNG_GIAN!F$7)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(10,r-7,c-1),1.0), X.eq(E.v(10,r-6,c-1),0.0), X.eq(X.MOD(X.add(E.v(23,7,c-1),2.0), 5.0),E.v(23,7,c))))) else 0.0)

def f50(E, r, c):
    # Ô C54, D54, E54, F54
    # = IF(C49=1,MAX(0,C41-P_TRANH),IF(C51=1,P_CAPHOA/2,IF(C50=1,P_DUA-IF(C53=1,P_HOP_KHAC,0),C41)))
    return (X.MAX(0.0, X.sub(E.v(10,r-13,c),E.v(24,81,2))) if X.truthy(X.eq(E.v(10,r-5,c),1.0)) else (X.div(E.v(24,93,2),2.0) if X.truthy(X.eq(E.v(10,r-3,c),1.0)) else (X.sub(E.v(24,76,2),(E.v(24,80,2) if X.truthy(X.eq(E.v(10,r-1,c),1.0)) else 0.0)) if X.truthy(X.eq(E.v(10,r-4,c),1.0)) else E.v(10,r-13,c))))

def f51(E, r, c):
    # Ô C55, D55, E55, F55
    # = IF(AND(C49=0,C50=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-6,c),0.0), X.eq(E.v(10,r-5,c),0.0))) else 0.0)

def f52(E, r, c):
    # Ô D57, E58, F59
    # = IF(AND(C$55=1,D$55=1,MOD(C$52+2,5)=D$52),P_KHAC_GAN,0)
    return (E.v(24,78,2) if X.truthy(X.AND(X.eq(E.v(10,55,c-1),1.0), X.eq(E.v(10,55,c),1.0), X.eq(X.MOD(X.add(E.v(10,52,c-1),2.0), 5.0),E.v(10,52,c)))) else 0.0)

def f53(E, r, c):
    # Ô E57, F58
    # = IF(AND(C$55=1,E$55=1,MOD(C$52+2,5)=E$52),P_KHAC_CACH,0)
    return (E.v(24,79,2) if X.truthy(X.AND(X.eq(E.v(10,55,c-2),1.0), X.eq(E.v(10,55,c),1.0), X.eq(X.MOD(X.add(E.v(10,52,c-2),2.0), 5.0),E.v(10,52,c)))) else 0.0)

def f54(E, r, c):
    # Ô F57
    # = IF(AND(C$55=1,F$55=1,MOD(C$52+2,5)=F$52),0,0)
    return (0.0 if X.truthy(X.AND(X.eq(E.v(10,55,c-3),1.0), X.eq(E.v(10,55,c),1.0), X.eq(X.MOD(X.add(E.v(10,52,c-3),2.0), 5.0),E.v(10,52,c)))) else 0.0)

def f55(E, r, c):
    # Ô C58, D59, E60
    # = IF(AND(D$55=1,C$55=1,MOD(D$52+2,5)=C$52),P_KHAC_GAN,0)
    return (E.v(24,78,2) if X.truthy(X.AND(X.eq(E.v(10,55,c+1),1.0), X.eq(E.v(10,55,c),1.0), X.eq(X.MOD(X.add(E.v(10,52,c+1),2.0), 5.0),E.v(10,52,c)))) else 0.0)

def f56(E, r, c):
    # Ô C59, D60
    # = IF(AND(E$55=1,C$55=1,MOD(E$52+2,5)=C$52),P_KHAC_CACH,0)
    return (E.v(24,79,2) if X.truthy(X.AND(X.eq(E.v(10,55,c+2),1.0), X.eq(E.v(10,55,c),1.0), X.eq(X.MOD(X.add(E.v(10,52,c+2),2.0), 5.0),E.v(10,52,c)))) else 0.0)

def f57(E, r, c):
    # Ô C60
    # = IF(AND(F$55=1,C$55=1,MOD(F$52+2,5)=C$52),0,0)
    return (0.0 if X.truthy(X.AND(X.eq(E.v(10,55,c+3),1.0), X.eq(E.v(10,55,c),1.0), X.eq(X.MOD(X.add(E.v(10,52,c+3),2.0), 5.0),E.v(10,52,c)))) else 0.0)

def f58(E, r, c):
    # Ô C62, D62, E62, F62
    # = IF(AND(C55=1,C40=1,C54>=P_KHA_KHAC),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-7,c),1.0), X.eq(E.v(10,r-22,c),1.0), X.ge(E.v(10,r-8,c),E.v(24,85,2)))) else 0.0)

def f59(E, r, c):
    # Ô C63, D63, E63, F63
    # = $C$62*C57+$D$62*C58+$E$62*C59+$F$62*C60
    return X.add(X.add(X.add(X.mul(E.v(10,62,3),E.v(10,r-6,c)),X.mul(E.v(10,62,4),E.v(10,r-5,c))),X.mul(E.v(10,62,5),E.v(10,r-4,c))),X.mul(E.v(10,62,6),E.v(10,r-3,c)))

def f60(E, r, c):
    # Ô C64, D64, E64, F64
    # = IF(AND(C55=1,C40=1,C54-C63>=P_KHA_KHAC),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-9,c),1.0), X.eq(E.v(10,r-24,c),1.0), X.ge(X.sub(E.v(10,r-10,c),E.v(10,r-1,c)),E.v(24,85,2)))) else 0.0)

def f61(E, r, c):
    # Ô C65, D65, E65, F65
    # = $C$64*C57+$D$64*C58+$E$64*C59+$F$64*C60
    return X.add(X.add(X.add(X.mul(E.v(10,64,3),E.v(10,r-8,c)),X.mul(E.v(10,64,4),E.v(10,r-7,c))),X.mul(E.v(10,64,5),E.v(10,r-6,c))),X.mul(E.v(10,64,6),E.v(10,r-5,c)))

def f62(E, r, c):
    # Ô C66, D66, E66, F66
    # = IF(AND(C55=1,C40=1,C54-C65>=P_KHA_KHAC),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(10,r-11,c),1.0), X.eq(E.v(10,r-26,c),1.0), X.ge(X.sub(E.v(10,r-12,c),E.v(10,r-1,c)),E.v(24,85,2)))) else 0.0)

def f63(E, r, c):
    # Ô C67, D67, E67, F67
    # = $C$66*C57+$D$66*C58+$E$66*C59+$F$66*C60
    return X.add(X.add(X.add(X.mul(E.v(10,66,3),E.v(10,r-10,c)),X.mul(E.v(10,66,4),E.v(10,r-9,c))),X.mul(E.v(10,66,5),E.v(10,r-8,c))),X.mul(E.v(10,66,6),E.v(10,r-7,c)))

def f64(E, r, c):
    # Ô C68, D68, E68, F68
    # = IF(AND(C55=1,C67>0),MAX(0,C54-C67),C54)
    return (X.MAX(0.0, X.sub(E.v(10,r-14,c),E.v(10,r-1,c))) if X.truthy(X.AND(X.eq(E.v(10,r-13,c),1.0), X.gt(E.v(10,r-1,c),0.0))) else E.v(10,r-14,c))

def f65(E, r, c):
    # Ô G68
    # = IF(AND(C66=C64,D66=D64,E66=E64,F66=F64),"Kiểm hội tụ A3: ỔN ĐỊNH sau 3 vòng – mỗi lần khắc chỉ trừ một lần (P2 tr.339-340)","⚠ NEED_REVIEW – A3 chưa h
    #   ội tụ sau 3 vòng: tập can đủ lực khắc vẫn thay đổi")
    return ('Kiểm hội tụ A3: ỔN ĐỊNH sau 3 vòng – mỗi lần khắc chỉ trừ một lần (P2 tr.339-340)' if X.truthy(X.AND(X.eq(E.v(10,r-2,c-4),E.v(10,r-4,c-4)), X.eq(E.v(10,r-2,c-3),E.v(10,r-4,c-3)), X.eq(E.v(10,r-2,c-2),E.v(10,r-4,c-2)), X.eq(E.v(10,r-2,c-1),E.v(10,r-4,c-1)))) else '⚠ NEED_REVIEW – A3 chưa hội tụ sau 3 vòng: tập can đủ lực khắc vẫn thay đổi')

def f66(E, r, c):
    # Ô C70, D70, E70, F70
    # = IF(C24>=0,C24,TRUNG_GIAN!C$9)
    return (E.v(10,r-46,c) if X.truthy(X.ge(E.v(10,r-46,c),0.0)) else E.v(23,9,c))

def f67(E, r, c):
    # Ô C71, D71, E71, F71
    # = IF(C51=1,0,IF(MOD(C52+1,5)=C70,-P_CAN_SINH_CHI,IF(MOD(C52+2,5)=C70,-P_CAN_KHAC_CHI,IF(MOD(C70+2,5)=C52,-P_CHI_KHAC_CAN,0))))
    return (0.0 if X.truthy(X.eq(E.v(10,r-20,c),1.0)) else (X.neg(E.v(24,82,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,r-19,c),1.0), 5.0),E.v(10,r-1,c))) else (X.neg(E.v(24,83,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,r-19,c),2.0), 5.0),E.v(10,r-1,c))) else (X.neg(E.v(24,84,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,r-1,c),2.0), 5.0),E.v(10,r-19,c))) else 0.0))))

def f68(E, r, c):
    # Ô C72, D72, E72, F72
    # = MAX(0,C68+C71)
    return X.MAX(0.0, X.add(E.v(10,r-4,c),E.v(10,r-1,c)))

def f69(E, r, c):
    # Ô C75, D75, E75, F75
    # = C24
    return E.v(10,r-51,c)

def f70(E, r, c):
    # Ô C76, D76, E76, F76
    # = IF(C24<0,0,IF(AND(TRUNG_GIAN!C$52=1,$C$21=1),IF(TRUNG_GIAN!$C$45="Tam hội",P_HOI,P_HOP)/3,IF(TRUNG_GIAN!C$59=1,P_BANHOP/2,P_LUCHOP/2)))
    return (0.0 if X.truthy(X.lt(E.v(10,r-52,c),0.0)) else (X.div((E.v(24,89,2) if X.truthy(X.eq(E.v(23,45,3),'Tam hội')) else E.v(24,90,2)),3.0) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(10,21,3),1.0))) else (X.div(E.v(24,91,2),2.0) if X.truthy(X.eq(E.v(23,59,c),1.0)) else X.div(E.v(24,92,2),2.0))))

def f71(E, r, c):
    # Ô C77, D77, E77, F77
    # = IF(C24<0,1,0)
    return (1.0 if X.truthy(X.lt(E.v(10,r-53,c),0.0)) else 0.0)

def f72(E, r, c):
    # Ô C78, D78, E78, F78
    # = IF(OR(TRUNG_GIAN!C$60=1,TRUNG_GIAN!C$69=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,60,c),1.0), X.eq(E.v(23,69,c),1.0))) else 0.0)

def f73(E, r, c):
    # Ô C79, D79, E79, F79
    # = TRUNG_GIAN!C$21
    return E.v(23,21,c)

def f74(E, r, c):
    # Ô C80, D80, E80, F80
    # = TRUNG_GIAN!C$16
    return E.v(23,16,c)

def f75(E, r, c):
    # Ô C81, D81, E81, F81
    # = IF(OR(C77=0,C51=1,C68<P_TRUC_DINH),0,IF(OR(C52=TRUNG_GIAN!C$9,MOD(C52+1,5)=TRUNG_GIAN!C$9),IF(C78=0,P_TD_CONG,0),IF(MOD(C52+2,5)=TRUNG_GIAN!C$9,-P_TD_
    #   KHAC,0)))
    return (0.0 if X.truthy(X.OR(X.eq(E.v(10,r-4,c),0.0), X.eq(E.v(10,r-30,c),1.0), X.lt(E.v(10,r-13,c),E.v(24,86,2)))) else ((E.v(24,87,2) if X.truthy(X.eq(E.v(10,r-3,c),0.0)) else 0.0) if X.truthy(X.OR(X.eq(E.v(10,r-29,c),E.v(23,9,c)), X.eq(X.MOD(X.add(E.v(10,r-29,c),1.0), 5.0),E.v(23,9,c)))) else (X.neg(E.v(24,88,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,r-29,c),2.0), 5.0),E.v(23,9,c))) else 0.0)))

def f76(E, r, c):
    # Ô C82
    # = IF(C32=0,0,IF(C32>=2,P_KEP_GIUA,IF(D32>=2,P_KEP_BIEN,IF(TRUNG_GIAN!C$74=1,0,MAX(0,INDEX(M_XL,$C$6+1,TRUNG_GIAN!C$6+1))))))
    return (0.0 if X.truthy(X.eq(E.v(10,r-50,c),0.0)) else (E.v(24,96,2) if X.truthy(X.ge(E.v(10,r-50,c),2.0)) else (E.v(24,97,2) if X.truthy(X.ge(E.v(10,r-50,c+1),2.0)) else (0.0 if X.truthy(X.eq(E.v(23,74,c),1.0)) else X.MAX(0.0, X.INDEX(E.rg(24,107,16,118,27), X.add(E.v(10,6,3),1.0), X.add(E.v(23,6,c),1.0)))))))

def f77(E, r, c):
    # Ô D82, E82
    # = IF(D32=0,0,IF(D32>=2,P_KEP_GIUA,IF(IF(C31=1,C32,E32)>=2,P_KEP_BIEN,IF(IF(C31=1,TRUNG_GIAN!C$74,TRUNG_GIAN!D$74)=1,0,MAX(0,INDEX(M_XL,$C$6+1,TRUNG_GIAN
    #   !D$6+1))))))
    return (0.0 if X.truthy(X.eq(E.v(10,r-50,c),0.0)) else (E.v(24,96,2) if X.truthy(X.ge(E.v(10,r-50,c),2.0)) else (E.v(24,97,2) if X.truthy(X.ge((E.v(10,r-50,c-1) if X.truthy(X.eq(E.v(10,r-51,c-1),1.0)) else E.v(10,r-50,c+1)),2.0)) else (0.0 if X.truthy(X.eq((E.v(23,74,c-1) if X.truthy(X.eq(E.v(10,r-51,c-1),1.0)) else E.v(23,74,c)),1.0)) else X.MAX(0.0, X.INDEX(E.rg(24,107,16,118,27), X.add(E.v(10,6,3),1.0), X.add(E.v(23,6,c),1.0)))))))

def f78(E, r, c):
    # Ô F82
    # = IF(F32=0,0,IF(F32>=2,P_KEP_GIUA,IF(E32>=2,P_KEP_BIEN,IF(TRUNG_GIAN!E$74=1,0,MAX(0,INDEX(M_XL,$C$6+1,TRUNG_GIAN!F$6+1))))))
    return (0.0 if X.truthy(X.eq(E.v(10,r-50,c),0.0)) else (E.v(24,96,2) if X.truthy(X.ge(E.v(10,r-50,c),2.0)) else (E.v(24,97,2) if X.truthy(X.ge(E.v(10,r-50,c-1),2.0)) else (0.0 if X.truthy(X.eq(E.v(23,74,c-1),1.0)) else X.MAX(0.0, X.INDEX(E.rg(24,107,16,118,27), X.add(E.v(10,6,3),1.0), X.add(E.v(23,6,c),1.0)))))))

def f79(E, r, c):
    # Ô C83, D83, E83, F83
    # = IF(C77=1,(C80+C81)*(1-C82),0)
    return (X.mul(X.add(E.v(10,r-3,c),E.v(10,r-2,c)),X.sub(1.0,E.v(10,r-1,c))) if X.truthy(X.eq(E.v(10,r-6,c),1.0)) else 0.0)

def f80(E, r, c):
    # Ô C84, D84, E84, F84
    # = IF(AND(C77=1,C33=0),TRUNG_GIAN!C$22,-1)
    return (E.v(23,22,c) if X.truthy(X.AND(X.eq(E.v(10,r-7,c),1.0), X.eq(E.v(10,r-51,c),0.0))) else X.neg(1.0))

def f81(E, r, c):
    # Ô C85, D85, E85, F85
    # = IF(C84>=0,TRUNG_GIAN!C$18,0)
    return (E.v(23,18,c) if X.truthy(X.ge(E.v(10,r-1,c),0.0)) else 0.0)

def f82(E, r, c):
    # Ô C86, D86, E86, F86
    # = IF(AND(C77=1,C33=0),TRUNG_GIAN!C$23,-1)
    return (E.v(23,23,c) if X.truthy(X.AND(X.eq(E.v(10,r-9,c),1.0), X.eq(E.v(10,r-53,c),0.0))) else X.neg(1.0))

def f83(E, r, c):
    # Ô C87, D87, E87, F87
    # = IF(C86>=0,TRUNG_GIAN!C$20,0)
    return (E.v(23,20,c) if X.truthy(X.ge(E.v(10,r-1,c),0.0)) else 0.0)

def f84(E, r, c):
    # Ô C91
    # = SUMPRODUCT(($C$52:$F$52=0)*$C$72:$F$72)
    return X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),0.0),E.rg(10,72,3,72,6)))

def f85(E, r, c):
    # Ô D91
    # = SUMPRODUCT(($C$75:$F$75=0)*$C$76:$F$76)+SUMPRODUCT(($C$79:$F$79=0)*$C$83:$F$83)+SUMPRODUCT(($C$84:$F$84=0)*$C$85:$F$85)+SUMPRODUCT(($C$86:$F$86=0)*$C$
    #   87:$F$87)
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,75,3,75,6),0.0),E.rg(10,76,3,76,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,79,3,79,6),0.0),E.rg(10,83,3,83,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,84,3,84,6),0.0),E.rg(10,85,3,85,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,86,3,86,6),0.0),E.rg(10,87,3,87,6))))

def f86(E, r, c):
    # Ô E91
    # = IF(0=$C$25,1+P_LENH,IF(MOD($C$25+2,5)=0,1-P_LENH,1))
    return (X.add(1.0,E.v(24,94,2)) if X.truthy(X.eq(0.0,E.v(10,25,3))) else (X.sub(1.0,E.v(24,94,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,25,3),2.0), 5.0),0.0)) else 1.0))

def f87(E, r, c):
    # Ô F91, F92, F93, F94, F95
    # = (C91+D91)*E91
    return X.mul(X.add(E.v(10,r,c-3),E.v(10,r,c-2)),E.v(10,r,c-1))

def f88(E, r, c):
    # Ô C92
    # = SUMPRODUCT(($C$52:$F$52=1)*$C$72:$F$72)
    return X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),1.0),E.rg(10,72,3,72,6)))

def f89(E, r, c):
    # Ô D92
    # = SUMPRODUCT(($C$75:$F$75=1)*$C$76:$F$76)+SUMPRODUCT(($C$79:$F$79=1)*$C$83:$F$83)+SUMPRODUCT(($C$84:$F$84=1)*$C$85:$F$85)+SUMPRODUCT(($C$86:$F$86=1)*$C$
    #   87:$F$87)
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,75,3,75,6),1.0),E.rg(10,76,3,76,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,79,3,79,6),1.0),E.rg(10,83,3,83,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,84,3,84,6),1.0),E.rg(10,85,3,85,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,86,3,86,6),1.0),E.rg(10,87,3,87,6))))

def f90(E, r, c):
    # Ô E92
    # = IF(1=$C$25,1+P_LENH,IF(MOD($C$25+2,5)=1,1-P_LENH,1))
    return (X.add(1.0,E.v(24,94,2)) if X.truthy(X.eq(1.0,E.v(10,25,3))) else (X.sub(1.0,E.v(24,94,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,25,3),2.0), 5.0),1.0)) else 1.0))

def f91(E, r, c):
    # Ô C93
    # = SUMPRODUCT(($C$52:$F$52=2)*$C$72:$F$72)
    return X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),2.0),E.rg(10,72,3,72,6)))

def f92(E, r, c):
    # Ô D93
    # = SUMPRODUCT(($C$75:$F$75=2)*$C$76:$F$76)+SUMPRODUCT(($C$79:$F$79=2)*$C$83:$F$83)+SUMPRODUCT(($C$84:$F$84=2)*$C$85:$F$85)+SUMPRODUCT(($C$86:$F$86=2)*$C$
    #   87:$F$87)
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,75,3,75,6),2.0),E.rg(10,76,3,76,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,79,3,79,6),2.0),E.rg(10,83,3,83,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,84,3,84,6),2.0),E.rg(10,85,3,85,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,86,3,86,6),2.0),E.rg(10,87,3,87,6))))

def f93(E, r, c):
    # Ô E93
    # = IF(2=$C$25,1+P_LENH,IF(MOD($C$25+2,5)=2,1-P_LENH,1))
    return (X.add(1.0,E.v(24,94,2)) if X.truthy(X.eq(2.0,E.v(10,25,3))) else (X.sub(1.0,E.v(24,94,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,25,3),2.0), 5.0),2.0)) else 1.0))

def f94(E, r, c):
    # Ô C94
    # = SUMPRODUCT(($C$52:$F$52=3)*$C$72:$F$72)
    return X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),3.0),E.rg(10,72,3,72,6)))

def f95(E, r, c):
    # Ô D94
    # = SUMPRODUCT(($C$75:$F$75=3)*$C$76:$F$76)+SUMPRODUCT(($C$79:$F$79=3)*$C$83:$F$83)+SUMPRODUCT(($C$84:$F$84=3)*$C$85:$F$85)+SUMPRODUCT(($C$86:$F$86=3)*$C$
    #   87:$F$87)
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,75,3,75,6),3.0),E.rg(10,76,3,76,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,79,3,79,6),3.0),E.rg(10,83,3,83,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,84,3,84,6),3.0),E.rg(10,85,3,85,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,86,3,86,6),3.0),E.rg(10,87,3,87,6))))

def f96(E, r, c):
    # Ô E94
    # = IF(3=$C$25,1+P_LENH,IF(MOD($C$25+2,5)=3,1-P_LENH,1))
    return (X.add(1.0,E.v(24,94,2)) if X.truthy(X.eq(3.0,E.v(10,25,3))) else (X.sub(1.0,E.v(24,94,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,25,3),2.0), 5.0),3.0)) else 1.0))

def f97(E, r, c):
    # Ô C95
    # = SUMPRODUCT(($C$52:$F$52=4)*$C$72:$F$72)
    return X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),4.0),E.rg(10,72,3,72,6)))

def f98(E, r, c):
    # Ô D95
    # = SUMPRODUCT(($C$75:$F$75=4)*$C$76:$F$76)+SUMPRODUCT(($C$79:$F$79=4)*$C$83:$F$83)+SUMPRODUCT(($C$84:$F$84=4)*$C$85:$F$85)+SUMPRODUCT(($C$86:$F$86=4)*$C$
    #   87:$F$87)
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,75,3,75,6),4.0),E.rg(10,76,3,76,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,79,3,79,6),4.0),E.rg(10,83,3,83,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,84,3,84,6),4.0),E.rg(10,85,3,85,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,86,3,86,6),4.0),E.rg(10,87,3,87,6))))

def f99(E, r, c):
    # Ô E95
    # = IF(4=$C$25,1+P_LENH,IF(MOD($C$25+2,5)=4,1-P_LENH,1))
    return (X.add(1.0,E.v(24,94,2)) if X.truthy(X.eq(4.0,E.v(10,25,3))) else (X.sub(1.0,E.v(24,94,2)) if X.truthy(X.eq(X.MOD(X.add(E.v(10,25,3),2.0), 5.0),4.0)) else 1.0))

def f100(E, r, c):
    # Ô C97
    # = INDEX(HANH_TEN,$C$25+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,25,3),1.0))

def f101(E, r, c):
    # Ô C98
    # = $E$52
    return E.v(10,52,5)

def f102(E, r, c):
    # Ô D98
    # = INDEX(HANH_TEN,C98+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,r,c-1),1.0))

def f103(E, r, c):
    # Ô C99
    # = INDEX($F$91:$F$95,C98+1)+INDEX($F$91:$F$95,MOD(C98+4,5)+1)
    return X.add(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(10,r-1,c),1.0)),X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,r-1,c),4.0), 5.0),1.0)))

def f104(E, r, c):
    # Ô C100
    # = SUM($F$91:$F$95)
    return X.SUM(E.rg(10,91,6,95,6))

def f105(E, r, c):
    # Ô C101
    # = IF(C100=0,0,C99/C100)
    return (0.0 if X.truthy(X.eq(E.v(10,r-1,c),0.0)) else X.div(E.v(10,r-2,c),E.v(10,r-1,c)))

def f106(E, r, c):
    # Ô C102
    # = P_NGUONG*C100
    return X.mul(E.v(24,95,2),E.v(10,r-2,c))

def f107(E, r, c):
    # Ô C103
    # = IF(C100<=0,"LỖI DỮ LIỆU – tổng độ bằng 0",IF(C100-C99<=0.000000001,"Độc vượng",IF(C101>=P_NGUONG,"Nhật chủ vượng","Nhật chủ nhược")))
    return ('LỖI DỮ LIỆU – tổng độ bằng 0' if X.truthy(X.le(E.v(10,r-3,c),0.0)) else ('Độc vượng' if X.truthy(X.le(X.sub(E.v(10,r-3,c),E.v(10,r-4,c)),1e-09)) else ('Nhật chủ vượng' if X.truthy(X.ge(E.v(10,r-2,c),E.v(24,95,2))) else 'Nhật chủ nhược')))

def f108(E, r, c):
    # Ô D106
    # = MAX(0,INDEX($F$91:$F$95,MOD($C$98+0,5)+1)-$E$72*INDEX($E$91:$E$95,$C$98+1))
    return X.MAX(0.0, X.sub(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,98,3),0.0), 5.0),1.0)),X.mul(E.v(10,72,5),X.INDEX(E.rg(10,91,5,95,5), X.add(E.v(10,98,3),1.0)))))

def f109(E, r, c):
    # Ô E106, E107, E108, E109, E110
    # = IF($C$100=0,0,D106/$C$100)
    return (0.0 if X.truthy(X.eq(E.v(10,100,3),0.0)) else X.div(E.v(10,r,c-1),E.v(10,100,3)))

def f110(E, r, c):
    # Ô D107
    # = MAX(0,INDEX($F$91:$F$95,MOD($C$98+1,5)+1))
    return X.MAX(0.0, X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,98,3),1.0), 5.0),1.0)))

def f111(E, r, c):
    # Ô D108
    # = MAX(0,INDEX($F$91:$F$95,MOD($C$98+2,5)+1))
    return X.MAX(0.0, X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,98,3),2.0), 5.0),1.0)))

def f112(E, r, c):
    # Ô D109
    # = MAX(0,INDEX($F$91:$F$95,MOD($C$98+3,5)+1))
    return X.MAX(0.0, X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,98,3),3.0), 5.0),1.0)))

def f113(E, r, c):
    # Ô D110
    # = MAX(0,INDEX($F$91:$F$95,MOD($C$98+4,5)+1))
    return X.MAX(0.0, X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,98,3),4.0), 5.0),1.0)))

def f114(E, r, c):
    # Ô C114
    # = "Độ theo hành (đã nhân hệ số lệnh): Mộc "&TEXT(F91,"0.0")&" (×"&E91&"), Hỏa "&TEXT(F92,"0.0")&" (×"&E92&"), Thổ "&TEXT(F93,"0.0")&" (×"&E93&"), Kim "&
    #   TEXT(F94,"0.0")&" (×"&E94&"), Thủy "&TEXT(F95,"0.0")&" (×"&E95&"); tổng "&TEXT(C100,"0.0")&" độ. Hành nắm lệnh: "&C97&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Độ theo hành (đã nhân hệ số lệnh): Mộc ',X.TEXT(E.v(10,r-23,c+3), '0.0')),' (×'),E.v(10,r-23,c+2)),'), Hỏa '),X.TEXT(E.v(10,r-22,c+3), '0.0')),' (×'),E.v(10,r-22,c+2)),'), Thổ '),X.TEXT(E.v(10,r-21,c+3), '0.0')),' (×'),E.v(10,r-21,c+2)),'), Kim '),X.TEXT(E.v(10,r-20,c+3), '0.0')),' (×'),E.v(10,r-20,c+2)),'), Thủy '),X.TEXT(E.v(10,r-19,c+3), '0.0')),' (×'),E.v(10,r-19,c+2)),'); tổng '),X.TEXT(E.v(10,r-14,c), '0.0')),' độ. Hành nắm lệnh: '),E.v(10,r-17,c)),'.')

def f115(E, r, c):
    # Ô C115
    # = "Phe mình = hành nhật chủ "&D98&" ("&TEXT(INDEX($F$91:$F$95,MOD(C98,5)+1),"0.0")&" độ) + hành sinh nhật chủ "&INDEX(HANH_TEN,MOD(C98+4,5)+1)&" ("&TEXT
    #   (INDEX($F$91:$F$95,MOD(C98+4,5)+1),"0.0")&" độ) = "&TEXT(C99,"0.0")&" độ = "&TEXT(C101,"0.0%")&" tổng."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Phe mình = hành nhật chủ ',E.v(10,r-17,c+1)),' ('),X.TEXT(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(E.v(10,r-17,c), 5.0),1.0)), '0.0')),' độ) + hành sinh nhật chủ '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(10,r-17,c),4.0), 5.0),1.0))),' ('),X.TEXT(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.add(E.v(10,r-17,c),4.0), 5.0),1.0)), '0.0')),' độ) = '),X.TEXT(E.v(10,r-16,c), '0.0')),' độ = '),X.TEXT(E.v(10,r-14,c), '0.0%')),' tổng.')

def f116(E, r, c):
    # Ô C116
    # = "Ngưỡng 40% = "&TEXT(C102,"0.0")&" độ. "&IF(C101>=0.4,"Phe mình vượt ngưỡng "&TEXT(C99-C102,"0.0")&" độ → VƯỢNG.","Phe mình còn thiếu "&TEXT(C102-C99,
    #   "0.0")&" độ → NHƯỢC.")&" "&IF(ABS(C101-0.4)<0.05,"Chênh với ngưỡng dưới 5 điểm % → SÁT NGƯỠNG, kết luận PP2 kém chắc chắn (nên xem DOI_CHIEU mục khuyế
    #   n nghị).","Cách ngưỡng rõ → kết luận PP2 chắc chắn.")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Ngưỡng 40% = ',X.TEXT(E.v(10,r-14,c), '0.0')),' độ. '),(X.cat(X.cat('Phe mình vượt ngưỡng ',X.TEXT(X.sub(E.v(10,r-17,c),E.v(10,r-14,c)), '0.0')),' độ → VƯỢNG.') if X.truthy(X.ge(E.v(10,r-15,c),0.4)) else X.cat(X.cat('Phe mình còn thiếu ',X.TEXT(X.sub(E.v(10,r-14,c),E.v(10,r-17,c)), '0.0')),' độ → NHƯỢC.'))),' '),('Chênh với ngưỡng dưới 5 điểm % → SÁT NGƯỠNG, kết luận PP2 kém chắc chắn (nên xem DOI_CHIEU mục khuyến nghị).' if X.truthy(X.lt(X.ABS(X.sub(E.v(10,r-15,c),0.4)),0.05)) else 'Cách ngưỡng rõ → kết luận PP2 chắc chắn.'))

def f117(E, r, c):
    # Ô C117
    # = "Thực thương "&TEXT(D107,"0.0")&" độ, tài "&TEXT(D108,"0.0")&" độ, quan sát "&TEXT(D109,"0.0")&" độ – lớn nhất là "&INDEX(B107:B109,MATCH(MAX(D107:D10
    #   9),D107:D109,0))&" (nhóm này thường là kỵ thần theo cân bằng khi nhật chủ nhược – P1 tr.99, 180)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thực thương ',X.TEXT(E.v(10,r-10,c+1), '0.0')),' độ, tài '),X.TEXT(E.v(10,r-9,c+1), '0.0')),' độ, quan sát '),X.TEXT(E.v(10,r-8,c+1), '0.0')),' độ – lớn nhất là '),X.INDEX(E.rg(10,r-10,c-1,r-8,c-1), X.MATCH(X.MAX(E.rg(10,r-10,c+1,r-8,c+1)), E.rg(10,r-10,c+1,r-8,c+1), 0.0))),' (nhóm này thường là kỵ thần theo cân bằng khi nhật chủ nhược – P1 tr.99, 180).')

def f118(E, r, c):
    # Ô C118
    # = "Nhật chủ "&D98&" – "&UPPER(C103)&": phe mình "&TEXT(C101,"0.0%")&IF(C101>=0.4," ≥ 40%"," < 40%")&" tổng độ."&IF(ISNUMBER(SEARCH("nhược",C103))," Nhật
    #    chủ yếu → cần ấn, tỉ kiếp làm dụng thần.",IF(ISNUMBER(SEARCH("vượng",C103))," Nhật chủ mạnh → cần quan sát, thực thương, tài làm dụng thần (trừ khi đ
    #   ộc vượng – theo thế)."," "))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',E.v(10,r-20,c+1)),' – '),X.UPPER(E.v(10,r-15,c))),': phe mình '),X.TEXT(E.v(10,r-17,c), '0.0%')),(' ≥ 40%' if X.truthy(X.ge(E.v(10,r-17,c),0.4)) else ' < 40%')),' tổng độ.'),(' Nhật chủ yếu → cần ấn, tỉ kiếp làm dụng thần.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('nhược', E.v(10,r-15,c)))) else (' Nhật chủ mạnh → cần quan sát, thực thương, tài làm dụng thần (trừ khi độc vượng – theo thế).' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('vượng', E.v(10,r-15,c)))) else ' ')))

CELLS = {
    'C6': (f0, 6, 3),
    'C7': (f1, 7, 3),
    'D7': (f2, 7, 4),
    'C8': (f3, 8, 3),
    'C11': (f4, 11, 3),
    'C12': (f5, 12, 3),
    'D12': (f5, 12, 4),
    'E12': (f5, 12, 5),
    'C13': (f6, 13, 3),
    'D13': (f6, 13, 4),
    'E13': (f6, 13, 5),
    'C14': (f7, 14, 3),
    'D14': (f8, 14, 4),
    'E14': (f8, 14, 5),
    'F14': (f9, 14, 6),
    'C15': (f10, 15, 3),
    'C16': (f11, 16, 3),
    'D16': (f11, 16, 4),
    'E16': (f11, 16, 5),
    'C17': (f12, 17, 3),
    'D17': (f12, 17, 4),
    'E17': (f12, 17, 5),
    'C18': (f13, 18, 3),
    'D18': (f13, 18, 4),
    'E18': (f13, 18, 5),
    'C21': (f14, 21, 3),
    'C22': (f15, 22, 3),
    'D22': (f15, 22, 4),
    'E22': (f15, 22, 5),
    'C23': (f16, 23, 3),
    'D23': (f16, 23, 4),
    'E23': (f16, 23, 5),
    'C24': (f17, 24, 3),
    'D24': (f18, 24, 4),
    'E24': (f18, 24, 5),
    'F24': (f19, 24, 6),
    'C25': (f20, 25, 3),
    'C26': (f21, 26, 3),
    'D26': (f21, 26, 4),
    'E26': (f21, 26, 5),
    'C27': (f22, 27, 3),
    'D27': (f22, 27, 4),
    'E27': (f22, 27, 5),
    'C28': (f13, 28, 3),
    'D28': (f13, 28, 4),
    'E28': (f13, 28, 5),
    'C31': (f23, 31, 3),
    'D31': (f23, 31, 4),
    'E31': (f23, 31, 5),
    'C32': (f24, 32, 3),
    'D32': (f25, 32, 4),
    'E32': (f25, 32, 5),
    'F32': (f26, 32, 6),
    'C33': (f27, 33, 3),
    'D33': (f27, 33, 4),
    'E33': (f27, 33, 5),
    'F33': (f27, 33, 6),
    'C36': (f28, 36, 3),
    'D36': (f28, 36, 4),
    'E36': (f28, 36, 5),
    'F36': (f28, 36, 6),
    'C37': (f29, 37, 3),
    'D37': (f29, 37, 4),
    'E37': (f29, 37, 5),
    'F37': (f29, 37, 6),
    'C38': (f30, 38, 3),
    'D38': (f30, 38, 4),
    'E38': (f30, 38, 5),
    'F38': (f30, 38, 6),
    'C40': (f31, 40, 3),
    'D40': (f31, 40, 4),
    'E40': (f31, 40, 5),
    'F40': (f31, 40, 6),
    'C41': (f32, 41, 3),
    'D41': (f32, 41, 4),
    'E41': (f32, 41, 5),
    'F41': (f32, 41, 6),
    'C43': (f33, 43, 3),
    'D43': (f33, 43, 4),
    'E43': (f33, 43, 5),
    'C44': (f34, 44, 3),
    'D44': (f34, 44, 4),
    'E44': (f34, 44, 5),
    'C45': (f35, 45, 3),
    'D45': (f35, 45, 4),
    'E45': (f35, 45, 5),
    'C46': (f36, 46, 3),
    'D46': (f36, 46, 4),
    'E46': (f36, 46, 5),
    'C47': (f37, 47, 3),
    'D47': (f37, 47, 4),
    'E47': (f37, 47, 5),
    'C49': (f38, 49, 3),
    'D49': (f39, 49, 4),
    'E49': (f39, 49, 5),
    'F49': (f40, 49, 6),
    'C50': (f41, 50, 3),
    'D50': (f42, 50, 4),
    'E50': (f42, 50, 5),
    'F50': (f43, 50, 6),
    'C51': (f41, 51, 3),
    'D51': (f42, 51, 4),
    'E51': (f42, 51, 5),
    'F51': (f43, 51, 6),
    'C52': (f44, 52, 3),
    'D52': (f45, 52, 4),
    'E52': (f45, 52, 5),
    'F52': (f46, 52, 6),
    'C53': (f47, 53, 3),
    'D53': (f48, 53, 4),
    'E53': (f48, 53, 5),
    'F53': (f49, 53, 6),
    'C54': (f50, 54, 3),
    'D54': (f50, 54, 4),
    'E54': (f50, 54, 5),
    'F54': (f50, 54, 6),
    'C55': (f51, 55, 3),
    'D55': (f51, 55, 4),
    'E55': (f51, 55, 5),
    'F55': (f51, 55, 6),
    'D57': (f52, 57, 4),
    'E57': (f53, 57, 5),
    'F57': (f54, 57, 6),
    'C58': (f55, 58, 3),
    'E58': (f52, 58, 5),
    'F58': (f53, 58, 6),
    'C59': (f56, 59, 3),
    'D59': (f55, 59, 4),
    'F59': (f52, 59, 6),
    'C60': (f57, 60, 3),
    'D60': (f56, 60, 4),
    'E60': (f55, 60, 5),
    'C62': (f58, 62, 3),
    'D62': (f58, 62, 4),
    'E62': (f58, 62, 5),
    'F62': (f58, 62, 6),
    'C63': (f59, 63, 3),
    'D63': (f59, 63, 4),
    'E63': (f59, 63, 5),
    'F63': (f59, 63, 6),
    'C64': (f60, 64, 3),
    'D64': (f60, 64, 4),
    'E64': (f60, 64, 5),
    'F64': (f60, 64, 6),
    'C65': (f61, 65, 3),
    'D65': (f61, 65, 4),
    'E65': (f61, 65, 5),
    'F65': (f61, 65, 6),
    'C66': (f62, 66, 3),
    'D66': (f62, 66, 4),
    'E66': (f62, 66, 5),
    'F66': (f62, 66, 6),
    'C67': (f63, 67, 3),
    'D67': (f63, 67, 4),
    'E67': (f63, 67, 5),
    'F67': (f63, 67, 6),
    'C68': (f64, 68, 3),
    'D68': (f64, 68, 4),
    'E68': (f64, 68, 5),
    'F68': (f64, 68, 6),
    'G68': (f65, 68, 7),
    'C70': (f66, 70, 3),
    'D70': (f66, 70, 4),
    'E70': (f66, 70, 5),
    'F70': (f66, 70, 6),
    'C71': (f67, 71, 3),
    'D71': (f67, 71, 4),
    'E71': (f67, 71, 5),
    'F71': (f67, 71, 6),
    'C72': (f68, 72, 3),
    'D72': (f68, 72, 4),
    'E72': (f68, 72, 5),
    'F72': (f68, 72, 6),
    'C75': (f69, 75, 3),
    'D75': (f69, 75, 4),
    'E75': (f69, 75, 5),
    'F75': (f69, 75, 6),
    'C76': (f70, 76, 3),
    'D76': (f70, 76, 4),
    'E76': (f70, 76, 5),
    'F76': (f70, 76, 6),
    'C77': (f71, 77, 3),
    'D77': (f71, 77, 4),
    'E77': (f71, 77, 5),
    'F77': (f71, 77, 6),
    'C78': (f72, 78, 3),
    'D78': (f72, 78, 4),
    'E78': (f72, 78, 5),
    'F78': (f72, 78, 6),
    'C79': (f73, 79, 3),
    'D79': (f73, 79, 4),
    'E79': (f73, 79, 5),
    'F79': (f73, 79, 6),
    'C80': (f74, 80, 3),
    'D80': (f74, 80, 4),
    'E80': (f74, 80, 5),
    'F80': (f74, 80, 6),
    'C81': (f75, 81, 3),
    'D81': (f75, 81, 4),
    'E81': (f75, 81, 5),
    'F81': (f75, 81, 6),
    'C82': (f76, 82, 3),
    'D82': (f77, 82, 4),
    'E82': (f77, 82, 5),
    'F82': (f78, 82, 6),
    'C83': (f79, 83, 3),
    'D83': (f79, 83, 4),
    'E83': (f79, 83, 5),
    'F83': (f79, 83, 6),
    'C84': (f80, 84, 3),
    'D84': (f80, 84, 4),
    'E84': (f80, 84, 5),
    'F84': (f80, 84, 6),
    'C85': (f81, 85, 3),
    'D85': (f81, 85, 4),
    'E85': (f81, 85, 5),
    'F85': (f81, 85, 6),
    'C86': (f82, 86, 3),
    'D86': (f82, 86, 4),
    'E86': (f82, 86, 5),
    'F86': (f82, 86, 6),
    'C87': (f83, 87, 3),
    'D87': (f83, 87, 4),
    'E87': (f83, 87, 5),
    'F87': (f83, 87, 6),
    'C91': (f84, 91, 3),
    'D91': (f85, 91, 4),
    'E91': (f86, 91, 5),
    'F91': (f87, 91, 6),
    'C92': (f88, 92, 3),
    'D92': (f89, 92, 4),
    'E92': (f90, 92, 5),
    'F92': (f87, 92, 6),
    'C93': (f91, 93, 3),
    'D93': (f92, 93, 4),
    'E93': (f93, 93, 5),
    'F93': (f87, 93, 6),
    'C94': (f94, 94, 3),
    'D94': (f95, 94, 4),
    'E94': (f96, 94, 5),
    'F94': (f87, 94, 6),
    'C95': (f97, 95, 3),
    'D95': (f98, 95, 4),
    'E95': (f99, 95, 5),
    'F95': (f87, 95, 6),
    'C97': (f100, 97, 3),
    'C98': (f101, 98, 3),
    'D98': (f102, 98, 4),
    'C99': (f103, 99, 3),
    'C100': (f104, 100, 3),
    'C101': (f105, 101, 3),
    'C102': (f106, 102, 3),
    'C103': (f107, 103, 3),
    'D106': (f108, 106, 4),
    'E106': (f109, 106, 5),
    'D107': (f110, 107, 4),
    'E107': (f109, 107, 5),
    'D108': (f111, 108, 4),
    'E108': (f109, 108, 5),
    'D109': (f112, 109, 4),
    'E109': (f109, 109, 5),
    'D110': (f113, 110, 4),
    'E110': (f109, 110, 5),
    'C114': (f114, 114, 3),
    'C115': (f115, 115, 3),
    'C116': (f116, 116, 3),
    'C117': (f117, 117, 3),
    'C118': (f118, 118, 3),
}
