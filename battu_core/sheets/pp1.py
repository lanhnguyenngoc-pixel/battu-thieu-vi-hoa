# -*- coding: utf-8 -*-
"""Module sheet PP1_DINH_TINH – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: 6.14.xlsx (SHA-256 b6c1126b9b23cc29eb1d8e3cf5b84b07a74bfc629e1cb6a6cc9f6b4d74f758cd).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'PP1_DINH_TINH'
SI = 9

def f0(E, r, c):
    # Ô C6
    # = IF(TRUNG_GIAN!$C$44=0,0,IF(COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!$C$46)>0,1,0))
    return (0.0 if X.truthy(X.eq(E.v(23,44,3),0.0)) else (1.0 if X.truthy(X.gt(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,46,3)),0.0)) else 0.0))

def f1(E, r, c):
    # Ô C7
    # = IF(AND(TRUNG_GIAN!C$52=1,TRUNG_GIAN!C$6=TRUNG_GIAN!$C$50,OR(AND(TRUNG_GIAN!C$73=1,TRUNG_GIAN!D$52=0))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(23,6,c),E.v(23,50,3)), X.OR(X.AND(X.eq(E.v(23,73,c),1.0), X.eq(E.v(23,52,c+1),0.0))))) else 0.0)

def f2(E, r, c):
    # Ô D7, E7
    # = IF(AND(TRUNG_GIAN!D$52=1,TRUNG_GIAN!D$6=TRUNG_GIAN!$C$50,OR(AND(TRUNG_GIAN!C$73=1,TRUNG_GIAN!C$52=0),AND(TRUNG_GIAN!D$73=1,TRUNG_GIAN!E$52=0))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(23,6,c),E.v(23,50,3)), X.OR(X.AND(X.eq(E.v(23,73,c-1),1.0), X.eq(E.v(23,52,c-1),0.0)), X.AND(X.eq(E.v(23,73,c),1.0), X.eq(E.v(23,52,c+1),0.0))))) else 0.0)

def f3(E, r, c):
    # Ô F7
    # = IF(AND(TRUNG_GIAN!F$52=1,TRUNG_GIAN!F$6=TRUNG_GIAN!$C$50,OR(AND(TRUNG_GIAN!E$73=1,TRUNG_GIAN!E$52=0))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(23,6,c),E.v(23,50,3)), X.OR(X.AND(X.eq(E.v(23,73,c-1),1.0), X.eq(E.v(23,52,c-1),0.0))))) else 0.0)

def f4(E, r, c):
    # Ô C8
    # = IF(AND(TRUNG_GIAN!$C$45="Tam hợp",SUM(C7:F7)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,45,3),'Tam hợp'), X.gt(X.SUM(E.rg(9,r-1,c,r-1,c+3)),0.0))) else 0.0)

def f5(E, r, c):
    # Ô C10, D10, E10
    # = IF(TRUNG_GIAN!C$57=1,CHOOSE(TRUNG_GIAN!C$55+1,3,6,-1,9,0),-1)
    return (X.CHOOSE(X.add(E.v(23,55,c),1.0), lambda: 3.0, lambda: 6.0, lambda: X.neg(1.0), lambda: 9.0, lambda: 0.0) if X.truthy(X.eq(E.v(23,57,c),1.0)) else X.neg(1.0))

def f6(E, r, c):
    # Ô C11
    # = IF(AND(TRUNG_GIAN!C$57=1,OR(AND(TRUNG_GIAN!D$6=C10,TRUNG_GIAN!D$73=1),AND(TRUNG_GIAN!C$6=C10,INDEX(M_XU,TRUNG_GIAN!C$6+1,TRUNG_GIAN!E$6+1)=1),AND(TRUN
    #   G_GIAN!D$6=C10,INDEX(M_XU,TRUNG_GIAN!D$6+1,TRUNG_GIAN!F$6+1)=1))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,57,c),1.0), X.OR(X.AND(X.eq(E.v(23,6,c+1),E.v(9,r-1,c)), X.eq(E.v(23,73,c+1),1.0)), X.AND(X.eq(E.v(23,6,c),E.v(9,r-1,c)), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c),1.0), X.add(E.v(23,6,c+2),1.0)),1.0)), X.AND(X.eq(E.v(23,6,c+1),E.v(9,r-1,c)), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c+1),1.0), X.add(E.v(23,6,c+3),1.0)),1.0))))) else 0.0)

def f7(E, r, c):
    # Ô D11
    # = IF(AND(TRUNG_GIAN!D$57=1,OR(AND(TRUNG_GIAN!D$6=D10,TRUNG_GIAN!C$73=1),AND(TRUNG_GIAN!E$6=D10,TRUNG_GIAN!E$73=1),AND(TRUNG_GIAN!D$6=D10,INDEX(M_XU,TRUN
    #   G_GIAN!D$6+1,TRUNG_GIAN!F$6+1)=1),AND(TRUNG_GIAN!E$6=D10,INDEX(M_XU,TRUNG_GIAN!E$6+1,TRUNG_GIAN!C$6+1)=1))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,57,c),1.0), X.OR(X.AND(X.eq(E.v(23,6,c),E.v(9,r-1,c)), X.eq(E.v(23,73,c-1),1.0)), X.AND(X.eq(E.v(23,6,c+1),E.v(9,r-1,c)), X.eq(E.v(23,73,c+1),1.0)), X.AND(X.eq(E.v(23,6,c),E.v(9,r-1,c)), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c),1.0), X.add(E.v(23,6,c+2),1.0)),1.0)), X.AND(X.eq(E.v(23,6,c+1),E.v(9,r-1,c)), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c+1),1.0), X.add(E.v(23,6,c-1),1.0)),1.0))))) else 0.0)

def f8(E, r, c):
    # Ô E11
    # = IF(AND(TRUNG_GIAN!E$57=1,OR(AND(TRUNG_GIAN!E$6=E10,TRUNG_GIAN!D$73=1),AND(TRUNG_GIAN!E$6=E10,INDEX(M_XU,TRUNG_GIAN!E$6+1,TRUNG_GIAN!C$6+1)=1),AND(TRUN
    #   G_GIAN!F$6=E10,INDEX(M_XU,TRUNG_GIAN!F$6+1,TRUNG_GIAN!D$6+1)=1))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,57,c),1.0), X.OR(X.AND(X.eq(E.v(23,6,c),E.v(9,r-1,c)), X.eq(E.v(23,73,c-1),1.0)), X.AND(X.eq(E.v(23,6,c),E.v(9,r-1,c)), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c),1.0), X.add(E.v(23,6,c-2),1.0)),1.0)), X.AND(X.eq(E.v(23,6,c+1),E.v(9,r-1,c)), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c+1),1.0), X.add(E.v(23,6,c-1),1.0)),1.0))))) else 0.0)

def f9(E, r, c):
    # Ô C12, D12, E12
    # = IF(AND(TRUNG_GIAN!C$57=1,C11=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,57,c),1.0), X.eq(E.v(9,r-1,c),0.0))) else 0.0)

def f10(E, r, c):
    # Ô C13, D13, E13
    # = IF(AND(C12=1,COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!C$55)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.gt(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,55,c)),0.0))) else 0.0)

def f11(E, r, c):
    # Ô C14
    # = IF(OR(AND(TRUNG_GIAN!C$52=1,$C$8=0),C12=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(9,r-2,c),1.0))) else 0.0)

def f12(E, r, c):
    # Ô D14, E14
    # = IF(OR(AND(TRUNG_GIAN!D$52=1,$C$8=0),C12=1,D12=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(9,r-2,c-1),1.0), X.eq(E.v(9,r-2,c),1.0))) else 0.0)

def f13(E, r, c):
    # Ô F14
    # = IF(OR(AND(TRUNG_GIAN!F$52=1,$C$8=0),E12=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(9,r-2,c-1),1.0))) else 0.0)

def f14(E, r, c):
    # Ô C15
    # = IF(AND(TRUNG_GIAN!C$52=1,$C$8=0),TRUNG_GIAN!$C$46,IF(C12=1,TRUNG_GIAN!C$55,-1))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0))) else (E.v(23,55,c) if X.truthy(X.eq(E.v(9,r-3,c),1.0)) else X.neg(1.0)))

def f15(E, r, c):
    # Ô D15, E15
    # = IF(AND(TRUNG_GIAN!D$52=1,$C$8=0),TRUNG_GIAN!$C$46,IF(C12=1,TRUNG_GIAN!C$55,IF(D12=1,TRUNG_GIAN!D$55,-1)))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0))) else (E.v(23,55,c-1) if X.truthy(X.eq(E.v(9,r-3,c-1),1.0)) else (E.v(23,55,c) if X.truthy(X.eq(E.v(9,r-3,c),1.0)) else X.neg(1.0))))

def f16(E, r, c):
    # Ô F15
    # = IF(AND(TRUNG_GIAN!F$52=1,$C$8=0),TRUNG_GIAN!$C$46,IF(E12=1,TRUNG_GIAN!E$55,-1))
    return (E.v(23,46,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0))) else (E.v(23,55,c-1) if X.truthy(X.eq(E.v(9,r-3,c-1),1.0)) else X.neg(1.0)))

def f17(E, r, c):
    # Ô C16
    # = IF(AND(TRUNG_GIAN!C$52=1,$C$8=0),$C$6,IF(C12=1,C13,0))
    return (E.v(9,6,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0))) else (E.v(9,r-3,c) if X.truthy(X.eq(E.v(9,r-4,c),1.0)) else 0.0))

def f18(E, r, c):
    # Ô D16, E16
    # = IF(AND(TRUNG_GIAN!D$52=1,$C$8=0),$C$6,IF(C12=1,C13,IF(D12=1,D13,0)))
    return (E.v(9,6,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0))) else (E.v(9,r-3,c-1) if X.truthy(X.eq(E.v(9,r-4,c-1),1.0)) else (E.v(9,r-3,c) if X.truthy(X.eq(E.v(9,r-4,c),1.0)) else 0.0)))

def f19(E, r, c):
    # Ô F16
    # = IF(AND(TRUNG_GIAN!F$52=1,$C$8=0),$C$6,IF(E12=1,E13,0))
    return (E.v(9,6,3) if X.truthy(X.AND(X.eq(E.v(23,52,c),1.0), X.eq(E.v(9,8,3),0.0))) else (E.v(9,r-3,c-1) if X.truthy(X.eq(E.v(9,r-4,c-1),1.0)) else 0.0))

def f20(E, r, c):
    # Ô C18, D18, E18
    # = IF(AND(TRUNG_GIAN!C$64>=0,C14=0,D14=0),1,0)
    return (1.0 if X.truthy(X.AND(X.ge(E.v(23,64,c),0.0), X.eq(E.v(9,r-4,c),0.0), X.eq(E.v(9,r-4,c+1),0.0))) else 0.0)

def f21(E, r, c):
    # Ô C19, C25
    # = C18
    return E.v(9,r-1,c)

def f22(E, r, c):
    # Ô D19, E19, D25, E25
    # = C18+D18
    return X.add(E.v(9,r-1,c-1),E.v(9,r-1,c))

def f23(E, r, c):
    # Ô F19, F25
    # = E18
    return E.v(9,r-1,c-1)

def f24(E, r, c):
    # Ô C20, D20, E20
    # = IF(AND(C18=1,OR(C19>1,D19>1)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-2,c),1.0), X.OR(X.gt(E.v(9,r-1,c),1.0), X.gt(E.v(9,r-1,c+1),1.0)))) else 0.0)

def f25(E, r, c):
    # Ô C21, D21, E21
    # = IF(AND(C18=1,C20=0,COUNTIF(TRUNG_GIAN!$C$7:$F$7,TRUNG_GIAN!C$64)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.eq(E.v(9,r-1,c),0.0), X.gt(X.COUNTIF(E.rg(23,7,3,7,6), E.v(23,64,c)),0.0))) else 0.0)

def f26(E, r, c):
    # Ô C22
    # = IF(OR(C21=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-1,c),1.0))) else 0.0)

def f27(E, r, c):
    # Ô D22, E22
    # = IF(OR(C21=1,D21=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-1,c-1),1.0), X.eq(E.v(9,r-1,c),1.0))) else 0.0)

def f28(E, r, c):
    # Ô F22
    # = IF(OR(E21=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-1,c-1),1.0))) else 0.0)

def f29(E, r, c):
    # Ô C24, D24, E24
    # = IF(AND(TRUNG_GIAN!C$73=1,C14=0,D14=0,C22=0,D22=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,73,c),1.0), X.eq(E.v(9,r-10,c),0.0), X.eq(E.v(9,r-10,c+1),0.0), X.eq(E.v(9,r-2,c),0.0), X.eq(E.v(9,r-2,c+1),0.0))) else 0.0)

def f30(E, r, c):
    # Ô C26, D26, E26
    # = IF(AND(C18=1,OR(C21=1,AND(C25=0,D25=0))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-8,c),1.0), X.OR(X.eq(E.v(9,r-5,c),1.0), X.AND(X.eq(E.v(9,r-1,c),0.0), X.eq(E.v(9,r-1,c+1),0.0))))) else 0.0)

def f31(E, r, c):
    # Ô C27
    # = IF(AND(C14=1,C16=1),C15,IF(AND(C26=1,C21=1),TRUNG_GIAN!C$64,-1))
    return (E.v(9,r-12,c) if X.truthy(X.AND(X.eq(E.v(9,r-13,c),1.0), X.eq(E.v(9,r-11,c),1.0))) else (E.v(23,64,c) if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.eq(E.v(9,r-6,c),1.0))) else X.neg(1.0)))

def f32(E, r, c):
    # Ô D27, E27
    # = IF(AND(D14=1,D16=1),D15,IF(AND(C26=1,C21=1),TRUNG_GIAN!C$64,IF(AND(D26=1,D21=1),TRUNG_GIAN!D$64,-1)))
    return (E.v(9,r-12,c) if X.truthy(X.AND(X.eq(E.v(9,r-13,c),1.0), X.eq(E.v(9,r-11,c),1.0))) else (E.v(23,64,c-1) if X.truthy(X.AND(X.eq(E.v(9,r-1,c-1),1.0), X.eq(E.v(9,r-6,c-1),1.0))) else (E.v(23,64,c) if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.eq(E.v(9,r-6,c),1.0))) else X.neg(1.0))))

def f33(E, r, c):
    # Ô F27
    # = IF(AND(F14=1,F16=1),F15,IF(AND(E26=1,E21=1),TRUNG_GIAN!E$64,-1))
    return (E.v(9,r-12,c) if X.truthy(X.AND(X.eq(E.v(9,r-13,c),1.0), X.eq(E.v(9,r-11,c),1.0))) else (E.v(23,64,c-1) if X.truthy(X.AND(X.eq(E.v(9,r-1,c-1),1.0), X.eq(E.v(9,r-6,c-1),1.0))) else X.neg(1.0)))

def f34(E, r, c):
    # Ô C28
    # = IF(OR(AND(C14=1,C16=0),AND(C26=1,C21=0,C20=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-14,c),1.0), X.eq(E.v(9,r-12,c),0.0)), X.AND(X.eq(E.v(9,r-2,c),1.0), X.eq(E.v(9,r-7,c),0.0), X.eq(E.v(9,r-8,c),0.0)))) else 0.0)

def f35(E, r, c):
    # Ô D28, E28
    # = IF(OR(AND(D14=1,D16=0),AND(C26=1,C21=0,C20=0),AND(D26=1,D21=0,D20=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-14,c),1.0), X.eq(E.v(9,r-12,c),0.0)), X.AND(X.eq(E.v(9,r-2,c-1),1.0), X.eq(E.v(9,r-7,c-1),0.0), X.eq(E.v(9,r-8,c-1),0.0)), X.AND(X.eq(E.v(9,r-2,c),1.0), X.eq(E.v(9,r-7,c),0.0), X.eq(E.v(9,r-8,c),0.0)))) else 0.0)

def f36(E, r, c):
    # Ô F28
    # = IF(OR(AND(F14=1,F16=0),AND(E26=1,E21=0,E20=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-14,c),1.0), X.eq(E.v(9,r-12,c),0.0)), X.AND(X.eq(E.v(9,r-2,c-1),1.0), X.eq(E.v(9,r-7,c-1),0.0), X.eq(E.v(9,r-8,c-1),0.0)))) else 0.0)

def f37(E, r, c):
    # Ô C29, D29, E29, F29
    # = IF(C25>0,1,0)
    return (1.0 if X.truthy(X.gt(E.v(9,r-4,c),0.0)) else 0.0)

def f38(E, r, c):
    # Ô C30
    # = IF(OR(AND(C24=1,TRUNG_GIAN!C$74=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-6,c),1.0), X.eq(E.v(23,74,c),1.0)))) else 0.0)

def f39(E, r, c):
    # Ô D30, E30
    # = IF(OR(AND(C24=1,TRUNG_GIAN!C$74=1),AND(D24=1,TRUNG_GIAN!D$74=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-6,c-1),1.0), X.eq(E.v(23,74,c-1),1.0)), X.AND(X.eq(E.v(9,r-6,c),1.0), X.eq(E.v(23,74,c),1.0)))) else 0.0)

def f40(E, r, c):
    # Ô F30
    # = IF(OR(AND(E24=1,TRUNG_GIAN!E$74=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-6,c-1),1.0), X.eq(E.v(23,74,c-1),1.0)))) else 0.0)

def f41(E, r, c):
    # Ô C33, D33, E33
    # = IF(TRUNG_GIAN!C$76>=0,1,0)
    return (1.0 if X.truthy(X.ge(E.v(23,76,c),0.0)) else 0.0)

def f42(E, r, c):
    # Ô C34
    # = IF(AND(PP2_DINH_LUONG!C$47=1,NOT(OR(TRUNG_GIAN!C$21=PP2_DINH_LUONG!C$28,TRUNG_GIAN!C$22=PP2_DINH_LUONG!C$28,TRUNG_GIAN!C$23=PP2_DINH_LUONG!C$28))),0,P
    #   P2_DINH_LUONG!C$47)
    return (0.0 if X.truthy(X.AND(X.eq(E.v(10,47,c),1.0), X.NOT(X.OR(X.eq(E.v(23,21,c),E.v(10,28,c)), X.eq(E.v(23,22,c),E.v(10,28,c)), X.eq(E.v(23,23,c),E.v(10,28,c)))))) else E.v(10,47,c))

def f43(E, r, c):
    # Ô D34, E34
    # = PP2_DINH_LUONG!D$47
    return E.v(10,47,c)

def f44(E, r, c):
    # Ô C35
    # = IF(C33=0,"",INDEX(CAN_TEN,TRUNG_GIAN!C$5+1)&"–"&INDEX(CAN_TEN,TRUNG_GIAN!D$5+1)&IF(TRUNG_GIAN!C$78=1,": tranh hợp, không hóa",IF(C34=1,": HÓA "&INDEX(
    #   HANH_TEN,PP2_DINH_LUONG!C$52+1),IF(PP2_DINH_LUONG!C$47=1,": hợp không hóa (PP1) – đủ điều kiện P2 tr.336-337 nhưng chi năm "&TRUNG_GIAN!C$11&" không c
    #   ó căn của hóa thần "&INDEX(HANH_TEN,PP2_DINH_LUONG!C$28+1)&" (P1 tr.62; G81) – giữ hành gốc; khắc gần, can bị khắc còn 2/3 lực (P2 tr.339)",": hợp khô
    #   ng hóa – khắc gần, can bị khắc còn 2/3 lực, can khắc không hao (P2 tr.339)"))))
    return ('' if X.truthy(X.eq(E.v(9,r-2,c),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,c),1.0)),'–'),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,c+1),1.0))),(': tranh hợp, không hóa' if X.truthy(X.eq(E.v(23,78,c),1.0)) else (X.cat(': HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,c),1.0))) if X.truthy(X.eq(E.v(9,r-1,c),1.0)) else (X.cat(X.cat(X.cat(X.cat(': hợp không hóa (PP1) – đủ điều kiện P2 tr.336-337 nhưng chi năm ',E.v(23,11,c)),' không có căn của hóa thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,28,c),1.0))),' (P1 tr.62; G81) – giữ hành gốc; khắc gần, can bị khắc còn 2/3 lực (P2 tr.339)') if X.truthy(X.eq(E.v(10,47,c),1.0)) else ': hợp không hóa – khắc gần, can bị khắc còn 2/3 lực, can khắc không hao (P2 tr.339)')))))

def f45(E, r, c):
    # Ô D35, E35
    # = IF(D33=0,"",INDEX(CAN_TEN,TRUNG_GIAN!D$5+1)&"–"&INDEX(CAN_TEN,TRUNG_GIAN!E$5+1)&IF(TRUNG_GIAN!D$78=1,": tranh hợp, không hóa",IF(D34=1,": HÓA "&INDEX(
    #   HANH_TEN,PP2_DINH_LUONG!D$52+1),": hợp không hóa – khắc gần, can bị khắc còn 2/3 lực, can khắc không hao (P2 tr.339)")))
    return ('' if X.truthy(X.eq(E.v(9,r-2,c),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,c),1.0)),'–'),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,c+1),1.0))),(': tranh hợp, không hóa' if X.truthy(X.eq(E.v(23,78,c),1.0)) else (X.cat(': HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,c),1.0))) if X.truthy(X.eq(E.v(9,r-1,c),1.0)) else ': hợp không hóa – khắc gần, can bị khắc còn 2/3 lực, can khắc không hao (P2 tr.339)'))))

def f46(E, r, c):
    # Ô C36
    # = IF(OR(AND(C33=1,C34=0,OR(TRUNG_GIAN!C$78=1,MOD(TRUNG_GIAN!D$7+2,5)=TRUNG_GIAN!C$7))),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-3,c),1.0), X.eq(E.v(9,r-2,c),0.0), X.OR(X.eq(E.v(23,78,c),1.0), X.eq(X.MOD(X.add(E.v(23,7,c+1),2.0), 5.0),E.v(23,7,c)))))) else 0.0)

def f47(E, r, c):
    # Ô D36
    # = IF(OR(AND(C33=1,C34=0,OR(TRUNG_GIAN!C$78=1,MOD(TRUNG_GIAN!C$7+2,5)=TRUNG_GIAN!D$7)),AND(D33=1,D34=0,OR(TRUNG_GIAN!D$78=1,MOD(TRUNG_GIAN!E$7+2,5)=TRUNG
    #   _GIAN!D$7))),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-3,c-1),1.0), X.eq(E.v(9,r-2,c-1),0.0), X.OR(X.eq(E.v(23,78,c-1),1.0), X.eq(X.MOD(X.add(E.v(23,7,c-1),2.0), 5.0),E.v(23,7,c)))), X.AND(X.eq(E.v(9,r-3,c),1.0), X.eq(E.v(9,r-2,c),0.0), X.OR(X.eq(E.v(23,78,c),1.0), X.eq(X.MOD(X.add(E.v(23,7,c+1),2.0), 5.0),E.v(23,7,c)))))) else 0.0)

def f48(E, r, c):
    # Ô E36, E41, D45, D46, E48, E49, E57, E58
    # = 0
    return 0.0

def f49(E, r, c):
    # Ô F36
    # = IF(OR(AND(E33=1,E34=0,OR(TRUNG_GIAN!E$78=1,MOD(TRUNG_GIAN!E$7+2,5)=TRUNG_GIAN!F$7))),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-3,c-1),1.0), X.eq(E.v(9,r-2,c-1),0.0), X.OR(X.eq(E.v(23,78,c-1),1.0), X.eq(X.MOD(X.add(E.v(23,7,c-1),2.0), 5.0),E.v(23,7,c)))))) else 0.0)

def f50(E, r, c):
    # Ô C37, D37, E37, F37
    # = IF(C39=1,PP2_DINH_LUONG!C$52,TRUNG_GIAN!C$7)
    return (E.v(10,52,c) if X.truthy(X.eq(E.v(9,r+2,c),1.0)) else E.v(23,7,c))

def f51(E, r, c):
    # Ô C38
    # = IF(OR(AND(C33=1,C34=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-5,c),1.0), X.eq(E.v(9,r-4,c),0.0)))) else 0.0)

def f52(E, r, c):
    # Ô D38, E38
    # = IF(OR(AND(C33=1,C34=0),AND(D33=1,D34=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-5,c-1),1.0), X.eq(E.v(9,r-4,c-1),0.0)), X.AND(X.eq(E.v(9,r-5,c),1.0), X.eq(E.v(9,r-4,c),0.0)))) else 0.0)

def f53(E, r, c):
    # Ô F38
    # = IF(OR(AND(E33=1,E34=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(9,r-5,c-1),1.0), X.eq(E.v(9,r-4,c-1),0.0)))) else 0.0)

def f54(E, r, c):
    # Ô C39
    # = IF(OR(C34=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-5,c),1.0))) else 0.0)

def f55(E, r, c):
    # Ô D39, E39
    # = IF(OR(C34=1,D34=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-5,c-1),1.0), X.eq(E.v(9,r-5,c),1.0))) else 0.0)

def f56(E, r, c):
    # Ô F39
    # = IF(OR(E34=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-5,c-1),1.0))) else 0.0)

def f57(E, r, c):
    # Ô C40
    # = IF(AND(C38=0,C39=0,OR(AND(D38=0,MOD(D37+2,5)=C37))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-2,c),0.0), X.eq(E.v(9,r-1,c),0.0), X.OR(X.AND(X.eq(E.v(9,r-2,c+1),0.0), X.eq(X.MOD(X.add(E.v(9,r-3,c+1),2.0), 5.0),E.v(9,r-3,c)))))) else 0.0)

def f58(E, r, c):
    # Ô D40, E40
    # = IF(AND(D38=0,D39=0,OR(AND(C38=0,MOD(C37+2,5)=D37),AND(E38=0,MOD(E37+2,5)=D37))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-2,c),0.0), X.eq(E.v(9,r-1,c),0.0), X.OR(X.AND(X.eq(E.v(9,r-2,c-1),0.0), X.eq(X.MOD(X.add(E.v(9,r-3,c-1),2.0), 5.0),E.v(9,r-3,c))), X.AND(X.eq(E.v(9,r-2,c+1),0.0), X.eq(X.MOD(X.add(E.v(9,r-3,c+1),2.0), 5.0),E.v(9,r-3,c)))))) else 0.0)

def f59(E, r, c):
    # Ô F40
    # = IF(AND(F38=0,F39=0,OR(AND(E38=0,MOD(E37+2,5)=F37))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-2,c),0.0), X.eq(E.v(9,r-1,c),0.0), X.OR(X.AND(X.eq(E.v(9,r-2,c-1),0.0), X.eq(X.MOD(X.add(E.v(9,r-3,c-1),2.0), 5.0),E.v(9,r-3,c)))))) else 0.0)

def f60(E, r, c):
    # Ô C41, D41, F41
    # = IF(AND(TRUNG_GIAN!C$30<=4,OR(TRUNG_GIAN!C$9=TRUNG_GIAN!C$7,MOD(TRUNG_GIAN!C$9+1,5)=TRUNG_GIAN!C$7)),1,0)
    return (1.0 if X.truthy(X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(23,9,c),E.v(23,7,c)), X.eq(X.MOD(X.add(E.v(23,9,c),1.0), 5.0),E.v(23,7,c))))) else 0.0)

def f61(E, r, c):
    # Ô C43
    # = IF(TRUNG_GIAN!$D$28<=4,1,0)
    return (1.0 if X.truthy(X.le(E.v(23,28,4),4.0)) else 0.0)

def f62(E, r, c):
    # Ô D43
    # = "Nhật can "&TRUNG_GIAN!$E$10&" ở "&TRUNG_GIAN!$D$29&" tại tháng "&TRUNG_GIAN!$D$11
    return X.cat(X.cat(X.cat(X.cat(X.cat('Nhật can ',E.v(23,10,5)),' ở '),E.v(23,29,4)),' tại tháng '),E.v(23,11,4))

def f63(E, r, c):
    # Ô C45, E45, F45
    # = IF(C27>=0,IF(C27=TRUNG_GIAN!$E$7,1,0),IF(C28=1,0,IF(OR(AND(OR(TRUNG_GIAN!C$28=3,TRUNG_GIAN!C$28=4),TRUNG_GIAN!C$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$
    #   8=0,OR(TRUNG_GIAN!C$28=0,TRUNG_GIAN!C$28=8)),AND(P_DUKHI=1,C29=0,OR(AND(TRUNG_GIAN!$E$5=1,TRUNG_GIAN!C$6=4),AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!C$6=7),AN
    #   D(TRUNG_GIAN!$E$5=7,TRUNG_GIAN!C$6=10),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!C$6=1)))),1,0)))
    return ((1.0 if X.truthy(X.eq(E.v(9,r-18,c),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(9,r-18,c),0.0)) else (0.0 if X.truthy(X.eq(E.v(9,r-17,c),1.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,28,c),3.0), X.eq(E.v(23,28,c),4.0)), X.eq(E.v(23,21,c),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,c),0.0), X.eq(E.v(23,28,c),8.0))), X.AND(X.eq(E.v(24,76,6),1.0), X.eq(E.v(9,r-16,c),0.0), X.OR(X.AND(X.eq(E.v(23,5,5),1.0), X.eq(E.v(23,6,c),4.0)), X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,c),7.0)), X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,6,c),10.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,c),1.0)))))) else 0.0)))

def f64(E, r, c):
    # Ô C46
    # = IF(C45=1,IF(C27>=0,IF(ABS(0-2)<=1,1,0),IF(AND(OR(C29=0,C30=1),ABS(0-2)<=1),1,0)),0)
    return (((1.0 if X.truthy(X.le(X.ABS(X.sub(0.0,2.0)),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-19,c),0.0)) else (1.0 if X.truthy(X.AND(X.OR(X.eq(E.v(9,r-17,c),0.0), X.eq(E.v(9,r-16,c),1.0)), X.le(X.ABS(X.sub(0.0,2.0)),1.0))) else 0.0)) if X.truthy(X.eq(E.v(9,r-1,c),1.0)) else 0.0)

def f65(E, r, c):
    # Ô E46
    # = IF(E45=1,IF(E27>=0,IF(ABS(2-2)<=1,1,0),IF(AND(OR(E29=0,E30=1),ABS(2-2)<=1),1,0)),0)
    return (((1.0 if X.truthy(X.le(X.ABS(X.sub(2.0,2.0)),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-19,c),0.0)) else (1.0 if X.truthy(X.AND(X.OR(X.eq(E.v(9,r-17,c),0.0), X.eq(E.v(9,r-16,c),1.0)), X.le(X.ABS(X.sub(2.0,2.0)),1.0))) else 0.0)) if X.truthy(X.eq(E.v(9,r-1,c),1.0)) else 0.0)

def f66(E, r, c):
    # Ô F46
    # = IF(F45=1,IF(F27>=0,IF(ABS(3-2)<=1,1,0),IF(AND(OR(F29=0,F30=1),ABS(3-2)<=1),1,0)),0)
    return (((1.0 if X.truthy(X.le(X.ABS(X.sub(3.0,2.0)),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-19,c),0.0)) else (1.0 if X.truthy(X.AND(X.OR(X.eq(E.v(9,r-17,c),0.0), X.eq(E.v(9,r-16,c),1.0)), X.le(X.ABS(X.sub(3.0,2.0)),1.0))) else 0.0)) if X.truthy(X.eq(E.v(9,r-1,c),1.0)) else 0.0)

def f67(E, r, c):
    # Ô C47
    # = "Nhật can "&TRUNG_GIAN!$E$10&" tại các chi – "&TRUNG_GIAN!$C$11&" (năm): "&TRUNG_GIAN!$C$29&"; "&TRUNG_GIAN!$E$11&" (ngày): "&TRUNG_GIAN!$E$29&"; "&TR
    #   UNG_GIAN!$F$11&" (giờ): "&TRUNG_GIAN!$F$29&". "&IF(SUM($C$45:$F$45)>0,"→ ĐẮC ĐỊA ở "&TRIM(IF(PP1_DINH_TINH!$C$45=1,TRUNG_GIAN!$C$11&" (năm) ","")&IF(P
    #   P1_DINH_TINH!$E$45=1,TRUNG_GIAN!$E$11&" (ngày) ","")&IF(PP1_DINH_TINH!$F$45=1,TRUNG_GIAN!$F$11&" (giờ) ",""))&".","→ KHÔNG gặp Trường sinh (can dương)
    #   , Lộc, Kình dương (bản khí là tỉ/kiếp), Mộ khố (can dương)"&IF(P_DUKHI=1,", dư khí (can âm: Ất–Thìn, Đinh–Mùi, Tân–Tuất, Quý–Sửu – mở rộng Tử Bình Châ
    #   n Thuyên, đang BẬT)","")&" nên KHÔNG đắc địa.")&IF(TRIM($AA$47&$AB$47&$AC$47&$AD$47)="",""," Tàng can cùng hành nhật chủ: "&TRIM($AA$47&$AB$47&$AC$47&
    #   $AD$47)&" Đây là THÔNG CĂN / có gốc (xem THONG_CAN; được cộng độ ở PP2), KHÔNG phải đắc địa: sách chỉ gọi đắc địa khi nhật can gặp Trường sinh, Lộc, K
    #   ình dương, Mộ khố ở chi (P1 tr.70); "&IF(P_DUKHI=1,"công cụ đang BẬT phần mở rộng dư khí của can âm (Tử Bình Chân Thuyên ch.6); ","phần mở rộng dư khí
    #    của can âm (Tử Bình Chân Thuyên) đang TẮT – CSDL P_DUKHI; ")&"tỉ k"&"iếp tàng trong chi cũng không tính vào ④ được trợ (chỉ tính thiên can lộ).")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật can ',E.v(23,10,5)),' tại các chi – '),E.v(23,11,3)),' (năm): '),E.v(23,29,3)),'; '),E.v(23,11,5)),' (ngày): '),E.v(23,29,5)),'; '),E.v(23,11,6)),' (giờ): '),E.v(23,29,6)),'. '),(X.cat(X.cat('→ ĐẮC ĐỊA ở ',X.TRIM(X.cat(X.cat((X.cat(E.v(23,11,3),' (năm) ') if X.truthy(X.eq(E.v(9,45,3),1.0)) else ''),(X.cat(E.v(23,11,5),' (ngày) ') if X.truthy(X.eq(E.v(9,45,5),1.0)) else '')),(X.cat(E.v(23,11,6),' (giờ) ') if X.truthy(X.eq(E.v(9,45,6),1.0)) else '')))),'.') if X.truthy(X.gt(X.SUM(E.rg(9,45,3,45,6)),0.0)) else X.cat(X.cat('→ KHÔNG gặp Trường sinh (can dương), Lộc, Kình dương (bản khí là tỉ/kiếp), Mộ khố (can dương)',(', dư khí (can âm: Ất–Thìn, Đinh–Mùi, Tân–Tuất, Quý–Sửu – mở rộng Tử Bình Chân Thuyên, đang BẬT)' if X.truthy(X.eq(E.v(24,76,6),1.0)) else '')),' nên KHÔNG đắc địa.'))),('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(E.v(9,47,27),E.v(9,47,28)),E.v(9,47,29)),E.v(9,47,30))),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(' Tàng can cùng hành nhật chủ: ',X.TRIM(X.cat(X.cat(X.cat(E.v(9,47,27),E.v(9,47,28)),E.v(9,47,29)),E.v(9,47,30)))),' Đây là THÔNG CĂN / có gốc (xem THONG_CAN; được cộng độ ở PP2), KHÔNG phải đắc địa: sách chỉ gọi đắc địa khi nhật can gặp Trường sinh, Lộc, Kình dương, Mộ khố ở chi (P1 tr.70); '),('công cụ đang BẬT phần mở rộng dư khí của can âm (Tử Bình Chân Thuyên ch.6); ' if X.truthy(X.eq(E.v(24,76,6),1.0)) else 'phần mở rộng dư khí của can âm (Tử Bình Chân Thuyên) đang TẮT – CSDL P_DUKHI; ')),'tỉ k'),'iếp tàng trong chi cũng không tính vào ④ được trợ (chỉ tính thiên can lộ).')))

def f68(E, r, c):
    # Ô AA47
    # = IF(TRUNG_GIAN!$C$21=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" trong "&TRUNG_GIAN!$C$11&" (năm, bản khí)"&IF(PP1_DINH_TINH!$C$27>=0,IF(PP1_DI
    #   NH_TINH!$C$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1)&" trong hợp cụ
    #   c, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$C$22=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" trong "&TRUNG_GIAN!$C$11&" (năm, trun
    #   g khí)"&IF(PP1_DINH_TINH!$C$27>=0,IF(PP1_DINH_TINH!$C$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HAN
    #   H_TEN,PP1_DINH_TINH!$C$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$C$23=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$C$19
    #   +1)&" trong "&TRUNG_GIAN!$C$11&" (năm, dư khí)"&IF(PP1_DINH_TINH!$C$27>=0,IF(PP1_DINH_TINH!$C$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (l
    #   à gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$C$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")
    return X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)),' trong '),E.v(23,11,3)),' (năm, bản khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,3),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,3),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,21,3),E.v(23,7,5))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)),' trong '),E.v(23,11,3)),' (năm, trung khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,3),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,3),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,22,3),E.v(23,7,5))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)),' trong '),E.v(23,11,3)),' (năm, dư khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,3),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,3),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,3),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,23,3),E.v(23,7,5))) else ''))

def f69(E, r, c):
    # Ô AB47
    # = IF(TRUNG_GIAN!$D$21=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" trong "&TRUNG_GIAN!$D$11&" (tháng, bản khí)"&IF(PP1_DINH_TINH!$D$27>=0,IF(PP1_
    #   DINH_TINH!$D$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1)&" trong hợp 
    #   cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$D$22=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" trong "&TRUNG_GIAN!$D$11&" (tháng, 
    #   trung khí)"&IF(PP1_DINH_TINH!$D$27>=0,IF(PP1_DINH_TINH!$D$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX
    #   (HANH_TEN,PP1_DINH_TINH!$D$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$D$23=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$
    #   D$19+1)&" trong "&TRUNG_GIAN!$D$11&" (tháng, dư khí)"&IF(PP1_DINH_TINH!$D$27>=0,IF(PP1_DINH_TINH!$D$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật 
    #   chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$D$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")
    return X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)),' trong '),E.v(23,11,4)),' (tháng, bản khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,4),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,4),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,21,4),E.v(23,7,5))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)),' trong '),E.v(23,11,4)),' (tháng, trung khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,4),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,4),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,22,4),E.v(23,7,5))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)),' trong '),E.v(23,11,4)),' (tháng, dư khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,4),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,4),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,4),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,23,4),E.v(23,7,5))) else ''))

def f70(E, r, c):
    # Ô AC47
    # = IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" trong "&TRUNG_GIAN!$E$11&" (ngày, bản khí)"&IF(PP1_DINH_TINH!$E$27>=0,IF(PP1_D
    #   INH_TINH!$E$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1)&" trong hợp c
    #   ục, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" trong "&TRUNG_GIAN!$E$11&" (ngày, tr
    #   ung khí)"&IF(PP1_DINH_TINH!$E$27>=0,IF(PP1_DINH_TINH!$E$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(H
    #   ANH_TEN,PP1_DINH_TINH!$E$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$E$
    #   19+1)&" trong "&TRUNG_GIAN!$E$11&" (ngày, dư khí)"&IF(PP1_DINH_TINH!$E$27>=0,IF(PP1_DINH_TINH!$E$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ
    #    (là gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$E$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")
    return X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)),' trong '),E.v(23,11,5)),' (ngày, bản khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,5),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,5),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,5))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)),' trong '),E.v(23,11,5)),' (ngày, trung khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,5),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,5),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,5))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)),' trong '),E.v(23,11,5)),' (ngày, dư khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,5),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,5),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,5),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,5))) else ''))

def f71(E, r, c):
    # Ô AD47
    # = IF(TRUNG_GIAN!$F$21=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" trong "&TRUNG_GIAN!$F$11&" (giờ, bản khí)"&IF(PP1_DINH_TINH!$F$27>=0,IF(PP1_DI
    #   NH_TINH!$F$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1)&" trong hợp cụ
    #   c, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$F$22=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" trong "&TRUNG_GIAN!$F$11&" (giờ, trun
    #   g khí)"&IF(PP1_DINH_TINH!$F$27>=0,IF(PP1_DINH_TINH!$F$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (là gốc)"," – nhưng chi đã hóa "&INDEX(HAN
    #   H_TEN,PP1_DINH_TINH!$F$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")&IF(TRUNG_GIAN!$F$23=TRUNG_GIAN!$E$7,INDEX(CAN_TEN,TRUNG_GIAN!$F$19
    #   +1)&" trong "&TRUNG_GIAN!$F$11&" (giờ, dư khí)"&IF(PP1_DINH_TINH!$F$27>=0,IF(PP1_DINH_TINH!$F$27=TRUNG_GIAN!$E$7," – chi đã hóa thành hành nhật chủ (l
    #   à gốc)"," – nhưng chi đã hóa "&INDEX(HANH_TEN,PP1_DINH_TINH!$F$27+1)&" trong hợp cục, KHÔNG còn trợ nhật chủ"),"")&"; ","")
    return X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)),' trong '),E.v(23,11,6)),' (giờ, bản khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,6),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,6),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,21,6),E.v(23,7,5))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)),' trong '),E.v(23,11,6)),' (giờ, trung khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,6),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,6),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,22,6),E.v(23,7,5))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)),' trong '),E.v(23,11,6)),' (giờ, dư khí)'),((' – chi đã hóa thành hành nhật chủ (là gốc)' if X.truthy(X.eq(E.v(9,27,6),E.v(23,7,5))) else X.cat(X.cat(' – nhưng chi đã hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,27,6),1.0))),' trong hợp cục, KHÔNG còn trợ nhật chủ')) if X.truthy(X.ge(E.v(9,27,6),0.0)) else '')),'; ') if X.truthy(X.eq(E.v(23,23,6),E.v(23,7,5))) else ''))

def f72(E, r, c):
    # Ô C48, D48, F48
    # = IF(C37=MOD(TRUNG_GIAN!$E$7+4,5),1,0)
    return (1.0 if X.truthy(X.eq(E.v(9,r-11,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))) else 0.0)

def f73(E, r, c):
    # Ô C49
    # = IF(AND(C48=1,AND(TRUNG_GIAN!C$30<=4,OR(C41=1,AND(C36=0,C40=0)),NOT(OR(THONG_CAN!$C$69="Không có gốc",THONG_CAN!$C$69="Gốc dễ bị nhổ",THONG_CAN!$C$69="
    #   Gốc bị hợp (đứng độc lập)")),ABS(0-2)<2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(9,r-8,c),1.0), X.AND(X.eq(E.v(9,r-13,c),0.0), X.eq(E.v(9,r-9,c),0.0))), X.NOT(X.OR(X.eq(E.v(11,69,3),'Không có gốc'), X.eq(E.v(11,69,3),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,3),'Gốc bị hợp (đứng độc lập)'))), X.lt(X.ABS(X.sub(0.0,2.0)),2.0)))) else 0.0)

def f74(E, r, c):
    # Ô D49
    # = IF(AND(D48=1,AND(TRUNG_GIAN!D$30<=4,OR(D41=1,AND(D36=0,D40=0)),NOT(OR(THONG_CAN!$D$69="Không có gốc",THONG_CAN!$D$69="Gốc dễ bị nhổ",THONG_CAN!$D$69="
    #   Gốc bị hợp (đứng độc lập)")),ABS(1-2)<2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(9,r-8,c),1.0), X.AND(X.eq(E.v(9,r-13,c),0.0), X.eq(E.v(9,r-9,c),0.0))), X.NOT(X.OR(X.eq(E.v(11,69,4),'Không có gốc'), X.eq(E.v(11,69,4),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,4),'Gốc bị hợp (đứng độc lập)'))), X.lt(X.ABS(X.sub(1.0,2.0)),2.0)))) else 0.0)

def f75(E, r, c):
    # Ô F49
    # = IF(AND(F48=1,AND(TRUNG_GIAN!F$30<=4,OR(F41=1,AND(F36=0,F40=0)),NOT(OR(THONG_CAN!$F$69="Không có gốc",THONG_CAN!$F$69="Gốc dễ bị nhổ",THONG_CAN!$F$69="
    #   Gốc bị hợp (đứng độc lập)")),ABS(3-2)<2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(9,r-8,c),1.0), X.AND(X.eq(E.v(9,r-13,c),0.0), X.eq(E.v(9,r-9,c),0.0))), X.NOT(X.OR(X.eq(E.v(11,69,6),'Không có gốc'), X.eq(E.v(11,69,6),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,6),'Gốc bị hợp (đứng độc lập)'))), X.lt(X.ABS(X.sub(3.0,2.0)),2.0)))) else 0.0)

def f76(E, r, c):
    # Ô C50, D50, E50, F50
    # = IF(C27=MOD(TRUNG_GIAN!$E$7+4,5),1,0)
    return (1.0 if X.truthy(X.eq(E.v(9,r-23,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))) else 0.0)

def f77(E, r, c):
    # Ô C51
    # = IF(AND(C50=1,OR(C29=0,C30=1),ABS(0-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.OR(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-21,c),1.0)), X.le(X.ABS(X.sub(0.0,2.0)),1.0))) else 0.0)

def f78(E, r, c):
    # Ô D51
    # = IF(AND(D50=1,OR(D29=0,D30=1),ABS(1-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.OR(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-21,c),1.0)), X.le(X.ABS(X.sub(1.0,2.0)),1.0))) else 0.0)

def f79(E, r, c):
    # Ô E51
    # = IF(AND(E50=1,OR(E29=0,E30=1),ABS(2-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.OR(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-21,c),1.0)), X.le(X.ABS(X.sub(2.0,2.0)),1.0))) else 0.0)

def f80(E, r, c):
    # Ô F51
    # = IF(AND(F50=1,OR(F29=0,F30=1),ABS(3-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.OR(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-21,c),1.0)), X.le(X.ABS(X.sub(3.0,2.0)),1.0))) else 0.0)

def f81(E, r, c):
    # Ô C52, D52, E52, F52
    # = IF(AND(C27<0,C28=0,TRUNG_GIAN!C$21=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)
    return (1.0 if X.truthy(X.AND(X.lt(E.v(9,r-25,c),0.0), X.eq(E.v(9,r-24,c),0.0), X.eq(E.v(23,21,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)

def f82(E, r, c):
    # Ô C53, D53, E53, F53
    # = IF(AND(C27<0,C28=0,TRUNG_GIAN!C$22=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)
    return (1.0 if X.truthy(X.AND(X.lt(E.v(9,r-26,c),0.0), X.eq(E.v(9,r-25,c),0.0), X.eq(E.v(23,22,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)

def f83(E, r, c):
    # Ô C54, D54, E54, F54
    # = IF(AND(C27<0,C28=0,TRUNG_GIAN!C$23=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)
    return (1.0 if X.truthy(X.AND(X.lt(E.v(9,r-27,c),0.0), X.eq(E.v(9,r-26,c),0.0), X.eq(E.v(23,23,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)

def f84(E, r, c):
    # Ô C55
    # = IF(AND(C52=1,INDEX(M_TS,TRUNG_GIAN!C$15+1,TRUNG_GIAN!$D$6+1)<=4,OR(C29=0,C30=1),ABS(0-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,c),1.0), X.add(E.v(23,6,4),1.0)),4.0), X.OR(X.eq(E.v(9,r-26,c),0.0), X.eq(E.v(9,r-25,c),1.0)), X.le(X.ABS(X.sub(0.0,2.0)),1.0))) else 0.0)

def f85(E, r, c):
    # Ô D55
    # = IF(AND(D52=1,INDEX(M_TS,TRUNG_GIAN!D$15+1,TRUNG_GIAN!$D$6+1)<=4,OR(D29=0,D30=1),ABS(1-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,c),1.0), X.add(E.v(23,6,4),1.0)),4.0), X.OR(X.eq(E.v(9,r-26,c),0.0), X.eq(E.v(9,r-25,c),1.0)), X.le(X.ABS(X.sub(1.0,2.0)),1.0))) else 0.0)

def f86(E, r, c):
    # Ô E55
    # = IF(AND(E52=1,INDEX(M_TS,TRUNG_GIAN!E$15+1,TRUNG_GIAN!$D$6+1)<=4,OR(E29=0,E30=1),ABS(2-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,c),1.0), X.add(E.v(23,6,4),1.0)),4.0), X.OR(X.eq(E.v(9,r-26,c),0.0), X.eq(E.v(9,r-25,c),1.0)), X.le(X.ABS(X.sub(2.0,2.0)),1.0))) else 0.0)

def f87(E, r, c):
    # Ô F55
    # = IF(AND(F52=1,INDEX(M_TS,TRUNG_GIAN!F$15+1,TRUNG_GIAN!$D$6+1)<=4,OR(F29=0,F30=1),ABS(3-2)<=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,c),1.0), X.add(E.v(23,6,4),1.0)),4.0), X.OR(X.eq(E.v(9,r-26,c),0.0), X.eq(E.v(9,r-25,c),1.0)), X.le(X.ABS(X.sub(3.0,2.0)),1.0))) else 0.0)

def f88(E, r, c):
    # Ô C56, D56, F56
    # = IF(OR(C37=MOD(TRUNG_GIAN!$E$7+4,5),C27=MOD(TRUNG_GIAN!$E$7+4,5),AND(C27<0,OR(TRUNG_GIAN!C$21=MOD(TRUNG_GIAN!$E$7+4,5),TRUNG_GIAN!C$22=MOD(TRUNG_GIAN!$
    #   E$7+4,5),TRUNG_GIAN!C$23=MOD(TRUNG_GIAN!$E$7+4,5)))),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-19,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(E.v(9,r-29,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.AND(X.lt(E.v(9,r-29,c),0.0), X.OR(X.eq(E.v(23,21,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(E.v(23,22,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(E.v(23,23,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))))) else 0.0)

def f89(E, r, c):
    # Ô E56
    # = IF(OR(FALSE(),E27=MOD(TRUNG_GIAN!$E$7+4,5),AND(E27<0,OR(TRUNG_GIAN!E$21=MOD(TRUNG_GIAN!$E$7+4,5),TRUNG_GIAN!E$22=MOD(TRUNG_GIAN!$E$7+4,5),TRUNG_GIAN!E
    #   $23=MOD(TRUNG_GIAN!$E$7+4,5)))),1,0)
    return (1.0 if X.truthy(X.OR(X.FALSE(), X.eq(E.v(9,r-29,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.AND(X.lt(E.v(9,r-29,c),0.0), X.OR(X.eq(E.v(23,21,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(E.v(23,22,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(E.v(23,23,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))))) else 0.0)

def f90(E, r, c):
    # Ô C57, D57, F57
    # = IF(C37=TRUNG_GIAN!$E$7,1,0)
    return (1.0 if X.truthy(X.eq(E.v(9,r-20,c),E.v(23,7,5))) else 0.0)

def f91(E, r, c):
    # Ô C58
    # = IF(AND(C57=1,AND(TRUNG_GIAN!C$30<=4,OR(C41=1,AND(C36=0,C40=0)),NOT(OR(THONG_CAN!$C$69="Không có gốc",THONG_CAN!$C$69="Gốc dễ bị nhổ",THONG_CAN!$C$69="
    #   Gốc bị hợp (đứng độc lập)")),ABS(0-2)<2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(9,r-17,c),1.0), X.AND(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-18,c),0.0))), X.NOT(X.OR(X.eq(E.v(11,69,3),'Không có gốc'), X.eq(E.v(11,69,3),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,3),'Gốc bị hợp (đứng độc lập)'))), X.lt(X.ABS(X.sub(0.0,2.0)),2.0)))) else 0.0)

def f92(E, r, c):
    # Ô D58
    # = IF(AND(D57=1,AND(TRUNG_GIAN!D$30<=4,OR(D41=1,AND(D36=0,D40=0)),NOT(OR(THONG_CAN!$D$69="Không có gốc",THONG_CAN!$D$69="Gốc dễ bị nhổ",THONG_CAN!$D$69="
    #   Gốc bị hợp (đứng độc lập)")),ABS(1-2)<2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(9,r-17,c),1.0), X.AND(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-18,c),0.0))), X.NOT(X.OR(X.eq(E.v(11,69,4),'Không có gốc'), X.eq(E.v(11,69,4),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,4),'Gốc bị hợp (đứng độc lập)'))), X.lt(X.ABS(X.sub(1.0,2.0)),2.0)))) else 0.0)

def f93(E, r, c):
    # Ô F58
    # = IF(AND(F57=1,AND(TRUNG_GIAN!F$30<=4,OR(F41=1,AND(F36=0,F40=0)),NOT(OR(THONG_CAN!$F$69="Không có gốc",THONG_CAN!$F$69="Gốc dễ bị nhổ",THONG_CAN!$F$69="
    #   Gốc bị hợp (đứng độc lập)")),ABS(3-2)<2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.AND(X.le(E.v(23,30,c),4.0), X.OR(X.eq(E.v(9,r-17,c),1.0), X.AND(X.eq(E.v(9,r-22,c),0.0), X.eq(E.v(9,r-18,c),0.0))), X.NOT(X.OR(X.eq(E.v(11,69,6),'Không có gốc'), X.eq(E.v(11,69,6),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,6),'Gốc bị hợp (đứng độc lập)'))), X.lt(X.ABS(X.sub(3.0,2.0)),2.0)))) else 0.0)

def f94(E, r, c):
    # Ô C62
    # = C43
    return E.v(9,r-19,c)

def f95(E, r, c):
    # Ô D62
    # = C43
    return E.v(9,r-19,c-1)

def f96(E, r, c):
    # Ô C63
    # = IF(SUM(C45:F45)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUM(E.rg(9,r-18,c,r-18,c+3)),0.0)) else 0.0)

def f97(E, r, c):
    # Ô D63
    # = IF(SUM(C46:F46)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUM(E.rg(9,r-17,c-1,r-17,c+2)),0.0)) else 0.0)

def f98(E, r, c):
    # Ô C64
    # = IF(SUM(C48:F48)+SUM(C50:F50)+SUM(C52:F54)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.add(X.SUM(E.rg(9,r-16,c,r-16,c+3)),X.SUM(E.rg(9,r-14,c,r-14,c+3))),X.SUM(E.rg(9,r-12,c,r-10,c+3))),0.0)) else 0.0)

def f99(E, r, c):
    # Ô D64
    # = IF(SUM(C49:F49)+SUM(C51:F51)+SUM(C55:F55)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.add(X.SUM(E.rg(9,r-15,c-1,r-15,c+2)),X.SUM(E.rg(9,r-13,c-1,r-13,c+2))),X.SUM(E.rg(9,r-9,c-1,r-9,c+2))),0.0)) else 0.0)

def f100(E, r, c):
    # Ô C65
    # = IF(SUM(C57:F57)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUM(E.rg(9,r-8,c,r-8,c+3)),0.0)) else 0.0)

def f101(E, r, c):
    # Ô D65
    # = IF(SUM(C58:F58)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUM(E.rg(9,r-7,c-1,r-7,c+2)),0.0)) else 0.0)

def f102(E, r, c):
    # Ô C66, D66
    # = C63+C64+C65
    return X.add(X.add(E.v(9,r-3,c),E.v(9,r-2,c)),E.v(9,r-1,c))

def f103(E, r, c):
    # Ô C67
    # = IF(AND(TRUNG_GIAN!$C$44>0,$C$8=0,$C$6=1,OR(TRUNG_GIAN!$C$46=TRUNG_GIAN!$E$7,TRUNG_GIAN!$C$46=MOD(TRUNG_GIAN!$E$7+4,5))),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(9,8,3),0.0), X.eq(E.v(9,6,3),1.0), X.OR(X.eq(E.v(23,46,3),E.v(23,7,5)), X.eq(E.v(23,46,3),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else 0.0)

def f104(E, r, c):
    # Ô C68
    # = IF(AND(SUM(C56:F56)=0,C65=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.SUM(E.rg(9,r-12,c,r-12,c+3)),0.0), X.eq(E.v(9,r-3,c),0.0))) else 0.0)

def f105(E, r, c):
    # Ô C71
    # = SIGN(IF(C34=1,IF(OR(TRUNG_GIAN!C$76=TRUNG_GIAN!$E$7,TRUNG_GIAN!C$76=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(D34=1,IF(OR(TRUNG_GIAN!D$76=TRUNG_GIAN!$E$7,
    #   TRUNG_GIAN!D$76=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(E34=1,IF(OR(TRUNG_GIAN!E$76=TRUNG_GIAN!$E$7,TRUNG_GIAN!E$76=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+I
    #   F(AND(TRUNG_GIAN!$C$44>0,$C$8=0,$C$6=1),IF(OR(TRUNG_GIAN!$C$46=TRUNG_GIAN!$E$7,TRUNG_GIAN!$C$46=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(AND(C26=1,C21=1)
    #   ,IF(OR(TRUNG_GIAN!C$64=TRUNG_GIAN!$E$7,TRUNG_GIAN!C$64=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(AND(D26=1,D21=1),IF(OR(TRUNG_GIAN!D$64=TRUNG_GIAN!$E$7,TR
    #   UNG_GIAN!D$64=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(AND(E26=1,E21=1),IF(OR(TRUNG_GIAN!E$64=TRUNG_GIAN!$E$7,TRUNG_GIAN!E$64=MOD(TRUNG_GIAN!$E$7+4,5)),1
    #   ,-1),0))
    return X.SIGN(X.add(X.add(X.add(X.add(X.add(X.add(((1.0 if X.truthy(X.OR(X.eq(E.v(23,76,c),E.v(23,7,5)), X.eq(E.v(23,76,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r-37,c),1.0)) else 0.0),((1.0 if X.truthy(X.OR(X.eq(E.v(23,76,c+1),E.v(23,7,5)), X.eq(E.v(23,76,c+1),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r-37,c+1),1.0)) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(23,76,c+2),E.v(23,7,5)), X.eq(E.v(23,76,c+2),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r-37,c+2),1.0)) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(23,46,3),E.v(23,7,5)), X.eq(E.v(23,46,3),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(9,8,3),0.0), X.eq(E.v(9,6,3),1.0))) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(23,64,c),E.v(23,7,5)), X.eq(E.v(23,64,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.AND(X.eq(E.v(9,r-45,c),1.0), X.eq(E.v(9,r-50,c),1.0))) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(23,64,c+1),E.v(23,7,5)), X.eq(E.v(23,64,c+1),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.AND(X.eq(E.v(9,r-45,c+1),1.0), X.eq(E.v(9,r-50,c+1),1.0))) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(23,64,c+2),E.v(23,7,5)), X.eq(E.v(23,64,c+2),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.AND(X.eq(E.v(9,r-45,c+2),1.0), X.eq(E.v(9,r-50,c+2),1.0))) else 0.0)))

def f106(E, r, c):
    # Ô C72
    # = SIGN(IF(C12=1,IF(OR(TRUNG_GIAN!C$55=TRUNG_GIAN!$E$7,TRUNG_GIAN!C$55=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(D12=1,IF(OR(TRUNG_GIAN!D$55=TRUNG_GIAN!$E$7,
    #   TRUNG_GIAN!D$55=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(E12=1,IF(OR(TRUNG_GIAN!E$55=TRUNG_GIAN!$E$7,TRUNG_GIAN!E$55=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+I
    #   F(C123=1,IF(OR(C122=TRUNG_GIAN!$E$7,C122=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)+IF(D123=1,IF(OR(D122=TRUNG_GIAN!$E$7,D122=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),
    #   0)+IF(E123=1,IF(OR(E122=TRUNG_GIAN!$E$7,E122=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0))
    return X.SIGN(X.add(X.add(X.add(X.add(X.add(((1.0 if X.truthy(X.OR(X.eq(E.v(23,55,c),E.v(23,7,5)), X.eq(E.v(23,55,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r-60,c),1.0)) else 0.0),((1.0 if X.truthy(X.OR(X.eq(E.v(23,55,c+1),E.v(23,7,5)), X.eq(E.v(23,55,c+1),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r-60,c+1),1.0)) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(23,55,c+2),E.v(23,7,5)), X.eq(E.v(23,55,c+2),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r-60,c+2),1.0)) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(9,r+50,c),E.v(23,7,5)), X.eq(E.v(9,r+50,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r+51,c),1.0)) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(9,r+50,c+1),E.v(23,7,5)), X.eq(E.v(9,r+50,c+1),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r+51,c+1),1.0)) else 0.0)),((1.0 if X.truthy(X.OR(X.eq(E.v(9,r+50,c+2),E.v(23,7,5)), X.eq(E.v(9,r+50,c+2),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else X.neg(1.0)) if X.truthy(X.eq(E.v(9,r+51,c+2),1.0)) else 0.0)))

def f107(E, r, c):
    # Ô C74, D74, F74
    # = IF(OR(C37=TRUNG_GIAN!$E$7,C37=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-37,c),E.v(23,7,5)), X.eq(E.v(9,r-37,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)

def f108(E, r, c):
    # Ô E74, E75, E76
    # = ""
    return ''

def f109(E, r, c):
    # Ô C75, D75, F75
    # = IF(TRUNG_GIAN!C$30<=4,1,0)
    return (1.0 if X.truthy(X.le(E.v(23,30,c),4.0)) else 0.0)

def f110(E, r, c):
    # Ô C76
    # = IF(OR(C38=1,C40=1,ABS(0-2)>=2),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-38,c),1.0), X.eq(E.v(9,r-36,c),1.0), X.ge(X.ABS(X.sub(0.0,2.0)),2.0))) else 0.0)

def f111(E, r, c):
    # Ô D76
    # = IF(OR(D38=1,D40=1,ABS(1-2)>=2),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-38,c),1.0), X.eq(E.v(9,r-36,c),1.0), X.ge(X.ABS(X.sub(1.0,2.0)),2.0))) else 0.0)

def f112(E, r, c):
    # Ô F76
    # = IF(OR(F38=1,F40=1,ABS(3-2)>=2),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,r-38,c),1.0), X.eq(E.v(9,r-36,c),1.0), X.ge(X.ABS(X.sub(3.0,2.0)),2.0))) else 0.0)

def f113(E, r, c):
    # Ô C77, D77, E77, F77
    # = IF(C27<0,1,0)
    return (1.0 if X.truthy(X.lt(E.v(9,r-50,c),0.0)) else 0.0)

def f114(E, r, c):
    # Ô C78, D78, E78, F78
    # = IF(AND(C77=1,OR(TRUNG_GIAN!C$21=TRUNG_GIAN!$E$7,TRUNG_GIAN!C$21=MOD(TRUNG_GIAN!$E$7+4,5))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-1,c),1.0), X.OR(X.eq(E.v(23,21,c),E.v(23,7,5)), X.eq(E.v(23,21,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else 0.0)

def f115(E, r, c):
    # Ô C79, D79, E79, F79
    # = IF(AND(C77=1,INDEX(M_TS,TRUNG_GIAN!C$15+1,TRUNG_GIAN!$D$6+1)<=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-2,c),1.0), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,c),1.0), X.add(E.v(23,6,4),1.0)),4.0))) else 0.0)

def f116(E, r, c):
    # Ô C80
    # = IF(AND(C77=1,OR(AND(C29=1,C30=0),ABS(0-2)>=2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.OR(X.AND(X.eq(E.v(9,r-51,c),1.0), X.eq(E.v(9,r-50,c),0.0)), X.ge(X.ABS(X.sub(0.0,2.0)),2.0)))) else 0.0)

def f117(E, r, c):
    # Ô D80
    # = IF(AND(D77=1,OR(AND(D29=1,D30=0),ABS(1-2)>=2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.OR(X.AND(X.eq(E.v(9,r-51,c),1.0), X.eq(E.v(9,r-50,c),0.0)), X.ge(X.ABS(X.sub(1.0,2.0)),2.0)))) else 0.0)

def f118(E, r, c):
    # Ô E80
    # = IF(AND(E77=1,OR(AND(E29=1,E30=0),ABS(2-2)>=2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.OR(X.AND(X.eq(E.v(9,r-51,c),1.0), X.eq(E.v(9,r-50,c),0.0)), X.ge(X.ABS(X.sub(2.0,2.0)),2.0)))) else 0.0)

def f119(E, r, c):
    # Ô F80
    # = IF(AND(F77=1,OR(AND(F29=1,F30=0),ABS(3-2)>=2)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(9,r-3,c),1.0), X.OR(X.AND(X.eq(E.v(9,r-51,c),1.0), X.eq(E.v(9,r-50,c),0.0)), X.ge(X.ABS(X.sub(3.0,2.0)),2.0)))) else 0.0)

def f120(E, r, c):
    # Ô C82
    # = C74*C75+D74*D75+F74*F75+SUMPRODUCT($C$78:$F$78,$C$79:$F$79)
    return X.add(X.add(X.add(X.mul(E.v(9,r-8,c),E.v(9,r-7,c)),X.mul(E.v(9,r-8,c+1),E.v(9,r-7,c+1))),X.mul(E.v(9,r-8,c+3),E.v(9,r-7,c+3))),X.SUMPRODUCT(E.rg(9,78,3,78,6), E.rg(9,79,3,79,6)))

def f121(E, r, c):
    # Ô C83
    # = (1-C74)*C75+(1-D74)*D75+(1-F74)*F75+SUMPRODUCT(($C$77:$F$77-$C$78:$F$78),$C$79:$F$79)
    return X.add(X.add(X.add(X.mul(X.sub(1.0,E.v(9,r-9,c)),E.v(9,r-8,c)),X.mul(X.sub(1.0,E.v(9,r-9,c+1)),E.v(9,r-8,c+1))),X.mul(X.sub(1.0,E.v(9,r-9,c+3)),E.v(9,r-8,c+3))),X.SUMPRODUCT(X.A('sub',E.rg(9,77,3,77,6),E.rg(9,78,3,78,6)), E.rg(9,79,3,79,6)))

def f122(E, r, c):
    # Ô C84
    # = SIGN(C82-C83)
    return X.SIGN(X.sub(E.v(9,r-2,c),E.v(9,r-1,c)))

def f123(E, r, c):
    # Ô C85
    # = IF((C74+D74+F74+SUM($C$78:$F$78))=0,1,(C74*C76+D74*D76+F74*F76+SUMPRODUCT($C$78:$F$78,$C$80:$F$80))/(C74+D74+F74+SUM($C$78:$F$78)))
    return (1.0 if X.truthy(X.eq(X.add(X.add(X.add(E.v(9,r-11,c),E.v(9,r-11,c+1)),E.v(9,r-11,c+3)),X.SUM(E.rg(9,78,3,78,6))),0.0)) else X.div(X.add(X.add(X.add(X.mul(E.v(9,r-11,c),E.v(9,r-9,c)),X.mul(E.v(9,r-11,c+1),E.v(9,r-9,c+1))),X.mul(E.v(9,r-11,c+3),E.v(9,r-9,c+3))),X.SUMPRODUCT(E.rg(9,78,3,78,6), E.rg(9,80,3,80,6))),X.add(X.add(X.add(E.v(9,r-11,c),E.v(9,r-11,c+1)),E.v(9,r-11,c+3)),X.SUM(E.rg(9,78,3,78,6)))))

def f124(E, r, c):
    # Ô C86
    # = IF(((1-C74)+(1-D74)+(1-F74)+SUM($C$77:$F$77)-SUM($C$78:$F$78))=0,1,((1-C74)*C76+(1-D74)*D76+(1-F74)*F76+SUMPRODUCT(($C$77:$F$77-$C$78:$F$78),$C$80:$F$
    #   80))/((1-C74)+(1-D74)+(1-F74)+SUM($C$77:$F$77)-SUM($C$78:$F$78)))
    return (1.0 if X.truthy(X.eq(X.sub(X.add(X.add(X.add(X.sub(1.0,E.v(9,r-12,c)),X.sub(1.0,E.v(9,r-12,c+1))),X.sub(1.0,E.v(9,r-12,c+3))),X.SUM(E.rg(9,77,3,77,6))),X.SUM(E.rg(9,78,3,78,6))),0.0)) else X.div(X.add(X.add(X.add(X.mul(X.sub(1.0,E.v(9,r-12,c)),E.v(9,r-10,c)),X.mul(X.sub(1.0,E.v(9,r-12,c+1)),E.v(9,r-10,c+1))),X.mul(X.sub(1.0,E.v(9,r-12,c+3)),E.v(9,r-10,c+3))),X.SUMPRODUCT(X.A('sub',E.rg(9,77,3,77,6),E.rg(9,78,3,78,6)), E.rg(9,80,3,80,6))),X.sub(X.add(X.add(X.add(X.sub(1.0,E.v(9,r-12,c)),X.sub(1.0,E.v(9,r-12,c+1))),X.sub(1.0,E.v(9,r-12,c+3))),X.SUM(E.rg(9,77,3,77,6))),X.SUM(E.rg(9,78,3,78,6)))))

def f125(E, r, c):
    # Ô C87
    # = SIGN(C86-C85)
    return X.SIGN(X.sub(E.v(9,r-1,c),E.v(9,r-2,c)))

def f126(E, r, c):
    # Ô C88
    # = C71+C72+C84+C87
    return X.add(X.add(X.add(E.v(9,r-17,c),E.v(9,r-16,c)),E.v(9,r-4,c)),E.v(9,r-1,c))

def f127(E, r, c):
    # Ô C89
    # = C74+D74+F74+SUM(C78:F78)+IF(AND(C27>=0,OR(C27=TRUNG_GIAN!$E$7,C27=MOD(TRUNG_GIAN!$E$7+4,5))),1,0)+IF(AND(D27>=0,OR(D27=TRUNG_GIAN!$E$7,D27=MOD(TRUNG_G
    #   IAN!$E$7+4,5))),1,0)+IF(AND(E27>=0,OR(E27=TRUNG_GIAN!$E$7,E27=MOD(TRUNG_GIAN!$E$7+4,5))),1,0)+IF(AND(F27>=0,OR(F27=TRUNG_GIAN!$E$7,F27=MOD(TRUNG_GIAN!
    #   $E$7+4,5))),1,0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(E.v(9,r-15,c),E.v(9,r-15,c+1)),E.v(9,r-15,c+3)),X.SUM(E.rg(9,r-11,c,r-11,c+3))),(1.0 if X.truthy(X.AND(X.ge(E.v(9,r-62,c),0.0), X.OR(X.eq(E.v(9,r-62,c),E.v(23,7,5)), X.eq(E.v(9,r-62,c),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(9,r-62,c+1),0.0), X.OR(X.eq(E.v(9,r-62,c+1),E.v(23,7,5)), X.eq(E.v(9,r-62,c+1),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(9,r-62,c+2),0.0), X.OR(X.eq(E.v(9,r-62,c+2),E.v(23,7,5)), X.eq(E.v(9,r-62,c+2),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(9,r-62,c+3),0.0), X.OR(X.eq(E.v(9,r-62,c+3),E.v(23,7,5)), X.eq(E.v(9,r-62,c+3),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else 0.0))

def f128(E, r, c):
    # Ô D89
    # = 7-C89
    return X.sub(7.0,E.v(9,r,c-1))

def f129(E, r, c):
    # Ô C91
    # = IF($C$117=-1,"Thân nhược",IF($C$117=1,"Thân vượng",$C$115))
    return ('Thân nhược' if X.truthy(X.eq(E.v(9,117,3),X.neg(1.0))) else ('Thân vượng' if X.truthy(X.eq(E.v(9,117,3),1.0)) else E.v(9,115,3)))

def f130(E, r, c):
    # Ô C92
    # = $C$116&IF(OR($C$117=-1,$C$117=1)," → bước 5 kiểm chứng: "&IF($C$117=-1,"đổi thành thân nhược","đổi thành thân vượng")&" (P1 tr.70–72)","")
    return X.cat(E.v(9,116,3),(X.cat(X.cat(' → bước 5 kiểm chứng: ',('đổi thành thân nhược' if X.truthy(X.eq(E.v(9,117,3),X.neg(1.0))) else 'đổi thành thân vượng')),' (P1 tr.70–72)') if X.truthy(X.OR(X.eq(E.v(9,117,3),X.neg(1.0)), X.eq(E.v(9,117,3),1.0))) else ''))

def f131(E, r, c):
    # Ô D96
    # = IF(MOD(C37-TRUNG_GIAN!$E$7,5)=0,1,0)+IF(MOD(D37-TRUNG_GIAN!$E$7,5)=0,1,0)+IF(MOD(F37-TRUNG_GIAN!$E$7,5)=0,1,0)
    return X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-59,c-1),E.v(23,7,5)), 5.0),0.0)) else 0.0),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-59,c),E.v(23,7,5)), 5.0),0.0)) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-59,c+2),E.v(23,7,5)), 5.0),0.0)) else 0.0))

def f132(E, r, c):
    # Ô E96
    # = IF(C27>=0,IF(MOD(C27-TRUNG_GIAN!$E$7,5)=0,1,0),IF(AND(TRUNG_GIAN!C$21>=0,MOD(TRUNG_GIAN!C$21-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!C$22>=0,MOD(
    #   TRUNG_GIAN!C$22-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!C$23>=0,MOD(TRUNG_GIAN!C$23-TRUNG_GIAN!$E$7,5)=0),1,0))+IF(D27>=0,IF(MOD(D27-TRUNG_GIAN!$
    #   E$7,5)=0,1,0),IF(AND(TRUNG_GIAN!D$21>=0,MOD(TRUNG_GIAN!D$21-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!D$22>=0,MOD(TRUNG_GIAN!D$22-TRUNG_GIAN!$E$7,5
    #   )=0),1,0)+IF(AND(TRUNG_GIAN!D$23>=0,MOD(TRUNG_GIAN!D$23-TRUNG_GIAN!$E$7,5)=0),1,0))+IF(E27>=0,IF(MOD(E27-TRUNG_GIAN!$E$7,5)=0,1,0),IF(AND(TRUNG_GIAN!E
    #   $21>=0,MOD(TRUNG_GIAN!E$21-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!E$22>=0,MOD(TRUNG_GIAN!E$22-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!E$23>
    #   =0,MOD(TRUNG_GIAN!E$23-TRUNG_GIAN!$E$7,5)=0),1,0))+IF(F27>=0,IF(MOD(F27-TRUNG_GIAN!$E$7,5)=0,1,0),IF(AND(TRUNG_GIAN!F$21>=0,MOD(TRUNG_GIAN!F$21-TRUNG_
    #   GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!F$22>=0,MOD(TRUNG_GIAN!F$22-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!F$23>=0,MOD(TRUNG_GIAN!F$23-TRUNG_GIAN
    #   !$E$7,5)=0),1,0))
    return X.add(X.add(X.add(((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-69,c-2),E.v(23,7,5)), 5.0),0.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-69,c-2),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-2),E.v(23,7,5)), 5.0),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-2),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-2),E.v(23,7,5)), 5.0),0.0))) else 0.0))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-69,c-1),E.v(23,7,5)), 5.0),0.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-69,c-1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-1),E.v(23,7,5)), 5.0),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-1),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-1),E.v(23,7,5)), 5.0),0.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-69,c),E.v(23,7,5)), 5.0),0.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-69,c),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c),E.v(23,7,5)), 5.0),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c),E.v(23,7,5)), 5.0),0.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-69,c+1),E.v(23,7,5)), 5.0),0.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-69,c+1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c+1),E.v(23,7,5)), 5.0),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c+1),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c+1),E.v(23,7,5)), 5.0),0.0))) else 0.0))))

def f133(E, r, c):
    # Ô F96, F97, F98, F99, F100
    # = D96+E96
    return X.add(E.v(9,r,c-2),E.v(9,r,c-1))

def f134(E, r, c):
    # Ô D97
    # = IF(MOD(C37-TRUNG_GIAN!$E$7,5)=1,1,0)+IF(MOD(D37-TRUNG_GIAN!$E$7,5)=1,1,0)+IF(MOD(F37-TRUNG_GIAN!$E$7,5)=1,1,0)
    return X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-60,c-1),E.v(23,7,5)), 5.0),1.0)) else 0.0),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-60,c),E.v(23,7,5)), 5.0),1.0)) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-60,c+2),E.v(23,7,5)), 5.0),1.0)) else 0.0))

def f135(E, r, c):
    # Ô E97
    # = IF(C27>=0,IF(MOD(C27-TRUNG_GIAN!$E$7,5)=1,1,0),IF(AND(TRUNG_GIAN!C$21>=0,MOD(TRUNG_GIAN!C$21-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!C$22>=0,MOD(
    #   TRUNG_GIAN!C$22-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!C$23>=0,MOD(TRUNG_GIAN!C$23-TRUNG_GIAN!$E$7,5)=1),1,0))+IF(D27>=0,IF(MOD(D27-TRUNG_GIAN!$
    #   E$7,5)=1,1,0),IF(AND(TRUNG_GIAN!D$21>=0,MOD(TRUNG_GIAN!D$21-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!D$22>=0,MOD(TRUNG_GIAN!D$22-TRUNG_GIAN!$E$7,5
    #   )=1),1,0)+IF(AND(TRUNG_GIAN!D$23>=0,MOD(TRUNG_GIAN!D$23-TRUNG_GIAN!$E$7,5)=1),1,0))+IF(E27>=0,IF(MOD(E27-TRUNG_GIAN!$E$7,5)=1,1,0),IF(AND(TRUNG_GIAN!E
    #   $21>=0,MOD(TRUNG_GIAN!E$21-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!E$22>=0,MOD(TRUNG_GIAN!E$22-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!E$23>
    #   =0,MOD(TRUNG_GIAN!E$23-TRUNG_GIAN!$E$7,5)=1),1,0))+IF(F27>=0,IF(MOD(F27-TRUNG_GIAN!$E$7,5)=1,1,0),IF(AND(TRUNG_GIAN!F$21>=0,MOD(TRUNG_GIAN!F$21-TRUNG_
    #   GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!F$22>=0,MOD(TRUNG_GIAN!F$22-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!F$23>=0,MOD(TRUNG_GIAN!F$23-TRUNG_GIAN
    #   !$E$7,5)=1),1,0))
    return X.add(X.add(X.add(((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-70,c-2),E.v(23,7,5)), 5.0),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-70,c-2),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-2),E.v(23,7,5)), 5.0),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-2),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-2),E.v(23,7,5)), 5.0),1.0))) else 0.0))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-70,c-1),E.v(23,7,5)), 5.0),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-70,c-1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-1),E.v(23,7,5)), 5.0),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-1),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-1),E.v(23,7,5)), 5.0),1.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-70,c),E.v(23,7,5)), 5.0),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-70,c),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c),E.v(23,7,5)), 5.0),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c),E.v(23,7,5)), 5.0),1.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-70,c+1),E.v(23,7,5)), 5.0),1.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-70,c+1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c+1),E.v(23,7,5)), 5.0),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c+1),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c+1),E.v(23,7,5)), 5.0),1.0))) else 0.0))))

def f136(E, r, c):
    # Ô D98
    # = IF(MOD(C37-TRUNG_GIAN!$E$7,5)=2,1,0)+IF(MOD(D37-TRUNG_GIAN!$E$7,5)=2,1,0)+IF(MOD(F37-TRUNG_GIAN!$E$7,5)=2,1,0)
    return X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-61,c-1),E.v(23,7,5)), 5.0),2.0)) else 0.0),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-61,c),E.v(23,7,5)), 5.0),2.0)) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-61,c+2),E.v(23,7,5)), 5.0),2.0)) else 0.0))

def f137(E, r, c):
    # Ô E98
    # = IF(C27>=0,IF(MOD(C27-TRUNG_GIAN!$E$7,5)=2,1,0),IF(AND(TRUNG_GIAN!C$21>=0,MOD(TRUNG_GIAN!C$21-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!C$22>=0,MOD(
    #   TRUNG_GIAN!C$22-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!C$23>=0,MOD(TRUNG_GIAN!C$23-TRUNG_GIAN!$E$7,5)=2),1,0))+IF(D27>=0,IF(MOD(D27-TRUNG_GIAN!$
    #   E$7,5)=2,1,0),IF(AND(TRUNG_GIAN!D$21>=0,MOD(TRUNG_GIAN!D$21-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!D$22>=0,MOD(TRUNG_GIAN!D$22-TRUNG_GIAN!$E$7,5
    #   )=2),1,0)+IF(AND(TRUNG_GIAN!D$23>=0,MOD(TRUNG_GIAN!D$23-TRUNG_GIAN!$E$7,5)=2),1,0))+IF(E27>=0,IF(MOD(E27-TRUNG_GIAN!$E$7,5)=2,1,0),IF(AND(TRUNG_GIAN!E
    #   $21>=0,MOD(TRUNG_GIAN!E$21-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!E$22>=0,MOD(TRUNG_GIAN!E$22-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!E$23>
    #   =0,MOD(TRUNG_GIAN!E$23-TRUNG_GIAN!$E$7,5)=2),1,0))+IF(F27>=0,IF(MOD(F27-TRUNG_GIAN!$E$7,5)=2,1,0),IF(AND(TRUNG_GIAN!F$21>=0,MOD(TRUNG_GIAN!F$21-TRUNG_
    #   GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!F$22>=0,MOD(TRUNG_GIAN!F$22-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!F$23>=0,MOD(TRUNG_GIAN!F$23-TRUNG_GIAN
    #   !$E$7,5)=2),1,0))
    return X.add(X.add(X.add(((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-71,c-2),E.v(23,7,5)), 5.0),2.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-71,c-2),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-2),E.v(23,7,5)), 5.0),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-2),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-2),E.v(23,7,5)), 5.0),2.0))) else 0.0))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-71,c-1),E.v(23,7,5)), 5.0),2.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-71,c-1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-1),E.v(23,7,5)), 5.0),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-1),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-1),E.v(23,7,5)), 5.0),2.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-71,c),E.v(23,7,5)), 5.0),2.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-71,c),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c),E.v(23,7,5)), 5.0),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c),E.v(23,7,5)), 5.0),2.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-71,c+1),E.v(23,7,5)), 5.0),2.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-71,c+1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c+1),E.v(23,7,5)), 5.0),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c+1),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c+1),E.v(23,7,5)), 5.0),2.0))) else 0.0))))

def f138(E, r, c):
    # Ô D99
    # = IF(MOD(C37-TRUNG_GIAN!$E$7,5)=3,1,0)+IF(MOD(D37-TRUNG_GIAN!$E$7,5)=3,1,0)+IF(MOD(F37-TRUNG_GIAN!$E$7,5)=3,1,0)
    return X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-62,c-1),E.v(23,7,5)), 5.0),3.0)) else 0.0),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-62,c),E.v(23,7,5)), 5.0),3.0)) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-62,c+2),E.v(23,7,5)), 5.0),3.0)) else 0.0))

def f139(E, r, c):
    # Ô E99
    # = IF(C27>=0,IF(MOD(C27-TRUNG_GIAN!$E$7,5)=3,1,0),IF(AND(TRUNG_GIAN!C$21>=0,MOD(TRUNG_GIAN!C$21-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!C$22>=0,MOD(
    #   TRUNG_GIAN!C$22-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!C$23>=0,MOD(TRUNG_GIAN!C$23-TRUNG_GIAN!$E$7,5)=3),1,0))+IF(D27>=0,IF(MOD(D27-TRUNG_GIAN!$
    #   E$7,5)=3,1,0),IF(AND(TRUNG_GIAN!D$21>=0,MOD(TRUNG_GIAN!D$21-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!D$22>=0,MOD(TRUNG_GIAN!D$22-TRUNG_GIAN!$E$7,5
    #   )=3),1,0)+IF(AND(TRUNG_GIAN!D$23>=0,MOD(TRUNG_GIAN!D$23-TRUNG_GIAN!$E$7,5)=3),1,0))+IF(E27>=0,IF(MOD(E27-TRUNG_GIAN!$E$7,5)=3,1,0),IF(AND(TRUNG_GIAN!E
    #   $21>=0,MOD(TRUNG_GIAN!E$21-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!E$22>=0,MOD(TRUNG_GIAN!E$22-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!E$23>
    #   =0,MOD(TRUNG_GIAN!E$23-TRUNG_GIAN!$E$7,5)=3),1,0))+IF(F27>=0,IF(MOD(F27-TRUNG_GIAN!$E$7,5)=3,1,0),IF(AND(TRUNG_GIAN!F$21>=0,MOD(TRUNG_GIAN!F$21-TRUNG_
    #   GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!F$22>=0,MOD(TRUNG_GIAN!F$22-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!F$23>=0,MOD(TRUNG_GIAN!F$23-TRUNG_GIAN
    #   !$E$7,5)=3),1,0))
    return X.add(X.add(X.add(((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-72,c-2),E.v(23,7,5)), 5.0),3.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-72,c-2),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-2),E.v(23,7,5)), 5.0),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-2),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-2),E.v(23,7,5)), 5.0),3.0))) else 0.0))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-72,c-1),E.v(23,7,5)), 5.0),3.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-72,c-1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-1),E.v(23,7,5)), 5.0),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-1),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-1),E.v(23,7,5)), 5.0),3.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-72,c),E.v(23,7,5)), 5.0),3.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-72,c),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c),E.v(23,7,5)), 5.0),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c),E.v(23,7,5)), 5.0),3.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-72,c+1),E.v(23,7,5)), 5.0),3.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-72,c+1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c+1),E.v(23,7,5)), 5.0),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c+1),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c+1),E.v(23,7,5)), 5.0),3.0))) else 0.0))))

def f140(E, r, c):
    # Ô D100
    # = IF(MOD(C37-TRUNG_GIAN!$E$7,5)=4,1,0)+IF(MOD(D37-TRUNG_GIAN!$E$7,5)=4,1,0)+IF(MOD(F37-TRUNG_GIAN!$E$7,5)=4,1,0)
    return X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-63,c-1),E.v(23,7,5)), 5.0),4.0)) else 0.0),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-63,c),E.v(23,7,5)), 5.0),4.0)) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-63,c+2),E.v(23,7,5)), 5.0),4.0)) else 0.0))

def f141(E, r, c):
    # Ô E100
    # = IF(C27>=0,IF(MOD(C27-TRUNG_GIAN!$E$7,5)=4,1,0),IF(AND(TRUNG_GIAN!C$21>=0,MOD(TRUNG_GIAN!C$21-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!C$22>=0,MOD(
    #   TRUNG_GIAN!C$22-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!C$23>=0,MOD(TRUNG_GIAN!C$23-TRUNG_GIAN!$E$7,5)=4),1,0))+IF(D27>=0,IF(MOD(D27-TRUNG_GIAN!$
    #   E$7,5)=4,1,0),IF(AND(TRUNG_GIAN!D$21>=0,MOD(TRUNG_GIAN!D$21-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!D$22>=0,MOD(TRUNG_GIAN!D$22-TRUNG_GIAN!$E$7,5
    #   )=4),1,0)+IF(AND(TRUNG_GIAN!D$23>=0,MOD(TRUNG_GIAN!D$23-TRUNG_GIAN!$E$7,5)=4),1,0))+IF(E27>=0,IF(MOD(E27-TRUNG_GIAN!$E$7,5)=4,1,0),IF(AND(TRUNG_GIAN!E
    #   $21>=0,MOD(TRUNG_GIAN!E$21-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!E$22>=0,MOD(TRUNG_GIAN!E$22-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!E$23>
    #   =0,MOD(TRUNG_GIAN!E$23-TRUNG_GIAN!$E$7,5)=4),1,0))+IF(F27>=0,IF(MOD(F27-TRUNG_GIAN!$E$7,5)=4,1,0),IF(AND(TRUNG_GIAN!F$21>=0,MOD(TRUNG_GIAN!F$21-TRUNG_
    #   GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!F$22>=0,MOD(TRUNG_GIAN!F$22-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!F$23>=0,MOD(TRUNG_GIAN!F$23-TRUNG_GIAN
    #   !$E$7,5)=4),1,0))
    return X.add(X.add(X.add(((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-73,c-2),E.v(23,7,5)), 5.0),4.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-73,c-2),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-2),E.v(23,7,5)), 5.0),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-2),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-2),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-2),E.v(23,7,5)), 5.0),4.0))) else 0.0))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-73,c-1),E.v(23,7,5)), 5.0),4.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-73,c-1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c-1),E.v(23,7,5)), 5.0),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c-1),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c-1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c-1),E.v(23,7,5)), 5.0),4.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-73,c),E.v(23,7,5)), 5.0),4.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-73,c),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c),E.v(23,7,5)), 5.0),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c),E.v(23,7,5)), 5.0),4.0))) else 0.0)))),((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(9,r-73,c+1),E.v(23,7,5)), 5.0),4.0)) else 0.0) if X.truthy(X.ge(E.v(9,r-73,c+1),0.0)) else X.add(X.add((1.0 if X.truthy(X.AND(X.ge(E.v(23,21,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,21,c+1),E.v(23,7,5)), 5.0),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,22,c+1),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,c+1),0.0), X.eq(X.MOD(X.sub(E.v(23,23,c+1),E.v(23,7,5)), 5.0),4.0))) else 0.0))))

def f142(E, r, c):
    # Ô C104
    # = "Chi tháng "&TRUNG_GIAN!$D$11&"; nhật can "&TRUNG_GIAN!$E$10&" ở giai đoạn "&TRUNG_GIAN!$D$29&" tại chi tháng. "&IF(C43=1,"Giai đoạn này thuộc 5 giai 
    #   đoạn Trường sinh – Mộc dục – Quan đái – Lâm quan – Đế vượng → ĐƯỢC LỆNH: khí mùa sinh ủng hộ nhật can (P1 tr.70).","Giai đoạn này KHÔNG thuộc 5 giai đ
    #   oạn Trường sinh – Mộc dục – Quan đái – Lâm quan – Đế vượng → KHÔNG ĐƯỢC LỆNH: mùa sinh không phải mùa của nhật can; đây là tiêu chí nặng nhất (P1 tr.7
    #   0).")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chi tháng ',E.v(23,11,4)),'; nhật can '),E.v(23,10,5)),' ở giai đoạn '),E.v(23,29,4)),' tại chi tháng. '),('Giai đoạn này thuộc 5 giai đoạn Trường sinh – Mộc dục – Quan đái – Lâm quan – Đế vượng → ĐƯỢC LỆNH: khí mùa sinh ủng hộ nhật can (P1 tr.70).' if X.truthy(X.eq(E.v(9,r-61,c),1.0)) else 'Giai đoạn này KHÔNG thuộc 5 giai đoạn Trường sinh – Mộc dục – Quan đái – Lâm quan – Đế vượng → KHÔNG ĐƯỢC LỆNH: mùa sinh không phải mùa của nhật can; đây là tiêu chí nặng nhất (P1 tr.70).'))

def f143(E, r, c):
    # Ô C105
    # = IF(TRIM(IF(C45=1,"Chi Năm "&TRUNG_GIAN!$C$11&": "&IF(C27>=0,"chi đã hợp hóa thành hành nhật chủ (gốc)","nhật can ở "&TRUNG_GIAN!$C$29&" tại chi này → 
    #   là gốc")&IF(C46=1," – MẠNH (không bị xung, gần nhật can)"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(C29=1,C30=0),"chi bị xung → gốc dễ bị nhổ (P1 tr.
    #   71; P2 tr.424), ","")&IF(ABS(0-2)>1,"trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ",IF(AND(C28=1,OR(AND(OR(TRUNG_GIAN!$C$28=
    #   3,TRUNG_GIAN!$C$28=4),TRUNG_GIAN!$C$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(TRUNG_GIAN!$C$28=0,TRUNG_GIAN!$C$28=8)))),"Chi Năm "&TRUNG_GIAN!$C$11
    #   &": là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ",""))&IF(E45=1,"Chi Ngày "&TRUNG_GIAN!$E$11&": "&IF(E27>=0,"chi đã hợp hóa thàn
    #   h hành nhật chủ (gốc)","nhật can ở "&TRUNG_GIAN!$E$29&" tại chi này → là gốc")&IF(E46=1," – MẠNH (không bị xung, gần nhật can)"," – KHÔNG MẠNH vì "&SU
    #   BSTITUTE(TRIM(IF(AND(E29=1,E30=0),"chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ","")&IF(ABS(2-2)>1,"trụ Ngày cách xa trụ ngày → lực yếu (P1 tr.
    #   71), ",""))&"#",",#",""))&"; ",IF(AND(E28=1,OR(AND(OR(TRUNG_GIAN!$E$28=3,TRUNG_GIAN!$E$28=4),TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,O
    #   R(TRUNG_GIAN!$E$28=0,TRUNG_GIAN!$E$28=8)))),"Chi Ngày "&TRUNG_GIAN!$E$11&": là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ",""))&I
    #   F(F45=1,"Chi Giờ "&TRUNG_GIAN!$F$11&": "&IF(F27>=0,"chi đã hợp hóa thành hành nhật chủ (gốc)","nhật can ở "&TRUNG_GIAN!$F$29&" tại chi này → là gốc")&
    #   IF(F46=1," – MẠNH (không bị xung, gần nhật can)"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(F29=1,F30=0),"chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr
    #   .424), ","")&IF(ABS(3-2)>1,"trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ",IF(AND(F28=1,OR(AND(OR(TRUNG_GIAN!$F$28=3,TRUNG_G
    #   IAN!$F$28=4),TRUNG_GIAN!$F$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(TRUNG_GIAN!$F$28=0,TRUNG_GIAN!$F$28=8)))),"Chi Giờ "&TRUNG_GIAN!$F$11&": là gố
    #   c nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ","")))="","Không có chi nào (ngoài chi tháng) là gốc của nhật can – gốc là Lộc, Kình dươ
    #   ng (chỉ khi bản khí của chi là tỉ/kiếp – P1 tr.70, 72); với can dương thêm Trường sinh và Mộ khố (P1 tr.70); "&IF(P_DUKHI=1,"với can âm thêm dư khí Ất
    #   –Thìn, Đinh–Mùi, Tân–Tuất, Quý–Sửu (mở rộng Tử Bình Chân Thuyên ch.6 – đang BẬT).","phần mở rộng dư khí của can âm (Tử Bình Chân Thuyên) đang TẮT – CS
    #   DL P_DUKHI."),TRIM(IF(C45=1,"Chi Năm "&TRUNG_GIAN!$C$11&": "&IF(C27>=0,"chi đã hợp hóa thành hành nhật chủ (gốc)","nhật can ở "&TRUNG_GIAN!$C$29&" tại
    #    chi này → là gốc")&IF(C46=1," – MẠNH (không bị xung, gần nhật can)"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(C29=1,C30=0),"chi bị xung → gốc dễ bị 
    #   nhổ (P1 tr.71; P2 tr.424), ","")&IF(ABS(0-2)>1,"trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ",IF(AND(C28=1,OR(AND(OR(TRUNG_
    #   GIAN!$C$28=3,TRUNG_GIAN!$C$28=4),TRUNG_GIAN!$C$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(TRUNG_GIAN!$C$28=0,TRUNG_GIAN!$C$28=8)))),"Chi Năm "&TRUNG
    #   _GIAN!$C$11&": là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ",""))&IF(E45=1,"Chi Ngày "&TRUNG_GIAN!$E$11&": "&IF(E27>=0,"chi đã h
    #   ợp hóa thành hành nhật chủ (gốc)","nhật can ở "&TRUNG_GIAN!$E$29&" tại chi này → là gốc")&IF(E46=1," – MẠNH (không bị xung, gần nhật can)"," – KHÔNG M
    #   ẠNH vì "&SUBSTITUTE(TRIM(IF(AND(E29=1,E30=0),"chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ","")&IF(ABS(2-2)>1,"trụ Ngày cách xa trụ ngày → lực 
    #   yếu (P1 tr.71), ",""))&"#",",#",""))&"; ",IF(AND(E28=1,OR(AND(OR(TRUNG_GIAN!$E$28=3,TRUNG_GIAN!$E$28=4),TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7),AND(TRUNG_GI
    #   AN!$E$8=0,OR(TRUNG_GIAN!$E$28=0,TRUNG_GIAN!$E$28=8)))),"Chi Ngày "&TRUNG_GIAN!$E$11&": là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62
    #   ); ",""))&IF(F45=1,"Chi Giờ "&TRUNG_GIAN!$F$11&": "&IF(F27>=0,"chi đã hợp hóa thành hành nhật chủ (gốc)","nhật can ở "&TRUNG_GIAN!$F$29&" tại chi này 
    #   → là gốc")&IF(F46=1," – MẠNH (không bị xung, gần nhật can)"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(F29=1,F30=0),"chi bị xung → gốc dễ bị nhổ (P1 t
    #   r.71; P2 tr.424), ","")&IF(ABS(3-2)>1,"trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ",IF(AND(F28=1,OR(AND(OR(TRUNG_GIAN!$F$2
    #   8=3,TRUNG_GIAN!$F$28=4),TRUNG_GIAN!$F$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(TRUNG_GIAN!$F$28=0,TRUNG_GIAN!$F$28=8)))),"Chi Giờ "&TRUNG_GIAN!$F$
    #   11&": là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ",""))))&" (Chi tháng không xét ở đây vì đã xét ở ①.) → Kết luận ②: "&IF(C63=0
    #   ,"KHÔNG đắc địa.",IF(D63=1,"CÓ và MẠNH.","CÓ nhưng KHÔNG MẠNH."))
    return X.cat(X.cat((X.cat('Không có chi nào (ngoài chi tháng) là gốc của nhật can – gốc là Lộc, Kình dương (chỉ khi bản khí của chi là tỉ/kiếp – P1 tr.70, 72); với can dương thêm Trường sinh và Mộ khố (P1 tr.70); ',('với can âm thêm dư khí Ất–Thìn, Đinh–Mùi, Tân–Tuất, Quý–Sửu (mở rộng Tử Bình Chân Thuyên ch.6 – đang BẬT).' if X.truthy(X.eq(E.v(24,76,6),1.0)) else 'phần mở rộng dư khí của can âm (Tử Bình Chân Thuyên) đang TẮT – CSDL P_DUKHI.')) if X.truthy(X.eq(X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('Chi Năm ',E.v(23,11,3)),': '),('chi đã hợp hóa thành hành nhật chủ (gốc)' if X.truthy(X.ge(E.v(9,r-78,c),0.0)) else X.cat(X.cat('nhật can ở ',E.v(23,29,3)),' tại chi này → là gốc'))),(' – MẠNH (không bị xung, gần nhật can)' if X.truthy(X.eq(E.v(9,r-59,c),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ' if X.truthy(X.AND(X.eq(E.v(9,r-76,c),1.0), X.eq(E.v(9,r-75,c),0.0))) else ''),('trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(0.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-60,c),1.0)) else (X.cat(X.cat('Chi Năm ',E.v(23,11,3)),': là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ') if X.truthy(X.AND(X.eq(E.v(9,r-77,c),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,28,3),3.0), X.eq(E.v(23,28,3),4.0)), X.eq(E.v(23,21,3),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,3),0.0), X.eq(E.v(23,28,3),8.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('Chi Ngày ',E.v(23,11,5)),': '),('chi đã hợp hóa thành hành nhật chủ (gốc)' if X.truthy(X.ge(E.v(9,r-78,c+2),0.0)) else X.cat(X.cat('nhật can ở ',E.v(23,29,5)),' tại chi này → là gốc'))),(' – MẠNH (không bị xung, gần nhật can)' if X.truthy(X.eq(E.v(9,r-59,c+2),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ' if X.truthy(X.AND(X.eq(E.v(9,r-76,c+2),1.0), X.eq(E.v(9,r-75,c+2),0.0))) else ''),('trụ Ngày cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(2.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-60,c+2),1.0)) else (X.cat(X.cat('Chi Ngày ',E.v(23,11,5)),': là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ') if X.truthy(X.AND(X.eq(E.v(9,r-77,c+2),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,28,5),3.0), X.eq(E.v(23,28,5),4.0)), X.eq(E.v(23,21,5),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,5),0.0), X.eq(E.v(23,28,5),8.0)))))) else ''))),(X.cat(X.cat(X.cat(X.cat(X.cat('Chi Giờ ',E.v(23,11,6)),': '),('chi đã hợp hóa thành hành nhật chủ (gốc)' if X.truthy(X.ge(E.v(9,r-78,c+3),0.0)) else X.cat(X.cat('nhật can ở ',E.v(23,29,6)),' tại chi này → là gốc'))),(' – MẠNH (không bị xung, gần nhật can)' if X.truthy(X.eq(E.v(9,r-59,c+3),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ' if X.truthy(X.AND(X.eq(E.v(9,r-76,c+3),1.0), X.eq(E.v(9,r-75,c+3),0.0))) else ''),('trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(3.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-60,c+3),1.0)) else (X.cat(X.cat('Chi Giờ ',E.v(23,11,6)),': là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ') if X.truthy(X.AND(X.eq(E.v(9,r-77,c+3),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,28,6),3.0), X.eq(E.v(23,28,6),4.0)), X.eq(E.v(23,21,6),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,6),0.0), X.eq(E.v(23,28,6),8.0)))))) else '')))),'')) else X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('Chi Năm ',E.v(23,11,3)),': '),('chi đã hợp hóa thành hành nhật chủ (gốc)' if X.truthy(X.ge(E.v(9,r-78,c),0.0)) else X.cat(X.cat('nhật can ở ',E.v(23,29,3)),' tại chi này → là gốc'))),(' – MẠNH (không bị xung, gần nhật can)' if X.truthy(X.eq(E.v(9,r-59,c),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ' if X.truthy(X.AND(X.eq(E.v(9,r-76,c),1.0), X.eq(E.v(9,r-75,c),0.0))) else ''),('trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(0.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-60,c),1.0)) else (X.cat(X.cat('Chi Năm ',E.v(23,11,3)),': là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ') if X.truthy(X.AND(X.eq(E.v(9,r-77,c),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,28,3),3.0), X.eq(E.v(23,28,3),4.0)), X.eq(E.v(23,21,3),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,3),0.0), X.eq(E.v(23,28,3),8.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('Chi Ngày ',E.v(23,11,5)),': '),('chi đã hợp hóa thành hành nhật chủ (gốc)' if X.truthy(X.ge(E.v(9,r-78,c+2),0.0)) else X.cat(X.cat('nhật can ở ',E.v(23,29,5)),' tại chi này → là gốc'))),(' – MẠNH (không bị xung, gần nhật can)' if X.truthy(X.eq(E.v(9,r-59,c+2),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ' if X.truthy(X.AND(X.eq(E.v(9,r-76,c+2),1.0), X.eq(E.v(9,r-75,c+2),0.0))) else ''),('trụ Ngày cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(2.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-60,c+2),1.0)) else (X.cat(X.cat('Chi Ngày ',E.v(23,11,5)),': là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ') if X.truthy(X.AND(X.eq(E.v(9,r-77,c+2),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,28,5),3.0), X.eq(E.v(23,28,5),4.0)), X.eq(E.v(23,21,5),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,5),0.0), X.eq(E.v(23,28,5),8.0)))))) else ''))),(X.cat(X.cat(X.cat(X.cat(X.cat('Chi Giờ ',E.v(23,11,6)),': '),('chi đã hợp hóa thành hành nhật chủ (gốc)' if X.truthy(X.ge(E.v(9,r-78,c+3),0.0)) else X.cat(X.cat('nhật can ở ',E.v(23,29,6)),' tại chi này → là gốc'))),(' – MẠNH (không bị xung, gần nhật can)' if X.truthy(X.eq(E.v(9,r-59,c+3),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung → gốc dễ bị nhổ (P1 tr.71; P2 tr.424), ' if X.truthy(X.AND(X.eq(E.v(9,r-76,c+3),1.0), X.eq(E.v(9,r-75,c+3),0.0))) else ''),('trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(3.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-60,c+3),1.0)) else (X.cat(X.cat('Chi Giờ ',E.v(23,11,6)),': là gốc nhưng hợp không hóa – đứng độc lập, không tính (P1 tr.62); ') if X.truthy(X.AND(X.eq(E.v(9,r-77,c+3),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,28,6),3.0), X.eq(E.v(23,28,6),4.0)), X.eq(E.v(23,21,6),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,6),0.0), X.eq(E.v(23,28,6),8.0)))))) else ''))))),' (Chi tháng không xét ở đây vì đã xét ở ①.) → Kết luận ②: '),('KHÔNG đắc địa.' if X.truthy(X.eq(E.v(9,r-42,c),0.0)) else ('CÓ và MẠNH.' if X.truthy(X.eq(E.v(9,r-42,c+1),1.0)) else 'CÓ nhưng KHÔNG MẠNH.')))

def f144(E, r, c):
    # Ô C106
    # = IF(TRIM(J106&" "&J107&" "&J108&" "&J109)="","Không có ấn (hành sinh nhật can) lộ ở can năm, tháng, giờ, cũng không tàng trong chi năm, tháng, ngày, gi
    #   ờ (ấn ở các can chi khác trong tứ trụ – P1 tr.70-72).",TRIM(J106&" "&J107&" "&J108&" "&J109))&" → Kết luận ③: "&IF(C64=0,"KHÔNG được sinh.",IF(D64=1,"
    #   CÓ và MẠNH.","CÓ nhưng KHÔNG MẠNH."))
    return X.cat(X.cat(('Không có ấn (hành sinh nhật can) lộ ở can năm, tháng, giờ, cũng không tàng trong chi năm, tháng, ngày, giờ (ấn ở các can chi khác trong tứ trụ – P1 tr.70-72).' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(9,r,c+7),' '),E.v(9,r+1,c+7)),' '),E.v(9,r+2,c+7)),' '),E.v(9,r+3,c+7))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(9,r,c+7),' '),E.v(9,r+1,c+7)),' '),E.v(9,r+2,c+7)),' '),E.v(9,r+3,c+7)))),' → Kết luận ③: '),('KHÔNG được sinh.' if X.truthy(X.eq(E.v(9,r-42,c),0.0)) else ('CÓ và MẠNH.' if X.truthy(X.eq(E.v(9,r-42,c+1),1.0)) else 'CÓ nhưng KHÔNG MẠNH.')))

def f145(E, r, c):
    # Ô J106
    # = TRIM(IF(C48=1,"Can Năm "&TRUNG_GIAN!$C$10&" ("&TRUNG_GIAN!$C$12&") là ấn lộ"&IF(C49=1," – MẠNH"&IF(AND(C41=1,OR(C38=1,C40=1))," (tuy hợp/bị khắc gần n
    #   hưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$C$30>4,"khôn
    #   g vượng tại lệnh (ở "&TRUNG_GIAN!$C$31&" tại tháng), ","")&IF(C41=1,"",IF(C36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(C40=1,"b
    #   ị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$C$69="Không có gốc",THONG_CAN!$C$69="Gốc dễ bị nhổ",THONG_CAN!$C$69="Gốc bị hợp (đứng độc lập)"),LOWER(THON
    #   G_CAN!$C$69)&", ","")&IF(ABS(0-2)>=2,"cách xa nhật can, ",""))&"#",",#",""))&"; ","")&IF(D48=1,"Can Tháng "&TRUNG_GIAN!$D$10&" ("&TRUNG_GIAN!$D$12&") 
    #   là ấn lộ"&IF(D49=1," – MẠNH"&IF(AND(D41=1,OR(D38=1,D40=1))," (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.
    #   32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$D$30>4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$D$31&" tại tháng), ","")&IF(D41=1
    #   ,"",IF(D36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(D40=1,"bị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$D$69="Không có gốc",THON
    #   G_CAN!$D$69="Gốc dễ bị nhổ",THONG_CAN!$D$69="Gốc bị hợp (đứng độc lập)"),LOWER(THONG_CAN!$D$69)&", ","")&IF(ABS(1-2)>=2,"cách xa nhật can, ",""))&"#",
    #   ",#",""))&"; ","")&IF(F48=1,"Can Giờ "&TRUNG_GIAN!$F$10&" ("&TRUNG_GIAN!$F$12&") là ấn lộ"&IF(F49=1," – MẠNH"&IF(AND(F41=1,OR(F38=1,F40=1))," (tuy hợp
    #   /bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN
    #   !$F$30>4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$F$31&" tại tháng), ","")&IF(F41=1,"",IF(F36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","
    #   ")&IF(F40=1,"bị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$F$69="Không có gốc",THONG_CAN!$F$69="Gốc dễ bị nhổ",THONG_CAN!$F$69="Gốc bị hợp (đứng độc lập
    #   )"),LOWER(THONG_CAN!$F$69)&", ","")&IF(ABS(3-2)>=2,"cách xa nhật can, ",""))&"#",",#",""))&"; ","")&IF(C50=1,"Chi Năm "&TRUNG_GIAN!$C$11&" hóa thành ấ
    #   n"&IF(C51=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(C29=1,C30=0),"chi bị xung, ","")&IF(ABS(0-2)>1,"trụ Năm cách xa trụ ngày → lực yếu (P
    #   1 tr.71), ",""))&"#",",#",""))&"; ","")&IF(C52=1,"Bản khí "&INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" trong chi Năm "&TRUNG_GIAN!$C$11&" là ấn"&IF(C55=1," –
    #    MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(INDEX(M_TS,TRUNG_GIAN!$C$15+1,TRUNG_GIAN!$D$6+1)>4,"không vượng tại lệnh, ","")&IF(AND(C29=1,C30=0),"chi
    #    bị xung, ","")&IF(ABS(0-2)>1,"trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ","")&IF(C53=1,"Trung khí "&INDEX(CAN_TEN,TRUNG_
    #   GIAN!$C$17+1)&" trong chi Năm "&TRUNG_GIAN!$C$11&" là ấn (trung khí – không tính mạnh); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Năm ',E.v(23,10,3)),' ('),E.v(23,12,3)),') là ấn lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-65,c-7),1.0), X.OR(X.eq(E.v(9,r-68,c-7),1.0), X.eq(E.v(9,r-66,c-7),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-57,c-7),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,3)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,3),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-65,c-7),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-70,c-7),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-66,c-7),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,3)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,3),'Không có gốc'), X.eq(E.v(11,69,3),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,3),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(0.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-58,c-7),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Tháng ',E.v(23,10,4)),' ('),E.v(23,12,4)),') là ấn lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-65,c-6),1.0), X.OR(X.eq(E.v(9,r-68,c-6),1.0), X.eq(E.v(9,r-66,c-6),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-57,c-6),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,4)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,4),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-65,c-6),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-70,c-6),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-66,c-6),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,4)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,4),'Không có gốc'), X.eq(E.v(11,69,4),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,4),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(1.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-58,c-6),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Giờ ',E.v(23,10,6)),' ('),E.v(23,12,6)),') là ấn lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-65,c-4),1.0), X.OR(X.eq(E.v(9,r-68,c-4),1.0), X.eq(E.v(9,r-66,c-4),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-57,c-4),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,6)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,6),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-65,c-4),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-70,c-4),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-66,c-4),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,6)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,6),'Không có gốc'), X.eq(E.v(11,69,6),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,6),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(3.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-58,c-4),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Chi Năm ',E.v(23,11,3)),' hóa thành ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-55,c-7),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-77,c-7),1.0), X.eq(E.v(9,r-76,c-7),0.0))) else ''),('trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(0.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-56,c-7),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bản khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0))),' trong chi Năm '),E.v(23,11,3)),' là ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-51,c-7),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(('không vượng tại lệnh, ' if X.truthy(X.gt(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else ''),('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-77,c-7),1.0), X.eq(E.v(9,r-76,c-7),0.0))) else '')),('trụ Năm cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(0.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-54,c-7),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Trung khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0))),' trong chi Năm '),E.v(23,11,3)),' là ấn (trung khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-53,c-7),1.0)) else '')))

def f146(E, r, c):
    # Ô C107
    # = IF(TRIM(IF(C57=1,"Can Năm "&TRUNG_GIAN!$C$10&" ("&TRUNG_GIAN!$C$12&") là tỉ kiếp lộ"&IF(C58=1," – MẠNH"&IF(AND(C41=1,OR(C38=1,C40=1))," (tuy hợp/bị kh
    #   ắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$C$30
    #   >4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$C$31&" tại tháng), ","")&IF(C41=1,"",IF(C36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(
    #   C40=1,"bị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$C$69="Không có gốc",THONG_CAN!$C$69="Gốc dễ bị nhổ",THONG_CAN!$C$69="Gốc bị hợp (đứng độc lập)"),LO
    #   WER(THONG_CAN!$C$69)&", ","")&IF(ABS(0-2)>=2,"cách xa nhật can, ",""))&"#",",#",""))&"; ","")&IF(D57=1,"Can Tháng "&TRUNG_GIAN!$D$10&" ("&TRUNG_GIAN!$
    #   D$12&") là tỉ kiếp lộ"&IF(D58=1," – MẠNH"&IF(AND(D41=1,OR(D38=1,D40=1))," (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền
    #    chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$D$30>4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$D$31&" tại tháng), "
    #   ,"")&IF(D41=1,"",IF(D36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(D40=1,"bị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$D$69="Không
    #    có gốc",THONG_CAN!$D$69="Gốc dễ bị nhổ",THONG_CAN!$D$69="Gốc bị hợp (đứng độc lập)"),LOWER(THONG_CAN!$D$69)&", ","")&IF(ABS(1-2)>=2,"cách xa nhật can
    #   , ",""))&"#",",#",""))&"; ","")&IF(F57=1,"Can Giờ "&TRUNG_GIAN!$F$10&" ("&TRUNG_GIAN!$F$12&") là tỉ kiếp lộ"&IF(F58=1," – MẠNH"&IF(AND(F41=1,OR(F38=1,
    #   F40=1))," (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(
    #   TRIM(IF(TRUNG_GIAN!$F$30>4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$F$31&" tại tháng), ","")&IF(F41=1,"",IF(F36=1,"hợp không hóa, bị khắc gần còn 2/3 lự
    #   c (P2 tr.339), ","")&IF(F40=1,"bị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$F$69="Không có gốc",THONG_CAN!$F$69="Gốc dễ bị nhổ",THONG_CAN!$F$69="Gốc bị
    #    hợp (đứng độc lập)"),LOWER(THONG_CAN!$F$69)&", ","")&IF(ABS(3-2)>=2,"cách xa nhật can, ",""))&"#",",#",""))&"; ",""))="","Không có tỉ kiếp lộ ở thiên
    #    can (sách chỉ tính các thiên can cùng hành nhật can – P1 tr.70; tỉ kiếp tàng trong chi đã tính ở ② đắc địa).",TRIM(IF(C57=1,"Can Năm "&TRUNG_GIAN!$C$
    #   10&" ("&TRUNG_GIAN!$C$12&") là tỉ kiếp lộ"&IF(C58=1," – MẠNH"&IF(AND(C41=1,OR(C38=1,C40=1))," (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở ch
    #   i cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$C$30>4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$
    #   C$31&" tại tháng), ","")&IF(C41=1,"",IF(C36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(C40=1,"bị can kề khắc chế, ",""))&IF(OR(TH
    #   ONG_CAN!$C$69="Không có gốc",THONG_CAN!$C$69="Gốc dễ bị nhổ",THONG_CAN!$C$69="Gốc bị hợp (đứng độc lập)"),LOWER(THONG_CAN!$C$69)&", ","")&IF(ABS(0-2)>
    #   =2,"cách xa nhật can, ",""))&"#",",#",""))&"; ","")&IF(D57=1,"Can Tháng "&TRUNG_GIAN!$D$10&" ("&TRUNG_GIAN!$D$12&") là tỉ kiếp lộ"&IF(D58=1," – MẠNH"&
    #   IF(AND(D41=1,OR(D38=1,D40=1))," (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG
    #    MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$D$30>4,"không vượng tại lệnh (ở "&TRUNG_GIAN!$D$31&" tại tháng), ","")&IF(D41=1,"",IF(D36=1,"hợp không hóa, 
    #   bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(D40=1,"bị can kề khắc chế, ",""))&IF(OR(THONG_CAN!$D$69="Không có gốc",THONG_CAN!$D$69="Gốc dễ bị nhổ",T
    #   HONG_CAN!$D$69="Gốc bị hợp (đứng độc lập)"),LOWER(THONG_CAN!$D$69)&", ","")&IF(ABS(1-2)>=2,"cách xa nhật can, ",""))&"#",",#",""))&"; ","")&IF(F57=1,"
    #   Can Giờ "&TRUNG_GIAN!$F$10&" ("&TRUNG_GIAN!$F$12&") là tỉ kiếp lộ"&IF(F58=1," – MẠNH"&IF(AND(F41=1,OR(F38=1,F40=1))," (tuy hợp/bị khắc gần nhưng vượng
    #    tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)","")," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(TRUNG_GIAN!$F$30>4,"không vượng tạ
    #   i lệnh (ở "&TRUNG_GIAN!$F$31&" tại tháng), ","")&IF(F41=1,"",IF(F36=1,"hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ","")&IF(F40=1,"bị can kề k
    #   hắc chế, ",""))&IF(OR(THONG_CAN!$F$69="Không có gốc",THONG_CAN!$F$69="Gốc dễ bị nhổ",THONG_CAN!$F$69="Gốc bị hợp (đứng độc lập)"),LOWER(THONG_CAN!$F$6
    #   9)&", ","")&IF(ABS(3-2)>=2,"cách xa nhật can, ",""))&"#",",#",""))&"; ","")))&" → Kết luận ④: "&IF(C65=0,"KHÔNG được trợ.",IF(D65=1,"CÓ và MẠNH.","CÓ 
    #   nhưng KHÔNG MẠNH."))
    return X.cat(X.cat(('Không có tỉ kiếp lộ ở thiên can (sách chỉ tính các thiên can cùng hành nhật can – P1 tr.70; tỉ kiếp tàng trong chi đã tính ở ② đắc địa).' if X.truthy(X.eq(X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Năm ',E.v(23,10,3)),' ('),E.v(23,12,3)),') là tỉ kiếp lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-66,c),1.0), X.OR(X.eq(E.v(9,r-69,c),1.0), X.eq(E.v(9,r-67,c),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-49,c),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,3)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,3),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-66,c),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-71,c),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-67,c),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,3)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,3),'Không có gốc'), X.eq(E.v(11,69,3),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,3),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(0.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-50,c),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Tháng ',E.v(23,10,4)),' ('),E.v(23,12,4)),') là tỉ kiếp lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-66,c+1),1.0), X.OR(X.eq(E.v(9,r-69,c+1),1.0), X.eq(E.v(9,r-67,c+1),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-49,c+1),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,4)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,4),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-66,c+1),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-71,c+1),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-67,c+1),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,4)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,4),'Không có gốc'), X.eq(E.v(11,69,4),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,4),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(1.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-50,c+1),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Giờ ',E.v(23,10,6)),' ('),E.v(23,12,6)),') là tỉ kiếp lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-66,c+3),1.0), X.OR(X.eq(E.v(9,r-69,c+3),1.0), X.eq(E.v(9,r-67,c+3),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-49,c+3),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,6)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,6),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-66,c+3),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-71,c+3),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-67,c+3),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,6)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,6),'Không có gốc'), X.eq(E.v(11,69,6),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,6),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(3.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-50,c+3),1.0)) else ''))),'')) else X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Năm ',E.v(23,10,3)),' ('),E.v(23,12,3)),') là tỉ kiếp lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-66,c),1.0), X.OR(X.eq(E.v(9,r-69,c),1.0), X.eq(E.v(9,r-67,c),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-49,c),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,3)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,3),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-66,c),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-71,c),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-67,c),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,3)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,3),'Không có gốc'), X.eq(E.v(11,69,3),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,3),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(0.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-50,c),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Tháng ',E.v(23,10,4)),' ('),E.v(23,12,4)),') là tỉ kiếp lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-66,c+1),1.0), X.OR(X.eq(E.v(9,r-69,c+1),1.0), X.eq(E.v(9,r-67,c+1),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-49,c+1),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,4)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,4),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-66,c+1),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-71,c+1),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-67,c+1),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,4)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,4),'Không có gốc'), X.eq(E.v(11,69,4),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,4),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(1.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-50,c+1),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can Giờ ',E.v(23,10,6)),' ('),E.v(23,12,6)),') là tỉ kiếp lộ'),(X.cat(' – MẠNH',(' (tuy hợp/bị khắc gần nhưng vượng tại lệnh và có gốc ở chi cùng trụ – gốc bền chặt, P1 tr.32; P2 tr.339)' if X.truthy(X.AND(X.eq(E.v(9,r-66,c+3),1.0), X.OR(X.eq(E.v(9,r-69,c+3),1.0), X.eq(E.v(9,r-67,c+3),1.0)))) else '')) if X.truthy(X.eq(E.v(9,r-49,c+3),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('không vượng tại lệnh (ở ',E.v(23,31,6)),' tại tháng), ') if X.truthy(X.gt(E.v(23,30,6),4.0)) else ''),('' if X.truthy(X.eq(E.v(9,r-66,c+3),1.0)) else X.cat(('hợp không hóa, bị khắc gần còn 2/3 lực (P2 tr.339), ' if X.truthy(X.eq(E.v(9,r-71,c+3),1.0)) else ''),('bị can kề khắc chế, ' if X.truthy(X.eq(E.v(9,r-67,c+3),1.0)) else '')))),(X.cat(X.LOWER(E.v(11,69,6)),', ') if X.truthy(X.OR(X.eq(E.v(11,69,6),'Không có gốc'), X.eq(E.v(11,69,6),'Gốc dễ bị nhổ'), X.eq(E.v(11,69,6),'Gốc bị hợp (đứng độc lập)'))) else '')),('cách xa nhật can, ' if X.truthy(X.ge(X.ABS(X.sub(3.0,2.0)),2.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-50,c+3),1.0)) else '')))),' → Kết luận ④: '),('KHÔNG được trợ.' if X.truthy(X.eq(E.v(9,r-42,c),0.0)) else ('CÓ và MẠNH.' if X.truthy(X.eq(E.v(9,r-42,c+1),1.0)) else 'CÓ nhưng KHÔNG MẠNH.')))

def f147(E, r, c):
    # Ô J107
    # = TRIM(IF(C54=1,"Dư khí "&INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" trong chi Năm "&TRUNG_GIAN!$C$11&" là ấn (dư khí – không tính mạnh); ","")&IF(D50=1,"Chi T
    #   háng "&TRUNG_GIAN!$D$11&" hóa thành ấn"&IF(D51=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(D29=1,D30=0),"chi bị xung, ","")&IF(ABS(1-2)>1,"
    #   trụ Tháng cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ","")&IF(D52=1,"Bản khí "&INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" trong chi Tháng
    #    "&TRUNG_GIAN!$D$11&" là ấn"&IF(D55=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(INDEX(M_TS,TRUNG_GIAN!$D$15+1,TRUNG_GIAN!$D$6+1)>4,"không vượng
    #    tại lệnh, ","")&IF(AND(D29=1,D30=0),"chi bị xung, ","")&IF(ABS(1-2)>1,"trụ Tháng cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ","")
    #   &IF(D53=1,"Trung khí "&INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" trong chi Tháng "&TRUNG_GIAN!$D$11&" là ấn (trung khí – không tính mạnh); ","")&IF(D54=1,"D
    #   ư khí "&INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" trong chi Tháng "&TRUNG_GIAN!$D$11&" là ấn (dư khí – không tính mạnh); ","")&IF(E50=1,"Chi Ngày "&TRUNG_GI
    #   AN!$E$11&" hóa thành ấn"&IF(E51=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(AND(E29=1,E30=0),"chi bị xung, ","")&IF(ABS(2-2)>1,"trụ Ngày cách x
    #   a trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('Dư khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0))),' trong chi Năm '),E.v(23,11,3)),' là ấn (dư khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-53,c-7),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat('Chi Tháng ',E.v(23,11,4)),' hóa thành ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-56,c-6),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-78,c-6),1.0), X.eq(E.v(9,r-77,c-6),0.0))) else ''),('trụ Tháng cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(1.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-57,c-6),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bản khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0))),' trong chi Tháng '),E.v(23,11,4)),' là ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-52,c-6),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(('không vượng tại lệnh, ' if X.truthy(X.gt(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else ''),('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-78,c-6),1.0), X.eq(E.v(9,r-77,c-6),0.0))) else '')),('trụ Tháng cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(1.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-55,c-6),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Trung khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0))),' trong chi Tháng '),E.v(23,11,4)),' là ấn (trung khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-54,c-6),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Dư khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0))),' trong chi Tháng '),E.v(23,11,4)),' là ấn (dư khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-53,c-6),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Chi Ngày ',E.v(23,11,5)),' hóa thành ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-56,c-5),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-78,c-5),1.0), X.eq(E.v(9,r-77,c-5),0.0))) else ''),('trụ Ngày cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(2.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-57,c-5),1.0)) else '')))

def f148(E, r, c):
    # Ô C108
    # = "Được lệnh: "&IF(C62=1,"CÓ","KHÔNG")&"; đắc địa: "&IF(C63=0,"không",IF(D63=1,"có – mạnh","có – không mạnh"))&"; được sinh: "&IF(C64=0,"không",IF(D64=1
    #   ,"có – mạnh","có – không mạnh"))&"; được trợ: "&IF(C65=0,"không",IF(D65=1,"có – mạnh","có – không mạnh"))&" → có "&C66&"/3 điều ②③④, trong đó "&D66&" 
    #   điều mạnh."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Được lệnh: ',('CÓ' if X.truthy(X.eq(E.v(9,r-46,c),1.0)) else 'KHÔNG')),'; đắc địa: '),('không' if X.truthy(X.eq(E.v(9,r-45,c),0.0)) else ('có – mạnh' if X.truthy(X.eq(E.v(9,r-45,c+1),1.0)) else 'có – không mạnh'))),'; được sinh: '),('không' if X.truthy(X.eq(E.v(9,r-44,c),0.0)) else ('có – mạnh' if X.truthy(X.eq(E.v(9,r-44,c+1),1.0)) else 'có – không mạnh'))),'; được trợ: '),('không' if X.truthy(X.eq(E.v(9,r-43,c),0.0)) else ('có – mạnh' if X.truthy(X.eq(E.v(9,r-43,c+1),1.0)) else 'có – không mạnh'))),' → có '),E.v(9,r-42,c)),'/3 điều ②③④, trong đó '),E.v(9,r-42,c+1)),' điều mạnh.')

def f149(E, r, c):
    # Ô J108
    # = TRIM(IF(E52=1,"Bản khí "&INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" trong chi Ngày "&TRUNG_GIAN!$E$11&" là ấn"&IF(E55=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTIT
    #   UTE(TRIM(IF(INDEX(M_TS,TRUNG_GIAN!$E$15+1,TRUNG_GIAN!$D$6+1)>4,"không vượng tại lệnh, ","")&IF(AND(E29=1,E30=0),"chi bị xung, ","")&IF(ABS(2-2)>1,"trụ
    #    Ngày cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ","")&IF(E53=1,"Trung khí "&INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" trong chi Ngày "&
    #   TRUNG_GIAN!$E$11&" là ấn (trung khí – không tính mạnh); ","")&IF(E54=1,"Dư khí "&INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" trong chi Ngày "&TRUNG_GIAN!$E$11
    #   &" là ấn (dư khí – không tính mạnh); ","")&IF(F50=1,"Chi Giờ "&TRUNG_GIAN!$F$11&" hóa thành ấn"&IF(F51=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM
    #   (IF(AND(F29=1,F30=0),"chi bị xung, ","")&IF(ABS(3-2)>1,"trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ","")&IF(F52=1,"Bản khí
    #    "&INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" trong chi Giờ "&TRUNG_GIAN!$F$11&" là ấn"&IF(F55=1," – MẠNH"," – KHÔNG MẠNH vì "&SUBSTITUTE(TRIM(IF(INDEX(M_TS,
    #   TRUNG_GIAN!$F$15+1,TRUNG_GIAN!$D$6+1)>4,"không vượng tại lệnh, ","")&IF(AND(F29=1,F30=0),"chi bị xung, ","")&IF(ABS(3-2)>1,"trụ Giờ cách xa trụ ngày →
    #    lực yếu (P1 tr.71), ",""))&"#",",#",""))&"; ","")&IF(F53=1,"Trung khí "&INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" trong chi Giờ "&TRUNG_GIAN!$F$11&" là ấn 
    #   (trung khí – không tính mạnh); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bản khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0))),' trong chi Ngày '),E.v(23,11,5)),' là ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-53,c-5),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(('không vượng tại lệnh, ' if X.truthy(X.gt(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else ''),('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-79,c-5),1.0), X.eq(E.v(9,r-78,c-5),0.0))) else '')),('trụ Ngày cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(2.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-56,c-5),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat('Trung khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0))),' trong chi Ngày '),E.v(23,11,5)),' là ấn (trung khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-55,c-5),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Dư khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0))),' trong chi Ngày '),E.v(23,11,5)),' là ấn (dư khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-54,c-5),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Chi Giờ ',E.v(23,11,6)),' hóa thành ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-57,c-4),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-79,c-4),1.0), X.eq(E.v(9,r-78,c-4),0.0))) else ''),('trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(3.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-58,c-4),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bản khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0))),' trong chi Giờ '),E.v(23,11,6)),' là ấn'),(' – MẠNH' if X.truthy(X.eq(E.v(9,r-53,c-4),1.0)) else X.cat(' – KHÔNG MẠNH vì ',X.SUBSTITUTE(X.cat(X.TRIM(X.cat(X.cat(('không vượng tại lệnh, ' if X.truthy(X.gt(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else ''),('chi bị xung, ' if X.truthy(X.AND(X.eq(E.v(9,r-79,c-4),1.0), X.eq(E.v(9,r-78,c-4),0.0))) else '')),('trụ Giờ cách xa trụ ngày → lực yếu (P1 tr.71), ' if X.truthy(X.gt(X.ABS(X.sub(3.0,2.0)),1.0)) else ''))),'#'), ',#', '')))),'; ') if X.truthy(X.eq(E.v(9,r-56,c-4),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Trung khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0))),' trong chi Giờ '),E.v(23,11,6)),' là ấn (trung khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-55,c-4),1.0)) else '')))

def f150(E, r, c):
    # Ô C109
    # = "Quy tắc áp dụng: "&C92&". "&IF(ISNUMBER(SEARCH("phân xử",C92)),"Vì không được lệnh mà vẫn có ≥ 2 điều nhưng các điều đó không đủ mạnh, sách gọi là ""
    #   nửa vượng nửa nhược"" và phải phân xử theo 4 mặt (P1 tr.71): a) "&IF(C71>0,"hợp hóa tạo ra hành sinh/trợ nhật chủ → +1",IF(C71<0,"hợp hóa tạo ra hành 
    #   khắc/tiết/hao nhật chủ → −1","không có hợp hóa liên quan → 0"))&"; b) "&IF(C72>0,"bán hợp sinh/trợ nhật chủ → +1",IF(C72<0,"bán hợp khắc/tiết/hao nhật
    #    chủ → −1","không có bán hợp liên quan → 0"))&IF(LEN(C125)>0," (có tính bán hội "&C125&" – suy luận ngoài sách, chưa được sách xác nhận)","")&"; c) ph
    #   e mình có "&C82&" yếu tố vượng tại lệnh, phe khắc–tiết–hao có "&C83&" → "&IF(C84>0,"+1",IF(C84<0,"−1","0"))&"; d) tỉ lệ yếu tố phe mình bị vô hiệu (hợ
    #   p mất, khắc chế, xung, ở xa) "&TEXT(C85,"0%")&", phe kia "&TEXT(C86,"0%")&" → "&IF(C87>0,"+1",IF(C87<0,"−1","0"))&". Tổng "&C88&" → "&IF(C88>0,"nghiên
    #   g VƯỢNG",IF(C88<0,"nghiêng NHƯỢC","cân bằng – TRUNG HÒA"))&".",IF(C62=1,"Được lệnh là điều kiện nặng nhất; chỉ cần thêm 1 trong 3 điều còn lại là vượn
    #   g (P1 tr.70).","Không được lệnh thì phải có ≥ 2 điều và các điều đó mạnh mới là vượng (P1 tr.70-71)."))
    return X.cat(X.cat(X.cat('Quy tắc áp dụng: ',E.v(9,r-17,c)),'. '),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vì không được lệnh mà vẫn có ≥ 2 điều nhưng các điều đó không đủ mạnh, sách gọi là "nửa vượng nửa nhược" và phải phân xử theo 4 mặt (P1 tr.71): a) ',('hợp hóa tạo ra hành sinh/trợ nhật chủ → +1' if X.truthy(X.gt(E.v(9,r-38,c),0.0)) else ('hợp hóa tạo ra hành khắc/tiết/hao nhật chủ → −1' if X.truthy(X.lt(E.v(9,r-38,c),0.0)) else 'không có hợp hóa liên quan → 0'))),'; b) '),('bán hợp sinh/trợ nhật chủ → +1' if X.truthy(X.gt(E.v(9,r-37,c),0.0)) else ('bán hợp khắc/tiết/hao nhật chủ → −1' if X.truthy(X.lt(E.v(9,r-37,c),0.0)) else 'không có bán hợp liên quan → 0'))),(X.cat(X.cat(' (có tính bán hội ',E.v(9,r+16,c)),' – suy luận ngoài sách, chưa được sách xác nhận)') if X.truthy(X.gt(X.LEN(E.v(9,r+16,c)),0.0)) else '')),'; c) phe mình có '),E.v(9,r-27,c)),' yếu tố vượng tại lệnh, phe khắc–tiết–hao có '),E.v(9,r-26,c)),' → '),('+1' if X.truthy(X.gt(E.v(9,r-25,c),0.0)) else ('−1' if X.truthy(X.lt(E.v(9,r-25,c),0.0)) else '0'))),'; d) tỉ lệ yếu tố phe mình bị vô hiệu (hợp mất, khắc chế, xung, ở xa) '),X.TEXT(E.v(9,r-24,c), '0%')),', phe kia '),X.TEXT(E.v(9,r-23,c), '0%')),' → '),('+1' if X.truthy(X.gt(E.v(9,r-22,c),0.0)) else ('−1' if X.truthy(X.lt(E.v(9,r-22,c),0.0)) else '0'))),'. Tổng '),E.v(9,r-21,c)),' → '),('nghiêng VƯỢNG' if X.truthy(X.gt(E.v(9,r-21,c),0.0)) else ('nghiêng NHƯỢC' if X.truthy(X.lt(E.v(9,r-21,c),0.0)) else 'cân bằng – TRUNG HÒA'))),'.') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('phân xử', E.v(9,r-17,c)))) else ('Được lệnh là điều kiện nặng nhất; chỉ cần thêm 1 trong 3 điều còn lại là vượng (P1 tr.70).' if X.truthy(X.eq(E.v(9,r-47,c),1.0)) else 'Không được lệnh thì phải có ≥ 2 điều và các điều đó mạnh mới là vượng (P1 tr.70-71).')))

def f151(E, r, c):
    # Ô J109
    # = TRIM(IF(F54=1,"Dư khí "&INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" trong chi Giờ "&TRUNG_GIAN!$F$11&" là ấn (dư khí – không tính mạnh); ",""))
    return X.TRIM((X.cat(X.cat(X.cat(X.cat('Dư khí ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0))),' trong chi Giờ '),E.v(23,11,6)),' là ấn (dư khí – không tính mạnh); ') if X.truthy(X.eq(E.v(9,r-55,c-4),1.0)) else ''))

def f152(E, r, c):
    # Ô C110
    # = "Nhật chủ "&TRUNG_GIAN!$E$10&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&") – "&UPPER(C91)&". Lý do: "&IF(C62=1,"được lệnh","không được lệnh")&", "&C66&"/3
    #    điều còn lại có mặt, "&D66&" điều mạnh"&IF(ISNUMBER(SEARCH("phân xử",C92)),"; phân xử a–d được "&C88,"")&". "&IF(OR(C91="Thân nhược",C91="Quá nhược")
    #   ,"Nhật chủ yếu → cần hành sinh – trợ (ấn, tỉ kiếp) làm dụng thần.",IF(C91="Trung hòa","Nhật chủ cân bằng → xem DUNG_HY_KY phần trung hòa.","Nhật chủ m
    #   ạnh → cần hành khắc – tiết – hao (quan sát, thực thương, tài) làm dụng thần."))&IF($C$117<>0," Bước 5: "&$C$118,"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',E.v(23,10,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),') – '),X.UPPER(E.v(9,r-19,c))),'. Lý do: '),('được lệnh' if X.truthy(X.eq(E.v(9,r-48,c),1.0)) else 'không được lệnh')),', '),E.v(9,r-44,c)),'/3 điều còn lại có mặt, '),E.v(9,r-44,c+1)),' điều mạnh'),(X.cat('; phân xử a–d được ',E.v(9,r-22,c)) if X.truthy(X.ISNUMBER(lambda: X.SEARCH('phân xử', E.v(9,r-18,c)))) else '')),'. '),('Nhật chủ yếu → cần hành sinh – trợ (ấn, tỉ kiếp) làm dụng thần.' if X.truthy(X.OR(X.eq(E.v(9,r-19,c),'Thân nhược'), X.eq(E.v(9,r-19,c),'Quá nhược'))) else ('Nhật chủ cân bằng → xem DUNG_HY_KY phần trung hòa.' if X.truthy(X.eq(E.v(9,r-19,c),'Trung hòa')) else 'Nhật chủ mạnh → cần hành khắc – tiết – hao (quan sát, thực thương, tài) làm dụng thần.'))),(X.cat(' Bước 5: ',E.v(9,118,3)) if X.truthy(X.ne(E.v(9,117,3),0.0)) else ''))

def f153(E, r, c):
    # Ô C111
    # = IF($C$117=2,"⚠ VƯỢNG CẬN RANH GIỚI: "&$C$118,"")
    return (X.cat('⚠ VƯỢNG CẬN RANH GIỚI: ',E.v(9,118,3)) if X.truthy(X.eq(E.v(9,117,3),2.0)) else '')

def f154(E, r, c):
    # Ô C112
    # = IF(OR(C111<>"",AND(ISNUMBER(SEARCH("phân xử",C92)),C88=0)),"NEED_REVIEW – cần luận giải (P1 tr.71): "&IF(C111<>"","vượng cận ranh giới (bước 5)","nửa 
    #   vượng nửa nhược, phân xử a–d hòa – kết luận theo đếm chữ"),"AUTO_CONCLUSION – kết luận theo quy tắc P1 tr.70-71")
    return (X.cat('NEED_REVIEW – cần luận giải (P1 tr.71): ',('vượng cận ranh giới (bước 5)' if X.truthy(X.ne(E.v(9,r-1,c),'')) else 'nửa vượng nửa nhược, phân xử a–d hòa – kết luận theo đếm chữ')) if X.truthy(X.OR(X.ne(E.v(9,r-1,c),''), X.AND(X.ISNUMBER(lambda: X.SEARCH('phân xử', E.v(9,r-20,c))), X.eq(E.v(9,r-24,c),0.0)))) else 'AUTO_CONCLUSION – kết luận theo quy tắc P1 tr.70-71')

def f155(E, r, c):
    # Ô C115
    # = IF(AND(C62=1,C66=3),"Quá vượng",IF(AND(C62=1,C66=2),"Rất vượng",IF(AND(C62=1,C66=1),"Thân vượng",IF(AND(C62=1,C66=0),"Thân nhược",IF(AND(C66>=2,D66>=2
    #   ),IF(D66=3,"Rất vượng","Thân vượng"),IF(AND(C66=1,C67=1),"Thân vượng",IF(C66>=2,IF(C88>0,"Thân vượng",IF(C88<0,"Thân nhược",IF(C89>D89,"Thân vượng",IF
    #   (C89<D89,"Thân nhược","Trung hòa")))),IF(AND(C66=0,C68=1),"Quá nhược","Thân nhược"))))))))
    return ('Quá vượng' if X.truthy(X.AND(X.eq(E.v(9,r-53,c),1.0), X.eq(E.v(9,r-49,c),3.0))) else ('Rất vượng' if X.truthy(X.AND(X.eq(E.v(9,r-53,c),1.0), X.eq(E.v(9,r-49,c),2.0))) else ('Thân vượng' if X.truthy(X.AND(X.eq(E.v(9,r-53,c),1.0), X.eq(E.v(9,r-49,c),1.0))) else ('Thân nhược' if X.truthy(X.AND(X.eq(E.v(9,r-53,c),1.0), X.eq(E.v(9,r-49,c),0.0))) else (('Rất vượng' if X.truthy(X.eq(E.v(9,r-49,c+1),3.0)) else 'Thân vượng') if X.truthy(X.AND(X.ge(E.v(9,r-49,c),2.0), X.ge(E.v(9,r-49,c+1),2.0))) else ('Thân vượng' if X.truthy(X.AND(X.eq(E.v(9,r-49,c),1.0), X.eq(E.v(9,r-48,c),1.0))) else (('Thân vượng' if X.truthy(X.gt(E.v(9,r-27,c),0.0)) else ('Thân nhược' if X.truthy(X.lt(E.v(9,r-27,c),0.0)) else ('Thân vượng' if X.truthy(X.gt(E.v(9,r-26,c),E.v(9,r-26,c+1))) else ('Thân nhược' if X.truthy(X.lt(E.v(9,r-26,c),E.v(9,r-26,c+1))) else 'Trung hòa')))) if X.truthy(X.ge(E.v(9,r-49,c),2.0)) else ('Quá nhược' if X.truthy(X.AND(X.eq(E.v(9,r-49,c),0.0), X.eq(E.v(9,r-47,c),1.0))) else 'Thân nhược'))))))))

def f156(E, r, c):
    # Ô C116
    # = IF(AND(C62=1,C66=3),"Được lệnh + đủ 3 điều (P1 tr.70)",IF(AND(C62=1,C66=2),"Được lệnh + 2 trong 3 điều (P1 tr.70)",IF(AND(C62=1,C66=1),"Được lệnh + 1 
    #   trong 3 điều (P1 tr.70)",IF(AND(C62=1,C66=0),"Được lệnh nhưng không đắc địa, không được sinh, không được trợ → không được xem là vượng (P1 tr.70)",IF(
    #   AND(C66>=2,D66>=2),"Không lệnh nhưng có ≥2 điều và các điều đó mạnh (P1 tr.70-71, quy tắc 3, 5)",IF(AND(C66=1,C67=1),"Không lệnh, 1 điều, có tam hợp/t
    #   am hội ấn cục hoặc tỉ cục (P1 tr.70, quy tắc 4)",IF(C66>=2,IF(C88=0,"Không lệnh, ≥2 điều nhưng không mạnh → phân xử a–d hòa (0) → đếm chữ: phe mình "&
    #   C89&", phe khác "&D89&" (P1 tr.71; P2 tr.455-457)","Không lệnh, ≥2 điều nhưng không mạnh → nửa vượng nửa nhược, phân xử a–d, tổng "&C88&" (P1 tr.71)")
    #   ,IF(AND(C66=0,C68=1),"Không lệnh, không đắc địa, không có ấn, không có tỉ kiếp (P1 tr.172)","Không lệnh, ≤1 điều (P1 tr.70)"))))))))
    return ('Được lệnh + đủ 3 điều (P1 tr.70)' if X.truthy(X.AND(X.eq(E.v(9,r-54,c),1.0), X.eq(E.v(9,r-50,c),3.0))) else ('Được lệnh + 2 trong 3 điều (P1 tr.70)' if X.truthy(X.AND(X.eq(E.v(9,r-54,c),1.0), X.eq(E.v(9,r-50,c),2.0))) else ('Được lệnh + 1 trong 3 điều (P1 tr.70)' if X.truthy(X.AND(X.eq(E.v(9,r-54,c),1.0), X.eq(E.v(9,r-50,c),1.0))) else ('Được lệnh nhưng không đắc địa, không được sinh, không được trợ → không được xem là vượng (P1 tr.70)' if X.truthy(X.AND(X.eq(E.v(9,r-54,c),1.0), X.eq(E.v(9,r-50,c),0.0))) else ('Không lệnh nhưng có ≥2 điều và các điều đó mạnh (P1 tr.70-71, quy tắc 3, 5)' if X.truthy(X.AND(X.ge(E.v(9,r-50,c),2.0), X.ge(E.v(9,r-50,c+1),2.0))) else ('Không lệnh, 1 điều, có tam hợp/tam hội ấn cục hoặc tỉ cục (P1 tr.70, quy tắc 4)' if X.truthy(X.AND(X.eq(E.v(9,r-50,c),1.0), X.eq(E.v(9,r-49,c),1.0))) else ((X.cat(X.cat(X.cat(X.cat('Không lệnh, ≥2 điều nhưng không mạnh → phân xử a–d hòa (0) → đếm chữ: phe mình ',E.v(9,r-27,c)),', phe khác '),E.v(9,r-27,c+1)),' (P1 tr.71; P2 tr.455-457)') if X.truthy(X.eq(E.v(9,r-28,c),0.0)) else X.cat(X.cat('Không lệnh, ≥2 điều nhưng không mạnh → nửa vượng nửa nhược, phân xử a–d, tổng ',E.v(9,r-28,c)),' (P1 tr.71)')) if X.truthy(X.ge(E.v(9,r-50,c),2.0)) else ('Không lệnh, không đắc địa, không có ấn, không có tỉ kiếp (P1 tr.172)' if X.truthy(X.AND(X.eq(E.v(9,r-50,c),0.0), X.eq(E.v(9,r-48,c),1.0))) else 'Không lệnh, ≤1 điều (P1 tr.70)'))))))))

def f157(E, r, c):
    # Ô C117
    # = IF(ISNUMBER(SEARCH("phân xử",$C$116)),0,IF($C$62=0,IF(AND(ISNUMBER(SEARCH("vượng",$C$115)),$C$88<0,$D$89>$C$89),-1,IF(AND(NOT(ISNUMBER(SEARCH("vượng",
    #   $C$115))),$C$88>0,$C$89>$D$89),1,0)),IF(AND(ISNUMBER(SEARCH("vượng",$C$115)),$C$88<0,$D$89>$C$89),2,0)))
    return (0.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('phân xử', E.v(9,116,3)))) else ((X.neg(1.0) if X.truthy(X.AND(X.ISNUMBER(lambda: X.SEARCH('vượng', E.v(9,115,3))), X.lt(E.v(9,88,3),0.0), X.gt(E.v(9,89,4),E.v(9,89,3)))) else (1.0 if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('vượng', E.v(9,115,3)))), X.gt(E.v(9,88,3),0.0), X.gt(E.v(9,89,3),E.v(9,89,4)))) else 0.0)) if X.truthy(X.eq(E.v(9,62,3),0.0)) else (2.0 if X.truthy(X.AND(X.ISNUMBER(lambda: X.SEARCH('vượng', E.v(9,115,3))), X.lt(E.v(9,88,3),0.0), X.gt(E.v(9,89,4),E.v(9,89,3)))) else 0.0)))

def f158(E, r, c):
    # Ô C118
    # = IF(ISNUMBER(SEARCH("phân xử",$C$116)),"Kết luận sơ bộ đã dựa trên phân xử a–d (nửa vượng nửa nhược) nên giữ nguyên.",CHOOSE($C$117+2,"Không được lệnh;
    #    các điều sinh trợ không mạnh, không giúp ích nhiều: phân xử a–d nghiêng về phe khắc (tổng "&$C$88&") và phe khắc–tiết–hao nhiều chữ hơn ("&$D$89&" so
    #    với "&$C$89&") → đổi thành THÂN NHƯỢC.",IF($C$62=1,"Được lệnh; phân xử a–d và đếm chữ không cùng phủ định kết luận sơ bộ → giữ nguyên.","Phân xử a–d 
    #   và đếm chữ không cùng phủ định kết luận sơ bộ → giữ nguyên."),"Không được lệnh nhưng phân xử a–d nghiêng về phe sinh–trợ (tổng "&$C$88&") và phe sinh–
    #   trợ nhiều chữ hơn ("&$C$89&" so với "&$D$89&") → đổi thành THÂN VƯỢNG.","Được lệnh nên giữ THÂN VƯỢNG (lệnh là quan trọng nhất), nhưng phân xử a–d (tổ
    #   ng "&$C$88&") và đếm chữ ("&$D$89&" chữ phe khắc so với "&$C$89&") đều nghiêng về phe khắc–tiết–hao → vượng cận ranh giới: gặp vận khắc–tiết dễ chuyển
    #    nhược."))
    return ('Kết luận sơ bộ đã dựa trên phân xử a–d (nửa vượng nửa nhược) nên giữ nguyên.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('phân xử', E.v(9,116,3)))) else X.CHOOSE(X.add(E.v(9,117,3),2.0), lambda: X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không được lệnh; các điều sinh trợ không mạnh, không giúp ích nhiều: phân xử a–d nghiêng về phe khắc (tổng ',E.v(9,88,3)),') và phe khắc–tiết–hao nhiều chữ hơn ('),E.v(9,89,4)),' so với '),E.v(9,89,3)),') → đổi thành THÂN NHƯỢC.'), lambda: ('Được lệnh; phân xử a–d và đếm chữ không cùng phủ định kết luận sơ bộ → giữ nguyên.' if X.truthy(X.eq(E.v(9,62,3),1.0)) else 'Phân xử a–d và đếm chữ không cùng phủ định kết luận sơ bộ → giữ nguyên.'), lambda: X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không được lệnh nhưng phân xử a–d nghiêng về phe sinh–trợ (tổng ',E.v(9,88,3)),') và phe sinh–trợ nhiều chữ hơn ('),E.v(9,89,3)),' so với '),E.v(9,89,4)),') → đổi thành THÂN VƯỢNG.'), lambda: X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Được lệnh nên giữ THÂN VƯỢNG (lệnh là quan trọng nhất), nhưng phân xử a–d (tổng ',E.v(9,88,3)),') và đếm chữ ('),E.v(9,89,4)),' chữ phe khắc so với '),E.v(9,89,3)),') đều nghiêng về phe khắc–tiết–hao → vượng cận ranh giới: gặp vận khắc–tiết dễ chuyển nhược.')))

def f159(E, r, c):
    # Ô C119
    # = "P1 tr.70 (quy tắc 3: các điều phải mạnh, giúp ích nhiều); P1 tr.71–72 (so sánh mức độ được sinh – bị khắc; phán đoán tổng hợp)"
    return 'P1 tr.70 (quy tắc 3: các điều phải mạnh, giúp ích nhiều); P1 tr.71–72 (so sánh mức độ được sinh – bị khắc; phán đoán tổng hợp)'

def f160(E, r, c):
    # Ô C122, D122
    # = IF(OR(AND(TRUNG_GIAN!C$6=2,TRUNG_GIAN!D$6=3),AND(TRUNG_GIAN!C$6=3,TRUNG_GIAN!D$6=2)),0,IF(OR(AND(TRUNG_GIAN!C$6=5,TRUNG_GIAN!D$6=6),AND(TRUNG_GIAN!C$6
    #   =6,TRUNG_GIAN!D$6=5)),1,IF(OR(AND(TRUNG_GIAN!C$6=8,TRUNG_GIAN!D$6=9),AND(TRUNG_GIAN!C$6=9,TRUNG_GIAN!D$6=8)),3,IF(OR(AND(TRUNG_GIAN!C$6=11,TRUNG_GIAN!
    #   D$6=0),AND(TRUNG_GIAN!C$6=0,TRUNG_GIAN!D$6=11)),4,-1))))
    return (0.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),2.0), X.eq(E.v(23,6,c+1),3.0)), X.AND(X.eq(E.v(23,6,c),3.0), X.eq(E.v(23,6,c+1),2.0)))) else (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),5.0), X.eq(E.v(23,6,c+1),6.0)), X.AND(X.eq(E.v(23,6,c),6.0), X.eq(E.v(23,6,c+1),5.0)))) else (3.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),8.0), X.eq(E.v(23,6,c+1),9.0)), X.AND(X.eq(E.v(23,6,c),9.0), X.eq(E.v(23,6,c+1),8.0)))) else (4.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),11.0), X.eq(E.v(23,6,c+1),0.0)), X.AND(X.eq(E.v(23,6,c),0.0), X.eq(E.v(23,6,c+1),11.0)))) else X.neg(1.0)))))

def f161(E, r, c):
    # Ô E122
    # = IF(NHAP_LIEU!$C$9="Không rõ giờ",-1,IF(OR(AND(TRUNG_GIAN!E$6=2,TRUNG_GIAN!F$6=3),AND(TRUNG_GIAN!E$6=3,TRUNG_GIAN!F$6=2)),0,IF(OR(AND(TRUNG_GIAN!E$6=5,
    #   TRUNG_GIAN!F$6=6),AND(TRUNG_GIAN!E$6=6,TRUNG_GIAN!F$6=5)),1,IF(OR(AND(TRUNG_GIAN!E$6=8,TRUNG_GIAN!F$6=9),AND(TRUNG_GIAN!E$6=9,TRUNG_GIAN!F$6=8)),3,IF(
    #   OR(AND(TRUNG_GIAN!E$6=11,TRUNG_GIAN!F$6=0),AND(TRUNG_GIAN!E$6=0,TRUNG_GIAN!F$6=11)),4,-1)))))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else (0.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),2.0), X.eq(E.v(23,6,c+1),3.0)), X.AND(X.eq(E.v(23,6,c),3.0), X.eq(E.v(23,6,c+1),2.0)))) else (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),5.0), X.eq(E.v(23,6,c+1),6.0)), X.AND(X.eq(E.v(23,6,c),6.0), X.eq(E.v(23,6,c+1),5.0)))) else (3.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),8.0), X.eq(E.v(23,6,c+1),9.0)), X.AND(X.eq(E.v(23,6,c),9.0), X.eq(E.v(23,6,c+1),8.0)))) else (4.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,c),11.0), X.eq(E.v(23,6,c+1),0.0)), X.AND(X.eq(E.v(23,6,c),0.0), X.eq(E.v(23,6,c+1),11.0)))) else X.neg(1.0))))))

def f162(E, r, c):
    # Ô C123, D123, E123
    # = IF(AND(C122>=0,TRUNG_GIAN!C56=1,TRUNG_GIAN!C57=0),1,0)
    return (1.0 if X.truthy(X.AND(X.ge(E.v(9,r-1,c),0.0), X.eq(E.v(23,r-67,c),1.0), X.eq(E.v(23,r-66,c),0.0))) else 0.0)

def f163(E, r, c):
    # Ô C124, D124, E124
    # = IF(C123=1,INDEX(CHI_TEN,TRUNG_GIAN!C$6+1)&"–"&INDEX(CHI_TEN,TRUNG_GIAN!D$6+1)&" ("&INDEX(HANH_TEN,C122+1)&")","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,c),1.0)),'–'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,c+1),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(9,r-2,c),1.0))),')') if X.truthy(X.eq(E.v(9,r-1,c),1.0)) else '')

def f164(E, r, c):
    # Ô C125
    # = TRIM(C124&" "&D124&" "&E124)
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(E.v(9,r-1,c),' '),E.v(9,r-1,c+1)),' '),E.v(9,r-1,c+2)))

CELLS = {
    'C6': (f0, 6, 3),
    'C7': (f1, 7, 3),
    'D7': (f2, 7, 4),
    'E7': (f2, 7, 5),
    'F7': (f3, 7, 6),
    'C8': (f4, 8, 3),
    'C10': (f5, 10, 3),
    'D10': (f5, 10, 4),
    'E10': (f5, 10, 5),
    'C11': (f6, 11, 3),
    'D11': (f7, 11, 4),
    'E11': (f8, 11, 5),
    'C12': (f9, 12, 3),
    'D12': (f9, 12, 4),
    'E12': (f9, 12, 5),
    'C13': (f10, 13, 3),
    'D13': (f10, 13, 4),
    'E13': (f10, 13, 5),
    'C14': (f11, 14, 3),
    'D14': (f12, 14, 4),
    'E14': (f12, 14, 5),
    'F14': (f13, 14, 6),
    'C15': (f14, 15, 3),
    'D15': (f15, 15, 4),
    'E15': (f15, 15, 5),
    'F15': (f16, 15, 6),
    'C16': (f17, 16, 3),
    'D16': (f18, 16, 4),
    'E16': (f18, 16, 5),
    'F16': (f19, 16, 6),
    'C18': (f20, 18, 3),
    'D18': (f20, 18, 4),
    'E18': (f20, 18, 5),
    'C19': (f21, 19, 3),
    'D19': (f22, 19, 4),
    'E19': (f22, 19, 5),
    'F19': (f23, 19, 6),
    'C20': (f24, 20, 3),
    'D20': (f24, 20, 4),
    'E20': (f24, 20, 5),
    'C21': (f25, 21, 3),
    'D21': (f25, 21, 4),
    'E21': (f25, 21, 5),
    'C22': (f26, 22, 3),
    'D22': (f27, 22, 4),
    'E22': (f27, 22, 5),
    'F22': (f28, 22, 6),
    'C24': (f29, 24, 3),
    'D24': (f29, 24, 4),
    'E24': (f29, 24, 5),
    'C25': (f21, 25, 3),
    'D25': (f22, 25, 4),
    'E25': (f22, 25, 5),
    'F25': (f23, 25, 6),
    'C26': (f30, 26, 3),
    'D26': (f30, 26, 4),
    'E26': (f30, 26, 5),
    'C27': (f31, 27, 3),
    'D27': (f32, 27, 4),
    'E27': (f32, 27, 5),
    'F27': (f33, 27, 6),
    'C28': (f34, 28, 3),
    'D28': (f35, 28, 4),
    'E28': (f35, 28, 5),
    'F28': (f36, 28, 6),
    'C29': (f37, 29, 3),
    'D29': (f37, 29, 4),
    'E29': (f37, 29, 5),
    'F29': (f37, 29, 6),
    'C30': (f38, 30, 3),
    'D30': (f39, 30, 4),
    'E30': (f39, 30, 5),
    'F30': (f40, 30, 6),
    'C33': (f41, 33, 3),
    'D33': (f41, 33, 4),
    'E33': (f41, 33, 5),
    'C34': (f42, 34, 3),
    'D34': (f43, 34, 4),
    'E34': (f43, 34, 5),
    'C35': (f44, 35, 3),
    'D35': (f45, 35, 4),
    'E35': (f45, 35, 5),
    'C36': (f46, 36, 3),
    'D36': (f47, 36, 4),
    'E36': (f48, 36, 5),
    'F36': (f49, 36, 6),
    'C37': (f50, 37, 3),
    'D37': (f50, 37, 4),
    'E37': (f50, 37, 5),
    'F37': (f50, 37, 6),
    'C38': (f51, 38, 3),
    'D38': (f52, 38, 4),
    'E38': (f52, 38, 5),
    'F38': (f53, 38, 6),
    'C39': (f54, 39, 3),
    'D39': (f55, 39, 4),
    'E39': (f55, 39, 5),
    'F39': (f56, 39, 6),
    'C40': (f57, 40, 3),
    'D40': (f58, 40, 4),
    'E40': (f58, 40, 5),
    'F40': (f59, 40, 6),
    'C41': (f60, 41, 3),
    'D41': (f60, 41, 4),
    'E41': (f48, 41, 5),
    'F41': (f60, 41, 6),
    'C43': (f61, 43, 3),
    'D43': (f62, 43, 4),
    'C45': (f63, 45, 3),
    'D45': (f48, 45, 4),
    'E45': (f63, 45, 5),
    'F45': (f63, 45, 6),
    'C46': (f64, 46, 3),
    'D46': (f48, 46, 4),
    'E46': (f65, 46, 5),
    'F46': (f66, 46, 6),
    'C47': (f67, 47, 3),
    'AA47': (f68, 47, 27),
    'AB47': (f69, 47, 28),
    'AC47': (f70, 47, 29),
    'AD47': (f71, 47, 30),
    'C48': (f72, 48, 3),
    'D48': (f72, 48, 4),
    'E48': (f48, 48, 5),
    'F48': (f72, 48, 6),
    'C49': (f73, 49, 3),
    'D49': (f74, 49, 4),
    'E49': (f48, 49, 5),
    'F49': (f75, 49, 6),
    'C50': (f76, 50, 3),
    'D50': (f76, 50, 4),
    'E50': (f76, 50, 5),
    'F50': (f76, 50, 6),
    'C51': (f77, 51, 3),
    'D51': (f78, 51, 4),
    'E51': (f79, 51, 5),
    'F51': (f80, 51, 6),
    'C52': (f81, 52, 3),
    'D52': (f81, 52, 4),
    'E52': (f81, 52, 5),
    'F52': (f81, 52, 6),
    'C53': (f82, 53, 3),
    'D53': (f82, 53, 4),
    'E53': (f82, 53, 5),
    'F53': (f82, 53, 6),
    'C54': (f83, 54, 3),
    'D54': (f83, 54, 4),
    'E54': (f83, 54, 5),
    'F54': (f83, 54, 6),
    'C55': (f84, 55, 3),
    'D55': (f85, 55, 4),
    'E55': (f86, 55, 5),
    'F55': (f87, 55, 6),
    'C56': (f88, 56, 3),
    'D56': (f88, 56, 4),
    'E56': (f89, 56, 5),
    'F56': (f88, 56, 6),
    'C57': (f90, 57, 3),
    'D57': (f90, 57, 4),
    'E57': (f48, 57, 5),
    'F57': (f90, 57, 6),
    'C58': (f91, 58, 3),
    'D58': (f92, 58, 4),
    'E58': (f48, 58, 5),
    'F58': (f93, 58, 6),
    'C62': (f94, 62, 3),
    'D62': (f95, 62, 4),
    'C63': (f96, 63, 3),
    'D63': (f97, 63, 4),
    'C64': (f98, 64, 3),
    'D64': (f99, 64, 4),
    'C65': (f100, 65, 3),
    'D65': (f101, 65, 4),
    'C66': (f102, 66, 3),
    'D66': (f102, 66, 4),
    'C67': (f103, 67, 3),
    'C68': (f104, 68, 3),
    'C71': (f105, 71, 3),
    'C72': (f106, 72, 3),
    'C74': (f107, 74, 3),
    'D74': (f107, 74, 4),
    'E74': (f108, 74, 5),
    'F74': (f107, 74, 6),
    'C75': (f109, 75, 3),
    'D75': (f109, 75, 4),
    'E75': (f108, 75, 5),
    'F75': (f109, 75, 6),
    'C76': (f110, 76, 3),
    'D76': (f111, 76, 4),
    'E76': (f108, 76, 5),
    'F76': (f112, 76, 6),
    'C77': (f113, 77, 3),
    'D77': (f113, 77, 4),
    'E77': (f113, 77, 5),
    'F77': (f113, 77, 6),
    'C78': (f114, 78, 3),
    'D78': (f114, 78, 4),
    'E78': (f114, 78, 5),
    'F78': (f114, 78, 6),
    'C79': (f115, 79, 3),
    'D79': (f115, 79, 4),
    'E79': (f115, 79, 5),
    'F79': (f115, 79, 6),
    'C80': (f116, 80, 3),
    'D80': (f117, 80, 4),
    'E80': (f118, 80, 5),
    'F80': (f119, 80, 6),
    'C82': (f120, 82, 3),
    'C83': (f121, 83, 3),
    'C84': (f122, 84, 3),
    'C85': (f123, 85, 3),
    'C86': (f124, 86, 3),
    'C87': (f125, 87, 3),
    'C88': (f126, 88, 3),
    'C89': (f127, 89, 3),
    'D89': (f128, 89, 4),
    'C91': (f129, 91, 3),
    'C92': (f130, 92, 3),
    'D96': (f131, 96, 4),
    'E96': (f132, 96, 5),
    'F96': (f133, 96, 6),
    'D97': (f134, 97, 4),
    'E97': (f135, 97, 5),
    'F97': (f133, 97, 6),
    'D98': (f136, 98, 4),
    'E98': (f137, 98, 5),
    'F98': (f133, 98, 6),
    'D99': (f138, 99, 4),
    'E99': (f139, 99, 5),
    'F99': (f133, 99, 6),
    'D100': (f140, 100, 4),
    'E100': (f141, 100, 5),
    'F100': (f133, 100, 6),
    'C104': (f142, 104, 3),
    'C105': (f143, 105, 3),
    'C106': (f144, 106, 3),
    'J106': (f145, 106, 10),
    'C107': (f146, 107, 3),
    'J107': (f147, 107, 10),
    'C108': (f148, 108, 3),
    'J108': (f149, 108, 10),
    'C109': (f150, 109, 3),
    'J109': (f151, 109, 10),
    'C110': (f152, 110, 3),
    'C111': (f153, 111, 3),
    'C112': (f154, 112, 3),
    'C115': (f155, 115, 3),
    'C116': (f156, 116, 3),
    'C117': (f157, 117, 3),
    'C118': (f158, 118, 3),
    'C119': (f159, 119, 3),
    'C122': (f160, 122, 3),
    'D122': (f160, 122, 4),
    'E122': (f161, 122, 5),
    'C123': (f162, 123, 3),
    'D123': (f162, 123, 4),
    'E123': (f162, 123, 5),
    'C124': (f163, 124, 3),
    'D124': (f163, 124, 4),
    'E124': (f163, 124, 5),
    'C125': (f164, 125, 3),
}
