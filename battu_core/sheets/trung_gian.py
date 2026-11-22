# -*- coding: utf-8 -*-
"""Module sheet TRUNG_GIAN – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: B69n.xlsx (SHA-256 9228d1ab613379f50ca7126e315efcbb940599c6e1784212c489756674ca8b67).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'TRUNG_GIAN'
SI = 23

def f0(E, r, c):
    # Ô C5
    # = NHAP_LIEU!$C$20
    return E.v(2,20,3)

def f1(E, r, c):
    # Ô D5
    # = NHAP_LIEU!$D$20
    return E.v(2,20,4)

def f2(E, r, c):
    # Ô E5
    # = NHAP_LIEU!$E$20
    return E.v(2,20,5)

def f3(E, r, c):
    # Ô F5
    # = NHAP_LIEU!$F$20
    return E.v(2,20,6)

def f4(E, r, c):
    # Ô C6
    # = NHAP_LIEU!$C$21
    return E.v(2,21,3)

def f5(E, r, c):
    # Ô D6
    # = NHAP_LIEU!$D$21
    return E.v(2,21,4)

def f6(E, r, c):
    # Ô E6
    # = NHAP_LIEU!$E$21
    return E.v(2,21,5)

def f7(E, r, c):
    # Ô F6
    # = NHAP_LIEU!$F$21
    return E.v(2,21,6)

def f8(E, r, c):
    # Ô C7, D7, E7, F7
    # = INDEX(CSDL!$C$5:$C$14,C5+1)
    return X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,r-2,c),1.0))

def f9(E, r, c):
    # Ô C8, D8, E8, F8
    # = INDEX(CSDL!$D$5:$D$14,C5+1)
    return X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,r-3,c),1.0))

def f10(E, r, c):
    # Ô C9, D9, E9, F9
    # = INDEX(CSDL!$C$18:$C$29,C6+1)
    return X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,r-3,c),1.0))

def f11(E, r, c):
    # Ô C10, D10, E10, F10
    # = INDEX(CAN_TEN,C5+1)
    return X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,r-5,c),1.0))

def f12(E, r, c):
    # Ô C11, D11, E11, F11
    # = INDEX(CHI_TEN,C6+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,r-5,c),1.0))

def f13(E, r, c):
    # Ô C12, D12, F12
    # = INDEX(M_TT,$E$5+1,C5+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,r-7,c),1.0))

def f14(E, r, c):
    # Ô E12
    # = "Nhật chủ"
    return 'Nhật chủ'

def f15(E, r, c):
    # Ô C13, D13, E13, F13
    # = MOD(C7-$E$7,5)
    return X.MOD(X.sub(E.v(23,r-6,c),E.v(23,7,5)), 5.0)

def f16(E, r, c):
    # Ô C15, D15, E15, F15
    # = INDEX(CSDL!$D$18:$D$29,C6+1)
    return X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,r-9,c),1.0))

def f17(E, r, c):
    # Ô C16, D16, E16, F16
    # = INDEX(CSDL!$E$18:$E$29,C6+1)
    return X.INDEX(E.rg(24,18,5,29,5), X.add(E.v(23,r-10,c),1.0))

def f18(E, r, c):
    # Ô C17, D17, E17, F17
    # = INDEX(CSDL!$F$18:$F$29,C6+1)
    return X.INDEX(E.rg(24,18,6,29,6), X.add(E.v(23,r-11,c),1.0))

def f19(E, r, c):
    # Ô C18, D18, E18, F18
    # = INDEX(CSDL!$G$18:$G$29,C6+1)
    return X.INDEX(E.rg(24,18,7,29,7), X.add(E.v(23,r-12,c),1.0))

def f20(E, r, c):
    # Ô C19, D19, E19, F19
    # = INDEX(CSDL!$H$18:$H$29,C6+1)
    return X.INDEX(E.rg(24,18,8,29,8), X.add(E.v(23,r-13,c),1.0))

def f21(E, r, c):
    # Ô C20, D20, E20, F20
    # = INDEX(CSDL!$I$18:$I$29,C6+1)
    return X.INDEX(E.rg(24,18,9,29,9), X.add(E.v(23,r-14,c),1.0))

def f22(E, r, c):
    # Ô C21, D21, E21, F21
    # = IF(C15<0,-1,INDEX(CSDL!$C$5:$C$14,C15+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,r-6,c),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,r-6,c),1.0)))

def f23(E, r, c):
    # Ô C22, D22, E22, F22
    # = IF(C17<0,-1,INDEX(CSDL!$C$5:$C$14,C17+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,r-5,c),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,r-5,c),1.0)))

def f24(E, r, c):
    # Ô C23, D23, E23, F23
    # = IF(C19<0,-1,INDEX(CSDL!$C$5:$C$14,C19+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,r-4,c),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,r-4,c),1.0)))

def f25(E, r, c):
    # Ô C24, D24, E24, F24
    # = IF(C15<0,"",INDEX(CAN_TEN,C15+1)&" – "&INDEX(M_TT,$E$5+1,C15+1))
    return ('' if X.truthy(X.lt(E.v(23,r-9,c),0.0)) else X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,r-9,c),1.0)),' – '),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,r-9,c),1.0))))

def f26(E, r, c):
    # Ô C25, D25, E25, F25
    # = IF(C17<0,"",INDEX(CAN_TEN,C17+1)&" – "&INDEX(M_TT,$E$5+1,C17+1))
    return ('' if X.truthy(X.lt(E.v(23,r-8,c),0.0)) else X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,r-8,c),1.0)),' – '),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,r-8,c),1.0))))

def f27(E, r, c):
    # Ô C26, D26, E26, F26
    # = IF(C19<0,"",INDEX(CAN_TEN,C19+1)&" – "&INDEX(M_TT,$E$5+1,C19+1))
    return ('' if X.truthy(X.lt(E.v(23,r-7,c),0.0)) else X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,r-7,c),1.0)),' – '),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,r-7,c),1.0))))

def f28(E, r, c):
    # Ô C28, D28, E28, F28
    # = INDEX(M_TS,$E$5+1,C6+1)
    return X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,r-22,c),1.0))

def f29(E, r, c):
    # Ô C29, D29, E29, F29, C31, D31, E31, F31
    # = INDEX(STAGE_TEN,C28+1)
    return X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(23,r-1,c),1.0))

def f30(E, r, c):
    # Ô C30, D30, E30, F30
    # = INDEX(M_TS,C5+1,$D$6+1)
    return X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,r-25,c),1.0), X.add(E.v(23,6,4),1.0))

def f31(E, r, c):
    # Ô B35
    # = CSDL!$D$62&" "&INDEX(CHI_TEN,CSDL!$E$62+1)&"-"&INDEX(CHI_TEN,CSDL!$F$62+1)&"-"&INDEX(CHI_TEN,CSDL!$G$62+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,62,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,62,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,62,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,62,7),1.0)))

def f32(E, r, c):
    # Ô C35
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$62)>0,COUNTIF($C$6:$F$6,CSDL!$F$62)>0,COUNTIF($C$6:$F$6,CSDL!$G$62)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,62,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,62,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,62,7)),0.0))) else 0.0)

def f33(E, r, c):
    # Ô D35
    # = IFERROR(MATCH(CSDL!$E$62,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,62,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f34(E, r, c):
    # Ô E35
    # = IFERROR(MATCH(CSDL!$F$62,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,62,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f35(E, r, c):
    # Ô F35
    # = IFERROR(MATCH(CSDL!$G$62,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,62,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f36(E, r, c):
    # Ô G35
    # = CSDL!$H$62
    return E.v(24,62,8)

def f37(E, r, c):
    # Ô B36
    # = CSDL!$D$63&" "&INDEX(CHI_TEN,CSDL!$E$63+1)&"-"&INDEX(CHI_TEN,CSDL!$F$63+1)&"-"&INDEX(CHI_TEN,CSDL!$G$63+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,63,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,63,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,63,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,63,7),1.0)))

def f38(E, r, c):
    # Ô C36
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$63)>0,COUNTIF($C$6:$F$6,CSDL!$F$63)>0,COUNTIF($C$6:$F$6,CSDL!$G$63)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,63,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,63,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,63,7)),0.0))) else 0.0)

def f39(E, r, c):
    # Ô D36
    # = IFERROR(MATCH(CSDL!$E$63,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,63,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f40(E, r, c):
    # Ô E36
    # = IFERROR(MATCH(CSDL!$F$63,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,63,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f41(E, r, c):
    # Ô F36
    # = IFERROR(MATCH(CSDL!$G$63,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,63,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f42(E, r, c):
    # Ô G36
    # = CSDL!$H$63
    return E.v(24,63,8)

def f43(E, r, c):
    # Ô B37
    # = CSDL!$D$64&" "&INDEX(CHI_TEN,CSDL!$E$64+1)&"-"&INDEX(CHI_TEN,CSDL!$F$64+1)&"-"&INDEX(CHI_TEN,CSDL!$G$64+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,64,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,64,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,64,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,64,7),1.0)))

def f44(E, r, c):
    # Ô C37
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$64)>0,COUNTIF($C$6:$F$6,CSDL!$F$64)>0,COUNTIF($C$6:$F$6,CSDL!$G$64)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,64,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,64,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,64,7)),0.0))) else 0.0)

def f45(E, r, c):
    # Ô D37
    # = IFERROR(MATCH(CSDL!$E$64,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,64,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f46(E, r, c):
    # Ô E37
    # = IFERROR(MATCH(CSDL!$F$64,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,64,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f47(E, r, c):
    # Ô F37
    # = IFERROR(MATCH(CSDL!$G$64,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,64,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f48(E, r, c):
    # Ô G37
    # = CSDL!$H$64
    return E.v(24,64,8)

def f49(E, r, c):
    # Ô B38
    # = CSDL!$D$65&" "&INDEX(CHI_TEN,CSDL!$E$65+1)&"-"&INDEX(CHI_TEN,CSDL!$F$65+1)&"-"&INDEX(CHI_TEN,CSDL!$G$65+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,65,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,65,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,65,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,65,7),1.0)))

def f50(E, r, c):
    # Ô C38
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$65)>0,COUNTIF($C$6:$F$6,CSDL!$F$65)>0,COUNTIF($C$6:$F$6,CSDL!$G$65)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,65,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,65,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,65,7)),0.0))) else 0.0)

def f51(E, r, c):
    # Ô D38
    # = IFERROR(MATCH(CSDL!$E$65,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,65,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f52(E, r, c):
    # Ô E38
    # = IFERROR(MATCH(CSDL!$F$65,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,65,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f53(E, r, c):
    # Ô F38
    # = IFERROR(MATCH(CSDL!$G$65,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,65,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f54(E, r, c):
    # Ô G38
    # = CSDL!$H$65
    return E.v(24,65,8)

def f55(E, r, c):
    # Ô B39
    # = CSDL!$D$66&" "&INDEX(CHI_TEN,CSDL!$E$66+1)&"-"&INDEX(CHI_TEN,CSDL!$F$66+1)&"-"&INDEX(CHI_TEN,CSDL!$G$66+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,66,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,66,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,66,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,66,7),1.0)))

def f56(E, r, c):
    # Ô C39
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$66)>0,COUNTIF($C$6:$F$6,CSDL!$F$66)>0,COUNTIF($C$6:$F$6,CSDL!$G$66)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,66,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,66,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,66,7)),0.0))) else 0.0)

def f57(E, r, c):
    # Ô D39
    # = IFERROR(MATCH(CSDL!$E$66,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,66,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f58(E, r, c):
    # Ô E39
    # = IFERROR(MATCH(CSDL!$F$66,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,66,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f59(E, r, c):
    # Ô F39
    # = IFERROR(MATCH(CSDL!$G$66,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,66,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f60(E, r, c):
    # Ô G39
    # = CSDL!$H$66
    return E.v(24,66,8)

def f61(E, r, c):
    # Ô B40
    # = CSDL!$D$67&" "&INDEX(CHI_TEN,CSDL!$E$67+1)&"-"&INDEX(CHI_TEN,CSDL!$F$67+1)&"-"&INDEX(CHI_TEN,CSDL!$G$67+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,67,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,67,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,67,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,67,7),1.0)))

def f62(E, r, c):
    # Ô C40
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$67)>0,COUNTIF($C$6:$F$6,CSDL!$F$67)>0,COUNTIF($C$6:$F$6,CSDL!$G$67)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,67,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,67,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,67,7)),0.0))) else 0.0)

def f63(E, r, c):
    # Ô D40
    # = IFERROR(MATCH(CSDL!$E$67,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,67,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f64(E, r, c):
    # Ô E40
    # = IFERROR(MATCH(CSDL!$F$67,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,67,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f65(E, r, c):
    # Ô F40
    # = IFERROR(MATCH(CSDL!$G$67,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,67,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f66(E, r, c):
    # Ô G40
    # = CSDL!$H$67
    return E.v(24,67,8)

def f67(E, r, c):
    # Ô B41
    # = CSDL!$D$68&" "&INDEX(CHI_TEN,CSDL!$E$68+1)&"-"&INDEX(CHI_TEN,CSDL!$F$68+1)&"-"&INDEX(CHI_TEN,CSDL!$G$68+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,68,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,68,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,68,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,68,7),1.0)))

def f68(E, r, c):
    # Ô C41
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$68)>0,COUNTIF($C$6:$F$6,CSDL!$F$68)>0,COUNTIF($C$6:$F$6,CSDL!$G$68)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,68,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,68,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,68,7)),0.0))) else 0.0)

def f69(E, r, c):
    # Ô D41
    # = IFERROR(MATCH(CSDL!$E$68,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,68,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f70(E, r, c):
    # Ô E41
    # = IFERROR(MATCH(CSDL!$F$68,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,68,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f71(E, r, c):
    # Ô F41
    # = IFERROR(MATCH(CSDL!$G$68,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,68,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f72(E, r, c):
    # Ô G41
    # = CSDL!$H$68
    return E.v(24,68,8)

def f73(E, r, c):
    # Ô B42
    # = CSDL!$D$69&" "&INDEX(CHI_TEN,CSDL!$E$69+1)&"-"&INDEX(CHI_TEN,CSDL!$F$69+1)&"-"&INDEX(CHI_TEN,CSDL!$G$69+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(24,69,4),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,69,5),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,69,6),1.0))),'-'),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(24,69,7),1.0)))

def f74(E, r, c):
    # Ô C42
    # = IF(AND(COUNTIF($C$6:$F$6,CSDL!$E$69)>0,COUNTIF($C$6:$F$6,CSDL!$F$69)>0,COUNTIF($C$6:$F$6,CSDL!$G$69)>0),1,0)
    return (1.0 if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,69,5)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,69,6)),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(24,69,7)),0.0))) else 0.0)

def f75(E, r, c):
    # Ô D42
    # = IFERROR(MATCH(CSDL!$E$69,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,69,5), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f76(E, r, c):
    # Ô E42
    # = IFERROR(MATCH(CSDL!$F$69,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,69,6), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f77(E, r, c):
    # Ô F42
    # = IFERROR(MATCH(CSDL!$G$69,$C$6:$F$6,0)-1,-1)
    return X.IFERROR(lambda: X.sub(X.MATCH(E.v(24,69,7), E.rg(23,6,3,6,6), 0.0),1.0), lambda: X.neg(1.0))

def f78(E, r, c):
    # Ô G42
    # = CSDL!$H$69
    return E.v(24,69,8)

def f79(E, r, c):
    # Ô C44
    # = IFERROR(MATCH(1,$C$35:$C$42,0),0)
    return X.IFERROR(lambda: X.MATCH(1.0, E.rg(23,35,3,42,3), 0.0), lambda: 0.0)

def f80(E, r, c):
    # Ô C45
    # = IF($C$44=0,"",INDEX(CSDL!$D$62:$D$69,$C$44))
    return ('' if X.truthy(X.eq(E.v(23,44,3),0.0)) else X.INDEX(E.rg(24,62,4,69,4), E.v(23,44,3)))

def f81(E, r, c):
    # Ô C46
    # = IF($C$44=0,-1,INDEX(CSDL!$H$62:$H$69,$C$44))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(23,44,3),0.0)) else X.INDEX(E.rg(24,62,8,69,8), E.v(23,44,3)))

def f82(E, r, c):
    # Ô C47
    # = IF($C$44=0,-1,INDEX($D$35:$D$42,$C$44))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(23,44,3),0.0)) else X.INDEX(E.rg(23,35,4,42,4), E.v(23,44,3)))

def f83(E, r, c):
    # Ô C48
    # = IF($C$44=0,-1,INDEX($E$35:$E$42,$C$44))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(23,44,3),0.0)) else X.INDEX(E.rg(23,35,5,42,5), E.v(23,44,3)))

def f84(E, r, c):
    # Ô C49
    # = IF($C$44=0,-1,INDEX($F$35:$F$42,$C$44))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(23,44,3),0.0)) else X.INDEX(E.rg(23,35,6,42,6), E.v(23,44,3)))

def f85(E, r, c):
    # Ô C50
    # = IF($C$44=0,-1,INDEX(CSDL!$I$62:$I$69,$C$44))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(23,44,3),0.0)) else X.INDEX(E.rg(24,62,9,69,9), E.v(23,44,3)))

def f86(E, r, c):
    # Ô C52, D52, E52, F52
    # = IF(OR(C$4=$C$47,C$4=$C$48,C$4=$C$49),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,4,c),E.v(23,47,3)), X.eq(E.v(23,4,c),E.v(23,48,3)), X.eq(E.v(23,4,c),E.v(23,49,3)))) else 0.0)

def f87(E, r, c):
    # Ô C55, D55, E55
    # = INDEX(M_BH,C$6+1,D$6+1)
    return X.INDEX(E.rg(24,62,16,73,27), X.add(E.v(23,6,c),1.0), X.add(E.v(23,6,c+1),1.0))

def f88(E, r, c):
    # Ô C56, D56, E56
    # = IF(AND(C$52=0,D$52=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,52,c),0.0), X.eq(E.v(23,52,c+1),0.0))) else 0.0)

def f89(E, r, c):
    # Ô C57
    # = IF(AND(C55>=0,C56=1),1,0)
    return (1.0 if X.truthy(X.AND(X.ge(E.v(23,r-2,c),0.0), X.eq(E.v(23,r-1,c),1.0))) else 0.0)

def f90(E, r, c):
    # Ô D57, E57
    # = IF(AND(D55>=0,D56=1,C57=0),1,0)
    return (1.0 if X.truthy(X.AND(X.ge(E.v(23,r-2,c),0.0), X.eq(E.v(23,r-1,c),1.0), X.eq(E.v(23,r,c-1),0.0))) else 0.0)

def f91(E, r, c):
    # Ô C59, C67
    # = C57
    return E.v(23,r-2,c)

def f92(E, r, c):
    # Ô D59, E59, D67, E67
    # = C57+D57
    return X.add(E.v(23,r-2,c-1),E.v(23,r-2,c))

def f93(E, r, c):
    # Ô F59, F67
    # = E57
    return E.v(23,r-2,c-1)

def f94(E, r, c):
    # Ô C60, D60, E60, F60
    # = MAX(C52,C59)
    return X.MAX(E.rg(23,r-8,c,r-8,c), E.rg(23,r-1,c,r-1,c))

def f95(E, r, c):
    # Ô C61
    # = IF(C52=1,$C$46,IF(FALSE(),0,IF(C57=1,C55,-1)))
    return (E.v(23,46,3) if X.truthy(X.eq(E.v(23,r-9,c),1.0)) else (0.0 if X.truthy(X.FALSE()) else (E.v(23,r-6,c) if X.truthy(X.eq(E.v(23,r-4,c),1.0)) else X.neg(1.0))))

def f96(E, r, c):
    # Ô D61, E61
    # = IF(D52=1,$C$46,IF(C57=1,C55,IF(D57=1,D55,-1)))
    return (E.v(23,46,3) if X.truthy(X.eq(E.v(23,r-9,c),1.0)) else (E.v(23,r-6,c-1) if X.truthy(X.eq(E.v(23,r-4,c-1),1.0)) else (E.v(23,r-6,c) if X.truthy(X.eq(E.v(23,r-4,c),1.0)) else X.neg(1.0))))

def f97(E, r, c):
    # Ô F61
    # = IF(F52=1,$C$46,IF(E57=1,E55,-1))
    return (E.v(23,46,3) if X.truthy(X.eq(E.v(23,r-9,c),1.0)) else (E.v(23,r-6,c-1) if X.truthy(X.eq(E.v(23,r-4,c-1),1.0)) else X.neg(1.0)))

def f98(E, r, c):
    # Ô C62, D62, E62, F62
    # = IF(C52=1,$C$45,IF(C59=1,"Bán hợp",""))
    return (E.v(23,45,3) if X.truthy(X.eq(E.v(23,r-10,c),1.0)) else ('Bán hợp' if X.truthy(X.eq(E.v(23,r-3,c),1.0)) else ''))

def f99(E, r, c):
    # Ô C64, D64, E64
    # = INDEX(M_LH,C$6+1,D$6+1)
    return X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,c),1.0), X.add(E.v(23,6,c+1),1.0))

def f100(E, r, c):
    # Ô C65, D65, E65
    # = IF(AND(C64>=0,C$60=0,D$60=0),1,0)
    return (1.0 if X.truthy(X.AND(X.ge(E.v(23,r-1,c),0.0), X.eq(E.v(23,60,c),0.0), X.eq(E.v(23,60,c+1),0.0))) else 0.0)

def f101(E, r, c):
    # Ô C68, D68, E68
    # = IF(AND(C65=1,OR(C67>1,D67>1)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,r-3,c),1.0), X.OR(X.gt(E.v(23,r-1,c),1.0), X.gt(E.v(23,r-1,c+1),1.0)))) else 0.0)

def f102(E, r, c):
    # Ô C69, D69, E69, F69
    # = IF(C67>0,1,0)
    return (1.0 if X.truthy(X.gt(E.v(23,r-2,c),0.0)) else 0.0)

def f103(E, r, c):
    # Ô C70
    # = IF(FALSE(),0,IF(C65=1,C64,-1))
    return (0.0 if X.truthy(X.FALSE()) else (E.v(23,r-6,c) if X.truthy(X.eq(E.v(23,r-5,c),1.0)) else X.neg(1.0)))

def f104(E, r, c):
    # Ô D70, E70
    # = IF(C65=1,C64,IF(D65=1,D64,-1))
    return (E.v(23,r-6,c-1) if X.truthy(X.eq(E.v(23,r-5,c-1),1.0)) else (E.v(23,r-6,c) if X.truthy(X.eq(E.v(23,r-5,c),1.0)) else X.neg(1.0)))

def f105(E, r, c):
    # Ô F70
    # = IF(E65=1,E64,-1)
    return (E.v(23,r-6,c-1) if X.truthy(X.eq(E.v(23,r-5,c-1),1.0)) else X.neg(1.0))

def f106(E, r, c):
    # Ô C71
    # = MAX(0,C68)
    return X.MAX(0.0, E.rg(23,r-3,c,r-3,c))

def f107(E, r, c):
    # Ô D71, E71
    # = MAX(0,C68,D68)
    return X.MAX(0.0, E.rg(23,r-3,c-1,r-3,c-1), E.rg(23,r-3,c,r-3,c))

def f108(E, r, c):
    # Ô F71
    # = MAX(0,E68)
    return X.MAX(0.0, E.rg(23,r-3,c-1,r-3,c-1))

def f109(E, r, c):
    # Ô C73, D73, E73
    # = INDEX(M_XU,C$6+1,D$6+1)
    return X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,c),1.0), X.add(E.v(23,6,c+1),1.0))

def f110(E, r, c):
    # Ô C74, D74, E74
    # = IF(AND(C73=1,OR(C$6=4,C$6=10,C$6=1,C$6=7)),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(23,r-1,c),1.0), X.OR(X.eq(E.v(23,6,c),4.0), X.eq(E.v(23,6,c),10.0), X.eq(E.v(23,6,c),1.0), X.eq(E.v(23,6,c),7.0)))) else 0.0)

def f111(E, r, c):
    # Ô C76, D76, E76
    # = INDEX(M_NH,C$5+1,D$5+1)
    return X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,c),1.0), X.add(E.v(23,5,c+1),1.0))

def f112(E, r, c):
    # Ô C77
    # = IF(C76>=0,1,0)
    return (1.0 if X.truthy(X.ge(E.v(23,r-1,c),0.0)) else 0.0)

def f113(E, r, c):
    # Ô D77, E77
    # = IF(C76>=0,1,0)+IF(D76>=0,1,0)
    return X.add((1.0 if X.truthy(X.ge(E.v(23,r-1,c-1),0.0)) else 0.0),(1.0 if X.truthy(X.ge(E.v(23,r-1,c),0.0)) else 0.0))

def f114(E, r, c):
    # Ô F77
    # = IF(E76>=0,1,0)
    return (1.0 if X.truthy(X.ge(E.v(23,r-1,c-1),0.0)) else 0.0)

def f115(E, r, c):
    # Ô C78, D78, E78
    # = IF(AND(C76>=0,OR(C77>1,D77>1)),1,0)
    return (1.0 if X.truthy(X.AND(X.ge(E.v(23,r-2,c),0.0), X.OR(X.gt(E.v(23,r-1,c),1.0), X.gt(E.v(23,r-1,c+1),1.0)))) else 0.0)

CELLS = {
    'C5': (f0, 5, 3),
    'D5': (f1, 5, 4),
    'E5': (f2, 5, 5),
    'F5': (f3, 5, 6),
    'C6': (f4, 6, 3),
    'D6': (f5, 6, 4),
    'E6': (f6, 6, 5),
    'F6': (f7, 6, 6),
    'C7': (f8, 7, 3),
    'D7': (f8, 7, 4),
    'E7': (f8, 7, 5),
    'F7': (f8, 7, 6),
    'C8': (f9, 8, 3),
    'D8': (f9, 8, 4),
    'E8': (f9, 8, 5),
    'F8': (f9, 8, 6),
    'C9': (f10, 9, 3),
    'D9': (f10, 9, 4),
    'E9': (f10, 9, 5),
    'F9': (f10, 9, 6),
    'C10': (f11, 10, 3),
    'D10': (f11, 10, 4),
    'E10': (f11, 10, 5),
    'F10': (f11, 10, 6),
    'C11': (f12, 11, 3),
    'D11': (f12, 11, 4),
    'E11': (f12, 11, 5),
    'F11': (f12, 11, 6),
    'C12': (f13, 12, 3),
    'D12': (f13, 12, 4),
    'E12': (f14, 12, 5),
    'F12': (f13, 12, 6),
    'C13': (f15, 13, 3),
    'D13': (f15, 13, 4),
    'E13': (f15, 13, 5),
    'F13': (f15, 13, 6),
    'C15': (f16, 15, 3),
    'D15': (f16, 15, 4),
    'E15': (f16, 15, 5),
    'F15': (f16, 15, 6),
    'C16': (f17, 16, 3),
    'D16': (f17, 16, 4),
    'E16': (f17, 16, 5),
    'F16': (f17, 16, 6),
    'C17': (f18, 17, 3),
    'D17': (f18, 17, 4),
    'E17': (f18, 17, 5),
    'F17': (f18, 17, 6),
    'C18': (f19, 18, 3),
    'D18': (f19, 18, 4),
    'E18': (f19, 18, 5),
    'F18': (f19, 18, 6),
    'C19': (f20, 19, 3),
    'D19': (f20, 19, 4),
    'E19': (f20, 19, 5),
    'F19': (f20, 19, 6),
    'C20': (f21, 20, 3),
    'D20': (f21, 20, 4),
    'E20': (f21, 20, 5),
    'F20': (f21, 20, 6),
    'C21': (f22, 21, 3),
    'D21': (f22, 21, 4),
    'E21': (f22, 21, 5),
    'F21': (f22, 21, 6),
    'C22': (f23, 22, 3),
    'D22': (f23, 22, 4),
    'E22': (f23, 22, 5),
    'F22': (f23, 22, 6),
    'C23': (f24, 23, 3),
    'D23': (f24, 23, 4),
    'E23': (f24, 23, 5),
    'F23': (f24, 23, 6),
    'C24': (f25, 24, 3),
    'D24': (f25, 24, 4),
    'E24': (f25, 24, 5),
    'F24': (f25, 24, 6),
    'C25': (f26, 25, 3),
    'D25': (f26, 25, 4),
    'E25': (f26, 25, 5),
    'F25': (f26, 25, 6),
    'C26': (f27, 26, 3),
    'D26': (f27, 26, 4),
    'E26': (f27, 26, 5),
    'F26': (f27, 26, 6),
    'C28': (f28, 28, 3),
    'D28': (f28, 28, 4),
    'E28': (f28, 28, 5),
    'F28': (f28, 28, 6),
    'C29': (f29, 29, 3),
    'D29': (f29, 29, 4),
    'E29': (f29, 29, 5),
    'F29': (f29, 29, 6),
    'C30': (f30, 30, 3),
    'D30': (f30, 30, 4),
    'E30': (f30, 30, 5),
    'F30': (f30, 30, 6),
    'C31': (f29, 31, 3),
    'D31': (f29, 31, 4),
    'E31': (f29, 31, 5),
    'F31': (f29, 31, 6),
    'B35': (f31, 35, 2),
    'C35': (f32, 35, 3),
    'D35': (f33, 35, 4),
    'E35': (f34, 35, 5),
    'F35': (f35, 35, 6),
    'G35': (f36, 35, 7),
    'B36': (f37, 36, 2),
    'C36': (f38, 36, 3),
    'D36': (f39, 36, 4),
    'E36': (f40, 36, 5),
    'F36': (f41, 36, 6),
    'G36': (f42, 36, 7),
    'B37': (f43, 37, 2),
    'C37': (f44, 37, 3),
    'D37': (f45, 37, 4),
    'E37': (f46, 37, 5),
    'F37': (f47, 37, 6),
    'G37': (f48, 37, 7),
    'B38': (f49, 38, 2),
    'C38': (f50, 38, 3),
    'D38': (f51, 38, 4),
    'E38': (f52, 38, 5),
    'F38': (f53, 38, 6),
    'G38': (f54, 38, 7),
    'B39': (f55, 39, 2),
    'C39': (f56, 39, 3),
    'D39': (f57, 39, 4),
    'E39': (f58, 39, 5),
    'F39': (f59, 39, 6),
    'G39': (f60, 39, 7),
    'B40': (f61, 40, 2),
    'C40': (f62, 40, 3),
    'D40': (f63, 40, 4),
    'E40': (f64, 40, 5),
    'F40': (f65, 40, 6),
    'G40': (f66, 40, 7),
    'B41': (f67, 41, 2),
    'C41': (f68, 41, 3),
    'D41': (f69, 41, 4),
    'E41': (f70, 41, 5),
    'F41': (f71, 41, 6),
    'G41': (f72, 41, 7),
    'B42': (f73, 42, 2),
    'C42': (f74, 42, 3),
    'D42': (f75, 42, 4),
    'E42': (f76, 42, 5),
    'F42': (f77, 42, 6),
    'G42': (f78, 42, 7),
    'C44': (f79, 44, 3),
    'C45': (f80, 45, 3),
    'C46': (f81, 46, 3),
    'C47': (f82, 47, 3),
    'C48': (f83, 48, 3),
    'C49': (f84, 49, 3),
    'C50': (f85, 50, 3),
    'C52': (f86, 52, 3),
    'D52': (f86, 52, 4),
    'E52': (f86, 52, 5),
    'F52': (f86, 52, 6),
    'C55': (f87, 55, 3),
    'D55': (f87, 55, 4),
    'E55': (f87, 55, 5),
    'C56': (f88, 56, 3),
    'D56': (f88, 56, 4),
    'E56': (f88, 56, 5),
    'C57': (f89, 57, 3),
    'D57': (f90, 57, 4),
    'E57': (f90, 57, 5),
    'C59': (f91, 59, 3),
    'D59': (f92, 59, 4),
    'E59': (f92, 59, 5),
    'F59': (f93, 59, 6),
    'C60': (f94, 60, 3),
    'D60': (f94, 60, 4),
    'E60': (f94, 60, 5),
    'F60': (f94, 60, 6),
    'C61': (f95, 61, 3),
    'D61': (f96, 61, 4),
    'E61': (f96, 61, 5),
    'F61': (f97, 61, 6),
    'C62': (f98, 62, 3),
    'D62': (f98, 62, 4),
    'E62': (f98, 62, 5),
    'F62': (f98, 62, 6),
    'C64': (f99, 64, 3),
    'D64': (f99, 64, 4),
    'E64': (f99, 64, 5),
    'C65': (f100, 65, 3),
    'D65': (f100, 65, 4),
    'E65': (f100, 65, 5),
    'C67': (f91, 67, 3),
    'D67': (f92, 67, 4),
    'E67': (f92, 67, 5),
    'F67': (f93, 67, 6),
    'C68': (f101, 68, 3),
    'D68': (f101, 68, 4),
    'E68': (f101, 68, 5),
    'C69': (f102, 69, 3),
    'D69': (f102, 69, 4),
    'E69': (f102, 69, 5),
    'F69': (f102, 69, 6),
    'C70': (f103, 70, 3),
    'D70': (f104, 70, 4),
    'E70': (f104, 70, 5),
    'F70': (f105, 70, 6),
    'C71': (f106, 71, 3),
    'D71': (f107, 71, 4),
    'E71': (f107, 71, 5),
    'F71': (f108, 71, 6),
    'C73': (f109, 73, 3),
    'D73': (f109, 73, 4),
    'E73': (f109, 73, 5),
    'C74': (f110, 74, 3),
    'D74': (f110, 74, 4),
    'E74': (f110, 74, 5),
    'C76': (f111, 76, 3),
    'D76': (f111, 76, 4),
    'E76': (f111, 76, 5),
    'C77': (f112, 77, 3),
    'D77': (f113, 77, 4),
    'E77': (f113, 77, 5),
    'F77': (f114, 77, 6),
    'C78': (f115, 78, 3),
    'D78': (f115, 78, 4),
    'E78': (f115, 78, 5),
}
