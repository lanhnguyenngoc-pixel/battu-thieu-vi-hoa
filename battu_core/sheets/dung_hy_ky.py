# -*- coding: utf-8 -*-
"""Module sheet DUNG_HY_KY – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.5.xlsx (SHA-256 f19f3f74b720c1a067a640a31d58441793bf64bfdd12629f17ba4be6623c56f4).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'DUNG_HY_KY'
SI = 8

def f0(E, r, c):
    # Ô C5
    # = TRUNG_GIAN!$E$7
    return E.v(23,7,5)

def f1(E, r, c):
    # Ô D5
    # = PP2_DINH_LUONG!$C$98
    return E.v(10,98,3)

def f2(E, r, c):
    # Ô C6
    # = PP1_DINH_TINH!$C$91
    return E.v(9,91,3)

def f3(E, r, c):
    # Ô D6
    # = PP2_DINH_LUONG!$C$103
    return E.v(10,103,3)

def f4(E, r, c):
    # Ô C7
    # = IF(C6="Trung hòa",2,IF(OR(C6="Thân nhược",C6="Quá nhược"),0,1))
    return (2.0 if X.truthy(X.eq(E.v(8,r-1,c),'Trung hòa')) else (0.0 if X.truthy(X.OR(X.eq(E.v(8,r-1,c),'Thân nhược'), X.eq(E.v(8,r-1,c),'Quá nhược'))) else 1.0))

def f5(E, r, c):
    # Ô D7
    # = IF(D6="Nhật chủ nhược",0,1)
    return (0.0 if X.truthy(X.eq(E.v(8,r-1,c),'Nhật chủ nhược')) else 1.0)

def f6(E, r, c):
    # Ô C8
    # = PP2_DINH_LUONG!$C$7
    return E.v(10,7,3)

def f7(E, r, c):
    # Ô D8
    # = PP2_DINH_LUONG!$C$25
    return E.v(10,25,3)

def f8(E, r, c):
    # Ô C9
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+4,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))+AND(TRUNG_GIAN!$C$21=MOD(TRUNG_GI
    #   AN!$E$7+4,5),INDEX(M_TS,TRUNG_GIAN!$C$15+1,TRUNG_GIAN!$D$6+1)<=4)*1+AND(TRUNG_GIAN!$D$21=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(M_TS,TRUNG_GIAN!$D$15+1,TRUNG_
    #   GIAN!$D$6+1)<=4)*1+AND(TRUNG_GIAN!$E$21=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(M_TS,TRUNG_GIAN!$E$15+1,TRUNG_GIAN!$D$6+1)<=4)*1+AND(TRUNG_GIAN!$F$21=MOD(TRUNG
    #   _GIAN!$E$7+4,5),INDEX(M_TS,TRUNG_GIAN!$F$15+1,TRUNG_GIAN!$D$6+1)<=4)*1>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),4.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),X.mul(X.AND(X.eq(E.v(23,21,3),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)),1.0)),X.mul(X.AND(X.eq(E.v(23,21,4),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)),1.0)),X.mul(X.AND(X.eq(E.v(23,21,5),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)),1.0)),X.mul(X.AND(X.eq(E.v(23,21,6),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)),1.0)),0.0)) else 0.0)

def f9(E, r, c):
    # Ô D9, D10
    # = C9
    return E.v(8,r,c-1)

def f10(E, r, c):
    # Ô C10
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+4,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),4.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else 0.0)

def f11(E, r, c):
    # Ô C12
    # = PP1_DINH_TINH!$F$96
    return E.v(9,96,6)

def f12(E, r, c):
    # Ô D12
    # = PP2_DINH_LUONG!$D$106
    return E.v(10,106,4)

def f13(E, r, c):
    # Ô C13
    # = PP1_DINH_TINH!$F$97
    return E.v(9,97,6)

def f14(E, r, c):
    # Ô D13
    # = PP2_DINH_LUONG!$D$107
    return E.v(10,107,4)

def f15(E, r, c):
    # Ô C14
    # = PP1_DINH_TINH!$F$98
    return E.v(9,98,6)

def f16(E, r, c):
    # Ô D14
    # = PP2_DINH_LUONG!$D$108
    return E.v(10,108,4)

def f17(E, r, c):
    # Ô C15
    # = PP1_DINH_TINH!$F$99
    return E.v(9,99,6)

def f18(E, r, c):
    # Ô D15
    # = PP2_DINH_LUONG!$D$109
    return E.v(10,109,4)

def f19(E, r, c):
    # Ô C16
    # = PP1_DINH_TINH!$F$100
    return E.v(9,100,6)

def f20(E, r, c):
    # Ô D16
    # = PP2_DINH_LUONG!$D$110
    return E.v(10,110,4)

def f21(E, r, c):
    # Ô C18, C19, C20, C21, C22
    # = IF(C12>0,1,0)
    return (1.0 if X.truthy(X.gt(E.v(8,r-6,c),0.0)) else 0.0)

def f22(E, r, c):
    # Ô D18, D19, D20, D21, D22
    # = IF(D12>0.000000001,1,0)
    return (1.0 if X.truthy(X.gt(E.v(8,r-6,c),1e-09)) else 0.0)

def f23(E, r, c):
    # Ô C24, D24
    # = INDEX(C$12:C$16,MOD(0-C$5,5)+1)
    return X.INDEX(E.rg(8,12,c,16,c), X.add(X.MOD(X.sub(0.0,E.v(8,5,c)), 5.0),1.0))

def f24(E, r, c):
    # Ô C25, D25
    # = INDEX(C$12:C$16,MOD(1-C$5,5)+1)
    return X.INDEX(E.rg(8,12,c,16,c), X.add(X.MOD(X.sub(1.0,E.v(8,5,c)), 5.0),1.0))

def f25(E, r, c):
    # Ô C26, D26
    # = INDEX(C$12:C$16,MOD(2-C$5,5)+1)
    return X.INDEX(E.rg(8,12,c,16,c), X.add(X.MOD(X.sub(2.0,E.v(8,5,c)), 5.0),1.0))

def f26(E, r, c):
    # Ô C27, D27
    # = INDEX(C$12:C$16,MOD(3-C$5,5)+1)
    return X.INDEX(E.rg(8,12,c,16,c), X.add(X.MOD(X.sub(3.0,E.v(8,5,c)), 5.0),1.0))

def f27(E, r, c):
    # Ô C28, D28
    # = INDEX(C$12:C$16,MOD(4-C$5,5)+1)
    return X.INDEX(E.rg(8,12,c,16,c), X.add(X.MOD(X.sub(4.0,E.v(8,5,c)), 5.0),1.0))

def f28(E, r, c):
    # Ô C29, D29
    # = MATCH(MAX(C24:C28),C24:C28,0)-1
    return X.sub(X.MATCH(X.MAX(E.rg(8,r-5,c,r-1,c)), E.rg(8,r-5,c,r-1,c), 0.0),1.0)

def f29(E, r, c):
    # Ô C30, D30
    # = INDEX(HANH_TEN,C29+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,r-1,c),1.0))

def f30(E, r, c):
    # Ô C31, D31
    # = IF(MOD(C$5+0,5)=C$8,0,100)+(1-IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+0,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,1,0))*10+MIN(IF(TRU
    #   NG_GIAN!$C$7=MOD(C$5+0,5),2,9),IF(TRUNG_GIAN!$D$7=MOD(C$5+0,5),1,9),IF(TRUNG_GIAN!$F$7=MOD(C$5+0,5),1,9),IF(TRUNG_GIAN!$C$9=MOD(C$5+0,5),2,9),IF(TRUNG
    #   _GIAN!$D$9=MOD(C$5+0,5),1,9),IF(TRUNG_GIAN!$E$9=MOD(C$5+0,5),0,9),IF(TRUNG_GIAN!$F$9=MOD(C$5+0,5),1,9))
    return X.add(X.add((0.0 if X.truthy(X.eq(X.MOD(X.add(E.v(8,5,c),0.0), 5.0),E.v(8,8,c))) else 100.0),X.mul(X.sub(1.0,(1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),0.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else 0.0)),10.0)),X.MIN((2.0 if X.truthy(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0), (2.0 if X.truthy(X.eq(E.v(23,9,3),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,4),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0), (0.0 if X.truthy(X.eq(E.v(23,9,5),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,6),X.MOD(X.add(E.v(8,5,c),0.0), 5.0))) else 9.0)))

def f31(E, r, c):
    # Ô C32, D32
    # = IF(MOD(C$5+1,5)=C$8,0,100)+(1-IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,1,0))*10+MIN(IF(TRU
    #   NG_GIAN!$C$7=MOD(C$5+1,5),2,9),IF(TRUNG_GIAN!$D$7=MOD(C$5+1,5),1,9),IF(TRUNG_GIAN!$F$7=MOD(C$5+1,5),1,9),IF(TRUNG_GIAN!$C$9=MOD(C$5+1,5),2,9),IF(TRUNG
    #   _GIAN!$D$9=MOD(C$5+1,5),1,9),IF(TRUNG_GIAN!$E$9=MOD(C$5+1,5),0,9),IF(TRUNG_GIAN!$F$9=MOD(C$5+1,5),1,9))
    return X.add(X.add((0.0 if X.truthy(X.eq(X.MOD(X.add(E.v(8,5,c),1.0), 5.0),E.v(8,8,c))) else 100.0),X.mul(X.sub(1.0,(1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else 0.0)),10.0)),X.MIN((2.0 if X.truthy(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0), (2.0 if X.truthy(X.eq(E.v(23,9,3),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,4),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0), (0.0 if X.truthy(X.eq(E.v(23,9,5),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,6),X.MOD(X.add(E.v(8,5,c),1.0), 5.0))) else 9.0)))

def f32(E, r, c):
    # Ô C33, D33
    # = IF(MOD(C$5+2,5)=C$8,0,100)+(1-IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+2,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,1,0))*10+MIN(IF(TRU
    #   NG_GIAN!$C$7=MOD(C$5+2,5),2,9),IF(TRUNG_GIAN!$D$7=MOD(C$5+2,5),1,9),IF(TRUNG_GIAN!$F$7=MOD(C$5+2,5),1,9),IF(TRUNG_GIAN!$C$9=MOD(C$5+2,5),2,9),IF(TRUNG
    #   _GIAN!$D$9=MOD(C$5+2,5),1,9),IF(TRUNG_GIAN!$E$9=MOD(C$5+2,5),0,9),IF(TRUNG_GIAN!$F$9=MOD(C$5+2,5),1,9))
    return X.add(X.add((0.0 if X.truthy(X.eq(X.MOD(X.add(E.v(8,5,c),2.0), 5.0),E.v(8,8,c))) else 100.0),X.mul(X.sub(1.0,(1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),2.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else 0.0)),10.0)),X.MIN((2.0 if X.truthy(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0), (2.0 if X.truthy(X.eq(E.v(23,9,3),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,4),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0), (0.0 if X.truthy(X.eq(E.v(23,9,5),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,6),X.MOD(X.add(E.v(8,5,c),2.0), 5.0))) else 9.0)))

def f33(E, r, c):
    # Ô C34, D34
    # = IF(MOD(C$5+3,5)=C$8,0,100)+(1-IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+3,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,1,0))*10+MIN(IF(TRU
    #   NG_GIAN!$C$7=MOD(C$5+3,5),2,9),IF(TRUNG_GIAN!$D$7=MOD(C$5+3,5),1,9),IF(TRUNG_GIAN!$F$7=MOD(C$5+3,5),1,9),IF(TRUNG_GIAN!$C$9=MOD(C$5+3,5),2,9),IF(TRUNG
    #   _GIAN!$D$9=MOD(C$5+3,5),1,9),IF(TRUNG_GIAN!$E$9=MOD(C$5+3,5),0,9),IF(TRUNG_GIAN!$F$9=MOD(C$5+3,5),1,9))
    return X.add(X.add((0.0 if X.truthy(X.eq(X.MOD(X.add(E.v(8,5,c),3.0), 5.0),E.v(8,8,c))) else 100.0),X.mul(X.sub(1.0,(1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else 0.0)),10.0)),X.MIN((2.0 if X.truthy(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0), (2.0 if X.truthy(X.eq(E.v(23,9,3),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,4),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0), (0.0 if X.truthy(X.eq(E.v(23,9,5),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,6),X.MOD(X.add(E.v(8,5,c),3.0), 5.0))) else 9.0)))

def f34(E, r, c):
    # Ô C35, D35
    # = IF(MOD(C$5+4,5)=C$8,0,100)+(1-IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+4,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,1,0))*10+MIN(IF(TRU
    #   NG_GIAN!$C$7=MOD(C$5+4,5),2,9),IF(TRUNG_GIAN!$D$7=MOD(C$5+4,5),1,9),IF(TRUNG_GIAN!$F$7=MOD(C$5+4,5),1,9),IF(TRUNG_GIAN!$C$9=MOD(C$5+4,5),2,9),IF(TRUNG
    #   _GIAN!$D$9=MOD(C$5+4,5),1,9),IF(TRUNG_GIAN!$E$9=MOD(C$5+4,5),0,9),IF(TRUNG_GIAN!$F$9=MOD(C$5+4,5),1,9))
    return X.add(X.add((0.0 if X.truthy(X.eq(X.MOD(X.add(E.v(8,5,c),4.0), 5.0),E.v(8,8,c))) else 100.0),X.mul(X.sub(1.0,(1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else 0.0)),10.0)),X.MIN((2.0 if X.truthy(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0), (2.0 if X.truthy(X.eq(E.v(23,9,3),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,4),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0), (0.0 if X.truthy(X.eq(E.v(23,9,5),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0), (1.0 if X.truthy(X.eq(E.v(23,9,6),X.MOD(X.add(E.v(8,5,c),4.0), 5.0))) else 9.0)))

def f35(E, r, c):
    # Ô C39
    # = PP2_DINH_LUONG!$C$40
    return E.v(10,40,3)

def f36(E, r, c):
    # Ô D39
    # = PP2_DINH_LUONG!$D$40
    return E.v(10,40,4)

def f37(E, r, c):
    # Ô E39
    # = 0
    return 0.0

def f38(E, r, c):
    # Ô F39
    # = PP2_DINH_LUONG!$F$40
    return E.v(10,40,6)

def f39(E, r, c):
    # Ô C40
    # = TRUNG_GIAN!$C$21
    return E.v(23,21,3)

def f40(E, r, c):
    # Ô D40
    # = TRUNG_GIAN!$D$21
    return E.v(23,21,4)

def f41(E, r, c):
    # Ô E40
    # = TRUNG_GIAN!$E$21
    return E.v(23,21,5)

def f42(E, r, c):
    # Ô F40
    # = TRUNG_GIAN!$F$21
    return E.v(23,21,6)

def f43(E, r, c):
    # Ô C41
    # = IF(OR(AND(TRUNG_GIAN!$C$52=1,PP1_DINH_TINH!$C$8=0,OR(PP1_DINH_TINH!$C$6=1,TRUNG_GIAN!$C$45="Tam hội")),PP1_DINH_TINH!$C$13=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,3),1.0), X.eq(E.v(9,8,3),0.0), X.OR(X.eq(E.v(9,6,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(9,13,3),1.0))) else 0.0)

def f44(E, r, c):
    # Ô D41
    # = IF(OR(AND(TRUNG_GIAN!$D$52=1,PP1_DINH_TINH!$C$8=0,OR(PP1_DINH_TINH!$C$6=1,TRUNG_GIAN!$C$45="Tam hội")),PP1_DINH_TINH!$C$13=1,PP1_DINH_TINH!$D$13=1),1,
    #   0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,4),1.0), X.eq(E.v(9,8,3),0.0), X.OR(X.eq(E.v(9,6,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(9,13,3),1.0), X.eq(E.v(9,13,4),1.0))) else 0.0)

def f45(E, r, c):
    # Ô E41
    # = IF(OR(AND(TRUNG_GIAN!$E$52=1,PP1_DINH_TINH!$C$8=0,OR(PP1_DINH_TINH!$C$6=1,TRUNG_GIAN!$C$45="Tam hội")),PP1_DINH_TINH!$D$13=1,PP1_DINH_TINH!$E$13=1),1,
    #   0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,5),1.0), X.eq(E.v(9,8,3),0.0), X.OR(X.eq(E.v(9,6,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(9,13,4),1.0), X.eq(E.v(9,13,5),1.0))) else 0.0)

def f46(E, r, c):
    # Ô F41
    # = IF(OR(AND(TRUNG_GIAN!$F$52=1,PP1_DINH_TINH!$C$8=0,OR(PP1_DINH_TINH!$C$6=1,TRUNG_GIAN!$C$45="Tam hội")),PP1_DINH_TINH!$E$13=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,6),1.0), X.eq(E.v(9,8,3),0.0), X.OR(X.eq(E.v(9,6,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(9,13,5),1.0))) else 0.0)

def f47(E, r, c):
    # Ô C42
    # = IF(C41=0,-1,PP1_DINH_TINH!$C$15)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(9,15,3))

def f48(E, r, c):
    # Ô D42
    # = IF(D41=0,-1,PP1_DINH_TINH!$D$15)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(9,15,4))

def f49(E, r, c):
    # Ô E42
    # = IF(E41=0,-1,PP1_DINH_TINH!$E$15)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(9,15,5))

def f50(E, r, c):
    # Ô F42
    # = IF(F41=0,-1,PP1_DINH_TINH!$F$15)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(9,15,6))

def f51(E, r, c):
    # Ô C43
    # = PP1_DINH_TINH!$C$27
    return E.v(9,27,3)

def f52(E, r, c):
    # Ô D43
    # = PP1_DINH_TINH!$D$27
    return E.v(9,27,4)

def f53(E, r, c):
    # Ô E43
    # = PP1_DINH_TINH!$E$27
    return E.v(9,27,5)

def f54(E, r, c):
    # Ô F43
    # = PP1_DINH_TINH!$F$27
    return E.v(9,27,6)

def f55(E, r, c):
    # Ô C44, C50
    # = IF(C43>=0,IF(C43=TRUNG_GIAN!$E$7,1,0),IF(OR(AND(OR(TRUNG_GIAN!$C$28=3,TRUNG_GIAN!$C$28=4),TRUNG_GIAN!$C$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(T
    #   RUNG_GIAN!$C$28=0,TRUNG_GIAN!$C$28=8))),1,0))
    return ((1.0 if X.truthy(X.eq(E.v(8,r-1,c),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(8,r-1,c),0.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,28,3),3.0), X.eq(E.v(23,28,3),4.0)), X.eq(E.v(23,21,3),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,3),0.0), X.eq(E.v(23,28,3),8.0))))) else 0.0))

def f56(E, r, c):
    # Ô D44, D50
    # = IF(D43>=0,IF(D43=TRUNG_GIAN!$E$7,1,0),IF(OR(AND(OR(TRUNG_GIAN!$D$28=3,TRUNG_GIAN!$D$28=4),TRUNG_GIAN!$D$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(T
    #   RUNG_GIAN!$D$28=0,TRUNG_GIAN!$D$28=8))),1,0))
    return ((1.0 if X.truthy(X.eq(E.v(8,r-1,c),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(8,r-1,c),0.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,28,4),3.0), X.eq(E.v(23,28,4),4.0)), X.eq(E.v(23,21,4),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,4),0.0), X.eq(E.v(23,28,4),8.0))))) else 0.0))

def f57(E, r, c):
    # Ô E44, E50
    # = IF(E43>=0,IF(E43=TRUNG_GIAN!$E$7,1,0),IF(OR(AND(OR(TRUNG_GIAN!$E$28=3,TRUNG_GIAN!$E$28=4),TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(T
    #   RUNG_GIAN!$E$28=0,TRUNG_GIAN!$E$28=8))),1,0))
    return ((1.0 if X.truthy(X.eq(E.v(8,r-1,c),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(8,r-1,c),0.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,28,5),3.0), X.eq(E.v(23,28,5),4.0)), X.eq(E.v(23,21,5),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,5),0.0), X.eq(E.v(23,28,5),8.0))))) else 0.0))

def f58(E, r, c):
    # Ô F44, F50
    # = IF(F43>=0,IF(F43=TRUNG_GIAN!$E$7,1,0),IF(OR(AND(OR(TRUNG_GIAN!$F$28=3,TRUNG_GIAN!$F$28=4),TRUNG_GIAN!$F$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(T
    #   RUNG_GIAN!$F$28=0,TRUNG_GIAN!$F$28=8))),1,0))
    return ((1.0 if X.truthy(X.eq(E.v(8,r-1,c),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(8,r-1,c),0.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,28,6),3.0), X.eq(E.v(23,28,6),4.0)), X.eq(E.v(23,21,6),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq(E.v(23,28,6),0.0), X.eq(E.v(23,28,6),8.0))))) else 0.0))

def f59(E, r, c):
    # Ô C45, C51
    # = IF(AND(C41=0,OR(TRUNG_GIAN!$C$6=CHOOSE(TRUNG_GIAN!$E$7+1,8,11,2,6,10),TRUNG_GIAN!$C$6=CHOOSE(TRUNG_GIAN!$E$7+1,9,0,3,5,7))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(8,r-4,c),0.0), X.OR(X.eq(E.v(23,6,3),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 8.0, lambda: 11.0, lambda: 2.0, lambda: 6.0, lambda: 10.0)), X.eq(E.v(23,6,3),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 9.0, lambda: 0.0, lambda: 3.0, lambda: 5.0, lambda: 7.0))))) else 0.0)

def f60(E, r, c):
    # Ô D45, D51
    # = IF(AND(D41=0,OR(TRUNG_GIAN!$D$6=CHOOSE(TRUNG_GIAN!$E$7+1,8,11,2,6,10),TRUNG_GIAN!$D$6=CHOOSE(TRUNG_GIAN!$E$7+1,9,0,3,5,7))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(8,r-4,c),0.0), X.OR(X.eq(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 8.0, lambda: 11.0, lambda: 2.0, lambda: 6.0, lambda: 10.0)), X.eq(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 9.0, lambda: 0.0, lambda: 3.0, lambda: 5.0, lambda: 7.0))))) else 0.0)

def f61(E, r, c):
    # Ô E45, E51
    # = IF(AND(E41=0,OR(TRUNG_GIAN!$E$6=CHOOSE(TRUNG_GIAN!$E$7+1,8,11,2,6,10),TRUNG_GIAN!$E$6=CHOOSE(TRUNG_GIAN!$E$7+1,9,0,3,5,7))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(8,r-4,c),0.0), X.OR(X.eq(E.v(23,6,5),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 8.0, lambda: 11.0, lambda: 2.0, lambda: 6.0, lambda: 10.0)), X.eq(E.v(23,6,5),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 9.0, lambda: 0.0, lambda: 3.0, lambda: 5.0, lambda: 7.0))))) else 0.0)

def f62(E, r, c):
    # Ô F45, F51
    # = IF(AND(F41=0,OR(TRUNG_GIAN!$F$6=CHOOSE(TRUNG_GIAN!$E$7+1,8,11,2,6,10),TRUNG_GIAN!$F$6=CHOOSE(TRUNG_GIAN!$E$7+1,9,0,3,5,7))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(8,r-4,c),0.0), X.OR(X.eq(E.v(23,6,6),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 8.0, lambda: 11.0, lambda: 2.0, lambda: 6.0, lambda: 10.0)), X.eq(E.v(23,6,6),X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 9.0, lambda: 0.0, lambda: 3.0, lambda: 5.0, lambda: 7.0))))) else 0.0)

def f63(E, r, c):
    # Ô C47
    # = IF(OR(AND(TRUNG_GIAN!$C$52=1,OR(PP2_DINH_LUONG!$C$21=1,TRUNG_GIAN!$C$45="Tam hội")),PP2_DINH_LUONG!$C$22=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,3),1.0), X.OR(X.eq(E.v(10,21,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(10,22,3),1.0))) else 0.0)

def f64(E, r, c):
    # Ô D47
    # = IF(OR(AND(TRUNG_GIAN!$D$52=1,OR(PP2_DINH_LUONG!$C$21=1,TRUNG_GIAN!$C$45="Tam hội")),PP2_DINH_LUONG!$C$22=1,PP2_DINH_LUONG!$D$22=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,4),1.0), X.OR(X.eq(E.v(10,21,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(10,22,3),1.0), X.eq(E.v(10,22,4),1.0))) else 0.0)

def f65(E, r, c):
    # Ô E47
    # = IF(OR(AND(TRUNG_GIAN!$E$52=1,OR(PP2_DINH_LUONG!$C$21=1,TRUNG_GIAN!$C$45="Tam hội")),PP2_DINH_LUONG!$D$22=1,PP2_DINH_LUONG!$E$22=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,5),1.0), X.OR(X.eq(E.v(10,21,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(10,22,4),1.0), X.eq(E.v(10,22,5),1.0))) else 0.0)

def f66(E, r, c):
    # Ô F47
    # = IF(OR(AND(TRUNG_GIAN!$F$52=1,OR(PP2_DINH_LUONG!$C$21=1,TRUNG_GIAN!$C$45="Tam hội")),PP2_DINH_LUONG!$E$22=1),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(23,52,6),1.0), X.OR(X.eq(E.v(10,21,3),1.0), X.eq(E.v(23,45,3),'Tam hội'))), X.eq(E.v(10,22,5),1.0))) else 0.0)

def f67(E, r, c):
    # Ô C48
    # = IF(C47=0,-1,TRUNG_GIAN!$C$61)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(23,61,3))

def f68(E, r, c):
    # Ô D48
    # = IF(D47=0,-1,TRUNG_GIAN!$D$61)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(23,61,4))

def f69(E, r, c):
    # Ô E48
    # = IF(E47=0,-1,TRUNG_GIAN!$E$61)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(23,61,5))

def f70(E, r, c):
    # Ô F48
    # = IF(F47=0,-1,TRUNG_GIAN!$F$61)
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else E.v(23,61,6))

def f71(E, r, c):
    # Ô C49
    # = PP2_DINH_LUONG!$C$24
    return E.v(10,24,3)

def f72(E, r, c):
    # Ô D49
    # = PP2_DINH_LUONG!$D$24
    return E.v(10,24,4)

def f73(E, r, c):
    # Ô E49
    # = PP2_DINH_LUONG!$E$24
    return E.v(10,24,5)

def f74(E, r, c):
    # Ô F49
    # = PP2_DINH_LUONG!$F$24
    return E.v(10,24,6)

def f75(E, r, c):
    # Ô C53
    # = (COUNTIF(TRUNG_GIAN!$C$6:$F$6,4)>0)+(COUNTIF(TRUNG_GIAN!$C$6:$F$6,10)>0)+(COUNTIF(TRUNG_GIAN!$C$6:$F$6,1)>0)+(COUNTIF(TRUNG_GIAN!$C$6:$F$6,7)>0)
    return X.add(X.add(X.add(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 4.0),0.0),X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 10.0),0.0)),X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 1.0),0.0)),X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 7.0),0.0))

def f76(E, r, c):
    # Ô C57
    # = IF(CHOOSE(TRUNG_GIAN!$E$7+1,OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$D$6=4,$D$42=0),OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7
    #   ,$D$42=1),OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$D$6=7),OR(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$D$6=9),OR(TRUNG_GIAN!$D$6=11,T
    #   RUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=4,$D$42=4)),1,0)
    return (1.0 if X.truthy(X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,4),4.0), X.eq(E.v(8,42,4),0.0)), lambda: X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0), X.eq(E.v(8,42,4),1.0)), lambda: X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,4),7.0)), lambda: X.OR(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),9.0)), lambda: X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),4.0), X.eq(E.v(8,42,4),4.0)))) else 0.0)

def f77(E, r, c):
    # Ô D57
    # = IF(CHOOSE(TRUNG_GIAN!$E$7+1,OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$D$6=4,$D$48=0),OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7
    #   ,$D$48=1),OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$D$6=7),OR(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$D$6=9),OR(TRUNG_GIAN!$D$6=11,T
    #   RUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=4,$D$48=4)),1,0)
    return (1.0 if X.truthy(X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,4),4.0), X.eq(E.v(8,48,4),0.0)), lambda: X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0), X.eq(E.v(8,48,4),1.0)), lambda: X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,4),7.0)), lambda: X.OR(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),9.0)), lambda: X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),4.0), X.eq(E.v(8,48,4),4.0)))) else 0.0)

def f78(E, r, c):
    # Ô C58, D58
    # = SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(TRUNG_GIAN!$E$7+3,5))*$C$39:$F$39)
    return X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(23,7,5),3.0), 5.0)),E.rg(8,39,3,39,6)))

def f79(E, r, c):
    # Ô C59
    # = SUM($C$45:$F$45)
    return X.SUM(E.rg(8,45,3,45,6))

def f80(E, r, c):
    # Ô D59
    # = SUM($C$51:$F$51)
    return X.SUM(E.rg(8,51,3,51,6))

def f81(E, r, c):
    # Ô C60
    # = IF(CHOOSE(TRUNG_GIAN!$E$7+1,OR(AND(AND(TRUNG_GIAN!$C$44>0,PP1_DINH_TINH!$C$8=0),TRUNG_GIAN!$C$46=0),C29=0),OR(AND(AND(TRUNG_GIAN!$C$44>0,PP1_DINH_TINH
    #   !$C$8=0),TRUNG_GIAN!$C$46=1),C29=0,C29=1),OR($C$53>=3,COUNTIF(TRUNG_GIAN!$C$7:$F$7,2)=4),AND(AND(TRUNG_GIAN!$C$44>0,PP1_DINH_TINH!$C$8=0),TRUNG_GIAN!$
    #   C$46=3),OR(AND(AND(TRUNG_GIAN!$C$44>0,PP1_DINH_TINH!$C$8=0),TRUNG_GIAN!$C$46=4),C29=4)),1,0)
    return (1.0 if X.truthy(X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: X.OR(X.AND(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(23,46,3),0.0)), X.eq(E.v(8,r-31,c),0.0)), lambda: X.OR(X.AND(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(23,46,3),1.0)), X.eq(E.v(8,r-31,c),0.0), X.eq(E.v(8,r-31,c),1.0)), lambda: X.OR(X.ge(E.v(8,53,3),3.0), X.eq(X.COUNTIF(E.rg(23,7,3,7,6), 2.0),4.0)), lambda: X.AND(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(23,46,3),3.0)), lambda: X.OR(X.AND(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(9,8,3),0.0)), X.eq(E.v(23,46,3),4.0)), X.eq(E.v(8,r-31,c),4.0)))) else 0.0)

def f82(E, r, c):
    # Ô D60
    # = IF(CHOOSE(TRUNG_GIAN!$E$7+1,OR(AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=0),D29=0),OR(AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=1),D29=0,D29=1),OR($C$53>=
    #   3,COUNTIF(TRUNG_GIAN!$C$7:$F$7,2)=4),AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=3),OR(AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=4),D29=4)),1,0)
    return (1.0 if X.truthy(X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: X.OR(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),0.0)), X.eq(E.v(8,r-31,c),0.0)), lambda: X.OR(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),1.0)), X.eq(E.v(8,r-31,c),0.0), X.eq(E.v(8,r-31,c),1.0)), lambda: X.OR(X.ge(E.v(8,53,3),3.0), X.eq(X.COUNTIF(E.rg(23,7,3,7,6), 2.0),4.0)), lambda: X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),3.0)), lambda: X.OR(X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),4.0)), X.eq(E.v(8,r-31,c),4.0)))) else 0.0)

def f83(E, r, c):
    # Ô C61, D61
    # = IF(AND(C57=1,C58+C59=0,C60=1),CHOOSE(TRUNG_GIAN!$E$7+1,"Khúc trực","Viêm thượng","Gia tường","Tòng cách","Nhuận hạ"),"")
    return (X.CHOOSE(X.add(E.v(23,7,5),1.0), lambda: 'Khúc trực', lambda: 'Viêm thượng', lambda: 'Gia tường', lambda: 'Tòng cách', lambda: 'Nhuận hạ') if X.truthy(X.AND(X.eq(E.v(8,r-4,c),1.0), X.eq(X.add(E.v(8,r-3,c),E.v(8,r-2,c)),0.0), X.eq(E.v(8,r-1,c),1.0))) else '')

def f84(E, r, c):
    # Ô C62
    # = IF(SUM($C$44:$F$44)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUM(E.rg(8,44,3,44,6)),0.0)) else 0.0)

def f85(E, r, c):
    # Ô D62
    # = IF(SUM($C$50:$F$50)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.SUM(E.rg(8,50,3,50,6)),0.0)) else 0.0)

def f86(E, r, c):
    # Ô C63
    # = SUMPRODUCT((PP1_DINH_TINH!$C$37:$F$37=C$5)*$C$39:$F$39)+SUMPRODUCT((PP1_DINH_TINH!$C$37:$F$37=MOD(C$5+4,5))*$C$39:$F$39)+SUMPRODUCT(($C$40:$F$40=C$5)*
    #   (1-$C$41:$F$41))+SUMPRODUCT(($C$40:$F$40=MOD(C$5+4,5))*(1-$C$41:$F$41))
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(9,37,3,37,6),E.v(8,5,c)),E.rg(8,39,3,39,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(9,37,3,37,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),E.rg(8,39,3,39,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),E.v(8,5,c)),X.A('sub',1.0,E.rg(8,41,3,41,6))))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),X.A('sub',1.0,E.rg(8,41,3,41,6)))))

def f87(E, r, c):
    # Ô D63
    # = SUMPRODUCT((PP2_DINH_LUONG!$C$52:$F$52=D$5)*$C$39:$F$39)+SUMPRODUCT((PP2_DINH_LUONG!$C$52:$F$52=MOD(D$5+4,5))*$C$39:$F$39)+SUMPRODUCT(($C$40:$F$40=D$5
    #   )*(1-$C$47:$F$47))+SUMPRODUCT(($C$40:$F$40=MOD(D$5+4,5))*(1-$C$47:$F$47))
    return X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),E.v(8,5,c)),E.rg(8,39,3,39,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(10,52,3,52,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),E.rg(8,39,3,39,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),E.v(8,5,c)),X.A('sub',1.0,E.rg(8,47,3,47,6))))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),X.A('sub',1.0,E.rg(8,47,3,47,6)))))

def f88(E, r, c):
    # Ô C64, D64
    # = IF(OR(INDEX(C$24:C$28,MOD(C$5+2,5)+1)>0,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+2,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0),1,0)
    return (1.0 if X.truthy(X.OR(X.gt(X.INDEX(E.rg(8,24,c,28,c), X.add(X.MOD(X.add(E.v(8,5,c),2.0), 5.0),1.0)),0.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),2.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0))) else 0.0)

def f89(E, r, c):
    # Ô C65, D65
    # = IF(AND(C62=0,C63=0,C29=MOD(C$5+2,5)),"Tòng tài","")
    return ('Tòng tài' if X.truthy(X.AND(X.eq(E.v(8,r-3,c),0.0), X.eq(E.v(8,r-2,c),0.0), X.eq(E.v(8,r-36,c),X.MOD(X.add(E.v(8,5,c),2.0), 5.0)))) else '')

def f90(E, r, c):
    # Ô C66, D66
    # = IF(AND(C62=0,C63=0,C29=MOD(C$5+3,5),SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*$C$39:$F$39)=0,C64=1),"Tòng sát","")
    return ('Tòng sát' if X.truthy(X.AND(X.eq(E.v(8,r-4,c),0.0), X.eq(E.v(8,r-3,c),0.0), X.eq(E.v(8,r-37,c),X.MOD(X.add(E.v(8,5,c),3.0), 5.0)), X.eq(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),E.rg(8,39,3,39,6))),0.0), X.eq(E.v(8,r-2,c),1.0))) else '')

def f91(E, r, c):
    # Ô C67, D67
    # = IF(OR(TRUNG_GIAN!$D$21=MOD(C$5+1,5),TRUNG_GIAN!$D$22=MOD(C$5+1,5),TRUNG_GIAN!$D$23=MOD(C$5+1,5)),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,21,4),X.MOD(X.add(E.v(8,5,c),1.0), 5.0)), X.eq(E.v(23,22,4),X.MOD(X.add(E.v(8,5,c),1.0), 5.0)), X.eq(E.v(23,23,4),X.MOD(X.add(E.v(8,5,c),1.0), 5.0)))) else 0.0)

def f92(E, r, c):
    # Ô C68
    # = IF(AND(C67=1,C29=MOD(C$5+1,5),C64=1,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+3,5))*$C$39:$F$39)+SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+4,5))*$C$39:$
    #   F$39)+SUMPRODUCT(($C$40:$F$40=MOD(C$5+3,5))*(1-$C$41:$F$41))+SUMPRODUCT(($C$40:$F$40=MOD(C$5+4,5))*(1-$C$41:$F$41))=0),"Tòng nhi","")
    return ('Tòng nhi' if X.truthy(X.AND(X.eq(E.v(8,r-1,c),1.0), X.eq(E.v(8,r-39,c),X.MOD(X.add(E.v(8,5,c),1.0), 5.0)), X.eq(E.v(8,r-4,c),1.0), X.eq(X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),E.rg(8,39,3,39,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),E.rg(8,39,3,39,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),X.A('sub',1.0,E.rg(8,41,3,41,6))))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),X.A('sub',1.0,E.rg(8,41,3,41,6))))),0.0))) else '')

def f93(E, r, c):
    # Ô D68
    # = IF(AND(D67=1,D29=MOD(D$5+1,5),D64=1,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+3,5))*$C$39:$F$39)+SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+4,5))*$C$39:$
    #   F$39)+SUMPRODUCT(($C$40:$F$40=MOD(D$5+3,5))*(1-$C$47:$F$47))+SUMPRODUCT(($C$40:$F$40=MOD(D$5+4,5))*(1-$C$47:$F$47))=0),"Tòng nhi","")
    return ('Tòng nhi' if X.truthy(X.AND(X.eq(E.v(8,r-1,c),1.0), X.eq(E.v(8,r-39,c),X.MOD(X.add(E.v(8,5,c),1.0), 5.0)), X.eq(E.v(8,r-4,c),1.0), X.eq(X.add(X.add(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),E.rg(8,39,3,39,6))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),E.rg(8,39,3,39,6)))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),X.A('sub',1.0,E.rg(8,47,3,47,6))))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(8,40,3,40,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),X.A('sub',1.0,E.rg(8,47,3,47,6))))),0.0))) else '')

def f94(E, r, c):
    # Ô C69
    # = IF(AND(PP1_DINH_TINH!$D$34=1,OR(TRUNG_GIAN!$D$9=TRUNG_GIAN!$D$76,$D$42=TRUNG_GIAN!$D$76),C29=TRUNG_GIAN!$D$76),CHOOSE(TRUNG_GIAN!$D$76+1,"Hóa mộc","Hó
    #   a hỏa","Hóa thổ","Hóa kim","Hóa thủy"),"")
    return (X.CHOOSE(X.add(E.v(23,76,4),1.0), lambda: 'Hóa mộc', lambda: 'Hóa hỏa', lambda: 'Hóa thổ', lambda: 'Hóa kim', lambda: 'Hóa thủy') if X.truthy(X.AND(X.eq(E.v(9,34,4),1.0), X.OR(X.eq(E.v(23,9,4),E.v(23,76,4)), X.eq(E.v(8,42,4),E.v(23,76,4))), X.eq(E.v(8,r-40,c),E.v(23,76,4)))) else '')

def f95(E, r, c):
    # Ô D69
    # = IF(AND(PP2_DINH_LUONG!$D$47=1,OR(TRUNG_GIAN!$D$9=PP2_DINH_LUONG!$D$28,$D$48=PP2_DINH_LUONG!$D$28),D29=PP2_DINH_LUONG!$D$28),CHOOSE(PP2_DINH_LUONG!$D$2
    #   8+1,"Hóa mộc","Hóa hỏa","Hóa thổ","Hóa kim","Hóa thủy"),"")
    return (X.CHOOSE(X.add(E.v(10,28,4),1.0), lambda: 'Hóa mộc', lambda: 'Hóa hỏa', lambda: 'Hóa thổ', lambda: 'Hóa kim', lambda: 'Hóa thủy') if X.truthy(X.AND(X.eq(E.v(10,47,4),1.0), X.OR(X.eq(E.v(23,9,4),E.v(10,28,4)), X.eq(E.v(8,48,4),E.v(10,28,4))), X.eq(E.v(8,r-40,c),E.v(10,28,4)))) else '')

def f96(E, r, c):
    # Ô C70
    # = IF(AND(PP1_DINH_TINH!$E$34=1,OR(TRUNG_GIAN!$D$9=TRUNG_GIAN!$E$76,$D$42=TRUNG_GIAN!$E$76),C29=TRUNG_GIAN!$E$76),CHOOSE(TRUNG_GIAN!$E$76+1,"Hóa mộc","Hó
    #   a hỏa","Hóa thổ","Hóa kim","Hóa thủy"),"")
    return (X.CHOOSE(X.add(E.v(23,76,5),1.0), lambda: 'Hóa mộc', lambda: 'Hóa hỏa', lambda: 'Hóa thổ', lambda: 'Hóa kim', lambda: 'Hóa thủy') if X.truthy(X.AND(X.eq(E.v(9,34,5),1.0), X.OR(X.eq(E.v(23,9,4),E.v(23,76,5)), X.eq(E.v(8,42,4),E.v(23,76,5))), X.eq(E.v(8,r-41,c),E.v(23,76,5)))) else '')

def f97(E, r, c):
    # Ô D70
    # = IF(AND(PP2_DINH_LUONG!$E$47=1,OR(TRUNG_GIAN!$D$9=PP2_DINH_LUONG!$E$28,$D$48=PP2_DINH_LUONG!$E$28),D29=PP2_DINH_LUONG!$E$28),CHOOSE(PP2_DINH_LUONG!$E$2
    #   8+1,"Hóa mộc","Hóa hỏa","Hóa thổ","Hóa kim","Hóa thủy"),"")
    return (X.CHOOSE(X.add(E.v(10,28,5),1.0), lambda: 'Hóa mộc', lambda: 'Hóa hỏa', lambda: 'Hóa thổ', lambda: 'Hóa kim', lambda: 'Hóa thủy') if X.truthy(X.AND(X.eq(E.v(10,47,5),1.0), X.OR(X.eq(E.v(23,9,4),E.v(10,28,5)), X.eq(E.v(8,48,4),E.v(10,28,5))), X.eq(E.v(8,r-41,c),E.v(10,28,5)))) else '')

def f98(E, r, c):
    # Ô C71
    # = ""
    return ''

def f99(E, r, c):
    # Ô D71
    # = IF(D6="Độc vượng","Tòng vượng (độc vượng)","")
    return ('Tòng vượng (độc vượng)' if X.truthy(X.eq(E.v(8,r-65,c),'Độc vượng')) else '')

def f100(E, r, c):
    # Ô C72
    # = IF($C$205<>"","",$C$202)
    return ('' if X.truthy(X.ne(E.v(8,205,3),'')) else E.v(8,202,3))

def f101(E, r, c):
    # Ô D72
    # = IF($D$205<>"","",$D$202)
    return ('' if X.truthy(X.ne(E.v(8,205,4),'')) else E.v(8,202,4))

def f102(E, r, c):
    # Ô C73, D73
    # = IF(C72="",0,MATCH(C72,SP_TEN,0))
    return (0.0 if X.truthy(X.eq(E.v(8,r-1,c),'')) else X.MATCH(E.v(8,r-1,c), E.rg(24,103,1,116,1), 0.0))

def f103(E, r, c):
    # Ô C74, D74
    # = IF(C73=0,-1,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$C$103:$C$116,C$73)-C$5,5),INDEX(CSDL!$C$103:$C$116,C$73)))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),0.0)) else (X.MOD(X.sub(X.INDEX(E.rg(24,103,3,116,3), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,3,116,3), E.v(8,73,c))))

def f104(E, r, c):
    # Ô C75, D75
    # = IF(C73=0,-1,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$D$103:$D$116,C$73)-C$5,5),INDEX(CSDL!$D$103:$D$116,C$73)))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-2,c),0.0)) else (X.MOD(X.sub(X.INDEX(E.rg(24,103,4,116,4), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,4,116,4), E.v(8,73,c))))

def f105(E, r, c):
    # Ô C76, D76
    # = IF(C73=0,-1,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$E$103:$E$116,C$73)-C$5,5),INDEX(CSDL!$E$103:$E$116,C$73)))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-3,c),0.0)) else (X.MOD(X.sub(X.INDEX(E.rg(24,103,5,116,5), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,5,116,5), E.v(8,73,c))))

def f106(E, r, c):
    # Ô C77, D77
    # = IF(C73=0,"",IF(INDEX(CSDL!$F$103:$F$116,C$73)="","",IF(0>0,", ","")&INDEX(HANH_TEN,MOD(IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",INDEX(CSDL!$F$103:$F$116,
    #   C$73),INDEX(CSDL!$F$103:$F$116,C$73)+C$5),5)+1)&"/"&INDEX(NHOM_TEN,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$F$103:$F$116,C$73)-C$5,5),IND
    #   EX(CSDL!$F$103:$F$116,C$73))+1))&IF(INDEX(CSDL!$G$103:$G$116,C$73)="","",IF(1>0,", ","")&INDEX(HANH_TEN,MOD(IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",INDE
    #   X(CSDL!$G$103:$G$116,C$73),INDEX(CSDL!$G$103:$G$116,C$73)+C$5),5)+1)&"/"&INDEX(NHOM_TEN,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$G$103:$G
    #   $116,C$73)-C$5,5),INDEX(CSDL!$G$103:$G$116,C$73))+1))&IF(INDEX(CSDL!$H$103:$H$116,C$73)="","",IF(2>0,", ","")&INDEX(HANH_TEN,MOD(IF(INDEX(CSDL!$B$103:
    #   $B$116,C$73)="H",INDEX(CSDL!$H$103:$H$116,C$73),INDEX(CSDL!$H$103:$H$116,C$73)+C$5),5)+1)&"/"&INDEX(NHOM_TEN,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD
    #   (INDEX(CSDL!$H$103:$H$116,C$73)-C$5,5),INDEX(CSDL!$H$103:$H$116,C$73))+1)))
    return ('' if X.truthy(X.eq(E.v(8,r-4,c),0.0)) else X.cat(X.cat(('' if X.truthy(X.eq(X.INDEX(E.rg(24,103,6,116,6), E.v(8,73,c)),'')) else X.cat(X.cat(X.cat((', ' if X.truthy(X.gt(0.0,0.0)) else ''),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,103,6,116,6), E.v(8,73,c)) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.add(X.INDEX(E.rg(24,103,6,116,6), E.v(8,73,c)),E.v(8,5,c))), 5.0),1.0))),'/'),X.INDEX(E.rg(24,54,1,58,1), X.add((X.MOD(X.sub(X.INDEX(E.rg(24,103,6,116,6), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,6,116,6), E.v(8,73,c))),1.0)))),('' if X.truthy(X.eq(X.INDEX(E.rg(24,103,7,116,7), E.v(8,73,c)),'')) else X.cat(X.cat(X.cat((', ' if X.truthy(X.gt(1.0,0.0)) else ''),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,103,7,116,7), E.v(8,73,c)) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.add(X.INDEX(E.rg(24,103,7,116,7), E.v(8,73,c)),E.v(8,5,c))), 5.0),1.0))),'/'),X.INDEX(E.rg(24,54,1,58,1), X.add((X.MOD(X.sub(X.INDEX(E.rg(24,103,7,116,7), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,7,116,7), E.v(8,73,c))),1.0))))),('' if X.truthy(X.eq(X.INDEX(E.rg(24,103,8,116,8), E.v(8,73,c)),'')) else X.cat(X.cat(X.cat((', ' if X.truthy(X.gt(2.0,0.0)) else ''),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,103,8,116,8), E.v(8,73,c)) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.add(X.INDEX(E.rg(24,103,8,116,8), E.v(8,73,c)),E.v(8,5,c))), 5.0),1.0))),'/'),X.INDEX(E.rg(24,54,1,58,1), X.add((X.MOD(X.sub(X.INDEX(E.rg(24,103,8,116,8), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,8,116,8), E.v(8,73,c))),1.0))))))

def f107(E, r, c):
    # Ô C78, D78
    # = IF(C73=0,"",IF(INDEX(CSDL!$I$103:$I$116,C$73)="","",IF(0>0,", ","")&INDEX(HANH_TEN,MOD(IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",INDEX(CSDL!$I$103:$I$116,
    #   C$73),INDEX(CSDL!$I$103:$I$116,C$73)+C$5),5)+1)&"/"&INDEX(NHOM_TEN,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$I$103:$I$116,C$73)-C$5,5),IND
    #   EX(CSDL!$I$103:$I$116,C$73))+1))&IF(INDEX(CSDL!$J$103:$J$116,C$73)="","",IF(1>0,", ","")&INDEX(HANH_TEN,MOD(IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",INDE
    #   X(CSDL!$J$103:$J$116,C$73),INDEX(CSDL!$J$103:$J$116,C$73)+C$5),5)+1)&"/"&INDEX(NHOM_TEN,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD(INDEX(CSDL!$J$103:$J
    #   $116,C$73)-C$5,5),INDEX(CSDL!$J$103:$J$116,C$73))+1))&IF(INDEX(CSDL!$K$103:$K$116,C$73)="","",IF(2>0,", ","")&INDEX(HANH_TEN,MOD(IF(INDEX(CSDL!$B$103:
    #   $B$116,C$73)="H",INDEX(CSDL!$K$103:$K$116,C$73),INDEX(CSDL!$K$103:$K$116,C$73)+C$5),5)+1)&"/"&INDEX(NHOM_TEN,IF(INDEX(CSDL!$B$103:$B$116,C$73)="H",MOD
    #   (INDEX(CSDL!$K$103:$K$116,C$73)-C$5,5),INDEX(CSDL!$K$103:$K$116,C$73))+1)))
    return ('' if X.truthy(X.eq(E.v(8,r-5,c),0.0)) else X.cat(X.cat(('' if X.truthy(X.eq(X.INDEX(E.rg(24,103,9,116,9), E.v(8,73,c)),'')) else X.cat(X.cat(X.cat((', ' if X.truthy(X.gt(0.0,0.0)) else ''),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,103,9,116,9), E.v(8,73,c)) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.add(X.INDEX(E.rg(24,103,9,116,9), E.v(8,73,c)),E.v(8,5,c))), 5.0),1.0))),'/'),X.INDEX(E.rg(24,54,1,58,1), X.add((X.MOD(X.sub(X.INDEX(E.rg(24,103,9,116,9), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,9,116,9), E.v(8,73,c))),1.0)))),('' if X.truthy(X.eq(X.INDEX(E.rg(24,103,10,116,10), E.v(8,73,c)),'')) else X.cat(X.cat(X.cat((', ' if X.truthy(X.gt(1.0,0.0)) else ''),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,103,10,116,10), E.v(8,73,c)) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.add(X.INDEX(E.rg(24,103,10,116,10), E.v(8,73,c)),E.v(8,5,c))), 5.0),1.0))),'/'),X.INDEX(E.rg(24,54,1,58,1), X.add((X.MOD(X.sub(X.INDEX(E.rg(24,103,10,116,10), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,10,116,10), E.v(8,73,c))),1.0))))),('' if X.truthy(X.eq(X.INDEX(E.rg(24,103,11,116,11), E.v(8,73,c)),'')) else X.cat(X.cat(X.cat((', ' if X.truthy(X.gt(2.0,0.0)) else ''),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,103,11,116,11), E.v(8,73,c)) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.add(X.INDEX(E.rg(24,103,11,116,11), E.v(8,73,c)),E.v(8,5,c))), 5.0),1.0))),'/'),X.INDEX(E.rg(24,54,1,58,1), X.add((X.MOD(X.sub(X.INDEX(E.rg(24,103,11,116,11), E.v(8,73,c)),E.v(8,5,c)), 5.0) if X.truthy(X.eq(X.INDEX(E.rg(24,103,2,116,2), E.v(8,73,c)),'H')) else X.INDEX(E.rg(24,103,11,116,11), E.v(8,73,c))),1.0))))))

def f108(E, r, c):
    # Ô C81, D81
    # = IF(ABS(C$15-MAX(C$15,C$14,C$13))<0.000000001,0,1000)+C$34+0/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,15,c),X.MAX(E.rg(8,15,c,15,c), E.rg(8,14,c,14,c), E.rg(8,13,c,13,c)))),1e-09)) else 1000.0),E.v(8,34,c)),X.div(0.0,10.0))

def f109(E, r, c):
    # Ô C82, D82
    # = IF(ABS(C$14-MAX(C$15,C$14,C$13))<0.000000001,0,1000)+C$33+1/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,14,c),X.MAX(E.rg(8,15,c,15,c), E.rg(8,14,c,14,c), E.rg(8,13,c,13,c)))),1e-09)) else 1000.0),E.v(8,33,c)),X.div(1.0,10.0))

def f110(E, r, c):
    # Ô C83, D83
    # = IF(ABS(C$13-MAX(C$15,C$14,C$13))<0.000000001,0,1000)+C$32+2/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,13,c),X.MAX(E.rg(8,15,c,15,c), E.rg(8,14,c,14,c), E.rg(8,13,c,13,c)))),1e-09)) else 1000.0),E.v(8,32,c)),X.div(2.0,10.0))

def f111(E, r, c):
    # Ô C84, D84
    # = IF(ABS(C$16-MAX(C$16,C$12))<0.000000001,0,1000)+C$35+0/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,16,c),X.MAX(E.rg(8,16,c,16,c), E.rg(8,12,c,12,c)))),1e-09)) else 1000.0),E.v(8,35,c)),X.div(0.0,10.0))

def f112(E, r, c):
    # Ô C85, D85
    # = IF(ABS(C$12-MAX(C$16,C$12))<0.000000001,0,1000)+C$31+1/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,12,c),X.MAX(E.rg(8,16,c,16,c), E.rg(8,12,c,12,c)))),1e-09)) else 1000.0),E.v(8,31,c)),X.div(1.0,10.0))

def f113(E, r, c):
    # Ô C86, D86
    # = IF(ABS(C$14-MAX(C$14,C$15))<0.000000001,0,1000)+C$33+0/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,14,c),X.MAX(E.rg(8,14,c,14,c), E.rg(8,15,c,15,c)))),1e-09)) else 1000.0),E.v(8,33,c)),X.div(0.0,10.0))

def f114(E, r, c):
    # Ô C87, D87
    # = IF(ABS(C$15-MAX(C$14,C$15))<0.000000001,0,1000)+C$34+1/10
    return X.add(X.add((0.0 if X.truthy(X.lt(X.ABS(X.sub(E.v(8,15,c),X.MAX(E.rg(8,14,c,14,c), E.rg(8,15,c,15,c)))),1e-09)) else 1000.0),E.v(8,34,c)),X.div(1.0,10.0))

def f115(E, r, c):
    # Ô C88, D88
    # = CHOOSE(MATCH(MIN(C81:C83),C81:C83,0),3,2,1)
    return X.CHOOSE(X.MATCH(X.MIN(E.rg(8,r-7,c,r-5,c)), E.rg(8,r-7,c,r-5,c), 0.0), lambda: 3.0, lambda: 2.0, lambda: 1.0)

def f116(E, r, c):
    # Ô C89, D89
    # = CHOOSE(MATCH(MIN(C84:C85),C84:C85,0),4,0)
    return X.CHOOSE(X.MATCH(X.MIN(E.rg(8,r-5,c,r-4,c)), E.rg(8,r-5,c,r-4,c), 0.0), lambda: 4.0, lambda: 0.0)

def f117(E, r, c):
    # Ô C90, D90
    # = CHOOSE(MATCH(MIN(C86:C87),C86:C87,0),2,3)
    return X.CHOOSE(X.MATCH(X.MIN(E.rg(8,r-4,c,r-3,c)), E.rg(8,r-4,c,r-3,c), 0.0), lambda: 2.0, lambda: 3.0)

def f118(E, r, c):
    # Ô C91, D91
    # = IF(C7=2,"TH",IF(C7=0,"NHUOC",IF(MAX(C14,C15)>INDEX(C12:C16,C89+1),IF(C90=3,"V_QS","V_TAI"),IF(AND(C89=0,C10=1,C20=1),"V_ATG",IF(C89=4,"V_AN","V_TK")))
    #   ))
    return ('TH' if X.truthy(X.eq(E.v(8,r-84,c),2.0)) else ('NHUOC' if X.truthy(X.eq(E.v(8,r-84,c),0.0)) else (('V_QS' if X.truthy(X.eq(E.v(8,r-1,c),3.0)) else 'V_TAI') if X.truthy(X.gt(X.MAX(E.rg(8,r-77,c,r-77,c), E.rg(8,r-76,c,r-76,c)),X.INDEX(E.rg(8,r-79,c,r-75,c), X.add(E.v(8,r-2,c),1.0)))) else ('V_ATG' if X.truthy(X.AND(X.eq(E.v(8,r-2,c),0.0), X.eq(E.v(8,r-81,c),1.0), X.eq(E.v(8,r-71,c),1.0))) else ('V_AN' if X.truthy(X.eq(E.v(8,r-2,c),4.0)) else 'V_TK')))))

def f119(E, r, c):
    # Ô C92, D92
    # = IF(C91="TH",-1,IF(C91="NHUOC",C88,IF(OR(C91="V_QS",C91="V_TAI"),C90,IF(C91="V_ATG",4,C89))))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-1,c),'TH')) else (E.v(8,r-4,c) if X.truthy(X.eq(E.v(8,r-1,c),'NHUOC')) else (E.v(8,r-2,c) if X.truthy(X.OR(X.eq(E.v(8,r-1,c),'V_QS'), X.eq(E.v(8,r-1,c),'V_TAI'))) else (4.0 if X.truthy(X.eq(E.v(8,r-1,c),'V_ATG')) else E.v(8,r-3,c)))))

def f120(E, r, c):
    # Ô C93, D93
    # = CHOOSE(MATCH(C91,{"NHUOC","V_QS","V_TAI","V_ATG","V_AN","V_TK","TH"},0),IF(C88=2,0,4),1,3,2,2,3,2)
    return X.CHOOSE(X.MATCH(E.v(8,r-2,c), X.Arr([['NHUOC', 'V_QS', 'V_TAI', 'V_ATG', 'V_AN', 'V_TK', 'TH']]), 0.0), lambda: (0.0 if X.truthy(X.eq(E.v(8,r-5,c),2.0)) else 4.0), lambda: 1.0, lambda: 3.0, lambda: 2.0, lambda: 2.0, lambda: 3.0, lambda: 2.0)

def f121(E, r, c):
    # Ô C94, D94
    # = CHOOSE(MATCH(C91,{"NHUOC","V_QS","V_TAI","V_ATG","V_AN","V_TK","TH"},0),IF(C88=2,4,0),-1,1,1,IF(C9=0,3,1),1,-1)
    return X.CHOOSE(X.MATCH(E.v(8,r-3,c), X.Arr([['NHUOC', 'V_QS', 'V_TAI', 'V_ATG', 'V_AN', 'V_TK', 'TH']]), 0.0), lambda: (4.0 if X.truthy(X.eq(E.v(8,r-6,c),2.0)) else 0.0), lambda: X.neg(1.0), lambda: 1.0, lambda: 1.0, lambda: (3.0 if X.truthy(X.eq(E.v(8,r-85,c),0.0)) else 1.0), lambda: 1.0, lambda: X.neg(1.0))

def f122(E, r, c):
    # Ô C95, D95
    # = CHOOSE(MATCH(C91,{"NHUOC","V_QS","V_TAI","V_ATG","V_AN","V_TK","TH"},0),-1,-1,-1,-1,IF(C9=0,1,-1),2,-1)
    return X.CHOOSE(X.MATCH(E.v(8,r-4,c), X.Arr([['NHUOC', 'V_QS', 'V_TAI', 'V_ATG', 'V_AN', 'V_TK', 'TH']]), 0.0), lambda: X.neg(1.0), lambda: X.neg(1.0), lambda: X.neg(1.0), lambda: X.neg(1.0), lambda: (1.0 if X.truthy(X.eq(E.v(8,r-86,c),0.0)) else X.neg(1.0)), lambda: 2.0, lambda: X.neg(1.0))

def f123(E, r, c):
    # Ô C96, D96
    # = C217
    return E.v(8,r+121,c)

def f124(E, r, c):
    # Ô C97, D97
    # = IF(C96<0,1,IF(INDEX(C$209:C$213,C96+1)=1,0,1))
    return (1.0 if X.truthy(X.lt(E.v(8,r-1,c),0.0)) else (0.0 if X.truthy(X.eq(X.INDEX(E.rg(8,209,c,213,c), X.add(E.v(8,r-1,c),1.0)),1.0)) else 1.0))

def f125(E, r, c):
    # Ô C98, D98
    # = IF(C96<0,-1,IF(IF(C7=0,AND(MOD(C96-1,5)>=1,MOD(C96-1,5)<=3),OR(MOD(C96-1,5)=0,MOD(C96-1,5)=4)),MOD(C96+1,5),MOD(C96-1,5)))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(8,r-2,c),0.0)) else (X.MOD(X.add(E.v(8,r-2,c),1.0), 5.0) if X.truthy((X.AND(X.ge(X.MOD(X.sub(E.v(8,r-2,c),1.0), 5.0),1.0), X.le(X.MOD(X.sub(E.v(8,r-2,c),1.0), 5.0),3.0)) if X.truthy(X.eq(E.v(8,r-91,c),0.0)) else X.OR(X.eq(X.MOD(X.sub(E.v(8,r-2,c),1.0), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(8,r-2,c),1.0), 5.0),4.0)))) else X.MOD(X.sub(E.v(8,r-2,c),1.0), 5.0)))

def f126(E, r, c):
    # Ô C99, D99
    # = IF(C96<0,-1,MOD(C96-2,5))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(8,r-3,c),0.0)) else X.MOD(X.sub(E.v(8,r-3,c),2.0), 5.0))

def f127(E, r, c):
    # Ô C102, D102
    # = IF(C72<>"","Cách đặc biệt: "&C72,IF(C7=2,"Trung hòa","Chính cách"))
    return (X.cat('Cách đặc biệt: ',E.v(8,r-30,c)) if X.truthy(X.ne(E.v(8,r-30,c),'')) else ('Trung hòa' if X.truthy(X.eq(E.v(8,r-95,c),2.0)) else 'Chính cách'))

def f128(E, r, c):
    # Ô C103, D103
    # = IF(C72<>"",C74,C96)
    return (E.v(8,r-29,c) if X.truthy(X.ne(E.v(8,r-31,c),'')) else E.v(8,r-7,c))

def f129(E, r, c):
    # Ô C104, D104
    # = IF(C72<>"",C75,IF(C103<0,-1,IF(C7=2,1,C98)))
    return (E.v(8,r-29,c) if X.truthy(X.ne(E.v(8,r-32,c),'')) else (X.neg(1.0) if X.truthy(X.lt(E.v(8,r-1,c),0.0)) else (1.0 if X.truthy(X.eq(E.v(8,r-97,c),2.0)) else E.v(8,r-6,c))))

def f130(E, r, c):
    # Ô C105, D105
    # = IF(C72<>"",C76,IF(C103<0,-1,IF(C7=2,0,C99)))
    return (E.v(8,r-29,c) if X.truthy(X.ne(E.v(8,r-33,c),'')) else (X.neg(1.0) if X.truthy(X.lt(E.v(8,r-2,c),0.0)) else (0.0 if X.truthy(X.eq(E.v(8,r-98,c),2.0)) else E.v(8,r-6,c))))

def f131(E, r, c):
    # Ô C106, D106, C107, D107, C108, D108
    # = IF(C103<0,"—",INDEX(HANH_TEN,MOD(C$5+C103,5)+1)&" ("&INDEX(NHOM_TEN,C103+1)&")")
    return ('—' if X.truthy(X.lt(E.v(8,r-3,c),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),E.v(8,r-3,c)), 5.0),1.0)),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-3,c),1.0))),')'))

def f132(E, r, c):
    # Ô C109, D109, C110, D110, C111, D111
    # = IF(C103<0,-1,MOD(C$5+C103,5))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(8,r-6,c),0.0)) else X.MOD(X.add(E.v(8,5,c),E.v(8,r-6,c)), 5.0))

def f133(E, r, c):
    # Ô C112, D112
    # = IF(OR(C72<>"",C92<0),"—",INDEX(HANH_TEN,MOD(C$5+C92,5)+1)&" ("&INDEX(NHOM_TEN,C92+1)&")")
    return ('—' if X.truthy(X.OR(X.ne(E.v(8,r-40,c),''), X.lt(E.v(8,r-20,c),0.0))) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),E.v(8,r-20,c)), 5.0),1.0)),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-20,c),1.0))),')'))

def f134(E, r, c):
    # Ô C113, D113
    # = IF(OR(C72<>"",C7=2),"—",IF(C7=0,"Quan sát, Tài, Thực thương","Ấn, Tỉ kiếp"))
    return ('—' if X.truthy(X.OR(X.ne(E.v(8,r-41,c),''), X.eq(E.v(8,r-106,c),2.0))) else ('Quan sát, Tài, Thực thương' if X.truthy(X.eq(E.v(8,r-106,c),0.0)) else 'Ấn, Tỉ kiếp'))

def f135(E, r, c):
    # Ô C114, D114
    # = IF(C72<>"","Không",IF(OR(C103<0,C97=1),"Có – trong tứ trụ không có nhóm này (chờ đại vận/lưu niên bổ sung, P1 tr.114-116)","Không"))
    return ('Không' if X.truthy(X.ne(E.v(8,r-42,c),'')) else ('Có – trong tứ trụ không có nhóm này (chờ đại vận/lưu niên bổ sung, P1 tr.114-116)' if X.truthy(X.OR(X.lt(E.v(8,r-11,c),0.0), X.eq(E.v(8,r-17,c),1.0))) else 'Không'))

def f136(E, r, c):
    # Ô C115, D115
    # = IF(C72<>"","Cách "&C72&": dụng thần theo P2 tr.381-382; vận tốt: "&C77&"; vận xấu: "&C78&" (P2 tr.385-386)",CHOOSE(MATCH(C91,{"NHUOC","V_QS","V_TAI","
    #   V_ATG","V_AN","V_TK","TH"},0),"Nhật nhược, "&INDEX(NHOM_TEN,C88+1)&" nhiều → dụng "&INDEX(NHOM_TEN,C93+1)&", không có thì "&INDEX(NHOM_TEN,C94+1)&" (P
    #   1 tr.99, 101-105, 108)","Nhật vượng, quan sát nhiều hơn ấn/tỉ → dụng thực thương (P2 tr.380, 384)","Nhật vượng, tài nhiều hơn ấn/tỉ → dụng quan sát, k
    #   hông có thì thực thương (P2 tr.381, 384)","Nhật vượng, tỉ kiếp nhiều nhưng ấn thấu can có gốc: lấy quan sát thì dụng thần sinh kỵ thần, nên lấy ấn làm
    #    kỵ, dụng tài (P1 tr.106, 190)","Nhật vượng, ấn nhiều → dụng tài; không có tài thì quan sát (chỉ khi ấn không vượng), rồi thực thương (P1 tr.99, 105-1
    #   07, 108)","Nhật vượng, tỉ kiếp nhiều → dụng quan sát, rồi thực thương, rồi tài (P1 tr.99, 107-109, 108)","Trung hòa: sách không có bảng chọn; tham khả
    #   o ""thân và tài ngang nhau thì tài là dụng thần"" (P1 tr.87)"))
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Cách ',E.v(8,r-43,c)),': dụng thần theo P2 tr.381-382; vận tốt: '),E.v(8,r-38,c)),'; vận xấu: '),E.v(8,r-37,c)),' (P2 tr.385-386)') if X.truthy(X.ne(E.v(8,r-43,c),'')) else X.CHOOSE(X.MATCH(E.v(8,r-24,c), X.Arr([['NHUOC', 'V_QS', 'V_TAI', 'V_ATG', 'V_AN', 'V_TK', 'TH']]), 0.0), lambda: X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật nhược, ',X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-27,c),1.0))),' nhiều → dụng '),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-22,c),1.0))),', không có thì '),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-21,c),1.0))),' (P1 tr.99, 101-105, 108)'), lambda: 'Nhật vượng, quan sát nhiều hơn ấn/tỉ → dụng thực thương (P2 tr.380, 384)', lambda: 'Nhật vượng, tài nhiều hơn ấn/tỉ → dụng quan sát, không có thì thực thương (P2 tr.381, 384)', lambda: 'Nhật vượng, tỉ kiếp nhiều nhưng ấn thấu can có gốc: lấy quan sát thì dụng thần sinh kỵ thần, nên lấy ấn làm kỵ, dụng tài (P1 tr.106, 190)', lambda: 'Nhật vượng, ấn nhiều → dụng tài; không có tài thì quan sát (chỉ khi ấn không vượng), rồi thực thương (P1 tr.99, 105-107, 108)', lambda: 'Nhật vượng, tỉ kiếp nhiều → dụng quan sát, rồi thực thương, rồi tài (P1 tr.99, 107-109, 108)', lambda: 'Trung hòa: sách không có bảng chọn; tham khảo "thân và tài ngang nhau thì tài là dụng thần" (P1 tr.87)'))

def f137(E, r, c):
    # Ô C120
    # = IF(TRUNG_GIAN!$C$5<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$C$5+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,5,3),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,5,3),1.0)))

def f138(E, r, c):
    # Ô D120
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$5+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" thấu ở can Năm"&IF(THONG_CAN!$C$71=1," (có gốc)",IF(THONG_CAN!$
    #   C$69="Gốc dễ bị nhổ"," (gốc dễ bị nhổ – chi bị xung)"," (không có gốc)"))
    return X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,5,3),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0))),' thấu ở can Năm'),(' (có gốc)' if X.truthy(X.eq(E.v(11,71,3),1.0)) else (' (gốc dễ bị nhổ – chi bị xung)' if X.truthy(X.eq(E.v(11,69,3),'Gốc dễ bị nhổ')) else ' (không có gốc)')))

def f139(E, r, c):
    # Ô E120
    # = IF(OR(TRUNG_GIAN!$C$5<0,$C$109<0),1000000000,IF(C120<>$C$109,1000000000,IF(THONG_CAN!$C$71=1,0,1)*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN
    #   !$C$5+1)<>TRUNG_GIAN!$E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$5+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+0))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,5,3),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul((0.0 if X.truthy(X.eq(E.v(11,71,3),1.0)) else 1.0),1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),0.0)))

def f140(E, r, c):
    # Ô F120
    # = IF(OR(TRUNG_GIAN!$C$5<0,$D$109<0),1000000000,IF(C120<>$D$109,1000000000,IF(THONG_CAN!$C$71=1,0,1)*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN
    #   !$C$5+1)<>TRUNG_GIAN!$E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$5+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+0))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,5,3),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul((0.0 if X.truthy(X.eq(E.v(11,71,3),1.0)) else 1.0),1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),0.0)))

def f141(E, r, c):
    # Ô C121
    # = IF(TRUNG_GIAN!$D$5<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$D$5+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,5,4),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,5,4),1.0)))

def f142(E, r, c):
    # Ô D121
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$5+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" thấu ở can Tháng"&IF(THONG_CAN!$D$71=1," (có gốc)",IF(THONG_CAN
    #   !$D$69="Gốc dễ bị nhổ"," (gốc dễ bị nhổ – chi bị xung)"," (không có gốc)"))
    return X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,5,4),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0))),' thấu ở can Tháng'),(' (có gốc)' if X.truthy(X.eq(E.v(11,71,4),1.0)) else (' (gốc dễ bị nhổ – chi bị xung)' if X.truthy(X.eq(E.v(11,69,4),'Gốc dễ bị nhổ')) else ' (không có gốc)')))

def f143(E, r, c):
    # Ô E121
    # = IF(OR(TRUNG_GIAN!$D$5<0,$C$109<0),1000000000,IF(C121<>$C$109,1000000000,IF(THONG_CAN!$D$71=1,0,1)*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN
    #   !$D$5+1)<>TRUNG_GIAN!$E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$5+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+3))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,5,4),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul((0.0 if X.truthy(X.eq(E.v(11,71,4),1.0)) else 1.0),1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),3.0)))

def f144(E, r, c):
    # Ô F121
    # = IF(OR(TRUNG_GIAN!$D$5<0,$D$109<0),1000000000,IF(C121<>$D$109,1000000000,IF(THONG_CAN!$D$71=1,0,1)*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN
    #   !$D$5+1)<>TRUNG_GIAN!$E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$5+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+3))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,5,4),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul((0.0 if X.truthy(X.eq(E.v(11,71,4),1.0)) else 1.0),1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),3.0)))

def f145(E, r, c):
    # Ô C122
    # = IF(TRUNG_GIAN!$F$5<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$F$5+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,5,6),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,5,6),1.0)))

def f146(E, r, c):
    # Ô D122
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$5+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" thấu ở can Giờ"&IF(THONG_CAN!$F$71=1," (có gốc)",IF(THONG_CAN!$
    #   F$69="Gốc dễ bị nhổ"," (gốc dễ bị nhổ – chi bị xung)"," (không có gốc)"))
    return X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,5,6),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0))),' thấu ở can Giờ'),(' (có gốc)' if X.truthy(X.eq(E.v(11,71,6),1.0)) else (' (gốc dễ bị nhổ – chi bị xung)' if X.truthy(X.eq(E.v(11,69,6),'Gốc dễ bị nhổ')) else ' (không có gốc)')))

def f147(E, r, c):
    # Ô E122
    # = IF(OR(TRUNG_GIAN!$F$5<0,$C$109<0),1000000000,IF(C122<>$C$109,1000000000,IF(THONG_CAN!$F$71=1,0,1)*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN
    #   !$F$5+1)<>TRUNG_GIAN!$E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$5+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+9))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,5,6),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul((0.0 if X.truthy(X.eq(E.v(11,71,6),1.0)) else 1.0),1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),9.0)))

def f148(E, r, c):
    # Ô F122
    # = IF(OR(TRUNG_GIAN!$F$5<0,$D$109<0),1000000000,IF(C122<>$D$109,1000000000,IF(THONG_CAN!$F$71=1,0,1)*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN
    #   !$F$5+1)<>TRUNG_GIAN!$E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$5+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+9))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,5,6),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul((0.0 if X.truthy(X.eq(E.v(11,71,6),1.0)) else 1.0),1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),9.0)))

def f149(E, r, c):
    # Ô C123
    # = IF(TRUNG_GIAN!$C$15<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$C$15+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,15,3),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,15,3),1.0)))

def f150(E, r, c):
    # Ô D123
    # = IF(TRUNG_GIAN!$C$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$15+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" tàng (bản khí) trong chi Năm "&TRUNG
    #   _GIAN!$C$11)
    return ('' if X.truthy(X.lt(E.v(23,15,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,3),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0))),' tàng (bản khí) trong chi Năm '),E.v(23,11,3)))

def f151(E, r, c):
    # Ô E123
    # = IF(OR(TRUNG_GIAN!$C$15<0,$C$109<0),1000000000,IF(C123<>$C$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$C$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+12))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,3),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),12.0)))

def f152(E, r, c):
    # Ô F123
    # = IF(OR(TRUNG_GIAN!$C$15<0,$D$109<0),1000000000,IF(C123<>$D$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$C$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+12))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,3),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),12.0)))

def f153(E, r, c):
    # Ô C124
    # = IF(TRUNG_GIAN!$C$17<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$C$17+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,17,3),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,17,3),1.0)))

def f154(E, r, c):
    # Ô D124
    # = IF(TRUNG_GIAN!$C$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$17+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" tàng (trung khí) trong chi Năm "&TRU
    #   NG_GIAN!$C$11)
    return ('' if X.truthy(X.lt(E.v(23,17,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,3),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0))),' tàng (trung khí) trong chi Năm '),E.v(23,11,3)))

def f155(E, r, c):
    # Ô E124
    # = IF(OR(TRUNG_GIAN!$C$17<0,$C$109<0),1000000000,IF(C124<>$C$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$C$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+13))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,3),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),13.0)))

def f156(E, r, c):
    # Ô F124
    # = IF(OR(TRUNG_GIAN!$C$17<0,$D$109<0),1000000000,IF(C124<>$D$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$C$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+13))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,3),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),13.0)))

def f157(E, r, c):
    # Ô C125
    # = IF(TRUNG_GIAN!$C$19<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$C$19+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,19,3),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,19,3),1.0)))

def f158(E, r, c):
    # Ô D125
    # = IF(TRUNG_GIAN!$C$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$19+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" tàng (dư khí) trong chi Năm "&TRUNG_
    #   GIAN!$C$11)
    return ('' if X.truthy(X.lt(E.v(23,19,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,3),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0))),' tàng (dư khí) trong chi Năm '),E.v(23,11,3)))

def f159(E, r, c):
    # Ô E125
    # = IF(OR(TRUNG_GIAN!$C$19<0,$C$109<0),1000000000,IF(C125<>$C$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$C$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+14))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,3),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),14.0)))

def f160(E, r, c):
    # Ô F125
    # = IF(OR(TRUNG_GIAN!$C$19<0,$D$109<0),1000000000,IF(C125<>$D$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$C$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$C$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+2*100+14))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,3),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,3),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,3),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(2.0,100.0)),14.0)))

def f161(E, r, c):
    # Ô C126
    # = IF(TRUNG_GIAN!$D$15<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$D$15+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,15,4),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,15,4),1.0)))

def f162(E, r, c):
    # Ô D126
    # = IF(TRUNG_GIAN!$D$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$15+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" tàng (bản khí) trong chi Tháng "&TRU
    #   NG_GIAN!$D$11)
    return ('' if X.truthy(X.lt(E.v(23,15,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,4),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0))),' tàng (bản khí) trong chi Tháng '),E.v(23,11,4)))

def f163(E, r, c):
    # Ô E126
    # = IF(OR(TRUNG_GIAN!$D$15<0,$C$109<0),1000000000,IF(C126<>$C$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$D$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+15))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,4),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),15.0)))

def f164(E, r, c):
    # Ô F126
    # = IF(OR(TRUNG_GIAN!$D$15<0,$D$109<0),1000000000,IF(C126<>$D$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$D$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+15))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,4),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),15.0)))

def f165(E, r, c):
    # Ô C127
    # = IF(TRUNG_GIAN!$D$17<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$D$17+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,17,4),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,17,4),1.0)))

def f166(E, r, c):
    # Ô D127
    # = IF(TRUNG_GIAN!$D$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$17+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" tàng (trung khí) trong chi Tháng "&T
    #   RUNG_GIAN!$D$11)
    return ('' if X.truthy(X.lt(E.v(23,17,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,4),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0))),' tàng (trung khí) trong chi Tháng '),E.v(23,11,4)))

def f167(E, r, c):
    # Ô E127
    # = IF(OR(TRUNG_GIAN!$D$17<0,$C$109<0),1000000000,IF(C127<>$C$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$D$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+16))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,4),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),16.0)))

def f168(E, r, c):
    # Ô F127
    # = IF(OR(TRUNG_GIAN!$D$17<0,$D$109<0),1000000000,IF(C127<>$D$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$D$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+16))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,4),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),16.0)))

def f169(E, r, c):
    # Ô C128
    # = IF(TRUNG_GIAN!$D$19<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$D$19+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,19,4),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,19,4),1.0)))

def f170(E, r, c):
    # Ô D128
    # = IF(TRUNG_GIAN!$D$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$19+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" tàng (dư khí) trong chi Tháng "&TRUN
    #   G_GIAN!$D$11)
    return ('' if X.truthy(X.lt(E.v(23,19,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,4),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0))),' tàng (dư khí) trong chi Tháng '),E.v(23,11,4)))

def f171(E, r, c):
    # Ô E128
    # = IF(OR(TRUNG_GIAN!$D$19<0,$C$109<0),1000000000,IF(C128<>$C$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$D$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+17))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,4),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),17.0)))

def f172(E, r, c):
    # Ô F128
    # = IF(OR(TRUNG_GIAN!$D$19<0,$D$109<0),1000000000,IF(C128<>$D$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$D$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$D$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+17))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,4),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,4),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,4),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),17.0)))

def f173(E, r, c):
    # Ô C129
    # = IF(TRUNG_GIAN!$E$15<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$E$15+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,15,5),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,15,5),1.0)))

def f174(E, r, c):
    # Ô D129
    # = IF(TRUNG_GIAN!$E$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$15+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" tàng (bản khí) trong chi Ngày "&TRUN
    #   G_GIAN!$E$11)
    return ('' if X.truthy(X.lt(E.v(23,15,5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,5),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0))),' tàng (bản khí) trong chi Ngày '),E.v(23,11,5)))

def f175(E, r, c):
    # Ô E129
    # = IF(OR(TRUNG_GIAN!$E$15<0,$C$109<0),1000000000,IF(C129<>$C$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$E$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+0*100+18))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,5),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,5),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(0.0,100.0)),18.0)))

def f176(E, r, c):
    # Ô F129
    # = IF(OR(TRUNG_GIAN!$E$15<0,$D$109<0),1000000000,IF(C129<>$D$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$E$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+0*100+18))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,5),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,5),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(0.0,100.0)),18.0)))

def f177(E, r, c):
    # Ô C130
    # = IF(TRUNG_GIAN!$E$17<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$E$17+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,17,5),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,17,5),1.0)))

def f178(E, r, c):
    # Ô D130
    # = IF(TRUNG_GIAN!$E$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$17+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" tàng (trung khí) trong chi Ngày "&TR
    #   UNG_GIAN!$E$11)
    return ('' if X.truthy(X.lt(E.v(23,17,5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,5),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0))),' tàng (trung khí) trong chi Ngày '),E.v(23,11,5)))

def f179(E, r, c):
    # Ô E130
    # = IF(OR(TRUNG_GIAN!$E$17<0,$C$109<0),1000000000,IF(C130<>$C$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$E$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+0*100+19))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,5),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,5),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(0.0,100.0)),19.0)))

def f180(E, r, c):
    # Ô F130
    # = IF(OR(TRUNG_GIAN!$E$17<0,$D$109<0),1000000000,IF(C130<>$D$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$E$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+0*100+19))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,5),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,5),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(0.0,100.0)),19.0)))

def f181(E, r, c):
    # Ô C131
    # = IF(TRUNG_GIAN!$E$19<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$E$19+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,19,5),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,19,5),1.0)))

def f182(E, r, c):
    # Ô D131
    # = IF(TRUNG_GIAN!$E$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$19+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" tàng (dư khí) trong chi Ngày "&TRUNG
    #   _GIAN!$E$11)
    return ('' if X.truthy(X.lt(E.v(23,19,5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,5),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0))),' tàng (dư khí) trong chi Ngày '),E.v(23,11,5)))

def f183(E, r, c):
    # Ô E131
    # = IF(OR(TRUNG_GIAN!$E$19<0,$C$109<0),1000000000,IF(C131<>$C$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$E$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+0*100+20))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,5),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,5),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(0.0,100.0)),20.0)))

def f184(E, r, c):
    # Ô F131
    # = IF(OR(TRUNG_GIAN!$E$19<0,$D$109<0),1000000000,IF(C131<>$D$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$E$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+0*100+20))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,5),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,5),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,5),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(0.0,100.0)),20.0)))

def f185(E, r, c):
    # Ô C132
    # = IF(TRUNG_GIAN!$F$15<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$F$15+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,15,6),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,15,6),1.0)))

def f186(E, r, c):
    # Ô D132
    # = IF(TRUNG_GIAN!$F$15<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$15+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" tàng (bản khí) trong chi Giờ "&TRUNG
    #   _GIAN!$F$11)
    return ('' if X.truthy(X.lt(E.v(23,15,6),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,15,6),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0))),' tàng (bản khí) trong chi Giờ '),E.v(23,11,6)))

def f187(E, r, c):
    # Ô E132
    # = IF(OR(TRUNG_GIAN!$F$15<0,$C$109<0),1000000000,IF(C132<>$C$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$F$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+21))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,6),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),21.0)))

def f188(E, r, c):
    # Ô F132
    # = IF(OR(TRUNG_GIAN!$F$15<0,$D$109<0),1000000000,IF(C132<>$D$109,1000000000,2*1000000+0*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$F$15+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$15+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+21))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,15,6),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(0.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,15,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,15,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),21.0)))

def f189(E, r, c):
    # Ô C133
    # = IF(TRUNG_GIAN!$F$17<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$F$17+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,17,6),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,17,6),1.0)))

def f190(E, r, c):
    # Ô D133
    # = IF(TRUNG_GIAN!$F$17<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$17+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" tàng (trung khí) trong chi Giờ "&TRU
    #   NG_GIAN!$F$11)
    return ('' if X.truthy(X.lt(E.v(23,17,6),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,17,6),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0))),' tàng (trung khí) trong chi Giờ '),E.v(23,11,6)))

def f191(E, r, c):
    # Ô E133
    # = IF(OR(TRUNG_GIAN!$F$17<0,$C$109<0),1000000000,IF(C133<>$C$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$F$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+22))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,6),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),22.0)))

def f192(E, r, c):
    # Ô F133
    # = IF(OR(TRUNG_GIAN!$F$17<0,$D$109<0),1000000000,IF(C133<>$D$109,1000000000,2*1000000+1*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$F$17+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$17+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+22))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,17,6),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(1.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,17,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,17,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),22.0)))

def f193(E, r, c):
    # Ô C134
    # = IF(TRUNG_GIAN!$F$19<0,-1,INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$F$19+1))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(23,19,6),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,19,6),1.0)))

def f194(E, r, c):
    # Ô D134
    # = IF(TRUNG_GIAN!$F$19<0,"",INDEX(M_TT,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$19+1)&" "&INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" tàng (dư khí) trong chi Giờ "&TRUNG_
    #   GIAN!$F$11)
    return ('' if X.truthy(X.lt(E.v(23,19,6),0.0)) else X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(23,19,6),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0))),' tàng (dư khí) trong chi Giờ '),E.v(23,11,6)))

def f195(E, r, c):
    # Ô E134
    # = IF(OR(TRUNG_GIAN!$F$19<0,$C$109<0),1000000000,IF(C134<>$C$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$F$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+23))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,6),0.0), X.lt(E.v(8,109,3),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-2),E.v(8,109,3))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),23.0)))

def f196(E, r, c):
    # Ô F134
    # = IF(OR(TRUNG_GIAN!$F$19<0,$D$109<0),1000000000,IF(C134<>$D$109,1000000000,2*1000000+2*100000+IF(INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$F$19+1)<>TRUNG_GIAN!$
    #   E$8,0,1)*10000+IF(INDEX(M_TS,TRUNG_GIAN!$F$19+1,TRUNG_GIAN!$D$6+1)<=4,0,1)*1000+1*100+23))
    return (1000000000.0 if X.truthy(X.OR(X.lt(E.v(23,19,6),0.0), X.lt(E.v(8,109,4),0.0))) else (1000000000.0 if X.truthy(X.ne(E.v(8,r,c-3),E.v(8,109,4))) else X.add(X.add(X.add(X.add(X.add(X.mul(2.0,1000000.0),X.mul(2.0,100000.0)),X.mul((0.0 if X.truthy(X.ne(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,19,6),1.0)),E.v(23,8,5))) else 1.0),10000.0)),X.mul((0.0 if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,19,6),1.0), X.add(E.v(23,6,4),1.0)),4.0)) else 1.0),1000.0)),X.mul(1.0,100.0)),23.0)))

def f197(E, r, c):
    # Ô C136
    # = IF(MIN($E$120:$E$134)>=1000000000,"(không có trong tứ trụ)",INDEX($D$120:$D$134,MATCH(MIN($E$120:$E$134),$E$120:$E$134,0)))
    return ('(không có trong tứ trụ)' if X.truthy(X.ge(X.MIN(E.rg(8,120,5,134,5)),1000000000.0)) else X.INDEX(E.rg(8,120,4,134,4), X.MATCH(X.MIN(E.rg(8,120,5,134,5)), E.rg(8,120,5,134,5), 0.0)))

def f198(E, r, c):
    # Ô C137
    # = IF(MIN($F$120:$F$134)>=1000000000,"(không có trong tứ trụ)",INDEX($D$120:$D$134,MATCH(MIN($F$120:$F$134),$F$120:$F$134,0)))
    return ('(không có trong tứ trụ)' if X.truthy(X.ge(X.MIN(E.rg(8,120,6,134,6)),1000000000.0)) else X.INDEX(E.rg(8,120,4,134,4), X.MATCH(X.MIN(E.rg(8,120,6,134,6)), E.rg(8,120,6,134,6), 0.0)))

def f199(E, r, c):
    # Ô C140
    # = $C$171
    return E.v(8,171,3)

def f200(E, r, c):
    # Ô C141
    # = IF(AND($C$14=$C$16,$C$14>=2,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD($C$5+2,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0,SUMPRODUCT((TRUNG
    #   _GIAN!$C$7:$F$7=MOD($C$5+4,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0),"Tài và Ấn cùng vượng, cùng tương tranh → cần xét thông quan 
    #   bằng Quan sát (P1 tr.100, 110-111). Độ PP2: Tài "&TEXT($D$14,"0.0")&", Ấn "&TEXT($D$16,"0.0")&IF($C$21+$D$21>0," → Tứ trụ CÓ Quan sát: thông quan được
    #   ."," → KHÔNG có Quan sát: mệnh cục trục trặc, chờ vận Quan sát."),"")
    return (X.cat(X.cat(X.cat(X.cat('Tài và Ấn cùng vượng, cùng tương tranh → cần xét thông quan bằng Quan sát (P1 tr.100, 110-111). Độ PP2: Tài ',X.TEXT(E.v(8,14,4), '0.0')),', Ấn '),X.TEXT(E.v(8,16,4), '0.0')),(' → Tứ trụ CÓ Quan sát: thông quan được.' if X.truthy(X.gt(X.add(E.v(8,21,3),E.v(8,21,4)),0.0)) else ' → KHÔNG có Quan sát: mệnh cục trục trặc, chờ vận Quan sát.')) if X.truthy(X.AND(X.eq(E.v(8,14,3),E.v(8,16,3)), X.ge(E.v(8,14,3),2.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),2.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),4.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0))) else '')

def f201(E, r, c):
    # Ô C142
    # = IF(AND($C$15=$C$12,$C$15>=2,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD($C$5+3,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0,SUMPRODUCT((TRUNG
    #   _GIAN!$C$7:$F$7=MOD($C$5+0,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0),"Quan sát và Tỉ kiếp cùng vượng, cùng tương tranh → cần xét t
    #   hông quan bằng Ấn (P1 tr.100, 110-111). Độ PP2: Quan sát "&TEXT($D$15,"0.0")&", Tỉ kiếp "&TEXT($D$12,"0.0")&IF($C$22+$D$22>0," → Tứ trụ CÓ Ấn: thông q
    #   uan được."," → KHÔNG có Ấn: mệnh cục trục trặc, chờ vận Ấn."),"")
    return (X.cat(X.cat(X.cat(X.cat('Quan sát và Tỉ kiếp cùng vượng, cùng tương tranh → cần xét thông quan bằng Ấn (P1 tr.100, 110-111). Độ PP2: Quan sát ',X.TEXT(E.v(8,15,4), '0.0')),', Tỉ kiếp '),X.TEXT(E.v(8,12,4), '0.0')),(' → Tứ trụ CÓ Ấn: thông quan được.' if X.truthy(X.gt(X.add(E.v(8,22,3),E.v(8,22,4)),0.0)) else ' → KHÔNG có Ấn: mệnh cục trục trặc, chờ vận Ấn.')) if X.truthy(X.AND(X.eq(E.v(8,15,3),E.v(8,12,3)), X.ge(E.v(8,15,3),2.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),3.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),0.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0))) else '')

def f202(E, r, c):
    # Ô C143
    # = IF(AND($C$12=$C$14,$C$12>=2,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD($C$5+0,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0,SUMPRODUCT((TRUNG
    #   _GIAN!$C$7:$F$7=MOD($C$5+2,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0),"Tỉ kiếp và Tài cùng vượng, cùng tương tranh → cần xét thông 
    #   quan bằng Thực thương (P1 tr.100, 110-111). Độ PP2: Tỉ kiếp "&TEXT($D$12,"0.0")&", Tài "&TEXT($D$14,"0.0")&IF($C$155=1,IF($C$19+$D$19>0," → Nhật can v
    #   ượng, CÓ Thực thương: dùng Thực thương thông quan làm dụng thần."," → Nhật can vượng nhưng KHÔNG có Thực thương: chờ vận Thực thương.")," → Nhật can n
    #   hược: lấy Thực thương thông quan KHÔNG có tác dụng như dụng thần (P1 tr.111)."),"")
    return (X.cat(X.cat(X.cat(X.cat('Tỉ kiếp và Tài cùng vượng, cùng tương tranh → cần xét thông quan bằng Thực thương (P1 tr.100, 110-111). Độ PP2: Tỉ kiếp ',X.TEXT(E.v(8,12,4), '0.0')),', Tài '),X.TEXT(E.v(8,14,4), '0.0')),((' → Nhật can vượng, CÓ Thực thương: dùng Thực thương thông quan làm dụng thần.' if X.truthy(X.gt(X.add(E.v(8,19,3),E.v(8,19,4)),0.0)) else ' → Nhật can vượng nhưng KHÔNG có Thực thương: chờ vận Thực thương.') if X.truthy(X.eq(E.v(8,155,3),1.0)) else ' → Nhật can nhược: lấy Thực thương thông quan KHÔNG có tác dụng như dụng thần (P1 tr.111).')) if X.truthy(X.AND(X.eq(E.v(8,12,3),E.v(8,14,3)), X.ge(E.v(8,12,3),2.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),0.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),2.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0))) else '')

def f203(E, r, c):
    # Ô C144
    # = IF(AND($C$16=$C$13,$C$16>=2,SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD($C$5+4,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0,SUMPRODUCT((TRUNG
    #   _GIAN!$C$7:$F$7=MOD($C$5+1,5))*(TRUNG_GIAN!$C$30:$F$30<=4)*(TRUNG_GIAN!$C$4:$F$4<>2))>0),"Ấn và Thực thương cùng vượng, cùng tương tranh → cần xét thô
    #   ng quan bằng nhật can được lệnh (P1 tr.100, 110-111). Độ PP2: Ấn "&TEXT($D$16,"0.0")&", Thực thương "&TEXT($D$13,"0.0")&IF(TRUNG_GIAN!$D$28<=4," → Nhậ
    #   t can được lệnh: thông quan được."," → Nhật can KHÔNG được lệnh: không thông quan được."),"")
    return (X.cat(X.cat(X.cat(X.cat('Ấn và Thực thương cùng vượng, cùng tương tranh → cần xét thông quan bằng nhật can được lệnh (P1 tr.100, 110-111). Độ PP2: Ấn ',X.TEXT(E.v(8,16,4), '0.0')),', Thực thương '),X.TEXT(E.v(8,13,4), '0.0')),(' → Nhật can được lệnh: thông quan được.' if X.truthy(X.le(E.v(23,28,4),4.0)) else ' → Nhật can KHÔNG được lệnh: không thông quan được.')) if X.truthy(X.AND(X.eq(E.v(8,16,3),E.v(8,13,3)), X.ge(E.v(8,16,3),2.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),4.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0), X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,3),1.0), 5.0)),X.A('le',E.rg(23,30,3,30,6),4.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0))) else '')

def f204(E, r, c):
    # Ô C145
    # = IF(AND(C72="",OR(C91="V_TK",C91="V_ATG")),"PP1: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương ("&INDEX(HANH_TEN,MOD(C$5
    #   +1,5)+1)&") nếu thực thương có lực; lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"th
    #   ực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương lộ can nhưng không có
    #    gốc – lực yếu",IF(INDEX(C$18:C$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực thương")))&". Công cụ vẫn chọn theo Thiệu Vĩ 
    #   Hoa (P1 tr.99, 106).",IF(AND(C61<>"",C72=C61),"PP1: cách "&C61&" – Trích Thiên Tủy ch.13: có thể dụng thực thương ("&INDEX(HANH_TEN,MOD(C$5+1,5)+1)&")
    #    làm tú khí (không dụng quan sát); lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thự
    #   c thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương lộ can nhưng không có 
    #   gốc – lực yếu",IF(INDEX(C$18:C$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực thương")))&". Công cụ vẫn theo P2 tr.369-373."
    #   ,""))&IF(AND(IF(AND(C72="",OR(C91="V_TK",C91="V_ATG")),"PP1: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương ("&INDEX(HAN
    #   H_TEN,MOD(C$5+1,5)+1)&") nếu thực thương có lực; lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F
    #   $4<>2))>0,"thực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương lộ can n
    #   hưng không có gốc – lực yếu",IF(INDEX(C$18:C$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực thương")))&". Công cụ vẫn chọn t
    #   heo Thiệu Vĩ Hoa (P1 tr.99, 106).",IF(AND(C61<>"",C72=C61),"PP1: cách "&C61&" – Trích Thiên Tủy ch.13: có thể dụng thực thương ("&INDEX(HANH_TEN,MOD(C
    #   $5+1,5)+1)&") làm tú khí (không dụng quan sát); lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$
    #   4<>2))>0,"thực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương lộ can nh
    #   ưng không có gốc – lực yếu",IF(INDEX(C$18:C$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực thương")))&". Công cụ vẫn theo P2
    #    tr.369-373.",""))<>"",IF(AND(D72="",OR(D91="V_TK",D91="V_ATG")),"PP2: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương ("
    #   &INDEX(HANH_TEN,MOD(D$5+1,5)+1)&") nếu thực thương có lực; lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GI
    #   AN!$C$4:$F$4<>2))>0,"thực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thươn
    #   g lộ can nhưng không có gốc – lực yếu",IF(INDEX(D$18:D$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực thương")))&". Công cụ 
    #   vẫn chọn theo Thiệu Vĩ Hoa (P1 tr.99, 106).",IF(AND(D61<>"",D72=D61),"PP2: cách "&D61&" – Trích Thiên Tủy ch.13: có thể dụng thực thương ("&INDEX(HANH
    #   _TEN,MOD(D$5+1,5)+1)&") làm tú khí (không dụng quan sát); lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*THONG_CAN!$C$71:$F$71*(TRUNG_GIA
    #   N!$C$4:$F$4<>2))>0,"thực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương
    #    lộ can nhưng không có gốc – lực yếu",IF(INDEX(D$18:D$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực thương")))&". Công cụ v
    #   ẫn theo P2 tr.369-373.",""))<>""),CHAR(10),"")&IF(AND(D72="",OR(D91="V_TK",D91="V_ATG")),"PP2: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có
    #    thể dụng thực thương ("&INDEX(HANH_TEN,MOD(D$5+1,5)+1)&") nếu thực thương có lực; lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*THONG_C
    #   AN!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*(TRUNG_GIAN!$C$4
    #   :$F$4<>2))>0,"thực thương lộ can nhưng không có gốc – lực yếu",IF(INDEX(D$18:D$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thự
    #   c thương")))&". Công cụ vẫn chọn theo Thiệu Vĩ Hoa (P1 tr.99, 106).",IF(AND(D61<>"",D72=D61),"PP2: cách "&D61&" – Trích Thiên Tủy ch.13: có thể dụng t
    #   hực thương ("&INDEX(HANH_TEN,MOD(D$5+1,5)+1)&") làm tú khí (không dụng quan sát); lá này: "&IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*THONG_CA
    #   N!$C$71:$F$71*(TRUNG_GIAN!$C$4:$F$4<>2))>0,"thực thương lộ can và có gốc – có lực",IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(D$5+1,5))*(TRUNG_GIAN!$C$4:
    #   $F$4<>2))>0,"thực thương lộ can nhưng không có gốc – lực yếu",IF(INDEX(D$18:D$22,2)=1,"thực thương chỉ tàng trong chi – lực yếu","tứ trụ không có thực
    #    thương")))&". Công cụ vẫn theo P2 tr.369-373.",""))
    return X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('PP1: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),1.0), 5.0),1.0))),') nếu thực thương có lực; lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c,22,c), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn chọn theo Thiệu Vĩ Hoa (P1 tr.99, 106).') if X.truthy(X.AND(X.eq(E.v(8,r-73,c),''), X.OR(X.eq(E.v(8,r-54,c),'V_TK'), X.eq(E.v(8,r-54,c),'V_ATG')))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: cách ',E.v(8,r-84,c)),' – Trích Thiên Tủy ch.13: có thể dụng thực thương ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),1.0), 5.0),1.0))),') làm tú khí (không dụng quan sát); lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c,22,c), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn theo P2 tr.369-373.') if X.truthy(X.AND(X.ne(E.v(8,r-84,c),''), X.eq(E.v(8,r-73,c),E.v(8,r-84,c)))) else '')),(X.CHAR(10.0) if X.truthy(X.AND(X.ne((X.cat(X.cat(X.cat(X.cat('PP1: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),1.0), 5.0),1.0))),') nếu thực thương có lực; lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c,22,c), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn chọn theo Thiệu Vĩ Hoa (P1 tr.99, 106).') if X.truthy(X.AND(X.eq(E.v(8,r-73,c),''), X.OR(X.eq(E.v(8,r-54,c),'V_TK'), X.eq(E.v(8,r-54,c),'V_ATG')))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: cách ',E.v(8,r-84,c)),' – Trích Thiên Tủy ch.13: có thể dụng thực thương ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),1.0), 5.0),1.0))),') làm tú khí (không dụng quan sát); lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c,22,c), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn theo P2 tr.369-373.') if X.truthy(X.AND(X.ne(E.v(8,r-84,c),''), X.eq(E.v(8,r-73,c),E.v(8,r-84,c)))) else '')),''), X.ne((X.cat(X.cat(X.cat(X.cat('PP2: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c+1),1.0), 5.0),1.0))),') nếu thực thương có lực; lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c+1,22,c+1), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn chọn theo Thiệu Vĩ Hoa (P1 tr.99, 106).') if X.truthy(X.AND(X.eq(E.v(8,r-73,c+1),''), X.OR(X.eq(E.v(8,r-54,c+1),'V_TK'), X.eq(E.v(8,r-54,c+1),'V_ATG')))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: cách ',E.v(8,r-84,c+1)),' – Trích Thiên Tủy ch.13: có thể dụng thực thương ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c+1),1.0), 5.0),1.0))),') làm tú khí (không dụng quan sát); lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c+1,22,c+1), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn theo P2 tr.369-373.') if X.truthy(X.AND(X.ne(E.v(8,r-84,c+1),''), X.eq(E.v(8,r-73,c+1),E.v(8,r-84,c+1)))) else '')),''))) else '')),(X.cat(X.cat(X.cat(X.cat('PP2: nhật vượng do tỉ kiếp nhiều – Trích Thiên Tủy ch.13: có thể dụng thực thương (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c+1),1.0), 5.0),1.0))),') nếu thực thương có lực; lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c+1,22,c+1), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn chọn theo Thiệu Vĩ Hoa (P1 tr.99, 106).') if X.truthy(X.AND(X.eq(E.v(8,r-73,c+1),''), X.OR(X.eq(E.v(8,r-54,c+1),'V_TK'), X.eq(E.v(8,r-54,c+1),'V_ATG')))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: cách ',E.v(8,r-84,c+1)),' – Trích Thiên Tủy ch.13: có thể dụng thực thương ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c+1),1.0), 5.0),1.0))),') làm tú khí (không dụng quan sát); lá này: '),('thực thương lộ can và có gốc – có lực' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),E.rg(11,71,3,71,6)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương lộ can nhưng không có gốc – lực yếu' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c+1),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),0.0)) else ('thực thương chỉ tàng trong chi – lực yếu' if X.truthy(X.eq(X.INDEX(E.rg(8,18,c+1,22,c+1), 2.0),1.0)) else 'tứ trụ không có thực thương')))),'. Công cụ vẫn theo P2 tr.369-373.') if X.truthy(X.AND(X.ne(E.v(8,r-84,c+1),''), X.eq(E.v(8,r-73,c+1),E.v(8,r-84,c+1)))) else '')))

def f205(E, r, c):
    # Ô C147
    # = "PP1: "&IF(C102="Chính cách","Nhật chủ "&CHOOSE(C7+1,"nhược","vượng","trung hòa")&" → "&IF(C7=0,"cần SINH – TRỢ (bổ cho nhật chủ)",IF(C7=1,"cần KHẮC –
    #    TIẾT – HAO (bớt cho nhật chủ)","giữ cân bằng"))&". Phe gây mất cân bằng: "&C113&"; nhóm nhiều nhất: "&C112&" (""lấy cái nhiều làm kỵ"" – P1 tr.180). 
    #   Quy tắc: "&C115&" → DỤNG THẦN "&C106&". ","Lá số thuộc "&C102&": chọn dụng theo thế của cách (thuận thế), không theo phù ức (P2 tr.369-394) → DỤNG THẦ
    #   N "&C106&". ")&"HỶ THẦN "&C107&IF(C110<0,"",IF(C110=MOD(C109+4,5)," = hành sinh ra dụng thần (P2 tr.382)."," – hành sinh dụng thần thuộc phe gây mất c
    #   ân bằng nên lấy hành khắc kỵ thần (P2 tr.382; P1 tr.112)."))&" KỴ THẦN "&C108&" = hành khắc dụng thần (P2 tr.382)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',X.CHOOSE(X.add(E.v(8,r-140,c),1.0), lambda: 'nhược', lambda: 'vượng', lambda: 'trung hòa')),' → '),('cần SINH – TRỢ (bổ cho nhật chủ)' if X.truthy(X.eq(E.v(8,r-140,c),0.0)) else ('cần KHẮC – TIẾT – HAO (bớt cho nhật chủ)' if X.truthy(X.eq(E.v(8,r-140,c),1.0)) else 'giữ cân bằng'))),'. Phe gây mất cân bằng: '),E.v(8,r-34,c)),'; nhóm nhiều nhất: '),E.v(8,r-35,c)),' ("lấy cái nhiều làm kỵ" – P1 tr.180). Quy tắc: '),E.v(8,r-32,c)),' → DỤNG THẦN '),E.v(8,r-41,c)),'. ') if X.truthy(X.eq(E.v(8,r-45,c),'Chính cách')) else X.cat(X.cat(X.cat(X.cat('Lá số thuộc ',E.v(8,r-45,c)),': chọn dụng theo thế của cách (thuận thế), không theo phù ức (P2 tr.369-394) → DỤNG THẦN '),E.v(8,r-41,c)),'. '))),'HỶ THẦN '),E.v(8,r-40,c)),('' if X.truthy(X.lt(E.v(8,r-37,c),0.0)) else (' = hành sinh ra dụng thần (P2 tr.382).' if X.truthy(X.eq(E.v(8,r-37,c),X.MOD(X.add(E.v(8,r-38,c),4.0), 5.0))) else ' – hành sinh dụng thần thuộc phe gây mất cân bằng nên lấy hành khắc kỵ thần (P2 tr.382; P1 tr.112).'))),' KỴ THẦN '),E.v(8,r-39,c)),' = hành khắc dụng thần (P2 tr.382).')

def f206(E, r, c):
    # Ô C148
    # = "PP2: "&IF(D102="Chính cách","Nhật chủ "&CHOOSE(D7+1,"nhược","vượng","trung hòa")&" → "&IF(D7=0,"cần SINH – TRỢ (bổ cho nhật chủ)",IF(D7=1,"cần KHẮC –
    #    TIẾT – HAO (bớt cho nhật chủ)","giữ cân bằng"))&". Phe gây mất cân bằng: "&D113&"; nhóm nhiều nhất: "&D112&" (""lấy cái nhiều làm kỵ"" – P1 tr.180). 
    #   Quy tắc: "&D115&" → DỤNG THẦN "&D106&". ","Lá số thuộc "&D102&": chọn dụng theo thế của cách (thuận thế), không theo phù ức (P2 tr.369-394) → DỤNG THẦ
    #   N "&D106&". ")&"HỶ THẦN "&D107&IF(D110<0,"",IF(D110=MOD(D109+4,5)," = hành sinh ra dụng thần (P2 tr.382)."," – hành sinh dụng thần thuộc phe gây mất c
    #   ân bằng nên lấy hành khắc kỵ thần (P2 tr.382; P1 tr.112)."))&" KỴ THẦN "&D108&" = hành khắc dụng thần (P2 tr.382)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: ',(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',X.CHOOSE(X.add(E.v(8,r-141,c+1),1.0), lambda: 'nhược', lambda: 'vượng', lambda: 'trung hòa')),' → '),('cần SINH – TRỢ (bổ cho nhật chủ)' if X.truthy(X.eq(E.v(8,r-141,c+1),0.0)) else ('cần KHẮC – TIẾT – HAO (bớt cho nhật chủ)' if X.truthy(X.eq(E.v(8,r-141,c+1),1.0)) else 'giữ cân bằng'))),'. Phe gây mất cân bằng: '),E.v(8,r-35,c+1)),'; nhóm nhiều nhất: '),E.v(8,r-36,c+1)),' ("lấy cái nhiều làm kỵ" – P1 tr.180). Quy tắc: '),E.v(8,r-33,c+1)),' → DỤNG THẦN '),E.v(8,r-42,c+1)),'. ') if X.truthy(X.eq(E.v(8,r-46,c+1),'Chính cách')) else X.cat(X.cat(X.cat(X.cat('Lá số thuộc ',E.v(8,r-46,c+1)),': chọn dụng theo thế của cách (thuận thế), không theo phù ức (P2 tr.369-394) → DỤNG THẦN '),E.v(8,r-42,c+1)),'. '))),'HỶ THẦN '),E.v(8,r-41,c+1)),('' if X.truthy(X.lt(E.v(8,r-38,c+1),0.0)) else (' = hành sinh ra dụng thần (P2 tr.382).' if X.truthy(X.eq(E.v(8,r-38,c+1),X.MOD(X.add(E.v(8,r-39,c+1),4.0), 5.0))) else ' – hành sinh dụng thần thuộc phe gây mất cân bằng nên lấy hành khắc kỵ thần (P2 tr.382; P1 tr.112).'))),' KỴ THẦN '),E.v(8,r-40,c+1)),' = hành khắc dụng thần (P2 tr.382).')

def f207(E, r, c):
    # Ô C149
    # = "Dùng theo "&DOI_CHIEU!$D$33&": nên bồi bổ hành "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1)&IF(IF(DOI_CHIEU!$C$33=2,
    #   DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)>=0," và "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1),"")&" (màu sắc, phương vị, 
    #   nghề nghiệp; năm – vận mang hành này là thuận), hạn chế hành "&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,INDEX(HANH_TEN,IF(DOI_C
    #   HIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1),"—")&" (năm – vận mang hành này dễ trắc trở). Chữ đại diện của dụng thần trong tứ trụ: "&IF(DOI_C
    #   HIEU!$C$33=2,C137,C136)&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dùng theo ',E.v(7,33,4)),': nên bồi bổ hành '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),(X.cat(' và ',X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))) if X.truthy(X.ge((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else '')),' (màu sắc, phương vị, nghề nghiệp; năm – vận mang hành này là thuận), hạn chế hành '),(X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0)) if X.truthy(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0)) else '—')),' (năm – vận mang hành này dễ trắc trở). Chữ đại diện của dụng thần trong tứ trụ: '),(E.v(8,r-12,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,r-13,c))),'.')

def f208(E, r, c):
    # Ô C151
    # = F151&IF(AND(F151<>"",G151<>""),CHAR(10),"")&G151&IF(F151&G151<>""," Điểm mấu chốt là xác định lại thân vượng hay nhược (P1 tr.114). [Ngưỡng 40%: sách 
    #   không cho số, công cụ lấy theo thang phe mình ≥ 40% là vượng của P2 tr.346. Chỉ cảnh báo, không đổi dụng thần.]","")
    return X.cat(X.cat(X.cat(E.v(8,r,c+3),(X.CHAR(10.0) if X.truthy(X.AND(X.ne(E.v(8,r,c+3),''), X.ne(E.v(8,r,c+4),''))) else '')),E.v(8,r,c+4)),(' Điểm mấu chốt là xác định lại thân vượng hay nhược (P1 tr.114). [Ngưỡng 40%: sách không cho số, công cụ lấy theo thang phe mình ≥ 40% là vượng của P2 tr.346. Chỉ cảnh báo, không đổi dụng thần.]' if X.truthy(X.ne(X.cat(E.v(8,r,c+3),E.v(8,r,c+4)),'')) else ''))

def f209(E, r, c):
    # Ô F151
    # = IF(AND(C72="",C7<>2,C103>=0,C97=0,PP2_DINH_LUONG!$C$100>0),IF(INDEX(PP2_DINH_LUONG!$F$91:$F$95,C109+1)/PP2_DINH_LUONG!$C$100>=0.4,IF(AND(C7=0,C103=4),
    #   "PP1: Ấn đang thái quá ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG!$F$91:$F$95,C109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.113: nhật nhược đ
    #   ược sinh phù thái quá thì có thể biến thành vượng; khi đó chọn dụng như thân vượng, ấn vượng: dụng tài để chế ấn (P1 tr.106, 109). Đối chiếu PP2."&" →
    #    Nên xem thêm "&INDEX(HANH_TEN,MOD(C$5+2,5)+1)&" (tài) làm dụng thần thay thế; "&IF(C14>0.000000001,"tứ trụ có "&IF(ABS(C14-ROUND(C14,0))<0.000000001,
    #   TEXT(C14,"0"),SUBSTITUTE(TEXT(C14,"0.0"),".",","))&" chữ.","tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97)."),IF(AND(C7=0,C103=0),"PP1: Tỉ kiếp
    #    đang thái quá ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG!$F$91:$F$95,C109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.113: phù trợ mạnh quá thì
    #    hành chế áp sự phù trợ là dụng; nhật nhược thành vượng thì chọn như nhật vượng, tỉ kiếp vượng: lấy quan sát làm dụng, không cần quan sát thật mạnh. Đ
    #   ối chiếu PP2."&" → Nên xem thêm "&INDEX(HANH_TEN,MOD(C$5+3,5)+1)&" (quan sát) làm dụng thần thay thế; "&IF(C15>0.000000001,"tứ trụ có "&IF(ABS(C15-ROU
    #   ND(C15,0))<0.000000001,TEXT(C15,"0"),SUBSTITUTE(TEXT(C15,"0.0"),".",","))&" chữ.","tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97)."),IF(AND(C7=
    #   1,C103=3),"PP1: Quan sát đang thái quá ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG!$F$91:$F$95,C109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.1
    #   13–114: chế áp mạnh quá mà không có ấn thông quan thì dụng thực thương để khống chế quan sát; thực thương không được quá mạnh vì còn làm xì hơi thân."
    #   &" → Nên xem thêm "&INDEX(HANH_TEN,MOD(C$5+1,5)+1)&" (thực thương) làm dụng thần thay thế; "&IF(C13>0.000000001,"tứ trụ có "&IF(ABS(C13-ROUND(C13,0))<
    #   0.000000001,TEXT(C13,"0"),SUBSTITUTE(TEXT(C13,"0.0"),".",","))&" chữ.","tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).")&IF(C16>0.000000001," 
    #   Tứ trụ có ấn: xét thông quan trước (P1 tr.114).",""),IF(AND(C7=1,C103=2,OR(C15>0,C13>0)),"PP1: Tài đang mạnh nhất ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LU
    #   ONG!$F$91:$F$95,C109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.114: tài dụng để sinh quan không nên vượng quá; nếu mạnh quá thì thân chuyển
    #    nhược, không thắng nổi tài quan.","")))),""),"")
    return (((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: Ấn đang thái quá (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.113: nhật nhược được sinh phù thái quá thì có thể biến thành vượng; khi đó chọn dụng như thân vượng, ấn vượng: dụng tài để chế ấn (P1 tr.106, 109). Đối chiếu PP2.'),' → Nên xem thêm '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c-3),2.0), 5.0),1.0))),' (tài) làm dụng thần thay thế; '),(X.cat(X.cat('tứ trụ có ',(X.TEXT(E.v(8,r-137,c-3), '0') if X.truthy(X.lt(X.ABS(X.sub(E.v(8,r-137,c-3),X.ROUND(E.v(8,r-137,c-3), 0.0))),1e-09)) else X.SUBSTITUTE(X.TEXT(E.v(8,r-137,c-3), '0.0'), '.', ','))),' chữ.') if X.truthy(X.gt(E.v(8,r-137,c-3),1e-09)) else 'tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).')) if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),0.0), X.eq(E.v(8,r-48,c-3),4.0))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: Tỉ kiếp đang thái quá (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.113: phù trợ mạnh quá thì hành chế áp sự phù trợ là dụng; nhật nhược thành vượng thì chọn như nhật vượng, tỉ kiếp vượng: lấy quan sát làm dụng, không cần quan sát thật mạnh. Đối chiếu PP2.'),' → Nên xem thêm '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c-3),3.0), 5.0),1.0))),' (quan sát) làm dụng thần thay thế; '),(X.cat(X.cat('tứ trụ có ',(X.TEXT(E.v(8,r-136,c-3), '0') if X.truthy(X.lt(X.ABS(X.sub(E.v(8,r-136,c-3),X.ROUND(E.v(8,r-136,c-3), 0.0))),1e-09)) else X.SUBSTITUTE(X.TEXT(E.v(8,r-136,c-3), '0.0'), '.', ','))),' chữ.') if X.truthy(X.gt(E.v(8,r-136,c-3),1e-09)) else 'tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).')) if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),0.0), X.eq(E.v(8,r-48,c-3),0.0))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: Quan sát đang thái quá (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.113–114: chế áp mạnh quá mà không có ấn thông quan thì dụng thực thương để khống chế quan sát; thực thương không được quá mạnh vì còn làm xì hơi thân.'),' → Nên xem thêm '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c-3),1.0), 5.0),1.0))),' (thực thương) làm dụng thần thay thế; '),(X.cat(X.cat('tứ trụ có ',(X.TEXT(E.v(8,r-138,c-3), '0') if X.truthy(X.lt(X.ABS(X.sub(E.v(8,r-138,c-3),X.ROUND(E.v(8,r-138,c-3), 0.0))),1e-09)) else X.SUBSTITUTE(X.TEXT(E.v(8,r-138,c-3), '0.0'), '.', ','))),' chữ.') if X.truthy(X.gt(E.v(8,r-138,c-3),1e-09)) else 'tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).')),(' Tứ trụ có ấn: xét thông quan trước (P1 tr.114).' if X.truthy(X.gt(E.v(8,r-135,c-3),1e-09)) else '')) if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),1.0), X.eq(E.v(8,r-48,c-3),3.0))) else (X.cat(X.cat('PP1: Tài đang mạnh nhất (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.114: tài dụng để sinh quan không nên vượng quá; nếu mạnh quá thì thân chuyển nhược, không thắng nổi tài quan.') if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),1.0), X.eq(E.v(8,r-48,c-3),2.0), X.OR(X.gt(E.v(8,r-136,c-3),0.0), X.gt(E.v(8,r-138,c-3),0.0)))) else '')))) if X.truthy(X.ge(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)),0.4)) else '') if X.truthy(X.AND(X.eq(E.v(8,r-79,c-3),''), X.ne(E.v(8,r-144,c-3),2.0), X.ge(E.v(8,r-48,c-3),0.0), X.eq(E.v(8,r-54,c-3),0.0), X.gt(E.v(10,100,3),0.0))) else '')

def f210(E, r, c):
    # Ô G151
    # = IF(AND(D72="",D7<>2,D103>=0,D97=0,PP2_DINH_LUONG!$C$100>0),IF(INDEX(PP2_DINH_LUONG!$F$91:$F$95,D109+1)/PP2_DINH_LUONG!$C$100>=0.4,IF(AND(D7=0,D103=4),
    #   "PP2: Ấn đang thái quá ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG!$F$91:$F$95,D109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.113: nhật nhược đ
    #   ược sinh phù thái quá thì có thể biến thành vượng; khi đó chọn dụng như thân vượng, ấn vượng: dụng tài để chế ấn (P1 tr.106, 109). Đối chiếu PP1."&" →
    #    Nên xem thêm "&INDEX(HANH_TEN,MOD(D$5+2,5)+1)&" (tài) làm dụng thần thay thế; "&IF(D14>0.000000001,"tứ trụ có "&IF(ABS(D14-ROUND(D14,0))<0.000000001,
    #   TEXT(D14,"0"),SUBSTITUTE(TEXT(D14,"0.0"),".",","))&" độ.","tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97)."),IF(AND(D7=0,D103=0),"PP2: Tỉ kiếp 
    #   đang thái quá ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG!$F$91:$F$95,D109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.113: phù trợ mạnh quá thì 
    #   hành chế áp sự phù trợ là dụng; nhật nhược thành vượng thì chọn như nhật vượng, tỉ kiếp vượng: lấy quan sát làm dụng, không cần quan sát thật mạnh. Đố
    #   i chiếu PP1."&" → Nên xem thêm "&INDEX(HANH_TEN,MOD(D$5+3,5)+1)&" (quan sát) làm dụng thần thay thế; "&IF(D15>0.000000001,"tứ trụ có "&IF(ABS(D15-ROUN
    #   D(D15,0))<0.000000001,TEXT(D15,"0"),SUBSTITUTE(TEXT(D15,"0.0"),".",","))&" độ.","tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97)."),IF(AND(D7=1,
    #   D103=3),"PP2: Quan sát đang thái quá ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG!$F$91:$F$95,D109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.113
    #   –114: chế áp mạnh quá mà không có ấn thông quan thì dụng thực thương để khống chế quan sát; thực thương không được quá mạnh vì còn làm xì hơi thân."&"
    #    → Nên xem thêm "&INDEX(HANH_TEN,MOD(D$5+1,5)+1)&" (thực thương) làm dụng thần thay thế; "&IF(D13>0.000000001,"tứ trụ có "&IF(ABS(D13-ROUND(D13,0))<0.
    #   000000001,TEXT(D13,"0"),SUBSTITUTE(TEXT(D13,"0.0"),".",","))&" độ.","tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).")&IF(D16>0.000000001," Tứ 
    #   trụ có ấn: xét thông quan trước (P1 tr.114).",""),IF(AND(D7=1,D103=2,OR(D15>0,D13>0)),"PP2: Tài đang mạnh nhất ("&SUBSTITUTE(TEXT(INDEX(PP2_DINH_LUONG
    #   !$F$91:$F$95,D109+1)/PP2_DINH_LUONG!$C$100,"0.0%"),".",",")&") – P1 tr.114: tài dụng để sinh quan không nên vượng quá; nếu mạnh quá thì thân chuyển nh
    #   ược, không thắng nổi tài quan.","")))),""),"")
    return (((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: Ấn đang thái quá (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.113: nhật nhược được sinh phù thái quá thì có thể biến thành vượng; khi đó chọn dụng như thân vượng, ấn vượng: dụng tài để chế ấn (P1 tr.106, 109). Đối chiếu PP1.'),' → Nên xem thêm '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c-3),2.0), 5.0),1.0))),' (tài) làm dụng thần thay thế; '),(X.cat(X.cat('tứ trụ có ',(X.TEXT(E.v(8,r-137,c-3), '0') if X.truthy(X.lt(X.ABS(X.sub(E.v(8,r-137,c-3),X.ROUND(E.v(8,r-137,c-3), 0.0))),1e-09)) else X.SUBSTITUTE(X.TEXT(E.v(8,r-137,c-3), '0.0'), '.', ','))),' độ.') if X.truthy(X.gt(E.v(8,r-137,c-3),1e-09)) else 'tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).')) if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),0.0), X.eq(E.v(8,r-48,c-3),4.0))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: Tỉ kiếp đang thái quá (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.113: phù trợ mạnh quá thì hành chế áp sự phù trợ là dụng; nhật nhược thành vượng thì chọn như nhật vượng, tỉ kiếp vượng: lấy quan sát làm dụng, không cần quan sát thật mạnh. Đối chiếu PP1.'),' → Nên xem thêm '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c-3),3.0), 5.0),1.0))),' (quan sát) làm dụng thần thay thế; '),(X.cat(X.cat('tứ trụ có ',(X.TEXT(E.v(8,r-136,c-3), '0') if X.truthy(X.lt(X.ABS(X.sub(E.v(8,r-136,c-3),X.ROUND(E.v(8,r-136,c-3), 0.0))),1e-09)) else X.SUBSTITUTE(X.TEXT(E.v(8,r-136,c-3), '0.0'), '.', ','))),' độ.') if X.truthy(X.gt(E.v(8,r-136,c-3),1e-09)) else 'tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).')) if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),0.0), X.eq(E.v(8,r-48,c-3),0.0))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2: Quan sát đang thái quá (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.113–114: chế áp mạnh quá mà không có ấn thông quan thì dụng thực thương để khống chế quan sát; thực thương không được quá mạnh vì còn làm xì hơi thân.'),' → Nên xem thêm '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c-3),1.0), 5.0),1.0))),' (thực thương) làm dụng thần thay thế; '),(X.cat(X.cat('tứ trụ có ',(X.TEXT(E.v(8,r-138,c-3), '0') if X.truthy(X.lt(X.ABS(X.sub(E.v(8,r-138,c-3),X.ROUND(E.v(8,r-138,c-3), 0.0))),1e-09)) else X.SUBSTITUTE(X.TEXT(E.v(8,r-138,c-3), '0.0'), '.', ','))),' độ.') if X.truthy(X.gt(E.v(8,r-138,c-3),1e-09)) else 'tứ trụ không có, chỉ chờ đại vận, lưu niên (P1 tr.97).')),(' Tứ trụ có ấn: xét thông quan trước (P1 tr.114).' if X.truthy(X.gt(E.v(8,r-135,c-3),1e-09)) else '')) if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),1.0), X.eq(E.v(8,r-48,c-3),3.0))) else (X.cat(X.cat('PP2: Tài đang mạnh nhất (',X.SUBSTITUTE(X.TEXT(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)), '0.0%'), '.', ',')),') – P1 tr.114: tài dụng để sinh quan không nên vượng quá; nếu mạnh quá thì thân chuyển nhược, không thắng nổi tài quan.') if X.truthy(X.AND(X.eq(E.v(8,r-144,c-3),1.0), X.eq(E.v(8,r-48,c-3),2.0), X.OR(X.gt(E.v(8,r-136,c-3),0.0), X.gt(E.v(8,r-138,c-3),0.0)))) else '')))) if X.truthy(X.ge(X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(E.v(8,r-42,c-3),1.0)),E.v(10,100,3)),0.4)) else '') if X.truthy(X.AND(X.eq(E.v(8,r-79,c-3),''), X.ne(E.v(8,r-144,c-3),2.0), X.ge(E.v(8,r-48,c-3),0.0), X.eq(E.v(8,r-54,c-3),0.0), X.gt(E.v(10,100,3),0.0))) else '')

def f211(E, r, c):
    # Ô C154
    # = IF(DOI_CHIEU!$C$33=2,2,1)
    return (2.0 if X.truthy(X.eq(E.v(7,33,3),2.0)) else 1.0)

def f212(E, r, c):
    # Ô C155
    # = IF($C$154=2,D7,C7)
    return (E.v(8,r-148,c+1) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,r-148,c))

def f213(E, r, c):
    # Ô C156
    # = IF($C$154=2,D103,C103)
    return (E.v(8,r-53,c+1) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,r-53,c))

def f214(E, r, c):
    # Ô C157, C158, C159
    # = IF($C$154=2,D109,C109)
    return (E.v(8,r-48,c+1) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,r-48,c))

def f215(E, r, c):
    # Ô C160
    # = IF($C$154=2,D72,C72)&""
    return X.cat((E.v(8,r-88,c+1) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,r-88,c)),'')

def f216(E, r, c):
    # Ô C161
    # = "Chi tháng "&CHOOSE(TRUNG_GIAN!$D$6+1,"Tý – thấp (thuần thủy)","Sửu – thấp (tàng Quý thủy)","Dần – táo (tàng Bính hỏa)","Mão – trung hòa (trọng xuân)"
    #   ,"Thìn – thấp (tàng Quý thủy)","Tị – táo (tàng Bính hỏa)","Ngọ – táo (tàng Đinh hỏa)","Mùi – táo (tàng Đinh hỏa)","Thân – thấp (tàng Nhâm thủy)","Dậu 
    #   – trung hòa (trọng thu)","Tuất – táo (tàng Đinh hỏa)","Hợi – thấp (tàng Nhâm thủy)")&"; mùa: "&CHOOSE(TRUNG_GIAN!$D$6+1,"đông (hàn)","đông (hàn)","xuâ
    #   n","xuân","xuân (cuối)","hạ (nhiệt)","hạ (nhiệt)","hạ (nhiệt)","thu","thu","thu (cuối)","đông (hàn)")&". PP2: Hỏa "&TEXT(IF(PP2_DINH_LUONG!$C$100>0,PP
    #   2_DINH_LUONG!$F$92/PP2_DINH_LUONG!$C$100,0),"0%")&", Thủy "&TEXT(IF(PP2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$95/PP2_DINH_LUONG!$C$100,0),"0%")&" tổng
    #    độ."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chi tháng ',X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 'Tý – thấp (thuần thủy)', lambda: 'Sửu – thấp (tàng Quý thủy)', lambda: 'Dần – táo (tàng Bính hỏa)', lambda: 'Mão – trung hòa (trọng xuân)', lambda: 'Thìn – thấp (tàng Quý thủy)', lambda: 'Tị – táo (tàng Bính hỏa)', lambda: 'Ngọ – táo (tàng Đinh hỏa)', lambda: 'Mùi – táo (tàng Đinh hỏa)', lambda: 'Thân – thấp (tàng Nhâm thủy)', lambda: 'Dậu – trung hòa (trọng thu)', lambda: 'Tuất – táo (tàng Đinh hỏa)', lambda: 'Hợi – thấp (tàng Nhâm thủy)')),'; mùa: '),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 'đông (hàn)', lambda: 'đông (hàn)', lambda: 'xuân', lambda: 'xuân', lambda: 'xuân (cuối)', lambda: 'hạ (nhiệt)', lambda: 'hạ (nhiệt)', lambda: 'hạ (nhiệt)', lambda: 'thu', lambda: 'thu', lambda: 'thu (cuối)', lambda: 'đông (hàn)')),'. PP2: Hỏa '),X.TEXT((X.div(E.v(10,92,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),', Thủy '),X.TEXT((X.div(E.v(10,95,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),' tổng độ.')

def f217(E, r, c):
    # Ô C162
    # = IF($J$162>=0,$J$162,$K$162)
    return (E.v(8,162,10) if X.truthy(X.ge(E.v(8,162,10),0.0)) else E.v(8,162,11))

def f218(E, r, c):
    # Ô J162
    # = IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),4,IF(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),1,-1))
    return (4.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else X.neg(1.0)))

def f219(E, r, c):
    # Ô K162
    # = IF($J$162>=0,-1,IF(AND(SUMPRODUCT(--ISNUMBER(MATCH(TRUNG_GIAN!$C$6:$F$6,{0,1,4,8,11},0)))>=3,IF(PP2_DINH_LUONG!$C$100>0,INDEX(PP2_DINH_LUONG!$F$91:$F$
    #   95,1+1)/PP2_DINH_LUONG!$C$100,0)<0.1),1,IF(AND(SUMPRODUCT(--ISNUMBER(MATCH(TRUNG_GIAN!$C$6:$F$6,{2,5,6,7,10},0)))>=3,IF(PP2_DINH_LUONG!$C$100>0,INDEX(
    #   PP2_DINH_LUONG!$F$91:$F$95,4+1)/PP2_DINH_LUONG!$C$100,0)<0.1),4,-1)))
    return (X.neg(1.0) if X.truthy(X.ge(E.v(8,162,10),0.0)) else (1.0 if X.truthy(X.AND(X.ge(X.SUMPRODUCT(X.A('neg',X.A('neg',X.AF_ISNUMBER(lambda: X.AF_MATCH(E.rg(23,6,3,6,6), X.Arr([[0.0, 1.0, 4.0, 8.0, 11.0]]), 0.0))))),3.0), X.lt((X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(1.0,1.0)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.1))) else (4.0 if X.truthy(X.AND(X.ge(X.SUMPRODUCT(X.A('neg',X.A('neg',X.AF_ISNUMBER(lambda: X.AF_MATCH(E.rg(23,6,3,6,6), X.Arr([[2.0, 5.0, 6.0, 7.0, 10.0]]), 0.0))))),3.0), X.lt((X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(4.0,1.0)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.1))) else X.neg(1.0))))

def f220(E, r, c):
    # Ô L162
    # = SUMPRODUCT(--ISNUMBER(MATCH(TRUNG_GIAN!$C$6:$F$6,{0,1,4,8,11},0)))
    return X.SUMPRODUCT(X.A('neg',X.A('neg',X.AF_ISNUMBER(lambda: X.AF_MATCH(E.rg(23,6,3,6,6), X.Arr([[0.0, 1.0, 4.0, 8.0, 11.0]]), 0.0)))))

def f221(E, r, c):
    # Ô M162
    # = SUMPRODUCT(--ISNUMBER(MATCH(TRUNG_GIAN!$C$6:$F$6,{2,5,6,7,10},0)))
    return X.SUMPRODUCT(X.A('neg',X.A('neg',X.AF_ISNUMBER(lambda: X.AF_MATCH(E.rg(23,6,3,6,6), X.Arr([[2.0, 5.0, 6.0, 7.0, 10.0]]), 0.0)))))

def f222(E, r, c):
    # Ô C163
    # = IF($C$162<0,"KHÔNG CẦN",IF(AND($J$162>=0,TRUNG_GIAN!$D$28<=4,$C$155=1,OR(AND($C$162=4,OR(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$5=3)),AND($C$162=1,OR(TRUNG_G
    #   IAN!$E$5=8,TRUNG_GIAN!$E$5=9)))),"BẮT BUỘC",IF($J$162<0,"CẦN (bổ trợ)","CẦN")))
    return ('KHÔNG CẦN' if X.truthy(X.lt(E.v(8,162,3),0.0)) else ('BẮT BUỘC' if X.truthy(X.AND(X.ge(E.v(8,162,10),0.0), X.le(E.v(23,28,4),4.0), X.eq(E.v(8,155,3),1.0), X.OR(X.AND(X.eq(E.v(8,162,3),4.0), X.OR(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0))), X.AND(X.eq(E.v(8,162,3),1.0), X.OR(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,5,5),9.0)))))) else ('CẦN (bổ trợ)' if X.truthy(X.lt(E.v(8,162,10),0.0)) else 'CẦN')))

def f223(E, r, c):
    # Ô C164
    # = IF($C$162<0,0,COUNTIF(TRUNG_GIAN!$C$7:$F$7,$C$162)-IF(TRUNG_GIAN!$E$7=$C$162,1,0)+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=$C$162)*(PP2_DINH_LUONG!$C$24:$F$
    #   24<0))+COUNTIF(PP2_DINH_LUONG!$C$24:$F$24,$C$162))
    return (0.0 if X.truthy(X.lt(E.v(8,162,3),0.0)) else X.add(X.add(X.sub(X.COUNTIF(E.rg(23,7,3,7,6), E.v(8,162,3)),(1.0 if X.truthy(X.eq(E.v(23,7,5),E.v(8,162,3))) else 0.0)),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),E.v(8,162,3)),X.A('lt',E.rg(10,24,3,24,6),0.0)))),X.COUNTIF(E.rg(10,24,3,24,6), E.v(8,162,3))))

def f224(E, r, c):
    # Ô C165
    # = IF($C$162<0,"",TRIM(IF(AND(DAI_VAN!C$35<>"",DAI_VAN!C$36<=90,OR(DAI_VAN!C$42=$C$162,DAI_VAN!C$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!C$34>=8,DAI_VA
    #   N!C$34<=1),AND(DAI_VAN!C$34>=2,DAI_VAN!C$34<=7)))),DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_VAN!C$36+9)&" t.); ","")&IF(AND(DAI_VAN!D$35<>"",DAI_VAN!D$
    #   36<=90,OR(DAI_VAN!D$42=$C$162,DAI_VAN!D$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!D$34>=8,DAI_VAN!D$34<=1),AND(DAI_VAN!D$34>=2,DAI_VAN!D$34<=7)))),DAI
    #   _VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&" t.); ","")&IF(AND(DAI_VAN!E$35<>"",DAI_VAN!E$36<=90,OR(DAI_VAN!E$42=$C$162,DAI_VAN!E$42=MOD($C$162+
    #   4,5),IF($C$162=4,OR(DAI_VAN!E$34>=8,DAI_VAN!E$34<=1),AND(DAI_VAN!E$34>=2,DAI_VAN!E$34<=7)))),DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&" t.)
    #   ; ","")&IF(AND(DAI_VAN!F$35<>"",DAI_VAN!F$36<=90,OR(DAI_VAN!F$42=$C$162,DAI_VAN!F$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!F$34>=8,DAI_VAN!F$34<=1),A
    #   ND(DAI_VAN!F$34>=2,DAI_VAN!F$34<=7)))),DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&" t.); ","")&IF(AND(DAI_VAN!G$35<>"",DAI_VAN!G$36<=90,OR(DA
    #   I_VAN!G$42=$C$162,DAI_VAN!G$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!G$34>=8,DAI_VAN!G$34<=1),AND(DAI_VAN!G$34>=2,DAI_VAN!G$34<=7)))),DAI_VAN!G$35&" 
    #   ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&" t.); ","")&IF(AND(DAI_VAN!H$35<>"",DAI_VAN!H$36<=90,OR(DAI_VAN!H$42=$C$162,DAI_VAN!H$42=MOD($C$162+4,5),IF($C$1
    #   62=4,OR(DAI_VAN!H$34>=8,DAI_VAN!H$34<=1),AND(DAI_VAN!H$34>=2,DAI_VAN!H$34<=7)))),DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&" t.); ","")&IF(A
    #   ND(DAI_VAN!I$35<>"",DAI_VAN!I$36<=90,OR(DAI_VAN!I$42=$C$162,DAI_VAN!I$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!I$34>=8,DAI_VAN!I$34<=1),AND(DAI_VAN!I
    #   $34>=2,DAI_VAN!I$34<=7)))),DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&" t.); ","")&IF(AND(DAI_VAN!J$35<>"",DAI_VAN!J$36<=90,OR(DAI_VAN!J$42=$
    #   C$162,DAI_VAN!J$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!J$34>=8,DAI_VAN!J$34<=1),AND(DAI_VAN!J$34>=2,DAI_VAN!J$34<=7)))),DAI_VAN!J$35&" ("&DAI_VAN!J
    #   $36&"–"&(DAI_VAN!J$36+9)&" t.); ","")&IF(AND(DAI_VAN!K$35<>"",DAI_VAN!K$36<=90,OR(DAI_VAN!K$42=$C$162,DAI_VAN!K$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_
    #   VAN!K$34>=8,DAI_VAN!K$34<=1),AND(DAI_VAN!K$34>=2,DAI_VAN!K$34<=7)))),DAI_VAN!K$35&" ("&DAI_VAN!K$36&"–"&(DAI_VAN!K$36+9)&" t.); ","")&IF(AND(DAI_VAN!L
    #   $35<>"",DAI_VAN!L$36<=90,OR(DAI_VAN!L$42=$C$162,DAI_VAN!L$42=MOD($C$162+4,5),IF($C$162=4,OR(DAI_VAN!L$34>=8,DAI_VAN!L$34<=1),AND(DAI_VAN!L$34>=2,DAI_V
    #   AN!L$34<=7)))),DAI_VAN!L$35&" ("&DAI_VAN!L$36&"–"&(DAI_VAN!L$36+9)&" t.); ","")))
    return ('' if X.truthy(X.lt(E.v(8,162,3),0.0)) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c),''), X.le(E.v(14,36,c),90.0), X.OR(X.eq(E.v(14,42,c),E.v(8,162,3)), X.eq(E.v(14,42,c),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c),8.0), X.le(E.v(14,34,c),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c),2.0), X.le(E.v(14,34,c),7.0)))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+1),''), X.le(E.v(14,36,c+1),90.0), X.OR(X.eq(E.v(14,42,c+1),E.v(8,162,3)), X.eq(E.v(14,42,c+1),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+1),8.0), X.le(E.v(14,34,c+1),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+1),2.0), X.le(E.v(14,34,c+1),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+2),''), X.le(E.v(14,36,c+2),90.0), X.OR(X.eq(E.v(14,42,c+2),E.v(8,162,3)), X.eq(E.v(14,42,c+2),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+2),8.0), X.le(E.v(14,34,c+2),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+2),2.0), X.le(E.v(14,34,c+2),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+3),''), X.le(E.v(14,36,c+3),90.0), X.OR(X.eq(E.v(14,42,c+3),E.v(8,162,3)), X.eq(E.v(14,42,c+3),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+3),8.0), X.le(E.v(14,34,c+3),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+3),2.0), X.le(E.v(14,34,c+3),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+4),''), X.le(E.v(14,36,c+4),90.0), X.OR(X.eq(E.v(14,42,c+4),E.v(8,162,3)), X.eq(E.v(14,42,c+4),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+4),8.0), X.le(E.v(14,34,c+4),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+4),2.0), X.le(E.v(14,34,c+4),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+5),''), X.le(E.v(14,36,c+5),90.0), X.OR(X.eq(E.v(14,42,c+5),E.v(8,162,3)), X.eq(E.v(14,42,c+5),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+5),8.0), X.le(E.v(14,34,c+5),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+5),2.0), X.le(E.v(14,34,c+5),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+6),''), X.le(E.v(14,36,c+6),90.0), X.OR(X.eq(E.v(14,42,c+6),E.v(8,162,3)), X.eq(E.v(14,42,c+6),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+6),8.0), X.le(E.v(14,34,c+6),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+6),2.0), X.le(E.v(14,34,c+6),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+7),''), X.le(E.v(14,36,c+7),90.0), X.OR(X.eq(E.v(14,42,c+7),E.v(8,162,3)), X.eq(E.v(14,42,c+7),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+7),8.0), X.le(E.v(14,34,c+7),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+7),2.0), X.le(E.v(14,34,c+7),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+8),' ('),E.v(14,36,c+8)),'–'),X.add(E.v(14,36,c+8),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+8),''), X.le(E.v(14,36,c+8),90.0), X.OR(X.eq(E.v(14,42,c+8),E.v(8,162,3)), X.eq(E.v(14,42,c+8),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+8),8.0), X.le(E.v(14,34,c+8),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+8),2.0), X.le(E.v(14,34,c+8),7.0)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+9),' ('),E.v(14,36,c+9)),'–'),X.add(E.v(14,36,c+9),9.0)),' t.); ') if X.truthy(X.AND(X.ne(E.v(14,35,c+9),''), X.le(E.v(14,36,c+9),90.0), X.OR(X.eq(E.v(14,42,c+9),E.v(8,162,3)), X.eq(E.v(14,42,c+9),X.MOD(X.add(E.v(8,162,3),4.0), 5.0)), (X.OR(X.ge(E.v(14,34,c+9),8.0), X.le(E.v(14,34,c+9),1.0)) if X.truthy(X.eq(E.v(8,162,3),4.0)) else X.AND(X.ge(E.v(14,34,c+9),2.0), X.le(E.v(14,34,c+9),7.0)))))) else ''))))

def f225(E, r, c):
    # Ô C166
    # = TRUNG_GIAN!$E$5*12+MOD(TRUNG_GIAN!$D$6-2,12)+1
    return X.add(X.add(X.mul(E.v(23,5,5),12.0),X.MOD(X.sub(E.v(23,6,4),2.0), 12.0)),1.0)

def f226(E, r, c):
    # Ô C167
    # = INDEX(CHOOSE($M$187+1,CSDL!$AD$5:$AD$124,CSDL!$AN$5:$AN$124,CSDL!$AU$5:$AU$124,CSDL!$BB$5:$BB$124),$C$166)
    return X.INDEX(X.CHOOSE(X.add(E.v(8,187,13),1.0), lambda: E.rg(24,5,30,124,30), lambda: E.rg(24,5,40,124,40), lambda: E.rg(24,5,47,124,47), lambda: E.rg(24,5,54,124,54)), E.v(8,166,3))

def f227(E, r, c):
    # Ô D167
    # = INDEX(CHOOSE($M$187+1,CSDL!$AE$5:$AE$124,CSDL!$AO$5:$AO$124,CSDL!$AV$5:$AV$124,CSDL!$BC$5:$BC$124),$C$166)
    return X.INDEX(X.CHOOSE(X.add(E.v(8,187,13),1.0), lambda: E.rg(24,5,31,124,31), lambda: E.rg(24,5,41,124,41), lambda: E.rg(24,5,48,124,48), lambda: E.rg(24,5,55,124,55)), E.v(8,166,3))

def f228(E, r, c):
    # Ô E167
    # = INDEX(CHOOSE($M$187+1,CSDL!$AF$5:$AF$124,CSDL!$AP$5:$AP$124,CSDL!$AW$5:$AW$124,CSDL!$BD$5:$BD$124),$C$166)
    return X.INDEX(X.CHOOSE(X.add(E.v(8,187,13),1.0), lambda: E.rg(24,5,32,124,32), lambda: E.rg(24,5,42,124,42), lambda: E.rg(24,5,49,124,49), lambda: E.rg(24,5,56,124,56)), E.v(8,166,3))

def f229(E, r, c):
    # Ô F167
    # = INDEX(CHOOSE($M$187+1,CSDL!$AG$5:$AG$124,CSDL!$AQ$5:$AQ$124,CSDL!$AX$5:$AX$124,CSDL!$BE$5:$BE$124),$C$166)
    return X.INDEX(X.CHOOSE(X.add(E.v(8,187,13),1.0), lambda: E.rg(24,5,33,124,33), lambda: E.rg(24,5,43,124,43), lambda: E.rg(24,5,50,124,50), lambda: E.rg(24,5,57,124,57)), E.v(8,166,3))

def f230(E, r, c):
    # Ô G167
    # = INDEX(CHOOSE($M$187+1,CSDL!$AH$5:$AH$124,CSDL!$AR$5:$AR$124,CSDL!$AY$5:$AY$124,CSDL!$BF$5:$BF$124),$C$166)
    return X.INDEX(X.CHOOSE(X.add(E.v(8,187,13),1.0), lambda: E.rg(24,5,34,124,34), lambda: E.rg(24,5,44,124,44), lambda: E.rg(24,5,51,124,51), lambda: E.rg(24,5,58,124,58)), E.v(8,166,3))

def f231(E, r, c):
    # Ô C168, D168, E168, F168, G168
    # = IF(C167<0,-1,IF((TRUNG_GIAN!$C$5=C167)+(TRUNG_GIAN!$D$5=C167)+(TRUNG_GIAN!$F$5=C167)>0,2,IF(SUMPRODUCT(((TRUNG_GIAN!$C$15:$F$15=C167)+(TRUNG_GIAN!$C$1
    #   7:$F$17=C167)+(TRUNG_GIAN!$C$19:$F$19=C167))*(PP2_DINH_LUONG!$C$24:$F$24<0))>0,1,0)))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(8,r-1,c),0.0)) else (2.0 if X.truthy(X.gt(X.add(X.add(X.eq(E.v(23,5,3),E.v(8,r-1,c)),X.eq(E.v(23,5,4),E.v(8,r-1,c))),X.eq(E.v(23,5,6),E.v(8,r-1,c))),0.0)) else (1.0 if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('add',X.A('add',X.A('eq',E.rg(23,15,3,15,6),E.v(8,r-1,c)),X.A('eq',E.rg(23,17,3,17,6),E.v(8,r-1,c))),X.A('eq',E.rg(23,19,3,19,6),E.v(8,r-1,c))),X.A('lt',E.rg(10,24,3,24,6),0.0))),0.0)) else 0.0)))

def f232(E, r, c):
    # Ô C169
    # = IF(C168=2,C167,IF(D168=2,D167,IF(E168=2,E167,IF(F168=2,F167,IF(G168=2,G167,IF(C168=1,C167,IF(D168=1,D167,IF(E168=1,E167,IF(F168=1,F167,IF(G168=1,G167,
    #   -1))))))))))
    return (E.v(8,r-2,c) if X.truthy(X.eq(E.v(8,r-1,c),2.0)) else (E.v(8,r-2,c+1) if X.truthy(X.eq(E.v(8,r-1,c+1),2.0)) else (E.v(8,r-2,c+2) if X.truthy(X.eq(E.v(8,r-1,c+2),2.0)) else (E.v(8,r-2,c+3) if X.truthy(X.eq(E.v(8,r-1,c+3),2.0)) else (E.v(8,r-2,c+4) if X.truthy(X.eq(E.v(8,r-1,c+4),2.0)) else (E.v(8,r-2,c) if X.truthy(X.eq(E.v(8,r-1,c),1.0)) else (E.v(8,r-2,c+1) if X.truthy(X.eq(E.v(8,r-1,c+1),1.0)) else (E.v(8,r-2,c+2) if X.truthy(X.eq(E.v(8,r-1,c+2),1.0)) else (E.v(8,r-2,c+3) if X.truthy(X.eq(E.v(8,r-1,c+3),1.0)) else (E.v(8,r-2,c+4) if X.truthy(X.eq(E.v(8,r-1,c+4),1.0)) else X.neg(1.0)))))))))))

def f233(E, r, c):
    # Ô C170
    # = "Nhật can "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng "&(MOD(TRUNG_GIAN!$D$6-2,12)+1)&" theo tiết): sá
    #   ch cho dùng theo thứ tự "&INDEX(CSDL!$AI$5:$AI$124,$C$166)&". Chú thích: "&INDEX(CSDL!$AJ$5:$AJ$124,$C$166)&IF(INDEX(CSDL!$AL$5:$AL$124,$C$166)=1," [B
    #   ảng tr.389-390 ghi: "&INDEX(CSDL!$AK$5:$AK$124,$C$166)&"]","")&IF($M$187>0," → ĐIỀU KIỆN CHÚ THÍCH ĐỦ ("&$N$187&"): dùng thứ tự "&INDEX(CHOOSE($M$187,
    #   CSDL!$AS$5:$AS$124,CSDL!$AZ$5:$AZ$124,CSDL!$BG$5:$BG$124),$C$166)&".","")&" Trong tứ trụ: "&TRIM(IF(C167<0,"",INDEX(CAN_TEN,C167+1)&" "&CHOOSE(C168+1,
    #   "(không có)","(tàng chi)","(thấu can)")&"; ")&IF(D167<0,"",INDEX(CAN_TEN,D167+1)&" "&CHOOSE(D168+1,"(không có)","(tàng chi)","(thấu can)")&"; ")&IF(E1
    #   67<0,"",INDEX(CAN_TEN,E167+1)&" "&CHOOSE(E168+1,"(không có)","(tàng chi)","(thấu can)")&"; ")&IF(F167<0,"",INDEX(CAN_TEN,F167+1)&" "&CHOOSE(F168+1,"(k
    #   hông có)","(tàng chi)","(thấu can)")&"; ")&IF(G167<0,"",INDEX(CAN_TEN,G167+1)&" "&CHOOSE(G168+1,"(không có)","(tàng chi)","(thấu can)")&"; "))&" → "&I
    #   F($C$169<0,"KHÔNG có can nào: thiếu dụng thần điều hầu, chờ đại vận, lưu niên mang "&INDEX(CAN_TEN,$C$167+1)&" ("&INDEX(HANH_TEN,INT($C$167/2)+1)&")."
    #   ,"chọn "&INDEX(CAN_TEN,$C$169+1)&IF(INDEX(C168:G168,MATCH($C$169,C167:G167,0))=2," – thấu can"," – chỉ tàng trong chi")&" ("&INDEX(HANH_TEN,INT($C$169
    #   /2)+1)&")"&IF(INT($C$169/2)=$C$157," – TRÙNG dụng thần phù ức.",IF(INT($C$169/2)=$C$158," – trùng hỷ thần.",IF(INT($C$169/2)=$C$159," – ⚠ trùng kỵ thầ
    #   n phù ức: dụng thần sinh phù vẫn làm chủ.","."))))&IF(AND($C$169>=0,INT($C$169/2)=$O$187,INT($C$169/2)<>$C$157,INT($C$169/2)<>$C$158)," ⚠ Can được chọ
    #   n thuộc hành NHIỀU NHẤT của tứ trụ ("&INDEX(HANH_TEN,$O$187+1)&") – chính là bệnh của mệnh cục, không dùng làm thuốc.","")&IF(AND($C$169<0,$C$167>=0),
    #   IF(INT($C$167/2)=$C$159," ⚠ Hành này là KỴ THẦN phù ức – dụng thần sinh phù vẫn làm chủ.",""),"")&$C$186&" (P2 tr.387-397: bảng do độc giả tặng, tác g
    #   iả giới thiệu để tham khảo)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật can ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' sinh tháng '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' (tháng '),X.add(X.MOD(X.sub(E.v(23,6,4),2.0), 12.0),1.0)),' theo tiết): sách cho dùng theo thứ tự '),X.INDEX(E.rg(24,5,35,124,35), E.v(8,166,3))),'. Chú thích: '),X.INDEX(E.rg(24,5,36,124,36), E.v(8,166,3))),(X.cat(X.cat(' [Bảng tr.389-390 ghi: ',X.INDEX(E.rg(24,5,37,124,37), E.v(8,166,3))),']') if X.truthy(X.eq(X.INDEX(E.rg(24,5,38,124,38), E.v(8,166,3)),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(' → ĐIỀU KIỆN CHÚ THÍCH ĐỦ (',E.v(8,187,14)),'): dùng thứ tự '),X.INDEX(X.CHOOSE(E.v(8,187,13), lambda: E.rg(24,5,45,124,45), lambda: E.rg(24,5,52,124,52), lambda: E.rg(24,5,59,124,59)), E.v(8,166,3))),'.') if X.truthy(X.gt(E.v(8,187,13),0.0)) else '')),' Trong tứ trụ: '),X.TRIM(X.cat(X.cat(X.cat(X.cat(('' if X.truthy(X.lt(E.v(8,r-3,c),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,r-3,c),1.0)),' '),X.CHOOSE(X.add(E.v(8,r-2,c),1.0), lambda: '(không có)', lambda: '(tàng chi)', lambda: '(thấu can)')),'; ')),('' if X.truthy(X.lt(E.v(8,r-3,c+1),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,r-3,c+1),1.0)),' '),X.CHOOSE(X.add(E.v(8,r-2,c+1),1.0), lambda: '(không có)', lambda: '(tàng chi)', lambda: '(thấu can)')),'; '))),('' if X.truthy(X.lt(E.v(8,r-3,c+2),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,r-3,c+2),1.0)),' '),X.CHOOSE(X.add(E.v(8,r-2,c+2),1.0), lambda: '(không có)', lambda: '(tàng chi)', lambda: '(thấu can)')),'; '))),('' if X.truthy(X.lt(E.v(8,r-3,c+3),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,r-3,c+3),1.0)),' '),X.CHOOSE(X.add(E.v(8,r-2,c+3),1.0), lambda: '(không có)', lambda: '(tàng chi)', lambda: '(thấu can)')),'; '))),('' if X.truthy(X.lt(E.v(8,r-3,c+4),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,r-3,c+4),1.0)),' '),X.CHOOSE(X.add(E.v(8,r-2,c+4),1.0), lambda: '(không có)', lambda: '(tàng chi)', lambda: '(thấu can)')),'; '))))),' → '),(X.cat(X.cat(X.cat(X.cat('KHÔNG có can nào: thiếu dụng thần điều hầu, chờ đại vận, lưu niên mang ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,167,3),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(8,167,3),2.0)),1.0))),').') if X.truthy(X.lt(E.v(8,169,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('chọn ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(8,169,3),1.0))),(' – thấu can' if X.truthy(X.eq(X.INDEX(E.rg(8,r-2,c,r-2,c+4), X.MATCH(E.v(8,169,3), E.rg(8,r-3,c,r-3,c+4), 0.0)),2.0)) else ' – chỉ tàng trong chi')),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(8,169,3),2.0)),1.0))),')'),(' – TRÙNG dụng thần phù ức.' if X.truthy(X.eq(X.INT(X.div(E.v(8,169,3),2.0)),E.v(8,157,3))) else (' – trùng hỷ thần.' if X.truthy(X.eq(X.INT(X.div(E.v(8,169,3),2.0)),E.v(8,158,3))) else (' – ⚠ trùng kỵ thần phù ức: dụng thần sinh phù vẫn làm chủ.' if X.truthy(X.eq(X.INT(X.div(E.v(8,169,3),2.0)),E.v(8,159,3))) else '.')))))),(X.cat(X.cat(' ⚠ Can được chọn thuộc hành NHIỀU NHẤT của tứ trụ (',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,187,15),1.0))),') – chính là bệnh của mệnh cục, không dùng làm thuốc.') if X.truthy(X.AND(X.ge(E.v(8,169,3),0.0), X.eq(X.INT(X.div(E.v(8,169,3),2.0)),E.v(8,187,15)), X.ne(X.INT(X.div(E.v(8,169,3),2.0)),E.v(8,157,3)), X.ne(X.INT(X.div(E.v(8,169,3),2.0)),E.v(8,158,3)))) else '')),((' ⚠ Hành này là KỴ THẦN phù ức – dụng thần sinh phù vẫn làm chủ.' if X.truthy(X.eq(X.INT(X.div(E.v(8,167,3),2.0)),E.v(8,159,3))) else '') if X.truthy(X.AND(X.lt(E.v(8,169,3),0.0), X.ge(E.v(8,167,3),0.0))) else '')),E.v(8,186,3)),' (P2 tr.387-397: bảng do độc giả tặng, tác giả giới thiệu để tham khảo)')

def f234(E, r, c):
    # Ô C171
    # = IF($C$162<0,"KHÔNG CẦN điều hầu: sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (xuân, thu hoặc cuối mùa) nóng lạnh vừa phải – chỉ xét sinh khắc chế 
    #   hóa (P1 tr.100-101).",$C$163&" điều hầu bằng "&INDEX(HANH_TEN,$C$162+1)&" ("&IF($J$162<0,"sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" không phải m
    #   ùa đông/hạ nhưng tứ trụ "&IF($C$162=1,$L$162&"/4 chi thấp, Hỏa dưới 10% → thiên hàn thấp, cần ấm",$M$162&"/4 chi táo, Thủy dưới 10% → thiên táo nóng, 
    #   cần nhuận")&"; ngưỡng là quy ước kỹ thuật",IF($C$162=4,"tứ trụ sinh mùa hạ, thiên nóng – khô, cần nhuận","tứ trụ sinh mùa đông, thiên lạnh – ẩm, cần ấ
    #   m"))&"). "&IF($C$164>0,"Tứ trụ CÓ "&$C$164&" chữ hành "&INDEX(HANH_TEN,$C$162+1)&" → có thuốc điều hầu.","Tứ trụ KHUYẾT hành "&INDEX(HANH_TEN,$C$162+1
    #   )&" → "&IF($C$163="BẮT BUỘC","tứ trụ BỆNH NẶNG (P1 tr.112)","thiếu điều hầu")&"; vận bổ cứu: "&IF($C$165="","không có trong 10 đại vận",$C$165))&IF($C
    #   $162=4,IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,7)+COUNTIF(TRUNG_GIAN!$C$6:$F$6,10)>0," Có Mùi/Tuất (táo thổ) càng làm tăng thế hỏa.",""),IF(COUNTIF(TRUNG_GIAN
    #   !$C$6:$F$6,1)+COUNTIF(TRUNG_GIAN!$C$6:$F$6,4)>0," Có Sửu/Thìn (thấp thổ) không chặn được thủy mà còn tăng thế thủy.",""))&IF($C$162=$C$159," ⚠ Hành đi
    #   ều hầu TRÙNG kỵ thần phù ức: hai nhu cầu ngược nhau – chấm vận, năm vẫn theo dụng thần sinh phù.",IF($C$163="BẮT BUỘC",IF($C$162=4," Lợi đi về phương 
    #   Bắc, dùng màu đen."," Lợi đi về phương Nam, dùng màu đỏ."),""))&" Dụng thần sinh phù làm chủ, điều hầu là bổ trợ (P1 tr.100, 111-113).")
    return (X.cat(X.cat('KHÔNG CẦN điều hầu: sinh tháng ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' (xuân, thu hoặc cuối mùa) nóng lạnh vừa phải – chỉ xét sinh khắc chế hóa (P1 tr.100-101).') if X.truthy(X.lt(E.v(8,162,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(8,163,3),' điều hầu bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,162,3),1.0))),' ('),(X.cat(X.cat(X.cat(X.cat('sinh tháng ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' không phải mùa đông/hạ nhưng tứ trụ '),(X.cat(E.v(8,162,12),'/4 chi thấp, Hỏa dưới 10% → thiên hàn thấp, cần ấm') if X.truthy(X.eq(E.v(8,162,3),1.0)) else X.cat(E.v(8,162,13),'/4 chi táo, Thủy dưới 10% → thiên táo nóng, cần nhuận'))),'; ngưỡng là quy ước kỹ thuật') if X.truthy(X.lt(E.v(8,162,10),0.0)) else ('tứ trụ sinh mùa hạ, thiên nóng – khô, cần nhuận' if X.truthy(X.eq(E.v(8,162,3),4.0)) else 'tứ trụ sinh mùa đông, thiên lạnh – ẩm, cần ấm'))),'). '),(X.cat(X.cat(X.cat(X.cat('Tứ trụ CÓ ',E.v(8,164,3)),' chữ hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,162,3),1.0))),' → có thuốc điều hầu.') if X.truthy(X.gt(E.v(8,164,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat('Tứ trụ KHUYẾT hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,162,3),1.0))),' → '),('tứ trụ BỆNH NẶNG (P1 tr.112)' if X.truthy(X.eq(E.v(8,163,3),'BẮT BUỘC')) else 'thiếu điều hầu')),'; vận bổ cứu: '),('không có trong 10 đại vận' if X.truthy(X.eq(E.v(8,165,3),'')) else E.v(8,165,3))))),((' Có Mùi/Tuất (táo thổ) càng làm tăng thế hỏa.' if X.truthy(X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,6), 7.0),X.COUNTIF(E.rg(23,6,3,6,6), 10.0)),0.0)) else '') if X.truthy(X.eq(E.v(8,162,3),4.0)) else (' Có Sửu/Thìn (thấp thổ) không chặn được thủy mà còn tăng thế thủy.' if X.truthy(X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,6), 1.0),X.COUNTIF(E.rg(23,6,3,6,6), 4.0)),0.0)) else ''))),(' ⚠ Hành điều hầu TRÙNG kỵ thần phù ức: hai nhu cầu ngược nhau – chấm vận, năm vẫn theo dụng thần sinh phù.' if X.truthy(X.eq(E.v(8,162,3),E.v(8,159,3))) else ((' Lợi đi về phương Bắc, dùng màu đen.' if X.truthy(X.eq(E.v(8,162,3),4.0)) else ' Lợi đi về phương Nam, dùng màu đỏ.') if X.truthy(X.eq(E.v(8,163,3),'BẮT BUỘC')) else ''))),' Dụng thần sinh phù làm chủ, điều hầu là bổ trợ (P1 tr.100, 111-113).'))

def f235(E, r, c):
    # Ô C172
    # = IF($C$156<0,-1,MOD($C$156+3,5))
    return (X.neg(1.0) if X.truthy(X.lt(E.v(8,156,3),0.0)) else X.MOD(X.add(E.v(8,156,3),3.0), 5.0))

def f236(E, r, c):
    # Ô C173
    # = IF($C$172<0,0,IF($C$154=2,INDEX($D$18:$D$22,$C$172+1),INDEX($C$18:$C$22,$C$172+1)))
    return (0.0 if X.truthy(X.lt(E.v(8,172,3),0.0)) else (X.INDEX(E.rg(8,18,4,22,4), X.add(E.v(8,172,3),1.0)) if X.truthy(X.eq(E.v(8,154,3),2.0)) else X.INDEX(E.rg(8,18,3,22,3), X.add(E.v(8,172,3),1.0))))

def f237(E, r, c):
    # Ô D173
    # = IF($C$172<0,0,IF($C$154=2,INDEX($D$18:$D$22,MOD($C$156+1,5)+1),INDEX($C$18:$C$22,MOD($C$156+1,5)+1)))
    return (0.0 if X.truthy(X.lt(E.v(8,172,3),0.0)) else (X.INDEX(E.rg(8,18,4,22,4), X.add(X.MOD(X.add(E.v(8,156,3),1.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,154,3),2.0)) else X.INDEX(E.rg(8,18,3,22,3), X.add(X.MOD(X.add(E.v(8,156,3),1.0), 5.0),1.0))))

def f238(E, r, c):
    # Ô E173
    # = IF($C$172<0,0,IF($C$154=2,INDEX($D$18:$D$22,MOD($C$156+4,5)+1),INDEX($C$18:$C$22,MOD($C$156+4,5)+1)))
    return (0.0 if X.truthy(X.lt(E.v(8,172,3),0.0)) else (X.INDEX(E.rg(8,18,4,22,4), X.add(X.MOD(X.add(E.v(8,156,3),4.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,154,3),2.0)) else X.INDEX(E.rg(8,18,3,22,3), X.add(X.MOD(X.add(E.v(8,156,3),4.0), 5.0),1.0))))

def f239(E, r, c):
    # Ô F173
    # = IF($C$172<0,0,IF(OR(AND(ABS(TRUNG_GIAN!$C$5-TRUNG_GIAN!$D$5)=5,OR(AND(TRUNG_GIAN!$C$13=$C$172,TRUNG_GIAN!$D$13<>$C$156),AND(TRUNG_GIAN!$D$13=$C$172,TR
    #   UNG_GIAN!$C$13<>$C$156))),AND(ABS(TRUNG_GIAN!$D$5-TRUNG_GIAN!$E$5)=5,TRUNG_GIAN!$D$13=$C$172),AND(ABS(TRUNG_GIAN!$E$5-TRUNG_GIAN!$F$5)=5,TRUNG_GIAN!$F
    #   $13=$C$172)),1,0))
    return (0.0 if X.truthy(X.lt(E.v(8,172,3),0.0)) else (1.0 if X.truthy(X.OR(X.AND(X.eq(X.ABS(X.sub(E.v(23,5,3),E.v(23,5,4))),5.0), X.OR(X.AND(X.eq(E.v(23,13,3),E.v(8,172,3)), X.ne(E.v(23,13,4),E.v(8,156,3))), X.AND(X.eq(E.v(23,13,4),E.v(8,172,3)), X.ne(E.v(23,13,3),E.v(8,156,3))))), X.AND(X.eq(X.ABS(X.sub(E.v(23,5,4),E.v(23,5,5))),5.0), X.eq(E.v(23,13,4),E.v(8,172,3))), X.AND(X.eq(X.ABS(X.sub(E.v(23,5,5),E.v(23,5,6))),5.0), X.eq(E.v(23,13,6),E.v(8,172,3))))) else 0.0))

def f240(E, r, c):
    # Ô C174
    # = IF(OR($C$160<>"",$C$156<0),"Không áp dụng: cách đặc biệt hoặc trung hòa – dụng thần theo cách cục.","Dụng thần "&INDEX(HANH_TEN,$C$157+1)&" ("&INDEX(N
    #   HOM_TEN,$C$156+1)&")"&IF(IF($C$154=2,$D$97,$C$97)=1," – KHUYẾT trong tứ trụ: hỷ thần gánh thay, phải nhờ tuế vận bổ cứu (P1 tr.97)","")&". Kỵ thần khắ
    #   c dụng: "&INDEX(NHOM_TEN,$C$172+1)&" ("&INDEX(HANH_TEN,MOD($C$157+3,5)+1)&") – "&IF($C$173=0,"KHÔNG có trong tứ trụ → dụng thần không bị thương.","CÓ 
    #   trong tứ trụ → dụng thần bị thương. "&IF($D$173>0,INDEX(NHOM_TEN,MOD($C$156+1,5)+1)&" chế kỵ thần – có; ","")&IF($E$173>0,INDEX(NHOM_TEN,MOD($C$156+4,
    #   5)+1)&" tiết kỵ sinh dụng (thông quan) – có; ","")&IF($F$173=1,"can kỵ thần bị ngũ hợp với can kề (hợp mất kỵ) – có; ","")&IF($D$173+$E$173+$F$173>0,"
    #   → CÓ CỨU (bệnh có thuốc).","→ KHÔNG CÓ CỨU (bệnh không thuốc): chờ vận "&INDEX(NHOM_TEN,MOD($C$156+1,5)+1)&" ("&INDEX(HANH_TEN,MOD($C$157+1,5)+1)&") h
    #   oặc "&INDEX(NHOM_TEN,MOD($C$156+4,5)+1)&" ("&INDEX(HANH_TEN,MOD($C$157+4,5)+1)&")."))&$C$185&" (P1 tr.97, 114-117)")
    return ('Không áp dụng: cách đặc biệt hoặc trung hòa – dụng thần theo cách cục.' if X.truthy(X.OR(X.ne(E.v(8,160,3),''), X.lt(E.v(8,156,3),0.0))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dụng thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,157,3),1.0))),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,156,3),1.0))),')'),(' – KHUYẾT trong tứ trụ: hỷ thần gánh thay, phải nhờ tuế vận bổ cứu (P1 tr.97)' if X.truthy(X.eq((E.v(8,97,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,97,3)),1.0)) else '')),'. Kỵ thần khắc dụng: '),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,172,3),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,157,3),3.0), 5.0),1.0))),') – '),('KHÔNG có trong tứ trụ → dụng thần không bị thương.' if X.truthy(X.eq(E.v(8,173,3),0.0)) else X.cat(X.cat(X.cat(X.cat('CÓ trong tứ trụ → dụng thần bị thương. ',(X.cat(X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(E.v(8,156,3),1.0), 5.0),1.0)),' chế kỵ thần – có; ') if X.truthy(X.gt(E.v(8,173,4),0.0)) else '')),(X.cat(X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(E.v(8,156,3),4.0), 5.0),1.0)),' tiết kỵ sinh dụng (thông quan) – có; ') if X.truthy(X.gt(E.v(8,173,5),0.0)) else '')),('can kỵ thần bị ngũ hợp với can kề (hợp mất kỵ) – có; ' if X.truthy(X.eq(E.v(8,173,6),1.0)) else '')),('→ CÓ CỨU (bệnh có thuốc).' if X.truthy(X.gt(X.add(X.add(E.v(8,173,4),E.v(8,173,5)),E.v(8,173,6)),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('→ KHÔNG CÓ CỨU (bệnh không thuốc): chờ vận ',X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(E.v(8,156,3),1.0), 5.0),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,157,3),1.0), 5.0),1.0))),') hoặc '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(E.v(8,156,3),4.0), 5.0),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,157,3),4.0), 5.0),1.0))),').'))))),E.v(8,185,3)),' (P1 tr.97, 114-117)'))

def f241(E, r, c):
    # Ô C175
    # = IF(OR($C$160<>"",$C$156<0),"Không áp dụng (cách đặc biệt / trung hòa).","Vận dụng thần "&INDEX(HANH_TEN,$C$157+1)&" đến ngôi: "&IF($J$175="","không có
    #   ",$J$175)&" Vận kỵ thần "&INDEX(HANH_TEN,MOD($C$157+3,5)+1)&" đến khắc dụng: "&IF($K$175="","không có",$K$175)&" (P1 tr.179 bước 10; P2 tr.387: vận lợ
    #   i dụng thần mà tứ trụ có thần khắc/hợp chặt dụng thì tốt thành xấu – xem DAI_VAN dòng 53)")
    return ('Không áp dụng (cách đặc biệt / trung hòa).' if X.truthy(X.OR(X.ne(E.v(8,160,3),''), X.lt(E.v(8,156,3),0.0))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận dụng thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,157,3),1.0))),' đến ngôi: '),('không có' if X.truthy(X.eq(E.v(8,175,10),'')) else E.v(8,175,10))),' Vận kỵ thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,157,3),3.0), 5.0),1.0))),' đến khắc dụng: '),('không có' if X.truthy(X.eq(E.v(8,175,11),'')) else E.v(8,175,11))),' (P1 tr.179 bước 10; P2 tr.387: vận lợi dụng thần mà tứ trụ có thần khắc/hợp chặt dụng thì tốt thành xấu – xem DAI_VAN dòng 53)'))

def f242(E, r, c):
    # Ô J175
    # = TRIM(IF(AND(DAI_VAN!C$35<>"",DAI_VAN!C$36<=90,OR(DAI_VAN!C$42=$C$157,DAI_VAN!C$43=$C$157)),DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_VAN!C$36+9)&" t.)"&
    #   "; ","")&IF(AND(DAI_VAN!D$35<>"",DAI_VAN!D$36<=90,OR(DAI_VAN!D$42=$C$157,DAI_VAN!D$43=$C$157)),DAI_VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&" t
    #   .)"&"; ","")&IF(AND(DAI_VAN!E$35<>"",DAI_VAN!E$36<=90,OR(DAI_VAN!E$42=$C$157,DAI_VAN!E$43=$C$157)),DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)
    #   &" t.)"&"; ","")&IF(AND(DAI_VAN!F$35<>"",DAI_VAN!F$36<=90,OR(DAI_VAN!F$42=$C$157,DAI_VAN!F$43=$C$157)),DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$3
    #   6+9)&" t.)"&"; ","")&IF(AND(DAI_VAN!G$35<>"",DAI_VAN!G$36<=90,OR(DAI_VAN!G$42=$C$157,DAI_VAN!G$43=$C$157)),DAI_VAN!G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN
    #   !G$36+9)&" t.)"&"; ","")&IF(AND(DAI_VAN!H$35<>"",DAI_VAN!H$36<=90,OR(DAI_VAN!H$42=$C$157,DAI_VAN!H$43=$C$157)),DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI
    #   _VAN!H$36+9)&" t.)"&"; ","")&IF(AND(DAI_VAN!I$35<>"",DAI_VAN!I$36<=90,OR(DAI_VAN!I$42=$C$157,DAI_VAN!I$43=$C$157)),DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&
    #   (DAI_VAN!I$36+9)&" t.)"&"; ","")&IF(AND(DAI_VAN!J$35<>"",DAI_VAN!J$36<=90,OR(DAI_VAN!J$42=$C$157,DAI_VAN!J$43=$C$157)),DAI_VAN!J$35&" ("&DAI_VAN!J$36&
    #   "–"&(DAI_VAN!J$36+9)&" t.)"&"; ","")&IF(AND(DAI_VAN!K$35<>"",DAI_VAN!K$36<=90,OR(DAI_VAN!K$42=$C$157,DAI_VAN!K$43=$C$157)),DAI_VAN!K$35&" ("&DAI_VAN!K
    #   $36&"–"&(DAI_VAN!K$36+9)&" t.)"&"; ","")&IF(AND(DAI_VAN!L$35<>"",DAI_VAN!L$36<=90,OR(DAI_VAN!L$42=$C$157,DAI_VAN!L$43=$C$157)),DAI_VAN!L$35&" ("&DAI_V
    #   AN!L$36&"–"&(DAI_VAN!L$36+9)&" t.)"&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-7),' ('),E.v(14,36,c-7)),'–'),X.add(E.v(14,36,c-7),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-7),''), X.le(E.v(14,36,c-7),90.0), X.OR(X.eq(E.v(14,42,c-7),E.v(8,157,3)), X.eq(E.v(14,43,c-7),E.v(8,157,3))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-6),' ('),E.v(14,36,c-6)),'–'),X.add(E.v(14,36,c-6),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-6),''), X.le(E.v(14,36,c-6),90.0), X.OR(X.eq(E.v(14,42,c-6),E.v(8,157,3)), X.eq(E.v(14,43,c-6),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-5),' ('),E.v(14,36,c-5)),'–'),X.add(E.v(14,36,c-5),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-5),''), X.le(E.v(14,36,c-5),90.0), X.OR(X.eq(E.v(14,42,c-5),E.v(8,157,3)), X.eq(E.v(14,43,c-5),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-4),' ('),E.v(14,36,c-4)),'–'),X.add(E.v(14,36,c-4),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-4),''), X.le(E.v(14,36,c-4),90.0), X.OR(X.eq(E.v(14,42,c-4),E.v(8,157,3)), X.eq(E.v(14,43,c-4),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-3),' ('),E.v(14,36,c-3)),'–'),X.add(E.v(14,36,c-3),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-3),''), X.le(E.v(14,36,c-3),90.0), X.OR(X.eq(E.v(14,42,c-3),E.v(8,157,3)), X.eq(E.v(14,43,c-3),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-2),' ('),E.v(14,36,c-2)),'–'),X.add(E.v(14,36,c-2),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-2),''), X.le(E.v(14,36,c-2),90.0), X.OR(X.eq(E.v(14,42,c-2),E.v(8,157,3)), X.eq(E.v(14,43,c-2),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-1),' ('),E.v(14,36,c-1)),'–'),X.add(E.v(14,36,c-1),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-1),''), X.le(E.v(14,36,c-1),90.0), X.OR(X.eq(E.v(14,42,c-1),E.v(8,157,3)), X.eq(E.v(14,43,c-1),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c),''), X.le(E.v(14,36,c),90.0), X.OR(X.eq(E.v(14,42,c),E.v(8,157,3)), X.eq(E.v(14,43,c),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c+1),''), X.le(E.v(14,36,c+1),90.0), X.OR(X.eq(E.v(14,42,c+1),E.v(8,157,3)), X.eq(E.v(14,43,c+1),E.v(8,157,3))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),' t.)'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c+2),''), X.le(E.v(14,36,c+2),90.0), X.OR(X.eq(E.v(14,42,c+2),E.v(8,157,3)), X.eq(E.v(14,43,c+2),E.v(8,157,3))))) else '')))

def f243(E, r, c):
    # Ô K175
    # = TRIM(IF(AND(DAI_VAN!C$35<>"",DAI_VAN!C$36<=90,OR(DAI_VAN!C$42=MOD($C$157+3,5),DAI_VAN!C$43=MOD($C$157+3,5))),DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_V
    #   AN!C$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!C$42=MOD($C$157+1,5),DAI_VAN!C$43=MOD($C$157+1,5),DAI_VAN!C$42=MOD($C$157+4,5),DAI_VAN!C$43
    #   =MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!D$35<>"",DAI_VAN!D$36<=90,OR(DAI_VAN!D$42=MOD($C$157+3,5),DAI_VAN!D$43=MOD($C$157+3,
    #   5))),DAI_VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!D$42=MOD($C$157+1,5),DAI_VAN!D$43=MOD($C$157+
    #   1,5),DAI_VAN!D$42=MOD($C$157+4,5),DAI_VAN!D$43=MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!E$35<>"",DAI_VAN!E$36<=90,OR(DAI_VAN!E
    #   $42=MOD($C$157+3,5),DAI_VAN!E$43=MOD($C$157+3,5))),DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN
    #   !E$42=MOD($C$157+1,5),DAI_VAN!E$43=MOD($C$157+1,5),DAI_VAN!E$42=MOD($C$157+4,5),DAI_VAN!E$43=MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(
    #   DAI_VAN!F$35<>"",DAI_VAN!F$36<=90,OR(DAI_VAN!F$42=MOD($C$157+3,5),DAI_VAN!F$43=MOD($C$157+3,5))),DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&"
    #    t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!F$42=MOD($C$157+1,5),DAI_VAN!F$43=MOD($C$157+1,5),DAI_VAN!F$42=MOD($C$157+4,5),DAI_VAN!F$43=MOD($C$157+
    #   4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!G$35<>"",DAI_VAN!G$36<=90,OR(DAI_VAN!G$42=MOD($C$157+3,5),DAI_VAN!G$43=MOD($C$157+3,5))),DAI_VAN
    #   !G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!G$42=MOD($C$157+1,5),DAI_VAN!G$43=MOD($C$157+1,5),DAI_VAN
    #   !G$42=MOD($C$157+4,5),DAI_VAN!G$43=MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!H$35<>"",DAI_VAN!H$36<=90,OR(DAI_VAN!H$42=MOD($C$1
    #   57+3,5),DAI_VAN!H$43=MOD($C$157+3,5))),DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!H$42=MOD($C
    #   $157+1,5),DAI_VAN!H$43=MOD($C$157+1,5),DAI_VAN!H$42=MOD($C$157+4,5),DAI_VAN!H$43=MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!I$35
    #   <>"",DAI_VAN!I$36<=90,OR(DAI_VAN!I$42=MOD($C$157+3,5),DAI_VAN!I$43=MOD($C$157+3,5))),DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&" t.)"&IF(OR(
    #   $D$173>0,$E$173>0,$F$173=1,DAI_VAN!I$42=MOD($C$157+1,5),DAI_VAN!I$43=MOD($C$157+1,5),DAI_VAN!I$42=MOD($C$157+4,5),DAI_VAN!I$43=MOD($C$157+4,5))," có c
    #   ứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!J$35<>"",DAI_VAN!J$36<=90,OR(DAI_VAN!J$42=MOD($C$157+3,5),DAI_VAN!J$43=MOD($C$157+3,5))),DAI_VAN!J$35&" ("&D
    #   AI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!J$42=MOD($C$157+1,5),DAI_VAN!J$43=MOD($C$157+1,5),DAI_VAN!J$42=MOD($C
    #   $157+4,5),DAI_VAN!J$43=MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!K$35<>"",DAI_VAN!K$36<=90,OR(DAI_VAN!K$42=MOD($C$157+3,5),DAI_
    #   VAN!K$43=MOD($C$157+3,5))),DAI_VAN!K$35&" ("&DAI_VAN!K$36&"–"&(DAI_VAN!K$36+9)&" t.)"&IF(OR($D$173>0,$E$173>0,$F$173=1,DAI_VAN!K$42=MOD($C$157+1,5),DA
    #   I_VAN!K$43=MOD($C$157+1,5),DAI_VAN!K$42=MOD($C$157+4,5),DAI_VAN!K$43=MOD($C$157+4,5))," có cứu"," KHÔNG cứu")&"; ","")&IF(AND(DAI_VAN!L$35<>"",DAI_VAN
    #   !L$36<=90,OR(DAI_VAN!L$42=MOD($C$157+3,5),DAI_VAN!L$43=MOD($C$157+3,5))),DAI_VAN!L$35&" ("&DAI_VAN!L$36&"–"&(DAI_VAN!L$36+9)&" t.)"&IF(OR($D$173>0,$E$
    #   173>0,$F$173=1,DAI_VAN!L$42=MOD($C$157+1,5),DAI_VAN!L$43=MOD($C$157+1,5),DAI_VAN!L$42=MOD($C$157+4,5),DAI_VAN!L$43=MOD($C$157+4,5))," có cứu"," KHÔNG 
    #   cứu")&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-8),' ('),E.v(14,36,c-8)),'–'),X.add(E.v(14,36,c-8),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-8),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-8),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-8),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-8),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-8),''), X.le(E.v(14,36,c-8),90.0), X.OR(X.eq(E.v(14,42,c-8),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-8),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-7),' ('),E.v(14,36,c-7)),'–'),X.add(E.v(14,36,c-7),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-7),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-7),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-7),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-7),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-7),''), X.le(E.v(14,36,c-7),90.0), X.OR(X.eq(E.v(14,42,c-7),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-7),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-6),' ('),E.v(14,36,c-6)),'–'),X.add(E.v(14,36,c-6),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-6),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-6),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-6),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-6),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-6),''), X.le(E.v(14,36,c-6),90.0), X.OR(X.eq(E.v(14,42,c-6),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-6),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-5),' ('),E.v(14,36,c-5)),'–'),X.add(E.v(14,36,c-5),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-5),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-5),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-5),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-5),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-5),''), X.le(E.v(14,36,c-5),90.0), X.OR(X.eq(E.v(14,42,c-5),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-5),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-4),' ('),E.v(14,36,c-4)),'–'),X.add(E.v(14,36,c-4),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-4),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-4),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-4),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-4),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-4),''), X.le(E.v(14,36,c-4),90.0), X.OR(X.eq(E.v(14,42,c-4),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-4),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-3),' ('),E.v(14,36,c-3)),'–'),X.add(E.v(14,36,c-3),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-3),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-3),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-3),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-3),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-3),''), X.le(E.v(14,36,c-3),90.0), X.OR(X.eq(E.v(14,42,c-3),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-3),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-2),' ('),E.v(14,36,c-2)),'–'),X.add(E.v(14,36,c-2),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-2),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-2),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-2),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-2),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-2),''), X.le(E.v(14,36,c-2),90.0), X.OR(X.eq(E.v(14,42,c-2),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-2),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-1),' ('),E.v(14,36,c-1)),'–'),X.add(E.v(14,36,c-1),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c-1),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c-1),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c-1),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c-1),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-1),''), X.le(E.v(14,36,c-1),90.0), X.OR(X.eq(E.v(14,42,c-1),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c-1),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c),''), X.le(E.v(14,36,c),90.0), X.OR(X.eq(E.v(14,42,c),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),' t.)'),(' có cứu' if X.truthy(X.OR(X.gt(E.v(8,173,4),0.0), X.gt(E.v(8,173,5),0.0), X.eq(E.v(8,173,6),1.0), X.eq(E.v(14,42,c+1),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,43,c+1),X.MOD(X.add(E.v(8,157,3),1.0), 5.0)), X.eq(E.v(14,42,c+1),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)), X.eq(E.v(14,43,c+1),X.MOD(X.add(E.v(8,157,3),4.0), 5.0)))) else ' KHÔNG cứu')),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c+1),''), X.le(E.v(14,36,c+1),90.0), X.OR(X.eq(E.v(14,42,c+1),X.MOD(X.add(E.v(8,157,3),3.0), 5.0)), X.eq(E.v(14,43,c+1),X.MOD(X.add(E.v(8,157,3),3.0), 5.0))))) else '')))

def f244(E, r, c):
    # Ô C176
    # = IF($C$160<>"",-1,IF($C$155=2,0,IF(OR(AND(PP2_DINH_LUONG!$C$101>=0.6,NOT(IF($C$154=2,$D$7,$C$7)=0)),AND(PP2_DINH_LUONG!$C$101<0.3,IF($C$154=2,$D$7,$C$7
    #   )=0),IF($C$154=2,$D$6,$C$6)="Quá vượng",IF($C$154=2,$D$6,$C$6)="Quá nhược",IF($C$154=2,$D$6,$C$6)="Độc vượng"),2,1)))
    return (X.neg(1.0) if X.truthy(X.ne(E.v(8,160,3),'')) else (0.0 if X.truthy(X.eq(E.v(8,155,3),2.0)) else (2.0 if X.truthy(X.OR(X.AND(X.ge(E.v(10,101,3),0.6), X.NOT(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,7,3)),0.0))), X.AND(X.lt(E.v(10,101,3),0.3), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,7,3)),0.0)), X.eq((E.v(8,6,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,6,3)),'Quá vượng'), X.eq((E.v(8,6,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,6,3)),'Quá nhược'), X.eq((E.v(8,6,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,6,3)),'Độc vượng'))) else 1.0)))

def f245(E, r, c):
    # Ô C177
    # = IF(OR($C$160<>"",$C$156<0),-1,IF(IF($C$154=2,$D$97,$C$97)=1,0,IF(AND(OR($C$157=IF($C$154=2,$D$8,$C$8),AND(COUNTIF(TRUNG_GIAN!$C$7:$D$7,$C$157)+IF(TRUN
    #   G_GIAN!$F$7=$C$157,1,0)>0,SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=$C$157)*(PP2_DINH_LUONG!$C$24:$F$24<0))+COUNTIF(PP2_DINH_LUONG!$C$24:$F$24,$C$157)>0)),OR
    #   ($C$173=0,$D$173+$E$173+$F$173>0)),2,1)))
    return (X.neg(1.0) if X.truthy(X.OR(X.ne(E.v(8,160,3),''), X.lt(E.v(8,156,3),0.0))) else (0.0 if X.truthy(X.eq((E.v(8,97,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,97,3)),1.0)) else (2.0 if X.truthy(X.AND(X.OR(X.eq(E.v(8,157,3),(E.v(8,8,4) if X.truthy(X.eq(E.v(8,154,3),2.0)) else E.v(8,8,3))), X.AND(X.gt(X.add(X.COUNTIF(E.rg(23,7,3,7,4), E.v(8,157,3)),(1.0 if X.truthy(X.eq(E.v(23,7,6),E.v(8,157,3))) else 0.0)),0.0), X.gt(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),E.v(8,157,3)),X.A('lt',E.rg(10,24,3,24,6),0.0))),X.COUNTIF(E.rg(10,24,3,24,6), E.v(8,157,3))),0.0))), X.OR(X.eq(E.v(8,173,3),0.0), X.gt(X.add(X.add(E.v(8,173,4),E.v(8,173,5)),E.v(8,173,6)),0.0)))) else 1.0)))

def f246(E, r, c):
    # Ô C178
    # = IF($C$176<0,"Cách đặc biệt "&$C$160&": xét theo cách (P2 tr.375-386).","BỆNH: "&CHOOSE($C$176+1,"không rõ bệnh – tứ trụ gần trung hòa","NHẸ","NẶNG")&"
    #    (phe mình PP2 "&TEXT(PP2_DINH_LUONG!$C$101,"0%")&IF(AND($C$176=1,OR(PP2_DINH_LUONG!$C$101>=0.6,PP2_DINH_LUONG!$C$101<0.3))," – số độ PP2 lệch ngược c
    #   hiều kết luận vượng suy cuối nên chỉ xét bệnh nhẹ","")&"). THUỐC – dụng thần "&INDEX(HANH_TEN,MAX(0,$C$157)+1)&": "&CHOOSE(MAX(0,$C$177)+1,"KHUYẾT tro
    #   ng tứ trụ","có nhưng YẾU hoặc bị thương","CÓ LỰC")&". → "&IF($C$176=0,"Không bệnh không thuốc: đời bình thường, ít thăng trầm lớn.",IF($C$177=2,IF($C$
    #   176=2,"Bệnh nặng được thuốc: tiềm năng ĐẠI phú quý.","Bệnh nhẹ được thuốc: tiềm năng TIỂU phú quý."),IF($C$177=1,"Thuốc nhẹ: phải nhờ đại vận, lưu niê
    #   n mang dụng thần bổ cứu mới cân bằng.","Có bệnh mà không có thuốc: phải trông vào đại vận, lưu niên mang dụng thần; bổ sung bằng nhân tạo (nghề nghiệp
    #   , phương vị nơi ở, tên gọi theo hành dụng thần).")))&" (P2 tr.376-378, dẫn Mệnh lý thám nguyên)")
    return (X.cat(X.cat('Cách đặc biệt ',E.v(8,160,3)),': xét theo cách (P2 tr.375-386).') if X.truthy(X.lt(E.v(8,176,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('BỆNH: ',X.CHOOSE(X.add(E.v(8,176,3),1.0), lambda: 'không rõ bệnh – tứ trụ gần trung hòa', lambda: 'NHẸ', lambda: 'NẶNG')),' (phe mình PP2 '),X.TEXT(E.v(10,101,3), '0%')),(' – số độ PP2 lệch ngược chiều kết luận vượng suy cuối nên chỉ xét bệnh nhẹ' if X.truthy(X.AND(X.eq(E.v(8,176,3),1.0), X.OR(X.ge(E.v(10,101,3),0.6), X.lt(E.v(10,101,3),0.3)))) else '')),'). THUỐC – dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MAX(0.0, E.rg(8,157,3,157,3)),1.0))),': '),X.CHOOSE(X.add(X.MAX(0.0, E.rg(8,177,3,177,3)),1.0), lambda: 'KHUYẾT trong tứ trụ', lambda: 'có nhưng YẾU hoặc bị thương', lambda: 'CÓ LỰC')),'. → '),('Không bệnh không thuốc: đời bình thường, ít thăng trầm lớn.' if X.truthy(X.eq(E.v(8,176,3),0.0)) else (('Bệnh nặng được thuốc: tiềm năng ĐẠI phú quý.' if X.truthy(X.eq(E.v(8,176,3),2.0)) else 'Bệnh nhẹ được thuốc: tiềm năng TIỂU phú quý.') if X.truthy(X.eq(E.v(8,177,3),2.0)) else ('Thuốc nhẹ: phải nhờ đại vận, lưu niên mang dụng thần bổ cứu mới cân bằng.' if X.truthy(X.eq(E.v(8,177,3),1.0)) else 'Có bệnh mà không có thuốc: phải trông vào đại vận, lưu niên mang dụng thần; bổ sung bằng nhân tạo (nghề nghiệp, phương vị nơi ở, tên gọi theo hành dụng thần).')))),' (P2 tr.376-378, dẫn Mệnh lý thám nguyên)'))

def f247(E, r, c):
    # Ô C179
    # = "PP1 (đếm chữ, trừ nhật can): phe sinh trợ "&($C$12+$C$16)&" chữ – phe khắc tiết hao "&($C$13+$C$14+$C$15)&" chữ. PP2: phe mình "&TEXT(PP2_DINH_LUONG!
    #   $C$99,"0.0")&" – phe kia "&TEXT(PP2_DINH_LUONG!$C$100-PP2_DINH_LUONG!$C$99,"0.0")&" độ ("&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&"). "&IF(AND(PP2_DINH_LUO
    #   NG!$C$101>=0.35,PP2_DINH_LUONG!$C$101<0.45),"⚠ SÁT NGƯỠNG 40%: hai phe gần cân – cần luận giải thêm (gần–xa, sinh–khắc, hình–hợp–xung–hại). ","")&IF(O
    #   R(AND($C$7=0,$C$12+$C$16>$C$13+$C$14+$C$15),AND($C$7=1,$C$12+$C$16<$C$13+$C$14+$C$15)),"⚠ PP1 kết luận ngược tỉ lệ chữ (được lệnh/đắc địa quyết định) 
    #   – cần luận giải. ","")&IF(AND(NOT(AND(PP2_DINH_LUONG!$C$101>=0.35,PP2_DINH_LUONG!$C$101<0.45)),NOT(OR(AND($C$7=0,$C$12+$C$16>$C$13+$C$14+$C$15),AND($C
    #   $7=1,$C$12+$C$16<$C$13+$C$14+$C$15)))),"Hai phe chênh rõ, phù hợp kết luận. ","")&"(P1 tr.179 bước 7; mốc 35–45% là quy ước kỹ thuật)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 (đếm chữ, trừ nhật can): phe sinh trợ ',X.add(E.v(8,12,3),E.v(8,16,3))),' chữ – phe khắc tiết hao '),X.add(X.add(E.v(8,13,3),E.v(8,14,3)),E.v(8,15,3))),' chữ. PP2: phe mình '),X.TEXT(E.v(10,99,3), '0.0')),' – phe kia '),X.TEXT(X.sub(E.v(10,100,3),E.v(10,99,3)), '0.0')),' độ ('),X.TEXT(E.v(10,101,3), '0.0%')),'). '),('⚠ SÁT NGƯỠNG 40%: hai phe gần cân – cần luận giải thêm (gần–xa, sinh–khắc, hình–hợp–xung–hại). ' if X.truthy(X.AND(X.ge(E.v(10,101,3),0.35), X.lt(E.v(10,101,3),0.45))) else '')),('⚠ PP1 kết luận ngược tỉ lệ chữ (được lệnh/đắc địa quyết định) – cần luận giải. ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),0.0), X.gt(X.add(E.v(8,12,3),E.v(8,16,3)),X.add(X.add(E.v(8,13,3),E.v(8,14,3)),E.v(8,15,3)))), X.AND(X.eq(E.v(8,7,3),1.0), X.lt(X.add(E.v(8,12,3),E.v(8,16,3)),X.add(X.add(E.v(8,13,3),E.v(8,14,3)),E.v(8,15,3)))))) else '')),('Hai phe chênh rõ, phù hợp kết luận. ' if X.truthy(X.AND(X.NOT(X.AND(X.ge(E.v(10,101,3),0.35), X.lt(E.v(10,101,3),0.45))), X.NOT(X.OR(X.AND(X.eq(E.v(8,7,3),0.0), X.gt(X.add(E.v(8,12,3),E.v(8,16,3)),X.add(X.add(E.v(8,13,3),E.v(8,14,3)),E.v(8,15,3)))), X.AND(X.eq(E.v(8,7,3),1.0), X.lt(X.add(E.v(8,12,3),E.v(8,16,3)),X.add(X.add(E.v(8,13,3),E.v(8,14,3)),E.v(8,15,3)))))))) else '')),'(P1 tr.179 bước 7; mốc 35–45% là quy ước kỹ thuật)')

def f248(E, r, c):
    # Ô C180
    # = IF(NOT(ISNUMBER(DAI_VAN!$C$11)),"Chưa nhập ngày giờ sinh dương lịch (NHAP_LIEU C26:G26) – không kiểm tra được ranh giới.",IF($J$180+$K$180=0,"✔ Giờ si
    #   nh không sát ranh giới canh giờ; ngày sinh cách giao tiết hơn 1 ngày.",IF($J$180=1,"⚠ Giờ sinh cách ranh giới canh giờ (giờ lẻ) không quá 10 phút: trụ
    #    giờ có thể là giờ kề – nên xét cả hai."&IF(NHAP_LIEU!$F$26>=22," Sau 23 giờ là giờ Tý của ngày hôm sau (P1 tr.45).","")&" ","")&IF($K$180=1,"⚠ Sinh c
    #   ách giao tiết không quá 1 ngày ("&DAI_VAN!$C$12&" | "&DAI_VAN!$C$13&"): trụ tháng (và trụ năm nếu là Lập xuân) đổi nếu giờ sinh lệch – kiểm tra lại gi
    #   ờ sinh và múi giờ.","")))&" (P1 tr.178 bước 1; ngưỡng 10 phút, 1 ngày là quy ước kỹ thuật)"
    return X.cat(('Chưa nhập ngày giờ sinh dương lịch (NHAP_LIEU C26:G26) – không kiểm tra được ranh giới.' if X.truthy(X.NOT(X.ISNUMBER(lambda: E.v(14,11,3)))) else ('✔ Giờ sinh không sát ranh giới canh giờ; ngày sinh cách giao tiết hơn 1 ngày.' if X.truthy(X.eq(X.add(E.v(8,180,10),E.v(8,180,11)),0.0)) else X.cat((X.cat(X.cat('⚠ Giờ sinh cách ranh giới canh giờ (giờ lẻ) không quá 10 phút: trụ giờ có thể là giờ kề – nên xét cả hai.',(' Sau 23 giờ là giờ Tý của ngày hôm sau (P1 tr.45).' if X.truthy(X.ge(E.v(2,26,6),22.0)) else '')),' ') if X.truthy(X.eq(E.v(8,180,10),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat('⚠ Sinh cách giao tiết không quá 1 ngày (',E.v(14,12,3)),' | '),E.v(14,13,3)),'): trụ tháng (và trụ năm nếu là Lập xuân) đổi nếu giờ sinh lệch – kiểm tra lại giờ sinh và múi giờ.') if X.truthy(X.eq(E.v(8,180,11),1.0)) else '')))),' (P1 tr.178 bước 1; ngưỡng 10 phút, 1 ngày là quy ước kỹ thuật)')

def f249(E, r, c):
    # Ô J180
    # = IF(AND(NHAP_LIEU!$C$9<>"Không rõ giờ",ISNUMBER(NHAP_LIEU!$F$26)),IF(MIN(MOD(NHAP_LIEU!$F$26*60+N(NHAP_LIEU!$G$26)-60,120),120-MOD(NHAP_LIEU!$F$26*60+N
    #   (NHAP_LIEU!$G$26)-60,120))<=10,1,0),0)
    return ((1.0 if X.truthy(X.le(X.MIN(X.MOD(X.sub(X.add(X.mul(E.v(2,26,6),60.0),X.N(E.v(2,26,7))),60.0), 120.0), X.sub(120.0,X.MOD(X.sub(X.add(X.mul(E.v(2,26,6),60.0),X.N(E.v(2,26,7))),60.0), 120.0))),10.0)) else 0.0) if X.truthy(X.AND(X.ne(E.v(2,9,3),'Không rõ giờ'), X.ISNUMBER(lambda: E.v(2,26,6)))) else 0.0)

def f250(E, r, c):
    # Ô K180
    # = IF(ISNUMBER(DAI_VAN!$C$11),IF(MIN(INDEX(TIET_UTC,DAI_VAN!$C$11+1)-DAI_VAN!$C$10,DAI_VAN!$C$10-INDEX(TIET_UTC,DAI_VAN!$C$11))<=1,1,0),0)
    return ((1.0 if X.truthy(X.le(X.MIN(X.sub(X.INDEX(E.rg(25,4,1,2439,1), X.add(E.v(14,11,3),1.0)),E.v(14,10,3)), X.sub(E.v(14,10,3),X.INDEX(E.rg(25,4,1,2439,1), E.v(14,11,3)))),1.0)) else 0.0) if X.truthy(X.ISNUMBER(lambda: E.v(14,11,3))) else 0.0)

def f251(E, r, c):
    # Ô C181
    # = "Hàn – ôn: "&IF(OR($C$162=4,IF(PP2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$92/PP2_DINH_LUONG!$C$100,0)>=0.4),"tứ trụ ôn táo → thiên nhiệt tình, cởi mở, 
    #   dễ hòa hợp với mọi người",IF(OR($C$162=1,IF(PP2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$95/PP2_DINH_LUONG!$C$100,0)>=0.4),"tứ trụ hàn thấp → thiên điềm 
    #   đạm, lãnh đạm, giữ khoảng cách","nóng lạnh vừa phải → hòa nhã, chừng mực"))&". Sinh hóa: "&$J$181&" cặp kề nhau tương sinh, "&$K$181&" cặp tương khắc 
    #   → "&IF($J$181>$K$181,"sinh hóa có tình: trọng tình nghĩa",IF($J$181<$K$181,"khắc nhiều hơn sinh: tình cảm khô, dễ va chạm","tình nghĩa vừa phải"))&". 
    #   Cương – nhu: "&IF(DOI_CHIEU!$C$33=1,IF(DOI_CHIEU!$C$31="vượng","thân vượng (theo PP1) → cương, quyết đoán, dễ cố chấp",IF(DOI_CHIEU!$C$31="nhược","thâ
    #   n nhược (theo PP1) → nhu, mềm mỏng, dễ nhu nhược","tương đối trung chính"))&IF(DOI_CHIEU!$C$31<>DOI_CHIEU!$C$32," (PP2 đọc ngược lại)",""),IF(PP2_DINH
    #   _LUONG!$C$101>=0.6,"thân quá vượng → cương, quyết đoán, dễ cố chấp",IF(PP2_DINH_LUONG!$C$101<0.25,"thân quá nhược → nhu, mềm mỏng, dễ nhu nhược","tươn
    #   g đối trung chính")))&IF($L$181>=2,"; địa chi "&$L$181&" cặp xung → tổ hợp xấu nhiều, thiên khô, cần tu dưỡng","")&". (Tiêu chí theo P1 tr.181; cách l
    #   ượng hóa là quy ước kỹ thuật)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Hàn – ôn: ',('tứ trụ ôn táo → thiên nhiệt tình, cởi mở, dễ hòa hợp với mọi người' if X.truthy(X.OR(X.eq(E.v(8,162,3),4.0), X.ge((X.div(E.v(10,92,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.4))) else ('tứ trụ hàn thấp → thiên điềm đạm, lãnh đạm, giữ khoảng cách' if X.truthy(X.OR(X.eq(E.v(8,162,3),1.0), X.ge((X.div(E.v(10,95,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.4))) else 'nóng lạnh vừa phải → hòa nhã, chừng mực'))),'. Sinh hóa: '),E.v(8,181,10)),' cặp kề nhau tương sinh, '),E.v(8,181,11)),' cặp tương khắc → '),('sinh hóa có tình: trọng tình nghĩa' if X.truthy(X.gt(E.v(8,181,10),E.v(8,181,11))) else ('khắc nhiều hơn sinh: tình cảm khô, dễ va chạm' if X.truthy(X.lt(E.v(8,181,10),E.v(8,181,11))) else 'tình nghĩa vừa phải'))),'. Cương – nhu: '),(X.cat(('thân vượng (theo PP1) → cương, quyết đoán, dễ cố chấp' if X.truthy(X.eq(E.v(7,31,3),'vượng')) else ('thân nhược (theo PP1) → nhu, mềm mỏng, dễ nhu nhược' if X.truthy(X.eq(E.v(7,31,3),'nhược')) else 'tương đối trung chính')),(' (PP2 đọc ngược lại)' if X.truthy(X.ne(E.v(7,31,3),E.v(7,32,3))) else '')) if X.truthy(X.eq(E.v(7,33,3),1.0)) else ('thân quá vượng → cương, quyết đoán, dễ cố chấp' if X.truthy(X.ge(E.v(10,101,3),0.6)) else ('thân quá nhược → nhu, mềm mỏng, dễ nhu nhược' if X.truthy(X.lt(E.v(10,101,3),0.25)) else 'tương đối trung chính')))),(X.cat(X.cat('; địa chi ',E.v(8,181,12)),' cặp xung → tổ hợp xấu nhiều, thiên khô, cần tu dưỡng') if X.truthy(X.ge(E.v(8,181,12),2.0)) else '')),'. (Tiêu chí theo P1 tr.181; cách lượng hóa là quy ước kỹ thuật)')

def f252(E, r, c):
    # Ô J181
    # = IF(OR(MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$C$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$E$7,MOD(TRUNG
    #   _GIAN!$E$7+1,5)=TRUNG_GIAN!$D$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$F$7,MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7),1,0)+IF(OR(MOD(TRUNG_GI
    #   AN!$C$7+1,5)=TRUNG_GIAN!$C$9,MOD(TRUNG_GIAN!$C$9+1,5)=TRUNG_GIAN!$C$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$D$9,MOD(TRUNG_GIAN!$D$9+1,5)=TR
    #   UNG_GIAN!$D$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$E$9,MOD(TRUNG_GIAN!$E$9+1,5)=TRUNG_GIAN!$E$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG
    #   _GIAN!$F$9,MOD(TRUNG_GIAN!$F$9+1,5)=TRUNG_GIAN!$F$7),1,0)
    return X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,3)))) else 0.0),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,7,4)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,7,6)), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,9,3)), X.eq(X.MOD(X.add(E.v(23,9,3),1.0), 5.0),E.v(23,7,3)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,9,4)), X.eq(X.MOD(X.add(E.v(23,9,4),1.0), 5.0),E.v(23,7,4)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,9,5)), X.eq(X.MOD(X.add(E.v(23,9,5),1.0), 5.0),E.v(23,7,5)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,9,6)), X.eq(X.MOD(X.add(E.v(23,9,6),1.0), 5.0),E.v(23,7,6)))) else 0.0))

def f253(E, r, c):
    # Ô K181
    # = IF(OR(MOD(TRUNG_GIAN!$C$7+2,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+2,5)=TRUNG_GIAN!$C$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$D$7+2,5)=TRUNG_GIAN!$E$7,MOD(TRUNG
    #   _GIAN!$E$7+2,5)=TRUNG_GIAN!$D$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$F$7,MOD(TRUNG_GIAN!$F$7+2,5)=TRUNG_GIAN!$E$7),1,0)+IF(OR(MOD(TRUNG_GI
    #   AN!$C$7+2,5)=TRUNG_GIAN!$C$9,MOD(TRUNG_GIAN!$C$9+2,5)=TRUNG_GIAN!$C$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$D$7+2,5)=TRUNG_GIAN!$D$9,MOD(TRUNG_GIAN!$D$9+2,5)=TR
    #   UNG_GIAN!$D$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$E$9,MOD(TRUNG_GIAN!$E$9+2,5)=TRUNG_GIAN!$E$7),1,0)+IF(OR(MOD(TRUNG_GIAN!$F$7+2,5)=TRUNG
    #   _GIAN!$F$9,MOD(TRUNG_GIAN!$F$9+2,5)=TRUNG_GIAN!$F$7),1,0)
    return X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(23,7,3)))) else 0.0),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,7,4)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,7,6)), X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(23,7,5)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(23,9,3)), X.eq(X.MOD(X.add(E.v(23,9,3),2.0), 5.0),E.v(23,7,3)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(23,9,4)), X.eq(X.MOD(X.add(E.v(23,9,4),2.0), 5.0),E.v(23,7,4)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,9,5)), X.eq(X.MOD(X.add(E.v(23,9,5),2.0), 5.0),E.v(23,7,5)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(23,9,6)), X.eq(X.MOD(X.add(E.v(23,9,6),2.0), 5.0),E.v(23,7,6)))) else 0.0))

def f254(E, r, c):
    # Ô L181
    # = INDEX(M_XU,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$D$6+1)+INDEX(M_XU,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$E$6+1)+INDEX(M_XU,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$F$6+1)+INDEX(
    #   M_XU,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$E$6+1)+INDEX(M_XU,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$F$6+1)+INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)
    return X.add(X.add(X.add(X.add(X.add(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,4),1.0)),X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,5),1.0))),X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,6),1.0))),X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,5),1.0))),X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,6),1.0))),X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)))

def f255(E, r, c):
    # Ô C182
    # = "10 năm tới (từ năm xem) – TỐT NHẤT: "&IF(IF(IFERROR(LARGE(LUU_NIEN!$DE$30:$DE$129,1),-1)>0,IFERROR(INDEX(LUU_NIEN!$DG$30:$DG$129,MATCH(LARGE(LUU_NIEN
    #   !$DE$30:$DE$129,1),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")="","không có năm tốt rõ",TRIM(IF(IFERROR(LARGE(LUU_NIEN!$DE$30:$DE$129,1),-1)>0,IFERROR(IN
    #   DEX(LUU_NIEN!$DG$30:$DG$129,MATCH(LARGE(LUU_NIEN!$DE$30:$DE$129,1),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")&IF(IFERROR(LARGE(LUU_NIEN!$DE$30:$DE$129,2
    #   ),-1)>0,IFERROR(INDEX(LUU_NIEN!$DG$30:$DG$129,MATCH(LARGE(LUU_NIEN!$DE$30:$DE$129,2),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")&IF(IFERROR(LARGE(LUU_NIE
    #   N!$DE$30:$DE$129,3),-1)>0,IFERROR(INDEX(LUU_NIEN!$DG$30:$DG$129,MATCH(LARGE(LUU_NIEN!$DE$30:$DE$129,3),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")))&" XẤ
    #   U NHẤT: "&IF(IF(IFERROR(SMALL(LUU_NIEN!$DE$30:$DE$129,1),1)<0,IFERROR(INDEX(LUU_NIEN!$DG$30:$DG$129,MATCH(SMALL(LUU_NIEN!$DE$30:$DE$129,1),LUU_NIEN!$D
    #   E$30:$DE$129,0))&"; ",""),"")="","không có năm xấu rõ",TRIM(IF(IFERROR(SMALL(LUU_NIEN!$DE$30:$DE$129,1),1)<0,IFERROR(INDEX(LUU_NIEN!$DG$30:$DG$129,MAT
    #   CH(SMALL(LUU_NIEN!$DE$30:$DE$129,1),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")&IF(IFERROR(SMALL(LUU_NIEN!$DE$30:$DE$129,2),1)<0,IFERROR(INDEX(LUU_NIEN!$
    #   DG$30:$DG$129,MATCH(SMALL(LUU_NIEN!$DE$30:$DE$129,2),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")&IF(IFERROR(SMALL(LUU_NIEN!$DE$30:$DE$129,3),1)<0,IFERROR
    #   (INDEX(LUU_NIEN!$DG$30:$DG$129,MATCH(SMALL(LUU_NIEN!$DE$30:$DE$129,3),LUU_NIEN!$DE$30:$DE$129,0))&"; ",""),"")))&" Cả đời (100 năm) – tốt nhất: "&IFER
    #   ROR(INDEX(LUU_NIEN!$DG$30:$DG$129,MATCH(MAX(LUU_NIEN!$DD$30:$DD$129),LUU_NIEN!$DD$30:$DD$129,0)),"")&"; xấu nhất: "&IFERROR(INDEX(LUU_NIEN!$DG$30:$DG$
    #   129,MATCH(MIN(LUU_NIEN!$DD$30:$DD$129),LUU_NIEN!$DD$30:$DD$129,0)),"")&". (Xếp theo xu thế năm của phương pháp khuyến nghị, trừ điểm năm hạn nặng; việ
    #   c ứng theo P1 tr.98; P1 tr.180 bước 11)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('10 năm tới (từ năm xem) – TỐT NHẤT: ',('không có năm tốt rõ' if X.truthy(X.eq((X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.LARGE(E.rg(15,30,109,129,109), 1.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.gt(X.IFERROR(lambda: X.LARGE(E.rg(15,30,109,129,109), 1.0), lambda: X.neg(1.0)),0.0)) else ''),'')) else X.TRIM(X.cat(X.cat((X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.LARGE(E.rg(15,30,109,129,109), 1.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.gt(X.IFERROR(lambda: X.LARGE(E.rg(15,30,109,129,109), 1.0), lambda: X.neg(1.0)),0.0)) else ''),(X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.LARGE(E.rg(15,30,109,129,109), 2.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.gt(X.IFERROR(lambda: X.LARGE(E.rg(15,30,109,129,109), 2.0), lambda: X.neg(1.0)),0.0)) else '')),(X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.LARGE(E.rg(15,30,109,129,109), 3.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.gt(X.IFERROR(lambda: X.LARGE(E.rg(15,30,109,129,109), 3.0), lambda: X.neg(1.0)),0.0)) else ''))))),' XẤU NHẤT: '),('không có năm xấu rõ' if X.truthy(X.eq((X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.SMALL(E.rg(15,30,109,129,109), 1.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.lt(X.IFERROR(lambda: X.SMALL(E.rg(15,30,109,129,109), 1.0), lambda: 1.0),0.0)) else ''),'')) else X.TRIM(X.cat(X.cat((X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.SMALL(E.rg(15,30,109,129,109), 1.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.lt(X.IFERROR(lambda: X.SMALL(E.rg(15,30,109,129,109), 1.0), lambda: 1.0),0.0)) else ''),(X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.SMALL(E.rg(15,30,109,129,109), 2.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.lt(X.IFERROR(lambda: X.SMALL(E.rg(15,30,109,129,109), 2.0), lambda: 1.0),0.0)) else '')),(X.IFERROR(lambda: X.cat(X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.SMALL(E.rg(15,30,109,129,109), 3.0), E.rg(15,30,109,129,109), 0.0)),'; '), lambda: '') if X.truthy(X.lt(X.IFERROR(lambda: X.SMALL(E.rg(15,30,109,129,109), 3.0), lambda: 1.0),0.0)) else ''))))),' Cả đời (100 năm) – tốt nhất: '),X.IFERROR(lambda: X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.MAX(E.rg(15,30,108,129,108)), E.rg(15,30,108,129,108), 0.0)), lambda: '')),'; xấu nhất: '),X.IFERROR(lambda: X.INDEX(E.rg(15,30,111,129,111), X.MATCH(X.MIN(E.rg(15,30,108,129,108)), E.rg(15,30,108,129,108), 0.0)), lambda: '')),'. (Xếp theo xu thế năm của phương pháp khuyến nghị, trừ điểm năm hạn nặng; việc ứng theo P1 tr.98; P1 tr.180 bước 11)')

def f256(E, r, c):
    # Ô C183
    # = IF(C141&C142&C143&C144="","KHÔNG CẦN thông quan: không có cặp hành tương tranh ngang nhau (P1 tr.100, 110-111).",TRIM(C141&" "&C142&" "&C143&" "&C144)
    #   )
    return ('KHÔNG CẦN thông quan: không có cặp hành tương tranh ngang nhau (P1 tr.100, 110-111).' if X.truthy(X.eq(X.cat(X.cat(X.cat(E.v(8,r-42,c),E.v(8,r-41,c)),E.v(8,r-40,c)),E.v(8,r-39,c)),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(8,r-42,c),' '),E.v(8,r-41,c)),' '),E.v(8,r-40,c)),' '),E.v(8,r-39,c))))

def f257(E, r, c):
    # Ô C184
    # = "Điều hầu: "&$C$163&IF($C$162>=0," ("&INDEX(HANH_TEN,$C$162+1)&")","")&" · Thông quan: "&IF(C141&C142&C143&C144="","KHÔNG CẦN","CẦN XÉT")&" · Cứu ứng:
    #    "&IF(OR($C$160<>"",$C$156<0),"không áp dụng",IF($C$173=0,"KHÔNG CẦN (dụng không bị khắc)",IF($D$173+$E$173+$F$173>0,"CÓ CỨU","KHÔNG CÓ CỨU – chờ vận"
    #   )))&" · Thái quá: "&IF(C151="","không","CÓ – xem dòng 151")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Điều hầu: ',E.v(8,163,3)),(X.cat(X.cat(' (',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,162,3),1.0))),')') if X.truthy(X.ge(E.v(8,162,3),0.0)) else '')),' · Thông quan: '),('KHÔNG CẦN' if X.truthy(X.eq(X.cat(X.cat(X.cat(E.v(8,r-43,c),E.v(8,r-42,c)),E.v(8,r-41,c)),E.v(8,r-40,c)),'')) else 'CẦN XÉT')),' · Cứu ứng: '),('không áp dụng' if X.truthy(X.OR(X.ne(E.v(8,160,3),''), X.lt(E.v(8,156,3),0.0))) else ('KHÔNG CẦN (dụng không bị khắc)' if X.truthy(X.eq(E.v(8,173,3),0.0)) else ('CÓ CỨU' if X.truthy(X.gt(X.add(X.add(E.v(8,173,4),E.v(8,173,5)),E.v(8,173,6)),0.0)) else 'KHÔNG CÓ CỨU – chờ vận')))),' · Thái quá: '),('không' if X.truthy(X.eq(E.v(8,r-33,c),'')) else 'CÓ – xem dòng 151'))

def f258(E, r, c):
    # Ô C185
    # = IF(OR($C$160<>"",$C$156<0),"",IF(SUMPRODUCT((TRUNG_GIAN!$C$21:$F$21=$C$157)*(PP2_DINH_LUONG!$C$32:$F$32>0)*(MOD(TRUNG_GIAN!$C$6:$F$6,3)<>1))>0," Gốc d
    #   ụng thần ở chi bị XUNG mà không có hợp giải (đã xét tham hợp quên xung) → dụng thần thêm tổn thương, gốc không chắc; vận, năm có chi lục hợp hoặc tam 
    #   hợp với chi đó sẽ giải (P1 tr.116, trường hợp 3; xung tứ khố không tính – G80).",""))
    return ('' if X.truthy(X.OR(X.ne(E.v(8,160,3),''), X.lt(E.v(8,156,3),0.0))) else (' Gốc dụng thần ở chi bị XUNG mà không có hợp giải (đã xét tham hợp quên xung) → dụng thần thêm tổn thương, gốc không chắc; vận, năm có chi lục hợp hoặc tam hợp với chi đó sẽ giải (P1 tr.116, trường hợp 3; xung tứ khố không tính – G80).' if X.truthy(X.gt(X.SUMPRODUCT(X.A('mul',X.A('mul',X.A('eq',E.rg(23,21,3,21,6),E.v(8,157,3)),X.A('gt',E.rg(10,32,3,32,6),0.0)),X.A('ne',X.AF_MOD(E.rg(23,6,3,6,6), 3.0),1.0))),0.0)) else ''))

def f259(E, r, c):
    # Ô C186
    # = " [Dữ kiện lá này: "&IF(ISNUMBER(DAI_VAN!$C$11),"sinh "&TEXT(DAI_VAN!$C$10-INDEX(TIET_UTC,DAI_VAN!$C$11),"0")&" ngày sau tiết "&INDEX(TIET_TEN,DAI_VAN
    #   !$C$11)&" → "&IF(DAI_VAN!$C$10-INDEX(TIET_UTC,DAI_VAN!$C$11)<15,"thượng tuần (nửa đầu tháng)","hạ tuần (nửa sau tháng)"),"chưa nhập ngày giờ sinh – kh
    #   ông xác định thượng/hạ tuần")&"; độ PP2: Mộc "&TEXT(IF(PP2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$91/PP2_DINH_LUONG!$C$100,0),"0%")&", Hỏa "&TEXT(IF(PP
    #   2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$92/PP2_DINH_LUONG!$C$100,0),"0%")&", Thổ "&TEXT(IF(PP2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$93/PP2_DINH_LUONG
    #   !$C$100,0),"0%")&", Kim "&TEXT(IF(PP2_DINH_LUONG!$C$100>0,PP2_DINH_LUONG!$F$94/PP2_DINH_LUONG!$C$100,0),"0%")&", Thủy "&TEXT(IF(PP2_DINH_LUONG!$C$100>
    #   0,PP2_DINH_LUONG!$F$95/PP2_DINH_LUONG!$C$100,0),"0%")&"; "&IF(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$45&" "&INDEX(HANH_TEN,TRUNG_GIAN!$C$46+1)&IF(PP2_DINH_L
    #   UONG!$C$21=1," (đã hóa)"," (chưa hóa)"),"không có cục 3 chi")&IF(SUM(TRUNG_GIAN!$C$57:$E$57)>0,"; có bán hợp","")&".]"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' [Dữ kiện lá này: ',(X.cat(X.cat(X.cat(X.cat(X.cat('sinh ',X.TEXT(X.sub(E.v(14,10,3),X.INDEX(E.rg(25,4,1,2439,1), E.v(14,11,3))), '0')),' ngày sau tiết '),X.INDEX(E.rg(25,4,3,2439,3), E.v(14,11,3))),' → '),('thượng tuần (nửa đầu tháng)' if X.truthy(X.lt(X.sub(E.v(14,10,3),X.INDEX(E.rg(25,4,1,2439,1), E.v(14,11,3))),15.0)) else 'hạ tuần (nửa sau tháng)')) if X.truthy(X.ISNUMBER(lambda: E.v(14,11,3))) else 'chưa nhập ngày giờ sinh – không xác định thượng/hạ tuần')),'; độ PP2: Mộc '),X.TEXT((X.div(E.v(10,91,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),', Hỏa '),X.TEXT((X.div(E.v(10,92,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),', Thổ '),X.TEXT((X.div(E.v(10,93,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),', Kim '),X.TEXT((X.div(E.v(10,94,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),', Thủy '),X.TEXT((X.div(E.v(10,95,6),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),'; '),(X.cat(X.cat(X.cat(E.v(23,45,3),' '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,46,3),1.0))),(' (đã hóa)' if X.truthy(X.eq(E.v(10,21,3),1.0)) else ' (chưa hóa)')) if X.truthy(X.gt(E.v(23,44,3),0.0)) else 'không có cục 3 chi')),('; có bán hợp' if X.truthy(X.gt(X.SUM(E.rg(23,57,3,57,5)),0.0)) else '')),'.]')

def f260(E, r, c):
    # Ô J187
    # = IF(INDEX(CSDL!$AM$5:$AM$124,$C$166)=0,FALSE(),CHOOSE(INT(INDEX(CSDL!$AM$5:$AM$124,$C$166)/10),IF(PP2_DINH_LUONG!$C$100>0,INDEX(PP2_DINH_LUONG!$F$91:$F
    #   $95,MOD(INDEX(CSDL!$AM$5:$AM$124,$C$166),10)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4,AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=MOD(INDEX(CSDL!$AM$5:$AM$124,$C$
    #   166),10),PP2_DINH_LUONG!$C$21=1),AND(ISNUMBER(DAI_VAN!$C$11),N(DAI_VAN!$C$10)-INDEX(TIET_UTC,MAX(1,N(DAI_VAN!$C$11)))>=15),IF(PP2_DINH_LUONG!$C$100>0,
    #   (PP2_DINH_LUONG!$F$94+PP2_DINH_LUONG!$F$95)/PP2_DINH_LUONG!$C$100,0)>=0.5))
    return (X.FALSE() if X.truthy(X.eq(X.INDEX(E.rg(24,5,39,124,39), E.v(8,166,3)),0.0)) else X.CHOOSE(X.INT(X.div(X.INDEX(E.rg(24,5,39,124,39), E.v(8,166,3)),10.0)), lambda: X.ge((X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.INDEX(E.rg(24,5,39,124,39), E.v(8,166,3)), 10.0),1.0)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.4), lambda: X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),X.MOD(X.INDEX(E.rg(24,5,39,124,39), E.v(8,166,3)), 10.0)), X.eq(E.v(10,21,3),1.0)), lambda: X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ge(X.sub(X.N(E.v(14,10,3)),X.INDEX(E.rg(25,4,1,2439,1), X.MAX(1.0, X.N(E.v(14,11,3))))),15.0)), lambda: X.ge((X.div(X.add(E.v(10,94,6),E.v(10,95,6)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.5)))

def f261(E, r, c):
    # Ô K187
    # = IF(INDEX(CSDL!$AT$5:$AT$124,$C$166)=0,FALSE(),CHOOSE(INT(INDEX(CSDL!$AT$5:$AT$124,$C$166)/10),IF(PP2_DINH_LUONG!$C$100>0,INDEX(PP2_DINH_LUONG!$F$91:$F
    #   $95,MOD(INDEX(CSDL!$AT$5:$AT$124,$C$166),10)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4,AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=MOD(INDEX(CSDL!$AT$5:$AT$124,$C$
    #   166),10),PP2_DINH_LUONG!$C$21=1),AND(ISNUMBER(DAI_VAN!$C$11),N(DAI_VAN!$C$10)-INDEX(TIET_UTC,MAX(1,N(DAI_VAN!$C$11)))>=15),IF(PP2_DINH_LUONG!$C$100>0,
    #   (PP2_DINH_LUONG!$F$94+PP2_DINH_LUONG!$F$95)/PP2_DINH_LUONG!$C$100,0)>=0.5))
    return (X.FALSE() if X.truthy(X.eq(X.INDEX(E.rg(24,5,46,124,46), E.v(8,166,3)),0.0)) else X.CHOOSE(X.INT(X.div(X.INDEX(E.rg(24,5,46,124,46), E.v(8,166,3)),10.0)), lambda: X.ge((X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.INDEX(E.rg(24,5,46,124,46), E.v(8,166,3)), 10.0),1.0)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.4), lambda: X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),X.MOD(X.INDEX(E.rg(24,5,46,124,46), E.v(8,166,3)), 10.0)), X.eq(E.v(10,21,3),1.0)), lambda: X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ge(X.sub(X.N(E.v(14,10,3)),X.INDEX(E.rg(25,4,1,2439,1), X.MAX(1.0, X.N(E.v(14,11,3))))),15.0)), lambda: X.ge((X.div(X.add(E.v(10,94,6),E.v(10,95,6)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.5)))

def f262(E, r, c):
    # Ô L187
    # = IF(INDEX(CSDL!$BA$5:$BA$124,$C$166)=0,FALSE(),CHOOSE(INT(INDEX(CSDL!$BA$5:$BA$124,$C$166)/10),IF(PP2_DINH_LUONG!$C$100>0,INDEX(PP2_DINH_LUONG!$F$91:$F
    #   $95,MOD(INDEX(CSDL!$BA$5:$BA$124,$C$166),10)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4,AND(TRUNG_GIAN!$C$44>0,TRUNG_GIAN!$C$46=MOD(INDEX(CSDL!$BA$5:$BA$124,$C$
    #   166),10),PP2_DINH_LUONG!$C$21=1),AND(ISNUMBER(DAI_VAN!$C$11),N(DAI_VAN!$C$10)-INDEX(TIET_UTC,MAX(1,N(DAI_VAN!$C$11)))>=15),IF(PP2_DINH_LUONG!$C$100>0,
    #   (PP2_DINH_LUONG!$F$94+PP2_DINH_LUONG!$F$95)/PP2_DINH_LUONG!$C$100,0)>=0.5))
    return (X.FALSE() if X.truthy(X.eq(X.INDEX(E.rg(24,5,53,124,53), E.v(8,166,3)),0.0)) else X.CHOOSE(X.INT(X.div(X.INDEX(E.rg(24,5,53,124,53), E.v(8,166,3)),10.0)), lambda: X.ge((X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.INDEX(E.rg(24,5,53,124,53), E.v(8,166,3)), 10.0),1.0)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.4), lambda: X.AND(X.gt(E.v(23,44,3),0.0), X.eq(E.v(23,46,3),X.MOD(X.INDEX(E.rg(24,5,53,124,53), E.v(8,166,3)), 10.0)), X.eq(E.v(10,21,3),1.0)), lambda: X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ge(X.sub(X.N(E.v(14,10,3)),X.INDEX(E.rg(25,4,1,2439,1), X.MAX(1.0, X.N(E.v(14,11,3))))),15.0)), lambda: X.ge((X.div(X.add(E.v(10,94,6),E.v(10,95,6)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0),0.5)))

def f263(E, r, c):
    # Ô M187
    # = IF($J$187,1,IF($K$187,2,IF($L$187,3,0)))
    return (1.0 if X.truthy(E.v(8,187,10)) else (2.0 if X.truthy(E.v(8,187,11)) else (3.0 if X.truthy(E.v(8,187,12)) else 0.0)))

def f264(E, r, c):
    # Ô N187
    # = IF($M$187=0,"",CHOOSE(INT(INDEX(CHOOSE($M$187,CSDL!$AM$5:$AM$124,CSDL!$AT$5:$AT$124,CSDL!$BA$5:$BA$124),$C$166)/10),INDEX(HANH_TEN,MOD(INDEX(CHOOSE($M
    #   $187,CSDL!$AM$5:$AM$124,CSDL!$AT$5:$AT$124,CSDL!$BA$5:$BA$124),$C$166),10)+1)&" nhiều (" &TEXT(IF(PP2_DINH_LUONG!$C$100>0,INDEX(PP2_DINH_LUONG!$F$91:$
    #   F$95,MOD(INDEX(CHOOSE($M$187,CSDL!$AM$5:$AM$124,CSDL!$AT$5:$AT$124,CSDL!$BA$5:$BA$124),$C$166),10)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" độ PP2, ngưỡng 
    #   40%)","có cục "&INDEX(HANH_TEN,MOD(INDEX(CHOOSE($M$187,CSDL!$AM$5:$AM$124,CSDL!$AT$5:$AT$124,CSDL!$BA$5:$BA$124),$C$166),10)+1)&" đã hóa","sinh hạ tuầ
    #   n (≥ 15 ngày sau tiết)","Kim + Thủy nhiều (≥ 50% độ PP2)"))
    return ('' if X.truthy(X.eq(E.v(8,187,13),0.0)) else X.CHOOSE(X.INT(X.div(X.INDEX(X.CHOOSE(E.v(8,187,13), lambda: E.rg(24,5,39,124,39), lambda: E.rg(24,5,46,124,46), lambda: E.rg(24,5,53,124,53)), E.v(8,166,3)),10.0)), lambda: X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.INDEX(X.CHOOSE(E.v(8,187,13), lambda: E.rg(24,5,39,124,39), lambda: E.rg(24,5,46,124,46), lambda: E.rg(24,5,53,124,53)), E.v(8,166,3)), 10.0),1.0)),' nhiều ('),X.TEXT((X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.INDEX(X.CHOOSE(E.v(8,187,13), lambda: E.rg(24,5,39,124,39), lambda: E.rg(24,5,46,124,46), lambda: E.rg(24,5,53,124,53)), E.v(8,166,3)), 10.0),1.0)),E.v(10,100,3)) if X.truthy(X.gt(E.v(10,100,3),0.0)) else 0.0), '0%')),' độ PP2, ngưỡng 40%)'), lambda: X.cat(X.cat('có cục ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.INDEX(X.CHOOSE(E.v(8,187,13), lambda: E.rg(24,5,39,124,39), lambda: E.rg(24,5,46,124,46), lambda: E.rg(24,5,53,124,53)), E.v(8,166,3)), 10.0),1.0))),' đã hóa'), lambda: 'sinh hạ tuần (≥ 15 ngày sau tiết)', lambda: 'Kim + Thủy nhiều (≥ 50% độ PP2)'))

def f265(E, r, c):
    # Ô O187
    # = IF(PP2_DINH_LUONG!$C$100>0,MATCH(MAX(PP2_DINH_LUONG!$F$91:$F$95),PP2_DINH_LUONG!$F$91:$F$95,0)-1,-1)
    return (X.sub(X.MATCH(X.MAX(E.rg(10,91,6,95,6)), E.rg(10,91,6,95,6), 0.0),1.0) if X.truthy(X.gt(E.v(10,100,3),0.0)) else X.neg(1.0))

def f266(E, r, c):
    # Ô C188
    # = IF($M$188="","Cách đặc biệt – không xếp mức phú quý theo phù ức.","MỨC PHÚ QUÝ TỔNG HỢP: "&IF($M$188>=1.5,"GIÀU, PHÚ QUÝ LỚN",IF($M$188>=0.5,"KHÁ GIẢ"
    #   ,IF($M$188>-0.5,"TRUNG BÌNH",IF($M$188>-1.5,"VẤT VẢ CẦU TÀI","NGHÈO"))))&" (điểm "&TEXT($M$188,"0.0")&" = trung bình 4 căn cứ: bệnh – thuốc "&IF(ISNUM
    #   BER($J$188),TEXT($J$188,"+0;-0;0"),"–")&" [P2 tr.376-378]; bậc giàu "&IF(ISNUMBER($K$188),TEXT($K$188,"+0;-0;0"),"–")&" [P1 tr.95-96, 164-166]; bảng g
    #   iàu – nghèo "&IF(ISNUMBER($L$188),TEXT($L$188,"+0;-0;0"),"–")&" [P1 tr.173]; tài vận chương 18 "&IF(ISNUMBER($P$188),TEXT($P$188,"+0;-0;0"),"–")&" [P2
    #    tr.520-534]; thang −2…+2, quy ước kỹ thuật)."&IF(AND(COUNT($J$188:$L$188,$P$188)>=2,MAX($J$188:$L$188,$P$188)>0,MIN($J$188:$L$188,$P$188)<0)," Các că
    #   n cứ NGƯỢC CHIỀU: tiền tài có đến nhưng thân khó giữ – giàu hay nghèo tùy vận dụng thần.",""))
    return ('Cách đặc biệt – không xếp mức phú quý theo phù ức.' if X.truthy(X.eq(E.v(8,188,13),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('MỨC PHÚ QUÝ TỔNG HỢP: ',('GIÀU, PHÚ QUÝ LỚN' if X.truthy(X.ge(E.v(8,188,13),1.5)) else ('KHÁ GIẢ' if X.truthy(X.ge(E.v(8,188,13),0.5)) else ('TRUNG BÌNH' if X.truthy(X.gt(E.v(8,188,13),X.neg(0.5))) else ('VẤT VẢ CẦU TÀI' if X.truthy(X.gt(E.v(8,188,13),X.neg(1.5))) else 'NGHÈO'))))),' (điểm '),X.TEXT(E.v(8,188,13), '0.0')),' = trung bình 4 căn cứ: bệnh – thuốc '),(X.TEXT(E.v(8,188,10), '+0;-0;0') if X.truthy(X.ISNUMBER(lambda: E.v(8,188,10))) else '–')),' [P2 tr.376-378]; bậc giàu '),(X.TEXT(E.v(8,188,11), '+0;-0;0') if X.truthy(X.ISNUMBER(lambda: E.v(8,188,11))) else '–')),' [P1 tr.95-96, 164-166]; bảng giàu – nghèo '),(X.TEXT(E.v(8,188,12), '+0;-0;0') if X.truthy(X.ISNUMBER(lambda: E.v(8,188,12))) else '–')),' [P1 tr.173]; tài vận chương 18 '),(X.TEXT(E.v(8,188,16), '+0;-0;0') if X.truthy(X.ISNUMBER(lambda: E.v(8,188,16))) else '–')),' [P2 tr.520-534]; thang −2…+2, quy ước kỹ thuật).'),(' Các căn cứ NGƯỢC CHIỀU: tiền tài có đến nhưng thân khó giữ – giàu hay nghèo tùy vận dụng thần.' if X.truthy(X.AND(X.ge(X.COUNT(E.rg(8,188,10,188,12), E.rg(8,188,16,188,16)),2.0), X.gt(X.MAX(E.rg(8,188,10,188,12), E.rg(8,188,16,188,16)),0.0), X.lt(X.MIN(E.rg(8,188,10,188,12), E.rg(8,188,16,188,16)),0.0))) else '')))

def f267(E, r, c):
    # Ô J188
    # = IF($C$176<0,"",IF($C$176=0,0,IF($C$177=2,IF($C$176=2,2,1),-1)))
    return ('' if X.truthy(X.lt(E.v(8,176,3),0.0)) else (0.0 if X.truthy(X.eq(E.v(8,176,3),0.0)) else ((2.0 if X.truthy(X.eq(E.v(8,176,3),2.0)) else 1.0) if X.truthy(X.eq(E.v(8,177,3),2.0)) else X.neg(1.0))))

def f268(E, r, c):
    # Ô K188
    # = IF(TAI_LOC!$E$16=0,"",TAI_LOC!$E$16-3)
    return ('' if X.truthy(X.eq(E.v(19,16,5),0.0)) else X.sub(E.v(19,16,5),3.0))

def f269(E, r, c):
    # Ô L188
    # = IF(ISNUMBER(SEARCH("TRUNG BÌNH",LEFT(MENH_CUC!$D$18,IFERROR(FIND(CHAR(10),MENH_CUC!$D$18)-1,LEN(MENH_CUC!$D$18))))),0,IF(ISNUMBER(SEARCH("GIÀU",LEFT(M
    #   ENH_CUC!$D$18,IFERROR(FIND(CHAR(10),MENH_CUC!$D$18)-1,LEN(MENH_CUC!$D$18))))),1,IF(ISNUMBER(SEARCH("NGHÈO",LEFT(MENH_CUC!$D$18,IFERROR(FIND(CHAR(10),M
    #   ENH_CUC!$D$18)-1,LEN(MENH_CUC!$D$18))))),-1,0))*IF(ISNUMBER(SEARCH("RÕ RỆT",LEFT(MENH_CUC!$D$18,IFERROR(FIND(CHAR(10),MENH_CUC!$D$18)-1,LEN(MENH_CUC!$
    #   D$18))))),2,IF(ISNUMBER(SEARCH("NGHIÊNG NHẸ",LEFT(MENH_CUC!$D$18,IFERROR(FIND(CHAR(10),MENH_CUC!$D$18)-1,LEN(MENH_CUC!$D$18))))),0.5,1)))
    return (0.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(18,18,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(18,18,4)),1.0), lambda: X.LEN(E.v(18,18,4))))))) else X.mul((1.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('GIÀU', X.LEFT(E.v(18,18,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(18,18,4)),1.0), lambda: X.LEN(E.v(18,18,4))))))) else (X.neg(1.0) if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHÈO', X.LEFT(E.v(18,18,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(18,18,4)),1.0), lambda: X.LEN(E.v(18,18,4))))))) else 0.0)),(2.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('RÕ RỆT', X.LEFT(E.v(18,18,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(18,18,4)),1.0), lambda: X.LEN(E.v(18,18,4))))))) else (0.5 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG NHẸ', X.LEFT(E.v(18,18,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(18,18,4)),1.0), lambda: X.LEN(E.v(18,18,4))))))) else 1.0))))

def f270(E, r, c):
    # Ô M188
    # = IF(COUNT($J$188:$L$188,$P$188)=0,"",AVERAGE($J$188:$L$188,$P$188))
    return ('' if X.truthy(X.eq(X.COUNT(E.rg(8,188,10,188,12), E.rg(8,188,16,188,16)),0.0)) else X.AVERAGE(E.rg(8,188,10,188,12), E.rg(8,188,16,188,16)))

def f271(E, r, c):
    # Ô P188
    # = IF(LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))="BẤT LỢI",-2,IF(OR(LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_N
    #   GHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))="BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG",LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))="NG
    #   HIÊNG BẤT LỢI"),-1,IF(LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))="TRUNG BÌNH",0,IF(OR(LEFT(SU_NGHIEP!$J$6,IFERR
    #   OR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))="NGHIÊNG THUẬN",LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$
    #   6)))="THUẬN NHƯNG CÓ TRỞ NGẠI"),1,IF(LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))="THUẬN",2,"")))))
    return (X.neg(2.0) if X.truthy(X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'BẤT LỢI')) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG'), X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'NGHIÊNG BẤT LỢI'))) else (0.0 if X.truthy(X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'TRUNG BÌNH')) else (1.0 if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'THUẬN NHƯNG CÓ TRỞ NGẠI'))) else (2.0 if X.truthy(X.eq(X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10)))),'THUẬN')) else '')))))

def f272(E, r, c):
    # Ô C189
    # = "Đại vận TỐT (xu thế cuối): "&IF($J$189="","không có",$J$189)&" Đại vận XẤU: "&IF($K$189="","không có",$K$189)&IF($L$189=""," "," Can – chi TRÁI CHIỀU
    #    (một nửa dụng/hỷ, một nửa kỵ – 5 năm đầu theo can, 5 năm sau theo chi): "&$L$189)
    return X.cat(X.cat(X.cat(X.cat('Đại vận TỐT (xu thế cuối): ',('không có' if X.truthy(X.eq(E.v(8,189,10),'')) else E.v(8,189,10))),' Đại vận XẤU: '),('không có' if X.truthy(X.eq(E.v(8,189,11),'')) else E.v(8,189,11))),(' ' if X.truthy(X.eq(E.v(8,189,12),'')) else X.cat(' Can – chi TRÁI CHIỀU (một nửa dụng/hỷ, một nửa kỵ – 5 năm đầu theo can, 5 năm sau theo chi): ',E.v(8,189,12))))

def f273(E, r, c):
    # Ô J189
    # = TRIM(IF(AND(DAI_VAN!C$35<>"",DAI_VAN!C$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!C$69,DAI_VAN!C$54)))),DAI_VAN!C$35&" ("&DAI_VAN!C$36&
    #   "–"&(DAI_VAN!C$36+9)&")"&"; ","")&IF(AND(DAI_VAN!D$35<>"",DAI_VAN!D$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!D$69,DAI_VAN!D$54)))),DA
    #   I_VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&")"&"; ","")&IF(AND(DAI_VAN!E$35<>"",DAI_VAN!E$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI
    #   _VAN!E$69,DAI_VAN!E$54)))),DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&")"&"; ","")&IF(AND(DAI_VAN!F$35<>"",DAI_VAN!F$36<=90,ISNUMBER(SEARCH("
    #   Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!F$69,DAI_VAN!F$54)))),DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&")"&"; ","")&IF(AND(DAI_VAN!G$35<>"",DAI_V
    #   AN!G$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!G$69,DAI_VAN!G$54)))),DAI_VAN!G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&")"&"; ","")&
    #   IF(AND(DAI_VAN!H$35<>"",DAI_VAN!H$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!H$69,DAI_VAN!H$54)))),DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(
    #   DAI_VAN!H$36+9)&")"&"; ","")&IF(AND(DAI_VAN!I$35<>"",DAI_VAN!I$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!I$69,DAI_VAN!I$54)))),DAI_VAN
    #   !I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&")"&"; ","")&IF(AND(DAI_VAN!J$35<>"",DAI_VAN!J$36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!
    #   J$69,DAI_VAN!J$54)))),DAI_VAN!J$35&" ("&DAI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&")"&"; ","")&IF(AND(DAI_VAN!K$35<>"",DAI_VAN!K$36<=90,ISNUMBER(SEARCH("Tốt",
    #   IF(DOI_CHIEU!$C$33=2,DAI_VAN!K$69,DAI_VAN!K$54)))),DAI_VAN!K$35&" ("&DAI_VAN!K$36&"–"&(DAI_VAN!K$36+9)&")"&"; ","")&IF(AND(DAI_VAN!L$35<>"",DAI_VAN!L$
    #   36<=90,ISNUMBER(SEARCH("Tốt",IF(DOI_CHIEU!$C$33=2,DAI_VAN!L$69,DAI_VAN!L$54)))),DAI_VAN!L$35&" ("&DAI_VAN!L$36&"–"&(DAI_VAN!L$36+9)&")"&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-7),' ('),E.v(14,36,c-7)),'–'),X.add(E.v(14,36,c-7),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-7),''), X.le(E.v(14,36,c-7),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-7) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-7)))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-6),' ('),E.v(14,36,c-6)),'–'),X.add(E.v(14,36,c-6),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-6),''), X.le(E.v(14,36,c-6),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-6)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-5),' ('),E.v(14,36,c-5)),'–'),X.add(E.v(14,36,c-5),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-5),''), X.le(E.v(14,36,c-5),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-5) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-5)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-4),' ('),E.v(14,36,c-4)),'–'),X.add(E.v(14,36,c-4),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-4),''), X.le(E.v(14,36,c-4),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-4)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-3),' ('),E.v(14,36,c-3)),'–'),X.add(E.v(14,36,c-3),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-3),''), X.le(E.v(14,36,c-3),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-2),' ('),E.v(14,36,c-2)),'–'),X.add(E.v(14,36,c-2),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-2),''), X.le(E.v(14,36,c-2),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-2)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-1),' ('),E.v(14,36,c-1)),'–'),X.add(E.v(14,36,c-1),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-1),''), X.le(E.v(14,36,c-1),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-1)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c),''), X.le(E.v(14,36,c),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c+1),''), X.le(E.v(14,36,c+1),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c+1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c+1)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c+2),''), X.le(E.v(14,36,c+2),90.0), X.ISNUMBER(lambda: X.SEARCH('Tốt', (E.v(14,69,c+2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c+2)))))) else '')))

def f274(E, r, c):
    # Ô K189
    # = TRIM(IF(AND(DAI_VAN!C$35<>"",DAI_VAN!C$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!C$69,DAI_VAN!C$54)))),DAI_VAN!C$35&" ("&DAI_VAN!C$36&
    #   "–"&(DAI_VAN!C$36+9)&")"&"; ","")&IF(AND(DAI_VAN!D$35<>"",DAI_VAN!D$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!D$69,DAI_VAN!D$54)))),DA
    #   I_VAN!D$35&" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&")"&"; ","")&IF(AND(DAI_VAN!E$35<>"",DAI_VAN!E$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI
    #   _VAN!E$69,DAI_VAN!E$54)))),DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&")"&"; ","")&IF(AND(DAI_VAN!F$35<>"",DAI_VAN!F$36<=90,ISNUMBER(SEARCH("
    #   Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!F$69,DAI_VAN!F$54)))),DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&")"&"; ","")&IF(AND(DAI_VAN!G$35<>"",DAI_V
    #   AN!G$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!G$69,DAI_VAN!G$54)))),DAI_VAN!G$35&" ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&")"&"; ","")&
    #   IF(AND(DAI_VAN!H$35<>"",DAI_VAN!H$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!H$69,DAI_VAN!H$54)))),DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(
    #   DAI_VAN!H$36+9)&")"&"; ","")&IF(AND(DAI_VAN!I$35<>"",DAI_VAN!I$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!I$69,DAI_VAN!I$54)))),DAI_VAN
    #   !I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&")"&"; ","")&IF(AND(DAI_VAN!J$35<>"",DAI_VAN!J$36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!
    #   J$69,DAI_VAN!J$54)))),DAI_VAN!J$35&" ("&DAI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&")"&"; ","")&IF(AND(DAI_VAN!K$35<>"",DAI_VAN!K$36<=90,ISNUMBER(SEARCH("Xấu",
    #   IF(DOI_CHIEU!$C$33=2,DAI_VAN!K$69,DAI_VAN!K$54)))),DAI_VAN!K$35&" ("&DAI_VAN!K$36&"–"&(DAI_VAN!K$36+9)&")"&"; ","")&IF(AND(DAI_VAN!L$35<>"",DAI_VAN!L$
    #   36<=90,ISNUMBER(SEARCH("Xấu",IF(DOI_CHIEU!$C$33=2,DAI_VAN!L$69,DAI_VAN!L$54)))),DAI_VAN!L$35&" ("&DAI_VAN!L$36&"–"&(DAI_VAN!L$36+9)&")"&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-8),' ('),E.v(14,36,c-8)),'–'),X.add(E.v(14,36,c-8),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-8),''), X.le(E.v(14,36,c-8),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-8) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-8)))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-7),' ('),E.v(14,36,c-7)),'–'),X.add(E.v(14,36,c-7),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-7),''), X.le(E.v(14,36,c-7),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-7) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-7)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-6),' ('),E.v(14,36,c-6)),'–'),X.add(E.v(14,36,c-6),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-6),''), X.le(E.v(14,36,c-6),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-6)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-5),' ('),E.v(14,36,c-5)),'–'),X.add(E.v(14,36,c-5),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-5),''), X.le(E.v(14,36,c-5),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-5) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-5)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-4),' ('),E.v(14,36,c-4)),'–'),X.add(E.v(14,36,c-4),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-4),''), X.le(E.v(14,36,c-4),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-4)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-3),' ('),E.v(14,36,c-3)),'–'),X.add(E.v(14,36,c-3),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-3),''), X.le(E.v(14,36,c-3),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-2),' ('),E.v(14,36,c-2)),'–'),X.add(E.v(14,36,c-2),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-2),''), X.le(E.v(14,36,c-2),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-2)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-1),' ('),E.v(14,36,c-1)),'–'),X.add(E.v(14,36,c-1),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-1),''), X.le(E.v(14,36,c-1),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c-1)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c),''), X.le(E.v(14,36,c),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c+1),''), X.le(E.v(14,36,c+1),90.0), X.ISNUMBER(lambda: X.SEARCH('Xấu', (E.v(14,69,c+1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,54,c+1)))))) else '')))

def f275(E, r, c):
    # Ô L189
    # = TRIM(IF(AND(DAI_VAN!C$35<>"",DAI_VAN!C$36<=90,OR(AND(OR(DAI_VAN!C$42=$C$157,DAI_VAN!C$42=$C$158),DAI_VAN!C$43=$C$159),AND(OR(DAI_VAN!C$43=$C$157,DAI_V
    #   AN!C$43=$C$158),DAI_VAN!C$42=$C$159))),DAI_VAN!C$35&" ("&DAI_VAN!C$36&"–"&(DAI_VAN!C$36+9)&")"&"; ","")&IF(AND(DAI_VAN!D$35<>"",DAI_VAN!D$36<=90,OR(AN
    #   D(OR(DAI_VAN!D$42=$C$157,DAI_VAN!D$42=$C$158),DAI_VAN!D$43=$C$159),AND(OR(DAI_VAN!D$43=$C$157,DAI_VAN!D$43=$C$158),DAI_VAN!D$42=$C$159))),DAI_VAN!D$35
    #   &" ("&DAI_VAN!D$36&"–"&(DAI_VAN!D$36+9)&")"&"; ","")&IF(AND(DAI_VAN!E$35<>"",DAI_VAN!E$36<=90,OR(AND(OR(DAI_VAN!E$42=$C$157,DAI_VAN!E$42=$C$158),DAI_V
    #   AN!E$43=$C$159),AND(OR(DAI_VAN!E$43=$C$157,DAI_VAN!E$43=$C$158),DAI_VAN!E$42=$C$159))),DAI_VAN!E$35&" ("&DAI_VAN!E$36&"–"&(DAI_VAN!E$36+9)&")"&"; ",""
    #   )&IF(AND(DAI_VAN!F$35<>"",DAI_VAN!F$36<=90,OR(AND(OR(DAI_VAN!F$42=$C$157,DAI_VAN!F$42=$C$158),DAI_VAN!F$43=$C$159),AND(OR(DAI_VAN!F$43=$C$157,DAI_VAN!
    #   F$43=$C$158),DAI_VAN!F$42=$C$159))),DAI_VAN!F$35&" ("&DAI_VAN!F$36&"–"&(DAI_VAN!F$36+9)&")"&"; ","")&IF(AND(DAI_VAN!G$35<>"",DAI_VAN!G$36<=90,OR(AND(O
    #   R(DAI_VAN!G$42=$C$157,DAI_VAN!G$42=$C$158),DAI_VAN!G$43=$C$159),AND(OR(DAI_VAN!G$43=$C$157,DAI_VAN!G$43=$C$158),DAI_VAN!G$42=$C$159))),DAI_VAN!G$35&" 
    #   ("&DAI_VAN!G$36&"–"&(DAI_VAN!G$36+9)&")"&"; ","")&IF(AND(DAI_VAN!H$35<>"",DAI_VAN!H$36<=90,OR(AND(OR(DAI_VAN!H$42=$C$157,DAI_VAN!H$42=$C$158),DAI_VAN!
    #   H$43=$C$159),AND(OR(DAI_VAN!H$43=$C$157,DAI_VAN!H$43=$C$158),DAI_VAN!H$42=$C$159))),DAI_VAN!H$35&" ("&DAI_VAN!H$36&"–"&(DAI_VAN!H$36+9)&")"&"; ","")&I
    #   F(AND(DAI_VAN!I$35<>"",DAI_VAN!I$36<=90,OR(AND(OR(DAI_VAN!I$42=$C$157,DAI_VAN!I$42=$C$158),DAI_VAN!I$43=$C$159),AND(OR(DAI_VAN!I$43=$C$157,DAI_VAN!I$4
    #   3=$C$158),DAI_VAN!I$42=$C$159))),DAI_VAN!I$35&" ("&DAI_VAN!I$36&"–"&(DAI_VAN!I$36+9)&")"&"; ","")&IF(AND(DAI_VAN!J$35<>"",DAI_VAN!J$36<=90,OR(AND(OR(D
    #   AI_VAN!J$42=$C$157,DAI_VAN!J$42=$C$158),DAI_VAN!J$43=$C$159),AND(OR(DAI_VAN!J$43=$C$157,DAI_VAN!J$43=$C$158),DAI_VAN!J$42=$C$159))),DAI_VAN!J$35&" ("&
    #   DAI_VAN!J$36&"–"&(DAI_VAN!J$36+9)&")"&"; ","")&IF(AND(DAI_VAN!K$35<>"",DAI_VAN!K$36<=90,OR(AND(OR(DAI_VAN!K$42=$C$157,DAI_VAN!K$42=$C$158),DAI_VAN!K$4
    #   3=$C$159),AND(OR(DAI_VAN!K$43=$C$157,DAI_VAN!K$43=$C$158),DAI_VAN!K$42=$C$159))),DAI_VAN!K$35&" ("&DAI_VAN!K$36&"–"&(DAI_VAN!K$36+9)&")"&"; ","")&IF(A
    #   ND(DAI_VAN!L$35<>"",DAI_VAN!L$36<=90,OR(AND(OR(DAI_VAN!L$42=$C$157,DAI_VAN!L$42=$C$158),DAI_VAN!L$43=$C$159),AND(OR(DAI_VAN!L$43=$C$157,DAI_VAN!L$43=$
    #   C$158),DAI_VAN!L$42=$C$159))),DAI_VAN!L$35&" ("&DAI_VAN!L$36&"–"&(DAI_VAN!L$36+9)&")"&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-9),' ('),E.v(14,36,c-9)),'–'),X.add(E.v(14,36,c-9),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-9),''), X.le(E.v(14,36,c-9),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-9),E.v(8,157,3)), X.eq(E.v(14,42,c-9),E.v(8,158,3))), X.eq(E.v(14,43,c-9),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-9),E.v(8,157,3)), X.eq(E.v(14,43,c-9),E.v(8,158,3))), X.eq(E.v(14,42,c-9),E.v(8,159,3)))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-8),' ('),E.v(14,36,c-8)),'–'),X.add(E.v(14,36,c-8),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-8),''), X.le(E.v(14,36,c-8),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-8),E.v(8,157,3)), X.eq(E.v(14,42,c-8),E.v(8,158,3))), X.eq(E.v(14,43,c-8),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-8),E.v(8,157,3)), X.eq(E.v(14,43,c-8),E.v(8,158,3))), X.eq(E.v(14,42,c-8),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-7),' ('),E.v(14,36,c-7)),'–'),X.add(E.v(14,36,c-7),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-7),''), X.le(E.v(14,36,c-7),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-7),E.v(8,157,3)), X.eq(E.v(14,42,c-7),E.v(8,158,3))), X.eq(E.v(14,43,c-7),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-7),E.v(8,157,3)), X.eq(E.v(14,43,c-7),E.v(8,158,3))), X.eq(E.v(14,42,c-7),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-6),' ('),E.v(14,36,c-6)),'–'),X.add(E.v(14,36,c-6),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-6),''), X.le(E.v(14,36,c-6),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-6),E.v(8,157,3)), X.eq(E.v(14,42,c-6),E.v(8,158,3))), X.eq(E.v(14,43,c-6),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-6),E.v(8,157,3)), X.eq(E.v(14,43,c-6),E.v(8,158,3))), X.eq(E.v(14,42,c-6),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-5),' ('),E.v(14,36,c-5)),'–'),X.add(E.v(14,36,c-5),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-5),''), X.le(E.v(14,36,c-5),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-5),E.v(8,157,3)), X.eq(E.v(14,42,c-5),E.v(8,158,3))), X.eq(E.v(14,43,c-5),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-5),E.v(8,157,3)), X.eq(E.v(14,43,c-5),E.v(8,158,3))), X.eq(E.v(14,42,c-5),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-4),' ('),E.v(14,36,c-4)),'–'),X.add(E.v(14,36,c-4),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-4),''), X.le(E.v(14,36,c-4),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-4),E.v(8,157,3)), X.eq(E.v(14,42,c-4),E.v(8,158,3))), X.eq(E.v(14,43,c-4),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-4),E.v(8,157,3)), X.eq(E.v(14,43,c-4),E.v(8,158,3))), X.eq(E.v(14,42,c-4),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-3),' ('),E.v(14,36,c-3)),'–'),X.add(E.v(14,36,c-3),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-3),''), X.le(E.v(14,36,c-3),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-3),E.v(8,157,3)), X.eq(E.v(14,42,c-3),E.v(8,158,3))), X.eq(E.v(14,43,c-3),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-3),E.v(8,157,3)), X.eq(E.v(14,43,c-3),E.v(8,158,3))), X.eq(E.v(14,42,c-3),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-2),' ('),E.v(14,36,c-2)),'–'),X.add(E.v(14,36,c-2),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-2),''), X.le(E.v(14,36,c-2),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-2),E.v(8,157,3)), X.eq(E.v(14,42,c-2),E.v(8,158,3))), X.eq(E.v(14,43,c-2),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-2),E.v(8,157,3)), X.eq(E.v(14,43,c-2),E.v(8,158,3))), X.eq(E.v(14,42,c-2),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c-1),' ('),E.v(14,36,c-1)),'–'),X.add(E.v(14,36,c-1),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c-1),''), X.le(E.v(14,36,c-1),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c-1),E.v(8,157,3)), X.eq(E.v(14,42,c-1),E.v(8,158,3))), X.eq(E.v(14,43,c-1),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c-1),E.v(8,157,3)), X.eq(E.v(14,43,c-1),E.v(8,158,3))), X.eq(E.v(14,42,c-1),E.v(8,159,3)))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),')'),'; ') if X.truthy(X.AND(X.ne(E.v(14,35,c),''), X.le(E.v(14,36,c),90.0), X.OR(X.AND(X.OR(X.eq(E.v(14,42,c),E.v(8,157,3)), X.eq(E.v(14,42,c),E.v(8,158,3))), X.eq(E.v(14,43,c),E.v(8,159,3))), X.AND(X.OR(X.eq(E.v(14,43,c),E.v(8,157,3)), X.eq(E.v(14,43,c),E.v(8,158,3))), X.eq(E.v(14,42,c),E.v(8,159,3)))))) else '')))

def f276(E, r, c):
    # Ô D190
    # = " ► Lời khuyên: "&IF($C$155=0,"thân nhược – học tập, làm việc trong tập thể, có thầy, người cộng sự nâng đỡ; nhận việc vừa sức, tránh ôm đồm, tránh đầ
    #   u tư hay bảo lãnh vượt khả năng; giữ nếp ăn ngủ, rèn thể lực đều đặn.",IF($C$155=1,"thân vượng – chủ động gánh việc, tạo ra sản phẩm, giúp người; tiết
    #    chế cái tôi, tránh cố chấp, tranh giành, liều lĩnh.","tứ trụ gần cân bằng – giữ nhịp sống điều độ, tránh để một mặt nào lấn át."))&" [Lời khuyên của 
    #   công cụ – không phải câu sách]"
    return X.cat(X.cat(' ► Lời khuyên: ',('thân nhược – học tập, làm việc trong tập thể, có thầy, người cộng sự nâng đỡ; nhận việc vừa sức, tránh ôm đồm, tránh đầu tư hay bảo lãnh vượt khả năng; giữ nếp ăn ngủ, rèn thể lực đều đặn.' if X.truthy(X.eq(E.v(8,155,3),0.0)) else ('thân vượng – chủ động gánh việc, tạo ra sản phẩm, giúp người; tiết chế cái tôi, tránh cố chấp, tranh giành, liều lĩnh.' if X.truthy(X.eq(E.v(8,155,3),1.0)) else 'tứ trụ gần cân bằng – giữ nhịp sống điều độ, tránh để một mặt nào lấn át.'))),' [Lời khuyên của công cụ – không phải câu sách]')

def f277(E, r, c):
    # Ô D192
    # = IF($M$188="",""," ► Lời khuyên: "&IF($M$188<0.5,"tiền đến nhưng khó giữ – lập quỹ tiết kiệm cố định, không vay mượn, đầu cơ, bảo lãnh; đầu tư vào học 
    #   vấn, tay nghề; đẩy mạnh làm ăn trong vận dụng – hỷ thần, thu hẹp trong vận kỵ thần.","có nền tảng tài lộc – tích lũy dài hạn, đa dạng hóa, giữ kỷ luật
    #    chi tiêu; khi vào vận kỵ thần thì phòng thủ, không mở rộng lớn.")&" [Lời khuyên của công cụ – không phải câu sách]")
    return ('' if X.truthy(X.eq(E.v(8,188,13),'')) else X.cat(X.cat(' ► Lời khuyên: ',('tiền đến nhưng khó giữ – lập quỹ tiết kiệm cố định, không vay mượn, đầu cơ, bảo lãnh; đầu tư vào học vấn, tay nghề; đẩy mạnh làm ăn trong vận dụng – hỷ thần, thu hẹp trong vận kỵ thần.' if X.truthy(X.lt(E.v(8,188,13),0.5)) else 'có nền tảng tài lộc – tích lũy dài hạn, đa dạng hóa, giữ kỷ luật chi tiêu; khi vào vận kỵ thần thì phòng thủ, không mở rộng lớn.')),' [Lời khuyên của công cụ – không phải câu sách]'))

def f278(E, r, c):
    # Ô D193
    # = " ► Lời khuyên: trong vận TỐT – chủ động học lên, thi cử, chuyển việc, mở rộng có kiểm soát; "&"trong vận XẤU – giữ ổn định, không mạo hiểm lớn, tích 
    #   lũy và chuẩn bị cho vận sau; vận trái chiều – tận dụng nửa thuận, phòng thủ nửa nghịch. [Lời khuyên của công cụ – không phải câu sách]"
    return X.cat(' ► Lời khuyên: trong vận TỐT – chủ động học lên, thi cử, chuyển việc, mở rộng có kiểm soát; ','trong vận XẤU – giữ ổn định, không mạo hiểm lớn, tích lũy và chuẩn bị cho vận sau; vận trái chiều – tận dụng nửa thuận, phòng thủ nửa nghịch. [Lời khuyên của công cụ – không phải câu sách]')

def f279(E, r, c):
    # Ô D194
    # = " ► Lời khuyên: năm có dấu hiệu hạn nặng – khám sức khỏe định kỳ, đi lại cẩn thận, tránh nơi nguy hiểm, quan tâm cha mẹ và người thân lớn tuổi, hoãn q
    #   uyết định lớn (đầu tư, ký kết, bảo lãnh); năm tốt – dồn sức cho việc quan trọng."&" [Lời khuyên của công cụ – không phải câu sách]"
    return X.cat(' ► Lời khuyên: năm có dấu hiệu hạn nặng – khám sức khỏe định kỳ, đi lại cẩn thận, tránh nơi nguy hiểm, quan tâm cha mẹ và người thân lớn tuổi, hoãn quyết định lớn (đầu tư, ký kết, bảo lãnh); năm tốt – dồn sức cho việc quan trọng.',' [Lời khuyên của công cụ – không phải câu sách]')

def f280(E, r, c):
    # Ô D195
    # = " ► Lời khuyên: các nhận định sức khỏe, thọ – yểu là xu hướng theo bảng của sách, không phải tiên đoán; giữ nếp sống lành mạnh, khám định kỳ, chú ý cá
    #   c tạng phủ được nêu. Không thay thế ý kiến bác sĩ. [Lời khuyên của công cụ – không phải câu sách]"
    return ' ► Lời khuyên: các nhận định sức khỏe, thọ – yểu là xu hướng theo bảng của sách, không phải tiên đoán; giữ nếp sống lành mạnh, khám định kỳ, chú ý các tạng phủ được nêu. Không thay thế ý kiến bác sĩ. [Lời khuyên của công cụ – không phải câu sách]'

def f281(E, r, c):
    # Ô C198
    # = IF($D$198<0,"","Dụng thần theo cách bổ hành khuyết: "&INDEX(HANH_TEN,MOD($C$5+$D$198,5)+1)&" ("&INDEX(NHOM_TEN,$D$198+1)&") – tứ trụ không có chữ nào,
    #    lộ hay tàng, thuộc nhóm này. Khi giải ví dụ, sách nhiều lần lấy nhóm đang khuyết làm dụng thần (dụng thần khuyết thì chờ đại vận, lưu niên đưa tới – 
    #   P1 tr.97). Chỉ là ghi chú để cân nhắc, không thay dụng thần "&INDEX(HANH_TEN,$C$157+1)&" ("&INDEX(NHOM_TEN,$C$156+1)&") đã chọn theo bảng P1 tr.108, P
    #   2 tr.380-384.")
    return ('' if X.truthy(X.lt(E.v(8,198,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dụng thần theo cách bổ hành khuyết: ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,3),E.v(8,198,4)), 5.0),1.0))),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,198,4),1.0))),') – tứ trụ không có chữ nào, lộ hay tàng, thuộc nhóm này. Khi giải ví dụ, sách nhiều lần lấy nhóm đang khuyết làm dụng thần (dụng thần khuyết thì chờ đại vận, lưu niên đưa tới – P1 tr.97). Chỉ là ghi chú để cân nhắc, không thay dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,157,3),1.0))),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,156,3),1.0))),') đã chọn theo bảng P1 tr.108, P2 tr.380-384.'))

def f282(E, r, c):
    # Ô D198
    # = IF(OR($C$160<>"",$C$155=2,$C$156<0),-1,IF($C$155=1,IF(OR($C$156=1,$C$156=2,$C$156=3),IF(AND($C$156<>3,$C$21=0),3,IF(AND($C$156<>1,$C$19=0),1,IF(AND($C
    #   $156<>2,$C$20=0),2,-1))),-1),IF(OR($C$156=0,$C$156=4),IF(AND($C$156<>4,$C$22=0),4,IF(AND($C$156<>0,$C$18=0),0,-1)),-1)))
    return (X.neg(1.0) if X.truthy(X.OR(X.ne(E.v(8,160,3),''), X.eq(E.v(8,155,3),2.0), X.lt(E.v(8,156,3),0.0))) else (((3.0 if X.truthy(X.AND(X.ne(E.v(8,156,3),3.0), X.eq(E.v(8,21,3),0.0))) else (1.0 if X.truthy(X.AND(X.ne(E.v(8,156,3),1.0), X.eq(E.v(8,19,3),0.0))) else (2.0 if X.truthy(X.AND(X.ne(E.v(8,156,3),2.0), X.eq(E.v(8,20,3),0.0))) else X.neg(1.0)))) if X.truthy(X.OR(X.eq(E.v(8,156,3),1.0), X.eq(E.v(8,156,3),2.0), X.eq(E.v(8,156,3),3.0))) else X.neg(1.0)) if X.truthy(X.eq(E.v(8,155,3),1.0)) else ((4.0 if X.truthy(X.AND(X.ne(E.v(8,156,3),4.0), X.eq(E.v(8,22,3),0.0))) else (0.0 if X.truthy(X.AND(X.ne(E.v(8,156,3),0.0), X.eq(E.v(8,18,3),0.0))) else X.neg(1.0))) if X.truthy(X.OR(X.eq(E.v(8,156,3),0.0), X.eq(E.v(8,156,3),4.0))) else X.neg(1.0))))

def f283(E, r, c):
    # Ô C202, D202
    # = IF(C61<>"",C61,IF(C65<>"",C65,IF(C66<>"",C66,IF(C68<>"",C68,IF(C69<>"",C69,IF(C70<>"",C70,C71))))))
    return (E.v(8,r-141,c) if X.truthy(X.ne(E.v(8,r-141,c),'')) else (E.v(8,r-137,c) if X.truthy(X.ne(E.v(8,r-137,c),'')) else (E.v(8,r-136,c) if X.truthy(X.ne(E.v(8,r-136,c),'')) else (E.v(8,r-134,c) if X.truthy(X.ne(E.v(8,r-134,c),'')) else (E.v(8,r-133,c) if X.truthy(X.ne(E.v(8,r-133,c),'')) else (E.v(8,r-132,c) if X.truthy(X.ne(E.v(8,r-132,c),'')) else E.v(8,r-131,c)))))))

def f284(E, r, c):
    # Ô C203
    # = IF(OR(INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$6+1)=3,INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$C$6+1)=4,AND(INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!
    #   $C$6+1)=0,TRUNG_GIAN!$E$8=0),INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$6+1)=3,INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$D$6+1)=4,AND(INDEX(M_TS,TRUNG_
    #   GIAN!$E$5+1,TRUNG_GIAN!$D$6+1)=0,TRUNG_GIAN!$E$8=0),INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$6+1)=3,INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$6+1)=
    #   4,AND(INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$E$6+1)=0,TRUNG_GIAN!$E$8=0),INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$6+1)=3,INDEX(M_TS,TRUNG_GIAN!$E$
    #   5+1,TRUNG_GIAN!$F$6+1)=4,AND(INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$6+1)=0,TRUNG_GIAN!$E$8=0)),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,3),1.0)),3.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,3),1.0)),4.0), X.AND(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,3),1.0)),0.0), X.eq(E.v(23,8,5),0.0)), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,4),1.0)),3.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,4),1.0)),4.0), X.AND(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,4),1.0)),0.0), X.eq(E.v(23,8,5),0.0)), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,5),1.0)),3.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,5),1.0)),4.0), X.AND(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,5),1.0)),0.0), X.eq(E.v(23,8,5),0.0)), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,6),1.0)),3.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,6),1.0)),4.0), X.AND(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(23,6,6),1.0)),0.0), X.eq(E.v(23,8,5),0.0)))) else 0.0)

def f285(E, r, c):
    # Ô D203
    # = $C$203
    return E.v(8,203,3)

def f286(E, r, c):
    # Ô C204
    # = IF(OR(TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7,TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)+IF(OR(TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7,TRUNG_GIAN!$D$7=MOD(TRUNG_G
    #   IAN!$E$7+4,5)),1,0)+IF(OR(TRUNG_GIAN!$E$7=TRUNG_GIAN!$E$7,TRUNG_GIAN!$E$7=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)+IF(OR(TRUNG_GIAN!$F$7=TRUNG_GIAN!$E$7,TRUNG_G
    #   IAN!$F$7=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)+IF(OR(IF(PP1_DINH_TINH!$C$27>=0,PP1_DINH_TINH!$C$27,TRUNG_GIAN!$C$9)=TRUNG_GIAN!$E$7,IF(PP1_DINH_TINH!$C$27>=0
    #   ,PP1_DINH_TINH!$C$27,TRUNG_GIAN!$C$9)=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)+IF(OR(IF(PP1_DINH_TINH!$D$27>=0,PP1_DINH_TINH!$D$27,TRUNG_GIAN!$D$9)=TRUNG_GIAN!$
    #   E$7,IF(PP1_DINH_TINH!$D$27>=0,PP1_DINH_TINH!$D$27,TRUNG_GIAN!$D$9)=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)+IF(OR(IF(PP1_DINH_TINH!$E$27>=0,PP1_DINH_TINH!$E$27,
    #   TRUNG_GIAN!$E$9)=TRUNG_GIAN!$E$7,IF(PP1_DINH_TINH!$E$27>=0,PP1_DINH_TINH!$E$27,TRUNG_GIAN!$E$9)=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)+IF(OR(IF(PP1_DINH_TINH!
    #   $F$27>=0,PP1_DINH_TINH!$F$27,TRUNG_GIAN!$F$9)=TRUNG_GIAN!$E$7,IF(PP1_DINH_TINH!$F$27>=0,PP1_DINH_TINH!$F$27,TRUNG_GIAN!$F$9)=MOD(TRUNG_GIAN!$E$7+4,5))
    #   ,1,0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.OR(X.eq(E.v(23,7,3),E.v(23,7,5)), X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0),(1.0 if X.truthy(X.OR(X.eq(E.v(23,7,4),E.v(23,7,5)), X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(E.v(23,7,5),E.v(23,7,5)), X.eq(E.v(23,7,5),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(E.v(23,7,6),E.v(23,7,5)), X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq((E.v(9,27,3) if X.truthy(X.ge(E.v(9,27,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)), X.eq((E.v(9,27,3) if X.truthy(X.ge(E.v(9,27,3),0.0)) else E.v(23,9,3)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq((E.v(9,27,4) if X.truthy(X.ge(E.v(9,27,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)), X.eq((E.v(9,27,4) if X.truthy(X.ge(E.v(9,27,4),0.0)) else E.v(23,9,4)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq((E.v(9,27,5) if X.truthy(X.ge(E.v(9,27,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)), X.eq((E.v(9,27,5) if X.truthy(X.ge(E.v(9,27,5),0.0)) else E.v(23,9,5)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq((E.v(9,27,6) if X.truthy(X.ge(E.v(9,27,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)), X.eq((E.v(9,27,6) if X.truthy(X.ge(E.v(9,27,6),0.0)) else E.v(23,9,6)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else 0.0))

def f287(E, r, c):
    # Ô D204
    # = $C$204
    return E.v(8,204,3)

def f288(E, r, c):
    # Ô C205, D205
    # = IF(AND(C202="Tòng nhi",C203=1),"nhật can có gốc mạnh ở chi nên không tòng – giữ chính cách",IF(C202="Tòng vượng (độc vượng)","độc vượng theo số độ khô
    #   ng tự thành cách – chỉ thành cách khi đủ điều kiện cách chuyên vượng; giữ chính cách",IF(AND(C61<>"",C202=C61,C204<5),"hành nhật chủ và hành sinh nhật
    #    chủ mới có "&C204&"/8 chữ, chưa đủ điều kiện tứ trụ hành này nhiều – giữ chính cách","")))
    return ('nhật can có gốc mạnh ở chi nên không tòng – giữ chính cách' if X.truthy(X.AND(X.eq(E.v(8,r-3,c),'Tòng nhi'), X.eq(E.v(8,r-2,c),1.0))) else ('độc vượng theo số độ không tự thành cách – chỉ thành cách khi đủ điều kiện cách chuyên vượng; giữ chính cách' if X.truthy(X.eq(E.v(8,r-3,c),'Tòng vượng (độc vượng)')) else (X.cat(X.cat('hành nhật chủ và hành sinh nhật chủ mới có ',E.v(8,r-1,c)),'/8 chữ, chưa đủ điều kiện tứ trụ hành này nhiều – giữ chính cách') if X.truthy(X.AND(X.ne(E.v(8,r-144,c),''), X.eq(E.v(8,r-3,c),E.v(8,r-144,c)), X.lt(E.v(8,r-1,c),5.0))) else '')))

def f289(E, r, c):
    # Ô C206, D206
    # = IF(C205="","",IF(C202="Tòng nhi","P2 tr.373–375: cách tòng đòi nhật chủ không có khí gốc; các ví dụ tòng nhi của sách nhật can không có gốc mạnh",IF(C
    #   202="Tòng vượng (độc vượng)","P2 tr.353 chỉ gọi tên kết quả tính độ; điều kiện thành cách chuyên vượng ở P2 tr.369–373","P2 tr.369–373: cách chuyên vư
    #   ợng đòi tứ trụ hành nhật chủ (và hành sinh nó) nhiều – hiểu là quá nửa, ≥ 5/8 chữ")))
    return ('' if X.truthy(X.eq(E.v(8,r-1,c),'')) else ('P2 tr.373–375: cách tòng đòi nhật chủ không có khí gốc; các ví dụ tòng nhi của sách nhật can không có gốc mạnh' if X.truthy(X.eq(E.v(8,r-4,c),'Tòng nhi')) else ('P2 tr.353 chỉ gọi tên kết quả tính độ; điều kiện thành cách chuyên vượng ở P2 tr.369–373' if X.truthy(X.eq(E.v(8,r-4,c),'Tòng vượng (độc vượng)')) else 'P2 tr.369–373: cách chuyên vượng đòi tứ trụ hành nhật chủ (và hành sinh nó) nhiều – hiểu là quá nửa, ≥ 5/8 chữ')))

def f290(E, r, c):
    # Ô C209, D209
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+0,5))*(TRUNG_GIAN!$C$4:$F$4<>2))+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=MOD(C$5+0,5))*1)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),0.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),X.AF_MOD(X.A('add',E.v(8,5,c),0.0), 5.0)),1.0))),0.0)) else 0.0)

def f291(E, r, c):
    # Ô C210, D210
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+1,5))*(TRUNG_GIAN!$C$4:$F$4<>2))+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=MOD(C$5+1,5))*1)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),X.AF_MOD(X.A('add',E.v(8,5,c),1.0), 5.0)),1.0))),0.0)) else 0.0)

def f292(E, r, c):
    # Ô C211, D211
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+2,5))*(TRUNG_GIAN!$C$4:$F$4<>2))+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=MOD(C$5+2,5))*1)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),2.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),X.AF_MOD(X.A('add',E.v(8,5,c),2.0), 5.0)),1.0))),0.0)) else 0.0)

def f293(E, r, c):
    # Ô C212, D212
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+3,5))*(TRUNG_GIAN!$C$4:$F$4<>2))+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=MOD(C$5+3,5))*1)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),X.AF_MOD(X.A('add',E.v(8,5,c),3.0), 5.0)),1.0))),0.0)) else 0.0)

def f294(E, r, c):
    # Ô C213, D213
    # = IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=MOD(C$5+4,5))*(TRUNG_GIAN!$C$4:$F$4<>2))+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=MOD(C$5+4,5))*1)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.add(X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,7,3,7,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),X.A('ne',E.rg(23,4,3,4,6),2.0))),X.SUMPRODUCT(X.A('mul',X.A('eq',E.rg(23,21,3,23,6),X.AF_MOD(X.A('add',E.v(8,5,c),4.0), 5.0)),1.0))),0.0)) else 0.0)

def f295(E, r, c):
    # Ô C214, D214
    # = IF(AND(C93>=0,INDEX(C$209:C$213,C93+1)=1),C93,IF(AND(C94>=0,INDEX(C$209:C$213,MAX(0,C94)+1)=1),C94,IF(AND(C95>=0,INDEX(C$209:C$213,MAX(0,C95)+1)=1),C9
    #   5,IF(C91="TH",-1,C93))))
    return (E.v(8,r-121,c) if X.truthy(X.AND(X.ge(E.v(8,r-121,c),0.0), X.eq(X.INDEX(E.rg(8,209,c,213,c), X.add(E.v(8,r-121,c),1.0)),1.0))) else (E.v(8,r-120,c) if X.truthy(X.AND(X.ge(E.v(8,r-120,c),0.0), X.eq(X.INDEX(E.rg(8,209,c,213,c), X.add(X.MAX(0.0, E.rg(8,r-120,c,r-120,c)),1.0)),1.0))) else (E.v(8,r-119,c) if X.truthy(X.AND(X.ge(E.v(8,r-119,c),0.0), X.eq(X.INDEX(E.rg(8,209,c,213,c), X.add(X.MAX(0.0, E.rg(8,r-119,c,r-119,c)),1.0)),1.0))) else (X.neg(1.0) if X.truthy(X.eq(E.v(8,r-123,c),'TH')) else E.v(8,r-121,c)))))

def f296(E, r, c):
    # Ô C215
    # = IF(AND(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$D$6=2),OR(TRUNG_GIAN!$E$7=0,TRUNG_GIAN!$E$7=3,TRUNG_GIAN!$E$7=4)),1,IF(AN
    #   D(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),OR(TRUNG_GIAN!$E$7=1,TRUNG_GIAN!$E$7=2,TRUNG_GIAN!$E$7=3)),4,-1))
    return (1.0 if X.truthy(X.AND(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,4),2.0)), X.OR(X.eq(E.v(23,7,5),0.0), X.eq(E.v(23,7,5),3.0), X.eq(E.v(23,7,5),4.0)))) else (4.0 if X.truthy(X.AND(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0)), X.OR(X.eq(E.v(23,7,5),1.0), X.eq(E.v(23,7,5),2.0), X.eq(E.v(23,7,5),3.0)))) else X.neg(1.0)))

def f297(E, r, c):
    # Ô D215
    # = $C$215
    return E.v(8,215,3)

def f298(E, r, c):
    # Ô C216
    # = IF($C$215<0,"",IF(SUM($D$24:$D$28)=0,0,INDEX($D$24:$D$28,$C$215+1)/SUM($D$24:$D$28)))
    return ('' if X.truthy(X.lt(E.v(8,215,3),0.0)) else (0.0 if X.truthy(X.eq(X.SUM(E.rg(8,24,4,28,4)),0.0)) else X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,215,3),1.0)),X.SUM(E.rg(8,24,4,28,4)))))

def f299(E, r, c):
    # Ô D216
    # = $C$216
    return E.v(8,216,3)

def f300(E, r, c):
    # Ô C217, D217
    # = IF(AND(C$91<>"TH",$C$215>=0,N($C$216)<0.2,MOD($C$215-C$5,5)<>C214),MOD($C$215-C$5,5),C214)
    return (X.MOD(X.sub(E.v(8,215,3),E.v(8,5,c)), 5.0) if X.truthy(X.AND(X.ne(E.v(8,91,c),'TH'), X.ge(E.v(8,215,3),0.0), X.lt(X.N(E.v(8,216,3)),0.2), X.ne(X.MOD(X.sub(E.v(8,215,3),E.v(8,5,c)), 5.0),E.v(8,r-3,c)))) else E.v(8,r-3,c))

def f301(E, r, c):
    # Ô C218, D218
    # = IF(C217=C214,"","Nhật chủ "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" ("&IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!
    #   $D$6=6,TRUNG_GIAN!$D$6=7),"mùa nóng","mùa lạnh")&") cần "&INDEX(HANH_TEN,$C$215+1)&" để điều hầu; "&INDEX(HANH_TEN,$C$215+1)&" chỉ chiếm "&TEXT($C$216
    #   ,"0%")&" tổng độ nên lấy "&INDEX(HANH_TEN,$C$215+1)&" ("&INDEX(NHOM_TEN,C217+1)&") làm dụng thần thay cho "&INDEX(HANH_TEN,MOD(C$5+C214,5)+1)&" ("&IND
    #   EX(NHOM_TEN,C214+1)&").")
    return ('' if X.truthy(X.eq(E.v(8,r-1,c),E.v(8,r-4,c))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' sinh tháng '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' ('),('mùa nóng' if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else 'mùa lạnh')),') cần '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,215,3),1.0))),' để điều hầu; '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,215,3),1.0))),' chỉ chiếm '),X.TEXT(E.v(8,216,3), '0%')),' tổng độ nên lấy '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,215,3),1.0))),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-1,c),1.0))),') làm dụng thần thay cho '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,c),E.v(8,r-4,c)), 5.0),1.0))),' ('),X.INDEX(E.rg(24,54,1,58,1), X.add(E.v(8,r-4,c),1.0))),').'))

def f302(E, r, c):
    # Ô C219, D219
    # = IF(C218="","","P1 tr.76–79: sinh mùa lạnh cần hỏa sưởi ấm, sinh mùa nóng cần thủy điều hòa; P1 tr.100: điều hầu là một nguyên tắc chọn dụng thần")
    return ('' if X.truthy(X.eq(E.v(8,r-1,c),'')) else 'P1 tr.76–79: sinh mùa lạnh cần hỏa sưởi ấm, sinh mùa nóng cần thủy điều hòa; P1 tr.100: điều hầu là một nguyên tắc chọn dụng thần')

CELLS = {
    'C5': (f0, 5, 3),
    'D5': (f1, 5, 4),
    'C6': (f2, 6, 3),
    'D6': (f3, 6, 4),
    'C7': (f4, 7, 3),
    'D7': (f5, 7, 4),
    'C8': (f6, 8, 3),
    'D8': (f7, 8, 4),
    'C9': (f8, 9, 3),
    'D9': (f9, 9, 4),
    'C10': (f10, 10, 3),
    'D10': (f9, 10, 4),
    'C12': (f11, 12, 3),
    'D12': (f12, 12, 4),
    'C13': (f13, 13, 3),
    'D13': (f14, 13, 4),
    'C14': (f15, 14, 3),
    'D14': (f16, 14, 4),
    'C15': (f17, 15, 3),
    'D15': (f18, 15, 4),
    'C16': (f19, 16, 3),
    'D16': (f20, 16, 4),
    'C18': (f21, 18, 3),
    'D18': (f22, 18, 4),
    'C19': (f21, 19, 3),
    'D19': (f22, 19, 4),
    'C20': (f21, 20, 3),
    'D20': (f22, 20, 4),
    'C21': (f21, 21, 3),
    'D21': (f22, 21, 4),
    'C22': (f21, 22, 3),
    'D22': (f22, 22, 4),
    'C24': (f23, 24, 3),
    'D24': (f23, 24, 4),
    'C25': (f24, 25, 3),
    'D25': (f24, 25, 4),
    'C26': (f25, 26, 3),
    'D26': (f25, 26, 4),
    'C27': (f26, 27, 3),
    'D27': (f26, 27, 4),
    'C28': (f27, 28, 3),
    'D28': (f27, 28, 4),
    'C29': (f28, 29, 3),
    'D29': (f28, 29, 4),
    'C30': (f29, 30, 3),
    'D30': (f29, 30, 4),
    'C31': (f30, 31, 3),
    'D31': (f30, 31, 4),
    'C32': (f31, 32, 3),
    'D32': (f31, 32, 4),
    'C33': (f32, 33, 3),
    'D33': (f32, 33, 4),
    'C34': (f33, 34, 3),
    'D34': (f33, 34, 4),
    'C35': (f34, 35, 3),
    'D35': (f34, 35, 4),
    'C39': (f35, 39, 3),
    'D39': (f36, 39, 4),
    'E39': (f37, 39, 5),
    'F39': (f38, 39, 6),
    'C40': (f39, 40, 3),
    'D40': (f40, 40, 4),
    'E40': (f41, 40, 5),
    'F40': (f42, 40, 6),
    'C41': (f43, 41, 3),
    'D41': (f44, 41, 4),
    'E41': (f45, 41, 5),
    'F41': (f46, 41, 6),
    'C42': (f47, 42, 3),
    'D42': (f48, 42, 4),
    'E42': (f49, 42, 5),
    'F42': (f50, 42, 6),
    'C43': (f51, 43, 3),
    'D43': (f52, 43, 4),
    'E43': (f53, 43, 5),
    'F43': (f54, 43, 6),
    'C44': (f55, 44, 3),
    'D44': (f56, 44, 4),
    'E44': (f57, 44, 5),
    'F44': (f58, 44, 6),
    'C45': (f59, 45, 3),
    'D45': (f60, 45, 4),
    'E45': (f61, 45, 5),
    'F45': (f62, 45, 6),
    'C47': (f63, 47, 3),
    'D47': (f64, 47, 4),
    'E47': (f65, 47, 5),
    'F47': (f66, 47, 6),
    'C48': (f67, 48, 3),
    'D48': (f68, 48, 4),
    'E48': (f69, 48, 5),
    'F48': (f70, 48, 6),
    'C49': (f71, 49, 3),
    'D49': (f72, 49, 4),
    'E49': (f73, 49, 5),
    'F49': (f74, 49, 6),
    'C50': (f55, 50, 3),
    'D50': (f56, 50, 4),
    'E50': (f57, 50, 5),
    'F50': (f58, 50, 6),
    'C51': (f59, 51, 3),
    'D51': (f60, 51, 4),
    'E51': (f61, 51, 5),
    'F51': (f62, 51, 6),
    'C53': (f75, 53, 3),
    'C57': (f76, 57, 3),
    'D57': (f77, 57, 4),
    'C58': (f78, 58, 3),
    'D58': (f78, 58, 4),
    'C59': (f79, 59, 3),
    'D59': (f80, 59, 4),
    'C60': (f81, 60, 3),
    'D60': (f82, 60, 4),
    'C61': (f83, 61, 3),
    'D61': (f83, 61, 4),
    'C62': (f84, 62, 3),
    'D62': (f85, 62, 4),
    'C63': (f86, 63, 3),
    'D63': (f87, 63, 4),
    'C64': (f88, 64, 3),
    'D64': (f88, 64, 4),
    'C65': (f89, 65, 3),
    'D65': (f89, 65, 4),
    'C66': (f90, 66, 3),
    'D66': (f90, 66, 4),
    'C67': (f91, 67, 3),
    'D67': (f91, 67, 4),
    'C68': (f92, 68, 3),
    'D68': (f93, 68, 4),
    'C69': (f94, 69, 3),
    'D69': (f95, 69, 4),
    'C70': (f96, 70, 3),
    'D70': (f97, 70, 4),
    'C71': (f98, 71, 3),
    'D71': (f99, 71, 4),
    'C72': (f100, 72, 3),
    'D72': (f101, 72, 4),
    'C73': (f102, 73, 3),
    'D73': (f102, 73, 4),
    'C74': (f103, 74, 3),
    'D74': (f103, 74, 4),
    'C75': (f104, 75, 3),
    'D75': (f104, 75, 4),
    'C76': (f105, 76, 3),
    'D76': (f105, 76, 4),
    'C77': (f106, 77, 3),
    'D77': (f106, 77, 4),
    'C78': (f107, 78, 3),
    'D78': (f107, 78, 4),
    'C81': (f108, 81, 3),
    'D81': (f108, 81, 4),
    'C82': (f109, 82, 3),
    'D82': (f109, 82, 4),
    'C83': (f110, 83, 3),
    'D83': (f110, 83, 4),
    'C84': (f111, 84, 3),
    'D84': (f111, 84, 4),
    'C85': (f112, 85, 3),
    'D85': (f112, 85, 4),
    'C86': (f113, 86, 3),
    'D86': (f113, 86, 4),
    'C87': (f114, 87, 3),
    'D87': (f114, 87, 4),
    'C88': (f115, 88, 3),
    'D88': (f115, 88, 4),
    'C89': (f116, 89, 3),
    'D89': (f116, 89, 4),
    'C90': (f117, 90, 3),
    'D90': (f117, 90, 4),
    'C91': (f118, 91, 3),
    'D91': (f118, 91, 4),
    'C92': (f119, 92, 3),
    'D92': (f119, 92, 4),
    'C93': (f120, 93, 3),
    'D93': (f120, 93, 4),
    'C94': (f121, 94, 3),
    'D94': (f121, 94, 4),
    'C95': (f122, 95, 3),
    'D95': (f122, 95, 4),
    'C96': (f123, 96, 3),
    'D96': (f123, 96, 4),
    'C97': (f124, 97, 3),
    'D97': (f124, 97, 4),
    'C98': (f125, 98, 3),
    'D98': (f125, 98, 4),
    'C99': (f126, 99, 3),
    'D99': (f126, 99, 4),
    'C102': (f127, 102, 3),
    'D102': (f127, 102, 4),
    'C103': (f128, 103, 3),
    'D103': (f128, 103, 4),
    'C104': (f129, 104, 3),
    'D104': (f129, 104, 4),
    'C105': (f130, 105, 3),
    'D105': (f130, 105, 4),
    'C106': (f131, 106, 3),
    'D106': (f131, 106, 4),
    'C107': (f131, 107, 3),
    'D107': (f131, 107, 4),
    'C108': (f131, 108, 3),
    'D108': (f131, 108, 4),
    'C109': (f132, 109, 3),
    'D109': (f132, 109, 4),
    'C110': (f132, 110, 3),
    'D110': (f132, 110, 4),
    'C111': (f132, 111, 3),
    'D111': (f132, 111, 4),
    'C112': (f133, 112, 3),
    'D112': (f133, 112, 4),
    'C113': (f134, 113, 3),
    'D113': (f134, 113, 4),
    'C114': (f135, 114, 3),
    'D114': (f135, 114, 4),
    'C115': (f136, 115, 3),
    'D115': (f136, 115, 4),
    'C120': (f137, 120, 3),
    'D120': (f138, 120, 4),
    'E120': (f139, 120, 5),
    'F120': (f140, 120, 6),
    'C121': (f141, 121, 3),
    'D121': (f142, 121, 4),
    'E121': (f143, 121, 5),
    'F121': (f144, 121, 6),
    'C122': (f145, 122, 3),
    'D122': (f146, 122, 4),
    'E122': (f147, 122, 5),
    'F122': (f148, 122, 6),
    'C123': (f149, 123, 3),
    'D123': (f150, 123, 4),
    'E123': (f151, 123, 5),
    'F123': (f152, 123, 6),
    'C124': (f153, 124, 3),
    'D124': (f154, 124, 4),
    'E124': (f155, 124, 5),
    'F124': (f156, 124, 6),
    'C125': (f157, 125, 3),
    'D125': (f158, 125, 4),
    'E125': (f159, 125, 5),
    'F125': (f160, 125, 6),
    'C126': (f161, 126, 3),
    'D126': (f162, 126, 4),
    'E126': (f163, 126, 5),
    'F126': (f164, 126, 6),
    'C127': (f165, 127, 3),
    'D127': (f166, 127, 4),
    'E127': (f167, 127, 5),
    'F127': (f168, 127, 6),
    'C128': (f169, 128, 3),
    'D128': (f170, 128, 4),
    'E128': (f171, 128, 5),
    'F128': (f172, 128, 6),
    'C129': (f173, 129, 3),
    'D129': (f174, 129, 4),
    'E129': (f175, 129, 5),
    'F129': (f176, 129, 6),
    'C130': (f177, 130, 3),
    'D130': (f178, 130, 4),
    'E130': (f179, 130, 5),
    'F130': (f180, 130, 6),
    'C131': (f181, 131, 3),
    'D131': (f182, 131, 4),
    'E131': (f183, 131, 5),
    'F131': (f184, 131, 6),
    'C132': (f185, 132, 3),
    'D132': (f186, 132, 4),
    'E132': (f187, 132, 5),
    'F132': (f188, 132, 6),
    'C133': (f189, 133, 3),
    'D133': (f190, 133, 4),
    'E133': (f191, 133, 5),
    'F133': (f192, 133, 6),
    'C134': (f193, 134, 3),
    'D134': (f194, 134, 4),
    'E134': (f195, 134, 5),
    'F134': (f196, 134, 6),
    'C136': (f197, 136, 3),
    'C137': (f198, 137, 3),
    'C140': (f199, 140, 3),
    'C141': (f200, 141, 3),
    'C142': (f201, 142, 3),
    'C143': (f202, 143, 3),
    'C144': (f203, 144, 3),
    'C145': (f204, 145, 3),
    'C147': (f205, 147, 3),
    'C148': (f206, 148, 3),
    'C149': (f207, 149, 3),
    'C151': (f208, 151, 3),
    'F151': (f209, 151, 6),
    'G151': (f210, 151, 7),
    'C154': (f211, 154, 3),
    'C155': (f212, 155, 3),
    'C156': (f213, 156, 3),
    'C157': (f214, 157, 3),
    'C158': (f214, 158, 3),
    'C159': (f214, 159, 3),
    'C160': (f215, 160, 3),
    'C161': (f216, 161, 3),
    'C162': (f217, 162, 3),
    'J162': (f218, 162, 10),
    'K162': (f219, 162, 11),
    'L162': (f220, 162, 12),
    'M162': (f221, 162, 13),
    'C163': (f222, 163, 3),
    'C164': (f223, 164, 3),
    'C165': (f224, 165, 3),
    'C166': (f225, 166, 3),
    'C167': (f226, 167, 3),
    'D167': (f227, 167, 4),
    'E167': (f228, 167, 5),
    'F167': (f229, 167, 6),
    'G167': (f230, 167, 7),
    'C168': (f231, 168, 3),
    'D168': (f231, 168, 4),
    'E168': (f231, 168, 5),
    'F168': (f231, 168, 6),
    'G168': (f231, 168, 7),
    'C169': (f232, 169, 3),
    'C170': (f233, 170, 3),
    'C171': (f234, 171, 3),
    'C172': (f235, 172, 3),
    'C173': (f236, 173, 3),
    'D173': (f237, 173, 4),
    'E173': (f238, 173, 5),
    'F173': (f239, 173, 6),
    'C174': (f240, 174, 3),
    'C175': (f241, 175, 3),
    'J175': (f242, 175, 10),
    'K175': (f243, 175, 11),
    'C176': (f244, 176, 3),
    'C177': (f245, 177, 3),
    'C178': (f246, 178, 3),
    'C179': (f247, 179, 3),
    'C180': (f248, 180, 3),
    'J180': (f249, 180, 10),
    'K180': (f250, 180, 11),
    'C181': (f251, 181, 3),
    'J181': (f252, 181, 10),
    'K181': (f253, 181, 11),
    'L181': (f254, 181, 12),
    'C182': (f255, 182, 3),
    'C183': (f256, 183, 3),
    'C184': (f257, 184, 3),
    'C185': (f258, 185, 3),
    'C186': (f259, 186, 3),
    'J187': (f260, 187, 10),
    'K187': (f261, 187, 11),
    'L187': (f262, 187, 12),
    'M187': (f263, 187, 13),
    'N187': (f264, 187, 14),
    'O187': (f265, 187, 15),
    'C188': (f266, 188, 3),
    'J188': (f267, 188, 10),
    'K188': (f268, 188, 11),
    'L188': (f269, 188, 12),
    'M188': (f270, 188, 13),
    'P188': (f271, 188, 16),
    'C189': (f272, 189, 3),
    'J189': (f273, 189, 10),
    'K189': (f274, 189, 11),
    'L189': (f275, 189, 12),
    'D190': (f276, 190, 4),
    'D192': (f277, 192, 4),
    'D193': (f278, 193, 4),
    'D194': (f279, 194, 4),
    'D195': (f280, 195, 4),
    'C198': (f281, 198, 3),
    'D198': (f282, 198, 4),
    'C202': (f283, 202, 3),
    'D202': (f283, 202, 4),
    'C203': (f284, 203, 3),
    'D203': (f285, 203, 4),
    'C204': (f286, 204, 3),
    'D204': (f287, 204, 4),
    'C205': (f288, 205, 3),
    'D205': (f288, 205, 4),
    'C206': (f289, 206, 3),
    'D206': (f289, 206, 4),
    'C209': (f290, 209, 3),
    'D209': (f290, 209, 4),
    'C210': (f291, 210, 3),
    'D210': (f291, 210, 4),
    'C211': (f292, 211, 3),
    'D211': (f292, 211, 4),
    'C212': (f293, 212, 3),
    'D212': (f293, 212, 4),
    'C213': (f294, 213, 3),
    'D213': (f294, 213, 4),
    'C214': (f295, 214, 3),
    'D214': (f295, 214, 4),
    'C215': (f296, 215, 3),
    'D215': (f297, 215, 4),
    'C216': (f298, 216, 3),
    'D216': (f299, 216, 4),
    'C217': (f300, 217, 3),
    'D217': (f300, 217, 4),
    'C218': (f301, 218, 3),
    'D218': (f301, 218, 4),
    'C219': (f302, 219, 3),
    'D219': (f302, 219, 4),
}
