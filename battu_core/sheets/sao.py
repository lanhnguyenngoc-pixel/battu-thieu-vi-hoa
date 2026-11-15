# -*- coding: utf-8 -*-
"""Module sheet SAO – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.5.xlsx (SHA-256 f19f3f74b720c1a067a640a31d58441793bf64bfdd12629f17ba4be6623c56f4).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'SAO'
SI = 13

def f0(E, r, c):
    # Ô D5
    # = TRUNG_GIAN!$C$5
    return E.v(23,5,3)

def f1(E, r, c):
    # Ô E5, E6, E7, E8, E9, E10, E11, E12, E13, E14, E15, E16 … (16 ô)
    # = IF(X5=0,"",INDEX(CAN_TEN,D5+1))
    return ('' if X.truthy(X.eq(E.v(13,r,c+19),0.0)) else X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(13,r,c-1),1.0)))

def f2(E, r, c):
    # Ô F5, F6, F7, F8, F9, F10, F11, F12, F13, F14, F15, F16 … (16 ô)
    # = IF(X5=0,"",IF(Y5=1,"Nhật chủ",INDEX(M_TT,TRUNG_GIAN!$E$5+1,D5+1)))
    return ('' if X.truthy(X.eq(E.v(13,r,c+18),0.0)) else ('Nhật chủ' if X.truthy(X.eq(E.v(13,r,c+19),1.0)) else X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(13,r,c-2),1.0))))

def f3(E, r, c):
    # Ô G5, G6, G7, G8, G9, G10, G11, G12, G13, G14, G15, G16 … (16 ô)
    # = IF(X5=0,-1,MOD(INDEX(CSDL!$C$5:$C$14,D5+1)-TRUNG_GIAN!$E$7,5))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+17),0.0)) else X.MOD(X.sub(X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(13,r,c-3),1.0)),E.v(23,7,5)), 5.0))

def f4(E, r, c):
    # Ô H5, H6, H7, H8, H9, H10, H11, H12, H13, H14, H15, H16 … (16 ô)
    # = IF(X5=0,-1,INDEX(CSDL!$C$5:$C$14,D5+1))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+16),0.0)) else X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(13,r,c-4),1.0)))

def f5(E, r, c):
    # Ô I5, I6, I7, I8, I9, I10, I11, I12, I13, I14, I15, I16 … (16 ô)
    # = IF(X5=0,-1,INDEX(M_TS,D5+1,TRUNG_GIAN!$D$6+1))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+15),0.0)) else X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(13,r,c-5),1.0), X.add(E.v(23,6,4),1.0)))

def f6(E, r, c):
    # Ô J5, J6, J7, J8, J9, J10, J11, J12, J13, J14, J15, J16 … (16 ô)
    # = IF(AND(X5=1,I5<=4,I5>=0),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(13,r,c+14),1.0), X.le(E.v(13,r,c-1),4.0), X.ge(E.v(13,r,c-1),0.0))) else 0.0)

def f7(E, r, c):
    # Ô K5, K6, K7, K8
    # = IF(X5=0,-1,INDEX(M_TS,D5+1,TRUNG_GIAN!$C$6+1))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+13),0.0)) else X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(13,r,c-7),1.0), X.add(E.v(23,6,3),1.0)))

def f8(E, r, c):
    # Ô L5
    # = IF(OR(PP1_DINH_TINH!$C$38=1,PP1_DINH_TINH!$C$40=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(9,40,3),1.0))) else 0.0)

def f9(E, r, c):
    # Ô M5, M6, M7, M8
    # = IF(X5=0,0,PP1_DINH_TINH!$C$29)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+11),0.0)) else E.v(9,29,3))

def f10(E, r, c):
    # Ô N5, N6, N7, N8
    # = IF(X5=0,0,THAN_SAT!$C$13)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+10),0.0)) else E.v(12,13,3))

def f11(E, r, c):
    # Ô O5, O9, O13, O17
    # = IF(X5=0,0,IF(OR(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",D5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),D5=INDEX(CSDL!$D$122:$D$133,TRUN
    #   G_GIAN!$D$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+9),0.0)) else (1.0 if X.truthy(X.OR(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(13,r,c-11),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.eq(E.v(13,r,c-11),X.INDEX(E.rg(24,122,4,133,4), X.add(E.v(23,6,4),1.0))))) else 0.0))

def f12(E, r, c):
    # Ô P5, P6, P7, P8
    # = IF(X5=0,0,IF(OR(TRUNG_GIAN!$C$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$C$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$C$6=INDE
    #   X(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),TRUNG_GIAN!$C$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+8),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0))), X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0))))) else 0.0))

def f13(E, r, c):
    # Ô Q5, Q6, Q7, Q8
    # = IF(X5=0,0,IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$G$5:$G$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+7),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,7,14,7), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f14(E, r, c):
    # Ô R5, R6, R7, R8
    # = IF(X5=0,0,IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,8,14,8), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f15(E, r, c):
    # Ô S5, S6, S7, S8
    # = IF(X5=0,0,IF(OR(TRUNG_GIAN!$C$6=INDEX(CSDL!$G$122:$G$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+5),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,7,133,7), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f16(E, r, c):
    # Ô T5, T6, T7, T8
    # = IF(X5=0,0,IF(OR(TRUNG_GIAN!$C$6=INDEX(CSDL!$E$122:$E$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+4),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,5,133,5), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f17(E, r, c):
    # Ô U5, U6, U7, U8
    # = IF(X5=0,0,IF(OR(TRUNG_GIAN!$C$6=INDEX(CSDL!$H$122:$H$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+3),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,8,133,8), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f18(E, r, c):
    # Ô V5, V6, V7, V8
    # = IF(X5=0,0,IF(OR(TRUNG_GIAN!$C$6=INDEX(CSDL!$I$122:$I$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+2),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,9,133,9), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f19(E, r, c):
    # Ô W5, W6, W7, W8
    # = IF(X5=0,0,IF(OR(TRUNG_GIAN!$C$6=INDEX(CSDL!$F$122:$F$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+1),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,6,133,6), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f20(E, r, c):
    # Ô X5, X6, X7, X8, X9, X10, X11, X12, X13, X14, X15, X16 … (16 ô)
    # = IF(D5>=0,1,0)
    return (1.0 if X.truthy(X.ge(E.v(13,r,c-20),0.0)) else 0.0)

def f21(E, r, c):
    # Ô Y5, L6, Y6, L7, Y7, L8, Y8, Y9, L10, Y10, L11, Y11 … (27 ô)
    # = 0
    return 0.0

def f22(E, r, c):
    # Ô D6
    # = TRUNG_GIAN!$C$15
    return E.v(23,15,3)

def f23(E, r, c):
    # Ô O6, O7, O8
    # = IF(X6=0,0,IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$C$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+9),0.0)) else (1.0 if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else 0.0))

def f24(E, r, c):
    # Ô D7
    # = TRUNG_GIAN!$C$17
    return E.v(23,17,3)

def f25(E, r, c):
    # Ô D8
    # = TRUNG_GIAN!$C$19
    return E.v(23,19,3)

def f26(E, r, c):
    # Ô D9
    # = TRUNG_GIAN!$D$5
    return E.v(23,5,4)

def f27(E, r, c):
    # Ô K9, K10, K11, K12
    # = IF(X9=0,-1,INDEX(M_TS,D9+1,TRUNG_GIAN!$D$6+1))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+13),0.0)) else X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(13,r,c-7),1.0), X.add(E.v(23,6,4),1.0)))

def f28(E, r, c):
    # Ô L9
    # = IF(OR(PP1_DINH_TINH!$D$38=1,PP1_DINH_TINH!$D$40=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(9,40,4),1.0))) else 0.0)

def f29(E, r, c):
    # Ô M9, M10, M11, M12
    # = IF(X9=0,0,PP1_DINH_TINH!$D$29)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+11),0.0)) else E.v(9,29,4))

def f30(E, r, c):
    # Ô N9, N10, N11, N12
    # = IF(X9=0,0,THAN_SAT!$D$13)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+10),0.0)) else E.v(12,13,4))

def f31(E, r, c):
    # Ô P9, P10, P11, P12
    # = IF(X9=0,0,IF(OR(TRUNG_GIAN!$D$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$D$6=INDE
    #   X(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+8),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0))))) else 0.0))

def f32(E, r, c):
    # Ô Q9, Q10, Q11, Q12
    # = IF(X9=0,0,IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$G$5:$G$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+7),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,7,14,7), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f33(E, r, c):
    # Ô R9, R10, R11, R12
    # = IF(X9=0,0,IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,8,14,8), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f34(E, r, c):
    # Ô S9, S10, S11, S12
    # = IF(X9=0,0,IF(OR(TRUNG_GIAN!$D$6=INDEX(CSDL!$G$122:$G$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$G$122:$G$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+5),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,7,133,7), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,7,133,7), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f35(E, r, c):
    # Ô T9, T10, T11, T12
    # = IF(X9=0,0,IF(OR(TRUNG_GIAN!$D$6=INDEX(CSDL!$E$122:$E$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$E$122:$E$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+4),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,5,133,5), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,5,133,5), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f36(E, r, c):
    # Ô U9, U10, U11, U12
    # = IF(X9=0,0,IF(OR(TRUNG_GIAN!$D$6=INDEX(CSDL!$H$122:$H$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$H$122:$H$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+3),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,8,133,8), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,8,133,8), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f37(E, r, c):
    # Ô V9, V10, V11, V12
    # = IF(X9=0,0,IF(OR(TRUNG_GIAN!$D$6=INDEX(CSDL!$I$122:$I$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$I$122:$I$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+2),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,9,133,9), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,9,133,9), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f38(E, r, c):
    # Ô W9, W10, W11, W12
    # = IF(X9=0,0,IF(OR(TRUNG_GIAN!$D$6=INDEX(CSDL!$F$122:$F$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$D$6=INDEX(CSDL!$F$122:$F$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+1),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,6,133,6), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,6,133,6), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f39(E, r, c):
    # Ô D10
    # = TRUNG_GIAN!$D$15
    return E.v(23,15,4)

def f40(E, r, c):
    # Ô O10, O11, O12
    # = IF(X10=0,0,IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$D$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+9),0.0)) else (1.0 if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else 0.0))

def f41(E, r, c):
    # Ô D11
    # = TRUNG_GIAN!$D$17
    return E.v(23,17,4)

def f42(E, r, c):
    # Ô D12
    # = TRUNG_GIAN!$D$19
    return E.v(23,19,4)

def f43(E, r, c):
    # Ô D13
    # = TRUNG_GIAN!$E$5
    return E.v(23,5,5)

def f44(E, r, c):
    # Ô K13, K14, K15, K16
    # = IF(X13=0,-1,INDEX(M_TS,D13+1,TRUNG_GIAN!$E$6+1))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+13),0.0)) else X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(13,r,c-7),1.0), X.add(E.v(23,6,5),1.0)))

def f45(E, r, c):
    # Ô L13
    # = IF(OR(PP1_DINH_TINH!$E$38=1,PP1_DINH_TINH!$E$40=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,38,5),1.0), X.eq(E.v(9,40,5),1.0))) else 0.0)

def f46(E, r, c):
    # Ô M13, M14, M15, M16
    # = IF(X13=0,0,PP1_DINH_TINH!$E$29)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+11),0.0)) else E.v(9,29,5))

def f47(E, r, c):
    # Ô N13, N14, N15, N16
    # = IF(X13=0,0,THAN_SAT!$E$13)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+10),0.0)) else E.v(12,13,5))

def f48(E, r, c):
    # Ô P13, P14, P15, P16
    # = IF(X13=0,0,IF(OR(TRUNG_GIAN!$E$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$E$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$E$6=IND
    #   EX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),TRUNG_GIAN!$E$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+8),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0))), X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0))))) else 0.0))

def f49(E, r, c):
    # Ô Q13, Q14, Q15, Q16
    # = IF(X13=0,0,IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$G$5:$G$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+7),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,7,14,7), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f50(E, r, c):
    # Ô R13, R14, R15, R16
    # = IF(X13=0,0,IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,8,14,8), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f51(E, r, c):
    # Ô S13, S14, S15, S16
    # = IF(X13=0,0,IF(OR(TRUNG_GIAN!$E$6=INDEX(CSDL!$G$122:$G$133,TRUNG_GIAN!$C$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+5),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,7,133,7), X.add(E.v(23,6,3),1.0))))) else 0.0))

def f52(E, r, c):
    # Ô T13, T14, T15, T16
    # = IF(X13=0,0,IF(OR(TRUNG_GIAN!$E$6=INDEX(CSDL!$E$122:$E$133,TRUNG_GIAN!$C$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+4),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,5,133,5), X.add(E.v(23,6,3),1.0))))) else 0.0))

def f53(E, r, c):
    # Ô U13, U14, U15, U16
    # = IF(X13=0,0,IF(OR(TRUNG_GIAN!$E$6=INDEX(CSDL!$H$122:$H$133,TRUNG_GIAN!$C$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+3),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,8,133,8), X.add(E.v(23,6,3),1.0))))) else 0.0))

def f54(E, r, c):
    # Ô V13, V14, V15, V16
    # = IF(X13=0,0,IF(OR(TRUNG_GIAN!$E$6=INDEX(CSDL!$I$122:$I$133,TRUNG_GIAN!$C$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+2),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,9,133,9), X.add(E.v(23,6,3),1.0))))) else 0.0))

def f55(E, r, c):
    # Ô W13, W14, W15, W16
    # = IF(X13=0,0,IF(OR(TRUNG_GIAN!$E$6=INDEX(CSDL!$F$122:$F$133,TRUNG_GIAN!$C$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+1),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,6,133,6), X.add(E.v(23,6,3),1.0))))) else 0.0))

def f56(E, r, c):
    # Ô Y13
    # = 1
    return 1.0

def f57(E, r, c):
    # Ô D14
    # = TRUNG_GIAN!$E$15
    return E.v(23,15,5)

def f58(E, r, c):
    # Ô O14, O15, O16
    # = IF(X14=0,0,IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$E$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+9),0.0)) else (1.0 if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else 0.0))

def f59(E, r, c):
    # Ô D15
    # = TRUNG_GIAN!$E$17
    return E.v(23,17,5)

def f60(E, r, c):
    # Ô D16
    # = TRUNG_GIAN!$E$19
    return E.v(23,19,5)

def f61(E, r, c):
    # Ô D17
    # = TRUNG_GIAN!$F$5
    return E.v(23,5,6)

def f62(E, r, c):
    # Ô K17, K18, K19, K20
    # = IF(X17=0,-1,INDEX(M_TS,D17+1,TRUNG_GIAN!$F$6+1))
    return (X.neg(1.0) if X.truthy(X.eq(E.v(13,r,c+13),0.0)) else X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(13,r,c-7),1.0), X.add(E.v(23,6,6),1.0)))

def f63(E, r, c):
    # Ô L17
    # = IF(OR(PP1_DINH_TINH!$F$38=1,PP1_DINH_TINH!$F$40=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(9,38,6),1.0), X.eq(E.v(9,40,6),1.0))) else 0.0)

def f64(E, r, c):
    # Ô M17, M18, M19, M20
    # = IF(X17=0,0,PP1_DINH_TINH!$F$29)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+11),0.0)) else E.v(9,29,6))

def f65(E, r, c):
    # Ô N17, N18, N19, N20
    # = IF(X17=0,0,THAN_SAT!$F$13)
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+10),0.0)) else E.v(12,13,6))

def f66(E, r, c):
    # Ô P17, P18, P19, P20
    # = IF(X17=0,0,IF(OR(TRUNG_GIAN!$F$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),TRUNG_GIAN!$F$6=IND
    #   EX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+8),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0))))) else 0.0))

def f67(E, r, c):
    # Ô Q17, Q18, Q19, Q20
    # = IF(X17=0,0,IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$G$5:$G$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+7),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,7,14,7), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f68(E, r, c):
    # Ô R17, R18, R19, R20
    # = IF(X17=0,0,IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+6),0.0)) else (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,8,14,8), X.add(E.v(23,5,5),1.0)))) else 0.0))

def f69(E, r, c):
    # Ô S17, S18, S19, S20
    # = IF(X17=0,0,IF(OR(TRUNG_GIAN!$F$6=INDEX(CSDL!$G$122:$G$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$G$122:$G$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+5),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,7,133,7), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,7,133,7), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f70(E, r, c):
    # Ô T17, T18, T19, T20
    # = IF(X17=0,0,IF(OR(TRUNG_GIAN!$F$6=INDEX(CSDL!$E$122:$E$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$E$122:$E$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+4),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,5,133,5), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,5,133,5), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f71(E, r, c):
    # Ô U17, U18, U19, U20
    # = IF(X17=0,0,IF(OR(TRUNG_GIAN!$F$6=INDEX(CSDL!$H$122:$H$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$H$122:$H$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+3),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,8,133,8), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,8,133,8), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f72(E, r, c):
    # Ô V17, V18, V19, V20
    # = IF(X17=0,0,IF(OR(TRUNG_GIAN!$F$6=INDEX(CSDL!$I$122:$I$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$I$122:$I$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+2),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,9,133,9), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,9,133,9), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f73(E, r, c):
    # Ô W17, W18, W19, W20
    # = IF(X17=0,0,IF(OR(TRUNG_GIAN!$F$6=INDEX(CSDL!$F$122:$F$133,TRUNG_GIAN!$C$6+1),TRUNG_GIAN!$F$6=INDEX(CSDL!$F$122:$F$133,TRUNG_GIAN!$E$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+1),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,6,133,6), X.add(E.v(23,6,3),1.0))), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,6,133,6), X.add(E.v(23,6,5),1.0))))) else 0.0))

def f74(E, r, c):
    # Ô D18
    # = TRUNG_GIAN!$F$15
    return E.v(23,15,6)

def f75(E, r, c):
    # Ô O18, O19, O20
    # = IF(X18=0,0,IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$F$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),1,0))
    return (0.0 if X.truthy(X.eq(E.v(13,r,c+9),0.0)) else (1.0 if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else 0.0))

def f76(E, r, c):
    # Ô D19
    # = TRUNG_GIAN!$F$17
    return E.v(23,17,6)

def f77(E, r, c):
    # Ô D20
    # = TRUNG_GIAN!$F$19
    return E.v(23,19,6)

CELLS = {
    'D5': (f0, 5, 4),
    'E5': (f1, 5, 5),
    'F5': (f2, 5, 6),
    'G5': (f3, 5, 7),
    'H5': (f4, 5, 8),
    'I5': (f5, 5, 9),
    'J5': (f6, 5, 10),
    'K5': (f7, 5, 11),
    'L5': (f8, 5, 12),
    'M5': (f9, 5, 13),
    'N5': (f10, 5, 14),
    'O5': (f11, 5, 15),
    'P5': (f12, 5, 16),
    'Q5': (f13, 5, 17),
    'R5': (f14, 5, 18),
    'S5': (f15, 5, 19),
    'T5': (f16, 5, 20),
    'U5': (f17, 5, 21),
    'V5': (f18, 5, 22),
    'W5': (f19, 5, 23),
    'X5': (f20, 5, 24),
    'Y5': (f21, 5, 25),
    'D6': (f22, 6, 4),
    'E6': (f1, 6, 5),
    'F6': (f2, 6, 6),
    'G6': (f3, 6, 7),
    'H6': (f4, 6, 8),
    'I6': (f5, 6, 9),
    'J6': (f6, 6, 10),
    'K6': (f7, 6, 11),
    'L6': (f21, 6, 12),
    'M6': (f9, 6, 13),
    'N6': (f10, 6, 14),
    'O6': (f23, 6, 15),
    'P6': (f12, 6, 16),
    'Q6': (f13, 6, 17),
    'R6': (f14, 6, 18),
    'S6': (f15, 6, 19),
    'T6': (f16, 6, 20),
    'U6': (f17, 6, 21),
    'V6': (f18, 6, 22),
    'W6': (f19, 6, 23),
    'X6': (f20, 6, 24),
    'Y6': (f21, 6, 25),
    'D7': (f24, 7, 4),
    'E7': (f1, 7, 5),
    'F7': (f2, 7, 6),
    'G7': (f3, 7, 7),
    'H7': (f4, 7, 8),
    'I7': (f5, 7, 9),
    'J7': (f6, 7, 10),
    'K7': (f7, 7, 11),
    'L7': (f21, 7, 12),
    'M7': (f9, 7, 13),
    'N7': (f10, 7, 14),
    'O7': (f23, 7, 15),
    'P7': (f12, 7, 16),
    'Q7': (f13, 7, 17),
    'R7': (f14, 7, 18),
    'S7': (f15, 7, 19),
    'T7': (f16, 7, 20),
    'U7': (f17, 7, 21),
    'V7': (f18, 7, 22),
    'W7': (f19, 7, 23),
    'X7': (f20, 7, 24),
    'Y7': (f21, 7, 25),
    'D8': (f25, 8, 4),
    'E8': (f1, 8, 5),
    'F8': (f2, 8, 6),
    'G8': (f3, 8, 7),
    'H8': (f4, 8, 8),
    'I8': (f5, 8, 9),
    'J8': (f6, 8, 10),
    'K8': (f7, 8, 11),
    'L8': (f21, 8, 12),
    'M8': (f9, 8, 13),
    'N8': (f10, 8, 14),
    'O8': (f23, 8, 15),
    'P8': (f12, 8, 16),
    'Q8': (f13, 8, 17),
    'R8': (f14, 8, 18),
    'S8': (f15, 8, 19),
    'T8': (f16, 8, 20),
    'U8': (f17, 8, 21),
    'V8': (f18, 8, 22),
    'W8': (f19, 8, 23),
    'X8': (f20, 8, 24),
    'Y8': (f21, 8, 25),
    'D9': (f26, 9, 4),
    'E9': (f1, 9, 5),
    'F9': (f2, 9, 6),
    'G9': (f3, 9, 7),
    'H9': (f4, 9, 8),
    'I9': (f5, 9, 9),
    'J9': (f6, 9, 10),
    'K9': (f27, 9, 11),
    'L9': (f28, 9, 12),
    'M9': (f29, 9, 13),
    'N9': (f30, 9, 14),
    'O9': (f11, 9, 15),
    'P9': (f31, 9, 16),
    'Q9': (f32, 9, 17),
    'R9': (f33, 9, 18),
    'S9': (f34, 9, 19),
    'T9': (f35, 9, 20),
    'U9': (f36, 9, 21),
    'V9': (f37, 9, 22),
    'W9': (f38, 9, 23),
    'X9': (f20, 9, 24),
    'Y9': (f21, 9, 25),
    'D10': (f39, 10, 4),
    'E10': (f1, 10, 5),
    'F10': (f2, 10, 6),
    'G10': (f3, 10, 7),
    'H10': (f4, 10, 8),
    'I10': (f5, 10, 9),
    'J10': (f6, 10, 10),
    'K10': (f27, 10, 11),
    'L10': (f21, 10, 12),
    'M10': (f29, 10, 13),
    'N10': (f30, 10, 14),
    'O10': (f40, 10, 15),
    'P10': (f31, 10, 16),
    'Q10': (f32, 10, 17),
    'R10': (f33, 10, 18),
    'S10': (f34, 10, 19),
    'T10': (f35, 10, 20),
    'U10': (f36, 10, 21),
    'V10': (f37, 10, 22),
    'W10': (f38, 10, 23),
    'X10': (f20, 10, 24),
    'Y10': (f21, 10, 25),
    'D11': (f41, 11, 4),
    'E11': (f1, 11, 5),
    'F11': (f2, 11, 6),
    'G11': (f3, 11, 7),
    'H11': (f4, 11, 8),
    'I11': (f5, 11, 9),
    'J11': (f6, 11, 10),
    'K11': (f27, 11, 11),
    'L11': (f21, 11, 12),
    'M11': (f29, 11, 13),
    'N11': (f30, 11, 14),
    'O11': (f40, 11, 15),
    'P11': (f31, 11, 16),
    'Q11': (f32, 11, 17),
    'R11': (f33, 11, 18),
    'S11': (f34, 11, 19),
    'T11': (f35, 11, 20),
    'U11': (f36, 11, 21),
    'V11': (f37, 11, 22),
    'W11': (f38, 11, 23),
    'X11': (f20, 11, 24),
    'Y11': (f21, 11, 25),
    'D12': (f42, 12, 4),
    'E12': (f1, 12, 5),
    'F12': (f2, 12, 6),
    'G12': (f3, 12, 7),
    'H12': (f4, 12, 8),
    'I12': (f5, 12, 9),
    'J12': (f6, 12, 10),
    'K12': (f27, 12, 11),
    'L12': (f21, 12, 12),
    'M12': (f29, 12, 13),
    'N12': (f30, 12, 14),
    'O12': (f40, 12, 15),
    'P12': (f31, 12, 16),
    'Q12': (f32, 12, 17),
    'R12': (f33, 12, 18),
    'S12': (f34, 12, 19),
    'T12': (f35, 12, 20),
    'U12': (f36, 12, 21),
    'V12': (f37, 12, 22),
    'W12': (f38, 12, 23),
    'X12': (f20, 12, 24),
    'Y12': (f21, 12, 25),
    'D13': (f43, 13, 4),
    'E13': (f1, 13, 5),
    'F13': (f2, 13, 6),
    'G13': (f3, 13, 7),
    'H13': (f4, 13, 8),
    'I13': (f5, 13, 9),
    'J13': (f6, 13, 10),
    'K13': (f44, 13, 11),
    'L13': (f45, 13, 12),
    'M13': (f46, 13, 13),
    'N13': (f47, 13, 14),
    'O13': (f11, 13, 15),
    'P13': (f48, 13, 16),
    'Q13': (f49, 13, 17),
    'R13': (f50, 13, 18),
    'S13': (f51, 13, 19),
    'T13': (f52, 13, 20),
    'U13': (f53, 13, 21),
    'V13': (f54, 13, 22),
    'W13': (f55, 13, 23),
    'X13': (f20, 13, 24),
    'Y13': (f56, 13, 25),
    'D14': (f57, 14, 4),
    'E14': (f1, 14, 5),
    'F14': (f2, 14, 6),
    'G14': (f3, 14, 7),
    'H14': (f4, 14, 8),
    'I14': (f5, 14, 9),
    'J14': (f6, 14, 10),
    'K14': (f44, 14, 11),
    'L14': (f21, 14, 12),
    'M14': (f46, 14, 13),
    'N14': (f47, 14, 14),
    'O14': (f58, 14, 15),
    'P14': (f48, 14, 16),
    'Q14': (f49, 14, 17),
    'R14': (f50, 14, 18),
    'S14': (f51, 14, 19),
    'T14': (f52, 14, 20),
    'U14': (f53, 14, 21),
    'V14': (f54, 14, 22),
    'W14': (f55, 14, 23),
    'X14': (f20, 14, 24),
    'Y14': (f21, 14, 25),
    'D15': (f59, 15, 4),
    'E15': (f1, 15, 5),
    'F15': (f2, 15, 6),
    'G15': (f3, 15, 7),
    'H15': (f4, 15, 8),
    'I15': (f5, 15, 9),
    'J15': (f6, 15, 10),
    'K15': (f44, 15, 11),
    'L15': (f21, 15, 12),
    'M15': (f46, 15, 13),
    'N15': (f47, 15, 14),
    'O15': (f58, 15, 15),
    'P15': (f48, 15, 16),
    'Q15': (f49, 15, 17),
    'R15': (f50, 15, 18),
    'S15': (f51, 15, 19),
    'T15': (f52, 15, 20),
    'U15': (f53, 15, 21),
    'V15': (f54, 15, 22),
    'W15': (f55, 15, 23),
    'X15': (f20, 15, 24),
    'Y15': (f21, 15, 25),
    'D16': (f60, 16, 4),
    'E16': (f1, 16, 5),
    'F16': (f2, 16, 6),
    'G16': (f3, 16, 7),
    'H16': (f4, 16, 8),
    'I16': (f5, 16, 9),
    'J16': (f6, 16, 10),
    'K16': (f44, 16, 11),
    'L16': (f21, 16, 12),
    'M16': (f46, 16, 13),
    'N16': (f47, 16, 14),
    'O16': (f58, 16, 15),
    'P16': (f48, 16, 16),
    'Q16': (f49, 16, 17),
    'R16': (f50, 16, 18),
    'S16': (f51, 16, 19),
    'T16': (f52, 16, 20),
    'U16': (f53, 16, 21),
    'V16': (f54, 16, 22),
    'W16': (f55, 16, 23),
    'X16': (f20, 16, 24),
    'Y16': (f21, 16, 25),
    'D17': (f61, 17, 4),
    'E17': (f1, 17, 5),
    'F17': (f2, 17, 6),
    'G17': (f3, 17, 7),
    'H17': (f4, 17, 8),
    'I17': (f5, 17, 9),
    'J17': (f6, 17, 10),
    'K17': (f62, 17, 11),
    'L17': (f63, 17, 12),
    'M17': (f64, 17, 13),
    'N17': (f65, 17, 14),
    'O17': (f11, 17, 15),
    'P17': (f66, 17, 16),
    'Q17': (f67, 17, 17),
    'R17': (f68, 17, 18),
    'S17': (f69, 17, 19),
    'T17': (f70, 17, 20),
    'U17': (f71, 17, 21),
    'V17': (f72, 17, 22),
    'W17': (f73, 17, 23),
    'X17': (f20, 17, 24),
    'Y17': (f21, 17, 25),
    'D18': (f74, 18, 4),
    'E18': (f1, 18, 5),
    'F18': (f2, 18, 6),
    'G18': (f3, 18, 7),
    'H18': (f4, 18, 8),
    'I18': (f5, 18, 9),
    'J18': (f6, 18, 10),
    'K18': (f62, 18, 11),
    'L18': (f21, 18, 12),
    'M18': (f64, 18, 13),
    'N18': (f65, 18, 14),
    'O18': (f75, 18, 15),
    'P18': (f66, 18, 16),
    'Q18': (f67, 18, 17),
    'R18': (f68, 18, 18),
    'S18': (f69, 18, 19),
    'T18': (f70, 18, 20),
    'U18': (f71, 18, 21),
    'V18': (f72, 18, 22),
    'W18': (f73, 18, 23),
    'X18': (f20, 18, 24),
    'Y18': (f21, 18, 25),
    'D19': (f76, 19, 4),
    'E19': (f1, 19, 5),
    'F19': (f2, 19, 6),
    'G19': (f3, 19, 7),
    'H19': (f4, 19, 8),
    'I19': (f5, 19, 9),
    'J19': (f6, 19, 10),
    'K19': (f62, 19, 11),
    'L19': (f21, 19, 12),
    'M19': (f64, 19, 13),
    'N19': (f65, 19, 14),
    'O19': (f75, 19, 15),
    'P19': (f66, 19, 16),
    'Q19': (f67, 19, 17),
    'R19': (f68, 19, 18),
    'S19': (f69, 19, 19),
    'T19': (f70, 19, 20),
    'U19': (f71, 19, 21),
    'V19': (f72, 19, 22),
    'W19': (f73, 19, 23),
    'X19': (f20, 19, 24),
    'Y19': (f21, 19, 25),
    'D20': (f77, 20, 4),
    'E20': (f1, 20, 5),
    'F20': (f2, 20, 6),
    'G20': (f3, 20, 7),
    'H20': (f4, 20, 8),
    'I20': (f5, 20, 9),
    'J20': (f6, 20, 10),
    'K20': (f62, 20, 11),
    'L20': (f21, 20, 12),
    'M20': (f64, 20, 13),
    'N20': (f65, 20, 14),
    'O20': (f75, 20, 15),
    'P20': (f66, 20, 16),
    'Q20': (f67, 20, 17),
    'R20': (f68, 20, 18),
    'S20': (f69, 20, 19),
    'T20': (f70, 20, 20),
    'U20': (f71, 20, 21),
    'V20': (f72, 20, 22),
    'W20': (f73, 20, 23),
    'X20': (f20, 20, 24),
    'Y20': (f21, 20, 25),
}
