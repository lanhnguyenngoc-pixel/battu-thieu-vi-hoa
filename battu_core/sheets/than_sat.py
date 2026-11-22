# -*- coding: utf-8 -*-
"""Module sheet THAN_SAT – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: B69n.xlsx (SHA-256 9228d1ab613379f50ca7126e315efcbb940599c6e1784212c489756674ca8b67).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'THAN_SAT'
SI = 12

def f0(E, r, c):
    # Ô C6
    # = TRUNG_GIAN!$C$11
    return E.v(23,11,3)

def f1(E, r, c):
    # Ô D6
    # = TRUNG_GIAN!$D$11
    return E.v(23,11,4)

def f2(E, r, c):
    # Ô E6
    # = TRUNG_GIAN!$E$11
    return E.v(23,11,5)

def f3(E, r, c):
    # Ô F6
    # = TRUNG_GIAN!$F$11
    return E.v(23,11,6)

def f4(E, r, c):
    # Ô C7
    # = IF(INDEX(M_XU,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$D$6+1)=1,"xung "&TRUNG_GIAN!$D$11&" ","")
    return (X.cat(X.cat('xung ',E.v(23,11,4)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else '')

def f5(E, r, c):
    # Ô D7
    # = IF(INDEX(M_XU,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$C$6+1)=1,"xung "&TRUNG_GIAN!$C$11&" ","")&IF(INDEX(M_XU,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$E$6+1)=1,"xung "&TRU
    #   NG_GIAN!$E$11&" ","")
    return X.cat((X.cat(X.cat('xung ',E.v(23,11,3)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(X.cat(X.cat('xung ',E.v(23,11,5)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else ''))

def f6(E, r, c):
    # Ô E7
    # = IF(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1,"xung "&TRUNG_GIAN!$D$11&" ","")&IF(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)=1,"xung "&TRU
    #   NG_GIAN!$F$11&" ","")
    return X.cat((X.cat(X.cat('xung ',E.v(23,11,4)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else ''),(X.cat(X.cat('xung ',E.v(23,11,6)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f7(E, r, c):
    # Ô F7
    # = IF(INDEX(M_XU,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$E$6+1)=1,"xung "&TRUNG_GIAN!$E$11&" ","")
    return (X.cat(X.cat('xung ',E.v(23,11,5)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else '')

def f8(E, r, c):
    # Ô C8
    # = IF(INDEX(M_HINH,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$D$6+1)=1,IF(TRUNG_GIAN!$C$6=TRUNG_GIAN!$D$6,"tự hình ","hình "&TRUNG_GIAN!$D$11&" "),"")&IF(INDEX(M_HINH
    #   ,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$E$6+1)=1,IF(TRUNG_GIAN!$C$6=TRUNG_GIAN!$E$6,"tự hình ","hình "&TRUNG_GIAN!$E$11&" "),"")&IF(INDEX(M_HINH,TRUNG_GIAN!$C$
    #   6+1,TRUNG_GIAN!$F$6+1)=1,IF(TRUNG_GIAN!$C$6=TRUNG_GIAN!$F$6,"tự hình ","hình "&TRUNG_GIAN!$F$11&" "),"")
    return X.cat(X.cat((('tự hình ' if X.truthy(X.eq(E.v(23,6,3),E.v(23,6,4))) else X.cat(X.cat('hình ',E.v(23,11,4)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else ''),(('tự hình ' if X.truthy(X.eq(E.v(23,6,3),E.v(23,6,5))) else X.cat(X.cat('hình ',E.v(23,11,5)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else '')),(('tự hình ' if X.truthy(X.eq(E.v(23,6,3),E.v(23,6,6))) else X.cat(X.cat('hình ',E.v(23,11,6)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f9(E, r, c):
    # Ô D8
    # = IF(INDEX(M_HINH,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$C$6+1)=1,IF(TRUNG_GIAN!$D$6=TRUNG_GIAN!$C$6,"tự hình ","hình "&TRUNG_GIAN!$C$11&" "),"")&IF(INDEX(M_HINH
    #   ,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$E$6+1)=1,IF(TRUNG_GIAN!$D$6=TRUNG_GIAN!$E$6,"tự hình ","hình "&TRUNG_GIAN!$E$11&" "),"")&IF(INDEX(M_HINH,TRUNG_GIAN!$D$
    #   6+1,TRUNG_GIAN!$F$6+1)=1,IF(TRUNG_GIAN!$D$6=TRUNG_GIAN!$F$6,"tự hình ","hình "&TRUNG_GIAN!$F$11&" "),"")
    return X.cat(X.cat((('tự hình ' if X.truthy(X.eq(E.v(23,6,4),E.v(23,6,3))) else X.cat(X.cat('hình ',E.v(23,11,3)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(('tự hình ' if X.truthy(X.eq(E.v(23,6,4),E.v(23,6,5))) else X.cat(X.cat('hình ',E.v(23,11,5)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else '')),(('tự hình ' if X.truthy(X.eq(E.v(23,6,4),E.v(23,6,6))) else X.cat(X.cat('hình ',E.v(23,11,6)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f10(E, r, c):
    # Ô E8
    # = IF(INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1,IF(TRUNG_GIAN!$E$6=TRUNG_GIAN!$C$6,"tự hình ","hình "&TRUNG_GIAN!$C$11&" "),"")&IF(INDEX(M_HINH
    #   ,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1,IF(TRUNG_GIAN!$E$6=TRUNG_GIAN!$D$6,"tự hình ","hình "&TRUNG_GIAN!$D$11&" "),"")&IF(INDEX(M_HINH,TRUNG_GIAN!$E$
    #   6+1,TRUNG_GIAN!$F$6+1)=1,IF(TRUNG_GIAN!$E$6=TRUNG_GIAN!$F$6,"tự hình ","hình "&TRUNG_GIAN!$F$11&" "),"")
    return X.cat(X.cat((('tự hình ' if X.truthy(X.eq(E.v(23,6,5),E.v(23,6,3))) else X.cat(X.cat('hình ',E.v(23,11,3)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(('tự hình ' if X.truthy(X.eq(E.v(23,6,5),E.v(23,6,4))) else X.cat(X.cat('hình ',E.v(23,11,4)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else '')),(('tự hình ' if X.truthy(X.eq(E.v(23,6,5),E.v(23,6,6))) else X.cat(X.cat('hình ',E.v(23,11,6)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f11(E, r, c):
    # Ô F8
    # = IF(INDEX(M_HINH,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$C$6+1)=1,IF(TRUNG_GIAN!$F$6=TRUNG_GIAN!$C$6,"tự hình ","hình "&TRUNG_GIAN!$C$11&" "),"")&IF(INDEX(M_HINH
    #   ,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$D$6+1)=1,IF(TRUNG_GIAN!$F$6=TRUNG_GIAN!$D$6,"tự hình ","hình "&TRUNG_GIAN!$D$11&" "),"")&IF(INDEX(M_HINH,TRUNG_GIAN!$F$
    #   6+1,TRUNG_GIAN!$E$6+1)=1,IF(TRUNG_GIAN!$F$6=TRUNG_GIAN!$E$6,"tự hình ","hình "&TRUNG_GIAN!$E$11&" "),"")
    return X.cat(X.cat((('tự hình ' if X.truthy(X.eq(E.v(23,6,6),E.v(23,6,3))) else X.cat(X.cat('hình ',E.v(23,11,3)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(('tự hình ' if X.truthy(X.eq(E.v(23,6,6),E.v(23,6,4))) else X.cat(X.cat('hình ',E.v(23,11,4)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else '')),(('tự hình ' if X.truthy(X.eq(E.v(23,6,6),E.v(23,6,5))) else X.cat(X.cat('hình ',E.v(23,11,5)),' ')) if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else ''))

def f12(E, r, c):
    # Ô C9
    # = IF(INDEX(M_HAI,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$D$6+1)=1,"hại "&TRUNG_GIAN!$D$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$E$6+1)=1,"hại "&TRU
    #   NG_GIAN!$E$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$F$6+1)=1,"hại "&TRUNG_GIAN!$F$11&" ","")
    return X.cat(X.cat((X.cat(X.cat('hại ',E.v(23,11,4)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else ''),(X.cat(X.cat('hại ',E.v(23,11,5)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else '')),(X.cat(X.cat('hại ',E.v(23,11,6)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f13(E, r, c):
    # Ô D9
    # = IF(INDEX(M_HAI,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$C$6+1)=1,"hại "&TRUNG_GIAN!$C$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$E$6+1)=1,"hại "&TRU
    #   NG_GIAN!$E$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$F$6+1)=1,"hại "&TRUNG_GIAN!$F$11&" ","")
    return X.cat(X.cat((X.cat(X.cat('hại ',E.v(23,11,3)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(X.cat(X.cat('hại ',E.v(23,11,5)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else '')),(X.cat(X.cat('hại ',E.v(23,11,6)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f14(E, r, c):
    # Ô E9
    # = IF(INDEX(M_HAI,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1,"hại "&TRUNG_GIAN!$C$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1,"hại "&TRU
    #   NG_GIAN!$D$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)=1,"hại "&TRUNG_GIAN!$F$11&" ","")
    return X.cat(X.cat((X.cat(X.cat('hại ',E.v(23,11,3)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(X.cat(X.cat('hại ',E.v(23,11,4)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else '')),(X.cat(X.cat('hại ',E.v(23,11,6)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else ''))

def f15(E, r, c):
    # Ô F9
    # = IF(INDEX(M_HAI,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$C$6+1)=1,"hại "&TRUNG_GIAN!$C$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$D$6+1)=1,"hại "&TRU
    #   NG_GIAN!$D$11&" ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$E$6+1)=1,"hại "&TRUNG_GIAN!$E$11&" ","")
    return X.cat(X.cat((X.cat(X.cat('hại ',E.v(23,11,3)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),(X.cat(X.cat('hại ',E.v(23,11,4)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else '')),(X.cat(X.cat('hại ',E.v(23,11,5)),' ') if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else ''))

def f16(E, r, c):
    # Ô C10
    # = IF(MOD(TRUNG_GIAN!$C$7+2,5)=TRUNG_GIAN!$C$9,"bị can "&TRUNG_GIAN!$C$10&" khắc","")
    return (X.cat(X.cat('bị can ',E.v(23,10,3)),' khắc') if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(23,9,3))) else '')

def f17(E, r, c):
    # Ô D10
    # = IF(MOD(TRUNG_GIAN!$D$7+2,5)=TRUNG_GIAN!$D$9,"bị can "&TRUNG_GIAN!$D$10&" khắc","")
    return (X.cat(X.cat('bị can ',E.v(23,10,4)),' khắc') if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(23,9,4))) else '')

def f18(E, r, c):
    # Ô E10
    # = IF(MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$E$9,"bị can "&TRUNG_GIAN!$E$10&" khắc","")
    return (X.cat(X.cat('bị can ',E.v(23,10,5)),' khắc') if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,9,5))) else '')

def f19(E, r, c):
    # Ô F10
    # = IF(MOD(TRUNG_GIAN!$F$7+2,5)=TRUNG_GIAN!$F$9,"bị can "&TRUNG_GIAN!$F$10&" khắc","")
    return (X.cat(X.cat('bị can ',E.v(23,10,6)),' khắc') if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(23,9,6))) else '')

def f20(E, r, c):
    # Ô C11, D11, E11, F11
    # = IF(LEN(C7&C8&C9&C10)>0,1,0)
    return (1.0 if X.truthy(X.gt(X.LEN(X.cat(X.cat(X.cat(E.v(12,r-4,c),E.v(12,r-3,c)),E.v(12,r-2,c)),E.v(12,r-1,c))),0.0)) else 0.0)

def f21(E, r, c):
    # Ô C12
    # = TRUNG_GIAN!$C$29
    return E.v(23,29,3)

def f22(E, r, c):
    # Ô D12
    # = TRUNG_GIAN!$D$29
    return E.v(23,29,4)

def f23(E, r, c):
    # Ô E12
    # = TRUNG_GIAN!$E$29
    return E.v(23,29,5)

def f24(E, r, c):
    # Ô F12
    # = TRUNG_GIAN!$F$29
    return E.v(23,29,6)

def f25(E, r, c):
    # Ô C13
    # = IF(OR(TRUNG_GIAN!$C$6=$C$17,TRUNG_GIAN!$C$6=$C$18),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,3),E.v(12,17,3)), X.eq(E.v(23,6,3),E.v(12,18,3)))) else 0.0)

def f26(E, r, c):
    # Ô D13
    # = IF(OR(TRUNG_GIAN!$D$6=$C$17,TRUNG_GIAN!$D$6=$C$18),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),E.v(12,17,3)), X.eq(E.v(23,6,4),E.v(12,18,3)))) else 0.0)

def f27(E, r, c):
    # Ô E13
    # = IF(OR(TRUNG_GIAN!$E$6=$C$17,TRUNG_GIAN!$E$6=$C$18),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,5),E.v(12,17,3)), X.eq(E.v(23,6,5),E.v(12,18,3)))) else 0.0)

def f28(E, r, c):
    # Ô F13
    # = IF(OR(TRUNG_GIAN!$F$6=$C$17,TRUNG_GIAN!$F$6=$C$18),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,6),E.v(12,17,3)), X.eq(E.v(23,6,6),E.v(12,18,3)))) else 0.0)

def f29(E, r, c):
    # Ô C16
    # = MOD(TRUNG_GIAN!$E$6-TRUNG_GIAN!$E$5,12)
    return X.MOD(X.sub(E.v(23,6,5),E.v(23,5,5)), 12.0)

def f30(E, r, c):
    # Ô C17
    # = MOD(C16+10,12)
    return X.MOD(X.add(E.v(12,r-1,c),10.0), 12.0)

def f31(E, r, c):
    # Ô D17, D18
    # = INDEX(CHI_TEN,C17+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,r,c-1),1.0))

def f32(E, r, c):
    # Ô C18
    # = MOD(C16+11,12)
    return X.MOD(X.add(E.v(12,r-2,c),11.0), 12.0)

def f33(E, r, c):
    # Ô C19
    # = MOD(TRUNG_GIAN!$D$5+1,10)
    return X.MOD(X.add(E.v(23,5,4),1.0), 10.0)

def f34(E, r, c):
    # Ô D19
    # = MOD(TRUNG_GIAN!$D$6+3,12)
    return X.MOD(X.add(E.v(23,6,4),3.0), 12.0)

def f35(E, r, c):
    # Ô E19
    # = INDEX(CAN_TEN,C19+1)&" "&INDEX(CHI_TEN,D19+1)
    return X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(12,r,c-2),1.0)),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,r,c-1),1.0)))

def f36(E, r, c):
    # Ô C20
    # = MOD(TRUNG_GIAN!$D$6-2,12)+1
    return X.add(X.MOD(X.sub(E.v(23,6,4),2.0), 12.0),1.0)

def f37(E, r, c):
    # Ô C21
    # = MOD(MOD(-(C20-1),12)+MOD(3-TRUNG_GIAN!$F$6,12),12)
    return X.MOD(X.add(X.MOD(X.neg(X.sub(E.v(12,r-1,c),1.0)), 12.0),X.MOD(X.sub(3.0,E.v(23,6,6)), 12.0)), 12.0)

def f38(E, r, c):
    # Ô C22
    # = MOD(CHOOSE(TRUNG_GIAN!$C$5+1,2,4,6,8,0,2,4,6,8,0)+MOD(C21-2,12),10)
    return X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: 2.0, lambda: 4.0, lambda: 6.0, lambda: 8.0, lambda: 0.0, lambda: 2.0, lambda: 4.0, lambda: 6.0, lambda: 8.0, lambda: 0.0),X.MOD(X.sub(E.v(12,r-1,c),2.0), 12.0)), 10.0)

def f39(E, r, c):
    # Ô C23
    # = INDEX(CAN_TEN,C22+1)&" "&INDEX(CHI_TEN,C21+1)
    return X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(12,r-1,c),1.0)),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,r-2,c),1.0)))

def f40(E, r, c):
    # Ô C27, C28, C29, C30
    # = INDEX(CHI_TEN,INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$E$5+1)+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,5),1.0)),1.0))

def f41(E, r, c):
    # Ô E27
    # = IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f42(E, r, c):
    # Ô F27, F31, F35, F39, F43, F47
    # = IF(E27=0,"",IF($C$11=1,"Như không có","Có tác dụng"))
    return ('' if X.truthy(X.eq(E.v(12,r,c-1),0.0)) else ('Như không có' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'Có tác dụng'))

def f43(E, r, c):
    # Ô G27, G31, G35, G39, G43, G47
    # = IF(E27=0,"",TRIM($C$7&$C$8&$C$9&$C$10))
    return ('' if X.truthy(X.eq(E.v(12,r,c-2),0.0)) else X.TRIM(X.cat(X.cat(X.cat(E.v(12,7,3),E.v(12,8,3)),E.v(12,9,3)),E.v(12,10,3))))

def f44(E, r, c):
    # Ô H27, H31, H35, H39, H43, H47
    # = IF(E27=0,"",TRIM(IF(OR(TRUNG_GIAN!$C$28=5,TRUNG_GIAN!$C$28=6,TRUNG_GIAN!$C$28=7,TRUNG_GIAN!$C$28=9),TRUNG_GIAN!$C$29&" so với nhật can ","")&IF($C$13=
    #   1,"không vong","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-3),0.0)) else X.TRIM(X.cat((X.cat(E.v(23,29,3),' so với nhật can ') if X.truthy(X.OR(X.eq(E.v(23,28,3),5.0), X.eq(E.v(23,28,3),6.0), X.eq(E.v(23,28,3),7.0), X.eq(E.v(23,28,3),9.0))) else ''),('không vong' if X.truthy(X.eq(E.v(12,13,3),1.0)) else ''))))

def f45(E, r, c):
    # Ô I27, I31, I35, I39, I43, I47
    # = IF(E27=0,"",IF(TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$157,"dụng thần",IF(TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$158,"hỷ thần","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-4),0.0)) else ('dụng thần' if X.truthy(X.eq(E.v(23,9,3),E.v(8,157,3))) else ('hỷ thần' if X.truthy(X.eq(E.v(23,9,3),E.v(8,158,3))) else '')))

def f46(E, r, c):
    # Ô J27, J31, J35, J39, J43, J47
    # = IF(E27=0,"",C27&" ở trụ Năm"&" – "&F27&IF(G27<>""," ("&G27&")","")&IF(H27<>"","; "&H27,"")&IF(I27<>"","; trùng "&I27,""))
    return ('' if X.truthy(X.eq(E.v(12,r,c-5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(12,r,c-7),' ở trụ Năm'),' – '),E.v(12,r,c-4)),(X.cat(X.cat(' (',E.v(12,r,c-3)),')') if X.truthy(X.ne(E.v(12,r,c-3),'')) else '')),(X.cat('; ',E.v(12,r,c-2)) if X.truthy(X.ne(E.v(12,r,c-2),'')) else '')),(X.cat('; trùng ',E.v(12,r,c-1)) if X.truthy(X.ne(E.v(12,r,c-1),'')) else '')))

def f47(E, r, c):
    # Ô E28
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f48(E, r, c):
    # Ô F28, F32, F36, F40, F44, F48
    # = IF(E28=0,"",IF($D$11=1,"Như không có","Có tác dụng"))
    return ('' if X.truthy(X.eq(E.v(12,r,c-1),0.0)) else ('Như không có' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'Có tác dụng'))

def f49(E, r, c):
    # Ô G28, G32, G36, G40, G44, G48
    # = IF(E28=0,"",TRIM($D$7&$D$8&$D$9&$D$10))
    return ('' if X.truthy(X.eq(E.v(12,r,c-2),0.0)) else X.TRIM(X.cat(X.cat(X.cat(E.v(12,7,4),E.v(12,8,4)),E.v(12,9,4)),E.v(12,10,4))))

def f50(E, r, c):
    # Ô H28, H32, H36, H40, H44, H48
    # = IF(E28=0,"",TRIM(IF(OR(TRUNG_GIAN!$D$28=5,TRUNG_GIAN!$D$28=6,TRUNG_GIAN!$D$28=7,TRUNG_GIAN!$D$28=9),TRUNG_GIAN!$D$29&" so với nhật can ","")&IF($D$13=
    #   1,"không vong","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-3),0.0)) else X.TRIM(X.cat((X.cat(E.v(23,29,4),' so với nhật can ') if X.truthy(X.OR(X.eq(E.v(23,28,4),5.0), X.eq(E.v(23,28,4),6.0), X.eq(E.v(23,28,4),7.0), X.eq(E.v(23,28,4),9.0))) else ''),('không vong' if X.truthy(X.eq(E.v(12,13,4),1.0)) else ''))))

def f51(E, r, c):
    # Ô I28, I32, I36, I40, I44, I48
    # = IF(E28=0,"",IF(TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$157,"dụng thần",IF(TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$158,"hỷ thần","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-4),0.0)) else ('dụng thần' if X.truthy(X.eq(E.v(23,9,4),E.v(8,157,3))) else ('hỷ thần' if X.truthy(X.eq(E.v(23,9,4),E.v(8,158,3))) else '')))

def f52(E, r, c):
    # Ô J28, J32, J36, J40, J44, J48
    # = IF(E28=0,"",C28&" ở trụ Tháng"&" – "&F28&IF(G28<>""," ("&G28&")","")&IF(H28<>"","; "&H28,"")&IF(I28<>"","; trùng "&I28,""))
    return ('' if X.truthy(X.eq(E.v(12,r,c-5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(12,r,c-7),' ở trụ Tháng'),' – '),E.v(12,r,c-4)),(X.cat(X.cat(' (',E.v(12,r,c-3)),')') if X.truthy(X.ne(E.v(12,r,c-3),'')) else '')),(X.cat('; ',E.v(12,r,c-2)) if X.truthy(X.ne(E.v(12,r,c-2),'')) else '')),(X.cat('; trùng ',E.v(12,r,c-1)) if X.truthy(X.ne(E.v(12,r,c-1),'')) else '')))

def f53(E, r, c):
    # Ô E29
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f54(E, r, c):
    # Ô F29, F33, F37, F41, F45, F49
    # = IF(E29=0,"",IF($E$11=1,"Như không có","Có tác dụng"))
    return ('' if X.truthy(X.eq(E.v(12,r,c-1),0.0)) else ('Như không có' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'Có tác dụng'))

def f55(E, r, c):
    # Ô G29, G33, G37, G41, G45, G49
    # = IF(E29=0,"",TRIM($E$7&$E$8&$E$9&$E$10))
    return ('' if X.truthy(X.eq(E.v(12,r,c-2),0.0)) else X.TRIM(X.cat(X.cat(X.cat(E.v(12,7,5),E.v(12,8,5)),E.v(12,9,5)),E.v(12,10,5))))

def f56(E, r, c):
    # Ô H29, H33, H37, H41, H45, H49
    # = IF(E29=0,"",TRIM(IF(OR(TRUNG_GIAN!$E$28=5,TRUNG_GIAN!$E$28=6,TRUNG_GIAN!$E$28=7,TRUNG_GIAN!$E$28=9),TRUNG_GIAN!$E$29&" so với nhật can ","")&IF($E$13=
    #   1,"không vong","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-3),0.0)) else X.TRIM(X.cat((X.cat(E.v(23,29,5),' so với nhật can ') if X.truthy(X.OR(X.eq(E.v(23,28,5),5.0), X.eq(E.v(23,28,5),6.0), X.eq(E.v(23,28,5),7.0), X.eq(E.v(23,28,5),9.0))) else ''),('không vong' if X.truthy(X.eq(E.v(12,13,5),1.0)) else ''))))

def f57(E, r, c):
    # Ô I29, I33, I37, I41, I45, I49
    # = IF(E29=0,"",IF(TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$157,"dụng thần",IF(TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$158,"hỷ thần","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-4),0.0)) else ('dụng thần' if X.truthy(X.eq(E.v(23,9,5),E.v(8,157,3))) else ('hỷ thần' if X.truthy(X.eq(E.v(23,9,5),E.v(8,158,3))) else '')))

def f58(E, r, c):
    # Ô J29, J33, J37, J41, J45, J49
    # = IF(E29=0,"",C29&" ở trụ Ngày"&" – "&F29&IF(G29<>""," ("&G29&")","")&IF(H29<>"","; "&H29,"")&IF(I29<>"","; trùng "&I29,""))
    return ('' if X.truthy(X.eq(E.v(12,r,c-5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(12,r,c-7),' ở trụ Ngày'),' – '),E.v(12,r,c-4)),(X.cat(X.cat(' (',E.v(12,r,c-3)),')') if X.truthy(X.ne(E.v(12,r,c-3),'')) else '')),(X.cat('; ',E.v(12,r,c-2)) if X.truthy(X.ne(E.v(12,r,c-2),'')) else '')),(X.cat('; trùng ',E.v(12,r,c-1)) if X.truthy(X.ne(E.v(12,r,c-1),'')) else '')))

def f59(E, r, c):
    # Ô E30
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f60(E, r, c):
    # Ô F30, F34, F38, F42, F46, F50
    # = IF(E30=0,"",IF($F$11=1,"Như không có","Có tác dụng"))
    return ('' if X.truthy(X.eq(E.v(12,r,c-1),0.0)) else ('Như không có' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'Có tác dụng'))

def f61(E, r, c):
    # Ô G30, G34, G38, G42, G46, G50
    # = IF(E30=0,"",TRIM($F$7&$F$8&$F$9&$F$10))
    return ('' if X.truthy(X.eq(E.v(12,r,c-2),0.0)) else X.TRIM(X.cat(X.cat(X.cat(E.v(12,7,6),E.v(12,8,6)),E.v(12,9,6)),E.v(12,10,6))))

def f62(E, r, c):
    # Ô H30, H34, H38, H42, H46, H50
    # = IF(E30=0,"",TRIM(IF(OR(TRUNG_GIAN!$F$28=5,TRUNG_GIAN!$F$28=6,TRUNG_GIAN!$F$28=7,TRUNG_GIAN!$F$28=9),TRUNG_GIAN!$F$29&" so với nhật can ","")&IF($F$13=
    #   1,"không vong","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-3),0.0)) else X.TRIM(X.cat((X.cat(E.v(23,29,6),' so với nhật can ') if X.truthy(X.OR(X.eq(E.v(23,28,6),5.0), X.eq(E.v(23,28,6),6.0), X.eq(E.v(23,28,6),7.0), X.eq(E.v(23,28,6),9.0))) else ''),('không vong' if X.truthy(X.eq(E.v(12,13,6),1.0)) else ''))))

def f63(E, r, c):
    # Ô I30, I34, I38, I42, I46, I50
    # = IF(E30=0,"",IF(TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$157,"dụng thần",IF(TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$158,"hỷ thần","")))
    return ('' if X.truthy(X.eq(E.v(12,r,c-4),0.0)) else ('dụng thần' if X.truthy(X.eq(E.v(23,9,6),E.v(8,157,3))) else ('hỷ thần' if X.truthy(X.eq(E.v(23,9,6),E.v(8,158,3))) else '')))

def f64(E, r, c):
    # Ô J30, J34, J38, J42, J46, J50
    # = IF(E30=0,"",C30&" ở trụ Giờ"&" – "&F30&IF(G30<>""," ("&G30&")","")&IF(H30<>"","; "&H30,"")&IF(I30<>"","; trùng "&I30,""))
    return ('' if X.truthy(X.eq(E.v(12,r,c-5),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(12,r,c-7),' ở trụ Giờ'),' – '),E.v(12,r,c-4)),(X.cat(X.cat(' (',E.v(12,r,c-3)),')') if X.truthy(X.ne(E.v(12,r,c-3),'')) else '')),(X.cat('; ',E.v(12,r,c-2)) if X.truthy(X.ne(E.v(12,r,c-2),'')) else '')),(X.cat('; trùng ',E.v(12,r,c-1)) if X.truthy(X.ne(E.v(12,r,c-1),'')) else '')))

def f65(E, r, c):
    # Ô C31, C32, C33, C34
    # = INDEX(CHI_TEN,INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1)+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)),1.0))

def f66(E, r, c):
    # Ô E31
    # = IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f67(E, r, c):
    # Ô E32
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f68(E, r, c):
    # Ô E33
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f69(E, r, c):
    # Ô E34
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f70(E, r, c):
    # Ô C35, C36, C37, C38
    # = INDEX(CHI_TEN,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1)+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)),1.0))

def f71(E, r, c):
    # Ô E35
    # = IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f72(E, r, c):
    # Ô E36
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f73(E, r, c):
    # Ô E37
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f74(E, r, c):
    # Ô E38
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f75(E, r, c):
    # Ô C39, C40, C41, C42
    # = INDEX(CHI_TEN,INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$C$5+1)+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,3),1.0)),1.0))

def f76(E, r, c):
    # Ô E39
    # = IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f77(E, r, c):
    # Ô E40
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f78(E, r, c):
    # Ô E41
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f79(E, r, c):
    # Ô E42
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$I$5:$I$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,9,14,9), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f80(E, r, c):
    # Ô C43, C44, C45, C46
    # = INDEX(CHI_TEN,INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1)+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)),1.0))

def f81(E, r, c):
    # Ô E43
    # = IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f82(E, r, c):
    # Ô E44
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f83(E, r, c):
    # Ô E45
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f84(E, r, c):
    # Ô E46
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f85(E, r, c):
    # Ô C47, C48, C49, C50
    # = INDEX(CHI_TEN,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)),1.0))

def f86(E, r, c):
    # Ô E47
    # = IF(TRUNG_GIAN!$C$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,3),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f87(E, r, c):
    # Ô E48
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f88(E, r, c):
    # Ô E49
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f89(E, r, c):
    # Ô E50
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)))) else 0.0)

def f90(E, r, c):
    # Ô C53
    # = E27+E28+E29+E30+E39+E40+E41+E42
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(E.v(12,r-26,c+2),E.v(12,r-25,c+2)),E.v(12,r-24,c+2)),E.v(12,r-23,c+2)),E.v(12,r-14,c+2)),E.v(12,r-13,c+2)),E.v(12,r-12,c+2)),E.v(12,r-11,c+2))

def f91(E, r, c):
    # Ô C54
    # = E31+E32+E33+E34+E35+E36+E37+E38+E43+E44+E45+E46+E47+E48+E49+E50
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(E.v(12,r-23,c+2),E.v(12,r-22,c+2)),E.v(12,r-21,c+2)),E.v(12,r-20,c+2)),E.v(12,r-19,c+2)),E.v(12,r-18,c+2)),E.v(12,r-17,c+2)),E.v(12,r-16,c+2)),E.v(12,r-11,c+2)),E.v(12,r-10,c+2)),E.v(12,r-9,c+2)),E.v(12,r-8,c+2)),E.v(12,r-7,c+2)),E.v(12,r-6,c+2)),E.v(12,r-5,c+2)),E.v(12,r-4,c+2))

def f92(E, r, c):
    # Ô C55
    # = IF(C53=0,"Không có Văn Xương trong tứ trụ",TRIM(IF(E27=1,B27&": "&J27&". ","")&IF(E28=1,B28&": "&J28&". ","")&IF(E29=1,B29&": "&J29&". ","")&IF(E30=1,
    #   B30&": "&J30&". ","")&IF(E39=1,B39&": "&J39&". ","")&IF(E40=1,B40&": "&J40&". ","")&IF(E41=1,B41&": "&J41&". ","")&IF(E42=1,B42&": "&J42&". ","")))
    return ('Không có Văn Xương trong tứ trụ' if X.truthy(X.eq(E.v(12,r-2,c),0.0)) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(12,r-28,c-1),': '),E.v(12,r-28,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-28,c+2),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(12,r-27,c-1),': '),E.v(12,r-27,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-27,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-26,c-1),': '),E.v(12,r-26,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-26,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-25,c-1),': '),E.v(12,r-25,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-25,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-16,c-1),': '),E.v(12,r-16,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-16,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-15,c-1),': '),E.v(12,r-15,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-15,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-14,c-1),': '),E.v(12,r-14,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-14,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-13,c-1),': '),E.v(12,r-13,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-13,c+2),1.0)) else ''))))

def f93(E, r, c):
    # Ô C56
    # = IF(C54=0,"Không có quý nhân trong tứ trụ",TRIM(IF(E31=1,B31&": "&J31&". ","")&IF(E32=1,B32&": "&J32&". ","")&IF(E33=1,B33&": "&J33&". ","")&IF(E34=1,B
    #   34&": "&J34&". ","")&IF(E35=1,B35&": "&J35&". ","")&IF(E36=1,B36&": "&J36&". ","")&IF(E37=1,B37&": "&J37&". ","")&IF(E38=1,B38&": "&J38&". ","")&IF(E4
    #   3=1,B43&": "&J43&". ","")&IF(E44=1,B44&": "&J44&". ","")&IF(E45=1,B45&": "&J45&". ","")&IF(E46=1,B46&": "&J46&". ","")&IF(E47=1,B47&": "&J47&". ","")&
    #   IF(E48=1,B48&": "&J48&". ","")&IF(E49=1,B49&": "&J49&". ","")&IF(E50=1,B50&": "&J50&". ","")))
    return ('Không có quý nhân trong tứ trụ' if X.truthy(X.eq(E.v(12,r-2,c),0.0)) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(12,r-25,c-1),': '),E.v(12,r-25,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-25,c+2),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(12,r-24,c-1),': '),E.v(12,r-24,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-24,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-23,c-1),': '),E.v(12,r-23,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-23,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-22,c-1),': '),E.v(12,r-22,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-22,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-21,c-1),': '),E.v(12,r-21,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-21,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-20,c-1),': '),E.v(12,r-20,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-20,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-19,c-1),': '),E.v(12,r-19,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-19,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-18,c-1),': '),E.v(12,r-18,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-18,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-13,c-1),': '),E.v(12,r-13,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-13,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-12,c-1),': '),E.v(12,r-12,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-12,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-11,c-1),': '),E.v(12,r-11,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-11,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-10,c-1),': '),E.v(12,r-10,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-10,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-9,c-1),': '),E.v(12,r-9,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-9,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-8,c-1),': '),E.v(12,r-8,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-8,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-7,c-1),': '),E.v(12,r-7,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-7,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(12,r-6,c-1),': '),E.v(12,r-6,c+7)),'. ') if X.truthy(X.eq(E.v(12,r-6,c+2),1.0)) else ''))))

def f94(E, r, c):
    # Ô C57
    # = IF(C54>0,"(không cần)",IF(LEN(IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1)=$D$19,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)=$D$19),"quý nhân theo can 
    #   năm ở thai nguyên ("&INDEX(CHI_TEN,$D$19+1)&"); ","")&IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1)=$C$21,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)=$C
    #   $21),"quý nhân theo can năm ở cung mệnh ("&INDEX(CHI_TEN,$C$21+1)&"); ","")&IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1)=$D$19,INDEX(CSDL!$K$5:$K$14
    #   ,TRUNG_GIAN!$E$5+1)=$D$19),"quý nhân theo can ngày ở thai nguyên ("&INDEX(CHI_TEN,$D$19+1)&"); ","")&IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1)=$C
    #   $21,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1)=$C$21),"quý nhân theo can ngày ở cung mệnh ("&INDEX(CHI_TEN,$C$21+1)&"); ",""))=0,"Thai nguyên, cung mệnh
    #    cũng không có quý nhân",TRIM(IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1)=$D$19,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)=$D$19),"quý nhân theo can 
    #   năm ở thai nguyên ("&INDEX(CHI_TEN,$D$19+1)&"); ","")&IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1)=$C$21,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)=$C
    #   $21),"quý nhân theo can năm ở cung mệnh ("&INDEX(CHI_TEN,$C$21+1)&"); ","")&IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1)=$D$19,INDEX(CSDL!$K$5:$K$14
    #   ,TRUNG_GIAN!$E$5+1)=$D$19),"quý nhân theo can ngày ở thai nguyên ("&INDEX(CHI_TEN,$D$19+1)&"); ","")&IF(OR(INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1)=$C
    #   $21,INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1)=$C$21),"quý nhân theo can ngày ở cung mệnh ("&INDEX(CHI_TEN,$C$21+1)&"); ",""))))
    return ('(không cần)' if X.truthy(X.gt(E.v(12,r-3,c),0.0)) else ('Thai nguyên, cung mệnh cũng không có quý nhân' if X.truthy(X.eq(X.LEN(X.cat(X.cat(X.cat((X.cat(X.cat('quý nhân theo can năm ở thai nguyên (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,19,4),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)),E.v(12,19,4)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)),E.v(12,19,4)))) else ''),(X.cat(X.cat('quý nhân theo can năm ở cung mệnh (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,21,3),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)),E.v(12,21,3)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)),E.v(12,21,3)))) else '')),(X.cat(X.cat('quý nhân theo can ngày ở thai nguyên (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,19,4),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)),E.v(12,19,4)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)),E.v(12,19,4)))) else '')),(X.cat(X.cat('quý nhân theo can ngày ở cung mệnh (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,21,3),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)),E.v(12,21,3)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)),E.v(12,21,3)))) else ''))),0.0)) else X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('quý nhân theo can năm ở thai nguyên (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,19,4),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)),E.v(12,19,4)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)),E.v(12,19,4)))) else ''),(X.cat(X.cat('quý nhân theo can năm ở cung mệnh (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,21,3),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0)),E.v(12,21,3)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0)),E.v(12,21,3)))) else '')),(X.cat(X.cat('quý nhân theo can ngày ở thai nguyên (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,19,4),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)),E.v(12,19,4)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)),E.v(12,19,4)))) else '')),(X.cat(X.cat('quý nhân theo can ngày ở cung mệnh (',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(12,21,3),1.0))),'); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0)),E.v(12,21,3)), X.eq(X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0)),E.v(12,21,3)))) else '')))))

def f95(E, r, c):
    # Ô C85
    # = IF(OR(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRUNG_GIAN!$C$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),AND(INDEX(CSDL!$B$122:$B$133,
    #   TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$C$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1))),IF(IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRU
    #   NG_GIAN!$C$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),OR(SAO!$L$5=1,PP1_DINH_TINH!$C$29=1),THAN_SAT!$C$11=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy((X.OR(X.eq(E.v(13,5,12),1.0), X.eq(E.v(9,29,3),1.0)) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,3),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else X.eq(E.v(12,11,3),1.0))) else 'có') if X.truthy(X.OR(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,3),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,3),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))))) else '')

def f96(E, r, c):
    # Ô D85
    # = IF(OR(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRUNG_GIAN!$D$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),AND(INDEX(CSDL!$B$122:$B$133,
    #   TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$D$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1))),IF(IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRU
    #   NG_GIAN!$D$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),OR(SAO!$L$9=1,PP1_DINH_TINH!$D$29=1),THAN_SAT!$D$11=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy((X.OR(X.eq(E.v(13,9,12),1.0), X.eq(E.v(9,29,4),1.0)) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,4),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else X.eq(E.v(12,11,4),1.0))) else 'có') if X.truthy(X.OR(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,4),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))))) else '')

def f97(E, r, c):
    # Ô E85
    # = IF(OR(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRUNG_GIAN!$E$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),AND(INDEX(CSDL!$B$122:$B$133,
    #   TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$E$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1))),IF(IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRU
    #   NG_GIAN!$E$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),OR(SAO!$L$13=1,PP1_DINH_TINH!$E$29=1),THAN_SAT!$E$11=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy((X.OR(X.eq(E.v(13,13,12),1.0), X.eq(E.v(9,29,5),1.0)) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,5),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else X.eq(E.v(12,11,5),1.0))) else 'có') if X.truthy(X.OR(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,5),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))))) else '')

def f98(E, r, c):
    # Ô F85
    # = IF(OR(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRUNG_GIAN!$F$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),AND(INDEX(CSDL!$B$122:$B$133,
    #   TRUNG_GIAN!$D$6+1)="chi",TRUNG_GIAN!$F$6=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1))),IF(IF(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",TRU
    #   NG_GIAN!$F$5=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),OR(SAO!$L$17=1,PP1_DINH_TINH!$F$29=1),THAN_SAT!$F$11=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy((X.OR(X.eq(E.v(13,17,12),1.0), X.eq(E.v(9,29,6),1.0)) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,6),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0))))) else X.eq(E.v(12,11,6),1.0))) else 'có') if X.truthy(X.OR(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(23,5,6),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))))) else '')

def f99(E, r, c):
    # Ô H85, H86, H87, H88, H89, H90, H91, H92, H93, H96, H102, H103 … (40 ô)
    # = IF(COUNTIF(C85:F85,"có")>0,"Có",IF(COUNTIF(C85:F85,"có – vô hiệu")>0,"Như không có","Không có"))
    return ('Có' if X.truthy(X.gt(X.COUNTIF(E.rg(12,r,c-5,r,c-2), 'có'),0.0)) else ('Như không có' if X.truthy(X.gt(X.COUNTIF(E.rg(12,r,c-5,r,c-2), 'có – vô hiệu'),0.0)) else 'Không có'))

def f100(E, r, c):
    # Ô I85
    # = TRIM(IF(COUNTIF($C$85:$F$85,"có")>0,"• Nhân từ đôn hậu, phúc nhiều, gặp hung hóa cát. ","")&IF(AND(COUNTIF($C$85:$F$85,"có")>0,COUNTIF($C$86:$F$86,"có
    #   ")>0),"• Có cả hai đức: năng lực hóa hung rất mạnh, ít bệnh, không bị hình phạt. ","")&IF($E$85="có","• Đức ở trụ ngày: cuộc đời không gặp nguy hiểm. 
    #   ","")&IF(AND(COUNTIF($C$85:$F$85,"có")>0,NHAP_LIEU!$C$22="Nữ"),"• Nữ: lấy chồng tốt, dễ sinh đẻ. ","")&IF(AND(NOT(COUNTIF($C$85:$F$85,"có")>0),(COUNTI
    #   F($C$85:$F$85,"có*")>0)),"• Bị xung khắc → vô dụng. ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(('• Nhân từ đôn hậu, phúc nhiều, gặp hung hóa cát. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có'),0.0)) else ''),('• Có cả hai đức: năng lực hóa hung rất mạnh, ít bệnh, không bị hình phạt. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có'),0.0), X.gt(X.COUNTIF(E.rg(12,86,3,86,6), 'có'),0.0))) else '')),('• Đức ở trụ ngày: cuộc đời không gặp nguy hiểm. ' if X.truthy(X.eq(E.v(12,85,5),'có')) else '')),('• Nữ: lấy chồng tốt, dễ sinh đẻ. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có'),0.0), X.eq(E.v(2,22,3),'Nữ'))) else '')),('• Bị xung khắc → vô dụng. ' if X.truthy(X.AND(X.NOT(X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có'),0.0)), X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có*'),0.0))) else '')))

def f101(E, r, c):
    # Ô C86
    # = IF(TRUNG_GIAN!$C$5=INDEX(CSDL!$D$122:$D$133,TRUNG_GIAN!$D$6+1),IF(OR(SAO!$L$5=1,PP1_DINH_TINH!$C$29=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.eq(E.v(13,5,12),1.0), X.eq(E.v(9,29,3),1.0))) else 'có') if X.truthy(X.eq(E.v(23,5,3),X.INDEX(E.rg(24,122,4,133,4), X.add(E.v(23,6,4),1.0)))) else '')

def f102(E, r, c):
    # Ô D86
    # = IF(TRUNG_GIAN!$D$5=INDEX(CSDL!$D$122:$D$133,TRUNG_GIAN!$D$6+1),IF(OR(SAO!$L$9=1,PP1_DINH_TINH!$D$29=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.eq(E.v(13,9,12),1.0), X.eq(E.v(9,29,4),1.0))) else 'có') if X.truthy(X.eq(E.v(23,5,4),X.INDEX(E.rg(24,122,4,133,4), X.add(E.v(23,6,4),1.0)))) else '')

def f103(E, r, c):
    # Ô E86
    # = IF(TRUNG_GIAN!$E$5=INDEX(CSDL!$D$122:$D$133,TRUNG_GIAN!$D$6+1),IF(OR(SAO!$L$13=1,PP1_DINH_TINH!$E$29=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.eq(E.v(13,13,12),1.0), X.eq(E.v(9,29,5),1.0))) else 'có') if X.truthy(X.eq(E.v(23,5,5),X.INDEX(E.rg(24,122,4,133,4), X.add(E.v(23,6,4),1.0)))) else '')

def f104(E, r, c):
    # Ô F86
    # = IF(TRUNG_GIAN!$F$5=INDEX(CSDL!$D$122:$D$133,TRUNG_GIAN!$D$6+1),IF(OR(SAO!$L$17=1,PP1_DINH_TINH!$F$29=1),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.eq(E.v(13,17,12),1.0), X.eq(E.v(9,29,6),1.0))) else 'có') if X.truthy(X.eq(E.v(23,5,6),X.INDEX(E.rg(24,122,4,133,4), X.add(E.v(23,6,4),1.0)))) else '')

def f105(E, r, c):
    # Ô I86
    # = TRIM(IF(COUNTIF($C$86:$F$86,"có")>0,"• Cát tinh phúc thọ; hóa hung thành cát. ","")&IF($E$86="có","• Đức ở trụ ngày: cuộc đời không gặp nguy hiểm. ","
    #   ")&IF(AND(NOT(COUNTIF($C$86:$F$86,"có")>0),(COUNTIF($C$86:$F$86,"có*")>0)),"• Bị xung khắc → vô dụng. ",""))
    return X.TRIM(X.cat(X.cat(('• Cát tinh phúc thọ; hóa hung thành cát. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,86,3,86,6), 'có'),0.0)) else ''),('• Đức ở trụ ngày: cuộc đời không gặp nguy hiểm. ' if X.truthy(X.eq(E.v(12,86,5),'có')) else '')),('• Bị xung khắc → vô dụng. ' if X.truthy(X.AND(X.NOT(X.gt(X.COUNTIF(E.rg(12,86,3,86,6), 'có'),0.0)), X.gt(X.COUNTIF(E.rg(12,86,3,86,6), 'có*'),0.0))) else '')))

def f106(E, r, c):
    # Ô C87
    # = IF(SAO!$Q$5=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,5,17),1.0)) else '')

def f107(E, r, c):
    # Ô D87
    # = IF(SAO!$Q$9=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,9,17),1.0)) else '')

def f108(E, r, c):
    # Ô E87
    # = IF(SAO!$Q$13=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,13,17),1.0)) else '')

def f109(E, r, c):
    # Ô F87
    # = IF(SAO!$Q$17=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,17,17),1.0)) else '')

def f110(E, r, c):
    # Ô I87
    # = TRIM(IF(COUNTIF($C$87:$F$87,"có")>0,"• Công danh thuận lợi (không bị hình xung khắc phá). ","")&IF(AND(COUNTIF($C$87:$F$87,"có")>0,DUNG_HY_KY!$C$155=0
    #   ),"• Thân nhược mà có lộc không bị phá: quý mệnh. ","")&IF(AND((COUNTIF($C$87:$F$87,"có*")>0),DUNG_HY_KY!$C$155=1),"• Thân vượng lộc nhiều: nên có qua
    #   n/thực thương khắc tiết. ","")&IF(AND(NOT(COUNTIF($C$87:$F$87,"có")>0),(COUNTIF($C$87:$F$87,"có*")>0)),"• Lộc bị xung phá: quan thì mất chức, người th
    #   ường cơm áo không đủ, bôn ba. ","")&IF(AND((COUNTIF($C$87:$F$87,"có*")>0),COUNTIFS(SAO!$Q$5:$Q$20,1,SAO!$N$5:$N$20,1)>0),"• Lộc nhập không vong: có lộ
    #   c như không. ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(('• Công danh thuận lợi (không bị hình xung khắc phá). ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,87,3,87,6), 'có'),0.0)) else ''),('• Thân nhược mà có lộc không bị phá: quý mệnh. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,87,3,87,6), 'có'),0.0), X.eq(E.v(8,155,3),0.0))) else '')),('• Thân vượng lộc nhiều: nên có quan/thực thương khắc tiết. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,87,3,87,6), 'có*'),0.0), X.eq(E.v(8,155,3),1.0))) else '')),('• Lộc bị xung phá: quan thì mất chức, người thường cơm áo không đủ, bôn ba. ' if X.truthy(X.AND(X.NOT(X.gt(X.COUNTIF(E.rg(12,87,3,87,6), 'có'),0.0)), X.gt(X.COUNTIF(E.rg(12,87,3,87,6), 'có*'),0.0))) else '')),('• Lộc nhập không vong: có lộc như không. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,87,3,87,6), 'có*'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,17,20,17), 1.0, E.rg(13,5,14,20,14), 1.0),0.0))) else '')))

def f111(E, r, c):
    # Ô C88
    # = IF(SAO!$R$5=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,5,18),1.0)) else '')

def f112(E, r, c):
    # Ô D88
    # = IF(SAO!$R$9=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,9,18),1.0)) else '')

def f113(E, r, c):
    # Ô E88
    # = IF(SAO!$R$13=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,13,18),1.0)) else '')

def f114(E, r, c):
    # Ô F88
    # = IF(SAO!$R$17=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,17,18),1.0)) else '')

def f115(E, r, c):
    # Ô I88
    # = TRIM(IF((COUNTIF($C$88:$F$88,"có*")>0),"• Hung nhiều cát ít; dễ tai nạn khi bị hợp, xung, hình hoặc gặp tuế vận. ","")&IF(AND((COUNTIF($C$88:$F$88,"có
    #   *")>0),DUNG_HY_KY!$C$155=0),"• Thân nhược: kình dương giúp thân (nhưng ""như làm bạn với hổ""). ","")&IF(AND((COUNTIF($C$88:$F$88,"có*")>0),DUNG_HY_KY
    #   !$C$155=1),"• Thân vượng: cái xấu của kình dương càng tăng. ","")&IF(LEN($C$88)>0,"• Ở trụ năm: phá cái tổ tiên để lại. ","")&IF(LEN($D$88)>0,"• Ở trụ
    #    tháng: tính ngang ngược, khác người. ","")&IF(LEN($E$88)>0,"• Ở trụ ngày: tính nóng; vợ hay bệnh. ","")&IF(LEN($F$88)>0,"• Ở trụ giờ: khắc vợ con, cu
    #   ối đời hay gặp họa. ","")&IF(AND((COUNTIF($C$88:$F$88,"có*")>0),COUNTIF(SAO!$F$5:$F$20,"Thất sát")>0),"• Có cả thất sát: ""có sát có kình dương thì lậ
    #   p công kiến nghiệp"". ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('• Hung nhiều cát ít; dễ tai nạn khi bị hợp, xung, hình hoặc gặp tuế vận. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,88,3,88,6), 'có*'),0.0)) else ''),('• Thân nhược: kình dương giúp thân (nhưng "như làm bạn với hổ"). ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,88,3,88,6), 'có*'),0.0), X.eq(E.v(8,155,3),0.0))) else '')),('• Thân vượng: cái xấu của kình dương càng tăng. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,88,3,88,6), 'có*'),0.0), X.eq(E.v(8,155,3),1.0))) else '')),('• Ở trụ năm: phá cái tổ tiên để lại. ' if X.truthy(X.gt(X.LEN(E.v(12,88,3)),0.0)) else '')),('• Ở trụ tháng: tính ngang ngược, khác người. ' if X.truthy(X.gt(X.LEN(E.v(12,88,4)),0.0)) else '')),('• Ở trụ ngày: tính nóng; vợ hay bệnh. ' if X.truthy(X.gt(X.LEN(E.v(12,88,5)),0.0)) else '')),('• Ở trụ giờ: khắc vợ con, cuối đời hay gặp họa. ' if X.truthy(X.gt(X.LEN(E.v(12,88,6)),0.0)) else '')),('• Có cả thất sát: "có sát có kình dương thì lập công kiến nghiệp". ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,88,3,88,6), 'có*'),0.0), X.gt(X.COUNTIF(E.rg(13,5,6,20,6), 'Thất sát'),0.0))) else '')))

def f116(E, r, c):
    # Ô C89
    # = IF(SAO!$T$5=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,5,20),1.0)) else '')

def f117(E, r, c):
    # Ô D89
    # = IF(SAO!$T$9=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,9,20),1.0)) else '')

def f118(E, r, c):
    # Ô E89
    # = IF(SAO!$T$13=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,13,20),1.0)) else '')

def f119(E, r, c):
    # Ô F89
    # = IF(SAO!$T$17=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,17,20),1.0)) else '')

def f120(E, r, c):
    # Ô I89
    # = TRIM(IF((COUNTIF($C$89:$F$89,"có*")>0),"• Thông minh, giỏi nghệ thuật/kỹ thuật, có duyên khí công, tôn giáo; dễ cô quả. ","")&IF(AND((COUNTIF($C$89:$F
    #   $89,"có*")>0),COUNTIFS(SAO!$T$5:$T$20,1,SAO!$F$5:$F$20,"*ấn")>0,COUNTIFS(SAO!$T$5:$T$20,1,SAO!$J$5:$J$20,1)>0),"• Hoa cái gặp ấn và vượng tướng: thi c
    #   ử đỗ cao. ","")&IF(AND((COUNTIF($C$89:$F$89,"có*")>0),OR(NOT(COUNTIF($C$89:$F$89,"có")>0),COUNTIFS(SAO!$T$5:$T$20,1,SAO!$N$5:$N$20,1)>0)),"• Hoa cái b
    #   ị phá hại hoặc tuần không: dễ đi tu, cô quả, phiêu bạt. ",""))
    return X.TRIM(X.cat(X.cat(('• Thông minh, giỏi nghệ thuật/kỹ thuật, có duyên khí công, tôn giáo; dễ cô quả. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,89,3,89,6), 'có*'),0.0)) else ''),('• Hoa cái gặp ấn và vượng tướng: thi cử đỗ cao. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,89,3,89,6), 'có*'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,20,20,20), 1.0, E.rg(13,5,6,20,6), '*ấn'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,20,20,20), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '')),('• Hoa cái bị phá hại hoặc tuần không: dễ đi tu, cô quả, phiêu bạt. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,89,3,89,6), 'có*'),0.0), X.OR(X.NOT(X.gt(X.COUNTIF(E.rg(12,89,3,89,6), 'có'),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,20,20,20), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)))) else '')))

def f121(E, r, c):
    # Ô C90
    # = IF(SAO!$W$5=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,5,23),1.0)) else '')

def f122(E, r, c):
    # Ô D90
    # = IF(SAO!$W$9=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,9,23),1.0)) else '')

def f123(E, r, c):
    # Ô E90
    # = IF(SAO!$W$13=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,13,23),1.0)) else '')

def f124(E, r, c):
    # Ô F90
    # = IF(SAO!$W$17=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,17,23),1.0)) else '')

def f125(E, r, c):
    # Ô I90
    # = TRIM(IF((COUNTIF($C$90:$F$90,"có*")>0),"• Hung tinh: đề phòng bệnh tật, thương tật, hình pháp (""không gặp là tốt nhất""). ","")&IF(AND((COUNTIF($C$90
    #   :$F$90,"có*")>0),COUNTIFS(SAO!$W$5:$W$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$H$5:$H$20,DUNG_HY_KY!$C$159)>0),"• Kiếp sát là kỵ thần: tính cường bạo, đến t
    #   uế vận cướp đoạt rất hung. ","")&IF(AND((COUNTIF($C$90:$F$90,"có*")>0),OR(COUNTIFS(SAO!$W$5:$W$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$H$5:$H$20,DUNG_HY_KY
    #   !$C$157)>0,COUNTIFS(SAO!$W$5:$W$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$H$5:$H$20,DUNG_HY_KY!$C$158)>0)),"• Kiếp sát là dụng/hỷ thần: hiếu học, quả đoán, d
    #   ễ thành công. ",""))
    return X.TRIM(X.cat(X.cat(('• Hung tinh: đề phòng bệnh tật, thương tật, hình pháp ("không gặp là tốt nhất"). ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,90,3,90,6), 'có*'),0.0)) else ''),('• Kiếp sát là kỵ thần: tính cường bạo, đến tuế vận cướp đoạt rất hung. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,90,3,90,6), 'có*'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,23,20,23), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,8,20,8), E.v(8,159,3)),0.0))) else '')),('• Kiếp sát là dụng/hỷ thần: hiếu học, quả đoán, dễ thành công. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,90,3,90,6), 'có*'),0.0), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,23,20,23), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,8,20,8), E.v(8,157,3)),0.0), X.gt(X.COUNTIFS(E.rg(13,5,23,20,23), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,8,20,8), E.v(8,158,3)),0.0)))) else '')))

def f126(E, r, c):
    # Ô C91
    # = IF(SAO!$S$5=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,5,19),1.0)) else '')

def f127(E, r, c):
    # Ô D91
    # = IF(SAO!$S$9=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,9,19),1.0)) else '')

def f128(E, r, c):
    # Ô E91
    # = IF(SAO!$S$13=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,13,19),1.0)) else '')

def f129(E, r, c):
    # Ô F91
    # = IF(SAO!$S$17=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,17,19),1.0)) else '')

def f130(E, r, c):
    # Ô I91
    # = TRIM(IF((COUNTIF($C$91:$F$91,"có*")>0),"• Người đẹp, thông minh, nhiều tài nghệ; tổ hợp xấu thì phong lưu, quan hệ nam nữ mờ ám. ","")&IF(AND((COUNTIF
    #   ($C$91:$F$91,"có*")>0),COUNTIFS(SAO!$S$5:$S$20,1,SAO!$F$5:$F$20,"*ấn")>0),"• Đào hoa gặp ấn: văn chương. ","")&IF(AND((COUNTIF($C$91:$F$91,"có*")>0),C
    #   OUNTIFS(SAO!$S$5:$S$20,1,SAO!$F$5:$F$20,"Chính quan")+COUNTIFS(SAO!$S$5:$S$20,1,SAO!$F$5:$F$20,"Thất sát")>0),"• Đào hoa gặp quan: được thăng chức. ",
    #   "")&IF(AND((COUNTIF($C$91:$F$91,"có*")>0),COUNTIFS(SAO!$S$5:$S$20,1,SAO!$F$5:$F$20,"*tài")>0),"• Đào hoa gặp tài: quan quý. ","")&IF(LEN($D$91)>0,"• T
    #   háng sinh gặp hàm trì: vợ chồng dễ ngoại tình (P2 tr.471). ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(('• Người đẹp, thông minh, nhiều tài nghệ; tổ hợp xấu thì phong lưu, quan hệ nam nữ mờ ám. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,91,3,91,6), 'có*'),0.0)) else ''),('• Đào hoa gặp ấn: văn chương. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,91,3,91,6), 'có*'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,19,20,19), 1.0, E.rg(13,5,6,20,6), '*ấn'),0.0))) else '')),('• Đào hoa gặp quan: được thăng chức. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,91,3,91,6), 'có*'),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,19,20,19), 1.0, E.rg(13,5,6,20,6), 'Chính quan'),X.COUNTIFS(E.rg(13,5,19,20,19), 1.0, E.rg(13,5,6,20,6), 'Thất sát')),0.0))) else '')),('• Đào hoa gặp tài: quan quý. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,91,3,91,6), 'có*'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,19,20,19), 1.0, E.rg(13,5,6,20,6), '*tài'),0.0))) else '')),('• Tháng sinh gặp hàm trì: vợ chồng dễ ngoại tình (P2 tr.471). ' if X.truthy(X.gt(X.LEN(E.v(12,91,4)),0.0)) else '')))

def f131(E, r, c):
    # Ô C92
    # = IF(SAO!$U$5=1,IF(OR(INDEX(M_LH,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$D$6+1)>=0),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,4),1.0)),0.0))) else 'có') if X.truthy(X.eq(E.v(13,5,21),1.0)) else '')

def f132(E, r, c):
    # Ô D92
    # = IF(SAO!$U$9=1,IF(OR(INDEX(M_LH,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$C$6+1)>=0,INDEX(M_LH,TRUNG_GIAN!$D$6+1,TRUNG_GIAN!$E$6+1)>=0),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,3),1.0)),0.0), X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(23,6,5),1.0)),0.0))) else 'có') if X.truthy(X.eq(E.v(13,9,21),1.0)) else '')

def f133(E, r, c):
    # Ô E92
    # = IF(SAO!$U$13=1,IF(OR(INDEX(M_LH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)>=0,INDEX(M_LH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)>=0),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),0.0), X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),0.0))) else 'có') if X.truthy(X.eq(E.v(13,13,21),1.0)) else '')

def f134(E, r, c):
    # Ô F92
    # = IF(SAO!$U$17=1,IF(OR(INDEX(M_LH,TRUNG_GIAN!$F$6+1,TRUNG_GIAN!$E$6+1)>=0),"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.OR(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(23,6,5),1.0)),0.0))) else 'có') if X.truthy(X.eq(E.v(13,17,21),1.0)) else '')

def f135(E, r, c):
    # Ô I92
    # = TRIM(IF((COUNTIF($C$92:$F$92,"có")>0),"• Mã là động, là tượng bôn ba: hiếu động, đi khắp bốn phương (quân nhân, thương nhân, ngoại giao, hay đi công t
    #   ác); cũng là tiêu chí đi nước ngoài. ","")&IF(AND((COUNTIF($C$92:$F$92,"có")>0),COUNTIFS(SAO!$U$5:$U$20,1,SAO!$M$5:$M$20,1)>0),"• “Gặp xung thì như bị
    #    roi quất”; “mã tinh bị xung mất thì ngày đi nghìn dặm” → rời quê, đi xa rất rõ. ","")&IF(AND((COUNTIF($C$92:$F$92,"có*")>0),NOT((COUNTIF($C$92:$F$92,
    #   "có")>0))),"• “Mã tinh bị hợp chặt thì có cũng như không”. ","")&IF(AND((COUNTIF($C$92:$F$92,"có")>0),COUNTIFS(SAO!$U$5:$U$20,1,SAO!$Q$5:$Q$20,1)>0),"
    #   • Mã gặp lộc (cùng một chi): “gặp lộc thì phúc nhiều”; “gặp thiên lộc, chính lộc … thì thông minh cao quý”. ","")&IF(AND((COUNTIF($C$92:$F$92,"có")>0)
    #   ,COUNTIFS(SAO!$U$5:$U$20,1,SAO!$P$5:$P$20,1)>0),"• “Quý nhân mà gặp mã tinh thì thăng vượt”. ","")&IF(AND((COUNTIF($C$92:$F$92,"có")>0),LEN($F$92)>0),
    #   "• “Mã tinh ở trụ giờ thì người đó đi kinh doanh, hiếu động mới lợi cho cầu danh cầu lợi”. ","")&IF(AND((COUNTIF($C$92:$F$92,"có")>0),COUNTIFS(SAO!$U$
    #   5:$U$20,1,SAO!$R$5:$R$20,1)>0),"• Mã gặp kình dương: “sợ nhất là gặp kình dương đầy máu”. ","")&IF(AND((COUNTIF($C$92:$F$92,"có")>0),NHAP_LIEU!$C$22="
    #   Nữ"),"• “Nữ giới gặp mã tinh thì tâm bất an, không ở một chỗ”. ","")&IF((COUNTIF($C$92:$F$92,"có")>0),"• Mã là hỷ/dụng thì thăng tiến, thu lợi khi vận
    #    động; là kỵ thì bôn ba vất vả; “mã gặp tài của lưu niên, đại vận thì phát tài nhanh”. ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('• Mã là động, là tượng bôn ba: hiếu động, đi khắp bốn phương (quân nhân, thương nhân, ngoại giao, hay đi công tác); cũng là tiêu chí đi nước ngoài. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0)) else ''),('• “Gặp xung thì như bị roi quất”; “mã tinh bị xung mất thì ngày đi nghìn dặm” → rời quê, đi xa rất rõ. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,13,20,13), 1.0),0.0))) else '')),('• “Mã tinh bị hợp chặt thì có cũng như không”. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có*'),0.0), X.NOT(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0)))) else '')),('• Mã gặp lộc (cùng một chi): “gặp lộc thì phúc nhiều”; “gặp thiên lộc, chính lộc … thì thông minh cao quý”. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,17,20,17), 1.0),0.0))) else '')),('• “Quý nhân mà gặp mã tinh thì thăng vượt”. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,16,20,16), 1.0),0.0))) else '')),('• “Mã tinh ở trụ giờ thì người đó đi kinh doanh, hiếu động mới lợi cho cầu danh cầu lợi”. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0), X.gt(X.LEN(E.v(12,92,6)),0.0))) else '')),('• Mã gặp kình dương: “sợ nhất là gặp kình dương đầy máu”. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,18,20,18), 1.0),0.0))) else '')),('• “Nữ giới gặp mã tinh thì tâm bất an, không ở một chỗ”. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0), X.eq(E.v(2,22,3),'Nữ'))) else '')),('• Mã là hỷ/dụng thì thăng tiến, thu lợi khi vận động; là kỵ thì bôn ba vất vả; “mã gặp tài của lưu niên, đại vận thì phát tài nhanh”. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,92,3,92,6), 'có'),0.0)) else '')))

def f136(E, r, c):
    # Ô C93
    # = IF(SAO!$V$5=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,5,22),1.0)) else '')

def f137(E, r, c):
    # Ô D93
    # = IF(SAO!$V$9=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,9,22),1.0)) else '')

def f138(E, r, c):
    # Ô E93
    # = IF(SAO!$V$13=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,13,22),1.0)) else '')

def f139(E, r, c):
    # Ô F93
    # = IF(SAO!$V$17=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,17,22),1.0)) else '')

def f140(E, r, c):
    # Ô I93
    # = TRIM(IF((COUNTIF($C$93:$F$93,"có*")>0),"• Vừa văn vừa võ, có khả năng nắm quyền, được tín phục. ",""))
    return X.TRIM(('• Vừa văn vừa võ, có khả năng nắm quyền, được tín phục. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,93,3,93,6), 'có*'),0.0)) else ''))

def f141(E, r, c):
    # Ô C94, C95, C138, D138, F138, C139, D139, F139, C140, D140, F140, C141 … (20 ô)
    # = ""
    return ''

def f142(E, r, c):
    # Ô D94
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$J$122:$J$133,TRUNG_GIAN!$C$6+1),"có","")
    return ('có' if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,10,133,10), X.add(E.v(23,6,3),1.0)))) else '')

def f143(E, r, c):
    # Ô E94
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$J$122:$J$133,TRUNG_GIAN!$C$6+1),"có","")
    return ('có' if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,10,133,10), X.add(E.v(23,6,3),1.0)))) else '')

def f144(E, r, c):
    # Ô F94
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$J$122:$J$133,TRUNG_GIAN!$C$6+1),"có","")
    return ('có' if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,10,133,10), X.add(E.v(23,6,3),1.0)))) else '')

def f145(E, r, c):
    # Ô H94, H95, H133, H138, H139, H140, H141, H142, H143
    # = IF(COUNTIF(C94:F94,"có")>0,"Có","Không có")
    return ('Có' if X.truthy(X.gt(X.COUNTIF(E.rg(12,r,c-5,r,c-2), 'có'),0.0)) else 'Không có')

def f146(E, r, c):
    # Ô I94
    # = TRIM(IF(COUNTIF($C$94:$F$94,"có")>0,"• Chủ hôn nhân không thuận, khắc lục thân; gặp vận tốt thì hôn nhân muộn. ","")&IF(AND(COUNTIF($C$94:$F$94,"có")>
    #   0,NHAP_LIEU!$C$22="Nam"),"• Nam gặp cô thần: thường tha phương. ","")&IF(AND(COUNTIF($C$94:$F$94,"có")>0,OR(COUNTIF(SAO!$P$5:$P$20,1)>0,COUNTIF($C$85:
    #   $F$85,"có")>0,COUNTIF($C$86:$F$86,"có")>0)),"• Có quý thần: giảm hung (tứ trụ phối hợp tốt). ",""))
    return X.TRIM(X.cat(X.cat(('• Chủ hôn nhân không thuận, khắc lục thân; gặp vận tốt thì hôn nhân muộn. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,94,3,94,6), 'có'),0.0)) else ''),('• Nam gặp cô thần: thường tha phương. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,94,3,94,6), 'có'),0.0), X.eq(E.v(2,22,3),'Nam'))) else '')),('• Có quý thần: giảm hung (tứ trụ phối hợp tốt). ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,94,3,94,6), 'có'),0.0), X.OR(X.gt(X.COUNTIF(E.rg(13,5,16,20,16), 1.0),0.0), X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có'),0.0), X.gt(X.COUNTIF(E.rg(12,86,3,86,6), 'có'),0.0)))) else '')))

def f147(E, r, c):
    # Ô D95
    # = IF(TRUNG_GIAN!$D$6=INDEX(CSDL!$K$122:$K$133,TRUNG_GIAN!$C$6+1),"có","")
    return ('có' if X.truthy(X.eq(E.v(23,6,4),X.INDEX(E.rg(24,122,11,133,11), X.add(E.v(23,6,3),1.0)))) else '')

def f148(E, r, c):
    # Ô E95
    # = IF(TRUNG_GIAN!$E$6=INDEX(CSDL!$K$122:$K$133,TRUNG_GIAN!$C$6+1),"có","")
    return ('có' if X.truthy(X.eq(E.v(23,6,5),X.INDEX(E.rg(24,122,11,133,11), X.add(E.v(23,6,3),1.0)))) else '')

def f149(E, r, c):
    # Ô F95
    # = IF(TRUNG_GIAN!$F$6=INDEX(CSDL!$K$122:$K$133,TRUNG_GIAN!$C$6+1),"có","")
    return ('có' if X.truthy(X.eq(E.v(23,6,6),X.INDEX(E.rg(24,122,11,133,11), X.add(E.v(23,6,3),1.0)))) else '')

def f150(E, r, c):
    # Ô I95
    # = TRIM(IF(COUNTIF($C$95:$F$95,"có")>0,"• Chủ hôn nhân không thuận, khắc lục thân; gặp vận tốt thì hôn nhân muộn. ","")&IF(AND(COUNTIF($C$95:$F$95,"có")>
    #   0,NHAP_LIEU!$C$22="Nữ"),"• Nữ gặp quả tú: thường khó giai lão. ",""))
    return X.TRIM(X.cat(('• Chủ hôn nhân không thuận, khắc lục thân; gặp vận tốt thì hôn nhân muộn. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,95,3,95,6), 'có'),0.0)) else ''),('• Nữ gặp quả tú: thường khó giai lão. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,95,3,95,6), 'có'),0.0), X.eq(E.v(2,22,3),'Nữ'))) else '')))

def f151(E, r, c):
    # Ô C96
    # = IF(SAO!$P$6=1,IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.eq(E.v(13,6,16),1.0)) else '')

def f152(E, r, c):
    # Ô D96
    # = IF(SAO!$P$10=1,IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.eq(E.v(13,10,16),1.0)) else '')

def f153(E, r, c):
    # Ô E96
    # = IF(SAO!$P$14=1,IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.eq(E.v(13,14,16),1.0)) else '')

def f154(E, r, c):
    # Ô F96
    # = IF(SAO!$P$18=1,IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.eq(E.v(13,18,16),1.0)) else '')

def f155(E, r, c):
    # Ô I96
    # = TRIM(IF(COUNTIF($C$96:$F$96,"có")>0,"• Thông minh, gặp việc có người giúp, gặp hung hóa cát. ","")&IF(AND(COUNTIF($C$96:$F$96,"có")>0,DUNG_HY_KY!$C$15
    #   5=1),"• Thân vượng: phúc quý càng tăng, ít bệnh. ","")&IF(AND(COUNTIF($C$96:$F$96,"có")>0,DUNG_HY_KY!$C$155=0),"• Thân nhược: bệnh nhiều hoặc giảm phú
    #   c. ","")&IF(AND(COUNTIF($C$96:$F$96,"có")>0,OR(COUNTIF($C$85:$F$85,"có")>0,COUNTIF($C$86:$F$86,"có")>0)),"• Gặp cả thiên/nguyệt đức: quý vô cùng. ",""
    #   )&IF(AND((COUNTIF($C$96:$F$96,"có*")>0),COUNTIFS(SAO!$P$5:$P$20,1,SAO!$N$5:$N$20,1)>0),"• Quý nhân gặp tuần không: giảm lực. ","")&IF(AND(NOT(COUNTIF(
    #   $C$96:$F$96,"có")>0),(COUNTIF($C$96:$F$96,"có*")>0)),"• Bị hình xung khắc hại: như không có. ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(('• Thông minh, gặp việc có người giúp, gặp hung hóa cát. ' if X.truthy(X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có'),0.0)) else ''),('• Thân vượng: phúc quý càng tăng, ít bệnh. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có'),0.0), X.eq(E.v(8,155,3),1.0))) else '')),('• Thân nhược: bệnh nhiều hoặc giảm phúc. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có'),0.0), X.eq(E.v(8,155,3),0.0))) else '')),('• Gặp cả thiên/nguyệt đức: quý vô cùng. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có'),0.0), X.OR(X.gt(X.COUNTIF(E.rg(12,85,3,85,6), 'có'),0.0), X.gt(X.COUNTIF(E.rg(12,86,3,86,6), 'có'),0.0)))) else '')),('• Quý nhân gặp tuần không: giảm lực. ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có*'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,16,20,16), 1.0, E.rg(13,5,14,20,14), 1.0),0.0))) else '')),('• Bị hình xung khắc hại: như không có. ' if X.truthy(X.AND(X.NOT(X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có'),0.0)), X.gt(X.COUNTIF(E.rg(12,96,3,96,6), 'có*'),0.0))) else '')))

def f156(E, r, c):
    # Ô C102
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7|","|4|10|1|7|","|2|11|","|2
    #   |11|","|5|8|","|5|8|"))),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7|","|4|10|1|7|","|2|
    #   11|","|2|11|","|5|8|","|5|8|"))))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|')))))) else '')

def f157(E, r, c):
    # Ô D102
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7|","|4|10|1|7|","|2|11|","|2
    #   |11|","|5|8|","|5|8|"))),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7|","|4|10|1|7|","|2|
    #   11|","|2|11|","|5|8|","|5|8|"))))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|')))))) else '')

def f158(E, r, c):
    # Ô E102
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7|","|4|10|1|7|","|2|11|","|2
    #   |11|","|5|8|","|5|8|"))),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7|","|4|10|1|7|","|2|
    #   11|","|2|11|","|5|8|","|5|8|"))))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|')))))) else '')

def f159(E, r, c):
    # Ô F102
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|4|10|1|7
    #   |","|4|10|1|7|","|2|11|","|2|11|","|5|8|","|5|8|"))),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|6|","|0|6|","|3|9|","|3|9|","|
    #   4|10|1|7|","|4|10|1|7|","|2|11|","|2|11|","|5|8|","|5|8|"))))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|6|', lambda: '|0|6|', lambda: '|3|9|', lambda: '|3|9|', lambda: '|4|10|1|7|', lambda: '|4|10|1|7|', lambda: '|2|11|', lambda: '|2|11|', lambda: '|5|8|', lambda: '|5|8|')))))) else '')

def f160(E, r, c):
    # Ô I102
    # = IF(H102="Có",TRIM("• “Chủ về thông minh hiếu học … tính cách chính trực, làm việc chuyên chú, có đầu có đuôi” "),"")
    return (X.TRIM('• “Chủ về thông minh hiếu học … tính cách chính trực, làm việc chuyên chú, có đầu có đuôi” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f161(E, r, c):
    # Ô C103
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|4|","|3|1|"))
    #   ),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|4|","|3|1|"))))),IF(THAN_SAT
    #   !$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|')))))) else '')

def f162(E, r, c):
    # Ô D103
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|4|","|3|1|"))
    #   ),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|4|","|3|1|"))))),IF(THAN_SAT
    #   !$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|')))))) else '')

def f163(E, r, c):
    # Ô E103
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|4|","|3|1|"))
    #   ),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|4|","|3|1|"))))),IF(THAN_SAT
    #   !$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|')))))) else '')

def f164(E, r, c):
    # Ô F103
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|"
    #   ,"|6|","|5|","|4|","|3|1|"))),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|2|0|","|3|1|","|2|0|","|11|","|8|","|7|","|6|","|5|","|
    #   4|","|3|1|"))))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|2|0|', lambda: '|3|1|', lambda: '|2|0|', lambda: '|11|', lambda: '|8|', lambda: '|7|', lambda: '|6|', lambda: '|5|', lambda: '|4|', lambda: '|3|1|')))))) else '')

def f165(E, r, c):
    # Ô I103
    # = IF(H103="Có",TRIM("• “Cuộc đời phúc lộc dồi dào … phần nhiều chỉ về bình an, có phúc chứ không chỉ sự giàu sang” "),"")
    return (X.TRIM('• “Cuộc đời phúc lộc dồi dào … phần nhiều chỉ về bình an, có phúc chứ không chỉ sự giàu sang” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f166(E, r, c):
    # Ô C104
    # = IF(AND(TRUE(),AND(OR(0=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|","|100|82|76|58|")))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), '|100|82|76|58|'))))) else '')

def f167(E, r, c):
    # Ô D104
    # = IF(AND(TRUE(),AND(OR(1=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|","|100|82|76|58|")))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), '|100|82|76|58|'))))) else '')

def f168(E, r, c):
    # Ô E104
    # = IF(AND(TRUE(),AND(OR(2=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|100|82|76|58|")))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|100|82|76|58|'))))) else '')

def f169(E, r, c):
    # Ô F104
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|","|100|82|76|58|")))),IF(THAN_SAT!$F$1
    #   1=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), '|100|82|76|58|'))))) else '')

def f170(E, r, c):
    # Ô I104
    # = IF(H104="Có",TRIM("• “Tính cách thông tuệ, ưa sát phạt”; “việc hôn nhân thường trắc trở” "&IF(AND(OR((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=82,(TRUNG_GI
    #   AN!$E$5*12+TRUNG_GIAN!$E$6)=76),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1)>0)),"• Ngày Canh Tuất/Canh Thìn mà trụ có quan: “là không tốt”
    #    ","")&IF(AND(OR((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=58,(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=100),OR((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X
    #   $20,1)>0),(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$X$5:$X$20,1)>0))),"• Ngày Mậu Tuất/Nhâm Thìn mà trụ có tài: “là không tốt” ","")&IF(AND((IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$155)=1),NOT((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1)>0)),NOT((COUNTIFS(SAO!$F$5:$F$20,"Thất s
    #   át",SAO!$X$5:$X$20,1)>0))),"• “Thân vượng thì phát phúc to” ","")&IF(THAN_SAT!$E$11=1,"• “Có hình xung phá hại thì người đó nghèo rách, kiệt quệ” ",""
    #   )),"")
    return (X.TRIM(X.cat(X.cat(X.cat(X.cat('• “Tính cách thông tuệ, ưa sát phạt”; “việc hôn nhân thường trắc trở” ',('• Ngày Canh Tuất/Canh Thìn mà trụ có quan: “là không tốt” ' if X.truthy(X.AND(X.OR(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),82.0), X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),76.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0),0.0))) else '')),('• Ngày Mậu Tuất/Nhâm Thìn mà trụ có tài: “là không tốt” ' if X.truthy(X.AND(X.OR(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),58.0), X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),100.0)), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0),0.0)))) else '')),('• “Thân vượng thì phát phúc to” ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,155,3)),1.0), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0),0.0)))) else '')),('• “Có hình xung phá hại thì người đó nghèo rách, kiệt quệ” ' if X.truthy(X.eq(E.v(12,11,5),1.0)) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f171(E, r, c):
    # Ô C105
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|"))),ISNUM
    #   BER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|"))))),IF(THAN_SAT!$C$11=1,"có –
    #    vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|')))))) else '')

def f172(E, r, c):
    # Ô D105
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|"))),ISNUM
    #   BER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|"))))),IF(THAN_SAT!$D$11=1,"có –
    #    vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|')))))) else '')

def f173(E, r, c):
    # Ô E105
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|"))),ISNUM
    #   BER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|"))))),IF(THAN_SAT!$E$11=1,"có –
    #    vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|')))))) else '')

def f174(E, r, c):
    # Ô F105
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|
    #   ","|5|","|7|","|8|"))),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|10|","|11|","|1|","|2|","|1|","|2|","|4|","|5|","|7|","|8|")))
    #   )),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|', lambda: '|1|', lambda: '|2|', lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|')))))) else '')

def f175(E, r, c):
    # Ô I105
    # = IF(H105="Có",TRIM("• “Thành thực đáng tin, nghiêm túc thanh liêm, quy củ, hành sự theo lẽ công”; sinh vượng, không bị xung phá thì “làm chức to … nắm 
    #   quyền quản lý” "),"")
    return (X.TRIM('• “Thành thực đáng tin, nghiêm túc thanh liêm, quy củ, hành sự theo lẽ công”; sinh vượng, không bị xung phá thì “làm chức to … nắm quyền quản lý” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f176(E, r, c):
    # Ô C106
    # = IF(AND(TRUE(),OR(TRUNG_GIAN!$C$6=CHOOSE(INDEX(NAPAM_H,MOD(6*TRUNG_GIAN!$C$5-5*TRUNG_GIAN!$C$6,60)+1)+1,11,2,8,5,8),TRUNG_GIAN!$C$6=CHOOSE(INDEX(NAPAM_
    #   H,MOD(6*TRUNG_GIAN!$E$5-5*TRUNG_GIAN!$E$6,60)+1)+1,11,2,8,5,8))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0)), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,5)),X.mul(5.0,E.v(23,6,5))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0))))) else '')

def f177(E, r, c):
    # Ô D106
    # = IF(AND(TRUE(),OR(TRUNG_GIAN!$D$6=CHOOSE(INDEX(NAPAM_H,MOD(6*TRUNG_GIAN!$C$5-5*TRUNG_GIAN!$C$6,60)+1)+1,11,2,8,5,8),TRUNG_GIAN!$D$6=CHOOSE(INDEX(NAPAM_
    #   H,MOD(6*TRUNG_GIAN!$E$5-5*TRUNG_GIAN!$E$6,60)+1)+1,11,2,8,5,8))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0)), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,5)),X.mul(5.0,E.v(23,6,5))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0))))) else '')

def f178(E, r, c):
    # Ô E106
    # = IF(AND(TRUE(),OR(TRUNG_GIAN!$E$6=CHOOSE(INDEX(NAPAM_H,MOD(6*TRUNG_GIAN!$C$5-5*TRUNG_GIAN!$C$6,60)+1)+1,11,2,8,5,8),TRUNG_GIAN!$E$6=CHOOSE(INDEX(NAPAM_
    #   H,MOD(6*TRUNG_GIAN!$E$5-5*TRUNG_GIAN!$E$6,60)+1)+1,11,2,8,5,8))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0)), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,5)),X.mul(5.0,E.v(23,6,5))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0))))) else '')

def f179(E, r, c):
    # Ô F106
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=CHOOSE(INDEX(NAPAM_H,MOD(6*TRUNG_GIAN!$C$5-5*TRUNG_GIAN!$C$6,60)+1)+1,11,2,8,5,8),TRUNG_G
    #   IAN!$F$6=CHOOSE(INDEX(NAPAM_H,MOD(6*TRUNG_GIAN!$E$5-5*TRUNG_GIAN!$E$6,60)+1)+1,11,2,8,5,8))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0)), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(X.MOD(X.sub(X.mul(6.0,E.v(23,5,5)),X.mul(5.0,E.v(23,6,5))), 60.0),1.0)),1.0), lambda: 11.0, lambda: 2.0, lambda: 8.0, lambda: 5.0, lambda: 8.0))))) else '')

def f180(E, r, c):
    # Ô I106
    # = IF(H106="Có",TRIM("• “Chủ về thứ bậc thi đỗ hoặc tiền đồ học nghiệp”; “tú khí phát sinh, thông minh khéo léo, văn chương nổi tiếng” "),"")
    return (X.TRIM('• “Chủ về thứ bậc thi đỗ hoặc tiền đồ học nghiệp”; “tú khí phát sinh, thông minh khéo léo, văn chương nổi tiếng” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f181(E, r, c):
    # Ô C107
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|74|","|95|","|17|","|54|","|41|","|78|","|104|"
    #   ,"|117|","|119|","|106|"))),ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|74|","|95|","|17|","|54|","|41|","|7
    #   8|","|104|","|117|","|119|","|106|"))))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|')))))) else '')

def f182(E, r, c):
    # Ô D107
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|74|","|95|","|17|","|54|","|41|","|78|","|104|"
    #   ,"|117|","|119|","|106|"))),ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|74|","|95|","|17|","|54|","|41|","|7
    #   8|","|104|","|117|","|119|","|106|"))))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|')))))) else '')

def f183(E, r, c):
    # Ô E107
    # = IF(AND(TRUE(),OR(ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|74|","|95|","|17|","|54|","|41|","|78|","|104|"
    #   ,"|117|","|119|","|106|"))),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|74|","|95|","|17|","|54|","|41|","|7
    #   8|","|104|","|117|","|119|","|106|"))))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|')))))) else '')

def f184(E, r, c):
    # Ô F107
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|74|","|95|","|17|",
    #   "|54|","|41|","|78|","|104|","|117|","|119|","|106|"))),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|74|","|9
    #   5|","|17|","|54|","|41|","|78|","|104|","|117|","|119|","|106|"))))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|'))), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|74|', lambda: '|95|', lambda: '|17|', lambda: '|54|', lambda: '|41|', lambda: '|78|', lambda: '|104|', lambda: '|117|', lambda: '|119|', lambda: '|106|')))))) else '')

def f185(E, r, c):
    # Ô I107
    # = IF(H107="Có",TRIM("• “Học vấn tinh thông, văn chương giỏi”; “nên được sinh vượng không nên bị khắc, hại, xung, phá” "),"")
    return (X.TRIM('• “Học vấn tinh thông, văn chương giỏi”; “nên được sinh vượng không nên bị khắc, hại, xung, phá” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f186(E, r, c):
    # Ô C108
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|8|9|4|5|","|6|7|","|2|3|","|0|1|","|8|9|4|5|","|6|7|","|2|3|","|0|1|","
    #   |8|9|4|5|","|6|7|","|2|3|","|0|1|")))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,3)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|'))))) else '')

def f187(E, r, c):
    # Ô D108
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|8|9|4|5|","|6|7|","|2|3|","|0|1|","|8|9|4|5|","|6|7|","|2|3|","|0|1|","
    #   |8|9|4|5|","|6|7|","|2|3|","|0|1|")))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,4)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|'))))) else '')

def f188(E, r, c):
    # Ô E108
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|8|9|4|5|","|6|7|","|2|3|","|0|1|","|8|9|4|5|","|6|7|","|2|3|","|0|1|","
    #   |8|9|4|5|","|6|7|","|2|3|","|0|1|")))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,5)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|'))))) else '')

def f189(E, r, c):
    # Ô F108
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|8|9|4|5|","|6|7|","|2|3|","|0|1|","|8|9|4|5
    #   |","|6|7|","|2|3|","|0|1|","|8|9|4|5|","|6|7|","|2|3|","|0|1|")))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,6)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|', lambda: '|8|9|4|5|', lambda: '|6|7|', lambda: '|2|3|', lambda: '|0|1|'))))) else '')

def f190(E, r, c):
    # Ô I108
    # = IF(H108="Có",TRIM("• “Thần giải hung”; “tính thông minh, ôn hòa trung hậu”, “có thể biến hung thành cát” "),"")
    return (X.TRIM('• “Thần giải hung”; “tính thông minh, ôn hòa trung hậu”, “có thể biến hung thành cát” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f191(E, r, c):
    # Ô C109
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|2|7|0|5|","|1|6|","|4|9|","|3|8|","|2|7|0|5|","|1|6|","|4|9|","|3|8|","
    #   |2|7|0|5|","|1|6|","|4|9|","|3|8|")))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,3)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|'))))) else '')

def f192(E, r, c):
    # Ô D109
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|2|7|0|5|","|1|6|","|4|9|","|3|8|","|2|7|0|5|","|1|6|","|4|9|","|3|8|","
    #   |2|7|0|5|","|1|6|","|4|9|","|3|8|")))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,4)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|'))))) else '')

def f193(E, r, c):
    # Ô E109
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|2|7|0|5|","|1|6|","|4|9|","|3|8|","|2|7|0|5|","|1|6|","|4|9|","|3|8|","
    #   |2|7|0|5|","|1|6|","|4|9|","|3|8|")))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,5)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|'))))) else '')

def f194(E, r, c):
    # Ô F109
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$5&"|",CHOOSE(TRUNG_GIAN!$D$6+1,"|2|7|0|5|","|1|6|","|4|9|","|3|8|","|2|7|0|5
    #   |","|1|6|","|4|9|","|3|8|","|2|7|0|5|","|1|6|","|4|9|","|3|8|")))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,5,6)),'|'), X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|', lambda: '|2|7|0|5|', lambda: '|1|6|', lambda: '|4|9|', lambda: '|3|8|'))))) else '')

def f195(E, r, c):
    # Ô I109
    # = IF(H109="Có",TRIM("• “Khí thanh tú của trời đất”; “thành thật, tinh thần sáng sủa, nghĩa hiệp thanh cao, tài hoa xuất chúng” "),"")
    return (X.TRIM('• “Khí thanh tú của trời đất”; “thành thật, tinh thần sáng sủa, nghĩa hiệp thanh cao, tài hoa xuất chúng” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f196(E, r, c):
    # Ô C110
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|4|","|5|","|7|","|8|","|7|","|8|","|10|","|11|","|1|","|2|")))),IF(THAN
    #   _SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|', lambda: '|7|', lambda: '|8|', lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|'))))) else '')

def f197(E, r, c):
    # Ô D110
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|4|","|5|","|7|","|8|","|7|","|8|","|10|","|11|","|1|","|2|")))),IF(THAN
    #   _SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|', lambda: '|7|', lambda: '|8|', lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|'))))) else '')

def f198(E, r, c):
    # Ô E110
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|4|","|5|","|7|","|8|","|7|","|8|","|10|","|11|","|1|","|2|")))),IF(THAN
    #   _SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|', lambda: '|7|', lambda: '|8|', lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|'))))) else '')

def f199(E, r, c):
    # Ô F110
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|4|","|5|","|7|","|8|","|7|","|8|","|10|","|
    #   11|","|1|","|2|")))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|4|', lambda: '|5|', lambda: '|7|', lambda: '|8|', lambda: '|7|', lambda: '|8|', lambda: '|10|', lambda: '|11|', lambda: '|1|', lambda: '|2|'))))) else '')

def f200(E, r, c):
    # Ô I110
    # = IF(H110="Có",TRIM("• “Chủ về thông minh, phú quý, tính cách uyển chuyển ôn hòa, dáng người thanh thần” "&IF(NHAP_LIEU!$C$22="Nam","• Nam: “nhiều thê t
    #   hiếp, anh em hòa thuận, con cháu đông đúc” ","")&IF(NOT(NHAP_LIEU!$C$22="Nam"),"• Nữ: “nhiều phú quý” ","")),"")
    return (X.TRIM(X.cat(X.cat('• “Chủ về thông minh, phú quý, tính cách uyển chuyển ôn hòa, dáng người thanh thần” ',('• Nam: “nhiều thê thiếp, anh em hòa thuận, con cháu đông đúc” ' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else '')),('• Nữ: “nhiều phú quý” ' if X.truthy(X.NOT(X.eq(E.v(2,22,3),'Nam'))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f201(E, r, c):
    # Ô C111
    # = IF(AND(TRUE(),AND(OR(0=2,0=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|","|13|65|117|")))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),""
    #   )
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,2.0), X.eq(0.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), '|13|65|117|'))))) else '')

def f202(E, r, c):
    # Ô D111
    # = IF(AND(TRUE(),AND(OR(1=2,1=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|","|13|65|117|")))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),""
    #   )
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,2.0), X.eq(1.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), '|13|65|117|'))))) else '')

def f203(E, r, c):
    # Ô E111
    # = IF(AND(TRUE(),AND(OR(2=2,2=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|13|65|117|")))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),""
    #   )
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,2.0), X.eq(2.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|13|65|117|'))))) else '')

def f204(E, r, c):
    # Ô F111
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=2,3=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|","|13|65|117|")))),IF(THAN_SAT!$F$
    #   11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,2.0), X.eq(3.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), '|13|65|117|'))))) else '')

def f205(E, r, c):
    # Ô I111
    # = IF(H111="Có",TRIM("• “Không giàu thì sang, thông minh hiếu học, có tài, nhưng tính gấp, mạnh mẽ” "&IF(PP2_DINH_LUONG!$F$92>0,"• Trụ có Hỏa: “kim thần 
    #   gặp hỏa, uy trấn biên cương”; gặp vận Hỏa thì phát ","")&IF(AND(PP2_DINH_LUONG!$F$94/PP2_DINH_LUONG!$C$100>=0.3,PP2_DINH_LUONG!$F$95/PP2_DINH_LUONG!$C
    #   $100>=0.2),"• “Kim nhiều … thủy vượng là phá cách, chủ về họa nhiều mà thành công ít” (giả định G32: Kim ≥30%, Thủy ≥20% tổng độ PP2) ","")),"")
    return (X.TRIM(X.cat(X.cat('• “Không giàu thì sang, thông minh hiếu học, có tài, nhưng tính gấp, mạnh mẽ” ',('• Trụ có Hỏa: “kim thần gặp hỏa, uy trấn biên cương”; gặp vận Hỏa thì phát ' if X.truthy(X.gt(E.v(10,92,6),0.0)) else '')),('• “Kim nhiều … thủy vượng là phá cách, chủ về họa nhiều mà thành công ít” (giả định G32: Kim ≥30%, Thủy ≥20% tổng độ PP2) ' if X.truthy(X.AND(X.ge(X.div(E.v(10,94,6),E.v(10,100,3)),0.3), X.ge(X.div(E.v(10,95,6),E.v(10,100,3)),0.2))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f206(E, r, c):
    # Ô C112
    # = IF(AND(TRUE(),AND(0<>1,TRUNG_GIAN!$C$6=MOD(TRUNG_GIAN!$D$6+11,12))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(0.0,1.0), X.eq(E.v(23,6,3),X.MOD(X.add(E.v(23,6,4),11.0), 12.0))))) else '')

def f207(E, r, c):
    # Ô D112
    # = IF(AND(TRUE(),AND(1<>1,TRUNG_GIAN!$D$6=MOD(TRUNG_GIAN!$D$6+11,12))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(1.0,1.0), X.eq(E.v(23,6,4),X.MOD(X.add(E.v(23,6,4),11.0), 12.0))))) else '')

def f208(E, r, c):
    # Ô E112
    # = IF(AND(TRUE(),AND(2<>1,TRUNG_GIAN!$E$6=MOD(TRUNG_GIAN!$D$6+11,12))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(2.0,1.0), X.eq(E.v(23,6,5),X.MOD(X.add(E.v(23,6,4),11.0), 12.0))))) else '')

def f209(E, r, c):
    # Ô F112
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3<>1,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$D$6+11,12))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.ne(3.0,1.0), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(23,6,4),11.0), 12.0))))) else '')

def f210(E, r, c):
    # Ô I112
    # = IF(H112="Có",TRIM(IF(NOT(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$155)=1),(THAN_SAT!$H$96="Có"))),"• “Nếu không vượng lại không có cát 
    #   thần quý nhân phù trợ, thì không bệnh tật luôn cũng là người gầy mòn ốm yếu” ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$155)=1),
    #   (THAN_SAT!$H$96="Có")),"• “Sinh vượng và còn có quý nhân phù trợ thì … phù hợp với công việc nghiên cứu y học, y dược, … làm nghề y” ","")),"")
    return (X.TRIM(X.cat(('• “Nếu không vượng lại không có cát thần quý nhân phù trợ, thì không bệnh tật luôn cũng là người gầy mòn ốm yếu” ' if X.truthy(X.NOT(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,155,3)),1.0), X.eq(E.v(12,96,8),'Có')))) else ''),('• “Sinh vượng và còn có quý nhân phù trợ thì … phù hợp với công việc nghiên cứu y học, y dược, … làm nghề y” ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,155,3)),1.0), X.eq(E.v(12,96,8),'Có'))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f211(E, r, c):
    # Ô C113
    # = IF(AND(TRUE(),AND(OR(0=2,0=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=119,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=
    #   109),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=109,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=119),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=41,(TRUNG_GIAN!$F
    #   $5*12+TRUNG_GIAN!$F$6)=43),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=67,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=65),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$
    #   6)=52,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=54)))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,2.0), X.eq(0.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),119.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),109.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),109.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),119.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),41.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),43.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),67.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),65.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),52.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),54.0)))))) else '')

def f212(E, r, c):
    # Ô D113
    # = IF(AND(TRUE(),AND(OR(1=2,1=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=119,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=
    #   109),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=109,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=119),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=41,(TRUNG_GIAN!$F
    #   $5*12+TRUNG_GIAN!$F$6)=43),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=67,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=65),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$
    #   6)=52,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=54)))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,2.0), X.eq(1.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),119.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),109.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),109.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),119.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),41.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),43.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),67.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),65.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),52.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),54.0)))))) else '')

def f213(E, r, c):
    # Ô E113
    # = IF(AND(TRUE(),AND(OR(2=2,2=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=119,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=
    #   109),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=109,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=119),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=41,(TRUNG_GIAN!$F
    #   $5*12+TRUNG_GIAN!$F$6)=43),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=67,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=65),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$
    #   6)=52,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=54)))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,2.0), X.eq(2.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),119.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),109.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),109.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),119.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),41.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),43.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),67.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),65.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),52.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),54.0)))))) else '')

def f214(E, r, c):
    # Ô F113
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=2,3=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=119,(TRUNG_GI
    #   AN!$F$5*12+TRUNG_GIAN!$F$6)=109),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=109,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=119),AND((TRUNG_GIAN!$E$5*12+TRUNG_
    #   GIAN!$E$6)=41,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=43),AND((TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=67,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=65),AND((TRUNG
    #   _GIAN!$E$5*12+TRUNG_GIAN!$E$6)=52,(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)=54)))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,2.0), X.eq(3.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),119.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),109.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),109.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),119.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),41.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),43.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),67.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),65.0)), X.AND(X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),52.0), X.eq(X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6)),54.0)))))) else '')

def f215(E, r, c):
    # Ô I113
    # = IF(H113="Có",TRIM("• “Củng lộc, củng quý thuần túy là bậc vương hầu”; kị gặp quan sát, tuần không vong, kình dương ở chi tháng "),"")
    return (X.TRIM('• “Củng lộc, củng quý thuần túy là bậc vương hầu”; kị gặp quan sát, tuần không vong, kình dương ở chi tháng ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f216(E, r, c):
    # Ô C114
    # = IF(AND(TRUE(),AND(OR(0=2,0=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12),TRUN
    #   G_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12)),AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12),TRU
    #   NG_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12))))),IF(THAN_SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,2.0), X.eq(0.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0))), X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0))))))) else '')

def f217(E, r, c):
    # Ô D114
    # = IF(AND(TRUE(),AND(OR(1=2,1=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12),TRUN
    #   G_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12)),AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12),TRU
    #   NG_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12))))),IF(THAN_SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,2.0), X.eq(1.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0))), X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0))))))) else '')

def f218(E, r, c):
    # Ô E114
    # = IF(AND(TRUE(),AND(OR(2=2,2=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12),TRUN
    #   G_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12)),AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12),TRU
    #   NG_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12))))),IF(THAN_SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,2.0), X.eq(2.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0))), X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0))))))) else '')

def f219(E, r, c):
    # Ô F114
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=2,3=3),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5
    #   ,6,5,6,8,9,11,0)+11,12),TRUNG_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+1,12)),AND(TRUNG_GIAN!$E$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3
    #   ,5,6,5,6,8,9,11,0)+1,12),TRUNG_GIAN!$F$6=MOD(CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)+11,12))))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,2.0), X.eq(3.0,3.0)), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0))), X.AND(X.eq(E.v(23,6,5),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),1.0), 12.0)), X.eq(E.v(23,6,6),X.MOD(X.add(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0),11.0), 12.0))))))) else '')

def f220(E, r, c):
    # Ô I114
    # = IF(H114="Có",TRIM("• “Giáp lộc là có cát thần phù trợ, chủ về giàu thọ quan quý” "),"")
    return (X.TRIM('• “Giáp lộc là có cát thần phù trợ, chủ về giàu thọ quan quý” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f221(E, r, c):
    # Ô C115
    # = IF(AND(TRUE(),AND(0=2,(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,50,6,56,0))),IF(THAN_SAT!$C$11=1,"có – vô hiệu",
    #   "có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(0.0,2.0), X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: 50.0, lambda: 6.0, lambda: 56.0, lambda: 0.0))))) else '')

def f222(E, r, c):
    # Ô D115
    # = IF(AND(TRUE(),AND(1=2,(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,50,6,56,0))),IF(THAN_SAT!$D$11=1,"có – vô hiệu",
    #   "có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(1.0,2.0), X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: 50.0, lambda: 6.0, lambda: 56.0, lambda: 0.0))))) else '')

def f223(E, r, c):
    # Ô E115
    # = IF(AND(TRUE(),AND(2=2,(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,50,6,56,0))),IF(THAN_SAT!$E$11=1,"có – vô hiệu",
    #   "có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(2.0,2.0), X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: 50.0, lambda: 6.0, lambda: 56.0, lambda: 0.0))))) else '')

def f224(E, r, c):
    # Ô F115
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3=2,(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)=CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,50,6,56,0))),IF(THAN
    #   _SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.eq(3.0,2.0), X.eq(X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5)),X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: 50.0, lambda: 6.0, lambda: 56.0, lambda: 0.0))))) else '')

def f225(E, r, c):
    # Ô I115
    # = IF(H115="Có",TRIM("• “Gặp hung hóa cát, sao giải tai họa. Trong mệnh gặp nó là có cứu” "),"")
    return (X.TRIM('• “Gặp hung hóa cát, sao giải tai họa. Trong mệnh gặp nó là có cứu” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f226(E, r, c):
    # Ô C116
    # = IF(AND(TRUE(),OR(AND(0<=2,OR(AND(TRUNG_GIAN!$C$5=0,TRUNG_GIAN!$D$5=4,TRUNG_GIAN!$E$5=6),AND(TRUNG_GIAN!$C$5=1,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$E$5=3),AND
    #   (TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$D$5=9,TRUNG_GIAN!$E$5=7))),AND(0>=1,AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$5=0,TRUNG_GIAN!$E$5=4,T
    #   RUNG_GIAN!$F$5=6),AND(TRUNG_GIAN!$D$5=1,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$F$5=3),AND(TRUNG_GIAN!$D$5=8,TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$F$5=7)))))),IF(THAN_S
    #   AT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.le(0.0,2.0), X.OR(X.AND(X.eq(E.v(23,5,3),0.0), X.eq(E.v(23,5,4),4.0), X.eq(E.v(23,5,5),6.0)), X.AND(X.eq(E.v(23,5,3),1.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,5),3.0)), X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,5,4),9.0), X.eq(E.v(23,5,5),7.0)))), X.AND(X.ge(0.0,1.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,4),0.0), X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,6),6.0)), X.AND(X.eq(E.v(23,5,4),1.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,6),3.0)), X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,5,6),7.0)))))))) else '')

def f227(E, r, c):
    # Ô D116
    # = IF(AND(TRUE(),OR(AND(1<=2,OR(AND(TRUNG_GIAN!$C$5=0,TRUNG_GIAN!$D$5=4,TRUNG_GIAN!$E$5=6),AND(TRUNG_GIAN!$C$5=1,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$E$5=3),AND
    #   (TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$D$5=9,TRUNG_GIAN!$E$5=7))),AND(1>=1,AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$5=0,TRUNG_GIAN!$E$5=4,T
    #   RUNG_GIAN!$F$5=6),AND(TRUNG_GIAN!$D$5=1,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$F$5=3),AND(TRUNG_GIAN!$D$5=8,TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$F$5=7)))))),IF(THAN_S
    #   AT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.le(1.0,2.0), X.OR(X.AND(X.eq(E.v(23,5,3),0.0), X.eq(E.v(23,5,4),4.0), X.eq(E.v(23,5,5),6.0)), X.AND(X.eq(E.v(23,5,3),1.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,5),3.0)), X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,5,4),9.0), X.eq(E.v(23,5,5),7.0)))), X.AND(X.ge(1.0,1.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,4),0.0), X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,6),6.0)), X.AND(X.eq(E.v(23,5,4),1.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,6),3.0)), X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,5,6),7.0)))))))) else '')

def f228(E, r, c):
    # Ô E116
    # = IF(AND(TRUE(),OR(AND(2<=2,OR(AND(TRUNG_GIAN!$C$5=0,TRUNG_GIAN!$D$5=4,TRUNG_GIAN!$E$5=6),AND(TRUNG_GIAN!$C$5=1,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$E$5=3),AND
    #   (TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$D$5=9,TRUNG_GIAN!$E$5=7))),AND(2>=1,AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$5=0,TRUNG_GIAN!$E$5=4,T
    #   RUNG_GIAN!$F$5=6),AND(TRUNG_GIAN!$D$5=1,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$F$5=3),AND(TRUNG_GIAN!$D$5=8,TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$F$5=7)))))),IF(THAN_S
    #   AT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.le(2.0,2.0), X.OR(X.AND(X.eq(E.v(23,5,3),0.0), X.eq(E.v(23,5,4),4.0), X.eq(E.v(23,5,5),6.0)), X.AND(X.eq(E.v(23,5,3),1.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,5),3.0)), X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,5,4),9.0), X.eq(E.v(23,5,5),7.0)))), X.AND(X.ge(2.0,1.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,4),0.0), X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,6),6.0)), X.AND(X.eq(E.v(23,5,4),1.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,6),3.0)), X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,5,6),7.0)))))))) else '')

def f229(E, r, c):
    # Ô F116
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(3<=2,OR(AND(TRUNG_GIAN!$C$5=0,TRUNG_GIAN!$D$5=4,TRUNG_GIAN!$E$5=6),AND(TRUNG_GIAN!$C$5=1,TRUNG_GIAN!$
    #   D$5=2,TRUNG_GIAN!$E$5=3),AND(TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$D$5=9,TRUNG_GIAN!$E$5=7))),AND(3>=1,AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIA
    #   N!$D$5=0,TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$F$5=6),AND(TRUNG_GIAN!$D$5=1,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$F$5=3),AND(TRUNG_GIAN!$D$5=8,TRUNG_GIAN!$E$5=9,TRUNG
    #   _GIAN!$F$5=7)))))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.le(3.0,2.0), X.OR(X.AND(X.eq(E.v(23,5,3),0.0), X.eq(E.v(23,5,4),4.0), X.eq(E.v(23,5,5),6.0)), X.AND(X.eq(E.v(23,5,3),1.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,5),3.0)), X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,5,4),9.0), X.eq(E.v(23,5,5),7.0)))), X.AND(X.ge(3.0,1.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,4),0.0), X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,6),6.0)), X.AND(X.eq(E.v(23,5,4),1.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,6),3.0)), X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,5,6),7.0)))))))) else '')

def f230(E, r, c):
    # Ô I116
    # = IF(H116="Có",TRIM("• “Tinh thần người đó khác thường, hoài bão to lớn … biết rộng, nhiều tài” "&IF((THAN_SAT!$H$96="Có"),"• “Có thêm cả thiên ất quý n
    #   hân thì công danh siêu quần” ","")&IF((OR(THAN_SAT!$H$85="Có",THAN_SAT!$H$86="Có")),"• “Có thêm thiên đức, nguyệt đức thì không bao giờ gặp điều tai v
    #   ạ” ","")&IF(NOT((THAN_SAT!$H$96="Có")),"• “Nếu chỉ có tam kỳ, không có quý nhân, mệnh cục lung tung thì cũng sẽ bần cùng thấp kém” ","")),"")
    return (X.TRIM(X.cat(X.cat(X.cat('• “Tinh thần người đó khác thường, hoài bão to lớn … biết rộng, nhiều tài” ',('• “Có thêm cả thiên ất quý nhân thì công danh siêu quần” ' if X.truthy(X.eq(E.v(12,96,8),'Có')) else '')),('• “Có thêm thiên đức, nguyệt đức thì không bao giờ gặp điều tai vạ” ' if X.truthy(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))) else '')),('• “Nếu chỉ có tam kỳ, không có quý nhân, mệnh cục lung tung thì cũng sẽ bần cùng thấp kém” ' if X.truthy(X.NOT(X.eq(E.v(12,96,8),'Có'))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f231(E, r, c):
    # Ô C117
    # = IF(AND(TRUE(),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ
    #    giờ"),TRUNG_GIAN!$F$6=10)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$6=11))),OR(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$C$6=11))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))), X.OR(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,3),11.0))))) else '')

def f232(E, r, c):
    # Ô D117
    # = IF(AND(TRUE(),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ
    #    giờ"),TRUNG_GIAN!$F$6=10)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$6=11))),OR(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$D$6=11))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))), X.OR(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,4),11.0))))) else '')

def f233(E, r, c):
    # Ô E117
    # = IF(AND(TRUE(),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ
    #    giờ"),TRUNG_GIAN!$F$6=10)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$6=11))),OR(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$E$6=11))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))), X.OR(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,5),11.0))))) else '')

def f234(E, r, c):
    # Ô F117
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11
    #   ),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=11))),OR(TRUNG_GIAN!$F$6=10,TRUNG_GIAN!$F$6=11))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))), X.OR(X.eq(E.v(23,6,6),10.0), X.eq(E.v(23,6,6),11.0))))) else '')

def f235(E, r, c):
    # Ô I117
    # = IF(H117="Có",TRIM("• “Thường là người hay gặp nạn về hình pháp … còn chủ về tai nạn, bệnh tật” "&IF(NHAP_LIEU!$C$22="Nam","• “Nam sợ thiên la” ","")&I
    #   F((OR(THAN_SAT!$H$85="Có",THAN_SAT!$H$86="Có")),"• Có thiên/nguyệt đức: “có giải cứu … không gặp tai vạ” ","")),"")
    return (X.TRIM(X.cat(X.cat('• “Thường là người hay gặp nạn về hình pháp … còn chủ về tai nạn, bệnh tật” ',('• “Nam sợ thiên la” ' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else '')),('• Có thiên/nguyệt đức: “có giải cứu … không gặp tai vạ” ' if X.truthy(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f236(E, r, c):
    # Ô C118
    # = IF(AND(TRUE(),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ gi
    #   ờ"),TRUNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=5))),OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=5))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),5.0))))) else '')

def f237(E, r, c):
    # Ô D118
    # = IF(AND(TRUE(),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ gi
    #   ờ"),TRUNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=5))),OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=5))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),5.0))))) else '')

def f238(E, r, c):
    # Ô E118
    # = IF(AND(TRUE(),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ gi
    #   ờ"),TRUNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=5))),OR(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$E$6=5))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),5.0))))) else '')

def f239(E, r, c):
    # Ô F118
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT
    #   (NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(N
    #   OT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=5))),OR(TRUNG_GIAN!$F$6=4,TRUNG_GIAN!$F$6=5))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.eq(E.v(23,6,6),4.0), X.eq(E.v(23,6,6),5.0))))) else '')

def f240(E, r, c):
    # Ô I118
    # = IF(H118="Có",TRIM("• “Thường là người hay gặp nạn về hình pháp … còn chủ về tai nạn, bệnh tật” "&IF(NOT(NHAP_LIEU!$C$22="Nam"),"• “Nữ kị địa võng” ","
    #   ")&IF((OR(THAN_SAT!$H$85="Có",THAN_SAT!$H$86="Có")),"• Có thiên/nguyệt đức: “có giải cứu … không gặp tai vạ” ","")),"")
    return (X.TRIM(X.cat(X.cat('• “Thường là người hay gặp nạn về hình pháp … còn chủ về tai nạn, bệnh tật” ',('• “Nữ kị địa võng” ' if X.truthy(X.NOT(X.eq(E.v(2,22,3),'Nam'))) else '')),('• Có thiên/nguyệt đức: “có giải cứu … không gặp tai vạ” ' if X.truthy(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f241(E, r, c):
    # Ô C119
    # = IF(AND(TRUE(),AND(0<>0,TRUNG_GIAN!$C$6=CHOOSE(TRUNG_GIAN!$C$6+1,6,3,0,9,6,3,0,9,6,3,0,9))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(0.0,0.0), X.eq(E.v(23,6,3),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0))))) else '')

def f242(E, r, c):
    # Ô D119
    # = IF(AND(TRUE(),AND(1<>0,TRUNG_GIAN!$D$6=CHOOSE(TRUNG_GIAN!$C$6+1,6,3,0,9,6,3,0,9,6,3,0,9))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(1.0,0.0), X.eq(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0))))) else '')

def f243(E, r, c):
    # Ô E119
    # = IF(AND(TRUE(),AND(2<>0,TRUNG_GIAN!$E$6=CHOOSE(TRUNG_GIAN!$C$6+1,6,3,0,9,6,3,0,9,6,3,0,9))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(2.0,0.0), X.eq(E.v(23,6,5),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0))))) else '')

def f244(E, r, c):
    # Ô F119
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3<>0,TRUNG_GIAN!$F$6=CHOOSE(TRUNG_GIAN!$C$6+1,6,3,0,9,6,3,0,9,6,3,0,9))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.ne(3.0,0.0), X.eq(E.v(23,6,6),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0, lambda: 6.0, lambda: 3.0, lambda: 0.0, lambda: 9.0))))) else '')

def f245(E, r, c):
    # Ô I119
    # = IF(H119="Có",TRIM("• “Phúc ít họa nhiều. Sát chủ về máu me”; “đi với thủy, hỏa phải đề phòng bỏng, cháy, đi với kim mộc đề phòng bị đánh, đi với thổ p
    #   hải đề phòng đổ sập, ôn dịch” "&IF(OR((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1)>0),(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$X$5:$X$20,1)
    #   >0),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$X$5:$X$20,1)>0)),"• “Gặp quan tinh ấn thụ là tốt” ","")),"")
    return (X.TRIM(X.cat('• “Phúc ít họa nhiều. Sát chủ về máu me”; “đi với thủy, hỏa phải đề phòng bỏng, cháy, đi với kim mộc đề phòng bị đánh, đi với thổ phải đề phòng đổ sập, ôn dịch” ',('• “Gặp quan tinh ấn thụ là tốt” ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,24,20,24), 1.0),0.0))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f246(E, r, c):
    # Ô C120
    # = IF(AND(TRUE(),AND(0<>0,TRUNG_GIAN!$C$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRUNG
    #   _GIAN!$C$8=0))),3,-3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(0.0,0.0), X.eq(E.v(23,6,3),X.MOD(X.add(X.add(E.v(23,6,3),(3.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(3.0))),12.0), 12.0))))) else '')

def f247(E, r, c):
    # Ô D120
    # = IF(AND(TRUE(),AND(1<>0,TRUNG_GIAN!$D$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRUNG
    #   _GIAN!$C$8=0))),3,-3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(1.0,0.0), X.eq(E.v(23,6,4),X.MOD(X.add(X.add(E.v(23,6,3),(3.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(3.0))),12.0), 12.0))))) else '')

def f248(E, r, c):
    # Ô E120
    # = IF(AND(TRUE(),AND(2<>0,TRUNG_GIAN!$E$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRUNG
    #   _GIAN!$C$8=0))),3,-3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(2.0,0.0), X.eq(E.v(23,6,5),X.MOD(X.add(X.add(E.v(23,6,3),(3.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(3.0))),12.0), 12.0))))) else '')

def f249(E, r, c):
    # Ô F120
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3<>0,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP
    #   _LIEU!$C$22="Nam"),NOT(TRUNG_GIAN!$C$8=0))),3,-3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.ne(3.0,0.0), X.eq(E.v(23,6,6),X.MOD(X.add(X.add(E.v(23,6,3),(3.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(3.0))),12.0), 12.0))))) else '')

def f250(E, r, c):
    # Ô I120
    # = IF(H120="Có",TRIM("• “Nhiều cơ mưu, chủ về nắm các việc hình phạt…”; năm, vận đến đó “thường gặp chuyện cãi cọ, hình phạt” "),"")
    return (X.TRIM('• “Nhiều cơ mưu, chủ về nắm các việc hình phạt…”; năm, vận đến đó “thường gặp chuyện cãi cọ, hình phạt” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f251(E, r, c):
    # Ô C121
    # = IF(AND(TRUE(),AND(0<>0,TRUNG_GIAN!$C$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRUNG
    #   _GIAN!$C$8=0))),-3,3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(0.0,0.0), X.eq(E.v(23,6,3),X.MOD(X.add(X.add(E.v(23,6,3),(X.neg(3.0) if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else 3.0)),12.0), 12.0))))) else '')

def f252(E, r, c):
    # Ô D121
    # = IF(AND(TRUE(),AND(1<>0,TRUNG_GIAN!$D$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRUNG
    #   _GIAN!$C$8=0))),-3,3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(1.0,0.0), X.eq(E.v(23,6,4),X.MOD(X.add(X.add(E.v(23,6,3),(X.neg(3.0) if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else 3.0)),12.0), 12.0))))) else '')

def f253(E, r, c):
    # Ô E121
    # = IF(AND(TRUE(),AND(2<>0,TRUNG_GIAN!$E$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRUNG
    #   _GIAN!$C$8=0))),-3,3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(2.0,0.0), X.eq(E.v(23,6,5),X.MOD(X.add(X.add(E.v(23,6,3),(X.neg(3.0) if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else 3.0)),12.0), 12.0))))) else '')

def f254(E, r, c):
    # Ô F121
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3<>0,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$C$6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP
    #   _LIEU!$C$22="Nam"),NOT(TRUNG_GIAN!$C$8=0))),-3,3)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.ne(3.0,0.0), X.eq(E.v(23,6,6),X.MOD(X.add(X.add(E.v(23,6,3),(X.neg(3.0) if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else 3.0)),12.0), 12.0))))) else '')

def f255(E, r, c):
    # Ô I121
    # = IF(H121="Có",TRIM("• “Cả hai (câu, giảo) đều gặp càng nặng, gặp một thì nhẹ hơn, đi với quan sát thì nặng” "),"")
    return (X.TRIM('• “Cả hai (câu, giảo) đều gặp càng nặng, gặp một thì nhẹ hơn, đi với quan sát thì nặng” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f256(E, r, c):
    # Ô C122
    # = IF(AND(TRUE(),AND(SUM(TRUNG_GIAN!$C$60:$F$60)>0,OR(AND(0<>0,TRUNG_GIAN!$C$6=CHOOSE(TRUNG_GIAN!$C$6+1,11,8,5,2,11,8,5,2,11,8,5,2)),AND(0<>2,TRUNG_GIAN!
    #   $C$6=CHOOSE(TRUNG_GIAN!$E$6+1,11,8,5,2,11,8,5,2,11,8,5,2))))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.gt(X.SUM(E.rg(23,60,3,60,6)),0.0), X.OR(X.AND(X.ne(0.0,0.0), X.eq(E.v(23,6,3),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))), X.AND(X.ne(0.0,2.0), X.eq(E.v(23,6,3),X.CHOOSE(X.add(E.v(23,6,5),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))))))) else '')

def f257(E, r, c):
    # Ô D122
    # = IF(AND(TRUE(),AND(SUM(TRUNG_GIAN!$C$60:$F$60)>0,OR(AND(1<>0,TRUNG_GIAN!$D$6=CHOOSE(TRUNG_GIAN!$C$6+1,11,8,5,2,11,8,5,2,11,8,5,2)),AND(1<>2,TRUNG_GIAN!
    #   $D$6=CHOOSE(TRUNG_GIAN!$E$6+1,11,8,5,2,11,8,5,2,11,8,5,2))))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.gt(X.SUM(E.rg(23,60,3,60,6)),0.0), X.OR(X.AND(X.ne(1.0,0.0), X.eq(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))), X.AND(X.ne(1.0,2.0), X.eq(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,6,5),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))))))) else '')

def f258(E, r, c):
    # Ô E122
    # = IF(AND(TRUE(),AND(SUM(TRUNG_GIAN!$C$60:$F$60)>0,OR(AND(2<>0,TRUNG_GIAN!$E$6=CHOOSE(TRUNG_GIAN!$C$6+1,11,8,5,2,11,8,5,2,11,8,5,2)),AND(2<>2,TRUNG_GIAN!
    #   $E$6=CHOOSE(TRUNG_GIAN!$E$6+1,11,8,5,2,11,8,5,2,11,8,5,2))))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.gt(X.SUM(E.rg(23,60,3,60,6)),0.0), X.OR(X.AND(X.ne(2.0,0.0), X.eq(E.v(23,6,5),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))), X.AND(X.ne(2.0,2.0), X.eq(E.v(23,6,5),X.CHOOSE(X.add(E.v(23,6,5),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))))))) else '')

def f259(E, r, c):
    # Ô F122
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(SUM(TRUNG_GIAN!$C$60:$F$60)>0,OR(AND(3<>0,TRUNG_GIAN!$F$6=CHOOSE(TRUNG_GIAN!$C$6+1,11,8,5,2,11,8,5,2,11,
    #   8,5,2)),AND(3<>2,TRUNG_GIAN!$F$6=CHOOSE(TRUNG_GIAN!$E$6+1,11,8,5,2,11,8,5,2,11,8,5,2))))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.gt(X.SUM(E.rg(23,60,3,60,6)),0.0), X.OR(X.AND(X.ne(3.0,0.0), X.eq(E.v(23,6,6),X.CHOOSE(X.add(E.v(23,6,3),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))), X.AND(X.ne(3.0,2.0), X.eq(E.v(23,6,6),X.CHOOSE(X.add(E.v(23,6,5),1.0), lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0, lambda: 8.0, lambda: 5.0, lambda: 2.0))))))) else '')

def f260(E, r, c):
    # Ô I122
    # = IF(H122="Có",TRIM("• Có hợp cục: “gặp tài cục là tài vượng, gặp quan cục là quan vượng … thân nhược là hại, gặp ấn cục là tốt” "),"")
    return (X.TRIM('• Có hợp cục: “gặp tài cục là tài vượng, gặp quan cục là quan vượng … thân nhược là hại, gặp ấn cục là tốt” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f261(E, r, c):
    # Ô C123
    # = IF(AND(TRUE(),AND(0<>0,TRUNG_GIAN!$C$6=MOD(TRUNG_GIAN!$C$6+6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRU
    #   NG_GIAN!$C$8=0))),1,-1)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(0.0,0.0), X.eq(E.v(23,6,3),X.MOD(X.add(X.add(X.add(E.v(23,6,3),6.0),(1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(1.0))),12.0), 12.0))))) else '')

def f262(E, r, c):
    # Ô D123
    # = IF(AND(TRUE(),AND(1<>0,TRUNG_GIAN!$D$6=MOD(TRUNG_GIAN!$C$6+6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRU
    #   NG_GIAN!$C$8=0))),1,-1)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(1.0,0.0), X.eq(E.v(23,6,4),X.MOD(X.add(X.add(X.add(E.v(23,6,3),6.0),(1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(1.0))),12.0), 12.0))))) else '')

def f263(E, r, c):
    # Ô E123
    # = IF(AND(TRUE(),AND(2<>0,TRUNG_GIAN!$E$6=MOD(TRUNG_GIAN!$C$6+6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NHAP_LIEU!$C$22="Nam"),NOT(TRU
    #   NG_GIAN!$C$8=0))),1,-1)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(2.0,0.0), X.eq(E.v(23,6,5),X.MOD(X.add(X.add(X.add(E.v(23,6,3),6.0),(1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(1.0))),12.0), 12.0))))) else '')

def f264(E, r, c):
    # Ô F123
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3<>0,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$C$6+6+IF(OR(AND(NHAP_LIEU!$C$22="Nam",TRUNG_GIAN!$C$8=0),AND(NOT(NH
    #   AP_LIEU!$C$22="Nam"),NOT(TRUNG_GIAN!$C$8=0))),1,-1)+12,12))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.ne(3.0,0.0), X.eq(E.v(23,6,6),X.MOD(X.add(X.add(X.add(E.v(23,6,3),6.0),(1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.eq(E.v(23,8,3),0.0))))) else X.neg(1.0))),12.0), 12.0))))) else '')

def f265(E, r, c):
    # Ô I123
    # = IF(H123="Có",TRIM("• Tuế vận gặp “không được bình yên, không có bệnh trong thì gặp nạn ngoài” "&IF((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,
    #   1)>0),"• “Đã có thất sát làm hại, còn gặp phải nguyên thần thì hại càng nặng hơn” ","")),"")
    return (X.TRIM(X.cat('• Tuế vận gặp “không được bình yên, không có bệnh trong thì gặp nạn ngoài” ',('• “Đã có thất sát làm hại, còn gặp phải nguyên thần thì hại càng nặng hơn” ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0),0.0)) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f266(E, r, c):
    # Ô C124
    # = IF(AND(TRUE(),AND(OR(0=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|","|4|17|104|32|47|76|58|119|89|61|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), '|4|17|104|32|47|76|58|119|89|61|'))))) else '')

def f267(E, r, c):
    # Ô D124
    # = IF(AND(TRUE(),AND(OR(1=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|","|4|17|104|32|47|76|58|119|89|61|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), '|4|17|104|32|47|76|58|119|89|61|'))))) else '')

def f268(E, r, c):
    # Ô E124
    # = IF(AND(TRUE(),AND(OR(2=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|4|17|104|32|47|76|58|119|89|61|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|4|17|104|32|47|76|58|119|89|61|'))))) else '')

def f269(E, r, c):
    # Ô F124
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=2),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|","|4|17|104|32|47|76|58|119|89|61|")))
    #   ),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,2.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), '|4|17|104|32|47|76|58|119|89|61|'))))) else '')

def f270(E, r, c):
    # Ô I124
    # = IF(H124="Có",TRIM(IF(NOT((OR(THAN_SAT!$H$85="Có",THAN_SAT!$H$86="Có"))),"• “Kho vàng bạc hóa thành đất bụi” ","")&IF((OR(THAN_SAT!$H$85="Có",THAN_SAT!
    #   $H$86="Có")),"• “Nếu gặp thiên đức, nguyệt đức thì không còn là điều kị nữa” ","")),"")
    return (X.TRIM(X.cat(('• “Kho vàng bạc hóa thành đất bụi” ' if X.truthy(X.NOT(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có')))) else ''),('• “Nếu gặp thiên đức, nguyệt đức thì không còn là điều kị nữa” ' if X.truthy(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f271(E, r, c):
    # Ô C125
    # = IF(AND(TRUE(),AND(OR(0=2,0=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|","|17|41|95|56|98|54|96|30|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,2.0), X.eq(0.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), '|17|41|95|56|98|54|96|30|'))))) else '')

def f272(E, r, c):
    # Ô D125
    # = IF(AND(TRUE(),AND(OR(1=2,1=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|","|17|41|95|56|98|54|96|30|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,2.0), X.eq(1.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), '|17|41|95|56|98|54|96|30|'))))) else '')

def f273(E, r, c):
    # Ô E125
    # = IF(AND(TRUE(),AND(OR(2=2,2=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|17|41|95|56|98|54|96|30|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,2.0), X.eq(2.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|17|41|95|56|98|54|96|30|'))))) else '')

def f274(E, r, c):
    # Ô F125
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=2,3=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|","|17|41|95|56|98|54|96|30|")))),"
    #   có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,2.0), X.eq(3.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), '|17|41|95|56|98|54|96|30|'))))) else '')

def f275(E, r, c):
    # Ô I125
    # = IF(H125="Có",TRIM("• “Hôn nhân không thuận. Nam khắc vợ, nữ khắc chồng” (sách ghi chỉ để tham khảo; giả định G33: ngày hoặc giờ) "),"")
    return (X.TRIM('• “Hôn nhân không thuận. Nam khắc vợ, nữ khắc chồng” (sách ghi chỉ để tham khảo; giả định G33: ngày hoặc giờ) ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f276(E, r, c):
    # Ô C126
    # = IF(AND(TRUE(),AND(OR(0=1,0=2,0=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")),ISNUMBER(FI
    #   ND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(0.0,1.0), X.eq(0.0,2.0), X.eq(0.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|'))))) else '')

def f277(E, r, c):
    # Ô D126
    # = IF(AND(TRUE(),AND(OR(1=1,1=2,1=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")),ISNUMBER(FI
    #   ND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(1.0,1.0), X.eq(1.0,2.0), X.eq(1.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|'))))) else '')

def f278(E, r, c):
    # Ô E126
    # = IF(AND(TRUE(),AND(OR(2=1,2=2,2=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")),ISNUMBER(FI
    #   ND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.OR(X.eq(2.0,1.0), X.eq(2.0,2.0), X.eq(2.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|'))))) else '')

def f279(E, r, c):
    # Ô F126
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(OR(3=1,3=2,3=3),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|","|24|37|50|87|100|113|30|43|5
    #   6|93|106|119|")),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|","|24|37|50|87|100|113|30|43|56|93|106|119|")))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.OR(X.eq(3.0,1.0), X.eq(3.0,2.0), X.eq(3.0,3.0)), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), '|24|37|50|87|100|113|30|43|56|93|106|119|'))))) else '')

def f280(E, r, c):
    # Ô I126
    # = IF(H126="Có",TRIM(IF(NHAP_LIEU!$C$22="Nam","• Nam: “bị nhà vợ từ chối, hoặc cãi cọ mâu thuẫn với nhà vợ” ","")&IF(NOT(NHAP_LIEU!$C$22="Nam"),"• Nữ: “c
    #   hị em dâu không toàn vẹn, hoặc bị nhà chồng từ chối” ","")&"• “Tiêu chí thông tin rõ nhất về hôn nhân không thuận” "),"")
    return (X.TRIM(X.cat(X.cat(('• Nam: “bị nhà vợ từ chối, hoặc cãi cọ mâu thuẫn với nhà vợ” ' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else ''),('• Nữ: “chị em dâu không toàn vẹn, hoặc bị nhà chồng từ chối” ' if X.truthy(X.NOT(X.eq(E.v(2,22,3),'Nam'))) else '')),'• “Tiêu chí thông tin rõ nhất về hôn nhân không thuận” ')) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f281(E, r, c):
    # Ô C127
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&(TRUNG_GIAN!$C$5*12+TRUNG_GIAN!$C$6)&"|",CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,"|80|93|","|96|119|","|2|15|","|30
    #   |41|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,3),12.0),E.v(23,6,3))),'|'), X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: '|80|93|', lambda: '|96|119|', lambda: '|2|15|', lambda: '|30|41|'))))) else '')

def f282(E, r, c):
    # Ô D127
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&(TRUNG_GIAN!$D$5*12+TRUNG_GIAN!$D$6)&"|",CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,"|80|93|","|96|119|","|2|15|","|30
    #   |41|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,4),12.0),E.v(23,6,4))),'|'), X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: '|80|93|', lambda: '|96|119|', lambda: '|2|15|', lambda: '|30|41|'))))) else '')

def f283(E, r, c):
    # Ô E127
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&(TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$E$6)&"|",CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,"|80|93|","|96|119|","|2|15|","|30
    #   |41|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,5))),'|'), X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: '|80|93|', lambda: '|96|119|', lambda: '|2|15|', lambda: '|30|41|'))))) else '')

def f284(E, r, c):
    # Ô F127
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(FIND("|"&(TRUNG_GIAN!$F$5*12+TRUNG_GIAN!$F$6)&"|",CHOOSE(INT(MOD(TRUNG_GIAN!$D$6+10,12)/3)+1,"|80|9
    #   3|","|96|119|","|2|15|","|30|41|")))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',X.add(X.mul(E.v(23,5,6),12.0),E.v(23,6,6))),'|'), X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)),1.0), lambda: '|80|93|', lambda: '|96|119|', lambda: '|2|15|', lambda: '|30|41|'))))) else '')

def f285(E, r, c):
    # Ô I127
    # = IF(H127="Có",TRIM("• “Mọi việc không thành, có đầu không có đuôi”; “chủ về thân yếu, nhiều bệnh” "&IF(LEN($E$127)>0,"• Ở trụ ngày: “đặc biệt là nhật n
    #   guyên thì càng không tốt” ","")),"")
    return (X.TRIM(X.cat('• “Mọi việc không thành, có đầu không có đuôi”; “chủ về thân yếu, nhiều bệnh” ',('• Ở trụ ngày: “đặc biệt là nhật nguyên thì càng không tốt” ' if X.truthy(X.gt(X.LEN(E.v(12,127,5)),0.0)) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f286(E, r, c):
    # Ô C128
    # = IF(AND(TRUE(),AND(0<>2,THAN_SAT!$C$13=1)),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(0.0,2.0), X.eq(E.v(12,13,3),1.0)))) else '')

def f287(E, r, c):
    # Ô D128
    # = IF(AND(TRUE(),AND(1<>2,THAN_SAT!$D$13=1)),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(1.0,2.0), X.eq(E.v(12,13,4),1.0)))) else '')

def f288(E, r, c):
    # Ô E128
    # = IF(AND(TRUE(),AND(2<>2,THAN_SAT!$E$13=1)),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.ne(2.0,2.0), X.eq(E.v(12,13,5),1.0)))) else '')

def f289(E, r, c):
    # Ô F128
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3<>2,THAN_SAT!$F$13=1)),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.ne(3.0,2.0), X.eq(E.v(12,13,6),1.0)))) else '')

def f290(E, r, c):
    # Ô I128
    # = IF(H128="Có",TRIM(IF(LEN($C$128)>0,"• Ở chi năm: “tổ nghiệp không có gì … mẹ … có mẹ như không” ","")&IF(LEN($D$128)>0,"• Ở chi tháng: “không có anh c
    #   hị em hoặc có … nhưng không nương tựa được” ","")&IF(LEN($F$128)>0,"• Ở chi giờ: “sau khi kết hôn không thể có con ngay … hoặc có con mà không nương t
    #   ựa được” ","")),"")
    return (X.TRIM(X.cat(X.cat(('• Ở chi năm: “tổ nghiệp không có gì … mẹ … có mẹ như không” ' if X.truthy(X.gt(X.LEN(E.v(12,128,3)),0.0)) else ''),('• Ở chi tháng: “không có anh chị em hoặc có … nhưng không nương tựa được” ' if X.truthy(X.gt(X.LEN(E.v(12,128,4)),0.0)) else '')),('• Ở chi giờ: “sau khi kết hôn không thể có con ngay … hoặc có con mà không nương tựa được” ' if X.truthy(X.gt(X.LEN(E.v(12,128,6)),0.0)) else ''))) if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f291(E, r, c):
    # Ô C129
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|","|8|","|5|","|11|","|5|","|2|","|8|","|5|","|11|","|5|")))),IF(THAN
    #   _SAT!$C$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,3),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|', lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|'))))) else '')

def f292(E, r, c):
    # Ô D129
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|","|8|","|5|","|11|","|5|","|2|","|8|","|5|","|11|","|5|")))),IF(THAN
    #   _SAT!$D$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,4),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|', lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|'))))) else '')

def f293(E, r, c):
    # Ô E129
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|","|8|","|5|","|11|","|5|","|2|","|8|","|5|","|11|","|5|")))),IF(THAN
    #   _SAT!$E$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,5),1.0)) else 'có') if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|', lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|'))))) else '')

def f294(E, r, c):
    # Ô F129
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$C$5+1,"|2|","|8|","|5|","|11|","|5|","|2|","|8|","|
    #   5|","|11|","|5|")))),IF(THAN_SAT!$F$11=1,"có – vô hiệu","có"),"")
    return (('có – vô hiệu' if X.truthy(X.eq(E.v(12,11,6),1.0)) else 'có') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|', lambda: '|2|', lambda: '|8|', lambda: '|5|', lambda: '|11|', lambda: '|5|'))))) else '')

def f295(E, r, c):
    # Ô I129
    # = IF(H129="Có",TRIM("• [Tham khảo] “Đức thần là thần bảo hộ, là điềm lành, có thể hóa hung thành cát” (ngũ hành phải sinh vượng mới được hưởng phúc) "),
    #   "")
    return (X.TRIM('• [Tham khảo] “Đức thần là thần bảo hộ, là điềm lành, có thể hóa hung thành cát” (ngũ hành phải sinh vượng mới được hưởng phúc) ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f296(E, r, c):
    # Ô C130
    # = IF(AND(TRUE(),AND(0=3,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$E$6+2,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(0.0,3.0), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(23,6,5),2.0), 12.0))))) else '')

def f297(E, r, c):
    # Ô D130
    # = IF(AND(TRUE(),AND(1=3,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$E$6+2,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(1.0,3.0), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(23,6,5),2.0), 12.0))))) else '')

def f298(E, r, c):
    # Ô E130
    # = IF(AND(TRUE(),AND(2=3,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$E$6+2,12))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(2.0,3.0), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(23,6,5),2.0), 12.0))))) else '')

def f299(E, r, c):
    # Ô F130
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(3=3,TRUNG_GIAN!$F$6=MOD(TRUNG_GIAN!$E$6+2,12))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.eq(3.0,3.0), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(23,6,5),2.0), 12.0))))) else '')

def f300(E, r, c):
    # Ô I130
    # = IF(H130="Có",TRIM("• [Tham khảo] “Cách giác chủ về tai họa ngục tù, cốt nhục chia lìa” "),"")
    return (X.TRIM('• [Tham khảo] “Cách giác chủ về tai họa ngục tù, cốt nhục chia lìa” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f301(E, r, c):
    # Ô C131
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$C$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|","|5|","|3|","|8|","|3|","|8|","|6|","|11|","|9|","|2|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,3)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|', lambda: '|5|', lambda: '|3|', lambda: '|8|', lambda: '|3|', lambda: '|8|', lambda: '|6|', lambda: '|11|', lambda: '|9|', lambda: '|2|'))))) else '')

def f302(E, r, c):
    # Ô D131
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$D$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|","|5|","|3|","|8|","|3|","|8|","|6|","|11|","|9|","|2|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,4)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|', lambda: '|5|', lambda: '|3|', lambda: '|8|', lambda: '|3|', lambda: '|8|', lambda: '|6|', lambda: '|11|', lambda: '|9|', lambda: '|2|'))))) else '')

def f303(E, r, c):
    # Ô E131
    # = IF(AND(TRUE(),ISNUMBER(FIND("|"&TRUNG_GIAN!$E$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|","|5|","|3|","|8|","|3|","|8|","|6|","|11|","|9|","|2|")))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,5)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|', lambda: '|5|', lambda: '|3|', lambda: '|8|', lambda: '|3|', lambda: '|8|', lambda: '|6|', lambda: '|11|', lambda: '|9|', lambda: '|2|'))))) else '')

def f304(E, r, c):
    # Ô F131
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(FIND("|"&TRUNG_GIAN!$F$6&"|",CHOOSE(TRUNG_GIAN!$E$5+1,"|0|","|5|","|3|","|8|","|3|","|8|","|6|","|1
    #   1|","|9|","|2|")))),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.FIND(X.cat(X.cat('|',E.v(23,6,6)),'|'), X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '|0|', lambda: '|5|', lambda: '|3|', lambda: '|8|', lambda: '|3|', lambda: '|8|', lambda: '|6|', lambda: '|11|', lambda: '|9|', lambda: '|2|'))))) else '')

def f305(E, r, c):
    # Ô I131
    # = IF(H131="Có",TRIM("• [Tham khảo] “Chủ về sự hoan lạc nam nữ, dễ hấp dẫn người khác giới cũng như dễ bị lôi cuốn” "),"")
    return (X.TRIM('• [Tham khảo] “Chủ về sự hoan lạc nam nữ, dễ hấp dẫn người khác giới cũng như dễ bị lôi cuốn” ') if X.truthy(X.eq(E.v(12,r,c-1),'Có')) else '')

def f306(E, r, c):
    # Ô C133, D133, E133, F133
    # = IF(LEN(L133)>1,"có","")
    return ('có' if X.truthy(X.gt(X.LEN(E.v(12,r,c+9)),1.0)) else '')

def f307(E, r, c):
    # Ô I133
    # = TRIM(IF(LEN(L133)>1,"• Năm – "&L133&" ","")&IF(LEN(M133)>1,"• Tháng – "&M133&" ","")&IF(LEN(N133)>1,"• Ngày – "&N133&" ","")&IF(LEN(O133)>1,"• Giờ – "
    #   &O133&" ",""))
    return X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('• Năm – ',E.v(12,r,c+3)),' ') if X.truthy(X.gt(X.LEN(E.v(12,r,c+3)),1.0)) else ''),(X.cat(X.cat('• Tháng – ',E.v(12,r,c+4)),' ') if X.truthy(X.gt(X.LEN(E.v(12,r,c+4)),1.0)) else '')),(X.cat(X.cat('• Ngày – ',E.v(12,r,c+5)),' ') if X.truthy(X.gt(X.LEN(E.v(12,r,c+5)),1.0)) else '')),(X.cat(X.cat('• Giờ – ',E.v(12,r,c+6)),' ') if X.truthy(X.gt(X.LEN(E.v(12,r,c+6)),1.0)) else '')))

def f308(E, r, c):
    # Ô L133
    # = IF(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)),CHOOSE(TRUNG_GIAN!$E$5+1,CHOOSE(TRUNG_GIAN!$C$5+1,"Giáp Dần: “lộc trường
    #    sinh, đại cát”","-","Bính Dần: “lộc phúc tinh” – tốt","-","Mậu Dần: “lộc phục mã” – tốt","-","Canh Dần: “lộc phá, nửa tốt nửa xấu”","-","Nhâm Dần: “c
    #   hính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”","-"),CHOOSE(TRUNG_GIAN!$C$5+1,"-","Ất Mão: “lộc hỉ thần vượng, chỉ về cát”","-","Đinh 
    #   Mão: “cắt ngang đường, chủ về hung”","-","Kỷ Mão: “lộc tiến thần”","-","Tân Mão: “lộc phá … nửa tốt nửa xấu”","-","Quý Mão: kèm thiên ất “là lộc chết,
    #    tuy sang nhưng nghèo suốt đời”"),CHOOSE(TRUNG_GIAN!$C$5+1,"-","Ất Tị: “lộc mã” – tốt","-","Đinh Tị: “lộc khố” – tốt","-","Kỷ Tị: “lộc kho trời, chủ v
    #   ề cát”","-","Tân Tị: “đứt đường, coi như bị tuần không”","-","Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”"),CHOOSE(TRUNG_GIAN!$C$5+1,"Giáp Ngọ: “lộc 
    #   tiến thần” – tốt","-","Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu","-","Mậu Ngọ: “gặp kình dương, phần nhiều xấu”","-","Canh Ng
    #   ọ: “đứt đường giống như không vong, xấu”","-","Nhâm Ngọ: “lộc đức hợp” – tốt","-"),CHOOSE(TRUNG_GIAN!$C$5+1,"-","Ất Tị: “lộc đồng hương” – tốt","-","Đ
    #   inh Tị: “lộc kho” – tốt","-","Kỷ Tị: “lộc kho trời, tốt”","-","Tân Tị: “đứt đường, tuần không, vong”","-","Quý Tị: “lộc quý thần, hóa hợp với mậu quý 
    #   là được chức quan quan trọng”"),CHOOSE(TRUNG_GIAN!$C$5+1,"Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”","-","Bính Ngọ: “mừng gặp được lộc thần”","
    #   -","Mậu Ngọ: “lộc phục thần kình dương, xấu”","-","Canh Ngọ: “đứt đường, tuần không, vong” – xấu","-","Nhâm Ngọ: “lộc tử quý” – xấu","-"),CHOOSE(TRUNG
    #   _GIAN!$C$5+1,"Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu","-","Bính Thân: “lộc đại bại”","-","Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc ti
    #   nh quý nhân thì tốt”","-","Canh Thân: “lộc trường sinh, đại cát”","-","Nhâm Thân: “lộc đại bại” – xấu","-"),CHOOSE(TRUNG_GIAN!$C$5+1,"-","Ất Dậu: “lộc
    #    bị phá, phần nhiều thất bại”","-","Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”","-","Kỷ Dậu: “lộc tiến thần” – tốt"
    #   ,"-","Tân Dậu: “chính lộc” – tốt","-","Quý Dậu: “lộc phục thần, xấu”"),CHOOSE(TRUNG_GIAN!$C$5+1,"-","Ất Hợi: “lộc thiên đức” – tốt","-","Đinh Hợi: “lộ
    #   c hợp quý thần” – tốt","-","Kỷ Hợi: “lộc vượng” – tốt","-","Tân Hợi: “lộc cùng với mã” – tốt","-","Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”"),C
    #   HOOSE(TRUNG_GIAN!$C$5+1,"Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”","-","Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền
    #   ”","-","Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”","-","Canh Tí: “ấn lộc là cát”","-","Nhâm Tí: “chính lộc kình dương, xấu”","-")),"")
    return (X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: 'Giáp Dần: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Bính Dần: “lộc phúc tinh” – tốt', lambda: '-', lambda: 'Mậu Dần: “lộc phục mã” – tốt', lambda: '-', lambda: 'Canh Dần: “lộc phá, nửa tốt nửa xấu”', lambda: '-', lambda: 'Nhâm Dần: “chính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '-', lambda: 'Ất Mão: “lộc hỉ thần vượng, chỉ về cát”', lambda: '-', lambda: 'Đinh Mão: “cắt ngang đường, chủ về hung”', lambda: '-', lambda: 'Kỷ Mão: “lộc tiến thần”', lambda: '-', lambda: 'Tân Mão: “lộc phá … nửa tốt nửa xấu”', lambda: '-', lambda: 'Quý Mão: kèm thiên ất “là lộc chết, tuy sang nhưng nghèo suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '-', lambda: 'Ất Tị: “lộc mã” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc khố” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, chủ về cát”', lambda: '-', lambda: 'Tân Tị: “đứt đường, coi như bị tuần không”', lambda: '-', lambda: 'Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: 'Giáp Ngọ: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu', lambda: '-', lambda: 'Mậu Ngọ: “gặp kình dương, phần nhiều xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường giống như không vong, xấu”', lambda: '-', lambda: 'Nhâm Ngọ: “lộc đức hợp” – tốt', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '-', lambda: 'Ất Tị: “lộc đồng hương” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc kho” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, tốt”', lambda: '-', lambda: 'Tân Tị: “đứt đường, tuần không, vong”', lambda: '-', lambda: 'Quý Tị: “lộc quý thần, hóa hợp với mậu quý là được chức quan quan trọng”'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: 'Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp được lộc thần”', lambda: '-', lambda: 'Mậu Ngọ: “lộc phục thần kình dương, xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Nhâm Ngọ: “lộc tử quý” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: 'Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Bính Thân: “lộc đại bại”', lambda: '-', lambda: 'Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc tinh quý nhân thì tốt”', lambda: '-', lambda: 'Canh Thân: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Nhâm Thân: “lộc đại bại” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '-', lambda: 'Ất Dậu: “lộc bị phá, phần nhiều thất bại”', lambda: '-', lambda: 'Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”', lambda: '-', lambda: 'Kỷ Dậu: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Tân Dậu: “chính lộc” – tốt', lambda: '-', lambda: 'Quý Dậu: “lộc phục thần, xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: '-', lambda: 'Ất Hợi: “lộc thiên đức” – tốt', lambda: '-', lambda: 'Đinh Hợi: “lộc hợp quý thần” – tốt', lambda: '-', lambda: 'Kỷ Hợi: “lộc vượng” – tốt', lambda: '-', lambda: 'Tân Hợi: “lộc cùng với mã” – tốt', lambda: '-', lambda: 'Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,3),1.0), lambda: 'Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”', lambda: '-', lambda: 'Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền”', lambda: '-', lambda: 'Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”', lambda: '-', lambda: 'Canh Tí: “ấn lộc là cát”', lambda: '-', lambda: 'Nhâm Tí: “chính lộc kình dương, xấu”', lambda: '-')) if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0)))) else '')

def f309(E, r, c):
    # Ô M133
    # = IF(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)),CHOOSE(TRUNG_GIAN!$E$5+1,CHOOSE(TRUNG_GIAN!$D$5+1,"Giáp Dần: “lộc trường
    #    sinh, đại cát”","-","Bính Dần: “lộc phúc tinh” – tốt","-","Mậu Dần: “lộc phục mã” – tốt","-","Canh Dần: “lộc phá, nửa tốt nửa xấu”","-","Nhâm Dần: “c
    #   hính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”","-"),CHOOSE(TRUNG_GIAN!$D$5+1,"-","Ất Mão: “lộc hỉ thần vượng, chỉ về cát”","-","Đinh 
    #   Mão: “cắt ngang đường, chủ về hung”","-","Kỷ Mão: “lộc tiến thần”","-","Tân Mão: “lộc phá … nửa tốt nửa xấu”","-","Quý Mão: kèm thiên ất “là lộc chết,
    #    tuy sang nhưng nghèo suốt đời”"),CHOOSE(TRUNG_GIAN!$D$5+1,"-","Ất Tị: “lộc mã” – tốt","-","Đinh Tị: “lộc khố” – tốt","-","Kỷ Tị: “lộc kho trời, chủ v
    #   ề cát”","-","Tân Tị: “đứt đường, coi như bị tuần không”","-","Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”"),CHOOSE(TRUNG_GIAN!$D$5+1,"Giáp Ngọ: “lộc 
    #   tiến thần” – tốt","-","Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu","-","Mậu Ngọ: “gặp kình dương, phần nhiều xấu”","-","Canh Ng
    #   ọ: “đứt đường giống như không vong, xấu”","-","Nhâm Ngọ: “lộc đức hợp” – tốt","-"),CHOOSE(TRUNG_GIAN!$D$5+1,"-","Ất Tị: “lộc đồng hương” – tốt","-","Đ
    #   inh Tị: “lộc kho” – tốt","-","Kỷ Tị: “lộc kho trời, tốt”","-","Tân Tị: “đứt đường, tuần không, vong”","-","Quý Tị: “lộc quý thần, hóa hợp với mậu quý 
    #   là được chức quan quan trọng”"),CHOOSE(TRUNG_GIAN!$D$5+1,"Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”","-","Bính Ngọ: “mừng gặp được lộc thần”","
    #   -","Mậu Ngọ: “lộc phục thần kình dương, xấu”","-","Canh Ngọ: “đứt đường, tuần không, vong” – xấu","-","Nhâm Ngọ: “lộc tử quý” – xấu","-"),CHOOSE(TRUNG
    #   _GIAN!$D$5+1,"Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu","-","Bính Thân: “lộc đại bại”","-","Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc ti
    #   nh quý nhân thì tốt”","-","Canh Thân: “lộc trường sinh, đại cát”","-","Nhâm Thân: “lộc đại bại” – xấu","-"),CHOOSE(TRUNG_GIAN!$D$5+1,"-","Ất Dậu: “lộc
    #    bị phá, phần nhiều thất bại”","-","Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”","-","Kỷ Dậu: “lộc tiến thần” – tốt"
    #   ,"-","Tân Dậu: “chính lộc” – tốt","-","Quý Dậu: “lộc phục thần, xấu”"),CHOOSE(TRUNG_GIAN!$D$5+1,"-","Ất Hợi: “lộc thiên đức” – tốt","-","Đinh Hợi: “lộ
    #   c hợp quý thần” – tốt","-","Kỷ Hợi: “lộc vượng” – tốt","-","Tân Hợi: “lộc cùng với mã” – tốt","-","Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”"),C
    #   HOOSE(TRUNG_GIAN!$D$5+1,"Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”","-","Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền
    #   ”","-","Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”","-","Canh Tí: “ấn lộc là cát”","-","Nhâm Tí: “chính lộc kình dương, xấu”","-")),"")
    return (X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 'Giáp Dần: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Bính Dần: “lộc phúc tinh” – tốt', lambda: '-', lambda: 'Mậu Dần: “lộc phục mã” – tốt', lambda: '-', lambda: 'Canh Dần: “lộc phá, nửa tốt nửa xấu”', lambda: '-', lambda: 'Nhâm Dần: “chính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: '-', lambda: 'Ất Mão: “lộc hỉ thần vượng, chỉ về cát”', lambda: '-', lambda: 'Đinh Mão: “cắt ngang đường, chủ về hung”', lambda: '-', lambda: 'Kỷ Mão: “lộc tiến thần”', lambda: '-', lambda: 'Tân Mão: “lộc phá … nửa tốt nửa xấu”', lambda: '-', lambda: 'Quý Mão: kèm thiên ất “là lộc chết, tuy sang nhưng nghèo suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: '-', lambda: 'Ất Tị: “lộc mã” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc khố” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, chủ về cát”', lambda: '-', lambda: 'Tân Tị: “đứt đường, coi như bị tuần không”', lambda: '-', lambda: 'Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 'Giáp Ngọ: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu', lambda: '-', lambda: 'Mậu Ngọ: “gặp kình dương, phần nhiều xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường giống như không vong, xấu”', lambda: '-', lambda: 'Nhâm Ngọ: “lộc đức hợp” – tốt', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: '-', lambda: 'Ất Tị: “lộc đồng hương” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc kho” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, tốt”', lambda: '-', lambda: 'Tân Tị: “đứt đường, tuần không, vong”', lambda: '-', lambda: 'Quý Tị: “lộc quý thần, hóa hợp với mậu quý là được chức quan quan trọng”'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 'Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp được lộc thần”', lambda: '-', lambda: 'Mậu Ngọ: “lộc phục thần kình dương, xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Nhâm Ngọ: “lộc tử quý” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 'Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Bính Thân: “lộc đại bại”', lambda: '-', lambda: 'Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc tinh quý nhân thì tốt”', lambda: '-', lambda: 'Canh Thân: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Nhâm Thân: “lộc đại bại” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: '-', lambda: 'Ất Dậu: “lộc bị phá, phần nhiều thất bại”', lambda: '-', lambda: 'Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”', lambda: '-', lambda: 'Kỷ Dậu: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Tân Dậu: “chính lộc” – tốt', lambda: '-', lambda: 'Quý Dậu: “lộc phục thần, xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: '-', lambda: 'Ất Hợi: “lộc thiên đức” – tốt', lambda: '-', lambda: 'Đinh Hợi: “lộc hợp quý thần” – tốt', lambda: '-', lambda: 'Kỷ Hợi: “lộc vượng” – tốt', lambda: '-', lambda: 'Tân Hợi: “lộc cùng với mã” – tốt', lambda: '-', lambda: 'Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 'Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”', lambda: '-', lambda: 'Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền”', lambda: '-', lambda: 'Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”', lambda: '-', lambda: 'Canh Tí: “ấn lộc là cát”', lambda: '-', lambda: 'Nhâm Tí: “chính lộc kình dương, xấu”', lambda: '-')) if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0)))) else '')

def f310(E, r, c):
    # Ô N133
    # = IF(AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)),CHOOSE(TRUNG_GIAN!$E$5+1,CHOOSE(TRUNG_GIAN!$E$5+1,"Giáp Dần: “lộc trường
    #    sinh, đại cát”","-","Bính Dần: “lộc phúc tinh” – tốt","-","Mậu Dần: “lộc phục mã” – tốt","-","Canh Dần: “lộc phá, nửa tốt nửa xấu”","-","Nhâm Dần: “c
    #   hính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”","-"),CHOOSE(TRUNG_GIAN!$E$5+1,"-","Ất Mão: “lộc hỉ thần vượng, chỉ về cát”","-","Đinh 
    #   Mão: “cắt ngang đường, chủ về hung”","-","Kỷ Mão: “lộc tiến thần”","-","Tân Mão: “lộc phá … nửa tốt nửa xấu”","-","Quý Mão: kèm thiên ất “là lộc chết,
    #    tuy sang nhưng nghèo suốt đời”"),CHOOSE(TRUNG_GIAN!$E$5+1,"-","Ất Tị: “lộc mã” – tốt","-","Đinh Tị: “lộc khố” – tốt","-","Kỷ Tị: “lộc kho trời, chủ v
    #   ề cát”","-","Tân Tị: “đứt đường, coi như bị tuần không”","-","Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”"),CHOOSE(TRUNG_GIAN!$E$5+1,"Giáp Ngọ: “lộc 
    #   tiến thần” – tốt","-","Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu","-","Mậu Ngọ: “gặp kình dương, phần nhiều xấu”","-","Canh Ng
    #   ọ: “đứt đường giống như không vong, xấu”","-","Nhâm Ngọ: “lộc đức hợp” – tốt","-"),CHOOSE(TRUNG_GIAN!$E$5+1,"-","Ất Tị: “lộc đồng hương” – tốt","-","Đ
    #   inh Tị: “lộc kho” – tốt","-","Kỷ Tị: “lộc kho trời, tốt”","-","Tân Tị: “đứt đường, tuần không, vong”","-","Quý Tị: “lộc quý thần, hóa hợp với mậu quý 
    #   là được chức quan quan trọng”"),CHOOSE(TRUNG_GIAN!$E$5+1,"Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”","-","Bính Ngọ: “mừng gặp được lộc thần”","
    #   -","Mậu Ngọ: “lộc phục thần kình dương, xấu”","-","Canh Ngọ: “đứt đường, tuần không, vong” – xấu","-","Nhâm Ngọ: “lộc tử quý” – xấu","-"),CHOOSE(TRUNG
    #   _GIAN!$E$5+1,"Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu","-","Bính Thân: “lộc đại bại”","-","Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc ti
    #   nh quý nhân thì tốt”","-","Canh Thân: “lộc trường sinh, đại cát”","-","Nhâm Thân: “lộc đại bại” – xấu","-"),CHOOSE(TRUNG_GIAN!$E$5+1,"-","Ất Dậu: “lộc
    #    bị phá, phần nhiều thất bại”","-","Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”","-","Kỷ Dậu: “lộc tiến thần” – tốt"
    #   ,"-","Tân Dậu: “chính lộc” – tốt","-","Quý Dậu: “lộc phục thần, xấu”"),CHOOSE(TRUNG_GIAN!$E$5+1,"-","Ất Hợi: “lộc thiên đức” – tốt","-","Đinh Hợi: “lộ
    #   c hợp quý thần” – tốt","-","Kỷ Hợi: “lộc vượng” – tốt","-","Tân Hợi: “lộc cùng với mã” – tốt","-","Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”"),C
    #   HOOSE(TRUNG_GIAN!$E$5+1,"Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”","-","Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền
    #   ”","-","Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”","-","Canh Tí: “ấn lộc là cát”","-","Nhâm Tí: “chính lộc kình dương, xấu”","-")),"")
    return (X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 'Giáp Dần: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Bính Dần: “lộc phúc tinh” – tốt', lambda: '-', lambda: 'Mậu Dần: “lộc phục mã” – tốt', lambda: '-', lambda: 'Canh Dần: “lộc phá, nửa tốt nửa xấu”', lambda: '-', lambda: 'Nhâm Dần: “chính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '-', lambda: 'Ất Mão: “lộc hỉ thần vượng, chỉ về cát”', lambda: '-', lambda: 'Đinh Mão: “cắt ngang đường, chủ về hung”', lambda: '-', lambda: 'Kỷ Mão: “lộc tiến thần”', lambda: '-', lambda: 'Tân Mão: “lộc phá … nửa tốt nửa xấu”', lambda: '-', lambda: 'Quý Mão: kèm thiên ất “là lộc chết, tuy sang nhưng nghèo suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '-', lambda: 'Ất Tị: “lộc mã” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc khố” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, chủ về cát”', lambda: '-', lambda: 'Tân Tị: “đứt đường, coi như bị tuần không”', lambda: '-', lambda: 'Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 'Giáp Ngọ: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu', lambda: '-', lambda: 'Mậu Ngọ: “gặp kình dương, phần nhiều xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường giống như không vong, xấu”', lambda: '-', lambda: 'Nhâm Ngọ: “lộc đức hợp” – tốt', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '-', lambda: 'Ất Tị: “lộc đồng hương” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc kho” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, tốt”', lambda: '-', lambda: 'Tân Tị: “đứt đường, tuần không, vong”', lambda: '-', lambda: 'Quý Tị: “lộc quý thần, hóa hợp với mậu quý là được chức quan quan trọng”'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 'Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp được lộc thần”', lambda: '-', lambda: 'Mậu Ngọ: “lộc phục thần kình dương, xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Nhâm Ngọ: “lộc tử quý” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 'Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Bính Thân: “lộc đại bại”', lambda: '-', lambda: 'Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc tinh quý nhân thì tốt”', lambda: '-', lambda: 'Canh Thân: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Nhâm Thân: “lộc đại bại” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '-', lambda: 'Ất Dậu: “lộc bị phá, phần nhiều thất bại”', lambda: '-', lambda: 'Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”', lambda: '-', lambda: 'Kỷ Dậu: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Tân Dậu: “chính lộc” – tốt', lambda: '-', lambda: 'Quý Dậu: “lộc phục thần, xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: '-', lambda: 'Ất Hợi: “lộc thiên đức” – tốt', lambda: '-', lambda: 'Đinh Hợi: “lộc hợp quý thần” – tốt', lambda: '-', lambda: 'Kỷ Hợi: “lộc vượng” – tốt', lambda: '-', lambda: 'Tân Hợi: “lộc cùng với mã” – tốt', lambda: '-', lambda: 'Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 'Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”', lambda: '-', lambda: 'Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền”', lambda: '-', lambda: 'Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”', lambda: '-', lambda: 'Canh Tí: “ấn lộc là cát”', lambda: '-', lambda: 'Nhâm Tí: “chính lộc kình dương, xấu”', lambda: '-')) if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0)))) else '')

def f311(E, r, c):
    # Ô O133
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(TRUNG_GIAN!$E$5+1,2,3,5,6,5,6,8,9,11,0)),CHOOSE(TRUNG_GIAN!$E$5+1,CHOOSE(TRUNG_GIAN!$
    #   F$5+1,"Giáp Dần: “lộc trường sinh, đại cát”","-","Bính Dần: “lộc phúc tinh” – tốt","-","Mậu Dần: “lộc phục mã” – tốt","-","Canh Dần: “lộc phá, nửa tốt
    #    nửa xấu”","-","Nhâm Dần: “chính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”","-"),CHOOSE(TRUNG_GIAN!$F$5+1,"-","Ất Mão: “lộc hỉ thần vư
    #   ợng, chỉ về cát”","-","Đinh Mão: “cắt ngang đường, chủ về hung”","-","Kỷ Mão: “lộc tiến thần”","-","Tân Mão: “lộc phá … nửa tốt nửa xấu”","-","Quý Mão
    #   : kèm thiên ất “là lộc chết, tuy sang nhưng nghèo suốt đời”"),CHOOSE(TRUNG_GIAN!$F$5+1,"-","Ất Tị: “lộc mã” – tốt","-","Đinh Tị: “lộc khố” – tốt","-",
    #   "Kỷ Tị: “lộc kho trời, chủ về cát”","-","Tân Tị: “đứt đường, coi như bị tuần không”","-","Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”"),CHOOSE(TRUNG_
    #   GIAN!$F$5+1,"Giáp Ngọ: “lộc tiến thần” – tốt","-","Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu","-","Mậu Ngọ: “gặp kình dương, p
    #   hần nhiều xấu”","-","Canh Ngọ: “đứt đường giống như không vong, xấu”","-","Nhâm Ngọ: “lộc đức hợp” – tốt","-"),CHOOSE(TRUNG_GIAN!$F$5+1,"-","Ất Tị: “l
    #   ộc đồng hương” – tốt","-","Đinh Tị: “lộc kho” – tốt","-","Kỷ Tị: “lộc kho trời, tốt”","-","Tân Tị: “đứt đường, tuần không, vong”","-","Quý Tị: “lộc qu
    #   ý thần, hóa hợp với mậu quý là được chức quan quan trọng”"),CHOOSE(TRUNG_GIAN!$F$5+1,"Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”","-","Bính Ngọ:
    #    “mừng gặp được lộc thần”","-","Mậu Ngọ: “lộc phục thần kình dương, xấu”","-","Canh Ngọ: “đứt đường, tuần không, vong” – xấu","-","Nhâm Ngọ: “lộc tử q
    #   uý” – xấu","-"),CHOOSE(TRUNG_GIAN!$F$5+1,"Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu","-","Bính Thân: “lộc đại bại”","-","Mậu Thân: “lộc phúc 
    #   mã, trì trệ; nếu gặp phúc tinh quý nhân thì tốt”","-","Canh Thân: “lộc trường sinh, đại cát”","-","Nhâm Thân: “lộc đại bại” – xấu","-"),CHOOSE(TRUNG_G
    #   IAN!$F$5+1,"-","Ất Dậu: “lộc bị phá, phần nhiều thất bại”","-","Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”","-","Kỷ
    #    Dậu: “lộc tiến thần” – tốt","-","Tân Dậu: “chính lộc” – tốt","-","Quý Dậu: “lộc phục thần, xấu”"),CHOOSE(TRUNG_GIAN!$F$5+1,"-","Ất Hợi: “lộc thiên đứ
    #   c” – tốt","-","Đinh Hợi: “lộc hợp quý thần” – tốt","-","Kỷ Hợi: “lộc vượng” – tốt","-","Tân Hợi: “lộc cùng với mã” – tốt","-","Quý Hợi: “lộc đại bại, 
    #   chủ về đói khó suốt đời”"),CHOOSE(TRUNG_GIAN!$F$5+1,"Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”","-","Bính Tí: “lộc kình dương, nếu có cả thầ
    #   n tinh quý nhân thì có quyền”","-","Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”","-","Canh Tí: “ấn lộc là cát”","-","Nhâm Tí: “chính lộc kình 
    #   dương, xấu”","-")),"")
    return (X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: 'Giáp Dần: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Bính Dần: “lộc phúc tinh” – tốt', lambda: '-', lambda: 'Mậu Dần: “lộc phục mã” – tốt', lambda: '-', lambda: 'Canh Dần: “lộc phá, nửa tốt nửa xấu”', lambda: '-', lambda: 'Nhâm Dần: “chính lộc”; kèm tuần không vong “có nhiều khả năng đi theo đạo”', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: '-', lambda: 'Ất Mão: “lộc hỉ thần vượng, chỉ về cát”', lambda: '-', lambda: 'Đinh Mão: “cắt ngang đường, chủ về hung”', lambda: '-', lambda: 'Kỷ Mão: “lộc tiến thần”', lambda: '-', lambda: 'Tân Mão: “lộc phá … nửa tốt nửa xấu”', lambda: '-', lambda: 'Quý Mão: kèm thiên ất “là lộc chết, tuy sang nhưng nghèo suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: '-', lambda: 'Ất Tị: “lộc mã” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc khố” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, chủ về cát”', lambda: '-', lambda: 'Tân Tị: “đứt đường, coi như bị tuần không”', lambda: '-', lambda: 'Quý Tị: “lộc phục quý thần, nửa tốt nửa xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: 'Giáp Ngọ: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp lộc thần”, gặp kình dương thì nửa tốt nửa xấu', lambda: '-', lambda: 'Mậu Ngọ: “gặp kình dương, phần nhiều xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường giống như không vong, xấu”', lambda: '-', lambda: 'Nhâm Ngọ: “lộc đức hợp” – tốt', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: '-', lambda: 'Ất Tị: “lộc đồng hương” – tốt', lambda: '-', lambda: 'Đinh Tị: “lộc kho” – tốt', lambda: '-', lambda: 'Kỷ Tị: “lộc kho trời, tốt”', lambda: '-', lambda: 'Tân Tị: “đứt đường, tuần không, vong”', lambda: '-', lambda: 'Quý Tị: “lộc quý thần, hóa hợp với mậu quý là được chức quan quan trọng”'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: 'Giáp Ngọ: “lộc hợp tiến thần, là tượng hiển đạt”', lambda: '-', lambda: 'Bính Ngọ: “mừng gặp được lộc thần”', lambda: '-', lambda: 'Mậu Ngọ: “lộc phục thần kình dương, xấu”', lambda: '-', lambda: 'Canh Ngọ: “đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Nhâm Ngọ: “lộc tử quý” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: 'Giáp Thân: “lộc đứt đường, tuần không, vong” – xấu', lambda: '-', lambda: 'Bính Thân: “lộc đại bại”', lambda: '-', lambda: 'Mậu Thân: “lộc phúc mã, trì trệ; nếu gặp phúc tinh quý nhân thì tốt”', lambda: '-', lambda: 'Canh Thân: “lộc trường sinh, đại cát”', lambda: '-', lambda: 'Nhâm Thân: “lộc đại bại” – xấu', lambda: '-'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: '-', lambda: 'Ất Dậu: “lộc bị phá, phần nhiều thất bại”', lambda: '-', lambda: 'Đinh Dậu: “lộc quý thần nhưng tuần không vong … nếu gặp được hỉ thần thì tốt”', lambda: '-', lambda: 'Kỷ Dậu: “lộc tiến thần” – tốt', lambda: '-', lambda: 'Tân Dậu: “chính lộc” – tốt', lambda: '-', lambda: 'Quý Dậu: “lộc phục thần, xấu”'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: '-', lambda: 'Ất Hợi: “lộc thiên đức” – tốt', lambda: '-', lambda: 'Đinh Hợi: “lộc hợp quý thần” – tốt', lambda: '-', lambda: 'Kỷ Hợi: “lộc vượng” – tốt', lambda: '-', lambda: 'Tân Hợi: “lộc cùng với mã” – tốt', lambda: '-', lambda: 'Quý Hợi: “lộc đại bại, chủ về đói khó suốt đời”'), lambda: X.CHOOSE(X.add(E.v(23,5,6),1.0), lambda: 'Giáp Tí: “lộc tiến thần, chủ về đăng khoa tiến đạt”', lambda: '-', lambda: 'Bính Tí: “lộc kình dương, nếu có cả thần tinh quý nhân thì có quyền”', lambda: '-', lambda: 'Mậu Tí: “lộc kình dương phục nấp … chỉ tốt một nửa”', lambda: '-', lambda: 'Canh Tí: “ấn lộc là cát”', lambda: '-', lambda: 'Nhâm Tí: “chính lộc kình dương, xấu”', lambda: '-')) if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 2.0, lambda: 3.0, lambda: 5.0, lambda: 6.0, lambda: 5.0, lambda: 6.0, lambda: 8.0, lambda: 9.0, lambda: 11.0, lambda: 0.0)))) else '')

def f312(E, r, c):
    # Ô C134
    # = TRIM(IF(H133="Có","▸ LỘC THEO CAN CHI: "&I133&" [P2 tr.308-309]"&CHAR(10),"")&IF(H102="Có","▸ THÁI CỰC QUÝ NHÂN ("&TRIM(IF(C102<>"","năm ","")&IF(D102
    #   <>"","tháng ","")&IF(E102<>"","ngày ","")&IF(F102<>"","giờ ",""))&"): "&I102&" [P2 tr.296-297]"&CHAR(10),"")&IF(H103="Có","▸ PHÚC TINH QUÝ NHÂN ("&TRI
    #   M(IF(C103<>"","năm ","")&IF(D103<>"","tháng ","")&IF(E103<>"","ngày ","")&IF(F103<>"","giờ ",""))&"): "&I103&" [P2 tr.300]"&CHAR(10),"")&IF(H104="Có",
    #   "▸ KHÔI CANH QUÝ NHÂN ("&TRIM(IF(C104<>"","năm ","")&IF(D104<>"","tháng ","")&IF(E104<>"","ngày ","")&IF(F104<>"","giờ ",""))&"): "&I104&" [P2 tr.301]
    #   "&CHAR(10),"")&IF(H105="Có","▸ QUỐC ẤN QUÝ NHÂN ("&TRIM(IF(C105<>"","năm ","")&IF(D105<>"","tháng ","")&IF(E105<>"","ngày ","")&IF(F105<>"","giờ ","")
    #   )&"): "&I105&" [P2 tr.302]"&CHAR(10),"")&IF(H106="Có","▸ HỌC ĐƯỜNG ("&TRIM(IF(C106<>"","năm ","")&IF(D106<>"","tháng ","")&IF(E106<>"","ngày ","")&IF(
    #   F106<>"","giờ ",""))&"): "&I106&" [P2 tr.302-303]"&CHAR(10),"")&IF(H107="Có","▸ TỪ QUÁN ("&TRIM(IF(C107<>"","năm ","")&IF(D107<>"","tháng ","")&IF(E10
    #   7<>"","ngày ","")&IF(F107<>"","giờ ",""))&"): "&I107&" [P2 tr.302-303]"&CHAR(10),"")&IF(H108="Có","▸ ĐỨC QUÝ NHÂN ("&TRIM(IF(C108<>"","năm ","")&IF(D1
    #   08<>"","tháng ","")&IF(E108<>"","ngày ","")&IF(F108<>"","giờ ",""))&"): "&I108&" [P2 tr.303]"&CHAR(10),"")&IF(H109="Có","▸ TÚ QUÝ NHÂN ("&TRIM(IF(C109
    #   <>"","năm ","")&IF(D109<>"","tháng ","")&IF(E109<>"","ngày ","")&IF(F109<>"","giờ ",""))&"): "&I109&" [P2 tr.303]"&CHAR(10),"")&IF(H110="Có","▸ KIM DƯ
    #    ("&TRIM(IF(C110<>"","năm ","")&IF(D110<>"","tháng ","")&IF(E110<>"","ngày ","")&IF(F110<>"","giờ ",""))&"): "&I110&" [P2 tr.306]"&CHAR(10),"")&IF(H11
    #   1="Có","▸ KIM THẦN ("&TRIM(IF(C111<>"","năm ","")&IF(D111<>"","tháng ","")&IF(E111<>"","ngày ","")&IF(F111<>"","giờ ",""))&"): "&I111&" [P2 tr.307]"&C
    #   HAR(10),"")&IF(H112="Có","▸ THIÊN Y ("&TRIM(IF(C112<>"","năm ","")&IF(D112<>"","tháng ","")&IF(E112<>"","ngày ","")&IF(F112<>"","giờ ",""))&"): "&I112
    #   &" [P2 tr.307-308]"&CHAR(10),"")&IF(H113="Có","▸ CỦNG LỘC ("&TRIM(IF(C113<>"","năm ","")&IF(D113<>"","tháng ","")&IF(E113<>"","ngày ","")&IF(F113<>"",
    #   "giờ ",""))&"): "&I113&" [P2 tr.310]"&CHAR(10),"")&IF(H114="Có","▸ GIÁP LỘC ("&TRIM(IF(C114<>"","năm ","")&IF(D114<>"","tháng ","")&IF(E114<>"","ngày 
    #   ","")&IF(F114<>"","giờ ",""))&"): "&I114&" [P2 tr.310]"&CHAR(10),"")&IF(H115="Có","▸ THIÊN XÁ ("&TRIM(IF(C115<>"","năm ","")&IF(D115<>"","tháng ","")&
    #   IF(E115<>"","ngày ","")&IF(F115<>"","giờ ",""))&"): "&I115&" [P2 tr.311]"&CHAR(10),"")&IF(H116="Có","▸ TAM KỲ ("&TRIM(IF(C116<>"","năm ","")&IF(D116<>
    #   "","tháng ","")&IF(E116<>"","ngày ","")&IF(F116<>"","giờ ",""))&"): "&I116&" [P2 tr.299-300]"&CHAR(10),"")&IF(H117="Có","▸ THIÊN LA (TUẤT – HỢI) ("&TR
    #   IM(IF(C117<>"","năm ","")&IF(D117<>"","tháng ","")&IF(E117<>"","ngày ","")&IF(F117<>"","giờ ",""))&"): "&I117&" [P2 tr.313-314]"&CHAR(10),"")&IF(H118=
    #   "Có","▸ ĐỊA VÕNG (THÌN – TỊ) ("&TRIM(IF(C118<>"","năm ","")&IF(D118<>"","tháng ","")&IF(E118<>"","ngày ","")&IF(F118<>"","giờ ",""))&"): "&I118&" [P2 
    #   tr.313-314]"&CHAR(10),"")&IF(H119="Có","▸ TAI SÁT ("&TRIM(IF(C119<>"","năm ","")&IF(D119<>"","tháng ","")&IF(E119<>"","ngày ","")&IF(F119<>"","giờ ","
    #   "))&"): "&I119&" [P2 tr.317-318]"&CHAR(10),"")&IF(H120="Có","▸ CÂU (CẤU) ("&TRIM(IF(C120<>"","năm ","")&IF(D120<>"","tháng ","")&IF(E120<>"","ngày ","
    #   ")&IF(F120<>"","giờ ",""))&"): "&I120&" [P2 tr.318]"&CHAR(10),"")&IF(H121="Có","▸ GIẢO ("&TRIM(IF(C121<>"","năm ","")&IF(D121<>"","tháng ","")&IF(E121
    #   <>"","ngày ","")&IF(F121<>"","giờ ",""))&"): "&I121&" [P2 tr.318]"&CHAR(10),"")&IF(H122="Có","▸ VONG THẦN ("&TRIM(IF(C122<>"","năm ","")&IF(D122<>"","
    #   tháng ","")&IF(E122<>"","ngày ","")&IF(F122<>"","giờ ",""))&"): "&I122&" [P2 tr.320]"&CHAR(10),"")&IF(H123="Có","▸ NGUYÊN THẦN (ĐẠI HAO) ("&TRIM(IF(C1
    #   23<>"","năm ","")&IF(D123<>"","tháng ","")&IF(E123<>"","ngày ","")&IF(F123<>"","giờ ",""))&"): "&I123&" [P2 tr.320-321]"&CHAR(10),"")&IF(H124="Có","▸ 
    #   THẬP ÁC ĐẠI BẠI ("&TRIM(IF(C124<>"","năm ","")&IF(D124<>"","tháng ","")&IF(E124<>"","ngày ","")&IF(F124<>"","giờ ",""))&"): "&I124&" [P2 tr.323-324]"&
    #   CHAR(10),"")&IF(H125="Có","▸ CÔ LOAN ("&TRIM(IF(C125<>"","năm ","")&IF(D125<>"","tháng ","")&IF(E125<>"","ngày ","")&IF(F125<>"","giờ ",""))&"): "&I12
    #   5&" [P2 tr.325]"&CHAR(10),"")&IF(H126="Có","▸ ÂM DƯƠNG LỆCH NHAU ("&TRIM(IF(C126<>"","năm ","")&IF(D126<>"","tháng ","")&IF(E126<>"","ngày ","")&IF(F1
    #   26<>"","giờ ",""))&"): "&I126&" [P2 tr.326]"&CHAR(10),"")&IF(H127="Có","▸ TỨ PHẾ ("&TRIM(IF(C127<>"","năm ","")&IF(D127<>"","tháng ","")&IF(E127<>"","
    #   ngày ","")&IF(F127<>"","giờ ",""))&"): "&I127&" [P2 tr.326]"&CHAR(10),"")&IF(H128="Có","▸ KHÔNG VONG THEO TRỤ ("&TRIM(IF(C128<>"","năm ","")&IF(D128<>
    #   "","tháng ","")&IF(E128<>"","ngày ","")&IF(F128<>"","giờ ",""))&"): "&I128&" [P2 tr.321-323]"&CHAR(10),"")&IF(H129="Có","▸ [THAM KHẢO] ĐỨC THẦN ("&TRI
    #   M(IF(C129<>"","năm ","")&IF(D129<>"","tháng ","")&IF(E129<>"","ngày ","")&IF(F129<>"","giờ ",""))&"): "&I129&" [Hoàng Thăng Long tr.128 (ngoài P1/P2)]
    #   "&CHAR(10),"")&IF(H130="Có","▸ [THAM KHẢO] CÁCH GIÁC ("&TRIM(IF(C130<>"","năm ","")&IF(D130<>"","tháng ","")&IF(E130<>"","ngày ","")&IF(F130<>"","giờ 
    #   ",""))&"): "&I130&" [Hoàng Thăng Long tr.129 (ngoài P1/P2)]"&CHAR(10),"")&IF(H131="Có","▸ [THAM KHẢO] ĐÀO HOA SÁT ("&TRIM(IF(C131<>"","năm ","")&IF(D1
    #   31<>"","tháng ","")&IF(E131<>"","ngày ","")&IF(F131<>"","giờ ",""))&"): "&I131&" [Hoàng Thăng Long tr.108 (ngoài P1/P2)]"&CHAR(10),""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat('▸ LỘC THEO CAN CHI: ',E.v(12,r-1,c+6)),' [P2 tr.308-309]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-1,c+5),'Có')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THÁI CỰC QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-32,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-32,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-32,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-32,c+3),'')) else '')))),'): '),E.v(12,r-32,c+6)),' [P2 tr.296-297]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-32,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ PHÚC TINH QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-31,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-31,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-31,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-31,c+3),'')) else '')))),'): '),E.v(12,r-31,c+6)),' [P2 tr.300]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-31,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ KHÔI CANH QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-30,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-30,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-30,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-30,c+3),'')) else '')))),'): '),E.v(12,r-30,c+6)),' [P2 tr.301]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-30,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ QUỐC ẤN QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-29,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-29,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-29,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-29,c+3),'')) else '')))),'): '),E.v(12,r-29,c+6)),' [P2 tr.302]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-29,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ HỌC ĐƯỜNG (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-28,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-28,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-28,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-28,c+3),'')) else '')))),'): '),E.v(12,r-28,c+6)),' [P2 tr.302-303]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-28,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TỪ QUÁN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-27,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-27,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-27,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-27,c+3),'')) else '')))),'): '),E.v(12,r-27,c+6)),' [P2 tr.302-303]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-27,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ ĐỨC QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-26,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-26,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-26,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-26,c+3),'')) else '')))),'): '),E.v(12,r-26,c+6)),' [P2 tr.303]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-26,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TÚ QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-25,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-25,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-25,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-25,c+3),'')) else '')))),'): '),E.v(12,r-25,c+6)),' [P2 tr.303]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-25,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ KIM DƯ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-24,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-24,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-24,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-24,c+3),'')) else '')))),'): '),E.v(12,r-24,c+6)),' [P2 tr.306]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-24,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ KIM THẦN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-23,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-23,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-23,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-23,c+3),'')) else '')))),'): '),E.v(12,r-23,c+6)),' [P2 tr.307]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-23,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THIÊN Y (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-22,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-22,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-22,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-22,c+3),'')) else '')))),'): '),E.v(12,r-22,c+6)),' [P2 tr.307-308]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-22,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ CỦNG LỘC (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-21,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-21,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-21,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-21,c+3),'')) else '')))),'): '),E.v(12,r-21,c+6)),' [P2 tr.310]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-21,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ GIÁP LỘC (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-20,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-20,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-20,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-20,c+3),'')) else '')))),'): '),E.v(12,r-20,c+6)),' [P2 tr.310]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-20,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THIÊN XÁ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-19,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-19,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-19,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-19,c+3),'')) else '')))),'): '),E.v(12,r-19,c+6)),' [P2 tr.311]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-19,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TAM KỲ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-18,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-18,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-18,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-18,c+3),'')) else '')))),'): '),E.v(12,r-18,c+6)),' [P2 tr.299-300]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-18,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THIÊN LA (TUẤT – HỢI) (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-17,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-17,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-17,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-17,c+3),'')) else '')))),'): '),E.v(12,r-17,c+6)),' [P2 tr.313-314]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-17,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ ĐỊA VÕNG (THÌN – TỊ) (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-16,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-16,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-16,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-16,c+3),'')) else '')))),'): '),E.v(12,r-16,c+6)),' [P2 tr.313-314]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-16,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TAI SÁT (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-15,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-15,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-15,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-15,c+3),'')) else '')))),'): '),E.v(12,r-15,c+6)),' [P2 tr.317-318]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-15,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ CÂU (CẤU) (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-14,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-14,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-14,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-14,c+3),'')) else '')))),'): '),E.v(12,r-14,c+6)),' [P2 tr.318]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-14,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ GIẢO (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-13,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-13,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-13,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-13,c+3),'')) else '')))),'): '),E.v(12,r-13,c+6)),' [P2 tr.318]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-13,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ VONG THẦN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-12,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-12,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-12,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-12,c+3),'')) else '')))),'): '),E.v(12,r-12,c+6)),' [P2 tr.320]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-12,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ NGUYÊN THẦN (ĐẠI HAO) (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-11,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-11,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-11,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-11,c+3),'')) else '')))),'): '),E.v(12,r-11,c+6)),' [P2 tr.320-321]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-11,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THẬP ÁC ĐẠI BẠI (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-10,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-10,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-10,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-10,c+3),'')) else '')))),'): '),E.v(12,r-10,c+6)),' [P2 tr.323-324]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-10,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ CÔ LOAN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-9,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-9,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-9,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-9,c+3),'')) else '')))),'): '),E.v(12,r-9,c+6)),' [P2 tr.325]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-9,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ ÂM DƯƠNG LỆCH NHAU (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-8,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-8,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-8,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-8,c+3),'')) else '')))),'): '),E.v(12,r-8,c+6)),' [P2 tr.326]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-8,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TỨ PHẾ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-7,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-7,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-7,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-7,c+3),'')) else '')))),'): '),E.v(12,r-7,c+6)),' [P2 tr.326]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-7,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ KHÔNG VONG THEO TRỤ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-6,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-6,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-6,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-6,c+3),'')) else '')))),'): '),E.v(12,r-6,c+6)),' [P2 tr.321-323]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-6,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ [THAM KHẢO] ĐỨC THẦN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-5,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-5,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-5,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-5,c+3),'')) else '')))),'): '),E.v(12,r-5,c+6)),' [Hoàng Thăng Long tr.128 (ngoài P1/P2)]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-5,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ [THAM KHẢO] CÁCH GIÁC (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-4,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-4,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-4,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-4,c+3),'')) else '')))),'): '),E.v(12,r-4,c+6)),' [Hoàng Thăng Long tr.129 (ngoài P1/P2)]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-4,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('▸ [THAM KHẢO] ĐÀO HOA SÁT (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.ne(E.v(12,r-3,c),'')) else ''),('tháng ' if X.truthy(X.ne(E.v(12,r-3,c+1),'')) else '')),('ngày ' if X.truthy(X.ne(E.v(12,r-3,c+2),'')) else '')),('giờ ' if X.truthy(X.ne(E.v(12,r-3,c+3),'')) else '')))),'): '),E.v(12,r-3,c+6)),' [Hoàng Thăng Long tr.108 (ngoài P1/P2)]'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,r-3,c+5),'Có')) else '')))

def f313(E, r, c):
    # Ô E138
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG
    #   _GIAN!$E$5=4,TRUNG_GIAN!$E$6=10))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),10.0))))) else '')

def f314(E, r, c):
    # Ô E139
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!$E$6=9),AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!$E$6=11),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$E$6=5),AND(TRUNG
    #   _GIAN!$E$5=9,TRUNG_GIAN!$E$6=3))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,5),9.0)), X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,5),11.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,5),3.0))))) else '')

def f315(E, r, c):
    # Ô E140
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=4),AND(TRUNG_
    #   GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=10))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),10.0))))) else '')

def f316(E, r, c):
    # Ô E141
    # = IF(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=6),AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=6),AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=0))),"có","")
    return ('có' if X.truthy(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),6.0)), X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),6.0)), X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),0.0))))) else '')

def f317(E, r, c):
    # Ô F142
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$5,5)+1,8,6,4,2,0),TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$5,5)+
    #   1,8,6,4,2,0)+1)),"có","")
    return ('có' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(E.v(23,5,5), 5.0),1.0), lambda: 8.0, lambda: 6.0, lambda: 4.0, lambda: 2.0, lambda: 0.0)), X.eq(E.v(23,6,6),X.add(X.CHOOSE(X.add(X.MOD(E.v(23,5,5), 5.0),1.0), lambda: 8.0, lambda: 6.0, lambda: 4.0, lambda: 2.0, lambda: 0.0),1.0))))) else '')

def f318(E, r, c):
    # Ô E143
    # = IF(OR(AND(OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$D$6=4),OR(AND(TRUNG_GIAN!$E$5=1,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$E$5=7,TRUNG_GIAN!$E$6=
    #   3))),AND(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),OR(AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=6),AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=6
    #   ))),AND(OR(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=10),OR(AND(TRUNG_GIAN!$E$5=7,TRUNG_GIAN!$E$6=9),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$E$6=9
    #   ))),AND(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),OR(AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=0
    #   )))),"có","")
    return ('có' if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,4),4.0)), X.OR(X.AND(X.eq(E.v(23,5,5),1.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,6,5),3.0)))), X.AND(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0)), X.OR(X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),6.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),6.0)))), X.AND(X.OR(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),10.0)), X.OR(X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,6,5),9.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,5),9.0)))), X.AND(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.OR(X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),0.0)))))) else '')

def f319(E, r, c):
    # Ô C144
    # = TRIM(IF(H138="Có","• "&B138&": "&I138&" ("&J138&") ","")&IF(H139="Có","• "&B139&": "&I139&" ("&J139&") ","")&IF(H140="Có","• "&B140&": "&I140&" ("&J14
    #   0&") ","")&IF(H141="Có","• "&B141&": "&I141&" ("&J141&") ","")&IF(H142="Có","• "&B142&": "&I142&" ("&J142&") ","")&IF(H143="Có","• "&B143&": "&I143&" 
    #   ("&J143&") ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(12,r-6,c-1)),': '),E.v(12,r-6,c+6)),' ('),E.v(12,r-6,c+7)),') ') if X.truthy(X.eq(E.v(12,r-6,c+5),'Có')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(12,r-5,c-1)),': '),E.v(12,r-5,c+6)),' ('),E.v(12,r-5,c+7)),') ') if X.truthy(X.eq(E.v(12,r-5,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(12,r-4,c-1)),': '),E.v(12,r-4,c+6)),' ('),E.v(12,r-4,c+7)),') ') if X.truthy(X.eq(E.v(12,r-4,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(12,r-3,c-1)),': '),E.v(12,r-3,c+6)),' ('),E.v(12,r-3,c+7)),') ') if X.truthy(X.eq(E.v(12,r-3,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(12,r-2,c-1)),': '),E.v(12,r-2,c+6)),' ('),E.v(12,r-2,c+7)),') ') if X.truthy(X.eq(E.v(12,r-2,c+5),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(12,r-1,c-1)),': '),E.v(12,r-1,c+6)),' ('),E.v(12,r-1,c+7)),') ') if X.truthy(X.eq(E.v(12,r-1,c+5),'Có')) else '')))

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
    'D11': (f20, 11, 4),
    'E11': (f20, 11, 5),
    'F11': (f20, 11, 6),
    'C12': (f21, 12, 3),
    'D12': (f22, 12, 4),
    'E12': (f23, 12, 5),
    'F12': (f24, 12, 6),
    'C13': (f25, 13, 3),
    'D13': (f26, 13, 4),
    'E13': (f27, 13, 5),
    'F13': (f28, 13, 6),
    'C16': (f29, 16, 3),
    'C17': (f30, 17, 3),
    'D17': (f31, 17, 4),
    'C18': (f32, 18, 3),
    'D18': (f31, 18, 4),
    'C19': (f33, 19, 3),
    'D19': (f34, 19, 4),
    'E19': (f35, 19, 5),
    'C20': (f36, 20, 3),
    'C21': (f37, 21, 3),
    'C22': (f38, 22, 3),
    'C23': (f39, 23, 3),
    'C27': (f40, 27, 3),
    'E27': (f41, 27, 5),
    'F27': (f42, 27, 6),
    'G27': (f43, 27, 7),
    'H27': (f44, 27, 8),
    'I27': (f45, 27, 9),
    'J27': (f46, 27, 10),
    'C28': (f40, 28, 3),
    'E28': (f47, 28, 5),
    'F28': (f48, 28, 6),
    'G28': (f49, 28, 7),
    'H28': (f50, 28, 8),
    'I28': (f51, 28, 9),
    'J28': (f52, 28, 10),
    'C29': (f40, 29, 3),
    'E29': (f53, 29, 5),
    'F29': (f54, 29, 6),
    'G29': (f55, 29, 7),
    'H29': (f56, 29, 8),
    'I29': (f57, 29, 9),
    'J29': (f58, 29, 10),
    'C30': (f40, 30, 3),
    'E30': (f59, 30, 5),
    'F30': (f60, 30, 6),
    'G30': (f61, 30, 7),
    'H30': (f62, 30, 8),
    'I30': (f63, 30, 9),
    'J30': (f64, 30, 10),
    'C31': (f65, 31, 3),
    'E31': (f66, 31, 5),
    'F31': (f42, 31, 6),
    'G31': (f43, 31, 7),
    'H31': (f44, 31, 8),
    'I31': (f45, 31, 9),
    'J31': (f46, 31, 10),
    'C32': (f65, 32, 3),
    'E32': (f67, 32, 5),
    'F32': (f48, 32, 6),
    'G32': (f49, 32, 7),
    'H32': (f50, 32, 8),
    'I32': (f51, 32, 9),
    'J32': (f52, 32, 10),
    'C33': (f65, 33, 3),
    'E33': (f68, 33, 5),
    'F33': (f54, 33, 6),
    'G33': (f55, 33, 7),
    'H33': (f56, 33, 8),
    'I33': (f57, 33, 9),
    'J33': (f58, 33, 10),
    'C34': (f65, 34, 3),
    'E34': (f69, 34, 5),
    'F34': (f60, 34, 6),
    'G34': (f61, 34, 7),
    'H34': (f62, 34, 8),
    'I34': (f63, 34, 9),
    'J34': (f64, 34, 10),
    'C35': (f70, 35, 3),
    'E35': (f71, 35, 5),
    'F35': (f42, 35, 6),
    'G35': (f43, 35, 7),
    'H35': (f44, 35, 8),
    'I35': (f45, 35, 9),
    'J35': (f46, 35, 10),
    'C36': (f70, 36, 3),
    'E36': (f72, 36, 5),
    'F36': (f48, 36, 6),
    'G36': (f49, 36, 7),
    'H36': (f50, 36, 8),
    'I36': (f51, 36, 9),
    'J36': (f52, 36, 10),
    'C37': (f70, 37, 3),
    'E37': (f73, 37, 5),
    'F37': (f54, 37, 6),
    'G37': (f55, 37, 7),
    'H37': (f56, 37, 8),
    'I37': (f57, 37, 9),
    'J37': (f58, 37, 10),
    'C38': (f70, 38, 3),
    'E38': (f74, 38, 5),
    'F38': (f60, 38, 6),
    'G38': (f61, 38, 7),
    'H38': (f62, 38, 8),
    'I38': (f63, 38, 9),
    'J38': (f64, 38, 10),
    'C39': (f75, 39, 3),
    'E39': (f76, 39, 5),
    'F39': (f42, 39, 6),
    'G39': (f43, 39, 7),
    'H39': (f44, 39, 8),
    'I39': (f45, 39, 9),
    'J39': (f46, 39, 10),
    'C40': (f75, 40, 3),
    'E40': (f77, 40, 5),
    'F40': (f48, 40, 6),
    'G40': (f49, 40, 7),
    'H40': (f50, 40, 8),
    'I40': (f51, 40, 9),
    'J40': (f52, 40, 10),
    'C41': (f75, 41, 3),
    'E41': (f78, 41, 5),
    'F41': (f54, 41, 6),
    'G41': (f55, 41, 7),
    'H41': (f56, 41, 8),
    'I41': (f57, 41, 9),
    'J41': (f58, 41, 10),
    'C42': (f75, 42, 3),
    'E42': (f79, 42, 5),
    'F42': (f60, 42, 6),
    'G42': (f61, 42, 7),
    'H42': (f62, 42, 8),
    'I42': (f63, 42, 9),
    'J42': (f64, 42, 10),
    'C43': (f80, 43, 3),
    'E43': (f81, 43, 5),
    'F43': (f42, 43, 6),
    'G43': (f43, 43, 7),
    'H43': (f44, 43, 8),
    'I43': (f45, 43, 9),
    'J43': (f46, 43, 10),
    'C44': (f80, 44, 3),
    'E44': (f82, 44, 5),
    'F44': (f48, 44, 6),
    'G44': (f49, 44, 7),
    'H44': (f50, 44, 8),
    'I44': (f51, 44, 9),
    'J44': (f52, 44, 10),
    'C45': (f80, 45, 3),
    'E45': (f83, 45, 5),
    'F45': (f54, 45, 6),
    'G45': (f55, 45, 7),
    'H45': (f56, 45, 8),
    'I45': (f57, 45, 9),
    'J45': (f58, 45, 10),
    'C46': (f80, 46, 3),
    'E46': (f84, 46, 5),
    'F46': (f60, 46, 6),
    'G46': (f61, 46, 7),
    'H46': (f62, 46, 8),
    'I46': (f63, 46, 9),
    'J46': (f64, 46, 10),
    'C47': (f85, 47, 3),
    'E47': (f86, 47, 5),
    'F47': (f42, 47, 6),
    'G47': (f43, 47, 7),
    'H47': (f44, 47, 8),
    'I47': (f45, 47, 9),
    'J47': (f46, 47, 10),
    'C48': (f85, 48, 3),
    'E48': (f87, 48, 5),
    'F48': (f48, 48, 6),
    'G48': (f49, 48, 7),
    'H48': (f50, 48, 8),
    'I48': (f51, 48, 9),
    'J48': (f52, 48, 10),
    'C49': (f85, 49, 3),
    'E49': (f88, 49, 5),
    'F49': (f54, 49, 6),
    'G49': (f55, 49, 7),
    'H49': (f56, 49, 8),
    'I49': (f57, 49, 9),
    'J49': (f58, 49, 10),
    'C50': (f85, 50, 3),
    'E50': (f89, 50, 5),
    'F50': (f60, 50, 6),
    'G50': (f61, 50, 7),
    'H50': (f62, 50, 8),
    'I50': (f63, 50, 9),
    'J50': (f64, 50, 10),
    'C53': (f90, 53, 3),
    'C54': (f91, 54, 3),
    'C55': (f92, 55, 3),
    'C56': (f93, 56, 3),
    'C57': (f94, 57, 3),
    'C85': (f95, 85, 3),
    'D85': (f96, 85, 4),
    'E85': (f97, 85, 5),
    'F85': (f98, 85, 6),
    'H85': (f99, 85, 8),
    'I85': (f100, 85, 9),
    'C86': (f101, 86, 3),
    'D86': (f102, 86, 4),
    'E86': (f103, 86, 5),
    'F86': (f104, 86, 6),
    'H86': (f99, 86, 8),
    'I86': (f105, 86, 9),
    'C87': (f106, 87, 3),
    'D87': (f107, 87, 4),
    'E87': (f108, 87, 5),
    'F87': (f109, 87, 6),
    'H87': (f99, 87, 8),
    'I87': (f110, 87, 9),
    'C88': (f111, 88, 3),
    'D88': (f112, 88, 4),
    'E88': (f113, 88, 5),
    'F88': (f114, 88, 6),
    'H88': (f99, 88, 8),
    'I88': (f115, 88, 9),
    'C89': (f116, 89, 3),
    'D89': (f117, 89, 4),
    'E89': (f118, 89, 5),
    'F89': (f119, 89, 6),
    'H89': (f99, 89, 8),
    'I89': (f120, 89, 9),
    'C90': (f121, 90, 3),
    'D90': (f122, 90, 4),
    'E90': (f123, 90, 5),
    'F90': (f124, 90, 6),
    'H90': (f99, 90, 8),
    'I90': (f125, 90, 9),
    'C91': (f126, 91, 3),
    'D91': (f127, 91, 4),
    'E91': (f128, 91, 5),
    'F91': (f129, 91, 6),
    'H91': (f99, 91, 8),
    'I91': (f130, 91, 9),
    'C92': (f131, 92, 3),
    'D92': (f132, 92, 4),
    'E92': (f133, 92, 5),
    'F92': (f134, 92, 6),
    'H92': (f99, 92, 8),
    'I92': (f135, 92, 9),
    'C93': (f136, 93, 3),
    'D93': (f137, 93, 4),
    'E93': (f138, 93, 5),
    'F93': (f139, 93, 6),
    'H93': (f99, 93, 8),
    'I93': (f140, 93, 9),
    'C94': (f141, 94, 3),
    'D94': (f142, 94, 4),
    'E94': (f143, 94, 5),
    'F94': (f144, 94, 6),
    'H94': (f145, 94, 8),
    'I94': (f146, 94, 9),
    'C95': (f141, 95, 3),
    'D95': (f147, 95, 4),
    'E95': (f148, 95, 5),
    'F95': (f149, 95, 6),
    'H95': (f145, 95, 8),
    'I95': (f150, 95, 9),
    'C96': (f151, 96, 3),
    'D96': (f152, 96, 4),
    'E96': (f153, 96, 5),
    'F96': (f154, 96, 6),
    'H96': (f99, 96, 8),
    'I96': (f155, 96, 9),
    'C102': (f156, 102, 3),
    'D102': (f157, 102, 4),
    'E102': (f158, 102, 5),
    'F102': (f159, 102, 6),
    'H102': (f99, 102, 8),
    'I102': (f160, 102, 9),
    'C103': (f161, 103, 3),
    'D103': (f162, 103, 4),
    'E103': (f163, 103, 5),
    'F103': (f164, 103, 6),
    'H103': (f99, 103, 8),
    'I103': (f165, 103, 9),
    'C104': (f166, 104, 3),
    'D104': (f167, 104, 4),
    'E104': (f168, 104, 5),
    'F104': (f169, 104, 6),
    'H104': (f99, 104, 8),
    'I104': (f170, 104, 9),
    'C105': (f171, 105, 3),
    'D105': (f172, 105, 4),
    'E105': (f173, 105, 5),
    'F105': (f174, 105, 6),
    'H105': (f99, 105, 8),
    'I105': (f175, 105, 9),
    'C106': (f176, 106, 3),
    'D106': (f177, 106, 4),
    'E106': (f178, 106, 5),
    'F106': (f179, 106, 6),
    'H106': (f99, 106, 8),
    'I106': (f180, 106, 9),
    'C107': (f181, 107, 3),
    'D107': (f182, 107, 4),
    'E107': (f183, 107, 5),
    'F107': (f184, 107, 6),
    'H107': (f99, 107, 8),
    'I107': (f185, 107, 9),
    'C108': (f186, 108, 3),
    'D108': (f187, 108, 4),
    'E108': (f188, 108, 5),
    'F108': (f189, 108, 6),
    'H108': (f99, 108, 8),
    'I108': (f190, 108, 9),
    'C109': (f191, 109, 3),
    'D109': (f192, 109, 4),
    'E109': (f193, 109, 5),
    'F109': (f194, 109, 6),
    'H109': (f99, 109, 8),
    'I109': (f195, 109, 9),
    'C110': (f196, 110, 3),
    'D110': (f197, 110, 4),
    'E110': (f198, 110, 5),
    'F110': (f199, 110, 6),
    'H110': (f99, 110, 8),
    'I110': (f200, 110, 9),
    'C111': (f201, 111, 3),
    'D111': (f202, 111, 4),
    'E111': (f203, 111, 5),
    'F111': (f204, 111, 6),
    'H111': (f99, 111, 8),
    'I111': (f205, 111, 9),
    'C112': (f206, 112, 3),
    'D112': (f207, 112, 4),
    'E112': (f208, 112, 5),
    'F112': (f209, 112, 6),
    'H112': (f99, 112, 8),
    'I112': (f210, 112, 9),
    'C113': (f211, 113, 3),
    'D113': (f212, 113, 4),
    'E113': (f213, 113, 5),
    'F113': (f214, 113, 6),
    'H113': (f99, 113, 8),
    'I113': (f215, 113, 9),
    'C114': (f216, 114, 3),
    'D114': (f217, 114, 4),
    'E114': (f218, 114, 5),
    'F114': (f219, 114, 6),
    'H114': (f99, 114, 8),
    'I114': (f220, 114, 9),
    'C115': (f221, 115, 3),
    'D115': (f222, 115, 4),
    'E115': (f223, 115, 5),
    'F115': (f224, 115, 6),
    'H115': (f99, 115, 8),
    'I115': (f225, 115, 9),
    'C116': (f226, 116, 3),
    'D116': (f227, 116, 4),
    'E116': (f228, 116, 5),
    'F116': (f229, 116, 6),
    'H116': (f99, 116, 8),
    'I116': (f230, 116, 9),
    'C117': (f231, 117, 3),
    'D117': (f232, 117, 4),
    'E117': (f233, 117, 5),
    'F117': (f234, 117, 6),
    'H117': (f99, 117, 8),
    'I117': (f235, 117, 9),
    'C118': (f236, 118, 3),
    'D118': (f237, 118, 4),
    'E118': (f238, 118, 5),
    'F118': (f239, 118, 6),
    'H118': (f99, 118, 8),
    'I118': (f240, 118, 9),
    'C119': (f241, 119, 3),
    'D119': (f242, 119, 4),
    'E119': (f243, 119, 5),
    'F119': (f244, 119, 6),
    'H119': (f99, 119, 8),
    'I119': (f245, 119, 9),
    'C120': (f246, 120, 3),
    'D120': (f247, 120, 4),
    'E120': (f248, 120, 5),
    'F120': (f249, 120, 6),
    'H120': (f99, 120, 8),
    'I120': (f250, 120, 9),
    'C121': (f251, 121, 3),
    'D121': (f252, 121, 4),
    'E121': (f253, 121, 5),
    'F121': (f254, 121, 6),
    'H121': (f99, 121, 8),
    'I121': (f255, 121, 9),
    'C122': (f256, 122, 3),
    'D122': (f257, 122, 4),
    'E122': (f258, 122, 5),
    'F122': (f259, 122, 6),
    'H122': (f99, 122, 8),
    'I122': (f260, 122, 9),
    'C123': (f261, 123, 3),
    'D123': (f262, 123, 4),
    'E123': (f263, 123, 5),
    'F123': (f264, 123, 6),
    'H123': (f99, 123, 8),
    'I123': (f265, 123, 9),
    'C124': (f266, 124, 3),
    'D124': (f267, 124, 4),
    'E124': (f268, 124, 5),
    'F124': (f269, 124, 6),
    'H124': (f99, 124, 8),
    'I124': (f270, 124, 9),
    'C125': (f271, 125, 3),
    'D125': (f272, 125, 4),
    'E125': (f273, 125, 5),
    'F125': (f274, 125, 6),
    'H125': (f99, 125, 8),
    'I125': (f275, 125, 9),
    'C126': (f276, 126, 3),
    'D126': (f277, 126, 4),
    'E126': (f278, 126, 5),
    'F126': (f279, 126, 6),
    'H126': (f99, 126, 8),
    'I126': (f280, 126, 9),
    'C127': (f281, 127, 3),
    'D127': (f282, 127, 4),
    'E127': (f283, 127, 5),
    'F127': (f284, 127, 6),
    'H127': (f99, 127, 8),
    'I127': (f285, 127, 9),
    'C128': (f286, 128, 3),
    'D128': (f287, 128, 4),
    'E128': (f288, 128, 5),
    'F128': (f289, 128, 6),
    'H128': (f99, 128, 8),
    'I128': (f290, 128, 9),
    'C129': (f291, 129, 3),
    'D129': (f292, 129, 4),
    'E129': (f293, 129, 5),
    'F129': (f294, 129, 6),
    'H129': (f99, 129, 8),
    'I129': (f295, 129, 9),
    'C130': (f296, 130, 3),
    'D130': (f297, 130, 4),
    'E130': (f298, 130, 5),
    'F130': (f299, 130, 6),
    'H130': (f99, 130, 8),
    'I130': (f300, 130, 9),
    'C131': (f301, 131, 3),
    'D131': (f302, 131, 4),
    'E131': (f303, 131, 5),
    'F131': (f304, 131, 6),
    'H131': (f99, 131, 8),
    'I131': (f305, 131, 9),
    'C133': (f306, 133, 3),
    'D133': (f306, 133, 4),
    'E133': (f306, 133, 5),
    'F133': (f306, 133, 6),
    'H133': (f145, 133, 8),
    'I133': (f307, 133, 9),
    'L133': (f308, 133, 12),
    'M133': (f309, 133, 13),
    'N133': (f310, 133, 14),
    'O133': (f311, 133, 15),
    'C134': (f312, 134, 3),
    'C138': (f141, 138, 3),
    'D138': (f141, 138, 4),
    'E138': (f313, 138, 5),
    'F138': (f141, 138, 6),
    'H138': (f145, 138, 8),
    'C139': (f141, 139, 3),
    'D139': (f141, 139, 4),
    'E139': (f314, 139, 5),
    'F139': (f141, 139, 6),
    'H139': (f145, 139, 8),
    'C140': (f141, 140, 3),
    'D140': (f141, 140, 4),
    'E140': (f315, 140, 5),
    'F140': (f141, 140, 6),
    'H140': (f145, 140, 8),
    'C141': (f141, 141, 3),
    'D141': (f141, 141, 4),
    'E141': (f316, 141, 5),
    'F141': (f141, 141, 6),
    'H141': (f145, 141, 8),
    'C142': (f141, 142, 3),
    'D142': (f141, 142, 4),
    'E142': (f141, 142, 5),
    'F142': (f317, 142, 6),
    'H142': (f145, 142, 8),
    'C143': (f141, 143, 3),
    'D143': (f141, 143, 4),
    'E143': (f318, 143, 5),
    'F143': (f141, 143, 6),
    'H143': (f145, 143, 8),
    'C144': (f319, 144, 3),
}
