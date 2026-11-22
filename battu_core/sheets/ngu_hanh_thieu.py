# -*- coding: utf-8 -*-
"""Module sheet NGU_HANH_THIEU – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.9.xlsx (SHA-256 671245aea574f577f3cb312db9b3db0d1ce1973b107488ca7b828a9ceb4d60da).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'NGU_HANH_THIEU'
SI = 29

def f0(E, r, c):
    # Ô G3
    # = IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2","Mặc định: theo sách (giả, hợp, Mậu/Kỷ)","Tái tạo đúng trang tính")
    return ('Mặc định: theo sách (giả, hợp, Mậu/Kỷ)' if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else 'Tái tạo đúng trang tính')

def f1(E, r, c):
    # Ô B5
    # = " Tứ trụ: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" – "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" – "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" – "&IF(NHA
    #   P_LIEU!$C$9="Không rõ giờ","(không rõ giờ)",TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11)&". Chi tháng: "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$
    #   9+1)&"); can ngày: "&TRUNG_GIAN!$E$10&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&")."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' Tứ trụ: ',E.v(23,10,3)),' '),E.v(23,11,3)),' – '),E.v(23,10,4)),' '),E.v(23,11,4)),' – '),E.v(23,10,5)),' '),E.v(23,11,5)),' – '),('(không rõ giờ)' if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else X.cat(X.cat(E.v(23,10,6),' '),E.v(23,11,6)))),'. Chi tháng: '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),'); can ngày: '),E.v(23,10,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),').')

def f2(E, r, c):
    # Ô B7
    # = "Xét được lệnh (can tháng – chi tháng): "&INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" – "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)
    return X.cat(X.cat(X.cat('Xét được lệnh (can tháng – chi tháng): ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0))),' – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)))

def f3(E, r, c):
    # Ô C7
    # = INDEX(STAGE_TEN,IF(MOD(TRUNG_GIAN!$D$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$D$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$D$5+1,11,6,2
    #   ,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)
    return X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0))

def f4(E, r, c):
    # Ô D7
    # = IF(IF(MOD(TRUNG_GIAN!$D$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$D$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$D$5+1,11,6,2,9,8,3,5,0,8,
    #   3)-TRUNG_GIAN!$D$6,12))<=4,"Có","Không")
    return ('Có' if X.truthy(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0)) else 'Không')

def f5(E, r, c):
    # Ô E7
    # = IF(IF(MOD(TRUNG_GIAN!$D$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$D$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$D$5+1,11,6,2,9,8,3,5,0,8,
    #   3)-TRUNG_GIAN!$D$6,12))<=4,0.48,0)
    return (0.48 if X.truthy(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0)) else 0.0)

def f6(E, r, c):
    # Ô B8
    # = "Xét đắc địa – nhật can "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" tại chi năm, ngày, giờ (P1 tr.70)"
    return X.cat(X.cat('Xét đắc địa – nhật can ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' tại chi năm, ngày, giờ (P1 tr.70)')

def f7(E, r, c):
    # Ô B9
    # = IF(TRUE(),"   Chi năm: "&TRUNG_GIAN!$C$11,"   Chi năm: (không rõ giờ)")
    return (X.cat('   Chi năm: ',E.v(23,11,3)) if X.truthy(X.TRUE()) else '   Chi năm: (không rõ giờ)')

def f8(E, r, c):
    # Ô C9
    # = IF(TRUE(),TRUNG_GIAN!$C$29,"")
    return (E.v(23,29,3) if X.truthy(X.TRUE()) else '')

def f9(E, r, c):
    # Ô D9
    # = IF(NOT(TRUE()),"0",IF(NOT(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2"),"PA1: 0",IF(IF(PP1_DINH_TINH!C27>=0,IF(PP1_DINH_TINH!C27=TRUNG_GIAN!$E$7,1,0),IF(PP1_DINH
    #   _TINH!C28=1,0,IF(OR(AND(OR(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!C$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$
    #   5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!C$6,12))=3,IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!C$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(
    #   CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!C$6,12))=4),TRUNG_GIAN!C$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(IF(MOD(TRUNG_GIAN!$E$5
    #   ,2)=0,MOD(TRUNG_GIAN!C$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!C$6,12))=0,I
    #   F(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!C$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TR
    #   UNG_GIAN!C$6,12))=8)),AND(P_DUKHI=1,PP1_DINH_TINH!C29=0,OR(AND(TRUNG_GIAN!$E$5=1,TRUNG_GIAN!C$6=4),AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!C$6=7),AND(TRUNG_G
    #   IAN!$E$5=7,TRUNG_GIAN!C$6=10),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!C$6=1)))),1,0)))=1,"Có","0")))
    return ('0' if X.truthy(X.NOT(X.TRUE())) else ('PA1: 0' if X.truthy(X.NOT(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'))) else ('Có' if X.truthy(X.eq(((1.0 if X.truthy(X.eq(E.v(9,r+18,c-1),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(9,r+18,c-1),0.0)) else (0.0 if X.truthy(X.eq(E.v(9,r+19,c-1),1.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq((X.MOD(X.sub(E.v(23,6,c-1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c-1)), 12.0)),3.0), X.eq((X.MOD(X.sub(E.v(23,6,c-1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c-1)), 12.0)),4.0)), X.eq(E.v(23,21,c-1),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq((X.MOD(X.sub(E.v(23,6,c-1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c-1)), 12.0)),0.0), X.eq((X.MOD(X.sub(E.v(23,6,c-1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c-1)), 12.0)),8.0))), X.AND(X.eq(E.v(24,76,6),1.0), X.eq(E.v(9,r+20,c-1),0.0), X.OR(X.AND(X.eq(E.v(23,5,5),1.0), X.eq(E.v(23,6,c-1),4.0)), X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,c-1),7.0)), X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,6,c-1),10.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,c-1),1.0)))))) else 0.0))),1.0)) else '0')))

def f10(E, r, c):
    # Ô B10
    # = IF(TRUE(),"   Chi ngày: "&TRUNG_GIAN!$E$11,"   Chi ngày: (không rõ giờ)")
    return (X.cat('   Chi ngày: ',E.v(23,11,5)) if X.truthy(X.TRUE()) else '   Chi ngày: (không rõ giờ)')

def f11(E, r, c):
    # Ô C10
    # = IF(TRUE(),TRUNG_GIAN!$E$29,"")
    return (E.v(23,29,5) if X.truthy(X.TRUE()) else '')

def f12(E, r, c):
    # Ô D10
    # = IF(NOT(TRUE()),"0",IF(NOT(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2"),"PA1: 0",IF(IF(PP1_DINH_TINH!E27>=0,IF(PP1_DINH_TINH!E27=TRUNG_GIAN!$E$7,1,0),IF(PP1_DINH
    #   _TINH!E28=1,0,IF(OR(AND(OR(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!E$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$
    #   5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!E$6,12))=3,IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!E$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(
    #   CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!E$6,12))=4),TRUNG_GIAN!E$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8=0,OR(IF(MOD(TRUNG_GIAN!$E$5
    #   ,2)=0,MOD(TRUNG_GIAN!E$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!E$6,12))=0,I
    #   F(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!E$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TR
    #   UNG_GIAN!E$6,12))=8)),AND(P_DUKHI=1,PP1_DINH_TINH!E29=0,OR(AND(TRUNG_GIAN!$E$5=1,TRUNG_GIAN!E$6=4),AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!E$6=7),AND(TRUNG_G
    #   IAN!$E$5=7,TRUNG_GIAN!E$6=10),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!E$6=1)))),1,0)))=1,"Có","0")))
    return ('0' if X.truthy(X.NOT(X.TRUE())) else ('PA1: 0' if X.truthy(X.NOT(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'))) else ('Có' if X.truthy(X.eq(((1.0 if X.truthy(X.eq(E.v(9,r+17,c+1),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(9,r+17,c+1),0.0)) else (0.0 if X.truthy(X.eq(E.v(9,r+18,c+1),1.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq((X.MOD(X.sub(E.v(23,6,c+1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+1)), 12.0)),3.0), X.eq((X.MOD(X.sub(E.v(23,6,c+1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+1)), 12.0)),4.0)), X.eq(E.v(23,21,c+1),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq((X.MOD(X.sub(E.v(23,6,c+1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+1)), 12.0)),0.0), X.eq((X.MOD(X.sub(E.v(23,6,c+1),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+1)), 12.0)),8.0))), X.AND(X.eq(E.v(24,76,6),1.0), X.eq(E.v(9,r+19,c+1),0.0), X.OR(X.AND(X.eq(E.v(23,5,5),1.0), X.eq(E.v(23,6,c+1),4.0)), X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,c+1),7.0)), X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,6,c+1),10.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,c+1),1.0)))))) else 0.0))),1.0)) else '0')))

def f13(E, r, c):
    # Ô B11
    # = IF(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),"   Chi giờ: "&TRUNG_GIAN!$F$11,"   Chi giờ: (không rõ giờ)")
    return (X.cat('   Chi giờ: ',E.v(23,11,6)) if X.truthy(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))) else '   Chi giờ: (không rõ giờ)')

def f14(E, r, c):
    # Ô C11
    # = IF(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$29,"")
    return (E.v(23,29,6) if X.truthy(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))) else '')

def f15(E, r, c):
    # Ô D11
    # = IF(NOT(NOT(NHAP_LIEU!$C$9="Không rõ giờ")),"0",IF(NOT(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2"),"PA1: 0",IF(IF(PP1_DINH_TINH!F27>=0,IF(PP1_DINH_TINH!F27=TRUN
    #   G_GIAN!$E$7,1,0),IF(PP1_DINH_TINH!F28=1,0,IF(OR(AND(OR(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!F$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),1
    #   2),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!F$6,12))=3,IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!F$6-CHOOSE(TRUNG_GIAN!$E$5+1,11
    #   ,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!F$6,12))=4),TRUNG_GIAN!F$21=TRUNG_GIAN!$E$7),AND(TRUNG_GIAN!$E$8
    #   =0,OR(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!F$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,
    #   8,3)-TRUNG_GIAN!F$6,12))=0,IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!F$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$
    #   5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!F$6,12))=8)),AND(P_DUKHI=1,PP1_DINH_TINH!F29=0,OR(AND(TRUNG_GIAN!$E$5=1,TRUNG_GIAN!F$6=4),AND(TRUNG_GIAN!$E$5=3,T
    #   RUNG_GIAN!F$6=7),AND(TRUNG_GIAN!$E$5=7,TRUNG_GIAN!F$6=10),AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!F$6=1)))),1,0)))=1,"Có","0")))
    return ('0' if X.truthy(X.NOT(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')))) else ('PA1: 0' if X.truthy(X.NOT(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'))) else ('Có' if X.truthy(X.eq(((1.0 if X.truthy(X.eq(E.v(9,r+16,c+2),E.v(23,7,5))) else 0.0) if X.truthy(X.ge(E.v(9,r+16,c+2),0.0)) else (0.0 if X.truthy(X.eq(E.v(9,r+17,c+2),1.0)) else (1.0 if X.truthy(X.OR(X.AND(X.OR(X.eq((X.MOD(X.sub(E.v(23,6,c+2),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+2)), 12.0)),3.0), X.eq((X.MOD(X.sub(E.v(23,6,c+2),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+2)), 12.0)),4.0)), X.eq(E.v(23,21,c+2),E.v(23,7,5))), X.AND(X.eq(E.v(23,8,5),0.0), X.OR(X.eq((X.MOD(X.sub(E.v(23,6,c+2),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+2)), 12.0)),0.0), X.eq((X.MOD(X.sub(E.v(23,6,c+2),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,c+2)), 12.0)),8.0))), X.AND(X.eq(E.v(24,76,6),1.0), X.eq(E.v(9,r+18,c+2),0.0), X.OR(X.AND(X.eq(E.v(23,5,5),1.0), X.eq(E.v(23,6,c+2),4.0)), X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,6,c+2),7.0)), X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,6,c+2),10.0)), X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,c+2),1.0)))))) else 0.0))),1.0)) else '0')))

def f16(E, r, c):
    # Ô B12
    # = "   → Đắc địa ("&IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2","PA2 – theo sách","PA1 – theo trang: luôn 0")&")"
    return X.cat(X.cat('   → Đắc địa (',('PA2 – theo sách' if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else 'PA1 – theo trang: luôn 0')),')')

def f17(E, r, c):
    # Ô E12
    # = IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",IF(OR(D9="Có",D10="Có",D11="Có"),0.2,0),0)
    return ((0.2 if X.truthy(X.OR(X.eq(E.v(29,r-3,c-1),'Có'), X.eq(E.v(29,r-2,c-1),'Có'), X.eq(E.v(29,r-1,c-1),'Có'))) else 0.0) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else 0.0)

def f18(E, r, c):
    # Ô B13
    # = "Xét cùng hành (trang ghi “được sinh”) – can ngày "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&")"
    return X.cat(X.cat(X.cat(X.cat('Xét cùng hành (trang ghi “được sinh”) – can ngày ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),')')

def f19(E, r, c):
    # Ô C13
    # = IF(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),PP2_DINH_LUONG!$C$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND
    #   (TRUE(),PP2_DINH_LUONG!$D$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUON
    #   G!$F$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!
    #   $C$5+1)&" (năm) ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không
    #    rõ giờ"),TRUNG_GIAN!$F$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))="","—",IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TR
    #   UE(),PP2_DINH_LUONG!$C$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$52=TRUNG_GIAN!$E$7),INDEX(CA
    #   N_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5
    #   +1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$7=TR
    #   UNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$7=TRUNG_GIAN!$E$7),INDEX(CAN_T
    #   EN,TRUNG_GIAN!$F$5+1)&" (giờ) ",""))))
    return ('—' if X.truthy(X.eq((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,52,6),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),E.v(23,7,5)))) else '')))),'')) else (X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,52,6),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),E.v(23,7,5)))) else '')))))

def f20(E, r, c):
    # Ô D13
    # = IF(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),PP2_DINH_LUONG!$C$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND
    #   (TRUE(),PP2_DINH_LUONG!$D$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUON
    #   G!$F$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!
    #   $C$5+1)&" (năm) ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không
    #    rõ giờ"),TRUNG_GIAN!$F$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))="","0","Có")
    return ('0' if X.truthy(X.eq((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,52,6),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),E.v(23,7,5)))) else '')))),'')) else 'Có')

def f21(E, r, c):
    # Ô E13
    # = IF(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),PP2_DINH_LUONG!$C$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND
    #   (TRUE(),PP2_DINH_LUONG!$D$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUON
    #   G!$F$52=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!
    #   $C$5+1)&" (năm) ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không
    #    rõ giờ"),TRUNG_GIAN!$F$7=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))="",0,0.2)
    return (0.0 if X.truthy(X.eq((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,52,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,52,6),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,3),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,7,4),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,7,6),E.v(23,7,5)))) else '')))),'')) else 0.2)

def f22(E, r, c):
    # Ô B14
    # = "Xét được sinh (trang ghi “được trợ giúp”) – can ngày "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&")"
    return X.cat(X.cat(X.cat(X.cat('Xét được sinh (trang ghi “được trợ giúp”) – can ngày ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),')')

def f23(E, r, c):
    # Ô C14
    # = IF(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$C$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","
    #   ")&IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$D$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),MOD(PP2_DINH_LUONG!$F$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_
    #   GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (
    #   tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))="","
    #   —",IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$C$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","
    #   ")&IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$D$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),MOD(PP2_DINH_LUONG!$F$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_
    #   GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (
    #   tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))&" si
    #   nh "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1))
    return ('—' if X.truthy(X.eq((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(10,52,6),1.0), 5.0),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)))) else '')))),'')) else X.cat(X.cat((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(10,52,6),1.0), 5.0),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)))) else '')))),' sinh '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))))

def f24(E, r, c):
    # Ô D14
    # = IF(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$C$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","
    #   ")&IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$D$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),MOD(PP2_DINH_LUONG!$F$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_
    #   GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (
    #   tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))="","
    #   0","Có")
    return ('0' if X.truthy(X.eq((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(10,52,6),1.0), 5.0),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)))) else '')))),'')) else 'Có')

def f25(E, r, c):
    # Ô E14
    # = IF(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",TRIM(IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$C$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","
    #   ")&IF(AND(TRUE(),MOD(PP2_DINH_LUONG!$D$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),MOD(PP2_DINH_LUONG!$F$52+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")),TRIM(IF(AND(TRUE(),MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_
    #   GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (
    #   tháng) ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ","")))="",0
    #   ,0.12)
    return (0.0 if X.truthy(X.eq((X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(10,52,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(10,52,6),1.0), 5.0),E.v(23,7,5)))) else ''))) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)))) else '')))),'')) else 0.12)

def f26(E, r, c):
    # Ô C15
    # = INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)
    return X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))

def f27(E, r, c):
    # Ô E15
    # = E7+E12+E13+E14
    return X.add(X.add(X.add(E.v(29,r-8,c),E.v(29,r-3,c)),E.v(29,r-2,c)),E.v(29,r-1,c))

def f28(E, r, c):
    # Ô F15
    # = IF(E15<=$H$15,"Thiếu – cần BỔ SUNG",IF(E15>=$H$16,IF(E15>=0.8,"Quá vượng – cần TIẾT GIẢM","Cần TIẾT GIẢM"),"Cân bằng khí – không cần bổ sung"))
    return ('Thiếu – cần BỔ SUNG' if X.truthy(X.le(E.v(29,r,c-1),E.v(29,15,8))) else (('Quá vượng – cần TIẾT GIẢM' if X.truthy(X.ge(E.v(29,r,c-1),0.8)) else 'Cần TIẾT GIẢM') if X.truthy(X.ge(E.v(29,r,c-1),E.v(29,16,8))) else 'Cân bằng khí – không cần bổ sung'))

def f29(E, r, c):
    # Ô C16
    # = INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" – "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)
    return X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0)),' – '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0)))

def f30(E, r, c):
    # Ô F16
    # = IF(E15<=$H$15,"Bổ sung chính ngũ hành của can ngày: "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1),IF(E15>=$H$16,"Tiết giảm bằng ngũ hành do can ngày sinh ra: "&
    #   INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1),"Không cần bổ sung (46%–55%)"))
    return (X.cat('Bổ sung chính ngũ hành của can ngày: ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))) if X.truthy(X.le(E.v(29,r-1,c-1),E.v(29,15,8))) else (X.cat('Tiết giảm bằng ngũ hành do can ngày sinh ra: ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))) if X.truthy(X.ge(E.v(29,r-1,c-1),E.v(29,16,8))) else 'Không cần bổ sung (46%–55%)'))

def f31(E, r, c):
    # Ô B18
    # = "⚠ XUNG – HỢP – HÓA (công cụ tính theo Thiệu Vĩ Hoa – điểm ở trên KHÔNG trừ/cộng theo các quan hệ này, cần tự cân nhắc): "&IF(TRIM(TRIM(IF(AND(TRUE(),
    #   TRUE(),PP2_DINH_LUONG!$C$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) – "&INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ngũ hợp: "&IF(PP2_DINH_LUONG!$C
    #   $47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$52+1),IF(PP2_DINH_LUONG!$C$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$C$46=1,"hợp không hóa – giữ n
    #   gũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ","")&IF(AND(TRUE(),TRUE(),PP2_DINH_LUONG!$D$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)
    #   &" (tháng) – "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) ngũ hợp: "&IF(PP2_DINH_LUONG!$D$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$52+1),IF(PP2_DIN
    #   H_LUONG!$D$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$D$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp"))
    #   )&"; ","")&IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) – "&INDEX(CAN_TEN,TRUNG_
    #   GIAN!$F$5+1)&" (giờ) ngũ hợp: "&IF(PP2_DINH_LUONG!$E$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$52+1),IF(PP2_DINH_LUONG!$E$44=1,"tranh hợp – không h
    #   óa",IF(PP2_DINH_LUONG!$E$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ",""))&" "&TRIM(IF(AND(TRUE(),P
    #   P2_DINH_LUONG!$C$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUON
    #   G!$C$24=-1,PP2_DINH_LUONG!$C$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) "&IF(PP2_DINH_LUONG!$C$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí
    #   "&IF(PP2_DINH_LUONG!$C$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$C$82,"0%"),"")&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$24<>-1),INDEX(CHI_TEN,TRUNG_
    #   GIAN!$D$6+1)&" (tháng) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$24=-1,PP2_DINH_LUONG!$D$33=1),INDEX(CH
    #   I_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) "&IF(PP2_DINH_LUONG!$D$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$D$82>0,", bản khí tổ
    #   n "&TEXT(PP2_DINH_LUONG!$D$82,"0%"),"")&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$E$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) đã hóa "&INDEX(HANH
    #   _TEN,PP2_DINH_LUONG!$E$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$E$24=-1,PP2_DINH_LUONG!$E$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) "&IF(PP2
    #   _DINH_LUONG!$E$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$E$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$E$82,"0%"),"")&"; 
    #   ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LU
    #   ONG!$F$24+1)&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$24=-1,PP2_DINH_LUONG!$F$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (gi
    #   ờ) "&IF(PP2_DINH_LUONG!$F$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$F$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$F$82,"0
    #   %"),"")&"; ","")))="","không có.",TRIM(TRIM(IF(AND(TRUE(),TRUE(),PP2_DINH_LUONG!$C$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) – "&INDEX(CAN_TEN,TR
    #   UNG_GIAN!$D$5+1)&" (tháng) ngũ hợp: "&IF(PP2_DINH_LUONG!$C$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$52+1),IF(PP2_DINH_LUONG!$C$44=1,"tranh hợp – k
    #   hông hóa",IF(PP2_DINH_LUONG!$C$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ","")&IF(AND(TRUE(),TRUE(
    #   ),PP2_DINH_LUONG!$D$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) – "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) ngũ hợp: "&IF(PP2_DINH_LUONG!$D$47=
    #   1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$52+1),IF(PP2_DINH_LUONG!$D$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$D$46=1,"hợp không hóa – giữ ngũ h
    #   ành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ","")&IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1),INDEX(CA
    #   N_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) – "&INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ngũ hợp: "&IF(PP2_DINH_LUONG!$E$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUO
    #   NG!$E$52+1),IF(PP2_DINH_LUONG!$E$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$E$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339
    #   )","không thành hợp")))&"; ",""))&" "&TRIM(IF(AND(TRUE(),PP2_DINH_LUONG!$C$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) đã hóa "&INDEX(HANH_TEN,PP
    #   2_DINH_LUONG!$C$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$C$24=-1,PP2_DINH_LUONG!$C$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) "&IF(PP2_DINH_LU
    #   ONG!$C$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$C$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$C$82,"0%"),"")&"; ","")&IF
    #   (AND(TRUE(),PP2_DINH_LUONG!$D$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$24+1)&"; ","")&IF(AND(TRUE(
    #   ),PP2_DINH_LUONG!$D$24=-1,PP2_DINH_LUONG!$D$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) "&IF(PP2_DINH_LUONG!$D$32>0,"bị xung","hợp không hóa")&" 
    #   → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$D$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$D$82,"0%"),"")&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$E$24<>-1),IND
    #   EX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$E$24=-1,PP2_DINH_LUONG!$
    #   E$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) "&IF(PP2_DINH_LUONG!$E$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$E$82>
    #   0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$E$82,"0%"),"")&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$24<>-1),INDEX(CHI_TEN,TRUN
    #   G_GIAN!$F$6+1)&" (giờ) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$F$24+1)&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$24=-1,PP2_
    #   DINH_LUONG!$F$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ) "&IF(PP2_DINH_LUONG!$F$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_L
    #   UONG!$F$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$F$82,"0%"),"")&"; ",""))))&IF(NHAP_LIEU!$C$9="Không rõ giờ"," (Không rõ giờ: chỉ xét năm, tháng, ng
    #   ày.)","")
    return X.cat(X.cat('⚠ XUNG – HỢP – HÓA (công cụ tính theo Thiệu Vĩ Hoa – điểm ở trên KHÔNG trừ/cộng theo các quan hệ này, cần tự cân nhắc): ',('không có.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0))),' (tháng) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,3),1.0))) if X.truthy(X.eq(E.v(10,47,3),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,3),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,3),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,3),1.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' (ngày) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,4),1.0))) if X.truthy(X.eq(E.v(10,47,4),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,4),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,4),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0)),' (ngày) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0))),' (giờ) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,5),1.0))) if X.truthy(X.eq(E.v(10,47,5),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,5),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,5),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else ''))),' '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,3),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,3),X.neg(1.0)))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) '),('bị xung' if X.truthy(X.gt(E.v(10,32,3),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,3), '0%')) if X.truthy(X.gt(E.v(10,82,3),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,3),X.neg(1.0)), X.eq(E.v(10,33,3),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,4),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,4),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) '),('bị xung' if X.truthy(X.gt(E.v(10,32,4),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,4), '0%')) if X.truthy(X.gt(E.v(10,82,4),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,4),X.neg(1.0)), X.eq(E.v(10,33,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,5),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,5),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) '),('bị xung' if X.truthy(X.gt(E.v(10,32,5),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,5), '0%')) if X.truthy(X.gt(E.v(10,82,5),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,5),X.neg(1.0)), X.eq(E.v(10,33,5),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,6),1.0))),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ne(E.v(10,24,6),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) '),('bị xung' if X.truthy(X.gt(E.v(10,32,6),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,6), '0%')) if X.truthy(X.gt(E.v(10,82,6),0.0)) else '')),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,24,6),X.neg(1.0)), X.eq(E.v(10,33,6),1.0))) else ''))))),'')) else X.TRIM(X.cat(X.cat(X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0))),' (tháng) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,3),1.0))) if X.truthy(X.eq(E.v(10,47,3),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,3),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,3),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,3),1.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' (ngày) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,4),1.0))) if X.truthy(X.eq(E.v(10,47,4),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,4),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,4),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0)),' (ngày) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0))),' (giờ) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,5),1.0))) if X.truthy(X.eq(E.v(10,47,5),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,5),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,5),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else ''))),' '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,3),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,3),X.neg(1.0)))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) '),('bị xung' if X.truthy(X.gt(E.v(10,32,3),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,3), '0%')) if X.truthy(X.gt(E.v(10,82,3),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,3),X.neg(1.0)), X.eq(E.v(10,33,3),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,4),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,4),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) '),('bị xung' if X.truthy(X.gt(E.v(10,32,4),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,4), '0%')) if X.truthy(X.gt(E.v(10,82,4),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,4),X.neg(1.0)), X.eq(E.v(10,33,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,5),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,5),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) '),('bị xung' if X.truthy(X.gt(E.v(10,32,5),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,5), '0%')) if X.truthy(X.gt(E.v(10,82,5),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,5),X.neg(1.0)), X.eq(E.v(10,33,5),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,6),1.0))),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ne(E.v(10,24,6),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) '),('bị xung' if X.truthy(X.gt(E.v(10,32,6),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,6), '0%')) if X.truthy(X.gt(E.v(10,82,6),0.0)) else '')),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,24,6),X.neg(1.0)), X.eq(E.v(10,33,6),1.0))) else ''))))))),(' (Không rõ giờ: chỉ xét năm, tháng, ngày.)' if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else ''))

def f32(E, r, c):
    # Ô B21
    # = "Xét đắc lệnh (can ngày – chi tháng): "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" – "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)
    return X.cat(X.cat(X.cat('Xét đắc lệnh (can ngày – chi tháng): ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)))

def f33(E, r, c):
    # Ô C21
    # = INDEX(STAGE_TEN,IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2
    #   ,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(
    #   CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(TRUNG_GIAN!$E$5/2)
    #   ,MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(TRUNG_GIAN!$E$5/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(23,5,5),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(23,5,5),2.0))))))) else ''))

def f34(E, r, c):
    # Ô D21
    # = IF(AND(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,
    #   0,8,3)-TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(TRUNG_GIAN!$E$5/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,
    #   2,3,3,2,4)+1,5)=INT(TRUNG_GIAN!$E$5/2))),"Đắc lệnh","Không đắc lệnh")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(23,5,5),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(23,5,5),2.0)))))) else 'Không đắc lệnh')

def f35(E, r, c):
    # Ô E21
    # = IF(AND(IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,
    #   0,8,3)-TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(TRUNG_GIAN!$E$5/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,
    #   2,3,3,2,4)+1,5)=INT(TRUNG_GIAN!$E$5/2))),2,-2)
    return (2.0 if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(23,5,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(23,5,5),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(23,5,5),2.0)))))) else X.neg(2.0))

def f36(E, r, c):
    # Ô B22
    # = "Xét được sinh (can năm, tháng, giờ sinh can ngày)"
    return 'Xét được sinh (can năm, tháng, giờ sinh can ngày)'

def f37(E, r, c):
    # Ô C22
    # = IF((IF(AND(TRUE(),MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)+1,5)=TRUNG_GIAN!$E$7),1,0)+IF(AND(TRUE(),MOD(IF(LEFT(
    #   NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)+1,5)=TRUNG_GIAN!$E$7),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(IF(LEFT(NG
    #   U_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)+1,5)=TRUNG_GIAN!$E$7),1,0))=0,"—",TRIM(IF(AND(TRUE(),MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,
    #   3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) ","")&IF(AND(TRUE(),MOD(IF(LEFT(NGU_HANH
    #   _THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ","")&IF(AND(NOT(NHAP_LIEU
    #   !$C$9="Không rõ giờ"),MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)+1,5)=TRUNG_GIAN!$E$7),INDEX(CAN_TEN,TRUNG_GIAN!$F
    #   $5+1)&" (giờ) ",""))&" sinh "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1))
    return ('—' if X.truthy(X.eq(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add((E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3)),1.0), 5.0),E.v(23,7,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add((E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4)),1.0), 5.0),E.v(23,7,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add((E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6)),1.0), 5.0),E.v(23,7,5)))) else 0.0)),0.0)) else X.cat(X.cat(X.TRIM(X.cat(X.cat((X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add((E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3)),1.0), 5.0),E.v(23,7,5)))) else ''),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) ') if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add((E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4)),1.0), 5.0),E.v(23,7,5)))) else '')),(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0)),' (giờ) ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add((E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6)),1.0), 5.0),E.v(23,7,5)))) else ''))),' sinh '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))))

def f38(E, r, c):
    # Ô D22, E22
    # = (IF(AND(TRUE(),MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)+1,5)=TRUNG_GIAN!$E$7),1,0)+IF(AND(TRUE(),MOD(IF(LEFT(NGU
    #   _HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)+1,5)=TRUNG_GIAN!$E$7),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(IF(LEFT(NGU_H
    #   ANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)+1,5)=TRUNG_GIAN!$E$7),1,0))
    return X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add((E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3)),1.0), 5.0),E.v(23,7,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add((E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4)),1.0), 5.0),E.v(23,7,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add((E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6)),1.0), 5.0),E.v(23,7,5)))) else 0.0))

def f39(E, r, c):
    # Ô B23
    # = "Xét bị tiết (can ngày sinh cho can khác)"
    return 'Xét bị tiết (can ngày sinh cho can khác)'

def f40(E, r, c):
    # Ô C23
    # = IF((IF(AND(TRUE(),MOD(TRUNG_GIAN!$E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)),1,0)+IF(AND(TRUE(),MOD(TRUNG_GI
    #   AN!$E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN
    #   !$E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)),1,0))=0,"—","Bị tiết")
    return ('—' if X.truthy(X.eq(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6))))) else 0.0)),0.0)) else 'Bị tiết')

def f41(E, r, c):
    # Ô D23
    # = (IF(AND(TRUE(),MOD(TRUNG_GIAN!$E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)),1,0)+IF(AND(TRUE(),MOD(TRUNG_GIAN!
    #   $E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$E
    #   $7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)),1,0))
    return X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6))))) else 0.0))

def f42(E, r, c):
    # Ô E23
    # = -(IF(AND(TRUE(),MOD(TRUNG_GIAN!$E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)),1,0)+IF(AND(TRUE(),MOD(TRUNG_GIAN
    #   !$E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$
    #   E$7+1,5)=IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)),1,0))
    return X.neg(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6))))) else 0.0)))

def f43(E, r, c):
    # Ô B24
    # = "Xét bị khắc (can khác khắc can ngày)"
    return 'Xét bị khắc (can khác khắc can ngày)'

def f44(E, r, c):
    # Ô C24
    # = IF((IF(AND(TRUE(),AND(MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)+2,5)=TRUNG_GIAN!$E$7,NOT(AND(LEFT(NGU_HANH_THIEU!
    #   $E$3,3)="PA2",PP2_DINH_LUONG!$C$46=1)))),1,0)+IF(AND(TRUE(),AND(MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)+2,5)=TR
    #   UNG_GIAN!$E$7,NOT(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$46=1)))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(MOD(IF(LEFT(NGU_
    #   HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)+2,5)=TRUNG_GIAN!$E$7,NOT(FALSE()))),1,0))=0,"—","Bị khắc")
    return ('—' if X.truthy(X.eq(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(X.MOD(X.add((E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.eq(E.v(10,46,3),1.0)))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(X.MOD(X.add((E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.eq(E.v(10,46,3),1.0)))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.eq(X.MOD(X.add((E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.FALSE())))) else 0.0)),0.0)) else 'Bị khắc')

def f45(E, r, c):
    # Ô D24
    # = (IF(AND(TRUE(),AND(MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)+2,5)=TRUNG_GIAN!$E$7,NOT(AND(LEFT(NGU_HANH_THIEU!$E$
    #   3,3)="PA2",PP2_DINH_LUONG!$C$46=1)))),1,0)+IF(AND(TRUE(),AND(MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)+2,5)=TRUNG
    #   _GIAN!$E$7,NOT(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$46=1)))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(MOD(IF(LEFT(NGU_HAN
    #   H_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)+2,5)=TRUNG_GIAN!$E$7,NOT(FALSE()))),1,0))
    return X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(X.MOD(X.add((E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.eq(E.v(10,46,3),1.0)))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(X.MOD(X.add((E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.eq(E.v(10,46,3),1.0)))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.eq(X.MOD(X.add((E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.FALSE())))) else 0.0))

def f46(E, r, c):
    # Ô E24
    # = -(IF(AND(TRUE(),AND(MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)+2,5)=TRUNG_GIAN!$E$7,NOT(AND(LEFT(NGU_HANH_THIEU!$E
    #   $3,3)="PA2",PP2_DINH_LUONG!$C$46=1)))),1,0)+IF(AND(TRUE(),AND(MOD(IF(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)+2,5)=TRUN
    #   G_GIAN!$E$7,NOT(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$46=1)))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),AND(MOD(IF(LEFT(NGU_HA
    #   NH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)+2,5)=TRUNG_GIAN!$E$7,NOT(FALSE()))),1,0))
    return X.neg(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(X.MOD(X.add((E.v(10,52,3) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,3)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.eq(E.v(10,46,3),1.0)))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.AND(X.eq(X.MOD(X.add((E.v(10,52,4) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,4)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.eq(E.v(10,46,3),1.0)))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.AND(X.eq(X.MOD(X.add((E.v(10,52,6) if X.truthy(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2')) else E.v(23,7,6)),2.0), 5.0),E.v(23,7,5)), X.NOT(X.FALSE())))) else 0.0)))

def f47(E, r, c):
    # Ô B25
    # = "Thông căn – tàng can CÙNG HÀNH với can ngày trong CHI NGÀY ("&TRUNG_GIAN!$E$11&")"
    return X.cat(X.cat('Thông căn – tàng can CÙNG HÀNH với can ngày trong CHI NGÀY (',E.v(23,11,5)),')')

def f48(E, r, c):
    # Ô C25
    # = IF(IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$E$24<>-1,PP2_DINH_LUONG!$E$24<>TRUNG_GIAN!$E$7),0,(IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7,1,0
    #   )+IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$E$7,1,0)+IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$E$7,1,0)))=0,"Không thông căn","Thông căn")
    return ('Không thông căn' if X.truthy(X.eq((0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,5),X.neg(1.0)), X.ne(E.v(10,24,5),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,5))) else 0.0),(1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,5))) else 0.0)),(1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,5))) else 0.0))),0.0)) else 'Thông căn')

def f49(E, r, c):
    # Ô D25
    # = "N = "&IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$E$24<>-1,PP2_DINH_LUONG!$E$24<>TRUNG_GIAN!$E$7),0,(IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7
    #   ,1,0)+IF(TRUNG_GIAN!$E$22=TRUNG_GIAN!$E$7,1,0)+IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$E$7,1,0)))
    return X.cat('N = ',(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,5),X.neg(1.0)), X.ne(E.v(10,24,5),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,5))) else 0.0),(1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,5))) else 0.0)),(1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,5))) else 0.0))))

def f50(E, r, c):
    # Ô E25
    # = IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$E$24<>-1,PP2_DINH_LUONG!$E$24<>TRUNG_GIAN!$E$7),0,(IF(TRUNG_GIAN!$E$21=TRUNG_GIAN!$E$7,1,0)+I
    #   F(TRUNG_GIAN!$E$22=TRUNG_GIAN!$E$7,1,0)+IF(TRUNG_GIAN!$E$23=TRUNG_GIAN!$E$7,1,0)))
    return (0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,5),X.neg(1.0)), X.ne(E.v(10,24,5),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.eq(E.v(23,21,5),E.v(23,7,5))) else 0.0),(1.0 if X.truthy(X.eq(E.v(23,22,5),E.v(23,7,5))) else 0.0)),(1.0 if X.truthy(X.eq(E.v(23,23,5),E.v(23,7,5))) else 0.0)))

def f51(E, r, c):
    # Ô B26
    # = "Thông gốc – tàng can CÙNG TÊN với can ngày trong CHI NĂM, THÁNG, NGÀY"
    return 'Thông gốc – tàng can CÙNG TÊN với can ngày trong CHI NĂM, THÁNG, NGÀY'

def f52(E, r, c):
    # Ô C26
    # = IF((IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$24<>-1,PP2_DINH_LUONG!$C$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$C$15=TRUNG_GI
    #   AN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$17=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$19=TRUNG_GIAN!$E$5),1,0))+IF(AND(LEFT(NGU_HANH_THIEU!$E
    #   $3,3)="PA2",PP2_DINH_LUONG!$D$24<>-1,PP2_DINH_LUONG!$D$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$D$15=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG
    #   _GIAN!$D$17=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$19=TRUNG_GIAN!$E$5),1,0))+IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$E$24<
    #   >-1,PP2_DINH_LUONG!$E$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$E$15=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$17=TRUNG_GIAN!$E$5),1,0)
    #   +IF(AND(TRUE(),TRUNG_GIAN!$E$19=TRUNG_GIAN!$E$5),1,0)))=0,"Không thông gốc","Thông gốc")
    return ('Không thông gốc' if X.truthy(X.eq(X.add(X.add((0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,3),X.neg(1.0)), X.ne(E.v(10,24,3),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,3),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,3),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,3),E.v(23,5,5)))) else 0.0))),(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,4),X.neg(1.0)), X.ne(E.v(10,24,4),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,4),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,4),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,4),E.v(23,5,5)))) else 0.0)))),(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,5),X.neg(1.0)), X.ne(E.v(10,24,5),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,5),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,5),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,5),E.v(23,5,5)))) else 0.0)))),0.0)) else 'Thông gốc')

def f53(E, r, c):
    # Ô D26
    # = "N = "&(IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$24<>-1,PP2_DINH_LUONG!$C$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$C$15=TRUN
    #   G_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$17=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$19=TRUNG_GIAN!$E$5),1,0))+IF(AND(LEFT(NGU_HANH_THIE
    #   U!$E$3,3)="PA2",PP2_DINH_LUONG!$D$24<>-1,PP2_DINH_LUONG!$D$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$D$15=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),T
    #   RUNG_GIAN!$D$17=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$19=TRUNG_GIAN!$E$5),1,0))+IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$E
    #   $24<>-1,PP2_DINH_LUONG!$E$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$E$15=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$17=TRUNG_GIAN!$E$5),
    #   1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$19=TRUNG_GIAN!$E$5),1,0)))
    return X.cat('N = ',X.add(X.add((0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,3),X.neg(1.0)), X.ne(E.v(10,24,3),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,3),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,3),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,3),E.v(23,5,5)))) else 0.0))),(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,4),X.neg(1.0)), X.ne(E.v(10,24,4),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,4),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,4),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,4),E.v(23,5,5)))) else 0.0)))),(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,5),X.neg(1.0)), X.ne(E.v(10,24,5),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,5),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,5),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,5),E.v(23,5,5)))) else 0.0)))))

def f54(E, r, c):
    # Ô E26
    # = 0.75*(IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$C$24<>-1,PP2_DINH_LUONG!$C$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$C$15=TRUNG_
    #   GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$17=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$C$19=TRUNG_GIAN!$E$5),1,0))+IF(AND(LEFT(NGU_HANH_THIEU!
    #   $E$3,3)="PA2",PP2_DINH_LUONG!$D$24<>-1,PP2_DINH_LUONG!$D$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$D$15=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRU
    #   NG_GIAN!$D$17=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$19=TRUNG_GIAN!$E$5),1,0))+IF(AND(LEFT(NGU_HANH_THIEU!$E$3,3)="PA2",PP2_DINH_LUONG!$E$2
    #   4<>-1,PP2_DINH_LUONG!$E$24<>TRUNG_GIAN!$E$7),0,IF(AND(TRUE(),TRUNG_GIAN!$E$15=TRUNG_GIAN!$E$5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$17=TRUNG_GIAN!$E$5),1,
    #   0)+IF(AND(TRUE(),TRUNG_GIAN!$E$19=TRUNG_GIAN!$E$5),1,0)))
    return X.mul(0.75,X.add(X.add((0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,3),X.neg(1.0)), X.ne(E.v(10,24,3),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,3),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,3),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,3),E.v(23,5,5)))) else 0.0))),(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,4),X.neg(1.0)), X.ne(E.v(10,24,4),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,4),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,4),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,4),E.v(23,5,5)))) else 0.0)))),(0.0 if X.truthy(X.AND(X.eq(X.LEFT(E.v(29,3,5), 3.0),'PA2'), X.ne(E.v(10,24,5),X.neg(1.0)), X.ne(E.v(10,24,5),E.v(23,7,5)))) else X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,15,5),E.v(23,5,5)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,17,5),E.v(23,5,5)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,19,5),E.v(23,5,5)))) else 0.0)))))

def f55(E, r, c):
    # Ô C27
    # = INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)
    return X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0)),' '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0)))

def f56(E, r, c):
    # Ô E27
    # = SUM(E21:E26)
    return X.SUM(E.rg(29,r-6,c,r-1,c))

def f57(E, r, c):
    # Ô F27
    # = IF(E27>=3,"Thân quá vượng",IF(E27>=2,"Thân vượng",IF(E27>=1,"Thân hơi vượng",IF(E27>=0,"Thân trung hòa",IF(E27>-1,"Thân hơi yếu",IF(E27>=-2,"Thân yếu"
    #   ,IF(E27>=-3,"Thân quá yếu","Thân cực yếu")))))))
    return ('Thân quá vượng' if X.truthy(X.ge(E.v(29,r,c-1),3.0)) else ('Thân vượng' if X.truthy(X.ge(E.v(29,r,c-1),2.0)) else ('Thân hơi vượng' if X.truthy(X.ge(E.v(29,r,c-1),1.0)) else ('Thân trung hòa' if X.truthy(X.ge(E.v(29,r,c-1),0.0)) else ('Thân hơi yếu' if X.truthy(X.gt(E.v(29,r,c-1),X.neg(1.0))) else ('Thân yếu' if X.truthy(X.ge(E.v(29,r,c-1),X.neg(2.0))) else ('Thân quá yếu' if X.truthy(X.ge(E.v(29,r,c-1),X.neg(3.0))) else 'Thân cực yếu')))))))

def f58(E, r, c):
    # Ô F28
    # = "Theo mục I: "&IF((E15<=$H$15),"thiếu "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" – bổ sung "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1),IF((E15>=$H$16),"dư "&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$7+1)&" – tiết giảm bằng "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1),"cân bằng – không cần bổ sung hay tiết giảm"))
    return X.cat('Theo mục I: ',(X.cat(X.cat(X.cat('thiếu ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' – bổ sung '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))) if X.truthy(X.le(E.v(29,r-13,c-1),E.v(29,15,8))) else (X.cat(X.cat(X.cat('dư ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' – tiết giảm bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))) if X.truthy(X.ge(E.v(29,r-13,c-1),E.v(29,16,8))) else 'cân bằng – không cần bổ sung hay tiết giảm')))

def f59(E, r, c):
    # Ô F29
    # = UPPER(LEFT(IF(E25>0,"có thông căn","không thông căn")&", "&IF(E26>0,"có thông gốc","không thông gốc"),1))&MID(IF(E25>0,"có thông căn","không thông căn
    #   ")&", "&IF(E26>0,"có thông gốc","không thông gốc"),2,200)&"; "&LOWER(F27)&". "&IF(OR(E25<=0,E26<=0,E27<0),"Cần hóa giải – "&IF((E15<=$H$15),"bổ sung "
    #   &INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1),IF((E15>=$H$16),"tiết giảm "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" bằng "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1),
    #   "ngũ hành cân bằng, không cần bổ sung hay tiết giảm"))&"; bổ sung linh phù thông căn, thông gốc để thông khí cho tứ trụ.","Gốc nhật chủ đã thông, khôn
    #   g cần linh phù; "&IF(OR((E15<=$H$15),(E15>=$H$16)),"điều chỉnh ngũ hành theo mục I: "&IF((E15<=$H$15),"bổ sung "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1),IF(
    #   (E15>=$H$16),"tiết giảm "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" bằng "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1),"ngũ hành cân bằng, không cần bổ sung 
    #   hay tiết giảm"))&".","ngũ hành cân bằng – không cần hóa giải."))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.UPPER(X.LEFT(X.cat(X.cat(('có thông căn' if X.truthy(X.gt(E.v(29,r-4,c-1),0.0)) else 'không thông căn'),', '),('có thông gốc' if X.truthy(X.gt(E.v(29,r-3,c-1),0.0)) else 'không thông gốc')), 1.0)),X.MID(X.cat(X.cat(('có thông căn' if X.truthy(X.gt(E.v(29,r-4,c-1),0.0)) else 'không thông căn'),', '),('có thông gốc' if X.truthy(X.gt(E.v(29,r-3,c-1),0.0)) else 'không thông gốc')), 2.0, 200.0)),'; '),X.LOWER(E.v(29,r-2,c))),'. '),(X.cat(X.cat('Cần hóa giải – ',(X.cat('bổ sung ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))) if X.truthy(X.le(E.v(29,r-14,c-1),E.v(29,15,8))) else (X.cat(X.cat(X.cat('tiết giảm ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))) if X.truthy(X.ge(E.v(29,r-14,c-1),E.v(29,16,8))) else 'ngũ hành cân bằng, không cần bổ sung hay tiết giảm'))),'; bổ sung linh phù thông căn, thông gốc để thông khí cho tứ trụ.') if X.truthy(X.OR(X.le(E.v(29,r-4,c-1),0.0), X.le(E.v(29,r-3,c-1),0.0), X.lt(E.v(29,r-2,c-1),0.0))) else X.cat('Gốc nhật chủ đã thông, không cần linh phù; ',(X.cat(X.cat('điều chỉnh ngũ hành theo mục I: ',(X.cat('bổ sung ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))) if X.truthy(X.le(E.v(29,r-14,c-1),E.v(29,15,8))) else (X.cat(X.cat(X.cat('tiết giảm ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))) if X.truthy(X.ge(E.v(29,r-14,c-1),E.v(29,16,8))) else 'ngũ hành cân bằng, không cần bổ sung hay tiết giảm'))),'.') if X.truthy(X.OR(X.le(E.v(29,r-14,c-1),E.v(29,15,8)), X.ge(E.v(29,r-14,c-1),E.v(29,16,8)))) else 'ngũ hành cân bằng – không cần hóa giải.'))))

def f60(E, r, c):
    # Ô B31
    # = "⚠ XUNG – HỢP – HÓA (điểm ở trên KHÔNG trừ/cộng theo các quan hệ này): "&IF(TRIM(TRIM(IF(AND(TRUE(),TRUE(),PP2_DINH_LUONG!$C$43=1),INDEX(CAN_TEN,TRUNG
    #   _GIAN!$C$5+1)&" (năm) – "&INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ngũ hợp: "&IF(PP2_DINH_LUONG!$C$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$52+1
    #   ),IF(PP2_DINH_LUONG!$C$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$C$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không 
    #   thành hợp")))&"; ","")&IF(AND(TRUE(),TRUE(),PP2_DINH_LUONG!$D$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) – "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" 
    #   (ngày) ngũ hợp: "&IF(PP2_DINH_LUONG!$D$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$52+1),IF(PP2_DINH_LUONG!$D$44=1,"tranh hợp – không hóa",IF(PP2_DIN
    #   H_LUONG!$D$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ","")&IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không
    #    rõ giờ"),PP2_DINH_LUONG!$E$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) – "&INDEX(CAN_TEN,TRUNG_GIAN!$F$5+1)&" (giờ) ngũ hợp: "&IF(PP2_DINH_LUONG!
    #   $E$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$52+1),IF(PP2_DINH_LUONG!$E$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$E$46=1,"hợp không hóa – giữ
    #    ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ",""))&" "&TRIM(IF(AND(TRUE(),PP2_DINH_LUONG!$C$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!
    #   $C$6+1)&" (năm) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$C$24=-1,PP2_DINH_LUONG!$C$33=1),INDEX(CHI_TEN,T
    #   RUNG_GIAN!$C$6+1)&" (năm) "&IF(PP2_DINH_LUONG!$C$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$C$82>0,", bản khí tổn "&TEXT(
    #   PP2_DINH_LUONG!$C$82,"0%"),"")&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) đã hóa "&INDEX(HANH_TEN,PP2
    #   _DINH_LUONG!$D$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$24=-1,PP2_DINH_LUONG!$D$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) "&IF(PP2_DINH_L
    #   UONG!$D$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$D$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$D$82,"0%"),"")&"; ","")&I
    #   F(AND(TRUE(),PP2_DINH_LUONG!$E$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$24+1)&"; ","")&IF(AND(TRUE(
    #   ),PP2_DINH_LUONG!$E$24=-1,PP2_DINH_LUONG!$E$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) "&IF(PP2_DINH_LUONG!$E$32>0,"bị xung","hợp không hóa")&" →
    #    chỉ còn bản khí"&IF(PP2_DINH_LUONG!$E$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$E$82,"0%"),"")&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP
    #   2_DINH_LUONG!$F$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$F$24+1)&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="K
    #   hông rõ giờ"),PP2_DINH_LUONG!$F$24=-1,PP2_DINH_LUONG!$F$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ) "&IF(PP2_DINH_LUONG!$F$32>0,"bị xung","hợp khôn
    #   g hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$F$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$F$82,"0%"),"")&"; ","")))="","không có.",TRIM(TRIM(IF(AND
    #   (TRUE(),TRUE(),PP2_DINH_LUONG!$C$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$C$5+1)&" (năm) – "&INDEX(CAN_TEN,TRUNG_GIAN!$D$5+1)&" (tháng) ngũ hợp: "&IF(PP2_DINH_
    #   LUONG!$C$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$52+1),IF(PP2_DINH_LUONG!$C$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$C$46=1,"hợp không hóa
    #    – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ","")&IF(AND(TRUE(),TRUE(),PP2_DINH_LUONG!$D$43=1),INDEX(CAN_TEN,TRUNG_GIAN
    #   !$D$5+1)&" (tháng) – "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) ngũ hợp: "&IF(PP2_DINH_LUONG!$D$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$52+1),IF
    #   (PP2_DINH_LUONG!$D$44=1,"tranh hợp – không hóa",IF(PP2_DINH_LUONG!$D$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thàn
    #   h hợp")))&"; ","")&IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1),INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" (ngày) – "&INDEX(CAN_TE
    #   N,TRUNG_GIAN!$F$5+1)&" (giờ) ngũ hợp: "&IF(PP2_DINH_LUONG!$E$47=1,"HÓA "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$52+1),IF(PP2_DINH_LUONG!$E$44=1,"tranh hợp –
    #    không hóa",IF(PP2_DINH_LUONG!$E$46=1,"hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)","không thành hợp")))&"; ",""))&" "&TRIM(IF(AND(
    #   TRUE(),PP2_DINH_LUONG!$C$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$24+1)&"; ","")&IF(AND(TRUE(),PP2_D
    #   INH_LUONG!$C$24=-1,PP2_DINH_LUONG!$C$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) "&IF(PP2_DINH_LUONG!$C$32>0,"bị xung","hợp không hóa")&" → chỉ còn
    #    bản khí"&IF(PP2_DINH_LUONG!$C$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$C$82,"0%"),"")&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$24<>-1),INDEX(CHI_TE
    #   N,TRUNG_GIAN!$D$6+1)&" (tháng) đã hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$D$24=-1,PP2_DINH_LUONG!$D$33=1),
    #   INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) "&IF(PP2_DINH_LUONG!$D$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$D$82>0,", bả
    #   n khí tổn "&TEXT(PP2_DINH_LUONG!$D$82,"0%"),"")&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$E$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) đã hóa "&IN
    #   DEX(HANH_TEN,PP2_DINH_LUONG!$E$24+1)&"; ","")&IF(AND(TRUE(),PP2_DINH_LUONG!$E$24=-1,PP2_DINH_LUONG!$E$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) 
    #   "&IF(PP2_DINH_LUONG!$E$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$E$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!$E$82,"0%")
    #   ,"")&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$24<>-1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ) đã hóa "&INDEX(HANH_TEN,PP2
    #   _DINH_LUONG!$F$24+1)&"; ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$F$24=-1,PP2_DINH_LUONG!$F$33=1),INDEX(CHI_TEN,TRUNG_GIAN!$F$6+
    #   1)&" (giờ) "&IF(PP2_DINH_LUONG!$F$32>0,"bị xung","hợp không hóa")&" → chỉ còn bản khí"&IF(PP2_DINH_LUONG!$F$82>0,", bản khí tổn "&TEXT(PP2_DINH_LUONG!
    #   $F$82,"0%"),"")&"; ",""))))&IF(NHAP_LIEU!$C$9="Không rõ giờ"," (Không rõ giờ: chỉ xét năm, tháng, ngày.)","")
    return X.cat(X.cat('⚠ XUNG – HỢP – HÓA (điểm ở trên KHÔNG trừ/cộng theo các quan hệ này): ',('không có.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0))),' (tháng) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,3),1.0))) if X.truthy(X.eq(E.v(10,47,3),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,3),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,3),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,3),1.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' (ngày) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,4),1.0))) if X.truthy(X.eq(E.v(10,47,4),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,4),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,4),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0)),' (ngày) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0))),' (giờ) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,5),1.0))) if X.truthy(X.eq(E.v(10,47,5),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,5),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,5),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else ''))),' '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,3),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,3),X.neg(1.0)))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) '),('bị xung' if X.truthy(X.gt(E.v(10,32,3),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,3), '0%')) if X.truthy(X.gt(E.v(10,82,3),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,3),X.neg(1.0)), X.eq(E.v(10,33,3),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,4),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,4),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) '),('bị xung' if X.truthy(X.gt(E.v(10,32,4),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,4), '0%')) if X.truthy(X.gt(E.v(10,82,4),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,4),X.neg(1.0)), X.eq(E.v(10,33,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,5),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,5),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) '),('bị xung' if X.truthy(X.gt(E.v(10,32,5),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,5), '0%')) if X.truthy(X.gt(E.v(10,82,5),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,5),X.neg(1.0)), X.eq(E.v(10,33,5),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,6),1.0))),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ne(E.v(10,24,6),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) '),('bị xung' if X.truthy(X.gt(E.v(10,32,6),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,6), '0%')) if X.truthy(X.gt(E.v(10,82,6),0.0)) else '')),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,24,6),X.neg(1.0)), X.eq(E.v(10,33,6),1.0))) else ''))))),'')) else X.TRIM(X.cat(X.cat(X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,3),1.0)),' (năm) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0))),' (tháng) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,3),1.0))) if X.truthy(X.eq(E.v(10,47,3),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,3),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,3),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,3),1.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,4),1.0)),' (tháng) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' (ngày) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,4),1.0))) if X.truthy(X.eq(E.v(10,47,4),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,4),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,4),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(E.v(10,43,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0)),' (ngày) – '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,6),1.0))),' (giờ) ngũ hợp: '),(X.cat('HÓA ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,52,5),1.0))) if X.truthy(X.eq(E.v(10,47,5),1.0)) else ('tranh hợp – không hóa' if X.truthy(X.eq(E.v(10,44,5),1.0)) else ('hợp không hóa – giữ ngũ hành gốc, xem như khắc gần (P2 tr.339)' if X.truthy(X.eq(E.v(10,46,5),1.0)) else 'không thành hợp')))),'; ') if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else ''))),' '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,3),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,3),X.neg(1.0)))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),' (năm) '),('bị xung' if X.truthy(X.gt(E.v(10,32,3),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,3), '0%')) if X.truthy(X.gt(E.v(10,82,3),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,3),X.neg(1.0)), X.eq(E.v(10,33,3),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,4),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,4),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),' (tháng) '),('bị xung' if X.truthy(X.gt(E.v(10,32,4),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,4), '0%')) if X.truthy(X.gt(E.v(10,82,4),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,4),X.neg(1.0)), X.eq(E.v(10,33,4),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,5),1.0))),'; ') if X.truthy(X.AND(X.TRUE(), X.ne(E.v(10,24,5),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),' (ngày) '),('bị xung' if X.truthy(X.gt(E.v(10,32,5),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,5), '0%')) if X.truthy(X.gt(E.v(10,82,5),0.0)) else '')),'; ') if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,24,5),X.neg(1.0)), X.eq(E.v(10,33,5),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) đã hóa '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,6),1.0))),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ne(E.v(10,24,6),X.neg(1.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),' (giờ) '),('bị xung' if X.truthy(X.gt(E.v(10,32,6),0.0)) else 'hợp không hóa')),' → chỉ còn bản khí'),(X.cat(', bản khí tổn ',X.TEXT(E.v(10,82,6), '0%')) if X.truthy(X.gt(E.v(10,82,6),0.0)) else '')),'; ') if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,24,6),X.neg(1.0)), X.eq(E.v(10,33,6),1.0))) else ''))))))),(' (Không rõ giờ: chỉ xét năm, tháng, ngày.)' if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else ''))

def f61(E, r, c):
    # Ô C32
    # = IF(OR(E25<=0,E26<=0,E27<0),"Cần hóa giải",IF(OR((E15<=$H$15),(E15>=$H$16)),"Điều chỉnh theo mục I","Không cần hóa giải"))
    return ('Cần hóa giải' if X.truthy(X.OR(X.le(E.v(29,r-7,c+2),0.0), X.le(E.v(29,r-6,c+2),0.0), X.lt(E.v(29,r-5,c+2),0.0))) else ('Điều chỉnh theo mục I' if X.truthy(X.OR(X.le(E.v(29,r-17,c+2),E.v(29,15,8)), X.ge(E.v(29,r-17,c+2),E.v(29,16,8)))) else 'Không cần hóa giải'))

def f62(E, r, c):
    # Ô F32
    # = "(1) Mục I – cân bằng ngũ hành: "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" "&IF((E15<=$H$15),"thiếu "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" – bổ sung "&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$7+1),IF((E15>=$H$16),"dư "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" – tiết giảm bằng "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1),"c
    #   ân bằng – không cần bổ sung hay tiết giảm"))&". (2) Mục II – gốc nhật chủ: "&LOWER(F27)&"; "&IF(E25>0,"có thông căn","không thông căn")&", "&IF(E26>0,
    #   "có thông gốc","không thông gốc")&". "&IF(OR(E25<=0,E26<=0,E27<0),"→ Cần hóa giải: "&IF((E15<=$H$15),"bổ sung "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1),IF((
    #   E15>=$H$16),"tiết giảm "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" bằng "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1),"ngũ hành cân bằng, không cần bổ sung h
    #   ay tiết giảm"))&", đồng thời bổ sung linh phù thông căn, thông gốc để thông khí cho tứ trụ.","→ Gốc nhật chủ đã thông, không cần linh phù"&IF(OR((E15<
    #   =$H$15),(E15>=$H$16)),"; điều chỉnh ngũ hành theo mục I.","; ngũ hành cân bằng, không cần hóa giải."))&" Đối chiếu Thiệu Vĩ Hoa (DOI_CHIEU): "&DOI_CHI
    #   EU!$C$34&" – "&DOI_CHIEU!$C$37&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('(1) Mục I – cân bằng ngũ hành: ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' '),(X.cat(X.cat(X.cat('thiếu ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' – bổ sung '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))) if X.truthy(X.le(E.v(29,r-17,c-1),E.v(29,15,8))) else (X.cat(X.cat(X.cat('dư ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' – tiết giảm bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))) if X.truthy(X.ge(E.v(29,r-17,c-1),E.v(29,16,8))) else 'cân bằng – không cần bổ sung hay tiết giảm'))),'. (2) Mục II – gốc nhật chủ: '),X.LOWER(E.v(29,r-5,c))),'; '),('có thông căn' if X.truthy(X.gt(E.v(29,r-7,c-1),0.0)) else 'không thông căn')),', '),('có thông gốc' if X.truthy(X.gt(E.v(29,r-6,c-1),0.0)) else 'không thông gốc')),'. '),(X.cat(X.cat('→ Cần hóa giải: ',(X.cat('bổ sung ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))) if X.truthy(X.le(E.v(29,r-17,c-1),E.v(29,15,8))) else (X.cat(X.cat(X.cat('tiết giảm ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),' bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))) if X.truthy(X.ge(E.v(29,r-17,c-1),E.v(29,16,8))) else 'ngũ hành cân bằng, không cần bổ sung hay tiết giảm'))),', đồng thời bổ sung linh phù thông căn, thông gốc để thông khí cho tứ trụ.') if X.truthy(X.OR(X.le(E.v(29,r-7,c-1),0.0), X.le(E.v(29,r-6,c-1),0.0), X.lt(E.v(29,r-5,c-1),0.0))) else X.cat('→ Gốc nhật chủ đã thông, không cần linh phù',('; điều chỉnh ngũ hành theo mục I.' if X.truthy(X.OR(X.le(E.v(29,r-17,c-1),E.v(29,15,8)), X.ge(E.v(29,r-17,c-1),E.v(29,16,8)))) else '; ngũ hành cân bằng, không cần hóa giải.')))),' Đối chiếu Thiệu Vĩ Hoa (DOI_CHIEU): '),E.v(7,34,3)),' – '),E.v(7,37,3)),'.')

def f63(E, r, c):
    # Ô B35
    # = DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi; "&DAI_VAN!C37&"–"&DAI_VAN!C38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r,c+1),' ('),E.v(14,r+1,c+1)),'–'),X.add(E.v(14,r+1,c+1),9.0)),' tuổi; '),E.v(14,r+2,c+1)),'–'),E.v(14,r+3,c+1)),')')

def f64(E, r, c):
    # Ô C35
    # = DAI_VAN!$C$39
    return E.v(14,39,3)

def f65(E, r, c):
    # Ô D35, D70
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$C$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$C$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$C$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$C$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,3), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,3), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,3),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,3),2.0))))))) else ''))

def f66(E, r, c):
    # Ô E35
    # = IF(AND(IF(MOD(DAI_VAN!$C$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$C$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$C$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,3), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,3),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,3),2.0)))))) else 'Không')

def f67(E, r, c):
    # Ô F35
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$C$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,3),1.0)),1.0))

def f68(E, r, c):
    # Ô G35
    # = IF(OR(DAI_VAN!$C$39="Chính quan",DAI_VAN!$C$39="Chính tài",DAI_VAN!$C$39="Thiên tài",DAI_VAN!$C$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   C$39="Tỉ kiên",DAI_VAN!$C$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Chính quan'), X.eq(E.v(14,39,3),'Chính tài'), X.eq(E.v(14,39,3),'Thiên tài'), X.eq(E.v(14,39,3),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f69(E, r, c):
    # Ô H35
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$C$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$C$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,3),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,3),1.0)),4.0)) else ''))

def f70(E, r, c):
    # Ô B36
    # = DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi; "&DAI_VAN!D37&"–"&DAI_VAN!D38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-1,c+2),' ('),E.v(14,r,c+2)),'–'),X.add(E.v(14,r,c+2),9.0)),' tuổi; '),E.v(14,r+1,c+2)),'–'),E.v(14,r+2,c+2)),')')

def f71(E, r, c):
    # Ô C36
    # = DAI_VAN!$D$39
    return E.v(14,39,4)

def f72(E, r, c):
    # Ô D36, D71
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$D$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$D$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$D$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$D$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,4),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,4),2.0))))))) else ''))

def f73(E, r, c):
    # Ô E36
    # = IF(AND(IF(MOD(DAI_VAN!$D$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$D$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$D$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,4),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,4),2.0)))))) else 'Không')

def f74(E, r, c):
    # Ô F36
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$D$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,4),1.0)),1.0))

def f75(E, r, c):
    # Ô G36
    # = IF(OR(DAI_VAN!$D$39="Chính quan",DAI_VAN!$D$39="Chính tài",DAI_VAN!$D$39="Thiên tài",DAI_VAN!$D$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   D$39="Tỉ kiên",DAI_VAN!$D$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Chính quan'), X.eq(E.v(14,39,4),'Chính tài'), X.eq(E.v(14,39,4),'Thiên tài'), X.eq(E.v(14,39,4),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f76(E, r, c):
    # Ô H36
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$D$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$D$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,4),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,4),1.0)),4.0)) else ''))

def f77(E, r, c):
    # Ô B37
    # = DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi; "&DAI_VAN!E37&"–"&DAI_VAN!E38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+3),' ('),E.v(14,r-1,c+3)),'–'),X.add(E.v(14,r-1,c+3),9.0)),' tuổi; '),E.v(14,r,c+3)),'–'),E.v(14,r+1,c+3)),')')

def f78(E, r, c):
    # Ô C37
    # = DAI_VAN!$E$39
    return E.v(14,39,5)

def f79(E, r, c):
    # Ô D37, D72
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$E$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$E$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$E$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$E$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,5),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,5),2.0))))))) else ''))

def f80(E, r, c):
    # Ô E37
    # = IF(AND(IF(MOD(DAI_VAN!$E$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$E$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$E$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,5),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,5),2.0)))))) else 'Không')

def f81(E, r, c):
    # Ô F37
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$E$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,5),1.0)),1.0))

def f82(E, r, c):
    # Ô G37
    # = IF(OR(DAI_VAN!$E$39="Chính quan",DAI_VAN!$E$39="Chính tài",DAI_VAN!$E$39="Thiên tài",DAI_VAN!$E$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   E$39="Tỉ kiên",DAI_VAN!$E$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Chính quan'), X.eq(E.v(14,39,5),'Chính tài'), X.eq(E.v(14,39,5),'Thiên tài'), X.eq(E.v(14,39,5),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f83(E, r, c):
    # Ô H37
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$E$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$E$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,5),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,5),1.0)),4.0)) else ''))

def f84(E, r, c):
    # Ô B38
    # = DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi; "&DAI_VAN!F37&"–"&DAI_VAN!F38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+4),' ('),E.v(14,r-2,c+4)),'–'),X.add(E.v(14,r-2,c+4),9.0)),' tuổi; '),E.v(14,r-1,c+4)),'–'),E.v(14,r,c+4)),')')

def f85(E, r, c):
    # Ô C38
    # = DAI_VAN!$F$39
    return E.v(14,39,6)

def f86(E, r, c):
    # Ô D38, D73
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$F$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$F$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$F$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$F$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,6), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,6), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,6),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,6),2.0))))))) else ''))

def f87(E, r, c):
    # Ô E38
    # = IF(AND(IF(MOD(DAI_VAN!$F$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$F$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$F$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,6), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,6),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,6),2.0)))))) else 'Không')

def f88(E, r, c):
    # Ô F38
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$F$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,6),1.0)),1.0))

def f89(E, r, c):
    # Ô G38
    # = IF(OR(DAI_VAN!$F$39="Chính quan",DAI_VAN!$F$39="Chính tài",DAI_VAN!$F$39="Thiên tài",DAI_VAN!$F$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   F$39="Tỉ kiên",DAI_VAN!$F$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Chính quan'), X.eq(E.v(14,39,6),'Chính tài'), X.eq(E.v(14,39,6),'Thiên tài'), X.eq(E.v(14,39,6),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f90(E, r, c):
    # Ô H38
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$F$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$F$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,6),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,6),1.0)),4.0)) else ''))

def f91(E, r, c):
    # Ô B39
    # = DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi; "&DAI_VAN!G37&"–"&DAI_VAN!G38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-4,c+5),' ('),E.v(14,r-3,c+5)),'–'),X.add(E.v(14,r-3,c+5),9.0)),' tuổi; '),E.v(14,r-2,c+5)),'–'),E.v(14,r-1,c+5)),')')

def f92(E, r, c):
    # Ô C39
    # = DAI_VAN!$G$39
    return E.v(14,39,7)

def f93(E, r, c):
    # Ô D39, D74
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$G$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$G$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$G$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$G$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,7), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,7), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,7),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,7),2.0))))))) else ''))

def f94(E, r, c):
    # Ô E39
    # = IF(AND(IF(MOD(DAI_VAN!$G$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$G$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$G$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,7), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,7),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,7),2.0)))))) else 'Không')

def f95(E, r, c):
    # Ô F39
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$G$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,7),1.0)),1.0))

def f96(E, r, c):
    # Ô G39
    # = IF(OR(DAI_VAN!$G$39="Chính quan",DAI_VAN!$G$39="Chính tài",DAI_VAN!$G$39="Thiên tài",DAI_VAN!$G$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   G$39="Tỉ kiên",DAI_VAN!$G$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Chính quan'), X.eq(E.v(14,39,7),'Chính tài'), X.eq(E.v(14,39,7),'Thiên tài'), X.eq(E.v(14,39,7),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f97(E, r, c):
    # Ô H39
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$G$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$G$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,7),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,7),1.0)),4.0)) else ''))

def f98(E, r, c):
    # Ô B40
    # = DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi; "&DAI_VAN!H37&"–"&DAI_VAN!H38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-5,c+6),' ('),E.v(14,r-4,c+6)),'–'),X.add(E.v(14,r-4,c+6),9.0)),' tuổi; '),E.v(14,r-3,c+6)),'–'),E.v(14,r-2,c+6)),')')

def f99(E, r, c):
    # Ô C40
    # = DAI_VAN!$H$39
    return E.v(14,39,8)

def f100(E, r, c):
    # Ô D40, D75
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$H$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$H$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$H$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$H$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,8), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,8), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,8),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,8),2.0))))))) else ''))

def f101(E, r, c):
    # Ô E40
    # = IF(AND(IF(MOD(DAI_VAN!$H$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$H$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$H$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,8), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,8),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,8),2.0)))))) else 'Không')

def f102(E, r, c):
    # Ô F40
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$H$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,8),1.0)),1.0))

def f103(E, r, c):
    # Ô G40
    # = IF(OR(DAI_VAN!$H$39="Chính quan",DAI_VAN!$H$39="Chính tài",DAI_VAN!$H$39="Thiên tài",DAI_VAN!$H$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   H$39="Tỉ kiên",DAI_VAN!$H$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Chính quan'), X.eq(E.v(14,39,8),'Chính tài'), X.eq(E.v(14,39,8),'Thiên tài'), X.eq(E.v(14,39,8),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f104(E, r, c):
    # Ô H40
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$H$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$H$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,8),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,8),1.0)),4.0)) else ''))

def f105(E, r, c):
    # Ô B41
    # = DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi; "&DAI_VAN!I37&"–"&DAI_VAN!I38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-6,c+7),' ('),E.v(14,r-5,c+7)),'–'),X.add(E.v(14,r-5,c+7),9.0)),' tuổi; '),E.v(14,r-4,c+7)),'–'),E.v(14,r-3,c+7)),')')

def f106(E, r, c):
    # Ô C41
    # = DAI_VAN!$I$39
    return E.v(14,39,9)

def f107(E, r, c):
    # Ô D41, D76
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$I$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$I$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$I$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$I$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,9), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,9), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,9),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,9),2.0))))))) else ''))

def f108(E, r, c):
    # Ô E41
    # = IF(AND(IF(MOD(DAI_VAN!$I$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$I$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$I$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,9), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,9),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,9),2.0)))))) else 'Không')

def f109(E, r, c):
    # Ô F41
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$I$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,9),1.0)),1.0))

def f110(E, r, c):
    # Ô G41
    # = IF(OR(DAI_VAN!$I$39="Chính quan",DAI_VAN!$I$39="Chính tài",DAI_VAN!$I$39="Thiên tài",DAI_VAN!$I$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   I$39="Tỉ kiên",DAI_VAN!$I$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Chính quan'), X.eq(E.v(14,39,9),'Chính tài'), X.eq(E.v(14,39,9),'Thiên tài'), X.eq(E.v(14,39,9),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f111(E, r, c):
    # Ô H41
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$I$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$I$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,9),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,9),1.0)),4.0)) else ''))

def f112(E, r, c):
    # Ô B42
    # = DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi; "&DAI_VAN!J37&"–"&DAI_VAN!J38&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-7,c+8),' ('),E.v(14,r-6,c+8)),'–'),X.add(E.v(14,r-6,c+8),9.0)),' tuổi; '),E.v(14,r-5,c+8)),'–'),E.v(14,r-4,c+8)),')')

def f113(E, r, c):
    # Ô C42
    # = DAI_VAN!$J$39
    return E.v(14,39,10)

def f114(E, r, c):
    # Ô D42, D77
    # = INDEX(STAGE_TEN,IF(MOD(DAI_VAN!$J$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3
    #   ,5,0,8,3)-TRUNG_GIAN!$D$6,12))+1)&IF(AND(IF(MOD(DAI_VAN!$J$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI
    #   _VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))<=4,NOT(OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$J$33/2),MOD(CHOOSE(TR
    #   UNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)=INT(DAI_VAN!$J$33/2))))," Giả","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,10), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),1.0)),(' Giả' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,10), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.NOT(X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,10),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,10),2.0))))))) else ''))

def f115(E, r, c):
    # Ô E42
    # = IF(AND(IF(MOD(DAI_VAN!$J$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3)
    #   -TRUNG_GIAN!$D$6,12))<=4,OR(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)=INT(DAI_VAN!$J$33/2),MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,
    #   4)+1,5)=INT(DAI_VAN!$J$33/2))),"Đắc lệnh","Không")
    return ('Đắc lệnh' if X.truthy(X.AND(X.le((X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,10), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0)),4.0), X.OR(X.eq(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),X.INT(X.div(E.v(14,33,10),2.0))), X.eq(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),X.INT(X.div(E.v(14,33,10),2.0)))))) else 'Không')

def f116(E, r, c):
    # Ô F42
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!$J$34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,10),1.0)),1.0))

def f117(E, r, c):
    # Ô G42
    # = IF(OR(DAI_VAN!$J$39="Chính quan",DAI_VAN!$J$39="Chính tài",DAI_VAN!$J$39="Thiên tài",DAI_VAN!$J$39="Chính ấn"),"Quan–Tài–Thê–Ấn (tốt)",IF(OR(DAI_VAN!$
    #   J$39="Tỉ kiên",DAI_VAN!$J$39="Thực thần"),"Tỷ–Thực (trung bình)","Sát–Kiêu–Kiếp–Thương (xấu khi vượng)"))
    return ('Quan–Tài–Thê–Ấn (tốt)' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Chính quan'), X.eq(E.v(14,39,10),'Chính tài'), X.eq(E.v(14,39,10),'Thiên tài'), X.eq(E.v(14,39,10),'Chính ấn'))) else ('Tỷ–Thực (trung bình)' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else 'Sát–Kiêu–Kiếp–Thương (xấu khi vượng)'))

def f118(E, r, c):
    # Ô H42
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$J$34+1)+1)&IF(INDEX(M_TS,TRUNG_GIAN!$E$5+1,DAI_VAN!$J$34+1)<=4," – vượng","")
    return X.cat(X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,10),1.0)),1.0)),(' – vượng' if X.truthy(X.le(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,34,10),1.0)),4.0)) else ''))

def f119(E, r, c):
    # Ô B46
    # = " Nhật can: "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"); chi tháng: "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" – mùa "&INDE
    #   X(HANH_TEN,CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1)&" (Vượng: "&INDEX(HANH_TEN,CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1)&"; Tướ
    #   ng: "&INDEX(HANH_TEN,MOD(CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4)+1,5)+1)&")."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' Nhật can: ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'); chi tháng: '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' – mùa '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0))),' (Vượng: '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0))),'; Tướng: '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0),1.0), 5.0),1.0))),').')

def f120(E, r, c):
    # Ô C48, C49, C50, C51, C52, C53
    # = INDEX(CAN_TEN,$X48+1)&" ("&INDEX(HANH_TEN,INT($X48/2)+1)&")"
    return X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(29,r,24),1.0)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),')')

def f121(E, r, c):
    # Ô D48, D49, D50, D51, D52, D53
    # = INDEX(STAGE_TEN,$Y48+1)
    return X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))

def f122(E, r, c):
    # Ô E48, E49, E50, E51, E52, E53
    # = G48*H48
    return X.mul(E.v(29,r,c+2),E.v(29,r,c+3))

def f123(E, r, c):
    # Ô F48, F49, F52, F53
    # = "Thang A: "&D48&" = "&G48&IF(H48=2," × 2 ("&INDEX(HANH_TEN,INT($X48/2)+1)&" "&IF(MOD(INT($X48/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=
    #   0,"vượng","tướng")&" ở tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&")","")&IF(E48=0," – trang không hiển thị","")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Thang A: ',E.v(29,r,c-2)),' = '),E.v(29,r,c+1)),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),' ở tháng '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),')') if X.truthy(X.eq(E.v(29,r,c+2),2.0)) else '')),(' – trang không hiển thị' if X.truthy(X.eq(E.v(29,r,c-1),0.0)) else ''))

def f124(E, r, c):
    # Ô G48, G49, G52, G53
    # = CHOOSE($Y48+1,50,50,45,40,40,10,5,0,5,10,20,30)
    return X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0)

def f125(E, r, c):
    # Ô H48, H49, H50, H51, H52, H53
    # = IF(AND($Y48<=4,OR(MOD(INT($X48/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X48/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,
    #   2,4),5)=1)),2,1)
    return (2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0)

def f126(E, r, c):
    # Ô X48
    # = (2*MOD(INT(TRUNG_GIAN!$E$5/2)+3,5)+1-MOD(TRUNG_GIAN!$E$5,2))
    return X.sub(X.add(X.mul(2.0,X.MOD(X.add(X.INT(X.div(E.v(23,5,5),2.0)),3.0), 5.0)),1.0),X.MOD(E.v(23,5,5), 2.0))

def f127(E, r, c):
    # Ô Y48, Y49, Y50, Y51, Y52, Y53
    # = IF(MOD($X48,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE($X48+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE($X48+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(29,r,24),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(29,r,24), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(29,r,24),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f128(E, r, c):
    # Ô X49
    # = (2*MOD(INT(TRUNG_GIAN!$E$5/2)+4,5)+1-MOD(TRUNG_GIAN!$E$5,2))
    return X.sub(X.add(X.mul(2.0,X.MOD(X.add(X.INT(X.div(E.v(23,5,5),2.0)),4.0), 5.0)),1.0),X.MOD(E.v(23,5,5), 2.0))

def f129(E, r, c):
    # Ô F50, F51
    # = "Thang B: "&D50&" = "&G50&IF(H50=2," × 2 ("&INDEX(HANH_TEN,INT($X50/2)+1)&" "&IF(MOD(INT($X50/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=
    #   0,"vượng","tướng")&" ở tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&")","")&IF(E50=0," – trang không hiển thị","")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Thang B: ',E.v(29,r,c-2)),' = '),E.v(29,r,c+1)),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),' ở tháng '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),')') if X.truthy(X.eq(E.v(29,r,c+2),2.0)) else '')),(' – trang không hiển thị' if X.truthy(X.eq(E.v(29,r,c-1),0.0)) else ''))

def f130(E, r, c):
    # Ô G50, G51
    # = CHOOSE($Y50+1,25,25,20,15,15,5,0,0,0,5,10,15)
    return X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)

def f131(E, r, c):
    # Ô X50
    # = (2*MOD(INT(TRUNG_GIAN!$E$5/2)+0,5)+MOD(TRUNG_GIAN!$E$5,2))
    return X.add(X.mul(2.0,X.MOD(X.add(X.INT(X.div(E.v(23,5,5),2.0)),0.0), 5.0)),X.MOD(E.v(23,5,5), 2.0))

def f132(E, r, c):
    # Ô X51
    # = (2*MOD(INT(TRUNG_GIAN!$E$5/2)+1,5)+MOD(TRUNG_GIAN!$E$5,2))
    return X.add(X.mul(2.0,X.MOD(X.add(X.INT(X.div(E.v(23,5,5),2.0)),1.0), 5.0)),X.MOD(E.v(23,5,5), 2.0))

def f133(E, r, c):
    # Ô X52
    # = (2*MOD(INT(TRUNG_GIAN!$E$5/2)+2,5)+MOD(TRUNG_GIAN!$E$5,2))
    return X.add(X.mul(2.0,X.MOD(X.add(X.INT(X.div(E.v(23,5,5),2.0)),2.0), 5.0)),X.MOD(E.v(23,5,5), 2.0))

def f134(E, r, c):
    # Ô X53
    # = (2*MOD(INT(TRUNG_GIAN!$E$5/2)+2,5)+1-MOD(TRUNG_GIAN!$E$5,2))
    return X.sub(X.add(X.mul(2.0,X.MOD(X.add(X.INT(X.div(E.v(23,5,5),2.0)),2.0), 5.0)),1.0),X.MOD(E.v(23,5,5), 2.0))

def f135(E, r, c):
    # Ô B68
    # = " Mục II (thân vượng/yếu) = "&TEXT(E27,"0.00")&" → X = "&IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))&IF(AND(E27>0,E27<1.5,MOD(E27*2,2)=1),"  (t = 
    #   0,5 hoặc 1,5: chưa có mẫu – tạm làm tròn lên)","")
    return X.cat(X.cat(X.cat(X.cat(' Mục II (thân vượng/yếu) = ',X.TEXT(E.v(29,r-41,c+3), '0.00')),' → X = '),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-41,c+3),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-41,c+3),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-41,c+3), 0.0)))))),('  (t = 0,5 hoặc 1,5: chưa có mẫu – tạm làm tròn lên)' if X.truthy(X.AND(X.gt(E.v(29,r-41,c+3),0.0), X.lt(E.v(29,r-41,c+3),1.5), X.eq(X.MOD(X.mul(E.v(29,r-41,c+3),2.0), 2.0),1.0))) else ''))

def f136(E, r, c):
    # Ô B70
    # = DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-35,c+1),' ('),E.v(14,r-34,c+1)),'–'),X.add(E.v(14,r-34,c+1),9.0)),' tuổi)')

def f137(E, r, c):
    # Ô C70
    # = DAI_VAN!$C$39&" ("&IF(OR(DAI_VAN!$C$39="Tỉ kiên",DAI_VAN!$C$39="Thực thần"),"B",IF(DAI_VAN!$C$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$C$39="Thất 
    #   sát",DAI_VAN!$C$39="Thiên ấn",DAI_VAN!$C$39="Kiếp tài",DAI_VAN!$C$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,3),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,3),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,3),'Thất sát'), X.eq(E.v(14,39,3),'Thiên ấn'), X.eq(E.v(14,39,3),'Kiếp tài'), X.eq(E.v(14,39,3),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f138(E, r, c):
    # Ô E70
    # = IF(OR(DAI_VAN!$C$39="Thất sát",DAI_VAN!$C$39="Thiên ấn",DAI_VAN!$C$39="Kiếp tài",DAI_VAN!$C$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$C$39="Tỉ kiên",D
    #   AI_VAN!$C$39="Thực thần"),"B",IF(DAI_VAN!$C$39="Thương quan","C","A"))="A",CHOOSE($Y70+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$C$39="Tỉ 
    #   kiên",DAI_VAN!$C$39="Thực thần"),"B",IF(DAI_VAN!$C$39="Thương quan","C","A"))="B",CHOOSE($Y70+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y70+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y70<=4,OR(MOD(INT($X70/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X70/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,3),'Thất sát'), X.eq(E.v(14,39,3),'Thiên ấn'), X.eq(E.v(14,39,3),'Kiếp tài'), X.eq(E.v(14,39,3),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,3),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,3),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f139(E, r, c):
    # Ô F70
    # = "Thang "&IF(OR(DAI_VAN!$C$39="Tỉ kiên",DAI_VAN!$C$39="Thực thần"),"B",IF(DAI_VAN!$C$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y70+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$C$39="Tỉ kiên",DAI_VAN!$C$39="Thực thần"),"B",IF(DAI_VAN!$C$39="Thương quan","C","A"))="A",CHOOSE($Y70+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$C$39="Tỉ kiên",DAI_VAN!$C$39="Thực thần"),"B",IF(DAI_VAN!$C$39="Thương quan","C","A"))="B",CHOOSE($Y70+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y70+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y70<=4,OR(MOD(INT($X70/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X70/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X70/2)+1)&" "&IF(MOD(INT($X70/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$C$39="Thất sát",DAI_VAN!$C$39="Thiên ấn",DAI_VAN!$C$39="Kiếp tài",DAI_VA
    #   N!$C$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,3),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,3),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,3),'Tỉ kiên'), X.eq(E.v(14,39,3),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,3),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,3),'Thất sát'), X.eq(E.v(14,39,3),'Thiên ấn'), X.eq(E.v(14,39,3),'Kiếp tài'), X.eq(E.v(14,39,3),'Thương quan'))) else 1.0),0.0)) else ''))

def f140(E, r, c):
    # Ô G70
    # = E70+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-43,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-43,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-43,c-2), 0.0))))))

def f141(E, r, c):
    # Ô H70, H71, H72, H73, H74, H75, H76, H77
    # = IF(G70>=70,"Vượng phát",IF(G70>=45,"Tốt (trên tiêu chuẩn)",IF(G70>=0,"Dưới tiêu chuẩn",IF(G70>=-30,"Kém",IF(G70>=-70,"Phá sản","Dưới vực")))))
    return ('Vượng phát' if X.truthy(X.ge(E.v(29,r,c-1),70.0)) else ('Tốt (trên tiêu chuẩn)' if X.truthy(X.ge(E.v(29,r,c-1),45.0)) else ('Dưới tiêu chuẩn' if X.truthy(X.ge(E.v(29,r,c-1),0.0)) else ('Kém' if X.truthy(X.ge(E.v(29,r,c-1),X.neg(30.0))) else ('Phá sản' if X.truthy(X.ge(E.v(29,r,c-1),X.neg(70.0))) else 'Dưới vực')))))

def f142(E, r, c):
    # Ô X70
    # = DAI_VAN!$C$33
    return E.v(14,33,3)

def f143(E, r, c):
    # Ô Y70
    # = IF(MOD(DAI_VAN!$C$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$C$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,3), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,3),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f144(E, r, c):
    # Ô B71
    # = DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-36,c+2),' ('),E.v(14,r-35,c+2)),'–'),X.add(E.v(14,r-35,c+2),9.0)),' tuổi)')

def f145(E, r, c):
    # Ô C71
    # = DAI_VAN!$D$39&" ("&IF(OR(DAI_VAN!$D$39="Tỉ kiên",DAI_VAN!$D$39="Thực thần"),"B",IF(DAI_VAN!$D$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$D$39="Thất 
    #   sát",DAI_VAN!$D$39="Thiên ấn",DAI_VAN!$D$39="Kiếp tài",DAI_VAN!$D$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,4),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,4),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,4),'Thất sát'), X.eq(E.v(14,39,4),'Thiên ấn'), X.eq(E.v(14,39,4),'Kiếp tài'), X.eq(E.v(14,39,4),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f146(E, r, c):
    # Ô E71
    # = IF(OR(DAI_VAN!$D$39="Thất sát",DAI_VAN!$D$39="Thiên ấn",DAI_VAN!$D$39="Kiếp tài",DAI_VAN!$D$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$D$39="Tỉ kiên",D
    #   AI_VAN!$D$39="Thực thần"),"B",IF(DAI_VAN!$D$39="Thương quan","C","A"))="A",CHOOSE($Y71+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$D$39="Tỉ 
    #   kiên",DAI_VAN!$D$39="Thực thần"),"B",IF(DAI_VAN!$D$39="Thương quan","C","A"))="B",CHOOSE($Y71+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y71+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y71<=4,OR(MOD(INT($X71/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X71/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,4),'Thất sát'), X.eq(E.v(14,39,4),'Thiên ấn'), X.eq(E.v(14,39,4),'Kiếp tài'), X.eq(E.v(14,39,4),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,4),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,4),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f147(E, r, c):
    # Ô F71
    # = "Thang "&IF(OR(DAI_VAN!$D$39="Tỉ kiên",DAI_VAN!$D$39="Thực thần"),"B",IF(DAI_VAN!$D$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y71+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$D$39="Tỉ kiên",DAI_VAN!$D$39="Thực thần"),"B",IF(DAI_VAN!$D$39="Thương quan","C","A"))="A",CHOOSE($Y71+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$D$39="Tỉ kiên",DAI_VAN!$D$39="Thực thần"),"B",IF(DAI_VAN!$D$39="Thương quan","C","A"))="B",CHOOSE($Y71+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y71+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y71<=4,OR(MOD(INT($X71/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X71/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X71/2)+1)&" "&IF(MOD(INT($X71/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$D$39="Thất sát",DAI_VAN!$D$39="Thiên ấn",DAI_VAN!$D$39="Kiếp tài",DAI_VA
    #   N!$D$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,4),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,4),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,4),'Tỉ kiên'), X.eq(E.v(14,39,4),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,4),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,4),'Thất sát'), X.eq(E.v(14,39,4),'Thiên ấn'), X.eq(E.v(14,39,4),'Kiếp tài'), X.eq(E.v(14,39,4),'Thương quan'))) else 1.0),0.0)) else ''))

def f148(E, r, c):
    # Ô G71
    # = E71+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-44,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-44,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-44,c-2), 0.0))))))

def f149(E, r, c):
    # Ô X71
    # = DAI_VAN!$D$33
    return E.v(14,33,4)

def f150(E, r, c):
    # Ô Y71
    # = IF(MOD(DAI_VAN!$D$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$D$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,4), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,4),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f151(E, r, c):
    # Ô B72
    # = DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-37,c+3),' ('),E.v(14,r-36,c+3)),'–'),X.add(E.v(14,r-36,c+3),9.0)),' tuổi)')

def f152(E, r, c):
    # Ô C72
    # = DAI_VAN!$E$39&" ("&IF(OR(DAI_VAN!$E$39="Tỉ kiên",DAI_VAN!$E$39="Thực thần"),"B",IF(DAI_VAN!$E$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$E$39="Thất 
    #   sát",DAI_VAN!$E$39="Thiên ấn",DAI_VAN!$E$39="Kiếp tài",DAI_VAN!$E$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,5),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,5),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,5),'Thất sát'), X.eq(E.v(14,39,5),'Thiên ấn'), X.eq(E.v(14,39,5),'Kiếp tài'), X.eq(E.v(14,39,5),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f153(E, r, c):
    # Ô E72
    # = IF(OR(DAI_VAN!$E$39="Thất sát",DAI_VAN!$E$39="Thiên ấn",DAI_VAN!$E$39="Kiếp tài",DAI_VAN!$E$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$E$39="Tỉ kiên",D
    #   AI_VAN!$E$39="Thực thần"),"B",IF(DAI_VAN!$E$39="Thương quan","C","A"))="A",CHOOSE($Y72+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$E$39="Tỉ 
    #   kiên",DAI_VAN!$E$39="Thực thần"),"B",IF(DAI_VAN!$E$39="Thương quan","C","A"))="B",CHOOSE($Y72+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y72+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y72<=4,OR(MOD(INT($X72/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X72/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,5),'Thất sát'), X.eq(E.v(14,39,5),'Thiên ấn'), X.eq(E.v(14,39,5),'Kiếp tài'), X.eq(E.v(14,39,5),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,5),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,5),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f154(E, r, c):
    # Ô F72
    # = "Thang "&IF(OR(DAI_VAN!$E$39="Tỉ kiên",DAI_VAN!$E$39="Thực thần"),"B",IF(DAI_VAN!$E$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y72+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$E$39="Tỉ kiên",DAI_VAN!$E$39="Thực thần"),"B",IF(DAI_VAN!$E$39="Thương quan","C","A"))="A",CHOOSE($Y72+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$E$39="Tỉ kiên",DAI_VAN!$E$39="Thực thần"),"B",IF(DAI_VAN!$E$39="Thương quan","C","A"))="B",CHOOSE($Y72+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y72+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y72<=4,OR(MOD(INT($X72/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X72/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X72/2)+1)&" "&IF(MOD(INT($X72/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$E$39="Thất sát",DAI_VAN!$E$39="Thiên ấn",DAI_VAN!$E$39="Kiếp tài",DAI_VA
    #   N!$E$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,5),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,5),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,5),'Tỉ kiên'), X.eq(E.v(14,39,5),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,5),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,5),'Thất sát'), X.eq(E.v(14,39,5),'Thiên ấn'), X.eq(E.v(14,39,5),'Kiếp tài'), X.eq(E.v(14,39,5),'Thương quan'))) else 1.0),0.0)) else ''))

def f155(E, r, c):
    # Ô G72
    # = E72+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-45,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-45,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-45,c-2), 0.0))))))

def f156(E, r, c):
    # Ô X72
    # = DAI_VAN!$E$33
    return E.v(14,33,5)

def f157(E, r, c):
    # Ô Y72
    # = IF(MOD(DAI_VAN!$E$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$E$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,5), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,5),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f158(E, r, c):
    # Ô B73
    # = DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-38,c+4),' ('),E.v(14,r-37,c+4)),'–'),X.add(E.v(14,r-37,c+4),9.0)),' tuổi)')

def f159(E, r, c):
    # Ô C73
    # = DAI_VAN!$F$39&" ("&IF(OR(DAI_VAN!$F$39="Tỉ kiên",DAI_VAN!$F$39="Thực thần"),"B",IF(DAI_VAN!$F$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$F$39="Thất 
    #   sát",DAI_VAN!$F$39="Thiên ấn",DAI_VAN!$F$39="Kiếp tài",DAI_VAN!$F$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,6),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,6),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,6),'Thất sát'), X.eq(E.v(14,39,6),'Thiên ấn'), X.eq(E.v(14,39,6),'Kiếp tài'), X.eq(E.v(14,39,6),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f160(E, r, c):
    # Ô E73
    # = IF(OR(DAI_VAN!$F$39="Thất sát",DAI_VAN!$F$39="Thiên ấn",DAI_VAN!$F$39="Kiếp tài",DAI_VAN!$F$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$F$39="Tỉ kiên",D
    #   AI_VAN!$F$39="Thực thần"),"B",IF(DAI_VAN!$F$39="Thương quan","C","A"))="A",CHOOSE($Y73+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$F$39="Tỉ 
    #   kiên",DAI_VAN!$F$39="Thực thần"),"B",IF(DAI_VAN!$F$39="Thương quan","C","A"))="B",CHOOSE($Y73+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y73+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y73<=4,OR(MOD(INT($X73/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X73/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,6),'Thất sát'), X.eq(E.v(14,39,6),'Thiên ấn'), X.eq(E.v(14,39,6),'Kiếp tài'), X.eq(E.v(14,39,6),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,6),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,6),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f161(E, r, c):
    # Ô F73
    # = "Thang "&IF(OR(DAI_VAN!$F$39="Tỉ kiên",DAI_VAN!$F$39="Thực thần"),"B",IF(DAI_VAN!$F$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y73+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$F$39="Tỉ kiên",DAI_VAN!$F$39="Thực thần"),"B",IF(DAI_VAN!$F$39="Thương quan","C","A"))="A",CHOOSE($Y73+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$F$39="Tỉ kiên",DAI_VAN!$F$39="Thực thần"),"B",IF(DAI_VAN!$F$39="Thương quan","C","A"))="B",CHOOSE($Y73+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y73+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y73<=4,OR(MOD(INT($X73/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X73/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X73/2)+1)&" "&IF(MOD(INT($X73/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$F$39="Thất sát",DAI_VAN!$F$39="Thiên ấn",DAI_VAN!$F$39="Kiếp tài",DAI_VA
    #   N!$F$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,6),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,6),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,6),'Tỉ kiên'), X.eq(E.v(14,39,6),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,6),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,6),'Thất sát'), X.eq(E.v(14,39,6),'Thiên ấn'), X.eq(E.v(14,39,6),'Kiếp tài'), X.eq(E.v(14,39,6),'Thương quan'))) else 1.0),0.0)) else ''))

def f162(E, r, c):
    # Ô G73
    # = E73+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-46,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-46,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-46,c-2), 0.0))))))

def f163(E, r, c):
    # Ô X73
    # = DAI_VAN!$F$33
    return E.v(14,33,6)

def f164(E, r, c):
    # Ô Y73
    # = IF(MOD(DAI_VAN!$F$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$F$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,6), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,6),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f165(E, r, c):
    # Ô B74
    # = DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-39,c+5),' ('),E.v(14,r-38,c+5)),'–'),X.add(E.v(14,r-38,c+5),9.0)),' tuổi)')

def f166(E, r, c):
    # Ô C74
    # = DAI_VAN!$G$39&" ("&IF(OR(DAI_VAN!$G$39="Tỉ kiên",DAI_VAN!$G$39="Thực thần"),"B",IF(DAI_VAN!$G$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$G$39="Thất 
    #   sát",DAI_VAN!$G$39="Thiên ấn",DAI_VAN!$G$39="Kiếp tài",DAI_VAN!$G$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,7),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,7),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,7),'Thất sát'), X.eq(E.v(14,39,7),'Thiên ấn'), X.eq(E.v(14,39,7),'Kiếp tài'), X.eq(E.v(14,39,7),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f167(E, r, c):
    # Ô E74
    # = IF(OR(DAI_VAN!$G$39="Thất sát",DAI_VAN!$G$39="Thiên ấn",DAI_VAN!$G$39="Kiếp tài",DAI_VAN!$G$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$G$39="Tỉ kiên",D
    #   AI_VAN!$G$39="Thực thần"),"B",IF(DAI_VAN!$G$39="Thương quan","C","A"))="A",CHOOSE($Y74+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$G$39="Tỉ 
    #   kiên",DAI_VAN!$G$39="Thực thần"),"B",IF(DAI_VAN!$G$39="Thương quan","C","A"))="B",CHOOSE($Y74+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y74+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y74<=4,OR(MOD(INT($X74/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X74/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,7),'Thất sát'), X.eq(E.v(14,39,7),'Thiên ấn'), X.eq(E.v(14,39,7),'Kiếp tài'), X.eq(E.v(14,39,7),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,7),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,7),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f168(E, r, c):
    # Ô F74
    # = "Thang "&IF(OR(DAI_VAN!$G$39="Tỉ kiên",DAI_VAN!$G$39="Thực thần"),"B",IF(DAI_VAN!$G$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y74+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$G$39="Tỉ kiên",DAI_VAN!$G$39="Thực thần"),"B",IF(DAI_VAN!$G$39="Thương quan","C","A"))="A",CHOOSE($Y74+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$G$39="Tỉ kiên",DAI_VAN!$G$39="Thực thần"),"B",IF(DAI_VAN!$G$39="Thương quan","C","A"))="B",CHOOSE($Y74+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y74+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y74<=4,OR(MOD(INT($X74/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X74/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X74/2)+1)&" "&IF(MOD(INT($X74/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$G$39="Thất sát",DAI_VAN!$G$39="Thiên ấn",DAI_VAN!$G$39="Kiếp tài",DAI_VA
    #   N!$G$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,7),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,7),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,7),'Tỉ kiên'), X.eq(E.v(14,39,7),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,7),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,7),'Thất sát'), X.eq(E.v(14,39,7),'Thiên ấn'), X.eq(E.v(14,39,7),'Kiếp tài'), X.eq(E.v(14,39,7),'Thương quan'))) else 1.0),0.0)) else ''))

def f169(E, r, c):
    # Ô G74
    # = E74+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-47,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-47,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-47,c-2), 0.0))))))

def f170(E, r, c):
    # Ô X74
    # = DAI_VAN!$G$33
    return E.v(14,33,7)

def f171(E, r, c):
    # Ô Y74
    # = IF(MOD(DAI_VAN!$G$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$G$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,7), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,7),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f172(E, r, c):
    # Ô B75
    # = DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-40,c+6),' ('),E.v(14,r-39,c+6)),'–'),X.add(E.v(14,r-39,c+6),9.0)),' tuổi)')

def f173(E, r, c):
    # Ô C75
    # = DAI_VAN!$H$39&" ("&IF(OR(DAI_VAN!$H$39="Tỉ kiên",DAI_VAN!$H$39="Thực thần"),"B",IF(DAI_VAN!$H$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$H$39="Thất 
    #   sát",DAI_VAN!$H$39="Thiên ấn",DAI_VAN!$H$39="Kiếp tài",DAI_VAN!$H$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,8),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,8),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,8),'Thất sát'), X.eq(E.v(14,39,8),'Thiên ấn'), X.eq(E.v(14,39,8),'Kiếp tài'), X.eq(E.v(14,39,8),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f174(E, r, c):
    # Ô E75
    # = IF(OR(DAI_VAN!$H$39="Thất sát",DAI_VAN!$H$39="Thiên ấn",DAI_VAN!$H$39="Kiếp tài",DAI_VAN!$H$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$H$39="Tỉ kiên",D
    #   AI_VAN!$H$39="Thực thần"),"B",IF(DAI_VAN!$H$39="Thương quan","C","A"))="A",CHOOSE($Y75+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$H$39="Tỉ 
    #   kiên",DAI_VAN!$H$39="Thực thần"),"B",IF(DAI_VAN!$H$39="Thương quan","C","A"))="B",CHOOSE($Y75+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y75+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y75<=4,OR(MOD(INT($X75/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X75/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,8),'Thất sát'), X.eq(E.v(14,39,8),'Thiên ấn'), X.eq(E.v(14,39,8),'Kiếp tài'), X.eq(E.v(14,39,8),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,8),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,8),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f175(E, r, c):
    # Ô F75
    # = "Thang "&IF(OR(DAI_VAN!$H$39="Tỉ kiên",DAI_VAN!$H$39="Thực thần"),"B",IF(DAI_VAN!$H$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y75+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$H$39="Tỉ kiên",DAI_VAN!$H$39="Thực thần"),"B",IF(DAI_VAN!$H$39="Thương quan","C","A"))="A",CHOOSE($Y75+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$H$39="Tỉ kiên",DAI_VAN!$H$39="Thực thần"),"B",IF(DAI_VAN!$H$39="Thương quan","C","A"))="B",CHOOSE($Y75+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y75+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y75<=4,OR(MOD(INT($X75/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X75/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X75/2)+1)&" "&IF(MOD(INT($X75/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$H$39="Thất sát",DAI_VAN!$H$39="Thiên ấn",DAI_VAN!$H$39="Kiếp tài",DAI_VA
    #   N!$H$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,8),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,8),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,8),'Tỉ kiên'), X.eq(E.v(14,39,8),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,8),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,8),'Thất sát'), X.eq(E.v(14,39,8),'Thiên ấn'), X.eq(E.v(14,39,8),'Kiếp tài'), X.eq(E.v(14,39,8),'Thương quan'))) else 1.0),0.0)) else ''))

def f176(E, r, c):
    # Ô G75
    # = E75+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-48,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-48,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-48,c-2), 0.0))))))

def f177(E, r, c):
    # Ô X75
    # = DAI_VAN!$H$33
    return E.v(14,33,8)

def f178(E, r, c):
    # Ô Y75
    # = IF(MOD(DAI_VAN!$H$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$H$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,8), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,8),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f179(E, r, c):
    # Ô B76
    # = DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-41,c+7),' ('),E.v(14,r-40,c+7)),'–'),X.add(E.v(14,r-40,c+7),9.0)),' tuổi)')

def f180(E, r, c):
    # Ô C76
    # = DAI_VAN!$I$39&" ("&IF(OR(DAI_VAN!$I$39="Tỉ kiên",DAI_VAN!$I$39="Thực thần"),"B",IF(DAI_VAN!$I$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$I$39="Thất 
    #   sát",DAI_VAN!$I$39="Thiên ấn",DAI_VAN!$I$39="Kiếp tài",DAI_VAN!$I$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,9),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,9),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,9),'Thất sát'), X.eq(E.v(14,39,9),'Thiên ấn'), X.eq(E.v(14,39,9),'Kiếp tài'), X.eq(E.v(14,39,9),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f181(E, r, c):
    # Ô E76
    # = IF(OR(DAI_VAN!$I$39="Thất sát",DAI_VAN!$I$39="Thiên ấn",DAI_VAN!$I$39="Kiếp tài",DAI_VAN!$I$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$I$39="Tỉ kiên",D
    #   AI_VAN!$I$39="Thực thần"),"B",IF(DAI_VAN!$I$39="Thương quan","C","A"))="A",CHOOSE($Y76+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$I$39="Tỉ 
    #   kiên",DAI_VAN!$I$39="Thực thần"),"B",IF(DAI_VAN!$I$39="Thương quan","C","A"))="B",CHOOSE($Y76+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y76+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y76<=4,OR(MOD(INT($X76/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X76/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,9),'Thất sát'), X.eq(E.v(14,39,9),'Thiên ấn'), X.eq(E.v(14,39,9),'Kiếp tài'), X.eq(E.v(14,39,9),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,9),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,9),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f182(E, r, c):
    # Ô F76
    # = "Thang "&IF(OR(DAI_VAN!$I$39="Tỉ kiên",DAI_VAN!$I$39="Thực thần"),"B",IF(DAI_VAN!$I$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y76+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$I$39="Tỉ kiên",DAI_VAN!$I$39="Thực thần"),"B",IF(DAI_VAN!$I$39="Thương quan","C","A"))="A",CHOOSE($Y76+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$I$39="Tỉ kiên",DAI_VAN!$I$39="Thực thần"),"B",IF(DAI_VAN!$I$39="Thương quan","C","A"))="B",CHOOSE($Y76+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y76+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y76<=4,OR(MOD(INT($X76/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X76/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X76/2)+1)&" "&IF(MOD(INT($X76/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$I$39="Thất sát",DAI_VAN!$I$39="Thiên ấn",DAI_VAN!$I$39="Kiếp tài",DAI_VA
    #   N!$I$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,9),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,9),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,9),'Tỉ kiên'), X.eq(E.v(14,39,9),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,9),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,9),'Thất sát'), X.eq(E.v(14,39,9),'Thiên ấn'), X.eq(E.v(14,39,9),'Kiếp tài'), X.eq(E.v(14,39,9),'Thương quan'))) else 1.0),0.0)) else ''))

def f183(E, r, c):
    # Ô G76
    # = E76+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-49,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-49,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-49,c-2), 0.0))))))

def f184(E, r, c):
    # Ô X76
    # = DAI_VAN!$I$33
    return E.v(14,33,9)

def f185(E, r, c):
    # Ô Y76
    # = IF(MOD(DAI_VAN!$I$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$I$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,9), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,9),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

def f186(E, r, c):
    # Ô B77
    # = DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-42,c+8),' ('),E.v(14,r-41,c+8)),'–'),X.add(E.v(14,r-41,c+8),9.0)),' tuổi)')

def f187(E, r, c):
    # Ô C77
    # = DAI_VAN!$J$39&" ("&IF(OR(DAI_VAN!$J$39="Tỉ kiên",DAI_VAN!$J$39="Thực thần"),"B",IF(DAI_VAN!$J$39="Thương quan","C","A"))&IF(IF(OR(DAI_VAN!$J$39="Thất 
    #   sát",DAI_VAN!$J$39="Thiên ấn",DAI_VAN!$J$39="Kiếp tài",DAI_VAN!$J$39="Thương quan"),-1,1)<0,", −",", +")&")"
    return X.cat(X.cat(X.cat(X.cat(E.v(14,39,10),' ('),('B' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,10),'Thương quan')) else 'A'))),(', −' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,10),'Thất sát'), X.eq(E.v(14,39,10),'Thiên ấn'), X.eq(E.v(14,39,10),'Kiếp tài'), X.eq(E.v(14,39,10),'Thương quan'))) else 1.0),0.0)) else ', +')),')')

def f188(E, r, c):
    # Ô E77
    # = IF(OR(DAI_VAN!$J$39="Thất sát",DAI_VAN!$J$39="Thiên ấn",DAI_VAN!$J$39="Kiếp tài",DAI_VAN!$J$39="Thương quan"),-1,1)*IF(IF(OR(DAI_VAN!$J$39="Tỉ kiên",D
    #   AI_VAN!$J$39="Thực thần"),"B",IF(DAI_VAN!$J$39="Thương quan","C","A"))="A",CHOOSE($Y77+1,50,50,45,40,40,10,5,0,5,10,20,30),IF(IF(OR(DAI_VAN!$J$39="Tỉ 
    #   kiên",DAI_VAN!$J$39="Thực thần"),"B",IF(DAI_VAN!$J$39="Thương quan","C","A"))="B",CHOOSE($Y77+1,25,25,20,15,15,5,0,0,0,5,10,15),CHOOSE($Y77+1,45,45,40
    #   ,35,35,5,0,0,0,5,10,15)))*IF(AND($Y77<=4,OR(MOD(INT($X77/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(INT($X77/2)-CHOOSE(TRUNG_GIAN!$
    #   D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1)),2,1)
    return X.mul(X.mul((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,10),'Thất sát'), X.eq(E.v(14,39,10),'Thiên ấn'), X.eq(E.v(14,39,10),'Kiếp tài'), X.eq(E.v(14,39,10),'Thương quan'))) else 1.0),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,10),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,10),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(2.0 if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else 1.0))

def f189(E, r, c):
    # Ô F77
    # = "Thang "&IF(OR(DAI_VAN!$J$39="Tỉ kiên",DAI_VAN!$J$39="Thực thần"),"B",IF(DAI_VAN!$J$39="Thương quan","C","A"))&": "&INDEX(STAGE_TEN,$Y77+1)&" = "&IF(I
    #   F(OR(DAI_VAN!$J$39="Tỉ kiên",DAI_VAN!$J$39="Thực thần"),"B",IF(DAI_VAN!$J$39="Thương quan","C","A"))="A",CHOOSE($Y77+1,50,50,45,40,40,10,5,0,5,10,20,3
    #   0),IF(IF(OR(DAI_VAN!$J$39="Tỉ kiên",DAI_VAN!$J$39="Thực thần"),"B",IF(DAI_VAN!$J$39="Thương quan","C","A"))="B",CHOOSE($Y77+1,25,25,20,15,15,5,0,0,0,5
    #   ,10,15),CHOOSE($Y77+1,45,45,40,35,35,5,0,0,0,5,10,15)))&IF(AND($Y77<=4,OR(MOD(INT($X77/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,MOD(I
    #   NT($X77/2)-CHOOSE(TRUNG_GIAN!$D$6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=1))," × 2 ("&INDEX(HANH_TEN,INT($X77/2)+1)&" "&IF(MOD(INT($X77/2)-CHOOSE(TRUNG_GIAN!$D
    #   $6+1,4,2,0,0,2,1,1,2,3,3,2,4),5)=0,"vượng","tướng")&")","")&IF(IF(OR(DAI_VAN!$J$39="Thất sát",DAI_VAN!$J$39="Thiên ấn",DAI_VAN!$J$39="Kiếp tài",DAI_VA
    #   N!$J$39="Thương quan"),-1,1)<0," – lấy dấu âm","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thang ',('B' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,10),'Thương quan')) else 'A'))),': '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(29,r,25),1.0))),' = '),(X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 50.0, lambda: 50.0, lambda: 45.0, lambda: 40.0, lambda: 40.0, lambda: 10.0, lambda: 5.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 20.0, lambda: 30.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,10),'Thương quan')) else 'A')),'A')) else (X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 25.0, lambda: 25.0, lambda: 20.0, lambda: 15.0, lambda: 15.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0) if X.truthy(X.eq(('B' if X.truthy(X.OR(X.eq(E.v(14,39,10),'Tỉ kiên'), X.eq(E.v(14,39,10),'Thực thần'))) else ('C' if X.truthy(X.eq(E.v(14,39,10),'Thương quan')) else 'A')),'B')) else X.CHOOSE(X.add(E.v(29,r,25),1.0), lambda: 45.0, lambda: 45.0, lambda: 40.0, lambda: 35.0, lambda: 35.0, lambda: 5.0, lambda: 0.0, lambda: 0.0, lambda: 0.0, lambda: 5.0, lambda: 10.0, lambda: 15.0)))),(X.cat(X.cat(X.cat(X.cat(' × 2 (',X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(29,r,24),2.0)),1.0))),' '),('vượng' if X.truthy(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0)) else 'tướng')),')') if X.truthy(X.AND(X.le(E.v(29,r,25),4.0), X.OR(X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(E.v(29,r,24),2.0)),X.CHOOSE(X.add(E.v(23,6,4),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 0.0, lambda: 2.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 2.0, lambda: 4.0)), 5.0),1.0)))) else '')),(' – lấy dấu âm' if X.truthy(X.lt((X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(14,39,10),'Thất sát'), X.eq(E.v(14,39,10),'Thiên ấn'), X.eq(E.v(14,39,10),'Kiếp tài'), X.eq(E.v(14,39,10),'Thương quan'))) else 1.0),0.0)) else ''))

def f190(E, r, c):
    # Ô G77
    # = E77+IF(E27<0,-50,IF(E27>=2,-1,-30+10*ROUND(E27,0)))
    return X.add(E.v(29,r,c-2),(X.neg(50.0) if X.truthy(X.lt(E.v(29,r-50,c-2),0.0)) else (X.neg(1.0) if X.truthy(X.ge(E.v(29,r-50,c-2),2.0)) else X.add(X.neg(30.0),X.mul(10.0,X.ROUND(E.v(29,r-50,c-2), 0.0))))))

def f191(E, r, c):
    # Ô X77
    # = DAI_VAN!$J$33
    return E.v(14,33,10)

def f192(E, r, c):
    # Ô Y77
    # = IF(MOD(DAI_VAN!$J$33,2)=0,MOD(TRUNG_GIAN!$D$6-CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3),12),MOD(CHOOSE(DAI_VAN!$J$33+1,11,6,2,9,8,3,5,0,8,3)-TRUNG_
    #   GIAN!$D$6,12))
    return (X.MOD(X.sub(E.v(23,6,4),X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0)), 12.0) if X.truthy(X.eq(X.MOD(E.v(14,33,10), 2.0),0.0)) else X.MOD(X.sub(X.CHOOSE(X.add(E.v(14,33,10),1.0), lambda: 11.0, lambda: 6.0, lambda: 2.0, lambda: 9.0, lambda: 8.0, lambda: 3.0, lambda: 5.0, lambda: 0.0, lambda: 8.0, lambda: 3.0),E.v(23,6,4)), 12.0))

CELLS = {
    'G3': (f0, 3, 7),
    'B5': (f1, 5, 2),
    'B7': (f2, 7, 2),
    'C7': (f3, 7, 3),
    'D7': (f4, 7, 4),
    'E7': (f5, 7, 5),
    'B8': (f6, 8, 2),
    'B9': (f7, 9, 2),
    'C9': (f8, 9, 3),
    'D9': (f9, 9, 4),
    'B10': (f10, 10, 2),
    'C10': (f11, 10, 3),
    'D10': (f12, 10, 4),
    'B11': (f13, 11, 2),
    'C11': (f14, 11, 3),
    'D11': (f15, 11, 4),
    'B12': (f16, 12, 2),
    'E12': (f17, 12, 5),
    'B13': (f18, 13, 2),
    'C13': (f19, 13, 3),
    'D13': (f20, 13, 4),
    'E13': (f21, 13, 5),
    'B14': (f22, 14, 2),
    'C14': (f23, 14, 3),
    'D14': (f24, 14, 4),
    'E14': (f25, 14, 5),
    'C15': (f26, 15, 3),
    'E15': (f27, 15, 5),
    'F15': (f28, 15, 6),
    'C16': (f29, 16, 3),
    'F16': (f30, 16, 6),
    'B18': (f31, 18, 2),
    'B21': (f32, 21, 2),
    'C21': (f33, 21, 3),
    'D21': (f34, 21, 4),
    'E21': (f35, 21, 5),
    'B22': (f36, 22, 2),
    'C22': (f37, 22, 3),
    'D22': (f38, 22, 4),
    'E22': (f38, 22, 5),
    'B23': (f39, 23, 2),
    'C23': (f40, 23, 3),
    'D23': (f41, 23, 4),
    'E23': (f42, 23, 5),
    'B24': (f43, 24, 2),
    'C24': (f44, 24, 3),
    'D24': (f45, 24, 4),
    'E24': (f46, 24, 5),
    'B25': (f47, 25, 2),
    'C25': (f48, 25, 3),
    'D25': (f49, 25, 4),
    'E25': (f50, 25, 5),
    'B26': (f51, 26, 2),
    'C26': (f52, 26, 3),
    'D26': (f53, 26, 4),
    'E26': (f54, 26, 5),
    'C27': (f55, 27, 3),
    'E27': (f56, 27, 5),
    'F27': (f57, 27, 6),
    'F28': (f58, 28, 6),
    'F29': (f59, 29, 6),
    'B31': (f60, 31, 2),
    'C32': (f61, 32, 3),
    'F32': (f62, 32, 6),
    'B35': (f63, 35, 2),
    'C35': (f64, 35, 3),
    'D35': (f65, 35, 4),
    'E35': (f66, 35, 5),
    'F35': (f67, 35, 6),
    'G35': (f68, 35, 7),
    'H35': (f69, 35, 8),
    'B36': (f70, 36, 2),
    'C36': (f71, 36, 3),
    'D36': (f72, 36, 4),
    'E36': (f73, 36, 5),
    'F36': (f74, 36, 6),
    'G36': (f75, 36, 7),
    'H36': (f76, 36, 8),
    'B37': (f77, 37, 2),
    'C37': (f78, 37, 3),
    'D37': (f79, 37, 4),
    'E37': (f80, 37, 5),
    'F37': (f81, 37, 6),
    'G37': (f82, 37, 7),
    'H37': (f83, 37, 8),
    'B38': (f84, 38, 2),
    'C38': (f85, 38, 3),
    'D38': (f86, 38, 4),
    'E38': (f87, 38, 5),
    'F38': (f88, 38, 6),
    'G38': (f89, 38, 7),
    'H38': (f90, 38, 8),
    'B39': (f91, 39, 2),
    'C39': (f92, 39, 3),
    'D39': (f93, 39, 4),
    'E39': (f94, 39, 5),
    'F39': (f95, 39, 6),
    'G39': (f96, 39, 7),
    'H39': (f97, 39, 8),
    'B40': (f98, 40, 2),
    'C40': (f99, 40, 3),
    'D40': (f100, 40, 4),
    'E40': (f101, 40, 5),
    'F40': (f102, 40, 6),
    'G40': (f103, 40, 7),
    'H40': (f104, 40, 8),
    'B41': (f105, 41, 2),
    'C41': (f106, 41, 3),
    'D41': (f107, 41, 4),
    'E41': (f108, 41, 5),
    'F41': (f109, 41, 6),
    'G41': (f110, 41, 7),
    'H41': (f111, 41, 8),
    'B42': (f112, 42, 2),
    'C42': (f113, 42, 3),
    'D42': (f114, 42, 4),
    'E42': (f115, 42, 5),
    'F42': (f116, 42, 6),
    'G42': (f117, 42, 7),
    'H42': (f118, 42, 8),
    'B46': (f119, 46, 2),
    'C48': (f120, 48, 3),
    'D48': (f121, 48, 4),
    'E48': (f122, 48, 5),
    'F48': (f123, 48, 6),
    'G48': (f124, 48, 7),
    'H48': (f125, 48, 8),
    'X48': (f126, 48, 24),
    'Y48': (f127, 48, 25),
    'C49': (f120, 49, 3),
    'D49': (f121, 49, 4),
    'E49': (f122, 49, 5),
    'F49': (f123, 49, 6),
    'G49': (f124, 49, 7),
    'H49': (f125, 49, 8),
    'X49': (f128, 49, 24),
    'Y49': (f127, 49, 25),
    'C50': (f120, 50, 3),
    'D50': (f121, 50, 4),
    'E50': (f122, 50, 5),
    'F50': (f129, 50, 6),
    'G50': (f130, 50, 7),
    'H50': (f125, 50, 8),
    'X50': (f131, 50, 24),
    'Y50': (f127, 50, 25),
    'C51': (f120, 51, 3),
    'D51': (f121, 51, 4),
    'E51': (f122, 51, 5),
    'F51': (f129, 51, 6),
    'G51': (f130, 51, 7),
    'H51': (f125, 51, 8),
    'X51': (f132, 51, 24),
    'Y51': (f127, 51, 25),
    'C52': (f120, 52, 3),
    'D52': (f121, 52, 4),
    'E52': (f122, 52, 5),
    'F52': (f123, 52, 6),
    'G52': (f124, 52, 7),
    'H52': (f125, 52, 8),
    'X52': (f133, 52, 24),
    'Y52': (f127, 52, 25),
    'C53': (f120, 53, 3),
    'D53': (f121, 53, 4),
    'E53': (f122, 53, 5),
    'F53': (f123, 53, 6),
    'G53': (f124, 53, 7),
    'H53': (f125, 53, 8),
    'X53': (f134, 53, 24),
    'Y53': (f127, 53, 25),
    'B68': (f135, 68, 2),
    'B70': (f136, 70, 2),
    'C70': (f137, 70, 3),
    'D70': (f65, 70, 4),
    'E70': (f138, 70, 5),
    'F70': (f139, 70, 6),
    'G70': (f140, 70, 7),
    'H70': (f141, 70, 8),
    'X70': (f142, 70, 24),
    'Y70': (f143, 70, 25),
    'B71': (f144, 71, 2),
    'C71': (f145, 71, 3),
    'D71': (f72, 71, 4),
    'E71': (f146, 71, 5),
    'F71': (f147, 71, 6),
    'G71': (f148, 71, 7),
    'H71': (f141, 71, 8),
    'X71': (f149, 71, 24),
    'Y71': (f150, 71, 25),
    'B72': (f151, 72, 2),
    'C72': (f152, 72, 3),
    'D72': (f79, 72, 4),
    'E72': (f153, 72, 5),
    'F72': (f154, 72, 6),
    'G72': (f155, 72, 7),
    'H72': (f141, 72, 8),
    'X72': (f156, 72, 24),
    'Y72': (f157, 72, 25),
    'B73': (f158, 73, 2),
    'C73': (f159, 73, 3),
    'D73': (f86, 73, 4),
    'E73': (f160, 73, 5),
    'F73': (f161, 73, 6),
    'G73': (f162, 73, 7),
    'H73': (f141, 73, 8),
    'X73': (f163, 73, 24),
    'Y73': (f164, 73, 25),
    'B74': (f165, 74, 2),
    'C74': (f166, 74, 3),
    'D74': (f93, 74, 4),
    'E74': (f167, 74, 5),
    'F74': (f168, 74, 6),
    'G74': (f169, 74, 7),
    'H74': (f141, 74, 8),
    'X74': (f170, 74, 24),
    'Y74': (f171, 74, 25),
    'B75': (f172, 75, 2),
    'C75': (f173, 75, 3),
    'D75': (f100, 75, 4),
    'E75': (f174, 75, 5),
    'F75': (f175, 75, 6),
    'G75': (f176, 75, 7),
    'H75': (f141, 75, 8),
    'X75': (f177, 75, 24),
    'Y75': (f178, 75, 25),
    'B76': (f179, 76, 2),
    'C76': (f180, 76, 3),
    'D76': (f107, 76, 4),
    'E76': (f181, 76, 5),
    'F76': (f182, 76, 6),
    'G76': (f183, 76, 7),
    'H76': (f141, 76, 8),
    'X76': (f184, 76, 24),
    'Y76': (f185, 76, 25),
    'B77': (f186, 77, 2),
    'C77': (f187, 77, 3),
    'D77': (f114, 77, 4),
    'E77': (f188, 77, 5),
    'F77': (f189, 77, 6),
    'G77': (f190, 77, 7),
    'H77': (f141, 77, 8),
    'X77': (f191, 77, 24),
    'Y77': (f192, 77, 25),
}
