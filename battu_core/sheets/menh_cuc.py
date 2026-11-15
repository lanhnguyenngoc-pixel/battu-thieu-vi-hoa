# -*- coding: utf-8 -*-
"""Module sheet MENH_CUC – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.6.xlsx (SHA-256 c1e0254833a8f522615199e1e0140460e912227d0b9ffc88b9164ebc26d3e0b3).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'MENH_CUC'
SI = 18

def f0(E, r, c):
    # Ô C6
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+0,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),1.0))

def f1(E, r, c):
    # Ô D6
    # = TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,SAO!$G$5=0),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,SAO!$G$6=0),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$
    #   7=1,SAO!$Y$7=0,SAO!$G$7=0),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,SAO!$G$8=0),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0
    #   ,SAO!$G$9=0),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,SAO!$G$10=0),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,SAO!$G$1
    #   1=0),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,SAO!$G$12=0),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,SAO!$G$13=
    #   0),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,SAO!$G$14=0),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,SAO!$G$15=0),SAO!$E
    #   $15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,SAO!$G$16=0),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,SAO!$G$17=0),SAO!$E$17&
    #   "(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,SAO!$G$18=0),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,SAO!$G$19=0),SAO!$E$19&"(Giờ tàng)
    #    ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,SAO!$G$20=0),SAO!$E$20&"(Giờ tàng) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.eq(E.v(13,5,7),0.0))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.eq(E.v(13,6,7),0.0))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.eq(E.v(13,7,7),0.0))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.eq(E.v(13,8,7),0.0))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.eq(E.v(13,9,7),0.0))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.eq(E.v(13,10,7),0.0))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.eq(E.v(13,11,7),0.0))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.eq(E.v(13,12,7),0.0))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.eq(E.v(13,13,7),0.0))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.eq(E.v(13,14,7),0.0))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.eq(E.v(13,15,7),0.0))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.eq(E.v(13,16,7),0.0))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.eq(E.v(13,17,7),0.0))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.eq(E.v(13,18,7),0.0))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.eq(E.v(13,19,7),0.0))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.eq(E.v(13,20,7),0.0))) else '')))

def f2(E, r, c):
    # Ô E6
    # = PP1_DINH_TINH!$F$96
    return E.v(9,96,6)

def f3(E, r, c):
    # Ô F6
    # = PP1_DINH_TINH!$D$96
    return E.v(9,96,4)

def f4(E, r, c):
    # Ô G6
    # = TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(TRUNG_GIAN!$E$7+0,5)+1),"0.0")&" độ ("&IF(SUM(DUNG_HY_KY!$D$24:$D$28)=0,"–",TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MO
    #   D(TRUNG_GIAN!$E$7+0,5)+1)/SUM(DUNG_HY_KY!$D$24:$D$28),"0%"))&")"
    return X.cat(X.cat(X.cat(X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),1.0)), '0.0'),' độ ('),('–' if X.truthy(X.eq(X.SUM(E.rg(8,24,4,28,4)),0.0)) else X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),1.0)),X.SUM(E.rg(8,24,4,28,4))), '0%'))),')')

def f5(E, r, c):
    # Ô H6
    # = IF(MOD(TRUNG_GIAN!$E$7+0,5)=DUNG_HY_KY!$C$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+0,5)=DUNG_HY_KY!$C$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+0,5)=DUNG_HY_KY!$C$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)=MOD(TRUNG_GIAN!$E$7+0,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),E.v(8,109,3))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),E.v(8,110,3))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),E.v(8,111,3))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)))) else '—'))))

def f6(E, r, c):
    # Ô I6
    # = IF(MOD(TRUNG_GIAN!$E$7+0,5)=DUNG_HY_KY!$D$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+0,5)=DUNG_HY_KY!$D$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+0,5)=DUNG_HY_KY!$D$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)=MOD(TRUNG_GIAN!$E$7+0,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),E.v(8,109,4))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),E.v(8,110,4))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),E.v(8,111,4))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)))) else '—'))))

def f7(E, r, c):
    # Ô C7
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))

def f8(E, r, c):
    # Ô D7
    # = TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,SAO!$G$5=1),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,SAO!$G$6=1),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$
    #   7=1,SAO!$Y$7=0,SAO!$G$7=1),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,SAO!$G$8=1),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0
    #   ,SAO!$G$9=1),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,SAO!$G$10=1),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,SAO!$G$1
    #   1=1),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,SAO!$G$12=1),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,SAO!$G$13=
    #   1),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,SAO!$G$14=1),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,SAO!$G$15=1),SAO!$E
    #   $15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,SAO!$G$16=1),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,SAO!$G$17=1),SAO!$E$17&
    #   "(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,SAO!$G$18=1),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,SAO!$G$19=1),SAO!$E$19&"(Giờ tàng)
    #    ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,SAO!$G$20=1),SAO!$E$20&"(Giờ tàng) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.eq(E.v(13,5,7),1.0))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.eq(E.v(13,6,7),1.0))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.eq(E.v(13,7,7),1.0))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.eq(E.v(13,8,7),1.0))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.eq(E.v(13,9,7),1.0))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.eq(E.v(13,10,7),1.0))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.eq(E.v(13,11,7),1.0))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.eq(E.v(13,12,7),1.0))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.eq(E.v(13,13,7),1.0))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.eq(E.v(13,14,7),1.0))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.eq(E.v(13,15,7),1.0))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.eq(E.v(13,16,7),1.0))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.eq(E.v(13,17,7),1.0))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.eq(E.v(13,18,7),1.0))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.eq(E.v(13,19,7),1.0))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.eq(E.v(13,20,7),1.0))) else '')))

def f9(E, r, c):
    # Ô E7
    # = PP1_DINH_TINH!$F$97
    return E.v(9,97,6)

def f10(E, r, c):
    # Ô F7
    # = PP1_DINH_TINH!$D$97
    return E.v(9,97,4)

def f11(E, r, c):
    # Ô G7
    # = TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(TRUNG_GIAN!$E$7+1,5)+1),"0.0")&" độ ("&IF(SUM(DUNG_HY_KY!$D$24:$D$28)=0,"–",TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MO
    #   D(TRUNG_GIAN!$E$7+1,5)+1)/SUM(DUNG_HY_KY!$D$24:$D$28),"0%"))&")"
    return X.cat(X.cat(X.cat(X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0)), '0.0'),' độ ('),('–' if X.truthy(X.eq(X.SUM(E.rg(8,24,4,28,4)),0.0)) else X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0)),X.SUM(E.rg(8,24,4,28,4))), '0%'))),')')

def f12(E, r, c):
    # Ô H7
    # = IF(MOD(TRUNG_GIAN!$E$7+1,5)=DUNG_HY_KY!$C$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+1,5)=DUNG_HY_KY!$C$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+1,5)=DUNG_HY_KY!$C$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)=MOD(TRUNG_GIAN!$E$7+1,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(8,109,3))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(8,110,3))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(8,111,3))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),X.MOD(X.add(E.v(23,7,5),1.0), 5.0)))) else '—'))))

def f13(E, r, c):
    # Ô I7
    # = IF(MOD(TRUNG_GIAN!$E$7+1,5)=DUNG_HY_KY!$D$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+1,5)=DUNG_HY_KY!$D$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+1,5)=DUNG_HY_KY!$D$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)=MOD(TRUNG_GIAN!$E$7+1,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(8,109,4))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(8,110,4))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(8,111,4))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),X.MOD(X.add(E.v(23,7,5),1.0), 5.0)))) else '—'))))

def f14(E, r, c):
    # Ô C8
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))

def f15(E, r, c):
    # Ô D8
    # = TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,SAO!$G$5=2),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,SAO!$G$6=2),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$
    #   7=1,SAO!$Y$7=0,SAO!$G$7=2),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,SAO!$G$8=2),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0
    #   ,SAO!$G$9=2),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,SAO!$G$10=2),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,SAO!$G$1
    #   1=2),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,SAO!$G$12=2),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,SAO!$G$13=
    #   2),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,SAO!$G$14=2),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,SAO!$G$15=2),SAO!$E
    #   $15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,SAO!$G$16=2),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,SAO!$G$17=2),SAO!$E$17&
    #   "(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,SAO!$G$18=2),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,SAO!$G$19=2),SAO!$E$19&"(Giờ tàng)
    #    ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,SAO!$G$20=2),SAO!$E$20&"(Giờ tàng) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.eq(E.v(13,5,7),2.0))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.eq(E.v(13,6,7),2.0))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.eq(E.v(13,7,7),2.0))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.eq(E.v(13,8,7),2.0))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.eq(E.v(13,9,7),2.0))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.eq(E.v(13,10,7),2.0))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.eq(E.v(13,11,7),2.0))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.eq(E.v(13,12,7),2.0))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.eq(E.v(13,13,7),2.0))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.eq(E.v(13,14,7),2.0))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.eq(E.v(13,15,7),2.0))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.eq(E.v(13,16,7),2.0))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.eq(E.v(13,17,7),2.0))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.eq(E.v(13,18,7),2.0))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.eq(E.v(13,19,7),2.0))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.eq(E.v(13,20,7),2.0))) else '')))

def f16(E, r, c):
    # Ô E8
    # = PP1_DINH_TINH!$F$98
    return E.v(9,98,6)

def f17(E, r, c):
    # Ô F8
    # = PP1_DINH_TINH!$D$98
    return E.v(9,98,4)

def f18(E, r, c):
    # Ô G8
    # = TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(TRUNG_GIAN!$E$7+2,5)+1),"0.0")&" độ ("&IF(SUM(DUNG_HY_KY!$D$24:$D$28)=0,"–",TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MO
    #   D(TRUNG_GIAN!$E$7+2,5)+1)/SUM(DUNG_HY_KY!$D$24:$D$28),"0%"))&")"
    return X.cat(X.cat(X.cat(X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0)), '0.0'),' độ ('),('–' if X.truthy(X.eq(X.SUM(E.rg(8,24,4,28,4)),0.0)) else X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0)),X.SUM(E.rg(8,24,4,28,4))), '0%'))),')')

def f19(E, r, c):
    # Ô H8
    # = IF(MOD(TRUNG_GIAN!$E$7+2,5)=DUNG_HY_KY!$C$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+2,5)=DUNG_HY_KY!$C$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+2,5)=DUNG_HY_KY!$C$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)=MOD(TRUNG_GIAN!$E$7+2,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(8,109,3))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(8,110,3))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(8,111,3))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)))) else '—'))))

def f20(E, r, c):
    # Ô I8
    # = IF(MOD(TRUNG_GIAN!$E$7+2,5)=DUNG_HY_KY!$D$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+2,5)=DUNG_HY_KY!$D$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+2,5)=DUNG_HY_KY!$D$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)=MOD(TRUNG_GIAN!$E$7+2,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(8,109,4))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(8,110,4))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(8,111,4))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)))) else '—'))))

def f21(E, r, c):
    # Ô C9
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+3,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0))

def f22(E, r, c):
    # Ô D9
    # = TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,SAO!$G$5=3),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,SAO!$G$6=3),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$
    #   7=1,SAO!$Y$7=0,SAO!$G$7=3),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,SAO!$G$8=3),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0
    #   ,SAO!$G$9=3),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,SAO!$G$10=3),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,SAO!$G$1
    #   1=3),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,SAO!$G$12=3),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,SAO!$G$13=
    #   3),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,SAO!$G$14=3),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,SAO!$G$15=3),SAO!$E
    #   $15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,SAO!$G$16=3),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,SAO!$G$17=3),SAO!$E$17&
    #   "(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,SAO!$G$18=3),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,SAO!$G$19=3),SAO!$E$19&"(Giờ tàng)
    #    ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,SAO!$G$20=3),SAO!$E$20&"(Giờ tàng) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.eq(E.v(13,5,7),3.0))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.eq(E.v(13,6,7),3.0))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.eq(E.v(13,7,7),3.0))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.eq(E.v(13,8,7),3.0))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.eq(E.v(13,9,7),3.0))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.eq(E.v(13,10,7),3.0))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.eq(E.v(13,11,7),3.0))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.eq(E.v(13,12,7),3.0))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.eq(E.v(13,13,7),3.0))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.eq(E.v(13,14,7),3.0))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.eq(E.v(13,15,7),3.0))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.eq(E.v(13,16,7),3.0))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.eq(E.v(13,17,7),3.0))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.eq(E.v(13,18,7),3.0))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.eq(E.v(13,19,7),3.0))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.eq(E.v(13,20,7),3.0))) else '')))

def f23(E, r, c):
    # Ô E9
    # = PP1_DINH_TINH!$F$99
    return E.v(9,99,6)

def f24(E, r, c):
    # Ô F9
    # = PP1_DINH_TINH!$D$99
    return E.v(9,99,4)

def f25(E, r, c):
    # Ô G9
    # = TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(TRUNG_GIAN!$E$7+3,5)+1),"0.0")&" độ ("&IF(SUM(DUNG_HY_KY!$D$24:$D$28)=0,"–",TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MO
    #   D(TRUNG_GIAN!$E$7+3,5)+1)/SUM(DUNG_HY_KY!$D$24:$D$28),"0%"))&")"
    return X.cat(X.cat(X.cat(X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0)), '0.0'),' độ ('),('–' if X.truthy(X.eq(X.SUM(E.rg(8,24,4,28,4)),0.0)) else X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0)),X.SUM(E.rg(8,24,4,28,4))), '0%'))),')')

def f26(E, r, c):
    # Ô H9
    # = IF(MOD(TRUNG_GIAN!$E$7+3,5)=DUNG_HY_KY!$C$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+3,5)=DUNG_HY_KY!$C$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+3,5)=DUNG_HY_KY!$C$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)=MOD(TRUNG_GIAN!$E$7+3,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),E.v(8,109,3))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),E.v(8,110,3))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),E.v(8,111,3))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—'))))

def f27(E, r, c):
    # Ô I9
    # = IF(MOD(TRUNG_GIAN!$E$7+3,5)=DUNG_HY_KY!$D$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+3,5)=DUNG_HY_KY!$D$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+3,5)=DUNG_HY_KY!$D$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)=MOD(TRUNG_GIAN!$E$7+3,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),E.v(8,109,4))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),E.v(8,110,4))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),E.v(8,111,4))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—'))))

def f28(E, r, c):
    # Ô C10
    # = INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))

def f29(E, r, c):
    # Ô D10
    # = TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,SAO!$G$5=4),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,SAO!$G$6=4),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$
    #   7=1,SAO!$Y$7=0,SAO!$G$7=4),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,SAO!$G$8=4),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0
    #   ,SAO!$G$9=4),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,SAO!$G$10=4),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,SAO!$G$1
    #   1=4),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,SAO!$G$12=4),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,SAO!$G$13=
    #   4),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,SAO!$G$14=4),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,SAO!$G$15=4),SAO!$E
    #   $15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,SAO!$G$16=4),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,SAO!$G$17=4),SAO!$E$17&
    #   "(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,SAO!$G$18=4),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,SAO!$G$19=4),SAO!$E$19&"(Giờ tàng)
    #    ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,SAO!$G$20=4),SAO!$E$20&"(Giờ tàng) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.eq(E.v(13,5,7),4.0))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.eq(E.v(13,6,7),4.0))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.eq(E.v(13,7,7),4.0))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.eq(E.v(13,8,7),4.0))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.eq(E.v(13,9,7),4.0))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.eq(E.v(13,10,7),4.0))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.eq(E.v(13,11,7),4.0))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.eq(E.v(13,12,7),4.0))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.eq(E.v(13,13,7),4.0))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.eq(E.v(13,14,7),4.0))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.eq(E.v(13,15,7),4.0))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.eq(E.v(13,16,7),4.0))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.eq(E.v(13,17,7),4.0))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.eq(E.v(13,18,7),4.0))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.eq(E.v(13,19,7),4.0))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.eq(E.v(13,20,7),4.0))) else '')))

def f30(E, r, c):
    # Ô E10
    # = PP1_DINH_TINH!$F$100
    return E.v(9,100,6)

def f31(E, r, c):
    # Ô F10
    # = PP1_DINH_TINH!$D$100
    return E.v(9,100,4)

def f32(E, r, c):
    # Ô G10
    # = TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(TRUNG_GIAN!$E$7+4,5)+1),"0.0")&" độ ("&IF(SUM(DUNG_HY_KY!$D$24:$D$28)=0,"–",TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,MO
    #   D(TRUNG_GIAN!$E$7+4,5)+1)/SUM(DUNG_HY_KY!$D$24:$D$28),"0%"))&")"
    return X.cat(X.cat(X.cat(X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0)), '0.0'),' độ ('),('–' if X.truthy(X.eq(X.SUM(E.rg(8,24,4,28,4)),0.0)) else X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0)),X.SUM(E.rg(8,24,4,28,4))), '0%'))),')')

def f33(E, r, c):
    # Ô H10
    # = IF(MOD(TRUNG_GIAN!$E$7+4,5)=DUNG_HY_KY!$C$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+4,5)=DUNG_HY_KY!$C$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+4,5)=DUNG_HY_KY!$C$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)=MOD(TRUNG_GIAN!$E$7+4,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),E.v(8,109,3))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),E.v(8,110,3))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),E.v(8,111,3))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else '—'))))

def f34(E, r, c):
    # Ô I10
    # = IF(MOD(TRUNG_GIAN!$E$7+4,5)=DUNG_HY_KY!$D$109,"Dụng",IF(MOD(TRUNG_GIAN!$E$7+4,5)=DUNG_HY_KY!$D$110,"Hỷ",IF(MOD(TRUNG_GIAN!$E$7+4,5)=DUNG_HY_KY!$D$111,
    #   "Kỵ",IF(AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)=MOD(TRUNG_GIAN!$E$7+4,5)),"Kỵ (nhóm nhiều)","—"))))
    return ('Dụng' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),E.v(8,109,4))) else ('Hỷ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),E.v(8,110,4))) else ('Kỵ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),E.v(8,111,4))) else ('Kỵ (nhóm nhiều)' if X.truthy(X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)))) else '—'))))

def f35(E, r, c):
    # Ô D12
    # = "PP1: "&DUNG_HY_KY!$C$6&" | PP2: "&DUNG_HY_KY!$D$6
    return X.cat(X.cat(X.cat('PP1: ',E.v(8,6,3)),' | PP2: '),E.v(8,6,4))

def f36(E, r, c):
    # Ô D13
    # = "PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1),"—")&" / "&
    #   IF(DUNG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$111+1),"—")&" | PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&" / "&IF(
    #   DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$110+1),"—")&" / "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' | PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),' / '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),' / '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—'))

def f37(E, r, c):
    # Ô D14
    # = "PP1 (nhật "&IF(DUNG_HY_KY!$C$7=1,"mạnh",IF(DUNG_HY_KY!$C$7=0,"nhược","trung hòa – bảng không áp dụng"))&"): Giàu nếu "&IF(DUNG_HY_KY!$C$7=1,"tài","tỉ
    #    kiếp")&" là dụng; Sang nếu "&IF(DUNG_HY_KY!$C$7=1,"quan sát","ấn")&" là dụng; Thọ nếu "&IF(DUNG_HY_KY!$C$7=1,"thực thương","ấn")&" là dụng; Nghèo nếu
    #    "&IF(DUNG_HY_KY!$C$7=1,"tỉ kiếp","tài")&" là kỵ; Hèn nếu "&IF(DUNG_HY_KY!$C$7=1,"ấn","quan sát")&" là kỵ; Yểu nếu "&IF(DUNG_HY_KY!$C$7=1,"ấn","thực t
    #   hương")&" là kỵ. | PP2 (nhật "&IF(DUNG_HY_KY!$D$7=1,"vượng","nhược")&")."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,3),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,3),0.0)) else 'trung hòa – bảng không áp dụng'))),'): Giàu nếu '),('tài' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'tỉ kiếp')),' là dụng; Sang nếu '),('quan sát' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'ấn')),' là dụng; Thọ nếu '),('thực thương' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'ấn')),' là dụng; Nghèo nếu '),('tỉ kiếp' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'tài')),' là kỵ; Hèn nếu '),('ấn' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'quan sát')),' là kỵ; Yểu nếu '),('ấn' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'thực thương')),' là kỵ. | PP2 (nhật '),('vượng' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'nhược')),').')

def f38(E, r, c):
    # Ô D18, D19
    # = IF(AND(O18=R18,O18="CÂN BẰNG"),"MỨC TRUNG BÌNH: hai bên cân nhau (cả hai phương pháp)",IF(O18=R18,"NGHIÊNG "&IF(AND(O18=R18,O18<>"CÂN BẰNG",O18<>"TRUN
    #   G HÒA",P18>=2*Q18,M18>=2*N18),"RÕ RỆT ","")&"VỀ "&O18&" (hai phương pháp thống nhất)",IF(O18="TRUNG HÒA","Nhật chủ trung hòa (PP1) – bảng P1 tr.173 kh
    #   ông áp dụng; theo PP2: nghiêng về "&R18,IF(O18="CÂN BẰNG","NGHIÊNG NHẸ VỀ "&R18&" (PP1 cân bằng, PP2 nghiêng về "&R18&")",IF(R18="CÂN BẰNG","NGHIÊNG N
    #   HẸ VỀ "&O18&" (PP2 cân bằng, PP1 nghiêng về "&O18&")",IF(DOI_CHIEU!$C$33=1,"NGHIÊNG VỀ "&O18&" (theo PP1 – phương pháp được khuyến nghị về vượng suy ở
    #    DOI_CHIEU; PP2 nghiêng về "&R18&")",IF(DOI_CHIEU!$C$33=2,"NGHIÊNG VỀ "&R18&" (theo PP2 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP1 n
    #   ghiêng về "&O18&")","MỨC TRUNG BÌNH: PP1 nghiêng về "&O18&", PP2 nghiêng về "&R18&" – không bên nào trội; ứng theo vận")))))))
    return ('MỨC TRUNG BÌNH: hai bên cân nhau (cả hai phương pháp)' if X.truthy(X.AND(X.eq(E.v(18,r,c+11),E.v(18,r,c+14)), X.eq(E.v(18,r,c+11),'CÂN BẰNG'))) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG ',('RÕ RỆT ' if X.truthy(X.AND(X.eq(E.v(18,r,c+11),E.v(18,r,c+14)), X.ne(E.v(18,r,c+11),'CÂN BẰNG'), X.ne(E.v(18,r,c+11),'TRUNG HÒA'), X.ge(E.v(18,r,c+12),X.mul(2.0,E.v(18,r,c+13))), X.ge(E.v(18,r,c+9),X.mul(2.0,E.v(18,r,c+10))))) else '')),'VỀ '),E.v(18,r,c+11)),' (hai phương pháp thống nhất)') if X.truthy(X.eq(E.v(18,r,c+11),E.v(18,r,c+14))) else (X.cat('Nhật chủ trung hòa (PP1) – bảng P1 tr.173 không áp dụng; theo PP2: nghiêng về ',E.v(18,r,c+14)) if X.truthy(X.eq(E.v(18,r,c+11),'TRUNG HÒA')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG NHẸ VỀ ',E.v(18,r,c+14)),' (PP1 cân bằng, PP2 nghiêng về '),E.v(18,r,c+14)),')') if X.truthy(X.eq(E.v(18,r,c+11),'CÂN BẰNG')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG NHẸ VỀ ',E.v(18,r,c+11)),' (PP2 cân bằng, PP1 nghiêng về '),E.v(18,r,c+11)),')') if X.truthy(X.eq(E.v(18,r,c+14),'CÂN BẰNG')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG VỀ ',E.v(18,r,c+11)),' (theo PP1 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP2 nghiêng về '),E.v(18,r,c+14)),')') if X.truthy(X.eq(E.v(7,33,3),1.0)) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG VỀ ',E.v(18,r,c+14)),' (theo PP2 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP1 nghiêng về '),E.v(18,r,c+11)),')') if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.cat(X.cat(X.cat(X.cat('MỨC TRUNG BÌNH: PP1 nghiêng về ',E.v(18,r,c+11)),', PP2 nghiêng về '),E.v(18,r,c+14)),' – không bên nào trội; ứng theo vận'))))))))

def f39(E, r, c):
    # Ô G18
    # = S18&CHAR(10)&T18&CHAR(10)&"Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng giàu, nhóm bất lợi càng vượng càng nghèo; lấy bên mạnh hơn làm kết
    #    luận; mạnh gấp đôi trở lên ghi ""rõ rệt"" (giả định)."
    return X.cat(X.cat(X.cat(X.cat(E.v(18,r,c+12),X.CHAR(10.0)),E.v(18,r,c+13)),X.CHAR(10.0)),'Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng giàu, nhóm bất lợi càng vượng càng nghèo; lấy bên mạnh hơn làm kết luận; mạnh gấp đôi trở lên ghi "rõ rệt" (giả định).')

def f40(E, r, c):
    # Ô K18
    # = IF(U18="GIÀU","Sách nêu – khớp kết luận: "&IF(V18="","(chỉ theo bảng tóm tắt)",V18)&IF(W18<>"",CHAR(10)&"Yếu tố ngược chiều (nhẹ hơn): "&W18,""),IF(U1
    #   8="NGHÈO","Sách nêu – khớp kết luận: "&IF(W18="","(chỉ theo bảng tóm tắt)",W18)&IF(V18<>"",CHAR(10)&"Điểm sáng (không đủ đảo ngược): "&V18,""),"Không 
    #   bên nào trội, các ý của sách ứng theo vận:"&CHAR(10)&"– Vận dụng/hỷ thần: "&IF(V18="","—",V18)&CHAR(10)&"– Vận kỵ thần: "&IF(W18="","—",W18)))
    return (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),(X.cat(X.cat(X.CHAR(10.0),'Yếu tố ngược chiều (nhẹ hơn): '),E.v(18,r,c+12)) if X.truthy(X.ne(E.v(18,r,c+12),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'GIÀU')) else (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12))),(X.cat(X.cat(X.CHAR(10.0),'Điểm sáng (không đủ đảo ngược): '),E.v(18,r,c+11)) if X.truthy(X.ne(E.v(18,r,c+11),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'NGHÈO')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không bên nào trội, các ý của sách ứng theo vận:',X.CHAR(10.0)),'– Vận dụng/hỷ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),X.CHAR(10.0)),'– Vận kỵ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12)))))

def f41(E, r, c):
    # Ô M18
    # = IF(DUNG_HY_KY!$C$7=1,INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+2,5)+1),INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+0,5)+1))
    return (X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),2.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,3),1.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),0.0), 5.0),1.0)))

def f42(E, r, c):
    # Ô N18
    # = IF(DUNG_HY_KY!$C$7=1,INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+0,5)+1),INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+2,5)+1))
    return (X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),0.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,3),1.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),2.0), 5.0),1.0)))

def f43(E, r, c):
    # Ô O18
    # = IF(AND("C"="C",DUNG_HY_KY!$C$7=2),"TRUNG HÒA",IF(M18>N18,"GIÀU",IF(M18<N18,"NGHÈO","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('C','C'), X.eq(E.v(8,7,3),2.0))) else ('GIÀU' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('NGHÈO' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f44(E, r, c):
    # Ô P18
    # = IF(DUNG_HY_KY!$D$7=1,INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+2,5)+1),INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+0,5)+1))
    return (X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),2.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,4),1.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),0.0), 5.0),1.0)))

def f45(E, r, c):
    # Ô Q18
    # = IF(DUNG_HY_KY!$D$7=1,INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+0,5)+1),INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+2,5)+1))
    return (X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),0.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,4),1.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),2.0), 5.0),1.0)))

def f46(E, r, c):
    # Ô R18
    # = IF(AND("D"="C",DUNG_HY_KY!$D$7=2),"TRUNG HÒA",IF(P18>Q18,"GIÀU",IF(P18<Q18,"NGHÈO","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('D','C'), X.eq(E.v(8,7,4),2.0))) else ('GIÀU' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('NGHÈO' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f47(E, r, c):
    # Ô S18
    # = "PP1 (nhật "&IF(DUNG_HY_KY!$C$7=1,"mạnh",IF(DUNG_HY_KY!$C$7=0,"nhược","trung hòa"))&"): "&IF(DUNG_HY_KY!$C$7=1,"Tài","Tỉ kiếp")&" "&TEXT(M18,"0")&" so
    #    với "&IF(DUNG_HY_KY!$C$7=1,"Tỉ kiếp","Tài")&" "&TEXT(N18,"0")&" chữ → "&O18
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,3),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,3),0.0)) else 'trung hòa'))),'): '),('Tài' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'Tỉ kiếp')),' '),X.TEXT(E.v(18,r,c-6), '0')),' so với '),('Tỉ kiếp' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'Tài')),' '),X.TEXT(E.v(18,r,c-5), '0')),' chữ → '),E.v(18,r,c-4))

def f48(E, r, c):
    # Ô T18
    # = "PP2 (nhật "&IF(DUNG_HY_KY!$D$7=1,"mạnh",IF(DUNG_HY_KY!$D$7=0,"nhược","trung hòa"))&"): "&IF(DUNG_HY_KY!$D$7=1,"Tài","Tỉ kiếp")&" "&TEXT(P18,"0.0")&" 
    #   so với "&IF(DUNG_HY_KY!$D$7=1,"Tỉ kiếp","Tài")&" "&TEXT(Q18,"0.0")&" độ → "&R18
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,4),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,4),0.0)) else 'trung hòa'))),'): '),('Tài' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'Tỉ kiếp')),' '),X.TEXT(E.v(18,r,c-4), '0.0')),' so với '),('Tỉ kiếp' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'Tài')),' '),X.TEXT(E.v(18,r,c-3), '0.0')),' độ → '),E.v(18,r,c-2))

def f49(E, r, c):
    # Ô U18, U19, U20, U21
    # = IF(O18=R18,O18,IF(O18="CÂN BẰNG",R18,IF(OR(R18="CÂN BẰNG",R18="TRUNG HÒA"),O18,IF(O18="TRUNG HÒA",R18,IF(DOI_CHIEU!$C$33=1,O18,IF(DOI_CHIEU!$C$33=2,R1
    #   8,"TB"))))))
    return (E.v(18,r,c-6) if X.truthy(X.eq(E.v(18,r,c-6),E.v(18,r,c-3))) else (E.v(18,r,c-3) if X.truthy(X.eq(E.v(18,r,c-6),'CÂN BẰNG')) else (E.v(18,r,c-6) if X.truthy(X.OR(X.eq(E.v(18,r,c-3),'CÂN BẰNG'), X.eq(E.v(18,r,c-3),'TRUNG HÒA'))) else (E.v(18,r,c-3) if X.truthy(X.eq(E.v(18,r,c-6),'TRUNG HÒA')) else (E.v(18,r,c-6) if X.truthy(X.eq(E.v(7,33,3),1.0)) else (E.v(18,r,c-3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'TB'))))))

def f50(E, r, c):
    # Ô V18
    # = TRIM(IF(MENH_CUC!$J$62<>"","• "&MENH_CUC!$G$62&" ["&MENH_CUC!$B$62&IF(MENH_CUC!$J$62="cả hai","","; chỉ "&MENH_CUC!$J$62)&"] ","")&IF(MENH_CUC!$J$68<>
    #   "","• "&MENH_CUC!$G$68&" ["&MENH_CUC!$B$68&IF(MENH_CUC!$J$68="cả hai","","; chỉ "&MENH_CUC!$J$68)&"] ","")&IF(MENH_CUC!$J$69<>"","• "&MENH_CUC!$G$69&"
    #    ["&MENH_CUC!$B$69&IF(MENH_CUC!$J$69="cả hai","","; chỉ "&MENH_CUC!$J$69)&"] ","")&IF(MENH_CUC!$J$70<>"","• "&MENH_CUC!$G$70&" ["&MENH_CUC!$B$70&IF(ME
    #   NH_CUC!$J$70="cả hai","","; chỉ "&MENH_CUC!$J$70)&"] ","")&IF(MENH_CUC!$J$71<>"","• "&MENH_CUC!$G$71&" ["&MENH_CUC!$B$71&IF(MENH_CUC!$J$71="cả hai",""
    #   ,"; chỉ "&MENH_CUC!$J$71)&"] ","")&IF(MENH_CUC!$J$72<>"","• "&MENH_CUC!$G$72&" ["&MENH_CUC!$B$72&IF(MENH_CUC!$J$72="cả hai","","; chỉ "&MENH_CUC!$J$72
    #   )&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,62,7)),' ['),E.v(18,62,2)),('' if X.truthy(X.eq(E.v(18,62,10),'cả hai')) else X.cat('; chỉ ',E.v(18,62,10)))),'] ') if X.truthy(X.ne(E.v(18,62,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,68,7)),' ['),E.v(18,68,2)),('' if X.truthy(X.eq(E.v(18,68,10),'cả hai')) else X.cat('; chỉ ',E.v(18,68,10)))),'] ') if X.truthy(X.ne(E.v(18,68,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,69,7)),' ['),E.v(18,69,2)),('' if X.truthy(X.eq(E.v(18,69,10),'cả hai')) else X.cat('; chỉ ',E.v(18,69,10)))),'] ') if X.truthy(X.ne(E.v(18,69,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,70,7)),' ['),E.v(18,70,2)),('' if X.truthy(X.eq(E.v(18,70,10),'cả hai')) else X.cat('; chỉ ',E.v(18,70,10)))),'] ') if X.truthy(X.ne(E.v(18,70,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,71,7)),' ['),E.v(18,71,2)),('' if X.truthy(X.eq(E.v(18,71,10),'cả hai')) else X.cat('; chỉ ',E.v(18,71,10)))),'] ') if X.truthy(X.ne(E.v(18,71,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,72,7)),' ['),E.v(18,72,2)),('' if X.truthy(X.eq(E.v(18,72,10),'cả hai')) else X.cat('; chỉ ',E.v(18,72,10)))),'] ') if X.truthy(X.ne(E.v(18,72,10),'')) else '')))

def f51(E, r, c):
    # Ô W18
    # = TRIM(IF(MENH_CUC!$J$65<>"","• "&MENH_CUC!$G$65&" ["&MENH_CUC!$B$65&IF(MENH_CUC!$J$65="cả hai","","; chỉ "&MENH_CUC!$J$65)&"] ","")&IF(MENH_CUC!$J$73<>
    #   "","• "&MENH_CUC!$G$73&" ["&MENH_CUC!$B$73&IF(MENH_CUC!$J$73="cả hai","","; chỉ "&MENH_CUC!$J$73)&"] ","")&IF(MENH_CUC!$J$74<>"","• "&MENH_CUC!$G$74&"
    #    ["&MENH_CUC!$B$74&IF(MENH_CUC!$J$74="cả hai","","; chỉ "&MENH_CUC!$J$74)&"] ","")&IF(MENH_CUC!$J$75<>"","• "&MENH_CUC!$G$75&" ["&MENH_CUC!$B$75&IF(ME
    #   NH_CUC!$J$75="cả hai","","; chỉ "&MENH_CUC!$J$75)&"] ","")&IF(MENH_CUC!$J$76<>"","• "&MENH_CUC!$G$76&" ["&MENH_CUC!$B$76&IF(MENH_CUC!$J$76="cả hai",""
    #   ,"; chỉ "&MENH_CUC!$J$76)&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,65,7)),' ['),E.v(18,65,2)),('' if X.truthy(X.eq(E.v(18,65,10),'cả hai')) else X.cat('; chỉ ',E.v(18,65,10)))),'] ') if X.truthy(X.ne(E.v(18,65,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,73,7)),' ['),E.v(18,73,2)),('' if X.truthy(X.eq(E.v(18,73,10),'cả hai')) else X.cat('; chỉ ',E.v(18,73,10)))),'] ') if X.truthy(X.ne(E.v(18,73,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,74,7)),' ['),E.v(18,74,2)),('' if X.truthy(X.eq(E.v(18,74,10),'cả hai')) else X.cat('; chỉ ',E.v(18,74,10)))),'] ') if X.truthy(X.ne(E.v(18,74,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,75,7)),' ['),E.v(18,75,2)),('' if X.truthy(X.eq(E.v(18,75,10),'cả hai')) else X.cat('; chỉ ',E.v(18,75,10)))),'] ') if X.truthy(X.ne(E.v(18,75,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,76,7)),' ['),E.v(18,76,2)),('' if X.truthy(X.eq(E.v(18,76,10),'cả hai')) else X.cat('; chỉ ',E.v(18,76,10)))),'] ') if X.truthy(X.ne(E.v(18,76,10),'')) else '')))

def f52(E, r, c):
    # Ô G19
    # = S19&CHAR(10)&T19&CHAR(10)&"Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng sang, nhóm bất lợi càng vượng càng hèn; lấy bên mạnh hơn làm kết l
    #   uận; mạnh gấp đôi trở lên ghi ""rõ rệt"" (giả định)."
    return X.cat(X.cat(X.cat(X.cat(E.v(18,r,c+12),X.CHAR(10.0)),E.v(18,r,c+13)),X.CHAR(10.0)),'Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng sang, nhóm bất lợi càng vượng càng hèn; lấy bên mạnh hơn làm kết luận; mạnh gấp đôi trở lên ghi "rõ rệt" (giả định).')

def f53(E, r, c):
    # Ô K19
    # = IF(U19="SANG","Sách nêu – khớp kết luận: "&IF(V19="","(chỉ theo bảng tóm tắt)",V19)&IF(W19<>"",CHAR(10)&"Yếu tố ngược chiều (nhẹ hơn): "&W19,""),IF(U1
    #   9="HÈN","Sách nêu – khớp kết luận: "&IF(W19="","(chỉ theo bảng tóm tắt)",W19)&IF(V19<>"",CHAR(10)&"Điểm sáng (không đủ đảo ngược): "&V19,""),"Không bê
    #   n nào trội, các ý của sách ứng theo vận:"&CHAR(10)&"– Vận dụng/hỷ thần: "&IF(V19="","—",V19)&CHAR(10)&"– Vận kỵ thần: "&IF(W19="","—",W19)))
    return (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),(X.cat(X.cat(X.CHAR(10.0),'Yếu tố ngược chiều (nhẹ hơn): '),E.v(18,r,c+12)) if X.truthy(X.ne(E.v(18,r,c+12),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'SANG')) else (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12))),(X.cat(X.cat(X.CHAR(10.0),'Điểm sáng (không đủ đảo ngược): '),E.v(18,r,c+11)) if X.truthy(X.ne(E.v(18,r,c+11),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'HÈN')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không bên nào trội, các ý của sách ứng theo vận:',X.CHAR(10.0)),'– Vận dụng/hỷ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),X.CHAR(10.0)),'– Vận kỵ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12)))))

def f54(E, r, c):
    # Ô M19
    # = IF(DUNG_HY_KY!$C$7=1,INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+3,5)+1),INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+4,5)+1))
    return (X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),3.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,3),1.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),4.0), 5.0),1.0)))

def f55(E, r, c):
    # Ô N19
    # = IF(DUNG_HY_KY!$C$7=1,INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+4,5)+1),INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+3,5)+1))
    return (X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),4.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,3),1.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),3.0), 5.0),1.0)))

def f56(E, r, c):
    # Ô O19
    # = IF(AND("C"="C",DUNG_HY_KY!$C$7=2),"TRUNG HÒA",IF(M19>N19,"SANG",IF(M19<N19,"HÈN","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('C','C'), X.eq(E.v(8,7,3),2.0))) else ('SANG' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('HÈN' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f57(E, r, c):
    # Ô P19
    # = IF(DUNG_HY_KY!$D$7=1,INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+3,5)+1),INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+4,5)+1))
    return (X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),3.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,4),1.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),4.0), 5.0),1.0)))

def f58(E, r, c):
    # Ô Q19
    # = IF(DUNG_HY_KY!$D$7=1,INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+4,5)+1),INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+3,5)+1))
    return (X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),4.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,4),1.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),3.0), 5.0),1.0)))

def f59(E, r, c):
    # Ô R19
    # = IF(AND("D"="C",DUNG_HY_KY!$D$7=2),"TRUNG HÒA",IF(P19>Q19,"SANG",IF(P19<Q19,"HÈN","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('D','C'), X.eq(E.v(8,7,4),2.0))) else ('SANG' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('HÈN' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f60(E, r, c):
    # Ô S19
    # = "PP1 (nhật "&IF(DUNG_HY_KY!$C$7=1,"mạnh",IF(DUNG_HY_KY!$C$7=0,"nhược","trung hòa"))&"): "&IF(DUNG_HY_KY!$C$7=1,"Quan sát","Ấn")&" "&TEXT(M19,"0")&" so
    #    với "&IF(DUNG_HY_KY!$C$7=1,"Ấn","Quan sát")&" "&TEXT(N19,"0")&" chữ → "&O19
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,3),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,3),0.0)) else 'trung hòa'))),'): '),('Quan sát' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'Ấn')),' '),X.TEXT(E.v(18,r,c-6), '0')),' so với '),('Ấn' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'Quan sát')),' '),X.TEXT(E.v(18,r,c-5), '0')),' chữ → '),E.v(18,r,c-4))

def f61(E, r, c):
    # Ô T19
    # = "PP2 (nhật "&IF(DUNG_HY_KY!$D$7=1,"mạnh",IF(DUNG_HY_KY!$D$7=0,"nhược","trung hòa"))&"): "&IF(DUNG_HY_KY!$D$7=1,"Quan sát","Ấn")&" "&TEXT(P19,"0.0")&" 
    #   so với "&IF(DUNG_HY_KY!$D$7=1,"Ấn","Quan sát")&" "&TEXT(Q19,"0.0")&" độ → "&R19
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,4),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,4),0.0)) else 'trung hòa'))),'): '),('Quan sát' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'Ấn')),' '),X.TEXT(E.v(18,r,c-4), '0.0')),' so với '),('Ấn' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'Quan sát')),' '),X.TEXT(E.v(18,r,c-3), '0.0')),' độ → '),E.v(18,r,c-2))

def f62(E, r, c):
    # Ô V19
    # = TRIM(IF(MENH_CUC!$J$63<>"","• "&MENH_CUC!$G$63&" ["&MENH_CUC!$B$63&IF(MENH_CUC!$J$63="cả hai","","; chỉ "&MENH_CUC!$J$63)&"] ","")&IF(MENH_CUC!$J$77<>
    #   "","• "&MENH_CUC!$G$77&" ["&MENH_CUC!$B$77&IF(MENH_CUC!$J$77="cả hai","","; chỉ "&MENH_CUC!$J$77)&"] ","")&IF(MENH_CUC!$J$78<>"","• "&MENH_CUC!$G$78&"
    #    ["&MENH_CUC!$B$78&IF(MENH_CUC!$J$78="cả hai","","; chỉ "&MENH_CUC!$J$78)&"] ","")&IF(MENH_CUC!$J$79<>"","• "&MENH_CUC!$G$79&" ["&MENH_CUC!$B$79&IF(ME
    #   NH_CUC!$J$79="cả hai","","; chỉ "&MENH_CUC!$J$79)&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,63,7)),' ['),E.v(18,63,2)),('' if X.truthy(X.eq(E.v(18,63,10),'cả hai')) else X.cat('; chỉ ',E.v(18,63,10)))),'] ') if X.truthy(X.ne(E.v(18,63,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,77,7)),' ['),E.v(18,77,2)),('' if X.truthy(X.eq(E.v(18,77,10),'cả hai')) else X.cat('; chỉ ',E.v(18,77,10)))),'] ') if X.truthy(X.ne(E.v(18,77,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,78,7)),' ['),E.v(18,78,2)),('' if X.truthy(X.eq(E.v(18,78,10),'cả hai')) else X.cat('; chỉ ',E.v(18,78,10)))),'] ') if X.truthy(X.ne(E.v(18,78,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,79,7)),' ['),E.v(18,79,2)),('' if X.truthy(X.eq(E.v(18,79,10),'cả hai')) else X.cat('; chỉ ',E.v(18,79,10)))),'] ') if X.truthy(X.ne(E.v(18,79,10),'')) else '')))

def f63(E, r, c):
    # Ô W19
    # = TRIM(IF(MENH_CUC!$J$66<>"","• "&MENH_CUC!$G$66&" ["&MENH_CUC!$B$66&IF(MENH_CUC!$J$66="cả hai","","; chỉ "&MENH_CUC!$J$66)&"] ","")&IF(MENH_CUC!$J$80<>
    #   "","• "&MENH_CUC!$G$80&" ["&MENH_CUC!$B$80&IF(MENH_CUC!$J$80="cả hai","","; chỉ "&MENH_CUC!$J$80)&"] ","")&IF(MENH_CUC!$J$81<>"","• "&MENH_CUC!$G$81&"
    #    ["&MENH_CUC!$B$81&IF(MENH_CUC!$J$81="cả hai","","; chỉ "&MENH_CUC!$J$81)&"] ","")&IF(MENH_CUC!$J$82<>"","• "&MENH_CUC!$G$82&" ["&MENH_CUC!$B$82&IF(ME
    #   NH_CUC!$J$82="cả hai","","; chỉ "&MENH_CUC!$J$82)&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,66,7)),' ['),E.v(18,66,2)),('' if X.truthy(X.eq(E.v(18,66,10),'cả hai')) else X.cat('; chỉ ',E.v(18,66,10)))),'] ') if X.truthy(X.ne(E.v(18,66,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,80,7)),' ['),E.v(18,80,2)),('' if X.truthy(X.eq(E.v(18,80,10),'cả hai')) else X.cat('; chỉ ',E.v(18,80,10)))),'] ') if X.truthy(X.ne(E.v(18,80,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,81,7)),' ['),E.v(18,81,2)),('' if X.truthy(X.eq(E.v(18,81,10),'cả hai')) else X.cat('; chỉ ',E.v(18,81,10)))),'] ') if X.truthy(X.ne(E.v(18,81,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,82,7)),' ['),E.v(18,82,2)),('' if X.truthy(X.eq(E.v(18,82,10),'cả hai')) else X.cat('; chỉ ',E.v(18,82,10)))),'] ') if X.truthy(X.ne(E.v(18,82,10),'')) else '')))

def f64(E, r, c):
    # Ô D20
    # = IF(AND(O20=R20,O20="CÂN BẰNG"),"MỨC TRUNG BÌNH: hai bên cân nhau (cả hai phương pháp)",IF(O20=R20,"NGHIÊNG "&IF(AND(O20=R20,O20<>"CÂN BẰNG",O20<>"TRUN
    #   G HÒA",P20>=2*Q20,M20>=2*N20),"RÕ RỆT ","")&"VỀ "&O20&" (hai phương pháp thống nhất)",IF(O20="TRUNG HÒA","Nhật chủ trung hòa (PP1) – bảng P1 tr.173 kh
    #   ông áp dụng; theo PP2: nghiêng về "&R20,IF(O20="CÂN BẰNG","NGHIÊNG NHẸ VỀ "&R20&" (PP1 cân bằng, PP2 nghiêng về "&R20&")",IF(R20="CÂN BẰNG","NGHIÊNG N
    #   HẸ VỀ "&O20&" (PP2 cân bằng, PP1 nghiêng về "&O20&")",IF(DOI_CHIEU!$C$33=1,"NGHIÊNG VỀ "&O20&" (theo PP1 – phương pháp được khuyến nghị về vượng suy ở
    #    DOI_CHIEU; PP2 nghiêng về "&R20&")",IF(DOI_CHIEU!$C$33=2,"NGHIÊNG VỀ "&R20&" (theo PP2 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP1 n
    #   ghiêng về "&O20&")","MỨC TRUNG BÌNH: PP1 nghiêng về "&O20&", PP2 nghiêng về "&R20&" – không bên nào trội; ứng theo vận")))))))&IF(DUNG_HY_KY!$C$114<>"
    #   Không",". Dụng thần PP1 KHUYẾT trong tứ trụ → ""không có dụng thần … thì hung"" (P1 tr.171)","")
    return X.cat(('MỨC TRUNG BÌNH: hai bên cân nhau (cả hai phương pháp)' if X.truthy(X.AND(X.eq(E.v(18,r,c+11),E.v(18,r,c+14)), X.eq(E.v(18,r,c+11),'CÂN BẰNG'))) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG ',('RÕ RỆT ' if X.truthy(X.AND(X.eq(E.v(18,r,c+11),E.v(18,r,c+14)), X.ne(E.v(18,r,c+11),'CÂN BẰNG'), X.ne(E.v(18,r,c+11),'TRUNG HÒA'), X.ge(E.v(18,r,c+12),X.mul(2.0,E.v(18,r,c+13))), X.ge(E.v(18,r,c+9),X.mul(2.0,E.v(18,r,c+10))))) else '')),'VỀ '),E.v(18,r,c+11)),' (hai phương pháp thống nhất)') if X.truthy(X.eq(E.v(18,r,c+11),E.v(18,r,c+14))) else (X.cat('Nhật chủ trung hòa (PP1) – bảng P1 tr.173 không áp dụng; theo PP2: nghiêng về ',E.v(18,r,c+14)) if X.truthy(X.eq(E.v(18,r,c+11),'TRUNG HÒA')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG NHẸ VỀ ',E.v(18,r,c+14)),' (PP1 cân bằng, PP2 nghiêng về '),E.v(18,r,c+14)),')') if X.truthy(X.eq(E.v(18,r,c+11),'CÂN BẰNG')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG NHẸ VỀ ',E.v(18,r,c+11)),' (PP2 cân bằng, PP1 nghiêng về '),E.v(18,r,c+11)),')') if X.truthy(X.eq(E.v(18,r,c+14),'CÂN BẰNG')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG VỀ ',E.v(18,r,c+11)),' (theo PP1 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP2 nghiêng về '),E.v(18,r,c+14)),')') if X.truthy(X.eq(E.v(7,33,3),1.0)) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG VỀ ',E.v(18,r,c+14)),' (theo PP2 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP1 nghiêng về '),E.v(18,r,c+11)),')') if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.cat(X.cat(X.cat(X.cat('MỨC TRUNG BÌNH: PP1 nghiêng về ',E.v(18,r,c+11)),', PP2 nghiêng về '),E.v(18,r,c+14)),' – không bên nào trội; ứng theo vận')))))))),('. Dụng thần PP1 KHUYẾT trong tứ trụ → "không có dụng thần … thì hung" (P1 tr.171)' if X.truthy(X.ne(E.v(8,114,3),'Không')) else ''))

def f65(E, r, c):
    # Ô G20
    # = S20&CHAR(10)&T20&CHAR(10)&"Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng cát, nhóm bất lợi càng vượng càng hung; lấy bên mạnh hơn làm kết l
    #   uận; mạnh gấp đôi trở lên ghi ""rõ rệt"" (giả định)."
    return X.cat(X.cat(X.cat(X.cat(E.v(18,r,c+12),X.CHAR(10.0)),E.v(18,r,c+13)),X.CHAR(10.0)),'Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng cát, nhóm bất lợi càng vượng càng hung; lấy bên mạnh hơn làm kết luận; mạnh gấp đôi trở lên ghi "rõ rệt" (giả định).')

def f66(E, r, c):
    # Ô K20
    # = IF(U20="CÁT","Sách nêu – khớp kết luận: "&IF(V20="","(chỉ theo bảng tóm tắt)",V20)&IF(W20<>"",CHAR(10)&"Yếu tố ngược chiều (nhẹ hơn): "&W20,""),IF(U20
    #   ="HUNG","Sách nêu – khớp kết luận: "&IF(W20="","(chỉ theo bảng tóm tắt)",W20)&IF(V20<>"",CHAR(10)&"Điểm sáng (không đủ đảo ngược): "&V20,""),"Không bê
    #   n nào trội, các ý của sách ứng theo vận:"&CHAR(10)&"– Vận dụng/hỷ thần: "&IF(V20="","—",V20)&CHAR(10)&"– Vận kỵ thần: "&IF(W20="","—",W20)))
    return (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),(X.cat(X.cat(X.CHAR(10.0),'Yếu tố ngược chiều (nhẹ hơn): '),E.v(18,r,c+12)) if X.truthy(X.ne(E.v(18,r,c+12),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'CÁT')) else (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12))),(X.cat(X.cat(X.CHAR(10.0),'Điểm sáng (không đủ đảo ngược): '),E.v(18,r,c+11)) if X.truthy(X.ne(E.v(18,r,c+11),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'HUNG')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không bên nào trội, các ý của sách ứng theo vận:',X.CHAR(10.0)),'– Vận dụng/hỷ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),X.CHAR(10.0)),'– Vận kỵ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12)))))

def f67(E, r, c):
    # Ô M20
    # = (IF(DUNG_HY_KY!$C$109<0,0,INDEX(DUNG_HY_KY!$C$24:$C$28,DUNG_HY_KY!$C$109+1))+IF(DUNG_HY_KY!$C$110<0,0,INDEX(DUNG_HY_KY!$C$24:$C$28,DUNG_HY_KY!$C$110+1
    #   )))
    return X.add((0.0 if X.truthy(X.lt(E.v(8,109,3),0.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,109,3),1.0))),(0.0 if X.truthy(X.lt(E.v(8,110,3),0.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,110,3),1.0))))

def f68(E, r, c):
    # Ô N20
    # = IF(AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)<>DUNG_HY_KY!$C$111),IF(DUNG_HY_KY!$C$111<0,0,INDEX(DUNG_HY_KY!$
    #   C$24:$C$28,DUNG_HY_KY!$C$111+1))+INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92,5)+1),IF(DUNG_HY_KY!$C$111<0,0,INDEX(DUNG_HY_KY!$C$
    #   24:$C$28,DUNG_HY_KY!$C$111+1)))
    return (X.add((0.0 if X.truthy(X.lt(E.v(8,111,3),0.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,111,3),1.0))),X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),1.0))) if X.truthy(X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.ne(X.MOD(X.add(E.v(8,5,3),E.v(8,92,3)), 5.0),E.v(8,111,3)))) else (0.0 if X.truthy(X.lt(E.v(8,111,3),0.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,111,3),1.0))))

def f69(E, r, c):
    # Ô O20
    # = IF(AND("C"="C",DUNG_HY_KY!$C$7=2),"TRUNG HÒA",IF(M20>N20,"CÁT",IF(M20<N20,"HUNG","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('C','C'), X.eq(E.v(8,7,3),2.0))) else ('CÁT' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('HUNG' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f70(E, r, c):
    # Ô P20
    # = (IF(DUNG_HY_KY!$D$109<0,0,INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$109+1))+IF(DUNG_HY_KY!$D$110<0,0,INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$110+1
    #   )))
    return X.add((0.0 if X.truthy(X.lt(E.v(8,109,4),0.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,109,4),1.0))),(0.0 if X.truthy(X.lt(E.v(8,110,4),0.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,110,4),1.0))))

def f71(E, r, c):
    # Ô Q20
    # = IF(AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)<>DUNG_HY_KY!$D$111),IF(DUNG_HY_KY!$D$111<0,0,INDEX(DUNG_HY_KY!$
    #   D$24:$D$28,DUNG_HY_KY!$D$111+1))+INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92,5)+1),IF(DUNG_HY_KY!$D$111<0,0,INDEX(DUNG_HY_KY!$D$
    #   24:$D$28,DUNG_HY_KY!$D$111+1)))
    return (X.add((0.0 if X.truthy(X.lt(E.v(8,111,4),0.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,111,4),1.0))),X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),1.0))) if X.truthy(X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.ne(X.MOD(X.add(E.v(8,5,4),E.v(8,92,4)), 5.0),E.v(8,111,4)))) else (0.0 if X.truthy(X.lt(E.v(8,111,4),0.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,111,4),1.0))))

def f72(E, r, c):
    # Ô R20
    # = IF(AND("D"="C",DUNG_HY_KY!$D$7=2),"TRUNG HÒA",IF(P20>Q20,"CÁT",IF(P20<Q20,"HUNG","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('D','C'), X.eq(E.v(8,7,4),2.0))) else ('CÁT' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('HUNG' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f73(E, r, c):
    # Ô S20
    # = "PP1 (nhật "&IF(DUNG_HY_KY!$C$7=1,"mạnh",IF(DUNG_HY_KY!$C$7=0,"nhược","trung hòa"))&"): "&"dụng + hỷ"&" "&TEXT(M20,"0")&" so với "&"kỵ (+ nhóm nhiều)"
    #   &" "&TEXT(N20,"0")&" chữ → "&O20
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,3),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,3),0.0)) else 'trung hòa'))),'): '),'dụng + hỷ'),' '),X.TEXT(E.v(18,r,c-6), '0')),' so với '),'kỵ (+ nhóm nhiều)'),' '),X.TEXT(E.v(18,r,c-5), '0')),' chữ → '),E.v(18,r,c-4))

def f74(E, r, c):
    # Ô T20
    # = "PP2 (nhật "&IF(DUNG_HY_KY!$D$7=1,"mạnh",IF(DUNG_HY_KY!$D$7=0,"nhược","trung hòa"))&"): "&"dụng + hỷ"&" "&TEXT(P20,"0.0")&" so với "&"kỵ (+ nhóm nhiều
    #   )"&" "&TEXT(Q20,"0.0")&" độ → "&R20
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,4),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,4),0.0)) else 'trung hòa'))),'): '),'dụng + hỷ'),' '),X.TEXT(E.v(18,r,c-4), '0.0')),' so với '),'kỵ (+ nhóm nhiều)'),' '),X.TEXT(E.v(18,r,c-3), '0.0')),' độ → '),E.v(18,r,c-2))

def f75(E, r, c):
    # Ô V20
    # = TRIM(IF(MENH_CUC!$J$83<>"","• "&MENH_CUC!$G$83&" ["&MENH_CUC!$B$83&IF(MENH_CUC!$J$83="cả hai","","; chỉ "&MENH_CUC!$J$83)&"] ",""))
    return X.TRIM((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,83,7)),' ['),E.v(18,83,2)),('' if X.truthy(X.eq(E.v(18,83,10),'cả hai')) else X.cat('; chỉ ',E.v(18,83,10)))),'] ') if X.truthy(X.ne(E.v(18,83,10),'')) else ''))

def f76(E, r, c):
    # Ô W20
    # = TRIM(IF(MENH_CUC!$J$84<>"","• "&MENH_CUC!$G$84&" ["&MENH_CUC!$B$84&IF(MENH_CUC!$J$84="cả hai","","; chỉ "&MENH_CUC!$J$84)&"] ",""))
    return X.TRIM((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,84,7)),' ['),E.v(18,84,2)),('' if X.truthy(X.eq(E.v(18,84,10),'cả hai')) else X.cat('; chỉ ',E.v(18,84,10)))),'] ') if X.truthy(X.ne(E.v(18,84,10),'')) else ''))

def f77(E, r, c):
    # Ô D21
    # = IF(AND(O21=R21,O21="CÂN BẰNG"),"MỨC TRUNG BÌNH: hai bên cân nhau (cả hai phương pháp)",IF(O21=R21,"NGHIÊNG "&IF(AND(O21=R21,O21<>"CÂN BẰNG",O21<>"TRUN
    #   G HÒA",P21>=2*Q21,M21>=2*N21),"RÕ RỆT ","")&"VỀ "&O21&" (hai phương pháp thống nhất)",IF(O21="TRUNG HÒA","Nhật chủ trung hòa (PP1) – bảng P1 tr.173 kh
    #   ông áp dụng; theo PP2: nghiêng về "&R21,IF(O21="CÂN BẰNG","NGHIÊNG NHẸ VỀ "&R21&" (PP1 cân bằng, PP2 nghiêng về "&R21&")",IF(R21="CÂN BẰNG","NGHIÊNG N
    #   HẸ VỀ "&O21&" (PP2 cân bằng, PP1 nghiêng về "&O21&")",IF(DOI_CHIEU!$C$33=1,"NGHIÊNG VỀ "&O21&" (theo PP1 – phương pháp được khuyến nghị về vượng suy ở
    #    DOI_CHIEU; PP2 nghiêng về "&R21&")",IF(DOI_CHIEU!$C$33=2,"NGHIÊNG VỀ "&R21&" (theo PP2 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP1 n
    #   ghiêng về "&O21&")","MỨC TRUNG BÌNH: PP1 nghiêng về "&O21&", PP2 nghiêng về "&R21&" – không bên nào trội; ứng theo vận")))))))&IF(OR(DUNG_HY_KY!$C$7=2
    #   ,COUNTIF(PP1_DINH_TINH!$F$96:$F$100,0)=0),". Thêm: ngũ hành đủ hoặc trung hòa → ""trung hòa, thuần túy càng thọ"" (P1 tr.172)","")&". Thời điểm cần đề
    #    phòng: xem KET_LUAN dòng 144 (mục H)."
    return X.cat(X.cat(('MỨC TRUNG BÌNH: hai bên cân nhau (cả hai phương pháp)' if X.truthy(X.AND(X.eq(E.v(18,r,c+11),E.v(18,r,c+14)), X.eq(E.v(18,r,c+11),'CÂN BẰNG'))) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG ',('RÕ RỆT ' if X.truthy(X.AND(X.eq(E.v(18,r,c+11),E.v(18,r,c+14)), X.ne(E.v(18,r,c+11),'CÂN BẰNG'), X.ne(E.v(18,r,c+11),'TRUNG HÒA'), X.ge(E.v(18,r,c+12),X.mul(2.0,E.v(18,r,c+13))), X.ge(E.v(18,r,c+9),X.mul(2.0,E.v(18,r,c+10))))) else '')),'VỀ '),E.v(18,r,c+11)),' (hai phương pháp thống nhất)') if X.truthy(X.eq(E.v(18,r,c+11),E.v(18,r,c+14))) else (X.cat('Nhật chủ trung hòa (PP1) – bảng P1 tr.173 không áp dụng; theo PP2: nghiêng về ',E.v(18,r,c+14)) if X.truthy(X.eq(E.v(18,r,c+11),'TRUNG HÒA')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG NHẸ VỀ ',E.v(18,r,c+14)),' (PP1 cân bằng, PP2 nghiêng về '),E.v(18,r,c+14)),')') if X.truthy(X.eq(E.v(18,r,c+11),'CÂN BẰNG')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG NHẸ VỀ ',E.v(18,r,c+11)),' (PP2 cân bằng, PP1 nghiêng về '),E.v(18,r,c+11)),')') if X.truthy(X.eq(E.v(18,r,c+14),'CÂN BẰNG')) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG VỀ ',E.v(18,r,c+11)),' (theo PP1 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP2 nghiêng về '),E.v(18,r,c+14)),')') if X.truthy(X.eq(E.v(7,33,3),1.0)) else (X.cat(X.cat(X.cat(X.cat('NGHIÊNG VỀ ',E.v(18,r,c+14)),' (theo PP2 – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU; PP1 nghiêng về '),E.v(18,r,c+11)),')') if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.cat(X.cat(X.cat(X.cat('MỨC TRUNG BÌNH: PP1 nghiêng về ',E.v(18,r,c+11)),', PP2 nghiêng về '),E.v(18,r,c+14)),' – không bên nào trội; ứng theo vận')))))))),('. Thêm: ngũ hành đủ hoặc trung hòa → "trung hòa, thuần túy càng thọ" (P1 tr.172)' if X.truthy(X.OR(X.eq(E.v(8,7,3),2.0), X.eq(X.COUNTIF(E.rg(9,96,6,100,6), 0.0),0.0))) else '')),'. Thời điểm cần đề phòng: xem KET_LUAN dòng 144 (mục H).')

def f78(E, r, c):
    # Ô G21
    # = S21&CHAR(10)&T21&CHAR(10)&"Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng thọ, nhóm bất lợi càng vượng càng yểu; lấy bên mạnh hơn làm kết lu
    #   ận; mạnh gấp đôi trở lên ghi ""rõ rệt"" (giả định)."
    return X.cat(X.cat(X.cat(X.cat(E.v(18,r,c+12),X.CHAR(10.0)),E.v(18,r,c+13)),X.CHAR(10.0)),'Quy tắc: nhóm có lợi theo bảng P1 tr.173 càng vượng càng thọ, nhóm bất lợi càng vượng càng yểu; lấy bên mạnh hơn làm kết luận; mạnh gấp đôi trở lên ghi "rõ rệt" (giả định).')

def f79(E, r, c):
    # Ô K21
    # = IF(U21="THỌ","Sách nêu – khớp kết luận: "&IF(V21="","(chỉ theo bảng tóm tắt)",V21)&IF(W21<>"",CHAR(10)&"Yếu tố ngược chiều (nhẹ hơn): "&W21,""),IF(U21
    #   ="YỂU","Sách nêu – khớp kết luận: "&IF(W21="","(chỉ theo bảng tóm tắt)",W21)&IF(V21<>"",CHAR(10)&"Điểm sáng (không đủ đảo ngược): "&V21,""),"Không bên
    #    nào trội, các ý của sách ứng theo vận:"&CHAR(10)&"– Vận dụng/hỷ thần: "&IF(V21="","—",V21)&CHAR(10)&"– Vận kỵ thần: "&IF(W21="","—",W21)))
    return (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),(X.cat(X.cat(X.CHAR(10.0),'Yếu tố ngược chiều (nhẹ hơn): '),E.v(18,r,c+12)) if X.truthy(X.ne(E.v(18,r,c+12),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'THỌ')) else (X.cat(X.cat('Sách nêu – khớp kết luận: ',('(chỉ theo bảng tóm tắt)' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12))),(X.cat(X.cat(X.CHAR(10.0),'Điểm sáng (không đủ đảo ngược): '),E.v(18,r,c+11)) if X.truthy(X.ne(E.v(18,r,c+11),'')) else '')) if X.truthy(X.eq(E.v(18,r,c+10),'YỂU')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không bên nào trội, các ý của sách ứng theo vận:',X.CHAR(10.0)),'– Vận dụng/hỷ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+11),'')) else E.v(18,r,c+11))),X.CHAR(10.0)),'– Vận kỵ thần: '),('—' if X.truthy(X.eq(E.v(18,r,c+12),'')) else E.v(18,r,c+12)))))

def f80(E, r, c):
    # Ô M21
    # = IF(DUNG_HY_KY!$C$7=1,INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+1,5)+1),INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+4,5)+1))
    return (X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),1.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,3),1.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),4.0), 5.0),1.0)))

def f81(E, r, c):
    # Ô N21
    # = IF(DUNG_HY_KY!$C$7=1,INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+4,5)+1),INDEX(DUNG_HY_KY!$C$24:$C$28,MOD(DUNG_HY_KY!$C$5+1,5)+1))
    return (X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),4.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,3),1.0)) else X.INDEX(E.rg(8,24,3,28,3), X.add(X.MOD(X.add(E.v(8,5,3),1.0), 5.0),1.0)))

def f82(E, r, c):
    # Ô O21
    # = IF(AND("C"="C",DUNG_HY_KY!$C$7=2),"TRUNG HÒA",IF(M21>N21,"THỌ",IF(M21<N21,"YỂU","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('C','C'), X.eq(E.v(8,7,3),2.0))) else ('THỌ' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('YỂU' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f83(E, r, c):
    # Ô P21
    # = IF(DUNG_HY_KY!$D$7=1,INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+1,5)+1),INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+4,5)+1))
    return (X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),1.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,4),1.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),4.0), 5.0),1.0)))

def f84(E, r, c):
    # Ô Q21
    # = IF(DUNG_HY_KY!$D$7=1,INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+4,5)+1),INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+1,5)+1))
    return (X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),4.0), 5.0),1.0)) if X.truthy(X.eq(E.v(8,7,4),1.0)) else X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),1.0), 5.0),1.0)))

def f85(E, r, c):
    # Ô R21
    # = IF(AND("D"="C",DUNG_HY_KY!$D$7=2),"TRUNG HÒA",IF(P21>Q21,"THỌ",IF(P21<Q21,"YỂU","CÂN BẰNG")))
    return ('TRUNG HÒA' if X.truthy(X.AND(X.eq('D','C'), X.eq(E.v(8,7,4),2.0))) else ('THỌ' if X.truthy(X.gt(E.v(18,r,c-2),E.v(18,r,c-1))) else ('YỂU' if X.truthy(X.lt(E.v(18,r,c-2),E.v(18,r,c-1))) else 'CÂN BẰNG')))

def f86(E, r, c):
    # Ô S21
    # = "PP1 (nhật "&IF(DUNG_HY_KY!$C$7=1,"mạnh",IF(DUNG_HY_KY!$C$7=0,"nhược","trung hòa"))&"): "&IF(DUNG_HY_KY!$C$7=1,"Thực thương","Ấn")&" "&TEXT(M21,"0")&"
    #    so với "&IF(DUNG_HY_KY!$C$7=1,"Ấn","Thực thương")&" "&TEXT(N21,"0")&" chữ → "&O21
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,3),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,3),0.0)) else 'trung hòa'))),'): '),('Thực thương' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'Ấn')),' '),X.TEXT(E.v(18,r,c-6), '0')),' so với '),('Ấn' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'Thực thương')),' '),X.TEXT(E.v(18,r,c-5), '0')),' chữ → '),E.v(18,r,c-4))

def f87(E, r, c):
    # Ô T21
    # = "PP2 (nhật "&IF(DUNG_HY_KY!$D$7=1,"mạnh",IF(DUNG_HY_KY!$D$7=0,"nhược","trung hòa"))&"): "&IF(DUNG_HY_KY!$D$7=1,"Thực thương","Ấn")&" "&TEXT(P21,"0.0")
    #   &" so với "&IF(DUNG_HY_KY!$D$7=1,"Ấn","Thực thương")&" "&TEXT(Q21,"0.0")&" độ → "&R21
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP2 (nhật ',('mạnh' if X.truthy(X.eq(E.v(8,7,4),1.0)) else ('nhược' if X.truthy(X.eq(E.v(8,7,4),0.0)) else 'trung hòa'))),'): '),('Thực thương' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'Ấn')),' '),X.TEXT(E.v(18,r,c-4), '0.0')),' so với '),('Ấn' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'Thực thương')),' '),X.TEXT(E.v(18,r,c-3), '0.0')),' độ → '),E.v(18,r,c-2))

def f88(E, r, c):
    # Ô V21
    # = TRIM(IF(MENH_CUC!$J$64<>"","• "&MENH_CUC!$G$64&" ["&MENH_CUC!$B$64&IF(MENH_CUC!$J$64="cả hai","","; chỉ "&MENH_CUC!$J$64)&"] ","")&IF(MENH_CUC!$J$85<>
    #   "","• "&MENH_CUC!$G$85&" ["&MENH_CUC!$B$85&IF(MENH_CUC!$J$85="cả hai","","; chỉ "&MENH_CUC!$J$85)&"] ","")&IF(MENH_CUC!$J$86<>"","• "&MENH_CUC!$G$86&"
    #    ["&MENH_CUC!$B$86&IF(MENH_CUC!$J$86="cả hai","","; chỉ "&MENH_CUC!$J$86)&"] ","")&IF(MENH_CUC!$J$87<>"","• "&MENH_CUC!$G$87&" ["&MENH_CUC!$B$87&IF(ME
    #   NH_CUC!$J$87="cả hai","","; chỉ "&MENH_CUC!$J$87)&"] ","")&IF(MENH_CUC!$J$88<>"","• "&MENH_CUC!$G$88&" ["&MENH_CUC!$B$88&IF(MENH_CUC!$J$88="cả hai",""
    #   ,"; chỉ "&MENH_CUC!$J$88)&"] ","")&IF(MENH_CUC!$J$89<>"","• "&MENH_CUC!$G$89&" ["&MENH_CUC!$B$89&IF(MENH_CUC!$J$89="cả hai","","; chỉ "&MENH_CUC!$J$89
    #   )&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,64,7)),' ['),E.v(18,64,2)),('' if X.truthy(X.eq(E.v(18,64,10),'cả hai')) else X.cat('; chỉ ',E.v(18,64,10)))),'] ') if X.truthy(X.ne(E.v(18,64,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,85,7)),' ['),E.v(18,85,2)),('' if X.truthy(X.eq(E.v(18,85,10),'cả hai')) else X.cat('; chỉ ',E.v(18,85,10)))),'] ') if X.truthy(X.ne(E.v(18,85,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,86,7)),' ['),E.v(18,86,2)),('' if X.truthy(X.eq(E.v(18,86,10),'cả hai')) else X.cat('; chỉ ',E.v(18,86,10)))),'] ') if X.truthy(X.ne(E.v(18,86,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,87,7)),' ['),E.v(18,87,2)),('' if X.truthy(X.eq(E.v(18,87,10),'cả hai')) else X.cat('; chỉ ',E.v(18,87,10)))),'] ') if X.truthy(X.ne(E.v(18,87,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,88,7)),' ['),E.v(18,88,2)),('' if X.truthy(X.eq(E.v(18,88,10),'cả hai')) else X.cat('; chỉ ',E.v(18,88,10)))),'] ') if X.truthy(X.ne(E.v(18,88,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,89,7)),' ['),E.v(18,89,2)),('' if X.truthy(X.eq(E.v(18,89,10),'cả hai')) else X.cat('; chỉ ',E.v(18,89,10)))),'] ') if X.truthy(X.ne(E.v(18,89,10),'')) else '')))

def f89(E, r, c):
    # Ô W21
    # = TRIM(IF(MENH_CUC!$J$67<>"","• "&MENH_CUC!$G$67&" ["&MENH_CUC!$B$67&IF(MENH_CUC!$J$67="cả hai","","; chỉ "&MENH_CUC!$J$67)&"] ","")&IF(MENH_CUC!$J$90<>
    #   "","• "&MENH_CUC!$G$90&" ["&MENH_CUC!$B$90&IF(MENH_CUC!$J$90="cả hai","","; chỉ "&MENH_CUC!$J$90)&"] ","")&IF(MENH_CUC!$J$91<>"","• "&MENH_CUC!$G$91&"
    #    ["&MENH_CUC!$B$91&IF(MENH_CUC!$J$91="cả hai","","; chỉ "&MENH_CUC!$J$91)&"] ","")&IF(MENH_CUC!$J$92<>"","• "&MENH_CUC!$G$92&" ["&MENH_CUC!$B$92&IF(ME
    #   NH_CUC!$J$92="cả hai","","; chỉ "&MENH_CUC!$J$92)&"] ","")&IF(MENH_CUC!$J$93<>"","• "&MENH_CUC!$G$93&" ["&MENH_CUC!$B$93&IF(MENH_CUC!$J$93="cả hai",""
    #   ,"; chỉ "&MENH_CUC!$J$93)&"] ","")&IF(MENH_CUC!$J$94<>"","• "&MENH_CUC!$G$94&" ["&MENH_CUC!$B$94&IF(MENH_CUC!$J$94="cả hai","","; chỉ "&MENH_CUC!$J$94
    #   )&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,67,7)),' ['),E.v(18,67,2)),('' if X.truthy(X.eq(E.v(18,67,10),'cả hai')) else X.cat('; chỉ ',E.v(18,67,10)))),'] ') if X.truthy(X.ne(E.v(18,67,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,90,7)),' ['),E.v(18,90,2)),('' if X.truthy(X.eq(E.v(18,90,10),'cả hai')) else X.cat('; chỉ ',E.v(18,90,10)))),'] ') if X.truthy(X.ne(E.v(18,90,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,91,7)),' ['),E.v(18,91,2)),('' if X.truthy(X.eq(E.v(18,91,10),'cả hai')) else X.cat('; chỉ ',E.v(18,91,10)))),'] ') if X.truthy(X.ne(E.v(18,91,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,92,7)),' ['),E.v(18,92,2)),('' if X.truthy(X.eq(E.v(18,92,10),'cả hai')) else X.cat('; chỉ ',E.v(18,92,10)))),'] ') if X.truthy(X.ne(E.v(18,92,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,93,7)),' ['),E.v(18,93,2)),('' if X.truthy(X.eq(E.v(18,93,10),'cả hai')) else X.cat('; chỉ ',E.v(18,93,10)))),'] ') if X.truthy(X.ne(E.v(18,93,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(18,94,7)),' ['),E.v(18,94,2)),('' if X.truthy(X.eq(E.v(18,94,10),'cả hai')) else X.cat('; chỉ ',E.v(18,94,10)))),'] ') if X.truthy(X.ne(E.v(18,94,10),'')) else '')))

def f90(E, r, c):
    # Ô D26
    # = TRIM(IF(DAI_VAN!C78=1,"Vận "&DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi); ","")&IF(DAI_VAN!D78=1,"Vận "&DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&
    #   (DAI_VAN!D36+9)&" tuổi); ","")&IF(DAI_VAN!E78=1,"Vận "&DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi); ","")&IF(DAI_VAN!F78=1,"Vận "&DAI_VAN
    #   !F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi); ","")&IF(DAI_VAN!G78=1,"Vận "&DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi); ","")&IF(DA
    #   I_VAN!H78=1,"Vận "&DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi); ","")&IF(DAI_VAN!I78=1,"Vận "&DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I
    #   36+9)&" tuổi); ","")&IF(DAI_VAN!J78=1,"Vận "&DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi); ",""))&IF(COUNTIF(DAI_VAN!$C$78:$J$78,1)=0,"Khô
    #   ng có trong 8 vận","")
    return X.cat(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c-1)),' ('),E.v(14,r+10,c-1)),'–'),X.add(E.v(14,r+10,c-1),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c-1),1.0)) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c)),' ('),E.v(14,r+10,c)),'–'),X.add(E.v(14,r+10,c),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c+1)),' ('),E.v(14,r+10,c+1)),'–'),X.add(E.v(14,r+10,c+1),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c+1),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c+2)),' ('),E.v(14,r+10,c+2)),'–'),X.add(E.v(14,r+10,c+2),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c+2),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c+3)),' ('),E.v(14,r+10,c+3)),'–'),X.add(E.v(14,r+10,c+3),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c+3),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c+4)),' ('),E.v(14,r+10,c+4)),'–'),X.add(E.v(14,r+10,c+4),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c+4),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c+5)),' ('),E.v(14,r+10,c+5)),'–'),X.add(E.v(14,r+10,c+5),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c+5),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r+9,c+6)),' ('),E.v(14,r+10,c+6)),'–'),X.add(E.v(14,r+10,c+6),9.0)),' tuổi); ') if X.truthy(X.eq(E.v(14,r+52,c+6),1.0)) else ''))),('Không có trong 8 vận' if X.truthy(X.eq(X.COUNTIF(E.rg(14,78,3,78,10), 1.0),0.0)) else ''))

def f91(E, r, c):
    # Ô D27
    # = IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,TRUNG_GIAN!$D$6)>1,"Có – ""nếu chi tháng giống chi khác trong tứ trụ thì sinh mệnh thường không nguy"" (P1 tr.202)","K
    #   hông – khi gặp xung khắc đề cương cần lưu ý hơn (P1 tr.202)")
    return ('Có – "nếu chi tháng giống chi khác trong tứ trụ thì sinh mệnh thường không nguy" (P1 tr.202)' if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(23,6,4)),1.0)) else 'Không – khi gặp xung khắc đề cương cần lưu ý hơn (P1 tr.202)')

def f92(E, r, c):
    # Ô D28
    # = "PP1: "&IF(TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$109,"có – bị thiên khắc địa xung thì càng nguy (P1 tr.202)","không")&" | PP2: "&IF(TRUNG_GIAN!$D$9=DUNG_HY_KY
    #   !$D$109,"có","không")
    return X.cat(X.cat(X.cat('PP1: ',('có – bị thiên khắc địa xung thì càng nguy (P1 tr.202)' if X.truthy(X.eq(E.v(23,9,4),E.v(8,109,3))) else 'không')),' | PP2: '),('có' if X.truthy(X.eq(E.v(23,9,4),E.v(8,109,4))) else 'không'))

def f93(E, r, c):
    # Ô C33
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f94(E, r, c):
    # Ô D33
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(1,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f95(E, r, c):
    # Ô E33
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f96(E, r, c):
    # Ô F33
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f97(E, r, c):
    # Ô G33
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f98(E, r, c):
    # Ô K33
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(1,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(1.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f99(E, r, c):
    # Ô C34
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f100(E, r, c):
    # Ô D34
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(2,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f101(E, r, c):
    # Ô E34
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f102(E, r, c):
    # Ô F34
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f103(E, r, c):
    # Ô G34
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f104(E, r, c):
    # Ô K34
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(2,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(2.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f105(E, r, c):
    # Ô C35
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f106(E, r, c):
    # Ô D35
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(3,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f107(E, r, c):
    # Ô E35
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f108(E, r, c):
    # Ô F35
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f109(E, r, c):
    # Ô G35
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f110(E, r, c):
    # Ô K35
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(3,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(3.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f111(E, r, c):
    # Ô C36
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f112(E, r, c):
    # Ô D36
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(4,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f113(E, r, c):
    # Ô E36
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f114(E, r, c):
    # Ô F36
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f115(E, r, c):
    # Ô G36
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f116(E, r, c):
    # Ô K36
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(4,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(4.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f117(E, r, c):
    # Ô C37
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f118(E, r, c):
    # Ô D37
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(5,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f119(E, r, c):
    # Ô E37
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f120(E, r, c):
    # Ô F37
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f121(E, r, c):
    # Ô G37
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f122(E, r, c):
    # Ô K37
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(5,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(5.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f123(E, r, c):
    # Ô C38
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f124(E, r, c):
    # Ô D38
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(6,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f125(E, r, c):
    # Ô E38
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f126(E, r, c):
    # Ô F38
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f127(E, r, c):
    # Ô G38
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f128(E, r, c):
    # Ô K38
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(6,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(6.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f129(E, r, c):
    # Ô C39
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f130(E, r, c):
    # Ô D39
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(7,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f131(E, r, c):
    # Ô E39
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f132(E, r, c):
    # Ô F39
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f133(E, r, c):
    # Ô G39
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f134(E, r, c):
    # Ô K39
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(7,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(7.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f135(E, r, c):
    # Ô C40
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f136(E, r, c):
    # Ô D40
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(8,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f137(E, r, c):
    # Ô E40
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f138(E, r, c):
    # Ô F40
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f139(E, r, c):
    # Ô G40
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f140(E, r, c):
    # Ô K40
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(8,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(8.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f141(E, r, c):
    # Ô C41
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f142(E, r, c):
    # Ô D41
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(9,LUU_NIEN!$CA$30:$CA$1
    #   29,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f143(E, r, c):
    # Ô E41
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f144(E, r, c):
    # Ô F41
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f145(E, r, c):
    # Ô G41
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f146(E, r, c):
    # Ô K41
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(9,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(9.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f147(E, r, c):
    # Ô C42
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f148(E, r, c):
    # Ô D42
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(10,LUU_NIEN!$CA$30:$CA
    #   $129,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f149(E, r, c):
    # Ô E42
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f150(E, r, c):
    # Ô F42
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f151(E, r, c):
    # Ô G42
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f152(E, r, c):
    # Ô K42
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(10,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(10.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f153(E, r, c):
    # Ô C43
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f154(E, r, c):
    # Ô D43
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(11,LUU_NIEN!$CA$30:$CA
    #   $129,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f155(E, r, c):
    # Ô E43
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f156(E, r, c):
    # Ô F43
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f157(E, r, c):
    # Ô G43
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f158(E, r, c):
    # Ô K43
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(11,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(11.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f159(E, r, c):
    # Ô C44
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f160(E, r, c):
    # Ô D44
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(12,LUU_NIEN!$CA$30:$CA
    #   $129,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f161(E, r, c):
    # Ô E44
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f162(E, r, c):
    # Ô F44
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f163(E, r, c):
    # Ô G44
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f164(E, r, c):
    # Ô K44
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(12,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(12.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f165(E, r, c):
    # Ô C45
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f166(E, r, c):
    # Ô D45
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(13,LUU_NIEN!$CA$30:$CA
    #   $129,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f167(E, r, c):
    # Ô E45
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f168(E, r, c):
    # Ô F45
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f169(E, r, c):
    # Ô G45
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f170(E, r, c):
    # Ô K45
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(13,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(13.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f171(E, r, c):
    # Ô C46
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f172(E, r, c):
    # Ô D46
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(14,LUU_NIEN!$CA$30:$CA
    #   $129,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f173(E, r, c):
    # Ô E46
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f174(E, r, c):
    # Ô F46
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f175(E, r, c):
    # Ô G46
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f176(E, r, c):
    # Ô K46
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(14,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(14.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f177(E, r, c):
    # Ô C47
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f178(E, r, c):
    # Ô D47
    # = IFERROR("Lưu niên "&INDEX(LUU_NIEN!$L$30:$L$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(15,LUU_NIEN!$CA$30:$CA
    #   $129,0))&" / tiểu vận "&INDEX(LUU_NIEN!$N$30:$N$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat('Lưu niên ',X.INDEX(E.rg(15,30,12,129,12), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0))),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0))),' / tiểu vận '),X.INDEX(E.rg(15,30,14,129,14), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0))), lambda: '')

def f179(E, r, c):
    # Ô E47
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f180(E, r, c):
    # Ô F47
    # = IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f181(E, r, c):
    # Ô G47
    # = IFERROR(INDEX(LUU_NIEN!$BN$30:$BN$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,66,129,66), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f182(E, r, c):
    # Ô K47
    # = IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(15,LUU_NIEN!$CA$30:$CA$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(15.0, E.rg(15,30,79,129,79), 0.0)), lambda: '')

def f183(E, r, c):
    # Ô E62
    # = IF(OR(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,M
    #   OD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))))) else '—')

def f184(E, r, c):
    # Ô F62
    # = IF(OR(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,M
    #   OD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))))) else '—')

def f185(E, r, c):
    # Ô J62, J63, J64, J65, J66, J67, J68, J69, J70, J71, J72, J73 … (33 ô)
    # = IF(AND(E62="ĐỦ",F62="ĐỦ"),"cả hai",IF(E62="ĐỦ","PP1",IF(F62="ĐỦ","PP2","")))
    return ('cả hai' if X.truthy(X.AND(X.eq(E.v(18,r,c-5),'ĐỦ'), X.eq(E.v(18,r,c-4),'ĐỦ'))) else ('PP1' if X.truthy(X.eq(E.v(18,r,c-5),'ĐỦ')) else ('PP2' if X.truthy(X.eq(E.v(18,r,c-4),'ĐỦ')) else '')))

def f186(E, r, c):
    # Ô K62
    # = "Dữ kiện: "&"Tài: "&MENH_CUC!$E$8&" chữ, "&MENH_CUC!$G$8&IF(MENH_CUC!$D$8="",""," – "&MENH_CUC!$D$8)&"; "&"Tỉ kiếp: "&MENH_CUC!$E$6&" chữ, "&MENH_CUC!
    #   $G$6&IF(MENH_CUC!$D$6="",""," – "&MENH_CUC!$D$6)&"; "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG
    #   _HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ ngày: 
    #   "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.173): Giàu
    #   : nhật mạnh & tài là dụng; hoặc nhật nhược & tỉ kiếp là dụng → "&IF(J62="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J62&") ⇒ "
    #   &G62&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Tài: '),E.v(18,8,5)),' chữ, '),E.v(18,8,7)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(' – ',E.v(18,8,4)))),'; '),'Tỉ kiếp: '),E.v(18,6,5)),' chữ, '),E.v(18,6,7)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(' – ',E.v(18,6,4)))),'; '),'Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.173): Giàu: nhật mạnh & tài là dụng; hoặc nhật nhược & tỉ kiếp là dụng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f187(E, r, c):
    # Ô E63
    # = IF(OR(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,M
    #   OD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))))) else '—')

def f188(E, r, c):
    # Ô F63
    # = IF(OR(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,M
    #   OD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))))) else '—')

def f189(E, r, c):
    # Ô K63
    # = "Dữ kiện: "&"Quan sát: "&MENH_CUC!$E$9&" chữ, "&MENH_CUC!$G$9&IF(MENH_CUC!$D$9="",""," – "&MENH_CUC!$D$9)&"; "&"Ấn: "&MENH_CUC!$E$10&" chữ, "&MENH_CUC
    #   !$G$10&IF(MENH_CUC!$D$10="",""," – "&MENH_CUC!$D$10)&"; "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(
    #   DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ ng
    #   ày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.173 (t
    #   r.167)): Sang: nhật mạnh & quan sát là dụng (xem mâu thuẫn); hoặc nhật nhược & ấn là dụng → "&IF(J63="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J63&") ⇒ "&G63&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Quan sát: '),E.v(18,9,5)),' chữ, '),E.v(18,9,7)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(' – ',E.v(18,9,4)))),'; '),'Ấn: '),E.v(18,10,5)),' chữ, '),E.v(18,10,7)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(' – ',E.v(18,10,4)))),'; '),'Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.173 (tr.167)): Sang: nhật mạnh & quan sát là dụng (xem mâu thuẫn); hoặc nhật nhược & ấn là dụng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f190(E, r, c):
    # Ô E64
    # = IF(OR(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=1)),AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,M
    #   OD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),1.0))), X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))))) else '—')

def f191(E, r, c):
    # Ô F64
    # = IF(OR(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=1)),AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,M
    #   OD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),1.0))), X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))))) else '—')

def f192(E, r, c):
    # Ô K64
    # = "Dữ kiện: "&"Thực thương: "&MENH_CUC!$E$7&" chữ, "&MENH_CUC!$G$7&IF(MENH_CUC!$D$7="",""," – "&MENH_CUC!$D$7)&"; "&"Ấn: "&MENH_CUC!$E$10&" chữ, "&MENH_
    #   CUC!$G$10&IF(MENH_CUC!$D$10="",""," – "&MENH_CUC!$D$10)&"; "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&
    #   IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ
    #    ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.173
    #   ): Thọ: nhật mạnh & thực thương là dụng; hoặc nhật nhược & ấn là dụng → "&IF(J64="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J
    #   64&") ⇒ "&G64&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Thực thương: '),E.v(18,7,5)),' chữ, '),E.v(18,7,7)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(' – ',E.v(18,7,4)))),'; '),'Ấn: '),E.v(18,10,5)),' chữ, '),E.v(18,10,7)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(' – ',E.v(18,10,4)))),'; '),'Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.173): Thọ: nhật mạnh & thực thương là dụng; hoặc nhật nhược & ấn là dụng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f193(E, r, c):
    # Ô E65
    # = IF(OR(AND((DUNG_HY_KY!$C$7=1),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=0),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(
    #   DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=0))),AND((DUNG_HY_KY!$C$7=0),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=
    #   2),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=2)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),0.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),0.0)))), X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),2.0)))))) else '—')

def f194(E, r, c):
    # Ô F65
    # = IF(OR(AND((DUNG_HY_KY!$D$7=1),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=0),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(
    #   DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=0))),AND((DUNG_HY_KY!$D$7=0),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=
    #   2),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=2)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),0.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),0.0)))), X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),2.0)))))) else '—')

def f195(E, r, c):
    # Ô K65
    # = "Dữ kiện: "&"Tỉ kiếp: "&MENH_CUC!$E$6&" chữ, "&MENH_CUC!$G$6&IF(MENH_CUC!$D$6="",""," – "&MENH_CUC!$D$6)&"; "&"Tài: "&MENH_CUC!$E$8&" chữ, "&MENH_CUC!
    #   $G$8&IF(MENH_CUC!$D$8="",""," – "&MENH_CUC!$D$8)&"; "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Kỵ PP1: "&IF(DUNG_H
    #   Y_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Nhóm nhiều (
    #   kỵ theo cân bằng) PP1: "&DUNG_HY_KY!$C$112&" / PP2: "&DUNG_HY_KY!$D$112&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TR
    #   UNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.173): Nghèo: nhật mạnh & tỉ kiếp là kỵ; hoặc nhật nhược & tài là kỵ → "&
    #   IF(J65="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J65&") ⇒ "&G65&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Tỉ kiếp: '),E.v(18,6,5)),' chữ, '),E.v(18,6,7)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(' – ',E.v(18,6,4)))),'; '),'Tài: '),E.v(18,8,5)),' chữ, '),E.v(18,8,7)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(' – ',E.v(18,8,4)))),'; '),'Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Nhóm nhiều (kỵ theo cân bằng) PP1: '),E.v(8,112,3)),' / PP2: '),E.v(8,112,4)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.173): Nghèo: nhật mạnh & tỉ kiếp là kỵ; hoặc nhật nhược & tài là kỵ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f196(E, r, c):
    # Ô E66
    # = IF(OR(AND((DUNG_HY_KY!$C$7=1),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=4),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(
    #   DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=4))),AND((DUNG_HY_KY!$C$7=0),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=
    #   3),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),4.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),4.0)))), X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),3.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),3.0)))))) else '—')

def f197(E, r, c):
    # Ô F66
    # = IF(OR(AND((DUNG_HY_KY!$D$7=1),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=4),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(
    #   DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=4))),AND((DUNG_HY_KY!$D$7=0),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=
    #   3),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),4.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),4.0)))), X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),3.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),3.0)))))) else '—')

def f198(E, r, c):
    # Ô K66
    # = "Dữ kiện: "&"Ấn: "&MENH_CUC!$E$10&" chữ, "&MENH_CUC!$G$10&IF(MENH_CUC!$D$10="",""," – "&MENH_CUC!$D$10)&"; "&"Quan sát: "&MENH_CUC!$E$9&" chữ, "&MENH_
    #   CUC!$G$9&IF(MENH_CUC!$D$9="",""," – "&MENH_CUC!$D$9)&"; "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Kỵ PP1: "&IF(DU
    #   NG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Nhóm nhi
    #   ều (kỵ theo cân bằng) PP1: "&DUNG_HY_KY!$C$112&" / PP2: "&DUNG_HY_KY!$D$112&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TE
    #   N,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.173): Hèn: nhật mạnh & ấn là kỵ; hoặc nhật nhược & quan sát là kỵ → 
    #   "&IF(J66="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J66&") ⇒ "&G66&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Ấn: '),E.v(18,10,5)),' chữ, '),E.v(18,10,7)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(' – ',E.v(18,10,4)))),'; '),'Quan sát: '),E.v(18,9,5)),' chữ, '),E.v(18,9,7)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(' – ',E.v(18,9,4)))),'; '),'Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Nhóm nhiều (kỵ theo cân bằng) PP1: '),E.v(8,112,3)),' / PP2: '),E.v(8,112,4)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.173): Hèn: nhật mạnh & ấn là kỵ; hoặc nhật nhược & quan sát là kỵ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f199(E, r, c):
    # Ô E67
    # = IF(OR(AND((DUNG_HY_KY!$C$7=1),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=4),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(
    #   DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=4))),AND((DUNG_HY_KY!$C$7=0),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=
    #   1),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD(DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=1)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),4.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),4.0)))), X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),1.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),1.0)))))) else '—')

def f200(E, r, c):
    # Ô F67
    # = IF(OR(AND((DUNG_HY_KY!$D$7=1),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=4),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(
    #   DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=4))),AND((DUNG_HY_KY!$D$7=0),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=
    #   1),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD(DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=1)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),4.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),4.0)))), X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),1.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),1.0)))))) else '—')

def f201(E, r, c):
    # Ô K67
    # = "Dữ kiện: "&"Ấn: "&MENH_CUC!$E$10&" chữ, "&MENH_CUC!$G$10&IF(MENH_CUC!$D$10="",""," – "&MENH_CUC!$D$10)&"; "&"Thực thương: "&MENH_CUC!$E$7&" chữ, "&ME
    #   NH_CUC!$G$7&IF(MENH_CUC!$D$7="",""," – "&MENH_CUC!$D$7)&"; "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Kỵ PP1: "&IF
    #   (DUNG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Nhóm 
    #   nhiều (kỵ theo cân bằng) PP1: "&DUNG_HY_KY!$C$112&" / PP2: "&DUNG_HY_KY!$D$112&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.173): Yểu: nhật mạnh & ấn là kỵ; hoặc nhật nhược & thực thương là
    #    kỵ → "&IF(J67="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J67&") ⇒ "&G67&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Ấn: '),E.v(18,10,5)),' chữ, '),E.v(18,10,7)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(' – ',E.v(18,10,4)))),'; '),'Thực thương: '),E.v(18,7,5)),' chữ, '),E.v(18,7,7)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(' – ',E.v(18,7,4)))),'; '),'Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Nhóm nhiều (kỵ theo cân bằng) PP1: '),E.v(8,112,3)),' / PP2: '),E.v(8,112,4)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.173): Yểu: nhật mạnh & ấn là kỵ; hoặc nhật nhược & thực thương là kỵ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f202(E, r, c):
    # Ô E68
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$
    #   M$5:$M$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0))) else '—')

def f203(E, r, c):
    # Ô F68
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$
    #   M$5:$M$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0))) else '—')

def f204(E, r, c):
    # Ô K68
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" 
    #   ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Chính tài không bị khắc/hợp mất không bị xung: "&COU
    #   NTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên t
    #   ài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Thiên tài không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.164): Thân vượng, tài là dụng thần, t
    #   ài có mặt và không bị hợp/khắc/xung → "&IF(J68="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J68&") ⇒ "&G68&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P1 tr.164): Thân vượng, tài là dụng thần, tài có mặt và không bị hợp/khắc/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f205(E, r, c):
    # Ô E69
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$110>=0,MOD(DUNG_HY_KY!$C$110-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,110,3),0.0), X.eq(X.MOD(X.sub(E.v(8,110,3),E.v(23,7,5)), 5.0),2.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f206(E, r, c):
    # Ô F69
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$110>=0,MOD(DUNG_HY_KY!$D$110-TRUNG_GIAN!$E$7,5)=2)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,110,4),0.0), X.eq(X.MOD(X.sub(E.v(8,110,4),E.v(23,7,5)), 5.0),2.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f207(E, r, c):
    # Ô K69
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$
    #   C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$110+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("
    #   &INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO
    #   !$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0
    #   ,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr
    #   .165): Thân vượng, tài là hỷ thần (sinh dụng thần quan), tài vượng tại lệnh → "&IF(J69="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J69&") ⇒ "&G69&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.165): Thân vượng, tài là hỷ thần (sinh dụng thần quan), tài vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f208(E, r, c):
    # Ô E70, F70
    # = IF((PP1_DINH_TINH!$F$98=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(9,98,6),0.0)) else '—')

def f209(E, r, c):
    # Ô K70
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P1 tr.165): Không có tài (rõ và mờ) → "&IF(
    #   J70="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J70&") ⇒ "&G70&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P1 tr.165): Không có tài (rõ và mờ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f210(E, r, c):
    # Ô E71
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f211(E, r, c):
    # Ô F71
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f212(E, r, c):
    # Ô K71
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" 
    #   ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)
    #   >0,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1 
    #   tr.166): Thân nhược dùng tỉ kiếp, tài vượng → "&IF(J71="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J71&") ⇒ "&G71&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.166): Thân nhược dùng tỉ kiếp, tài vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f213(E, r, c):
    # Ô E72
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,S
    #   AO!$M$5:$M$20,0))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0))) else '—')

def f214(E, r, c):
    # Ô F72
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,S
    #   AO!$M$5:$M$20,0))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0))) else '—')

def f215(E, r, c):
    # Ô K72
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" 
    #   ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Chính tài không bị khắc/hợp mất không bị xung: "&COU
    #   NTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên t
    #   ài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Thiên tài không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Chính quan không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Thất sát không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.165-166): Thân nhược dùng ấn, có tài và quan không bị phá → "&IF(J72="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J72&") ⇒ "&G72&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')),(X.cat(X.cat('Chính quan không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')),(X.cat(X.cat('Thất sát không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P1 tr.165-166): Thân nhược dùng ấn, có tài và quan không bị phá → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f216(E, r, c):
    # Ô E73
    # = IF(AND(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1))>
    #   0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f217(E, r, c):
    # Ô F73
    # = IF(AND(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1))>
    #   0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f218(E, r, c):
    # Ô K73
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$L$5:$L$20,1)>0,"Chính tài thiên can bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1)>0
    #   ,"Thiên tài thiên can bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1)&
    #   " chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$F$5:$F$2
    #   0,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → 
    #   Quy tắc (P1 tr.166): Tài là dụng thần nhưng tài bị khắc/hợp mất/xung → "&IF(J73="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J7
    #   3&") ⇒ "&G73&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.166): Tài là dụng thần nhưng tài bị khắc/hợp mất/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f219(E, r, c):
    # Ô E74
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD
    #   (DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=2))),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),2.0)))), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f220(E, r, c):
    # Ô F74
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD
    #   (DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=2))),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),2.0)))), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f221(E, r, c):
    # Ô K74
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Kỵ PP1: "&IF(DUNG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$
    #   C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Nhóm nhiều (kỵ theo cân bằng) PP1: "&DUNG_HY_KY!$C$112
    #   &" / PP2: "&DUNG_HY_KY!$D$112&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" 
    #   chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy 
    #   tắc (P1 tr.166; P1 tr.204): Thân nhược, tài là nhóm nhiều nhất (tài mạnh) → "&IF(J74="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J74&") ⇒ "&G74&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Nhóm nhiều (kỵ theo cân bằng) PP1: '),E.v(8,112,3)),' / PP2: '),E.v(8,112,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.166; P1 tr.204): Thân nhược, tài là nhóm nhiều nhất (tài mạnh) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f222(E, r, c):
    # Ô E75
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$97=0)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.le(E.v(9,98,6),1.0)), X.eq(E.v(9,97,6),0.0)), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f223(E, r, c):
    # Ô F75
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$97=0)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.le(E.v(9,98,6),1.0)), X.eq(E.v(9,97,6),0.0)), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f224(E, r, c):
    # Ô K75
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TI
    #   NH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P1 tr.166): Thân vượng, tài yếu (≤1 chữ), k
    #   hông có thực thương, tỉ kiếp nhiều → "&IF(J75="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J75&") ⇒ "&G75&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.166): Thân vượng, tài yếu (≤1 chữ), không có thực thương, tỉ kiếp nhiều → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f225(E, r, c):
    # Ô E76
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f226(E, r, c):
    # Ô F76
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f227(E, r, c):
    # Ô K76
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.166): Thân nhược, dụng ấn, tài nhiều phá ấn → "&IF(J76="","chưa đủ điều kiện 
    #   ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J76&") ⇒ "&G76&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.166): Thân nhược, dụng ấn, tài nhiều phá ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f228(E, r, c):
    # Ô E77
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$
    #   M$5:$M$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0))) else '—')

def f229(E, r, c):
    # Ô F77
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$
    #   M$5:$M$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0))) else '—')

def f230(E, r, c):
    # Ô K77
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" 
    #   ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Chính quan không bị khắc/hợp mất không bị xung: "&C
    #   OUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất
    #    sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Thất sát không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,
    #   "Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.167): Thân vượng, quan là dụng thần, 
    #   có mặt và không bị hợp/khắc/xung → "&IF(J77="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J77&") ⇒ "&G77&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P1 tr.167): Thân vượng, quan là dụng thần, có mặt và không bị hợp/khắc/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f231(E, r, c):
    # Ô E78
    # = IF(AND(AND(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f232(E, r, c):
    # Ô F78
    # = IF(AND(AND(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f233(E, r, c):
    # Ô K78
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_
    #   GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(
    #   IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"C
    #   hính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$J$5:$J$20,1)>0,"Thất sát vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" →
    #    Quy tắc (P1 tr.168): Quan vượng làm dụng, có ấn bảo vệ → "&IF(J78="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J78&") ⇒ "&G78&
    #   ".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.168): Quan vượng làm dụng, có ấn bảo vệ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f234(E, r, c):
    # Ô E79
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f235(E, r, c):
    # Ô F79
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f236(E, r, c):
    # Ô K79
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thất sát vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1
    #   )>0,"Thực thần vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1
    #    tr.168): Thân vượng, sát mạnh, thực thần có lực (thực thần chế sát) → "&IF(J79="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J7
    #   9&") ⇒ "&G79&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.168): Thân vượng, sát mạnh, thực thần có lực (thực thần chế sát) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f237(E, r, c):
    # Ô E80, F80
    # = IF(AND((PP1_DINH_TINH!$F$99>=P_NHIEU),(PP1_DINH_TINH!$F$100=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,99,6),E.v(24,99,2)), X.eq(E.v(9,100,6),0.0))) else '—')

def f238(E, r, c):
    # Ô K80
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC
    #   !$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P1 tr.169): Quan sát nặng (≥ P_NHIEU), không có ấn → "&IF(J80
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J80&") ⇒ "&G80&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.169): Quan sát nặng (≥ P_NHIEU), không có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f239(E, r, c):
    # Ô E81
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$99<=1)),(PP1_DINH_TINH!$F$100>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.le(E.v(9,99,6),1.0)), X.ge(E.v(9,100,6),E.v(24,99,2)))) else '—')

def f240(E, r, c):
    # Ô F81
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$99<=1)),(PP1_DINH_TINH!$F$100>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.le(E.v(9,99,6),1.0)), X.ge(E.v(9,100,6),E.v(24,99,2)))) else '—')

def f241(E, r, c):
    # Ô K81
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""
    #   ," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ
    #   "&" → Quy tắc (P1 tr.169): Quan nhẹ (≤1), ấn nặng (≥ P_NHIEU), thân vượng → "&IF(J81="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J81&") ⇒ "&G81&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.169): Quan nhẹ (≤1), ấn nặng (≥ P_NHIEU), thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f242(E, r, c):
    # Ô E82, F82
    # = IF(AND(AND((PP1_DINH_TINH!$F$99<=1),(PP1_DINH_TINH!$F$96>=P_NHIEU)),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.le(E.v(9,99,6),1.0), X.ge(E.v(9,96,6),E.v(24,99,2))), X.eq(E.v(9,98,6),0.0))) else '—')

def f243(E, r, c):
    # Ô K82
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&
    #   P_NHIEU&" chữ"&" → Quy tắc (P1 tr.169): Quan nhược (≤1), kiếp nặng (≥ P_NHIEU), không có tài → "&IF(J82="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J82&") ⇒ "&G82&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.169): Quan nhược (≤1), kiếp nặng (≥ P_NHIEU), không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f244(E, r, c):
    # Ô E83
    # = IF(AND((DUNG_HY_KY!$C$114="Không"),IF(DUNG_HY_KY!$C$110<0,FALSE(),INDEX(DUNG_HY_KY!$C$24:$C$28,DUNG_HY_KY!$C$110+1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,114,3),'Không'), (X.FALSE() if X.truthy(X.lt(E.v(8,110,3),0.0)) else X.gt(X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,110,3),1.0)),0.0)))) else '—')

def f245(E, r, c):
    # Ô F83
    # = IF(AND((DUNG_HY_KY!$D$114="Không"),IF(DUNG_HY_KY!$D$110<0,FALSE(),INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$110+1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,114,4),'Không'), (X.FALSE() if X.truthy(X.lt(E.v(8,110,4),0.0)) else X.gt(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,110,4),1.0)),0.0)))) else '—')

def f246(E, r, c):
    # Ô K83
    # = "Dữ kiện: "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$
    #   D$110+1),"—")&"; "&"Dụng thần khuyết? PP1: "&DUNG_HY_KY!$C$114&" / PP2: "&DUNG_HY_KY!$D$114&" → Quy tắc (P1 tr.170): Dụng thần có mặt, có hỷ thần sinh
    #    (hỷ có mặt), dụng không bị khắc/xung → "&IF(J83="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J83&") ⇒ "&G83&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Dụng thần khuyết? PP1: '),E.v(8,114,3)),' / PP2: '),E.v(8,114,4)),' → Quy tắc (P1 tr.170): Dụng thần có mặt, có hỷ thần sinh (hỷ có mặt), dụng không bị khắc/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f247(E, r, c):
    # Ô E84
    # = IF((DUNG_HY_KY!$C$114<>"Không"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ne(E.v(8,114,3),'Không')) else '—')

def f248(E, r, c):
    # Ô F84
    # = IF((DUNG_HY_KY!$D$114<>"Không"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ne(E.v(8,114,4),'Không')) else '—')

def f249(E, r, c):
    # Ô K84
    # = "Dữ kiện: "&"Dụng thần khuyết? PP1: "&DUNG_HY_KY!$C$114&" / PP2: "&DUNG_HY_KY!$D$114&" → Quy tắc (P1 tr.171): Không có dụng thần trong tứ trụ (khuyết)
    #    → "&IF(J84="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J84&") ⇒ "&G84&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng thần khuyết? PP1: '),E.v(8,114,3)),' / PP2: '),E.v(8,114,4)),' → Quy tắc (P1 tr.171): Không có dụng thần trong tứ trụ (khuyết) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f250(E, r, c):
    # Ô E85
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.ge(E.v(9,98,6),1.0))) else '—')

def f251(E, r, c):
    # Ô F85
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.ge(E.v(9,98,6),1.0))) else '—')

def f252(E, r, c):
    # Ô K85
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$E$9+1)&")"&" → Quy tắc (P1 tr.172): Thân vượng, dụng là quan, có tài sinh quan → "&IF(J85="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J85&") ⇒ "&G85&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.172): Thân vượng, dụng là quan, có tài sinh quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f253(E, r, c):
    # Ô E86
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.ge(E.v(9,97,6),1.0))) else '—')

def f254(E, r, c):
    # Ô F86
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.ge(E.v(9,97,6),1.0))) else '—')

def f255(E, r, c):
    # Ô K86
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CU
    #   C!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN
    #   ,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.172): Thân vượng, dụng là tài, có thực thương sinh tài → "&IF(J86="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J86&") ⇒ "&G86&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.172): Thân vượng, dụng là tài, có thực thương sinh tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f256(E, r, c):
    # Ô E87
    # = IF(AND((DUNG_HY_KY!$C$7=0),(DUNG_HY_KY!$C$9=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.eq(E.v(8,9,3),1.0))) else '—')

def f257(E, r, c):
    # Ô F87
    # = IF(AND((DUNG_HY_KY!$D$7=0),(DUNG_HY_KY!$C$9=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.eq(E.v(8,9,3),1.0))) else '—')

def f258(E, r, c):
    # Ô K87
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Ấn vượng tại lệnh: "&IF(DUNG_HY_KY!$C$9=1,"có","không")&" → Q
    #   uy tắc (P1 tr.172): Thân nhược nhưng ấn vượng tại lệnh (ấn nắm quyền) → "&IF(J87="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J
    #   87&") ⇒ "&G87&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Ấn vượng tại lệnh: '),('có' if X.truthy(X.eq(E.v(8,9,3),1.0)) else 'không')),' → Quy tắc (P1 tr.172): Thân nhược nhưng ấn vượng tại lệnh (ấn nắm quyền) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f259(E, r, c):
    # Ô E88, F88
    # = IF((SUM(PP1_DINH_TINH!$C$24:$E$24)=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.SUM(E.rg(9,24,3,24,5)),0.0)) else '—')

def f260(E, r, c):
    # Ô K88
    # = "Dữ kiện: "&"Số cặp chi kề xung: "&SUM(PP1_DINH_TINH!$C$24:$E$24)&" → Quy tắc (P1 tr.172): Tứ trụ không có lục xung giữa các chi kề → "&IF(J88="","chư
    #   a đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J88&") ⇒ "&G88&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số cặp chi kề xung: '),X.SUM(E.rg(9,24,3,24,5))),' → Quy tắc (P1 tr.172): Tứ trụ không có lục xung giữa các chi kề → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f261(E, r, c):
    # Ô E89
    # = IF(OR((COUNTIF(PP1_DINH_TINH!$F$96:$F$100,0)=0),OR(DUNG_HY_KY!$C$6="Trung hòa")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(X.COUNTIF(E.rg(9,96,6,100,6), 0.0),0.0), X.OR(X.eq(E.v(8,6,3),'Trung hòa')))) else '—')

def f262(E, r, c):
    # Ô F89
    # = IF(OR((COUNTIF(PP1_DINH_TINH!$F$96:$F$100,0)=0),OR(DUNG_HY_KY!$D$6="Trung hòa")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(X.COUNTIF(E.rg(9,96,6,100,6), 0.0),0.0), X.OR(X.eq(E.v(8,6,4),'Trung hòa')))) else '—')

def f263(E, r, c):
    # Ô K89
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&" → Quy tắc (P1 tr.172): Ngũ hành đồng đều (đủ 5 hành) hoặc trung hòa → "&IF(J89="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J89&") ⇒ "&G89&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),' → Quy tắc (P1 tr.172): Ngũ hành đồng đều (đủ 5 hành) hoặc trung hòa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f264(E, r, c):
    # Ô E90
    # = IF(AND(AND(AND(OR(DUNG_HY_KY!$C$6="Quá vượng",DUNG_HY_KY!$C$6="Độc vượng"),(PP1_DINH_TINH!$F$97=0)),(PP1_DINH_TINH!$F$98=0)),(PP1_DINH_TINH!$F$99=0)),
    #   "ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.OR(X.eq(E.v(8,6,3),'Quá vượng'), X.eq(E.v(8,6,3),'Độc vượng')), X.eq(E.v(9,97,6),0.0)), X.eq(E.v(9,98,6),0.0)), X.eq(E.v(9,99,6),0.0))) else '—')

def f265(E, r, c):
    # Ô F90
    # = IF(AND(AND(AND(OR(DUNG_HY_KY!$D$6="Quá vượng",DUNG_HY_KY!$D$6="Độc vượng"),(PP1_DINH_TINH!$F$97=0)),(PP1_DINH_TINH!$F$98=0)),(PP1_DINH_TINH!$F$99=0)),
    #   "ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.OR(X.eq(E.v(8,6,4),'Quá vượng'), X.eq(E.v(8,6,4),'Độc vượng')), X.eq(E.v(9,97,6),0.0)), X.eq(E.v(9,98,6),0.0)), X.eq(E.v(9,99,6),0.0))) else '—')

def f266(E, r, c):
    # Ô K90
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_T
    #   INH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P1 tr.172): Thân quá vượng, không có thực thương, tài, quan sát → "&IF(J90="","c
    #   hưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J90&") ⇒ "&G90&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P1 tr.172): Thân quá vượng, không có thực thương, tài, quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f267(E, r, c):
    # Ô E91
    # = IF(AND(AND(OR(DUNG_HY_KY!$C$6="Quá nhược",DUNG_HY_KY!$C$6="Nhật chủ nhược"),(PP1_DINH_TINH!$F$100>=P_NHIEU)),(THONG_CAN!$E$71=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.eq(E.v(8,6,3),'Quá nhược'), X.eq(E.v(8,6,3),'Nhật chủ nhược')), X.ge(E.v(9,100,6),E.v(24,99,2))), X.eq(E.v(11,71,5),0.0))) else '—')

def f268(E, r, c):
    # Ô F91
    # = IF(AND(AND(OR(DUNG_HY_KY!$D$6="Quá nhược",DUNG_HY_KY!$D$6="Nhật chủ nhược"),(PP1_DINH_TINH!$F$100>=P_NHIEU)),(THONG_CAN!$E$71=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.eq(E.v(8,6,4),'Quá nhược'), X.eq(E.v(8,6,4),'Nhật chủ nhược')), X.ge(E.v(9,100,6),E.v(24,99,2))), X.eq(E.v(11,71,5),0.0))) else '—')

def f269(E, r, c):
    # Ô K91
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," (
    #   "&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Nhật chủ có gốc dùng được: "&IF(THONG_CAN!$E$71=1,"có","không")&" → Quy tắc (P1 tr.1
    #   72): Thân quá nhược, ấn quá nặng (≥ P_NHIEU) và nhật chủ không có gốc → "&IF(J91="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J
    #   91&") ⇒ "&G91&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Nhật chủ có gốc dùng được: '),('có' if X.truthy(X.eq(E.v(11,71,5),1.0)) else 'không')),' → Quy tắc (P1 tr.172): Thân quá nhược, ấn quá nặng (≥ P_NHIEU) và nhật chủ không có gốc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f270(E, r, c):
    # Ô E92
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$100=0)),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.eq(E.v(9,100,6),0.0)), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f271(E, r, c):
    # Ô F92
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$100=0)),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.eq(E.v(9,100,6),0.0)), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f272(E, r, c):
    # Ô K92
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" 
    #   chữ"&" → Quy tắc (P1 tr.172): Thân nhược, không có ấn, thực thương trùng điệp (≥ P_NHIEU) → "&IF(J92="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J92&") ⇒ "&G92&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.172): Thân nhược, không có ấn, thực thương trùng điệp (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f273(E, r, c):
    # Ô E93
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$100=0)),(PP1_DINH_TINH!$F$96=0)),((PP1_DINH_TINH!$F$98+PP1_DINH_TINH!$F$99)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.eq(E.v(9,100,6),0.0)), X.eq(E.v(9,96,6),0.0)), X.ge(X.add(E.v(9,98,6),E.v(9,99,6)),E.v(24,99,2)))) else '—')

def f274(E, r, c):
    # Ô F93
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$100=0)),(PP1_DINH_TINH!$F$96=0)),((PP1_DINH_TINH!$F$98+PP1_DINH_TINH!$F$99)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.eq(E.v(9,100,6),0.0)), X.eq(E.v(9,96,6),0.0)), X.ge(X.add(E.v(9,98,6),E.v(9,99,6)),E.v(24,99,2)))) else '—')

def f275(E, r, c):
    # Ô K93
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!
    #   $F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ng
    #   ưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P1 tr.172): Thân nhược, không có ấn và tỉ kiếp, tài + quan sát nhiều → "&IF(J93="","chưa đủ điều kiện ở cả 
    #   hai phương pháp – không áp dụng.","ĐỦ ("&J93&") ⇒ "&G93&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.172): Thân nhược, không có ấn và tỉ kiếp, tài + quan sát nhiều → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f276(E, r, c):
    # Ô E94, F94
    # = IF(OR(AND(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),INDEX(DUNG_HY_KY!$C$24:$C$28,2)=0),AND(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRU
    #   NG_GIAN!$D$6=7),INDEX(DUNG_HY_KY!$C$24:$C$28,5)=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.eq(X.INDEX(E.rg(8,24,3,28,3), 2.0),0.0)), X.AND(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0)), X.eq(X.INDEX(E.rg(8,24,3,28,3), 5.0),0.0)))) else '—')

def f277(E, r, c):
    # Ô K94
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&" → Q
    #   uy tắc (P1 tr.172): Sinh tháng Hợi, Tí, Sửu mà không có hỏa; hoặc sinh tháng Tị, Ngọ, Mùi mà không có thủy → "&IF(J94="","chưa đủ điều kiện ở cả hai p
    #   hương pháp – không áp dụng.","ĐỦ ("&J94&") ⇒ "&G94&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),' → Quy tắc (P1 tr.172): Sinh tháng Hợi, Tí, Sửu mà không có hỏa; hoặc sinh tháng Tị, Ngọ, Mùi mà không có thủy → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(18,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(18,r,c-1)),') ⇒ '),E.v(18,r,c-4)),'.')))

def f278(E, r, c):
    # Ô C97
    # = IF(LEFT(DOI_CHIEU!$C$34,5)="VƯỢNG",1,0)
    return (1.0 if X.truthy(X.eq(X.LEFT(E.v(7,34,3), 5.0),'VƯỢNG')) else 0.0)

def f279(E, r, c):
    # Ô C98
    # = IF(THAN_SAT!$H$86="Có",1,0)
    return (1.0 if X.truthy(X.eq(E.v(12,86,8),'Có')) else 0.0)

def f280(E, r, c):
    # Ô C99
    # = IF(THAN_SAT!$H$85="Có",1,0)
    return (1.0 if X.truthy(X.eq(E.v(12,85,8),'Có')) else 0.0)

def f281(E, r, c):
    # Ô C100
    # = IF(LEFT(SUC_KHOE!$J$6,5)="THUẬN",1,0)
    return (1.0 if X.truthy(X.eq(X.LEFT(E.v(22,6,10), 5.0),'THUẬN')) else 0.0)

def f282(E, r, c):
    # Ô C101
    # = IF(AND(ISNUMBER(SEARCH("CÓ CỨU",DUNG_HY_KY!$C$174)),NOT(ISNUMBER(SEARCH("KHÔNG CÓ CỨU",DUNG_HY_KY!$C$174))),NOT(ISNUMBER(SEARCH("KHÔNG CỨU",DUNG_HY_KY
    #   !$C$174)))),1,0)
    return (1.0 if X.truthy(X.AND(X.ISNUMBER(lambda: X.SEARCH('CÓ CỨU', E.v(8,174,3))), X.NOT(X.ISNUMBER(lambda: X.SEARCH('KHÔNG CÓ CỨU', E.v(8,174,3)))), X.NOT(X.ISNUMBER(lambda: X.SEARCH('KHÔNG CỨU', E.v(8,174,3)))))) else 0.0)

def f283(E, r, c):
    # Ô C102
    # = IF(ISNUMBER(SEARCH("ít bệnh",KET_LUAN!$C$156)),1,0)
    return (1.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ít bệnh', E.v(3,156,3)))) else 0.0)

def f284(E, r, c):
    # Ô C103
    # = SUM(C97:C102)
    return X.SUM(E.rg(18,r-6,c,r-1,c))

def f285(E, r, c):
    # Ô C104
    # = IF(AND($U$21="YỂU",C103>=2),"CĂN CỨ NGƯỢC CHIỀU","")
    return ('CĂN CỨ NGƯỢC CHIỀU' if X.truthy(X.AND(X.eq(E.v(18,21,21),'YỂU'), X.ge(E.v(18,r-1,c),2.0))) else '')

def f286(E, r, c):
    # Ô C105
    # = IF(C104="","","Bảng P1 tr.173 nghiêng về yểu (chỉ so Thực thương với Ấn), nhưng có "&C103&" dấu hiệu ngược chiều: "&MID(IF(C97=1,", thân vượng","")&IF
    #   (C98=1,", nguyệt đức","")&IF(C99=1,", thiên đức","")&IF(C100=1,", sức khỏe chung thuận","")&IF(C101=1,", có cứu","")&IF(C102=1,", bình sinh ít bệnh","
    #   "),3,300)&". Đọc là thể chất nền tốt, đời có giai đoạn thử thách sức khỏe – chú ý các năm hạn.")
    return ('' if X.truthy(X.eq(E.v(18,r-1,c),'')) else X.cat(X.cat(X.cat(X.cat('Bảng P1 tr.173 nghiêng về yểu (chỉ so Thực thương với Ấn), nhưng có ',E.v(18,r-2,c)),' dấu hiệu ngược chiều: '),X.MID(X.cat(X.cat(X.cat(X.cat(X.cat((', thân vượng' if X.truthy(X.eq(E.v(18,r-8,c),1.0)) else ''),(', nguyệt đức' if X.truthy(X.eq(E.v(18,r-7,c),1.0)) else '')),(', thiên đức' if X.truthy(X.eq(E.v(18,r-6,c),1.0)) else '')),(', sức khỏe chung thuận' if X.truthy(X.eq(E.v(18,r-5,c),1.0)) else '')),(', có cứu' if X.truthy(X.eq(E.v(18,r-4,c),1.0)) else '')),(', bình sinh ít bệnh' if X.truthy(X.eq(E.v(18,r-3,c),1.0)) else '')), 3.0, 300.0)),'. Đọc là thể chất nền tốt, đời có giai đoạn thử thách sức khỏe – chú ý các năm hạn.'))

CELLS = {
    'C6': (f0, 6, 3),
    'D6': (f1, 6, 4),
    'E6': (f2, 6, 5),
    'F6': (f3, 6, 6),
    'G6': (f4, 6, 7),
    'H6': (f5, 6, 8),
    'I6': (f6, 6, 9),
    'C7': (f7, 7, 3),
    'D7': (f8, 7, 4),
    'E7': (f9, 7, 5),
    'F7': (f10, 7, 6),
    'G7': (f11, 7, 7),
    'H7': (f12, 7, 8),
    'I7': (f13, 7, 9),
    'C8': (f14, 8, 3),
    'D8': (f15, 8, 4),
    'E8': (f16, 8, 5),
    'F8': (f17, 8, 6),
    'G8': (f18, 8, 7),
    'H8': (f19, 8, 8),
    'I8': (f20, 8, 9),
    'C9': (f21, 9, 3),
    'D9': (f22, 9, 4),
    'E9': (f23, 9, 5),
    'F9': (f24, 9, 6),
    'G9': (f25, 9, 7),
    'H9': (f26, 9, 8),
    'I9': (f27, 9, 9),
    'C10': (f28, 10, 3),
    'D10': (f29, 10, 4),
    'E10': (f30, 10, 5),
    'F10': (f31, 10, 6),
    'G10': (f32, 10, 7),
    'H10': (f33, 10, 8),
    'I10': (f34, 10, 9),
    'D12': (f35, 12, 4),
    'D13': (f36, 13, 4),
    'D14': (f37, 14, 4),
    'D18': (f38, 18, 4),
    'G18': (f39, 18, 7),
    'K18': (f40, 18, 11),
    'M18': (f41, 18, 13),
    'N18': (f42, 18, 14),
    'O18': (f43, 18, 15),
    'P18': (f44, 18, 16),
    'Q18': (f45, 18, 17),
    'R18': (f46, 18, 18),
    'S18': (f47, 18, 19),
    'T18': (f48, 18, 20),
    'U18': (f49, 18, 21),
    'V18': (f50, 18, 22),
    'W18': (f51, 18, 23),
    'D19': (f38, 19, 4),
    'G19': (f52, 19, 7),
    'K19': (f53, 19, 11),
    'M19': (f54, 19, 13),
    'N19': (f55, 19, 14),
    'O19': (f56, 19, 15),
    'P19': (f57, 19, 16),
    'Q19': (f58, 19, 17),
    'R19': (f59, 19, 18),
    'S19': (f60, 19, 19),
    'T19': (f61, 19, 20),
    'U19': (f49, 19, 21),
    'V19': (f62, 19, 22),
    'W19': (f63, 19, 23),
    'D20': (f64, 20, 4),
    'G20': (f65, 20, 7),
    'K20': (f66, 20, 11),
    'M20': (f67, 20, 13),
    'N20': (f68, 20, 14),
    'O20': (f69, 20, 15),
    'P20': (f70, 20, 16),
    'Q20': (f71, 20, 17),
    'R20': (f72, 20, 18),
    'S20': (f73, 20, 19),
    'T20': (f74, 20, 20),
    'U20': (f49, 20, 21),
    'V20': (f75, 20, 22),
    'W20': (f76, 20, 23),
    'D21': (f77, 21, 4),
    'G21': (f78, 21, 7),
    'K21': (f79, 21, 11),
    'M21': (f80, 21, 13),
    'N21': (f81, 21, 14),
    'O21': (f82, 21, 15),
    'P21': (f83, 21, 16),
    'Q21': (f84, 21, 17),
    'R21': (f85, 21, 18),
    'S21': (f86, 21, 19),
    'T21': (f87, 21, 20),
    'U21': (f49, 21, 21),
    'V21': (f88, 21, 22),
    'W21': (f89, 21, 23),
    'D26': (f90, 26, 4),
    'D27': (f91, 27, 4),
    'D28': (f92, 28, 4),
    'C33': (f93, 33, 3),
    'D33': (f94, 33, 4),
    'E33': (f95, 33, 5),
    'F33': (f96, 33, 6),
    'G33': (f97, 33, 7),
    'K33': (f98, 33, 11),
    'C34': (f99, 34, 3),
    'D34': (f100, 34, 4),
    'E34': (f101, 34, 5),
    'F34': (f102, 34, 6),
    'G34': (f103, 34, 7),
    'K34': (f104, 34, 11),
    'C35': (f105, 35, 3),
    'D35': (f106, 35, 4),
    'E35': (f107, 35, 5),
    'F35': (f108, 35, 6),
    'G35': (f109, 35, 7),
    'K35': (f110, 35, 11),
    'C36': (f111, 36, 3),
    'D36': (f112, 36, 4),
    'E36': (f113, 36, 5),
    'F36': (f114, 36, 6),
    'G36': (f115, 36, 7),
    'K36': (f116, 36, 11),
    'C37': (f117, 37, 3),
    'D37': (f118, 37, 4),
    'E37': (f119, 37, 5),
    'F37': (f120, 37, 6),
    'G37': (f121, 37, 7),
    'K37': (f122, 37, 11),
    'C38': (f123, 38, 3),
    'D38': (f124, 38, 4),
    'E38': (f125, 38, 5),
    'F38': (f126, 38, 6),
    'G38': (f127, 38, 7),
    'K38': (f128, 38, 11),
    'C39': (f129, 39, 3),
    'D39': (f130, 39, 4),
    'E39': (f131, 39, 5),
    'F39': (f132, 39, 6),
    'G39': (f133, 39, 7),
    'K39': (f134, 39, 11),
    'C40': (f135, 40, 3),
    'D40': (f136, 40, 4),
    'E40': (f137, 40, 5),
    'F40': (f138, 40, 6),
    'G40': (f139, 40, 7),
    'K40': (f140, 40, 11),
    'C41': (f141, 41, 3),
    'D41': (f142, 41, 4),
    'E41': (f143, 41, 5),
    'F41': (f144, 41, 6),
    'G41': (f145, 41, 7),
    'K41': (f146, 41, 11),
    'C42': (f147, 42, 3),
    'D42': (f148, 42, 4),
    'E42': (f149, 42, 5),
    'F42': (f150, 42, 6),
    'G42': (f151, 42, 7),
    'K42': (f152, 42, 11),
    'C43': (f153, 43, 3),
    'D43': (f154, 43, 4),
    'E43': (f155, 43, 5),
    'F43': (f156, 43, 6),
    'G43': (f157, 43, 7),
    'K43': (f158, 43, 11),
    'C44': (f159, 44, 3),
    'D44': (f160, 44, 4),
    'E44': (f161, 44, 5),
    'F44': (f162, 44, 6),
    'G44': (f163, 44, 7),
    'K44': (f164, 44, 11),
    'C45': (f165, 45, 3),
    'D45': (f166, 45, 4),
    'E45': (f167, 45, 5),
    'F45': (f168, 45, 6),
    'G45': (f169, 45, 7),
    'K45': (f170, 45, 11),
    'C46': (f171, 46, 3),
    'D46': (f172, 46, 4),
    'E46': (f173, 46, 5),
    'F46': (f174, 46, 6),
    'G46': (f175, 46, 7),
    'K46': (f176, 46, 11),
    'C47': (f177, 47, 3),
    'D47': (f178, 47, 4),
    'E47': (f179, 47, 5),
    'F47': (f180, 47, 6),
    'G47': (f181, 47, 7),
    'K47': (f182, 47, 11),
    'E62': (f183, 62, 5),
    'F62': (f184, 62, 6),
    'J62': (f185, 62, 10),
    'K62': (f186, 62, 11),
    'E63': (f187, 63, 5),
    'F63': (f188, 63, 6),
    'J63': (f185, 63, 10),
    'K63': (f189, 63, 11),
    'E64': (f190, 64, 5),
    'F64': (f191, 64, 6),
    'J64': (f185, 64, 10),
    'K64': (f192, 64, 11),
    'E65': (f193, 65, 5),
    'F65': (f194, 65, 6),
    'J65': (f185, 65, 10),
    'K65': (f195, 65, 11),
    'E66': (f196, 66, 5),
    'F66': (f197, 66, 6),
    'J66': (f185, 66, 10),
    'K66': (f198, 66, 11),
    'E67': (f199, 67, 5),
    'F67': (f200, 67, 6),
    'J67': (f185, 67, 10),
    'K67': (f201, 67, 11),
    'E68': (f202, 68, 5),
    'F68': (f203, 68, 6),
    'J68': (f185, 68, 10),
    'K68': (f204, 68, 11),
    'E69': (f205, 69, 5),
    'F69': (f206, 69, 6),
    'J69': (f185, 69, 10),
    'K69': (f207, 69, 11),
    'E70': (f208, 70, 5),
    'F70': (f208, 70, 6),
    'J70': (f185, 70, 10),
    'K70': (f209, 70, 11),
    'E71': (f210, 71, 5),
    'F71': (f211, 71, 6),
    'J71': (f185, 71, 10),
    'K71': (f212, 71, 11),
    'E72': (f213, 72, 5),
    'F72': (f214, 72, 6),
    'J72': (f185, 72, 10),
    'K72': (f215, 72, 11),
    'E73': (f216, 73, 5),
    'F73': (f217, 73, 6),
    'J73': (f185, 73, 10),
    'K73': (f218, 73, 11),
    'E74': (f219, 74, 5),
    'F74': (f220, 74, 6),
    'J74': (f185, 74, 10),
    'K74': (f221, 74, 11),
    'E75': (f222, 75, 5),
    'F75': (f223, 75, 6),
    'J75': (f185, 75, 10),
    'K75': (f224, 75, 11),
    'E76': (f225, 76, 5),
    'F76': (f226, 76, 6),
    'J76': (f185, 76, 10),
    'K76': (f227, 76, 11),
    'E77': (f228, 77, 5),
    'F77': (f229, 77, 6),
    'J77': (f185, 77, 10),
    'K77': (f230, 77, 11),
    'E78': (f231, 78, 5),
    'F78': (f232, 78, 6),
    'J78': (f185, 78, 10),
    'K78': (f233, 78, 11),
    'E79': (f234, 79, 5),
    'F79': (f235, 79, 6),
    'J79': (f185, 79, 10),
    'K79': (f236, 79, 11),
    'E80': (f237, 80, 5),
    'F80': (f237, 80, 6),
    'J80': (f185, 80, 10),
    'K80': (f238, 80, 11),
    'E81': (f239, 81, 5),
    'F81': (f240, 81, 6),
    'J81': (f185, 81, 10),
    'K81': (f241, 81, 11),
    'E82': (f242, 82, 5),
    'F82': (f242, 82, 6),
    'J82': (f185, 82, 10),
    'K82': (f243, 82, 11),
    'E83': (f244, 83, 5),
    'F83': (f245, 83, 6),
    'J83': (f185, 83, 10),
    'K83': (f246, 83, 11),
    'E84': (f247, 84, 5),
    'F84': (f248, 84, 6),
    'J84': (f185, 84, 10),
    'K84': (f249, 84, 11),
    'E85': (f250, 85, 5),
    'F85': (f251, 85, 6),
    'J85': (f185, 85, 10),
    'K85': (f252, 85, 11),
    'E86': (f253, 86, 5),
    'F86': (f254, 86, 6),
    'J86': (f185, 86, 10),
    'K86': (f255, 86, 11),
    'E87': (f256, 87, 5),
    'F87': (f257, 87, 6),
    'J87': (f185, 87, 10),
    'K87': (f258, 87, 11),
    'E88': (f259, 88, 5),
    'F88': (f259, 88, 6),
    'J88': (f185, 88, 10),
    'K88': (f260, 88, 11),
    'E89': (f261, 89, 5),
    'F89': (f262, 89, 6),
    'J89': (f185, 89, 10),
    'K89': (f263, 89, 11),
    'E90': (f264, 90, 5),
    'F90': (f265, 90, 6),
    'J90': (f185, 90, 10),
    'K90': (f266, 90, 11),
    'E91': (f267, 91, 5),
    'F91': (f268, 91, 6),
    'J91': (f185, 91, 10),
    'K91': (f269, 91, 11),
    'E92': (f270, 92, 5),
    'F92': (f271, 92, 6),
    'J92': (f185, 92, 10),
    'K92': (f272, 92, 11),
    'E93': (f273, 93, 5),
    'F93': (f274, 93, 6),
    'J93': (f185, 93, 10),
    'K93': (f275, 93, 11),
    'E94': (f276, 94, 5),
    'F94': (f276, 94, 6),
    'J94': (f185, 94, 10),
    'K94': (f277, 94, 11),
    'C97': (f278, 97, 3),
    'C98': (f279, 98, 3),
    'C99': (f280, 99, 3),
    'C100': (f281, 100, 3),
    'C101': (f282, 101, 3),
    'C102': (f283, 102, 3),
    'C103': (f284, 103, 3),
    'C104': (f285, 104, 3),
    'C105': (f286, 105, 3),
}
