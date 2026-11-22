# -*- coding: utf-8 -*-
"""Module sheet HON_NHAN – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: B69n.xlsx (SHA-256 9228d1ab613379f50ca7126e315efcbb940599c6e1784212c489756674ca8b67).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'HON_NHAN'
SI = 17

def f0(E, r, c):
    # Ô D5
    # = INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)
    return X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0))

def f1(E, r, c):
    # Ô E5
    # = "Bản khí "&SAO!$E$14&" = "&SAO!$F$14&"; nhật can tại chi ngày: "&INDEX(STAGE_TEN,SAO!$K$13+1)&IF(SAO!$M$14=1,"; chi ngày bị chi kề xung","")&IF(SAO!$N
    #   $14=1,"; tuần không","")&IF(SAO!$R$14=1,"; là kình dương","")&IF(SAO!$S$14=1,"; là đào hoa","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bản khí ',E.v(13,14,5)),' = '),E.v(13,14,6)),'; nhật can tại chi ngày: '),X.INDEX(E.rg(24,40,1,51,1), X.add(E.v(13,13,11),1.0))),('; chi ngày bị chi kề xung' if X.truthy(X.eq(E.v(13,14,13),1.0)) else '')),('; tuần không' if X.truthy(X.eq(E.v(13,14,14),1.0)) else '')),('; là kình dương' if X.truthy(X.eq(E.v(13,14,18),1.0)) else '')),('; là đào hoa' if X.truthy(X.eq(E.v(13,14,19),1.0)) else ''))

def f2(E, r, c):
    # Ô D6
    # = IF(NHAP_LIEU!$C$22="Nam","Tài (vợ)","Quan sát (chồng)")
    return ('Tài (vợ)' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else 'Quan sát (chồng)')

def f3(E, r, c):
    # Ô E6
    # = "Có "&IF(NHAP_LIEU!$C$22="Nam",(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$X$5:$X
    #   $20,1,SAO!$Y$5:$Y$20,0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,1,
    #   SAO!$Y$5:$Y$20,0)))&" chữ: "&IF(NHAP_LIEU!$C$22="Nam",TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,OR(SAO!$F$5="Chính tài",SAO!$F$5="Thiên tài")),SAO!$E$5&"(Năm)
    #    ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,OR(SAO!$F$6="Chính tài",SAO!$F$6="Thiên tài")),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$7=1,SAO!$Y$7=0,OR(SAO!$F$7
    #   ="Chính tài",SAO!$F$7="Thiên tài")),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,OR(SAO!$F$8="Chính tài",SAO!$F$8="Thiên tài")),SAO!$E$8&"(
    #   Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0,OR(SAO!$F$9="Chính tài",SAO!$F$9="Thiên tài")),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,OR(
    #   SAO!$F$10="Chính tài",SAO!$F$10="Thiên tài")),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,OR(SAO!$F$11="Chính tài",SAO!$F$11="Thiên t
    #   ài")),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,OR(SAO!$F$12="Chính tài",SAO!$F$12="Thiên tài")),SAO!$E$12&"(Tháng tàng) ","")&IF(A
    #   ND(SAO!$X$13=1,SAO!$Y$13=0,OR(SAO!$F$13="Chính tài",SAO!$F$13="Thiên tài")),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,OR(SAO!$F$14="Chính
    #    tài",SAO!$F$14="Thiên tài")),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,OR(SAO!$F$15="Chính tài",SAO!$F$15="Thiên tài")),SAO!$E$15&"
    #   (Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,OR(SAO!$F$16="Chính tài",SAO!$F$16="Thiên tài")),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!
    #   $Y$17=0,OR(SAO!$F$17="Chính tài",SAO!$F$17="Thiên tài")),SAO!$E$17&"(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,OR(SAO!$F$18="Chính tài",SAO!$F$18="Thi
    #   ên tài")),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,OR(SAO!$F$19="Chính tài",SAO!$F$19="Thiên tài")),SAO!$E$19&"(Giờ tàng) ","")&IF(A
    #   ND(SAO!$X$20=1,SAO!$Y$20=0,OR(SAO!$F$20="Chính tài",SAO!$F$20="Thiên tài")),SAO!$E$20&"(Giờ tàng) ","")),TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,OR(SAO!$F$5
    #   ="Chính quan",SAO!$F$5="Thất sát")),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,OR(SAO!$F$6="Chính quan",SAO!$F$6="Thất sát")),SAO!$E$6&"(Năm t
    #   àng) ","")&IF(AND(SAO!$X$7=1,SAO!$Y$7=0,OR(SAO!$F$7="Chính quan",SAO!$F$7="Thất sát")),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,OR(SAO!
    #   $F$8="Chính quan",SAO!$F$8="Thất sát")),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0,OR(SAO!$F$9="Chính quan",SAO!$F$9="Thất sát")),SAO!$E$
    #   9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,OR(SAO!$F$10="Chính quan",SAO!$F$10="Thất sát")),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!
    #   $Y$11=0,OR(SAO!$F$11="Chính quan",SAO!$F$11="Thất sát")),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,OR(SAO!$F$12="Chính quan",SAO!$F
    #   $12="Thất sát")),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,OR(SAO!$F$13="Chính quan",SAO!$F$13="Thất sát")),SAO!$E$13&"(Ngày) ","")
    #   &IF(AND(SAO!$X$14=1,SAO!$Y$14=0,OR(SAO!$F$14="Chính quan",SAO!$F$14="Thất sát")),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,OR(SAO!$F
    #   $15="Chính quan",SAO!$F$15="Thất sát")),SAO!$E$15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,OR(SAO!$F$16="Chính quan",SAO!$F$16="Thất sát")),S
    #   AO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,OR(SAO!$F$17="Chính quan",SAO!$F$17="Thất sát")),SAO!$E$17&"(Giờ) ","")&IF(AND(SAO!$X$18=1,
    #   SAO!$Y$18=0,OR(SAO!$F$18="Chính quan",SAO!$F$18="Thất sát")),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,OR(SAO!$F$19="Chính quan",SAO!
    #   $F$19="Thất sát")),SAO!$E$19&"(Giờ tàng) ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,OR(SAO!$F$20="Chính quan",SAO!$F$20="Thất sát")),SAO!$E$20&"(Giờ tàng) "
    #   ,"")))&"; vai trò PP1: "&IF(NHAP_LIEU!$C$22="Nam",MENH_CUC!$H$8,MENH_CUC!$H$9)&", PP2: "&IF(NHAP_LIEU!$C$22="Nam",MENH_CUC!$I$8,MENH_CUC!$I$9)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Có ',(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0)) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0)))),' chữ: '),(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.OR(X.eq(E.v(13,5,6),'Chính tài'), X.eq(E.v(13,5,6),'Thiên tài')))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.OR(X.eq(E.v(13,6,6),'Chính tài'), X.eq(E.v(13,6,6),'Thiên tài')))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.OR(X.eq(E.v(13,7,6),'Chính tài'), X.eq(E.v(13,7,6),'Thiên tài')))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.OR(X.eq(E.v(13,8,6),'Chính tài'), X.eq(E.v(13,8,6),'Thiên tài')))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.OR(X.eq(E.v(13,9,6),'Chính tài'), X.eq(E.v(13,9,6),'Thiên tài')))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.OR(X.eq(E.v(13,10,6),'Chính tài'), X.eq(E.v(13,10,6),'Thiên tài')))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.OR(X.eq(E.v(13,11,6),'Chính tài'), X.eq(E.v(13,11,6),'Thiên tài')))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.OR(X.eq(E.v(13,12,6),'Chính tài'), X.eq(E.v(13,12,6),'Thiên tài')))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.OR(X.eq(E.v(13,13,6),'Chính tài'), X.eq(E.v(13,13,6),'Thiên tài')))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.OR(X.eq(E.v(13,14,6),'Chính tài'), X.eq(E.v(13,14,6),'Thiên tài')))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.OR(X.eq(E.v(13,15,6),'Chính tài'), X.eq(E.v(13,15,6),'Thiên tài')))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.OR(X.eq(E.v(13,16,6),'Chính tài'), X.eq(E.v(13,16,6),'Thiên tài')))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.OR(X.eq(E.v(13,17,6),'Chính tài'), X.eq(E.v(13,17,6),'Thiên tài')))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.OR(X.eq(E.v(13,18,6),'Chính tài'), X.eq(E.v(13,18,6),'Thiên tài')))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.OR(X.eq(E.v(13,19,6),'Chính tài'), X.eq(E.v(13,19,6),'Thiên tài')))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.OR(X.eq(E.v(13,20,6),'Chính tài'), X.eq(E.v(13,20,6),'Thiên tài')))) else ''))) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.OR(X.eq(E.v(13,5,6),'Chính quan'), X.eq(E.v(13,5,6),'Thất sát')))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.OR(X.eq(E.v(13,6,6),'Chính quan'), X.eq(E.v(13,6,6),'Thất sát')))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.OR(X.eq(E.v(13,7,6),'Chính quan'), X.eq(E.v(13,7,6),'Thất sát')))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.OR(X.eq(E.v(13,8,6),'Chính quan'), X.eq(E.v(13,8,6),'Thất sát')))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.OR(X.eq(E.v(13,9,6),'Chính quan'), X.eq(E.v(13,9,6),'Thất sát')))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.OR(X.eq(E.v(13,10,6),'Chính quan'), X.eq(E.v(13,10,6),'Thất sát')))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.OR(X.eq(E.v(13,11,6),'Chính quan'), X.eq(E.v(13,11,6),'Thất sát')))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.OR(X.eq(E.v(13,12,6),'Chính quan'), X.eq(E.v(13,12,6),'Thất sát')))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.OR(X.eq(E.v(13,13,6),'Chính quan'), X.eq(E.v(13,13,6),'Thất sát')))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.OR(X.eq(E.v(13,14,6),'Chính quan'), X.eq(E.v(13,14,6),'Thất sát')))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.OR(X.eq(E.v(13,15,6),'Chính quan'), X.eq(E.v(13,15,6),'Thất sát')))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.OR(X.eq(E.v(13,16,6),'Chính quan'), X.eq(E.v(13,16,6),'Thất sát')))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.OR(X.eq(E.v(13,17,6),'Chính quan'), X.eq(E.v(13,17,6),'Thất sát')))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.OR(X.eq(E.v(13,18,6),'Chính quan'), X.eq(E.v(13,18,6),'Thất sát')))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.OR(X.eq(E.v(13,19,6),'Chính quan'), X.eq(E.v(13,19,6),'Thất sát')))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.OR(X.eq(E.v(13,20,6),'Chính quan'), X.eq(E.v(13,20,6),'Thất sát')))) else ''))))),'; vai trò PP1: '),(E.v(18,8,8) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else E.v(18,9,8))),', PP2: '),(E.v(18,8,9) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else E.v(18,9,9)))

def f4(E, r, c):
    # Ô D7
    # = IF(OR(LUC_THAN!$D$7=12,LUC_THAN!$D$7=13,LUC_THAN!$D$7=14,LUC_THAN!$D$7=27,LUC_THAN!$D$7=28,LUC_THAN!$D$7=29,LUC_THAN!$D$7=42,LUC_THAN!$D$7=43,LUC_THAN
    #   !$D$7=44,LUC_THAN!$D$7=57,LUC_THAN!$D$7=58,LUC_THAN!$D$7=59),"Có","Không")
    return ('Có' if X.truthy(X.OR(X.eq(E.v(16,7,4),12.0), X.eq(E.v(16,7,4),13.0), X.eq(E.v(16,7,4),14.0), X.eq(E.v(16,7,4),27.0), X.eq(E.v(16,7,4),28.0), X.eq(E.v(16,7,4),29.0), X.eq(E.v(16,7,4),42.0), X.eq(E.v(16,7,4),43.0), X.eq(E.v(16,7,4),44.0), X.eq(E.v(16,7,4),57.0), X.eq(E.v(16,7,4),58.0), X.eq(E.v(16,7,4),59.0))) else 'Không')

def f5(E, r, c):
    # Ô C11
    # = DAI_VAN!C36&"–"&(DAI_VAN!C36+9)
    return X.cat(X.cat(E.v(14,r+25,c),'–'),X.add(E.v(14,r+25,c),9.0))

def f6(E, r, c):
    # Ô D11
    # = DAI_VAN!C35&IF(DAI_VAN!C44<>""," ◀","")
    return X.cat(E.v(14,r+24,c-1),(' ◀' if X.truthy(X.ne(E.v(14,r+33,c-1),'')) else ''))

def f7(E, r, c):
    # Ô E11, F12
    # = IF(DAI_VAN!C84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+73,c-2),1.0)) else '')

def f8(E, r, c):
    # Ô F11
    # = IF(DAI_VAN!C85=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+74,c-3),1.0)) else '')

def f9(E, r, c):
    # Ô G11
    # = DAI_VAN!C54&" / "&DAI_VAN!C69
    return X.cat(X.cat(E.v(14,r+43,c-4),' / '),E.v(14,r+58,c-4))

def f10(E, r, c):
    # Ô H11, I12, J13
    # = IF(DAI_VAN!C86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+75,c-5),1.0)) else '')

def f11(E, r, c):
    # Ô I11, J12
    # = IF(DAI_VAN!C87=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+76,c-6),1.0)) else '')

def f12(E, r, c):
    # Ô J11
    # = IF(DAI_VAN!C88=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+77,c-7),1.0)) else '')

def f13(E, r, c):
    # Ô K11, K12, K13, K14, K15, K16, K17, K18
    # = TRIM(IF(OR(E11="có",F11="có"),"Thời kỳ thuận lợi kết hôn (P1 tr.201). ","")&IF(H11="có","Chi vận xung cung phối ngẫu: nhẹ thương tật, nặng khắc; ""chờ
    #    qua vận đó mới lấy vợ thì tránh được"" (P1 tr.156)"&IF(NHAP_LIEU!$C$22="Nữ"," – với nữ: lấy chồng","")&". ","")&IF(I11="có","Vận thương quan: nam khắ
    #   c vợ, nữ khắc chồng (P2 tr.471). ","")&IF(AND(J11="có",NHAP_LIEU!$C$22="Nam"),"Vận tỉ kiếp: tài nhược thì tổn vợ, hao tài (P1 tr.156; P2 tr.468). ",""
    #   ))
    return X.TRIM(X.cat(X.cat(X.cat(('Thời kỳ thuận lợi kết hôn (P1 tr.201). ' if X.truthy(X.OR(X.eq(E.v(17,r,c-6),'có'), X.eq(E.v(17,r,c-5),'có'))) else ''),(X.cat(X.cat('Chi vận xung cung phối ngẫu: nhẹ thương tật, nặng khắc; "chờ qua vận đó mới lấy vợ thì tránh được" (P1 tr.156)',(' – với nữ: lấy chồng' if X.truthy(X.eq(E.v(2,22,3),'Nữ')) else '')),'. ') if X.truthy(X.eq(E.v(17,r,c-3),'có')) else '')),('Vận thương quan: nam khắc vợ, nữ khắc chồng (P2 tr.471). ' if X.truthy(X.eq(E.v(17,r,c-2),'có')) else '')),('Vận tỉ kiếp: tài nhược thì tổn vợ, hao tài (P1 tr.156; P2 tr.468). ' if X.truthy(X.AND(X.eq(E.v(17,r,c-1),'có'), X.eq(E.v(2,22,3),'Nam'))) else '')))

def f14(E, r, c):
    # Ô C12
    # = DAI_VAN!D36&"–"&(DAI_VAN!D36+9)
    return X.cat(X.cat(E.v(14,r+24,c+1),'–'),X.add(E.v(14,r+24,c+1),9.0))

def f15(E, r, c):
    # Ô D12
    # = DAI_VAN!D35&IF(DAI_VAN!D44<>""," ◀","")
    return X.cat(E.v(14,r+23,c),(' ◀' if X.truthy(X.ne(E.v(14,r+32,c),'')) else ''))

def f16(E, r, c):
    # Ô E12, F13
    # = IF(DAI_VAN!D84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+72,c-1),1.0)) else '')

def f17(E, r, c):
    # Ô G12
    # = DAI_VAN!D54&" / "&DAI_VAN!D69
    return X.cat(X.cat(E.v(14,r+42,c-3),' / '),E.v(14,r+57,c-3))

def f18(E, r, c):
    # Ô H12, I13, J14
    # = IF(DAI_VAN!D86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+74,c-4),1.0)) else '')

def f19(E, r, c):
    # Ô C13
    # = DAI_VAN!E36&"–"&(DAI_VAN!E36+9)
    return X.cat(X.cat(E.v(14,r+23,c+2),'–'),X.add(E.v(14,r+23,c+2),9.0))

def f20(E, r, c):
    # Ô D13
    # = DAI_VAN!E35&IF(DAI_VAN!E44<>""," ◀","")
    return X.cat(E.v(14,r+22,c+1),(' ◀' if X.truthy(X.ne(E.v(14,r+31,c+1),'')) else ''))

def f21(E, r, c):
    # Ô E13, F14
    # = IF(DAI_VAN!E84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+71,c),1.0)) else '')

def f22(E, r, c):
    # Ô G13
    # = DAI_VAN!E54&" / "&DAI_VAN!E69
    return X.cat(X.cat(E.v(14,r+41,c-2),' / '),E.v(14,r+56,c-2))

def f23(E, r, c):
    # Ô H13, I14, J15
    # = IF(DAI_VAN!E86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+73,c-3),1.0)) else '')

def f24(E, r, c):
    # Ô C14
    # = DAI_VAN!F36&"–"&(DAI_VAN!F36+9)
    return X.cat(X.cat(E.v(14,r+22,c+3),'–'),X.add(E.v(14,r+22,c+3),9.0))

def f25(E, r, c):
    # Ô D14
    # = DAI_VAN!F35&IF(DAI_VAN!F44<>""," ◀","")
    return X.cat(E.v(14,r+21,c+2),(' ◀' if X.truthy(X.ne(E.v(14,r+30,c+2),'')) else ''))

def f26(E, r, c):
    # Ô E14, F15
    # = IF(DAI_VAN!F84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+70,c+1),1.0)) else '')

def f27(E, r, c):
    # Ô G14
    # = DAI_VAN!F54&" / "&DAI_VAN!F69
    return X.cat(X.cat(E.v(14,r+40,c-1),' / '),E.v(14,r+55,c-1))

def f28(E, r, c):
    # Ô H14, I15, J16
    # = IF(DAI_VAN!F86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+72,c-2),1.0)) else '')

def f29(E, r, c):
    # Ô C15
    # = DAI_VAN!G36&"–"&(DAI_VAN!G36+9)
    return X.cat(X.cat(E.v(14,r+21,c+4),'–'),X.add(E.v(14,r+21,c+4),9.0))

def f30(E, r, c):
    # Ô D15
    # = DAI_VAN!G35&IF(DAI_VAN!G44<>""," ◀","")
    return X.cat(E.v(14,r+20,c+3),(' ◀' if X.truthy(X.ne(E.v(14,r+29,c+3),'')) else ''))

def f31(E, r, c):
    # Ô E15, F16
    # = IF(DAI_VAN!G84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+69,c+2),1.0)) else '')

def f32(E, r, c):
    # Ô G15
    # = DAI_VAN!G54&" / "&DAI_VAN!G69
    return X.cat(X.cat(E.v(14,r+39,c),' / '),E.v(14,r+54,c))

def f33(E, r, c):
    # Ô H15, I16, J17
    # = IF(DAI_VAN!G86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+71,c-1),1.0)) else '')

def f34(E, r, c):
    # Ô C16
    # = DAI_VAN!H36&"–"&(DAI_VAN!H36+9)
    return X.cat(X.cat(E.v(14,r+20,c+5),'–'),X.add(E.v(14,r+20,c+5),9.0))

def f35(E, r, c):
    # Ô D16
    # = DAI_VAN!H35&IF(DAI_VAN!H44<>""," ◀","")
    return X.cat(E.v(14,r+19,c+4),(' ◀' if X.truthy(X.ne(E.v(14,r+28,c+4),'')) else ''))

def f36(E, r, c):
    # Ô E16, F17
    # = IF(DAI_VAN!H84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+68,c+3),1.0)) else '')

def f37(E, r, c):
    # Ô G16
    # = DAI_VAN!H54&" / "&DAI_VAN!H69
    return X.cat(X.cat(E.v(14,r+38,c+1),' / '),E.v(14,r+53,c+1))

def f38(E, r, c):
    # Ô H16, I17, J18
    # = IF(DAI_VAN!H86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+70,c),1.0)) else '')

def f39(E, r, c):
    # Ô C17
    # = DAI_VAN!I36&"–"&(DAI_VAN!I36+9)
    return X.cat(X.cat(E.v(14,r+19,c+6),'–'),X.add(E.v(14,r+19,c+6),9.0))

def f40(E, r, c):
    # Ô D17
    # = DAI_VAN!I35&IF(DAI_VAN!I44<>""," ◀","")
    return X.cat(E.v(14,r+18,c+5),(' ◀' if X.truthy(X.ne(E.v(14,r+27,c+5),'')) else ''))

def f41(E, r, c):
    # Ô E17, F18
    # = IF(DAI_VAN!I84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+67,c+4),1.0)) else '')

def f42(E, r, c):
    # Ô G17
    # = DAI_VAN!I54&" / "&DAI_VAN!I69
    return X.cat(X.cat(E.v(14,r+37,c+2),' / '),E.v(14,r+52,c+2))

def f43(E, r, c):
    # Ô H17, I18
    # = IF(DAI_VAN!I86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+69,c+1),1.0)) else '')

def f44(E, r, c):
    # Ô C18
    # = DAI_VAN!J36&"–"&(DAI_VAN!J36+9)
    return X.cat(X.cat(E.v(14,r+18,c+7),'–'),X.add(E.v(14,r+18,c+7),9.0))

def f45(E, r, c):
    # Ô D18
    # = DAI_VAN!J35&IF(DAI_VAN!J44<>""," ◀","")
    return X.cat(E.v(14,r+17,c+6),(' ◀' if X.truthy(X.ne(E.v(14,r+26,c+6),'')) else ''))

def f46(E, r, c):
    # Ô E18
    # = IF(DAI_VAN!J84=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+66,c+5),1.0)) else '')

def f47(E, r, c):
    # Ô G18
    # = DAI_VAN!J54&" / "&DAI_VAN!J69
    return X.cat(X.cat(E.v(14,r+36,c+3),' / '),E.v(14,r+51,c+3))

def f48(E, r, c):
    # Ô H18
    # = IF(DAI_VAN!J86=1,"có","")
    return ('có' if X.truthy(X.eq(E.v(14,r+68,c+2),1.0)) else '')

def f49(E, r, c):
    # Ô C22
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(1,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(1.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f50(E, r, c):
    # Ô D22
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(1,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(1,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(1.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(1.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f51(E, r, c):
    # Ô E22
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(1,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(1.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f52(E, r, c):
    # Ô G22
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(1,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(1.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f53(E, r, c):
    # Ô K22
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(1,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(1.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f54(E, r, c):
    # Ô C23
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(2,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(2.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f55(E, r, c):
    # Ô D23
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(2,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(2,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(2.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(2.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f56(E, r, c):
    # Ô E23
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(2,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(2.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f57(E, r, c):
    # Ô G23
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(2,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(2.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f58(E, r, c):
    # Ô K23
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(2,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(2.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f59(E, r, c):
    # Ô C24
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(3,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(3.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f60(E, r, c):
    # Ô D24
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(3,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(3,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(3.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(3.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f61(E, r, c):
    # Ô E24
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(3,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(3.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f62(E, r, c):
    # Ô G24
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(3,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(3.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f63(E, r, c):
    # Ô K24
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(3,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(3.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f64(E, r, c):
    # Ô C25
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(4,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(4.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f65(E, r, c):
    # Ô D25
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(4,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(4,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(4.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(4.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f66(E, r, c):
    # Ô E25
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(4,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(4.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f67(E, r, c):
    # Ô G25
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(4,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(4.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f68(E, r, c):
    # Ô K25
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(4,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(4.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f69(E, r, c):
    # Ô C26
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(5,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(5.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f70(E, r, c):
    # Ô D26
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(5,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(5,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(5.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(5.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f71(E, r, c):
    # Ô E26
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(5,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(5.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f72(E, r, c):
    # Ô G26
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(5,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(5.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f73(E, r, c):
    # Ô K26
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(5,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(5.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f74(E, r, c):
    # Ô C27
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(6,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(6.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f75(E, r, c):
    # Ô D27
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(6,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(6,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(6.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(6.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f76(E, r, c):
    # Ô E27
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(6,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(6.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f77(E, r, c):
    # Ô G27
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(6,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(6.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f78(E, r, c):
    # Ô K27
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(6,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(6.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f79(E, r, c):
    # Ô C28
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(7,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(7.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f80(E, r, c):
    # Ô D28
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(7,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(7,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(7.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(7.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f81(E, r, c):
    # Ô E28
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(7,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(7.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f82(E, r, c):
    # Ô G28
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(7,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(7.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f83(E, r, c):
    # Ô K28
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(7,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(7.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f84(E, r, c):
    # Ô C29
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(8,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(8.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f85(E, r, c):
    # Ô D29
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(8,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(8,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(8.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(8.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f86(E, r, c):
    # Ô E29
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(8,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(8.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f87(E, r, c):
    # Ô G29
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(8,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(8.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f88(E, r, c):
    # Ô K29
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(8,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(8.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f89(E, r, c):
    # Ô C30
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(9,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(9.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f90(E, r, c):
    # Ô D30
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(9,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(9,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(9.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(9.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f91(E, r, c):
    # Ô E30
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(9,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(9.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f92(E, r, c):
    # Ô G30
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(9,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(9.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f93(E, r, c):
    # Ô K30
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(9,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(9.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f94(E, r, c):
    # Ô C31
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(10,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(10.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f95(E, r, c):
    # Ô D31
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(10,LUU_NIEN!$BW$30:$BW$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(10,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(10.0, E.rg(15,30,75,129,75), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(10.0, E.rg(15,30,75,129,75), 0.0))), lambda: '')

def f96(E, r, c):
    # Ô E31
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(10,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(10.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f97(E, r, c):
    # Ô G31
    # = IFERROR(INDEX(LUU_NIEN!$BY$30:$BY$129,MATCH(10,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,77,129,77), X.MATCH(10.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f98(E, r, c):
    # Ô K31
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(10,LUU_NIEN!$BW$30:$BW$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(10.0, E.rg(15,30,75,129,75), 0.0)), lambda: '')

def f99(E, r, c):
    # Ô C35
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(1,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(1.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f100(E, r, c):
    # Ô D35
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(1,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(1,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(1.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(1.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f101(E, r, c):
    # Ô E35
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(1,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(1.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f102(E, r, c):
    # Ô G35
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(1,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(1.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f103(E, r, c):
    # Ô C36
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(2,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(2.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f104(E, r, c):
    # Ô D36
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(2,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(2,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(2.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(2.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f105(E, r, c):
    # Ô E36
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(2,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(2.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f106(E, r, c):
    # Ô G36
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(2,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(2.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f107(E, r, c):
    # Ô C37
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(3,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(3.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f108(E, r, c):
    # Ô D37
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(3,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(3,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(3.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(3.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f109(E, r, c):
    # Ô E37
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(3,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(3.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f110(E, r, c):
    # Ô G37
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(3,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(3.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f111(E, r, c):
    # Ô C38
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(4,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(4.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f112(E, r, c):
    # Ô D38
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(4,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(4,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(4.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(4.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f113(E, r, c):
    # Ô E38
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(4,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(4.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f114(E, r, c):
    # Ô G38
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(4,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(4.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f115(E, r, c):
    # Ô C39
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(5,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(5.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f116(E, r, c):
    # Ô D39
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(5,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(5,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(5.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(5.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f117(E, r, c):
    # Ô E39
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(5,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(5.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f118(E, r, c):
    # Ô G39
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(5,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(5.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f119(E, r, c):
    # Ô C40
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(6,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(6.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f120(E, r, c):
    # Ô D40
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(6,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(6,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(6.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(6.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f121(E, r, c):
    # Ô E40
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(6,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(6.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f122(E, r, c):
    # Ô G40
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(6,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(6.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f123(E, r, c):
    # Ô C41
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(7,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(7.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f124(E, r, c):
    # Ô D41
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(7,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(7,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(7.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(7.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f125(E, r, c):
    # Ô E41
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(7,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(7.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f126(E, r, c):
    # Ô G41
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(7,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(7.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f127(E, r, c):
    # Ô C42
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(8,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(8.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f128(E, r, c):
    # Ô D42
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(8,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(8,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(8.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(8.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f129(E, r, c):
    # Ô E42
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(8,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(8.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f130(E, r, c):
    # Ô G42
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(8,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(8.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f131(E, r, c):
    # Ô C43
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(9,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(9.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f132(E, r, c):
    # Ô D43
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(9,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(9,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(9.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(9.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f133(E, r, c):
    # Ô E43
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(9,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(9.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f134(E, r, c):
    # Ô G43
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(9,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(9.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f135(E, r, c):
    # Ô C44
    # = IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(10,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,1,129,1), X.MATCH(10.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f136(E, r, c):
    # Ô D44
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(10,LUU_NIEN!$BX$30:$BX$129,0))&" / ĐV "&INDEX(LUU_NIEN!$H$30:$H$129,MATCH(10,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), X.MATCH(10.0, E.rg(15,30,76,129,76), 0.0)),' / ĐV '),X.INDEX(E.rg(15,30,8,129,8), X.MATCH(10.0, E.rg(15,30,76,129,76), 0.0))), lambda: '')

def f137(E, r, c):
    # Ô E44
    # = IFERROR(INDEX(LUU_NIEN!$C$30:$C$129,MATCH(10,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,3,129,3), X.MATCH(10.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f138(E, r, c):
    # Ô G44
    # = IFERROR(INDEX(LUU_NIEN!$BZ$30:$BZ$129,MATCH(10,LUU_NIEN!$BX$30:$BX$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,78,129,78), X.MATCH(10.0, E.rg(15,30,76,129,76), 0.0)), lambda: '')

def f139(E, r, c):
    # Ô D47
    # = IF(OR(NHAP_LIEU!$C$32="",NHAP_LIEU!$C$33=""),"",INDEX(CHI_TEN,MOD(NHAP_LIEU!$C$32+NHAP_LIEU!$C$33,12)+1))
    return ('' if X.truthy(X.OR(X.eq(E.v(2,32,3),''), X.eq(E.v(2,33,3),''))) else X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(E.v(2,32,3),E.v(2,33,3)), 12.0),1.0)))

def f140(E, r, c):
    # Ô G47
    # = IF(D47="","","Tháng âm "&NHAP_LIEU!$C$32&" → chi "&INDEX(CHI_TEN,MOD(NHAP_LIEU!$C$32+1,12)+1)&" làm mồng 1, đếm thuận đến ngày "&NHAP_LIEU!$C$33&" → d
    #   ừng ở "&D47&" → trục "&D48&". Quy tắc: Tí/Ngọ → Bắc–Nam; Mão/Dậu → Đông–Tây; Sửu, Dần/Mùi, Thân → Đông Bắc–Tây Nam; Thìn, Tị/Tuất, Hợi → Tây Bắc–Đông 
    #   Nam; lấy nguyên quán làm tâm; tác giả ""chưa khẳng định được ở phía nào trên trục"" (P2 tr.494-499).")
    return ('' if X.truthy(X.eq(E.v(17,r,c-3),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tháng âm ',E.v(2,32,3)),' → chi '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(E.v(2,32,3),1.0), 12.0),1.0))),' làm mồng 1, đếm thuận đến ngày '),E.v(2,33,3)),' → dừng ở '),E.v(17,r,c-3)),' → trục '),E.v(17,r+1,c-3)),'. Quy tắc: Tí/Ngọ → Bắc–Nam; Mão/Dậu → Đông–Tây; Sửu, Dần/Mùi, Thân → Đông Bắc–Tây Nam; Thìn, Tị/Tuất, Hợi → Tây Bắc–Đông Nam; lấy nguyên quán làm tâm; tác giả "chưa khẳng định được ở phía nào trên trục" (P2 tr.494-499).'))

def f141(E, r, c):
    # Ô D48
    # = IF(D47="","(nhập tháng, ngày âm lịch ở NHAP_LIEU!C32:C33)",CHOOSE(MOD(NHAP_LIEU!$C$32+NHAP_LIEU!$C$33,12)+1,"Bắc – Nam","Đông Bắc – Tây Nam","Đông Bắc
    #    – Tây Nam","Đông – Tây","Tây Bắc – Đông Nam","Tây Bắc – Đông Nam","Bắc – Nam","Đông Bắc – Tây Nam","Đông Bắc – Tây Nam","Đông – Tây","Tây Bắc – Đông 
    #   Nam","Tây Bắc – Đông Nam"))
    return ('(nhập tháng, ngày âm lịch ở NHAP_LIEU!C32:C33)' if X.truthy(X.eq(E.v(17,r-1,c),'')) else X.CHOOSE(X.add(X.MOD(X.add(E.v(2,32,3),E.v(2,33,3)), 12.0),1.0), lambda: 'Bắc – Nam', lambda: 'Đông Bắc – Tây Nam', lambda: 'Đông Bắc – Tây Nam', lambda: 'Đông – Tây', lambda: 'Tây Bắc – Đông Nam', lambda: 'Tây Bắc – Đông Nam', lambda: 'Bắc – Nam', lambda: 'Đông Bắc – Tây Nam', lambda: 'Đông Bắc – Tây Nam', lambda: 'Đông – Tây', lambda: 'Tây Bắc – Đông Nam', lambda: 'Tây Bắc – Đông Nam'))

def f142(E, r, c):
    # Ô D51
    # = INDEX(HG_TEN,LUC_THAN!$D$5+1)&" – "&LUC_THAN!$E$5
    return X.cat(X.cat(X.INDEX(E.rg(24,140,2,199,2), X.add(E.v(16,5,4),1.0)),' – '),E.v(16,5,5))

def f143(E, r, c):
    # Ô D52
    # = IF(R52="","",IF(AND(ISNUMBER(NHAP_LIEU!$C$34),NHAP_LIEU!$C$34<=31,ISNUMBER(NHAP_LIEU!$D$34),ISNUMBER(NHAP_LIEU!$E$34)),IF(P52="","",MOD(P52-1984,60)),
    #   MOD(R52-1984,60)))
    return ('' if X.truthy(X.eq(E.v(17,r,c+14),'')) else (('' if X.truthy(X.eq(E.v(17,r,c+12),'')) else X.MOD(X.sub(E.v(17,r,c+12),1984.0), 60.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.le(E.v(2,34,3),31.0), X.ISNUMBER(lambda: E.v(2,34,4)), X.ISNUMBER(lambda: E.v(2,34,5)))) else X.MOD(X.sub(E.v(17,r,c+14),1984.0), 60.0)))

def f144(E, r, c):
    # Ô F52
    # = IF(R52="","",IF(AND(ISNUMBER(NHAP_LIEU!$C$34),NHAP_LIEU!$C$34<=31,ISNUMBER(NHAP_LIEU!$D$34),ISNUMBER(NHAP_LIEU!$E$34)),IF(M52="LỖI","LỖI: ngày sinh ng
    #   ười kia không hợp lệ (NHAP_LIEU C34:E34)",IF(P52="","Ngày sinh ngoài bảng tiết 1899–2101","Sinh "&TEXT(M52,"dd/mm/yyyy")&IF(ISNUMBER(NHAP_LIEU!$F$34),
    #   " "&TEXT(M52,"hh:mm"),"")&"; Lập xuân năm "&R52&": "&IF(Q52="","?",TEXT(Q52,"dd/mm/yyyy hh:mm"))&" → "&IF(P52<R52,"sinh TRƯỚC Lập xuân nên trụ năm thu
    #   ộc năm "&P52&" ("&INDEX(HG_TEN,MOD(P52-1984,60)+1)&"), không phải năm "&R52,"sinh sau Lập xuân → trụ năm "&INDEX(HG_TEN,MOD(P52-1984,60)+1))&" (P2 tr.
    #   268)."&IF(AND(NOT(ISNUMBER(NHAP_LIEU!$F$34)),Q52<>"",INT(Q52)=INT(M52))," CẢNH BÁO: sinh đúng ngày Lập xuân mà chưa nhập giờ – đang tạm lấy 12:00, hãy
    #    nhập giờ sinh.",""))),"Chỉ nhập năm "&R52&": công cụ coi như sinh sau Lập xuân ("&IF(Q52="","khoảng 4/2",TEXT(Q52,"dd/mm/yyyy hh:mm"))&") → "&INDEX(H
    #   G_TEN,MOD(R52-1984,60)+1)&". Nếu người kia sinh từ 1/1 đến trước thời điểm đó thì trụ năm là "&INDEX(HG_TEN,MOD(R52-1985,60)+1)&" – hãy nhập đủ ngày |
    #    tháng | năm ở NHAP_LIEU C34:E34."))
    return ('' if X.truthy(X.eq(E.v(17,r,c+12),'')) else (('LỖI: ngày sinh người kia không hợp lệ (NHAP_LIEU C34:E34)' if X.truthy(X.eq(E.v(17,r,c+7),'LỖI')) else ('Ngày sinh ngoài bảng tiết 1899–2101' if X.truthy(X.eq(E.v(17,r,c+10),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Sinh ',X.TEXT(E.v(17,r,c+7), 'dd/mm/yyyy')),(X.cat(' ',X.TEXT(E.v(17,r,c+7), 'hh:mm')) if X.truthy(X.ISNUMBER(lambda: E.v(2,34,6))) else '')),'; Lập xuân năm '),E.v(17,r,c+12)),': '),('?' if X.truthy(X.eq(E.v(17,r,c+11),'')) else X.TEXT(E.v(17,r,c+11), 'dd/mm/yyyy hh:mm'))),' → '),(X.cat(X.cat(X.cat(X.cat(X.cat('sinh TRƯỚC Lập xuân nên trụ năm thuộc năm ',E.v(17,r,c+10)),' ('),X.INDEX(E.rg(24,140,2,199,2), X.add(X.MOD(X.sub(E.v(17,r,c+10),1984.0), 60.0),1.0))),'), không phải năm '),E.v(17,r,c+12)) if X.truthy(X.lt(E.v(17,r,c+10),E.v(17,r,c+12))) else X.cat('sinh sau Lập xuân → trụ năm ',X.INDEX(E.rg(24,140,2,199,2), X.add(X.MOD(X.sub(E.v(17,r,c+10),1984.0), 60.0),1.0))))),' (P2 tr.268).'),(' CẢNH BÁO: sinh đúng ngày Lập xuân mà chưa nhập giờ – đang tạm lấy 12:00, hãy nhập giờ sinh.' if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: E.v(2,34,6))), X.ne(E.v(17,r,c+11),''), X.eq(X.INT(E.v(17,r,c+11)),X.INT(E.v(17,r,c+7))))) else '')))) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.le(E.v(2,34,3),31.0), X.ISNUMBER(lambda: E.v(2,34,4)), X.ISNUMBER(lambda: E.v(2,34,5)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chỉ nhập năm ',E.v(17,r,c+12)),': công cụ coi như sinh sau Lập xuân ('),('khoảng 4/2' if X.truthy(X.eq(E.v(17,r,c+11),'')) else X.TEXT(E.v(17,r,c+11), 'dd/mm/yyyy hh:mm'))),') → '),X.INDEX(E.rg(24,140,2,199,2), X.add(X.MOD(X.sub(E.v(17,r,c+12),1984.0), 60.0),1.0))),'. Nếu người kia sinh từ 1/1 đến trước thời điểm đó thì trụ năm là '),X.INDEX(E.rg(24,140,2,199,2), X.add(X.MOD(X.sub(E.v(17,r,c+12),1985.0), 60.0),1.0))),' – hãy nhập đủ ngày | tháng | năm ở NHAP_LIEU C34:E34.')))

def f145(E, r, c):
    # Ô M52
    # = IF(AND(ISNUMBER(NHAP_LIEU!$C$34),NHAP_LIEU!$C$34<=31,ISNUMBER(NHAP_LIEU!$D$34),ISNUMBER(NHAP_LIEU!$E$34)),IFERROR(IF(DAY(DATE(NHAP_LIEU!$E$34,NHAP_LIE
    #   U!$D$34,NHAP_LIEU!$C$34))<>NHAP_LIEU!$C$34,"LỖI",DATE(NHAP_LIEU!$E$34,NHAP_LIEU!$D$34,NHAP_LIEU!$C$34)+IF(ISNUMBER(NHAP_LIEU!$F$34),NHAP_LIEU!$F$34,12
    #   )/24+IF(ISNUMBER(NHAP_LIEU!$G$34),NHAP_LIEU!$G$34,0)/1440),"LỖI"),"")
    return (X.IFERROR(lambda: ('LỖI' if X.truthy(X.ne(X.DAY(X.DATE(E.v(2,34,5), E.v(2,34,4), E.v(2,34,3))),E.v(2,34,3))) else X.add(X.add(X.DATE(E.v(2,34,5), E.v(2,34,4), E.v(2,34,3)),X.div((E.v(2,34,6) if X.truthy(X.ISNUMBER(lambda: E.v(2,34,6))) else 12.0),24.0)),X.div((E.v(2,34,7) if X.truthy(X.ISNUMBER(lambda: E.v(2,34,7))) else 0.0),1440.0))), lambda: 'LỖI') if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.le(E.v(2,34,3),31.0), X.ISNUMBER(lambda: E.v(2,34,4)), X.ISNUMBER(lambda: E.v(2,34,5)))) else '')

def f146(E, r, c):
    # Ô N52
    # = IF(OR(M52="",M52="LỖI"),"",M52-DAI_VAN!$C$9/24)
    return ('' if X.truthy(X.OR(X.eq(E.v(17,r,c-1),''), X.eq(E.v(17,r,c-1),'LỖI'))) else X.sub(E.v(17,r,c-1),X.div(E.v(14,9,3),24.0)))

def f147(E, r, c):
    # Ô O52
    # = IF(N52="","",IFERROR(MATCH(N52,TIET_UTC,1),""))
    return ('' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.IFERROR(lambda: X.MATCH(E.v(17,r,c-1), E.rg(25,4,1,2439,1), 1.0), lambda: ''))

def f148(E, r, c):
    # Ô P52
    # = IF(O52="","",INDEX(TIET_NAM,O52))
    return ('' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.INDEX(E.rg(25,4,5,2439,5), E.v(17,r,c-1)))

def f149(E, r, c):
    # Ô Q52
    # = IF(IF(AND(ISNUMBER(NHAP_LIEU!$C$34),NHAP_LIEU!$C$34<=31,ISNUMBER(NHAP_LIEU!$D$34),ISNUMBER(NHAP_LIEU!$E$34)),NHAP_LIEU!$E$34,IF(AND(ISNUMBER(NHAP_LIEU
    #   !$C$34),NHAP_LIEU!$C$34>=1900,NOT(ISNUMBER(NHAP_LIEU!$E$34))),NHAP_LIEU!$C$34,""))="","",IFERROR(INDEX(TIET_UTC,MATCH(DATE(IF(AND(ISNUMBER(NHAP_LIEU!$
    #   C$34),NHAP_LIEU!$C$34<=31,ISNUMBER(NHAP_LIEU!$D$34),ISNUMBER(NHAP_LIEU!$E$34)),NHAP_LIEU!$E$34,IF(AND(ISNUMBER(NHAP_LIEU!$C$34),NHAP_LIEU!$C$34>=1900,
    #   NOT(ISNUMBER(NHAP_LIEU!$E$34))),NHAP_LIEU!$C$34,"")),2,15),TIET_UTC,1))+DAI_VAN!$C$9/24,""))
    return ('' if X.truthy(X.eq((E.v(2,34,5) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.le(E.v(2,34,3),31.0), X.ISNUMBER(lambda: E.v(2,34,4)), X.ISNUMBER(lambda: E.v(2,34,5)))) else (E.v(2,34,3) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.ge(E.v(2,34,3),1900.0), X.NOT(X.ISNUMBER(lambda: E.v(2,34,5))))) else '')),'')) else X.IFERROR(lambda: X.add(X.INDEX(E.rg(25,4,1,2439,1), X.MATCH(X.DATE((E.v(2,34,5) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.le(E.v(2,34,3),31.0), X.ISNUMBER(lambda: E.v(2,34,4)), X.ISNUMBER(lambda: E.v(2,34,5)))) else (E.v(2,34,3) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.ge(E.v(2,34,3),1900.0), X.NOT(X.ISNUMBER(lambda: E.v(2,34,5))))) else '')), 2.0, 15.0), E.rg(25,4,1,2439,1), 1.0)),X.div(E.v(14,9,3),24.0)), lambda: ''))

def f150(E, r, c):
    # Ô R52
    # = IF(AND(ISNUMBER(NHAP_LIEU!$C$34),NHAP_LIEU!$C$34<=31,ISNUMBER(NHAP_LIEU!$D$34),ISNUMBER(NHAP_LIEU!$E$34)),NHAP_LIEU!$E$34,IF(AND(ISNUMBER(NHAP_LIEU!$C
    #   $34),NHAP_LIEU!$C$34>=1900,NOT(ISNUMBER(NHAP_LIEU!$E$34))),NHAP_LIEU!$C$34,""))
    return (E.v(2,34,5) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.le(E.v(2,34,3),31.0), X.ISNUMBER(lambda: E.v(2,34,4)), X.ISNUMBER(lambda: E.v(2,34,5)))) else (E.v(2,34,3) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(2,34,3)), X.ge(E.v(2,34,3),1900.0), X.NOT(X.ISNUMBER(lambda: E.v(2,34,5))))) else ''))

def f151(E, r, c):
    # Ô D53
    # = IF(D52="","",INDEX(HG_TEN,D52+1)&" – "&INDEX(NAPAM_TEN,D52+1)&" ("&INDEX(HANH_TEN,INDEX(NAPAM_H,D52+1)+1)&")")
    return ('' if X.truthy(X.eq(E.v(17,r-1,c),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,140,2,199,2), X.add(E.v(17,r-1,c),1.0)),' – '),X.INDEX(E.rg(24,140,3,199,3), X.add(E.v(17,r-1,c),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(17,r-1,c),1.0)),1.0))),')'))

def f152(E, r, c):
    # Ô D54
    # = IF(D52="","",IF(INDEX(M_NH,TRUNG_GIAN!$C$5+1,MOD(D52,10)+1)>=0,"có","không"))
    return ('' if X.truthy(X.eq(E.v(17,r-2,c),'')) else ('có' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,3),1.0), X.add(X.MOD(E.v(17,r-2,c), 10.0),1.0)),0.0)) else 'không'))

def f153(E, r, c):
    # Ô D55
    # = IF(D52="","",IF(INDEX(M_LH,TRUNG_GIAN!$C$6+1,MOD(D52,12)+1)>=0,"có","không"))
    return ('' if X.truthy(X.eq(E.v(17,r-3,c),'')) else ('có' if X.truthy(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(X.MOD(E.v(17,r-3,c), 12.0),1.0)),0.0)) else 'không'))

def f154(E, r, c):
    # Ô D56
    # = IF(D52="","",IF(OR(MOD(INDEX(NAPAM_H,LUC_THAN!$D$5+1)+1,5)=INDEX(NAPAM_H,D52+1),MOD(INDEX(NAPAM_H,D52+1)+1,5)=INDEX(NAPAM_H,LUC_THAN!$D$5+1)),"tương s
    #   inh",IF(OR(MOD(INDEX(NAPAM_H,LUC_THAN!$D$5+1)+2,5)=INDEX(NAPAM_H,D52+1),MOD(INDEX(NAPAM_H,D52+1)+2,5)=INDEX(NAPAM_H,LUC_THAN!$D$5+1)),"tương khắc","cù
    #   ng hành")))
    return ('' if X.truthy(X.eq(E.v(17,r-4,c),'')) else ('tương sinh' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(16,5,4),1.0)),1.0), 5.0),X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(17,r-4,c),1.0))), X.eq(X.MOD(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(17,r-4,c),1.0)),1.0), 5.0),X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(16,5,4),1.0))))) else ('tương khắc' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(16,5,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(17,r-4,c),1.0))), X.eq(X.MOD(X.add(X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(17,r-4,c),1.0)),2.0), 5.0),X.INDEX(E.rg(24,140,4,199,4), X.add(E.v(16,5,4),1.0))))) else 'cùng hành')))

def f155(E, r, c):
    # Ô D57
    # = IF(D52="","",IF(AND(OR(MOD(INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$C$5+1)+2,5)=INDEX(CSDL!$C$5:$C$14,MOD(D52,10)+1),MOD(INDEX(CSDL!$C$5:$C$14,MOD(D52,10)+1)
    #   +2,5)=INDEX(CSDL!$C$5:$C$14,TRUNG_GIAN!$C$5+1)),INDEX(M_XU,TRUNG_GIAN!$C$6+1,MOD(D52,12)+1)=1),"CÓ","không"))
    return ('' if X.truthy(X.eq(E.v(17,r-5,c),'')) else ('CÓ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,5,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,5,3,14,3), X.add(X.MOD(E.v(17,r-5,c), 10.0),1.0))), X.eq(X.MOD(X.add(X.INDEX(E.rg(24,5,3,14,3), X.add(X.MOD(E.v(17,r-5,c), 10.0),1.0)),2.0), 5.0),X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(23,5,3),1.0)))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(X.MOD(E.v(17,r-5,c), 12.0),1.0)),1.0))) else 'không'))

def f156(E, r, c):
    # Ô D58
    # = IF(D52="","",TRIM(IF(INDEX(M_XU,TRUNG_GIAN!$C$6+1,MOD(D52,12)+1)=1,"xung ","")&IF(INDEX(M_HAI,TRUNG_GIAN!$C$6+1,MOD(D52,12)+1)=1,"hại ","")&IF(INDEX(M
    #   _HINH,TRUNG_GIAN!$C$6+1,MOD(D52,12)+1)=1,"hình ",""))&"")
    return ('' if X.truthy(X.eq(E.v(17,r-6,c),'')) else X.cat(X.TRIM(X.cat(X.cat(('xung ' if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(X.MOD(E.v(17,r-6,c), 12.0),1.0)),1.0)) else ''),('hại ' if X.truthy(X.eq(X.INDEX(E.rg(24,77,16,88,27), X.add(E.v(23,6,3),1.0), X.add(X.MOD(E.v(17,r-6,c), 12.0),1.0)),1.0)) else '')),('hình ' if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,3),1.0), X.add(X.MOD(E.v(17,r-6,c), 12.0),1.0)),1.0)) else ''))),''))

def f157(E, r, c):
    # Ô D59
    # = IF(D52="","(chưa nhập ngày/năm sinh người kia)",TRIM(IF(D57="CÓ","• Trụ năm thiên khắc địa xung: ""trong tất cả hôn nhân bất hạnh, mệnh năm thiên khắc
    #    địa xung là nặng nhất""; khó qua cửa ải tuổi 54 (P1 tr.175-176; P2 tr.489-490). ","")&IF(AND(D54="có",D55="có",D56="tương sinh"),"• Thiên hợp địa hợp
    #   , nạp âm tương sinh: hôn nhân tốt, bách niên giai lão (P1 tr.174; P2 tr.467). ","")&IF(AND(D54="có",D55="có",D56="tương khắc"),"• Thiên hợp địa hợp mà
    #    nạp âm khắc: hay cãi nhau nhưng không tan (P1 tr.175). ","")&IF(AND(OR(D54<>"có",D55<>"có"),D56="tương sinh"),"• Nạp âm năm tương sinh: gắn bó (P1 tr
    #   .176-177; P2 tr.488). ","")&IF(AND(OR(D54="có",D55="có"),NOT(AND(D54="có",D55="có"))),"• "&IF(D54="có","Can","Chi")&" năm tương hợp: thuận cho hôn nhâ
    #   n (P1 tr.174). ","")&IF(AND(D57<>"CÓ",OR(ISNUMBER(SEARCH("xung",D58)),ISNUMBER(SEARCH("hại",D58)))),"• 12 con giáp xung/hại: không nên (P2 tr.488). ",
    #   "")&IF(AND(D57<>"CÓ",D56="tương khắc",OR(D54<>"có",D55<>"có")),"• Nạp âm năm tương khắc: kém (P1 tr.176). ","")))
    return ('(chưa nhập ngày/năm sinh người kia)' if X.truthy(X.eq(E.v(17,r-7,c),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('• Trụ năm thiên khắc địa xung: "trong tất cả hôn nhân bất hạnh, mệnh năm thiên khắc địa xung là nặng nhất"; khó qua cửa ải tuổi 54 (P1 tr.175-176; P2 tr.489-490). ' if X.truthy(X.eq(E.v(17,r-2,c),'CÓ')) else ''),('• Thiên hợp địa hợp, nạp âm tương sinh: hôn nhân tốt, bách niên giai lão (P1 tr.174; P2 tr.467). ' if X.truthy(X.AND(X.eq(E.v(17,r-5,c),'có'), X.eq(E.v(17,r-4,c),'có'), X.eq(E.v(17,r-3,c),'tương sinh'))) else '')),('• Thiên hợp địa hợp mà nạp âm khắc: hay cãi nhau nhưng không tan (P1 tr.175). ' if X.truthy(X.AND(X.eq(E.v(17,r-5,c),'có'), X.eq(E.v(17,r-4,c),'có'), X.eq(E.v(17,r-3,c),'tương khắc'))) else '')),('• Nạp âm năm tương sinh: gắn bó (P1 tr.176-177; P2 tr.488). ' if X.truthy(X.AND(X.OR(X.ne(E.v(17,r-5,c),'có'), X.ne(E.v(17,r-4,c),'có')), X.eq(E.v(17,r-3,c),'tương sinh'))) else '')),(X.cat(X.cat('• ',('Can' if X.truthy(X.eq(E.v(17,r-5,c),'có')) else 'Chi')),' năm tương hợp: thuận cho hôn nhân (P1 tr.174). ') if X.truthy(X.AND(X.OR(X.eq(E.v(17,r-5,c),'có'), X.eq(E.v(17,r-4,c),'có')), X.NOT(X.AND(X.eq(E.v(17,r-5,c),'có'), X.eq(E.v(17,r-4,c),'có'))))) else '')),('• 12 con giáp xung/hại: không nên (P2 tr.488). ' if X.truthy(X.AND(X.ne(E.v(17,r-2,c),'CÓ'), X.OR(X.ISNUMBER(lambda: X.SEARCH('xung', E.v(17,r-1,c))), X.ISNUMBER(lambda: X.SEARCH('hại', E.v(17,r-1,c)))))) else '')),('• Nạp âm năm tương khắc: kém (P1 tr.176). ' if X.truthy(X.AND(X.ne(E.v(17,r-2,c),'CÓ'), X.eq(E.v(17,r-3,c),'tương khắc'), X.OR(X.ne(E.v(17,r-5,c),'có'), X.ne(E.v(17,r-4,c),'có')))) else ''))))

def f158(E, r, c):
    # Ô D64
    # = IF(AG64>0,SUBSTITUTE(SUBSTITUTE(AH64,"hôn nhân tương đối ổn, có lúc trục trặc nhỏ","bản thân sao phối ngẫu có lực; còn sự êm ấm của hôn nhân có "&AG64
    #   &" dấu hiệu bất lợi của sách (xem dòng dưới)"),"; hôn nhân bền","; riêng hôn nhân có "&AG64&" dấu hiệu bất lợi của sách (xem dòng dưới)"),AH64)
    return (X.SUBSTITUTE(X.SUBSTITUTE(E.v(17,r,c+30), 'hôn nhân tương đối ổn, có lúc trục trặc nhỏ', X.cat(X.cat('bản thân sao phối ngẫu có lực; còn sự êm ấm của hôn nhân có ',E.v(17,r,c+29)),' dấu hiệu bất lợi của sách (xem dòng dưới)')), '; hôn nhân bền', X.cat(X.cat('; riêng hôn nhân có ',E.v(17,r,c+29)),' dấu hiệu bất lợi của sách (xem dòng dưới)')) if X.truthy(X.gt(E.v(17,r,c+29),0.0)) else E.v(17,r,c+30))

def f159(E, r, c):
    # Ô G64
    # = "Sao "&IF(NHAP_LIEU!$C$22="Nam","Tài (vợ)","Quan sát (chồng)")&": "&IF(M64=0,"KHÔNG CÓ trong tứ trụ (−2)",M64&" chữ ("&AA64&"); được ngôi ở trụ ngày: 
    #   "&IF(N64=1,"có (+1)","không")&"; vượng tại lệnh: "&IF(O64=1,"có (+1)","không")&"; thấu can có lực: "&IF(P64=1,"có (+1)","không")&"; vai trò PP1 "&R64&
    #   ", PP2 "&S64&" (dụng/hỷ +1, kỵ −1)"&"; còn nguyên vẹn (không bị xung, khắc, hợp mất, tuần không): "&Q64&"/"&M64&IF(Q64=0," (−1)",""))&"; cung ngày ("&
    #   TRUNG_GIAN!$E$11&") bị xung/hình: "&IF(T64=1,"có (−1)","không")&" → điểm PP1 "&U64&", PP2 "&V64&" (P1 tr.159)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Sao ',('Tài (vợ)' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else 'Quan sát (chồng)')),': '),('KHÔNG CÓ trong tứ trụ (−2)' if X.truthy(X.eq(E.v(17,r,c+6),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(17,r,c+6),' chữ ('),E.v(17,r,c+20)),'); được ngôi ở trụ ngày: '),('có (+1)' if X.truthy(X.eq(E.v(17,r,c+7),1.0)) else 'không')),'; vượng tại lệnh: '),('có (+1)' if X.truthy(X.eq(E.v(17,r,c+8),1.0)) else 'không')),'; thấu can có lực: '),('có (+1)' if X.truthy(X.eq(E.v(17,r,c+9),1.0)) else 'không')),'; vai trò PP1 '),E.v(17,r,c+11)),', PP2 '),E.v(17,r,c+12)),' (dụng/hỷ +1, kỵ −1)'),'; còn nguyên vẹn (không bị xung, khắc, hợp mất, tuần không): '),E.v(17,r,c+10)),'/'),E.v(17,r,c+6)),(' (−1)' if X.truthy(X.eq(E.v(17,r,c+10),0.0)) else '')))),'; cung ngày ('),E.v(23,11,5)),') bị xung/hình: '),('có (−1)' if X.truthy(X.eq(E.v(17,r,c+13),1.0)) else 'không')),' → điểm PP1 '),E.v(17,r,c+14)),', PP2 '),E.v(17,r,c+15)),' (P1 tr.159).')

def f160(E, r, c):
    # Ô H64
    # = TRIM("Thuận lợi kết hôn: "&TRIM(IF(HON_NHAN!$C$22="","",HON_NHAN!$C$22&" ")&IF(HON_NHAN!$C$23="","",HON_NHAN!$C$23&" ")&IF(HON_NHAN!$C$24="","",HON_NH
    #   AN!$C$24&" ")&IF(HON_NHAN!$C$25="","",HON_NHAN!$C$25&" ")&IF(HON_NHAN!$C$26="","",HON_NHAN!$C$26&" ")&IF(HON_NHAN!$C$27="","",HON_NHAN!$C$27&" "))&CHA
    #   R(10)&"Cần thận trọng: "&TRIM(IF(HON_NHAN!$C$35="","",HON_NHAN!$C$35&" ")&IF(HON_NHAN!$C$36="","",HON_NHAN!$C$36&" ")&IF(HON_NHAN!$C$37="","",HON_NHAN
    #   !$C$37&" ")&IF(HON_NHAN!$C$38="","",HON_NHAN!$C$38&" ")&IF(HON_NHAN!$C$39="","",HON_NHAN!$C$39&" ")&IF(HON_NHAN!$C$40="","",HON_NHAN!$C$40&" ")))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat('Thuận lợi kết hôn: ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(17,22,3),'')) else X.cat(E.v(17,22,3),' ')),('' if X.truthy(X.eq(E.v(17,23,3),'')) else X.cat(E.v(17,23,3),' '))),('' if X.truthy(X.eq(E.v(17,24,3),'')) else X.cat(E.v(17,24,3),' '))),('' if X.truthy(X.eq(E.v(17,25,3),'')) else X.cat(E.v(17,25,3),' '))),('' if X.truthy(X.eq(E.v(17,26,3),'')) else X.cat(E.v(17,26,3),' '))),('' if X.truthy(X.eq(E.v(17,27,3),'')) else X.cat(E.v(17,27,3),' '))))),X.CHAR(10.0)),'Cần thận trọng: '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(17,35,3),'')) else X.cat(E.v(17,35,3),' ')),('' if X.truthy(X.eq(E.v(17,36,3),'')) else X.cat(E.v(17,36,3),' '))),('' if X.truthy(X.eq(E.v(17,37,3),'')) else X.cat(E.v(17,37,3),' '))),('' if X.truthy(X.eq(E.v(17,38,3),'')) else X.cat(E.v(17,38,3),' '))),('' if X.truthy(X.eq(E.v(17,39,3),'')) else X.cat(E.v(17,39,3),' '))),('' if X.truthy(X.eq(E.v(17,40,3),'')) else X.cat(E.v(17,40,3),' '))))))

def f161(E, r, c):
    # Ô K64
    # = IF(OR(X64="Tốt",X64="Khá"),"Sách nêu – khớp kết luận: "&IF(NHAP_LIEU!$C$22="Nữ","(lá số NỮ: đọc phần sách ghi ‘nữ’; phần ghi ‘nam’ không áp dụng) ","(
    #   lá số NAM: đọc phần sách ghi ‘nam’; phần ghi ‘nữ’ không áp dụng) ")&IF(AB64="","(không có dòng riêng)",AB64)&IF(AND(AC64<>"",Z64<>""),CHAR(10)&"Chỉ cầ
    #   n đề phòng khi gặp đại vận kỵ thần ("&IF(Z64="","không có",Z64)&"): "&AC64,""),IF(OR(X64="Kém",X64="Rất kém"),"Sách nêu – khớp kết luận: "&IF(NHAP_LIE
    #   U!$C$22="Nữ","(lá số NỮ: đọc phần sách ghi ‘nữ’; phần ghi ‘nam’ không áp dụng) ","(lá số NAM: đọc phần sách ghi ‘nam’; phần ghi ‘nữ’ không áp dụng) ")
    #   &IF(AC64="","(không có dòng riêng)",AC64)&IF(AND(AB64<>"",Y64<>""),CHAR(10)&"Chỉ cải thiện được khi gặp đại vận dụng/hỷ thần ("&IF(Y64="","không có",Y
    #   64)&"): "&AB64,""),"Không chiều nào trội, các ý của sách ứng theo vận:"&CHAR(10)&"– Đại vận dụng/hỷ thần ("&IF(Y64="","không có",Y64)&"): "&IF(OR(AB64
    #   ="",Y64=""),"—",AB64)&CHAR(10)&"– Đại vận kỵ thần ("&IF(Z64="","không có",Z64)&"): "&IF(OR(AC64="",Z64=""),"—",AC64)))&IF(AD64<>"",CHAR(10)&"Đặc điểm 
    #   khác: "&AD64,"")
    return X.cat((X.cat(X.cat(X.cat('Sách nêu – khớp kết luận: ',('(lá số NỮ: đọc phần sách ghi ‘nữ’; phần ghi ‘nam’ không áp dụng) ' if X.truthy(X.eq(E.v(2,22,3),'Nữ')) else '(lá số NAM: đọc phần sách ghi ‘nam’; phần ghi ‘nữ’ không áp dụng) ')),('(không có dòng riêng)' if X.truthy(X.eq(E.v(17,r,c+17),'')) else E.v(17,r,c+17))),(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'Chỉ cần đề phòng khi gặp đại vận kỵ thần ('),('không có' if X.truthy(X.eq(E.v(17,r,c+15),'')) else E.v(17,r,c+15))),'): '),E.v(17,r,c+18)) if X.truthy(X.AND(X.ne(E.v(17,r,c+18),''), X.ne(E.v(17,r,c+15),''))) else '')) if X.truthy(X.OR(X.eq(E.v(17,r,c+13),'Tốt'), X.eq(E.v(17,r,c+13),'Khá'))) else (X.cat(X.cat(X.cat('Sách nêu – khớp kết luận: ',('(lá số NỮ: đọc phần sách ghi ‘nữ’; phần ghi ‘nam’ không áp dụng) ' if X.truthy(X.eq(E.v(2,22,3),'Nữ')) else '(lá số NAM: đọc phần sách ghi ‘nam’; phần ghi ‘nữ’ không áp dụng) ')),('(không có dòng riêng)' if X.truthy(X.eq(E.v(17,r,c+18),'')) else E.v(17,r,c+18))),(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'Chỉ cải thiện được khi gặp đại vận dụng/hỷ thần ('),('không có' if X.truthy(X.eq(E.v(17,r,c+14),'')) else E.v(17,r,c+14))),'): '),E.v(17,r,c+17)) if X.truthy(X.AND(X.ne(E.v(17,r,c+17),''), X.ne(E.v(17,r,c+14),''))) else '')) if X.truthy(X.OR(X.eq(E.v(17,r,c+13),'Kém'), X.eq(E.v(17,r,c+13),'Rất kém'))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Không chiều nào trội, các ý của sách ứng theo vận:',X.CHAR(10.0)),'– Đại vận dụng/hỷ thần ('),('không có' if X.truthy(X.eq(E.v(17,r,c+14),'')) else E.v(17,r,c+14))),'): '),('—' if X.truthy(X.OR(X.eq(E.v(17,r,c+17),''), X.eq(E.v(17,r,c+14),''))) else E.v(17,r,c+17))),X.CHAR(10.0)),'– Đại vận kỵ thần ('),('không có' if X.truthy(X.eq(E.v(17,r,c+15),'')) else E.v(17,r,c+15))),'): '),('—' if X.truthy(X.OR(X.eq(E.v(17,r,c+18),''), X.eq(E.v(17,r,c+15),''))) else E.v(17,r,c+18))))),(X.cat(X.cat(X.CHAR(10.0),'Đặc điểm khác: '),E.v(17,r,c+19)) if X.truthy(X.ne(E.v(17,r,c+19),'')) else ''))

def f162(E, r, c):
    # Ô M64
    # = IF(NHAP_LIEU!$C$22="Nam",(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$X$5:$X$20,1,
    #   SAO!$Y$5:$Y$20,0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,1,SAO!$Y
    #   $5:$Y$20,0)))
    return (X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0)) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0)))

def f163(E, r, c):
    # Ô N64
    # = IF(IF(NHAP_LIEU!$C$22="Nam",(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tà
    #   i",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$B$5:$B$20,2)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$B$5:$B$20,2)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$B$5:$B$20,2)))>0,1,0)
    return (1.0 if X.truthy(X.gt((X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,2,20,2), 2.0)) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,2,20,2), 2.0))),0.0)) else 0.0)

def f164(E, r, c):
    # Ô O64
    # = IF(IF(NHAP_LIEU!$C$22="Nam",(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tà
    #   i",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$J$5:$J$20,1)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$J$5:$J$20,1)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$J$5:$J$20,1)))>0,1,0)
    return (1.0 if X.truthy(X.gt((X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,10,20,10), 1.0)) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,10,20,10), 1.0))),0.0)) else 0.0)

def f165(E, r, c):
    # Ô P64
    # = IF(IF(NHAP_LIEU!$C$22="Nam",(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0)+COUNTIFS(SAO
    #   !$F$5:$F$20,"Thiên tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0)),(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$X$5:$X$20
    #   ,1,SAO!$Y$5:$Y$20,0,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$C$5:$C$20,"Can",S
    #   AO!$L$5:$L$20,0)))>0,1,0)
    return (1.0 if X.truthy(X.gt((X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0))),0.0)) else 0.0)

def f166(E, r, c):
    # Ô Q64
    # = IF(NHAP_LIEU!$C$22="Nam",(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$M$5:$M$20,0,SAO!$L$5:$L$20,0,SAO!$N$5:$N$20,0)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$M$5:$M$20,0,SAO!$L$5:$L$20,0,SAO!$N$5:$N$20,0)),(COUNTIFS(SAO!$F$5:$F$20,"Chí
    #   nh quan",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$M$5:$M$20,0,SAO!$L$5:$L$20,0,SAO!$N$5:$N$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$X$5:$X$20,1,SAO
    #   !$Y$5:$Y$20,0,SAO!$M$5:$M$20,0,SAO!$L$5:$L$20,0,SAO!$N$5:$N$20,0)))
    return (X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,14,20,14), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,14,20,14), 0.0)) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,14,20,14), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,14,20,14), 0.0)))

def f167(E, r, c):
    # Ô R64
    # = IF(NHAP_LIEU!$C$22="Nam",MENH_CUC!$H$8,MENH_CUC!$H$9)
    return (E.v(18,8,8) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else E.v(18,9,8))

def f168(E, r, c):
    # Ô S64
    # = IF(NHAP_LIEU!$C$22="Nam",MENH_CUC!$I$8,MENH_CUC!$I$9)
    return (E.v(18,8,9) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else E.v(18,9,9))

def f169(E, r, c):
    # Ô T64
    # = IF(OR(SAO!$M$14=1,OR(INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1,INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1,INDEX(M_HINH,TRUNG_GIAN!$
    #   E$6+1,TRUNG_GIAN!$F$6+1)=1)),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(13,14,13),1.0), X.OR(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)))) else 0.0)

def f170(E, r, c):
    # Ô U64
    # = IF(M64=0,-2,N64+O64+P64-IF(Q64=0,1,0)+IF(OR(LEFT(R64,4)="Dụng",LEFT(R64,2)="Hỷ"),1,IF(LEFT(R64,2)="Kỵ",-1,0)))-T64
    return X.sub((X.neg(2.0) if X.truthy(X.eq(E.v(17,r,c-8),0.0)) else X.add(X.sub(X.add(X.add(E.v(17,r,c-7),E.v(17,r,c-6)),E.v(17,r,c-5)),(1.0 if X.truthy(X.eq(E.v(17,r,c-4),0.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.LEFT(E.v(17,r,c-3), 4.0),'Dụng'), X.eq(X.LEFT(E.v(17,r,c-3), 2.0),'Hỷ'))) else (X.neg(1.0) if X.truthy(X.eq(X.LEFT(E.v(17,r,c-3), 2.0),'Kỵ')) else 0.0)))),E.v(17,r,c-1))

def f171(E, r, c):
    # Ô V64
    # = IF(M64=0,-2,N64+O64+P64-IF(Q64=0,1,0)+IF(OR(LEFT(S64,4)="Dụng",LEFT(S64,2)="Hỷ"),1,IF(LEFT(S64,2)="Kỵ",-1,0)))-T64
    return X.sub((X.neg(2.0) if X.truthy(X.eq(E.v(17,r,c-9),0.0)) else X.add(X.sub(X.add(X.add(E.v(17,r,c-8),E.v(17,r,c-7)),E.v(17,r,c-6)),(1.0 if X.truthy(X.eq(E.v(17,r,c-5),0.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.LEFT(E.v(17,r,c-3), 4.0),'Dụng'), X.eq(X.LEFT(E.v(17,r,c-3), 2.0),'Hỷ'))) else (X.neg(1.0) if X.truthy(X.eq(X.LEFT(E.v(17,r,c-3), 2.0),'Kỵ')) else 0.0)))),E.v(17,r,c-2))

def f172(E, r, c):
    # Ô W64
    # = IF(DOI_CHIEU!$C$33=1,U64,IF(DOI_CHIEU!$C$33=2,V64,INT((U64+V64)/2)))
    return (E.v(17,r,c-2) if X.truthy(X.eq(E.v(7,33,3),1.0)) else (E.v(17,r,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INT(X.div(X.add(E.v(17,r,c-2),E.v(17,r,c-1)),2.0))))

def f173(E, r, c):
    # Ô X64
    # = IF(W64>=3,"Tốt",IF(W64=2,"Khá",IF(W64>=0,"Trung bình",IF(W64>=-2,"Kém","Rất kém"))))
    return ('Tốt' if X.truthy(X.ge(E.v(17,r,c-1),3.0)) else ('Khá' if X.truthy(X.eq(E.v(17,r,c-1),2.0)) else ('Trung bình' if X.truthy(X.ge(E.v(17,r,c-1),0.0)) else ('Kém' if X.truthy(X.ge(E.v(17,r,c-1),X.neg(2.0))) else 'Rất kém'))))

def f174(E, r, c):
    # Ô Y64
    # = TRIM(IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C63,DAI_VAN!C48)="Tốt",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI
    #   _VAN!D63,DAI_VAN!D48)="Tốt",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E63,DAI_VAN!E48)="Tốt",DAI_
    #   VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F63,DAI_VAN!F48)="Tốt",DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DA
    #   I_VAN!F36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G63,DAI_VAN!G48)="Tốt",DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(IF
    #   (DOI_CHIEU!$C$33=2,DAI_VAN!H63,DAI_VAN!H48)="Tốt",DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I63,D
    #   AI_VAN!I48)="Tốt",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J63,DAI_VAN!J48)="Tốt",DAI_VAN!J35&" 
    #   ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-22),' ('),E.v(14,r-28,c-22)),'–'),X.add(E.v(14,r-28,c-22),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-22) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-22)),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-21),' ('),E.v(14,r-28,c-21)),'–'),X.add(E.v(14,r-28,c-21),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-21) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-21)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-20),' ('),E.v(14,r-28,c-20)),'–'),X.add(E.v(14,r-28,c-20),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-20) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-20)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-19),' ('),E.v(14,r-28,c-19)),'–'),X.add(E.v(14,r-28,c-19),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-19) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-19)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-18),' ('),E.v(14,r-28,c-18)),'–'),X.add(E.v(14,r-28,c-18),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-18) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-18)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-17),' ('),E.v(14,r-28,c-17)),'–'),X.add(E.v(14,r-28,c-17),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-17) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-17)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-16),' ('),E.v(14,r-28,c-16)),'–'),X.add(E.v(14,r-28,c-16),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-16) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-16)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-15),' ('),E.v(14,r-28,c-15)),'–'),X.add(E.v(14,r-28,c-15),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-15) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-15)),'Tốt')) else '')))

def f175(E, r, c):
    # Ô Z64
    # = TRIM(IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C63,DAI_VAN!C48)="Xấu",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI
    #   _VAN!D63,DAI_VAN!D48)="Xấu",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E63,DAI_VAN!E48)="Xấu",DAI_
    #   VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F63,DAI_VAN!F48)="Xấu",DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DA
    #   I_VAN!F36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G63,DAI_VAN!G48)="Xấu",DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(IF
    #   (DOI_CHIEU!$C$33=2,DAI_VAN!H63,DAI_VAN!H48)="Xấu",DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I63,D
    #   AI_VAN!I48)="Xấu",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J63,DAI_VAN!J48)="Xấu",DAI_VAN!J35&" 
    #   ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-23),' ('),E.v(14,r-28,c-23)),'–'),X.add(E.v(14,r-28,c-23),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-23) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-23)),'Xấu')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-22),' ('),E.v(14,r-28,c-22)),'–'),X.add(E.v(14,r-28,c-22),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-22) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-22)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-21),' ('),E.v(14,r-28,c-21)),'–'),X.add(E.v(14,r-28,c-21),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-21) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-21)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-20),' ('),E.v(14,r-28,c-20)),'–'),X.add(E.v(14,r-28,c-20),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-20) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-20)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-19),' ('),E.v(14,r-28,c-19)),'–'),X.add(E.v(14,r-28,c-19),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-19) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-19)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-18),' ('),E.v(14,r-28,c-18)),'–'),X.add(E.v(14,r-28,c-18),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-18) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-18)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-17),' ('),E.v(14,r-28,c-17)),'–'),X.add(E.v(14,r-28,c-17),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-17) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-17)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-29,c-16),' ('),E.v(14,r-28,c-16)),'–'),X.add(E.v(14,r-28,c-16),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-1,c-16) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-16,c-16)),'Xấu')) else '')))

def f176(E, r, c):
    # Ô AA64
    # = IF(NHAP_LIEU!$C$22="Nam",TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,OR(SAO!$F$5="Chính tài",SAO!$F$5="Thiên tài")),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!
    #   $Y$6=0,OR(SAO!$F$6="Chính tài",SAO!$F$6="Thiên tài")),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$7=1,SAO!$Y$7=0,OR(SAO!$F$7="Chính tài",SAO!$F$7="Thiên 
    #   tài")),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,OR(SAO!$F$8="Chính tài",SAO!$F$8="Thiên tài")),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X
    #   $9=1,SAO!$Y$9=0,OR(SAO!$F$9="Chính tài",SAO!$F$9="Thiên tài")),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$X$10=1,SAO!$Y$10=0,OR(SAO!$F$10="Chính tài",SAO!$F$
    #   10="Thiên tài")),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,OR(SAO!$F$11="Chính tài",SAO!$F$11="Thiên tài")),SAO!$E$11&"(Tháng tàng)
    #    ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,OR(SAO!$F$12="Chính tài",SAO!$F$12="Thiên tài")),SAO!$E$12&"(Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,OR
    #   (SAO!$F$13="Chính tài",SAO!$F$13="Thiên tài")),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14=0,OR(SAO!$F$14="Chính tài",SAO!$F$14="Thiên tài"))
    #   ,SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,OR(SAO!$F$15="Chính tài",SAO!$F$15="Thiên tài")),SAO!$E$15&"(Ngày tàng) ","")&IF(AND(SAO!
    #   $X$16=1,SAO!$Y$16=0,OR(SAO!$F$16="Chính tài",SAO!$F$16="Thiên tài")),SAO!$E$16&"(Ngày tàng) ","")&IF(AND(SAO!$X$17=1,SAO!$Y$17=0,OR(SAO!$F$17="Chính t
    #   ài",SAO!$F$17="Thiên tài")),SAO!$E$17&"(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,OR(SAO!$F$18="Chính tài",SAO!$F$18="Thiên tài")),SAO!$E$18&"(Giờ tàn
    #   g) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,OR(SAO!$F$19="Chính tài",SAO!$F$19="Thiên tài")),SAO!$E$19&"(Giờ tàng) ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,OR
    #   (SAO!$F$20="Chính tài",SAO!$F$20="Thiên tài")),SAO!$E$20&"(Giờ tàng) ","")),TRIM(IF(AND(SAO!$X$5=1,SAO!$Y$5=0,OR(SAO!$F$5="Chính quan",SAO!$F$5="Thất 
    #   sát")),SAO!$E$5&"(Năm) ","")&IF(AND(SAO!$X$6=1,SAO!$Y$6=0,OR(SAO!$F$6="Chính quan",SAO!$F$6="Thất sát")),SAO!$E$6&"(Năm tàng) ","")&IF(AND(SAO!$X$7=1,
    #   SAO!$Y$7=0,OR(SAO!$F$7="Chính quan",SAO!$F$7="Thất sát")),SAO!$E$7&"(Năm tàng) ","")&IF(AND(SAO!$X$8=1,SAO!$Y$8=0,OR(SAO!$F$8="Chính quan",SAO!$F$8="T
    #   hất sát")),SAO!$E$8&"(Năm tàng) ","")&IF(AND(SAO!$X$9=1,SAO!$Y$9=0,OR(SAO!$F$9="Chính quan",SAO!$F$9="Thất sát")),SAO!$E$9&"(Tháng) ","")&IF(AND(SAO!$
    #   X$10=1,SAO!$Y$10=0,OR(SAO!$F$10="Chính quan",SAO!$F$10="Thất sát")),SAO!$E$10&"(Tháng tàng) ","")&IF(AND(SAO!$X$11=1,SAO!$Y$11=0,OR(SAO!$F$11="Chính q
    #   uan",SAO!$F$11="Thất sát")),SAO!$E$11&"(Tháng tàng) ","")&IF(AND(SAO!$X$12=1,SAO!$Y$12=0,OR(SAO!$F$12="Chính quan",SAO!$F$12="Thất sát")),SAO!$E$12&"(
    #   Tháng tàng) ","")&IF(AND(SAO!$X$13=1,SAO!$Y$13=0,OR(SAO!$F$13="Chính quan",SAO!$F$13="Thất sát")),SAO!$E$13&"(Ngày) ","")&IF(AND(SAO!$X$14=1,SAO!$Y$14
    #   =0,OR(SAO!$F$14="Chính quan",SAO!$F$14="Thất sát")),SAO!$E$14&"(Ngày tàng) ","")&IF(AND(SAO!$X$15=1,SAO!$Y$15=0,OR(SAO!$F$15="Chính quan",SAO!$F$15="T
    #   hất sát")),SAO!$E$15&"(Ngày tàng) ","")&IF(AND(SAO!$X$16=1,SAO!$Y$16=0,OR(SAO!$F$16="Chính quan",SAO!$F$16="Thất sát")),SAO!$E$16&"(Ngày tàng) ","")&I
    #   F(AND(SAO!$X$17=1,SAO!$Y$17=0,OR(SAO!$F$17="Chính quan",SAO!$F$17="Thất sát")),SAO!$E$17&"(Giờ) ","")&IF(AND(SAO!$X$18=1,SAO!$Y$18=0,OR(SAO!$F$18="Chí
    #   nh quan",SAO!$F$18="Thất sát")),SAO!$E$18&"(Giờ tàng) ","")&IF(AND(SAO!$X$19=1,SAO!$Y$19=0,OR(SAO!$F$19="Chính quan",SAO!$F$19="Thất sát")),SAO!$E$19&
    #   "(Giờ tàng) ","")&IF(AND(SAO!$X$20=1,SAO!$Y$20=0,OR(SAO!$F$20="Chính quan",SAO!$F$20="Thất sát")),SAO!$E$20&"(Giờ tàng) ","")))
    return (X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.OR(X.eq(E.v(13,5,6),'Chính tài'), X.eq(E.v(13,5,6),'Thiên tài')))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.OR(X.eq(E.v(13,6,6),'Chính tài'), X.eq(E.v(13,6,6),'Thiên tài')))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.OR(X.eq(E.v(13,7,6),'Chính tài'), X.eq(E.v(13,7,6),'Thiên tài')))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.OR(X.eq(E.v(13,8,6),'Chính tài'), X.eq(E.v(13,8,6),'Thiên tài')))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.OR(X.eq(E.v(13,9,6),'Chính tài'), X.eq(E.v(13,9,6),'Thiên tài')))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.OR(X.eq(E.v(13,10,6),'Chính tài'), X.eq(E.v(13,10,6),'Thiên tài')))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.OR(X.eq(E.v(13,11,6),'Chính tài'), X.eq(E.v(13,11,6),'Thiên tài')))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.OR(X.eq(E.v(13,12,6),'Chính tài'), X.eq(E.v(13,12,6),'Thiên tài')))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.OR(X.eq(E.v(13,13,6),'Chính tài'), X.eq(E.v(13,13,6),'Thiên tài')))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.OR(X.eq(E.v(13,14,6),'Chính tài'), X.eq(E.v(13,14,6),'Thiên tài')))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.OR(X.eq(E.v(13,15,6),'Chính tài'), X.eq(E.v(13,15,6),'Thiên tài')))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.OR(X.eq(E.v(13,16,6),'Chính tài'), X.eq(E.v(13,16,6),'Thiên tài')))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.OR(X.eq(E.v(13,17,6),'Chính tài'), X.eq(E.v(13,17,6),'Thiên tài')))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.OR(X.eq(E.v(13,18,6),'Chính tài'), X.eq(E.v(13,18,6),'Thiên tài')))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.OR(X.eq(E.v(13,19,6),'Chính tài'), X.eq(E.v(13,19,6),'Thiên tài')))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.OR(X.eq(E.v(13,20,6),'Chính tài'), X.eq(E.v(13,20,6),'Thiên tài')))) else ''))) if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(13,5,5),'(Năm) ') if X.truthy(X.AND(X.eq(E.v(13,5,24),1.0), X.eq(E.v(13,5,25),0.0), X.OR(X.eq(E.v(13,5,6),'Chính quan'), X.eq(E.v(13,5,6),'Thất sát')))) else ''),(X.cat(E.v(13,6,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,6,24),1.0), X.eq(E.v(13,6,25),0.0), X.OR(X.eq(E.v(13,6,6),'Chính quan'), X.eq(E.v(13,6,6),'Thất sát')))) else '')),(X.cat(E.v(13,7,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,7,24),1.0), X.eq(E.v(13,7,25),0.0), X.OR(X.eq(E.v(13,7,6),'Chính quan'), X.eq(E.v(13,7,6),'Thất sát')))) else '')),(X.cat(E.v(13,8,5),'(Năm tàng) ') if X.truthy(X.AND(X.eq(E.v(13,8,24),1.0), X.eq(E.v(13,8,25),0.0), X.OR(X.eq(E.v(13,8,6),'Chính quan'), X.eq(E.v(13,8,6),'Thất sát')))) else '')),(X.cat(E.v(13,9,5),'(Tháng) ') if X.truthy(X.AND(X.eq(E.v(13,9,24),1.0), X.eq(E.v(13,9,25),0.0), X.OR(X.eq(E.v(13,9,6),'Chính quan'), X.eq(E.v(13,9,6),'Thất sát')))) else '')),(X.cat(E.v(13,10,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,10,24),1.0), X.eq(E.v(13,10,25),0.0), X.OR(X.eq(E.v(13,10,6),'Chính quan'), X.eq(E.v(13,10,6),'Thất sát')))) else '')),(X.cat(E.v(13,11,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,11,24),1.0), X.eq(E.v(13,11,25),0.0), X.OR(X.eq(E.v(13,11,6),'Chính quan'), X.eq(E.v(13,11,6),'Thất sát')))) else '')),(X.cat(E.v(13,12,5),'(Tháng tàng) ') if X.truthy(X.AND(X.eq(E.v(13,12,24),1.0), X.eq(E.v(13,12,25),0.0), X.OR(X.eq(E.v(13,12,6),'Chính quan'), X.eq(E.v(13,12,6),'Thất sát')))) else '')),(X.cat(E.v(13,13,5),'(Ngày) ') if X.truthy(X.AND(X.eq(E.v(13,13,24),1.0), X.eq(E.v(13,13,25),0.0), X.OR(X.eq(E.v(13,13,6),'Chính quan'), X.eq(E.v(13,13,6),'Thất sát')))) else '')),(X.cat(E.v(13,14,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,14,24),1.0), X.eq(E.v(13,14,25),0.0), X.OR(X.eq(E.v(13,14,6),'Chính quan'), X.eq(E.v(13,14,6),'Thất sát')))) else '')),(X.cat(E.v(13,15,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,15,24),1.0), X.eq(E.v(13,15,25),0.0), X.OR(X.eq(E.v(13,15,6),'Chính quan'), X.eq(E.v(13,15,6),'Thất sát')))) else '')),(X.cat(E.v(13,16,5),'(Ngày tàng) ') if X.truthy(X.AND(X.eq(E.v(13,16,24),1.0), X.eq(E.v(13,16,25),0.0), X.OR(X.eq(E.v(13,16,6),'Chính quan'), X.eq(E.v(13,16,6),'Thất sát')))) else '')),(X.cat(E.v(13,17,5),'(Giờ) ') if X.truthy(X.AND(X.eq(E.v(13,17,24),1.0), X.eq(E.v(13,17,25),0.0), X.OR(X.eq(E.v(13,17,6),'Chính quan'), X.eq(E.v(13,17,6),'Thất sát')))) else '')),(X.cat(E.v(13,18,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,18,24),1.0), X.eq(E.v(13,18,25),0.0), X.OR(X.eq(E.v(13,18,6),'Chính quan'), X.eq(E.v(13,18,6),'Thất sát')))) else '')),(X.cat(E.v(13,19,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,19,24),1.0), X.eq(E.v(13,19,25),0.0), X.OR(X.eq(E.v(13,19,6),'Chính quan'), X.eq(E.v(13,19,6),'Thất sát')))) else '')),(X.cat(E.v(13,20,5),'(Giờ tàng) ') if X.truthy(X.AND(X.eq(E.v(13,20,24),1.0), X.eq(E.v(13,20,25),0.0), X.OR(X.eq(E.v(13,20,6),'Chính quan'), X.eq(E.v(13,20,6),'Thất sát')))) else ''))))

def f177(E, r, c):
    # Ô AB64
    # = TRIM(IF(HON_NHAN!$J$72<>"","• "&HON_NHAN!$G$72&" ["&HON_NHAN!$B$72&IF(HON_NHAN!$J$72="cả hai","","; chỉ "&HON_NHAN!$J$72)&"] ","")&IF(HON_NHAN!$J$73<>
    #   "","• "&HON_NHAN!$G$73&" ["&HON_NHAN!$B$73&IF(HON_NHAN!$J$73="cả hai","","; chỉ "&HON_NHAN!$J$73)&"] ","")&IF(HON_NHAN!$J$74<>"","• "&HON_NHAN!$G$74&"
    #    ["&HON_NHAN!$B$74&IF(HON_NHAN!$J$74="cả hai","","; chỉ "&HON_NHAN!$J$74)&"] ","")&IF(HON_NHAN!$J$75<>"","• "&HON_NHAN!$G$75&" ["&HON_NHAN!$B$75&IF(HO
    #   N_NHAN!$J$75="cả hai","","; chỉ "&HON_NHAN!$J$75)&"] ","")&IF(HON_NHAN!$J$76<>"","• "&HON_NHAN!$G$76&" ["&HON_NHAN!$B$76&IF(HON_NHAN!$J$76="cả hai",""
    #   ,"; chỉ "&HON_NHAN!$J$76)&"] ","")&IF(HON_NHAN!$J$84<>"","• "&HON_NHAN!$G$84&" ["&HON_NHAN!$B$84&IF(HON_NHAN!$J$84="cả hai","","; chỉ "&HON_NHAN!$J$84
    #   )&"] ","")&IF(HON_NHAN!$J$105<>"","• "&HON_NHAN!$G$105&" ["&HON_NHAN!$B$105&IF(HON_NHAN!$J$105="cả hai","","; chỉ "&HON_NHAN!$J$105)&"] ","")&IF(HON_N
    #   HAN!$J$106<>"","• "&HON_NHAN!$G$106&" ["&HON_NHAN!$B$106&IF(HON_NHAN!$J$106="cả hai","","; chỉ "&HON_NHAN!$J$106)&"] ","")&IF(HON_NHAN!$J$107<>"","• "
    #   &HON_NHAN!$G$107&" ["&HON_NHAN!$B$107&IF(HON_NHAN!$J$107="cả hai","","; chỉ "&HON_NHAN!$J$107)&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,72,7)),' ['),E.v(17,72,2)),('' if X.truthy(X.eq(E.v(17,72,10),'cả hai')) else X.cat('; chỉ ',E.v(17,72,10)))),'] ') if X.truthy(X.ne(E.v(17,72,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,73,7)),' ['),E.v(17,73,2)),('' if X.truthy(X.eq(E.v(17,73,10),'cả hai')) else X.cat('; chỉ ',E.v(17,73,10)))),'] ') if X.truthy(X.ne(E.v(17,73,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,74,7)),' ['),E.v(17,74,2)),('' if X.truthy(X.eq(E.v(17,74,10),'cả hai')) else X.cat('; chỉ ',E.v(17,74,10)))),'] ') if X.truthy(X.ne(E.v(17,74,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,75,7)),' ['),E.v(17,75,2)),('' if X.truthy(X.eq(E.v(17,75,10),'cả hai')) else X.cat('; chỉ ',E.v(17,75,10)))),'] ') if X.truthy(X.ne(E.v(17,75,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,76,7)),' ['),E.v(17,76,2)),('' if X.truthy(X.eq(E.v(17,76,10),'cả hai')) else X.cat('; chỉ ',E.v(17,76,10)))),'] ') if X.truthy(X.ne(E.v(17,76,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,84,7)),' ['),E.v(17,84,2)),('' if X.truthy(X.eq(E.v(17,84,10),'cả hai')) else X.cat('; chỉ ',E.v(17,84,10)))),'] ') if X.truthy(X.ne(E.v(17,84,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,105,7)),' ['),E.v(17,105,2)),('' if X.truthy(X.eq(E.v(17,105,10),'cả hai')) else X.cat('; chỉ ',E.v(17,105,10)))),'] ') if X.truthy(X.ne(E.v(17,105,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,106,7)),' ['),E.v(17,106,2)),('' if X.truthy(X.eq(E.v(17,106,10),'cả hai')) else X.cat('; chỉ ',E.v(17,106,10)))),'] ') if X.truthy(X.ne(E.v(17,106,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,107,7)),' ['),E.v(17,107,2)),('' if X.truthy(X.eq(E.v(17,107,10),'cả hai')) else X.cat('; chỉ ',E.v(17,107,10)))),'] ') if X.truthy(X.ne(E.v(17,107,10),'')) else '')))

def f178(E, r, c):
    # Ô AC64
    # = TRIM(IF(HON_NHAN!$J$77<>"","• "&HON_NHAN!$G$77&" ["&HON_NHAN!$B$77&IF(HON_NHAN!$J$77="cả hai","","; chỉ "&HON_NHAN!$J$77)&"] ","")&IF(HON_NHAN!$J$79<>
    #   "","• "&HON_NHAN!$G$79&" ["&HON_NHAN!$B$79&IF(HON_NHAN!$J$79="cả hai","","; chỉ "&HON_NHAN!$J$79)&"] ","")&IF(HON_NHAN!$J$80<>"","• "&HON_NHAN!$G$80&"
    #    ["&HON_NHAN!$B$80&IF(HON_NHAN!$J$80="cả hai","","; chỉ "&HON_NHAN!$J$80)&"] ","")&IF(HON_NHAN!$J$85<>"","• "&HON_NHAN!$G$85&" ["&HON_NHAN!$B$85&IF(HO
    #   N_NHAN!$J$85="cả hai","","; chỉ "&HON_NHAN!$J$85)&"] ","")&IF(HON_NHAN!$J$86<>"","• "&HON_NHAN!$G$86&" ["&HON_NHAN!$B$86&IF(HON_NHAN!$J$86="cả hai",""
    #   ,"; chỉ "&HON_NHAN!$J$86)&"] ","")&IF(HON_NHAN!$J$88<>"","• "&HON_NHAN!$G$88&" ["&HON_NHAN!$B$88&IF(HON_NHAN!$J$88="cả hai","","; chỉ "&HON_NHAN!$J$88
    #   )&"] ","")&IF(HON_NHAN!$J$89<>"","• "&HON_NHAN!$G$89&" ["&HON_NHAN!$B$89&IF(HON_NHAN!$J$89="cả hai","","; chỉ "&HON_NHAN!$J$89)&"] ","")&IF(HON_NHAN!$
    #   J$90<>"","• "&HON_NHAN!$G$90&" ["&HON_NHAN!$B$90&IF(HON_NHAN!$J$90="cả hai","","; chỉ "&HON_NHAN!$J$90)&"] ","")&IF(HON_NHAN!$J$91<>"","• "&HON_NHAN!$
    #   G$91&" ["&HON_NHAN!$B$91&IF(HON_NHAN!$J$91="cả hai","","; chỉ "&HON_NHAN!$J$91)&"] ","")&IF(HON_NHAN!$J$92<>"","• "&HON_NHAN!$G$92&" ["&HON_NHAN!$B$92
    #   &IF(HON_NHAN!$J$92="cả hai","","; chỉ "&HON_NHAN!$J$92)&"] ","")&IF(HON_NHAN!$J$93<>"","• "&HON_NHAN!$G$93&" ["&HON_NHAN!$B$93&IF(HON_NHAN!$J$93="cả h
    #   ai","","; chỉ "&HON_NHAN!$J$93)&"] ","")&IF(HON_NHAN!$J$94<>"","• "&HON_NHAN!$G$94&" ["&HON_NHAN!$B$94&IF(HON_NHAN!$J$94="cả hai","","; chỉ "&HON_NHAN
    #   !$J$94)&"] ","")&IF(HON_NHAN!$J$95<>"","• "&HON_NHAN!$G$95&" ["&HON_NHAN!$B$95&IF(HON_NHAN!$J$95="cả hai","","; chỉ "&HON_NHAN!$J$95)&"] ","")&IF(HON_
    #   NHAN!$J$96<>"","• "&HON_NHAN!$G$96&" ["&HON_NHAN!$B$96&IF(HON_NHAN!$J$96="cả hai","","; chỉ "&HON_NHAN!$J$96)&"] ","")&IF(HON_NHAN!$J$97<>"","• "&HON_
    #   NHAN!$G$97&" ["&HON_NHAN!$B$97&IF(HON_NHAN!$J$97="cả hai","","; chỉ "&HON_NHAN!$J$97)&"] ","")&IF(HON_NHAN!$J$98<>"","• "&HON_NHAN!$G$98&" ["&HON_NHAN
    #   !$B$98&IF(HON_NHAN!$J$98="cả hai","","; chỉ "&HON_NHAN!$J$98)&"] ","")&IF(HON_NHAN!$J$99<>"","• "&HON_NHAN!$G$99&" ["&HON_NHAN!$B$99&IF(HON_NHAN!$J$99
    #   ="cả hai","","; chỉ "&HON_NHAN!$J$99)&"] ","")&IF(HON_NHAN!$J$100<>"","• "&HON_NHAN!$G$100&" ["&HON_NHAN!$B$100&IF(HON_NHAN!$J$100="cả hai","","; chỉ 
    #   "&HON_NHAN!$J$100)&"] ","")&IF(HON_NHAN!$J$101<>"","• "&HON_NHAN!$G$101&" ["&HON_NHAN!$B$101&IF(HON_NHAN!$J$101="cả hai","","; chỉ "&HON_NHAN!$J$101)&
    #   "] ","")&IF(HON_NHAN!$J$102<>"","• "&HON_NHAN!$G$102&" ["&HON_NHAN!$B$102&IF(HON_NHAN!$J$102="cả hai","","; chỉ "&HON_NHAN!$J$102)&"] ","")&IF(HON_NHA
    #   N!$J$103<>"","• "&HON_NHAN!$G$103&" ["&HON_NHAN!$B$103&IF(HON_NHAN!$J$103="cả hai","","; chỉ "&HON_NHAN!$J$103)&"] ","")&IF(HON_NHAN!$J$104<>"","• "&H
    #   ON_NHAN!$G$104&" ["&HON_NHAN!$B$104&IF(HON_NHAN!$J$104="cả hai","","; chỉ "&HON_NHAN!$J$104)&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,77,7)),' ['),E.v(17,77,2)),('' if X.truthy(X.eq(E.v(17,77,10),'cả hai')) else X.cat('; chỉ ',E.v(17,77,10)))),'] ') if X.truthy(X.ne(E.v(17,77,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,79,7)),' ['),E.v(17,79,2)),('' if X.truthy(X.eq(E.v(17,79,10),'cả hai')) else X.cat('; chỉ ',E.v(17,79,10)))),'] ') if X.truthy(X.ne(E.v(17,79,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,80,7)),' ['),E.v(17,80,2)),('' if X.truthy(X.eq(E.v(17,80,10),'cả hai')) else X.cat('; chỉ ',E.v(17,80,10)))),'] ') if X.truthy(X.ne(E.v(17,80,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,85,7)),' ['),E.v(17,85,2)),('' if X.truthy(X.eq(E.v(17,85,10),'cả hai')) else X.cat('; chỉ ',E.v(17,85,10)))),'] ') if X.truthy(X.ne(E.v(17,85,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,86,7)),' ['),E.v(17,86,2)),('' if X.truthy(X.eq(E.v(17,86,10),'cả hai')) else X.cat('; chỉ ',E.v(17,86,10)))),'] ') if X.truthy(X.ne(E.v(17,86,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,88,7)),' ['),E.v(17,88,2)),('' if X.truthy(X.eq(E.v(17,88,10),'cả hai')) else X.cat('; chỉ ',E.v(17,88,10)))),'] ') if X.truthy(X.ne(E.v(17,88,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,89,7)),' ['),E.v(17,89,2)),('' if X.truthy(X.eq(E.v(17,89,10),'cả hai')) else X.cat('; chỉ ',E.v(17,89,10)))),'] ') if X.truthy(X.ne(E.v(17,89,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,90,7)),' ['),E.v(17,90,2)),('' if X.truthy(X.eq(E.v(17,90,10),'cả hai')) else X.cat('; chỉ ',E.v(17,90,10)))),'] ') if X.truthy(X.ne(E.v(17,90,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,91,7)),' ['),E.v(17,91,2)),('' if X.truthy(X.eq(E.v(17,91,10),'cả hai')) else X.cat('; chỉ ',E.v(17,91,10)))),'] ') if X.truthy(X.ne(E.v(17,91,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,92,7)),' ['),E.v(17,92,2)),('' if X.truthy(X.eq(E.v(17,92,10),'cả hai')) else X.cat('; chỉ ',E.v(17,92,10)))),'] ') if X.truthy(X.ne(E.v(17,92,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,93,7)),' ['),E.v(17,93,2)),('' if X.truthy(X.eq(E.v(17,93,10),'cả hai')) else X.cat('; chỉ ',E.v(17,93,10)))),'] ') if X.truthy(X.ne(E.v(17,93,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,94,7)),' ['),E.v(17,94,2)),('' if X.truthy(X.eq(E.v(17,94,10),'cả hai')) else X.cat('; chỉ ',E.v(17,94,10)))),'] ') if X.truthy(X.ne(E.v(17,94,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,95,7)),' ['),E.v(17,95,2)),('' if X.truthy(X.eq(E.v(17,95,10),'cả hai')) else X.cat('; chỉ ',E.v(17,95,10)))),'] ') if X.truthy(X.ne(E.v(17,95,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,96,7)),' ['),E.v(17,96,2)),('' if X.truthy(X.eq(E.v(17,96,10),'cả hai')) else X.cat('; chỉ ',E.v(17,96,10)))),'] ') if X.truthy(X.ne(E.v(17,96,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,97,7)),' ['),E.v(17,97,2)),('' if X.truthy(X.eq(E.v(17,97,10),'cả hai')) else X.cat('; chỉ ',E.v(17,97,10)))),'] ') if X.truthy(X.ne(E.v(17,97,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,98,7)),' ['),E.v(17,98,2)),('' if X.truthy(X.eq(E.v(17,98,10),'cả hai')) else X.cat('; chỉ ',E.v(17,98,10)))),'] ') if X.truthy(X.ne(E.v(17,98,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,99,7)),' ['),E.v(17,99,2)),('' if X.truthy(X.eq(E.v(17,99,10),'cả hai')) else X.cat('; chỉ ',E.v(17,99,10)))),'] ') if X.truthy(X.ne(E.v(17,99,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,100,7)),' ['),E.v(17,100,2)),('' if X.truthy(X.eq(E.v(17,100,10),'cả hai')) else X.cat('; chỉ ',E.v(17,100,10)))),'] ') if X.truthy(X.ne(E.v(17,100,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,101,7)),' ['),E.v(17,101,2)),('' if X.truthy(X.eq(E.v(17,101,10),'cả hai')) else X.cat('; chỉ ',E.v(17,101,10)))),'] ') if X.truthy(X.ne(E.v(17,101,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,102,7)),' ['),E.v(17,102,2)),('' if X.truthy(X.eq(E.v(17,102,10),'cả hai')) else X.cat('; chỉ ',E.v(17,102,10)))),'] ') if X.truthy(X.ne(E.v(17,102,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,103,7)),' ['),E.v(17,103,2)),('' if X.truthy(X.eq(E.v(17,103,10),'cả hai')) else X.cat('; chỉ ',E.v(17,103,10)))),'] ') if X.truthy(X.ne(E.v(17,103,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,104,7)),' ['),E.v(17,104,2)),('' if X.truthy(X.eq(E.v(17,104,10),'cả hai')) else X.cat('; chỉ ',E.v(17,104,10)))),'] ') if X.truthy(X.ne(E.v(17,104,10),'')) else '')))

def f179(E, r, c):
    # Ô AD64
    # = TRIM(IF(HON_NHAN!$J$78<>"","• "&HON_NHAN!$G$78&" ["&HON_NHAN!$B$78&IF(HON_NHAN!$J$78="cả hai","","; chỉ "&HON_NHAN!$J$78)&"] ","")&IF(HON_NHAN!$J$81<>
    #   "","• "&HON_NHAN!$G$81&" ["&HON_NHAN!$B$81&IF(HON_NHAN!$J$81="cả hai","","; chỉ "&HON_NHAN!$J$81)&"] ","")&IF(HON_NHAN!$J$82<>"","• "&HON_NHAN!$G$82&"
    #    ["&HON_NHAN!$B$82&IF(HON_NHAN!$J$82="cả hai","","; chỉ "&HON_NHAN!$J$82)&"] ","")&IF(HON_NHAN!$J$83<>"","• "&HON_NHAN!$G$83&" ["&HON_NHAN!$B$83&IF(HO
    #   N_NHAN!$J$83="cả hai","","; chỉ "&HON_NHAN!$J$83)&"] ","")&IF(HON_NHAN!$J$87<>"","• "&HON_NHAN!$G$87&" ["&HON_NHAN!$B$87&IF(HON_NHAN!$J$87="cả hai",""
    #   ,"; chỉ "&HON_NHAN!$J$87)&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,78,7)),' ['),E.v(17,78,2)),('' if X.truthy(X.eq(E.v(17,78,10),'cả hai')) else X.cat('; chỉ ',E.v(17,78,10)))),'] ') if X.truthy(X.ne(E.v(17,78,10),'')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,81,7)),' ['),E.v(17,81,2)),('' if X.truthy(X.eq(E.v(17,81,10),'cả hai')) else X.cat('; chỉ ',E.v(17,81,10)))),'] ') if X.truthy(X.ne(E.v(17,81,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,82,7)),' ['),E.v(17,82,2)),('' if X.truthy(X.eq(E.v(17,82,10),'cả hai')) else X.cat('; chỉ ',E.v(17,82,10)))),'] ') if X.truthy(X.ne(E.v(17,82,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,83,7)),' ['),E.v(17,83,2)),('' if X.truthy(X.eq(E.v(17,83,10),'cả hai')) else X.cat('; chỉ ',E.v(17,83,10)))),'] ') if X.truthy(X.ne(E.v(17,83,10),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,87,7)),' ['),E.v(17,87,2)),('' if X.truthy(X.eq(E.v(17,87,10),'cả hai')) else X.cat('; chỉ ',E.v(17,87,10)))),'] ') if X.truthy(X.ne(E.v(17,87,10),'')) else '')))

def f180(E, r, c):
    # Ô AF64
    # = TRIM(IF(AND(I72="Bất lợi",OR(J72="cả hai",J72=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G72&" ["&B72&", "&H72&"] ","")&IF(AND(I73="Bất lợi",OR(J73="cả 
    #   hai",J73=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G73&" ["&B73&", "&H73&"] ","")&IF(AND(I74="Bất lợi",OR(J74="cả hai",J74=IF(DOI_CHIEU!$C$33=2,"PP2","
    #   PP1"))),"• "&G74&" ["&B74&", "&H74&"] ","")&IF(AND(I75="Bất lợi",OR(J75="cả hai",J75=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G75&" ["&B75&", "&H75&"]
    #    ","")&IF(AND(I76="Bất lợi",OR(J76="cả hai",J76=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G76&" ["&B76&", "&H76&"] ","")&IF(AND(I77="Bất lợi",OR(J77="c
    #   ả hai",J77=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G77&" ["&B77&", "&H77&"] ","")&IF(AND(I78="Bất lợi",OR(J78="cả hai",J78=IF(DOI_CHIEU!$C$33=2,"PP2"
    #   ,"PP1"))),"• "&G78&" ["&B78&", "&H78&"] ","")&IF(AND(I79="Bất lợi",OR(J79="cả hai",J79=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G79&" ["&B79&", "&H79&
    #   "] ","")&IF(AND(I80="Bất lợi",OR(J80="cả hai",J80=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G80&" ["&B80&", "&H80&"] ","")&IF(AND(I81="Bất lợi",OR(J81=
    #   "cả hai",J81=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G81&" ["&B81&", "&H81&"] ","")&IF(AND(I82="Bất lợi",OR(J82="cả hai",J82=IF(DOI_CHIEU!$C$33=2,"PP
    #   2","PP1"))),"• "&G82&" ["&B82&", "&H82&"] ","")&IF(AND(I83="Bất lợi",OR(J83="cả hai",J83=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G83&" ["&B83&", "&H8
    #   3&"] ","")&IF(AND(I84="Bất lợi",OR(J84="cả hai",J84=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G84&" ["&B84&", "&H84&"] ","")&IF(AND(I85="Bất lợi",OR(J8
    #   5="cả hai",J85=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G85&" ["&B85&", "&H85&"] ","")&IF(AND(I86="Bất lợi",OR(J86="cả hai",J86=IF(DOI_CHIEU!$C$33=2,"
    #   PP2","PP1"))),"• "&G86&" ["&B86&", "&H86&"] ","")&IF(AND(I87="Bất lợi",OR(J87="cả hai",J87=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G87&" ["&B87&", "&
    #   H87&"] ","")&IF(AND(I88="Bất lợi",OR(J88="cả hai",J88=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G88&" ["&B88&", "&H88&"] ","")&IF(AND(I89="Bất lợi",OR(
    #   J89="cả hai",J89=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G89&" ["&B89&", "&H89&"] ","")&IF(AND(I90="Bất lợi",OR(J90="cả hai",J90=IF(DOI_CHIEU!$C$33=2
    #   ,"PP2","PP1"))),"• "&G90&" ["&B90&", "&H90&"] ","")&IF(AND(I91="Bất lợi",OR(J91="cả hai",J91=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G91&" ["&B91&", 
    #   "&H91&"] ","")&IF(AND(I92="Bất lợi",OR(J92="cả hai",J92=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G92&" ["&B92&", "&H92&"] ","")&IF(AND(I93="Bất lợi",O
    #   R(J93="cả hai",J93=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G93&" ["&B93&", "&H93&"] ","")&IF(AND(I94="Bất lợi",OR(J94="cả hai",J94=IF(DOI_CHIEU!$C$33
    #   =2,"PP2","PP1"))),"• "&G94&" ["&B94&", "&H94&"] ","")&IF(AND(I95="Bất lợi",OR(J95="cả hai",J95=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G95&" ["&B95&"
    #   , "&H95&"] ","")&IF(AND(I96="Bất lợi",OR(J96="cả hai",J96=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G96&" ["&B96&", "&H96&"] ","")&IF(AND(I97="Bất lợi"
    #   ,OR(J97="cả hai",J97=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G97&" ["&B97&", "&H97&"] ","")&IF(AND(I98="Bất lợi",OR(J98="cả hai",J98=IF(DOI_CHIEU!$C$
    #   33=2,"PP2","PP1"))),"• "&G98&" ["&B98&", "&H98&"] ","")&IF(AND(I99="Bất lợi",OR(J99="cả hai",J99=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G99&" ["&B99
    #   &", "&H99&"] ","")&IF(AND(I100="Bất lợi",OR(J100="cả hai",J100=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G100&" ["&B100&", "&H100&"] ","")&IF(AND(I101=
    #   "Bất lợi",OR(J101="cả hai",J101=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G101&" ["&B101&", "&H101&"] ","")&IF(AND(I102="Bất lợi",OR(J102="cả hai",J102
    #   =IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G102&" ["&B102&", "&H102&"] ","")&IF(AND(I103="Bất lợi",OR(J103="cả hai",J103=IF(DOI_CHIEU!$C$33=2,"PP2","PP
    #   1"))),"• "&G103&" ["&B103&", "&H103&"] ","")&IF(AND(I104="Bất lợi",OR(J104="cả hai",J104=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G104&" ["&B104&", "&
    #   H104&"] ","")&IF(AND(I105="Bất lợi",OR(J105="cả hai",J105=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G105&" ["&B105&", "&H105&"] ","")&IF(AND(I106="Bất 
    #   lợi",OR(J106="cả hai",J106=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G106&" ["&B106&", "&H106&"] ","")&IF(AND(I107="Bất lợi",OR(J107="cả hai",J107=IF(D
    #   OI_CHIEU!$C$33=2,"PP2","PP1"))),"• "&G107&" ["&B107&", "&H107&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+8,c-25)),' ['),E.v(17,r+8,c-30)),', '),E.v(17,r+8,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+8,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+8,c-22),'cả hai'), X.eq(E.v(17,r+8,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+9,c-25)),' ['),E.v(17,r+9,c-30)),', '),E.v(17,r+9,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+9,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+9,c-22),'cả hai'), X.eq(E.v(17,r+9,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+10,c-25)),' ['),E.v(17,r+10,c-30)),', '),E.v(17,r+10,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+10,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+10,c-22),'cả hai'), X.eq(E.v(17,r+10,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+11,c-25)),' ['),E.v(17,r+11,c-30)),', '),E.v(17,r+11,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+11,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+11,c-22),'cả hai'), X.eq(E.v(17,r+11,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+12,c-25)),' ['),E.v(17,r+12,c-30)),', '),E.v(17,r+12,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+12,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+12,c-22),'cả hai'), X.eq(E.v(17,r+12,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+13,c-25)),' ['),E.v(17,r+13,c-30)),', '),E.v(17,r+13,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+13,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+13,c-22),'cả hai'), X.eq(E.v(17,r+13,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+14,c-25)),' ['),E.v(17,r+14,c-30)),', '),E.v(17,r+14,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+14,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+14,c-22),'cả hai'), X.eq(E.v(17,r+14,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+15,c-25)),' ['),E.v(17,r+15,c-30)),', '),E.v(17,r+15,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+15,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+15,c-22),'cả hai'), X.eq(E.v(17,r+15,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+16,c-25)),' ['),E.v(17,r+16,c-30)),', '),E.v(17,r+16,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+16,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+16,c-22),'cả hai'), X.eq(E.v(17,r+16,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+17,c-25)),' ['),E.v(17,r+17,c-30)),', '),E.v(17,r+17,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+17,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+17,c-22),'cả hai'), X.eq(E.v(17,r+17,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+18,c-25)),' ['),E.v(17,r+18,c-30)),', '),E.v(17,r+18,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+18,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+18,c-22),'cả hai'), X.eq(E.v(17,r+18,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+19,c-25)),' ['),E.v(17,r+19,c-30)),', '),E.v(17,r+19,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+19,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+19,c-22),'cả hai'), X.eq(E.v(17,r+19,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+20,c-25)),' ['),E.v(17,r+20,c-30)),', '),E.v(17,r+20,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+20,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+20,c-22),'cả hai'), X.eq(E.v(17,r+20,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+21,c-25)),' ['),E.v(17,r+21,c-30)),', '),E.v(17,r+21,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+21,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+21,c-22),'cả hai'), X.eq(E.v(17,r+21,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+22,c-25)),' ['),E.v(17,r+22,c-30)),', '),E.v(17,r+22,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+22,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+22,c-22),'cả hai'), X.eq(E.v(17,r+22,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+23,c-25)),' ['),E.v(17,r+23,c-30)),', '),E.v(17,r+23,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+23,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+23,c-22),'cả hai'), X.eq(E.v(17,r+23,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+24,c-25)),' ['),E.v(17,r+24,c-30)),', '),E.v(17,r+24,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+24,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+24,c-22),'cả hai'), X.eq(E.v(17,r+24,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+25,c-25)),' ['),E.v(17,r+25,c-30)),', '),E.v(17,r+25,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+25,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+25,c-22),'cả hai'), X.eq(E.v(17,r+25,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+26,c-25)),' ['),E.v(17,r+26,c-30)),', '),E.v(17,r+26,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+26,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+26,c-22),'cả hai'), X.eq(E.v(17,r+26,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+27,c-25)),' ['),E.v(17,r+27,c-30)),', '),E.v(17,r+27,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+27,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+27,c-22),'cả hai'), X.eq(E.v(17,r+27,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+28,c-25)),' ['),E.v(17,r+28,c-30)),', '),E.v(17,r+28,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+28,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+28,c-22),'cả hai'), X.eq(E.v(17,r+28,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+29,c-25)),' ['),E.v(17,r+29,c-30)),', '),E.v(17,r+29,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+29,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+29,c-22),'cả hai'), X.eq(E.v(17,r+29,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+30,c-25)),' ['),E.v(17,r+30,c-30)),', '),E.v(17,r+30,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+30,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+30,c-22),'cả hai'), X.eq(E.v(17,r+30,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+31,c-25)),' ['),E.v(17,r+31,c-30)),', '),E.v(17,r+31,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+31,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+31,c-22),'cả hai'), X.eq(E.v(17,r+31,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+32,c-25)),' ['),E.v(17,r+32,c-30)),', '),E.v(17,r+32,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+32,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+32,c-22),'cả hai'), X.eq(E.v(17,r+32,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+33,c-25)),' ['),E.v(17,r+33,c-30)),', '),E.v(17,r+33,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+33,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+33,c-22),'cả hai'), X.eq(E.v(17,r+33,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+34,c-25)),' ['),E.v(17,r+34,c-30)),', '),E.v(17,r+34,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+34,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+34,c-22),'cả hai'), X.eq(E.v(17,r+34,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+35,c-25)),' ['),E.v(17,r+35,c-30)),', '),E.v(17,r+35,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+35,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+35,c-22),'cả hai'), X.eq(E.v(17,r+35,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+36,c-25)),' ['),E.v(17,r+36,c-30)),', '),E.v(17,r+36,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+36,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+36,c-22),'cả hai'), X.eq(E.v(17,r+36,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+37,c-25)),' ['),E.v(17,r+37,c-30)),', '),E.v(17,r+37,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+37,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+37,c-22),'cả hai'), X.eq(E.v(17,r+37,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+38,c-25)),' ['),E.v(17,r+38,c-30)),', '),E.v(17,r+38,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+38,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+38,c-22),'cả hai'), X.eq(E.v(17,r+38,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+39,c-25)),' ['),E.v(17,r+39,c-30)),', '),E.v(17,r+39,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+39,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+39,c-22),'cả hai'), X.eq(E.v(17,r+39,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+40,c-25)),' ['),E.v(17,r+40,c-30)),', '),E.v(17,r+40,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+40,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+40,c-22),'cả hai'), X.eq(E.v(17,r+40,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+41,c-25)),' ['),E.v(17,r+41,c-30)),', '),E.v(17,r+41,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+41,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+41,c-22),'cả hai'), X.eq(E.v(17,r+41,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+42,c-25)),' ['),E.v(17,r+42,c-30)),', '),E.v(17,r+42,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+42,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+42,c-22),'cả hai'), X.eq(E.v(17,r+42,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+43,c-25)),' ['),E.v(17,r+43,c-30)),', '),E.v(17,r+43,c-24)),'] ') if X.truthy(X.AND(X.eq(E.v(17,r+43,c-23),'Bất lợi'), X.OR(X.eq(E.v(17,r+43,c-22),'cả hai'), X.eq(E.v(17,r+43,c-22),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else '')))

def f181(E, r, c):
    # Ô AG64
    # = IF(AND(I72="Bất lợi",OR(J72="cả hai",J72=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I73="Bất lợi",OR(J73="cả hai",J73=IF(DOI_CHIEU!$C$33=2,"PP2",
    #   "PP1"))),1,0)+IF(AND(I74="Bất lợi",OR(J74="cả hai",J74=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I75="Bất lợi",OR(J75="cả hai",J75=IF(DOI_CHIEU!
    #   $C$33=2,"PP2","PP1"))),1,0)+IF(AND(I76="Bất lợi",OR(J76="cả hai",J76=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I77="Bất lợi",OR(J77="cả hai",J77
    #   =IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I78="Bất lợi",OR(J78="cả hai",J78=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I79="Bất lợi",OR(J7
    #   9="cả hai",J79=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I80="Bất lợi",OR(J80="cả hai",J80=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I81="
    #   Bất lợi",OR(J81="cả hai",J81=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I82="Bất lợi",OR(J82="cả hai",J82=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0
    #   )+IF(AND(I83="Bất lợi",OR(J83="cả hai",J83=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I84="Bất lợi",OR(J84="cả hai",J84=IF(DOI_CHIEU!$C$33=2,"PP2
    #   ","PP1"))),1,0)+IF(AND(I85="Bất lợi",OR(J85="cả hai",J85=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I86="Bất lợi",OR(J86="cả hai",J86=IF(DOI_CHIE
    #   U!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I87="Bất lợi",OR(J87="cả hai",J87=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I88="Bất lợi",OR(J88="cả hai",J
    #   88=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I89="Bất lợi",OR(J89="cả hai",J89=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I90="Bất lợi",OR(
    #   J90="cả hai",J90=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I91="Bất lợi",OR(J91="cả hai",J91=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I92
    #   ="Bất lợi",OR(J92="cả hai",J92=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I93="Bất lợi",OR(J93="cả hai",J93=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1
    #   ,0)+IF(AND(I94="Bất lợi",OR(J94="cả hai",J94=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I95="Bất lợi",OR(J95="cả hai",J95=IF(DOI_CHIEU!$C$33=2,"P
    #   P2","PP1"))),1,0)+IF(AND(I96="Bất lợi",OR(J96="cả hai",J96=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I97="Bất lợi",OR(J97="cả hai",J97=IF(DOI_CH
    #   IEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I98="Bất lợi",OR(J98="cả hai",J98=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I99="Bất lợi",OR(J99="cả hai"
    #   ,J99=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I100="Bất lợi",OR(J100="cả hai",J100=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I101="Bất lợ
    #   i",OR(J101="cả hai",J101=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I102="Bất lợi",OR(J102="cả hai",J102=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)
    #   +IF(AND(I103="Bất lợi",OR(J103="cả hai",J103=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I104="Bất lợi",OR(J104="cả hai",J104=IF(DOI_CHIEU!$C$33=2
    #   ,"PP2","PP1"))),1,0)+IF(AND(I105="Bất lợi",OR(J105="cả hai",J105=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I106="Bất lợi",OR(J106="cả hai",J106=
    #   IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I107="Bất lợi",OR(J107="cả hai",J107=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.eq(E.v(17,r+8,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+8,c-23),'cả hai'), X.eq(E.v(17,r+8,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+9,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+9,c-23),'cả hai'), X.eq(E.v(17,r+9,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+10,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+10,c-23),'cả hai'), X.eq(E.v(17,r+10,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+11,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+11,c-23),'cả hai'), X.eq(E.v(17,r+11,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+12,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+12,c-23),'cả hai'), X.eq(E.v(17,r+12,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+13,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+13,c-23),'cả hai'), X.eq(E.v(17,r+13,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+14,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+14,c-23),'cả hai'), X.eq(E.v(17,r+14,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+15,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+15,c-23),'cả hai'), X.eq(E.v(17,r+15,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+16,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+16,c-23),'cả hai'), X.eq(E.v(17,r+16,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+17,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+17,c-23),'cả hai'), X.eq(E.v(17,r+17,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+18,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+18,c-23),'cả hai'), X.eq(E.v(17,r+18,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+19,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+19,c-23),'cả hai'), X.eq(E.v(17,r+19,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+20,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+20,c-23),'cả hai'), X.eq(E.v(17,r+20,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+21,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+21,c-23),'cả hai'), X.eq(E.v(17,r+21,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+22,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+22,c-23),'cả hai'), X.eq(E.v(17,r+22,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+23,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+23,c-23),'cả hai'), X.eq(E.v(17,r+23,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+24,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+24,c-23),'cả hai'), X.eq(E.v(17,r+24,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+25,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+25,c-23),'cả hai'), X.eq(E.v(17,r+25,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+26,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+26,c-23),'cả hai'), X.eq(E.v(17,r+26,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+27,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+27,c-23),'cả hai'), X.eq(E.v(17,r+27,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+28,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+28,c-23),'cả hai'), X.eq(E.v(17,r+28,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+29,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+29,c-23),'cả hai'), X.eq(E.v(17,r+29,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+30,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+30,c-23),'cả hai'), X.eq(E.v(17,r+30,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+31,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+31,c-23),'cả hai'), X.eq(E.v(17,r+31,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+32,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+32,c-23),'cả hai'), X.eq(E.v(17,r+32,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+33,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+33,c-23),'cả hai'), X.eq(E.v(17,r+33,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+34,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+34,c-23),'cả hai'), X.eq(E.v(17,r+34,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+35,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+35,c-23),'cả hai'), X.eq(E.v(17,r+35,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+36,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+36,c-23),'cả hai'), X.eq(E.v(17,r+36,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+37,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+37,c-23),'cả hai'), X.eq(E.v(17,r+37,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+38,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+38,c-23),'cả hai'), X.eq(E.v(17,r+38,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+39,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+39,c-23),'cả hai'), X.eq(E.v(17,r+39,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+40,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+40,c-23),'cả hai'), X.eq(E.v(17,r+40,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+41,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+41,c-23),'cả hai'), X.eq(E.v(17,r+41,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+42,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+42,c-23),'cả hai'), X.eq(E.v(17,r+42,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+43,c-24),'Bất lợi'), X.OR(X.eq(E.v(17,r+43,c-23),'cả hai'), X.eq(E.v(17,r+43,c-23),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0))

def f182(E, r, c):
    # Ô AH64
    # = UPPER(X64)&": "&IF(NHAP_LIEU!$C$22="Nam",CHOOSE(MATCH(X64,{"Rất kém","Kém","Trung bình","Khá","Tốt"},0),"Hôn nhân bất lợi rõ: dễ ly tán, khắc vợ hoặc 
    #   khó lập gia đình.","Hôn nhân kém: vợ yếu thế hoặc không hợp, hay bất hòa; nên kết hôn muộn.","Hôn nhân bình thường, tốt xấu tùy vận; cần vun đắp.","Vợ
    #    khá: hôn nhân tương đối ổn, có lúc trục trặc nhỏ.","Vợ tốt và có lực: hiền, đảm đang, giúp chồng; hôn nhân bền."),CHOOSE(MATCH(X64,{"Rất kém","Kém","
    #   Trung bình","Khá","Tốt"},0),"Hôn nhân bất lợi rõ: dễ ly tán, khắc chồng hoặc khó lập gia đình.","Hôn nhân kém: chồng yếu thế hoặc không hợp, hay bất h
    #   òa; nên kết hôn muộn.","Hôn nhân bình thường, tốt xấu tùy vận; cần vun đắp.","Chồng khá: hôn nhân tương đối ổn, có lúc trục trặc nhỏ.","Chồng tốt và c
    #   ó lực: chồng giỏi giang, giúp đỡ vợ; hôn nhân bền."))&IF(X64="Trung bình"," Khởi sắc trong đại vận dụng/hỷ thần: "&IF(Y64="","không có",Y64)&"; kém đi
    #    trong đại vận kỵ thần: "&IF(Z64="","không có",Z64)&".","")&IF(IF(U64>=3,"Tốt",IF(U64=2,"Khá",IF(U64>=0,"Trung bình",IF(U64>=-2,"Kém","Rất kém"))))<>I
    #   F(V64>=3,"Tốt",IF(V64=2,"Khá",IF(V64>=0,"Trung bình",IF(V64>=-2,"Kém","Rất kém"))))," (PP1 xếp "&IF(U64>=3,"Tốt",IF(U64=2,"Khá",IF(U64>=0,"Trung bình"
    #   ,IF(U64>=-2,"Kém","Rất kém"))))&", PP2 xếp "&IF(V64>=3,"Tốt",IF(V64=2,"Khá",IF(V64>=0,"Trung bình",IF(V64>=-2,"Kém","Rất kém"))))&" – "&IF(DOI_CHIEU!$
    #   C$33=0,"hai PP cùng trạng thái nhật chủ: kết luận lấy trung bình hai điểm, làm tròn xuống","kết luận theo PP"&DOI_CHIEU!$C$33&" – phương pháp được khu
    #   yến nghị về vượng suy ở DOI_CHIEU dòng 33")&")","")
    return X.cat(X.cat(X.cat(X.cat(X.UPPER(E.v(17,r,c-10)),': '),(X.CHOOSE(X.MATCH(E.v(17,r,c-10), X.Arr([['Rất kém', 'Kém', 'Trung bình', 'Khá', 'Tốt']]), 0.0), lambda: 'Hôn nhân bất lợi rõ: dễ ly tán, khắc vợ hoặc khó lập gia đình.', lambda: 'Hôn nhân kém: vợ yếu thế hoặc không hợp, hay bất hòa; nên kết hôn muộn.', lambda: 'Hôn nhân bình thường, tốt xấu tùy vận; cần vun đắp.', lambda: 'Vợ khá: hôn nhân tương đối ổn, có lúc trục trặc nhỏ.', lambda: 'Vợ tốt và có lực: hiền, đảm đang, giúp chồng; hôn nhân bền.') if X.truthy(X.eq(E.v(2,22,3),'Nam')) else X.CHOOSE(X.MATCH(E.v(17,r,c-10), X.Arr([['Rất kém', 'Kém', 'Trung bình', 'Khá', 'Tốt']]), 0.0), lambda: 'Hôn nhân bất lợi rõ: dễ ly tán, khắc chồng hoặc khó lập gia đình.', lambda: 'Hôn nhân kém: chồng yếu thế hoặc không hợp, hay bất hòa; nên kết hôn muộn.', lambda: 'Hôn nhân bình thường, tốt xấu tùy vận; cần vun đắp.', lambda: 'Chồng khá: hôn nhân tương đối ổn, có lúc trục trặc nhỏ.', lambda: 'Chồng tốt và có lực: chồng giỏi giang, giúp đỡ vợ; hôn nhân bền.'))),(X.cat(X.cat(X.cat(X.cat(' Khởi sắc trong đại vận dụng/hỷ thần: ',('không có' if X.truthy(X.eq(E.v(17,r,c-9),'')) else E.v(17,r,c-9))),'; kém đi trong đại vận kỵ thần: '),('không có' if X.truthy(X.eq(E.v(17,r,c-8),'')) else E.v(17,r,c-8))),'.') if X.truthy(X.eq(E.v(17,r,c-10),'Trung bình')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' (PP1 xếp ',('Tốt' if X.truthy(X.ge(E.v(17,r,c-13),3.0)) else ('Khá' if X.truthy(X.eq(E.v(17,r,c-13),2.0)) else ('Trung bình' if X.truthy(X.ge(E.v(17,r,c-13),0.0)) else ('Kém' if X.truthy(X.ge(E.v(17,r,c-13),X.neg(2.0))) else 'Rất kém'))))),', PP2 xếp '),('Tốt' if X.truthy(X.ge(E.v(17,r,c-12),3.0)) else ('Khá' if X.truthy(X.eq(E.v(17,r,c-12),2.0)) else ('Trung bình' if X.truthy(X.ge(E.v(17,r,c-12),0.0)) else ('Kém' if X.truthy(X.ge(E.v(17,r,c-12),X.neg(2.0))) else 'Rất kém'))))),' – '),('hai PP cùng trạng thái nhật chủ: kết luận lấy trung bình hai điểm, làm tròn xuống' if X.truthy(X.eq(E.v(7,33,3),0.0)) else X.cat(X.cat('kết luận theo PP',E.v(7,33,3)),' – phương pháp được khuyến nghị về vượng suy ở DOI_CHIEU dòng 33'))),')') if X.truthy(X.ne(('Tốt' if X.truthy(X.ge(E.v(17,r,c-13),3.0)) else ('Khá' if X.truthy(X.eq(E.v(17,r,c-13),2.0)) else ('Trung bình' if X.truthy(X.ge(E.v(17,r,c-13),0.0)) else ('Kém' if X.truthy(X.ge(E.v(17,r,c-13),X.neg(2.0))) else 'Rất kém')))),('Tốt' if X.truthy(X.ge(E.v(17,r,c-12),3.0)) else ('Khá' if X.truthy(X.eq(E.v(17,r,c-12),2.0)) else ('Trung bình' if X.truthy(X.ge(E.v(17,r,c-12),0.0)) else ('Kém' if X.truthy(X.ge(E.v(17,r,c-12),X.neg(2.0))) else 'Rất kém')))))) else ''))

def f183(E, r, c):
    # Ô AI64
    # = IF(AND(I72="Thuận",OR(J72="cả hai",J72=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I73="Thuận",OR(J73="cả hai",J73=IF(DOI_CHIEU!$C$33=2,"PP2","PP1
    #   "))),1,0)+IF(AND(I74="Thuận",OR(J74="cả hai",J74=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I75="Thuận",OR(J75="cả hai",J75=IF(DOI_CHIEU!$C$33=2,
    #   "PP2","PP1"))),1,0)+IF(AND(I76="Thuận",OR(J76="cả hai",J76=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I77="Thuận",OR(J77="cả hai",J77=IF(DOI_CHIE
    #   U!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I78="Thuận",OR(J78="cả hai",J78=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I79="Thuận",OR(J79="cả hai",J79=I
    #   F(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I80="Thuận",OR(J80="cả hai",J80=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I81="Thuận",OR(J81="cả 
    #   hai",J81=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I82="Thuận",OR(J82="cả hai",J82=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I83="Thuận",O
    #   R(J83="cả hai",J83=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I84="Thuận",OR(J84="cả hai",J84=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I85
    #   ="Thuận",OR(J85="cả hai",J85=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I86="Thuận",OR(J86="cả hai",J86=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+
    #   IF(AND(I87="Thuận",OR(J87="cả hai",J87=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I88="Thuận",OR(J88="cả hai",J88=IF(DOI_CHIEU!$C$33=2,"PP2","PP1
    #   "))),1,0)+IF(AND(I89="Thuận",OR(J89="cả hai",J89=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I90="Thuận",OR(J90="cả hai",J90=IF(DOI_CHIEU!$C$33=2,
    #   "PP2","PP1"))),1,0)+IF(AND(I91="Thuận",OR(J91="cả hai",J91=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I92="Thuận",OR(J92="cả hai",J92=IF(DOI_CHIE
    #   U!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I93="Thuận",OR(J93="cả hai",J93=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I94="Thuận",OR(J94="cả hai",J94=I
    #   F(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I95="Thuận",OR(J95="cả hai",J95=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I96="Thuận",OR(J96="cả 
    #   hai",J96=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I97="Thuận",OR(J97="cả hai",J97=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I98="Thuận",O
    #   R(J98="cả hai",J98=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I99="Thuận",OR(J99="cả hai",J99=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I10
    #   0="Thuận",OR(J100="cả hai",J100=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I101="Thuận",OR(J101="cả hai",J101=IF(DOI_CHIEU!$C$33=2,"PP2","PP1")))
    #   ,1,0)+IF(AND(I102="Thuận",OR(J102="cả hai",J102=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I103="Thuận",OR(J103="cả hai",J103=IF(DOI_CHIEU!$C$33=
    #   2,"PP2","PP1"))),1,0)+IF(AND(I104="Thuận",OR(J104="cả hai",J104=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I105="Thuận",OR(J105="cả hai",J105=IF(
    #   DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I106="Thuận",OR(J106="cả hai",J106=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)+IF(AND(I107="Thuận",OR(J107="
    #   cả hai",J107=IF(DOI_CHIEU!$C$33=2,"PP2","PP1"))),1,0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.eq(E.v(17,r+8,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+8,c-25),'cả hai'), X.eq(E.v(17,r+8,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+9,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+9,c-25),'cả hai'), X.eq(E.v(17,r+9,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+10,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+10,c-25),'cả hai'), X.eq(E.v(17,r+10,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+11,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+11,c-25),'cả hai'), X.eq(E.v(17,r+11,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+12,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+12,c-25),'cả hai'), X.eq(E.v(17,r+12,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+13,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+13,c-25),'cả hai'), X.eq(E.v(17,r+13,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+14,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+14,c-25),'cả hai'), X.eq(E.v(17,r+14,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+15,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+15,c-25),'cả hai'), X.eq(E.v(17,r+15,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+16,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+16,c-25),'cả hai'), X.eq(E.v(17,r+16,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+17,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+17,c-25),'cả hai'), X.eq(E.v(17,r+17,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+18,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+18,c-25),'cả hai'), X.eq(E.v(17,r+18,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+19,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+19,c-25),'cả hai'), X.eq(E.v(17,r+19,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+20,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+20,c-25),'cả hai'), X.eq(E.v(17,r+20,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+21,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+21,c-25),'cả hai'), X.eq(E.v(17,r+21,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+22,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+22,c-25),'cả hai'), X.eq(E.v(17,r+22,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+23,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+23,c-25),'cả hai'), X.eq(E.v(17,r+23,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+24,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+24,c-25),'cả hai'), X.eq(E.v(17,r+24,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+25,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+25,c-25),'cả hai'), X.eq(E.v(17,r+25,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+26,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+26,c-25),'cả hai'), X.eq(E.v(17,r+26,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+27,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+27,c-25),'cả hai'), X.eq(E.v(17,r+27,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+28,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+28,c-25),'cả hai'), X.eq(E.v(17,r+28,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+29,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+29,c-25),'cả hai'), X.eq(E.v(17,r+29,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+30,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+30,c-25),'cả hai'), X.eq(E.v(17,r+30,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+31,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+31,c-25),'cả hai'), X.eq(E.v(17,r+31,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+32,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+32,c-25),'cả hai'), X.eq(E.v(17,r+32,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+33,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+33,c-25),'cả hai'), X.eq(E.v(17,r+33,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+34,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+34,c-25),'cả hai'), X.eq(E.v(17,r+34,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+35,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+35,c-25),'cả hai'), X.eq(E.v(17,r+35,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+36,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+36,c-25),'cả hai'), X.eq(E.v(17,r+36,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+37,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+37,c-25),'cả hai'), X.eq(E.v(17,r+37,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+38,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+38,c-25),'cả hai'), X.eq(E.v(17,r+38,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+39,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+39,c-25),'cả hai'), X.eq(E.v(17,r+39,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+40,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+40,c-25),'cả hai'), X.eq(E.v(17,r+40,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+41,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+41,c-25),'cả hai'), X.eq(E.v(17,r+41,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+42,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+42,c-25),'cả hai'), X.eq(E.v(17,r+42,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0)),(1.0 if X.truthy(X.AND(X.eq(E.v(17,r+43,c-26),'Thuận'), X.OR(X.eq(E.v(17,r+43,c-25),'cả hai'), X.eq(E.v(17,r+43,c-25),('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1'))))) else 0.0))

def f184(E, r, c):
    # Ô AJ64
    # = TRIM(IF(L112="ĐỦ","• "&G112&" ["&B112&", "&H112&"] ","")&IF(L113="ĐỦ","• "&G113&" ["&B113&", "&H113&"] ","")&IF(L114="ĐỦ","• "&G114&" ["&B114&", "&H11
    #   4&"] ","")&IF(L115="ĐỦ","• "&G115&" ["&B115&", "&H115&"] ","")&IF(L117="ĐỦ","• "&G117&" ["&B117&", "&H117&"] ","")&IF(L118="ĐỦ","• "&G118&" ["&B118&",
    #    "&H118&"] ","")&IF(L120="ĐỦ","• "&G120&" ["&B120&", "&H120&"] ","")&IF(L121="ĐỦ","• "&G121&" ["&B121&", "&H121&"] ","")&IF(L123="ĐỦ","• "&G123&" ["&B
    #   123&", "&H123&"] ","")&IF(L125="ĐỦ","• "&G125&" ["&B125&", "&H125&"] ","")&IF(L126="ĐỦ","• "&G126&" ["&B126&", "&H126&"] ","")&IF(L128="ĐỦ","• "&G128&
    #   " ["&B128&", "&H128&"] ","")&IF(L130="ĐỦ","• "&G130&" ["&B130&", "&H130&"] ","")&IF(L132="ĐỦ","• "&G132&" ["&B132&", "&H132&"] ","")&IF(L133="ĐỦ","• "
    #   &G133&" ["&B133&", "&H133&"] ","")&IF(L134="ĐỦ","• "&G134&" ["&B134&", "&H134&"] ","")&IF(L136="ĐỦ","• "&G136&" ["&B136&", "&H136&"] ","")&IF(L137="ĐỦ
    #   ","• "&G137&" ["&B137&", "&H137&"] ","")&IF(L163="ĐỦ","• "&G163&" ["&B163&", "&H163&"] ","")&IF(L177="ĐỦ","• "&G177&" ["&B177&", "&H177&"] ","")&IF(L1
    #   79="ĐỦ","• "&G179&" ["&B179&", "&H179&"] ","")&IF(L180="ĐỦ","• "&G180&" ["&B180&", "&H180&"] ","")&IF(L181="ĐỦ","• "&G181&" ["&B181&", "&H181&"] ","")
    #   &IF(L182="ĐỦ","• "&G182&" ["&B182&", "&H182&"] ","")&IF(L183="ĐỦ","• "&G183&" ["&B183&", "&H183&"] ","")&IF(L184="ĐỦ","• "&G184&" ["&B184&", "&H184&"]
    #    ","")&IF(L185="ĐỦ","• "&G185&" ["&B185&", "&H185&"] ","")&IF(L187="ĐỦ","• "&G187&" ["&B187&", "&H187&"] ","")&IF(L188="ĐỦ","• "&G188&" ["&B188&", "&H
    #   188&"] ","")&IF(L189="ĐỦ","• "&G189&" ["&B189&", "&H189&"] ","")&IF(L190="ĐỦ","• "&G190&" ["&B190&", "&H190&"] ","")&IF(L191="ĐỦ","• "&G191&" ["&B191&
    #   ", "&H191&"] ","")&IF(L192="ĐỦ","• "&G192&" ["&B192&", "&H192&"] ","")&IF(L193="ĐỦ","• "&G193&" ["&B193&", "&H193&"] ","")&IF(L194="ĐỦ","• "&G194&" ["
    #   &B194&", "&H194&"] ","")&IF(L195="ĐỦ","• "&G195&" ["&B195&", "&H195&"] ","")&IF(L196="ĐỦ","• "&G196&" ["&B196&", "&H196&"] ","")&IF(L197="ĐỦ","• "&G19
    #   7&" ["&B197&", "&H197&"] ","")&IF(L198="ĐỦ","• "&G198&" ["&B198&", "&H198&"] ","")&IF(L199="ĐỦ","• "&G199&" ["&B199&", "&H199&"] ","")&IF(L200="ĐỦ","•
    #    "&G200&" ["&B200&", "&H200&"] ","")&IF(L201="ĐỦ","• "&G201&" ["&B201&", "&H201&"] ","")&IF(L202="ĐỦ","• "&G202&" ["&B202&", "&H202&"] ","")&IF(L204="
    #   ĐỦ","• "&G204&" ["&B204&", "&H204&"] ","")&IF(L205="ĐỦ","• "&G205&" ["&B205&", "&H205&"] ","")&IF(L206="ĐỦ","• "&G206&" ["&B206&", "&H206&"] ","")&IF(
    #   L208="ĐỦ","• "&G208&" ["&B208&", "&H208&"] ","")&IF(L209="ĐỦ","• "&G209&" ["&B209&", "&H209&"] ","")&IF(L212="ĐỦ","• "&G212&" ["&B212&", "&H212&"] ","
    #   ")&IF(L215="ĐỦ","• "&G215&" ["&B215&", "&H215&"] ","")&IF(L216="ĐỦ","• "&G216&" ["&B216&", "&H216&"] ","")&IF(L217="ĐỦ","• "&G217&" ["&B217&", "&H217&
    #   "] ","")&IF(L218="ĐỦ","• "&G218&" ["&B218&", "&H218&"] ","")&IF(L253="ĐỦ","• "&G253&" ["&B253&", "&H253&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+48,c-29)),' ['),E.v(17,r+48,c-34)),', '),E.v(17,r+48,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+48,c-24),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+49,c-29)),' ['),E.v(17,r+49,c-34)),', '),E.v(17,r+49,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+49,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+50,c-29)),' ['),E.v(17,r+50,c-34)),', '),E.v(17,r+50,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+50,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+51,c-29)),' ['),E.v(17,r+51,c-34)),', '),E.v(17,r+51,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+51,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+53,c-29)),' ['),E.v(17,r+53,c-34)),', '),E.v(17,r+53,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+53,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+54,c-29)),' ['),E.v(17,r+54,c-34)),', '),E.v(17,r+54,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+54,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+56,c-29)),' ['),E.v(17,r+56,c-34)),', '),E.v(17,r+56,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+56,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+57,c-29)),' ['),E.v(17,r+57,c-34)),', '),E.v(17,r+57,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+57,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+59,c-29)),' ['),E.v(17,r+59,c-34)),', '),E.v(17,r+59,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+59,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+61,c-29)),' ['),E.v(17,r+61,c-34)),', '),E.v(17,r+61,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+61,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+62,c-29)),' ['),E.v(17,r+62,c-34)),', '),E.v(17,r+62,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+62,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+64,c-29)),' ['),E.v(17,r+64,c-34)),', '),E.v(17,r+64,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+64,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+66,c-29)),' ['),E.v(17,r+66,c-34)),', '),E.v(17,r+66,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+66,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+68,c-29)),' ['),E.v(17,r+68,c-34)),', '),E.v(17,r+68,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+68,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+69,c-29)),' ['),E.v(17,r+69,c-34)),', '),E.v(17,r+69,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+69,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+70,c-29)),' ['),E.v(17,r+70,c-34)),', '),E.v(17,r+70,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+70,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+72,c-29)),' ['),E.v(17,r+72,c-34)),', '),E.v(17,r+72,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+72,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+73,c-29)),' ['),E.v(17,r+73,c-34)),', '),E.v(17,r+73,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+73,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+99,c-29)),' ['),E.v(17,r+99,c-34)),', '),E.v(17,r+99,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+99,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+113,c-29)),' ['),E.v(17,r+113,c-34)),', '),E.v(17,r+113,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+113,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+115,c-29)),' ['),E.v(17,r+115,c-34)),', '),E.v(17,r+115,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+115,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+116,c-29)),' ['),E.v(17,r+116,c-34)),', '),E.v(17,r+116,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+116,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+117,c-29)),' ['),E.v(17,r+117,c-34)),', '),E.v(17,r+117,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+117,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+118,c-29)),' ['),E.v(17,r+118,c-34)),', '),E.v(17,r+118,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+118,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+119,c-29)),' ['),E.v(17,r+119,c-34)),', '),E.v(17,r+119,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+119,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+120,c-29)),' ['),E.v(17,r+120,c-34)),', '),E.v(17,r+120,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+120,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+121,c-29)),' ['),E.v(17,r+121,c-34)),', '),E.v(17,r+121,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+121,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+123,c-29)),' ['),E.v(17,r+123,c-34)),', '),E.v(17,r+123,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+123,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+124,c-29)),' ['),E.v(17,r+124,c-34)),', '),E.v(17,r+124,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+124,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+125,c-29)),' ['),E.v(17,r+125,c-34)),', '),E.v(17,r+125,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+125,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+126,c-29)),' ['),E.v(17,r+126,c-34)),', '),E.v(17,r+126,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+126,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+127,c-29)),' ['),E.v(17,r+127,c-34)),', '),E.v(17,r+127,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+127,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+128,c-29)),' ['),E.v(17,r+128,c-34)),', '),E.v(17,r+128,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+128,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+129,c-29)),' ['),E.v(17,r+129,c-34)),', '),E.v(17,r+129,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+129,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+130,c-29)),' ['),E.v(17,r+130,c-34)),', '),E.v(17,r+130,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+130,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+131,c-29)),' ['),E.v(17,r+131,c-34)),', '),E.v(17,r+131,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+131,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+132,c-29)),' ['),E.v(17,r+132,c-34)),', '),E.v(17,r+132,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+132,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+133,c-29)),' ['),E.v(17,r+133,c-34)),', '),E.v(17,r+133,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+133,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+134,c-29)),' ['),E.v(17,r+134,c-34)),', '),E.v(17,r+134,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+134,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+135,c-29)),' ['),E.v(17,r+135,c-34)),', '),E.v(17,r+135,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+135,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+136,c-29)),' ['),E.v(17,r+136,c-34)),', '),E.v(17,r+136,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+136,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+137,c-29)),' ['),E.v(17,r+137,c-34)),', '),E.v(17,r+137,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+137,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+138,c-29)),' ['),E.v(17,r+138,c-34)),', '),E.v(17,r+138,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+138,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+140,c-29)),' ['),E.v(17,r+140,c-34)),', '),E.v(17,r+140,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+140,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+141,c-29)),' ['),E.v(17,r+141,c-34)),', '),E.v(17,r+141,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+141,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+142,c-29)),' ['),E.v(17,r+142,c-34)),', '),E.v(17,r+142,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+142,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+144,c-29)),' ['),E.v(17,r+144,c-34)),', '),E.v(17,r+144,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+144,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+145,c-29)),' ['),E.v(17,r+145,c-34)),', '),E.v(17,r+145,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+145,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+148,c-29)),' ['),E.v(17,r+148,c-34)),', '),E.v(17,r+148,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+148,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+151,c-29)),' ['),E.v(17,r+151,c-34)),', '),E.v(17,r+151,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+151,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+152,c-29)),' ['),E.v(17,r+152,c-34)),', '),E.v(17,r+152,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+152,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+153,c-29)),' ['),E.v(17,r+153,c-34)),', '),E.v(17,r+153,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+153,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+154,c-29)),' ['),E.v(17,r+154,c-34)),', '),E.v(17,r+154,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+154,c-24),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+189,c-29)),' ['),E.v(17,r+189,c-34)),', '),E.v(17,r+189,c-28)),'] ') if X.truthy(X.eq(E.v(17,r+189,c-24),'ĐỦ')) else '')))

def f185(E, r, c):
    # Ô AK64
    # = TRIM(IF(L138="ĐỦ","• "&G138&" ["&B138&", "&H138&"] ","")&IF(L139="ĐỦ","• "&G139&" ["&B139&", "&H139&"] ","")&IF(L140="ĐỦ","• "&G140&" ["&B140&", "&H14
    #   0&"] ","")&IF(L141="ĐỦ","• "&G141&" ["&B141&", "&H141&"] ","")&IF(L142="ĐỦ","• "&G142&" ["&B142&", "&H142&"] ","")&IF(L143="ĐỦ","• "&G143&" ["&B143&",
    #    "&H143&"] ","")&IF(L144="ĐỦ","• "&G144&" ["&B144&", "&H144&"] ","")&IF(L145="ĐỦ","• "&G145&" ["&B145&", "&H145&"] ","")&IF(L146="ĐỦ","• "&G146&" ["&B
    #   146&", "&H146&"] ","")&IF(L147="ĐỦ","• "&G147&" ["&B147&", "&H147&"] ","")&IF(L148="ĐỦ","• "&G148&" ["&B148&", "&H148&"] ","")&IF(L149="ĐỦ","• "&G149&
    #   " ["&B149&", "&H149&"] ","")&IF(L150="ĐỦ","• "&G150&" ["&B150&", "&H150&"] ","")&IF(L151="ĐỦ","• "&G151&" ["&B151&", "&H151&"] ","")&IF(L152="ĐỦ","• "
    #   &G152&" ["&B152&", "&H152&"] ","")&IF(L153="ĐỦ","• "&G153&" ["&B153&", "&H153&"] ","")&IF(L154="ĐỦ","• "&G154&" ["&B154&", "&H154&"] ","")&IF(L155="ĐỦ
    #   ","• "&G155&" ["&B155&", "&H155&"] ","")&IF(L156="ĐỦ","• "&G156&" ["&B156&", "&H156&"] ","")&IF(L157="ĐỦ","• "&G157&" ["&B157&", "&H157&"] ","")&IF(L1
    #   58="ĐỦ","• "&G158&" ["&B158&", "&H158&"] ","")&IF(L159="ĐỦ","• "&G159&" ["&B159&", "&H159&"] ","")&IF(L162="ĐỦ","• "&G162&" ["&B162&", "&H162&"] ","")
    #   &IF(L164="ĐỦ","• "&G164&" ["&B164&", "&H164&"] ","")&IF(L165="ĐỦ","• "&G165&" ["&B165&", "&H165&"] ","")&IF(L166="ĐỦ","• "&G166&" ["&B166&", "&H166&"]
    #    ","")&IF(L167="ĐỦ","• "&G167&" ["&B167&", "&H167&"] ","")&IF(L168="ĐỦ","• "&G168&" ["&B168&", "&H168&"] ","")&IF(L169="ĐỦ","• "&G169&" ["&B169&", "&H
    #   169&"] ","")&IF(L172="ĐỦ","• "&G172&" ["&B172&", "&H172&"] ","")&IF(L173="ĐỦ","• "&G173&" ["&B173&", "&H173&"] ","")&IF(L174="ĐỦ","• "&G174&" ["&B174&
    #   ", "&H174&"] ","")&IF(L175="ĐỦ","• "&G175&" ["&B175&", "&H175&"] ","")&IF(L211="ĐỦ","• "&G211&" ["&B211&", "&H211&"] ","")&IF(L213="ĐỦ","• "&G213&" ["
    #   &B213&", "&H213&"] ","")&IF(L214="ĐỦ","• "&G214&" ["&B214&", "&H214&"] ","")&IF(L220="ĐỦ","• "&G220&" ["&B220&", "&H220&"] ","")&IF(L221="ĐỦ","• "&G22
    #   1&" ["&B221&", "&H221&"] ","")&IF(L222="ĐỦ","• "&G222&" ["&B222&", "&H222&"] ","")&IF(L223="ĐỦ","• "&G223&" ["&B223&", "&H223&"] ","")&IF(L224="ĐỦ","•
    #    "&G224&" ["&B224&", "&H224&"] ","")&IF(L225="ĐỦ","• "&G225&" ["&B225&", "&H225&"] ","")&IF(L227="ĐỦ","• "&G227&" ["&B227&", "&H227&"] ","")&IF(L228="
    #   ĐỦ","• "&G228&" ["&B228&", "&H228&"] ","")&IF(L229="ĐỦ","• "&G229&" ["&B229&", "&H229&"] ","")&IF(L230="ĐỦ","• "&G230&" ["&B230&", "&H230&"] ","")&IF(
    #   L231="ĐỦ","• "&G231&" ["&B231&", "&H231&"] ","")&IF(L232="ĐỦ","• "&G232&" ["&B232&", "&H232&"] ","")&IF(L233="ĐỦ","• "&G233&" ["&B233&", "&H233&"] ","
    #   ")&IF(L234="ĐỦ","• "&G234&" ["&B234&", "&H234&"] ","")&IF(L235="ĐỦ","• "&G235&" ["&B235&", "&H235&"] ","")&IF(L236="ĐỦ","• "&G236&" ["&B236&", "&H236&
    #   "] ","")&IF(L237="ĐỦ","• "&G237&" ["&B237&", "&H237&"] ","")&IF(L240="ĐỦ","• "&G240&" ["&B240&", "&H240&"] ","")&IF(L241="ĐỦ","• "&G241&" ["&B241&", "
    #   &H241&"] ","")&IF(L242="ĐỦ","• "&G242&" ["&B242&", "&H242&"] ","")&IF(L245="ĐỦ","• "&G245&" ["&B245&", "&H245&"] ","")&IF(L246="ĐỦ","• "&G246&" ["&B24
    #   6&", "&H246&"] ","")&IF(L247="ĐỦ","• "&G247&" ["&B247&", "&H247&"] ","")&IF(L249="ĐỦ","• "&G249&" ["&B249&", "&H249&"] ","")&IF(L250="ĐỦ","• "&G250&" 
    #   ["&B250&", "&H250&"] ","")&IF(L252="ĐỦ","• "&G252&" ["&B252&", "&H252&"] ","")&IF(L254="ĐỦ","• "&G254&" ["&B254&", "&H254&"] ","")&IF(L255="ĐỦ","• "&G
    #   255&" ["&B255&", "&H255&"] ","")&IF(L257="ĐỦ","• "&G257&" ["&B257&", "&H257&"] ","")&IF(L258="ĐỦ","• "&G258&" ["&B258&", "&H258&"] ","")&IF(L259="ĐỦ",
    #   "• "&G259&" ["&B259&", "&H259&"] ","")&IF(L260="ĐỦ","• "&G260&" ["&B260&", "&H260&"] ","")&IF(L261="ĐỦ","• "&G261&" ["&B261&", "&H261&"] ","")&IF(L262
    #   ="ĐỦ","• "&G262&" ["&B262&", "&H262&"] ","")&IF(L263="ĐỦ","• "&G263&" ["&B263&", "&H263&"] ","")&IF(L264="ĐỦ","• "&G264&" ["&B264&", "&H264&"] ","")&I
    #   F(L265="ĐỦ","• "&G265&" ["&B265&", "&H265&"] ","")&IF(L266="ĐỦ","• "&G266&" ["&B266&", "&H266&"] ","")&IF(L267="ĐỦ","• "&G267&" ["&B267&", "&H267&"] "
    #   ,"")&IF(L268="ĐỦ","• "&G268&" ["&B268&", "&H268&"] ","")&IF(L269="ĐỦ","• "&G269&" ["&B269&", "&H269&"] ","")&IF(L270="ĐỦ","• "&G270&" ["&B270&", "&H27
    #   0&"] ","")&IF(L271="ĐỦ","• "&G271&" ["&B271&", "&H271&"] ","")&IF(L272="ĐỦ","• "&G272&" ["&B272&", "&H272&"] ","")&IF(L273="ĐỦ","• "&G273&" ["&B273&",
    #    "&H273&"] ","")&IF(L274="ĐỦ","• "&G274&" ["&B274&", "&H274&"] ","")&IF(L275="ĐỦ","• "&G275&" ["&B275&", "&H275&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+74,c-30)),' ['),E.v(17,r+74,c-35)),', '),E.v(17,r+74,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+74,c-25),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+75,c-30)),' ['),E.v(17,r+75,c-35)),', '),E.v(17,r+75,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+75,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+76,c-30)),' ['),E.v(17,r+76,c-35)),', '),E.v(17,r+76,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+76,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+77,c-30)),' ['),E.v(17,r+77,c-35)),', '),E.v(17,r+77,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+77,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+78,c-30)),' ['),E.v(17,r+78,c-35)),', '),E.v(17,r+78,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+78,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+79,c-30)),' ['),E.v(17,r+79,c-35)),', '),E.v(17,r+79,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+79,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+80,c-30)),' ['),E.v(17,r+80,c-35)),', '),E.v(17,r+80,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+80,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+81,c-30)),' ['),E.v(17,r+81,c-35)),', '),E.v(17,r+81,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+81,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+82,c-30)),' ['),E.v(17,r+82,c-35)),', '),E.v(17,r+82,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+82,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+83,c-30)),' ['),E.v(17,r+83,c-35)),', '),E.v(17,r+83,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+83,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+84,c-30)),' ['),E.v(17,r+84,c-35)),', '),E.v(17,r+84,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+84,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+85,c-30)),' ['),E.v(17,r+85,c-35)),', '),E.v(17,r+85,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+85,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+86,c-30)),' ['),E.v(17,r+86,c-35)),', '),E.v(17,r+86,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+86,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+87,c-30)),' ['),E.v(17,r+87,c-35)),', '),E.v(17,r+87,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+87,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+88,c-30)),' ['),E.v(17,r+88,c-35)),', '),E.v(17,r+88,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+88,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+89,c-30)),' ['),E.v(17,r+89,c-35)),', '),E.v(17,r+89,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+89,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+90,c-30)),' ['),E.v(17,r+90,c-35)),', '),E.v(17,r+90,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+90,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+91,c-30)),' ['),E.v(17,r+91,c-35)),', '),E.v(17,r+91,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+91,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+92,c-30)),' ['),E.v(17,r+92,c-35)),', '),E.v(17,r+92,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+92,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+93,c-30)),' ['),E.v(17,r+93,c-35)),', '),E.v(17,r+93,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+93,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+94,c-30)),' ['),E.v(17,r+94,c-35)),', '),E.v(17,r+94,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+94,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+95,c-30)),' ['),E.v(17,r+95,c-35)),', '),E.v(17,r+95,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+95,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+98,c-30)),' ['),E.v(17,r+98,c-35)),', '),E.v(17,r+98,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+98,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+100,c-30)),' ['),E.v(17,r+100,c-35)),', '),E.v(17,r+100,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+100,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+101,c-30)),' ['),E.v(17,r+101,c-35)),', '),E.v(17,r+101,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+101,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+102,c-30)),' ['),E.v(17,r+102,c-35)),', '),E.v(17,r+102,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+102,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+103,c-30)),' ['),E.v(17,r+103,c-35)),', '),E.v(17,r+103,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+103,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+104,c-30)),' ['),E.v(17,r+104,c-35)),', '),E.v(17,r+104,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+104,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+105,c-30)),' ['),E.v(17,r+105,c-35)),', '),E.v(17,r+105,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+105,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+108,c-30)),' ['),E.v(17,r+108,c-35)),', '),E.v(17,r+108,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+108,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+109,c-30)),' ['),E.v(17,r+109,c-35)),', '),E.v(17,r+109,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+109,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+110,c-30)),' ['),E.v(17,r+110,c-35)),', '),E.v(17,r+110,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+110,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+111,c-30)),' ['),E.v(17,r+111,c-35)),', '),E.v(17,r+111,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+111,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+147,c-30)),' ['),E.v(17,r+147,c-35)),', '),E.v(17,r+147,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+147,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+149,c-30)),' ['),E.v(17,r+149,c-35)),', '),E.v(17,r+149,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+149,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+150,c-30)),' ['),E.v(17,r+150,c-35)),', '),E.v(17,r+150,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+150,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+156,c-30)),' ['),E.v(17,r+156,c-35)),', '),E.v(17,r+156,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+156,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+157,c-30)),' ['),E.v(17,r+157,c-35)),', '),E.v(17,r+157,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+157,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+158,c-30)),' ['),E.v(17,r+158,c-35)),', '),E.v(17,r+158,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+158,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+159,c-30)),' ['),E.v(17,r+159,c-35)),', '),E.v(17,r+159,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+159,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+160,c-30)),' ['),E.v(17,r+160,c-35)),', '),E.v(17,r+160,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+160,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+161,c-30)),' ['),E.v(17,r+161,c-35)),', '),E.v(17,r+161,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+161,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+163,c-30)),' ['),E.v(17,r+163,c-35)),', '),E.v(17,r+163,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+163,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+164,c-30)),' ['),E.v(17,r+164,c-35)),', '),E.v(17,r+164,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+164,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+165,c-30)),' ['),E.v(17,r+165,c-35)),', '),E.v(17,r+165,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+165,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+166,c-30)),' ['),E.v(17,r+166,c-35)),', '),E.v(17,r+166,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+166,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+167,c-30)),' ['),E.v(17,r+167,c-35)),', '),E.v(17,r+167,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+167,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+168,c-30)),' ['),E.v(17,r+168,c-35)),', '),E.v(17,r+168,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+168,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+169,c-30)),' ['),E.v(17,r+169,c-35)),', '),E.v(17,r+169,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+169,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+170,c-30)),' ['),E.v(17,r+170,c-35)),', '),E.v(17,r+170,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+170,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+171,c-30)),' ['),E.v(17,r+171,c-35)),', '),E.v(17,r+171,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+171,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+172,c-30)),' ['),E.v(17,r+172,c-35)),', '),E.v(17,r+172,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+172,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+173,c-30)),' ['),E.v(17,r+173,c-35)),', '),E.v(17,r+173,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+173,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+176,c-30)),' ['),E.v(17,r+176,c-35)),', '),E.v(17,r+176,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+176,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+177,c-30)),' ['),E.v(17,r+177,c-35)),', '),E.v(17,r+177,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+177,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+178,c-30)),' ['),E.v(17,r+178,c-35)),', '),E.v(17,r+178,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+178,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+181,c-30)),' ['),E.v(17,r+181,c-35)),', '),E.v(17,r+181,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+181,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+182,c-30)),' ['),E.v(17,r+182,c-35)),', '),E.v(17,r+182,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+182,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+183,c-30)),' ['),E.v(17,r+183,c-35)),', '),E.v(17,r+183,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+183,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+185,c-30)),' ['),E.v(17,r+185,c-35)),', '),E.v(17,r+185,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+185,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+186,c-30)),' ['),E.v(17,r+186,c-35)),', '),E.v(17,r+186,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+186,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+188,c-30)),' ['),E.v(17,r+188,c-35)),', '),E.v(17,r+188,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+188,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+190,c-30)),' ['),E.v(17,r+190,c-35)),', '),E.v(17,r+190,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+190,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+191,c-30)),' ['),E.v(17,r+191,c-35)),', '),E.v(17,r+191,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+191,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+193,c-30)),' ['),E.v(17,r+193,c-35)),', '),E.v(17,r+193,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+193,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+194,c-30)),' ['),E.v(17,r+194,c-35)),', '),E.v(17,r+194,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+194,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+195,c-30)),' ['),E.v(17,r+195,c-35)),', '),E.v(17,r+195,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+195,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+196,c-30)),' ['),E.v(17,r+196,c-35)),', '),E.v(17,r+196,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+196,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+197,c-30)),' ['),E.v(17,r+197,c-35)),', '),E.v(17,r+197,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+197,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+198,c-30)),' ['),E.v(17,r+198,c-35)),', '),E.v(17,r+198,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+198,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+199,c-30)),' ['),E.v(17,r+199,c-35)),', '),E.v(17,r+199,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+199,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+200,c-30)),' ['),E.v(17,r+200,c-35)),', '),E.v(17,r+200,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+200,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+201,c-30)),' ['),E.v(17,r+201,c-35)),', '),E.v(17,r+201,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+201,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+202,c-30)),' ['),E.v(17,r+202,c-35)),', '),E.v(17,r+202,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+202,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+203,c-30)),' ['),E.v(17,r+203,c-35)),', '),E.v(17,r+203,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+203,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+204,c-30)),' ['),E.v(17,r+204,c-35)),', '),E.v(17,r+204,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+204,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+205,c-30)),' ['),E.v(17,r+205,c-35)),', '),E.v(17,r+205,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+205,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+206,c-30)),' ['),E.v(17,r+206,c-35)),', '),E.v(17,r+206,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+206,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+207,c-30)),' ['),E.v(17,r+207,c-35)),', '),E.v(17,r+207,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+207,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+208,c-30)),' ['),E.v(17,r+208,c-35)),', '),E.v(17,r+208,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+208,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+209,c-30)),' ['),E.v(17,r+209,c-35)),', '),E.v(17,r+209,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+209,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+210,c-30)),' ['),E.v(17,r+210,c-35)),', '),E.v(17,r+210,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+210,c-25),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+211,c-30)),' ['),E.v(17,r+211,c-35)),', '),E.v(17,r+211,c-29)),'] ') if X.truthy(X.eq(E.v(17,r+211,c-25),'ĐỦ')) else '')))

def f186(E, r, c):
    # Ô AL64
    # = TRIM(IF(L116="ĐỦ","• "&G116&" ["&B116&", "&H116&"] ","")&IF(L119="ĐỦ","• "&G119&" ["&B119&", "&H119&"] ","")&IF(L122="ĐỦ","• "&G122&" ["&B122&", "&H12
    #   2&"] ","")&IF(L124="ĐỦ","• "&G124&" ["&B124&", "&H124&"] ","")&IF(L127="ĐỦ","• "&G127&" ["&B127&", "&H127&"] ","")&IF(L129="ĐỦ","• "&G129&" ["&B129&",
    #    "&H129&"] ","")&IF(L131="ĐỦ","• "&G131&" ["&B131&", "&H131&"] ","")&IF(L135="ĐỦ","• "&G135&" ["&B135&", "&H135&"] ","")&IF(L160="ĐỦ","• "&G160&" ["&B
    #   160&", "&H160&"] ","")&IF(L161="ĐỦ","• "&G161&" ["&B161&", "&H161&"] ","")&IF(L170="ĐỦ","• "&G170&" ["&B170&", "&H170&"] ","")&IF(L171="ĐỦ","• "&G171&
    #   " ["&B171&", "&H171&"] ","")&IF(L176="ĐỦ","• "&G176&" ["&B176&", "&H176&"] ","")&IF(L178="ĐỦ","• "&G178&" ["&B178&", "&H178&"] ","")&IF(L186="ĐỦ","• "
    #   &G186&" ["&B186&", "&H186&"] ","")&IF(L203="ĐỦ","• "&G203&" ["&B203&", "&H203&"] ","")&IF(L207="ĐỦ","• "&G207&" ["&B207&", "&H207&"] ","")&IF(L210="ĐỦ
    #   ","• "&G210&" ["&B210&", "&H210&"] ","")&IF(L219="ĐỦ","• "&G219&" ["&B219&", "&H219&"] ","")&IF(L226="ĐỦ","• "&G226&" ["&B226&", "&H226&"] ","")&IF(L2
    #   38="ĐỦ","• "&G238&" ["&B238&", "&H238&"] ","")&IF(L239="ĐỦ","• "&G239&" ["&B239&", "&H239&"] ","")&IF(L243="ĐỦ","• "&G243&" ["&B243&", "&H243&"] ","")
    #   &IF(L244="ĐỦ","• "&G244&" ["&B244&", "&H244&"] ","")&IF(L248="ĐỦ","• "&G248&" ["&B248&", "&H248&"] ","")&IF(L251="ĐỦ","• "&G251&" ["&B251&", "&H251&"]
    #    ","")&IF(L256="ĐỦ","• "&G256&" ["&B256&", "&H256&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+52,c-31)),' ['),E.v(17,r+52,c-36)),', '),E.v(17,r+52,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+52,c-26),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+55,c-31)),' ['),E.v(17,r+55,c-36)),', '),E.v(17,r+55,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+55,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+58,c-31)),' ['),E.v(17,r+58,c-36)),', '),E.v(17,r+58,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+58,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+60,c-31)),' ['),E.v(17,r+60,c-36)),', '),E.v(17,r+60,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+60,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+63,c-31)),' ['),E.v(17,r+63,c-36)),', '),E.v(17,r+63,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+63,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+65,c-31)),' ['),E.v(17,r+65,c-36)),', '),E.v(17,r+65,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+65,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+67,c-31)),' ['),E.v(17,r+67,c-36)),', '),E.v(17,r+67,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+67,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+71,c-31)),' ['),E.v(17,r+71,c-36)),', '),E.v(17,r+71,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+71,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+96,c-31)),' ['),E.v(17,r+96,c-36)),', '),E.v(17,r+96,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+96,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+97,c-31)),' ['),E.v(17,r+97,c-36)),', '),E.v(17,r+97,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+97,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+106,c-31)),' ['),E.v(17,r+106,c-36)),', '),E.v(17,r+106,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+106,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+107,c-31)),' ['),E.v(17,r+107,c-36)),', '),E.v(17,r+107,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+107,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+112,c-31)),' ['),E.v(17,r+112,c-36)),', '),E.v(17,r+112,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+112,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+114,c-31)),' ['),E.v(17,r+114,c-36)),', '),E.v(17,r+114,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+114,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+122,c-31)),' ['),E.v(17,r+122,c-36)),', '),E.v(17,r+122,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+122,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+139,c-31)),' ['),E.v(17,r+139,c-36)),', '),E.v(17,r+139,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+139,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+143,c-31)),' ['),E.v(17,r+143,c-36)),', '),E.v(17,r+143,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+143,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+146,c-31)),' ['),E.v(17,r+146,c-36)),', '),E.v(17,r+146,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+146,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+155,c-31)),' ['),E.v(17,r+155,c-36)),', '),E.v(17,r+155,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+155,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+162,c-31)),' ['),E.v(17,r+162,c-36)),', '),E.v(17,r+162,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+162,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+174,c-31)),' ['),E.v(17,r+174,c-36)),', '),E.v(17,r+174,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+174,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+175,c-31)),' ['),E.v(17,r+175,c-36)),', '),E.v(17,r+175,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+175,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+179,c-31)),' ['),E.v(17,r+179,c-36)),', '),E.v(17,r+179,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+179,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+180,c-31)),' ['),E.v(17,r+180,c-36)),', '),E.v(17,r+180,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+180,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+184,c-31)),' ['),E.v(17,r+184,c-36)),', '),E.v(17,r+184,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+184,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+187,c-31)),' ['),E.v(17,r+187,c-36)),', '),E.v(17,r+187,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+187,c-26),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(17,r+192,c-31)),' ['),E.v(17,r+192,c-36)),', '),E.v(17,r+192,c-30)),'] ') if X.truthy(X.eq(E.v(17,r+192,c-26),'ĐỦ')) else '')))

def f187(E, r, c):
    # Ô AM64
    # = TRIM(IF(AJ64="","","Thuận: "&AJ64&" ")&IF(AK64="","","Bất lợi: "&AK64&" ")&IF(AL64="","","Đặc điểm: "&AL64))
    return X.TRIM(X.cat(X.cat(('' if X.truthy(X.eq(E.v(17,r,c-3),'')) else X.cat(X.cat('Thuận: ',E.v(17,r,c-3)),' ')),('' if X.truthy(X.eq(E.v(17,r,c-2),'')) else X.cat(X.cat('Bất lợi: ',E.v(17,r,c-2)),' '))),('' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat('Đặc điểm: ',E.v(17,r,c-1)))))

def f188(E, r, c):
    # Ô D66
    # = IF(LUU_NIEN!$B$5<=18,"(Lưu ý tuổi: chưa thành niên – dòng này là đặc điểm vận hôn nhân ở tương lai xa, chưa xét lúc này.) ","")&IF(LUU_NIEN!$B$5>=40,"
    #   (Lưu ý tuổi: đã qua tuổi kết hôn thông thường – đọc dòng này như đặc điểm vận hôn nhân, không phải dự báo ngày kết hôn.) ","")&IF(HON_NHAN!$C$22="","K
    #   hông có năm thuận lợi trong danh sách","Năm thuận lợi gần nhất: "&HON_NHAN!$C$22&" ("&HON_NHAN!$G$22&")"&IF(HON_NHAN!$K$22<>""," – lưu ý cùng năm có d
    #   ấu hiệu bất lợi: "&HON_NHAN!$K$22,""))&IF(COUNTIFS(DAI_VAN!$C$86:$J$86,1,DAI_VAN!$C$36:$J$36,">=9")>0,CHAR(10)&"Tránh kết hôn trong đại vận xung cung 
    #   phối ngẫu: "&TRIM(IF(AND(DAI_VAN!C86=1,DAI_VAN!C36>=9),DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(AND(DAI_VAN!D86=1,DAI_VAN!D36
    #   >=9),DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(AND(DAI_VAN!E86=1,DAI_VAN!E36>=9),DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36
    #   +9)&" tuổi) ","")&IF(AND(DAI_VAN!F86=1,DAI_VAN!F36>=9),DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi) ","")&IF(AND(DAI_VAN!G86=1,DAI_VAN!G36
    #   >=9),DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(AND(DAI_VAN!H86=1,DAI_VAN!H36>=9),DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36
    #   +9)&" tuổi) ","")&IF(AND(DAI_VAN!I86=1,DAI_VAN!I36>=9),DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi) ","")&IF(AND(DAI_VAN!J86=1,DAI_VAN!J36
    #   >=9),DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))&" – ""chờ qua vận đó mới lấy vợ thì tránh được"" (P1 tr.156)"&IF(NHAP_LIEU!$C$22=
    #   "Nữ"," – với nữ: lấy chồng","")&". Chỉ liệt kê vận có tuổi kết thúc từ 18 trở lên.","")
    return X.cat(X.cat(X.cat(('(Lưu ý tuổi: chưa thành niên – dòng này là đặc điểm vận hôn nhân ở tương lai xa, chưa xét lúc này.) ' if X.truthy(X.le(E.v(15,5,2),18.0)) else ''),('(Lưu ý tuổi: đã qua tuổi kết hôn thông thường – đọc dòng này như đặc điểm vận hôn nhân, không phải dự báo ngày kết hôn.) ' if X.truthy(X.ge(E.v(15,5,2),40.0)) else '')),('Không có năm thuận lợi trong danh sách' if X.truthy(X.eq(E.v(17,22,3),'')) else X.cat(X.cat(X.cat(X.cat(X.cat('Năm thuận lợi gần nhất: ',E.v(17,22,3)),' ('),E.v(17,22,7)),')'),(X.cat(' – lưu ý cùng năm có dấu hiệu bất lợi: ',E.v(17,22,11)) if X.truthy(X.ne(E.v(17,22,11),'')) else '')))),(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'Tránh kết hôn trong đại vận xung cung phối ngẫu: '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c-1),' ('),E.v(14,r-30,c-1)),'–'),X.add(E.v(14,r-30,c-1),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c-1),1.0), X.ge(E.v(14,r-30,c-1),9.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c),' ('),E.v(14,r-30,c)),'–'),X.add(E.v(14,r-30,c),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c),1.0), X.ge(E.v(14,r-30,c),9.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c+1),' ('),E.v(14,r-30,c+1)),'–'),X.add(E.v(14,r-30,c+1),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c+1),1.0), X.ge(E.v(14,r-30,c+1),9.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c+2),' ('),E.v(14,r-30,c+2)),'–'),X.add(E.v(14,r-30,c+2),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c+2),1.0), X.ge(E.v(14,r-30,c+2),9.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c+3),' ('),E.v(14,r-30,c+3)),'–'),X.add(E.v(14,r-30,c+3),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c+3),1.0), X.ge(E.v(14,r-30,c+3),9.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c+4),' ('),E.v(14,r-30,c+4)),'–'),X.add(E.v(14,r-30,c+4),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c+4),1.0), X.ge(E.v(14,r-30,c+4),9.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c+5),' ('),E.v(14,r-30,c+5)),'–'),X.add(E.v(14,r-30,c+5),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c+5),1.0), X.ge(E.v(14,r-30,c+5),9.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-31,c+6),' ('),E.v(14,r-30,c+6)),'–'),X.add(E.v(14,r-30,c+6),9.0)),' tuổi) ') if X.truthy(X.AND(X.eq(E.v(14,r+20,c+6),1.0), X.ge(E.v(14,r-30,c+6),9.0))) else '')))),' – "chờ qua vận đó mới lấy vợ thì tránh được" (P1 tr.156)'),(' – với nữ: lấy chồng' if X.truthy(X.eq(E.v(2,22,3),'Nữ')) else '')),'. Chỉ liệt kê vận có tuổi kết thúc từ 18 trở lên.') if X.truthy(X.gt(X.COUNTIFS(E.rg(14,86,3,86,10), 1.0, E.rg(14,36,3,36,10), '>=9'),0.0)) else ''))

def f189(E, r, c):
    # Ô D67
    # = HON_NHAN!$D$59&CHAR(10)&"Phương hướng kết hôn: "&HON_NHAN!$D$48
    return X.cat(X.cat(X.cat(E.v(17,59,4),X.CHAR(10.0)),'Phương hướng kết hôn: '),E.v(17,48,4))

def f190(E, r, c):
    # Ô E72, F72
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)
    #   ),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0))) else '—')

def f191(E, r, c):
    # Ô J72, J73, J74, J75, J76, J77, J78, J79, J80, J81, J82, J83 … (203 ô)
    # = IF(AND(E72="ĐỦ",F72="ĐỦ"),"cả hai",IF(E72="ĐỦ","PP1",IF(F72="ĐỦ","PP2","")))
    return ('cả hai' if X.truthy(X.AND(X.eq(E.v(17,r,c-5),'ĐỦ'), X.eq(E.v(17,r,c-4),'ĐỦ'))) else ('PP1' if X.truthy(X.eq(E.v(17,r,c-5),'ĐỦ')) else ('PP2' if X.truthy(X.eq(E.v(17,r,c-4),'ĐỦ')) else '')))

def f192(E, r, c):
    # Ô K72
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính quan bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.154; P2 tr.464): Nam: chi ngày là chính quan → "&IF(J7
    #   2="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J72&") ⇒ "&G72&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P1 tr.154; P2 tr.464): Nam: chi ngày là chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f193(E, r, c):
    # Ô E73, F73
    # = IF(OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),NOT(((COUNTIFS(SAO!$F
    #   $5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0))),(SAO!$Q$14=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))), X.eq(E.v(13,14,17),1.0))) else '—')

def f194(E, r, c):
    # Ô K73
    # = "Dữ kiện: "&"Lộc ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$Q$14=1,"có","không")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thực thần bản khí ở ngày: "&COUNTIFS(SAO
    #   !$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P
    #   1 tr.154): Chi ngày là thực thần không bị kiêu đoạt, hoặc chi ngày là lộc → "&IF(J73="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J73&") ⇒ "&G73&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Lộc ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,17),1.0)) else 'không')),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.154): Chi ngày là thực thần không bị kiêu đoạt, hoặc chi ngày là lộc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f195(E, r, c):
    # Ô E74
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+
    #   COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),OR(AND(DUNG_HY_KY!$C$109>=0,MOD(
    #   DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$C$110>=0,MOD(DUNG_HY_KY!$C$110-TRUNG_GIAN!$E$7,5)=2))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.OR(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,110,3),0.0), X.eq(X.MOD(X.sub(E.v(8,110,3),E.v(23,7,5)), 5.0),2.0))))) else '—')

def f196(E, r, c):
    # Ô F74
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+
    #   COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),OR(AND(DUNG_HY_KY!$D$109>=0,MOD(
    #   DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$D$110>=0,MOD(DUNG_HY_KY!$D$110-TRUNG_GIAN!$E$7,5)=2))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.OR(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,110,4),0.0), X.eq(X.MOD(X.sub(E.v(8,110,4),E.v(23,7,5)), 5.0),2.0))))) else '—')

def f197(E, r, c):
    # Ô K74
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUN
    #   G_HY_KY!$D$110+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+
    #   1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,S
    #   AO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thiên tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$
    #   5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.154; P2 tr.463): Nam: chi ngày là tài và tài là hỷ/dụng → "&IF(J74="","chưa đủ đi
    #   ều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J74&") ⇒ "&G74&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P1 tr.154; P2 tr.463): Nam: chi ngày là tài và tài là hỷ/dụng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f198(E, r, c):
    # Ô E75, F75
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '—')

def f199(E, r, c):
    # Ô K75
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính ấn bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản k
    #   hí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.154; P2 tr.463): Chi ngày là chính ấn → "&IF(J75="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J75&") ⇒ "&G75&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P1 tr.154; P2 tr.463): Chi ngày là chính ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f200(E, r, c):
    # Ô E76
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),(DUNG_HY_KY!$C$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0), X.eq(E.v(8,7,3),0.0))) else '—')

def f201(E, r, c):
    # Ô F76
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),(DUNG_HY_KY!$D$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0), X.eq(E.v(8,7,4),0.0))) else '—')

def f202(E, r, c):
    # Ô K76
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Tỉ kiên bản khí ở ngày: "&COUNTIFS(SAO!$F$
    #   5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Kiếp tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.154): Chi ngày là tỉ kiếp và thân nhược → "&IF(J76="
    #   ","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J76&") ⇒ "&G76&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P1 tr.154): Chi ngày là tỉ kiếp và thân nhược → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f203(E, r, c):
    # Ô E77
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0), X.eq(E.v(8,7,3),1.0))) else '—')

def f204(E, r, c):
    # Ô F77
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0), X.eq(E.v(8,7,4),1.0))) else '—')

def f205(E, r, c):
    # Ô K77
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Tỉ kiên bản khí ở ngày: "&COUNTIFS(SAO!$F$
    #   5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Kiếp tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.155): Nhật vượng, chi ngày là tỉ kiếp → "&IF(J77="",
    #   "chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J77&") ⇒ "&G77&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P1 tr.155): Nhật vượng, chi ngày là tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f206(E, r, c):
    # Ô E78, F78, E167, F167
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '—')

def f207(E, r, c):
    # Ô K78
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thương quan bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.154): Chi ngày là thương quan → "&IF(J78="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J78&") ⇒ "&G78&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P1 tr.154): Chi ngày là thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f208(E, r, c):
    # Ô E79, F79
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '—')

def f209(E, r, c):
    # Ô K79
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thất sát bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản k
    #   hí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.155; P2 tr.471): Chi ngày là thất sát → "&IF(J79="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J79&") ⇒ "&G79&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P1 tr.155; P2 tr.471): Chi ngày là thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f210(E, r, c):
    # Ô E80, F80
    # = IF((SAO!$R$14=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(13,14,18),1.0)) else '—')

def f211(E, r, c):
    # Ô K80
    # = "Dữ kiện: "&"Kình dương ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$R$14=1,"có","không")&" → Quy tắc (P1 tr.155; P2 tr.469): Chi ngày là kình dương →
    #    "&IF(J80="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J80&") ⇒ "&G80&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Kình dương ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,18),1.0)) else 'không')),' → Quy tắc (P1 tr.155; P2 tr.469): Chi ngày là kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f212(E, r, c):
    # Ô E81, F81
    # = IF((SAO!$S$14=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(13,14,19),1.0)) else '—')

def f213(E, r, c):
    # Ô K81
    # = "Dữ kiện: "&"Đào hoa ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$S$14=1,"có","không")&" → Quy tắc (P1 tr.155; P2 tr.467): Chi ngày đóng đào hoa → "&I
    #   F(J81="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J81&") ⇒ "&G81&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Đào hoa ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,19),1.0)) else 'không')),' → Quy tắc (P1 tr.155; P2 tr.467): Chi ngày đóng đào hoa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f214(E, r, c):
    # Ô E82, F82
    # = IF((SAO!$T$14=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(13,14,20),1.0)) else '—')

def f215(E, r, c):
    # Ô K82
    # = "Dữ kiện: "&"Hoa cái ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$T$14=1,"có","không")&" → Quy tắc (P1 tr.154; P2 tr.469): Chi ngày đóng hoa cái → "&I
    #   F(J82="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J82&") ⇒ "&G82&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Hoa cái ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,20),1.0)) else 'không')),' → Quy tắc (P1 tr.154; P2 tr.469): Chi ngày đóng hoa cái → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f216(E, r, c):
    # Ô E83, F83
    # = IF(TRUE(),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.TRUE()) else '—')

def f217(E, r, c):
    # Ô G83
    # = "Chi ngày "&INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&": "&CHOOSE(TRUNG_GIAN!$E$6+1,"vợ đẹp","bình thường","vợ đôn hậu","vợ đẹp","bình thường","vợ đôn hậu","vợ
    #    đẹp","bình thường","vợ đôn hậu","vợ đẹp","bình thường","vợ đôn hậu")
    return X.cat(X.cat(X.cat('Chi ngày ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0))),': '),X.CHOOSE(X.add(E.v(23,6,5),1.0), lambda: 'vợ đẹp', lambda: 'bình thường', lambda: 'vợ đôn hậu', lambda: 'vợ đẹp', lambda: 'bình thường', lambda: 'vợ đôn hậu', lambda: 'vợ đẹp', lambda: 'bình thường', lambda: 'vợ đôn hậu', lambda: 'vợ đẹp', lambda: 'bình thường', lambda: 'vợ đôn hậu'))

def f218(E, r, c):
    # Ô K83
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.463): Chi ngày là Tí, Ngọ, Mão, Dậu / Dần, Thân, Tị, Hợi / Thìn, Tuất, Sửu, Mùi → "&IF(J83="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J83&") ⇒ "&G83&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.463): Chi ngày là Tí, Ngọ, Mão, Dậu / Dần, Thân, Tị, Hợi / Thìn, Tuất, Sửu, Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f219(E, r, c):
    # Ô E84, F84
    # = IF(OR((MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$E$9),(MOD(TRUNG_GIAN!$E$9+1,5)=TRUNG_GIAN!$E$7)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,9,5)), X.eq(X.MOD(X.add(E.v(23,9,5),1.0), 5.0),E.v(23,7,5)))) else '—')

def f220(E, r, c):
    # Ô K84
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.467): Can chi trụ ngày tương sinh → "&IF(J84="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J84&") ⇒ "&G84&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.467): Can chi trụ ngày tương sinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f221(E, r, c):
    # Ô E85, F85
    # = IF(OR(LUC_THAN!$D$7=12,LUC_THAN!$D$7=13,LUC_THAN!$D$7=14,LUC_THAN!$D$7=27,LUC_THAN!$D$7=28,LUC_THAN!$D$7=29,LUC_THAN!$D$7=42,LUC_THAN!$D$7=43,LUC_THAN
    #   !$D$7=44,LUC_THAN!$D$7=57,LUC_THAN!$D$7=58,LUC_THAN!$D$7=59),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(16,7,4),12.0), X.eq(E.v(16,7,4),13.0), X.eq(E.v(16,7,4),14.0), X.eq(E.v(16,7,4),27.0), X.eq(E.v(16,7,4),28.0), X.eq(E.v(16,7,4),29.0), X.eq(E.v(16,7,4),42.0), X.eq(E.v(16,7,4),43.0), X.eq(E.v(16,7,4),44.0), X.eq(E.v(16,7,4),57.0), X.eq(E.v(16,7,4),58.0), X.eq(E.v(16,7,4),59.0))) else '—')

def f222(E, r, c):
    # Ô K85
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P1 tr.144; P2 tr.326, 470): Ngày âm dương lệch (Bính Tí, Đinh Sửu, Mậu Dần, Tân Mão, Nhâm Thìn, Quý Tị, Bính Ngọ, Đinh Mùi, Mậu Thân, Tân Dậu, 
    #   Nhâm Tuất, Quý Hợi) → "&IF(J85="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J85&") ⇒ "&G85&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.144; P2 tr.326, 470): Ngày âm dương lệch (Bính Tí, Đinh Sửu, Mậu Dần, Tân Mão, Nhâm Thìn, Quý Tị, Bính Ngọ, Đinh Mùi, Mậu Thân, Tân Dậu, Nhâm Tuất, Quý Hợi) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f223(E, r, c):
    # Ô E86, F86
    # = IF(OR((LUC_THAN!$D$9=1),(LUC_THAN!$D$10=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(16,9,4),1.0), X.eq(E.v(16,10,4),1.0))) else '—')

def f224(E, r, c):
    # Ô K86
    # = "Dữ kiện: "&"Cô thần: "&IF((LUC_THAN!$D$9=1),"có","không")&"; "&"Quả tú: "&IF((LUC_THAN!$D$10=1),"có","không")&" → Quy tắc (P2 tr.318-319): Có cô thần
    #    hoặc quả tú (tra theo chi năm) → "&IF(J86="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J86&") ⇒ "&G86&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Cô thần: '),('có' if X.truthy(X.eq(E.v(16,9,4),1.0)) else 'không')),'; '),'Quả tú: '),('có' if X.truthy(X.eq(E.v(16,10,4),1.0)) else 'không')),' → Quy tắc (P2 tr.318-319): Có cô thần hoặc quả tú (tra theo chi năm) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f225(E, r, c):
    # Ô E87, F87
    # = IF(OR(LUC_THAN!$D$7=38,LUC_THAN!$D$7=41,LUC_THAN!$D$7=42,LUC_THAN!$D$7=44,LUC_THAN!$D$7=47,LUC_THAN!$D$7=48,LUC_THAN!$D$7=53,LUC_THAN!$D$7=54,LUC_THAN
    #   !$D$8=38,LUC_THAN!$D$8=41,LUC_THAN!$D$8=42,LUC_THAN!$D$8=44,LUC_THAN!$D$8=47,LUC_THAN!$D$8=48,LUC_THAN!$D$8=53,LUC_THAN!$D$8=54),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(16,7,4),38.0), X.eq(E.v(16,7,4),41.0), X.eq(E.v(16,7,4),42.0), X.eq(E.v(16,7,4),44.0), X.eq(E.v(16,7,4),47.0), X.eq(E.v(16,7,4),48.0), X.eq(E.v(16,7,4),53.0), X.eq(E.v(16,7,4),54.0), X.eq(E.v(16,8,4),38.0), X.eq(E.v(16,8,4),41.0), X.eq(E.v(16,8,4),42.0), X.eq(E.v(16,8,4),44.0), X.eq(E.v(16,8,4),47.0), X.eq(E.v(16,8,4),48.0), X.eq(E.v(16,8,4),53.0), X.eq(E.v(16,8,4),54.0))) else '—')

def f226(E, r, c):
    # Ô K87
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.3
    #   25): Trụ ngày hoặc giờ thuộc nhóm cô loan (tham khảo) → "&IF(J87="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J87&") ⇒ "&G87&".
    #   ")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.325): Trụ ngày hoặc giờ thuộc nhóm cô loan (tham khảo) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f227(E, r, c):
    # Ô E88, F88
    # = IF(OR(LUC_THAN!$D$7=10,LUC_THAN!$D$7=40),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(16,7,4),10.0), X.eq(E.v(16,7,4),40.0))) else '—')

def f228(E, r, c):
    # Ô K88
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.471): Ngày Giáp Thìn hoặc Giáp Tuất → "&IF(J88="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J88&") ⇒ "&G88&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.471): Ngày Giáp Thìn hoặc Giáp Tuất → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f229(E, r, c):
    # Ô E89, F89
    # = IF((SAO!$M$14=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(13,14,13),1.0)) else '—')

def f230(E, r, c):
    # Ô K89
    # = "Dữ kiện: "&"Bị xung ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$M$14=1,"có","không")&" → Quy tắc (P1 tr.174; P2 tr.470): Chi ngày bị chi kề xung → "
    #   &IF(J89="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J89&") ⇒ "&G89&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Bị xung ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,13),1.0)) else 'không')),' → Quy tắc (P1 tr.174; P2 tr.470): Chi ngày bị chi kề xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f231(E, r, c):
    # Ô E90, F90
    # = IF(OR(INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1,INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1,INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIA
    #   N!$F$6+1)=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0))) else '—')

def f232(E, r, c):
    # Ô K90
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.470): Chi ngày gặp hình (vớ
    #   i chi khác trong tứ trụ) → "&IF(J90="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J90&") ⇒ "&G90&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.470): Chi ngày gặp hình (với chi khác trong tứ trụ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f233(E, r, c):
    # Ô E91, F91
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f234(E, r, c):
    # Ô K91
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.470, 476): Có thương quan → "&IF(J9
    #   1="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J91&") ⇒ "&G91&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.470, 476): Có thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f235(E, r, c):
    # Ô E92, F92
    # = IF((PP1_DINH_TINH!$F$96>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,96,6),E.v(24,99,2))) else '—')

def f236(E, r, c):
    # Ô K92
    # = "Dữ kiện: "&"Số chữ tỉ kiếp: "&PP1_DINH_TINH!$F$96&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"ng
    #   ưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.471, 478): Tỉ kiếp nhiều (≥ P_NHIEU) → "&IF(J92="","chưa đủ điều kiện ở cả hai phương pháp – không áp
    #    dụng.","ĐỦ ("&J92&") ⇒ "&G92&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ tỉ kiếp: '),E.v(9,96,6)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.471, 478): Tỉ kiếp nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f237(E, r, c):
    # Ô E93, F93
    # = IF((COUNTIFS(SAO!$R$5:$R$20,1)>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '—')

def f238(E, r, c):
    # Ô K93
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   ))&" → Quy tắc (P2 tr.471, 477): Có kình dương trong tứ trụ → "&IF(J93="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J93&") ⇒ "&
    #   G93&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.471, 477): Có kình dương trong tứ trụ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f239(E, r, c):
    # Ô E94, F94
    # = IF(OR(SUMPRODUCT(MOD(TRUNG_GIAN!$C$5:$F$5,2))+SUMPRODUCT(MOD(TRUNG_GIAN!$C$6:$F$6,2))=0,SUMPRODUCT(MOD(TRUNG_GIAN!$C$5:$F$5,2))+SUMPRODUCT(MOD(TRUNG_G
    #   IAN!$C$6:$F$6,2))=8),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(X.add(X.SUMPRODUCT(X.AF_MOD(E.rg(23,5,3,5,6), 2.0)),X.SUMPRODUCT(X.AF_MOD(E.rg(23,6,3,6,6), 2.0))),0.0), X.eq(X.add(X.SUMPRODUCT(X.AF_MOD(E.rg(23,5,3,5,6), 2.0)),X.SUMPRODUCT(X.AF_MOD(E.rg(23,6,3,6,6), 2.0))),8.0))) else '—')

def f240(E, r, c):
    # Ô K94
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&" → Quy
    #    tắc (P2 tr.471): Tứ trụ thuần dương / thuần âm → "&IF(J94="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J94&") ⇒ "&G94&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),' → Quy tắc (P2 tr.471): Tứ trụ thuần dương / thuần âm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f241(E, r, c):
    # Ô E95, F95
    # = IF((INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else '—')

def f242(E, r, c):
    # Ô K95
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.4
    #   71): Ngày và giờ xung nhau → "&IF(J95="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J95&") ⇒ "&G95&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.471): Ngày và giờ xung nhau → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f243(E, r, c):
    # Ô E96, F96
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Thất
    #    sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f244(E, r, c):
    # Ô K96
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&"
    #    → Quy tắc (P1 tr.153; P2 tr.465, 476): Nữ: có cả chính quan và thất sát (quan sát hỗn tạp) → "&IF(J96="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J96&") ⇒ "&G96&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.153; P2 tr.465, 476): Nữ: có cả chính quan và thất sát (quan sát hỗn tạp) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f245(E, r, c):
    # Ô E97, F97
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(9,99,6),0.0))) else '—')

def f246(E, r, c):
    # Ô K97
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 
    #   tr.477, 483): Nữ: không có quan sát → "&IF(J97="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J97&") ⇒ "&G97&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.477, 483): Nữ: không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f247(E, r, c):
    # Ô E98, F98
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(9,98,6),0.0))) else '—')

def f248(E, r, c):
    # Ô K98
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.46
    #   9, 474): Nam: không có tài → "&IF(J98="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J98&") ⇒ "&G98&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.469, 474): Nam: không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f249(E, r, c):
    # Ô E99
    # = IF(AND(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$C$7=1)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),(PP1_DINH_TINH!$F$98<PP1_DINH_TINH!$F$96)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,3),1.0)), X.ge(E.v(9,96,6),E.v(24,99,2))), X.lt(E.v(9,98,6),E.v(9,96,6)))) else '—')

def f250(E, r, c):
    # Ô F99
    # = IF(AND(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$D$7=1)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),(PP1_DINH_TINH!$F$98<PP1_DINH_TINH!$F$96)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,4),1.0)), X.ge(E.v(9,96,6),E.v(24,99,2))), X.lt(E.v(9,98,6),E.v(9,96,6)))) else '—')

def f251(E, r, c):
    # Ô K99
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"
    #   ; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P1 tr.155; P2 tr.469): Nam: tài ít hơn tỉ kiếp và tỉ kiếp nhiều, thân vượng → "&IF(J99="","chưa đủ điều
    #    kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J99&") ⇒ "&G99&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P1 tr.155; P2 tr.469): Nam: tài ít hơn tỉ kiếp và tỉ kiếp nhiều, thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f252(E, r, c):
    # Ô E100, F100
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Chín
    #   h tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COU
    #   NTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0))) else '—')

def f253(E, r, c):
    # Ô K100
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Chính tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Chính tài lâm Tuyệt: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$K$5:$K$20,8)>0,"Thiên tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF
    #   (COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Thiên tài lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ",""))&" → Quy tắc (P1 tr.155; P2 tr.468): Nam: tài lâm Mộ hoặc Tuyệt → "&IF(J100="","chưa
    #    đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J100&") ⇒ "&G100&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else ''),(X.cat(X.cat('Chính tài lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P1 tr.155; P2 tr.468): Nam: tài lâm Mộ hoặc Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f254(E, r, c):
    # Ô E101, F101
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20
    #   ,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,
    #   8)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0))) else '—')

def f255(E, r, c):
    # Ô K101
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Chính quan lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$2
    #   0,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Chính quan lâm Tuyệt: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thất sát lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")
    #   &IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Thất sát lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ",""))&" → Quy tắc (P2 tr.478, 480): Nữ: quan sát lâm Mộ hoặc Tuyệt → "&IF(J101="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J101&") ⇒ "&G101&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else ''),(X.cat(X.cat('Chính quan lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thất sát lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Thất sát lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P2 tr.478, 480): Nữ: quan sát lâm Mộ hoặc Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f256(E, r, c):
    # Ô E102, F102
    # = IF((SAO!$S$10=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(13,10,19),1.0)) else '—')

def f257(E, r, c):
    # Ô K102
    # = "Dữ kiện: "&"Đào hoa ở tháng (chi "&TRUNG_GIAN!$D$11&"): "&IF(SAO!$S$10=1,"có","không")&" → Quy tắc (P2 tr.471): Chi tháng là đào hoa (tháng sinh gặp 
    #   hàm trì) → "&IF(J102="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J102&") ⇒ "&G102&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Đào hoa ở tháng (chi '),E.v(23,11,4)),'): '),('có' if X.truthy(X.eq(E.v(13,10,19),1.0)) else 'không')),' → Quy tắc (P2 tr.471): Chi tháng là đào hoa (tháng sinh gặp hàm trì) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f258(E, r, c):
    # Ô E103, F103
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f259(E, r, c):
    # Ô K103
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2
    #    tr.477): Nữ: thương quan nhiều (≥ P_NHIEU) → "&IF(J103="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J103&") ⇒ "&G103&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.477): Nữ: thương quan nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f260(E, r, c):
    # Ô E104, E127
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$C$7=0)),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,3),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f261(E, r, c):
    # Ô F104, F127
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$D$7=0)),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,4),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f262(E, r, c):
    # Ô K104
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P1 tr.155; P2 tr.468): Nam: thân nhược, tài 
    #   nhiều (≥ P_NHIEU) → "&IF(J104="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J104&") ⇒ "&G104&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P1 tr.155; P2 tr.468): Nam: thân nhược, tài nhiều (≥ P_NHIEU) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f263(E, r, c):
    # Ô E105, F105
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0))=1))
    #   ,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),1.0))) else '—')

def f264(E, r, c):
    # Ô K105
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0
    #   ,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.153; P2 tr.465): Nữ: chỉ một quan hoặc một sát lộ 
    #   rõ, không hỗn tạp → "&IF(J105="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J105&") ⇒ "&G105&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P1 tr.153; P2 tr.465): Nữ: chỉ một quan hoặc một sát lộ rõ, không hỗn tạp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f265(E, r, c):
    # Ô E106
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0)))) else '—')

def f266(E, r, c):
    # Ô F106
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0)))) else '—')

def f267(E, r, c):
    # Ô K106
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&
    #   INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.464, 466): Nữ: quan sát là dụng thần → "&IF(J106="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J106&") ⇒ "&G106&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.464, 466): Nữ: quan sát là dụng thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f268(E, r, c):
    # Ô E107
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)))) else '—')

def f269(E, r, c):
    # Ô F107
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)))) else '—')

def f270(E, r, c):
    # Ô K107
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&
    #   INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.464): Nam: tài là dụng thần → "&IF(J107="","chưa đủ điều kiện ở cả hai phương pháp – không áp
    #    dụng.","ĐỦ ("&J107&") ⇒ "&G107&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.464): Nam: tài là dụng thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f271(E, r, c):
    # Ô E112, F112
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),
    #   "ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0))) else '—')

def f272(E, r, c):
    # Ô K112
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Kiếp tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là kiếp tài → "&IF(J112="","chưa
    #    đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J112&") ⇒ "&G112&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là kiếp tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f273(E, r, c):
    # Ô L112, L113, L114, L115, L116, L117, L118, L119, L120, L121, L122, L123 … (167 ô)
    # = IF(DOI_CHIEU!$C$33=2,F112,E112)
    return (E.v(17,r,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(17,r,c-7))

def f274(E, r, c):
    # Ô E113, F113
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0))) else '—')

def f275(E, r, c):
    # Ô K113
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thiên tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là tài → "&IF(J113="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J113&") ⇒ "&G113&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f276(E, r, c):
    # Ô E114, F114
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+
    #   COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),AND(TRUE(),LEN(THAN_SAT!$E$93)>0
    #   )),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,93,5)),0.0)))) else '—')

def f277(E, r, c):
    # Ô K114
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thiên tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là tài và chi ngày có tướng tinh → "&IF(J114="","chưa đủ
    #    điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J114&") ⇒ "&G114&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là tài và chi ngày có tướng tinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f278(E, r, c):
    # Ô E115, F115
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNT
    #   IFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0))) else '—')

def f279(E, r, c):
    # Ô K115
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính ấn bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thiên ấn bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là chính ấn hoặc thiên ấn → "&IF(J115="","chưa đủ điều kiện ở 
    #   cả hai phương pháp – không áp dụng.","ĐỦ ("&J115&") ⇒ "&G115&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là chính ấn hoặc thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f280(E, r, c):
    # Ô E116, F116
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0))
    #   ,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0))) else '—')

def f281(E, r, c):
    # Ô K116
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thực thần bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là thực thần → "&IF(J116="","
    #   chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J116&") ⇒ "&G116&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P2 tr.463): Nam, bản khí chi ngày là thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f282(E, r, c):
    # Ô E117, F117
    # = IF(AND(TRUE(),SAO!$P$13=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0))) else '—')

def f283(E, r, c):
    # Ô K117
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.463-464): Chi ngày là thiên ất quý nhân → "&IF(J117="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J117&") ⇒ "&G117&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.463-464): Chi ngày là thiên ất quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f284(E, r, c):
    # Ô E118, F118
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+
    #   COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),NOT((SAO!$M$14=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.NOT(X.eq(E.v(13,14,13),1.0)))) else '—')

def f285(E, r, c):
    # Ô K118
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Bị xung ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$M$14=1,"có","không")&"; "&"Các chữ liên quan (chỉ
    #    liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"C
    #   hính tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","
    #   ")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thiên tài bản khí ở ngày: "&
    #   COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.463-
    #   464): Nam, bản khí chi ngày là tài và chi ngày không bị xung → "&IF(J118="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J118&") ⇒
    #    "&G118&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Bị xung ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,13),1.0)) else 'không')),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.463-464): Nam, bản khí chi ngày là tài và chi ngày không bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f286(E, r, c):
    # Ô E119
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$C$7=1)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO
    #   !$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,3),1.0)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f287(E, r, c):
    # Ô F119
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$D$7=1)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO
    #   !$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,4),1.0)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f288(E, r, c):
    # Ô K119
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM
    #   (IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh
    #   : "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.463-464): Nam, thân vượng và tài có lực → "&IF(J119="","chưa đủ điề
    #   u kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J119&") ⇒ "&G119&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.463-464): Nam, thân vượng và tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f289(E, r, c):
    # Ô E120, F120
    # = IF(AND(AND(TRUE(),SAO!$Q$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.TRUE(), X.eq(E.v(13,13,17),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0)))) else '—')

def f290(E, r, c):
    # Ô K120
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.463): Chi ngày là lộc và chi giờ là quý nhân → "&IF(J120="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J120&") ⇒ "&G120&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.463): Chi ngày là lộc và chi giờ là quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f291(E, r, c):
    # Ô E121, F121
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$Q$17=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,17),1.0)))) else '—')

def f292(E, r, c):
    # Ô K121
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.463): Nam, chi giờ vừa là quý nhân vừa là lộc → "&IF(J121="","chưa đủ điều kiện ở cả hai 
    #   phương pháp – không áp dụng.","ĐỦ ("&J121&") ⇒ "&G121&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.463): Nam, chi giờ vừa là quý nhân vừa là lộc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f293(E, r, c):
    # Ô E122, F122
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f294(E, r, c):
    # Ô K122
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5
    #   :$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.463): Nam, thiên tài vượng tại lệnh → "&IF(J122="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.
    #   ","ĐỦ ("&J122&") ⇒ "&G122&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.463): Nam, thiên tài vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f295(E, r, c):
    # Ô E123, F123
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f296(E, r, c):
    # Ô K123
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "
    #   &NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J
    #   $5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J
    #   $5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính 
    #   tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; 
    #   ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng t
    #   ại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc 
    #   (P2 tr.463): Nam, tài có lực → "&IF(J123="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J123&") ⇒ "&G123&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.463): Nam, tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f297(E, r, c):
    # Ô E124, F124
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),AND(TRUE(),LEN(THAN_SAT!$E$92)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0)))) else '—')

def f298(E, r, c):
    # Ô K124
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.463, 470): Nam, trạch mã ở chi ngày → "&IF(J124="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J124&") ⇒ "&G124&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.463, 470): Nam, trạch mã ở chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f299(E, r, c):
    # Ô E125, F125
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),AND(TRUE(),SAO!$Q$13=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.TRUE(), X.eq(E.v(13,13,17),1.0)))) else '—')

def f300(E, r, c):
    # Ô K125
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.463): Nam, chi ngày là lộc của nhật can → "&IF(J125="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J125&") ⇒ "&G125&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.463): Nam, chi ngày là lộc của nhật can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f301(E, r, c):
    # Ô E126, F126
    # = IF(AND(OR(AND(TRUE(),SAO!$Q$13=1),AND(TRUE(),LEN(THAN_SAT!$E$92)>0)),OR(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$Q$17=1),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),LEN(THAN_SAT!$F$92)>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,13,17),1.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0))), X.OR(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,17),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0))))) else '—')

def f302(E, r, c):
    # Ô K126
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.463): Chi ngày và chi giờ đều có lộc hoặc trạch mã → "&IF(J126="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J126&") ⇒ "&G126&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.463): Chi ngày và chi giờ đều có lộc hoặc trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f303(E, r, c):
    # Ô K127
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.463-464, 470): Nam, thân nhược và tài 
    #   ≥ P_NHIEU → "&IF(J127="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J127&") ⇒ "&G127&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.463-464, 470): Nam, thân nhược và tài ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f304(E, r, c):
    # Ô E128, F128
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),AND(TRUE(),LEN(THAN_SAT!$E$91)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,91,5)),0.0)))) else '—')

def f305(E, r, c):
    # Ô K128
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.464): Nam, đào hoa ở chi ngày → "&IF(J128="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J128&") ⇒ "&G128&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.464): Nam, đào hoa ở chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f306(E, r, c):
    # Ô E129, F129
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$5:$
    #   F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0)))) else '—')

def f307(E, r, c):
    # Ô K129
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Chính tài tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thiên tài tàng can: "
    #   &COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Chính tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,S
    #   AO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thiên tài thiên ca
    #   n: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.464, 469): Nam, tài 
    #   chỉ tàng trong chi (không lộ ở can) → "&IF(J129="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J129&") ⇒ "&G129&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')),(X.cat(X.cat('Chính tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.464, 469): Nam, tài chỉ tàng trong chi (không lộ ở can) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f308(E, r, c):
    # Ô E130, F130
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0))) else '—')

def f309(E, r, c):
    # Ô K130
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Chính tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thiên tài thiên can: "&COU
    #   NTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nam, tài lộ ở can → "&
    #   IF(J130="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J130&") ⇒ "&G130&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nam, tài lộ ở can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f310(E, r, c):
    # Ô E131, F131
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f311(E, r, c):
    # Ô K131
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1)>0,"Thiên tài: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.464): 
    #   Nam, có thiên tài → "&IF(J131="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J131&") ⇒ "&G131&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.464): Nam, có thiên tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f312(E, r, c):
    # Ô E132, F132
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),OR(OR(AND(TRUE(),SAO!$P$5=1),AND(TRUE(),SAO!$Q$5=1)),AND(TRUE(),LEN(THAN_SAT!$C$92
    #   )>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,5,17),1.0))), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,3)),0.0))))) else '—')

def f313(E, r, c):
    # Ô K132
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5
    #   :$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNT
    #   IFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nam, tài vượng tại lệnh và c
    #   hi năm có quý nhân/lộc/trạch mã → "&IF(J132="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J132&") ⇒ "&G132&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nam, tài vượng tại lệnh và chi năm có quý nhân/lộc/trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f314(E, r, c):
    # Ô E133, F133
    # = IF(AND(AND(OR(AND(TRUE(),SAO!$P$5=1),AND(TRUE(),SAO!$P$9=1),AND(TRUE(),SAO!$P$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1)),OR(AND(TRUE()
    #   ,SAO!$Q$5=1),AND(TRUE(),SAO!$Q$9=1),AND(TRUE(),SAO!$Q$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$Q$17=1))),(THAN_SAT!$H$92="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,9,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,17),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,9,17),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,17),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,17),1.0)))), X.eq(E.v(12,92,8),'Có'))) else '—')

def f315(E, r, c):
    # Ô K133
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.464): Có quý nhân, có lộc và có trạch mã → "&IF(J133="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J133&") ⇒ "&G133&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.464): Có quý nhân, có lộc và có trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f316(E, r, c):
    # Ô E134, F134
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—"
    #   )
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f317(E, r, c):
    # Ô K134
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên 
    #   quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5
    #   :$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nam, có thương quan và có ấn chế → "&IF(J134="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J134&") ⇒ "&G134&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.464): Nam, có thương quan và có ấn chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f318(E, r, c):
    # Ô E135, F135
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)
    #   )>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên 
    #   ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,0))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0)),AND(((COUNTI
    #   FS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,2))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B
    #   $20,3))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3))>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0)),0.0)), X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),0.0)), X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0)),0.0))))) else '—')

def f319(E, r, c):
    # Ô K135
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)>0,"Thương quan thiên can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$B$5:$B$20,0)>0,"Chính ấn bản khí ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$
    #   20,0)>0,"Thiên ấn bản khí ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,0)&" c
    #   hữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Thương quan thiên can ở
    #    tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Chính ấn bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F
    #   $20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Thiên ấn bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,2)>0,"Thương quan thiên can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$C$5:$C$20,"Can",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!
    #   $B$5:$B$20,2)>0,"Chính ấn bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$
    #   20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thiên ấn bản 
    #   khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS
    #   (SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)>0,"Thương quan thiên can ở giờ: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3)>0,"Chính ấn bản khí ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3)>0,"Thiên ấn bản khí ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$
    #   5:$C$20,"Bản khí",SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.464, 468): Nam, có trụ mà can là thương quan, bản khí chi là ấn → "&IF(J135="","c
    #   hưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J135&") ⇒ "&G135&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Thương quan thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Chính ấn bản khí ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Thương quan thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thương quan thiên can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thương quan thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.464, 468): Nam, có trụ mà can là thương quan, bản khí chi là ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f320(E, r, c):
    # Ô E136, F136
    # = IF(AND(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20
    #   ,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$97>=1)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,97,6),1.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f321(E, r, c):
    # Ô K136
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_
    #   CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Tỉ kiên vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Kiếp tài
    #    vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nam, 
    #   tỉ kiếp vượng tại lệnh, có thực thương và tài → "&IF(J136="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J136&") ⇒ "&G136&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nam, tỉ kiếp vượng tại lệnh, có thực thương và tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f322(E, r, c):
    # Ô E137, F137
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0)+COUNTIFS(SAO
    #   !$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0)),0.0))) else '—')

def f323(E, r, c):
    # Ô K137
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0)>0,"Chính tài không bị xung lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   M$5:$M$20,0,SAO!$K$5:$K$20,0)>0,"Thiên tài không bị xung lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $M$5:$M$20,0,SAO!$K$5:$K$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nam, tài lâm Trường sinh, không bị xung → "&IF(J137="","chưa đủ điều kiện ở cả h
    #   ai phương pháp – không áp dụng.","ĐỦ ("&J137&") ⇒ "&G137&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài không bị xung lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài không bị xung lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nam, tài lâm Trường sinh, không bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f324(E, r, c):
    # Ô E138
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$C$7=0)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO
    #   !$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,3),0.0)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f325(E, r, c):
    # Ô F138
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$D$7=0)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO
    #   !$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,4),0.0)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f326(E, r, c):
    # Ô K138
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM
    #   (IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh
    #   : "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.468): Nam, thân nhược và tài có lực → "&IF(J138="","chưa đủ điều ki
    #   ện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J138&") ⇒ "&G138&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.468): Nam, thân nhược và tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f327(E, r, c):
    # Ô E139, F139
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f328(E, r, c):
    # Ô K139
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1)>0,"Tỉ kiên: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp
    #    tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắ
    #   c (P2 tr.468): Nam, có cả tỉ kiên và kiếp tài → "&IF(J139="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J139&") ⇒ "&G139&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.468): Nam, có cả tỉ kiên và kiếp tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f329(E, r, c):
    # Ô E140, F140
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)+COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$
    #   20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 1.0)),0.0)))) else '—')

def f330(E, r, c):
    # Ô K140
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)>0,"Chính tài đào hoa: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,
    #   1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)>0,"Thiên tài đào hoa: "&COUNTIFS(SAO!$F$5:$
    #   F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$K$5:$K$20,1)>0,"Chính tài lâm Mộc dục: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,1)&" chữ; ",""
    #   )&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,1)>0,"Thiên tài lâm Mộc dục: "&COUNTIFS(SAO!$F$5:$F$20,"Thiê
    #   n tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, tài đóng cùng đào hoa hoặc lâm Mộc dục → "&IF(J
    #   140="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J140&") ⇒ "&G140&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài đào hoa: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài đào hoa: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài lâm Mộc dục: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Mộc dục: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, tài đóng cùng đào hoa hoặc lâm Mộc dục → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f331(E, r, c):
    # Ô E141, F141
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","
    #   —")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f332(E, r, c):
    # Ô K141
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)
    #   >0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, có kiếp tài và kình dương → "&IF(J141="","chưa đủ điều kiện ở
    #    cả hai phương pháp – không áp dụng.","ĐỦ ("&J141&") ⇒ "&G141&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, có kiếp tài và kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f333(E, r, c):
    # Ô E142, F142
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),(IF(AND(TRUE(),OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$C$6=7)),1,0)+IF(AND(TRUE(),OR(TRUN
    #   G_GIAN!$D$6=4,TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$D$6=7)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$E$6=1,TR
    #   UNG_GIAN!$E$6=7)),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=4,TRUNG_GIAN!$F$6=10,TRUNG_GIAN!$F$6=1,TRUNG_GIAN!$F$6=7)),1,0))>=
    #   3),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,3),7.0)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,4),7.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,5),7.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),4.0), X.eq(E.v(23,6,6),10.0), X.eq(E.v(23,6,6),1.0), X.eq(E.v(23,6,6),7.0)))) else 0.0)),3.0))) else '—')

def f334(E, r, c):
    # Ô K142
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.469): Nam, có ≥ 3 chi Thìn/Tuất/Sửu/Mùi → "&IF(J142="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J142&") ⇒ "&G142&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.469): Nam, có ≥ 3 chi Thìn/Tuất/Sửu/Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f335(E, r, c):
    # Ô E143, F143
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),0.0))) else '—')

def f336(E, r, c):
    # Ô K143
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$
    #   20,1)>0,"Chính tài tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Thiên tài tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Ch
    #   ính quan tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Thất sát tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469, 477): Có tài ở tuần không và có quan sát ở tuần không → "&IF(J143="","chưa đủ điều k
    #   iện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J143&") ⇒ "&G143&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469, 477): Có tài ở tuần không và có quan sát ở tuần không → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f337(E, r, c):
    # Ô E144, F144
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f338(E, r, c):
    # Ô K144
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1)>0,"Chính tài: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên tài: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&"
    #    → Quy tắc (P2 tr.469-470): Nam, có cả chính tài và thiên tài → "&IF(J144="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J144&") 
    #   ⇒ "&G144&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469-470): Nam, có cả chính tài và thiên tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f339(E, r, c):
    # Ô E145, F145
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(PP1_DINH_TINH!$F$98>=1)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.ge(E.v(9,98,6),1.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f340(E, r, c):
    # Ô K145
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên qu
    #   an (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại
    #    lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, có tài nhưng không chữ tài nào vượng tại lệnh → "&IF(J145="","chưa đủ điều kiện ở cả h
    #   ai phương pháp – không áp dụng.","ĐỦ ("&J145&") ⇒ "&G145&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, có tài nhưng không chữ tài nào vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f341(E, r, c):
    # Ô E146, F146
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),0.0)))) else '—')

def f342(E, r, c):
    # Ô K146
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Chính tài ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thiên tài ở ngày: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, chi ngày không tàng tài → "&IF(J146="",
    #   "chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J146&") ⇒ "&G146&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, chi ngày không tàng tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f343(E, r, c):
    # Ô E147, F147
    # = IF(AND(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2))), X.eq(X.INDEX(E.rg(8,24,3,28,3), 5.0),0.0))) else '—')

def f344(E, r, c):
    # Ô K147
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.469, 479): Hỏa ≥ P_NHIEU, Thổ ≥ P_NHIEU, không có Thủy → "&IF(J147="","chưa đủ điều ki
    #   ện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J147&") ⇒ "&G147&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.469, 479): Hỏa ≥ P_NHIEU, Thổ ≥ P_NHIEU, không có Thủy → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f345(E, r, c):
    # Ô E148, F148
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),(TRUNG_GIAN!$E$60=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.eq(E.v(23,60,5),1.0))) else '—')

def f346(E, r, c):
    # Ô K148
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Tỉ kiên bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Kiếp tài bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản
    #    khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, bản khí chi ngày là tỉ kiếp và chi ngày thuộc cục → "&IF(J148="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J148&") ⇒ "&G148&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, bản khí chi ngày là tỉ kiếp và chi ngày thuộc cục → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f347(E, r, c):
    # Ô E149, F149
    # = IF(AND(TRUNG_GIAN!$E$60=1,PP2_DINH_LUONG!$E$24=MOD(TRUNG_GIAN!$E$7+3,5)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,60,5),1.0), X.eq(E.v(10,24,5),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—')

def f348(E, r, c):
    # Ô K149
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.469, 473): Chi ngày thuộc cục hóa thành hành khắc nhật can → "&IF(J149="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ (
    #   "&J149&") ⇒ "&G149&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.469, 473): Chi ngày thuộc cục hóa thành hành khắc nhật can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f349(E, r, c):
    # Ô E150, F150
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1,SAO!$K$5:$K$20,9)+COUNTIFS(SAO
    #   !$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1,SAO!$K$5:$K$20,9))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0, E.rg(13,5,11,20,11), 9.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0))) else '—')

def f350(E, r, c):
    # Ô K150
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1,SAO!$K$5:$K$20,9)>0,"Chính tài ở tháng lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$B$5:$B$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1,
    #   SAO!$K$5:$K$20,9)>0,"Thiên tài ở tháng lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1,SAO!$K$5:$K
    #   $20,9)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, tài ở trụ tháng lâm Tuyệt → "&IF(J150="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.",
    #   "ĐỦ ("&J150&") ⇒ "&G150&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài ở tháng lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else ''),(X.cat(X.cat('Thiên tài ở tháng lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, tài ở trụ tháng lâm Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f351(E, r, c):
    # Ô E151, F151
    # = IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=4))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),4.0))))) else '—')

def f352(E, r, c):
    # Ô K151
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.4
    #   69, 478): Chi ngày và chi giờ là Thìn – Tuất xung nhau → "&IF(J151="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J151&") ⇒ "&G15
    #   1&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.469, 478): Chi ngày và chi giờ là Thìn – Tuất xung nhau → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f353(E, r, c):
    # Ô E152, F152
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(PP1_DINH_TINH!$F$98<=1)),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.le(E.v(9,98,6),1.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f354(E, r, c):
    # Ô K152
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên qu
    #   an (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"
    #   Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, tài ≤ 1 chữ và có kiếp tài → "&IF(J152="","chưa đủ điều kiện 
    #   ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J152&") ⇒ "&G152&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.469): Nam, tài ≤ 1 chữ và có kiếp tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f355(E, r, c):
    # Ô E153, F153
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))
    #   >0)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f356(E, r, c):
    # Ô K153
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thực thần bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"T
    #   hiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, bản khí chi ngày là thực 
    #   thần và có thiên ấn → "&IF(J153="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J153&") ⇒ "&G153&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Nam, bản khí chi ngày là thực thần và có thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f357(E, r, c):
    # Ô E154, F154
    # = IF(AND((THAN_SAT!$H$122="Có"),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(12,122,8),'Có'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f358(E, r, c):
    # Ô K154
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Có vong thần và thất sát → "&IF(J154="
    #   ","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J154&") ⇒ "&G154&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.469): Có vong thần và thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f359(E, r, c):
    # Ô E155, F155
    # = IF(AND(AND(AND(AND((NHAP_LIEU!$C$22="Nam"),OR(AND(TRUE(),TRUNG_GIAN!$C$5=0),AND(TRUE(),TRUNG_GIAN!$D$5=0),AND(TRUE(),TRUNG_GIAN!$E$5=0),AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=0))),OR(AND(TRUE(),TRUNG_GIAN!$C$5=1),AND(TRUE(),TRUNG_GIAN!$D$5=1),AND(TRUE(),TRUNG_GIAN!$E$5=1),AND(NOT(NHA
    #   P_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=1))),OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),AND(TRUE(),TRUNG_GIAN!$E$5=2),AND(NOT(
    #   NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=2))),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3),AND(TRUE(),TRUNG_GIAN!$E$5=3),AND(N
    #   OT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=3))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),0.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),1.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0))))) else '—')

def f360(E, r, c):
    # Ô K155
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.469): Nam, có đủ Giáp, Ất, Bính, Đinh ở thiên can → "&IF(J155="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J155&") ⇒ "&
    #   G155&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.469): Nam, có đủ Giáp, Ất, Bính, Đinh ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f361(E, r, c):
    # Ô E156, F156
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0)),(PP1_DINH_TINH!$F$96
    #   >=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)), X.ge(E.v(9,96,6),1.0))) else '—')

def f362(E, r, c):
    # Ô K156
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liê
    #   n quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Thiên tài ở giờ
    #   : "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Nam, thiên tài ở tr
    #   ụ giờ và có tỉ kiếp → "&IF(J156="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J156&") ⇒ "&G156&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else ''))),' → Quy tắc (P2 tr.469): Nam, thiên tài ở trụ giờ và có tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f363(E, r, c):
    # Ô E157, F157
    # = IF(AND(OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),(SAO!$R$14=1)),(SAO!$R$18=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),
    #   "ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0), X.eq(E.v(13,14,18),1.0)), X.eq(E.v(13,18,18),1.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f364(E, r, c):
    # Ô K157
    # = "Dữ kiện: "&"Kình dương ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$R$14=1,"có","không")&"; "&"Kình dương ở giờ (chi "&TRUNG_GIAN!$F$11&"): "&IF(SAO!
    #   $R$18=1,"có","không")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$B$5:$B$20,2)>0,"Thất sát ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNT
    #   IFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Thất sát ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(
    #   SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.469): Trụ ngày/giờ có thất sát hoặc kình dương, và có th
    #   iên ấn → "&IF(J157="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J157&") ⇒ "&G157&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Kình dương ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,18),1.0)) else 'không')),'; '),'Kình dương ở giờ (chi '),E.v(23,11,6)),'): '),('có' if X.truthy(X.eq(E.v(13,18,18),1.0)) else 'không')),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thất sát ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thất sát ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.469): Trụ ngày/giờ có thất sát hoặc kình dương, và có thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f365(E, r, c):
    # Ô E158, F158
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)+COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$20,"Thiên 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNT
    #   IFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.add(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0)) else '—')

def f366(E, r, c):
    # Ô K158
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$
    #   20,7)>0,"Chính tài lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5
    #   :$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Chính tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Chính tài l
    #   âm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)>0,"Thiên tài lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thiên tài lâm Mộ: "&COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Thiên tài lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" 
    #   chữ; ",""))&" → Quy tắc (P2 tr.470): Có tài lâm Tử, Mộ hoặc Tuyệt → "&IF(J158="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J158
    #   &") ⇒ "&G158&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),0.0)) else ''),(X.cat(X.cat('Chính tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Chính tài lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P2 tr.470): Có tài lâm Tử, Mộ hoặc Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f367(E, r, c):
    # Ô E159, F159
    # = IF(AND(TRUNG_GIAN!$C$5=TRUNG_GIAN!$E$5,TRUNG_GIAN!$C$6=TRUNG_GIAN!$E$6),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,3),E.v(23,5,5)), X.eq(E.v(23,6,3),E.v(23,6,5)))) else '—')

def f368(E, r, c):
    # Ô K159
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.4
    #   70, 478): Trụ năm và trụ ngày cùng can chi → "&IF(J159="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J159&") ⇒ "&G159&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.470, 478): Trụ năm và trụ ngày cùng can chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f369(E, r, c):
    # Ô E160, F160
    # = IF(OR(TRUNG_GIAN!$E$28=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,28,5),1.0))) else '—')

def f370(E, r, c):
    # Ô K160
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.470): Nhật can lâm Mộc dục tại chi ngày → "&IF(J160="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J160&") ⇒ "&G160&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.470): Nhật can lâm Mộc dục tại chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f371(E, r, c):
    # Ô E161, F161
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.ge(E.v(9,98,6),E.v(24,99,2))), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f372(E, r, c):
    # Ô K161
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.470): Nam, tài ≥ P_NHIE
    #   U và tỉ kiếp ≥ P_NHIEU → "&IF(J161="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J161&") ⇒ "&G161&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.470): Nam, tài ≥ P_NHIEU và tỉ kiếp ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f373(E, r, c):
    # Ô E162
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$C$7=1)),(SAO!$R$14=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,3),1.0)), X.eq(E.v(13,14,18),1.0))) else '—')

def f374(E, r, c):
    # Ô F162
    # = IF(AND(AND((NHAP_LIEU!$C$22="Nam"),(DUNG_HY_KY!$D$7=1)),(SAO!$R$14=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.eq(E.v(8,7,4),1.0)), X.eq(E.v(13,14,18),1.0))) else '—')

def f375(E, r, c):
    # Ô K162
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Kình dương ở ngày (chi "&T
    #   RUNG_GIAN!$E$11&"): "&IF(SAO!$R$14=1,"có","không")&" → Quy tắc (P2 tr.470): Nam, thân vượng và chi ngày là kình dương → "&IF(J162="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J162&") ⇒ "&G162&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Kình dương ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,18),1.0)) else 'không')),' → Quy tắc (P2 tr.470): Nam, thân vượng và chi ngày là kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f376(E, r, c):
    # Ô E163, F163
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),OR(AND(TRUE(),PP2_DINH_LUONG!$C$43=1,OR(AND(OR(TRUNG_GIAN!$C$12="Tỉ kiên"),OR(TRUNG_GIAN!$D$12="Chính tài")),AND(OR(TRU
    #   NG_GIAN!$C$12="Chính tài"),OR(TRUNG_GIAN!$D$12="Tỉ kiên")))),AND(TRUE(),PP2_DINH_LUONG!$D$43=1,OR(AND(OR(TRUNG_GIAN!$D$12="Tỉ kiên"),FALSE()),AND(OR(T
    #   RUNG_GIAN!$D$12="Chính tài"),TRUE()))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1,OR(AND(TRUE(),OR(TRUNG_GIAN!$F$12="Chính tài")),A
    #   ND(FALSE(),OR(TRUNG_GIAN!$F$12="Tỉ kiên")))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.OR(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,12,3),'Tỉ kiên')), X.OR(X.eq(E.v(23,12,4),'Chính tài'))), X.AND(X.OR(X.eq(E.v(23,12,3),'Chính tài')), X.OR(X.eq(E.v(23,12,4),'Tỉ kiên'))))), X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,12,4),'Tỉ kiên')), X.FALSE()), X.AND(X.OR(X.eq(E.v(23,12,4),'Chính tài')), X.TRUE()))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0), X.OR(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,12,6),'Chính tài'))), X.AND(X.FALSE(), X.OR(X.eq(E.v(23,12,6),'Tỉ kiên')))))))) else '—')

def f377(E, r, c):
    # Ô K163
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.470): Nam, tỉ kiên và chính tài kề nhau thành ngũ hợp → "&IF(J163="","chưa đủ điều kiện ở
    #    cả hai phương pháp – không áp dụng.","ĐỦ ("&J163&") ⇒ "&G163&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.470): Nam, tỉ kiên và chính tài kề nhau thành ngũ hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f378(E, r, c):
    # Ô E164, F164
    # = IF(AND(AND(AND(AND((NHAP_LIEU!$C$22="Nam"),(PP1_DINH_TINH!$F$98>=1)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),(PP1_DINH_TINH!$F$99=0)),(PP1_DINH_TINH!$F$97=0))
    #   ,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(2,22,3),'Nam'), X.ge(E.v(9,98,6),1.0)), X.ge(E.v(9,96,6),E.v(24,99,2))), X.eq(E.v(9,99,6),0.0)), X.eq(E.v(9,97,6),0.0))) else '—')

def f379(E, r, c):
    # Ô K164
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(M
    #   ENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sá
    #   t: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Q
    #   uy tắc (P2 tr.470): Nam, có tài, tỉ kiếp ≥ P_NHIEU, không có quan sát và thực thương cứu → "&IF(J164="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J164&") ⇒ "&G164&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.470): Nam, có tài, tỉ kiếp ≥ P_NHIEU, không có quan sát và thực thương cứu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f380(E, r, c):
    # Ô E165, F165
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),0.0))) else '—')

def f381(E, r, c):
    # Ô K165
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Tỉ kiên bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Kiếp tài bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"B
    #   ản khí",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.470): Nam, bản khí chi tháng là tỉ kiếp → "&IF(J165="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J165&") ⇒ "&G165&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Tỉ kiên bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.470): Nam, bản khí chi tháng là tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f382(E, r, c):
    # Ô E166, F166
    # = IF((TRUNG_GIAN!$E$7=TRUNG_GIAN!$E$9),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(23,7,5),E.v(23,9,5))) else '—')

def f383(E, r, c):
    # Ô K166
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.470): Can ngày và chi ngày cùng hành → "&IF(J166="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J166&") ⇒ "&G166&"."
    #   )
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.470): Can ngày và chi ngày cùng hành → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f384(E, r, c):
    # Ô K167
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thương quan bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.470, 480): Bản khí chi ngày là thương quan → "&IF(J167="","chưa đủ điều kiện ở cả hai 
    #   phương pháp – không áp dụng.","ĐỦ ("&J167&") ⇒ "&G167&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P2 tr.470, 480): Bản khí chi ngày là thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f385(E, r, c):
    # Ô E168, F168
    # = IF(AND((PP1_DINH_TINH!$F$99>=P_NHIEU),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,99,6),E.v(24,99,2)), X.eq(E.v(9,98,6),0.0))) else '—')

def f386(E, r, c):
    # Ô K168
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.471, 481): Quan sát ≥ P_NHIEU và không có tài → "&IF(J168
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J168&") ⇒ "&G168&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.471, 481): Quan sát ≥ P_NHIEU và không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f387(E, r, c):
    # Ô E169, F169
    # = IF(OR(AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=0),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9
    #   ="Không rõ giờ"),TRUNG_GIAN!$F$5=0),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5
    #   =5),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=5),1,0))>=2),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D
    #   $5=5),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=5),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=5),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=0
    #   ),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=0),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=0),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=0),1,0))>
    #   =2),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="
    #   Không rõ giờ"),TRUNG_GIAN!$F$5=1),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=6),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=6),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=6
    #   ),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=6),1,0))>=2),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=6),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5
    #   =6),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=6),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=6),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=1),
    #   1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=1),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=1),1,0))>=2
    #   ),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=2),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$5=2),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=7),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=7),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=7),
    #   1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=7),1,0))>=2),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=7),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=7
    #   ),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=7),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=7),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=2),1,
    #   0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=2),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=2),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=2),1,0))>=2),
    #   AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=3),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Khôn
    #   g rõ giờ"),TRUNG_GIAN!$F$5=3),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=8),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=8),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=8),1,
    #   0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8),1,0))>=2),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=8),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=8),
    #   1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=8),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=3),1,0)
    #   +IF(AND(TRUE(),TRUNG_GIAN!$D$5=3),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=3),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=3),1,0))>=2),AN
    #   D((IF(AND(TRUE(),TRUNG_GIAN!$C$5=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=4),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không 
    #   rõ giờ"),TRUNG_GIAN!$F$5=4),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=9),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=9),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=9),1,0)
    #   +IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=9),1,0))>=2),AND((IF(AND(TRUE(),TRUNG_GIAN!$C$5=9),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=9),1,
    #   0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=9),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=9),1,0))=1,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=4),1,0)+I
    #   F(AND(TRUE(),TRUNG_GIAN!$D$5=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=4),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=4),1,0))>=2)),"ĐỦ
    #   ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),0.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),5.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),5.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),5.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),5.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),5.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),5.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),5.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),5.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),0.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),0.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),1.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),1.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),8.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),8.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),4.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),9.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),9.0))) else 0.0)),2.0)), X.AND(X.eq(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),9.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),9.0))) else 0.0)),1.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),4.0))) else 0.0)),2.0)))) else '—')

def f388(E, r, c):
    # Ô K169
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.471, 478): Tranh hợp: một c
    #   an và hai can ngũ hợp với nó (vd 1 Quý 2 Mậu) → "&IF(J169="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J169&") ⇒ "&G169&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.471, 478): Tranh hợp: một can và hai can ngũ hợp với nó (vd 1 Quý 2 Mậu) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f389(E, r, c):
    # Ô E170, F170
    # = IF(MOD(TRUNG_GIAN!$E$7+2,5)=TRUNG_GIAN!$E$9,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(23,9,5))) else '—')

def f390(E, r, c):
    # Ô K170
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.471, 480): Hành can ngày khắc hành chi ngày (cung phối bị khắc) → "&IF(J170="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.",
    #   "ĐỦ ("&J170&") ⇒ "&G170&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.471, 480): Hành can ngày khắc hành chi ngày (cung phối bị khắc) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f391(E, r, c):
    # Ô E171, F171
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f392(E, r, c):
    # Ô K171
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$
    #   Y$20,0)>0,"Tỉ kiên: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.471): Tỉ kiên ≥ P_NHIEU → "&IF(J171="","chư
    #   a đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J171&") ⇒ "&G171&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Tỉ kiên: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.471): Tỉ kiên ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f393(E, r, c):
    # Ô E172, F172
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f394(E, r, c):
    # Ô K172
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5
    #   :$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")
    #   &IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" 
    #   → Quy tắc (P2 tr.471): Quan sát hỗn tạp và thương quan ≥ P_NHIEU → "&IF(J172="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J172&
    #   ") ⇒ "&G172&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.471): Quan sát hỗn tạp và thương quan ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f395(E, r, c):
    # Ô E173, F173
    # = IF((PP1_DINH_TINH!$F$98>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,98,6),E.v(24,99,2))) else '—')

def f396(E, r, c):
    # Ô K173
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr
    #   .471, 476): Tài ≥ P_NHIEU → "&IF(J173="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J173&") ⇒ "&G173&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.471, 476): Tài ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f397(E, r, c):
    # Ô E174, F174
    # = IF((PP1_DINH_TINH!$F$100>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,100,6),E.v(24,99,2))) else '—')

def f398(E, r, c):
    # Ô K174
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 
    #   tr.471, 477, 479): Ấn ≥ P_NHIEU → "&IF(J174="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J174&") ⇒ "&G174&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.471, 477, 479): Ấn ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f399(E, r, c):
    # Ô E175, F175
    # = IF(OR(AND(TRUE(),PP2_DINH_LUONG!$C$43=1,OR(AND(OR(TRUNG_GIAN!$C$12="Chính tài",TRUNG_GIAN!$C$12="Thiên tài"),OR(TRUNG_GIAN!$D$12="Chính quan",TRUNG_GI
    #   AN!$D$12="Thất sát")),AND(OR(TRUNG_GIAN!$C$12="Chính quan",TRUNG_GIAN!$C$12="Thất sát"),OR(TRUNG_GIAN!$D$12="Chính tài",TRUNG_GIAN!$D$12="Thiên tài"))
    #   )),AND(TRUE(),PP2_DINH_LUONG!$D$43=1,OR(AND(OR(TRUNG_GIAN!$D$12="Chính tài",TRUNG_GIAN!$D$12="Thiên tài"),FALSE()),AND(OR(TRUNG_GIAN!$D$12="Chính quan
    #   ",TRUNG_GIAN!$D$12="Thất sát"),FALSE()))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1,OR(AND(FALSE(),OR(TRUNG_GIAN!$F$12="Chính quan
    #   ",TRUNG_GIAN!$F$12="Thất sát")),AND(FALSE(),OR(TRUNG_GIAN!$F$12="Chính tài",TRUNG_GIAN!$F$12="Thiên tài"))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,12,3),'Chính tài'), X.eq(E.v(23,12,3),'Thiên tài')), X.OR(X.eq(E.v(23,12,4),'Chính quan'), X.eq(E.v(23,12,4),'Thất sát'))), X.AND(X.OR(X.eq(E.v(23,12,3),'Chính quan'), X.eq(E.v(23,12,3),'Thất sát')), X.OR(X.eq(E.v(23,12,4),'Chính tài'), X.eq(E.v(23,12,4),'Thiên tài'))))), X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0), X.OR(X.AND(X.OR(X.eq(E.v(23,12,4),'Chính tài'), X.eq(E.v(23,12,4),'Thiên tài')), X.FALSE()), X.AND(X.OR(X.eq(E.v(23,12,4),'Chính quan'), X.eq(E.v(23,12,4),'Thất sát')), X.FALSE()))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0), X.OR(X.AND(X.FALSE(), X.OR(X.eq(E.v(23,12,6),'Chính quan'), X.eq(E.v(23,12,6),'Thất sát'))), X.AND(X.FALSE(), X.OR(X.eq(E.v(23,12,6),'Chính tài'), X.eq(E.v(23,12,6),'Thiên tài'))))))) else '—')

def f400(E, r, c):
    # Ô K175
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.471): Tài và quan sát ở hai can kề nhau thành ngũ hợp → "&IF(J175="","chưa đủ điều kiện ở cả hai phươ
    #   ng pháp – không áp dụng.","ĐỦ ("&J175&") ⇒ "&G175&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.471): Tài và quan sát ở hai can kề nhau thành ngũ hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f401(E, r, c):
    # Ô E176, F176
    # = IF((IF(AND(TRUE(),PP2_DINH_LUONG!$C$43=1),1,0)+IF(AND(TRUE(),PP2_DINH_LUONG!$D$43=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$
    #   43=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$E$6,12)=1),1,0)+I
    #   F(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$E
    #   $6,12)=1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Khôn
    #   g rõ giờ"),MOD(TRUNG_GIAN!$E$6+TRUNG_GIAN!$F$6,12)=1),1,0))>=1,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))) else 0.0)),1.0)) else '—')

def f402(E, r, c):
    # Ô K176
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.471): Có ngũ hợp can hoặc l
    #   ục hợp chi → "&IF(J176="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J176&") ⇒ "&G176&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.471): Có ngũ hợp can hoặc lục hợp chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f403(E, r, c):
    # Ô E177, F177
    # = IF(AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),0.0))) else '—')

def f404(E, r, c):
    # Ô K177
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.467): Trụ ngày Bính Tí → "&IF(J177="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J177&") ⇒ "&G177&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.467): Trụ ngày Bính Tí → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f405(E, r, c):
    # Ô E178, F178
    # = IF(AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=6),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),6.0))) else '—')

def f406(E, r, c):
    # Ô K178
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.467): Trụ ngày Mậu Ngọ → "&IF(J178="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J178&") ⇒ "&G178&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.467): Trụ ngày Mậu Ngọ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f407(E, r, c):
    # Ô E179, F179
    # = IF(OR(TRUNG_GIAN!$E$28=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,28,5),0.0))) else '—')

def f408(E, r, c):
    # Ô K179
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.467): Nhật can lâm Trường sinh tại chi ngày → "&IF(J179="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J179&") ⇒ "&G179&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.467): Nhật can lâm Trường sinh tại chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f409(E, r, c):
    # Ô E180, F180
    # = IF((TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))) else '—')

def f410(E, r, c):
    # Ô K180
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.467): Chi ngày là chi Mộ (kho) của hành tài → "&IF(J180="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J180&") ⇒ "&G
    #   180&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.467): Chi ngày là chi Mộ (kho) của hành tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f411(E, r, c):
    # Ô E181, F181
    # = IF(AND(AND(TRUE(),OR(MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$C$7)),AND(TRUE(),OR(MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_
    #   GIAN!$E$7,MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$D$7)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$F$7,MOD(TRUNG_GIAN
    #   !$F$7+1,5)=TRUNG_GIAN!$E$7)),AND(TRUE(),OR(MOD(TRUNG_GIAN!$C$9+1,5)=TRUNG_GIAN!$D$9,MOD(TRUNG_GIAN!$D$9+1,5)=TRUNG_GIAN!$C$9)),AND(TRUE(),OR(MOD(TRUNG
    #   _GIAN!$D$9+1,5)=TRUNG_GIAN!$E$9,MOD(TRUNG_GIAN!$E$9+1,5)=TRUNG_GIAN!$D$9)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(MOD(TRUNG_GIAN!$E$9+1,5)=TRUNG_GI
    #   AN!$F$9,MOD(TRUNG_GIAN!$F$9+1,5)=TRUNG_GIAN!$E$9)),OR(MOD(TRUNG_GIAN!$E$7+1,5)=TRUNG_GIAN!$E$9,MOD(TRUNG_GIAN!$E$9+1,5)=TRUNG_GIAN!$E$7)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.TRUE(), X.OR(X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,3)))), X.AND(X.TRUE(), X.OR(X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,7,4)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,7,6)), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5)))), X.AND(X.TRUE(), X.OR(X.eq(X.MOD(X.add(E.v(23,9,3),1.0), 5.0),E.v(23,9,4)), X.eq(X.MOD(X.add(E.v(23,9,4),1.0), 5.0),E.v(23,9,3)))), X.AND(X.TRUE(), X.OR(X.eq(X.MOD(X.add(E.v(23,9,4),1.0), 5.0),E.v(23,9,5)), X.eq(X.MOD(X.add(E.v(23,9,5),1.0), 5.0),E.v(23,9,4)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(X.MOD(X.add(E.v(23,9,5),1.0), 5.0),E.v(23,9,6)), X.eq(X.MOD(X.add(E.v(23,9,6),1.0), 5.0),E.v(23,9,5)))), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),E.v(23,9,5)), X.eq(X.MOD(X.add(E.v(23,9,5),1.0), 5.0),E.v(23,7,5))))) else '—')

def f412(E, r, c):
    # Ô K181
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.467): Can kề nhau, chi kề n
    #   hau và can – chi trụ ngày đều tương sinh → "&IF(J181="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J181&") ⇒ "&G181&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.467): Can kề nhau, chi kề nhau và can – chi trụ ngày đều tương sinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f413(E, r, c):
    # Ô E182, F182
    # = IF(AND(OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))), X.ge(E.v(9,99,6),1.0))) else '—')

def f414(E, r, c):
    # Ô K182
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.467): Tài có lực và có quan sát → "&IF(J182="","chưa đủ điều kiện ở cả hai phương pháp
    #    – không áp dụng.","ĐỦ ("&J182&") ⇒ "&G182&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.467): Tài có lực và có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f415(E, r, c):
    # Ô E183, F183
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1))>0)),(PP1_DINH_TINH!$F$98>=3)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,98,6),3.0))) else '—')

def f416(E, r, c):
    # Ô K183
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.467): Có thất sát, có thực thần chế và tài ≥ 3 chữ → "&IF(J183="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J183&") ⇒ "&G183&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.467): Có thất sát, có thực thần chế và tài ≥ 3 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f417(E, r, c):
    # Ô E184, F184
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),1.0)) else '—')

def f418(E, r, c):
    # Ô K184
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0,"Thực thần: "&COUNTIFS(SAO
    #   !$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.467): Chỉ có đúng 1 chữ thực thần → "&IF(J184="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J184&") ⇒ "&G184&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.467): Chỉ có đúng 1 chữ thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f419(E, r, c):
    # Ô E185, F185
    # = IF(AND((TRUNG_GIAN!$E$7=3),OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)))) else '—')

def f420(E, r, c):
    # Ô K185
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&
    #   "Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr
    #   .467): Nhật chủ Kim, sinh mùa đông → "&IF(J185="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J185&") ⇒ "&G185&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.467): Nhật chủ Kim, sinh mùa đông → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f421(E, r, c):
    # Ô E186, F186
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp tà
    #   i",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+C
    #   OUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0), X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0))) else '—')

def f422(E, r, c):
    # Ô K186
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20
    #   ,0)>0,"Tỉ kiên ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,
    #   "Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Tỉ kiên ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Kiếp tài ở năm: "&COUNTI
    #   FS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Kiếp tài ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ;
    #    ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Chính tài ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chín
    #   h tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   B$5:$B$20,3)>0,"Chính tài ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SA
    #   O!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thiên tài ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Thiên
    #    tài ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.467): Tỉ kiếp ở
    #    trụ năm/tháng, tài ở trụ ngày/giờ → "&IF(J186="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J186&") ⇒ "&G186&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Tỉ kiên ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Tỉ kiên ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Chính tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thiên tài ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thiên tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.467): Tỉ kiếp ở trụ năm/tháng, tài ở trụ ngày/giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f423(E, r, c):
    # Ô E187, F187
    # = IF(OR((THAN_SAT!$H$85="Có"),(THAN_SAT!$H$86="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))) else '—')

def f424(E, r, c):
    # Ô K187
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.467): Có thiên/nguyệt đức có hiệu lực (không bị hình xung khắc hại) → "&IF(J187="","chưa đủ điều kiện
    #    ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J187&") ⇒ "&G187&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.467): Có thiên/nguyệt đức có hiệu lực (không bị hình xung khắc hại) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f425(E, r, c):
    # Ô E188, F188
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$
    #   20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),(IF(DOI_CHIEU!$C$33=2,DUNG_
    #   HY_KY!$D$109,DUNG_HY_KY!$C$109)=MOD(TRUNG_GIAN!$E$7+3,5))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—')

def f426(E, r, c):
    # Ô K188
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&
    #   INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Chính quan bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thất sát bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nữ, bản khí chi ngày là quan sát và quan sát là dụng thần → "&IF(J188="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J188&") ⇒ "&G188&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nữ, bản khí chi ngày là quan sát và quan sát là dụng thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f427(E, r, c):
    # Ô E189, F189
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(TRUE(),LEN(THAN_SAT!$E$92)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0)))) else '—')

def f428(E, r, c):
    # Ô K189
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.464): Nữ, chi ngày có trạch mã → "&IF(J189="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J189&") ⇒ "&G189&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.464): Nữ, chi ngày có trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f429(E, r, c):
    # Ô E190, F190
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f430(E, r, c):
    # Ô K190
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5
    #   :$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNT
    #   IFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nữ, tài vượng tại lệnh → "&I
    #   F(J190="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J190&") ⇒ "&G190&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nữ, tài vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f431(E, r, c):
    # Ô E191, F191
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f432(E, r, c):
    # Ô K191
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thực thần vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5
    #   :$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nữ, thực thần vượng tại lệnh → "&IF(J191="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J191&") ⇒ "&G191&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.464): Nữ, thực thần vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f433(E, r, c):
    # Ô E192, F192
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)+COUNTIFS(
    #   SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),0.0))) else '—')

def f434(E, r, c):
    # Ô K192
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)>0,"Thực thần lâm Trường sinh: "&COUNTIFS(SAO!$F
    #   $5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$K$5:$K$20,0)>0,"Thương quan lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,
    #   0)&" chữ; ",""))&" → Quy tắc (P2 tr.464): Nữ, quan sát vượng tại lệnh và thực thương lâm Trường sinh → "&IF(J192="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J192&") ⇒ "&G192&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thực thần lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),0.0)) else '')),(X.cat(X.cat('Thương quan lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.464): Nữ, quan sát vượng tại lệnh và thực thương lâm Trường sinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f435(E, r, c):
    # Ô E193, F193
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99>=1)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,99,6),1.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f436(E, r, c):
    # Ô K193
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.465-466): Nữ, có quan sát và có tài → "&IF(J193="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J193&") ⇒ "&G193&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.465-466): Nữ, có quan sát và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f437(E, r, c):
    # Ô E194, F194
    # = IF(AND(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0))),OR((THAN_SAT!$H$85="
    #   Có"),(THAN_SAT!$H$86="Có"))),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0
    #   ),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sá
    #   t",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))), X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f438(E, r, c):
    # Ô K194
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "
    #   &NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,
    #   "Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J
    #   $5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!
    #   $J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiê
    #   n tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ
    #   ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng
    #    tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SA
    #   O!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.465): Nữ, không có thất sát, có thiên/nguyệt đức, tài và qua
    #   n có lực → "&IF(J194="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J194&") ⇒ "&G194&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, không có thất sát, có thiên/nguyệt đức, tài và quan có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f439(E, r, c):
    # Ô E195, F195
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)+COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0)),0.0))) else '—')

def f440(E, r, c):
    # Ô K195
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)>0,"Chính quan đào hoa: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$
    #   20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)>0,"Thất sát đào hoa: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$S$5:$S$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.465): Nữ, quan sát đóng cùng đào hoa → "&IF(J195
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J195&") ⇒ "&G195&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan đào hoa: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát đào hoa: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,19,20,19), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, quan sát đóng cùng đào hoa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f441(E, r, c):
    # Ô E196, F196
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,
    #   "Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)))) else '—')

def f442(E, r, c):
    # Ô K196
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5
    #   :$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ;
    #    ",""))&" → Quy tắc (P2 tr.465): Nữ, có thương quan, không có chính quan → "&IF(J196="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J196&") ⇒ "&G196&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, có thương quan, không có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f443(E, r, c):
    # Ô E197, F197
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$F$5:$
    #   F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0))) else '—')

def f444(E, r, c):
    # Ô K197
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Chính quan thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thất sát thiên can: "&CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.465): Nữ, quan sát lộ ở can 
    #   → "&IF(J197="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J197&") ⇒ "&G197&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, quan sát lộ ở can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f445(E, r, c):
    # Ô E198, F198
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,
    #   1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính t
    #   ài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f446(E, r, c):
    # Ô K198
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "
    #   &NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J
    #   $5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J
    #   $5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính 
    #   tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; 
    #   ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng t
    #   ại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc 
    #   (P2 tr.465): Nữ, tài có lực → "&IF(J198="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J198&") ⇒ "&G198&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f447(E, r, c):
    # Ô E199, F199
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)
    #   +COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f448(E, r, c):
    # Ô K199
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ",""))&" → Quy tắc (P2 tr.465): Nữ, quan sát (can lộ hoặc bản khí) vượng tại lệnh → "&IF(J199="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J199&") ⇒ "&G199&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, quan sát (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f449(E, r, c):
    # Ô E200, F200
    # = IF(AND(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO
    #   !$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F
    #   $20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),(PP1_DINH_TINH!$F$99>=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thự
    #   c thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0))),"Đ
    #   Ủ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.ge(E.v(9,99,6),1.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)))) else '—')

def f450(E, r, c):
    # Ô K200
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại
    #    có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên ca
    #   n vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNT
    #   IFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNT
    #   IFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính
    #    tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chí
    #   nh tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thực thần bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $L$5:$L$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.465): Nữ, tài có lực, có quan sát, thực thần không bị khắc → "&IF(J200="","chưa đủ điều kiện ở cả hai p
    #   hương pháp – không áp dụng.","ĐỦ ("&J200&") ⇒ "&G200&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Thực thần bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, tài có lực, có quan sát, thực thần không bị khắc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f451(E, r, c):
    # Ô E201, F201
    # = IF(AND(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$100>=1)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0))),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),OR(
    #   (SAO!$O$14=1),(SAO!$O$13=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,100,6),1.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0))), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))), X.OR(X.eq(E.v(13,14,15),1.0), X.eq(E.v(13,13,15),1.0)))) else '—')

def f452(E, r, c):
    # Ô K201
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Thiên/nguyệt 
    #   đức ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$O$14=1,"có","không")&"; "&"Thiên/nguyệt đức ở ngày (can "&TRUNG_GIAN!$E$10&"): "&IF(SAO!$O$13=1,"có",
    #   "không")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L
    #   $20,1)>0,"Chính ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0
    #   ,"Chính ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.465): Nữ, ấn không bị khắc/xung và chi ngày là thiên/nguyệt đức → "&IF(J201="","chưa đủ điều ki
    #   ện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J201&") ⇒ "&G201&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Thiên/nguyệt đức ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,15),1.0)) else 'không')),'; '),'Thiên/nguyệt đức ở ngày (can '),E.v(23,10,5)),'): '),('có' if X.truthy(X.eq(E.v(13,13,15),1.0)) else 'không')),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, ấn không bị khắc/xung và chi ngày là thiên/nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f453(E, r, c):
    # Ô E202, F202
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Thực t
    #   hần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f454(E, r, c):
    # Ô K202
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"T
    #   hực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → 
    #   Quy tắc (P2 tr.465): Nữ, có thất sát và có thực thần chế → "&IF(J202="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J202&") ⇒ "&G
    #   202&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.465): Nữ, có thất sát và có thực thần chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f455(E, r, c):
    # Ô E203, F203
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(TRUNG_GIAN!$E$6=0,TRUNG_GIAN!$E$6=6,TRUNG_GIAN!$E$6=3,TRUNG_GIAN!$E$6=9)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.eq(E.v(23,6,5),0.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,6,5),3.0), X.eq(E.v(23,6,5),9.0)))) else '—')

def f456(E, r, c):
    # Ô K203
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.465): Nữ, chi ngày là Tí, Ngọ, Mão hoặc Dậu → "&IF(J203="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J203&") ⇒ "&G203&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.465): Nữ, chi ngày là Tí, Ngọ, Mão hoặc Dậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f457(E, r, c):
    # Ô E204, F204
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(
    #   SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f458(E, r, c):
    # Ô K204
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại lệnh: "&COUNTIFS(SAO!$F$
    #   5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ
    #   ; ",""))&" → Quy tắc (P2 tr.466): Nữ, quan sát và tài đều vượng tại lệnh → "&IF(J204="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J204&") ⇒ "&G204&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, quan sát và tài đều vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f459(E, r, c):
    # Ô E205, F205
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),0.0))) else '—')

def f460(E, r, c):
    # Ô K205
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)>0,"Thất sát lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$
    #   K$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, thất sát lâm Trường sinh → "&IF(J205="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J205&") ⇒ "&G205&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.466): Nữ, thất sát lâm Trường sinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f461(E, r, c):
    # Ô E206, F206
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$98>=1)),(PP1_DINH_TINH!$F$99>=1)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,98,6),1.0)), X.ge(E.v(9,99,6),1.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f462(E, r, c):
    # Ô K206
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_L
    #   IEU!$C$22&" → Quy tắc (P2 tr.466-467): Nữ, có đủ tài, quan sát, ấn → "&IF(J206="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J20
    #   6&") ⇒ "&G206&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.466-467): Nữ, có đủ tài, quan sát, ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f463(E, r, c):
    # Ô E207, F207
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))), X.eq(E.v(9,99,6),0.0))) else '—')

def f464(E, r, c):
    # Ô K207
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tí
    #   nh: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0,"Thực thần:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, thực thần ≥ P_NHIEU, không có quan sát → "&IF(J207
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J207&") ⇒ "&G207&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.466): Nữ, thực thần ≥ P_NHIEU, không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f465(E, r, c):
    # Ô E208, F208
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),0.0))) else '—')

def f466(E, r, c):
    # Ô K208
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)>0,"Chính ấn thiên/nguyệt đức: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:
    #   $O$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)>0,"Thiên ấn thiên/nguyệt đức: "&COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, ấn đóng cùng thiên/nguyệt 
    #   đức → "&IF(J208="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J208&") ⇒ "&G208&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, ấn đóng cùng thiên/nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f467(E, r, c):
    # Ô E209, F209
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99>=1)),NOT(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)))),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,99,6),1.0)), X.NOT(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)))), X.ge(E.v(9,100,6),1.0))) else '—')

def f468(E, r, c):
    # Ô K209
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC
    #   !$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$
    #   5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ
    #   ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, quan sát không hỗn tạp và có ấn → "&IF(J209="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J209&") ⇒ "&G209&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, quan sát không hỗn tạp và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f469(E, r, c):
    # Ô E210, F210
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,
    #   1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),AND((PP1_DINH_TINH!$F$99<=1),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,
    #   1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Th
    #   ất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.AND(X.le(E.v(9,99,6),1.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))))) else '—')

def f470(E, r, c):
    # Ô K210
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ li
    #   ên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,
    #   1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)
    #   &" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượ
    #   ng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SA
    #   O!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng tại lệnh: "&COUNTIFS(SA
    #   O!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:
    #   $J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất s
    #   át bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",
    #   ""))&" → Quy tắc (P2 tr.466): Nữ, ấn vượng tại lệnh, quan sát yếu → "&IF(J210="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J210
    #   &") ⇒ "&G210&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, ấn vượng tại lệnh, quan sát yếu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f471(E, r, c):
    # Ô E211, F211
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99=0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(9,99,6),0.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f472(E, r, c):
    # Ô K211
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.466): Nữ, không có quan sát nhưng có tài → "&IF(J211="","chư
    #   a đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J211&") ⇒ "&G211&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.466): Nữ, không có quan sát nhưng có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f473(E, r, c):
    # Ô E212, F212
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8))>0)),NOT(OR(((COUNTIF
    #   S(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)), X.NOT(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))))) else '—')

def f474(E, r, c):
    # Ô K212
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thất sát lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng
    #    tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, thất sát
    #    lâm Mộ, thất sát không vượng tại lệnh → "&IF(J212="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J212&") ⇒ "&G212&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thất sát lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, thất sát lâm Mộ, thất sát không vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f475(E, r, c):
    # Ô E213, F213
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,
    #   1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),OR(AND(TRUE(),TRUNG_GIA
    #   N!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIA
    #   N!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,1
    #   0,1,4)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)))))) else '—')

def f476(E, r, c):
    # Ô K213
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các c
    #   hữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$
    #   J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$
    #   20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản 
    #   khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" 
    #   → Quy tắc (P2 tr.478): Nữ, thất sát vượng tại lệnh và chi Mộ (kho) của hành quan có trong trụ → "&IF(J213="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J213&") ⇒ "&G213&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.478): Nữ, thất sát vượng tại lệnh và chi Mộ (kho) của hành quan có trong trụ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f477(E, r, c):
    # Ô E214, F214
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J
    #   $20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),((COUNTIFS(SAO!$F$5
    #   :$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$2
    #   0,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f478(E, r, c):
    # Ô K214
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS
    #   (SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$J$5:$J$20,1)>0,"Chính quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","
    #   "))&" → Quy tắc (P2 tr.466): Nữ, thất sát vượng tại lệnh, chính quan yếu → "&IF(J214="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J214&") ⇒ "&G214&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, thất sát vượng tại lệnh, chính quan yếu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f479(E, r, c):
    # Ô E215, F215
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=MOD(TRUNG_GIAN!$E$7+3,5))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—')

def f480(E, r, c):
    # Ô K215
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&
    #   INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Th
    #   ất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ
    #   ; ",""))&" → Quy tắc (P2 tr.466): Nữ, quan sát vượng tại lệnh và là dụng thần → "&IF(J215="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J215&") ⇒ "&G215&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, quan sát vượng tại lệnh và là dụng thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f481(E, r, c):
    # Ô E216, F216
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99>=1)),AND((PP1_DINH_TINH!$F$99<=1),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))))),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—"
    #   )
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,99,6),1.0)), X.AND(X.le(E.v(9,99,6),1.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))))), X.ge(E.v(9,98,6),1.0))) else '—')

def f482(E, r, c):
    # Ô K216
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$
    #   5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, quan sát yếu và có tài → "&IF(J216="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J216&") ⇒ "&G216&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, quan sát yếu và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f483(E, r, c):
    # Ô E217, F217
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0)),0.0))) else '—')

def f484(E, r, c):
    # Ô K217
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0)>0,"Chính quan không bị xung lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0,SAO!$K$5:$K$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$M$5:$M$20,0,SAO!$K$5:$K$20,0)>0,"Thất sát không bị xung lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $M$5:$M$20,0,SAO!$K$5:$K$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.466): Nữ, quan sát lâm Trường sinh, không bị xung → "&IF(J217="","chưa đủ điều kiện ở 
    #   cả hai phương pháp – không áp dụng.","ĐỦ ("&J217&") ⇒ "&G217&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan không bị xung lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát không bị xung lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,11,20,11), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, quan sát lâm Trường sinh, không bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f485(E, r, c):
    # Ô E218, F218
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.le(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),1.0))) else '—')

def f486(E, r, c):
    # Ô K218
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ
    #   ; ",""))&" → Quy tắc (P2 tr.466): Nữ, quan sát vượng tại lệnh, thương quan ≤ 1 chữ → "&IF(J218="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J218&") ⇒ "&G218&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.466): Nữ, quan sát vượng tại lệnh, thương quan ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f487(E, r, c):
    # Ô E219, F219
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUNG_GIAN!$E$5=2,OR((IF(AND(TRUE(),TRUNG_GIAN!$C$5=8),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=8),1,0)+IF(AND(TR
    #   UE(),TRUNG_GIAN!$E$5=8),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8),1,0))>0,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=9),1,0)+IF(AND(TRUE()
    #   ,TRUNG_GIAN!$D$5=9),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=9),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=9),1,0))>0)),AND(TRUNG_GIAN!$
    #   E$5=0,OR((IF(AND(TRUE(),TRUNG_GIAN!$C$5=6),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=6),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=6),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),TRUNG_GIAN!$F$5=6),1,0))>0,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=7),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=7),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$E$5=
    #   7),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=7),1,0))>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.eq(E.v(23,5,5),2.0), X.OR(X.gt(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),8.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0))) else 0.0)),0.0), X.gt(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),9.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),9.0))) else 0.0)),0.0))), X.AND(X.eq(E.v(23,5,5),0.0), X.OR(X.gt(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))) else 0.0)),0.0), X.gt(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))) else 0.0)),0.0)))))) else '—')

def f488(E, r, c):
    # Ô K219
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.466): Nữ, nhật can Bính có Nhâm/Quý, hoặc nhật can Giáp có Canh/Tân → "&IF(J219="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J219&") ⇒ "&G219&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.466): Nữ, nhật can Bính có Nhâm/Quý, hoặc nhật can Giáp có Canh/Tân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f489(E, r, c):
    # Ô E220, F220
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(IF(AND(TRUE(),PP2_DINH_LUONG!
    #   $C$43=1),1,0)+IF(AND(TRUE(),PP2_DINH_LUONG!$D$43=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1),1,0)+IF(AND(TRUE(),TRUE(),M
    #   OD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$E$6,12)=1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="
    #   Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$E$6,12)=1),1,0)+IF(AND(TRUE(),NOT(
    #   NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),1,0)+IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$E$6+TR
    #   UNG_GIAN!$F$6,12)=1),1,0))>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))) else 0.0)),2.0))) else '—')

def f490(E, r, c):
    # Ô K220
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các c
    #   hữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(
    #   SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.476): Nữ, có thương quan và ≥ 2 cặp hợp → "&IF(J220="
    #   ","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J220&") ⇒ "&G220&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.476): Nữ, có thương quan và ≥ 2 cặp hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f491(E, r, c):
    # Ô E221, F221
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),PP2_DINH_LUONG!$C$43=1,OR(AND(TRUNG_GIAN!$C$5=1,TRUNG_GIAN!$D$5=6),AND(TRUNG_GIAN!$C$5=6,TRUNG_GIAN!
    #   $D$5=1))),AND(TRUE(),PP2_DINH_LUONG!$D$43=1,OR(AND(TRUNG_GIAN!$D$5=1,TRUNG_GIAN!$E$5=6),AND(TRUNG_GIAN!$D$5=6,TRUNG_GIAN!$E$5=1))),AND(NOT(NHAP_LIEU!$
    #   C$9="Không rõ giờ"),PP2_DINH_LUONG!$E$43=1,OR(AND(TRUNG_GIAN!$E$5=1,TRUNG_GIAN!$F$5=6),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$F$5=1))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.eq(E.v(10,43,3),1.0), X.OR(X.AND(X.eq(E.v(23,5,3),1.0), X.eq(E.v(23,5,4),6.0)), X.AND(X.eq(E.v(23,5,3),6.0), X.eq(E.v(23,5,4),1.0)))), X.AND(X.TRUE(), X.eq(E.v(10,43,4),1.0), X.OR(X.AND(X.eq(E.v(23,5,4),1.0), X.eq(E.v(23,5,5),6.0)), X.AND(X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,5,5),1.0)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(10,43,5),1.0), X.OR(X.AND(X.eq(E.v(23,5,5),1.0), X.eq(E.v(23,5,6),6.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,5,6),1.0))))))) else '—')

def f492(E, r, c):
    # Ô K221
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.476): Nữ, Ất và Canh kề nhau thành ngũ hợp → "&IF(J221="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J221&") ⇒ "&G221&".
    #   ")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.476): Nữ, Ất và Canh kề nhau thành ngũ hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f493(E, r, c):
    # Ô E222, F222
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,7
    #   )+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,7)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can",SAO!$K$5:$K$20,9))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$
    #   5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 7.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 7.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 9.0)),0.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0)))) else '—')

def f494(E, r, c):
    # Ô K222
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,7)>0,"Chính quan thiên can lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can",SAO!$K$5:$K$20,9)>0,"Chính quan thiên can lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can",SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20
    #   ,7)>0,"Thất sát thiên can lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,7)&" chữ;
    #    ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,9)>0,"Thất sát thiên can lâm Tuyệt:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20
    #   ,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Chính quan tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"
    #   )>0,"Thất sát tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ",""))&" → Quy tắc (P2 t
    #   r.476): Nữ, quan sát lộ ở can nhưng không tàng ở chi, và lâm Tử/Tuyệt → "&IF(J222="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&
    #   J222&") ⇒ "&G222&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 7.0),0.0)) else ''),(X.cat(X.cat('Chính quan thiên can lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 7.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Chính quan tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')),(X.cat(X.cat('Thất sát tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')))),' → Quy tắc (P2 tr.476): Nữ, quan sát lộ ở can nhưng không tàng ở chi, và lâm Tử/Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f495(E, r, c):
    # Ô E223, F223
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,99,6),E.v(24,99,2)))) else '—')

def f496(E, r, c):
    # Ô K223
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tí
    #   nh: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.477): Nữ, quan sát ≥ P_NHIEU chữ → "&IF(J223="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ
    #    ("&J223&") ⇒ "&G223&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.477): Nữ, quan sát ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f497(E, r, c):
    # Ô E224, F224
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),0.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),0.0))) else '—')

def f498(E, r, c):
    # Ô K224
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Chính tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thiên tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1,SAO!$K$5:$K$20,8)>0,"Chính quan lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF
    #   (COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thất sát lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ",""))&" → Quy tắc (P2 tr.477): Nữ, có tài và quan sát lâm Mộ → "&IF(J224="","chưa đủ điều kiện 
    #   ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J224&") ⇒ "&G224&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else ''),(X.cat(X.cat('Thiên tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Chính quan lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Thất sát lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')))),' → Quy tắc (P2 tr.477): Nữ, có tài và quan sát lâm Mộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f499(E, r, c):
    # Ô E225, F225
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$
    #   20,7)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$
    #   20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,
    #   9))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0)), X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0))) else '—')

def f500(E, r, c):
    # Ô K225
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)>0,"Chính tài lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Chính tài lâm Tuyệt: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$K$5:$K$20,7)>0,"Thiên tài lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)&" chữ; ","")&IF
    #   (COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Thiên tài lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$
    #   K$20,7)>0,"Chính quan lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Chính quan lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)>0,"Th
    #   ất sát lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất 
    #   sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Thất sát lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$K$5:$K$20,9)&" chữ; ",""))&" → Quy tắc (P2 tr.477): Nữ, có tài và quan sát lâm Tử/Tuyệt → "&IF(J225="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J225&") ⇒ "&G225&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),0.0)) else ''),(X.cat(X.cat('Chính tài lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),0.0)) else '')),(X.cat(X.cat('Thiên tài lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Chính quan lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),0.0)) else '')),(X.cat(X.cat('Chính quan lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thất sát lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 7.0),0.0)) else '')),(X.cat(X.cat('Thất sát lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P2 tr.477): Nữ, có tài và quan sát lâm Tử/Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f501(E, r, c):
    # Ô E226, F226
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(TRUNG_GIAN!$E$5=9,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=4),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=4),1,0)+IF(AND(TRUE(),T
    #   RUNG_GIAN!$E$5=4),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=4),1,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.eq(E.v(23,5,5),9.0), X.gt(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),4.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),4.0))) else 0.0)),0.0)))) else '—')

def f502(E, r, c):
    # Ô K226
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.477): Nữ, nhật can Quý và có Mậu → "&IF(J226="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J226&") ⇒ "&G226&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.477): Nữ, nhật can Quý và có Mậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f503(E, r, c):
    # Ô E227, F227
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11),AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=11))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),AND(NOT(NH
    #   AP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=9))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0))))) else '—')

def f504(E, r, c):
    # Ô K227
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.477): Nữ, có cả Hợi và Dậu → "&IF(J227="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J227&") ⇒ "&G227&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.477): Nữ, có cả Hợi và Dậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f505(E, r, c):
    # Ô E228, F228
    # = IF(AND(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$98=0)),(PP1_DINH_TINH!$F$99=0)),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,
    #   0)=0)),(PP1_DINH_TINH!$F$100=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(9,98,6),0.0)), X.eq(E.v(9,99,6),0.0)), X.eq(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)), X.eq(E.v(9,100,6),0.0))) else '—')

def f506(E, r, c):
    # Ô K228
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_L
    #   IEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0,"Thực thần: "&COUNTIFS(
    #   SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.477): Nữ, không có tài, quan sát, thực thần và ấn → "&IF(J228="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J228&") ⇒ "&G228&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.477): Nữ, không có tài, quan sát, thực thần và ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f507(E, r, c):
    # Ô E229, F229
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99<=1)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.le(E.v(9,99,6),1.0)), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f508(E, r, c):
    # Ô K229
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH
    #   _CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.477): Nữ, quan sát
    #    ≤ 1 chữ, tỉ kiếp ≥ P_NHIEU → "&IF(J229="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J229&") ⇒ "&G229&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.477): Nữ, quan sát ≤ 1 chữ, tỉ kiếp ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f509(E, r, c):
    # Ô E230, F230
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$D$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$D$6=0),AND(
    #   TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$D$6=5),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIA
    #   N!$D$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$
    #   C$6=10,TRUNG_GIAN!$D$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=7)
    #   ,AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$D$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6
    #   =11)))),AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=5
    #   ),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6=2,TRU
    #   NG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG
    #   _GIAN!$C$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$
    #   E$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$E$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C
    #   $9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=5),AND(T
    #   RUNG_GIAN!$C$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN
    #   !$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$
    #   C$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$F$6=1),
    #   AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),TRUE(),OR(AND(TRUNG_GIA
    #   N!$D$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=2)
    #   ,AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUN
    #   G_GIAN!$E$6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG
    #   _GIAN!$D$6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$E$6
    #   ,OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$6=
    #   0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=2),AND(TR
    #   UNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!
    #   $F$6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$
    #   D$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$F$6,OR(TRU
    #   NG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=0,TRUNG
    #   _GIAN!$F$6=3),AND(TRUNG_GIAN!$E$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIA
    #   N!$E$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$F$6=2)
    #   ,AND(TRUNG_GIAN!$E$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,T
    #   RUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$E$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN
    #   !$E$6=4,TRUNG_GIAN!$E$6=6,TRUNG_GIAN!$E$6=9,TRUNG_GIAN!$E$6=11)))))),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5
    #   :$N$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,4),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,4),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,4)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,5),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,6,5),9.0), X.eq(E.v(23,6,5),11.0))))))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),0.0))) else '—')

def f510(E, r, c):
    # Ô K230
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các c
    #   hữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Chính qu
    #   an tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thấ
    #   t sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Thất sát tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1,SAO!$N$5:$N$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.477): Nữ, có chi hình và quan sát ở tuần không → "&IF(J230="","chưa đủ điều kiện ở cả hai phươn
    #   g pháp – không áp dụng.","ĐỦ ("&J230&") ⇒ "&G230&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.477): Nữ, có chi hình và quan sát ở tuần không → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f511(E, r, c):
    # Ô E231, F231
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(SAO!$R$14=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(13,14,18),1.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f512(E, r, c):
    # Ô K231
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Kình dương ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$R$14=1,"có","không")&"; "&"Các chữ liên quan (
    #   chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất
    #    sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.477, 479): Nữ, chi ngày là kình dương và có thất sát → "&IF(J231="","chưa đủ
    #    điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J231&") ⇒ "&G231&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Kình dương ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,18),1.0)) else 'không')),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.477, 479): Nữ, chi ngày là kình dương và có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f513(E, r, c):
    # Ô E232, F232
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0)),(COUNTIFS(SAO
    #   !$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f514(E, r, c):
    # Ô K232
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thương quan ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.477): Nữ, trụ tháng có thương quan và thương quan ≥ P_NHIEU → "&IF(J232=""
    #   ,"chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J232&") ⇒ "&G232&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.477): Nữ, trụ tháng có thương quan và thương quan ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f515(E, r, c):
    # Ô E233, F233
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f516(E, r, c):
    # Ô K233
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thương quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, thương quan vượng tại lệnh → "&IF(J233="","chưa đủ điều kiện ở cả hai phương pháp – không á
    #   p dụng.","ĐỦ ("&J233&") ⇒ "&G233&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.478): Nữ, thương quan vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f517(E, r, c):
    # Ô E234, F234
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),(PP1_DINH_TINH!$F$98=0)),(PP1_DINH_TIN
    #   H!$F$100=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))), X.eq(E.v(9,98,6),0.0)), X.eq(E.v(9,100,6),0.0))) else '—')

def f518(E, r, c):
    # Ô K234
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$1
    #   0="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có
    #    mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" c
    #   hữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, thương quan ≥ P_NHIEU, không có tài và ấn → "&IF(J234="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụ
    #   ng.","ĐỦ ("&J234&") ⇒ "&G234&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.478): Nữ, thương quan ≥ P_NHIEU, không có tài và ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f519(E, r, c):
    # Ô E235, F235
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),LEN(THAN_SAT!$E$89)>0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$89)>0))),(PP1_DINH
    #   _TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,89,5)),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,89,6)),0.0)))), X.ge(E.v(9,100,6),1.0))) else '—')

def f520(E, r, c):
    # Ô K235
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.
    #   478): Nữ, hoa cái ở chi ngày hoặc chi giờ và có ấn → "&IF(J235="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J235&") ⇒ "&G235&".
    #   ")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.478): Nữ, hoa cái ở chi ngày hoặc chi giờ và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f521(E, r, c):
    # Ô E236, F236
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"
    #   Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$100=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.eq(E.v(9,100,6),0.0))) else '—')

def f522(E, r, c):
    # Ô K236
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên 
    #   quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$
    #   F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thươn
    #   g quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, có chính quan, có thương
    #    quan, không có ấn → "&IF(J236="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J236&") ⇒ "&G236&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.478): Nữ, có chính quan, có thương quan, không có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f523(E, r, c):
    # Ô E237, F237
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)=0)),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)), X.eq(E.v(9,99,6),0.0))) else '—')

def f524(E, r, c):
    # Ô K237
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ li
    #   ên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần"
    #   ,SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, không có thực thần và không có quan sát → "&IF(J237="","chưa đủ điều kiện ở cả hai phươn
    #   g pháp – không áp dụng.","ĐỦ ("&J237&") ⇒ "&G237&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.478): Nữ, không có thực thần và không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f525(E, r, c):
    # Ô E238, F238
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),(PP1_DINH_TINH!$F$97=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.eq(E.v(9,97,6),0.0))) else '—')

def f526(E, r, c):
    # Ô K238
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ
    #    liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$
    #   J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5
    #   :$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiê
    #   n can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(CO
    #   UNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: 
    #   "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, quan sát vượn
    #   g tại lệnh, không có thực thương chế → "&IF(J238="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J238&") ⇒ "&G238&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.478): Nữ, quan sát vượng tại lệnh, không có thực thương chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f527(E, r, c):
    # Ô E239, F239
    # = IF(AND(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99>=1)),AND((PP1_DINH_TINH!$F$99<=1),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+
    #   COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))))),(PP1_DINH_TINH!$F$97>=P_NHIEU)
    #   ),(PP1_DINH_TINH!$F$98<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,99,6),1.0)), X.AND(X.le(E.v(9,99,6),1.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))))), X.ge(E.v(9,97,6),E.v(24,99,2))), X.le(E.v(9,98,6),1.0))) else '—')

def f528(E, r, c):
    # Ô K239
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_
    #   CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều 
    #   = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính qu
    #   an",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$
    #   5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$
    #   5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, quan sát yếu, thực thương ≥ P_NHIEU, tài ≤ 1 → "&IF(J239="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J239&") ⇒ "&G239&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.478): Nữ, quan sát yếu, thực thương ≥ P_NHIEU, tài ≤ 1 → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f529(E, r, c):
    # Ô E240, F240
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,
    #   2))>0),AND(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1))>0))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0), X.AND(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)))))) else '—')

def f530(E, r, c):
    # Ô K240
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thất sát bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J
    #   $20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J
    #   $20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, bản khí chi ngày là thất sát, hoặc thất sát vượng tại lệnh không có thực thầ
    #   n chế → "&IF(J240="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J240&") ⇒ "&G240&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thất sát bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.478): Nữ, bản khí chi ngày là thất sát, hoặc thất sát vượng tại lệnh không có thực thần chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f531(E, r, c):
    # Ô E241
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$C$7=1)),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,3),1.0)), X.eq(E.v(9,99,6),0.0))) else '—')

def f532(E, r, c):
    # Ô F241
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$D$7=1)),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,4),1.0)), X.eq(E.v(9,99,6),0.0))) else '—')

def f533(E, r, c):
    # Ô K241
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""
    #   ," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.478): Nữ, thân vượng và không có quan sát → "&IF(J241="","chưa đủ điều 
    #   kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J241&") ⇒ "&G241&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.478): Nữ, thân vượng và không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f534(E, r, c):
    # Ô E242, F242
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=8))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),8.0))))) else '—')

def f535(E, r, c):
    # Ô K242
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.478): Nữ, trụ ngày Giáp Dần hoặc Mậu Thân → "&IF(J242="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J242&") ⇒ "&G242&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.478): Nữ, trụ ngày Giáp Dần hoặc Mậu Thân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f536(E, r, c):
    # Ô E243, F243
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$98<=1)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5
    #   :$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.le(E.v(9,98,6),1.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f537(E, r, c):
    # Ô K243
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên qu
    #   an (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính tài vượng tại
    #    lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.478): Nữ, tài ≤ 1 chữ và không vượng tại lệnh → "&IF(J243="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J243&") ⇒ "&G243&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.478): Nữ, tài ≤ 1 chữ và không vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f538(E, r, c):
    # Ô E244, F244
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f539(E, r, c):
    # Ô K244
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.47
    #   8): Nữ, thực thần ≥ P_NHIEU → "&IF(J244="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J244&") ⇒ "&G244&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.478): Nữ, thực thần ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f540(E, r, c):
    # Ô E245, F245
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(TRUNG_GIAN!$E$5=3,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=8),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=8),1,0)+IF(AND(TRUE(),T
    #   RUNG_GIAN!$E$5=8),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8),1,0))>=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.eq(E.v(23,5,5),3.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),8.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),8.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0))) else 0.0)),2.0)))) else '—')

def f541(E, r, c):
    # Ô K245
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.478): Nữ, nhật can Đinh, có ≥ 2 Nhâm → "&IF(J245="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J245&") ⇒ "&G245&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.478): Nữ, nhật can Đinh, có ≥ 2 Nhâm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f542(E, r, c):
    # Ô E246, F246
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),(PP1_DINH_TINH!$F$97<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.le(E.v(9,97,6),1.0))) else '—')

def f543(E, r, c):
    # Ô K246
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ
    #    liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$
    #   J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5
    #   :$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiê
    #   n can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(CO
    #   UNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: 
    #   "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, quan sát vượn
    #   g tại lệnh, thực thương ≤ 1 chữ → "&IF(J246="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J246&") ⇒ "&G246&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.479): Nữ, quan sát vượng tại lệnh, thực thương ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f544(E, r, c):
    # Ô E247, F247
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),AND((PP1_DINH_TINH!$F$99<=1),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)
    #   +COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)), X.AND(X.le(E.v(9,99,6),1.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))))) else '—')

def f545(E, r, c):
    # Ô K247
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ li
    #   ên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thực thần vượn
    #   g tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thương quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J
    #   $20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:
    #   $J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên
    #    can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COU
    #   NTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "
    #   &COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$
    #   F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$
    #   F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, thực thương vư
    #   ợng tại lệnh, quan sát yếu → "&IF(J247="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J247&") ⇒ "&G247&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Thực thần vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.479): Nữ, thực thương vượng tại lệnh, quan sát yếu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f546(E, r, c):
    # Ô E248, F248
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99=1)),(PP1_DINH_TINH!$F$98>=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(9,99,6),1.0)), X.ge(E.v(9,98,6),2.0))) else '—')

def f547(E, r, c):
    # Ô K248
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.479): Nữ, đúng 1 chữ quan sát và ≥ 2 chữ tài → "&IF(J248="",
    #   "chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J248&") ⇒ "&G248&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.479): Nữ, đúng 1 chữ quan sát và ≥ 2 chữ tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f548(E, r, c):
    # Ô E249, F249
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:
    #   $J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Ch
    #   ính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)))),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)+COU
    #   NTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0))) else '—')

def f549(E, r, c):
    # Ô K249
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Chính quan bị khắc/hợp mất: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thất sát bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" c
    #   hữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, quan sát không vượng tại lệnh và bị khắc/hợp → "&IF(J249="","chưa đủ điều kiện ở cả hai phương pháp – không áp
    #    dụng.","ĐỦ ("&J249&") ⇒ "&G249&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.479): Nữ, quan sát không vượng tại lệnh và bị khắc/hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f550(E, r, c):
    # Ô E250, F250
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$100<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,98,6),E.v(24,99,2))), X.le(E.v(9,100,6),1.0))) else '—')

def f551(E, r, c):
    # Ô K250
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$1
    #   0="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.479): Nữ, tài ≥ P_NHIEU, 
    #   ấn ≤ 1 chữ → "&IF(J250="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J250&") ⇒ "&G250&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.479): Nữ, tài ≥ P_NHIEU, ấn ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f552(E, r, c):
    # Ô E251, F251
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),SAO!$P$5=1),AND(TRUE(),SAO!$P$9=1),AND(TRUE(),SAO!$P$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"
    #   ),SAO!$P$17=1))),(THAN_SAT!$H$92="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,9,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0)))), X.eq(E.v(12,92,8),'Có'))) else '—')

def f553(E, r, c):
    # Ô K251
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.479): Nữ, có quý nhân và trạch mã → "&IF(J251="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J251&") ⇒ "&G251&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.479): Nữ, có quý nhân và trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f554(E, r, c):
    # Ô E252, F252
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=0),AND(TRUE(),TRUNG_GIAN!$D$6=0),AND(TRUE(),TRUNG_GIAN!$E$6=0),AND(NOT(N
    #   HAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=0)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT
    #   (NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=6))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=3))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),A
    #   ND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=9)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),0.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0)))))) else '—')

def f555(E, r, c):
    # Ô K252
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.479): Nữ, có đủ Tí, Ngọ, Mão, Dậu → "&IF(J252="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J252&") ⇒ "&G252&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.479): Nữ, có đủ Tí, Ngọ, Mão, Dậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f556(E, r, c):
    # Ô E253, F253
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(N
    #   HAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8),AND(NOT
    #   (NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=8))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=5))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11
    #   ),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=11)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))))) else '—')

def f557(E, r, c):
    # Ô K253
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.479): Nữ, có đủ Dần, Thân, Tị, Hợi → "&IF(J253="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J253&") ⇒ "&G253&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.479): Nữ, có đủ Dần, Thân, Tị, Hợi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f558(E, r, c):
    # Ô E254, F254
    # = IF(AND(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),OR(TRUNG_GIAN!$D$6
    #   =11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2))), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))), X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)))) else '—')

def f559(E, r, c):
    # Ô K254
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HA
    #   NH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&" → Quy tắc (P2 tr.479): Nữ, Kim ≥ P_NHIEU, Thủy ≥ P_NHIEU, sinh mùa đông → "&IF(J
    #   254="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J254&") ⇒ "&G254&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),' → Quy tắc (P2 tr.479): Nữ, Kim ≥ P_NHIEU, Thủy ≥ P_NHIEU, sinh mùa đông → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f560(E, r, c):
    # Ô E255, F255
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$99>=P_NHIEU)),(PP1_DINH_TINH!$F$100>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(E.v(9,99,6),E.v(24,99,2))), X.ge(E.v(9,100,6),E.v(24,99,2)))) else '—')

def f561(E, r, c):
    # Ô K255
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC
    #   !$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.479): Nữ, quan sát ≥
    #    P_NHIEU và ấn ≥ P_NHIEU → "&IF(J255="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J255&") ⇒ "&G255&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.479): Nữ, quan sát ≥ P_NHIEU và ấn ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f562(E, r, c):
    # Ô E256, F256
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(PP1_DINH_TINH!$F$98<=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$2
    #   0,9)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.le(E.v(9,98,6),1.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),0.0))) else '—')

def f563(E, r, c):
    # Ô K256
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên qu
    #   an (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Chính ấn lâm Tuyệt: 
    #   "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,9)>0,"Thiên ấn lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20
    #   ,9)&" chữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, tài ≤ 1 và ấn lâm Tuyệt → "&IF(J256="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("
    #   &J256&") ⇒ "&G256&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P2 tr.479): Nữ, tài ≤ 1 và ấn lâm Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f564(E, r, c):
    # Ô E257, F257
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NO
    #   T(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=5))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8),A
    #   ND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=8))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIA
    #   N!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG
    #   _GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),
    #   TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7))))),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0)))))), X.ge(E.v(9,99,6),1.0))) else '—')

def f565(E, r, c):
    # Ô K257
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "
    #   &TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!
    #   $D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRU
    #   NG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&
    #   " ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.479): Nữ, có tam hình đủ 3 chi và có quan sát → "&
    #   IF(J257="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J257&") ⇒ "&G257&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.479): Nữ, có tam hình đủ 3 chi và có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f566(E, r, c):
    # Ô E258, F258
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8),AND(NOT(N
    #   HAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=8)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=0),AND(TRUE(),TRUNG_GIAN!$D$6=0),AND(TRUE(),TRUNG_GIAN!$E$6=0),AND(NOT
    #   (NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=0))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=4)))),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),0.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))))), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))) else '—')

def f567(E, r, c):
    # Ô K258
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_
    #   TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.479-480): Nữ, có Thân, Tí, Thìn (thủy cục) và Thủy ≥ P_NHIEU → "&IF(J258="","chưa đủ điều kiện ở cả hai 
    #   phương pháp – không áp dụng.","ĐỦ ("&J258&") ⇒ "&G258&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.479-480): Nữ, có Thân, Tí, Thìn (thủy cục) và Thủy ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f568(E, r, c):
    # Ô E259, F259
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GI
    #   AN!$E$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GI
    #   AN!$E$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),
    #   MOD(TRUNG_GIAN!$E$6+TRUNG_GIAN!$F$6,12)=1))),OR(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),A
    #   ND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5)
    #   ,AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=5))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6
    #   =8),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=8))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUN
    #   G_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),
    #   TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TR
    #   UE(),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0)))), X.OR(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))))) else '—')

def f569(E, r, c):
    # Ô K259
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.479): Nữ, có lục hợp và tam hình → "&IF(J259="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J259&") ⇒ "&G259&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.479): Nữ, có lục hợp và tam hình → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f570(E, r, c):
    # Ô E260, F260
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5
    #   :$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0),((COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f571(E, r, c):
    # Ô K260
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Chính tài bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$
    #   5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên tài bị khắc/hợp mất: "&COU
    #   NTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)
    #   &" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$L$5:$L$20,1)>0,"Chính ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")
    #   &IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thi
    #   ên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M
    #   $5:$M$20,1)>0,"Chính ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO
    #   !$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, tài và ấn đều bị khắc/hợp/xung → "&IF(J260="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J260&") ⇒ "&G260&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.479): Nữ, tài và ấn đều bị khắc/hợp/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f572(E, r, c):
    # Ô E261, F261
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f573(E, r, c):
    # Ô K261
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, quan sát vượng tại lệnh và có thương quan → "&IF(J261="","chưa đủ điều kiện ở 
    #   cả hai phương pháp – không áp dụng.","ĐỦ ("&J261&") ⇒ "&G261&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.479): Nữ, quan sát vượng tại lệnh và có thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f574(E, r, c):
    # Ô E262, F262
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0))) else '—')

def f575(E, r, c):
    # Ô K262
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Tỉ kiên thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can")&" chữ; ",""))&" → Quy tắc (P2 tr.479): Nữ, tỉ kiên lộ ở can → "&IF(J262="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J26
    #   2&") ⇒ "&G262&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Tỉ kiên thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''))),' → Quy tắc (P2 tr.479): Nữ, tỉ kiên lộ ở can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f576(E, r, c):
    # Ô E263, F263
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0)),"ĐỦ",
    #   "—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f577(E, r, c):
    # Ô K263
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",S
    #   AO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản kh
    #   í",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0
    #   ,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên ấn vượng tại lệnh: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.480): Nữ, quan sát vượng tại lệnh và thiên ấn vư
    #   ợng tại lệnh → "&IF(J263="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J263&") ⇒ "&G263&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.480): Nữ, quan sát vượng tại lệnh và thiên ấn vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f578(E, r, c):
    # Ô E264, F264
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(INDEX(M_XU,TRUNG_GIAN!$C$6+1,TRUNG_GIAN!$E$6+1)=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,3),1.0), X.add(E.v(23,6,5),1.0)),1.0))) else '—')

def f579(E, r, c):
    # Ô K264
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GI
    #   AN!$E$9+1)&")"&" → Quy tắc (P2 tr.480): Nữ, chi năm xung chi ngày → "&IF(J264="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J264
    #   &") ⇒ "&G264&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, chi năm xung chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f580(E, r, c):
    # Ô E265, F265
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(THAN_SAT!$H$91="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(12,91,8),'Có'))) else '—')

def f581(E, r, c):
    # Ô K265
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.480): Nữ, đào hoa có hiệu lực → "&IF(J265="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J265&") ⇒ "&G265&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.480): Nữ, đào hoa có hiệu lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f582(E, r, c):
    # Ô E266
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$C$7=1)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)=MOD(TRUNG_GIAN!$E$7+3,5))),"ĐỦ
    #   ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,3),1.0)), X.eq((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—')

def f583(E, r, c):
    # Ô F266
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$D$7=1)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)=MOD(TRUNG_GIAN!$E$7+3,5))),"ĐỦ
    #   ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,4),1.0)), X.eq((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '—')

def f584(E, r, c):
    # Ô K266
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Kỵ PP1: "&IF(DUNG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$
    #   C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_G
    #   IAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.480): Nữ, thân vượn
    #   g và quan sát là kỵ thần → "&IF(J266="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J266&") ⇒ "&G266&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, thân vượng và quan sát là kỵ thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f585(E, r, c):
    # Ô E267, F267
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(IF(OR(OR(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND
    #   (NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),A
    #   ND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=5))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8
    #   ),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=8))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_
    #   GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TR
    #   UNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE
    #   (),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7)))),OR(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),AND
    #   (TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$E$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),AND
    #   (TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$E$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),AND
    #   (TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$E$6+TRUNG_GIAN!$F$6,12)=1))),1,0)+IF(OR(THAN_SAT!$H$122="Có",THAN_SAT!$H$90="Có"),1,0)+IF(O
    #   R(THAN_SAT!$H$94="Có",THAN_SAT!$H$95="Có"),1,0))>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.add(X.add((1.0 if X.truthy(X.OR(X.OR(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))), X.OR(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))))) else 0.0),(1.0 if X.truthy(X.OR(X.eq(E.v(12,122,8),'Có'), X.eq(E.v(12,90,8),'Có'))) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(E.v(12,94,8),'Có'), X.eq(E.v(12,95,8),'Có'))) else 0.0)),2.0))) else '—')

def f586(E, r, c):
    # Ô K267
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.480): Nữ, có ≥ 2 trong 3 nhóm: tam hình/lục hợp, vong thần/kiếp sát, cô thần/quả tú → "&IF(J267="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J267&") ⇒ "&G267&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, có ≥ 2 trong 3 nhóm: tam hình/lục hợp, vong thần/kiếp sát, cô thần/quả tú → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f587(E, r, c):
    # Ô E268, F268
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,7)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,9)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5
    #   :$K$20,7)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,9))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.add(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 7.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 8.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 9.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 7.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 8.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 9.0)),0.0))) else '—')

def f588(E, r, c):
    # Ô K268
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,7)>0,"Chính quan ở ngày lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,
    #   SAO!$K$5:$K$20,8)>0,"Chính quan ở ngày lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$2
    #   0,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,9)>0,"Chính quan ở ngày lâ
    #   m Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,9)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5
    #   :$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,7)>0,"Thất sát ở ngày lâm Tử: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,7)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,8)>0,"Thất sát ở ngày lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,2,SAO!$K$5:$K$20,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,9)>0,"Thất
    #    sát ở ngày lâm Tuyệt: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2,SAO!$K$5:$K$20,9)&" chữ; ",""))&" → Quy
    #    tắc (P2 tr.480): Nữ, quan sát ở chi ngày lâm Tử/Mộ/Tuyệt → "&IF(J268="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J268&") ⇒ "&
    #   G268&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan ở ngày lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 7.0),0.0)) else ''),(X.cat(X.cat('Chính quan ở ngày lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Chính quan ở ngày lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở ngày lâm Tử: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 7.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 7.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở ngày lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở ngày lâm Tuyệt: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 9.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0, E.rg(13,5,11,20,11), 9.0),0.0)) else '')))),' → Quy tắc (P2 tr.480): Nữ, quan sát ở chi ngày lâm Tử/Mộ/Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f589(E, r, c):
    # Ô E269
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)+COUNTIFS(SAO!$F$5:$
    #   F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0)),(DUNG_HY_KY!$C$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),0.0)), X.eq(E.v(8,7,3),0.0))) else '—')

def f590(E, r, c):
    # Ô F269
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)+COUNTIFS(SAO!$F$5:$
    #   F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0)),(DUNG_HY_KY!$D$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),0.0)), X.eq(E.v(8,7,4),0.0))) else '—')

def f591(E, r, c):
    # Ô K269
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệ
    #   t kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Chính quan tuần không: "&COUN
    #   TIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Thất sát tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)
    #   &" chữ; ",""))&" → Quy tắc (P2 tr.480): Nữ, có quan sát ở tuần không và thân nhược → "&IF(J269="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J269&") ⇒ "&G269&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.480): Nữ, có quan sát ở tuần không và thân nhược → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f592(E, r, c):
    # Ô E270, F270
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(THAN_SAT!$H$111="Có")),OR(AND(TRUE(),TRUNG_GIAN!$C$5=0),AND(TRUE(),TRUNG_GIAN!$D$5=0),AND(TRUE(),TRUNG_GIAN!$
    #   E$5=0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(12,111,8),'Có')), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),0.0))))) else '—')

def f593(E, r, c):
    # Ô K270
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.480): Nữ, có kim thần và có Giáp → "&IF(J270="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J270&") ⇒ "&G270&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, có kim thần và có Giáp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f594(E, r, c):
    # Ô E271, F271
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(THAN_SAT!$H$91="Có")),(THAN_SAT!$H$90="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(12,91,8),'Có')), X.eq(E.v(12,90,8),'Có'))) else '—')

def f595(E, r, c):
    # Ô K271
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.480): Nữ, đào hoa và kiếp sát đều có hiệu lực → "&IF(J271="","chưa đủ điều kiện ở cả hai 
    #   phương pháp – không áp dụng.","ĐỦ ("&J271&") ⇒ "&G271&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.480): Nữ, đào hoa và kiếp sát đều có hiệu lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f596(E, r, c):
    # Ô E272, F272
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(AND(TRUE(),LEN(THAN_SAT!$C$91)>0),AND(TRUE(),LEN(THAN_SAT!$C$123)>0)),AND(AND(TRUE(),LEN(THAN_SAT!$D$91)>0)
    #   ,AND(TRUE(),LEN(THAN_SAT!$D$123)>0)),AND(AND(TRUE(),LEN(THAN_SAT!$E$91)>0),AND(TRUE(),LEN(THAN_SAT!$E$123)>0)),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ gi
    #   ờ"),LEN(THAN_SAT!$F$91)>0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$123)>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,91,3)),0.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,123,3)),0.0))), X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,91,4)),0.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,123,4)),0.0))), X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,91,5)),0.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,123,5)),0.0))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,91,6)),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,123,6)),0.0)))))) else '—')

def f597(E, r, c):
    # Ô K272
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.480): Nữ, đào hoa và nguyên thần ở cùng một trụ → "&IF(J272="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J272&") ⇒ "&G272&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.480): Nữ, đào hoa và nguyên thần ở cùng một trụ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f598(E, r, c):
    # Ô E273, F273
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(TRUNG_GIAN!$E$5=0,(IF(AND(TRUE(),TRUNG_GIAN!$C$5=7),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$5=7),1,0)+IF(AND(TRUE
    #   (),TRUNG_GIAN!$E$5=7),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=7),1,0))>=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_G
    #   IAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.eq(E.v(23,5,5),0.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))) else 0.0)),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))))) else '—')

def f599(E, r, c):
    # Ô K273
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.480): Nữ, nhật can Giáp, có ≥ 2 Tân và có Sửu → "&IF(J273="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J273&") ⇒ "&G273
    #   &".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, nhật can Giáp, có ≥ 2 Tân và có Sửu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f600(E, r, c):
    # Ô E274, F274
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN
    #   !$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=10))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$F$5=8,TRUNG_GIAN!$F$6
    #   =4),AND(TRUNG_GIAN!$F$5=6,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$F$5=6,TRUNG_GIAN!$F$6=4),AND(TRUNG_GIAN!$F$5=4,TRUNG_GIAN!$F$6=10))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),10.0)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,6),8.0), X.eq(E.v(23,6,6),4.0)), X.AND(X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,6,6),4.0)), X.AND(X.eq(E.v(23,5,6),4.0), X.eq(E.v(23,6,6),10.0))))))) else '—')

def f601(E, r, c):
    # Ô K274
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GI
    #   AN!$F$9+1)&")"&" → Quy tắc (P2 tr.480): Nữ, chi ngày hoặc chi giờ là khôi canh (trụ ngày/giờ Nhâm Thìn, Canh Tuất, Canh Thìn, Mậu Tuất) → "&IF(J274=""
    #   ,"chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J274&") ⇒ "&G274&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, chi ngày hoặc chi giờ là khôi canh (trụ ngày/giờ Nhâm Thìn, Canh Tuất, Canh Thìn, Mậu Tuất) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f602(E, r, c):
    # Ô E275, F275
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),(TRUNG_GIAN!$E$7=4)),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(23,7,5),4.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)))) else '—')

def f603(E, r, c):
    # Ô K275
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.480): Nữ, nhật chủ Thủy và Thổ ≥ P_NHIEU → "&IF(J275="","chưa đủ
    #    điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J275&") ⇒ "&G275&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.480): Nữ, nhật chủ Thủy và Thổ ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f604(E, r, c):
    # Ô E280, F280
    # = IF(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$28=5),AND(TRUE(),TRUNG_GIAN!$D$28=5),AND(TRUE(),TRUNG_GIAN!$E$28=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GI
    #   AN!$F$28=5)),OR(AND(TRUE(),TRUNG_GIAN!$C$28=7),AND(TRUE(),TRUNG_GIAN!$D$28=7),AND(TRUE(),TRUNG_GIAN!$E$28=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TR
    #   UNG_GIAN!$F$28=7))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,28,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,28,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,28,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,28,6),5.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,28,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,28,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,28,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,28,6),7.0))))) else '—')

def f605(E, r, c):
    # Ô K280
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.470; UH tr.280): Nhật can ở cung Suy tại một chi và ở cung Tử tại chi khác → "&IF(J280="","chưa đủ đi
    #   ều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J280&") ⇒ "&G280&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.470; UH tr.280): Nhật can ở cung Suy tại một chi và ở cung Tử tại chi khác → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f606(E, r, c):
    # Ô E281
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$C$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,3),0.0))) else '—')

def f607(E, r, c):
    # Ô F281
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$D$7=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,4),0.0))) else '—')

def f608(E, r, c):
    # Ô K281
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.465; UH tr.128): Nữ, thân nhược → "&IF(J281="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J281&") ⇒ "&G281&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.465; UH tr.128): Nữ, thân nhược → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

def f609(E, r, c):
    # Ô E282
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,3),1.0))) else '—')

def f610(E, r, c):
    # Ô F282
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.eq(E.v(8,7,4),1.0))) else '—')

def f611(E, r, c):
    # Ô K282
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.465; UH tr.128): Nữ, thân vượng → "&IF(J282="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J282&") ⇒ "&G282&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.465; UH tr.128): Nữ, thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(17,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(17,r,c-1)),') ⇒ '),E.v(17,r,c-4)),'.')))

CELLS = {
    'D5': (f0, 5, 4),
    'E5': (f1, 5, 5),
    'D6': (f2, 6, 4),
    'E6': (f3, 6, 5),
    'D7': (f4, 7, 4),
    'C11': (f5, 11, 3),
    'D11': (f6, 11, 4),
    'E11': (f7, 11, 5),
    'F11': (f8, 11, 6),
    'G11': (f9, 11, 7),
    'H11': (f10, 11, 8),
    'I11': (f11, 11, 9),
    'J11': (f12, 11, 10),
    'K11': (f13, 11, 11),
    'C12': (f14, 12, 3),
    'D12': (f15, 12, 4),
    'E12': (f16, 12, 5),
    'F12': (f7, 12, 6),
    'G12': (f17, 12, 7),
    'H12': (f18, 12, 8),
    'I12': (f10, 12, 9),
    'J12': (f11, 12, 10),
    'K12': (f13, 12, 11),
    'C13': (f19, 13, 3),
    'D13': (f20, 13, 4),
    'E13': (f21, 13, 5),
    'F13': (f16, 13, 6),
    'G13': (f22, 13, 7),
    'H13': (f23, 13, 8),
    'I13': (f18, 13, 9),
    'J13': (f10, 13, 10),
    'K13': (f13, 13, 11),
    'C14': (f24, 14, 3),
    'D14': (f25, 14, 4),
    'E14': (f26, 14, 5),
    'F14': (f21, 14, 6),
    'G14': (f27, 14, 7),
    'H14': (f28, 14, 8),
    'I14': (f23, 14, 9),
    'J14': (f18, 14, 10),
    'K14': (f13, 14, 11),
    'C15': (f29, 15, 3),
    'D15': (f30, 15, 4),
    'E15': (f31, 15, 5),
    'F15': (f26, 15, 6),
    'G15': (f32, 15, 7),
    'H15': (f33, 15, 8),
    'I15': (f28, 15, 9),
    'J15': (f23, 15, 10),
    'K15': (f13, 15, 11),
    'C16': (f34, 16, 3),
    'D16': (f35, 16, 4),
    'E16': (f36, 16, 5),
    'F16': (f31, 16, 6),
    'G16': (f37, 16, 7),
    'H16': (f38, 16, 8),
    'I16': (f33, 16, 9),
    'J16': (f28, 16, 10),
    'K16': (f13, 16, 11),
    'C17': (f39, 17, 3),
    'D17': (f40, 17, 4),
    'E17': (f41, 17, 5),
    'F17': (f36, 17, 6),
    'G17': (f42, 17, 7),
    'H17': (f43, 17, 8),
    'I17': (f38, 17, 9),
    'J17': (f33, 17, 10),
    'K17': (f13, 17, 11),
    'C18': (f44, 18, 3),
    'D18': (f45, 18, 4),
    'E18': (f46, 18, 5),
    'F18': (f41, 18, 6),
    'G18': (f47, 18, 7),
    'H18': (f48, 18, 8),
    'I18': (f43, 18, 9),
    'J18': (f38, 18, 10),
    'K18': (f13, 18, 11),
    'C22': (f49, 22, 3),
    'D22': (f50, 22, 4),
    'E22': (f51, 22, 5),
    'G22': (f52, 22, 7),
    'K22': (f53, 22, 11),
    'C23': (f54, 23, 3),
    'D23': (f55, 23, 4),
    'E23': (f56, 23, 5),
    'G23': (f57, 23, 7),
    'K23': (f58, 23, 11),
    'C24': (f59, 24, 3),
    'D24': (f60, 24, 4),
    'E24': (f61, 24, 5),
    'G24': (f62, 24, 7),
    'K24': (f63, 24, 11),
    'C25': (f64, 25, 3),
    'D25': (f65, 25, 4),
    'E25': (f66, 25, 5),
    'G25': (f67, 25, 7),
    'K25': (f68, 25, 11),
    'C26': (f69, 26, 3),
    'D26': (f70, 26, 4),
    'E26': (f71, 26, 5),
    'G26': (f72, 26, 7),
    'K26': (f73, 26, 11),
    'C27': (f74, 27, 3),
    'D27': (f75, 27, 4),
    'E27': (f76, 27, 5),
    'G27': (f77, 27, 7),
    'K27': (f78, 27, 11),
    'C28': (f79, 28, 3),
    'D28': (f80, 28, 4),
    'E28': (f81, 28, 5),
    'G28': (f82, 28, 7),
    'K28': (f83, 28, 11),
    'C29': (f84, 29, 3),
    'D29': (f85, 29, 4),
    'E29': (f86, 29, 5),
    'G29': (f87, 29, 7),
    'K29': (f88, 29, 11),
    'C30': (f89, 30, 3),
    'D30': (f90, 30, 4),
    'E30': (f91, 30, 5),
    'G30': (f92, 30, 7),
    'K30': (f93, 30, 11),
    'C31': (f94, 31, 3),
    'D31': (f95, 31, 4),
    'E31': (f96, 31, 5),
    'G31': (f97, 31, 7),
    'K31': (f98, 31, 11),
    'C35': (f99, 35, 3),
    'D35': (f100, 35, 4),
    'E35': (f101, 35, 5),
    'G35': (f102, 35, 7),
    'C36': (f103, 36, 3),
    'D36': (f104, 36, 4),
    'E36': (f105, 36, 5),
    'G36': (f106, 36, 7),
    'C37': (f107, 37, 3),
    'D37': (f108, 37, 4),
    'E37': (f109, 37, 5),
    'G37': (f110, 37, 7),
    'C38': (f111, 38, 3),
    'D38': (f112, 38, 4),
    'E38': (f113, 38, 5),
    'G38': (f114, 38, 7),
    'C39': (f115, 39, 3),
    'D39': (f116, 39, 4),
    'E39': (f117, 39, 5),
    'G39': (f118, 39, 7),
    'C40': (f119, 40, 3),
    'D40': (f120, 40, 4),
    'E40': (f121, 40, 5),
    'G40': (f122, 40, 7),
    'C41': (f123, 41, 3),
    'D41': (f124, 41, 4),
    'E41': (f125, 41, 5),
    'G41': (f126, 41, 7),
    'C42': (f127, 42, 3),
    'D42': (f128, 42, 4),
    'E42': (f129, 42, 5),
    'G42': (f130, 42, 7),
    'C43': (f131, 43, 3),
    'D43': (f132, 43, 4),
    'E43': (f133, 43, 5),
    'G43': (f134, 43, 7),
    'C44': (f135, 44, 3),
    'D44': (f136, 44, 4),
    'E44': (f137, 44, 5),
    'G44': (f138, 44, 7),
    'D47': (f139, 47, 4),
    'G47': (f140, 47, 7),
    'D48': (f141, 48, 4),
    'D51': (f142, 51, 4),
    'D52': (f143, 52, 4),
    'F52': (f144, 52, 6),
    'M52': (f145, 52, 13),
    'N52': (f146, 52, 14),
    'O52': (f147, 52, 15),
    'P52': (f148, 52, 16),
    'Q52': (f149, 52, 17),
    'R52': (f150, 52, 18),
    'D53': (f151, 53, 4),
    'D54': (f152, 54, 4),
    'D55': (f153, 55, 4),
    'D56': (f154, 56, 4),
    'D57': (f155, 57, 4),
    'D58': (f156, 58, 4),
    'D59': (f157, 59, 4),
    'D64': (f158, 64, 4),
    'G64': (f159, 64, 7),
    'H64': (f160, 64, 8),
    'K64': (f161, 64, 11),
    'M64': (f162, 64, 13),
    'N64': (f163, 64, 14),
    'O64': (f164, 64, 15),
    'P64': (f165, 64, 16),
    'Q64': (f166, 64, 17),
    'R64': (f167, 64, 18),
    'S64': (f168, 64, 19),
    'T64': (f169, 64, 20),
    'U64': (f170, 64, 21),
    'V64': (f171, 64, 22),
    'W64': (f172, 64, 23),
    'X64': (f173, 64, 24),
    'Y64': (f174, 64, 25),
    'Z64': (f175, 64, 26),
    'AA64': (f176, 64, 27),
    'AB64': (f177, 64, 28),
    'AC64': (f178, 64, 29),
    'AD64': (f179, 64, 30),
    'AF64': (f180, 64, 32),
    'AG64': (f181, 64, 33),
    'AH64': (f182, 64, 34),
    'AI64': (f183, 64, 35),
    'AJ64': (f184, 64, 36),
    'AK64': (f185, 64, 37),
    'AL64': (f186, 64, 38),
    'AM64': (f187, 64, 39),
    'D66': (f188, 66, 4),
    'D67': (f189, 67, 4),
    'E72': (f190, 72, 5),
    'F72': (f190, 72, 6),
    'J72': (f191, 72, 10),
    'K72': (f192, 72, 11),
    'E73': (f193, 73, 5),
    'F73': (f193, 73, 6),
    'J73': (f191, 73, 10),
    'K73': (f194, 73, 11),
    'E74': (f195, 74, 5),
    'F74': (f196, 74, 6),
    'J74': (f191, 74, 10),
    'K74': (f197, 74, 11),
    'E75': (f198, 75, 5),
    'F75': (f198, 75, 6),
    'J75': (f191, 75, 10),
    'K75': (f199, 75, 11),
    'E76': (f200, 76, 5),
    'F76': (f201, 76, 6),
    'J76': (f191, 76, 10),
    'K76': (f202, 76, 11),
    'E77': (f203, 77, 5),
    'F77': (f204, 77, 6),
    'J77': (f191, 77, 10),
    'K77': (f205, 77, 11),
    'E78': (f206, 78, 5),
    'F78': (f206, 78, 6),
    'J78': (f191, 78, 10),
    'K78': (f207, 78, 11),
    'E79': (f208, 79, 5),
    'F79': (f208, 79, 6),
    'J79': (f191, 79, 10),
    'K79': (f209, 79, 11),
    'E80': (f210, 80, 5),
    'F80': (f210, 80, 6),
    'J80': (f191, 80, 10),
    'K80': (f211, 80, 11),
    'E81': (f212, 81, 5),
    'F81': (f212, 81, 6),
    'J81': (f191, 81, 10),
    'K81': (f213, 81, 11),
    'E82': (f214, 82, 5),
    'F82': (f214, 82, 6),
    'J82': (f191, 82, 10),
    'K82': (f215, 82, 11),
    'E83': (f216, 83, 5),
    'F83': (f216, 83, 6),
    'G83': (f217, 83, 7),
    'J83': (f191, 83, 10),
    'K83': (f218, 83, 11),
    'E84': (f219, 84, 5),
    'F84': (f219, 84, 6),
    'J84': (f191, 84, 10),
    'K84': (f220, 84, 11),
    'E85': (f221, 85, 5),
    'F85': (f221, 85, 6),
    'J85': (f191, 85, 10),
    'K85': (f222, 85, 11),
    'E86': (f223, 86, 5),
    'F86': (f223, 86, 6),
    'J86': (f191, 86, 10),
    'K86': (f224, 86, 11),
    'E87': (f225, 87, 5),
    'F87': (f225, 87, 6),
    'J87': (f191, 87, 10),
    'K87': (f226, 87, 11),
    'E88': (f227, 88, 5),
    'F88': (f227, 88, 6),
    'J88': (f191, 88, 10),
    'K88': (f228, 88, 11),
    'E89': (f229, 89, 5),
    'F89': (f229, 89, 6),
    'J89': (f191, 89, 10),
    'K89': (f230, 89, 11),
    'E90': (f231, 90, 5),
    'F90': (f231, 90, 6),
    'J90': (f191, 90, 10),
    'K90': (f232, 90, 11),
    'E91': (f233, 91, 5),
    'F91': (f233, 91, 6),
    'J91': (f191, 91, 10),
    'K91': (f234, 91, 11),
    'E92': (f235, 92, 5),
    'F92': (f235, 92, 6),
    'J92': (f191, 92, 10),
    'K92': (f236, 92, 11),
    'E93': (f237, 93, 5),
    'F93': (f237, 93, 6),
    'J93': (f191, 93, 10),
    'K93': (f238, 93, 11),
    'E94': (f239, 94, 5),
    'F94': (f239, 94, 6),
    'J94': (f191, 94, 10),
    'K94': (f240, 94, 11),
    'E95': (f241, 95, 5),
    'F95': (f241, 95, 6),
    'J95': (f191, 95, 10),
    'K95': (f242, 95, 11),
    'E96': (f243, 96, 5),
    'F96': (f243, 96, 6),
    'J96': (f191, 96, 10),
    'K96': (f244, 96, 11),
    'E97': (f245, 97, 5),
    'F97': (f245, 97, 6),
    'J97': (f191, 97, 10),
    'K97': (f246, 97, 11),
    'E98': (f247, 98, 5),
    'F98': (f247, 98, 6),
    'J98': (f191, 98, 10),
    'K98': (f248, 98, 11),
    'E99': (f249, 99, 5),
    'F99': (f250, 99, 6),
    'J99': (f191, 99, 10),
    'K99': (f251, 99, 11),
    'E100': (f252, 100, 5),
    'F100': (f252, 100, 6),
    'J100': (f191, 100, 10),
    'K100': (f253, 100, 11),
    'E101': (f254, 101, 5),
    'F101': (f254, 101, 6),
    'J101': (f191, 101, 10),
    'K101': (f255, 101, 11),
    'E102': (f256, 102, 5),
    'F102': (f256, 102, 6),
    'J102': (f191, 102, 10),
    'K102': (f257, 102, 11),
    'E103': (f258, 103, 5),
    'F103': (f258, 103, 6),
    'J103': (f191, 103, 10),
    'K103': (f259, 103, 11),
    'E104': (f260, 104, 5),
    'F104': (f261, 104, 6),
    'J104': (f191, 104, 10),
    'K104': (f262, 104, 11),
    'E105': (f263, 105, 5),
    'F105': (f263, 105, 6),
    'J105': (f191, 105, 10),
    'K105': (f264, 105, 11),
    'E106': (f265, 106, 5),
    'F106': (f266, 106, 6),
    'J106': (f191, 106, 10),
    'K106': (f267, 106, 11),
    'E107': (f268, 107, 5),
    'F107': (f269, 107, 6),
    'J107': (f191, 107, 10),
    'K107': (f270, 107, 11),
    'E112': (f271, 112, 5),
    'F112': (f271, 112, 6),
    'J112': (f191, 112, 10),
    'K112': (f272, 112, 11),
    'L112': (f273, 112, 12),
    'E113': (f274, 113, 5),
    'F113': (f274, 113, 6),
    'J113': (f191, 113, 10),
    'K113': (f275, 113, 11),
    'L113': (f273, 113, 12),
    'E114': (f276, 114, 5),
    'F114': (f276, 114, 6),
    'J114': (f191, 114, 10),
    'K114': (f277, 114, 11),
    'L114': (f273, 114, 12),
    'E115': (f278, 115, 5),
    'F115': (f278, 115, 6),
    'J115': (f191, 115, 10),
    'K115': (f279, 115, 11),
    'L115': (f273, 115, 12),
    'E116': (f280, 116, 5),
    'F116': (f280, 116, 6),
    'J116': (f191, 116, 10),
    'K116': (f281, 116, 11),
    'L116': (f273, 116, 12),
    'E117': (f282, 117, 5),
    'F117': (f282, 117, 6),
    'J117': (f191, 117, 10),
    'K117': (f283, 117, 11),
    'L117': (f273, 117, 12),
    'E118': (f284, 118, 5),
    'F118': (f284, 118, 6),
    'J118': (f191, 118, 10),
    'K118': (f285, 118, 11),
    'L118': (f273, 118, 12),
    'E119': (f286, 119, 5),
    'F119': (f287, 119, 6),
    'J119': (f191, 119, 10),
    'K119': (f288, 119, 11),
    'L119': (f273, 119, 12),
    'E120': (f289, 120, 5),
    'F120': (f289, 120, 6),
    'J120': (f191, 120, 10),
    'K120': (f290, 120, 11),
    'L120': (f273, 120, 12),
    'E121': (f291, 121, 5),
    'F121': (f291, 121, 6),
    'J121': (f191, 121, 10),
    'K121': (f292, 121, 11),
    'L121': (f273, 121, 12),
    'E122': (f293, 122, 5),
    'F122': (f293, 122, 6),
    'J122': (f191, 122, 10),
    'K122': (f294, 122, 11),
    'L122': (f273, 122, 12),
    'E123': (f295, 123, 5),
    'F123': (f295, 123, 6),
    'J123': (f191, 123, 10),
    'K123': (f296, 123, 11),
    'L123': (f273, 123, 12),
    'E124': (f297, 124, 5),
    'F124': (f297, 124, 6),
    'J124': (f191, 124, 10),
    'K124': (f298, 124, 11),
    'L124': (f273, 124, 12),
    'E125': (f299, 125, 5),
    'F125': (f299, 125, 6),
    'J125': (f191, 125, 10),
    'K125': (f300, 125, 11),
    'L125': (f273, 125, 12),
    'E126': (f301, 126, 5),
    'F126': (f301, 126, 6),
    'J126': (f191, 126, 10),
    'K126': (f302, 126, 11),
    'L126': (f273, 126, 12),
    'E127': (f260, 127, 5),
    'F127': (f261, 127, 6),
    'J127': (f191, 127, 10),
    'K127': (f303, 127, 11),
    'L127': (f273, 127, 12),
    'E128': (f304, 128, 5),
    'F128': (f304, 128, 6),
    'J128': (f191, 128, 10),
    'K128': (f305, 128, 11),
    'L128': (f273, 128, 12),
    'E129': (f306, 129, 5),
    'F129': (f306, 129, 6),
    'J129': (f191, 129, 10),
    'K129': (f307, 129, 11),
    'L129': (f273, 129, 12),
    'E130': (f308, 130, 5),
    'F130': (f308, 130, 6),
    'J130': (f191, 130, 10),
    'K130': (f309, 130, 11),
    'L130': (f273, 130, 12),
    'E131': (f310, 131, 5),
    'F131': (f310, 131, 6),
    'J131': (f191, 131, 10),
    'K131': (f311, 131, 11),
    'L131': (f273, 131, 12),
    'E132': (f312, 132, 5),
    'F132': (f312, 132, 6),
    'J132': (f191, 132, 10),
    'K132': (f313, 132, 11),
    'L132': (f273, 132, 12),
    'E133': (f314, 133, 5),
    'F133': (f314, 133, 6),
    'J133': (f191, 133, 10),
    'K133': (f315, 133, 11),
    'L133': (f273, 133, 12),
    'E134': (f316, 134, 5),
    'F134': (f316, 134, 6),
    'J134': (f191, 134, 10),
    'K134': (f317, 134, 11),
    'L134': (f273, 134, 12),
    'E135': (f318, 135, 5),
    'F135': (f318, 135, 6),
    'J135': (f191, 135, 10),
    'K135': (f319, 135, 11),
    'L135': (f273, 135, 12),
    'E136': (f320, 136, 5),
    'F136': (f320, 136, 6),
    'J136': (f191, 136, 10),
    'K136': (f321, 136, 11),
    'L136': (f273, 136, 12),
    'E137': (f322, 137, 5),
    'F137': (f322, 137, 6),
    'J137': (f191, 137, 10),
    'K137': (f323, 137, 11),
    'L137': (f273, 137, 12),
    'E138': (f324, 138, 5),
    'F138': (f325, 138, 6),
    'J138': (f191, 138, 10),
    'K138': (f326, 138, 11),
    'L138': (f273, 138, 12),
    'E139': (f327, 139, 5),
    'F139': (f327, 139, 6),
    'J139': (f191, 139, 10),
    'K139': (f328, 139, 11),
    'L139': (f273, 139, 12),
    'E140': (f329, 140, 5),
    'F140': (f329, 140, 6),
    'J140': (f191, 140, 10),
    'K140': (f330, 140, 11),
    'L140': (f273, 140, 12),
    'E141': (f331, 141, 5),
    'F141': (f331, 141, 6),
    'J141': (f191, 141, 10),
    'K141': (f332, 141, 11),
    'L141': (f273, 141, 12),
    'E142': (f333, 142, 5),
    'F142': (f333, 142, 6),
    'J142': (f191, 142, 10),
    'K142': (f334, 142, 11),
    'L142': (f273, 142, 12),
    'E143': (f335, 143, 5),
    'F143': (f335, 143, 6),
    'J143': (f191, 143, 10),
    'K143': (f336, 143, 11),
    'L143': (f273, 143, 12),
    'E144': (f337, 144, 5),
    'F144': (f337, 144, 6),
    'J144': (f191, 144, 10),
    'K144': (f338, 144, 11),
    'L144': (f273, 144, 12),
    'E145': (f339, 145, 5),
    'F145': (f339, 145, 6),
    'J145': (f191, 145, 10),
    'K145': (f340, 145, 11),
    'L145': (f273, 145, 12),
    'E146': (f341, 146, 5),
    'F146': (f341, 146, 6),
    'J146': (f191, 146, 10),
    'K146': (f342, 146, 11),
    'L146': (f273, 146, 12),
    'E147': (f343, 147, 5),
    'F147': (f343, 147, 6),
    'J147': (f191, 147, 10),
    'K147': (f344, 147, 11),
    'L147': (f273, 147, 12),
    'E148': (f345, 148, 5),
    'F148': (f345, 148, 6),
    'J148': (f191, 148, 10),
    'K148': (f346, 148, 11),
    'L148': (f273, 148, 12),
    'E149': (f347, 149, 5),
    'F149': (f347, 149, 6),
    'J149': (f191, 149, 10),
    'K149': (f348, 149, 11),
    'L149': (f273, 149, 12),
    'E150': (f349, 150, 5),
    'F150': (f349, 150, 6),
    'J150': (f191, 150, 10),
    'K150': (f350, 150, 11),
    'L150': (f273, 150, 12),
    'E151': (f351, 151, 5),
    'F151': (f351, 151, 6),
    'J151': (f191, 151, 10),
    'K151': (f352, 151, 11),
    'L151': (f273, 151, 12),
    'E152': (f353, 152, 5),
    'F152': (f353, 152, 6),
    'J152': (f191, 152, 10),
    'K152': (f354, 152, 11),
    'L152': (f273, 152, 12),
    'E153': (f355, 153, 5),
    'F153': (f355, 153, 6),
    'J153': (f191, 153, 10),
    'K153': (f356, 153, 11),
    'L153': (f273, 153, 12),
    'E154': (f357, 154, 5),
    'F154': (f357, 154, 6),
    'J154': (f191, 154, 10),
    'K154': (f358, 154, 11),
    'L154': (f273, 154, 12),
    'E155': (f359, 155, 5),
    'F155': (f359, 155, 6),
    'J155': (f191, 155, 10),
    'K155': (f360, 155, 11),
    'L155': (f273, 155, 12),
    'E156': (f361, 156, 5),
    'F156': (f361, 156, 6),
    'J156': (f191, 156, 10),
    'K156': (f362, 156, 11),
    'L156': (f273, 156, 12),
    'E157': (f363, 157, 5),
    'F157': (f363, 157, 6),
    'J157': (f191, 157, 10),
    'K157': (f364, 157, 11),
    'L157': (f273, 157, 12),
    'E158': (f365, 158, 5),
    'F158': (f365, 158, 6),
    'J158': (f191, 158, 10),
    'K158': (f366, 158, 11),
    'L158': (f273, 158, 12),
    'E159': (f367, 159, 5),
    'F159': (f367, 159, 6),
    'J159': (f191, 159, 10),
    'K159': (f368, 159, 11),
    'L159': (f273, 159, 12),
    'E160': (f369, 160, 5),
    'F160': (f369, 160, 6),
    'J160': (f191, 160, 10),
    'K160': (f370, 160, 11),
    'L160': (f273, 160, 12),
    'E161': (f371, 161, 5),
    'F161': (f371, 161, 6),
    'J161': (f191, 161, 10),
    'K161': (f372, 161, 11),
    'L161': (f273, 161, 12),
    'E162': (f373, 162, 5),
    'F162': (f374, 162, 6),
    'J162': (f191, 162, 10),
    'K162': (f375, 162, 11),
    'L162': (f273, 162, 12),
    'E163': (f376, 163, 5),
    'F163': (f376, 163, 6),
    'J163': (f191, 163, 10),
    'K163': (f377, 163, 11),
    'L163': (f273, 163, 12),
    'E164': (f378, 164, 5),
    'F164': (f378, 164, 6),
    'J164': (f191, 164, 10),
    'K164': (f379, 164, 11),
    'L164': (f273, 164, 12),
    'E165': (f380, 165, 5),
    'F165': (f380, 165, 6),
    'J165': (f191, 165, 10),
    'K165': (f381, 165, 11),
    'L165': (f273, 165, 12),
    'E166': (f382, 166, 5),
    'F166': (f382, 166, 6),
    'J166': (f191, 166, 10),
    'K166': (f383, 166, 11),
    'L166': (f273, 166, 12),
    'E167': (f206, 167, 5),
    'F167': (f206, 167, 6),
    'J167': (f191, 167, 10),
    'K167': (f384, 167, 11),
    'L167': (f273, 167, 12),
    'E168': (f385, 168, 5),
    'F168': (f385, 168, 6),
    'J168': (f191, 168, 10),
    'K168': (f386, 168, 11),
    'L168': (f273, 168, 12),
    'E169': (f387, 169, 5),
    'F169': (f387, 169, 6),
    'J169': (f191, 169, 10),
    'K169': (f388, 169, 11),
    'L169': (f273, 169, 12),
    'E170': (f389, 170, 5),
    'F170': (f389, 170, 6),
    'J170': (f191, 170, 10),
    'K170': (f390, 170, 11),
    'L170': (f273, 170, 12),
    'E171': (f391, 171, 5),
    'F171': (f391, 171, 6),
    'J171': (f191, 171, 10),
    'K171': (f392, 171, 11),
    'L171': (f273, 171, 12),
    'E172': (f393, 172, 5),
    'F172': (f393, 172, 6),
    'J172': (f191, 172, 10),
    'K172': (f394, 172, 11),
    'L172': (f273, 172, 12),
    'E173': (f395, 173, 5),
    'F173': (f395, 173, 6),
    'J173': (f191, 173, 10),
    'K173': (f396, 173, 11),
    'L173': (f273, 173, 12),
    'E174': (f397, 174, 5),
    'F174': (f397, 174, 6),
    'J174': (f191, 174, 10),
    'K174': (f398, 174, 11),
    'L174': (f273, 174, 12),
    'E175': (f399, 175, 5),
    'F175': (f399, 175, 6),
    'J175': (f191, 175, 10),
    'K175': (f400, 175, 11),
    'L175': (f273, 175, 12),
    'E176': (f401, 176, 5),
    'F176': (f401, 176, 6),
    'J176': (f191, 176, 10),
    'K176': (f402, 176, 11),
    'L176': (f273, 176, 12),
    'E177': (f403, 177, 5),
    'F177': (f403, 177, 6),
    'J177': (f191, 177, 10),
    'K177': (f404, 177, 11),
    'L177': (f273, 177, 12),
    'E178': (f405, 178, 5),
    'F178': (f405, 178, 6),
    'J178': (f191, 178, 10),
    'K178': (f406, 178, 11),
    'L178': (f273, 178, 12),
    'E179': (f407, 179, 5),
    'F179': (f407, 179, 6),
    'J179': (f191, 179, 10),
    'K179': (f408, 179, 11),
    'L179': (f273, 179, 12),
    'E180': (f409, 180, 5),
    'F180': (f409, 180, 6),
    'J180': (f191, 180, 10),
    'K180': (f410, 180, 11),
    'L180': (f273, 180, 12),
    'E181': (f411, 181, 5),
    'F181': (f411, 181, 6),
    'J181': (f191, 181, 10),
    'K181': (f412, 181, 11),
    'L181': (f273, 181, 12),
    'E182': (f413, 182, 5),
    'F182': (f413, 182, 6),
    'J182': (f191, 182, 10),
    'K182': (f414, 182, 11),
    'L182': (f273, 182, 12),
    'E183': (f415, 183, 5),
    'F183': (f415, 183, 6),
    'J183': (f191, 183, 10),
    'K183': (f416, 183, 11),
    'L183': (f273, 183, 12),
    'E184': (f417, 184, 5),
    'F184': (f417, 184, 6),
    'J184': (f191, 184, 10),
    'K184': (f418, 184, 11),
    'L184': (f273, 184, 12),
    'E185': (f419, 185, 5),
    'F185': (f419, 185, 6),
    'J185': (f191, 185, 10),
    'K185': (f420, 185, 11),
    'L185': (f273, 185, 12),
    'E186': (f421, 186, 5),
    'F186': (f421, 186, 6),
    'J186': (f191, 186, 10),
    'K186': (f422, 186, 11),
    'L186': (f273, 186, 12),
    'E187': (f423, 187, 5),
    'F187': (f423, 187, 6),
    'J187': (f191, 187, 10),
    'K187': (f424, 187, 11),
    'L187': (f273, 187, 12),
    'E188': (f425, 188, 5),
    'F188': (f425, 188, 6),
    'J188': (f191, 188, 10),
    'K188': (f426, 188, 11),
    'L188': (f273, 188, 12),
    'E189': (f427, 189, 5),
    'F189': (f427, 189, 6),
    'J189': (f191, 189, 10),
    'K189': (f428, 189, 11),
    'L189': (f273, 189, 12),
    'E190': (f429, 190, 5),
    'F190': (f429, 190, 6),
    'J190': (f191, 190, 10),
    'K190': (f430, 190, 11),
    'L190': (f273, 190, 12),
    'E191': (f431, 191, 5),
    'F191': (f431, 191, 6),
    'J191': (f191, 191, 10),
    'K191': (f432, 191, 11),
    'L191': (f273, 191, 12),
    'E192': (f433, 192, 5),
    'F192': (f433, 192, 6),
    'J192': (f191, 192, 10),
    'K192': (f434, 192, 11),
    'L192': (f273, 192, 12),
    'E193': (f435, 193, 5),
    'F193': (f435, 193, 6),
    'J193': (f191, 193, 10),
    'K193': (f436, 193, 11),
    'L193': (f273, 193, 12),
    'E194': (f437, 194, 5),
    'F194': (f437, 194, 6),
    'J194': (f191, 194, 10),
    'K194': (f438, 194, 11),
    'L194': (f273, 194, 12),
    'E195': (f439, 195, 5),
    'F195': (f439, 195, 6),
    'J195': (f191, 195, 10),
    'K195': (f440, 195, 11),
    'L195': (f273, 195, 12),
    'E196': (f441, 196, 5),
    'F196': (f441, 196, 6),
    'J196': (f191, 196, 10),
    'K196': (f442, 196, 11),
    'L196': (f273, 196, 12),
    'E197': (f443, 197, 5),
    'F197': (f443, 197, 6),
    'J197': (f191, 197, 10),
    'K197': (f444, 197, 11),
    'L197': (f273, 197, 12),
    'E198': (f445, 198, 5),
    'F198': (f445, 198, 6),
    'J198': (f191, 198, 10),
    'K198': (f446, 198, 11),
    'L198': (f273, 198, 12),
    'E199': (f447, 199, 5),
    'F199': (f447, 199, 6),
    'J199': (f191, 199, 10),
    'K199': (f448, 199, 11),
    'L199': (f273, 199, 12),
    'E200': (f449, 200, 5),
    'F200': (f449, 200, 6),
    'J200': (f191, 200, 10),
    'K200': (f450, 200, 11),
    'L200': (f273, 200, 12),
    'E201': (f451, 201, 5),
    'F201': (f451, 201, 6),
    'J201': (f191, 201, 10),
    'K201': (f452, 201, 11),
    'L201': (f273, 201, 12),
    'E202': (f453, 202, 5),
    'F202': (f453, 202, 6),
    'J202': (f191, 202, 10),
    'K202': (f454, 202, 11),
    'L202': (f273, 202, 12),
    'E203': (f455, 203, 5),
    'F203': (f455, 203, 6),
    'J203': (f191, 203, 10),
    'K203': (f456, 203, 11),
    'L203': (f273, 203, 12),
    'E204': (f457, 204, 5),
    'F204': (f457, 204, 6),
    'J204': (f191, 204, 10),
    'K204': (f458, 204, 11),
    'L204': (f273, 204, 12),
    'E205': (f459, 205, 5),
    'F205': (f459, 205, 6),
    'J205': (f191, 205, 10),
    'K205': (f460, 205, 11),
    'L205': (f273, 205, 12),
    'E206': (f461, 206, 5),
    'F206': (f461, 206, 6),
    'J206': (f191, 206, 10),
    'K206': (f462, 206, 11),
    'L206': (f273, 206, 12),
    'E207': (f463, 207, 5),
    'F207': (f463, 207, 6),
    'J207': (f191, 207, 10),
    'K207': (f464, 207, 11),
    'L207': (f273, 207, 12),
    'E208': (f465, 208, 5),
    'F208': (f465, 208, 6),
    'J208': (f191, 208, 10),
    'K208': (f466, 208, 11),
    'L208': (f273, 208, 12),
    'E209': (f467, 209, 5),
    'F209': (f467, 209, 6),
    'J209': (f191, 209, 10),
    'K209': (f468, 209, 11),
    'L209': (f273, 209, 12),
    'E210': (f469, 210, 5),
    'F210': (f469, 210, 6),
    'J210': (f191, 210, 10),
    'K210': (f470, 210, 11),
    'L210': (f273, 210, 12),
    'E211': (f471, 211, 5),
    'F211': (f471, 211, 6),
    'J211': (f191, 211, 10),
    'K211': (f472, 211, 11),
    'L211': (f273, 211, 12),
    'E212': (f473, 212, 5),
    'F212': (f473, 212, 6),
    'J212': (f191, 212, 10),
    'K212': (f474, 212, 11),
    'L212': (f273, 212, 12),
    'E213': (f475, 213, 5),
    'F213': (f475, 213, 6),
    'J213': (f191, 213, 10),
    'K213': (f476, 213, 11),
    'L213': (f273, 213, 12),
    'E214': (f477, 214, 5),
    'F214': (f477, 214, 6),
    'J214': (f191, 214, 10),
    'K214': (f478, 214, 11),
    'L214': (f273, 214, 12),
    'E215': (f479, 215, 5),
    'F215': (f479, 215, 6),
    'J215': (f191, 215, 10),
    'K215': (f480, 215, 11),
    'L215': (f273, 215, 12),
    'E216': (f481, 216, 5),
    'F216': (f481, 216, 6),
    'J216': (f191, 216, 10),
    'K216': (f482, 216, 11),
    'L216': (f273, 216, 12),
    'E217': (f483, 217, 5),
    'F217': (f483, 217, 6),
    'J217': (f191, 217, 10),
    'K217': (f484, 217, 11),
    'L217': (f273, 217, 12),
    'E218': (f485, 218, 5),
    'F218': (f485, 218, 6),
    'J218': (f191, 218, 10),
    'K218': (f486, 218, 11),
    'L218': (f273, 218, 12),
    'E219': (f487, 219, 5),
    'F219': (f487, 219, 6),
    'J219': (f191, 219, 10),
    'K219': (f488, 219, 11),
    'L219': (f273, 219, 12),
    'E220': (f489, 220, 5),
    'F220': (f489, 220, 6),
    'J220': (f191, 220, 10),
    'K220': (f490, 220, 11),
    'L220': (f273, 220, 12),
    'E221': (f491, 221, 5),
    'F221': (f491, 221, 6),
    'J221': (f191, 221, 10),
    'K221': (f492, 221, 11),
    'L221': (f273, 221, 12),
    'E222': (f493, 222, 5),
    'F222': (f493, 222, 6),
    'J222': (f191, 222, 10),
    'K222': (f494, 222, 11),
    'L222': (f273, 222, 12),
    'E223': (f495, 223, 5),
    'F223': (f495, 223, 6),
    'J223': (f191, 223, 10),
    'K223': (f496, 223, 11),
    'L223': (f273, 223, 12),
    'E224': (f497, 224, 5),
    'F224': (f497, 224, 6),
    'J224': (f191, 224, 10),
    'K224': (f498, 224, 11),
    'L224': (f273, 224, 12),
    'E225': (f499, 225, 5),
    'F225': (f499, 225, 6),
    'J225': (f191, 225, 10),
    'K225': (f500, 225, 11),
    'L225': (f273, 225, 12),
    'E226': (f501, 226, 5),
    'F226': (f501, 226, 6),
    'J226': (f191, 226, 10),
    'K226': (f502, 226, 11),
    'L226': (f273, 226, 12),
    'E227': (f503, 227, 5),
    'F227': (f503, 227, 6),
    'J227': (f191, 227, 10),
    'K227': (f504, 227, 11),
    'L227': (f273, 227, 12),
    'E228': (f505, 228, 5),
    'F228': (f505, 228, 6),
    'J228': (f191, 228, 10),
    'K228': (f506, 228, 11),
    'L228': (f273, 228, 12),
    'E229': (f507, 229, 5),
    'F229': (f507, 229, 6),
    'J229': (f191, 229, 10),
    'K229': (f508, 229, 11),
    'L229': (f273, 229, 12),
    'E230': (f509, 230, 5),
    'F230': (f509, 230, 6),
    'J230': (f191, 230, 10),
    'K230': (f510, 230, 11),
    'L230': (f273, 230, 12),
    'E231': (f511, 231, 5),
    'F231': (f511, 231, 6),
    'J231': (f191, 231, 10),
    'K231': (f512, 231, 11),
    'L231': (f273, 231, 12),
    'E232': (f513, 232, 5),
    'F232': (f513, 232, 6),
    'J232': (f191, 232, 10),
    'K232': (f514, 232, 11),
    'L232': (f273, 232, 12),
    'E233': (f515, 233, 5),
    'F233': (f515, 233, 6),
    'J233': (f191, 233, 10),
    'K233': (f516, 233, 11),
    'L233': (f273, 233, 12),
    'E234': (f517, 234, 5),
    'F234': (f517, 234, 6),
    'J234': (f191, 234, 10),
    'K234': (f518, 234, 11),
    'L234': (f273, 234, 12),
    'E235': (f519, 235, 5),
    'F235': (f519, 235, 6),
    'J235': (f191, 235, 10),
    'K235': (f520, 235, 11),
    'L235': (f273, 235, 12),
    'E236': (f521, 236, 5),
    'F236': (f521, 236, 6),
    'J236': (f191, 236, 10),
    'K236': (f522, 236, 11),
    'L236': (f273, 236, 12),
    'E237': (f523, 237, 5),
    'F237': (f523, 237, 6),
    'J237': (f191, 237, 10),
    'K237': (f524, 237, 11),
    'L237': (f273, 237, 12),
    'E238': (f525, 238, 5),
    'F238': (f525, 238, 6),
    'J238': (f191, 238, 10),
    'K238': (f526, 238, 11),
    'L238': (f273, 238, 12),
    'E239': (f527, 239, 5),
    'F239': (f527, 239, 6),
    'J239': (f191, 239, 10),
    'K239': (f528, 239, 11),
    'L239': (f273, 239, 12),
    'E240': (f529, 240, 5),
    'F240': (f529, 240, 6),
    'J240': (f191, 240, 10),
    'K240': (f530, 240, 11),
    'L240': (f273, 240, 12),
    'E241': (f531, 241, 5),
    'F241': (f532, 241, 6),
    'J241': (f191, 241, 10),
    'K241': (f533, 241, 11),
    'L241': (f273, 241, 12),
    'E242': (f534, 242, 5),
    'F242': (f534, 242, 6),
    'J242': (f191, 242, 10),
    'K242': (f535, 242, 11),
    'L242': (f273, 242, 12),
    'E243': (f536, 243, 5),
    'F243': (f536, 243, 6),
    'J243': (f191, 243, 10),
    'K243': (f537, 243, 11),
    'L243': (f273, 243, 12),
    'E244': (f538, 244, 5),
    'F244': (f538, 244, 6),
    'J244': (f191, 244, 10),
    'K244': (f539, 244, 11),
    'L244': (f273, 244, 12),
    'E245': (f540, 245, 5),
    'F245': (f540, 245, 6),
    'J245': (f191, 245, 10),
    'K245': (f541, 245, 11),
    'L245': (f273, 245, 12),
    'E246': (f542, 246, 5),
    'F246': (f542, 246, 6),
    'J246': (f191, 246, 10),
    'K246': (f543, 246, 11),
    'L246': (f273, 246, 12),
    'E247': (f544, 247, 5),
    'F247': (f544, 247, 6),
    'J247': (f191, 247, 10),
    'K247': (f545, 247, 11),
    'L247': (f273, 247, 12),
    'E248': (f546, 248, 5),
    'F248': (f546, 248, 6),
    'J248': (f191, 248, 10),
    'K248': (f547, 248, 11),
    'L248': (f273, 248, 12),
    'E249': (f548, 249, 5),
    'F249': (f548, 249, 6),
    'J249': (f191, 249, 10),
    'K249': (f549, 249, 11),
    'L249': (f273, 249, 12),
    'E250': (f550, 250, 5),
    'F250': (f550, 250, 6),
    'J250': (f191, 250, 10),
    'K250': (f551, 250, 11),
    'L250': (f273, 250, 12),
    'E251': (f552, 251, 5),
    'F251': (f552, 251, 6),
    'J251': (f191, 251, 10),
    'K251': (f553, 251, 11),
    'L251': (f273, 251, 12),
    'E252': (f554, 252, 5),
    'F252': (f554, 252, 6),
    'J252': (f191, 252, 10),
    'K252': (f555, 252, 11),
    'L252': (f273, 252, 12),
    'E253': (f556, 253, 5),
    'F253': (f556, 253, 6),
    'J253': (f191, 253, 10),
    'K253': (f557, 253, 11),
    'L253': (f273, 253, 12),
    'E254': (f558, 254, 5),
    'F254': (f558, 254, 6),
    'J254': (f191, 254, 10),
    'K254': (f559, 254, 11),
    'L254': (f273, 254, 12),
    'E255': (f560, 255, 5),
    'F255': (f560, 255, 6),
    'J255': (f191, 255, 10),
    'K255': (f561, 255, 11),
    'L255': (f273, 255, 12),
    'E256': (f562, 256, 5),
    'F256': (f562, 256, 6),
    'J256': (f191, 256, 10),
    'K256': (f563, 256, 11),
    'L256': (f273, 256, 12),
    'E257': (f564, 257, 5),
    'F257': (f564, 257, 6),
    'J257': (f191, 257, 10),
    'K257': (f565, 257, 11),
    'L257': (f273, 257, 12),
    'E258': (f566, 258, 5),
    'F258': (f566, 258, 6),
    'J258': (f191, 258, 10),
    'K258': (f567, 258, 11),
    'L258': (f273, 258, 12),
    'E259': (f568, 259, 5),
    'F259': (f568, 259, 6),
    'J259': (f191, 259, 10),
    'K259': (f569, 259, 11),
    'L259': (f273, 259, 12),
    'E260': (f570, 260, 5),
    'F260': (f570, 260, 6),
    'J260': (f191, 260, 10),
    'K260': (f571, 260, 11),
    'L260': (f273, 260, 12),
    'E261': (f572, 261, 5),
    'F261': (f572, 261, 6),
    'J261': (f191, 261, 10),
    'K261': (f573, 261, 11),
    'L261': (f273, 261, 12),
    'E262': (f574, 262, 5),
    'F262': (f574, 262, 6),
    'J262': (f191, 262, 10),
    'K262': (f575, 262, 11),
    'L262': (f273, 262, 12),
    'E263': (f576, 263, 5),
    'F263': (f576, 263, 6),
    'J263': (f191, 263, 10),
    'K263': (f577, 263, 11),
    'L263': (f273, 263, 12),
    'E264': (f578, 264, 5),
    'F264': (f578, 264, 6),
    'J264': (f191, 264, 10),
    'K264': (f579, 264, 11),
    'L264': (f273, 264, 12),
    'E265': (f580, 265, 5),
    'F265': (f580, 265, 6),
    'J265': (f191, 265, 10),
    'K265': (f581, 265, 11),
    'L265': (f273, 265, 12),
    'E266': (f582, 266, 5),
    'F266': (f583, 266, 6),
    'J266': (f191, 266, 10),
    'K266': (f584, 266, 11),
    'L266': (f273, 266, 12),
    'E267': (f585, 267, 5),
    'F267': (f585, 267, 6),
    'J267': (f191, 267, 10),
    'K267': (f586, 267, 11),
    'L267': (f273, 267, 12),
    'E268': (f587, 268, 5),
    'F268': (f587, 268, 6),
    'J268': (f191, 268, 10),
    'K268': (f588, 268, 11),
    'L268': (f273, 268, 12),
    'E269': (f589, 269, 5),
    'F269': (f590, 269, 6),
    'J269': (f191, 269, 10),
    'K269': (f591, 269, 11),
    'L269': (f273, 269, 12),
    'E270': (f592, 270, 5),
    'F270': (f592, 270, 6),
    'J270': (f191, 270, 10),
    'K270': (f593, 270, 11),
    'L270': (f273, 270, 12),
    'E271': (f594, 271, 5),
    'F271': (f594, 271, 6),
    'J271': (f191, 271, 10),
    'K271': (f595, 271, 11),
    'L271': (f273, 271, 12),
    'E272': (f596, 272, 5),
    'F272': (f596, 272, 6),
    'J272': (f191, 272, 10),
    'K272': (f597, 272, 11),
    'L272': (f273, 272, 12),
    'E273': (f598, 273, 5),
    'F273': (f598, 273, 6),
    'J273': (f191, 273, 10),
    'K273': (f599, 273, 11),
    'L273': (f273, 273, 12),
    'E274': (f600, 274, 5),
    'F274': (f600, 274, 6),
    'J274': (f191, 274, 10),
    'K274': (f601, 274, 11),
    'L274': (f273, 274, 12),
    'E275': (f602, 275, 5),
    'F275': (f602, 275, 6),
    'J275': (f191, 275, 10),
    'K275': (f603, 275, 11),
    'L275': (f273, 275, 12),
    'E280': (f604, 280, 5),
    'F280': (f604, 280, 6),
    'J280': (f191, 280, 10),
    'K280': (f605, 280, 11),
    'L280': (f273, 280, 12),
    'E281': (f606, 281, 5),
    'F281': (f607, 281, 6),
    'J281': (f191, 281, 10),
    'K281': (f608, 281, 11),
    'L281': (f273, 281, 12),
    'E282': (f609, 282, 5),
    'F282': (f610, 282, 6),
    'J282': (f191, 282, 10),
    'K282': (f611, 282, 11),
    'L282': (f273, 282, 12),
}
