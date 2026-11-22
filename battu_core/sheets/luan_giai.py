# -*- coding: utf-8 -*-
"""Module sheet LUAN_GIAI – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.9.xlsx (SHA-256 5486ea8b3eb631dc513aed16150152722659ca3ca34dfc6c7af268da54b63453).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'LUAN_GIAI'
SI = 4

def f0(E, r, c):
    # Ô S5
    # = TRUNG_GIAN!$E$5*12+TRUNG_GIAN!$D$6
    return X.add(X.mul(E.v(23,5,5),12.0),E.v(23,6,4))

def f1(E, r, c):
    # Ô S6
    # = INDEX($N$5:$N$124,$S$5+1)
    return X.INDEX(E.rg(4,5,14,124,14), X.add(E.v(4,5,19),1.0))

def f2(E, r, c):
    # Ô T6, T7, T8, T9, T10
    # = IF(S6<0,"",INDEX(CAN_TEN,S6+1)&" ("&INDEX(HANH_TEN,INT(S6/2)+1)&" – "&IF((COUNTIF(TRUNG_GIAN!$C$5:$D$5,S6)+COUNTIF(TRUNG_GIAN!$F$5,S6)>0),"có, lộ ở th
    #   iên can",IF((COUNTIF(TRUNG_GIAN!$C$15:$F$15,S6)+COUNTIF(TRUNG_GIAN!$C$17:$F$17,S6)+COUNTIF(TRUNG_GIAN!$C$19:$F$19,S6)>0),"có, tàng trong chi","không c
    #   ó trong tứ trụ"))&")")
    return ('' if X.truthy(X.lt(E.v(4,r,c-1),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(4,r,c-1),1.0)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.INT(X.div(E.v(4,r,c-1),2.0)),1.0))),' – '),('có, lộ ở thiên can' if X.truthy(X.gt(X.add(X.COUNTIF(E.rg(23,5,3,5,4), E.v(4,r,c-1)),X.COUNTIF(E.rg(23,5,6,5,6), E.v(4,r,c-1))),0.0)) else ('có, tàng trong chi' if X.truthy(X.gt(X.add(X.add(X.COUNTIF(E.rg(23,15,3,15,6), E.v(4,r,c-1)),X.COUNTIF(E.rg(23,17,3,17,6), E.v(4,r,c-1))),X.COUNTIF(E.rg(23,19,3,19,6), E.v(4,r,c-1))),0.0)) else 'không có trong tứ trụ'))),')'))

def f3(E, r, c):
    # Ô S7
    # = INDEX($O$5:$O$124,$S$5+1)
    return X.INDEX(E.rg(4,5,15,124,15), X.add(E.v(4,5,19),1.0))

def f4(E, r, c):
    # Ô C8
    # = "• Vượng suy: "&DOI_CHIEU!$C$34&" – độ tin "&LOWER(DOI_CHIEU!$C$35)&"."&CHAR(10)&"• Cách cục: "&IFERROR(LEFT(IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$35,KET_L
    #   UAN!$C$35),FIND(CHAR(10),IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$35,KET_LUAN!$C$35))-1),IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$35,KET_LUAN!$C$35))&"."&CHAR(10)&"• 
    #   Dụng thần: "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&" · hỷ thần: "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$C$37)&" · kỵ thần: "&IF(
    #   DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)&"."&CHAR(10)&"• Nói đơn giản: "&LUAN_GIAI!$F$106&CHAR(10)&"• Mệnh cục (P1 tr.164-173): "&IFERROR(LEFT
    #   (MENH_CUC!$D$18,FIND(" (",MENH_CUC!$D$18)-1),MENH_CUC!$D$18)&"; "&IFERROR(LEFT(MENH_CUC!$D$19,FIND(" (",MENH_CUC!$D$19)-1),MENH_CUC!$D$19)&"; "&IFERRO
    #   R(LEFT(MENH_CUC!$D$20,FIND(" (",MENH_CUC!$D$20)-1),MENH_CUC!$D$20)&"; "&IFERROR(LEFT(MENH_CUC!$D$21,FIND(" (",MENH_CUC!$D$21)-1),MENH_CUC!$D$21)&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Vượng suy: ',E.v(7,34,3)),' – độ tin '),X.LOWER(E.v(7,35,3))),'.'),X.CHAR(10.0)),'• Cách cục: '),X.IFERROR(lambda: X.LEFT((E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3)), X.sub(X.FIND(X.CHAR(10.0), (E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3))),1.0)), lambda: (E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3)))),'.'),X.CHAR(10.0)),'• Dụng thần: '),(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),' · hỷ thần: '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),' · kỵ thần: '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3))),'.'),X.CHAR(10.0)),'• Nói đơn giản: '),E.v(4,106,6)),X.CHAR(10.0)),'• Mệnh cục (P1 tr.164-173): '),X.IFERROR(lambda: X.LEFT(E.v(18,18,4), X.sub(X.FIND(' (', E.v(18,18,4)),1.0)), lambda: E.v(18,18,4))),'; '),X.IFERROR(lambda: X.LEFT(E.v(18,19,4), X.sub(X.FIND(' (', E.v(18,19,4)),1.0)), lambda: E.v(18,19,4))),'; '),X.IFERROR(lambda: X.LEFT(E.v(18,20,4), X.sub(X.FIND(' (', E.v(18,20,4)),1.0)), lambda: E.v(18,20,4))),'; '),X.IFERROR(lambda: X.LEFT(E.v(18,21,4), X.sub(X.FIND(' (', E.v(18,21,4)),1.0)), lambda: E.v(18,21,4))),'.')

def f5(E, r, c):
    # Ô S8
    # = INDEX($P$5:$P$124,$S$5+1)
    return X.INDEX(E.rg(4,5,16,124,16), X.add(E.v(4,5,19),1.0))

def f6(E, r, c):
    # Ô C9, C18, C25, C36, C45, C57, C67, C84, C92
    # = IF(TRIM(J9)="","Mục này chưa có điểm cát tường nổi bật theo sách – xem chi tiết bên dưới.",TRIM(J9))
    return ('Mục này chưa có điểm cát tường nổi bật theo sách – xem chi tiết bên dưới.' if X.truthy(X.eq(X.TRIM(E.v(4,r,c+7)),'')) else X.TRIM(E.v(4,r,c+7)))

def f7(E, r, c):
    # Ô J9
    # = "• Hành mang may mắn: dụng thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&", hỷ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$C$37)
    #   &" – nên tăng cường trong nghề nghiệp, màu sắc, phương hướng, sinh hoạt. "&IF(LUC_THAN!$Y$31="","","• Giai đoạn thuận nhất (đại vận dụng/hỷ): "&LUC_TH
    #   AN!$Y$31&" ")&IF(OR(LEFT(MENH_CUC!$D$18,5)="THUẬN",LEFT(MENH_CUC!$D$18,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$18,3)="TỐT",MENH_CUC!$D$18="Khá",ISNUMBER(
    #   SEARCH("VỀ GIÀU",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ THỌ",MENH_C
    #   UC!$D$18))),"• Giàu – nghèo: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$18,FIND(" (",MENH_CUC!$D$18)-1),MENH_CUC!$D$18))&" ","")&IF(OR(LEFT(MENH_CUC!$D$19,5)="T
    #   HUẬN",LEFT(MENH_CUC!$D$19,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$19,3)="TỐT",MENH_CUC!$D$19="Khá",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$19)),ISNUMBER(SE
    #   ARCH("VỀ SANG",MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$19))),"• Sang – hèn: "&LOWER(IFERROR(LE
    #   FT(MENH_CUC!$D$19,FIND(" (",MENH_CUC!$D$19)-1),MENH_CUC!$D$19))&" ","")&IF(OR(LEFT(MENH_CUC!$D$20,5)="THUẬN",LEFT(MENH_CUC!$D$20,13)="NGHIÊNG THUẬN",L
    #   EFT(MENH_CUC!$D$20,3)="TỐT",MENH_CUC!$D$20="Khá",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$20)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$20)),ISNUMBER(SEARCH
    #   ("VỀ CÁT",MENH_CUC!$D$20)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$20))),"• Cát – hung: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$20,FIND(" (",MENH_CUC!$D$20)-1),
    #   MENH_CUC!$D$20))&" ","")&IF(OR(LEFT(MENH_CUC!$D$21,5)="THUẬN",LEFT(MENH_CUC!$D$21,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$21,3)="TỐT",MENH_CUC!$D$21="Khá
    #   ",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ 
    #   THỌ",MENH_CUC!$D$21))),"• Thọ – yểu: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$21,FIND(" (",MENH_CUC!$D$21)-1),MENH_CUC!$D$21))&" ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Hành mang may mắn: dụng thần ',(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),', hỷ thần '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),' – nên tăng cường trong nghề nghiệp, màu sắc, phương hướng, sinh hoạt. '),('' if X.truthy(X.eq(E.v(16,31,25),'')) else X.cat(X.cat('• Giai đoạn thuận nhất (đại vận dụng/hỷ): ',E.v(16,31,25)),' '))),(X.cat(X.cat('• Giàu – nghèo: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,18,4), X.sub(X.FIND(' (', E.v(18,18,4)),1.0)), lambda: E.v(18,18,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,18,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,18,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,18,4), 3.0),'TỐT'), X.eq(E.v(18,18,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,18,4))))) else '')),(X.cat(X.cat('• Sang – hèn: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,19,4), X.sub(X.FIND(' (', E.v(18,19,4)),1.0)), lambda: E.v(18,19,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,19,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,19,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,19,4), 3.0),'TỐT'), X.eq(E.v(18,19,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,19,4))))) else '')),(X.cat(X.cat('• Cát – hung: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,20,4), X.sub(X.FIND(' (', E.v(18,20,4)),1.0)), lambda: E.v(18,20,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,20,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,20,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,20,4), 3.0),'TỐT'), X.eq(E.v(18,20,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,20,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,20,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,20,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,20,4))))) else '')),(X.cat(X.cat('• Thọ – yểu: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,21,4), X.sub(X.FIND(' (', E.v(18,21,4)),1.0)), lambda: E.v(18,21,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,21,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,21,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,21,4), 3.0),'TỐT'), X.eq(E.v(18,21,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,21,4))))) else ''))

def f8(E, r, c):
    # Ô S9
    # = INDEX($Q$5:$Q$124,$S$5+1)
    return X.INDEX(E.rg(4,5,17,124,17), X.add(E.v(4,5,19),1.0))

def f9(E, r, c):
    # Ô C10
    # = KET_LUAN!$C$202
    return E.v(3,202,3)

def f10(E, r, c):
    # Ô S10
    # = INDEX($R$5:$R$124,$S$5+1)
    return X.INDEX(E.rg(4,5,18,124,18), X.add(E.v(4,5,19),1.0))

def f11(E, r, c):
    # Ô C11
    # = "Vì sao chọn dụng – hỷ – kỵ: "&IF(LEN(KET_LUAN!$C$39)>700,LEFT(KET_LUAN!$C$39,700)&"…",KET_LUAN!$C$39)
    return X.cat('Vì sao chọn dụng – hỷ – kỵ: ',(X.cat(X.LEFT(E.v(3,39,3), 700.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,39,3)),700.0)) else E.v(3,39,3)))

def f12(E, r, c):
    # Ô C12
    # = "Nói đơn giản: "&LUAN_GIAI!$F$106&" Độ tin vượng suy: "&LUAN_GIAI!$I$106
    return X.cat(X.cat(X.cat('Nói đơn giản: ',E.v(4,106,6)),' Độ tin vượng suy: '),E.v(4,106,9))

def f13(E, r, c):
    # Ô S12
    # = INT(S6/2)
    return X.INT(X.div(E.v(4,r-6,c),2.0))

def f14(E, r, c):
    # Ô C13
    # = "Trích Thiên Tủy về nhật can "&TRUNG_GIAN!$E$10&": "&INDEX($L$5:$L$14,TRUNG_GIAN!$E$5+1)&" – "&INDEX($M$5:$M$14,TRUNG_GIAN!$E$5+1)
    return X.cat(X.cat(X.cat(X.cat(X.cat('Trích Thiên Tủy về nhật can ',E.v(23,10,5)),': '),X.INDEX(E.rg(4,5,12,14,12), X.add(E.v(23,5,5),1.0))),' – '),X.INDEX(E.rg(4,5,13,14,13), X.add(E.v(23,5,5),1.0)))

def f15(E, r, c):
    # Ô C14
    # = "Cùng Thông Bảo Giám (bảng điều hầu): "&TRUNG_GIAN!$E$10&" sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" cần trước hết "&$T$6&IF($T$7="","",", sau đ
    #   ó "&$T$7)&IF($T$8="","","; "&$T$8)&IF($T$9="","","; "&$T$9)&IF($T$10="","","; "&$T$10)&IF(INDEX($U$5:$U$124,$S$5+1)=1," (ô này đã đối chiếu nguyên văn
    #   )"," (bản tóm lược phổ biến)")&". So với sách (phù ức): hành điều hầu thứ nhất "&INDEX(HANH_TEN,$S$12+1)&" "&IF($S$12=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!
    #   $D$109,DUNG_HY_KY!$C$109),"trùng dụng thần → hai trường phái cùng hướng, dụng thần càng đáng tin.",IF($S$12=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUN
    #   G_HY_KY!$C$110),"trùng hỷ thần → khá đồng thuận.",IF($S$12=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"lại là kỵ thần theo phù ức → hai
    #    nhu cầu ngược nhau: sách lấy phù ức làm chủ (P1 tr.113), nhưng vận/năm mang hành này vẫn có mặt tốt về điều hòa khí hậu (sức khỏe, tinh thần).","là n
    #   hàn thần theo sách → bổ trợ, không mâu thuẫn.")))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Cùng Thông Bảo Giám (bảng điều hầu): ',E.v(23,10,5)),' sinh tháng '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' cần trước hết '),E.v(4,6,20)),('' if X.truthy(X.eq(E.v(4,7,20),'')) else X.cat(', sau đó ',E.v(4,7,20)))),('' if X.truthy(X.eq(E.v(4,8,20),'')) else X.cat('; ',E.v(4,8,20)))),('' if X.truthy(X.eq(E.v(4,9,20),'')) else X.cat('; ',E.v(4,9,20)))),('' if X.truthy(X.eq(E.v(4,10,20),'')) else X.cat('; ',E.v(4,10,20)))),(' (ô này đã đối chiếu nguyên văn)' if X.truthy(X.eq(X.INDEX(E.rg(4,5,21,124,21), X.add(E.v(4,5,19),1.0)),1.0)) else ' (bản tóm lược phổ biến)')),'. So với sách (phù ức): hành điều hầu thứ nhất '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(4,12,19),1.0))),' '),('trùng dụng thần → hai trường phái cùng hướng, dụng thần càng đáng tin.' if X.truthy(X.eq(E.v(4,12,19),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('trùng hỷ thần → khá đồng thuận.' if X.truthy(X.eq(E.v(4,12,19),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('lại là kỵ thần theo phù ức → hai nhu cầu ngược nhau: sách lấy phù ức làm chủ (P1 tr.113), nhưng vận/năm mang hành này vẫn có mặt tốt về điều hòa khí hậu (sức khỏe, tinh thần).' if X.truthy(X.eq(E.v(4,12,19),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'là nhàn thần theo sách → bổ trợ, không mâu thuẫn.'))))

def f16(E, r, c):
    # Ô C15
    # = "Nhật chủ ngồi trên chi ngày ở cung "&CHOOSE(TRUNG_GIAN!$E$28+1,"Trường sinh – khởi đầu thuận, được nâng đỡ, lạc quan","Mộc dục – nhạy cảm, thích cái 
    #   đẹp, tình cảm nhiều biến động","Quan đới – tự trọng, coi trọng thể diện, chỉn chu","Lâm quan – tự lực, chủ động, làm được việc lớn","Đế vượng – bản lĩ
    #   nh, cái tôi lớn, khó chịu bị chỉ đạo","Suy – điềm đạm, thận trọng, ít mạo hiểm","Bệnh – hay lo nghĩ, cần giữ sức khỏe","Tử – trầm tĩnh, ít bộc lộ, nên
    #    dựa vào người nâng đỡ và môi trường tốt","Mộ – kín đáo, giỏi tích lũy, tiết kiệm","Tuyệt – đời nhiều thay đổi, hay phải làm lại từ đầu, thích nghi gi
    #   ỏi","Thai – nhiều ý tưởng, đang định hình, cần kiên trì","Dưỡng – được chăm lo, ôn hòa, phát triển chậm mà chắc")&". (Đọc 12 cung trường sinh của nhật
    #    chủ tại chi ngày theo kinh nghiệm thực hành; trạng thái tính ở TRUNG_GIAN dòng 28.)"
    return X.cat(X.cat('Nhật chủ ngồi trên chi ngày ở cung ',X.CHOOSE(X.add(E.v(23,28,5),1.0), lambda: 'Trường sinh – khởi đầu thuận, được nâng đỡ, lạc quan', lambda: 'Mộc dục – nhạy cảm, thích cái đẹp, tình cảm nhiều biến động', lambda: 'Quan đới – tự trọng, coi trọng thể diện, chỉn chu', lambda: 'Lâm quan – tự lực, chủ động, làm được việc lớn', lambda: 'Đế vượng – bản lĩnh, cái tôi lớn, khó chịu bị chỉ đạo', lambda: 'Suy – điềm đạm, thận trọng, ít mạo hiểm', lambda: 'Bệnh – hay lo nghĩ, cần giữ sức khỏe', lambda: 'Tử – trầm tĩnh, ít bộc lộ, nên dựa vào người nâng đỡ và môi trường tốt', lambda: 'Mộ – kín đáo, giỏi tích lũy, tiết kiệm', lambda: 'Tuyệt – đời nhiều thay đổi, hay phải làm lại từ đầu, thích nghi giỏi', lambda: 'Thai – nhiều ý tưởng, đang định hình, cần kiên trì', lambda: 'Dưỡng – được chăm lo, ôn hòa, phát triển chậm mà chắc')),'. (Đọc 12 cung trường sinh của nhật chủ tại chi ngày theo kinh nghiệm thực hành; trạng thái tính ở TRUNG_GIAN dòng 28.)')

def f17(E, r, c):
    # Ô C17
    # = "• Tố chất cốt lõi: "&TO_CHAT!$C$8&CHAR(10)&"• Tính cách theo sách: "&IF(LEN(KET_LUAN!$C$73)>380,LEFT(KET_LUAN!$C$73,380)&"…",KET_LUAN!$C$73)&CHAR(10)
    #   &"• Năng khiếu, điểm mạnh: "&IF(LEN(IFERROR(SUBSTITUTE(MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET_LUAN!$C$71)+11,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71
    #   ),LEN(KET_LUAN!$C$71)+1)-FIND("Điểm mạnh:",KET_LUAN!$C$71)-11),CHAR(10)," "),""))>260,LEFT(IFERROR(SUBSTITUTE(MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET
    #   _LUAN!$C$71)+11,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71),LEN(KET_LUAN!$C$71)+1)-FIND("Điểm mạnh:",KET_LUAN!$C$71)-11),CHAR(10)," "),""),260)&"…",IFERR
    #   OR(SUBSTITUTE(MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET_LUAN!$C$71)+11,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71),LEN(KET_LUAN!$C$71)+1)-FIND("Điểm mạnh:
    #   ",KET_LUAN!$C$71)-11),CHAR(10)," "),""))&CHAR(10)&"• Cần lưu ý: "&IF(LEN(IFERROR(MID(KET_LUAN!$C$71,FIND("Điểm yếu:",KET_LUAN!$C$71)+10,400),""))>200,
    #   LEFT(IFERROR(MID(KET_LUAN!$C$71,FIND("Điểm yếu:",KET_LUAN!$C$71)+10,400),""),200)&"…",IFERROR(MID(KET_LUAN!$C$71,FIND("Điểm yếu:",KET_LUAN!$C$71)+10,4
    #   00),""))&" "&IF(LEN(IFERROR(MID(KET_LUAN!$C$74,FIND("Điểm cần lưu ý:",KET_LUAN!$C$74)+16,400),""))>220,LEFT(IFERROR(MID(KET_LUAN!$C$74,FIND("Điểm cần 
    #   lưu ý:",KET_LUAN!$C$74)+16,400),""),220)&"…",IFERROR(MID(KET_LUAN!$C$74,FIND("Điểm cần lưu ý:",KET_LUAN!$C$74)+16,400),""))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Tố chất cốt lõi: ',E.v(20,8,3)),X.CHAR(10.0)),'• Tính cách theo sách: '),(X.cat(X.LEFT(E.v(3,73,3), 380.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,73,3)),380.0)) else E.v(3,73,3))),X.CHAR(10.0)),'• Năng khiếu, điểm mạnh: '),(X.cat(X.LEFT(X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: ''), 260.0),'…') if X.truthy(X.gt(X.LEN(X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: '')),260.0)) else X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: ''))),X.CHAR(10.0)),'• Cần lưu ý: '),(X.cat(X.LEFT(X.IFERROR(lambda: X.MID(E.v(3,71,3), X.add(X.FIND('Điểm yếu:', E.v(3,71,3)),10.0), 400.0), lambda: ''), 200.0),'…') if X.truthy(X.gt(X.LEN(X.IFERROR(lambda: X.MID(E.v(3,71,3), X.add(X.FIND('Điểm yếu:', E.v(3,71,3)),10.0), 400.0), lambda: '')),200.0)) else X.IFERROR(lambda: X.MID(E.v(3,71,3), X.add(X.FIND('Điểm yếu:', E.v(3,71,3)),10.0), 400.0), lambda: ''))),' '),(X.cat(X.LEFT(X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm cần lưu ý:', E.v(3,74,3)),16.0), 400.0), lambda: ''), 220.0),'…') if X.truthy(X.gt(X.LEN(X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm cần lưu ý:', E.v(3,74,3)),16.0), 400.0), lambda: '')),220.0)) else X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm cần lưu ý:', E.v(3,74,3)),16.0), 400.0), lambda: '')))

def f18(E, r, c):
    # Ô J18
    # = IF(IFERROR(SUBSTITUTE(MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET_LUAN!$C$71)+11,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71),LEN(KET_LUAN!$C$71)+1)-FIND("Đi
    #   ểm mạnh:",KET_LUAN!$C$71)-11),CHAR(10)," "),"")="","","• Điểm mạnh: "&IF(LEN(IFERROR(SUBSTITUTE(MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET_LUAN!$C$71)+1
    #   1,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71),LEN(KET_LUAN!$C$71)+1)-FIND("Điểm mạnh:",KET_LUAN!$C$71)-11),CHAR(10)," "),""))>300,LEFT(IFERROR(SUBSTITUTE
    #   (MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET_LUAN!$C$71)+11,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71),LEN(KET_LUAN!$C$71)+1)-FIND("Điểm mạnh:",KET_LUAN!$C
    #   $71)-11),CHAR(10)," "),""),300)&"…",IFERROR(SUBSTITUTE(MID(KET_LUAN!$C$71,FIND("Điểm mạnh:",KET_LUAN!$C$71)+11,IFERROR(FIND("Điểm yếu:",KET_LUAN!$C$71
    #   ),LEN(KET_LUAN!$C$71)+1)-FIND("Điểm mạnh:",KET_LUAN!$C$71)-11),CHAR(10)," "),""))&" ")&IF(IFERROR(MID(KET_LUAN!$C$74,FIND("Điểm tốt:",KET_LUAN!$C$74)+
    #   10,IFERROR(FIND("| Điểm cần lưu ý:",KET_LUAN!$C$74),LEN(KET_LUAN!$C$74)+1)-FIND("Điểm tốt:",KET_LUAN!$C$74)-10),"")="","","• Đức tính tốt: "&IF(LEN(IF
    #   ERROR(MID(KET_LUAN!$C$74,FIND("Điểm tốt:",KET_LUAN!$C$74)+10,IFERROR(FIND("| Điểm cần lưu ý:",KET_LUAN!$C$74),LEN(KET_LUAN!$C$74)+1)-FIND("Điểm tốt:",
    #   KET_LUAN!$C$74)-10),""))>420,LEFT(IFERROR(MID(KET_LUAN!$C$74,FIND("Điểm tốt:",KET_LUAN!$C$74)+10,IFERROR(FIND("| Điểm cần lưu ý:",KET_LUAN!$C$74),LEN(
    #   KET_LUAN!$C$74)+1)-FIND("Điểm tốt:",KET_LUAN!$C$74)-10),""),420)&"…",IFERROR(MID(KET_LUAN!$C$74,FIND("Điểm tốt:",KET_LUAN!$C$74)+10,IFERROR(FIND("| Đi
    #   ểm cần lưu ý:",KET_LUAN!$C$74),LEN(KET_LUAN!$C$74)+1)-FIND("Điểm tốt:",KET_LUAN!$C$74)-10),""))&" ")
    return X.cat(('' if X.truthy(X.eq(X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: ''),'')) else X.cat(X.cat('• Điểm mạnh: ',(X.cat(X.LEFT(X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: ''), 300.0),'…') if X.truthy(X.gt(X.LEN(X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: '')),300.0)) else X.IFERROR(lambda: X.SUBSTITUTE(X.MID(E.v(3,71,3), X.add(X.FIND('Điểm mạnh:', E.v(3,71,3)),11.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('Điểm yếu:', E.v(3,71,3)), lambda: X.add(X.LEN(E.v(3,71,3)),1.0)),X.FIND('Điểm mạnh:', E.v(3,71,3))),11.0)), X.CHAR(10.0), ' '), lambda: ''))),' ')),('' if X.truthy(X.eq(X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm tốt:', E.v(3,74,3)),10.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('| Điểm cần lưu ý:', E.v(3,74,3)), lambda: X.add(X.LEN(E.v(3,74,3)),1.0)),X.FIND('Điểm tốt:', E.v(3,74,3))),10.0)), lambda: ''),'')) else X.cat(X.cat('• Đức tính tốt: ',(X.cat(X.LEFT(X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm tốt:', E.v(3,74,3)),10.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('| Điểm cần lưu ý:', E.v(3,74,3)), lambda: X.add(X.LEN(E.v(3,74,3)),1.0)),X.FIND('Điểm tốt:', E.v(3,74,3))),10.0)), lambda: ''), 420.0),'…') if X.truthy(X.gt(X.LEN(X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm tốt:', E.v(3,74,3)),10.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('| Điểm cần lưu ý:', E.v(3,74,3)), lambda: X.add(X.LEN(E.v(3,74,3)),1.0)),X.FIND('Điểm tốt:', E.v(3,74,3))),10.0)), lambda: '')),420.0)) else X.IFERROR(lambda: X.MID(E.v(3,74,3), X.add(X.FIND('Điểm tốt:', E.v(3,74,3)),10.0), X.sub(X.sub(X.IFERROR(lambda: X.FIND('| Điểm cần lưu ý:', E.v(3,74,3)), lambda: X.add(X.LEN(E.v(3,74,3)),1.0)),X.FIND('Điểm tốt:', E.v(3,74,3))),10.0)), lambda: ''))),' ')))

def f19(E, r, c):
    # Ô C19
    # = "Tố chất cốt lõi: "&KET_LUAN!$C$70&" Tính cách: "&IF(LEN(KET_LUAN!$C$73)>400,LEFT(KET_LUAN!$C$73,400)&"…",KET_LUAN!$C$73)&IF(KET_LUAN!$C$76="","",CHAR
    #   (10)&"Bổ sung (Uyên Hải Tử Bình): "&KET_LUAN!$C$76)
    return X.cat(X.cat(X.cat(X.cat('Tố chất cốt lõi: ',E.v(3,70,3)),' Tính cách: '),(X.cat(X.LEFT(E.v(3,73,3), 400.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,73,3)),400.0)) else E.v(3,73,3))),('' if X.truthy(X.eq(E.v(3,76,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,76,3))))

def f20(E, r, c):
    # Ô C20
    # = "Câu sách chương 22: "&IF(LEN(KET_LUAN!$C$74)>900,LEFT(KET_LUAN!$C$74,900)&"…",KET_LUAN!$C$74)
    return X.cat('Câu sách chương 22: ',(X.cat(X.LEFT(E.v(3,74,3), 900.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,74,3)),900.0)) else E.v(3,74,3)))

def f21(E, r, c):
    # Ô C21
    # = "Trí tuệ – năng khiếu: "&IF(LEN(KET_LUAN!$C$71)>500,LEFT(KET_LUAN!$C$71,500)&"…",KET_LUAN!$C$71)
    return X.cat('Trí tuệ – năng khiếu: ',(X.cat(X.LEFT(E.v(3,71,3), 500.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,71,3)),500.0)) else E.v(3,71,3)))

def f22(E, r, c):
    # Ô C22
    # = "Nhóm thập thần nổi trội: "&INDEX(NHOM_TEN,MATCH(MAX(PP1_DINH_TINH!$F$96:$F$100),PP1_DINH_TINH!$F$96:$F$100,0)-1+1)&" ("&MAX(PP1_DINH_TINH!$F$96:$F$10
    #   0)&" chữ) → "&CHOOSE(MATCH(MAX(PP1_DINH_TINH!$F$96:$F$100),PP1_DINH_TINH!$F$96:$F$100,0)-1+1,"tự chủ, ý chí mạnh, coi trọng bạn bè – anh em, thích cạn
    #   h tranh; mặt trái: dễ cố chấp, chi tiêu theo cảm hứng","giàu ý tưởng, diễn đạt tốt, thích tự do sáng tạo; mặt trái: dễ chán việc lặp lại, không thích 
    #   bị áp đặt, hay va chạm cấp trên","thực tế, nhạy với cơ hội và tiền bạc, giỏi xoay xở; mặt trái: dễ ôm đồm, nặng lo toan vật chất","trách nhiệm, kỷ luậ
    #   t, coi trọng danh dự – địa vị; mặt trái: dễ căng thẳng, khắt khe với bản thân","ham học, suy tư, giàu lòng trắc ẩn, cần cảm giác an toàn; mặt trái: dễ
    #    thụ động, chần chừ, ỷ lại")&IF(COUNTIF(PP1_DINH_TINH!$F$96:$F$100,0)>0,". Thiếu hẳn nhóm "&INDEX(NHOM_TEN,MATCH(0,PP1_DINH_TINH!$F$96:$F$100,0))&" → 
    #   "&CHOOSE(MATCH(0,PP1_DINH_TINH!$F$96:$F$100,0)-1+1,"ít người đồng hành, quen tự làm một mình","ít bộc lộ, cần rèn kỹ năng trình bày","chưa nhạy tiền b
    #   ạc, nên học quản lý tài chính sớm","ít chịu khuôn khổ, cần tự đặt kỷ luật","ít chỗ dựa, phải tự lập sớm"),"")&". (Đọc thập thần theo tâm lý – kinh ngh
    #   iệm thực hành; số chữ lấy ở PP1_DINH_TINH mục G.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhóm thập thần nổi trội: ',X.INDEX(E.rg(24,54,1,58,1), X.add(X.sub(X.MATCH(X.MAX(E.rg(9,96,6,100,6)), E.rg(9,96,6,100,6), 0.0),1.0),1.0))),' ('),X.MAX(E.rg(9,96,6,100,6))),' chữ) → '),X.CHOOSE(X.add(X.sub(X.MATCH(X.MAX(E.rg(9,96,6,100,6)), E.rg(9,96,6,100,6), 0.0),1.0),1.0), lambda: 'tự chủ, ý chí mạnh, coi trọng bạn bè – anh em, thích cạnh tranh; mặt trái: dễ cố chấp, chi tiêu theo cảm hứng', lambda: 'giàu ý tưởng, diễn đạt tốt, thích tự do sáng tạo; mặt trái: dễ chán việc lặp lại, không thích bị áp đặt, hay va chạm cấp trên', lambda: 'thực tế, nhạy với cơ hội và tiền bạc, giỏi xoay xở; mặt trái: dễ ôm đồm, nặng lo toan vật chất', lambda: 'trách nhiệm, kỷ luật, coi trọng danh dự – địa vị; mặt trái: dễ căng thẳng, khắt khe với bản thân', lambda: 'ham học, suy tư, giàu lòng trắc ẩn, cần cảm giác an toàn; mặt trái: dễ thụ động, chần chừ, ỷ lại')),(X.cat(X.cat(X.cat('. Thiếu hẳn nhóm ',X.INDEX(E.rg(24,54,1,58,1), X.MATCH(0.0, E.rg(9,96,6,100,6), 0.0))),' → '),X.CHOOSE(X.add(X.sub(X.MATCH(0.0, E.rg(9,96,6,100,6), 0.0),1.0),1.0), lambda: 'ít người đồng hành, quen tự làm một mình', lambda: 'ít bộc lộ, cần rèn kỹ năng trình bày', lambda: 'chưa nhạy tiền bạc, nên học quản lý tài chính sớm', lambda: 'ít chịu khuôn khổ, cần tự đặt kỷ luật', lambda: 'ít chỗ dựa, phải tự lập sớm')) if X.truthy(X.gt(X.COUNTIF(E.rg(9,96,6,100,6), 0.0),0.0)) else '')),'. (Đọc thập thần theo tâm lý – kinh nghiệm thực hành; số chữ lấy ở PP1_DINH_TINH mục G.)')

def f23(E, r, c):
    # Ô C24
    # = "• Học tập – thi cử: "&IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8)&CHAR(10)&"• Quan vận – chức vụ: "&IFERROR(LEFT(SU_
    #   NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7)&CHAR(10)&"• Danh tiếng: "&IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU
    #   _NGHIEP!$J$9)&CHAR(10)&"• Sang – hèn: "&IFERROR(LEFT(MENH_CUC!$D$19,FIND(" (",MENH_CUC!$D$19)-1),MENH_CUC!$D$19)&CHAR(10)&"• Nghề hợp dụng thần (P2): 
    #   "&IF(LEN(TO_CHAT!$C$17)>200,LEFT(TO_CHAT!$C$17,200)&"…",TO_CHAT!$C$17)&CHAR(10)&"• Môi trường làm việc (Hóa giải toàn thư): "&IF(KET_LUAN!$C$102="","—
    #   ",IF(LEN(KET_LUAN!$C$102)>420,LEFT(KET_LUAN!$C$102,420)&"…",KET_LUAN!$C$102))&CHAR(10)&"• Đề bạt (Hóa giải toàn thư): "&IF(KET_LUAN!$C$103="","không c
    #   ó dấu hiệu riêng",IF(LEN(KET_LUAN!$C$103)>260,LEFT(KET_LUAN!$C$103,260)&"…",KET_LUAN!$C$103))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Học tập – thi cử: ',X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10))),X.CHAR(10.0)),'• Quan vận – chức vụ: '),X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10))),X.CHAR(10.0)),'• Danh tiếng: '),X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10))),X.CHAR(10.0)),'• Sang – hèn: '),X.IFERROR(lambda: X.LEFT(E.v(18,19,4), X.sub(X.FIND(' (', E.v(18,19,4)),1.0)), lambda: E.v(18,19,4))),X.CHAR(10.0)),'• Nghề hợp dụng thần (P2): '),(X.cat(X.LEFT(E.v(20,17,3), 200.0),'…') if X.truthy(X.gt(X.LEN(E.v(20,17,3)),200.0)) else E.v(20,17,3))),X.CHAR(10.0)),'• Môi trường làm việc (Hóa giải toàn thư): '),('—' if X.truthy(X.eq(E.v(3,102,3),'')) else (X.cat(X.LEFT(E.v(3,102,3), 420.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,102,3)),420.0)) else E.v(3,102,3)))),X.CHAR(10.0)),'• Đề bạt (Hóa giải toàn thư): '),('không có dấu hiệu riêng' if X.truthy(X.eq(E.v(3,103,3),'')) else (X.cat(X.LEFT(E.v(3,103,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,103,3)),260.0)) else E.v(3,103,3))))

def f24(E, r, c):
    # Ô J25
    # = IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8),5)="THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_N
    #   GHIEP!$J$8)-1),SU_NGHIEP!$J$8),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8),3)="TỐT",IFERROR(
    #   LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHI
    #   EP!$J$8)-1),SU_NGHIEP!$J$8))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))),ISNUMBER(SEARCH
    #   ("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHA
    #   R(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8)))),"• Học tập: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8),FI
    #   ND(" (",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))-1),IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),
    #   SU_NGHIEP!$J$8))&" ","")&IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7),5)="THUẬN",LEFT(IFERROR(LEFT(SU_NGHIE
    #   P!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIE
    #   P!$J$7),3)="TỐT",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SU_NGHIEP!$
    #   J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIE
    #   P!$J$7))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEF
    #   T(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7)))),"• Quan vận: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$
    #   J$7)-1),SU_NGHIEP!$J$7),FIND(" (",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))-1),IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CH
    #   AR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))&" ","")&IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),5)="THUẬN",
    #   LEFT(IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),S
    #   U_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),3)="TỐT",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9)="Khá",ISNUMBER(SEARCH("VỀ GIÀU
    #   ",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),S
    #   U_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))),ISNUMBER(S
    #   EARCH("VỀ THỌ",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9)))),"• Danh tiếng: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J
    #   $9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),FIND(" (",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))-1),IFERROR
    #   (LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))&" ","")&IF(OR(LEFT(MENH_CUC!$D$19,5)="THUẬN",LEFT(MENH_CUC!$D$19,13)="NGHIÊNG T
    #   HUẬN",LEFT(MENH_CUC!$D$19,3)="TỐT",MENH_CUC!$D$19="Khá",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$19)),ISNUMBER
    #   (SEARCH("VỀ CÁT",MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$19))),"• Sang – hèn: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$19,FIND(" (",MENH_CUC!$D$
    #   19)-1),MENH_CUC!$D$19))&" ","")&KET_LUAN!$C$105&" "
    return X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• Học tập: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))))) else ''),(X.cat(X.cat('• Quan vận: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))))) else '')),(X.cat(X.cat('• Danh tiếng: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))))) else '')),(X.cat(X.cat('• Sang – hèn: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,19,4), X.sub(X.FIND(' (', E.v(18,19,4)),1.0)), lambda: E.v(18,19,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,19,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,19,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,19,4), 3.0),'TỐT'), X.eq(E.v(18,19,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,19,4))))) else '')),E.v(3,105,3)),' ')

def f25(E, r, c):
    # Ô C26
    # = "Học tập – thi cử: "&SUBSTITUTE(IFERROR(LEFT(KET_LUAN!$C$96,MIN(330,FIND("Vì sao",KET_LUAN!$C$96)-1)),IF(LEN(KET_LUAN!$C$96)>330,LEFT(KET_LUAN!$C$96,3
    #   30)&"…",KET_LUAN!$C$96)),CHAR(10)," ")
    return X.cat('Học tập – thi cử: ',X.SUBSTITUTE(X.IFERROR(lambda: X.LEFT(E.v(3,96,3), X.MIN(330.0, X.sub(X.FIND('Vì sao', E.v(3,96,3)),1.0))), lambda: (X.cat(X.LEFT(E.v(3,96,3), 330.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,96,3)),330.0)) else E.v(3,96,3))), X.CHAR(10.0), ' '))

def f26(E, r, c):
    # Ô C27
    # = "Quan vận – chức vụ: "&SUBSTITUTE(IFERROR(LEFT(KET_LUAN!$C$95,MIN(330,FIND("Vì sao",KET_LUAN!$C$95)-1)),IF(LEN(KET_LUAN!$C$95)>330,LEFT(KET_LUAN!$C$95
    #   ,330)&"…",KET_LUAN!$C$95)),CHAR(10)," ")&CHAR(10)&"Danh tiếng: "&SUBSTITUTE(IFERROR(LEFT(KET_LUAN!$C$97,MIN(200,FIND("Vì sao",KET_LUAN!$C$97)-1)),IF(L
    #   EN(KET_LUAN!$C$97)>200,LEFT(KET_LUAN!$C$97,200)&"…",KET_LUAN!$C$97)),CHAR(10)," ")&IF(KET_LUAN!$C$101="","",CHAR(10)&"Bổ sung (Uyên Hải Tử Bình): "&KE
    #   T_LUAN!$C$101)
    return X.cat(X.cat(X.cat(X.cat(X.cat('Quan vận – chức vụ: ',X.SUBSTITUTE(X.IFERROR(lambda: X.LEFT(E.v(3,95,3), X.MIN(330.0, X.sub(X.FIND('Vì sao', E.v(3,95,3)),1.0))), lambda: (X.cat(X.LEFT(E.v(3,95,3), 330.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,95,3)),330.0)) else E.v(3,95,3))), X.CHAR(10.0), ' ')),X.CHAR(10.0)),'Danh tiếng: '),X.SUBSTITUTE(X.IFERROR(lambda: X.LEFT(E.v(3,97,3), X.MIN(200.0, X.sub(X.FIND('Vì sao', E.v(3,97,3)),1.0))), lambda: (X.cat(X.LEFT(E.v(3,97,3), 200.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,97,3)),200.0)) else E.v(3,97,3))), X.CHAR(10.0), ' ')),('' if X.truthy(X.eq(E.v(3,101,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,101,3))))

def f27(E, r, c):
    # Ô C28
    # = "Sang – hèn: "&IF(LEN(KET_LUAN!$C$99)>260,LEFT(KET_LUAN!$C$99,260)&"…",KET_LUAN!$C$99)
    return X.cat('Sang – hèn: ',(X.cat(X.LEFT(E.v(3,99,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,99,3)),260.0)) else E.v(3,99,3)))

def f28(E, r, c):
    # Ô C29
    # = "Thiên hướng nghề (sách): "&IF(LEN(KET_LUAN!$C$72)>700,LEFT(KET_LUAN!$C$72,700)&"…",KET_LUAN!$C$72)
    return X.cat('Thiên hướng nghề (sách): ',(X.cat(X.LEFT(E.v(3,72,3), 700.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,72,3)),700.0)) else E.v(3,72,3)))

def f29(E, r, c):
    # Ô C30
    # = "Môi trường giúp phát triển (theo nhóm dụng thần "&INDEX(NHOM_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7+5,5)+1
    #   )&"): "&CHOOSE(MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7+5,5)+1,"làm việc nhóm, hợp tác, có bạn đồng hành; rèn ý c
    #   hí, thể thao","phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình","công việc gắn với kết quả thực tế, quản lý tiền, kinh doan
    #   h","tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý","học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi trường ổn đ
    #   ịnh")&". Nghề hiện đại hợp hành dụng thần "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1)&": "&CHOOSE(IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1,"giáo dục – đào tạo, y dược, xuất bản – văn hóa, nông lâm, môi trường, thiết kế – nội thất gỗ","điện – điện 
    #   tử, công nghệ thông tin, năng lượng, truyền thông – quảng cáo, giải trí, ẩm thực, làm đẹp","bất động sản, xây dựng, nông nghiệp, quản lý – hành chính,
    #    bảo hiểm, tư vấn, kho bãi","tài chính – ngân hàng, cơ khí – ô tô, kỹ thuật, pháp luật – thanh tra, quân đội – công an, thiết bị","thương mại – xuất n
    #   hập khẩu, logistics – vận tải, du lịch, truyền thông số, nghiên cứu, dịch vụ lưu động")&"; hợp hành hỷ thần "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG
    #   _HY_KY!$D$110,DUNG_HY_KY!$C$110)+1)&": "&CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1,"giáo dục – đào tạo, y dược, xuất bản – vă
    #   n hóa, nông lâm, môi trường, thiết kế – nội thất gỗ","điện – điện tử, công nghệ thông tin, năng lượng, truyền thông – quảng cáo, giải trí, ẩm thực, là
    #   m đẹp","bất động sản, xây dựng, nông nghiệp, quản lý – hành chính, bảo hiểm, tư vấn, kho bãi","tài chính – ngân hàng, cơ khí – ô tô, kỹ thuật, pháp lu
    #   ật – thanh tra, quân đội – công an, thiết bị","thương mại – xuất nhập khẩu, logistics – vận tải, du lịch, truyền thông số, nghiên cứu, dịch vụ lưu độn
    #   g")&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Môi trường giúp phát triển (theo nhóm dụng thần ',X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)),5.0), 5.0),1.0))),'): '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'làm việc nhóm, hợp tác, có bạn đồng hành; rèn ý chí, thể thao', lambda: 'phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình', lambda: 'công việc gắn với kết quả thực tế, quản lý tiền, kinh doanh', lambda: 'tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý', lambda: 'học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi trường ổn định')),'. Nghề hiện đại hợp hành dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),': '),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'giáo dục – đào tạo, y dược, xuất bản – văn hóa, nông lâm, môi trường, thiết kế – nội thất gỗ', lambda: 'điện – điện tử, công nghệ thông tin, năng lượng, truyền thông – quảng cáo, giải trí, ẩm thực, làm đẹp', lambda: 'bất động sản, xây dựng, nông nghiệp, quản lý – hành chính, bảo hiểm, tư vấn, kho bãi', lambda: 'tài chính – ngân hàng, cơ khí – ô tô, kỹ thuật, pháp luật – thanh tra, quân đội – công an, thiết bị', lambda: 'thương mại – xuất nhập khẩu, logistics – vận tải, du lịch, truyền thông số, nghiên cứu, dịch vụ lưu động')),'; hợp hành hỷ thần '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))),': '),X.CHOOSE(X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0), lambda: 'giáo dục – đào tạo, y dược, xuất bản – văn hóa, nông lâm, môi trường, thiết kế – nội thất gỗ', lambda: 'điện – điện tử, công nghệ thông tin, năng lượng, truyền thông – quảng cáo, giải trí, ẩm thực, làm đẹp', lambda: 'bất động sản, xây dựng, nông nghiệp, quản lý – hành chính, bảo hiểm, tư vấn, kho bãi', lambda: 'tài chính – ngân hàng, cơ khí – ô tô, kỹ thuật, pháp luật – thanh tra, quân đội – công an, thiết bị', lambda: 'thương mại – xuất nhập khẩu, logistics – vận tải, du lịch, truyền thông số, nghiên cứu, dịch vụ lưu động')),'.')

def f30(E, r, c):
    # Ô C31
    # = IF(""&IF(AND(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_
    #   HY_KY!$D$110,DUNG_HY_KY!$C$110)),OR(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GI
    #   AN!$E$7+1,5),MOD(TRUNG_GIAN!$E$7+1,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+1,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$110,DUNG_HY_KY!$C$110)))),"• Quan sát là dụng/hỷ nhưng thực thương là kỵ: tài năng mạnh mà dễ va chạm cấp trên; công danh lên xuống theo năm q
    #   uan tinh (lên) và năm thực thương (xuống) – nên chọn vai trò chuyên gia, quản lý chuyên môn, giữ hòa khí với cấp trên. "&CHAR(10),"")&IF(OR(MOD(TRUNG_
    #   GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_
    #   CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"• Quan sát l
    #   à kỵ: tránh môi trường áp lực cao, cấp bậc cứng nhắc; hợp làm chuyên môn độc lập hoặc nơi có người nâng đỡ. "&CHAR(10),"")&IF(OR(MOD(TRUNG_GIAN!$E$7+4
    #   ,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"• Ấ
    #   n là dụng/hỷ: học vấn, bằng cấp, người thầy là đòn bẩy chính – nên học suốt đời. "&CHAR(10),"")&IF(OR(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DU
    #   NG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"• Thực thương là dụng/hỷ: phát
    #    triển bằng kỹ năng, sản phẩm, sáng tạo; có thể làm nghề tự do, kinh doanh nhỏ dựa trên tay nghề. "&CHAR(10),"")&IF(AND(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF
    #   (DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUN
    #   G_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),PP1_DINH_TINH!$F$98>=P_NHIEU),
    #   IF(DOI_CHIEU!$C$33=2,IF(DUNG_HY_KY!$D$7=0,"• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. "&CHAR(10),""),IF(DUNG_
    #   HY_KY!$C$7=0,"• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. "&CHAR(10),"")),"")="","(Không có mẫu kinh nghiệm đặ
    #   c biệt.)",""&IF(AND(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=
    #   2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),OR(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(T
    #   RUNG_GIAN!$E$7+1,5),MOD(TRUNG_GIAN!$E$7+1,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+1,5)<>IF(DOI_CHIEU!$C$33=2
    #   ,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"• Quan sát là dụng/hỷ nhưng thực thương là kỵ: tài năng mạnh mà dễ va chạm cấp trên; công danh lên xuống the
    #   o năm quan tinh (lên) và năm thực thương (xuống) – nên chọn vai trò chuyên gia, quản lý chuyên môn, giữ hòa khí với cấp trên. "&CHAR(10),"")&IF(OR(MOD
    #   (TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"• Qua
    #   n sát là kỵ: tránh môi trường áp lực cao, cấp bậc cứng nhắc; hợp làm chuyên môn độc lập hoặc nơi có người nâng đỡ. "&CHAR(10),"")&IF(OR(MOD(TRUNG_GIAN
    #   !$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110
    #   )),"• Ấn là dụng/hỷ: học vấn, bằng cấp, người thầy là đòn bẩy chính – nên học suốt đời. "&CHAR(10),"")&IF(OR(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$
    #   33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"• Thực thương là dụng/h
    #   ỷ: phát triển bằng kỹ năng, sản phẩm, sáng tạo; có thể làm nghề tự do, kinh doanh nhỏ dựa trên tay nghề. "&CHAR(10),"")&IF(AND(OR(MOD(TRUNG_GIAN!$E$7+
    #   2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$3
    #   3=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),PP1_DINH_TINH!$F$98>=P_
    #   NHIEU),IF(DOI_CHIEU!$C$33=2,IF(DUNG_HY_KY!$D$7=0,"• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. "&CHAR(10),""),I
    #   F(DUNG_HY_KY!$C$7=0,"• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. "&CHAR(10),"")),""))
    return ('(Không có mẫu kinh nghiệm đặc biệt.)' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat('',(X.cat('• Quan sát là dụng/hỷ nhưng thực thương là kỵ: tài năng mạnh mà dễ va chạm cấp trên; công danh lên xuống theo năm quan tinh (lên) và năm thực thương (xuống) – nên chọn vai trò chuyên gia, quản lý chuyên môn, giữ hòa khí với cấp trên. ',X.CHAR(10.0)) if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),1.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else '')),(X.cat('• Quan sát là kỵ: tránh môi trường áp lực cao, cấp bậc cứng nhắc; hợp làm chuyên môn độc lập hoặc nơi có người nâng đỡ. ',X.CHAR(10.0)) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')),(X.cat('• Ấn là dụng/hỷ: học vấn, bằng cấp, người thầy là đòn bẩy chính – nên học suốt đời. ',X.CHAR(10.0)) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat('• Thực thương là dụng/hỷ: phát triển bằng kỹ năng, sản phẩm, sáng tạo; có thể làm nghề tự do, kinh doanh nhỏ dựa trên tay nghề. ',X.CHAR(10.0)) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(((X.cat('• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. ',X.CHAR(10.0)) if X.truthy(X.eq(E.v(8,7,4),0.0)) else '') if X.truthy(X.eq(E.v(7,33,3),2.0)) else (X.cat('• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. ',X.CHAR(10.0)) if X.truthy(X.eq(E.v(8,7,3),0.0)) else '')) if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '')),'')) else X.cat(X.cat(X.cat(X.cat(X.cat('',(X.cat('• Quan sát là dụng/hỷ nhưng thực thương là kỵ: tài năng mạnh mà dễ va chạm cấp trên; công danh lên xuống theo năm quan tinh (lên) và năm thực thương (xuống) – nên chọn vai trò chuyên gia, quản lý chuyên môn, giữ hòa khí với cấp trên. ',X.CHAR(10.0)) if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),1.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else '')),(X.cat('• Quan sát là kỵ: tránh môi trường áp lực cao, cấp bậc cứng nhắc; hợp làm chuyên môn độc lập hoặc nơi có người nâng đỡ. ',X.CHAR(10.0)) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')),(X.cat('• Ấn là dụng/hỷ: học vấn, bằng cấp, người thầy là đòn bẩy chính – nên học suốt đời. ',X.CHAR(10.0)) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat('• Thực thương là dụng/hỷ: phát triển bằng kỹ năng, sản phẩm, sáng tạo; có thể làm nghề tự do, kinh doanh nhỏ dựa trên tay nghề. ',X.CHAR(10.0)) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(((X.cat('• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. ',X.CHAR(10.0)) if X.truthy(X.eq(E.v(8,7,4),0.0)) else '') if X.truthy(X.eq(E.v(7,33,3),2.0)) else (X.cat('• Tài nhiều thân nhược: nhận việc vừa sức, tránh ôm đồm; hợp tác với người tin cậy. ',X.CHAR(10.0)) if X.truthy(X.eq(E.v(8,7,3),0.0)) else '')) if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '')))

def f31(E, r, c):
    # Ô C32
    # = "MỐC CÔNG DANH (theo can năm – P1 tr.167-169; P2 tr.541-542; chi tiết sheet THEO_DOI)"&CHAR(10)&"↑ Năm thuận công danh – đã qua: "&IF(THEO_DOI!$AG$111
    #   ="","không có",THEO_DOI!$AG$111)&CHAR(10)&"   Sắp tới (20 năm): "&IF(THEO_DOI!$AH$111="","không có",THEO_DOI!$AH$111)&CHAR(10)&"↓ Năm dễ biến động chứ
    #   c vụ / áp lực – đã qua: "&IF(THEO_DOI!$AI$111="","không có",THEO_DOI!$AI$111)&CHAR(10)&"   Sắp tới (20 năm): "&IF(THEO_DOI!$AJ$111="","không có",THEO_
    #   DOI!$AJ$111)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('MỐC CÔNG DANH (theo can năm – P1 tr.167-169; P2 tr.541-542; chi tiết sheet THEO_DOI)',X.CHAR(10.0)),'↑ Năm thuận công danh – đã qua: '),('không có' if X.truthy(X.eq(E.v(5,111,33),'')) else E.v(5,111,33))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,34),'')) else E.v(5,111,34))),X.CHAR(10.0)),'↓ Năm dễ biến động chức vụ / áp lực – đã qua: '),('không có' if X.truthy(X.eq(E.v(5,111,35),'')) else E.v(5,111,35))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,36),'')) else E.v(5,111,36)))

def f32(E, r, c):
    # Ô C33
    # = "MỐC HỌC TẬP (năm ấn là dụng/hỷ, tuổi 6–30) – đã qua: "&IF(THEO_DOI!$AO$111="","không có",THEO_DOI!$AO$111)&CHAR(10)&"   Sắp tới (20 năm): "&IF(THEO_D
    #   OI!$AP$111="","không có",THEO_DOI!$AP$111)
    return X.cat(X.cat(X.cat(X.cat('MỐC HỌC TẬP (năm ấn là dụng/hỷ, tuổi 6–30) – đã qua: ',('không có' if X.truthy(X.eq(E.v(5,111,41),'')) else E.v(5,111,41))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,42),'')) else E.v(5,111,42)))

def f33(E, r, c):
    # Ô C35
    # = "• Bậc giàu: "&TAI_LOC!$E$17&CHAR(10)&"• Tài vận: "&IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6)&CHAR(10)&"• Giàu – ng
    #   hèo: "&IFERROR(LEFT(MENH_CUC!$D$18,FIND(" (",MENH_CUC!$D$18)-1),MENH_CUC!$D$18)&CHAR(10)&"• Vận phát tài: "&IF(TAI_LOC!$C$37="","chưa thấy vận phát tà
    #   i rõ",IF(LEN(TAI_LOC!$C$37)>320,LEFT(TAI_LOC!$C$37,320)&"…",TAI_LOC!$C$37))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Bậc giàu: ',E.v(19,17,5)),X.CHAR(10.0)),'• Tài vận: '),X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10))),X.CHAR(10.0)),'• Giàu – nghèo: '),X.IFERROR(lambda: X.LEFT(E.v(18,18,4), X.sub(X.FIND(' (', E.v(18,18,4)),1.0)), lambda: E.v(18,18,4))),X.CHAR(10.0)),'• Vận phát tài: '),('chưa thấy vận phát tài rõ' if X.truthy(X.eq(E.v(19,37,3),'')) else (X.cat(X.LEFT(E.v(19,37,3), 320.0),'…') if X.truthy(X.gt(X.LEN(E.v(19,37,3)),320.0)) else E.v(19,37,3))))

def f34(E, r, c):
    # Ô J36
    # = IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),5)="THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_N
    #   GHIEP!$J$6)-1),SU_NGHIEP!$J$6),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),3)="TỐT",IFERROR(
    #   LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHI
    #   EP!$J$6)-1),SU_NGHIEP!$J$6))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))),ISNUMBER(SEARCH
    #   ("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHA
    #   R(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6)))),"• Tài vận: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),FI
    #   ND(" (",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))-1),IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),
    #   SU_NGHIEP!$J$6))&" ","")&IF(OR(LEFT(MENH_CUC!$D$18,5)="THUẬN",LEFT(MENH_CUC!$D$18,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$18,3)="TỐT",MENH_CUC!$D$18="Khá
    #   ",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ 
    #   THỌ",MENH_CUC!$D$18))),"• Giàu – nghèo: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$18,FIND(" (",MENH_CUC!$D$18)-1),MENH_CUC!$D$18))&" ","")&IF(TAI_LOC!$C$37="",
    #   "","• Vận phát tài: "&IF(LEN(TAI_LOC!$C$37)>320,LEFT(TAI_LOC!$C$37,320)&"…",TAI_LOC!$C$37)&" ")
    return X.cat(X.cat((X.cat(X.cat('• Tài vận: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))))) else ''),(X.cat(X.cat('• Giàu – nghèo: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,18,4), X.sub(X.FIND(' (', E.v(18,18,4)),1.0)), lambda: E.v(18,18,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,18,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,18,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,18,4), 3.0),'TỐT'), X.eq(E.v(18,18,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,18,4))))) else '')),('' if X.truthy(X.eq(E.v(19,37,3),'')) else X.cat(X.cat('• Vận phát tài: ',(X.cat(X.LEFT(E.v(19,37,3), 320.0),'…') if X.truthy(X.gt(X.LEN(E.v(19,37,3)),320.0)) else E.v(19,37,3))),' ')))

def f35(E, r, c):
    # Ô C37
    # = "Bậc giàu: "&IF(LEN(KET_LUAN!$C$81)>180,LEFT(KET_LUAN!$C$81,180)&"…",KET_LUAN!$C$81)&CHAR(10)&"Giàu – nghèo: "&IF(LEN(KET_LUAN!$C$87)>200,LEFT(KET_LUA
    #   N!$C$87,200)&"…",KET_LUAN!$C$87)
    return X.cat(X.cat(X.cat(X.cat('Bậc giàu: ',(X.cat(X.LEFT(E.v(3,81,3), 180.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,81,3)),180.0)) else E.v(3,81,3))),X.CHAR(10.0)),'Giàu – nghèo: '),(X.cat(X.LEFT(E.v(3,87,3), 200.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,87,3)),200.0)) else E.v(3,87,3)))

def f36(E, r, c):
    # Ô C38
    # = "Tài vận: "&SUBSTITUTE(IFERROR(LEFT(KET_LUAN!$C$88,MIN(380,FIND("Vì sao",KET_LUAN!$C$88)-1)),IF(LEN(KET_LUAN!$C$88)>380,LEFT(KET_LUAN!$C$88,380)&"…",K
    #   ET_LUAN!$C$88)),CHAR(10)," ")
    return X.cat('Tài vận: ',X.SUBSTITUTE(X.IFERROR(lambda: X.LEFT(E.v(3,88,3), X.MIN(380.0, X.sub(X.FIND('Vì sao', E.v(3,88,3)),1.0))), lambda: (X.cat(X.LEFT(E.v(3,88,3), 380.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,88,3)),380.0)) else E.v(3,88,3))), X.CHAR(10.0), ' '))

def f37(E, r, c):
    # Ô C39
    # = "Nguồn tiền: "&IF(LEN(KET_LUAN!$C$83)>600,LEFT(KET_LUAN!$C$83,600)&"…",KET_LUAN!$C$83)&IF(KET_LUAN!$C$90="","",CHAR(10)&"Bổ sung (Uyên Hải Tử Bình): "
    #   &KET_LUAN!$C$90)
    return X.cat(X.cat('Nguồn tiền: ',(X.cat(X.LEFT(E.v(3,83,3), 600.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,83,3)),600.0)) else E.v(3,83,3))),('' if X.truthy(X.eq(E.v(3,90,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,90,3))))

def f38(E, r, c):
    # Ô C40
    # = "Vận phát tài / cẩn trọng: "&IF(LEN(KET_LUAN!$C$85)>400,LEFT(KET_LUAN!$C$85,400)&"…",KET_LUAN!$C$85)
    return X.cat('Vận phát tài / cẩn trọng: ',(X.cat(X.LEFT(E.v(3,85,3), 400.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,85,3)),400.0)) else E.v(3,85,3)))

def f39(E, r, c):
    # Ô C41
    # = IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$1
    #   10,DUNG_HY_KY!$C$110)),"Tài là dụng/hỷ: chủ động tạo thêm nguồn thu, tích lũy tài sản thực (nhà, đất, tài sản sinh lời) vào các vận/năm tài; vẫn giữ q
    #   uỹ dự phòng.","Tài không phải dụng/hỷ: ưu tiên thu nhập ổn định (lương, chuyên môn), lập quỹ dự phòng 6–12 tháng, không vay để đầu tư, không đứng ra b
    #   ảo lãnh; giao tài sản lớn cho người thân tín cùng quản lý.")&IF(AND(OR(MOD(TRUNG_GIAN!$E$7+0,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$1
    #   11),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+0,5),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$
    #   7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$1
    #   09),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))," Tỉ kiếp là kỵ: tránh hùn vốn, cho bạn bè vay lớn.","")
    return X.cat(('Tài là dụng/hỷ: chủ động tạo thêm nguồn thu, tích lũy tài sản thực (nhà, đất, tài sản sinh lời) vào các vận/năm tài; vẫn giữ quỹ dự phòng.' if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'Tài không phải dụng/hỷ: ưu tiên thu nhập ổn định (lương, chuyên môn), lập quỹ dự phòng 6–12 tháng, không vay để đầu tư, không đứng ra bảo lãnh; giao tài sản lớn cho người thân tín cùng quản lý.'),(' Tỉ kiếp là kỵ: tránh hùn vốn, cho bạn bè vay lớn.' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else ''))

def f40(E, r, c):
    # Ô C42
    # = "MỐC TÀI CHÍNH (P1 tr.156, 164-165; P2 tr.470, 521-527; chi tiết THEO_DOI)"&CHAR(10)&"↑ Năm thu nhập thuận – đã qua: "&IF(THEO_DOI!$AK$111="","không c
    #   ó",THEO_DOI!$AK$111)&CHAR(10)&"   Sắp tới (20 năm): "&IF(THEO_DOI!$AL$111="","không có",THEO_DOI!$AL$111)&CHAR(10)&"↓ Năm cần giữ tiền / dễ hao – đã q
    #   ua: "&IF(THEO_DOI!$AM$111="","không có",THEO_DOI!$AM$111)&CHAR(10)&"   Sắp tới (20 năm): "&IF(THEO_DOI!$AN$111="","không có",THEO_DOI!$AN$111)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('MỐC TÀI CHÍNH (P1 tr.156, 164-165; P2 tr.470, 521-527; chi tiết THEO_DOI)',X.CHAR(10.0)),'↑ Năm thu nhập thuận – đã qua: '),('không có' if X.truthy(X.eq(E.v(5,111,37),'')) else E.v(5,111,37))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,38),'')) else E.v(5,111,38))),X.CHAR(10.0)),'↓ Năm cần giữ tiền / dễ hao – đã qua: '),('không có' if X.truthy(X.eq(E.v(5,111,39),'')) else E.v(5,111,39))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,40),'')) else E.v(5,111,40)))

def f41(E, r, c):
    # Ô C44
    # = "• "&"Ông bà: "&LOWER(LUC_THAN!$X$30)&" · "&"Cha: "&LOWER(LUC_THAN!$X$31)&" · "&"Mẹ: "&LOWER(LUC_THAN!$X$32)&" · "&"Anh chị em: "&LOWER(LUC_THAN!$X$33
    #   )&" · "&"Con cái: "&LOWER(LUC_THAN!$X$34)&CHAR(10)&"• Hôn nhân: "&LOWER(HON_NHAN!$X$64)&" ("&HON_NHAN!$AG$64&" dấu hiệu bất lợi)."&CHAR(10)&"• Con đầu
    #    (P1 tr.159): "&LOWER(IFERROR(LEFT(KET_LUAN!$C$115,FIND(CHAR(10),KET_LUAN!$C$115)-1),KET_LUAN!$C$115))&CHAR(10)&"• Thời điểm kết hôn: "&IF(LEN(KET_LUA
    #   N!$C$119)>240,LEFT(KET_LUAN!$C$119,240)&"…",KET_LUAN!$C$119)&CHAR(10)&"• Con cái (Hóa giải toàn thư): "&IF(KET_LUAN!$C$128="","không có dấu hiệu riêng
    #   ",IF(LEN(KET_LUAN!$C$128)>450,LEFT(KET_LUAN!$C$128,450)&"…",KET_LUAN!$C$128))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ','Ông bà: '),X.LOWER(E.v(16,30,24))),' · '),'Cha: '),X.LOWER(E.v(16,31,24))),' · '),'Mẹ: '),X.LOWER(E.v(16,32,24))),' · '),'Anh chị em: '),X.LOWER(E.v(16,33,24))),' · '),'Con cái: '),X.LOWER(E.v(16,34,24))),X.CHAR(10.0)),'• Hôn nhân: '),X.LOWER(E.v(17,64,24))),' ('),E.v(17,64,33)),' dấu hiệu bất lợi).'),X.CHAR(10.0)),'• Con đầu (P1 tr.159): '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(3,115,3), X.sub(X.FIND(X.CHAR(10.0), E.v(3,115,3)),1.0)), lambda: E.v(3,115,3)))),X.CHAR(10.0)),'• Thời điểm kết hôn: '),(X.cat(X.LEFT(E.v(3,119,3), 240.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,119,3)),240.0)) else E.v(3,119,3))),X.CHAR(10.0)),'• Con cái (Hóa giải toàn thư): '),('không có dấu hiệu riêng' if X.truthy(X.eq(E.v(3,128,3),'')) else (X.cat(X.LEFT(E.v(3,128,3), 450.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,128,3)),450.0)) else E.v(3,128,3))))

def f42(E, r, c):
    # Ô J45
    # = IF(OR(LUC_THAN!$X$30="Tốt",LUC_THAN!$X$30="Khá"),"• Ông bà: "&LOWER(LUC_THAN!$X$30)&" ","")&IF(OR(LUC_THAN!$X$31="Tốt",LUC_THAN!$X$31="Khá"),"• Cha: "
    #   &LOWER(LUC_THAN!$X$31)&" ","")&IF(OR(LUC_THAN!$X$32="Tốt",LUC_THAN!$X$32="Khá"),"• Mẹ: "&LOWER(LUC_THAN!$X$32)&" ","")&IF(OR(LUC_THAN!$X$33="Tốt",LUC_
    #   THAN!$X$33="Khá"),"• Anh chị em: "&LOWER(LUC_THAN!$X$33)&" ","")&IF(OR(LUC_THAN!$X$34="Tốt",LUC_THAN!$X$34="Khá"),"• Con cái: "&LOWER(LUC_THAN!$X$34)&
    #   " ","")&IF(OR(HON_NHAN!$X$64="Tốt",HON_NHAN!$X$64="Khá"),"• Hôn nhân: "&LOWER(HON_NHAN!$X$64)&" ","")&KET_LUAN!$C$131&" "
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• Ông bà: ',X.LOWER(E.v(16,30,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,30,24),'Tốt'), X.eq(E.v(16,30,24),'Khá'))) else ''),(X.cat(X.cat('• Cha: ',X.LOWER(E.v(16,31,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,31,24),'Tốt'), X.eq(E.v(16,31,24),'Khá'))) else '')),(X.cat(X.cat('• Mẹ: ',X.LOWER(E.v(16,32,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,32,24),'Tốt'), X.eq(E.v(16,32,24),'Khá'))) else '')),(X.cat(X.cat('• Anh chị em: ',X.LOWER(E.v(16,33,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,33,24),'Tốt'), X.eq(E.v(16,33,24),'Khá'))) else '')),(X.cat(X.cat('• Con cái: ',X.LOWER(E.v(16,34,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,34,24),'Tốt'), X.eq(E.v(16,34,24),'Khá'))) else '')),(X.cat(X.cat('• Hôn nhân: ',X.LOWER(E.v(17,64,24))),' ') if X.truthy(X.OR(X.eq(E.v(17,64,24),'Tốt'), X.eq(E.v(17,64,24),'Khá'))) else '')),E.v(3,131,3)),' ')

def f43(E, r, c):
    # Ô C46
    # = "Ông bà – tổ nghiệp: "&IF(LEN(KET_LUAN!$C$110)>220,LEFT(KET_LUAN!$C$110,220)&"…",KET_LUAN!$C$110)&CHAR(10)&"Cha: "&IF(LEN(KET_LUAN!$C$111)>260,LEFT(KE
    #   T_LUAN!$C$111,260)&"…",KET_LUAN!$C$111)&CHAR(10)&"Mẹ: "&IF(LEN(KET_LUAN!$C$112)>260,LEFT(KET_LUAN!$C$112,260)&"…",KET_LUAN!$C$112)&IF(KET_LUAN!$C$125=
    #   "","",CHAR(10)&"Bổ sung (Uyên Hải Tử Bình): "&KET_LUAN!$C$125)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Ông bà – tổ nghiệp: ',(X.cat(X.LEFT(E.v(3,110,3), 220.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,110,3)),220.0)) else E.v(3,110,3))),X.CHAR(10.0)),'Cha: '),(X.cat(X.LEFT(E.v(3,111,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,111,3)),260.0)) else E.v(3,111,3))),X.CHAR(10.0)),'Mẹ: '),(X.cat(X.LEFT(E.v(3,112,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,112,3)),260.0)) else E.v(3,112,3))),('' if X.truthy(X.eq(E.v(3,125,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,125,3))))

def f44(E, r, c):
    # Ô C47
    # = "Năm có dấu hiệu bất lợi cho cha: "&LUC_THAN!$D$25&CHAR(10)&"Cho mẹ: "&LUC_THAN!$D$26&" (P2 tr.440, 446 – chỉ là dấu hiệu cần quan tâm)."
    return X.cat(X.cat(X.cat(X.cat(X.cat('Năm có dấu hiệu bất lợi cho cha: ',E.v(16,25,4)),X.CHAR(10.0)),'Cho mẹ: '),E.v(16,26,4)),' (P2 tr.440, 446 – chỉ là dấu hiệu cần quan tâm).')

def f45(E, r, c):
    # Ô C48
    # = "Anh chị em: "&IF(LEN(KET_LUAN!$C$113)>240,LEFT(KET_LUAN!$C$113,240)&"…",KET_LUAN!$C$113)&CHAR(10)&"Con cái: "&IF(LEN(KET_LUAN!$C$114)>260,LEFT(KET_LU
    #   AN!$C$114,260)&"…",KET_LUAN!$C$114)&CHAR(10)&"Con đầu lòng: "&IF(LEN(KET_LUAN!$C$115)>120,LEFT(KET_LUAN!$C$115,120)&"…",KET_LUAN!$C$115)&IF(KET_LUAN!$
    #   C$126="","",CHAR(10)&"Bổ sung (Uyên Hải Tử Bình): "&KET_LUAN!$C$126)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Anh chị em: ',(X.cat(X.LEFT(E.v(3,113,3), 240.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,113,3)),240.0)) else E.v(3,113,3))),X.CHAR(10.0)),'Con cái: '),(X.cat(X.LEFT(E.v(3,114,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,114,3)),260.0)) else E.v(3,114,3))),X.CHAR(10.0)),'Con đầu lòng: '),(X.cat(X.LEFT(E.v(3,115,3), 120.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,115,3)),120.0)) else E.v(3,115,3))),('' if X.truthy(X.eq(E.v(3,126,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,126,3))))

def f46(E, r, c):
    # Ô C49
    # = "Hôn nhân: "&IF(LEN(KET_LUAN!$C$117)>650,LEFT(KET_LUAN!$C$117,650)&"…",KET_LUAN!$C$117)&IF(KET_LUAN!$C$127="","",CHAR(10)&"Bổ sung (Uyên Hải Tử Bình):
    #    "&KET_LUAN!$C$127)
    return X.cat(X.cat('Hôn nhân: ',(X.cat(X.LEFT(E.v(3,117,3), 650.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,117,3)),650.0)) else E.v(3,117,3))),('' if X.truthy(X.eq(E.v(3,127,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,127,3))))

def f47(E, r, c):
    # Ô C50
    # = "Thời điểm kết hôn: "&IF(LEN(KET_LUAN!$C$119)>400,LEFT(KET_LUAN!$C$119,400)&"…",KET_LUAN!$C$119)
    return X.cat('Thời điểm kết hôn: ',(X.cat(X.LEFT(E.v(3,119,3), 400.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,119,3)),400.0)) else E.v(3,119,3)))

def f48(E, r, c):
    # Ô C51
    # = "Kinh nghiệm hợp hôn – gia đạo (“dụng thần tương bổ”): người có nhật can thuộc hành "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY
    #   !$C$109)+1)&" (dụng thần) hoặc "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1)&" (hỷ thần) thường bổ trợ tốt cho lá số n
    #   ày; nhật can thuộc hành "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1)&" (kỵ) dễ va chạm. Áp dụng cả với con cái, cộng 
    #   sự. Không có trong sách – cần xét đầy đủ lá số hai người."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Kinh nghiệm hợp hôn – gia đạo (“dụng thần tương bổ”): người có nhật can thuộc hành ',X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),' (dụng thần) hoặc '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))),' (hỷ thần) thường bổ trợ tốt cho lá số này; nhật can thuộc hành '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0))),' (kỵ) dễ va chạm. Áp dụng cả với con cái, cộng sự. Không có trong sách – cần xét đầy đủ lá số hai người.')

def f49(E, r, c):
    # Ô C52
    # = "Cung phối ngẫu (chi ngày, bản khí "&TRUNG_GIAN!$E$24&"): "&CHOOSE(MOD(TRUNG_GIAN!$E$21-TRUNG_GIAN!$E$7+5,5)+1,"phối ngẫu như người bạn ngang hàng, cá
    #    tính; dễ tranh luận, nên rạch ròi tiền bạc chung","phối ngẫu có tài, hoạt bát; lời qua tiếng lại nhiều, cần mềm mỏng","phối ngẫu thực tế, giỏi vun vé
    #   n kinh tế gia đình","phối ngẫu nề nếp, có địa vị hoặc nghiêm khắc; gia đình nhiều khuôn phép","phối ngẫu chu đáo, che chở như người thân; dễ ỷ lại nha
    #   u")&"; hành này "&IF(TRUNG_GIAN!$E$21=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần → chỗ dựa tốt",IF(TRUNG_GIAN!$E$21=IF(DO
    #   I_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần → hỗ trợ tốt",IF(TRUNG_GIAN!$E$21=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$
    #   C$111),"là kỵ thần → dễ va chạm, cần vun đắp","là nhàn thần → trung tính")))&"."&CHAR(10)&"Cung con cái – hậu vận (chi giờ, bản khí "&TRUNG_GIAN!$F$24
    #   &"): "&CHOOSE(MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!$E$7+5,5)+1,"con cái độc lập, có chính kiến; về già nên tự chủ","con cái thông minh, khéo léo, có năng k
    #   hiếu","con cái thực tế, về sau có điều kiện kinh tế","con cái nề nếp, có chí tiến thủ","con cái hiền, gần gũi, chăm lo cha mẹ")&"; hành này "&IF(TRUNG
    #   _GIAN!$F$21=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần → chỗ dựa tốt",IF(TRUNG_GIAN!$F$21=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY
    #   !$D$110,DUNG_HY_KY!$C$110),"là hỷ thần → hỗ trợ tốt",IF(TRUNG_GIAN!$F$21=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần → dễ va
    #    chạm, cần vun đắp","là nhàn thần → trung tính")))&". (Đọc cung theo thập thần bản khí – kinh nghiệm thực hành.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Cung phối ngẫu (chi ngày, bản khí ',E.v(23,24,5)),'): '),X.CHOOSE(X.add(X.MOD(X.add(X.sub(E.v(23,21,5),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'phối ngẫu như người bạn ngang hàng, cá tính; dễ tranh luận, nên rạch ròi tiền bạc chung', lambda: 'phối ngẫu có tài, hoạt bát; lời qua tiếng lại nhiều, cần mềm mỏng', lambda: 'phối ngẫu thực tế, giỏi vun vén kinh tế gia đình', lambda: 'phối ngẫu nề nếp, có địa vị hoặc nghiêm khắc; gia đình nhiều khuôn phép', lambda: 'phối ngẫu chu đáo, che chở như người thân; dễ ỷ lại nhau')),'; hành này '),('là dụng thần → chỗ dựa tốt' if X.truthy(X.eq(E.v(23,21,5),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần → hỗ trợ tốt' if X.truthy(X.eq(E.v(23,21,5),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần → dễ va chạm, cần vun đắp' if X.truthy(X.eq(E.v(23,21,5),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'là nhàn thần → trung tính')))),'.'),X.CHAR(10.0)),'Cung con cái – hậu vận (chi giờ, bản khí '),E.v(23,24,6)),'): '),X.CHOOSE(X.add(X.MOD(X.add(X.sub(E.v(23,21,6),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'con cái độc lập, có chính kiến; về già nên tự chủ', lambda: 'con cái thông minh, khéo léo, có năng khiếu', lambda: 'con cái thực tế, về sau có điều kiện kinh tế', lambda: 'con cái nề nếp, có chí tiến thủ', lambda: 'con cái hiền, gần gũi, chăm lo cha mẹ')),'; hành này '),('là dụng thần → chỗ dựa tốt' if X.truthy(X.eq(E.v(23,21,6),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần → hỗ trợ tốt' if X.truthy(X.eq(E.v(23,21,6),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần → dễ va chạm, cần vun đắp' if X.truthy(X.eq(E.v(23,21,6),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'là nhàn thần → trung tính')))),'. (Đọc cung theo thập thần bản khí – kinh nghiệm thực hành.)')

def f50(E, r, c):
    # Ô C53
    # = "MỐC HÔN NHÂN – năm thuận kết hôn / tin vui (P1 tr.201; tuổi 18–50) – đã qua: "&IF(THEO_DOI!$AQ$111="","không có",THEO_DOI!$AQ$111)&CHAR(10)&"   Sắp t
    #   ới (20 năm): "&IF(THEO_DOI!$AR$111="","không có",THEO_DOI!$AR$111)
    return X.cat(X.cat(X.cat(X.cat('MỐC HÔN NHÂN – năm thuận kết hôn / tin vui (P1 tr.201; tuổi 18–50) – đã qua: ',('không có' if X.truthy(X.eq(E.v(5,111,43),'')) else E.v(5,111,43))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,44),'')) else E.v(5,111,44)))

def f51(E, r, c):
    # Ô C54
    # = KET_LUAN!$C$122&CHAR(10)&KET_LUAN!$C$123&CHAR(10)&KET_LUAN!$C$124
    return X.cat(X.cat(X.cat(X.cat(E.v(3,122,3),X.CHAR(10.0)),E.v(3,123,3)),X.CHAR(10.0)),E.v(3,124,3))

def f52(E, r, c):
    # Ô C56
    # = "• "&"Sức khỏe chung: "&LOWER(IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6))&" · "&"Bệnh tật: "&LOWER(IFERROR(LEFT(SUC_KHO
    #   E!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7))&" · "&"Tai họa theo vận: "&LOWER(IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC
    #   _KHOE!$J$8))&CHAR(10)&"• Thọ – yểu: "&IFERROR(LEFT(MENH_CUC!$D$21,FIND(" (",MENH_CUC!$D$21)-1),MENH_CUC!$D$21)&CHAR(10)&"• Năm cần đề phòng: "&IFERROR
    #   (LEFT(KET_LUAN!$C$140,FIND(CHAR(10),KET_LUAN!$C$140)-1),KET_LUAN!$C$140)&CHAR(10)&"• Hóa giải toàn thư: "&IF(KET_LUAN!$C$156="","không có dấu hiệu riê
    #   ng",IF(LEN(KET_LUAN!$C$156)>300,LEFT(KET_LUAN!$C$156,300)&"…",KET_LUAN!$C$156))&CHAR(10)&"Lưu ý: đây là xu hướng theo sách mệnh lý, không phải chẩn đo
    #   án y khoa – sức khỏe cần khám và theo dõi của bác sĩ."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ','Sức khỏe chung: '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))),' · '),'Bệnh tật: '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)))),' · '),'Tai họa theo vận: '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)))),X.CHAR(10.0)),'• Thọ – yểu: '),X.IFERROR(lambda: X.LEFT(E.v(18,21,4), X.sub(X.FIND(' (', E.v(18,21,4)),1.0)), lambda: E.v(18,21,4))),X.CHAR(10.0)),'• Năm cần đề phòng: '),X.IFERROR(lambda: X.LEFT(E.v(3,140,3), X.sub(X.FIND(X.CHAR(10.0), E.v(3,140,3)),1.0)), lambda: E.v(3,140,3))),X.CHAR(10.0)),'• Hóa giải toàn thư: '),('không có dấu hiệu riêng' if X.truthy(X.eq(E.v(3,156,3),'')) else (X.cat(X.LEFT(E.v(3,156,3), 300.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,156,3)),300.0)) else E.v(3,156,3)))),X.CHAR(10.0)),'Lưu ý: đây là xu hướng theo sách mệnh lý, không phải chẩn đoán y khoa – sức khỏe cần khám và theo dõi của bác sĩ.')

def f53(E, r, c):
    # Ô J57
    # = IF(OR(LEFT(IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6),5)="THUẬN",LEFT(IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE
    #   !$J$6)-1),SUC_KHOE!$J$6),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6),3)="TỐT",IFERROR(LEFT(SUC_
    #   KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SU
    #   C_KHOE!$J$6))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(
    #   LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1
    #   ),SUC_KHOE!$J$6)))),"• Sức khỏe chung: "&IFERROR(LEFT(IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6),FIND(" (",IFERROR(LEFT
    #   (SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6))-1),IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6))&" ","")&I
    #   F(OR(LEFT(IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7),5)="THUẬN",LEFT(IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!
    #   $J$7)-1),SUC_KHOE!$J$7),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7),3)="TỐT",IFERROR(LEFT(SUC_K
    #   HOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC
    #   _KHOE!$J$7))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(L
    #   EFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1)
    #   ,SUC_KHOE!$J$7)))),"• Bệnh tật: "&IFERROR(LEFT(IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7),FIND(" (",IFERROR(LEFT(SUC_KH
    #   OE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7))-1),IFERROR(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7)-1),SUC_KHOE!$J$7))&" ","")&IF(OR(LE
    #   FT(IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8),5)="THUẬN",LEFT(IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1
    #   ),SUC_KHOE!$J$8),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8),3)="TỐT",IFERROR(LEFT(SUC_KHOE!$J$
    #   8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$
    #   J$8))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(LEFT(SUC
    #   _KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KH
    #   OE!$J$8)))),"• Tai họa theo vận: "&IFERROR(LEFT(IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8),FIND(" (",IFERROR(LEFT(SUC_K
    #   HOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8))-1),IFERROR(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8)-1),SUC_KHOE!$J$8))&" ","")&IF(OR(L
    #   EFT(MENH_CUC!$D$21,5)="THUẬN",LEFT(MENH_CUC!$D$21,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$21,3)="TỐT",MENH_CUC!$D$21="Khá",ISNUMBER(SEARCH("VỀ GIÀU",MENH
    #   _CUC!$D$21)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$21))),"• Thọ – 
    #   yểu: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$21,FIND(" (",MENH_CUC!$D$21)-1),MENH_CUC!$D$21))&" ","")&KET_LUAN!$C$157&" "&KET_LUAN!$C$132&" "
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• Sức khỏe chung: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))))) else ''),(X.cat(X.cat('• Bệnh tật: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0)), lambda: E.v(22,7,10)))))) else '')),(X.cat(X.cat('• Tai họa theo vận: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0)), lambda: E.v(22,8,10)))))) else '')),(X.cat(X.cat('• Thọ – yểu: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,21,4), X.sub(X.FIND(' (', E.v(18,21,4)),1.0)), lambda: E.v(18,21,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,21,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,21,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,21,4), 3.0),'TỐT'), X.eq(E.v(18,21,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,21,4))))) else '')),E.v(3,157,3)),' '),E.v(3,132,3)),' ')

def f54(E, r, c):
    # Ô C58
    # = "Sức khỏe chung: "&SUBSTITUTE(IFERROR(LEFT(KET_LUAN!$C$137,MIN(260,FIND("Vì sao",KET_LUAN!$C$137)-1)),IF(LEN(KET_LUAN!$C$137)>260,LEFT(KET_LUAN!$C$137
    #   ,260)&"…",KET_LUAN!$C$137)),CHAR(10)," ")&CHAR(10)&"Thọ – yểu: "&IF(LEN(KET_LUAN!$C$143)>200,LEFT(KET_LUAN!$C$143,200)&"…",KET_LUAN!$C$143)
    return X.cat(X.cat(X.cat(X.cat('Sức khỏe chung: ',X.SUBSTITUTE(X.IFERROR(lambda: X.LEFT(E.v(3,137,3), X.MIN(260.0, X.sub(X.FIND('Vì sao', E.v(3,137,3)),1.0))), lambda: (X.cat(X.LEFT(E.v(3,137,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,137,3)),260.0)) else E.v(3,137,3))), X.CHAR(10.0), ' ')),X.CHAR(10.0)),'Thọ – yểu: '),(X.cat(X.LEFT(E.v(3,143,3), 200.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,143,3)),200.0)) else E.v(3,143,3)))

def f55(E, r, c):
    # Ô C59
    # = "Bệnh tật: "&SUBSTITUTE(IFERROR(LEFT(KET_LUAN!$C$138,MIN(200,FIND("Vì sao",KET_LUAN!$C$138)-1)),IF(LEN(KET_LUAN!$C$138)>200,LEFT(KET_LUAN!$C$138,200)&
    #   "…",KET_LUAN!$C$138)),CHAR(10)," ")&CHAR(10)&IFERROR(MID(KET_LUAN!$C$138,FIND("Tạng phủ cần chú ý",KET_LUAN!$C$138),500),"")
    return X.cat(X.cat(X.cat('Bệnh tật: ',X.SUBSTITUTE(X.IFERROR(lambda: X.LEFT(E.v(3,138,3), X.MIN(200.0, X.sub(X.FIND('Vì sao', E.v(3,138,3)),1.0))), lambda: (X.cat(X.LEFT(E.v(3,138,3), 200.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,138,3)),200.0)) else E.v(3,138,3))), X.CHAR(10.0), ' ')),X.CHAR(10.0)),X.IFERROR(lambda: X.MID(E.v(3,138,3), X.FIND('Tạng phủ cần chú ý', E.v(3,138,3)), 500.0), lambda: ''))

def f56(E, r, c):
    # Ô C60
    # = "Tai họa theo vận: "&IFERROR(MID(KET_LUAN!$C$139,FIND("Đại vận cần đề phòng",KET_LUAN!$C$139),700),IF(LEN(KET_LUAN!$C$139)>500,LEFT(KET_LUAN!$C$139,50
    #   0)&"…",KET_LUAN!$C$139))
    return X.cat('Tai họa theo vận: ',X.IFERROR(lambda: X.MID(E.v(3,139,3), X.FIND('Đại vận cần đề phòng', E.v(3,139,3)), 700.0), lambda: (X.cat(X.LEFT(E.v(3,139,3), 500.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,139,3)),500.0)) else E.v(3,139,3))))

def f57(E, r, c):
    # Ô C61
    # = "Năm cần đề phòng (năm xem + 5): "&KET_LUAN!$C$140
    return X.cat('Năm cần đề phòng (năm xem + 5): ',E.v(3,140,3))

def f58(E, r, c):
    # Ô C62
    # = "Thói quen nên giữ – hành kỵ thần "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1)&" (thái quá): "&CHOOSE(IF(DOI_CHIEU!$C
    #   $33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1,"gan – mật, gân cơ, mắt: ngủ trước 23 giờ, hạn chế rượu bia, tập giãn cơ","tim – mạch, huyết áp, giấc ngủ
    #   : kiểm soát căng thẳng, đo huyết áp định kỳ, hạn chế cà phê và đồ cay nóng","dạ dày – tì vị, tiêu hóa: ăn đúng bữa, nhai kỹ, hạn chế đồ ngọt béo","phổ
    #   i – họng, đại tràng, da: tránh khói thuốc, giữ ấm cổ họng, ăn nhiều chất xơ","thận – bàng quang, xương khớp, tai: uống đủ nước, giữ ấm lưng – chân, kh
    #   ông nhịn tiểu")&". Hành bị kỵ thần khắc "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+2,5)+1)&": "&CHOOSE(MOD(IF(DOI_
    #   CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+2,5)+1,"gan – mật, gân cơ, mắt: ngủ trước 23 giờ, hạn chế rượu bia, tập giãn cơ","tim – mạch, huyết
    #    áp, giấc ngủ: kiểm soát căng thẳng, đo huyết áp định kỳ, hạn chế cà phê và đồ cay nóng","dạ dày – tì vị, tiêu hóa: ăn đúng bữa, nhai kỹ, hạn chế đồ n
    #   gọt béo","phổi – họng, đại tràng, da: tránh khói thuốc, giữ ấm cổ họng, ăn nhiều chất xơ","thận – bàng quang, xương khớp, tai: uống đủ nước, giữ ấm lư
    #   ng – chân, không nhịn tiểu")&"."&IF(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1)," Sinh mùa đông: giữ ấm, vận động ngoài trời, tránh đồ 
    #   lạnh.",IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7)," Sinh mùa hè: giải nhiệt, uống đủ nước, tránh nóng nảy và thức khuya.",""))&" Khô
    #   ng thay cho khám bệnh."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thói quen nên giữ – hành kỵ thần ',X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0))),' (thái quá): '),X.CHOOSE(X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0), lambda: 'gan – mật, gân cơ, mắt: ngủ trước 23 giờ, hạn chế rượu bia, tập giãn cơ', lambda: 'tim – mạch, huyết áp, giấc ngủ: kiểm soát căng thẳng, đo huyết áp định kỳ, hạn chế cà phê và đồ cay nóng', lambda: 'dạ dày – tì vị, tiêu hóa: ăn đúng bữa, nhai kỹ, hạn chế đồ ngọt béo', lambda: 'phổi – họng, đại tràng, da: tránh khói thuốc, giữ ấm cổ họng, ăn nhiều chất xơ', lambda: 'thận – bàng quang, xương khớp, tai: uống đủ nước, giữ ấm lưng – chân, không nhịn tiểu')),'. Hành bị kỵ thần khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),2.0), 5.0),1.0))),': '),X.CHOOSE(X.add(X.MOD(X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),2.0), 5.0),1.0), lambda: 'gan – mật, gân cơ, mắt: ngủ trước 23 giờ, hạn chế rượu bia, tập giãn cơ', lambda: 'tim – mạch, huyết áp, giấc ngủ: kiểm soát căng thẳng, đo huyết áp định kỳ, hạn chế cà phê và đồ cay nóng', lambda: 'dạ dày – tì vị, tiêu hóa: ăn đúng bữa, nhai kỹ, hạn chế đồ ngọt béo', lambda: 'phổi – họng, đại tràng, da: tránh khói thuốc, giữ ấm cổ họng, ăn nhiều chất xơ', lambda: 'thận – bàng quang, xương khớp, tai: uống đủ nước, giữ ấm lưng – chân, không nhịn tiểu')),'.'),(' Sinh mùa đông: giữ ấm, vận động ngoài trời, tránh đồ lạnh.' if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else (' Sinh mùa hè: giải nhiệt, uống đủ nước, tránh nóng nảy và thức khuya.' if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else ''))),' Không thay cho khám bệnh.')

def f59(E, r, c):
    # Ô C63
    # = "MỐC SỨC KHỎE – năm tạng phủ mức NẶNG (SUC_KHOE mục H) – đã qua: "&IF(THEO_DOI!$AW$111="","không có",THEO_DOI!$AW$111)&CHAR(10)&"   Sắp tới (20 năm): 
    #   "&IF(THEO_DOI!$AX$111="","không có",THEO_DOI!$AX$111)
    return X.cat(X.cat(X.cat(X.cat('MỐC SỨC KHỎE – năm tạng phủ mức NẶNG (SUC_KHOE mục H) – đã qua: ',('không có' if X.truthy(X.eq(E.v(5,111,49),'')) else E.v(5,111,49))),X.CHAR(10.0)),'   Sắp tới (20 năm): '),('không có' if X.truthy(X.eq(E.v(5,111,50),'')) else E.v(5,111,50)))

def f60(E, r, c):
    # Ô C64
    # = "Cấu trúc dễ sinh ung tượng (KET_LUAN mục H): "&IF(LEN(KET_LUAN!$C$155)>420,LEFT(KET_LUAN!$C$155,420)&"…",KET_LUAN!$C$155)&CHAR(10)&"Năm kích hoạt: "&
    #   IF(LEN(KET_LUAN!$C$154)>200,LEFT(KET_LUAN!$C$154,200)&"…",KET_LUAN!$C$154)
    return X.cat(X.cat(X.cat(X.cat('Cấu trúc dễ sinh ung tượng (KET_LUAN mục H): ',(X.cat(X.LEFT(E.v(3,155,3), 420.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,155,3)),420.0)) else E.v(3,155,3))),X.CHAR(10.0)),'Năm kích hoạt: '),(X.cat(X.LEFT(E.v(3,154,3), 200.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,154,3)),200.0)) else E.v(3,154,3)))

def f61(E, r, c):
    # Ô C66
    # = "• Đại vận đang hành: "&IF(LEN(KET_LUAN!$C$163)>260,LEFT(KET_LUAN!$C$163,260)&"…",KET_LUAN!$C$163)&CHAR(10)&"• Đại vận thuận (dụng/hỷ): "&IF(LUC_THAN!
    #   $Y$31="","(không có)",LUC_THAN!$Y$31)&CHAR(10)&"• Đại vận cần thận trọng: "&IF(LUC_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31)&CHAR(10)&"• Vận đang hàn
    #   h theo can – chi (Hóa giải toàn thư): "&IF(DAI_VAN!$C$25<1,"chưa vào vận",INDEX(DAI_VAN!$C$102:$L$102,MIN(10,DAI_VAN!$C$25)))&CHAR(10)&"• Năm thuận (n
    #   ăm xem → +10): "&IF(TRIM(IF($L$71=1,$A$71&" ","")&IF($L$72=1,$A$72&" ","")&IF($L$73=1,$A$73&" ","")&IF($L$74=1,$A$74&" ","")&IF($L$75=1,$A$75&" ","")&
    #   IF($L$76=1,$A$76&" ","")&IF($L$77=1,$A$77&" ","")&IF($L$78=1,$A$78&" ","")&IF($L$79=1,$A$79&" ","")&IF($L$80=1,$A$80&" ","")&IF($L$81=1,$A$81&" ",""))
    #   ="","không có năm tốt rõ",TRIM(IF($L$71=1,$A$71&" ","")&IF($L$72=1,$A$72&" ","")&IF($L$73=1,$A$73&" ","")&IF($L$74=1,$A$74&" ","")&IF($L$75=1,$A$75&" 
    #   ","")&IF($L$76=1,$A$76&" ","")&IF($L$77=1,$A$77&" ","")&IF($L$78=1,$A$78&" ","")&IF($L$79=1,$A$79&" ","")&IF($L$80=1,$A$80&" ","")&IF($L$81=1,$A$81&" 
    #   ","")))&CHAR(10)&"• Năm cần thận trọng: "&IF(TRIM(IF($L$71=-1,$A$71&" ","")&IF($L$72=-1,$A$72&" ","")&IF($L$73=-1,$A$73&" ","")&IF($L$74=-1,$A$74&" ",
    #   "")&IF($L$75=-1,$A$75&" ","")&IF($L$76=-1,$A$76&" ","")&IF($L$77=-1,$A$77&" ","")&IF($L$78=-1,$A$78&" ","")&IF($L$79=-1,$A$79&" ","")&IF($L$80=-1,$A$8
    #   0&" ","")&IF($L$81=-1,$A$81&" ",""))="","không có năm xấu rõ",TRIM(IF($L$71=-1,$A$71&" ","")&IF($L$72=-1,$A$72&" ","")&IF($L$73=-1,$A$73&" ","")&IF($L
    #   $74=-1,$A$74&" ","")&IF($L$75=-1,$A$75&" ","")&IF($L$76=-1,$A$76&" ","")&IF($L$77=-1,$A$77&" ","")&IF($L$78=-1,$A$78&" ","")&IF($L$79=-1,$A$79&" ","")
    #   &IF($L$80=-1,$A$80&" ","")&IF($L$81=-1,$A$81&" ","")))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Đại vận đang hành: ',(X.cat(X.LEFT(E.v(3,163,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,163,3)),260.0)) else E.v(3,163,3))),X.CHAR(10.0)),'• Đại vận thuận (dụng/hỷ): '),('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),X.CHAR(10.0)),'• Đại vận cần thận trọng: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))),X.CHAR(10.0)),'• Vận đang hành theo can – chi (Hóa giải toàn thư): '),('chưa vào vận' if X.truthy(X.lt(E.v(14,25,3),1.0)) else X.INDEX(E.rg(14,102,3,102,12), X.MIN(10.0, E.rg(14,25,3,25,3))))),X.CHAR(10.0)),'• Năm thuận (năm xem → +10): '),('không có năm tốt rõ' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),1.0)) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),1.0)) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),1.0)) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),1.0)) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),1.0)) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),1.0)) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),1.0)) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),1.0)) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),1.0)) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),1.0)) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),1.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),1.0)) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),1.0)) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),1.0)) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),1.0)) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),1.0)) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),1.0)) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),1.0)) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),1.0)) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),1.0)) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),1.0)) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),1.0)) else ''))))),X.CHAR(10.0)),'• Năm cần thận trọng: '),('không có năm xấu rõ' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),X.neg(1.0))) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),X.neg(1.0))) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),X.neg(1.0))) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),X.neg(1.0))) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),X.neg(1.0))) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),X.neg(1.0))) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),X.neg(1.0))) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),X.neg(1.0))) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),X.neg(1.0))) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),X.neg(1.0))) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),X.neg(1.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),X.neg(1.0))) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),X.neg(1.0))) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),X.neg(1.0))) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),X.neg(1.0))) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),X.neg(1.0))) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),X.neg(1.0))) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),X.neg(1.0))) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),X.neg(1.0))) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),X.neg(1.0))) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),X.neg(1.0))) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),X.neg(1.0))) else '')))))

def f62(E, r, c):
    # Ô J67
    # = IF(LUC_THAN!$Y$31="","","• Đại vận thuận: "&LUC_THAN!$Y$31&" ")&IF(TRIM(IF($L$71=1,$A$71&" ","")&IF($L$72=1,$A$72&" ","")&IF($L$73=1,$A$73&" ","")&IF(
    #   $L$74=1,$A$74&" ","")&IF($L$75=1,$A$75&" ","")&IF($L$76=1,$A$76&" ","")&IF($L$77=1,$A$77&" ","")&IF($L$78=1,$A$78&" ","")&IF($L$79=1,$A$79&" ","")&IF(
    #   $L$80=1,$A$80&" ","")&IF($L$81=1,$A$81&" ",""))="","","• Năm thuận: "&TRIM(IF($L$71=1,$A$71&" ","")&IF($L$72=1,$A$72&" ","")&IF($L$73=1,$A$73&" ","")&
    #   IF($L$74=1,$A$74&" ","")&IF($L$75=1,$A$75&" ","")&IF($L$76=1,$A$76&" ","")&IF($L$77=1,$A$77&" ","")&IF($L$78=1,$A$78&" ","")&IF($L$79=1,$A$79&" ","")&
    #   IF($L$80=1,$A$80&" ","")&IF($L$81=1,$A$81&" ",""))&" ")&IF(DAI_VAN!$C$25<1,"",IF(ISNUMBER(SEARCH("cát",INDEX(DAI_VAN!$C$102:$L$102,MIN(10,DAI_VAN!$C$2
    #   5)))),"• Vận đang hành theo can – chi: "&INDEX(DAI_VAN!$C$102:$L$102,MIN(10,DAI_VAN!$C$25))&" ",""))
    return X.cat(X.cat(('' if X.truthy(X.eq(E.v(16,31,25),'')) else X.cat(X.cat('• Đại vận thuận: ',E.v(16,31,25)),' ')),('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),1.0)) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),1.0)) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),1.0)) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),1.0)) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),1.0)) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),1.0)) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),1.0)) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),1.0)) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),1.0)) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),1.0)) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),1.0)) else ''))),'')) else X.cat(X.cat('• Năm thuận: ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),1.0)) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),1.0)) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),1.0)) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),1.0)) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),1.0)) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),1.0)) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),1.0)) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),1.0)) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),1.0)) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),1.0)) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),1.0)) else '')))),' '))),('' if X.truthy(X.lt(E.v(14,25,3),1.0)) else (X.cat(X.cat('• Vận đang hành theo can – chi: ',X.INDEX(E.rg(14,102,3,102,12), X.MIN(10.0, E.rg(14,25,3,25,3)))),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('cát', X.INDEX(E.rg(14,102,3,102,12), X.MIN(10.0, E.rg(14,25,3,25,3)))))) else '')))

def f63(E, r, c):
    # Ô C68
    # = "Khởi vận: "&IF(LEN(KET_LUAN!$C$162)>300,LEFT(KET_LUAN!$C$162,300)&"…",KET_LUAN!$C$162)&CHAR(10)&"Đại vận đang hành: "&IF(LEN(KET_LUAN!$C$163)>450,LEF
    #   T(KET_LUAN!$C$163,450)&"…",KET_LUAN!$C$163)
    return X.cat(X.cat(X.cat(X.cat('Khởi vận: ',(X.cat(X.LEFT(E.v(3,162,3), 300.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,162,3)),300.0)) else E.v(3,162,3))),X.CHAR(10.0)),'Đại vận đang hành: '),(X.cat(X.LEFT(E.v(3,163,3), 450.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,163,3)),450.0)) else E.v(3,163,3)))

def f64(E, r, c):
    # Ô C69
    # = "Đại vận tốt (dụng/hỷ): "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; đại vận xấu (kỵ): "&IF(LUC_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31)&C
    #   HAR(10)&IF(LEN(KET_LUAN!$C$171)>500,LEFT(KET_LUAN!$C$171,500)&"…",KET_LUAN!$C$171)&IF(KET_LUAN!$C$174="","",CHAR(10)&"Bổ sung (Uyên Hải Tử Bình): "&KE
    #   T_LUAN!$C$174)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận tốt (dụng/hỷ): ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; đại vận xấu (kỵ): '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))),X.CHAR(10.0)),(X.cat(X.LEFT(E.v(3,171,3), 500.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,171,3)),500.0)) else E.v(3,171,3))),('' if X.truthy(X.eq(E.v(3,174,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Bổ sung (Uyên Hải Tử Bình): '),E.v(3,174,3))))

def f65(E, r, c):
    # Ô A71
    # = (N(NHAP_LIEU!$C$31)+0)
    return X.add(X.N(E.v(2,31,3)),0.0)

def f66(E, r, c):
    # Ô B71, B72, B73, B74, B75, B76, B77, B78, B79, B80, B81
    # = IF($K71="","",INDEX(LUU_NIEN!$L$30:$L$129,$K71)&CHAR(10)&"ĐV: "&INDEX(LUU_NIEN!$H$30:$H$129,$K71))
    return ('' if X.truthy(X.eq(E.v(4,r,11),'')) else X.cat(X.cat(X.cat(X.INDEX(E.rg(15,30,12,129,12), E.v(4,r,11)),X.CHAR(10.0)),'ĐV: '),X.INDEX(E.rg(15,30,8,129,8), E.v(4,r,11))))

def f67(E, r, c):
    # Ô C71, C72, C73, C74, C75, C76, C77, C78, C79, C80, C81
    # = IF($K71="","",IF(DOI_CHIEU!$C$33=2,INDEX(LUU_NIEN!$AL$30:$AL$129,$K71),INDEX(LUU_NIEN!$AA$30:$AA$129,$K71))&IF(INDEX(LUU_NIEN!$BH$30:$BH$129,$K71)=1,C
    #   HAR(10)&"Dấu hiệu bất lợi cho cha (P2 tr.440)","")&IF(INDEX(LUU_NIEN!$BI$30:$BI$129,$K71)=1,CHAR(10)&"Dấu hiệu bất lợi cho mẹ (P2 tr.440)","")&IF($K71
    #   ="","",CHAR(10)&INDEX(LUU_NIEN!$CV$30:$CV$129,$K71)))
    return ('' if X.truthy(X.eq(E.v(4,r,11),'')) else X.cat(X.cat(X.cat((X.INDEX(E.rg(15,30,38,129,38), E.v(4,r,11)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(15,30,27,129,27), E.v(4,r,11))),(X.cat(X.CHAR(10.0),'Dấu hiệu bất lợi cho cha (P2 tr.440)') if X.truthy(X.eq(X.INDEX(E.rg(15,30,60,129,60), E.v(4,r,11)),1.0)) else '')),(X.cat(X.CHAR(10.0),'Dấu hiệu bất lợi cho mẹ (P2 tr.440)') if X.truthy(X.eq(X.INDEX(E.rg(15,30,61,129,61), E.v(4,r,11)),1.0)) else '')),('' if X.truthy(X.eq(E.v(4,r,11),'')) else X.cat(X.CHAR(10.0),X.INDEX(E.rg(15,30,100,129,100), E.v(4,r,11))))))

def f68(E, r, c):
    # Ô D71, D72, D73, D74, D75, D76, D77, D78, D79, D80, D81
    # = IF($K71="","","Can năm "&INDEX(LUU_NIEN!$O$30:$O$129,$K71)&" – "&IF(INDEX(LUU_NIEN!$P$30:$P$129,$K71)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_K
    #   Y!$C$109),"dụng thần (thuận)",IF(INDEX(LUU_NIEN!$P$30:$P$129,$K71)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần (thuận)",IF(INDE
    #   X(LUU_NIEN!$P$30:$P$129,$K71)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần (bất lợi)","nhàn thần")))&": "&CHOOSE(MOD(INDEX(LUU_N
    #   IEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè, đồng nghiệp; cạnh tranh, chi tiêu","tài năng, ý tưởng, lời nói, con cái (nữ); dễ va chạ
    #   m cấp trên","tiền bạc, tài sản, vợ (nam), cha","công việc, chức vụ, áp lực, chồng (nữ), con (nam)","học tập, giấy tờ, nhà cửa, mẹ, người nâng đỡ")&"."
    #   &IF(AND(MOD(INDEX(LUU_NIEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$E$7+5,5)=3,OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$
    #   109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))," → Năm quan tinh là dụng/hỷ: thuận công danh, dễ được giao 
    #   việc, bổ nhiệm (P1 tr.167-169).",IF(AND(MOD(INDEX(LUU_NIEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$E$7+5,5)=1,OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,
    #   DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))," → Thực thương gặp quan (qu
    #   an là dụng/hỷ): dễ biến động chức vụ, va chạm cấp trên (P2 tr.541-542).",IF(AND(MOD(INDEX(LUU_NIEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$E$7+5,5)=0,OR(MOD(TR
    #   UNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_K
    #   Y!$C$110)))," → Tỉ kiếp tranh tài: giữ tiền, tránh cho vay, hùn vốn (P1 tr.156; P2 tr.470).",IF(AND(MOD(INDEX(LUU_NIEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$
    #   E$7+5,5)=1,OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$110,DUNG_HY_KY!$C$110)))," → Thực thương sinh tài: thu nhập khá (P2 tr.521-522).",IF(AND(MOD(INDEX(LUU_NIEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$E$7+
    #   5,5)=2,OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!
    #   $D$110,DUNG_HY_KY!$C$110)))," → Năm tài là dụng/hỷ: dễ có thêm thu nhập (P1 tr.164-165).",IF(AND(MOD(INDEX(LUU_NIEN!$P$30:$P$129,$K71)-TRUNG_GIAN!$E$7
    #   +5,5)=4,OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY
    #   !$D$110,DUNG_HY_KY!$C$110)))," → Năm ấn là dụng/hỷ: thuận học hành, giấy tờ, có quý nhân.","")))))))
    return ('' if X.truthy(X.eq(E.v(4,r,11),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can năm ',X.INDEX(E.rg(15,30,15,129,15), E.v(4,r,11))),' – '),('dụng thần (thuận)' if X.truthy(X.eq(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần (thuận)' if X.truthy(X.eq(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần (bất lợi)' if X.truthy(X.eq(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'nhàn thần')))),': '),X.CHOOSE(X.add(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè, đồng nghiệp; cạnh tranh, chi tiêu', lambda: 'tài năng, ý tưởng, lời nói, con cái (nữ); dễ va chạm cấp trên', lambda: 'tiền bạc, tài sản, vợ (nam), cha', lambda: 'công việc, chức vụ, áp lực, chồng (nữ), con (nam)', lambda: 'học tập, giấy tờ, nhà cửa, mẹ, người nâng đỡ')),'.'),(' → Năm quan tinh là dụng/hỷ: thuận công danh, dễ được giao việc, bổ nhiệm (P1 tr.167-169).' if X.truthy(X.AND(X.eq(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),3.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (' → Thực thương gặp quan (quan là dụng/hỷ): dễ biến động chức vụ, va chạm cấp trên (P2 tr.541-542).' if X.truthy(X.AND(X.eq(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (' → Tỉ kiếp tranh tài: giữ tiền, tránh cho vay, hùn vốn (P1 tr.156; P2 tr.470).' if X.truthy(X.AND(X.eq(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),0.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (' → Thực thương sinh tài: thu nhập khá (P2 tr.521-522).' if X.truthy(X.AND(X.eq(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (' → Năm tài là dụng/hỷ: dễ có thêm thu nhập (P1 tr.164-165).' if X.truthy(X.AND(X.eq(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),2.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (' → Năm ấn là dụng/hỷ: thuận học hành, giấy tờ, có quý nhân.' if X.truthy(X.AND(X.eq(X.MOD(X.add(X.sub(X.INDEX(E.rg(15,30,16,129,16), E.v(4,r,11)),E.v(23,7,5)),5.0), 5.0),4.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else ''))))))))

def f69(E, r, c):
    # Ô E71, E72, E73, E74, E75, E76, E77, E78, E79, E80, E81
    # = IF($K71="","",IF(INDEX(LUU_NIEN!$BN$30:$BN$129,$K71)="","—",IF(LEN(INDEX(LUU_NIEN!$BN$30:$BN$129,$K71))>420,LEFT(INDEX(LUU_NIEN!$BN$30:$BN$129,$K71),4
    #   20)&"…",INDEX(LUU_NIEN!$BN$30:$BN$129,$K71))))
    return ('' if X.truthy(X.eq(E.v(4,r,11),'')) else ('—' if X.truthy(X.eq(X.INDEX(E.rg(15,30,66,129,66), E.v(4,r,11)),'')) else (X.cat(X.LEFT(X.INDEX(E.rg(15,30,66,129,66), E.v(4,r,11)), 420.0),'…') if X.truthy(X.gt(X.LEN(X.INDEX(E.rg(15,30,66,129,66), E.v(4,r,11))),420.0)) else X.INDEX(E.rg(15,30,66,129,66), E.v(4,r,11)))))

def f70(E, r, c):
    # Ô F71, F72, F73, F74, F75, F76, F77, F78, F79, F80, F81
    # = IF($K71="","",IF(INDEX(LUU_NIEN!$CN$30:$CN$129,$K71)="","—",IFERROR(LEFT(INDEX(LUU_NIEN!$CN$30:$CN$129,$K71),FIND(":",INDEX(LUU_NIEN!$CN$30:$CN$129,$K
    #   71))-1),IF(LEN(INDEX(LUU_NIEN!$CN$30:$CN$129,$K71))>120,LEFT(INDEX(LUU_NIEN!$CN$30:$CN$129,$K71),120)&"…",INDEX(LUU_NIEN!$CN$30:$CN$129,$K71)))))
    return ('' if X.truthy(X.eq(E.v(4,r,11),'')) else ('—' if X.truthy(X.eq(X.INDEX(E.rg(15,30,92,129,92), E.v(4,r,11)),'')) else X.IFERROR(lambda: X.LEFT(X.INDEX(E.rg(15,30,92,129,92), E.v(4,r,11)), X.sub(X.FIND(':', X.INDEX(E.rg(15,30,92,129,92), E.v(4,r,11))),1.0)), lambda: (X.cat(X.LEFT(X.INDEX(E.rg(15,30,92,129,92), E.v(4,r,11)), 120.0),'…') if X.truthy(X.gt(X.LEN(X.INDEX(E.rg(15,30,92,129,92), E.v(4,r,11))),120.0)) else X.INDEX(E.rg(15,30,92,129,92), E.v(4,r,11))))))

def f71(E, r, c):
    # Ô G71, G72, G73, G74, G75, G76, G77, G78, G79, G80, G81
    # = IFERROR(INDEX(THEO_DOI!$O$12:$O$111,MATCH($A71,THEO_DOI!$A$12:$A$111,0))&"","")
    return X.IFERROR(lambda: X.cat(X.INDEX(E.rg(5,12,15,111,15), X.MATCH(E.v(4,r,1), E.rg(5,12,1,111,1), 0.0)),''), lambda: '')

def f72(E, r, c):
    # Ô H71, H72, H73, H74, H75, H76, H77, H78, H79, H80, H81
    # = IFERROR(INDEX(THEO_DOI!$P$12:$P$111,MATCH($A71,THEO_DOI!$A$12:$A$111,0))&"","")
    return X.IFERROR(lambda: X.cat(X.INDEX(E.rg(5,12,16,111,16), X.MATCH(E.v(4,r,1), E.rg(5,12,1,111,1), 0.0)),''), lambda: '')

def f73(E, r, c):
    # Ô I71, I72, I73, I74, I75, I76, I77, I78, I79, I80, I81
    # = IFERROR(INDEX(THEO_DOI!$N$12:$N$111,MATCH($A71,THEO_DOI!$A$12:$A$111,0))&"","")
    return X.IFERROR(lambda: X.cat(X.INDEX(E.rg(5,12,14,111,14), X.MATCH(E.v(4,r,1), E.rg(5,12,1,111,1), 0.0)),''), lambda: '')

def f74(E, r, c):
    # Ô K71
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+0),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),0.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f75(E, r, c):
    # Ô L71, L72, L73, L74, L75, L76, L77, L78, L79, L80, L81
    # = IF(OR(ISNUMBER(SEARCH("Rất tốt",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C71))),ISNUMBER(SEARCH("Tốt vừa",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C71))),
    #   ISNUMBER(SEARCH("Tốt nhiều xấu ít",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C71))),AND(ISNUMBER(SEARCH("Tốt ít",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C
    #   71))),NOT(ISNUMBER(SEARCH("Tốt ít xấu nhiều",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C71)))))),1,IF(OR(ISNUMBER(SEARCH("Xấu vừa",IFERROR(LEFT(C71,FIND(
    #   CHAR(10),C71)-1),C71))),ISNUMBER(SEARCH("Xấu nhiều",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C71))),ISNUMBER(SEARCH("Tốt ít xấu nhiều",IFERROR(LEFT(C71,
    #   FIND(CHAR(10),C71)-1),C71))),AND(ISNUMBER(SEARCH("Xấu ít",IFERROR(LEFT(C71,FIND(CHAR(10),C71)-1),C71))),NOT(ISNUMBER(SEARCH("Tốt nhiều xấu ít",IFERROR
    #   (LEFT(C71,FIND(CHAR(10),C71)-1),C71)))))),-1,0))
    return (1.0 if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Rất tốt', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.ISNUMBER(lambda: X.SEARCH('Tốt vừa', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.ISNUMBER(lambda: X.SEARCH('Tốt nhiều xấu ít', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.AND(X.ISNUMBER(lambda: X.SEARCH('Tốt ít', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.NOT(X.ISNUMBER(lambda: X.SEARCH('Tốt ít xấu nhiều', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))))))) else (X.neg(1.0) if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Xấu vừa', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.ISNUMBER(lambda: X.SEARCH('Xấu nhiều', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.ISNUMBER(lambda: X.SEARCH('Tốt ít xấu nhiều', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.AND(X.ISNUMBER(lambda: X.SEARCH('Xấu ít', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))), X.NOT(X.ISNUMBER(lambda: X.SEARCH('Tốt nhiều xấu ít', X.IFERROR(lambda: X.LEFT(E.v(4,r,c-9), X.sub(X.FIND(X.CHAR(10.0), E.v(4,r,c-9)),1.0)), lambda: E.v(4,r,c-9)))))))) else 0.0))

def f76(E, r, c):
    # Ô A72
    # = (N(NHAP_LIEU!$C$31)+1)
    return X.add(X.N(E.v(2,31,3)),1.0)

def f77(E, r, c):
    # Ô K72
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+1),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),1.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f78(E, r, c):
    # Ô A73
    # = (N(NHAP_LIEU!$C$31)+2)
    return X.add(X.N(E.v(2,31,3)),2.0)

def f79(E, r, c):
    # Ô K73
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+2),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),2.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f80(E, r, c):
    # Ô A74
    # = (N(NHAP_LIEU!$C$31)+3)
    return X.add(X.N(E.v(2,31,3)),3.0)

def f81(E, r, c):
    # Ô K74
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+3),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),3.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f82(E, r, c):
    # Ô A75
    # = (N(NHAP_LIEU!$C$31)+4)
    return X.add(X.N(E.v(2,31,3)),4.0)

def f83(E, r, c):
    # Ô K75
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+4),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),4.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f84(E, r, c):
    # Ô A76
    # = (N(NHAP_LIEU!$C$31)+5)
    return X.add(X.N(E.v(2,31,3)),5.0)

def f85(E, r, c):
    # Ô K76
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+5),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),5.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f86(E, r, c):
    # Ô A77
    # = (N(NHAP_LIEU!$C$31)+6)
    return X.add(X.N(E.v(2,31,3)),6.0)

def f87(E, r, c):
    # Ô K77
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+6),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),6.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f88(E, r, c):
    # Ô A78
    # = (N(NHAP_LIEU!$C$31)+7)
    return X.add(X.N(E.v(2,31,3)),7.0)

def f89(E, r, c):
    # Ô K78
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+7),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),7.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f90(E, r, c):
    # Ô A79
    # = (N(NHAP_LIEU!$C$31)+8)
    return X.add(X.N(E.v(2,31,3)),8.0)

def f91(E, r, c):
    # Ô K79
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+8),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),8.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f92(E, r, c):
    # Ô A80
    # = (N(NHAP_LIEU!$C$31)+9)
    return X.add(X.N(E.v(2,31,3)),9.0)

def f93(E, r, c):
    # Ô K80
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+9),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),9.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f94(E, r, c):
    # Ô A81
    # = (N(NHAP_LIEU!$C$31)+10)
    return X.add(X.N(E.v(2,31,3)),10.0)

def f95(E, r, c):
    # Ô K81
    # = IFERROR(MATCH((N(NHAP_LIEU!$C$31)+10),LUU_NIEN!$A$30:$A$129,0),"")
    return X.IFERROR(lambda: X.MATCH(X.add(X.N(E.v(2,31,3)),10.0), E.rg(15,30,1,129,1), 0.0), lambda: '')

def f96(E, r, c):
    # Ô C83
    # = "• Nguyên tắc chung: "&LUAN_GIAI!$H$106&CHAR(10)&"• Nên tăng cường hành "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&" và "&IF(DOI_CHIEU!$C$3
    #   3=2,KET_LUAN!$D$37,KET_LUAN!$C$37)&"; hạn chế hành "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)&"."&CHAR(10)&"• Năm xem: "&IF(LEN(KET_LUAN!$C
    #   $141)>260,LEFT(KET_LUAN!$C$141,260)&"…",KET_LUAN!$C$141)&CHAR(10)&"• Hóa giải theo phương vị, tên gọi, phong thủy, sinh hoạt: xem SUC_KHOE mục I và KE
    #   T_LUAN mục I."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Nguyên tắc chung: ',E.v(4,106,8)),X.CHAR(10.0)),'• Nên tăng cường hành '),(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),' và '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'; hạn chế hành '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3))),'.'),X.CHAR(10.0)),'• Năm xem: '),(X.cat(X.LEFT(E.v(3,141,3), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,141,3)),260.0)) else E.v(3,141,3))),X.CHAR(10.0)),'• Hóa giải theo phương vị, tên gọi, phong thủy, sinh hoạt: xem SUC_KHOE mục I và KET_LUAN mục I.')

def f97(E, r, c):
    # Ô J84
    # = IF(TRIM(IF($L$71=1,$A$71&" ","")&IF($L$72=1,$A$72&" ","")&IF($L$73=1,$A$73&" ","")&IF($L$74=1,$A$74&" ","")&IF($L$75=1,$A$75&" ","")&IF($L$76=1,$A$76&
    #   " ","")&IF($L$77=1,$A$77&" ","")&IF($L$78=1,$A$78&" ","")&IF($L$79=1,$A$79&" ","")&IF($L$80=1,$A$80&" ","")&IF($L$81=1,$A$81&" ",""))="","","• Thời đi
    #   ểm nên hành động: "&TRIM(IF($L$71=1,$A$71&" ","")&IF($L$72=1,$A$72&" ","")&IF($L$73=1,$A$73&" ","")&IF($L$74=1,$A$74&" ","")&IF($L$75=1,$A$75&" ","")&
    #   IF($L$76=1,$A$76&" ","")&IF($L$77=1,$A$77&" ","")&IF($L$78=1,$A$78&" ","")&IF($L$79=1,$A$79&" ","")&IF($L$80=1,$A$80&" ","")&IF($L$81=1,$A$81&" ",""))
    #   &" ")&IF(LUC_THAN!$Y$31="","","• Giai đoạn nên dồn sức: "&LUC_THAN!$Y$31&" ")&"• Nghề hợp hỷ dụng thần: "&SU_NGHIEP!$D$313
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),1.0)) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),1.0)) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),1.0)) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),1.0)) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),1.0)) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),1.0)) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),1.0)) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),1.0)) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),1.0)) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),1.0)) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),1.0)) else ''))),'')) else X.cat(X.cat('• Thời điểm nên hành động: ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(4,71,1),' ') if X.truthy(X.eq(E.v(4,71,12),1.0)) else ''),(X.cat(E.v(4,72,1),' ') if X.truthy(X.eq(E.v(4,72,12),1.0)) else '')),(X.cat(E.v(4,73,1),' ') if X.truthy(X.eq(E.v(4,73,12),1.0)) else '')),(X.cat(E.v(4,74,1),' ') if X.truthy(X.eq(E.v(4,74,12),1.0)) else '')),(X.cat(E.v(4,75,1),' ') if X.truthy(X.eq(E.v(4,75,12),1.0)) else '')),(X.cat(E.v(4,76,1),' ') if X.truthy(X.eq(E.v(4,76,12),1.0)) else '')),(X.cat(E.v(4,77,1),' ') if X.truthy(X.eq(E.v(4,77,12),1.0)) else '')),(X.cat(E.v(4,78,1),' ') if X.truthy(X.eq(E.v(4,78,12),1.0)) else '')),(X.cat(E.v(4,79,1),' ') if X.truthy(X.eq(E.v(4,79,12),1.0)) else '')),(X.cat(E.v(4,80,1),' ') if X.truthy(X.eq(E.v(4,80,12),1.0)) else '')),(X.cat(E.v(4,81,1),' ') if X.truthy(X.eq(E.v(4,81,12),1.0)) else '')))),' ')),('' if X.truthy(X.eq(E.v(16,31,25),'')) else X.cat(X.cat('• Giai đoạn nên dồn sức: ',E.v(16,31,25)),' '))),'• Nghề hợp hỷ dụng thần: '),E.v(21,313,4))

def f98(E, r, c):
    # Ô C85
    # = "Nguyên tắc chung: "&LUAN_GIAI!$H$106
    return X.cat('Nguyên tắc chung: ',E.v(4,106,8))

def f99(E, r, c):
    # Ô C86
    # = "Hóa giải năm xem: "&IF(LEN(KET_LUAN!$C$141)>600,LEFT(KET_LUAN!$C$141,600)&"…",KET_LUAN!$C$141)
    return X.cat('Hóa giải năm xem: ',(X.cat(X.LEFT(E.v(3,141,3), 600.0),'…') if X.truthy(X.gt(X.LEN(E.v(3,141,3)),600.0)) else E.v(3,141,3)))

def f100(E, r, c):
    # Ô C87
    # = "Nên phát triển: "&CHOOSE(MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7+5,5)+1,"làm việc nhóm, hợp tác, có bạn đồng hà
    #   nh; rèn ý chí, thể thao","phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình","công việc gắn với kết quả thực tế, quản lý tiền
    #   , kinh doanh","tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý","học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi 
    #   trường ổn định")&"; "&CHOOSE(MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)-TRUNG_GIAN!$E$7+5,5)+1,"làm việc nhóm, hợp tác, có bạn đồng
    #    hành; rèn ý chí, thể thao","phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình","công việc gắn với kết quả thực tế, quản lý t
    #   iền, kinh doanh","tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý","học tập liên tục, bằng cấp, người thầy – người đỡ đầu, m
    #   ôi trường ổn định")&"."&CHAR(10)&"Nên tránh: "&IFERROR(CHOOSE(MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)-TRUNG_GIAN!$E$7+5,5)+1,"hù
    #   n hạp, chung vốn, cạnh tranh trực diện; chi tiêu theo bạn bè","nói thẳng quá mức, đối đầu cấp trên, ôm quá nhiều việc cùng lúc","đầu tư mạo hiểm, bảo 
    #   lãnh vay, ôm đồm việc tiền bạc","môi trường áp lực cao, cấp bậc cứng nhắc, nhận trách nhiệm quá sức","ỷ lại, chần chừ, học mãi không làm, phụ thuộc ng
    #   ười khác"),"—")&"."&CHAR(10)&"Lên kế hoạch lớn vào đại vận, năm dụng/hỷ (bảng mục 7); năm xấu thì giữ ổn định, ưu tiên sức khỏe và quan hệ."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nên phát triển: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'làm việc nhóm, hợp tác, có bạn đồng hành; rèn ý chí, thể thao', lambda: 'phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình', lambda: 'công việc gắn với kết quả thực tế, quản lý tiền, kinh doanh', lambda: 'tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý', lambda: 'học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi trường ổn định')),'; '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'làm việc nhóm, hợp tác, có bạn đồng hành; rèn ý chí, thể thao', lambda: 'phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình', lambda: 'công việc gắn với kết quả thực tế, quản lý tiền, kinh doanh', lambda: 'tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý', lambda: 'học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi trường ổn định')),'.'),X.CHAR(10.0)),'Nên tránh: '),X.IFERROR(lambda: X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'hùn hạp, chung vốn, cạnh tranh trực diện; chi tiêu theo bạn bè', lambda: 'nói thẳng quá mức, đối đầu cấp trên, ôm quá nhiều việc cùng lúc', lambda: 'đầu tư mạo hiểm, bảo lãnh vay, ôm đồm việc tiền bạc', lambda: 'môi trường áp lực cao, cấp bậc cứng nhắc, nhận trách nhiệm quá sức', lambda: 'ỷ lại, chần chừ, học mãi không làm, phụ thuộc người khác'), lambda: '—')),'.'),X.CHAR(10.0)),'Lên kế hoạch lớn vào đại vận, năm dụng/hỷ (bảng mục 7); năm xấu thì giữ ổn định, ưu tiên sức khỏe và quan hệ.')

def f101(E, r, c):
    # Ô C88
    # = "Màu sắc, vật dụng hợp dụng thần "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1)&": "&CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$109,DUNG_HY_KY!$C$109)+1,"xanh lá","đỏ, hồng, tím","vàng, nâu đất","trắng, bạc, ánh kim","đen, xanh dương")&"; hỷ thần "&INDEX(HANH_TEN,IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1)&": "&CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1,"xanh lá","đỏ, hồng, tí
    #   m","vàng, nâu đất","trắng, bạc, ánh kim","đen, xanh dương")&"; hạn chế "&CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1,"xanh lá",
    #   "đỏ, hồng, tím","vàng, nâu đất","trắng, bạc, ánh kim","đen, xanh dương")&" (kỵ thần). Gợi ý dân gian, độ tin thấp. Kế hoạch hóa giải tổng hợp (sách + 
    #   công cụ + kinh nghiệm): KET_LUAN mục I."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Màu sắc, vật dụng hợp dụng thần ',X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),': '),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'xanh lá', lambda: 'đỏ, hồng, tím', lambda: 'vàng, nâu đất', lambda: 'trắng, bạc, ánh kim', lambda: 'đen, xanh dương')),'; hỷ thần '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))),': '),X.CHOOSE(X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0), lambda: 'xanh lá', lambda: 'đỏ, hồng, tím', lambda: 'vàng, nâu đất', lambda: 'trắng, bạc, ánh kim', lambda: 'đen, xanh dương')),'; hạn chế '),X.CHOOSE(X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0), lambda: 'xanh lá', lambda: 'đỏ, hồng, tím', lambda: 'vàng, nâu đất', lambda: 'trắng, bạc, ánh kim', lambda: 'đen, xanh dương')),' (kỵ thần). Gợi ý dân gian, độ tin thấp. Kế hoạch hóa giải tổng hợp (sách + công cụ + kinh nghiệm): KET_LUAN mục I.')

def f102(E, r, c):
    # Ô C89
    # = "Độ tin cậy và cách kiểm chứng: "&LUAN_GIAI!$C$140&" Ghi diễn biến thực tế vào sheet THEO_DOI (cả đời) – công cụ tự so khớp PP1/PP2 và đưa kết quả vào
    #    mục 9."
    return X.cat(X.cat('Độ tin cậy và cách kiểm chứng: ',E.v(4,140,3)),' Ghi diễn biến thực tế vào sheet THEO_DOI (cả đời) – công cụ tự so khớp PP1/PP2 và đưa kết quả vào mục 9.')

def f103(E, r, c):
    # Ô C91
    # = "• Bản mệnh: "&DOI_CHIEU!$C$34&"; dụng thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&", hỷ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET
    #   _LUAN!$C$37)&"."&CHAR(10)&"• Bốn trục mệnh cục: "&IFERROR(LEFT(MENH_CUC!$D$18,FIND(" (",MENH_CUC!$D$18)-1),MENH_CUC!$D$18)&"; "&IFERROR(LEFT(MENH_CUC!
    #   $D$19,FIND(" (",MENH_CUC!$D$19)-1),MENH_CUC!$D$19)&"; "&IFERROR(LEFT(MENH_CUC!$D$20,FIND(" (",MENH_CUC!$D$20)-1),MENH_CUC!$D$20)&"; "&IFERROR(LEFT(MEN
    #   H_CUC!$D$21,FIND(" (",MENH_CUC!$D$21)-1),MENH_CUC!$D$21)&"."&CHAR(10)&"• Sự nghiệp: quan vận "&LOWER(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHI
    #   EP!$J$7)-1),SU_NGHIEP!$J$7))&"; học tập "&LOWER(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))&"; tài vận "&LOWER(IFERR
    #   OR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))&"; bậc giàu: "&TAI_LOC!$E$17&"."&CHAR(10)&"• Gia đạo: hôn nhân "&LOWER(HON_NH
    #   AN!$X$64)&"; con cái "&LOWER(LUC_THAN!$X$34)&". Sức khỏe: "&LOWER(IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6))&"."&CHAR(
    #   10)&"• Giai đoạn thuận: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; cần thận trọng: "&IF(LUC_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31)&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Bản mệnh: ',E.v(7,34,3)),'; dụng thần '),(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),', hỷ thần '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'.'),X.CHAR(10.0)),'• Bốn trục mệnh cục: '),X.IFERROR(lambda: X.LEFT(E.v(18,18,4), X.sub(X.FIND(' (', E.v(18,18,4)),1.0)), lambda: E.v(18,18,4))),'; '),X.IFERROR(lambda: X.LEFT(E.v(18,19,4), X.sub(X.FIND(' (', E.v(18,19,4)),1.0)), lambda: E.v(18,19,4))),'; '),X.IFERROR(lambda: X.LEFT(E.v(18,20,4), X.sub(X.FIND(' (', E.v(18,20,4)),1.0)), lambda: E.v(18,20,4))),'; '),X.IFERROR(lambda: X.LEFT(E.v(18,21,4), X.sub(X.FIND(' (', E.v(18,21,4)),1.0)), lambda: E.v(18,21,4))),'.'),X.CHAR(10.0)),'• Sự nghiệp: quan vận '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))),'; học tập '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))),'; tài vận '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))),'; bậc giàu: '),E.v(19,17,5)),'.'),X.CHAR(10.0)),'• Gia đạo: hôn nhân '),X.LOWER(E.v(17,64,24))),'; con cái '),X.LOWER(E.v(16,34,24))),'. Sức khỏe: '),X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0)), lambda: E.v(22,6,10)))),'.'),X.CHAR(10.0)),'• Giai đoạn thuận: '),('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; cần thận trọng: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))),'.')

def f104(E, r, c):
    # Ô J92
    # = IF(OR(LEFT(MENH_CUC!$D$18,5)="THUẬN",LEFT(MENH_CUC!$D$18,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$18,3)="TỐT",MENH_CUC!$D$18="Khá",ISNUMBER(SEARCH("VỀ GIÀ
    #   U",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$18)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$18))),"•
    #    Giàu – nghèo: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$18,FIND(" (",MENH_CUC!$D$18)-1),MENH_CUC!$D$18))&" ","")&IF(OR(LEFT(MENH_CUC!$D$19,5)="THUẬN",LEFT(MEN
    #   H_CUC!$D$19,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$19,3)="TỐT",MENH_CUC!$D$19="Khá",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ SANG"
    #   ,MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$19)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$19))),"• Sang – hèn: "&LOWER(IFERROR(LEFT(MENH_CUC!$D
    #   $19,FIND(" (",MENH_CUC!$D$19)-1),MENH_CUC!$D$19))&" ","")&IF(OR(LEFT(MENH_CUC!$D$20,5)="THUẬN",LEFT(MENH_CUC!$D$20,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$
    #   D$20,3)="TỐT",MENH_CUC!$D$20="Khá",ISNUMBER(SEARCH("VỀ GIÀU",MENH_CUC!$D$20)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$20)),ISNUMBER(SEARCH("VỀ CÁT",MENH
    #   _CUC!$D$20)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!$D$20))),"• Cát – hung: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$20,FIND(" (",MENH_CUC!$D$20)-1),MENH_CUC!$D$20
    #   ))&" ","")&IF(OR(LEFT(MENH_CUC!$D$21,5)="THUẬN",LEFT(MENH_CUC!$D$21,13)="NGHIÊNG THUẬN",LEFT(MENH_CUC!$D$21,3)="TỐT",MENH_CUC!$D$21="Khá",ISNUMBER(SEA
    #   RCH("VỀ GIÀU",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ SANG",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ CÁT",MENH_CUC!$D$21)),ISNUMBER(SEARCH("VỀ THỌ",MENH_CUC!
    #   $D$21))),"• Thọ – yểu: "&LOWER(IFERROR(LEFT(MENH_CUC!$D$21,FIND(" (",MENH_CUC!$D$21)-1),MENH_CUC!$D$21))&" ","")&IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$
    #   6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),5)="THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),13)="N
    #   GHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),3)="TỐT",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_
    #   NGHIEP!$J$6)-1),SU_NGHIEP!$J$6)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))),ISNUMBE
    #   R(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$6
    #   ,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J
    #   $6)))),"• Tài vận: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6),FIND(" (",IFERROR(LEFT(SU_NGHIEP!$J$6,F
    #   IND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))-1),IFERROR(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6)-1),SU_NGHIEP!$J$6))&" ","")&IF(OR(LEFT(I
    #   FERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7),5)="THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-
    #   1),SU_NGHIEP!$J$7),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7),3)="TỐT",IFERROR(LEFT(SU_NGHI
    #   EP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),
    #   SU_NGHIEP!$J$7))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))),ISNUMBER(SEARCH("VỀ CÁT",IF
    #   ERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGH
    #   IEP!$J$7)-1),SU_NGHIEP!$J$7)))),"• Quan vận: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7),FIND(" (",IFE
    #   RROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$J$7))-1),IFERROR(LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7)-1),SU_NGHIEP!$
    #   J$7))&" ","")&IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8),5)="THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND
    #   (CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8),3)=
    #   "TỐT",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CH
    #   AR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))),I
    #   SNUMBER(SEARCH("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))),ISNUMBER(SEARCH("VỀ THỌ",IFERROR(LEFT(SU_NGHIEP
    #   !$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8)))),"• Học tập: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_N
    #   GHIEP!$J$8),FIND(" (",IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8)-1),SU_NGHIEP!$J$8))-1),IFERROR(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NG
    #   HIEP!$J$8)-1),SU_NGHIEP!$J$8))&" ","")&IF(OR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),5)="THUẬN",LEFT(IFERROR
    #   (LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),13)="NGHIÊNG THUẬN",LEFT(IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$
    #   9)-1),SU_NGHIEP!$J$9),3)="TỐT",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9)="Khá",ISNUMBER(SEARCH("VỀ GIÀU",IFERROR(LE
    #   FT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))),ISNUMBER(SEARCH("VỀ SANG",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$
    #   9)-1),SU_NGHIEP!$J$9))),ISNUMBER(SEARCH("VỀ CÁT",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))),ISNUMBER(SEARCH("VỀ TH
    #   Ọ",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9)))),"• Danh tiếng: "&IFERROR(LEFT(IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR
    #   (10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9),FIND(" (",IFERROR(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))-1),IFERROR(LEFT(SU_NGH
    #   IEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9)-1),SU_NGHIEP!$J$9))&" ","")&IF(OR(LUC_THAN!$X$30="Tốt",LUC_THAN!$X$30="Khá"),"• Ông bà: "&LOWER(LUC_THAN!$X$30)
    #   &" ","")&IF(OR(LUC_THAN!$X$31="Tốt",LUC_THAN!$X$31="Khá"),"• Cha: "&LOWER(LUC_THAN!$X$31)&" ","")&IF(OR(LUC_THAN!$X$32="Tốt",LUC_THAN!$X$32="Khá"),"• 
    #   Mẹ: "&LOWER(LUC_THAN!$X$32)&" ","")&IF(OR(LUC_THAN!$X$33="Tốt",LUC_THAN!$X$33="Khá"),"• Anh chị em: "&LOWER(LUC_THAN!$X$33)&" ","")&IF(OR(LUC_THAN!$X$
    #   34="Tốt",LUC_THAN!$X$34="Khá"),"• Con cái: "&LOWER(LUC_THAN!$X$34)&" ","")&IF(OR(HON_NHAN!$X$64="Tốt",HON_NHAN!$X$64="Khá"),"• Hôn nhân: "&LOWER(HON_N
    #   HAN!$X$64)&" ","")&IF(LUC_THAN!$Y$31="","","• Đại vận thuận: "&LUC_THAN!$Y$31&" ")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• Giàu – nghèo: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,18,4), X.sub(X.FIND(' (', E.v(18,18,4)),1.0)), lambda: E.v(18,18,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,18,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,18,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,18,4), 3.0),'TỐT'), X.eq(E.v(18,18,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,18,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,18,4))))) else ''),(X.cat(X.cat('• Sang – hèn: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,19,4), X.sub(X.FIND(' (', E.v(18,19,4)),1.0)), lambda: E.v(18,19,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,19,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,19,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,19,4), 3.0),'TỐT'), X.eq(E.v(18,19,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,19,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,19,4))))) else '')),(X.cat(X.cat('• Cát – hung: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,20,4), X.sub(X.FIND(' (', E.v(18,20,4)),1.0)), lambda: E.v(18,20,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,20,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,20,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,20,4), 3.0),'TỐT'), X.eq(E.v(18,20,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,20,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,20,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,20,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,20,4))))) else '')),(X.cat(X.cat('• Thọ – yểu: ',X.LOWER(X.IFERROR(lambda: X.LEFT(E.v(18,21,4), X.sub(X.FIND(' (', E.v(18,21,4)),1.0)), lambda: E.v(18,21,4)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(18,21,4), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,21,4), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,21,4), 3.0),'TỐT'), X.eq(E.v(18,21,4),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,21,4))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,21,4))))) else '')),(X.cat(X.cat('• Tài vận: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0)), lambda: E.v(21,6,10)))))) else '')),(X.cat(X.cat('• Quan vận: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0)), lambda: E.v(21,7,10)))))) else '')),(X.cat(X.cat('• Học tập: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0)), lambda: E.v(21,8,10)))))) else '')),(X.cat(X.cat('• Danh tiếng: ',X.IFERROR(lambda: X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), X.sub(X.FIND(' (', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10))),1.0)), lambda: X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))),' ') if X.truthy(X.OR(X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), 5.0),'THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)), 3.0),'TỐT'), X.eq(X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', X.IFERROR(lambda: X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0)), lambda: E.v(21,9,10)))))) else '')),(X.cat(X.cat('• Ông bà: ',X.LOWER(E.v(16,30,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,30,24),'Tốt'), X.eq(E.v(16,30,24),'Khá'))) else '')),(X.cat(X.cat('• Cha: ',X.LOWER(E.v(16,31,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,31,24),'Tốt'), X.eq(E.v(16,31,24),'Khá'))) else '')),(X.cat(X.cat('• Mẹ: ',X.LOWER(E.v(16,32,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,32,24),'Tốt'), X.eq(E.v(16,32,24),'Khá'))) else '')),(X.cat(X.cat('• Anh chị em: ',X.LOWER(E.v(16,33,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,33,24),'Tốt'), X.eq(E.v(16,33,24),'Khá'))) else '')),(X.cat(X.cat('• Con cái: ',X.LOWER(E.v(16,34,24))),' ') if X.truthy(X.OR(X.eq(E.v(16,34,24),'Tốt'), X.eq(E.v(16,34,24),'Khá'))) else '')),(X.cat(X.cat('• Hôn nhân: ',X.LOWER(E.v(17,64,24))),' ') if X.truthy(X.OR(X.eq(E.v(17,64,24),'Tốt'), X.eq(E.v(17,64,24),'Khá'))) else '')),('' if X.truthy(X.eq(E.v(16,31,25),'')) else X.cat(X.cat('• Đại vận thuận: ',E.v(16,31,25)),' ')))

def f105(E, r, c):
    # Ô C93
    # = "Chủ đề cuộc đời (theo nhóm dụng thần "&INDEX(NHOM_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7+5,5)+1)&"): "&CHO
    #   OSE(MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7+5,5)+1,"tự lực và hợp tác – thành công nhờ ý chí và người đồng hành"
    #   ,"tài năng và sáng tạo – thành công nhờ tay nghề, sản phẩm của chính mình","thực tế và tài chính – thành công nhờ nắm bắt cơ hội, quản lý tiền","trách
    #    nhiệm và địa vị – thành công nhờ kỷ luật, tổ chức, chức vụ","học vấn và quý nhân – thành công nhờ tri thức, bằng cấp, người nâng đỡ")&"."
    return X.cat(X.cat(X.cat(X.cat('Chủ đề cuộc đời (theo nhóm dụng thần ',X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)),5.0), 5.0),1.0))),'): '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tự lực và hợp tác – thành công nhờ ý chí và người đồng hành', lambda: 'tài năng và sáng tạo – thành công nhờ tay nghề, sản phẩm của chính mình', lambda: 'thực tế và tài chính – thành công nhờ nắm bắt cơ hội, quản lý tiền', lambda: 'trách nhiệm và địa vị – thành công nhờ kỷ luật, tổ chức, chức vụ', lambda: 'học vấn và quý nhân – thành công nhờ tri thức, bằng cấp, người nâng đỡ')),'.')

def f106(E, r, c):
    # Ô C94
    # = "Mệnh cục (P1 tr.164-173): "&LEFT(KET_LUAN!$C$87,IFERROR(FIND(CHAR(10),KET_LUAN!$C$87)-1,LEN(KET_LUAN!$C$87)))&" · "&LEFT(KET_LUAN!$C$99,IFERROR(FIND(
    #   CHAR(10),KET_LUAN!$C$99)-1,LEN(KET_LUAN!$C$99)))&" · "&LEFT(KET_LUAN!$C$173,IFERROR(FIND(CHAR(10),KET_LUAN!$C$173)-1,LEN(KET_LUAN!$C$173)))&" · "&LEFT
    #   (KET_LUAN!$C$143,IFERROR(FIND(CHAR(10),KET_LUAN!$C$143)-1,LEN(KET_LUAN!$C$143)))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mệnh cục (P1 tr.164-173): ',X.LEFT(E.v(3,87,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,87,3)),1.0), lambda: X.LEN(E.v(3,87,3))))),' · '),X.LEFT(E.v(3,99,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,99,3)),1.0), lambda: X.LEN(E.v(3,99,3))))),' · '),X.LEFT(E.v(3,173,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,173,3)),1.0), lambda: X.LEN(E.v(3,173,3))))),' · '),X.LEFT(E.v(3,143,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,143,3)),1.0), lambda: X.LEN(E.v(3,143,3)))))

def f107(E, r, c):
    # Ô C95
    # = "Bốn giai đoạn đời (P2 tr.284):"&CHAR(10)&IFERROR(TRIM(SUBSTITUTE(MID(KET_LUAN!$C$172,FIND("■ TRỤ NĂM",KET_LUAN!$C$172),FIND("Vì sao",KET_LUAN!$C$172,
    #   FIND("■ TRỤ NĂM",KET_LUAN!$C$172))-FIND("■ TRỤ NĂM",KET_LUAN!$C$172)),CHAR(10)," ")),"")&CHAR(10)&IFERROR(TRIM(SUBSTITUTE(MID(KET_LUAN!$C$172,FIND("■ 
    #   TRỤ THÁNG",KET_LUAN!$C$172),FIND("Vì sao",KET_LUAN!$C$172,FIND("■ TRỤ THÁNG",KET_LUAN!$C$172))-FIND("■ TRỤ THÁNG",KET_LUAN!$C$172)),CHAR(10)," ")),"")
    #   &CHAR(10)&IFERROR(TRIM(SUBSTITUTE(MID(KET_LUAN!$C$172,FIND("■ TRỤ NGÀY",KET_LUAN!$C$172),FIND("Vì sao",KET_LUAN!$C$172,FIND("■ TRỤ NGÀY",KET_LUAN!$C$1
    #   72))-FIND("■ TRỤ NGÀY",KET_LUAN!$C$172)),CHAR(10)," ")),"")&CHAR(10)&IFERROR(TRIM(SUBSTITUTE(MID(KET_LUAN!$C$172,FIND("■ TRỤ GIỜ",KET_LUAN!$C$172),FIN
    #   D("Vì sao",KET_LUAN!$C$172,FIND("■ TRỤ GIỜ",KET_LUAN!$C$172))-FIND("■ TRỤ GIỜ",KET_LUAN!$C$172)),CHAR(10)," ")),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bốn giai đoạn đời (P2 tr.284):',X.CHAR(10.0)),X.IFERROR(lambda: X.TRIM(X.SUBSTITUTE(X.MID(E.v(3,172,3), X.FIND('■ TRỤ NĂM', E.v(3,172,3)), X.sub(X.FIND('Vì sao', E.v(3,172,3), X.FIND('■ TRỤ NĂM', E.v(3,172,3))),X.FIND('■ TRỤ NĂM', E.v(3,172,3)))), X.CHAR(10.0), ' ')), lambda: '')),X.CHAR(10.0)),X.IFERROR(lambda: X.TRIM(X.SUBSTITUTE(X.MID(E.v(3,172,3), X.FIND('■ TRỤ THÁNG', E.v(3,172,3)), X.sub(X.FIND('Vì sao', E.v(3,172,3), X.FIND('■ TRỤ THÁNG', E.v(3,172,3))),X.FIND('■ TRỤ THÁNG', E.v(3,172,3)))), X.CHAR(10.0), ' ')), lambda: '')),X.CHAR(10.0)),X.IFERROR(lambda: X.TRIM(X.SUBSTITUTE(X.MID(E.v(3,172,3), X.FIND('■ TRỤ NGÀY', E.v(3,172,3)), X.sub(X.FIND('Vì sao', E.v(3,172,3), X.FIND('■ TRỤ NGÀY', E.v(3,172,3))),X.FIND('■ TRỤ NGÀY', E.v(3,172,3)))), X.CHAR(10.0), ' ')), lambda: '')),X.CHAR(10.0)),X.IFERROR(lambda: X.TRIM(X.SUBSTITUTE(X.MID(E.v(3,172,3), X.FIND('■ TRỤ GIỜ', E.v(3,172,3)), X.sub(X.FIND('Vì sao', E.v(3,172,3), X.FIND('■ TRỤ GIỜ', E.v(3,172,3))),X.FIND('■ TRỤ GIỜ', E.v(3,172,3)))), X.CHAR(10.0), ' ')), lambda: ''))

def f108(E, r, c):
    # Ô C96
    # = "Thế mạnh: "&CHOOSE(MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7+5,5)+1,"làm việc nhóm, hợp tác, có bạn đồng hành; rè
    #   n ý chí, thể thao","phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình","công việc gắn với kết quả thực tế, quản lý tiền, kinh
    #    doanh","tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý","học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi trường
    #    ổn định")&". Nút thắt: kỵ thần "&INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1)&" ("&IFERROR(INDEX(NHOM_TEN,MOD(IF(DOI_C
    #   HIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)-TRUNG_GIAN!$E$7+5,5)+1),"")&") – vận, năm mang hành này dễ “mất điểm”: "&IFERROR(CHOOSE(MOD(IF(DOI_C
    #   HIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)-TRUNG_GIAN!$E$7+5,5)+1,"hùn hạp, chung vốn, cạnh tranh trực diện; chi tiêu theo bạn bè","nói thẳng q
    #   uá mức, đối đầu cấp trên, ôm quá nhiều việc cùng lúc","đầu tư mạo hiểm, bảo lãnh vay, ôm đồm việc tiền bạc","môi trường áp lực cao, cấp bậc cứng nhắc,
    #    nhận trách nhiệm quá sức","ỷ lại, chần chừ, học mãi không làm, phụ thuộc người khác"),"")&"."&CHAR(10)&"Giai đoạn vàng (đại vận dụng/hỷ): "&IF(LUC_TH
    #   AN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; giai đoạn cần thận trọng (đại vận kỵ): "&IF(LUC_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31)&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thế mạnh: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'làm việc nhóm, hợp tác, có bạn đồng hành; rèn ý chí, thể thao', lambda: 'phát huy kỹ năng, chuyên môn, sáng tạo, trình bày; sản phẩm của chính mình', lambda: 'công việc gắn với kết quả thực tế, quản lý tiền, kinh doanh', lambda: 'tổ chức có kỷ luật, lộ trình thăng tiến rõ, vai trò trách nhiệm – quản lý', lambda: 'học tập liên tục, bằng cấp, người thầy – người đỡ đầu, môi trường ổn định')),'. Nút thắt: kỵ thần '),X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0))),' ('),X.IFERROR(lambda: X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),E.v(23,7,5)),5.0), 5.0),1.0)), lambda: '')),') – vận, năm mang hành này dễ “mất điểm”: '),X.IFERROR(lambda: X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'hùn hạp, chung vốn, cạnh tranh trực diện; chi tiêu theo bạn bè', lambda: 'nói thẳng quá mức, đối đầu cấp trên, ôm quá nhiều việc cùng lúc', lambda: 'đầu tư mạo hiểm, bảo lãnh vay, ôm đồm việc tiền bạc', lambda: 'môi trường áp lực cao, cấp bậc cứng nhắc, nhận trách nhiệm quá sức', lambda: 'ỷ lại, chần chừ, học mãi không làm, phụ thuộc người khác'), lambda: '')),'.'),X.CHAR(10.0)),'Giai đoạn vàng (đại vận dụng/hỷ): '),('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; giai đoạn cần thận trọng (đại vận kỵ): '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))),'.')

def f109(E, r, c):
    # Ô C97
    # = "Kiểm chứng bằng sự việc (THEO_DOI; P1 tr.125): đã ghi "&THEO_DOI!$E$4&" năm thực tế"&IF(THEO_DOI!$E$4=0,"",", khớp xu thế PP1 "&TEXT(THEO_DOI!$E$5,"0
    #   %")&", PP2 "&TEXT(THEO_DOI!$E$6,"0%"))&". "&THEO_DOI!$E$7&" (Xu thế tổng hợp gộp mọi mặt của năm nên thường khớp thấp hơn tín hiệu từng lĩnh vực – xem
    #    cột H–L và đánh giá cột P của THEO_DOI.)"&CHAR(10)&"Đánh giá luận giải theo từng sự kiện: "&THEO_DOI!$E$8
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Kiểm chứng bằng sự việc (THEO_DOI; P1 tr.125): đã ghi ',E.v(5,4,5)),' năm thực tế'),('' if X.truthy(X.eq(E.v(5,4,5),0.0)) else X.cat(X.cat(X.cat(', khớp xu thế PP1 ',X.TEXT(E.v(5,5,5), '0%')),', PP2 '),X.TEXT(E.v(5,6,5), '0%')))),'. '),E.v(5,7,5)),' (Xu thế tổng hợp gộp mọi mặt của năm nên thường khớp thấp hơn tín hiệu từng lĩnh vực – xem cột H–L và đánh giá cột P của THEO_DOI.)'),X.CHAR(10.0)),'Đánh giá luận giải theo từng sự kiện: '),E.v(5,8,5))

def f110(E, r, c):
    # Ô C98
    # = "Nguồn ngoài sách: Cùng Thông Bảo Giám – Giáp tháng Hợi https://www.dajiazhao.com/sm/qtbj/6088.html ; Bính hỏa https://blog.udn.com/DejaVu/3154129 ; M
    #   ậu thổ https://blog.udn.com/DejaVu/3154155 ; Tân kim https://m.163.com/dy/article/K7QOK6FQ0"&"55612AP.html ; Trích Thiên Tủy – Thiên can luận (10 câu 
    #   Hán-Việt ở cột ẩn L–M)."
    return X.cat('Nguồn ngoài sách: Cùng Thông Bảo Giám – Giáp tháng Hợi https://www.dajiazhao.com/sm/qtbj/6088.html ; Bính hỏa https://blog.udn.com/DejaVu/3154129 ; Mậu thổ https://blog.udn.com/DejaVu/3154155 ; Tân kim https://m.163.com/dy/article/K7QOK6FQ0','55612AP.html ; Trích Thiên Tủy – Thiên can luận (10 câu Hán-Việt ở cột ẩn L–M).')

def f111(E, r, c):
    # Ô C106
    # = "Nhật chủ "&TRUNG_GIAN!$E$10&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&") – "&DOI_CHIEU!$C$34
    return X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',E.v(23,10,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),') – '),E.v(7,34,3))

def f112(E, r, c):
    # Ô D106
    # = KET_LUAN!$C$34
    return E.v(3,34,3)

def f113(E, r, c):
    # Ô E106
    # = "Nhật chủ là can của ngày sinh, đại diện cho chính mình. “Vượng” nghĩa là khí của bản thân mạnh: sinh đúng mùa của hành mình (được lệnh), có gốc ở các
    #    chi, có nhiều chữ cùng phe (cùng hành hoặc sinh cho mình). “Nhược” là ngược lại. PP1 đếm đ"&"iều kiện (P1 tr.70); PP2 cộng độ, phe mình ≥ 40% là vượn
    #   g (P2 tr.332-352)."
    return X.cat('Nhật chủ là can của ngày sinh, đại diện cho chính mình. “Vượng” nghĩa là khí của bản thân mạnh: sinh đúng mùa của hành mình (được lệnh), có gốc ở các chi, có nhiều chữ cùng phe (cùng hành hoặc sinh cho mình). “Nhược” là ngược lại. PP1 đếm đ','iều kiện (P1 tr.70); PP2 cộng độ, phe mình ≥ 40% là vượng (P2 tr.332-352).')

def f114(E, r, c):
    # Ô F106
    # = IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),"Sức lực, ý chí tự thân mạnh; chịu được áp lực, thích tự quyết. Năng lượng thừa cần có ch
    #   ỗ “xả” (công việc, trách nhiệm, chi tiêu hợp lý) thì tốt.",IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"Sức tự thân có hạn; dễ mệt, d
    #   ễ bị hoàn cảnh chi phối. Làm việc cần chỗ dựa (học vấn, người giúp, cộng sự cùng chí hướng); không nên ôm đồm.","Khá cân bằng; ít thái quá, tùy vận mà
    #    nghiêng bên nào."))
    return ('Sức lực, ý chí tự thân mạnh; chịu được áp lực, thích tự quyết. Năng lượng thừa cần có chỗ “xả” (công việc, trách nhiệm, chi tiêu hợp lý) thì tốt.' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0)) else ('Sức tự thân có hạn; dễ mệt, dễ bị hoàn cảnh chi phối. Làm việc cần chỗ dựa (học vấn, người giúp, cộng sự cùng chí hướng); không nên ôm đồm.' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else 'Khá cân bằng; ít thái quá, tùy vận mà nghiêng bên nào.'))

def f115(E, r, c):
    # Ô G106, G107, G108, G116, G117, G118, G119, G123, G124, G125, G126, G144 … (30 ô)
    # = "Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LUC_THAN!$Z$31="","(khô
    #   ng có)",LUC_THAN!$Z$31)
    return X.cat(X.cat(X.cat('Rõ nhất trong đại vận dụng/hỷ thần: ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26)))

def f116(E, r, c):
    # Ô H106
    # = "Bổ dụng thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&", hỷ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$C$37)&"; hạn chế hành k
    #   ỵ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)&" (P2 tr.631-647). Nghề, phương hướng hợp dụng thần: "&LEFT(TO_CHAT!$C$17,IFERROR(FIND(CH
    #   AR(10),TO_CHAT!$C$17)-1,LEN(TO_CHAT!$C$17)))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bổ dụng thần ',(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),', hỷ thần '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'; hạn chế hành kỵ thần '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3))),' (P2 tr.631-647). Nghề, phương hướng hợp dụng thần: '),X.LEFT(E.v(20,17,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(20,17,3)),1.0), lambda: X.LEN(E.v(20,17,3)))))

def f117(E, r, c):
    # Ô I106
    # = DOI_CHIEU!$C$35&" – "&IF(DOI_CHIEU!$C$33=0,"hai phương pháp thống nhất","hai phương pháp khác nhau, đang theo "&IF(DOI_CHIEU!$C$33=2,"PP2","PP1")&" (c
    #   ó thể đổi ở NHAP_LIEU C14)")
    return X.cat(X.cat(E.v(7,35,3),' – '),('hai phương pháp thống nhất' if X.truthy(X.eq(E.v(7,33,3),0.0)) else X.cat(X.cat('hai phương pháp khác nhau, đang theo ',('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1')),' (có thể đổi ở NHAP_LIEU C14)')))

def f118(E, r, c):
    # Ô J106, J107, J108, J109, J110, J111, J112, J113, J114, J115, J116, J117 … (25 ô)
    # = "(a) "&C106&CHAR(10)&"(b) "&LEFT(D106,600)&CHAR(10)&"   Lời thường: "&E106&CHAR(10)&"(c) "&F106&CHAR(10)&"(d) "&G106&CHAR(10)&"(e) "&H106&CHAR(10)&"(f
    #   ) Độ tin cậy: "&I106
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('(a) ',E.v(4,r,c-7)),X.CHAR(10.0)),'(b) '),X.LEFT(E.v(4,r,c-6), 600.0)),X.CHAR(10.0)),'   Lời thường: '),E.v(4,r,c-5)),X.CHAR(10.0)),'(c) '),E.v(4,r,c-4)),X.CHAR(10.0)),'(d) '),E.v(4,r,c-3)),X.CHAR(10.0)),'(e) '),E.v(4,r,c-2)),X.CHAR(10.0)),'(f) Độ tin cậy: '),E.v(4,r,c-1))

def f119(E, r, c):
    # Ô C107
    # = "Dụng thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&"; hỷ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$C$37)&"; kỵ thần "&IF(DOI_
    #   CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dụng thần ',(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),'; hỷ thần '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'; kỵ thần '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3)))

def f120(E, r, c):
    # Ô D107
    # = KET_LUAN!$C$39
    return E.v(3,39,3)

def f121(E, r, c):
    # Ô E107
    # = "Dụng thần là hành mà lá số cần nhất để cân bằng (thân vượng thì cần hành làm hao bớt, thân nhược thì cần hành sinh – giúp mình); hỷ thần là hành hỗ t
    #   rợ dụng thần; kỵ thần là hành làm lá số lệch thêm (P1 tr.124-133)."
    return 'Dụng thần là hành mà lá số cần nhất để cân bằng (thân vượng thì cần hành làm hao bớt, thân nhược thì cần hành sinh – giúp mình); hỷ thần là hành hỗ trợ dụng thần; kỵ thần là hành làm lá số lệch thêm (P1 tr.124-133).'

def f122(E, r, c):
    # Ô F107
    # = "Những năm, vận, nghề, phương hướng, môi trường mang hành dụng/hỷ thần thường thuận lợi, dễ được giúp; mang hành kỵ thần thì dễ trở ngại, hao tổn."
    return 'Những năm, vận, nghề, phương hướng, môi trường mang hành dụng/hỷ thần thường thuận lợi, dễ được giúp; mang hành kỵ thần thì dễ trở ngại, hao tổn.'

def f123(E, r, c):
    # Ô H107
    # = "Chọn nghề, nơi ở, phương hướng, màu sắc thuộc hành "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&" hoặc "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37
    #   ,KET_LUAN!$C$37)&"; tránh dồn thêm hành "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)&" (P2 tr.631-647; TO_CHAT, SUC_KHOE mục E)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chọn nghề, nơi ở, phương hướng, màu sắc thuộc hành ',(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),' hoặc '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'; tránh dồn thêm hành '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3))),' (P2 tr.631-647; TO_CHAT, SUC_KHOE mục E).')

def f124(E, r, c):
    # Ô I107
    # = DOI_CHIEU!$C$35&" (dụng thần phụ thuộc vào kết luận vượng suy)"
    return X.cat(E.v(7,35,3),' (dụng thần phụ thuộc vào kết luận vượng suy)')

def f125(E, r, c):
    # Ô C108
    # = LEFT(IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$35,KET_LUAN!$C$35),IFERROR(FIND(CHAR(10),IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$35,KET_LUAN!$C$35))-1,LEN(IF(DOI_CHIEU
    #   !$C$33=2,KET_LUAN!$D$35,KET_LUAN!$C$35))))
    return X.LEFT((E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3)), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), (E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3))),1.0), lambda: X.LEN((E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3)))))

def f126(E, r, c):
    # Ô D108
    # = IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$35,KET_LUAN!$C$35)
    return (E.v(3,35,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,35,3))

def f127(E, r, c):
    # Ô E108
    # = "Cách cục là “dáng” chung của lá số: thường (chính cách, xét vượng suy) hay đặc biệt (tòng, hóa, chuyên vượng…) khi một phe quá áp đảo (P1 tr.104-123)
    #   ."
    return 'Cách cục là “dáng” chung của lá số: thường (chính cách, xét vượng suy) hay đặc biệt (tòng, hóa, chuyên vượng…) khi một phe quá áp đảo (P1 tr.104-123).'

def f128(E, r, c):
    # Ô F108
    # = "Cách thường: luận theo cân bằng vượng – suy. Cách đặc biệt: thuận theo thế mạnh nhất thì tốt, đi ngược thì xấu."
    return 'Cách thường: luận theo cân bằng vượng – suy. Cách đặc biệt: thuận theo thế mạnh nhất thì tốt, đi ngược thì xấu.'

def f129(E, r, c):
    # Ô H108
    # = "Theo dụng thần đã chọn (ở trên)."
    return 'Theo dụng thần đã chọn (ở trên).'

def f130(E, r, c):
    # Ô I108, I121, I130, I144, I145, I146, I147, I148, I149, I152, I153, I155 … (31 ô)
    # = DOI_CHIEU!$C$35
    return E.v(7,35,3)

def f131(E, r, c):
    # Ô C109
    # = KET_LUAN!$C$62&IF(KET_LUAN!$C$65="","",CHAR(10)&KET_LUAN!$C$65)
    return X.cat(E.v(3,62,3),('' if X.truthy(X.eq(E.v(3,65,3),'')) else X.cat(X.CHAR(10.0),E.v(3,65,3))))

def f132(E, r, c):
    # Ô D109
    # = KET_LUAN!$C$64
    return E.v(3,64,3)

def f133(E, r, c):
    # Ô E109
    # = "Thần sát là các “sao” tra theo bảng (quý nhân, văn xương, đào hoa, kình dương…). Sách coi đây là yếu tố PHỤ, sao tốt gặp hình xung khắc hại thì “như 
    #   không có” (P1 tr.137-140; P2 tr.294-326)."
    return 'Thần sát là các “sao” tra theo bảng (quý nhân, văn xương, đào hoa, kình dương…). Sách coi đây là yếu tố PHỤ, sao tốt gặp hình xung khắc hại thì “như không có” (P1 tr.137-140; P2 tr.294-326).'

def f134(E, r, c):
    # Ô F109
    # = "Sao tốt: thêm may mắn, được giúp; sao xấu: thêm rủi ro ở mặt sao đó chủ. Không quyết định cả lá số."
    return 'Sao tốt: thêm may mắn, được giúp; sao xấu: thêm rủi ro ở mặt sao đó chủ. Không quyết định cả lá số.'

def f135(E, r, c):
    # Ô G109
    # = "Ứng rõ khi đại vận/lưu niên chạm vào trụ có sao hoặc mang lại chính sao đó."
    return 'Ứng rõ khi đại vận/lưu niên chạm vào trụ có sao hoặc mang lại chính sao đó.'

def f136(E, r, c):
    # Ô H109
    # = "Dùng như lời nhắc: phát huy sao tốt (học hành, quan hệ), đề phòng mặt sao xấu chủ."
    return 'Dùng như lời nhắc: phát huy sao tốt (học hành, quan hệ), đề phòng mặt sao xấu chủ.'

def f137(E, r, c):
    # Ô I109, I173
    # = "Thấp – Vừa (yếu tố phụ)"
    return 'Thấp – Vừa (yếu tố phụ)'

def f138(E, r, c):
    # Ô C110
    # = LEFT(LUC_THAN!$D$30,IFERROR(FIND(CHAR(10),LUC_THAN!$D$30)-1,LEN(LUC_THAN!$D$30)))
    return X.LEFT(E.v(16,30,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(16,30,4)),1.0), lambda: X.LEN(E.v(16,30,4))))

def f139(E, r, c):
    # Ô D110
    # = LUC_THAN!$G$30
    return E.v(16,30,7)

def f140(E, r, c):
    # Ô E110
    # = "Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: trụ năm là cung tổ tiên;
    #    tài, quan, ấn ở trụ năm tốt thì gốc gác tốt."
    return 'Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: trụ năm là cung tổ tiên; tài, quan, ấn ở trụ năm tốt thì gốc gác tốt.'

def f141(E, r, c):
    # Ô F110
    # = IF(OR(LEFT(LUC_THAN!$X$30,5)="THUẬN",LEFT(LUC_THAN!$X$30,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(LUC_THAN!$X$30,20))),"Duyên mỏng hoặc người thâ
    #   n vất vả; dễ xa cách, phải tự lập nhiều hơn.","Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn."),IF(LEFT(LUC_THAN!$X$30,4)="BẤT ","Duyên mỏ
    #   ng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$30,3)="TỐT",LEFT(LUC_THAN!$X$30,3)="Tốt",LEFT(LUC_THAN!$X$30,3)=
    #   "Khá"),"Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$30,3)="KÉM",LEFT(LUC_THAN!$X$30,3)="Kém",LEFT(LUC_THAN!$X$3
    #   0,6)="Rất k"),"Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.","Bình thường, tốt xấu tùy giai đoạn."))))
    return (('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(16,30,24), 20.0)))) else 'Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,30,24), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(16,30,24), 5.0),'NGHIÊ'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.eq(X.LEFT(E.v(16,30,24), 4.0),'BẤT ')) else ('Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,30,24), 3.0),'TỐT'), X.eq(X.LEFT(E.v(16,30,24), 3.0),'Tốt'), X.eq(X.LEFT(E.v(16,30,24), 3.0),'Khá'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,30,24), 3.0),'KÉM'), X.eq(X.LEFT(E.v(16,30,24), 3.0),'Kém'), X.eq(X.LEFT(E.v(16,30,24), 6.0),'Rất k'))) else 'Bình thường, tốt xấu tùy giai đoạn.'))))

def f142(E, r, c):
    # Ô G110
    # = IF(LUC_THAN!$H$30="","Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LU
    #   C_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31),LUC_THAN!$H$30)
    return (X.cat(X.cat(X.cat('Rõ nhất trong đại vận dụng/hỷ thần: ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))) if X.truthy(X.eq(E.v(16,30,8),'')) else E.v(16,30,8))

def f143(E, r, c):
    # Ô H110, H111, H112, H113, H114
    # = "Giữ liên lạc, chăm sóc nhiều hơn trong những năm/vận được báo; vận kỵ thần nên tránh xung đột, chú ý sức khỏe người thân."
    return 'Giữ liên lạc, chăm sóc nhiều hơn trong những năm/vận được báo; vận kỵ thần nên tránh xung đột, chú ý sức khỏe người thân.'

def f144(E, r, c):
    # Ô I110
    # = IF(ABS(LUC_THAN!$W$30)>=3,"Khá",IF(ABS(LUC_THAN!$W$30)>=1,"Vừa","Thấp"))&" (điểm chấm "&LUC_THAN!$W$30&"); vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(E.v(16,30,23)),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(E.v(16,30,23)),1.0)) else 'Thấp')),' (điểm chấm '),E.v(16,30,23)),'); vượng suy: '),E.v(7,35,3))

def f145(E, r, c):
    # Ô C111
    # = LEFT(LUC_THAN!$D$31,IFERROR(FIND(CHAR(10),LUC_THAN!$D$31)-1,LEN(LUC_THAN!$D$31)))
    return X.LEFT(E.v(16,31,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(16,31,4)),1.0), lambda: X.LEN(E.v(16,31,4))))

def f146(E, r, c):
    # Ô D111
    # = LUC_THAN!$G$31
    return E.v(16,31,7)

def f147(E, r, c):
    # Ô E111
    # = "Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao cha là thiên tài; đư
    #   ợc ngôi, vượng, là dụng/hỷ thì cha có lực, mình được nhờ."
    return 'Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao cha là thiên tài; được ngôi, vượng, là dụng/hỷ thì cha có lực, mình được nhờ.'

def f148(E, r, c):
    # Ô F111
    # = IF(OR(LEFT(LUC_THAN!$X$31,5)="THUẬN",LEFT(LUC_THAN!$X$31,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(LUC_THAN!$X$31,20))),"Duyên mỏng hoặc người thâ
    #   n vất vả; dễ xa cách, phải tự lập nhiều hơn.","Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn."),IF(LEFT(LUC_THAN!$X$31,4)="BẤT ","Duyên mỏ
    #   ng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$31,3)="TỐT",LEFT(LUC_THAN!$X$31,3)="Tốt",LEFT(LUC_THAN!$X$31,3)=
    #   "Khá"),"Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$31,3)="KÉM",LEFT(LUC_THAN!$X$31,3)="Kém",LEFT(LUC_THAN!$X$3
    #   1,6)="Rất k"),"Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.","Bình thường, tốt xấu tùy giai đoạn."))))
    return (('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(16,31,24), 20.0)))) else 'Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,31,24), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(16,31,24), 5.0),'NGHIÊ'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.eq(X.LEFT(E.v(16,31,24), 4.0),'BẤT ')) else ('Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,31,24), 3.0),'TỐT'), X.eq(X.LEFT(E.v(16,31,24), 3.0),'Tốt'), X.eq(X.LEFT(E.v(16,31,24), 3.0),'Khá'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,31,24), 3.0),'KÉM'), X.eq(X.LEFT(E.v(16,31,24), 3.0),'Kém'), X.eq(X.LEFT(E.v(16,31,24), 6.0),'Rất k'))) else 'Bình thường, tốt xấu tùy giai đoạn.'))))

def f149(E, r, c):
    # Ô G111
    # = IF(LUC_THAN!$H$31="","Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LU
    #   C_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31),LUC_THAN!$H$31)
    return (X.cat(X.cat(X.cat('Rõ nhất trong đại vận dụng/hỷ thần: ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))) if X.truthy(X.eq(E.v(16,31,8),'')) else E.v(16,31,8))

def f150(E, r, c):
    # Ô I111
    # = IF(ABS(LUC_THAN!$W$31)>=3,"Khá",IF(ABS(LUC_THAN!$W$31)>=1,"Vừa","Thấp"))&" (điểm chấm "&LUC_THAN!$W$31&"); vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(E.v(16,31,23)),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(E.v(16,31,23)),1.0)) else 'Thấp')),' (điểm chấm '),E.v(16,31,23)),'); vượng suy: '),E.v(7,35,3))

def f151(E, r, c):
    # Ô C112
    # = LEFT(LUC_THAN!$D$32,IFERROR(FIND(CHAR(10),LUC_THAN!$D$32)-1,LEN(LUC_THAN!$D$32)))
    return X.LEFT(E.v(16,32,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(16,32,4)),1.0), lambda: X.LEN(E.v(16,32,4))))

def f152(E, r, c):
    # Ô D112
    # = LUC_THAN!$G$32
    return E.v(16,32,7)

def f153(E, r, c):
    # Ô E112
    # = "Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao mẹ là chính ấn (khôn
    #   g có thì thiên ấn); được ngôi, vượng, là dụng/hỷ thì mẹ có lực."
    return 'Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao mẹ là chính ấn (không có thì thiên ấn); được ngôi, vượng, là dụng/hỷ thì mẹ có lực.'

def f154(E, r, c):
    # Ô F112
    # = IF(OR(LEFT(LUC_THAN!$X$32,5)="THUẬN",LEFT(LUC_THAN!$X$32,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(LUC_THAN!$X$32,20))),"Duyên mỏng hoặc người thâ
    #   n vất vả; dễ xa cách, phải tự lập nhiều hơn.","Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn."),IF(LEFT(LUC_THAN!$X$32,4)="BẤT ","Duyên mỏ
    #   ng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$32,3)="TỐT",LEFT(LUC_THAN!$X$32,3)="Tốt",LEFT(LUC_THAN!$X$32,3)=
    #   "Khá"),"Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$32,3)="KÉM",LEFT(LUC_THAN!$X$32,3)="Kém",LEFT(LUC_THAN!$X$3
    #   2,6)="Rất k"),"Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.","Bình thường, tốt xấu tùy giai đoạn."))))
    return (('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(16,32,24), 20.0)))) else 'Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,32,24), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(16,32,24), 5.0),'NGHIÊ'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.eq(X.LEFT(E.v(16,32,24), 4.0),'BẤT ')) else ('Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,32,24), 3.0),'TỐT'), X.eq(X.LEFT(E.v(16,32,24), 3.0),'Tốt'), X.eq(X.LEFT(E.v(16,32,24), 3.0),'Khá'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,32,24), 3.0),'KÉM'), X.eq(X.LEFT(E.v(16,32,24), 3.0),'Kém'), X.eq(X.LEFT(E.v(16,32,24), 6.0),'Rất k'))) else 'Bình thường, tốt xấu tùy giai đoạn.'))))

def f155(E, r, c):
    # Ô G112
    # = IF(LUC_THAN!$H$32="","Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LU
    #   C_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31),LUC_THAN!$H$32)
    return (X.cat(X.cat(X.cat('Rõ nhất trong đại vận dụng/hỷ thần: ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))) if X.truthy(X.eq(E.v(16,32,8),'')) else E.v(16,32,8))

def f156(E, r, c):
    # Ô I112
    # = IF(ABS(LUC_THAN!$W$32)>=3,"Khá",IF(ABS(LUC_THAN!$W$32)>=1,"Vừa","Thấp"))&" (điểm chấm "&LUC_THAN!$W$32&"); vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(E.v(16,32,23)),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(E.v(16,32,23)),1.0)) else 'Thấp')),' (điểm chấm '),E.v(16,32,23)),'); vượng suy: '),E.v(7,35,3))

def f157(E, r, c):
    # Ô C113
    # = LEFT(LUC_THAN!$D$33,IFERROR(FIND(CHAR(10),LUC_THAN!$D$33)-1,LEN(LUC_THAN!$D$33)))
    return X.LEFT(E.v(16,33,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(16,33,4)),1.0), lambda: X.LEN(E.v(16,33,4))))

def f158(E, r, c):
    # Ô D113
    # = LUC_THAN!$G$33
    return E.v(16,33,7)

def f159(E, r, c):
    # Ô E113
    # = "Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao anh em là tỉ kiên, k
    #   iếp tài; cung anh em là trụ tháng."
    return 'Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao anh em là tỉ kiên, kiếp tài; cung anh em là trụ tháng.'

def f160(E, r, c):
    # Ô F113
    # = IF(OR(LEFT(LUC_THAN!$X$33,5)="THUẬN",LEFT(LUC_THAN!$X$33,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(LUC_THAN!$X$33,20))),"Duyên mỏng hoặc người thâ
    #   n vất vả; dễ xa cách, phải tự lập nhiều hơn.","Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn."),IF(LEFT(LUC_THAN!$X$33,4)="BẤT ","Duyên mỏ
    #   ng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$33,3)="TỐT",LEFT(LUC_THAN!$X$33,3)="Tốt",LEFT(LUC_THAN!$X$33,3)=
    #   "Khá"),"Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$33,3)="KÉM",LEFT(LUC_THAN!$X$33,3)="Kém",LEFT(LUC_THAN!$X$3
    #   3,6)="Rất k"),"Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.","Bình thường, tốt xấu tùy giai đoạn."))))
    return (('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(16,33,24), 20.0)))) else 'Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,33,24), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(16,33,24), 5.0),'NGHIÊ'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.eq(X.LEFT(E.v(16,33,24), 4.0),'BẤT ')) else ('Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,33,24), 3.0),'TỐT'), X.eq(X.LEFT(E.v(16,33,24), 3.0),'Tốt'), X.eq(X.LEFT(E.v(16,33,24), 3.0),'Khá'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,33,24), 3.0),'KÉM'), X.eq(X.LEFT(E.v(16,33,24), 3.0),'Kém'), X.eq(X.LEFT(E.v(16,33,24), 6.0),'Rất k'))) else 'Bình thường, tốt xấu tùy giai đoạn.'))))

def f161(E, r, c):
    # Ô G113
    # = IF(LUC_THAN!$H$33="","Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LU
    #   C_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31),LUC_THAN!$H$33)
    return (X.cat(X.cat(X.cat('Rõ nhất trong đại vận dụng/hỷ thần: ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))) if X.truthy(X.eq(E.v(16,33,8),'')) else E.v(16,33,8))

def f162(E, r, c):
    # Ô I113
    # = IF(ABS(LUC_THAN!$W$33)>=3,"Khá",IF(ABS(LUC_THAN!$W$33)>=1,"Vừa","Thấp"))&" (điểm chấm "&LUC_THAN!$W$33&"); vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(E.v(16,33,23)),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(E.v(16,33,23)),1.0)) else 'Thấp')),' (điểm chấm '),E.v(16,33,23)),'); vượng suy: '),E.v(7,35,3))

def f163(E, r, c):
    # Ô C114
    # = LEFT(LUC_THAN!$D$34,IFERROR(FIND(CHAR(10),LUC_THAN!$D$34)-1,LEN(LUC_THAN!$D$34)))
    return X.LEFT(E.v(16,34,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(16,34,4)),1.0), lambda: X.LEN(E.v(16,34,4))))

def f164(E, r, c):
    # Ô D114
    # = LUC_THAN!$G$34
    return E.v(16,34,7)

def f165(E, r, c):
    # Ô E114
    # = "Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao con theo P1 là thực 
    #   thần, thương quan; cung con là trụ giờ."
    return 'Cách chấm của sách (P1 tr.159): “đóng đúng bản cung là được ngôi, gặp nguyệt lệnh là vượng, gặp dụng thần là có lực”. Ở đây: sao con theo P1 là thực thần, thương quan; cung con là trụ giờ.'

def f166(E, r, c):
    # Ô F114
    # = IF(OR(LEFT(LUC_THAN!$X$34,5)="THUẬN",LEFT(LUC_THAN!$X$34,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(LUC_THAN!$X$34,20))),"Duyên mỏng hoặc người thâ
    #   n vất vả; dễ xa cách, phải tự lập nhiều hơn.","Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn."),IF(LEFT(LUC_THAN!$X$34,4)="BẤT ","Duyên mỏ
    #   ng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$34,3)="TỐT",LEFT(LUC_THAN!$X$34,3)="Tốt",LEFT(LUC_THAN!$X$34,3)=
    #   "Khá"),"Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.",IF(OR(LEFT(LUC_THAN!$X$34,3)="KÉM",LEFT(LUC_THAN!$X$34,3)="Kém",LEFT(LUC_THAN!$X$3
    #   4,6)="Rất k"),"Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.","Bình thường, tốt xấu tùy giai đoạn."))))
    return (('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(16,34,24), 20.0)))) else 'Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,34,24), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(16,34,24), 5.0),'NGHIÊ'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.eq(X.LEFT(E.v(16,34,24), 4.0),'BẤT ')) else ('Quan hệ thuận, người thân có lực, mình được nhờ nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,34,24), 3.0),'TỐT'), X.eq(X.LEFT(E.v(16,34,24), 3.0),'Tốt'), X.eq(X.LEFT(E.v(16,34,24), 3.0),'Khá'))) else ('Duyên mỏng hoặc người thân vất vả; dễ xa cách, phải tự lập nhiều hơn.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(16,34,24), 3.0),'KÉM'), X.eq(X.LEFT(E.v(16,34,24), 3.0),'Kém'), X.eq(X.LEFT(E.v(16,34,24), 6.0),'Rất k'))) else 'Bình thường, tốt xấu tùy giai đoạn.'))))

def f167(E, r, c):
    # Ô G114
    # = IF(LUC_THAN!$H$34="","Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LU
    #   C_THAN!$Z$31="","(không có)",LUC_THAN!$Z$31),LUC_THAN!$H$34)
    return (X.cat(X.cat(X.cat('Rõ nhất trong đại vận dụng/hỷ thần: ',('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))) if X.truthy(X.eq(E.v(16,34,8),'')) else E.v(16,34,8))

def f168(E, r, c):
    # Ô I114
    # = IF(ABS(LUC_THAN!$W$34)>=3,"Khá",IF(ABS(LUC_THAN!$W$34)>=1,"Vừa","Thấp"))&" (điểm chấm "&LUC_THAN!$W$34&"); vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(E.v(16,34,23)),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(E.v(16,34,23)),1.0)) else 'Thấp')),' (điểm chấm '),E.v(16,34,23)),'); vượng suy: '),E.v(7,35,3))

def f169(E, r, c):
    # Ô C115
    # = LEFT(HON_NHAN!$D$64,IFERROR(FIND(CHAR(10),HON_NHAN!$D$64)-1,LEN(HON_NHAN!$D$64)))
    return X.LEFT(E.v(17,64,4), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(17,64,4)),1.0), lambda: X.LEN(E.v(17,64,4))))

def f170(E, r, c):
    # Ô D115
    # = HON_NHAN!$G$64
    return E.v(17,64,7)

def f171(E, r, c):
    # Ô E115
    # = "Nam lấy tài làm sao vợ, nữ lấy quan sát làm sao chồng; chi ngày là cung phối ngẫu (P1 tr.153-155; P2 tr.463-481). Mức ở đây đánh giá người phối ngẫu 
    #   có lực hay không; sự êm ấm xem thêm các dấu hiệu bất lợi."
    return 'Nam lấy tài làm sao vợ, nữ lấy quan sát làm sao chồng; chi ngày là cung phối ngẫu (P1 tr.153-155; P2 tr.463-481). Mức ở đây đánh giá người phối ngẫu có lực hay không; sự êm ấm xem thêm các dấu hiệu bất lợi.'

def f172(E, r, c):
    # Ô F115
    # = "Sao "&IF(NHAP_LIEU!$C$22="Nam","vợ (sao Tài)","chồng (sao Quan sát)")&" mức "&LOWER(HON_NHAN!$X$64)&"; sách có "&HON_NHAN!$AI$64&" dấu hiệu thuận, "&
    #   HON_NHAN!$AG$64&" dấu hiệu bất lợi về hôn nhân → "&IF(OR(HON_NHAN!$X$64="Tốt",HON_NHAN!$X$64="Khá"),IF(HON_NHAN!$AG$64>HON_NHAN!$AI$64,"người phối ngẫ
    #   u có lực nhưng sự êm ấm dễ va chạm, xa cách; cần nhường nhịn, vun đắp.","người phối ngẫu có lực, giúp được mình; hôn nhân tương đối thuận."),IF(HON_NH
    #   AN!$X$64="Trung bình",IF(HON_NHAN!$AG$64>HON_NHAN!$AI$64,"hôn nhân bình thường nhưng dễ va chạm; cần nhường nhịn, vun đắp.","hôn nhân bình thường, tùy
    #    vận."),IF(HON_NHAN!$AG$64>0,"người phối ngẫu ít trợ lực, hôn nhân nhiều trắc trở (dễ va chạm, xa cách); cần nhường nhịn, vun đắp.","người phối ngẫu í
    #   t trợ lực hoặc hôn nhân nhiều trắc trở; cần nhường nhịn, vun đắp."))&IF(LUU_NIEN!$B$5>=40," (Lá số đã qua tuổi kết hôn thông thường: các ý như ‘nên kế
    #   t hôn muộn’ đọc là đặc điểm, không còn là lời khuyên.)",""))&IF(LUU_NIEN!$B$5<=18," (Lưu ý tuổi: chưa thành niên – ý về hôn nhân ở đây là xu hướng xa,
    #    chưa xét lúc này.)","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Sao ',('vợ (sao Tài)' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else 'chồng (sao Quan sát)')),' mức '),X.LOWER(E.v(17,64,24))),'; sách có '),E.v(17,64,35)),' dấu hiệu thuận, '),E.v(17,64,33)),' dấu hiệu bất lợi về hôn nhân → '),(('người phối ngẫu có lực nhưng sự êm ấm dễ va chạm, xa cách; cần nhường nhịn, vun đắp.' if X.truthy(X.gt(E.v(17,64,33),E.v(17,64,35))) else 'người phối ngẫu có lực, giúp được mình; hôn nhân tương đối thuận.') if X.truthy(X.OR(X.eq(E.v(17,64,24),'Tốt'), X.eq(E.v(17,64,24),'Khá'))) else X.cat((('hôn nhân bình thường nhưng dễ va chạm; cần nhường nhịn, vun đắp.' if X.truthy(X.gt(E.v(17,64,33),E.v(17,64,35))) else 'hôn nhân bình thường, tùy vận.') if X.truthy(X.eq(E.v(17,64,24),'Trung bình')) else ('người phối ngẫu ít trợ lực, hôn nhân nhiều trắc trở (dễ va chạm, xa cách); cần nhường nhịn, vun đắp.' if X.truthy(X.gt(E.v(17,64,33),0.0)) else 'người phối ngẫu ít trợ lực hoặc hôn nhân nhiều trắc trở; cần nhường nhịn, vun đắp.')),(' (Lá số đã qua tuổi kết hôn thông thường: các ý như ‘nên kết hôn muộn’ đọc là đặc điểm, không còn là lời khuyên.)' if X.truthy(X.ge(E.v(15,5,2),40.0)) else '')))),(' (Lưu ý tuổi: chưa thành niên – ý về hôn nhân ở đây là xu hướng xa, chưa xét lúc này.)' if X.truthy(X.le(E.v(15,5,2),18.0)) else ''))

def f173(E, r, c):
    # Ô G115
    # = HON_NHAN!$D$66&" "&HON_NHAN!$H$64
    return X.cat(X.cat(E.v(17,66,4),' '),E.v(17,64,8))

def f174(E, r, c):
    # Ô H115
    # = "Kết hôn, quyết định lớn về gia đình nên chọn năm thuận (HON_NHAN); năm bất lợi nên giữ hòa khí, tránh quyết định vội."
    return 'Kết hôn, quyết định lớn về gia đình nên chọn năm thuận (HON_NHAN); năm bất lợi nên giữ hòa khí, tránh quyết định vội.'

def f175(E, r, c):
    # Ô I115
    # = IF(ABS(HON_NHAN!$W$64)>=3,"Khá",IF(ABS(HON_NHAN!$W$64)>=1,"Vừa","Thấp"))&"; vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(E.v(17,64,23)),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(E.v(17,64,23)),1.0)) else 'Thấp')),'; vượng suy: '),E.v(7,35,3))

def f176(E, r, c):
    # Ô C116
    # = MENH_CUC!$D$18
    return E.v(18,18,4)

def f177(E, r, c):
    # Ô D116
    # = MENH_CUC!$G$18
    return E.v(18,18,7)

def f178(E, r, c):
    # Ô E116, E117, E118, E119
    # = "Bảng tóm tắt P1 tr.173: dựa vào vượng suy nhật chủ và nhóm sao là dụng hay kỵ (ví dụ giàu: thân mạnh mà tài là dụng)."
    return 'Bảng tóm tắt P1 tr.173: dựa vào vượng suy nhật chủ và nhóm sao là dụng hay kỵ (ví dụ giàu: thân mạnh mà tài là dụng).'

def f179(E, r, c):
    # Ô F116
    # = IF(OR(ISNUMBER(SEARCH("GIÀU",MENH_CUC!$U$18)),ISNUMBER(SEARCH("SANG",MENH_CUC!$U$18)),ISNUMBER(SEARCH("CÁT",MENH_CUC!$U$18)),ISNUMBER(SEARCH("THỌ",MEN
    #   H_CUC!$U$18))),"Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.",IF(OR(ISNUMBER(SEARCH("NGHÈO",MENH_CUC!$U$18)),ISNUMBER(SEARCH("HÈN",MENH_CUC!
    #   $U$18)),ISNUMBER(SEARCH("HUNG",MENH_CUC!$U$18)),ISNUMBER(SEARCH("YỂU",MENH_CUC!$U$18))),"Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.",
    #   "Mức trung bình; tốt xấu tùy vận."))
    return ('Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('GIÀU', E.v(18,18,21))), X.ISNUMBER(lambda: X.SEARCH('SANG', E.v(18,18,21))), X.ISNUMBER(lambda: X.SEARCH('CÁT', E.v(18,18,21))), X.ISNUMBER(lambda: X.SEARCH('THỌ', E.v(18,18,21))))) else ('Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('NGHÈO', E.v(18,18,21))), X.ISNUMBER(lambda: X.SEARCH('HÈN', E.v(18,18,21))), X.ISNUMBER(lambda: X.SEARCH('HUNG', E.v(18,18,21))), X.ISNUMBER(lambda: X.SEARCH('YỂU', E.v(18,18,21))))) else 'Mức trung bình; tốt xấu tùy vận.'))

def f180(E, r, c):
    # Ô H116, H117, H118, H119
    # = "Theo dụng thần: làm việc, kinh doanh thuộc hành dụng/hỷ; tránh mạo hiểm trong vận kỵ thần."
    return 'Theo dụng thần: làm việc, kinh doanh thuộc hành dụng/hỷ; tránh mạo hiểm trong vận kỵ thần.'

def f181(E, r, c):
    # Ô I116
    # = IF(ISNUMBER(SEARCH("PP2 nghiêng",MENH_CUC!$D$18)),"Vừa – hai phương pháp cho kết quả khác nhau","Khá")&"; vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – hai phương pháp cho kết quả khác nhau' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('PP2 nghiêng', E.v(18,18,4)))) else 'Khá'),'; vượng suy: '),E.v(7,35,3))

def f182(E, r, c):
    # Ô C117
    # = MENH_CUC!$D$19
    return E.v(18,19,4)

def f183(E, r, c):
    # Ô D117
    # = MENH_CUC!$G$19
    return E.v(18,19,7)

def f184(E, r, c):
    # Ô F117
    # = IF(OR(ISNUMBER(SEARCH("GIÀU",MENH_CUC!$U$19)),ISNUMBER(SEARCH("SANG",MENH_CUC!$U$19)),ISNUMBER(SEARCH("CÁT",MENH_CUC!$U$19)),ISNUMBER(SEARCH("THỌ",MEN
    #   H_CUC!$U$19))),"Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.",IF(OR(ISNUMBER(SEARCH("NGHÈO",MENH_CUC!$U$19)),ISNUMBER(SEARCH("HÈN",MENH_CUC!
    #   $U$19)),ISNUMBER(SEARCH("HUNG",MENH_CUC!$U$19)),ISNUMBER(SEARCH("YỂU",MENH_CUC!$U$19))),"Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.",
    #   "Mức trung bình; tốt xấu tùy vận."))
    return ('Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('GIÀU', E.v(18,19,21))), X.ISNUMBER(lambda: X.SEARCH('SANG', E.v(18,19,21))), X.ISNUMBER(lambda: X.SEARCH('CÁT', E.v(18,19,21))), X.ISNUMBER(lambda: X.SEARCH('THỌ', E.v(18,19,21))))) else ('Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('NGHÈO', E.v(18,19,21))), X.ISNUMBER(lambda: X.SEARCH('HÈN', E.v(18,19,21))), X.ISNUMBER(lambda: X.SEARCH('HUNG', E.v(18,19,21))), X.ISNUMBER(lambda: X.SEARCH('YỂU', E.v(18,19,21))))) else 'Mức trung bình; tốt xấu tùy vận.'))

def f185(E, r, c):
    # Ô I117
    # = IF(ISNUMBER(SEARCH("PP2 nghiêng",MENH_CUC!$D$19)),"Vừa – hai phương pháp cho kết quả khác nhau","Khá")&"; vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – hai phương pháp cho kết quả khác nhau' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('PP2 nghiêng', E.v(18,19,4)))) else 'Khá'),'; vượng suy: '),E.v(7,35,3))

def f186(E, r, c):
    # Ô C118
    # = MENH_CUC!$D$20
    return E.v(18,20,4)

def f187(E, r, c):
    # Ô D118
    # = MENH_CUC!$G$20
    return E.v(18,20,7)

def f188(E, r, c):
    # Ô F118
    # = IF(OR(ISNUMBER(SEARCH("GIÀU",MENH_CUC!$U$20)),ISNUMBER(SEARCH("SANG",MENH_CUC!$U$20)),ISNUMBER(SEARCH("CÁT",MENH_CUC!$U$20)),ISNUMBER(SEARCH("THỌ",MEN
    #   H_CUC!$U$20))),"Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.",IF(OR(ISNUMBER(SEARCH("NGHÈO",MENH_CUC!$U$20)),ISNUMBER(SEARCH("HÈN",MENH_CUC!
    #   $U$20)),ISNUMBER(SEARCH("HUNG",MENH_CUC!$U$20)),ISNUMBER(SEARCH("YỂU",MENH_CUC!$U$20))),"Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.",
    #   "Mức trung bình; tốt xấu tùy vận."))
    return ('Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('GIÀU', E.v(18,20,21))), X.ISNUMBER(lambda: X.SEARCH('SANG', E.v(18,20,21))), X.ISNUMBER(lambda: X.SEARCH('CÁT', E.v(18,20,21))), X.ISNUMBER(lambda: X.SEARCH('THỌ', E.v(18,20,21))))) else ('Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('NGHÈO', E.v(18,20,21))), X.ISNUMBER(lambda: X.SEARCH('HÈN', E.v(18,20,21))), X.ISNUMBER(lambda: X.SEARCH('HUNG', E.v(18,20,21))), X.ISNUMBER(lambda: X.SEARCH('YỂU', E.v(18,20,21))))) else 'Mức trung bình; tốt xấu tùy vận.'))

def f189(E, r, c):
    # Ô I118
    # = IF(ISNUMBER(SEARCH("PP2 nghiêng",MENH_CUC!$D$20)),"Vừa – hai phương pháp cho kết quả khác nhau","Khá")&"; vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – hai phương pháp cho kết quả khác nhau' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('PP2 nghiêng', E.v(18,20,4)))) else 'Khá'),'; vượng suy: '),E.v(7,35,3))

def f190(E, r, c):
    # Ô C119
    # = MENH_CUC!$D$21
    return E.v(18,21,4)

def f191(E, r, c):
    # Ô D119
    # = MENH_CUC!$G$21
    return E.v(18,21,7)

def f192(E, r, c):
    # Ô F119
    # = IF(OR(ISNUMBER(SEARCH("GIÀU",MENH_CUC!$C$126)),ISNUMBER(SEARCH("SANG",MENH_CUC!$C$126)),ISNUMBER(SEARCH("CÁT",MENH_CUC!$C$126)),ISNUMBER(SEARCH("THỌ",
    #   MENH_CUC!$C$126))),"Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.",IF(OR(ISNUMBER(SEARCH("NGHÈO",MENH_CUC!$C$126)),ISNUMBER(SEARCH("HÈN",MENH
    #   _CUC!$C$126)),ISNUMBER(SEARCH("HUNG",MENH_CUC!$C$126)),ISNUMBER(SEARCH("YỂU",MENH_CUC!$C$126))),"Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vậ
    #   n tốt.","Mức trung bình; tốt xấu tùy vận."))
    return ('Nền tảng thuận ở mặt này; vận tốt đến thì phát huy rõ.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('GIÀU', E.v(18,126,3))), X.ISNUMBER(lambda: X.SEARCH('SANG', E.v(18,126,3))), X.ISNUMBER(lambda: X.SEARCH('CÁT', E.v(18,126,3))), X.ISNUMBER(lambda: X.SEARCH('THỌ', E.v(18,126,3))))) else ('Nền tảng kém ở mặt này; phải nỗ lực nhiều, dựa vào vận tốt.' if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('NGHÈO', E.v(18,126,3))), X.ISNUMBER(lambda: X.SEARCH('HÈN', E.v(18,126,3))), X.ISNUMBER(lambda: X.SEARCH('HUNG', E.v(18,126,3))), X.ISNUMBER(lambda: X.SEARCH('YỂU', E.v(18,126,3))))) else 'Mức trung bình; tốt xấu tùy vận.'))

def f193(E, r, c):
    # Ô I119
    # = IF(ISNUMBER(SEARCH("PP2 nghiêng",MENH_CUC!$D$21)),"Vừa – hai phương pháp cho kết quả khác nhau","Khá")&"; vượng suy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – hai phương pháp cho kết quả khác nhau' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('PP2 nghiêng', E.v(18,21,4)))) else 'Khá'),'; vượng suy: '),E.v(7,35,3))

def f194(E, r, c):
    # Ô C120
    # = TAI_LOC!$E$17
    return E.v(19,17,5)

def f195(E, r, c):
    # Ô D120
    # = KET_LUAN!$C$86
    return E.v(3,86,3)

def f196(E, r, c):
    # Ô E120
    # = "Tài là hành do nhật chủ khắc (của cải mình chế ngự được). Giàu khi thân đủ mạnh để “gánh” tài và tài là hỷ dụng (P2 tr.520-530)."
    return 'Tài là hành do nhật chủ khắc (của cải mình chế ngự được). Giàu khi thân đủ mạnh để “gánh” tài và tài là hỷ dụng (P2 tr.520-530).'

def f197(E, r, c):
    # Ô F120
    # = IF(OR(LEFT(SU_NGHIEP!$J$6,5)="THUẬN",LEFT(SU_NGHIEP!$J$6,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SU_NGHIEP!$J$6,20))),"Tiền đến rồi đi, dễ hao; 
    #   cần kỷ luật chi tiêu, tránh đầu cơ.","Có khả năng kiếm và giữ tiền; nên đầu tư vào lĩnh vực hợp dụng thần."),IF(LEFT(SU_NGHIEP!$J$6,4)="BẤT ","Tiền đế
    #   n rồi đi, dễ hao; cần kỷ luật chi tiêu, tránh đầu cơ.",IF(OR(LEFT(SU_NGHIEP!$J$6,3)="TỐT",LEFT(SU_NGHIEP!$J$6,3)="Tốt",LEFT(SU_NGHIEP!$J$6,3)="Khá"),"
    #   Có khả năng kiếm và giữ tiền; nên đầu tư vào lĩnh vực hợp dụng thần.",IF(OR(LEFT(SU_NGHIEP!$J$6,3)="KÉM",LEFT(SU_NGHIEP!$J$6,3)="Kém",LEFT(SU_NGHIEP!$
    #   J$6,6)="Rất k"),"Tiền đến rồi đi, dễ hao; cần kỷ luật chi tiêu, tránh đầu cơ.","Thu nhập ổn định ở mức vừa; tăng khi gặp vận tốt."))))
    return (('Tiền đến rồi đi, dễ hao; cần kỷ luật chi tiêu, tránh đầu cơ.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(21,6,10), 20.0)))) else 'Có khả năng kiếm và giữ tiền; nên đầu tư vào lĩnh vực hợp dụng thần.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(21,6,10), 5.0),'NGHIÊ'))) else ('Tiền đến rồi đi, dễ hao; cần kỷ luật chi tiêu, tránh đầu cơ.' if X.truthy(X.eq(X.LEFT(E.v(21,6,10), 4.0),'BẤT ')) else ('Có khả năng kiếm và giữ tiền; nên đầu tư vào lĩnh vực hợp dụng thần.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(21,6,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(21,6,10), 3.0),'Khá'))) else ('Tiền đến rồi đi, dễ hao; cần kỷ luật chi tiêu, tránh đầu cơ.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(21,6,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(21,6,10), 6.0),'Rất k'))) else 'Thu nhập ổn định ở mức vừa; tăng khi gặp vận tốt.'))))

def f198(E, r, c):
    # Ô G120
    # = "Dễ phát tài: "&IF(TAI_LOC!$C$37="","—",TAI_LOC!$C$37)
    return X.cat('Dễ phát tài: ',('—' if X.truthy(X.eq(E.v(19,37,3),'')) else E.v(19,37,3)))

def f199(E, r, c):
    # Ô H120
    # = "Giữ tiền trong vận kỵ; mở rộng làm ăn trong vận dụng/hỷ; nguồn tiền hợp: xem TAI_LOC."
    return 'Giữ tiền trong vận kỵ; mở rộng làm ăn trong vận dụng/hỷ; nguồn tiền hợp: xem TAI_LOC.'

def f200(E, r, c):
    # Ô I120, I123, I182, I183, I184, I185, I186
    # = IF(SU_NGHIEP!$M$6+SU_NGHIEP!$N$6=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(AB
    #   S(SU_NGHIEP!$M$6-SU_NGHIEP!$N$6)>=3,"Khá",IF(ABS(SU_NGHIEP!$M$6-SU_NGHIEP!$N$6)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SU_NGHIEP!$M
    #   $6&"/"&SU_NGHIEP!$N$6)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(21,6,13),E.v(21,6,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,6,13),E.v(21,6,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,6,13),E.v(21,6,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(21,6,13)),'/'),E.v(21,6,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f201(E, r, c):
    # Ô C121
    # = TO_CHAT!$C$8
    return E.v(20,8,3)

def f202(E, r, c):
    # Ô D121
    # = TO_CHAT!$C$9
    return E.v(20,9,3)

def f203(E, r, c):
    # Ô E121
    # = "Bảng P2 tr.578-581 tra theo hành nhật chủ và mức vượng suy: mỗi hành có ưu điểm khi vừa phải và nhược điểm khi thái quá hoặc bất cập."
    return 'Bảng P2 tr.578-581 tra theo hành nhật chủ và mức vượng suy: mỗi hành có ưu điểm khi vừa phải và nhược điểm khi thái quá hoặc bất cập.'

def f204(E, r, c):
    # Ô F121
    # = "Đây là thiên hướng bẩm sinh; môi trường, giáo dục làm thay đổi được (P1 tr.24)."
    return 'Đây là thiên hướng bẩm sinh; môi trường, giáo dục làm thay đổi được (P1 tr.24).'

def f205(E, r, c):
    # Ô G121
    # = "Bộc lộ suốt đời, rõ hơn trong vận làm hành nhật chủ thái quá hoặc bất cập."
    return 'Bộc lộ suốt đời, rõ hơn trong vận làm hành nhật chủ thái quá hoặc bất cập.'

def f206(E, r, c):
    # Ô H121
    # = "Phát huy điểm mạnh; chọn nghề hợp dụng thần: "&LEFT(TO_CHAT!$C$17,IFERROR(FIND(CHAR(10),TO_CHAT!$C$17)-1,LEN(TO_CHAT!$C$17)))
    return X.cat('Phát huy điểm mạnh; chọn nghề hợp dụng thần: ',X.LEFT(E.v(20,17,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(20,17,3)),1.0), lambda: X.LEN(E.v(20,17,3)))))

def f207(E, r, c):
    # Ô C122
    # = LEFT(TO_CHAT!$C$25,IFERROR(FIND(CHAR(10),TO_CHAT!$C$25)-1,LEN(TO_CHAT!$C$25)))
    return X.LEFT(E.v(20,25,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(20,25,3)),1.0), lambda: X.LEN(E.v(20,25,3))))

def f208(E, r, c):
    # Ô D122
    # = TO_CHAT!$C$25
    return E.v(20,25,3)

def f209(E, r, c):
    # Ô E122
    # = "Mỗi câu là một dấu hiệu trong sách (P2 tr.573-577); hình dáng theo hành thịnh/suy (P1 tr.25-26). Hành thịnh là nét trội, hành suy là nét chìm (P1 tr.
    #   24)."
    return 'Mỗi câu là một dấu hiệu trong sách (P2 tr.573-577); hình dáng theo hành thịnh/suy (P1 tr.25-26). Hành thịnh là nét trội, hành suy là nét chìm (P1 tr.24).'

def f210(E, r, c):
    # Ô F122
    # = "Dùng để hiểu mình: nét trội là thế mạnh, nét cần lưu ý là chỗ nên tự điều chỉnh. Hình dáng chỉ là gợi ý, sách dặn phải kết hợp diện tướng."
    return 'Dùng để hiểu mình: nét trội là thế mạnh, nét cần lưu ý là chỗ nên tự điều chỉnh. Hình dáng chỉ là gợi ý, sách dặn phải kết hợp diện tướng.'

def f211(E, r, c):
    # Ô G122
    # = "Rõ trong giao tiếp hằng ngày; nét xấu dễ lộ khi gặp vận kỵ thần."
    return 'Rõ trong giao tiếp hằng ngày; nét xấu dễ lộ khi gặp vận kỵ thần.'

def f212(E, r, c):
    # Ô H122
    # = "Tự rèn ở những điểm “cần lưu ý”; môi trường, nghề hợp dụng thần giúp cân bằng tính cách."
    return 'Tự rèn ở những điểm “cần lưu ý”; môi trường, nghề hợp dụng thần giúp cân bằng tính cách.'

def f213(E, r, c):
    # Ô I122
    # = "Vừa (nhiều câu sách, mỗi câu một dấu hiệu)"
    return 'Vừa (nhiều câu sách, mỗi câu một dấu hiệu)'

def f214(E, r, c):
    # Ô C123
    # = LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))
    return X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10))))

def f215(E, r, c):
    # Ô D123
    # = "Thuận: "&LEFT(SU_NGHIEP!$C$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$C$6)-1,LEN(SU_NGHIEP!$C$6)))&" … Bất lợi: "&LEFT(SU_NGHIEP!$F$6,IFERROR(FIND(CHAR(10),S
    #   U_NGHIEP!$F$6)-1,LEN(SU_NGHIEP!$F$6)))&" … (đủ ở SU_NGHIEP)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(21,6,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,3)),1.0), lambda: X.LEN(E.v(21,6,3))))),' … Bất lợi: '),X.LEFT(E.v(21,6,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,6)),1.0), lambda: X.LEN(E.v(21,6,6))))),' … (đủ ở SU_NGHIEP)')

def f216(E, r, c):
    # Ô E123
    # = "Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: tài là của cải; chủ sự là sao tài."
    return 'Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: tài là của cải; chủ sự là sao tài.'

def f217(E, r, c):
    # Ô F123
    # = IF(OR(LEFT(SU_NGHIEP!$J$6,5)="THUẬN",LEFT(SU_NGHIEP!$J$6,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SU_NGHIEP!$J$6,20))),"Mặt này trở ngại; cần bền
    #    bỉ, chọn hướng đi hợp dụng thần.","Mặt này thuận; nỗ lực dễ có kết quả."),IF(LEFT(SU_NGHIEP!$J$6,4)="BẤT ","Mặt này trở ngại; cần bền bỉ, chọn hướng 
    #   đi hợp dụng thần.",IF(OR(LEFT(SU_NGHIEP!$J$6,3)="TỐT",LEFT(SU_NGHIEP!$J$6,3)="Tốt",LEFT(SU_NGHIEP!$J$6,3)="Khá"),"Mặt này thuận; nỗ lực dễ có kết quả.
    #   ",IF(OR(LEFT(SU_NGHIEP!$J$6,3)="KÉM",LEFT(SU_NGHIEP!$J$6,3)="Kém",LEFT(SU_NGHIEP!$J$6,6)="Rất k"),"Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụn
    #   g thần.","Mức trung bình; kết quả tùy vận."))))
    return (('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(21,6,10), 20.0)))) else 'Mặt này thuận; nỗ lực dễ có kết quả.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(21,6,10), 5.0),'NGHIÊ'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.eq(X.LEFT(E.v(21,6,10), 4.0),'BẤT ')) else ('Mặt này thuận; nỗ lực dễ có kết quả.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(21,6,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(21,6,10), 3.0),'Khá'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,6,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(21,6,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(21,6,10), 6.0),'Rất k'))) else 'Mức trung bình; kết quả tùy vận.'))))

def f218(E, r, c):
    # Ô H123, H124, H125, H126
    # = "Đẩy mạnh trong vận dụng/hỷ; giữ vững, tránh mạo hiểm trong vận kỵ."
    return 'Đẩy mạnh trong vận dụng/hỷ; giữ vững, tránh mạo hiểm trong vận kỵ.'

def f219(E, r, c):
    # Ô C124
    # = LEFT(SU_NGHIEP!$J$7,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$7)-1,LEN(SU_NGHIEP!$J$7)))
    return X.LEFT(E.v(21,7,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0), lambda: X.LEN(E.v(21,7,10))))

def f220(E, r, c):
    # Ô D124
    # = "Thuận: "&LEFT(SU_NGHIEP!$C$7,IFERROR(FIND(CHAR(10),SU_NGHIEP!$C$7)-1,LEN(SU_NGHIEP!$C$7)))&" … Bất lợi: "&LEFT(SU_NGHIEP!$F$7,IFERROR(FIND(CHAR(10),S
    #   U_NGHIEP!$F$7)-1,LEN(SU_NGHIEP!$F$7)))&" … (đủ ở SU_NGHIEP)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(21,7,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,3)),1.0), lambda: X.LEN(E.v(21,7,3))))),' … Bất lợi: '),X.LEFT(E.v(21,7,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,6)),1.0), lambda: X.LEN(E.v(21,7,6))))),' … (đủ ở SU_NGHIEP)')

def f221(E, r, c):
    # Ô E124
    # = "Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: quan sát là chức quyền, kỷ luật, sự quả
    #   n lý."
    return 'Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: quan sát là chức quyền, kỷ luật, sự quản lý.'

def f222(E, r, c):
    # Ô F124
    # = IF(OR(LEFT(SU_NGHIEP!$J$7,5)="THUẬN",LEFT(SU_NGHIEP!$J$7,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SU_NGHIEP!$J$7,20))),"Mặt này trở ngại; cần bền
    #    bỉ, chọn hướng đi hợp dụng thần.","Mặt này thuận; nỗ lực dễ có kết quả."),IF(LEFT(SU_NGHIEP!$J$7,4)="BẤT ","Mặt này trở ngại; cần bền bỉ, chọn hướng 
    #   đi hợp dụng thần.",IF(OR(LEFT(SU_NGHIEP!$J$7,3)="TỐT",LEFT(SU_NGHIEP!$J$7,3)="Tốt",LEFT(SU_NGHIEP!$J$7,3)="Khá"),"Mặt này thuận; nỗ lực dễ có kết quả.
    #   ",IF(OR(LEFT(SU_NGHIEP!$J$7,3)="KÉM",LEFT(SU_NGHIEP!$J$7,3)="Kém",LEFT(SU_NGHIEP!$J$7,6)="Rất k"),"Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụn
    #   g thần.","Mức trung bình; kết quả tùy vận."))))
    return (('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(21,7,10), 20.0)))) else 'Mặt này thuận; nỗ lực dễ có kết quả.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,7,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(21,7,10), 5.0),'NGHIÊ'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.eq(X.LEFT(E.v(21,7,10), 4.0),'BẤT ')) else ('Mặt này thuận; nỗ lực dễ có kết quả.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,7,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(21,7,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(21,7,10), 3.0),'Khá'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,7,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(21,7,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(21,7,10), 6.0),'Rất k'))) else 'Mức trung bình; kết quả tùy vận.'))))

def f223(E, r, c):
    # Ô I124, I191
    # = IF(SU_NGHIEP!$M$7+SU_NGHIEP!$N$7=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(AB
    #   S(SU_NGHIEP!$M$7-SU_NGHIEP!$N$7)>=3,"Khá",IF(ABS(SU_NGHIEP!$M$7-SU_NGHIEP!$N$7)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SU_NGHIEP!$M
    #   $7&"/"&SU_NGHIEP!$N$7)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(21,7,13),E.v(21,7,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,7,13),E.v(21,7,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,7,13),E.v(21,7,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(21,7,13)),'/'),E.v(21,7,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f224(E, r, c):
    # Ô C125
    # = LEFT(SU_NGHIEP!$J$8,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$8)-1,LEN(SU_NGHIEP!$J$8)))
    return X.LEFT(E.v(21,8,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0), lambda: X.LEN(E.v(21,8,10))))

def f225(E, r, c):
    # Ô D125
    # = "Thuận: "&LEFT(SU_NGHIEP!$C$8,IFERROR(FIND(CHAR(10),SU_NGHIEP!$C$8)-1,LEN(SU_NGHIEP!$C$8)))&" … Bất lợi: "&LEFT(SU_NGHIEP!$F$8,IFERROR(FIND(CHAR(10),S
    #   U_NGHIEP!$F$8)-1,LEN(SU_NGHIEP!$F$8)))&" … (đủ ở SU_NGHIEP)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(21,8,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,3)),1.0), lambda: X.LEN(E.v(21,8,3))))),' … Bất lợi: '),X.LEFT(E.v(21,8,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,6)),1.0), lambda: X.LEN(E.v(21,8,6))))),' … (đủ ở SU_NGHIEP)')

def f226(E, r, c):
    # Ô E125
    # = "Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: ấn là học vấn, bằng cấp; quan sát là th
    #   i đỗ."
    return 'Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: ấn là học vấn, bằng cấp; quan sát là thi đỗ.'

def f227(E, r, c):
    # Ô F125
    # = IF(OR(LEFT(SU_NGHIEP!$J$8,5)="THUẬN",LEFT(SU_NGHIEP!$J$8,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SU_NGHIEP!$J$8,20))),"Mặt này trở ngại; cần bền
    #    bỉ, chọn hướng đi hợp dụng thần.","Mặt này thuận; nỗ lực dễ có kết quả."),IF(LEFT(SU_NGHIEP!$J$8,4)="BẤT ","Mặt này trở ngại; cần bền bỉ, chọn hướng 
    #   đi hợp dụng thần.",IF(OR(LEFT(SU_NGHIEP!$J$8,3)="TỐT",LEFT(SU_NGHIEP!$J$8,3)="Tốt",LEFT(SU_NGHIEP!$J$8,3)="Khá"),"Mặt này thuận; nỗ lực dễ có kết quả.
    #   ",IF(OR(LEFT(SU_NGHIEP!$J$8,3)="KÉM",LEFT(SU_NGHIEP!$J$8,3)="Kém",LEFT(SU_NGHIEP!$J$8,6)="Rất k"),"Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụn
    #   g thần.","Mức trung bình; kết quả tùy vận."))))
    return (('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(21,8,10), 20.0)))) else 'Mặt này thuận; nỗ lực dễ có kết quả.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,8,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(21,8,10), 5.0),'NGHIÊ'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.eq(X.LEFT(E.v(21,8,10), 4.0),'BẤT ')) else ('Mặt này thuận; nỗ lực dễ có kết quả.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,8,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(21,8,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(21,8,10), 3.0),'Khá'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,8,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(21,8,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(21,8,10), 6.0),'Rất k'))) else 'Mức trung bình; kết quả tùy vận.'))))

def f228(E, r, c):
    # Ô I125
    # = IF(SU_NGHIEP!$M$8+SU_NGHIEP!$N$8=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(AB
    #   S(SU_NGHIEP!$M$8-SU_NGHIEP!$N$8)>=3,"Khá",IF(ABS(SU_NGHIEP!$M$8-SU_NGHIEP!$N$8)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SU_NGHIEP!$M
    #   $8&"/"&SU_NGHIEP!$N$8)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(21,8,13),E.v(21,8,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,8,13),E.v(21,8,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,8,13),E.v(21,8,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(21,8,13)),'/'),E.v(21,8,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f229(E, r, c):
    # Ô C126
    # = LEFT(SU_NGHIEP!$J$9,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$9)-1,LEN(SU_NGHIEP!$J$9)))
    return X.LEFT(E.v(21,9,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0), lambda: X.LEN(E.v(21,9,10))))

def f230(E, r, c):
    # Ô D126
    # = "Thuận: "&LEFT(SU_NGHIEP!$C$9,IFERROR(FIND(CHAR(10),SU_NGHIEP!$C$9)-1,LEN(SU_NGHIEP!$C$9)))&" … Bất lợi: "&LEFT(SU_NGHIEP!$F$9,IFERROR(FIND(CHAR(10),S
    #   U_NGHIEP!$F$9)-1,LEN(SU_NGHIEP!$F$9)))&" … (đủ ở SU_NGHIEP)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(21,9,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,3)),1.0), lambda: X.LEN(E.v(21,9,3))))),' … Bất lợi: '),X.LEFT(E.v(21,9,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,6)),1.0), lambda: X.LEN(E.v(21,9,6))))),' … (đủ ở SU_NGHIEP)')

def f231(E, r, c):
    # Ô E126
    # = "Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: quan, ấn, thực thương là danh tiếng, tà
    #   i năng được biết đến."
    return 'Kết luận dựa vào vai trò của sao chủ sự theo dụng thần (giả định G17), cộng các câu sách đủ điều kiện. Ở đây: quan, ấn, thực thương là danh tiếng, tài năng được biết đến.'

def f232(E, r, c):
    # Ô F126
    # = IF(OR(LEFT(SU_NGHIEP!$J$9,5)="THUẬN",LEFT(SU_NGHIEP!$J$9,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SU_NGHIEP!$J$9,20))),"Mặt này trở ngại; cần bền
    #    bỉ, chọn hướng đi hợp dụng thần.","Mặt này thuận; nỗ lực dễ có kết quả."),IF(LEFT(SU_NGHIEP!$J$9,4)="BẤT ","Mặt này trở ngại; cần bền bỉ, chọn hướng 
    #   đi hợp dụng thần.",IF(OR(LEFT(SU_NGHIEP!$J$9,3)="TỐT",LEFT(SU_NGHIEP!$J$9,3)="Tốt",LEFT(SU_NGHIEP!$J$9,3)="Khá"),"Mặt này thuận; nỗ lực dễ có kết quả.
    #   ",IF(OR(LEFT(SU_NGHIEP!$J$9,3)="KÉM",LEFT(SU_NGHIEP!$J$9,3)="Kém",LEFT(SU_NGHIEP!$J$9,6)="Rất k"),"Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụn
    #   g thần.","Mức trung bình; kết quả tùy vận."))))
    return (('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(21,9,10), 20.0)))) else 'Mặt này thuận; nỗ lực dễ có kết quả.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,9,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(21,9,10), 5.0),'NGHIÊ'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.eq(X.LEFT(E.v(21,9,10), 4.0),'BẤT ')) else ('Mặt này thuận; nỗ lực dễ có kết quả.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,9,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(21,9,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(21,9,10), 3.0),'Khá'))) else ('Mặt này trở ngại; cần bền bỉ, chọn hướng đi hợp dụng thần.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(21,9,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(21,9,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(21,9,10), 6.0),'Rất k'))) else 'Mức trung bình; kết quả tùy vận.'))))

def f233(E, r, c):
    # Ô I126
    # = IF(SU_NGHIEP!$M$9+SU_NGHIEP!$N$9=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(AB
    #   S(SU_NGHIEP!$M$9-SU_NGHIEP!$N$9)>=3,"Khá",IF(ABS(SU_NGHIEP!$M$9-SU_NGHIEP!$N$9)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SU_NGHIEP!$M
    #   $9&"/"&SU_NGHIEP!$N$9)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(21,9,13),E.v(21,9,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,9,13),E.v(21,9,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(21,9,13),E.v(21,9,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(21,9,13)),'/'),E.v(21,9,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f234(E, r, c):
    # Ô C127
    # = LEFT(SUC_KHOE!$J$6,IFERROR(FIND(CHAR(10),SUC_KHOE!$J$6)-1,LEN(SUC_KHOE!$J$6)))
    return X.LEFT(E.v(22,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0), lambda: X.LEN(E.v(22,6,10))))

def f235(E, r, c):
    # Ô D127
    # = "Thuận: "&LEFT(SUC_KHOE!$C$6,IFERROR(FIND(CHAR(10),SUC_KHOE!$C$6)-1,LEN(SUC_KHOE!$C$6)))&" … Bất lợi: "&LEFT(SUC_KHOE!$F$6,IFERROR(FIND(CHAR(10),SUC_K
    #   HOE!$F$6)-1,LEN(SUC_KHOE!$F$6)))&" … (đủ ở SUC_KHOE)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(22,6,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,3)),1.0), lambda: X.LEN(E.v(22,6,3))))),' … Bất lợi: '),X.LEFT(E.v(22,6,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,6)),1.0), lambda: X.LEN(E.v(22,6,6))))),' … (đủ ở SUC_KHOE)')

def f236(E, r, c):
    # Ô E127, E128, E129
    # = "Ngũ hành ứng với tạng phủ (Mộc – gan mật, Hỏa – tim máu mắt, Thổ – tì vị, Kim – phổi đại tràng, Thủy – thận bàng quang); hành bị khắc hoặc ở đất Tử/T
    #   uyệt là chỗ yếu (P2 tr.557-562)."
    return 'Ngũ hành ứng với tạng phủ (Mộc – gan mật, Hỏa – tim máu mắt, Thổ – tì vị, Kim – phổi đại tràng, Thủy – thận bàng quang); hành bị khắc hoặc ở đất Tử/Tuyệt là chỗ yếu (P2 tr.557-562).'

def f237(E, r, c):
    # Ô F127
    # = IF(OR(LEFT(SUC_KHOE!$J$6,5)="THUẬN",LEFT(SUC_KHOE!$J$6,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SUC_KHOE!$J$6,20))),"Có điểm yếu cần theo dõi (xe
    #   m tạng phủ ở SUC_KHOE mục B, H).","Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ."),IF(LEFT(SUC_KHOE!$J$6,4)="BẤT ","Có điểm yếu cần theo dõi (xem tạ
    #   ng phủ ở SUC_KHOE mục B, H).",IF(OR(LEFT(SUC_KHOE!$J$6,3)="TỐT",LEFT(SUC_KHOE!$J$6,3)="Tốt",LEFT(SUC_KHOE!$J$6,3)="Khá"),"Nền sức khỏe tốt; vẫn nên gi
    #   ữ nếp sống điều độ.",IF(OR(LEFT(SUC_KHOE!$J$6,3)="KÉM",LEFT(SUC_KHOE!$J$6,3)="Kém",LEFT(SUC_KHOE!$J$6,6)="Rất k"),"Có điểm yếu cần theo dõi (xem tạng 
    #   phủ ở SUC_KHOE mục B, H).","Bình thường; chú ý trong năm được báo."))))
    return (('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(22,6,10), 20.0)))) else 'Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,6,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(22,6,10), 5.0),'NGHIÊ'))) else ('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.eq(X.LEFT(E.v(22,6,10), 4.0),'BẤT ')) else ('Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,6,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(22,6,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(22,6,10), 3.0),'Khá'))) else ('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,6,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(22,6,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(22,6,10), 6.0),'Rất k'))) else 'Bình thường; chú ý trong năm được báo.'))))

def f238(E, r, c):
    # Ô G127, G128, G129, G194
    # = SUC_KHOE!$C$157
    return E.v(22,157,3)

def f239(E, r, c):
    # Ô H127, H128, H129
    # = "Khám sức khỏe định kỳ, nhất là năm mức ĐỀ PHÒNG/NẶNG; ăn uống, sinh hoạt bổ hành dụng thần. Không thay chẩn đoán y khoa."
    return 'Khám sức khỏe định kỳ, nhất là năm mức ĐỀ PHÒNG/NẶNG; ăn uống, sinh hoạt bổ hành dụng thần. Không thay chẩn đoán y khoa.'

def f240(E, r, c):
    # Ô I127, I194
    # = IF(SUC_KHOE!$M$6+SUC_KHOE!$N$6=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(ABS(
    #   SUC_KHOE!$M$6-SUC_KHOE!$N$6)>=3,"Khá",IF(ABS(SUC_KHOE!$M$6-SUC_KHOE!$N$6)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SUC_KHOE!$M$6&"/"&
    #   SUC_KHOE!$N$6)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(22,6,13),E.v(22,6,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(22,6,13),E.v(22,6,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(22,6,13),E.v(22,6,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(22,6,13)),'/'),E.v(22,6,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f241(E, r, c):
    # Ô C128
    # = LEFT(SUC_KHOE!$J$7,IFERROR(FIND(CHAR(10),SUC_KHOE!$J$7)-1,LEN(SUC_KHOE!$J$7)))
    return X.LEFT(E.v(22,7,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0), lambda: X.LEN(E.v(22,7,10))))

def f242(E, r, c):
    # Ô D128
    # = "Thuận: "&LEFT(SUC_KHOE!$C$7,IFERROR(FIND(CHAR(10),SUC_KHOE!$C$7)-1,LEN(SUC_KHOE!$C$7)))&" … Bất lợi: "&LEFT(SUC_KHOE!$F$7,IFERROR(FIND(CHAR(10),SUC_K
    #   HOE!$F$7)-1,LEN(SUC_KHOE!$F$7)))&" … (đủ ở SUC_KHOE)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(22,7,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,3)),1.0), lambda: X.LEN(E.v(22,7,3))))),' … Bất lợi: '),X.LEFT(E.v(22,7,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,6)),1.0), lambda: X.LEN(E.v(22,7,6))))),' … (đủ ở SUC_KHOE)')

def f243(E, r, c):
    # Ô F128
    # = IF(OR(LEFT(SUC_KHOE!$J$7,5)="THUẬN",LEFT(SUC_KHOE!$J$7,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SUC_KHOE!$J$7,20))),"Có điểm yếu cần theo dõi (xe
    #   m tạng phủ ở SUC_KHOE mục B, H).","Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ."),IF(LEFT(SUC_KHOE!$J$7,4)="BẤT ","Có điểm yếu cần theo dõi (xem tạ
    #   ng phủ ở SUC_KHOE mục B, H).",IF(OR(LEFT(SUC_KHOE!$J$7,3)="TỐT",LEFT(SUC_KHOE!$J$7,3)="Tốt",LEFT(SUC_KHOE!$J$7,3)="Khá"),"Nền sức khỏe tốt; vẫn nên gi
    #   ữ nếp sống điều độ.",IF(OR(LEFT(SUC_KHOE!$J$7,3)="KÉM",LEFT(SUC_KHOE!$J$7,3)="Kém",LEFT(SUC_KHOE!$J$7,6)="Rất k"),"Có điểm yếu cần theo dõi (xem tạng 
    #   phủ ở SUC_KHOE mục B, H).","Bình thường; chú ý trong năm được báo."))))
    return (('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(22,7,10), 20.0)))) else 'Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,7,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(22,7,10), 5.0),'NGHIÊ'))) else ('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.eq(X.LEFT(E.v(22,7,10), 4.0),'BẤT ')) else ('Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,7,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(22,7,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(22,7,10), 3.0),'Khá'))) else ('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,7,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(22,7,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(22,7,10), 6.0),'Rất k'))) else 'Bình thường; chú ý trong năm được báo.'))))

def f244(E, r, c):
    # Ô I128
    # = IF(SUC_KHOE!$M$7+SUC_KHOE!$N$7=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(ABS(
    #   SUC_KHOE!$M$7-SUC_KHOE!$N$7)>=3,"Khá",IF(ABS(SUC_KHOE!$M$7-SUC_KHOE!$N$7)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SUC_KHOE!$M$7&"/"&
    #   SUC_KHOE!$N$7)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(22,7,13),E.v(22,7,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(22,7,13),E.v(22,7,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(22,7,13),E.v(22,7,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(22,7,13)),'/'),E.v(22,7,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f245(E, r, c):
    # Ô C129
    # = LEFT(SUC_KHOE!$J$8,IFERROR(FIND(CHAR(10),SUC_KHOE!$J$8)-1,LEN(SUC_KHOE!$J$8)))
    return X.LEFT(E.v(22,8,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0), lambda: X.LEN(E.v(22,8,10))))

def f246(E, r, c):
    # Ô D129
    # = "Thuận: "&LEFT(SUC_KHOE!$C$8,IFERROR(FIND(CHAR(10),SUC_KHOE!$C$8)-1,LEN(SUC_KHOE!$C$8)))&" … Bất lợi: "&LEFT(SUC_KHOE!$F$8,IFERROR(FIND(CHAR(10),SUC_K
    #   HOE!$F$8)-1,LEN(SUC_KHOE!$F$8)))&" … (đủ ở SUC_KHOE)"
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',X.LEFT(E.v(22,8,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,3)),1.0), lambda: X.LEN(E.v(22,8,3))))),' … Bất lợi: '),X.LEFT(E.v(22,8,6), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,6)),1.0), lambda: X.LEN(E.v(22,8,6))))),' … (đủ ở SUC_KHOE)')

def f247(E, r, c):
    # Ô F129
    # = IF(OR(LEFT(SUC_KHOE!$J$8,5)="THUẬN",LEFT(SUC_KHOE!$J$8,5)="NGHIÊ"),IF(ISNUMBER(SEARCH("BẤT LỢI",LEFT(SUC_KHOE!$J$8,20))),"Có điểm yếu cần theo dõi (xe
    #   m tạng phủ ở SUC_KHOE mục B, H).","Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ."),IF(LEFT(SUC_KHOE!$J$8,4)="BẤT ","Có điểm yếu cần theo dõi (xem tạ
    #   ng phủ ở SUC_KHOE mục B, H).",IF(OR(LEFT(SUC_KHOE!$J$8,3)="TỐT",LEFT(SUC_KHOE!$J$8,3)="Tốt",LEFT(SUC_KHOE!$J$8,3)="Khá"),"Nền sức khỏe tốt; vẫn nên gi
    #   ữ nếp sống điều độ.",IF(OR(LEFT(SUC_KHOE!$J$8,3)="KÉM",LEFT(SUC_KHOE!$J$8,3)="Kém",LEFT(SUC_KHOE!$J$8,6)="Rất k"),"Có điểm yếu cần theo dõi (xem tạng 
    #   phủ ở SUC_KHOE mục B, H).","Bình thường; chú ý trong năm được báo."))))
    return (('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', X.LEFT(E.v(22,8,10), 20.0)))) else 'Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ.') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,8,10), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(22,8,10), 5.0),'NGHIÊ'))) else ('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.eq(X.LEFT(E.v(22,8,10), 4.0),'BẤT ')) else ('Nền sức khỏe tốt; vẫn nên giữ nếp sống điều độ.' if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,8,10), 3.0),'TỐT'), X.eq(X.LEFT(E.v(22,8,10), 3.0),'Tốt'), X.eq(X.LEFT(E.v(22,8,10), 3.0),'Khá'))) else ('Có điểm yếu cần theo dõi (xem tạng phủ ở SUC_KHOE mục B, H).' if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,8,10), 3.0),'KÉM'), X.eq(X.LEFT(E.v(22,8,10), 3.0),'Kém'), X.eq(X.LEFT(E.v(22,8,10), 6.0),'Rất k'))) else 'Bình thường; chú ý trong năm được báo.'))))

def f248(E, r, c):
    # Ô I129
    # = IF(SUC_KHOE!$M$8+SUC_KHOE!$N$8=0,"Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)",IF(ABS(
    #   SUC_KHOE!$M$8-SUC_KHOE!$N$8)>=3,"Khá",IF(ABS(SUC_KHOE!$M$8-SUC_KHOE!$N$8)>=1,"Vừa","Thấp"))&" – chênh lệch dấu hiệu thuận/bất lợi "&SUC_KHOE!$M$8&"/"&
    #   SUC_KHOE!$N$8)&"; vượng suy (nền của kết luận) tin cậy: "&DOI_CHIEU!$C$35
    return X.cat(X.cat(('Vừa – không có câu sách nào đủ điều kiện, kết luận dựa vào yếu tố then chốt (vượng suy, vai trò sao chủ sự)' if X.truthy(X.eq(X.add(E.v(22,8,13),E.v(22,8,14)),0.0)) else X.cat(X.cat(X.cat(X.cat(('Khá' if X.truthy(X.ge(X.ABS(X.sub(E.v(22,8,13),E.v(22,8,14))),3.0)) else ('Vừa' if X.truthy(X.ge(X.ABS(X.sub(E.v(22,8,13),E.v(22,8,14))),1.0)) else 'Thấp')),' – chênh lệch dấu hiệu thuận/bất lợi '),E.v(22,8,13)),'/'),E.v(22,8,14))),'; vượng suy (nền của kết luận) tin cậy: '),E.v(7,35,3))

def f249(E, r, c):
    # Ô C130
    # = LEFT(KET_LUAN!$C$171,IFERROR(FIND(CHAR(10),KET_LUAN!$C$171)-1,LEN(KET_LUAN!$C$171)))
    return X.LEFT(E.v(3,171,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,171,3)),1.0), lambda: X.LEN(E.v(3,171,3))))

def f250(E, r, c):
    # Ô D130
    # = KET_LUAN!$C$171
    return E.v(3,171,3)

def f251(E, r, c):
    # Ô E130
    # = "Đại vận đổi mỗi 10 năm, lưu niên là từng năm. Vận/năm mang hành dụng, hỷ thần là tốt, mang kỵ thần là xấu; kết hợp vận và năm theo bảng P1 tr.133."
    return 'Đại vận đổi mỗi 10 năm, lưu niên là từng năm. Vận/năm mang hành dụng, hỷ thần là tốt, mang kỵ thần là xấu; kết hợp vận và năm theo bảng P1 tr.133.'

def f252(E, r, c):
    # Ô F130
    # = "Giai đoạn tốt nên chủ động phát triển; giai đoạn xấu nên giữ gìn, chuẩn bị."
    return 'Giai đoạn tốt nên chủ động phát triển; giai đoạn xấu nên giữ gìn, chuẩn bị.'

def f253(E, r, c):
    # Ô G130
    # = "Đại vận đang hành: "&LEFT(KET_LUAN!$C$163,IFERROR(FIND(CHAR(10),KET_LUAN!$C$163)-1,LEN(KET_LUAN!$C$163)))
    return X.cat('Đại vận đang hành: ',X.LEFT(E.v(3,163,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,163,3)),1.0), lambda: X.LEN(E.v(3,163,3)))))

def f254(E, r, c):
    # Ô H130
    # = "Lên kế hoạch lớn vào năm/vận tốt; năm xấu: tránh đầu tư mạo hiểm, chú ý sức khỏe, quan hệ."
    return 'Lên kế hoạch lớn vào năm/vận tốt; năm xấu: tránh đầu tư mạo hiểm, chú ý sức khỏe, quan hệ.'

def f255(E, r, c):
    # Ô C133
    # = "Người sinh với tứ trụ "&KET_LUAN!$C$15&" "&KET_LUAN!$D$15&" "&KET_LUAN!$E$15&" "&KET_LUAN!$F$15&", nhật chủ là "&C106&". Nói đơn giản: "&F106&" Vì th
    #   ế lá số cần dụng thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&" (hỷ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$C$37)&"), kỵ "&
    #   IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)&". Cách cục: "&C108&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Người sinh với tứ trụ ',E.v(3,15,3)),' '),E.v(3,15,4)),' '),E.v(3,15,5)),' '),E.v(3,15,6)),', nhật chủ là '),E.v(4,r-27,c)),'. Nói đơn giản: '),E.v(4,r-27,c+3)),' Vì thế lá số cần dụng thần '),(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),' (hỷ thần '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'), kỵ '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3))),'. Cách cục: '),E.v(4,r-25,c)),'.')

def f256(E, r, c):
    # Ô C134
    # = "Tố chất cốt lõi: "&C121&" Theo các câu của sách: "&C122&". "&TO_CHAT!$C$198
    return X.cat(X.cat(X.cat(X.cat(X.cat('Tố chất cốt lõi: ',E.v(4,r-13,c)),' Theo các câu của sách: '),E.v(4,r-12,c)),'. '),E.v(20,198,3))

def f257(E, r, c):
    # Ô C135
    # = "Tài vận "&LOWER(C123)&"; quan vận "&LOWER(C124)&"; học tập "&LOWER(C125)&"; danh tiếng "&LOWER(C126)&". Bậc giàu: "&C120&". "&F123&" Nghề hợp: "&LEFT
    #   (TO_CHAT!$C$17,IFERROR(FIND(CHAR(10),TO_CHAT!$C$17)-1,LEN(TO_CHAT!$C$17)))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tài vận ',X.LOWER(E.v(4,r-12,c))),'; quan vận '),X.LOWER(E.v(4,r-11,c))),'; học tập '),X.LOWER(E.v(4,r-10,c))),'; danh tiếng '),X.LOWER(E.v(4,r-9,c))),'. Bậc giàu: '),E.v(4,r-15,c)),'. '),E.v(4,r-12,c+3)),' Nghề hợp: '),X.LEFT(E.v(20,17,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(20,17,3)),1.0), lambda: X.LEN(E.v(20,17,3)))))

def f258(E, r, c):
    # Ô C136
    # = "Ông bà – tổ nghiệp: "&C110&". Cha: "&C111&". Mẹ: "&C112&". Anh chị em: "&C113&". Con cái: "&C114&". Hôn nhân: "&C115&" – "&F115
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Ông bà – tổ nghiệp: ',E.v(4,r-26,c)),'. Cha: '),E.v(4,r-25,c)),'. Mẹ: '),E.v(4,r-24,c)),'. Anh chị em: '),E.v(4,r-23,c)),'. Con cái: '),E.v(4,r-22,c)),'. Hôn nhân: '),E.v(4,r-21,c)),' – '),E.v(4,r-21,c+3))

def f259(E, r, c):
    # Ô C137
    # = "Sức khỏe chung: "&LOWER(C127)&"; bệnh tật: "&LOWER(C128)&"; thương tật – tai họa: "&LOWER(C129)&". "&SUC_KHOE!$C$157
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Sức khỏe chung: ',X.LOWER(E.v(4,r-10,c))),'; bệnh tật: '),X.LOWER(E.v(4,r-9,c))),'; thương tật – tai họa: '),X.LOWER(E.v(4,r-8,c))),'. '),E.v(22,157,3))

def f260(E, r, c):
    # Ô C138
    # = G130&". "&"Rõ nhất trong đại vận dụng/hỷ thần: "&IF(LUC_THAN!$Y$31="","(không có)",LUC_THAN!$Y$31)&"; kém đi trong đại vận kỵ thần: "&IF(LUC_THAN!$Z$3
    #   1="","(không có)",LUC_THAN!$Z$31)&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(4,r-8,c+4),'. '),'Rõ nhất trong đại vận dụng/hỷ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; kém đi trong đại vận kỵ thần: '),('(không có)' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))),'.')

def f261(E, r, c):
    # Ô C139
    # = H106&". "&H130
    return X.cat(X.cat(E.v(4,r-33,c+5),'. '),E.v(4,r-9,c+5))

def f262(E, r, c):
    # Ô C140
    # = "Vượng suy (nền của mọi kết luận): "&DOI_CHIEU!$C$35&". "&IF(DOI_CHIEU!$C$33=0,"PP1 và PP2 thống nhất.","PP1 và PP2 khác nhau; công cụ đang theo "&IF(
    #   DOI_CHIEU!$C$33=2,"PP2","PP1")&" – nếu sự việc thực tế khớp với phương pháp kia, đổi ở NHAP_LIEU C14 rồi xem lại.")&" Các câu sách là dấu hiệu, không 
    #   phải định mệnh; phần (c), (e) là cách nói dễ hiểu. Công cụ không kết luận năm mất và không thay chẩn đoán y khoa."
    return X.cat(X.cat(X.cat(X.cat('Vượng suy (nền của mọi kết luận): ',E.v(7,35,3)),'. '),('PP1 và PP2 thống nhất.' if X.truthy(X.eq(E.v(7,33,3),0.0)) else X.cat(X.cat('PP1 và PP2 khác nhau; công cụ đang theo ',('PP2' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP1')),' – nếu sự việc thực tế khớp với phương pháp kia, đổi ở NHAP_LIEU C14 rồi xem lại.'))),' Các câu sách là dấu hiệu, không phải định mệnh; phần (c), (e) là cách nói dễ hiểu. Công cụ không kết luận năm mất và không thay chẩn đoán y khoa.')

def f263(E, r, c):
    # Ô C144
    # = "PP1: "&KET_LUAN!$C$31&" | PP2: "&KET_LUAN!$D$31&" → đang dùng: "&DOI_CHIEU!$C$34
    return X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(3,31,3)),' | PP2: '),E.v(3,31,4)),' → đang dùng: '),E.v(7,34,3))

def f264(E, r, c):
    # Ô D144
    # = KET_LUAN!$F$31
    return E.v(3,31,6)

def f265(E, r, c):
    # Ô E144
    # = "PP1 (định tính) xét 4 điều kiện: được lệnh, đắc địa, được sinh, được trợ (P1 tr.70). PP2 (định lượng) cộng độ khí từng chữ theo tháng sinh rồi so phe
    #    mình với ngưỡng 40% (P2 tr.332-352). Hai cách có thể cho kết quả khác nhau."
    return 'PP1 (định tính) xét 4 điều kiện: được lệnh, đắc địa, được sinh, được trợ (P1 tr.70). PP2 (định lượng) cộng độ khí từng chữ theo tháng sinh rồi so phe mình với ngưỡng 40% (P2 tr.332-352). Hai cách có thể cho kết quả khác nhau.'

def f266(E, r, c):
    # Ô F144, F152, F153
    # = LUAN_GIAI!$F$106
    return E.v(4,106,6)

def f267(E, r, c):
    # Ô H144
    # = "Nếu hai phương pháp khác nhau, đối chiếu với các năm tốt/xấu đã qua để chọn ở NHAP_LIEU C14 (P1 tr.125, 201)."
    return 'Nếu hai phương pháp khác nhau, đối chiếu với các năm tốt/xấu đã qua để chọn ở NHAP_LIEU C14 (P1 tr.125, 201).'

def f268(E, r, c):
    # Ô J144, J145, J146, J147, J148, J149, J150, J151, J152, J153, J154, J155 … (64 ô)
    # = "(a) "&C144&CHAR(10)&"(b) "&LEFT(D144,600)&CHAR(10)&"   Lời thường: "&E144&CHAR(10)&"(c) "&F144&CHAR(10)&"(d) "&LEFT(G144,400)&CHAR(10)&"(e) "&LEFT(H1
    #   44,500)&CHAR(10)&"(f) Độ tin cậy: "&I144
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('(a) ',E.v(4,r,c-7)),X.CHAR(10.0)),'(b) '),X.LEFT(E.v(4,r,c-6), 600.0)),X.CHAR(10.0)),'   Lời thường: '),E.v(4,r,c-5)),X.CHAR(10.0)),'(c) '),E.v(4,r,c-4)),X.CHAR(10.0)),'(d) '),X.LEFT(E.v(4,r,c-3), 400.0)),X.CHAR(10.0)),'(e) '),X.LEFT(E.v(4,r,c-2), 500.0)),X.CHAR(10.0)),'(f) Độ tin cậy: '),E.v(4,r,c-1))

def f269(E, r, c):
    # Ô C145
    # = KET_LUAN!$C$31
    return E.v(3,31,3)

def f270(E, r, c):
    # Ô D145
    # = LEFT(KET_LUAN!$C$32,700)
    return X.LEFT(E.v(3,32,3), 700.0)

def f271(E, r, c):
    # Ô E145
    # = "Được lệnh: sinh vào mùa hành của nhật chủ vượng; đắc địa: chi khác có lộc, kình dương, trường sinh hoặc mộ của nhật can; được sinh: có chữ sinh cho n
    #   hật chủ; được trợ: có chữ cùng hành. Đạt nhiều điều kiện thì vượng (P1 tr.70)."
    return 'Được lệnh: sinh vào mùa hành của nhật chủ vượng; đắc địa: chi khác có lộc, kình dương, trường sinh hoặc mộ của nhật can; được sinh: có chữ sinh cho nhật chủ; được trợ: có chữ cùng hành. Đạt nhiều điều kiện thì vượng (P1 tr.70).'

def f272(E, r, c):
    # Ô F145
    # = "Cho biết sức của bản thân đến từ đâu (mùa sinh, gốc rễ, người giúp) và thiếu ở đâu."
    return 'Cho biết sức của bản thân đến từ đâu (mùa sinh, gốc rễ, người giúp) và thiếu ở đâu.'

def f273(E, r, c):
    # Ô G145, G146
    # = "Không đổi theo thời gian – là nền của lá số."
    return 'Không đổi theo thời gian – là nền của lá số.'

def f274(E, r, c):
    # Ô H145
    # = "Đọc cùng PP2; nếu có tiêu chí “phân xử” thì xem DOI_CHIEU dòng 28-39."
    return 'Đọc cùng PP2; nếu có tiêu chí “phân xử” thì xem DOI_CHIEU dòng 28-39.'

def f275(E, r, c):
    # Ô C146
    # = KET_LUAN!$D$31
    return E.v(3,31,4)

def f276(E, r, c):
    # Ô D146
    # = LEFT(KET_LUAN!$C$33,700)
    return X.LEFT(E.v(3,33,3), 700.0)

def f277(E, r, c):
    # Ô E146
    # = "Mỗi chữ được cho số “độ” khí theo tháng sinh (bảng P2 tr.332-345), trừ khi bị hợp, xung, khắc. Phe mình = hành nhật chủ + hành sinh cho nhật chủ; phe
    #    mình ≥ 40% tổng độ là vượng."
    return 'Mỗi chữ được cho số “độ” khí theo tháng sinh (bảng P2 tr.332-345), trừ khi bị hợp, xung, khắc. Phe mình = hành nhật chủ + hành sinh cho nhật chủ; phe mình ≥ 40% tổng độ là vượng.'

def f278(E, r, c):
    # Ô F146
    # = "Cho biết mức mạnh – yếu bằng con số; càng xa ngưỡng 40% càng chắc chắn."
    return 'Cho biết mức mạnh – yếu bằng con số; càng xa ngưỡng 40% càng chắc chắn.'

def f279(E, r, c):
    # Ô H146
    # = "Nếu tỉ lệ gần 40% (35-45%) thì nên kiểm chứng bằng sự việc trước khi chọn phương pháp."
    return 'Nếu tỉ lệ gần 40% (35-45%) thì nên kiểm chứng bằng sự việc trước khi chọn phương pháp.'

def f280(E, r, c):
    # Ô C147
    # = "PP1: "&KET_LUAN!$C$36&" | PP2: "&KET_LUAN!$D$36&" → đang dùng: "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)
    return X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(3,36,3)),' | PP2: '),E.v(3,36,4)),' → đang dùng: '),(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3)))

def f281(E, r, c):
    # Ô D147
    # = KET_LUAN!$F$36
    return E.v(3,36,6)

def f282(E, r, c):
    # Ô E147
    # = "Dụng thần là hành lá số cần nhất để cân bằng: thân vượng cần hành làm hao bớt (tài, quan, thực thương); thân nhược cần hành sinh – giúp (ấn, tỉ kiếp)
    #    (P1 tr.124-133)."
    return 'Dụng thần là hành lá số cần nhất để cân bằng: thân vượng cần hành làm hao bớt (tài, quan, thực thương); thân nhược cần hành sinh – giúp (ấn, tỉ kiếp) (P1 tr.124-133).'

def f283(E, r, c):
    # Ô F147
    # = "Đây là “thuốc” của lá số: gặp nhiều thì thuận, thiếu thì khó."
    return 'Đây là “thuốc” của lá số: gặp nhiều thì thuận, thiếu thì khó.'

def f284(E, r, c):
    # Ô H147, H148, H149, H155, H156, H157, H158, H159, H166, H173, H178, H195 … (17 ô)
    # = "Theo nguyên tắc chung: bổ dụng thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$36,KET_LUAN!$C$36)&", hỷ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$
    #   C$37)&"; hạn chế kỵ thần "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)&" (P2 tr.631-647)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Theo nguyên tắc chung: bổ dụng thần ',(E.v(3,36,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,36,3))),', hỷ thần '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3))),'; hạn chế kỵ thần '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3))),' (P2 tr.631-647).')

def f285(E, r, c):
    # Ô C148
    # = "PP1: "&KET_LUAN!$C$37&" | PP2: "&KET_LUAN!$D$37&" → đang dùng: "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$37,KET_LUAN!$C$37)
    return X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(3,37,3)),' | PP2: '),E.v(3,37,4)),' → đang dùng: '),(E.v(3,37,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,37,3)))

def f286(E, r, c):
    # Ô D148
    # = KET_LUAN!$F$37
    return E.v(3,37,6)

def f287(E, r, c):
    # Ô E148
    # = "Hỷ thần là hành sinh ra hoặc hỗ trợ dụng thần (P1 tr.124)."
    return 'Hỷ thần là hành sinh ra hoặc hỗ trợ dụng thần (P1 tr.124).'

def f288(E, r, c):
    # Ô F148
    # = "Hành hỗ trợ: có thì dụng thần mạnh hơn, mọi việc thuận hơn."
    return 'Hành hỗ trợ: có thì dụng thần mạnh hơn, mọi việc thuận hơn.'

def f289(E, r, c):
    # Ô C149
    # = "PP1: "&KET_LUAN!$C$38&" | PP2: "&KET_LUAN!$D$38&" → đang dùng: "&IF(DOI_CHIEU!$C$33=2,KET_LUAN!$D$38,KET_LUAN!$C$38)
    return X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(3,38,3)),' | PP2: '),E.v(3,38,4)),' → đang dùng: '),(E.v(3,38,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,38,3)))

def f290(E, r, c):
    # Ô D149
    # = KET_LUAN!$F$38
    return E.v(3,38,6)

def f291(E, r, c):
    # Ô E149
    # = "Kỵ thần là hành khắc dụng thần hoặc làm lá số lệch thêm (P2 tr.382)."
    return 'Kỵ thần là hành khắc dụng thần hoặc làm lá số lệch thêm (P2 tr.382).'

def f292(E, r, c):
    # Ô F149
    # = "Hành gây trở ngại: gặp nhiều (vận, năm, môi trường) thì dễ hao tổn."
    return 'Hành gây trở ngại: gặp nhiều (vận, năm, môi trường) thì dễ hao tổn.'

def f293(E, r, c):
    # Ô C150
    # = LEFT(KET_LUAN!$C$40,IFERROR(FIND(CHAR(10),KET_LUAN!$C$40)-1,LEN(KET_LUAN!$C$40)))
    return X.LEFT(E.v(3,40,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,40,3)),1.0), lambda: X.LEN(E.v(3,40,3))))

def f294(E, r, c):
    # Ô D150
    # = LEFT(KET_LUAN!$C$40,700)
    return X.LEFT(E.v(3,40,3), 700.0)

def f295(E, r, c):
    # Ô E150
    # = "Văn xương là sao chủ học vấn, văn chương, tra theo can năm hoặc can ngày (P2 tr.300-301); gặp xung phá thì giảm tác dụng."
    return 'Văn xương là sao chủ học vấn, văn chương, tra theo can năm hoặc can ngày (P2 tr.300-301); gặp xung phá thì giảm tác dụng.'

def f296(E, r, c):
    # Ô F150
    # = IF(ISNUMBER(SEARCH("Có tác dụng",KET_LUAN!$C$40)),"Thuận cho học hành, thi cử, viết lách, nghiên cứu.","Không có trợ lực rõ từ sao này; học vấn dựa và
    #   o ấn và sự nỗ lực.")
    return ('Thuận cho học hành, thi cử, viết lách, nghiên cứu.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Có tác dụng', E.v(3,40,3)))) else 'Không có trợ lực rõ từ sao này; học vấn dựa vào ấn và sự nỗ lực.')

def f297(E, r, c):
    # Ô G150
    # = "Rõ khi vận/năm chạm vào trụ có văn xương hoặc vận ấn."
    return 'Rõ khi vận/năm chạm vào trụ có văn xương hoặc vận ấn.'

def f298(E, r, c):
    # Ô H150
    # = "Phát huy qua học tập, nghề chữ nghĩa nếu sao có tác dụng."
    return 'Phát huy qua học tập, nghề chữ nghĩa nếu sao có tác dụng.'

def f299(E, r, c):
    # Ô I150, I151
    # = "Thấp – Vừa (thần sát là yếu tố phụ)"
    return 'Thấp – Vừa (thần sát là yếu tố phụ)'

def f300(E, r, c):
    # Ô C151
    # = LEFT(KET_LUAN!$C$41,IFERROR(FIND(CHAR(10),KET_LUAN!$C$41)-1,LEN(KET_LUAN!$C$41)))
    return X.LEFT(E.v(3,41,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,41,3)),1.0), lambda: X.LEN(E.v(3,41,3))))

def f301(E, r, c):
    # Ô D151
    # = LEFT(KET_LUAN!$C$41,700)
    return X.LEFT(E.v(3,41,3), 700.0)

def f302(E, r, c):
    # Ô E151
    # = "Thiên ất quý nhân tra theo can ngày/can năm ra chi (P2 tr.294-296); là cát thần giúp gặp người đỡ đầu, hóa giải hung; gặp hình xung thì như không có 
    #   (P1 tr.139-140)."
    return 'Thiên ất quý nhân tra theo can ngày/can năm ra chi (P2 tr.294-296); là cát thần giúp gặp người đỡ đầu, hóa giải hung; gặp hình xung thì như không có (P1 tr.139-140).'

def f303(E, r, c):
    # Ô F151
    # = IF(LEFT(KET_LUAN!$C$41,2)="Có","Dễ gặp người giúp đỡ lúc khó khăn; việc khó dễ được hóa giải.","Ít trợ lực từ sao này; nên chủ động xây dựng quan hệ."
    #   )
    return ('Dễ gặp người giúp đỡ lúc khó khăn; việc khó dễ được hóa giải.' if X.truthy(X.eq(X.LEFT(E.v(3,41,3), 2.0),'Có')) else 'Ít trợ lực từ sao này; nên chủ động xây dựng quan hệ.')

def f304(E, r, c):
    # Ô G151
    # = "Rõ khi vận/năm mang chi quý nhân hoặc chạm trụ có quý nhân."
    return 'Rõ khi vận/năm mang chi quý nhân hoặc chạm trụ có quý nhân.'

def f305(E, r, c):
    # Ô H151
    # = "Giữ quan hệ tốt với người lớn tuổi, cấp trên, thầy."
    return 'Giữ quan hệ tốt với người lớn tuổi, cấp trên, thầy.'

def f306(E, r, c):
    # Ô C152
    # = LEFT(KET_LUAN!$C$42,IFERROR(FIND(CHAR(10),KET_LUAN!$C$42)-1,LEN(KET_LUAN!$C$42)))
    return X.LEFT(E.v(3,42,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,42,3)),1.0), lambda: X.LEN(E.v(3,42,3))))

def f307(E, r, c):
    # Ô D152
    # = DOI_CHIEU!$C$36
    return E.v(7,36,3)

def f308(E, r, c):
    # Ô E152
    # = "Khi PP1 và PP2 khác nhau, công cụ phân xử theo quy tắc ở DOI_CHIEU (dòng 28-39); người dùng có thể chọn lại ở NHAP_LIEU C14."
    return 'Khi PP1 và PP2 khác nhau, công cụ phân xử theo quy tắc ở DOI_CHIEU (dòng 28-39); người dùng có thể chọn lại ở NHAP_LIEU C14.'

def f309(E, r, c):
    # Ô G152
    # = "Kiểm chứng bằng các năm đã qua: năm tốt/xấu thực tế khớp phương pháp nào thì tin phương pháp đó."
    return 'Kiểm chứng bằng các năm đã qua: năm tốt/xấu thực tế khớp phương pháp nào thì tin phương pháp đó.'

def f310(E, r, c):
    # Ô H152
    # = "Liệt kê 3-5 sự kiện lớn đã xảy ra, so với LUU_NIEN cột PP1/PP2 tổng hợp để chọn."
    return 'Liệt kê 3-5 sự kiện lớn đã xảy ra, so với LUU_NIEN cột PP1/PP2 tổng hợp để chọn.'

def f311(E, r, c):
    # Ô C153
    # = LEFT(KET_LUAN!$C$43,IFERROR(FIND(CHAR(10),KET_LUAN!$C$43)-1,LEN(KET_LUAN!$C$43)))
    return X.LEFT(E.v(3,43,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,43,3)),1.0), lambda: X.LEN(E.v(3,43,3))))

def f312(E, r, c):
    # Ô D153
    # = LEFT(KET_LUAN!$C$43,700)
    return X.LEFT(E.v(3,43,3), 700.0)

def f313(E, r, c):
    # Ô E153
    # = "Ghép vượng suy và dụng – hỷ – kỵ của hai phương pháp thành một kết luận."
    return 'Ghép vượng suy và dụng – hỷ – kỵ của hai phương pháp thành một kết luận.'

def f314(E, r, c):
    # Ô H153
    # = LUAN_GIAI!$H$106
    return E.v(4,106,8)

def f315(E, r, c):
    # Ô C154
    # = LEFT(KET_LUAN!$C$172,300)
    return X.LEFT(E.v(3,172,3), 300.0)

def f316(E, r, c):
    # Ô D154
    # = LEFT(KET_LUAN!$C$172,700)
    return X.LEFT(E.v(3,172,3), 700.0)

def f317(E, r, c):
    # Ô E154
    # = "Trụ năm ứng tuổi nhỏ, trụ tháng tuổi thanh niên, trụ ngày trung niên, trụ giờ tuổi già (P2 tr.284; P1 tr.149). Trụ mang dụng/hỷ thần thì giai đoạn đó
    #    thuận, mang kỵ thần thì kém."
    return 'Trụ năm ứng tuổi nhỏ, trụ tháng tuổi thanh niên, trụ ngày trung niên, trụ giờ tuổi già (P2 tr.284; P1 tr.149). Trụ mang dụng/hỷ thần thì giai đoạn đó thuận, mang kỵ thần thì kém.'

def f318(E, r, c):
    # Ô F154
    # = "Cho biết giai đoạn nào của đời thuận hơn, giai đoạn nào cần chuẩn bị."
    return 'Cho biết giai đoạn nào của đời thuận hơn, giai đoạn nào cần chuẩn bị.'

def f319(E, r, c):
    # Ô G154
    # = "Theo các mốc tuổi ghi ở kết quả; đại vận trong giai đoạn đó làm rõ thêm."
    return 'Theo các mốc tuổi ghi ở kết quả; đại vận trong giai đoạn đó làm rõ thêm.'

def f320(E, r, c):
    # Ô H154
    # = "Chuẩn bị trước (sức khỏe, tài chính, quan hệ) cho giai đoạn kém."
    return 'Chuẩn bị trước (sức khỏe, tài chính, quan hệ) cho giai đoạn kém.'

def f321(E, r, c):
    # Ô I154
    # = "Vừa (mốc tuổi là gần đúng)"
    return 'Vừa (mốc tuổi là gần đúng)'

def f322(E, r, c):
    # Ô C155
    # = LEFT(KET_LUAN!$C$56,IFERROR(FIND(CHAR(10),KET_LUAN!$C$56)-1,LEN(KET_LUAN!$C$56)))
    return X.LEFT(E.v(3,56,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,56,3)),1.0), lambda: X.LEN(E.v(3,56,3))))

def f323(E, r, c):
    # Ô D155
    # = LEFT(KET_LUAN!$C$56,700)
    return X.LEFT(E.v(3,56,3), 700.0)

def f324(E, r, c):
    # Ô E155, E156, E157, E158
    # = "Thông căn là can lộ có gốc ở chi (chi tàng can cùng hành). Có gốc thì can có sức; không gốc thì “hư phù”. Gốc gần (cùng trụ, trụ kề) mạnh hơn gốc xa;
    #    gốc bị xung, bị hợp hóa thì giảm (P1 tr.32, 59-60; P2 tr.401-424)."
    return 'Thông căn là can lộ có gốc ở chi (chi tàng can cùng hành). Có gốc thì can có sức; không gốc thì “hư phù”. Gốc gần (cùng trụ, trụ kề) mạnh hơn gốc xa; gốc bị xung, bị hợp hóa thì giảm (P1 tr.32, 59-60; P2 tr.401-424).'

def f325(E, r, c):
    # Ô F155, F156, F157, F158
    # = "Nhật chủ có gốc thì bền sức, chịu được áp lực; không gốc thì dễ dao động theo hoàn cảnh."
    return 'Nhật chủ có gốc thì bền sức, chịu được áp lực; không gốc thì dễ dao động theo hoàn cảnh.'

def f326(E, r, c):
    # Ô G155, G156, G157, G158, G159
    # = "Lúc vận/năm xung chi chứa gốc là lúc gốc yếu nhất."
    return 'Lúc vận/năm xung chi chứa gốc là lúc gốc yếu nhất.'

def f327(E, r, c):
    # Ô C156
    # = LEFT(KET_LUAN!$C$57,IFERROR(FIND(CHAR(10),KET_LUAN!$C$57)-1,LEN(KET_LUAN!$C$57)))
    return X.LEFT(E.v(3,57,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,57,3)),1.0), lambda: X.LEN(E.v(3,57,3))))

def f328(E, r, c):
    # Ô D156
    # = LEFT(KET_LUAN!$C$57,700)
    return X.LEFT(E.v(3,57,3), 700.0)

def f329(E, r, c):
    # Ô C157
    # = LEFT(KET_LUAN!$C$58,IFERROR(FIND(CHAR(10),KET_LUAN!$C$58)-1,LEN(KET_LUAN!$C$58)))
    return X.LEFT(E.v(3,58,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,58,3)),1.0), lambda: X.LEN(E.v(3,58,3))))

def f330(E, r, c):
    # Ô D157
    # = LEFT(KET_LUAN!$C$58,700)
    return X.LEFT(E.v(3,58,3), 700.0)

def f331(E, r, c):
    # Ô C158
    # = LEFT(KET_LUAN!$C$59,IFERROR(FIND(CHAR(10),KET_LUAN!$C$59)-1,LEN(KET_LUAN!$C$59)))
    return X.LEFT(E.v(3,59,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,59,3)),1.0), lambda: X.LEN(E.v(3,59,3))))

def f332(E, r, c):
    # Ô D158
    # = LEFT(KET_LUAN!$C$59,700)
    return X.LEFT(E.v(3,59,3), 700.0)

def f333(E, r, c):
    # Ô C159
    # = LEFT(KET_LUAN!$C$60,IFERROR(FIND(CHAR(10),KET_LUAN!$C$60)-1,LEN(KET_LUAN!$C$60)))
    return X.LEFT(E.v(3,60,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,60,3)),1.0), lambda: X.LEN(E.v(3,60,3))))

def f334(E, r, c):
    # Ô D159
    # = LEFT(KET_LUAN!$C$60,700)
    return X.LEFT(E.v(3,60,3), 700.0)

def f335(E, r, c):
    # Ô E159
    # = "Tổng hợp tình trạng gốc của nhật chủ và các can khác."
    return 'Tổng hợp tình trạng gốc của nhật chủ và các can khác.'

def f336(E, r, c):
    # Ô F159
    # = "Cho biết chữ nào trong lá số “có thực lực”, chữ nào chỉ là hình thức."
    return 'Cho biết chữ nào trong lá số “có thực lực”, chữ nào chỉ là hình thức.'

def f337(E, r, c):
    # Ô C160
    # = LEFT(KET_LUAN!$C$47,IFERROR(FIND(CHAR(10),KET_LUAN!$C$47)-1,LEN(KET_LUAN!$C$47)))
    return X.LEFT(E.v(3,47,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,47,3)),1.0), lambda: X.LEN(E.v(3,47,3))))

def f338(E, r, c):
    # Ô D160
    # = LEFT(KET_LUAN!$C$47,700)
    return X.LEFT(E.v(3,47,3), 700.0)

def f339(E, r, c):
    # Ô E160
    # = "Điều hầu là điều chỉnh nóng – lạnh, khô – ẩm theo mùa sinh: sinh mùa đông cần Hỏa sưởi ấm, mùa hạ cần Thủy làm mát (P1 tr.124; P2 tr.379-381)."
    return 'Điều hầu là điều chỉnh nóng – lạnh, khô – ẩm theo mùa sinh: sinh mùa đông cần Hỏa sưởi ấm, mùa hạ cần Thủy làm mát (P1 tr.124; P2 tr.379-381).'

def f340(E, r, c):
    # Ô F160
    # = "Thiếu điều hầu thì dù cân bằng vẫn có điểm “khó chịu” về sức khỏe, tâm trạng theo mùa."
    return 'Thiếu điều hầu thì dù cân bằng vẫn có điểm “khó chịu” về sức khỏe, tâm trạng theo mùa.'

def f341(E, r, c):
    # Ô G160
    # = "Rõ vào mùa khắc nghiệt và khi vận/năm mang hành điều hầu."
    return 'Rõ vào mùa khắc nghiệt và khi vận/năm mang hành điều hầu.'

def f342(E, r, c):
    # Ô H160
    # = "Bổ hành điều hầu trong sinh hoạt (môi trường sống, màu sắc, phương hướng)."
    return 'Bổ hành điều hầu trong sinh hoạt (môi trường sống, màu sắc, phương hướng).'

def f343(E, r, c):
    # Ô I160, I161, I162, I163, I164, I174, I175, I177, I179, I187, I188, I189 … (14 ô)
    # = "Vừa"
    return 'Vừa'

def f344(E, r, c):
    # Ô C161
    # = IF(LEN(KET_LUAN!$C$48)<3,"Không có",LEFT(KET_LUAN!$C$48,IFERROR(FIND(CHAR(10),KET_LUAN!$C$48)-1,LEN(KET_LUAN!$C$48))))
    return ('Không có' if X.truthy(X.lt(X.LEN(E.v(3,48,3)),3.0)) else X.LEFT(E.v(3,48,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,48,3)),1.0), lambda: X.LEN(E.v(3,48,3)))))

def f345(E, r, c):
    # Ô D161
    # = LEFT(KET_LUAN!$C$48,700)
    return X.LEFT(E.v(3,48,3), 700.0)

def f346(E, r, c):
    # Ô E161, E162, E163, E164
    # = "Thông quan là hành đứng giữa hai hành đang khắc nhau, chuyển khắc thành sinh (ví dụ Kim khắc Mộc thì dùng Thủy thông quan) (P1 tr.124)."
    return 'Thông quan là hành đứng giữa hai hành đang khắc nhau, chuyển khắc thành sinh (ví dụ Kim khắc Mộc thì dùng Thủy thông quan) (P1 tr.124).'

def f347(E, r, c):
    # Ô F161
    # = IF(LEN(KET_LUAN!$C$48)<3,"Không có cặp hành khắc nhau cần thông quan.","Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.")
    return ('Không có cặp hành khắc nhau cần thông quan.' if X.truthy(X.lt(X.LEN(E.v(3,48,3)),3.0)) else 'Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.')

def f348(E, r, c):
    # Ô G161, G162, G163, G164
    # = "Rõ khi vận/năm mang hành thông quan (tốt) hoặc làm cặp khắc gay gắt hơn (xấu)."
    return 'Rõ khi vận/năm mang hành thông quan (tốt) hoặc làm cặp khắc gay gắt hơn (xấu).'

def f349(E, r, c):
    # Ô H161, H162, H163, H164
    # = "Dùng hành thông quan trong nghề nghiệp, môi trường để giảm xung đột."
    return 'Dùng hành thông quan trong nghề nghiệp, môi trường để giảm xung đột.'

def f350(E, r, c):
    # Ô C162
    # = IF(LEN(KET_LUAN!$C$49)<3,"Không có",LEFT(KET_LUAN!$C$49,IFERROR(FIND(CHAR(10),KET_LUAN!$C$49)-1,LEN(KET_LUAN!$C$49))))
    return ('Không có' if X.truthy(X.lt(X.LEN(E.v(3,49,3)),3.0)) else X.LEFT(E.v(3,49,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,49,3)),1.0), lambda: X.LEN(E.v(3,49,3)))))

def f351(E, r, c):
    # Ô D162
    # = LEFT(KET_LUAN!$C$49,700)
    return X.LEFT(E.v(3,49,3), 700.0)

def f352(E, r, c):
    # Ô F162
    # = IF(LEN(KET_LUAN!$C$49)<3,"Không có cặp hành khắc nhau cần thông quan.","Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.")
    return ('Không có cặp hành khắc nhau cần thông quan.' if X.truthy(X.lt(X.LEN(E.v(3,49,3)),3.0)) else 'Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.')

def f353(E, r, c):
    # Ô C163
    # = IF(LEN(KET_LUAN!$C$50)<3,"Không có",LEFT(KET_LUAN!$C$50,IFERROR(FIND(CHAR(10),KET_LUAN!$C$50)-1,LEN(KET_LUAN!$C$50))))
    return ('Không có' if X.truthy(X.lt(X.LEN(E.v(3,50,3)),3.0)) else X.LEFT(E.v(3,50,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,50,3)),1.0), lambda: X.LEN(E.v(3,50,3)))))

def f354(E, r, c):
    # Ô D163
    # = LEFT(KET_LUAN!$C$50,700)
    return X.LEFT(E.v(3,50,3), 700.0)

def f355(E, r, c):
    # Ô F163
    # = IF(LEN(KET_LUAN!$C$50)<3,"Không có cặp hành khắc nhau cần thông quan.","Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.")
    return ('Không có cặp hành khắc nhau cần thông quan.' if X.truthy(X.lt(X.LEN(E.v(3,50,3)),3.0)) else 'Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.')

def f356(E, r, c):
    # Ô C164
    # = IF(LEN(KET_LUAN!$C$51)<3,"Không có",LEFT(KET_LUAN!$C$51,IFERROR(FIND(CHAR(10),KET_LUAN!$C$51)-1,LEN(KET_LUAN!$C$51))))
    return ('Không có' if X.truthy(X.lt(X.LEN(E.v(3,51,3)),3.0)) else X.LEFT(E.v(3,51,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,51,3)),1.0), lambda: X.LEN(E.v(3,51,3)))))

def f357(E, r, c):
    # Ô D164
    # = LEFT(KET_LUAN!$C$51,700)
    return X.LEFT(E.v(3,51,3), 700.0)

def f358(E, r, c):
    # Ô F164
    # = IF(LEN(KET_LUAN!$C$51)<3,"Không có cặp hành khắc nhau cần thông quan.","Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.")
    return ('Không có cặp hành khắc nhau cần thông quan.' if X.truthy(X.lt(X.LEN(E.v(3,51,3)),3.0)) else 'Có hai thế lực xung đột trong lá số; hành thông quan giúp hòa giải.')

def f359(E, r, c):
    # Ô C165
    # = LEFT(KET_LUAN!$C$162,IFERROR(FIND(CHAR(10),KET_LUAN!$C$162)-1,LEN(KET_LUAN!$C$162)))
    return X.LEFT(E.v(3,162,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,162,3)),1.0), lambda: X.LEN(E.v(3,162,3))))

def f360(E, r, c):
    # Ô D165
    # = LEFT(KET_LUAN!$C$162,700)
    return X.LEFT(E.v(3,162,3), 700.0)

def f361(E, r, c):
    # Ô E165
    # = "Tuổi khởi vận tính từ số ngày giữa ngày sinh và tiết khí gần nhất (3 ngày = 1 năm), đi thuận hay nghịch theo giới tính và can năm (P2 tr.270-272)."
    return 'Tuổi khởi vận tính từ số ngày giữa ngày sinh và tiết khí gần nhất (3 ngày = 1 năm), đi thuận hay nghịch theo giới tính và can năm (P2 tr.270-272).'

def f362(E, r, c):
    # Ô F165
    # = "Từ tuổi này đại vận bắt đầu tác động; trước đó xét tiểu vận và lưu niên."
    return 'Từ tuổi này đại vận bắt đầu tác động; trước đó xét tiểu vận và lưu niên.'

def f363(E, r, c):
    # Ô G165
    # = "Mỗi đại vận 10 năm, tính từ mốc khởi vận."
    return 'Mỗi đại vận 10 năm, tính từ mốc khởi vận.'

def f364(E, r, c):
    # Ô H165
    # = "Kiểm tra ngày giờ sinh chính xác vì tuổi khởi vận phụ thuộc số ngày đến tiết."
    return 'Kiểm tra ngày giờ sinh chính xác vì tuổi khởi vận phụ thuộc số ngày đến tiết.'

def f365(E, r, c):
    # Ô I165
    # = "Cao nếu ngày giờ sinh chính xác"
    return 'Cao nếu ngày giờ sinh chính xác'

def f366(E, r, c):
    # Ô C166
    # = LEFT(KET_LUAN!$C$163,IFERROR(FIND(CHAR(10),KET_LUAN!$C$163)-1,LEN(KET_LUAN!$C$163)))
    return X.LEFT(E.v(3,163,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,163,3)),1.0), lambda: X.LEN(E.v(3,163,3))))

def f367(E, r, c):
    # Ô D166
    # = LEFT(KET_LUAN!$C$163,700)
    return X.LEFT(E.v(3,163,3), 700.0)

def f368(E, r, c):
    # Ô E166
    # = "Đại vận 10 năm: can và chi cùng tác dụng cả 10 năm; vận mang dụng/hỷ thần là tốt, kỵ thần là xấu (P2 tr.281; P1 tr.120-123)."
    return 'Đại vận 10 năm: can và chi cùng tác dụng cả 10 năm; vận mang dụng/hỷ thần là tốt, kỵ thần là xấu (P2 tr.281; P1 tr.120-123).'

def f369(E, r, c):
    # Ô F166
    # = IF(ISNUMBER(SEARCH("Tốt",KET_LUAN!$C$163)),"Giai đoạn thuận: nên chủ động phát triển.",IF(ISNUMBER(SEARCH("Xấu",KET_LUAN!$C$163)),"Giai đoạn khó: nên 
    #   giữ gìn, chuẩn bị.","Giai đoạn bình thường: tốt xấu tùy từng năm."))
    return ('Giai đoạn thuận: nên chủ động phát triển.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Tốt', E.v(3,163,3)))) else ('Giai đoạn khó: nên giữ gìn, chuẩn bị.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Xấu', E.v(3,163,3)))) else 'Giai đoạn bình thường: tốt xấu tùy từng năm.'))

def f370(E, r, c):
    # Ô G166
    # = "Trong 10 năm của đại vận này."
    return 'Trong 10 năm của đại vận này.'

def f371(E, r, c):
    # Ô C167
    # = LUU_NIEN!$A$5&" ("&LUU_NIEN!$C$5&"): "&IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$5,LUU_NIEN!$G$5)
    return X.cat(X.cat(X.cat(X.cat(E.v(15,5,1),' ('),E.v(15,5,3)),'): '),(E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7)))

def f372(E, r, c):
    # Ô D167
    # = IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$L$5,LUU_NIEN!$K$5)
    return (E.v(15,5,12) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,11))

def f373(E, r, c):
    # Ô E167, E168, E169, E170, E171, E172
    # = "Lưu niên là can chi của năm; so hành can/chi năm với dụng – hỷ – kỵ rồi ghép với đại vận theo bảng P1 tr.133; các dấu hiệu đặc biệt (tuế vận tịnh lâm
    #   , thiên khắc địa xung…) theo P2 tr.282-289."
    return 'Lưu niên là can chi của năm; so hành can/chi năm với dụng – hỷ – kỵ rồi ghép với đại vận theo bảng P1 tr.133; các dấu hiệu đặc biệt (tuế vận tịnh lâm, thiên khắc địa xung…) theo P2 tr.282-289.'

def f374(E, r, c):
    # Ô F167
    # = IF(ISNUMBER(SEARCH("tốt",IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$5,LUU_NIEN!$G$5))),"Năm thuận: nên triển khai việc lớn.",IF(ISNUMBER(SEARCH("xấu",IF(DOI_CHI
    #   EU!$C$33=2,LUU_NIEN!$I$5,LUU_NIEN!$G$5))),"Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.","Năm bình thường."))
    return ('Năm thuận: nên triển khai việc lớn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('tốt', (E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7))))) else ('Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('xấu', (E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7))))) else 'Năm bình thường.'))

def f375(E, r, c):
    # Ô G167
    # = "Cả năm "&LUU_NIEN!$A$5&IF(LUU_NIEN!$J$5="","",". Dấu hiệu: "&LEFT(LUU_NIEN!$J$5,300))
    return X.cat(X.cat('Cả năm ',E.v(15,5,1)),('' if X.truthy(X.eq(E.v(15,5,10),'')) else X.cat('. Dấu hiệu: ',X.LEFT(E.v(15,5,10), 300.0))))

def f376(E, r, c):
    # Ô H167
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$5,LUU_NIEN!$A$30:$A$129,0)),"")="","Theo dụng thần; không có tạng phủ nào bị báo riêng tron
    #   g năm.","Sức khỏe năm này: "&LEFT(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$5,LUU_NIEN!$A$30:$A$129,0)),""),400))
    return ('Theo dụng thần; không có tạng phủ nào bị báo riêng trong năm.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,5,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.cat('Sức khỏe năm này: ',X.LEFT(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,5,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''), 400.0)))

def f377(E, r, c):
    # Ô C168
    # = LUU_NIEN!$A$6&" ("&LUU_NIEN!$C$6&"): "&IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$G$6)
    return X.cat(X.cat(X.cat(X.cat(E.v(15,6,1),' ('),E.v(15,6,3)),'): '),(E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7)))

def f378(E, r, c):
    # Ô D168
    # = IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$L$6,LUU_NIEN!$K$6)
    return (E.v(15,6,12) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,11))

def f379(E, r, c):
    # Ô F168
    # = IF(ISNUMBER(SEARCH("tốt",IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$G$6))),"Năm thuận: nên triển khai việc lớn.",IF(ISNUMBER(SEARCH("xấu",IF(DOI_CHI
    #   EU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$G$6))),"Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.","Năm bình thường."))
    return ('Năm thuận: nên triển khai việc lớn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('tốt', (E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7))))) else ('Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('xấu', (E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7))))) else 'Năm bình thường.'))

def f380(E, r, c):
    # Ô G168
    # = "Cả năm "&LUU_NIEN!$A$6&IF(LUU_NIEN!$J$6="","",". Dấu hiệu: "&LEFT(LUU_NIEN!$J$6,300))
    return X.cat(X.cat('Cả năm ',E.v(15,6,1)),('' if X.truthy(X.eq(E.v(15,6,10),'')) else X.cat('. Dấu hiệu: ',X.LEFT(E.v(15,6,10), 300.0))))

def f381(E, r, c):
    # Ô H168
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$6,LUU_NIEN!$A$30:$A$129,0)),"")="","Theo dụng thần; không có tạng phủ nào bị báo riêng tron
    #   g năm.","Sức khỏe năm này: "&LEFT(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$6,LUU_NIEN!$A$30:$A$129,0)),""),400))
    return ('Theo dụng thần; không có tạng phủ nào bị báo riêng trong năm.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,6,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.cat('Sức khỏe năm này: ',X.LEFT(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,6,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''), 400.0)))

def f382(E, r, c):
    # Ô C169
    # = LUU_NIEN!$A$7&" ("&LUU_NIEN!$C$7&"): "&IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$7,LUU_NIEN!$G$7)
    return X.cat(X.cat(X.cat(X.cat(E.v(15,7,1),' ('),E.v(15,7,3)),'): '),(E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7)))

def f383(E, r, c):
    # Ô D169
    # = IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$L$7,LUU_NIEN!$K$7)
    return (E.v(15,7,12) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,11))

def f384(E, r, c):
    # Ô F169
    # = IF(ISNUMBER(SEARCH("tốt",IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$7,LUU_NIEN!$G$7))),"Năm thuận: nên triển khai việc lớn.",IF(ISNUMBER(SEARCH("xấu",IF(DOI_CHI
    #   EU!$C$33=2,LUU_NIEN!$I$7,LUU_NIEN!$G$7))),"Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.","Năm bình thường."))
    return ('Năm thuận: nên triển khai việc lớn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('tốt', (E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7))))) else ('Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('xấu', (E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7))))) else 'Năm bình thường.'))

def f385(E, r, c):
    # Ô G169
    # = "Cả năm "&LUU_NIEN!$A$7&IF(LUU_NIEN!$J$7="","",". Dấu hiệu: "&LEFT(LUU_NIEN!$J$7,300))
    return X.cat(X.cat('Cả năm ',E.v(15,7,1)),('' if X.truthy(X.eq(E.v(15,7,10),'')) else X.cat('. Dấu hiệu: ',X.LEFT(E.v(15,7,10), 300.0))))

def f386(E, r, c):
    # Ô H169
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$7,LUU_NIEN!$A$30:$A$129,0)),"")="","Theo dụng thần; không có tạng phủ nào bị báo riêng tron
    #   g năm.","Sức khỏe năm này: "&LEFT(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$7,LUU_NIEN!$A$30:$A$129,0)),""),400))
    return ('Theo dụng thần; không có tạng phủ nào bị báo riêng trong năm.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,7,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.cat('Sức khỏe năm này: ',X.LEFT(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,7,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''), 400.0)))

def f387(E, r, c):
    # Ô C170
    # = LUU_NIEN!$A$8&" ("&LUU_NIEN!$C$8&"): "&IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8)
    return X.cat(X.cat(X.cat(X.cat(E.v(15,8,1),' ('),E.v(15,8,3)),'): '),(E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7)))

def f388(E, r, c):
    # Ô D170
    # = IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$L$8,LUU_NIEN!$K$8)
    return (E.v(15,8,12) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,11))

def f389(E, r, c):
    # Ô F170
    # = IF(ISNUMBER(SEARCH("tốt",IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8))),"Năm thuận: nên triển khai việc lớn.",IF(ISNUMBER(SEARCH("xấu",IF(DOI_CHI
    #   EU!$C$33=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8))),"Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.","Năm bình thường."))
    return ('Năm thuận: nên triển khai việc lớn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('tốt', (E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7))))) else ('Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('xấu', (E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7))))) else 'Năm bình thường.'))

def f390(E, r, c):
    # Ô G170
    # = "Cả năm "&LUU_NIEN!$A$8&IF(LUU_NIEN!$J$8="","",". Dấu hiệu: "&LEFT(LUU_NIEN!$J$8,300))
    return X.cat(X.cat('Cả năm ',E.v(15,8,1)),('' if X.truthy(X.eq(E.v(15,8,10),'')) else X.cat('. Dấu hiệu: ',X.LEFT(E.v(15,8,10), 300.0))))

def f391(E, r, c):
    # Ô H170
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$8,LUU_NIEN!$A$30:$A$129,0)),"")="","Theo dụng thần; không có tạng phủ nào bị báo riêng tron
    #   g năm.","Sức khỏe năm này: "&LEFT(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$8,LUU_NIEN!$A$30:$A$129,0)),""),400))
    return ('Theo dụng thần; không có tạng phủ nào bị báo riêng trong năm.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,8,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.cat('Sức khỏe năm này: ',X.LEFT(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,8,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''), 400.0)))

def f392(E, r, c):
    # Ô C171
    # = LUU_NIEN!$A$9&" ("&LUU_NIEN!$C$9&"): "&IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$9,LUU_NIEN!$G$9)
    return X.cat(X.cat(X.cat(X.cat(E.v(15,9,1),' ('),E.v(15,9,3)),'): '),(E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7)))

def f393(E, r, c):
    # Ô D171
    # = IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$L$9,LUU_NIEN!$K$9)
    return (E.v(15,9,12) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,11))

def f394(E, r, c):
    # Ô F171
    # = IF(ISNUMBER(SEARCH("tốt",IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$9,LUU_NIEN!$G$9))),"Năm thuận: nên triển khai việc lớn.",IF(ISNUMBER(SEARCH("xấu",IF(DOI_CHI
    #   EU!$C$33=2,LUU_NIEN!$I$9,LUU_NIEN!$G$9))),"Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.","Năm bình thường."))
    return ('Năm thuận: nên triển khai việc lớn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('tốt', (E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7))))) else ('Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('xấu', (E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7))))) else 'Năm bình thường.'))

def f395(E, r, c):
    # Ô G171
    # = "Cả năm "&LUU_NIEN!$A$9&IF(LUU_NIEN!$J$9="","",". Dấu hiệu: "&LEFT(LUU_NIEN!$J$9,300))
    return X.cat(X.cat('Cả năm ',E.v(15,9,1)),('' if X.truthy(X.eq(E.v(15,9,10),'')) else X.cat('. Dấu hiệu: ',X.LEFT(E.v(15,9,10), 300.0))))

def f396(E, r, c):
    # Ô H171
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$9,LUU_NIEN!$A$30:$A$129,0)),"")="","Theo dụng thần; không có tạng phủ nào bị báo riêng tron
    #   g năm.","Sức khỏe năm này: "&LEFT(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$9,LUU_NIEN!$A$30:$A$129,0)),""),400))
    return ('Theo dụng thần; không có tạng phủ nào bị báo riêng trong năm.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,9,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.cat('Sức khỏe năm này: ',X.LEFT(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,9,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''), 400.0)))

def f397(E, r, c):
    # Ô C172
    # = LUU_NIEN!$A$10&" ("&LUU_NIEN!$C$10&"): "&IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$10,LUU_NIEN!$G$10)
    return X.cat(X.cat(X.cat(X.cat(E.v(15,10,1),' ('),E.v(15,10,3)),'): '),(E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7)))

def f398(E, r, c):
    # Ô D172
    # = IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$L$10,LUU_NIEN!$K$10)
    return (E.v(15,10,12) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,11))

def f399(E, r, c):
    # Ô F172
    # = IF(ISNUMBER(SEARCH("tốt",IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$10,LUU_NIEN!$G$10))),"Năm thuận: nên triển khai việc lớn.",IF(ISNUMBER(SEARCH("xấu",IF(DOI_C
    #   HIEU!$C$33=2,LUU_NIEN!$I$10,LUU_NIEN!$G$10))),"Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.","Năm bình thường."))
    return ('Năm thuận: nên triển khai việc lớn.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('tốt', (E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7))))) else ('Năm khó: thận trọng, giữ gìn sức khỏe, tránh quyết định vội.' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('xấu', (E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7))))) else 'Năm bình thường.'))

def f400(E, r, c):
    # Ô G172
    # = "Cả năm "&LUU_NIEN!$A$10&IF(LUU_NIEN!$J$10="","",". Dấu hiệu: "&LEFT(LUU_NIEN!$J$10,300))
    return X.cat(X.cat('Cả năm ',E.v(15,10,1)),('' if X.truthy(X.eq(E.v(15,10,10),'')) else X.cat('. Dấu hiệu: ',X.LEFT(E.v(15,10,10), 300.0))))

def f401(E, r, c):
    # Ô H172
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$10,LUU_NIEN!$A$30:$A$129,0)),"")="","Theo dụng thần; không có tạng phủ nào bị báo riêng tro
    #   ng năm.","Sức khỏe năm này: "&LEFT(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH(LUU_NIEN!$A$10,LUU_NIEN!$A$30:$A$129,0)),""),400))
    return ('Theo dụng thần; không có tạng phủ nào bị báo riêng trong năm.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,10,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.cat('Sức khỏe năm này: ',X.LEFT(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(E.v(15,10,1), E.rg(15,30,1,129,1), 0.0)), lambda: ''), 400.0)))

def f402(E, r, c):
    # Ô C173
    # = LEFT(KET_LUAN!$C$62,300)
    return X.LEFT(E.v(3,62,3), 300.0)

def f403(E, r, c):
    # Ô D173
    # = LEFT(KET_LUAN!$C$62,700)
    return X.LEFT(E.v(3,62,3), 700.0)

def f404(E, r, c):
    # Ô E173
    # = "Thần sát còn hiệu lực là sao không bị hình, xung, khắc, hại (P1 tr.137-140). Sao tốt thêm may mắn, sao xấu thêm rủi ro ở mặt sao đó chủ."
    return 'Thần sát còn hiệu lực là sao không bị hình, xung, khắc, hại (P1 tr.137-140). Sao tốt thêm may mắn, sao xấu thêm rủi ro ở mặt sao đó chủ.'

def f405(E, r, c):
    # Ô F173
    # = "Dùng như lời nhắc về điểm mạnh (sao tốt) và điểm cần đề phòng (sao xấu)."
    return 'Dùng như lời nhắc về điểm mạnh (sao tốt) và điểm cần đề phòng (sao xấu).'

def f406(E, r, c):
    # Ô G173
    # = "Rõ khi vận/năm chạm trụ có sao."
    return 'Rõ khi vận/năm chạm trụ có sao.'

def f407(E, r, c):
    # Ô C174
    # = LEFT(KET_LUAN!$C$63,IFERROR(FIND(CHAR(10),KET_LUAN!$C$63)-1,LEN(KET_LUAN!$C$63)))
    return X.LEFT(E.v(3,63,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,63,3)),1.0), lambda: X.LEN(E.v(3,63,3))))

def f408(E, r, c):
    # Ô D174
    # = LEFT(KET_LUAN!$C$63,700)
    return X.LEFT(E.v(3,63,3), 700.0)

def f409(E, r, c):
    # Ô E174
    # = "Cát thần gặp hình xung khắc hại thì “như không có” (P1 tr.139-140)."
    return 'Cát thần gặp hình xung khắc hại thì “như không có” (P1 tr.139-140).'

def f410(E, r, c):
    # Ô F174
    # = "Những sao này không nên trông cậy."
    return 'Những sao này không nên trông cậy.'

def f411(E, r, c):
    # Ô G174
    # = "—"
    return '—'

def f412(E, r, c):
    # Ô H174
    # = "Không cần làm gì riêng."
    return 'Không cần làm gì riêng.'

def f413(E, r, c):
    # Ô C175
    # = LEFT(KET_LUAN!$C$115,IFERROR(FIND(CHAR(10),KET_LUAN!$C$115)-1,LEN(KET_LUAN!$C$115)))
    return X.LEFT(E.v(3,115,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,115,3)),1.0), lambda: X.LEN(E.v(3,115,3))))

def f414(E, r, c):
    # Ô D175
    # = LUC_THAN!$G$23
    return E.v(16,23,7)

def f415(E, r, c):
    # Ô E175
    # = "Cách của sách: nam can giờ khắc can ngày → con trai; can ngày khắc can giờ, hoặc tương sinh, tỉ hòa → con gái; nữ thì ngược lại (P2 tr.504; sách nói 
    #   đã kiểm hơn 250 trường hợp)."
    return 'Cách của sách: nam can giờ khắc can ngày → con trai; can ngày khắc can giờ, hoặc tương sinh, tỉ hòa → con gái; nữ thì ngược lại (P2 tr.504; sách nói đã kiểm hơn 250 trường hợp).'

def f416(E, r, c):
    # Ô F175
    # = "Chỉ là dự đoán lần sinh đầu, để tham khảo."
    return 'Chỉ là dự đoán lần sinh đầu, để tham khảo.'

def f417(E, r, c):
    # Ô G175
    # = "Lần sinh con đầu."
    return 'Lần sinh con đầu.'

def f418(E, r, c):
    # Ô H175
    # = "Không dùng để quyết định việc sinh con."
    return 'Không dùng để quyết định việc sinh con.'

def f419(E, r, c):
    # Ô C176
    # = LEFT(KET_LUAN!$C$118,IFERROR(FIND(CHAR(10),KET_LUAN!$C$118)-1,LEN(KET_LUAN!$C$118)))
    return X.LEFT(E.v(3,118,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,118,3)),1.0), lambda: X.LEN(E.v(3,118,3))))

def f420(E, r, c):
    # Ô D176
    # = LEFT(KET_LUAN!$C$118,700)
    return X.LEFT(E.v(3,118,3), 700.0)

def f421(E, r, c):
    # Ô E176
    # = "Cung phối ngẫu là chi ngày; sao vợ của nam là tài, sao chồng của nữ là quan sát (P1 tr.153-155). Sao được ngôi, vượng, là dụng/hỷ thì người phối ngẫu
    #    có lực."
    return 'Cung phối ngẫu là chi ngày; sao vợ của nam là tài, sao chồng của nữ là quan sát (P1 tr.153-155). Sao được ngôi, vượng, là dụng/hỷ thì người phối ngẫu có lực.'

def f422(E, r, c):
    # Ô F176
    # = LUAN_GIAI!$F$115
    return E.v(4,115,6)

def f423(E, r, c):
    # Ô H176
    # = "Vun đắp quan hệ, nhất là trong vận/năm bất lợi cho hôn nhân."
    return 'Vun đắp quan hệ, nhất là trong vận/năm bất lợi cho hôn nhân.'

def f424(E, r, c):
    # Ô C177
    # = LEFT(KET_LUAN!$C$119,IFERROR(FIND(CHAR(10),KET_LUAN!$C$119)-1,LEN(KET_LUAN!$C$119)))
    return X.LEFT(E.v(3,119,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,119,3)),1.0), lambda: X.LEN(E.v(3,119,3))))

def f425(E, r, c):
    # Ô D177
    # = LEFT(KET_LUAN!$C$119,700)
    return X.LEFT(E.v(3,119,3), 700.0)

def f426(E, r, c):
    # Ô E177
    # = "Năm thuận kết hôn: lưu niên hợp trụ ngày, sao phối ngẫu đến, hoặc nhật nhược gặp năm sinh trợ (P1 tr.201; P2 tr.463)."
    return 'Năm thuận kết hôn: lưu niên hợp trụ ngày, sao phối ngẫu đến, hoặc nhật nhược gặp năm sinh trợ (P1 tr.201; P2 tr.463).'

def f427(E, r, c):
    # Ô F177
    # = "Các năm có “duyên” rõ hơn; ngoài các năm này vẫn có thể kết hôn."
    return 'Các năm có “duyên” rõ hơn; ngoài các năm này vẫn có thể kết hôn.'

def f428(E, r, c):
    # Ô G177
    # = KET_LUAN!$C$119
    return E.v(3,119,3)

def f429(E, r, c):
    # Ô H177
    # = "Ưu tiên chọn năm thuận; tránh năm bất lợi cho hôn nhân."
    return 'Ưu tiên chọn năm thuận; tránh năm bất lợi cho hôn nhân.'

def f430(E, r, c):
    # Ô C178
    # = LEFT(KET_LUAN!$C$120,IFERROR(FIND(CHAR(10),KET_LUAN!$C$120)-1,LEN(KET_LUAN!$C$120)))
    return X.LEFT(E.v(3,120,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,120,3)),1.0), lambda: X.LEN(E.v(3,120,3))))

def f431(E, r, c):
    # Ô D178
    # = LEFT(KET_LUAN!$C$120,700)
    return X.LEFT(E.v(3,120,3), 700.0)

def f432(E, r, c):
    # Ô E178
    # = "Tuổi hợp theo tam hợp, lục hợp với chi ngày/chi năm; phương hướng theo dụng thần (P2 tr.487-499, 631-632)."
    return 'Tuổi hợp theo tam hợp, lục hợp với chi ngày/chi năm; phương hướng theo dụng thần (P2 tr.487-499, 631-632).'

def f433(E, r, c):
    # Ô F178
    # = "Gợi ý tuổi và hướng thuận; không phải điều kiện bắt buộc."
    return 'Gợi ý tuổi và hướng thuận; không phải điều kiện bắt buộc.'

def f434(E, r, c):
    # Ô G178
    # = "Khi chọn người, nơi ở, hướng làm việc."
    return 'Khi chọn người, nơi ở, hướng làm việc.'

def f435(E, r, c):
    # Ô I178
    # = "Thấp – Vừa"
    return 'Thấp – Vừa'

def f436(E, r, c):
    # Ô C179
    # = LEFT(KET_LUAN!$C$121,400)
    return X.LEFT(E.v(3,121,3), 400.0)

def f437(E, r, c):
    # Ô D179
    # = LEFT(KET_LUAN!$C$121,700)
    return X.LEFT(E.v(3,121,3), 700.0)

def f438(E, r, c):
    # Ô E179
    # = "Ghép kết luận từng người thân (chấm theo P1 tr.159) và hôn nhân."
    return 'Ghép kết luận từng người thân (chấm theo P1 tr.159) và hôn nhân.'

def f439(E, r, c):
    # Ô F179
    # = "Tổng quan duyên với gia đình: ai có lực, ai cần quan tâm hơn."
    return 'Tổng quan duyên với gia đình: ai có lực, ai cần quan tâm hơn.'

def f440(E, r, c):
    # Ô H179
    # = "Quan tâm nhiều hơn tới người thân có mức Kém, nhất là trong năm có dấu hiệu bất lợi (LUC_THAN cột H)."
    return 'Quan tâm nhiều hơn tới người thân có mức Kém, nhất là trong năm có dấu hiệu bất lợi (LUC_THAN cột H).'

def f441(E, r, c):
    # Ô C180
    # = LEFT(KET_LUAN!$C$144,IFERROR(FIND(CHAR(10),KET_LUAN!$C$144)-1,LEN(KET_LUAN!$C$144)))
    return X.LEFT(E.v(3,144,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,144,3)),1.0), lambda: X.LEN(E.v(3,144,3))))

def f442(E, r, c):
    # Ô D180
    # = LEFT(KET_LUAN!$C$144,700)
    return X.LEFT(E.v(3,144,3), 700.0)

def f443(E, r, c):
    # Ô E180
    # = "Các năm có nhiều dấu hiệu nặng của sách: tuế vận tịnh lâm, thiên khắc địa xung trụ ngày, vận kỵ + năm kỵ, nhật can nhập mộ… (P1 tr.201-203; P2 tr.282
    #   -289, 566)."
    return 'Các năm có nhiều dấu hiệu nặng của sách: tuế vận tịnh lâm, thiên khắc địa xung trụ ngày, vận kỵ + năm kỵ, nhật can nhập mộ… (P1 tr.201-203; P2 tr.282-289, 566).'

def f444(E, r, c):
    # Ô F180
    # = "Là năm cần thận trọng mọi mặt (sức khỏe, đi lại, quyết định lớn); không phải dự báo chắc chắn điều xấu."
    return 'Là năm cần thận trọng mọi mặt (sức khỏe, đi lại, quyết định lớn); không phải dự báo chắc chắn điều xấu.'

def f445(E, r, c):
    # Ô G180
    # = KET_LUAN!$C$144
    return E.v(3,144,3)

def f446(E, r, c):
    # Ô H180
    # = "Khám sức khỏe, tránh mạo hiểm, giữ hòa khí trong các năm này."
    return 'Khám sức khỏe, tránh mạo hiểm, giữ hòa khí trong các năm này.'

def f447(E, r, c):
    # Ô I180
    # = "Vừa – công cụ không kết luận năm mất (P1 tr.201)"
    return 'Vừa – công cụ không kết luận năm mất (P1 tr.201)'

def f448(E, r, c):
    # Ô C181
    # = LEFT(KET_LUAN!$C$45,IFERROR(FIND(CHAR(10),KET_LUAN!$C$45)-1,LEN(KET_LUAN!$C$45)))
    return X.LEFT(E.v(3,45,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,45,3)),1.0), lambda: X.LEN(E.v(3,45,3))))

def f449(E, r, c):
    # Ô D181
    # = LEFT(KET_LUAN!$C$45,700)
    return X.LEFT(E.v(3,45,3), 700.0)

def f450(E, r, c):
    # Ô E181
    # = "Bảng P1 tr.173: giàu – nghèo, sang – hèn, cát – hung, thọ – yểu dựa vào vượng suy và nhóm sao là dụng hay kỵ."
    return 'Bảng P1 tr.173: giàu – nghèo, sang – hèn, cát – hung, thọ – yểu dựa vào vượng suy và nhóm sao là dụng hay kỵ.'

def f451(E, r, c):
    # Ô F181
    # = "Bức tranh chung về nền tảng của đời người."
    return 'Bức tranh chung về nền tảng của đời người.'

def f452(E, r, c):
    # Ô H181
    # = LUAN_GIAI!$H$116
    return E.v(4,116,8)

def f453(E, r, c):
    # Ô C182
    # = LEFT(KET_LUAN!$C$81,IFERROR(FIND(CHAR(10),KET_LUAN!$C$81)-1,LEN(KET_LUAN!$C$81)))
    return X.LEFT(E.v(3,81,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,81,3)),1.0), lambda: X.LEN(E.v(3,81,3))))

def f454(E, r, c):
    # Ô D182
    # = LEFT(KET_LUAN!$C$81,700)
    return X.LEFT(E.v(3,81,3), 700.0)

def f455(E, r, c):
    # Ô E182
    # = "Xếp bậc giàu theo vượng suy nhật chủ và sức của sao tài (P2 tr.520-530; P1 tr.173)."
    return 'Xếp bậc giàu theo vượng suy nhật chủ và sức của sao tài (P2 tr.520-530; P1 tr.173).'

def f456(E, r, c):
    # Ô F182
    # = "Mức của cải có thể đạt được khi gặp vận thuận."
    return 'Mức của cải có thể đạt được khi gặp vận thuận.'

def f457(E, r, c):
    # Ô H182, H183, H184, H185, H186
    # = "Đẩy mạnh làm ăn trong vận dễ phát tài; giữ tiền, tránh đầu cơ trong vận cẩn trọng."
    return 'Đẩy mạnh làm ăn trong vận dễ phát tài; giữ tiền, tránh đầu cơ trong vận cẩn trọng.'

def f458(E, r, c):
    # Ô C183
    # = LEFT(KET_LUAN!$C$82,IFERROR(FIND(CHAR(10),KET_LUAN!$C$82)-1,LEN(KET_LUAN!$C$82)))
    return X.LEFT(E.v(3,82,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,82,3)),1.0), lambda: X.LEN(E.v(3,82,3))))

def f459(E, r, c):
    # Ô D183
    # = LEFT(KET_LUAN!$C$82,700)
    return X.LEFT(E.v(3,82,3), 700.0)

def f460(E, r, c):
    # Ô E183
    # = "Tỉ lệ độ sao tài trên tổng độ (PP2) – chỉ là số đo tham khảo."
    return 'Tỉ lệ độ sao tài trên tổng độ (PP2) – chỉ là số đo tham khảo.'

def f461(E, r, c):
    # Ô F183
    # = "Sao tài càng có lực thì khả năng kiếm tiền càng rõ."
    return 'Sao tài càng có lực thì khả năng kiếm tiền càng rõ.'

def f462(E, r, c):
    # Ô C184
    # = LEFT(KET_LUAN!$C$83,IFERROR(FIND(CHAR(10),KET_LUAN!$C$83)-1,LEN(KET_LUAN!$C$83)))
    return X.LEFT(E.v(3,83,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,83,3)),1.0), lambda: X.LEN(E.v(3,83,3))))

def f463(E, r, c):
    # Ô D184
    # = LEFT(KET_LUAN!$C$83,700)
    return X.LEFT(E.v(3,83,3), 700.0)

def f464(E, r, c):
    # Ô E184
    # = "Nguồn tiền theo các câu sách: tài từ đâu (nghề, người, thừa kế…) (P2 tr.520-530)."
    return 'Nguồn tiền theo các câu sách: tài từ đâu (nghề, người, thừa kế…) (P2 tr.520-530).'

def f465(E, r, c):
    # Ô F184
    # = "Gợi ý kênh kiếm tiền phù hợp."
    return 'Gợi ý kênh kiếm tiền phù hợp.'

def f466(E, r, c):
    # Ô C185
    # = LEFT(KET_LUAN!$C$84,IFERROR(FIND(CHAR(10),KET_LUAN!$C$84)-1,LEN(KET_LUAN!$C$84)))
    return X.LEFT(E.v(3,84,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,84,3)),1.0), lambda: X.LEN(E.v(3,84,3))))

def f467(E, r, c):
    # Ô D185
    # = LEFT(KET_LUAN!$C$84,700)
    return X.LEFT(E.v(3,84,3), 700.0)

def f468(E, r, c):
    # Ô E185
    # = "Các dấu hiệu giữ của hay hao tài của sách (kho tài, tỉ kiếp tranh tài…)."
    return 'Các dấu hiệu giữ của hay hao tài của sách (kho tài, tỉ kiếp tranh tài…).'

def f469(E, r, c):
    # Ô F185
    # = "Cho biết có giữ được tiền hay dễ hao."
    return 'Cho biết có giữ được tiền hay dễ hao.'

def f470(E, r, c):
    # Ô C186
    # = LEFT(KET_LUAN!$C$85,IFERROR(FIND(CHAR(10),KET_LUAN!$C$85)-1,LEN(KET_LUAN!$C$85)))
    return X.LEFT(E.v(3,85,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,85,3)),1.0), lambda: X.LEN(E.v(3,85,3))))

def f471(E, r, c):
    # Ô D186
    # = LEFT(KET_LUAN!$C$85,700)
    return X.LEFT(E.v(3,85,3), 700.0)

def f472(E, r, c):
    # Ô E186
    # = "Đại vận phát tài: vận tài khi thân vượng, vận tỉ kiếp khi thân nhược tài nhiều… (P2 tr.521-523)."
    return 'Đại vận phát tài: vận tài khi thân vượng, vận tỉ kiếp khi thân nhược tài nhiều… (P2 tr.521-523).'

def f473(E, r, c):
    # Ô F186
    # = "Giai đoạn nên đẩy mạnh làm ăn và giai đoạn nên thận trọng."
    return 'Giai đoạn nên đẩy mạnh làm ăn và giai đoạn nên thận trọng.'

def f474(E, r, c):
    # Ô G186
    # = KET_LUAN!$C$85
    return E.v(3,85,3)

def f475(E, r, c):
    # Ô C187
    # = LEFT(KET_LUAN!$C$71,IFERROR(FIND(CHAR(10),KET_LUAN!$C$71)-1,LEN(KET_LUAN!$C$71)))
    return X.LEFT(E.v(3,71,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,71,3)),1.0), lambda: X.LEN(E.v(3,71,3))))

def f476(E, r, c):
    # Ô D187
    # = LEFT(KET_LUAN!$C$71,700)
    return X.LEFT(E.v(3,71,3), 700.0)

def f477(E, r, c):
    # Ô E187
    # = "Các câu sách về thông minh, năng khiếu (ấn, thực thần, thương quan, văn xương…) (P1 tr.87-95; P2 tr.573-577)."
    return 'Các câu sách về thông minh, năng khiếu (ấn, thực thần, thương quan, văn xương…) (P1 tr.87-95; P2 tr.573-577).'

def f478(E, r, c):
    # Ô F187
    # = "Thế mạnh trí tuệ để định hướng học và nghề."
    return 'Thế mạnh trí tuệ để định hướng học và nghề.'

def f479(E, r, c):
    # Ô G187, G188, G189, G190
    # = "Bộc lộ suốt đời; rõ hơn trong vận dụng/hỷ thần."
    return 'Bộc lộ suốt đời; rõ hơn trong vận dụng/hỷ thần.'

def f480(E, r, c):
    # Ô H187, H188, H189, H190
    # = "Chọn học, nghề theo thế mạnh; nghề hợp dụng thần: "&LEFT(TO_CHAT!$C$17,IFERROR(FIND(CHAR(10),TO_CHAT!$C$17)-1,LEN(TO_CHAT!$C$17)))
    return X.cat('Chọn học, nghề theo thế mạnh; nghề hợp dụng thần: ',X.LEFT(E.v(20,17,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(20,17,3)),1.0), lambda: X.LEN(E.v(20,17,3)))))

def f481(E, r, c):
    # Ô C188
    # = LEFT(KET_LUAN!$C$72,IFERROR(FIND(CHAR(10),KET_LUAN!$C$72)-1,LEN(KET_LUAN!$C$72)))
    return X.LEFT(E.v(3,72,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,72,3)),1.0), lambda: X.LEN(E.v(3,72,3))))

def f482(E, r, c):
    # Ô D188
    # = LEFT(KET_LUAN!$C$72,700)
    return X.LEFT(E.v(3,72,3), 700.0)

def f483(E, r, c):
    # Ô E188
    # = "Nghề theo thập thần và theo hành dụng thần (P1 tr.24-26)."
    return 'Nghề theo thập thần và theo hành dụng thần (P1 tr.24-26).'

def f484(E, r, c):
    # Ô F188
    # = "Gợi ý nghề hợp với lá số."
    return 'Gợi ý nghề hợp với lá số.'

def f485(E, r, c):
    # Ô C189
    # = LEFT(KET_LUAN!$C$73,IFERROR(FIND(CHAR(10),KET_LUAN!$C$73)-1,LEN(KET_LUAN!$C$73)))
    return X.LEFT(E.v(3,73,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,73,3)),1.0), lambda: X.LEN(E.v(3,73,3))))

def f486(E, r, c):
    # Ô D189
    # = LEFT(KET_LUAN!$C$73,700)
    return X.LEFT(E.v(3,73,3), 700.0)

def f487(E, r, c):
    # Ô E189
    # = "Tính cách theo hành nhật chủ và mức vượng suy (P2 tr.578-581)."
    return 'Tính cách theo hành nhật chủ và mức vượng suy (P2 tr.578-581).'

def f488(E, r, c):
    # Ô F189
    # = "Nét tính cách bẩm sinh, có thể rèn luyện."
    return 'Nét tính cách bẩm sinh, có thể rèn luyện.'

def f489(E, r, c):
    # Ô C190
    # = LEFT(KET_LUAN!$C$75,IFERROR(FIND(CHAR(10),KET_LUAN!$C$75)-1,LEN(KET_LUAN!$C$75)))
    return X.LEFT(E.v(3,75,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,75,3)),1.0), lambda: X.LEN(E.v(3,75,3))))

def f490(E, r, c):
    # Ô D190
    # = LEFT(KET_LUAN!$C$75,700)
    return X.LEFT(E.v(3,75,3), 700.0)

def f491(E, r, c):
    # Ô E190
    # = "Tổng hợp tố chất cốt lõi và nghề hợp dụng thần."
    return 'Tổng hợp tố chất cốt lõi và nghề hợp dụng thần.'

def f492(E, r, c):
    # Ô F190
    # = "Định hướng học tập, nghề nghiệp phù hợp."
    return 'Định hướng học tập, nghề nghiệp phù hợp.'

def f493(E, r, c):
    # Ô C191
    # = LEFT(KET_LUAN!$C$98,300)
    return X.LEFT(E.v(3,98,3), 300.0)

def f494(E, r, c):
    # Ô D191
    # = LEFT(KET_LUAN!$C$98,700)
    return X.LEFT(E.v(3,98,3), 700.0)

def f495(E, r, c):
    # Ô E191
    # = "Ghép tài vận, quan vận, học tập, danh tiếng (chương 18-20); mức theo vai trò sao chủ sự (G17)."
    return 'Ghép tài vận, quan vận, học tập, danh tiếng (chương 18-20); mức theo vai trò sao chủ sự (G17).'

def f496(E, r, c):
    # Ô F191
    # = "Bức tranh chung về công việc và địa vị."
    return 'Bức tranh chung về công việc và địa vị.'

def f497(E, r, c):
    # Ô H191
    # = "Đẩy mạnh trong vận dụng/hỷ; giữ vững trong vận kỵ."
    return 'Đẩy mạnh trong vận dụng/hỷ; giữ vững trong vận kỵ.'

def f498(E, r, c):
    # Ô C192
    # = LEFT(KET_LUAN!$C$140,IFERROR(FIND(CHAR(10),KET_LUAN!$C$140)-1,LEN(KET_LUAN!$C$140)))
    return X.LEFT(E.v(3,140,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,140,3)),1.0), lambda: X.LEN(E.v(3,140,3))))

def f499(E, r, c):
    # Ô D192
    # = LEFT(KET_LUAN!$C$140,700)
    return X.LEFT(E.v(3,140,3), 700.0)

def f500(E, r, c):
    # Ô E192
    # = "Xếp năm theo 15 loại năm hạn của sách (P2 tr.652-653) và phản/phục ngâm (P2 tr.657-659)."
    return 'Xếp năm theo 15 loại năm hạn của sách (P2 tr.652-653) và phản/phục ngâm (P2 tr.657-659).'

def f501(E, r, c):
    # Ô F192
    # = "Năm cần thận trọng hơn về sức khỏe, đi lại, quyết định lớn."
    return 'Năm cần thận trọng hơn về sức khỏe, đi lại, quyết định lớn.'

def f502(E, r, c):
    # Ô G192
    # = KET_LUAN!$C$140
    return E.v(3,140,3)

def f503(E, r, c):
    # Ô H192
    # = "Xem cách hóa giải ở dòng dưới (chương 24)."
    return 'Xem cách hóa giải ở dòng dưới (chương 24).'

def f504(E, r, c):
    # Ô C193
    # = LEFT(KET_LUAN!$C$141,IFERROR(FIND(CHAR(10),KET_LUAN!$C$141)-1,LEN(KET_LUAN!$C$141)))
    return X.LEFT(E.v(3,141,3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,141,3)),1.0), lambda: X.LEN(E.v(3,141,3))))

def f505(E, r, c):
    # Ô D193
    # = LEFT(KET_LUAN!$C$141,700)
    return X.LEFT(E.v(3,141,3), 700.0)

def f506(E, r, c):
    # Ô E193
    # = "Cách hóa giải của chương 24: dùng phương vị, màu sắc, nghề, môi trường thuộc dụng thần; tránh hướng thái tuế (P2 tr.631-662)."
    return 'Cách hóa giải của chương 24: dùng phương vị, màu sắc, nghề, môi trường thuộc dụng thần; tránh hướng thái tuế (P2 tr.631-662).'

def f507(E, r, c):
    # Ô F193
    # = "Những việc cụ thể làm được để giảm bất lợi trong năm."
    return 'Những việc cụ thể làm được để giảm bất lợi trong năm.'

def f508(E, r, c):
    # Ô G193
    # = "Trong năm xem."
    return 'Trong năm xem.'

def f509(E, r, c):
    # Ô H193
    # = KET_LUAN!$C$141
    return E.v(3,141,3)

def f510(E, r, c):
    # Ô I193
    # = "Vừa (sách nêu nguyên tắc)"
    return 'Vừa (sách nêu nguyên tắc)'

def f511(E, r, c):
    # Ô C194
    # = LEFT(KET_LUAN!$C$142,300)
    return X.LEFT(E.v(3,142,3), 300.0)

def f512(E, r, c):
    # Ô D194
    # = LEFT(KET_LUAN!$C$142,700)
    return X.LEFT(E.v(3,142,3), 700.0)

def f513(E, r, c):
    # Ô E194
    # = "Ghép sức khỏe chung, bệnh tật – tạng phủ, thương tật theo vận (chương 21)."
    return 'Ghép sức khỏe chung, bệnh tật – tạng phủ, thương tật theo vận (chương 21).'

def f514(E, r, c):
    # Ô F194
    # = "Tổng quan sức khỏe và điểm yếu cần theo dõi."
    return 'Tổng quan sức khỏe và điểm yếu cần theo dõi.'

def f515(E, r, c):
    # Ô H194
    # = "Khám định kỳ, nhất là các năm được báo; không thay chẩn đoán y khoa."
    return 'Khám định kỳ, nhất là các năm được báo; không thay chẩn đoán y khoa.'

def f516(E, r, c):
    # Ô C195
    # = LEFT(KET_LUAN!$C$202,300)
    return X.LEFT(E.v(3,202,3), 300.0)

def f517(E, r, c):
    # Ô D195
    # = LEFT(KET_LUAN!$C$202,700)
    return X.LEFT(E.v(3,202,3), 700.0)

def f518(E, r, c):
    # Ô E195, E196, E197, E198, E199, E200
    # = "Ghép từ các kết luận đã có căn cứ ở từng sheet, không thêm quy tắc."
    return 'Ghép từ các kết luận đã có căn cứ ở từng sheet, không thêm quy tắc.'

def f519(E, r, c):
    # Ô F195, F196, F197, F198, F199, F200
    # = "Xem bài luận giải liền mạch (KET_LUAN cột H dòng 208-209) để đọc thành văn."
    return 'Xem bài luận giải liền mạch (KET_LUAN cột H dòng 208-209) để đọc thành văn.'

def f520(E, r, c):
    # Ô C196
    # = LEFT(KET_LUAN!$C$203,300)
    return X.LEFT(E.v(3,203,3), 300.0)

def f521(E, r, c):
    # Ô D196
    # = LEFT(KET_LUAN!$C$203,700)
    return X.LEFT(E.v(3,203,3), 700.0)

def f522(E, r, c):
    # Ô C197
    # = LEFT(KET_LUAN!$C$204,300)
    return X.LEFT(E.v(3,204,3), 300.0)

def f523(E, r, c):
    # Ô D197
    # = LEFT(KET_LUAN!$C$204,700)
    return X.LEFT(E.v(3,204,3), 700.0)

def f524(E, r, c):
    # Ô C198
    # = LEFT(KET_LUAN!$C$205,300)
    return X.LEFT(E.v(3,205,3), 300.0)

def f525(E, r, c):
    # Ô D198
    # = LEFT(KET_LUAN!$C$205,700)
    return X.LEFT(E.v(3,205,3), 700.0)

def f526(E, r, c):
    # Ô C199
    # = LEFT(KET_LUAN!$C$206,300)
    return X.LEFT(E.v(3,206,3), 300.0)

def f527(E, r, c):
    # Ô D199
    # = LEFT(KET_LUAN!$C$206,700)
    return X.LEFT(E.v(3,206,3), 700.0)

def f528(E, r, c):
    # Ô C200
    # = LEFT(KET_LUAN!$C$207,300)
    return X.LEFT(E.v(3,207,3), 300.0)

def f529(E, r, c):
    # Ô D200
    # = LEFT(KET_LUAN!$C$207,700)
    return X.LEFT(E.v(3,207,3), 700.0)

def f530(E, r, c):
    # Ô C201
    # = LEFT(KET_LUAN!$C$195,300)
    return X.LEFT(E.v(3,195,3), 300.0)

def f531(E, r, c):
    # Ô D201
    # = LEFT(KET_LUAN!$C$195,700)
    return X.LEFT(E.v(3,195,3), 700.0)

def f532(E, r, c):
    # Ô E201, E202, E203, E204, E205, E206, E207
    # = "NGOÀI sách Thiệu Vĩ Hoa: quy luật rút từ tài liệu “Phân tích thí dụ thực tế”, đã thử trên 5 lá có sự việc (THAM_KHAO)."
    return 'NGOÀI sách Thiệu Vĩ Hoa: quy luật rút từ tài liệu “Phân tích thí dụ thực tế”, đã thử trên 5 lá có sự việc (THAM_KHAO).'

def f533(E, r, c):
    # Ô F201, F202, F203, F204, F205, F206, F207
    # = "Chỉ dùng để gợi ý và hỏi lại người được xem; không dùng để sửa kết luận chính."
    return 'Chỉ dùng để gợi ý và hỏi lại người được xem; không dùng để sửa kết luận chính.'

def f534(E, r, c):
    # Ô G201, G202, G203, G204, G205, G206, G207
    # = "Theo từng quy luật (THAM_KHAO cột E)."
    return 'Theo từng quy luật (THAM_KHAO cột E).'

def f535(E, r, c):
    # Ô H201, H202, H203, H204, H205, H206, H207
    # = "Kiểm chứng bằng sự việc thực tế trước khi tin."
    return 'Kiểm chứng bằng sự việc thực tế trước khi tin.'

def f536(E, r, c):
    # Ô I201, I202, I203, I204, I205, I206, I207
    # = "Thấp – Trung bình (ghi ở đầu dòng)"
    return 'Thấp – Trung bình (ghi ở đầu dòng)'

def f537(E, r, c):
    # Ô C202
    # = LEFT(KET_LUAN!$C$122,300)
    return X.LEFT(E.v(3,122,3), 300.0)

def f538(E, r, c):
    # Ô D202
    # = LEFT(KET_LUAN!$C$122,700)
    return X.LEFT(E.v(3,122,3), 700.0)

def f539(E, r, c):
    # Ô C203
    # = LEFT(KET_LUAN!$C$123,300)
    return X.LEFT(E.v(3,123,3), 300.0)

def f540(E, r, c):
    # Ô D203
    # = LEFT(KET_LUAN!$C$123,700)
    return X.LEFT(E.v(3,123,3), 700.0)

def f541(E, r, c):
    # Ô C204
    # = LEFT(KET_LUAN!$C$124,300)
    return X.LEFT(E.v(3,124,3), 300.0)

def f542(E, r, c):
    # Ô D204
    # = LEFT(KET_LUAN!$C$124,700)
    return X.LEFT(E.v(3,124,3), 700.0)

def f543(E, r, c):
    # Ô C205
    # = LEFT(KET_LUAN!$C$100,300)
    return X.LEFT(E.v(3,100,3), 300.0)

def f544(E, r, c):
    # Ô D205
    # = LEFT(KET_LUAN!$C$100,700)
    return X.LEFT(E.v(3,100,3), 700.0)

def f545(E, r, c):
    # Ô C206
    # = LEFT(KET_LUAN!$C$145,300)
    return X.LEFT(E.v(3,145,3), 300.0)

def f546(E, r, c):
    # Ô D206
    # = LEFT(KET_LUAN!$C$145,700)
    return X.LEFT(E.v(3,145,3), 700.0)

def f547(E, r, c):
    # Ô C207
    # = LEFT(KET_LUAN!$C$89,300)
    return X.LEFT(E.v(3,89,3), 300.0)

def f548(E, r, c):
    # Ô D207
    # = LEFT(KET_LUAN!$C$89,700)
    return X.LEFT(E.v(3,89,3), 700.0)

def f549(E, r, c):
    # Ô C210
    # = IF(KET_LUAN!$C$76="","(không có)",KET_LUAN!$C$76)
    return ('(không có)' if X.truthy(X.eq(E.v(3,76,3),'')) else E.v(3,76,3))

def f550(E, r, c):
    # Ô C211
    # = IF(KET_LUAN!$C$101="","(không có)",KET_LUAN!$C$101)
    return ('(không có)' if X.truthy(X.eq(E.v(3,101,3),'')) else E.v(3,101,3))

def f551(E, r, c):
    # Ô C212
    # = IF(KET_LUAN!$C$90="","(không có)",KET_LUAN!$C$90)
    return ('(không có)' if X.truthy(X.eq(E.v(3,90,3),'')) else E.v(3,90,3))

def f552(E, r, c):
    # Ô C213
    # = IF(KET_LUAN!$C$125="","(không có)",KET_LUAN!$C$125)
    return ('(không có)' if X.truthy(X.eq(E.v(3,125,3),'')) else E.v(3,125,3))

def f553(E, r, c):
    # Ô C214
    # = IF(KET_LUAN!$C$126="","(không có)",KET_LUAN!$C$126)
    return ('(không có)' if X.truthy(X.eq(E.v(3,126,3),'')) else E.v(3,126,3))

def f554(E, r, c):
    # Ô C215
    # = IF(KET_LUAN!$C$127="","(không có)",KET_LUAN!$C$127)
    return ('(không có)' if X.truthy(X.eq(E.v(3,127,3),'')) else E.v(3,127,3))

def f555(E, r, c):
    # Ô C216
    # = IF(KET_LUAN!$C$174="","(không có)",KET_LUAN!$C$174)
    return ('(không có)' if X.truthy(X.eq(E.v(3,174,3),'')) else E.v(3,174,3))

def f556(E, r, c):
    # Ô C217
    # = IF(KET_LUAN!$C$65="","(không có)",KET_LUAN!$C$65)
    return ('(không có)' if X.truthy(X.eq(E.v(3,65,3),'')) else E.v(3,65,3))

def f557(E, r, c):
    # Ô C220
    # = IF(KET_LUAN!$C$128="","(không có)",KET_LUAN!$C$128)
    return ('(không có)' if X.truthy(X.eq(E.v(3,128,3),'')) else E.v(3,128,3))

def f558(E, r, c):
    # Ô C221
    # = IF(KET_LUAN!$C$129="","(không có)",KET_LUAN!$C$129)
    return ('(không có)' if X.truthy(X.eq(E.v(3,129,3),'')) else E.v(3,129,3))

def f559(E, r, c):
    # Ô C222
    # = IF(KET_LUAN!$C$130="","(không có)",KET_LUAN!$C$130)
    return ('(không có)' if X.truthy(X.eq(E.v(3,130,3),'')) else E.v(3,130,3))

def f560(E, r, c):
    # Ô C223
    # = IF(KET_LUAN!$C$156="","(không có)",KET_LUAN!$C$156)
    return ('(không có)' if X.truthy(X.eq(E.v(3,156,3),'')) else E.v(3,156,3))

def f561(E, r, c):
    # Ô C224
    # = IF(KET_LUAN!$C$102="","(không có)",KET_LUAN!$C$102)
    return ('(không có)' if X.truthy(X.eq(E.v(3,102,3),'')) else E.v(3,102,3))

def f562(E, r, c):
    # Ô C225
    # = IF(KET_LUAN!$C$103="","(không có)",KET_LUAN!$C$103)
    return ('(không có)' if X.truthy(X.eq(E.v(3,103,3),'')) else E.v(3,103,3))

def f563(E, r, c):
    # Ô C226
    # = IF(KET_LUAN!$C$104="","(không có)",KET_LUAN!$C$104)
    return ('(không có)' if X.truthy(X.eq(E.v(3,104,3),'')) else E.v(3,104,3))

def f564(E, r, c):
    # Ô C227
    # = IF(KET_LUAN!$C$175="","(không có)",KET_LUAN!$C$175)
    return ('(không có)' if X.truthy(X.eq(E.v(3,175,3),'')) else E.v(3,175,3))

def f565(E, r, c):
    # Ô C228
    # = IF(KET_LUAN!$C$176="","(không có)",KET_LUAN!$C$176)
    return ('(không có)' if X.truthy(X.eq(E.v(3,176,3),'')) else E.v(3,176,3))

CELLS = {
    'S5': (f0, 5, 19),
    'S6': (f1, 6, 19),
    'T6': (f2, 6, 20),
    'S7': (f3, 7, 19),
    'T7': (f2, 7, 20),
    'C8': (f4, 8, 3),
    'S8': (f5, 8, 19),
    'T8': (f2, 8, 20),
    'C9': (f6, 9, 3),
    'J9': (f7, 9, 10),
    'S9': (f8, 9, 19),
    'T9': (f2, 9, 20),
    'C10': (f9, 10, 3),
    'S10': (f10, 10, 19),
    'T10': (f2, 10, 20),
    'C11': (f11, 11, 3),
    'C12': (f12, 12, 3),
    'S12': (f13, 12, 19),
    'C13': (f14, 13, 3),
    'C14': (f15, 14, 3),
    'C15': (f16, 15, 3),
    'C17': (f17, 17, 3),
    'C18': (f6, 18, 3),
    'J18': (f18, 18, 10),
    'C19': (f19, 19, 3),
    'C20': (f20, 20, 3),
    'C21': (f21, 21, 3),
    'C22': (f22, 22, 3),
    'C24': (f23, 24, 3),
    'C25': (f6, 25, 3),
    'J25': (f24, 25, 10),
    'C26': (f25, 26, 3),
    'C27': (f26, 27, 3),
    'C28': (f27, 28, 3),
    'C29': (f28, 29, 3),
    'C30': (f29, 30, 3),
    'C31': (f30, 31, 3),
    'C32': (f31, 32, 3),
    'C33': (f32, 33, 3),
    'C35': (f33, 35, 3),
    'C36': (f6, 36, 3),
    'J36': (f34, 36, 10),
    'C37': (f35, 37, 3),
    'C38': (f36, 38, 3),
    'C39': (f37, 39, 3),
    'C40': (f38, 40, 3),
    'C41': (f39, 41, 3),
    'C42': (f40, 42, 3),
    'C44': (f41, 44, 3),
    'C45': (f6, 45, 3),
    'J45': (f42, 45, 10),
    'C46': (f43, 46, 3),
    'C47': (f44, 47, 3),
    'C48': (f45, 48, 3),
    'C49': (f46, 49, 3),
    'C50': (f47, 50, 3),
    'C51': (f48, 51, 3),
    'C52': (f49, 52, 3),
    'C53': (f50, 53, 3),
    'C54': (f51, 54, 3),
    'C56': (f52, 56, 3),
    'C57': (f6, 57, 3),
    'J57': (f53, 57, 10),
    'C58': (f54, 58, 3),
    'C59': (f55, 59, 3),
    'C60': (f56, 60, 3),
    'C61': (f57, 61, 3),
    'C62': (f58, 62, 3),
    'C63': (f59, 63, 3),
    'C64': (f60, 64, 3),
    'C66': (f61, 66, 3),
    'C67': (f6, 67, 3),
    'J67': (f62, 67, 10),
    'C68': (f63, 68, 3),
    'C69': (f64, 69, 3),
    'A71': (f65, 71, 1),
    'B71': (f66, 71, 2),
    'C71': (f67, 71, 3),
    'D71': (f68, 71, 4),
    'E71': (f69, 71, 5),
    'F71': (f70, 71, 6),
    'G71': (f71, 71, 7),
    'H71': (f72, 71, 8),
    'I71': (f73, 71, 9),
    'K71': (f74, 71, 11),
    'L71': (f75, 71, 12),
    'A72': (f76, 72, 1),
    'B72': (f66, 72, 2),
    'C72': (f67, 72, 3),
    'D72': (f68, 72, 4),
    'E72': (f69, 72, 5),
    'F72': (f70, 72, 6),
    'G72': (f71, 72, 7),
    'H72': (f72, 72, 8),
    'I72': (f73, 72, 9),
    'K72': (f77, 72, 11),
    'L72': (f75, 72, 12),
    'A73': (f78, 73, 1),
    'B73': (f66, 73, 2),
    'C73': (f67, 73, 3),
    'D73': (f68, 73, 4),
    'E73': (f69, 73, 5),
    'F73': (f70, 73, 6),
    'G73': (f71, 73, 7),
    'H73': (f72, 73, 8),
    'I73': (f73, 73, 9),
    'K73': (f79, 73, 11),
    'L73': (f75, 73, 12),
    'A74': (f80, 74, 1),
    'B74': (f66, 74, 2),
    'C74': (f67, 74, 3),
    'D74': (f68, 74, 4),
    'E74': (f69, 74, 5),
    'F74': (f70, 74, 6),
    'G74': (f71, 74, 7),
    'H74': (f72, 74, 8),
    'I74': (f73, 74, 9),
    'K74': (f81, 74, 11),
    'L74': (f75, 74, 12),
    'A75': (f82, 75, 1),
    'B75': (f66, 75, 2),
    'C75': (f67, 75, 3),
    'D75': (f68, 75, 4),
    'E75': (f69, 75, 5),
    'F75': (f70, 75, 6),
    'G75': (f71, 75, 7),
    'H75': (f72, 75, 8),
    'I75': (f73, 75, 9),
    'K75': (f83, 75, 11),
    'L75': (f75, 75, 12),
    'A76': (f84, 76, 1),
    'B76': (f66, 76, 2),
    'C76': (f67, 76, 3),
    'D76': (f68, 76, 4),
    'E76': (f69, 76, 5),
    'F76': (f70, 76, 6),
    'G76': (f71, 76, 7),
    'H76': (f72, 76, 8),
    'I76': (f73, 76, 9),
    'K76': (f85, 76, 11),
    'L76': (f75, 76, 12),
    'A77': (f86, 77, 1),
    'B77': (f66, 77, 2),
    'C77': (f67, 77, 3),
    'D77': (f68, 77, 4),
    'E77': (f69, 77, 5),
    'F77': (f70, 77, 6),
    'G77': (f71, 77, 7),
    'H77': (f72, 77, 8),
    'I77': (f73, 77, 9),
    'K77': (f87, 77, 11),
    'L77': (f75, 77, 12),
    'A78': (f88, 78, 1),
    'B78': (f66, 78, 2),
    'C78': (f67, 78, 3),
    'D78': (f68, 78, 4),
    'E78': (f69, 78, 5),
    'F78': (f70, 78, 6),
    'G78': (f71, 78, 7),
    'H78': (f72, 78, 8),
    'I78': (f73, 78, 9),
    'K78': (f89, 78, 11),
    'L78': (f75, 78, 12),
    'A79': (f90, 79, 1),
    'B79': (f66, 79, 2),
    'C79': (f67, 79, 3),
    'D79': (f68, 79, 4),
    'E79': (f69, 79, 5),
    'F79': (f70, 79, 6),
    'G79': (f71, 79, 7),
    'H79': (f72, 79, 8),
    'I79': (f73, 79, 9),
    'K79': (f91, 79, 11),
    'L79': (f75, 79, 12),
    'A80': (f92, 80, 1),
    'B80': (f66, 80, 2),
    'C80': (f67, 80, 3),
    'D80': (f68, 80, 4),
    'E80': (f69, 80, 5),
    'F80': (f70, 80, 6),
    'G80': (f71, 80, 7),
    'H80': (f72, 80, 8),
    'I80': (f73, 80, 9),
    'K80': (f93, 80, 11),
    'L80': (f75, 80, 12),
    'A81': (f94, 81, 1),
    'B81': (f66, 81, 2),
    'C81': (f67, 81, 3),
    'D81': (f68, 81, 4),
    'E81': (f69, 81, 5),
    'F81': (f70, 81, 6),
    'G81': (f71, 81, 7),
    'H81': (f72, 81, 8),
    'I81': (f73, 81, 9),
    'K81': (f95, 81, 11),
    'L81': (f75, 81, 12),
    'C83': (f96, 83, 3),
    'C84': (f6, 84, 3),
    'J84': (f97, 84, 10),
    'C85': (f98, 85, 3),
    'C86': (f99, 86, 3),
    'C87': (f100, 87, 3),
    'C88': (f101, 88, 3),
    'C89': (f102, 89, 3),
    'C91': (f103, 91, 3),
    'C92': (f6, 92, 3),
    'J92': (f104, 92, 10),
    'C93': (f105, 93, 3),
    'C94': (f106, 94, 3),
    'C95': (f107, 95, 3),
    'C96': (f108, 96, 3),
    'C97': (f109, 97, 3),
    'C98': (f110, 98, 3),
    'C106': (f111, 106, 3),
    'D106': (f112, 106, 4),
    'E106': (f113, 106, 5),
    'F106': (f114, 106, 6),
    'G106': (f115, 106, 7),
    'H106': (f116, 106, 8),
    'I106': (f117, 106, 9),
    'J106': (f118, 106, 10),
    'C107': (f119, 107, 3),
    'D107': (f120, 107, 4),
    'E107': (f121, 107, 5),
    'F107': (f122, 107, 6),
    'G107': (f115, 107, 7),
    'H107': (f123, 107, 8),
    'I107': (f124, 107, 9),
    'J107': (f118, 107, 10),
    'C108': (f125, 108, 3),
    'D108': (f126, 108, 4),
    'E108': (f127, 108, 5),
    'F108': (f128, 108, 6),
    'G108': (f115, 108, 7),
    'H108': (f129, 108, 8),
    'I108': (f130, 108, 9),
    'J108': (f118, 108, 10),
    'C109': (f131, 109, 3),
    'D109': (f132, 109, 4),
    'E109': (f133, 109, 5),
    'F109': (f134, 109, 6),
    'G109': (f135, 109, 7),
    'H109': (f136, 109, 8),
    'I109': (f137, 109, 9),
    'J109': (f118, 109, 10),
    'C110': (f138, 110, 3),
    'D110': (f139, 110, 4),
    'E110': (f140, 110, 5),
    'F110': (f141, 110, 6),
    'G110': (f142, 110, 7),
    'H110': (f143, 110, 8),
    'I110': (f144, 110, 9),
    'J110': (f118, 110, 10),
    'C111': (f145, 111, 3),
    'D111': (f146, 111, 4),
    'E111': (f147, 111, 5),
    'F111': (f148, 111, 6),
    'G111': (f149, 111, 7),
    'H111': (f143, 111, 8),
    'I111': (f150, 111, 9),
    'J111': (f118, 111, 10),
    'C112': (f151, 112, 3),
    'D112': (f152, 112, 4),
    'E112': (f153, 112, 5),
    'F112': (f154, 112, 6),
    'G112': (f155, 112, 7),
    'H112': (f143, 112, 8),
    'I112': (f156, 112, 9),
    'J112': (f118, 112, 10),
    'C113': (f157, 113, 3),
    'D113': (f158, 113, 4),
    'E113': (f159, 113, 5),
    'F113': (f160, 113, 6),
    'G113': (f161, 113, 7),
    'H113': (f143, 113, 8),
    'I113': (f162, 113, 9),
    'J113': (f118, 113, 10),
    'C114': (f163, 114, 3),
    'D114': (f164, 114, 4),
    'E114': (f165, 114, 5),
    'F114': (f166, 114, 6),
    'G114': (f167, 114, 7),
    'H114': (f143, 114, 8),
    'I114': (f168, 114, 9),
    'J114': (f118, 114, 10),
    'C115': (f169, 115, 3),
    'D115': (f170, 115, 4),
    'E115': (f171, 115, 5),
    'F115': (f172, 115, 6),
    'G115': (f173, 115, 7),
    'H115': (f174, 115, 8),
    'I115': (f175, 115, 9),
    'J115': (f118, 115, 10),
    'C116': (f176, 116, 3),
    'D116': (f177, 116, 4),
    'E116': (f178, 116, 5),
    'F116': (f179, 116, 6),
    'G116': (f115, 116, 7),
    'H116': (f180, 116, 8),
    'I116': (f181, 116, 9),
    'J116': (f118, 116, 10),
    'C117': (f182, 117, 3),
    'D117': (f183, 117, 4),
    'E117': (f178, 117, 5),
    'F117': (f184, 117, 6),
    'G117': (f115, 117, 7),
    'H117': (f180, 117, 8),
    'I117': (f185, 117, 9),
    'J117': (f118, 117, 10),
    'C118': (f186, 118, 3),
    'D118': (f187, 118, 4),
    'E118': (f178, 118, 5),
    'F118': (f188, 118, 6),
    'G118': (f115, 118, 7),
    'H118': (f180, 118, 8),
    'I118': (f189, 118, 9),
    'J118': (f118, 118, 10),
    'C119': (f190, 119, 3),
    'D119': (f191, 119, 4),
    'E119': (f178, 119, 5),
    'F119': (f192, 119, 6),
    'G119': (f115, 119, 7),
    'H119': (f180, 119, 8),
    'I119': (f193, 119, 9),
    'J119': (f118, 119, 10),
    'C120': (f194, 120, 3),
    'D120': (f195, 120, 4),
    'E120': (f196, 120, 5),
    'F120': (f197, 120, 6),
    'G120': (f198, 120, 7),
    'H120': (f199, 120, 8),
    'I120': (f200, 120, 9),
    'J120': (f118, 120, 10),
    'C121': (f201, 121, 3),
    'D121': (f202, 121, 4),
    'E121': (f203, 121, 5),
    'F121': (f204, 121, 6),
    'G121': (f205, 121, 7),
    'H121': (f206, 121, 8),
    'I121': (f130, 121, 9),
    'J121': (f118, 121, 10),
    'C122': (f207, 122, 3),
    'D122': (f208, 122, 4),
    'E122': (f209, 122, 5),
    'F122': (f210, 122, 6),
    'G122': (f211, 122, 7),
    'H122': (f212, 122, 8),
    'I122': (f213, 122, 9),
    'J122': (f118, 122, 10),
    'C123': (f214, 123, 3),
    'D123': (f215, 123, 4),
    'E123': (f216, 123, 5),
    'F123': (f217, 123, 6),
    'G123': (f115, 123, 7),
    'H123': (f218, 123, 8),
    'I123': (f200, 123, 9),
    'J123': (f118, 123, 10),
    'C124': (f219, 124, 3),
    'D124': (f220, 124, 4),
    'E124': (f221, 124, 5),
    'F124': (f222, 124, 6),
    'G124': (f115, 124, 7),
    'H124': (f218, 124, 8),
    'I124': (f223, 124, 9),
    'J124': (f118, 124, 10),
    'C125': (f224, 125, 3),
    'D125': (f225, 125, 4),
    'E125': (f226, 125, 5),
    'F125': (f227, 125, 6),
    'G125': (f115, 125, 7),
    'H125': (f218, 125, 8),
    'I125': (f228, 125, 9),
    'J125': (f118, 125, 10),
    'C126': (f229, 126, 3),
    'D126': (f230, 126, 4),
    'E126': (f231, 126, 5),
    'F126': (f232, 126, 6),
    'G126': (f115, 126, 7),
    'H126': (f218, 126, 8),
    'I126': (f233, 126, 9),
    'J126': (f118, 126, 10),
    'C127': (f234, 127, 3),
    'D127': (f235, 127, 4),
    'E127': (f236, 127, 5),
    'F127': (f237, 127, 6),
    'G127': (f238, 127, 7),
    'H127': (f239, 127, 8),
    'I127': (f240, 127, 9),
    'J127': (f118, 127, 10),
    'C128': (f241, 128, 3),
    'D128': (f242, 128, 4),
    'E128': (f236, 128, 5),
    'F128': (f243, 128, 6),
    'G128': (f238, 128, 7),
    'H128': (f239, 128, 8),
    'I128': (f244, 128, 9),
    'J128': (f118, 128, 10),
    'C129': (f245, 129, 3),
    'D129': (f246, 129, 4),
    'E129': (f236, 129, 5),
    'F129': (f247, 129, 6),
    'G129': (f238, 129, 7),
    'H129': (f239, 129, 8),
    'I129': (f248, 129, 9),
    'J129': (f118, 129, 10),
    'C130': (f249, 130, 3),
    'D130': (f250, 130, 4),
    'E130': (f251, 130, 5),
    'F130': (f252, 130, 6),
    'G130': (f253, 130, 7),
    'H130': (f254, 130, 8),
    'I130': (f130, 130, 9),
    'J130': (f118, 130, 10),
    'C133': (f255, 133, 3),
    'C134': (f256, 134, 3),
    'C135': (f257, 135, 3),
    'C136': (f258, 136, 3),
    'C137': (f259, 137, 3),
    'C138': (f260, 138, 3),
    'C139': (f261, 139, 3),
    'C140': (f262, 140, 3),
    'C144': (f263, 144, 3),
    'D144': (f264, 144, 4),
    'E144': (f265, 144, 5),
    'F144': (f266, 144, 6),
    'G144': (f115, 144, 7),
    'H144': (f267, 144, 8),
    'I144': (f130, 144, 9),
    'J144': (f268, 144, 10),
    'C145': (f269, 145, 3),
    'D145': (f270, 145, 4),
    'E145': (f271, 145, 5),
    'F145': (f272, 145, 6),
    'G145': (f273, 145, 7),
    'H145': (f274, 145, 8),
    'I145': (f130, 145, 9),
    'J145': (f268, 145, 10),
    'C146': (f275, 146, 3),
    'D146': (f276, 146, 4),
    'E146': (f277, 146, 5),
    'F146': (f278, 146, 6),
    'G146': (f273, 146, 7),
    'H146': (f279, 146, 8),
    'I146': (f130, 146, 9),
    'J146': (f268, 146, 10),
    'C147': (f280, 147, 3),
    'D147': (f281, 147, 4),
    'E147': (f282, 147, 5),
    'F147': (f283, 147, 6),
    'G147': (f115, 147, 7),
    'H147': (f284, 147, 8),
    'I147': (f130, 147, 9),
    'J147': (f268, 147, 10),
    'C148': (f285, 148, 3),
    'D148': (f286, 148, 4),
    'E148': (f287, 148, 5),
    'F148': (f288, 148, 6),
    'G148': (f115, 148, 7),
    'H148': (f284, 148, 8),
    'I148': (f130, 148, 9),
    'J148': (f268, 148, 10),
    'C149': (f289, 149, 3),
    'D149': (f290, 149, 4),
    'E149': (f291, 149, 5),
    'F149': (f292, 149, 6),
    'G149': (f115, 149, 7),
    'H149': (f284, 149, 8),
    'I149': (f130, 149, 9),
    'J149': (f268, 149, 10),
    'C150': (f293, 150, 3),
    'D150': (f294, 150, 4),
    'E150': (f295, 150, 5),
    'F150': (f296, 150, 6),
    'G150': (f297, 150, 7),
    'H150': (f298, 150, 8),
    'I150': (f299, 150, 9),
    'J150': (f268, 150, 10),
    'C151': (f300, 151, 3),
    'D151': (f301, 151, 4),
    'E151': (f302, 151, 5),
    'F151': (f303, 151, 6),
    'G151': (f304, 151, 7),
    'H151': (f305, 151, 8),
    'I151': (f299, 151, 9),
    'J151': (f268, 151, 10),
    'C152': (f306, 152, 3),
    'D152': (f307, 152, 4),
    'E152': (f308, 152, 5),
    'F152': (f266, 152, 6),
    'G152': (f309, 152, 7),
    'H152': (f310, 152, 8),
    'I152': (f130, 152, 9),
    'J152': (f268, 152, 10),
    'C153': (f311, 153, 3),
    'D153': (f312, 153, 4),
    'E153': (f313, 153, 5),
    'F153': (f266, 153, 6),
    'G153': (f115, 153, 7),
    'H153': (f314, 153, 8),
    'I153': (f130, 153, 9),
    'J153': (f268, 153, 10),
    'C154': (f315, 154, 3),
    'D154': (f316, 154, 4),
    'E154': (f317, 154, 5),
    'F154': (f318, 154, 6),
    'G154': (f319, 154, 7),
    'H154': (f320, 154, 8),
    'I154': (f321, 154, 9),
    'J154': (f268, 154, 10),
    'C155': (f322, 155, 3),
    'D155': (f323, 155, 4),
    'E155': (f324, 155, 5),
    'F155': (f325, 155, 6),
    'G155': (f326, 155, 7),
    'H155': (f284, 155, 8),
    'I155': (f130, 155, 9),
    'J155': (f268, 155, 10),
    'C156': (f327, 156, 3),
    'D156': (f328, 156, 4),
    'E156': (f324, 156, 5),
    'F156': (f325, 156, 6),
    'G156': (f326, 156, 7),
    'H156': (f284, 156, 8),
    'I156': (f130, 156, 9),
    'J156': (f268, 156, 10),
    'C157': (f329, 157, 3),
    'D157': (f330, 157, 4),
    'E157': (f324, 157, 5),
    'F157': (f325, 157, 6),
    'G157': (f326, 157, 7),
    'H157': (f284, 157, 8),
    'I157': (f130, 157, 9),
    'J157': (f268, 157, 10),
    'C158': (f331, 158, 3),
    'D158': (f332, 158, 4),
    'E158': (f324, 158, 5),
    'F158': (f325, 158, 6),
    'G158': (f326, 158, 7),
    'H158': (f284, 158, 8),
    'I158': (f130, 158, 9),
    'J158': (f268, 158, 10),
    'C159': (f333, 159, 3),
    'D159': (f334, 159, 4),
    'E159': (f335, 159, 5),
    'F159': (f336, 159, 6),
    'G159': (f326, 159, 7),
    'H159': (f284, 159, 8),
    'I159': (f130, 159, 9),
    'J159': (f268, 159, 10),
    'C160': (f337, 160, 3),
    'D160': (f338, 160, 4),
    'E160': (f339, 160, 5),
    'F160': (f340, 160, 6),
    'G160': (f341, 160, 7),
    'H160': (f342, 160, 8),
    'I160': (f343, 160, 9),
    'J160': (f268, 160, 10),
    'C161': (f344, 161, 3),
    'D161': (f345, 161, 4),
    'E161': (f346, 161, 5),
    'F161': (f347, 161, 6),
    'G161': (f348, 161, 7),
    'H161': (f349, 161, 8),
    'I161': (f343, 161, 9),
    'J161': (f268, 161, 10),
    'C162': (f350, 162, 3),
    'D162': (f351, 162, 4),
    'E162': (f346, 162, 5),
    'F162': (f352, 162, 6),
    'G162': (f348, 162, 7),
    'H162': (f349, 162, 8),
    'I162': (f343, 162, 9),
    'J162': (f268, 162, 10),
    'C163': (f353, 163, 3),
    'D163': (f354, 163, 4),
    'E163': (f346, 163, 5),
    'F163': (f355, 163, 6),
    'G163': (f348, 163, 7),
    'H163': (f349, 163, 8),
    'I163': (f343, 163, 9),
    'J163': (f268, 163, 10),
    'C164': (f356, 164, 3),
    'D164': (f357, 164, 4),
    'E164': (f346, 164, 5),
    'F164': (f358, 164, 6),
    'G164': (f348, 164, 7),
    'H164': (f349, 164, 8),
    'I164': (f343, 164, 9),
    'J164': (f268, 164, 10),
    'C165': (f359, 165, 3),
    'D165': (f360, 165, 4),
    'E165': (f361, 165, 5),
    'F165': (f362, 165, 6),
    'G165': (f363, 165, 7),
    'H165': (f364, 165, 8),
    'I165': (f365, 165, 9),
    'J165': (f268, 165, 10),
    'C166': (f366, 166, 3),
    'D166': (f367, 166, 4),
    'E166': (f368, 166, 5),
    'F166': (f369, 166, 6),
    'G166': (f370, 166, 7),
    'H166': (f284, 166, 8),
    'I166': (f130, 166, 9),
    'J166': (f268, 166, 10),
    'C167': (f371, 167, 3),
    'D167': (f372, 167, 4),
    'E167': (f373, 167, 5),
    'F167': (f374, 167, 6),
    'G167': (f375, 167, 7),
    'H167': (f376, 167, 8),
    'I167': (f130, 167, 9),
    'J167': (f268, 167, 10),
    'C168': (f377, 168, 3),
    'D168': (f378, 168, 4),
    'E168': (f373, 168, 5),
    'F168': (f379, 168, 6),
    'G168': (f380, 168, 7),
    'H168': (f381, 168, 8),
    'I168': (f130, 168, 9),
    'J168': (f268, 168, 10),
    'C169': (f382, 169, 3),
    'D169': (f383, 169, 4),
    'E169': (f373, 169, 5),
    'F169': (f384, 169, 6),
    'G169': (f385, 169, 7),
    'H169': (f386, 169, 8),
    'I169': (f130, 169, 9),
    'J169': (f268, 169, 10),
    'C170': (f387, 170, 3),
    'D170': (f388, 170, 4),
    'E170': (f373, 170, 5),
    'F170': (f389, 170, 6),
    'G170': (f390, 170, 7),
    'H170': (f391, 170, 8),
    'I170': (f130, 170, 9),
    'J170': (f268, 170, 10),
    'C171': (f392, 171, 3),
    'D171': (f393, 171, 4),
    'E171': (f373, 171, 5),
    'F171': (f394, 171, 6),
    'G171': (f395, 171, 7),
    'H171': (f396, 171, 8),
    'I171': (f130, 171, 9),
    'J171': (f268, 171, 10),
    'C172': (f397, 172, 3),
    'D172': (f398, 172, 4),
    'E172': (f373, 172, 5),
    'F172': (f399, 172, 6),
    'G172': (f400, 172, 7),
    'H172': (f401, 172, 8),
    'I172': (f130, 172, 9),
    'J172': (f268, 172, 10),
    'C173': (f402, 173, 3),
    'D173': (f403, 173, 4),
    'E173': (f404, 173, 5),
    'F173': (f405, 173, 6),
    'G173': (f406, 173, 7),
    'H173': (f284, 173, 8),
    'I173': (f137, 173, 9),
    'J173': (f268, 173, 10),
    'C174': (f407, 174, 3),
    'D174': (f408, 174, 4),
    'E174': (f409, 174, 5),
    'F174': (f410, 174, 6),
    'G174': (f411, 174, 7),
    'H174': (f412, 174, 8),
    'I174': (f343, 174, 9),
    'J174': (f268, 174, 10),
    'C175': (f413, 175, 3),
    'D175': (f414, 175, 4),
    'E175': (f415, 175, 5),
    'F175': (f416, 175, 6),
    'G175': (f417, 175, 7),
    'H175': (f418, 175, 8),
    'I175': (f343, 175, 9),
    'J175': (f268, 175, 10),
    'C176': (f419, 176, 3),
    'D176': (f420, 176, 4),
    'E176': (f421, 176, 5),
    'F176': (f422, 176, 6),
    'G176': (f115, 176, 7),
    'H176': (f423, 176, 8),
    'I176': (f130, 176, 9),
    'J176': (f268, 176, 10),
    'C177': (f424, 177, 3),
    'D177': (f425, 177, 4),
    'E177': (f426, 177, 5),
    'F177': (f427, 177, 6),
    'G177': (f428, 177, 7),
    'H177': (f429, 177, 8),
    'I177': (f343, 177, 9),
    'J177': (f268, 177, 10),
    'C178': (f430, 178, 3),
    'D178': (f431, 178, 4),
    'E178': (f432, 178, 5),
    'F178': (f433, 178, 6),
    'G178': (f434, 178, 7),
    'H178': (f284, 178, 8),
    'I178': (f435, 178, 9),
    'J178': (f268, 178, 10),
    'C179': (f436, 179, 3),
    'D179': (f437, 179, 4),
    'E179': (f438, 179, 5),
    'F179': (f439, 179, 6),
    'G179': (f115, 179, 7),
    'H179': (f440, 179, 8),
    'I179': (f343, 179, 9),
    'J179': (f268, 179, 10),
    'C180': (f441, 180, 3),
    'D180': (f442, 180, 4),
    'E180': (f443, 180, 5),
    'F180': (f444, 180, 6),
    'G180': (f445, 180, 7),
    'H180': (f446, 180, 8),
    'I180': (f447, 180, 9),
    'J180': (f268, 180, 10),
    'C181': (f448, 181, 3),
    'D181': (f449, 181, 4),
    'E181': (f450, 181, 5),
    'F181': (f451, 181, 6),
    'G181': (f115, 181, 7),
    'H181': (f452, 181, 8),
    'I181': (f130, 181, 9),
    'J181': (f268, 181, 10),
    'C182': (f453, 182, 3),
    'D182': (f454, 182, 4),
    'E182': (f455, 182, 5),
    'F182': (f456, 182, 6),
    'G182': (f115, 182, 7),
    'H182': (f457, 182, 8),
    'I182': (f200, 182, 9),
    'J182': (f268, 182, 10),
    'C183': (f458, 183, 3),
    'D183': (f459, 183, 4),
    'E183': (f460, 183, 5),
    'F183': (f461, 183, 6),
    'G183': (f115, 183, 7),
    'H183': (f457, 183, 8),
    'I183': (f200, 183, 9),
    'J183': (f268, 183, 10),
    'C184': (f462, 184, 3),
    'D184': (f463, 184, 4),
    'E184': (f464, 184, 5),
    'F184': (f465, 184, 6),
    'G184': (f115, 184, 7),
    'H184': (f457, 184, 8),
    'I184': (f200, 184, 9),
    'J184': (f268, 184, 10),
    'C185': (f466, 185, 3),
    'D185': (f467, 185, 4),
    'E185': (f468, 185, 5),
    'F185': (f469, 185, 6),
    'G185': (f115, 185, 7),
    'H185': (f457, 185, 8),
    'I185': (f200, 185, 9),
    'J185': (f268, 185, 10),
    'C186': (f470, 186, 3),
    'D186': (f471, 186, 4),
    'E186': (f472, 186, 5),
    'F186': (f473, 186, 6),
    'G186': (f474, 186, 7),
    'H186': (f457, 186, 8),
    'I186': (f200, 186, 9),
    'J186': (f268, 186, 10),
    'C187': (f475, 187, 3),
    'D187': (f476, 187, 4),
    'E187': (f477, 187, 5),
    'F187': (f478, 187, 6),
    'G187': (f479, 187, 7),
    'H187': (f480, 187, 8),
    'I187': (f343, 187, 9),
    'J187': (f268, 187, 10),
    'C188': (f481, 188, 3),
    'D188': (f482, 188, 4),
    'E188': (f483, 188, 5),
    'F188': (f484, 188, 6),
    'G188': (f479, 188, 7),
    'H188': (f480, 188, 8),
    'I188': (f343, 188, 9),
    'J188': (f268, 188, 10),
    'C189': (f485, 189, 3),
    'D189': (f486, 189, 4),
    'E189': (f487, 189, 5),
    'F189': (f488, 189, 6),
    'G189': (f479, 189, 7),
    'H189': (f480, 189, 8),
    'I189': (f343, 189, 9),
    'J189': (f268, 189, 10),
    'C190': (f489, 190, 3),
    'D190': (f490, 190, 4),
    'E190': (f491, 190, 5),
    'F190': (f492, 190, 6),
    'G190': (f479, 190, 7),
    'H190': (f480, 190, 8),
    'I190': (f343, 190, 9),
    'J190': (f268, 190, 10),
    'C191': (f493, 191, 3),
    'D191': (f494, 191, 4),
    'E191': (f495, 191, 5),
    'F191': (f496, 191, 6),
    'G191': (f115, 191, 7),
    'H191': (f497, 191, 8),
    'I191': (f223, 191, 9),
    'J191': (f268, 191, 10),
    'C192': (f498, 192, 3),
    'D192': (f499, 192, 4),
    'E192': (f500, 192, 5),
    'F192': (f501, 192, 6),
    'G192': (f502, 192, 7),
    'H192': (f503, 192, 8),
    'I192': (f343, 192, 9),
    'J192': (f268, 192, 10),
    'C193': (f504, 193, 3),
    'D193': (f505, 193, 4),
    'E193': (f506, 193, 5),
    'F193': (f507, 193, 6),
    'G193': (f508, 193, 7),
    'H193': (f509, 193, 8),
    'I193': (f510, 193, 9),
    'J193': (f268, 193, 10),
    'C194': (f511, 194, 3),
    'D194': (f512, 194, 4),
    'E194': (f513, 194, 5),
    'F194': (f514, 194, 6),
    'G194': (f238, 194, 7),
    'H194': (f515, 194, 8),
    'I194': (f240, 194, 9),
    'J194': (f268, 194, 10),
    'C195': (f516, 195, 3),
    'D195': (f517, 195, 4),
    'E195': (f518, 195, 5),
    'F195': (f519, 195, 6),
    'G195': (f115, 195, 7),
    'H195': (f284, 195, 8),
    'I195': (f130, 195, 9),
    'J195': (f268, 195, 10),
    'C196': (f520, 196, 3),
    'D196': (f521, 196, 4),
    'E196': (f518, 196, 5),
    'F196': (f519, 196, 6),
    'G196': (f115, 196, 7),
    'H196': (f284, 196, 8),
    'I196': (f130, 196, 9),
    'J196': (f268, 196, 10),
    'C197': (f522, 197, 3),
    'D197': (f523, 197, 4),
    'E197': (f518, 197, 5),
    'F197': (f519, 197, 6),
    'G197': (f115, 197, 7),
    'H197': (f284, 197, 8),
    'I197': (f130, 197, 9),
    'J197': (f268, 197, 10),
    'C198': (f524, 198, 3),
    'D198': (f525, 198, 4),
    'E198': (f518, 198, 5),
    'F198': (f519, 198, 6),
    'G198': (f115, 198, 7),
    'H198': (f284, 198, 8),
    'I198': (f130, 198, 9),
    'J198': (f268, 198, 10),
    'C199': (f526, 199, 3),
    'D199': (f527, 199, 4),
    'E199': (f518, 199, 5),
    'F199': (f519, 199, 6),
    'G199': (f115, 199, 7),
    'H199': (f284, 199, 8),
    'I199': (f130, 199, 9),
    'J199': (f268, 199, 10),
    'C200': (f528, 200, 3),
    'D200': (f529, 200, 4),
    'E200': (f518, 200, 5),
    'F200': (f519, 200, 6),
    'G200': (f115, 200, 7),
    'H200': (f284, 200, 8),
    'I200': (f130, 200, 9),
    'J200': (f268, 200, 10),
    'C201': (f530, 201, 3),
    'D201': (f531, 201, 4),
    'E201': (f532, 201, 5),
    'F201': (f533, 201, 6),
    'G201': (f534, 201, 7),
    'H201': (f535, 201, 8),
    'I201': (f536, 201, 9),
    'J201': (f268, 201, 10),
    'C202': (f537, 202, 3),
    'D202': (f538, 202, 4),
    'E202': (f532, 202, 5),
    'F202': (f533, 202, 6),
    'G202': (f534, 202, 7),
    'H202': (f535, 202, 8),
    'I202': (f536, 202, 9),
    'J202': (f268, 202, 10),
    'C203': (f539, 203, 3),
    'D203': (f540, 203, 4),
    'E203': (f532, 203, 5),
    'F203': (f533, 203, 6),
    'G203': (f534, 203, 7),
    'H203': (f535, 203, 8),
    'I203': (f536, 203, 9),
    'J203': (f268, 203, 10),
    'C204': (f541, 204, 3),
    'D204': (f542, 204, 4),
    'E204': (f532, 204, 5),
    'F204': (f533, 204, 6),
    'G204': (f534, 204, 7),
    'H204': (f535, 204, 8),
    'I204': (f536, 204, 9),
    'J204': (f268, 204, 10),
    'C205': (f543, 205, 3),
    'D205': (f544, 205, 4),
    'E205': (f532, 205, 5),
    'F205': (f533, 205, 6),
    'G205': (f534, 205, 7),
    'H205': (f535, 205, 8),
    'I205': (f536, 205, 9),
    'J205': (f268, 205, 10),
    'C206': (f545, 206, 3),
    'D206': (f546, 206, 4),
    'E206': (f532, 206, 5),
    'F206': (f533, 206, 6),
    'G206': (f534, 206, 7),
    'H206': (f535, 206, 8),
    'I206': (f536, 206, 9),
    'J206': (f268, 206, 10),
    'C207': (f547, 207, 3),
    'D207': (f548, 207, 4),
    'E207': (f532, 207, 5),
    'F207': (f533, 207, 6),
    'G207': (f534, 207, 7),
    'H207': (f535, 207, 8),
    'I207': (f536, 207, 9),
    'J207': (f268, 207, 10),
    'C210': (f549, 210, 3),
    'C211': (f550, 211, 3),
    'C212': (f551, 212, 3),
    'C213': (f552, 213, 3),
    'C214': (f553, 214, 3),
    'C215': (f554, 215, 3),
    'C216': (f555, 216, 3),
    'C217': (f556, 217, 3),
    'C220': (f557, 220, 3),
    'C221': (f558, 221, 3),
    'C222': (f559, 222, 3),
    'C223': (f560, 223, 3),
    'C224': (f561, 224, 3),
    'C225': (f562, 225, 3),
    'C226': (f563, 226, 3),
    'C227': (f564, 227, 3),
    'C228': (f565, 228, 3),
}
