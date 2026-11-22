# -*- coding: utf-8 -*-
"""Module sheet TAI_LOC – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.8.xlsx (SHA-256 b171b485540d6fb7f1b77d5397c9a0f6a2ddce344baf90abfa50bdfaaf431a5d).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'TAI_LOC'
SI = 19

def f0(E, r, c):
    # Ô C6
    # = INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+2,5)+1)
    return X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),2.0), 5.0),1.0))

def f1(E, r, c):
    # Ô D6
    # = PP1_DINH_TINH!$F$98
    return E.v(9,98,6)

def f2(E, r, c):
    # Ô E6
    # = "Hành tài = hành nhật chủ khắc: "&INDEX(HANH_TEN,MOD(DUNG_HY_KY!$D$5+2,5)+1)&" (PP2) / "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&" (PP1). PP1: tài 
    #   lộ "&PP1_DINH_TINH!$D$98&" chữ, tàng "&PP1_DINH_TINH!$E$98&" chữ."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Hành tài = hành nhật chủ khắc: ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(8,5,4),2.0), 5.0),1.0))),' (PP2) / '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),' (PP1). PP1: tài lộ '),E.v(9,98,4)),' chữ, tàng '),E.v(9,98,5)),' chữ.')

def f3(E, r, c):
    # Ô C7
    # = SUM(DUNG_HY_KY!$D$24:$D$28)
    return X.SUM(E.rg(8,24,4,28,4))

def f4(E, r, c):
    # Ô D7
    # = SUM(PP1_DINH_TINH!$F$96:$F$100)
    return X.SUM(E.rg(9,96,6,100,6))

def f5(E, r, c):
    # Ô E7
    # = "PP2 cộng độ năm hành (P2 tr.332-352); PP1 đếm chữ rõ + mờ (P2 tr.455-457). Cả hai không tính chữ nhật can, như cách tính lượng nhóm ở DUNG_HY_KY dòng
    #    12-16."
    return 'PP2 cộng độ năm hành (P2 tr.332-352); PP1 đếm chữ rõ + mờ (P2 tr.455-457). Cả hai không tính chữ nhật can, như cách tính lượng nhóm ở DUNG_HY_KY dòng 12-16.'

def f6(E, r, c):
    # Ô C8, D8
    # = IF(C7=0,0,C6/C7)
    return (0.0 if X.truthy(X.eq(E.v(19,r-1,c),0.0)) else X.div(E.v(19,r-2,c),E.v(19,r-1,c)))

def f7(E, r, c):
    # Ô E8
    # = "Số đo tham khảo: phần của sao tài trong toàn bộ khí của tứ trụ. Sách không cho ngưỡng nào để quy ra mức giàu."
    return 'Số đo tham khảo: phần của sao tài trong toàn bộ khí của tứ trụ. Sách không cho ngưỡng nào để quy ra mức giàu.'

def f8(E, r, c):
    # Ô C9
    # = IF(C7=0,0,(INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$5+1)+INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+4,5)+1))/C7)
    return (0.0 if X.truthy(X.eq(E.v(19,r-2,c),0.0)) else X.div(X.add(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,5,4),1.0)),X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),4.0), 5.0),1.0))),E.v(19,r-2,c)))

def f9(E, r, c):
    # Ô D9
    # = IF(D7=0,0,(PP1_DINH_TINH!$F$96+PP1_DINH_TINH!$F$100)/D7)
    return (0.0 if X.truthy(X.eq(E.v(19,r-2,c),0.0)) else X.div(X.add(E.v(9,96,6),E.v(9,100,6)),E.v(19,r-2,c)))

def f10(E, r, c):
    # Ô E9
    # = "Sức ""gánh tài"" của nhật chủ. Sách: ""người tài nhiều thì sợ thân nhược. Người thân mạnh có thể thắng được tài"" (P1 tr.95). Ngưỡng vượng PP2 = 40% 
    #   (P2 tr.352)."
    return 'Sức "gánh tài" của nhật chủ. Sách: "người tài nhiều thì sợ thân nhược. Người thân mạnh có thể thắng được tài" (P1 tr.95). Ngưỡng vượng PP2 = 40% (P2 tr.352).'

def f11(E, r, c):
    # Ô C10, D10
    # = IF(C9=0,"—",C8/C9)
    return ('—' if X.truthy(X.eq(E.v(19,r-1,c),0.0)) else X.div(E.v(19,r-2,c),E.v(19,r-1,c)))

def f12(E, r, c):
    # Ô E10
    # = "≈1: thân và tài ngang nhau – ""chỉ có thân và tài cân bằng mệnh mới tốt"" (P1 tr.95); ""thân và tài ngang nhau thì tài là dụng thần"" (P1 tr.87). >1:
    #    tài nặng hơn thân; <1: thân nặng hơn tài."
    return '≈1: thân và tài ngang nhau – "chỉ có thân và tài cân bằng mệnh mới tốt" (P1 tr.95); "thân và tài ngang nhau thì tài là dụng thần" (P1 tr.87). >1: tài nặng hơn thân; <1: thân nặng hơn tài.'

def f13(E, r, c):
    # Ô C11
    # = RANK(INDEX(DUNG_HY_KY!$D$24:$D$28,MOD(DUNG_HY_KY!$D$5+2,5)+1),DUNG_HY_KY!$D$24:$D$28)
    return X.RANK(X.INDEX(E.rg(8,24,4,28,4), X.add(X.MOD(X.add(E.v(8,5,4),2.0), 5.0),1.0)), E.rg(8,24,4,28,4))

def f14(E, r, c):
    # Ô D11
    # = RANK(PP1_DINH_TINH!$F$98,PP1_DINH_TINH!$F$96:$F$100)
    return X.RANK(E.v(9,98,6), E.rg(9,96,6,100,6))

def f15(E, r, c):
    # Ô E11
    # = "1 = hành nhiều nhất trong tứ trụ."
    return '1 = hành nhiều nhất trong tứ trụ.'

def f16(E, r, c):
    # Ô B12
    # = IF(DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5,"Lưu ý: PP2 coi nhật chủ đã hợp hóa thành "&INDEX(HANH_TEN,DUNG_HY_KY!$D$5+1)&" nên ""sao tài"" của PP2 là hành do
    #    hóa thần khắc.","")
    return (X.cat(X.cat('Lưu ý: PP2 coi nhật chủ đã hợp hóa thành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,5,4),1.0))),' nên "sao tài" của PP2 là hành do hóa thần khắc.') if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else '')

def f17(E, r, c):
    # Ô C16
    # = IF(DUNG_HY_KY!$C$72<>"",0,IF(OR(E48="ĐỦ",E49="ĐỦ"),1,IF(OR(E42="ĐỦ",E43="ĐỦ"),IF(E55="ĐỦ",4,5),IF(E44="ĐỦ",4,IF(OR(E52="ĐỦ",E53="ĐỦ",E54="ĐỦ"),2,IF(OR
    #   (E46="ĐỦ",E47="ĐỦ",E45="ĐỦ",E56="ĐỦ"),4,IF(E51="ĐỦ",2,3)))))))
    return (0.0 if X.truthy(X.ne(E.v(8,72,3),'')) else (1.0 if X.truthy(X.OR(X.eq(E.v(19,r+32,c+2),'ĐỦ'), X.eq(E.v(19,r+33,c+2),'ĐỦ'))) else ((4.0 if X.truthy(X.eq(E.v(19,r+39,c+2),'ĐỦ')) else 5.0) if X.truthy(X.OR(X.eq(E.v(19,r+26,c+2),'ĐỦ'), X.eq(E.v(19,r+27,c+2),'ĐỦ'))) else (4.0 if X.truthy(X.eq(E.v(19,r+28,c+2),'ĐỦ')) else (2.0 if X.truthy(X.OR(X.eq(E.v(19,r+36,c+2),'ĐỦ'), X.eq(E.v(19,r+37,c+2),'ĐỦ'), X.eq(E.v(19,r+38,c+2),'ĐỦ'))) else (4.0 if X.truthy(X.OR(X.eq(E.v(19,r+30,c+2),'ĐỦ'), X.eq(E.v(19,r+31,c+2),'ĐỦ'), X.eq(E.v(19,r+29,c+2),'ĐỦ'), X.eq(E.v(19,r+40,c+2),'ĐỦ'))) else (2.0 if X.truthy(X.eq(E.v(19,r+35,c+2),'ĐỦ')) else 3.0)))))))

def f18(E, r, c):
    # Ô D16
    # = IF(DUNG_HY_KY!$D$72<>"",0,IF(OR(F48="ĐỦ",F49="ĐỦ"),1,IF(OR(F42="ĐỦ",F43="ĐỦ"),IF(F55="ĐỦ",4,5),IF(F44="ĐỦ",4,IF(OR(F52="ĐỦ",F53="ĐỦ",F54="ĐỦ"),2,IF(OR
    #   (F46="ĐỦ",F47="ĐỦ",F45="ĐỦ",F56="ĐỦ"),4,IF(F51="ĐỦ",2,3)))))))
    return (0.0 if X.truthy(X.ne(E.v(8,72,4),'')) else (1.0 if X.truthy(X.OR(X.eq(E.v(19,r+32,c+2),'ĐỦ'), X.eq(E.v(19,r+33,c+2),'ĐỦ'))) else ((4.0 if X.truthy(X.eq(E.v(19,r+39,c+2),'ĐỦ')) else 5.0) if X.truthy(X.OR(X.eq(E.v(19,r+26,c+2),'ĐỦ'), X.eq(E.v(19,r+27,c+2),'ĐỦ'))) else (4.0 if X.truthy(X.eq(E.v(19,r+28,c+2),'ĐỦ')) else (2.0 if X.truthy(X.OR(X.eq(E.v(19,r+36,c+2),'ĐỦ'), X.eq(E.v(19,r+37,c+2),'ĐỦ'), X.eq(E.v(19,r+38,c+2),'ĐỦ'))) else (4.0 if X.truthy(X.OR(X.eq(E.v(19,r+30,c+2),'ĐỦ'), X.eq(E.v(19,r+31,c+2),'ĐỦ'), X.eq(E.v(19,r+29,c+2),'ĐỦ'), X.eq(E.v(19,r+40,c+2),'ĐỦ'))) else (2.0 if X.truthy(X.eq(E.v(19,r+35,c+2),'ĐỦ')) else 3.0)))))))

def f19(E, r, c):
    # Ô E16
    # = IF(DOI_CHIEU!$C$33=2,D16,C16)
    return (E.v(19,r,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(19,r,c-2))

def f20(E, r, c):
    # Ô F16
    # = "Dữ kiện: vượng suy theo khuyến nghị = "&DOI_CHIEU!$C$34&"; tài "&PP1_DINH_TINH!$F$98&" chữ (lộ "&PP1_DINH_TINH!$D$98&"), "&IF(COUNTIFS(SAO!$G$5:$G$20
    #   ,2,SAO!$J$5:$J$20,1,SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$G$5:$G$20,2,SAO!$J$5:$J$20,1,SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20
    #   ,0,SAO!$C$5:$C$20,"Bản khí")>0,"có","không có")&" chữ tài (thiên can/bản khí) vượng tại lệnh, "&IF(COUNTIFS(SAO!$G$5:$G$20,2,SAO!$L$5:$L$20,0,SAO!$M$5
    #   :$M$20,0,SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)>0,"còn","không còn")&" chữ tài nguyên vẹn; tỉ kiếp "&PP1_DINH_TINH!$F$96&" chữ (lộ "&PP1_DINH_TINH!$D$96&"
    #   ); quan sát "&PP1_DINH_TINH!$F$99&" chữ; tài/tổng (PP2) "&TEXT(C8,"0.0%")&". Quy tắc – chương ""Bàn về giàu nghèo"" (P1 tr.164-166: ""then chốt là phả
    #   i xem tài tinh đó là hỉ hay kị, có nguồn gốc từ đâu, có kiện toàn đầy đủ không""), thứ tự xét là giả định G12: cách đặc biệt → không xếp; không có tài
    #   , hoặc tài là d"&"ụng mà bị phá → Nghèo; tài là dụng (hoặc hỷ sinh quan) vượng, không bị phá → Giàu lớn (hạ xuống Khá giả nếu tỉ kiếp nhiều áp chế tài
    #    – P2 tr.397); thân vượng tài là dụng không bị phá → Khá giả (""mới có khả năng giàu""); thân vượng tài yếu "&"bị tỉ kiếp đoạt / thân nhược dùng kiếp 
    #   mà kiếp bị quan chế / thân nhược tài nặng thực thương nặng → Vất vả cầu tài; thân nhược dùng tỉ kiếp hoặc dùng ấn mà tài quan không bị phá, thân vượng
    #    tài vượng có nguồn hoặc được bảo vệ, trung hòa có "&"tài → Khá giả; thân nhược mà tài là kỵ và vượng → Vất vả cầu tài; còn lại → Trung bình. Kết luận
    #   : "&E17&". Quy tắc đã khớp: "&TRIM(IF(L42="ĐỦ","GN1 ","")&IF(L43="ĐỦ","GN2 ","")&IF(L45="ĐỦ","GN3 ","")&IF(L46="ĐỦ","GN4 ","")&IF(L47="ĐỦ","GN5 ","")&
    #   IF(L48="ĐỦ","GN6 ","")&IF(L49="ĐỦ","GN7 ","")&IF(L50="ĐỦ","GN8 ","")&IF(L51="ĐỦ","GN9 ","")&IF(L52="ĐỦ","GN10 ","")&IF(L53="ĐỦ","GN11 ","")&IF(L54="ĐỦ
    #   ","GN12 ","")&IF(L55="ĐỦ","GN13 ","")&IF(L56="ĐỦ","GN14 ","")&IF(L57="ĐỦ","GN15 ","")&IF(L44="ĐỦ","GN16 ",""))&" (xem mục F)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: vượng suy theo khuyến nghị = ',E.v(7,34,3)),'; tài '),E.v(9,98,6)),' chữ (lộ '),E.v(9,98,4)),'), '),('có' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,7,20,7), 2.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,7,20,7), 2.0, E.rg(13,5,10,20,10), 1.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0, E.rg(13,5,3,20,3), 'Bản khí')),0.0)) else 'không có')),' chữ tài (thiên can/bản khí) vượng tại lệnh, '),('còn' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,7,20,7), 2.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,25,20,25), 0.0),0.0)) else 'không còn')),' chữ tài nguyên vẹn; tỉ kiếp '),E.v(9,96,6)),' chữ (lộ '),E.v(9,96,4)),'); quan sát '),E.v(9,99,6)),' chữ; tài/tổng (PP2) '),X.TEXT(E.v(19,r-8,c-3), '0.0%')),'. Quy tắc – chương "Bàn về giàu nghèo" (P1 tr.164-166: "then chốt là phải xem tài tinh đó là hỉ hay kị, có nguồn gốc từ đâu, có kiện toàn đầy đủ không"), thứ tự xét là giả định G12: cách đặc biệt → không xếp; không có tài, hoặc tài là d'),'ụng mà bị phá → Nghèo; tài là dụng (hoặc hỷ sinh quan) vượng, không bị phá → Giàu lớn (hạ xuống Khá giả nếu tỉ kiếp nhiều áp chế tài – P2 tr.397); thân vượng tài là dụng không bị phá → Khá giả ("mới có khả năng giàu"); thân vượng tài yếu '),'bị tỉ kiếp đoạt / thân nhược dùng kiếp mà kiếp bị quan chế / thân nhược tài nặng thực thương nặng → Vất vả cầu tài; thân nhược dùng tỉ kiếp hoặc dùng ấn mà tài quan không bị phá, thân vượng tài vượng có nguồn hoặc được bảo vệ, trung hòa có '),'tài → Khá giả; thân nhược mà tài là kỵ và vượng → Vất vả cầu tài; còn lại → Trung bình. Kết luận: '),E.v(19,r+1,c-1)),'. Quy tắc đã khớp: '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('GN1 ' if X.truthy(X.eq(E.v(19,r+26,c+6),'ĐỦ')) else ''),('GN2 ' if X.truthy(X.eq(E.v(19,r+27,c+6),'ĐỦ')) else '')),('GN3 ' if X.truthy(X.eq(E.v(19,r+29,c+6),'ĐỦ')) else '')),('GN4 ' if X.truthy(X.eq(E.v(19,r+30,c+6),'ĐỦ')) else '')),('GN5 ' if X.truthy(X.eq(E.v(19,r+31,c+6),'ĐỦ')) else '')),('GN6 ' if X.truthy(X.eq(E.v(19,r+32,c+6),'ĐỦ')) else '')),('GN7 ' if X.truthy(X.eq(E.v(19,r+33,c+6),'ĐỦ')) else '')),('GN8 ' if X.truthy(X.eq(E.v(19,r+34,c+6),'ĐỦ')) else '')),('GN9 ' if X.truthy(X.eq(E.v(19,r+35,c+6),'ĐỦ')) else '')),('GN10 ' if X.truthy(X.eq(E.v(19,r+36,c+6),'ĐỦ')) else '')),('GN11 ' if X.truthy(X.eq(E.v(19,r+37,c+6),'ĐỦ')) else '')),('GN12 ' if X.truthy(X.eq(E.v(19,r+38,c+6),'ĐỦ')) else '')),('GN13 ' if X.truthy(X.eq(E.v(19,r+39,c+6),'ĐỦ')) else '')),('GN14 ' if X.truthy(X.eq(E.v(19,r+40,c+6),'ĐỦ')) else '')),('GN15 ' if X.truthy(X.eq(E.v(19,r+41,c+6),'ĐỦ')) else '')),('GN16 ' if X.truthy(X.eq(E.v(19,r+28,c+6),'ĐỦ')) else '')))),' (xem mục F).')

def f21(E, r, c):
    # Ô C17, D17
    # = INDEX({"Không xếp bậc (cách đặc biệt)","Nghèo","Vất vả cầu tài – dễ nghèo","Trung bình","Khá giả (giàu vừa)","Giàu lớn"},C16+1)&IF(AND(C16=5,E43="ĐỦ")
    #   ," – phú quý song toàn","")
    return X.cat(X.INDEX(X.Arr([['Không xếp bậc (cách đặc biệt)', 'Nghèo', 'Vất vả cầu tài – dễ nghèo', 'Trung bình', 'Khá giả (giàu vừa)', 'Giàu lớn']]), X.add(E.v(19,r-1,c),1.0)),(' – phú quý song toàn' if X.truthy(X.AND(X.eq(E.v(19,r-1,c),5.0), X.eq(E.v(19,r+26,c+2),'ĐỦ'))) else ''))

def f22(E, r, c):
    # Ô E17
    # = IF(AND(DOI_CHIEU!$C$33=0,C16<>D16),INDEX({"Không xếp bậc (cách đặc biệt)","Nghèo","Vất vả cầu tài – dễ nghèo","Trung bình","Khá giả (giàu vừa)","Giàu 
    #   lớn"},C16+1)&" (PP1) / "&INDEX({"Không xếp bậc (cách đặc biệt)","Nghèo","Vất vả cầu tài – dễ nghèo","Trung bình","Khá giả (giàu vừa)","Giàu lớn"},D16+
    #   1)&" (PP2)",INDEX({"Không xếp bậc (cách đặc biệt)","Nghèo","Vất vả cầu tài – dễ nghèo","Trung bình","Khá giả (giàu vừa)","Giàu lớn"},E16+1)&IF(AND(E16
    #   =5,L43="ĐỦ")," – phú quý song toàn",""))
    return (X.cat(X.cat(X.cat(X.INDEX(X.Arr([['Không xếp bậc (cách đặc biệt)', 'Nghèo', 'Vất vả cầu tài – dễ nghèo', 'Trung bình', 'Khá giả (giàu vừa)', 'Giàu lớn']]), X.add(E.v(19,r-1,c-2),1.0)),' (PP1) / '),X.INDEX(X.Arr([['Không xếp bậc (cách đặc biệt)', 'Nghèo', 'Vất vả cầu tài – dễ nghèo', 'Trung bình', 'Khá giả (giàu vừa)', 'Giàu lớn']]), X.add(E.v(19,r-1,c-1),1.0))),' (PP2)') if X.truthy(X.AND(X.eq(E.v(7,33,3),0.0), X.ne(E.v(19,r-1,c-2),E.v(19,r-1,c-1)))) else X.cat(X.INDEX(X.Arr([['Không xếp bậc (cách đặc biệt)', 'Nghèo', 'Vất vả cầu tài – dễ nghèo', 'Trung bình', 'Khá giả (giàu vừa)', 'Giàu lớn']]), X.add(E.v(19,r-1,c),1.0)),(' – phú quý song toàn' if X.truthy(X.AND(X.eq(E.v(19,r-1,c),5.0), X.eq(E.v(19,r+26,c+7),'ĐỦ'))) else '')))

def f23(E, r, c):
    # Ô C18
    # = TRIM(IF(TAI_LOC!$L$42="ĐỦ","• "&TAI_LOC!$G$42&" (P1 tr.164, 95) ","")&IF(TAI_LOC!$L$43="ĐỦ","• "&TAI_LOC!$G$43&" (P1 tr.165) ","")&IF(TAI_LOC!$L$44="Đ
    #   Ủ","• "&TAI_LOC!$G$44&" (P1 tr.164) ","")&IF(TAI_LOC!$L$45="ĐỦ","• "&TAI_LOC!$G$45&" (P1 tr.165) ","")&IF(TAI_LOC!$L$46="ĐỦ","• "&TAI_LOC!$G$46&" (P1 
    #   tr.166) ","")&IF(TAI_LOC!$L$47="ĐỦ","• "&TAI_LOC!$G$47&" (P1 tr.165-166) ","")&IF(TAI_LOC!$L$48="ĐỦ","• "&TAI_LOC!$G$48&" (P1 tr.165) ","")&IF(TAI_LOC
    #   !$L$49="ĐỦ","• "&TAI_LOC!$G$49&" (P1 tr.166) ","")&IF(TAI_LOC!$L$50="ĐỦ","• "&TAI_LOC!$G$50&" (P1 tr.166) ","")&IF(TAI_LOC!$L$51="ĐỦ","• "&TAI_LOC!$G$
    #   51&" (P1 tr.166, 95) ","")&IF(TAI_LOC!$L$52="ĐỦ","• "&TAI_LOC!$G$52&" (P1 tr.166) ","")&IF(TAI_LOC!$L$53="ĐỦ","• "&TAI_LOC!$G$53&" (P1 tr.166) ","")&I
    #   F(TAI_LOC!$L$54="ĐỦ","• "&TAI_LOC!$G$54&" (P1 tr.166) ","")&IF(TAI_LOC!$L$55="ĐỦ","• "&TAI_LOC!$G$55&" (P2 tr.397; P1 tr.94) ","")&IF(TAI_LOC!$L$56="Đ
    #   Ủ","• "&TAI_LOC!$G$56&" (P1 tr.95, 87) ","")&IF(TAI_LOC!$L$57="ĐỦ","• "&TAI_LOC!$G$57&" (P1 tr.87, 95) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(19,42,7)),' (P1 tr.164, 95) ') if X.truthy(X.eq(E.v(19,42,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(19,43,7)),' (P1 tr.165) ') if X.truthy(X.eq(E.v(19,43,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,44,7)),' (P1 tr.164) ') if X.truthy(X.eq(E.v(19,44,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,45,7)),' (P1 tr.165) ') if X.truthy(X.eq(E.v(19,45,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,46,7)),' (P1 tr.166) ') if X.truthy(X.eq(E.v(19,46,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,47,7)),' (P1 tr.165-166) ') if X.truthy(X.eq(E.v(19,47,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,48,7)),' (P1 tr.165) ') if X.truthy(X.eq(E.v(19,48,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,49,7)),' (P1 tr.166) ') if X.truthy(X.eq(E.v(19,49,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,50,7)),' (P1 tr.166) ') if X.truthy(X.eq(E.v(19,50,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,51,7)),' (P1 tr.166, 95) ') if X.truthy(X.eq(E.v(19,51,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,52,7)),' (P1 tr.166) ') if X.truthy(X.eq(E.v(19,52,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,53,7)),' (P1 tr.166) ') if X.truthy(X.eq(E.v(19,53,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,54,7)),' (P1 tr.166) ') if X.truthy(X.eq(E.v(19,54,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,55,7)),' (P2 tr.397; P1 tr.94) ') if X.truthy(X.eq(E.v(19,55,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,56,7)),' (P1 tr.95, 87) ') if X.truthy(X.eq(E.v(19,56,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,57,7)),' (P1 tr.87, 95) ') if X.truthy(X.eq(E.v(19,57,12),'ĐỦ')) else '')))

def f24(E, r, c):
    # Ô C19
    # = MENH_CUC!D18&IF(OR(AND(E16>=4,MENH_CUC!$U$18="NGHÈO"),AND(E16>=1,E16<=2,MENH_CUC!$U$18="GIÀU"))," → LƯU Ý: bảng tóm tắt P1 tr.173 (so nhóm có lợi với 
    #   nhóm bất lợi) cho chiều ngược với bậc giàu ở trên (xét theo câu chữ P1 tr.95-96). Sách không nói cách nào trội; nên coi mức giàu là ở giữa hai kết luậ
    #   n và ứng theo vận.","")
    return X.cat(E.v(18,r-1,c+1),(' → LƯU Ý: bảng tóm tắt P1 tr.173 (so nhóm có lợi với nhóm bất lợi) cho chiều ngược với bậc giàu ở trên (xét theo câu chữ P1 tr.95-96). Sách không nói cách nào trội; nên coi mức giàu là ở giữa hai kết luận và ứng theo vận.' if X.truthy(X.OR(X.AND(X.ge(E.v(19,r-3,c+2),4.0), X.eq(E.v(18,18,21),'NGHÈO')), X.AND(X.ge(E.v(19,r-3,c+2),1.0), X.le(E.v(19,r-3,c+2),2.0), X.eq(E.v(18,18,21),'GIÀU')))) else ''))

def f25(E, r, c):
    # Ô C21
    # = TRIM(IF(TAI_LOC!$L$58="ĐỦ","• "&TAI_LOC!$G$58&" (P1 tr.96) ","")&IF(TAI_LOC!$L$59="ĐỦ","• "&TAI_LOC!$G$59&" (P1 tr.96) ","")&IF(TAI_LOC!$L$60="ĐỦ","• 
    #   "&TAI_LOC!$G$60&" (P1 tr.96) ","")&IF(TAI_LOC!$L$61="ĐỦ","• "&TAI_LOC!$G$61&" (P1 tr.96) ","")&IF(TAI_LOC!$L$62="ĐỦ","• "&TAI_LOC!$G$62&" (P1 tr.96) "
    #   ,"")&IF(TAI_LOC!$L$63="ĐỦ","• "&TAI_LOC!$G$63&" (P1 tr.96) ","")&IF(TAI_LOC!$L$65="ĐỦ","• "&TAI_LOC!$G$65&" (P1 tr.88) ","")&IF(TAI_LOC!$L$66="ĐỦ","• 
    #   "&TAI_LOC!$G$66&" (P1 tr.88, 96) ","")&IF(TAI_LOC!$L$67="ĐỦ","• "&TAI_LOC!$G$67&" (P2 tr.378; P1 tr.94) ","")&IF(TAI_LOC!$L$68="ĐỦ","• "&TAI_LOC!$G$68
    #   &" (P1 tr.95) ","")&IF(TAI_LOC!$L$69="ĐỦ","• "&TAI_LOC!$G$69&" (P1 tr.94) ","")&IF(TAI_LOC!$L$70="ĐỦ","• "&TAI_LOC!$G$70&" (P1 tr.95) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(19,58,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,58,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(19,59,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,59,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,60,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,60,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,61,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,61,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,62,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,62,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,63,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,63,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,65,7)),' (P1 tr.88) ') if X.truthy(X.eq(E.v(19,65,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,66,7)),' (P1 tr.88, 96) ') if X.truthy(X.eq(E.v(19,66,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,67,7)),' (P2 tr.378; P1 tr.94) ') if X.truthy(X.eq(E.v(19,67,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,68,7)),' (P1 tr.95) ') if X.truthy(X.eq(E.v(19,68,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,69,7)),' (P1 tr.94) ') if X.truthy(X.eq(E.v(19,69,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,70,7)),' (P1 tr.95) ') if X.truthy(X.eq(E.v(19,70,12),'ĐỦ')) else '')))

def f26(E, r, c):
    # Ô C22
    # = TRIM(IF(TAI_LOC!$L$64="ĐỦ","• "&TAI_LOC!$G$64&" (P1 tr.96) ",""))
    return X.TRIM((X.cat(X.cat('• ',E.v(19,64,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,64,12),'ĐỦ')) else ''))

def f27(E, r, c):
    # Ô C25
    # = TRIM(IF(TAI_LOC!$L$71="ĐỦ","• "&TAI_LOC!$G$71&" (P1 tr.95, 123) ","")&IF(TAI_LOC!$L$73="ĐỦ","• "&TAI_LOC!$G$73&" (P1 tr.123) ","")&IF(TAI_LOC!$L$74="Đ
    #   Ủ","• "&TAI_LOC!$G$74&" (P1 tr.95, 123) ",""))
    return X.TRIM(X.cat(X.cat((X.cat(X.cat('• ',E.v(19,71,7)),' (P1 tr.95, 123) ') if X.truthy(X.eq(E.v(19,71,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(19,73,7)),' (P1 tr.123) ') if X.truthy(X.eq(E.v(19,73,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,74,7)),' (P1 tr.95, 123) ') if X.truthy(X.eq(E.v(19,74,12),'ĐỦ')) else '')))

def f28(E, r, c):
    # Ô C26
    # = TRIM(IF(TAI_LOC!$L$72="ĐỦ","• "&TAI_LOC!$G$72&" (P1 tr.95) ","")&IF(TAI_LOC!$L$75="ĐỦ","• "&TAI_LOC!$G$75&" (P1 tr.94) ","")&IF(TAI_LOC!$L$76="ĐỦ","• 
    #   "&TAI_LOC!$G$76&" (P1 tr.95) ","")&IF(TAI_LOC!$L$77="ĐỦ","• "&TAI_LOC!$G$77&" (P1 tr.96) ","")&IF(TAI_LOC!$L$78="ĐỦ","• "&TAI_LOC!$G$78&" (P1 tr.94) "
    #   ,"")&IF(TAI_LOC!$L$79="ĐỦ","• "&TAI_LOC!$G$79&" (P1 tr.123) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(19,72,7)),' (P1 tr.95) ') if X.truthy(X.eq(E.v(19,72,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(19,75,7)),' (P1 tr.94) ') if X.truthy(X.eq(E.v(19,75,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,76,7)),' (P1 tr.95) ') if X.truthy(X.eq(E.v(19,76,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,77,7)),' (P1 tr.96) ') if X.truthy(X.eq(E.v(19,77,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,78,7)),' (P1 tr.94) ') if X.truthy(X.eq(E.v(19,78,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(19,79,7)),' (P1 tr.123) ') if X.truthy(X.eq(E.v(19,79,12),'ĐỦ')) else '')))

def f29(E, r, c):
    # Ô C27
    # = INDEX(CHI_TEN,CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)+1)&" (mộ của hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&") – "&IF(COUNTIF(TRUNG_GI
    #   AN!$C$6:$F$6,CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4))>0,"CÓ trong tứ trụ: gặp vận/năm xung chi này dễ phát tài (P1 tr.95)","không có trong tứ t
    #   rụ")
    return X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0),1.0)),' (mộ của hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),') – '),('CÓ trong tứ trụ: gặp vận/năm xung chi này dễ phát tài (P1 tr.95)' if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)),0.0)) else 'không có trong tứ trụ'))

def f30(E, r, c):
    # Ô C31, D31, E31, F31, G31, H31, I31, J31
    # = DAI_VAN!C35&" ("&DAI_VAN!C36&")"
    return X.cat(X.cat(X.cat(E.v(14,r+4,c),' ('),E.v(14,r+5,c)),')')

def f31(E, r, c):
    # Ô C32, D32, E32, F32, G32, H32, I32, J32
    # = IF(AND(OR(DAI_VAN!C42=MOD(TRUNG_GIAN!$E$7+2,5),DAI_VAN!C43=MOD(TRUNG_GIAN!$E$7+2,5)),OR(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)>
    #   =0,MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)),AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)>=
    #   0,MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),1,0)
    return (1.0 if X.truthy(X.AND(X.OR(X.eq(E.v(14,r+10,c),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(E.v(14,r+11,c),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.OR(X.AND(X.ge((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))), X.AND(X.ge((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 0.0)

def f32(E, r, c):
    # Ô C33, D33, E33, F33, G33, H33, I33, J33
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$72,DUNG_HY_KY!$C$72)=""),OR(DAI_VAN!C42=TRUNG_GIAN!$E$7,DAI_VAN!C43=TRUNG_GIAN!$E$7),$L$51="ĐỦ"),1,0)
    return (1.0 if X.truthy(X.AND(X.eq((E.v(8,72,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,72,3)),''), X.OR(X.eq(E.v(14,r+9,c),E.v(23,7,5)), X.eq(E.v(14,r+10,c),E.v(23,7,5))), X.eq(E.v(19,51,12),'ĐỦ'))) else 0.0)

def f33(E, r, c):
    # Ô C34, D34, E34, F34, G34, H34, I34, J34
    # = IF(AND($L$74="ĐỦ",INDEX(M_XU,DAI_VAN!C34+1,CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)+1)=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(19,74,12),'ĐỦ'), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r,c),1.0), X.add(X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0),1.0)),1.0))) else 0.0)

def f34(E, r, c):
    # Ô C35, D35, E35, F35, G35, H35, I35, J35
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$72,DUNG_HY_KY!$C$72)=""),OR(DAI_VAN!C42=MOD(TRUNG_GIAN!$E$7+2,5),DAI_VAN!C43=MOD(TRUNG_GIAN!$E$7+2,5)),(IF(
    #   DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)
    return (1.0 if X.truthy(X.AND(X.eq((E.v(8,72,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,72,3)),''), X.OR(X.eq(E.v(14,r+7,c),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(E.v(14,r+8,c),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)

def f35(E, r, c):
    # Ô C36, D36, E36, F36, G36, H36, I36, J36
    # = C89
    return E.v(19,r+53,c)

def f36(E, r, c):
    # Ô C37
    # = TRIM(IF(C36="Tốt",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi: "&C90&") ","")&IF(D36="Tốt",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9
    #   )&" tuổi: "&D90&") ","")&IF(E36="Tốt",DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi: "&E90&") ","")&IF(F36="Tốt",DAI_VAN!F35&" ("&DAI_VAN!F3
    #   6&"–"&(DAI_VAN!F36+9)&" tuổi: "&F90&") ","")&IF(G36="Tốt",DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi: "&G90&") ","")&IF(H36="Tốt",DAI_VAN
    #   !H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi: "&H90&") ","")&IF(I36="Tốt",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi: "&I90&") ","")&
    #   IF(J36="Tốt",DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi: "&J90&") ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c),' ('),E.v(14,r-1,c)),'–'),X.add(E.v(14,r-1,c),9.0)),' tuổi: '),E.v(19,r+53,c)),') ') if X.truthy(X.eq(E.v(19,r-1,c),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+1),' ('),E.v(14,r-1,c+1)),'–'),X.add(E.v(14,r-1,c+1),9.0)),' tuổi: '),E.v(19,r+53,c+1)),') ') if X.truthy(X.eq(E.v(19,r-1,c+1),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+2),' ('),E.v(14,r-1,c+2)),'–'),X.add(E.v(14,r-1,c+2),9.0)),' tuổi: '),E.v(19,r+53,c+2)),') ') if X.truthy(X.eq(E.v(19,r-1,c+2),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+3),' ('),E.v(14,r-1,c+3)),'–'),X.add(E.v(14,r-1,c+3),9.0)),' tuổi: '),E.v(19,r+53,c+3)),') ') if X.truthy(X.eq(E.v(19,r-1,c+3),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+4),' ('),E.v(14,r-1,c+4)),'–'),X.add(E.v(14,r-1,c+4),9.0)),' tuổi: '),E.v(19,r+53,c+4)),') ') if X.truthy(X.eq(E.v(19,r-1,c+4),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+5),' ('),E.v(14,r-1,c+5)),'–'),X.add(E.v(14,r-1,c+5),9.0)),' tuổi: '),E.v(19,r+53,c+5)),') ') if X.truthy(X.eq(E.v(19,r-1,c+5),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+6),' ('),E.v(14,r-1,c+6)),'–'),X.add(E.v(14,r-1,c+6),9.0)),' tuổi: '),E.v(19,r+53,c+6)),') ') if X.truthy(X.eq(E.v(19,r-1,c+6),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-2,c+7),' ('),E.v(14,r-1,c+7)),'–'),X.add(E.v(14,r-1,c+7),9.0)),' tuổi: '),E.v(19,r+53,c+7)),') ') if X.truthy(X.eq(E.v(19,r-1,c+7),'Tốt')) else '')))

def f37(E, r, c):
    # Ô C38
    # = TRIM(IF(C36="Xấu",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi: "&C90&") ","")&IF(D36="Xấu",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9
    #   )&" tuổi: "&D90&") ","")&IF(E36="Xấu",DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi: "&E90&") ","")&IF(F36="Xấu",DAI_VAN!F35&" ("&DAI_VAN!F3
    #   6&"–"&(DAI_VAN!F36+9)&" tuổi: "&F90&") ","")&IF(G36="Xấu",DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi: "&G90&") ","")&IF(H36="Xấu",DAI_VAN
    #   !H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi: "&H90&") ","")&IF(I36="Xấu",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi: "&I90&") ","")&
    #   IF(J36="Xấu",DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi: "&J90&") ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c),' ('),E.v(14,r-2,c)),'–'),X.add(E.v(14,r-2,c),9.0)),' tuổi: '),E.v(19,r+52,c)),') ') if X.truthy(X.eq(E.v(19,r-2,c),'Xấu')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+1),' ('),E.v(14,r-2,c+1)),'–'),X.add(E.v(14,r-2,c+1),9.0)),' tuổi: '),E.v(19,r+52,c+1)),') ') if X.truthy(X.eq(E.v(19,r-2,c+1),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+2),' ('),E.v(14,r-2,c+2)),'–'),X.add(E.v(14,r-2,c+2),9.0)),' tuổi: '),E.v(19,r+52,c+2)),') ') if X.truthy(X.eq(E.v(19,r-2,c+2),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+3),' ('),E.v(14,r-2,c+3)),'–'),X.add(E.v(14,r-2,c+3),9.0)),' tuổi: '),E.v(19,r+52,c+3)),') ') if X.truthy(X.eq(E.v(19,r-2,c+3),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+4),' ('),E.v(14,r-2,c+4)),'–'),X.add(E.v(14,r-2,c+4),9.0)),' tuổi: '),E.v(19,r+52,c+4)),') ') if X.truthy(X.eq(E.v(19,r-2,c+4),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+5),' ('),E.v(14,r-2,c+5)),'–'),X.add(E.v(14,r-2,c+5),9.0)),' tuổi: '),E.v(19,r+52,c+5)),') ') if X.truthy(X.eq(E.v(19,r-2,c+5),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+6),' ('),E.v(14,r-2,c+6)),'–'),X.add(E.v(14,r-2,c+6),9.0)),' tuổi: '),E.v(19,r+52,c+6)),') ') if X.truthy(X.eq(E.v(19,r-2,c+6),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-3,c+7),' ('),E.v(14,r-2,c+7)),'–'),X.add(E.v(14,r-2,c+7),9.0)),' tuổi: '),E.v(19,r+52,c+7)),') ') if X.truthy(X.eq(E.v(19,r-2,c+7),'Xấu')) else '')))

def f38(E, r, c):
    # Ô E42
    # = IF(AND(AND(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$
    #   20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),NOT(OR(AND(TRUNG_GIAN!$C
    #   $7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$
    #   D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),"Đ
    #   Ủ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))))) else '—')

def f39(E, r, c):
    # Ô F42
    # = IF(AND(AND(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$
    #   20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),NOT(OR(AND(TRUNG_GIAN!$C
    #   $7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$
    #   D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),"Đ
    #   Ủ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))))) else '—')

def f40(E, r, c):
    # Ô J42, J43, J44, J45, J46, J47, J48, J49, J50, J51, J52, J53 … (40 ô)
    # = IF(AND(E42="ĐỦ",F42="ĐỦ"),"cả hai",IF(E42="ĐỦ","PP1",IF(F42="ĐỦ","PP2","")))
    return ('cả hai' if X.truthy(X.AND(X.eq(E.v(19,r,c-5),'ĐỦ'), X.eq(E.v(19,r,c-4),'ĐỦ'))) else ('PP1' if X.truthy(X.eq(E.v(19,r,c-5),'ĐỦ')) else ('PP2' if X.truthy(X.eq(E.v(19,r,c-4),'ĐỦ')) else '')))

def f41(E, r, c):
    # Ô K42
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ
    #    năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUN
    #   G_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&
    #   " "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN
    #   !$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(C
    #   OUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&C
    #   OUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0
    #   ,"Chính tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.164, 95): Tài là dụng thần, tài vượng (thiên can/bản khí vượng tại lệnh hoặc ≥ P_NHIEU chữ
    #   ), tài không bị phá (không bị hợp mất, không bị tỉ kiếp lộ kề khắc, chi tài không bị xung) → "&IF(J42="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J42&") ⇒ "&G42&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.164, 95): Tài là dụng thần, tài vượng (thiên can/bản khí vượng tại lệnh hoặc ≥ P_NHIEU chữ), tài không bị phá (không bị hợp mất, không bị tỉ kiếp lộ kề khắc, chi tài không bị xung) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f42(E, r, c):
    # Ô L42, L43, L44, L45, L46, L47, L48, L49, L50, L51, L52, L53 … (40 ô)
    # = IF(DOI_CHIEU!$C$33=2,F42,E42)
    return (E.v(19,r,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(19,r,c-7))

def f43(E, r, c):
    # Ô E43
    # = IF(AND(AND(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=3)),AND(DUNG_HY_KY!$C$110>=0,MOD(DUNG_HY_KY!$
    #   C$110-TRUNG_GIAN!$E$7,5)=2)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),NOT(OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=
    #   1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_
    #   GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1
    #   )+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),3.0))), X.AND(X.ge(E.v(8,110,3),0.0), X.eq(X.MOD(X.sub(E.v(8,110,3),E.v(23,7,5)), 5.0),2.0))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))))) else '—')

def f44(E, r, c):
    # Ô F43
    # = IF(AND(AND(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=3)),AND(DUNG_HY_KY!$D$110>=0,MOD(DUNG_HY_KY!$
    #   D$110-TRUNG_GIAN!$E$7,5)=2)),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUN
    #   TIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),NOT(OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=
    #   1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_
    #   GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1
    #   )+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),3.0))), X.AND(X.ge(E.v(8,110,4),0.0), X.eq(X.MOD(X.sub(E.v(8,110,4),E.v(23,7,5)), 5.0),2.0))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))))) else '—')

def f45(E, r, c):
    # Ô K43
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUN
    #   G_HY_KY!$C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$110+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC
    #   !$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,T
    #   RUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D
    #   $7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&IND
    #   EX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TR
    #   UNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SA
    #   O!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SA
    #   O!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Ch
    #   ính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" c
    #   hữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượ
    #   ng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,
    #   "Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.165): T
    #   hân vượng, dụng thần là quan sát, tài là hỷ thần; tài vượng, không bị khắc/hợp/xung → "&IF(J43="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J43&") ⇒ "&G43&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.165): Thân vượng, dụng thần là quan sát, tài là hỷ thần; tài vượng, không bị khắc/hợp/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f46(E, r, c):
    # Ô E44
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=1),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2)),NOT(OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$
    #   7+2,5),OR(PP1_DINH_TINH!$C$38=1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$
    #   7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),NOT(OR(OR(((COUNTIFS(SA
    #   O!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"B
    #   ản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_D
    #   INH_TINH!$F$98>=P_NHIEU)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))), X.NOT(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))))) else '—')

def f47(E, r, c):
    # Ô F44
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=1),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2)),NOT(OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$
    #   7+2,5),OR(PP1_DINH_TINH!$C$38=1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$
    #   7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),NOT(OR(OR(((COUNTIFS(SA
    #   O!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"B
    #   ản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_D
    #   INH_TINH!$F$98>=P_NHIEU)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))), X.NOT(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))))) else '—')

def f48(E, r, c):
    # Ô K44
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GI
    #   AN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"
    #   /"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIA
    #   N!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SA
    #   O!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SA
    #   O!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên 
    #   tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ",
    #   "")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại
    #    lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO
    #   !$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.164): Thân vượ
    #   ng, tài là dụng thần, tài không bị phá (chưa đủ điều kiện tài vượng) → "&IF(J44="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J4
    #   4&") ⇒ "&G44&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.164): Thân vượng, tài là dụng thần, tài không bị phá (chưa đủ điều kiện tài vượng) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f49(E, r, c):
    # Ô E45
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tà
    #   i",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),NOT(OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!
    #   $C$38=1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(
    #   TRUNG_GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$
    #   M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),OR((PP1_DINH_TINH!$F$97>=1),(PP1_DINH_TINH!$F$9
    #   9>=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))), X.OR(X.ge(E.v(9,97,6),1.0), X.ge(E.v(9,99,6),1.0)))) else '—')

def f50(E, r, c):
    # Ô F45
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tà
    #   i",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),NOT(OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!
    #   $C$38=1,TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(
    #   TRUNG_GIAN!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$
    #   M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)))),OR((PP1_DINH_TINH!$F$97>=1),(PP1_DINH_TINH!$F$9
    #   9>=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.NOT(X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))), X.OR(X.ge(E.v(9,97,6),1.0), X.ge(E.v(9,99,6),1.0)))) else '—')

def f51(E, r, c):
    # Ô K45
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_T
    #   INH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&
    #   " ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG
    #   _GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&
    #   "/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SA
    #   O!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",
    #   SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"
    #   Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20
    #   ,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Q
    #   uy tắc (P1 tr.165): Thân vượng, tài vượng, có thực thương (nguồn của tài) hoặc quan sát (bảo vệ tài), tài không bị phá → "&IF(J45="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J45&") ⇒ "&G45&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.165): Thân vượng, tài vượng, có thực thương (nguồn của tài) hoặc quan sát (bảo vệ tài), tài không bị phá → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f52(E, r, c):
    # Ô E46
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.ge(E.v(9,98,6),1.0))) else '—')

def f53(E, r, c):
    # Ô F46
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.ge(E.v(9,98,6),1.0))) else '—')

def f54(E, r, c):
    # Ô K46
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$E$9+1)&")"&" → Quy tắc (P1 tr.166): Thân nhược, dụng thần là tỉ kiếp, có tài → "&IF(J46="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J46&") ⇒ "&G46&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.166): Thân nhược, dụng thần là tỉ kiếp, có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f55(E, r, c):
    # Ô E47
    # = IF(AND(AND(AND(AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20
    #   ,0,SAO!$M$5:$M$20,0))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$
    #   5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0))>0)),NOT(OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU)))),"ĐỦ","
    #   —")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)), X.NOT(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))))) else '—')

def f56(E, r, c):
    # Ô F47
    # = IF(AND(AND(AND(AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20
    #   ,0,SAO!$M$5:$M$20,0))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$
    #   5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0))>0)),NOT(OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU)))),"ĐỦ","
    #   —")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)), X.NOT(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))))) else '—')

def f57(E, r, c):
    # Ô K47
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Chính tài không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chí
    #   nh tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)>0,"Thiên tài không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5
    #   :$L$20,0,SAO!$M$5:$M$20,0)>0,"Chính quan không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$
    #   20,0)>0,"Thất sát không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$
    #   5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài t
    #   hiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&I
    #   F(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: 
    #   "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20
    #   ,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.165-166): Thân nhược, dụng thần là ấn; có tà
    #   i và quan; tài, quan không bị khắc/hợp/xung; tài không vượng → "&IF(J47="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J47&") ⇒ "
    #   &G47&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')),(X.cat(X.cat('Chính quan không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')),(X.cat(X.cat('Thất sát không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')),(X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.165-166): Thân nhược, dụng thần là ấn; có tài và quan; tài, quan không bị khắc/hợp/xung; tài không vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f58(E, r, c):
    # Ô E48, F48
    # = IF((PP1_DINH_TINH!$F$98=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(9,98,6),0.0)) else '—')

def f59(E, r, c):
    # Ô K48
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P1 tr.165): Không có tài (rõ và mờ) → "&IF(
    #   J48="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J48&") ⇒ "&G48&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P1 tr.165): Không có tài (rõ và mờ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f60(E, r, c):
    # Ô E49
    # = IF(AND(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=1,T
    #   RUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIA
    #   N!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)), X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f61(E, r, c):
    # Ô F49
    # = IF(AND(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=1,T
    #   RUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIA
    #   N!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)), X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f62(E, r, c):
    # Ô K49
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")
    #   "&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ 
    #   ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_
    #   GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê 
    #   loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","
    #   "))&" → Quy tắc (P1 tr.166): Tài là dụng thần nhưng tài bị khắc/hợp mất hoặc bị xung → "&IF(J49="","chưa đủ điều kiện ở cả hai phương pháp – không áp 
    #   dụng.","ĐỦ ("&J49&") ⇒ "&G49&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.166): Tài là dụng thần nhưng tài bị khắc/hợp mất hoặc bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f63(E, r, c):
    # Ô E50
    # = IF(AND(AND(DUNG_HY_KY!$C$110>=0,MOD(DUNG_HY_KY!$C$110-TRUNG_GIAN!$E$7,5)=2),OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=1,T
    #   RUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIA
    #   N!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(X.MOD(X.sub(E.v(8,110,3),E.v(23,7,5)), 5.0),2.0)), X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f64(E, r, c):
    # Ô F50
    # = IF(AND(AND(DUNG_HY_KY!$D$110>=0,MOD(DUNG_HY_KY!$D$110-TRUNG_GIAN!$E$7,5)=2),OR(AND(TRUNG_GIAN!$C$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$C$38=1,T
    #   RUNG_GIAN!$D$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIAN!$D$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$D$38=1,TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7)),AND(TRUNG_GIA
    #   N!$F$7=MOD(TRUNG_GIAN!$E$7+2,5),OR(PP1_DINH_TINH!$F$38=1)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(X.MOD(X.sub(E.v(8,110,4),E.v(23,7,5)), 5.0),2.0)), X.OR(X.AND(X.eq(E.v(23,7,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,3),1.0), X.eq(E.v(23,7,4),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,4),1.0), X.eq(E.v(23,7,3),E.v(23,7,5)))), X.AND(X.eq(E.v(23,7,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.OR(X.eq(E.v(9,38,6),1.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)))) else '—')

def f65(E, r, c):
    # Ô K50
    # = "Dữ kiện: "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$
    #   D$110+1),"—")&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&
    #   "; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ng
    #   ày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GI
    #   AN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê lo
    #   ại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Chính tài bị xung: "&COUNTIFS(SAO!$F$
    #   5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên tài bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")
    #   )&" → Quy tắc (P1 tr.166): Tài là hỷ thần nhưng bị khắc/hợp mất hoặc bị xung → "&IF(J50="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J50&") ⇒ "&G50&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.166): Tài là hỷ thần nhưng bị khắc/hợp mất hoặc bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f66(E, r, c):
    # Ô E51
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),OR(AND(DUNG_HY_KY!$C$111>=0,MOD(DUNG_HY_KY!$C$111-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$C$92>=0,DUNG_HY_KY!$C$72="",MOD
    #   (DUNG_HY_KY!$C$5+DUNG_HY_KY!$C$92-TRUNG_GIAN!$E$7,5)=2))),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(X.MOD(X.sub(E.v(8,111,3),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,92,3),0.0), X.eq(E.v(8,72,3),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,3),E.v(8,92,3)),E.v(23,7,5)), 5.0),2.0)))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f67(E, r, c):
    # Ô F51
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),OR(AND(DUNG_HY_KY!$D$111>=0,MOD(DUNG_HY_KY!$D$111-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$D$92>=0,DUNG_HY_KY!$D$72="",MOD
    #   (DUNG_HY_KY!$D$5+DUNG_HY_KY!$D$92-TRUNG_GIAN!$E$7,5)=2))),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(X.MOD(X.sub(E.v(8,111,4),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,92,4),0.0), X.eq(E.v(8,72,4),''), X.eq(X.MOD(X.sub(X.add(E.v(8,5,4),E.v(8,92,4)),E.v(23,7,5)), 5.0),2.0)))), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f68(E, r, c):
    # Ô K51
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Kỵ PP1: "&IF(DUNG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$
    #   C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Nhóm nhiều (kỵ theo cân bằng) PP1: "&DUNG_HY_KY!$C$112
    #   &" / PP2: "&DUNG_HY_KY!$D$112&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" 
    #   chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Cá
    #   c chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J
    #   $5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J
    #   $5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài 
    #   thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&
    #   IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lện
    #   h: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$
    #   5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.166, 95): Thân như
    #   ợc, tài là kỵ thần (hoặc nhóm nhiều), tài vượng → "&IF(J51="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J51&") ⇒ "&G51&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Nhóm nhiều (kỵ theo cân bằng) PP1: '),E.v(8,112,3)),' / PP2: '),E.v(8,112,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.166, 95): Thân nhược, tài là kỵ thần (hoặc nhóm nhiều), tài vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f69(E, r, c):
    # Ô E52
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$97=0)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.le(E.v(9,98,6),1.0)), X.eq(E.v(9,97,6),0.0)), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f70(E, r, c):
    # Ô F52
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$97=0)),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.le(E.v(9,98,6),1.0)), X.eq(E.v(9,97,6),0.0)), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f71(E, r, c):
    # Ô K52
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TI
    #   NH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P1 tr.166): Thân vượng, tài ≤ 1 chữ, không 
    #   có thực thương, tỉ kiếp ≥ P_NHIEU → "&IF(J52="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J52&") ⇒ "&G52&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.166): Thân vượng, tài ≤ 1 chữ, không có thực thương, tỉ kiếp ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f72(E, r, c):
    # Ô E53
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=0),AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=0)),(PP1_DINH_TINH!$D$99>=1)),(PP1_DINH_TINH!$F$9
    #   8>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),0.0))), X.ge(E.v(9,99,4),1.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f73(E, r, c):
    # Ô F53
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=0),AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=0)),(PP1_DINH_TINH!$D$99>=1)),(PP1_DINH_TINH!$F$9
    #   8>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),0.0))), X.ge(E.v(9,99,4),1.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f74(E, r, c):
    # Ô K53
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát lộ rõ: "&PP1_DINH_TINH!$D$99&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.166): Thân nhược, dụng thần là tỉ kiếp, có quan sát lộ ở thiên c
    #   an chế tỉ kiếp → "&IF(J53="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J53&") ⇒ "&G53&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát lộ rõ: '),E.v(9,99,4)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.166): Thân nhược, dụng thần là tỉ kiếp, có quan sát lộ ở thiên can chế tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f75(E, r, c):
    # Ô E54
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f76(E, r, c):
    # Ô F54
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f77(E, r, c):
    # Ô K54
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" ch
    #   ữ"&" → Quy tắc (P1 tr.166): Thân nhược, tài ≥ P_NHIEU và thực thương ≥ P_NHIEU → "&IF(J54="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J54&") ⇒ "&G54&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P1 tr.166): Thân nhược, tài ≥ P_NHIEU và thực thương ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f78(E, r, c):
    # Ô E55
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f79(E, r, c):
    # Ô F55
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),(PP1_DINH_TINH!$F$96>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.ge(E.v(9,96,6),E.v(24,99,2)))) else '—')

def f80(E, r, c):
    # Ô K55
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"
    #   ; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"
    #   ,SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"
    #   ,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thi
    #   ên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ;
    #    ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng 
    #   tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIF
    #   S(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.397; P1 tr.
    #   94): Thân vượng, tài vượng nhưng tỉ kiếp ≥ P_NHIEU (tỉ kiếp áp chế tài) → "&IF(J55="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("
    #   &J55&") ⇒ "&G55&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.397; P1 tr.94): Thân vượng, tài vượng nhưng tỉ kiếp ≥ P_NHIEU (tỉ kiếp áp chế tài) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f81(E, r, c):
    # Ô E56
    # = IF(AND(OR(DUNG_HY_KY!$C$6="Trung hòa"),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(8,6,3),'Trung hòa')), X.ge(E.v(9,98,6),1.0))) else '—')

def f82(E, r, c):
    # Ô F56
    # = IF(AND(OR(DUNG_HY_KY!$D$6="Trung hòa"),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(8,6,4),'Trung hòa')), X.ge(E.v(9,98,6),1.0))) else '—')

def f83(E, r, c):
    # Ô K56
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&" → Quy tắc (P1 tr.95, 87): Nhật chủ trung hòa và có tài → "&IF(J56="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J56&") ⇒ "&G56&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P1 tr.95, 87): Nhật chủ trung hòa và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f84(E, r, c):
    # Ô E57
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$98=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.eq(E.v(9,98,6),1.0))) else '—')

def f85(E, r, c):
    # Ô F57
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$98=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.eq(E.v(9,98,6),1.0))) else '—')

def f86(E, r, c):
    # Ô K57
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&" → Quy tắc (P1 tr.87, 95): Thân vượng, tài chỉ 1 chữ → "&IF(J57="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ 
    #   ("&J57&") ⇒ "&G57&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P1 tr.87, 95): Thân vượng, tài chỉ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f87(E, r, c):
    # Ô E58
    # = IF(AND((DUNG_HY_KY!$C$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0))) else '—')

def f88(E, r, c):
    # Ô F58
    # = IF(AND((DUNG_HY_KY!$D$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0))) else '—')

def f89(E, r, c):
    # Ô K58
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Chính tài ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.96): Thân vượng và có chính tài ở trụ năm → "&IF(J58="","chưa đủ điều kiện 
    #   ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J58&") ⇒ "&G58&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.96): Thân vượng và có chính tài ở trụ năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f90(E, r, c):
    # Ô E59, F59
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '—')

def f91(E, r, c):
    # Ô K59
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,1)>0,"Chính tài ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr
    #   .96): Chính tài ở trụ tháng → "&IF(J59="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J59&") ⇒ "&G59&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''))),' → Quy tắc (P1 tr.96): Chính tài ở trụ tháng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f92(E, r, c):
    # Ô E60, F60
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '—')

def f93(E, r, c):
    # Ô K60
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"<>Can",SAO!$B$5:$B$20,2)>0,"Chính tài tàng can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>C
    #   an",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P1 tr.96): Chính tài ở chi ngày (cung phối ngẫu) → "&IF(J60="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J60&") ⇒ "&G60&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P1 tr.96): Chính tài ở chi ngày (cung phối ngẫu) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f94(E, r, c):
    # Ô E61, F61
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '—')

def f95(E, r, c):
    # Ô K61
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,3)>0,"Chính tài ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P1 tr.9
    #   6): Chính tài ở trụ giờ → "&IF(J61="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J61&") ⇒ "&G61&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else ''))),' → Quy tắc (P1 tr.96): Chính tài ở trụ giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f96(E, r, c):
    # Ô E62, F62
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,3))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 3.0)),0.0)) else '—')

def f97(E, r, c):
    # Ô K62
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$
    #   20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)>0,"Thiên tài không bị khắc/hợp mất không bị xung ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,3)>0,"Thiên tài không bị khắc/hợp mất không bị xung ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tà
    #   i",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P1 tr.96): Thiên tài ở trụ ngày ho
    #   ặc giờ, không bị xung, không bị khắc → "&IF(J62="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J62&") ⇒ "&G62&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên tài không bị khắc/hợp mất không bị xung ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài không bị khắc/hợp mất không bị xung ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P1 tr.96): Thiên tài ở trụ ngày hoặc giờ, không bị xung, không bị khắc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f98(E, r, c):
    # Ô E63, F63
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0))) else '—')

def f99(E, r, c):
    # Ô K63
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can")>0,"Thiên tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thiên tài tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ",""))&" → Quy tắc (P1 tr.96): Thiên tài có cả ở thiên can và trong địa chi → "&IF(J63
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J63&") ⇒ "&G63&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')))),' → Quy tắc (P1 tr.96): Thiên tài có cả ở thiên can và trong địa chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f100(E, r, c):
    # Ô E64, F64
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else '—')

def f101(E, r, c):
    # Ô K64
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$B$5:$B$20,0)>0,"Thiên tài thiên can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",
    #   SAO!$B$5:$B$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.96): Thiên tài ở can năm → "&IF(J64="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J64&") ⇒ "&G64&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.96): Thiên tài ở can năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f102(E, r, c):
    # Ô E65, F65
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f103(E, r, c):
    # Ô K65
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính tà
    #   i: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.88): Có chính tài → "&IF(J65="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J65&") ⇒ "&G65&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P1 tr.88): Có chính tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f104(E, r, c):
    # Ô E66, F66
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f105(E, r, c):
    # Ô K66
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên tà
    #   i: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.88, 96): Có thiên tài → "&IF(J66="","chưa
    #    đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J66&") ⇒ "&G66&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P1 tr.88, 96): Có thiên tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f106(E, r, c):
    # Ô E67, F67
    # = IF(AND((PP1_DINH_TINH!$F$97>=1),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,97,6),1.0), X.ge(E.v(9,98,6),1.0))) else '—')

def f107(E, r, c):
    # Ô K67
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_
    #   CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P2 tr.378; P1 tr.94): Có thực thương và có tài (thực thương sinh tài) → "&IF(J67="","chưa đủ điều 
    #   kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J67&") ⇒ "&G67&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P2 tr.378; P1 tr.94): Có thực thương và có tài (thực thương sinh tài) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f108(E, r, c):
    # Ô E68, F68
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"C
    #   hính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),0.0))) else '—')

def f109(E, r, c):
    # Ô K68
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$B$5:$B$20,1)>0,"Thực thần thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can
    #   ",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)>0,
    #   "Chính quan tàng can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)&" chữ;
    #    ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)>0,"Thất sát tàng can ở tháng: 
    #   "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.95): 
    #   Can tháng là thực thần, chi tháng có quan → "&IF(J68="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J68&") ⇒ "&G68&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thực thần thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan tàng can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát tàng can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.95): Can tháng là thực thần, chi tháng có quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f110(E, r, c):
    # Ô E69, F69
    # = IF(OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0))>0),((COUNTIFS(SAO!$F$5:$F$20
    #   ,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>
    #   Can",SAO!$B$5:$B$20,1))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,2))>0),((C
    #   OUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực 
    #   thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3))>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),0.0)), X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)), X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)), X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),0.0)))) else '—')

def f111(E, r, c):
    # Ô K69
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$B$5:$B$20,0)>0,"Thực thần thiên can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",
    #   SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)>0,"Kiế
    #   p tài tàng can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(
    #   COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Thực thần thiên can ở tháng: "&COUNTIF
    #   S(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tà
    #   i",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)>0,"Kiếp tài tàng can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,2)>0,"Thực thần thiên can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Can",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$
    #   B$5:$B$20,2)>0,"Kiếp tài tàng can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20
    #   ,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)>0,"Thực thần thiên ca
    #   n ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$
    #   F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)>0,"Kiếp tài tàng can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,
    #   "Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P1 tr.94): Cùng một trụ: can là thực 
    #   thần, chi có kiếp tài → "&IF(J69="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J69&") ⇒ "&G69&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Thực thần thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài tàng can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Thực thần thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài tàng can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thực thần thiên can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thực thần thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài tàng can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P1 tr.94): Cùng một trụ: can là thực thần, chi có kiếp tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f112(E, r, c):
    # Ô E70, F70
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '—')

def f113(E, r, c):
    # Ô K70
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,0)>0,"Thực thần ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.9
    #   5): Thực thần đóng ở trụ năm → "&IF(J70="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J70&") ⇒ "&G70&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.95): Thực thần đóng ở trụ năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f114(E, r, c):
    # Ô E71, F71
    # = IF(AND((PP1_DINH_TINH!$F$98>=1),(PP1_DINH_TINH!$D$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),1.0), X.eq(E.v(9,98,4),0.0))) else '—')

def f115(E, r, c):
    # Ô K71
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&" → Quy tắ
    #   c (P1 tr.95, 123): Tài chỉ tàng trong địa chi, không lộ → "&IF(J71="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J71&") ⇒ "&G71&
    #   ".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Tài lộ rõ: '),E.v(9,98,4)),' → Quy tắc (P1 tr.95, 123): Tài chỉ tàng trong địa chi, không lộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f116(E, r, c):
    # Ô E72, F72
    # = IF(AND((PP1_DINH_TINH!$D$98>=1),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$5:
    #   $F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,4),1.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0)))) else '—')

def f117(E, r, c):
    # Ô K72
    # = "Dữ kiện: "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Chính tài tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thiên tài 
    #   tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ",""))&" → Quy tắc (P1 tr.95): Tài ch
    #   ỉ lộ ở thiên can, không tàng → "&IF(J72="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J72&") ⇒ "&G72&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')))),' → Quy tắc (P1 tr.95): Tài chỉ lộ ở thiên can, không tàng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f118(E, r, c):
    # Ô E73, F73
    # = IF(AND((PP1_DINH_TINH!$D$98>=1),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,4),1.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0))) else '—')

def f119(E, r, c):
    # Ô K73
    # = "Dữ kiện: "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Chính tài tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thiên tài 
    #   tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ",""))&" → Quy tắc (P1 tr.123): Tài v
    #   ừa lộ vừa tàng → "&IF(J73="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J73&") ⇒ "&G73&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')))),' → Quy tắc (P1 tr.123): Tài vừa lộ vừa tàng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f120(E, r, c):
    # Ô E74, F74
    # = IF(AND((COUNTIF(TRUNG_GIAN!$C$6:$F$6,CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4))>0),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)),0.0), X.ge(E.v(9,98,6),1.0))) else '—')

def f121(E, r, c):
    # Ô K74
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&"
    #    ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.95, 123): Có chi kho (mộ) của hành tài trong tứ trụ → "&IF(J74="
    #   ","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J74&") ⇒ "&G74&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.95, 123): Có chi kho (mộ) của hành tài trong tứ trụ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f122(E, r, c):
    # Ô E75, F75
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '—')

def f123(E, r, c):
    # Ô K75
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$2
    #   0,1)>0,"Kiếp tài ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.94
    #   ): Kiếp tài ở trụ tháng → "&IF(J75="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J75&") ⇒ "&G75&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''))),' → Quy tắc (P1 tr.94): Kiếp tài ở trụ tháng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f124(E, r, c):
    # Ô E76, F76
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f125(E, r, c):
    # Ô K76
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính tà
    #   i: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P1 tr.95): Chính tài và ki
    #   ếp tài cùng xuất hiện → "&IF(J76="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J76&") ⇒ "&G76&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P1 tr.95): Chính tài và kiếp tài cùng xuất hiện → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f126(E, r, c):
    # Ô E77, F77
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0))) else '—')

def f127(E, r, c):
    # Ô K77
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,1)>0,"Thiên tài ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$
    #   5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Tỉ kiên ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Kiếp tài ở giờ: "&
    #   COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P1 tr.96): Trụ tháng có thiên tài, t
    #   rụ giờ có tỉ kiếp → "&IF(J77="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J77&") ⇒ "&G77&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thiên tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Tỉ kiên ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P1 tr.96): Trụ tháng có thiên tài, trụ giờ có tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f128(E, r, c):
    # Ô E78, F78
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f129(E, r, c):
    # Ô K78
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:
    #   $Y$20,0)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P1 tr.94): Kiếp tài nhiều (≥ P_NHIEU chữ) → "&
    #   IF(J78="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J78&") ⇒ "&G78&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P1 tr.94): Kiếp tài nhiều (≥ P_NHIEU chữ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f130(E, r, c):
    # Ô E79
    # = IF(AND(AND((PP1_DINH_TINH!$D$98>=1),(PP1_DINH_TINH!$D$96>=1)),OR(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$
    #   C$110>=0,MOD(DUNG_HY_KY!$C$110-TRUNG_GIAN!$E$7,5)=2))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,98,4),1.0), X.ge(E.v(9,96,4),1.0)), X.OR(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,110,3),0.0), X.eq(X.MOD(X.sub(E.v(8,110,3),E.v(23,7,5)), 5.0),2.0))))) else '—')

def f131(E, r, c):
    # Ô F79
    # = IF(AND(AND((PP1_DINH_TINH!$D$98>=1),(PP1_DINH_TINH!$D$96>=1)),OR(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),AND(DUNG_HY_KY!$
    #   D$110>=0,MOD(DUNG_HY_KY!$D$110-TRUNG_GIAN!$E$7,5)=2))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,98,4),1.0), X.ge(E.v(9,96,4),1.0)), X.OR(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)), X.AND(X.ge(E.v(8,110,4),0.0), X.eq(X.MOD(X.sub(E.v(8,110,4),E.v(23,7,5)), 5.0),2.0))))) else '—')

def f132(E, r, c):
    # Ô K79
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUN
    #   G_HY_KY!$D$110+1),"—")&"; "&"Số chữ Tỉ kiếp lộ rõ: "&PP1_DINH_TINH!$D$96&"; "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"Trụ ngày: "&TRUNG_GIAN!$E
    #   $10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P1 tr.123): Tài lộ, tỉ kiếp cũn
    #   g lộ, và tài là dụng/hỷ thần → "&IF(J79="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J79&") ⇒ "&G79&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Số chữ Tỉ kiếp lộ rõ: '),E.v(9,96,4)),'; '),'Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P1 tr.123): Tài lộ, tỉ kiếp cũng lộ, và tài là dụng/hỷ thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f133(E, r, c):
    # Ô E84, F84
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)))) else '—')

def f134(E, r, c):
    # Ô K84
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521; UH 
    #   tr.239): Có thương quan và tứ trụ không có chính quan (thương tận – UH tr.239) → "&IF(J84="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J84&") ⇒ "&G84&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.521; UH tr.239): Có thương quan và tứ trụ không có chính quan (thương tận – UH tr.239) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f135(E, r, c):
    # Ô E85, F85
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")+C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0)))) else '—')

def f136(E, r, c):
    # Ô K85
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"<>Can")>0,"Chính tài tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COU
    #   NTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thiên tài tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Can")>0,"Chính tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&I
    #   F(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thiên tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên
    #    tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.522; UH tr.173): Tài chỉ có trong tàng can, không lộ ở 
    #   thiên can (""địa chi phục tài, gọi là ám sinh"") → "&IF(J85="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J85&") ⇒ "&G85&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')),(X.cat(X.cat('Chính tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.522; UH tr.173): Tài chỉ có trong tàng can, không lộ ở thiên can ("địa chi phục tài, gọi là ám sinh") → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(19,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(19,r,c-1)),') ⇒ '),E.v(19,r,c-4)),'.')))

def f137(E, r, c):
    # Ô C88, D88, E88, F88, G88, H88, I88, J88
    # = IF(AND(NOT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!$C$110,DAI_VAN!$C$109)=1),MOD(DAI_VAN!C$43-TRUNG_GIAN!$E$7,5)=2),1,IF(AND(IF(DOI_CHIEU!$C$33=2,DAI_VAN!$C$110,
    #   DAI_VAN!$C$109)=1,MOD(DAI_VAN!C$42-TRUNG_GIAN!$E$7,5)=2,OR(MOD(DAI_VAN!C$43-TRUNG_GIAN!$E$7,5)=0,MOD(DAI_VAN!C$43-TRUNG_GIAN!$E$7,5)=4)),2,IF(AND(NOT(
    #   IF(DOI_CHIEU!$C$33=2,DAI_VAN!$C$110,DAI_VAN!$C$109)=1),MOD(DAI_VAN!C$42-TRUNG_GIAN!$E$7,5)=0),3,IF(AND(IF(DOI_CHIEU!$C$33=2,DAI_VAN!$C$110,DAI_VAN!$C$
    #   109)=1,MOD(DAI_VAN!C$42-TRUNG_GIAN!$E$7,5)=2),4,0))))
    return (1.0 if X.truthy(X.AND(X.NOT(X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0)), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),2.0))) else (2.0 if X.truthy(X.AND(X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),2.0), X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)))) else (3.0 if X.truthy(X.AND(X.NOT(X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0)), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0))) else (4.0 if X.truthy(X.AND(X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),2.0))) else 0.0))))

def f138(E, r, c):
    # Ô C89, D89, E89, F89, G89, H89, I89, J89
    # = IF(OR(C88=1,C88=2),"Tốt",IF(OR(C88=3,C88=4),"Xấu",IF(DAI_VAN!C$119>0,"Tốt","Xấu")))
    return ('Tốt' if X.truthy(X.OR(X.eq(E.v(19,r-1,c),1.0), X.eq(E.v(19,r-1,c),2.0))) else ('Xấu' if X.truthy(X.OR(X.eq(E.v(19,r-1,c),3.0), X.eq(E.v(19,r-1,c),4.0))) else ('Tốt' if X.truthy(X.gt(E.v(14,119,c),0.0)) else 'Xấu')))

def f139(E, r, c):
    # Ô C90, D90, E90, F90, G90, H90, I90, J90
    # = CHOOSE(C88+1,"theo xu thế chung của vận","thân vượng, chi vận là tài nên 5 năm sau có tài","vận tài nhưng chi vận giúp thân gánh được tài nên có thể t
    #   ụ tài","thân vượng gặp vận tỉ kiếp nên tranh tài, hao tài","thân nhược gặp vận tài nên tiền đến nhưng khó giữ")
    return X.CHOOSE(X.add(E.v(19,r-2,c),1.0), lambda: 'theo xu thế chung của vận', lambda: 'thân vượng, chi vận là tài nên 5 năm sau có tài', lambda: 'vận tài nhưng chi vận giúp thân gánh được tài nên có thể tụ tài', lambda: 'thân vượng gặp vận tỉ kiếp nên tranh tài, hao tài', lambda: 'thân nhược gặp vận tài nên tiền đến nhưng khó giữ')

def f140(E, r, c):
    # Ô C91, D91, E91, F91, G91, H91, I91, J91
    # = CHOOSE(C88+1,"P1 tr.124–125 (xu thế chung của vận)","P2 tr.522: trong trụ có tài, hành đến đất tài tất đại phát","P2 tr.416 (VD23): tài vận phải gặp v
    #   ận thân vượng mới tụ tài; P2 tr.521: thân nhược tài nhiều, hành vận tỉ kiếp mới phát tài","P2 tr.527: thân vượng hành đến vận tỉ kiếp thì phá tài bại 
    #   nghiệp","P2 tr.527: thân nhược không thắng tài thì vì của mà sinh nạn")
    return X.CHOOSE(X.add(E.v(19,r-3,c),1.0), lambda: 'P1 tr.124–125 (xu thế chung của vận)', lambda: 'P2 tr.522: trong trụ có tài, hành đến đất tài tất đại phát', lambda: 'P2 tr.416 (VD23): tài vận phải gặp vận thân vượng mới tụ tài; P2 tr.521: thân nhược tài nhiều, hành vận tỉ kiếp mới phát tài', lambda: 'P2 tr.527: thân vượng hành đến vận tỉ kiếp thì phá tài bại nghiệp', lambda: 'P2 tr.527: thân nhược không thắng tài thì vì của mà sinh nạn')

CELLS = {
    'C6': (f0, 6, 3),
    'D6': (f1, 6, 4),
    'E6': (f2, 6, 5),
    'C7': (f3, 7, 3),
    'D7': (f4, 7, 4),
    'E7': (f5, 7, 5),
    'C8': (f6, 8, 3),
    'D8': (f6, 8, 4),
    'E8': (f7, 8, 5),
    'C9': (f8, 9, 3),
    'D9': (f9, 9, 4),
    'E9': (f10, 9, 5),
    'C10': (f11, 10, 3),
    'D10': (f11, 10, 4),
    'E10': (f12, 10, 5),
    'C11': (f13, 11, 3),
    'D11': (f14, 11, 4),
    'E11': (f15, 11, 5),
    'B12': (f16, 12, 2),
    'C16': (f17, 16, 3),
    'D16': (f18, 16, 4),
    'E16': (f19, 16, 5),
    'F16': (f20, 16, 6),
    'C17': (f21, 17, 3),
    'D17': (f21, 17, 4),
    'E17': (f22, 17, 5),
    'C18': (f23, 18, 3),
    'C19': (f24, 19, 3),
    'C21': (f25, 21, 3),
    'C22': (f26, 22, 3),
    'C25': (f27, 25, 3),
    'C26': (f28, 26, 3),
    'C27': (f29, 27, 3),
    'C31': (f30, 31, 3),
    'D31': (f30, 31, 4),
    'E31': (f30, 31, 5),
    'F31': (f30, 31, 6),
    'G31': (f30, 31, 7),
    'H31': (f30, 31, 8),
    'I31': (f30, 31, 9),
    'J31': (f30, 31, 10),
    'C32': (f31, 32, 3),
    'D32': (f31, 32, 4),
    'E32': (f31, 32, 5),
    'F32': (f31, 32, 6),
    'G32': (f31, 32, 7),
    'H32': (f31, 32, 8),
    'I32': (f31, 32, 9),
    'J32': (f31, 32, 10),
    'C33': (f32, 33, 3),
    'D33': (f32, 33, 4),
    'E33': (f32, 33, 5),
    'F33': (f32, 33, 6),
    'G33': (f32, 33, 7),
    'H33': (f32, 33, 8),
    'I33': (f32, 33, 9),
    'J33': (f32, 33, 10),
    'C34': (f33, 34, 3),
    'D34': (f33, 34, 4),
    'E34': (f33, 34, 5),
    'F34': (f33, 34, 6),
    'G34': (f33, 34, 7),
    'H34': (f33, 34, 8),
    'I34': (f33, 34, 9),
    'J34': (f33, 34, 10),
    'C35': (f34, 35, 3),
    'D35': (f34, 35, 4),
    'E35': (f34, 35, 5),
    'F35': (f34, 35, 6),
    'G35': (f34, 35, 7),
    'H35': (f34, 35, 8),
    'I35': (f34, 35, 9),
    'J35': (f34, 35, 10),
    'C36': (f35, 36, 3),
    'D36': (f35, 36, 4),
    'E36': (f35, 36, 5),
    'F36': (f35, 36, 6),
    'G36': (f35, 36, 7),
    'H36': (f35, 36, 8),
    'I36': (f35, 36, 9),
    'J36': (f35, 36, 10),
    'C37': (f36, 37, 3),
    'C38': (f37, 38, 3),
    'E42': (f38, 42, 5),
    'F42': (f39, 42, 6),
    'J42': (f40, 42, 10),
    'K42': (f41, 42, 11),
    'L42': (f42, 42, 12),
    'E43': (f43, 43, 5),
    'F43': (f44, 43, 6),
    'J43': (f40, 43, 10),
    'K43': (f45, 43, 11),
    'L43': (f42, 43, 12),
    'E44': (f46, 44, 5),
    'F44': (f47, 44, 6),
    'J44': (f40, 44, 10),
    'K44': (f48, 44, 11),
    'L44': (f42, 44, 12),
    'E45': (f49, 45, 5),
    'F45': (f50, 45, 6),
    'J45': (f40, 45, 10),
    'K45': (f51, 45, 11),
    'L45': (f42, 45, 12),
    'E46': (f52, 46, 5),
    'F46': (f53, 46, 6),
    'J46': (f40, 46, 10),
    'K46': (f54, 46, 11),
    'L46': (f42, 46, 12),
    'E47': (f55, 47, 5),
    'F47': (f56, 47, 6),
    'J47': (f40, 47, 10),
    'K47': (f57, 47, 11),
    'L47': (f42, 47, 12),
    'E48': (f58, 48, 5),
    'F48': (f58, 48, 6),
    'J48': (f40, 48, 10),
    'K48': (f59, 48, 11),
    'L48': (f42, 48, 12),
    'E49': (f60, 49, 5),
    'F49': (f61, 49, 6),
    'J49': (f40, 49, 10),
    'K49': (f62, 49, 11),
    'L49': (f42, 49, 12),
    'E50': (f63, 50, 5),
    'F50': (f64, 50, 6),
    'J50': (f40, 50, 10),
    'K50': (f65, 50, 11),
    'L50': (f42, 50, 12),
    'E51': (f66, 51, 5),
    'F51': (f67, 51, 6),
    'J51': (f40, 51, 10),
    'K51': (f68, 51, 11),
    'L51': (f42, 51, 12),
    'E52': (f69, 52, 5),
    'F52': (f70, 52, 6),
    'J52': (f40, 52, 10),
    'K52': (f71, 52, 11),
    'L52': (f42, 52, 12),
    'E53': (f72, 53, 5),
    'F53': (f73, 53, 6),
    'J53': (f40, 53, 10),
    'K53': (f74, 53, 11),
    'L53': (f42, 53, 12),
    'E54': (f75, 54, 5),
    'F54': (f76, 54, 6),
    'J54': (f40, 54, 10),
    'K54': (f77, 54, 11),
    'L54': (f42, 54, 12),
    'E55': (f78, 55, 5),
    'F55': (f79, 55, 6),
    'J55': (f40, 55, 10),
    'K55': (f80, 55, 11),
    'L55': (f42, 55, 12),
    'E56': (f81, 56, 5),
    'F56': (f82, 56, 6),
    'J56': (f40, 56, 10),
    'K56': (f83, 56, 11),
    'L56': (f42, 56, 12),
    'E57': (f84, 57, 5),
    'F57': (f85, 57, 6),
    'J57': (f40, 57, 10),
    'K57': (f86, 57, 11),
    'L57': (f42, 57, 12),
    'E58': (f87, 58, 5),
    'F58': (f88, 58, 6),
    'J58': (f40, 58, 10),
    'K58': (f89, 58, 11),
    'L58': (f42, 58, 12),
    'E59': (f90, 59, 5),
    'F59': (f90, 59, 6),
    'J59': (f40, 59, 10),
    'K59': (f91, 59, 11),
    'L59': (f42, 59, 12),
    'E60': (f92, 60, 5),
    'F60': (f92, 60, 6),
    'J60': (f40, 60, 10),
    'K60': (f93, 60, 11),
    'L60': (f42, 60, 12),
    'E61': (f94, 61, 5),
    'F61': (f94, 61, 6),
    'J61': (f40, 61, 10),
    'K61': (f95, 61, 11),
    'L61': (f42, 61, 12),
    'E62': (f96, 62, 5),
    'F62': (f96, 62, 6),
    'J62': (f40, 62, 10),
    'K62': (f97, 62, 11),
    'L62': (f42, 62, 12),
    'E63': (f98, 63, 5),
    'F63': (f98, 63, 6),
    'J63': (f40, 63, 10),
    'K63': (f99, 63, 11),
    'L63': (f42, 63, 12),
    'E64': (f100, 64, 5),
    'F64': (f100, 64, 6),
    'J64': (f40, 64, 10),
    'K64': (f101, 64, 11),
    'L64': (f42, 64, 12),
    'E65': (f102, 65, 5),
    'F65': (f102, 65, 6),
    'J65': (f40, 65, 10),
    'K65': (f103, 65, 11),
    'L65': (f42, 65, 12),
    'E66': (f104, 66, 5),
    'F66': (f104, 66, 6),
    'J66': (f40, 66, 10),
    'K66': (f105, 66, 11),
    'L66': (f42, 66, 12),
    'E67': (f106, 67, 5),
    'F67': (f106, 67, 6),
    'J67': (f40, 67, 10),
    'K67': (f107, 67, 11),
    'L67': (f42, 67, 12),
    'E68': (f108, 68, 5),
    'F68': (f108, 68, 6),
    'J68': (f40, 68, 10),
    'K68': (f109, 68, 11),
    'L68': (f42, 68, 12),
    'E69': (f110, 69, 5),
    'F69': (f110, 69, 6),
    'J69': (f40, 69, 10),
    'K69': (f111, 69, 11),
    'L69': (f42, 69, 12),
    'E70': (f112, 70, 5),
    'F70': (f112, 70, 6),
    'J70': (f40, 70, 10),
    'K70': (f113, 70, 11),
    'L70': (f42, 70, 12),
    'E71': (f114, 71, 5),
    'F71': (f114, 71, 6),
    'J71': (f40, 71, 10),
    'K71': (f115, 71, 11),
    'L71': (f42, 71, 12),
    'E72': (f116, 72, 5),
    'F72': (f116, 72, 6),
    'J72': (f40, 72, 10),
    'K72': (f117, 72, 11),
    'L72': (f42, 72, 12),
    'E73': (f118, 73, 5),
    'F73': (f118, 73, 6),
    'J73': (f40, 73, 10),
    'K73': (f119, 73, 11),
    'L73': (f42, 73, 12),
    'E74': (f120, 74, 5),
    'F74': (f120, 74, 6),
    'J74': (f40, 74, 10),
    'K74': (f121, 74, 11),
    'L74': (f42, 74, 12),
    'E75': (f122, 75, 5),
    'F75': (f122, 75, 6),
    'J75': (f40, 75, 10),
    'K75': (f123, 75, 11),
    'L75': (f42, 75, 12),
    'E76': (f124, 76, 5),
    'F76': (f124, 76, 6),
    'J76': (f40, 76, 10),
    'K76': (f125, 76, 11),
    'L76': (f42, 76, 12),
    'E77': (f126, 77, 5),
    'F77': (f126, 77, 6),
    'J77': (f40, 77, 10),
    'K77': (f127, 77, 11),
    'L77': (f42, 77, 12),
    'E78': (f128, 78, 5),
    'F78': (f128, 78, 6),
    'J78': (f40, 78, 10),
    'K78': (f129, 78, 11),
    'L78': (f42, 78, 12),
    'E79': (f130, 79, 5),
    'F79': (f131, 79, 6),
    'J79': (f40, 79, 10),
    'K79': (f132, 79, 11),
    'L79': (f42, 79, 12),
    'E84': (f133, 84, 5),
    'F84': (f133, 84, 6),
    'J84': (f40, 84, 10),
    'K84': (f134, 84, 11),
    'L84': (f42, 84, 12),
    'E85': (f135, 85, 5),
    'F85': (f135, 85, 6),
    'J85': (f40, 85, 10),
    'K85': (f136, 85, 11),
    'L85': (f42, 85, 12),
    'C88': (f137, 88, 3),
    'D88': (f137, 88, 4),
    'E88': (f137, 88, 5),
    'F88': (f137, 88, 6),
    'G88': (f137, 88, 7),
    'H88': (f137, 88, 8),
    'I88': (f137, 88, 9),
    'J88': (f137, 88, 10),
    'C89': (f138, 89, 3),
    'D89': (f138, 89, 4),
    'E89': (f138, 89, 5),
    'F89': (f138, 89, 6),
    'G89': (f138, 89, 7),
    'H89': (f138, 89, 8),
    'I89': (f138, 89, 9),
    'J89': (f138, 89, 10),
    'C90': (f139, 90, 3),
    'D90': (f139, 90, 4),
    'E90': (f139, 90, 5),
    'F90': (f139, 90, 6),
    'G90': (f139, 90, 7),
    'H90': (f139, 90, 8),
    'I90': (f139, 90, 9),
    'J90': (f139, 90, 10),
    'C91': (f140, 91, 3),
    'D91': (f140, 91, 4),
    'E91': (f140, 91, 5),
    'F91': (f140, 91, 6),
    'G91': (f140, 91, 7),
    'H91': (f140, 91, 8),
    'I91': (f140, 91, 9),
    'J91': (f140, 91, 10),
}
