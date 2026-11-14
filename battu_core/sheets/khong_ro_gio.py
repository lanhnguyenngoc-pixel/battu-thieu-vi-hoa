# -*- coding: utf-8 -*-
"""Module sheet KHONG_RO_GIO – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.3.xlsx (SHA-256 7ae759658d290172bbbf7b6642e24532239e3b9ea54ca71e2696802411bca4c5).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'KHONG_RO_GIO'
SI = 6

def f0(E, r, c):
    # Ô C5
    # = NHAP_LIEU!$C$9
    return E.v(2,9,3)

def f1(E, r, c):
    # Ô C7
    # = IF(OR(DAI_VAN!$C$8="",DAI_VAN!$C$8="LỖI",NOT((NHAP_LIEU!$C$9="Không rõ giờ"))),"",IFERROR(IF(MATCH((DAI_VAN!$C$8-0.5-DAI_VAN!$C$9/24),TIET_UTC,1)<>MAT
    #   CH((DAI_VAN!$C$8-0.5-DAI_VAN!$C$9/24)+1-1/1440,TIET_UTC,1),"CẢNH BÁO: tiết "&INDEX(TIET_TEN,MATCH((DAI_VAN!$C$8-0.5-DAI_VAN!$C$9/24)+1-1/1440,TIET_UTC
    #   ,1))&" giao lúc "&TEXT(INDEX(TIET_UTC,MATCH((DAI_VAN!$C$8-0.5-DAI_VAN!$C$9/24)+1-1/1440,TIET_UTC,1))+DAI_VAN!$C$9/24,"dd/mm/yyyy hh:mm")&" – trụ tháng
    #    (có thể cả trụ năm) phụ thuộc giờ sinh; bảng 12 giờ dưới đây đang lấy trụ tháng lúc 12:00.","Không – trụ năm, tháng, ngày không phụ thuộc giờ."),""))
    return ('' if X.truthy(X.OR(X.eq(E.v(14,8,3),''), X.eq(E.v(14,8,3),'LỖI'), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')))) else X.IFERROR(lambda: (X.cat(X.cat(X.cat(X.cat('CẢNH BÁO: tiết ',X.INDEX(E.rg(25,4,3,2439,3), X.MATCH(X.sub(X.add(X.sub(X.sub(E.v(14,8,3),0.5),X.div(E.v(14,9,3),24.0)),1.0),X.div(1.0,1440.0)), E.rg(25,4,1,2439,1), 1.0))),' giao lúc '),X.TEXT(X.add(X.INDEX(E.rg(25,4,1,2439,1), X.MATCH(X.sub(X.add(X.sub(X.sub(E.v(14,8,3),0.5),X.div(E.v(14,9,3),24.0)),1.0),X.div(1.0,1440.0)), E.rg(25,4,1,2439,1), 1.0)),X.div(E.v(14,9,3),24.0)), 'dd/mm/yyyy hh:mm')),' – trụ tháng (có thể cả trụ năm) phụ thuộc giờ sinh; bảng 12 giờ dưới đây đang lấy trụ tháng lúc 12:00.') if X.truthy(X.ne(X.MATCH(X.sub(X.sub(E.v(14,8,3),0.5),X.div(E.v(14,9,3),24.0)), E.rg(25,4,1,2439,1), 1.0),X.MATCH(X.sub(X.add(X.sub(X.sub(E.v(14,8,3),0.5),X.div(E.v(14,9,3),24.0)),1.0),X.div(1.0,1440.0)), E.rg(25,4,1,2439,1), 1.0))) else 'Không – trụ năm, tháng, ngày không phụ thuộc giờ.'), lambda: ''))

def f2(E, r, c):
    # Ô F11, F12, F13, F14, F15, F16, F17, F18, F19, F20, F21, F22
    # = IF(OR($C11="",$D11=""),"",IFERROR(IF(CHOOSE(MATCH($D11,{"Kết hôn","Ly hôn – vợ chồng bất hòa","Mất cha – cha bệnh nạn","Mất mẹ – mẹ bệnh nạn","Tai nạn
    #    – bệnh nặng","Thăng tiến – phát tài","Thất bại – phá sản – kiện tụng","Sinh con"},0),OR(INDEX(LUU_NIEN!$BE$30:$BE$129,MATCH($C11,LUU_NIEN!$A$30:$A$12
    #   9,0))=1,INDEX(LUU_NIEN!$BF$30:$BF$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1,INDEX(LUU_NIEN!$BA$30:$BA$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1),IND
    #   EX(LUU_NIEN!$BG$30:$BG$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1,OR(INDEX(LUU_NIEN!$BH$30:$BH$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1,INDEX(LUU_NI
    #   EN!$BJ$30:$BJ$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1),INDEX(LUU_NIEN!$BI$30:$BI$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1,OR(INDEX(LUU_NIEN!$BM$3
    #   0:$BM$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))>=1,IF(DOI_CHIEU!$C$33=2,INDEX(LUU_NIEN!$AM$30:$AM$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0)),INDEX(LUU_NI
    #   EN!$AB$30:$AB$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0)))=1),OR(LEFT(IF(DOI_CHIEU!$C$33=2,INDEX(LUU_NIEN!$AL$30:$AL$129,MATCH($C11,LUU_NIEN!$A$30:$A$129
    #   ,0)),INDEX(LUU_NIEN!$AA$30:$AA$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))),3)="Tốt",ISNUMBER(SEARCH(": Tốt",IF(DOI_CHIEU!$C$33=2,INDEX(LUU_NIEN!$AL$30:$
    #   AL$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0)),INDEX(LUU_NIEN!$AA$30:$AA$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0)))))),OR(LEFT(IF(DOI_CHIEU!$C$33=2,INDEX(
    #   LUU_NIEN!$AL$30:$AL$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0)),INDEX(LUU_NIEN!$AA$30:$AA$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))),3)="Xấu",ISNUMBER(SEA
    #   RCH(": Xấu",IF(DOI_CHIEU!$C$33=2,INDEX(LUU_NIEN!$AL$30:$AL$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0)),INDEX(LUU_NIEN!$AA$30:$AA$129,MATCH($C11,LUU_NIEN!
    #   $A$30:$A$129,0))))),INDEX(LUU_NIEN!$BL$30:$BL$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))=1),OR(INDEX(LUU_NIEN!$O$30:$O$129,MATCH($C11,LUU_NIEN!$A$30:$A$
    #   129,0))="Thực thần",INDEX(LUU_NIEN!$O$30:$O$129,MATCH($C11,LUU_NIEN!$A$30:$A$129,0))="Thương quan")),"Khớp","Không"),"ngoài 100 năm"))
    return ('' if X.truthy(X.OR(X.eq(E.v(6,r,3),''), X.eq(E.v(6,r,4),''))) else X.IFERROR(lambda: ('Khớp' if X.truthy(X.CHOOSE(X.MATCH(E.v(6,r,4), X.Arr([['Kết hôn', 'Ly hôn – vợ chồng bất hòa', 'Mất cha – cha bệnh nạn', 'Mất mẹ – mẹ bệnh nạn', 'Tai nạn – bệnh nặng', 'Thăng tiến – phát tài', 'Thất bại – phá sản – kiện tụng', 'Sinh con']]), 0.0), lambda: X.OR(X.eq(X.INDEX(E.rg(15,30,57,129,57), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0), X.eq(X.INDEX(E.rg(15,30,58,129,58), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0), X.eq(X.INDEX(E.rg(15,30,53,129,53), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0)), lambda: X.eq(X.INDEX(E.rg(15,30,59,129,59), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0), lambda: X.OR(X.eq(X.INDEX(E.rg(15,30,60,129,60), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0), X.eq(X.INDEX(E.rg(15,30,62,129,62), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0)), lambda: X.eq(X.INDEX(E.rg(15,30,61,129,61), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0), lambda: X.OR(X.ge(X.INDEX(E.rg(15,30,65,129,65), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0), X.eq((X.INDEX(E.rg(15,30,39,129,39), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(15,30,28,129,28), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0))),1.0)), lambda: X.OR(X.eq(X.LEFT((X.INDEX(E.rg(15,30,38,129,38), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(15,30,27,129,27), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0))), 3.0),'Tốt'), X.ISNUMBER(lambda: X.SEARCH(': Tốt', (X.INDEX(E.rg(15,30,38,129,38), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(15,30,27,129,27), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)))))), lambda: X.OR(X.eq(X.LEFT((X.INDEX(E.rg(15,30,38,129,38), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(15,30,27,129,27), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0))), 3.0),'Xấu'), X.ISNUMBER(lambda: X.SEARCH(': Xấu', (X.INDEX(E.rg(15,30,38,129,38), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.INDEX(E.rg(15,30,27,129,27), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0))))), X.eq(X.INDEX(E.rg(15,30,64,129,64), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),1.0)), lambda: X.OR(X.eq(X.INDEX(E.rg(15,30,15,129,15), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),'Thực thần'), X.eq(X.INDEX(E.rg(15,30,15,129,15), X.MATCH(E.v(6,r,3), E.rg(15,30,1,129,1), 0.0)),'Thương quan')))) else 'Không'), lambda: 'ngoài 100 năm'))

def f3(E, r, c):
    # Ô G11, G12, G13, G14, G15, G16, G17, G18, G19, G20, G21, G22
    # = IF($D11="","",CHOOSE(MATCH($D11,{"Kết hôn","Ly hôn – vợ chồng bất hòa","Mất cha – cha bệnh nạn","Mất mẹ – mẹ bệnh nạn","Tai nạn – bệnh nặng","Thăng ti
    #   ến – phát tài","Thất bại – phá sản – kiện tụng","Sinh con"},0),"năm thuận lợi kết hôn / tam hợp tài / lưu niên hợp trụ ngày (P1 tr.201; P2 tr.289, 463
    #   )","bất lợi hôn nhân, xung cung phối ngẫu, năm thương quan (P1 tr.156; P2 tr.470-471)","dấu hiệu bất lợi cho cha, kiêu đoạt thực (P2 tr.440, 446)","dấ
    #   u hiệu bất lợi cho mẹ (P2 tr.440)","dấu hiệu nặng: thiên khắc địa xung, nhật can nhập mộ, vận kỵ + năm kỵ (P1 tr.201-202; P2 tr.282-288)","năm tổng hợ
    #   p tốt theo bảng P1 tr.133","năm tổng hợp xấu (P1 tr.133), lộc bị xung (P1 tr.139)","năm thực thương – sao con (P1 tr.157); giả định G15, không phân bi
    #   ệt được giờ"))
    return ('' if X.truthy(X.eq(E.v(6,r,4),'')) else X.CHOOSE(X.MATCH(E.v(6,r,4), X.Arr([['Kết hôn', 'Ly hôn – vợ chồng bất hòa', 'Mất cha – cha bệnh nạn', 'Mất mẹ – mẹ bệnh nạn', 'Tai nạn – bệnh nặng', 'Thăng tiến – phát tài', 'Thất bại – phá sản – kiện tụng', 'Sinh con']]), 0.0), lambda: 'năm thuận lợi kết hôn / tam hợp tài / lưu niên hợp trụ ngày (P1 tr.201; P2 tr.289, 463)', lambda: 'bất lợi hôn nhân, xung cung phối ngẫu, năm thương quan (P1 tr.156; P2 tr.470-471)', lambda: 'dấu hiệu bất lợi cho cha, kiêu đoạt thực (P2 tr.440, 446)', lambda: 'dấu hiệu bất lợi cho mẹ (P2 tr.440)', lambda: 'dấu hiệu nặng: thiên khắc địa xung, nhật can nhập mộ, vận kỵ + năm kỵ (P1 tr.201-202; P2 tr.282-288)', lambda: 'năm tổng hợp tốt theo bảng P1 tr.133', lambda: 'năm tổng hợp xấu (P1 tr.133), lộc bị xung (P1 tr.139)', lambda: 'năm thực thương – sao con (P1 tr.157); giả định G15, không phân biệt được giờ'))

def f4(E, r, c):
    # Ô E23
    # = COUNTIFS(C11:C22,">0",D11:D22,"<>")
    return X.COUNTIFS(E.rg(6,r-12,c-2,r-1,c-2), '>0', E.rg(6,r-12,c-1,r-1,c-1), '<>')

def f5(E, r, c):
    # Ô F23
    # = COUNTIF(F11:F22,"Khớp")
    return X.COUNTIF(E.rg(6,r-12,c,r-1,c), 'Khớp')

def f6(E, r, c):
    # Ô C27, C117
    # = NHAP_LIEU!$F$23
    return E.v(2,23,6)

def f7(E, r, c):
    # Ô D27, D117
    # = DUNG_HY_KY!$C$6
    return E.v(8,6,3)

def f8(E, r, c):
    # Ô E27, C131
    # = PP2_DINH_LUONG!$C$101
    return E.v(10,101,3)

def f9(E, r, c):
    # Ô F27, D131
    # = IF(DOI_CHIEU!$C$33=2,DOI_CHIEU!$C$32,DOI_CHIEU!$C$31)
    return (E.v(7,32,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(7,31,3))

def f10(E, r, c):
    # Ô G27, C145
    # = DOI_CHIEU!$C$35
    return E.v(7,35,3)

def f11(E, r, c):
    # Ô H27, D145
    # = IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)>=0,INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1),"—")
    return (X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0)) if X.truthy(X.ge((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0)) else '—')

def f12(E, r, c):
    # Ô I27, C159
    # = IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)>=0,INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1),"—")
    return (X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0)) if X.truthy(X.ge((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else '—')

def f13(E, r, c):
    # Ô J27, D159
    # = IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1),"—")
    return (X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0)) if X.truthy(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0)) else '—')

def f14(E, r, c):
    # Ô K27, C173
    # = IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$102,DUNG_HY_KY!$C$102)
    return (E.v(8,102,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,102,3))

def f15(E, r, c):
    # Ô L27
    # = IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C69,DAI_VAN!C54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C69,DAI_VAN!C54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DA
    #   I_VAN!D69,DAI_VAN!D54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!D69,DAI_VAN!D54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E69,DAI_VAN!E54)="Tố
    #   t","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E69,DAI_VAN!E54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F69,DAI_VAN!F54)="Tốt","T",IF(IF(DOI_CHIEU!$C$
    #   33=2,DAI_VAN!F69,DAI_VAN!F54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G69,DAI_VAN!G54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G69,DAI_VAN!G
    #   54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!H69,DAI_VAN!H54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!H69,DAI_VAN!H54)="Xấu","X","·"))&IF(IF(
    #   DOI_CHIEU!$C$33=2,DAI_VAN!I69,DAI_VAN!I54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I69,DAI_VAN!I54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!
    #   J69,DAI_VAN!J54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J69,DAI_VAN!J54)="Xấu","X","·"))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('T' if X.truthy(X.eq((E.v(14,r+42,c-9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-9)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-9)),'Xấu')) else '·')),('T' if X.truthy(X.eq((E.v(14,r+42,c-8) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-8)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-8) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-8)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r+42,c-7) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-7)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-7) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-7)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r+42,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-6)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-6)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r+42,c-5) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-5)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-5) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-5)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r+42,c-4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-4)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-4)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r+42,c-3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-3)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-3)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r+42,c-2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-2)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r+42,c-2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r+27,c-2)),'Xấu')) else '·')))

def f16(E, r, c):
    # Ô M27, C187
    # = INDEX({"Không xếp","Nghèo","Vất vả","Trung bình","Khá giả","Giàu lớn"},TAI_LOC!$E$16+1)
    return X.INDEX(X.Arr([['Không xếp', 'Nghèo', 'Vất vả', 'Trung bình', 'Khá giả', 'Giàu lớn']]), X.add(E.v(19,16,5),1.0))

def f17(E, r, c):
    # Ô N27, D187
    # = LUC_THAN!$X$30
    return E.v(16,30,24)

def f18(E, r, c):
    # Ô O27, C201
    # = LUC_THAN!$X$31
    return E.v(16,31,24)

def f19(E, r, c):
    # Ô P27, D201
    # = LUC_THAN!$X$32
    return E.v(16,32,24)

def f20(E, r, c):
    # Ô Q27, C215
    # = LUC_THAN!$X$33
    return E.v(16,33,24)

def f21(E, r, c):
    # Ô R27, D215
    # = LUC_THAN!$X$34
    return E.v(16,34,24)

def f22(E, r, c):
    # Ô S27, C229
    # = HON_NHAN!$X$64
    return E.v(17,64,24)

def f23(E, r, c):
    # Ô T27, D229
    # = MENH_CUC!$U$18
    return E.v(18,18,21)

def f24(E, r, c):
    # Ô U27, C243
    # = MENH_CUC!$U$19
    return E.v(18,19,21)

def f25(E, r, c):
    # Ô V27, D243
    # = MENH_CUC!$U$20
    return E.v(18,20,21)

def f26(E, r, c):
    # Ô W27, C257
    # = MENH_CUC!$U$21
    return E.v(18,21,21)

def f27(E, r, c):
    # Ô X27, D257
    # = INDEX(HANH_TEN,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+1)&" – "&TO_CHAT!$D$7
    return X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),1.0)),' – '),E.v(20,7,4))

def f28(E, r, c):
    # Ô Y27, C271
    # = $F$23
    return E.v(6,23,6)

def f29(E, r, c):
    # Ô Z27, D271
    # = TRIM(IF($F$11="Khớp",$C$11&" ","")&IF($F$12="Khớp",$C$12&" ","")&IF($F$13="Khớp",$C$13&" ","")&IF($F$14="Khớp",$C$14&" ","")&IF($F$15="Khớp",$C$15&" "
    #   ,"")&IF($F$16="Khớp",$C$16&" ","")&IF($F$17="Khớp",$C$17&" ","")&IF($F$18="Khớp",$C$18&" ","")&IF($F$19="Khớp",$C$19&" ","")&IF($F$20="Khớp",$C$20&" "
    #   ,"")&IF($F$21="Khớp",$C$21&" ","")&IF($F$22="Khớp",$C$22&" ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(6,11,3),' ') if X.truthy(X.eq(E.v(6,11,6),'Khớp')) else ''),(X.cat(E.v(6,12,3),' ') if X.truthy(X.eq(E.v(6,12,6),'Khớp')) else '')),(X.cat(E.v(6,13,3),' ') if X.truthy(X.eq(E.v(6,13,6),'Khớp')) else '')),(X.cat(E.v(6,14,3),' ') if X.truthy(X.eq(E.v(6,14,6),'Khớp')) else '')),(X.cat(E.v(6,15,3),' ') if X.truthy(X.eq(E.v(6,15,6),'Khớp')) else '')),(X.cat(E.v(6,16,3),' ') if X.truthy(X.eq(E.v(6,16,6),'Khớp')) else '')),(X.cat(E.v(6,17,3),' ') if X.truthy(X.eq(E.v(6,17,6),'Khớp')) else '')),(X.cat(E.v(6,18,3),' ') if X.truthy(X.eq(E.v(6,18,6),'Khớp')) else '')),(X.cat(E.v(6,19,3),' ') if X.truthy(X.eq(E.v(6,19,6),'Khớp')) else '')),(X.cat(E.v(6,20,3),' ') if X.truthy(X.eq(E.v(6,20,6),'Khớp')) else '')),(X.cat(E.v(6,21,3),' ') if X.truthy(X.eq(E.v(6,21,6),'Khớp')) else '')),(X.cat(E.v(6,22,3),' ') if X.truthy(X.eq(E.v(6,22,6),'Khớp')) else '')))

def f30(E, r, c):
    # Ô C28, D28, C29, D29, C30, D30, C31, D31, C32, D32, C33, D33 … (24 ô)
    # = C118
    return E.v(6,r+90,c)

def f31(E, r, c):
    # Ô E28, F28, E29, F29, E30, F30, E31, F31, E32, F32, E33, F33 … (24 ô)
    # = C132
    return E.v(6,r+104,c-2)

def f32(E, r, c):
    # Ô G28, H28, G29, H29, G30, H30, G31, H31, G32, H32, G33, H33 … (24 ô)
    # = C146
    return E.v(6,r+118,c-4)

def f33(E, r, c):
    # Ô I28, J28, I29, J29, I30, J30, I31, J31, I32, J32, I33, J33 … (24 ô)
    # = C160
    return E.v(6,r+132,c-6)

def f34(E, r, c):
    # Ô K28, L28, K29, L29, K30, L30, K31, L31, K32, L32, K33, L33 … (24 ô)
    # = C174
    return E.v(6,r+146,c-8)

def f35(E, r, c):
    # Ô M28, N28, M29, N29, M30, N30, M31, N31, M32, N32, M33, N33 … (24 ô)
    # = C188
    return E.v(6,r+160,c-10)

def f36(E, r, c):
    # Ô O28, P28, O29, P29, O30, P30, O31, P31, O32, P32, O33, P33 … (24 ô)
    # = C202
    return E.v(6,r+174,c-12)

def f37(E, r, c):
    # Ô Q28, R28, Q29, R29, Q30, R30, Q31, R31, Q32, R32, Q33, R33 … (24 ô)
    # = C216
    return E.v(6,r+188,c-14)

def f38(E, r, c):
    # Ô S28, T28, S29, T29, S30, T30, S31, T31, S32, T32, S33, T33 … (24 ô)
    # = C230
    return E.v(6,r+202,c-16)

def f39(E, r, c):
    # Ô U28, V28, U29, V29, U30, V30, U31, V31, U32, V32, U33, V33 … (24 ô)
    # = C244
    return E.v(6,r+216,c-18)

def f40(E, r, c):
    # Ô W28, X28, W29, X29, W30, X30, W31, X31, W32, X32, W33, X33 … (24 ô)
    # = C258
    return E.v(6,r+230,c-20)

def f41(E, r, c):
    # Ô Y28, Z28, Y29, Z29, Y30, Z30, Y31, Z31, Y32, Z32, Y33, Z33 … (24 ô)
    # = C272
    return E.v(6,r+244,c-22)

def f42(E, r, c):
    # Ô AB28
    # = N(Y28)*100+(12-0)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,0.0))

def f43(E, r, c):
    # Ô AB29
    # = N(Y29)*100+(12-1)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,1.0))

def f44(E, r, c):
    # Ô AB30
    # = N(Y30)*100+(12-2)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,2.0))

def f45(E, r, c):
    # Ô AB31
    # = N(Y31)*100+(12-3)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,3.0))

def f46(E, r, c):
    # Ô AB32
    # = N(Y32)*100+(12-4)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,4.0))

def f47(E, r, c):
    # Ô AB33
    # = N(Y33)*100+(12-5)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,5.0))

def f48(E, r, c):
    # Ô AB34
    # = N(Y34)*100+(12-6)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,6.0))

def f49(E, r, c):
    # Ô AB35
    # = N(Y35)*100+(12-7)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,7.0))

def f50(E, r, c):
    # Ô AB36
    # = N(Y36)*100+(12-8)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,8.0))

def f51(E, r, c):
    # Ô AB37
    # = N(Y37)*100+(12-9)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,9.0))

def f52(E, r, c):
    # Ô AB38
    # = N(Y38)*100+(12-10)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,10.0))

def f53(E, r, c):
    # Ô AB39
    # = N(Y39)*100+(12-11)
    return X.add(X.mul(X.N(E.v(6,r,c-3)),100.0),X.sub(12.0,11.0))

def f54(E, r, c):
    # Ô C43
    # = IF(COUNTA($C$28:$C$39)=0,"",IF(COUNTIF($C$28:$C$39,C28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,3,39,3)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-15,c)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f55(E, r, c):
    # Ô D43
    # = TRIM(IF(COUNTIF(C$28:C28,C28)=1,C28&" ("&COUNTIF($C$28:$C$39,C28)&"/12); ","")&IF(COUNTIF(C$28:C29,C29)=1,C29&" ("&COUNTIF($C$28:$C$39,C29)&"/12); ","
    #   ")&IF(COUNTIF(C$28:C30,C30)=1,C30&" ("&COUNTIF($C$28:$C$39,C30)&"/12); ","")&IF(COUNTIF(C$28:C31,C31)=1,C31&" ("&COUNTIF($C$28:$C$39,C31)&"/12); ","")
    #   &IF(COUNTIF(C$28:C32,C32)=1,C32&" ("&COUNTIF($C$28:$C$39,C32)&"/12); ","")&IF(COUNTIF(C$28:C33,C33)=1,C33&" ("&COUNTIF($C$28:$C$39,C33)&"/12); ","")&I
    #   F(COUNTIF(C$28:C34,C34)=1,C34&" ("&COUNTIF($C$28:$C$39,C34)&"/12); ","")&IF(COUNTIF(C$28:C35,C35)=1,C35&" ("&COUNTIF($C$28:$C$39,C35)&"/12); ","")&IF(
    #   COUNTIF(C$28:C36,C36)=1,C36&" ("&COUNTIF($C$28:$C$39,C36)&"/12); ","")&IF(COUNTIF(C$28:C37,C37)=1,C37&" ("&COUNTIF($C$28:$C$39,C37)&"/12); ","")&IF(CO
    #   UNTIF(C$28:C38,C38)=1,C38&" ("&COUNTIF($C$28:$C$39,C38)&"/12); ","")&IF(COUNTIF(C$28:C39,C39)=1,C39&" ("&COUNTIF($C$28:$C$39,C39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-15,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-15,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-15,c-1), E.v(6,r-15,c-1)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-14,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-14,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-14,c-1), E.v(6,r-14,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-13,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-13,c-1), E.v(6,r-13,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-12,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-12,c-1), E.v(6,r-12,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-11,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-11,c-1), E.v(6,r-11,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-10,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-10,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-10,c-1), E.v(6,r-10,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-9,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-9,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-9,c-1), E.v(6,r-9,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-8,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-8,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-8,c-1), E.v(6,r-8,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-7,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-7,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-7,c-1), E.v(6,r-7,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-6,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-6,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-6,c-1), E.v(6,r-6,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-5,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-5,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-5,c-1), E.v(6,r-5,c-1)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-4,c-1),' ('),X.COUNTIF(E.rg(6,28,3,39,3), E.v(6,r-4,c-1))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c-1,r-4,c-1), E.v(6,r-4,c-1)),1.0)) else '')))

def f56(E, r, c):
    # Ô C44
    # = IF(COUNTA($D$28:$D$39)=0,"",IF(COUNTIF($D$28:$D$39,D28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,4,39,4)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-16,c+1)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f57(E, r, c):
    # Ô D44
    # = TRIM(IF(COUNTIF(D$28:D28,D28)=1,D28&" ("&COUNTIF($D$28:$D$39,D28)&"/12); ","")&IF(COUNTIF(D$28:D29,D29)=1,D29&" ("&COUNTIF($D$28:$D$39,D29)&"/12); ","
    #   ")&IF(COUNTIF(D$28:D30,D30)=1,D30&" ("&COUNTIF($D$28:$D$39,D30)&"/12); ","")&IF(COUNTIF(D$28:D31,D31)=1,D31&" ("&COUNTIF($D$28:$D$39,D31)&"/12); ","")
    #   &IF(COUNTIF(D$28:D32,D32)=1,D32&" ("&COUNTIF($D$28:$D$39,D32)&"/12); ","")&IF(COUNTIF(D$28:D33,D33)=1,D33&" ("&COUNTIF($D$28:$D$39,D33)&"/12); ","")&I
    #   F(COUNTIF(D$28:D34,D34)=1,D34&" ("&COUNTIF($D$28:$D$39,D34)&"/12); ","")&IF(COUNTIF(D$28:D35,D35)=1,D35&" ("&COUNTIF($D$28:$D$39,D35)&"/12); ","")&IF(
    #   COUNTIF(D$28:D36,D36)=1,D36&" ("&COUNTIF($D$28:$D$39,D36)&"/12); ","")&IF(COUNTIF(D$28:D37,D37)=1,D37&" ("&COUNTIF($D$28:$D$39,D37)&"/12); ","")&IF(CO
    #   UNTIF(D$28:D38,D38)=1,D38&" ("&COUNTIF($D$28:$D$39,D38)&"/12); ","")&IF(COUNTIF(D$28:D39,D39)=1,D39&" ("&COUNTIF($D$28:$D$39,D39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-16,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-16,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-16,c), E.v(6,r-16,c)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-15,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-15,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-15,c), E.v(6,r-15,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-14,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-14,c), E.v(6,r-14,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-13,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-13,c), E.v(6,r-13,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-12,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-12,c), E.v(6,r-12,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-11,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-11,c), E.v(6,r-11,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-10,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-10,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-10,c), E.v(6,r-10,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-9,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-9,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-9,c), E.v(6,r-9,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-8,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-8,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-8,c), E.v(6,r-8,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-7,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-7,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-7,c), E.v(6,r-7,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-6,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-6,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-6,c), E.v(6,r-6,c)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-5,c),' ('),X.COUNTIF(E.rg(6,28,4,39,4), E.v(6,r-5,c))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c,r-5,c), E.v(6,r-5,c)),1.0)) else '')))

def f58(E, r, c):
    # Ô C45
    # = IF(COUNT($E$28:$E$39)=0,"",IF(OR(MAX($E$28:$E$39)<P_NGUONG,MIN($E$28:$E$39)>=P_NGUONG),"ỔN ĐỊNH (cùng phía ngưỡng 40%)","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNT(E.rg(6,28,5,39,5)),0.0)) else ('ỔN ĐỊNH (cùng phía ngưỡng 40%)' if X.truthy(X.OR(X.lt(X.MAX(E.rg(6,28,5,39,5)),E.v(24,95,2)), X.ge(X.MIN(E.rg(6,28,5,39,5)),E.v(24,95,2)))) else 'PHỤ THUỘC GIỜ'))

def f59(E, r, c):
    # Ô D45
    # = IF(COUNT($E$28:$E$39)=0,"","từ "&TEXT(MIN($E$28:$E$39),"0.0%")&" đến "&TEXT(MAX($E$28:$E$39),"0.0%"))
    return ('' if X.truthy(X.eq(X.COUNT(E.rg(6,28,5,39,5)),0.0)) else X.cat(X.cat(X.cat('từ ',X.TEXT(X.MIN(E.rg(6,28,5,39,5)), '0.0%')),' đến '),X.TEXT(X.MAX(E.rg(6,28,5,39,5)), '0.0%')))

def f60(E, r, c):
    # Ô C46
    # = IF(COUNTA($F$28:$F$39)=0,"",IF(COUNTIF($F$28:$F$39,F28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,6,39,6)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-18,c+3)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f61(E, r, c):
    # Ô D46
    # = TRIM(IF(COUNTIF(F$28:F28,F28)=1,F28&" ("&COUNTIF($F$28:$F$39,F28)&"/12); ","")&IF(COUNTIF(F$28:F29,F29)=1,F29&" ("&COUNTIF($F$28:$F$39,F29)&"/12); ","
    #   ")&IF(COUNTIF(F$28:F30,F30)=1,F30&" ("&COUNTIF($F$28:$F$39,F30)&"/12); ","")&IF(COUNTIF(F$28:F31,F31)=1,F31&" ("&COUNTIF($F$28:$F$39,F31)&"/12); ","")
    #   &IF(COUNTIF(F$28:F32,F32)=1,F32&" ("&COUNTIF($F$28:$F$39,F32)&"/12); ","")&IF(COUNTIF(F$28:F33,F33)=1,F33&" ("&COUNTIF($F$28:$F$39,F33)&"/12); ","")&I
    #   F(COUNTIF(F$28:F34,F34)=1,F34&" ("&COUNTIF($F$28:$F$39,F34)&"/12); ","")&IF(COUNTIF(F$28:F35,F35)=1,F35&" ("&COUNTIF($F$28:$F$39,F35)&"/12); ","")&IF(
    #   COUNTIF(F$28:F36,F36)=1,F36&" ("&COUNTIF($F$28:$F$39,F36)&"/12); ","")&IF(COUNTIF(F$28:F37,F37)=1,F37&" ("&COUNTIF($F$28:$F$39,F37)&"/12); ","")&IF(CO
    #   UNTIF(F$28:F38,F38)=1,F38&" ("&COUNTIF($F$28:$F$39,F38)&"/12); ","")&IF(COUNTIF(F$28:F39,F39)=1,F39&" ("&COUNTIF($F$28:$F$39,F39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-18,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-18,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-18,c+2), E.v(6,r-18,c+2)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-17,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-17,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-17,c+2), E.v(6,r-17,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-16,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-16,c+2), E.v(6,r-16,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-15,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-15,c+2), E.v(6,r-15,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-14,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-14,c+2), E.v(6,r-14,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-13,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-13,c+2), E.v(6,r-13,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-12,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-12,c+2), E.v(6,r-12,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-11,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-11,c+2), E.v(6,r-11,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-10,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-10,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-10,c+2), E.v(6,r-10,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-9,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-9,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-9,c+2), E.v(6,r-9,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-8,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-8,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-8,c+2), E.v(6,r-8,c+2)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-7,c+2),' ('),X.COUNTIF(E.rg(6,28,6,39,6), E.v(6,r-7,c+2))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+2,r-7,c+2), E.v(6,r-7,c+2)),1.0)) else '')))

def f62(E, r, c):
    # Ô C47
    # = IF(COUNTA($G$28:$G$39)=0,"",IF(COUNTIF($G$28:$G$39,G28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,7,39,7)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-19,c+4)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f63(E, r, c):
    # Ô D47
    # = TRIM(IF(COUNTIF(G$28:G28,G28)=1,G28&" ("&COUNTIF($G$28:$G$39,G28)&"/12); ","")&IF(COUNTIF(G$28:G29,G29)=1,G29&" ("&COUNTIF($G$28:$G$39,G29)&"/12); ","
    #   ")&IF(COUNTIF(G$28:G30,G30)=1,G30&" ("&COUNTIF($G$28:$G$39,G30)&"/12); ","")&IF(COUNTIF(G$28:G31,G31)=1,G31&" ("&COUNTIF($G$28:$G$39,G31)&"/12); ","")
    #   &IF(COUNTIF(G$28:G32,G32)=1,G32&" ("&COUNTIF($G$28:$G$39,G32)&"/12); ","")&IF(COUNTIF(G$28:G33,G33)=1,G33&" ("&COUNTIF($G$28:$G$39,G33)&"/12); ","")&I
    #   F(COUNTIF(G$28:G34,G34)=1,G34&" ("&COUNTIF($G$28:$G$39,G34)&"/12); ","")&IF(COUNTIF(G$28:G35,G35)=1,G35&" ("&COUNTIF($G$28:$G$39,G35)&"/12); ","")&IF(
    #   COUNTIF(G$28:G36,G36)=1,G36&" ("&COUNTIF($G$28:$G$39,G36)&"/12); ","")&IF(COUNTIF(G$28:G37,G37)=1,G37&" ("&COUNTIF($G$28:$G$39,G37)&"/12); ","")&IF(CO
    #   UNTIF(G$28:G38,G38)=1,G38&" ("&COUNTIF($G$28:$G$39,G38)&"/12); ","")&IF(COUNTIF(G$28:G39,G39)=1,G39&" ("&COUNTIF($G$28:$G$39,G39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-19,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-19,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-19,c+3), E.v(6,r-19,c+3)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-18,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-18,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-18,c+3), E.v(6,r-18,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-17,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-17,c+3), E.v(6,r-17,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-16,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-16,c+3), E.v(6,r-16,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-15,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-15,c+3), E.v(6,r-15,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-14,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-14,c+3), E.v(6,r-14,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-13,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-13,c+3), E.v(6,r-13,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-12,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-12,c+3), E.v(6,r-12,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-11,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-11,c+3), E.v(6,r-11,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-10,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-10,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-10,c+3), E.v(6,r-10,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-9,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-9,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-9,c+3), E.v(6,r-9,c+3)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-8,c+3),' ('),X.COUNTIF(E.rg(6,28,7,39,7), E.v(6,r-8,c+3))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+3,r-8,c+3), E.v(6,r-8,c+3)),1.0)) else '')))

def f64(E, r, c):
    # Ô C48
    # = IF(COUNTA($H$28:$H$39)=0,"",IF(COUNTIF($H$28:$H$39,H28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,8,39,8)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-20,c+5)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f65(E, r, c):
    # Ô D48
    # = TRIM(IF(COUNTIF(H$28:H28,H28)=1,H28&" ("&COUNTIF($H$28:$H$39,H28)&"/12); ","")&IF(COUNTIF(H$28:H29,H29)=1,H29&" ("&COUNTIF($H$28:$H$39,H29)&"/12); ","
    #   ")&IF(COUNTIF(H$28:H30,H30)=1,H30&" ("&COUNTIF($H$28:$H$39,H30)&"/12); ","")&IF(COUNTIF(H$28:H31,H31)=1,H31&" ("&COUNTIF($H$28:$H$39,H31)&"/12); ","")
    #   &IF(COUNTIF(H$28:H32,H32)=1,H32&" ("&COUNTIF($H$28:$H$39,H32)&"/12); ","")&IF(COUNTIF(H$28:H33,H33)=1,H33&" ("&COUNTIF($H$28:$H$39,H33)&"/12); ","")&I
    #   F(COUNTIF(H$28:H34,H34)=1,H34&" ("&COUNTIF($H$28:$H$39,H34)&"/12); ","")&IF(COUNTIF(H$28:H35,H35)=1,H35&" ("&COUNTIF($H$28:$H$39,H35)&"/12); ","")&IF(
    #   COUNTIF(H$28:H36,H36)=1,H36&" ("&COUNTIF($H$28:$H$39,H36)&"/12); ","")&IF(COUNTIF(H$28:H37,H37)=1,H37&" ("&COUNTIF($H$28:$H$39,H37)&"/12); ","")&IF(CO
    #   UNTIF(H$28:H38,H38)=1,H38&" ("&COUNTIF($H$28:$H$39,H38)&"/12); ","")&IF(COUNTIF(H$28:H39,H39)=1,H39&" ("&COUNTIF($H$28:$H$39,H39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-20,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-20,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-20,c+4), E.v(6,r-20,c+4)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-19,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-19,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-19,c+4), E.v(6,r-19,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-18,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-18,c+4), E.v(6,r-18,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-17,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-17,c+4), E.v(6,r-17,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-16,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-16,c+4), E.v(6,r-16,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-15,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-15,c+4), E.v(6,r-15,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-14,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-14,c+4), E.v(6,r-14,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-13,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-13,c+4), E.v(6,r-13,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-12,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-12,c+4), E.v(6,r-12,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-11,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-11,c+4), E.v(6,r-11,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-10,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-10,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-10,c+4), E.v(6,r-10,c+4)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-9,c+4),' ('),X.COUNTIF(E.rg(6,28,8,39,8), E.v(6,r-9,c+4))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+4,r-9,c+4), E.v(6,r-9,c+4)),1.0)) else '')))

def f66(E, r, c):
    # Ô C49
    # = IF(COUNTA($I$28:$I$39)=0,"",IF(COUNTIF($I$28:$I$39,I28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,9,39,9)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-21,c+6)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f67(E, r, c):
    # Ô D49
    # = TRIM(IF(COUNTIF(I$28:I28,I28)=1,I28&" ("&COUNTIF($I$28:$I$39,I28)&"/12); ","")&IF(COUNTIF(I$28:I29,I29)=1,I29&" ("&COUNTIF($I$28:$I$39,I29)&"/12); ","
    #   ")&IF(COUNTIF(I$28:I30,I30)=1,I30&" ("&COUNTIF($I$28:$I$39,I30)&"/12); ","")&IF(COUNTIF(I$28:I31,I31)=1,I31&" ("&COUNTIF($I$28:$I$39,I31)&"/12); ","")
    #   &IF(COUNTIF(I$28:I32,I32)=1,I32&" ("&COUNTIF($I$28:$I$39,I32)&"/12); ","")&IF(COUNTIF(I$28:I33,I33)=1,I33&" ("&COUNTIF($I$28:$I$39,I33)&"/12); ","")&I
    #   F(COUNTIF(I$28:I34,I34)=1,I34&" ("&COUNTIF($I$28:$I$39,I34)&"/12); ","")&IF(COUNTIF(I$28:I35,I35)=1,I35&" ("&COUNTIF($I$28:$I$39,I35)&"/12); ","")&IF(
    #   COUNTIF(I$28:I36,I36)=1,I36&" ("&COUNTIF($I$28:$I$39,I36)&"/12); ","")&IF(COUNTIF(I$28:I37,I37)=1,I37&" ("&COUNTIF($I$28:$I$39,I37)&"/12); ","")&IF(CO
    #   UNTIF(I$28:I38,I38)=1,I38&" ("&COUNTIF($I$28:$I$39,I38)&"/12); ","")&IF(COUNTIF(I$28:I39,I39)=1,I39&" ("&COUNTIF($I$28:$I$39,I39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-21,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-21,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-21,c+5), E.v(6,r-21,c+5)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-20,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-20,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-20,c+5), E.v(6,r-20,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-19,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-19,c+5), E.v(6,r-19,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-18,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-18,c+5), E.v(6,r-18,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-17,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-17,c+5), E.v(6,r-17,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-16,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-16,c+5), E.v(6,r-16,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-15,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-15,c+5), E.v(6,r-15,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-14,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-14,c+5), E.v(6,r-14,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-13,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-13,c+5), E.v(6,r-13,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-12,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-12,c+5), E.v(6,r-12,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-11,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-11,c+5), E.v(6,r-11,c+5)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-10,c+5),' ('),X.COUNTIF(E.rg(6,28,9,39,9), E.v(6,r-10,c+5))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+5,r-10,c+5), E.v(6,r-10,c+5)),1.0)) else '')))

def f68(E, r, c):
    # Ô C50
    # = IF(COUNTA($J$28:$J$39)=0,"",IF(COUNTIF($J$28:$J$39,J28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,10,39,10)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-22,c+7)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f69(E, r, c):
    # Ô D50
    # = TRIM(IF(COUNTIF(J$28:J28,J28)=1,J28&" ("&COUNTIF($J$28:$J$39,J28)&"/12); ","")&IF(COUNTIF(J$28:J29,J29)=1,J29&" ("&COUNTIF($J$28:$J$39,J29)&"/12); ","
    #   ")&IF(COUNTIF(J$28:J30,J30)=1,J30&" ("&COUNTIF($J$28:$J$39,J30)&"/12); ","")&IF(COUNTIF(J$28:J31,J31)=1,J31&" ("&COUNTIF($J$28:$J$39,J31)&"/12); ","")
    #   &IF(COUNTIF(J$28:J32,J32)=1,J32&" ("&COUNTIF($J$28:$J$39,J32)&"/12); ","")&IF(COUNTIF(J$28:J33,J33)=1,J33&" ("&COUNTIF($J$28:$J$39,J33)&"/12); ","")&I
    #   F(COUNTIF(J$28:J34,J34)=1,J34&" ("&COUNTIF($J$28:$J$39,J34)&"/12); ","")&IF(COUNTIF(J$28:J35,J35)=1,J35&" ("&COUNTIF($J$28:$J$39,J35)&"/12); ","")&IF(
    #   COUNTIF(J$28:J36,J36)=1,J36&" ("&COUNTIF($J$28:$J$39,J36)&"/12); ","")&IF(COUNTIF(J$28:J37,J37)=1,J37&" ("&COUNTIF($J$28:$J$39,J37)&"/12); ","")&IF(CO
    #   UNTIF(J$28:J38,J38)=1,J38&" ("&COUNTIF($J$28:$J$39,J38)&"/12); ","")&IF(COUNTIF(J$28:J39,J39)=1,J39&" ("&COUNTIF($J$28:$J$39,J39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-22,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-22,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-22,c+6), E.v(6,r-22,c+6)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-21,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-21,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-21,c+6), E.v(6,r-21,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-20,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-20,c+6), E.v(6,r-20,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-19,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-19,c+6), E.v(6,r-19,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-18,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-18,c+6), E.v(6,r-18,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-17,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-17,c+6), E.v(6,r-17,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-16,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-16,c+6), E.v(6,r-16,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-15,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-15,c+6), E.v(6,r-15,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-14,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-14,c+6), E.v(6,r-14,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-13,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-13,c+6), E.v(6,r-13,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-12,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-12,c+6), E.v(6,r-12,c+6)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-11,c+6),' ('),X.COUNTIF(E.rg(6,28,10,39,10), E.v(6,r-11,c+6))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+6,r-11,c+6), E.v(6,r-11,c+6)),1.0)) else '')))

def f70(E, r, c):
    # Ô C51
    # = IF(COUNTA($K$28:$K$39)=0,"",IF(COUNTIF($K$28:$K$39,K28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,11,39,11)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-23,c+8)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f71(E, r, c):
    # Ô D51
    # = TRIM(IF(COUNTIF(K$28:K28,K28)=1,K28&" ("&COUNTIF($K$28:$K$39,K28)&"/12); ","")&IF(COUNTIF(K$28:K29,K29)=1,K29&" ("&COUNTIF($K$28:$K$39,K29)&"/12); ","
    #   ")&IF(COUNTIF(K$28:K30,K30)=1,K30&" ("&COUNTIF($K$28:$K$39,K30)&"/12); ","")&IF(COUNTIF(K$28:K31,K31)=1,K31&" ("&COUNTIF($K$28:$K$39,K31)&"/12); ","")
    #   &IF(COUNTIF(K$28:K32,K32)=1,K32&" ("&COUNTIF($K$28:$K$39,K32)&"/12); ","")&IF(COUNTIF(K$28:K33,K33)=1,K33&" ("&COUNTIF($K$28:$K$39,K33)&"/12); ","")&I
    #   F(COUNTIF(K$28:K34,K34)=1,K34&" ("&COUNTIF($K$28:$K$39,K34)&"/12); ","")&IF(COUNTIF(K$28:K35,K35)=1,K35&" ("&COUNTIF($K$28:$K$39,K35)&"/12); ","")&IF(
    #   COUNTIF(K$28:K36,K36)=1,K36&" ("&COUNTIF($K$28:$K$39,K36)&"/12); ","")&IF(COUNTIF(K$28:K37,K37)=1,K37&" ("&COUNTIF($K$28:$K$39,K37)&"/12); ","")&IF(CO
    #   UNTIF(K$28:K38,K38)=1,K38&" ("&COUNTIF($K$28:$K$39,K38)&"/12); ","")&IF(COUNTIF(K$28:K39,K39)=1,K39&" ("&COUNTIF($K$28:$K$39,K39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-23,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-23,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-23,c+7), E.v(6,r-23,c+7)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-22,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-22,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-22,c+7), E.v(6,r-22,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-21,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-21,c+7), E.v(6,r-21,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-20,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-20,c+7), E.v(6,r-20,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-19,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-19,c+7), E.v(6,r-19,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-18,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-18,c+7), E.v(6,r-18,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-17,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-17,c+7), E.v(6,r-17,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-16,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-16,c+7), E.v(6,r-16,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-15,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-15,c+7), E.v(6,r-15,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-14,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-14,c+7), E.v(6,r-14,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-13,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-13,c+7), E.v(6,r-13,c+7)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-12,c+7),' ('),X.COUNTIF(E.rg(6,28,11,39,11), E.v(6,r-12,c+7))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+7,r-12,c+7), E.v(6,r-12,c+7)),1.0)) else '')))

def f72(E, r, c):
    # Ô C52
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(COUNTIF($L$28:$L$39,L28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-24,c+9)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f73(E, r, c):
    # Ô D52
    # = TRIM(IF(COUNTIF(L$28:L28,L28)=1,L28&" ("&COUNTIF($L$28:$L$39,L28)&"/12); ","")&IF(COUNTIF(L$28:L29,L29)=1,L29&" ("&COUNTIF($L$28:$L$39,L29)&"/12); ","
    #   ")&IF(COUNTIF(L$28:L30,L30)=1,L30&" ("&COUNTIF($L$28:$L$39,L30)&"/12); ","")&IF(COUNTIF(L$28:L31,L31)=1,L31&" ("&COUNTIF($L$28:$L$39,L31)&"/12); ","")
    #   &IF(COUNTIF(L$28:L32,L32)=1,L32&" ("&COUNTIF($L$28:$L$39,L32)&"/12); ","")&IF(COUNTIF(L$28:L33,L33)=1,L33&" ("&COUNTIF($L$28:$L$39,L33)&"/12); ","")&I
    #   F(COUNTIF(L$28:L34,L34)=1,L34&" ("&COUNTIF($L$28:$L$39,L34)&"/12); ","")&IF(COUNTIF(L$28:L35,L35)=1,L35&" ("&COUNTIF($L$28:$L$39,L35)&"/12); ","")&IF(
    #   COUNTIF(L$28:L36,L36)=1,L36&" ("&COUNTIF($L$28:$L$39,L36)&"/12); ","")&IF(COUNTIF(L$28:L37,L37)=1,L37&" ("&COUNTIF($L$28:$L$39,L37)&"/12); ","")&IF(CO
    #   UNTIF(L$28:L38,L38)=1,L38&" ("&COUNTIF($L$28:$L$39,L38)&"/12); ","")&IF(COUNTIF(L$28:L39,L39)=1,L39&" ("&COUNTIF($L$28:$L$39,L39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-24,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-24,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-24,c+8), E.v(6,r-24,c+8)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-23,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-23,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-23,c+8), E.v(6,r-23,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-22,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-22,c+8), E.v(6,r-22,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-21,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-21,c+8), E.v(6,r-21,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-20,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-20,c+8), E.v(6,r-20,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-19,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-19,c+8), E.v(6,r-19,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-18,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-18,c+8), E.v(6,r-18,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-17,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-17,c+8), E.v(6,r-17,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-16,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-16,c+8), E.v(6,r-16,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-15,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-15,c+8), E.v(6,r-15,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-14,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-14,c+8), E.v(6,r-14,c+8)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-13,c+8),' ('),X.COUNTIF(E.rg(6,28,12,39,12), E.v(6,r-13,c+8))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+8,r-13,c+8), E.v(6,r-13,c+8)),1.0)) else '')))

def f74(E, r, c):
    # Ô C53
    # = IF(COUNTA($M$28:$M$39)=0,"",IF(COUNTIF($M$28:$M$39,M28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,13,39,13)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-25,c+10)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f75(E, r, c):
    # Ô D53
    # = TRIM(IF(COUNTIF(M$28:M28,M28)=1,M28&" ("&COUNTIF($M$28:$M$39,M28)&"/12); ","")&IF(COUNTIF(M$28:M29,M29)=1,M29&" ("&COUNTIF($M$28:$M$39,M29)&"/12); ","
    #   ")&IF(COUNTIF(M$28:M30,M30)=1,M30&" ("&COUNTIF($M$28:$M$39,M30)&"/12); ","")&IF(COUNTIF(M$28:M31,M31)=1,M31&" ("&COUNTIF($M$28:$M$39,M31)&"/12); ","")
    #   &IF(COUNTIF(M$28:M32,M32)=1,M32&" ("&COUNTIF($M$28:$M$39,M32)&"/12); ","")&IF(COUNTIF(M$28:M33,M33)=1,M33&" ("&COUNTIF($M$28:$M$39,M33)&"/12); ","")&I
    #   F(COUNTIF(M$28:M34,M34)=1,M34&" ("&COUNTIF($M$28:$M$39,M34)&"/12); ","")&IF(COUNTIF(M$28:M35,M35)=1,M35&" ("&COUNTIF($M$28:$M$39,M35)&"/12); ","")&IF(
    #   COUNTIF(M$28:M36,M36)=1,M36&" ("&COUNTIF($M$28:$M$39,M36)&"/12); ","")&IF(COUNTIF(M$28:M37,M37)=1,M37&" ("&COUNTIF($M$28:$M$39,M37)&"/12); ","")&IF(CO
    #   UNTIF(M$28:M38,M38)=1,M38&" ("&COUNTIF($M$28:$M$39,M38)&"/12); ","")&IF(COUNTIF(M$28:M39,M39)=1,M39&" ("&COUNTIF($M$28:$M$39,M39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-25,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-25,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-25,c+9), E.v(6,r-25,c+9)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-24,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-24,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-24,c+9), E.v(6,r-24,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-23,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-23,c+9), E.v(6,r-23,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-22,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-22,c+9), E.v(6,r-22,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-21,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-21,c+9), E.v(6,r-21,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-20,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-20,c+9), E.v(6,r-20,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-19,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-19,c+9), E.v(6,r-19,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-18,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-18,c+9), E.v(6,r-18,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-17,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-17,c+9), E.v(6,r-17,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-16,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-16,c+9), E.v(6,r-16,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-15,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-15,c+9), E.v(6,r-15,c+9)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-14,c+9),' ('),X.COUNTIF(E.rg(6,28,13,39,13), E.v(6,r-14,c+9))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+9,r-14,c+9), E.v(6,r-14,c+9)),1.0)) else '')))

def f76(E, r, c):
    # Ô C54
    # = IF(COUNTA($N$28:$N$39)=0,"",IF(COUNTIF($N$28:$N$39,N28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,14,39,14)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-26,c+11)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f77(E, r, c):
    # Ô D54
    # = TRIM(IF(COUNTIF(N$28:N28,N28)=1,N28&" ("&COUNTIF($N$28:$N$39,N28)&"/12); ","")&IF(COUNTIF(N$28:N29,N29)=1,N29&" ("&COUNTIF($N$28:$N$39,N29)&"/12); ","
    #   ")&IF(COUNTIF(N$28:N30,N30)=1,N30&" ("&COUNTIF($N$28:$N$39,N30)&"/12); ","")&IF(COUNTIF(N$28:N31,N31)=1,N31&" ("&COUNTIF($N$28:$N$39,N31)&"/12); ","")
    #   &IF(COUNTIF(N$28:N32,N32)=1,N32&" ("&COUNTIF($N$28:$N$39,N32)&"/12); ","")&IF(COUNTIF(N$28:N33,N33)=1,N33&" ("&COUNTIF($N$28:$N$39,N33)&"/12); ","")&I
    #   F(COUNTIF(N$28:N34,N34)=1,N34&" ("&COUNTIF($N$28:$N$39,N34)&"/12); ","")&IF(COUNTIF(N$28:N35,N35)=1,N35&" ("&COUNTIF($N$28:$N$39,N35)&"/12); ","")&IF(
    #   COUNTIF(N$28:N36,N36)=1,N36&" ("&COUNTIF($N$28:$N$39,N36)&"/12); ","")&IF(COUNTIF(N$28:N37,N37)=1,N37&" ("&COUNTIF($N$28:$N$39,N37)&"/12); ","")&IF(CO
    #   UNTIF(N$28:N38,N38)=1,N38&" ("&COUNTIF($N$28:$N$39,N38)&"/12); ","")&IF(COUNTIF(N$28:N39,N39)=1,N39&" ("&COUNTIF($N$28:$N$39,N39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-26,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-26,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-26,c+10), E.v(6,r-26,c+10)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-25,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-25,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-25,c+10), E.v(6,r-25,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-24,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-24,c+10), E.v(6,r-24,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-23,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-23,c+10), E.v(6,r-23,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-22,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-22,c+10), E.v(6,r-22,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-21,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-21,c+10), E.v(6,r-21,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-20,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-20,c+10), E.v(6,r-20,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-19,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-19,c+10), E.v(6,r-19,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-18,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-18,c+10), E.v(6,r-18,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-17,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-17,c+10), E.v(6,r-17,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-16,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-16,c+10), E.v(6,r-16,c+10)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-15,c+10),' ('),X.COUNTIF(E.rg(6,28,14,39,14), E.v(6,r-15,c+10))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+10,r-15,c+10), E.v(6,r-15,c+10)),1.0)) else '')))

def f78(E, r, c):
    # Ô C55
    # = IF(COUNTA($O$28:$O$39)=0,"",IF(COUNTIF($O$28:$O$39,O28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,15,39,15)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-27,c+12)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f79(E, r, c):
    # Ô D55
    # = TRIM(IF(COUNTIF(O$28:O28,O28)=1,O28&" ("&COUNTIF($O$28:$O$39,O28)&"/12); ","")&IF(COUNTIF(O$28:O29,O29)=1,O29&" ("&COUNTIF($O$28:$O$39,O29)&"/12); ","
    #   ")&IF(COUNTIF(O$28:O30,O30)=1,O30&" ("&COUNTIF($O$28:$O$39,O30)&"/12); ","")&IF(COUNTIF(O$28:O31,O31)=1,O31&" ("&COUNTIF($O$28:$O$39,O31)&"/12); ","")
    #   &IF(COUNTIF(O$28:O32,O32)=1,O32&" ("&COUNTIF($O$28:$O$39,O32)&"/12); ","")&IF(COUNTIF(O$28:O33,O33)=1,O33&" ("&COUNTIF($O$28:$O$39,O33)&"/12); ","")&I
    #   F(COUNTIF(O$28:O34,O34)=1,O34&" ("&COUNTIF($O$28:$O$39,O34)&"/12); ","")&IF(COUNTIF(O$28:O35,O35)=1,O35&" ("&COUNTIF($O$28:$O$39,O35)&"/12); ","")&IF(
    #   COUNTIF(O$28:O36,O36)=1,O36&" ("&COUNTIF($O$28:$O$39,O36)&"/12); ","")&IF(COUNTIF(O$28:O37,O37)=1,O37&" ("&COUNTIF($O$28:$O$39,O37)&"/12); ","")&IF(CO
    #   UNTIF(O$28:O38,O38)=1,O38&" ("&COUNTIF($O$28:$O$39,O38)&"/12); ","")&IF(COUNTIF(O$28:O39,O39)=1,O39&" ("&COUNTIF($O$28:$O$39,O39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-27,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-27,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-27,c+11), E.v(6,r-27,c+11)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-26,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-26,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-26,c+11), E.v(6,r-26,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-25,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-25,c+11), E.v(6,r-25,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-24,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-24,c+11), E.v(6,r-24,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-23,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-23,c+11), E.v(6,r-23,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-22,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-22,c+11), E.v(6,r-22,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-21,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-21,c+11), E.v(6,r-21,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-20,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-20,c+11), E.v(6,r-20,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-19,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-19,c+11), E.v(6,r-19,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-18,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-18,c+11), E.v(6,r-18,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-17,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-17,c+11), E.v(6,r-17,c+11)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-16,c+11),' ('),X.COUNTIF(E.rg(6,28,15,39,15), E.v(6,r-16,c+11))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+11,r-16,c+11), E.v(6,r-16,c+11)),1.0)) else '')))

def f80(E, r, c):
    # Ô C56
    # = IF(COUNTA($P$28:$P$39)=0,"",IF(COUNTIF($P$28:$P$39,P28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,16,39,16)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-28,c+13)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f81(E, r, c):
    # Ô D56
    # = TRIM(IF(COUNTIF(P$28:P28,P28)=1,P28&" ("&COUNTIF($P$28:$P$39,P28)&"/12); ","")&IF(COUNTIF(P$28:P29,P29)=1,P29&" ("&COUNTIF($P$28:$P$39,P29)&"/12); ","
    #   ")&IF(COUNTIF(P$28:P30,P30)=1,P30&" ("&COUNTIF($P$28:$P$39,P30)&"/12); ","")&IF(COUNTIF(P$28:P31,P31)=1,P31&" ("&COUNTIF($P$28:$P$39,P31)&"/12); ","")
    #   &IF(COUNTIF(P$28:P32,P32)=1,P32&" ("&COUNTIF($P$28:$P$39,P32)&"/12); ","")&IF(COUNTIF(P$28:P33,P33)=1,P33&" ("&COUNTIF($P$28:$P$39,P33)&"/12); ","")&I
    #   F(COUNTIF(P$28:P34,P34)=1,P34&" ("&COUNTIF($P$28:$P$39,P34)&"/12); ","")&IF(COUNTIF(P$28:P35,P35)=1,P35&" ("&COUNTIF($P$28:$P$39,P35)&"/12); ","")&IF(
    #   COUNTIF(P$28:P36,P36)=1,P36&" ("&COUNTIF($P$28:$P$39,P36)&"/12); ","")&IF(COUNTIF(P$28:P37,P37)=1,P37&" ("&COUNTIF($P$28:$P$39,P37)&"/12); ","")&IF(CO
    #   UNTIF(P$28:P38,P38)=1,P38&" ("&COUNTIF($P$28:$P$39,P38)&"/12); ","")&IF(COUNTIF(P$28:P39,P39)=1,P39&" ("&COUNTIF($P$28:$P$39,P39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-28,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-28,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-28,c+12), E.v(6,r-28,c+12)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-27,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-27,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-27,c+12), E.v(6,r-27,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-26,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-26,c+12), E.v(6,r-26,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-25,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-25,c+12), E.v(6,r-25,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-24,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-24,c+12), E.v(6,r-24,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-23,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-23,c+12), E.v(6,r-23,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-22,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-22,c+12), E.v(6,r-22,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-21,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-21,c+12), E.v(6,r-21,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-20,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-20,c+12), E.v(6,r-20,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-19,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-19,c+12), E.v(6,r-19,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-18,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-18,c+12), E.v(6,r-18,c+12)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-17,c+12),' ('),X.COUNTIF(E.rg(6,28,16,39,16), E.v(6,r-17,c+12))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+12,r-17,c+12), E.v(6,r-17,c+12)),1.0)) else '')))

def f82(E, r, c):
    # Ô C57
    # = IF(COUNTA($Q$28:$Q$39)=0,"",IF(COUNTIF($Q$28:$Q$39,Q28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,17,39,17)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-29,c+14)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f83(E, r, c):
    # Ô D57
    # = TRIM(IF(COUNTIF(Q$28:Q28,Q28)=1,Q28&" ("&COUNTIF($Q$28:$Q$39,Q28)&"/12); ","")&IF(COUNTIF(Q$28:Q29,Q29)=1,Q29&" ("&COUNTIF($Q$28:$Q$39,Q29)&"/12); ","
    #   ")&IF(COUNTIF(Q$28:Q30,Q30)=1,Q30&" ("&COUNTIF($Q$28:$Q$39,Q30)&"/12); ","")&IF(COUNTIF(Q$28:Q31,Q31)=1,Q31&" ("&COUNTIF($Q$28:$Q$39,Q31)&"/12); ","")
    #   &IF(COUNTIF(Q$28:Q32,Q32)=1,Q32&" ("&COUNTIF($Q$28:$Q$39,Q32)&"/12); ","")&IF(COUNTIF(Q$28:Q33,Q33)=1,Q33&" ("&COUNTIF($Q$28:$Q$39,Q33)&"/12); ","")&I
    #   F(COUNTIF(Q$28:Q34,Q34)=1,Q34&" ("&COUNTIF($Q$28:$Q$39,Q34)&"/12); ","")&IF(COUNTIF(Q$28:Q35,Q35)=1,Q35&" ("&COUNTIF($Q$28:$Q$39,Q35)&"/12); ","")&IF(
    #   COUNTIF(Q$28:Q36,Q36)=1,Q36&" ("&COUNTIF($Q$28:$Q$39,Q36)&"/12); ","")&IF(COUNTIF(Q$28:Q37,Q37)=1,Q37&" ("&COUNTIF($Q$28:$Q$39,Q37)&"/12); ","")&IF(CO
    #   UNTIF(Q$28:Q38,Q38)=1,Q38&" ("&COUNTIF($Q$28:$Q$39,Q38)&"/12); ","")&IF(COUNTIF(Q$28:Q39,Q39)=1,Q39&" ("&COUNTIF($Q$28:$Q$39,Q39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-29,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-29,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-29,c+13), E.v(6,r-29,c+13)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-28,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-28,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-28,c+13), E.v(6,r-28,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-27,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-27,c+13), E.v(6,r-27,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-26,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-26,c+13), E.v(6,r-26,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-25,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-25,c+13), E.v(6,r-25,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-24,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-24,c+13), E.v(6,r-24,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-23,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-23,c+13), E.v(6,r-23,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-22,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-22,c+13), E.v(6,r-22,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-21,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-21,c+13), E.v(6,r-21,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-20,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-20,c+13), E.v(6,r-20,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-19,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-19,c+13), E.v(6,r-19,c+13)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-18,c+13),' ('),X.COUNTIF(E.rg(6,28,17,39,17), E.v(6,r-18,c+13))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+13,r-18,c+13), E.v(6,r-18,c+13)),1.0)) else '')))

def f84(E, r, c):
    # Ô C58
    # = IF(COUNTA($R$28:$R$39)=0,"",IF(COUNTIF($R$28:$R$39,R28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,18,39,18)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-30,c+15)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f85(E, r, c):
    # Ô D58
    # = TRIM(IF(COUNTIF(R$28:R28,R28)=1,R28&" ("&COUNTIF($R$28:$R$39,R28)&"/12); ","")&IF(COUNTIF(R$28:R29,R29)=1,R29&" ("&COUNTIF($R$28:$R$39,R29)&"/12); ","
    #   ")&IF(COUNTIF(R$28:R30,R30)=1,R30&" ("&COUNTIF($R$28:$R$39,R30)&"/12); ","")&IF(COUNTIF(R$28:R31,R31)=1,R31&" ("&COUNTIF($R$28:$R$39,R31)&"/12); ","")
    #   &IF(COUNTIF(R$28:R32,R32)=1,R32&" ("&COUNTIF($R$28:$R$39,R32)&"/12); ","")&IF(COUNTIF(R$28:R33,R33)=1,R33&" ("&COUNTIF($R$28:$R$39,R33)&"/12); ","")&I
    #   F(COUNTIF(R$28:R34,R34)=1,R34&" ("&COUNTIF($R$28:$R$39,R34)&"/12); ","")&IF(COUNTIF(R$28:R35,R35)=1,R35&" ("&COUNTIF($R$28:$R$39,R35)&"/12); ","")&IF(
    #   COUNTIF(R$28:R36,R36)=1,R36&" ("&COUNTIF($R$28:$R$39,R36)&"/12); ","")&IF(COUNTIF(R$28:R37,R37)=1,R37&" ("&COUNTIF($R$28:$R$39,R37)&"/12); ","")&IF(CO
    #   UNTIF(R$28:R38,R38)=1,R38&" ("&COUNTIF($R$28:$R$39,R38)&"/12); ","")&IF(COUNTIF(R$28:R39,R39)=1,R39&" ("&COUNTIF($R$28:$R$39,R39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-30,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-30,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-30,c+14), E.v(6,r-30,c+14)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-29,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-29,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-29,c+14), E.v(6,r-29,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-28,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-28,c+14), E.v(6,r-28,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-27,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-27,c+14), E.v(6,r-27,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-26,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-26,c+14), E.v(6,r-26,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-25,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-25,c+14), E.v(6,r-25,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-24,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-24,c+14), E.v(6,r-24,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-23,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-23,c+14), E.v(6,r-23,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-22,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-22,c+14), E.v(6,r-22,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-21,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-21,c+14), E.v(6,r-21,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-20,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-20,c+14), E.v(6,r-20,c+14)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-19,c+14),' ('),X.COUNTIF(E.rg(6,28,18,39,18), E.v(6,r-19,c+14))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+14,r-19,c+14), E.v(6,r-19,c+14)),1.0)) else '')))

def f86(E, r, c):
    # Ô C59
    # = IF(COUNTA($S$28:$S$39)=0,"",IF(COUNTIF($S$28:$S$39,S28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,19,39,19)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-31,c+16)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f87(E, r, c):
    # Ô D59
    # = TRIM(IF(COUNTIF(S$28:S28,S28)=1,S28&" ("&COUNTIF($S$28:$S$39,S28)&"/12); ","")&IF(COUNTIF(S$28:S29,S29)=1,S29&" ("&COUNTIF($S$28:$S$39,S29)&"/12); ","
    #   ")&IF(COUNTIF(S$28:S30,S30)=1,S30&" ("&COUNTIF($S$28:$S$39,S30)&"/12); ","")&IF(COUNTIF(S$28:S31,S31)=1,S31&" ("&COUNTIF($S$28:$S$39,S31)&"/12); ","")
    #   &IF(COUNTIF(S$28:S32,S32)=1,S32&" ("&COUNTIF($S$28:$S$39,S32)&"/12); ","")&IF(COUNTIF(S$28:S33,S33)=1,S33&" ("&COUNTIF($S$28:$S$39,S33)&"/12); ","")&I
    #   F(COUNTIF(S$28:S34,S34)=1,S34&" ("&COUNTIF($S$28:$S$39,S34)&"/12); ","")&IF(COUNTIF(S$28:S35,S35)=1,S35&" ("&COUNTIF($S$28:$S$39,S35)&"/12); ","")&IF(
    #   COUNTIF(S$28:S36,S36)=1,S36&" ("&COUNTIF($S$28:$S$39,S36)&"/12); ","")&IF(COUNTIF(S$28:S37,S37)=1,S37&" ("&COUNTIF($S$28:$S$39,S37)&"/12); ","")&IF(CO
    #   UNTIF(S$28:S38,S38)=1,S38&" ("&COUNTIF($S$28:$S$39,S38)&"/12); ","")&IF(COUNTIF(S$28:S39,S39)=1,S39&" ("&COUNTIF($S$28:$S$39,S39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-31,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-31,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-31,c+15), E.v(6,r-31,c+15)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-30,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-30,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-30,c+15), E.v(6,r-30,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-29,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-29,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-29,c+15), E.v(6,r-29,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-28,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-28,c+15), E.v(6,r-28,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-27,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-27,c+15), E.v(6,r-27,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-26,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-26,c+15), E.v(6,r-26,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-25,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-25,c+15), E.v(6,r-25,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-24,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-24,c+15), E.v(6,r-24,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-23,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-23,c+15), E.v(6,r-23,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-22,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-22,c+15), E.v(6,r-22,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-21,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-21,c+15), E.v(6,r-21,c+15)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-20,c+15),' ('),X.COUNTIF(E.rg(6,28,19,39,19), E.v(6,r-20,c+15))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+15,r-20,c+15), E.v(6,r-20,c+15)),1.0)) else '')))

def f88(E, r, c):
    # Ô C60
    # = IF(COUNTA($T$28:$T$39)=0,"",IF(COUNTIF($T$28:$T$39,T28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,20,39,20)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-32,c+17)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f89(E, r, c):
    # Ô D60
    # = TRIM(IF(COUNTIF(T$28:T28,T28)=1,T28&" ("&COUNTIF($T$28:$T$39,T28)&"/12); ","")&IF(COUNTIF(T$28:T29,T29)=1,T29&" ("&COUNTIF($T$28:$T$39,T29)&"/12); ","
    #   ")&IF(COUNTIF(T$28:T30,T30)=1,T30&" ("&COUNTIF($T$28:$T$39,T30)&"/12); ","")&IF(COUNTIF(T$28:T31,T31)=1,T31&" ("&COUNTIF($T$28:$T$39,T31)&"/12); ","")
    #   &IF(COUNTIF(T$28:T32,T32)=1,T32&" ("&COUNTIF($T$28:$T$39,T32)&"/12); ","")&IF(COUNTIF(T$28:T33,T33)=1,T33&" ("&COUNTIF($T$28:$T$39,T33)&"/12); ","")&I
    #   F(COUNTIF(T$28:T34,T34)=1,T34&" ("&COUNTIF($T$28:$T$39,T34)&"/12); ","")&IF(COUNTIF(T$28:T35,T35)=1,T35&" ("&COUNTIF($T$28:$T$39,T35)&"/12); ","")&IF(
    #   COUNTIF(T$28:T36,T36)=1,T36&" ("&COUNTIF($T$28:$T$39,T36)&"/12); ","")&IF(COUNTIF(T$28:T37,T37)=1,T37&" ("&COUNTIF($T$28:$T$39,T37)&"/12); ","")&IF(CO
    #   UNTIF(T$28:T38,T38)=1,T38&" ("&COUNTIF($T$28:$T$39,T38)&"/12); ","")&IF(COUNTIF(T$28:T39,T39)=1,T39&" ("&COUNTIF($T$28:$T$39,T39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-32,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-32,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-32,c+16), E.v(6,r-32,c+16)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-31,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-31,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-31,c+16), E.v(6,r-31,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-30,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-30,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-30,c+16), E.v(6,r-30,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-29,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-29,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-29,c+16), E.v(6,r-29,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-28,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-28,c+16), E.v(6,r-28,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-27,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-27,c+16), E.v(6,r-27,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-26,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-26,c+16), E.v(6,r-26,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-25,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-25,c+16), E.v(6,r-25,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-24,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-24,c+16), E.v(6,r-24,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-23,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-23,c+16), E.v(6,r-23,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-22,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-22,c+16), E.v(6,r-22,c+16)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-21,c+16),' ('),X.COUNTIF(E.rg(6,28,20,39,20), E.v(6,r-21,c+16))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+16,r-21,c+16), E.v(6,r-21,c+16)),1.0)) else '')))

def f90(E, r, c):
    # Ô C61
    # = IF(COUNTA($U$28:$U$39)=0,"",IF(COUNTIF($U$28:$U$39,U28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,21,39,21)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-33,c+18)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f91(E, r, c):
    # Ô D61
    # = TRIM(IF(COUNTIF(U$28:U28,U28)=1,U28&" ("&COUNTIF($U$28:$U$39,U28)&"/12); ","")&IF(COUNTIF(U$28:U29,U29)=1,U29&" ("&COUNTIF($U$28:$U$39,U29)&"/12); ","
    #   ")&IF(COUNTIF(U$28:U30,U30)=1,U30&" ("&COUNTIF($U$28:$U$39,U30)&"/12); ","")&IF(COUNTIF(U$28:U31,U31)=1,U31&" ("&COUNTIF($U$28:$U$39,U31)&"/12); ","")
    #   &IF(COUNTIF(U$28:U32,U32)=1,U32&" ("&COUNTIF($U$28:$U$39,U32)&"/12); ","")&IF(COUNTIF(U$28:U33,U33)=1,U33&" ("&COUNTIF($U$28:$U$39,U33)&"/12); ","")&I
    #   F(COUNTIF(U$28:U34,U34)=1,U34&" ("&COUNTIF($U$28:$U$39,U34)&"/12); ","")&IF(COUNTIF(U$28:U35,U35)=1,U35&" ("&COUNTIF($U$28:$U$39,U35)&"/12); ","")&IF(
    #   COUNTIF(U$28:U36,U36)=1,U36&" ("&COUNTIF($U$28:$U$39,U36)&"/12); ","")&IF(COUNTIF(U$28:U37,U37)=1,U37&" ("&COUNTIF($U$28:$U$39,U37)&"/12); ","")&IF(CO
    #   UNTIF(U$28:U38,U38)=1,U38&" ("&COUNTIF($U$28:$U$39,U38)&"/12); ","")&IF(COUNTIF(U$28:U39,U39)=1,U39&" ("&COUNTIF($U$28:$U$39,U39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-33,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-33,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-33,c+17), E.v(6,r-33,c+17)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-32,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-32,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-32,c+17), E.v(6,r-32,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-31,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-31,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-31,c+17), E.v(6,r-31,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-30,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-30,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-30,c+17), E.v(6,r-30,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-29,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-29,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-29,c+17), E.v(6,r-29,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-28,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-28,c+17), E.v(6,r-28,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-27,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-27,c+17), E.v(6,r-27,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-26,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-26,c+17), E.v(6,r-26,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-25,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-25,c+17), E.v(6,r-25,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-24,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-24,c+17), E.v(6,r-24,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-23,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-23,c+17), E.v(6,r-23,c+17)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-22,c+17),' ('),X.COUNTIF(E.rg(6,28,21,39,21), E.v(6,r-22,c+17))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+17,r-22,c+17), E.v(6,r-22,c+17)),1.0)) else '')))

def f92(E, r, c):
    # Ô C62
    # = IF(COUNTA($V$28:$V$39)=0,"",IF(COUNTIF($V$28:$V$39,V28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,22,39,22)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-34,c+19)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f93(E, r, c):
    # Ô D62
    # = TRIM(IF(COUNTIF(V$28:V28,V28)=1,V28&" ("&COUNTIF($V$28:$V$39,V28)&"/12); ","")&IF(COUNTIF(V$28:V29,V29)=1,V29&" ("&COUNTIF($V$28:$V$39,V29)&"/12); ","
    #   ")&IF(COUNTIF(V$28:V30,V30)=1,V30&" ("&COUNTIF($V$28:$V$39,V30)&"/12); ","")&IF(COUNTIF(V$28:V31,V31)=1,V31&" ("&COUNTIF($V$28:$V$39,V31)&"/12); ","")
    #   &IF(COUNTIF(V$28:V32,V32)=1,V32&" ("&COUNTIF($V$28:$V$39,V32)&"/12); ","")&IF(COUNTIF(V$28:V33,V33)=1,V33&" ("&COUNTIF($V$28:$V$39,V33)&"/12); ","")&I
    #   F(COUNTIF(V$28:V34,V34)=1,V34&" ("&COUNTIF($V$28:$V$39,V34)&"/12); ","")&IF(COUNTIF(V$28:V35,V35)=1,V35&" ("&COUNTIF($V$28:$V$39,V35)&"/12); ","")&IF(
    #   COUNTIF(V$28:V36,V36)=1,V36&" ("&COUNTIF($V$28:$V$39,V36)&"/12); ","")&IF(COUNTIF(V$28:V37,V37)=1,V37&" ("&COUNTIF($V$28:$V$39,V37)&"/12); ","")&IF(CO
    #   UNTIF(V$28:V38,V38)=1,V38&" ("&COUNTIF($V$28:$V$39,V38)&"/12); ","")&IF(COUNTIF(V$28:V39,V39)=1,V39&" ("&COUNTIF($V$28:$V$39,V39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-34,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-34,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-34,c+18), E.v(6,r-34,c+18)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-33,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-33,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-33,c+18), E.v(6,r-33,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-32,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-32,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-32,c+18), E.v(6,r-32,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-31,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-31,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-31,c+18), E.v(6,r-31,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-30,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-30,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-30,c+18), E.v(6,r-30,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-29,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-29,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-29,c+18), E.v(6,r-29,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-28,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-28,c+18), E.v(6,r-28,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-27,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-27,c+18), E.v(6,r-27,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-26,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-26,c+18), E.v(6,r-26,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-25,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-25,c+18), E.v(6,r-25,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-24,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-24,c+18), E.v(6,r-24,c+18)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-23,c+18),' ('),X.COUNTIF(E.rg(6,28,22,39,22), E.v(6,r-23,c+18))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+18,r-23,c+18), E.v(6,r-23,c+18)),1.0)) else '')))

def f94(E, r, c):
    # Ô C63
    # = IF(COUNTA($W$28:$W$39)=0,"",IF(COUNTIF($W$28:$W$39,W28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,23,39,23)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-35,c+20)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f95(E, r, c):
    # Ô D63
    # = TRIM(IF(COUNTIF(W$28:W28,W28)=1,W28&" ("&COUNTIF($W$28:$W$39,W28)&"/12); ","")&IF(COUNTIF(W$28:W29,W29)=1,W29&" ("&COUNTIF($W$28:$W$39,W29)&"/12); ","
    #   ")&IF(COUNTIF(W$28:W30,W30)=1,W30&" ("&COUNTIF($W$28:$W$39,W30)&"/12); ","")&IF(COUNTIF(W$28:W31,W31)=1,W31&" ("&COUNTIF($W$28:$W$39,W31)&"/12); ","")
    #   &IF(COUNTIF(W$28:W32,W32)=1,W32&" ("&COUNTIF($W$28:$W$39,W32)&"/12); ","")&IF(COUNTIF(W$28:W33,W33)=1,W33&" ("&COUNTIF($W$28:$W$39,W33)&"/12); ","")&I
    #   F(COUNTIF(W$28:W34,W34)=1,W34&" ("&COUNTIF($W$28:$W$39,W34)&"/12); ","")&IF(COUNTIF(W$28:W35,W35)=1,W35&" ("&COUNTIF($W$28:$W$39,W35)&"/12); ","")&IF(
    #   COUNTIF(W$28:W36,W36)=1,W36&" ("&COUNTIF($W$28:$W$39,W36)&"/12); ","")&IF(COUNTIF(W$28:W37,W37)=1,W37&" ("&COUNTIF($W$28:$W$39,W37)&"/12); ","")&IF(CO
    #   UNTIF(W$28:W38,W38)=1,W38&" ("&COUNTIF($W$28:$W$39,W38)&"/12); ","")&IF(COUNTIF(W$28:W39,W39)=1,W39&" ("&COUNTIF($W$28:$W$39,W39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-35,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-35,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-35,c+19), E.v(6,r-35,c+19)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-34,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-34,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-34,c+19), E.v(6,r-34,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-33,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-33,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-33,c+19), E.v(6,r-33,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-32,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-32,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-32,c+19), E.v(6,r-32,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-31,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-31,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-31,c+19), E.v(6,r-31,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-30,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-30,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-30,c+19), E.v(6,r-30,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-29,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-29,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-29,c+19), E.v(6,r-29,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-28,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-28,c+19), E.v(6,r-28,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-27,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-27,c+19), E.v(6,r-27,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-26,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-26,c+19), E.v(6,r-26,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-25,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-25,c+19), E.v(6,r-25,c+19)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-24,c+19),' ('),X.COUNTIF(E.rg(6,28,23,39,23), E.v(6,r-24,c+19))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+19,r-24,c+19), E.v(6,r-24,c+19)),1.0)) else '')))

def f96(E, r, c):
    # Ô C64
    # = IF(COUNTA($X$28:$X$39)=0,"",IF(COUNTIF($X$28:$X$39,X28)=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,24,39,24)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-36,c+21)),12.0)) else 'PHỤ THUỘC GIỜ'))

def f97(E, r, c):
    # Ô D64
    # = TRIM(IF(COUNTIF(X$28:X28,X28)=1,X28&" ("&COUNTIF($X$28:$X$39,X28)&"/12); ","")&IF(COUNTIF(X$28:X29,X29)=1,X29&" ("&COUNTIF($X$28:$X$39,X29)&"/12); ","
    #   ")&IF(COUNTIF(X$28:X30,X30)=1,X30&" ("&COUNTIF($X$28:$X$39,X30)&"/12); ","")&IF(COUNTIF(X$28:X31,X31)=1,X31&" ("&COUNTIF($X$28:$X$39,X31)&"/12); ","")
    #   &IF(COUNTIF(X$28:X32,X32)=1,X32&" ("&COUNTIF($X$28:$X$39,X32)&"/12); ","")&IF(COUNTIF(X$28:X33,X33)=1,X33&" ("&COUNTIF($X$28:$X$39,X33)&"/12); ","")&I
    #   F(COUNTIF(X$28:X34,X34)=1,X34&" ("&COUNTIF($X$28:$X$39,X34)&"/12); ","")&IF(COUNTIF(X$28:X35,X35)=1,X35&" ("&COUNTIF($X$28:$X$39,X35)&"/12); ","")&IF(
    #   COUNTIF(X$28:X36,X36)=1,X36&" ("&COUNTIF($X$28:$X$39,X36)&"/12); ","")&IF(COUNTIF(X$28:X37,X37)=1,X37&" ("&COUNTIF($X$28:$X$39,X37)&"/12); ","")&IF(CO
    #   UNTIF(X$28:X38,X38)=1,X38&" ("&COUNTIF($X$28:$X$39,X38)&"/12); ","")&IF(COUNTIF(X$28:X39,X39)=1,X39&" ("&COUNTIF($X$28:$X$39,X39)&"/12); ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(6,r-36,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-36,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-36,c+20), E.v(6,r-36,c+20)),1.0)) else ''),(X.cat(X.cat(X.cat(E.v(6,r-35,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-35,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-35,c+20), E.v(6,r-35,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-34,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-34,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-34,c+20), E.v(6,r-34,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-33,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-33,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-33,c+20), E.v(6,r-33,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-32,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-32,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-32,c+20), E.v(6,r-32,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-31,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-31,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-31,c+20), E.v(6,r-31,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-30,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-30,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-30,c+20), E.v(6,r-30,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-29,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-29,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-29,c+20), E.v(6,r-29,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-28,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-28,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-28,c+20), E.v(6,r-28,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-27,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-27,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-27,c+20), E.v(6,r-27,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-26,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-26,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-26,c+20), E.v(6,r-26,c+20)),1.0)) else '')),(X.cat(X.cat(X.cat(E.v(6,r-25,c+20),' ('),X.COUNTIF(E.rg(6,28,24,39,24), E.v(6,r-25,c+20))),'/12); ') if X.truthy(X.eq(X.COUNTIF(E.rg(6,28,c+20,r-25,c+20), E.v(6,r-25,c+20)),1.0)) else '')))

def f98(E, r, c):
    # Ô B65
    # = "Đại vận "&DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-30,c+1)),' ('),E.v(14,r-29,c+1)),'–'),X.add(E.v(14,r-29,c+1),9.0)),' tuổi)')

def f99(E, r, c):
    # Ô C65
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,1,1)=MID(L28,1,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 1.0, 1.0),X.AF_MID(E.v(6,r-37,c+9), 1.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f100(E, r, c):
    # Ô D65
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C65="ỔN ĐỊNH",IF(MID(L28,1,1)="T","Tốt",IF(MID(L28,1,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,1,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,1,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,1,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-37,c+8), 1.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-37,c+8), 1.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 1.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 1.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 1.0, 1.0),'X'))))),'/12')))

def f101(E, r, c):
    # Ô B66
    # = "Đại vận "&DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-31,c+2)),' ('),E.v(14,r-30,c+2)),'–'),X.add(E.v(14,r-30,c+2),9.0)),' tuổi)')

def f102(E, r, c):
    # Ô C66
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,2,1)=MID(L28,2,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 2.0, 1.0),X.AF_MID(E.v(6,r-38,c+9), 2.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f103(E, r, c):
    # Ô D66
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C66="ỔN ĐỊNH",IF(MID(L28,2,1)="T","Tốt",IF(MID(L28,2,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,2,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,2,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,2,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-38,c+8), 2.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-38,c+8), 2.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 2.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 2.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 2.0, 1.0),'X'))))),'/12')))

def f104(E, r, c):
    # Ô B67
    # = "Đại vận "&DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-32,c+3)),' ('),E.v(14,r-31,c+3)),'–'),X.add(E.v(14,r-31,c+3),9.0)),' tuổi)')

def f105(E, r, c):
    # Ô C67
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,3,1)=MID(L28,3,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 3.0, 1.0),X.AF_MID(E.v(6,r-39,c+9), 3.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f106(E, r, c):
    # Ô D67
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C67="ỔN ĐỊNH",IF(MID(L28,3,1)="T","Tốt",IF(MID(L28,3,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,3,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,3,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,3,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-39,c+8), 3.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-39,c+8), 3.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 3.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 3.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 3.0, 1.0),'X'))))),'/12')))

def f107(E, r, c):
    # Ô B68
    # = "Đại vận "&DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-33,c+4)),' ('),E.v(14,r-32,c+4)),'–'),X.add(E.v(14,r-32,c+4),9.0)),' tuổi)')

def f108(E, r, c):
    # Ô C68
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,4,1)=MID(L28,4,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 4.0, 1.0),X.AF_MID(E.v(6,r-40,c+9), 4.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f109(E, r, c):
    # Ô D68
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C68="ỔN ĐỊNH",IF(MID(L28,4,1)="T","Tốt",IF(MID(L28,4,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,4,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,4,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,4,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-40,c+8), 4.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-40,c+8), 4.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 4.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 4.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 4.0, 1.0),'X'))))),'/12')))

def f110(E, r, c):
    # Ô B69
    # = "Đại vận "&DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-34,c+5)),' ('),E.v(14,r-33,c+5)),'–'),X.add(E.v(14,r-33,c+5),9.0)),' tuổi)')

def f111(E, r, c):
    # Ô C69
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,5,1)=MID(L28,5,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 5.0, 1.0),X.AF_MID(E.v(6,r-41,c+9), 5.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f112(E, r, c):
    # Ô D69
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C69="ỔN ĐỊNH",IF(MID(L28,5,1)="T","Tốt",IF(MID(L28,5,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,5,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,5,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,5,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-41,c+8), 5.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-41,c+8), 5.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 5.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 5.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 5.0, 1.0),'X'))))),'/12')))

def f113(E, r, c):
    # Ô B70
    # = "Đại vận "&DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-35,c+6)),' ('),E.v(14,r-34,c+6)),'–'),X.add(E.v(14,r-34,c+6),9.0)),' tuổi)')

def f114(E, r, c):
    # Ô C70
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,6,1)=MID(L28,6,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 6.0, 1.0),X.AF_MID(E.v(6,r-42,c+9), 6.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f115(E, r, c):
    # Ô D70
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C70="ỔN ĐỊNH",IF(MID(L28,6,1)="T","Tốt",IF(MID(L28,6,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,6,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,6,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,6,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-42,c+8), 6.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-42,c+8), 6.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 6.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 6.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 6.0, 1.0),'X'))))),'/12')))

def f116(E, r, c):
    # Ô B71
    # = "Đại vận "&DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-36,c+7)),' ('),E.v(14,r-35,c+7)),'–'),X.add(E.v(14,r-35,c+7),9.0)),' tuổi)')

def f117(E, r, c):
    # Ô C71
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,7,1)=MID(L28,7,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 7.0, 1.0),X.AF_MID(E.v(6,r-43,c+9), 7.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f118(E, r, c):
    # Ô D71
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C71="ỔN ĐỊNH",IF(MID(L28,7,1)="T","Tốt",IF(MID(L28,7,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,7,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,7,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,7,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-43,c+8), 7.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-43,c+8), 7.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 7.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 7.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 7.0, 1.0),'X'))))),'/12')))

def f119(E, r, c):
    # Ô B72
    # = "Đại vận "&DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận ',E.v(14,r-37,c+8)),' ('),E.v(14,r-36,c+8)),'–'),X.add(E.v(14,r-36,c+8),9.0)),' tuổi)')

def f120(E, r, c):
    # Ô C72
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(SUMPRODUCT(--(MID($L$28:$L$39,8,1)=MID(L28,8,1)))=12,"ỔN ĐỊNH","PHỤ THUỘC GIỜ"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else ('ỔN ĐỊNH' if X.truthy(X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 8.0, 1.0),X.AF_MID(E.v(6,r-44,c+9), 8.0, 1.0))))),12.0)) else 'PHỤ THUỘC GIỜ'))

def f121(E, r, c):
    # Ô D72
    # = IF(COUNTA($L$28:$L$39)=0,"",IF(C72="ỔN ĐỊNH",IF(MID(L28,8,1)="T","Tốt",IF(MID(L28,8,1)="X","Xấu","Thường"))&" (12/12)","Tốt "&SUMPRODUCT(--(MID($L$28:
    #   $L$39,8,1)="T"))&"/12; Thường "&SUMPRODUCT(--(MID($L$28:$L$39,8,1)="·"))&"/12; Xấu "&SUMPRODUCT(--(MID($L$28:$L$39,8,1)="X"))&"/12"))
    return ('' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,12,39,12)),0.0)) else (X.cat(('Tốt' if X.truthy(X.eq(X.MID(E.v(6,r-44,c+8), 8.0, 1.0),'T')) else ('Xấu' if X.truthy(X.eq(X.MID(E.v(6,r-44,c+8), 8.0, 1.0),'X')) else 'Thường')),' (12/12)') if X.truthy(X.eq(E.v(6,r,c-1),'ỔN ĐỊNH')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tốt ',X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 8.0, 1.0),'T'))))),'/12; Thường '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 8.0, 1.0),'·'))))),'/12; Xấu '),X.SUMPRODUCT(X.A('neg',X.A('neg',X.A('eq',X.AF_MID(E.rg(6,28,12,39,12), 8.0, 1.0),'X'))))),'/12')))

def f122(E, r, c):
    # Ô C76
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,1),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 1.0), E.rg(6,28,28,39,28), 0.0))

def f123(E, r, c):
    # Ô D76
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,1),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 1.0), E.rg(6,28,28,39,28), 0.0))

def f124(E, r, c):
    # Ô E76, E77, E78, E79, E80, E81, E82, E83, E84, E85, E86, E87
    # = $E$23
    return E.v(6,23,5)

def f125(E, r, c):
    # Ô F76
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,1),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 1.0), E.rg(6,28,28,39,28), 0.0))

def f126(E, r, c):
    # Ô G76
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,1),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 1.0), E.rg(6,28,28,39,28), 0.0))

def f127(E, r, c):
    # Ô H76
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,1),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 1.0), E.rg(6,28,28,39,28), 0.0))

def f128(E, r, c):
    # Ô C77
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,2),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 2.0), E.rg(6,28,28,39,28), 0.0))

def f129(E, r, c):
    # Ô D77
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,2),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 2.0), E.rg(6,28,28,39,28), 0.0))

def f130(E, r, c):
    # Ô F77
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,2),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 2.0), E.rg(6,28,28,39,28), 0.0))

def f131(E, r, c):
    # Ô G77
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,2),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 2.0), E.rg(6,28,28,39,28), 0.0))

def f132(E, r, c):
    # Ô H77
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,2),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 2.0), E.rg(6,28,28,39,28), 0.0))

def f133(E, r, c):
    # Ô C78
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,3),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 3.0), E.rg(6,28,28,39,28), 0.0))

def f134(E, r, c):
    # Ô D78
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,3),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 3.0), E.rg(6,28,28,39,28), 0.0))

def f135(E, r, c):
    # Ô F78
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,3),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 3.0), E.rg(6,28,28,39,28), 0.0))

def f136(E, r, c):
    # Ô G78
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,3),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 3.0), E.rg(6,28,28,39,28), 0.0))

def f137(E, r, c):
    # Ô H78
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,3),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 3.0), E.rg(6,28,28,39,28), 0.0))

def f138(E, r, c):
    # Ô C79
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,4),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 4.0), E.rg(6,28,28,39,28), 0.0))

def f139(E, r, c):
    # Ô D79
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,4),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 4.0), E.rg(6,28,28,39,28), 0.0))

def f140(E, r, c):
    # Ô F79
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,4),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 4.0), E.rg(6,28,28,39,28), 0.0))

def f141(E, r, c):
    # Ô G79
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,4),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 4.0), E.rg(6,28,28,39,28), 0.0))

def f142(E, r, c):
    # Ô H79
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,4),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 4.0), E.rg(6,28,28,39,28), 0.0))

def f143(E, r, c):
    # Ô C80
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,5),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 5.0), E.rg(6,28,28,39,28), 0.0))

def f144(E, r, c):
    # Ô D80
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,5),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 5.0), E.rg(6,28,28,39,28), 0.0))

def f145(E, r, c):
    # Ô F80
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,5),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 5.0), E.rg(6,28,28,39,28), 0.0))

def f146(E, r, c):
    # Ô G80
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,5),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 5.0), E.rg(6,28,28,39,28), 0.0))

def f147(E, r, c):
    # Ô H80
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,5),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 5.0), E.rg(6,28,28,39,28), 0.0))

def f148(E, r, c):
    # Ô C81
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,6),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 6.0), E.rg(6,28,28,39,28), 0.0))

def f149(E, r, c):
    # Ô D81
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,6),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 6.0), E.rg(6,28,28,39,28), 0.0))

def f150(E, r, c):
    # Ô F81
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,6),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 6.0), E.rg(6,28,28,39,28), 0.0))

def f151(E, r, c):
    # Ô G81
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,6),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 6.0), E.rg(6,28,28,39,28), 0.0))

def f152(E, r, c):
    # Ô H81
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,6),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 6.0), E.rg(6,28,28,39,28), 0.0))

def f153(E, r, c):
    # Ô C82
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,7),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 7.0), E.rg(6,28,28,39,28), 0.0))

def f154(E, r, c):
    # Ô D82
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,7),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 7.0), E.rg(6,28,28,39,28), 0.0))

def f155(E, r, c):
    # Ô F82
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,7),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 7.0), E.rg(6,28,28,39,28), 0.0))

def f156(E, r, c):
    # Ô G82
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,7),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 7.0), E.rg(6,28,28,39,28), 0.0))

def f157(E, r, c):
    # Ô H82
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,7),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 7.0), E.rg(6,28,28,39,28), 0.0))

def f158(E, r, c):
    # Ô C83
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,8),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 8.0), E.rg(6,28,28,39,28), 0.0))

def f159(E, r, c):
    # Ô D83
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,8),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 8.0), E.rg(6,28,28,39,28), 0.0))

def f160(E, r, c):
    # Ô F83
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,8),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 8.0), E.rg(6,28,28,39,28), 0.0))

def f161(E, r, c):
    # Ô G83
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,8),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 8.0), E.rg(6,28,28,39,28), 0.0))

def f162(E, r, c):
    # Ô H83
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,8),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 8.0), E.rg(6,28,28,39,28), 0.0))

def f163(E, r, c):
    # Ô C84
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,9),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 9.0), E.rg(6,28,28,39,28), 0.0))

def f164(E, r, c):
    # Ô D84
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,9),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 9.0), E.rg(6,28,28,39,28), 0.0))

def f165(E, r, c):
    # Ô F84
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,9),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 9.0), E.rg(6,28,28,39,28), 0.0))

def f166(E, r, c):
    # Ô G84
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,9),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 9.0), E.rg(6,28,28,39,28), 0.0))

def f167(E, r, c):
    # Ô H84
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,9),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 9.0), E.rg(6,28,28,39,28), 0.0))

def f168(E, r, c):
    # Ô C85
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,10),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 10.0), E.rg(6,28,28,39,28), 0.0))

def f169(E, r, c):
    # Ô D85
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,10),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 10.0), E.rg(6,28,28,39,28), 0.0))

def f170(E, r, c):
    # Ô F85
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,10),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 10.0), E.rg(6,28,28,39,28), 0.0))

def f171(E, r, c):
    # Ô G85
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,10),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 10.0), E.rg(6,28,28,39,28), 0.0))

def f172(E, r, c):
    # Ô H85
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,10),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 10.0), E.rg(6,28,28,39,28), 0.0))

def f173(E, r, c):
    # Ô C86
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,11),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 11.0), E.rg(6,28,28,39,28), 0.0))

def f174(E, r, c):
    # Ô D86
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,11),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 11.0), E.rg(6,28,28,39,28), 0.0))

def f175(E, r, c):
    # Ô F86
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,11),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 11.0), E.rg(6,28,28,39,28), 0.0))

def f176(E, r, c):
    # Ô G86
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,11),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 11.0), E.rg(6,28,28,39,28), 0.0))

def f177(E, r, c):
    # Ô H86
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,11),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 11.0), E.rg(6,28,28,39,28), 0.0))

def f178(E, r, c):
    # Ô C87
    # = INDEX($A$28:$A$39,MATCH(LARGE($AB$28:$AB$39,12),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,1,39,1), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 12.0), E.rg(6,28,28,39,28), 0.0))

def f179(E, r, c):
    # Ô D87
    # = INDEX($Y$28:$Y$39,MATCH(LARGE($AB$28:$AB$39,12),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,25,39,25), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 12.0), E.rg(6,28,28,39,28), 0.0))

def f180(E, r, c):
    # Ô F87
    # = INDEX($Z$28:$Z$39,MATCH(LARGE($AB$28:$AB$39,12),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,26,39,26), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 12.0), E.rg(6,28,28,39,28), 0.0))

def f181(E, r, c):
    # Ô G87
    # = INDEX($F$28:$F$39,MATCH(LARGE($AB$28:$AB$39,12),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,6,39,6), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 12.0), E.rg(6,28,28,39,28), 0.0))

def f182(E, r, c):
    # Ô H87
    # = INDEX($H$28:$H$39,MATCH(LARGE($AB$28:$AB$39,12),$AB$28:$AB$39,0))
    return X.INDEX(E.rg(6,28,8,39,8), X.MATCH(X.LARGE(E.rg(6,28,28,39,28), 12.0), E.rg(6,28,28,39,28), 0.0))

def f183(E, r, c):
    # Ô C92
    # = IF(NOT((NHAP_LIEU!$C$9="Không rõ giờ")),"(Đang ở chế độ ""Biết giờ"" – xem KET_LUAN)",IF(COUNTA($C$28:$C$39)=0,"(Bảng 12 giờ chưa tính – nhấn F9)",IF(
    #   LEFT(C46,8)="ỔN ĐỊNH","• "&"Vượng suy nhật chủ"&": "&D46&CHAR(10),"")&IF(LEFT(C48,8)="ỔN ĐỊNH","• "&"Dụng thần"&": "&D48&CHAR(10),"")&IF(LEFT(C49,8)="
    #   ỔN ĐỊNH","• "&"Hỷ thần"&": "&D49&CHAR(10),"")&IF(LEFT(C50,8)="ỔN ĐỊNH","• "&"Kỵ thần"&": "&D50&CHAR(10),"")&IF(LEFT(C51,8)="ỔN ĐỊNH","• "&"Cách cục"&"
    #   : "&D51&CHAR(10),"")&IF(LEFT(C53,8)="ỔN ĐỊNH","• "&"Bậc giàu"&": "&D53&CHAR(10),"")&IF(LEFT(C54,8)="ỔN ĐỊNH","• "&"Ông bà – tổ nghiệp"&": "&D54&CHAR(1
    #   0),"")&IF(LEFT(C55,8)="ỔN ĐỊNH","• "&"Cha"&": "&D55&CHAR(10),"")&IF(LEFT(C56,8)="ỔN ĐỊNH","• "&"Mẹ"&": "&D56&CHAR(10),"")&IF(LEFT(C57,8)="ỔN ĐỊNH","• 
    #   "&"Anh chị em"&": "&D57&CHAR(10),"")&IF(LEFT(C58,8)="ỔN ĐỊNH","• "&"Con cái"&": "&D58&CHAR(10),"")&IF(LEFT(C59,8)="ỔN ĐỊNH","• "&"Hôn nhân"&": "&D59&C
    #   HAR(10),"")&IF(LEFT(C60,8)="ỔN ĐỊNH","• "&"Giàu – nghèo"&": "&D60&CHAR(10),"")&IF(LEFT(C61,8)="ỔN ĐỊNH","• "&"Sang – hèn"&": "&D61&CHAR(10),"")&IF(LEF
    #   T(C62,8)="ỔN ĐỊNH","• "&"Cát – hung"&": "&D62&CHAR(10),"")&IF(LEFT(C63,8)="ỔN ĐỊNH","• "&"Thọ – yểu"&": "&D63&CHAR(10),"")&IF(LEFT(C64,8)="ỔN ĐỊNH","•
    #    "&"Tố chất cốt lõi"&": "&D64&CHAR(10),"")&IF(LEFT(C65,8)="ỔN ĐỊNH","• "&B65&": "&D65&CHAR(10),"")&IF(LEFT(C66,8)="ỔN ĐỊNH","• "&B66&": "&D66&CHAR(10)
    #   ,"")&IF(LEFT(C67,8)="ỔN ĐỊNH","• "&B67&": "&D67&CHAR(10),"")&IF(LEFT(C68,8)="ỔN ĐỊNH","• "&B68&": "&D68&CHAR(10),"")&IF(LEFT(C69,8)="ỔN ĐỊNH","• "&B69
    #   &": "&D69&CHAR(10),"")&IF(LEFT(C70,8)="ỔN ĐỊNH","• "&B70&": "&D70&CHAR(10),"")&IF(LEFT(C71,8)="ỔN ĐỊNH","• "&B71&": "&D71&CHAR(10),"")&IF(LEFT(C72,8)=
    #   "ỔN ĐỊNH","• "&B72&": "&D72&CHAR(10),"")))
    return ('(Đang ở chế độ "Biết giờ" – xem KET_LUAN)' if X.truthy(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))) else ('(Bảng 12 giờ chưa tính – nhấn F9)' if X.truthy(X.eq(X.COUNTA(E.rg(6,28,3,39,3)),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ','Vượng suy nhật chủ'),': '),E.v(6,r-46,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-46,c), 8.0),'ỔN ĐỊNH')) else ''),(X.cat(X.cat(X.cat(X.cat('• ','Dụng thần'),': '),E.v(6,r-44,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-44,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Hỷ thần'),': '),E.v(6,r-43,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-43,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Kỵ thần'),': '),E.v(6,r-42,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-42,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Cách cục'),': '),E.v(6,r-41,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-41,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Bậc giàu'),': '),E.v(6,r-39,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-39,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Ông bà – tổ nghiệp'),': '),E.v(6,r-38,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-38,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Cha'),': '),E.v(6,r-37,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-37,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Mẹ'),': '),E.v(6,r-36,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-36,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Anh chị em'),': '),E.v(6,r-35,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-35,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Con cái'),': '),E.v(6,r-34,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-34,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Hôn nhân'),': '),E.v(6,r-33,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-33,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Giàu – nghèo'),': '),E.v(6,r-32,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-32,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Sang – hèn'),': '),E.v(6,r-31,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-31,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Cát – hung'),': '),E.v(6,r-30,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-30,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Thọ – yểu'),': '),E.v(6,r-29,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-29,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ','Tố chất cốt lõi'),': '),E.v(6,r-28,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-28,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-27,c-1)),': '),E.v(6,r-27,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-27,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-26,c-1)),': '),E.v(6,r-26,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-26,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-25,c-1)),': '),E.v(6,r-25,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-25,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-24,c-1)),': '),E.v(6,r-24,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-24,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-23,c-1)),': '),E.v(6,r-23,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-23,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-22,c-1)),': '),E.v(6,r-22,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-22,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-21,c-1)),': '),E.v(6,r-21,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-21,c), 8.0),'ỔN ĐỊNH')) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(6,r-20,c-1)),': '),E.v(6,r-20,c+1)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(6,r-20,c), 8.0),'ỔN ĐỊNH')) else ''))))

def f184(E, r, c):
    # Ô C93
    # = IF(NOT((NHAP_LIEU!$C$9="Không rõ giờ")),"",TRIM(IF(C46="PHỤ THUỘC GIỜ","Vượng suy nhật chủ"&"; ","")&IF(C48="PHỤ THUỘC GIỜ","Dụng thần"&"; ","")&IF(C4
    #   9="PHỤ THUỘC GIỜ","Hỷ thần"&"; ","")&IF(C50="PHỤ THUỘC GIỜ","Kỵ thần"&"; ","")&IF(C51="PHỤ THUỘC GIỜ","Cách cục"&"; ","")&IF(C53="PHỤ THUỘC GIỜ","Bậc 
    #   giàu"&"; ","")&IF(C54="PHỤ THUỘC GIỜ","Ông bà – tổ nghiệp"&"; ","")&IF(C55="PHỤ THUỘC GIỜ","Cha"&"; ","")&IF(C56="PHỤ THUỘC GIỜ","Mẹ"&"; ","")&IF(C57=
    #   "PHỤ THUỘC GIỜ","Anh chị em"&"; ","")&IF(C58="PHỤ THUỘC GIỜ","Con cái"&"; ","")&IF(C59="PHỤ THUỘC GIỜ","Hôn nhân"&"; ","")&IF(C60="PHỤ THUỘC GIỜ","Già
    #   u – nghèo"&"; ","")&IF(C61="PHỤ THUỘC GIỜ","Sang – hèn"&"; ","")&IF(C62="PHỤ THUỘC GIỜ","Cát – hung"&"; ","")&IF(C63="PHỤ THUỘC GIỜ","Thọ – yểu"&"; ",
    #   "")&IF(C64="PHỤ THUỘC GIỜ","Tố chất cốt lõi"&"; ","")&IF(C65="PHỤ THUỘC GIỜ",B65&"; ","")&IF(C66="PHỤ THUỘC GIỜ",B66&"; ","")&IF(C67="PHỤ THUỘC GIỜ",B
    #   67&"; ","")&IF(C68="PHỤ THUỘC GIỜ",B68&"; ","")&IF(C69="PHỤ THUỘC GIỜ",B69&"; ","")&IF(C70="PHỤ THUỘC GIỜ",B70&"; ","")&IF(C71="PHỤ THUỘC GIỜ",B71&"; 
    #   ","")&IF(C72="PHỤ THUỘC GIỜ",B72&"; ","")))
    return ('' if X.truthy(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat('Vượng suy nhật chủ','; ') if X.truthy(X.eq(E.v(6,r-47,c),'PHỤ THUỘC GIỜ')) else ''),(X.cat('Dụng thần','; ') if X.truthy(X.eq(E.v(6,r-45,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Hỷ thần','; ') if X.truthy(X.eq(E.v(6,r-44,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Kỵ thần','; ') if X.truthy(X.eq(E.v(6,r-43,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Cách cục','; ') if X.truthy(X.eq(E.v(6,r-42,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Bậc giàu','; ') if X.truthy(X.eq(E.v(6,r-40,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Ông bà – tổ nghiệp','; ') if X.truthy(X.eq(E.v(6,r-39,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Cha','; ') if X.truthy(X.eq(E.v(6,r-38,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Mẹ','; ') if X.truthy(X.eq(E.v(6,r-37,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Anh chị em','; ') if X.truthy(X.eq(E.v(6,r-36,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Con cái','; ') if X.truthy(X.eq(E.v(6,r-35,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Hôn nhân','; ') if X.truthy(X.eq(E.v(6,r-34,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Giàu – nghèo','; ') if X.truthy(X.eq(E.v(6,r-33,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Sang – hèn','; ') if X.truthy(X.eq(E.v(6,r-32,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Cát – hung','; ') if X.truthy(X.eq(E.v(6,r-31,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Thọ – yểu','; ') if X.truthy(X.eq(E.v(6,r-30,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat('Tố chất cốt lõi','; ') if X.truthy(X.eq(E.v(6,r-29,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-28,c-1),'; ') if X.truthy(X.eq(E.v(6,r-28,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-27,c-1),'; ') if X.truthy(X.eq(E.v(6,r-27,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-26,c-1),'; ') if X.truthy(X.eq(E.v(6,r-26,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-25,c-1),'; ') if X.truthy(X.eq(E.v(6,r-25,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-24,c-1),'; ') if X.truthy(X.eq(E.v(6,r-24,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-23,c-1),'; ') if X.truthy(X.eq(E.v(6,r-23,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-22,c-1),'; ') if X.truthy(X.eq(E.v(6,r-22,c),'PHỤ THUỘC GIỜ')) else '')),(X.cat(E.v(6,r-21,c-1),'; ') if X.truthy(X.eq(E.v(6,r-21,c),'PHỤ THUỘC GIỜ')) else ''))))

def f185(E, r, c):
    # Ô C94
    # = IF(OR(NOT((NHAP_LIEU!$C$9="Không rõ giờ")),$E$23=0),"(chưa nhập việc nào ở mục B)",IF(D76=0,"Không giờ nào khớp việc nào – kiểm tra lại năm và loại vi
    #   ệc",C76&": khớp "&D76&"/"&E76&" việc"&IF(D76=D77," – NGANG ĐIỂM với "&C77&IF(D77=D78,", "&C78&"…","")&": chưa phân biệt được, cần nhập thêm việc."," –
    #    hơn giờ thứ hai ("&C77&", "&D77&" việc). Có thể chọn giờ này ở ô C6 để xem toàn bộ lá số.")))
    return ('(chưa nhập việc nào ở mục B)' if X.truthy(X.OR(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(6,23,5),0.0))) else ('Không giờ nào khớp việc nào – kiểm tra lại năm và loại việc' if X.truthy(X.eq(E.v(6,r-18,c+1),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(6,r-18,c),': khớp '),E.v(6,r-18,c+1)),'/'),E.v(6,r-18,c+2)),' việc'),(X.cat(X.cat(X.cat(' – NGANG ĐIỂM với ',E.v(6,r-17,c)),(X.cat(X.cat(', ',E.v(6,r-16,c)),'…') if X.truthy(X.eq(E.v(6,r-17,c+1),E.v(6,r-16,c+1))) else '')),': chưa phân biệt được, cần nhập thêm việc.') if X.truthy(X.eq(E.v(6,r-18,c+1),E.v(6,r-17,c+1))) else X.cat(X.cat(X.cat(X.cat(' – hơn giờ thứ hai (',E.v(6,r-17,c)),', '),E.v(6,r-17,c+1)),' việc). Có thể chọn giờ này ở ô C6 để xem toàn bộ lá số.')))))

def f186(E, r, c):
    # Ô D173
    # = IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C69,DAI_VAN!C54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C69,DAI_VAN!C54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DA
    #   I_VAN!D69,DAI_VAN!D54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!D69,DAI_VAN!D54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E69,DAI_VAN!E54)="Tố
    #   t","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E69,DAI_VAN!E54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F69,DAI_VAN!F54)="Tốt","T",IF(IF(DOI_CHIEU!$C$
    #   33=2,DAI_VAN!F69,DAI_VAN!F54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G69,DAI_VAN!G54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G69,DAI_VAN!G
    #   54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!H69,DAI_VAN!H54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!H69,DAI_VAN!H54)="Xấu","X","·"))&IF(IF(
    #   DOI_CHIEU!$C$33=2,DAI_VAN!I69,DAI_VAN!I54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I69,DAI_VAN!I54)="Xấu","X","·"))&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!
    #   J69,DAI_VAN!J54)="Tốt","T",IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J69,DAI_VAN!J54)="Xấu","X","·"))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('T' if X.truthy(X.eq((E.v(14,r-104,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c-1)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c-1)),'Xấu')) else '·')),('T' if X.truthy(X.eq((E.v(14,r-104,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r-104,c+1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+1)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c+1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+1)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r-104,c+2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+2)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c+2) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+2)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r-104,c+3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+3)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c+3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+3)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r-104,c+4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+4)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c+4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+4)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r-104,c+5) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+5)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c+5) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+5)),'Xấu')) else '·'))),('T' if X.truthy(X.eq((E.v(14,r-104,c+6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+6)),'Tốt')) else ('X' if X.truthy(X.eq((E.v(14,r-104,c+6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-119,c+6)),'Xấu')) else '·')))

CELLS = {
    'C5': (f0, 5, 3),
    'C7': (f1, 7, 3),
    'F11': (f2, 11, 6),
    'G11': (f3, 11, 7),
    'F12': (f2, 12, 6),
    'G12': (f3, 12, 7),
    'F13': (f2, 13, 6),
    'G13': (f3, 13, 7),
    'F14': (f2, 14, 6),
    'G14': (f3, 14, 7),
    'F15': (f2, 15, 6),
    'G15': (f3, 15, 7),
    'F16': (f2, 16, 6),
    'G16': (f3, 16, 7),
    'F17': (f2, 17, 6),
    'G17': (f3, 17, 7),
    'F18': (f2, 18, 6),
    'G18': (f3, 18, 7),
    'F19': (f2, 19, 6),
    'G19': (f3, 19, 7),
    'F20': (f2, 20, 6),
    'G20': (f3, 20, 7),
    'F21': (f2, 21, 6),
    'G21': (f3, 21, 7),
    'F22': (f2, 22, 6),
    'G22': (f3, 22, 7),
    'E23': (f4, 23, 5),
    'F23': (f5, 23, 6),
    'C27': (f6, 27, 3),
    'D27': (f7, 27, 4),
    'E27': (f8, 27, 5),
    'F27': (f9, 27, 6),
    'G27': (f10, 27, 7),
    'H27': (f11, 27, 8),
    'I27': (f12, 27, 9),
    'J27': (f13, 27, 10),
    'K27': (f14, 27, 11),
    'L27': (f15, 27, 12),
    'M27': (f16, 27, 13),
    'N27': (f17, 27, 14),
    'O27': (f18, 27, 15),
    'P27': (f19, 27, 16),
    'Q27': (f20, 27, 17),
    'R27': (f21, 27, 18),
    'S27': (f22, 27, 19),
    'T27': (f23, 27, 20),
    'U27': (f24, 27, 21),
    'V27': (f25, 27, 22),
    'W27': (f26, 27, 23),
    'X27': (f27, 27, 24),
    'Y27': (f28, 27, 25),
    'Z27': (f29, 27, 26),
    'C28': (f30, 28, 3),
    'D28': (f30, 28, 4),
    'E28': (f31, 28, 5),
    'F28': (f31, 28, 6),
    'G28': (f32, 28, 7),
    'H28': (f32, 28, 8),
    'I28': (f33, 28, 9),
    'J28': (f33, 28, 10),
    'K28': (f34, 28, 11),
    'L28': (f34, 28, 12),
    'M28': (f35, 28, 13),
    'N28': (f35, 28, 14),
    'O28': (f36, 28, 15),
    'P28': (f36, 28, 16),
    'Q28': (f37, 28, 17),
    'R28': (f37, 28, 18),
    'S28': (f38, 28, 19),
    'T28': (f38, 28, 20),
    'U28': (f39, 28, 21),
    'V28': (f39, 28, 22),
    'W28': (f40, 28, 23),
    'X28': (f40, 28, 24),
    'Y28': (f41, 28, 25),
    'Z28': (f41, 28, 26),
    'AB28': (f42, 28, 28),
    'C29': (f30, 29, 3),
    'D29': (f30, 29, 4),
    'E29': (f31, 29, 5),
    'F29': (f31, 29, 6),
    'G29': (f32, 29, 7),
    'H29': (f32, 29, 8),
    'I29': (f33, 29, 9),
    'J29': (f33, 29, 10),
    'K29': (f34, 29, 11),
    'L29': (f34, 29, 12),
    'M29': (f35, 29, 13),
    'N29': (f35, 29, 14),
    'O29': (f36, 29, 15),
    'P29': (f36, 29, 16),
    'Q29': (f37, 29, 17),
    'R29': (f37, 29, 18),
    'S29': (f38, 29, 19),
    'T29': (f38, 29, 20),
    'U29': (f39, 29, 21),
    'V29': (f39, 29, 22),
    'W29': (f40, 29, 23),
    'X29': (f40, 29, 24),
    'Y29': (f41, 29, 25),
    'Z29': (f41, 29, 26),
    'AB29': (f43, 29, 28),
    'C30': (f30, 30, 3),
    'D30': (f30, 30, 4),
    'E30': (f31, 30, 5),
    'F30': (f31, 30, 6),
    'G30': (f32, 30, 7),
    'H30': (f32, 30, 8),
    'I30': (f33, 30, 9),
    'J30': (f33, 30, 10),
    'K30': (f34, 30, 11),
    'L30': (f34, 30, 12),
    'M30': (f35, 30, 13),
    'N30': (f35, 30, 14),
    'O30': (f36, 30, 15),
    'P30': (f36, 30, 16),
    'Q30': (f37, 30, 17),
    'R30': (f37, 30, 18),
    'S30': (f38, 30, 19),
    'T30': (f38, 30, 20),
    'U30': (f39, 30, 21),
    'V30': (f39, 30, 22),
    'W30': (f40, 30, 23),
    'X30': (f40, 30, 24),
    'Y30': (f41, 30, 25),
    'Z30': (f41, 30, 26),
    'AB30': (f44, 30, 28),
    'C31': (f30, 31, 3),
    'D31': (f30, 31, 4),
    'E31': (f31, 31, 5),
    'F31': (f31, 31, 6),
    'G31': (f32, 31, 7),
    'H31': (f32, 31, 8),
    'I31': (f33, 31, 9),
    'J31': (f33, 31, 10),
    'K31': (f34, 31, 11),
    'L31': (f34, 31, 12),
    'M31': (f35, 31, 13),
    'N31': (f35, 31, 14),
    'O31': (f36, 31, 15),
    'P31': (f36, 31, 16),
    'Q31': (f37, 31, 17),
    'R31': (f37, 31, 18),
    'S31': (f38, 31, 19),
    'T31': (f38, 31, 20),
    'U31': (f39, 31, 21),
    'V31': (f39, 31, 22),
    'W31': (f40, 31, 23),
    'X31': (f40, 31, 24),
    'Y31': (f41, 31, 25),
    'Z31': (f41, 31, 26),
    'AB31': (f45, 31, 28),
    'C32': (f30, 32, 3),
    'D32': (f30, 32, 4),
    'E32': (f31, 32, 5),
    'F32': (f31, 32, 6),
    'G32': (f32, 32, 7),
    'H32': (f32, 32, 8),
    'I32': (f33, 32, 9),
    'J32': (f33, 32, 10),
    'K32': (f34, 32, 11),
    'L32': (f34, 32, 12),
    'M32': (f35, 32, 13),
    'N32': (f35, 32, 14),
    'O32': (f36, 32, 15),
    'P32': (f36, 32, 16),
    'Q32': (f37, 32, 17),
    'R32': (f37, 32, 18),
    'S32': (f38, 32, 19),
    'T32': (f38, 32, 20),
    'U32': (f39, 32, 21),
    'V32': (f39, 32, 22),
    'W32': (f40, 32, 23),
    'X32': (f40, 32, 24),
    'Y32': (f41, 32, 25),
    'Z32': (f41, 32, 26),
    'AB32': (f46, 32, 28),
    'C33': (f30, 33, 3),
    'D33': (f30, 33, 4),
    'E33': (f31, 33, 5),
    'F33': (f31, 33, 6),
    'G33': (f32, 33, 7),
    'H33': (f32, 33, 8),
    'I33': (f33, 33, 9),
    'J33': (f33, 33, 10),
    'K33': (f34, 33, 11),
    'L33': (f34, 33, 12),
    'M33': (f35, 33, 13),
    'N33': (f35, 33, 14),
    'O33': (f36, 33, 15),
    'P33': (f36, 33, 16),
    'Q33': (f37, 33, 17),
    'R33': (f37, 33, 18),
    'S33': (f38, 33, 19),
    'T33': (f38, 33, 20),
    'U33': (f39, 33, 21),
    'V33': (f39, 33, 22),
    'W33': (f40, 33, 23),
    'X33': (f40, 33, 24),
    'Y33': (f41, 33, 25),
    'Z33': (f41, 33, 26),
    'AB33': (f47, 33, 28),
    'C34': (f30, 34, 3),
    'D34': (f30, 34, 4),
    'E34': (f31, 34, 5),
    'F34': (f31, 34, 6),
    'G34': (f32, 34, 7),
    'H34': (f32, 34, 8),
    'I34': (f33, 34, 9),
    'J34': (f33, 34, 10),
    'K34': (f34, 34, 11),
    'L34': (f34, 34, 12),
    'M34': (f35, 34, 13),
    'N34': (f35, 34, 14),
    'O34': (f36, 34, 15),
    'P34': (f36, 34, 16),
    'Q34': (f37, 34, 17),
    'R34': (f37, 34, 18),
    'S34': (f38, 34, 19),
    'T34': (f38, 34, 20),
    'U34': (f39, 34, 21),
    'V34': (f39, 34, 22),
    'W34': (f40, 34, 23),
    'X34': (f40, 34, 24),
    'Y34': (f41, 34, 25),
    'Z34': (f41, 34, 26),
    'AB34': (f48, 34, 28),
    'C35': (f30, 35, 3),
    'D35': (f30, 35, 4),
    'E35': (f31, 35, 5),
    'F35': (f31, 35, 6),
    'G35': (f32, 35, 7),
    'H35': (f32, 35, 8),
    'I35': (f33, 35, 9),
    'J35': (f33, 35, 10),
    'K35': (f34, 35, 11),
    'L35': (f34, 35, 12),
    'M35': (f35, 35, 13),
    'N35': (f35, 35, 14),
    'O35': (f36, 35, 15),
    'P35': (f36, 35, 16),
    'Q35': (f37, 35, 17),
    'R35': (f37, 35, 18),
    'S35': (f38, 35, 19),
    'T35': (f38, 35, 20),
    'U35': (f39, 35, 21),
    'V35': (f39, 35, 22),
    'W35': (f40, 35, 23),
    'X35': (f40, 35, 24),
    'Y35': (f41, 35, 25),
    'Z35': (f41, 35, 26),
    'AB35': (f49, 35, 28),
    'C36': (f30, 36, 3),
    'D36': (f30, 36, 4),
    'E36': (f31, 36, 5),
    'F36': (f31, 36, 6),
    'G36': (f32, 36, 7),
    'H36': (f32, 36, 8),
    'I36': (f33, 36, 9),
    'J36': (f33, 36, 10),
    'K36': (f34, 36, 11),
    'L36': (f34, 36, 12),
    'M36': (f35, 36, 13),
    'N36': (f35, 36, 14),
    'O36': (f36, 36, 15),
    'P36': (f36, 36, 16),
    'Q36': (f37, 36, 17),
    'R36': (f37, 36, 18),
    'S36': (f38, 36, 19),
    'T36': (f38, 36, 20),
    'U36': (f39, 36, 21),
    'V36': (f39, 36, 22),
    'W36': (f40, 36, 23),
    'X36': (f40, 36, 24),
    'Y36': (f41, 36, 25),
    'Z36': (f41, 36, 26),
    'AB36': (f50, 36, 28),
    'C37': (f30, 37, 3),
    'D37': (f30, 37, 4),
    'E37': (f31, 37, 5),
    'F37': (f31, 37, 6),
    'G37': (f32, 37, 7),
    'H37': (f32, 37, 8),
    'I37': (f33, 37, 9),
    'J37': (f33, 37, 10),
    'K37': (f34, 37, 11),
    'L37': (f34, 37, 12),
    'M37': (f35, 37, 13),
    'N37': (f35, 37, 14),
    'O37': (f36, 37, 15),
    'P37': (f36, 37, 16),
    'Q37': (f37, 37, 17),
    'R37': (f37, 37, 18),
    'S37': (f38, 37, 19),
    'T37': (f38, 37, 20),
    'U37': (f39, 37, 21),
    'V37': (f39, 37, 22),
    'W37': (f40, 37, 23),
    'X37': (f40, 37, 24),
    'Y37': (f41, 37, 25),
    'Z37': (f41, 37, 26),
    'AB37': (f51, 37, 28),
    'C38': (f30, 38, 3),
    'D38': (f30, 38, 4),
    'E38': (f31, 38, 5),
    'F38': (f31, 38, 6),
    'G38': (f32, 38, 7),
    'H38': (f32, 38, 8),
    'I38': (f33, 38, 9),
    'J38': (f33, 38, 10),
    'K38': (f34, 38, 11),
    'L38': (f34, 38, 12),
    'M38': (f35, 38, 13),
    'N38': (f35, 38, 14),
    'O38': (f36, 38, 15),
    'P38': (f36, 38, 16),
    'Q38': (f37, 38, 17),
    'R38': (f37, 38, 18),
    'S38': (f38, 38, 19),
    'T38': (f38, 38, 20),
    'U38': (f39, 38, 21),
    'V38': (f39, 38, 22),
    'W38': (f40, 38, 23),
    'X38': (f40, 38, 24),
    'Y38': (f41, 38, 25),
    'Z38': (f41, 38, 26),
    'AB38': (f52, 38, 28),
    'C39': (f30, 39, 3),
    'D39': (f30, 39, 4),
    'E39': (f31, 39, 5),
    'F39': (f31, 39, 6),
    'G39': (f32, 39, 7),
    'H39': (f32, 39, 8),
    'I39': (f33, 39, 9),
    'J39': (f33, 39, 10),
    'K39': (f34, 39, 11),
    'L39': (f34, 39, 12),
    'M39': (f35, 39, 13),
    'N39': (f35, 39, 14),
    'O39': (f36, 39, 15),
    'P39': (f36, 39, 16),
    'Q39': (f37, 39, 17),
    'R39': (f37, 39, 18),
    'S39': (f38, 39, 19),
    'T39': (f38, 39, 20),
    'U39': (f39, 39, 21),
    'V39': (f39, 39, 22),
    'W39': (f40, 39, 23),
    'X39': (f40, 39, 24),
    'Y39': (f41, 39, 25),
    'Z39': (f41, 39, 26),
    'AB39': (f53, 39, 28),
    'C43': (f54, 43, 3),
    'D43': (f55, 43, 4),
    'C44': (f56, 44, 3),
    'D44': (f57, 44, 4),
    'C45': (f58, 45, 3),
    'D45': (f59, 45, 4),
    'C46': (f60, 46, 3),
    'D46': (f61, 46, 4),
    'C47': (f62, 47, 3),
    'D47': (f63, 47, 4),
    'C48': (f64, 48, 3),
    'D48': (f65, 48, 4),
    'C49': (f66, 49, 3),
    'D49': (f67, 49, 4),
    'C50': (f68, 50, 3),
    'D50': (f69, 50, 4),
    'C51': (f70, 51, 3),
    'D51': (f71, 51, 4),
    'C52': (f72, 52, 3),
    'D52': (f73, 52, 4),
    'C53': (f74, 53, 3),
    'D53': (f75, 53, 4),
    'C54': (f76, 54, 3),
    'D54': (f77, 54, 4),
    'C55': (f78, 55, 3),
    'D55': (f79, 55, 4),
    'C56': (f80, 56, 3),
    'D56': (f81, 56, 4),
    'C57': (f82, 57, 3),
    'D57': (f83, 57, 4),
    'C58': (f84, 58, 3),
    'D58': (f85, 58, 4),
    'C59': (f86, 59, 3),
    'D59': (f87, 59, 4),
    'C60': (f88, 60, 3),
    'D60': (f89, 60, 4),
    'C61': (f90, 61, 3),
    'D61': (f91, 61, 4),
    'C62': (f92, 62, 3),
    'D62': (f93, 62, 4),
    'C63': (f94, 63, 3),
    'D63': (f95, 63, 4),
    'C64': (f96, 64, 3),
    'D64': (f97, 64, 4),
    'B65': (f98, 65, 2),
    'C65': (f99, 65, 3),
    'D65': (f100, 65, 4),
    'B66': (f101, 66, 2),
    'C66': (f102, 66, 3),
    'D66': (f103, 66, 4),
    'B67': (f104, 67, 2),
    'C67': (f105, 67, 3),
    'D67': (f106, 67, 4),
    'B68': (f107, 68, 2),
    'C68': (f108, 68, 3),
    'D68': (f109, 68, 4),
    'B69': (f110, 69, 2),
    'C69': (f111, 69, 3),
    'D69': (f112, 69, 4),
    'B70': (f113, 70, 2),
    'C70': (f114, 70, 3),
    'D70': (f115, 70, 4),
    'B71': (f116, 71, 2),
    'C71': (f117, 71, 3),
    'D71': (f118, 71, 4),
    'B72': (f119, 72, 2),
    'C72': (f120, 72, 3),
    'D72': (f121, 72, 4),
    'C76': (f122, 76, 3),
    'D76': (f123, 76, 4),
    'E76': (f124, 76, 5),
    'F76': (f125, 76, 6),
    'G76': (f126, 76, 7),
    'H76': (f127, 76, 8),
    'C77': (f128, 77, 3),
    'D77': (f129, 77, 4),
    'E77': (f124, 77, 5),
    'F77': (f130, 77, 6),
    'G77': (f131, 77, 7),
    'H77': (f132, 77, 8),
    'C78': (f133, 78, 3),
    'D78': (f134, 78, 4),
    'E78': (f124, 78, 5),
    'F78': (f135, 78, 6),
    'G78': (f136, 78, 7),
    'H78': (f137, 78, 8),
    'C79': (f138, 79, 3),
    'D79': (f139, 79, 4),
    'E79': (f124, 79, 5),
    'F79': (f140, 79, 6),
    'G79': (f141, 79, 7),
    'H79': (f142, 79, 8),
    'C80': (f143, 80, 3),
    'D80': (f144, 80, 4),
    'E80': (f124, 80, 5),
    'F80': (f145, 80, 6),
    'G80': (f146, 80, 7),
    'H80': (f147, 80, 8),
    'C81': (f148, 81, 3),
    'D81': (f149, 81, 4),
    'E81': (f124, 81, 5),
    'F81': (f150, 81, 6),
    'G81': (f151, 81, 7),
    'H81': (f152, 81, 8),
    'C82': (f153, 82, 3),
    'D82': (f154, 82, 4),
    'E82': (f124, 82, 5),
    'F82': (f155, 82, 6),
    'G82': (f156, 82, 7),
    'H82': (f157, 82, 8),
    'C83': (f158, 83, 3),
    'D83': (f159, 83, 4),
    'E83': (f124, 83, 5),
    'F83': (f160, 83, 6),
    'G83': (f161, 83, 7),
    'H83': (f162, 83, 8),
    'C84': (f163, 84, 3),
    'D84': (f164, 84, 4),
    'E84': (f124, 84, 5),
    'F84': (f165, 84, 6),
    'G84': (f166, 84, 7),
    'H84': (f167, 84, 8),
    'C85': (f168, 85, 3),
    'D85': (f169, 85, 4),
    'E85': (f124, 85, 5),
    'F85': (f170, 85, 6),
    'G85': (f171, 85, 7),
    'H85': (f172, 85, 8),
    'C86': (f173, 86, 3),
    'D86': (f174, 86, 4),
    'E86': (f124, 86, 5),
    'F86': (f175, 86, 6),
    'G86': (f176, 86, 7),
    'H86': (f177, 86, 8),
    'C87': (f178, 87, 3),
    'D87': (f179, 87, 4),
    'E87': (f124, 87, 5),
    'F87': (f180, 87, 6),
    'G87': (f181, 87, 7),
    'H87': (f182, 87, 8),
    'C92': (f183, 92, 3),
    'C93': (f184, 93, 3),
    'C94': (f185, 94, 3),
    'C117': (f6, 117, 3),
    'D117': (f7, 117, 4),
    'C131': (f8, 131, 3),
    'D131': (f9, 131, 4),
    'C145': (f10, 145, 3),
    'D145': (f11, 145, 4),
    'C159': (f12, 159, 3),
    'D159': (f13, 159, 4),
    'C173': (f14, 173, 3),
    'D173': (f186, 173, 4),
    'C187': (f16, 187, 3),
    'D187': (f17, 187, 4),
    'C201': (f18, 201, 3),
    'D201': (f19, 201, 4),
    'C215': (f20, 215, 3),
    'D215': (f21, 215, 4),
    'C229': (f22, 229, 3),
    'D229': (f23, 229, 4),
    'C243': (f24, 243, 3),
    'D243': (f25, 243, 4),
    'C257': (f26, 257, 3),
    'D257': (f27, 257, 4),
    'C271': (f28, 271, 3),
    'D271': (f29, 271, 4),
}
