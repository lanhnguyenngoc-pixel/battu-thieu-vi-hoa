# -*- coding: utf-8 -*-
"""Module sheet THEO_DOI – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.9.xlsx (SHA-256 5486ea8b3eb631dc513aed16150152722659ca3ca34dfc6c7af268da54b63453).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'THEO_DOI'
SI = 5

def f0(E, r, c):
    # Ô E4
    # = COUNTA(N12:N111)
    return X.COUNTA(E.rg(5,r+8,c+9,r+107,c+9))

def f1(E, r, c):
    # Ô E5
    # = IF(E4=0,"",SUM(Q12:Q111)/E4)
    return ('' if X.truthy(X.eq(E.v(5,r-1,c),0.0)) else X.div(X.SUM(E.rg(5,r+7,c+12,r+106,c+12)),E.v(5,r-1,c)))

def f2(E, r, c):
    # Ô E6
    # = IF(E4=0,"",SUM(R12:R111)/E4)
    return ('' if X.truthy(X.eq(E.v(5,r-2,c),0.0)) else X.div(X.SUM(E.rg(5,r+6,c+13,r+105,c+13)),E.v(5,r-2,c)))

def f3(E, r, c):
    # Ô E7
    # = IF(E4<5,"Ghi ít nhất 5 năm đã qua có sự việc rõ (tốt / xấu) để so PP1 – PP2 (P1 tr.125).",IF(DOI_CHIEU!$C$33=0,"PP1 và PP2 thống nhất vượng suy; tỉ lệ
    #    khớp dùng để đánh giá độ tin của lá số.",IF(E5>=E6+0.15,"Sự việc khớp PP1 rõ hơn → nên chọn NHAP_LIEU C14 = Theo PP1.",IF(E6>=E5+0.15,"Sự việc khớp P
    #   P2 rõ hơn → nên chọn NHAP_LIEU C14 = Theo PP2.","Chưa phân biệt rõ (chênh dưới 15%) – ghi thêm năm có sự việc rõ."))))
    return ('Ghi ít nhất 5 năm đã qua có sự việc rõ (tốt / xấu) để so PP1 – PP2 (P1 tr.125).' if X.truthy(X.lt(E.v(5,r-3,c),5.0)) else ('PP1 và PP2 thống nhất vượng suy; tỉ lệ khớp dùng để đánh giá độ tin của lá số.' if X.truthy(X.eq(E.v(7,33,3),0.0)) else ('Sự việc khớp PP1 rõ hơn → nên chọn NHAP_LIEU C14 = Theo PP1.' if X.truthy(X.ge(E.v(5,r-2,c),X.add(E.v(5,r-1,c),0.15))) else ('Sự việc khớp PP2 rõ hơn → nên chọn NHAP_LIEU C14 = Theo PP2.' if X.truthy(X.ge(E.v(5,r-1,c),X.add(E.v(5,r-2,c),0.15))) else 'Chưa phân biệt rõ (chênh dưới 15%) – ghi thêm năm có sự việc rõ.'))))

def f4(E, r, c):
    # Ô E8
    # = "Đúng "&COUNTIF(P12:P111,"Đúng")&" · Một phần "&COUNTIF(P12:P111,"Một phần")&" · Sai "&COUNTIF(P12:P111,"Sai")&" · Chưa rõ "&COUNTIF(P12:P111,"Chưa rõ
    #   ")&IF(COUNTIF(P12:P111,"Đúng")+COUNTIF(P12:P111,"Một phần")+COUNTIF(P12:P111,"Sai")=0,"",". Tỉ lệ đúng (một phần tính ½): "&TEXT((COUNTIF(P12:P111,"Đú
    #   ng")+0.5*COUNTIF(P12:P111,"Một phần"))/(COUNTIF(P12:P111,"Đúng")+COUNTIF(P12:P111,"Một phần")+COUNTIF(P12:P111,"Sai")),"0%"))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đúng ',X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Đúng')),' · Một phần '),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Một phần')),' · Sai '),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Sai')),' · Chưa rõ '),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Chưa rõ')),('' if X.truthy(X.eq(X.add(X.add(X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Đúng'),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Một phần')),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Sai')),0.0)) else X.cat('. Tỉ lệ đúng (một phần tính ½): ',X.TEXT(X.div(X.add(X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Đúng'),X.mul(0.5,X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Một phần'))),X.add(X.add(X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Đúng'),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Một phần')),X.COUNTIF(E.rg(5,r+4,c+11,r+103,c+11), 'Sai'))), '0%'))))

def f5(E, r, c):
    # Ô E9
    # = N(NHAP_LIEU!$C$31)
    return X.N(E.v(2,31,3))

def f6(E, r, c):
    # Ô A12
    # = IF((ISNUMBER(LUU_NIEN!$A$30)),LUU_NIEN!$A$30,"")
    return (E.v(15,30,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,30,1))) else '')

def f7(E, r, c):
    # Ô B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, B22, B23 … (100 ô)
    # = IF(A12="","",U12+1&IF(A12=N(NHAP_LIEU!$C$31)," ◄",""))
    return ('' if X.truthy(X.eq(E.v(5,r,c-1),'')) else X.cat(X.add(E.v(5,r,c+19),1.0),(' ◄' if X.truthy(X.eq(E.v(5,r,c-1),X.N(E.v(2,31,3)))) else '')))

def f8(E, r, c):
    # Ô C12
    # = IF(A12="","",LUU_NIEN!$L$30)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,30,12))

def f9(E, r, c):
    # Ô D12
    # = IF(A12="","",LUU_NIEN!$H$30)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,30,8))

def f10(E, r, c):
    # Ô E12
    # = IF(A12="","",LUU_NIEN!$AA$30)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,30,27))

def f11(E, r, c):
    # Ô F12
    # = IF(A12="","",LUU_NIEN!$AL$30)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,30,38))

def f12(E, r, c):
    # Ô G12, G13, G14, G15, G16, G17, G18, G19, G20, G21, G22, G23 … (100 ô)
    # = IF(A12="","",IF(DOI_CHIEU!$C$33=2,F12,E12))
    return ('' if X.truthy(X.eq(E.v(5,r,c-6),'')) else (E.v(5,r,c-1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(5,r,c-2)))

def f13(E, r, c):
    # Ô H12, H13, H14, H15, H16, H17, H18, H19, H20, H21, H22, H23 … (100 ô)
    # = IF(A12="","",IF(V12,"↑ Năm quan tinh là dụng/hỷ: thuận công danh, dễ được giao việc, bổ nhiệm (P1 tr.167-169)",IF(W12,IF($T12=1,"↓ Thực thương gặp qua
    #   n: dễ biến động chức vụ, va chạm cấp trên (P2 tr.541-542)","↓ Quan sát là kỵ: áp lực, dễ bị chèn ép (P1 tr.169-170)"),"")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-7),'')) else ('↑ Năm quan tinh là dụng/hỷ: thuận công danh, dễ được giao việc, bổ nhiệm (P1 tr.167-169)' if X.truthy(E.v(5,r,c+14)) else (('↓ Thực thương gặp quan: dễ biến động chức vụ, va chạm cấp trên (P2 tr.541-542)' if X.truthy(X.eq(E.v(5,r,20),1.0)) else '↓ Quan sát là kỵ: áp lực, dễ bị chèn ép (P1 tr.169-170)') if X.truthy(E.v(5,r,c+15)) else '')))

def f14(E, r, c):
    # Ô I12, I13, I14, I15, I16, I17, I18, I19, I20, I21, I22, I23 … (100 ô)
    # = IF(A12="","",IF(X12,"↑ "&IF($T12=1,"Thực thương sinh tài: thu nhập khá (P2 tr.521-522)","Năm tài là dụng/hỷ: dễ có thêm thu nhập (P1 tr.164-165)"),IF(
    #   Y12,"↓ "&IF($T12=0,"Tỉ kiếp tranh tài: giữ tiền, tránh cho vay, hùn vốn (P1 tr.156; P2 tr.470)","Năm tài là kỵ: vì tiền mà phiền, hao tài (P2 tr.526-5
    #   27)"),"")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-8),'')) else (X.cat('↑ ',('Thực thương sinh tài: thu nhập khá (P2 tr.521-522)' if X.truthy(X.eq(E.v(5,r,20),1.0)) else 'Năm tài là dụng/hỷ: dễ có thêm thu nhập (P1 tr.164-165)')) if X.truthy(E.v(5,r,c+15)) else (X.cat('↓ ',('Tỉ kiếp tranh tài: giữ tiền, tránh cho vay, hùn vốn (P1 tr.156; P2 tr.470)' if X.truthy(X.eq(E.v(5,r,20),0.0)) else 'Năm tài là kỵ: vì tiền mà phiền, hao tài (P2 tr.526-527)')) if X.truthy(E.v(5,r,c+16)) else '')))

def f15(E, r, c):
    # Ô J12, J13, J14, J15, J16, J17, J18, J19, J20, J21, J22, J23 … (100 ô)
    # = IF(A12="","",IF(Z12,"↑ Năm ấn là dụng/hỷ: thuận học hành, thi cử, giấy tờ",""))
    return ('' if X.truthy(X.eq(E.v(5,r,c-9),'')) else ('↑ Năm ấn là dụng/hỷ: thuận học hành, thi cử, giấy tờ' if X.truthy(E.v(5,r,c+16)) else ''))

def f16(E, r, c):
    # Ô K12
    # = IF(A12="","",TRIM(IF(AB12,"Dấu hiệu bất lợi cho cha. ","")&IF(AC12,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA12,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U12>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$30))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,30,66))))) else ''))))

def f17(E, r, c):
    # Ô L12
    # = IF(OR(A12="",LUU_NIEN!$CN$30=""),"",IFERROR(LEFT(LUU_NIEN!$CN$30,FIND(":",LUU_NIEN!$CN$30)-1),LEFT(LUU_NIEN!$CN$30,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,30,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,30,92), X.sub(X.FIND(':', E.v(15,30,92)),1.0)), lambda: X.LEFT(E.v(15,30,92), 120.0)))

def f18(E, r, c):
    # Ô M12
    # = IF(A12="","",IF(LEN(LUU_NIEN!$BN$30)>260,LEFT(LUU_NIEN!$BN$30,260)&"…",LUU_NIEN!$BN$30))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,30,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,30,66)),260.0)) else E.v(15,30,66)))

def f19(E, r, c):
    # Ô Q12, Q13, Q14, Q15, Q16, Q17, Q18, Q19, Q20, Q21, Q22, Q23 … (100 ô)
    # = IF(OR(A12="",N12=""),"",IF(IF(ISNUMBER(SEARCH("Tốt",E12)),1,IF(ISNUMBER(SEARCH("Xấu",E12)),-1,0))=IF(N12="Tốt",1,IF(N12="Xấu",-1,0)),1,0))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-16),''), X.eq(E.v(5,r,c-3),''))) else (1.0 if X.truthy(X.eq((1.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Tốt', E.v(5,r,c-12)))) else (X.neg(1.0) if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Xấu', E.v(5,r,c-12)))) else 0.0)),(1.0 if X.truthy(X.eq(E.v(5,r,c-3),'Tốt')) else (X.neg(1.0) if X.truthy(X.eq(E.v(5,r,c-3),'Xấu')) else 0.0)))) else 0.0))

def f20(E, r, c):
    # Ô R12, R13, R14, R15, R16, R17, R18, R19, R20, R21, R22, R23 … (100 ô)
    # = IF(OR(A12="",N12=""),"",IF(IF(ISNUMBER(SEARCH("Tốt",F12)),1,IF(ISNUMBER(SEARCH("Xấu",F12)),-1,0))=IF(N12="Tốt",1,IF(N12="Xấu",-1,0)),1,0))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-17),''), X.eq(E.v(5,r,c-4),''))) else (1.0 if X.truthy(X.eq((1.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Tốt', E.v(5,r,c-12)))) else (X.neg(1.0) if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Xấu', E.v(5,r,c-12)))) else 0.0)),(1.0 if X.truthy(X.eq(E.v(5,r,c-4),'Tốt')) else (X.neg(1.0) if X.truthy(X.eq(E.v(5,r,c-4),'Xấu')) else 0.0)))) else 0.0))

def f21(E, r, c):
    # Ô T12
    # = IF(OR(A12="",LUU_NIEN!$P$30=""),-1,MOD(LUU_NIEN!$P$30-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,30,16),''))) else X.MOD(X.add(X.sub(E.v(15,30,16),E.v(23,7,5)),5.0), 5.0))

def f22(E, r, c):
    # Ô U12, U13, U14, U15, U16, U17, U18, U19, U20, U21, U22, U23 … (100 ô)
    # = IF(A12="","",A12-$A$12)
    return ('' if X.truthy(X.eq(E.v(5,r,c-20),'')) else X.sub(E.v(5,r,c-20),E.v(5,12,1)))

def f23(E, r, c):
    # Ô V12, V13, V14, V15, V16, V17, V18, V19, V20, V21, V22, V23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=18,$U12<=75,AND($T12=3,OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIA
    #   N!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-21),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),75.0), X.AND(X.eq(E.v(5,r,20),3.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))

def f24(E, r, c):
    # Ô W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22, W23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=18,$U12<=75,OR(AND($T12=1,OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_
    #   GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),AND($T12=3,OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$1
    #   11,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),
    #   MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-22),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),75.0), X.OR(X.AND(X.eq(E.v(5,r,20),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.AND(X.eq(E.v(5,r,20),3.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))))

def f25(E, r, c):
    # Ô X12, X13, X14, X15, X16, X17, X18, X19, X20, X21, X22, X23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=18,$U12<=90,AND(OR($T12=2,$T12=1),OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MO
    #   D(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-23),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),90.0), X.AND(X.OR(X.eq(E.v(5,r,20),2.0), X.eq(E.v(5,r,20),1.0)), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))

def f26(E, r, c):
    # Ô Y12, Y13, Y14, Y15, Y16, Y17, Y18, Y19, Y20, Y21, Y22, Y23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=18,$U12<=90,OR(AND($T12=0,OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_
    #   GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),AND($T12=2,OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$1
    #   11,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),
    #   MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-24),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),90.0), X.OR(X.AND(X.eq(E.v(5,r,20),0.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.AND(X.eq(E.v(5,r,20),2.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))))

def f27(E, r, c):
    # Ô Z12, Z13, Z14, Z15, Z16, Z17, Z18, Z19, Z20, Z21, Z22, Z23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=6,$U12<=30,AND($T12=4,OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN
    #   !$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-25),'')) else X.AND(X.ge(E.v(5,r,21),6.0), X.le(E.v(5,r,21),30.0), X.AND(X.eq(E.v(5,r,20),4.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))

def f28(E, r, c):
    # Ô AA12
    # = IF(A12="",FALSE(),AND($U12>=18,$U12<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$30))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,30,66)))))

def f29(E, r, c):
    # Ô AB12
    # = IF(A12="",FALSE(),AND($U12>=0,$U12<=120,LUU_NIEN!$BH$30=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,30,60),1.0)))

def f30(E, r, c):
    # Ô AC12
    # = IF(A12="",FALSE(),AND($U12>=0,$U12<=120,LUU_NIEN!$BI$30=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,30,61),1.0)))

def f31(E, r, c):
    # Ô AD12
    # = IF(A12="",FALSE(),AND($U12>=0,$U12<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$30))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,30,92)))))

def f32(E, r, c):
    # Ô AE12, AE13, AE14, AE15, AE16, AE17, AE18, AE19, AE20, AE21, AE22, AE23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=0,$U12<=120,IF(ISNUMBER(SEARCH("Tốt",G12)),1,IF(ISNUMBER(SEARCH("Xấu",G12)),-1,0))=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-30),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq((1.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Tốt', E.v(5,r,c-24)))) else (X.neg(1.0) if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Xấu', E.v(5,r,c-24)))) else 0.0)),1.0)))

def f33(E, r, c):
    # Ô AF12, AF13, AF14, AF15, AF16, AF17, AF18, AF19, AF20, AF21, AF22, AF23 … (100 ô)
    # = IF(A12="",FALSE(),AND($U12>=0,$U12<=120,IF(ISNUMBER(SEARCH("Tốt",G12)),1,IF(ISNUMBER(SEARCH("Xấu",G12)),-1,0))=-1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-31),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq((1.0 if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Tốt', E.v(5,r,c-25)))) else (X.neg(1.0) if X.truthy(X.ISNUMBER(lambda: X.SEARCH('Xấu', E.v(5,r,c-25)))) else 0.0)),X.neg(1.0))))

def f34(E, r, c):
    # Ô AG12
    # = ""&IF(AND(V12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-32),' ') if X.truthy(X.AND(E.rg(5,r,c-11,r,c-11), X.le(E.v(5,r,c-32),X.N(E.v(2,31,3))))) else ''))

def f35(E, r, c):
    # Ô AH12
    # = ""&IF(AND(V12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-33),' ') if X.truthy(X.AND(E.rg(5,r,c-12,r,c-12), X.gt(E.v(5,r,c-33),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-33),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f36(E, r, c):
    # Ô AI12
    # = ""&IF(AND(W12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-34),' ') if X.truthy(X.AND(E.rg(5,r,c-12,r,c-12), X.le(E.v(5,r,c-34),X.N(E.v(2,31,3))))) else ''))

def f37(E, r, c):
    # Ô AJ12
    # = ""&IF(AND(W12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-35),' ') if X.truthy(X.AND(E.rg(5,r,c-13,r,c-13), X.gt(E.v(5,r,c-35),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-35),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f38(E, r, c):
    # Ô AK12
    # = ""&IF(AND(X12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-36),' ') if X.truthy(X.AND(E.rg(5,r,c-13,r,c-13), X.le(E.v(5,r,c-36),X.N(E.v(2,31,3))))) else ''))

def f39(E, r, c):
    # Ô AL12
    # = ""&IF(AND(X12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-37),' ') if X.truthy(X.AND(E.rg(5,r,c-14,r,c-14), X.gt(E.v(5,r,c-37),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-37),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f40(E, r, c):
    # Ô AM12
    # = ""&IF(AND(Y12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-38),' ') if X.truthy(X.AND(E.rg(5,r,c-14,r,c-14), X.le(E.v(5,r,c-38),X.N(E.v(2,31,3))))) else ''))

def f41(E, r, c):
    # Ô AN12
    # = ""&IF(AND(Y12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-39),' ') if X.truthy(X.AND(E.rg(5,r,c-15,r,c-15), X.gt(E.v(5,r,c-39),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-39),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f42(E, r, c):
    # Ô AO12
    # = ""&IF(AND(Z12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-40),' ') if X.truthy(X.AND(E.rg(5,r,c-15,r,c-15), X.le(E.v(5,r,c-40),X.N(E.v(2,31,3))))) else ''))

def f43(E, r, c):
    # Ô AP12
    # = ""&IF(AND(Z12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-41),' ') if X.truthy(X.AND(E.rg(5,r,c-16,r,c-16), X.gt(E.v(5,r,c-41),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-41),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f44(E, r, c):
    # Ô AQ12
    # = ""&IF(AND(AA12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-42),' ') if X.truthy(X.AND(E.rg(5,r,c-16,r,c-16), X.le(E.v(5,r,c-42),X.N(E.v(2,31,3))))) else ''))

def f45(E, r, c):
    # Ô AR12
    # = ""&IF(AND(AA12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-43),' ') if X.truthy(X.AND(E.rg(5,r,c-17,r,c-17), X.gt(E.v(5,r,c-43),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-43),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f46(E, r, c):
    # Ô AS12
    # = ""&IF(AND(AB12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-44),' ') if X.truthy(X.AND(E.rg(5,r,c-17,r,c-17), X.le(E.v(5,r,c-44),X.N(E.v(2,31,3))))) else ''))

def f47(E, r, c):
    # Ô AT12
    # = ""&IF(AND(AB12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-45),' ') if X.truthy(X.AND(E.rg(5,r,c-18,r,c-18), X.gt(E.v(5,r,c-45),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-45),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f48(E, r, c):
    # Ô AU12
    # = ""&IF(AND(AC12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-46),' ') if X.truthy(X.AND(E.rg(5,r,c-18,r,c-18), X.le(E.v(5,r,c-46),X.N(E.v(2,31,3))))) else ''))

def f49(E, r, c):
    # Ô AV12
    # = ""&IF(AND(AC12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-47),' ') if X.truthy(X.AND(E.rg(5,r,c-19,r,c-19), X.gt(E.v(5,r,c-47),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-47),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f50(E, r, c):
    # Ô AW12
    # = ""&IF(AND(AD12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-48),' ') if X.truthy(X.AND(E.rg(5,r,c-19,r,c-19), X.le(E.v(5,r,c-48),X.N(E.v(2,31,3))))) else ''))

def f51(E, r, c):
    # Ô AX12
    # = ""&IF(AND(AD12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-49),' ') if X.truthy(X.AND(E.rg(5,r,c-20,r,c-20), X.gt(E.v(5,r,c-49),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-49),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f52(E, r, c):
    # Ô AY12
    # = ""&IF(AND(AE12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-50),' ') if X.truthy(X.AND(E.rg(5,r,c-20,r,c-20), X.le(E.v(5,r,c-50),X.N(E.v(2,31,3))))) else ''))

def f53(E, r, c):
    # Ô AZ12
    # = ""&IF(AND(AE12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-51),' ') if X.truthy(X.AND(E.rg(5,r,c-21,r,c-21), X.gt(E.v(5,r,c-51),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-51),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f54(E, r, c):
    # Ô BA12
    # = ""&IF(AND(AF12,A12<=N(NHAP_LIEU!$C$31)),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-52),' ') if X.truthy(X.AND(E.rg(5,r,c-21,r,c-21), X.le(E.v(5,r,c-52),X.N(E.v(2,31,3))))) else ''))

def f55(E, r, c):
    # Ô BB12
    # = ""&IF(AND(AF12,A12>N(NHAP_LIEU!$C$31),A12<=N(NHAP_LIEU!$C$31)+20),A12&" ","")
    return X.cat('',(X.cat(E.v(5,r,c-53),' ') if X.truthy(X.AND(E.rg(5,r,c-22,r,c-22), X.gt(E.v(5,r,c-53),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-53),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f56(E, r, c):
    # Ô A13
    # = IF((ISNUMBER(LUU_NIEN!$A$31)),LUU_NIEN!$A$31,"")
    return (E.v(15,31,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,31,1))) else '')

def f57(E, r, c):
    # Ô C13
    # = IF(A13="","",LUU_NIEN!$L$31)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,31,12))

def f58(E, r, c):
    # Ô D13
    # = IF(A13="","",LUU_NIEN!$H$31)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,31,8))

def f59(E, r, c):
    # Ô E13
    # = IF(A13="","",LUU_NIEN!$AA$31)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,31,27))

def f60(E, r, c):
    # Ô F13
    # = IF(A13="","",LUU_NIEN!$AL$31)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,31,38))

def f61(E, r, c):
    # Ô K13
    # = IF(A13="","",TRIM(IF(AB13,"Dấu hiệu bất lợi cho cha. ","")&IF(AC13,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA13,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U13>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$31))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,31,66))))) else ''))))

def f62(E, r, c):
    # Ô L13
    # = IF(OR(A13="",LUU_NIEN!$CN$31=""),"",IFERROR(LEFT(LUU_NIEN!$CN$31,FIND(":",LUU_NIEN!$CN$31)-1),LEFT(LUU_NIEN!$CN$31,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,31,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,31,92), X.sub(X.FIND(':', E.v(15,31,92)),1.0)), lambda: X.LEFT(E.v(15,31,92), 120.0)))

def f63(E, r, c):
    # Ô M13
    # = IF(A13="","",IF(LEN(LUU_NIEN!$BN$31)>260,LEFT(LUU_NIEN!$BN$31,260)&"…",LUU_NIEN!$BN$31))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,31,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,31,66)),260.0)) else E.v(15,31,66)))

def f64(E, r, c):
    # Ô T13
    # = IF(OR(A13="",LUU_NIEN!$P$31=""),-1,MOD(LUU_NIEN!$P$31-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,31,16),''))) else X.MOD(X.add(X.sub(E.v(15,31,16),E.v(23,7,5)),5.0), 5.0))

def f65(E, r, c):
    # Ô AA13
    # = IF(A13="",FALSE(),AND($U13>=18,$U13<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$31))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,31,66)))))

def f66(E, r, c):
    # Ô AB13
    # = IF(A13="",FALSE(),AND($U13>=0,$U13<=120,LUU_NIEN!$BH$31=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,31,60),1.0)))

def f67(E, r, c):
    # Ô AC13
    # = IF(A13="",FALSE(),AND($U13>=0,$U13<=120,LUU_NIEN!$BI$31=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,31,61),1.0)))

def f68(E, r, c):
    # Ô AD13
    # = IF(A13="",FALSE(),AND($U13>=0,$U13<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$31))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,31,92)))))

def f69(E, r, c):
    # Ô AG13, AG14, AG15, AG16, AG17, AG18, AG19, AG20, AG21, AG22, AG23, AG24 … (99 ô)
    # = AG12&IF(AND(V13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-32),' ') if X.truthy(X.AND(E.rg(5,r,c-11,r,c-11), X.le(E.v(5,r,c-32),X.N(E.v(2,31,3))))) else ''))

def f70(E, r, c):
    # Ô AH13, AH14, AH15, AH16, AH17, AH18, AH19, AH20, AH21, AH22, AH23, AH24 … (99 ô)
    # = AH12&IF(AND(V13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-33),' ') if X.truthy(X.AND(E.rg(5,r,c-12,r,c-12), X.gt(E.v(5,r,c-33),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-33),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f71(E, r, c):
    # Ô AI13, AI14, AI15, AI16, AI17, AI18, AI19, AI20, AI21, AI22, AI23, AI24 … (99 ô)
    # = AI12&IF(AND(W13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-34),' ') if X.truthy(X.AND(E.rg(5,r,c-12,r,c-12), X.le(E.v(5,r,c-34),X.N(E.v(2,31,3))))) else ''))

def f72(E, r, c):
    # Ô AJ13, AJ14, AJ15, AJ16, AJ17, AJ18, AJ19, AJ20, AJ21, AJ22, AJ23, AJ24 … (99 ô)
    # = AJ12&IF(AND(W13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-35),' ') if X.truthy(X.AND(E.rg(5,r,c-13,r,c-13), X.gt(E.v(5,r,c-35),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-35),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f73(E, r, c):
    # Ô AK13, AK14, AK15, AK16, AK17, AK18, AK19, AK20, AK21, AK22, AK23, AK24 … (99 ô)
    # = AK12&IF(AND(X13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-36),' ') if X.truthy(X.AND(E.rg(5,r,c-13,r,c-13), X.le(E.v(5,r,c-36),X.N(E.v(2,31,3))))) else ''))

def f74(E, r, c):
    # Ô AL13, AL14, AL15, AL16, AL17, AL18, AL19, AL20, AL21, AL22, AL23, AL24 … (99 ô)
    # = AL12&IF(AND(X13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-37),' ') if X.truthy(X.AND(E.rg(5,r,c-14,r,c-14), X.gt(E.v(5,r,c-37),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-37),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f75(E, r, c):
    # Ô AM13, AM14, AM15, AM16, AM17, AM18, AM19, AM20, AM21, AM22, AM23, AM24 … (99 ô)
    # = AM12&IF(AND(Y13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-38),' ') if X.truthy(X.AND(E.rg(5,r,c-14,r,c-14), X.le(E.v(5,r,c-38),X.N(E.v(2,31,3))))) else ''))

def f76(E, r, c):
    # Ô AN13, AN14, AN15, AN16, AN17, AN18, AN19, AN20, AN21, AN22, AN23, AN24 … (99 ô)
    # = AN12&IF(AND(Y13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-39),' ') if X.truthy(X.AND(E.rg(5,r,c-15,r,c-15), X.gt(E.v(5,r,c-39),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-39),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f77(E, r, c):
    # Ô AO13, AO14, AO15, AO16, AO17, AO18, AO19, AO20, AO21, AO22, AO23, AO24 … (99 ô)
    # = AO12&IF(AND(Z13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-40),' ') if X.truthy(X.AND(E.rg(5,r,c-15,r,c-15), X.le(E.v(5,r,c-40),X.N(E.v(2,31,3))))) else ''))

def f78(E, r, c):
    # Ô AP13, AP14, AP15, AP16, AP17, AP18, AP19, AP20, AP21, AP22, AP23, AP24 … (99 ô)
    # = AP12&IF(AND(Z13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-41),' ') if X.truthy(X.AND(E.rg(5,r,c-16,r,c-16), X.gt(E.v(5,r,c-41),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-41),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f79(E, r, c):
    # Ô AQ13, AQ14, AQ15, AQ16, AQ17, AQ18, AQ19, AQ20, AQ21, AQ22, AQ23, AQ24 … (99 ô)
    # = AQ12&IF(AND(AA13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-42),' ') if X.truthy(X.AND(E.rg(5,r,c-16,r,c-16), X.le(E.v(5,r,c-42),X.N(E.v(2,31,3))))) else ''))

def f80(E, r, c):
    # Ô AR13, AR14, AR15, AR16, AR17, AR18, AR19, AR20, AR21, AR22, AR23, AR24 … (99 ô)
    # = AR12&IF(AND(AA13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-43),' ') if X.truthy(X.AND(E.rg(5,r,c-17,r,c-17), X.gt(E.v(5,r,c-43),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-43),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f81(E, r, c):
    # Ô AS13, AS14, AS15, AS16, AS17, AS18, AS19, AS20, AS21, AS22, AS23, AS24 … (99 ô)
    # = AS12&IF(AND(AB13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-44),' ') if X.truthy(X.AND(E.rg(5,r,c-17,r,c-17), X.le(E.v(5,r,c-44),X.N(E.v(2,31,3))))) else ''))

def f82(E, r, c):
    # Ô AT13, AT14, AT15, AT16, AT17, AT18, AT19, AT20, AT21, AT22, AT23, AT24 … (99 ô)
    # = AT12&IF(AND(AB13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-45),' ') if X.truthy(X.AND(E.rg(5,r,c-18,r,c-18), X.gt(E.v(5,r,c-45),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-45),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f83(E, r, c):
    # Ô AU13, AU14, AU15, AU16, AU17, AU18, AU19, AU20, AU21, AU22, AU23, AU24 … (99 ô)
    # = AU12&IF(AND(AC13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-46),' ') if X.truthy(X.AND(E.rg(5,r,c-18,r,c-18), X.le(E.v(5,r,c-46),X.N(E.v(2,31,3))))) else ''))

def f84(E, r, c):
    # Ô AV13, AV14, AV15, AV16, AV17, AV18, AV19, AV20, AV21, AV22, AV23, AV24 … (99 ô)
    # = AV12&IF(AND(AC13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-47),' ') if X.truthy(X.AND(E.rg(5,r,c-19,r,c-19), X.gt(E.v(5,r,c-47),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-47),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f85(E, r, c):
    # Ô AW13, AW14, AW15, AW16, AW17, AW18, AW19, AW20, AW21, AW22, AW23, AW24 … (99 ô)
    # = AW12&IF(AND(AD13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-48),' ') if X.truthy(X.AND(E.rg(5,r,c-19,r,c-19), X.le(E.v(5,r,c-48),X.N(E.v(2,31,3))))) else ''))

def f86(E, r, c):
    # Ô AX13, AX14, AX15, AX16, AX17, AX18, AX19, AX20, AX21, AX22, AX23, AX24 … (99 ô)
    # = AX12&IF(AND(AD13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-49),' ') if X.truthy(X.AND(E.rg(5,r,c-20,r,c-20), X.gt(E.v(5,r,c-49),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-49),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f87(E, r, c):
    # Ô AY13, AY14, AY15, AY16, AY17, AY18, AY19, AY20, AY21, AY22, AY23, AY24 … (99 ô)
    # = AY12&IF(AND(AE13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-50),' ') if X.truthy(X.AND(E.rg(5,r,c-20,r,c-20), X.le(E.v(5,r,c-50),X.N(E.v(2,31,3))))) else ''))

def f88(E, r, c):
    # Ô AZ13, AZ14, AZ15, AZ16, AZ17, AZ18, AZ19, AZ20, AZ21, AZ22, AZ23, AZ24 … (99 ô)
    # = AZ12&IF(AND(AE13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-51),' ') if X.truthy(X.AND(E.rg(5,r,c-21,r,c-21), X.gt(E.v(5,r,c-51),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-51),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f89(E, r, c):
    # Ô BA13, BA14, BA15, BA16, BA17, BA18, BA19, BA20, BA21, BA22, BA23, BA24 … (99 ô)
    # = BA12&IF(AND(AF13,A13<=N(NHAP_LIEU!$C$31)),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-52),' ') if X.truthy(X.AND(E.rg(5,r,c-21,r,c-21), X.le(E.v(5,r,c-52),X.N(E.v(2,31,3))))) else ''))

def f90(E, r, c):
    # Ô BB13, BB14, BB15, BB16, BB17, BB18, BB19, BB20, BB21, BB22, BB23, BB24 … (99 ô)
    # = BB12&IF(AND(AF13,A13>N(NHAP_LIEU!$C$31),A13<=N(NHAP_LIEU!$C$31)+20),A13&" ","")
    return X.cat(E.v(5,r-1,c),(X.cat(E.v(5,r,c-53),' ') if X.truthy(X.AND(E.rg(5,r,c-22,r,c-22), X.gt(E.v(5,r,c-53),X.N(E.v(2,31,3))), X.le(E.v(5,r,c-53),X.add(X.N(E.v(2,31,3)),20.0)))) else ''))

def f91(E, r, c):
    # Ô A14
    # = IF((ISNUMBER(LUU_NIEN!$A$32)),LUU_NIEN!$A$32,"")
    return (E.v(15,32,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,32,1))) else '')

def f92(E, r, c):
    # Ô C14
    # = IF(A14="","",LUU_NIEN!$L$32)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,32,12))

def f93(E, r, c):
    # Ô D14
    # = IF(A14="","",LUU_NIEN!$H$32)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,32,8))

def f94(E, r, c):
    # Ô E14
    # = IF(A14="","",LUU_NIEN!$AA$32)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,32,27))

def f95(E, r, c):
    # Ô F14
    # = IF(A14="","",LUU_NIEN!$AL$32)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,32,38))

def f96(E, r, c):
    # Ô K14
    # = IF(A14="","",TRIM(IF(AB14,"Dấu hiệu bất lợi cho cha. ","")&IF(AC14,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA14,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U14>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$32))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,32,66))))) else ''))))

def f97(E, r, c):
    # Ô L14
    # = IF(OR(A14="",LUU_NIEN!$CN$32=""),"",IFERROR(LEFT(LUU_NIEN!$CN$32,FIND(":",LUU_NIEN!$CN$32)-1),LEFT(LUU_NIEN!$CN$32,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,32,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,32,92), X.sub(X.FIND(':', E.v(15,32,92)),1.0)), lambda: X.LEFT(E.v(15,32,92), 120.0)))

def f98(E, r, c):
    # Ô M14
    # = IF(A14="","",IF(LEN(LUU_NIEN!$BN$32)>260,LEFT(LUU_NIEN!$BN$32,260)&"…",LUU_NIEN!$BN$32))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,32,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,32,66)),260.0)) else E.v(15,32,66)))

def f99(E, r, c):
    # Ô T14
    # = IF(OR(A14="",LUU_NIEN!$P$32=""),-1,MOD(LUU_NIEN!$P$32-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,32,16),''))) else X.MOD(X.add(X.sub(E.v(15,32,16),E.v(23,7,5)),5.0), 5.0))

def f100(E, r, c):
    # Ô AA14
    # = IF(A14="",FALSE(),AND($U14>=18,$U14<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$32))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,32,66)))))

def f101(E, r, c):
    # Ô AB14
    # = IF(A14="",FALSE(),AND($U14>=0,$U14<=120,LUU_NIEN!$BH$32=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,32,60),1.0)))

def f102(E, r, c):
    # Ô AC14
    # = IF(A14="",FALSE(),AND($U14>=0,$U14<=120,LUU_NIEN!$BI$32=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,32,61),1.0)))

def f103(E, r, c):
    # Ô AD14
    # = IF(A14="",FALSE(),AND($U14>=0,$U14<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$32))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,32,92)))))

def f104(E, r, c):
    # Ô A15
    # = IF((ISNUMBER(LUU_NIEN!$A$33)),LUU_NIEN!$A$33,"")
    return (E.v(15,33,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,33,1))) else '')

def f105(E, r, c):
    # Ô C15
    # = IF(A15="","",LUU_NIEN!$L$33)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,33,12))

def f106(E, r, c):
    # Ô D15
    # = IF(A15="","",LUU_NIEN!$H$33)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,33,8))

def f107(E, r, c):
    # Ô E15
    # = IF(A15="","",LUU_NIEN!$AA$33)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,33,27))

def f108(E, r, c):
    # Ô F15
    # = IF(A15="","",LUU_NIEN!$AL$33)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,33,38))

def f109(E, r, c):
    # Ô K15
    # = IF(A15="","",TRIM(IF(AB15,"Dấu hiệu bất lợi cho cha. ","")&IF(AC15,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA15,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U15>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$33))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,33,66))))) else ''))))

def f110(E, r, c):
    # Ô L15
    # = IF(OR(A15="",LUU_NIEN!$CN$33=""),"",IFERROR(LEFT(LUU_NIEN!$CN$33,FIND(":",LUU_NIEN!$CN$33)-1),LEFT(LUU_NIEN!$CN$33,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,33,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,33,92), X.sub(X.FIND(':', E.v(15,33,92)),1.0)), lambda: X.LEFT(E.v(15,33,92), 120.0)))

def f111(E, r, c):
    # Ô M15
    # = IF(A15="","",IF(LEN(LUU_NIEN!$BN$33)>260,LEFT(LUU_NIEN!$BN$33,260)&"…",LUU_NIEN!$BN$33))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,33,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,33,66)),260.0)) else E.v(15,33,66)))

def f112(E, r, c):
    # Ô T15
    # = IF(OR(A15="",LUU_NIEN!$P$33=""),-1,MOD(LUU_NIEN!$P$33-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,33,16),''))) else X.MOD(X.add(X.sub(E.v(15,33,16),E.v(23,7,5)),5.0), 5.0))

def f113(E, r, c):
    # Ô AA15
    # = IF(A15="",FALSE(),AND($U15>=18,$U15<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$33))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,33,66)))))

def f114(E, r, c):
    # Ô AB15
    # = IF(A15="",FALSE(),AND($U15>=0,$U15<=120,LUU_NIEN!$BH$33=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,33,60),1.0)))

def f115(E, r, c):
    # Ô AC15
    # = IF(A15="",FALSE(),AND($U15>=0,$U15<=120,LUU_NIEN!$BI$33=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,33,61),1.0)))

def f116(E, r, c):
    # Ô AD15
    # = IF(A15="",FALSE(),AND($U15>=0,$U15<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$33))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,33,92)))))

def f117(E, r, c):
    # Ô A16
    # = IF((ISNUMBER(LUU_NIEN!$A$34)),LUU_NIEN!$A$34,"")
    return (E.v(15,34,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,34,1))) else '')

def f118(E, r, c):
    # Ô C16
    # = IF(A16="","",LUU_NIEN!$L$34)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,34,12))

def f119(E, r, c):
    # Ô D16
    # = IF(A16="","",LUU_NIEN!$H$34)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,34,8))

def f120(E, r, c):
    # Ô E16
    # = IF(A16="","",LUU_NIEN!$AA$34)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,34,27))

def f121(E, r, c):
    # Ô F16
    # = IF(A16="","",LUU_NIEN!$AL$34)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,34,38))

def f122(E, r, c):
    # Ô K16
    # = IF(A16="","",TRIM(IF(AB16,"Dấu hiệu bất lợi cho cha. ","")&IF(AC16,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA16,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U16>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$34))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,34,66))))) else ''))))

def f123(E, r, c):
    # Ô L16
    # = IF(OR(A16="",LUU_NIEN!$CN$34=""),"",IFERROR(LEFT(LUU_NIEN!$CN$34,FIND(":",LUU_NIEN!$CN$34)-1),LEFT(LUU_NIEN!$CN$34,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,34,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,34,92), X.sub(X.FIND(':', E.v(15,34,92)),1.0)), lambda: X.LEFT(E.v(15,34,92), 120.0)))

def f124(E, r, c):
    # Ô M16
    # = IF(A16="","",IF(LEN(LUU_NIEN!$BN$34)>260,LEFT(LUU_NIEN!$BN$34,260)&"…",LUU_NIEN!$BN$34))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,34,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,34,66)),260.0)) else E.v(15,34,66)))

def f125(E, r, c):
    # Ô T16
    # = IF(OR(A16="",LUU_NIEN!$P$34=""),-1,MOD(LUU_NIEN!$P$34-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,34,16),''))) else X.MOD(X.add(X.sub(E.v(15,34,16),E.v(23,7,5)),5.0), 5.0))

def f126(E, r, c):
    # Ô AA16
    # = IF(A16="",FALSE(),AND($U16>=18,$U16<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$34))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,34,66)))))

def f127(E, r, c):
    # Ô AB16
    # = IF(A16="",FALSE(),AND($U16>=0,$U16<=120,LUU_NIEN!$BH$34=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,34,60),1.0)))

def f128(E, r, c):
    # Ô AC16
    # = IF(A16="",FALSE(),AND($U16>=0,$U16<=120,LUU_NIEN!$BI$34=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,34,61),1.0)))

def f129(E, r, c):
    # Ô AD16
    # = IF(A16="",FALSE(),AND($U16>=0,$U16<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$34))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,34,92)))))

def f130(E, r, c):
    # Ô A17
    # = IF((ISNUMBER(LUU_NIEN!$A$35)),LUU_NIEN!$A$35,"")
    return (E.v(15,35,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,35,1))) else '')

def f131(E, r, c):
    # Ô C17
    # = IF(A17="","",LUU_NIEN!$L$35)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,35,12))

def f132(E, r, c):
    # Ô D17
    # = IF(A17="","",LUU_NIEN!$H$35)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,35,8))

def f133(E, r, c):
    # Ô E17
    # = IF(A17="","",LUU_NIEN!$AA$35)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,35,27))

def f134(E, r, c):
    # Ô F17
    # = IF(A17="","",LUU_NIEN!$AL$35)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,35,38))

def f135(E, r, c):
    # Ô K17
    # = IF(A17="","",TRIM(IF(AB17,"Dấu hiệu bất lợi cho cha. ","")&IF(AC17,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA17,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U17>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$35))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,35,66))))) else ''))))

def f136(E, r, c):
    # Ô L17
    # = IF(OR(A17="",LUU_NIEN!$CN$35=""),"",IFERROR(LEFT(LUU_NIEN!$CN$35,FIND(":",LUU_NIEN!$CN$35)-1),LEFT(LUU_NIEN!$CN$35,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,35,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,35,92), X.sub(X.FIND(':', E.v(15,35,92)),1.0)), lambda: X.LEFT(E.v(15,35,92), 120.0)))

def f137(E, r, c):
    # Ô M17
    # = IF(A17="","",IF(LEN(LUU_NIEN!$BN$35)>260,LEFT(LUU_NIEN!$BN$35,260)&"…",LUU_NIEN!$BN$35))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,35,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,35,66)),260.0)) else E.v(15,35,66)))

def f138(E, r, c):
    # Ô T17
    # = IF(OR(A17="",LUU_NIEN!$P$35=""),-1,MOD(LUU_NIEN!$P$35-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,35,16),''))) else X.MOD(X.add(X.sub(E.v(15,35,16),E.v(23,7,5)),5.0), 5.0))

def f139(E, r, c):
    # Ô AA17
    # = IF(A17="",FALSE(),AND($U17>=18,$U17<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$35))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,35,66)))))

def f140(E, r, c):
    # Ô AB17
    # = IF(A17="",FALSE(),AND($U17>=0,$U17<=120,LUU_NIEN!$BH$35=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,35,60),1.0)))

def f141(E, r, c):
    # Ô AC17
    # = IF(A17="",FALSE(),AND($U17>=0,$U17<=120,LUU_NIEN!$BI$35=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,35,61),1.0)))

def f142(E, r, c):
    # Ô AD17
    # = IF(A17="",FALSE(),AND($U17>=0,$U17<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$35))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,35,92)))))

def f143(E, r, c):
    # Ô A18
    # = IF((ISNUMBER(LUU_NIEN!$A$36)),LUU_NIEN!$A$36,"")
    return (E.v(15,36,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,36,1))) else '')

def f144(E, r, c):
    # Ô C18
    # = IF(A18="","",LUU_NIEN!$L$36)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,36,12))

def f145(E, r, c):
    # Ô D18
    # = IF(A18="","",LUU_NIEN!$H$36)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,36,8))

def f146(E, r, c):
    # Ô E18
    # = IF(A18="","",LUU_NIEN!$AA$36)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,36,27))

def f147(E, r, c):
    # Ô F18
    # = IF(A18="","",LUU_NIEN!$AL$36)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,36,38))

def f148(E, r, c):
    # Ô K18
    # = IF(A18="","",TRIM(IF(AB18,"Dấu hiệu bất lợi cho cha. ","")&IF(AC18,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA18,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U18>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$36))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,36,66))))) else ''))))

def f149(E, r, c):
    # Ô L18
    # = IF(OR(A18="",LUU_NIEN!$CN$36=""),"",IFERROR(LEFT(LUU_NIEN!$CN$36,FIND(":",LUU_NIEN!$CN$36)-1),LEFT(LUU_NIEN!$CN$36,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,36,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,36,92), X.sub(X.FIND(':', E.v(15,36,92)),1.0)), lambda: X.LEFT(E.v(15,36,92), 120.0)))

def f150(E, r, c):
    # Ô M18
    # = IF(A18="","",IF(LEN(LUU_NIEN!$BN$36)>260,LEFT(LUU_NIEN!$BN$36,260)&"…",LUU_NIEN!$BN$36))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,36,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,36,66)),260.0)) else E.v(15,36,66)))

def f151(E, r, c):
    # Ô T18
    # = IF(OR(A18="",LUU_NIEN!$P$36=""),-1,MOD(LUU_NIEN!$P$36-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,36,16),''))) else X.MOD(X.add(X.sub(E.v(15,36,16),E.v(23,7,5)),5.0), 5.0))

def f152(E, r, c):
    # Ô AA18
    # = IF(A18="",FALSE(),AND($U18>=18,$U18<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$36))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,36,66)))))

def f153(E, r, c):
    # Ô AB18
    # = IF(A18="",FALSE(),AND($U18>=0,$U18<=120,LUU_NIEN!$BH$36=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,36,60),1.0)))

def f154(E, r, c):
    # Ô AC18
    # = IF(A18="",FALSE(),AND($U18>=0,$U18<=120,LUU_NIEN!$BI$36=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,36,61),1.0)))

def f155(E, r, c):
    # Ô AD18
    # = IF(A18="",FALSE(),AND($U18>=0,$U18<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$36))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,36,92)))))

def f156(E, r, c):
    # Ô A19
    # = IF((ISNUMBER(LUU_NIEN!$A$37)),LUU_NIEN!$A$37,"")
    return (E.v(15,37,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,37,1))) else '')

def f157(E, r, c):
    # Ô C19
    # = IF(A19="","",LUU_NIEN!$L$37)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,37,12))

def f158(E, r, c):
    # Ô D19
    # = IF(A19="","",LUU_NIEN!$H$37)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,37,8))

def f159(E, r, c):
    # Ô E19
    # = IF(A19="","",LUU_NIEN!$AA$37)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,37,27))

def f160(E, r, c):
    # Ô F19
    # = IF(A19="","",LUU_NIEN!$AL$37)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,37,38))

def f161(E, r, c):
    # Ô K19
    # = IF(A19="","",TRIM(IF(AB19,"Dấu hiệu bất lợi cho cha. ","")&IF(AC19,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA19,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U19>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$37))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,37,66))))) else ''))))

def f162(E, r, c):
    # Ô L19
    # = IF(OR(A19="",LUU_NIEN!$CN$37=""),"",IFERROR(LEFT(LUU_NIEN!$CN$37,FIND(":",LUU_NIEN!$CN$37)-1),LEFT(LUU_NIEN!$CN$37,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,37,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,37,92), X.sub(X.FIND(':', E.v(15,37,92)),1.0)), lambda: X.LEFT(E.v(15,37,92), 120.0)))

def f163(E, r, c):
    # Ô M19
    # = IF(A19="","",IF(LEN(LUU_NIEN!$BN$37)>260,LEFT(LUU_NIEN!$BN$37,260)&"…",LUU_NIEN!$BN$37))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,37,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,37,66)),260.0)) else E.v(15,37,66)))

def f164(E, r, c):
    # Ô T19
    # = IF(OR(A19="",LUU_NIEN!$P$37=""),-1,MOD(LUU_NIEN!$P$37-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,37,16),''))) else X.MOD(X.add(X.sub(E.v(15,37,16),E.v(23,7,5)),5.0), 5.0))

def f165(E, r, c):
    # Ô AA19
    # = IF(A19="",FALSE(),AND($U19>=18,$U19<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$37))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,37,66)))))

def f166(E, r, c):
    # Ô AB19
    # = IF(A19="",FALSE(),AND($U19>=0,$U19<=120,LUU_NIEN!$BH$37=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,37,60),1.0)))

def f167(E, r, c):
    # Ô AC19
    # = IF(A19="",FALSE(),AND($U19>=0,$U19<=120,LUU_NIEN!$BI$37=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,37,61),1.0)))

def f168(E, r, c):
    # Ô AD19
    # = IF(A19="",FALSE(),AND($U19>=0,$U19<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$37))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,37,92)))))

def f169(E, r, c):
    # Ô A20
    # = IF((ISNUMBER(LUU_NIEN!$A$38)),LUU_NIEN!$A$38,"")
    return (E.v(15,38,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,38,1))) else '')

def f170(E, r, c):
    # Ô C20
    # = IF(A20="","",LUU_NIEN!$L$38)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,38,12))

def f171(E, r, c):
    # Ô D20
    # = IF(A20="","",LUU_NIEN!$H$38)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,38,8))

def f172(E, r, c):
    # Ô E20
    # = IF(A20="","",LUU_NIEN!$AA$38)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,38,27))

def f173(E, r, c):
    # Ô F20
    # = IF(A20="","",LUU_NIEN!$AL$38)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,38,38))

def f174(E, r, c):
    # Ô K20
    # = IF(A20="","",TRIM(IF(AB20,"Dấu hiệu bất lợi cho cha. ","")&IF(AC20,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA20,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U20>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$38))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,38,66))))) else ''))))

def f175(E, r, c):
    # Ô L20
    # = IF(OR(A20="",LUU_NIEN!$CN$38=""),"",IFERROR(LEFT(LUU_NIEN!$CN$38,FIND(":",LUU_NIEN!$CN$38)-1),LEFT(LUU_NIEN!$CN$38,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,38,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,38,92), X.sub(X.FIND(':', E.v(15,38,92)),1.0)), lambda: X.LEFT(E.v(15,38,92), 120.0)))

def f176(E, r, c):
    # Ô M20
    # = IF(A20="","",IF(LEN(LUU_NIEN!$BN$38)>260,LEFT(LUU_NIEN!$BN$38,260)&"…",LUU_NIEN!$BN$38))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,38,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,38,66)),260.0)) else E.v(15,38,66)))

def f177(E, r, c):
    # Ô T20
    # = IF(OR(A20="",LUU_NIEN!$P$38=""),-1,MOD(LUU_NIEN!$P$38-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,38,16),''))) else X.MOD(X.add(X.sub(E.v(15,38,16),E.v(23,7,5)),5.0), 5.0))

def f178(E, r, c):
    # Ô AA20
    # = IF(A20="",FALSE(),AND($U20>=18,$U20<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$38))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,38,66)))))

def f179(E, r, c):
    # Ô AB20
    # = IF(A20="",FALSE(),AND($U20>=0,$U20<=120,LUU_NIEN!$BH$38=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,38,60),1.0)))

def f180(E, r, c):
    # Ô AC20
    # = IF(A20="",FALSE(),AND($U20>=0,$U20<=120,LUU_NIEN!$BI$38=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,38,61),1.0)))

def f181(E, r, c):
    # Ô AD20
    # = IF(A20="",FALSE(),AND($U20>=0,$U20<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$38))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,38,92)))))

def f182(E, r, c):
    # Ô A21
    # = IF((ISNUMBER(LUU_NIEN!$A$39)),LUU_NIEN!$A$39,"")
    return (E.v(15,39,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,39,1))) else '')

def f183(E, r, c):
    # Ô C21
    # = IF(A21="","",LUU_NIEN!$L$39)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,39,12))

def f184(E, r, c):
    # Ô D21
    # = IF(A21="","",LUU_NIEN!$H$39)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,39,8))

def f185(E, r, c):
    # Ô E21
    # = IF(A21="","",LUU_NIEN!$AA$39)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,39,27))

def f186(E, r, c):
    # Ô F21
    # = IF(A21="","",LUU_NIEN!$AL$39)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,39,38))

def f187(E, r, c):
    # Ô K21
    # = IF(A21="","",TRIM(IF(AB21,"Dấu hiệu bất lợi cho cha. ","")&IF(AC21,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA21,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U21>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$39))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,39,66))))) else ''))))

def f188(E, r, c):
    # Ô L21
    # = IF(OR(A21="",LUU_NIEN!$CN$39=""),"",IFERROR(LEFT(LUU_NIEN!$CN$39,FIND(":",LUU_NIEN!$CN$39)-1),LEFT(LUU_NIEN!$CN$39,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,39,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,39,92), X.sub(X.FIND(':', E.v(15,39,92)),1.0)), lambda: X.LEFT(E.v(15,39,92), 120.0)))

def f189(E, r, c):
    # Ô M21
    # = IF(A21="","",IF(LEN(LUU_NIEN!$BN$39)>260,LEFT(LUU_NIEN!$BN$39,260)&"…",LUU_NIEN!$BN$39))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,39,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,39,66)),260.0)) else E.v(15,39,66)))

def f190(E, r, c):
    # Ô T21
    # = IF(OR(A21="",LUU_NIEN!$P$39=""),-1,MOD(LUU_NIEN!$P$39-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,39,16),''))) else X.MOD(X.add(X.sub(E.v(15,39,16),E.v(23,7,5)),5.0), 5.0))

def f191(E, r, c):
    # Ô AA21
    # = IF(A21="",FALSE(),AND($U21>=18,$U21<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$39))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,39,66)))))

def f192(E, r, c):
    # Ô AB21
    # = IF(A21="",FALSE(),AND($U21>=0,$U21<=120,LUU_NIEN!$BH$39=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,39,60),1.0)))

def f193(E, r, c):
    # Ô AC21
    # = IF(A21="",FALSE(),AND($U21>=0,$U21<=120,LUU_NIEN!$BI$39=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,39,61),1.0)))

def f194(E, r, c):
    # Ô AD21
    # = IF(A21="",FALSE(),AND($U21>=0,$U21<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$39))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,39,92)))))

def f195(E, r, c):
    # Ô A22
    # = IF((ISNUMBER(LUU_NIEN!$A$40)),LUU_NIEN!$A$40,"")
    return (E.v(15,40,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,40,1))) else '')

def f196(E, r, c):
    # Ô C22
    # = IF(A22="","",LUU_NIEN!$L$40)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,40,12))

def f197(E, r, c):
    # Ô D22
    # = IF(A22="","",LUU_NIEN!$H$40)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,40,8))

def f198(E, r, c):
    # Ô E22
    # = IF(A22="","",LUU_NIEN!$AA$40)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,40,27))

def f199(E, r, c):
    # Ô F22
    # = IF(A22="","",LUU_NIEN!$AL$40)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,40,38))

def f200(E, r, c):
    # Ô K22
    # = IF(A22="","",TRIM(IF(AB22,"Dấu hiệu bất lợi cho cha. ","")&IF(AC22,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA22,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U22>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$40))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,40,66))))) else ''))))

def f201(E, r, c):
    # Ô L22
    # = IF(OR(A22="",LUU_NIEN!$CN$40=""),"",IFERROR(LEFT(LUU_NIEN!$CN$40,FIND(":",LUU_NIEN!$CN$40)-1),LEFT(LUU_NIEN!$CN$40,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,40,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,40,92), X.sub(X.FIND(':', E.v(15,40,92)),1.0)), lambda: X.LEFT(E.v(15,40,92), 120.0)))

def f202(E, r, c):
    # Ô M22
    # = IF(A22="","",IF(LEN(LUU_NIEN!$BN$40)>260,LEFT(LUU_NIEN!$BN$40,260)&"…",LUU_NIEN!$BN$40))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,40,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,40,66)),260.0)) else E.v(15,40,66)))

def f203(E, r, c):
    # Ô T22
    # = IF(OR(A22="",LUU_NIEN!$P$40=""),-1,MOD(LUU_NIEN!$P$40-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,40,16),''))) else X.MOD(X.add(X.sub(E.v(15,40,16),E.v(23,7,5)),5.0), 5.0))

def f204(E, r, c):
    # Ô AA22
    # = IF(A22="",FALSE(),AND($U22>=18,$U22<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$40))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,40,66)))))

def f205(E, r, c):
    # Ô AB22
    # = IF(A22="",FALSE(),AND($U22>=0,$U22<=120,LUU_NIEN!$BH$40=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,40,60),1.0)))

def f206(E, r, c):
    # Ô AC22
    # = IF(A22="",FALSE(),AND($U22>=0,$U22<=120,LUU_NIEN!$BI$40=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,40,61),1.0)))

def f207(E, r, c):
    # Ô AD22
    # = IF(A22="",FALSE(),AND($U22>=0,$U22<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$40))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,40,92)))))

def f208(E, r, c):
    # Ô A23
    # = IF((ISNUMBER(LUU_NIEN!$A$41)),LUU_NIEN!$A$41,"")
    return (E.v(15,41,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,41,1))) else '')

def f209(E, r, c):
    # Ô C23
    # = IF(A23="","",LUU_NIEN!$L$41)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,41,12))

def f210(E, r, c):
    # Ô D23
    # = IF(A23="","",LUU_NIEN!$H$41)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,41,8))

def f211(E, r, c):
    # Ô E23
    # = IF(A23="","",LUU_NIEN!$AA$41)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,41,27))

def f212(E, r, c):
    # Ô F23
    # = IF(A23="","",LUU_NIEN!$AL$41)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,41,38))

def f213(E, r, c):
    # Ô K23
    # = IF(A23="","",TRIM(IF(AB23,"Dấu hiệu bất lợi cho cha. ","")&IF(AC23,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA23,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U23>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$41))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,41,66))))) else ''))))

def f214(E, r, c):
    # Ô L23
    # = IF(OR(A23="",LUU_NIEN!$CN$41=""),"",IFERROR(LEFT(LUU_NIEN!$CN$41,FIND(":",LUU_NIEN!$CN$41)-1),LEFT(LUU_NIEN!$CN$41,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,41,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,41,92), X.sub(X.FIND(':', E.v(15,41,92)),1.0)), lambda: X.LEFT(E.v(15,41,92), 120.0)))

def f215(E, r, c):
    # Ô M23
    # = IF(A23="","",IF(LEN(LUU_NIEN!$BN$41)>260,LEFT(LUU_NIEN!$BN$41,260)&"…",LUU_NIEN!$BN$41))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,41,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,41,66)),260.0)) else E.v(15,41,66)))

def f216(E, r, c):
    # Ô T23
    # = IF(OR(A23="",LUU_NIEN!$P$41=""),-1,MOD(LUU_NIEN!$P$41-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,41,16),''))) else X.MOD(X.add(X.sub(E.v(15,41,16),E.v(23,7,5)),5.0), 5.0))

def f217(E, r, c):
    # Ô AA23
    # = IF(A23="",FALSE(),AND($U23>=18,$U23<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$41))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,41,66)))))

def f218(E, r, c):
    # Ô AB23
    # = IF(A23="",FALSE(),AND($U23>=0,$U23<=120,LUU_NIEN!$BH$41=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,41,60),1.0)))

def f219(E, r, c):
    # Ô AC23
    # = IF(A23="",FALSE(),AND($U23>=0,$U23<=120,LUU_NIEN!$BI$41=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,41,61),1.0)))

def f220(E, r, c):
    # Ô AD23
    # = IF(A23="",FALSE(),AND($U23>=0,$U23<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$41))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,41,92)))))

def f221(E, r, c):
    # Ô A24
    # = IF((ISNUMBER(LUU_NIEN!$A$42)),LUU_NIEN!$A$42,"")
    return (E.v(15,42,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,42,1))) else '')

def f222(E, r, c):
    # Ô C24
    # = IF(A24="","",LUU_NIEN!$L$42)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,42,12))

def f223(E, r, c):
    # Ô D24
    # = IF(A24="","",LUU_NIEN!$H$42)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,42,8))

def f224(E, r, c):
    # Ô E24
    # = IF(A24="","",LUU_NIEN!$AA$42)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,42,27))

def f225(E, r, c):
    # Ô F24
    # = IF(A24="","",LUU_NIEN!$AL$42)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,42,38))

def f226(E, r, c):
    # Ô K24
    # = IF(A24="","",TRIM(IF(AB24,"Dấu hiệu bất lợi cho cha. ","")&IF(AC24,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA24,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U24>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$42))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,42,66))))) else ''))))

def f227(E, r, c):
    # Ô L24
    # = IF(OR(A24="",LUU_NIEN!$CN$42=""),"",IFERROR(LEFT(LUU_NIEN!$CN$42,FIND(":",LUU_NIEN!$CN$42)-1),LEFT(LUU_NIEN!$CN$42,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,42,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,42,92), X.sub(X.FIND(':', E.v(15,42,92)),1.0)), lambda: X.LEFT(E.v(15,42,92), 120.0)))

def f228(E, r, c):
    # Ô M24
    # = IF(A24="","",IF(LEN(LUU_NIEN!$BN$42)>260,LEFT(LUU_NIEN!$BN$42,260)&"…",LUU_NIEN!$BN$42))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,42,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,42,66)),260.0)) else E.v(15,42,66)))

def f229(E, r, c):
    # Ô T24
    # = IF(OR(A24="",LUU_NIEN!$P$42=""),-1,MOD(LUU_NIEN!$P$42-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,42,16),''))) else X.MOD(X.add(X.sub(E.v(15,42,16),E.v(23,7,5)),5.0), 5.0))

def f230(E, r, c):
    # Ô AA24
    # = IF(A24="",FALSE(),AND($U24>=18,$U24<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$42))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,42,66)))))

def f231(E, r, c):
    # Ô AB24
    # = IF(A24="",FALSE(),AND($U24>=0,$U24<=120,LUU_NIEN!$BH$42=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,42,60),1.0)))

def f232(E, r, c):
    # Ô AC24
    # = IF(A24="",FALSE(),AND($U24>=0,$U24<=120,LUU_NIEN!$BI$42=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,42,61),1.0)))

def f233(E, r, c):
    # Ô AD24
    # = IF(A24="",FALSE(),AND($U24>=0,$U24<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$42))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,42,92)))))

def f234(E, r, c):
    # Ô A25
    # = IF((ISNUMBER(LUU_NIEN!$A$43)),LUU_NIEN!$A$43,"")
    return (E.v(15,43,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,43,1))) else '')

def f235(E, r, c):
    # Ô C25
    # = IF(A25="","",LUU_NIEN!$L$43)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,43,12))

def f236(E, r, c):
    # Ô D25
    # = IF(A25="","",LUU_NIEN!$H$43)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,43,8))

def f237(E, r, c):
    # Ô E25
    # = IF(A25="","",LUU_NIEN!$AA$43)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,43,27))

def f238(E, r, c):
    # Ô F25
    # = IF(A25="","",LUU_NIEN!$AL$43)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,43,38))

def f239(E, r, c):
    # Ô K25
    # = IF(A25="","",TRIM(IF(AB25,"Dấu hiệu bất lợi cho cha. ","")&IF(AC25,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA25,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U25>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$43))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,43,66))))) else ''))))

def f240(E, r, c):
    # Ô L25
    # = IF(OR(A25="",LUU_NIEN!$CN$43=""),"",IFERROR(LEFT(LUU_NIEN!$CN$43,FIND(":",LUU_NIEN!$CN$43)-1),LEFT(LUU_NIEN!$CN$43,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,43,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,43,92), X.sub(X.FIND(':', E.v(15,43,92)),1.0)), lambda: X.LEFT(E.v(15,43,92), 120.0)))

def f241(E, r, c):
    # Ô M25
    # = IF(A25="","",IF(LEN(LUU_NIEN!$BN$43)>260,LEFT(LUU_NIEN!$BN$43,260)&"…",LUU_NIEN!$BN$43))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,43,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,43,66)),260.0)) else E.v(15,43,66)))

def f242(E, r, c):
    # Ô T25
    # = IF(OR(A25="",LUU_NIEN!$P$43=""),-1,MOD(LUU_NIEN!$P$43-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,43,16),''))) else X.MOD(X.add(X.sub(E.v(15,43,16),E.v(23,7,5)),5.0), 5.0))

def f243(E, r, c):
    # Ô AA25
    # = IF(A25="",FALSE(),AND($U25>=18,$U25<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$43))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,43,66)))))

def f244(E, r, c):
    # Ô AB25
    # = IF(A25="",FALSE(),AND($U25>=0,$U25<=120,LUU_NIEN!$BH$43=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,43,60),1.0)))

def f245(E, r, c):
    # Ô AC25
    # = IF(A25="",FALSE(),AND($U25>=0,$U25<=120,LUU_NIEN!$BI$43=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,43,61),1.0)))

def f246(E, r, c):
    # Ô AD25
    # = IF(A25="",FALSE(),AND($U25>=0,$U25<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$43))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,43,92)))))

def f247(E, r, c):
    # Ô A26
    # = IF((ISNUMBER(LUU_NIEN!$A$44)),LUU_NIEN!$A$44,"")
    return (E.v(15,44,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,44,1))) else '')

def f248(E, r, c):
    # Ô C26
    # = IF(A26="","",LUU_NIEN!$L$44)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,44,12))

def f249(E, r, c):
    # Ô D26
    # = IF(A26="","",LUU_NIEN!$H$44)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,44,8))

def f250(E, r, c):
    # Ô E26
    # = IF(A26="","",LUU_NIEN!$AA$44)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,44,27))

def f251(E, r, c):
    # Ô F26
    # = IF(A26="","",LUU_NIEN!$AL$44)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,44,38))

def f252(E, r, c):
    # Ô K26
    # = IF(A26="","",TRIM(IF(AB26,"Dấu hiệu bất lợi cho cha. ","")&IF(AC26,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA26,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U26>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$44))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,44,66))))) else ''))))

def f253(E, r, c):
    # Ô L26
    # = IF(OR(A26="",LUU_NIEN!$CN$44=""),"",IFERROR(LEFT(LUU_NIEN!$CN$44,FIND(":",LUU_NIEN!$CN$44)-1),LEFT(LUU_NIEN!$CN$44,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,44,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,44,92), X.sub(X.FIND(':', E.v(15,44,92)),1.0)), lambda: X.LEFT(E.v(15,44,92), 120.0)))

def f254(E, r, c):
    # Ô M26
    # = IF(A26="","",IF(LEN(LUU_NIEN!$BN$44)>260,LEFT(LUU_NIEN!$BN$44,260)&"…",LUU_NIEN!$BN$44))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,44,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,44,66)),260.0)) else E.v(15,44,66)))

def f255(E, r, c):
    # Ô T26
    # = IF(OR(A26="",LUU_NIEN!$P$44=""),-1,MOD(LUU_NIEN!$P$44-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,44,16),''))) else X.MOD(X.add(X.sub(E.v(15,44,16),E.v(23,7,5)),5.0), 5.0))

def f256(E, r, c):
    # Ô AA26
    # = IF(A26="",FALSE(),AND($U26>=18,$U26<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$44))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,44,66)))))

def f257(E, r, c):
    # Ô AB26
    # = IF(A26="",FALSE(),AND($U26>=0,$U26<=120,LUU_NIEN!$BH$44=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,44,60),1.0)))

def f258(E, r, c):
    # Ô AC26
    # = IF(A26="",FALSE(),AND($U26>=0,$U26<=120,LUU_NIEN!$BI$44=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,44,61),1.0)))

def f259(E, r, c):
    # Ô AD26
    # = IF(A26="",FALSE(),AND($U26>=0,$U26<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$44))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,44,92)))))

def f260(E, r, c):
    # Ô A27
    # = IF((ISNUMBER(LUU_NIEN!$A$45)),LUU_NIEN!$A$45,"")
    return (E.v(15,45,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,45,1))) else '')

def f261(E, r, c):
    # Ô C27
    # = IF(A27="","",LUU_NIEN!$L$45)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,45,12))

def f262(E, r, c):
    # Ô D27
    # = IF(A27="","",LUU_NIEN!$H$45)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,45,8))

def f263(E, r, c):
    # Ô E27
    # = IF(A27="","",LUU_NIEN!$AA$45)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,45,27))

def f264(E, r, c):
    # Ô F27
    # = IF(A27="","",LUU_NIEN!$AL$45)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,45,38))

def f265(E, r, c):
    # Ô K27
    # = IF(A27="","",TRIM(IF(AB27,"Dấu hiệu bất lợi cho cha. ","")&IF(AC27,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA27,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U27>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$45))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,45,66))))) else ''))))

def f266(E, r, c):
    # Ô L27
    # = IF(OR(A27="",LUU_NIEN!$CN$45=""),"",IFERROR(LEFT(LUU_NIEN!$CN$45,FIND(":",LUU_NIEN!$CN$45)-1),LEFT(LUU_NIEN!$CN$45,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,45,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,45,92), X.sub(X.FIND(':', E.v(15,45,92)),1.0)), lambda: X.LEFT(E.v(15,45,92), 120.0)))

def f267(E, r, c):
    # Ô M27
    # = IF(A27="","",IF(LEN(LUU_NIEN!$BN$45)>260,LEFT(LUU_NIEN!$BN$45,260)&"…",LUU_NIEN!$BN$45))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,45,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,45,66)),260.0)) else E.v(15,45,66)))

def f268(E, r, c):
    # Ô T27
    # = IF(OR(A27="",LUU_NIEN!$P$45=""),-1,MOD(LUU_NIEN!$P$45-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,45,16),''))) else X.MOD(X.add(X.sub(E.v(15,45,16),E.v(23,7,5)),5.0), 5.0))

def f269(E, r, c):
    # Ô AA27
    # = IF(A27="",FALSE(),AND($U27>=18,$U27<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$45))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,45,66)))))

def f270(E, r, c):
    # Ô AB27
    # = IF(A27="",FALSE(),AND($U27>=0,$U27<=120,LUU_NIEN!$BH$45=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,45,60),1.0)))

def f271(E, r, c):
    # Ô AC27
    # = IF(A27="",FALSE(),AND($U27>=0,$U27<=120,LUU_NIEN!$BI$45=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,45,61),1.0)))

def f272(E, r, c):
    # Ô AD27
    # = IF(A27="",FALSE(),AND($U27>=0,$U27<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$45))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,45,92)))))

def f273(E, r, c):
    # Ô A28
    # = IF((ISNUMBER(LUU_NIEN!$A$46)),LUU_NIEN!$A$46,"")
    return (E.v(15,46,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,46,1))) else '')

def f274(E, r, c):
    # Ô C28
    # = IF(A28="","",LUU_NIEN!$L$46)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,46,12))

def f275(E, r, c):
    # Ô D28
    # = IF(A28="","",LUU_NIEN!$H$46)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,46,8))

def f276(E, r, c):
    # Ô E28
    # = IF(A28="","",LUU_NIEN!$AA$46)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,46,27))

def f277(E, r, c):
    # Ô F28
    # = IF(A28="","",LUU_NIEN!$AL$46)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,46,38))

def f278(E, r, c):
    # Ô K28
    # = IF(A28="","",TRIM(IF(AB28,"Dấu hiệu bất lợi cho cha. ","")&IF(AC28,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA28,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U28>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$46))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,46,66))))) else ''))))

def f279(E, r, c):
    # Ô L28
    # = IF(OR(A28="",LUU_NIEN!$CN$46=""),"",IFERROR(LEFT(LUU_NIEN!$CN$46,FIND(":",LUU_NIEN!$CN$46)-1),LEFT(LUU_NIEN!$CN$46,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,46,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,46,92), X.sub(X.FIND(':', E.v(15,46,92)),1.0)), lambda: X.LEFT(E.v(15,46,92), 120.0)))

def f280(E, r, c):
    # Ô M28
    # = IF(A28="","",IF(LEN(LUU_NIEN!$BN$46)>260,LEFT(LUU_NIEN!$BN$46,260)&"…",LUU_NIEN!$BN$46))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,46,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,46,66)),260.0)) else E.v(15,46,66)))

def f281(E, r, c):
    # Ô T28
    # = IF(OR(A28="",LUU_NIEN!$P$46=""),-1,MOD(LUU_NIEN!$P$46-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,46,16),''))) else X.MOD(X.add(X.sub(E.v(15,46,16),E.v(23,7,5)),5.0), 5.0))

def f282(E, r, c):
    # Ô AA28
    # = IF(A28="",FALSE(),AND($U28>=18,$U28<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$46))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,46,66)))))

def f283(E, r, c):
    # Ô AB28
    # = IF(A28="",FALSE(),AND($U28>=0,$U28<=120,LUU_NIEN!$BH$46=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,46,60),1.0)))

def f284(E, r, c):
    # Ô AC28
    # = IF(A28="",FALSE(),AND($U28>=0,$U28<=120,LUU_NIEN!$BI$46=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,46,61),1.0)))

def f285(E, r, c):
    # Ô AD28
    # = IF(A28="",FALSE(),AND($U28>=0,$U28<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$46))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,46,92)))))

def f286(E, r, c):
    # Ô A29
    # = IF((ISNUMBER(LUU_NIEN!$A$47)),LUU_NIEN!$A$47,"")
    return (E.v(15,47,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,47,1))) else '')

def f287(E, r, c):
    # Ô C29
    # = IF(A29="","",LUU_NIEN!$L$47)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,47,12))

def f288(E, r, c):
    # Ô D29
    # = IF(A29="","",LUU_NIEN!$H$47)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,47,8))

def f289(E, r, c):
    # Ô E29
    # = IF(A29="","",LUU_NIEN!$AA$47)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,47,27))

def f290(E, r, c):
    # Ô F29
    # = IF(A29="","",LUU_NIEN!$AL$47)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,47,38))

def f291(E, r, c):
    # Ô K29
    # = IF(A29="","",TRIM(IF(AB29,"Dấu hiệu bất lợi cho cha. ","")&IF(AC29,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA29,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U29>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$47))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,47,66))))) else ''))))

def f292(E, r, c):
    # Ô L29
    # = IF(OR(A29="",LUU_NIEN!$CN$47=""),"",IFERROR(LEFT(LUU_NIEN!$CN$47,FIND(":",LUU_NIEN!$CN$47)-1),LEFT(LUU_NIEN!$CN$47,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,47,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,47,92), X.sub(X.FIND(':', E.v(15,47,92)),1.0)), lambda: X.LEFT(E.v(15,47,92), 120.0)))

def f293(E, r, c):
    # Ô M29
    # = IF(A29="","",IF(LEN(LUU_NIEN!$BN$47)>260,LEFT(LUU_NIEN!$BN$47,260)&"…",LUU_NIEN!$BN$47))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,47,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,47,66)),260.0)) else E.v(15,47,66)))

def f294(E, r, c):
    # Ô T29
    # = IF(OR(A29="",LUU_NIEN!$P$47=""),-1,MOD(LUU_NIEN!$P$47-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,47,16),''))) else X.MOD(X.add(X.sub(E.v(15,47,16),E.v(23,7,5)),5.0), 5.0))

def f295(E, r, c):
    # Ô AA29
    # = IF(A29="",FALSE(),AND($U29>=18,$U29<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$47))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,47,66)))))

def f296(E, r, c):
    # Ô AB29
    # = IF(A29="",FALSE(),AND($U29>=0,$U29<=120,LUU_NIEN!$BH$47=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,47,60),1.0)))

def f297(E, r, c):
    # Ô AC29
    # = IF(A29="",FALSE(),AND($U29>=0,$U29<=120,LUU_NIEN!$BI$47=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,47,61),1.0)))

def f298(E, r, c):
    # Ô AD29
    # = IF(A29="",FALSE(),AND($U29>=0,$U29<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$47))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,47,92)))))

def f299(E, r, c):
    # Ô A30
    # = IF((ISNUMBER(LUU_NIEN!$A$48)),LUU_NIEN!$A$48,"")
    return (E.v(15,48,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,48,1))) else '')

def f300(E, r, c):
    # Ô C30
    # = IF(A30="","",LUU_NIEN!$L$48)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,48,12))

def f301(E, r, c):
    # Ô D30
    # = IF(A30="","",LUU_NIEN!$H$48)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,48,8))

def f302(E, r, c):
    # Ô E30
    # = IF(A30="","",LUU_NIEN!$AA$48)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,48,27))

def f303(E, r, c):
    # Ô F30
    # = IF(A30="","",LUU_NIEN!$AL$48)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,48,38))

def f304(E, r, c):
    # Ô K30
    # = IF(A30="","",TRIM(IF(AB30,"Dấu hiệu bất lợi cho cha. ","")&IF(AC30,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA30,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U30>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$48))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,48,66))))) else ''))))

def f305(E, r, c):
    # Ô L30
    # = IF(OR(A30="",LUU_NIEN!$CN$48=""),"",IFERROR(LEFT(LUU_NIEN!$CN$48,FIND(":",LUU_NIEN!$CN$48)-1),LEFT(LUU_NIEN!$CN$48,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,48,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,48,92), X.sub(X.FIND(':', E.v(15,48,92)),1.0)), lambda: X.LEFT(E.v(15,48,92), 120.0)))

def f306(E, r, c):
    # Ô M30
    # = IF(A30="","",IF(LEN(LUU_NIEN!$BN$48)>260,LEFT(LUU_NIEN!$BN$48,260)&"…",LUU_NIEN!$BN$48))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,48,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,48,66)),260.0)) else E.v(15,48,66)))

def f307(E, r, c):
    # Ô T30
    # = IF(OR(A30="",LUU_NIEN!$P$48=""),-1,MOD(LUU_NIEN!$P$48-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,48,16),''))) else X.MOD(X.add(X.sub(E.v(15,48,16),E.v(23,7,5)),5.0), 5.0))

def f308(E, r, c):
    # Ô AA30
    # = IF(A30="",FALSE(),AND($U30>=18,$U30<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$48))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,48,66)))))

def f309(E, r, c):
    # Ô AB30
    # = IF(A30="",FALSE(),AND($U30>=0,$U30<=120,LUU_NIEN!$BH$48=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,48,60),1.0)))

def f310(E, r, c):
    # Ô AC30
    # = IF(A30="",FALSE(),AND($U30>=0,$U30<=120,LUU_NIEN!$BI$48=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,48,61),1.0)))

def f311(E, r, c):
    # Ô AD30
    # = IF(A30="",FALSE(),AND($U30>=0,$U30<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$48))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,48,92)))))

def f312(E, r, c):
    # Ô A31
    # = IF((ISNUMBER(LUU_NIEN!$A$49)),LUU_NIEN!$A$49,"")
    return (E.v(15,49,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,49,1))) else '')

def f313(E, r, c):
    # Ô C31
    # = IF(A31="","",LUU_NIEN!$L$49)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,49,12))

def f314(E, r, c):
    # Ô D31
    # = IF(A31="","",LUU_NIEN!$H$49)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,49,8))

def f315(E, r, c):
    # Ô E31
    # = IF(A31="","",LUU_NIEN!$AA$49)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,49,27))

def f316(E, r, c):
    # Ô F31
    # = IF(A31="","",LUU_NIEN!$AL$49)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,49,38))

def f317(E, r, c):
    # Ô K31
    # = IF(A31="","",TRIM(IF(AB31,"Dấu hiệu bất lợi cho cha. ","")&IF(AC31,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA31,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U31>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$49))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,49,66))))) else ''))))

def f318(E, r, c):
    # Ô L31
    # = IF(OR(A31="",LUU_NIEN!$CN$49=""),"",IFERROR(LEFT(LUU_NIEN!$CN$49,FIND(":",LUU_NIEN!$CN$49)-1),LEFT(LUU_NIEN!$CN$49,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,49,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,49,92), X.sub(X.FIND(':', E.v(15,49,92)),1.0)), lambda: X.LEFT(E.v(15,49,92), 120.0)))

def f319(E, r, c):
    # Ô M31
    # = IF(A31="","",IF(LEN(LUU_NIEN!$BN$49)>260,LEFT(LUU_NIEN!$BN$49,260)&"…",LUU_NIEN!$BN$49))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,49,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,49,66)),260.0)) else E.v(15,49,66)))

def f320(E, r, c):
    # Ô T31
    # = IF(OR(A31="",LUU_NIEN!$P$49=""),-1,MOD(LUU_NIEN!$P$49-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,49,16),''))) else X.MOD(X.add(X.sub(E.v(15,49,16),E.v(23,7,5)),5.0), 5.0))

def f321(E, r, c):
    # Ô AA31
    # = IF(A31="",FALSE(),AND($U31>=18,$U31<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$49))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,49,66)))))

def f322(E, r, c):
    # Ô AB31
    # = IF(A31="",FALSE(),AND($U31>=0,$U31<=120,LUU_NIEN!$BH$49=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,49,60),1.0)))

def f323(E, r, c):
    # Ô AC31
    # = IF(A31="",FALSE(),AND($U31>=0,$U31<=120,LUU_NIEN!$BI$49=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,49,61),1.0)))

def f324(E, r, c):
    # Ô AD31
    # = IF(A31="",FALSE(),AND($U31>=0,$U31<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$49))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,49,92)))))

def f325(E, r, c):
    # Ô A32
    # = IF((ISNUMBER(LUU_NIEN!$A$50)),LUU_NIEN!$A$50,"")
    return (E.v(15,50,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,50,1))) else '')

def f326(E, r, c):
    # Ô C32
    # = IF(A32="","",LUU_NIEN!$L$50)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,50,12))

def f327(E, r, c):
    # Ô D32
    # = IF(A32="","",LUU_NIEN!$H$50)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,50,8))

def f328(E, r, c):
    # Ô E32
    # = IF(A32="","",LUU_NIEN!$AA$50)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,50,27))

def f329(E, r, c):
    # Ô F32
    # = IF(A32="","",LUU_NIEN!$AL$50)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,50,38))

def f330(E, r, c):
    # Ô K32
    # = IF(A32="","",TRIM(IF(AB32,"Dấu hiệu bất lợi cho cha. ","")&IF(AC32,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA32,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U32>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$50))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,50,66))))) else ''))))

def f331(E, r, c):
    # Ô L32
    # = IF(OR(A32="",LUU_NIEN!$CN$50=""),"",IFERROR(LEFT(LUU_NIEN!$CN$50,FIND(":",LUU_NIEN!$CN$50)-1),LEFT(LUU_NIEN!$CN$50,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,50,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,50,92), X.sub(X.FIND(':', E.v(15,50,92)),1.0)), lambda: X.LEFT(E.v(15,50,92), 120.0)))

def f332(E, r, c):
    # Ô M32
    # = IF(A32="","",IF(LEN(LUU_NIEN!$BN$50)>260,LEFT(LUU_NIEN!$BN$50,260)&"…",LUU_NIEN!$BN$50))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,50,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,50,66)),260.0)) else E.v(15,50,66)))

def f333(E, r, c):
    # Ô T32
    # = IF(OR(A32="",LUU_NIEN!$P$50=""),-1,MOD(LUU_NIEN!$P$50-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,50,16),''))) else X.MOD(X.add(X.sub(E.v(15,50,16),E.v(23,7,5)),5.0), 5.0))

def f334(E, r, c):
    # Ô AA32
    # = IF(A32="",FALSE(),AND($U32>=18,$U32<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$50))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,50,66)))))

def f335(E, r, c):
    # Ô AB32
    # = IF(A32="",FALSE(),AND($U32>=0,$U32<=120,LUU_NIEN!$BH$50=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,50,60),1.0)))

def f336(E, r, c):
    # Ô AC32
    # = IF(A32="",FALSE(),AND($U32>=0,$U32<=120,LUU_NIEN!$BI$50=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,50,61),1.0)))

def f337(E, r, c):
    # Ô AD32
    # = IF(A32="",FALSE(),AND($U32>=0,$U32<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$50))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,50,92)))))

def f338(E, r, c):
    # Ô A33
    # = IF((ISNUMBER(LUU_NIEN!$A$51)),LUU_NIEN!$A$51,"")
    return (E.v(15,51,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,51,1))) else '')

def f339(E, r, c):
    # Ô C33
    # = IF(A33="","",LUU_NIEN!$L$51)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,51,12))

def f340(E, r, c):
    # Ô D33
    # = IF(A33="","",LUU_NIEN!$H$51)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,51,8))

def f341(E, r, c):
    # Ô E33
    # = IF(A33="","",LUU_NIEN!$AA$51)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,51,27))

def f342(E, r, c):
    # Ô F33
    # = IF(A33="","",LUU_NIEN!$AL$51)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,51,38))

def f343(E, r, c):
    # Ô K33
    # = IF(A33="","",TRIM(IF(AB33,"Dấu hiệu bất lợi cho cha. ","")&IF(AC33,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA33,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U33>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$51))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,51,66))))) else ''))))

def f344(E, r, c):
    # Ô L33
    # = IF(OR(A33="",LUU_NIEN!$CN$51=""),"",IFERROR(LEFT(LUU_NIEN!$CN$51,FIND(":",LUU_NIEN!$CN$51)-1),LEFT(LUU_NIEN!$CN$51,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,51,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,51,92), X.sub(X.FIND(':', E.v(15,51,92)),1.0)), lambda: X.LEFT(E.v(15,51,92), 120.0)))

def f345(E, r, c):
    # Ô M33
    # = IF(A33="","",IF(LEN(LUU_NIEN!$BN$51)>260,LEFT(LUU_NIEN!$BN$51,260)&"…",LUU_NIEN!$BN$51))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,51,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,51,66)),260.0)) else E.v(15,51,66)))

def f346(E, r, c):
    # Ô T33
    # = IF(OR(A33="",LUU_NIEN!$P$51=""),-1,MOD(LUU_NIEN!$P$51-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,51,16),''))) else X.MOD(X.add(X.sub(E.v(15,51,16),E.v(23,7,5)),5.0), 5.0))

def f347(E, r, c):
    # Ô AA33
    # = IF(A33="",FALSE(),AND($U33>=18,$U33<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$51))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,51,66)))))

def f348(E, r, c):
    # Ô AB33
    # = IF(A33="",FALSE(),AND($U33>=0,$U33<=120,LUU_NIEN!$BH$51=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,51,60),1.0)))

def f349(E, r, c):
    # Ô AC33
    # = IF(A33="",FALSE(),AND($U33>=0,$U33<=120,LUU_NIEN!$BI$51=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,51,61),1.0)))

def f350(E, r, c):
    # Ô AD33
    # = IF(A33="",FALSE(),AND($U33>=0,$U33<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$51))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,51,92)))))

def f351(E, r, c):
    # Ô A34
    # = IF((ISNUMBER(LUU_NIEN!$A$52)),LUU_NIEN!$A$52,"")
    return (E.v(15,52,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,52,1))) else '')

def f352(E, r, c):
    # Ô C34
    # = IF(A34="","",LUU_NIEN!$L$52)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,52,12))

def f353(E, r, c):
    # Ô D34
    # = IF(A34="","",LUU_NIEN!$H$52)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,52,8))

def f354(E, r, c):
    # Ô E34
    # = IF(A34="","",LUU_NIEN!$AA$52)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,52,27))

def f355(E, r, c):
    # Ô F34
    # = IF(A34="","",LUU_NIEN!$AL$52)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,52,38))

def f356(E, r, c):
    # Ô K34
    # = IF(A34="","",TRIM(IF(AB34,"Dấu hiệu bất lợi cho cha. ","")&IF(AC34,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA34,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U34>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$52))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,52,66))))) else ''))))

def f357(E, r, c):
    # Ô L34
    # = IF(OR(A34="",LUU_NIEN!$CN$52=""),"",IFERROR(LEFT(LUU_NIEN!$CN$52,FIND(":",LUU_NIEN!$CN$52)-1),LEFT(LUU_NIEN!$CN$52,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,52,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,52,92), X.sub(X.FIND(':', E.v(15,52,92)),1.0)), lambda: X.LEFT(E.v(15,52,92), 120.0)))

def f358(E, r, c):
    # Ô M34
    # = IF(A34="","",IF(LEN(LUU_NIEN!$BN$52)>260,LEFT(LUU_NIEN!$BN$52,260)&"…",LUU_NIEN!$BN$52))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,52,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,52,66)),260.0)) else E.v(15,52,66)))

def f359(E, r, c):
    # Ô T34
    # = IF(OR(A34="",LUU_NIEN!$P$52=""),-1,MOD(LUU_NIEN!$P$52-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,52,16),''))) else X.MOD(X.add(X.sub(E.v(15,52,16),E.v(23,7,5)),5.0), 5.0))

def f360(E, r, c):
    # Ô AA34
    # = IF(A34="",FALSE(),AND($U34>=18,$U34<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$52))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,52,66)))))

def f361(E, r, c):
    # Ô AB34
    # = IF(A34="",FALSE(),AND($U34>=0,$U34<=120,LUU_NIEN!$BH$52=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,52,60),1.0)))

def f362(E, r, c):
    # Ô AC34
    # = IF(A34="",FALSE(),AND($U34>=0,$U34<=120,LUU_NIEN!$BI$52=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,52,61),1.0)))

def f363(E, r, c):
    # Ô AD34
    # = IF(A34="",FALSE(),AND($U34>=0,$U34<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$52))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,52,92)))))

def f364(E, r, c):
    # Ô A35
    # = IF((ISNUMBER(LUU_NIEN!$A$53)),LUU_NIEN!$A$53,"")
    return (E.v(15,53,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,53,1))) else '')

def f365(E, r, c):
    # Ô C35
    # = IF(A35="","",LUU_NIEN!$L$53)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,53,12))

def f366(E, r, c):
    # Ô D35
    # = IF(A35="","",LUU_NIEN!$H$53)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,53,8))

def f367(E, r, c):
    # Ô E35
    # = IF(A35="","",LUU_NIEN!$AA$53)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,53,27))

def f368(E, r, c):
    # Ô F35
    # = IF(A35="","",LUU_NIEN!$AL$53)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,53,38))

def f369(E, r, c):
    # Ô K35
    # = IF(A35="","",TRIM(IF(AB35,"Dấu hiệu bất lợi cho cha. ","")&IF(AC35,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA35,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U35>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$53))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,53,66))))) else ''))))

def f370(E, r, c):
    # Ô L35
    # = IF(OR(A35="",LUU_NIEN!$CN$53=""),"",IFERROR(LEFT(LUU_NIEN!$CN$53,FIND(":",LUU_NIEN!$CN$53)-1),LEFT(LUU_NIEN!$CN$53,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,53,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,53,92), X.sub(X.FIND(':', E.v(15,53,92)),1.0)), lambda: X.LEFT(E.v(15,53,92), 120.0)))

def f371(E, r, c):
    # Ô M35
    # = IF(A35="","",IF(LEN(LUU_NIEN!$BN$53)>260,LEFT(LUU_NIEN!$BN$53,260)&"…",LUU_NIEN!$BN$53))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,53,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,53,66)),260.0)) else E.v(15,53,66)))

def f372(E, r, c):
    # Ô T35
    # = IF(OR(A35="",LUU_NIEN!$P$53=""),-1,MOD(LUU_NIEN!$P$53-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,53,16),''))) else X.MOD(X.add(X.sub(E.v(15,53,16),E.v(23,7,5)),5.0), 5.0))

def f373(E, r, c):
    # Ô AA35
    # = IF(A35="",FALSE(),AND($U35>=18,$U35<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$53))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,53,66)))))

def f374(E, r, c):
    # Ô AB35
    # = IF(A35="",FALSE(),AND($U35>=0,$U35<=120,LUU_NIEN!$BH$53=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,53,60),1.0)))

def f375(E, r, c):
    # Ô AC35
    # = IF(A35="",FALSE(),AND($U35>=0,$U35<=120,LUU_NIEN!$BI$53=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,53,61),1.0)))

def f376(E, r, c):
    # Ô AD35
    # = IF(A35="",FALSE(),AND($U35>=0,$U35<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$53))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,53,92)))))

def f377(E, r, c):
    # Ô A36
    # = IF((ISNUMBER(LUU_NIEN!$A$54)),LUU_NIEN!$A$54,"")
    return (E.v(15,54,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,54,1))) else '')

def f378(E, r, c):
    # Ô C36
    # = IF(A36="","",LUU_NIEN!$L$54)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,54,12))

def f379(E, r, c):
    # Ô D36
    # = IF(A36="","",LUU_NIEN!$H$54)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,54,8))

def f380(E, r, c):
    # Ô E36
    # = IF(A36="","",LUU_NIEN!$AA$54)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,54,27))

def f381(E, r, c):
    # Ô F36
    # = IF(A36="","",LUU_NIEN!$AL$54)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,54,38))

def f382(E, r, c):
    # Ô K36
    # = IF(A36="","",TRIM(IF(AB36,"Dấu hiệu bất lợi cho cha. ","")&IF(AC36,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA36,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U36>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$54))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,54,66))))) else ''))))

def f383(E, r, c):
    # Ô L36
    # = IF(OR(A36="",LUU_NIEN!$CN$54=""),"",IFERROR(LEFT(LUU_NIEN!$CN$54,FIND(":",LUU_NIEN!$CN$54)-1),LEFT(LUU_NIEN!$CN$54,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,54,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,54,92), X.sub(X.FIND(':', E.v(15,54,92)),1.0)), lambda: X.LEFT(E.v(15,54,92), 120.0)))

def f384(E, r, c):
    # Ô M36
    # = IF(A36="","",IF(LEN(LUU_NIEN!$BN$54)>260,LEFT(LUU_NIEN!$BN$54,260)&"…",LUU_NIEN!$BN$54))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,54,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,54,66)),260.0)) else E.v(15,54,66)))

def f385(E, r, c):
    # Ô T36
    # = IF(OR(A36="",LUU_NIEN!$P$54=""),-1,MOD(LUU_NIEN!$P$54-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,54,16),''))) else X.MOD(X.add(X.sub(E.v(15,54,16),E.v(23,7,5)),5.0), 5.0))

def f386(E, r, c):
    # Ô AA36
    # = IF(A36="",FALSE(),AND($U36>=18,$U36<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$54))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,54,66)))))

def f387(E, r, c):
    # Ô AB36
    # = IF(A36="",FALSE(),AND($U36>=0,$U36<=120,LUU_NIEN!$BH$54=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,54,60),1.0)))

def f388(E, r, c):
    # Ô AC36
    # = IF(A36="",FALSE(),AND($U36>=0,$U36<=120,LUU_NIEN!$BI$54=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,54,61),1.0)))

def f389(E, r, c):
    # Ô AD36
    # = IF(A36="",FALSE(),AND($U36>=0,$U36<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$54))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,54,92)))))

def f390(E, r, c):
    # Ô A37
    # = IF((ISNUMBER(LUU_NIEN!$A$55)),LUU_NIEN!$A$55,"")
    return (E.v(15,55,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,55,1))) else '')

def f391(E, r, c):
    # Ô C37
    # = IF(A37="","",LUU_NIEN!$L$55)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,55,12))

def f392(E, r, c):
    # Ô D37
    # = IF(A37="","",LUU_NIEN!$H$55)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,55,8))

def f393(E, r, c):
    # Ô E37
    # = IF(A37="","",LUU_NIEN!$AA$55)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,55,27))

def f394(E, r, c):
    # Ô F37
    # = IF(A37="","",LUU_NIEN!$AL$55)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,55,38))

def f395(E, r, c):
    # Ô K37
    # = IF(A37="","",TRIM(IF(AB37,"Dấu hiệu bất lợi cho cha. ","")&IF(AC37,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA37,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U37>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$55))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,55,66))))) else ''))))

def f396(E, r, c):
    # Ô L37
    # = IF(OR(A37="",LUU_NIEN!$CN$55=""),"",IFERROR(LEFT(LUU_NIEN!$CN$55,FIND(":",LUU_NIEN!$CN$55)-1),LEFT(LUU_NIEN!$CN$55,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,55,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,55,92), X.sub(X.FIND(':', E.v(15,55,92)),1.0)), lambda: X.LEFT(E.v(15,55,92), 120.0)))

def f397(E, r, c):
    # Ô M37
    # = IF(A37="","",IF(LEN(LUU_NIEN!$BN$55)>260,LEFT(LUU_NIEN!$BN$55,260)&"…",LUU_NIEN!$BN$55))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,55,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,55,66)),260.0)) else E.v(15,55,66)))

def f398(E, r, c):
    # Ô T37
    # = IF(OR(A37="",LUU_NIEN!$P$55=""),-1,MOD(LUU_NIEN!$P$55-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,55,16),''))) else X.MOD(X.add(X.sub(E.v(15,55,16),E.v(23,7,5)),5.0), 5.0))

def f399(E, r, c):
    # Ô AA37
    # = IF(A37="",FALSE(),AND($U37>=18,$U37<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$55))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,55,66)))))

def f400(E, r, c):
    # Ô AB37
    # = IF(A37="",FALSE(),AND($U37>=0,$U37<=120,LUU_NIEN!$BH$55=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,55,60),1.0)))

def f401(E, r, c):
    # Ô AC37
    # = IF(A37="",FALSE(),AND($U37>=0,$U37<=120,LUU_NIEN!$BI$55=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,55,61),1.0)))

def f402(E, r, c):
    # Ô AD37
    # = IF(A37="",FALSE(),AND($U37>=0,$U37<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$55))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,55,92)))))

def f403(E, r, c):
    # Ô A38
    # = IF((ISNUMBER(LUU_NIEN!$A$56)),LUU_NIEN!$A$56,"")
    return (E.v(15,56,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,56,1))) else '')

def f404(E, r, c):
    # Ô C38
    # = IF(A38="","",LUU_NIEN!$L$56)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,56,12))

def f405(E, r, c):
    # Ô D38
    # = IF(A38="","",LUU_NIEN!$H$56)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,56,8))

def f406(E, r, c):
    # Ô E38
    # = IF(A38="","",LUU_NIEN!$AA$56)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,56,27))

def f407(E, r, c):
    # Ô F38
    # = IF(A38="","",LUU_NIEN!$AL$56)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,56,38))

def f408(E, r, c):
    # Ô K38
    # = IF(A38="","",TRIM(IF(AB38,"Dấu hiệu bất lợi cho cha. ","")&IF(AC38,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA38,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U38>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$56))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,56,66))))) else ''))))

def f409(E, r, c):
    # Ô L38
    # = IF(OR(A38="",LUU_NIEN!$CN$56=""),"",IFERROR(LEFT(LUU_NIEN!$CN$56,FIND(":",LUU_NIEN!$CN$56)-1),LEFT(LUU_NIEN!$CN$56,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,56,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,56,92), X.sub(X.FIND(':', E.v(15,56,92)),1.0)), lambda: X.LEFT(E.v(15,56,92), 120.0)))

def f410(E, r, c):
    # Ô M38
    # = IF(A38="","",IF(LEN(LUU_NIEN!$BN$56)>260,LEFT(LUU_NIEN!$BN$56,260)&"…",LUU_NIEN!$BN$56))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,56,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,56,66)),260.0)) else E.v(15,56,66)))

def f411(E, r, c):
    # Ô T38
    # = IF(OR(A38="",LUU_NIEN!$P$56=""),-1,MOD(LUU_NIEN!$P$56-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,56,16),''))) else X.MOD(X.add(X.sub(E.v(15,56,16),E.v(23,7,5)),5.0), 5.0))

def f412(E, r, c):
    # Ô AA38
    # = IF(A38="",FALSE(),AND($U38>=18,$U38<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$56))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,56,66)))))

def f413(E, r, c):
    # Ô AB38
    # = IF(A38="",FALSE(),AND($U38>=0,$U38<=120,LUU_NIEN!$BH$56=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,56,60),1.0)))

def f414(E, r, c):
    # Ô AC38
    # = IF(A38="",FALSE(),AND($U38>=0,$U38<=120,LUU_NIEN!$BI$56=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,56,61),1.0)))

def f415(E, r, c):
    # Ô AD38
    # = IF(A38="",FALSE(),AND($U38>=0,$U38<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$56))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,56,92)))))

def f416(E, r, c):
    # Ô A39
    # = IF((ISNUMBER(LUU_NIEN!$A$57)),LUU_NIEN!$A$57,"")
    return (E.v(15,57,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,57,1))) else '')

def f417(E, r, c):
    # Ô C39
    # = IF(A39="","",LUU_NIEN!$L$57)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,57,12))

def f418(E, r, c):
    # Ô D39
    # = IF(A39="","",LUU_NIEN!$H$57)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,57,8))

def f419(E, r, c):
    # Ô E39
    # = IF(A39="","",LUU_NIEN!$AA$57)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,57,27))

def f420(E, r, c):
    # Ô F39
    # = IF(A39="","",LUU_NIEN!$AL$57)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,57,38))

def f421(E, r, c):
    # Ô K39
    # = IF(A39="","",TRIM(IF(AB39,"Dấu hiệu bất lợi cho cha. ","")&IF(AC39,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA39,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U39>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$57))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,57,66))))) else ''))))

def f422(E, r, c):
    # Ô L39
    # = IF(OR(A39="",LUU_NIEN!$CN$57=""),"",IFERROR(LEFT(LUU_NIEN!$CN$57,FIND(":",LUU_NIEN!$CN$57)-1),LEFT(LUU_NIEN!$CN$57,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,57,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,57,92), X.sub(X.FIND(':', E.v(15,57,92)),1.0)), lambda: X.LEFT(E.v(15,57,92), 120.0)))

def f423(E, r, c):
    # Ô M39
    # = IF(A39="","",IF(LEN(LUU_NIEN!$BN$57)>260,LEFT(LUU_NIEN!$BN$57,260)&"…",LUU_NIEN!$BN$57))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,57,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,57,66)),260.0)) else E.v(15,57,66)))

def f424(E, r, c):
    # Ô T39
    # = IF(OR(A39="",LUU_NIEN!$P$57=""),-1,MOD(LUU_NIEN!$P$57-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,57,16),''))) else X.MOD(X.add(X.sub(E.v(15,57,16),E.v(23,7,5)),5.0), 5.0))

def f425(E, r, c):
    # Ô AA39
    # = IF(A39="",FALSE(),AND($U39>=18,$U39<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$57))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,57,66)))))

def f426(E, r, c):
    # Ô AB39
    # = IF(A39="",FALSE(),AND($U39>=0,$U39<=120,LUU_NIEN!$BH$57=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,57,60),1.0)))

def f427(E, r, c):
    # Ô AC39
    # = IF(A39="",FALSE(),AND($U39>=0,$U39<=120,LUU_NIEN!$BI$57=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,57,61),1.0)))

def f428(E, r, c):
    # Ô AD39
    # = IF(A39="",FALSE(),AND($U39>=0,$U39<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$57))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,57,92)))))

def f429(E, r, c):
    # Ô A40
    # = IF((ISNUMBER(LUU_NIEN!$A$58)),LUU_NIEN!$A$58,"")
    return (E.v(15,58,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,58,1))) else '')

def f430(E, r, c):
    # Ô C40
    # = IF(A40="","",LUU_NIEN!$L$58)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,58,12))

def f431(E, r, c):
    # Ô D40
    # = IF(A40="","",LUU_NIEN!$H$58)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,58,8))

def f432(E, r, c):
    # Ô E40
    # = IF(A40="","",LUU_NIEN!$AA$58)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,58,27))

def f433(E, r, c):
    # Ô F40
    # = IF(A40="","",LUU_NIEN!$AL$58)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,58,38))

def f434(E, r, c):
    # Ô K40
    # = IF(A40="","",TRIM(IF(AB40,"Dấu hiệu bất lợi cho cha. ","")&IF(AC40,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA40,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U40>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$58))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,58,66))))) else ''))))

def f435(E, r, c):
    # Ô L40
    # = IF(OR(A40="",LUU_NIEN!$CN$58=""),"",IFERROR(LEFT(LUU_NIEN!$CN$58,FIND(":",LUU_NIEN!$CN$58)-1),LEFT(LUU_NIEN!$CN$58,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,58,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,58,92), X.sub(X.FIND(':', E.v(15,58,92)),1.0)), lambda: X.LEFT(E.v(15,58,92), 120.0)))

def f436(E, r, c):
    # Ô M40
    # = IF(A40="","",IF(LEN(LUU_NIEN!$BN$58)>260,LEFT(LUU_NIEN!$BN$58,260)&"…",LUU_NIEN!$BN$58))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,58,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,58,66)),260.0)) else E.v(15,58,66)))

def f437(E, r, c):
    # Ô T40
    # = IF(OR(A40="",LUU_NIEN!$P$58=""),-1,MOD(LUU_NIEN!$P$58-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,58,16),''))) else X.MOD(X.add(X.sub(E.v(15,58,16),E.v(23,7,5)),5.0), 5.0))

def f438(E, r, c):
    # Ô AA40
    # = IF(A40="",FALSE(),AND($U40>=18,$U40<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$58))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,58,66)))))

def f439(E, r, c):
    # Ô AB40
    # = IF(A40="",FALSE(),AND($U40>=0,$U40<=120,LUU_NIEN!$BH$58=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,58,60),1.0)))

def f440(E, r, c):
    # Ô AC40
    # = IF(A40="",FALSE(),AND($U40>=0,$U40<=120,LUU_NIEN!$BI$58=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,58,61),1.0)))

def f441(E, r, c):
    # Ô AD40
    # = IF(A40="",FALSE(),AND($U40>=0,$U40<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$58))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,58,92)))))

def f442(E, r, c):
    # Ô A41
    # = IF((ISNUMBER(LUU_NIEN!$A$59)),LUU_NIEN!$A$59,"")
    return (E.v(15,59,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,59,1))) else '')

def f443(E, r, c):
    # Ô C41
    # = IF(A41="","",LUU_NIEN!$L$59)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,59,12))

def f444(E, r, c):
    # Ô D41
    # = IF(A41="","",LUU_NIEN!$H$59)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,59,8))

def f445(E, r, c):
    # Ô E41
    # = IF(A41="","",LUU_NIEN!$AA$59)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,59,27))

def f446(E, r, c):
    # Ô F41
    # = IF(A41="","",LUU_NIEN!$AL$59)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,59,38))

def f447(E, r, c):
    # Ô K41
    # = IF(A41="","",TRIM(IF(AB41,"Dấu hiệu bất lợi cho cha. ","")&IF(AC41,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA41,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U41>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$59))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,59,66))))) else ''))))

def f448(E, r, c):
    # Ô L41
    # = IF(OR(A41="",LUU_NIEN!$CN$59=""),"",IFERROR(LEFT(LUU_NIEN!$CN$59,FIND(":",LUU_NIEN!$CN$59)-1),LEFT(LUU_NIEN!$CN$59,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,59,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,59,92), X.sub(X.FIND(':', E.v(15,59,92)),1.0)), lambda: X.LEFT(E.v(15,59,92), 120.0)))

def f449(E, r, c):
    # Ô M41
    # = IF(A41="","",IF(LEN(LUU_NIEN!$BN$59)>260,LEFT(LUU_NIEN!$BN$59,260)&"…",LUU_NIEN!$BN$59))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,59,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,59,66)),260.0)) else E.v(15,59,66)))

def f450(E, r, c):
    # Ô T41
    # = IF(OR(A41="",LUU_NIEN!$P$59=""),-1,MOD(LUU_NIEN!$P$59-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,59,16),''))) else X.MOD(X.add(X.sub(E.v(15,59,16),E.v(23,7,5)),5.0), 5.0))

def f451(E, r, c):
    # Ô AA41
    # = IF(A41="",FALSE(),AND($U41>=18,$U41<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$59))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,59,66)))))

def f452(E, r, c):
    # Ô AB41
    # = IF(A41="",FALSE(),AND($U41>=0,$U41<=120,LUU_NIEN!$BH$59=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,59,60),1.0)))

def f453(E, r, c):
    # Ô AC41
    # = IF(A41="",FALSE(),AND($U41>=0,$U41<=120,LUU_NIEN!$BI$59=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,59,61),1.0)))

def f454(E, r, c):
    # Ô AD41
    # = IF(A41="",FALSE(),AND($U41>=0,$U41<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$59))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,59,92)))))

def f455(E, r, c):
    # Ô A42
    # = IF((ISNUMBER(LUU_NIEN!$A$60)),LUU_NIEN!$A$60,"")
    return (E.v(15,60,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,60,1))) else '')

def f456(E, r, c):
    # Ô C42
    # = IF(A42="","",LUU_NIEN!$L$60)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,60,12))

def f457(E, r, c):
    # Ô D42
    # = IF(A42="","",LUU_NIEN!$H$60)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,60,8))

def f458(E, r, c):
    # Ô E42
    # = IF(A42="","",LUU_NIEN!$AA$60)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,60,27))

def f459(E, r, c):
    # Ô F42
    # = IF(A42="","",LUU_NIEN!$AL$60)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,60,38))

def f460(E, r, c):
    # Ô K42
    # = IF(A42="","",TRIM(IF(AB42,"Dấu hiệu bất lợi cho cha. ","")&IF(AC42,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA42,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U42>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$60))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,60,66))))) else ''))))

def f461(E, r, c):
    # Ô L42
    # = IF(OR(A42="",LUU_NIEN!$CN$60=""),"",IFERROR(LEFT(LUU_NIEN!$CN$60,FIND(":",LUU_NIEN!$CN$60)-1),LEFT(LUU_NIEN!$CN$60,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,60,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,60,92), X.sub(X.FIND(':', E.v(15,60,92)),1.0)), lambda: X.LEFT(E.v(15,60,92), 120.0)))

def f462(E, r, c):
    # Ô M42
    # = IF(A42="","",IF(LEN(LUU_NIEN!$BN$60)>260,LEFT(LUU_NIEN!$BN$60,260)&"…",LUU_NIEN!$BN$60))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,60,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,60,66)),260.0)) else E.v(15,60,66)))

def f463(E, r, c):
    # Ô T42
    # = IF(OR(A42="",LUU_NIEN!$P$60=""),-1,MOD(LUU_NIEN!$P$60-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,60,16),''))) else X.MOD(X.add(X.sub(E.v(15,60,16),E.v(23,7,5)),5.0), 5.0))

def f464(E, r, c):
    # Ô AA42
    # = IF(A42="",FALSE(),AND($U42>=18,$U42<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$60))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,60,66)))))

def f465(E, r, c):
    # Ô AB42
    # = IF(A42="",FALSE(),AND($U42>=0,$U42<=120,LUU_NIEN!$BH$60=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,60,60),1.0)))

def f466(E, r, c):
    # Ô AC42
    # = IF(A42="",FALSE(),AND($U42>=0,$U42<=120,LUU_NIEN!$BI$60=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,60,61),1.0)))

def f467(E, r, c):
    # Ô AD42
    # = IF(A42="",FALSE(),AND($U42>=0,$U42<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$60))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,60,92)))))

def f468(E, r, c):
    # Ô A43
    # = IF((ISNUMBER(LUU_NIEN!$A$61)),LUU_NIEN!$A$61,"")
    return (E.v(15,61,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,61,1))) else '')

def f469(E, r, c):
    # Ô C43
    # = IF(A43="","",LUU_NIEN!$L$61)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,61,12))

def f470(E, r, c):
    # Ô D43
    # = IF(A43="","",LUU_NIEN!$H$61)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,61,8))

def f471(E, r, c):
    # Ô E43
    # = IF(A43="","",LUU_NIEN!$AA$61)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,61,27))

def f472(E, r, c):
    # Ô F43
    # = IF(A43="","",LUU_NIEN!$AL$61)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,61,38))

def f473(E, r, c):
    # Ô K43
    # = IF(A43="","",TRIM(IF(AB43,"Dấu hiệu bất lợi cho cha. ","")&IF(AC43,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA43,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U43>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$61))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,61,66))))) else ''))))

def f474(E, r, c):
    # Ô L43
    # = IF(OR(A43="",LUU_NIEN!$CN$61=""),"",IFERROR(LEFT(LUU_NIEN!$CN$61,FIND(":",LUU_NIEN!$CN$61)-1),LEFT(LUU_NIEN!$CN$61,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,61,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,61,92), X.sub(X.FIND(':', E.v(15,61,92)),1.0)), lambda: X.LEFT(E.v(15,61,92), 120.0)))

def f475(E, r, c):
    # Ô M43
    # = IF(A43="","",IF(LEN(LUU_NIEN!$BN$61)>260,LEFT(LUU_NIEN!$BN$61,260)&"…",LUU_NIEN!$BN$61))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,61,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,61,66)),260.0)) else E.v(15,61,66)))

def f476(E, r, c):
    # Ô T43
    # = IF(OR(A43="",LUU_NIEN!$P$61=""),-1,MOD(LUU_NIEN!$P$61-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,61,16),''))) else X.MOD(X.add(X.sub(E.v(15,61,16),E.v(23,7,5)),5.0), 5.0))

def f477(E, r, c):
    # Ô AA43
    # = IF(A43="",FALSE(),AND($U43>=18,$U43<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$61))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,61,66)))))

def f478(E, r, c):
    # Ô AB43
    # = IF(A43="",FALSE(),AND($U43>=0,$U43<=120,LUU_NIEN!$BH$61=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,61,60),1.0)))

def f479(E, r, c):
    # Ô AC43
    # = IF(A43="",FALSE(),AND($U43>=0,$U43<=120,LUU_NIEN!$BI$61=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,61,61),1.0)))

def f480(E, r, c):
    # Ô AD43
    # = IF(A43="",FALSE(),AND($U43>=0,$U43<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$61))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,61,92)))))

def f481(E, r, c):
    # Ô A44
    # = IF((ISNUMBER(LUU_NIEN!$A$62)),LUU_NIEN!$A$62,"")
    return (E.v(15,62,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,62,1))) else '')

def f482(E, r, c):
    # Ô C44
    # = IF(A44="","",LUU_NIEN!$L$62)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,62,12))

def f483(E, r, c):
    # Ô D44
    # = IF(A44="","",LUU_NIEN!$H$62)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,62,8))

def f484(E, r, c):
    # Ô E44
    # = IF(A44="","",LUU_NIEN!$AA$62)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,62,27))

def f485(E, r, c):
    # Ô F44
    # = IF(A44="","",LUU_NIEN!$AL$62)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,62,38))

def f486(E, r, c):
    # Ô K44
    # = IF(A44="","",TRIM(IF(AB44,"Dấu hiệu bất lợi cho cha. ","")&IF(AC44,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA44,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U44>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$62))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,62,66))))) else ''))))

def f487(E, r, c):
    # Ô L44
    # = IF(OR(A44="",LUU_NIEN!$CN$62=""),"",IFERROR(LEFT(LUU_NIEN!$CN$62,FIND(":",LUU_NIEN!$CN$62)-1),LEFT(LUU_NIEN!$CN$62,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,62,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,62,92), X.sub(X.FIND(':', E.v(15,62,92)),1.0)), lambda: X.LEFT(E.v(15,62,92), 120.0)))

def f488(E, r, c):
    # Ô M44
    # = IF(A44="","",IF(LEN(LUU_NIEN!$BN$62)>260,LEFT(LUU_NIEN!$BN$62,260)&"…",LUU_NIEN!$BN$62))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,62,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,62,66)),260.0)) else E.v(15,62,66)))

def f489(E, r, c):
    # Ô T44
    # = IF(OR(A44="",LUU_NIEN!$P$62=""),-1,MOD(LUU_NIEN!$P$62-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,62,16),''))) else X.MOD(X.add(X.sub(E.v(15,62,16),E.v(23,7,5)),5.0), 5.0))

def f490(E, r, c):
    # Ô AA44
    # = IF(A44="",FALSE(),AND($U44>=18,$U44<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$62))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,62,66)))))

def f491(E, r, c):
    # Ô AB44
    # = IF(A44="",FALSE(),AND($U44>=0,$U44<=120,LUU_NIEN!$BH$62=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,62,60),1.0)))

def f492(E, r, c):
    # Ô AC44
    # = IF(A44="",FALSE(),AND($U44>=0,$U44<=120,LUU_NIEN!$BI$62=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,62,61),1.0)))

def f493(E, r, c):
    # Ô AD44
    # = IF(A44="",FALSE(),AND($U44>=0,$U44<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$62))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,62,92)))))

def f494(E, r, c):
    # Ô A45
    # = IF((ISNUMBER(LUU_NIEN!$A$63)),LUU_NIEN!$A$63,"")
    return (E.v(15,63,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,63,1))) else '')

def f495(E, r, c):
    # Ô C45
    # = IF(A45="","",LUU_NIEN!$L$63)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,63,12))

def f496(E, r, c):
    # Ô D45
    # = IF(A45="","",LUU_NIEN!$H$63)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,63,8))

def f497(E, r, c):
    # Ô E45
    # = IF(A45="","",LUU_NIEN!$AA$63)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,63,27))

def f498(E, r, c):
    # Ô F45
    # = IF(A45="","",LUU_NIEN!$AL$63)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,63,38))

def f499(E, r, c):
    # Ô K45
    # = IF(A45="","",TRIM(IF(AB45,"Dấu hiệu bất lợi cho cha. ","")&IF(AC45,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA45,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U45>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$63))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,63,66))))) else ''))))

def f500(E, r, c):
    # Ô L45
    # = IF(OR(A45="",LUU_NIEN!$CN$63=""),"",IFERROR(LEFT(LUU_NIEN!$CN$63,FIND(":",LUU_NIEN!$CN$63)-1),LEFT(LUU_NIEN!$CN$63,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,63,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,63,92), X.sub(X.FIND(':', E.v(15,63,92)),1.0)), lambda: X.LEFT(E.v(15,63,92), 120.0)))

def f501(E, r, c):
    # Ô M45
    # = IF(A45="","",IF(LEN(LUU_NIEN!$BN$63)>260,LEFT(LUU_NIEN!$BN$63,260)&"…",LUU_NIEN!$BN$63))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,63,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,63,66)),260.0)) else E.v(15,63,66)))

def f502(E, r, c):
    # Ô T45
    # = IF(OR(A45="",LUU_NIEN!$P$63=""),-1,MOD(LUU_NIEN!$P$63-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,63,16),''))) else X.MOD(X.add(X.sub(E.v(15,63,16),E.v(23,7,5)),5.0), 5.0))

def f503(E, r, c):
    # Ô AA45
    # = IF(A45="",FALSE(),AND($U45>=18,$U45<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$63))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,63,66)))))

def f504(E, r, c):
    # Ô AB45
    # = IF(A45="",FALSE(),AND($U45>=0,$U45<=120,LUU_NIEN!$BH$63=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,63,60),1.0)))

def f505(E, r, c):
    # Ô AC45
    # = IF(A45="",FALSE(),AND($U45>=0,$U45<=120,LUU_NIEN!$BI$63=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,63,61),1.0)))

def f506(E, r, c):
    # Ô AD45
    # = IF(A45="",FALSE(),AND($U45>=0,$U45<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$63))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,63,92)))))

def f507(E, r, c):
    # Ô A46
    # = IF((ISNUMBER(LUU_NIEN!$A$64)),LUU_NIEN!$A$64,"")
    return (E.v(15,64,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,64,1))) else '')

def f508(E, r, c):
    # Ô C46
    # = IF(A46="","",LUU_NIEN!$L$64)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,64,12))

def f509(E, r, c):
    # Ô D46
    # = IF(A46="","",LUU_NIEN!$H$64)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,64,8))

def f510(E, r, c):
    # Ô E46
    # = IF(A46="","",LUU_NIEN!$AA$64)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,64,27))

def f511(E, r, c):
    # Ô F46
    # = IF(A46="","",LUU_NIEN!$AL$64)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,64,38))

def f512(E, r, c):
    # Ô K46
    # = IF(A46="","",TRIM(IF(AB46,"Dấu hiệu bất lợi cho cha. ","")&IF(AC46,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA46,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U46>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$64))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,64,66))))) else ''))))

def f513(E, r, c):
    # Ô L46
    # = IF(OR(A46="",LUU_NIEN!$CN$64=""),"",IFERROR(LEFT(LUU_NIEN!$CN$64,FIND(":",LUU_NIEN!$CN$64)-1),LEFT(LUU_NIEN!$CN$64,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,64,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,64,92), X.sub(X.FIND(':', E.v(15,64,92)),1.0)), lambda: X.LEFT(E.v(15,64,92), 120.0)))

def f514(E, r, c):
    # Ô M46
    # = IF(A46="","",IF(LEN(LUU_NIEN!$BN$64)>260,LEFT(LUU_NIEN!$BN$64,260)&"…",LUU_NIEN!$BN$64))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,64,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,64,66)),260.0)) else E.v(15,64,66)))

def f515(E, r, c):
    # Ô T46
    # = IF(OR(A46="",LUU_NIEN!$P$64=""),-1,MOD(LUU_NIEN!$P$64-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,64,16),''))) else X.MOD(X.add(X.sub(E.v(15,64,16),E.v(23,7,5)),5.0), 5.0))

def f516(E, r, c):
    # Ô AA46
    # = IF(A46="",FALSE(),AND($U46>=18,$U46<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$64))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,64,66)))))

def f517(E, r, c):
    # Ô AB46
    # = IF(A46="",FALSE(),AND($U46>=0,$U46<=120,LUU_NIEN!$BH$64=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,64,60),1.0)))

def f518(E, r, c):
    # Ô AC46
    # = IF(A46="",FALSE(),AND($U46>=0,$U46<=120,LUU_NIEN!$BI$64=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,64,61),1.0)))

def f519(E, r, c):
    # Ô AD46
    # = IF(A46="",FALSE(),AND($U46>=0,$U46<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$64))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,64,92)))))

def f520(E, r, c):
    # Ô A47
    # = IF((ISNUMBER(LUU_NIEN!$A$65)),LUU_NIEN!$A$65,"")
    return (E.v(15,65,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,65,1))) else '')

def f521(E, r, c):
    # Ô C47
    # = IF(A47="","",LUU_NIEN!$L$65)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,65,12))

def f522(E, r, c):
    # Ô D47
    # = IF(A47="","",LUU_NIEN!$H$65)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,65,8))

def f523(E, r, c):
    # Ô E47
    # = IF(A47="","",LUU_NIEN!$AA$65)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,65,27))

def f524(E, r, c):
    # Ô F47
    # = IF(A47="","",LUU_NIEN!$AL$65)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,65,38))

def f525(E, r, c):
    # Ô K47
    # = IF(A47="","",TRIM(IF(AB47,"Dấu hiệu bất lợi cho cha. ","")&IF(AC47,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA47,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U47>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$65))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,65,66))))) else ''))))

def f526(E, r, c):
    # Ô L47
    # = IF(OR(A47="",LUU_NIEN!$CN$65=""),"",IFERROR(LEFT(LUU_NIEN!$CN$65,FIND(":",LUU_NIEN!$CN$65)-1),LEFT(LUU_NIEN!$CN$65,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,65,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,65,92), X.sub(X.FIND(':', E.v(15,65,92)),1.0)), lambda: X.LEFT(E.v(15,65,92), 120.0)))

def f527(E, r, c):
    # Ô M47
    # = IF(A47="","",IF(LEN(LUU_NIEN!$BN$65)>260,LEFT(LUU_NIEN!$BN$65,260)&"…",LUU_NIEN!$BN$65))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,65,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,65,66)),260.0)) else E.v(15,65,66)))

def f528(E, r, c):
    # Ô T47
    # = IF(OR(A47="",LUU_NIEN!$P$65=""),-1,MOD(LUU_NIEN!$P$65-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,65,16),''))) else X.MOD(X.add(X.sub(E.v(15,65,16),E.v(23,7,5)),5.0), 5.0))

def f529(E, r, c):
    # Ô AA47
    # = IF(A47="",FALSE(),AND($U47>=18,$U47<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$65))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,65,66)))))

def f530(E, r, c):
    # Ô AB47
    # = IF(A47="",FALSE(),AND($U47>=0,$U47<=120,LUU_NIEN!$BH$65=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,65,60),1.0)))

def f531(E, r, c):
    # Ô AC47
    # = IF(A47="",FALSE(),AND($U47>=0,$U47<=120,LUU_NIEN!$BI$65=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,65,61),1.0)))

def f532(E, r, c):
    # Ô AD47
    # = IF(A47="",FALSE(),AND($U47>=0,$U47<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$65))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,65,92)))))

def f533(E, r, c):
    # Ô A48
    # = IF((ISNUMBER(LUU_NIEN!$A$66)),LUU_NIEN!$A$66,"")
    return (E.v(15,66,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,66,1))) else '')

def f534(E, r, c):
    # Ô C48
    # = IF(A48="","",LUU_NIEN!$L$66)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,66,12))

def f535(E, r, c):
    # Ô D48
    # = IF(A48="","",LUU_NIEN!$H$66)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,66,8))

def f536(E, r, c):
    # Ô E48
    # = IF(A48="","",LUU_NIEN!$AA$66)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,66,27))

def f537(E, r, c):
    # Ô F48
    # = IF(A48="","",LUU_NIEN!$AL$66)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,66,38))

def f538(E, r, c):
    # Ô K48
    # = IF(A48="","",TRIM(IF(AB48,"Dấu hiệu bất lợi cho cha. ","")&IF(AC48,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA48,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U48>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$66))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,66,66))))) else ''))))

def f539(E, r, c):
    # Ô L48
    # = IF(OR(A48="",LUU_NIEN!$CN$66=""),"",IFERROR(LEFT(LUU_NIEN!$CN$66,FIND(":",LUU_NIEN!$CN$66)-1),LEFT(LUU_NIEN!$CN$66,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,66,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,66,92), X.sub(X.FIND(':', E.v(15,66,92)),1.0)), lambda: X.LEFT(E.v(15,66,92), 120.0)))

def f540(E, r, c):
    # Ô M48
    # = IF(A48="","",IF(LEN(LUU_NIEN!$BN$66)>260,LEFT(LUU_NIEN!$BN$66,260)&"…",LUU_NIEN!$BN$66))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,66,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,66,66)),260.0)) else E.v(15,66,66)))

def f541(E, r, c):
    # Ô T48
    # = IF(OR(A48="",LUU_NIEN!$P$66=""),-1,MOD(LUU_NIEN!$P$66-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,66,16),''))) else X.MOD(X.add(X.sub(E.v(15,66,16),E.v(23,7,5)),5.0), 5.0))

def f542(E, r, c):
    # Ô AA48
    # = IF(A48="",FALSE(),AND($U48>=18,$U48<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$66))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,66,66)))))

def f543(E, r, c):
    # Ô AB48
    # = IF(A48="",FALSE(),AND($U48>=0,$U48<=120,LUU_NIEN!$BH$66=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,66,60),1.0)))

def f544(E, r, c):
    # Ô AC48
    # = IF(A48="",FALSE(),AND($U48>=0,$U48<=120,LUU_NIEN!$BI$66=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,66,61),1.0)))

def f545(E, r, c):
    # Ô AD48
    # = IF(A48="",FALSE(),AND($U48>=0,$U48<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$66))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,66,92)))))

def f546(E, r, c):
    # Ô A49
    # = IF((ISNUMBER(LUU_NIEN!$A$67)),LUU_NIEN!$A$67,"")
    return (E.v(15,67,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,67,1))) else '')

def f547(E, r, c):
    # Ô C49
    # = IF(A49="","",LUU_NIEN!$L$67)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,67,12))

def f548(E, r, c):
    # Ô D49
    # = IF(A49="","",LUU_NIEN!$H$67)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,67,8))

def f549(E, r, c):
    # Ô E49
    # = IF(A49="","",LUU_NIEN!$AA$67)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,67,27))

def f550(E, r, c):
    # Ô F49
    # = IF(A49="","",LUU_NIEN!$AL$67)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,67,38))

def f551(E, r, c):
    # Ô K49
    # = IF(A49="","",TRIM(IF(AB49,"Dấu hiệu bất lợi cho cha. ","")&IF(AC49,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA49,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U49>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$67))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,67,66))))) else ''))))

def f552(E, r, c):
    # Ô L49
    # = IF(OR(A49="",LUU_NIEN!$CN$67=""),"",IFERROR(LEFT(LUU_NIEN!$CN$67,FIND(":",LUU_NIEN!$CN$67)-1),LEFT(LUU_NIEN!$CN$67,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,67,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,67,92), X.sub(X.FIND(':', E.v(15,67,92)),1.0)), lambda: X.LEFT(E.v(15,67,92), 120.0)))

def f553(E, r, c):
    # Ô M49
    # = IF(A49="","",IF(LEN(LUU_NIEN!$BN$67)>260,LEFT(LUU_NIEN!$BN$67,260)&"…",LUU_NIEN!$BN$67))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,67,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,67,66)),260.0)) else E.v(15,67,66)))

def f554(E, r, c):
    # Ô T49
    # = IF(OR(A49="",LUU_NIEN!$P$67=""),-1,MOD(LUU_NIEN!$P$67-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,67,16),''))) else X.MOD(X.add(X.sub(E.v(15,67,16),E.v(23,7,5)),5.0), 5.0))

def f555(E, r, c):
    # Ô AA49
    # = IF(A49="",FALSE(),AND($U49>=18,$U49<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$67))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,67,66)))))

def f556(E, r, c):
    # Ô AB49
    # = IF(A49="",FALSE(),AND($U49>=0,$U49<=120,LUU_NIEN!$BH$67=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,67,60),1.0)))

def f557(E, r, c):
    # Ô AC49
    # = IF(A49="",FALSE(),AND($U49>=0,$U49<=120,LUU_NIEN!$BI$67=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,67,61),1.0)))

def f558(E, r, c):
    # Ô AD49
    # = IF(A49="",FALSE(),AND($U49>=0,$U49<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$67))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,67,92)))))

def f559(E, r, c):
    # Ô A50
    # = IF((ISNUMBER(LUU_NIEN!$A$68)),LUU_NIEN!$A$68,"")
    return (E.v(15,68,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,68,1))) else '')

def f560(E, r, c):
    # Ô C50
    # = IF(A50="","",LUU_NIEN!$L$68)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,68,12))

def f561(E, r, c):
    # Ô D50
    # = IF(A50="","",LUU_NIEN!$H$68)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,68,8))

def f562(E, r, c):
    # Ô E50
    # = IF(A50="","",LUU_NIEN!$AA$68)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,68,27))

def f563(E, r, c):
    # Ô F50
    # = IF(A50="","",LUU_NIEN!$AL$68)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,68,38))

def f564(E, r, c):
    # Ô K50
    # = IF(A50="","",TRIM(IF(AB50,"Dấu hiệu bất lợi cho cha. ","")&IF(AC50,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA50,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U50>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$68))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,68,66))))) else ''))))

def f565(E, r, c):
    # Ô L50
    # = IF(OR(A50="",LUU_NIEN!$CN$68=""),"",IFERROR(LEFT(LUU_NIEN!$CN$68,FIND(":",LUU_NIEN!$CN$68)-1),LEFT(LUU_NIEN!$CN$68,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,68,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,68,92), X.sub(X.FIND(':', E.v(15,68,92)),1.0)), lambda: X.LEFT(E.v(15,68,92), 120.0)))

def f566(E, r, c):
    # Ô M50
    # = IF(A50="","",IF(LEN(LUU_NIEN!$BN$68)>260,LEFT(LUU_NIEN!$BN$68,260)&"…",LUU_NIEN!$BN$68))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,68,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,68,66)),260.0)) else E.v(15,68,66)))

def f567(E, r, c):
    # Ô T50
    # = IF(OR(A50="",LUU_NIEN!$P$68=""),-1,MOD(LUU_NIEN!$P$68-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,68,16),''))) else X.MOD(X.add(X.sub(E.v(15,68,16),E.v(23,7,5)),5.0), 5.0))

def f568(E, r, c):
    # Ô AA50
    # = IF(A50="",FALSE(),AND($U50>=18,$U50<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$68))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,68,66)))))

def f569(E, r, c):
    # Ô AB50
    # = IF(A50="",FALSE(),AND($U50>=0,$U50<=120,LUU_NIEN!$BH$68=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,68,60),1.0)))

def f570(E, r, c):
    # Ô AC50
    # = IF(A50="",FALSE(),AND($U50>=0,$U50<=120,LUU_NIEN!$BI$68=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,68,61),1.0)))

def f571(E, r, c):
    # Ô AD50
    # = IF(A50="",FALSE(),AND($U50>=0,$U50<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$68))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,68,92)))))

def f572(E, r, c):
    # Ô A51
    # = IF((ISNUMBER(LUU_NIEN!$A$69)),LUU_NIEN!$A$69,"")
    return (E.v(15,69,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,69,1))) else '')

def f573(E, r, c):
    # Ô C51
    # = IF(A51="","",LUU_NIEN!$L$69)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,69,12))

def f574(E, r, c):
    # Ô D51
    # = IF(A51="","",LUU_NIEN!$H$69)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,69,8))

def f575(E, r, c):
    # Ô E51
    # = IF(A51="","",LUU_NIEN!$AA$69)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,69,27))

def f576(E, r, c):
    # Ô F51
    # = IF(A51="","",LUU_NIEN!$AL$69)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,69,38))

def f577(E, r, c):
    # Ô K51
    # = IF(A51="","",TRIM(IF(AB51,"Dấu hiệu bất lợi cho cha. ","")&IF(AC51,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA51,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U51>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$69))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,69,66))))) else ''))))

def f578(E, r, c):
    # Ô L51
    # = IF(OR(A51="",LUU_NIEN!$CN$69=""),"",IFERROR(LEFT(LUU_NIEN!$CN$69,FIND(":",LUU_NIEN!$CN$69)-1),LEFT(LUU_NIEN!$CN$69,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,69,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,69,92), X.sub(X.FIND(':', E.v(15,69,92)),1.0)), lambda: X.LEFT(E.v(15,69,92), 120.0)))

def f579(E, r, c):
    # Ô M51
    # = IF(A51="","",IF(LEN(LUU_NIEN!$BN$69)>260,LEFT(LUU_NIEN!$BN$69,260)&"…",LUU_NIEN!$BN$69))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,69,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,69,66)),260.0)) else E.v(15,69,66)))

def f580(E, r, c):
    # Ô T51
    # = IF(OR(A51="",LUU_NIEN!$P$69=""),-1,MOD(LUU_NIEN!$P$69-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,69,16),''))) else X.MOD(X.add(X.sub(E.v(15,69,16),E.v(23,7,5)),5.0), 5.0))

def f581(E, r, c):
    # Ô AA51
    # = IF(A51="",FALSE(),AND($U51>=18,$U51<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$69))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,69,66)))))

def f582(E, r, c):
    # Ô AB51
    # = IF(A51="",FALSE(),AND($U51>=0,$U51<=120,LUU_NIEN!$BH$69=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,69,60),1.0)))

def f583(E, r, c):
    # Ô AC51
    # = IF(A51="",FALSE(),AND($U51>=0,$U51<=120,LUU_NIEN!$BI$69=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,69,61),1.0)))

def f584(E, r, c):
    # Ô AD51
    # = IF(A51="",FALSE(),AND($U51>=0,$U51<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$69))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,69,92)))))

def f585(E, r, c):
    # Ô A52
    # = IF((ISNUMBER(LUU_NIEN!$A$70)),LUU_NIEN!$A$70,"")
    return (E.v(15,70,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,70,1))) else '')

def f586(E, r, c):
    # Ô C52
    # = IF(A52="","",LUU_NIEN!$L$70)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,70,12))

def f587(E, r, c):
    # Ô D52
    # = IF(A52="","",LUU_NIEN!$H$70)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,70,8))

def f588(E, r, c):
    # Ô E52
    # = IF(A52="","",LUU_NIEN!$AA$70)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,70,27))

def f589(E, r, c):
    # Ô F52
    # = IF(A52="","",LUU_NIEN!$AL$70)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,70,38))

def f590(E, r, c):
    # Ô K52
    # = IF(A52="","",TRIM(IF(AB52,"Dấu hiệu bất lợi cho cha. ","")&IF(AC52,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA52,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U52>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$70))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,70,66))))) else ''))))

def f591(E, r, c):
    # Ô L52
    # = IF(OR(A52="",LUU_NIEN!$CN$70=""),"",IFERROR(LEFT(LUU_NIEN!$CN$70,FIND(":",LUU_NIEN!$CN$70)-1),LEFT(LUU_NIEN!$CN$70,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,70,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,70,92), X.sub(X.FIND(':', E.v(15,70,92)),1.0)), lambda: X.LEFT(E.v(15,70,92), 120.0)))

def f592(E, r, c):
    # Ô M52
    # = IF(A52="","",IF(LEN(LUU_NIEN!$BN$70)>260,LEFT(LUU_NIEN!$BN$70,260)&"…",LUU_NIEN!$BN$70))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,70,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,70,66)),260.0)) else E.v(15,70,66)))

def f593(E, r, c):
    # Ô T52
    # = IF(OR(A52="",LUU_NIEN!$P$70=""),-1,MOD(LUU_NIEN!$P$70-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,70,16),''))) else X.MOD(X.add(X.sub(E.v(15,70,16),E.v(23,7,5)),5.0), 5.0))

def f594(E, r, c):
    # Ô AA52
    # = IF(A52="",FALSE(),AND($U52>=18,$U52<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$70))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,70,66)))))

def f595(E, r, c):
    # Ô AB52
    # = IF(A52="",FALSE(),AND($U52>=0,$U52<=120,LUU_NIEN!$BH$70=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,70,60),1.0)))

def f596(E, r, c):
    # Ô AC52
    # = IF(A52="",FALSE(),AND($U52>=0,$U52<=120,LUU_NIEN!$BI$70=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,70,61),1.0)))

def f597(E, r, c):
    # Ô AD52
    # = IF(A52="",FALSE(),AND($U52>=0,$U52<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$70))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,70,92)))))

def f598(E, r, c):
    # Ô A53
    # = IF((ISNUMBER(LUU_NIEN!$A$71)),LUU_NIEN!$A$71,"")
    return (E.v(15,71,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,71,1))) else '')

def f599(E, r, c):
    # Ô C53
    # = IF(A53="","",LUU_NIEN!$L$71)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,71,12))

def f600(E, r, c):
    # Ô D53
    # = IF(A53="","",LUU_NIEN!$H$71)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,71,8))

def f601(E, r, c):
    # Ô E53
    # = IF(A53="","",LUU_NIEN!$AA$71)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,71,27))

def f602(E, r, c):
    # Ô F53
    # = IF(A53="","",LUU_NIEN!$AL$71)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,71,38))

def f603(E, r, c):
    # Ô K53
    # = IF(A53="","",TRIM(IF(AB53,"Dấu hiệu bất lợi cho cha. ","")&IF(AC53,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA53,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U53>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$71))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,71,66))))) else ''))))

def f604(E, r, c):
    # Ô L53
    # = IF(OR(A53="",LUU_NIEN!$CN$71=""),"",IFERROR(LEFT(LUU_NIEN!$CN$71,FIND(":",LUU_NIEN!$CN$71)-1),LEFT(LUU_NIEN!$CN$71,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,71,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,71,92), X.sub(X.FIND(':', E.v(15,71,92)),1.0)), lambda: X.LEFT(E.v(15,71,92), 120.0)))

def f605(E, r, c):
    # Ô M53
    # = IF(A53="","",IF(LEN(LUU_NIEN!$BN$71)>260,LEFT(LUU_NIEN!$BN$71,260)&"…",LUU_NIEN!$BN$71))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,71,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,71,66)),260.0)) else E.v(15,71,66)))

def f606(E, r, c):
    # Ô T53
    # = IF(OR(A53="",LUU_NIEN!$P$71=""),-1,MOD(LUU_NIEN!$P$71-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,71,16),''))) else X.MOD(X.add(X.sub(E.v(15,71,16),E.v(23,7,5)),5.0), 5.0))

def f607(E, r, c):
    # Ô AA53
    # = IF(A53="",FALSE(),AND($U53>=18,$U53<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$71))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,71,66)))))

def f608(E, r, c):
    # Ô AB53
    # = IF(A53="",FALSE(),AND($U53>=0,$U53<=120,LUU_NIEN!$BH$71=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,71,60),1.0)))

def f609(E, r, c):
    # Ô AC53
    # = IF(A53="",FALSE(),AND($U53>=0,$U53<=120,LUU_NIEN!$BI$71=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,71,61),1.0)))

def f610(E, r, c):
    # Ô AD53
    # = IF(A53="",FALSE(),AND($U53>=0,$U53<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$71))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,71,92)))))

def f611(E, r, c):
    # Ô A54
    # = IF((ISNUMBER(LUU_NIEN!$A$72)),LUU_NIEN!$A$72,"")
    return (E.v(15,72,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,72,1))) else '')

def f612(E, r, c):
    # Ô C54
    # = IF(A54="","",LUU_NIEN!$L$72)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,72,12))

def f613(E, r, c):
    # Ô D54
    # = IF(A54="","",LUU_NIEN!$H$72)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,72,8))

def f614(E, r, c):
    # Ô E54
    # = IF(A54="","",LUU_NIEN!$AA$72)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,72,27))

def f615(E, r, c):
    # Ô F54
    # = IF(A54="","",LUU_NIEN!$AL$72)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,72,38))

def f616(E, r, c):
    # Ô K54
    # = IF(A54="","",TRIM(IF(AB54,"Dấu hiệu bất lợi cho cha. ","")&IF(AC54,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA54,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U54>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$72))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,72,66))))) else ''))))

def f617(E, r, c):
    # Ô L54
    # = IF(OR(A54="",LUU_NIEN!$CN$72=""),"",IFERROR(LEFT(LUU_NIEN!$CN$72,FIND(":",LUU_NIEN!$CN$72)-1),LEFT(LUU_NIEN!$CN$72,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,72,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,72,92), X.sub(X.FIND(':', E.v(15,72,92)),1.0)), lambda: X.LEFT(E.v(15,72,92), 120.0)))

def f618(E, r, c):
    # Ô M54
    # = IF(A54="","",IF(LEN(LUU_NIEN!$BN$72)>260,LEFT(LUU_NIEN!$BN$72,260)&"…",LUU_NIEN!$BN$72))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,72,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,72,66)),260.0)) else E.v(15,72,66)))

def f619(E, r, c):
    # Ô T54
    # = IF(OR(A54="",LUU_NIEN!$P$72=""),-1,MOD(LUU_NIEN!$P$72-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,72,16),''))) else X.MOD(X.add(X.sub(E.v(15,72,16),E.v(23,7,5)),5.0), 5.0))

def f620(E, r, c):
    # Ô AA54
    # = IF(A54="",FALSE(),AND($U54>=18,$U54<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$72))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,72,66)))))

def f621(E, r, c):
    # Ô AB54
    # = IF(A54="",FALSE(),AND($U54>=0,$U54<=120,LUU_NIEN!$BH$72=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,72,60),1.0)))

def f622(E, r, c):
    # Ô AC54
    # = IF(A54="",FALSE(),AND($U54>=0,$U54<=120,LUU_NIEN!$BI$72=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,72,61),1.0)))

def f623(E, r, c):
    # Ô AD54
    # = IF(A54="",FALSE(),AND($U54>=0,$U54<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$72))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,72,92)))))

def f624(E, r, c):
    # Ô A55
    # = IF((ISNUMBER(LUU_NIEN!$A$73)),LUU_NIEN!$A$73,"")
    return (E.v(15,73,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,73,1))) else '')

def f625(E, r, c):
    # Ô C55
    # = IF(A55="","",LUU_NIEN!$L$73)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,73,12))

def f626(E, r, c):
    # Ô D55
    # = IF(A55="","",LUU_NIEN!$H$73)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,73,8))

def f627(E, r, c):
    # Ô E55
    # = IF(A55="","",LUU_NIEN!$AA$73)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,73,27))

def f628(E, r, c):
    # Ô F55
    # = IF(A55="","",LUU_NIEN!$AL$73)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,73,38))

def f629(E, r, c):
    # Ô K55
    # = IF(A55="","",TRIM(IF(AB55,"Dấu hiệu bất lợi cho cha. ","")&IF(AC55,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA55,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U55>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$73))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,73,66))))) else ''))))

def f630(E, r, c):
    # Ô L55
    # = IF(OR(A55="",LUU_NIEN!$CN$73=""),"",IFERROR(LEFT(LUU_NIEN!$CN$73,FIND(":",LUU_NIEN!$CN$73)-1),LEFT(LUU_NIEN!$CN$73,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,73,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,73,92), X.sub(X.FIND(':', E.v(15,73,92)),1.0)), lambda: X.LEFT(E.v(15,73,92), 120.0)))

def f631(E, r, c):
    # Ô M55
    # = IF(A55="","",IF(LEN(LUU_NIEN!$BN$73)>260,LEFT(LUU_NIEN!$BN$73,260)&"…",LUU_NIEN!$BN$73))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,73,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,73,66)),260.0)) else E.v(15,73,66)))

def f632(E, r, c):
    # Ô T55
    # = IF(OR(A55="",LUU_NIEN!$P$73=""),-1,MOD(LUU_NIEN!$P$73-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,73,16),''))) else X.MOD(X.add(X.sub(E.v(15,73,16),E.v(23,7,5)),5.0), 5.0))

def f633(E, r, c):
    # Ô AA55
    # = IF(A55="",FALSE(),AND($U55>=18,$U55<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$73))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,73,66)))))

def f634(E, r, c):
    # Ô AB55
    # = IF(A55="",FALSE(),AND($U55>=0,$U55<=120,LUU_NIEN!$BH$73=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,73,60),1.0)))

def f635(E, r, c):
    # Ô AC55
    # = IF(A55="",FALSE(),AND($U55>=0,$U55<=120,LUU_NIEN!$BI$73=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,73,61),1.0)))

def f636(E, r, c):
    # Ô AD55
    # = IF(A55="",FALSE(),AND($U55>=0,$U55<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$73))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,73,92)))))

def f637(E, r, c):
    # Ô A56
    # = IF((ISNUMBER(LUU_NIEN!$A$74)),LUU_NIEN!$A$74,"")
    return (E.v(15,74,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,74,1))) else '')

def f638(E, r, c):
    # Ô C56
    # = IF(A56="","",LUU_NIEN!$L$74)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,74,12))

def f639(E, r, c):
    # Ô D56
    # = IF(A56="","",LUU_NIEN!$H$74)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,74,8))

def f640(E, r, c):
    # Ô E56
    # = IF(A56="","",LUU_NIEN!$AA$74)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,74,27))

def f641(E, r, c):
    # Ô F56
    # = IF(A56="","",LUU_NIEN!$AL$74)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,74,38))

def f642(E, r, c):
    # Ô K56
    # = IF(A56="","",TRIM(IF(AB56,"Dấu hiệu bất lợi cho cha. ","")&IF(AC56,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA56,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U56>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$74))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,74,66))))) else ''))))

def f643(E, r, c):
    # Ô L56
    # = IF(OR(A56="",LUU_NIEN!$CN$74=""),"",IFERROR(LEFT(LUU_NIEN!$CN$74,FIND(":",LUU_NIEN!$CN$74)-1),LEFT(LUU_NIEN!$CN$74,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,74,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,74,92), X.sub(X.FIND(':', E.v(15,74,92)),1.0)), lambda: X.LEFT(E.v(15,74,92), 120.0)))

def f644(E, r, c):
    # Ô M56
    # = IF(A56="","",IF(LEN(LUU_NIEN!$BN$74)>260,LEFT(LUU_NIEN!$BN$74,260)&"…",LUU_NIEN!$BN$74))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,74,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,74,66)),260.0)) else E.v(15,74,66)))

def f645(E, r, c):
    # Ô T56
    # = IF(OR(A56="",LUU_NIEN!$P$74=""),-1,MOD(LUU_NIEN!$P$74-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,74,16),''))) else X.MOD(X.add(X.sub(E.v(15,74,16),E.v(23,7,5)),5.0), 5.0))

def f646(E, r, c):
    # Ô AA56
    # = IF(A56="",FALSE(),AND($U56>=18,$U56<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$74))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,74,66)))))

def f647(E, r, c):
    # Ô AB56
    # = IF(A56="",FALSE(),AND($U56>=0,$U56<=120,LUU_NIEN!$BH$74=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,74,60),1.0)))

def f648(E, r, c):
    # Ô AC56
    # = IF(A56="",FALSE(),AND($U56>=0,$U56<=120,LUU_NIEN!$BI$74=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,74,61),1.0)))

def f649(E, r, c):
    # Ô AD56
    # = IF(A56="",FALSE(),AND($U56>=0,$U56<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$74))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,74,92)))))

def f650(E, r, c):
    # Ô A57
    # = IF((ISNUMBER(LUU_NIEN!$A$75)),LUU_NIEN!$A$75,"")
    return (E.v(15,75,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,75,1))) else '')

def f651(E, r, c):
    # Ô C57
    # = IF(A57="","",LUU_NIEN!$L$75)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,75,12))

def f652(E, r, c):
    # Ô D57
    # = IF(A57="","",LUU_NIEN!$H$75)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,75,8))

def f653(E, r, c):
    # Ô E57
    # = IF(A57="","",LUU_NIEN!$AA$75)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,75,27))

def f654(E, r, c):
    # Ô F57
    # = IF(A57="","",LUU_NIEN!$AL$75)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,75,38))

def f655(E, r, c):
    # Ô K57
    # = IF(A57="","",TRIM(IF(AB57,"Dấu hiệu bất lợi cho cha. ","")&IF(AC57,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA57,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U57>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$75))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,75,66))))) else ''))))

def f656(E, r, c):
    # Ô L57
    # = IF(OR(A57="",LUU_NIEN!$CN$75=""),"",IFERROR(LEFT(LUU_NIEN!$CN$75,FIND(":",LUU_NIEN!$CN$75)-1),LEFT(LUU_NIEN!$CN$75,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,75,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,75,92), X.sub(X.FIND(':', E.v(15,75,92)),1.0)), lambda: X.LEFT(E.v(15,75,92), 120.0)))

def f657(E, r, c):
    # Ô M57
    # = IF(A57="","",IF(LEN(LUU_NIEN!$BN$75)>260,LEFT(LUU_NIEN!$BN$75,260)&"…",LUU_NIEN!$BN$75))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,75,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,75,66)),260.0)) else E.v(15,75,66)))

def f658(E, r, c):
    # Ô T57
    # = IF(OR(A57="",LUU_NIEN!$P$75=""),-1,MOD(LUU_NIEN!$P$75-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,75,16),''))) else X.MOD(X.add(X.sub(E.v(15,75,16),E.v(23,7,5)),5.0), 5.0))

def f659(E, r, c):
    # Ô AA57
    # = IF(A57="",FALSE(),AND($U57>=18,$U57<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$75))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,75,66)))))

def f660(E, r, c):
    # Ô AB57
    # = IF(A57="",FALSE(),AND($U57>=0,$U57<=120,LUU_NIEN!$BH$75=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,75,60),1.0)))

def f661(E, r, c):
    # Ô AC57
    # = IF(A57="",FALSE(),AND($U57>=0,$U57<=120,LUU_NIEN!$BI$75=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,75,61),1.0)))

def f662(E, r, c):
    # Ô AD57
    # = IF(A57="",FALSE(),AND($U57>=0,$U57<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$75))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,75,92)))))

def f663(E, r, c):
    # Ô A58
    # = IF((ISNUMBER(LUU_NIEN!$A$76)),LUU_NIEN!$A$76,"")
    return (E.v(15,76,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,76,1))) else '')

def f664(E, r, c):
    # Ô C58
    # = IF(A58="","",LUU_NIEN!$L$76)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,76,12))

def f665(E, r, c):
    # Ô D58
    # = IF(A58="","",LUU_NIEN!$H$76)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,76,8))

def f666(E, r, c):
    # Ô E58
    # = IF(A58="","",LUU_NIEN!$AA$76)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,76,27))

def f667(E, r, c):
    # Ô F58
    # = IF(A58="","",LUU_NIEN!$AL$76)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,76,38))

def f668(E, r, c):
    # Ô K58
    # = IF(A58="","",TRIM(IF(AB58,"Dấu hiệu bất lợi cho cha. ","")&IF(AC58,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA58,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U58>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$76))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,76,66))))) else ''))))

def f669(E, r, c):
    # Ô L58
    # = IF(OR(A58="",LUU_NIEN!$CN$76=""),"",IFERROR(LEFT(LUU_NIEN!$CN$76,FIND(":",LUU_NIEN!$CN$76)-1),LEFT(LUU_NIEN!$CN$76,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,76,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,76,92), X.sub(X.FIND(':', E.v(15,76,92)),1.0)), lambda: X.LEFT(E.v(15,76,92), 120.0)))

def f670(E, r, c):
    # Ô M58
    # = IF(A58="","",IF(LEN(LUU_NIEN!$BN$76)>260,LEFT(LUU_NIEN!$BN$76,260)&"…",LUU_NIEN!$BN$76))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,76,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,76,66)),260.0)) else E.v(15,76,66)))

def f671(E, r, c):
    # Ô T58
    # = IF(OR(A58="",LUU_NIEN!$P$76=""),-1,MOD(LUU_NIEN!$P$76-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,76,16),''))) else X.MOD(X.add(X.sub(E.v(15,76,16),E.v(23,7,5)),5.0), 5.0))

def f672(E, r, c):
    # Ô AA58
    # = IF(A58="",FALSE(),AND($U58>=18,$U58<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$76))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,76,66)))))

def f673(E, r, c):
    # Ô AB58
    # = IF(A58="",FALSE(),AND($U58>=0,$U58<=120,LUU_NIEN!$BH$76=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,76,60),1.0)))

def f674(E, r, c):
    # Ô AC58
    # = IF(A58="",FALSE(),AND($U58>=0,$U58<=120,LUU_NIEN!$BI$76=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,76,61),1.0)))

def f675(E, r, c):
    # Ô AD58
    # = IF(A58="",FALSE(),AND($U58>=0,$U58<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$76))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,76,92)))))

def f676(E, r, c):
    # Ô A59
    # = IF((ISNUMBER(LUU_NIEN!$A$77)),LUU_NIEN!$A$77,"")
    return (E.v(15,77,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,77,1))) else '')

def f677(E, r, c):
    # Ô C59
    # = IF(A59="","",LUU_NIEN!$L$77)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,77,12))

def f678(E, r, c):
    # Ô D59
    # = IF(A59="","",LUU_NIEN!$H$77)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,77,8))

def f679(E, r, c):
    # Ô E59
    # = IF(A59="","",LUU_NIEN!$AA$77)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,77,27))

def f680(E, r, c):
    # Ô F59
    # = IF(A59="","",LUU_NIEN!$AL$77)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,77,38))

def f681(E, r, c):
    # Ô K59
    # = IF(A59="","",TRIM(IF(AB59,"Dấu hiệu bất lợi cho cha. ","")&IF(AC59,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA59,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U59>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$77))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,77,66))))) else ''))))

def f682(E, r, c):
    # Ô L59
    # = IF(OR(A59="",LUU_NIEN!$CN$77=""),"",IFERROR(LEFT(LUU_NIEN!$CN$77,FIND(":",LUU_NIEN!$CN$77)-1),LEFT(LUU_NIEN!$CN$77,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,77,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,77,92), X.sub(X.FIND(':', E.v(15,77,92)),1.0)), lambda: X.LEFT(E.v(15,77,92), 120.0)))

def f683(E, r, c):
    # Ô M59
    # = IF(A59="","",IF(LEN(LUU_NIEN!$BN$77)>260,LEFT(LUU_NIEN!$BN$77,260)&"…",LUU_NIEN!$BN$77))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,77,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,77,66)),260.0)) else E.v(15,77,66)))

def f684(E, r, c):
    # Ô T59
    # = IF(OR(A59="",LUU_NIEN!$P$77=""),-1,MOD(LUU_NIEN!$P$77-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,77,16),''))) else X.MOD(X.add(X.sub(E.v(15,77,16),E.v(23,7,5)),5.0), 5.0))

def f685(E, r, c):
    # Ô AA59
    # = IF(A59="",FALSE(),AND($U59>=18,$U59<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$77))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,77,66)))))

def f686(E, r, c):
    # Ô AB59
    # = IF(A59="",FALSE(),AND($U59>=0,$U59<=120,LUU_NIEN!$BH$77=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,77,60),1.0)))

def f687(E, r, c):
    # Ô AC59
    # = IF(A59="",FALSE(),AND($U59>=0,$U59<=120,LUU_NIEN!$BI$77=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,77,61),1.0)))

def f688(E, r, c):
    # Ô AD59
    # = IF(A59="",FALSE(),AND($U59>=0,$U59<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$77))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,77,92)))))

def f689(E, r, c):
    # Ô A60
    # = IF((ISNUMBER(LUU_NIEN!$A$78)),LUU_NIEN!$A$78,"")
    return (E.v(15,78,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,78,1))) else '')

def f690(E, r, c):
    # Ô C60
    # = IF(A60="","",LUU_NIEN!$L$78)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,78,12))

def f691(E, r, c):
    # Ô D60
    # = IF(A60="","",LUU_NIEN!$H$78)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,78,8))

def f692(E, r, c):
    # Ô E60
    # = IF(A60="","",LUU_NIEN!$AA$78)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,78,27))

def f693(E, r, c):
    # Ô F60
    # = IF(A60="","",LUU_NIEN!$AL$78)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,78,38))

def f694(E, r, c):
    # Ô K60
    # = IF(A60="","",TRIM(IF(AB60,"Dấu hiệu bất lợi cho cha. ","")&IF(AC60,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA60,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U60>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$78))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,78,66))))) else ''))))

def f695(E, r, c):
    # Ô L60
    # = IF(OR(A60="",LUU_NIEN!$CN$78=""),"",IFERROR(LEFT(LUU_NIEN!$CN$78,FIND(":",LUU_NIEN!$CN$78)-1),LEFT(LUU_NIEN!$CN$78,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,78,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,78,92), X.sub(X.FIND(':', E.v(15,78,92)),1.0)), lambda: X.LEFT(E.v(15,78,92), 120.0)))

def f696(E, r, c):
    # Ô M60
    # = IF(A60="","",IF(LEN(LUU_NIEN!$BN$78)>260,LEFT(LUU_NIEN!$BN$78,260)&"…",LUU_NIEN!$BN$78))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,78,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,78,66)),260.0)) else E.v(15,78,66)))

def f697(E, r, c):
    # Ô T60
    # = IF(OR(A60="",LUU_NIEN!$P$78=""),-1,MOD(LUU_NIEN!$P$78-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,78,16),''))) else X.MOD(X.add(X.sub(E.v(15,78,16),E.v(23,7,5)),5.0), 5.0))

def f698(E, r, c):
    # Ô AA60
    # = IF(A60="",FALSE(),AND($U60>=18,$U60<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$78))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,78,66)))))

def f699(E, r, c):
    # Ô AB60
    # = IF(A60="",FALSE(),AND($U60>=0,$U60<=120,LUU_NIEN!$BH$78=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,78,60),1.0)))

def f700(E, r, c):
    # Ô AC60
    # = IF(A60="",FALSE(),AND($U60>=0,$U60<=120,LUU_NIEN!$BI$78=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,78,61),1.0)))

def f701(E, r, c):
    # Ô AD60
    # = IF(A60="",FALSE(),AND($U60>=0,$U60<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$78))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,78,92)))))

def f702(E, r, c):
    # Ô A61
    # = IF((ISNUMBER(LUU_NIEN!$A$79)),LUU_NIEN!$A$79,"")
    return (E.v(15,79,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,79,1))) else '')

def f703(E, r, c):
    # Ô C61
    # = IF(A61="","",LUU_NIEN!$L$79)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,79,12))

def f704(E, r, c):
    # Ô D61
    # = IF(A61="","",LUU_NIEN!$H$79)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,79,8))

def f705(E, r, c):
    # Ô E61
    # = IF(A61="","",LUU_NIEN!$AA$79)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,79,27))

def f706(E, r, c):
    # Ô F61
    # = IF(A61="","",LUU_NIEN!$AL$79)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,79,38))

def f707(E, r, c):
    # Ô K61
    # = IF(A61="","",TRIM(IF(AB61,"Dấu hiệu bất lợi cho cha. ","")&IF(AC61,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA61,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U61>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$79))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,79,66))))) else ''))))

def f708(E, r, c):
    # Ô L61
    # = IF(OR(A61="",LUU_NIEN!$CN$79=""),"",IFERROR(LEFT(LUU_NIEN!$CN$79,FIND(":",LUU_NIEN!$CN$79)-1),LEFT(LUU_NIEN!$CN$79,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,79,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,79,92), X.sub(X.FIND(':', E.v(15,79,92)),1.0)), lambda: X.LEFT(E.v(15,79,92), 120.0)))

def f709(E, r, c):
    # Ô M61
    # = IF(A61="","",IF(LEN(LUU_NIEN!$BN$79)>260,LEFT(LUU_NIEN!$BN$79,260)&"…",LUU_NIEN!$BN$79))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,79,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,79,66)),260.0)) else E.v(15,79,66)))

def f710(E, r, c):
    # Ô T61
    # = IF(OR(A61="",LUU_NIEN!$P$79=""),-1,MOD(LUU_NIEN!$P$79-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,79,16),''))) else X.MOD(X.add(X.sub(E.v(15,79,16),E.v(23,7,5)),5.0), 5.0))

def f711(E, r, c):
    # Ô AA61
    # = IF(A61="",FALSE(),AND($U61>=18,$U61<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$79))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,79,66)))))

def f712(E, r, c):
    # Ô AB61
    # = IF(A61="",FALSE(),AND($U61>=0,$U61<=120,LUU_NIEN!$BH$79=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,79,60),1.0)))

def f713(E, r, c):
    # Ô AC61
    # = IF(A61="",FALSE(),AND($U61>=0,$U61<=120,LUU_NIEN!$BI$79=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,79,61),1.0)))

def f714(E, r, c):
    # Ô AD61
    # = IF(A61="",FALSE(),AND($U61>=0,$U61<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$79))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,79,92)))))

def f715(E, r, c):
    # Ô A62
    # = IF((ISNUMBER(LUU_NIEN!$A$80)),LUU_NIEN!$A$80,"")
    return (E.v(15,80,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,80,1))) else '')

def f716(E, r, c):
    # Ô C62
    # = IF(A62="","",LUU_NIEN!$L$80)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,80,12))

def f717(E, r, c):
    # Ô D62
    # = IF(A62="","",LUU_NIEN!$H$80)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,80,8))

def f718(E, r, c):
    # Ô E62
    # = IF(A62="","",LUU_NIEN!$AA$80)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,80,27))

def f719(E, r, c):
    # Ô F62
    # = IF(A62="","",LUU_NIEN!$AL$80)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,80,38))

def f720(E, r, c):
    # Ô K62
    # = IF(A62="","",TRIM(IF(AB62,"Dấu hiệu bất lợi cho cha. ","")&IF(AC62,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA62,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U62>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$80))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,80,66))))) else ''))))

def f721(E, r, c):
    # Ô L62
    # = IF(OR(A62="",LUU_NIEN!$CN$80=""),"",IFERROR(LEFT(LUU_NIEN!$CN$80,FIND(":",LUU_NIEN!$CN$80)-1),LEFT(LUU_NIEN!$CN$80,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,80,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,80,92), X.sub(X.FIND(':', E.v(15,80,92)),1.0)), lambda: X.LEFT(E.v(15,80,92), 120.0)))

def f722(E, r, c):
    # Ô M62
    # = IF(A62="","",IF(LEN(LUU_NIEN!$BN$80)>260,LEFT(LUU_NIEN!$BN$80,260)&"…",LUU_NIEN!$BN$80))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,80,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,80,66)),260.0)) else E.v(15,80,66)))

def f723(E, r, c):
    # Ô T62
    # = IF(OR(A62="",LUU_NIEN!$P$80=""),-1,MOD(LUU_NIEN!$P$80-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,80,16),''))) else X.MOD(X.add(X.sub(E.v(15,80,16),E.v(23,7,5)),5.0), 5.0))

def f724(E, r, c):
    # Ô AA62
    # = IF(A62="",FALSE(),AND($U62>=18,$U62<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$80))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,80,66)))))

def f725(E, r, c):
    # Ô AB62
    # = IF(A62="",FALSE(),AND($U62>=0,$U62<=120,LUU_NIEN!$BH$80=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,80,60),1.0)))

def f726(E, r, c):
    # Ô AC62
    # = IF(A62="",FALSE(),AND($U62>=0,$U62<=120,LUU_NIEN!$BI$80=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,80,61),1.0)))

def f727(E, r, c):
    # Ô AD62
    # = IF(A62="",FALSE(),AND($U62>=0,$U62<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$80))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,80,92)))))

def f728(E, r, c):
    # Ô A63
    # = IF((ISNUMBER(LUU_NIEN!$A$81)),LUU_NIEN!$A$81,"")
    return (E.v(15,81,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,81,1))) else '')

def f729(E, r, c):
    # Ô C63
    # = IF(A63="","",LUU_NIEN!$L$81)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,81,12))

def f730(E, r, c):
    # Ô D63
    # = IF(A63="","",LUU_NIEN!$H$81)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,81,8))

def f731(E, r, c):
    # Ô E63
    # = IF(A63="","",LUU_NIEN!$AA$81)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,81,27))

def f732(E, r, c):
    # Ô F63
    # = IF(A63="","",LUU_NIEN!$AL$81)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,81,38))

def f733(E, r, c):
    # Ô K63
    # = IF(A63="","",TRIM(IF(AB63,"Dấu hiệu bất lợi cho cha. ","")&IF(AC63,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA63,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U63>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$81))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,81,66))))) else ''))))

def f734(E, r, c):
    # Ô L63
    # = IF(OR(A63="",LUU_NIEN!$CN$81=""),"",IFERROR(LEFT(LUU_NIEN!$CN$81,FIND(":",LUU_NIEN!$CN$81)-1),LEFT(LUU_NIEN!$CN$81,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,81,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,81,92), X.sub(X.FIND(':', E.v(15,81,92)),1.0)), lambda: X.LEFT(E.v(15,81,92), 120.0)))

def f735(E, r, c):
    # Ô M63
    # = IF(A63="","",IF(LEN(LUU_NIEN!$BN$81)>260,LEFT(LUU_NIEN!$BN$81,260)&"…",LUU_NIEN!$BN$81))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,81,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,81,66)),260.0)) else E.v(15,81,66)))

def f736(E, r, c):
    # Ô T63
    # = IF(OR(A63="",LUU_NIEN!$P$81=""),-1,MOD(LUU_NIEN!$P$81-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,81,16),''))) else X.MOD(X.add(X.sub(E.v(15,81,16),E.v(23,7,5)),5.0), 5.0))

def f737(E, r, c):
    # Ô AA63
    # = IF(A63="",FALSE(),AND($U63>=18,$U63<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$81))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,81,66)))))

def f738(E, r, c):
    # Ô AB63
    # = IF(A63="",FALSE(),AND($U63>=0,$U63<=120,LUU_NIEN!$BH$81=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,81,60),1.0)))

def f739(E, r, c):
    # Ô AC63
    # = IF(A63="",FALSE(),AND($U63>=0,$U63<=120,LUU_NIEN!$BI$81=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,81,61),1.0)))

def f740(E, r, c):
    # Ô AD63
    # = IF(A63="",FALSE(),AND($U63>=0,$U63<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$81))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,81,92)))))

def f741(E, r, c):
    # Ô A64
    # = IF((ISNUMBER(LUU_NIEN!$A$82)),LUU_NIEN!$A$82,"")
    return (E.v(15,82,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,82,1))) else '')

def f742(E, r, c):
    # Ô C64
    # = IF(A64="","",LUU_NIEN!$L$82)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,82,12))

def f743(E, r, c):
    # Ô D64
    # = IF(A64="","",LUU_NIEN!$H$82)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,82,8))

def f744(E, r, c):
    # Ô E64
    # = IF(A64="","",LUU_NIEN!$AA$82)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,82,27))

def f745(E, r, c):
    # Ô F64
    # = IF(A64="","",LUU_NIEN!$AL$82)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,82,38))

def f746(E, r, c):
    # Ô K64
    # = IF(A64="","",TRIM(IF(AB64,"Dấu hiệu bất lợi cho cha. ","")&IF(AC64,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA64,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U64>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$82))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,82,66))))) else ''))))

def f747(E, r, c):
    # Ô L64
    # = IF(OR(A64="",LUU_NIEN!$CN$82=""),"",IFERROR(LEFT(LUU_NIEN!$CN$82,FIND(":",LUU_NIEN!$CN$82)-1),LEFT(LUU_NIEN!$CN$82,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,82,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,82,92), X.sub(X.FIND(':', E.v(15,82,92)),1.0)), lambda: X.LEFT(E.v(15,82,92), 120.0)))

def f748(E, r, c):
    # Ô M64
    # = IF(A64="","",IF(LEN(LUU_NIEN!$BN$82)>260,LEFT(LUU_NIEN!$BN$82,260)&"…",LUU_NIEN!$BN$82))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,82,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,82,66)),260.0)) else E.v(15,82,66)))

def f749(E, r, c):
    # Ô T64
    # = IF(OR(A64="",LUU_NIEN!$P$82=""),-1,MOD(LUU_NIEN!$P$82-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,82,16),''))) else X.MOD(X.add(X.sub(E.v(15,82,16),E.v(23,7,5)),5.0), 5.0))

def f750(E, r, c):
    # Ô AA64
    # = IF(A64="",FALSE(),AND($U64>=18,$U64<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$82))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,82,66)))))

def f751(E, r, c):
    # Ô AB64
    # = IF(A64="",FALSE(),AND($U64>=0,$U64<=120,LUU_NIEN!$BH$82=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,82,60),1.0)))

def f752(E, r, c):
    # Ô AC64
    # = IF(A64="",FALSE(),AND($U64>=0,$U64<=120,LUU_NIEN!$BI$82=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,82,61),1.0)))

def f753(E, r, c):
    # Ô AD64
    # = IF(A64="",FALSE(),AND($U64>=0,$U64<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$82))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,82,92)))))

def f754(E, r, c):
    # Ô A65
    # = IF((ISNUMBER(LUU_NIEN!$A$83)),LUU_NIEN!$A$83,"")
    return (E.v(15,83,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,83,1))) else '')

def f755(E, r, c):
    # Ô C65
    # = IF(A65="","",LUU_NIEN!$L$83)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,83,12))

def f756(E, r, c):
    # Ô D65
    # = IF(A65="","",LUU_NIEN!$H$83)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,83,8))

def f757(E, r, c):
    # Ô E65
    # = IF(A65="","",LUU_NIEN!$AA$83)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,83,27))

def f758(E, r, c):
    # Ô F65
    # = IF(A65="","",LUU_NIEN!$AL$83)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,83,38))

def f759(E, r, c):
    # Ô K65
    # = IF(A65="","",TRIM(IF(AB65,"Dấu hiệu bất lợi cho cha. ","")&IF(AC65,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA65,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U65>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$83))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,83,66))))) else ''))))

def f760(E, r, c):
    # Ô L65
    # = IF(OR(A65="",LUU_NIEN!$CN$83=""),"",IFERROR(LEFT(LUU_NIEN!$CN$83,FIND(":",LUU_NIEN!$CN$83)-1),LEFT(LUU_NIEN!$CN$83,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,83,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,83,92), X.sub(X.FIND(':', E.v(15,83,92)),1.0)), lambda: X.LEFT(E.v(15,83,92), 120.0)))

def f761(E, r, c):
    # Ô M65
    # = IF(A65="","",IF(LEN(LUU_NIEN!$BN$83)>260,LEFT(LUU_NIEN!$BN$83,260)&"…",LUU_NIEN!$BN$83))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,83,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,83,66)),260.0)) else E.v(15,83,66)))

def f762(E, r, c):
    # Ô T65
    # = IF(OR(A65="",LUU_NIEN!$P$83=""),-1,MOD(LUU_NIEN!$P$83-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,83,16),''))) else X.MOD(X.add(X.sub(E.v(15,83,16),E.v(23,7,5)),5.0), 5.0))

def f763(E, r, c):
    # Ô AA65
    # = IF(A65="",FALSE(),AND($U65>=18,$U65<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$83))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,83,66)))))

def f764(E, r, c):
    # Ô AB65
    # = IF(A65="",FALSE(),AND($U65>=0,$U65<=120,LUU_NIEN!$BH$83=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,83,60),1.0)))

def f765(E, r, c):
    # Ô AC65
    # = IF(A65="",FALSE(),AND($U65>=0,$U65<=120,LUU_NIEN!$BI$83=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,83,61),1.0)))

def f766(E, r, c):
    # Ô AD65
    # = IF(A65="",FALSE(),AND($U65>=0,$U65<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$83))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,83,92)))))

def f767(E, r, c):
    # Ô A66
    # = IF((ISNUMBER(LUU_NIEN!$A$84)),LUU_NIEN!$A$84,"")
    return (E.v(15,84,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,84,1))) else '')

def f768(E, r, c):
    # Ô C66
    # = IF(A66="","",LUU_NIEN!$L$84)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,84,12))

def f769(E, r, c):
    # Ô D66
    # = IF(A66="","",LUU_NIEN!$H$84)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,84,8))

def f770(E, r, c):
    # Ô E66
    # = IF(A66="","",LUU_NIEN!$AA$84)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,84,27))

def f771(E, r, c):
    # Ô F66
    # = IF(A66="","",LUU_NIEN!$AL$84)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,84,38))

def f772(E, r, c):
    # Ô K66
    # = IF(A66="","",TRIM(IF(AB66,"Dấu hiệu bất lợi cho cha. ","")&IF(AC66,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA66,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U66>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$84))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,84,66))))) else ''))))

def f773(E, r, c):
    # Ô L66
    # = IF(OR(A66="",LUU_NIEN!$CN$84=""),"",IFERROR(LEFT(LUU_NIEN!$CN$84,FIND(":",LUU_NIEN!$CN$84)-1),LEFT(LUU_NIEN!$CN$84,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,84,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,84,92), X.sub(X.FIND(':', E.v(15,84,92)),1.0)), lambda: X.LEFT(E.v(15,84,92), 120.0)))

def f774(E, r, c):
    # Ô M66
    # = IF(A66="","",IF(LEN(LUU_NIEN!$BN$84)>260,LEFT(LUU_NIEN!$BN$84,260)&"…",LUU_NIEN!$BN$84))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,84,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,84,66)),260.0)) else E.v(15,84,66)))

def f775(E, r, c):
    # Ô T66
    # = IF(OR(A66="",LUU_NIEN!$P$84=""),-1,MOD(LUU_NIEN!$P$84-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,84,16),''))) else X.MOD(X.add(X.sub(E.v(15,84,16),E.v(23,7,5)),5.0), 5.0))

def f776(E, r, c):
    # Ô AA66
    # = IF(A66="",FALSE(),AND($U66>=18,$U66<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$84))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,84,66)))))

def f777(E, r, c):
    # Ô AB66
    # = IF(A66="",FALSE(),AND($U66>=0,$U66<=120,LUU_NIEN!$BH$84=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,84,60),1.0)))

def f778(E, r, c):
    # Ô AC66
    # = IF(A66="",FALSE(),AND($U66>=0,$U66<=120,LUU_NIEN!$BI$84=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,84,61),1.0)))

def f779(E, r, c):
    # Ô AD66
    # = IF(A66="",FALSE(),AND($U66>=0,$U66<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$84))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,84,92)))))

def f780(E, r, c):
    # Ô A67
    # = IF((ISNUMBER(LUU_NIEN!$A$85)),LUU_NIEN!$A$85,"")
    return (E.v(15,85,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,85,1))) else '')

def f781(E, r, c):
    # Ô C67
    # = IF(A67="","",LUU_NIEN!$L$85)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,85,12))

def f782(E, r, c):
    # Ô D67
    # = IF(A67="","",LUU_NIEN!$H$85)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,85,8))

def f783(E, r, c):
    # Ô E67
    # = IF(A67="","",LUU_NIEN!$AA$85)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,85,27))

def f784(E, r, c):
    # Ô F67
    # = IF(A67="","",LUU_NIEN!$AL$85)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,85,38))

def f785(E, r, c):
    # Ô K67
    # = IF(A67="","",TRIM(IF(AB67,"Dấu hiệu bất lợi cho cha. ","")&IF(AC67,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA67,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U67>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$85))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,85,66))))) else ''))))

def f786(E, r, c):
    # Ô L67
    # = IF(OR(A67="",LUU_NIEN!$CN$85=""),"",IFERROR(LEFT(LUU_NIEN!$CN$85,FIND(":",LUU_NIEN!$CN$85)-1),LEFT(LUU_NIEN!$CN$85,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,85,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,85,92), X.sub(X.FIND(':', E.v(15,85,92)),1.0)), lambda: X.LEFT(E.v(15,85,92), 120.0)))

def f787(E, r, c):
    # Ô M67
    # = IF(A67="","",IF(LEN(LUU_NIEN!$BN$85)>260,LEFT(LUU_NIEN!$BN$85,260)&"…",LUU_NIEN!$BN$85))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,85,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,85,66)),260.0)) else E.v(15,85,66)))

def f788(E, r, c):
    # Ô T67
    # = IF(OR(A67="",LUU_NIEN!$P$85=""),-1,MOD(LUU_NIEN!$P$85-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,85,16),''))) else X.MOD(X.add(X.sub(E.v(15,85,16),E.v(23,7,5)),5.0), 5.0))

def f789(E, r, c):
    # Ô AA67
    # = IF(A67="",FALSE(),AND($U67>=18,$U67<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$85))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,85,66)))))

def f790(E, r, c):
    # Ô AB67
    # = IF(A67="",FALSE(),AND($U67>=0,$U67<=120,LUU_NIEN!$BH$85=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,85,60),1.0)))

def f791(E, r, c):
    # Ô AC67
    # = IF(A67="",FALSE(),AND($U67>=0,$U67<=120,LUU_NIEN!$BI$85=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,85,61),1.0)))

def f792(E, r, c):
    # Ô AD67
    # = IF(A67="",FALSE(),AND($U67>=0,$U67<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$85))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,85,92)))))

def f793(E, r, c):
    # Ô A68
    # = IF((ISNUMBER(LUU_NIEN!$A$86)),LUU_NIEN!$A$86,"")
    return (E.v(15,86,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,86,1))) else '')

def f794(E, r, c):
    # Ô C68
    # = IF(A68="","",LUU_NIEN!$L$86)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,86,12))

def f795(E, r, c):
    # Ô D68
    # = IF(A68="","",LUU_NIEN!$H$86)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,86,8))

def f796(E, r, c):
    # Ô E68
    # = IF(A68="","",LUU_NIEN!$AA$86)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,86,27))

def f797(E, r, c):
    # Ô F68
    # = IF(A68="","",LUU_NIEN!$AL$86)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,86,38))

def f798(E, r, c):
    # Ô K68
    # = IF(A68="","",TRIM(IF(AB68,"Dấu hiệu bất lợi cho cha. ","")&IF(AC68,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA68,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U68>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$86))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,86,66))))) else ''))))

def f799(E, r, c):
    # Ô L68
    # = IF(OR(A68="",LUU_NIEN!$CN$86=""),"",IFERROR(LEFT(LUU_NIEN!$CN$86,FIND(":",LUU_NIEN!$CN$86)-1),LEFT(LUU_NIEN!$CN$86,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,86,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,86,92), X.sub(X.FIND(':', E.v(15,86,92)),1.0)), lambda: X.LEFT(E.v(15,86,92), 120.0)))

def f800(E, r, c):
    # Ô M68
    # = IF(A68="","",IF(LEN(LUU_NIEN!$BN$86)>260,LEFT(LUU_NIEN!$BN$86,260)&"…",LUU_NIEN!$BN$86))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,86,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,86,66)),260.0)) else E.v(15,86,66)))

def f801(E, r, c):
    # Ô T68
    # = IF(OR(A68="",LUU_NIEN!$P$86=""),-1,MOD(LUU_NIEN!$P$86-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,86,16),''))) else X.MOD(X.add(X.sub(E.v(15,86,16),E.v(23,7,5)),5.0), 5.0))

def f802(E, r, c):
    # Ô AA68
    # = IF(A68="",FALSE(),AND($U68>=18,$U68<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$86))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,86,66)))))

def f803(E, r, c):
    # Ô AB68
    # = IF(A68="",FALSE(),AND($U68>=0,$U68<=120,LUU_NIEN!$BH$86=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,86,60),1.0)))

def f804(E, r, c):
    # Ô AC68
    # = IF(A68="",FALSE(),AND($U68>=0,$U68<=120,LUU_NIEN!$BI$86=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,86,61),1.0)))

def f805(E, r, c):
    # Ô AD68
    # = IF(A68="",FALSE(),AND($U68>=0,$U68<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$86))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,86,92)))))

def f806(E, r, c):
    # Ô A69
    # = IF((ISNUMBER(LUU_NIEN!$A$87)),LUU_NIEN!$A$87,"")
    return (E.v(15,87,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,87,1))) else '')

def f807(E, r, c):
    # Ô C69
    # = IF(A69="","",LUU_NIEN!$L$87)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,87,12))

def f808(E, r, c):
    # Ô D69
    # = IF(A69="","",LUU_NIEN!$H$87)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,87,8))

def f809(E, r, c):
    # Ô E69
    # = IF(A69="","",LUU_NIEN!$AA$87)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,87,27))

def f810(E, r, c):
    # Ô F69
    # = IF(A69="","",LUU_NIEN!$AL$87)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,87,38))

def f811(E, r, c):
    # Ô K69
    # = IF(A69="","",TRIM(IF(AB69,"Dấu hiệu bất lợi cho cha. ","")&IF(AC69,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA69,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U69>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$87))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,87,66))))) else ''))))

def f812(E, r, c):
    # Ô L69
    # = IF(OR(A69="",LUU_NIEN!$CN$87=""),"",IFERROR(LEFT(LUU_NIEN!$CN$87,FIND(":",LUU_NIEN!$CN$87)-1),LEFT(LUU_NIEN!$CN$87,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,87,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,87,92), X.sub(X.FIND(':', E.v(15,87,92)),1.0)), lambda: X.LEFT(E.v(15,87,92), 120.0)))

def f813(E, r, c):
    # Ô M69
    # = IF(A69="","",IF(LEN(LUU_NIEN!$BN$87)>260,LEFT(LUU_NIEN!$BN$87,260)&"…",LUU_NIEN!$BN$87))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,87,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,87,66)),260.0)) else E.v(15,87,66)))

def f814(E, r, c):
    # Ô T69
    # = IF(OR(A69="",LUU_NIEN!$P$87=""),-1,MOD(LUU_NIEN!$P$87-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,87,16),''))) else X.MOD(X.add(X.sub(E.v(15,87,16),E.v(23,7,5)),5.0), 5.0))

def f815(E, r, c):
    # Ô AA69
    # = IF(A69="",FALSE(),AND($U69>=18,$U69<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$87))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,87,66)))))

def f816(E, r, c):
    # Ô AB69
    # = IF(A69="",FALSE(),AND($U69>=0,$U69<=120,LUU_NIEN!$BH$87=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,87,60),1.0)))

def f817(E, r, c):
    # Ô AC69
    # = IF(A69="",FALSE(),AND($U69>=0,$U69<=120,LUU_NIEN!$BI$87=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,87,61),1.0)))

def f818(E, r, c):
    # Ô AD69
    # = IF(A69="",FALSE(),AND($U69>=0,$U69<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$87))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,87,92)))))

def f819(E, r, c):
    # Ô A70
    # = IF((ISNUMBER(LUU_NIEN!$A$88)),LUU_NIEN!$A$88,"")
    return (E.v(15,88,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,88,1))) else '')

def f820(E, r, c):
    # Ô C70
    # = IF(A70="","",LUU_NIEN!$L$88)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,88,12))

def f821(E, r, c):
    # Ô D70
    # = IF(A70="","",LUU_NIEN!$H$88)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,88,8))

def f822(E, r, c):
    # Ô E70
    # = IF(A70="","",LUU_NIEN!$AA$88)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,88,27))

def f823(E, r, c):
    # Ô F70
    # = IF(A70="","",LUU_NIEN!$AL$88)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,88,38))

def f824(E, r, c):
    # Ô K70
    # = IF(A70="","",TRIM(IF(AB70,"Dấu hiệu bất lợi cho cha. ","")&IF(AC70,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA70,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U70>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$88))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,88,66))))) else ''))))

def f825(E, r, c):
    # Ô L70
    # = IF(OR(A70="",LUU_NIEN!$CN$88=""),"",IFERROR(LEFT(LUU_NIEN!$CN$88,FIND(":",LUU_NIEN!$CN$88)-1),LEFT(LUU_NIEN!$CN$88,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,88,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,88,92), X.sub(X.FIND(':', E.v(15,88,92)),1.0)), lambda: X.LEFT(E.v(15,88,92), 120.0)))

def f826(E, r, c):
    # Ô M70
    # = IF(A70="","",IF(LEN(LUU_NIEN!$BN$88)>260,LEFT(LUU_NIEN!$BN$88,260)&"…",LUU_NIEN!$BN$88))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,88,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,88,66)),260.0)) else E.v(15,88,66)))

def f827(E, r, c):
    # Ô T70
    # = IF(OR(A70="",LUU_NIEN!$P$88=""),-1,MOD(LUU_NIEN!$P$88-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,88,16),''))) else X.MOD(X.add(X.sub(E.v(15,88,16),E.v(23,7,5)),5.0), 5.0))

def f828(E, r, c):
    # Ô AA70
    # = IF(A70="",FALSE(),AND($U70>=18,$U70<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$88))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,88,66)))))

def f829(E, r, c):
    # Ô AB70
    # = IF(A70="",FALSE(),AND($U70>=0,$U70<=120,LUU_NIEN!$BH$88=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,88,60),1.0)))

def f830(E, r, c):
    # Ô AC70
    # = IF(A70="",FALSE(),AND($U70>=0,$U70<=120,LUU_NIEN!$BI$88=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,88,61),1.0)))

def f831(E, r, c):
    # Ô AD70
    # = IF(A70="",FALSE(),AND($U70>=0,$U70<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$88))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,88,92)))))

def f832(E, r, c):
    # Ô A71
    # = IF((ISNUMBER(LUU_NIEN!$A$89)),LUU_NIEN!$A$89,"")
    return (E.v(15,89,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,89,1))) else '')

def f833(E, r, c):
    # Ô C71
    # = IF(A71="","",LUU_NIEN!$L$89)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,89,12))

def f834(E, r, c):
    # Ô D71
    # = IF(A71="","",LUU_NIEN!$H$89)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,89,8))

def f835(E, r, c):
    # Ô E71
    # = IF(A71="","",LUU_NIEN!$AA$89)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,89,27))

def f836(E, r, c):
    # Ô F71
    # = IF(A71="","",LUU_NIEN!$AL$89)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,89,38))

def f837(E, r, c):
    # Ô K71
    # = IF(A71="","",TRIM(IF(AB71,"Dấu hiệu bất lợi cho cha. ","")&IF(AC71,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA71,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U71>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$89))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,89,66))))) else ''))))

def f838(E, r, c):
    # Ô L71
    # = IF(OR(A71="",LUU_NIEN!$CN$89=""),"",IFERROR(LEFT(LUU_NIEN!$CN$89,FIND(":",LUU_NIEN!$CN$89)-1),LEFT(LUU_NIEN!$CN$89,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,89,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,89,92), X.sub(X.FIND(':', E.v(15,89,92)),1.0)), lambda: X.LEFT(E.v(15,89,92), 120.0)))

def f839(E, r, c):
    # Ô M71
    # = IF(A71="","",IF(LEN(LUU_NIEN!$BN$89)>260,LEFT(LUU_NIEN!$BN$89,260)&"…",LUU_NIEN!$BN$89))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,89,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,89,66)),260.0)) else E.v(15,89,66)))

def f840(E, r, c):
    # Ô T71
    # = IF(OR(A71="",LUU_NIEN!$P$89=""),-1,MOD(LUU_NIEN!$P$89-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,89,16),''))) else X.MOD(X.add(X.sub(E.v(15,89,16),E.v(23,7,5)),5.0), 5.0))

def f841(E, r, c):
    # Ô AA71
    # = IF(A71="",FALSE(),AND($U71>=18,$U71<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$89))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,89,66)))))

def f842(E, r, c):
    # Ô AB71
    # = IF(A71="",FALSE(),AND($U71>=0,$U71<=120,LUU_NIEN!$BH$89=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,89,60),1.0)))

def f843(E, r, c):
    # Ô AC71
    # = IF(A71="",FALSE(),AND($U71>=0,$U71<=120,LUU_NIEN!$BI$89=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,89,61),1.0)))

def f844(E, r, c):
    # Ô AD71
    # = IF(A71="",FALSE(),AND($U71>=0,$U71<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$89))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,89,92)))))

def f845(E, r, c):
    # Ô A72
    # = IF((ISNUMBER(LUU_NIEN!$A$90)),LUU_NIEN!$A$90,"")
    return (E.v(15,90,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,90,1))) else '')

def f846(E, r, c):
    # Ô C72
    # = IF(A72="","",LUU_NIEN!$L$90)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,90,12))

def f847(E, r, c):
    # Ô D72
    # = IF(A72="","",LUU_NIEN!$H$90)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,90,8))

def f848(E, r, c):
    # Ô E72
    # = IF(A72="","",LUU_NIEN!$AA$90)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,90,27))

def f849(E, r, c):
    # Ô F72
    # = IF(A72="","",LUU_NIEN!$AL$90)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,90,38))

def f850(E, r, c):
    # Ô K72
    # = IF(A72="","",TRIM(IF(AB72,"Dấu hiệu bất lợi cho cha. ","")&IF(AC72,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA72,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U72>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$90))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,90,66))))) else ''))))

def f851(E, r, c):
    # Ô L72
    # = IF(OR(A72="",LUU_NIEN!$CN$90=""),"",IFERROR(LEFT(LUU_NIEN!$CN$90,FIND(":",LUU_NIEN!$CN$90)-1),LEFT(LUU_NIEN!$CN$90,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,90,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,90,92), X.sub(X.FIND(':', E.v(15,90,92)),1.0)), lambda: X.LEFT(E.v(15,90,92), 120.0)))

def f852(E, r, c):
    # Ô M72
    # = IF(A72="","",IF(LEN(LUU_NIEN!$BN$90)>260,LEFT(LUU_NIEN!$BN$90,260)&"…",LUU_NIEN!$BN$90))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,90,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,90,66)),260.0)) else E.v(15,90,66)))

def f853(E, r, c):
    # Ô T72
    # = IF(OR(A72="",LUU_NIEN!$P$90=""),-1,MOD(LUU_NIEN!$P$90-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,90,16),''))) else X.MOD(X.add(X.sub(E.v(15,90,16),E.v(23,7,5)),5.0), 5.0))

def f854(E, r, c):
    # Ô AA72
    # = IF(A72="",FALSE(),AND($U72>=18,$U72<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$90))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,90,66)))))

def f855(E, r, c):
    # Ô AB72
    # = IF(A72="",FALSE(),AND($U72>=0,$U72<=120,LUU_NIEN!$BH$90=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,90,60),1.0)))

def f856(E, r, c):
    # Ô AC72
    # = IF(A72="",FALSE(),AND($U72>=0,$U72<=120,LUU_NIEN!$BI$90=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,90,61),1.0)))

def f857(E, r, c):
    # Ô AD72
    # = IF(A72="",FALSE(),AND($U72>=0,$U72<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$90))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,90,92)))))

def f858(E, r, c):
    # Ô A73
    # = IF((ISNUMBER(LUU_NIEN!$A$91)),LUU_NIEN!$A$91,"")
    return (E.v(15,91,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,91,1))) else '')

def f859(E, r, c):
    # Ô C73
    # = IF(A73="","",LUU_NIEN!$L$91)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,91,12))

def f860(E, r, c):
    # Ô D73
    # = IF(A73="","",LUU_NIEN!$H$91)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,91,8))

def f861(E, r, c):
    # Ô E73
    # = IF(A73="","",LUU_NIEN!$AA$91)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,91,27))

def f862(E, r, c):
    # Ô F73
    # = IF(A73="","",LUU_NIEN!$AL$91)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,91,38))

def f863(E, r, c):
    # Ô K73
    # = IF(A73="","",TRIM(IF(AB73,"Dấu hiệu bất lợi cho cha. ","")&IF(AC73,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA73,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U73>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$91))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,91,66))))) else ''))))

def f864(E, r, c):
    # Ô L73
    # = IF(OR(A73="",LUU_NIEN!$CN$91=""),"",IFERROR(LEFT(LUU_NIEN!$CN$91,FIND(":",LUU_NIEN!$CN$91)-1),LEFT(LUU_NIEN!$CN$91,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,91,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,91,92), X.sub(X.FIND(':', E.v(15,91,92)),1.0)), lambda: X.LEFT(E.v(15,91,92), 120.0)))

def f865(E, r, c):
    # Ô M73
    # = IF(A73="","",IF(LEN(LUU_NIEN!$BN$91)>260,LEFT(LUU_NIEN!$BN$91,260)&"…",LUU_NIEN!$BN$91))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,91,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,91,66)),260.0)) else E.v(15,91,66)))

def f866(E, r, c):
    # Ô T73
    # = IF(OR(A73="",LUU_NIEN!$P$91=""),-1,MOD(LUU_NIEN!$P$91-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,91,16),''))) else X.MOD(X.add(X.sub(E.v(15,91,16),E.v(23,7,5)),5.0), 5.0))

def f867(E, r, c):
    # Ô AA73
    # = IF(A73="",FALSE(),AND($U73>=18,$U73<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$91))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,91,66)))))

def f868(E, r, c):
    # Ô AB73
    # = IF(A73="",FALSE(),AND($U73>=0,$U73<=120,LUU_NIEN!$BH$91=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,91,60),1.0)))

def f869(E, r, c):
    # Ô AC73
    # = IF(A73="",FALSE(),AND($U73>=0,$U73<=120,LUU_NIEN!$BI$91=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,91,61),1.0)))

def f870(E, r, c):
    # Ô AD73
    # = IF(A73="",FALSE(),AND($U73>=0,$U73<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$91))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,91,92)))))

def f871(E, r, c):
    # Ô A74
    # = IF((ISNUMBER(LUU_NIEN!$A$92)),LUU_NIEN!$A$92,"")
    return (E.v(15,92,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,92,1))) else '')

def f872(E, r, c):
    # Ô C74
    # = IF(A74="","",LUU_NIEN!$L$92)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,92,12))

def f873(E, r, c):
    # Ô D74
    # = IF(A74="","",LUU_NIEN!$H$92)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,92,8))

def f874(E, r, c):
    # Ô E74
    # = IF(A74="","",LUU_NIEN!$AA$92)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,92,27))

def f875(E, r, c):
    # Ô F74
    # = IF(A74="","",LUU_NIEN!$AL$92)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,92,38))

def f876(E, r, c):
    # Ô K74
    # = IF(A74="","",TRIM(IF(AB74,"Dấu hiệu bất lợi cho cha. ","")&IF(AC74,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA74,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U74>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$92))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,92,66))))) else ''))))

def f877(E, r, c):
    # Ô L74
    # = IF(OR(A74="",LUU_NIEN!$CN$92=""),"",IFERROR(LEFT(LUU_NIEN!$CN$92,FIND(":",LUU_NIEN!$CN$92)-1),LEFT(LUU_NIEN!$CN$92,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,92,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,92,92), X.sub(X.FIND(':', E.v(15,92,92)),1.0)), lambda: X.LEFT(E.v(15,92,92), 120.0)))

def f878(E, r, c):
    # Ô M74
    # = IF(A74="","",IF(LEN(LUU_NIEN!$BN$92)>260,LEFT(LUU_NIEN!$BN$92,260)&"…",LUU_NIEN!$BN$92))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,92,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,92,66)),260.0)) else E.v(15,92,66)))

def f879(E, r, c):
    # Ô T74
    # = IF(OR(A74="",LUU_NIEN!$P$92=""),-1,MOD(LUU_NIEN!$P$92-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,92,16),''))) else X.MOD(X.add(X.sub(E.v(15,92,16),E.v(23,7,5)),5.0), 5.0))

def f880(E, r, c):
    # Ô AA74
    # = IF(A74="",FALSE(),AND($U74>=18,$U74<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$92))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,92,66)))))

def f881(E, r, c):
    # Ô AB74
    # = IF(A74="",FALSE(),AND($U74>=0,$U74<=120,LUU_NIEN!$BH$92=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,92,60),1.0)))

def f882(E, r, c):
    # Ô AC74
    # = IF(A74="",FALSE(),AND($U74>=0,$U74<=120,LUU_NIEN!$BI$92=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,92,61),1.0)))

def f883(E, r, c):
    # Ô AD74
    # = IF(A74="",FALSE(),AND($U74>=0,$U74<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$92))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,92,92)))))

def f884(E, r, c):
    # Ô A75
    # = IF((ISNUMBER(LUU_NIEN!$A$93)),LUU_NIEN!$A$93,"")
    return (E.v(15,93,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,93,1))) else '')

def f885(E, r, c):
    # Ô C75
    # = IF(A75="","",LUU_NIEN!$L$93)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,93,12))

def f886(E, r, c):
    # Ô D75
    # = IF(A75="","",LUU_NIEN!$H$93)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,93,8))

def f887(E, r, c):
    # Ô E75
    # = IF(A75="","",LUU_NIEN!$AA$93)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,93,27))

def f888(E, r, c):
    # Ô F75
    # = IF(A75="","",LUU_NIEN!$AL$93)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,93,38))

def f889(E, r, c):
    # Ô K75
    # = IF(A75="","",TRIM(IF(AB75,"Dấu hiệu bất lợi cho cha. ","")&IF(AC75,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA75,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U75>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$93))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,93,66))))) else ''))))

def f890(E, r, c):
    # Ô L75
    # = IF(OR(A75="",LUU_NIEN!$CN$93=""),"",IFERROR(LEFT(LUU_NIEN!$CN$93,FIND(":",LUU_NIEN!$CN$93)-1),LEFT(LUU_NIEN!$CN$93,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,93,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,93,92), X.sub(X.FIND(':', E.v(15,93,92)),1.0)), lambda: X.LEFT(E.v(15,93,92), 120.0)))

def f891(E, r, c):
    # Ô M75
    # = IF(A75="","",IF(LEN(LUU_NIEN!$BN$93)>260,LEFT(LUU_NIEN!$BN$93,260)&"…",LUU_NIEN!$BN$93))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,93,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,93,66)),260.0)) else E.v(15,93,66)))

def f892(E, r, c):
    # Ô T75
    # = IF(OR(A75="",LUU_NIEN!$P$93=""),-1,MOD(LUU_NIEN!$P$93-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,93,16),''))) else X.MOD(X.add(X.sub(E.v(15,93,16),E.v(23,7,5)),5.0), 5.0))

def f893(E, r, c):
    # Ô AA75
    # = IF(A75="",FALSE(),AND($U75>=18,$U75<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$93))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,93,66)))))

def f894(E, r, c):
    # Ô AB75
    # = IF(A75="",FALSE(),AND($U75>=0,$U75<=120,LUU_NIEN!$BH$93=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,93,60),1.0)))

def f895(E, r, c):
    # Ô AC75
    # = IF(A75="",FALSE(),AND($U75>=0,$U75<=120,LUU_NIEN!$BI$93=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,93,61),1.0)))

def f896(E, r, c):
    # Ô AD75
    # = IF(A75="",FALSE(),AND($U75>=0,$U75<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$93))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,93,92)))))

def f897(E, r, c):
    # Ô A76
    # = IF((ISNUMBER(LUU_NIEN!$A$94)),LUU_NIEN!$A$94,"")
    return (E.v(15,94,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,94,1))) else '')

def f898(E, r, c):
    # Ô C76
    # = IF(A76="","",LUU_NIEN!$L$94)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,94,12))

def f899(E, r, c):
    # Ô D76
    # = IF(A76="","",LUU_NIEN!$H$94)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,94,8))

def f900(E, r, c):
    # Ô E76
    # = IF(A76="","",LUU_NIEN!$AA$94)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,94,27))

def f901(E, r, c):
    # Ô F76
    # = IF(A76="","",LUU_NIEN!$AL$94)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,94,38))

def f902(E, r, c):
    # Ô K76
    # = IF(A76="","",TRIM(IF(AB76,"Dấu hiệu bất lợi cho cha. ","")&IF(AC76,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA76,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U76>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$94))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,94,66))))) else ''))))

def f903(E, r, c):
    # Ô L76
    # = IF(OR(A76="",LUU_NIEN!$CN$94=""),"",IFERROR(LEFT(LUU_NIEN!$CN$94,FIND(":",LUU_NIEN!$CN$94)-1),LEFT(LUU_NIEN!$CN$94,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,94,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,94,92), X.sub(X.FIND(':', E.v(15,94,92)),1.0)), lambda: X.LEFT(E.v(15,94,92), 120.0)))

def f904(E, r, c):
    # Ô M76
    # = IF(A76="","",IF(LEN(LUU_NIEN!$BN$94)>260,LEFT(LUU_NIEN!$BN$94,260)&"…",LUU_NIEN!$BN$94))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,94,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,94,66)),260.0)) else E.v(15,94,66)))

def f905(E, r, c):
    # Ô T76
    # = IF(OR(A76="",LUU_NIEN!$P$94=""),-1,MOD(LUU_NIEN!$P$94-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,94,16),''))) else X.MOD(X.add(X.sub(E.v(15,94,16),E.v(23,7,5)),5.0), 5.0))

def f906(E, r, c):
    # Ô AA76
    # = IF(A76="",FALSE(),AND($U76>=18,$U76<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$94))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,94,66)))))

def f907(E, r, c):
    # Ô AB76
    # = IF(A76="",FALSE(),AND($U76>=0,$U76<=120,LUU_NIEN!$BH$94=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,94,60),1.0)))

def f908(E, r, c):
    # Ô AC76
    # = IF(A76="",FALSE(),AND($U76>=0,$U76<=120,LUU_NIEN!$BI$94=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,94,61),1.0)))

def f909(E, r, c):
    # Ô AD76
    # = IF(A76="",FALSE(),AND($U76>=0,$U76<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$94))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,94,92)))))

def f910(E, r, c):
    # Ô A77
    # = IF((ISNUMBER(LUU_NIEN!$A$95)),LUU_NIEN!$A$95,"")
    return (E.v(15,95,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,95,1))) else '')

def f911(E, r, c):
    # Ô C77
    # = IF(A77="","",LUU_NIEN!$L$95)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,95,12))

def f912(E, r, c):
    # Ô D77
    # = IF(A77="","",LUU_NIEN!$H$95)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,95,8))

def f913(E, r, c):
    # Ô E77
    # = IF(A77="","",LUU_NIEN!$AA$95)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,95,27))

def f914(E, r, c):
    # Ô F77
    # = IF(A77="","",LUU_NIEN!$AL$95)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,95,38))

def f915(E, r, c):
    # Ô K77
    # = IF(A77="","",TRIM(IF(AB77,"Dấu hiệu bất lợi cho cha. ","")&IF(AC77,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA77,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U77>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$95))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,95,66))))) else ''))))

def f916(E, r, c):
    # Ô L77
    # = IF(OR(A77="",LUU_NIEN!$CN$95=""),"",IFERROR(LEFT(LUU_NIEN!$CN$95,FIND(":",LUU_NIEN!$CN$95)-1),LEFT(LUU_NIEN!$CN$95,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,95,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,95,92), X.sub(X.FIND(':', E.v(15,95,92)),1.0)), lambda: X.LEFT(E.v(15,95,92), 120.0)))

def f917(E, r, c):
    # Ô M77
    # = IF(A77="","",IF(LEN(LUU_NIEN!$BN$95)>260,LEFT(LUU_NIEN!$BN$95,260)&"…",LUU_NIEN!$BN$95))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,95,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,95,66)),260.0)) else E.v(15,95,66)))

def f918(E, r, c):
    # Ô T77
    # = IF(OR(A77="",LUU_NIEN!$P$95=""),-1,MOD(LUU_NIEN!$P$95-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,95,16),''))) else X.MOD(X.add(X.sub(E.v(15,95,16),E.v(23,7,5)),5.0), 5.0))

def f919(E, r, c):
    # Ô AA77
    # = IF(A77="",FALSE(),AND($U77>=18,$U77<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$95))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,95,66)))))

def f920(E, r, c):
    # Ô AB77
    # = IF(A77="",FALSE(),AND($U77>=0,$U77<=120,LUU_NIEN!$BH$95=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,95,60),1.0)))

def f921(E, r, c):
    # Ô AC77
    # = IF(A77="",FALSE(),AND($U77>=0,$U77<=120,LUU_NIEN!$BI$95=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,95,61),1.0)))

def f922(E, r, c):
    # Ô AD77
    # = IF(A77="",FALSE(),AND($U77>=0,$U77<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$95))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,95,92)))))

def f923(E, r, c):
    # Ô A78
    # = IF((ISNUMBER(LUU_NIEN!$A$96)),LUU_NIEN!$A$96,"")
    return (E.v(15,96,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,96,1))) else '')

def f924(E, r, c):
    # Ô C78
    # = IF(A78="","",LUU_NIEN!$L$96)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,96,12))

def f925(E, r, c):
    # Ô D78
    # = IF(A78="","",LUU_NIEN!$H$96)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,96,8))

def f926(E, r, c):
    # Ô E78
    # = IF(A78="","",LUU_NIEN!$AA$96)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,96,27))

def f927(E, r, c):
    # Ô F78
    # = IF(A78="","",LUU_NIEN!$AL$96)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,96,38))

def f928(E, r, c):
    # Ô K78
    # = IF(A78="","",TRIM(IF(AB78,"Dấu hiệu bất lợi cho cha. ","")&IF(AC78,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA78,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U78>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$96))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,96,66))))) else ''))))

def f929(E, r, c):
    # Ô L78
    # = IF(OR(A78="",LUU_NIEN!$CN$96=""),"",IFERROR(LEFT(LUU_NIEN!$CN$96,FIND(":",LUU_NIEN!$CN$96)-1),LEFT(LUU_NIEN!$CN$96,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,96,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,96,92), X.sub(X.FIND(':', E.v(15,96,92)),1.0)), lambda: X.LEFT(E.v(15,96,92), 120.0)))

def f930(E, r, c):
    # Ô M78
    # = IF(A78="","",IF(LEN(LUU_NIEN!$BN$96)>260,LEFT(LUU_NIEN!$BN$96,260)&"…",LUU_NIEN!$BN$96))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,96,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,96,66)),260.0)) else E.v(15,96,66)))

def f931(E, r, c):
    # Ô T78
    # = IF(OR(A78="",LUU_NIEN!$P$96=""),-1,MOD(LUU_NIEN!$P$96-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,96,16),''))) else X.MOD(X.add(X.sub(E.v(15,96,16),E.v(23,7,5)),5.0), 5.0))

def f932(E, r, c):
    # Ô AA78
    # = IF(A78="",FALSE(),AND($U78>=18,$U78<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$96))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,96,66)))))

def f933(E, r, c):
    # Ô AB78
    # = IF(A78="",FALSE(),AND($U78>=0,$U78<=120,LUU_NIEN!$BH$96=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,96,60),1.0)))

def f934(E, r, c):
    # Ô AC78
    # = IF(A78="",FALSE(),AND($U78>=0,$U78<=120,LUU_NIEN!$BI$96=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,96,61),1.0)))

def f935(E, r, c):
    # Ô AD78
    # = IF(A78="",FALSE(),AND($U78>=0,$U78<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$96))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,96,92)))))

def f936(E, r, c):
    # Ô A79
    # = IF((ISNUMBER(LUU_NIEN!$A$97)),LUU_NIEN!$A$97,"")
    return (E.v(15,97,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,97,1))) else '')

def f937(E, r, c):
    # Ô C79
    # = IF(A79="","",LUU_NIEN!$L$97)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,97,12))

def f938(E, r, c):
    # Ô D79
    # = IF(A79="","",LUU_NIEN!$H$97)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,97,8))

def f939(E, r, c):
    # Ô E79
    # = IF(A79="","",LUU_NIEN!$AA$97)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,97,27))

def f940(E, r, c):
    # Ô F79
    # = IF(A79="","",LUU_NIEN!$AL$97)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,97,38))

def f941(E, r, c):
    # Ô K79
    # = IF(A79="","",TRIM(IF(AB79,"Dấu hiệu bất lợi cho cha. ","")&IF(AC79,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA79,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U79>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$97))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,97,66))))) else ''))))

def f942(E, r, c):
    # Ô L79
    # = IF(OR(A79="",LUU_NIEN!$CN$97=""),"",IFERROR(LEFT(LUU_NIEN!$CN$97,FIND(":",LUU_NIEN!$CN$97)-1),LEFT(LUU_NIEN!$CN$97,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,97,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,97,92), X.sub(X.FIND(':', E.v(15,97,92)),1.0)), lambda: X.LEFT(E.v(15,97,92), 120.0)))

def f943(E, r, c):
    # Ô M79
    # = IF(A79="","",IF(LEN(LUU_NIEN!$BN$97)>260,LEFT(LUU_NIEN!$BN$97,260)&"…",LUU_NIEN!$BN$97))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,97,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,97,66)),260.0)) else E.v(15,97,66)))

def f944(E, r, c):
    # Ô T79
    # = IF(OR(A79="",LUU_NIEN!$P$97=""),-1,MOD(LUU_NIEN!$P$97-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,97,16),''))) else X.MOD(X.add(X.sub(E.v(15,97,16),E.v(23,7,5)),5.0), 5.0))

def f945(E, r, c):
    # Ô AA79
    # = IF(A79="",FALSE(),AND($U79>=18,$U79<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$97))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,97,66)))))

def f946(E, r, c):
    # Ô AB79
    # = IF(A79="",FALSE(),AND($U79>=0,$U79<=120,LUU_NIEN!$BH$97=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,97,60),1.0)))

def f947(E, r, c):
    # Ô AC79
    # = IF(A79="",FALSE(),AND($U79>=0,$U79<=120,LUU_NIEN!$BI$97=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,97,61),1.0)))

def f948(E, r, c):
    # Ô AD79
    # = IF(A79="",FALSE(),AND($U79>=0,$U79<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$97))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,97,92)))))

def f949(E, r, c):
    # Ô A80
    # = IF((ISNUMBER(LUU_NIEN!$A$98)),LUU_NIEN!$A$98,"")
    return (E.v(15,98,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,98,1))) else '')

def f950(E, r, c):
    # Ô C80
    # = IF(A80="","",LUU_NIEN!$L$98)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,98,12))

def f951(E, r, c):
    # Ô D80
    # = IF(A80="","",LUU_NIEN!$H$98)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,98,8))

def f952(E, r, c):
    # Ô E80
    # = IF(A80="","",LUU_NIEN!$AA$98)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,98,27))

def f953(E, r, c):
    # Ô F80
    # = IF(A80="","",LUU_NIEN!$AL$98)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,98,38))

def f954(E, r, c):
    # Ô K80
    # = IF(A80="","",TRIM(IF(AB80,"Dấu hiệu bất lợi cho cha. ","")&IF(AC80,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA80,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U80>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$98))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,98,66))))) else ''))))

def f955(E, r, c):
    # Ô L80
    # = IF(OR(A80="",LUU_NIEN!$CN$98=""),"",IFERROR(LEFT(LUU_NIEN!$CN$98,FIND(":",LUU_NIEN!$CN$98)-1),LEFT(LUU_NIEN!$CN$98,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,98,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,98,92), X.sub(X.FIND(':', E.v(15,98,92)),1.0)), lambda: X.LEFT(E.v(15,98,92), 120.0)))

def f956(E, r, c):
    # Ô M80
    # = IF(A80="","",IF(LEN(LUU_NIEN!$BN$98)>260,LEFT(LUU_NIEN!$BN$98,260)&"…",LUU_NIEN!$BN$98))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,98,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,98,66)),260.0)) else E.v(15,98,66)))

def f957(E, r, c):
    # Ô T80
    # = IF(OR(A80="",LUU_NIEN!$P$98=""),-1,MOD(LUU_NIEN!$P$98-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,98,16),''))) else X.MOD(X.add(X.sub(E.v(15,98,16),E.v(23,7,5)),5.0), 5.0))

def f958(E, r, c):
    # Ô AA80
    # = IF(A80="",FALSE(),AND($U80>=18,$U80<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$98))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,98,66)))))

def f959(E, r, c):
    # Ô AB80
    # = IF(A80="",FALSE(),AND($U80>=0,$U80<=120,LUU_NIEN!$BH$98=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,98,60),1.0)))

def f960(E, r, c):
    # Ô AC80
    # = IF(A80="",FALSE(),AND($U80>=0,$U80<=120,LUU_NIEN!$BI$98=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,98,61),1.0)))

def f961(E, r, c):
    # Ô AD80
    # = IF(A80="",FALSE(),AND($U80>=0,$U80<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$98))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,98,92)))))

def f962(E, r, c):
    # Ô A81
    # = IF((ISNUMBER(LUU_NIEN!$A$99)),LUU_NIEN!$A$99,"")
    return (E.v(15,99,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,99,1))) else '')

def f963(E, r, c):
    # Ô C81
    # = IF(A81="","",LUU_NIEN!$L$99)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,99,12))

def f964(E, r, c):
    # Ô D81
    # = IF(A81="","",LUU_NIEN!$H$99)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,99,8))

def f965(E, r, c):
    # Ô E81
    # = IF(A81="","",LUU_NIEN!$AA$99)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,99,27))

def f966(E, r, c):
    # Ô F81
    # = IF(A81="","",LUU_NIEN!$AL$99)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,99,38))

def f967(E, r, c):
    # Ô K81
    # = IF(A81="","",TRIM(IF(AB81,"Dấu hiệu bất lợi cho cha. ","")&IF(AC81,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA81,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U81>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$99))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,99,66))))) else ''))))

def f968(E, r, c):
    # Ô L81
    # = IF(OR(A81="",LUU_NIEN!$CN$99=""),"",IFERROR(LEFT(LUU_NIEN!$CN$99,FIND(":",LUU_NIEN!$CN$99)-1),LEFT(LUU_NIEN!$CN$99,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,99,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,99,92), X.sub(X.FIND(':', E.v(15,99,92)),1.0)), lambda: X.LEFT(E.v(15,99,92), 120.0)))

def f969(E, r, c):
    # Ô M81
    # = IF(A81="","",IF(LEN(LUU_NIEN!$BN$99)>260,LEFT(LUU_NIEN!$BN$99,260)&"…",LUU_NIEN!$BN$99))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,99,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,99,66)),260.0)) else E.v(15,99,66)))

def f970(E, r, c):
    # Ô T81
    # = IF(OR(A81="",LUU_NIEN!$P$99=""),-1,MOD(LUU_NIEN!$P$99-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,99,16),''))) else X.MOD(X.add(X.sub(E.v(15,99,16),E.v(23,7,5)),5.0), 5.0))

def f971(E, r, c):
    # Ô AA81
    # = IF(A81="",FALSE(),AND($U81>=18,$U81<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$99))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,99,66)))))

def f972(E, r, c):
    # Ô AB81
    # = IF(A81="",FALSE(),AND($U81>=0,$U81<=120,LUU_NIEN!$BH$99=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,99,60),1.0)))

def f973(E, r, c):
    # Ô AC81
    # = IF(A81="",FALSE(),AND($U81>=0,$U81<=120,LUU_NIEN!$BI$99=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,99,61),1.0)))

def f974(E, r, c):
    # Ô AD81
    # = IF(A81="",FALSE(),AND($U81>=0,$U81<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$99))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,99,92)))))

def f975(E, r, c):
    # Ô A82
    # = IF((ISNUMBER(LUU_NIEN!$A$100)),LUU_NIEN!$A$100,"")
    return (E.v(15,100,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,100,1))) else '')

def f976(E, r, c):
    # Ô C82
    # = IF(A82="","",LUU_NIEN!$L$100)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,100,12))

def f977(E, r, c):
    # Ô D82
    # = IF(A82="","",LUU_NIEN!$H$100)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,100,8))

def f978(E, r, c):
    # Ô E82
    # = IF(A82="","",LUU_NIEN!$AA$100)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,100,27))

def f979(E, r, c):
    # Ô F82
    # = IF(A82="","",LUU_NIEN!$AL$100)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,100,38))

def f980(E, r, c):
    # Ô K82
    # = IF(A82="","",TRIM(IF(AB82,"Dấu hiệu bất lợi cho cha. ","")&IF(AC82,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA82,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U82>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$100))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,100,66))))) else ''))))

def f981(E, r, c):
    # Ô L82
    # = IF(OR(A82="",LUU_NIEN!$CN$100=""),"",IFERROR(LEFT(LUU_NIEN!$CN$100,FIND(":",LUU_NIEN!$CN$100)-1),LEFT(LUU_NIEN!$CN$100,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,100,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,100,92), X.sub(X.FIND(':', E.v(15,100,92)),1.0)), lambda: X.LEFT(E.v(15,100,92), 120.0)))

def f982(E, r, c):
    # Ô M82
    # = IF(A82="","",IF(LEN(LUU_NIEN!$BN$100)>260,LEFT(LUU_NIEN!$BN$100,260)&"…",LUU_NIEN!$BN$100))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,100,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,100,66)),260.0)) else E.v(15,100,66)))

def f983(E, r, c):
    # Ô T82
    # = IF(OR(A82="",LUU_NIEN!$P$100=""),-1,MOD(LUU_NIEN!$P$100-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,100,16),''))) else X.MOD(X.add(X.sub(E.v(15,100,16),E.v(23,7,5)),5.0), 5.0))

def f984(E, r, c):
    # Ô AA82
    # = IF(A82="",FALSE(),AND($U82>=18,$U82<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$100))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,100,66)))))

def f985(E, r, c):
    # Ô AB82
    # = IF(A82="",FALSE(),AND($U82>=0,$U82<=120,LUU_NIEN!$BH$100=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,100,60),1.0)))

def f986(E, r, c):
    # Ô AC82
    # = IF(A82="",FALSE(),AND($U82>=0,$U82<=120,LUU_NIEN!$BI$100=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,100,61),1.0)))

def f987(E, r, c):
    # Ô AD82
    # = IF(A82="",FALSE(),AND($U82>=0,$U82<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$100))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,100,92)))))

def f988(E, r, c):
    # Ô A83
    # = IF((ISNUMBER(LUU_NIEN!$A$101)),LUU_NIEN!$A$101,"")
    return (E.v(15,101,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,101,1))) else '')

def f989(E, r, c):
    # Ô C83
    # = IF(A83="","",LUU_NIEN!$L$101)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,101,12))

def f990(E, r, c):
    # Ô D83
    # = IF(A83="","",LUU_NIEN!$H$101)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,101,8))

def f991(E, r, c):
    # Ô E83
    # = IF(A83="","",LUU_NIEN!$AA$101)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,101,27))

def f992(E, r, c):
    # Ô F83
    # = IF(A83="","",LUU_NIEN!$AL$101)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,101,38))

def f993(E, r, c):
    # Ô K83
    # = IF(A83="","",TRIM(IF(AB83,"Dấu hiệu bất lợi cho cha. ","")&IF(AC83,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA83,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U83>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$101))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,101,66))))) else ''))))

def f994(E, r, c):
    # Ô L83
    # = IF(OR(A83="",LUU_NIEN!$CN$101=""),"",IFERROR(LEFT(LUU_NIEN!$CN$101,FIND(":",LUU_NIEN!$CN$101)-1),LEFT(LUU_NIEN!$CN$101,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,101,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,101,92), X.sub(X.FIND(':', E.v(15,101,92)),1.0)), lambda: X.LEFT(E.v(15,101,92), 120.0)))

def f995(E, r, c):
    # Ô M83
    # = IF(A83="","",IF(LEN(LUU_NIEN!$BN$101)>260,LEFT(LUU_NIEN!$BN$101,260)&"…",LUU_NIEN!$BN$101))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,101,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,101,66)),260.0)) else E.v(15,101,66)))

def f996(E, r, c):
    # Ô T83
    # = IF(OR(A83="",LUU_NIEN!$P$101=""),-1,MOD(LUU_NIEN!$P$101-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,101,16),''))) else X.MOD(X.add(X.sub(E.v(15,101,16),E.v(23,7,5)),5.0), 5.0))

def f997(E, r, c):
    # Ô AA83
    # = IF(A83="",FALSE(),AND($U83>=18,$U83<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$101))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,101,66)))))

def f998(E, r, c):
    # Ô AB83
    # = IF(A83="",FALSE(),AND($U83>=0,$U83<=120,LUU_NIEN!$BH$101=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,101,60),1.0)))

def f999(E, r, c):
    # Ô AC83
    # = IF(A83="",FALSE(),AND($U83>=0,$U83<=120,LUU_NIEN!$BI$101=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,101,61),1.0)))

def f1000(E, r, c):
    # Ô AD83
    # = IF(A83="",FALSE(),AND($U83>=0,$U83<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$101))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,101,92)))))

def f1001(E, r, c):
    # Ô A84
    # = IF((ISNUMBER(LUU_NIEN!$A$102)),LUU_NIEN!$A$102,"")
    return (E.v(15,102,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,102,1))) else '')

def f1002(E, r, c):
    # Ô C84
    # = IF(A84="","",LUU_NIEN!$L$102)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,102,12))

def f1003(E, r, c):
    # Ô D84
    # = IF(A84="","",LUU_NIEN!$H$102)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,102,8))

def f1004(E, r, c):
    # Ô E84
    # = IF(A84="","",LUU_NIEN!$AA$102)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,102,27))

def f1005(E, r, c):
    # Ô F84
    # = IF(A84="","",LUU_NIEN!$AL$102)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,102,38))

def f1006(E, r, c):
    # Ô K84
    # = IF(A84="","",TRIM(IF(AB84,"Dấu hiệu bất lợi cho cha. ","")&IF(AC84,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA84,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U84>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$102))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,102,66))))) else ''))))

def f1007(E, r, c):
    # Ô L84
    # = IF(OR(A84="",LUU_NIEN!$CN$102=""),"",IFERROR(LEFT(LUU_NIEN!$CN$102,FIND(":",LUU_NIEN!$CN$102)-1),LEFT(LUU_NIEN!$CN$102,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,102,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,102,92), X.sub(X.FIND(':', E.v(15,102,92)),1.0)), lambda: X.LEFT(E.v(15,102,92), 120.0)))

def f1008(E, r, c):
    # Ô M84
    # = IF(A84="","",IF(LEN(LUU_NIEN!$BN$102)>260,LEFT(LUU_NIEN!$BN$102,260)&"…",LUU_NIEN!$BN$102))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,102,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,102,66)),260.0)) else E.v(15,102,66)))

def f1009(E, r, c):
    # Ô T84
    # = IF(OR(A84="",LUU_NIEN!$P$102=""),-1,MOD(LUU_NIEN!$P$102-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,102,16),''))) else X.MOD(X.add(X.sub(E.v(15,102,16),E.v(23,7,5)),5.0), 5.0))

def f1010(E, r, c):
    # Ô AA84
    # = IF(A84="",FALSE(),AND($U84>=18,$U84<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$102))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,102,66)))))

def f1011(E, r, c):
    # Ô AB84
    # = IF(A84="",FALSE(),AND($U84>=0,$U84<=120,LUU_NIEN!$BH$102=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,102,60),1.0)))

def f1012(E, r, c):
    # Ô AC84
    # = IF(A84="",FALSE(),AND($U84>=0,$U84<=120,LUU_NIEN!$BI$102=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,102,61),1.0)))

def f1013(E, r, c):
    # Ô AD84
    # = IF(A84="",FALSE(),AND($U84>=0,$U84<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$102))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,102,92)))))

def f1014(E, r, c):
    # Ô A85
    # = IF((ISNUMBER(LUU_NIEN!$A$103)),LUU_NIEN!$A$103,"")
    return (E.v(15,103,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,103,1))) else '')

def f1015(E, r, c):
    # Ô C85
    # = IF(A85="","",LUU_NIEN!$L$103)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,103,12))

def f1016(E, r, c):
    # Ô D85
    # = IF(A85="","",LUU_NIEN!$H$103)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,103,8))

def f1017(E, r, c):
    # Ô E85
    # = IF(A85="","",LUU_NIEN!$AA$103)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,103,27))

def f1018(E, r, c):
    # Ô F85
    # = IF(A85="","",LUU_NIEN!$AL$103)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,103,38))

def f1019(E, r, c):
    # Ô K85
    # = IF(A85="","",TRIM(IF(AB85,"Dấu hiệu bất lợi cho cha. ","")&IF(AC85,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA85,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U85>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$103))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,103,66))))) else ''))))

def f1020(E, r, c):
    # Ô L85
    # = IF(OR(A85="",LUU_NIEN!$CN$103=""),"",IFERROR(LEFT(LUU_NIEN!$CN$103,FIND(":",LUU_NIEN!$CN$103)-1),LEFT(LUU_NIEN!$CN$103,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,103,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,103,92), X.sub(X.FIND(':', E.v(15,103,92)),1.0)), lambda: X.LEFT(E.v(15,103,92), 120.0)))

def f1021(E, r, c):
    # Ô M85
    # = IF(A85="","",IF(LEN(LUU_NIEN!$BN$103)>260,LEFT(LUU_NIEN!$BN$103,260)&"…",LUU_NIEN!$BN$103))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,103,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,103,66)),260.0)) else E.v(15,103,66)))

def f1022(E, r, c):
    # Ô T85
    # = IF(OR(A85="",LUU_NIEN!$P$103=""),-1,MOD(LUU_NIEN!$P$103-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,103,16),''))) else X.MOD(X.add(X.sub(E.v(15,103,16),E.v(23,7,5)),5.0), 5.0))

def f1023(E, r, c):
    # Ô AA85
    # = IF(A85="",FALSE(),AND($U85>=18,$U85<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$103))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,103,66)))))

def f1024(E, r, c):
    # Ô AB85
    # = IF(A85="",FALSE(),AND($U85>=0,$U85<=120,LUU_NIEN!$BH$103=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,103,60),1.0)))

def f1025(E, r, c):
    # Ô AC85
    # = IF(A85="",FALSE(),AND($U85>=0,$U85<=120,LUU_NIEN!$BI$103=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,103,61),1.0)))

def f1026(E, r, c):
    # Ô AD85
    # = IF(A85="",FALSE(),AND($U85>=0,$U85<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$103))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,103,92)))))

def f1027(E, r, c):
    # Ô A86
    # = IF((ISNUMBER(LUU_NIEN!$A$104)),LUU_NIEN!$A$104,"")
    return (E.v(15,104,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,104,1))) else '')

def f1028(E, r, c):
    # Ô C86
    # = IF(A86="","",LUU_NIEN!$L$104)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,104,12))

def f1029(E, r, c):
    # Ô D86
    # = IF(A86="","",LUU_NIEN!$H$104)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,104,8))

def f1030(E, r, c):
    # Ô E86
    # = IF(A86="","",LUU_NIEN!$AA$104)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,104,27))

def f1031(E, r, c):
    # Ô F86
    # = IF(A86="","",LUU_NIEN!$AL$104)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,104,38))

def f1032(E, r, c):
    # Ô K86
    # = IF(A86="","",TRIM(IF(AB86,"Dấu hiệu bất lợi cho cha. ","")&IF(AC86,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA86,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U86>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$104))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,104,66))))) else ''))))

def f1033(E, r, c):
    # Ô L86
    # = IF(OR(A86="",LUU_NIEN!$CN$104=""),"",IFERROR(LEFT(LUU_NIEN!$CN$104,FIND(":",LUU_NIEN!$CN$104)-1),LEFT(LUU_NIEN!$CN$104,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,104,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,104,92), X.sub(X.FIND(':', E.v(15,104,92)),1.0)), lambda: X.LEFT(E.v(15,104,92), 120.0)))

def f1034(E, r, c):
    # Ô M86
    # = IF(A86="","",IF(LEN(LUU_NIEN!$BN$104)>260,LEFT(LUU_NIEN!$BN$104,260)&"…",LUU_NIEN!$BN$104))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,104,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,104,66)),260.0)) else E.v(15,104,66)))

def f1035(E, r, c):
    # Ô T86
    # = IF(OR(A86="",LUU_NIEN!$P$104=""),-1,MOD(LUU_NIEN!$P$104-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,104,16),''))) else X.MOD(X.add(X.sub(E.v(15,104,16),E.v(23,7,5)),5.0), 5.0))

def f1036(E, r, c):
    # Ô AA86
    # = IF(A86="",FALSE(),AND($U86>=18,$U86<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$104))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,104,66)))))

def f1037(E, r, c):
    # Ô AB86
    # = IF(A86="",FALSE(),AND($U86>=0,$U86<=120,LUU_NIEN!$BH$104=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,104,60),1.0)))

def f1038(E, r, c):
    # Ô AC86
    # = IF(A86="",FALSE(),AND($U86>=0,$U86<=120,LUU_NIEN!$BI$104=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,104,61),1.0)))

def f1039(E, r, c):
    # Ô AD86
    # = IF(A86="",FALSE(),AND($U86>=0,$U86<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$104))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,104,92)))))

def f1040(E, r, c):
    # Ô A87
    # = IF((ISNUMBER(LUU_NIEN!$A$105)),LUU_NIEN!$A$105,"")
    return (E.v(15,105,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,105,1))) else '')

def f1041(E, r, c):
    # Ô C87
    # = IF(A87="","",LUU_NIEN!$L$105)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,105,12))

def f1042(E, r, c):
    # Ô D87
    # = IF(A87="","",LUU_NIEN!$H$105)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,105,8))

def f1043(E, r, c):
    # Ô E87
    # = IF(A87="","",LUU_NIEN!$AA$105)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,105,27))

def f1044(E, r, c):
    # Ô F87
    # = IF(A87="","",LUU_NIEN!$AL$105)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,105,38))

def f1045(E, r, c):
    # Ô K87
    # = IF(A87="","",TRIM(IF(AB87,"Dấu hiệu bất lợi cho cha. ","")&IF(AC87,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA87,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U87>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$105))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,105,66))))) else ''))))

def f1046(E, r, c):
    # Ô L87
    # = IF(OR(A87="",LUU_NIEN!$CN$105=""),"",IFERROR(LEFT(LUU_NIEN!$CN$105,FIND(":",LUU_NIEN!$CN$105)-1),LEFT(LUU_NIEN!$CN$105,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,105,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,105,92), X.sub(X.FIND(':', E.v(15,105,92)),1.0)), lambda: X.LEFT(E.v(15,105,92), 120.0)))

def f1047(E, r, c):
    # Ô M87
    # = IF(A87="","",IF(LEN(LUU_NIEN!$BN$105)>260,LEFT(LUU_NIEN!$BN$105,260)&"…",LUU_NIEN!$BN$105))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,105,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,105,66)),260.0)) else E.v(15,105,66)))

def f1048(E, r, c):
    # Ô T87
    # = IF(OR(A87="",LUU_NIEN!$P$105=""),-1,MOD(LUU_NIEN!$P$105-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,105,16),''))) else X.MOD(X.add(X.sub(E.v(15,105,16),E.v(23,7,5)),5.0), 5.0))

def f1049(E, r, c):
    # Ô AA87
    # = IF(A87="",FALSE(),AND($U87>=18,$U87<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$105))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,105,66)))))

def f1050(E, r, c):
    # Ô AB87
    # = IF(A87="",FALSE(),AND($U87>=0,$U87<=120,LUU_NIEN!$BH$105=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,105,60),1.0)))

def f1051(E, r, c):
    # Ô AC87
    # = IF(A87="",FALSE(),AND($U87>=0,$U87<=120,LUU_NIEN!$BI$105=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,105,61),1.0)))

def f1052(E, r, c):
    # Ô AD87
    # = IF(A87="",FALSE(),AND($U87>=0,$U87<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$105))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,105,92)))))

def f1053(E, r, c):
    # Ô A88
    # = IF((ISNUMBER(LUU_NIEN!$A$106)),LUU_NIEN!$A$106,"")
    return (E.v(15,106,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,106,1))) else '')

def f1054(E, r, c):
    # Ô C88
    # = IF(A88="","",LUU_NIEN!$L$106)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,106,12))

def f1055(E, r, c):
    # Ô D88
    # = IF(A88="","",LUU_NIEN!$H$106)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,106,8))

def f1056(E, r, c):
    # Ô E88
    # = IF(A88="","",LUU_NIEN!$AA$106)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,106,27))

def f1057(E, r, c):
    # Ô F88
    # = IF(A88="","",LUU_NIEN!$AL$106)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,106,38))

def f1058(E, r, c):
    # Ô K88
    # = IF(A88="","",TRIM(IF(AB88,"Dấu hiệu bất lợi cho cha. ","")&IF(AC88,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA88,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U88>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$106))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,106,66))))) else ''))))

def f1059(E, r, c):
    # Ô L88
    # = IF(OR(A88="",LUU_NIEN!$CN$106=""),"",IFERROR(LEFT(LUU_NIEN!$CN$106,FIND(":",LUU_NIEN!$CN$106)-1),LEFT(LUU_NIEN!$CN$106,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,106,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,106,92), X.sub(X.FIND(':', E.v(15,106,92)),1.0)), lambda: X.LEFT(E.v(15,106,92), 120.0)))

def f1060(E, r, c):
    # Ô M88
    # = IF(A88="","",IF(LEN(LUU_NIEN!$BN$106)>260,LEFT(LUU_NIEN!$BN$106,260)&"…",LUU_NIEN!$BN$106))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,106,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,106,66)),260.0)) else E.v(15,106,66)))

def f1061(E, r, c):
    # Ô T88
    # = IF(OR(A88="",LUU_NIEN!$P$106=""),-1,MOD(LUU_NIEN!$P$106-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,106,16),''))) else X.MOD(X.add(X.sub(E.v(15,106,16),E.v(23,7,5)),5.0), 5.0))

def f1062(E, r, c):
    # Ô AA88
    # = IF(A88="",FALSE(),AND($U88>=18,$U88<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$106))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,106,66)))))

def f1063(E, r, c):
    # Ô AB88
    # = IF(A88="",FALSE(),AND($U88>=0,$U88<=120,LUU_NIEN!$BH$106=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,106,60),1.0)))

def f1064(E, r, c):
    # Ô AC88
    # = IF(A88="",FALSE(),AND($U88>=0,$U88<=120,LUU_NIEN!$BI$106=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,106,61),1.0)))

def f1065(E, r, c):
    # Ô AD88
    # = IF(A88="",FALSE(),AND($U88>=0,$U88<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$106))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,106,92)))))

def f1066(E, r, c):
    # Ô A89
    # = IF((ISNUMBER(LUU_NIEN!$A$107)),LUU_NIEN!$A$107,"")
    return (E.v(15,107,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,107,1))) else '')

def f1067(E, r, c):
    # Ô C89
    # = IF(A89="","",LUU_NIEN!$L$107)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,107,12))

def f1068(E, r, c):
    # Ô D89
    # = IF(A89="","",LUU_NIEN!$H$107)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,107,8))

def f1069(E, r, c):
    # Ô E89
    # = IF(A89="","",LUU_NIEN!$AA$107)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,107,27))

def f1070(E, r, c):
    # Ô F89
    # = IF(A89="","",LUU_NIEN!$AL$107)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,107,38))

def f1071(E, r, c):
    # Ô K89
    # = IF(A89="","",TRIM(IF(AB89,"Dấu hiệu bất lợi cho cha. ","")&IF(AC89,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA89,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U89>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$107))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,107,66))))) else ''))))

def f1072(E, r, c):
    # Ô L89
    # = IF(OR(A89="",LUU_NIEN!$CN$107=""),"",IFERROR(LEFT(LUU_NIEN!$CN$107,FIND(":",LUU_NIEN!$CN$107)-1),LEFT(LUU_NIEN!$CN$107,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,107,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,107,92), X.sub(X.FIND(':', E.v(15,107,92)),1.0)), lambda: X.LEFT(E.v(15,107,92), 120.0)))

def f1073(E, r, c):
    # Ô M89
    # = IF(A89="","",IF(LEN(LUU_NIEN!$BN$107)>260,LEFT(LUU_NIEN!$BN$107,260)&"…",LUU_NIEN!$BN$107))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,107,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,107,66)),260.0)) else E.v(15,107,66)))

def f1074(E, r, c):
    # Ô T89
    # = IF(OR(A89="",LUU_NIEN!$P$107=""),-1,MOD(LUU_NIEN!$P$107-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,107,16),''))) else X.MOD(X.add(X.sub(E.v(15,107,16),E.v(23,7,5)),5.0), 5.0))

def f1075(E, r, c):
    # Ô AA89
    # = IF(A89="",FALSE(),AND($U89>=18,$U89<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$107))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,107,66)))))

def f1076(E, r, c):
    # Ô AB89
    # = IF(A89="",FALSE(),AND($U89>=0,$U89<=120,LUU_NIEN!$BH$107=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,107,60),1.0)))

def f1077(E, r, c):
    # Ô AC89
    # = IF(A89="",FALSE(),AND($U89>=0,$U89<=120,LUU_NIEN!$BI$107=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,107,61),1.0)))

def f1078(E, r, c):
    # Ô AD89
    # = IF(A89="",FALSE(),AND($U89>=0,$U89<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$107))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,107,92)))))

def f1079(E, r, c):
    # Ô A90
    # = IF((ISNUMBER(LUU_NIEN!$A$108)),LUU_NIEN!$A$108,"")
    return (E.v(15,108,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,108,1))) else '')

def f1080(E, r, c):
    # Ô C90
    # = IF(A90="","",LUU_NIEN!$L$108)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,108,12))

def f1081(E, r, c):
    # Ô D90
    # = IF(A90="","",LUU_NIEN!$H$108)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,108,8))

def f1082(E, r, c):
    # Ô E90
    # = IF(A90="","",LUU_NIEN!$AA$108)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,108,27))

def f1083(E, r, c):
    # Ô F90
    # = IF(A90="","",LUU_NIEN!$AL$108)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,108,38))

def f1084(E, r, c):
    # Ô K90
    # = IF(A90="","",TRIM(IF(AB90,"Dấu hiệu bất lợi cho cha. ","")&IF(AC90,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA90,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U90>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$108))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,108,66))))) else ''))))

def f1085(E, r, c):
    # Ô L90
    # = IF(OR(A90="",LUU_NIEN!$CN$108=""),"",IFERROR(LEFT(LUU_NIEN!$CN$108,FIND(":",LUU_NIEN!$CN$108)-1),LEFT(LUU_NIEN!$CN$108,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,108,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,108,92), X.sub(X.FIND(':', E.v(15,108,92)),1.0)), lambda: X.LEFT(E.v(15,108,92), 120.0)))

def f1086(E, r, c):
    # Ô M90
    # = IF(A90="","",IF(LEN(LUU_NIEN!$BN$108)>260,LEFT(LUU_NIEN!$BN$108,260)&"…",LUU_NIEN!$BN$108))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,108,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,108,66)),260.0)) else E.v(15,108,66)))

def f1087(E, r, c):
    # Ô T90
    # = IF(OR(A90="",LUU_NIEN!$P$108=""),-1,MOD(LUU_NIEN!$P$108-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,108,16),''))) else X.MOD(X.add(X.sub(E.v(15,108,16),E.v(23,7,5)),5.0), 5.0))

def f1088(E, r, c):
    # Ô AA90
    # = IF(A90="",FALSE(),AND($U90>=18,$U90<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$108))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,108,66)))))

def f1089(E, r, c):
    # Ô AB90
    # = IF(A90="",FALSE(),AND($U90>=0,$U90<=120,LUU_NIEN!$BH$108=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,108,60),1.0)))

def f1090(E, r, c):
    # Ô AC90
    # = IF(A90="",FALSE(),AND($U90>=0,$U90<=120,LUU_NIEN!$BI$108=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,108,61),1.0)))

def f1091(E, r, c):
    # Ô AD90
    # = IF(A90="",FALSE(),AND($U90>=0,$U90<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$108))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,108,92)))))

def f1092(E, r, c):
    # Ô A91
    # = IF((ISNUMBER(LUU_NIEN!$A$109)),LUU_NIEN!$A$109,"")
    return (E.v(15,109,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,109,1))) else '')

def f1093(E, r, c):
    # Ô C91
    # = IF(A91="","",LUU_NIEN!$L$109)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,109,12))

def f1094(E, r, c):
    # Ô D91
    # = IF(A91="","",LUU_NIEN!$H$109)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,109,8))

def f1095(E, r, c):
    # Ô E91
    # = IF(A91="","",LUU_NIEN!$AA$109)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,109,27))

def f1096(E, r, c):
    # Ô F91
    # = IF(A91="","",LUU_NIEN!$AL$109)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,109,38))

def f1097(E, r, c):
    # Ô K91
    # = IF(A91="","",TRIM(IF(AB91,"Dấu hiệu bất lợi cho cha. ","")&IF(AC91,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA91,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U91>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$109))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,109,66))))) else ''))))

def f1098(E, r, c):
    # Ô L91
    # = IF(OR(A91="",LUU_NIEN!$CN$109=""),"",IFERROR(LEFT(LUU_NIEN!$CN$109,FIND(":",LUU_NIEN!$CN$109)-1),LEFT(LUU_NIEN!$CN$109,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,109,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,109,92), X.sub(X.FIND(':', E.v(15,109,92)),1.0)), lambda: X.LEFT(E.v(15,109,92), 120.0)))

def f1099(E, r, c):
    # Ô M91
    # = IF(A91="","",IF(LEN(LUU_NIEN!$BN$109)>260,LEFT(LUU_NIEN!$BN$109,260)&"…",LUU_NIEN!$BN$109))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,109,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,109,66)),260.0)) else E.v(15,109,66)))

def f1100(E, r, c):
    # Ô T91
    # = IF(OR(A91="",LUU_NIEN!$P$109=""),-1,MOD(LUU_NIEN!$P$109-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,109,16),''))) else X.MOD(X.add(X.sub(E.v(15,109,16),E.v(23,7,5)),5.0), 5.0))

def f1101(E, r, c):
    # Ô AA91
    # = IF(A91="",FALSE(),AND($U91>=18,$U91<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$109))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,109,66)))))

def f1102(E, r, c):
    # Ô AB91
    # = IF(A91="",FALSE(),AND($U91>=0,$U91<=120,LUU_NIEN!$BH$109=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,109,60),1.0)))

def f1103(E, r, c):
    # Ô AC91
    # = IF(A91="",FALSE(),AND($U91>=0,$U91<=120,LUU_NIEN!$BI$109=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,109,61),1.0)))

def f1104(E, r, c):
    # Ô AD91
    # = IF(A91="",FALSE(),AND($U91>=0,$U91<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$109))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,109,92)))))

def f1105(E, r, c):
    # Ô A92
    # = IF((ISNUMBER(LUU_NIEN!$A$110)),LUU_NIEN!$A$110,"")
    return (E.v(15,110,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,110,1))) else '')

def f1106(E, r, c):
    # Ô C92
    # = IF(A92="","",LUU_NIEN!$L$110)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,110,12))

def f1107(E, r, c):
    # Ô D92
    # = IF(A92="","",LUU_NIEN!$H$110)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,110,8))

def f1108(E, r, c):
    # Ô E92
    # = IF(A92="","",LUU_NIEN!$AA$110)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,110,27))

def f1109(E, r, c):
    # Ô F92
    # = IF(A92="","",LUU_NIEN!$AL$110)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,110,38))

def f1110(E, r, c):
    # Ô K92
    # = IF(A92="","",TRIM(IF(AB92,"Dấu hiệu bất lợi cho cha. ","")&IF(AC92,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA92,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U92>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$110))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,110,66))))) else ''))))

def f1111(E, r, c):
    # Ô L92
    # = IF(OR(A92="",LUU_NIEN!$CN$110=""),"",IFERROR(LEFT(LUU_NIEN!$CN$110,FIND(":",LUU_NIEN!$CN$110)-1),LEFT(LUU_NIEN!$CN$110,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,110,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,110,92), X.sub(X.FIND(':', E.v(15,110,92)),1.0)), lambda: X.LEFT(E.v(15,110,92), 120.0)))

def f1112(E, r, c):
    # Ô M92
    # = IF(A92="","",IF(LEN(LUU_NIEN!$BN$110)>260,LEFT(LUU_NIEN!$BN$110,260)&"…",LUU_NIEN!$BN$110))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,110,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,110,66)),260.0)) else E.v(15,110,66)))

def f1113(E, r, c):
    # Ô T92
    # = IF(OR(A92="",LUU_NIEN!$P$110=""),-1,MOD(LUU_NIEN!$P$110-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,110,16),''))) else X.MOD(X.add(X.sub(E.v(15,110,16),E.v(23,7,5)),5.0), 5.0))

def f1114(E, r, c):
    # Ô AA92
    # = IF(A92="",FALSE(),AND($U92>=18,$U92<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$110))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,110,66)))))

def f1115(E, r, c):
    # Ô AB92
    # = IF(A92="",FALSE(),AND($U92>=0,$U92<=120,LUU_NIEN!$BH$110=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,110,60),1.0)))

def f1116(E, r, c):
    # Ô AC92
    # = IF(A92="",FALSE(),AND($U92>=0,$U92<=120,LUU_NIEN!$BI$110=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,110,61),1.0)))

def f1117(E, r, c):
    # Ô AD92
    # = IF(A92="",FALSE(),AND($U92>=0,$U92<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$110))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,110,92)))))

def f1118(E, r, c):
    # Ô A93
    # = IF((ISNUMBER(LUU_NIEN!$A$111)),LUU_NIEN!$A$111,"")
    return (E.v(15,111,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,111,1))) else '')

def f1119(E, r, c):
    # Ô C93
    # = IF(A93="","",LUU_NIEN!$L$111)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,111,12))

def f1120(E, r, c):
    # Ô D93
    # = IF(A93="","",LUU_NIEN!$H$111)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,111,8))

def f1121(E, r, c):
    # Ô E93
    # = IF(A93="","",LUU_NIEN!$AA$111)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,111,27))

def f1122(E, r, c):
    # Ô F93
    # = IF(A93="","",LUU_NIEN!$AL$111)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,111,38))

def f1123(E, r, c):
    # Ô K93
    # = IF(A93="","",TRIM(IF(AB93,"Dấu hiệu bất lợi cho cha. ","")&IF(AC93,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA93,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U93>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$111))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,111,66))))) else ''))))

def f1124(E, r, c):
    # Ô L93
    # = IF(OR(A93="",LUU_NIEN!$CN$111=""),"",IFERROR(LEFT(LUU_NIEN!$CN$111,FIND(":",LUU_NIEN!$CN$111)-1),LEFT(LUU_NIEN!$CN$111,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,111,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,111,92), X.sub(X.FIND(':', E.v(15,111,92)),1.0)), lambda: X.LEFT(E.v(15,111,92), 120.0)))

def f1125(E, r, c):
    # Ô M93
    # = IF(A93="","",IF(LEN(LUU_NIEN!$BN$111)>260,LEFT(LUU_NIEN!$BN$111,260)&"…",LUU_NIEN!$BN$111))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,111,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,111,66)),260.0)) else E.v(15,111,66)))

def f1126(E, r, c):
    # Ô T93
    # = IF(OR(A93="",LUU_NIEN!$P$111=""),-1,MOD(LUU_NIEN!$P$111-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,111,16),''))) else X.MOD(X.add(X.sub(E.v(15,111,16),E.v(23,7,5)),5.0), 5.0))

def f1127(E, r, c):
    # Ô AA93
    # = IF(A93="",FALSE(),AND($U93>=18,$U93<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$111))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,111,66)))))

def f1128(E, r, c):
    # Ô AB93
    # = IF(A93="",FALSE(),AND($U93>=0,$U93<=120,LUU_NIEN!$BH$111=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,111,60),1.0)))

def f1129(E, r, c):
    # Ô AC93
    # = IF(A93="",FALSE(),AND($U93>=0,$U93<=120,LUU_NIEN!$BI$111=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,111,61),1.0)))

def f1130(E, r, c):
    # Ô AD93
    # = IF(A93="",FALSE(),AND($U93>=0,$U93<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$111))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,111,92)))))

def f1131(E, r, c):
    # Ô A94
    # = IF((ISNUMBER(LUU_NIEN!$A$112)),LUU_NIEN!$A$112,"")
    return (E.v(15,112,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,112,1))) else '')

def f1132(E, r, c):
    # Ô C94
    # = IF(A94="","",LUU_NIEN!$L$112)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,112,12))

def f1133(E, r, c):
    # Ô D94
    # = IF(A94="","",LUU_NIEN!$H$112)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,112,8))

def f1134(E, r, c):
    # Ô E94
    # = IF(A94="","",LUU_NIEN!$AA$112)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,112,27))

def f1135(E, r, c):
    # Ô F94
    # = IF(A94="","",LUU_NIEN!$AL$112)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,112,38))

def f1136(E, r, c):
    # Ô K94
    # = IF(A94="","",TRIM(IF(AB94,"Dấu hiệu bất lợi cho cha. ","")&IF(AC94,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA94,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U94>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$112))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,112,66))))) else ''))))

def f1137(E, r, c):
    # Ô L94
    # = IF(OR(A94="",LUU_NIEN!$CN$112=""),"",IFERROR(LEFT(LUU_NIEN!$CN$112,FIND(":",LUU_NIEN!$CN$112)-1),LEFT(LUU_NIEN!$CN$112,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,112,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,112,92), X.sub(X.FIND(':', E.v(15,112,92)),1.0)), lambda: X.LEFT(E.v(15,112,92), 120.0)))

def f1138(E, r, c):
    # Ô M94
    # = IF(A94="","",IF(LEN(LUU_NIEN!$BN$112)>260,LEFT(LUU_NIEN!$BN$112,260)&"…",LUU_NIEN!$BN$112))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,112,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,112,66)),260.0)) else E.v(15,112,66)))

def f1139(E, r, c):
    # Ô T94
    # = IF(OR(A94="",LUU_NIEN!$P$112=""),-1,MOD(LUU_NIEN!$P$112-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,112,16),''))) else X.MOD(X.add(X.sub(E.v(15,112,16),E.v(23,7,5)),5.0), 5.0))

def f1140(E, r, c):
    # Ô AA94
    # = IF(A94="",FALSE(),AND($U94>=18,$U94<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$112))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,112,66)))))

def f1141(E, r, c):
    # Ô AB94
    # = IF(A94="",FALSE(),AND($U94>=0,$U94<=120,LUU_NIEN!$BH$112=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,112,60),1.0)))

def f1142(E, r, c):
    # Ô AC94
    # = IF(A94="",FALSE(),AND($U94>=0,$U94<=120,LUU_NIEN!$BI$112=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,112,61),1.0)))

def f1143(E, r, c):
    # Ô AD94
    # = IF(A94="",FALSE(),AND($U94>=0,$U94<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$112))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,112,92)))))

def f1144(E, r, c):
    # Ô A95
    # = IF((ISNUMBER(LUU_NIEN!$A$113)),LUU_NIEN!$A$113,"")
    return (E.v(15,113,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,113,1))) else '')

def f1145(E, r, c):
    # Ô C95
    # = IF(A95="","",LUU_NIEN!$L$113)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,113,12))

def f1146(E, r, c):
    # Ô D95
    # = IF(A95="","",LUU_NIEN!$H$113)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,113,8))

def f1147(E, r, c):
    # Ô E95
    # = IF(A95="","",LUU_NIEN!$AA$113)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,113,27))

def f1148(E, r, c):
    # Ô F95
    # = IF(A95="","",LUU_NIEN!$AL$113)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,113,38))

def f1149(E, r, c):
    # Ô K95
    # = IF(A95="","",TRIM(IF(AB95,"Dấu hiệu bất lợi cho cha. ","")&IF(AC95,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA95,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U95>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$113))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,113,66))))) else ''))))

def f1150(E, r, c):
    # Ô L95
    # = IF(OR(A95="",LUU_NIEN!$CN$113=""),"",IFERROR(LEFT(LUU_NIEN!$CN$113,FIND(":",LUU_NIEN!$CN$113)-1),LEFT(LUU_NIEN!$CN$113,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,113,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,113,92), X.sub(X.FIND(':', E.v(15,113,92)),1.0)), lambda: X.LEFT(E.v(15,113,92), 120.0)))

def f1151(E, r, c):
    # Ô M95
    # = IF(A95="","",IF(LEN(LUU_NIEN!$BN$113)>260,LEFT(LUU_NIEN!$BN$113,260)&"…",LUU_NIEN!$BN$113))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,113,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,113,66)),260.0)) else E.v(15,113,66)))

def f1152(E, r, c):
    # Ô T95
    # = IF(OR(A95="",LUU_NIEN!$P$113=""),-1,MOD(LUU_NIEN!$P$113-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,113,16),''))) else X.MOD(X.add(X.sub(E.v(15,113,16),E.v(23,7,5)),5.0), 5.0))

def f1153(E, r, c):
    # Ô AA95
    # = IF(A95="",FALSE(),AND($U95>=18,$U95<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$113))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,113,66)))))

def f1154(E, r, c):
    # Ô AB95
    # = IF(A95="",FALSE(),AND($U95>=0,$U95<=120,LUU_NIEN!$BH$113=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,113,60),1.0)))

def f1155(E, r, c):
    # Ô AC95
    # = IF(A95="",FALSE(),AND($U95>=0,$U95<=120,LUU_NIEN!$BI$113=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,113,61),1.0)))

def f1156(E, r, c):
    # Ô AD95
    # = IF(A95="",FALSE(),AND($U95>=0,$U95<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$113))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,113,92)))))

def f1157(E, r, c):
    # Ô A96
    # = IF((ISNUMBER(LUU_NIEN!$A$114)),LUU_NIEN!$A$114,"")
    return (E.v(15,114,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,114,1))) else '')

def f1158(E, r, c):
    # Ô C96
    # = IF(A96="","",LUU_NIEN!$L$114)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,114,12))

def f1159(E, r, c):
    # Ô D96
    # = IF(A96="","",LUU_NIEN!$H$114)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,114,8))

def f1160(E, r, c):
    # Ô E96
    # = IF(A96="","",LUU_NIEN!$AA$114)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,114,27))

def f1161(E, r, c):
    # Ô F96
    # = IF(A96="","",LUU_NIEN!$AL$114)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,114,38))

def f1162(E, r, c):
    # Ô K96
    # = IF(A96="","",TRIM(IF(AB96,"Dấu hiệu bất lợi cho cha. ","")&IF(AC96,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA96,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U96>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$114))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,114,66))))) else ''))))

def f1163(E, r, c):
    # Ô L96
    # = IF(OR(A96="",LUU_NIEN!$CN$114=""),"",IFERROR(LEFT(LUU_NIEN!$CN$114,FIND(":",LUU_NIEN!$CN$114)-1),LEFT(LUU_NIEN!$CN$114,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,114,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,114,92), X.sub(X.FIND(':', E.v(15,114,92)),1.0)), lambda: X.LEFT(E.v(15,114,92), 120.0)))

def f1164(E, r, c):
    # Ô M96
    # = IF(A96="","",IF(LEN(LUU_NIEN!$BN$114)>260,LEFT(LUU_NIEN!$BN$114,260)&"…",LUU_NIEN!$BN$114))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,114,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,114,66)),260.0)) else E.v(15,114,66)))

def f1165(E, r, c):
    # Ô T96
    # = IF(OR(A96="",LUU_NIEN!$P$114=""),-1,MOD(LUU_NIEN!$P$114-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,114,16),''))) else X.MOD(X.add(X.sub(E.v(15,114,16),E.v(23,7,5)),5.0), 5.0))

def f1166(E, r, c):
    # Ô AA96
    # = IF(A96="",FALSE(),AND($U96>=18,$U96<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$114))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,114,66)))))

def f1167(E, r, c):
    # Ô AB96
    # = IF(A96="",FALSE(),AND($U96>=0,$U96<=120,LUU_NIEN!$BH$114=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,114,60),1.0)))

def f1168(E, r, c):
    # Ô AC96
    # = IF(A96="",FALSE(),AND($U96>=0,$U96<=120,LUU_NIEN!$BI$114=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,114,61),1.0)))

def f1169(E, r, c):
    # Ô AD96
    # = IF(A96="",FALSE(),AND($U96>=0,$U96<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$114))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,114,92)))))

def f1170(E, r, c):
    # Ô A97
    # = IF((ISNUMBER(LUU_NIEN!$A$115)),LUU_NIEN!$A$115,"")
    return (E.v(15,115,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,115,1))) else '')

def f1171(E, r, c):
    # Ô C97
    # = IF(A97="","",LUU_NIEN!$L$115)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,115,12))

def f1172(E, r, c):
    # Ô D97
    # = IF(A97="","",LUU_NIEN!$H$115)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,115,8))

def f1173(E, r, c):
    # Ô E97
    # = IF(A97="","",LUU_NIEN!$AA$115)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,115,27))

def f1174(E, r, c):
    # Ô F97
    # = IF(A97="","",LUU_NIEN!$AL$115)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,115,38))

def f1175(E, r, c):
    # Ô K97
    # = IF(A97="","",TRIM(IF(AB97,"Dấu hiệu bất lợi cho cha. ","")&IF(AC97,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA97,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U97>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$115))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,115,66))))) else ''))))

def f1176(E, r, c):
    # Ô L97
    # = IF(OR(A97="",LUU_NIEN!$CN$115=""),"",IFERROR(LEFT(LUU_NIEN!$CN$115,FIND(":",LUU_NIEN!$CN$115)-1),LEFT(LUU_NIEN!$CN$115,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,115,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,115,92), X.sub(X.FIND(':', E.v(15,115,92)),1.0)), lambda: X.LEFT(E.v(15,115,92), 120.0)))

def f1177(E, r, c):
    # Ô M97
    # = IF(A97="","",IF(LEN(LUU_NIEN!$BN$115)>260,LEFT(LUU_NIEN!$BN$115,260)&"…",LUU_NIEN!$BN$115))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,115,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,115,66)),260.0)) else E.v(15,115,66)))

def f1178(E, r, c):
    # Ô T97
    # = IF(OR(A97="",LUU_NIEN!$P$115=""),-1,MOD(LUU_NIEN!$P$115-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,115,16),''))) else X.MOD(X.add(X.sub(E.v(15,115,16),E.v(23,7,5)),5.0), 5.0))

def f1179(E, r, c):
    # Ô AA97
    # = IF(A97="",FALSE(),AND($U97>=18,$U97<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$115))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,115,66)))))

def f1180(E, r, c):
    # Ô AB97
    # = IF(A97="",FALSE(),AND($U97>=0,$U97<=120,LUU_NIEN!$BH$115=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,115,60),1.0)))

def f1181(E, r, c):
    # Ô AC97
    # = IF(A97="",FALSE(),AND($U97>=0,$U97<=120,LUU_NIEN!$BI$115=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,115,61),1.0)))

def f1182(E, r, c):
    # Ô AD97
    # = IF(A97="",FALSE(),AND($U97>=0,$U97<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$115))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,115,92)))))

def f1183(E, r, c):
    # Ô A98
    # = IF((ISNUMBER(LUU_NIEN!$A$116)),LUU_NIEN!$A$116,"")
    return (E.v(15,116,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,116,1))) else '')

def f1184(E, r, c):
    # Ô C98
    # = IF(A98="","",LUU_NIEN!$L$116)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,116,12))

def f1185(E, r, c):
    # Ô D98
    # = IF(A98="","",LUU_NIEN!$H$116)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,116,8))

def f1186(E, r, c):
    # Ô E98
    # = IF(A98="","",LUU_NIEN!$AA$116)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,116,27))

def f1187(E, r, c):
    # Ô F98
    # = IF(A98="","",LUU_NIEN!$AL$116)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,116,38))

def f1188(E, r, c):
    # Ô K98
    # = IF(A98="","",TRIM(IF(AB98,"Dấu hiệu bất lợi cho cha. ","")&IF(AC98,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA98,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U98>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$116))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,116,66))))) else ''))))

def f1189(E, r, c):
    # Ô L98
    # = IF(OR(A98="",LUU_NIEN!$CN$116=""),"",IFERROR(LEFT(LUU_NIEN!$CN$116,FIND(":",LUU_NIEN!$CN$116)-1),LEFT(LUU_NIEN!$CN$116,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,116,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,116,92), X.sub(X.FIND(':', E.v(15,116,92)),1.0)), lambda: X.LEFT(E.v(15,116,92), 120.0)))

def f1190(E, r, c):
    # Ô M98
    # = IF(A98="","",IF(LEN(LUU_NIEN!$BN$116)>260,LEFT(LUU_NIEN!$BN$116,260)&"…",LUU_NIEN!$BN$116))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,116,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,116,66)),260.0)) else E.v(15,116,66)))

def f1191(E, r, c):
    # Ô T98
    # = IF(OR(A98="",LUU_NIEN!$P$116=""),-1,MOD(LUU_NIEN!$P$116-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,116,16),''))) else X.MOD(X.add(X.sub(E.v(15,116,16),E.v(23,7,5)),5.0), 5.0))

def f1192(E, r, c):
    # Ô AA98
    # = IF(A98="",FALSE(),AND($U98>=18,$U98<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$116))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,116,66)))))

def f1193(E, r, c):
    # Ô AB98
    # = IF(A98="",FALSE(),AND($U98>=0,$U98<=120,LUU_NIEN!$BH$116=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,116,60),1.0)))

def f1194(E, r, c):
    # Ô AC98
    # = IF(A98="",FALSE(),AND($U98>=0,$U98<=120,LUU_NIEN!$BI$116=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,116,61),1.0)))

def f1195(E, r, c):
    # Ô AD98
    # = IF(A98="",FALSE(),AND($U98>=0,$U98<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$116))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,116,92)))))

def f1196(E, r, c):
    # Ô A99
    # = IF((ISNUMBER(LUU_NIEN!$A$117)),LUU_NIEN!$A$117,"")
    return (E.v(15,117,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,117,1))) else '')

def f1197(E, r, c):
    # Ô C99
    # = IF(A99="","",LUU_NIEN!$L$117)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,117,12))

def f1198(E, r, c):
    # Ô D99
    # = IF(A99="","",LUU_NIEN!$H$117)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,117,8))

def f1199(E, r, c):
    # Ô E99
    # = IF(A99="","",LUU_NIEN!$AA$117)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,117,27))

def f1200(E, r, c):
    # Ô F99
    # = IF(A99="","",LUU_NIEN!$AL$117)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,117,38))

def f1201(E, r, c):
    # Ô K99
    # = IF(A99="","",TRIM(IF(AB99,"Dấu hiệu bất lợi cho cha. ","")&IF(AC99,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA99,"Thuận kết hôn / tin vui hôn nhân. ","")&IF
    #   (AND(U99>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$117))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,117,66))))) else ''))))

def f1202(E, r, c):
    # Ô L99
    # = IF(OR(A99="",LUU_NIEN!$CN$117=""),"",IFERROR(LEFT(LUU_NIEN!$CN$117,FIND(":",LUU_NIEN!$CN$117)-1),LEFT(LUU_NIEN!$CN$117,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,117,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,117,92), X.sub(X.FIND(':', E.v(15,117,92)),1.0)), lambda: X.LEFT(E.v(15,117,92), 120.0)))

def f1203(E, r, c):
    # Ô M99
    # = IF(A99="","",IF(LEN(LUU_NIEN!$BN$117)>260,LEFT(LUU_NIEN!$BN$117,260)&"…",LUU_NIEN!$BN$117))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,117,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,117,66)),260.0)) else E.v(15,117,66)))

def f1204(E, r, c):
    # Ô T99
    # = IF(OR(A99="",LUU_NIEN!$P$117=""),-1,MOD(LUU_NIEN!$P$117-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,117,16),''))) else X.MOD(X.add(X.sub(E.v(15,117,16),E.v(23,7,5)),5.0), 5.0))

def f1205(E, r, c):
    # Ô AA99
    # = IF(A99="",FALSE(),AND($U99>=18,$U99<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$117))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,117,66)))))

def f1206(E, r, c):
    # Ô AB99
    # = IF(A99="",FALSE(),AND($U99>=0,$U99<=120,LUU_NIEN!$BH$117=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,117,60),1.0)))

def f1207(E, r, c):
    # Ô AC99
    # = IF(A99="",FALSE(),AND($U99>=0,$U99<=120,LUU_NIEN!$BI$117=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,117,61),1.0)))

def f1208(E, r, c):
    # Ô AD99
    # = IF(A99="",FALSE(),AND($U99>=0,$U99<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$117))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,117,92)))))

def f1209(E, r, c):
    # Ô A100
    # = IF((ISNUMBER(LUU_NIEN!$A$118)),LUU_NIEN!$A$118,"")
    return (E.v(15,118,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,118,1))) else '')

def f1210(E, r, c):
    # Ô C100
    # = IF(A100="","",LUU_NIEN!$L$118)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,118,12))

def f1211(E, r, c):
    # Ô D100
    # = IF(A100="","",LUU_NIEN!$H$118)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,118,8))

def f1212(E, r, c):
    # Ô E100
    # = IF(A100="","",LUU_NIEN!$AA$118)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,118,27))

def f1213(E, r, c):
    # Ô F100
    # = IF(A100="","",LUU_NIEN!$AL$118)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,118,38))

def f1214(E, r, c):
    # Ô K100
    # = IF(A100="","",TRIM(IF(AB100,"Dấu hiệu bất lợi cho cha. ","")&IF(AC100,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA100,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U100>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$118))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,118,66))))) else ''))))

def f1215(E, r, c):
    # Ô L100
    # = IF(OR(A100="",LUU_NIEN!$CN$118=""),"",IFERROR(LEFT(LUU_NIEN!$CN$118,FIND(":",LUU_NIEN!$CN$118)-1),LEFT(LUU_NIEN!$CN$118,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,118,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,118,92), X.sub(X.FIND(':', E.v(15,118,92)),1.0)), lambda: X.LEFT(E.v(15,118,92), 120.0)))

def f1216(E, r, c):
    # Ô M100
    # = IF(A100="","",IF(LEN(LUU_NIEN!$BN$118)>260,LEFT(LUU_NIEN!$BN$118,260)&"…",LUU_NIEN!$BN$118))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,118,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,118,66)),260.0)) else E.v(15,118,66)))

def f1217(E, r, c):
    # Ô T100
    # = IF(OR(A100="",LUU_NIEN!$P$118=""),-1,MOD(LUU_NIEN!$P$118-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,118,16),''))) else X.MOD(X.add(X.sub(E.v(15,118,16),E.v(23,7,5)),5.0), 5.0))

def f1218(E, r, c):
    # Ô AA100
    # = IF(A100="",FALSE(),AND($U100>=18,$U100<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$118))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,118,66)))))

def f1219(E, r, c):
    # Ô AB100
    # = IF(A100="",FALSE(),AND($U100>=0,$U100<=120,LUU_NIEN!$BH$118=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,118,60),1.0)))

def f1220(E, r, c):
    # Ô AC100
    # = IF(A100="",FALSE(),AND($U100>=0,$U100<=120,LUU_NIEN!$BI$118=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,118,61),1.0)))

def f1221(E, r, c):
    # Ô AD100
    # = IF(A100="",FALSE(),AND($U100>=0,$U100<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$118))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,118,92)))))

def f1222(E, r, c):
    # Ô A101
    # = IF((ISNUMBER(LUU_NIEN!$A$119)),LUU_NIEN!$A$119,"")
    return (E.v(15,119,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,119,1))) else '')

def f1223(E, r, c):
    # Ô C101
    # = IF(A101="","",LUU_NIEN!$L$119)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,119,12))

def f1224(E, r, c):
    # Ô D101
    # = IF(A101="","",LUU_NIEN!$H$119)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,119,8))

def f1225(E, r, c):
    # Ô E101
    # = IF(A101="","",LUU_NIEN!$AA$119)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,119,27))

def f1226(E, r, c):
    # Ô F101
    # = IF(A101="","",LUU_NIEN!$AL$119)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,119,38))

def f1227(E, r, c):
    # Ô K101
    # = IF(A101="","",TRIM(IF(AB101,"Dấu hiệu bất lợi cho cha. ","")&IF(AC101,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA101,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U101>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$119))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,119,66))))) else ''))))

def f1228(E, r, c):
    # Ô L101
    # = IF(OR(A101="",LUU_NIEN!$CN$119=""),"",IFERROR(LEFT(LUU_NIEN!$CN$119,FIND(":",LUU_NIEN!$CN$119)-1),LEFT(LUU_NIEN!$CN$119,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,119,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,119,92), X.sub(X.FIND(':', E.v(15,119,92)),1.0)), lambda: X.LEFT(E.v(15,119,92), 120.0)))

def f1229(E, r, c):
    # Ô M101
    # = IF(A101="","",IF(LEN(LUU_NIEN!$BN$119)>260,LEFT(LUU_NIEN!$BN$119,260)&"…",LUU_NIEN!$BN$119))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,119,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,119,66)),260.0)) else E.v(15,119,66)))

def f1230(E, r, c):
    # Ô T101
    # = IF(OR(A101="",LUU_NIEN!$P$119=""),-1,MOD(LUU_NIEN!$P$119-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,119,16),''))) else X.MOD(X.add(X.sub(E.v(15,119,16),E.v(23,7,5)),5.0), 5.0))

def f1231(E, r, c):
    # Ô AA101
    # = IF(A101="",FALSE(),AND($U101>=18,$U101<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$119))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,119,66)))))

def f1232(E, r, c):
    # Ô AB101
    # = IF(A101="",FALSE(),AND($U101>=0,$U101<=120,LUU_NIEN!$BH$119=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,119,60),1.0)))

def f1233(E, r, c):
    # Ô AC101
    # = IF(A101="",FALSE(),AND($U101>=0,$U101<=120,LUU_NIEN!$BI$119=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,119,61),1.0)))

def f1234(E, r, c):
    # Ô AD101
    # = IF(A101="",FALSE(),AND($U101>=0,$U101<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$119))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,119,92)))))

def f1235(E, r, c):
    # Ô A102
    # = IF((ISNUMBER(LUU_NIEN!$A$120)),LUU_NIEN!$A$120,"")
    return (E.v(15,120,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,120,1))) else '')

def f1236(E, r, c):
    # Ô C102
    # = IF(A102="","",LUU_NIEN!$L$120)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,120,12))

def f1237(E, r, c):
    # Ô D102
    # = IF(A102="","",LUU_NIEN!$H$120)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,120,8))

def f1238(E, r, c):
    # Ô E102
    # = IF(A102="","",LUU_NIEN!$AA$120)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,120,27))

def f1239(E, r, c):
    # Ô F102
    # = IF(A102="","",LUU_NIEN!$AL$120)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,120,38))

def f1240(E, r, c):
    # Ô K102
    # = IF(A102="","",TRIM(IF(AB102,"Dấu hiệu bất lợi cho cha. ","")&IF(AC102,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA102,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U102>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$120))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,120,66))))) else ''))))

def f1241(E, r, c):
    # Ô L102
    # = IF(OR(A102="",LUU_NIEN!$CN$120=""),"",IFERROR(LEFT(LUU_NIEN!$CN$120,FIND(":",LUU_NIEN!$CN$120)-1),LEFT(LUU_NIEN!$CN$120,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,120,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,120,92), X.sub(X.FIND(':', E.v(15,120,92)),1.0)), lambda: X.LEFT(E.v(15,120,92), 120.0)))

def f1242(E, r, c):
    # Ô M102
    # = IF(A102="","",IF(LEN(LUU_NIEN!$BN$120)>260,LEFT(LUU_NIEN!$BN$120,260)&"…",LUU_NIEN!$BN$120))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,120,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,120,66)),260.0)) else E.v(15,120,66)))

def f1243(E, r, c):
    # Ô T102
    # = IF(OR(A102="",LUU_NIEN!$P$120=""),-1,MOD(LUU_NIEN!$P$120-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,120,16),''))) else X.MOD(X.add(X.sub(E.v(15,120,16),E.v(23,7,5)),5.0), 5.0))

def f1244(E, r, c):
    # Ô AA102
    # = IF(A102="",FALSE(),AND($U102>=18,$U102<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$120))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,120,66)))))

def f1245(E, r, c):
    # Ô AB102
    # = IF(A102="",FALSE(),AND($U102>=0,$U102<=120,LUU_NIEN!$BH$120=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,120,60),1.0)))

def f1246(E, r, c):
    # Ô AC102
    # = IF(A102="",FALSE(),AND($U102>=0,$U102<=120,LUU_NIEN!$BI$120=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,120,61),1.0)))

def f1247(E, r, c):
    # Ô AD102
    # = IF(A102="",FALSE(),AND($U102>=0,$U102<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$120))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,120,92)))))

def f1248(E, r, c):
    # Ô A103
    # = IF((ISNUMBER(LUU_NIEN!$A$121)),LUU_NIEN!$A$121,"")
    return (E.v(15,121,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,121,1))) else '')

def f1249(E, r, c):
    # Ô C103
    # = IF(A103="","",LUU_NIEN!$L$121)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,121,12))

def f1250(E, r, c):
    # Ô D103
    # = IF(A103="","",LUU_NIEN!$H$121)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,121,8))

def f1251(E, r, c):
    # Ô E103
    # = IF(A103="","",LUU_NIEN!$AA$121)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,121,27))

def f1252(E, r, c):
    # Ô F103
    # = IF(A103="","",LUU_NIEN!$AL$121)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,121,38))

def f1253(E, r, c):
    # Ô K103
    # = IF(A103="","",TRIM(IF(AB103,"Dấu hiệu bất lợi cho cha. ","")&IF(AC103,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA103,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U103>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$121))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,121,66))))) else ''))))

def f1254(E, r, c):
    # Ô L103
    # = IF(OR(A103="",LUU_NIEN!$CN$121=""),"",IFERROR(LEFT(LUU_NIEN!$CN$121,FIND(":",LUU_NIEN!$CN$121)-1),LEFT(LUU_NIEN!$CN$121,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,121,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,121,92), X.sub(X.FIND(':', E.v(15,121,92)),1.0)), lambda: X.LEFT(E.v(15,121,92), 120.0)))

def f1255(E, r, c):
    # Ô M103
    # = IF(A103="","",IF(LEN(LUU_NIEN!$BN$121)>260,LEFT(LUU_NIEN!$BN$121,260)&"…",LUU_NIEN!$BN$121))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,121,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,121,66)),260.0)) else E.v(15,121,66)))

def f1256(E, r, c):
    # Ô T103
    # = IF(OR(A103="",LUU_NIEN!$P$121=""),-1,MOD(LUU_NIEN!$P$121-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,121,16),''))) else X.MOD(X.add(X.sub(E.v(15,121,16),E.v(23,7,5)),5.0), 5.0))

def f1257(E, r, c):
    # Ô AA103
    # = IF(A103="",FALSE(),AND($U103>=18,$U103<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$121))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,121,66)))))

def f1258(E, r, c):
    # Ô AB103
    # = IF(A103="",FALSE(),AND($U103>=0,$U103<=120,LUU_NIEN!$BH$121=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,121,60),1.0)))

def f1259(E, r, c):
    # Ô AC103
    # = IF(A103="",FALSE(),AND($U103>=0,$U103<=120,LUU_NIEN!$BI$121=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,121,61),1.0)))

def f1260(E, r, c):
    # Ô AD103
    # = IF(A103="",FALSE(),AND($U103>=0,$U103<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$121))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,121,92)))))

def f1261(E, r, c):
    # Ô A104
    # = IF((ISNUMBER(LUU_NIEN!$A$122)),LUU_NIEN!$A$122,"")
    return (E.v(15,122,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,122,1))) else '')

def f1262(E, r, c):
    # Ô C104
    # = IF(A104="","",LUU_NIEN!$L$122)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,122,12))

def f1263(E, r, c):
    # Ô D104
    # = IF(A104="","",LUU_NIEN!$H$122)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,122,8))

def f1264(E, r, c):
    # Ô E104
    # = IF(A104="","",LUU_NIEN!$AA$122)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,122,27))

def f1265(E, r, c):
    # Ô F104
    # = IF(A104="","",LUU_NIEN!$AL$122)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,122,38))

def f1266(E, r, c):
    # Ô K104
    # = IF(A104="","",TRIM(IF(AB104,"Dấu hiệu bất lợi cho cha. ","")&IF(AC104,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA104,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U104>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$122))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,122,66))))) else ''))))

def f1267(E, r, c):
    # Ô L104
    # = IF(OR(A104="",LUU_NIEN!$CN$122=""),"",IFERROR(LEFT(LUU_NIEN!$CN$122,FIND(":",LUU_NIEN!$CN$122)-1),LEFT(LUU_NIEN!$CN$122,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,122,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,122,92), X.sub(X.FIND(':', E.v(15,122,92)),1.0)), lambda: X.LEFT(E.v(15,122,92), 120.0)))

def f1268(E, r, c):
    # Ô M104
    # = IF(A104="","",IF(LEN(LUU_NIEN!$BN$122)>260,LEFT(LUU_NIEN!$BN$122,260)&"…",LUU_NIEN!$BN$122))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,122,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,122,66)),260.0)) else E.v(15,122,66)))

def f1269(E, r, c):
    # Ô T104
    # = IF(OR(A104="",LUU_NIEN!$P$122=""),-1,MOD(LUU_NIEN!$P$122-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,122,16),''))) else X.MOD(X.add(X.sub(E.v(15,122,16),E.v(23,7,5)),5.0), 5.0))

def f1270(E, r, c):
    # Ô AA104
    # = IF(A104="",FALSE(),AND($U104>=18,$U104<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$122))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,122,66)))))

def f1271(E, r, c):
    # Ô AB104
    # = IF(A104="",FALSE(),AND($U104>=0,$U104<=120,LUU_NIEN!$BH$122=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,122,60),1.0)))

def f1272(E, r, c):
    # Ô AC104
    # = IF(A104="",FALSE(),AND($U104>=0,$U104<=120,LUU_NIEN!$BI$122=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,122,61),1.0)))

def f1273(E, r, c):
    # Ô AD104
    # = IF(A104="",FALSE(),AND($U104>=0,$U104<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$122))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,122,92)))))

def f1274(E, r, c):
    # Ô A105
    # = IF((ISNUMBER(LUU_NIEN!$A$123)),LUU_NIEN!$A$123,"")
    return (E.v(15,123,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,123,1))) else '')

def f1275(E, r, c):
    # Ô C105
    # = IF(A105="","",LUU_NIEN!$L$123)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,123,12))

def f1276(E, r, c):
    # Ô D105
    # = IF(A105="","",LUU_NIEN!$H$123)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,123,8))

def f1277(E, r, c):
    # Ô E105
    # = IF(A105="","",LUU_NIEN!$AA$123)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,123,27))

def f1278(E, r, c):
    # Ô F105
    # = IF(A105="","",LUU_NIEN!$AL$123)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,123,38))

def f1279(E, r, c):
    # Ô K105
    # = IF(A105="","",TRIM(IF(AB105,"Dấu hiệu bất lợi cho cha. ","")&IF(AC105,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA105,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U105>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$123))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,123,66))))) else ''))))

def f1280(E, r, c):
    # Ô L105
    # = IF(OR(A105="",LUU_NIEN!$CN$123=""),"",IFERROR(LEFT(LUU_NIEN!$CN$123,FIND(":",LUU_NIEN!$CN$123)-1),LEFT(LUU_NIEN!$CN$123,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,123,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,123,92), X.sub(X.FIND(':', E.v(15,123,92)),1.0)), lambda: X.LEFT(E.v(15,123,92), 120.0)))

def f1281(E, r, c):
    # Ô M105
    # = IF(A105="","",IF(LEN(LUU_NIEN!$BN$123)>260,LEFT(LUU_NIEN!$BN$123,260)&"…",LUU_NIEN!$BN$123))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,123,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,123,66)),260.0)) else E.v(15,123,66)))

def f1282(E, r, c):
    # Ô T105
    # = IF(OR(A105="",LUU_NIEN!$P$123=""),-1,MOD(LUU_NIEN!$P$123-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,123,16),''))) else X.MOD(X.add(X.sub(E.v(15,123,16),E.v(23,7,5)),5.0), 5.0))

def f1283(E, r, c):
    # Ô AA105
    # = IF(A105="",FALSE(),AND($U105>=18,$U105<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$123))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,123,66)))))

def f1284(E, r, c):
    # Ô AB105
    # = IF(A105="",FALSE(),AND($U105>=0,$U105<=120,LUU_NIEN!$BH$123=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,123,60),1.0)))

def f1285(E, r, c):
    # Ô AC105
    # = IF(A105="",FALSE(),AND($U105>=0,$U105<=120,LUU_NIEN!$BI$123=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,123,61),1.0)))

def f1286(E, r, c):
    # Ô AD105
    # = IF(A105="",FALSE(),AND($U105>=0,$U105<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$123))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,123,92)))))

def f1287(E, r, c):
    # Ô A106
    # = IF((ISNUMBER(LUU_NIEN!$A$124)),LUU_NIEN!$A$124,"")
    return (E.v(15,124,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,124,1))) else '')

def f1288(E, r, c):
    # Ô C106
    # = IF(A106="","",LUU_NIEN!$L$124)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,124,12))

def f1289(E, r, c):
    # Ô D106
    # = IF(A106="","",LUU_NIEN!$H$124)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,124,8))

def f1290(E, r, c):
    # Ô E106
    # = IF(A106="","",LUU_NIEN!$AA$124)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,124,27))

def f1291(E, r, c):
    # Ô F106
    # = IF(A106="","",LUU_NIEN!$AL$124)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,124,38))

def f1292(E, r, c):
    # Ô K106
    # = IF(A106="","",TRIM(IF(AB106,"Dấu hiệu bất lợi cho cha. ","")&IF(AC106,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA106,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U106>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$124))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,124,66))))) else ''))))

def f1293(E, r, c):
    # Ô L106
    # = IF(OR(A106="",LUU_NIEN!$CN$124=""),"",IFERROR(LEFT(LUU_NIEN!$CN$124,FIND(":",LUU_NIEN!$CN$124)-1),LEFT(LUU_NIEN!$CN$124,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,124,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,124,92), X.sub(X.FIND(':', E.v(15,124,92)),1.0)), lambda: X.LEFT(E.v(15,124,92), 120.0)))

def f1294(E, r, c):
    # Ô M106
    # = IF(A106="","",IF(LEN(LUU_NIEN!$BN$124)>260,LEFT(LUU_NIEN!$BN$124,260)&"…",LUU_NIEN!$BN$124))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,124,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,124,66)),260.0)) else E.v(15,124,66)))

def f1295(E, r, c):
    # Ô T106
    # = IF(OR(A106="",LUU_NIEN!$P$124=""),-1,MOD(LUU_NIEN!$P$124-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,124,16),''))) else X.MOD(X.add(X.sub(E.v(15,124,16),E.v(23,7,5)),5.0), 5.0))

def f1296(E, r, c):
    # Ô AA106
    # = IF(A106="",FALSE(),AND($U106>=18,$U106<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$124))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,124,66)))))

def f1297(E, r, c):
    # Ô AB106
    # = IF(A106="",FALSE(),AND($U106>=0,$U106<=120,LUU_NIEN!$BH$124=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,124,60),1.0)))

def f1298(E, r, c):
    # Ô AC106
    # = IF(A106="",FALSE(),AND($U106>=0,$U106<=120,LUU_NIEN!$BI$124=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,124,61),1.0)))

def f1299(E, r, c):
    # Ô AD106
    # = IF(A106="",FALSE(),AND($U106>=0,$U106<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$124))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,124,92)))))

def f1300(E, r, c):
    # Ô A107
    # = IF((ISNUMBER(LUU_NIEN!$A$125)),LUU_NIEN!$A$125,"")
    return (E.v(15,125,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,125,1))) else '')

def f1301(E, r, c):
    # Ô C107
    # = IF(A107="","",LUU_NIEN!$L$125)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,125,12))

def f1302(E, r, c):
    # Ô D107
    # = IF(A107="","",LUU_NIEN!$H$125)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,125,8))

def f1303(E, r, c):
    # Ô E107
    # = IF(A107="","",LUU_NIEN!$AA$125)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,125,27))

def f1304(E, r, c):
    # Ô F107
    # = IF(A107="","",LUU_NIEN!$AL$125)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,125,38))

def f1305(E, r, c):
    # Ô K107
    # = IF(A107="","",TRIM(IF(AB107,"Dấu hiệu bất lợi cho cha. ","")&IF(AC107,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA107,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U107>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$125))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,125,66))))) else ''))))

def f1306(E, r, c):
    # Ô L107
    # = IF(OR(A107="",LUU_NIEN!$CN$125=""),"",IFERROR(LEFT(LUU_NIEN!$CN$125,FIND(":",LUU_NIEN!$CN$125)-1),LEFT(LUU_NIEN!$CN$125,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,125,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,125,92), X.sub(X.FIND(':', E.v(15,125,92)),1.0)), lambda: X.LEFT(E.v(15,125,92), 120.0)))

def f1307(E, r, c):
    # Ô M107
    # = IF(A107="","",IF(LEN(LUU_NIEN!$BN$125)>260,LEFT(LUU_NIEN!$BN$125,260)&"…",LUU_NIEN!$BN$125))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,125,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,125,66)),260.0)) else E.v(15,125,66)))

def f1308(E, r, c):
    # Ô T107
    # = IF(OR(A107="",LUU_NIEN!$P$125=""),-1,MOD(LUU_NIEN!$P$125-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,125,16),''))) else X.MOD(X.add(X.sub(E.v(15,125,16),E.v(23,7,5)),5.0), 5.0))

def f1309(E, r, c):
    # Ô AA107
    # = IF(A107="",FALSE(),AND($U107>=18,$U107<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$125))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,125,66)))))

def f1310(E, r, c):
    # Ô AB107
    # = IF(A107="",FALSE(),AND($U107>=0,$U107<=120,LUU_NIEN!$BH$125=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,125,60),1.0)))

def f1311(E, r, c):
    # Ô AC107
    # = IF(A107="",FALSE(),AND($U107>=0,$U107<=120,LUU_NIEN!$BI$125=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,125,61),1.0)))

def f1312(E, r, c):
    # Ô AD107
    # = IF(A107="",FALSE(),AND($U107>=0,$U107<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$125))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,125,92)))))

def f1313(E, r, c):
    # Ô A108
    # = IF((ISNUMBER(LUU_NIEN!$A$126)),LUU_NIEN!$A$126,"")
    return (E.v(15,126,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,126,1))) else '')

def f1314(E, r, c):
    # Ô C108
    # = IF(A108="","",LUU_NIEN!$L$126)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,126,12))

def f1315(E, r, c):
    # Ô D108
    # = IF(A108="","",LUU_NIEN!$H$126)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,126,8))

def f1316(E, r, c):
    # Ô E108
    # = IF(A108="","",LUU_NIEN!$AA$126)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,126,27))

def f1317(E, r, c):
    # Ô F108
    # = IF(A108="","",LUU_NIEN!$AL$126)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,126,38))

def f1318(E, r, c):
    # Ô K108
    # = IF(A108="","",TRIM(IF(AB108,"Dấu hiệu bất lợi cho cha. ","")&IF(AC108,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA108,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U108>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$126))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,126,66))))) else ''))))

def f1319(E, r, c):
    # Ô L108
    # = IF(OR(A108="",LUU_NIEN!$CN$126=""),"",IFERROR(LEFT(LUU_NIEN!$CN$126,FIND(":",LUU_NIEN!$CN$126)-1),LEFT(LUU_NIEN!$CN$126,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,126,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,126,92), X.sub(X.FIND(':', E.v(15,126,92)),1.0)), lambda: X.LEFT(E.v(15,126,92), 120.0)))

def f1320(E, r, c):
    # Ô M108
    # = IF(A108="","",IF(LEN(LUU_NIEN!$BN$126)>260,LEFT(LUU_NIEN!$BN$126,260)&"…",LUU_NIEN!$BN$126))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,126,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,126,66)),260.0)) else E.v(15,126,66)))

def f1321(E, r, c):
    # Ô T108
    # = IF(OR(A108="",LUU_NIEN!$P$126=""),-1,MOD(LUU_NIEN!$P$126-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,126,16),''))) else X.MOD(X.add(X.sub(E.v(15,126,16),E.v(23,7,5)),5.0), 5.0))

def f1322(E, r, c):
    # Ô AA108
    # = IF(A108="",FALSE(),AND($U108>=18,$U108<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$126))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,126,66)))))

def f1323(E, r, c):
    # Ô AB108
    # = IF(A108="",FALSE(),AND($U108>=0,$U108<=120,LUU_NIEN!$BH$126=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,126,60),1.0)))

def f1324(E, r, c):
    # Ô AC108
    # = IF(A108="",FALSE(),AND($U108>=0,$U108<=120,LUU_NIEN!$BI$126=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,126,61),1.0)))

def f1325(E, r, c):
    # Ô AD108
    # = IF(A108="",FALSE(),AND($U108>=0,$U108<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$126))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,126,92)))))

def f1326(E, r, c):
    # Ô A109
    # = IF((ISNUMBER(LUU_NIEN!$A$127)),LUU_NIEN!$A$127,"")
    return (E.v(15,127,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,127,1))) else '')

def f1327(E, r, c):
    # Ô C109
    # = IF(A109="","",LUU_NIEN!$L$127)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,127,12))

def f1328(E, r, c):
    # Ô D109
    # = IF(A109="","",LUU_NIEN!$H$127)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,127,8))

def f1329(E, r, c):
    # Ô E109
    # = IF(A109="","",LUU_NIEN!$AA$127)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,127,27))

def f1330(E, r, c):
    # Ô F109
    # = IF(A109="","",LUU_NIEN!$AL$127)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,127,38))

def f1331(E, r, c):
    # Ô K109
    # = IF(A109="","",TRIM(IF(AB109,"Dấu hiệu bất lợi cho cha. ","")&IF(AC109,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA109,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U109>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$127))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,127,66))))) else ''))))

def f1332(E, r, c):
    # Ô L109
    # = IF(OR(A109="",LUU_NIEN!$CN$127=""),"",IFERROR(LEFT(LUU_NIEN!$CN$127,FIND(":",LUU_NIEN!$CN$127)-1),LEFT(LUU_NIEN!$CN$127,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,127,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,127,92), X.sub(X.FIND(':', E.v(15,127,92)),1.0)), lambda: X.LEFT(E.v(15,127,92), 120.0)))

def f1333(E, r, c):
    # Ô M109
    # = IF(A109="","",IF(LEN(LUU_NIEN!$BN$127)>260,LEFT(LUU_NIEN!$BN$127,260)&"…",LUU_NIEN!$BN$127))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,127,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,127,66)),260.0)) else E.v(15,127,66)))

def f1334(E, r, c):
    # Ô T109
    # = IF(OR(A109="",LUU_NIEN!$P$127=""),-1,MOD(LUU_NIEN!$P$127-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,127,16),''))) else X.MOD(X.add(X.sub(E.v(15,127,16),E.v(23,7,5)),5.0), 5.0))

def f1335(E, r, c):
    # Ô AA109
    # = IF(A109="",FALSE(),AND($U109>=18,$U109<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$127))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,127,66)))))

def f1336(E, r, c):
    # Ô AB109
    # = IF(A109="",FALSE(),AND($U109>=0,$U109<=120,LUU_NIEN!$BH$127=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,127,60),1.0)))

def f1337(E, r, c):
    # Ô AC109
    # = IF(A109="",FALSE(),AND($U109>=0,$U109<=120,LUU_NIEN!$BI$127=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,127,61),1.0)))

def f1338(E, r, c):
    # Ô AD109
    # = IF(A109="",FALSE(),AND($U109>=0,$U109<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$127))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,127,92)))))

def f1339(E, r, c):
    # Ô A110
    # = IF((ISNUMBER(LUU_NIEN!$A$128)),LUU_NIEN!$A$128,"")
    return (E.v(15,128,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,128,1))) else '')

def f1340(E, r, c):
    # Ô C110
    # = IF(A110="","",LUU_NIEN!$L$128)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,128,12))

def f1341(E, r, c):
    # Ô D110
    # = IF(A110="","",LUU_NIEN!$H$128)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,128,8))

def f1342(E, r, c):
    # Ô E110
    # = IF(A110="","",LUU_NIEN!$AA$128)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,128,27))

def f1343(E, r, c):
    # Ô F110
    # = IF(A110="","",LUU_NIEN!$AL$128)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,128,38))

def f1344(E, r, c):
    # Ô K110
    # = IF(A110="","",TRIM(IF(AB110,"Dấu hiệu bất lợi cho cha. ","")&IF(AC110,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA110,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U110>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$128))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,128,66))))) else ''))))

def f1345(E, r, c):
    # Ô L110
    # = IF(OR(A110="",LUU_NIEN!$CN$128=""),"",IFERROR(LEFT(LUU_NIEN!$CN$128,FIND(":",LUU_NIEN!$CN$128)-1),LEFT(LUU_NIEN!$CN$128,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,128,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,128,92), X.sub(X.FIND(':', E.v(15,128,92)),1.0)), lambda: X.LEFT(E.v(15,128,92), 120.0)))

def f1346(E, r, c):
    # Ô M110
    # = IF(A110="","",IF(LEN(LUU_NIEN!$BN$128)>260,LEFT(LUU_NIEN!$BN$128,260)&"…",LUU_NIEN!$BN$128))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,128,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,128,66)),260.0)) else E.v(15,128,66)))

def f1347(E, r, c):
    # Ô T110
    # = IF(OR(A110="",LUU_NIEN!$P$128=""),-1,MOD(LUU_NIEN!$P$128-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,128,16),''))) else X.MOD(X.add(X.sub(E.v(15,128,16),E.v(23,7,5)),5.0), 5.0))

def f1348(E, r, c):
    # Ô AA110
    # = IF(A110="",FALSE(),AND($U110>=18,$U110<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$128))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,128,66)))))

def f1349(E, r, c):
    # Ô AB110
    # = IF(A110="",FALSE(),AND($U110>=0,$U110<=120,LUU_NIEN!$BH$128=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,128,60),1.0)))

def f1350(E, r, c):
    # Ô AC110
    # = IF(A110="",FALSE(),AND($U110>=0,$U110<=120,LUU_NIEN!$BI$128=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,128,61),1.0)))

def f1351(E, r, c):
    # Ô AD110
    # = IF(A110="",FALSE(),AND($U110>=0,$U110<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$128))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,128,92)))))

def f1352(E, r, c):
    # Ô A111
    # = IF((ISNUMBER(LUU_NIEN!$A$129)),LUU_NIEN!$A$129,"")
    return (E.v(15,129,1) if X.truthy(X.ISNUMBER(lambda: E.v(15,129,1))) else '')

def f1353(E, r, c):
    # Ô C111
    # = IF(A111="","",LUU_NIEN!$L$129)
    return ('' if X.truthy(X.eq(E.v(5,r,c-2),'')) else E.v(15,129,12))

def f1354(E, r, c):
    # Ô D111
    # = IF(A111="","",LUU_NIEN!$H$129)
    return ('' if X.truthy(X.eq(E.v(5,r,c-3),'')) else E.v(15,129,8))

def f1355(E, r, c):
    # Ô E111
    # = IF(A111="","",LUU_NIEN!$AA$129)
    return ('' if X.truthy(X.eq(E.v(5,r,c-4),'')) else E.v(15,129,27))

def f1356(E, r, c):
    # Ô F111
    # = IF(A111="","",LUU_NIEN!$AL$129)
    return ('' if X.truthy(X.eq(E.v(5,r,c-5),'')) else E.v(15,129,38))

def f1357(E, r, c):
    # Ô K111
    # = IF(A111="","",TRIM(IF(AB111,"Dấu hiệu bất lợi cho cha. ","")&IF(AC111,"Dấu hiệu bất lợi cho mẹ. ","")&IF(AA111,"Thuận kết hôn / tin vui hôn nhân. ",""
    #   )&IF(AND(U111>=18,ISNUMBER(SEARCH("Bất lợi hôn nhân",LUU_NIEN!$BN$129))),"Bất lợi hôn nhân – cần nhường nhịn.","")))
    return ('' if X.truthy(X.eq(E.v(5,r,c-10),'')) else X.TRIM(X.cat(X.cat(X.cat(('Dấu hiệu bất lợi cho cha. ' if X.truthy(E.v(5,r,c+17)) else ''),('Dấu hiệu bất lợi cho mẹ. ' if X.truthy(E.v(5,r,c+18)) else '')),('Thuận kết hôn / tin vui hôn nhân. ' if X.truthy(E.v(5,r,c+16)) else '')),('Bất lợi hôn nhân – cần nhường nhịn.' if X.truthy(X.AND(X.ge(E.v(5,r,c+10),18.0), X.ISNUMBER(lambda: X.SEARCH('Bất lợi hôn nhân', E.v(15,129,66))))) else ''))))

def f1358(E, r, c):
    # Ô L111
    # = IF(OR(A111="",LUU_NIEN!$CN$129=""),"",IFERROR(LEFT(LUU_NIEN!$CN$129,FIND(":",LUU_NIEN!$CN$129)-1),LEFT(LUU_NIEN!$CN$129,120)))
    return ('' if X.truthy(X.OR(X.eq(E.v(5,r,c-11),''), X.eq(E.v(15,129,92),''))) else X.IFERROR(lambda: X.LEFT(E.v(15,129,92), X.sub(X.FIND(':', E.v(15,129,92)),1.0)), lambda: X.LEFT(E.v(15,129,92), 120.0)))

def f1359(E, r, c):
    # Ô M111
    # = IF(A111="","",IF(LEN(LUU_NIEN!$BN$129)>260,LEFT(LUU_NIEN!$BN$129,260)&"…",LUU_NIEN!$BN$129))
    return ('' if X.truthy(X.eq(E.v(5,r,c-12),'')) else (X.cat(X.LEFT(E.v(15,129,66), 260.0),'…') if X.truthy(X.gt(X.LEN(E.v(15,129,66)),260.0)) else E.v(15,129,66)))

def f1360(E, r, c):
    # Ô T111
    # = IF(OR(A111="",LUU_NIEN!$P$129=""),-1,MOD(LUU_NIEN!$P$129-TRUNG_GIAN!$E$7+5,5))
    return (X.neg(1.0) if X.truthy(X.OR(X.eq(E.v(5,r,c-19),''), X.eq(E.v(15,129,16),''))) else X.MOD(X.add(X.sub(E.v(15,129,16),E.v(23,7,5)),5.0), 5.0))

def f1361(E, r, c):
    # Ô AA111
    # = IF(A111="",FALSE(),AND($U111>=18,$U111<=50,ISNUMBER(SEARCH("thuận lợi kết hôn",LUU_NIEN!$BN$129))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-26),'')) else X.AND(X.ge(E.v(5,r,21),18.0), X.le(E.v(5,r,21),50.0), X.ISNUMBER(lambda: X.SEARCH('thuận lợi kết hôn', E.v(15,129,66)))))

def f1362(E, r, c):
    # Ô AB111
    # = IF(A111="",FALSE(),AND($U111>=0,$U111<=120,LUU_NIEN!$BH$129=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-27),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,129,60),1.0)))

def f1363(E, r, c):
    # Ô AC111
    # = IF(A111="",FALSE(),AND($U111>=0,$U111<=120,LUU_NIEN!$BI$129=1))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-28),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.eq(E.v(15,129,61),1.0)))

def f1364(E, r, c):
    # Ô AD111
    # = IF(A111="",FALSE(),AND($U111>=0,$U111<=120,ISNUMBER(SEARCH("NẶNG",LUU_NIEN!$CN$129))))
    return (X.FALSE() if X.truthy(X.eq(E.v(5,r,c-29),'')) else X.AND(X.ge(E.v(5,r,21),0.0), X.le(E.v(5,r,21),120.0), X.ISNUMBER(lambda: X.SEARCH('NẶNG', E.v(15,129,92)))))

CELLS = {
    'E4': (f0, 4, 5),
    'E5': (f1, 5, 5),
    'E6': (f2, 6, 5),
    'E7': (f3, 7, 5),
    'E8': (f4, 8, 5),
    'E9': (f5, 9, 5),
    'A12': (f6, 12, 1),
    'B12': (f7, 12, 2),
    'C12': (f8, 12, 3),
    'D12': (f9, 12, 4),
    'E12': (f10, 12, 5),
    'F12': (f11, 12, 6),
    'G12': (f12, 12, 7),
    'H12': (f13, 12, 8),
    'I12': (f14, 12, 9),
    'J12': (f15, 12, 10),
    'K12': (f16, 12, 11),
    'L12': (f17, 12, 12),
    'M12': (f18, 12, 13),
    'Q12': (f19, 12, 17),
    'R12': (f20, 12, 18),
    'T12': (f21, 12, 20),
    'U12': (f22, 12, 21),
    'V12': (f23, 12, 22),
    'W12': (f24, 12, 23),
    'X12': (f25, 12, 24),
    'Y12': (f26, 12, 25),
    'Z12': (f27, 12, 26),
    'AA12': (f28, 12, 27),
    'AB12': (f29, 12, 28),
    'AC12': (f30, 12, 29),
    'AD12': (f31, 12, 30),
    'AE12': (f32, 12, 31),
    'AF12': (f33, 12, 32),
    'AG12': (f34, 12, 33),
    'AH12': (f35, 12, 34),
    'AI12': (f36, 12, 35),
    'AJ12': (f37, 12, 36),
    'AK12': (f38, 12, 37),
    'AL12': (f39, 12, 38),
    'AM12': (f40, 12, 39),
    'AN12': (f41, 12, 40),
    'AO12': (f42, 12, 41),
    'AP12': (f43, 12, 42),
    'AQ12': (f44, 12, 43),
    'AR12': (f45, 12, 44),
    'AS12': (f46, 12, 45),
    'AT12': (f47, 12, 46),
    'AU12': (f48, 12, 47),
    'AV12': (f49, 12, 48),
    'AW12': (f50, 12, 49),
    'AX12': (f51, 12, 50),
    'AY12': (f52, 12, 51),
    'AZ12': (f53, 12, 52),
    'BA12': (f54, 12, 53),
    'BB12': (f55, 12, 54),
    'A13': (f56, 13, 1),
    'B13': (f7, 13, 2),
    'C13': (f57, 13, 3),
    'D13': (f58, 13, 4),
    'E13': (f59, 13, 5),
    'F13': (f60, 13, 6),
    'G13': (f12, 13, 7),
    'H13': (f13, 13, 8),
    'I13': (f14, 13, 9),
    'J13': (f15, 13, 10),
    'K13': (f61, 13, 11),
    'L13': (f62, 13, 12),
    'M13': (f63, 13, 13),
    'Q13': (f19, 13, 17),
    'R13': (f20, 13, 18),
    'T13': (f64, 13, 20),
    'U13': (f22, 13, 21),
    'V13': (f23, 13, 22),
    'W13': (f24, 13, 23),
    'X13': (f25, 13, 24),
    'Y13': (f26, 13, 25),
    'Z13': (f27, 13, 26),
    'AA13': (f65, 13, 27),
    'AB13': (f66, 13, 28),
    'AC13': (f67, 13, 29),
    'AD13': (f68, 13, 30),
    'AE13': (f32, 13, 31),
    'AF13': (f33, 13, 32),
    'AG13': (f69, 13, 33),
    'AH13': (f70, 13, 34),
    'AI13': (f71, 13, 35),
    'AJ13': (f72, 13, 36),
    'AK13': (f73, 13, 37),
    'AL13': (f74, 13, 38),
    'AM13': (f75, 13, 39),
    'AN13': (f76, 13, 40),
    'AO13': (f77, 13, 41),
    'AP13': (f78, 13, 42),
    'AQ13': (f79, 13, 43),
    'AR13': (f80, 13, 44),
    'AS13': (f81, 13, 45),
    'AT13': (f82, 13, 46),
    'AU13': (f83, 13, 47),
    'AV13': (f84, 13, 48),
    'AW13': (f85, 13, 49),
    'AX13': (f86, 13, 50),
    'AY13': (f87, 13, 51),
    'AZ13': (f88, 13, 52),
    'BA13': (f89, 13, 53),
    'BB13': (f90, 13, 54),
    'A14': (f91, 14, 1),
    'B14': (f7, 14, 2),
    'C14': (f92, 14, 3),
    'D14': (f93, 14, 4),
    'E14': (f94, 14, 5),
    'F14': (f95, 14, 6),
    'G14': (f12, 14, 7),
    'H14': (f13, 14, 8),
    'I14': (f14, 14, 9),
    'J14': (f15, 14, 10),
    'K14': (f96, 14, 11),
    'L14': (f97, 14, 12),
    'M14': (f98, 14, 13),
    'Q14': (f19, 14, 17),
    'R14': (f20, 14, 18),
    'T14': (f99, 14, 20),
    'U14': (f22, 14, 21),
    'V14': (f23, 14, 22),
    'W14': (f24, 14, 23),
    'X14': (f25, 14, 24),
    'Y14': (f26, 14, 25),
    'Z14': (f27, 14, 26),
    'AA14': (f100, 14, 27),
    'AB14': (f101, 14, 28),
    'AC14': (f102, 14, 29),
    'AD14': (f103, 14, 30),
    'AE14': (f32, 14, 31),
    'AF14': (f33, 14, 32),
    'AG14': (f69, 14, 33),
    'AH14': (f70, 14, 34),
    'AI14': (f71, 14, 35),
    'AJ14': (f72, 14, 36),
    'AK14': (f73, 14, 37),
    'AL14': (f74, 14, 38),
    'AM14': (f75, 14, 39),
    'AN14': (f76, 14, 40),
    'AO14': (f77, 14, 41),
    'AP14': (f78, 14, 42),
    'AQ14': (f79, 14, 43),
    'AR14': (f80, 14, 44),
    'AS14': (f81, 14, 45),
    'AT14': (f82, 14, 46),
    'AU14': (f83, 14, 47),
    'AV14': (f84, 14, 48),
    'AW14': (f85, 14, 49),
    'AX14': (f86, 14, 50),
    'AY14': (f87, 14, 51),
    'AZ14': (f88, 14, 52),
    'BA14': (f89, 14, 53),
    'BB14': (f90, 14, 54),
    'A15': (f104, 15, 1),
    'B15': (f7, 15, 2),
    'C15': (f105, 15, 3),
    'D15': (f106, 15, 4),
    'E15': (f107, 15, 5),
    'F15': (f108, 15, 6),
    'G15': (f12, 15, 7),
    'H15': (f13, 15, 8),
    'I15': (f14, 15, 9),
    'J15': (f15, 15, 10),
    'K15': (f109, 15, 11),
    'L15': (f110, 15, 12),
    'M15': (f111, 15, 13),
    'Q15': (f19, 15, 17),
    'R15': (f20, 15, 18),
    'T15': (f112, 15, 20),
    'U15': (f22, 15, 21),
    'V15': (f23, 15, 22),
    'W15': (f24, 15, 23),
    'X15': (f25, 15, 24),
    'Y15': (f26, 15, 25),
    'Z15': (f27, 15, 26),
    'AA15': (f113, 15, 27),
    'AB15': (f114, 15, 28),
    'AC15': (f115, 15, 29),
    'AD15': (f116, 15, 30),
    'AE15': (f32, 15, 31),
    'AF15': (f33, 15, 32),
    'AG15': (f69, 15, 33),
    'AH15': (f70, 15, 34),
    'AI15': (f71, 15, 35),
    'AJ15': (f72, 15, 36),
    'AK15': (f73, 15, 37),
    'AL15': (f74, 15, 38),
    'AM15': (f75, 15, 39),
    'AN15': (f76, 15, 40),
    'AO15': (f77, 15, 41),
    'AP15': (f78, 15, 42),
    'AQ15': (f79, 15, 43),
    'AR15': (f80, 15, 44),
    'AS15': (f81, 15, 45),
    'AT15': (f82, 15, 46),
    'AU15': (f83, 15, 47),
    'AV15': (f84, 15, 48),
    'AW15': (f85, 15, 49),
    'AX15': (f86, 15, 50),
    'AY15': (f87, 15, 51),
    'AZ15': (f88, 15, 52),
    'BA15': (f89, 15, 53),
    'BB15': (f90, 15, 54),
    'A16': (f117, 16, 1),
    'B16': (f7, 16, 2),
    'C16': (f118, 16, 3),
    'D16': (f119, 16, 4),
    'E16': (f120, 16, 5),
    'F16': (f121, 16, 6),
    'G16': (f12, 16, 7),
    'H16': (f13, 16, 8),
    'I16': (f14, 16, 9),
    'J16': (f15, 16, 10),
    'K16': (f122, 16, 11),
    'L16': (f123, 16, 12),
    'M16': (f124, 16, 13),
    'Q16': (f19, 16, 17),
    'R16': (f20, 16, 18),
    'T16': (f125, 16, 20),
    'U16': (f22, 16, 21),
    'V16': (f23, 16, 22),
    'W16': (f24, 16, 23),
    'X16': (f25, 16, 24),
    'Y16': (f26, 16, 25),
    'Z16': (f27, 16, 26),
    'AA16': (f126, 16, 27),
    'AB16': (f127, 16, 28),
    'AC16': (f128, 16, 29),
    'AD16': (f129, 16, 30),
    'AE16': (f32, 16, 31),
    'AF16': (f33, 16, 32),
    'AG16': (f69, 16, 33),
    'AH16': (f70, 16, 34),
    'AI16': (f71, 16, 35),
    'AJ16': (f72, 16, 36),
    'AK16': (f73, 16, 37),
    'AL16': (f74, 16, 38),
    'AM16': (f75, 16, 39),
    'AN16': (f76, 16, 40),
    'AO16': (f77, 16, 41),
    'AP16': (f78, 16, 42),
    'AQ16': (f79, 16, 43),
    'AR16': (f80, 16, 44),
    'AS16': (f81, 16, 45),
    'AT16': (f82, 16, 46),
    'AU16': (f83, 16, 47),
    'AV16': (f84, 16, 48),
    'AW16': (f85, 16, 49),
    'AX16': (f86, 16, 50),
    'AY16': (f87, 16, 51),
    'AZ16': (f88, 16, 52),
    'BA16': (f89, 16, 53),
    'BB16': (f90, 16, 54),
    'A17': (f130, 17, 1),
    'B17': (f7, 17, 2),
    'C17': (f131, 17, 3),
    'D17': (f132, 17, 4),
    'E17': (f133, 17, 5),
    'F17': (f134, 17, 6),
    'G17': (f12, 17, 7),
    'H17': (f13, 17, 8),
    'I17': (f14, 17, 9),
    'J17': (f15, 17, 10),
    'K17': (f135, 17, 11),
    'L17': (f136, 17, 12),
    'M17': (f137, 17, 13),
    'Q17': (f19, 17, 17),
    'R17': (f20, 17, 18),
    'T17': (f138, 17, 20),
    'U17': (f22, 17, 21),
    'V17': (f23, 17, 22),
    'W17': (f24, 17, 23),
    'X17': (f25, 17, 24),
    'Y17': (f26, 17, 25),
    'Z17': (f27, 17, 26),
    'AA17': (f139, 17, 27),
    'AB17': (f140, 17, 28),
    'AC17': (f141, 17, 29),
    'AD17': (f142, 17, 30),
    'AE17': (f32, 17, 31),
    'AF17': (f33, 17, 32),
    'AG17': (f69, 17, 33),
    'AH17': (f70, 17, 34),
    'AI17': (f71, 17, 35),
    'AJ17': (f72, 17, 36),
    'AK17': (f73, 17, 37),
    'AL17': (f74, 17, 38),
    'AM17': (f75, 17, 39),
    'AN17': (f76, 17, 40),
    'AO17': (f77, 17, 41),
    'AP17': (f78, 17, 42),
    'AQ17': (f79, 17, 43),
    'AR17': (f80, 17, 44),
    'AS17': (f81, 17, 45),
    'AT17': (f82, 17, 46),
    'AU17': (f83, 17, 47),
    'AV17': (f84, 17, 48),
    'AW17': (f85, 17, 49),
    'AX17': (f86, 17, 50),
    'AY17': (f87, 17, 51),
    'AZ17': (f88, 17, 52),
    'BA17': (f89, 17, 53),
    'BB17': (f90, 17, 54),
    'A18': (f143, 18, 1),
    'B18': (f7, 18, 2),
    'C18': (f144, 18, 3),
    'D18': (f145, 18, 4),
    'E18': (f146, 18, 5),
    'F18': (f147, 18, 6),
    'G18': (f12, 18, 7),
    'H18': (f13, 18, 8),
    'I18': (f14, 18, 9),
    'J18': (f15, 18, 10),
    'K18': (f148, 18, 11),
    'L18': (f149, 18, 12),
    'M18': (f150, 18, 13),
    'Q18': (f19, 18, 17),
    'R18': (f20, 18, 18),
    'T18': (f151, 18, 20),
    'U18': (f22, 18, 21),
    'V18': (f23, 18, 22),
    'W18': (f24, 18, 23),
    'X18': (f25, 18, 24),
    'Y18': (f26, 18, 25),
    'Z18': (f27, 18, 26),
    'AA18': (f152, 18, 27),
    'AB18': (f153, 18, 28),
    'AC18': (f154, 18, 29),
    'AD18': (f155, 18, 30),
    'AE18': (f32, 18, 31),
    'AF18': (f33, 18, 32),
    'AG18': (f69, 18, 33),
    'AH18': (f70, 18, 34),
    'AI18': (f71, 18, 35),
    'AJ18': (f72, 18, 36),
    'AK18': (f73, 18, 37),
    'AL18': (f74, 18, 38),
    'AM18': (f75, 18, 39),
    'AN18': (f76, 18, 40),
    'AO18': (f77, 18, 41),
    'AP18': (f78, 18, 42),
    'AQ18': (f79, 18, 43),
    'AR18': (f80, 18, 44),
    'AS18': (f81, 18, 45),
    'AT18': (f82, 18, 46),
    'AU18': (f83, 18, 47),
    'AV18': (f84, 18, 48),
    'AW18': (f85, 18, 49),
    'AX18': (f86, 18, 50),
    'AY18': (f87, 18, 51),
    'AZ18': (f88, 18, 52),
    'BA18': (f89, 18, 53),
    'BB18': (f90, 18, 54),
    'A19': (f156, 19, 1),
    'B19': (f7, 19, 2),
    'C19': (f157, 19, 3),
    'D19': (f158, 19, 4),
    'E19': (f159, 19, 5),
    'F19': (f160, 19, 6),
    'G19': (f12, 19, 7),
    'H19': (f13, 19, 8),
    'I19': (f14, 19, 9),
    'J19': (f15, 19, 10),
    'K19': (f161, 19, 11),
    'L19': (f162, 19, 12),
    'M19': (f163, 19, 13),
    'Q19': (f19, 19, 17),
    'R19': (f20, 19, 18),
    'T19': (f164, 19, 20),
    'U19': (f22, 19, 21),
    'V19': (f23, 19, 22),
    'W19': (f24, 19, 23),
    'X19': (f25, 19, 24),
    'Y19': (f26, 19, 25),
    'Z19': (f27, 19, 26),
    'AA19': (f165, 19, 27),
    'AB19': (f166, 19, 28),
    'AC19': (f167, 19, 29),
    'AD19': (f168, 19, 30),
    'AE19': (f32, 19, 31),
    'AF19': (f33, 19, 32),
    'AG19': (f69, 19, 33),
    'AH19': (f70, 19, 34),
    'AI19': (f71, 19, 35),
    'AJ19': (f72, 19, 36),
    'AK19': (f73, 19, 37),
    'AL19': (f74, 19, 38),
    'AM19': (f75, 19, 39),
    'AN19': (f76, 19, 40),
    'AO19': (f77, 19, 41),
    'AP19': (f78, 19, 42),
    'AQ19': (f79, 19, 43),
    'AR19': (f80, 19, 44),
    'AS19': (f81, 19, 45),
    'AT19': (f82, 19, 46),
    'AU19': (f83, 19, 47),
    'AV19': (f84, 19, 48),
    'AW19': (f85, 19, 49),
    'AX19': (f86, 19, 50),
    'AY19': (f87, 19, 51),
    'AZ19': (f88, 19, 52),
    'BA19': (f89, 19, 53),
    'BB19': (f90, 19, 54),
    'A20': (f169, 20, 1),
    'B20': (f7, 20, 2),
    'C20': (f170, 20, 3),
    'D20': (f171, 20, 4),
    'E20': (f172, 20, 5),
    'F20': (f173, 20, 6),
    'G20': (f12, 20, 7),
    'H20': (f13, 20, 8),
    'I20': (f14, 20, 9),
    'J20': (f15, 20, 10),
    'K20': (f174, 20, 11),
    'L20': (f175, 20, 12),
    'M20': (f176, 20, 13),
    'Q20': (f19, 20, 17),
    'R20': (f20, 20, 18),
    'T20': (f177, 20, 20),
    'U20': (f22, 20, 21),
    'V20': (f23, 20, 22),
    'W20': (f24, 20, 23),
    'X20': (f25, 20, 24),
    'Y20': (f26, 20, 25),
    'Z20': (f27, 20, 26),
    'AA20': (f178, 20, 27),
    'AB20': (f179, 20, 28),
    'AC20': (f180, 20, 29),
    'AD20': (f181, 20, 30),
    'AE20': (f32, 20, 31),
    'AF20': (f33, 20, 32),
    'AG20': (f69, 20, 33),
    'AH20': (f70, 20, 34),
    'AI20': (f71, 20, 35),
    'AJ20': (f72, 20, 36),
    'AK20': (f73, 20, 37),
    'AL20': (f74, 20, 38),
    'AM20': (f75, 20, 39),
    'AN20': (f76, 20, 40),
    'AO20': (f77, 20, 41),
    'AP20': (f78, 20, 42),
    'AQ20': (f79, 20, 43),
    'AR20': (f80, 20, 44),
    'AS20': (f81, 20, 45),
    'AT20': (f82, 20, 46),
    'AU20': (f83, 20, 47),
    'AV20': (f84, 20, 48),
    'AW20': (f85, 20, 49),
    'AX20': (f86, 20, 50),
    'AY20': (f87, 20, 51),
    'AZ20': (f88, 20, 52),
    'BA20': (f89, 20, 53),
    'BB20': (f90, 20, 54),
    'A21': (f182, 21, 1),
    'B21': (f7, 21, 2),
    'C21': (f183, 21, 3),
    'D21': (f184, 21, 4),
    'E21': (f185, 21, 5),
    'F21': (f186, 21, 6),
    'G21': (f12, 21, 7),
    'H21': (f13, 21, 8),
    'I21': (f14, 21, 9),
    'J21': (f15, 21, 10),
    'K21': (f187, 21, 11),
    'L21': (f188, 21, 12),
    'M21': (f189, 21, 13),
    'Q21': (f19, 21, 17),
    'R21': (f20, 21, 18),
    'T21': (f190, 21, 20),
    'U21': (f22, 21, 21),
    'V21': (f23, 21, 22),
    'W21': (f24, 21, 23),
    'X21': (f25, 21, 24),
    'Y21': (f26, 21, 25),
    'Z21': (f27, 21, 26),
    'AA21': (f191, 21, 27),
    'AB21': (f192, 21, 28),
    'AC21': (f193, 21, 29),
    'AD21': (f194, 21, 30),
    'AE21': (f32, 21, 31),
    'AF21': (f33, 21, 32),
    'AG21': (f69, 21, 33),
    'AH21': (f70, 21, 34),
    'AI21': (f71, 21, 35),
    'AJ21': (f72, 21, 36),
    'AK21': (f73, 21, 37),
    'AL21': (f74, 21, 38),
    'AM21': (f75, 21, 39),
    'AN21': (f76, 21, 40),
    'AO21': (f77, 21, 41),
    'AP21': (f78, 21, 42),
    'AQ21': (f79, 21, 43),
    'AR21': (f80, 21, 44),
    'AS21': (f81, 21, 45),
    'AT21': (f82, 21, 46),
    'AU21': (f83, 21, 47),
    'AV21': (f84, 21, 48),
    'AW21': (f85, 21, 49),
    'AX21': (f86, 21, 50),
    'AY21': (f87, 21, 51),
    'AZ21': (f88, 21, 52),
    'BA21': (f89, 21, 53),
    'BB21': (f90, 21, 54),
    'A22': (f195, 22, 1),
    'B22': (f7, 22, 2),
    'C22': (f196, 22, 3),
    'D22': (f197, 22, 4),
    'E22': (f198, 22, 5),
    'F22': (f199, 22, 6),
    'G22': (f12, 22, 7),
    'H22': (f13, 22, 8),
    'I22': (f14, 22, 9),
    'J22': (f15, 22, 10),
    'K22': (f200, 22, 11),
    'L22': (f201, 22, 12),
    'M22': (f202, 22, 13),
    'Q22': (f19, 22, 17),
    'R22': (f20, 22, 18),
    'T22': (f203, 22, 20),
    'U22': (f22, 22, 21),
    'V22': (f23, 22, 22),
    'W22': (f24, 22, 23),
    'X22': (f25, 22, 24),
    'Y22': (f26, 22, 25),
    'Z22': (f27, 22, 26),
    'AA22': (f204, 22, 27),
    'AB22': (f205, 22, 28),
    'AC22': (f206, 22, 29),
    'AD22': (f207, 22, 30),
    'AE22': (f32, 22, 31),
    'AF22': (f33, 22, 32),
    'AG22': (f69, 22, 33),
    'AH22': (f70, 22, 34),
    'AI22': (f71, 22, 35),
    'AJ22': (f72, 22, 36),
    'AK22': (f73, 22, 37),
    'AL22': (f74, 22, 38),
    'AM22': (f75, 22, 39),
    'AN22': (f76, 22, 40),
    'AO22': (f77, 22, 41),
    'AP22': (f78, 22, 42),
    'AQ22': (f79, 22, 43),
    'AR22': (f80, 22, 44),
    'AS22': (f81, 22, 45),
    'AT22': (f82, 22, 46),
    'AU22': (f83, 22, 47),
    'AV22': (f84, 22, 48),
    'AW22': (f85, 22, 49),
    'AX22': (f86, 22, 50),
    'AY22': (f87, 22, 51),
    'AZ22': (f88, 22, 52),
    'BA22': (f89, 22, 53),
    'BB22': (f90, 22, 54),
    'A23': (f208, 23, 1),
    'B23': (f7, 23, 2),
    'C23': (f209, 23, 3),
    'D23': (f210, 23, 4),
    'E23': (f211, 23, 5),
    'F23': (f212, 23, 6),
    'G23': (f12, 23, 7),
    'H23': (f13, 23, 8),
    'I23': (f14, 23, 9),
    'J23': (f15, 23, 10),
    'K23': (f213, 23, 11),
    'L23': (f214, 23, 12),
    'M23': (f215, 23, 13),
    'Q23': (f19, 23, 17),
    'R23': (f20, 23, 18),
    'T23': (f216, 23, 20),
    'U23': (f22, 23, 21),
    'V23': (f23, 23, 22),
    'W23': (f24, 23, 23),
    'X23': (f25, 23, 24),
    'Y23': (f26, 23, 25),
    'Z23': (f27, 23, 26),
    'AA23': (f217, 23, 27),
    'AB23': (f218, 23, 28),
    'AC23': (f219, 23, 29),
    'AD23': (f220, 23, 30),
    'AE23': (f32, 23, 31),
    'AF23': (f33, 23, 32),
    'AG23': (f69, 23, 33),
    'AH23': (f70, 23, 34),
    'AI23': (f71, 23, 35),
    'AJ23': (f72, 23, 36),
    'AK23': (f73, 23, 37),
    'AL23': (f74, 23, 38),
    'AM23': (f75, 23, 39),
    'AN23': (f76, 23, 40),
    'AO23': (f77, 23, 41),
    'AP23': (f78, 23, 42),
    'AQ23': (f79, 23, 43),
    'AR23': (f80, 23, 44),
    'AS23': (f81, 23, 45),
    'AT23': (f82, 23, 46),
    'AU23': (f83, 23, 47),
    'AV23': (f84, 23, 48),
    'AW23': (f85, 23, 49),
    'AX23': (f86, 23, 50),
    'AY23': (f87, 23, 51),
    'AZ23': (f88, 23, 52),
    'BA23': (f89, 23, 53),
    'BB23': (f90, 23, 54),
    'A24': (f221, 24, 1),
    'B24': (f7, 24, 2),
    'C24': (f222, 24, 3),
    'D24': (f223, 24, 4),
    'E24': (f224, 24, 5),
    'F24': (f225, 24, 6),
    'G24': (f12, 24, 7),
    'H24': (f13, 24, 8),
    'I24': (f14, 24, 9),
    'J24': (f15, 24, 10),
    'K24': (f226, 24, 11),
    'L24': (f227, 24, 12),
    'M24': (f228, 24, 13),
    'Q24': (f19, 24, 17),
    'R24': (f20, 24, 18),
    'T24': (f229, 24, 20),
    'U24': (f22, 24, 21),
    'V24': (f23, 24, 22),
    'W24': (f24, 24, 23),
    'X24': (f25, 24, 24),
    'Y24': (f26, 24, 25),
    'Z24': (f27, 24, 26),
    'AA24': (f230, 24, 27),
    'AB24': (f231, 24, 28),
    'AC24': (f232, 24, 29),
    'AD24': (f233, 24, 30),
    'AE24': (f32, 24, 31),
    'AF24': (f33, 24, 32),
    'AG24': (f69, 24, 33),
    'AH24': (f70, 24, 34),
    'AI24': (f71, 24, 35),
    'AJ24': (f72, 24, 36),
    'AK24': (f73, 24, 37),
    'AL24': (f74, 24, 38),
    'AM24': (f75, 24, 39),
    'AN24': (f76, 24, 40),
    'AO24': (f77, 24, 41),
    'AP24': (f78, 24, 42),
    'AQ24': (f79, 24, 43),
    'AR24': (f80, 24, 44),
    'AS24': (f81, 24, 45),
    'AT24': (f82, 24, 46),
    'AU24': (f83, 24, 47),
    'AV24': (f84, 24, 48),
    'AW24': (f85, 24, 49),
    'AX24': (f86, 24, 50),
    'AY24': (f87, 24, 51),
    'AZ24': (f88, 24, 52),
    'BA24': (f89, 24, 53),
    'BB24': (f90, 24, 54),
    'A25': (f234, 25, 1),
    'B25': (f7, 25, 2),
    'C25': (f235, 25, 3),
    'D25': (f236, 25, 4),
    'E25': (f237, 25, 5),
    'F25': (f238, 25, 6),
    'G25': (f12, 25, 7),
    'H25': (f13, 25, 8),
    'I25': (f14, 25, 9),
    'J25': (f15, 25, 10),
    'K25': (f239, 25, 11),
    'L25': (f240, 25, 12),
    'M25': (f241, 25, 13),
    'Q25': (f19, 25, 17),
    'R25': (f20, 25, 18),
    'T25': (f242, 25, 20),
    'U25': (f22, 25, 21),
    'V25': (f23, 25, 22),
    'W25': (f24, 25, 23),
    'X25': (f25, 25, 24),
    'Y25': (f26, 25, 25),
    'Z25': (f27, 25, 26),
    'AA25': (f243, 25, 27),
    'AB25': (f244, 25, 28),
    'AC25': (f245, 25, 29),
    'AD25': (f246, 25, 30),
    'AE25': (f32, 25, 31),
    'AF25': (f33, 25, 32),
    'AG25': (f69, 25, 33),
    'AH25': (f70, 25, 34),
    'AI25': (f71, 25, 35),
    'AJ25': (f72, 25, 36),
    'AK25': (f73, 25, 37),
    'AL25': (f74, 25, 38),
    'AM25': (f75, 25, 39),
    'AN25': (f76, 25, 40),
    'AO25': (f77, 25, 41),
    'AP25': (f78, 25, 42),
    'AQ25': (f79, 25, 43),
    'AR25': (f80, 25, 44),
    'AS25': (f81, 25, 45),
    'AT25': (f82, 25, 46),
    'AU25': (f83, 25, 47),
    'AV25': (f84, 25, 48),
    'AW25': (f85, 25, 49),
    'AX25': (f86, 25, 50),
    'AY25': (f87, 25, 51),
    'AZ25': (f88, 25, 52),
    'BA25': (f89, 25, 53),
    'BB25': (f90, 25, 54),
    'A26': (f247, 26, 1),
    'B26': (f7, 26, 2),
    'C26': (f248, 26, 3),
    'D26': (f249, 26, 4),
    'E26': (f250, 26, 5),
    'F26': (f251, 26, 6),
    'G26': (f12, 26, 7),
    'H26': (f13, 26, 8),
    'I26': (f14, 26, 9),
    'J26': (f15, 26, 10),
    'K26': (f252, 26, 11),
    'L26': (f253, 26, 12),
    'M26': (f254, 26, 13),
    'Q26': (f19, 26, 17),
    'R26': (f20, 26, 18),
    'T26': (f255, 26, 20),
    'U26': (f22, 26, 21),
    'V26': (f23, 26, 22),
    'W26': (f24, 26, 23),
    'X26': (f25, 26, 24),
    'Y26': (f26, 26, 25),
    'Z26': (f27, 26, 26),
    'AA26': (f256, 26, 27),
    'AB26': (f257, 26, 28),
    'AC26': (f258, 26, 29),
    'AD26': (f259, 26, 30),
    'AE26': (f32, 26, 31),
    'AF26': (f33, 26, 32),
    'AG26': (f69, 26, 33),
    'AH26': (f70, 26, 34),
    'AI26': (f71, 26, 35),
    'AJ26': (f72, 26, 36),
    'AK26': (f73, 26, 37),
    'AL26': (f74, 26, 38),
    'AM26': (f75, 26, 39),
    'AN26': (f76, 26, 40),
    'AO26': (f77, 26, 41),
    'AP26': (f78, 26, 42),
    'AQ26': (f79, 26, 43),
    'AR26': (f80, 26, 44),
    'AS26': (f81, 26, 45),
    'AT26': (f82, 26, 46),
    'AU26': (f83, 26, 47),
    'AV26': (f84, 26, 48),
    'AW26': (f85, 26, 49),
    'AX26': (f86, 26, 50),
    'AY26': (f87, 26, 51),
    'AZ26': (f88, 26, 52),
    'BA26': (f89, 26, 53),
    'BB26': (f90, 26, 54),
    'A27': (f260, 27, 1),
    'B27': (f7, 27, 2),
    'C27': (f261, 27, 3),
    'D27': (f262, 27, 4),
    'E27': (f263, 27, 5),
    'F27': (f264, 27, 6),
    'G27': (f12, 27, 7),
    'H27': (f13, 27, 8),
    'I27': (f14, 27, 9),
    'J27': (f15, 27, 10),
    'K27': (f265, 27, 11),
    'L27': (f266, 27, 12),
    'M27': (f267, 27, 13),
    'Q27': (f19, 27, 17),
    'R27': (f20, 27, 18),
    'T27': (f268, 27, 20),
    'U27': (f22, 27, 21),
    'V27': (f23, 27, 22),
    'W27': (f24, 27, 23),
    'X27': (f25, 27, 24),
    'Y27': (f26, 27, 25),
    'Z27': (f27, 27, 26),
    'AA27': (f269, 27, 27),
    'AB27': (f270, 27, 28),
    'AC27': (f271, 27, 29),
    'AD27': (f272, 27, 30),
    'AE27': (f32, 27, 31),
    'AF27': (f33, 27, 32),
    'AG27': (f69, 27, 33),
    'AH27': (f70, 27, 34),
    'AI27': (f71, 27, 35),
    'AJ27': (f72, 27, 36),
    'AK27': (f73, 27, 37),
    'AL27': (f74, 27, 38),
    'AM27': (f75, 27, 39),
    'AN27': (f76, 27, 40),
    'AO27': (f77, 27, 41),
    'AP27': (f78, 27, 42),
    'AQ27': (f79, 27, 43),
    'AR27': (f80, 27, 44),
    'AS27': (f81, 27, 45),
    'AT27': (f82, 27, 46),
    'AU27': (f83, 27, 47),
    'AV27': (f84, 27, 48),
    'AW27': (f85, 27, 49),
    'AX27': (f86, 27, 50),
    'AY27': (f87, 27, 51),
    'AZ27': (f88, 27, 52),
    'BA27': (f89, 27, 53),
    'BB27': (f90, 27, 54),
    'A28': (f273, 28, 1),
    'B28': (f7, 28, 2),
    'C28': (f274, 28, 3),
    'D28': (f275, 28, 4),
    'E28': (f276, 28, 5),
    'F28': (f277, 28, 6),
    'G28': (f12, 28, 7),
    'H28': (f13, 28, 8),
    'I28': (f14, 28, 9),
    'J28': (f15, 28, 10),
    'K28': (f278, 28, 11),
    'L28': (f279, 28, 12),
    'M28': (f280, 28, 13),
    'Q28': (f19, 28, 17),
    'R28': (f20, 28, 18),
    'T28': (f281, 28, 20),
    'U28': (f22, 28, 21),
    'V28': (f23, 28, 22),
    'W28': (f24, 28, 23),
    'X28': (f25, 28, 24),
    'Y28': (f26, 28, 25),
    'Z28': (f27, 28, 26),
    'AA28': (f282, 28, 27),
    'AB28': (f283, 28, 28),
    'AC28': (f284, 28, 29),
    'AD28': (f285, 28, 30),
    'AE28': (f32, 28, 31),
    'AF28': (f33, 28, 32),
    'AG28': (f69, 28, 33),
    'AH28': (f70, 28, 34),
    'AI28': (f71, 28, 35),
    'AJ28': (f72, 28, 36),
    'AK28': (f73, 28, 37),
    'AL28': (f74, 28, 38),
    'AM28': (f75, 28, 39),
    'AN28': (f76, 28, 40),
    'AO28': (f77, 28, 41),
    'AP28': (f78, 28, 42),
    'AQ28': (f79, 28, 43),
    'AR28': (f80, 28, 44),
    'AS28': (f81, 28, 45),
    'AT28': (f82, 28, 46),
    'AU28': (f83, 28, 47),
    'AV28': (f84, 28, 48),
    'AW28': (f85, 28, 49),
    'AX28': (f86, 28, 50),
    'AY28': (f87, 28, 51),
    'AZ28': (f88, 28, 52),
    'BA28': (f89, 28, 53),
    'BB28': (f90, 28, 54),
    'A29': (f286, 29, 1),
    'B29': (f7, 29, 2),
    'C29': (f287, 29, 3),
    'D29': (f288, 29, 4),
    'E29': (f289, 29, 5),
    'F29': (f290, 29, 6),
    'G29': (f12, 29, 7),
    'H29': (f13, 29, 8),
    'I29': (f14, 29, 9),
    'J29': (f15, 29, 10),
    'K29': (f291, 29, 11),
    'L29': (f292, 29, 12),
    'M29': (f293, 29, 13),
    'Q29': (f19, 29, 17),
    'R29': (f20, 29, 18),
    'T29': (f294, 29, 20),
    'U29': (f22, 29, 21),
    'V29': (f23, 29, 22),
    'W29': (f24, 29, 23),
    'X29': (f25, 29, 24),
    'Y29': (f26, 29, 25),
    'Z29': (f27, 29, 26),
    'AA29': (f295, 29, 27),
    'AB29': (f296, 29, 28),
    'AC29': (f297, 29, 29),
    'AD29': (f298, 29, 30),
    'AE29': (f32, 29, 31),
    'AF29': (f33, 29, 32),
    'AG29': (f69, 29, 33),
    'AH29': (f70, 29, 34),
    'AI29': (f71, 29, 35),
    'AJ29': (f72, 29, 36),
    'AK29': (f73, 29, 37),
    'AL29': (f74, 29, 38),
    'AM29': (f75, 29, 39),
    'AN29': (f76, 29, 40),
    'AO29': (f77, 29, 41),
    'AP29': (f78, 29, 42),
    'AQ29': (f79, 29, 43),
    'AR29': (f80, 29, 44),
    'AS29': (f81, 29, 45),
    'AT29': (f82, 29, 46),
    'AU29': (f83, 29, 47),
    'AV29': (f84, 29, 48),
    'AW29': (f85, 29, 49),
    'AX29': (f86, 29, 50),
    'AY29': (f87, 29, 51),
    'AZ29': (f88, 29, 52),
    'BA29': (f89, 29, 53),
    'BB29': (f90, 29, 54),
    'A30': (f299, 30, 1),
    'B30': (f7, 30, 2),
    'C30': (f300, 30, 3),
    'D30': (f301, 30, 4),
    'E30': (f302, 30, 5),
    'F30': (f303, 30, 6),
    'G30': (f12, 30, 7),
    'H30': (f13, 30, 8),
    'I30': (f14, 30, 9),
    'J30': (f15, 30, 10),
    'K30': (f304, 30, 11),
    'L30': (f305, 30, 12),
    'M30': (f306, 30, 13),
    'Q30': (f19, 30, 17),
    'R30': (f20, 30, 18),
    'T30': (f307, 30, 20),
    'U30': (f22, 30, 21),
    'V30': (f23, 30, 22),
    'W30': (f24, 30, 23),
    'X30': (f25, 30, 24),
    'Y30': (f26, 30, 25),
    'Z30': (f27, 30, 26),
    'AA30': (f308, 30, 27),
    'AB30': (f309, 30, 28),
    'AC30': (f310, 30, 29),
    'AD30': (f311, 30, 30),
    'AE30': (f32, 30, 31),
    'AF30': (f33, 30, 32),
    'AG30': (f69, 30, 33),
    'AH30': (f70, 30, 34),
    'AI30': (f71, 30, 35),
    'AJ30': (f72, 30, 36),
    'AK30': (f73, 30, 37),
    'AL30': (f74, 30, 38),
    'AM30': (f75, 30, 39),
    'AN30': (f76, 30, 40),
    'AO30': (f77, 30, 41),
    'AP30': (f78, 30, 42),
    'AQ30': (f79, 30, 43),
    'AR30': (f80, 30, 44),
    'AS30': (f81, 30, 45),
    'AT30': (f82, 30, 46),
    'AU30': (f83, 30, 47),
    'AV30': (f84, 30, 48),
    'AW30': (f85, 30, 49),
    'AX30': (f86, 30, 50),
    'AY30': (f87, 30, 51),
    'AZ30': (f88, 30, 52),
    'BA30': (f89, 30, 53),
    'BB30': (f90, 30, 54),
    'A31': (f312, 31, 1),
    'B31': (f7, 31, 2),
    'C31': (f313, 31, 3),
    'D31': (f314, 31, 4),
    'E31': (f315, 31, 5),
    'F31': (f316, 31, 6),
    'G31': (f12, 31, 7),
    'H31': (f13, 31, 8),
    'I31': (f14, 31, 9),
    'J31': (f15, 31, 10),
    'K31': (f317, 31, 11),
    'L31': (f318, 31, 12),
    'M31': (f319, 31, 13),
    'Q31': (f19, 31, 17),
    'R31': (f20, 31, 18),
    'T31': (f320, 31, 20),
    'U31': (f22, 31, 21),
    'V31': (f23, 31, 22),
    'W31': (f24, 31, 23),
    'X31': (f25, 31, 24),
    'Y31': (f26, 31, 25),
    'Z31': (f27, 31, 26),
    'AA31': (f321, 31, 27),
    'AB31': (f322, 31, 28),
    'AC31': (f323, 31, 29),
    'AD31': (f324, 31, 30),
    'AE31': (f32, 31, 31),
    'AF31': (f33, 31, 32),
    'AG31': (f69, 31, 33),
    'AH31': (f70, 31, 34),
    'AI31': (f71, 31, 35),
    'AJ31': (f72, 31, 36),
    'AK31': (f73, 31, 37),
    'AL31': (f74, 31, 38),
    'AM31': (f75, 31, 39),
    'AN31': (f76, 31, 40),
    'AO31': (f77, 31, 41),
    'AP31': (f78, 31, 42),
    'AQ31': (f79, 31, 43),
    'AR31': (f80, 31, 44),
    'AS31': (f81, 31, 45),
    'AT31': (f82, 31, 46),
    'AU31': (f83, 31, 47),
    'AV31': (f84, 31, 48),
    'AW31': (f85, 31, 49),
    'AX31': (f86, 31, 50),
    'AY31': (f87, 31, 51),
    'AZ31': (f88, 31, 52),
    'BA31': (f89, 31, 53),
    'BB31': (f90, 31, 54),
    'A32': (f325, 32, 1),
    'B32': (f7, 32, 2),
    'C32': (f326, 32, 3),
    'D32': (f327, 32, 4),
    'E32': (f328, 32, 5),
    'F32': (f329, 32, 6),
    'G32': (f12, 32, 7),
    'H32': (f13, 32, 8),
    'I32': (f14, 32, 9),
    'J32': (f15, 32, 10),
    'K32': (f330, 32, 11),
    'L32': (f331, 32, 12),
    'M32': (f332, 32, 13),
    'Q32': (f19, 32, 17),
    'R32': (f20, 32, 18),
    'T32': (f333, 32, 20),
    'U32': (f22, 32, 21),
    'V32': (f23, 32, 22),
    'W32': (f24, 32, 23),
    'X32': (f25, 32, 24),
    'Y32': (f26, 32, 25),
    'Z32': (f27, 32, 26),
    'AA32': (f334, 32, 27),
    'AB32': (f335, 32, 28),
    'AC32': (f336, 32, 29),
    'AD32': (f337, 32, 30),
    'AE32': (f32, 32, 31),
    'AF32': (f33, 32, 32),
    'AG32': (f69, 32, 33),
    'AH32': (f70, 32, 34),
    'AI32': (f71, 32, 35),
    'AJ32': (f72, 32, 36),
    'AK32': (f73, 32, 37),
    'AL32': (f74, 32, 38),
    'AM32': (f75, 32, 39),
    'AN32': (f76, 32, 40),
    'AO32': (f77, 32, 41),
    'AP32': (f78, 32, 42),
    'AQ32': (f79, 32, 43),
    'AR32': (f80, 32, 44),
    'AS32': (f81, 32, 45),
    'AT32': (f82, 32, 46),
    'AU32': (f83, 32, 47),
    'AV32': (f84, 32, 48),
    'AW32': (f85, 32, 49),
    'AX32': (f86, 32, 50),
    'AY32': (f87, 32, 51),
    'AZ32': (f88, 32, 52),
    'BA32': (f89, 32, 53),
    'BB32': (f90, 32, 54),
    'A33': (f338, 33, 1),
    'B33': (f7, 33, 2),
    'C33': (f339, 33, 3),
    'D33': (f340, 33, 4),
    'E33': (f341, 33, 5),
    'F33': (f342, 33, 6),
    'G33': (f12, 33, 7),
    'H33': (f13, 33, 8),
    'I33': (f14, 33, 9),
    'J33': (f15, 33, 10),
    'K33': (f343, 33, 11),
    'L33': (f344, 33, 12),
    'M33': (f345, 33, 13),
    'Q33': (f19, 33, 17),
    'R33': (f20, 33, 18),
    'T33': (f346, 33, 20),
    'U33': (f22, 33, 21),
    'V33': (f23, 33, 22),
    'W33': (f24, 33, 23),
    'X33': (f25, 33, 24),
    'Y33': (f26, 33, 25),
    'Z33': (f27, 33, 26),
    'AA33': (f347, 33, 27),
    'AB33': (f348, 33, 28),
    'AC33': (f349, 33, 29),
    'AD33': (f350, 33, 30),
    'AE33': (f32, 33, 31),
    'AF33': (f33, 33, 32),
    'AG33': (f69, 33, 33),
    'AH33': (f70, 33, 34),
    'AI33': (f71, 33, 35),
    'AJ33': (f72, 33, 36),
    'AK33': (f73, 33, 37),
    'AL33': (f74, 33, 38),
    'AM33': (f75, 33, 39),
    'AN33': (f76, 33, 40),
    'AO33': (f77, 33, 41),
    'AP33': (f78, 33, 42),
    'AQ33': (f79, 33, 43),
    'AR33': (f80, 33, 44),
    'AS33': (f81, 33, 45),
    'AT33': (f82, 33, 46),
    'AU33': (f83, 33, 47),
    'AV33': (f84, 33, 48),
    'AW33': (f85, 33, 49),
    'AX33': (f86, 33, 50),
    'AY33': (f87, 33, 51),
    'AZ33': (f88, 33, 52),
    'BA33': (f89, 33, 53),
    'BB33': (f90, 33, 54),
    'A34': (f351, 34, 1),
    'B34': (f7, 34, 2),
    'C34': (f352, 34, 3),
    'D34': (f353, 34, 4),
    'E34': (f354, 34, 5),
    'F34': (f355, 34, 6),
    'G34': (f12, 34, 7),
    'H34': (f13, 34, 8),
    'I34': (f14, 34, 9),
    'J34': (f15, 34, 10),
    'K34': (f356, 34, 11),
    'L34': (f357, 34, 12),
    'M34': (f358, 34, 13),
    'Q34': (f19, 34, 17),
    'R34': (f20, 34, 18),
    'T34': (f359, 34, 20),
    'U34': (f22, 34, 21),
    'V34': (f23, 34, 22),
    'W34': (f24, 34, 23),
    'X34': (f25, 34, 24),
    'Y34': (f26, 34, 25),
    'Z34': (f27, 34, 26),
    'AA34': (f360, 34, 27),
    'AB34': (f361, 34, 28),
    'AC34': (f362, 34, 29),
    'AD34': (f363, 34, 30),
    'AE34': (f32, 34, 31),
    'AF34': (f33, 34, 32),
    'AG34': (f69, 34, 33),
    'AH34': (f70, 34, 34),
    'AI34': (f71, 34, 35),
    'AJ34': (f72, 34, 36),
    'AK34': (f73, 34, 37),
    'AL34': (f74, 34, 38),
    'AM34': (f75, 34, 39),
    'AN34': (f76, 34, 40),
    'AO34': (f77, 34, 41),
    'AP34': (f78, 34, 42),
    'AQ34': (f79, 34, 43),
    'AR34': (f80, 34, 44),
    'AS34': (f81, 34, 45),
    'AT34': (f82, 34, 46),
    'AU34': (f83, 34, 47),
    'AV34': (f84, 34, 48),
    'AW34': (f85, 34, 49),
    'AX34': (f86, 34, 50),
    'AY34': (f87, 34, 51),
    'AZ34': (f88, 34, 52),
    'BA34': (f89, 34, 53),
    'BB34': (f90, 34, 54),
    'A35': (f364, 35, 1),
    'B35': (f7, 35, 2),
    'C35': (f365, 35, 3),
    'D35': (f366, 35, 4),
    'E35': (f367, 35, 5),
    'F35': (f368, 35, 6),
    'G35': (f12, 35, 7),
    'H35': (f13, 35, 8),
    'I35': (f14, 35, 9),
    'J35': (f15, 35, 10),
    'K35': (f369, 35, 11),
    'L35': (f370, 35, 12),
    'M35': (f371, 35, 13),
    'Q35': (f19, 35, 17),
    'R35': (f20, 35, 18),
    'T35': (f372, 35, 20),
    'U35': (f22, 35, 21),
    'V35': (f23, 35, 22),
    'W35': (f24, 35, 23),
    'X35': (f25, 35, 24),
    'Y35': (f26, 35, 25),
    'Z35': (f27, 35, 26),
    'AA35': (f373, 35, 27),
    'AB35': (f374, 35, 28),
    'AC35': (f375, 35, 29),
    'AD35': (f376, 35, 30),
    'AE35': (f32, 35, 31),
    'AF35': (f33, 35, 32),
    'AG35': (f69, 35, 33),
    'AH35': (f70, 35, 34),
    'AI35': (f71, 35, 35),
    'AJ35': (f72, 35, 36),
    'AK35': (f73, 35, 37),
    'AL35': (f74, 35, 38),
    'AM35': (f75, 35, 39),
    'AN35': (f76, 35, 40),
    'AO35': (f77, 35, 41),
    'AP35': (f78, 35, 42),
    'AQ35': (f79, 35, 43),
    'AR35': (f80, 35, 44),
    'AS35': (f81, 35, 45),
    'AT35': (f82, 35, 46),
    'AU35': (f83, 35, 47),
    'AV35': (f84, 35, 48),
    'AW35': (f85, 35, 49),
    'AX35': (f86, 35, 50),
    'AY35': (f87, 35, 51),
    'AZ35': (f88, 35, 52),
    'BA35': (f89, 35, 53),
    'BB35': (f90, 35, 54),
    'A36': (f377, 36, 1),
    'B36': (f7, 36, 2),
    'C36': (f378, 36, 3),
    'D36': (f379, 36, 4),
    'E36': (f380, 36, 5),
    'F36': (f381, 36, 6),
    'G36': (f12, 36, 7),
    'H36': (f13, 36, 8),
    'I36': (f14, 36, 9),
    'J36': (f15, 36, 10),
    'K36': (f382, 36, 11),
    'L36': (f383, 36, 12),
    'M36': (f384, 36, 13),
    'Q36': (f19, 36, 17),
    'R36': (f20, 36, 18),
    'T36': (f385, 36, 20),
    'U36': (f22, 36, 21),
    'V36': (f23, 36, 22),
    'W36': (f24, 36, 23),
    'X36': (f25, 36, 24),
    'Y36': (f26, 36, 25),
    'Z36': (f27, 36, 26),
    'AA36': (f386, 36, 27),
    'AB36': (f387, 36, 28),
    'AC36': (f388, 36, 29),
    'AD36': (f389, 36, 30),
    'AE36': (f32, 36, 31),
    'AF36': (f33, 36, 32),
    'AG36': (f69, 36, 33),
    'AH36': (f70, 36, 34),
    'AI36': (f71, 36, 35),
    'AJ36': (f72, 36, 36),
    'AK36': (f73, 36, 37),
    'AL36': (f74, 36, 38),
    'AM36': (f75, 36, 39),
    'AN36': (f76, 36, 40),
    'AO36': (f77, 36, 41),
    'AP36': (f78, 36, 42),
    'AQ36': (f79, 36, 43),
    'AR36': (f80, 36, 44),
    'AS36': (f81, 36, 45),
    'AT36': (f82, 36, 46),
    'AU36': (f83, 36, 47),
    'AV36': (f84, 36, 48),
    'AW36': (f85, 36, 49),
    'AX36': (f86, 36, 50),
    'AY36': (f87, 36, 51),
    'AZ36': (f88, 36, 52),
    'BA36': (f89, 36, 53),
    'BB36': (f90, 36, 54),
    'A37': (f390, 37, 1),
    'B37': (f7, 37, 2),
    'C37': (f391, 37, 3),
    'D37': (f392, 37, 4),
    'E37': (f393, 37, 5),
    'F37': (f394, 37, 6),
    'G37': (f12, 37, 7),
    'H37': (f13, 37, 8),
    'I37': (f14, 37, 9),
    'J37': (f15, 37, 10),
    'K37': (f395, 37, 11),
    'L37': (f396, 37, 12),
    'M37': (f397, 37, 13),
    'Q37': (f19, 37, 17),
    'R37': (f20, 37, 18),
    'T37': (f398, 37, 20),
    'U37': (f22, 37, 21),
    'V37': (f23, 37, 22),
    'W37': (f24, 37, 23),
    'X37': (f25, 37, 24),
    'Y37': (f26, 37, 25),
    'Z37': (f27, 37, 26),
    'AA37': (f399, 37, 27),
    'AB37': (f400, 37, 28),
    'AC37': (f401, 37, 29),
    'AD37': (f402, 37, 30),
    'AE37': (f32, 37, 31),
    'AF37': (f33, 37, 32),
    'AG37': (f69, 37, 33),
    'AH37': (f70, 37, 34),
    'AI37': (f71, 37, 35),
    'AJ37': (f72, 37, 36),
    'AK37': (f73, 37, 37),
    'AL37': (f74, 37, 38),
    'AM37': (f75, 37, 39),
    'AN37': (f76, 37, 40),
    'AO37': (f77, 37, 41),
    'AP37': (f78, 37, 42),
    'AQ37': (f79, 37, 43),
    'AR37': (f80, 37, 44),
    'AS37': (f81, 37, 45),
    'AT37': (f82, 37, 46),
    'AU37': (f83, 37, 47),
    'AV37': (f84, 37, 48),
    'AW37': (f85, 37, 49),
    'AX37': (f86, 37, 50),
    'AY37': (f87, 37, 51),
    'AZ37': (f88, 37, 52),
    'BA37': (f89, 37, 53),
    'BB37': (f90, 37, 54),
    'A38': (f403, 38, 1),
    'B38': (f7, 38, 2),
    'C38': (f404, 38, 3),
    'D38': (f405, 38, 4),
    'E38': (f406, 38, 5),
    'F38': (f407, 38, 6),
    'G38': (f12, 38, 7),
    'H38': (f13, 38, 8),
    'I38': (f14, 38, 9),
    'J38': (f15, 38, 10),
    'K38': (f408, 38, 11),
    'L38': (f409, 38, 12),
    'M38': (f410, 38, 13),
    'Q38': (f19, 38, 17),
    'R38': (f20, 38, 18),
    'T38': (f411, 38, 20),
    'U38': (f22, 38, 21),
    'V38': (f23, 38, 22),
    'W38': (f24, 38, 23),
    'X38': (f25, 38, 24),
    'Y38': (f26, 38, 25),
    'Z38': (f27, 38, 26),
    'AA38': (f412, 38, 27),
    'AB38': (f413, 38, 28),
    'AC38': (f414, 38, 29),
    'AD38': (f415, 38, 30),
    'AE38': (f32, 38, 31),
    'AF38': (f33, 38, 32),
    'AG38': (f69, 38, 33),
    'AH38': (f70, 38, 34),
    'AI38': (f71, 38, 35),
    'AJ38': (f72, 38, 36),
    'AK38': (f73, 38, 37),
    'AL38': (f74, 38, 38),
    'AM38': (f75, 38, 39),
    'AN38': (f76, 38, 40),
    'AO38': (f77, 38, 41),
    'AP38': (f78, 38, 42),
    'AQ38': (f79, 38, 43),
    'AR38': (f80, 38, 44),
    'AS38': (f81, 38, 45),
    'AT38': (f82, 38, 46),
    'AU38': (f83, 38, 47),
    'AV38': (f84, 38, 48),
    'AW38': (f85, 38, 49),
    'AX38': (f86, 38, 50),
    'AY38': (f87, 38, 51),
    'AZ38': (f88, 38, 52),
    'BA38': (f89, 38, 53),
    'BB38': (f90, 38, 54),
    'A39': (f416, 39, 1),
    'B39': (f7, 39, 2),
    'C39': (f417, 39, 3),
    'D39': (f418, 39, 4),
    'E39': (f419, 39, 5),
    'F39': (f420, 39, 6),
    'G39': (f12, 39, 7),
    'H39': (f13, 39, 8),
    'I39': (f14, 39, 9),
    'J39': (f15, 39, 10),
    'K39': (f421, 39, 11),
    'L39': (f422, 39, 12),
    'M39': (f423, 39, 13),
    'Q39': (f19, 39, 17),
    'R39': (f20, 39, 18),
    'T39': (f424, 39, 20),
    'U39': (f22, 39, 21),
    'V39': (f23, 39, 22),
    'W39': (f24, 39, 23),
    'X39': (f25, 39, 24),
    'Y39': (f26, 39, 25),
    'Z39': (f27, 39, 26),
    'AA39': (f425, 39, 27),
    'AB39': (f426, 39, 28),
    'AC39': (f427, 39, 29),
    'AD39': (f428, 39, 30),
    'AE39': (f32, 39, 31),
    'AF39': (f33, 39, 32),
    'AG39': (f69, 39, 33),
    'AH39': (f70, 39, 34),
    'AI39': (f71, 39, 35),
    'AJ39': (f72, 39, 36),
    'AK39': (f73, 39, 37),
    'AL39': (f74, 39, 38),
    'AM39': (f75, 39, 39),
    'AN39': (f76, 39, 40),
    'AO39': (f77, 39, 41),
    'AP39': (f78, 39, 42),
    'AQ39': (f79, 39, 43),
    'AR39': (f80, 39, 44),
    'AS39': (f81, 39, 45),
    'AT39': (f82, 39, 46),
    'AU39': (f83, 39, 47),
    'AV39': (f84, 39, 48),
    'AW39': (f85, 39, 49),
    'AX39': (f86, 39, 50),
    'AY39': (f87, 39, 51),
    'AZ39': (f88, 39, 52),
    'BA39': (f89, 39, 53),
    'BB39': (f90, 39, 54),
    'A40': (f429, 40, 1),
    'B40': (f7, 40, 2),
    'C40': (f430, 40, 3),
    'D40': (f431, 40, 4),
    'E40': (f432, 40, 5),
    'F40': (f433, 40, 6),
    'G40': (f12, 40, 7),
    'H40': (f13, 40, 8),
    'I40': (f14, 40, 9),
    'J40': (f15, 40, 10),
    'K40': (f434, 40, 11),
    'L40': (f435, 40, 12),
    'M40': (f436, 40, 13),
    'Q40': (f19, 40, 17),
    'R40': (f20, 40, 18),
    'T40': (f437, 40, 20),
    'U40': (f22, 40, 21),
    'V40': (f23, 40, 22),
    'W40': (f24, 40, 23),
    'X40': (f25, 40, 24),
    'Y40': (f26, 40, 25),
    'Z40': (f27, 40, 26),
    'AA40': (f438, 40, 27),
    'AB40': (f439, 40, 28),
    'AC40': (f440, 40, 29),
    'AD40': (f441, 40, 30),
    'AE40': (f32, 40, 31),
    'AF40': (f33, 40, 32),
    'AG40': (f69, 40, 33),
    'AH40': (f70, 40, 34),
    'AI40': (f71, 40, 35),
    'AJ40': (f72, 40, 36),
    'AK40': (f73, 40, 37),
    'AL40': (f74, 40, 38),
    'AM40': (f75, 40, 39),
    'AN40': (f76, 40, 40),
    'AO40': (f77, 40, 41),
    'AP40': (f78, 40, 42),
    'AQ40': (f79, 40, 43),
    'AR40': (f80, 40, 44),
    'AS40': (f81, 40, 45),
    'AT40': (f82, 40, 46),
    'AU40': (f83, 40, 47),
    'AV40': (f84, 40, 48),
    'AW40': (f85, 40, 49),
    'AX40': (f86, 40, 50),
    'AY40': (f87, 40, 51),
    'AZ40': (f88, 40, 52),
    'BA40': (f89, 40, 53),
    'BB40': (f90, 40, 54),
    'A41': (f442, 41, 1),
    'B41': (f7, 41, 2),
    'C41': (f443, 41, 3),
    'D41': (f444, 41, 4),
    'E41': (f445, 41, 5),
    'F41': (f446, 41, 6),
    'G41': (f12, 41, 7),
    'H41': (f13, 41, 8),
    'I41': (f14, 41, 9),
    'J41': (f15, 41, 10),
    'K41': (f447, 41, 11),
    'L41': (f448, 41, 12),
    'M41': (f449, 41, 13),
    'Q41': (f19, 41, 17),
    'R41': (f20, 41, 18),
    'T41': (f450, 41, 20),
    'U41': (f22, 41, 21),
    'V41': (f23, 41, 22),
    'W41': (f24, 41, 23),
    'X41': (f25, 41, 24),
    'Y41': (f26, 41, 25),
    'Z41': (f27, 41, 26),
    'AA41': (f451, 41, 27),
    'AB41': (f452, 41, 28),
    'AC41': (f453, 41, 29),
    'AD41': (f454, 41, 30),
    'AE41': (f32, 41, 31),
    'AF41': (f33, 41, 32),
    'AG41': (f69, 41, 33),
    'AH41': (f70, 41, 34),
    'AI41': (f71, 41, 35),
    'AJ41': (f72, 41, 36),
    'AK41': (f73, 41, 37),
    'AL41': (f74, 41, 38),
    'AM41': (f75, 41, 39),
    'AN41': (f76, 41, 40),
    'AO41': (f77, 41, 41),
    'AP41': (f78, 41, 42),
    'AQ41': (f79, 41, 43),
    'AR41': (f80, 41, 44),
    'AS41': (f81, 41, 45),
    'AT41': (f82, 41, 46),
    'AU41': (f83, 41, 47),
    'AV41': (f84, 41, 48),
    'AW41': (f85, 41, 49),
    'AX41': (f86, 41, 50),
    'AY41': (f87, 41, 51),
    'AZ41': (f88, 41, 52),
    'BA41': (f89, 41, 53),
    'BB41': (f90, 41, 54),
    'A42': (f455, 42, 1),
    'B42': (f7, 42, 2),
    'C42': (f456, 42, 3),
    'D42': (f457, 42, 4),
    'E42': (f458, 42, 5),
    'F42': (f459, 42, 6),
    'G42': (f12, 42, 7),
    'H42': (f13, 42, 8),
    'I42': (f14, 42, 9),
    'J42': (f15, 42, 10),
    'K42': (f460, 42, 11),
    'L42': (f461, 42, 12),
    'M42': (f462, 42, 13),
    'Q42': (f19, 42, 17),
    'R42': (f20, 42, 18),
    'T42': (f463, 42, 20),
    'U42': (f22, 42, 21),
    'V42': (f23, 42, 22),
    'W42': (f24, 42, 23),
    'X42': (f25, 42, 24),
    'Y42': (f26, 42, 25),
    'Z42': (f27, 42, 26),
    'AA42': (f464, 42, 27),
    'AB42': (f465, 42, 28),
    'AC42': (f466, 42, 29),
    'AD42': (f467, 42, 30),
    'AE42': (f32, 42, 31),
    'AF42': (f33, 42, 32),
    'AG42': (f69, 42, 33),
    'AH42': (f70, 42, 34),
    'AI42': (f71, 42, 35),
    'AJ42': (f72, 42, 36),
    'AK42': (f73, 42, 37),
    'AL42': (f74, 42, 38),
    'AM42': (f75, 42, 39),
    'AN42': (f76, 42, 40),
    'AO42': (f77, 42, 41),
    'AP42': (f78, 42, 42),
    'AQ42': (f79, 42, 43),
    'AR42': (f80, 42, 44),
    'AS42': (f81, 42, 45),
    'AT42': (f82, 42, 46),
    'AU42': (f83, 42, 47),
    'AV42': (f84, 42, 48),
    'AW42': (f85, 42, 49),
    'AX42': (f86, 42, 50),
    'AY42': (f87, 42, 51),
    'AZ42': (f88, 42, 52),
    'BA42': (f89, 42, 53),
    'BB42': (f90, 42, 54),
    'A43': (f468, 43, 1),
    'B43': (f7, 43, 2),
    'C43': (f469, 43, 3),
    'D43': (f470, 43, 4),
    'E43': (f471, 43, 5),
    'F43': (f472, 43, 6),
    'G43': (f12, 43, 7),
    'H43': (f13, 43, 8),
    'I43': (f14, 43, 9),
    'J43': (f15, 43, 10),
    'K43': (f473, 43, 11),
    'L43': (f474, 43, 12),
    'M43': (f475, 43, 13),
    'Q43': (f19, 43, 17),
    'R43': (f20, 43, 18),
    'T43': (f476, 43, 20),
    'U43': (f22, 43, 21),
    'V43': (f23, 43, 22),
    'W43': (f24, 43, 23),
    'X43': (f25, 43, 24),
    'Y43': (f26, 43, 25),
    'Z43': (f27, 43, 26),
    'AA43': (f477, 43, 27),
    'AB43': (f478, 43, 28),
    'AC43': (f479, 43, 29),
    'AD43': (f480, 43, 30),
    'AE43': (f32, 43, 31),
    'AF43': (f33, 43, 32),
    'AG43': (f69, 43, 33),
    'AH43': (f70, 43, 34),
    'AI43': (f71, 43, 35),
    'AJ43': (f72, 43, 36),
    'AK43': (f73, 43, 37),
    'AL43': (f74, 43, 38),
    'AM43': (f75, 43, 39),
    'AN43': (f76, 43, 40),
    'AO43': (f77, 43, 41),
    'AP43': (f78, 43, 42),
    'AQ43': (f79, 43, 43),
    'AR43': (f80, 43, 44),
    'AS43': (f81, 43, 45),
    'AT43': (f82, 43, 46),
    'AU43': (f83, 43, 47),
    'AV43': (f84, 43, 48),
    'AW43': (f85, 43, 49),
    'AX43': (f86, 43, 50),
    'AY43': (f87, 43, 51),
    'AZ43': (f88, 43, 52),
    'BA43': (f89, 43, 53),
    'BB43': (f90, 43, 54),
    'A44': (f481, 44, 1),
    'B44': (f7, 44, 2),
    'C44': (f482, 44, 3),
    'D44': (f483, 44, 4),
    'E44': (f484, 44, 5),
    'F44': (f485, 44, 6),
    'G44': (f12, 44, 7),
    'H44': (f13, 44, 8),
    'I44': (f14, 44, 9),
    'J44': (f15, 44, 10),
    'K44': (f486, 44, 11),
    'L44': (f487, 44, 12),
    'M44': (f488, 44, 13),
    'Q44': (f19, 44, 17),
    'R44': (f20, 44, 18),
    'T44': (f489, 44, 20),
    'U44': (f22, 44, 21),
    'V44': (f23, 44, 22),
    'W44': (f24, 44, 23),
    'X44': (f25, 44, 24),
    'Y44': (f26, 44, 25),
    'Z44': (f27, 44, 26),
    'AA44': (f490, 44, 27),
    'AB44': (f491, 44, 28),
    'AC44': (f492, 44, 29),
    'AD44': (f493, 44, 30),
    'AE44': (f32, 44, 31),
    'AF44': (f33, 44, 32),
    'AG44': (f69, 44, 33),
    'AH44': (f70, 44, 34),
    'AI44': (f71, 44, 35),
    'AJ44': (f72, 44, 36),
    'AK44': (f73, 44, 37),
    'AL44': (f74, 44, 38),
    'AM44': (f75, 44, 39),
    'AN44': (f76, 44, 40),
    'AO44': (f77, 44, 41),
    'AP44': (f78, 44, 42),
    'AQ44': (f79, 44, 43),
    'AR44': (f80, 44, 44),
    'AS44': (f81, 44, 45),
    'AT44': (f82, 44, 46),
    'AU44': (f83, 44, 47),
    'AV44': (f84, 44, 48),
    'AW44': (f85, 44, 49),
    'AX44': (f86, 44, 50),
    'AY44': (f87, 44, 51),
    'AZ44': (f88, 44, 52),
    'BA44': (f89, 44, 53),
    'BB44': (f90, 44, 54),
    'A45': (f494, 45, 1),
    'B45': (f7, 45, 2),
    'C45': (f495, 45, 3),
    'D45': (f496, 45, 4),
    'E45': (f497, 45, 5),
    'F45': (f498, 45, 6),
    'G45': (f12, 45, 7),
    'H45': (f13, 45, 8),
    'I45': (f14, 45, 9),
    'J45': (f15, 45, 10),
    'K45': (f499, 45, 11),
    'L45': (f500, 45, 12),
    'M45': (f501, 45, 13),
    'Q45': (f19, 45, 17),
    'R45': (f20, 45, 18),
    'T45': (f502, 45, 20),
    'U45': (f22, 45, 21),
    'V45': (f23, 45, 22),
    'W45': (f24, 45, 23),
    'X45': (f25, 45, 24),
    'Y45': (f26, 45, 25),
    'Z45': (f27, 45, 26),
    'AA45': (f503, 45, 27),
    'AB45': (f504, 45, 28),
    'AC45': (f505, 45, 29),
    'AD45': (f506, 45, 30),
    'AE45': (f32, 45, 31),
    'AF45': (f33, 45, 32),
    'AG45': (f69, 45, 33),
    'AH45': (f70, 45, 34),
    'AI45': (f71, 45, 35),
    'AJ45': (f72, 45, 36),
    'AK45': (f73, 45, 37),
    'AL45': (f74, 45, 38),
    'AM45': (f75, 45, 39),
    'AN45': (f76, 45, 40),
    'AO45': (f77, 45, 41),
    'AP45': (f78, 45, 42),
    'AQ45': (f79, 45, 43),
    'AR45': (f80, 45, 44),
    'AS45': (f81, 45, 45),
    'AT45': (f82, 45, 46),
    'AU45': (f83, 45, 47),
    'AV45': (f84, 45, 48),
    'AW45': (f85, 45, 49),
    'AX45': (f86, 45, 50),
    'AY45': (f87, 45, 51),
    'AZ45': (f88, 45, 52),
    'BA45': (f89, 45, 53),
    'BB45': (f90, 45, 54),
    'A46': (f507, 46, 1),
    'B46': (f7, 46, 2),
    'C46': (f508, 46, 3),
    'D46': (f509, 46, 4),
    'E46': (f510, 46, 5),
    'F46': (f511, 46, 6),
    'G46': (f12, 46, 7),
    'H46': (f13, 46, 8),
    'I46': (f14, 46, 9),
    'J46': (f15, 46, 10),
    'K46': (f512, 46, 11),
    'L46': (f513, 46, 12),
    'M46': (f514, 46, 13),
    'Q46': (f19, 46, 17),
    'R46': (f20, 46, 18),
    'T46': (f515, 46, 20),
    'U46': (f22, 46, 21),
    'V46': (f23, 46, 22),
    'W46': (f24, 46, 23),
    'X46': (f25, 46, 24),
    'Y46': (f26, 46, 25),
    'Z46': (f27, 46, 26),
    'AA46': (f516, 46, 27),
    'AB46': (f517, 46, 28),
    'AC46': (f518, 46, 29),
    'AD46': (f519, 46, 30),
    'AE46': (f32, 46, 31),
    'AF46': (f33, 46, 32),
    'AG46': (f69, 46, 33),
    'AH46': (f70, 46, 34),
    'AI46': (f71, 46, 35),
    'AJ46': (f72, 46, 36),
    'AK46': (f73, 46, 37),
    'AL46': (f74, 46, 38),
    'AM46': (f75, 46, 39),
    'AN46': (f76, 46, 40),
    'AO46': (f77, 46, 41),
    'AP46': (f78, 46, 42),
    'AQ46': (f79, 46, 43),
    'AR46': (f80, 46, 44),
    'AS46': (f81, 46, 45),
    'AT46': (f82, 46, 46),
    'AU46': (f83, 46, 47),
    'AV46': (f84, 46, 48),
    'AW46': (f85, 46, 49),
    'AX46': (f86, 46, 50),
    'AY46': (f87, 46, 51),
    'AZ46': (f88, 46, 52),
    'BA46': (f89, 46, 53),
    'BB46': (f90, 46, 54),
    'A47': (f520, 47, 1),
    'B47': (f7, 47, 2),
    'C47': (f521, 47, 3),
    'D47': (f522, 47, 4),
    'E47': (f523, 47, 5),
    'F47': (f524, 47, 6),
    'G47': (f12, 47, 7),
    'H47': (f13, 47, 8),
    'I47': (f14, 47, 9),
    'J47': (f15, 47, 10),
    'K47': (f525, 47, 11),
    'L47': (f526, 47, 12),
    'M47': (f527, 47, 13),
    'Q47': (f19, 47, 17),
    'R47': (f20, 47, 18),
    'T47': (f528, 47, 20),
    'U47': (f22, 47, 21),
    'V47': (f23, 47, 22),
    'W47': (f24, 47, 23),
    'X47': (f25, 47, 24),
    'Y47': (f26, 47, 25),
    'Z47': (f27, 47, 26),
    'AA47': (f529, 47, 27),
    'AB47': (f530, 47, 28),
    'AC47': (f531, 47, 29),
    'AD47': (f532, 47, 30),
    'AE47': (f32, 47, 31),
    'AF47': (f33, 47, 32),
    'AG47': (f69, 47, 33),
    'AH47': (f70, 47, 34),
    'AI47': (f71, 47, 35),
    'AJ47': (f72, 47, 36),
    'AK47': (f73, 47, 37),
    'AL47': (f74, 47, 38),
    'AM47': (f75, 47, 39),
    'AN47': (f76, 47, 40),
    'AO47': (f77, 47, 41),
    'AP47': (f78, 47, 42),
    'AQ47': (f79, 47, 43),
    'AR47': (f80, 47, 44),
    'AS47': (f81, 47, 45),
    'AT47': (f82, 47, 46),
    'AU47': (f83, 47, 47),
    'AV47': (f84, 47, 48),
    'AW47': (f85, 47, 49),
    'AX47': (f86, 47, 50),
    'AY47': (f87, 47, 51),
    'AZ47': (f88, 47, 52),
    'BA47': (f89, 47, 53),
    'BB47': (f90, 47, 54),
    'A48': (f533, 48, 1),
    'B48': (f7, 48, 2),
    'C48': (f534, 48, 3),
    'D48': (f535, 48, 4),
    'E48': (f536, 48, 5),
    'F48': (f537, 48, 6),
    'G48': (f12, 48, 7),
    'H48': (f13, 48, 8),
    'I48': (f14, 48, 9),
    'J48': (f15, 48, 10),
    'K48': (f538, 48, 11),
    'L48': (f539, 48, 12),
    'M48': (f540, 48, 13),
    'Q48': (f19, 48, 17),
    'R48': (f20, 48, 18),
    'T48': (f541, 48, 20),
    'U48': (f22, 48, 21),
    'V48': (f23, 48, 22),
    'W48': (f24, 48, 23),
    'X48': (f25, 48, 24),
    'Y48': (f26, 48, 25),
    'Z48': (f27, 48, 26),
    'AA48': (f542, 48, 27),
    'AB48': (f543, 48, 28),
    'AC48': (f544, 48, 29),
    'AD48': (f545, 48, 30),
    'AE48': (f32, 48, 31),
    'AF48': (f33, 48, 32),
    'AG48': (f69, 48, 33),
    'AH48': (f70, 48, 34),
    'AI48': (f71, 48, 35),
    'AJ48': (f72, 48, 36),
    'AK48': (f73, 48, 37),
    'AL48': (f74, 48, 38),
    'AM48': (f75, 48, 39),
    'AN48': (f76, 48, 40),
    'AO48': (f77, 48, 41),
    'AP48': (f78, 48, 42),
    'AQ48': (f79, 48, 43),
    'AR48': (f80, 48, 44),
    'AS48': (f81, 48, 45),
    'AT48': (f82, 48, 46),
    'AU48': (f83, 48, 47),
    'AV48': (f84, 48, 48),
    'AW48': (f85, 48, 49),
    'AX48': (f86, 48, 50),
    'AY48': (f87, 48, 51),
    'AZ48': (f88, 48, 52),
    'BA48': (f89, 48, 53),
    'BB48': (f90, 48, 54),
    'A49': (f546, 49, 1),
    'B49': (f7, 49, 2),
    'C49': (f547, 49, 3),
    'D49': (f548, 49, 4),
    'E49': (f549, 49, 5),
    'F49': (f550, 49, 6),
    'G49': (f12, 49, 7),
    'H49': (f13, 49, 8),
    'I49': (f14, 49, 9),
    'J49': (f15, 49, 10),
    'K49': (f551, 49, 11),
    'L49': (f552, 49, 12),
    'M49': (f553, 49, 13),
    'Q49': (f19, 49, 17),
    'R49': (f20, 49, 18),
    'T49': (f554, 49, 20),
    'U49': (f22, 49, 21),
    'V49': (f23, 49, 22),
    'W49': (f24, 49, 23),
    'X49': (f25, 49, 24),
    'Y49': (f26, 49, 25),
    'Z49': (f27, 49, 26),
    'AA49': (f555, 49, 27),
    'AB49': (f556, 49, 28),
    'AC49': (f557, 49, 29),
    'AD49': (f558, 49, 30),
    'AE49': (f32, 49, 31),
    'AF49': (f33, 49, 32),
    'AG49': (f69, 49, 33),
    'AH49': (f70, 49, 34),
    'AI49': (f71, 49, 35),
    'AJ49': (f72, 49, 36),
    'AK49': (f73, 49, 37),
    'AL49': (f74, 49, 38),
    'AM49': (f75, 49, 39),
    'AN49': (f76, 49, 40),
    'AO49': (f77, 49, 41),
    'AP49': (f78, 49, 42),
    'AQ49': (f79, 49, 43),
    'AR49': (f80, 49, 44),
    'AS49': (f81, 49, 45),
    'AT49': (f82, 49, 46),
    'AU49': (f83, 49, 47),
    'AV49': (f84, 49, 48),
    'AW49': (f85, 49, 49),
    'AX49': (f86, 49, 50),
    'AY49': (f87, 49, 51),
    'AZ49': (f88, 49, 52),
    'BA49': (f89, 49, 53),
    'BB49': (f90, 49, 54),
    'A50': (f559, 50, 1),
    'B50': (f7, 50, 2),
    'C50': (f560, 50, 3),
    'D50': (f561, 50, 4),
    'E50': (f562, 50, 5),
    'F50': (f563, 50, 6),
    'G50': (f12, 50, 7),
    'H50': (f13, 50, 8),
    'I50': (f14, 50, 9),
    'J50': (f15, 50, 10),
    'K50': (f564, 50, 11),
    'L50': (f565, 50, 12),
    'M50': (f566, 50, 13),
    'Q50': (f19, 50, 17),
    'R50': (f20, 50, 18),
    'T50': (f567, 50, 20),
    'U50': (f22, 50, 21),
    'V50': (f23, 50, 22),
    'W50': (f24, 50, 23),
    'X50': (f25, 50, 24),
    'Y50': (f26, 50, 25),
    'Z50': (f27, 50, 26),
    'AA50': (f568, 50, 27),
    'AB50': (f569, 50, 28),
    'AC50': (f570, 50, 29),
    'AD50': (f571, 50, 30),
    'AE50': (f32, 50, 31),
    'AF50': (f33, 50, 32),
    'AG50': (f69, 50, 33),
    'AH50': (f70, 50, 34),
    'AI50': (f71, 50, 35),
    'AJ50': (f72, 50, 36),
    'AK50': (f73, 50, 37),
    'AL50': (f74, 50, 38),
    'AM50': (f75, 50, 39),
    'AN50': (f76, 50, 40),
    'AO50': (f77, 50, 41),
    'AP50': (f78, 50, 42),
    'AQ50': (f79, 50, 43),
    'AR50': (f80, 50, 44),
    'AS50': (f81, 50, 45),
    'AT50': (f82, 50, 46),
    'AU50': (f83, 50, 47),
    'AV50': (f84, 50, 48),
    'AW50': (f85, 50, 49),
    'AX50': (f86, 50, 50),
    'AY50': (f87, 50, 51),
    'AZ50': (f88, 50, 52),
    'BA50': (f89, 50, 53),
    'BB50': (f90, 50, 54),
    'A51': (f572, 51, 1),
    'B51': (f7, 51, 2),
    'C51': (f573, 51, 3),
    'D51': (f574, 51, 4),
    'E51': (f575, 51, 5),
    'F51': (f576, 51, 6),
    'G51': (f12, 51, 7),
    'H51': (f13, 51, 8),
    'I51': (f14, 51, 9),
    'J51': (f15, 51, 10),
    'K51': (f577, 51, 11),
    'L51': (f578, 51, 12),
    'M51': (f579, 51, 13),
    'Q51': (f19, 51, 17),
    'R51': (f20, 51, 18),
    'T51': (f580, 51, 20),
    'U51': (f22, 51, 21),
    'V51': (f23, 51, 22),
    'W51': (f24, 51, 23),
    'X51': (f25, 51, 24),
    'Y51': (f26, 51, 25),
    'Z51': (f27, 51, 26),
    'AA51': (f581, 51, 27),
    'AB51': (f582, 51, 28),
    'AC51': (f583, 51, 29),
    'AD51': (f584, 51, 30),
    'AE51': (f32, 51, 31),
    'AF51': (f33, 51, 32),
    'AG51': (f69, 51, 33),
    'AH51': (f70, 51, 34),
    'AI51': (f71, 51, 35),
    'AJ51': (f72, 51, 36),
    'AK51': (f73, 51, 37),
    'AL51': (f74, 51, 38),
    'AM51': (f75, 51, 39),
    'AN51': (f76, 51, 40),
    'AO51': (f77, 51, 41),
    'AP51': (f78, 51, 42),
    'AQ51': (f79, 51, 43),
    'AR51': (f80, 51, 44),
    'AS51': (f81, 51, 45),
    'AT51': (f82, 51, 46),
    'AU51': (f83, 51, 47),
    'AV51': (f84, 51, 48),
    'AW51': (f85, 51, 49),
    'AX51': (f86, 51, 50),
    'AY51': (f87, 51, 51),
    'AZ51': (f88, 51, 52),
    'BA51': (f89, 51, 53),
    'BB51': (f90, 51, 54),
    'A52': (f585, 52, 1),
    'B52': (f7, 52, 2),
    'C52': (f586, 52, 3),
    'D52': (f587, 52, 4),
    'E52': (f588, 52, 5),
    'F52': (f589, 52, 6),
    'G52': (f12, 52, 7),
    'H52': (f13, 52, 8),
    'I52': (f14, 52, 9),
    'J52': (f15, 52, 10),
    'K52': (f590, 52, 11),
    'L52': (f591, 52, 12),
    'M52': (f592, 52, 13),
    'Q52': (f19, 52, 17),
    'R52': (f20, 52, 18),
    'T52': (f593, 52, 20),
    'U52': (f22, 52, 21),
    'V52': (f23, 52, 22),
    'W52': (f24, 52, 23),
    'X52': (f25, 52, 24),
    'Y52': (f26, 52, 25),
    'Z52': (f27, 52, 26),
    'AA52': (f594, 52, 27),
    'AB52': (f595, 52, 28),
    'AC52': (f596, 52, 29),
    'AD52': (f597, 52, 30),
    'AE52': (f32, 52, 31),
    'AF52': (f33, 52, 32),
    'AG52': (f69, 52, 33),
    'AH52': (f70, 52, 34),
    'AI52': (f71, 52, 35),
    'AJ52': (f72, 52, 36),
    'AK52': (f73, 52, 37),
    'AL52': (f74, 52, 38),
    'AM52': (f75, 52, 39),
    'AN52': (f76, 52, 40),
    'AO52': (f77, 52, 41),
    'AP52': (f78, 52, 42),
    'AQ52': (f79, 52, 43),
    'AR52': (f80, 52, 44),
    'AS52': (f81, 52, 45),
    'AT52': (f82, 52, 46),
    'AU52': (f83, 52, 47),
    'AV52': (f84, 52, 48),
    'AW52': (f85, 52, 49),
    'AX52': (f86, 52, 50),
    'AY52': (f87, 52, 51),
    'AZ52': (f88, 52, 52),
    'BA52': (f89, 52, 53),
    'BB52': (f90, 52, 54),
    'A53': (f598, 53, 1),
    'B53': (f7, 53, 2),
    'C53': (f599, 53, 3),
    'D53': (f600, 53, 4),
    'E53': (f601, 53, 5),
    'F53': (f602, 53, 6),
    'G53': (f12, 53, 7),
    'H53': (f13, 53, 8),
    'I53': (f14, 53, 9),
    'J53': (f15, 53, 10),
    'K53': (f603, 53, 11),
    'L53': (f604, 53, 12),
    'M53': (f605, 53, 13),
    'Q53': (f19, 53, 17),
    'R53': (f20, 53, 18),
    'T53': (f606, 53, 20),
    'U53': (f22, 53, 21),
    'V53': (f23, 53, 22),
    'W53': (f24, 53, 23),
    'X53': (f25, 53, 24),
    'Y53': (f26, 53, 25),
    'Z53': (f27, 53, 26),
    'AA53': (f607, 53, 27),
    'AB53': (f608, 53, 28),
    'AC53': (f609, 53, 29),
    'AD53': (f610, 53, 30),
    'AE53': (f32, 53, 31),
    'AF53': (f33, 53, 32),
    'AG53': (f69, 53, 33),
    'AH53': (f70, 53, 34),
    'AI53': (f71, 53, 35),
    'AJ53': (f72, 53, 36),
    'AK53': (f73, 53, 37),
    'AL53': (f74, 53, 38),
    'AM53': (f75, 53, 39),
    'AN53': (f76, 53, 40),
    'AO53': (f77, 53, 41),
    'AP53': (f78, 53, 42),
    'AQ53': (f79, 53, 43),
    'AR53': (f80, 53, 44),
    'AS53': (f81, 53, 45),
    'AT53': (f82, 53, 46),
    'AU53': (f83, 53, 47),
    'AV53': (f84, 53, 48),
    'AW53': (f85, 53, 49),
    'AX53': (f86, 53, 50),
    'AY53': (f87, 53, 51),
    'AZ53': (f88, 53, 52),
    'BA53': (f89, 53, 53),
    'BB53': (f90, 53, 54),
    'A54': (f611, 54, 1),
    'B54': (f7, 54, 2),
    'C54': (f612, 54, 3),
    'D54': (f613, 54, 4),
    'E54': (f614, 54, 5),
    'F54': (f615, 54, 6),
    'G54': (f12, 54, 7),
    'H54': (f13, 54, 8),
    'I54': (f14, 54, 9),
    'J54': (f15, 54, 10),
    'K54': (f616, 54, 11),
    'L54': (f617, 54, 12),
    'M54': (f618, 54, 13),
    'Q54': (f19, 54, 17),
    'R54': (f20, 54, 18),
    'T54': (f619, 54, 20),
    'U54': (f22, 54, 21),
    'V54': (f23, 54, 22),
    'W54': (f24, 54, 23),
    'X54': (f25, 54, 24),
    'Y54': (f26, 54, 25),
    'Z54': (f27, 54, 26),
    'AA54': (f620, 54, 27),
    'AB54': (f621, 54, 28),
    'AC54': (f622, 54, 29),
    'AD54': (f623, 54, 30),
    'AE54': (f32, 54, 31),
    'AF54': (f33, 54, 32),
    'AG54': (f69, 54, 33),
    'AH54': (f70, 54, 34),
    'AI54': (f71, 54, 35),
    'AJ54': (f72, 54, 36),
    'AK54': (f73, 54, 37),
    'AL54': (f74, 54, 38),
    'AM54': (f75, 54, 39),
    'AN54': (f76, 54, 40),
    'AO54': (f77, 54, 41),
    'AP54': (f78, 54, 42),
    'AQ54': (f79, 54, 43),
    'AR54': (f80, 54, 44),
    'AS54': (f81, 54, 45),
    'AT54': (f82, 54, 46),
    'AU54': (f83, 54, 47),
    'AV54': (f84, 54, 48),
    'AW54': (f85, 54, 49),
    'AX54': (f86, 54, 50),
    'AY54': (f87, 54, 51),
    'AZ54': (f88, 54, 52),
    'BA54': (f89, 54, 53),
    'BB54': (f90, 54, 54),
    'A55': (f624, 55, 1),
    'B55': (f7, 55, 2),
    'C55': (f625, 55, 3),
    'D55': (f626, 55, 4),
    'E55': (f627, 55, 5),
    'F55': (f628, 55, 6),
    'G55': (f12, 55, 7),
    'H55': (f13, 55, 8),
    'I55': (f14, 55, 9),
    'J55': (f15, 55, 10),
    'K55': (f629, 55, 11),
    'L55': (f630, 55, 12),
    'M55': (f631, 55, 13),
    'Q55': (f19, 55, 17),
    'R55': (f20, 55, 18),
    'T55': (f632, 55, 20),
    'U55': (f22, 55, 21),
    'V55': (f23, 55, 22),
    'W55': (f24, 55, 23),
    'X55': (f25, 55, 24),
    'Y55': (f26, 55, 25),
    'Z55': (f27, 55, 26),
    'AA55': (f633, 55, 27),
    'AB55': (f634, 55, 28),
    'AC55': (f635, 55, 29),
    'AD55': (f636, 55, 30),
    'AE55': (f32, 55, 31),
    'AF55': (f33, 55, 32),
    'AG55': (f69, 55, 33),
    'AH55': (f70, 55, 34),
    'AI55': (f71, 55, 35),
    'AJ55': (f72, 55, 36),
    'AK55': (f73, 55, 37),
    'AL55': (f74, 55, 38),
    'AM55': (f75, 55, 39),
    'AN55': (f76, 55, 40),
    'AO55': (f77, 55, 41),
    'AP55': (f78, 55, 42),
    'AQ55': (f79, 55, 43),
    'AR55': (f80, 55, 44),
    'AS55': (f81, 55, 45),
    'AT55': (f82, 55, 46),
    'AU55': (f83, 55, 47),
    'AV55': (f84, 55, 48),
    'AW55': (f85, 55, 49),
    'AX55': (f86, 55, 50),
    'AY55': (f87, 55, 51),
    'AZ55': (f88, 55, 52),
    'BA55': (f89, 55, 53),
    'BB55': (f90, 55, 54),
    'A56': (f637, 56, 1),
    'B56': (f7, 56, 2),
    'C56': (f638, 56, 3),
    'D56': (f639, 56, 4),
    'E56': (f640, 56, 5),
    'F56': (f641, 56, 6),
    'G56': (f12, 56, 7),
    'H56': (f13, 56, 8),
    'I56': (f14, 56, 9),
    'J56': (f15, 56, 10),
    'K56': (f642, 56, 11),
    'L56': (f643, 56, 12),
    'M56': (f644, 56, 13),
    'Q56': (f19, 56, 17),
    'R56': (f20, 56, 18),
    'T56': (f645, 56, 20),
    'U56': (f22, 56, 21),
    'V56': (f23, 56, 22),
    'W56': (f24, 56, 23),
    'X56': (f25, 56, 24),
    'Y56': (f26, 56, 25),
    'Z56': (f27, 56, 26),
    'AA56': (f646, 56, 27),
    'AB56': (f647, 56, 28),
    'AC56': (f648, 56, 29),
    'AD56': (f649, 56, 30),
    'AE56': (f32, 56, 31),
    'AF56': (f33, 56, 32),
    'AG56': (f69, 56, 33),
    'AH56': (f70, 56, 34),
    'AI56': (f71, 56, 35),
    'AJ56': (f72, 56, 36),
    'AK56': (f73, 56, 37),
    'AL56': (f74, 56, 38),
    'AM56': (f75, 56, 39),
    'AN56': (f76, 56, 40),
    'AO56': (f77, 56, 41),
    'AP56': (f78, 56, 42),
    'AQ56': (f79, 56, 43),
    'AR56': (f80, 56, 44),
    'AS56': (f81, 56, 45),
    'AT56': (f82, 56, 46),
    'AU56': (f83, 56, 47),
    'AV56': (f84, 56, 48),
    'AW56': (f85, 56, 49),
    'AX56': (f86, 56, 50),
    'AY56': (f87, 56, 51),
    'AZ56': (f88, 56, 52),
    'BA56': (f89, 56, 53),
    'BB56': (f90, 56, 54),
    'A57': (f650, 57, 1),
    'B57': (f7, 57, 2),
    'C57': (f651, 57, 3),
    'D57': (f652, 57, 4),
    'E57': (f653, 57, 5),
    'F57': (f654, 57, 6),
    'G57': (f12, 57, 7),
    'H57': (f13, 57, 8),
    'I57': (f14, 57, 9),
    'J57': (f15, 57, 10),
    'K57': (f655, 57, 11),
    'L57': (f656, 57, 12),
    'M57': (f657, 57, 13),
    'Q57': (f19, 57, 17),
    'R57': (f20, 57, 18),
    'T57': (f658, 57, 20),
    'U57': (f22, 57, 21),
    'V57': (f23, 57, 22),
    'W57': (f24, 57, 23),
    'X57': (f25, 57, 24),
    'Y57': (f26, 57, 25),
    'Z57': (f27, 57, 26),
    'AA57': (f659, 57, 27),
    'AB57': (f660, 57, 28),
    'AC57': (f661, 57, 29),
    'AD57': (f662, 57, 30),
    'AE57': (f32, 57, 31),
    'AF57': (f33, 57, 32),
    'AG57': (f69, 57, 33),
    'AH57': (f70, 57, 34),
    'AI57': (f71, 57, 35),
    'AJ57': (f72, 57, 36),
    'AK57': (f73, 57, 37),
    'AL57': (f74, 57, 38),
    'AM57': (f75, 57, 39),
    'AN57': (f76, 57, 40),
    'AO57': (f77, 57, 41),
    'AP57': (f78, 57, 42),
    'AQ57': (f79, 57, 43),
    'AR57': (f80, 57, 44),
    'AS57': (f81, 57, 45),
    'AT57': (f82, 57, 46),
    'AU57': (f83, 57, 47),
    'AV57': (f84, 57, 48),
    'AW57': (f85, 57, 49),
    'AX57': (f86, 57, 50),
    'AY57': (f87, 57, 51),
    'AZ57': (f88, 57, 52),
    'BA57': (f89, 57, 53),
    'BB57': (f90, 57, 54),
    'A58': (f663, 58, 1),
    'B58': (f7, 58, 2),
    'C58': (f664, 58, 3),
    'D58': (f665, 58, 4),
    'E58': (f666, 58, 5),
    'F58': (f667, 58, 6),
    'G58': (f12, 58, 7),
    'H58': (f13, 58, 8),
    'I58': (f14, 58, 9),
    'J58': (f15, 58, 10),
    'K58': (f668, 58, 11),
    'L58': (f669, 58, 12),
    'M58': (f670, 58, 13),
    'Q58': (f19, 58, 17),
    'R58': (f20, 58, 18),
    'T58': (f671, 58, 20),
    'U58': (f22, 58, 21),
    'V58': (f23, 58, 22),
    'W58': (f24, 58, 23),
    'X58': (f25, 58, 24),
    'Y58': (f26, 58, 25),
    'Z58': (f27, 58, 26),
    'AA58': (f672, 58, 27),
    'AB58': (f673, 58, 28),
    'AC58': (f674, 58, 29),
    'AD58': (f675, 58, 30),
    'AE58': (f32, 58, 31),
    'AF58': (f33, 58, 32),
    'AG58': (f69, 58, 33),
    'AH58': (f70, 58, 34),
    'AI58': (f71, 58, 35),
    'AJ58': (f72, 58, 36),
    'AK58': (f73, 58, 37),
    'AL58': (f74, 58, 38),
    'AM58': (f75, 58, 39),
    'AN58': (f76, 58, 40),
    'AO58': (f77, 58, 41),
    'AP58': (f78, 58, 42),
    'AQ58': (f79, 58, 43),
    'AR58': (f80, 58, 44),
    'AS58': (f81, 58, 45),
    'AT58': (f82, 58, 46),
    'AU58': (f83, 58, 47),
    'AV58': (f84, 58, 48),
    'AW58': (f85, 58, 49),
    'AX58': (f86, 58, 50),
    'AY58': (f87, 58, 51),
    'AZ58': (f88, 58, 52),
    'BA58': (f89, 58, 53),
    'BB58': (f90, 58, 54),
    'A59': (f676, 59, 1),
    'B59': (f7, 59, 2),
    'C59': (f677, 59, 3),
    'D59': (f678, 59, 4),
    'E59': (f679, 59, 5),
    'F59': (f680, 59, 6),
    'G59': (f12, 59, 7),
    'H59': (f13, 59, 8),
    'I59': (f14, 59, 9),
    'J59': (f15, 59, 10),
    'K59': (f681, 59, 11),
    'L59': (f682, 59, 12),
    'M59': (f683, 59, 13),
    'Q59': (f19, 59, 17),
    'R59': (f20, 59, 18),
    'T59': (f684, 59, 20),
    'U59': (f22, 59, 21),
    'V59': (f23, 59, 22),
    'W59': (f24, 59, 23),
    'X59': (f25, 59, 24),
    'Y59': (f26, 59, 25),
    'Z59': (f27, 59, 26),
    'AA59': (f685, 59, 27),
    'AB59': (f686, 59, 28),
    'AC59': (f687, 59, 29),
    'AD59': (f688, 59, 30),
    'AE59': (f32, 59, 31),
    'AF59': (f33, 59, 32),
    'AG59': (f69, 59, 33),
    'AH59': (f70, 59, 34),
    'AI59': (f71, 59, 35),
    'AJ59': (f72, 59, 36),
    'AK59': (f73, 59, 37),
    'AL59': (f74, 59, 38),
    'AM59': (f75, 59, 39),
    'AN59': (f76, 59, 40),
    'AO59': (f77, 59, 41),
    'AP59': (f78, 59, 42),
    'AQ59': (f79, 59, 43),
    'AR59': (f80, 59, 44),
    'AS59': (f81, 59, 45),
    'AT59': (f82, 59, 46),
    'AU59': (f83, 59, 47),
    'AV59': (f84, 59, 48),
    'AW59': (f85, 59, 49),
    'AX59': (f86, 59, 50),
    'AY59': (f87, 59, 51),
    'AZ59': (f88, 59, 52),
    'BA59': (f89, 59, 53),
    'BB59': (f90, 59, 54),
    'A60': (f689, 60, 1),
    'B60': (f7, 60, 2),
    'C60': (f690, 60, 3),
    'D60': (f691, 60, 4),
    'E60': (f692, 60, 5),
    'F60': (f693, 60, 6),
    'G60': (f12, 60, 7),
    'H60': (f13, 60, 8),
    'I60': (f14, 60, 9),
    'J60': (f15, 60, 10),
    'K60': (f694, 60, 11),
    'L60': (f695, 60, 12),
    'M60': (f696, 60, 13),
    'Q60': (f19, 60, 17),
    'R60': (f20, 60, 18),
    'T60': (f697, 60, 20),
    'U60': (f22, 60, 21),
    'V60': (f23, 60, 22),
    'W60': (f24, 60, 23),
    'X60': (f25, 60, 24),
    'Y60': (f26, 60, 25),
    'Z60': (f27, 60, 26),
    'AA60': (f698, 60, 27),
    'AB60': (f699, 60, 28),
    'AC60': (f700, 60, 29),
    'AD60': (f701, 60, 30),
    'AE60': (f32, 60, 31),
    'AF60': (f33, 60, 32),
    'AG60': (f69, 60, 33),
    'AH60': (f70, 60, 34),
    'AI60': (f71, 60, 35),
    'AJ60': (f72, 60, 36),
    'AK60': (f73, 60, 37),
    'AL60': (f74, 60, 38),
    'AM60': (f75, 60, 39),
    'AN60': (f76, 60, 40),
    'AO60': (f77, 60, 41),
    'AP60': (f78, 60, 42),
    'AQ60': (f79, 60, 43),
    'AR60': (f80, 60, 44),
    'AS60': (f81, 60, 45),
    'AT60': (f82, 60, 46),
    'AU60': (f83, 60, 47),
    'AV60': (f84, 60, 48),
    'AW60': (f85, 60, 49),
    'AX60': (f86, 60, 50),
    'AY60': (f87, 60, 51),
    'AZ60': (f88, 60, 52),
    'BA60': (f89, 60, 53),
    'BB60': (f90, 60, 54),
    'A61': (f702, 61, 1),
    'B61': (f7, 61, 2),
    'C61': (f703, 61, 3),
    'D61': (f704, 61, 4),
    'E61': (f705, 61, 5),
    'F61': (f706, 61, 6),
    'G61': (f12, 61, 7),
    'H61': (f13, 61, 8),
    'I61': (f14, 61, 9),
    'J61': (f15, 61, 10),
    'K61': (f707, 61, 11),
    'L61': (f708, 61, 12),
    'M61': (f709, 61, 13),
    'Q61': (f19, 61, 17),
    'R61': (f20, 61, 18),
    'T61': (f710, 61, 20),
    'U61': (f22, 61, 21),
    'V61': (f23, 61, 22),
    'W61': (f24, 61, 23),
    'X61': (f25, 61, 24),
    'Y61': (f26, 61, 25),
    'Z61': (f27, 61, 26),
    'AA61': (f711, 61, 27),
    'AB61': (f712, 61, 28),
    'AC61': (f713, 61, 29),
    'AD61': (f714, 61, 30),
    'AE61': (f32, 61, 31),
    'AF61': (f33, 61, 32),
    'AG61': (f69, 61, 33),
    'AH61': (f70, 61, 34),
    'AI61': (f71, 61, 35),
    'AJ61': (f72, 61, 36),
    'AK61': (f73, 61, 37),
    'AL61': (f74, 61, 38),
    'AM61': (f75, 61, 39),
    'AN61': (f76, 61, 40),
    'AO61': (f77, 61, 41),
    'AP61': (f78, 61, 42),
    'AQ61': (f79, 61, 43),
    'AR61': (f80, 61, 44),
    'AS61': (f81, 61, 45),
    'AT61': (f82, 61, 46),
    'AU61': (f83, 61, 47),
    'AV61': (f84, 61, 48),
    'AW61': (f85, 61, 49),
    'AX61': (f86, 61, 50),
    'AY61': (f87, 61, 51),
    'AZ61': (f88, 61, 52),
    'BA61': (f89, 61, 53),
    'BB61': (f90, 61, 54),
    'A62': (f715, 62, 1),
    'B62': (f7, 62, 2),
    'C62': (f716, 62, 3),
    'D62': (f717, 62, 4),
    'E62': (f718, 62, 5),
    'F62': (f719, 62, 6),
    'G62': (f12, 62, 7),
    'H62': (f13, 62, 8),
    'I62': (f14, 62, 9),
    'J62': (f15, 62, 10),
    'K62': (f720, 62, 11),
    'L62': (f721, 62, 12),
    'M62': (f722, 62, 13),
    'Q62': (f19, 62, 17),
    'R62': (f20, 62, 18),
    'T62': (f723, 62, 20),
    'U62': (f22, 62, 21),
    'V62': (f23, 62, 22),
    'W62': (f24, 62, 23),
    'X62': (f25, 62, 24),
    'Y62': (f26, 62, 25),
    'Z62': (f27, 62, 26),
    'AA62': (f724, 62, 27),
    'AB62': (f725, 62, 28),
    'AC62': (f726, 62, 29),
    'AD62': (f727, 62, 30),
    'AE62': (f32, 62, 31),
    'AF62': (f33, 62, 32),
    'AG62': (f69, 62, 33),
    'AH62': (f70, 62, 34),
    'AI62': (f71, 62, 35),
    'AJ62': (f72, 62, 36),
    'AK62': (f73, 62, 37),
    'AL62': (f74, 62, 38),
    'AM62': (f75, 62, 39),
    'AN62': (f76, 62, 40),
    'AO62': (f77, 62, 41),
    'AP62': (f78, 62, 42),
    'AQ62': (f79, 62, 43),
    'AR62': (f80, 62, 44),
    'AS62': (f81, 62, 45),
    'AT62': (f82, 62, 46),
    'AU62': (f83, 62, 47),
    'AV62': (f84, 62, 48),
    'AW62': (f85, 62, 49),
    'AX62': (f86, 62, 50),
    'AY62': (f87, 62, 51),
    'AZ62': (f88, 62, 52),
    'BA62': (f89, 62, 53),
    'BB62': (f90, 62, 54),
    'A63': (f728, 63, 1),
    'B63': (f7, 63, 2),
    'C63': (f729, 63, 3),
    'D63': (f730, 63, 4),
    'E63': (f731, 63, 5),
    'F63': (f732, 63, 6),
    'G63': (f12, 63, 7),
    'H63': (f13, 63, 8),
    'I63': (f14, 63, 9),
    'J63': (f15, 63, 10),
    'K63': (f733, 63, 11),
    'L63': (f734, 63, 12),
    'M63': (f735, 63, 13),
    'Q63': (f19, 63, 17),
    'R63': (f20, 63, 18),
    'T63': (f736, 63, 20),
    'U63': (f22, 63, 21),
    'V63': (f23, 63, 22),
    'W63': (f24, 63, 23),
    'X63': (f25, 63, 24),
    'Y63': (f26, 63, 25),
    'Z63': (f27, 63, 26),
    'AA63': (f737, 63, 27),
    'AB63': (f738, 63, 28),
    'AC63': (f739, 63, 29),
    'AD63': (f740, 63, 30),
    'AE63': (f32, 63, 31),
    'AF63': (f33, 63, 32),
    'AG63': (f69, 63, 33),
    'AH63': (f70, 63, 34),
    'AI63': (f71, 63, 35),
    'AJ63': (f72, 63, 36),
    'AK63': (f73, 63, 37),
    'AL63': (f74, 63, 38),
    'AM63': (f75, 63, 39),
    'AN63': (f76, 63, 40),
    'AO63': (f77, 63, 41),
    'AP63': (f78, 63, 42),
    'AQ63': (f79, 63, 43),
    'AR63': (f80, 63, 44),
    'AS63': (f81, 63, 45),
    'AT63': (f82, 63, 46),
    'AU63': (f83, 63, 47),
    'AV63': (f84, 63, 48),
    'AW63': (f85, 63, 49),
    'AX63': (f86, 63, 50),
    'AY63': (f87, 63, 51),
    'AZ63': (f88, 63, 52),
    'BA63': (f89, 63, 53),
    'BB63': (f90, 63, 54),
    'A64': (f741, 64, 1),
    'B64': (f7, 64, 2),
    'C64': (f742, 64, 3),
    'D64': (f743, 64, 4),
    'E64': (f744, 64, 5),
    'F64': (f745, 64, 6),
    'G64': (f12, 64, 7),
    'H64': (f13, 64, 8),
    'I64': (f14, 64, 9),
    'J64': (f15, 64, 10),
    'K64': (f746, 64, 11),
    'L64': (f747, 64, 12),
    'M64': (f748, 64, 13),
    'Q64': (f19, 64, 17),
    'R64': (f20, 64, 18),
    'T64': (f749, 64, 20),
    'U64': (f22, 64, 21),
    'V64': (f23, 64, 22),
    'W64': (f24, 64, 23),
    'X64': (f25, 64, 24),
    'Y64': (f26, 64, 25),
    'Z64': (f27, 64, 26),
    'AA64': (f750, 64, 27),
    'AB64': (f751, 64, 28),
    'AC64': (f752, 64, 29),
    'AD64': (f753, 64, 30),
    'AE64': (f32, 64, 31),
    'AF64': (f33, 64, 32),
    'AG64': (f69, 64, 33),
    'AH64': (f70, 64, 34),
    'AI64': (f71, 64, 35),
    'AJ64': (f72, 64, 36),
    'AK64': (f73, 64, 37),
    'AL64': (f74, 64, 38),
    'AM64': (f75, 64, 39),
    'AN64': (f76, 64, 40),
    'AO64': (f77, 64, 41),
    'AP64': (f78, 64, 42),
    'AQ64': (f79, 64, 43),
    'AR64': (f80, 64, 44),
    'AS64': (f81, 64, 45),
    'AT64': (f82, 64, 46),
    'AU64': (f83, 64, 47),
    'AV64': (f84, 64, 48),
    'AW64': (f85, 64, 49),
    'AX64': (f86, 64, 50),
    'AY64': (f87, 64, 51),
    'AZ64': (f88, 64, 52),
    'BA64': (f89, 64, 53),
    'BB64': (f90, 64, 54),
    'A65': (f754, 65, 1),
    'B65': (f7, 65, 2),
    'C65': (f755, 65, 3),
    'D65': (f756, 65, 4),
    'E65': (f757, 65, 5),
    'F65': (f758, 65, 6),
    'G65': (f12, 65, 7),
    'H65': (f13, 65, 8),
    'I65': (f14, 65, 9),
    'J65': (f15, 65, 10),
    'K65': (f759, 65, 11),
    'L65': (f760, 65, 12),
    'M65': (f761, 65, 13),
    'Q65': (f19, 65, 17),
    'R65': (f20, 65, 18),
    'T65': (f762, 65, 20),
    'U65': (f22, 65, 21),
    'V65': (f23, 65, 22),
    'W65': (f24, 65, 23),
    'X65': (f25, 65, 24),
    'Y65': (f26, 65, 25),
    'Z65': (f27, 65, 26),
    'AA65': (f763, 65, 27),
    'AB65': (f764, 65, 28),
    'AC65': (f765, 65, 29),
    'AD65': (f766, 65, 30),
    'AE65': (f32, 65, 31),
    'AF65': (f33, 65, 32),
    'AG65': (f69, 65, 33),
    'AH65': (f70, 65, 34),
    'AI65': (f71, 65, 35),
    'AJ65': (f72, 65, 36),
    'AK65': (f73, 65, 37),
    'AL65': (f74, 65, 38),
    'AM65': (f75, 65, 39),
    'AN65': (f76, 65, 40),
    'AO65': (f77, 65, 41),
    'AP65': (f78, 65, 42),
    'AQ65': (f79, 65, 43),
    'AR65': (f80, 65, 44),
    'AS65': (f81, 65, 45),
    'AT65': (f82, 65, 46),
    'AU65': (f83, 65, 47),
    'AV65': (f84, 65, 48),
    'AW65': (f85, 65, 49),
    'AX65': (f86, 65, 50),
    'AY65': (f87, 65, 51),
    'AZ65': (f88, 65, 52),
    'BA65': (f89, 65, 53),
    'BB65': (f90, 65, 54),
    'A66': (f767, 66, 1),
    'B66': (f7, 66, 2),
    'C66': (f768, 66, 3),
    'D66': (f769, 66, 4),
    'E66': (f770, 66, 5),
    'F66': (f771, 66, 6),
    'G66': (f12, 66, 7),
    'H66': (f13, 66, 8),
    'I66': (f14, 66, 9),
    'J66': (f15, 66, 10),
    'K66': (f772, 66, 11),
    'L66': (f773, 66, 12),
    'M66': (f774, 66, 13),
    'Q66': (f19, 66, 17),
    'R66': (f20, 66, 18),
    'T66': (f775, 66, 20),
    'U66': (f22, 66, 21),
    'V66': (f23, 66, 22),
    'W66': (f24, 66, 23),
    'X66': (f25, 66, 24),
    'Y66': (f26, 66, 25),
    'Z66': (f27, 66, 26),
    'AA66': (f776, 66, 27),
    'AB66': (f777, 66, 28),
    'AC66': (f778, 66, 29),
    'AD66': (f779, 66, 30),
    'AE66': (f32, 66, 31),
    'AF66': (f33, 66, 32),
    'AG66': (f69, 66, 33),
    'AH66': (f70, 66, 34),
    'AI66': (f71, 66, 35),
    'AJ66': (f72, 66, 36),
    'AK66': (f73, 66, 37),
    'AL66': (f74, 66, 38),
    'AM66': (f75, 66, 39),
    'AN66': (f76, 66, 40),
    'AO66': (f77, 66, 41),
    'AP66': (f78, 66, 42),
    'AQ66': (f79, 66, 43),
    'AR66': (f80, 66, 44),
    'AS66': (f81, 66, 45),
    'AT66': (f82, 66, 46),
    'AU66': (f83, 66, 47),
    'AV66': (f84, 66, 48),
    'AW66': (f85, 66, 49),
    'AX66': (f86, 66, 50),
    'AY66': (f87, 66, 51),
    'AZ66': (f88, 66, 52),
    'BA66': (f89, 66, 53),
    'BB66': (f90, 66, 54),
    'A67': (f780, 67, 1),
    'B67': (f7, 67, 2),
    'C67': (f781, 67, 3),
    'D67': (f782, 67, 4),
    'E67': (f783, 67, 5),
    'F67': (f784, 67, 6),
    'G67': (f12, 67, 7),
    'H67': (f13, 67, 8),
    'I67': (f14, 67, 9),
    'J67': (f15, 67, 10),
    'K67': (f785, 67, 11),
    'L67': (f786, 67, 12),
    'M67': (f787, 67, 13),
    'Q67': (f19, 67, 17),
    'R67': (f20, 67, 18),
    'T67': (f788, 67, 20),
    'U67': (f22, 67, 21),
    'V67': (f23, 67, 22),
    'W67': (f24, 67, 23),
    'X67': (f25, 67, 24),
    'Y67': (f26, 67, 25),
    'Z67': (f27, 67, 26),
    'AA67': (f789, 67, 27),
    'AB67': (f790, 67, 28),
    'AC67': (f791, 67, 29),
    'AD67': (f792, 67, 30),
    'AE67': (f32, 67, 31),
    'AF67': (f33, 67, 32),
    'AG67': (f69, 67, 33),
    'AH67': (f70, 67, 34),
    'AI67': (f71, 67, 35),
    'AJ67': (f72, 67, 36),
    'AK67': (f73, 67, 37),
    'AL67': (f74, 67, 38),
    'AM67': (f75, 67, 39),
    'AN67': (f76, 67, 40),
    'AO67': (f77, 67, 41),
    'AP67': (f78, 67, 42),
    'AQ67': (f79, 67, 43),
    'AR67': (f80, 67, 44),
    'AS67': (f81, 67, 45),
    'AT67': (f82, 67, 46),
    'AU67': (f83, 67, 47),
    'AV67': (f84, 67, 48),
    'AW67': (f85, 67, 49),
    'AX67': (f86, 67, 50),
    'AY67': (f87, 67, 51),
    'AZ67': (f88, 67, 52),
    'BA67': (f89, 67, 53),
    'BB67': (f90, 67, 54),
    'A68': (f793, 68, 1),
    'B68': (f7, 68, 2),
    'C68': (f794, 68, 3),
    'D68': (f795, 68, 4),
    'E68': (f796, 68, 5),
    'F68': (f797, 68, 6),
    'G68': (f12, 68, 7),
    'H68': (f13, 68, 8),
    'I68': (f14, 68, 9),
    'J68': (f15, 68, 10),
    'K68': (f798, 68, 11),
    'L68': (f799, 68, 12),
    'M68': (f800, 68, 13),
    'Q68': (f19, 68, 17),
    'R68': (f20, 68, 18),
    'T68': (f801, 68, 20),
    'U68': (f22, 68, 21),
    'V68': (f23, 68, 22),
    'W68': (f24, 68, 23),
    'X68': (f25, 68, 24),
    'Y68': (f26, 68, 25),
    'Z68': (f27, 68, 26),
    'AA68': (f802, 68, 27),
    'AB68': (f803, 68, 28),
    'AC68': (f804, 68, 29),
    'AD68': (f805, 68, 30),
    'AE68': (f32, 68, 31),
    'AF68': (f33, 68, 32),
    'AG68': (f69, 68, 33),
    'AH68': (f70, 68, 34),
    'AI68': (f71, 68, 35),
    'AJ68': (f72, 68, 36),
    'AK68': (f73, 68, 37),
    'AL68': (f74, 68, 38),
    'AM68': (f75, 68, 39),
    'AN68': (f76, 68, 40),
    'AO68': (f77, 68, 41),
    'AP68': (f78, 68, 42),
    'AQ68': (f79, 68, 43),
    'AR68': (f80, 68, 44),
    'AS68': (f81, 68, 45),
    'AT68': (f82, 68, 46),
    'AU68': (f83, 68, 47),
    'AV68': (f84, 68, 48),
    'AW68': (f85, 68, 49),
    'AX68': (f86, 68, 50),
    'AY68': (f87, 68, 51),
    'AZ68': (f88, 68, 52),
    'BA68': (f89, 68, 53),
    'BB68': (f90, 68, 54),
    'A69': (f806, 69, 1),
    'B69': (f7, 69, 2),
    'C69': (f807, 69, 3),
    'D69': (f808, 69, 4),
    'E69': (f809, 69, 5),
    'F69': (f810, 69, 6),
    'G69': (f12, 69, 7),
    'H69': (f13, 69, 8),
    'I69': (f14, 69, 9),
    'J69': (f15, 69, 10),
    'K69': (f811, 69, 11),
    'L69': (f812, 69, 12),
    'M69': (f813, 69, 13),
    'Q69': (f19, 69, 17),
    'R69': (f20, 69, 18),
    'T69': (f814, 69, 20),
    'U69': (f22, 69, 21),
    'V69': (f23, 69, 22),
    'W69': (f24, 69, 23),
    'X69': (f25, 69, 24),
    'Y69': (f26, 69, 25),
    'Z69': (f27, 69, 26),
    'AA69': (f815, 69, 27),
    'AB69': (f816, 69, 28),
    'AC69': (f817, 69, 29),
    'AD69': (f818, 69, 30),
    'AE69': (f32, 69, 31),
    'AF69': (f33, 69, 32),
    'AG69': (f69, 69, 33),
    'AH69': (f70, 69, 34),
    'AI69': (f71, 69, 35),
    'AJ69': (f72, 69, 36),
    'AK69': (f73, 69, 37),
    'AL69': (f74, 69, 38),
    'AM69': (f75, 69, 39),
    'AN69': (f76, 69, 40),
    'AO69': (f77, 69, 41),
    'AP69': (f78, 69, 42),
    'AQ69': (f79, 69, 43),
    'AR69': (f80, 69, 44),
    'AS69': (f81, 69, 45),
    'AT69': (f82, 69, 46),
    'AU69': (f83, 69, 47),
    'AV69': (f84, 69, 48),
    'AW69': (f85, 69, 49),
    'AX69': (f86, 69, 50),
    'AY69': (f87, 69, 51),
    'AZ69': (f88, 69, 52),
    'BA69': (f89, 69, 53),
    'BB69': (f90, 69, 54),
    'A70': (f819, 70, 1),
    'B70': (f7, 70, 2),
    'C70': (f820, 70, 3),
    'D70': (f821, 70, 4),
    'E70': (f822, 70, 5),
    'F70': (f823, 70, 6),
    'G70': (f12, 70, 7),
    'H70': (f13, 70, 8),
    'I70': (f14, 70, 9),
    'J70': (f15, 70, 10),
    'K70': (f824, 70, 11),
    'L70': (f825, 70, 12),
    'M70': (f826, 70, 13),
    'Q70': (f19, 70, 17),
    'R70': (f20, 70, 18),
    'T70': (f827, 70, 20),
    'U70': (f22, 70, 21),
    'V70': (f23, 70, 22),
    'W70': (f24, 70, 23),
    'X70': (f25, 70, 24),
    'Y70': (f26, 70, 25),
    'Z70': (f27, 70, 26),
    'AA70': (f828, 70, 27),
    'AB70': (f829, 70, 28),
    'AC70': (f830, 70, 29),
    'AD70': (f831, 70, 30),
    'AE70': (f32, 70, 31),
    'AF70': (f33, 70, 32),
    'AG70': (f69, 70, 33),
    'AH70': (f70, 70, 34),
    'AI70': (f71, 70, 35),
    'AJ70': (f72, 70, 36),
    'AK70': (f73, 70, 37),
    'AL70': (f74, 70, 38),
    'AM70': (f75, 70, 39),
    'AN70': (f76, 70, 40),
    'AO70': (f77, 70, 41),
    'AP70': (f78, 70, 42),
    'AQ70': (f79, 70, 43),
    'AR70': (f80, 70, 44),
    'AS70': (f81, 70, 45),
    'AT70': (f82, 70, 46),
    'AU70': (f83, 70, 47),
    'AV70': (f84, 70, 48),
    'AW70': (f85, 70, 49),
    'AX70': (f86, 70, 50),
    'AY70': (f87, 70, 51),
    'AZ70': (f88, 70, 52),
    'BA70': (f89, 70, 53),
    'BB70': (f90, 70, 54),
    'A71': (f832, 71, 1),
    'B71': (f7, 71, 2),
    'C71': (f833, 71, 3),
    'D71': (f834, 71, 4),
    'E71': (f835, 71, 5),
    'F71': (f836, 71, 6),
    'G71': (f12, 71, 7),
    'H71': (f13, 71, 8),
    'I71': (f14, 71, 9),
    'J71': (f15, 71, 10),
    'K71': (f837, 71, 11),
    'L71': (f838, 71, 12),
    'M71': (f839, 71, 13),
    'Q71': (f19, 71, 17),
    'R71': (f20, 71, 18),
    'T71': (f840, 71, 20),
    'U71': (f22, 71, 21),
    'V71': (f23, 71, 22),
    'W71': (f24, 71, 23),
    'X71': (f25, 71, 24),
    'Y71': (f26, 71, 25),
    'Z71': (f27, 71, 26),
    'AA71': (f841, 71, 27),
    'AB71': (f842, 71, 28),
    'AC71': (f843, 71, 29),
    'AD71': (f844, 71, 30),
    'AE71': (f32, 71, 31),
    'AF71': (f33, 71, 32),
    'AG71': (f69, 71, 33),
    'AH71': (f70, 71, 34),
    'AI71': (f71, 71, 35),
    'AJ71': (f72, 71, 36),
    'AK71': (f73, 71, 37),
    'AL71': (f74, 71, 38),
    'AM71': (f75, 71, 39),
    'AN71': (f76, 71, 40),
    'AO71': (f77, 71, 41),
    'AP71': (f78, 71, 42),
    'AQ71': (f79, 71, 43),
    'AR71': (f80, 71, 44),
    'AS71': (f81, 71, 45),
    'AT71': (f82, 71, 46),
    'AU71': (f83, 71, 47),
    'AV71': (f84, 71, 48),
    'AW71': (f85, 71, 49),
    'AX71': (f86, 71, 50),
    'AY71': (f87, 71, 51),
    'AZ71': (f88, 71, 52),
    'BA71': (f89, 71, 53),
    'BB71': (f90, 71, 54),
    'A72': (f845, 72, 1),
    'B72': (f7, 72, 2),
    'C72': (f846, 72, 3),
    'D72': (f847, 72, 4),
    'E72': (f848, 72, 5),
    'F72': (f849, 72, 6),
    'G72': (f12, 72, 7),
    'H72': (f13, 72, 8),
    'I72': (f14, 72, 9),
    'J72': (f15, 72, 10),
    'K72': (f850, 72, 11),
    'L72': (f851, 72, 12),
    'M72': (f852, 72, 13),
    'Q72': (f19, 72, 17),
    'R72': (f20, 72, 18),
    'T72': (f853, 72, 20),
    'U72': (f22, 72, 21),
    'V72': (f23, 72, 22),
    'W72': (f24, 72, 23),
    'X72': (f25, 72, 24),
    'Y72': (f26, 72, 25),
    'Z72': (f27, 72, 26),
    'AA72': (f854, 72, 27),
    'AB72': (f855, 72, 28),
    'AC72': (f856, 72, 29),
    'AD72': (f857, 72, 30),
    'AE72': (f32, 72, 31),
    'AF72': (f33, 72, 32),
    'AG72': (f69, 72, 33),
    'AH72': (f70, 72, 34),
    'AI72': (f71, 72, 35),
    'AJ72': (f72, 72, 36),
    'AK72': (f73, 72, 37),
    'AL72': (f74, 72, 38),
    'AM72': (f75, 72, 39),
    'AN72': (f76, 72, 40),
    'AO72': (f77, 72, 41),
    'AP72': (f78, 72, 42),
    'AQ72': (f79, 72, 43),
    'AR72': (f80, 72, 44),
    'AS72': (f81, 72, 45),
    'AT72': (f82, 72, 46),
    'AU72': (f83, 72, 47),
    'AV72': (f84, 72, 48),
    'AW72': (f85, 72, 49),
    'AX72': (f86, 72, 50),
    'AY72': (f87, 72, 51),
    'AZ72': (f88, 72, 52),
    'BA72': (f89, 72, 53),
    'BB72': (f90, 72, 54),
    'A73': (f858, 73, 1),
    'B73': (f7, 73, 2),
    'C73': (f859, 73, 3),
    'D73': (f860, 73, 4),
    'E73': (f861, 73, 5),
    'F73': (f862, 73, 6),
    'G73': (f12, 73, 7),
    'H73': (f13, 73, 8),
    'I73': (f14, 73, 9),
    'J73': (f15, 73, 10),
    'K73': (f863, 73, 11),
    'L73': (f864, 73, 12),
    'M73': (f865, 73, 13),
    'Q73': (f19, 73, 17),
    'R73': (f20, 73, 18),
    'T73': (f866, 73, 20),
    'U73': (f22, 73, 21),
    'V73': (f23, 73, 22),
    'W73': (f24, 73, 23),
    'X73': (f25, 73, 24),
    'Y73': (f26, 73, 25),
    'Z73': (f27, 73, 26),
    'AA73': (f867, 73, 27),
    'AB73': (f868, 73, 28),
    'AC73': (f869, 73, 29),
    'AD73': (f870, 73, 30),
    'AE73': (f32, 73, 31),
    'AF73': (f33, 73, 32),
    'AG73': (f69, 73, 33),
    'AH73': (f70, 73, 34),
    'AI73': (f71, 73, 35),
    'AJ73': (f72, 73, 36),
    'AK73': (f73, 73, 37),
    'AL73': (f74, 73, 38),
    'AM73': (f75, 73, 39),
    'AN73': (f76, 73, 40),
    'AO73': (f77, 73, 41),
    'AP73': (f78, 73, 42),
    'AQ73': (f79, 73, 43),
    'AR73': (f80, 73, 44),
    'AS73': (f81, 73, 45),
    'AT73': (f82, 73, 46),
    'AU73': (f83, 73, 47),
    'AV73': (f84, 73, 48),
    'AW73': (f85, 73, 49),
    'AX73': (f86, 73, 50),
    'AY73': (f87, 73, 51),
    'AZ73': (f88, 73, 52),
    'BA73': (f89, 73, 53),
    'BB73': (f90, 73, 54),
    'A74': (f871, 74, 1),
    'B74': (f7, 74, 2),
    'C74': (f872, 74, 3),
    'D74': (f873, 74, 4),
    'E74': (f874, 74, 5),
    'F74': (f875, 74, 6),
    'G74': (f12, 74, 7),
    'H74': (f13, 74, 8),
    'I74': (f14, 74, 9),
    'J74': (f15, 74, 10),
    'K74': (f876, 74, 11),
    'L74': (f877, 74, 12),
    'M74': (f878, 74, 13),
    'Q74': (f19, 74, 17),
    'R74': (f20, 74, 18),
    'T74': (f879, 74, 20),
    'U74': (f22, 74, 21),
    'V74': (f23, 74, 22),
    'W74': (f24, 74, 23),
    'X74': (f25, 74, 24),
    'Y74': (f26, 74, 25),
    'Z74': (f27, 74, 26),
    'AA74': (f880, 74, 27),
    'AB74': (f881, 74, 28),
    'AC74': (f882, 74, 29),
    'AD74': (f883, 74, 30),
    'AE74': (f32, 74, 31),
    'AF74': (f33, 74, 32),
    'AG74': (f69, 74, 33),
    'AH74': (f70, 74, 34),
    'AI74': (f71, 74, 35),
    'AJ74': (f72, 74, 36),
    'AK74': (f73, 74, 37),
    'AL74': (f74, 74, 38),
    'AM74': (f75, 74, 39),
    'AN74': (f76, 74, 40),
    'AO74': (f77, 74, 41),
    'AP74': (f78, 74, 42),
    'AQ74': (f79, 74, 43),
    'AR74': (f80, 74, 44),
    'AS74': (f81, 74, 45),
    'AT74': (f82, 74, 46),
    'AU74': (f83, 74, 47),
    'AV74': (f84, 74, 48),
    'AW74': (f85, 74, 49),
    'AX74': (f86, 74, 50),
    'AY74': (f87, 74, 51),
    'AZ74': (f88, 74, 52),
    'BA74': (f89, 74, 53),
    'BB74': (f90, 74, 54),
    'A75': (f884, 75, 1),
    'B75': (f7, 75, 2),
    'C75': (f885, 75, 3),
    'D75': (f886, 75, 4),
    'E75': (f887, 75, 5),
    'F75': (f888, 75, 6),
    'G75': (f12, 75, 7),
    'H75': (f13, 75, 8),
    'I75': (f14, 75, 9),
    'J75': (f15, 75, 10),
    'K75': (f889, 75, 11),
    'L75': (f890, 75, 12),
    'M75': (f891, 75, 13),
    'Q75': (f19, 75, 17),
    'R75': (f20, 75, 18),
    'T75': (f892, 75, 20),
    'U75': (f22, 75, 21),
    'V75': (f23, 75, 22),
    'W75': (f24, 75, 23),
    'X75': (f25, 75, 24),
    'Y75': (f26, 75, 25),
    'Z75': (f27, 75, 26),
    'AA75': (f893, 75, 27),
    'AB75': (f894, 75, 28),
    'AC75': (f895, 75, 29),
    'AD75': (f896, 75, 30),
    'AE75': (f32, 75, 31),
    'AF75': (f33, 75, 32),
    'AG75': (f69, 75, 33),
    'AH75': (f70, 75, 34),
    'AI75': (f71, 75, 35),
    'AJ75': (f72, 75, 36),
    'AK75': (f73, 75, 37),
    'AL75': (f74, 75, 38),
    'AM75': (f75, 75, 39),
    'AN75': (f76, 75, 40),
    'AO75': (f77, 75, 41),
    'AP75': (f78, 75, 42),
    'AQ75': (f79, 75, 43),
    'AR75': (f80, 75, 44),
    'AS75': (f81, 75, 45),
    'AT75': (f82, 75, 46),
    'AU75': (f83, 75, 47),
    'AV75': (f84, 75, 48),
    'AW75': (f85, 75, 49),
    'AX75': (f86, 75, 50),
    'AY75': (f87, 75, 51),
    'AZ75': (f88, 75, 52),
    'BA75': (f89, 75, 53),
    'BB75': (f90, 75, 54),
    'A76': (f897, 76, 1),
    'B76': (f7, 76, 2),
    'C76': (f898, 76, 3),
    'D76': (f899, 76, 4),
    'E76': (f900, 76, 5),
    'F76': (f901, 76, 6),
    'G76': (f12, 76, 7),
    'H76': (f13, 76, 8),
    'I76': (f14, 76, 9),
    'J76': (f15, 76, 10),
    'K76': (f902, 76, 11),
    'L76': (f903, 76, 12),
    'M76': (f904, 76, 13),
    'Q76': (f19, 76, 17),
    'R76': (f20, 76, 18),
    'T76': (f905, 76, 20),
    'U76': (f22, 76, 21),
    'V76': (f23, 76, 22),
    'W76': (f24, 76, 23),
    'X76': (f25, 76, 24),
    'Y76': (f26, 76, 25),
    'Z76': (f27, 76, 26),
    'AA76': (f906, 76, 27),
    'AB76': (f907, 76, 28),
    'AC76': (f908, 76, 29),
    'AD76': (f909, 76, 30),
    'AE76': (f32, 76, 31),
    'AF76': (f33, 76, 32),
    'AG76': (f69, 76, 33),
    'AH76': (f70, 76, 34),
    'AI76': (f71, 76, 35),
    'AJ76': (f72, 76, 36),
    'AK76': (f73, 76, 37),
    'AL76': (f74, 76, 38),
    'AM76': (f75, 76, 39),
    'AN76': (f76, 76, 40),
    'AO76': (f77, 76, 41),
    'AP76': (f78, 76, 42),
    'AQ76': (f79, 76, 43),
    'AR76': (f80, 76, 44),
    'AS76': (f81, 76, 45),
    'AT76': (f82, 76, 46),
    'AU76': (f83, 76, 47),
    'AV76': (f84, 76, 48),
    'AW76': (f85, 76, 49),
    'AX76': (f86, 76, 50),
    'AY76': (f87, 76, 51),
    'AZ76': (f88, 76, 52),
    'BA76': (f89, 76, 53),
    'BB76': (f90, 76, 54),
    'A77': (f910, 77, 1),
    'B77': (f7, 77, 2),
    'C77': (f911, 77, 3),
    'D77': (f912, 77, 4),
    'E77': (f913, 77, 5),
    'F77': (f914, 77, 6),
    'G77': (f12, 77, 7),
    'H77': (f13, 77, 8),
    'I77': (f14, 77, 9),
    'J77': (f15, 77, 10),
    'K77': (f915, 77, 11),
    'L77': (f916, 77, 12),
    'M77': (f917, 77, 13),
    'Q77': (f19, 77, 17),
    'R77': (f20, 77, 18),
    'T77': (f918, 77, 20),
    'U77': (f22, 77, 21),
    'V77': (f23, 77, 22),
    'W77': (f24, 77, 23),
    'X77': (f25, 77, 24),
    'Y77': (f26, 77, 25),
    'Z77': (f27, 77, 26),
    'AA77': (f919, 77, 27),
    'AB77': (f920, 77, 28),
    'AC77': (f921, 77, 29),
    'AD77': (f922, 77, 30),
    'AE77': (f32, 77, 31),
    'AF77': (f33, 77, 32),
    'AG77': (f69, 77, 33),
    'AH77': (f70, 77, 34),
    'AI77': (f71, 77, 35),
    'AJ77': (f72, 77, 36),
    'AK77': (f73, 77, 37),
    'AL77': (f74, 77, 38),
    'AM77': (f75, 77, 39),
    'AN77': (f76, 77, 40),
    'AO77': (f77, 77, 41),
    'AP77': (f78, 77, 42),
    'AQ77': (f79, 77, 43),
    'AR77': (f80, 77, 44),
    'AS77': (f81, 77, 45),
    'AT77': (f82, 77, 46),
    'AU77': (f83, 77, 47),
    'AV77': (f84, 77, 48),
    'AW77': (f85, 77, 49),
    'AX77': (f86, 77, 50),
    'AY77': (f87, 77, 51),
    'AZ77': (f88, 77, 52),
    'BA77': (f89, 77, 53),
    'BB77': (f90, 77, 54),
    'A78': (f923, 78, 1),
    'B78': (f7, 78, 2),
    'C78': (f924, 78, 3),
    'D78': (f925, 78, 4),
    'E78': (f926, 78, 5),
    'F78': (f927, 78, 6),
    'G78': (f12, 78, 7),
    'H78': (f13, 78, 8),
    'I78': (f14, 78, 9),
    'J78': (f15, 78, 10),
    'K78': (f928, 78, 11),
    'L78': (f929, 78, 12),
    'M78': (f930, 78, 13),
    'Q78': (f19, 78, 17),
    'R78': (f20, 78, 18),
    'T78': (f931, 78, 20),
    'U78': (f22, 78, 21),
    'V78': (f23, 78, 22),
    'W78': (f24, 78, 23),
    'X78': (f25, 78, 24),
    'Y78': (f26, 78, 25),
    'Z78': (f27, 78, 26),
    'AA78': (f932, 78, 27),
    'AB78': (f933, 78, 28),
    'AC78': (f934, 78, 29),
    'AD78': (f935, 78, 30),
    'AE78': (f32, 78, 31),
    'AF78': (f33, 78, 32),
    'AG78': (f69, 78, 33),
    'AH78': (f70, 78, 34),
    'AI78': (f71, 78, 35),
    'AJ78': (f72, 78, 36),
    'AK78': (f73, 78, 37),
    'AL78': (f74, 78, 38),
    'AM78': (f75, 78, 39),
    'AN78': (f76, 78, 40),
    'AO78': (f77, 78, 41),
    'AP78': (f78, 78, 42),
    'AQ78': (f79, 78, 43),
    'AR78': (f80, 78, 44),
    'AS78': (f81, 78, 45),
    'AT78': (f82, 78, 46),
    'AU78': (f83, 78, 47),
    'AV78': (f84, 78, 48),
    'AW78': (f85, 78, 49),
    'AX78': (f86, 78, 50),
    'AY78': (f87, 78, 51),
    'AZ78': (f88, 78, 52),
    'BA78': (f89, 78, 53),
    'BB78': (f90, 78, 54),
    'A79': (f936, 79, 1),
    'B79': (f7, 79, 2),
    'C79': (f937, 79, 3),
    'D79': (f938, 79, 4),
    'E79': (f939, 79, 5),
    'F79': (f940, 79, 6),
    'G79': (f12, 79, 7),
    'H79': (f13, 79, 8),
    'I79': (f14, 79, 9),
    'J79': (f15, 79, 10),
    'K79': (f941, 79, 11),
    'L79': (f942, 79, 12),
    'M79': (f943, 79, 13),
    'Q79': (f19, 79, 17),
    'R79': (f20, 79, 18),
    'T79': (f944, 79, 20),
    'U79': (f22, 79, 21),
    'V79': (f23, 79, 22),
    'W79': (f24, 79, 23),
    'X79': (f25, 79, 24),
    'Y79': (f26, 79, 25),
    'Z79': (f27, 79, 26),
    'AA79': (f945, 79, 27),
    'AB79': (f946, 79, 28),
    'AC79': (f947, 79, 29),
    'AD79': (f948, 79, 30),
    'AE79': (f32, 79, 31),
    'AF79': (f33, 79, 32),
    'AG79': (f69, 79, 33),
    'AH79': (f70, 79, 34),
    'AI79': (f71, 79, 35),
    'AJ79': (f72, 79, 36),
    'AK79': (f73, 79, 37),
    'AL79': (f74, 79, 38),
    'AM79': (f75, 79, 39),
    'AN79': (f76, 79, 40),
    'AO79': (f77, 79, 41),
    'AP79': (f78, 79, 42),
    'AQ79': (f79, 79, 43),
    'AR79': (f80, 79, 44),
    'AS79': (f81, 79, 45),
    'AT79': (f82, 79, 46),
    'AU79': (f83, 79, 47),
    'AV79': (f84, 79, 48),
    'AW79': (f85, 79, 49),
    'AX79': (f86, 79, 50),
    'AY79': (f87, 79, 51),
    'AZ79': (f88, 79, 52),
    'BA79': (f89, 79, 53),
    'BB79': (f90, 79, 54),
    'A80': (f949, 80, 1),
    'B80': (f7, 80, 2),
    'C80': (f950, 80, 3),
    'D80': (f951, 80, 4),
    'E80': (f952, 80, 5),
    'F80': (f953, 80, 6),
    'G80': (f12, 80, 7),
    'H80': (f13, 80, 8),
    'I80': (f14, 80, 9),
    'J80': (f15, 80, 10),
    'K80': (f954, 80, 11),
    'L80': (f955, 80, 12),
    'M80': (f956, 80, 13),
    'Q80': (f19, 80, 17),
    'R80': (f20, 80, 18),
    'T80': (f957, 80, 20),
    'U80': (f22, 80, 21),
    'V80': (f23, 80, 22),
    'W80': (f24, 80, 23),
    'X80': (f25, 80, 24),
    'Y80': (f26, 80, 25),
    'Z80': (f27, 80, 26),
    'AA80': (f958, 80, 27),
    'AB80': (f959, 80, 28),
    'AC80': (f960, 80, 29),
    'AD80': (f961, 80, 30),
    'AE80': (f32, 80, 31),
    'AF80': (f33, 80, 32),
    'AG80': (f69, 80, 33),
    'AH80': (f70, 80, 34),
    'AI80': (f71, 80, 35),
    'AJ80': (f72, 80, 36),
    'AK80': (f73, 80, 37),
    'AL80': (f74, 80, 38),
    'AM80': (f75, 80, 39),
    'AN80': (f76, 80, 40),
    'AO80': (f77, 80, 41),
    'AP80': (f78, 80, 42),
    'AQ80': (f79, 80, 43),
    'AR80': (f80, 80, 44),
    'AS80': (f81, 80, 45),
    'AT80': (f82, 80, 46),
    'AU80': (f83, 80, 47),
    'AV80': (f84, 80, 48),
    'AW80': (f85, 80, 49),
    'AX80': (f86, 80, 50),
    'AY80': (f87, 80, 51),
    'AZ80': (f88, 80, 52),
    'BA80': (f89, 80, 53),
    'BB80': (f90, 80, 54),
    'A81': (f962, 81, 1),
    'B81': (f7, 81, 2),
    'C81': (f963, 81, 3),
    'D81': (f964, 81, 4),
    'E81': (f965, 81, 5),
    'F81': (f966, 81, 6),
    'G81': (f12, 81, 7),
    'H81': (f13, 81, 8),
    'I81': (f14, 81, 9),
    'J81': (f15, 81, 10),
    'K81': (f967, 81, 11),
    'L81': (f968, 81, 12),
    'M81': (f969, 81, 13),
    'Q81': (f19, 81, 17),
    'R81': (f20, 81, 18),
    'T81': (f970, 81, 20),
    'U81': (f22, 81, 21),
    'V81': (f23, 81, 22),
    'W81': (f24, 81, 23),
    'X81': (f25, 81, 24),
    'Y81': (f26, 81, 25),
    'Z81': (f27, 81, 26),
    'AA81': (f971, 81, 27),
    'AB81': (f972, 81, 28),
    'AC81': (f973, 81, 29),
    'AD81': (f974, 81, 30),
    'AE81': (f32, 81, 31),
    'AF81': (f33, 81, 32),
    'AG81': (f69, 81, 33),
    'AH81': (f70, 81, 34),
    'AI81': (f71, 81, 35),
    'AJ81': (f72, 81, 36),
    'AK81': (f73, 81, 37),
    'AL81': (f74, 81, 38),
    'AM81': (f75, 81, 39),
    'AN81': (f76, 81, 40),
    'AO81': (f77, 81, 41),
    'AP81': (f78, 81, 42),
    'AQ81': (f79, 81, 43),
    'AR81': (f80, 81, 44),
    'AS81': (f81, 81, 45),
    'AT81': (f82, 81, 46),
    'AU81': (f83, 81, 47),
    'AV81': (f84, 81, 48),
    'AW81': (f85, 81, 49),
    'AX81': (f86, 81, 50),
    'AY81': (f87, 81, 51),
    'AZ81': (f88, 81, 52),
    'BA81': (f89, 81, 53),
    'BB81': (f90, 81, 54),
    'A82': (f975, 82, 1),
    'B82': (f7, 82, 2),
    'C82': (f976, 82, 3),
    'D82': (f977, 82, 4),
    'E82': (f978, 82, 5),
    'F82': (f979, 82, 6),
    'G82': (f12, 82, 7),
    'H82': (f13, 82, 8),
    'I82': (f14, 82, 9),
    'J82': (f15, 82, 10),
    'K82': (f980, 82, 11),
    'L82': (f981, 82, 12),
    'M82': (f982, 82, 13),
    'Q82': (f19, 82, 17),
    'R82': (f20, 82, 18),
    'T82': (f983, 82, 20),
    'U82': (f22, 82, 21),
    'V82': (f23, 82, 22),
    'W82': (f24, 82, 23),
    'X82': (f25, 82, 24),
    'Y82': (f26, 82, 25),
    'Z82': (f27, 82, 26),
    'AA82': (f984, 82, 27),
    'AB82': (f985, 82, 28),
    'AC82': (f986, 82, 29),
    'AD82': (f987, 82, 30),
    'AE82': (f32, 82, 31),
    'AF82': (f33, 82, 32),
    'AG82': (f69, 82, 33),
    'AH82': (f70, 82, 34),
    'AI82': (f71, 82, 35),
    'AJ82': (f72, 82, 36),
    'AK82': (f73, 82, 37),
    'AL82': (f74, 82, 38),
    'AM82': (f75, 82, 39),
    'AN82': (f76, 82, 40),
    'AO82': (f77, 82, 41),
    'AP82': (f78, 82, 42),
    'AQ82': (f79, 82, 43),
    'AR82': (f80, 82, 44),
    'AS82': (f81, 82, 45),
    'AT82': (f82, 82, 46),
    'AU82': (f83, 82, 47),
    'AV82': (f84, 82, 48),
    'AW82': (f85, 82, 49),
    'AX82': (f86, 82, 50),
    'AY82': (f87, 82, 51),
    'AZ82': (f88, 82, 52),
    'BA82': (f89, 82, 53),
    'BB82': (f90, 82, 54),
    'A83': (f988, 83, 1),
    'B83': (f7, 83, 2),
    'C83': (f989, 83, 3),
    'D83': (f990, 83, 4),
    'E83': (f991, 83, 5),
    'F83': (f992, 83, 6),
    'G83': (f12, 83, 7),
    'H83': (f13, 83, 8),
    'I83': (f14, 83, 9),
    'J83': (f15, 83, 10),
    'K83': (f993, 83, 11),
    'L83': (f994, 83, 12),
    'M83': (f995, 83, 13),
    'Q83': (f19, 83, 17),
    'R83': (f20, 83, 18),
    'T83': (f996, 83, 20),
    'U83': (f22, 83, 21),
    'V83': (f23, 83, 22),
    'W83': (f24, 83, 23),
    'X83': (f25, 83, 24),
    'Y83': (f26, 83, 25),
    'Z83': (f27, 83, 26),
    'AA83': (f997, 83, 27),
    'AB83': (f998, 83, 28),
    'AC83': (f999, 83, 29),
    'AD83': (f1000, 83, 30),
    'AE83': (f32, 83, 31),
    'AF83': (f33, 83, 32),
    'AG83': (f69, 83, 33),
    'AH83': (f70, 83, 34),
    'AI83': (f71, 83, 35),
    'AJ83': (f72, 83, 36),
    'AK83': (f73, 83, 37),
    'AL83': (f74, 83, 38),
    'AM83': (f75, 83, 39),
    'AN83': (f76, 83, 40),
    'AO83': (f77, 83, 41),
    'AP83': (f78, 83, 42),
    'AQ83': (f79, 83, 43),
    'AR83': (f80, 83, 44),
    'AS83': (f81, 83, 45),
    'AT83': (f82, 83, 46),
    'AU83': (f83, 83, 47),
    'AV83': (f84, 83, 48),
    'AW83': (f85, 83, 49),
    'AX83': (f86, 83, 50),
    'AY83': (f87, 83, 51),
    'AZ83': (f88, 83, 52),
    'BA83': (f89, 83, 53),
    'BB83': (f90, 83, 54),
    'A84': (f1001, 84, 1),
    'B84': (f7, 84, 2),
    'C84': (f1002, 84, 3),
    'D84': (f1003, 84, 4),
    'E84': (f1004, 84, 5),
    'F84': (f1005, 84, 6),
    'G84': (f12, 84, 7),
    'H84': (f13, 84, 8),
    'I84': (f14, 84, 9),
    'J84': (f15, 84, 10),
    'K84': (f1006, 84, 11),
    'L84': (f1007, 84, 12),
    'M84': (f1008, 84, 13),
    'Q84': (f19, 84, 17),
    'R84': (f20, 84, 18),
    'T84': (f1009, 84, 20),
    'U84': (f22, 84, 21),
    'V84': (f23, 84, 22),
    'W84': (f24, 84, 23),
    'X84': (f25, 84, 24),
    'Y84': (f26, 84, 25),
    'Z84': (f27, 84, 26),
    'AA84': (f1010, 84, 27),
    'AB84': (f1011, 84, 28),
    'AC84': (f1012, 84, 29),
    'AD84': (f1013, 84, 30),
    'AE84': (f32, 84, 31),
    'AF84': (f33, 84, 32),
    'AG84': (f69, 84, 33),
    'AH84': (f70, 84, 34),
    'AI84': (f71, 84, 35),
    'AJ84': (f72, 84, 36),
    'AK84': (f73, 84, 37),
    'AL84': (f74, 84, 38),
    'AM84': (f75, 84, 39),
    'AN84': (f76, 84, 40),
    'AO84': (f77, 84, 41),
    'AP84': (f78, 84, 42),
    'AQ84': (f79, 84, 43),
    'AR84': (f80, 84, 44),
    'AS84': (f81, 84, 45),
    'AT84': (f82, 84, 46),
    'AU84': (f83, 84, 47),
    'AV84': (f84, 84, 48),
    'AW84': (f85, 84, 49),
    'AX84': (f86, 84, 50),
    'AY84': (f87, 84, 51),
    'AZ84': (f88, 84, 52),
    'BA84': (f89, 84, 53),
    'BB84': (f90, 84, 54),
    'A85': (f1014, 85, 1),
    'B85': (f7, 85, 2),
    'C85': (f1015, 85, 3),
    'D85': (f1016, 85, 4),
    'E85': (f1017, 85, 5),
    'F85': (f1018, 85, 6),
    'G85': (f12, 85, 7),
    'H85': (f13, 85, 8),
    'I85': (f14, 85, 9),
    'J85': (f15, 85, 10),
    'K85': (f1019, 85, 11),
    'L85': (f1020, 85, 12),
    'M85': (f1021, 85, 13),
    'Q85': (f19, 85, 17),
    'R85': (f20, 85, 18),
    'T85': (f1022, 85, 20),
    'U85': (f22, 85, 21),
    'V85': (f23, 85, 22),
    'W85': (f24, 85, 23),
    'X85': (f25, 85, 24),
    'Y85': (f26, 85, 25),
    'Z85': (f27, 85, 26),
    'AA85': (f1023, 85, 27),
    'AB85': (f1024, 85, 28),
    'AC85': (f1025, 85, 29),
    'AD85': (f1026, 85, 30),
    'AE85': (f32, 85, 31),
    'AF85': (f33, 85, 32),
    'AG85': (f69, 85, 33),
    'AH85': (f70, 85, 34),
    'AI85': (f71, 85, 35),
    'AJ85': (f72, 85, 36),
    'AK85': (f73, 85, 37),
    'AL85': (f74, 85, 38),
    'AM85': (f75, 85, 39),
    'AN85': (f76, 85, 40),
    'AO85': (f77, 85, 41),
    'AP85': (f78, 85, 42),
    'AQ85': (f79, 85, 43),
    'AR85': (f80, 85, 44),
    'AS85': (f81, 85, 45),
    'AT85': (f82, 85, 46),
    'AU85': (f83, 85, 47),
    'AV85': (f84, 85, 48),
    'AW85': (f85, 85, 49),
    'AX85': (f86, 85, 50),
    'AY85': (f87, 85, 51),
    'AZ85': (f88, 85, 52),
    'BA85': (f89, 85, 53),
    'BB85': (f90, 85, 54),
    'A86': (f1027, 86, 1),
    'B86': (f7, 86, 2),
    'C86': (f1028, 86, 3),
    'D86': (f1029, 86, 4),
    'E86': (f1030, 86, 5),
    'F86': (f1031, 86, 6),
    'G86': (f12, 86, 7),
    'H86': (f13, 86, 8),
    'I86': (f14, 86, 9),
    'J86': (f15, 86, 10),
    'K86': (f1032, 86, 11),
    'L86': (f1033, 86, 12),
    'M86': (f1034, 86, 13),
    'Q86': (f19, 86, 17),
    'R86': (f20, 86, 18),
    'T86': (f1035, 86, 20),
    'U86': (f22, 86, 21),
    'V86': (f23, 86, 22),
    'W86': (f24, 86, 23),
    'X86': (f25, 86, 24),
    'Y86': (f26, 86, 25),
    'Z86': (f27, 86, 26),
    'AA86': (f1036, 86, 27),
    'AB86': (f1037, 86, 28),
    'AC86': (f1038, 86, 29),
    'AD86': (f1039, 86, 30),
    'AE86': (f32, 86, 31),
    'AF86': (f33, 86, 32),
    'AG86': (f69, 86, 33),
    'AH86': (f70, 86, 34),
    'AI86': (f71, 86, 35),
    'AJ86': (f72, 86, 36),
    'AK86': (f73, 86, 37),
    'AL86': (f74, 86, 38),
    'AM86': (f75, 86, 39),
    'AN86': (f76, 86, 40),
    'AO86': (f77, 86, 41),
    'AP86': (f78, 86, 42),
    'AQ86': (f79, 86, 43),
    'AR86': (f80, 86, 44),
    'AS86': (f81, 86, 45),
    'AT86': (f82, 86, 46),
    'AU86': (f83, 86, 47),
    'AV86': (f84, 86, 48),
    'AW86': (f85, 86, 49),
    'AX86': (f86, 86, 50),
    'AY86': (f87, 86, 51),
    'AZ86': (f88, 86, 52),
    'BA86': (f89, 86, 53),
    'BB86': (f90, 86, 54),
    'A87': (f1040, 87, 1),
    'B87': (f7, 87, 2),
    'C87': (f1041, 87, 3),
    'D87': (f1042, 87, 4),
    'E87': (f1043, 87, 5),
    'F87': (f1044, 87, 6),
    'G87': (f12, 87, 7),
    'H87': (f13, 87, 8),
    'I87': (f14, 87, 9),
    'J87': (f15, 87, 10),
    'K87': (f1045, 87, 11),
    'L87': (f1046, 87, 12),
    'M87': (f1047, 87, 13),
    'Q87': (f19, 87, 17),
    'R87': (f20, 87, 18),
    'T87': (f1048, 87, 20),
    'U87': (f22, 87, 21),
    'V87': (f23, 87, 22),
    'W87': (f24, 87, 23),
    'X87': (f25, 87, 24),
    'Y87': (f26, 87, 25),
    'Z87': (f27, 87, 26),
    'AA87': (f1049, 87, 27),
    'AB87': (f1050, 87, 28),
    'AC87': (f1051, 87, 29),
    'AD87': (f1052, 87, 30),
    'AE87': (f32, 87, 31),
    'AF87': (f33, 87, 32),
    'AG87': (f69, 87, 33),
    'AH87': (f70, 87, 34),
    'AI87': (f71, 87, 35),
    'AJ87': (f72, 87, 36),
    'AK87': (f73, 87, 37),
    'AL87': (f74, 87, 38),
    'AM87': (f75, 87, 39),
    'AN87': (f76, 87, 40),
    'AO87': (f77, 87, 41),
    'AP87': (f78, 87, 42),
    'AQ87': (f79, 87, 43),
    'AR87': (f80, 87, 44),
    'AS87': (f81, 87, 45),
    'AT87': (f82, 87, 46),
    'AU87': (f83, 87, 47),
    'AV87': (f84, 87, 48),
    'AW87': (f85, 87, 49),
    'AX87': (f86, 87, 50),
    'AY87': (f87, 87, 51),
    'AZ87': (f88, 87, 52),
    'BA87': (f89, 87, 53),
    'BB87': (f90, 87, 54),
    'A88': (f1053, 88, 1),
    'B88': (f7, 88, 2),
    'C88': (f1054, 88, 3),
    'D88': (f1055, 88, 4),
    'E88': (f1056, 88, 5),
    'F88': (f1057, 88, 6),
    'G88': (f12, 88, 7),
    'H88': (f13, 88, 8),
    'I88': (f14, 88, 9),
    'J88': (f15, 88, 10),
    'K88': (f1058, 88, 11),
    'L88': (f1059, 88, 12),
    'M88': (f1060, 88, 13),
    'Q88': (f19, 88, 17),
    'R88': (f20, 88, 18),
    'T88': (f1061, 88, 20),
    'U88': (f22, 88, 21),
    'V88': (f23, 88, 22),
    'W88': (f24, 88, 23),
    'X88': (f25, 88, 24),
    'Y88': (f26, 88, 25),
    'Z88': (f27, 88, 26),
    'AA88': (f1062, 88, 27),
    'AB88': (f1063, 88, 28),
    'AC88': (f1064, 88, 29),
    'AD88': (f1065, 88, 30),
    'AE88': (f32, 88, 31),
    'AF88': (f33, 88, 32),
    'AG88': (f69, 88, 33),
    'AH88': (f70, 88, 34),
    'AI88': (f71, 88, 35),
    'AJ88': (f72, 88, 36),
    'AK88': (f73, 88, 37),
    'AL88': (f74, 88, 38),
    'AM88': (f75, 88, 39),
    'AN88': (f76, 88, 40),
    'AO88': (f77, 88, 41),
    'AP88': (f78, 88, 42),
    'AQ88': (f79, 88, 43),
    'AR88': (f80, 88, 44),
    'AS88': (f81, 88, 45),
    'AT88': (f82, 88, 46),
    'AU88': (f83, 88, 47),
    'AV88': (f84, 88, 48),
    'AW88': (f85, 88, 49),
    'AX88': (f86, 88, 50),
    'AY88': (f87, 88, 51),
    'AZ88': (f88, 88, 52),
    'BA88': (f89, 88, 53),
    'BB88': (f90, 88, 54),
    'A89': (f1066, 89, 1),
    'B89': (f7, 89, 2),
    'C89': (f1067, 89, 3),
    'D89': (f1068, 89, 4),
    'E89': (f1069, 89, 5),
    'F89': (f1070, 89, 6),
    'G89': (f12, 89, 7),
    'H89': (f13, 89, 8),
    'I89': (f14, 89, 9),
    'J89': (f15, 89, 10),
    'K89': (f1071, 89, 11),
    'L89': (f1072, 89, 12),
    'M89': (f1073, 89, 13),
    'Q89': (f19, 89, 17),
    'R89': (f20, 89, 18),
    'T89': (f1074, 89, 20),
    'U89': (f22, 89, 21),
    'V89': (f23, 89, 22),
    'W89': (f24, 89, 23),
    'X89': (f25, 89, 24),
    'Y89': (f26, 89, 25),
    'Z89': (f27, 89, 26),
    'AA89': (f1075, 89, 27),
    'AB89': (f1076, 89, 28),
    'AC89': (f1077, 89, 29),
    'AD89': (f1078, 89, 30),
    'AE89': (f32, 89, 31),
    'AF89': (f33, 89, 32),
    'AG89': (f69, 89, 33),
    'AH89': (f70, 89, 34),
    'AI89': (f71, 89, 35),
    'AJ89': (f72, 89, 36),
    'AK89': (f73, 89, 37),
    'AL89': (f74, 89, 38),
    'AM89': (f75, 89, 39),
    'AN89': (f76, 89, 40),
    'AO89': (f77, 89, 41),
    'AP89': (f78, 89, 42),
    'AQ89': (f79, 89, 43),
    'AR89': (f80, 89, 44),
    'AS89': (f81, 89, 45),
    'AT89': (f82, 89, 46),
    'AU89': (f83, 89, 47),
    'AV89': (f84, 89, 48),
    'AW89': (f85, 89, 49),
    'AX89': (f86, 89, 50),
    'AY89': (f87, 89, 51),
    'AZ89': (f88, 89, 52),
    'BA89': (f89, 89, 53),
    'BB89': (f90, 89, 54),
    'A90': (f1079, 90, 1),
    'B90': (f7, 90, 2),
    'C90': (f1080, 90, 3),
    'D90': (f1081, 90, 4),
    'E90': (f1082, 90, 5),
    'F90': (f1083, 90, 6),
    'G90': (f12, 90, 7),
    'H90': (f13, 90, 8),
    'I90': (f14, 90, 9),
    'J90': (f15, 90, 10),
    'K90': (f1084, 90, 11),
    'L90': (f1085, 90, 12),
    'M90': (f1086, 90, 13),
    'Q90': (f19, 90, 17),
    'R90': (f20, 90, 18),
    'T90': (f1087, 90, 20),
    'U90': (f22, 90, 21),
    'V90': (f23, 90, 22),
    'W90': (f24, 90, 23),
    'X90': (f25, 90, 24),
    'Y90': (f26, 90, 25),
    'Z90': (f27, 90, 26),
    'AA90': (f1088, 90, 27),
    'AB90': (f1089, 90, 28),
    'AC90': (f1090, 90, 29),
    'AD90': (f1091, 90, 30),
    'AE90': (f32, 90, 31),
    'AF90': (f33, 90, 32),
    'AG90': (f69, 90, 33),
    'AH90': (f70, 90, 34),
    'AI90': (f71, 90, 35),
    'AJ90': (f72, 90, 36),
    'AK90': (f73, 90, 37),
    'AL90': (f74, 90, 38),
    'AM90': (f75, 90, 39),
    'AN90': (f76, 90, 40),
    'AO90': (f77, 90, 41),
    'AP90': (f78, 90, 42),
    'AQ90': (f79, 90, 43),
    'AR90': (f80, 90, 44),
    'AS90': (f81, 90, 45),
    'AT90': (f82, 90, 46),
    'AU90': (f83, 90, 47),
    'AV90': (f84, 90, 48),
    'AW90': (f85, 90, 49),
    'AX90': (f86, 90, 50),
    'AY90': (f87, 90, 51),
    'AZ90': (f88, 90, 52),
    'BA90': (f89, 90, 53),
    'BB90': (f90, 90, 54),
    'A91': (f1092, 91, 1),
    'B91': (f7, 91, 2),
    'C91': (f1093, 91, 3),
    'D91': (f1094, 91, 4),
    'E91': (f1095, 91, 5),
    'F91': (f1096, 91, 6),
    'G91': (f12, 91, 7),
    'H91': (f13, 91, 8),
    'I91': (f14, 91, 9),
    'J91': (f15, 91, 10),
    'K91': (f1097, 91, 11),
    'L91': (f1098, 91, 12),
    'M91': (f1099, 91, 13),
    'Q91': (f19, 91, 17),
    'R91': (f20, 91, 18),
    'T91': (f1100, 91, 20),
    'U91': (f22, 91, 21),
    'V91': (f23, 91, 22),
    'W91': (f24, 91, 23),
    'X91': (f25, 91, 24),
    'Y91': (f26, 91, 25),
    'Z91': (f27, 91, 26),
    'AA91': (f1101, 91, 27),
    'AB91': (f1102, 91, 28),
    'AC91': (f1103, 91, 29),
    'AD91': (f1104, 91, 30),
    'AE91': (f32, 91, 31),
    'AF91': (f33, 91, 32),
    'AG91': (f69, 91, 33),
    'AH91': (f70, 91, 34),
    'AI91': (f71, 91, 35),
    'AJ91': (f72, 91, 36),
    'AK91': (f73, 91, 37),
    'AL91': (f74, 91, 38),
    'AM91': (f75, 91, 39),
    'AN91': (f76, 91, 40),
    'AO91': (f77, 91, 41),
    'AP91': (f78, 91, 42),
    'AQ91': (f79, 91, 43),
    'AR91': (f80, 91, 44),
    'AS91': (f81, 91, 45),
    'AT91': (f82, 91, 46),
    'AU91': (f83, 91, 47),
    'AV91': (f84, 91, 48),
    'AW91': (f85, 91, 49),
    'AX91': (f86, 91, 50),
    'AY91': (f87, 91, 51),
    'AZ91': (f88, 91, 52),
    'BA91': (f89, 91, 53),
    'BB91': (f90, 91, 54),
    'A92': (f1105, 92, 1),
    'B92': (f7, 92, 2),
    'C92': (f1106, 92, 3),
    'D92': (f1107, 92, 4),
    'E92': (f1108, 92, 5),
    'F92': (f1109, 92, 6),
    'G92': (f12, 92, 7),
    'H92': (f13, 92, 8),
    'I92': (f14, 92, 9),
    'J92': (f15, 92, 10),
    'K92': (f1110, 92, 11),
    'L92': (f1111, 92, 12),
    'M92': (f1112, 92, 13),
    'Q92': (f19, 92, 17),
    'R92': (f20, 92, 18),
    'T92': (f1113, 92, 20),
    'U92': (f22, 92, 21),
    'V92': (f23, 92, 22),
    'W92': (f24, 92, 23),
    'X92': (f25, 92, 24),
    'Y92': (f26, 92, 25),
    'Z92': (f27, 92, 26),
    'AA92': (f1114, 92, 27),
    'AB92': (f1115, 92, 28),
    'AC92': (f1116, 92, 29),
    'AD92': (f1117, 92, 30),
    'AE92': (f32, 92, 31),
    'AF92': (f33, 92, 32),
    'AG92': (f69, 92, 33),
    'AH92': (f70, 92, 34),
    'AI92': (f71, 92, 35),
    'AJ92': (f72, 92, 36),
    'AK92': (f73, 92, 37),
    'AL92': (f74, 92, 38),
    'AM92': (f75, 92, 39),
    'AN92': (f76, 92, 40),
    'AO92': (f77, 92, 41),
    'AP92': (f78, 92, 42),
    'AQ92': (f79, 92, 43),
    'AR92': (f80, 92, 44),
    'AS92': (f81, 92, 45),
    'AT92': (f82, 92, 46),
    'AU92': (f83, 92, 47),
    'AV92': (f84, 92, 48),
    'AW92': (f85, 92, 49),
    'AX92': (f86, 92, 50),
    'AY92': (f87, 92, 51),
    'AZ92': (f88, 92, 52),
    'BA92': (f89, 92, 53),
    'BB92': (f90, 92, 54),
    'A93': (f1118, 93, 1),
    'B93': (f7, 93, 2),
    'C93': (f1119, 93, 3),
    'D93': (f1120, 93, 4),
    'E93': (f1121, 93, 5),
    'F93': (f1122, 93, 6),
    'G93': (f12, 93, 7),
    'H93': (f13, 93, 8),
    'I93': (f14, 93, 9),
    'J93': (f15, 93, 10),
    'K93': (f1123, 93, 11),
    'L93': (f1124, 93, 12),
    'M93': (f1125, 93, 13),
    'Q93': (f19, 93, 17),
    'R93': (f20, 93, 18),
    'T93': (f1126, 93, 20),
    'U93': (f22, 93, 21),
    'V93': (f23, 93, 22),
    'W93': (f24, 93, 23),
    'X93': (f25, 93, 24),
    'Y93': (f26, 93, 25),
    'Z93': (f27, 93, 26),
    'AA93': (f1127, 93, 27),
    'AB93': (f1128, 93, 28),
    'AC93': (f1129, 93, 29),
    'AD93': (f1130, 93, 30),
    'AE93': (f32, 93, 31),
    'AF93': (f33, 93, 32),
    'AG93': (f69, 93, 33),
    'AH93': (f70, 93, 34),
    'AI93': (f71, 93, 35),
    'AJ93': (f72, 93, 36),
    'AK93': (f73, 93, 37),
    'AL93': (f74, 93, 38),
    'AM93': (f75, 93, 39),
    'AN93': (f76, 93, 40),
    'AO93': (f77, 93, 41),
    'AP93': (f78, 93, 42),
    'AQ93': (f79, 93, 43),
    'AR93': (f80, 93, 44),
    'AS93': (f81, 93, 45),
    'AT93': (f82, 93, 46),
    'AU93': (f83, 93, 47),
    'AV93': (f84, 93, 48),
    'AW93': (f85, 93, 49),
    'AX93': (f86, 93, 50),
    'AY93': (f87, 93, 51),
    'AZ93': (f88, 93, 52),
    'BA93': (f89, 93, 53),
    'BB93': (f90, 93, 54),
    'A94': (f1131, 94, 1),
    'B94': (f7, 94, 2),
    'C94': (f1132, 94, 3),
    'D94': (f1133, 94, 4),
    'E94': (f1134, 94, 5),
    'F94': (f1135, 94, 6),
    'G94': (f12, 94, 7),
    'H94': (f13, 94, 8),
    'I94': (f14, 94, 9),
    'J94': (f15, 94, 10),
    'K94': (f1136, 94, 11),
    'L94': (f1137, 94, 12),
    'M94': (f1138, 94, 13),
    'Q94': (f19, 94, 17),
    'R94': (f20, 94, 18),
    'T94': (f1139, 94, 20),
    'U94': (f22, 94, 21),
    'V94': (f23, 94, 22),
    'W94': (f24, 94, 23),
    'X94': (f25, 94, 24),
    'Y94': (f26, 94, 25),
    'Z94': (f27, 94, 26),
    'AA94': (f1140, 94, 27),
    'AB94': (f1141, 94, 28),
    'AC94': (f1142, 94, 29),
    'AD94': (f1143, 94, 30),
    'AE94': (f32, 94, 31),
    'AF94': (f33, 94, 32),
    'AG94': (f69, 94, 33),
    'AH94': (f70, 94, 34),
    'AI94': (f71, 94, 35),
    'AJ94': (f72, 94, 36),
    'AK94': (f73, 94, 37),
    'AL94': (f74, 94, 38),
    'AM94': (f75, 94, 39),
    'AN94': (f76, 94, 40),
    'AO94': (f77, 94, 41),
    'AP94': (f78, 94, 42),
    'AQ94': (f79, 94, 43),
    'AR94': (f80, 94, 44),
    'AS94': (f81, 94, 45),
    'AT94': (f82, 94, 46),
    'AU94': (f83, 94, 47),
    'AV94': (f84, 94, 48),
    'AW94': (f85, 94, 49),
    'AX94': (f86, 94, 50),
    'AY94': (f87, 94, 51),
    'AZ94': (f88, 94, 52),
    'BA94': (f89, 94, 53),
    'BB94': (f90, 94, 54),
    'A95': (f1144, 95, 1),
    'B95': (f7, 95, 2),
    'C95': (f1145, 95, 3),
    'D95': (f1146, 95, 4),
    'E95': (f1147, 95, 5),
    'F95': (f1148, 95, 6),
    'G95': (f12, 95, 7),
    'H95': (f13, 95, 8),
    'I95': (f14, 95, 9),
    'J95': (f15, 95, 10),
    'K95': (f1149, 95, 11),
    'L95': (f1150, 95, 12),
    'M95': (f1151, 95, 13),
    'Q95': (f19, 95, 17),
    'R95': (f20, 95, 18),
    'T95': (f1152, 95, 20),
    'U95': (f22, 95, 21),
    'V95': (f23, 95, 22),
    'W95': (f24, 95, 23),
    'X95': (f25, 95, 24),
    'Y95': (f26, 95, 25),
    'Z95': (f27, 95, 26),
    'AA95': (f1153, 95, 27),
    'AB95': (f1154, 95, 28),
    'AC95': (f1155, 95, 29),
    'AD95': (f1156, 95, 30),
    'AE95': (f32, 95, 31),
    'AF95': (f33, 95, 32),
    'AG95': (f69, 95, 33),
    'AH95': (f70, 95, 34),
    'AI95': (f71, 95, 35),
    'AJ95': (f72, 95, 36),
    'AK95': (f73, 95, 37),
    'AL95': (f74, 95, 38),
    'AM95': (f75, 95, 39),
    'AN95': (f76, 95, 40),
    'AO95': (f77, 95, 41),
    'AP95': (f78, 95, 42),
    'AQ95': (f79, 95, 43),
    'AR95': (f80, 95, 44),
    'AS95': (f81, 95, 45),
    'AT95': (f82, 95, 46),
    'AU95': (f83, 95, 47),
    'AV95': (f84, 95, 48),
    'AW95': (f85, 95, 49),
    'AX95': (f86, 95, 50),
    'AY95': (f87, 95, 51),
    'AZ95': (f88, 95, 52),
    'BA95': (f89, 95, 53),
    'BB95': (f90, 95, 54),
    'A96': (f1157, 96, 1),
    'B96': (f7, 96, 2),
    'C96': (f1158, 96, 3),
    'D96': (f1159, 96, 4),
    'E96': (f1160, 96, 5),
    'F96': (f1161, 96, 6),
    'G96': (f12, 96, 7),
    'H96': (f13, 96, 8),
    'I96': (f14, 96, 9),
    'J96': (f15, 96, 10),
    'K96': (f1162, 96, 11),
    'L96': (f1163, 96, 12),
    'M96': (f1164, 96, 13),
    'Q96': (f19, 96, 17),
    'R96': (f20, 96, 18),
    'T96': (f1165, 96, 20),
    'U96': (f22, 96, 21),
    'V96': (f23, 96, 22),
    'W96': (f24, 96, 23),
    'X96': (f25, 96, 24),
    'Y96': (f26, 96, 25),
    'Z96': (f27, 96, 26),
    'AA96': (f1166, 96, 27),
    'AB96': (f1167, 96, 28),
    'AC96': (f1168, 96, 29),
    'AD96': (f1169, 96, 30),
    'AE96': (f32, 96, 31),
    'AF96': (f33, 96, 32),
    'AG96': (f69, 96, 33),
    'AH96': (f70, 96, 34),
    'AI96': (f71, 96, 35),
    'AJ96': (f72, 96, 36),
    'AK96': (f73, 96, 37),
    'AL96': (f74, 96, 38),
    'AM96': (f75, 96, 39),
    'AN96': (f76, 96, 40),
    'AO96': (f77, 96, 41),
    'AP96': (f78, 96, 42),
    'AQ96': (f79, 96, 43),
    'AR96': (f80, 96, 44),
    'AS96': (f81, 96, 45),
    'AT96': (f82, 96, 46),
    'AU96': (f83, 96, 47),
    'AV96': (f84, 96, 48),
    'AW96': (f85, 96, 49),
    'AX96': (f86, 96, 50),
    'AY96': (f87, 96, 51),
    'AZ96': (f88, 96, 52),
    'BA96': (f89, 96, 53),
    'BB96': (f90, 96, 54),
    'A97': (f1170, 97, 1),
    'B97': (f7, 97, 2),
    'C97': (f1171, 97, 3),
    'D97': (f1172, 97, 4),
    'E97': (f1173, 97, 5),
    'F97': (f1174, 97, 6),
    'G97': (f12, 97, 7),
    'H97': (f13, 97, 8),
    'I97': (f14, 97, 9),
    'J97': (f15, 97, 10),
    'K97': (f1175, 97, 11),
    'L97': (f1176, 97, 12),
    'M97': (f1177, 97, 13),
    'Q97': (f19, 97, 17),
    'R97': (f20, 97, 18),
    'T97': (f1178, 97, 20),
    'U97': (f22, 97, 21),
    'V97': (f23, 97, 22),
    'W97': (f24, 97, 23),
    'X97': (f25, 97, 24),
    'Y97': (f26, 97, 25),
    'Z97': (f27, 97, 26),
    'AA97': (f1179, 97, 27),
    'AB97': (f1180, 97, 28),
    'AC97': (f1181, 97, 29),
    'AD97': (f1182, 97, 30),
    'AE97': (f32, 97, 31),
    'AF97': (f33, 97, 32),
    'AG97': (f69, 97, 33),
    'AH97': (f70, 97, 34),
    'AI97': (f71, 97, 35),
    'AJ97': (f72, 97, 36),
    'AK97': (f73, 97, 37),
    'AL97': (f74, 97, 38),
    'AM97': (f75, 97, 39),
    'AN97': (f76, 97, 40),
    'AO97': (f77, 97, 41),
    'AP97': (f78, 97, 42),
    'AQ97': (f79, 97, 43),
    'AR97': (f80, 97, 44),
    'AS97': (f81, 97, 45),
    'AT97': (f82, 97, 46),
    'AU97': (f83, 97, 47),
    'AV97': (f84, 97, 48),
    'AW97': (f85, 97, 49),
    'AX97': (f86, 97, 50),
    'AY97': (f87, 97, 51),
    'AZ97': (f88, 97, 52),
    'BA97': (f89, 97, 53),
    'BB97': (f90, 97, 54),
    'A98': (f1183, 98, 1),
    'B98': (f7, 98, 2),
    'C98': (f1184, 98, 3),
    'D98': (f1185, 98, 4),
    'E98': (f1186, 98, 5),
    'F98': (f1187, 98, 6),
    'G98': (f12, 98, 7),
    'H98': (f13, 98, 8),
    'I98': (f14, 98, 9),
    'J98': (f15, 98, 10),
    'K98': (f1188, 98, 11),
    'L98': (f1189, 98, 12),
    'M98': (f1190, 98, 13),
    'Q98': (f19, 98, 17),
    'R98': (f20, 98, 18),
    'T98': (f1191, 98, 20),
    'U98': (f22, 98, 21),
    'V98': (f23, 98, 22),
    'W98': (f24, 98, 23),
    'X98': (f25, 98, 24),
    'Y98': (f26, 98, 25),
    'Z98': (f27, 98, 26),
    'AA98': (f1192, 98, 27),
    'AB98': (f1193, 98, 28),
    'AC98': (f1194, 98, 29),
    'AD98': (f1195, 98, 30),
    'AE98': (f32, 98, 31),
    'AF98': (f33, 98, 32),
    'AG98': (f69, 98, 33),
    'AH98': (f70, 98, 34),
    'AI98': (f71, 98, 35),
    'AJ98': (f72, 98, 36),
    'AK98': (f73, 98, 37),
    'AL98': (f74, 98, 38),
    'AM98': (f75, 98, 39),
    'AN98': (f76, 98, 40),
    'AO98': (f77, 98, 41),
    'AP98': (f78, 98, 42),
    'AQ98': (f79, 98, 43),
    'AR98': (f80, 98, 44),
    'AS98': (f81, 98, 45),
    'AT98': (f82, 98, 46),
    'AU98': (f83, 98, 47),
    'AV98': (f84, 98, 48),
    'AW98': (f85, 98, 49),
    'AX98': (f86, 98, 50),
    'AY98': (f87, 98, 51),
    'AZ98': (f88, 98, 52),
    'BA98': (f89, 98, 53),
    'BB98': (f90, 98, 54),
    'A99': (f1196, 99, 1),
    'B99': (f7, 99, 2),
    'C99': (f1197, 99, 3),
    'D99': (f1198, 99, 4),
    'E99': (f1199, 99, 5),
    'F99': (f1200, 99, 6),
    'G99': (f12, 99, 7),
    'H99': (f13, 99, 8),
    'I99': (f14, 99, 9),
    'J99': (f15, 99, 10),
    'K99': (f1201, 99, 11),
    'L99': (f1202, 99, 12),
    'M99': (f1203, 99, 13),
    'Q99': (f19, 99, 17),
    'R99': (f20, 99, 18),
    'T99': (f1204, 99, 20),
    'U99': (f22, 99, 21),
    'V99': (f23, 99, 22),
    'W99': (f24, 99, 23),
    'X99': (f25, 99, 24),
    'Y99': (f26, 99, 25),
    'Z99': (f27, 99, 26),
    'AA99': (f1205, 99, 27),
    'AB99': (f1206, 99, 28),
    'AC99': (f1207, 99, 29),
    'AD99': (f1208, 99, 30),
    'AE99': (f32, 99, 31),
    'AF99': (f33, 99, 32),
    'AG99': (f69, 99, 33),
    'AH99': (f70, 99, 34),
    'AI99': (f71, 99, 35),
    'AJ99': (f72, 99, 36),
    'AK99': (f73, 99, 37),
    'AL99': (f74, 99, 38),
    'AM99': (f75, 99, 39),
    'AN99': (f76, 99, 40),
    'AO99': (f77, 99, 41),
    'AP99': (f78, 99, 42),
    'AQ99': (f79, 99, 43),
    'AR99': (f80, 99, 44),
    'AS99': (f81, 99, 45),
    'AT99': (f82, 99, 46),
    'AU99': (f83, 99, 47),
    'AV99': (f84, 99, 48),
    'AW99': (f85, 99, 49),
    'AX99': (f86, 99, 50),
    'AY99': (f87, 99, 51),
    'AZ99': (f88, 99, 52),
    'BA99': (f89, 99, 53),
    'BB99': (f90, 99, 54),
    'A100': (f1209, 100, 1),
    'B100': (f7, 100, 2),
    'C100': (f1210, 100, 3),
    'D100': (f1211, 100, 4),
    'E100': (f1212, 100, 5),
    'F100': (f1213, 100, 6),
    'G100': (f12, 100, 7),
    'H100': (f13, 100, 8),
    'I100': (f14, 100, 9),
    'J100': (f15, 100, 10),
    'K100': (f1214, 100, 11),
    'L100': (f1215, 100, 12),
    'M100': (f1216, 100, 13),
    'Q100': (f19, 100, 17),
    'R100': (f20, 100, 18),
    'T100': (f1217, 100, 20),
    'U100': (f22, 100, 21),
    'V100': (f23, 100, 22),
    'W100': (f24, 100, 23),
    'X100': (f25, 100, 24),
    'Y100': (f26, 100, 25),
    'Z100': (f27, 100, 26),
    'AA100': (f1218, 100, 27),
    'AB100': (f1219, 100, 28),
    'AC100': (f1220, 100, 29),
    'AD100': (f1221, 100, 30),
    'AE100': (f32, 100, 31),
    'AF100': (f33, 100, 32),
    'AG100': (f69, 100, 33),
    'AH100': (f70, 100, 34),
    'AI100': (f71, 100, 35),
    'AJ100': (f72, 100, 36),
    'AK100': (f73, 100, 37),
    'AL100': (f74, 100, 38),
    'AM100': (f75, 100, 39),
    'AN100': (f76, 100, 40),
    'AO100': (f77, 100, 41),
    'AP100': (f78, 100, 42),
    'AQ100': (f79, 100, 43),
    'AR100': (f80, 100, 44),
    'AS100': (f81, 100, 45),
    'AT100': (f82, 100, 46),
    'AU100': (f83, 100, 47),
    'AV100': (f84, 100, 48),
    'AW100': (f85, 100, 49),
    'AX100': (f86, 100, 50),
    'AY100': (f87, 100, 51),
    'AZ100': (f88, 100, 52),
    'BA100': (f89, 100, 53),
    'BB100': (f90, 100, 54),
    'A101': (f1222, 101, 1),
    'B101': (f7, 101, 2),
    'C101': (f1223, 101, 3),
    'D101': (f1224, 101, 4),
    'E101': (f1225, 101, 5),
    'F101': (f1226, 101, 6),
    'G101': (f12, 101, 7),
    'H101': (f13, 101, 8),
    'I101': (f14, 101, 9),
    'J101': (f15, 101, 10),
    'K101': (f1227, 101, 11),
    'L101': (f1228, 101, 12),
    'M101': (f1229, 101, 13),
    'Q101': (f19, 101, 17),
    'R101': (f20, 101, 18),
    'T101': (f1230, 101, 20),
    'U101': (f22, 101, 21),
    'V101': (f23, 101, 22),
    'W101': (f24, 101, 23),
    'X101': (f25, 101, 24),
    'Y101': (f26, 101, 25),
    'Z101': (f27, 101, 26),
    'AA101': (f1231, 101, 27),
    'AB101': (f1232, 101, 28),
    'AC101': (f1233, 101, 29),
    'AD101': (f1234, 101, 30),
    'AE101': (f32, 101, 31),
    'AF101': (f33, 101, 32),
    'AG101': (f69, 101, 33),
    'AH101': (f70, 101, 34),
    'AI101': (f71, 101, 35),
    'AJ101': (f72, 101, 36),
    'AK101': (f73, 101, 37),
    'AL101': (f74, 101, 38),
    'AM101': (f75, 101, 39),
    'AN101': (f76, 101, 40),
    'AO101': (f77, 101, 41),
    'AP101': (f78, 101, 42),
    'AQ101': (f79, 101, 43),
    'AR101': (f80, 101, 44),
    'AS101': (f81, 101, 45),
    'AT101': (f82, 101, 46),
    'AU101': (f83, 101, 47),
    'AV101': (f84, 101, 48),
    'AW101': (f85, 101, 49),
    'AX101': (f86, 101, 50),
    'AY101': (f87, 101, 51),
    'AZ101': (f88, 101, 52),
    'BA101': (f89, 101, 53),
    'BB101': (f90, 101, 54),
    'A102': (f1235, 102, 1),
    'B102': (f7, 102, 2),
    'C102': (f1236, 102, 3),
    'D102': (f1237, 102, 4),
    'E102': (f1238, 102, 5),
    'F102': (f1239, 102, 6),
    'G102': (f12, 102, 7),
    'H102': (f13, 102, 8),
    'I102': (f14, 102, 9),
    'J102': (f15, 102, 10),
    'K102': (f1240, 102, 11),
    'L102': (f1241, 102, 12),
    'M102': (f1242, 102, 13),
    'Q102': (f19, 102, 17),
    'R102': (f20, 102, 18),
    'T102': (f1243, 102, 20),
    'U102': (f22, 102, 21),
    'V102': (f23, 102, 22),
    'W102': (f24, 102, 23),
    'X102': (f25, 102, 24),
    'Y102': (f26, 102, 25),
    'Z102': (f27, 102, 26),
    'AA102': (f1244, 102, 27),
    'AB102': (f1245, 102, 28),
    'AC102': (f1246, 102, 29),
    'AD102': (f1247, 102, 30),
    'AE102': (f32, 102, 31),
    'AF102': (f33, 102, 32),
    'AG102': (f69, 102, 33),
    'AH102': (f70, 102, 34),
    'AI102': (f71, 102, 35),
    'AJ102': (f72, 102, 36),
    'AK102': (f73, 102, 37),
    'AL102': (f74, 102, 38),
    'AM102': (f75, 102, 39),
    'AN102': (f76, 102, 40),
    'AO102': (f77, 102, 41),
    'AP102': (f78, 102, 42),
    'AQ102': (f79, 102, 43),
    'AR102': (f80, 102, 44),
    'AS102': (f81, 102, 45),
    'AT102': (f82, 102, 46),
    'AU102': (f83, 102, 47),
    'AV102': (f84, 102, 48),
    'AW102': (f85, 102, 49),
    'AX102': (f86, 102, 50),
    'AY102': (f87, 102, 51),
    'AZ102': (f88, 102, 52),
    'BA102': (f89, 102, 53),
    'BB102': (f90, 102, 54),
    'A103': (f1248, 103, 1),
    'B103': (f7, 103, 2),
    'C103': (f1249, 103, 3),
    'D103': (f1250, 103, 4),
    'E103': (f1251, 103, 5),
    'F103': (f1252, 103, 6),
    'G103': (f12, 103, 7),
    'H103': (f13, 103, 8),
    'I103': (f14, 103, 9),
    'J103': (f15, 103, 10),
    'K103': (f1253, 103, 11),
    'L103': (f1254, 103, 12),
    'M103': (f1255, 103, 13),
    'Q103': (f19, 103, 17),
    'R103': (f20, 103, 18),
    'T103': (f1256, 103, 20),
    'U103': (f22, 103, 21),
    'V103': (f23, 103, 22),
    'W103': (f24, 103, 23),
    'X103': (f25, 103, 24),
    'Y103': (f26, 103, 25),
    'Z103': (f27, 103, 26),
    'AA103': (f1257, 103, 27),
    'AB103': (f1258, 103, 28),
    'AC103': (f1259, 103, 29),
    'AD103': (f1260, 103, 30),
    'AE103': (f32, 103, 31),
    'AF103': (f33, 103, 32),
    'AG103': (f69, 103, 33),
    'AH103': (f70, 103, 34),
    'AI103': (f71, 103, 35),
    'AJ103': (f72, 103, 36),
    'AK103': (f73, 103, 37),
    'AL103': (f74, 103, 38),
    'AM103': (f75, 103, 39),
    'AN103': (f76, 103, 40),
    'AO103': (f77, 103, 41),
    'AP103': (f78, 103, 42),
    'AQ103': (f79, 103, 43),
    'AR103': (f80, 103, 44),
    'AS103': (f81, 103, 45),
    'AT103': (f82, 103, 46),
    'AU103': (f83, 103, 47),
    'AV103': (f84, 103, 48),
    'AW103': (f85, 103, 49),
    'AX103': (f86, 103, 50),
    'AY103': (f87, 103, 51),
    'AZ103': (f88, 103, 52),
    'BA103': (f89, 103, 53),
    'BB103': (f90, 103, 54),
    'A104': (f1261, 104, 1),
    'B104': (f7, 104, 2),
    'C104': (f1262, 104, 3),
    'D104': (f1263, 104, 4),
    'E104': (f1264, 104, 5),
    'F104': (f1265, 104, 6),
    'G104': (f12, 104, 7),
    'H104': (f13, 104, 8),
    'I104': (f14, 104, 9),
    'J104': (f15, 104, 10),
    'K104': (f1266, 104, 11),
    'L104': (f1267, 104, 12),
    'M104': (f1268, 104, 13),
    'Q104': (f19, 104, 17),
    'R104': (f20, 104, 18),
    'T104': (f1269, 104, 20),
    'U104': (f22, 104, 21),
    'V104': (f23, 104, 22),
    'W104': (f24, 104, 23),
    'X104': (f25, 104, 24),
    'Y104': (f26, 104, 25),
    'Z104': (f27, 104, 26),
    'AA104': (f1270, 104, 27),
    'AB104': (f1271, 104, 28),
    'AC104': (f1272, 104, 29),
    'AD104': (f1273, 104, 30),
    'AE104': (f32, 104, 31),
    'AF104': (f33, 104, 32),
    'AG104': (f69, 104, 33),
    'AH104': (f70, 104, 34),
    'AI104': (f71, 104, 35),
    'AJ104': (f72, 104, 36),
    'AK104': (f73, 104, 37),
    'AL104': (f74, 104, 38),
    'AM104': (f75, 104, 39),
    'AN104': (f76, 104, 40),
    'AO104': (f77, 104, 41),
    'AP104': (f78, 104, 42),
    'AQ104': (f79, 104, 43),
    'AR104': (f80, 104, 44),
    'AS104': (f81, 104, 45),
    'AT104': (f82, 104, 46),
    'AU104': (f83, 104, 47),
    'AV104': (f84, 104, 48),
    'AW104': (f85, 104, 49),
    'AX104': (f86, 104, 50),
    'AY104': (f87, 104, 51),
    'AZ104': (f88, 104, 52),
    'BA104': (f89, 104, 53),
    'BB104': (f90, 104, 54),
    'A105': (f1274, 105, 1),
    'B105': (f7, 105, 2),
    'C105': (f1275, 105, 3),
    'D105': (f1276, 105, 4),
    'E105': (f1277, 105, 5),
    'F105': (f1278, 105, 6),
    'G105': (f12, 105, 7),
    'H105': (f13, 105, 8),
    'I105': (f14, 105, 9),
    'J105': (f15, 105, 10),
    'K105': (f1279, 105, 11),
    'L105': (f1280, 105, 12),
    'M105': (f1281, 105, 13),
    'Q105': (f19, 105, 17),
    'R105': (f20, 105, 18),
    'T105': (f1282, 105, 20),
    'U105': (f22, 105, 21),
    'V105': (f23, 105, 22),
    'W105': (f24, 105, 23),
    'X105': (f25, 105, 24),
    'Y105': (f26, 105, 25),
    'Z105': (f27, 105, 26),
    'AA105': (f1283, 105, 27),
    'AB105': (f1284, 105, 28),
    'AC105': (f1285, 105, 29),
    'AD105': (f1286, 105, 30),
    'AE105': (f32, 105, 31),
    'AF105': (f33, 105, 32),
    'AG105': (f69, 105, 33),
    'AH105': (f70, 105, 34),
    'AI105': (f71, 105, 35),
    'AJ105': (f72, 105, 36),
    'AK105': (f73, 105, 37),
    'AL105': (f74, 105, 38),
    'AM105': (f75, 105, 39),
    'AN105': (f76, 105, 40),
    'AO105': (f77, 105, 41),
    'AP105': (f78, 105, 42),
    'AQ105': (f79, 105, 43),
    'AR105': (f80, 105, 44),
    'AS105': (f81, 105, 45),
    'AT105': (f82, 105, 46),
    'AU105': (f83, 105, 47),
    'AV105': (f84, 105, 48),
    'AW105': (f85, 105, 49),
    'AX105': (f86, 105, 50),
    'AY105': (f87, 105, 51),
    'AZ105': (f88, 105, 52),
    'BA105': (f89, 105, 53),
    'BB105': (f90, 105, 54),
    'A106': (f1287, 106, 1),
    'B106': (f7, 106, 2),
    'C106': (f1288, 106, 3),
    'D106': (f1289, 106, 4),
    'E106': (f1290, 106, 5),
    'F106': (f1291, 106, 6),
    'G106': (f12, 106, 7),
    'H106': (f13, 106, 8),
    'I106': (f14, 106, 9),
    'J106': (f15, 106, 10),
    'K106': (f1292, 106, 11),
    'L106': (f1293, 106, 12),
    'M106': (f1294, 106, 13),
    'Q106': (f19, 106, 17),
    'R106': (f20, 106, 18),
    'T106': (f1295, 106, 20),
    'U106': (f22, 106, 21),
    'V106': (f23, 106, 22),
    'W106': (f24, 106, 23),
    'X106': (f25, 106, 24),
    'Y106': (f26, 106, 25),
    'Z106': (f27, 106, 26),
    'AA106': (f1296, 106, 27),
    'AB106': (f1297, 106, 28),
    'AC106': (f1298, 106, 29),
    'AD106': (f1299, 106, 30),
    'AE106': (f32, 106, 31),
    'AF106': (f33, 106, 32),
    'AG106': (f69, 106, 33),
    'AH106': (f70, 106, 34),
    'AI106': (f71, 106, 35),
    'AJ106': (f72, 106, 36),
    'AK106': (f73, 106, 37),
    'AL106': (f74, 106, 38),
    'AM106': (f75, 106, 39),
    'AN106': (f76, 106, 40),
    'AO106': (f77, 106, 41),
    'AP106': (f78, 106, 42),
    'AQ106': (f79, 106, 43),
    'AR106': (f80, 106, 44),
    'AS106': (f81, 106, 45),
    'AT106': (f82, 106, 46),
    'AU106': (f83, 106, 47),
    'AV106': (f84, 106, 48),
    'AW106': (f85, 106, 49),
    'AX106': (f86, 106, 50),
    'AY106': (f87, 106, 51),
    'AZ106': (f88, 106, 52),
    'BA106': (f89, 106, 53),
    'BB106': (f90, 106, 54),
    'A107': (f1300, 107, 1),
    'B107': (f7, 107, 2),
    'C107': (f1301, 107, 3),
    'D107': (f1302, 107, 4),
    'E107': (f1303, 107, 5),
    'F107': (f1304, 107, 6),
    'G107': (f12, 107, 7),
    'H107': (f13, 107, 8),
    'I107': (f14, 107, 9),
    'J107': (f15, 107, 10),
    'K107': (f1305, 107, 11),
    'L107': (f1306, 107, 12),
    'M107': (f1307, 107, 13),
    'Q107': (f19, 107, 17),
    'R107': (f20, 107, 18),
    'T107': (f1308, 107, 20),
    'U107': (f22, 107, 21),
    'V107': (f23, 107, 22),
    'W107': (f24, 107, 23),
    'X107': (f25, 107, 24),
    'Y107': (f26, 107, 25),
    'Z107': (f27, 107, 26),
    'AA107': (f1309, 107, 27),
    'AB107': (f1310, 107, 28),
    'AC107': (f1311, 107, 29),
    'AD107': (f1312, 107, 30),
    'AE107': (f32, 107, 31),
    'AF107': (f33, 107, 32),
    'AG107': (f69, 107, 33),
    'AH107': (f70, 107, 34),
    'AI107': (f71, 107, 35),
    'AJ107': (f72, 107, 36),
    'AK107': (f73, 107, 37),
    'AL107': (f74, 107, 38),
    'AM107': (f75, 107, 39),
    'AN107': (f76, 107, 40),
    'AO107': (f77, 107, 41),
    'AP107': (f78, 107, 42),
    'AQ107': (f79, 107, 43),
    'AR107': (f80, 107, 44),
    'AS107': (f81, 107, 45),
    'AT107': (f82, 107, 46),
    'AU107': (f83, 107, 47),
    'AV107': (f84, 107, 48),
    'AW107': (f85, 107, 49),
    'AX107': (f86, 107, 50),
    'AY107': (f87, 107, 51),
    'AZ107': (f88, 107, 52),
    'BA107': (f89, 107, 53),
    'BB107': (f90, 107, 54),
    'A108': (f1313, 108, 1),
    'B108': (f7, 108, 2),
    'C108': (f1314, 108, 3),
    'D108': (f1315, 108, 4),
    'E108': (f1316, 108, 5),
    'F108': (f1317, 108, 6),
    'G108': (f12, 108, 7),
    'H108': (f13, 108, 8),
    'I108': (f14, 108, 9),
    'J108': (f15, 108, 10),
    'K108': (f1318, 108, 11),
    'L108': (f1319, 108, 12),
    'M108': (f1320, 108, 13),
    'Q108': (f19, 108, 17),
    'R108': (f20, 108, 18),
    'T108': (f1321, 108, 20),
    'U108': (f22, 108, 21),
    'V108': (f23, 108, 22),
    'W108': (f24, 108, 23),
    'X108': (f25, 108, 24),
    'Y108': (f26, 108, 25),
    'Z108': (f27, 108, 26),
    'AA108': (f1322, 108, 27),
    'AB108': (f1323, 108, 28),
    'AC108': (f1324, 108, 29),
    'AD108': (f1325, 108, 30),
    'AE108': (f32, 108, 31),
    'AF108': (f33, 108, 32),
    'AG108': (f69, 108, 33),
    'AH108': (f70, 108, 34),
    'AI108': (f71, 108, 35),
    'AJ108': (f72, 108, 36),
    'AK108': (f73, 108, 37),
    'AL108': (f74, 108, 38),
    'AM108': (f75, 108, 39),
    'AN108': (f76, 108, 40),
    'AO108': (f77, 108, 41),
    'AP108': (f78, 108, 42),
    'AQ108': (f79, 108, 43),
    'AR108': (f80, 108, 44),
    'AS108': (f81, 108, 45),
    'AT108': (f82, 108, 46),
    'AU108': (f83, 108, 47),
    'AV108': (f84, 108, 48),
    'AW108': (f85, 108, 49),
    'AX108': (f86, 108, 50),
    'AY108': (f87, 108, 51),
    'AZ108': (f88, 108, 52),
    'BA108': (f89, 108, 53),
    'BB108': (f90, 108, 54),
    'A109': (f1326, 109, 1),
    'B109': (f7, 109, 2),
    'C109': (f1327, 109, 3),
    'D109': (f1328, 109, 4),
    'E109': (f1329, 109, 5),
    'F109': (f1330, 109, 6),
    'G109': (f12, 109, 7),
    'H109': (f13, 109, 8),
    'I109': (f14, 109, 9),
    'J109': (f15, 109, 10),
    'K109': (f1331, 109, 11),
    'L109': (f1332, 109, 12),
    'M109': (f1333, 109, 13),
    'Q109': (f19, 109, 17),
    'R109': (f20, 109, 18),
    'T109': (f1334, 109, 20),
    'U109': (f22, 109, 21),
    'V109': (f23, 109, 22),
    'W109': (f24, 109, 23),
    'X109': (f25, 109, 24),
    'Y109': (f26, 109, 25),
    'Z109': (f27, 109, 26),
    'AA109': (f1335, 109, 27),
    'AB109': (f1336, 109, 28),
    'AC109': (f1337, 109, 29),
    'AD109': (f1338, 109, 30),
    'AE109': (f32, 109, 31),
    'AF109': (f33, 109, 32),
    'AG109': (f69, 109, 33),
    'AH109': (f70, 109, 34),
    'AI109': (f71, 109, 35),
    'AJ109': (f72, 109, 36),
    'AK109': (f73, 109, 37),
    'AL109': (f74, 109, 38),
    'AM109': (f75, 109, 39),
    'AN109': (f76, 109, 40),
    'AO109': (f77, 109, 41),
    'AP109': (f78, 109, 42),
    'AQ109': (f79, 109, 43),
    'AR109': (f80, 109, 44),
    'AS109': (f81, 109, 45),
    'AT109': (f82, 109, 46),
    'AU109': (f83, 109, 47),
    'AV109': (f84, 109, 48),
    'AW109': (f85, 109, 49),
    'AX109': (f86, 109, 50),
    'AY109': (f87, 109, 51),
    'AZ109': (f88, 109, 52),
    'BA109': (f89, 109, 53),
    'BB109': (f90, 109, 54),
    'A110': (f1339, 110, 1),
    'B110': (f7, 110, 2),
    'C110': (f1340, 110, 3),
    'D110': (f1341, 110, 4),
    'E110': (f1342, 110, 5),
    'F110': (f1343, 110, 6),
    'G110': (f12, 110, 7),
    'H110': (f13, 110, 8),
    'I110': (f14, 110, 9),
    'J110': (f15, 110, 10),
    'K110': (f1344, 110, 11),
    'L110': (f1345, 110, 12),
    'M110': (f1346, 110, 13),
    'Q110': (f19, 110, 17),
    'R110': (f20, 110, 18),
    'T110': (f1347, 110, 20),
    'U110': (f22, 110, 21),
    'V110': (f23, 110, 22),
    'W110': (f24, 110, 23),
    'X110': (f25, 110, 24),
    'Y110': (f26, 110, 25),
    'Z110': (f27, 110, 26),
    'AA110': (f1348, 110, 27),
    'AB110': (f1349, 110, 28),
    'AC110': (f1350, 110, 29),
    'AD110': (f1351, 110, 30),
    'AE110': (f32, 110, 31),
    'AF110': (f33, 110, 32),
    'AG110': (f69, 110, 33),
    'AH110': (f70, 110, 34),
    'AI110': (f71, 110, 35),
    'AJ110': (f72, 110, 36),
    'AK110': (f73, 110, 37),
    'AL110': (f74, 110, 38),
    'AM110': (f75, 110, 39),
    'AN110': (f76, 110, 40),
    'AO110': (f77, 110, 41),
    'AP110': (f78, 110, 42),
    'AQ110': (f79, 110, 43),
    'AR110': (f80, 110, 44),
    'AS110': (f81, 110, 45),
    'AT110': (f82, 110, 46),
    'AU110': (f83, 110, 47),
    'AV110': (f84, 110, 48),
    'AW110': (f85, 110, 49),
    'AX110': (f86, 110, 50),
    'AY110': (f87, 110, 51),
    'AZ110': (f88, 110, 52),
    'BA110': (f89, 110, 53),
    'BB110': (f90, 110, 54),
    'A111': (f1352, 111, 1),
    'B111': (f7, 111, 2),
    'C111': (f1353, 111, 3),
    'D111': (f1354, 111, 4),
    'E111': (f1355, 111, 5),
    'F111': (f1356, 111, 6),
    'G111': (f12, 111, 7),
    'H111': (f13, 111, 8),
    'I111': (f14, 111, 9),
    'J111': (f15, 111, 10),
    'K111': (f1357, 111, 11),
    'L111': (f1358, 111, 12),
    'M111': (f1359, 111, 13),
    'Q111': (f19, 111, 17),
    'R111': (f20, 111, 18),
    'T111': (f1360, 111, 20),
    'U111': (f22, 111, 21),
    'V111': (f23, 111, 22),
    'W111': (f24, 111, 23),
    'X111': (f25, 111, 24),
    'Y111': (f26, 111, 25),
    'Z111': (f27, 111, 26),
    'AA111': (f1361, 111, 27),
    'AB111': (f1362, 111, 28),
    'AC111': (f1363, 111, 29),
    'AD111': (f1364, 111, 30),
    'AE111': (f32, 111, 31),
    'AF111': (f33, 111, 32),
    'AG111': (f69, 111, 33),
    'AH111': (f70, 111, 34),
    'AI111': (f71, 111, 35),
    'AJ111': (f72, 111, 36),
    'AK111': (f73, 111, 37),
    'AL111': (f74, 111, 38),
    'AM111': (f75, 111, 39),
    'AN111': (f76, 111, 40),
    'AO111': (f77, 111, 41),
    'AP111': (f78, 111, 42),
    'AQ111': (f79, 111, 43),
    'AR111': (f80, 111, 44),
    'AS111': (f81, 111, 45),
    'AT111': (f82, 111, 46),
    'AU111': (f83, 111, 47),
    'AV111': (f84, 111, 48),
    'AW111': (f85, 111, 49),
    'AX111': (f86, 111, 50),
    'AY111': (f87, 111, 51),
    'AZ111': (f88, 111, 52),
    'BA111': (f89, 111, 53),
    'BB111': (f90, 111, 54),
}
