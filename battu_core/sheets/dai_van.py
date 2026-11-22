# -*- coding: utf-8 -*-
"""Module sheet DAI_VAN – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.8.xlsx (SHA-256 b171b485540d6fb7f1b77d5397c9a0f6a2ddce344baf90abfa50bdfaaf431a5d).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'DAI_VAN'
SI = 14

def f0(E, r, c):
    # Ô C5
    # = NHAP_LIEU!$C$22
    return E.v(2,22,3)

def f1(E, r, c):
    # Ô C6
    # = IF(TRUNG_GIAN!$C$8=0,"Năm dương","Năm âm")
    return ('Năm dương' if X.truthy(X.eq(E.v(23,8,3),0.0)) else 'Năm âm')

def f2(E, r, c):
    # Ô C7
    # = IF(OR(AND(C5="Nam",TRUNG_GIAN!$C$8=0),AND(C5="Nữ",TRUNG_GIAN!$C$8=1)),1,-1)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(14,r-2,c),'Nam'), X.eq(E.v(23,8,3),0.0)), X.AND(X.eq(E.v(14,r-2,c),'Nữ'), X.eq(E.v(23,8,3),1.0)))) else X.neg(1.0))

def f3(E, r, c):
    # Ô C8
    # = IFERROR(IF(NHAP_LIEU!$C$26="","",IF(ISTEXT(NHAP_LIEU!$C$26),DATE(VALUE(MID(TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIE
    #   U!$C$26))+1)+1,4)),VALUE(MID(TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIEU!$C$26))+1,FIND("/",TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIEU!$C$26))+1)
    #   -FIND("/",TRIM(NHAP_LIEU!$C$26))-1)),VALUE(LEFT(TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIEU!$C$26))-1)))+IF((NHAP_LIEU!$C$9="Không rõ giờ"),0.5,IFER
    #   ROR(TIMEVALUE(TRIM(MID(TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIEU!$C$26),FIND("/",TRIM(NHAP_LIEU!$C$26))+1)+5,10))),0)),IF(NHAP_LIEU!$C$26>31,NHAP_
    #   LIEU!$C$26,IF(OR(NHAP_LIEU!$D$26="",NHAP_LIEU!$E$26=""),"",IF(OR(DAY(DATE(NHAP_LIEU!$E$26,NHAP_LIEU!$D$26,NHAP_LIEU!$C$26))<>NHAP_LIEU!$C$26,N(NHAP_LI
    #   EU!$F$26)>23,N(NHAP_LIEU!$G$26)>59),"LỖI",DATE(NHAP_LIEU!$E$26,NHAP_LIEU!$D$26,NHAP_LIEU!$C$26)+IF((NHAP_LIEU!$C$9="Không rõ giờ"),0.5,(N(NHAP_LIEU!$F
    #   $26)*60+N(NHAP_LIEU!$G$26))/1440)))))),"LỖI")
    return X.IFERROR(lambda: ('' if X.truthy(X.eq(E.v(2,26,3),'')) else (X.add(X.DATE(X.VALUE(X.MID(X.TRIM(E.v(2,26,3)), X.add(X.FIND('/', X.TRIM(E.v(2,26,3)), X.add(X.FIND('/', X.TRIM(E.v(2,26,3))),1.0)),1.0), 4.0)), X.VALUE(X.MID(X.TRIM(E.v(2,26,3)), X.add(X.FIND('/', X.TRIM(E.v(2,26,3))),1.0), X.sub(X.sub(X.FIND('/', X.TRIM(E.v(2,26,3)), X.add(X.FIND('/', X.TRIM(E.v(2,26,3))),1.0)),X.FIND('/', X.TRIM(E.v(2,26,3)))),1.0))), X.VALUE(X.LEFT(X.TRIM(E.v(2,26,3)), X.sub(X.FIND('/', X.TRIM(E.v(2,26,3))),1.0)))),(0.5 if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else X.IFERROR(lambda: X.TIMEVALUE(X.TRIM(X.MID(X.TRIM(E.v(2,26,3)), X.add(X.FIND('/', X.TRIM(E.v(2,26,3)), X.add(X.FIND('/', X.TRIM(E.v(2,26,3))),1.0)),5.0), 10.0))), lambda: 0.0))) if X.truthy(X.ISTEXT(lambda: E.v(2,26,3))) else (E.v(2,26,3) if X.truthy(X.gt(E.v(2,26,3),31.0)) else ('' if X.truthy(X.OR(X.eq(E.v(2,26,4),''), X.eq(E.v(2,26,5),''))) else ('LỖI' if X.truthy(X.OR(X.ne(X.DAY(X.DATE(E.v(2,26,5), E.v(2,26,4), E.v(2,26,3))),E.v(2,26,3)), X.gt(X.N(E.v(2,26,6)),23.0), X.gt(X.N(E.v(2,26,7)),59.0))) else X.add(X.DATE(E.v(2,26,5), E.v(2,26,4), E.v(2,26,3)),(0.5 if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else X.div(X.add(X.mul(X.N(E.v(2,26,6)),60.0),X.N(E.v(2,26,7))),1440.0)))))))), lambda: 'LỖI')

def f4(E, r, c):
    # Ô C9
    # = IF(NHAP_LIEU!$C$27="",7,NHAP_LIEU!$C$27)
    return (7.0 if X.truthy(X.eq(E.v(2,27,3),'')) else E.v(2,27,3))

def f5(E, r, c):
    # Ô C10
    # = IF(OR(C8="",C8="LỖI"),"",C8-C9/24)
    return ('' if X.truthy(X.OR(X.eq(E.v(14,r-2,c),''), X.eq(E.v(14,r-2,c),'LỖI'))) else X.sub(E.v(14,r-2,c),X.div(E.v(14,r-1,c),24.0)))

def f6(E, r, c):
    # Ô C11
    # = IF(C10="","",IFERROR(MATCH(C10,TIET_UTC,1),""))
    return ('' if X.truthy(X.eq(E.v(14,r-1,c),'')) else X.IFERROR(lambda: X.MATCH(E.v(14,r-1,c), E.rg(25,4,1,2439,1), 1.0), lambda: ''))

def f7(E, r, c):
    # Ô C12
    # = IF(C11="","",INDEX(TIET_TEN,C11)&" "&TEXT(INDEX(TIET_UTC,C11)+C9/24,"dd/mm/yyyy hh:mm"))
    return ('' if X.truthy(X.eq(E.v(14,r-1,c),'')) else X.cat(X.cat(X.INDEX(E.rg(25,4,3,2439,3), E.v(14,r-1,c)),' '),X.TEXT(X.add(X.INDEX(E.rg(25,4,1,2439,1), E.v(14,r-1,c)),X.div(E.v(14,r-3,c),24.0)), 'dd/mm/yyyy hh:mm')))

def f8(E, r, c):
    # Ô C13
    # = IF(C11="","",INDEX(TIET_TEN,C11+1)&" "&TEXT(INDEX(TIET_UTC,C11+1)+C9/24,"dd/mm/yyyy hh:mm"))
    return ('' if X.truthy(X.eq(E.v(14,r-2,c),'')) else X.cat(X.cat(X.INDEX(E.rg(25,4,3,2439,3), X.add(E.v(14,r-2,c),1.0)),' '),X.TEXT(X.add(X.INDEX(E.rg(25,4,1,2439,1), X.add(E.v(14,r-2,c),1.0)),X.div(E.v(14,r-4,c),24.0)), 'dd/mm/yyyy hh:mm')))

def f9(E, r, c):
    # Ô C14
    # = IF(C11="","",INDEX(TIET_NAM,C11))
    return ('' if X.truthy(X.eq(E.v(14,r-3,c),'')) else X.INDEX(E.rg(25,4,5,2439,5), E.v(14,r-3,c)))

def f10(E, r, c):
    # Ô C15
    # = IF(C11="","",INDEX(TIET_CHI,C11))
    return ('' if X.truthy(X.eq(E.v(14,r-4,c),'')) else X.INDEX(E.rg(25,4,4,2439,4), E.v(14,r-4,c)))

def f11(E, r, c):
    # Ô C16
    # = IF(C8="LỖI","LỖI: ngày giờ sinh không hợp lệ – kiểm tra NHAP_LIEU C26:G26",IF(AND(C8<>"",C11=""),"Ngày sinh ngoài phạm vi bảng tiết 1899–2101",""))&IF
    #   (C11="",IF(C8="","(không nhập ngày giờ sinh – không kiểm tra)",""),IF(TRIM(IF(OR(NHAP_LIEU!$C$37<>NHAP_LIEU!$C$16,NHAP_LIEU!$C$38<>NHAP_LIEU!$C$17),"n
    #   ăm ","")&IF(OR(NHAP_LIEU!$D$37<>NHAP_LIEU!$D$16,NHAP_LIEU!$D$38<>NHAP_LIEU!$D$17),"tháng ","")&IF(OR(NHAP_LIEU!$E$37<>NHAP_LIEU!$E$16,NHAP_LIEU!$E$38<
    #   >NHAP_LIEU!$E$17),"ngày ","")&IF(AND(NHAP_LIEU!$C$9<>"Không rõ giờ",OR(NHAP_LIEU!$F$37<>NHAP_LIEU!$F$16,NHAP_LIEU!$F$38<>NHAP_LIEU!$F$17)),"giờ ",""))
    #   ="","Khớp đủ "&IF(NHAP_LIEU!$C$9="Không rõ giờ","3 trụ năm, tháng, ngày (không rõ giờ)","4 trụ năm, tháng, ngày, giờ"),"CẢNH BÁO: ngày giờ sinh cho tứ
    #    trụ "&NHAP_LIEU!$C$37&" "&NHAP_LIEU!$C$38&", "&NHAP_LIEU!$D$37&" "&NHAP_LIEU!$D$38&", "&NHAP_LIEU!$E$37&" "&NHAP_LIEU!$E$38&", "&NHAP_LIEU!$F$37&" "&
    #   NHAP_LIEU!$F$38&" – KHÁC tứ trụ đang dùng ở trụ "&TRIM(IF(OR(NHAP_LIEU!$C$37<>NHAP_LIEU!$C$16,NHAP_LIEU!$C$38<>NHAP_LIEU!$C$17),"năm ","")&IF(OR(NHAP_
    #   LIEU!$D$37<>NHAP_LIEU!$D$16,NHAP_LIEU!$D$38<>NHAP_LIEU!$D$17),"tháng ","")&IF(OR(NHAP_LIEU!$E$37<>NHAP_LIEU!$E$16,NHAP_LIEU!$E$38<>NHAP_LIEU!$E$17),"n
    #   gày ","")&IF(AND(NHAP_LIEU!$C$9<>"Không rõ giờ",OR(NHAP_LIEU!$F$37<>NHAP_LIEU!$F$16,NHAP_LIEU!$F$38<>NHAP_LIEU!$F$17)),"giờ ",""))&". Khởi vận và năm 
    #   sinh đang tính theo ngày giờ sinh nên KHÔNG khớp với tứ trụ đang dùng: sửa ngày giờ sinh (C26:G26), hoặc xóa C26:G26 và nhập năm sinh ở C30 để công cụ
    #    tự suy ngày sinh từ tứ trụ."))
    return X.cat(('LỖI: ngày giờ sinh không hợp lệ – kiểm tra NHAP_LIEU C26:G26' if X.truthy(X.eq(E.v(14,r-8,c),'LỖI')) else ('Ngày sinh ngoài phạm vi bảng tiết 1899–2101' if X.truthy(X.AND(X.ne(E.v(14,r-8,c),''), X.eq(E.v(14,r-5,c),''))) else '')),(('(không nhập ngày giờ sinh – không kiểm tra)' if X.truthy(X.eq(E.v(14,r-8,c),'')) else '') if X.truthy(X.eq(E.v(14,r-5,c),'')) else (X.cat('Khớp đủ ',('3 trụ năm, tháng, ngày (không rõ giờ)' if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else '4 trụ năm, tháng, ngày, giờ')) if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.OR(X.ne(E.v(2,37,3),E.v(2,16,3)), X.ne(E.v(2,38,3),E.v(2,17,3)))) else ''),('tháng ' if X.truthy(X.OR(X.ne(E.v(2,37,4),E.v(2,16,4)), X.ne(E.v(2,38,4),E.v(2,17,4)))) else '')),('ngày ' if X.truthy(X.OR(X.ne(E.v(2,37,5),E.v(2,16,5)), X.ne(E.v(2,38,5),E.v(2,17,5)))) else '')),('giờ ' if X.truthy(X.AND(X.ne(E.v(2,9,3),'Không rõ giờ'), X.OR(X.ne(E.v(2,37,6),E.v(2,16,6)), X.ne(E.v(2,38,6),E.v(2,17,6))))) else ''))),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('CẢNH BÁO: ngày giờ sinh cho tứ trụ ',E.v(2,37,3)),' '),E.v(2,38,3)),', '),E.v(2,37,4)),' '),E.v(2,38,4)),', '),E.v(2,37,5)),' '),E.v(2,38,5)),', '),E.v(2,37,6)),' '),E.v(2,38,6)),' – KHÁC tứ trụ đang dùng ở trụ '),X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.OR(X.ne(E.v(2,37,3),E.v(2,16,3)), X.ne(E.v(2,38,3),E.v(2,17,3)))) else ''),('tháng ' if X.truthy(X.OR(X.ne(E.v(2,37,4),E.v(2,16,4)), X.ne(E.v(2,38,4),E.v(2,17,4)))) else '')),('ngày ' if X.truthy(X.OR(X.ne(E.v(2,37,5),E.v(2,16,5)), X.ne(E.v(2,38,5),E.v(2,17,5)))) else '')),('giờ ' if X.truthy(X.AND(X.ne(E.v(2,9,3),'Không rõ giờ'), X.OR(X.ne(E.v(2,37,6),E.v(2,16,6)), X.ne(E.v(2,38,6),E.v(2,17,6))))) else '')))),'. Khởi vận và năm sinh đang tính theo ngày giờ sinh nên KHÔNG khớp với tứ trụ đang dùng: sửa ngày giờ sinh (C26:G26), hoặc xóa C26:G26 và nhập năm sinh ở C30 để công cụ tự suy ngày sinh từ tứ trụ.'))))

def f12(E, r, c):
    # Ô C17
    # = IF(C11="","",IF(C7=1,INDEX(TIET_UTC,C11+1)-C10,C10-INDEX(TIET_UTC,C11)))
    return ('' if X.truthy(X.eq(E.v(14,r-6,c),'')) else (X.sub(X.INDEX(E.rg(25,4,1,2439,1), X.add(E.v(14,r-6,c),1.0)),E.v(14,r-7,c)) if X.truthy(X.eq(E.v(14,r-10,c),1.0)) else X.sub(E.v(14,r-7,c),X.INDEX(E.rg(25,4,1,2439,1), E.v(14,r-6,c)))))

def f13(E, r, c):
    # Ô C18
    # = IF(NHAP_LIEU!$C$28="","",NHAP_LIEU!$C$28+IF(NHAP_LIEU!$C$29="",0,NHAP_LIEU!$C$29)/24)
    return ('' if X.truthy(X.eq(E.v(2,28,3),'')) else X.add(E.v(2,28,3),X.div((0.0 if X.truthy(X.eq(E.v(2,29,3),'')) else E.v(2,29,3)),24.0)))

def f14(E, r, c):
    # Ô C19
    # = IF(C17<>"",C17,IF(C18<>"",C18,IF(ISNUMBER(NHAP_LIEU!$C$53),NHAP_LIEU!$C$53,"")))
    return (E.v(14,r-2,c) if X.truthy(X.ne(E.v(14,r-2,c),'')) else (E.v(14,r-1,c) if X.truthy(X.ne(E.v(14,r-1,c),'')) else (E.v(2,53,3) if X.truthy(X.ISNUMBER(lambda: E.v(2,53,3))) else '')))

def f15(E, r, c):
    # Ô C20
    # = IF(C19="","",C19/3)
    return ('' if X.truthy(X.eq(E.v(14,r-1,c),'')) else X.div(E.v(14,r-1,c),3.0))

def f16(E, r, c):
    # Ô D20
    # = IF(C20="","",INT(C20)&" tuổi "&ROUND((C20-INT(C20))*12,0)&" tháng")
    return ('' if X.truthy(X.eq(E.v(14,r,c-1),'')) else X.cat(X.cat(X.cat(X.INT(E.v(14,r,c-1)),' tuổi '),X.ROUND(X.mul(X.sub(E.v(14,r,c-1),X.INT(E.v(14,r,c-1))),12.0), 0.0)),' tháng'))

def f17(E, r, c):
    # Ô C21
    # = IF(C20="",1,MAX(1,ROUND(C20,0)))
    return (1.0 if X.truthy(X.eq(E.v(14,r-1,c),'')) else X.MAX(1.0, X.ROUND(E.v(14,r-1,c), 0.0)))

def f18(E, r, c):
    # Ô C22
    # = IF(C14<>"",C14,IF(ISNUMBER(NHAP_LIEU!$C$54),NHAP_LIEU!$C$54,IF(AND(NHAP_LIEU!$C$30<>"",NHAP_LIEU!$C$45<>0),NHAP_LIEU!$C$30,C23-MOD(C23-1984-MOD(6*TRUN
    #   G_GIAN!$C$5-5*TRUNG_GIAN!$C$6,60),60))))
    return (E.v(14,r-8,c) if X.truthy(X.ne(E.v(14,r-8,c),'')) else (E.v(2,54,3) if X.truthy(X.ISNUMBER(lambda: E.v(2,54,3))) else (E.v(2,30,3) if X.truthy(X.AND(X.ne(E.v(2,30,3),''), X.ne(E.v(2,45,3),0.0))) else X.sub(E.v(14,r+1,c),X.MOD(X.sub(X.sub(E.v(14,r+1,c),1984.0),X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0)), 60.0)))))

def f19(E, r, c):
    # Ô D22
    # = IF(AND(C14="",NHAP_LIEU!$C$30=""),"GIẢ ĐỊNH: chưa nhập năm sinh → lấy năm gần nhất ≤ năm xem có cùng can chi trụ năm; hãy nhập NHAP_LIEU!C30",IF(MOD(C
    #   22-4,12)<>TRUNG_GIAN!$C$6,"CẢNH BÁO: chi của năm sinh khác chi trụ năm – năm sinh ở đây phải là năm theo Lập xuân: người sinh từ 1/1 đến trước Lập xuâ
    #   n (khoảng 4/2) thuộc năm trước (vd. sinh 13/1/1981 → nhập 1980, năm Canh Thân). Tốt nhất nhập đủ ngày giờ sinh ở NHAP"&"_LIEU C26:G26 (P2 tr.268).",""
    #   ))
    return ('GIẢ ĐỊNH: chưa nhập năm sinh → lấy năm gần nhất ≤ năm xem có cùng can chi trụ năm; hãy nhập NHAP_LIEU!C30' if X.truthy(X.AND(X.eq(E.v(14,r-8,c-1),''), X.eq(E.v(2,30,3),''))) else (X.cat('CẢNH BÁO: chi của năm sinh khác chi trụ năm – năm sinh ở đây phải là năm theo Lập xuân: người sinh từ 1/1 đến trước Lập xuân (khoảng 4/2) thuộc năm trước (vd. sinh 13/1/1981 → nhập 1980, năm Canh Thân). Tốt nhất nhập đủ ngày giờ sinh ở NHAP','_LIEU C26:G26 (P2 tr.268).') if X.truthy(X.ne(X.MOD(X.sub(E.v(14,r,c-1),4.0), 12.0),E.v(23,6,3))) else ''))

def f20(E, r, c):
    # Ô C23
    # = IF(NHAP_LIEU!$C$31="",2026,NHAP_LIEU!$C$31)
    return (2026.0 if X.truthy(X.eq(E.v(2,31,3),'')) else E.v(2,31,3))

def f21(E, r, c):
    # Ô C24
    # = C23-C22
    return X.sub(E.v(14,r-1,c),E.v(14,r-2,c))

def f22(E, r, c):
    # Ô D24
    # = C24+1
    return X.add(E.v(14,r,c-1),1.0)

def f23(E, r, c):
    # Ô C25
    # = IF(C24<C21,0,INT((C24-C21)/10)+1)
    return (0.0 if X.truthy(X.lt(E.v(14,r-1,c),E.v(14,r-4,c))) else X.add(X.INT(X.div(X.sub(E.v(14,r-1,c),E.v(14,r-4,c)),10.0)),1.0))

def f24(E, r, c):
    # Ô C26
    # = MOD(6*TRUNG_GIAN!$D$5-5*TRUNG_GIAN!$D$6,60)
    return X.MOD(X.sub(X.mul(6.0,E.v(23,5,4)),X.mul(5.0,E.v(23,6,4))), 60.0)

def f25(E, r, c):
    # Ô C27
    # = MOD(6*TRUNG_GIAN!$F$5-5*TRUNG_GIAN!$F$6,60)
    return X.MOD(X.sub(X.mul(6.0,E.v(23,5,6)),X.mul(5.0,E.v(23,6,6))), 60.0)

def f26(E, r, c):
    # Ô C28
    # = MATCH(8,INDEX(M_TS,TRUNG_GIAN!$E$5+1,0),0)-1
    return X.sub(X.MATCH(8.0, X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), 0.0), 0.0),1.0)

def f27(E, r, c):
    # Ô C32, D32, E32, F32, G32, H32, I32, J32, K32, L32
    # = MOD($C$26+$C$7*C31,60)
    return X.MOD(X.add(E.v(14,26,3),X.mul(E.v(14,7,3),E.v(14,r-1,c))), 60.0)

def f28(E, r, c):
    # Ô C33, D33, E33, F33, G33, H33, I33, J33, K33, L33
    # = MOD(C32,10)
    return X.MOD(E.v(14,r-1,c), 10.0)

def f29(E, r, c):
    # Ô C34, D34, E34, F34, G34, H34, I34, J34, K34, L34
    # = MOD(C32,12)
    return X.MOD(E.v(14,r-2,c), 12.0)

def f30(E, r, c):
    # Ô C35, D35, E35, F35, G35, H35, I35, J35, K35, L35
    # = INDEX(HG_TEN,C32+1)
    return X.INDEX(E.rg(24,140,2,199,2), X.add(E.v(14,r-3,c),1.0))

def f31(E, r, c):
    # Ô C36, D36, E36, F36, G36, H36, I36, J36, K36, L36
    # = $C$21+10*(C31-1)
    return X.add(E.v(14,21,3),X.mul(10.0,X.sub(E.v(14,r-5,c),1.0)))

def f32(E, r, c):
    # Ô C37, D37, E37, F37, G37, H37, I37, J37, K37, L37
    # = $C$22+C36
    return X.add(E.v(14,22,3),E.v(14,r-1,c))

def f33(E, r, c):
    # Ô C38, D38, E38, F38, G38, H38, I38, J38, K38, L38
    # = C37+9
    return X.add(E.v(14,r-1,c),9.0)

def f34(E, r, c):
    # Ô C39, D39, E39, F39, G39, H39, I39, J39, K39, L39
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,C33+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(E.v(14,r-6,c),1.0))

def f35(E, r, c):
    # Ô C40, D40, E40, F40, G40, H40, I40, J40, K40, L40
    # = TRIM(IF(INDEX(CSDL!$D$18:$D$29,C34+1)<0,"",INDEX(CAN_TEN,INDEX(CSDL!$D$18:$D$29,C34+1)+1)&"("&INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C34+
    #   1)+1)&") ")&IF(INDEX(CSDL!$F$18:$F$29,C34+1)<0,"",INDEX(CAN_TEN,INDEX(CSDL!$F$18:$F$29,C34+1)+1)&"("&INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$F$18:$F$
    #   29,C34+1)+1)&") ")&IF(INDEX(CSDL!$H$18:$H$29,C34+1)<0,"",INDEX(CAN_TEN,INDEX(CSDL!$H$18:$H$29,C34+1)+1)&"("&INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$H
    #   $18:$H$29,C34+1)+1)&") "))
    return X.TRIM(X.cat(X.cat(('' if X.truthy(X.lt(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,r-6,c),1.0)),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,r-6,c),1.0)),1.0)),'('),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,r-6,c),1.0)),1.0))),') ')),('' if X.truthy(X.lt(X.INDEX(E.rg(24,18,6,29,6), X.add(E.v(14,r-6,c),1.0)),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(X.INDEX(E.rg(24,18,6,29,6), X.add(E.v(14,r-6,c),1.0)),1.0)),'('),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,6,29,6), X.add(E.v(14,r-6,c),1.0)),1.0))),') '))),('' if X.truthy(X.lt(X.INDEX(E.rg(24,18,8,29,8), X.add(E.v(14,r-6,c),1.0)),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(X.INDEX(E.rg(24,18,8,29,8), X.add(E.v(14,r-6,c),1.0)),1.0)),'('),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,8,29,8), X.add(E.v(14,r-6,c),1.0)),1.0))),') '))))

def f36(E, r, c):
    # Ô C41, D41, E41, F41, G41, H41, I41, J41, K41, L41
    # = INDEX(STAGE_TEN,INDEX(M_TS,TRUNG_GIAN!$E$5+1,C34+1)+1)
    return X.INDEX(E.rg(24,40,1,51,1), X.add(X.INDEX(E.rg(24,5,16,14,27), X.add(E.v(23,5,5),1.0), X.add(E.v(14,r-7,c),1.0)),1.0))

def f37(E, r, c):
    # Ô C42, D42, E42, F42, G42, H42, I42, J42, K42, L42
    # = INDEX(CSDL!$C$5:$C$14,C33+1)
    return X.INDEX(E.rg(24,5,3,14,3), X.add(E.v(14,r-9,c),1.0))

def f38(E, r, c):
    # Ô C43, D43, E43, F43, G43, H43, I43, J43, K43, L43
    # = INDEX(CSDL!$C$18:$C$29,C34+1)
    return X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,r-9,c),1.0))

def f39(E, r, c):
    # Ô C44, D44, E44, F44, G44, H44, I44, J44, K44, L44
    # = IF(C31=$C$25,"◀ ĐANG HÀNH","")
    return ('◀ ĐANG HÀNH' if X.truthy(X.eq(E.v(14,r-13,c),E.v(14,25,3))) else '')

def f40(E, r, c):
    # Ô C47, D47, E47, F47, G47, H47, I47, J47, K47, L47
    # = IF(AND(DUNG_HY_KY!$C$109>=0,C42=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,C42=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,C42=DUNG
    #   _HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))
    return ('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(14,r-5,c),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(14,r-5,c),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(14,r-5,c),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ')))

def f41(E, r, c):
    # Ô C48, D48, E48, F48, G48, H48, I48, J48, K48, L48, C63, D63 … (20 ô)
    # = IF(OR(C47="dụng",C47="hỷ"),"Tốt",IF(C47="kỵ","Xấu","Thường"))
    return ('Tốt' if X.truthy(X.OR(X.eq(E.v(14,r-1,c),'dụng'), X.eq(E.v(14,r-1,c),'hỷ'))) else ('Xấu' if X.truthy(X.eq(E.v(14,r-1,c),'kỵ')) else 'Thường'))

def f42(E, r, c):
    # Ô C49, D49, E49, F49, G49, H49, I49, J49, K49, L49
    # = IF(OR(C43=C42,MOD(C43+1,5)=C42),"chi sinh phù can (+)",IF(MOD(C43+2,5)=C42,"chi khắc can (−)",IF(MOD(C42+2,5)=C43,"can khắc chi: can hao (−)","can sin
    #   h chi: xì hơi (−)")))
    return ('chi sinh phù can (+)' if X.truthy(X.OR(X.eq(E.v(14,r-6,c),E.v(14,r-7,c)), X.eq(X.MOD(X.add(E.v(14,r-6,c),1.0), 5.0),E.v(14,r-7,c)))) else ('chi khắc can (−)' if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-6,c),2.0), 5.0),E.v(14,r-7,c))) else ('can khắc chi: can hao (−)' if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-7,c),2.0), 5.0),E.v(14,r-6,c))) else 'can sinh chi: xì hơi (−)')))

def f43(E, r, c):
    # Ô C50, D50, E50, F50, G50, H50, I50, J50, K50, L50
    # = IF(AND(DUNG_HY_KY!$C$109>=0,C43=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,C43=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,C43=DUNG
    #   _HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))
    return ('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(14,r-7,c),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(14,r-7,c),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(14,r-7,c),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ')))

def f44(E, r, c):
    # Ô C51, D51, E51, F51, G51, H51, I51, J51, K51, L51
    # = TRIM(IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$C$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$C$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$C$9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_
    #   GIAN!$C$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$C$9=DU
    #   NG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="kỵ")),"Năm ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$D$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$D$9=2)),MOD(
    #   C43+2,5)=TRUNG_GIAN!$D$9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$D$9=DUNG_HY_
    #   KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="kỵ")),"Tháng ","")&IF(AND(INDEX(M_XU,C
    #   34+1,TRUNG_GIAN!$E$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$E$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$E$9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$
    #   109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$111),"kỵ","k
    #   hông phải dụng/hỷ/kỵ")))="kỵ")),"Ngày ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$F$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$F$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$F$
    #   9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(
    #   DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="kỵ")),"Giờ ",""))
    return X.TRIM(X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-17,c),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-8,c),2.0), X.eq(E.v(23,9,3),2.0))), X.eq(X.MOD(X.add(E.v(14,r-8,c),2.0), 5.0),E.v(23,9,3)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,3),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,3),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,3),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else ''),('Tháng ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-17,c),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-8,c),2.0), X.eq(E.v(23,9,4),2.0))), X.eq(X.MOD(X.add(E.v(14,r-8,c),2.0), 5.0),E.v(23,9,4)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,4),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,4),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,4),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else '')),('Ngày ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-17,c),1.0), X.add(E.v(23,6,5),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-8,c),2.0), X.eq(E.v(23,9,5),2.0))), X.eq(X.MOD(X.add(E.v(14,r-8,c),2.0), 5.0),E.v(23,9,5)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,5),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,5),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,5),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else '')),('Giờ ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-17,c),1.0), X.add(E.v(23,6,6),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-8,c),2.0), X.eq(E.v(23,9,6),2.0))), X.eq(X.MOD(X.add(E.v(14,r-8,c),2.0), 5.0),E.v(23,9,6)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,6),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,6),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,6),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else '')))

def f45(E, r, c):
    # Ô C52, D52, E52, F52, G52, H52, I52, J52, K52, L52
    # = TRIM(IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$C$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$C$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$C$9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_
    #   GIAN!$C$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$C$9=DU
    #   NG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110
    #   >=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="hỷ")),"Năm
    #    ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$D$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$D$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$D$9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,TRUN
    #   G_GIAN!$D$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$D$9=
    #   DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$1
    #   10>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="hỷ")),"T
    #   háng ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$E$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$E$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$E$9,OR(IF(AND(DUNG_HY_KY!$C$109>=0,
    #   TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$
    #   E$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!
    #   $C$110>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="hỷ")
    #   ),"Ngày ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$F$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$F$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$F$9,OR(IF(AND(DUNG_HY_KY!$C$109>
    #   =0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_KY!$C$110>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIA
    #   N!$F$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$C$109>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$109),"dụng",IF(AND(DUNG_HY_
    #   KY!$C$110>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$110),"hỷ",IF(AND(DUNG_HY_KY!$C$111>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$C$111),"kỵ","không phải dụng/hỷ/kỵ")))="h
    #   ỷ")),"Giờ ",""))
    return X.TRIM(X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-18,c),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-9,c),2.0), X.eq(E.v(23,9,3),2.0))), X.eq(X.MOD(X.add(E.v(14,r-9,c),2.0), 5.0),E.v(23,9,3)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,3),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,3),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,3),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,3),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,3),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,3),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else ''),('Tháng ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-18,c),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-9,c),2.0), X.eq(E.v(23,9,4),2.0))), X.eq(X.MOD(X.add(E.v(14,r-9,c),2.0), 5.0),E.v(23,9,4)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,4),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,4),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,4),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,4),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,4),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,4),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else '')),('Ngày ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-18,c),1.0), X.add(E.v(23,6,5),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-9,c),2.0), X.eq(E.v(23,9,5),2.0))), X.eq(X.MOD(X.add(E.v(14,r-9,c),2.0), 5.0),E.v(23,9,5)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,5),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,5),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,5),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,5),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,5),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,5),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else '')),('Giờ ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-18,c),1.0), X.add(E.v(23,6,6),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-9,c),2.0), X.eq(E.v(23,9,6),2.0))), X.eq(X.MOD(X.add(E.v(14,r-9,c),2.0), 5.0),E.v(23,9,6)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,6),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,6),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,6),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(23,9,6),E.v(8,109,3)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(23,9,6),E.v(8,110,3)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(23,9,6),E.v(8,111,3)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else '')))

def f46(E, r, c):
    # Ô C53, D53, E53, F53, G53, H53, I53, J53, K53, L53
    # = TRIM(IF(AND(MOD(TRUNG_GIAN!$C$7+2,5)=C42,PP2_DINH_LUONG!$C$40=1,PP1_DINH_TINH!$C$38=0),"Năm khắc; ","")&IF(INDEX(M_NH,TRUNG_GIAN!$C$5+1,C33+1)>=0,"Năm
    #    hợp; ","")&IF(AND(MOD(TRUNG_GIAN!$D$7+2,5)=C42,PP2_DINH_LUONG!$D$40=1,PP1_DINH_TINH!$D$38=0),"Tháng khắc; ","")&IF(INDEX(M_NH,TRUNG_GIAN!$D$5+1,C33+1
    #   )>=0,"Tháng hợp; ","")&IF(AND(MOD(TRUNG_GIAN!$F$7+2,5)=C42,PP2_DINH_LUONG!$F$40=1,PP1_DINH_TINH!$F$38=0),"Giờ khắc; ","")&IF(INDEX(M_NH,TRUNG_GIAN!$F$
    #   5+1,C33+1)>=0,"Giờ hợp; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(('Năm khắc; ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(14,r-11,c)), X.eq(E.v(10,40,3),1.0), X.eq(E.v(9,38,3),0.0))) else ''),('Năm hợp; ' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,3),1.0), X.add(E.v(14,r-20,c),1.0)),0.0)) else '')),('Tháng khắc; ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(14,r-11,c)), X.eq(E.v(10,40,4),1.0), X.eq(E.v(9,38,4),0.0))) else '')),('Tháng hợp; ' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,4),1.0), X.add(E.v(14,r-20,c),1.0)),0.0)) else '')),('Giờ khắc; ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(14,r-11,c)), X.eq(E.v(10,40,6),1.0), X.eq(E.v(9,38,6),0.0))) else '')),('Giờ hợp; ' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,6),1.0), X.add(E.v(14,r-20,c),1.0)),0.0)) else '')))

def f47(E, r, c):
    # Ô C54, D54, E54, F54, G54, H54, I54, J54, K54, L54
    # = IF(C115>0,"Tốt","Xấu")
    return ('Tốt' if X.truthy(X.gt(E.v(14,r+61,c),0.0)) else 'Xấu')

def f48(E, r, c):
    # Ô C55, D55, E55, F55, G55, H55, I55, J55, K55, L55, C70, D70 … (20 ô)
    # = IF(AND(LEN(C52)>0,LEN(C51)=0),"đ.2: xung mất dụng/hỷ thần ở "&C52,IF(AND(LEN(C51)>0,LEN(C52)=0),"đ.1: xung mất kỵ thần ở "&C51,IF(AND(C48="Tốt",LEN(C5
    #   3)>0),"đ.3: vận tốt bị mệnh cục khắc/hợp ("&C53&") → tốt không thực",IF(AND(C48="Xấu",LEN(C53)>0),"đ.4: vận xấu bị mệnh cục chế ("&C53&")",IF(AND(LEN(
    #   C51)>0,LEN(C52)>0),"vừa mất kỵ vừa mất dụng – giữ xu thế gốc","không điều chỉnh")))))
    return (X.cat('đ.2: xung mất dụng/hỷ thần ở ',E.v(14,r-3,c)) if X.truthy(X.AND(X.gt(X.LEN(E.v(14,r-3,c)),0.0), X.eq(X.LEN(E.v(14,r-4,c)),0.0))) else (X.cat('đ.1: xung mất kỵ thần ở ',E.v(14,r-4,c)) if X.truthy(X.AND(X.gt(X.LEN(E.v(14,r-4,c)),0.0), X.eq(X.LEN(E.v(14,r-3,c)),0.0))) else (X.cat(X.cat('đ.3: vận tốt bị mệnh cục khắc/hợp (',E.v(14,r-2,c)),') → tốt không thực') if X.truthy(X.AND(X.eq(E.v(14,r-7,c),'Tốt'), X.gt(X.LEN(E.v(14,r-2,c)),0.0))) else (X.cat(X.cat('đ.4: vận xấu bị mệnh cục chế (',E.v(14,r-2,c)),')') if X.truthy(X.AND(X.eq(E.v(14,r-7,c),'Xấu'), X.gt(X.LEN(E.v(14,r-2,c)),0.0))) else ('vừa mất kỵ vừa mất dụng – giữ xu thế gốc' if X.truthy(X.AND(X.gt(X.LEN(E.v(14,r-4,c)),0.0), X.gt(X.LEN(E.v(14,r-3,c)),0.0))) else 'không điều chỉnh')))))

def f49(E, r, c):
    # Ô C56, D56, E56, F56, G56, H56, I56, J56, K56, L56
    # = IF(DUNG_HY_KY!$C$7=2,"—",IF(OR(MOD(C42-DUNG_HY_KY!$C$5,5)=0,MOD(C42-DUNG_HY_KY!$C$5,5)=4),IF(DUNG_HY_KY!$C$7=1,"bại vận (sinh trợ thân vượng)","vận tố
    #   t (sinh trợ thân nhược)"),IF(DUNG_HY_KY!$C$7=1,"vận tốt (khắc tiết hao thân vượng)","bại vận (khắc tiết hao thân nhược)")))
    return ('—' if X.truthy(X.eq(E.v(8,7,3),2.0)) else (('bại vận (sinh trợ thân vượng)' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'vận tốt (sinh trợ thân nhược)') if X.truthy(X.OR(X.eq(X.MOD(X.sub(E.v(14,r-14,c),E.v(8,5,3)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,r-14,c),E.v(8,5,3)), 5.0),4.0))) else ('vận tốt (khắc tiết hao thân vượng)' if X.truthy(X.eq(E.v(8,7,3),1.0)) else 'bại vận (khắc tiết hao thân nhược)')))

def f50(E, r, c):
    # Ô C57, D57, E57, F57, G57, H57, I57, J57, K57, L57
    # = IF(DUNG_HY_KY!$C$109<0,0,IF(C34=INDEX(CSDL!$C$33:$C$37,DUNG_HY_KY!$C$109+1),1,0))
    return (0.0 if X.truthy(X.lt(E.v(8,109,3),0.0)) else (1.0 if X.truthy(X.eq(E.v(14,r-23,c),X.INDEX(E.rg(24,33,3,37,3), X.add(E.v(8,109,3),1.0)))) else 0.0))

def f51(E, r, c):
    # Ô C58, D58, E58, F58, G58, H58, I58, J58, K58, L58
    # = "Vận "&C35&" ("&C36&"–"&(C36+9)&" tuổi): can "&INDEX(CAN_TEN,C33+1)&" là "&C39&", hành "&INDEX(HANH_TEN,C42+1)&" = "&C47&" → "&C48&" (P2 tr.281; P1 tr
    #   .120). Chi "&INDEX(CHI_TEN,C34+1)&" hành "&INDEX(HANH_TEN,C43+1)&" ("&C50&"), "&C49&" (P1 tr.123). Điều chỉnh: "&C55&" → xu thế "&C54&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r-23,c)),' ('),E.v(14,r-22,c)),'–'),X.add(E.v(14,r-22,c),9.0)),' tuổi): can '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(14,r-25,c),1.0))),' là '),E.v(14,r-19,c)),', hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,r-16,c),1.0))),' = '),E.v(14,r-11,c)),' → '),E.v(14,r-10,c)),' (P2 tr.281; P1 tr.120). Chi '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-24,c),1.0))),' hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,r-15,c),1.0))),' ('),E.v(14,r-8,c)),'), '),E.v(14,r-9,c)),' (P1 tr.123). Điều chỉnh: '),E.v(14,r-3,c)),' → xu thế '),E.v(14,r-4,c)),'.')

def f52(E, r, c):
    # Ô C62, D62, E62, F62, G62, H62, I62, J62, K62, L62
    # = IF(AND(DUNG_HY_KY!$D$109>=0,C42=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,C42=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,C42=DUNG
    #   _HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))
    return ('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(14,r-20,c),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(14,r-20,c),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(14,r-20,c),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ')))

def f53(E, r, c):
    # Ô C64, D64, E64, F64, G64, H64, I64, J64, K64, L64
    # = IF(OR(C43=C42,MOD(C43+1,5)=C42),"chi sinh phù can (+)",IF(MOD(C43+2,5)=C42,"chi khắc can (−)",IF(MOD(C42+2,5)=C43,"can khắc chi: can hao (−)","can sin
    #   h chi: xì hơi (−)")))
    return ('chi sinh phù can (+)' if X.truthy(X.OR(X.eq(E.v(14,r-21,c),E.v(14,r-22,c)), X.eq(X.MOD(X.add(E.v(14,r-21,c),1.0), 5.0),E.v(14,r-22,c)))) else ('chi khắc can (−)' if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-21,c),2.0), 5.0),E.v(14,r-22,c))) else ('can khắc chi: can hao (−)' if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-22,c),2.0), 5.0),E.v(14,r-21,c))) else 'can sinh chi: xì hơi (−)')))

def f54(E, r, c):
    # Ô C65, D65, E65, F65, G65, H65, I65, J65, K65, L65
    # = IF(AND(DUNG_HY_KY!$D$109>=0,C43=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,C43=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,C43=DUNG
    #   _HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))
    return ('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(14,r-22,c),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(14,r-22,c),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(14,r-22,c),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ')))

def f55(E, r, c):
    # Ô C66, D66, E66, F66, G66, H66, I66, J66, K66, L66
    # = TRIM(IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$C$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$C$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$C$9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_
    #   GIAN!$C$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$C$9=DU
    #   NG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="kỵ")),"Năm ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$D$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$D$9=2)),MOD(
    #   C43+2,5)=TRUNG_GIAN!$D$9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$D$9=DUNG_HY_
    #   KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="kỵ")),"Tháng ","")&IF(AND(INDEX(M_XU,C
    #   34+1,TRUNG_GIAN!$E$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$E$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$E$9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$
    #   109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$111),"kỵ","k
    #   hông phải dụng/hỷ/kỵ")))="kỵ")),"Ngày ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$F$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$F$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$F$
    #   9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(
    #   DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="kỵ")),"Giờ ",""))
    return X.TRIM(X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-32,c),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-23,c),2.0), X.eq(E.v(23,9,3),2.0))), X.eq(X.MOD(X.add(E.v(14,r-23,c),2.0), 5.0),E.v(23,9,3)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,3),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,3),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,3),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else ''),('Tháng ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-32,c),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-23,c),2.0), X.eq(E.v(23,9,4),2.0))), X.eq(X.MOD(X.add(E.v(14,r-23,c),2.0), 5.0),E.v(23,9,4)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,4),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,4),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,4),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else '')),('Ngày ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-32,c),1.0), X.add(E.v(23,6,5),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-23,c),2.0), X.eq(E.v(23,9,5),2.0))), X.eq(X.MOD(X.add(E.v(14,r-23,c),2.0), 5.0),E.v(23,9,5)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,5),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,5),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,5),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else '')),('Giờ ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-32,c),1.0), X.add(E.v(23,6,6),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-23,c),2.0), X.eq(E.v(23,9,6),2.0))), X.eq(X.MOD(X.add(E.v(14,r-23,c),2.0), 5.0),E.v(23,9,6)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,6),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,6),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,6),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'kỵ')))) else '')))

def f56(E, r, c):
    # Ô C67, D67, E67, F67, G67, H67, I67, J67, K67, L67
    # = TRIM(IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$C$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$C$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$C$9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_
    #   GIAN!$C$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$C$9=DU
    #   NG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110
    #   >=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$C$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="hỷ")),"Năm
    #    ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$D$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$D$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$D$9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,TRUN
    #   G_GIAN!$D$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$D$9=
    #   DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$1
    #   10>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$D$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="hỷ")),"T
    #   háng ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$E$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$E$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$E$9,OR(IF(AND(DUNG_HY_KY!$D$109>=0,
    #   TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$
    #   E$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!
    #   $D$110>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$E$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="hỷ")
    #   ),"Ngày ","")&IF(AND(INDEX(M_XU,C34+1,TRUNG_GIAN!$F$6+1)=1,NOT(AND(C43=2,TRUNG_GIAN!$F$9=2)),MOD(C43+2,5)=TRUNG_GIAN!$F$9,OR(IF(AND(DUNG_HY_KY!$D$109>
    #   =0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_KY!$D$110>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIA
    #   N!$F$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="dụng",IF(AND(DUNG_HY_KY!$D$109>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$109),"dụng",IF(AND(DUNG_HY_
    #   KY!$D$110>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$110),"hỷ",IF(AND(DUNG_HY_KY!$D$111>=0,TRUNG_GIAN!$F$9=DUNG_HY_KY!$D$111),"kỵ","không phải dụng/hỷ/kỵ")))="h
    #   ỷ")),"Giờ ",""))
    return X.TRIM(X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-33,c),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-24,c),2.0), X.eq(E.v(23,9,3),2.0))), X.eq(X.MOD(X.add(E.v(14,r-24,c),2.0), 5.0),E.v(23,9,3)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,3),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,3),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,3),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,3),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,3),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,3),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else ''),('Tháng ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-33,c),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-24,c),2.0), X.eq(E.v(23,9,4),2.0))), X.eq(X.MOD(X.add(E.v(14,r-24,c),2.0), 5.0),E.v(23,9,4)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,4),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,4),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,4),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,4),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,4),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,4),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else '')),('Ngày ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-33,c),1.0), X.add(E.v(23,6,5),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-24,c),2.0), X.eq(E.v(23,9,5),2.0))), X.eq(X.MOD(X.add(E.v(14,r-24,c),2.0), 5.0),E.v(23,9,5)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,5),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,5),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,5),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,5),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,5),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,5),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else '')),('Giờ ' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-33,c),1.0), X.add(E.v(23,6,6),1.0)),1.0), X.NOT(X.AND(X.eq(E.v(14,r-24,c),2.0), X.eq(E.v(23,9,6),2.0))), X.eq(X.MOD(X.add(E.v(14,r-24,c),2.0), 5.0),E.v(23,9,6)), X.OR(X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,6),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,6),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,6),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'dụng'), X.eq(('dụng' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(23,9,6),E.v(8,109,4)))) else ('hỷ' if X.truthy(X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(23,9,6),E.v(8,110,4)))) else ('kỵ' if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(23,9,6),E.v(8,111,4)))) else 'không phải dụng/hỷ/kỵ'))),'hỷ')))) else '')))

def f57(E, r, c):
    # Ô C68, D68, E68, F68, G68, H68, I68, J68, K68, L68
    # = TRIM(IF(AND(MOD(TRUNG_GIAN!$C$7+2,5)=C42,PP2_DINH_LUONG!$C$40=1,PP1_DINH_TINH!$C$38=0),"Năm khắc; ","")&IF(INDEX(M_NH,TRUNG_GIAN!$C$5+1,C33+1)>=0,"Năm
    #    hợp; ","")&IF(AND(MOD(TRUNG_GIAN!$D$7+2,5)=C42,PP2_DINH_LUONG!$D$40=1,PP1_DINH_TINH!$D$38=0),"Tháng khắc; ","")&IF(INDEX(M_NH,TRUNG_GIAN!$D$5+1,C33+1
    #   )>=0,"Tháng hợp; ","")&IF(AND(MOD(TRUNG_GIAN!$F$7+2,5)=C42,PP2_DINH_LUONG!$F$40=1,PP1_DINH_TINH!$F$38=0),"Giờ khắc; ","")&IF(INDEX(M_NH,TRUNG_GIAN!$F$
    #   5+1,C33+1)>=0,"Giờ hợp; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(('Năm khắc; ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(14,r-26,c)), X.eq(E.v(10,40,3),1.0), X.eq(E.v(9,38,3),0.0))) else ''),('Năm hợp; ' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,3),1.0), X.add(E.v(14,r-35,c),1.0)),0.0)) else '')),('Tháng khắc; ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(14,r-26,c)), X.eq(E.v(10,40,4),1.0), X.eq(E.v(9,38,4),0.0))) else '')),('Tháng hợp; ' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,4),1.0), X.add(E.v(14,r-35,c),1.0)),0.0)) else '')),('Giờ khắc; ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(14,r-26,c)), X.eq(E.v(10,40,6),1.0), X.eq(E.v(9,38,6),0.0))) else '')),('Giờ hợp; ' if X.truthy(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,6),1.0), X.add(E.v(14,r-35,c),1.0)),0.0)) else '')))

def f58(E, r, c):
    # Ô C69, D69, E69, F69, G69, H69, I69, J69, K69, L69
    # = IF(C117>0,"Tốt","Xấu")
    return ('Tốt' if X.truthy(X.gt(E.v(14,r+48,c),0.0)) else 'Xấu')

def f59(E, r, c):
    # Ô C71, D71, E71, F71, G71, H71, I71, J71, K71, L71
    # = IF(DUNG_HY_KY!$D$7=2,"—",IF(OR(MOD(C42-DUNG_HY_KY!$D$5,5)=0,MOD(C42-DUNG_HY_KY!$D$5,5)=4),IF(DUNG_HY_KY!$D$7=1,"bại vận (sinh trợ thân vượng)","vận tố
    #   t (sinh trợ thân nhược)"),IF(DUNG_HY_KY!$D$7=1,"vận tốt (khắc tiết hao thân vượng)","bại vận (khắc tiết hao thân nhược)")))
    return ('—' if X.truthy(X.eq(E.v(8,7,4),2.0)) else (('bại vận (sinh trợ thân vượng)' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'vận tốt (sinh trợ thân nhược)') if X.truthy(X.OR(X.eq(X.MOD(X.sub(E.v(14,r-29,c),E.v(8,5,4)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,r-29,c),E.v(8,5,4)), 5.0),4.0))) else ('vận tốt (khắc tiết hao thân vượng)' if X.truthy(X.eq(E.v(8,7,4),1.0)) else 'bại vận (khắc tiết hao thân nhược)')))

def f60(E, r, c):
    # Ô C72, D72, E72, F72, G72, H72, I72, J72, K72, L72
    # = IF(DUNG_HY_KY!$D$109<0,0,IF(C34=INDEX(CSDL!$C$33:$C$37,DUNG_HY_KY!$D$109+1),1,0))
    return (0.0 if X.truthy(X.lt(E.v(8,109,4),0.0)) else (1.0 if X.truthy(X.eq(E.v(14,r-38,c),X.INDEX(E.rg(24,33,3,37,3), X.add(E.v(8,109,4),1.0)))) else 0.0))

def f61(E, r, c):
    # Ô C73, D73, E73, F73, G73, H73, I73, J73, K73, L73
    # = "Vận "&C35&" ("&C36&"–"&(C36+9)&" tuổi): can "&INDEX(CAN_TEN,C33+1)&" là "&C39&", hành "&INDEX(HANH_TEN,C42+1)&" = "&C62&" → "&C63&" (P2 tr.281; P1 tr
    #   .120). Chi "&INDEX(CHI_TEN,C34+1)&" hành "&INDEX(HANH_TEN,C43+1)&" ("&C65&"), "&C64&" (P1 tr.123). Điều chỉnh: "&C70&" → xu thế "&C69&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,r-38,c)),' ('),E.v(14,r-37,c)),'–'),X.add(E.v(14,r-37,c),9.0)),' tuổi): can '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(14,r-40,c),1.0))),' là '),E.v(14,r-34,c)),', hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,r-31,c),1.0))),' = '),E.v(14,r-11,c)),' → '),E.v(14,r-10,c)),' (P2 tr.281; P1 tr.120). Chi '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-39,c),1.0))),' hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,r-30,c),1.0))),' ('),E.v(14,r-8,c)),'), '),E.v(14,r-9,c)),' (P1 tr.123). Điều chỉnh: '),E.v(14,r-3,c)),' → xu thế '),E.v(14,r-4,c)),'.')

def f62(E, r, c):
    # Ô C77, D77, E77, F77, G77, H77, I77, J77, K77, L77
    # = TRIM(IF(AND(OR(MOD(C42+2,5)=TRUNG_GIAN!$C$7,MOD(TRUNG_GIAN!$C$7+2,5)=C42),INDEX(M_XU,C34+1,TRUNG_GIAN!$C$6+1)=1),"Năm ","")&IF(AND(OR(MOD(C42+2,5)=TRU
    #   NG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+2,5)=C42),INDEX(M_XU,C34+1,TRUNG_GIAN!$D$6+1)=1),"Tháng ","")&IF(AND(OR(MOD(C42+2,5)=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!$
    #   E$7+2,5)=C42),INDEX(M_XU,C34+1,TRUNG_GIAN!$E$6+1)=1),"Ngày ","")&IF(AND(OR(MOD(C42+2,5)=TRUNG_GIAN!$F$7,MOD(TRUNG_GIAN!$F$7+2,5)=C42),INDEX(M_XU,C34+1
    #   ,TRUNG_GIAN!$F$6+1)=1),"Giờ ",""))
    return X.TRIM(X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(14,r-35,c),2.0), 5.0),E.v(23,7,3)), X.eq(X.MOD(X.add(E.v(23,7,3),2.0), 5.0),E.v(14,r-35,c))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-43,c),1.0), X.add(E.v(23,6,3),1.0)),1.0))) else ''),('Tháng ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(14,r-35,c),2.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(14,r-35,c))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-43,c),1.0), X.add(E.v(23,6,4),1.0)),1.0))) else '')),('Ngày ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(14,r-35,c),2.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),E.v(14,r-35,c))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-43,c),1.0), X.add(E.v(23,6,5),1.0)),1.0))) else '')),('Giờ ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(14,r-35,c),2.0), 5.0),E.v(23,7,6)), X.eq(X.MOD(X.add(E.v(23,7,6),2.0), 5.0),E.v(14,r-35,c))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-43,c),1.0), X.add(E.v(23,6,6),1.0)),1.0))) else '')))

def f63(E, r, c):
    # Ô C78, D78, E78, F78, G78, H78, I78, J78, K78, L78
    # = IF(AND(OR(MOD(C42+2,5)=TRUNG_GIAN!$D$7,MOD(TRUNG_GIAN!$D$7+2,5)=C42),INDEX(M_XU,C34+1,TRUNG_GIAN!$D$6+1)=1),1,0)
    return (1.0 if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(14,r-36,c),2.0), 5.0),E.v(23,7,4)), X.eq(X.MOD(X.add(E.v(23,7,4),2.0), 5.0),E.v(14,r-36,c))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(14,r-44,c),1.0), X.add(E.v(23,6,4),1.0)),1.0))) else 0.0)

def f64(E, r, c):
    # Ô C79, D79, E79, F79, G79, H79, I79, J79, K79, L79
    # = TRIM(IF(INDEX(M_HINH,C34+1,TRUNG_GIAN!$C$6+1)=1,"Năm ","")&IF(INDEX(M_HINH,C34+1,TRUNG_GIAN!$D$6+1)=1,"Tháng ","")&IF(INDEX(M_HINH,C34+1,TRUNG_GIAN!$E
    #   $6+1)=1,"Ngày ","")&IF(INDEX(M_HINH,C34+1,TRUNG_GIAN!$F$6+1)=1,"Giờ ",""))
    return X.TRIM(X.cat(X.cat(X.cat(('Năm ' if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(14,r-45,c),1.0), X.add(E.v(23,6,3),1.0)),1.0)) else ''),('Tháng ' if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(14,r-45,c),1.0), X.add(E.v(23,6,4),1.0)),1.0)) else '')),('Ngày ' if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(14,r-45,c),1.0), X.add(E.v(23,6,5),1.0)),1.0)) else '')),('Giờ ' if X.truthy(X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(14,r-45,c),1.0), X.add(E.v(23,6,6),1.0)),1.0)) else '')))

def f65(E, r, c):
    # Ô C80, D80, E80, F80, G80, H80, I80, J80, K80, L80
    # = IF(C34=INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1),1,0)
    return (1.0 if X.truthy(X.eq(E.v(14,r-46,c),X.INDEX(E.rg(24,5,8,14,8), X.add(E.v(23,5,5),1.0)))) else 0.0)

def f66(E, r, c):
    # Ô C81, D81, E81, F81, G81, H81, I81, J81, K81, L81
    # = IF(C34=$C$28,1,0)
    return (1.0 if X.truthy(X.eq(E.v(14,r-47,c),E.v(14,28,3))) else 0.0)

def f67(E, r, c):
    # Ô C82, D82, E82, F82, G82, H82, I82, J82, K82, L82
    # = IF(OR(AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$D$6+1)="can",C33=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),AND(INDEX(CSDL!$B$122:$B$133,TRUNG_GIAN!$
    #   D$6+1)="chi",C34=INDEX(CSDL!$C$122:$C$133,TRUNG_GIAN!$D$6+1)),C33=INDEX(CSDL!$D$122:$D$133,TRUNG_GIAN!$D$6+1)),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'can'), X.eq(E.v(14,r-49,c),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.AND(X.eq(X.INDEX(E.rg(24,122,2,133,2), X.add(E.v(23,6,4),1.0)),'chi'), X.eq(E.v(14,r-48,c),X.INDEX(E.rg(24,122,3,133,3), X.add(E.v(23,6,4),1.0)))), X.eq(E.v(14,r-49,c),X.INDEX(E.rg(24,122,4,133,4), X.add(E.v(23,6,4),1.0))))) else 0.0)

def f68(E, r, c):
    # Ô C83, D83, E83, F83, G83, H83, I83, J83, K83, L83
    # = IF(OR(C34=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$E$5+1),C34=INDEX(CSDL!$K$5:$K$14,TRUNG_GIAN!$E$5+1),C34=INDEX(CSDL!$J$5:$J$14,TRUNG_GIAN!$C$5+1),C34=INDEX
    #   (CSDL!$K$5:$K$14,TRUNG_GIAN!$C$5+1)),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(14,r-49,c),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,5),1.0))), X.eq(E.v(14,r-49,c),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,5),1.0))), X.eq(E.v(14,r-49,c),X.INDEX(E.rg(24,5,10,14,10), X.add(E.v(23,5,3),1.0))), X.eq(E.v(14,r-49,c),X.INDEX(E.rg(24,5,11,14,11), X.add(E.v(23,5,3),1.0))))) else 0.0)

def f69(E, r, c):
    # Ô C84, D84, E84, F84, G84, H84, I84, J84, K84, L84
    # = IF(OR(INDEX(M_NH,TRUNG_GIAN!$E$5+1,C33+1)>=0,INDEX(M_LH,TRUNG_GIAN!$E$6+1,C34+1)>=0),1,0)
    return (1.0 if X.truthy(X.OR(X.ge(X.INDEX(E.rg(24,19,16,28,25), X.add(E.v(23,5,5),1.0), X.add(E.v(14,r-51,c),1.0)),0.0), X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,r-50,c),1.0)),0.0))) else 0.0)

def f70(E, r, c):
    # Ô C85, D85, E85, F85, G85, H85, I85, J85, K85, L85
    # = IF(OR(AND(NHAP_LIEU!$C$22="Nam",OR(C39="Chính tài",C39="Thiên tài")),AND(NHAP_LIEU!$C$22="Nữ",OR(C39="Chính quan",C39="Thất sát"))),1,0)
    return (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(2,22,3),'Nam'), X.OR(X.eq(E.v(14,r-46,c),'Chính tài'), X.eq(E.v(14,r-46,c),'Thiên tài'))), X.AND(X.eq(E.v(2,22,3),'Nữ'), X.OR(X.eq(E.v(14,r-46,c),'Chính quan'), X.eq(E.v(14,r-46,c),'Thất sát'))))) else 0.0)

def f71(E, r, c):
    # Ô C86, D86, E86, F86, G86, H86, I86, J86, K86, L86
    # = IF(INDEX(M_XU,TRUNG_GIAN!$E$6+1,C34+1)=1,1,0)
    return (1.0 if X.truthy(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,r-52,c),1.0)),1.0)) else 0.0)

def f72(E, r, c):
    # Ô C87, D87, E87, F87, G87, H87, I87, J87, K87, L87
    # = IF(C39="Thương quan",1,0)
    return (1.0 if X.truthy(X.eq(E.v(14,r-48,c),'Thương quan')) else 0.0)

def f73(E, r, c):
    # Ô C88, D88, E88, F88, G88, H88, I88, J88, K88, L88
    # = IF(OR(C39="Tỉ kiên",C39="Kiếp tài"),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(E.v(14,r-49,c),'Tỉ kiên'), X.eq(E.v(14,r-49,c),'Kiếp tài'))) else 0.0)

def f74(E, r, c):
    # Ô C89, D89, E89, F89, G89, H89, I89, J89, K89, L89
    # = IF(C39="Thiên ấn",1,0)
    return (1.0 if X.truthy(X.eq(E.v(14,r-50,c),'Thiên ấn')) else 0.0)

def f75(E, r, c):
    # Ô C90, D90, E90, F90, G90, H90, I90, J90, K90, L90
    # = IF(C39="Thất sát",1,0)
    return (1.0 if X.truthy(X.eq(E.v(14,r-51,c),'Thất sát')) else 0.0)

def f76(E, r, c):
    # Ô C92, D92, E92, F92, G92, H92, I92, J92, K92, L92
    # = TRIM(IF(LEN(C77)>0,"• Thiên khắc địa xung trụ "&C77&" (P1 tr.133; P2 tr.282). ","")&IF(C78=1,"• Xung khắc đề cương (lệnh tháng) (P2 tr.288; P1 tr.202)
    #   . ","")&IF(LEN(C79)>0,"• Hình chi "&C79&" (P2 tr.282 đ.6). ","")&IF(C80=1,"• Gặp kình dương: hao tán của cải, tai nạn (P2 tr.316). ","")&IF(C81=1,"• N
    #   hật can nhập mộ: đề phòng sức khỏe, tính mạng (P1 tr.202). ","")&IF(C82=1,"• Có thiên/nguyệt đức cứu giải (P2 tr.283). ","")&IF(C83=1,"• Gặp thiên ất 
    #   quý nhân: dễ thăng tiến (P2 tr.295). ","")&IF(C84=1,"• Lợi kết hôn (hợp với trụ ngày) (P1 tr.201). ","")&IF(C85=1,"• Lợi kết hôn (sao phối ngẫu đến) (
    #   P1 tr.201). ","")&IF(C86=1,"• Xung cung phối ngẫu: bất lợi vợ/chồng (P1 tr.156). ","")&IF(C87=1,"• Vận thương quan: bất lợi phối ngẫu, con (P2 tr.471,
    #    507). ","")&IF(C88=1,"• Vận tỉ kiếp: bất lợi cha, tài (P2 tr.442). ","")&IF(C90=1,"• Vận thất sát: đề phòng bệnh (nếu thân nhược) (P2 tr.559). ",""))
    #   &IF(LEN(C93)>0," "&C93,"")&" • Can – chi (Hóa giải toàn thư tr.83): "&C102
    return X.cat(X.cat(X.cat(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• Thiên khắc địa xung trụ ',E.v(14,r-15,c)),' (P1 tr.133; P2 tr.282). ') if X.truthy(X.gt(X.LEN(E.v(14,r-15,c)),0.0)) else ''),('• Xung khắc đề cương (lệnh tháng) (P2 tr.288; P1 tr.202). ' if X.truthy(X.eq(E.v(14,r-14,c),1.0)) else '')),(X.cat(X.cat('• Hình chi ',E.v(14,r-13,c)),' (P2 tr.282 đ.6). ') if X.truthy(X.gt(X.LEN(E.v(14,r-13,c)),0.0)) else '')),('• Gặp kình dương: hao tán của cải, tai nạn (P2 tr.316). ' if X.truthy(X.eq(E.v(14,r-12,c),1.0)) else '')),('• Nhật can nhập mộ: đề phòng sức khỏe, tính mạng (P1 tr.202). ' if X.truthy(X.eq(E.v(14,r-11,c),1.0)) else '')),('• Có thiên/nguyệt đức cứu giải (P2 tr.283). ' if X.truthy(X.eq(E.v(14,r-10,c),1.0)) else '')),('• Gặp thiên ất quý nhân: dễ thăng tiến (P2 tr.295). ' if X.truthy(X.eq(E.v(14,r-9,c),1.0)) else '')),('• Lợi kết hôn (hợp với trụ ngày) (P1 tr.201). ' if X.truthy(X.eq(E.v(14,r-8,c),1.0)) else '')),('• Lợi kết hôn (sao phối ngẫu đến) (P1 tr.201). ' if X.truthy(X.eq(E.v(14,r-7,c),1.0)) else '')),('• Xung cung phối ngẫu: bất lợi vợ/chồng (P1 tr.156). ' if X.truthy(X.eq(E.v(14,r-6,c),1.0)) else '')),('• Vận thương quan: bất lợi phối ngẫu, con (P2 tr.471, 507). ' if X.truthy(X.eq(E.v(14,r-5,c),1.0)) else '')),('• Vận tỉ kiếp: bất lợi cha, tài (P2 tr.442). ' if X.truthy(X.eq(E.v(14,r-4,c),1.0)) else '')),('• Vận thất sát: đề phòng bệnh (nếu thân nhược) (P2 tr.559). ' if X.truthy(X.eq(E.v(14,r-2,c),1.0)) else ''))),(X.cat(' ',E.v(14,r+1,c)) if X.truthy(X.gt(X.LEN(E.v(14,r+1,c)),0.0)) else '')),' • Can – chi (Hóa giải toàn thư tr.83): '),E.v(14,r+10,c))

def f77(E, r, c):
    # Ô C93
    # = TRIM(IF(AND(OR(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$5=1),OR(C34=8,C34=9,C34=10),IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),"• Vận Giáp Ất gặp
    #    phương tây, thân lại vượng thì sẽ nổi tiếng (P2 tr.552; UH tr.194: “Vận Giáp Ất nhập Tây phương, thân vượng công danh có thể hứa”). ","")&IF(INT(MOD(
    #   C34+10,12)/3)<>INT(MOD(TRUNG_GIAN!$D$6+10,12)/3),"• [Uyên Hải] Giao vận đổi phương (chuyển giác tiếp mộc) năm "&C37&": lúc đổi vận dễ có biến, người c
    #   ao tuổi đại kỵ (UH tr.138-139). ",""))
    return X.TRIM(X.cat(('• Vận Giáp Ất gặp phương tây, thân lại vượng thì sẽ nổi tiếng (P2 tr.552; UH tr.194: “Vận Giáp Ất nhập Tây phương, thân vượng công danh có thể hứa”). ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,5,5),1.0)), X.OR(X.eq(E.v(14,r-59,c),8.0), X.eq(E.v(14,r-59,c),9.0), X.eq(E.v(14,r-59,c),10.0)), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else ''),(X.cat(X.cat('• [Uyên Hải] Giao vận đổi phương (chuyển giác tiếp mộc) năm ',E.v(14,r-56,c)),': lúc đổi vận dễ có biến, người cao tuổi đại kỵ (UH tr.138-139). ') if X.truthy(X.ne(X.INT(X.div(X.MOD(X.add(E.v(14,r-59,c),10.0), 12.0),3.0)),X.INT(X.div(X.MOD(X.add(E.v(23,6,4),10.0), 12.0),3.0)))) else '')))

def f78(E, r, c):
    # Ô D93, E93, F93, G93, H93, I93, J93, K93, L93
    # = TRIM(IF(AND(OR(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$5=1),OR(D34=8,D34=9,D34=10),IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),"• Vận Giáp Ất gặp
    #    phương tây, thân lại vượng thì sẽ nổi tiếng (P2 tr.552; UH tr.194: “Vận Giáp Ất nhập Tây phương, thân vượng công danh có thể hứa”). ","")&IF(INT(MOD(
    #   D34+10,12)/3)<>INT(MOD(C34+10,12)/3),"• [Uyên Hải] Giao vận đổi phương (chuyển giác tiếp mộc) năm "&D37&": lúc đổi vận dễ có biến, người cao tuổi đại 
    #   kỵ (UH tr.138-139). ",""))
    return X.TRIM(X.cat(('• Vận Giáp Ất gặp phương tây, thân lại vượng thì sẽ nổi tiếng (P2 tr.552; UH tr.194: “Vận Giáp Ất nhập Tây phương, thân vượng công danh có thể hứa”). ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,5,5),1.0)), X.OR(X.eq(E.v(14,r-59,c),8.0), X.eq(E.v(14,r-59,c),9.0), X.eq(E.v(14,r-59,c),10.0)), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else ''),(X.cat(X.cat('• [Uyên Hải] Giao vận đổi phương (chuyển giác tiếp mộc) năm ',E.v(14,r-56,c)),': lúc đổi vận dễ có biến, người cao tuổi đại kỵ (UH tr.138-139). ') if X.truthy(X.ne(X.INT(X.div(X.MOD(X.add(E.v(14,r-59,c),10.0), 12.0),3.0)),X.INT(X.div(X.MOD(X.add(E.v(14,r-59,c-1),10.0), 12.0),3.0)))) else '')))

def f79(E, r, c):
    # Ô C94
    # = "Đại vận rất tốt: "&IF(TRIM(IF(C121="RẤT TỐT",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(D121="RẤT TỐT",D$35&" ("&D$36&"–"&(D$36+9)&") ","")&IF(E121="RẤT
    #    TỐT",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(F121="RẤT TỐT",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&IF(G121="RẤT TỐT",G$35&" ("&G$36&"–"&(G$36+9)&") ","
    #   ")&IF(H121="RẤT TỐT",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(I121="RẤT TỐT",I$35&" ("&I$36&"–"&(I$36+9)&") ","")&IF(J121="RẤT TỐT",J$35&" ("&J$36&"–"&
    #   (J$36+9)&") ",""))="","không có",TRIM(IF(C121="RẤT TỐT",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(D121="RẤT TỐT",D$35&" ("&D$36&"–"&(D$36+9)&") ","")&IF
    #   (E121="RẤT TỐT",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(F121="RẤT TỐT",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&IF(G121="RẤT TỐT",G$35&" ("&G$36&"–"&(G$36
    #   +9)&") ","")&IF(H121="RẤT TỐT",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(I121="RẤT TỐT",I$35&" ("&I$36&"–"&(I$36+9)&") ","")&IF(J121="RẤT TỐT",J$35&" ("
    #   &J$36&"–"&(J$36+9)&") ","")))&"; tốt: "&IF(TRIM(IF(LEFT(C121,3)="TỐT",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(LEFT(D121,3)="TỐT",D$35&" ("&D$36&"–"&(D
    #   $36+9)&") ","")&IF(LEFT(E121,3)="TỐT",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(LEFT(F121,3)="TỐT",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&IF(LEFT(G121,3)=
    #   "TỐT",G$35&" ("&G$36&"–"&(G$36+9)&") ","")&IF(LEFT(H121,3)="TỐT",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(LEFT(I121,3)="TỐT",I$35&" ("&I$36&"–"&(I$36+9
    #   )&") ","")&IF(LEFT(J121,3)="TỐT",J$35&" ("&J$36&"–"&(J$36+9)&") ",""))="","không có",TRIM(IF(LEFT(C121,3)="TỐT",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&I
    #   F(LEFT(D121,3)="TỐT",D$35&" ("&D$36&"–"&(D$36+9)&") ","")&IF(LEFT(E121,3)="TỐT",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(LEFT(F121,3)="TỐT",F$35&" ("&F
    #   $36&"–"&(F$36+9)&") ","")&IF(LEFT(G121,3)="TỐT",G$35&" ("&G$36&"–"&(G$36+9)&") ","")&IF(LEFT(H121,3)="TỐT",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(LEF
    #   T(I121,3)="TỐT",I$35&" ("&I$36&"–"&(I$36+9)&") ","")&IF(LEFT(J121,3)="TỐT",J$35&" ("&J$36&"–"&(J$36+9)&") ","")))&"; xấu: "&IF(TRIM(IF(LEFT(C121,3)="X
    #   ẤU",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(LEFT(D121,3)="XẤU",D$35&" ("&D$36&"–"&(D$36+9)&") ","")&IF(LEFT(E121,3)="XẤU",E$35&" ("&E$36&"–"&(E$36+9)&
    #   ") ","")&IF(LEFT(F121,3)="XẤU",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&IF(LEFT(G121,3)="XẤU",G$35&" ("&G$36&"–"&(G$36+9)&") ","")&IF(LEFT(H121,3)="XẤU",H
    #   $35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(LEFT(I121,3)="XẤU",I$35&" ("&I$36&"–"&(I$36+9)&") ","")&IF(LEFT(J121,3)="XẤU",J$35&" ("&J$36&"–"&(J$36+9)&") ",
    #   ""))="","không có",TRIM(IF(LEFT(C121,3)="XẤU",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(LEFT(D121,3)="XẤU",D$35&" ("&D$36&"–"&(D$36+9)&") ","")&IF(LEFT(
    #   E121,3)="XẤU",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(LEFT(F121,3)="XẤU",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&IF(LEFT(G121,3)="XẤU",G$35&" ("&G$36&"–"
    #   &(G$36+9)&") ","")&IF(LEFT(H121,3)="XẤU",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(LEFT(I121,3)="XẤU",I$35&" ("&I$36&"–"&(I$36+9)&") ","")&IF(LEFT(J121,
    #   3)="XẤU",J$35&" ("&J$36&"–"&(J$36+9)&") ","")))&"; rất xấu: "&IF(TRIM(IF(C121="RẤT XẤU",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(D121="RẤT XẤU",D$35&" 
    #   ("&D$36&"–"&(D$36+9)&") ","")&IF(E121="RẤT XẤU",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(F121="RẤT XẤU",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&IF(G121="R
    #   ẤT XẤU",G$35&" ("&G$36&"–"&(G$36+9)&") ","")&IF(H121="RẤT XẤU",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(I121="RẤT XẤU",I$35&" ("&I$36&"–"&(I$36+9)&") "
    #   ,"")&IF(J121="RẤT XẤU",J$35&" ("&J$36&"–"&(J$36+9)&") ",""))="","không có",TRIM(IF(C121="RẤT XẤU",C$35&" ("&C$36&"–"&(C$36+9)&") ","")&IF(D121="RẤT XẤ
    #   U",D$35&" ("&D$36&"–"&(D$36+9)&") ","")&IF(E121="RẤT XẤU",E$35&" ("&E$36&"–"&(E$36+9)&") ","")&IF(F121="RẤT XẤU",F$35&" ("&F$36&"–"&(F$36+9)&") ","")&
    #   IF(G121="RẤT XẤU",G$35&" ("&G$36&"–"&(G$36+9)&") ","")&IF(H121="RẤT XẤU",H$35&" ("&H$36&"–"&(H$36+9)&") ","")&IF(I121="RẤT XẤU",I$35&" ("&I$36&"–"&(I$
    #   36+9)&") ","")&IF(J121="RẤT XẤU",J$35&" ("&J$36&"–"&(J$36+9)&") ","")))&". Đang hành: "&LEFT(KET_LUAN!$C$163,MIN(LEN(KET_LUAN!$C$163),300))&" …"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đại vận rất tốt: ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c),'RẤT TỐT')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+1),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+2),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+3),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+4),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+5),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+6),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+7),'RẤT TỐT')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c),'RẤT TỐT')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+1),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+2),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+3),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+4),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+5),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+6),'RẤT TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+7),'RẤT TỐT')) else ''))))),'; tốt: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c), 3.0),'TỐT')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+1), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+2), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+3), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+4), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+5), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+6), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+7), 3.0),'TỐT')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c), 3.0),'TỐT')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+1), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+2), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+3), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+4), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+5), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+6), 3.0),'TỐT')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+7), 3.0),'TỐT')) else ''))))),'; xấu: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c), 3.0),'XẤU')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+1), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+2), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+3), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+4), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+5), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+6), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+7), 3.0),'XẤU')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c), 3.0),'XẤU')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+1), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+2), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+3), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+4), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+5), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+6), 3.0),'XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(X.LEFT(E.v(14,r+27,c+7), 3.0),'XẤU')) else ''))))),'; rất xấu: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c),'RẤT XẤU')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+1),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+2),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+3),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+4),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+5),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+6),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+7),'RẤT XẤU')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c),' ('),E.v(14,36,c)),'–'),X.add(E.v(14,36,c),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c),'RẤT XẤU')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+1),' ('),E.v(14,36,c+1)),'–'),X.add(E.v(14,36,c+1),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+1),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+2),' ('),E.v(14,36,c+2)),'–'),X.add(E.v(14,36,c+2),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+2),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+3),' ('),E.v(14,36,c+3)),'–'),X.add(E.v(14,36,c+3),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+3),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+4),' ('),E.v(14,36,c+4)),'–'),X.add(E.v(14,36,c+4),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+4),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+5),' ('),E.v(14,36,c+5)),'–'),X.add(E.v(14,36,c+5),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+5),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+6),' ('),E.v(14,36,c+6)),'–'),X.add(E.v(14,36,c+6),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+6),'RẤT XẤU')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,35,c+7),' ('),E.v(14,36,c+7)),'–'),X.add(E.v(14,36,c+7),9.0)),') ') if X.truthy(X.eq(E.v(14,r+27,c+7),'RẤT XẤU')) else ''))))),'. Đang hành: '),X.LEFT(E.v(3,163,3), X.MIN(X.LEN(E.v(3,163,3)), 300.0))),' …')

def f80(E, r, c):
    # Ô C97, D97, E97, F97, G97, H97, I97, J97, K97, L97
    # = IF(IF(DOI_CHIEU!$C$33=2,C62,C47)="","",IF(OR(IF(DOI_CHIEU!$C$33=2,C62,C47)="dụng",IF(DOI_CHIEU!$C$33=2,C62,C47)="hỷ"),1,IF(IF(DOI_CHIEU!$C$33=2,C62,C4
    #   7)="kỵ",-1,0)))
    return ('' if X.truthy(X.eq((E.v(14,r-35,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-50,c)),'')) else (1.0 if X.truthy(X.OR(X.eq((E.v(14,r-35,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-50,c)),'dụng'), X.eq((E.v(14,r-35,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-50,c)),'hỷ'))) else (X.neg(1.0) if X.truthy(X.eq((E.v(14,r-35,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-50,c)),'kỵ')) else 0.0)))

def f81(E, r, c):
    # Ô C98, D98, E98, F98, G98, H98, I98, J98, K98, L98
    # = IF(IF(DOI_CHIEU!$C$33=2,C65,C50)="","",IF(OR(IF(DOI_CHIEU!$C$33=2,C65,C50)="dụng",IF(DOI_CHIEU!$C$33=2,C65,C50)="hỷ"),1,IF(IF(DOI_CHIEU!$C$33=2,C65,C5
    #   0)="kỵ",-1,0)))
    return ('' if X.truthy(X.eq((E.v(14,r-33,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-48,c)),'')) else (1.0 if X.truthy(X.OR(X.eq((E.v(14,r-33,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-48,c)),'dụng'), X.eq((E.v(14,r-33,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-48,c)),'hỷ'))) else (X.neg(1.0) if X.truthy(X.eq((E.v(14,r-33,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-48,c)),'kỵ')) else 0.0)))

def f82(E, r, c):
    # Ô C99, D99, E99, F99, G99, H99, I99, J99, K99, L99
    # = IF(C42=C43,0,IF(MOD(C42+1,5)=C43,1,IF(MOD(C43+1,5)=C42,2,IF(MOD(C42+2,5)=C43,3,4))))
    return (0.0 if X.truthy(X.eq(E.v(14,r-57,c),E.v(14,r-56,c))) else (1.0 if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-57,c),1.0), 5.0),E.v(14,r-56,c))) else (2.0 if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-56,c),1.0), 5.0),E.v(14,r-57,c))) else (3.0 if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-57,c),2.0), 5.0),E.v(14,r-56,c))) else 4.0))))

def f83(E, r, c):
    # Ô C100, D100, E100, F100, G100, H100, I100, J100, K100, L100
    # = CHOOSE(C99+1,"Can chi cùng hành – lực mạnh nhất","Can sinh chi: can 4 phần, chi 6 phần","Chi sinh can: chi 4 phần, can 6 phần","Can khắc chi (“cái đầu
    #   ”): chi gần như mất lực","Chi khắc can (“tiệt cước”): can gần như mất lực")
    return X.CHOOSE(X.add(E.v(14,r-1,c),1.0), lambda: 'Can chi cùng hành – lực mạnh nhất', lambda: 'Can sinh chi: can 4 phần, chi 6 phần', lambda: 'Chi sinh can: chi 4 phần, can 6 phần', lambda: 'Can khắc chi (“cái đầu”): chi gần như mất lực', lambda: 'Chi khắc can (“tiệt cước”): can gần như mất lực')

def f84(E, r, c):
    # Ô C101, D101, E101, F101, G101, H101, I101, J101, K101, L101
    # = IFERROR(CHOOSE(C99+1,0.5*C97+0.5*C98,0.4*C97+0.6*C98,0.6*C97+0.4*C98,0.5*C97,0.5*C98),0)
    return X.IFERROR(lambda: X.CHOOSE(X.add(E.v(14,r-2,c),1.0), lambda: X.add(X.mul(0.5,E.v(14,r-4,c)),X.mul(0.5,E.v(14,r-3,c))), lambda: X.add(X.mul(0.4,E.v(14,r-4,c)),X.mul(0.6,E.v(14,r-3,c))), lambda: X.add(X.mul(0.6,E.v(14,r-4,c)),X.mul(0.4,E.v(14,r-3,c))), lambda: X.mul(0.5,E.v(14,r-4,c)), lambda: X.mul(0.5,E.v(14,r-3,c))), lambda: 0.0)

def f85(E, r, c):
    # Ô C102, D102, E102, F102, G102, H102, I102, J102, K102, L102
    # = IFERROR(IF(C97*C98=1,IF(C97=1,"Đại cát","Đại hung"),IF(C97*C98=-1,IF(C101>0,"Bình, thiên cát",IF(C101<0,"Bình, thiên hung","Nửa cát nửa hung")),IF(C10
    #   1>0,"Tiểu cát",IF(C101<0,"Tiểu hung","Bình")))),"")
    return X.IFERROR(lambda: (('Đại cát' if X.truthy(X.eq(E.v(14,r-5,c),1.0)) else 'Đại hung') if X.truthy(X.eq(X.mul(E.v(14,r-5,c),E.v(14,r-4,c)),1.0)) else (('Bình, thiên cát' if X.truthy(X.gt(E.v(14,r-1,c),0.0)) else ('Bình, thiên hung' if X.truthy(X.lt(E.v(14,r-1,c),0.0)) else 'Nửa cát nửa hung')) if X.truthy(X.eq(X.mul(E.v(14,r-5,c),E.v(14,r-4,c)),X.neg(1.0))) else ('Tiểu cát' if X.truthy(X.gt(E.v(14,r-1,c),0.0)) else ('Tiểu hung' if X.truthy(X.lt(E.v(14,r-1,c),0.0)) else 'Bình')))), lambda: '')

def f86(E, r, c):
    # Ô C103, D103, E103, F103, G103, H103, I103, J103, K103, L103
    # = IFERROR("5 năm đầu ("&C37&"–"&(C37+4)&"): can "&INDEX(CAN_TEN,C33+1)&" chủ 7 phần – "&CHOOSE(C97+2,"kỵ thần","nhàn thần","dụng/hỷ thần")&IF(C99=4," nh
    #   ưng bị chi khắc (tiệt cước) nên lực rất yếu","")&". 5 năm sau ("&(C37+5)&"–"&(C37+9)&"): chi "&INDEX(CHI_TEN,C34+1)&" chủ 7 phần – "&CHOOSE(C98+2,"kỵ 
    #   thần","nhàn thần","dụng/hỷ thần")&IF(C99=3," nhưng bị can khắc (cái đầu) nên lực rất yếu","")&".","")
    return X.IFERROR(lambda: X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('5 năm đầu (',E.v(14,r-66,c)),'–'),X.add(E.v(14,r-66,c),4.0)),'): can '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(14,r-70,c),1.0))),' chủ 7 phần – '),X.CHOOSE(X.add(E.v(14,r-6,c),2.0), lambda: 'kỵ thần', lambda: 'nhàn thần', lambda: 'dụng/hỷ thần')),(' nhưng bị chi khắc (tiệt cước) nên lực rất yếu' if X.truthy(X.eq(E.v(14,r-4,c),4.0)) else '')),'. 5 năm sau ('),X.add(E.v(14,r-66,c),5.0)),'–'),X.add(E.v(14,r-66,c),9.0)),'): chi '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-69,c),1.0))),' chủ 7 phần – '),X.CHOOSE(X.add(E.v(14,r-5,c),2.0), lambda: 'kỵ thần', lambda: 'nhàn thần', lambda: 'dụng/hỷ thần')),(' nhưng bị can khắc (cái đầu) nên lực rất yếu' if X.truthy(X.eq(E.v(14,r-4,c),3.0)) else '')),'.'), lambda: '')

def f87(E, r, c):
    # Ô C104, D104, E104, F104, G104, H104, I104, J104, K104, L104
    # = IFERROR(0.7*IF(C99=4,0,C97)+0.3*IF(C99=3,0,C98),0)
    return X.IFERROR(lambda: X.add(X.mul(0.7,(0.0 if X.truthy(X.eq(E.v(14,r-5,c),4.0)) else E.v(14,r-7,c))),X.mul(0.3,(0.0 if X.truthy(X.eq(E.v(14,r-5,c),3.0)) else E.v(14,r-6,c)))), lambda: 0.0)

def f88(E, r, c):
    # Ô C105, D105, E105, F105, G105, H105, I105, J105, K105, L105
    # = IFERROR(0.3*IF(C99=4,0,C97)+0.7*IF(C99=3,0,C98),0)
    return X.IFERROR(lambda: X.add(X.mul(0.3,(0.0 if X.truthy(X.eq(E.v(14,r-6,c),4.0)) else E.v(14,r-8,c))),X.mul(0.7,(0.0 if X.truthy(X.eq(E.v(14,r-6,c),3.0)) else E.v(14,r-7,c)))), lambda: 0.0)

def f89(E, r, c):
    # Ô C106, D106, E106, F106, G106, H106, I106, J106, K106, L106
    # = IF(IF(DOI_CHIEU!$C$33=2,C69,C54)="Tốt",1,0)
    return (1.0 if X.truthy(X.eq((E.v(14,r-37,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-52,c)),'Tốt')) else 0.0)

def f90(E, r, c):
    # Ô C109
    # = IF(DUNG_HY_KY!$C$7=0,1,0)
    return (1.0 if X.truthy(X.eq(E.v(8,7,3),0.0)) else 0.0)

def f91(E, r, c):
    # Ô C110
    # = IF(DUNG_HY_KY!$D$7=0,1,0)
    return (1.0 if X.truthy(X.eq(E.v(8,7,4),0.0)) else 0.0)

def f92(E, r, c):
    # Ô C111
    # = DUNG_HY_KY!$C$215
    return E.v(8,215,3)

def f93(E, r, c):
    # Ô C112
    # = IF(AND(OR(TRUNG_GIAN!$C$12="Chính quan",TRUNG_GIAN!$D$12="Chính quan",TRUNG_GIAN!$F$12="Chính quan"),NOT(OR(TRUNG_GIAN!$C$12="Chính ấn",TRUNG_GIAN!$C$
    #   12="Thiên ấn",TRUNG_GIAN!$D$12="Chính ấn",TRUNG_GIAN!$D$12="Thiên ấn",TRUNG_GIAN!$F$12="Chính ấn",TRUNG_GIAN!$F$12="Thiên ấn"))),1,0)
    return (1.0 if X.truthy(X.AND(X.OR(X.eq(E.v(23,12,3),'Chính quan'), X.eq(E.v(23,12,4),'Chính quan'), X.eq(E.v(23,12,6),'Chính quan')), X.NOT(X.OR(X.eq(E.v(23,12,3),'Chính ấn'), X.eq(E.v(23,12,3),'Thiên ấn'), X.eq(E.v(23,12,4),'Chính ấn'), X.eq(E.v(23,12,4),'Thiên ấn'), X.eq(E.v(23,12,6),'Chính ấn'), X.eq(E.v(23,12,6),'Thiên ấn'))))) else 0.0)

def f94(E, r, c):
    # Ô C113
    # = IF($C$111<0,"",INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&IF($C$111=1,", mùa lạnh, cần hỏa sưởi ấm",", mùa nóng
    #   , cần thủy điều hòa"))
    return ('' if X.truthy(X.lt(E.v(14,111,3),0.0)) else X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0)),' sinh tháng '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),(', mùa lạnh, cần hỏa sưởi ấm' if X.truthy(X.eq(E.v(14,111,3),1.0)) else ', mùa nóng, cần thủy điều hòa')))

def f95(E, r, c):
    # Ô C115, D115, E115, F115, G115, H115, I115, J115, K115, L115
    # = IF(AND(DUNG_HY_KY!$C$109>=0,IF(AND(C$39="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GI
    #   AN!$E$7,5)=4)=($C$109=1))),1,IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$109=1)),1,-1)))>0,C$42=DUNG_HY_KY!$C$109),2,IF(AN
    #   D(DUNG_HY_KY!$C$109>=0,IF(AND(C$39="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E
    #   $7,5)=4)=($C$109=1))),1,IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$109=1)),1,-1)))<0,MOD(C$42+2,5)=DUNG_HY_KY!$C$109),-2,
    #   IF(AND(C$39="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$109=1))),1
    #   ,IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$109=1)),1,-1)))))
    return (2.0 if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.gt((X.neg(1.0) if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))) else X.neg(1.0)))),0.0), X.eq(E.v(14,42,c),E.v(8,109,3)))) else (X.neg(2.0) if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.lt((X.neg(1.0) if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))) else X.neg(1.0)))),0.0), X.eq(X.MOD(X.add(E.v(14,42,c),2.0), 5.0),E.v(8,109,3)))) else (X.neg(1.0) if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))) else X.neg(1.0))))))

def f96(E, r, c):
    # Ô C116, D116, E116, F116, G116, H116, I116, J116, K116, L116
    # = IF(AND(DUNG_HY_KY!$C$109>=0,IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$43=$C$
    #   111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$109=1))),1,IF((OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5
    #   )=4)=($C$109=1)),1,-1)))>0,C$43=DUNG_HY_KY!$C$109),2,IF(AND(DUNG_HY_KY!$C$109>=0,IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)
    #   ="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$43=$C$111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$109=1))),1,IF((OR(MOD
    #   (C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$109=1)),1,-1)))<0,MOD(C$43+2,5)=DUNG_HY_KY!$C$109),-2,IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+
    #   1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$43=$C$111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$
    #   E$7,5)=4)=($C$109=1))),1,IF((OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$109=1)),1,-1)))))
    return (2.0 if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.gt((X.neg(1.0) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))) else X.neg(1.0)))),0.0), X.eq(E.v(14,43,c),E.v(8,109,3)))) else (X.neg(2.0) if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.lt((X.neg(1.0) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))) else X.neg(1.0)))),0.0), X.eq(X.MOD(X.add(E.v(14,43,c),2.0), 5.0),E.v(8,109,3)))) else (X.neg(1.0) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,109,3),1.0))) else X.neg(1.0))))))

def f97(E, r, c):
    # Ô C117, D117, E117, F117, G117, H117, I117, J117, K117, L117
    # = IF(AND(DUNG_HY_KY!$D$109>=0,IF(AND(C$39="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GI
    #   AN!$E$7,5)=4)=($C$110=1))),1,IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$110=1)),1,-1)))>0,C$42=DUNG_HY_KY!$D$109),2,IF(AN
    #   D(DUNG_HY_KY!$D$109>=0,IF(AND(C$39="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E
    #   $7,5)=4)=($C$110=1))),1,IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$110=1)),1,-1)))<0,MOD(C$42+2,5)=DUNG_HY_KY!$D$109),-2,
    #   IF(AND(C$39="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$110=1))),1
    #   ,IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=($C$110=1)),1,-1)))))
    return (2.0 if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.gt((X.neg(1.0) if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))) else X.neg(1.0)))),0.0), X.eq(E.v(14,42,c),E.v(8,109,4)))) else (X.neg(2.0) if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.lt((X.neg(1.0) if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))) else X.neg(1.0)))),0.0), X.eq(X.MOD(X.add(E.v(14,42,c),2.0), 5.0),E.v(8,109,4)))) else (X.neg(1.0) if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))) else X.neg(1.0))))))

def f98(E, r, c):
    # Ô C118, D118, E118, F118, G118, H118, I118, J118, K118, L118
    # = IF(AND(DUNG_HY_KY!$D$109>=0,IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$43=$C$
    #   111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$110=1))),1,IF((OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5
    #   )=4)=($C$110=1)),1,-1)))>0,C$43=DUNG_HY_KY!$D$109),2,IF(AND(DUNG_HY_KY!$D$109>=0,IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)
    #   ="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$43=$C$111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$110=1))),1,IF((OR(MOD
    #   (C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$110=1)),1,-1)))<0,MOD(C$43+2,5)=DUNG_HY_KY!$D$109),-2,IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+
    #   1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)="Thương quan",$C$112=1),-1,IF(AND($C$111>=0,C$43=$C$111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$
    #   E$7,5)=4)=($C$110=1))),1,IF((OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=($C$110=1)),1,-1)))))
    return (2.0 if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.gt((X.neg(1.0) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))) else X.neg(1.0)))),0.0), X.eq(E.v(14,43,c),E.v(8,109,4)))) else (X.neg(2.0) if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.lt((X.neg(1.0) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))) else X.neg(1.0)))),0.0), X.eq(X.MOD(X.add(E.v(14,43,c),2.0), 5.0),E.v(8,109,4)))) else (X.neg(1.0) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else (1.0 if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))))) else (1.0 if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq(E.v(14,110,3),1.0))) else X.neg(1.0))))))

def f99(E, r, c):
    # Ô C119, D119, E119, F119, G119, H119, I119, J119, K119, L119, C120, D120 … (20 ô)
    # = IF(DOI_CHIEU!$C$33=2,C117,C115)
    return (E.v(14,r-2,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-4,c))

def f100(E, r, c):
    # Ô C121, D121, E121, F121, G121, H121, I121, J121, K121, L121
    # = IF(C119=0,IF((C119>0)=(C120>0),CHOOSE(IF(ABS(C119)>=ABS(C120),C119,C120)+3,"RẤT XẤU","XẤU","BÌNH THƯỜNG","TỐT","RẤT TỐT"),CHOOSE(C119+3,"RẤT XẤU","XẤU
    #   ","BÌNH THƯỜNG","TỐT","RẤT TỐT")),IF(C119>0,CHOOSE(C137+1,"TỐT VỪA","TỐT","RẤT TỐT"),CHOOSE(C137+1,"XẤU VỪA","XẤU","RẤT XẤU")))
    return ((X.CHOOSE(X.add((E.v(14,r-2,c) if X.truthy(X.ge(X.ABS(E.v(14,r-2,c)),X.ABS(E.v(14,r-1,c)))) else E.v(14,r-1,c)),3.0), lambda: 'RẤT XẤU', lambda: 'XẤU', lambda: 'BÌNH THƯỜNG', lambda: 'TỐT', lambda: 'RẤT TỐT') if X.truthy(X.eq(X.gt(E.v(14,r-2,c),0.0),X.gt(E.v(14,r-1,c),0.0))) else X.CHOOSE(X.add(E.v(14,r-2,c),3.0), lambda: 'RẤT XẤU', lambda: 'XẤU', lambda: 'BÌNH THƯỜNG', lambda: 'TỐT', lambda: 'RẤT TỐT')) if X.truthy(X.eq(E.v(14,r-2,c),0.0)) else (X.CHOOSE(X.add(E.v(14,r+16,c),1.0), lambda: 'TỐT VỪA', lambda: 'TỐT', lambda: 'RẤT TỐT') if X.truthy(X.gt(E.v(14,r-2,c),0.0)) else X.CHOOSE(X.add(E.v(14,r+16,c),1.0), lambda: 'XẤU VỪA', lambda: 'XẤU', lambda: 'RẤT XẤU')))

def f101(E, r, c):
    # Ô C122, D122, E122, F122, G122, H122, I122, J122, K122, L122
    # = "cả 10 năm "&LOWER(C121)
    return X.cat('cả 10 năm ',X.LOWER(E.v(14,r-1,c)))

def f102(E, r, c):
    # Ô C123, D123, E123, F123, G123, H123, I123, J123, K123, L123
    # = "Can "&INDEX(CAN_TEN,C$33+1)&" là "&C$39&" ("&INDEX(HANH_TEN,C$42+1)&")"&" – "&IF((OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=(IF
    #   (DOI_CHIEU!$C$33=2,$C$110,$C$109)=1)),"thuận","nghịch")&" với thân "&IF(IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)=1,"nhược","vượng")&IF(AND($C$111>=0,C$42=$
    #   C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=(IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)=1))),"; nhưng "&INDEX(HANH_TEN,C$42+1)&
    #   " là hành điều hầu cần thiết ("&$C$113&") nên chuyển thành tốt","")&IF(AND(C$39="Thương quan",$C$112=1),"; nhưng là thương quan gặp chính quan lộ tron
    #   g trụ, không có ấn hóa giải nên xấu","")&IF(C119=2,"; hành "&INDEX(HANH_TEN,C$42+1)&" là dụng thần lên ngôi nên rất tốt",IF(C119=-2,"; hành "&INDEX(HA
    #   NH_TEN,C$42+1)&" khắc dụng thần "&INDEX(HANH_TEN,DUNG_HY_KY!$C$157+1)&" nên rất xấu",""))&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(14,33,c),1.0))),' là '),E.v(14,39,c)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,42,c),1.0))),')'),' – '),('thuận' if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0))) else 'nghịch')),' với thân '),('nhược' if X.truthy(X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0)) else 'vượng')),(X.cat(X.cat(X.cat(X.cat('; nhưng ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,42,c),1.0))),' là hành điều hầu cần thiết ('),E.v(14,113,3)),') nên chuyển thành tốt') if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0))))) else '')),('; nhưng là thương quan gặp chính quan lộ trong trụ, không có ấn hóa giải nên xấu' if X.truthy(X.AND(X.eq(E.v(14,39,c),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else '')),(X.cat(X.cat('; hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,42,c),1.0))),' là dụng thần lên ngôi nên rất tốt') if X.truthy(X.eq(E.v(14,r-4,c),2.0)) else (X.cat(X.cat(X.cat(X.cat('; hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,42,c),1.0))),' khắc dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,157,3),1.0))),' nên rất xấu') if X.truthy(X.eq(E.v(14,r-4,c),X.neg(2.0))) else ''))),'.')

def f103(E, r, c):
    # Ô C124, D124, E124, F124, G124, H124, I124, J124, K124, L124
    # = "Chi "&INDEX(CHI_TEN,C$34+1)&" (bản khí "&INDEX(CAN_TEN,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)&" là "&INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C
    #   $34+1)+1)&", hành "&INDEX(HANH_TEN,C$43+1)&")"&" – "&IF((OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=(IF(DOI_CHIEU!$C$33=2,$C$110,
    #   $C$109)=1)),"thuận","nghịch")&" với thân "&IF(IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)=1,"nhược","vượng")&IF(AND($C$111>=0,C$43=$C$111,NOT(OR(MOD(C$43-TRUN
    #   G_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=(IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)=1))),"; nhưng "&INDEX(HANH_TEN,C$43+1)&" là hành điều hầu cần thi
    #   ết ("&$C$113&") nên chuyển thành tốt","")&IF(AND(INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)="Thương quan",$C$112=1),"; nhưng là th
    #   ương quan gặp chính quan lộ trong trụ, không có ấn hóa giải nên xấu","")&IF(C120=2,"; hành "&INDEX(HANH_TEN,C$43+1)&" là dụng thần lên ngôi nên rất tố
    #   t",IF(C120=-2,"; hành "&INDEX(HANH_TEN,C$43+1)&" khắc dụng thần "&INDEX(HANH_TEN,DUNG_HY_KY!$C$157+1)&" nên rất xấu",""))&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chi ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,34,c),1.0))),' (bản khí '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0))),' là '),X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0))),', hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,43,c),1.0))),')'),' – '),('thuận' if X.truthy(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0))) else 'nghịch')),' với thân '),('nhược' if X.truthy(X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0)) else 'vượng')),(X.cat(X.cat(X.cat(X.cat('; nhưng ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,43,c),1.0))),' là hành điều hầu cần thiết ('),E.v(14,113,3)),') nên chuyển thành tốt') if X.truthy(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0))))) else '')),('; nhưng là thương quan gặp chính quan lộ trong trụ, không có ấn hóa giải nên xấu' if X.truthy(X.AND(X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan'), X.eq(E.v(14,112,3),1.0))) else '')),(X.cat(X.cat('; hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,43,c),1.0))),' là dụng thần lên ngôi nên rất tốt') if X.truthy(X.eq(E.v(14,r-4,c),2.0)) else (X.cat(X.cat(X.cat(X.cat('; hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,43,c),1.0))),' khắc dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,157,3),1.0))),' nên rất xấu') if X.truthy(X.eq(E.v(14,r-4,c),X.neg(2.0))) else ''))),'.')

def f104(E, r, c):
    # Ô C125, D125, E125, F125, G125, H125, I125, J125, K125, L125
    # = "P1 tr.124–125: thân nhược gặp vận sinh trợ, thân vượng gặp vận khắc–tiết–hao là thuận; dụng thần lên ngôi là tốt nhất, vận khắc dụng thần là xấu nhất
    #   "&IF(OR(AND($C$111>=0,C$42=$C$111,NOT(OR(MOD(C$42-TRUNG_GIAN!$E$7,5)=0,MOD(C$42-TRUNG_GIAN!$E$7,5)=4)=(IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)=1))),AND($C
    #   $111>=0,C$43=$C$111,NOT(OR(MOD(C$43-TRUNG_GIAN!$E$7,5)=0,MOD(C$43-TRUNG_GIAN!$E$7,5)=4)=(IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)=1)))),"; P1 tr.76–79: sin
    #   h mùa lạnh cần hỏa sưởi ấm, sinh mùa nóng cần thủy điều hòa","")&IF(AND($C$112=1,OR(C$39="Thương quan",INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$
    #   D$29,C$34+1)+1)="Thương quan")),"; P2 tr.541–542: thương quan gặp quan tinh là điều kị nhất","")&"; 5 năm đầu theo can, 5 năm sau theo chi (tham khảo 
    #   Hóa giải toàn thư tr.68)"
    return X.cat(X.cat(X.cat('P1 tr.124–125: thân nhược gặp vận sinh trợ, thân vượng gặp vận khắc–tiết–hao là thuận; dụng thần lên ngôi là tốt nhất, vận khắc dụng thần là xấu nhất',('; P1 tr.76–79: sinh mùa lạnh cần hỏa sưởi ấm, sinh mùa nóng cần thủy điều hòa' if X.truthy(X.OR(X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,42,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,42,c),E.v(23,7,5)), 5.0),4.0)),X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0)))), X.AND(X.ge(E.v(14,111,3),0.0), X.eq(E.v(14,43,c),E.v(14,111,3)), X.NOT(X.eq(X.OR(X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),0.0), X.eq(X.MOD(X.sub(E.v(14,43,c),E.v(23,7,5)), 5.0),4.0)),X.eq((E.v(14,110,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,109,3)),1.0)))))) else '')),('; P2 tr.541–542: thương quan gặp quan tinh là điều kị nhất' if X.truthy(X.AND(X.eq(E.v(14,112,3),1.0), X.OR(X.eq(E.v(14,39,c),'Thương quan'), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan')))) else '')),'; 5 năm đầu theo can, 5 năm sau theo chi (tham khảo Hóa giải toàn thư tr.68)')

def f105(E, r, c):
    # Ô C126, D126, E126, F126, G126, H126, I126, J126
    # = TAI_LOC!C$89&" – "&TAI_LOC!C$90
    return X.cat(X.cat(E.v(19,89,c),' – '),E.v(19,90,c))

def f106(E, r, c):
    # Ô C127, D127, E127, F127, G127, H127, I127, J127, K127, L127
    # = IF(C119>0,"Tốt","Xấu")&" – "&IF(AND($C$112=1,OR(C$39="Thương quan",INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,C$34+1)+1)="Thương quan")),"thư
    #   ơng quan gặp quan nên dễ mất chức, thị phi","theo xu thế chung của vận")
    return X.cat(X.cat(('Tốt' if X.truthy(X.gt(E.v(14,r-8,c),0.0)) else 'Xấu'),' – '),('thương quan gặp quan nên dễ mất chức, thị phi' if X.truthy(X.AND(X.eq(E.v(14,112,3),1.0), X.OR(X.eq(E.v(14,39,c),'Thương quan'), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,34,c),1.0)),1.0)),'Thương quan')))) else 'theo xu thế chung của vận'))

def f107(E, r, c):
    # Ô C130, D130, E130, F130, G130, H130, I130, J130, K130, L130
    # = IF(AND(INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$34+1)>=0,OR(TRUNG_GIAN!$C$7=INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$34+1),TRUNG_GIAN!$D$7=INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$
    #   34+1),TRUNG_GIAN!$E$7=INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$34+1),TRUNG_GIAN!$F$7=INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$34+1),C$42=INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$34
    #   +1))),INDEX(M_LH,TRUNG_GIAN!$C$6+1,C$34+1),IF(AND(INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$34+1)>=0,OR(TRUNG_GIAN!$C$7=INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$34+1),TRUN
    #   G_GIAN!$D$7=INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$34+1),TRUNG_GIAN!$E$7=INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$34+1),TRUNG_GIAN!$F$7=INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$3
    #   4+1),C$42=INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$34+1))),INDEX(M_LH,TRUNG_GIAN!$D$6+1,C$34+1),IF(AND(INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1)>=0,OR(TRUNG_GIAN!$C$
    #   7=INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1),TRUNG_GIAN!$D$7=INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1),TRUNG_GIAN!$E$7=INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1),TRUNG
    #   _GIAN!$F$7=INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1),C$42=INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1))),INDEX(M_LH,TRUNG_GIAN!$E$6+1,C$34+1),IF(AND(NOT(NHAP_LIEU!
    #   $C$9="Không rõ giờ"),INDEX(M_LH,TRUNG_GIAN!$F$6+1,C$34+1)>=0,OR(TRUNG_GIAN!$C$7=INDEX(M_LH,TRUNG_GIAN!$F$6+1,C$34+1),TRUNG_GIAN!$D$7=INDEX(M_LH,TRUNG_
    #   GIAN!$F$6+1,C$34+1),TRUNG_GIAN!$E$7=INDEX(M_LH,TRUNG_GIAN!$F$6+1,C$34+1),TRUNG_GIAN!$F$7=INDEX(M_LH,TRUNG_GIAN!$F$6+1,C$34+1),C$42=INDEX(M_LH,TRUNG_GI
    #   AN!$F$6+1,C$34+1))),INDEX(M_LH,TRUNG_GIAN!$F$6+1,C$34+1),-1))))
    return (X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0)) if X.truthy(X.AND(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0)),0.0), X.OR(X.eq(E.v(23,7,3),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,4),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,5),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,6),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(14,42,c),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,3),1.0), X.add(E.v(14,34,c),1.0)))))) else (X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0)) if X.truthy(X.AND(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0)),0.0), X.OR(X.eq(E.v(23,7,3),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,4),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,5),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,6),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(14,42,c),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,4),1.0), X.add(E.v(14,34,c),1.0)))))) else (X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0)) if X.truthy(X.AND(X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0)),0.0), X.OR(X.eq(E.v(23,7,3),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,4),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,5),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,6),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(14,42,c),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,5),1.0), X.add(E.v(14,34,c),1.0)))))) else (X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0)) if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ge(X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0)),0.0), X.OR(X.eq(E.v(23,7,3),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,4),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,5),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(23,7,6),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0))), X.eq(E.v(14,42,c),X.INDEX(E.rg(24,32,16,43,27), X.add(E.v(23,6,6),1.0), X.add(E.v(14,34,c),1.0)))))) else X.neg(1.0)))))

def f108(E, r, c):
    # Ô C131, D131, E131, F131, G131, H131, I131, J131, K131, L131
    # = IF(C130>=0,C130,C43)
    return (E.v(14,r-1,c) if X.truthy(X.ge(E.v(14,r-1,c),0.0)) else E.v(14,r-88,c))

def f109(E, r, c):
    # Ô C132, D132, E132, F132, G132, H132, I132, J132, K132, L132
    # = IF(DOI_CHIEU!$C$33=2,IF(OR(AND(DUNG_HY_KY!$D$109>=0,C131=DUNG_HY_KY!$D$109),AND(DUNG_HY_KY!$D$110>=0,C131=DUNG_HY_KY!$D$110)),1,IF(AND(DUNG_HY_KY!$D$1
    #   11>=0,C131=DUNG_HY_KY!$D$111),-1,0)),IF(OR(AND(DUNG_HY_KY!$C$109>=0,C131=DUNG_HY_KY!$C$109),AND(DUNG_HY_KY!$C$110>=0,C131=DUNG_HY_KY!$C$110)),1,IF(AND
    #   (DUNG_HY_KY!$C$111>=0,C131=DUNG_HY_KY!$C$111),-1,0)))
    return ((1.0 if X.truthy(X.OR(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(E.v(14,r-1,c),E.v(8,109,4))), X.AND(X.ge(E.v(8,110,4),0.0), X.eq(E.v(14,r-1,c),E.v(8,110,4))))) else (X.neg(1.0) if X.truthy(X.AND(X.ge(E.v(8,111,4),0.0), X.eq(E.v(14,r-1,c),E.v(8,111,4)))) else 0.0)) if X.truthy(X.eq(E.v(7,33,3),2.0)) else (1.0 if X.truthy(X.OR(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(E.v(14,r-1,c),E.v(8,109,3))), X.AND(X.ge(E.v(8,110,3),0.0), X.eq(E.v(14,r-1,c),E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.AND(X.ge(E.v(8,111,3),0.0), X.eq(E.v(14,r-1,c),E.v(8,111,3)))) else 0.0)))

def f110(E, r, c):
    # Ô C133, D133, E133, F133, G133, H133, I133, J133, K133, L133
    # = IF(LEN(IF(DOI_CHIEU!$C$33=2,C67,C52))>0,-1,IF(LEN(IF(DOI_CHIEU!$C$33=2,C66,C51))>0,1,0))
    return (X.neg(1.0) if X.truthy(X.gt(X.LEN((E.v(14,r-66,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-81,c))),0.0)) else (1.0 if X.truthy(X.gt(X.LEN((E.v(14,r-67,c) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-82,c))),0.0)) else 0.0))

def f111(E, r, c):
    # Ô C134, D134, E134, F134, G134, H134, I134, J134, K134, L134
    # = IF(C133<>0,C133,C132)
    return (E.v(14,r-1,c) if X.truthy(X.ne(E.v(14,r-1,c),0.0)) else E.v(14,r-2,c))

def f112(E, r, c):
    # Ô C135, D135, E135, F135, G135, H135, I135, J135, K135, L135
    # = IF(C119=0,"",IF(OR(C139=1,C139=3,C139=4,AND(C134<>0,C134=-SIGN(C119))),1,0))
    return ('' if X.truthy(X.eq(E.v(14,r-16,c),0.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(14,r+4,c),1.0), X.eq(E.v(14,r+4,c),3.0), X.eq(E.v(14,r+4,c),4.0), X.AND(X.ne(E.v(14,r-1,c),0.0), X.eq(E.v(14,r-1,c),X.neg(X.SIGN(E.v(14,r-16,c))))))) else 0.0))

def f113(E, r, c):
    # Ô C136, D136, E136, F136, G136, H136, I136, J136, K136, L136
    # = IF(C119=0,"",IF(C135=1,0,IF(OR(C139=0,C139=2,AND(C134<>0,C134=SIGN(C119))),1,0)))
    return ('' if X.truthy(X.eq(E.v(14,r-17,c),0.0)) else (0.0 if X.truthy(X.eq(E.v(14,r-1,c),1.0)) else (1.0 if X.truthy(X.OR(X.eq(E.v(14,r+3,c),0.0), X.eq(E.v(14,r+3,c),2.0), X.AND(X.ne(E.v(14,r-2,c),0.0), X.eq(E.v(14,r-2,c),X.SIGN(E.v(14,r-17,c)))))) else 0.0)))

def f114(E, r, c):
    # Ô C137, D137, E137, F137, G137, H137, I137, J137, K137, L137
    # = IF(C119=0,"",IF(C135=1,MAX(0,IF(ABS(C119)>=2,2,1)-1),IF(C136=1,MIN(2,IF(ABS(C119)>=2,2,1)+1),IF(ABS(C119)>=2,2,1))))
    return ('' if X.truthy(X.eq(E.v(14,r-18,c),0.0)) else (X.MAX(0.0, X.sub((2.0 if X.truthy(X.ge(X.ABS(E.v(14,r-18,c)),2.0)) else 1.0),1.0)) if X.truthy(X.eq(E.v(14,r-2,c),1.0)) else (X.MIN(2.0, X.add((2.0 if X.truthy(X.ge(X.ABS(E.v(14,r-18,c)),2.0)) else 1.0),1.0)) if X.truthy(X.eq(E.v(14,r-1,c),1.0)) else (2.0 if X.truthy(X.ge(X.ABS(E.v(14,r-18,c)),2.0)) else 1.0))))

def f115(E, r, c):
    # Ô C138, D138, E138, F138, G138, H138, I138, J138, K138, L138
    # = IF(C119=0,C121&" – can vận nhàn thần, xét theo chi vận (P1 tr.123).",C121&" – can "&INDEX(CAN_TEN,C$33+1)&" ("&C$39&", hành "&INDEX(HANH_TEN,C$42+1)&"
    #   ) làm chủ cả 10 năm; chi "&INDEX(CHI_TEN,C$34+1)&IF(C130>=0,", hợp hóa thành "&INDEX(HANH_TEN,C130+1),", hành "&INDEX(HANH_TEN,C$43+1))&IF(C134>0," – 
    #   cùng chiều can",IF(C134<0," – trái chiều can",""))&IF(C133=-1,"; xung mất chữ mang dụng/hỷ",IF(C133=1,"; xung mất chữ mang kỵ",""))&"; "&CHOOSE(C139+1
    #   ,"can chi cùng hành","can sinh chi (can xì hơi)","chi sinh can (sinh phù)","can khắc chi (can hao)","chi khắc can (áp chế can)")&" → "&IF(C135=1,"cái 
    #   "&IF(C119>0,"tốt","xấu")&" giảm một mức",IF(C136=1,"cái "&IF(C119>0,"tốt","xấu")&" tăng thêm","giữ nguyên mức"))&" (P1 tr.122–124; P2 tr.282).")
    return (X.cat(E.v(14,r-17,c),' – can vận nhàn thần, xét theo chi vận (P1 tr.123).') if X.truthy(X.eq(E.v(14,r-19,c),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-17,c),' – can '),X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(14,33,c),1.0))),' ('),E.v(14,39,c)),', hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,42,c),1.0))),') làm chủ cả 10 năm; chi '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,34,c),1.0))),(X.cat(', hợp hóa thành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,r-8,c),1.0))) if X.truthy(X.ge(E.v(14,r-8,c),0.0)) else X.cat(', hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(14,43,c),1.0))))),(' – cùng chiều can' if X.truthy(X.gt(E.v(14,r-4,c),0.0)) else (' – trái chiều can' if X.truthy(X.lt(E.v(14,r-4,c),0.0)) else ''))),('; xung mất chữ mang dụng/hỷ' if X.truthy(X.eq(E.v(14,r-5,c),X.neg(1.0))) else ('; xung mất chữ mang kỵ' if X.truthy(X.eq(E.v(14,r-5,c),1.0)) else ''))),'; '),X.CHOOSE(X.add(E.v(14,r+1,c),1.0), lambda: 'can chi cùng hành', lambda: 'can sinh chi (can xì hơi)', lambda: 'chi sinh can (sinh phù)', lambda: 'can khắc chi (can hao)', lambda: 'chi khắc can (áp chế can)')),' → '),(X.cat(X.cat('cái ',('tốt' if X.truthy(X.gt(E.v(14,r-19,c),0.0)) else 'xấu')),' giảm một mức') if X.truthy(X.eq(E.v(14,r-3,c),1.0)) else (X.cat(X.cat('cái ',('tốt' if X.truthy(X.gt(E.v(14,r-19,c),0.0)) else 'xấu')),' tăng thêm') if X.truthy(X.eq(E.v(14,r-2,c),1.0)) else 'giữ nguyên mức'))),' (P1 tr.122–124; P2 tr.282).'))

def f116(E, r, c):
    # Ô C139, D139, E139, F139, G139, H139, I139, J139, K139, L139
    # = IF(C42=C131,0,IF(MOD(C42+1,5)=C131,1,IF(MOD(C131+1,5)=C42,2,IF(MOD(C42+2,5)=C131,3,4))))
    return (0.0 if X.truthy(X.eq(E.v(14,r-97,c),E.v(14,r-8,c))) else (1.0 if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-97,c),1.0), 5.0),E.v(14,r-8,c))) else (2.0 if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-8,c),1.0), 5.0),E.v(14,r-97,c))) else (3.0 if X.truthy(X.eq(X.MOD(X.add(E.v(14,r-97,c),2.0), 5.0),E.v(14,r-8,c))) else 4.0))))

CELLS = {
    'C5': (f0, 5, 3),
    'C6': (f1, 6, 3),
    'C7': (f2, 7, 3),
    'C8': (f3, 8, 3),
    'C9': (f4, 9, 3),
    'C10': (f5, 10, 3),
    'C11': (f6, 11, 3),
    'C12': (f7, 12, 3),
    'C13': (f8, 13, 3),
    'C14': (f9, 14, 3),
    'C15': (f10, 15, 3),
    'C16': (f11, 16, 3),
    'C17': (f12, 17, 3),
    'C18': (f13, 18, 3),
    'C19': (f14, 19, 3),
    'C20': (f15, 20, 3),
    'D20': (f16, 20, 4),
    'C21': (f17, 21, 3),
    'C22': (f18, 22, 3),
    'D22': (f19, 22, 4),
    'C23': (f20, 23, 3),
    'C24': (f21, 24, 3),
    'D24': (f22, 24, 4),
    'C25': (f23, 25, 3),
    'C26': (f24, 26, 3),
    'C27': (f25, 27, 3),
    'C28': (f26, 28, 3),
    'C32': (f27, 32, 3),
    'D32': (f27, 32, 4),
    'E32': (f27, 32, 5),
    'F32': (f27, 32, 6),
    'G32': (f27, 32, 7),
    'H32': (f27, 32, 8),
    'I32': (f27, 32, 9),
    'J32': (f27, 32, 10),
    'K32': (f27, 32, 11),
    'L32': (f27, 32, 12),
    'C33': (f28, 33, 3),
    'D33': (f28, 33, 4),
    'E33': (f28, 33, 5),
    'F33': (f28, 33, 6),
    'G33': (f28, 33, 7),
    'H33': (f28, 33, 8),
    'I33': (f28, 33, 9),
    'J33': (f28, 33, 10),
    'K33': (f28, 33, 11),
    'L33': (f28, 33, 12),
    'C34': (f29, 34, 3),
    'D34': (f29, 34, 4),
    'E34': (f29, 34, 5),
    'F34': (f29, 34, 6),
    'G34': (f29, 34, 7),
    'H34': (f29, 34, 8),
    'I34': (f29, 34, 9),
    'J34': (f29, 34, 10),
    'K34': (f29, 34, 11),
    'L34': (f29, 34, 12),
    'C35': (f30, 35, 3),
    'D35': (f30, 35, 4),
    'E35': (f30, 35, 5),
    'F35': (f30, 35, 6),
    'G35': (f30, 35, 7),
    'H35': (f30, 35, 8),
    'I35': (f30, 35, 9),
    'J35': (f30, 35, 10),
    'K35': (f30, 35, 11),
    'L35': (f30, 35, 12),
    'C36': (f31, 36, 3),
    'D36': (f31, 36, 4),
    'E36': (f31, 36, 5),
    'F36': (f31, 36, 6),
    'G36': (f31, 36, 7),
    'H36': (f31, 36, 8),
    'I36': (f31, 36, 9),
    'J36': (f31, 36, 10),
    'K36': (f31, 36, 11),
    'L36': (f31, 36, 12),
    'C37': (f32, 37, 3),
    'D37': (f32, 37, 4),
    'E37': (f32, 37, 5),
    'F37': (f32, 37, 6),
    'G37': (f32, 37, 7),
    'H37': (f32, 37, 8),
    'I37': (f32, 37, 9),
    'J37': (f32, 37, 10),
    'K37': (f32, 37, 11),
    'L37': (f32, 37, 12),
    'C38': (f33, 38, 3),
    'D38': (f33, 38, 4),
    'E38': (f33, 38, 5),
    'F38': (f33, 38, 6),
    'G38': (f33, 38, 7),
    'H38': (f33, 38, 8),
    'I38': (f33, 38, 9),
    'J38': (f33, 38, 10),
    'K38': (f33, 38, 11),
    'L38': (f33, 38, 12),
    'C39': (f34, 39, 3),
    'D39': (f34, 39, 4),
    'E39': (f34, 39, 5),
    'F39': (f34, 39, 6),
    'G39': (f34, 39, 7),
    'H39': (f34, 39, 8),
    'I39': (f34, 39, 9),
    'J39': (f34, 39, 10),
    'K39': (f34, 39, 11),
    'L39': (f34, 39, 12),
    'C40': (f35, 40, 3),
    'D40': (f35, 40, 4),
    'E40': (f35, 40, 5),
    'F40': (f35, 40, 6),
    'G40': (f35, 40, 7),
    'H40': (f35, 40, 8),
    'I40': (f35, 40, 9),
    'J40': (f35, 40, 10),
    'K40': (f35, 40, 11),
    'L40': (f35, 40, 12),
    'C41': (f36, 41, 3),
    'D41': (f36, 41, 4),
    'E41': (f36, 41, 5),
    'F41': (f36, 41, 6),
    'G41': (f36, 41, 7),
    'H41': (f36, 41, 8),
    'I41': (f36, 41, 9),
    'J41': (f36, 41, 10),
    'K41': (f36, 41, 11),
    'L41': (f36, 41, 12),
    'C42': (f37, 42, 3),
    'D42': (f37, 42, 4),
    'E42': (f37, 42, 5),
    'F42': (f37, 42, 6),
    'G42': (f37, 42, 7),
    'H42': (f37, 42, 8),
    'I42': (f37, 42, 9),
    'J42': (f37, 42, 10),
    'K42': (f37, 42, 11),
    'L42': (f37, 42, 12),
    'C43': (f38, 43, 3),
    'D43': (f38, 43, 4),
    'E43': (f38, 43, 5),
    'F43': (f38, 43, 6),
    'G43': (f38, 43, 7),
    'H43': (f38, 43, 8),
    'I43': (f38, 43, 9),
    'J43': (f38, 43, 10),
    'K43': (f38, 43, 11),
    'L43': (f38, 43, 12),
    'C44': (f39, 44, 3),
    'D44': (f39, 44, 4),
    'E44': (f39, 44, 5),
    'F44': (f39, 44, 6),
    'G44': (f39, 44, 7),
    'H44': (f39, 44, 8),
    'I44': (f39, 44, 9),
    'J44': (f39, 44, 10),
    'K44': (f39, 44, 11),
    'L44': (f39, 44, 12),
    'C47': (f40, 47, 3),
    'D47': (f40, 47, 4),
    'E47': (f40, 47, 5),
    'F47': (f40, 47, 6),
    'G47': (f40, 47, 7),
    'H47': (f40, 47, 8),
    'I47': (f40, 47, 9),
    'J47': (f40, 47, 10),
    'K47': (f40, 47, 11),
    'L47': (f40, 47, 12),
    'C48': (f41, 48, 3),
    'D48': (f41, 48, 4),
    'E48': (f41, 48, 5),
    'F48': (f41, 48, 6),
    'G48': (f41, 48, 7),
    'H48': (f41, 48, 8),
    'I48': (f41, 48, 9),
    'J48': (f41, 48, 10),
    'K48': (f41, 48, 11),
    'L48': (f41, 48, 12),
    'C49': (f42, 49, 3),
    'D49': (f42, 49, 4),
    'E49': (f42, 49, 5),
    'F49': (f42, 49, 6),
    'G49': (f42, 49, 7),
    'H49': (f42, 49, 8),
    'I49': (f42, 49, 9),
    'J49': (f42, 49, 10),
    'K49': (f42, 49, 11),
    'L49': (f42, 49, 12),
    'C50': (f43, 50, 3),
    'D50': (f43, 50, 4),
    'E50': (f43, 50, 5),
    'F50': (f43, 50, 6),
    'G50': (f43, 50, 7),
    'H50': (f43, 50, 8),
    'I50': (f43, 50, 9),
    'J50': (f43, 50, 10),
    'K50': (f43, 50, 11),
    'L50': (f43, 50, 12),
    'C51': (f44, 51, 3),
    'D51': (f44, 51, 4),
    'E51': (f44, 51, 5),
    'F51': (f44, 51, 6),
    'G51': (f44, 51, 7),
    'H51': (f44, 51, 8),
    'I51': (f44, 51, 9),
    'J51': (f44, 51, 10),
    'K51': (f44, 51, 11),
    'L51': (f44, 51, 12),
    'C52': (f45, 52, 3),
    'D52': (f45, 52, 4),
    'E52': (f45, 52, 5),
    'F52': (f45, 52, 6),
    'G52': (f45, 52, 7),
    'H52': (f45, 52, 8),
    'I52': (f45, 52, 9),
    'J52': (f45, 52, 10),
    'K52': (f45, 52, 11),
    'L52': (f45, 52, 12),
    'C53': (f46, 53, 3),
    'D53': (f46, 53, 4),
    'E53': (f46, 53, 5),
    'F53': (f46, 53, 6),
    'G53': (f46, 53, 7),
    'H53': (f46, 53, 8),
    'I53': (f46, 53, 9),
    'J53': (f46, 53, 10),
    'K53': (f46, 53, 11),
    'L53': (f46, 53, 12),
    'C54': (f47, 54, 3),
    'D54': (f47, 54, 4),
    'E54': (f47, 54, 5),
    'F54': (f47, 54, 6),
    'G54': (f47, 54, 7),
    'H54': (f47, 54, 8),
    'I54': (f47, 54, 9),
    'J54': (f47, 54, 10),
    'K54': (f47, 54, 11),
    'L54': (f47, 54, 12),
    'C55': (f48, 55, 3),
    'D55': (f48, 55, 4),
    'E55': (f48, 55, 5),
    'F55': (f48, 55, 6),
    'G55': (f48, 55, 7),
    'H55': (f48, 55, 8),
    'I55': (f48, 55, 9),
    'J55': (f48, 55, 10),
    'K55': (f48, 55, 11),
    'L55': (f48, 55, 12),
    'C56': (f49, 56, 3),
    'D56': (f49, 56, 4),
    'E56': (f49, 56, 5),
    'F56': (f49, 56, 6),
    'G56': (f49, 56, 7),
    'H56': (f49, 56, 8),
    'I56': (f49, 56, 9),
    'J56': (f49, 56, 10),
    'K56': (f49, 56, 11),
    'L56': (f49, 56, 12),
    'C57': (f50, 57, 3),
    'D57': (f50, 57, 4),
    'E57': (f50, 57, 5),
    'F57': (f50, 57, 6),
    'G57': (f50, 57, 7),
    'H57': (f50, 57, 8),
    'I57': (f50, 57, 9),
    'J57': (f50, 57, 10),
    'K57': (f50, 57, 11),
    'L57': (f50, 57, 12),
    'C58': (f51, 58, 3),
    'D58': (f51, 58, 4),
    'E58': (f51, 58, 5),
    'F58': (f51, 58, 6),
    'G58': (f51, 58, 7),
    'H58': (f51, 58, 8),
    'I58': (f51, 58, 9),
    'J58': (f51, 58, 10),
    'K58': (f51, 58, 11),
    'L58': (f51, 58, 12),
    'C62': (f52, 62, 3),
    'D62': (f52, 62, 4),
    'E62': (f52, 62, 5),
    'F62': (f52, 62, 6),
    'G62': (f52, 62, 7),
    'H62': (f52, 62, 8),
    'I62': (f52, 62, 9),
    'J62': (f52, 62, 10),
    'K62': (f52, 62, 11),
    'L62': (f52, 62, 12),
    'C63': (f41, 63, 3),
    'D63': (f41, 63, 4),
    'E63': (f41, 63, 5),
    'F63': (f41, 63, 6),
    'G63': (f41, 63, 7),
    'H63': (f41, 63, 8),
    'I63': (f41, 63, 9),
    'J63': (f41, 63, 10),
    'K63': (f41, 63, 11),
    'L63': (f41, 63, 12),
    'C64': (f53, 64, 3),
    'D64': (f53, 64, 4),
    'E64': (f53, 64, 5),
    'F64': (f53, 64, 6),
    'G64': (f53, 64, 7),
    'H64': (f53, 64, 8),
    'I64': (f53, 64, 9),
    'J64': (f53, 64, 10),
    'K64': (f53, 64, 11),
    'L64': (f53, 64, 12),
    'C65': (f54, 65, 3),
    'D65': (f54, 65, 4),
    'E65': (f54, 65, 5),
    'F65': (f54, 65, 6),
    'G65': (f54, 65, 7),
    'H65': (f54, 65, 8),
    'I65': (f54, 65, 9),
    'J65': (f54, 65, 10),
    'K65': (f54, 65, 11),
    'L65': (f54, 65, 12),
    'C66': (f55, 66, 3),
    'D66': (f55, 66, 4),
    'E66': (f55, 66, 5),
    'F66': (f55, 66, 6),
    'G66': (f55, 66, 7),
    'H66': (f55, 66, 8),
    'I66': (f55, 66, 9),
    'J66': (f55, 66, 10),
    'K66': (f55, 66, 11),
    'L66': (f55, 66, 12),
    'C67': (f56, 67, 3),
    'D67': (f56, 67, 4),
    'E67': (f56, 67, 5),
    'F67': (f56, 67, 6),
    'G67': (f56, 67, 7),
    'H67': (f56, 67, 8),
    'I67': (f56, 67, 9),
    'J67': (f56, 67, 10),
    'K67': (f56, 67, 11),
    'L67': (f56, 67, 12),
    'C68': (f57, 68, 3),
    'D68': (f57, 68, 4),
    'E68': (f57, 68, 5),
    'F68': (f57, 68, 6),
    'G68': (f57, 68, 7),
    'H68': (f57, 68, 8),
    'I68': (f57, 68, 9),
    'J68': (f57, 68, 10),
    'K68': (f57, 68, 11),
    'L68': (f57, 68, 12),
    'C69': (f58, 69, 3),
    'D69': (f58, 69, 4),
    'E69': (f58, 69, 5),
    'F69': (f58, 69, 6),
    'G69': (f58, 69, 7),
    'H69': (f58, 69, 8),
    'I69': (f58, 69, 9),
    'J69': (f58, 69, 10),
    'K69': (f58, 69, 11),
    'L69': (f58, 69, 12),
    'C70': (f48, 70, 3),
    'D70': (f48, 70, 4),
    'E70': (f48, 70, 5),
    'F70': (f48, 70, 6),
    'G70': (f48, 70, 7),
    'H70': (f48, 70, 8),
    'I70': (f48, 70, 9),
    'J70': (f48, 70, 10),
    'K70': (f48, 70, 11),
    'L70': (f48, 70, 12),
    'C71': (f59, 71, 3),
    'D71': (f59, 71, 4),
    'E71': (f59, 71, 5),
    'F71': (f59, 71, 6),
    'G71': (f59, 71, 7),
    'H71': (f59, 71, 8),
    'I71': (f59, 71, 9),
    'J71': (f59, 71, 10),
    'K71': (f59, 71, 11),
    'L71': (f59, 71, 12),
    'C72': (f60, 72, 3),
    'D72': (f60, 72, 4),
    'E72': (f60, 72, 5),
    'F72': (f60, 72, 6),
    'G72': (f60, 72, 7),
    'H72': (f60, 72, 8),
    'I72': (f60, 72, 9),
    'J72': (f60, 72, 10),
    'K72': (f60, 72, 11),
    'L72': (f60, 72, 12),
    'C73': (f61, 73, 3),
    'D73': (f61, 73, 4),
    'E73': (f61, 73, 5),
    'F73': (f61, 73, 6),
    'G73': (f61, 73, 7),
    'H73': (f61, 73, 8),
    'I73': (f61, 73, 9),
    'J73': (f61, 73, 10),
    'K73': (f61, 73, 11),
    'L73': (f61, 73, 12),
    'C77': (f62, 77, 3),
    'D77': (f62, 77, 4),
    'E77': (f62, 77, 5),
    'F77': (f62, 77, 6),
    'G77': (f62, 77, 7),
    'H77': (f62, 77, 8),
    'I77': (f62, 77, 9),
    'J77': (f62, 77, 10),
    'K77': (f62, 77, 11),
    'L77': (f62, 77, 12),
    'C78': (f63, 78, 3),
    'D78': (f63, 78, 4),
    'E78': (f63, 78, 5),
    'F78': (f63, 78, 6),
    'G78': (f63, 78, 7),
    'H78': (f63, 78, 8),
    'I78': (f63, 78, 9),
    'J78': (f63, 78, 10),
    'K78': (f63, 78, 11),
    'L78': (f63, 78, 12),
    'C79': (f64, 79, 3),
    'D79': (f64, 79, 4),
    'E79': (f64, 79, 5),
    'F79': (f64, 79, 6),
    'G79': (f64, 79, 7),
    'H79': (f64, 79, 8),
    'I79': (f64, 79, 9),
    'J79': (f64, 79, 10),
    'K79': (f64, 79, 11),
    'L79': (f64, 79, 12),
    'C80': (f65, 80, 3),
    'D80': (f65, 80, 4),
    'E80': (f65, 80, 5),
    'F80': (f65, 80, 6),
    'G80': (f65, 80, 7),
    'H80': (f65, 80, 8),
    'I80': (f65, 80, 9),
    'J80': (f65, 80, 10),
    'K80': (f65, 80, 11),
    'L80': (f65, 80, 12),
    'C81': (f66, 81, 3),
    'D81': (f66, 81, 4),
    'E81': (f66, 81, 5),
    'F81': (f66, 81, 6),
    'G81': (f66, 81, 7),
    'H81': (f66, 81, 8),
    'I81': (f66, 81, 9),
    'J81': (f66, 81, 10),
    'K81': (f66, 81, 11),
    'L81': (f66, 81, 12),
    'C82': (f67, 82, 3),
    'D82': (f67, 82, 4),
    'E82': (f67, 82, 5),
    'F82': (f67, 82, 6),
    'G82': (f67, 82, 7),
    'H82': (f67, 82, 8),
    'I82': (f67, 82, 9),
    'J82': (f67, 82, 10),
    'K82': (f67, 82, 11),
    'L82': (f67, 82, 12),
    'C83': (f68, 83, 3),
    'D83': (f68, 83, 4),
    'E83': (f68, 83, 5),
    'F83': (f68, 83, 6),
    'G83': (f68, 83, 7),
    'H83': (f68, 83, 8),
    'I83': (f68, 83, 9),
    'J83': (f68, 83, 10),
    'K83': (f68, 83, 11),
    'L83': (f68, 83, 12),
    'C84': (f69, 84, 3),
    'D84': (f69, 84, 4),
    'E84': (f69, 84, 5),
    'F84': (f69, 84, 6),
    'G84': (f69, 84, 7),
    'H84': (f69, 84, 8),
    'I84': (f69, 84, 9),
    'J84': (f69, 84, 10),
    'K84': (f69, 84, 11),
    'L84': (f69, 84, 12),
    'C85': (f70, 85, 3),
    'D85': (f70, 85, 4),
    'E85': (f70, 85, 5),
    'F85': (f70, 85, 6),
    'G85': (f70, 85, 7),
    'H85': (f70, 85, 8),
    'I85': (f70, 85, 9),
    'J85': (f70, 85, 10),
    'K85': (f70, 85, 11),
    'L85': (f70, 85, 12),
    'C86': (f71, 86, 3),
    'D86': (f71, 86, 4),
    'E86': (f71, 86, 5),
    'F86': (f71, 86, 6),
    'G86': (f71, 86, 7),
    'H86': (f71, 86, 8),
    'I86': (f71, 86, 9),
    'J86': (f71, 86, 10),
    'K86': (f71, 86, 11),
    'L86': (f71, 86, 12),
    'C87': (f72, 87, 3),
    'D87': (f72, 87, 4),
    'E87': (f72, 87, 5),
    'F87': (f72, 87, 6),
    'G87': (f72, 87, 7),
    'H87': (f72, 87, 8),
    'I87': (f72, 87, 9),
    'J87': (f72, 87, 10),
    'K87': (f72, 87, 11),
    'L87': (f72, 87, 12),
    'C88': (f73, 88, 3),
    'D88': (f73, 88, 4),
    'E88': (f73, 88, 5),
    'F88': (f73, 88, 6),
    'G88': (f73, 88, 7),
    'H88': (f73, 88, 8),
    'I88': (f73, 88, 9),
    'J88': (f73, 88, 10),
    'K88': (f73, 88, 11),
    'L88': (f73, 88, 12),
    'C89': (f74, 89, 3),
    'D89': (f74, 89, 4),
    'E89': (f74, 89, 5),
    'F89': (f74, 89, 6),
    'G89': (f74, 89, 7),
    'H89': (f74, 89, 8),
    'I89': (f74, 89, 9),
    'J89': (f74, 89, 10),
    'K89': (f74, 89, 11),
    'L89': (f74, 89, 12),
    'C90': (f75, 90, 3),
    'D90': (f75, 90, 4),
    'E90': (f75, 90, 5),
    'F90': (f75, 90, 6),
    'G90': (f75, 90, 7),
    'H90': (f75, 90, 8),
    'I90': (f75, 90, 9),
    'J90': (f75, 90, 10),
    'K90': (f75, 90, 11),
    'L90': (f75, 90, 12),
    'C92': (f76, 92, 3),
    'D92': (f76, 92, 4),
    'E92': (f76, 92, 5),
    'F92': (f76, 92, 6),
    'G92': (f76, 92, 7),
    'H92': (f76, 92, 8),
    'I92': (f76, 92, 9),
    'J92': (f76, 92, 10),
    'K92': (f76, 92, 11),
    'L92': (f76, 92, 12),
    'C93': (f77, 93, 3),
    'D93': (f78, 93, 4),
    'E93': (f78, 93, 5),
    'F93': (f78, 93, 6),
    'G93': (f78, 93, 7),
    'H93': (f78, 93, 8),
    'I93': (f78, 93, 9),
    'J93': (f78, 93, 10),
    'K93': (f78, 93, 11),
    'L93': (f78, 93, 12),
    'C94': (f79, 94, 3),
    'C97': (f80, 97, 3),
    'D97': (f80, 97, 4),
    'E97': (f80, 97, 5),
    'F97': (f80, 97, 6),
    'G97': (f80, 97, 7),
    'H97': (f80, 97, 8),
    'I97': (f80, 97, 9),
    'J97': (f80, 97, 10),
    'K97': (f80, 97, 11),
    'L97': (f80, 97, 12),
    'C98': (f81, 98, 3),
    'D98': (f81, 98, 4),
    'E98': (f81, 98, 5),
    'F98': (f81, 98, 6),
    'G98': (f81, 98, 7),
    'H98': (f81, 98, 8),
    'I98': (f81, 98, 9),
    'J98': (f81, 98, 10),
    'K98': (f81, 98, 11),
    'L98': (f81, 98, 12),
    'C99': (f82, 99, 3),
    'D99': (f82, 99, 4),
    'E99': (f82, 99, 5),
    'F99': (f82, 99, 6),
    'G99': (f82, 99, 7),
    'H99': (f82, 99, 8),
    'I99': (f82, 99, 9),
    'J99': (f82, 99, 10),
    'K99': (f82, 99, 11),
    'L99': (f82, 99, 12),
    'C100': (f83, 100, 3),
    'D100': (f83, 100, 4),
    'E100': (f83, 100, 5),
    'F100': (f83, 100, 6),
    'G100': (f83, 100, 7),
    'H100': (f83, 100, 8),
    'I100': (f83, 100, 9),
    'J100': (f83, 100, 10),
    'K100': (f83, 100, 11),
    'L100': (f83, 100, 12),
    'C101': (f84, 101, 3),
    'D101': (f84, 101, 4),
    'E101': (f84, 101, 5),
    'F101': (f84, 101, 6),
    'G101': (f84, 101, 7),
    'H101': (f84, 101, 8),
    'I101': (f84, 101, 9),
    'J101': (f84, 101, 10),
    'K101': (f84, 101, 11),
    'L101': (f84, 101, 12),
    'C102': (f85, 102, 3),
    'D102': (f85, 102, 4),
    'E102': (f85, 102, 5),
    'F102': (f85, 102, 6),
    'G102': (f85, 102, 7),
    'H102': (f85, 102, 8),
    'I102': (f85, 102, 9),
    'J102': (f85, 102, 10),
    'K102': (f85, 102, 11),
    'L102': (f85, 102, 12),
    'C103': (f86, 103, 3),
    'D103': (f86, 103, 4),
    'E103': (f86, 103, 5),
    'F103': (f86, 103, 6),
    'G103': (f86, 103, 7),
    'H103': (f86, 103, 8),
    'I103': (f86, 103, 9),
    'J103': (f86, 103, 10),
    'K103': (f86, 103, 11),
    'L103': (f86, 103, 12),
    'C104': (f87, 104, 3),
    'D104': (f87, 104, 4),
    'E104': (f87, 104, 5),
    'F104': (f87, 104, 6),
    'G104': (f87, 104, 7),
    'H104': (f87, 104, 8),
    'I104': (f87, 104, 9),
    'J104': (f87, 104, 10),
    'K104': (f87, 104, 11),
    'L104': (f87, 104, 12),
    'C105': (f88, 105, 3),
    'D105': (f88, 105, 4),
    'E105': (f88, 105, 5),
    'F105': (f88, 105, 6),
    'G105': (f88, 105, 7),
    'H105': (f88, 105, 8),
    'I105': (f88, 105, 9),
    'J105': (f88, 105, 10),
    'K105': (f88, 105, 11),
    'L105': (f88, 105, 12),
    'C106': (f89, 106, 3),
    'D106': (f89, 106, 4),
    'E106': (f89, 106, 5),
    'F106': (f89, 106, 6),
    'G106': (f89, 106, 7),
    'H106': (f89, 106, 8),
    'I106': (f89, 106, 9),
    'J106': (f89, 106, 10),
    'K106': (f89, 106, 11),
    'L106': (f89, 106, 12),
    'C109': (f90, 109, 3),
    'C110': (f91, 110, 3),
    'C111': (f92, 111, 3),
    'C112': (f93, 112, 3),
    'C113': (f94, 113, 3),
    'C115': (f95, 115, 3),
    'D115': (f95, 115, 4),
    'E115': (f95, 115, 5),
    'F115': (f95, 115, 6),
    'G115': (f95, 115, 7),
    'H115': (f95, 115, 8),
    'I115': (f95, 115, 9),
    'J115': (f95, 115, 10),
    'K115': (f95, 115, 11),
    'L115': (f95, 115, 12),
    'C116': (f96, 116, 3),
    'D116': (f96, 116, 4),
    'E116': (f96, 116, 5),
    'F116': (f96, 116, 6),
    'G116': (f96, 116, 7),
    'H116': (f96, 116, 8),
    'I116': (f96, 116, 9),
    'J116': (f96, 116, 10),
    'K116': (f96, 116, 11),
    'L116': (f96, 116, 12),
    'C117': (f97, 117, 3),
    'D117': (f97, 117, 4),
    'E117': (f97, 117, 5),
    'F117': (f97, 117, 6),
    'G117': (f97, 117, 7),
    'H117': (f97, 117, 8),
    'I117': (f97, 117, 9),
    'J117': (f97, 117, 10),
    'K117': (f97, 117, 11),
    'L117': (f97, 117, 12),
    'C118': (f98, 118, 3),
    'D118': (f98, 118, 4),
    'E118': (f98, 118, 5),
    'F118': (f98, 118, 6),
    'G118': (f98, 118, 7),
    'H118': (f98, 118, 8),
    'I118': (f98, 118, 9),
    'J118': (f98, 118, 10),
    'K118': (f98, 118, 11),
    'L118': (f98, 118, 12),
    'C119': (f99, 119, 3),
    'D119': (f99, 119, 4),
    'E119': (f99, 119, 5),
    'F119': (f99, 119, 6),
    'G119': (f99, 119, 7),
    'H119': (f99, 119, 8),
    'I119': (f99, 119, 9),
    'J119': (f99, 119, 10),
    'K119': (f99, 119, 11),
    'L119': (f99, 119, 12),
    'C120': (f99, 120, 3),
    'D120': (f99, 120, 4),
    'E120': (f99, 120, 5),
    'F120': (f99, 120, 6),
    'G120': (f99, 120, 7),
    'H120': (f99, 120, 8),
    'I120': (f99, 120, 9),
    'J120': (f99, 120, 10),
    'K120': (f99, 120, 11),
    'L120': (f99, 120, 12),
    'C121': (f100, 121, 3),
    'D121': (f100, 121, 4),
    'E121': (f100, 121, 5),
    'F121': (f100, 121, 6),
    'G121': (f100, 121, 7),
    'H121': (f100, 121, 8),
    'I121': (f100, 121, 9),
    'J121': (f100, 121, 10),
    'K121': (f100, 121, 11),
    'L121': (f100, 121, 12),
    'C122': (f101, 122, 3),
    'D122': (f101, 122, 4),
    'E122': (f101, 122, 5),
    'F122': (f101, 122, 6),
    'G122': (f101, 122, 7),
    'H122': (f101, 122, 8),
    'I122': (f101, 122, 9),
    'J122': (f101, 122, 10),
    'K122': (f101, 122, 11),
    'L122': (f101, 122, 12),
    'C123': (f102, 123, 3),
    'D123': (f102, 123, 4),
    'E123': (f102, 123, 5),
    'F123': (f102, 123, 6),
    'G123': (f102, 123, 7),
    'H123': (f102, 123, 8),
    'I123': (f102, 123, 9),
    'J123': (f102, 123, 10),
    'K123': (f102, 123, 11),
    'L123': (f102, 123, 12),
    'C124': (f103, 124, 3),
    'D124': (f103, 124, 4),
    'E124': (f103, 124, 5),
    'F124': (f103, 124, 6),
    'G124': (f103, 124, 7),
    'H124': (f103, 124, 8),
    'I124': (f103, 124, 9),
    'J124': (f103, 124, 10),
    'K124': (f103, 124, 11),
    'L124': (f103, 124, 12),
    'C125': (f104, 125, 3),
    'D125': (f104, 125, 4),
    'E125': (f104, 125, 5),
    'F125': (f104, 125, 6),
    'G125': (f104, 125, 7),
    'H125': (f104, 125, 8),
    'I125': (f104, 125, 9),
    'J125': (f104, 125, 10),
    'K125': (f104, 125, 11),
    'L125': (f104, 125, 12),
    'C126': (f105, 126, 3),
    'D126': (f105, 126, 4),
    'E126': (f105, 126, 5),
    'F126': (f105, 126, 6),
    'G126': (f105, 126, 7),
    'H126': (f105, 126, 8),
    'I126': (f105, 126, 9),
    'J126': (f105, 126, 10),
    'C127': (f106, 127, 3),
    'D127': (f106, 127, 4),
    'E127': (f106, 127, 5),
    'F127': (f106, 127, 6),
    'G127': (f106, 127, 7),
    'H127': (f106, 127, 8),
    'I127': (f106, 127, 9),
    'J127': (f106, 127, 10),
    'K127': (f106, 127, 11),
    'L127': (f106, 127, 12),
    'C130': (f107, 130, 3),
    'D130': (f107, 130, 4),
    'E130': (f107, 130, 5),
    'F130': (f107, 130, 6),
    'G130': (f107, 130, 7),
    'H130': (f107, 130, 8),
    'I130': (f107, 130, 9),
    'J130': (f107, 130, 10),
    'K130': (f107, 130, 11),
    'L130': (f107, 130, 12),
    'C131': (f108, 131, 3),
    'D131': (f108, 131, 4),
    'E131': (f108, 131, 5),
    'F131': (f108, 131, 6),
    'G131': (f108, 131, 7),
    'H131': (f108, 131, 8),
    'I131': (f108, 131, 9),
    'J131': (f108, 131, 10),
    'K131': (f108, 131, 11),
    'L131': (f108, 131, 12),
    'C132': (f109, 132, 3),
    'D132': (f109, 132, 4),
    'E132': (f109, 132, 5),
    'F132': (f109, 132, 6),
    'G132': (f109, 132, 7),
    'H132': (f109, 132, 8),
    'I132': (f109, 132, 9),
    'J132': (f109, 132, 10),
    'K132': (f109, 132, 11),
    'L132': (f109, 132, 12),
    'C133': (f110, 133, 3),
    'D133': (f110, 133, 4),
    'E133': (f110, 133, 5),
    'F133': (f110, 133, 6),
    'G133': (f110, 133, 7),
    'H133': (f110, 133, 8),
    'I133': (f110, 133, 9),
    'J133': (f110, 133, 10),
    'K133': (f110, 133, 11),
    'L133': (f110, 133, 12),
    'C134': (f111, 134, 3),
    'D134': (f111, 134, 4),
    'E134': (f111, 134, 5),
    'F134': (f111, 134, 6),
    'G134': (f111, 134, 7),
    'H134': (f111, 134, 8),
    'I134': (f111, 134, 9),
    'J134': (f111, 134, 10),
    'K134': (f111, 134, 11),
    'L134': (f111, 134, 12),
    'C135': (f112, 135, 3),
    'D135': (f112, 135, 4),
    'E135': (f112, 135, 5),
    'F135': (f112, 135, 6),
    'G135': (f112, 135, 7),
    'H135': (f112, 135, 8),
    'I135': (f112, 135, 9),
    'J135': (f112, 135, 10),
    'K135': (f112, 135, 11),
    'L135': (f112, 135, 12),
    'C136': (f113, 136, 3),
    'D136': (f113, 136, 4),
    'E136': (f113, 136, 5),
    'F136': (f113, 136, 6),
    'G136': (f113, 136, 7),
    'H136': (f113, 136, 8),
    'I136': (f113, 136, 9),
    'J136': (f113, 136, 10),
    'K136': (f113, 136, 11),
    'L136': (f113, 136, 12),
    'C137': (f114, 137, 3),
    'D137': (f114, 137, 4),
    'E137': (f114, 137, 5),
    'F137': (f114, 137, 6),
    'G137': (f114, 137, 7),
    'H137': (f114, 137, 8),
    'I137': (f114, 137, 9),
    'J137': (f114, 137, 10),
    'K137': (f114, 137, 11),
    'L137': (f114, 137, 12),
    'C138': (f115, 138, 3),
    'D138': (f115, 138, 4),
    'E138': (f115, 138, 5),
    'F138': (f115, 138, 6),
    'G138': (f115, 138, 7),
    'H138': (f115, 138, 8),
    'I138': (f115, 138, 9),
    'J138': (f115, 138, 10),
    'K138': (f115, 138, 11),
    'L138': (f115, 138, 12),
    'C139': (f116, 139, 3),
    'D139': (f116, 139, 4),
    'E139': (f116, 139, 5),
    'F139': (f116, 139, 6),
    'G139': (f116, 139, 7),
    'H139': (f116, 139, 8),
    'I139': (f116, 139, 9),
    'J139': (f116, 139, 10),
    'K139': (f116, 139, 11),
    'L139': (f116, 139, 12),
}
