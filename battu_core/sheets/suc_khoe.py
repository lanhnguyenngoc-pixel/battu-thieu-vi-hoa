# -*- coding: utf-8 -*-
"""Module sheet SUC_KHOE – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: 6.10_G.xlsx (SHA-256 a34b9bd88b9052807174d240d6d3d83136557f1d078a7e3a3ba2d95e1eadc33f).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'SUC_KHOE'
SI = 22

def f0(E, r, c):
    # Ô C6
    # = IF(TRIM(IF(L62="ĐỦ","• "&G62&" ["&B62&", "&H62&"] ","")&IF(L63="ĐỦ","• "&G63&" ["&B63&", "&H63&"] ","")&IF(L64="ĐỦ","• "&G64&" ["&B64&", "&H64&"] ",""
    #   )&IF(L65="ĐỦ","• "&G65&" ["&B65&", "&H65&"] ","")&IF(L66="ĐỦ","• "&G66&" ["&B66&", "&H66&"] ","")&IF(L67="ĐỦ","• "&G67&" ["&B67&", "&H67&"] ",""))="",
    #   "(không có dấu hiệu thuận nào của sách đủ điều kiện)",TRIM(IF(L62="ĐỦ","• "&G62&" ["&B62&", "&H62&"] ","")&IF(L63="ĐỦ","• "&G63&" ["&B63&", "&H63&"] "
    #   ,"")&IF(L64="ĐỦ","• "&G64&" ["&B64&", "&H64&"] ","")&IF(L65="ĐỦ","• "&G65&" ["&B65&", "&H65&"] ","")&IF(L66="ĐỦ","• "&G66&" ["&B66&", "&H66&"] ","")&I
    #   F(L67="ĐỦ","• "&G67&" ["&B67&", "&H67&"] ","")))
    return ('(không có dấu hiệu thuận nào của sách đủ điều kiện)' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+56,c+4)),' ['),E.v(22,r+56,c-1)),', '),E.v(22,r+56,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+56,c+9),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+57,c+4)),' ['),E.v(22,r+57,c-1)),', '),E.v(22,r+57,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+57,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+58,c+4)),' ['),E.v(22,r+58,c-1)),', '),E.v(22,r+58,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+58,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+59,c+4)),' ['),E.v(22,r+59,c-1)),', '),E.v(22,r+59,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+59,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+60,c+4)),' ['),E.v(22,r+60,c-1)),', '),E.v(22,r+60,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+60,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+61,c+4)),' ['),E.v(22,r+61,c-1)),', '),E.v(22,r+61,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+61,c+9),'ĐỦ')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+56,c+4)),' ['),E.v(22,r+56,c-1)),', '),E.v(22,r+56,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+56,c+9),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+57,c+4)),' ['),E.v(22,r+57,c-1)),', '),E.v(22,r+57,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+57,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+58,c+4)),' ['),E.v(22,r+58,c-1)),', '),E.v(22,r+58,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+58,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+59,c+4)),' ['),E.v(22,r+59,c-1)),', '),E.v(22,r+59,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+59,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+60,c+4)),' ['),E.v(22,r+60,c-1)),', '),E.v(22,r+60,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+60,c+9),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+61,c+4)),' ['),E.v(22,r+61,c-1)),', '),E.v(22,r+61,c+5)),'] ') if X.truthy(X.eq(E.v(22,r+61,c+9),'ĐỦ')) else ''))))

def f1(E, r, c):
    # Ô F6
    # = IF(""="","(không có dấu hiệu bất lợi nào của sách đủ điều kiện)","")
    return ('(không có dấu hiệu bất lợi nào của sách đủ điều kiện)' if X.truthy(X.eq('','')) else '')

def f2(E, r, c):
    # Ô J6
    # = SUBSTITUTE(IF(IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),1,IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),-1,0))=1,IF(
    #   M6>=N6,"THUẬN","THUẬN NHƯNG CÓ TRỞ NGẠI"),IF(IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),1,IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,D
    #   UNG_HY_KY!$C$7)=0),-1,0))=-1,IF(N6>=M6,"BẤT LỢI","BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M6>N6,"NGHIÊNG THUẬN",IF(N6>M6,"NGHIÊNG BẤT LỢI","TRUNG BÌNH"))))&CH
    #   AR(10)&"("&M6&" dấu hiệu thuận, "&N6&" bất lợi; yếu tố then chốt: "&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),"thân vượng",IF((IF(D
    #   OI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"thân nhược","thân trung hòa"))&" – sách: ""trụ ngày cao cường thì cuộc đời ít bệnh"""&")","(0 dấ
    #   u hiệu thuận, 0 bất lợi; yếu tố then chốt: ","(KHÔNG có dấu hiệu cụ thể nào của sách – mức chỉ dựa vào yếu tố then chốt: ")
    return X.SUBSTITUTE(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(22,r,c+3),E.v(22,r,c+4))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq((1.0 if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0)) else (X.neg(1.0) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else 0.0)),1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(22,r,c+4),E.v(22,r,c+3))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq((1.0 if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0)) else (X.neg(1.0) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else 0.0)),X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(22,r,c+3),E.v(22,r,c+4))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(22,r,c+4),E.v(22,r,c+3))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(22,r,c+3)),' dấu hiệu thuận, '),E.v(22,r,c+4)),' bất lợi; yếu tố then chốt: '),('thân vượng' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0)) else ('thân nhược' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else 'thân trung hòa'))),' – sách: "trụ ngày cao cường thì cuộc đời ít bệnh"'),')'), '(0 dấu hiệu thuận, 0 bất lợi; yếu tố then chốt: ', '(KHÔNG có dấu hiệu cụ thể nào của sách – mức chỉ dựa vào yếu tố then chốt: ')

def f3(E, r, c):
    # Ô M6
    # = ((L62="ĐỦ")+(L63="ĐỦ")+(L64="ĐỦ")+(L65="ĐỦ")+(L66="ĐỦ")+(L67="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.eq(E.v(22,r+56,c-1),'ĐỦ'),X.eq(E.v(22,r+57,c-1),'ĐỦ')),X.eq(E.v(22,r+58,c-1),'ĐỦ')),X.eq(E.v(22,r+59,c-1),'ĐỦ')),X.eq(E.v(22,r+60,c-1),'ĐỦ')),X.eq(E.v(22,r+61,c-1),'ĐỦ'))

def f4(E, r, c):
    # Ô N6
    # = 0
    return 0.0

def f5(E, r, c):
    # Ô C7, C8
    # = IF(TRIM(AB7&" "&AC7)="","(không có dấu hiệu thuận nào của sách đủ điều kiện)",TRIM(AB7&" "&AC7))
    return ('(không có dấu hiệu thuận nào của sách đủ điều kiện)' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(22,r,c+25),' '),E.v(22,r,c+26))),'')) else X.TRIM(X.cat(X.cat(E.v(22,r,c+25),' '),E.v(22,r,c+26))))

def f6(E, r, c):
    # Ô F7, F8
    # = IF(TRIM(AD7&" "&AE7)="","(không có dấu hiệu bất lợi nào của sách đủ điều kiện)",TRIM(AD7&" "&AE7))
    return ('(không có dấu hiệu bất lợi nào của sách đủ điều kiện)' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(22,r,c+24),' '),E.v(22,r,c+25))),'')) else X.TRIM(X.cat(X.cat(E.v(22,r,c+24),' '),E.v(22,r,c+25))))

def f7(E, r, c):
    # Ô J7
    # = SUBSTITUTE(IF(0=1,IF(M7>=N7,"THUẬN","THUẬN NHƯNG CÓ TRỞ NGẠI"),IF(0=-1,IF(N7>=M7,"BẤT LỢI","BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M7>N7,"NGHIÊNG THUẬN",IF(N
    #   7>M7,"NGHIÊNG BẤT LỢI","TRUNG BÌNH"))))&CHAR(10)&"("&M7&" dấu hiệu thuận, "&N7&" bất lợi; yếu tố then chốt: "&"so số dấu hiệu bệnh tật của sách"&")","
    #   (0 dấu hiệu thuận, 0 bất lợi; yếu tố then chốt: so số dấu hiệu bệnh tật của sách)","(KHÔNG có dấu hiệu bệnh tật cụ thể nào của sách → để mức TRUNG BÌN
    #   H, không suy thêm)")
    return X.SUBSTITUTE(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(22,r,c+3),E.v(22,r,c+4))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq(0.0,1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(22,r,c+4),E.v(22,r,c+3))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq(0.0,X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(22,r,c+3),E.v(22,r,c+4))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(22,r,c+4),E.v(22,r,c+3))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(22,r,c+3)),' dấu hiệu thuận, '),E.v(22,r,c+4)),' bất lợi; yếu tố then chốt: '),'so số dấu hiệu bệnh tật của sách'),')'), '(0 dấu hiệu thuận, 0 bất lợi; yếu tố then chốt: so số dấu hiệu bệnh tật của sách)', '(KHÔNG có dấu hiệu bệnh tật cụ thể nào của sách → để mức TRUNG BÌNH, không suy thêm)')

def f8(E, r, c):
    # Ô M7
    # = ((L83="ĐỦ"))+0
    return X.add(X.eq(E.v(22,r+76,c-1),'ĐỦ'),0.0)

def f9(E, r, c):
    # Ô N7
    # = ((L68="ĐỦ")+(L69="ĐỦ")+(L70="ĐỦ")+(L71="ĐỦ")+(L72="ĐỦ")+(L73="ĐỦ")+(L74="ĐỦ")+(L75="ĐỦ")+(L76="ĐỦ")+(L77="ĐỦ")+(L78="ĐỦ")+(L79="ĐỦ")+(L80="ĐỦ")+(L81="
    #   ĐỦ")+(L82="ĐỦ")+(L90="ĐỦ")+(L91="ĐỦ"))+((L96="ĐỦ")+(L97="ĐỦ")+(L98="ĐỦ")+(L99="ĐỦ")+(L100="ĐỦ")+(L101="ĐỦ")+(L102="ĐỦ")+(L103="ĐỦ")+(L104="ĐỦ")+(L105=
    #   "ĐỦ")+(L106="ĐỦ")+(L107="ĐỦ")+(L108="ĐỦ")+(L109="ĐỦ")+(L110="ĐỦ")+(L111="ĐỦ")+(L112="ĐỦ")+(L113="ĐỦ")+(L114="ĐỦ")+(L115="ĐỦ")+(L116="ĐỦ")+(L117="ĐỦ")+
    #   (L118="ĐỦ")+(L119="ĐỦ")+(L120="ĐỦ")+(L121="ĐỦ")+(L122="ĐỦ")+(L123="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(22,r+61,c-2),'ĐỦ'),X.eq(E.v(22,r+62,c-2),'ĐỦ')),X.eq(E.v(22,r+63,c-2),'ĐỦ')),X.eq(E.v(22,r+64,c-2),'ĐỦ')),X.eq(E.v(22,r+65,c-2),'ĐỦ')),X.eq(E.v(22,r+66,c-2),'ĐỦ')),X.eq(E.v(22,r+67,c-2),'ĐỦ')),X.eq(E.v(22,r+68,c-2),'ĐỦ')),X.eq(E.v(22,r+69,c-2),'ĐỦ')),X.eq(E.v(22,r+70,c-2),'ĐỦ')),X.eq(E.v(22,r+71,c-2),'ĐỦ')),X.eq(E.v(22,r+72,c-2),'ĐỦ')),X.eq(E.v(22,r+73,c-2),'ĐỦ')),X.eq(E.v(22,r+74,c-2),'ĐỦ')),X.eq(E.v(22,r+75,c-2),'ĐỦ')),X.eq(E.v(22,r+83,c-2),'ĐỦ')),X.eq(E.v(22,r+84,c-2),'ĐỦ')),X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(22,r+89,c-2),'ĐỦ'),X.eq(E.v(22,r+90,c-2),'ĐỦ')),X.eq(E.v(22,r+91,c-2),'ĐỦ')),X.eq(E.v(22,r+92,c-2),'ĐỦ')),X.eq(E.v(22,r+93,c-2),'ĐỦ')),X.eq(E.v(22,r+94,c-2),'ĐỦ')),X.eq(E.v(22,r+95,c-2),'ĐỦ')),X.eq(E.v(22,r+96,c-2),'ĐỦ')),X.eq(E.v(22,r+97,c-2),'ĐỦ')),X.eq(E.v(22,r+98,c-2),'ĐỦ')),X.eq(E.v(22,r+99,c-2),'ĐỦ')),X.eq(E.v(22,r+100,c-2),'ĐỦ')),X.eq(E.v(22,r+101,c-2),'ĐỦ')),X.eq(E.v(22,r+102,c-2),'ĐỦ')),X.eq(E.v(22,r+103,c-2),'ĐỦ')),X.eq(E.v(22,r+104,c-2),'ĐỦ')),X.eq(E.v(22,r+105,c-2),'ĐỦ')),X.eq(E.v(22,r+106,c-2),'ĐỦ')),X.eq(E.v(22,r+107,c-2),'ĐỦ')),X.eq(E.v(22,r+108,c-2),'ĐỦ')),X.eq(E.v(22,r+109,c-2),'ĐỦ')),X.eq(E.v(22,r+110,c-2),'ĐỦ')),X.eq(E.v(22,r+111,c-2),'ĐỦ')),X.eq(E.v(22,r+112,c-2),'ĐỦ')),X.eq(E.v(22,r+113,c-2),'ĐỦ')),X.eq(E.v(22,r+114,c-2),'ĐỦ')),X.eq(E.v(22,r+115,c-2),'ĐỦ')),X.eq(E.v(22,r+116,c-2),'ĐỦ')))

def f10(E, r, c):
    # Ô AB7
    # = TRIM(IF(L83="ĐỦ","• "&G83&" ["&B83&", "&H83&"] ",""))
    return X.TRIM((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+76,c-21)),' ['),E.v(22,r+76,c-26)),', '),E.v(22,r+76,c-20)),'] ') if X.truthy(X.eq(E.v(22,r+76,c-16),'ĐỦ')) else ''))

def f11(E, r, c):
    # Ô AC7, AB8, AC8
    # = ""
    return ''

def f12(E, r, c):
    # Ô AD7
    # = TRIM(IF(L68="ĐỦ","• "&G68&" ["&B68&", "&H68&"] ","")&IF(L69="ĐỦ","• "&G69&" ["&B69&", "&H69&"] ","")&IF(L70="ĐỦ","• "&G70&" ["&B70&", "&H70&"] ","")&I
    #   F(L71="ĐỦ","• "&G71&" ["&B71&", "&H71&"] ","")&IF(L72="ĐỦ","• "&G72&" ["&B72&", "&H72&"] ","")&IF(L73="ĐỦ","• "&G73&" ["&B73&", "&H73&"] ","")&IF(L74=
    #   "ĐỦ","• "&G74&" ["&B74&", "&H74&"] ","")&IF(L75="ĐỦ","• "&G75&" ["&B75&", "&H75&"] ","")&IF(L76="ĐỦ","• "&G76&" ["&B76&", "&H76&"] ","")&IF(L77="ĐỦ","
    #   • "&G77&" ["&B77&", "&H77&"] ","")&IF(L78="ĐỦ","• "&G78&" ["&B78&", "&H78&"] ","")&IF(L79="ĐỦ","• "&G79&" ["&B79&", "&H79&"] ","")&IF(L80="ĐỦ","• "&G8
    #   0&" ["&B80&", "&H80&"] ","")&IF(L81="ĐỦ","• "&G81&" ["&B81&", "&H81&"] ","")&IF(L82="ĐỦ","• "&G82&" ["&B82&", "&H82&"] ","")&IF(L90="ĐỦ","• "&G90&" ["
    #   &B90&", "&H90&"] ","")&IF(L91="ĐỦ","• "&G91&" ["&B91&", "&H91&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+61,c-23)),' ['),E.v(22,r+61,c-28)),', '),E.v(22,r+61,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+61,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+62,c-23)),' ['),E.v(22,r+62,c-28)),', '),E.v(22,r+62,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+62,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+63,c-23)),' ['),E.v(22,r+63,c-28)),', '),E.v(22,r+63,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+63,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+64,c-23)),' ['),E.v(22,r+64,c-28)),', '),E.v(22,r+64,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+64,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+65,c-23)),' ['),E.v(22,r+65,c-28)),', '),E.v(22,r+65,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+65,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+66,c-23)),' ['),E.v(22,r+66,c-28)),', '),E.v(22,r+66,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+66,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+67,c-23)),' ['),E.v(22,r+67,c-28)),', '),E.v(22,r+67,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+67,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+68,c-23)),' ['),E.v(22,r+68,c-28)),', '),E.v(22,r+68,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+68,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+69,c-23)),' ['),E.v(22,r+69,c-28)),', '),E.v(22,r+69,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+69,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+70,c-23)),' ['),E.v(22,r+70,c-28)),', '),E.v(22,r+70,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+70,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+71,c-23)),' ['),E.v(22,r+71,c-28)),', '),E.v(22,r+71,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+71,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+72,c-23)),' ['),E.v(22,r+72,c-28)),', '),E.v(22,r+72,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+72,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+73,c-23)),' ['),E.v(22,r+73,c-28)),', '),E.v(22,r+73,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+73,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+74,c-23)),' ['),E.v(22,r+74,c-28)),', '),E.v(22,r+74,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+74,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+75,c-23)),' ['),E.v(22,r+75,c-28)),', '),E.v(22,r+75,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+75,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+83,c-23)),' ['),E.v(22,r+83,c-28)),', '),E.v(22,r+83,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+83,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+84,c-23)),' ['),E.v(22,r+84,c-28)),', '),E.v(22,r+84,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+84,c-18),'ĐỦ')) else '')))

def f13(E, r, c):
    # Ô AE7
    # = TRIM(IF(L96="ĐỦ","• "&G96&" ["&B96&", "&H96&"] ","")&IF(L97="ĐỦ","• "&G97&" ["&B97&", "&H97&"] ","")&IF(L98="ĐỦ","• "&G98&" ["&B98&", "&H98&"] ","")&I
    #   F(L99="ĐỦ","• "&G99&" ["&B99&", "&H99&"] ","")&IF(L100="ĐỦ","• "&G100&" ["&B100&", "&H100&"] ","")&IF(L101="ĐỦ","• "&G101&" ["&B101&", "&H101&"] ","")
    #   &IF(L102="ĐỦ","• "&G102&" ["&B102&", "&H102&"] ","")&IF(L103="ĐỦ","• "&G103&" ["&B103&", "&H103&"] ","")&IF(L104="ĐỦ","• "&G104&" ["&B104&", "&H104&"]
    #    ","")&IF(L105="ĐỦ","• "&G105&" ["&B105&", "&H105&"] ","")&IF(L106="ĐỦ","• "&G106&" ["&B106&", "&H106&"] ","")&IF(L107="ĐỦ","• "&G107&" ["&B107&", "&H
    #   107&"] ","")&IF(L108="ĐỦ","• "&G108&" ["&B108&", "&H108&"] ","")&IF(L109="ĐỦ","• "&G109&" ["&B109&", "&H109&"] ","")&IF(L110="ĐỦ","• "&G110&" ["&B110&
    #   ", "&H110&"] ","")&IF(L111="ĐỦ","• "&G111&" ["&B111&", "&H111&"] ","")&IF(L112="ĐỦ","• "&G112&" ["&B112&", "&H112&"] ","")&IF(L113="ĐỦ","• "&G113&" ["
    #   &B113&", "&H113&"] ","")&IF(L114="ĐỦ","• "&G114&" ["&B114&", "&H114&"] ","")&IF(L115="ĐỦ","• "&G115&" ["&B115&", "&H115&"] ","")&IF(L116="ĐỦ","• "&G11
    #   6&" ["&B116&", "&H116&"] ","")&IF(L117="ĐỦ","• "&G117&" ["&B117&", "&H117&"] ","")&IF(L118="ĐỦ","• "&G118&" ["&B118&", "&H118&"] ","")&IF(L119="ĐỦ","•
    #    "&G119&" ["&B119&", "&H119&"] ","")&IF(L120="ĐỦ","• "&G120&" ["&B120&", "&H120&"] ","")&IF(L121="ĐỦ","• "&G121&" ["&B121&", "&H121&"] ","")&IF(L122="
    #   ĐỦ","• "&G122&" ["&B122&", "&H122&"] ","")&IF(L123="ĐỦ","• "&G123&" ["&B123&", "&H123&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+89,c-24)),' ['),E.v(22,r+89,c-29)),', '),E.v(22,r+89,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+89,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+90,c-24)),' ['),E.v(22,r+90,c-29)),', '),E.v(22,r+90,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+90,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+91,c-24)),' ['),E.v(22,r+91,c-29)),', '),E.v(22,r+91,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+91,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+92,c-24)),' ['),E.v(22,r+92,c-29)),', '),E.v(22,r+92,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+92,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+93,c-24)),' ['),E.v(22,r+93,c-29)),', '),E.v(22,r+93,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+93,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+94,c-24)),' ['),E.v(22,r+94,c-29)),', '),E.v(22,r+94,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+94,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+95,c-24)),' ['),E.v(22,r+95,c-29)),', '),E.v(22,r+95,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+95,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+96,c-24)),' ['),E.v(22,r+96,c-29)),', '),E.v(22,r+96,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+96,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+97,c-24)),' ['),E.v(22,r+97,c-29)),', '),E.v(22,r+97,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+97,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+98,c-24)),' ['),E.v(22,r+98,c-29)),', '),E.v(22,r+98,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+98,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+99,c-24)),' ['),E.v(22,r+99,c-29)),', '),E.v(22,r+99,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+99,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+100,c-24)),' ['),E.v(22,r+100,c-29)),', '),E.v(22,r+100,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+100,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+101,c-24)),' ['),E.v(22,r+101,c-29)),', '),E.v(22,r+101,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+101,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+102,c-24)),' ['),E.v(22,r+102,c-29)),', '),E.v(22,r+102,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+102,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+103,c-24)),' ['),E.v(22,r+103,c-29)),', '),E.v(22,r+103,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+103,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+104,c-24)),' ['),E.v(22,r+104,c-29)),', '),E.v(22,r+104,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+104,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+105,c-24)),' ['),E.v(22,r+105,c-29)),', '),E.v(22,r+105,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+105,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+106,c-24)),' ['),E.v(22,r+106,c-29)),', '),E.v(22,r+106,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+106,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+107,c-24)),' ['),E.v(22,r+107,c-29)),', '),E.v(22,r+107,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+107,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+108,c-24)),' ['),E.v(22,r+108,c-29)),', '),E.v(22,r+108,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+108,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+109,c-24)),' ['),E.v(22,r+109,c-29)),', '),E.v(22,r+109,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+109,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+110,c-24)),' ['),E.v(22,r+110,c-29)),', '),E.v(22,r+110,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+110,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+111,c-24)),' ['),E.v(22,r+111,c-29)),', '),E.v(22,r+111,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+111,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+112,c-24)),' ['),E.v(22,r+112,c-29)),', '),E.v(22,r+112,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+112,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+113,c-24)),' ['),E.v(22,r+113,c-29)),', '),E.v(22,r+113,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+113,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+114,c-24)),' ['),E.v(22,r+114,c-29)),', '),E.v(22,r+114,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+114,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+115,c-24)),' ['),E.v(22,r+115,c-29)),', '),E.v(22,r+115,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+115,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+116,c-24)),' ['),E.v(22,r+116,c-29)),', '),E.v(22,r+116,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+116,c-19),'ĐỦ')) else '')))

def f14(E, r, c):
    # Ô J8
    # = SUBSTITUTE(IF(IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG
    #   _HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRU
    #   NG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),-1,0)=1,IF(M8>=N8,"THUẬN","THUẬN NHƯNG CÓ TRỞ NGẠI"),IF(IF(AND((IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP
    #   !$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),-1,0)=-1,IF(N8>=M8,"BẤT LỢI","BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M8>N8,"NGHIÊNG THUẬN",IF(N8>M8,"NGHI
    #   ÊNG BẤT LỢI","TRUNG BÌNH"))))&CHAR(10)&"("&M8&" dấu hiệu thuận, "&N8&" bất lợi; yếu tố then chốt: "&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_K
    #   Y!$C$7)=1),"thân vượng",IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"thân nhược","thân trung hòa"))&", quan sát "&IF(MOD(TRUNG_GIAN!$
    #   E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,D
    #   UNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$
    #   2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!
    #   $C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))
    #   &")","(0 dấu hiệu thuận, 0 bất lợi; yếu tố then chốt: ","(KHÔNG có dấu hiệu cụ thể nào của sách – mức chỉ dựa vào yếu tố then chốt: ")
    return X.SUBSTITUTE(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(22,r,c+3),E.v(22,r,c+4))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq((X.neg(1.0) if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 0.0),1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(22,r,c+4),E.v(22,r,c+3))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq((X.neg(1.0) if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 0.0),X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(22,r,c+3),E.v(22,r,c+4))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(22,r,c+4),E.v(22,r,c+3))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(22,r,c+3)),' dấu hiệu thuận, '),E.v(22,r,c+4)),' bất lợi; yếu tố then chốt: '),('thân vượng' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0)) else ('thân nhược' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else 'thân trung hòa'))),', quan sát '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),')'), '(0 dấu hiệu thuận, 0 bất lợi; yếu tố then chốt: ', '(KHÔNG có dấu hiệu cụ thể nào của sách – mức chỉ dựa vào yếu tố then chốt: ')

def f15(E, r, c):
    # Ô M8
    # = 0+0
    return X.add(0.0,0.0)

def f16(E, r, c):
    # Ô N8
    # = ((L84="ĐỦ")+(L85="ĐỦ")+(L86="ĐỦ")+(L87="ĐỦ")+(L88="ĐỦ")+(L89="ĐỦ")+(L92="ĐỦ"))+((L124="ĐỦ")+(L125="ĐỦ")+(L126="ĐỦ")+(L127="ĐỦ")+(L128="ĐỦ")+(L129="ĐỦ"
    #   )+(L130="ĐỦ")+(L131="ĐỦ")+(L132="ĐỦ")+(L133="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(22,r+76,c-2),'ĐỦ'),X.eq(E.v(22,r+77,c-2),'ĐỦ')),X.eq(E.v(22,r+78,c-2),'ĐỦ')),X.eq(E.v(22,r+79,c-2),'ĐỦ')),X.eq(E.v(22,r+80,c-2),'ĐỦ')),X.eq(E.v(22,r+81,c-2),'ĐỦ')),X.eq(E.v(22,r+84,c-2),'ĐỦ')),X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(22,r+116,c-2),'ĐỦ'),X.eq(E.v(22,r+117,c-2),'ĐỦ')),X.eq(E.v(22,r+118,c-2),'ĐỦ')),X.eq(E.v(22,r+119,c-2),'ĐỦ')),X.eq(E.v(22,r+120,c-2),'ĐỦ')),X.eq(E.v(22,r+121,c-2),'ĐỦ')),X.eq(E.v(22,r+122,c-2),'ĐỦ')),X.eq(E.v(22,r+123,c-2),'ĐỦ')),X.eq(E.v(22,r+124,c-2),'ĐỦ')),X.eq(E.v(22,r+125,c-2),'ĐỦ')))

def f17(E, r, c):
    # Ô AD8
    # = TRIM(IF(L84="ĐỦ","• "&G84&" ["&B84&", "&H84&"] ","")&IF(L85="ĐỦ","• "&G85&" ["&B85&", "&H85&"] ","")&IF(L86="ĐỦ","• "&G86&" ["&B86&", "&H86&"] ","")&I
    #   F(L87="ĐỦ","• "&G87&" ["&B87&", "&H87&"] ","")&IF(L88="ĐỦ","• "&G88&" ["&B88&", "&H88&"] ","")&IF(L89="ĐỦ","• "&G89&" ["&B89&", "&H89&"] ","")&IF(L92=
    #   "ĐỦ","• "&G92&" ["&B92&", "&H92&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+76,c-23)),' ['),E.v(22,r+76,c-28)),', '),E.v(22,r+76,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+76,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+77,c-23)),' ['),E.v(22,r+77,c-28)),', '),E.v(22,r+77,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+77,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+78,c-23)),' ['),E.v(22,r+78,c-28)),', '),E.v(22,r+78,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+78,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+79,c-23)),' ['),E.v(22,r+79,c-28)),', '),E.v(22,r+79,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+79,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+80,c-23)),' ['),E.v(22,r+80,c-28)),', '),E.v(22,r+80,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+80,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+81,c-23)),' ['),E.v(22,r+81,c-28)),', '),E.v(22,r+81,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+81,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+84,c-23)),' ['),E.v(22,r+84,c-28)),', '),E.v(22,r+84,c-22)),'] ') if X.truthy(X.eq(E.v(22,r+84,c-18),'ĐỦ')) else '')))

def f18(E, r, c):
    # Ô AE8
    # = TRIM(IF(L124="ĐỦ","• "&G124&" ["&B124&", "&H124&"] ","")&IF(L125="ĐỦ","• "&G125&" ["&B125&", "&H125&"] ","")&IF(L126="ĐỦ","• "&G126&" ["&B126&", "&H12
    #   6&"] ","")&IF(L127="ĐỦ","• "&G127&" ["&B127&", "&H127&"] ","")&IF(L128="ĐỦ","• "&G128&" ["&B128&", "&H128&"] ","")&IF(L129="ĐỦ","• "&G129&" ["&B129&",
    #    "&H129&"] ","")&IF(L130="ĐỦ","• "&G130&" ["&B130&", "&H130&"] ","")&IF(L131="ĐỦ","• "&G131&" ["&B131&", "&H131&"] ","")&IF(L132="ĐỦ","• "&G132&" ["&B
    #   132&", "&H132&"] ","")&IF(L133="ĐỦ","• "&G133&" ["&B133&", "&H133&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+116,c-24)),' ['),E.v(22,r+116,c-29)),', '),E.v(22,r+116,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+116,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+117,c-24)),' ['),E.v(22,r+117,c-29)),', '),E.v(22,r+117,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+117,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+118,c-24)),' ['),E.v(22,r+118,c-29)),', '),E.v(22,r+118,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+118,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+119,c-24)),' ['),E.v(22,r+119,c-29)),', '),E.v(22,r+119,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+119,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+120,c-24)),' ['),E.v(22,r+120,c-29)),', '),E.v(22,r+120,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+120,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+121,c-24)),' ['),E.v(22,r+121,c-29)),', '),E.v(22,r+121,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+121,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+122,c-24)),' ['),E.v(22,r+122,c-29)),', '),E.v(22,r+122,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+122,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+123,c-24)),' ['),E.v(22,r+123,c-29)),', '),E.v(22,r+123,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+123,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+124,c-24)),' ['),E.v(22,r+124,c-29)),', '),E.v(22,r+124,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+124,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(22,r+125,c-24)),' ['),E.v(22,r+125,c-29)),', '),E.v(22,r+125,c-23)),'] ') if X.truthy(X.eq(E.v(22,r+125,c-19),'ĐỦ')) else '')))

def f19(E, r, c):
    # Ô B13
    # = INDEX(HANH_TEN,MOD(0,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(0.0, 5.0),1.0))

def f20(E, r, c):
    # Ô C13
    # = CHOOSE(MOD(0,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại t
    #   ràng","thận – bàng quang, các tạng vùng bụng, chi dưới")
    return X.CHOOSE(X.add(X.MOD(0.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')

def f21(E, r, c):
    # Ô D13
    # = INDEX(DUNG_HY_KY!$C$24:$C$28,1)
    return X.INDEX(E.rg(8,24,3,28,3), 1.0)

def f22(E, r, c):
    # Ô E13
    # = IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF
    #   (0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần","nhàn thần")))
    return ('dụng thần' if X.truthy(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq(0.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'nhàn thần')))

def f23(E, r, c):
    # Ô F13
    # = IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1,INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU),"BẤT CẬP – bị Kim ("&INDEX(DUNG_HY_KY!$C$24:$C$28,4)&" chữ) khắc",
    #   IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU,0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),"THÁI QUÁ – là kỵ thần",IF(INDEX(DUNG_HY_
    #   KY!$C$24:$C$28,1)>=P_NHIEU,"Vượng (nhiều chữ)",IF(INDEX(DUNG_HY_KY!$C$24:$C$28,1)=0,"Không có trong tứ trụ","Bình hòa"))))
    return (X.cat(X.cat('BẤT CẬP – bị Kim (',X.INDEX(E.rg(8,24,3,28,3), 4.0)),' chữ) khắc') if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)))) else ('THÁI QUÁ – là kỵ thần' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)), X.eq(0.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ('Vượng (nhiều chữ)' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2))) else ('Không có trong tứ trụ' if X.truthy(X.eq(X.INDEX(E.rg(8,24,3,28,3), 1.0),0.0)) else 'Bình hòa'))))

def f24(E, r, c):
    # Ô G13
    # = IF(LEFT(F13,3)="BẤT","""Mộc bị kim khắc thì vùng thắt lưng bị bệnh""; ""người có bệnh đau gân cốt là vì mộc kim làm thương tổn"" (P2 tr.558)"&" → chú 
    #   ý "&CHOOSE(MOD(0,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đ
    #   ại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&".",IF(LEFT(F13,3)="THÁ","Hành kỵ thần quá nhiều: tạng "&CHOOSE(MOD(0,5)+1,"gan – mật (vù
    #   ng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạ
    #   ng vùng bụng, chi dưới")&" dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).",IF(LEFT(F13,3)="Khô","Thi
    #   ếu hành này: khi năm/vận khắc mạnh hành "&INDEX(HANH_TEN,MOD(0,5)+1)&" cần lưu ý "&CHOOSE(MOD(0,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim kh
    #   ắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" (suy luậ
    #   n từ P2 tr.558 – giả định).","Không có câu bệnh tật nào của sách ứng với tình trạng này.")))
    return (X.cat(X.cat(X.cat('"Mộc bị kim khắc thì vùng thắt lưng bị bệnh"; "người có bệnh đau gân cốt là vì mộc kim làm thương tổn" (P2 tr.558)',' → chú ý '),X.CHOOSE(X.add(X.MOD(0.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'.') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'BẤT')) else (X.cat(X.cat('Hành kỵ thần quá nhiều: tạng ',X.CHOOSE(X.add(X.MOD(0.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'THÁ')) else (X.cat(X.cat(X.cat(X.cat('Thiếu hành này: khi năm/vận khắc mạnh hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(0.0, 5.0),1.0))),' cần lưu ý '),X.CHOOSE(X.add(X.MOD(0.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' (suy luận từ P2 tr.558 – giả định).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'Khô')) else 'Không có câu bệnh tật nào của sách ứng với tình trạng này.')))

def f25(E, r, c):
    # Ô B14
    # = INDEX(HANH_TEN,MOD(1,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(1.0, 5.0),1.0))

def f26(E, r, c):
    # Ô C14
    # = CHOOSE(MOD(1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại t
    #   ràng","thận – bàng quang, các tạng vùng bụng, chi dưới")
    return X.CHOOSE(X.add(X.MOD(1.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')

def f27(E, r, c):
    # Ô D14
    # = INDEX(DUNG_HY_KY!$C$24:$C$28,2)
    return X.INDEX(E.rg(8,24,3,28,3), 2.0)

def f28(E, r, c):
    # Ô E14
    # = IF(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF
    #   (1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần","nhàn thần")))
    return ('dụng thần' if X.truthy(X.eq(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq(1.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq(1.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'nhàn thần')))

def f29(E, r, c):
    # Ô F14
    # = IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,2)<=1,INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),"BẤT CẬP – bị Thủy ("&INDEX(DUNG_HY_KY!$C$24:$C$28,5)&" chữ) khắc"
    #   ,IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU,1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),"THÁI QUÁ – là kỵ thần",IF(INDEX(DUNG_HY
    #   _KY!$C$24:$C$28,2)>=P_NHIEU,"Vượng (nhiều chữ)",IF(INDEX(DUNG_HY_KY!$C$24:$C$28,2)=0,"Không có trong tứ trụ","Bình hòa"))))
    return (X.cat(X.cat('BẤT CẬP – bị Thủy (',X.INDEX(E.rg(8,24,3,28,3), 5.0)),' chữ) khắc') if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 2.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))) else ('THÁI QUÁ – là kỵ thần' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.eq(1.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ('Vượng (nhiều chữ)' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2))) else ('Không có trong tứ trụ' if X.truthy(X.eq(X.INDEX(E.rg(8,24,3,28,3), 2.0),0.0)) else 'Bình hòa'))))

def f30(E, r, c):
    # Ô G14
    # = IF(LEFT(F14,3)="BẤT","""Hoả bị thuỷ khắc thì bị bệnh về mắt""; ""mắt mờ là hoả bị thuỷ khắc"" (P2 tr.558)"&" → chú ý "&CHOOSE(MOD(1,5)+1,"gan – mật (v
    #   ùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các t
    #   ạng vùng bụng, chi dưới")&".",IF(LEFT(F14,3)="THÁ","Hành kỵ thần quá nhiều: tạng "&CHOOSE(MOD(1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim kh
    #   ắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" dễ thái 
    #   quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).",IF(LEFT(F14,3)="Khô","Thiếu hành này: khi năm/vận khắc mạnh 
    #   hành "&INDEX(HANH_TEN,MOD(1,5)+1)&" cần lưu ý "&CHOOSE(MOD(1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị 
    #   (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" (suy luận từ P2 tr.558 – giả định).","Không
    #    có câu bệnh tật nào của sách ứng với tình trạng này.")))
    return (X.cat(X.cat(X.cat('"Hoả bị thuỷ khắc thì bị bệnh về mắt"; "mắt mờ là hoả bị thuỷ khắc" (P2 tr.558)',' → chú ý '),X.CHOOSE(X.add(X.MOD(1.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'.') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'BẤT')) else (X.cat(X.cat('Hành kỵ thần quá nhiều: tạng ',X.CHOOSE(X.add(X.MOD(1.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'THÁ')) else (X.cat(X.cat(X.cat(X.cat('Thiếu hành này: khi năm/vận khắc mạnh hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(1.0, 5.0),1.0))),' cần lưu ý '),X.CHOOSE(X.add(X.MOD(1.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' (suy luận từ P2 tr.558 – giả định).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'Khô')) else 'Không có câu bệnh tật nào của sách ứng với tình trạng này.')))

def f31(E, r, c):
    # Ô B15
    # = INDEX(HANH_TEN,MOD(2,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(2.0, 5.0),1.0))

def f32(E, r, c):
    # Ô C15
    # = CHOOSE(MOD(2,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại t
    #   ràng","thận – bàng quang, các tạng vùng bụng, chi dưới")
    return X.CHOOSE(X.add(X.MOD(2.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')

def f33(E, r, c):
    # Ô D15
    # = INDEX(DUNG_HY_KY!$C$24:$C$28,3)
    return X.INDEX(E.rg(8,24,3,28,3), 3.0)

def f34(E, r, c):
    # Ô E15
    # = IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF
    #   (2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần","nhàn thần")))
    return ('dụng thần' if X.truthy(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq(2.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'nhàn thần')))

def f35(E, r, c):
    # Ô F15
    # = IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1,INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU),"BẤT CẬP – bị Mộc ("&INDEX(DUNG_HY_KY!$C$24:$C$28,1)&" chữ) khắc",
    #   IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU,2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),"THÁI QUÁ – là kỵ thần",IF(INDEX(DUNG_HY_
    #   KY!$C$24:$C$28,3)>=P_NHIEU,"Vượng (nhiều chữ)",IF(INDEX(DUNG_HY_KY!$C$24:$C$28,3)=0,"Không có trong tứ trụ","Bình hòa"))))
    return (X.cat(X.cat('BẤT CẬP – bị Mộc (',X.INDEX(E.rg(8,24,3,28,3), 1.0)),' chữ) khắc') if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)))) else ('THÁI QUÁ – là kỵ thần' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)), X.eq(2.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ('Vượng (nhiều chữ)' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2))) else ('Không có trong tứ trụ' if X.truthy(X.eq(X.INDEX(E.rg(8,24,3,28,3), 3.0),0.0)) else 'Bình hòa'))))

def f36(E, r, c):
    # Ô G15
    # = IF(LEFT(F15,3)="BẤT","""Thổ hư gặp mộc vượng là tì vị bị tổn thương"" (P2 tr.558)"&" → chú ý "&CHOOSE(MOD(2,5)+1,"gan – mật (vùng thắt lưng, gân cốt k
    #   hi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới"
    #   )&".",IF(LEFT(F15,3)="THÁ","Hành kỵ thần quá nhiều: tạng "&CHOOSE(MOD(2,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, m
    #   ắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" dễ thái quá; khi gặp năm/vận cùn
    #   g hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).",IF(LEFT(F15,3)="Khô","Thiếu hành này: khi năm/vận khắc mạnh hành "&INDEX(HANH_TEN,MO
    #   D(2,5)+1)&" cần lưu ý "&CHOOSE(MOD(2,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu h
    #   óa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" (suy luận từ P2 tr.558 – giả định).","Không có câu bệnh tật nào của
    #    sách ứng với tình trạng này.")))
    return (X.cat(X.cat(X.cat('"Thổ hư gặp mộc vượng là tì vị bị tổn thương" (P2 tr.558)',' → chú ý '),X.CHOOSE(X.add(X.MOD(2.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'.') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'BẤT')) else (X.cat(X.cat('Hành kỵ thần quá nhiều: tạng ',X.CHOOSE(X.add(X.MOD(2.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'THÁ')) else (X.cat(X.cat(X.cat(X.cat('Thiếu hành này: khi năm/vận khắc mạnh hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(2.0, 5.0),1.0))),' cần lưu ý '),X.CHOOSE(X.add(X.MOD(2.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' (suy luận từ P2 tr.558 – giả định).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'Khô')) else 'Không có câu bệnh tật nào của sách ứng với tình trạng này.')))

def f37(E, r, c):
    # Ô B16
    # = INDEX(HANH_TEN,MOD(3,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(3.0, 5.0),1.0))

def f38(E, r, c):
    # Ô C16
    # = CHOOSE(MOD(3,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại t
    #   ràng","thận – bàng quang, các tạng vùng bụng, chi dưới")
    return X.CHOOSE(X.add(X.MOD(3.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')

def f39(E, r, c):
    # Ô D16
    # = INDEX(DUNG_HY_KY!$C$24:$C$28,4)
    return X.INDEX(E.rg(8,24,3,28,3), 4.0)

def f40(E, r, c):
    # Ô E16
    # = IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF
    #   (3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần","nhàn thần")))
    return ('dụng thần' if X.truthy(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq(3.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'nhàn thần')))

def f41(E, r, c):
    # Ô F16
    # = IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,4)<=1,INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),"BẤT CẬP – bị Hỏa ("&INDEX(DUNG_HY_KY!$C$24:$C$28,2)&" chữ) khắc",
    #   IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU,3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),"THÁI QUÁ – là kỵ thần",IF(INDEX(DUNG_HY_
    #   KY!$C$24:$C$28,4)>=P_NHIEU,"Vượng (nhiều chữ)",IF(INDEX(DUNG_HY_KY!$C$24:$C$28,4)=0,"Không có trong tứ trụ","Bình hòa"))))
    return (X.cat(X.cat('BẤT CẬP – bị Hỏa (',X.INDEX(E.rg(8,24,3,28,3), 2.0)),' chữ) khắc') if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 4.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else ('THÁI QUÁ – là kỵ thần' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)), X.eq(3.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ('Vượng (nhiều chữ)' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2))) else ('Không có trong tứ trụ' if X.truthy(X.eq(X.INDEX(E.rg(8,24,3,28,3), 4.0),0.0)) else 'Bình hòa'))))

def f42(E, r, c):
    # Ô G16
    # = IF(LEFT(F16,3)="BẤT","""Kim yếu gặp hoả vượng là có bệnh về máu""; ""hoả khắc kim thì phổi bị bệnh"" (P2 tr.558)"&" → chú ý "&CHOOSE(MOD(3,5)+1,"gan –
    #    mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang
    #   , các tạng vùng bụng, chi dưới")&".",IF(LEFT(F16,3)="THÁ","Hành kỵ thần quá nhiều: tạng "&CHOOSE(MOD(3,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị
    #    kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" d
    #   ễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).",IF(LEFT(F16,3)="Khô","Thiếu hành này: khi năm/vận khắ
    #   c mạnh hành "&INDEX(HANH_TEN,MOD(3,5)+1)&" cần lưu ý "&CHOOSE(MOD(3,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt",
    #   "tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" (suy luận từ P2 tr.558 – giả định)."
    #   ,"Không có câu bệnh tật nào của sách ứng với tình trạng này.")))
    return (X.cat(X.cat(X.cat('"Kim yếu gặp hoả vượng là có bệnh về máu"; "hoả khắc kim thì phổi bị bệnh" (P2 tr.558)',' → chú ý '),X.CHOOSE(X.add(X.MOD(3.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'.') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'BẤT')) else (X.cat(X.cat('Hành kỵ thần quá nhiều: tạng ',X.CHOOSE(X.add(X.MOD(3.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'THÁ')) else (X.cat(X.cat(X.cat(X.cat('Thiếu hành này: khi năm/vận khắc mạnh hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(3.0, 5.0),1.0))),' cần lưu ý '),X.CHOOSE(X.add(X.MOD(3.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' (suy luận từ P2 tr.558 – giả định).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'Khô')) else 'Không có câu bệnh tật nào của sách ứng với tình trạng này.')))

def f43(E, r, c):
    # Ô B17
    # = INDEX(HANH_TEN,MOD(4,5)+1)
    return X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(4.0, 5.0),1.0))

def f44(E, r, c):
    # Ô C17
    # = CHOOSE(MOD(4,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại t
    #   ràng","thận – bàng quang, các tạng vùng bụng, chi dưới")
    return X.CHOOSE(X.add(X.MOD(4.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')

def f45(E, r, c):
    # Ô D17
    # = INDEX(DUNG_HY_KY!$C$24:$C$28,5)
    return X.INDEX(E.rg(8,24,3,28,3), 5.0)

def f46(E, r, c):
    # Ô E17
    # = IF(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF
    #   (4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần","nhàn thần")))
    return ('dụng thần' if X.truthy(X.eq(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq(4.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq(4.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else 'nhàn thần')))

def f47(E, r, c):
    # Ô F17
    # = IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1,INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU),"BẤT CẬP – bị Thổ ("&INDEX(DUNG_HY_KY!$C$24:$C$28,3)&" chữ) khắc",
    #   IF(AND(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU,4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),"THÁI QUÁ – là kỵ thần",IF(INDEX(DUNG_HY_
    #   KY!$C$24:$C$28,5)>=P_NHIEU,"Vượng (nhiều chữ)",IF(INDEX(DUNG_HY_KY!$C$24:$C$28,5)=0,"Không có trong tứ trụ","Bình hòa"))))
    return (X.cat(X.cat('BẤT CẬP – bị Thổ (',X.INDEX(E.rg(8,24,3,28,3), 3.0)),' chữ) khắc') if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)))) else ('THÁI QUÁ – là kỵ thần' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)), X.eq(4.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ('Vượng (nhiều chữ)' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))) else ('Không có trong tứ trụ' if X.truthy(X.eq(X.INDEX(E.rg(8,24,3,28,3), 5.0),0.0)) else 'Bình hòa'))))

def f48(E, r, c):
    # Ô G17
    # = IF(LEFT(F17,3)="BẤT","""Thổ khắc thuỷ thì các tạng vùng bụng bị bệnh""; ""nhâm quý thêm thổ vượng là chi dưới bị thương tổn"" (P2 tr.558-559)"&" → chú
    #    ý "&CHOOSE(MOD(4,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – 
    #   đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&".",IF(LEFT(F17,3)="THÁ","Hành kỵ thần quá nhiều: tạng "&CHOOSE(MOD(4,5)+1,"gan – mật (v
    #   ùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các t
    #   ạng vùng bụng, chi dưới")&" dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).",IF(LEFT(F17,3)="Khô","Th
    #   iếu hành này: khi năm/vận khắc mạnh hành "&INDEX(HANH_TEN,MOD(4,5)+1)&" cần lưu ý "&CHOOSE(MOD(4,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim k
    #   hắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" (suy lu
    #   ận từ P2 tr.558 – giả định).","Không có câu bệnh tật nào của sách ứng với tình trạng này.")))
    return (X.cat(X.cat(X.cat('"Thổ khắc thuỷ thì các tạng vùng bụng bị bệnh"; "nhâm quý thêm thổ vượng là chi dưới bị thương tổn" (P2 tr.558-559)',' → chú ý '),X.CHOOSE(X.add(X.MOD(4.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'.') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'BẤT')) else (X.cat(X.cat('Hành kỵ thần quá nhiều: tạng ',X.CHOOSE(X.add(X.MOD(4.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' dễ thái quá; khi gặp năm/vận cùng hành càng nặng (đề cương phối thang mục 1, 3 – P2 tr.646-647).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'THÁ')) else (X.cat(X.cat(X.cat(X.cat('Thiếu hành này: khi năm/vận khắc mạnh hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(4.0, 5.0),1.0))),' cần lưu ý '),X.CHOOSE(X.add(X.MOD(4.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' (suy luận từ P2 tr.558 – giả định).') if X.truthy(X.eq(X.LEFT(E.v(22,r,c-1), 3.0),'Khô')) else 'Không có câu bệnh tật nào của sách ứng với tình trạng này.')))

def f49(E, r, c):
    # Ô C18
    # = IF(TRIM(IF(OR(LEFT(F13,3)="BẤT",LEFT(F13,3)="THÁ"),CHOOSE(MOD(0,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì 
    #   vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F13&"); ","")&IF(OR(LEFT(F14,3)="BẤT"
    #   ,LEFT(F14,3)="THÁ"),CHOOSE(MOD(1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)"
    #   ,"phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F14&"); ","")&IF(OR(LEFT(F15,3)="BẤT",LEFT(F15,3)="THÁ"),CHOOSE(MOD(
    #   2,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận
    #    – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F15&"); ","")&IF(OR(LEFT(F16,3)="BẤT",LEFT(F16,3)="THÁ"),CHOOSE(MOD(3,5)+1,"gan – mật (vùng thắt lư
    #   ng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụ
    #   ng, chi dưới")&" ("&F16&"); ","")&IF(OR(LEFT(F17,3)="BẤT",LEFT(F17,3)="THÁ"),CHOOSE(MOD(4,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","
    #   tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F17&"); ","
    #   "))="","Không hành nào bất cập bị khắc hay thái quá là kỵ thần → theo sách không có tạng phủ nào bị nhấn mạnh.","Tạng phủ cần chú ý: "&TRIM(IF(OR(LEFT
    #   (F13,3)="BẤT",LEFT(F13,3)="THÁ"),CHOOSE(MOD(0,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dà
    #   y, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F13&"); ","")&IF(OR(LEFT(F14,3)="BẤT",LEFT(F14,3)="THÁ"
    #   ),CHOOSE(MOD(1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại
    #    tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F14&"); ","")&IF(OR(LEFT(F15,3)="BẤT",LEFT(F15,3)="THÁ"),CHOOSE(MOD(2,5)+1,"gan – mật 
    #   (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các
    #    tạng vùng bụng, chi dưới")&" ("&F15&"); ","")&IF(OR(LEFT(F16,3)="BẤT",LEFT(F16,3)="THÁ"),CHOOSE(MOD(3,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị
    #    kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" (
    #   "&F16&"); ","")&IF(OR(LEFT(F17,3)="BẤT",LEFT(F17,3)="THÁ"),CHOOSE(MOD(4,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, m
    #   ắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" ("&F17&"); ","")))&IF(L90="ĐỦ","
    #    Ngoài phép xét theo hành, sách còn có câu riêng: Canh bị Bính/Đinh khắc → đại tràng (BT29, P2 tr.559).","")&IF(L91="ĐỦ"," Ngoài phép xét theo hành, s
    #   ách còn có câu riêng: Giáp gặp hỏa nhiều → thần kinh (BT30, P2 tr.559).","")&IF(L92="ĐỦ"," Ngoài phép xét theo hành, sách còn có câu riêng: Giáp bị kh
    #   ắc → đề phòng thương ở đầu (BT31, P2 tr.565).","")
    return X.cat(X.cat(X.cat(('Không hành nào bất cập bị khắc hay thái quá là kỵ thần → theo sách không có tạng phủ nào bị nhấn mạnh.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(0.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-5,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-5,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-5,c+3), 3.0),'THÁ'))) else ''),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(1.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-4,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-4,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-4,c+3), 3.0),'THÁ'))) else '')),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(2.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-3,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-3,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-3,c+3), 3.0),'THÁ'))) else '')),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(3.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-2,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-2,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-2,c+3), 3.0),'THÁ'))) else '')),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(4.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-1,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-1,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-1,c+3), 3.0),'THÁ'))) else ''))),'')) else X.cat('Tạng phủ cần chú ý: ',X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(0.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-5,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-5,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-5,c+3), 3.0),'THÁ'))) else ''),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(1.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-4,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-4,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-4,c+3), 3.0),'THÁ'))) else '')),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(2.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-3,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-3,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-3,c+3), 3.0),'THÁ'))) else '')),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(3.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-2,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-2,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-2,c+3), 3.0),'THÁ'))) else '')),(X.cat(X.cat(X.cat(X.CHOOSE(X.add(X.MOD(4.0, 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới'),' ('),E.v(22,r-1,c+3)),'); ') if X.truthy(X.OR(X.eq(X.LEFT(E.v(22,r-1,c+3), 3.0),'BẤT'), X.eq(X.LEFT(E.v(22,r-1,c+3), 3.0),'THÁ'))) else ''))))),(' Ngoài phép xét theo hành, sách còn có câu riêng: Canh bị Bính/Đinh khắc → đại tràng (BT29, P2 tr.559).' if X.truthy(X.eq(E.v(22,r+72,c+9),'ĐỦ')) else '')),(' Ngoài phép xét theo hành, sách còn có câu riêng: Giáp gặp hỏa nhiều → thần kinh (BT30, P2 tr.559).' if X.truthy(X.eq(E.v(22,r+73,c+9),'ĐỦ')) else '')),(' Ngoài phép xét theo hành, sách còn có câu riêng: Giáp bị khắc → đề phòng thương ở đầu (BT31, P2 tr.565).' if X.truthy(X.eq(E.v(22,r+74,c+9),'ĐỦ')) else ''))

def f50(E, r, c):
    # Ô C21
    # = "Vận "&1
    return X.cat('Vận ',1.0)

def f51(E, r, c):
    # Ô D21
    # = "Vận "&2
    return X.cat('Vận ',2.0)

def f52(E, r, c):
    # Ô E21
    # = "Vận "&3
    return X.cat('Vận ',3.0)

def f53(E, r, c):
    # Ô F21
    # = "Vận "&4
    return X.cat('Vận ',4.0)

def f54(E, r, c):
    # Ô G21
    # = "Vận "&5
    return X.cat('Vận ',5.0)

def f55(E, r, c):
    # Ô H21
    # = "Vận "&6
    return X.cat('Vận ',6.0)

def f56(E, r, c):
    # Ô I21
    # = "Vận "&7
    return X.cat('Vận ',7.0)

def f57(E, r, c):
    # Ô J21
    # = "Vận "&8
    return X.cat('Vận ',8.0)

def f58(E, r, c):
    # Ô C22, D22, E22, F22, G22, H22, I22, J22
    # = DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r+13,c),' ('),E.v(14,r+14,c)),'–'),X.add(E.v(14,r+14,c),9.0)),')')

def f59(E, r, c):
    # Ô C23, D23, E23, F23, G23, H23, I23, J23
    # = DAI_VAN!C39
    return E.v(14,r+16,c)

def f60(E, r, c):
    # Ô C24, D24, E24, F24, G24, H24, I24, J24
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!C34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,r+10,c),1.0)),1.0))

def f61(E, r, c):
    # Ô C25, D25, E25, F25, G25, H25, I25, J25, C26, D26, E26, F26 … (16 ô)
    # = MOD(DAI_VAN!C42-TRUNG_GIAN!$E$7,5)
    return X.MOD(X.sub(E.v(14,r+17,c),E.v(23,7,5)), 5.0)

def f62(E, r, c):
    # Ô C27, D27, E27, F27, G27, H27, I27, J27
    # = "Can "&DAI_VAN!C39&": "&IF(AND(C25=1,C23="Thương quan",(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0)),"⚠ TAI HỌA BẤT NGỜ – thương quan gặ
    #   p quan (""quan gặp thương quan, không bệnh cũng tai họa""; ""thương quan gặp quan, là tai họa bất ngờ"" – P2 tr.563, 565)",IF(AND(C25=1,C23="Thương qu
    #   an",(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0)),"⚠ BỆNH LIÊN MIÊN – thương quan gặp vận thương quan (""thương quan gặp thương vận là 
    #   bệnh tật liên miên"" – P2 tr.558)",IF(AND(C25=3,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),"⚠ ĐỀ PHÒNG TAI HỌA, BỆNH NẶNG – thân nhược
    #    gặp vận quan sát (""thân nhược gặp sát mạnh hoặc gặp sát địa nếu không chết cũng tai hoạ nặng"" – P2 tr.559; ""nhật chủ nhược, lại gặp vận hay năm qu
    #   an sát … nếu không đau ốm thì cũng bị … tai nạn,"&" hay kiện tụng"" – P2 tr.646)",IF(AND(C25=2,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=
    #   0)),IF(PP1_DINH_TINH!$F$98>=P_NHIEU,"⚠ NHIỀU BỆNH TẬT – tài nhiều, thân nhược lại hành vận tài (""tài nhiều thân nhược lại hành vận tài thì nhiều bệnh
    #    tật"" – P2 tr.559)","⚠ Khí nhật chủ bất túc – thân nhược gặp vận tài (đề cương phối thang mục 5: ""không ốm đau thì cũng gặp điều xấu"" – P2 tr.647)"
    #   ),IF(AND(C25=1,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),"⚠ Hao tổn sức khỏe – thân nhược gặp vận thực thương (đề cương phối thang mụ
    #   c 4: ""không ốm đau cũng gặp rủi ro"" – P2 tr.647)",IF(AND(C25=0,OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+0
    #   ,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+0,5),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))))),"⚠ Vì bệnh mà hao của 
    #   / tranh chấp – vận tỉ kiếp khi tỉ kiếp là kỵ (""vận gặp tỉ kiếp thì vì bệnh tật mà hết của"" – P2 tr.558; mục 1 – P2 tr.646; giả định G18)",IF(AND(C25
    #   =4,C23="Thiên ấn",(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0)),"⚠ Kiêu đoạt thực – vận thiên ấn gặp thực thần trong mệnh (""kiêu đoạt th
    #   ực thần là có tật"" – P2 tr.559)",IF(AND(C25=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1)),"⚠ Khí thái quá – thân vượng gặp vận ấn (đề 
    #   cương phối thang mục 3: ""không ốm đau thì cũng gặp điều xấu"" – P2 tr.647)",IF(OR(AND(C25=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)
    #   ),AND(C25=0,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0))),"Tốt cho sức khỏe – vận ấn/tỉ kiếp sinh phù thân nhược (""ấn thụ nhiều sinh th
    #   ân là người khoẻ""; ""ngang vai nhiều là người ít bệnh"" – P2 tr.558)","Không có câu bệnh – tai nào của sách ứng với vận này")))))))))&CHAR(10)&"Chi "
    #   &C24&": "&IF(AND(C26=1,C24="Thương quan",(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0)),"⚠ TAI HỌA BẤT NGỜ – thương quan gặp quan (""quan
    #    gặp thương quan, không bệnh cũng tai họa""; ""thương quan gặp quan, là tai họa bất ngờ"" – P2 tr.563, 565)",IF(AND(C26=1,C24="Thương quan",(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0)),"⚠ BỆNH LIÊN MIÊN – thương quan gặp vận thương quan (""thương quan gặp thương vận là bệnh tật liên 
    #   miên"" – P2 tr.558)",IF(AND(C26=3,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),"⚠ ĐỀ PHÒNG TAI HỌA, BỆNH NẶNG – thân nhược gặp vận quan 
    #   sát (""thân nhược gặp sát mạnh hoặc gặp sát địa nếu không chết cũng tai hoạ nặng"" – P2 tr.559; ""nhật chủ nhược, lại gặp vận hay năm quan sát … nếu k
    #   hông đau ốm thì cũng bị … tai nạn,"&" hay kiện tụng"" – P2 tr.646)",IF(AND(C26=2,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),IF(PP1_DIN
    #   H_TINH!$F$98>=P_NHIEU,"⚠ NHIỀU BỆNH TẬT – tài nhiều, thân nhược lại hành vận tài (""tài nhiều thân nhược lại hành vận tài thì nhiều bệnh tật"" – P2 tr
    #   .559)","⚠ Khí nhật chủ bất túc – thân nhược gặp vận tài (đề cương phối thang mục 5: ""không ốm đau thì cũng gặp điều xấu"" – P2 tr.647)"),IF(AND(C26=1
    #   ,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),"⚠ Hao tổn sức khỏe – thân nhược gặp vận thực thương (đề cương phối thang mục 4: ""không ố
    #   m đau cũng gặp rủi ro"" – P2 tr.647)",IF(AND(C26=0,OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+0,5)=IF(DOI_CHI
    #   EU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+0,5),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!
    #   $D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))))),"⚠ Vì bệnh mà hao của / tranh chấp –
    #    vận tỉ kiếp khi tỉ kiếp là kỵ (""vận gặp tỉ kiếp thì vì bệnh tật mà hết của"" – P2 tr.558; mục 1 – P2 tr.646; giả định G18)",IF(AND(C26=4,C24="Thiên 
    #   ấn",(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0)>0)),"⚠ Kiêu đoạt thực – vận thiên ấn gặp thực thần trong mệnh (""kiêu đoạt thực thần là có 
    #   tật"" – P2 tr.559)",IF(AND(C26=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1)),"⚠ Khí thái quá – thân vượng gặp vận ấn (đề cương phối tha
    #   ng mục 3: ""không ốm đau thì cũng gặp điều xấu"" – P2 tr.647)",IF(OR(AND(C26=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),AND(C26=0,(I
    #   F(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0))),"Tốt cho sức khỏe – vận ấn/tỉ kiếp sinh phù thân nhược (""ấn thụ nhiều sinh thân là người kh
    #   oẻ""; ""ngang vai nhiều là người ít bệnh"" – P2 tr.558)","Không có câu bệnh – tai nào của sách ứng với vận này")))))))))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can ',E.v(14,r+12,c)),': '),('⚠ TAI HỌA BẤT NGỜ – thương quan gặp quan ("quan gặp thương quan, không bệnh cũng tai họa"; "thương quan gặp quan, là tai họa bất ngờ" – P2 tr.563, 565)' if X.truthy(X.AND(X.eq(E.v(22,r-2,c),1.0), X.eq(E.v(22,r-4,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0))) else ('⚠ BỆNH LIÊN MIÊN – thương quan gặp vận thương quan ("thương quan gặp thương vận là bệnh tật liên miên" – P2 tr.558)' if X.truthy(X.AND(X.eq(E.v(22,r-2,c),1.0), X.eq(E.v(22,r-4,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0))) else (X.cat('⚠ ĐỀ PHÒNG TAI HỌA, BỆNH NẶNG – thân nhược gặp vận quan sát ("thân nhược gặp sát mạnh hoặc gặp sát địa nếu không chết cũng tai hoạ nặng" – P2 tr.559; "nhật chủ nhược, lại gặp vận hay năm quan sát … nếu không đau ốm thì cũng bị … tai nạn,',' hay kiện tụng" – P2 tr.646)') if X.truthy(X.AND(X.eq(E.v(22,r-2,c),3.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else (('⚠ NHIỀU BỆNH TẬT – tài nhiều, thân nhược lại hành vận tài ("tài nhiều thân nhược lại hành vận tài thì nhiều bệnh tật" – P2 tr.559)' if X.truthy(X.ge(E.v(9,98,6),E.v(24,99,2))) else '⚠ Khí nhật chủ bất túc – thân nhược gặp vận tài (đề cương phối thang mục 5: "không ốm đau thì cũng gặp điều xấu" – P2 tr.647)') if X.truthy(X.AND(X.eq(E.v(22,r-2,c),2.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ('⚠ Hao tổn sức khỏe – thân nhược gặp vận thực thương (đề cương phối thang mục 4: "không ốm đau cũng gặp rủi ro" – P2 tr.647)' if X.truthy(X.AND(X.eq(E.v(22,r-2,c),1.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ('⚠ Vì bệnh mà hao của / tranh chấp – vận tỉ kiếp khi tỉ kiếp là kỵ ("vận gặp tỉ kiếp thì vì bệnh tật mà hết của" – P2 tr.558; mục 1 – P2 tr.646; giả định G18)' if X.truthy(X.AND(X.eq(E.v(22,r-2,c),0.0), X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))) else ('⚠ Kiêu đoạt thực – vận thiên ấn gặp thực thần trong mệnh ("kiêu đoạt thực thần là có tật" – P2 tr.559)' if X.truthy(X.AND(X.eq(E.v(22,r-2,c),4.0), X.eq(E.v(22,r-4,c),'Thiên ấn'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0))) else ('⚠ Khí thái quá – thân vượng gặp vận ấn (đề cương phối thang mục 3: "không ốm đau thì cũng gặp điều xấu" – P2 tr.647)' if X.truthy(X.AND(X.eq(E.v(22,r-2,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else ('Tốt cho sức khỏe – vận ấn/tỉ kiếp sinh phù thân nhược ("ấn thụ nhiều sinh thân là người khoẻ"; "ngang vai nhiều là người ít bệnh" – P2 tr.558)' if X.truthy(X.OR(X.AND(X.eq(E.v(22,r-2,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)), X.AND(X.eq(E.v(22,r-2,c),0.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)))) else 'Không có câu bệnh – tai nào của sách ứng với vận này')))))))))),X.CHAR(10.0)),'Chi '),E.v(22,r-3,c)),': '),('⚠ TAI HỌA BẤT NGỜ – thương quan gặp quan ("quan gặp thương quan, không bệnh cũng tai họa"; "thương quan gặp quan, là tai họa bất ngờ" – P2 tr.563, 565)' if X.truthy(X.AND(X.eq(E.v(22,r-1,c),1.0), X.eq(E.v(22,r-3,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0))) else ('⚠ BỆNH LIÊN MIÊN – thương quan gặp vận thương quan ("thương quan gặp thương vận là bệnh tật liên miên" – P2 tr.558)' if X.truthy(X.AND(X.eq(E.v(22,r-1,c),1.0), X.eq(E.v(22,r-3,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0))) else (X.cat('⚠ ĐỀ PHÒNG TAI HỌA, BỆNH NẶNG – thân nhược gặp vận quan sát ("thân nhược gặp sát mạnh hoặc gặp sát địa nếu không chết cũng tai hoạ nặng" – P2 tr.559; "nhật chủ nhược, lại gặp vận hay năm quan sát … nếu không đau ốm thì cũng bị … tai nạn,',' hay kiện tụng" – P2 tr.646)') if X.truthy(X.AND(X.eq(E.v(22,r-1,c),3.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else (('⚠ NHIỀU BỆNH TẬT – tài nhiều, thân nhược lại hành vận tài ("tài nhiều thân nhược lại hành vận tài thì nhiều bệnh tật" – P2 tr.559)' if X.truthy(X.ge(E.v(9,98,6),E.v(24,99,2))) else '⚠ Khí nhật chủ bất túc – thân nhược gặp vận tài (đề cương phối thang mục 5: "không ốm đau thì cũng gặp điều xấu" – P2 tr.647)') if X.truthy(X.AND(X.eq(E.v(22,r-1,c),2.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ('⚠ Hao tổn sức khỏe – thân nhược gặp vận thực thương (đề cương phối thang mục 4: "không ốm đau cũng gặp rủi ro" – P2 tr.647)' if X.truthy(X.AND(X.eq(E.v(22,r-1,c),1.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ('⚠ Vì bệnh mà hao của / tranh chấp – vận tỉ kiếp khi tỉ kiếp là kỵ ("vận gặp tỉ kiếp thì vì bệnh tật mà hết của" – P2 tr.558; mục 1 – P2 tr.646; giả định G18)' if X.truthy(X.AND(X.eq(E.v(22,r-1,c),0.0), X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))))) else ('⚠ Kiêu đoạt thực – vận thiên ấn gặp thực thần trong mệnh ("kiêu đoạt thực thần là có tật" – P2 tr.559)' if X.truthy(X.AND(X.eq(E.v(22,r-1,c),4.0), X.eq(E.v(22,r-3,c),'Thiên ấn'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0),0.0))) else ('⚠ Khí thái quá – thân vượng gặp vận ấn (đề cương phối thang mục 3: "không ốm đau thì cũng gặp điều xấu" – P2 tr.647)' if X.truthy(X.AND(X.eq(E.v(22,r-1,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else ('Tốt cho sức khỏe – vận ấn/tỉ kiếp sinh phù thân nhược ("ấn thụ nhiều sinh thân là người khoẻ"; "ngang vai nhiều là người ít bệnh" – P2 tr.558)' if X.truthy(X.OR(X.AND(X.eq(E.v(22,r-1,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)), X.AND(X.eq(E.v(22,r-1,c),0.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)))) else 'Không có câu bệnh – tai nào của sách ứng với vận này'))))))))))

def f63(E, r, c):
    # Ô C28
    # = "Đại vận cần đề phòng bệnh tật – tai họa: "&IF(AND(SUMPRODUCT(--ISNUMBER(SEARCH("⚠",C27:J27)))>=3,SUMPRODUCT(--ISNUMBER(SEARCH("⚠",C27:J27)))=COUNTIF(
    #   C22:J22,"?*")),"cả "&COUNTIF(C22:J22,"?*")&" đại vận đều có dấu hiệu – không phân biệt được vận nặng nhẹ (xem từng vận ở dòng BỆNH TẬT – TAI HỌA bên t
    #   rên)",IF(TRIM(IF(ISNUMBER(SEARCH("⚠",C27)),C22&" ","")&IF(ISNUMBER(SEARCH("⚠",D27)),D22&" ","")&IF(ISNUMBER(SEARCH("⚠",E27)),E22&" ","")&IF(ISNUMBER(S
    #   EARCH("⚠",F27)),F22&" ","")&IF(ISNUMBER(SEARCH("⚠",G27)),G22&" ","")&IF(ISNUMBER(SEARCH("⚠",H27)),H22&" ","")&IF(ISNUMBER(SEARCH("⚠",I27)),I22&" ","")
    #   &IF(ISNUMBER(SEARCH("⚠",J27)),J22&" ",""))="","không có",TRIM(IF(ISNUMBER(SEARCH("⚠",C27)),C22&" ","")&IF(ISNUMBER(SEARCH("⚠",D27)),D22&" ","")&IF(ISN
    #   UMBER(SEARCH("⚠",E27)),E22&" ","")&IF(ISNUMBER(SEARCH("⚠",F27)),F22&" ","")&IF(ISNUMBER(SEARCH("⚠",G27)),G22&" ","")&IF(ISNUMBER(SEARCH("⚠",H27)),H22&
    #   " ","")&IF(ISNUMBER(SEARCH("⚠",I27)),I22&" ","")&IF(ISNUMBER(SEARCH("⚠",J27)),J22&" ",""))))&CHAR(10)&"Đại vận thuận cho sức khỏe: "&IF(TRIM(IF(AND(NO
    #   T(ISNUMBER(SEARCH("⚠",C27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",C27))),C22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",D27))),ISNUMBER(SEARCH("Tốt cho sức 
    #   khỏe",D27))),D22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",E27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",E27))),E22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",F2
    #   7))),ISNUMBER(SEARCH("Tốt cho sức khỏe",F27))),F22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",G27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",G27))),G22&" ","")
    #   &IF(AND(NOT(ISNUMBER(SEARCH("⚠",H27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",H27))),H22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",I27))),ISNUMBER(SEARCH("Tố
    #   t cho sức khỏe",I27))),I22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",J27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",J27))),J22&" ",""))="","không có",TRIM(IF(
    #   AND(NOT(ISNUMBER(SEARCH("⚠",C27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",C27))),C22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",D27))),ISNUMBER(SEARCH("Tốt ch
    #   o sức khỏe",D27))),D22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",E27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",E27))),E22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH(
    #   "⚠",F27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",F27))),F22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",G27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",G27))),G22&"
    #    ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",H27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",H27))),H22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",I27))),ISNUMBER(SEAR
    #   CH("Tốt cho sức khỏe",I27))),I22&" ","")&IF(AND(NOT(ISNUMBER(SEARCH("⚠",J27))),ISNUMBER(SEARCH("Tốt cho sức khỏe",J27))),J22&" ","")))
    return X.cat(X.cat(X.cat(X.cat('Đại vận cần đề phòng bệnh tật – tai họa: ',(X.cat(X.cat('cả ',X.COUNTIF(E.rg(22,r-6,c,r-6,c+7), '?*')),' đại vận đều có dấu hiệu – không phân biệt được vận nặng nhẹ (xem từng vận ở dòng BỆNH TẬT – TAI HỌA bên trên)') if X.truthy(X.AND(X.ge(X.SUMPRODUCT(X.A('neg',X.A('neg',X.AF_ISNUMBER(lambda: X.AF_SEARCH('⚠', E.rg(22,r-1,c,r-1,c+7)))))),3.0), X.eq(X.SUMPRODUCT(X.A('neg',X.A('neg',X.AF_ISNUMBER(lambda: X.AF_SEARCH('⚠', E.rg(22,r-1,c,r-1,c+7)))))),X.COUNTIF(E.rg(22,r-6,c,r-6,c+7), '?*')))) else ('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-6,c),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c)))) else ''),(X.cat(E.v(22,r-6,c+1),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+1)))) else '')),(X.cat(E.v(22,r-6,c+2),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+2)))) else '')),(X.cat(E.v(22,r-6,c+3),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+3)))) else '')),(X.cat(E.v(22,r-6,c+4),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+4)))) else '')),(X.cat(E.v(22,r-6,c+5),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+5)))) else '')),(X.cat(E.v(22,r-6,c+6),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+6)))) else '')),(X.cat(E.v(22,r-6,c+7),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+7)))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-6,c),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c)))) else ''),(X.cat(E.v(22,r-6,c+1),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+1)))) else '')),(X.cat(E.v(22,r-6,c+2),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+2)))) else '')),(X.cat(E.v(22,r-6,c+3),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+3)))) else '')),(X.cat(E.v(22,r-6,c+4),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+4)))) else '')),(X.cat(E.v(22,r-6,c+5),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+5)))) else '')),(X.cat(E.v(22,r-6,c+6),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+6)))) else '')),(X.cat(E.v(22,r-6,c+7),' ') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+7)))) else '')))))),X.CHAR(10.0)),'Đại vận thuận cho sức khỏe: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-6,c),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c))))) else ''),(X.cat(E.v(22,r-6,c+1),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+1)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+1))))) else '')),(X.cat(E.v(22,r-6,c+2),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+2)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+2))))) else '')),(X.cat(E.v(22,r-6,c+3),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+3)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+3))))) else '')),(X.cat(E.v(22,r-6,c+4),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+4)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+4))))) else '')),(X.cat(E.v(22,r-6,c+5),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+5)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+5))))) else '')),(X.cat(E.v(22,r-6,c+6),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+6)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+6))))) else '')),(X.cat(E.v(22,r-6,c+7),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+7)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+7))))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-6,c),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c))))) else ''),(X.cat(E.v(22,r-6,c+1),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+1)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+1))))) else '')),(X.cat(E.v(22,r-6,c+2),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+2)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+2))))) else '')),(X.cat(E.v(22,r-6,c+3),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+3)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+3))))) else '')),(X.cat(E.v(22,r-6,c+4),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+4)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+4))))) else '')),(X.cat(E.v(22,r-6,c+5),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+5)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+5))))) else '')),(X.cat(E.v(22,r-6,c+6),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+6)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+6))))) else '')),(X.cat(E.v(22,r-6,c+7),' ') if X.truthy(X.AND(X.NOT(X.ISNUMBER(lambda: X.SEARCH('⚠', E.v(22,r-1,c+7)))), X.ISNUMBER(lambda: X.SEARCH('Tốt cho sức khỏe', E.v(22,r-1,c+7))))) else '')))))

def f64(E, r, c):
    # Ô B32, B146
    # = DAI_VAN!$C$23+0
    return X.add(E.v(14,23,3),0.0)

def f65(E, r, c):
    # Ô C32, C33, C34, C35, C36, C37
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(E.v(22,r,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f66(E, r, c):
    # Ô D32, D33, D34, D35, D36, D37
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(E.v(22,r,c-2), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f67(E, r, c):
    # Ô E32, E33, E34, E35, E36, E37
    # = IF(AF32=0,"Không khớp loại năm hạn nào",TRIM(IF(R32=1,"đ.1 dụng thần bị khắc (quan gặp thương quan)"&IF("R"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR
    #   $129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&IF(S32=1,"đ.2 tuế vận tịnh lâm (lưu niên trùng đại vận/tiểu vận)"&IF("S"="U"," ("&IFERRO
    #   R(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&IF(T32=1,"đ.3 đại vận xung đề cương (chi tháng) và năm lại xu
    #   ng một trụ"&IF("T"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&IF(U32=1,"đ.4 lưu niên thiê
    #   n khắc địa xung trụ"&IF("U"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&IF(V32=1,"đ.5 năm 
    #   bản mệnh (chi năm = chi năm sinh)"&IF("V"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&IF(W
    #   32=1,"đ.6 trụ ngày xung khắc lưu niên"&IF("W"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&
    #   IF(X32=1,"đ.8 ba xung một / một xung ba"&IF("X"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ",""
    #   )&IF(Y32=1,"đ.10 năm giao nhau của hai đại vận"&IF("Y"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&
    #   "; ","")&IF(Z32=1,"đ.11 nhật can nhập mộ / mộ bị xung"&IF("Z"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&"
    #   )","")&"; ","")&IF(AA32=1,"đ.12 hung sát (kình dương gặp/xung tuế quân)"&IF("AA"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$
    #   30:$A$129,0)),"")&")","")&"; ","")&IF(AB32=1,"đ.13 năm kề năm hung (năm trước/sau có ≥ 2 dấu hiệu nặng)"&IF("AB"="U"," ("&IFERROR(INDEX(LUU_NIEN!$AR$3
    #   0:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")&IF(AC32=1,"đ.14 đại vận không tốt / đại vận và lưu niên đều là kỵ"&IF("AC"="U"," ("
    #   &IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")&")","")&"; ","")))
    return ('Không khớp loại năm hạn nào' if X.truthy(X.eq(E.v(22,r,c+27),0.0)) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('đ.1 dụng thần bị khắc (quan gặp thương quan)',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('R','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+13),1.0)) else ''),(X.cat(X.cat('đ.2 tuế vận tịnh lâm (lưu niên trùng đại vận/tiểu vận)',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('S','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+14),1.0)) else '')),(X.cat(X.cat('đ.3 đại vận xung đề cương (chi tháng) và năm lại xung một trụ',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('T','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+15),1.0)) else '')),(X.cat(X.cat('đ.4 lưu niên thiên khắc địa xung trụ',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('U','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+16),1.0)) else '')),(X.cat(X.cat('đ.5 năm bản mệnh (chi năm = chi năm sinh)',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('V','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+17),1.0)) else '')),(X.cat(X.cat('đ.6 trụ ngày xung khắc lưu niên',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('W','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+18),1.0)) else '')),(X.cat(X.cat('đ.8 ba xung một / một xung ba',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('X','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+19),1.0)) else '')),(X.cat(X.cat('đ.10 năm giao nhau của hai đại vận',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('Y','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+20),1.0)) else '')),(X.cat(X.cat('đ.11 nhật can nhập mộ / mộ bị xung',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('Z','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+21),1.0)) else '')),(X.cat(X.cat('đ.12 hung sát (kình dương gặp/xung tuế quân)',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('AA','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+22),1.0)) else '')),(X.cat(X.cat('đ.13 năm kề năm hung (năm trước/sau có ≥ 2 dấu hiệu nặng)',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('AB','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+23),1.0)) else '')),(X.cat(X.cat('đ.14 đại vận không tốt / đại vận và lưu niên đều là kỵ',(X.cat(X.cat(' (',X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),')') if X.truthy(X.eq('AC','U')) else '')),'; ') if X.truthy(X.eq(E.v(22,r,c+24),1.0)) else ''))))

def f68(E, r, c):
    # Ô F32, F33, F34, F35, F36, F37
    # = IF(AF32>=3,"HẠN NẶNG",IF(AF32=2,"Cần đề phòng",IF(AF32=1,"Chú ý nhẹ","Bình thường")))&IF(AND(AF32>=1,IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(B32,L
    #   UU_NIEN!$A$30:$A$129,0)),"")<>""),CHAR(10)&"Có cứu: "&IFERROR(INDEX(LUU_NIEN!$CB$30:$CB$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""),"")
    return X.cat(('HẠN NẶNG' if X.truthy(X.ge(E.v(22,r,c+26),3.0)) else ('Cần đề phòng' if X.truthy(X.eq(E.v(22,r,c+26),2.0)) else ('Chú ý nhẹ' if X.truthy(X.eq(E.v(22,r,c+26),1.0)) else 'Bình thường'))),(X.cat(X.cat(X.CHAR(10.0),'Có cứu: '),X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(E.v(22,r,c-4), E.rg(15,30,1,129,1), 0.0)), lambda: '')) if X.truthy(X.AND(X.ge(E.v(22,r,c+26),1.0), X.ne(X.IFERROR(lambda: X.INDEX(E.rg(15,30,80,129,80), X.MATCH(E.v(22,r,c-4), E.rg(15,30,1,129,1), 0.0)), lambda: ''),''))) else ''))

def f69(E, r, c):
    # Ô G32, G33, G34, G35, G36, G37
    # = TRIM(IF(AD32=1,"PHẢN NGÂM (tuế/vận thiên khắc địa xung trụ ngày) ","")&IF(AE32=1,"PHỤC NGÂM (tuế/vận trùng can chi trụ ngày)",""))&IF(AND(AD32=0,AE32=
    #   0),"Không","")
    return X.cat(X.TRIM(X.cat(('PHẢN NGÂM (tuế/vận thiên khắc địa xung trụ ngày) ' if X.truthy(X.eq(E.v(22,r,c+23),1.0)) else ''),('PHỤC NGÂM (tuế/vận trùng can chi trụ ngày)' if X.truthy(X.eq(E.v(22,r,c+24),1.0)) else ''))),('Không' if X.truthy(X.AND(X.eq(E.v(22,r,c+23),0.0), X.eq(E.v(22,r,c+24),0.0))) else ''))

def f70(E, r, c):
    # Ô H32, H33, H34, H35, H36, H37
    # = IF(AND((LEFT(F32,3)="Bìn"),AD32=0,AE32=0),"Không cần hóa giải riêng; duy trì sinh hoạt ở phương dụng/hỷ thần.",TRIM(IF(OR(AD32=1,AE32=1,U32=1,W32=1),"
    #   • Nên di chuyển / dời chỗ ở, đổi chỗ giường (""dù là phản ngâm hay phục ngâm, về nguyên tắc nên di chuyển"" – P2 tr.659): tránh phương thái tuế "&INDE
    #   X(CHI_TEN,N32+1)&" ("&INDEX({"Bắc","Đông bắc","Đông bắc","Đông","Đông nam","Đông nam","Nam","Tây nam","Tây nam","Tây","Tây bắc","Tây bắc"},N32+1)&") v
    #   à tuế phá "&INDEX(CHI_TEN,MOD(N32+6,12)+1)&" ("&INDEX({"Bắc","Đông bắc","Đông bắc","Đông","Đông nam","Đông nam","Nam","Tây nam","Tây nam","Tây","Tây b
    #   ắc","Tây bắc"},MOD(N32+6,12)+1)&")"&IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,N32)>0," – tứ trụ có chi "&INDEX(CHI_TEN,N32+1)&" nên kỵ nhất phương thái tuế",IF(
    #   COUNTIF(TRUNG_GIAN!$C$6:$F$6,MOD(N32+6,12))>0," – tứ trụ có chi xung thái tuế nên kỵ nhất phương tuế phá",""))&" (P2 tr.633). ","")&"• Hoạt động nhiều
    #    ở phương dụng thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),5)+1)&" ("&CHOOSE((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!
    #   $D$109,DUNG_HY_KY!$C$109))+1,"Đông","Nam","Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc")&")"&IF(IF(DOI_CHIEU!$C$3
    #   3=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)>=0,", hỷ thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),5)+1)&" ("&CHOOS
    #   E((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))+1,"Đông","Nam","Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","T
    #   ây","Bắc")&")","")&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,"; tránh phương kỵ thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,D
    #   UNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),5)+1)&" ("&CHOOSE((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))+1,"Đông","Nam","Trung cung (Thìn–Tuấ
    #   t–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc")&")","")&" (P2 tr.631-632). "&"• Chọn ngày giải hạn từ tháng 11–12 năm "&(B32-1)&" (thán
    #   g tính theo tiết – P2 tr.660-661)"&IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,MOD(N32+6,12))>0,", ưu tiên ngày/giờ "&INDEX(CHI_TEN,MOD(13-N32,12)+1)&" (lục hợp v
    #   ới chi năm) hoặc "&INDEX(CHI_TEN,MOD(13-MOD(N32+6,12),12)+1)&" (lục hợp với chi bị xung)","")&". "&IF(R32=1,"• Quan – thương xung khắc: chọn ngày giờ 
    #   thuộc tài của mệnh (hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&") (P2 tr.661). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C
    #   $7)=0),IFERROR(MOD(IFERROR(INDEX(LUU_NIEN!$P$30:$P$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")-TRUNG_GIAN!$E$7,5)=3,FALSE())),"• Quan sát khắc thân: c
    #   họn ngày giờ thuộc ấn (hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)&") (P2 tr.661). ","")))
    return ('Không cần hóa giải riêng; duy trì sinh hoạt ở phương dụng/hỷ thần.' if X.truthy(X.AND(X.eq(X.LEFT(E.v(22,r,c-2), 3.0),'Bìn'), X.eq(E.v(22,r,c+22),0.0), X.eq(E.v(22,r,c+23),0.0))) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• Nên di chuyển / dời chỗ ở, đổi chỗ giường ("dù là phản ngâm hay phục ngâm, về nguyên tắc nên di chuyển" – P2 tr.659): tránh phương thái tuế ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(22,r,c+6),1.0))),' ('),X.INDEX(X.Arr([['Bắc', 'Đông bắc', 'Đông bắc', 'Đông', 'Đông nam', 'Đông nam', 'Nam', 'Tây nam', 'Tây nam', 'Tây', 'Tây bắc', 'Tây bắc']]), X.add(E.v(22,r,c+6),1.0))),') và tuế phá '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(E.v(22,r,c+6),6.0), 12.0),1.0))),' ('),X.INDEX(X.Arr([['Bắc', 'Đông bắc', 'Đông bắc', 'Đông', 'Đông nam', 'Đông nam', 'Nam', 'Tây nam', 'Tây nam', 'Tây', 'Tây bắc', 'Tây bắc']]), X.add(X.MOD(X.add(E.v(22,r,c+6),6.0), 12.0),1.0))),')'),(X.cat(X.cat(' – tứ trụ có chi ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(22,r,c+6),1.0))),' nên kỵ nhất phương thái tuế') if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), E.v(22,r,c+6)),0.0)) else (' – tứ trụ có chi xung thái tuế nên kỵ nhất phương tuế phá' if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.add(E.v(22,r,c+6),6.0), 12.0)),0.0)) else ''))),' (P2 tr.633). ') if X.truthy(X.OR(X.eq(E.v(22,r,c+22),1.0), X.eq(E.v(22,r,c+23),1.0), X.eq(E.v(22,r,c+13),1.0), X.eq(E.v(22,r,c+15),1.0))) else ''),'• Hoạt động nhiều ở phương dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),')'),(X.cat(X.cat(X.cat(X.cat(', hỷ thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),')') if X.truthy(X.ge((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else '')),(X.cat(X.cat(X.cat(X.cat('; tránh phương kỵ thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),')') if X.truthy(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0)) else '')),' (P2 tr.631-632). '),'• Chọn ngày giải hạn từ tháng 11–12 năm '),X.sub(E.v(22,r,c-6),1.0)),' (tháng tính theo tiết – P2 tr.660-661)'),(X.cat(X.cat(X.cat(X.cat(', ưu tiên ngày/giờ ',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(13.0,E.v(22,r,c+6)), 12.0),1.0))),' (lục hợp với chi năm) hoặc '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(13.0,X.MOD(X.add(E.v(22,r,c+6),6.0), 12.0)), 12.0),1.0))),' (lục hợp với chi bị xung)') if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.add(E.v(22,r,c+6),6.0), 12.0)),0.0)) else '')),'. '),(X.cat(X.cat('• Quan – thương xung khắc: chọn ngày giờ thuộc tài của mệnh (hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),') (P2 tr.661). ') if X.truthy(X.eq(E.v(22,r,c+10),1.0)) else '')),(X.cat(X.cat('• Quan sát khắc thân: chọn ngày giờ thuộc ấn (hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),') (P2 tr.661). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.IFERROR(lambda: X.eq(X.MOD(X.sub(X.IFERROR(lambda: X.INDEX(E.rg(15,30,16,129,16), X.MATCH(E.v(22,r,c-6), E.rg(15,30,1,129,1), 0.0)), lambda: ''),E.v(23,7,5)), 5.0),3.0), lambda: X.FALSE()))) else ''))))

def f71(E, r, c):
    # Ô N32, N33, N34, N35, N36, N37
    # = IFERROR(INDEX(LUU_NIEN!$K$30:$K$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,11,129,11), X.MATCH(E.v(22,r,c-12), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f72(E, r, c):
    # Ô O32, O33, O34, O35, O36, O37
    # = IFERROR(INDEX(LUU_NIEN!$J$30:$J$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,10,129,10), X.MATCH(E.v(22,r,c-13), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f73(E, r, c):
    # Ô P32, P33, P34, P35, P36, P37
    # = IFERROR(INDEX(LUU_NIEN!$G$30:$G$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,7,129,7), X.MATCH(E.v(22,r,c-14), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f74(E, r, c):
    # Ô Q32, Q33, Q34, Q35, Q36, Q37
    # = IFERROR(INDEX(LUU_NIEN!$F$30:$F$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,6,129,6), X.MATCH(E.v(22,r,c-15), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f75(E, r, c):
    # Ô R32, R33, R34, R35, R36, R37
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$BK$30:$BK$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))=1,1,0)
    return (1.0 if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,63,129,63), X.MATCH(E.v(22,r,c-16), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else 0.0)

def f76(E, r, c):
    # Ô S32, S33, S34, S35, S36, S37
    # = IF(OR(N(IFERROR(INDEX(LUU_NIEN!$AP$30:$AP$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))=1,N(IFERROR(INDEX(LUU_NIEN!$AQ$30:$AQ$129,MATCH(B32,LUU_NIEN!$A
    #   $30:$A$129,0)),""))=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,42,129,42), X.MATCH(E.v(22,r,c-17), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0), X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,43,129,43), X.MATCH(E.v(22,r,c-17), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0))) else 0.0)

def f77(E, r, c):
    # Ô T32, T33, T34, T35, T36, T37
    # = IFERROR(IF(AND(P32>=0,INDEX(M_XU,P32+1,TRUNG_GIAN!$D$6+1)=1,OR(INDEX(M_XU,N32+1,TRUNG_GIAN!$C$6+1)=1,INDEX(M_XU,N32+1,TRUNG_GIAN!$D$6+1)=1,INDEX(M_XU,
    #   N32+1,TRUNG_GIAN!$E$6+1)=1,INDEX(M_XU,N32+1,TRUNG_GIAN!$F$6+1)=1)),1,0),0)
    return X.IFERROR(lambda: (1.0 if X.truthy(X.AND(X.ge(E.v(22,r,c-4),0.0), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(22,r,c-4),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.OR(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(22,r,c-6),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(22,r,c-6),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(22,r,c-6),1.0), X.add(E.v(23,6,5),1.0)),1.0), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(22,r,c-6),1.0), X.add(E.v(23,6,6),1.0)),1.0)))) else 0.0), lambda: 0.0)

def f78(E, r, c):
    # Ô U32, U33, U34, U35, U36, U37
    # = IF(IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")<>"",1,0)
    return (1.0 if X.truthy(X.ne(X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-19), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else 0.0)

def f79(E, r, c):
    # Ô V32, V33, V34, V35, V36, V37
    # = IF(N32=TRUNG_GIAN!$C$6,1,0)
    return (1.0 if X.truthy(X.eq(E.v(22,r,c-8),E.v(23,6,3))) else 0.0)

def f80(E, r, c):
    # Ô W32, W33, W34, W35, W36, W37
    # = IF(OR(ISNUMBER(SEARCH("Ngày",IFERROR(INDEX(LUU_NIEN!$AR$30:$AR$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))),N(IFERROR(INDEX(LUU_NIEN!$AT$30:$AT$129,M
    #   ATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))=1),1,0)
    return (1.0 if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Ngày', X.IFERROR(lambda: X.INDEX(E.rg(15,30,44,129,44), X.MATCH(E.v(22,r,c-21), E.rg(15,30,1,129,1), 0.0)), lambda: ''))), X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,46,129,46), X.MATCH(E.v(22,r,c-21), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0))) else 0.0)

def f81(E, r, c):
    # Ô X32, X33, X34, X35, X36, X37
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$AU$30:$AU$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))=2,1,0)
    return (1.0 if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,47,129,47), X.MATCH(E.v(22,r,c-22), E.rg(15,30,1,129,1), 0.0)), lambda: '')),2.0)) else 0.0)

def f82(E, r, c):
    # Ô Y32, Y33, Y34, Y35, Y36, Y37
    # = IFERROR(IF(AND(N(IFERROR(INDEX(LUU_NIEN!$D$30:$D$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))>=1,OR(IFERROR(INDEX(LUU_NIEN!$D$30:$D$129,MATCH(B32,LUU_
    #   NIEN!$A$30:$A$129,0)),"")<>IFERROR(INDEX(LUU_NIEN!$D$30:$D$129,MATCH(B32-1,LUU_NIEN!$A$30:$A$129,0)),""),IFERROR(INDEX(LUU_NIEN!$D$30:$D$129,MATCH(B32
    #   ,LUU_NIEN!$A$30:$A$129,0)),"")<>IFERROR(INDEX(LUU_NIEN!$D$30:$D$129,MATCH(B32+1,LUU_NIEN!$A$30:$A$129,0)),""))),1,0),0)
    return X.IFERROR(lambda: (1.0 if X.truthy(X.AND(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,4,129,4), X.MATCH(E.v(22,r,c-23), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0), X.OR(X.ne(X.IFERROR(lambda: X.INDEX(E.rg(15,30,4,129,4), X.MATCH(E.v(22,r,c-23), E.rg(15,30,1,129,1), 0.0)), lambda: ''),X.IFERROR(lambda: X.INDEX(E.rg(15,30,4,129,4), X.MATCH(X.sub(E.v(22,r,c-23),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')), X.ne(X.IFERROR(lambda: X.INDEX(E.rg(15,30,4,129,4), X.MATCH(E.v(22,r,c-23), E.rg(15,30,1,129,1), 0.0)), lambda: ''),X.IFERROR(lambda: X.INDEX(E.rg(15,30,4,129,4), X.MATCH(X.add(E.v(22,r,c-23),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))))) else 0.0), lambda: 0.0)

def f83(E, r, c):
    # Ô Z32, Z33, Z34, Z35, Z36, Z37
    # = IF(OR(N(IFERROR(INDEX(LUU_NIEN!$BB$30:$BB$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))=1,N(IFERROR(INDEX(LUU_NIEN!$BC$30:$BC$129,MATCH(B32,LUU_NIEN!$A
    #   $30:$A$129,0)),""))=1),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,54,129,54), X.MATCH(E.v(22,r,c-24), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0), X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,55,129,55), X.MATCH(E.v(22,r,c-24), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0))) else 0.0)

def f84(E, r, c):
    # Ô AA32, AA33, AA34, AA35, AA36, AA37
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$AZ$30:$AZ$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))=1,1,0)
    return (1.0 if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,52,129,52), X.MATCH(E.v(22,r,c-25), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else 0.0)

def f85(E, r, c):
    # Ô AB32, AB33, AB34, AB35, AB36, AB37
    # = IF(OR(N(IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(B32-1,LUU_NIEN!$A$30:$A$129,0)),""))>=2,N(IFERROR(INDEX(LUU_NIEN!$BM$30:$BM$129,MATCH(B32+1,LUU_NI
    #   EN!$A$30:$A$129,0)),""))>=2),1,0)
    return (1.0 if X.truthy(X.OR(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(X.sub(E.v(22,r,c-26),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),2.0), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,65,129,65), X.MATCH(X.add(E.v(22,r,c-26),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),2.0))) else 0.0)

def f86(E, r, c):
    # Ô AC32, AC33, AC34, AC35, AC36, AC37
    # = IF(OR(N(IF(DOI_CHIEU!$C$33=2,IFERROR(INDEX(LUU_NIEN!$AM$30:$AM$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""),IFERROR(INDEX(LUU_NIEN!$AB$30:$AB$129,MATCH
    #   (B32,LUU_NIEN!$A$30:$A$129,0)),"")))=1,IF(DOI_CHIEU!$C$33=2,IFERROR(INDEX(LUU_NIEN!$AK$30:$AK$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""),IFERROR(INDE
    #   X(LUU_NIEN!$Z$30:$Z$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),""))="Xấu"),1,0)
    return (1.0 if X.truthy(X.OR(X.eq(X.N((X.IFERROR(lambda: X.INDEX(E.rg(15,30,39,129,39), X.MATCH(E.v(22,r,c-27), E.rg(15,30,1,129,1), 0.0)), lambda: '') if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,28,129,28), X.MATCH(E.v(22,r,c-27), E.rg(15,30,1,129,1), 0.0)), lambda: ''))),1.0), X.eq((X.IFERROR(lambda: X.INDEX(E.rg(15,30,37,129,37), X.MATCH(E.v(22,r,c-27), E.rg(15,30,1,129,1), 0.0)), lambda: '') if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,26,129,26), X.MATCH(E.v(22,r,c-27), E.rg(15,30,1,129,1), 0.0)), lambda: '')),'Xấu'))) else 0.0)

def f87(E, r, c):
    # Ô AD32, AD33, AD34, AD35, AD36, AD37
    # = IFERROR(IF(OR(W32=1,AND(P32>=0,OR(MOD(IFERROR(INDEX(LUU_NIEN!$S$30:$S$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")+2,5)=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!
    #   $E$7+2,5)=IFERROR(INDEX(LUU_NIEN!$S$30:$S$129,MATCH(B32,LUU_NIEN!$A$30:$A$129,0)),"")),INDEX(M_XU,P32+1,TRUNG_GIAN!$E$6+1)=1)),1,0),0)
    return X.IFERROR(lambda: (1.0 if X.truthy(X.OR(X.eq(E.v(22,r,c-7),1.0), X.AND(X.ge(E.v(22,r,c-14),0.0), X.OR(X.eq(X.MOD(X.add(X.IFERROR(lambda: X.INDEX(E.rg(15,30,19,129,19), X.MATCH(E.v(22,r,c-28), E.rg(15,30,1,129,1), 0.0)), lambda: ''),2.0), 5.0),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),X.IFERROR(lambda: X.INDEX(E.rg(15,30,19,129,19), X.MATCH(E.v(22,r,c-28), E.rg(15,30,1,129,1), 0.0)), lambda: ''))), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(22,r,c-14),1.0), X.add(E.v(23,6,5),1.0)),1.0)))) else 0.0), lambda: 0.0)

def f88(E, r, c):
    # Ô AE32, AE33, AE34, AE35, AE36, AE37
    # = IFERROR(IF(OR(AND(O32=TRUNG_GIAN!$E$5,N32=TRUNG_GIAN!$E$6),AND(Q32=TRUNG_GIAN!$E$5,P32=TRUNG_GIAN!$E$6)),1,0),0)
    return X.IFERROR(lambda: (1.0 if X.truthy(X.OR(X.AND(X.eq(E.v(22,r,c-16),E.v(23,5,5)), X.eq(E.v(22,r,c-17),E.v(23,6,5))), X.AND(X.eq(E.v(22,r,c-14),E.v(23,5,5)), X.eq(E.v(22,r,c-15),E.v(23,6,5))))) else 0.0), lambda: 0.0)

def f89(E, r, c):
    # Ô AF32, AF33, AF34, AF35, AF36, AF37
    # = R32+S32+T32+U32+V32+W32+X32+Y32+Z32+AA32+AB32+AC32
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(E.v(22,r,c-14),E.v(22,r,c-13)),E.v(22,r,c-12)),E.v(22,r,c-11)),E.v(22,r,c-10)),E.v(22,r,c-9)),E.v(22,r,c-8)),E.v(22,r,c-7)),E.v(22,r,c-6)),E.v(22,r,c-5)),E.v(22,r,c-4)),E.v(22,r,c-3))

def f90(E, r, c):
    # Ô B33, B147
    # = DAI_VAN!$C$23+1
    return X.add(E.v(14,23,3),1.0)

def f91(E, r, c):
    # Ô B34, B148
    # = DAI_VAN!$C$23+2
    return X.add(E.v(14,23,3),2.0)

def f92(E, r, c):
    # Ô B35, B149
    # = DAI_VAN!$C$23+3
    return X.add(E.v(14,23,3),3.0)

def f93(E, r, c):
    # Ô B36, B150
    # = DAI_VAN!$C$23+4
    return X.add(E.v(14,23,3),4.0)

def f94(E, r, c):
    # Ô B37, B151
    # = DAI_VAN!$C$23+5
    return X.add(E.v(14,23,3),5.0)

def f95(E, r, c):
    # Ô C38
    # = "Năm cần đề phòng (≥ 2 loại năm hạn): "&IF(TRIM(IF(AF32>=2,B32&" ("&C32&": "&LEFT(F32,IFERROR(FIND(CHAR(10),F32)-1,LEN(F32)))&IF(OR(AD32=1,AE32=1),", 
    #   "&G32,"")&") ","")&IF(AF33>=2,B33&" ("&C33&": "&LEFT(F33,IFERROR(FIND(CHAR(10),F33)-1,LEN(F33)))&IF(OR(AD33=1,AE33=1),", "&G33,"")&") ","")&IF(AF34>=2
    #   ,B34&" ("&C34&": "&LEFT(F34,IFERROR(FIND(CHAR(10),F34)-1,LEN(F34)))&IF(OR(AD34=1,AE34=1),", "&G34,"")&") ","")&IF(AF35>=2,B35&" ("&C35&": "&LEFT(F35,I
    #   FERROR(FIND(CHAR(10),F35)-1,LEN(F35)))&IF(OR(AD35=1,AE35=1),", "&G35,"")&") ","")&IF(AF36>=2,B36&" ("&C36&": "&LEFT(F36,IFERROR(FIND(CHAR(10),F36)-1,L
    #   EN(F36)))&IF(OR(AD36=1,AE36=1),", "&G36,"")&") ","")&IF(AF37>=2,B37&" ("&C37&": "&LEFT(F37,IFERROR(FIND(CHAR(10),F37)-1,LEN(F37)))&IF(OR(AD37=1,AE37=1
    #   ),", "&G37,"")&") ",""))="","không có",TRIM(IF(AF32>=2,B32&" ("&C32&": "&LEFT(F32,IFERROR(FIND(CHAR(10),F32)-1,LEN(F32)))&IF(OR(AD32=1,AE32=1),", "&G3
    #   2,"")&") ","")&IF(AF33>=2,B33&" ("&C33&": "&LEFT(F33,IFERROR(FIND(CHAR(10),F33)-1,LEN(F33)))&IF(OR(AD33=1,AE33=1),", "&G33,"")&") ","")&IF(AF34>=2,B34
    #   &" ("&C34&": "&LEFT(F34,IFERROR(FIND(CHAR(10),F34)-1,LEN(F34)))&IF(OR(AD34=1,AE34=1),", "&G34,"")&") ","")&IF(AF35>=2,B35&" ("&C35&": "&LEFT(F35,IFERR
    #   OR(FIND(CHAR(10),F35)-1,LEN(F35)))&IF(OR(AD35=1,AE35=1),", "&G35,"")&") ","")&IF(AF36>=2,B36&" ("&C36&": "&LEFT(F36,IFERROR(FIND(CHAR(10),F36)-1,LEN(F
    #   36)))&IF(OR(AD36=1,AE36=1),", "&G36,"")&") ","")&IF(AF37>=2,B37&" ("&C37&": "&LEFT(F37,IFERROR(FIND(CHAR(10),F37)-1,LEN(F37)))&IF(OR(AD37=1,AE37=1),",
    #    "&G37,"")&") ","")))&CHAR(10)&"Năm chỉ cần chú ý nhẹ (1 loại): "&IF(TRIM(IF(AND(AF32=1,AD32=0,AE32=0),B32&" ","")&IF(AND(AF33=1,AD33=0,AE33=0),B33&" 
    #   ","")&IF(AND(AF34=1,AD34=0,AE34=0),B34&" ","")&IF(AND(AF35=1,AD35=0,AE35=0),B35&" ","")&IF(AND(AF36=1,AD36=0,AE36=0),B36&" ","")&IF(AND(AF37=1,AD37=0,
    #   AE37=0),B37&" ",""))="","không có",TRIM(IF(AND(AF32=1,AD32=0,AE32=0),B32&" ","")&IF(AND(AF33=1,AD33=0,AE33=0),B33&" ","")&IF(AND(AF34=1,AD34=0,AE34=0)
    #   ,B34&" ","")&IF(AND(AF35=1,AD35=0,AE35=0),B35&" ","")&IF(AND(AF36=1,AD36=0,AE36=0),B36&" ","")&IF(AND(AF37=1,AD37=0,AE37=0),B37&" ","")))
    return X.cat(X.cat(X.cat(X.cat('Năm cần đề phòng (≥ 2 loại năm hạn): ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-6,c-1),' ('),E.v(22,r-6,c)),': '),X.LEFT(E.v(22,r-6,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-6,c+3)),1.0), lambda: X.LEN(E.v(22,r-6,c+3))))),(X.cat(', ',E.v(22,r-6,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-6,c+27),1.0), X.eq(E.v(22,r-6,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-6,c+29),2.0)) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-5,c-1),' ('),E.v(22,r-5,c)),': '),X.LEFT(E.v(22,r-5,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-5,c+3)),1.0), lambda: X.LEN(E.v(22,r-5,c+3))))),(X.cat(', ',E.v(22,r-5,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-5,c+27),1.0), X.eq(E.v(22,r-5,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-5,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-4,c-1),' ('),E.v(22,r-4,c)),': '),X.LEFT(E.v(22,r-4,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-4,c+3)),1.0), lambda: X.LEN(E.v(22,r-4,c+3))))),(X.cat(', ',E.v(22,r-4,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-4,c+27),1.0), X.eq(E.v(22,r-4,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-4,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-3,c-1),' ('),E.v(22,r-3,c)),': '),X.LEFT(E.v(22,r-3,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-3,c+3)),1.0), lambda: X.LEN(E.v(22,r-3,c+3))))),(X.cat(', ',E.v(22,r-3,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-3,c+27),1.0), X.eq(E.v(22,r-3,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-3,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-2,c-1),' ('),E.v(22,r-2,c)),': '),X.LEFT(E.v(22,r-2,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-2,c+3)),1.0), lambda: X.LEN(E.v(22,r-2,c+3))))),(X.cat(', ',E.v(22,r-2,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-2,c+27),1.0), X.eq(E.v(22,r-2,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-2,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-1,c-1),' ('),E.v(22,r-1,c)),': '),X.LEFT(E.v(22,r-1,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-1,c+3)),1.0), lambda: X.LEN(E.v(22,r-1,c+3))))),(X.cat(', ',E.v(22,r-1,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-1,c+27),1.0), X.eq(E.v(22,r-1,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-1,c+29),2.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-6,c-1),' ('),E.v(22,r-6,c)),': '),X.LEFT(E.v(22,r-6,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-6,c+3)),1.0), lambda: X.LEN(E.v(22,r-6,c+3))))),(X.cat(', ',E.v(22,r-6,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-6,c+27),1.0), X.eq(E.v(22,r-6,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-6,c+29),2.0)) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-5,c-1),' ('),E.v(22,r-5,c)),': '),X.LEFT(E.v(22,r-5,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-5,c+3)),1.0), lambda: X.LEN(E.v(22,r-5,c+3))))),(X.cat(', ',E.v(22,r-5,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-5,c+27),1.0), X.eq(E.v(22,r-5,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-5,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-4,c-1),' ('),E.v(22,r-4,c)),': '),X.LEFT(E.v(22,r-4,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-4,c+3)),1.0), lambda: X.LEN(E.v(22,r-4,c+3))))),(X.cat(', ',E.v(22,r-4,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-4,c+27),1.0), X.eq(E.v(22,r-4,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-4,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-3,c-1),' ('),E.v(22,r-3,c)),': '),X.LEFT(E.v(22,r-3,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-3,c+3)),1.0), lambda: X.LEN(E.v(22,r-3,c+3))))),(X.cat(', ',E.v(22,r-3,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-3,c+27),1.0), X.eq(E.v(22,r-3,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-3,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-2,c-1),' ('),E.v(22,r-2,c)),': '),X.LEFT(E.v(22,r-2,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-2,c+3)),1.0), lambda: X.LEN(E.v(22,r-2,c+3))))),(X.cat(', ',E.v(22,r-2,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-2,c+27),1.0), X.eq(E.v(22,r-2,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-2,c+29),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r-1,c-1),' ('),E.v(22,r-1,c)),': '),X.LEFT(E.v(22,r-1,c+3), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-1,c+3)),1.0), lambda: X.LEN(E.v(22,r-1,c+3))))),(X.cat(', ',E.v(22,r-1,c+4)) if X.truthy(X.OR(X.eq(E.v(22,r-1,c+27),1.0), X.eq(E.v(22,r-1,c+28),1.0))) else '')),') ') if X.truthy(X.ge(E.v(22,r-1,c+29),2.0)) else ''))))),X.CHAR(10.0)),'Năm chỉ cần chú ý nhẹ (1 loại): '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-6,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-6,c+29),1.0), X.eq(E.v(22,r-6,c+27),0.0), X.eq(E.v(22,r-6,c+28),0.0))) else ''),(X.cat(E.v(22,r-5,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-5,c+29),1.0), X.eq(E.v(22,r-5,c+27),0.0), X.eq(E.v(22,r-5,c+28),0.0))) else '')),(X.cat(E.v(22,r-4,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-4,c+29),1.0), X.eq(E.v(22,r-4,c+27),0.0), X.eq(E.v(22,r-4,c+28),0.0))) else '')),(X.cat(E.v(22,r-3,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-3,c+29),1.0), X.eq(E.v(22,r-3,c+27),0.0), X.eq(E.v(22,r-3,c+28),0.0))) else '')),(X.cat(E.v(22,r-2,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-2,c+29),1.0), X.eq(E.v(22,r-2,c+27),0.0), X.eq(E.v(22,r-2,c+28),0.0))) else '')),(X.cat(E.v(22,r-1,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-1,c+29),1.0), X.eq(E.v(22,r-1,c+27),0.0), X.eq(E.v(22,r-1,c+28),0.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-6,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-6,c+29),1.0), X.eq(E.v(22,r-6,c+27),0.0), X.eq(E.v(22,r-6,c+28),0.0))) else ''),(X.cat(E.v(22,r-5,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-5,c+29),1.0), X.eq(E.v(22,r-5,c+27),0.0), X.eq(E.v(22,r-5,c+28),0.0))) else '')),(X.cat(E.v(22,r-4,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-4,c+29),1.0), X.eq(E.v(22,r-4,c+27),0.0), X.eq(E.v(22,r-4,c+28),0.0))) else '')),(X.cat(E.v(22,r-3,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-3,c+29),1.0), X.eq(E.v(22,r-3,c+27),0.0), X.eq(E.v(22,r-3,c+28),0.0))) else '')),(X.cat(E.v(22,r-2,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-2,c+29),1.0), X.eq(E.v(22,r-2,c+27),0.0), X.eq(E.v(22,r-2,c+28),0.0))) else '')),(X.cat(E.v(22,r-1,c-1),' ') if X.truthy(X.AND(X.eq(E.v(22,r-1,c+29),1.0), X.eq(E.v(22,r-1,c+27),0.0), X.eq(E.v(22,r-1,c+28),0.0))) else '')))))

def f96(E, r, c):
    # Ô C42
    # = "Nguyên tắc: ""tránh phương kỵ thần, đi theo phương hỷ, dụng thần"". "&"• Hoạt động nhiều ở phương dụng thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2
    #   ,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),5)+1)&" ("&CHOOSE((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109))+1,"Đông","Nam","Trung cung (Thìn–T
    #   uất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc")&")"&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)>=0,", hỷ thần "&INDE
    #   X(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),5)+1)&" ("&CHOOSE((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))
    #   +1,"Đông","Nam","Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc")&")","")&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,
    #   DUNG_HY_KY!$C$111)>=0,"; tránh phương kỵ thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),5)+1)&" ("&CHOOSE((IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))+1,"Đông","Nam","Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc"
    #   )&")","")&" (P2 tr.631-632). "&"Tâm phương vị: cả đời lấy nơi sinh; năm gặp hạn lấy nơi đang sống (bán kính 20–30 km là chưa đi xa); trong nhà lấy tâm
    #    nhà (giao điểm hai đường chéo nền nhà). Sách: nhiều khi đổi phòng, đổi chỗ giường hiệu quả hơn đi xa."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nguyên tắc: "tránh phương kỵ thần, đi theo phương hỷ, dụng thần". ','• Hoạt động nhiều ở phương dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),')'),(X.cat(X.cat(X.cat(X.cat(', hỷ thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),')') if X.truthy(X.ge((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else '')),(X.cat(X.cat(X.cat(X.cat('; tránh phương kỵ thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),')') if X.truthy(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0)) else '')),' (P2 tr.631-632). '),'Tâm phương vị: cả đời lấy nơi sinh; năm gặp hạn lấy nơi đang sống (bán kính 20–30 km là chưa đi xa); trong nhà lấy tâm nhà (giao điểm hai đường chéo nền nhà). Sách: nhiều khi đổi phòng, đổi chỗ giường hiệu quả hơn đi xa.')

def f97(E, r, c):
    # Ô M42
    # = MOD(-MOD(DAI_VAN!$C$23-1864,9),9)+1
    return X.add(X.MOD(X.neg(X.MOD(X.sub(E.v(14,23,3),1864.0), 9.0)), 9.0),1.0)

def f98(E, r, c):
    # Ô N42
    # = IF(MOD((MOD($M$43-0-1,9)+1)-$M$42,3)=0,"tháng 1 ("&INDEX(CHI_TEN,MOD(2,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-0-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-0-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 1 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(2.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),0.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),0.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),0.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f99(E, r, c):
    # Ô O42
    # = TRIM(IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD(DAI_VAN!$C$23-1984,10)/2)-TRUNG_GIAN!$E$7,5)=0),"Mục 1 (năm tỉ kiếp,
    #    nhật chủ vượng): khí hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7,5)+1)&" ("&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị 
    #   kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&") t
    #   hái quá – dễ phát nhiệt, bực tức, tranh chấp, phá tài; nguyên tắc: tiết bớt hành nhật chủ, dưỡng hành bị nó khắc ("&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7
    #   +2,5)+1)&" – "&CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày,
    #    tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&"). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY
    #   _KY!$C$7)=0),MOD(INT(MOD(DAI_VAN!$C$23-1984,10)/2)-TRUNG_GIAN!$E$7,5)=3),"Mục 2 (năm quan sát, nhật chủ nhược): khí "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,
    #   "gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng
    #    quang, các tạng vùng bụng, chi dưới")&" bất túc, sức đề kháng yếu, ""nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng""; nguyên tắc: bồi bổ hành
    #    nhật chủ "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7,5)+1)&", dùng hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)&" (ấn – hóa sát) hoặc "&INDEX(HANH_TEN,
    #   MOD(TRUNG_GIAN!$E$7+1,5)+1)&" (thực thương – chế sát). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD(DAI_VAN!$C$2
    #   3-1984,10)/2)-TRUNG_GIAN!$E$7,5)=4),"Mục 3 (năm ấn, nhật chủ vượng): hành ấn "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)&" và hành nhật chủ thái quá;
    #    nguyên tắc: giảm hành ấn, làm yếu hành nhật chủ, dưỡng hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1)&" ("&CHOOSE(MOD(TRUNG_GIAN!$E$7+1,5)+1,"gan 
    #   – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quan
    #   g, các tạng vùng bụng, chi dưới")&"). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD(DAI_VAN!$C$23-1984,10)/2)-TRU
    #   NG_GIAN!$E$7,5)=1),"Mục 4 (năm thực thương, nhật chủ nhược): khí bị tiết, ""không ốm đau cũng gặp rủi ro""; nguyên tắc (giả định): bổ trợ hành nhật ch
    #   ủ "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7,5)+1)&" ("&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – m
    #   áu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&"). ","")&IF(AND((IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD(DAI_VAN!$C$23-1984,10)/2)-TRUNG_GIAN!$E$7,5)=2),"Mục 5 (năm tài, nhật chủ nhược): khí "&CHOO
    #   SE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (ph
    #   ế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" bất túc; nguyên tắc: bổ hành nhật chủ "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7,5)+1)&
    #   ", chú ý dưỡng "&CHOOSE(MOD(TRUNG_GIAN!$E$7+1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dà
    #   y, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&". ","")&IF(N(IFERROR(INDEX(LUU_NIEN!$BK$30:$BK$129,MATCH(DA
    #   I_VAN!$C$23,LUU_NIEN!$A$30:$A$129,0)),""))=1,"Mục 6 (thương quan gặp quan): dùng hành thông quan "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&" (tài) 
    #   để điều hòa; nhật chủ nhược thì kiêm bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt"
    #   ,"tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&". ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 1 (năm tỉ kiếp, nhật chủ vượng): khí hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),' ('),X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),') thái quá – dễ phát nhiệt, bực tức, tranh chấp, phá tài; nguyên tắc: tiết bớt hành nhật chủ, dưỡng hành bị nó khắc ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),' – '),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.sub(X.INT(X.div(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),2.0)),E.v(23,7,5)), 5.0),0.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 2 (năm quan sát, nhật chủ nhược): khí ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' bất túc, sức đề kháng yếu, "nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng"; nguyên tắc: bồi bổ hành nhật chủ '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),', dùng hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),' (ấn – hóa sát) hoặc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))),' (thực thương – chế sát). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),2.0)),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 3 (năm ấn, nhật chủ vượng): hành ấn ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),' và hành nhật chủ thái quá; nguyên tắc: giảm hành ấn, làm yếu hành nhật chủ, dưỡng hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))),' ('),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.sub(X.INT(X.div(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),2.0)),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.cat('Mục 4 (năm thực thương, nhật chủ nhược): khí bị tiết, "không ốm đau cũng gặp rủi ro"; nguyên tắc (giả định): bổ trợ hành nhật chủ ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),' ('),X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),2.0)),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 5 (năm tài, nhật chủ nhược): khí ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' bất túc; nguyên tắc: bổ hành nhật chủ '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),', chú ý dưỡng '),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'. ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.sub(X.INT(X.div(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),2.0)),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.cat('Mục 6 (thương quan gặp quan): dùng hành thông quan ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),' (tài) để điều hòa; nhật chủ nhược thì kiêm bổ '),X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'. ') if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,63,129,63), X.MATCH(E.v(14,23,3), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')))

def f100(E, r, c):
    # Ô C43
    # = "Năm "&DAI_VAN!$C$23&" ("&INDEX(CAN_TEN,MOD(DAI_VAN!$C$23-1984,10)+1)&" "&INDEX(CHI_TEN,MOD(DAI_VAN!$C$23-1984,12)+1)&"): phương thái tuế "&INDEX(CHI_
    #   TEN,MOD(DAI_VAN!$C$23-1984,12)+1)&" ("&INDEX({"Bắc","Đông bắc","Đông bắc","Đông","Đông nam","Đông nam","Nam","Tây nam","Tây nam","Tây","Tây bắc","Tây 
    #   bắc"},MOD(DAI_VAN!$C$23-1984,12)+1)&"), phương tuế phá "&INDEX(CHI_TEN,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12)+1)&" ("&INDEX({"Bắc","Đông bắc","Đông bắc"
    #   ,"Đông","Đông nam","Đông nam","Nam","Tây nam","Tây nam","Tây","Tây bắc","Tây bắc"},MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12)+1)&"). "&IF(COUNTIF(TRUNG_GIAN
    #   !$C$6:$F$6,MOD(DAI_VAN!$C$23-1984,12))>0,"Tứ trụ có chi "&INDEX(CHI_TEN,MOD(DAI_VAN!$C$23-1984,12)+1)&" → phục ngâm thái tuế: giường nằm không đặt ở p
    #   hương thái tuế và tuế phá, KỴ NHẤT phương thái tuế "&INDEX(CHI_TEN,MOD(DAI_VAN!$C$23-1984,12)+1)&".",IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,MOD(MOD(DAI_VAN!$
    #   C$23-1984,12)+6,12))>0,"Tứ trụ có chi "&INDEX(CHI_TEN,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12)+1)&" xung thái tuế: giường nằm không đặt ở phương thái tuế 
    #   và tuế phá, KỴ NHẤT phương tuế phá "&INDEX(CHI_TEN,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12)+1)&".","Tứ trụ không phục ngâm, không xung thái tuế năm này – 
    #   chỉ cần tránh hoạt động nhiều ở hai phương trên khi năm có hạn."))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',E.v(14,23,3)),' ('),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),'): phương thái tuế '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),' ('),X.INDEX(X.Arr([['Bắc', 'Đông bắc', 'Đông bắc', 'Đông', 'Đông nam', 'Đông nam', 'Nam', 'Tây nam', 'Tây nam', 'Tây', 'Tây bắc', 'Tây bắc']]), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),'), phương tuế phá '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0),1.0))),' ('),X.INDEX(X.Arr([['Bắc', 'Đông bắc', 'Đông bắc', 'Đông', 'Đông nam', 'Đông nam', 'Nam', 'Tây nam', 'Tây nam', 'Tây', 'Tây bắc', 'Tây bắc']]), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0),1.0))),'). '),(X.cat(X.cat(X.cat(X.cat('Tứ trụ có chi ',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),' → phục ngâm thái tuế: giường nằm không đặt ở phương thái tuế và tuế phá, KỴ NHẤT phương thái tuế '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),'.') if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0)),0.0)) else (X.cat(X.cat(X.cat(X.cat('Tứ trụ có chi ',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0),1.0))),' xung thái tuế: giường nằm không đặt ở phương thái tuế và tuế phá, KỴ NHẤT phương tuế phá '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0),1.0))),'.') if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0)),0.0)) else 'Tứ trụ không phục ngâm, không xung thái tuế năm này – chỉ cần tránh hoạt động nhiều ở hai phương trên khi năm có hạn.')))

def f101(E, r, c):
    # Ô M43
    # = CHOOSE(MOD(MOD(DAI_VAN!$C$23-1984,12),3)+1,8,5,2)
    return X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0), 3.0),1.0), lambda: 8.0, lambda: 5.0, lambda: 2.0)

def f102(E, r, c):
    # Ô N43
    # = IF(MOD((MOD($M$43-1-1,9)+1)-$M$42,3)=0,"tháng 2 ("&INDEX(CHI_TEN,MOD(3,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-1-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-1-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 2 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(3.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),1.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),1.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),1.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f103(E, r, c):
    # Ô O43
    # = TRIM(IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),IF(DAI_VAN!$C$25>=1,MOD(IF(DAI_VAN!$C$25>=1,INDEX(DAI_VAN!$C$42:$L$42,MIN(10,DAI
    #   _VAN!$C$25)),-1)-TRUNG_GIAN!$E$7,5),-1)=0),"Mục 1 (vận tỉ kiếp, nhật chủ vượng): khí hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7,5)+1)&" ("&CHOOSE(MOD(T
    #   RUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại
    #    tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&") thái quá – dễ phát nhiệt, bực tức, tranh chấp, phá tài; nguyên tắc: tiết bớt hành nhật c
    #   hủ, dưỡng hành bị nó khắc ("&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&" – "&CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,"gan – mật (vùng thắt lưng, gân cốt kh
    #   i bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")
    #   &"). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),IF(DAI_VAN!$C$25>=1,MOD(IF(DAI_VAN!$C$25>=1,INDEX(DAI_VAN!$C$42:$L$42,MIN(
    #   10,DAI_VAN!$C$25)),-1)-TRUNG_GIAN!$E$7,5),-1)=3),"Mục 2 (vận quan sát, nhật chủ nhược): khí "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lư
    #   ng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụ
    #   ng, chi dưới")&" bất túc, sức đề kháng yếu, ""nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng""; nguyên tắc: bồi bổ hành nhật chủ "&INDEX(HANH_T
    #   EN,MOD(TRUNG_GIAN!$E$7,5)+1)&", dùng hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)&" (ấn – hóa sát) hoặc "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)
    #   +1)&" (thực thương – chế sát). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),IF(DAI_VAN!$C$25>=1,MOD(IF(DAI_VAN!$C$25>=1,INDE
    #   X(DAI_VAN!$C$42:$L$42,MIN(10,DAI_VAN!$C$25)),-1)-TRUNG_GIAN!$E$7,5),-1)=4),"Mục 3 (vận ấn, nhật chủ vượng): hành ấn "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E
    #   $7+4,5)+1)&" và hành nhật chủ thái quá; nguyên tắc: giảm hành ấn, làm yếu hành nhật chủ, dưỡng hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+1,5)+1)&" ("&
    #   CHOOSE(MOD(TRUNG_GIAN!$E$7+1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","ph
    #   ổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&"). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),IF
    #   (DAI_VAN!$C$25>=1,MOD(IF(DAI_VAN!$C$25>=1,INDEX(DAI_VAN!$C$42:$L$42,MIN(10,DAI_VAN!$C$25)),-1)-TRUNG_GIAN!$E$7,5),-1)=1),"Mục 4 (vận thực thương, nhật
    #    chủ nhược): khí bị tiết, ""không ốm đau cũng gặp rủi ro""; nguyên tắc (giả định): bổ trợ hành nhật chủ "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7,5)+1)&" (
    #   "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","ph
    #   ổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&"). ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),IF
    #   (DAI_VAN!$C$25>=1,MOD(IF(DAI_VAN!$C$25>=1,INDEX(DAI_VAN!$C$42:$L$42,MIN(10,DAI_VAN!$C$25)),-1)-TRUNG_GIAN!$E$7,5),-1)=2),"Mục 5 (vận tài, nhật chủ như
    #   ợc): khí "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu 
    #   hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&" bất túc; nguyên tắc: bổ hành nhật chủ "&INDEX(HANH_TEN,MOD(TRUNG_G
    #   IAN!$E$7,5)+1)&", chú ý dưỡng "&CHOOSE(MOD(TRUNG_GIAN!$E$7+1,5)+1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị 
    #   (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – bàng quang, các tạng vùng bụng, chi dưới")&". ","")&IF(N(0)=1,"Mục 6 (thương quan gặp qu
    #   an): dùng hành thông quan "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&" (tài) để điều hòa; nhật chủ nhược thì kiêm bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)
    #   +1,"gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)","tim (tâm) – máu, mắt","tì vị (lá lách, dạ dày, tiêu hóa)","phổi (phế) – đại tràng","thận – b
    #   àng quang, các tạng vùng bụng, chi dưới")&". ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 1 (vận tỉ kiếp, nhật chủ vượng): khí hành ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),' ('),X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),') thái quá – dễ phát nhiệt, bực tức, tranh chấp, phá tài; nguyên tắc: tiết bớt hành nhật chủ, dưỡng hành bị nó khắc ('),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),' – '),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq((X.MOD(X.sub((X.INDEX(E.rg(14,42,3,42,12), X.MIN(10.0, E.rg(14,25,3,25,3))) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),E.v(23,7,5)), 5.0) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),0.0))) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 2 (vận quan sát, nhật chủ nhược): khí ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' bất túc, sức đề kháng yếu, "nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng"; nguyên tắc: bồi bổ hành nhật chủ '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),', dùng hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),' (ấn – hóa sát) hoặc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))),' (thực thương – chế sát). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq((X.MOD(X.sub((X.INDEX(E.rg(14,42,3,42,12), X.MIN(10.0, E.rg(14,25,3,25,3))) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),E.v(23,7,5)), 5.0) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),3.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 3 (vận ấn, nhật chủ vượng): hành ấn ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),' và hành nhật chủ thái quá; nguyên tắc: giảm hành ấn, làm yếu hành nhật chủ, dưỡng hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0))),' ('),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq((X.MOD(X.sub((X.INDEX(E.rg(14,42,3,42,12), X.MIN(10.0, E.rg(14,25,3,25,3))) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),E.v(23,7,5)), 5.0) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),4.0))) else '')),(X.cat(X.cat(X.cat(X.cat('Mục 4 (vận thực thương, nhật chủ nhược): khí bị tiết, "không ốm đau cũng gặp rủi ro"; nguyên tắc (giả định): bổ trợ hành nhật chủ ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),' ('),X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'). ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq((X.MOD(X.sub((X.INDEX(E.rg(14,42,3,42,12), X.MIN(10.0, E.rg(14,25,3,25,3))) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),E.v(23,7,5)), 5.0) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),1.0))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mục 5 (vận tài, nhật chủ nhược): khí ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),' bất túc; nguyên tắc: bổ hành nhật chủ '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(E.v(23,7,5), 5.0),1.0))),', chú ý dưỡng '),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'. ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq((X.MOD(X.sub((X.INDEX(E.rg(14,42,3,42,12), X.MIN(10.0, E.rg(14,25,3,25,3))) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),E.v(23,7,5)), 5.0) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),2.0))) else '')),(X.cat(X.cat(X.cat(X.cat('Mục 6 (thương quan gặp quan): dùng hành thông quan ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),' (tài) để điều hòa; nhật chủ nhược thì kiêm bổ '),X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật (vùng thắt lưng, gân cốt khi bị kim khắc)', lambda: 'tim (tâm) – máu, mắt', lambda: 'tì vị (lá lách, dạ dày, tiêu hóa)', lambda: 'phổi (phế) – đại tràng', lambda: 'thận – bàng quang, các tạng vùng bụng, chi dưới')),'. ') if X.truthy(X.eq(X.N(0.0),1.0)) else '')))

def f104(E, r, c):
    # Ô C44
    # = "Tam sát năm "&INDEX(CHI_TEN,MOD(DAI_VAN!$C$23-1984,12)+1)&" (tuyệt – thai – dưỡng của tam hợp cục): "&INDEX(CHI_TEN,MOD(CHOOSE(MOD(MOD(DAI_VAN!$C$23-
    #   1984,12),4)+1,8,5,2,11)+9,12)+1)&", "&INDEX(CHI_TEN,MOD(CHOOSE(MOD(MOD(DAI_VAN!$C$23-1984,12),4)+1,8,5,2,11)+10,12)+1)&", "&INDEX(CHI_TEN,MOD(CHOOSE(M
    #   OD(MOD(DAI_VAN!$C$23-1984,12),4)+1,8,5,2,11)+11,12)+1)&" – phương "&INDEX({"Bắc","Đông bắc","Đông bắc","Đông","Đông nam","Đông nam","Nam","Tây nam","T
    #   ây nam","Tây","Tây bắc","Tây bắc"},MOD(CHOOSE(MOD(MOD(DAI_VAN!$C$23-1984,12),4)+1,8,5,2,11)+10,12)+1)&"; kiêng đặt giường ở phương này."&IF(MOD(MOD(DA
    #   I_VAN!$C$23-1984,12),4)=0," (Sách in dòng năm Thân Tí Thìn → Dần Mão Thìn, trùng với dòng Tị Dậu Sửu; theo định nghĩa tuyệt–thai–dưỡng của Thủy cục là
    #    Tị Ngọ Mùi – công cụ dùng định nghĩa, xem MAU_THUAN M22.)","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tam sát năm ',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),' (tuyệt – thai – dưỡng của tam hợp cục): '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0), 4.0),1.0), lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0),9.0), 12.0),1.0))),', '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0), 4.0),1.0), lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0),10.0), 12.0),1.0))),', '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0), 4.0),1.0), lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0),11.0), 12.0),1.0))),' – phương '),X.INDEX(X.Arr([['Bắc', 'Đông bắc', 'Đông bắc', 'Đông', 'Đông nam', 'Đông nam', 'Nam', 'Tây nam', 'Tây nam', 'Tây', 'Tây bắc', 'Tây bắc']]), X.add(X.MOD(X.add(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0), 4.0),1.0), lambda: 8.0, lambda: 5.0, lambda: 2.0, lambda: 11.0),10.0), 12.0),1.0))),'; kiêng đặt giường ở phương này.'),(' (Sách in dòng năm Thân Tí Thìn → Dần Mão Thìn, trùng với dòng Tị Dậu Sửu; theo định nghĩa tuyệt–thai–dưỡng của Thủy cục là Tị Ngọ Mùi – công cụ dùng định nghĩa, xem MAU_THUAN M22.)' if X.truthy(X.eq(X.MOD(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0), 4.0),0.0)) else ''))

def f105(E, r, c):
    # Ô N44
    # = IF(MOD((MOD($M$43-2-1,9)+1)-$M$42,3)=0,"tháng 3 ("&INDEX(CHI_TEN,MOD(4,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-2-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-2-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 3 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(4.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),2.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),2.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),2.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f106(E, r, c):
    # Ô C45
    # = "Phương tuế hợp can: "&INDEX(CAN_TEN,MOD(MOD(DAI_VAN!$C$23-1984,10)+5,10)+1)&" ("&CHOOSE(INT(MOD(MOD(MOD(DAI_VAN!$C$23-1984,10)+5,10),10)/2)+1,"Đông",
    #   "Nam","trung tâm – Mậu Kỷ thuần thổ, hướng không rõ","Tây","Bắc")&"); phương lục hợp: "&INDEX(CHI_TEN,MOD(13-MOD(DAI_VAN!$C$23-1984,12),12)+1)&" ("&IN
    #   DEX({"Bắc","Đông bắc","Đông bắc","Đông","Đông nam","Đông nam","Nam","Tây nam","Tây nam","Tây","Tây bắc","Tây bắc"},MOD(13-MOD(DAI_VAN!$C$23-1984,12),1
    #   2)+1)&"). Sách: ""hợp thái tuế chủ về điều mừng""; năm gặp hạn (xung, hình thái tuế) đặt giường ở phương hợp để ""tham hợp vong xung, hóa hung thành c
    #   át""."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Phương tuế hợp can: ',X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),5.0), 10.0),1.0))),' ('),X.CHOOSE(X.add(X.INT(X.div(X.MOD(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),5.0), 10.0), 10.0),2.0)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'trung tâm – Mậu Kỷ thuần thổ, hướng không rõ', lambda: 'Tây', lambda: 'Bắc')),'); phương lục hợp: '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(13.0,X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0)), 12.0),1.0))),' ('),X.INDEX(X.Arr([['Bắc', 'Đông bắc', 'Đông bắc', 'Đông', 'Đông nam', 'Đông nam', 'Nam', 'Tây nam', 'Tây nam', 'Tây', 'Tây bắc', 'Tây bắc']]), X.add(X.MOD(X.sub(13.0,X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0)), 12.0),1.0))),'). Sách: "hợp thái tuế chủ về điều mừng"; năm gặp hạn (xung, hình thái tuế) đặt giường ở phương hợp để "tham hợp vong xung, hóa hung thành cát".')

def f107(E, r, c):
    # Ô N45
    # = IF(MOD((MOD($M$43-3-1,9)+1)-$M$42,3)=0,"tháng 4 ("&INDEX(CHI_TEN,MOD(5,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-3-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-3-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 4 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(5.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),3.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),3.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),3.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f108(E, r, c):
    # Ô C46
    # = "Năm "&DAI_VAN!$C$23&": "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},$M$42)&" nhập cung 
    #   Giữa; NGŨ HOÀNG ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)",
    #   "Tốn (đông nam)"},MOD(5-($M$42),9)+1)&", NHỊ HẮC ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)
    #   ","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-($M$42),9)+1)&" (sao tai ách, bệnh tật – phòng ngủ, giường, cửa chính kỵ các phương này); LỤC
    #    BẠCH ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông
    #    nam)"},MOD(6-($M$42),9)+1)&" (nên sinh hoạt nhiều ở phương này). Sao năm đổi tại Lập xuân."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',E.v(14,23,3)),': '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), E.v(22,42,13))),' nhập cung Giữa; NGŨ HOÀNG ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,E.v(22,42,13)), 9.0),1.0))),', NHỊ HẮC ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,E.v(22,42,13)), 9.0),1.0))),' (sao tai ách, bệnh tật – phòng ngủ, giường, cửa chính kỵ các phương này); LỤC BẠCH ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(6.0,E.v(22,42,13)), 9.0),1.0))),' (nên sinh hoạt nhiều ở phương này). Sao năm đổi tại Lập xuân.')

def f109(E, r, c):
    # Ô N46
    # = IF(MOD((MOD($M$43-4-1,9)+1)-$M$42,3)=0,"tháng 5 ("&INDEX(CHI_TEN,MOD(6,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-4-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-4-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 5 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(6.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),4.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),4.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),4.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f110(E, r, c):
    # Ô C47
    # = IF(TRIM($N$42&$N$43&$N$44&$N$45&$N$46&$N$47&$N$48&$N$49&$N$50&$N$51&$N$52&$N$53)="","Không có tháng nào Ngũ hoàng/Nhị hắc của tháng trùng cung với năm
    #   .","Tháng có Ngũ hoàng/Nhị hắc của tháng đến đúng cung Ngũ hoàng/Nhị hắc của năm (""hai sao cùng đến tất sẽ nặng hơn""): "&TRIM($N$42&$N$43&$N$44&$N$4
    #   5&$N$46&$N$47&$N$48&$N$49&$N$50&$N$51&$N$52&$N$53)&" Tháng tính theo tiết (tháng 1 bắt đầu từ Lập xuân). Cách giải sách nêu: tạm chuyển phòng hoặc đón
    #   g cửa phương đó, đeo trang sức vàng bạc, sinh hoạt ở phương Lục bạch.")
    return ('Không có tháng nào Ngũ hoàng/Nhị hắc của tháng trùng cung với năm.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,42,14),E.v(22,43,14)),E.v(22,44,14)),E.v(22,45,14)),E.v(22,46,14)),E.v(22,47,14)),E.v(22,48,14)),E.v(22,49,14)),E.v(22,50,14)),E.v(22,51,14)),E.v(22,52,14)),E.v(22,53,14))),'')) else X.cat(X.cat('Tháng có Ngũ hoàng/Nhị hắc của tháng đến đúng cung Ngũ hoàng/Nhị hắc của năm ("hai sao cùng đến tất sẽ nặng hơn"): ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,42,14),E.v(22,43,14)),E.v(22,44,14)),E.v(22,45,14)),E.v(22,46,14)),E.v(22,47,14)),E.v(22,48,14)),E.v(22,49,14)),E.v(22,50,14)),E.v(22,51,14)),E.v(22,52,14)),E.v(22,53,14)))),' Tháng tính theo tiết (tháng 1 bắt đầu từ Lập xuân). Cách giải sách nêu: tạm chuyển phòng hoặc đóng cửa phương đó, đeo trang sức vàng bạc, sinh hoạt ở phương Lục bạch.'))

def f111(E, r, c):
    # Ô N47
    # = IF(MOD((MOD($M$43-5-1,9)+1)-$M$42,3)=0,"tháng 6 ("&INDEX(CHI_TEN,MOD(7,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-5-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-5-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 6 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(7.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),5.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),5.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),5.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f112(E, r, c):
    # Ô C48
    # = "Năm xem: "&G32&". Sách: phản ngâm là tuế/vận thiên khắc địa xung trụ ngày, phục ngâm là tuế/vận trùng can chi trụ ngày; ""dù là phản ngâm hay phục ng
    #   âm, về nguyên tắc nên di chuyển … di chuyển lớn là đi xa, ít ra thì dời chỗ ở trong nhà, cố tránh xa thái t"&"uế và trường lực đối xung. Đi đến những 
    #   phương có lợi""."
    return X.cat(X.cat(X.cat('Năm xem: ',E.v(22,r-16,c+4)),'. Sách: phản ngâm là tuế/vận thiên khắc địa xung trụ ngày, phục ngâm là tuế/vận trùng can chi trụ ngày; "dù là phản ngâm hay phục ngâm, về nguyên tắc nên di chuyển … di chuyển lớn là đi xa, ít ra thì dời chỗ ở trong nhà, cố tránh xa thái t'),'uế và trường lực đối xung. Đi đến những phương có lợi".')

def f113(E, r, c):
    # Ô N48
    # = IF(MOD((MOD($M$43-6-1,9)+1)-$M$42,3)=0,"tháng 7 ("&INDEX(CHI_TEN,MOD(8,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-6-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-6-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 7 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(8.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),6.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),6.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),6.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f114(E, r, c):
    # Ô C49
    # = "Chọn ngày tốt từ tháng 11–12 của năm trước năm hạn (năm "&(DAI_VAN!$C$23-1)&" cho năm "&DAI_VAN!$C$23&"); gấp thì chọn ngày tốt ngay trong tháng. Ưu 
    #   tiên ngày/giờ thiên đức, nguyệt đức, hợp can năm ("&INDEX(CAN_TEN,MOD(MOD(DAI_VAN!$C$23-1984,10)+5,10)+1)&"), hợp chi năm ("&INDEX(CHI_TEN,MOD(13-MOD(
    #   DAI_VAN!$C$23-1984,12),12)+1)&"), thiên ất quý nhân; không xung phạm bản mệnh. "&IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12))
    #   >0,"Năm "&INDEX(CHI_TEN,MOD(DAI_VAN!$C$23-1984,12)+1)&" xung chi "&INDEX(CHI_TEN,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12)+1)&" của tứ trụ → có thể chọn ng
    #   ày/giờ "&INDEX(CHI_TEN,MOD(13-MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12),12)+1)&" hoặc "&INDEX(CHI_TEN,MOD(13-MOD(DAI_VAN!$C$23-1984,12),12)+1)&" (lục hợp),
    #    dâng hương/phóng sinh ở đúng phương đó. ","")&"Quan – thương xung khắc: chọn ngày giờ thuộc tài (hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+2,5)+1)&")
    #   ; quan sát khắc thân: chọn cục ấn (hành "&INDEX(HANH_TEN,MOD(TRUNG_GIAN!$E$7+4,5)+1)&"). Giờ nên là giờ quý nhân, tạo thành dụng/hỷ thần của mệnh."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chọn ngày tốt từ tháng 11–12 của năm trước năm hạn (năm ',X.sub(E.v(14,23,3),1.0)),' cho năm '),E.v(14,23,3)),'); gấp thì chọn ngày tốt ngay trong tháng. Ưu tiên ngày/giờ thiên đức, nguyệt đức, hợp can năm ('),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 10.0),5.0), 10.0),1.0))),'), hợp chi năm ('),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(13.0,X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0)), 12.0),1.0))),'), thiên ất quý nhân; không xung phạm bản mệnh. '),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0))),' xung chi '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0),1.0))),' của tứ trụ → có thể chọn ngày/giờ '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(13.0,X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0)), 12.0),1.0))),' hoặc '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(13.0,X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0)), 12.0),1.0))),' (lục hợp), dâng hương/phóng sinh ở đúng phương đó. ') if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0)),0.0)) else '')),'Quan – thương xung khắc: chọn ngày giờ thuộc tài (hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0))),'); quan sát khắc thân: chọn cục ấn (hành '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),1.0))),'). Giờ nên là giờ quý nhân, tạo thành dụng/hỷ thần của mệnh.')

def f115(E, r, c):
    # Ô N49
    # = IF(MOD((MOD($M$43-7-1,9)+1)-$M$42,3)=0,"tháng 8 ("&INDEX(CHI_TEN,MOD(9,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây
    #   )","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-7-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX({
    #   "cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((M
    #   OD($M$43-7-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 8 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(9.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),7.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),7.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),7.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f116(E, r, c):
    # Ô C50
    # = "Phóng sinh theo phương hỷ thần: "&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)>=0,INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!
    #   $D$110,DUNG_HY_KY!$C$110),5)+1)&" – "&CHOOSE((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))+1,"Đông","Nam","Trung cung (Thìn–Tuất–Sửu–Mùi
    #   : đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc"),"(không có hỷ thần riêng – dùng phương dụng thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_
    #   HY_KY!$D$109,DUNG_HY_KY!$C$109),5)+1)&" – "&CHOOSE((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109))+1,"Đông","Nam","Trung cung (Thìn–Tuất–S
    #   ửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc")&", giả định)")&". Sách: ""phóng sinh vật nuôi theo phương hỉ thần của chủ … nếu phương vị 
    #   phóng sinh sai có thể sẽ đưa lại rủi ro""; tốt nhất là rùa hoặc ba ba; chọn ngày tốt, đốt hương rồi phóng sinh."
    return X.cat(X.cat('Phóng sinh theo phương hỷ thần: ',(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)), 5.0),1.0)),' – '),X.CHOOSE(X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')) if X.truthy(X.ge((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else X.cat(X.cat(X.cat(X.cat('(không có hỷ thần riêng – dùng phương dụng thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)), 5.0),1.0))),' – '),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),', giả định)'))),'. Sách: "phóng sinh vật nuôi theo phương hỉ thần của chủ … nếu phương vị phóng sinh sai có thể sẽ đưa lại rủi ro"; tốt nhất là rùa hoặc ba ba; chọn ngày tốt, đốt hương rồi phóng sinh.')

def f117(E, r, c):
    # Ô N50
    # = IF(MOD((MOD($M$43-8-1,9)+1)-$M$42,3)=0,"tháng 9 ("&INDEX(CHI_TEN,MOD(10,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tâ
    #   y)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-8-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX(
    #   {"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-((
    #   MOD($M$43-8-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 9 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(10.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),8.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),8.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),8.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f118(E, r, c):
    # Ô C51
    # = IF(TRIM($O$42&IF(AND(OR(MOD(DAI_VAN!$C$23-1984,12)=1,MOD(DAI_VAN!$C$23-1984,12)=4,MOD(DAI_VAN!$C$23-1984,12)=7,MOD(DAI_VAN!$C$23-1984,12)=10),COUNTIF(
    #   TRUNG_GIAN!$C$6:$F$6,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12))>0),"Mục 7 (năm Thìn/Tuất/Sửu/Mùi xung chi tứ trụ): ""dễ bị bệnh ở tì, vị và thận, dùng thuỷ
    #    tư nhuận hoặc dùng mộc để giảm nhẹ"". ",""))="","Năm xem không thuộc 7 trường hợp đề cương phối thang của sách.","NĂM "&DAI_VAN!$C$23&": "&TRIM($O$42
    #   &IF(AND(OR(MOD(DAI_VAN!$C$23-1984,12)=1,MOD(DAI_VAN!$C$23-1984,12)=4,MOD(DAI_VAN!$C$23-1984,12)=7,MOD(DAI_VAN!$C$23-1984,12)=10),COUNTIF(TRUNG_GIAN!$C
    #   $6:$F$6,MOD(MOD(DAI_VAN!$C$23-1984,12)+6,12))>0),"Mục 7 (năm Thìn/Tuất/Sửu/Mùi xung chi tứ trụ): ""dễ bị bệnh ở tì, vị và thận, dùng thuỷ tư nhuận hoặ
    #   c dùng mộc để giảm nhẹ"". ","")))&IF(IF(DAI_VAN!$C$25>=1,MOD(IF(DAI_VAN!$C$25>=1,INDEX(DAI_VAN!$C$42:$L$42,MIN(10,DAI_VAN!$C$25)),-1)-TRUNG_GIAN!$E$7,
    #   5),-1)<0,"",CHAR(10)&"ĐẠI VẬN HIỆN TẠI: "&IF($O$43="","không thuộc trường hợp nào.",$O$43))&CHAR(10)&"⚠ Chỉ là nguyên tắc ngũ hành sách nêu để tham kh
    #   ảo; công cụ KHÔNG đưa tên thuốc – mọi việc dùng thuốc phải do thầy thuốc chẩn đoán và kê đơn."
    return X.cat(X.cat(X.cat(('Năm xem không thuộc 7 trường hợp đề cương phối thang của sách.' if X.truthy(X.eq(X.TRIM(X.cat(E.v(22,42,15),('Mục 7 (năm Thìn/Tuất/Sửu/Mùi xung chi tứ trụ): "dễ bị bệnh ở tì, vị và thận, dùng thuỷ tư nhuận hoặc dùng mộc để giảm nhẹ". ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0), X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),4.0), X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),7.0), X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),10.0)), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0)),0.0))) else ''))),'')) else X.cat(X.cat(X.cat('NĂM ',E.v(14,23,3)),': '),X.TRIM(X.cat(E.v(22,42,15),('Mục 7 (năm Thìn/Tuất/Sửu/Mùi xung chi tứ trụ): "dễ bị bệnh ở tì, vị và thận, dùng thuỷ tư nhuận hoặc dùng mộc để giảm nhẹ". ' if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),1.0), X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),4.0), X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),7.0), X.eq(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),10.0)), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.MOD(X.add(X.MOD(X.sub(E.v(14,23,3),1984.0), 12.0),6.0), 12.0)),0.0))) else ''))))),('' if X.truthy(X.lt((X.MOD(X.sub((X.INDEX(E.rg(14,42,3,42,12), X.MIN(10.0, E.rg(14,25,3,25,3))) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),E.v(23,7,5)), 5.0) if X.truthy(X.ge(E.v(14,25,3),1.0)) else X.neg(1.0)),0.0)) else X.cat(X.cat(X.CHAR(10.0),'ĐẠI VẬN HIỆN TẠI: '),('không thuộc trường hợp nào.' if X.truthy(X.eq(E.v(22,43,15),'')) else E.v(22,43,15))))),X.CHAR(10.0)),'⚠ Chỉ là nguyên tắc ngũ hành sách nêu để tham khảo; công cụ KHÔNG đưa tên thuốc – mọi việc dùng thuốc phải do thầy thuốc chẩn đoán và kê đơn.')

def f119(E, r, c):
    # Ô N51
    # = IF(MOD((MOD($M$43-9-1,9)+1)-$M$42,3)=0,"tháng 10 ("&INDEX(CHI_TEN,MOD(11,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (t
    #   ây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-9-1,9)+1)),9)+1)&", Nhị hắc ở "&INDEX
    #   ({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-(
    #   (MOD($M$43-9-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 10 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(11.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),9.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),9.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),9.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f120(E, r, c):
    # Ô C52
    # = "Sách: ""người hành thiện tích đức, trong lòng thư thái, tính tình cởi mở, khí huyết hài hoà nên mạnh khoẻ, ít bệnh … mọi việc dễ hanh thông"". Đây là
    #    cách hóa giải áp dụng cho mọi lá số."
    return 'Sách: "người hành thiện tích đức, trong lòng thư thái, tính tình cởi mở, khí huyết hài hoà nên mạnh khoẻ, ít bệnh … mọi việc dễ hanh thông". Đây là cách hóa giải áp dụng cho mọi lá số.'

def f121(E, r, c):
    # Ô N52
    # = IF(MOD((MOD($M$43-10-1,9)+1)-$M$42,3)=0,"tháng 11 ("&INDEX(CHI_TEN,MOD(12,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (
    #   tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-10-1,9)+1)),9)+1)&", Nhị hắc ở "&IND
    #   EX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2
    #   -((MOD($M$43-10-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 11 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(12.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),10.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),10.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),10.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f122(E, r, c):
    # Ô C53
    # = """Khi gặp những Tứ trụ không hoàn toàn nắm vững thì không được hoá giải. Yêu cầu mệnh chủ phải thành tâm phối hợp, đối chiếu với quá khứ mới có thể r
    #   út ra kết luận"" → trước khi áp dụng, đối chiếu các năm đã qua (LUU_NIEN, KHONG_RO_GIO mục "&"sự kiện) để chắc lá số và dụng thần đúng."
    return X.cat('"Khi gặp những Tứ trụ không hoàn toàn nắm vững thì không được hoá giải. Yêu cầu mệnh chủ phải thành tâm phối hợp, đối chiếu với quá khứ mới có thể rút ra kết luận" → trước khi áp dụng, đối chiếu các năm đã qua (LUU_NIEN, KHONG_RO_GIO mục ','sự kiện) để chắc lá số và dụng thần đúng.')

def f123(E, r, c):
    # Ô N53
    # = IF(MOD((MOD($M$43-11-1,9)+1)-$M$42,3)=0,"tháng 12 ("&INDEX(CHI_TEN,MOD(13,12)+1)&": Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (
    #   tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-((MOD($M$43-11-1,9)+1)),9)+1)&", Nhị hắc ở "&IND
    #   EX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2
    #   -((MOD($M$43-11-1,9)+1)),9)+1)&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('tháng 12 (',X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(13.0, 12.0),1.0))),': Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),11.0),1.0), 9.0),1.0)), 9.0),1.0))),', Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),11.0),1.0), 9.0),1.0)), 9.0),1.0))),'); ') if X.truthy(X.eq(X.MOD(X.sub(X.add(X.MOD(X.sub(X.sub(E.v(22,43,13),11.0),1.0), 9.0),1.0),E.v(22,42,13)), 3.0),0.0)) else '')

def f124(E, r, c):
    # Ô C56
    # = "• SỨC KHỎE CHUNG: "&LEFT(J6,FIND(CHAR(10),J6)-1)&" ("&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),"thân vượng",IF((IF(DOI_CHIEU!$C$3
    #   3=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"thân nhược","thân trung hòa"))&"). • BỆNH TẬT: "&LEFT(J7,FIND(CHAR(10),J7)-1)&"; "&C18&CHAR(10)&"• THƯƠNG TẬT
    #    – TAI HỌA: "&LEFT(J8,FIND(CHAR(10),J8)-1)&". "&SUBSTITUTE(C28,CHAR(10),"; ")&CHAR(10)&"• "&SUBSTITUTE(C38,CHAR(10),"; ")&CHAR(10)&"• HÓA GIẢI CHÍNH: 
    #   hoạt động ở phương dụng thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),5)+1)&" ("&CHOOSE((IF(DOI_CHIEU!$C$33=2,DU
    #   NG_HY_KY!$D$109,DUNG_HY_KY!$C$109))+1,"Đông","Nam","Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)","Tây","Bắc")&"), tránh phươn
    #   g kỵ thần; năm "&DAI_VAN!$C$23&" tránh Ngũ hoàng ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn (đông bắc)","Ly (nam)","Khảm (bắc)
    #   ","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(5-($M$42),9)+1)&" và Nhị hắc ở "&INDEX({"cung Giữa (tâm nhà)","Càn (tây bắc)","Đoài (tây)","Cấn
    #    (đông bắc)","Ly (nam)","Khảm (bắc)","Khôn (tây nam)","Chấn (đông)","Tốn (đông nam)"},MOD(2-($M$42),9)+1)&"; năm có phản/phục ngâm nên di chuyển, dời 
    #   chỗ ở (P2 tr.631-659)."&IF(OR(MOD(5-$M$42,9)=CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1,7,4,0,2,5),MOD(2-$M$42,9)=CHOOSE(IF(DO
    #   I_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1,7,4,0,2,5))," Lưu ý: năm "&DAI_VAN!$C$23&" sao xấu (Ngũ hoàng/Nhị hắc) rơi đúng phương dụng thầ
    #   n → trong nhà KHÔNG đặt giường, cửa ở phương đó mà sinh hoạt ở phương Lục bạch (P2 tr.638); còn khi đi xa vẫn theo phương dụng thần (P2 tr.631-632) – 
    #   hai việc khác nhau, không mâu thuẫn.","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• SỨC KHỎE CHUNG: ',X.LEFT(E.v(22,r-50,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-50,c+7)),1.0))),' ('),('thân vượng' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0)) else ('thân nhược' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else 'thân trung hòa'))),'). • BỆNH TẬT: '),X.LEFT(E.v(22,r-49,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-49,c+7)),1.0))),'; '),E.v(22,r-38,c)),X.CHAR(10.0)),'• THƯƠNG TẬT – TAI HỌA: '),X.LEFT(E.v(22,r-48,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(22,r-48,c+7)),1.0))),'. '),X.SUBSTITUTE(E.v(22,r-28,c), X.CHAR(10.0), '; ')),X.CHAR(10.0)),'• '),X.SUBSTITUTE(E.v(22,r-18,c), X.CHAR(10.0), '; ')),X.CHAR(10.0)),'• HÓA GIẢI CHÍNH: hoạt động ở phương dụng thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)), 5.0),1.0))),' ('),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'Đông', lambda: 'Nam', lambda: 'Trung cung (Thìn–Tuất–Sửu–Mùi: đông nam, tây bắc, đông bắc, tây nam)', lambda: 'Tây', lambda: 'Bắc')),'), tránh phương kỵ thần; năm '),E.v(14,23,3)),' tránh Ngũ hoàng ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(5.0,E.v(22,42,13)), 9.0),1.0))),' và Nhị hắc ở '),X.INDEX(X.Arr([['cung Giữa (tâm nhà)', 'Càn (tây bắc)', 'Đoài (tây)', 'Cấn (đông bắc)', 'Ly (nam)', 'Khảm (bắc)', 'Khôn (tây nam)', 'Chấn (đông)', 'Tốn (đông nam)']]), X.add(X.MOD(X.sub(2.0,E.v(22,42,13)), 9.0),1.0))),'; năm có phản/phục ngâm nên di chuyển, dời chỗ ở (P2 tr.631-659).'),(X.cat(X.cat(' Lưu ý: năm ',E.v(14,23,3)),' sao xấu (Ngũ hoàng/Nhị hắc) rơi đúng phương dụng thần → trong nhà KHÔNG đặt giường, cửa ở phương đó mà sinh hoạt ở phương Lục bạch (P2 tr.638); còn khi đi xa vẫn theo phương dụng thần (P2 tr.631-632) – hai việc khác nhau, không mâu thuẫn.') if X.truthy(X.OR(X.eq(X.MOD(X.sub(5.0,E.v(22,42,13)), 9.0),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 7.0, lambda: 4.0, lambda: 0.0, lambda: 2.0, lambda: 5.0)), X.eq(X.MOD(X.sub(2.0,E.v(22,42,13)), 9.0),X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 7.0, lambda: 4.0, lambda: 0.0, lambda: 2.0, lambda: 5.0)))) else ''))

def f125(E, r, c):
    # Ô E62, F62
    # = IF((PP1_DINH_TINH!$F$96>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,96,6),E.v(24,99,2))) else '—')

def f126(E, r, c):
    # Ô J62, J63, J64, J65, J66, J67, J68, J69, J70, J71, J72, J73 … (75 ô)
    # = IF(AND(E62="ĐỦ",F62="ĐỦ"),"cả hai",IF(E62="ĐỦ","PP1",IF(F62="ĐỦ","PP2","")))
    return ('cả hai' if X.truthy(X.AND(X.eq(E.v(22,r,c-5),'ĐỦ'), X.eq(E.v(22,r,c-4),'ĐỦ'))) else ('PP1' if X.truthy(X.eq(E.v(22,r,c-5),'ĐỦ')) else ('PP2' if X.truthy(X.eq(E.v(22,r,c-4),'ĐỦ')) else '')))

def f127(E, r, c):
    # Ô K62
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P
    #   2 tr.558): Tỉ kiếp ≥ P_NHIEU chữ → "&IF(J62="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J62&") ⇒ "&G62&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Tỉ kiếp ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f128(E, r, c):
    # Ô L62, L63, L64, L65, L66, L67, L68, L69, L70, L71, L72, L73 … (75 ô)
    # = IF(DOI_CHIEU!$C$33=2,F62,E62)
    return (E.v(22,r,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(22,r,c-7))

def f129(E, r, c):
    # Ô E63
    # = IF((DUNG_HY_KY!$C$7=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(8,7,3),1.0)) else '—')

def f130(E, r, c):
    # Ô F63
    # = IF((DUNG_HY_KY!$D$7=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(8,7,4),1.0)) else '—')

def f131(E, r, c):
    # Ô K63
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&" → Quy tắc (P2 tr.558): Thân vượng → "&IF(J63="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J63&") ⇒ "&G63&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),' → Quy tắc (P2 tr.558): Thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f132(E, r, c):
    # Ô E64, F64
    # = IF((PP1_DINH_TINH!$F$100>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,100,6),E.v(24,99,2))) else '—')

def f133(E, r, c):
    # Ô K64
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 
    #   tr.558): Ấn ≥ P_NHIEU chữ → "&IF(J64="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J64&") ⇒ "&G64&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f134(E, r, c):
    # Ô E65
    # = IF(AND((DUNG_HY_KY!$C$7=1),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f135(E, r, c):
    # Ô F65
    # = IF(AND((DUNG_HY_KY!$D$7=1),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f136(E, r, c):
    # Ô K65
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS
    #   (SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.558): Thân vượng và ấn vượng tại lệnh → "&IF(J65="","chưa đủ điều kiện ở cả hai phương ph
    #   áp – không áp dụng.","ĐỦ ("&J65&") ⇒ "&G65&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.558): Thân vượng và ấn vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f137(E, r, c):
    # Ô E66, F66
    # = IF(OR(TRUNG_GIAN!$E$28=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(23,28,5),0.0))) else '—')

def f138(E, r, c):
    # Ô K66
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.558): Nhật can ở đất Trường sinh tại chi ngày → "&IF(J66="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J66&") ⇒ "&G66&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.558): Nhật can ở đất Trường sinh tại chi ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f139(E, r, c):
    # Ô E67, F67
    # = IF(OR((THAN_SAT!$H$85="Có"),(THAN_SAT!$H$86="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có'))) else '—')

def f140(E, r, c):
    # Ô K67
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.558): Có Thiên đức hoặc Nguyệt đức → "&IF(J67="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J67&") ⇒ "&G67&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.558): Có Thiên đức hoặc Nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f141(E, r, c):
    # Ô E68, F68
    # = IF(AND((COUNTIFS(SAO!$R$5:$R$20,1)>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f142(E, r, c):
    # Ô K68
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   )&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.558): Có kình dương (theo nhật can) ở chi và có kiếp tài → "&IF(J68="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J68&") ⇒ "&G68&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.558): Có kình dương (theo nhật can) ở chi và có kiếp tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f143(E, r, c):
    # Ô E69, F69
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,4)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 4.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f144(E, r, c):
    # Ô K69
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.558): Kim ≤ 1 chữ, Hỏa ≥ P_NHIEU chữ → "&IF(J69="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J69&") ⇒ "&G69&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Kim ≤ 1 chữ, Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f145(E, r, c):
    # Ô E70, F70
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)))) else '—')

def f146(E, r, c):
    # Ô K70
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.558): Thổ ≤ 1 chữ, Mộc ≥ P_NHIEU chữ → "&IF(J70="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J70&") ⇒ "&G70&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Thổ ≤ 1 chữ, Mộc ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f147(E, r, c):
    # Ô E71, F71
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 2.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))) else '—')

def f148(E, r, c):
    # Ô K71
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.558): Hỏa ≤ 1 chữ, Thủy ≥ P_NHIEU chữ → "&IF(J71="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J71&") ⇒ "&G71&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Hỏa ≤ 1 chữ, Thủy ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f149(E, r, c):
    # Ô E72, F72
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)))) else '—')

def f150(E, r, c):
    # Ô K72
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.558): Mộc ≤ 1 chữ, Kim ≥ P_NHIEU chữ → "&IF(J72="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J72&") ⇒ "&G72&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Mộc ≤ 1 chữ, Kim ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f151(E, r, c):
    # Ô E73, F73
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)))) else '—')

def f152(E, r, c):
    # Ô K73
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.558): Thủy ≤ 1 chữ, Thổ ≥ P_NHIEU chữ → "&IF(J73="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J73&") ⇒ "&G73&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.558): Thủy ≤ 1 chữ, Thổ ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f153(E, r, c):
    # Ô E74, F74
    # = IF(AND(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.eq(X.INDEX(E.rg(8,24,3,28,3), 2.0),0.0))) else '—')

def f154(E, r, c):
    # Ô K74
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&" → Q
    #   uy tắc (P2 tr.558): Sinh tháng Hợi, Tí, Sửu (mùa đông) và không có Hỏa → "&IF(J74="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&
    #   J74&") ⇒ "&G74&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),' → Quy tắc (P2 tr.558): Sinh tháng Hợi, Tí, Sửu (mùa đông) và không có Hỏa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f155(E, r, c):
    # Ô E75
    # = IF(AND((DUNG_HY_KY!$C$7=0),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f156(E, r, c):
    # Ô F75
    # = IF(AND((DUNG_HY_KY!$D$7=0),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f157(E, r, c):
    # Ô K75
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ 
    #   liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,
    #   0)&" chữ; ",""))&" → Quy tắc (P2 tr.558): Thân nhược, thiên ấn ≥ P_NHIEU chữ → "&IF(J75="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J75&") ⇒ "&G75&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.558): Thân nhược, thiên ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f158(E, r, c):
    # Ô E76
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))
    #   >0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất 
    #   sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f159(E, r, c):
    # Ô F76
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))
    #   >0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất 
    #   sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f160(E, r, c):
    # Ô K76
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính 
    #   quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; 
    #   ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại l
    #   ệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$
    #   F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.559): Thân vượng, tài và quan sát đều có
    #    lực → "&IF(J76="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J76&") ⇒ "&G76&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.559): Thân vượng, tài và quan sát đều có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f161(E, r, c):
    # Ô E77, F77
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f162(E, r, c):
    # Ô K77
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.559): Có thiên ấn và
    #    có thực thần (kiêu đoạt thực) → "&IF(J77="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J77&") ⇒ "&G77&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.559): Có thiên ấn và có thực thần (kiêu đoạt thực) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f163(E, r, c):
    # Ô E78, F78
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2))>0),(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,
    #   0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f164(E, r, c):
    # Ô K78
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thiên ấn ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""
    #   ))&" → Quy tắc (P2 tr.559): Thiên ấn ở trụ ngày hoặc thiên ấn ≥ P_NHIEU chữ → "&IF(J78="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J78&") ⇒ "&G78&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.559): Thiên ấn ở trụ ngày hoặc thiên ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f165(E, r, c):
    # Ô E79
    # = IF(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f166(E, r, c):
    # Ô F79
    # = IF(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f167(E, r, c):
    # Ô K79
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.559, 563): Thân nhược, thực thương ≥ P_NHIEU chữ → "&IF(J79="","c
    #   hưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J79&") ⇒ "&G79&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.559, 563): Thân nhược, thực thương ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f168(E, r, c):
    # Ô E80, F80
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2))) else '—')

def f169(E, r, c):
    # Ô K80
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.560): Hỏa ≥ P_NHIEU chữ → "&IF(J80="","chưa đủ điều kiện ở cả hai phương pháp – không 
    #   áp dụng.","ĐỦ ("&J80&") ⇒ "&G80&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.560): Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f170(E, r, c):
    # Ô E81, F81
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0))) else '—')

def f171(E, r, c):
    # Ô K81
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,2)>0,"Chính quan ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$
    #   F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Chính quan ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thất sá
    #   t ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Thất sát ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5
    #   :$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.559): Trụ ngày và giờ có cả chính quan và thất sát → "&IF(J81="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J81&") ⇒ "&G81&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Chính quan ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.559): Trụ ngày và giờ có cả chính quan và thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f172(E, r, c):
    # Ô E82, F82
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0))) else '—')

def f173(E, r, c):
    # Ô K82
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Thương quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.560): Nữ, thương quan ở trụ năm → "&IF(J82="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("
    #   &J82&") ⇒ "&G82&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.560): Nữ, thương quan ở trụ năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f174(E, r, c):
    # Ô E83, F83
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR((THAN_SAT!$H$85="Có"),(THAN_SAT!$H$86="Có"))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có')))) else '—')

def f175(E, r, c):
    # Ô K83
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.560): Nữ, có Thiên đức hoặc Nguyệt đức → "&IF(J83="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J83&") ⇒ "&G83&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.560): Nữ, có Thiên đức hoặc Nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f176(E, r, c):
    # Ô E84
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUN
    #   TIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f177(E, r, c):
    # Ô F84
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUN
    #   TIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f178(E, r, c):
    # Ô K84
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS
    #   (SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.559-560): Thân nhược, thất sát (can lộ 
    #   hoặc bản khí) vượng tại lệnh → "&IF(J84="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J84&") ⇒ "&G84&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.559-560): Thân nhược, thất sát (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f179(E, r, c):
    # Ô E85
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f180(E, r, c):
    # Ô F85
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f181(E, r, c):
    # Ô K85
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUN
    #   TIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thấ
    #   t sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất 
    #   sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.560): Thân nhược, quan sát vượng tại lệnh → "&IF(J85="","chưa đủ điều kiện ở 
    #   cả hai phương pháp – không áp dụng.","ĐỦ ("&J85&") ⇒ "&G85&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Thân nhược, quan sát vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f182(E, r, c):
    # Ô E86, F86
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f183(E, r, c):
    # Ô K86
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y
    #   $5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.560, 564): Thương quan ≥ P_NHIE
    #   U chữ → "&IF(J86="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J86&") ⇒ "&G86&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.560, 564): Thương quan ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f184(E, r, c):
    # Ô E87, F87
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","
    #   —")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)), X.eq(E.v(9,98,6),0.0))) else '—')

def f185(E, r, c):
    # Ô K87
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.560): Có kiếp tài, c
    #   ó kình dương ở chi, không có tài → "&IF(J87="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J87&") ⇒ "&G87&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Có kiếp tài, có kình dương ở chi, không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f186(E, r, c):
    # Ô E88, F88
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f187(E, r, c):
    # Ô K88
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.565): Có thất sát và có
    #    thiên ấn → "&IF(J88="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J88&") ⇒ "&G88&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.565): Có thất sát và có thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f188(E, r, c):
    # Ô E89, F89
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f189(E, r, c):
    # Ô K89
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.563, 565
    #   ): Có thương quan và có chính quan → "&IF(J89="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J89&") ⇒ "&G89&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.563, 565): Có thương quan và có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f190(E, r, c):
    # Ô E90, F90
    # = IF(OR(AND(TRUNG_GIAN!$C$5=6,OR(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$D$5=3)),AND(TRUNG_GIAN!$D$5=6,OR(TRUNG_GIAN!$D$6=5,TR
    #   UNG_GIAN!$D$6=6,TRUNG_GIAN!$C$5=2,TRUNG_GIAN!$C$5=3,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$5=3)),AND(TRUNG_GIAN!$E$5=6,OR(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$E$6=6
    #   ,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$D$5=3,TRUNG_GIAN!$F$5=2,TRUNG_GIAN!$F$5=3)),AND(TRUNG_GIAN!$F$5=6,OR(TRUNG_GIAN!$F$6=5,TRUNG_GIAN!$F$6=6,TRUNG_GIAN!$E$
    #   5=2,TRUNG_GIAN!$E$5=3))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,5,3),6.0), X.OR(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,4),3.0))), X.AND(X.eq(E.v(23,5,4),6.0), X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,5,3),2.0), X.eq(E.v(23,5,3),3.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0))), X.AND(X.eq(E.v(23,5,5),6.0), X.OR(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,4),3.0), X.eq(E.v(23,5,6),2.0), X.eq(E.v(23,5,6),3.0))), X.AND(X.eq(E.v(23,5,6),6.0), X.OR(X.eq(E.v(23,6,6),5.0), X.eq(E.v(23,6,6),6.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0))))) else '—')

def f191(E, r, c):
    # Ô K90
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.559): Canh lộ ở thiên can, 
    #   bị Bính/Đinh khắc (can kề là Bính/Đinh, hoặc chi cùng trụ là Tị/Ngọ) → "&IF(J90="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J9
    #   0&") ⇒ "&G90&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.559): Canh lộ ở thiên can, bị Bính/Đinh khắc (can kề là Bính/Đinh, hoặc chi cùng trụ là Tị/Ngọ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f192(E, r, c):
    # Ô E91, F91
    # = IF(AND((TRUNG_GIAN!$E$5=0),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),0.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f193(E, r, c):
    # Ô K91
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.559): Nhật can Giáp, Hỏa ≥ P_NHIEU chữ → "&IF(J91="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J91&") ⇒ "&G91&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.559): Nhật can Giáp, Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f194(E, r, c):
    # Ô E92, F92
    # = IF(OR(AND(TRUNG_GIAN!$C$5=0,OR(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$D$5=6,TRUNG_GIAN!$D$5=7)),AND(TRUNG_GIAN!$D$5=0,OR(TRUNG_GIAN!$D$6=8,TR
    #   UNG_GIAN!$D$6=9,TRUNG_GIAN!$C$5=6,TRUNG_GIAN!$C$5=7,TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$5=7)),AND(TRUNG_GIAN!$E$5=0,OR(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$E$6=9
    #   ,TRUNG_GIAN!$D$5=6,TRUNG_GIAN!$D$5=7,TRUNG_GIAN!$F$5=6,TRUNG_GIAN!$F$5=7)),AND(TRUNG_GIAN!$F$5=0,OR(TRUNG_GIAN!$F$6=8,TRUNG_GIAN!$F$6=9,TRUNG_GIAN!$E$
    #   5=6,TRUNG_GIAN!$E$5=7))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,5,3),0.0), X.OR(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,5,4),7.0))), X.AND(X.eq(E.v(23,5,4),0.0), X.OR(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,5,3),6.0), X.eq(E.v(23,5,3),7.0), X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,5,5),7.0))), X.AND(X.eq(E.v(23,5,5),0.0), X.OR(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,5),9.0), X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,5,4),7.0), X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,5,6),7.0))), X.AND(X.eq(E.v(23,5,6),0.0), X.OR(X.eq(E.v(23,6,6),8.0), X.eq(E.v(23,6,6),9.0), X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,5,5),7.0))))) else '—')

def f195(E, r, c):
    # Ô K92
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.565): Giáp lộ ở thiên can, 
    #   bị Canh/Tân khắc (can kề là Canh/Tân, hoặc chi cùng trụ là Thân/Dậu) → "&IF(J92="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J9
    #   2&") ⇒ "&G92&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.565): Giáp lộ ở thiên can, bị Canh/Tân khắc (can kề là Canh/Tân, hoặc chi cùng trụ là Thân/Dậu) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f196(E, r, c):
    # Ô E96, F96
    # = IF(OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=6,OR(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$D$5=3)),AND(TRUE(),TRUNG_GIAN!$D$5=6,OR(TRU
    #   NG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$C$5=2,TRUNG_GIAN!$C$5=3,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$5=3)),AND(TRUE(),TRUNG_GIAN!$E$5=6,OR(TRUNG_GIAN!
    #   $E$6=5,TRUNG_GIAN!$E$6=6,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$D$5=3,TRUNG_GIAN!$F$5=2,TRUNG_GIAN!$F$5=3)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$
    #   F$5=6,OR(TRUNG_GIAN!$F$6=5,TRUNG_GIAN!$F$6=6,TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$5=3))),OR(AND(TRUE(),TRUNG_GIAN!$C$5=7,OR(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$C
    #   $6=6,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$D$5=3)),AND(TRUE(),TRUNG_GIAN!$D$5=7,OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$C$5=2,TRUNG_GIAN!$C$5=3,TRU
    #   NG_GIAN!$E$5=2,TRUNG_GIAN!$E$5=3)),AND(TRUE(),TRUNG_GIAN!$E$5=7,OR(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$E$6=6,TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$D$5=3,TRUNG_GIAN!
    #   $F$5=2,TRUNG_GIAN!$F$5=3)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=7,OR(TRUNG_GIAN!$F$6=5,TRUNG_GIAN!$F$6=6,TRUNG_GIAN!$E$5=2,TRUNG_GIA
    #   N!$E$5=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0), X.OR(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,4),3.0))), X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0), X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,5,3),2.0), X.eq(E.v(23,5,3),3.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0))), X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0), X.OR(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,4),3.0), X.eq(E.v(23,5,6),2.0), X.eq(E.v(23,5,6),3.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0), X.OR(X.eq(E.v(23,6,6),5.0), X.eq(E.v(23,6,6),6.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0), X.OR(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,4),3.0))), X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0), X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,5,3),2.0), X.eq(E.v(23,5,3),3.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0))), X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0), X.OR(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,4),3.0), X.eq(E.v(23,5,6),2.0), X.eq(E.v(23,5,6),3.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0), X.OR(X.eq(E.v(23,6,6),5.0), X.eq(E.v(23,6,6),6.0), X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,5),3.0)))))) else '—')

def f197(E, r, c):
    # Ô K96
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.558): Canh hoặc Tân lộ ở th
    #   iên can bị Bính/Đinh khắc (can kề, hoặc chi cùng trụ Tị/Ngọ) → "&IF(J96="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J96&") ⇒ "
    #   &G96&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.558): Canh hoặc Tân lộ ở thiên can bị Bính/Đinh khắc (can kề, hoặc chi cùng trụ Tị/Ngọ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f198(E, r, c):
    # Ô E97, F97
    # = IF(AND(OR(OR(AND(TRUE(),TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(TRUE(),NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6)=6),AND(TRUE(),TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(TRUE(),NOT(NHAP_LIEU!$C
    #   $9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6)),OR(A
    #   ND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$D$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$D$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$D$6=5),AND(TRU
    #   NG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$
    #   D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$D$6=1),AND(TRUNG_GIAN!$C$
    #   6=10,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=1),AN
    #   D(TRUNG_GIAN!$C$6=TRUNG_GIAN!$D$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!
    #   $C$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6=2),A
    #   ND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_
    #   GIAN!$E$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG_G
    #   IAN!$C$6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$E$6,O
    #   R(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$C$6=0,
    #   TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$F$6=2),AND(TRUN
    #   G_GIAN!$C$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$F
    #   $6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$C$
    #   6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$F$6,OR(TRUNG
    #   _GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11)))),AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$
    #   D$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=8),AN
    #   D(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_G
    #   IAN!$E$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_G
    #   IAN!$D$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$E$6,OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUN
    #   G_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$D$6=3,T
    #   RUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG
    #   _GIAN!$D$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$F$
    #   6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$D$
    #   6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!
    #   $D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$E$6=3,TRUNG_GI
    #   AN!$F$6=0),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$
    #   E$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$E$6=1,TRUNG_GIAN!$F$6=10),A
    #   ND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$E$6=1,TRU
    #   NG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$E$6=6,TRUNG_GIAN!$E$6=9,
    #   TRUNG_GIAN!$E$6=11)))))),OR(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$E$6,12)=1),
    #   AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$E$6,12)=1),
    #   AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!
    #   $E$6+TRUNG_GIAN!$F$6,12)=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.OR(X.AND(X.TRUE(), X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0))), X.OR(X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,4),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,4),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,4)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,5),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,6,5),9.0), X.eq(E.v(23,6,5),11.0))))))), X.OR(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))))) else '—')

def f199(E, r, c):
    # Ô K97
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.558): Có cặp chi xung hoặc 
    #   hình, đồng thời có cặp chi lục hợp → "&IF(J97="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J97&") ⇒ "&G97&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.558): Có cặp chi xung hoặc hình, đồng thời có cặp chi lục hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f200(E, r, c):
    # Ô E98, F98
    # = IF(AND(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRU
    #   NG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),T
    #   RUNG_GIAN!$F$6=6))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),TRUNG_GIAN!$F$6=10))),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f201(E, r, c):
    # Ô K98
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUN
    #   G_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1
    #   )&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy 
    #   tắc (P2 tr.558): Tứ trụ đủ Dần, Ngọ, Tuất (tam hợp hỏa cục) và Hỏa ≥ P_NHIEU chữ → "&IF(J98="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng
    #   .","ĐỦ ("&J98&") ⇒ "&G98&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.558): Tứ trụ đủ Dần, Ngọ, Tuất (tam hợp hỏa cục) và Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f202(E, r, c):
    # Ô E99, F99
    # = IF(AND((COUNTIFS(SAO!$R$5:$R$20,1)>0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f203(E, r, c):
    # Ô K99
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.559): Có kình dương và có ấn → "&IF(J99
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J99&") ⇒ "&G99&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.559): Có kình dương và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f204(E, r, c):
    # Ô E100, F100
    # = IF(AND(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU)),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2))), X.eq(X.INDEX(E.rg(8,24,3,28,3), 5.0),0.0))) else '—')

def f205(E, r, c):
    # Ô K100
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.559): Hỏa ≥ P_NHIEU, Thổ ≥ P_NHIEU và không có Thủy → "&IF(J100="","chưa đủ điều kiện 
    #   ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J100&") ⇒ "&G100&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.559): Hỏa ≥ P_NHIEU, Thổ ≥ P_NHIEU và không có Thủy → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f206(E, r, c):
    # Ô E101, F101
    # = IF(AND(AND((TRUNG_GIAN!$E$7=3),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),OR(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$
    #   6=0,TRUNG_GIAN!$D$6=1),OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,7,5),3.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))))) else '—')

def f207(E, r, c):
    # Ô K101
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&
    #   "Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên
    #    quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$
    #   5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.559): Nhật chủ Kim, có thương quan (Thủy), sinh mùa đông (hà
    #   n) hoặc mùa hạ (nhiệt) → "&IF(J101="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J101&") ⇒ "&G101&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.559): Nhật chủ Kim, có thương quan (Thủy), sinh mùa đông (hàn) hoặc mùa hạ (nhiệt) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f208(E, r, c):
    # Ô E102, F102
    # = IF(AND((TRUNG_GIAN!$E$7=2),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),2.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f209(E, r, c):
    # Ô K102
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.559): Nhật chủ Thổ, ấn là Hỏa ≥ P_NHIEU chữ (hỏa thổ ấn thụ, khô nóng) → "&IF(J102="","chưa đủ điều
    #    kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J102&") ⇒ "&G102&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.559): Nhật chủ Thổ, ấn là Hỏa ≥ P_NHIEU chữ (hỏa thổ ấn thụ, khô nóng) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f210(E, r, c):
    # Ô E103, F103
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)))) else '—')

def f211(E, r, c):
    # Ô K103
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.559-560): Thủy ≥ P_NHIEU và Mộc ≥ P_NHIEU (thủy mộc tương sinh) → "&IF(J103="","chưa đ
    #   ủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J103&") ⇒ "&G103&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.559-560): Thủy ≥ P_NHIEU và Mộc ≥ P_NHIEU (thủy mộc tương sinh) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f212(E, r, c):
    # Ô E104, F104
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),AND(TRUE(),TRUNG_GIAN!$E$5=2),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3),AND(TRUE(),TRUNG_GIAN!$E$5=3),AN
    #   D(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0)))))) else '—')

def f213(E, r, c):
    # Ô K104
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUN
    #   G_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1
    #   )&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy 
    #   tắc (P2 tr.559): Thổ ≥ P_NHIEU chữ và Hỏa lộ ở thiên can (Bính/Đinh) → "&IF(J104="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J
    #   104&") ⇒ "&G104&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.559): Thổ ≥ P_NHIEU chữ và Hỏa lộ ở thiên can (Bính/Đinh) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f214(E, r, c):
    # Ô E105, F105
    # = IF(AND((THAN_SAT!$H$111="Có"),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(12,111,8),'Có'), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0))) else '—')

def f215(E, r, c):
    # Ô K105
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.559): Có kim thần và có Thủy → "&IF(J105="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J105&") ⇒ "&G105&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.559): Có kim thần và có Thủy → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f216(E, r, c):
    # Ô E106, F106
    # = IF(AND(OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),AND(TRUE(),TRUNG_GIAN!$E$5=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GI
    #   AN!$F$5=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3),AND(TRUE(),TRUNG_GIAN!$E$5=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_
    #   GIAN!$F$5=3))),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0)))), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f217(E, r, c):
    # Ô K106
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUN
    #   G_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1
    #   )&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy 
    #   tắc (P2 tr.559): Bính hoặc Đinh lộ ở thiên can và Hỏa ≥ P_NHIEU chữ → "&IF(J106="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J1
    #   06&") ⇒ "&G106&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.559): Bính hoặc Đinh lộ ở thiên can và Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f218(E, r, c):
    # Ô E107, F107
    # = IF(AND(AND(TRUE(),OR(TRUNG_GIAN!$E$28=5,TRUNG_GIAN!$E$28=6,TRUNG_GIAN!$E$28=7,TRUNG_GIAN!$E$28=8,TRUNG_GIAN!$E$28=9)),AND(NOT(NHAP_LIEU!$C$9="Không rõ
    #    giờ"),OR(TRUNG_GIAN!$F$28=5,TRUNG_GIAN!$F$28=6,TRUNG_GIAN!$F$28=7,TRUNG_GIAN!$F$28=8,TRUNG_GIAN!$F$28=9))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,28,5),5.0), X.eq(E.v(23,28,5),6.0), X.eq(E.v(23,28,5),7.0), X.eq(E.v(23,28,5),8.0), X.eq(E.v(23,28,5),9.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,28,6),5.0), X.eq(E.v(23,28,6),6.0), X.eq(E.v(23,28,6),7.0), X.eq(E.v(23,28,6),8.0), X.eq(E.v(23,28,6),9.0))))) else '—')

def f219(E, r, c):
    # Ô K107
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.559): Nhật can tại chi ngày và tại chi giờ đều ở Suy – Bệnh – Tử – Mộ – Tuyệt → "&IF(J107="","chưa đủ
    #    điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J107&") ⇒ "&G107&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.559): Nhật can tại chi ngày và tại chi giờ đều ở Suy – Bệnh – Tử – Mộ – Tuyệt → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f220(E, r, c):
    # Ô E108, F108
    # = IF(AND((TRUNG_GIAN!$E$7=4),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),4.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)))) else '—')

def f221(E, r, c):
    # Ô K108
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.559): Nhật chủ Thủy (Nhâm/Quý) và Thổ ≥ P_NHIEU chữ → "&IF(J108="","chưa đủ điều kiện ở cả hai phươ
    #   ng pháp – không áp dụng.","ĐỦ ("&J108&") ⇒ "&G108&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.559): Nhật chủ Thủy (Nhâm/Quý) và Thổ ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f222(E, r, c):
    # Ô E109, F109
    # = IF(AND((TRUNG_GIAN!$E$7=3),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))) else '—')

def f223(E, r, c):
    # Ô K109
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(
    #   HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.559): Nhật chủ Kim và Thủy ≥ P_NHIEU chữ → "&IF(J109="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J109&") ⇒ "&G109&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.559): Nhật chủ Kim và Thủy ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f224(E, r, c):
    # Ô E110, F110
    # = IF(AND(OR(OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!$D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG
    #   _GIAN!$F$6=11)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=0),AND(TRUE(),TRUNG_GIAN!$D$6=0),AND(TRUE(),TRUNG_GIAN!$E$6=0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TR
    #   UNG_GIAN!$F$6=0))),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ gi
    #   ờ"),TRUNG_GIAN!$F$6=5)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=6)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),0.0)))), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0)))))) else '—')

def f225(E, r, c):
    # Ô K110
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.559): Có Hợi hoặc Tí, và có
    #    Tị hoặc Ngọ → "&IF(J110="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J110&") ⇒ "&G110&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.559): Có Hợi hoặc Tí, và có Tị hoặc Ngọ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f226(E, r, c):
    # Ô E111, F111
    # = IF((AND(TRUE(),OR(TRUNG_GIAN!$C$6=11,TRUNG_GIAN!$C$6=0))+AND(TRUE(),OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0))+AND(TRUE(),OR(TRUNG_GIAN!$E$6=11,TRUNG_G
    #   IAN!$E$6=0))+AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=11,TRUNG_GIAN!$F$6=0)))>=2,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),11.0), X.eq(E.v(23,6,3),0.0))),X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0)))),X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),11.0), X.eq(E.v(23,6,5),0.0)))),X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),11.0), X.eq(E.v(23,6,6),0.0)))),2.0)) else '—')

def f227(E, r, c):
    # Ô K111
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.559): Có từ 2 chi Hợi/Tí tr
    #   ở lên → "&IF(J111="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J111&") ⇒ "&G111&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.559): Có từ 2 chi Hợi/Tí trở lên → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f228(E, r, c):
    # Ô E112, F112
    # = IF(AND((AND(TRUE(),TRUNG_GIAN!$C$5=3)+AND(TRUE(),TRUNG_GIAN!$D$5=3)+AND(TRUE(),TRUNG_GIAN!$E$5=3)+AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F
    #   $5=3)+(AND(TRUE(),OR(TRUNG_GIAN!$C$6=6))+AND(TRUE(),OR(TRUNG_GIAN!$D$6=6))+AND(TRUE(),OR(TRUNG_GIAN!$E$6=6))+AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR
    #   (TRUNG_GIAN!$F$6=6))))>=2,OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không r
    #   õ giờ"),TRUNG_GIAN!$F$6=7))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.add(X.add(X.add(X.add(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)),X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0))),X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0))),X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0))),X.add(X.add(X.add(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),6.0))),X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),6.0)))),X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),6.0)))),X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),6.0))))),2.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))) else '—')

def f229(E, r, c):
    # Ô K112
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.560): Số chữ Đinh (can) + N
    #   gọ (chi) ≥ 2 và có Mùi → "&IF(J112="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J112&") ⇒ "&G112&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.560): Số chữ Đinh (can) + Ngọ (chi) ≥ 2 và có Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f230(E, r, c):
    # Ô E113, F113
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))) else '—')

def f231(E, r, c):
    # Ô K113
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.560): Mộc ≥ P_NHIEU và Hỏa ≥ P_NHIEU → "&IF(J113="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J113&") ⇒ "&G113&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.560): Mộc ≥ P_NHIEU và Hỏa ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f232(E, r, c):
    # Ô E114, F114
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))) else '—')

def f233(E, r, c):
    # Ô K114
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.560): Kim ≥ P_NHIEU và Thủy ≥ P_NHIEU → "&IF(J114="","chưa đủ điều kiện ở cả hai phươn
    #   g pháp – không áp dụng.","ĐỦ ("&J114&") ⇒ "&G114&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.560): Kim ≥ P_NHIEU và Thủy ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f234(E, r, c):
    # Ô E115, F115
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$97>=P_NHIEU)),(PP1_DINH_TINH!$F$97>COUNTIFS(S
    #   AO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,97,6),E.v(24,99,2))), X.gt(E.v(9,97,6),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)))) else '—')

def f235(E, r, c):
    # Ô K115
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các 
    #   chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.560): Có thất sát, thực thương ≥ P_NHIEU và nhiều hơn thất sát (chế phục t
    #   hái quá) → "&IF(J115="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J115&") ⇒ "&G115&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Có thất sát, thực thương ≥ P_NHIEU và nhiều hơn thất sát (chế phục thái quá) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f236(E, r, c):
    # Ô E116, F116
    # = IF(AND(AND((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)=1),(PP1_DINH_TINH!$F$98=3)),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),1.0), X.eq(E.v(9,98,6),3.0)), X.eq(X.INDEX(E.rg(8,24,3,28,3), 2.0),2.0))) else '—')

def f237(E, r, c):
    # Ô K116
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắ
    #   c (P2 tr.561): Đúng 1 chữ thất sát, đúng 3 chữ tài, đúng 2 chữ Hỏa → "&IF(J116="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J11
    #   6&") ⇒ "&G116&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.561): Đúng 1 chữ thất sát, đúng 3 chữ tài, đúng 2 chữ Hỏa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f238(E, r, c):
    # Ô E117, F117
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),((COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0))) else '—')

def f239(E, r, c):
    # Ô K117
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thất sát bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản k
    #   hí",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)>
    #   0,"Thiên ấn bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; "
    #   ,""))&" → Quy tắc (P2 tr.560): Bản khí chi ngày là thất sát hoặc thiên ấn → "&IF(J117="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ
    #    ("&J117&") ⇒ "&G117&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Bản khí chi ngày là thất sát hoặc thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f240(E, r, c):
    # Ô E118, F118
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DO
    #   I_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+4,5),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else '—')

def f241(E, r, c):
    # Ô K118
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Hỷ PP1: "&IF(DUNG_HY_KY!$C$110>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$110>=0,INDEX(HANH_TEN,DUN
    #   G_HY_KY!$D$110+1),"—")&"; "&"Kỵ PP1: "&IF(DUNG_HY_KY!$C$111>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$111+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$111>=0,INDEX(HANH_
    #   TEN,DUNG_HY_KY!$D$111+1),"—")&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIA
    #   N!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr
    #   .560): Nữ, có thiên ấn và hành ấn là kỵ thần → "&IF(J118="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J118&") ⇒ "&G118&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Hỷ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)) if X.truthy(X.ge(E.v(8,110,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0)) if X.truthy(X.ge(E.v(8,110,4),0.0)) else '—')),'; '),'Kỵ PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)) if X.truthy(X.ge(E.v(8,111,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0)) if X.truthy(X.ge(E.v(8,111,4),0.0)) else '—')),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.560): Nữ, có thiên ấn và hành ấn là kỵ thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f242(E, r, c):
    # Ô E119, F119
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(COUNTIFS(SAO!$F$5:$F$20,"Thiên 
    #   ấn",SAO!$Y$5:$Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f243(E, r, c):
    # Ô K119
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&
    #   " → Quy tắc (P2 tr.560): Nữ, có thực thần và thiên ấn ≥ P_NHIEU chữ → "&IF(J119="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J1
    #   19&") ⇒ "&G119&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Giới tính: '),E.v(2,22,3)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Nữ, có thực thần và thiên ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f244(E, r, c):
    # Ô E120, F120
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),AND(TRUNG_GIAN!$C$12="Thương quan",NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$R$17=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.AND(X.eq(E.v(23,12,3),'Thương quan'), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,18),1.0)))) else '—')

def f245(E, r, c):
    # Ô K120
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.560): Nữ, can năm là thương quan và chi giờ là kình dương → "&IF(J120="","chưa đủ điều ki
    #   ện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J120&") ⇒ "&G120&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.560): Nữ, can năm là thương quan và chi giờ là kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f246(E, r, c):
    # Ô E121, F121
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(OR(AND(TRUE(),LEN(THAN_SAT!$E$120)>0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$120)>0)),OR(AND(TR
    #   UE(),LEN(THAN_SAT!$E$121)>0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$121)>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.OR(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,120,5)),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,120,6)),0.0))), X.OR(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,121,5)),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,121,6)),0.0)))))) else '—')

def f247(E, r, c):
    # Ô K121
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.560): Nữ, sao Câu hoặc Giảo ở trụ ngày hoặc giờ → "&IF(J121="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J121&") ⇒ "&G121&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.560): Nữ, sao Câu hoặc Giảo ở trụ ngày hoặc giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f248(E, r, c):
    # Ô E122, F122
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),(AND(TRUE(),SAO!$R$5=1)+AND(TRUE(),SAO!$R$9=1)+AND(TRUE(),SAO!$R$13=1)+AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!
    #   $R$17=1))>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.ge(X.add(X.add(X.add(X.AND(X.TRUE(), X.eq(E.v(13,5,18),1.0)),X.AND(X.TRUE(), X.eq(E.v(13,9,18),1.0))),X.AND(X.TRUE(), X.eq(E.v(13,13,18),1.0))),X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,18),1.0))),2.0))) else '—')

def f249(E, r, c):
    # Ô K122
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&" → Quy tắc (P2 tr.560): Nữ, kình dương ≥ 2 chi → "&IF(J122="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J122&") ⇒ "&G122&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),' → Quy tắc (P2 tr.560): Nữ, kình dương ≥ 2 chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f250(E, r, c):
    # Ô E123, F123
    # = IF(AND(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU
    #   !$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=3))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=9))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0))))) else '—')

def f251(E, r, c):
    # Ô K123
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.560): Nữ, có cả Mão và Dậu → "&IF(J123="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J123&") ⇒ "&G123&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.560): Nữ, có cả Mão và Dậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f252(E, r, c):
    # Ô E124, F124
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thi
    #   ên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f253(E, r, c):
    # Ô K124
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1
    #   )>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1
    #   )&" chữ; ",""))&" → Quy tắc (P2 tr.560): Thiên ấn (can lộ hoặc bản khí) vượng tại lệnh → "&IF(J124="","chưa đủ điều kiện ở cả hai phương pháp – không 
    #   áp dụng.","ĐỦ ("&J124&") ⇒ "&G124&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Thiên ấn (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f254(E, r, c):
    # Ô E125, F125
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))) else '—')

def f255(E, r, c):
    # Ô K125
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.560): Thổ ≤ 1 chữ và Thủy ≥ P_NHIEU chữ → "&IF(J125="","chưa đủ điều kiện ở cả hai phư
    #   ơng pháp – không áp dụng.","ĐỦ ("&J125&") ⇒ "&G125&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.560): Thổ ≤ 1 chữ và Thủy ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f256(E, r, c):
    # Ô E126, F126
    # = IF((AND(TRUE(),SAO!$R$5=1)+AND(TRUE(),SAO!$R$9=1)+AND(TRUE(),SAO!$R$13=1)+AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$R$17=1))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.AND(X.TRUE(), X.eq(E.v(13,5,18),1.0)),X.AND(X.TRUE(), X.eq(E.v(13,9,18),1.0))),X.AND(X.TRUE(), X.eq(E.v(13,13,18),1.0))),X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,18),1.0))),3.0)) else '—')

def f257(E, r, c):
    # Ô K126
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.560): Kình dương ≥ 3 chi → "&IF(J126="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ
    #    ("&J126&") ⇒ "&G126&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.560): Kình dương ≥ 3 chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f258(E, r, c):
    # Ô E127, F127
    # = IF(AND((COUNTIFS(SAO!$R$5:$R$20,1)>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f259(E, r, c):
    # Ô K127
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   )&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.560): Có kình dương và có thất sát → "&IF(J127="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J127&") ⇒ "&G127&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.560): Có kình dương và có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f260(E, r, c):
    # Ô E128, F128
    # = IF(AND(OR(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),
    #   TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"
    #   ),TRUNG_GIAN!$F$6=5))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8),AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),TRUNG_GIAN!$F$6=8))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!
    #   $C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7)))),OR(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$D$6,12)=1),AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6+T
    #   RUNG_GIAN!$E$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6+TRUNG_GIAN!$F$6,12)=1),AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6+T
    #   RUNG_GIAN!$E$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6+TRUNG_GIAN!$F$6,12)=1),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ
    #    giờ"),MOD(TRUNG_GIAN!$E$6+TRUNG_GIAN!$F$6,12)=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))), X.OR(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,4)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,3),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,5)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,4),E.v(23,6,6)), 12.0),1.0)), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(E.v(23,6,5),E.v(23,6,6)), 12.0),1.0))))) else '—')

def f261(E, r, c):
    # Ô K128
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.560): Có tam hình đủ 3 chi 
    #   (Dần–Tị–Thân hoặc Sửu–Tuất–Mùi) và có cặp chi lục hợp → "&IF(J128="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J128&") ⇒ "&G128
    #   &".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.560): Có tam hình đủ 3 chi (Dần–Tị–Thân hoặc Sửu–Tuất–Mùi) và có cặp chi lục hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f262(E, r, c):
    # Ô E129, F129
    # = IF((AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$C$6=0,TRUNG_GIAN!$D$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$D$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$D$6=5),AN
    #   D(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_G
    #   IAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$D$6=1),AND(TRUNG_GIA
    #   N!$C$6=10,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$D$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$D$6=
    #   1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$D$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11))))+AND(TRUE(),TRUE(),OR(AND(TRUNG_
    #   GIAN!$C$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6
    #   =2),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$C$6=8,T
    #   RUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$E$6=7),AND(TR
    #   UNG_GIAN!$C$6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$
    #   E$6,OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11))))+AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$C
    #   $6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$C$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$F$6=2),AND
    #   (TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GI
    #   AN!$F$6=2),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIA
    #   N!$C$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$C$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$C$6=TRUNG_GIAN!$F$6,OR(
    #   TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=6,TRUNG_GIAN!$C$6=9,TRUNG_GIAN!$C$6=11))))+AND(TRUE(),TRUE(),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$E$6=3),AND(TRUNG_G
    #   IAN!$D$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=
    #   8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$D$6=1,TR
    #   UNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=10),AND(TR
    #   UNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$E$6,OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6
    #   ,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11))))+AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$D$
    #   6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=8),AND(
    #   TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIA
    #   N!$F$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIA
    #   N!$D$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUNG_
    #   GIAN!$D$6=9,TRUNG_GIAN!$D$6=11))))+AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$E$6=0,TRUNG_GIAN!$F$6=3),AND(TRUNG_GIAN!$E$6=3,TRU
    #   NG_GIAN!$F$6=0),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$F$6=8),AND(TRUNG_G
    #   IAN!$E$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$E$6=1,TRUNG_GIAN!$F$6=
    #   10),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$E$6=
    #   1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$E$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$E$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$E$6=6,TRUNG_GIAN!$E
    #   $6=9,TRUNG_GIAN!$E$6=11)))))>=2,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.add(X.add(X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,4),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,4),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,4),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,4),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,4)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0))))),X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0)))))),X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,3),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,3),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,3),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,3),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),6.0), X.eq(E.v(23,6,3),9.0), X.eq(E.v(23,6,3),11.0)))))),X.AND(X.TRUE(), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0)))))),X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0)))))),X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,5),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,5),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,5),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,5),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),6.0), X.eq(E.v(23,6,5),9.0), X.eq(E.v(23,6,5),11.0)))))),2.0)) else '—')

def f263(E, r, c):
    # Ô K129
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.560): Có từ 2 cặp chi hình 
    #   trở lên (kể cả tự hình) → "&IF(J129="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J129&") ⇒ "&G129&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.560): Có từ 2 cặp chi hình trở lên (kể cả tự hình) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f264(E, r, c):
    # Ô E130, F130
    # = IF(OR(AND(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$D$6=8),AND(TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$D$6=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,4),8.0)), X.AND(X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,4),2.0)))) else '—')

def f265(E, r, c):
    # Ô K130
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&" → Quy tắc (P2 tr.
    #   561): Chi năm và chi tháng là Dần – Thân → "&IF(J130="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J130&") ⇒ "&G130&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),' → Quy tắc (P2 tr.561): Chi năm và chi tháng là Dần – Thân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f266(E, r, c):
    # Ô E131, F131
    # = IF((OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$
    #   6=4))+OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIA
    #   N!$F$6=10))+OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_
    #   GIAN!$F$6=1))+OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUN
    #   G_GIAN!$F$6=7)))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))),X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))),X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0)))),X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0)))),3.0)) else '—')

def f267(E, r, c):
    # Ô K131
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.561): Có ít nhất 3 trong 4 
    #   chi Thìn, Tuất, Sửu, Mùi → "&IF(J131="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J131&") ⇒ "&G131&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.561): Có ít nhất 3 trong 4 chi Thìn, Tuất, Sửu, Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f268(E, r, c):
    # Ô E132, F132
    # = IF(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!
    #   $F$6=3)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIA
    #   N!$F$6=9))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0))))) else '—')

def f269(E, r, c):
    # Ô K132
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.561): Có cả Mão và Dậu (xun
    #   g) → "&IF(J132="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J132&") ⇒ "&G132&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.561): Có cả Mão và Dậu (xung) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f270(E, r, c):
    # Ô E133, F133
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else '—')

def f271(E, r, c):
    # Ô K133
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Bản khí",SAO!$B$5:$B$20,2)>0,"Thương quan bản khí ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Bản khí",SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.561): Bản khí chi ngày là thương quan → "&IF(J133="","chưa đủ điều kiện ở cả hai phươn
    #   g pháp – không áp dụng.","ĐỦ ("&J133&") ⇒ "&G133&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan bản khí ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0)) else ''))),' → Quy tắc (P2 tr.561): Bản khí chi ngày là thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f272(E, r, c):
    # Ô C137, D137, E137, F137, G137, H137, I137, J137
    # = DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-102,c),' ('),E.v(14,r-101,c)),'–'),X.add(E.v(14,r-101,c),9.0)),')')

def f273(E, r, c):
    # Ô C138, D138, E138, F138, G138, H138, I138, J138
    # = IF(TRIM(C139&" "&C140&" "&C141&" "&C142&" "&C143)="","Không có tạng nào bị khắc rõ trong vận này.",TRIM(C139&" "&C140&" "&C141&" "&C142&" "&C143))
    return ('Không có tạng nào bị khắc rõ trong vận này.' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+1,c),' '),E.v(22,r+2,c)),' '),E.v(22,r+3,c)),' '),E.v(22,r+4,c)),' '),E.v(22,r+5,c))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+1,c),' '),E.v(22,r+2,c)),' '),E.v(22,r+3,c)),' '),E.v(22,r+4,c)),' '),E.v(22,r+5,c))))

def f274(E, r, c):
    # Ô C139, D139, E139, F139, G139, H139, I139, J139
    # = IF(OR(AND(OR((DAI_VAN!C42=3),(DAI_VAN!C43=3)),OR(OR(INDEX(M_TS,0+1,DAI_VAN!C34+1)=7,INDEX(M_TS,0+1,DAI_VAN!C34+1)=9),IF(TRUNG_GIAN!$E$7=0,(IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU))),AND(OR(INDEX(M_TS,0+1,D
    #   AI_VAN!C34+1)=7,INDEX(M_TS,0+1,DAI_VAN!C34+1)=9),OR(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$10
    #   9,DUNG_HY_KY!$C$109)=0)))),"• "&IF((IF((DAI_VAN!C42=3),1,0)+IF((DAI_VAN!C43=3),1,0)+IF(OR(INDEX(M_TS,0+1,DAI_VAN!C34+1)=7,INDEX(M_TS,0+1,DAI_VAN!C34+1
    #   )=9),1,0)+IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)
    #   ,DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=0,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY
    #   _KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)
    #   =0)),1,0))>=4,"NẶNG",IF((IF((DAI_VAN!C42=3),1,0)+IF((DAI_VAN!C43=3),1,0)+IF(OR(INDEX(M_TS,0+1,DAI_VAN!C34+1)=7,INDEX(M_TS,0+1,DAI_VAN!C34+1)=9),1,0)+I
    #   F(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C4
    #   3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=0,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0
    #   ),INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=0)),1,0))=
    #   3,"ĐỀ PHÒNG","CHÚ Ý"))&" – gan – mật (thắt lưng, gân cốt): "&IF((DAI_VAN!C42=3),"can mang Kim khắc; ","")&IF((DAI_VAN!C43=3),"chi mang Kim khắc; ","")
    #   &IF(OR(INDEX(M_TS,0+1,DAI_VAN!C34+1)=7,INDEX(M_TS,0+1,DAI_VAN!C34+1)=9),"Mộc ở đất "&IF(INDEX(M_TS,0+1,DAI_VAN!C34+1)=7,"Tử","Tuyệt")&" tại "&INDEX(CH
    #   I_TEN,DAI_VAN!C34+1)&" (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ","")&IF(IF(TRUNG_GIAN!$E$7=0,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$
    #   7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1),"Mộc yếu trong mệnh"&IF(AND(TRUNG_GIAN!$E$7=0,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0))," (
    #   hành nhật chủ, thân nhược)","")&"; ","")&IF((INDEX(DUNG_HY_KY!$C$24:$C$28,4)>=P_NHIEU),"Kim vượng sẵn trong mệnh; ","")&IF(AND(IF(DOI_CHIEU!$C$33=2,DU
    #   NG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$111,DUNG_HY_KY!$C$111))),"vận/năm mang hành kỵ thần; ","")&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=0),"Mộc là dụng thần;
    #    ",""),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',('NẶNG' if X.truthy(X.ge(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-97,c),3.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-96,c),3.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-96,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),0.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0))) else 0.0)),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-97,c),3.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-96,c),3.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-96,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),0.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0))) else 0.0)),3.0)) else 'CHÚ Ý'))),' – gan – mật (thắt lưng, gân cốt): '),('can mang Kim khắc; ' if X.truthy(X.eq(E.v(14,r-97,c),3.0)) else '')),('chi mang Kim khắc; ' if X.truthy(X.eq(E.v(14,r-96,c),3.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Mộc ở đất ',('Tử' if X.truthy(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),7.0)) else 'Tuyệt')),' tại '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-105,c),1.0))),' (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),9.0))) else '')),(X.cat(X.cat('Mộc yếu trong mệnh',(' (hành nhật chủ, thân nhược)' if X.truthy(X.AND(X.eq(E.v(23,7,5),0.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else '')),'; ') if X.truthy((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),0.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0))) else '')),('Kim vượng sẵn trong mệnh; ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2))) else '')),('vận/năm mang hành kỵ thần; ' if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-96,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else '')),('Mộc là dụng thần; ' if X.truthy(X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0)) else '')) if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(14,r-97,c),3.0), X.eq(E.v(14,r-96,c),3.0)), X.OR(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),9.0)), (X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),0.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 4.0),E.v(24,99,2)))), X.AND(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(0.0,1.0), X.add(E.v(14,r-105,c),1.0)),9.0)), X.OR(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-96,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0))))) else '')

def f275(E, r, c):
    # Ô C140, D140, E140, F140, G140, H140, I140, J140
    # = IF(OR(AND(OR((DAI_VAN!C42=4),(DAI_VAN!C43=4)),OR(OR(INDEX(M_TS,2+1,DAI_VAN!C34+1)=7,INDEX(M_TS,2+1,DAI_VAN!C34+1)=9),IF(TRUNG_GIAN!$E$7=1,(IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,2)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU))),AND(OR(INDEX(M_TS,2+1,D
    #   AI_VAN!C34+1)=7,INDEX(M_TS,2+1,DAI_VAN!C34+1)=9),OR(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$10
    #   9,DUNG_HY_KY!$C$109)=1)))),"• "&IF((IF((DAI_VAN!C42=4),1,0)+IF((DAI_VAN!C43=4),1,0)+IF(OR(INDEX(M_TS,2+1,DAI_VAN!C34+1)=7,INDEX(M_TS,2+1,DAI_VAN!C34+1
    #   )=9),1,0)+IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)
    #   ,DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=1,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY
    #   _KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,2)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)
    #   =1)),1,0))>=4,"NẶNG",IF((IF((DAI_VAN!C42=4),1,0)+IF((DAI_VAN!C43=4),1,0)+IF(OR(INDEX(M_TS,2+1,DAI_VAN!C34+1)=7,INDEX(M_TS,2+1,DAI_VAN!C34+1)=9),1,0)+I
    #   F(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C4
    #   3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=1,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0
    #   ),INDEX(DUNG_HY_KY!$C$24:$C$28,2)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=1)),1,0))=
    #   3,"ĐỀ PHÒNG","CHÚ Ý"))&" – tim – máu, mắt: "&IF((DAI_VAN!C42=4),"can mang Thủy khắc; ","")&IF((DAI_VAN!C43=4),"chi mang Thủy khắc; ","")&IF(OR(INDEX(M
    #   _TS,2+1,DAI_VAN!C34+1)=7,INDEX(M_TS,2+1,DAI_VAN!C34+1)=9),"Hỏa ở đất "&IF(INDEX(M_TS,2+1,DAI_VAN!C34+1)=7,"Tử","Tuyệt")&" tại "&INDEX(CHI_TEN,DAI_VAN!
    #   C34+1)&" (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ","")&IF(IF(TRUNG_GIAN!$E$7=1,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DU
    #   NG_HY_KY!$C$24:$C$28,2)<=1),"Hỏa yếu trong mệnh"&IF(AND(TRUNG_GIAN!$E$7=1,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0))," (hành nhật chủ,
    #    thân nhược)","")&"; ","")&IF((INDEX(DUNG_HY_KY!$C$24:$C$28,5)>=P_NHIEU),"Thủy vượng sẵn trong mệnh; ","")&IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$1
    #   11,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,D
    #   UNG_HY_KY!$C$111))),"vận/năm mang hành kỵ thần; ","")&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=1),"Hỏa là dụng thần; ",""),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',('NẶNG' if X.truthy(X.ge(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-98,c),4.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-97,c),4.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),1.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 2.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))) else 0.0)),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-98,c),4.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-97,c),4.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),1.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 2.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))) else 0.0)),3.0)) else 'CHÚ Ý'))),' – tim – máu, mắt: '),('can mang Thủy khắc; ' if X.truthy(X.eq(E.v(14,r-98,c),4.0)) else '')),('chi mang Thủy khắc; ' if X.truthy(X.eq(E.v(14,r-97,c),4.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Hỏa ở đất ',('Tử' if X.truthy(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),7.0)) else 'Tuyệt')),' tại '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-106,c),1.0))),' (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),9.0))) else '')),(X.cat(X.cat('Hỏa yếu trong mệnh',(' (hành nhật chủ, thân nhược)' if X.truthy(X.AND(X.eq(E.v(23,7,5),1.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else '')),'; ') if X.truthy((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),1.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 2.0),1.0))) else '')),('Thủy vượng sẵn trong mệnh; ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2))) else '')),('vận/năm mang hành kỵ thần; ' if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else '')),('Hỏa là dụng thần; ' if X.truthy(X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0)) else '')) if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(14,r-98,c),4.0), X.eq(E.v(14,r-97,c),4.0)), X.OR(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),9.0)), (X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),1.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 2.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 5.0),E.v(24,99,2)))), X.AND(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(2.0,1.0), X.add(E.v(14,r-106,c),1.0)),9.0)), X.OR(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-97,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))))) else '')

def f276(E, r, c):
    # Ô C141, D141, E141, F141, G141, H141, I141, J141
    # = IF(OR(AND(OR((DAI_VAN!C42=0),(DAI_VAN!C43=0)),OR(OR(INDEX(M_TS,4+1,DAI_VAN!C34+1)=7,INDEX(M_TS,4+1,DAI_VAN!C34+1)=9),IF(TRUNG_GIAN!$E$7=2,(IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU))),AND(OR(INDEX(M_TS,4+1,D
    #   AI_VAN!C34+1)=7,INDEX(M_TS,4+1,DAI_VAN!C34+1)=9),OR(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$10
    #   9,DUNG_HY_KY!$C$109)=2)))),"• "&IF((IF((DAI_VAN!C42=0),1,0)+IF((DAI_VAN!C43=0),1,0)+IF(OR(INDEX(M_TS,4+1,DAI_VAN!C34+1)=7,INDEX(M_TS,4+1,DAI_VAN!C34+1
    #   )=9),1,0)+IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)
    #   ,DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=2,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY
    #   _KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)
    #   =2)),1,0))>=4,"NẶNG",IF((IF((DAI_VAN!C42=0),1,0)+IF((DAI_VAN!C43=0),1,0)+IF(OR(INDEX(M_TS,4+1,DAI_VAN!C34+1)=7,INDEX(M_TS,4+1,DAI_VAN!C34+1)=9),1,0)+I
    #   F(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C4
    #   3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=2,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0
    #   ),INDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=2)),1,0))=
    #   3,"ĐỀ PHÒNG","CHÚ Ý"))&" – tì vị (lá lách, dạ dày): "&IF((DAI_VAN!C42=0),"can mang Mộc khắc; ","")&IF((DAI_VAN!C43=0),"chi mang Mộc khắc; ","")&IF(OR(
    #   INDEX(M_TS,4+1,DAI_VAN!C34+1)=7,INDEX(M_TS,4+1,DAI_VAN!C34+1)=9),"Thổ ở đất "&IF(INDEX(M_TS,4+1,DAI_VAN!C34+1)=7,"Tử","Tuyệt")&" tại "&INDEX(CHI_TEN,D
    #   AI_VAN!C34+1)&" (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ","")&IF(IF(TRUNG_GIAN!$E$7=2,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),I
    #   NDEX(DUNG_HY_KY!$C$24:$C$28,3)<=1),"Thổ yếu trong mệnh"&IF(AND(TRUNG_GIAN!$E$7=2,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0))," (hành nh
    #   ật chủ, thân nhược)","")&"; ","")&IF((INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU),"Mộc vượng sẵn trong mệnh; ","")&IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_K
    #   Y!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D
    #   $111,DUNG_HY_KY!$C$111))),"vận/năm mang hành kỵ thần; ","")&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=2),"Thổ là dụng thần; ",""),
    #   "")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',('NẶNG' if X.truthy(X.ge(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-99,c),0.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-98,c),0.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),2.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),2.0))) else 0.0)),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-99,c),0.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-98,c),0.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),2.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),2.0))) else 0.0)),3.0)) else 'CHÚ Ý'))),' – tì vị (lá lách, dạ dày): '),('can mang Mộc khắc; ' if X.truthy(X.eq(E.v(14,r-99,c),0.0)) else '')),('chi mang Mộc khắc; ' if X.truthy(X.eq(E.v(14,r-98,c),0.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Thổ ở đất ',('Tử' if X.truthy(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),7.0)) else 'Tuyệt')),' tại '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-107,c),1.0))),' (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),9.0))) else '')),(X.cat(X.cat('Thổ yếu trong mệnh',(' (hành nhật chủ, thân nhược)' if X.truthy(X.AND(X.eq(E.v(23,7,5),2.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else '')),'; ') if X.truthy((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),2.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0))) else '')),('Mộc vượng sẵn trong mệnh; ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2))) else '')),('vận/năm mang hành kỵ thần; ' if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else '')),('Thổ là dụng thần; ' if X.truthy(X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),2.0)) else '')) if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(14,r-99,c),0.0), X.eq(E.v(14,r-98,c),0.0)), X.OR(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),9.0)), (X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),2.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 3.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)))), X.AND(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(4.0,1.0), X.add(E.v(14,r-107,c),1.0)),9.0)), X.OR(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-98,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),2.0))))) else '')

def f277(E, r, c):
    # Ô C142, D142, E142, F142, G142, H142, I142, J142
    # = IF(OR(AND(OR((DAI_VAN!C42=1),(DAI_VAN!C43=1)),OR(OR(INDEX(M_TS,6+1,DAI_VAN!C34+1)=7,INDEX(M_TS,6+1,DAI_VAN!C34+1)=9),IF(TRUNG_GIAN!$E$7=3,(IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,4)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU))),AND(OR(INDEX(M_TS,6+1,D
    #   AI_VAN!C34+1)=7,INDEX(M_TS,6+1,DAI_VAN!C34+1)=9),OR(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$10
    #   9,DUNG_HY_KY!$C$109)=3)))),"• "&IF((IF((DAI_VAN!C42=1),1,0)+IF((DAI_VAN!C43=1),1,0)+IF(OR(INDEX(M_TS,6+1,DAI_VAN!C34+1)=7,INDEX(M_TS,6+1,DAI_VAN!C34+1
    #   )=9),1,0)+IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)
    #   ,DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=3,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY
    #   _KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,4)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)
    #   =3)),1,0))>=4,"NẶNG",IF((IF((DAI_VAN!C42=1),1,0)+IF((DAI_VAN!C43=1),1,0)+IF(OR(INDEX(M_TS,6+1,DAI_VAN!C34+1)=7,INDEX(M_TS,6+1,DAI_VAN!C34+1)=9),1,0)+I
    #   F(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C4
    #   3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=3,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0
    #   ),INDEX(DUNG_HY_KY!$C$24:$C$28,4)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=3)),1,0))=
    #   3,"ĐỀ PHÒNG","CHÚ Ý"))&" – phổi – đại tràng: "&IF((DAI_VAN!C42=1),"can mang Hỏa khắc; ","")&IF((DAI_VAN!C43=1),"chi mang Hỏa khắc; ","")&IF(OR(INDEX(M
    #   _TS,6+1,DAI_VAN!C34+1)=7,INDEX(M_TS,6+1,DAI_VAN!C34+1)=9),"Kim ở đất "&IF(INDEX(M_TS,6+1,DAI_VAN!C34+1)=7,"Tử","Tuyệt")&" tại "&INDEX(CHI_TEN,DAI_VAN!
    #   C34+1)&" (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ","")&IF(IF(TRUNG_GIAN!$E$7=3,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DU
    #   NG_HY_KY!$C$24:$C$28,4)<=1),"Kim yếu trong mệnh"&IF(AND(TRUNG_GIAN!$E$7=3,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0))," (hành nhật chủ,
    #    thân nhược)","")&"; ","")&IF((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),"Hỏa vượng sẵn trong mệnh; ","")&IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$11
    #   1,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DU
    #   NG_HY_KY!$C$111))),"vận/năm mang hành kỵ thần; ","")&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=3),"Kim là dụng thần; ",""),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',('NẶNG' if X.truthy(X.ge(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-100,c),1.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-99,c),1.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),3.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 4.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),3.0))) else 0.0)),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-100,c),1.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-99,c),1.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),3.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 4.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),3.0))) else 0.0)),3.0)) else 'CHÚ Ý'))),' – phổi – đại tràng: '),('can mang Hỏa khắc; ' if X.truthy(X.eq(E.v(14,r-100,c),1.0)) else '')),('chi mang Hỏa khắc; ' if X.truthy(X.eq(E.v(14,r-99,c),1.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Kim ở đất ',('Tử' if X.truthy(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),7.0)) else 'Tuyệt')),' tại '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-108,c),1.0))),' (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),9.0))) else '')),(X.cat(X.cat('Kim yếu trong mệnh',(' (hành nhật chủ, thân nhược)' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else '')),'; ') if X.truthy((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),3.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 4.0),1.0))) else '')),('Hỏa vượng sẵn trong mệnh; ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2))) else '')),('vận/năm mang hành kỵ thần; ' if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else '')),('Kim là dụng thần; ' if X.truthy(X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),3.0)) else '')) if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(14,r-100,c),1.0), X.eq(E.v(14,r-99,c),1.0)), X.OR(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),9.0)), (X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),3.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 4.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)))), X.AND(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(6.0,1.0), X.add(E.v(14,r-108,c),1.0)),9.0)), X.OR(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-99,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),3.0))))) else '')

def f278(E, r, c):
    # Ô C143, D143, E143, F143, G143, H143, I143, J143
    # = IF(OR(AND(OR((DAI_VAN!C42=2),(DAI_VAN!C43=2)),OR(OR(INDEX(M_TS,8+1,DAI_VAN!C34+1)=7,INDEX(M_TS,8+1,DAI_VAN!C34+1)=9),IF(TRUNG_GIAN!$E$7=4,(IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU))),AND(OR(INDEX(M_TS,8+1,D
    #   AI_VAN!C34+1)=7,INDEX(M_TS,8+1,DAI_VAN!C34+1)=9),OR(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33
    #   =2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$10
    #   9,DUNG_HY_KY!$C$109)=4)))),"• "&IF((IF((DAI_VAN!C42=2),1,0)+IF((DAI_VAN!C43=2),1,0)+IF(OR(INDEX(M_TS,8+1,DAI_VAN!C34+1)=7,INDEX(M_TS,8+1,DAI_VAN!C34+1
    #   )=9),1,0)+IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)
    #   ,DAI_VAN!C43=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY
    #   _KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)
    #   =4)),1,0))>=4,"NẶNG",IF((IF((DAI_VAN!C42=2),1,0)+IF((DAI_VAN!C43=2),1,0)+IF(OR(INDEX(M_TS,8+1,DAI_VAN!C34+1)=7,INDEX(M_TS,8+1,DAI_VAN!C34+1)=9),1,0)+I
    #   F(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C4
    #   3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),1,0)+IF(OR(IF(TRUNG_GIAN!$E$7=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0
    #   ),INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=4)),1,0))=
    #   3,"ĐỀ PHÒNG","CHÚ Ý"))&" – thận – bàng quang, vùng bụng, chi dưới: "&IF((DAI_VAN!C42=2),"can mang Thổ khắc; ","")&IF((DAI_VAN!C43=2),"chi mang Thổ khắ
    #   c; ","")&IF(OR(INDEX(M_TS,8+1,DAI_VAN!C34+1)=7,INDEX(M_TS,8+1,DAI_VAN!C34+1)=9),"Thủy ở đất "&IF(INDEX(M_TS,8+1,DAI_VAN!C34+1)=7,"Tử","Tuyệt")&" tại "
    #   &INDEX(CHI_TEN,DAI_VAN!C34+1)&" (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ","")&IF(IF(TRUNG_GIAN!$E$7=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_
    #   HY_KY!$C$7)=0),INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1),"Thủy yếu trong mệnh"&IF(AND(TRUNG_GIAN!$E$7=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$
    #   7)=0))," (hành nhật chủ, thân nhược)","")&"; ","")&IF((INDEX(DUNG_HY_KY!$C$24:$C$28,3)>=P_NHIEU),"Thổ vượng sẵn trong mệnh; ","")&IF(AND(IF(DOI_CHIEU!
    #   $C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)>=0,OR(DAI_VAN!C42=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),DAI_VAN!C43=IF(DOI_CHIEU!$C$3
    #   3=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111))),"vận/năm mang hành kỵ thần; ","")&IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=4),"Thủy là
    #    dụng thần; ",""),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',('NẶNG' if X.truthy(X.ge(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-101,c),2.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-100,c),2.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-101,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),4.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),4.0))) else 0.0)),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(E.v(14,r-101,c),2.0)) else 0.0),(1.0 if X.truthy(X.eq(E.v(14,r-100,c),2.0)) else 0.0)),(1.0 if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),9.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-101,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else 0.0)),(1.0 if X.truthy(X.OR((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),4.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),4.0))) else 0.0)),3.0)) else 'CHÚ Ý'))),' – thận – bàng quang, vùng bụng, chi dưới: '),('can mang Thổ khắc; ' if X.truthy(X.eq(E.v(14,r-101,c),2.0)) else '')),('chi mang Thổ khắc; ' if X.truthy(X.eq(E.v(14,r-100,c),2.0)) else '')),(X.cat(X.cat(X.cat(X.cat('Thủy ở đất ',('Tử' if X.truthy(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),7.0)) else 'Tuyệt')),' tại '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,r-109,c),1.0))),' (như ví dụ sách “mộc tử ở ngọ” – P2 tr.561); ') if X.truthy(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),9.0))) else '')),(X.cat(X.cat('Thủy yếu trong mệnh',(' (hành nhật chủ, thân nhược)' if X.truthy(X.AND(X.eq(E.v(23,7,5),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else '')),'; ') if X.truthy((X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),4.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0))) else '')),('Thổ vượng sẵn trong mệnh; ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2))) else '')),('vận/năm mang hành kỵ thần; ' if X.truthy(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-101,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else '')),('Thủy là dụng thần; ' if X.truthy(X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),4.0)) else '')) if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(14,r-101,c),2.0), X.eq(E.v(14,r-100,c),2.0)), X.OR(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),9.0)), (X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0) if X.truthy(X.eq(E.v(23,7,5),4.0)) else X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 3.0),E.v(24,99,2)))), X.AND(X.OR(X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),7.0), X.eq(X.INDEX(E.rg(24,5,16,14,27), X.add(8.0,1.0), X.add(E.v(14,r-109,c),1.0)),9.0)), X.OR(X.AND(X.ge((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),0.0), X.OR(X.eq(E.v(14,r-101,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.eq(E.v(14,r-100,c),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),4.0))))) else '')

def f279(E, r, c):
    # Ô C144, D144, E144, F144, G144, H144, I144, J144
    # = TRIM(IF(AND(OR(TRUNG_GIAN!$D$12="Thất sát",TRUNG_GIAN!$D$12="Thương quan",INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$D$6+1)+1)="T
    #   hất sát",INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,TRUNG_GIAN!$D$6+1)+1)="Thương quan"),OR(DAI_VAN!C39="Thương quan",C24="Thương quan")),"⚠ 
    #   Trụ tháng có sát/thương quan, hành vận thương quan: ""phải đề phòng bệnh mắt"". ","")&IF(AND((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)+
    #   COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0))>=P_NHIEU,NOT((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0)),OR(DAI_VAN!C39="Chính q
    #   uan",C24="Chính quan")),"⚠ Thương sát nhiều, không có quan, vận vào quan địa: năm gặp tuế quân ""nếu không bị bệnh về mắt thì sẽ gặp điều rủi ro khác"
    #   ". ",""))
    return X.TRIM(X.cat(('⚠ Trụ tháng có sát/thương quan, hành vận thương quan: "phải đề phòng bệnh mắt". ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,12,4),'Thất sát'), X.eq(E.v(23,12,4),'Thương quan'), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,4),1.0)),1.0)),'Thất sát'), X.eq(X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(23,6,4),1.0)),1.0)),'Thương quan')), X.OR(X.eq(E.v(14,r-105,c),'Thương quan'), X.eq(E.v(22,r-120,c),'Thương quan')))) else ''),('⚠ Thương sát nhiều, không có quan, vận vào quan địa: năm gặp tuế quân "nếu không bị bệnh về mắt thì sẽ gặp điều rủi ro khác". ' if X.truthy(X.AND(X.ge(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),E.v(24,99,2)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0)), X.OR(X.eq(E.v(14,r-105,c),'Chính quan'), X.eq(E.v(22,r-120,c),'Chính quan')))) else '')))

def f280(E, r, c):
    # Ô C146
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f281(E, r, c):
    # Ô D146
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f282(E, r, c):
    # Ô E146
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f283(E, r, c):
    # Ô K146
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f284(E, r, c):
    # Ô L146
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+0),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f285(E, r, c):
    # Ô C147
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f286(E, r, c):
    # Ô D147
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f287(E, r, c):
    # Ô E147
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f288(E, r, c):
    # Ô K147
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f289(E, r, c):
    # Ô L147
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+1),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f290(E, r, c):
    # Ô C148
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f291(E, r, c):
    # Ô D148
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f292(E, r, c):
    # Ô E148
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f293(E, r, c):
    # Ô K148
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f294(E, r, c):
    # Ô L148
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+2),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f295(E, r, c):
    # Ô C149
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f296(E, r, c):
    # Ô D149
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f297(E, r, c):
    # Ô E149
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f298(E, r, c):
    # Ô K149
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f299(E, r, c):
    # Ô L149
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+3),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f300(E, r, c):
    # Ô C150
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f301(E, r, c):
    # Ô D150
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f302(E, r, c):
    # Ô E150
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f303(E, r, c):
    # Ô K150
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f304(E, r, c):
    # Ô L150
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+4),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f305(E, r, c):
    # Ô C151
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f306(E, r, c):
    # Ô D151
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f307(E, r, c):
    # Ô E151
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f308(E, r, c):
    # Ô K151
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f309(E, r, c):
    # Ô L151
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+5),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f310(E, r, c):
    # Ô B152
    # = (DAI_VAN!$C$23+6)
    return X.add(E.v(14,23,3),6.0)

def f311(E, r, c):
    # Ô C152
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f312(E, r, c):
    # Ô D152
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f313(E, r, c):
    # Ô E152
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f314(E, r, c):
    # Ô K152
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f315(E, r, c):
    # Ô L152
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+6),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),6.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f316(E, r, c):
    # Ô B153
    # = (DAI_VAN!$C$23+7)
    return X.add(E.v(14,23,3),7.0)

def f317(E, r, c):
    # Ô C153
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f318(E, r, c):
    # Ô D153
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f319(E, r, c):
    # Ô E153
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f320(E, r, c):
    # Ô K153
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f321(E, r, c):
    # Ô L153
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+7),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),7.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f322(E, r, c):
    # Ô B154
    # = (DAI_VAN!$C$23+8)
    return X.add(E.v(14,23,3),8.0)

def f323(E, r, c):
    # Ô C154
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f324(E, r, c):
    # Ô D154
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f325(E, r, c):
    # Ô E154
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f326(E, r, c):
    # Ô K154
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f327(E, r, c):
    # Ô L154
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+8),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),8.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f328(E, r, c):
    # Ô B155
    # = (DAI_VAN!$C$23+9)
    return X.add(E.v(14,23,3),9.0)

def f329(E, r, c):
    # Ô C155
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f330(E, r, c):
    # Ô D155
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f331(E, r, c):
    # Ô E155
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERRO
    #   R(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f332(E, r, c):
    # Ô K155
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MAT
    #   CH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129
    #   ,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f333(E, r, c):
    # Ô L155
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+9),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụn
    #   g thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),9.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f334(E, r, c):
    # Ô B156
    # = (DAI_VAN!$C$23+10)
    return X.add(E.v(14,23,3),10.0)

def f335(E, r, c):
    # Ô C156
    # = IFERROR(INDEX(LUU_NIEN!$L$30:$L$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,12,129,12), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f336(E, r, c):
    # Ô D156
    # = IFERROR(INDEX(LUU_NIEN!$H$30:$H$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),"")
    return X.IFERROR(lambda: X.INDEX(E.rg(15,30,8,129,8), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')

def f337(E, r, c):
    # Ô E156
    # = IF(IFERROR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),"")="","Không có tạng nào bị khắc rõ trong năm này.",IFERR
    #   OR(INDEX(LUU_NIEN!$CN$30:$CN$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),""))
    return ('Không có tạng nào bị khắc rõ trong năm này.' if X.truthy(X.eq(X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''),'')) else X.IFERROR(lambda: X.INDEX(E.rg(15,30,92,129,92), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: ''))

def f338(E, r, c):
    # Ô K156
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),""))=0,"—",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MA
    #   TCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),""))>=4,"NẶNG",IF(N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$
    #   129,0)),""))=3,"ĐỀ PHÒNG","CHÚ Ý")))
    return ('—' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),0.0)) else ('NẶNG' if X.truthy(X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0)) else ('ĐỀ PHÒNG' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0)) else 'CHÚ Ý')))

def f339(E, r, c):
    # Ô L156
    # = IF(N(IFERROR(INDEX(LUU_NIEN!$CP$30:$CP$129,MATCH((DAI_VAN!$C$23+10),LUU_NIEN!$A$30:$A$129,0)),""))=1,"⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụ
    #   ng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)","")
    return ('⚠ CÓ – năm vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng: cần khám sức khỏe định kỳ, tránh rủi ro; không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,94,129,94), X.MATCH(X.add(E.v(14,23,3),10.0), E.rg(15,30,1,129,1), 0.0)), lambda: '')),1.0)) else '')

def f340(E, r, c):
    # Ô C157
    # = "• TẠNG PHỦ THEO NĂM (năm xem và 10 năm sau): năm mức NẶNG: "&IF(TRIM(IF(AND(K146<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B146,LUU_NIEN!$A$
    #   30:$A$129,0)),""))>=4),B146&" ("&C146&") ","")&IF(AND(K147<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B147,LUU_NIEN!$A$30:$A$129,0)),""))>=4),
    #   B147&" ("&C147&") ","")&IF(AND(K148<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B148,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B148&" ("&C148&") ","")
    #   &IF(AND(K149<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B149,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B149&" ("&C149&") ","")&IF(AND(K150<>"—",N(IFE
    #   RROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B150,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B150&" ("&C150&") ","")&IF(AND(K151<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO
    #   $30:$CO$129,MATCH(B151,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B151&" ("&C151&") ","")&IF(AND(K152<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B152,
    #   LUU_NIEN!$A$30:$A$129,0)),""))>=4),B152&" ("&C152&") ","")&IF(AND(K153<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B153,LUU_NIEN!$A$30:$A$129,0
    #   )),""))>=4),B153&" ("&C153&") ","")&IF(AND(K154<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B154,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B154&" ("&C
    #   154&") ","")&IF(AND(K155<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B155,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B155&" ("&C155&") ","")&IF(AND(K15
    #   6<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B156,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B156&" ("&C156&") ",""))="","không có",TRIM(IF(AND(K146<>
    #   "—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B146,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B146&" ("&C146&") ","")&IF(AND(K147<>"—",N(IFERROR(INDEX(LUU
    #   _NIEN!$CO$30:$CO$129,MATCH(B147,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B147&" ("&C147&") ","")&IF(AND(K148<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MA
    #   TCH(B148,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B148&" ("&C148&") ","")&IF(AND(K149<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B149,LUU_NIEN!$A$30
    #   :$A$129,0)),""))>=4),B149&" ("&C149&") ","")&IF(AND(K150<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B150,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B1
    #   50&" ("&C150&") ","")&IF(AND(K151<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B151,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B151&" ("&C151&") ","")&I
    #   F(AND(K152<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B152,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B152&" ("&C152&") ","")&IF(AND(K153<>"—",N(IFERR
    #   OR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B153,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B153&" ("&C153&") ","")&IF(AND(K154<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$3
    #   0:$CO$129,MATCH(B154,LUU_NIEN!$A$30:$A$129,0)),""))>=4),B154&" ("&C154&") ","")&IF(AND(K155<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B155,LU
    #   U_NIEN!$A$30:$A$129,0)),""))>=4),B155&" ("&C155&") ","")&IF(AND(K156<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B156,LUU_NIEN!$A$30:$A$129,0))
    #   ,""))>=4),B156&" ("&C156&") ","")))&"; mức ĐỀ PHÒNG trở lên: "&IF(TRIM(IF(AND(K146<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B146,LUU_NIEN!$A
    #   $30:$A$129,0)),""))>=3),B146&" ("&C146&") ","")&IF(AND(K147<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B147,LUU_NIEN!$A$30:$A$129,0)),""))>=3)
    #   ,B147&" ("&C147&") ","")&IF(AND(K148<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B148,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B148&" ("&C148&") ",""
    #   )&IF(AND(K149<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B149,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B149&" ("&C149&") ","")&IF(AND(K150<>"—",N(IF
    #   ERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B150,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B150&" ("&C150&") ","")&IF(AND(K151<>"—",N(IFERROR(INDEX(LUU_NIEN!$C
    #   O$30:$CO$129,MATCH(B151,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B151&" ("&C151&") ","")&IF(AND(K152<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B152
    #   ,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B152&" ("&C152&") ","")&IF(AND(K153<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B153,LUU_NIEN!$A$30:$A$129,
    #   0)),""))>=3),B153&" ("&C153&") ","")&IF(AND(K154<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B154,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B154&" ("&
    #   C154&") ","")&IF(AND(K155<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B155,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B155&" ("&C155&") ","")&IF(AND(K1
    #   56<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B156,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B156&" ("&C156&") ",""))="","không có",TRIM(IF(AND(K146<
    #   >"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B146,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B146&" ("&C146&") ","")&IF(AND(K147<>"—",N(IFERROR(INDEX(LU
    #   U_NIEN!$CO$30:$CO$129,MATCH(B147,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B147&" ("&C147&") ","")&IF(AND(K148<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,M
    #   ATCH(B148,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B148&" ("&C148&") ","")&IF(AND(K149<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B149,LUU_NIEN!$A$3
    #   0:$A$129,0)),""))>=3),B149&" ("&C149&") ","")&IF(AND(K150<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B150,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B
    #   150&" ("&C150&") ","")&IF(AND(K151<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B151,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B151&" ("&C151&") ","")&
    #   IF(AND(K152<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B152,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B152&" ("&C152&") ","")&IF(AND(K153<>"—",N(IFER
    #   ROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B153,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B153&" ("&C153&") ","")&IF(AND(K154<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$
    #   30:$CO$129,MATCH(B154,LUU_NIEN!$A$30:$A$129,0)),""))>=3),B154&" ("&C154&") ","")&IF(AND(K155<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B155,L
    #   UU_NIEN!$A$30:$A$129,0)),""))>=3),B155&" ("&C155&") ","")&IF(AND(K156<>"—",N(IFERROR(INDEX(LUU_NIEN!$CO$30:$CO$129,MATCH(B156,LUU_NIEN!$A$30:$A$129,0)
    #   ),""))>=3),B156&" ("&C156&") ","")))&IF(TRIM(IF(L146<>"",B146&" ","")&IF(L147<>"",B147&" ","")&IF(L148<>"",B148&" ","")&IF(L149<>"",B149&" ","")&IF(L1
    #   50<>"",B150&" ","")&IF(L151<>"",B151&" ","")&IF(L152<>"",B152&" ","")&IF(L153<>"",B153&" ","")&IF(L154<>"",B154&" ","")&IF(L155<>"",B155&" ","")&IF(L1
    #   56<>"",B156&" ",""))="","",CHAR(10)&"⚠ Năm cần đề phòng nghiêm túc (vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng – P1 tr.201; không kết l
    #   uận năm mất): "&TRIM(IF(L146<>"",B146&" ","")&IF(L147<>"",B147&" ","")&IF(L148<>"",B148&" ","")&IF(L149<>"",B149&" ","")&IF(L150<>"",B150&" ","")&IF(L
    #   151<>"",B151&" ","")&IF(L152<>"",B152&" ","")&IF(L153<>"",B153&" ","")&IF(L154<>"",B154&" ","")&IF(L155<>"",B155&" ","")&IF(L156<>"",B156&" ","")))&CH
    #   AR(10)&"Chi tiết từng năm: SUC_KHOE mục H; cả đời: LUU_NIEN cột CN. Không thay cho chẩn đoán y khoa."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• TẠNG PHỦ THEO NĂM (năm xem và 10 năm sau): năm mức NẶNG: ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(22,r-11,c-1),' ('),E.v(22,r-11,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-11,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-11,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else ''),(X.cat(X.cat(X.cat(E.v(22,r-10,c-1),' ('),E.v(22,r-10,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-10,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-10,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-9,c-1),' ('),E.v(22,r-9,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-9,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-9,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-8,c-1),' ('),E.v(22,r-8,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-8,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-8,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-7,c-1),' ('),E.v(22,r-7,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-7,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-7,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-6,c-1),' ('),E.v(22,r-6,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-6,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-6,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-5,c-1),' ('),E.v(22,r-5,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-5,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-5,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-4,c-1),' ('),E.v(22,r-4,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-4,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-4,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-3,c-1),' ('),E.v(22,r-3,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-3,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-3,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-2,c-1),' ('),E.v(22,r-2,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-2,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-2,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-1,c-1),' ('),E.v(22,r-1,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-1,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-1,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(22,r-11,c-1),' ('),E.v(22,r-11,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-11,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-11,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else ''),(X.cat(X.cat(X.cat(E.v(22,r-10,c-1),' ('),E.v(22,r-10,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-10,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-10,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-9,c-1),' ('),E.v(22,r-9,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-9,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-9,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-8,c-1),' ('),E.v(22,r-8,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-8,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-8,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-7,c-1),' ('),E.v(22,r-7,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-7,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-7,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-6,c-1),' ('),E.v(22,r-6,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-6,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-6,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-5,c-1),' ('),E.v(22,r-5,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-5,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-5,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-4,c-1),' ('),E.v(22,r-4,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-4,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-4,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-3,c-1),' ('),E.v(22,r-3,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-3,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-3,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-2,c-1),' ('),E.v(22,r-2,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-2,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-2,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-1,c-1),' ('),E.v(22,r-1,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-1,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-1,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),4.0))) else ''))))),'; mức ĐỀ PHÒNG trở lên: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(22,r-11,c-1),' ('),E.v(22,r-11,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-11,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-11,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else ''),(X.cat(X.cat(X.cat(E.v(22,r-10,c-1),' ('),E.v(22,r-10,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-10,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-10,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-9,c-1),' ('),E.v(22,r-9,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-9,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-9,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-8,c-1),' ('),E.v(22,r-8,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-8,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-8,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-7,c-1),' ('),E.v(22,r-7,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-7,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-7,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-6,c-1),' ('),E.v(22,r-6,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-6,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-6,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-5,c-1),' ('),E.v(22,r-5,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-5,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-5,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-4,c-1),' ('),E.v(22,r-4,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-4,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-4,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-3,c-1),' ('),E.v(22,r-3,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-3,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-3,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-2,c-1),' ('),E.v(22,r-2,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-2,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-2,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-1,c-1),' ('),E.v(22,r-1,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-1,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-1,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(E.v(22,r-11,c-1),' ('),E.v(22,r-11,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-11,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-11,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else ''),(X.cat(X.cat(X.cat(E.v(22,r-10,c-1),' ('),E.v(22,r-10,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-10,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-10,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-9,c-1),' ('),E.v(22,r-9,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-9,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-9,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-8,c-1),' ('),E.v(22,r-8,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-8,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-8,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-7,c-1),' ('),E.v(22,r-7,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-7,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-7,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-6,c-1),' ('),E.v(22,r-6,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-6,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-6,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-5,c-1),' ('),E.v(22,r-5,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-5,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-5,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-4,c-1),' ('),E.v(22,r-4,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-4,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-4,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-3,c-1),' ('),E.v(22,r-3,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-3,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-3,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-2,c-1),' ('),E.v(22,r-2,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-2,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-2,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else '')),(X.cat(X.cat(X.cat(E.v(22,r-1,c-1),' ('),E.v(22,r-1,c)),') ') if X.truthy(X.AND(X.ne(E.v(22,r-1,c+8),'—'), X.ge(X.N(X.IFERROR(lambda: X.INDEX(E.rg(15,30,93,129,93), X.MATCH(E.v(22,r-1,c-1), E.rg(15,30,1,129,1), 0.0)), lambda: '')),3.0))) else ''))))),('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-11,c-1),' ') if X.truthy(X.ne(E.v(22,r-11,c+9),'')) else ''),(X.cat(E.v(22,r-10,c-1),' ') if X.truthy(X.ne(E.v(22,r-10,c+9),'')) else '')),(X.cat(E.v(22,r-9,c-1),' ') if X.truthy(X.ne(E.v(22,r-9,c+9),'')) else '')),(X.cat(E.v(22,r-8,c-1),' ') if X.truthy(X.ne(E.v(22,r-8,c+9),'')) else '')),(X.cat(E.v(22,r-7,c-1),' ') if X.truthy(X.ne(E.v(22,r-7,c+9),'')) else '')),(X.cat(E.v(22,r-6,c-1),' ') if X.truthy(X.ne(E.v(22,r-6,c+9),'')) else '')),(X.cat(E.v(22,r-5,c-1),' ') if X.truthy(X.ne(E.v(22,r-5,c+9),'')) else '')),(X.cat(E.v(22,r-4,c-1),' ') if X.truthy(X.ne(E.v(22,r-4,c+9),'')) else '')),(X.cat(E.v(22,r-3,c-1),' ') if X.truthy(X.ne(E.v(22,r-3,c+9),'')) else '')),(X.cat(E.v(22,r-2,c-1),' ') if X.truthy(X.ne(E.v(22,r-2,c+9),'')) else '')),(X.cat(E.v(22,r-1,c-1),' ') if X.truthy(X.ne(E.v(22,r-1,c+9),'')) else ''))),'')) else X.cat(X.cat(X.CHAR(10.0),'⚠ Năm cần đề phòng nghiêm túc (vận kỵ + năm kỵ, tạng của nhật chủ/dụng thần bị khắc nặng – P1 tr.201; không kết luận năm mất): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(22,r-11,c-1),' ') if X.truthy(X.ne(E.v(22,r-11,c+9),'')) else ''),(X.cat(E.v(22,r-10,c-1),' ') if X.truthy(X.ne(E.v(22,r-10,c+9),'')) else '')),(X.cat(E.v(22,r-9,c-1),' ') if X.truthy(X.ne(E.v(22,r-9,c+9),'')) else '')),(X.cat(E.v(22,r-8,c-1),' ') if X.truthy(X.ne(E.v(22,r-8,c+9),'')) else '')),(X.cat(E.v(22,r-7,c-1),' ') if X.truthy(X.ne(E.v(22,r-7,c+9),'')) else '')),(X.cat(E.v(22,r-6,c-1),' ') if X.truthy(X.ne(E.v(22,r-6,c+9),'')) else '')),(X.cat(E.v(22,r-5,c-1),' ') if X.truthy(X.ne(E.v(22,r-5,c+9),'')) else '')),(X.cat(E.v(22,r-4,c-1),' ') if X.truthy(X.ne(E.v(22,r-4,c+9),'')) else '')),(X.cat(E.v(22,r-3,c-1),' ') if X.truthy(X.ne(E.v(22,r-3,c+9),'')) else '')),(X.cat(E.v(22,r-2,c-1),' ') if X.truthy(X.ne(E.v(22,r-2,c+9),'')) else '')),(X.cat(E.v(22,r-1,c-1),' ') if X.truthy(X.ne(E.v(22,r-1,c+9),'')) else '')))))),X.CHAR(10.0)),'Chi tiết từng năm: SUC_KHOE mục H; cả đời: LUU_NIEN cột CN. Không thay cho chẩn đoán y khoa.')

def f341(E, r, c):
    # Ô C160
    # = "Năm "&(DAI_VAN!$C$23+0)&" ("&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+0)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+0)-4,12)+1)&"): "&INDEX({"Nhất bạch","
    #   Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1))&" nhập cung Giữa"&"; Bắc 
    #   (Khảm): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),
    #   9),9)+1)-1+5,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","ha
    #   o tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1))&")"&"; Đông Bắc (Cấn): "&INDEX({"Nhất bạch","
    #   Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1))&" ("&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1))&")"&"; Đông (Chấn): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ
    #    hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tậ
    #   t","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-
    #   MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1))&")"&"; Đông Nam (Tốn): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích"
    #   ,"Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn x
    #   ương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)
    #   -1+8,9)+1))&")"&"; Nam (Ly): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-M
    #   OD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách",
    #   "tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1))&")"&"; Tây Nam (Khôn): "
    #   &INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-
    #   1+6,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mấ
    #   t mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1))&")"&"; Tây (Đoài): "&INDEX({"Nhất bạch","Nhị hắc","Tam
    #    bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1))&" ("&INDEX({"tốt – quý 
    #   nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ k
    #   hánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1))&")"&"; Tây Bắc (Càn): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lụ
    #   c bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi
    #   , tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+1,9)+1))&")"&". Sách: Ngũ hoàng, Nhị hắc là sao tai ách, bệnh tật – giường, phòng ngủ, cửa chính kỵ các phương này; Lục bạch nên
    #    sinh hoạt nhiều. Ý nghĩa các sao khác theo kinh nghiệm phong thủy phổ biến. Sao đổi tại Lập xuân."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',X.add(E.v(14,23,3),0.0)),' ('),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0))),'): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0))),' nhập cung Giữa'),'; Bắc (Khảm): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),')'),'; Đông Bắc (Cấn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),')'),'; Đông (Chấn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),')'),'; Đông Nam (Tốn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),')'),'; Nam (Ly): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),')'),'; Tây Nam (Khôn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),')'),'; Tây (Đoài): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),')'),'; Tây Bắc (Càn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),')'),'. Sách: Ngũ hoàng, Nhị hắc là sao tai ách, bệnh tật – giường, phòng ngủ, cửa chính kỵ các phương này; Lục bạch nên sinh hoạt nhiều. Ý nghĩa các sao khác theo kinh nghiệm phong thủy phổ biến. Sao đổi tại Lập xuân.')

def f342(E, r, c):
    # Ô C161
    # = "Năm "&(DAI_VAN!$C$23+1)&" ("&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+1)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+1)-4,12)+1)&"): "&INDEX({"Nhất bạch","
    #   Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1))&" nhập cung Giữa"&"; Bắc 
    #   (Khảm): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),
    #   9),9)+1)-1+5,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","ha
    #   o tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1))&")"&"; Đông Bắc (Cấn): "&INDEX({"Nhất bạch","
    #   Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1))&" ("&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1))&")"&"; Đông (Chấn): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ
    #    hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tậ
    #   t","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-
    #   MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1))&")"&"; Đông Nam (Tốn): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích"
    #   ,"Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn x
    #   ương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)
    #   -1+8,9)+1))&")"&"; Nam (Ly): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-M
    #   OD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách",
    #   "tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1))&")"&"; Tây Nam (Khôn): "
    #   &INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-
    #   1+6,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mấ
    #   t mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1))&")"&"; Tây (Đoài): "&INDEX({"Nhất bạch","Nhị hắc","Tam
    #    bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1))&" ("&INDEX({"tốt – quý 
    #   nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ k
    #   hánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1))&")"&"; Tây Bắc (Càn): "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lụ
    #   c bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1))&" ("&INDEX({"tốt – quý nhân","XẤU – bệnh tật","thị phi
    #   , tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+1,9)+1))&")"&". Sách: Ngũ hoàng, Nhị hắc là sao tai ách, bệnh tật – giường, phòng ngủ, cửa chính kỵ các phương này; Lục bạch nên
    #    sinh hoạt nhiều. Ý nghĩa các sao khác theo kinh nghiệm phong thủy phổ biến. Sao đổi tại Lập xuân."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',X.add(E.v(14,23,3),1.0)),' ('),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0))),'): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0))),' nhập cung Giữa'),'; Bắc (Khảm): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),')'),'; Đông Bắc (Cấn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),')'),'; Đông (Chấn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),')'),'; Đông Nam (Tốn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),')'),'; Nam (Ly): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),')'),'; Tây Nam (Khôn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),')'),'; Tây (Đoài): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),')'),'; Tây Bắc (Càn): '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),' ('),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),')'),'. Sách: Ngũ hoàng, Nhị hắc là sao tai ách, bệnh tật – giường, phòng ngủ, cửa chính kỵ các phương này; Lục bạch nên sinh hoạt nhiều. Ý nghĩa các sao khác theo kinh nghiệm phong thủy phổ biến. Sao đổi tại Lập xuân.')

def f343(E, r, c):
    # Ô C162
    # = "Năm "&(DAI_VAN!$C$23+0)&" ("&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+0)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+0)-4,12)+1)&"). ĐẶT GIƯỜNG, PHÒNG NGỦ,
    #    NƠI LÀM VIỆC NÊN Ở: "&IF(AL189&AL190&AL191&AL192&AL193&AL194&AL195&AL196="","(không phương dụng/hỷ nào sạch hung sát – dùng phương phụ dưới)",AL189&A
    #   L190&AL191&AL192&AL193&AL194&AL195&AL196)&CHAR(10)&"Phương phụ (hành trung tính, có sao tốt): "&IF(AM189&AM190&AM191&AM192&AM193&AM194&AM195&AM196="",
    #   "—",AM189&AM190&AM191&AM192&AM193&AM194&AM195&AM196)&CHAR(10)&"TRÁNH: "&IF(AK189&AK190&AK191&AK192&AK193&AK194&AK195&AK196="","—",AK189&AK190&AK191&AK
    #   192&AK193&AK194&AK195&AK196)&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD
    #   ((DAI_VAN!$C$23+0)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$
    #   9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"giờ ",""))<>"",CHAR(10)&"⚠ Tứ trụ có chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6
    #   =MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRU
    #   NG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+0)-4,
    #   12)+6,12)),"giờ ","")&"xung thái tuế → KỴ NHẤT phương tuế phá.","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD((DAI_VAN!$C$23+0)-4,12)),"năm ","")&IF(AND(T
    #   RUE(),TRUNG_GIAN!$D$6=MOD((DAI_VAN!$C$23+0)-4,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD((DAI_VAN!$C$23+0)-4,12)),"ngày ","")&IF(AND(NOT(NHAP
    #   _LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD((DAI_VAN!$C$23+0)-4,12)),"giờ ",""))<>"",CHAR(10)&"⚠ Chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD((DAI_V
    #   AN!$C$23+0)-4,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD((DAI_VAN!$C$23+0)-4,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD((DAI_VAN!$C$23
    #   +0)-4,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD((DAI_VAN!$C$23+0)-4,12)),"giờ ","")&"trùng thái tuế (phục ngâm) →
    #    KỴ NHẤT phương thái tuế; “nên di chuyển” (P2 tr.659).","")&CHAR(10)&"Căn cứ: tránh phương kỵ thần, đi theo phương hỷ, dụng thần (P2 tr.631); thái tuế
    #   , tuế phá, tam sát (P2 tr.632-633); cửu tinh (P2 tr.634-638); “đổi chỗ giường nằm là tiện nhất” (P2 tr.632)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',X.add(E.v(14,23,3),0.0)),' ('),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0))),'). ĐẶT GIƯỜNG, PHÒNG NGỦ, NƠI LÀM VIỆC NÊN Ở: '),('(không phương dụng/hỷ nào sạch hung sát – dùng phương phụ dưới)' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+27,c+35),E.v(22,r+28,c+35)),E.v(22,r+29,c+35)),E.v(22,r+30,c+35)),E.v(22,r+31,c+35)),E.v(22,r+32,c+35)),E.v(22,r+33,c+35)),E.v(22,r+34,c+35)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+27,c+35),E.v(22,r+28,c+35)),E.v(22,r+29,c+35)),E.v(22,r+30,c+35)),E.v(22,r+31,c+35)),E.v(22,r+32,c+35)),E.v(22,r+33,c+35)),E.v(22,r+34,c+35)))),X.CHAR(10.0)),'Phương phụ (hành trung tính, có sao tốt): '),('—' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+27,c+36),E.v(22,r+28,c+36)),E.v(22,r+29,c+36)),E.v(22,r+30,c+36)),E.v(22,r+31,c+36)),E.v(22,r+32,c+36)),E.v(22,r+33,c+36)),E.v(22,r+34,c+36)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+27,c+36),E.v(22,r+28,c+36)),E.v(22,r+29,c+36)),E.v(22,r+30,c+36)),E.v(22,r+31,c+36)),E.v(22,r+32,c+36)),E.v(22,r+33,c+36)),E.v(22,r+34,c+36)))),X.CHAR(10.0)),'TRÁNH: '),('—' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+27,c+34),E.v(22,r+28,c+34)),E.v(22,r+29,c+34)),E.v(22,r+30,c+34)),E.v(22,r+31,c+34)),E.v(22,r+32,c+34)),E.v(22,r+33,c+34)),E.v(22,r+34,c+34)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+27,c+34),E.v(22,r+28,c+34)),E.v(22,r+29,c+34)),E.v(22,r+30,c+34)),E.v(22,r+31,c+34)),E.v(22,r+32,c+34)),E.v(22,r+33,c+34)),E.v(22,r+34,c+34)))),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Tứ trụ có chi trụ '),('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),'xung thái tuế → KỴ NHẤT phương tuế phá.') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Chi trụ '),('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),'trùng thái tuế (phục ngâm) → KỴ NHẤT phương thái tuế; “nên di chuyển” (P2 tr.659).') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0)))) else '')),'')) else '')),X.CHAR(10.0)),'Căn cứ: tránh phương kỵ thần, đi theo phương hỷ, dụng thần (P2 tr.631); thái tuế, tuế phá, tam sát (P2 tr.632-633); cửu tinh (P2 tr.634-638); “đổi chỗ giường nằm là tiện nhất” (P2 tr.632).')

def f344(E, r, c):
    # Ô C163
    # = "Năm "&(DAI_VAN!$C$23+1)&" ("&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+1)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+1)-4,12)+1)&"). ĐẶT GIƯỜNG, PHÒNG NGỦ,
    #    NƠI LÀM VIỆC NÊN Ở: "&IF(AL197&AL198&AL199&AL200&AL201&AL202&AL203&AL204="","(không phương dụng/hỷ nào sạch hung sát – dùng phương phụ dưới)",AL197&A
    #   L198&AL199&AL200&AL201&AL202&AL203&AL204)&CHAR(10)&"Phương phụ (hành trung tính, có sao tốt): "&IF(AM197&AM198&AM199&AM200&AM201&AM202&AM203&AM204="",
    #   "—",AM197&AM198&AM199&AM200&AM201&AM202&AM203&AM204)&CHAR(10)&"TRÁNH: "&IF(AK197&AK198&AK199&AK200&AK201&AK202&AK203&AK204="","—",AK197&AK198&AK199&AK
    #   200&AK201&AK202&AK203&AK204)&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD
    #   ((DAI_VAN!$C$23+1)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$
    #   9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"giờ ",""))<>"",CHAR(10)&"⚠ Tứ trụ có chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6
    #   =MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRU
    #   NG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+1)-4,
    #   12)+6,12)),"giờ ","")&"xung thái tuế → KỴ NHẤT phương tuế phá.","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD((DAI_VAN!$C$23+1)-4,12)),"năm ","")&IF(AND(T
    #   RUE(),TRUNG_GIAN!$D$6=MOD((DAI_VAN!$C$23+1)-4,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD((DAI_VAN!$C$23+1)-4,12)),"ngày ","")&IF(AND(NOT(NHAP
    #   _LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD((DAI_VAN!$C$23+1)-4,12)),"giờ ",""))<>"",CHAR(10)&"⚠ Chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD((DAI_V
    #   AN!$C$23+1)-4,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD((DAI_VAN!$C$23+1)-4,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD((DAI_VAN!$C$23
    #   +1)-4,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD((DAI_VAN!$C$23+1)-4,12)),"giờ ","")&"trùng thái tuế (phục ngâm) →
    #    KỴ NHẤT phương thái tuế; “nên di chuyển” (P2 tr.659).","")&CHAR(10)&"Căn cứ: tránh phương kỵ thần, đi theo phương hỷ, dụng thần (P2 tr.631); thái tuế
    #   , tuế phá, tam sát (P2 tr.632-633); cửu tinh (P2 tr.634-638); “đổi chỗ giường nằm là tiện nhất” (P2 tr.632)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm ',X.add(E.v(14,23,3),1.0)),' ('),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0))),'). ĐẶT GIƯỜNG, PHÒNG NGỦ, NƠI LÀM VIỆC NÊN Ở: '),('(không phương dụng/hỷ nào sạch hung sát – dùng phương phụ dưới)' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+34,c+35),E.v(22,r+35,c+35)),E.v(22,r+36,c+35)),E.v(22,r+37,c+35)),E.v(22,r+38,c+35)),E.v(22,r+39,c+35)),E.v(22,r+40,c+35)),E.v(22,r+41,c+35)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+34,c+35),E.v(22,r+35,c+35)),E.v(22,r+36,c+35)),E.v(22,r+37,c+35)),E.v(22,r+38,c+35)),E.v(22,r+39,c+35)),E.v(22,r+40,c+35)),E.v(22,r+41,c+35)))),X.CHAR(10.0)),'Phương phụ (hành trung tính, có sao tốt): '),('—' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+34,c+36),E.v(22,r+35,c+36)),E.v(22,r+36,c+36)),E.v(22,r+37,c+36)),E.v(22,r+38,c+36)),E.v(22,r+39,c+36)),E.v(22,r+40,c+36)),E.v(22,r+41,c+36)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+34,c+36),E.v(22,r+35,c+36)),E.v(22,r+36,c+36)),E.v(22,r+37,c+36)),E.v(22,r+38,c+36)),E.v(22,r+39,c+36)),E.v(22,r+40,c+36)),E.v(22,r+41,c+36)))),X.CHAR(10.0)),'TRÁNH: '),('—' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+34,c+34),E.v(22,r+35,c+34)),E.v(22,r+36,c+34)),E.v(22,r+37,c+34)),E.v(22,r+38,c+34)),E.v(22,r+39,c+34)),E.v(22,r+40,c+34)),E.v(22,r+41,c+34)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+34,c+34),E.v(22,r+35,c+34)),E.v(22,r+36,c+34)),E.v(22,r+37,c+34)),E.v(22,r+38,c+34)),E.v(22,r+39,c+34)),E.v(22,r+40,c+34)),E.v(22,r+41,c+34)))),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Tứ trụ có chi trụ '),('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),'xung thái tuế → KỴ NHẤT phương tuế phá.') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Chi trụ '),('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),'trùng thái tuế (phục ngâm) → KỴ NHẤT phương thái tuế; “nên di chuyển” (P2 tr.659).') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0)))) else '')),'')) else '')),X.CHAR(10.0)),'Căn cứ: tránh phương kỵ thần, đi theo phương hỷ, dụng thần (P2 tr.631); thái tuế, tuế phá, tam sát (P2 tr.632-633); cửu tinh (P2 tr.634-638); “đổi chỗ giường nằm là tiện nhất” (P2 tr.632).')

def f345(E, r, c):
    # Ô C164
    # = AK209&AK210&AK211&AK212&AK213&AK214&"(Vượng suy theo phương pháp khuyến nghị. Công cụ KHÔNG ghi tên thuốc – sách liệt kê vị thuốc theo tạng ở P2 tr.64
    #   2-646, chỉ dùng khi thầy thuốc chẩn đoán và kê đơn.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,r+45,c+34),E.v(22,r+46,c+34)),E.v(22,r+47,c+34)),E.v(22,r+48,c+34)),E.v(22,r+49,c+34)),E.v(22,r+50,c+34)),'(Vượng suy theo phương pháp khuyến nghị. Công cụ KHÔNG ghi tên thuốc – sách liệt kê vị thuốc theo tạng ở P2 tr.642-646, chỉ dùng khi thầy thuốc chẩn đoán và kê đơn.)')

def f346(E, r, c):
    # Ô AJ189
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=0,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=0,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=0,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   5,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),0.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),0.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),0.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),2.0)) else ''))

def f347(E, r, c):
    # Ô AK189, AK197
    # = IF(AJ189<>"","Bắc (Khảm) ("&LEFT(AJ189,LEN(AJ189)-2)&"); ","")
    return (X.cat(X.cat('Bắc (Khảm) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f348(E, r, c):
    # Ô AL189
    # = IF(AND(AJ189="",OR(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Bắc (Khả
    #   m) – "&INDEX(HANH_TEN,MOD(4,5)+1)&IF(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị h
    #   ắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1))&" – "&INDEX({"t
    #   ốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","t
    #   ốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bắc (Khảm) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(4.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(4.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f349(E, r, c):
    # Ô AM189
    # = IF(AND(AJ189="",4<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),4<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),4<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+5,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+5,9)+1)=9)),"Bắc (Khảm) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","
    #   Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+5,9)+1))&"; ","")
    return (X.cat(X.cat('Bắc (Khảm) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(4.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(4.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),9.0)))) else '')

def f350(E, r, c):
    # Ô AJ190
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=1,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=1,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=1,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   3,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),1.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),1.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),1.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),2.0)) else ''))

def f351(E, r, c):
    # Ô AK190, AK198
    # = IF(AJ190<>"","Đông Bắc (Cấn) ("&LEFT(AJ190,LEN(AJ190)-2)&"); ","")
    return (X.cat(X.cat('Đông Bắc (Cấn) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f352(E, r, c):
    # Ô AL190
    # = IF(AND(AJ190="",OR(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Đông Bắc
    #    (Cấn) – "&INDEX(HANH_TEN,MOD(2,5)+1)&IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","N
    #   hị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1))&" – "&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đông Bắc (Cấn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(2.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f353(E, r, c):
    # Ô AM190
    # = IF(AND(AJ190="",2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+3,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+3,9)+1)=9)),"Đông Bắc (Cấn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạc
    #   h","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+3,9)+1))&"; ","")
    return (X.cat(X.cat('Đông Bắc (Cấn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(2.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),9.0)))) else '')

def f354(E, r, c):
    # Ô AJ191
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=2,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=2,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=2,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   7,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),2.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),2.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),2.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),2.0)) else ''))

def f355(E, r, c):
    # Ô AK191, AK199
    # = IF(AJ191<>"","Đông (Chấn) ("&LEFT(AJ191,LEN(AJ191)-2)&"); ","")
    return (X.cat(X.cat('Đông (Chấn) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f356(E, r, c):
    # Ô AL191
    # = IF(AND(AJ191="",OR(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Đông (Ch
    #   ấn) – "&INDEX(HANH_TEN,MOD(0,5)+1)&IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị 
    #   hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1))&" – "&INDEX({"
    #   tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","
    #   tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đông (Chấn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(0.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f357(E, r, c):
    # Ô AM191
    # = IF(AND(AJ191="",0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+7,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+7,9)+1)=9)),"Đông (Chấn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch",
    #   "Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+7,9)+1))&"; ","")
    return (X.cat(X.cat('Đông (Chấn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(0.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),9.0)))) else '')

def f358(E, r, c):
    # Ô AJ192
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=3,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=3,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=3,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   8,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),3.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),3.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),3.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),2.0)) else ''))

def f359(E, r, c):
    # Ô AK192, AK200
    # = IF(AJ192<>"","Đông Nam (Tốn) ("&LEFT(AJ192,LEN(AJ192)-2)&"); ","")
    return (X.cat(X.cat('Đông Nam (Tốn) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f360(E, r, c):
    # Ô AL192
    # = IF(AND(AJ192="",OR(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Đông Nam
    #    (Tốn) – "&INDEX(HANH_TEN,MOD(0,5)+1)&IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","N
    #   hị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1))&" – "&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đông Nam (Tốn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(0.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f361(E, r, c):
    # Ô AM192
    # = IF(AND(AJ192="",0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+8,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+8,9)+1)=9)),"Đông Nam (Tốn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạc
    #   h","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+8,9)+1))&"; ","")
    return (X.cat(X.cat('Đông Nam (Tốn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(0.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),9.0)))) else '')

def f362(E, r, c):
    # Ô AJ193
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=4,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=4,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=4,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   4,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),4.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),4.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),4.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),2.0)) else ''))

def f363(E, r, c):
    # Ô AK193, AK201
    # = IF(AJ193<>"","Nam (Ly) ("&LEFT(AJ193,LEN(AJ193)-2)&"); ","")
    return (X.cat(X.cat('Nam (Ly) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f364(E, r, c):
    # Ô AL193
    # = IF(AND(AJ193="",OR(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Nam (Ly)
    #    – "&INDEX(HANH_TEN,MOD(1,5)+1)&IF(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị hắc
    #   ","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1))&" – "&INDEX({"tốt
    #    – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt
    #    – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nam (Ly) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(1.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(1.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f365(E, r, c):
    # Ô AM193
    # = IF(AND(AJ193="",1<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),1<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),1<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+4,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+4,9)+1)=9)),"Nam (Ly) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cử
    #   u tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+4,9)+1))&"; ","")
    return (X.cat(X.cat('Nam (Ly) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(1.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(1.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),9.0)))) else '')

def f366(E, r, c):
    # Ô AJ194
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=5,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=5,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=5,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   6,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),5.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),5.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),5.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),2.0)) else ''))

def f367(E, r, c):
    # Ô AK194, AK202
    # = IF(AJ194<>"","Tây Nam (Khôn) ("&LEFT(AJ194,LEN(AJ194)-2)&"); ","")
    return (X.cat(X.cat('Tây Nam (Khôn) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f368(E, r, c):
    # Ô AL194
    # = IF(AND(AJ194="",OR(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Tây Nam 
    #   (Khôn) – "&INDEX(HANH_TEN,MOD(2,5)+1)&IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","N
    #   hị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1))&" – "&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tây Nam (Khôn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(2.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f369(E, r, c):
    # Ô AM194
    # = IF(AND(AJ194="",2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+6,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+6,9)+1)=9)),"Tây Nam (Khôn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạc
    #   h","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+6,9)+1))&"; ","")
    return (X.cat(X.cat('Tây Nam (Khôn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(2.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),9.0)))) else '')

def f370(E, r, c):
    # Ô AJ195
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=6,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=6,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=6,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   2,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),6.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),6.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),6.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),2.0)) else ''))

def f371(E, r, c):
    # Ô AK195, AK203
    # = IF(AJ195<>"","Tây (Đoài) ("&LEFT(AJ195,LEN(AJ195)-2)&"); ","")
    return (X.cat(X.cat('Tây (Đoài) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f372(E, r, c):
    # Ô AL195
    # = IF(AND(AJ195="",OR(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Tây (Đoà
    #   i) – "&INDEX(HANH_TEN,MOD(3,5)+1)&IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị h
    #   ắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1))&" – "&INDEX({"t
    #   ốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","t
    #   ốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tây (Đoài) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(3.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f373(E, r, c):
    # Ô AM195
    # = IF(AND(AJ195="",3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+2,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+2,9)+1)=9)),"Tây (Đoài) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","
    #   Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+2,9)+1))&"; ","")
    return (X.cat(X.cat('Tây (Đoài) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(3.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),9.0)))) else '')

def f374(E, r, c):
    # Ô AJ196
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+0)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=7,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=7,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+0)-4,12),4)+1,4,2,0,6)=7,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+
    #   1,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),7.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),7.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),7.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),2.0)) else ''))

def f375(E, r, c):
    # Ô AK196, AK204
    # = IF(AJ196<>"","Tây Bắc (Càn) ("&LEFT(AJ196,LEN(AJ196)-2)&"); ","")
    return (X.cat(X.cat('Tây Bắc (Càn) (',X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0))),'); ') if X.truthy(X.ne(E.v(22,r,c-1),'')) else '')

def f376(E, r, c):
    # Ô AL196
    # = IF(AND(AJ196="",OR(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Tây Bắc 
    #   (Càn) – "&INDEX(HANH_TEN,MOD(3,5)+1)&IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nh
    #   ị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1))&" – "&INDEX(
    #   {"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc"
    #   ,"tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tây Bắc (Càn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(3.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f377(E, r, c):
    # Ô AM196
    # = IF(AND(AJ196="",3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1
    #   )-1+1,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+0),9),9)+1)-1+1,9)+1)=9)),"Tây Bắc (Càn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch
    #   ","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+0),9),9)+1)-1+1,9)+1))&"; ","")
    return (X.cat(X.cat('Tây Bắc (Càn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(3.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),0.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),9.0)))) else '')

def f378(E, r, c):
    # Ô AJ197
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=0,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=0,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=0,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   5,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),0.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),0.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),0.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),2.0)) else ''))

def f379(E, r, c):
    # Ô AL197
    # = IF(AND(AJ197="",OR(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Bắc (Khả
    #   m) – "&INDEX(HANH_TEN,MOD(4,5)+1)&IF(4=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị h
    #   ắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1))&" – "&INDEX({"t
    #   ốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","t
    #   ốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bắc (Khảm) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(4.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(4.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f380(E, r, c):
    # Ô AM197
    # = IF(AND(AJ197="",4<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),4<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),4<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+5,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+5,9)+1)=9)),"Bắc (Khảm) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","
    #   Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+5,9)+1))&"; ","")
    return (X.cat(X.cat('Bắc (Khảm) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(4.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(4.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(4.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),5.0), 9.0),1.0),9.0)))) else '')

def f381(E, r, c):
    # Ô AJ198
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=1,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=1,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=1,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   3,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),1.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),1.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),1.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),2.0)) else ''))

def f382(E, r, c):
    # Ô AL198
    # = IF(AND(AJ198="",OR(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Đông Bắc
    #    (Cấn) – "&INDEX(HANH_TEN,MOD(2,5)+1)&IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","N
    #   hị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1))&" – "&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đông Bắc (Cấn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(2.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f383(E, r, c):
    # Ô AM198
    # = IF(AND(AJ198="",2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+3,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+3,9)+1)=9)),"Đông Bắc (Cấn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạc
    #   h","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+3,9)+1))&"; ","")
    return (X.cat(X.cat('Đông Bắc (Cấn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(2.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),3.0), 9.0),1.0),9.0)))) else '')

def f384(E, r, c):
    # Ô AJ199
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=2,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=2,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=2,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   7,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),2.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),2.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),2.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),2.0)) else ''))

def f385(E, r, c):
    # Ô AL199
    # = IF(AND(AJ199="",OR(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Đông (Ch
    #   ấn) – "&INDEX(HANH_TEN,MOD(0,5)+1)&IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị 
    #   hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1))&" – "&INDEX({"
    #   tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","
    #   tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đông (Chấn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(0.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f386(E, r, c):
    # Ô AM199
    # = IF(AND(AJ199="",0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+7,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+7,9)+1)=9)),"Đông (Chấn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch",
    #   "Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+7,9)+1))&"; ","")
    return (X.cat(X.cat('Đông (Chấn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(0.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),7.0), 9.0),1.0),9.0)))) else '')

def f387(E, r, c):
    # Ô AJ200
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=3,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=3,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=3,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   8,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),3.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),3.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),3.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),2.0)) else ''))

def f388(E, r, c):
    # Ô AL200
    # = IF(AND(AJ200="",OR(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Đông Nam
    #    (Tốn) – "&INDEX(HANH_TEN,MOD(0,5)+1)&IF(0=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","N
    #   hị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1))&" – "&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đông Nam (Tốn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(0.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f389(E, r, c):
    # Ô AM200
    # = IF(AND(AJ200="",0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),0<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),0<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+8,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+8,9)+1)=9)),"Đông Nam (Tốn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạc
    #   h","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+8,9)+1))&"; ","")
    return (X.cat(X.cat('Đông Nam (Tốn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(0.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(0.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(0.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),8.0), 9.0),1.0),9.0)))) else '')

def f390(E, r, c):
    # Ô AJ201
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=4,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=4,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=4,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   4,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),4.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),4.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),4.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),2.0)) else ''))

def f391(E, r, c):
    # Ô AL201
    # = IF(AND(AJ201="",OR(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Nam (Ly)
    #    – "&INDEX(HANH_TEN,MOD(1,5)+1)&IF(1=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị hắc
    #   ","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1))&" – "&INDEX({"tốt
    #    – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","tốt
    #    – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nam (Ly) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(1.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(1.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f392(E, r, c):
    # Ô AM201
    # = IF(AND(AJ201="",1<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),1<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),1<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+4,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+4,9)+1)=9)),"Nam (Ly) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cử
    #   u tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+4,9)+1))&"; ","")
    return (X.cat(X.cat('Nam (Ly) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(1.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(1.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(1.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),4.0), 9.0),1.0),9.0)))) else '')

def f393(E, r, c):
    # Ô AJ202
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=5,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=5,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=5,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   6,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),5.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),5.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),5.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),2.0)) else ''))

def f394(E, r, c):
    # Ô AL202
    # = IF(AND(AJ202="",OR(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Tây Nam 
    #   (Khôn) – "&INDEX(HANH_TEN,MOD(2,5)+1)&IF(2=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","N
    #   hị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1))&" – "&INDEX
    #   ({"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc
    #   ","tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tây Nam (Khôn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(2.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f395(E, r, c):
    # Ô AM202
    # = IF(AND(AJ202="",2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),2<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),2<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+6,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+6,9)+1)=9)),"Tây Nam (Khôn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạc
    #   h","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+6,9)+1))&"; ","")
    return (X.cat(X.cat('Tây Nam (Khôn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(2.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(2.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(2.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),6.0), 9.0),1.0),9.0)))) else '')

def f396(E, r, c):
    # Ô AJ203
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=6,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=6,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=6,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   2,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),6.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),6.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),6.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),2.0)) else ''))

def f397(E, r, c):
    # Ô AL203
    # = IF(AND(AJ203="",OR(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Tây (Đoà
    #   i) – "&INDEX(HANH_TEN,MOD(3,5)+1)&IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nhị h
    #   ắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1))&" – "&INDEX({"t
    #   ốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc","t
    #   ốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tây (Đoài) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(3.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f398(E, r, c):
    # Ô AM203
    # = IF(AND(AJ203="",3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+2,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+2,9)+1)=9)),"Tây (Đoài) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","
    #   Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+2,9)+1))&"; ","")
    return (X.cat(X.cat('Tây (Đoài) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(3.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),2.0), 9.0),1.0),9.0)))) else '')

def f399(E, r, c):
    # Ô AJ204
    # = IF(CHOOSE(MOD((DAI_VAN!$C$23+1)-4,12)+1,0,1,1,2,3,3,4,5,5,6,7,7)=7,"thái tuế, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)+1,0,1,1,2,3,3,4,5,
    #   5,6,7,7)=7,"tuế phá, ","")&IF(CHOOSE(MOD(MOD((DAI_VAN!$C$23+1)-4,12),4)+1,4,2,0,6)=7,"tam sát, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+
    #   1,9)+1)=5,"Ngũ hoàng, ","")&IF((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1)=2,"Nhị hắc, ","")
    return X.cat(X.cat(X.cat(X.cat(('thái tuế, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),7.0)) else ''),('tuế phá, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0),1.0), lambda: 0.0, lambda: 1.0, lambda: 1.0, lambda: 2.0, lambda: 3.0, lambda: 3.0, lambda: 4.0, lambda: 5.0, lambda: 5.0, lambda: 6.0, lambda: 7.0, lambda: 7.0),7.0)) else '')),('tam sát, ' if X.truthy(X.eq(X.CHOOSE(X.add(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0), 4.0),1.0), lambda: 4.0, lambda: 2.0, lambda: 0.0, lambda: 6.0),7.0)) else '')),('Ngũ hoàng, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),5.0)) else '')),('Nhị hắc, ' if X.truthy(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),2.0)) else ''))

def f400(E, r, c):
    # Ô AL204
    # = IF(AND(AJ204="",OR(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Tây Bắc 
    #   (Càn) – "&INDEX(HANH_TEN,MOD(3,5)+1)&IF(3=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (dụng)"," (hỷ)")&", sao "&INDEX({"Nhất bạch","Nh
    #   ị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1))&" – "&INDEX(
    #   {"tốt – quý nhân","XẤU – bệnh tật","thị phi, tranh cãi","văn xương – tốt cho học","XẤU – tai ách","tốt – quyền lộc","hao tổn, mất mát","tốt – tài lộc"
    #   ,"tốt – hỷ khánh"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1))&"; ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tây Bắc (Càn) – ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(3.0, 5.0),1.0))),(' (dụng)' if X.truthy(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ' (hỷ)')),', sao '),X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),' – '),X.INDEX(X.Arr([['tốt – quý nhân', 'XẤU – bệnh tật', 'thị phi, tranh cãi', 'văn xương – tốt cho học', 'XẤU – tai ách', 'tốt – quyền lộc', 'hao tổn, mất mát', 'tốt – tài lộc', 'tốt – hỷ khánh']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-2),''), X.OR(X.eq(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else '')

def f401(E, r, c):
    # Ô AM204
    # = IF(AND(AJ204="",3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),3<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),3<>IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),OR((MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1)=1,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1
    #   )-1+1,9)+1)=4,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1)=6,(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1)=8,(MOD((MOD(10-MOD((DAI_VA
    #   N!$C$23+1),9),9)+1)-1+1,9)+1)=9)),"Tây Bắc (Càn) – sao "&INDEX({"Nhất bạch","Nhị hắc","Tam bích","Tứ lục","Ngũ hoàng","Lục bạch","Thất xích","Bát bạch
    #   ","Cửu tử"},(MOD((MOD(10-MOD((DAI_VAN!$C$23+1),9),9)+1)-1+1,9)+1))&"; ","")
    return (X.cat(X.cat('Tây Bắc (Càn) – sao ',X.INDEX(X.Arr([['Nhất bạch', 'Nhị hắc', 'Tam bích', 'Tứ lục', 'Ngũ hoàng', 'Lục bạch', 'Thất xích', 'Bát bạch', 'Cửu tử']]), X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0))),'; ') if X.truthy(X.AND(X.eq(E.v(22,r,c-3),''), X.ne(3.0,(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.ne(3.0,(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(3.0,(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.OR(X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),1.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),4.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),6.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),8.0), X.eq(X.add(X.MOD(X.add(X.sub(X.add(X.MOD(X.sub(10.0,X.MOD(X.add(E.v(14,23,3),1.0), 9.0)), 9.0),1.0),1.0),1.0), 9.0),1.0),9.0)))) else '')

def f402(E, r, c):
    # Ô AJ209
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=0),"mục 1 – vượng gặp năm
    #    tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(M
    #   OD((DAI_VAN!$C$23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3),"mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay 
    #   kiện tụng” – bồi bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(AND
    #   ((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=4),"mục 3 – vượng gặp năm ấn: “
    #   không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C
    #   $23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1),"mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C$23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=2),"mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điề
    #   u xấu” – giữ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(OR(AND(MOD
    #   (INT(MOD((DAI_VAN!$C$23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3,NOT(MOD(MOD((DAI_VAN!$C$23+0)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(),TRUNG_GIAN!$C
    #   $12="Thương quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Thương quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Thương quan"),1
    #   ,0))>0),AND(MOD(INT(MOD((DAI_VAN!$C$23+0)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1,NOT(MOD(MOD((DAI_VAN!$C$23+0)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(
    #   ),TRUNG_GIAN!$C$12="Chính quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Chính quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Ch
    #   ính quan"),1,0))>0)),"mục 6 – thương quan gặp quan: nên thông quan; ","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"nă
    #   m ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12
    #   )+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"giờ ",""))<>"","mục 7 – chi năm
    #    xung chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+0
    #   )-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+0)-4,12)+6,12)),"giờ ","")&"(xung thái tuế): chọn cách thông quan"&IF(OR(MOD((DAI_VAN!$C$23+0)-4,12)=1,MOD((DA
    #   I_VAN!$C$23+0)-4,12)=4,MOD((DAI_VAN!$C$23+0)-4,12)=7,MOD((DAI_VAN!$C$23+0)-4,12)=10),"; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận","")&"; ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('mục 1 – vượng gặp năm tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),0.0))) else ''),(X.cat(X.cat('mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng” – bồi bổ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0))) else '')),('mục 3 – vượng gặp năm ấn: “không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),4.0))) else '')),('mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0))) else '')),(X.cat(X.cat('mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điều xấu” – giữ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),2.0))) else '')),('mục 6 – thương quan gặp quan: nên thông quan; ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thương quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thương quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thương quan'))) else 0.0)),0.0)), X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Chính quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Chính quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính quan'))) else 0.0)),0.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('mục 7 – chi năm xung chi trụ ',('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),'(xung thái tuế): chọn cách thông quan'),('; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận' if X.truthy(X.OR(X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),4.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),7.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),10.0))) else '')),'; ') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else ''))

def f403(E, r, c):
    # Ô AK209
    # = "• "&(DAI_VAN!$C$23+0)&" "&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+0)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+0)-4,12)+1)&": "&IF(AJ209="","không thuộc
    #    trường hợp nào",LEFT(AJ209,LEN(AJ209)-2))&CHAR(10)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.add(E.v(14,23,3),0.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),0.0),4.0), 12.0),1.0))),': '),('không thuộc trường hợp nào' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0)))),X.CHAR(10.0))

def f404(E, r, c):
    # Ô AJ210
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=0),"mục 1 – vượng gặp năm
    #    tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(M
    #   OD((DAI_VAN!$C$23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3),"mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay 
    #   kiện tụng” – bồi bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(AND
    #   ((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=4),"mục 3 – vượng gặp năm ấn: “
    #   không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C
    #   $23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1),"mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C$23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=2),"mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điề
    #   u xấu” – giữ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(OR(AND(MOD
    #   (INT(MOD((DAI_VAN!$C$23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3,NOT(MOD(MOD((DAI_VAN!$C$23+1)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(),TRUNG_GIAN!$C
    #   $12="Thương quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Thương quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Thương quan"),1
    #   ,0))>0),AND(MOD(INT(MOD((DAI_VAN!$C$23+1)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1,NOT(MOD(MOD((DAI_VAN!$C$23+1)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(
    #   ),TRUNG_GIAN!$C$12="Chính quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Chính quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Ch
    #   ính quan"),1,0))>0)),"mục 6 – thương quan gặp quan: nên thông quan; ","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"nă
    #   m ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12
    #   )+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"giờ ",""))<>"","mục 7 – chi năm
    #    xung chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+1
    #   )-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+1)-4,12)+6,12)),"giờ ","")&"(xung thái tuế): chọn cách thông quan"&IF(OR(MOD((DAI_VAN!$C$23+1)-4,12)=1,MOD((DA
    #   I_VAN!$C$23+1)-4,12)=4,MOD((DAI_VAN!$C$23+1)-4,12)=7,MOD((DAI_VAN!$C$23+1)-4,12)=10),"; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận","")&"; ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('mục 1 – vượng gặp năm tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),0.0))) else ''),(X.cat(X.cat('mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng” – bồi bổ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0))) else '')),('mục 3 – vượng gặp năm ấn: “không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),4.0))) else '')),('mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0))) else '')),(X.cat(X.cat('mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điều xấu” – giữ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),2.0))) else '')),('mục 6 – thương quan gặp quan: nên thông quan; ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thương quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thương quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thương quan'))) else 0.0)),0.0)), X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Chính quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Chính quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính quan'))) else 0.0)),0.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('mục 7 – chi năm xung chi trụ ',('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),'(xung thái tuế): chọn cách thông quan'),('; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận' if X.truthy(X.OR(X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),4.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),7.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),10.0))) else '')),'; ') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else ''))

def f405(E, r, c):
    # Ô AK210
    # = "• "&(DAI_VAN!$C$23+1)&" "&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+1)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+1)-4,12)+1)&": "&IF(AJ210="","không thuộc
    #    trường hợp nào",LEFT(AJ210,LEN(AJ210)-2))&CHAR(10)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.add(E.v(14,23,3),1.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),1.0),4.0), 12.0),1.0))),': '),('không thuộc trường hợp nào' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0)))),X.CHAR(10.0))

def f406(E, r, c):
    # Ô AJ211
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=0),"mục 1 – vượng gặp năm
    #    tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(M
    #   OD((DAI_VAN!$C$23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3),"mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay 
    #   kiện tụng” – bồi bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(AND
    #   ((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=4),"mục 3 – vượng gặp năm ấn: “
    #   không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C
    #   $23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1),"mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C$23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=2),"mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điề
    #   u xấu” – giữ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(OR(AND(MOD
    #   (INT(MOD((DAI_VAN!$C$23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3,NOT(MOD(MOD((DAI_VAN!$C$23+2)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(),TRUNG_GIAN!$C
    #   $12="Thương quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Thương quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Thương quan"),1
    #   ,0))>0),AND(MOD(INT(MOD((DAI_VAN!$C$23+2)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1,NOT(MOD(MOD((DAI_VAN!$C$23+2)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(
    #   ),TRUNG_GIAN!$C$12="Chính quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Chính quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Ch
    #   ính quan"),1,0))>0)),"mục 6 – thương quan gặp quan: nên thông quan; ","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12)+6,12)),"nă
    #   m ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12
    #   )+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12)+6,12)),"giờ ",""))<>"","mục 7 – chi năm
    #    xung chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+2
    #   )-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+2)-4,12)+6,12)),"giờ ","")&"(xung thái tuế): chọn cách thông quan"&IF(OR(MOD((DAI_VAN!$C$23+2)-4,12)=1,MOD((DA
    #   I_VAN!$C$23+2)-4,12)=4,MOD((DAI_VAN!$C$23+2)-4,12)=7,MOD((DAI_VAN!$C$23+2)-4,12)=10),"; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận","")&"; ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('mục 1 – vượng gặp năm tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),0.0))) else ''),(X.cat(X.cat('mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng” – bồi bổ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0))) else '')),('mục 3 – vượng gặp năm ấn: “không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),4.0))) else '')),('mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0))) else '')),(X.cat(X.cat('mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điều xấu” – giữ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),2.0))) else '')),('mục 6 – thương quan gặp quan: nên thông quan; ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thương quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thương quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thương quan'))) else 0.0)),0.0)), X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Chính quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Chính quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính quan'))) else 0.0)),0.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('mục 7 – chi năm xung chi trụ ',('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),'(xung thái tuế): chọn cách thông quan'),('; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận' if X.truthy(X.OR(X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),1.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),4.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),7.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),10.0))) else '')),'; ') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else ''))

def f407(E, r, c):
    # Ô AK211
    # = "• "&(DAI_VAN!$C$23+2)&" "&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+2)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+2)-4,12)+1)&": "&IF(AJ211="","không thuộc
    #    trường hợp nào",LEFT(AJ211,LEN(AJ211)-2))&CHAR(10)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.add(E.v(14,23,3),2.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),2.0),4.0), 12.0),1.0))),': '),('không thuộc trường hợp nào' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0)))),X.CHAR(10.0))

def f408(E, r, c):
    # Ô AJ212
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=0),"mục 1 – vượng gặp năm
    #    tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(M
    #   OD((DAI_VAN!$C$23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3),"mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay 
    #   kiện tụng” – bồi bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(AND
    #   ((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=4),"mục 3 – vượng gặp năm ấn: “
    #   không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C
    #   $23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1),"mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C$23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=2),"mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điề
    #   u xấu” – giữ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(OR(AND(MOD
    #   (INT(MOD((DAI_VAN!$C$23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3,NOT(MOD(MOD((DAI_VAN!$C$23+3)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(),TRUNG_GIAN!$C
    #   $12="Thương quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Thương quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Thương quan"),1
    #   ,0))>0),AND(MOD(INT(MOD((DAI_VAN!$C$23+3)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1,NOT(MOD(MOD((DAI_VAN!$C$23+3)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(
    #   ),TRUNG_GIAN!$C$12="Chính quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Chính quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Ch
    #   ính quan"),1,0))>0)),"mục 6 – thương quan gặp quan: nên thông quan; ","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12)+6,12)),"nă
    #   m ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12
    #   )+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12)+6,12)),"giờ ",""))<>"","mục 7 – chi năm
    #    xung chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+3
    #   )-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+3)-4,12)+6,12)),"giờ ","")&"(xung thái tuế): chọn cách thông quan"&IF(OR(MOD((DAI_VAN!$C$23+3)-4,12)=1,MOD((DA
    #   I_VAN!$C$23+3)-4,12)=4,MOD((DAI_VAN!$C$23+3)-4,12)=7,MOD((DAI_VAN!$C$23+3)-4,12)=10),"; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận","")&"; ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('mục 1 – vượng gặp năm tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),0.0))) else ''),(X.cat(X.cat('mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng” – bồi bổ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0))) else '')),('mục 3 – vượng gặp năm ấn: “không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),4.0))) else '')),('mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0))) else '')),(X.cat(X.cat('mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điều xấu” – giữ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),2.0))) else '')),('mục 6 – thương quan gặp quan: nên thông quan; ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thương quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thương quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thương quan'))) else 0.0)),0.0)), X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Chính quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Chính quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính quan'))) else 0.0)),0.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('mục 7 – chi năm xung chi trụ ',('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),'(xung thái tuế): chọn cách thông quan'),('; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận' if X.truthy(X.OR(X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),1.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),4.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),7.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),10.0))) else '')),'; ') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else ''))

def f409(E, r, c):
    # Ô AK212
    # = "• "&(DAI_VAN!$C$23+3)&" "&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+3)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+3)-4,12)+1)&": "&IF(AJ212="","không thuộc
    #    trường hợp nào",LEFT(AJ212,LEN(AJ212)-2))&CHAR(10)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.add(E.v(14,23,3),3.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),3.0),4.0), 12.0),1.0))),': '),('không thuộc trường hợp nào' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0)))),X.CHAR(10.0))

def f410(E, r, c):
    # Ô AJ213
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=0),"mục 1 – vượng gặp năm
    #    tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(M
    #   OD((DAI_VAN!$C$23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3),"mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay 
    #   kiện tụng” – bồi bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(AND
    #   ((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=4),"mục 3 – vượng gặp năm ấn: “
    #   không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C
    #   $23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1),"mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C$23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=2),"mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điề
    #   u xấu” – giữ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(OR(AND(MOD
    #   (INT(MOD((DAI_VAN!$C$23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3,NOT(MOD(MOD((DAI_VAN!$C$23+4)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(),TRUNG_GIAN!$C
    #   $12="Thương quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Thương quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Thương quan"),1
    #   ,0))>0),AND(MOD(INT(MOD((DAI_VAN!$C$23+4)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1,NOT(MOD(MOD((DAI_VAN!$C$23+4)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(
    #   ),TRUNG_GIAN!$C$12="Chính quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Chính quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Ch
    #   ính quan"),1,0))>0)),"mục 6 – thương quan gặp quan: nên thông quan; ","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12)+6,12)),"nă
    #   m ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12
    #   )+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12)+6,12)),"giờ ",""))<>"","mục 7 – chi năm
    #    xung chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+4
    #   )-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+4)-4,12)+6,12)),"giờ ","")&"(xung thái tuế): chọn cách thông quan"&IF(OR(MOD((DAI_VAN!$C$23+4)-4,12)=1,MOD((DA
    #   I_VAN!$C$23+4)-4,12)=4,MOD((DAI_VAN!$C$23+4)-4,12)=7,MOD((DAI_VAN!$C$23+4)-4,12)=10),"; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận","")&"; ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('mục 1 – vượng gặp năm tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),0.0))) else ''),(X.cat(X.cat('mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng” – bồi bổ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0))) else '')),('mục 3 – vượng gặp năm ấn: “không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),4.0))) else '')),('mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0))) else '')),(X.cat(X.cat('mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điều xấu” – giữ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),2.0))) else '')),('mục 6 – thương quan gặp quan: nên thông quan; ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thương quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thương quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thương quan'))) else 0.0)),0.0)), X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Chính quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Chính quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính quan'))) else 0.0)),0.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('mục 7 – chi năm xung chi trụ ',('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),'(xung thái tuế): chọn cách thông quan'),('; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận' if X.truthy(X.OR(X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),1.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),4.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),7.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),10.0))) else '')),'; ') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else ''))

def f411(E, r, c):
    # Ô AK213
    # = "• "&(DAI_VAN!$C$23+4)&" "&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+4)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+4)-4,12)+1)&": "&IF(AJ213="","không thuộc
    #    trường hợp nào",LEFT(AJ213,LEN(AJ213)-2))&CHAR(10)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.add(E.v(14,23,3),4.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),4.0),4.0), 12.0),1.0))),': '),('không thuộc trường hợp nào' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0)))),X.CHAR(10.0))

def f412(E, r, c):
    # Ô AJ214
    # = IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=0),"mục 1 – vượng gặp năm
    #    tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(M
    #   OD((DAI_VAN!$C$23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3),"mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay 
    #   kiện tụng” – bồi bổ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(AND
    #   ((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),MOD(INT(MOD((DAI_VAN!$C$23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=4),"mục 3 – vượng gặp năm ấn: “
    #   không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C
    #   $23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1),"mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ","")&IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$7,DUNG_HY_KY!$C$7)=0),MOD(INT(MOD((DAI_VAN!$C$23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=2),"mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điề
    #   u xấu” – giữ "&CHOOSE(MOD(TRUNG_GIAN!$E$7,5)+1,"gan – mật","tim – tiểu tràng","tì – vị","phổi – đại tràng","thận – bàng quang")&"; ","")&IF(OR(AND(MOD
    #   (INT(MOD((DAI_VAN!$C$23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=3,NOT(MOD(MOD((DAI_VAN!$C$23+5)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(),TRUNG_GIAN!$C
    #   $12="Thương quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Thương quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Thương quan"),1
    #   ,0))>0),AND(MOD(INT(MOD((DAI_VAN!$C$23+5)-4,10)/2)-TRUNG_GIAN!$E$7+5,5)=1,NOT(MOD(MOD((DAI_VAN!$C$23+5)-4,10),2)=MOD(TRUNG_GIAN!$E$5,2)),(IF(AND(TRUE(
    #   ),TRUNG_GIAN!$C$12="Chính quan"),1,0)+IF(AND(TRUE(),TRUNG_GIAN!$D$12="Chính quan"),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$12="Ch
    #   ính quan"),1,0))>0)),"mục 6 – thương quan gặp quan: nên thông quan; ","")&IF((IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12)+6,12)),"nă
    #   m ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12
    #   )+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12)+6,12)),"giờ ",""))<>"","mục 7 – chi năm
    #    xung chi trụ "&IF(AND(TRUE(),TRUNG_GIAN!$C$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12)+6,12)),"năm ","")&IF(AND(TRUE(),TRUNG_GIAN!$D$6=MOD(MOD((DAI_VAN!$C$23+5
    #   )-4,12)+6,12)),"tháng ","")&IF(AND(TRUE(),TRUNG_GIAN!$E$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12)+6,12)),"ngày ","")&IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,TRUNG_GIAN!$F$6=MOD(MOD((DAI_VAN!$C$23+5)-4,12)+6,12)),"giờ ","")&"(xung thái tuế): chọn cách thông quan"&IF(OR(MOD((DAI_VAN!$C$23+5)-4,12)=1,MOD((DA
    #   I_VAN!$C$23+5)-4,12)=4,MOD((DAI_VAN!$C$23+5)-4,12)=7,MOD((DAI_VAN!$C$23+5)-4,12)=10),"; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận","")&"; ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('mục 1 – vượng gặp năm tỉ kiếp: “nếu không có bệnh thì cũng phát sinh tranh chấp, phá tài”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),0.0))) else ''),(X.cat(X.cat('mục 2 – nhược gặp năm quan sát: sức đề kháng yếu, “nếu không đau ốm thì cũng bị … tai nạn, hay kiện tụng” – bồi bổ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0))) else '')),('mục 3 – vượng gặp năm ấn: “không ốm đau thì cũng gặp điều xấu” – giảm bớt hành ấn; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),4.0))) else '')),('mục 4 – nhược gặp năm thực thương: “không ốm đau cũng gặp rủi ro”; ' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0))) else '')),(X.cat(X.cat('mục 5 – nhược gặp năm tài: “không ốm đau thì cũng gặp điều xấu” – giữ ',X.CHOOSE(X.add(X.MOD(E.v(23,7,5), 5.0),1.0), lambda: 'gan – mật', lambda: 'tim – tiểu tràng', lambda: 'tì – vị', lambda: 'phổi – đại tràng', lambda: 'thận – bàng quang')),'; ') if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),2.0))) else '')),('mục 6 – thương quan gặp quan: nên thông quan; ' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),3.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thương quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thương quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thương quan'))) else 0.0)),0.0)), X.AND(X.eq(X.MOD(X.add(X.sub(X.INT(X.div(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),2.0)),E.v(23,7,5)),5.0), 5.0),1.0), X.NOT(X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0), 2.0),X.MOD(E.v(23,5,5), 2.0))), X.gt(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Chính quan'))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Chính quan'))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Chính quan'))) else 0.0)),0.0)))) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('mục 7 – chi năm xung chi trụ ',('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),'(xung thái tuế): chọn cách thông quan'),('; Thìn Tuất Sửu Mùi xung – dễ bệnh tì, vị, thận' if X.truthy(X.OR(X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),1.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),4.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),7.0), X.eq(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),10.0))) else '')),'; ') if X.truthy(X.ne(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else ''),('tháng ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),('ngày ' if X.truthy(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),('giờ ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),6.0), 12.0)))) else '')),'')) else ''))

def f413(E, r, c):
    # Ô AK214
    # = "• "&(DAI_VAN!$C$23+5)&" "&INDEX(CAN_TEN,MOD((DAI_VAN!$C$23+5)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((DAI_VAN!$C$23+5)-4,12)+1)&": "&IF(AJ214="","không thuộc
    #    trường hợp nào",LEFT(AJ214,LEN(AJ214)-2))&CHAR(10)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.add(E.v(14,23,3),5.0)),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(14,23,3),5.0),4.0), 12.0),1.0))),': '),('không thuộc trường hợp nào' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.LEFT(E.v(22,r,c-1), X.sub(X.LEN(E.v(22,r,c-1)),2.0)))),X.CHAR(10.0))

def f414(E, r, c):
    # Ô E219
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f415(E, r, c):
    # Ô F219
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f416(E, r, c):
    # Ô K219
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (HG tr.201): Thân vượng và có ấn → "&IF(J2
    #   19="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J219&") ⇒ "&G219&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (HG tr.201): Thân vượng và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f417(E, r, c):
    # Ô E220, F220
    # = IF(AND((PP1_DINH_TINH!$F$96>=P_NHIEU),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,96,6),E.v(24,99,2)), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f418(E, r, c):
    # Ô K220
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (HG tr.201): Tỉ kiếp ≥ P_NHIEU và tài ≥ P_NHIEU → "&IF(J220="","c
    #   hưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J220&") ⇒ "&G220&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (HG tr.201): Tỉ kiếp ≥ P_NHIEU và tài ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f419(E, r, c):
    # Ô E221, F221
    # = IF((COUNTIFS(SAO!$O$5:$O$20,1,SAO!$X$5:$X$20,1)>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,15,20,15), 1.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f420(E, r, c):
    # Ô K221
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$O$5:$O$20,1,SAO!$X$5:$X$20,1)>0,"thiên/nguyệt đức: "&COUNTIFS(SAO!$O
    #   $5:$O$20,1,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (HG tr.201): Tứ trụ có thiên đức hoặc nguyệt đức → "&IF(J221="","chưa đủ điều kiện ở cả hai phư
    #   ơng pháp – không áp dụng.","ĐỦ ("&J221&") ⇒ "&G221&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,15,20,15), 1.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,15,20,15), 1.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (HG tr.201): Tứ trụ có thiên đức hoặc nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f421(E, r, c):
    # Ô E222, F222
    # = IF((PP1_DINH_TINH!$C$67=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(9,67,3),1.0)) else '—')

def f422(E, r, c):
    # Ô K222
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (HG tr.201): Có tam hợp/tam hội cục hành ấn hoặc tỉ kiếp đã hóa (PP1_DINH_TINH C67) → "&IF(J222="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J222&") ⇒ "&G222&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (HG tr.201): Có tam hợp/tam hội cục hành ấn hoặc tỉ kiếp đã hóa (PP1_DINH_TINH C67) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f423(E, r, c):
    # Ô E223
    # = IF(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f424(E, r, c):
    # Ô F223
    # = IF(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f425(E, r, c):
    # Ô K223
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (HG tr.202): Thân nhược và có ấn sinh → "&
    #   IF(J223="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J223&") ⇒ "&G223&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (HG tr.202): Thân nhược và có ấn sinh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

def f426(E, r, c):
    # Ô E224, F224
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),0.0)) else '—')

def f427(E, r, c):
    # Ô K224
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Bản khí",SAO!$B$5:$B$20,1)>0,"Chính ấn bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản 
    #   khí",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)
    #   >0,"Thiên ấn bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ;
    #    ",""))&" → Quy tắc (HG tr.202): Bản khí chi tháng là ấn → "&IF(J224="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J224&") ⇒ "&G
    #   224&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (HG tr.202): Bản khí chi tháng là ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(22,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(22,r,c-1)),') ⇒ '),E.v(22,r,c-4)),'.')))

CELLS = {
    'C6': (f0, 6, 3),
    'F6': (f1, 6, 6),
    'J6': (f2, 6, 10),
    'M6': (f3, 6, 13),
    'N6': (f4, 6, 14),
    'C7': (f5, 7, 3),
    'F7': (f6, 7, 6),
    'J7': (f7, 7, 10),
    'M7': (f8, 7, 13),
    'N7': (f9, 7, 14),
    'AB7': (f10, 7, 28),
    'AC7': (f11, 7, 29),
    'AD7': (f12, 7, 30),
    'AE7': (f13, 7, 31),
    'C8': (f5, 8, 3),
    'F8': (f6, 8, 6),
    'J8': (f14, 8, 10),
    'M8': (f15, 8, 13),
    'N8': (f16, 8, 14),
    'AB8': (f11, 8, 28),
    'AC8': (f11, 8, 29),
    'AD8': (f17, 8, 30),
    'AE8': (f18, 8, 31),
    'B13': (f19, 13, 2),
    'C13': (f20, 13, 3),
    'D13': (f21, 13, 4),
    'E13': (f22, 13, 5),
    'F13': (f23, 13, 6),
    'G13': (f24, 13, 7),
    'B14': (f25, 14, 2),
    'C14': (f26, 14, 3),
    'D14': (f27, 14, 4),
    'E14': (f28, 14, 5),
    'F14': (f29, 14, 6),
    'G14': (f30, 14, 7),
    'B15': (f31, 15, 2),
    'C15': (f32, 15, 3),
    'D15': (f33, 15, 4),
    'E15': (f34, 15, 5),
    'F15': (f35, 15, 6),
    'G15': (f36, 15, 7),
    'B16': (f37, 16, 2),
    'C16': (f38, 16, 3),
    'D16': (f39, 16, 4),
    'E16': (f40, 16, 5),
    'F16': (f41, 16, 6),
    'G16': (f42, 16, 7),
    'B17': (f43, 17, 2),
    'C17': (f44, 17, 3),
    'D17': (f45, 17, 4),
    'E17': (f46, 17, 5),
    'F17': (f47, 17, 6),
    'G17': (f48, 17, 7),
    'C18': (f49, 18, 3),
    'C21': (f50, 21, 3),
    'D21': (f51, 21, 4),
    'E21': (f52, 21, 5),
    'F21': (f53, 21, 6),
    'G21': (f54, 21, 7),
    'H21': (f55, 21, 8),
    'I21': (f56, 21, 9),
    'J21': (f57, 21, 10),
    'C22': (f58, 22, 3),
    'D22': (f58, 22, 4),
    'E22': (f58, 22, 5),
    'F22': (f58, 22, 6),
    'G22': (f58, 22, 7),
    'H22': (f58, 22, 8),
    'I22': (f58, 22, 9),
    'J22': (f58, 22, 10),
    'C23': (f59, 23, 3),
    'D23': (f59, 23, 4),
    'E23': (f59, 23, 5),
    'F23': (f59, 23, 6),
    'G23': (f59, 23, 7),
    'H23': (f59, 23, 8),
    'I23': (f59, 23, 9),
    'J23': (f59, 23, 10),
    'C24': (f60, 24, 3),
    'D24': (f60, 24, 4),
    'E24': (f60, 24, 5),
    'F24': (f60, 24, 6),
    'G24': (f60, 24, 7),
    'H24': (f60, 24, 8),
    'I24': (f60, 24, 9),
    'J24': (f60, 24, 10),
    'C25': (f61, 25, 3),
    'D25': (f61, 25, 4),
    'E25': (f61, 25, 5),
    'F25': (f61, 25, 6),
    'G25': (f61, 25, 7),
    'H25': (f61, 25, 8),
    'I25': (f61, 25, 9),
    'J25': (f61, 25, 10),
    'C26': (f61, 26, 3),
    'D26': (f61, 26, 4),
    'E26': (f61, 26, 5),
    'F26': (f61, 26, 6),
    'G26': (f61, 26, 7),
    'H26': (f61, 26, 8),
    'I26': (f61, 26, 9),
    'J26': (f61, 26, 10),
    'C27': (f62, 27, 3),
    'D27': (f62, 27, 4),
    'E27': (f62, 27, 5),
    'F27': (f62, 27, 6),
    'G27': (f62, 27, 7),
    'H27': (f62, 27, 8),
    'I27': (f62, 27, 9),
    'J27': (f62, 27, 10),
    'C28': (f63, 28, 3),
    'B32': (f64, 32, 2),
    'C32': (f65, 32, 3),
    'D32': (f66, 32, 4),
    'E32': (f67, 32, 5),
    'F32': (f68, 32, 6),
    'G32': (f69, 32, 7),
    'H32': (f70, 32, 8),
    'N32': (f71, 32, 14),
    'O32': (f72, 32, 15),
    'P32': (f73, 32, 16),
    'Q32': (f74, 32, 17),
    'R32': (f75, 32, 18),
    'S32': (f76, 32, 19),
    'T32': (f77, 32, 20),
    'U32': (f78, 32, 21),
    'V32': (f79, 32, 22),
    'W32': (f80, 32, 23),
    'X32': (f81, 32, 24),
    'Y32': (f82, 32, 25),
    'Z32': (f83, 32, 26),
    'AA32': (f84, 32, 27),
    'AB32': (f85, 32, 28),
    'AC32': (f86, 32, 29),
    'AD32': (f87, 32, 30),
    'AE32': (f88, 32, 31),
    'AF32': (f89, 32, 32),
    'B33': (f90, 33, 2),
    'C33': (f65, 33, 3),
    'D33': (f66, 33, 4),
    'E33': (f67, 33, 5),
    'F33': (f68, 33, 6),
    'G33': (f69, 33, 7),
    'H33': (f70, 33, 8),
    'N33': (f71, 33, 14),
    'O33': (f72, 33, 15),
    'P33': (f73, 33, 16),
    'Q33': (f74, 33, 17),
    'R33': (f75, 33, 18),
    'S33': (f76, 33, 19),
    'T33': (f77, 33, 20),
    'U33': (f78, 33, 21),
    'V33': (f79, 33, 22),
    'W33': (f80, 33, 23),
    'X33': (f81, 33, 24),
    'Y33': (f82, 33, 25),
    'Z33': (f83, 33, 26),
    'AA33': (f84, 33, 27),
    'AB33': (f85, 33, 28),
    'AC33': (f86, 33, 29),
    'AD33': (f87, 33, 30),
    'AE33': (f88, 33, 31),
    'AF33': (f89, 33, 32),
    'B34': (f91, 34, 2),
    'C34': (f65, 34, 3),
    'D34': (f66, 34, 4),
    'E34': (f67, 34, 5),
    'F34': (f68, 34, 6),
    'G34': (f69, 34, 7),
    'H34': (f70, 34, 8),
    'N34': (f71, 34, 14),
    'O34': (f72, 34, 15),
    'P34': (f73, 34, 16),
    'Q34': (f74, 34, 17),
    'R34': (f75, 34, 18),
    'S34': (f76, 34, 19),
    'T34': (f77, 34, 20),
    'U34': (f78, 34, 21),
    'V34': (f79, 34, 22),
    'W34': (f80, 34, 23),
    'X34': (f81, 34, 24),
    'Y34': (f82, 34, 25),
    'Z34': (f83, 34, 26),
    'AA34': (f84, 34, 27),
    'AB34': (f85, 34, 28),
    'AC34': (f86, 34, 29),
    'AD34': (f87, 34, 30),
    'AE34': (f88, 34, 31),
    'AF34': (f89, 34, 32),
    'B35': (f92, 35, 2),
    'C35': (f65, 35, 3),
    'D35': (f66, 35, 4),
    'E35': (f67, 35, 5),
    'F35': (f68, 35, 6),
    'G35': (f69, 35, 7),
    'H35': (f70, 35, 8),
    'N35': (f71, 35, 14),
    'O35': (f72, 35, 15),
    'P35': (f73, 35, 16),
    'Q35': (f74, 35, 17),
    'R35': (f75, 35, 18),
    'S35': (f76, 35, 19),
    'T35': (f77, 35, 20),
    'U35': (f78, 35, 21),
    'V35': (f79, 35, 22),
    'W35': (f80, 35, 23),
    'X35': (f81, 35, 24),
    'Y35': (f82, 35, 25),
    'Z35': (f83, 35, 26),
    'AA35': (f84, 35, 27),
    'AB35': (f85, 35, 28),
    'AC35': (f86, 35, 29),
    'AD35': (f87, 35, 30),
    'AE35': (f88, 35, 31),
    'AF35': (f89, 35, 32),
    'B36': (f93, 36, 2),
    'C36': (f65, 36, 3),
    'D36': (f66, 36, 4),
    'E36': (f67, 36, 5),
    'F36': (f68, 36, 6),
    'G36': (f69, 36, 7),
    'H36': (f70, 36, 8),
    'N36': (f71, 36, 14),
    'O36': (f72, 36, 15),
    'P36': (f73, 36, 16),
    'Q36': (f74, 36, 17),
    'R36': (f75, 36, 18),
    'S36': (f76, 36, 19),
    'T36': (f77, 36, 20),
    'U36': (f78, 36, 21),
    'V36': (f79, 36, 22),
    'W36': (f80, 36, 23),
    'X36': (f81, 36, 24),
    'Y36': (f82, 36, 25),
    'Z36': (f83, 36, 26),
    'AA36': (f84, 36, 27),
    'AB36': (f85, 36, 28),
    'AC36': (f86, 36, 29),
    'AD36': (f87, 36, 30),
    'AE36': (f88, 36, 31),
    'AF36': (f89, 36, 32),
    'B37': (f94, 37, 2),
    'C37': (f65, 37, 3),
    'D37': (f66, 37, 4),
    'E37': (f67, 37, 5),
    'F37': (f68, 37, 6),
    'G37': (f69, 37, 7),
    'H37': (f70, 37, 8),
    'N37': (f71, 37, 14),
    'O37': (f72, 37, 15),
    'P37': (f73, 37, 16),
    'Q37': (f74, 37, 17),
    'R37': (f75, 37, 18),
    'S37': (f76, 37, 19),
    'T37': (f77, 37, 20),
    'U37': (f78, 37, 21),
    'V37': (f79, 37, 22),
    'W37': (f80, 37, 23),
    'X37': (f81, 37, 24),
    'Y37': (f82, 37, 25),
    'Z37': (f83, 37, 26),
    'AA37': (f84, 37, 27),
    'AB37': (f85, 37, 28),
    'AC37': (f86, 37, 29),
    'AD37': (f87, 37, 30),
    'AE37': (f88, 37, 31),
    'AF37': (f89, 37, 32),
    'C38': (f95, 38, 3),
    'C42': (f96, 42, 3),
    'M42': (f97, 42, 13),
    'N42': (f98, 42, 14),
    'O42': (f99, 42, 15),
    'C43': (f100, 43, 3),
    'M43': (f101, 43, 13),
    'N43': (f102, 43, 14),
    'O43': (f103, 43, 15),
    'C44': (f104, 44, 3),
    'N44': (f105, 44, 14),
    'C45': (f106, 45, 3),
    'N45': (f107, 45, 14),
    'C46': (f108, 46, 3),
    'N46': (f109, 46, 14),
    'C47': (f110, 47, 3),
    'N47': (f111, 47, 14),
    'C48': (f112, 48, 3),
    'N48': (f113, 48, 14),
    'C49': (f114, 49, 3),
    'N49': (f115, 49, 14),
    'C50': (f116, 50, 3),
    'N50': (f117, 50, 14),
    'C51': (f118, 51, 3),
    'N51': (f119, 51, 14),
    'C52': (f120, 52, 3),
    'N52': (f121, 52, 14),
    'C53': (f122, 53, 3),
    'N53': (f123, 53, 14),
    'C56': (f124, 56, 3),
    'E62': (f125, 62, 5),
    'F62': (f125, 62, 6),
    'J62': (f126, 62, 10),
    'K62': (f127, 62, 11),
    'L62': (f128, 62, 12),
    'E63': (f129, 63, 5),
    'F63': (f130, 63, 6),
    'J63': (f126, 63, 10),
    'K63': (f131, 63, 11),
    'L63': (f128, 63, 12),
    'E64': (f132, 64, 5),
    'F64': (f132, 64, 6),
    'J64': (f126, 64, 10),
    'K64': (f133, 64, 11),
    'L64': (f128, 64, 12),
    'E65': (f134, 65, 5),
    'F65': (f135, 65, 6),
    'J65': (f126, 65, 10),
    'K65': (f136, 65, 11),
    'L65': (f128, 65, 12),
    'E66': (f137, 66, 5),
    'F66': (f137, 66, 6),
    'J66': (f126, 66, 10),
    'K66': (f138, 66, 11),
    'L66': (f128, 66, 12),
    'E67': (f139, 67, 5),
    'F67': (f139, 67, 6),
    'J67': (f126, 67, 10),
    'K67': (f140, 67, 11),
    'L67': (f128, 67, 12),
    'E68': (f141, 68, 5),
    'F68': (f141, 68, 6),
    'J68': (f126, 68, 10),
    'K68': (f142, 68, 11),
    'L68': (f128, 68, 12),
    'E69': (f143, 69, 5),
    'F69': (f143, 69, 6),
    'J69': (f126, 69, 10),
    'K69': (f144, 69, 11),
    'L69': (f128, 69, 12),
    'E70': (f145, 70, 5),
    'F70': (f145, 70, 6),
    'J70': (f126, 70, 10),
    'K70': (f146, 70, 11),
    'L70': (f128, 70, 12),
    'E71': (f147, 71, 5),
    'F71': (f147, 71, 6),
    'J71': (f126, 71, 10),
    'K71': (f148, 71, 11),
    'L71': (f128, 71, 12),
    'E72': (f149, 72, 5),
    'F72': (f149, 72, 6),
    'J72': (f126, 72, 10),
    'K72': (f150, 72, 11),
    'L72': (f128, 72, 12),
    'E73': (f151, 73, 5),
    'F73': (f151, 73, 6),
    'J73': (f126, 73, 10),
    'K73': (f152, 73, 11),
    'L73': (f128, 73, 12),
    'E74': (f153, 74, 5),
    'F74': (f153, 74, 6),
    'J74': (f126, 74, 10),
    'K74': (f154, 74, 11),
    'L74': (f128, 74, 12),
    'E75': (f155, 75, 5),
    'F75': (f156, 75, 6),
    'J75': (f126, 75, 10),
    'K75': (f157, 75, 11),
    'L75': (f128, 75, 12),
    'E76': (f158, 76, 5),
    'F76': (f159, 76, 6),
    'J76': (f126, 76, 10),
    'K76': (f160, 76, 11),
    'L76': (f128, 76, 12),
    'E77': (f161, 77, 5),
    'F77': (f161, 77, 6),
    'J77': (f126, 77, 10),
    'K77': (f162, 77, 11),
    'L77': (f128, 77, 12),
    'E78': (f163, 78, 5),
    'F78': (f163, 78, 6),
    'J78': (f126, 78, 10),
    'K78': (f164, 78, 11),
    'L78': (f128, 78, 12),
    'E79': (f165, 79, 5),
    'F79': (f166, 79, 6),
    'J79': (f126, 79, 10),
    'K79': (f167, 79, 11),
    'L79': (f128, 79, 12),
    'E80': (f168, 80, 5),
    'F80': (f168, 80, 6),
    'J80': (f126, 80, 10),
    'K80': (f169, 80, 11),
    'L80': (f128, 80, 12),
    'E81': (f170, 81, 5),
    'F81': (f170, 81, 6),
    'J81': (f126, 81, 10),
    'K81': (f171, 81, 11),
    'L81': (f128, 81, 12),
    'E82': (f172, 82, 5),
    'F82': (f172, 82, 6),
    'J82': (f126, 82, 10),
    'K82': (f173, 82, 11),
    'L82': (f128, 82, 12),
    'E83': (f174, 83, 5),
    'F83': (f174, 83, 6),
    'J83': (f126, 83, 10),
    'K83': (f175, 83, 11),
    'L83': (f128, 83, 12),
    'E84': (f176, 84, 5),
    'F84': (f177, 84, 6),
    'J84': (f126, 84, 10),
    'K84': (f178, 84, 11),
    'L84': (f128, 84, 12),
    'E85': (f179, 85, 5),
    'F85': (f180, 85, 6),
    'J85': (f126, 85, 10),
    'K85': (f181, 85, 11),
    'L85': (f128, 85, 12),
    'E86': (f182, 86, 5),
    'F86': (f182, 86, 6),
    'J86': (f126, 86, 10),
    'K86': (f183, 86, 11),
    'L86': (f128, 86, 12),
    'E87': (f184, 87, 5),
    'F87': (f184, 87, 6),
    'J87': (f126, 87, 10),
    'K87': (f185, 87, 11),
    'L87': (f128, 87, 12),
    'E88': (f186, 88, 5),
    'F88': (f186, 88, 6),
    'J88': (f126, 88, 10),
    'K88': (f187, 88, 11),
    'L88': (f128, 88, 12),
    'E89': (f188, 89, 5),
    'F89': (f188, 89, 6),
    'J89': (f126, 89, 10),
    'K89': (f189, 89, 11),
    'L89': (f128, 89, 12),
    'E90': (f190, 90, 5),
    'F90': (f190, 90, 6),
    'J90': (f126, 90, 10),
    'K90': (f191, 90, 11),
    'L90': (f128, 90, 12),
    'E91': (f192, 91, 5),
    'F91': (f192, 91, 6),
    'J91': (f126, 91, 10),
    'K91': (f193, 91, 11),
    'L91': (f128, 91, 12),
    'E92': (f194, 92, 5),
    'F92': (f194, 92, 6),
    'J92': (f126, 92, 10),
    'K92': (f195, 92, 11),
    'L92': (f128, 92, 12),
    'E96': (f196, 96, 5),
    'F96': (f196, 96, 6),
    'J96': (f126, 96, 10),
    'K96': (f197, 96, 11),
    'L96': (f128, 96, 12),
    'E97': (f198, 97, 5),
    'F97': (f198, 97, 6),
    'J97': (f126, 97, 10),
    'K97': (f199, 97, 11),
    'L97': (f128, 97, 12),
    'E98': (f200, 98, 5),
    'F98': (f200, 98, 6),
    'J98': (f126, 98, 10),
    'K98': (f201, 98, 11),
    'L98': (f128, 98, 12),
    'E99': (f202, 99, 5),
    'F99': (f202, 99, 6),
    'J99': (f126, 99, 10),
    'K99': (f203, 99, 11),
    'L99': (f128, 99, 12),
    'E100': (f204, 100, 5),
    'F100': (f204, 100, 6),
    'J100': (f126, 100, 10),
    'K100': (f205, 100, 11),
    'L100': (f128, 100, 12),
    'E101': (f206, 101, 5),
    'F101': (f206, 101, 6),
    'J101': (f126, 101, 10),
    'K101': (f207, 101, 11),
    'L101': (f128, 101, 12),
    'E102': (f208, 102, 5),
    'F102': (f208, 102, 6),
    'J102': (f126, 102, 10),
    'K102': (f209, 102, 11),
    'L102': (f128, 102, 12),
    'E103': (f210, 103, 5),
    'F103': (f210, 103, 6),
    'J103': (f126, 103, 10),
    'K103': (f211, 103, 11),
    'L103': (f128, 103, 12),
    'E104': (f212, 104, 5),
    'F104': (f212, 104, 6),
    'J104': (f126, 104, 10),
    'K104': (f213, 104, 11),
    'L104': (f128, 104, 12),
    'E105': (f214, 105, 5),
    'F105': (f214, 105, 6),
    'J105': (f126, 105, 10),
    'K105': (f215, 105, 11),
    'L105': (f128, 105, 12),
    'E106': (f216, 106, 5),
    'F106': (f216, 106, 6),
    'J106': (f126, 106, 10),
    'K106': (f217, 106, 11),
    'L106': (f128, 106, 12),
    'E107': (f218, 107, 5),
    'F107': (f218, 107, 6),
    'J107': (f126, 107, 10),
    'K107': (f219, 107, 11),
    'L107': (f128, 107, 12),
    'E108': (f220, 108, 5),
    'F108': (f220, 108, 6),
    'J108': (f126, 108, 10),
    'K108': (f221, 108, 11),
    'L108': (f128, 108, 12),
    'E109': (f222, 109, 5),
    'F109': (f222, 109, 6),
    'J109': (f126, 109, 10),
    'K109': (f223, 109, 11),
    'L109': (f128, 109, 12),
    'E110': (f224, 110, 5),
    'F110': (f224, 110, 6),
    'J110': (f126, 110, 10),
    'K110': (f225, 110, 11),
    'L110': (f128, 110, 12),
    'E111': (f226, 111, 5),
    'F111': (f226, 111, 6),
    'J111': (f126, 111, 10),
    'K111': (f227, 111, 11),
    'L111': (f128, 111, 12),
    'E112': (f228, 112, 5),
    'F112': (f228, 112, 6),
    'J112': (f126, 112, 10),
    'K112': (f229, 112, 11),
    'L112': (f128, 112, 12),
    'E113': (f230, 113, 5),
    'F113': (f230, 113, 6),
    'J113': (f126, 113, 10),
    'K113': (f231, 113, 11),
    'L113': (f128, 113, 12),
    'E114': (f232, 114, 5),
    'F114': (f232, 114, 6),
    'J114': (f126, 114, 10),
    'K114': (f233, 114, 11),
    'L114': (f128, 114, 12),
    'E115': (f234, 115, 5),
    'F115': (f234, 115, 6),
    'J115': (f126, 115, 10),
    'K115': (f235, 115, 11),
    'L115': (f128, 115, 12),
    'E116': (f236, 116, 5),
    'F116': (f236, 116, 6),
    'J116': (f126, 116, 10),
    'K116': (f237, 116, 11),
    'L116': (f128, 116, 12),
    'E117': (f238, 117, 5),
    'F117': (f238, 117, 6),
    'J117': (f126, 117, 10),
    'K117': (f239, 117, 11),
    'L117': (f128, 117, 12),
    'E118': (f240, 118, 5),
    'F118': (f240, 118, 6),
    'J118': (f126, 118, 10),
    'K118': (f241, 118, 11),
    'L118': (f128, 118, 12),
    'E119': (f242, 119, 5),
    'F119': (f242, 119, 6),
    'J119': (f126, 119, 10),
    'K119': (f243, 119, 11),
    'L119': (f128, 119, 12),
    'E120': (f244, 120, 5),
    'F120': (f244, 120, 6),
    'J120': (f126, 120, 10),
    'K120': (f245, 120, 11),
    'L120': (f128, 120, 12),
    'E121': (f246, 121, 5),
    'F121': (f246, 121, 6),
    'J121': (f126, 121, 10),
    'K121': (f247, 121, 11),
    'L121': (f128, 121, 12),
    'E122': (f248, 122, 5),
    'F122': (f248, 122, 6),
    'J122': (f126, 122, 10),
    'K122': (f249, 122, 11),
    'L122': (f128, 122, 12),
    'E123': (f250, 123, 5),
    'F123': (f250, 123, 6),
    'J123': (f126, 123, 10),
    'K123': (f251, 123, 11),
    'L123': (f128, 123, 12),
    'E124': (f252, 124, 5),
    'F124': (f252, 124, 6),
    'J124': (f126, 124, 10),
    'K124': (f253, 124, 11),
    'L124': (f128, 124, 12),
    'E125': (f254, 125, 5),
    'F125': (f254, 125, 6),
    'J125': (f126, 125, 10),
    'K125': (f255, 125, 11),
    'L125': (f128, 125, 12),
    'E126': (f256, 126, 5),
    'F126': (f256, 126, 6),
    'J126': (f126, 126, 10),
    'K126': (f257, 126, 11),
    'L126': (f128, 126, 12),
    'E127': (f258, 127, 5),
    'F127': (f258, 127, 6),
    'J127': (f126, 127, 10),
    'K127': (f259, 127, 11),
    'L127': (f128, 127, 12),
    'E128': (f260, 128, 5),
    'F128': (f260, 128, 6),
    'J128': (f126, 128, 10),
    'K128': (f261, 128, 11),
    'L128': (f128, 128, 12),
    'E129': (f262, 129, 5),
    'F129': (f262, 129, 6),
    'J129': (f126, 129, 10),
    'K129': (f263, 129, 11),
    'L129': (f128, 129, 12),
    'E130': (f264, 130, 5),
    'F130': (f264, 130, 6),
    'J130': (f126, 130, 10),
    'K130': (f265, 130, 11),
    'L130': (f128, 130, 12),
    'E131': (f266, 131, 5),
    'F131': (f266, 131, 6),
    'J131': (f126, 131, 10),
    'K131': (f267, 131, 11),
    'L131': (f128, 131, 12),
    'E132': (f268, 132, 5),
    'F132': (f268, 132, 6),
    'J132': (f126, 132, 10),
    'K132': (f269, 132, 11),
    'L132': (f128, 132, 12),
    'E133': (f270, 133, 5),
    'F133': (f270, 133, 6),
    'J133': (f126, 133, 10),
    'K133': (f271, 133, 11),
    'L133': (f128, 133, 12),
    'C137': (f272, 137, 3),
    'D137': (f272, 137, 4),
    'E137': (f272, 137, 5),
    'F137': (f272, 137, 6),
    'G137': (f272, 137, 7),
    'H137': (f272, 137, 8),
    'I137': (f272, 137, 9),
    'J137': (f272, 137, 10),
    'C138': (f273, 138, 3),
    'D138': (f273, 138, 4),
    'E138': (f273, 138, 5),
    'F138': (f273, 138, 6),
    'G138': (f273, 138, 7),
    'H138': (f273, 138, 8),
    'I138': (f273, 138, 9),
    'J138': (f273, 138, 10),
    'C139': (f274, 139, 3),
    'D139': (f274, 139, 4),
    'E139': (f274, 139, 5),
    'F139': (f274, 139, 6),
    'G139': (f274, 139, 7),
    'H139': (f274, 139, 8),
    'I139': (f274, 139, 9),
    'J139': (f274, 139, 10),
    'C140': (f275, 140, 3),
    'D140': (f275, 140, 4),
    'E140': (f275, 140, 5),
    'F140': (f275, 140, 6),
    'G140': (f275, 140, 7),
    'H140': (f275, 140, 8),
    'I140': (f275, 140, 9),
    'J140': (f275, 140, 10),
    'C141': (f276, 141, 3),
    'D141': (f276, 141, 4),
    'E141': (f276, 141, 5),
    'F141': (f276, 141, 6),
    'G141': (f276, 141, 7),
    'H141': (f276, 141, 8),
    'I141': (f276, 141, 9),
    'J141': (f276, 141, 10),
    'C142': (f277, 142, 3),
    'D142': (f277, 142, 4),
    'E142': (f277, 142, 5),
    'F142': (f277, 142, 6),
    'G142': (f277, 142, 7),
    'H142': (f277, 142, 8),
    'I142': (f277, 142, 9),
    'J142': (f277, 142, 10),
    'C143': (f278, 143, 3),
    'D143': (f278, 143, 4),
    'E143': (f278, 143, 5),
    'F143': (f278, 143, 6),
    'G143': (f278, 143, 7),
    'H143': (f278, 143, 8),
    'I143': (f278, 143, 9),
    'J143': (f278, 143, 10),
    'C144': (f279, 144, 3),
    'D144': (f279, 144, 4),
    'E144': (f279, 144, 5),
    'F144': (f279, 144, 6),
    'G144': (f279, 144, 7),
    'H144': (f279, 144, 8),
    'I144': (f279, 144, 9),
    'J144': (f279, 144, 10),
    'B146': (f64, 146, 2),
    'C146': (f280, 146, 3),
    'D146': (f281, 146, 4),
    'E146': (f282, 146, 5),
    'K146': (f283, 146, 11),
    'L146': (f284, 146, 12),
    'B147': (f90, 147, 2),
    'C147': (f285, 147, 3),
    'D147': (f286, 147, 4),
    'E147': (f287, 147, 5),
    'K147': (f288, 147, 11),
    'L147': (f289, 147, 12),
    'B148': (f91, 148, 2),
    'C148': (f290, 148, 3),
    'D148': (f291, 148, 4),
    'E148': (f292, 148, 5),
    'K148': (f293, 148, 11),
    'L148': (f294, 148, 12),
    'B149': (f92, 149, 2),
    'C149': (f295, 149, 3),
    'D149': (f296, 149, 4),
    'E149': (f297, 149, 5),
    'K149': (f298, 149, 11),
    'L149': (f299, 149, 12),
    'B150': (f93, 150, 2),
    'C150': (f300, 150, 3),
    'D150': (f301, 150, 4),
    'E150': (f302, 150, 5),
    'K150': (f303, 150, 11),
    'L150': (f304, 150, 12),
    'B151': (f94, 151, 2),
    'C151': (f305, 151, 3),
    'D151': (f306, 151, 4),
    'E151': (f307, 151, 5),
    'K151': (f308, 151, 11),
    'L151': (f309, 151, 12),
    'B152': (f310, 152, 2),
    'C152': (f311, 152, 3),
    'D152': (f312, 152, 4),
    'E152': (f313, 152, 5),
    'K152': (f314, 152, 11),
    'L152': (f315, 152, 12),
    'B153': (f316, 153, 2),
    'C153': (f317, 153, 3),
    'D153': (f318, 153, 4),
    'E153': (f319, 153, 5),
    'K153': (f320, 153, 11),
    'L153': (f321, 153, 12),
    'B154': (f322, 154, 2),
    'C154': (f323, 154, 3),
    'D154': (f324, 154, 4),
    'E154': (f325, 154, 5),
    'K154': (f326, 154, 11),
    'L154': (f327, 154, 12),
    'B155': (f328, 155, 2),
    'C155': (f329, 155, 3),
    'D155': (f330, 155, 4),
    'E155': (f331, 155, 5),
    'K155': (f332, 155, 11),
    'L155': (f333, 155, 12),
    'B156': (f334, 156, 2),
    'C156': (f335, 156, 3),
    'D156': (f336, 156, 4),
    'E156': (f337, 156, 5),
    'K156': (f338, 156, 11),
    'L156': (f339, 156, 12),
    'C157': (f340, 157, 3),
    'C160': (f341, 160, 3),
    'C161': (f342, 161, 3),
    'C162': (f343, 162, 3),
    'C163': (f344, 163, 3),
    'C164': (f345, 164, 3),
    'AJ189': (f346, 189, 36),
    'AK189': (f347, 189, 37),
    'AL189': (f348, 189, 38),
    'AM189': (f349, 189, 39),
    'AJ190': (f350, 190, 36),
    'AK190': (f351, 190, 37),
    'AL190': (f352, 190, 38),
    'AM190': (f353, 190, 39),
    'AJ191': (f354, 191, 36),
    'AK191': (f355, 191, 37),
    'AL191': (f356, 191, 38),
    'AM191': (f357, 191, 39),
    'AJ192': (f358, 192, 36),
    'AK192': (f359, 192, 37),
    'AL192': (f360, 192, 38),
    'AM192': (f361, 192, 39),
    'AJ193': (f362, 193, 36),
    'AK193': (f363, 193, 37),
    'AL193': (f364, 193, 38),
    'AM193': (f365, 193, 39),
    'AJ194': (f366, 194, 36),
    'AK194': (f367, 194, 37),
    'AL194': (f368, 194, 38),
    'AM194': (f369, 194, 39),
    'AJ195': (f370, 195, 36),
    'AK195': (f371, 195, 37),
    'AL195': (f372, 195, 38),
    'AM195': (f373, 195, 39),
    'AJ196': (f374, 196, 36),
    'AK196': (f375, 196, 37),
    'AL196': (f376, 196, 38),
    'AM196': (f377, 196, 39),
    'AJ197': (f378, 197, 36),
    'AK197': (f347, 197, 37),
    'AL197': (f379, 197, 38),
    'AM197': (f380, 197, 39),
    'AJ198': (f381, 198, 36),
    'AK198': (f351, 198, 37),
    'AL198': (f382, 198, 38),
    'AM198': (f383, 198, 39),
    'AJ199': (f384, 199, 36),
    'AK199': (f355, 199, 37),
    'AL199': (f385, 199, 38),
    'AM199': (f386, 199, 39),
    'AJ200': (f387, 200, 36),
    'AK200': (f359, 200, 37),
    'AL200': (f388, 200, 38),
    'AM200': (f389, 200, 39),
    'AJ201': (f390, 201, 36),
    'AK201': (f363, 201, 37),
    'AL201': (f391, 201, 38),
    'AM201': (f392, 201, 39),
    'AJ202': (f393, 202, 36),
    'AK202': (f367, 202, 37),
    'AL202': (f394, 202, 38),
    'AM202': (f395, 202, 39),
    'AJ203': (f396, 203, 36),
    'AK203': (f371, 203, 37),
    'AL203': (f397, 203, 38),
    'AM203': (f398, 203, 39),
    'AJ204': (f399, 204, 36),
    'AK204': (f375, 204, 37),
    'AL204': (f400, 204, 38),
    'AM204': (f401, 204, 39),
    'AJ209': (f402, 209, 36),
    'AK209': (f403, 209, 37),
    'AJ210': (f404, 210, 36),
    'AK210': (f405, 210, 37),
    'AJ211': (f406, 211, 36),
    'AK211': (f407, 211, 37),
    'AJ212': (f408, 212, 36),
    'AK212': (f409, 212, 37),
    'AJ213': (f410, 213, 36),
    'AK213': (f411, 213, 37),
    'AJ214': (f412, 214, 36),
    'AK214': (f413, 214, 37),
    'E219': (f414, 219, 5),
    'F219': (f415, 219, 6),
    'J219': (f126, 219, 10),
    'K219': (f416, 219, 11),
    'L219': (f128, 219, 12),
    'E220': (f417, 220, 5),
    'F220': (f417, 220, 6),
    'J220': (f126, 220, 10),
    'K220': (f418, 220, 11),
    'L220': (f128, 220, 12),
    'E221': (f419, 221, 5),
    'F221': (f419, 221, 6),
    'J221': (f126, 221, 10),
    'K221': (f420, 221, 11),
    'L221': (f128, 221, 12),
    'E222': (f421, 222, 5),
    'F222': (f421, 222, 6),
    'J222': (f126, 222, 10),
    'K222': (f422, 222, 11),
    'L222': (f128, 222, 12),
    'E223': (f423, 223, 5),
    'F223': (f424, 223, 6),
    'J223': (f126, 223, 10),
    'K223': (f425, 223, 11),
    'L223': (f128, 223, 12),
    'E224': (f426, 224, 5),
    'F224': (f426, 224, 6),
    'J224': (f126, 224, 10),
    'K224': (f427, 224, 11),
    'L224': (f128, 224, 12),
}
