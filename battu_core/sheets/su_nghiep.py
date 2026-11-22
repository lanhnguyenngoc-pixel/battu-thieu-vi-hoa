# -*- coding: utf-8 -*-
"""Module sheet SU_NGHIEP – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.9.xlsx (SHA-256 5486ea8b3eb631dc513aed16150152722659ca3ca34dfc6c7af268da54b63453).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'SU_NGHIEP'
SI = 21

def f0(E, r, c):
    # Ô Z2
    # = IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$72,DUNG_HY_KY!$C$72)="",IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$92,DUNG_HY_KY!$C$92)>=1,IF(DOI_CHIEU!$C$33=2,DUNG
    #   _HY_KY!$D$92,DUNG_HY_KY!$C$92)<=3),MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$5,DUNG_HY_KY!$C$5)+IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$92,DUNG_HY_KY!$C$92),5
    #   ),-1)
    return (X.MOD(X.add((E.v(8,5,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,5,3)),(E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3))), 5.0) if X.truthy(X.AND(X.eq((E.v(8,72,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,72,3)),''), X.ge((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),1.0), X.le((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),3.0))) else X.neg(1.0))

def f1(E, r, c):
    # Ô C6, C7, C8, C9
    # = IF(TRIM(AB6&" "&AC6)="","(không có dấu hiệu thuận nào của sách đủ điều kiện)",TRIM(AB6&" "&AC6))
    return ('(không có dấu hiệu thuận nào của sách đủ điều kiện)' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(21,r,c+25),' '),E.v(21,r,c+26))),'')) else X.TRIM(X.cat(X.cat(E.v(21,r,c+25),' '),E.v(21,r,c+26))))

def f2(E, r, c):
    # Ô F6, F7, F8, F9
    # = IF(TRIM(AD6&" "&AE6)="","(không có dấu hiệu bất lợi nào của sách đủ điều kiện)",TRIM(AD6&" "&AE6))
    return ('(không có dấu hiệu bất lợi nào của sách đủ điều kiện)' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(E.v(21,r,c+24),' '),E.v(21,r,c+25))),'')) else X.TRIM(X.cat(X.cat(E.v(21,r,c+24),' '),E.v(21,r,c+25))))

def f3(E, r, c):
    # Ô J6
    # = IF(IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$
    #   D$110,DUNG_HY_KY!$C$110)),1,IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN
    #   !$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_
    #   KY!$D$110,DUNG_HY_KY!$C$110))),-1,0))=1,IF(M6>=N6,"THUẬN","THUẬN NHƯNG CÓ TRỞ NGẠI"),IF(IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY
    #   !$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),-1,0))=-1,IF(N6>=M6,"BẤT LỢI","BẤT
    #    LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M6>N6,"NGHIÊNG THUẬN",IF(N6>M6,"NGHIÊNG BẤT LỢI","TRUNG BÌNH"))))&CHAR(10)&"("&M6&" dấu hiệu thuận, "&N6&" bất lợi; yếu t
    #   ố then chốt: "&"tài "&IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)
    #   =IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!
    #   $C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109
    #   ),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","kh
    #   ông phải dụng/hỷ/kỵ thần (nhàn thần)"))))&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(21,r,c+3),E.v(21,r,c+4))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else 0.0)),1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(21,r,c+4),E.v(21,r,c+3))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else 0.0)),X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(21,r,c+3),E.v(21,r,c+4))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(21,r,c+4),E.v(21,r,c+3))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(21,r,c+3)),' dấu hiệu thuận, '),E.v(21,r,c+4)),' bất lợi; yếu tố then chốt: '),'tài '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),')')

def f4(E, r, c):
    # Ô M6
    # = ((L62="ĐỦ")+(L63="ĐỦ")+(L64="ĐỦ")+(L65="ĐỦ")+(L66="ĐỦ")+(L67="ĐỦ")+(L68="ĐỦ")+(L69="ĐỦ")+(L70="ĐỦ")+(L71="ĐỦ")+(L72="ĐỦ")+(L73="ĐỦ")+(L74="ĐỦ")+(L75="
    #   ĐỦ")+(L76="ĐỦ")+(L77="ĐỦ")+(L78="ĐỦ")+(L79="ĐỦ")+(L80="ĐỦ")+(L81="ĐỦ")+(L83="ĐỦ"))+((L246="ĐỦ")+(L247="ĐỦ")+(L248="ĐỦ")+(L249="ĐỦ")+(L252="ĐỦ")+(L254=
    #   "ĐỦ")+(L255="ĐỦ")+(L256="ĐỦ")+(L258="ĐỦ")+(L259="ĐỦ")+(L260="ĐỦ")+(L261="ĐỦ")+(L262="ĐỦ")+(L263="ĐỦ")+(L264="ĐỦ")+(L265="ĐỦ")+(L266="ĐỦ")+(L267="ĐỦ")+
    #   (L268="ĐỦ")+(L269="ĐỦ")+(L270="ĐỦ")+(L272="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+56,c-1),'ĐỦ'),X.eq(E.v(21,r+57,c-1),'ĐỦ')),X.eq(E.v(21,r+58,c-1),'ĐỦ')),X.eq(E.v(21,r+59,c-1),'ĐỦ')),X.eq(E.v(21,r+60,c-1),'ĐỦ')),X.eq(E.v(21,r+61,c-1),'ĐỦ')),X.eq(E.v(21,r+62,c-1),'ĐỦ')),X.eq(E.v(21,r+63,c-1),'ĐỦ')),X.eq(E.v(21,r+64,c-1),'ĐỦ')),X.eq(E.v(21,r+65,c-1),'ĐỦ')),X.eq(E.v(21,r+66,c-1),'ĐỦ')),X.eq(E.v(21,r+67,c-1),'ĐỦ')),X.eq(E.v(21,r+68,c-1),'ĐỦ')),X.eq(E.v(21,r+69,c-1),'ĐỦ')),X.eq(E.v(21,r+70,c-1),'ĐỦ')),X.eq(E.v(21,r+71,c-1),'ĐỦ')),X.eq(E.v(21,r+72,c-1),'ĐỦ')),X.eq(E.v(21,r+73,c-1),'ĐỦ')),X.eq(E.v(21,r+74,c-1),'ĐỦ')),X.eq(E.v(21,r+75,c-1),'ĐỦ')),X.eq(E.v(21,r+77,c-1),'ĐỦ')),X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+240,c-1),'ĐỦ'),X.eq(E.v(21,r+241,c-1),'ĐỦ')),X.eq(E.v(21,r+242,c-1),'ĐỦ')),X.eq(E.v(21,r+243,c-1),'ĐỦ')),X.eq(E.v(21,r+246,c-1),'ĐỦ')),X.eq(E.v(21,r+248,c-1),'ĐỦ')),X.eq(E.v(21,r+249,c-1),'ĐỦ')),X.eq(E.v(21,r+250,c-1),'ĐỦ')),X.eq(E.v(21,r+252,c-1),'ĐỦ')),X.eq(E.v(21,r+253,c-1),'ĐỦ')),X.eq(E.v(21,r+254,c-1),'ĐỦ')),X.eq(E.v(21,r+255,c-1),'ĐỦ')),X.eq(E.v(21,r+256,c-1),'ĐỦ')),X.eq(E.v(21,r+257,c-1),'ĐỦ')),X.eq(E.v(21,r+258,c-1),'ĐỦ')),X.eq(E.v(21,r+259,c-1),'ĐỦ')),X.eq(E.v(21,r+260,c-1),'ĐỦ')),X.eq(E.v(21,r+261,c-1),'ĐỦ')),X.eq(E.v(21,r+262,c-1),'ĐỦ')),X.eq(E.v(21,r+263,c-1),'ĐỦ')),X.eq(E.v(21,r+264,c-1),'ĐỦ')),X.eq(E.v(21,r+266,c-1),'ĐỦ')))

def f5(E, r, c):
    # Ô N6
    # = ((L86="ĐỦ")+(L87="ĐỦ")+(L88="ĐỦ")+(L89="ĐỦ")+(L90="ĐỦ")+(L91="ĐỦ")+(L92="ĐỦ")+(L93="ĐỦ")+(L94="ĐỦ")+(L95="ĐỦ")+(L96="ĐỦ")+(L97="ĐỦ"))+((L273="ĐỦ")+(L2
    #   74="ĐỦ")+(L276="ĐỦ")+(L278="ĐỦ")+(L279="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+80,c-2),'ĐỦ'),X.eq(E.v(21,r+81,c-2),'ĐỦ')),X.eq(E.v(21,r+82,c-2),'ĐỦ')),X.eq(E.v(21,r+83,c-2),'ĐỦ')),X.eq(E.v(21,r+84,c-2),'ĐỦ')),X.eq(E.v(21,r+85,c-2),'ĐỦ')),X.eq(E.v(21,r+86,c-2),'ĐỦ')),X.eq(E.v(21,r+87,c-2),'ĐỦ')),X.eq(E.v(21,r+88,c-2),'ĐỦ')),X.eq(E.v(21,r+89,c-2),'ĐỦ')),X.eq(E.v(21,r+90,c-2),'ĐỦ')),X.eq(E.v(21,r+91,c-2),'ĐỦ')),X.add(X.add(X.add(X.add(X.eq(E.v(21,r+267,c-2),'ĐỦ'),X.eq(E.v(21,r+268,c-2),'ĐỦ')),X.eq(E.v(21,r+270,c-2),'ĐỦ')),X.eq(E.v(21,r+272,c-2),'ĐỦ')),X.eq(E.v(21,r+273,c-2),'ĐỦ')))

def f6(E, r, c):
    # Ô AB6
    # = TRIM(IF(L62="ĐỦ","• "&G62&" ["&B62&", "&H62&"] ","")&IF(L63="ĐỦ","• "&G63&" ["&B63&", "&H63&"] ","")&IF(L64="ĐỦ","• "&G64&" ["&B64&", "&H64&"] ","")&I
    #   F(L65="ĐỦ","• "&G65&" ["&B65&", "&H65&"] ","")&IF(L66="ĐỦ","• "&G66&" ["&B66&", "&H66&"] ","")&IF(L67="ĐỦ","• "&G67&" ["&B67&", "&H67&"] ","")&IF(L68=
    #   "ĐỦ","• "&G68&" ["&B68&", "&H68&"] ","")&IF(L69="ĐỦ","• "&G69&" ["&B69&", "&H69&"] ","")&IF(L70="ĐỦ","• "&G70&" ["&B70&", "&H70&"] ","")&IF(L71="ĐỦ","
    #   • "&G71&" ["&B71&", "&H71&"] ","")&IF(L72="ĐỦ","• "&G72&" ["&B72&", "&H72&"] ","")&IF(L73="ĐỦ","• "&G73&" ["&B73&", "&H73&"] ","")&IF(L74="ĐỦ","• "&G7
    #   4&" ["&B74&", "&H74&"] ","")&IF(L75="ĐỦ","• "&G75&" ["&B75&", "&H75&"] ","")&IF(L76="ĐỦ","• "&G76&" ["&B76&", "&H76&"] ","")&IF(L77="ĐỦ","• "&G77&" ["
    #   &B77&", "&H77&"] ","")&IF(L78="ĐỦ","• "&G78&" ["&B78&", "&H78&"] ","")&IF(L79="ĐỦ","• "&G79&" ["&B79&", "&H79&"] ","")&IF(L80="ĐỦ","• "&G80&" ["&B80&"
    #   , "&H80&"] ","")&IF(L81="ĐỦ","• "&G81&" ["&B81&", "&H81&"] ","")&IF(L83="ĐỦ","• "&G83&" ["&B83&", "&H83&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+56,c-21)),' ['),E.v(21,r+56,c-26)),', '),E.v(21,r+56,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+56,c-16),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+57,c-21)),' ['),E.v(21,r+57,c-26)),', '),E.v(21,r+57,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+57,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+58,c-21)),' ['),E.v(21,r+58,c-26)),', '),E.v(21,r+58,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+58,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+59,c-21)),' ['),E.v(21,r+59,c-26)),', '),E.v(21,r+59,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+59,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+60,c-21)),' ['),E.v(21,r+60,c-26)),', '),E.v(21,r+60,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+60,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+61,c-21)),' ['),E.v(21,r+61,c-26)),', '),E.v(21,r+61,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+61,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+62,c-21)),' ['),E.v(21,r+62,c-26)),', '),E.v(21,r+62,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+62,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+63,c-21)),' ['),E.v(21,r+63,c-26)),', '),E.v(21,r+63,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+63,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+64,c-21)),' ['),E.v(21,r+64,c-26)),', '),E.v(21,r+64,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+64,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+65,c-21)),' ['),E.v(21,r+65,c-26)),', '),E.v(21,r+65,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+65,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+66,c-21)),' ['),E.v(21,r+66,c-26)),', '),E.v(21,r+66,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+66,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+67,c-21)),' ['),E.v(21,r+67,c-26)),', '),E.v(21,r+67,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+67,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+68,c-21)),' ['),E.v(21,r+68,c-26)),', '),E.v(21,r+68,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+68,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+69,c-21)),' ['),E.v(21,r+69,c-26)),', '),E.v(21,r+69,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+69,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+70,c-21)),' ['),E.v(21,r+70,c-26)),', '),E.v(21,r+70,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+70,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+71,c-21)),' ['),E.v(21,r+71,c-26)),', '),E.v(21,r+71,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+71,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+72,c-21)),' ['),E.v(21,r+72,c-26)),', '),E.v(21,r+72,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+72,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+73,c-21)),' ['),E.v(21,r+73,c-26)),', '),E.v(21,r+73,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+73,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+74,c-21)),' ['),E.v(21,r+74,c-26)),', '),E.v(21,r+74,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+74,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+75,c-21)),' ['),E.v(21,r+75,c-26)),', '),E.v(21,r+75,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+75,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+77,c-21)),' ['),E.v(21,r+77,c-26)),', '),E.v(21,r+77,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+77,c-16),'ĐỦ')) else '')))

def f7(E, r, c):
    # Ô AC6
    # = TRIM(IF(L246="ĐỦ","• "&G246&" ["&B246&", "&H246&"] ","")&IF(L247="ĐỦ","• "&G247&" ["&B247&", "&H247&"] ","")&IF(L248="ĐỦ","• "&G248&" ["&B248&", "&H24
    #   8&"] ","")&IF(L249="ĐỦ","• "&G249&" ["&B249&", "&H249&"] ","")&IF(L252="ĐỦ","• "&G252&" ["&B252&", "&H252&"] ","")&IF(L254="ĐỦ","• "&G254&" ["&B254&",
    #    "&H254&"] ","")&IF(L255="ĐỦ","• "&G255&" ["&B255&", "&H255&"] ","")&IF(L256="ĐỦ","• "&G256&" ["&B256&", "&H256&"] ","")&IF(L258="ĐỦ","• "&G258&" ["&B
    #   258&", "&H258&"] ","")&IF(L259="ĐỦ","• "&G259&" ["&B259&", "&H259&"] ","")&IF(L260="ĐỦ","• "&G260&" ["&B260&", "&H260&"] ","")&IF(L261="ĐỦ","• "&G261&
    #   " ["&B261&", "&H261&"] ","")&IF(L262="ĐỦ","• "&G262&" ["&B262&", "&H262&"] ","")&IF(L263="ĐỦ","• "&G263&" ["&B263&", "&H263&"] ","")&IF(L264="ĐỦ","• "
    #   &G264&" ["&B264&", "&H264&"] ","")&IF(L265="ĐỦ","• "&G265&" ["&B265&", "&H265&"] ","")&IF(L266="ĐỦ","• "&G266&" ["&B266&", "&H266&"] ","")&IF(L267="ĐỦ
    #   ","• "&G267&" ["&B267&", "&H267&"] ","")&IF(L268="ĐỦ","• "&G268&" ["&B268&", "&H268&"] ","")&IF(L269="ĐỦ","• "&G269&" ["&B269&", "&H269&"] ","")&IF(L2
    #   70="ĐỦ","• "&G270&" ["&B270&", "&H270&"] ","")&IF(L272="ĐỦ","• "&G272&" ["&B272&", "&H272&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+240,c-22)),' ['),E.v(21,r+240,c-27)),', '),E.v(21,r+240,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+240,c-17),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+241,c-22)),' ['),E.v(21,r+241,c-27)),', '),E.v(21,r+241,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+241,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+242,c-22)),' ['),E.v(21,r+242,c-27)),', '),E.v(21,r+242,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+242,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+243,c-22)),' ['),E.v(21,r+243,c-27)),', '),E.v(21,r+243,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+243,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+246,c-22)),' ['),E.v(21,r+246,c-27)),', '),E.v(21,r+246,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+246,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+248,c-22)),' ['),E.v(21,r+248,c-27)),', '),E.v(21,r+248,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+248,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+249,c-22)),' ['),E.v(21,r+249,c-27)),', '),E.v(21,r+249,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+249,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+250,c-22)),' ['),E.v(21,r+250,c-27)),', '),E.v(21,r+250,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+250,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+252,c-22)),' ['),E.v(21,r+252,c-27)),', '),E.v(21,r+252,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+252,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+253,c-22)),' ['),E.v(21,r+253,c-27)),', '),E.v(21,r+253,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+253,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+254,c-22)),' ['),E.v(21,r+254,c-27)),', '),E.v(21,r+254,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+254,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+255,c-22)),' ['),E.v(21,r+255,c-27)),', '),E.v(21,r+255,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+255,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+256,c-22)),' ['),E.v(21,r+256,c-27)),', '),E.v(21,r+256,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+256,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+257,c-22)),' ['),E.v(21,r+257,c-27)),', '),E.v(21,r+257,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+257,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+258,c-22)),' ['),E.v(21,r+258,c-27)),', '),E.v(21,r+258,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+258,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+259,c-22)),' ['),E.v(21,r+259,c-27)),', '),E.v(21,r+259,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+259,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+260,c-22)),' ['),E.v(21,r+260,c-27)),', '),E.v(21,r+260,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+260,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+261,c-22)),' ['),E.v(21,r+261,c-27)),', '),E.v(21,r+261,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+261,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+262,c-22)),' ['),E.v(21,r+262,c-27)),', '),E.v(21,r+262,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+262,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+263,c-22)),' ['),E.v(21,r+263,c-27)),', '),E.v(21,r+263,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+263,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+264,c-22)),' ['),E.v(21,r+264,c-27)),', '),E.v(21,r+264,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+264,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+266,c-22)),' ['),E.v(21,r+266,c-27)),', '),E.v(21,r+266,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+266,c-17),'ĐỦ')) else '')))

def f8(E, r, c):
    # Ô AD6
    # = TRIM(IF(L86="ĐỦ","• "&G86&" ["&B86&", "&H86&"] ","")&IF(L87="ĐỦ","• "&G87&" ["&B87&", "&H87&"] ","")&IF(L88="ĐỦ","• "&G88&" ["&B88&", "&H88&"] ","")&I
    #   F(L89="ĐỦ","• "&G89&" ["&B89&", "&H89&"] ","")&IF(L90="ĐỦ","• "&G90&" ["&B90&", "&H90&"] ","")&IF(L91="ĐỦ","• "&G91&" ["&B91&", "&H91&"] ","")&IF(L92=
    #   "ĐỦ","• "&G92&" ["&B92&", "&H92&"] ","")&IF(L93="ĐỦ","• "&G93&" ["&B93&", "&H93&"] ","")&IF(L94="ĐỦ","• "&G94&" ["&B94&", "&H94&"] ","")&IF(L95="ĐỦ","
    #   • "&G95&" ["&B95&", "&H95&"] ","")&IF(L96="ĐỦ","• "&G96&" ["&B96&", "&H96&"] ","")&IF(L97="ĐỦ","• "&G97&" ["&B97&", "&H97&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+80,c-23)),' ['),E.v(21,r+80,c-28)),', '),E.v(21,r+80,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+80,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+81,c-23)),' ['),E.v(21,r+81,c-28)),', '),E.v(21,r+81,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+81,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+82,c-23)),' ['),E.v(21,r+82,c-28)),', '),E.v(21,r+82,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+82,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+83,c-23)),' ['),E.v(21,r+83,c-28)),', '),E.v(21,r+83,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+83,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+84,c-23)),' ['),E.v(21,r+84,c-28)),', '),E.v(21,r+84,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+84,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+85,c-23)),' ['),E.v(21,r+85,c-28)),', '),E.v(21,r+85,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+85,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+86,c-23)),' ['),E.v(21,r+86,c-28)),', '),E.v(21,r+86,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+86,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+87,c-23)),' ['),E.v(21,r+87,c-28)),', '),E.v(21,r+87,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+87,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+88,c-23)),' ['),E.v(21,r+88,c-28)),', '),E.v(21,r+88,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+88,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+89,c-23)),' ['),E.v(21,r+89,c-28)),', '),E.v(21,r+89,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+89,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+90,c-23)),' ['),E.v(21,r+90,c-28)),', '),E.v(21,r+90,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+90,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+91,c-23)),' ['),E.v(21,r+91,c-28)),', '),E.v(21,r+91,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+91,c-18),'ĐỦ')) else '')))

def f9(E, r, c):
    # Ô AE6
    # = TRIM(IF(L273="ĐỦ","• "&G273&" ["&B273&", "&H273&"] ","")&IF(L274="ĐỦ","• "&G274&" ["&B274&", "&H274&"] ","")&IF(L276="ĐỦ","• "&G276&" ["&B276&", "&H27
    #   6&"] ","")&IF(L278="ĐỦ","• "&G278&" ["&B278&", "&H278&"] ","")&IF(L279="ĐỦ","• "&G279&" ["&B279&", "&H279&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+267,c-24)),' ['),E.v(21,r+267,c-29)),', '),E.v(21,r+267,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+267,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+268,c-24)),' ['),E.v(21,r+268,c-29)),', '),E.v(21,r+268,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+268,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+270,c-24)),' ['),E.v(21,r+270,c-29)),', '),E.v(21,r+270,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+270,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+272,c-24)),' ['),E.v(21,r+272,c-29)),', '),E.v(21,r+272,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+272,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+273,c-24)),' ['),E.v(21,r+273,c-29)),', '),E.v(21,r+273,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+273,c-19),'ĐỦ')) else '')))

def f10(E, r, c):
    # Ô J7
    # = IF(AND(LEFT(AA7,7)="BẤT LỢI",L100="ĐỦ",M7>=2*N7,M7>=4),"TRUNG BÌNH – SÁT CÓ CHẾ"&MID(AA7,FIND(CHAR(10),AA7&CHAR(10)),2000)&CHAR(10)&"(quan sát là kỵ t
    #   hần nhưng đã bị chế ngự – “thất sát có chế ngự là quân tử” (P2 tr.535) – và dấu hiệu thuận gấp đôi bất lợi trở lên → không xếp BẤT LỢI; giả định G26)"
    #   ,AA7)
    return (X.cat(X.cat(X.cat('TRUNG BÌNH – SÁT CÓ CHẾ',X.MID(E.v(21,r,c+17), X.FIND(X.CHAR(10.0), X.cat(E.v(21,r,c+17),X.CHAR(10.0))), 2000.0)),X.CHAR(10.0)),'(quan sát là kỵ thần nhưng đã bị chế ngự – “thất sát có chế ngự là quân tử” (P2 tr.535) – và dấu hiệu thuận gấp đôi bất lợi trở lên → không xếp BẤT LỢI; giả định G26)') if X.truthy(X.AND(X.eq(X.LEFT(E.v(21,r,c+17), 7.0),'BẤT LỢI'), X.eq(E.v(21,r+93,c+2),'ĐỦ'), X.ge(E.v(21,r,c+3),X.mul(2.0,E.v(21,r,c+4))), X.ge(E.v(21,r,c+3),4.0))) else E.v(21,r,c+17))

def f11(E, r, c):
    # Ô M7
    # = ((L98="ĐỦ")+(L99="ĐỦ")+(L100="ĐỦ")+(L101="ĐỦ")+(L102="ĐỦ")+(L103="ĐỦ")+(L104="ĐỦ")+(L105="ĐỦ")+(L106="ĐỦ")+(L107="ĐỦ")+(L108="ĐỦ")+(L109="ĐỦ")+(L111="
    #   ĐỦ"))+((L160="ĐỦ")+(L161="ĐỦ")+(L162="ĐỦ")+(L163="ĐỦ")+(L164="ĐỦ")+(L165="ĐỦ")+(L166="ĐỦ")+(L167="ĐỦ")+(L168="ĐỦ")+(L169="ĐỦ")+(L171="ĐỦ")+(L172="ĐỦ")
    #   +(L173="ĐỦ")+(L174="ĐỦ")+(L175="ĐỦ")+(L176="ĐỦ")+(L178="ĐỦ")+(L179="ĐỦ")+(L180="ĐỦ")+(L181="ĐỦ")+(L182="ĐỦ")+(L183="ĐỦ")+(L185="ĐỦ")+(L186="ĐỦ")+(L187
    #   ="ĐỦ")+(L188="ĐỦ")+(L190="ĐỦ")+(L191="ĐỦ")+(L192="ĐỦ")+(L194="ĐỦ")+(L195="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+91,c-1),'ĐỦ'),X.eq(E.v(21,r+92,c-1),'ĐỦ')),X.eq(E.v(21,r+93,c-1),'ĐỦ')),X.eq(E.v(21,r+94,c-1),'ĐỦ')),X.eq(E.v(21,r+95,c-1),'ĐỦ')),X.eq(E.v(21,r+96,c-1),'ĐỦ')),X.eq(E.v(21,r+97,c-1),'ĐỦ')),X.eq(E.v(21,r+98,c-1),'ĐỦ')),X.eq(E.v(21,r+99,c-1),'ĐỦ')),X.eq(E.v(21,r+100,c-1),'ĐỦ')),X.eq(E.v(21,r+101,c-1),'ĐỦ')),X.eq(E.v(21,r+102,c-1),'ĐỦ')),X.eq(E.v(21,r+104,c-1),'ĐỦ')),X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+153,c-1),'ĐỦ'),X.eq(E.v(21,r+154,c-1),'ĐỦ')),X.eq(E.v(21,r+155,c-1),'ĐỦ')),X.eq(E.v(21,r+156,c-1),'ĐỦ')),X.eq(E.v(21,r+157,c-1),'ĐỦ')),X.eq(E.v(21,r+158,c-1),'ĐỦ')),X.eq(E.v(21,r+159,c-1),'ĐỦ')),X.eq(E.v(21,r+160,c-1),'ĐỦ')),X.eq(E.v(21,r+161,c-1),'ĐỦ')),X.eq(E.v(21,r+162,c-1),'ĐỦ')),X.eq(E.v(21,r+164,c-1),'ĐỦ')),X.eq(E.v(21,r+165,c-1),'ĐỦ')),X.eq(E.v(21,r+166,c-1),'ĐỦ')),X.eq(E.v(21,r+167,c-1),'ĐỦ')),X.eq(E.v(21,r+168,c-1),'ĐỦ')),X.eq(E.v(21,r+169,c-1),'ĐỦ')),X.eq(E.v(21,r+171,c-1),'ĐỦ')),X.eq(E.v(21,r+172,c-1),'ĐỦ')),X.eq(E.v(21,r+173,c-1),'ĐỦ')),X.eq(E.v(21,r+174,c-1),'ĐỦ')),X.eq(E.v(21,r+175,c-1),'ĐỦ')),X.eq(E.v(21,r+176,c-1),'ĐỦ')),X.eq(E.v(21,r+178,c-1),'ĐỦ')),X.eq(E.v(21,r+179,c-1),'ĐỦ')),X.eq(E.v(21,r+180,c-1),'ĐỦ')),X.eq(E.v(21,r+181,c-1),'ĐỦ')),X.eq(E.v(21,r+183,c-1),'ĐỦ')),X.eq(E.v(21,r+184,c-1),'ĐỦ')),X.eq(E.v(21,r+185,c-1),'ĐỦ')),X.eq(E.v(21,r+187,c-1),'ĐỦ')),X.eq(E.v(21,r+188,c-1),'ĐỦ')))

def f12(E, r, c):
    # Ô N7
    # = ((L112="ĐỦ")+(L113="ĐỦ")+(L114="ĐỦ")+(L115="ĐỦ")+(L116="ĐỦ")+(L117="ĐỦ")+(L118="ĐỦ")+(L119="ĐỦ")+(L120="ĐỦ")+(L121="ĐỦ")+(L122="ĐỦ")+(L123="ĐỦ")+(L124
    #   ="ĐỦ"))+((L196="ĐỦ")+(L197="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+105,c-2),'ĐỦ'),X.eq(E.v(21,r+106,c-2),'ĐỦ')),X.eq(E.v(21,r+107,c-2),'ĐỦ')),X.eq(E.v(21,r+108,c-2),'ĐỦ')),X.eq(E.v(21,r+109,c-2),'ĐỦ')),X.eq(E.v(21,r+110,c-2),'ĐỦ')),X.eq(E.v(21,r+111,c-2),'ĐỦ')),X.eq(E.v(21,r+112,c-2),'ĐỦ')),X.eq(E.v(21,r+113,c-2),'ĐỦ')),X.eq(E.v(21,r+114,c-2),'ĐỦ')),X.eq(E.v(21,r+115,c-2),'ĐỦ')),X.eq(E.v(21,r+116,c-2),'ĐỦ')),X.eq(E.v(21,r+117,c-2),'ĐỦ')),X.add(X.eq(E.v(21,r+189,c-2),'ĐỦ'),X.eq(E.v(21,r+190,c-2),'ĐỦ')))

def f13(E, r, c):
    # Ô AA7
    # = IF(IF(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$
    #   D$110,DUNG_HY_KY!$C$110)),1,IF(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN
    #   !$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_
    #   KY!$D$110,DUNG_HY_KY!$C$110))),-1,0))=1,IF(M7>=N7,"THUẬN","THUẬN NHƯNG CÓ TRỞ NGẠI"),IF(IF(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY
    #   !$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),-1,0))=-1,IF(N7>=M7,"BẤT LỢI","BẤT
    #    LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M7>N7,"NGHIÊNG THUẬN",IF(N7>M7,"NGHIÊNG BẤT LỢI","TRUNG BÌNH"))))&CHAR(10)&"("&M7&" dấu hiệu thuận, "&N7&" bất lợi; yếu t
    #   ố then chốt: "&"quan sát "&IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7
    #   +3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_H
    #   Y_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$
    #   C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)
    #   ","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(21,r,c-14),E.v(21,r,c-13))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else 0.0)),1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(21,r,c-13),E.v(21,r,c-14))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else 0.0)),X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(21,r,c-14),E.v(21,r,c-13))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(21,r,c-13),E.v(21,r,c-14))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(21,r,c-14)),' dấu hiệu thuận, '),E.v(21,r,c-13)),' bất lợi; yếu tố then chốt: '),'quan sát '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),')')

def f14(E, r, c):
    # Ô AB7
    # = TRIM(IF(L98="ĐỦ","• "&G98&" ["&B98&", "&H98&"] ","")&IF(L99="ĐỦ","• "&G99&" ["&B99&", "&H99&"] ","")&IF(L100="ĐỦ","• "&G100&" ["&B100&", "&H100&"] ","
    #   ")&IF(L101="ĐỦ","• "&G101&" ["&B101&", "&H101&"] ","")&IF(L102="ĐỦ","• "&G102&" ["&B102&", "&H102&"] ","")&IF(L103="ĐỦ","• "&G103&" ["&B103&", "&H103&
    #   "] ","")&IF(L104="ĐỦ","• "&G104&" ["&B104&", "&H104&"] ","")&IF(L105="ĐỦ","• "&G105&" ["&B105&", "&H105&"] ","")&IF(L106="ĐỦ","• "&G106&" ["&B106&", "
    #   &H106&"] ","")&IF(L107="ĐỦ","• "&G107&" ["&B107&", "&H107&"] ","")&IF(L108="ĐỦ","• "&G108&" ["&B108&", "&H108&"] ","")&IF(L109="ĐỦ","• "&G109&" ["&B10
    #   9&", "&H109&"] ","")&IF(L111="ĐỦ","• "&G111&" ["&B111&", "&H111&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+91,c-21)),' ['),E.v(21,r+91,c-26)),', '),E.v(21,r+91,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+91,c-16),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+92,c-21)),' ['),E.v(21,r+92,c-26)),', '),E.v(21,r+92,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+92,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+93,c-21)),' ['),E.v(21,r+93,c-26)),', '),E.v(21,r+93,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+93,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+94,c-21)),' ['),E.v(21,r+94,c-26)),', '),E.v(21,r+94,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+94,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+95,c-21)),' ['),E.v(21,r+95,c-26)),', '),E.v(21,r+95,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+95,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+96,c-21)),' ['),E.v(21,r+96,c-26)),', '),E.v(21,r+96,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+96,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+97,c-21)),' ['),E.v(21,r+97,c-26)),', '),E.v(21,r+97,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+97,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+98,c-21)),' ['),E.v(21,r+98,c-26)),', '),E.v(21,r+98,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+98,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+99,c-21)),' ['),E.v(21,r+99,c-26)),', '),E.v(21,r+99,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+99,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+100,c-21)),' ['),E.v(21,r+100,c-26)),', '),E.v(21,r+100,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+100,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+101,c-21)),' ['),E.v(21,r+101,c-26)),', '),E.v(21,r+101,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+101,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+102,c-21)),' ['),E.v(21,r+102,c-26)),', '),E.v(21,r+102,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+102,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+104,c-21)),' ['),E.v(21,r+104,c-26)),', '),E.v(21,r+104,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+104,c-16),'ĐỦ')) else '')))

def f15(E, r, c):
    # Ô AC7
    # = TRIM(IF(L160="ĐỦ","• "&G160&" ["&B160&", "&H160&"] ","")&IF(L161="ĐỦ","• "&G161&" ["&B161&", "&H161&"] ","")&IF(L162="ĐỦ","• "&G162&" ["&B162&", "&H16
    #   2&"] ","")&IF(L163="ĐỦ","• "&G163&" ["&B163&", "&H163&"] ","")&IF(L164="ĐỦ","• "&G164&" ["&B164&", "&H164&"] ","")&IF(L165="ĐỦ","• "&G165&" ["&B165&",
    #    "&H165&"] ","")&IF(L166="ĐỦ","• "&G166&" ["&B166&", "&H166&"] ","")&IF(L167="ĐỦ","• "&G167&" ["&B167&", "&H167&"] ","")&IF(L168="ĐỦ","• "&G168&" ["&B
    #   168&", "&H168&"] ","")&IF(L169="ĐỦ","• "&G169&" ["&B169&", "&H169&"] ","")&IF(L171="ĐỦ","• "&G171&" ["&B171&", "&H171&"] ","")&IF(L172="ĐỦ","• "&G172&
    #   " ["&B172&", "&H172&"] ","")&IF(L173="ĐỦ","• "&G173&" ["&B173&", "&H173&"] ","")&IF(L174="ĐỦ","• "&G174&" ["&B174&", "&H174&"] ","")&IF(L175="ĐỦ","• "
    #   &G175&" ["&B175&", "&H175&"] ","")&IF(L176="ĐỦ","• "&G176&" ["&B176&", "&H176&"] ","")&IF(L178="ĐỦ","• "&G178&" ["&B178&", "&H178&"] ","")&IF(L179="ĐỦ
    #   ","• "&G179&" ["&B179&", "&H179&"] ","")&IF(L180="ĐỦ","• "&G180&" ["&B180&", "&H180&"] ","")&IF(L181="ĐỦ","• "&G181&" ["&B181&", "&H181&"] ","")&IF(L1
    #   82="ĐỦ","• "&G182&" ["&B182&", "&H182&"] ","")&IF(L183="ĐỦ","• "&G183&" ["&B183&", "&H183&"] ","")&IF(L185="ĐỦ","• "&G185&" ["&B185&", "&H185&"] ","")
    #   &IF(L186="ĐỦ","• "&G186&" ["&B186&", "&H186&"] ","")&IF(L187="ĐỦ","• "&G187&" ["&B187&", "&H187&"] ","")&IF(L188="ĐỦ","• "&G188&" ["&B188&", "&H188&"]
    #    ","")&IF(L190="ĐỦ","• "&G190&" ["&B190&", "&H190&"] ","")&IF(L191="ĐỦ","• "&G191&" ["&B191&", "&H191&"] ","")&IF(L192="ĐỦ","• "&G192&" ["&B192&", "&H
    #   192&"] ","")&IF(L194="ĐỦ","• "&G194&" ["&B194&", "&H194&"] ","")&IF(L195="ĐỦ","• "&G195&" ["&B195&", "&H195&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+153,c-22)),' ['),E.v(21,r+153,c-27)),', '),E.v(21,r+153,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+153,c-17),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+154,c-22)),' ['),E.v(21,r+154,c-27)),', '),E.v(21,r+154,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+154,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+155,c-22)),' ['),E.v(21,r+155,c-27)),', '),E.v(21,r+155,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+155,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+156,c-22)),' ['),E.v(21,r+156,c-27)),', '),E.v(21,r+156,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+156,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+157,c-22)),' ['),E.v(21,r+157,c-27)),', '),E.v(21,r+157,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+157,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+158,c-22)),' ['),E.v(21,r+158,c-27)),', '),E.v(21,r+158,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+158,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+159,c-22)),' ['),E.v(21,r+159,c-27)),', '),E.v(21,r+159,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+159,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+160,c-22)),' ['),E.v(21,r+160,c-27)),', '),E.v(21,r+160,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+160,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+161,c-22)),' ['),E.v(21,r+161,c-27)),', '),E.v(21,r+161,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+161,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+162,c-22)),' ['),E.v(21,r+162,c-27)),', '),E.v(21,r+162,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+162,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+164,c-22)),' ['),E.v(21,r+164,c-27)),', '),E.v(21,r+164,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+164,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+165,c-22)),' ['),E.v(21,r+165,c-27)),', '),E.v(21,r+165,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+165,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+166,c-22)),' ['),E.v(21,r+166,c-27)),', '),E.v(21,r+166,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+166,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+167,c-22)),' ['),E.v(21,r+167,c-27)),', '),E.v(21,r+167,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+167,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+168,c-22)),' ['),E.v(21,r+168,c-27)),', '),E.v(21,r+168,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+168,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+169,c-22)),' ['),E.v(21,r+169,c-27)),', '),E.v(21,r+169,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+169,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+171,c-22)),' ['),E.v(21,r+171,c-27)),', '),E.v(21,r+171,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+171,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+172,c-22)),' ['),E.v(21,r+172,c-27)),', '),E.v(21,r+172,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+172,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+173,c-22)),' ['),E.v(21,r+173,c-27)),', '),E.v(21,r+173,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+173,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+174,c-22)),' ['),E.v(21,r+174,c-27)),', '),E.v(21,r+174,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+174,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+175,c-22)),' ['),E.v(21,r+175,c-27)),', '),E.v(21,r+175,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+175,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+176,c-22)),' ['),E.v(21,r+176,c-27)),', '),E.v(21,r+176,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+176,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+178,c-22)),' ['),E.v(21,r+178,c-27)),', '),E.v(21,r+178,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+178,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+179,c-22)),' ['),E.v(21,r+179,c-27)),', '),E.v(21,r+179,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+179,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+180,c-22)),' ['),E.v(21,r+180,c-27)),', '),E.v(21,r+180,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+180,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+181,c-22)),' ['),E.v(21,r+181,c-27)),', '),E.v(21,r+181,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+181,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+183,c-22)),' ['),E.v(21,r+183,c-27)),', '),E.v(21,r+183,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+183,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+184,c-22)),' ['),E.v(21,r+184,c-27)),', '),E.v(21,r+184,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+184,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+185,c-22)),' ['),E.v(21,r+185,c-27)),', '),E.v(21,r+185,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+185,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+187,c-22)),' ['),E.v(21,r+187,c-27)),', '),E.v(21,r+187,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+187,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+188,c-22)),' ['),E.v(21,r+188,c-27)),', '),E.v(21,r+188,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+188,c-17),'ĐỦ')) else '')))

def f16(E, r, c):
    # Ô AD7
    # = TRIM(IF(L112="ĐỦ","• "&G112&" ["&B112&", "&H112&"] ","")&IF(L113="ĐỦ","• "&G113&" ["&B113&", "&H113&"] ","")&IF(L114="ĐỦ","• "&G114&" ["&B114&", "&H11
    #   4&"] ","")&IF(L115="ĐỦ","• "&G115&" ["&B115&", "&H115&"] ","")&IF(L116="ĐỦ","• "&G116&" ["&B116&", "&H116&"] ","")&IF(L117="ĐỦ","• "&G117&" ["&B117&",
    #    "&H117&"] ","")&IF(L118="ĐỦ","• "&G118&" ["&B118&", "&H118&"] ","")&IF(L119="ĐỦ","• "&G119&" ["&B119&", "&H119&"] ","")&IF(L120="ĐỦ","• "&G120&" ["&B
    #   120&", "&H120&"] ","")&IF(L121="ĐỦ","• "&G121&" ["&B121&", "&H121&"] ","")&IF(L122="ĐỦ","• "&G122&" ["&B122&", "&H122&"] ","")&IF(L123="ĐỦ","• "&G123&
    #   " ["&B123&", "&H123&"] ","")&IF(L124="ĐỦ","• "&G124&" ["&B124&", "&H124&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+105,c-23)),' ['),E.v(21,r+105,c-28)),', '),E.v(21,r+105,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+105,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+106,c-23)),' ['),E.v(21,r+106,c-28)),', '),E.v(21,r+106,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+106,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+107,c-23)),' ['),E.v(21,r+107,c-28)),', '),E.v(21,r+107,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+107,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+108,c-23)),' ['),E.v(21,r+108,c-28)),', '),E.v(21,r+108,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+108,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+109,c-23)),' ['),E.v(21,r+109,c-28)),', '),E.v(21,r+109,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+109,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+110,c-23)),' ['),E.v(21,r+110,c-28)),', '),E.v(21,r+110,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+110,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+111,c-23)),' ['),E.v(21,r+111,c-28)),', '),E.v(21,r+111,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+111,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+112,c-23)),' ['),E.v(21,r+112,c-28)),', '),E.v(21,r+112,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+112,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+113,c-23)),' ['),E.v(21,r+113,c-28)),', '),E.v(21,r+113,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+113,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+114,c-23)),' ['),E.v(21,r+114,c-28)),', '),E.v(21,r+114,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+114,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+115,c-23)),' ['),E.v(21,r+115,c-28)),', '),E.v(21,r+115,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+115,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+116,c-23)),' ['),E.v(21,r+116,c-28)),', '),E.v(21,r+116,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+116,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+117,c-23)),' ['),E.v(21,r+117,c-28)),', '),E.v(21,r+117,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+117,c-18),'ĐỦ')) else '')))

def f17(E, r, c):
    # Ô AE7
    # = TRIM(IF(L196="ĐỦ","• "&G196&" ["&B196&", "&H196&"] ","")&IF(L197="ĐỦ","• "&G197&" ["&B197&", "&H197&"] ",""))
    return X.TRIM(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+189,c-24)),' ['),E.v(21,r+189,c-29)),', '),E.v(21,r+189,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+189,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+190,c-24)),' ['),E.v(21,r+190,c-29)),', '),E.v(21,r+190,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+190,c-19),'ĐỦ')) else '')))

def f18(E, r, c):
    # Ô J8
    # = IF(IF(OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$
    #   D$110,DUNG_HY_KY!$C$110)),1,IF(OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN
    #   !$E$7+4,5),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_
    #   KY!$D$110,DUNG_HY_KY!$C$110))),-1,IF(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF
    #   (DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,0)))=1,IF(M8>=N8,"THUẬN","THUẬN NHƯNG CÓ TRỞ NGẠI"),IF(IF(OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DO
    #   I_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(OR(MOD(T
    #   RUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+4,5),MOD(TRUNG_GIAN!$E$7+4,5)<>IF
    #   (DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),-1,IF(OR
    #   (MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DU
    #   NG_HY_KY!$C$110)),1,0)))=-1,IF(N8>=M8,"BẤT LỢI","BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M8>N8,"NGHIÊNG THUẬN",IF(N8>M8,"NGHIÊNG BẤT LỢI","TRUNG BÌNH"))))&CHA
    #   R(10)&"("&M8&" dấu hiệu thuận, "&N8&" bất lợi; yếu tố then chốt: "&"ấn "&IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY
    #   !$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+4,5)
    #   =IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+4,5),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DO
    #   I_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhi
    #   ều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&", quan sát "&IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C
    #   $33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ
    #    thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,
    #   5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$11
    #   0,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(21,r,c+3),E.v(21,r,c+4))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 0.0))),1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(21,r,c+4),E.v(21,r,c+3))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 0.0))),X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(21,r,c+3),E.v(21,r,c+4))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(21,r,c+4),E.v(21,r,c+3))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(21,r,c+3)),' dấu hiệu thuận, '),E.v(21,r,c+4)),' bất lợi; yếu tố then chốt: '),'ấn '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),', quan sát '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),')')

def f19(E, r, c):
    # Ô M8
    # = ((L125="ĐỦ")+(L126="ĐỦ")+(L127="ĐỦ")+(L128="ĐỦ")+(L129="ĐỦ")+(L130="ĐỦ")+(L131="ĐỦ")+(L132="ĐỦ")+(L133="ĐỦ")+(L134="ĐỦ")+(L135="ĐỦ")+(L136="ĐỦ")+(L137
    #   ="ĐỦ")+(L138="ĐỦ"))+((L201="ĐỦ")+(L202="ĐỦ")+(L203="ĐỦ")+(L205="ĐỦ")+(L206="ĐỦ")+(L207="ĐỦ")+(L208="ĐỦ")+(L209="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+117,c-1),'ĐỦ'),X.eq(E.v(21,r+118,c-1),'ĐỦ')),X.eq(E.v(21,r+119,c-1),'ĐỦ')),X.eq(E.v(21,r+120,c-1),'ĐỦ')),X.eq(E.v(21,r+121,c-1),'ĐỦ')),X.eq(E.v(21,r+122,c-1),'ĐỦ')),X.eq(E.v(21,r+123,c-1),'ĐỦ')),X.eq(E.v(21,r+124,c-1),'ĐỦ')),X.eq(E.v(21,r+125,c-1),'ĐỦ')),X.eq(E.v(21,r+126,c-1),'ĐỦ')),X.eq(E.v(21,r+127,c-1),'ĐỦ')),X.eq(E.v(21,r+128,c-1),'ĐỦ')),X.eq(E.v(21,r+129,c-1),'ĐỦ')),X.eq(E.v(21,r+130,c-1),'ĐỦ')),X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+193,c-1),'ĐỦ'),X.eq(E.v(21,r+194,c-1),'ĐỦ')),X.eq(E.v(21,r+195,c-1),'ĐỦ')),X.eq(E.v(21,r+197,c-1),'ĐỦ')),X.eq(E.v(21,r+198,c-1),'ĐỦ')),X.eq(E.v(21,r+199,c-1),'ĐỦ')),X.eq(E.v(21,r+200,c-1),'ĐỦ')),X.eq(E.v(21,r+201,c-1),'ĐỦ')))

def f20(E, r, c):
    # Ô N8
    # = ((L139="ĐỦ")+(L140="ĐỦ")+(L141="ĐỦ")+(L142="ĐỦ")+(L143="ĐỦ")+(L144="ĐỦ"))+((L210="ĐỦ")+(L212="ĐỦ")+(L213="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+131,c-2),'ĐỦ'),X.eq(E.v(21,r+132,c-2),'ĐỦ')),X.eq(E.v(21,r+133,c-2),'ĐỦ')),X.eq(E.v(21,r+134,c-2),'ĐỦ')),X.eq(E.v(21,r+135,c-2),'ĐỦ')),X.eq(E.v(21,r+136,c-2),'ĐỦ')),X.add(X.add(X.eq(E.v(21,r+202,c-2),'ĐỦ'),X.eq(E.v(21,r+204,c-2),'ĐỦ')),X.eq(E.v(21,r+205,c-2),'ĐỦ')))

def f21(E, r, c):
    # Ô AB8
    # = TRIM(IF(L125="ĐỦ","• "&G125&" ["&B125&", "&H125&"] ","")&IF(L126="ĐỦ","• "&G126&" ["&B126&", "&H126&"] ","")&IF(L127="ĐỦ","• "&G127&" ["&B127&", "&H12
    #   7&"] ","")&IF(L128="ĐỦ","• "&G128&" ["&B128&", "&H128&"] ","")&IF(L129="ĐỦ","• "&G129&" ["&B129&", "&H129&"] ","")&IF(L130="ĐỦ","• "&G130&" ["&B130&",
    #    "&H130&"] ","")&IF(L131="ĐỦ","• "&G131&" ["&B131&", "&H131&"] ","")&IF(L132="ĐỦ","• "&G132&" ["&B132&", "&H132&"] ","")&IF(L133="ĐỦ","• "&G133&" ["&B
    #   133&", "&H133&"] ","")&IF(L134="ĐỦ","• "&G134&" ["&B134&", "&H134&"] ","")&IF(L135="ĐỦ","• "&G135&" ["&B135&", "&H135&"] ","")&IF(L136="ĐỦ","• "&G136&
    #   " ["&B136&", "&H136&"] ","")&IF(L137="ĐỦ","• "&G137&" ["&B137&", "&H137&"] ","")&IF(L138="ĐỦ","• "&G138&" ["&B138&", "&H138&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+117,c-21)),' ['),E.v(21,r+117,c-26)),', '),E.v(21,r+117,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+117,c-16),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+118,c-21)),' ['),E.v(21,r+118,c-26)),', '),E.v(21,r+118,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+118,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+119,c-21)),' ['),E.v(21,r+119,c-26)),', '),E.v(21,r+119,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+119,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+120,c-21)),' ['),E.v(21,r+120,c-26)),', '),E.v(21,r+120,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+120,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+121,c-21)),' ['),E.v(21,r+121,c-26)),', '),E.v(21,r+121,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+121,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+122,c-21)),' ['),E.v(21,r+122,c-26)),', '),E.v(21,r+122,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+122,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+123,c-21)),' ['),E.v(21,r+123,c-26)),', '),E.v(21,r+123,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+123,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+124,c-21)),' ['),E.v(21,r+124,c-26)),', '),E.v(21,r+124,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+124,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+125,c-21)),' ['),E.v(21,r+125,c-26)),', '),E.v(21,r+125,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+125,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+126,c-21)),' ['),E.v(21,r+126,c-26)),', '),E.v(21,r+126,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+126,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+127,c-21)),' ['),E.v(21,r+127,c-26)),', '),E.v(21,r+127,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+127,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+128,c-21)),' ['),E.v(21,r+128,c-26)),', '),E.v(21,r+128,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+128,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+129,c-21)),' ['),E.v(21,r+129,c-26)),', '),E.v(21,r+129,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+129,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+130,c-21)),' ['),E.v(21,r+130,c-26)),', '),E.v(21,r+130,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+130,c-16),'ĐỦ')) else '')))

def f22(E, r, c):
    # Ô AC8
    # = TRIM(IF(L201="ĐỦ","• "&G201&" ["&B201&", "&H201&"] ","")&IF(L202="ĐỦ","• "&G202&" ["&B202&", "&H202&"] ","")&IF(L203="ĐỦ","• "&G203&" ["&B203&", "&H20
    #   3&"] ","")&IF(L205="ĐỦ","• "&G205&" ["&B205&", "&H205&"] ","")&IF(L206="ĐỦ","• "&G206&" ["&B206&", "&H206&"] ","")&IF(L207="ĐỦ","• "&G207&" ["&B207&",
    #    "&H207&"] ","")&IF(L208="ĐỦ","• "&G208&" ["&B208&", "&H208&"] ","")&IF(L209="ĐỦ","• "&G209&" ["&B209&", "&H209&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+193,c-22)),' ['),E.v(21,r+193,c-27)),', '),E.v(21,r+193,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+193,c-17),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+194,c-22)),' ['),E.v(21,r+194,c-27)),', '),E.v(21,r+194,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+194,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+195,c-22)),' ['),E.v(21,r+195,c-27)),', '),E.v(21,r+195,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+195,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+197,c-22)),' ['),E.v(21,r+197,c-27)),', '),E.v(21,r+197,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+197,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+198,c-22)),' ['),E.v(21,r+198,c-27)),', '),E.v(21,r+198,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+198,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+199,c-22)),' ['),E.v(21,r+199,c-27)),', '),E.v(21,r+199,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+199,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+200,c-22)),' ['),E.v(21,r+200,c-27)),', '),E.v(21,r+200,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+200,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+201,c-22)),' ['),E.v(21,r+201,c-27)),', '),E.v(21,r+201,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+201,c-17),'ĐỦ')) else '')))

def f23(E, r, c):
    # Ô AD8
    # = TRIM(IF(L139="ĐỦ","• "&G139&" ["&B139&", "&H139&"] ","")&IF(L140="ĐỦ","• "&G140&" ["&B140&", "&H140&"] ","")&IF(L141="ĐỦ","• "&G141&" ["&B141&", "&H14
    #   1&"] ","")&IF(L142="ĐỦ","• "&G142&" ["&B142&", "&H142&"] ","")&IF(L143="ĐỦ","• "&G143&" ["&B143&", "&H143&"] ","")&IF(L144="ĐỦ","• "&G144&" ["&B144&",
    #    "&H144&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+131,c-23)),' ['),E.v(21,r+131,c-28)),', '),E.v(21,r+131,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+131,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+132,c-23)),' ['),E.v(21,r+132,c-28)),', '),E.v(21,r+132,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+132,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+133,c-23)),' ['),E.v(21,r+133,c-28)),', '),E.v(21,r+133,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+133,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+134,c-23)),' ['),E.v(21,r+134,c-28)),', '),E.v(21,r+134,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+134,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+135,c-23)),' ['),E.v(21,r+135,c-28)),', '),E.v(21,r+135,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+135,c-18),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+136,c-23)),' ['),E.v(21,r+136,c-28)),', '),E.v(21,r+136,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+136,c-18),'ĐỦ')) else '')))

def f24(E, r, c):
    # Ô AE8
    # = TRIM(IF(L210="ĐỦ","• "&G210&" ["&B210&", "&H210&"] ","")&IF(L212="ĐỦ","• "&G212&" ["&B212&", "&H212&"] ","")&IF(L213="ĐỦ","• "&G213&" ["&B213&", "&H21
    #   3&"] ",""))
    return X.TRIM(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+202,c-24)),' ['),E.v(21,r+202,c-29)),', '),E.v(21,r+202,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+202,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+204,c-24)),' ['),E.v(21,r+204,c-29)),', '),E.v(21,r+204,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+204,c-19),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+205,c-24)),' ['),E.v(21,r+205,c-29)),', '),E.v(21,r+205,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+205,c-19),'ĐỦ')) else '')))

def f25(E, r, c):
    # Ô J9
    # = IF(IF(OR(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_K
    #   Y!$D$110,DUNG_HY_KY!$C$110)),OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHI
    #   EU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),1,IF(AND(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(S
    #   U_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>I
    #   F(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),-1,0))=1,IF(M9>=N9,"THUẬN","THUẬ
    #   N NHƯNG CÓ TRỞ NGẠI"),IF(IF(OR(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_C
    #   HIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIA
    #   N!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),1,IF(AND(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DU
    #   NG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(T
    #   RUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),-1,0))=-1,
    #   IF(N9>=M9,"BẤT LỢI","BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG"),IF(M9>N9,"NGHIÊNG THUẬN",IF(N9>M9,"NGHIÊNG BẤT LỢI","TRUNG BÌNH"))))&CHAR(10)&"("&M9&" dấu hiệu thuậ
    #   n, "&N9&" bất lợi; yếu tố then chốt: "&"quan sát "&IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần
    #   ",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,
    #   DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ t
    #   heo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&", ấn "&IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG
    #   _HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$
    #   7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+4,5),MOD(TRUNG_GIAN!$E$7+4,5)<
    #   >IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nh
    #   óm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((('THUẬN' if X.truthy(X.ge(E.v(21,r,c+3),E.v(21,r,c+4))) else 'THUẬN NHƯNG CÓ TRỞ NGẠI') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (X.neg(1.0) if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else 0.0)),1.0)) else (('BẤT LỢI' if X.truthy(X.ge(E.v(21,r,c+4),E.v(21,r,c+3))) else 'BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG') if X.truthy(X.eq((1.0 if X.truthy(X.OR(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else (X.neg(1.0) if X.truthy(X.AND(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else 0.0)),X.neg(1.0))) else ('NGHIÊNG THUẬN' if X.truthy(X.gt(E.v(21,r,c+3),E.v(21,r,c+4))) else ('NGHIÊNG BẤT LỢI' if X.truthy(X.gt(E.v(21,r,c+4),E.v(21,r,c+3))) else 'TRUNG BÌNH')))),X.CHAR(10.0)),'('),E.v(21,r,c+3)),' dấu hiệu thuận, '),E.v(21,r,c+4)),' bất lợi; yếu tố then chốt: '),'quan sát '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),', ấn '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),')')

def f26(E, r, c):
    # Ô M9
    # = ((L145="ĐỦ")+(L146="ĐỦ")+(L147="ĐỦ")+(L148="ĐỦ")+(L149="ĐỦ")+(L150="ĐỦ")+(L151="ĐỦ")+(L152="ĐỦ")+(L153="ĐỦ"))+((L214="ĐỦ")+(L215="ĐỦ")+(L216="ĐỦ")+(L2
    #   17="ĐỦ")+(L218="ĐỦ")+(L219="ĐỦ")+(L220="ĐỦ")+(L221="ĐỦ")+(L222="ĐỦ")+(L223="ĐỦ")+(L224="ĐỦ")+(L225="ĐỦ")+(L226="ĐỦ")+(L227="ĐỦ")+(L228="ĐỦ")+(L229="ĐỦ
    #   ")+(L230="ĐỦ")+(L231="ĐỦ")+(L232="ĐỦ")+(L233="ĐỦ")+(L234="ĐỦ")+(L235="ĐỦ")+(L236="ĐỦ")+(L237="ĐỦ")+(L238="ĐỦ")+(L239="ĐỦ")+(L242="ĐỦ"))
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+136,c-1),'ĐỦ'),X.eq(E.v(21,r+137,c-1),'ĐỦ')),X.eq(E.v(21,r+138,c-1),'ĐỦ')),X.eq(E.v(21,r+139,c-1),'ĐỦ')),X.eq(E.v(21,r+140,c-1),'ĐỦ')),X.eq(E.v(21,r+141,c-1),'ĐỦ')),X.eq(E.v(21,r+142,c-1),'ĐỦ')),X.eq(E.v(21,r+143,c-1),'ĐỦ')),X.eq(E.v(21,r+144,c-1),'ĐỦ')),X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.eq(E.v(21,r+205,c-1),'ĐỦ'),X.eq(E.v(21,r+206,c-1),'ĐỦ')),X.eq(E.v(21,r+207,c-1),'ĐỦ')),X.eq(E.v(21,r+208,c-1),'ĐỦ')),X.eq(E.v(21,r+209,c-1),'ĐỦ')),X.eq(E.v(21,r+210,c-1),'ĐỦ')),X.eq(E.v(21,r+211,c-1),'ĐỦ')),X.eq(E.v(21,r+212,c-1),'ĐỦ')),X.eq(E.v(21,r+213,c-1),'ĐỦ')),X.eq(E.v(21,r+214,c-1),'ĐỦ')),X.eq(E.v(21,r+215,c-1),'ĐỦ')),X.eq(E.v(21,r+216,c-1),'ĐỦ')),X.eq(E.v(21,r+217,c-1),'ĐỦ')),X.eq(E.v(21,r+218,c-1),'ĐỦ')),X.eq(E.v(21,r+219,c-1),'ĐỦ')),X.eq(E.v(21,r+220,c-1),'ĐỦ')),X.eq(E.v(21,r+221,c-1),'ĐỦ')),X.eq(E.v(21,r+222,c-1),'ĐỦ')),X.eq(E.v(21,r+223,c-1),'ĐỦ')),X.eq(E.v(21,r+224,c-1),'ĐỦ')),X.eq(E.v(21,r+225,c-1),'ĐỦ')),X.eq(E.v(21,r+226,c-1),'ĐỦ')),X.eq(E.v(21,r+227,c-1),'ĐỦ')),X.eq(E.v(21,r+228,c-1),'ĐỦ')),X.eq(E.v(21,r+229,c-1),'ĐỦ')),X.eq(E.v(21,r+230,c-1),'ĐỦ')),X.eq(E.v(21,r+233,c-1),'ĐỦ')))

def f27(E, r, c):
    # Ô N9
    # = ((L154="ĐỦ")+(L155="ĐỦ"))+((L240="ĐỦ")+(L241="ĐỦ"))
    return X.add(X.add(X.eq(E.v(21,r+145,c-2),'ĐỦ'),X.eq(E.v(21,r+146,c-2),'ĐỦ')),X.add(X.eq(E.v(21,r+231,c-2),'ĐỦ'),X.eq(E.v(21,r+232,c-2),'ĐỦ')))

def f28(E, r, c):
    # Ô AB9
    # = TRIM(IF(L145="ĐỦ","• "&G145&" ["&B145&", "&H145&"] ","")&IF(L146="ĐỦ","• "&G146&" ["&B146&", "&H146&"] ","")&IF(L147="ĐỦ","• "&G147&" ["&B147&", "&H14
    #   7&"] ","")&IF(L148="ĐỦ","• "&G148&" ["&B148&", "&H148&"] ","")&IF(L149="ĐỦ","• "&G149&" ["&B149&", "&H149&"] ","")&IF(L150="ĐỦ","• "&G150&" ["&B150&",
    #    "&H150&"] ","")&IF(L151="ĐỦ","• "&G151&" ["&B151&", "&H151&"] ","")&IF(L152="ĐỦ","• "&G152&" ["&B152&", "&H152&"] ","")&IF(L153="ĐỦ","• "&G153&" ["&B
    #   153&", "&H153&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+136,c-21)),' ['),E.v(21,r+136,c-26)),', '),E.v(21,r+136,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+136,c-16),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+137,c-21)),' ['),E.v(21,r+137,c-26)),', '),E.v(21,r+137,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+137,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+138,c-21)),' ['),E.v(21,r+138,c-26)),', '),E.v(21,r+138,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+138,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+139,c-21)),' ['),E.v(21,r+139,c-26)),', '),E.v(21,r+139,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+139,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+140,c-21)),' ['),E.v(21,r+140,c-26)),', '),E.v(21,r+140,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+140,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+141,c-21)),' ['),E.v(21,r+141,c-26)),', '),E.v(21,r+141,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+141,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+142,c-21)),' ['),E.v(21,r+142,c-26)),', '),E.v(21,r+142,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+142,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+143,c-21)),' ['),E.v(21,r+143,c-26)),', '),E.v(21,r+143,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+143,c-16),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+144,c-21)),' ['),E.v(21,r+144,c-26)),', '),E.v(21,r+144,c-20)),'] ') if X.truthy(X.eq(E.v(21,r+144,c-16),'ĐỦ')) else '')))

def f29(E, r, c):
    # Ô AC9
    # = TRIM(IF(L214="ĐỦ","• "&G214&" ["&B214&", "&H214&"] ","")&IF(L215="ĐỦ","• "&G215&" ["&B215&", "&H215&"] ","")&IF(L216="ĐỦ","• "&G216&" ["&B216&", "&H21
    #   6&"] ","")&IF(L217="ĐỦ","• "&G217&" ["&B217&", "&H217&"] ","")&IF(L218="ĐỦ","• "&G218&" ["&B218&", "&H218&"] ","")&IF(L219="ĐỦ","• "&G219&" ["&B219&",
    #    "&H219&"] ","")&IF(L220="ĐỦ","• "&G220&" ["&B220&", "&H220&"] ","")&IF(L221="ĐỦ","• "&G221&" ["&B221&", "&H221&"] ","")&IF(L222="ĐỦ","• "&G222&" ["&B
    #   222&", "&H222&"] ","")&IF(L223="ĐỦ","• "&G223&" ["&B223&", "&H223&"] ","")&IF(L224="ĐỦ","• "&G224&" ["&B224&", "&H224&"] ","")&IF(L225="ĐỦ","• "&G225&
    #   " ["&B225&", "&H225&"] ","")&IF(L226="ĐỦ","• "&G226&" ["&B226&", "&H226&"] ","")&IF(L227="ĐỦ","• "&G227&" ["&B227&", "&H227&"] ","")&IF(L228="ĐỦ","• "
    #   &G228&" ["&B228&", "&H228&"] ","")&IF(L229="ĐỦ","• "&G229&" ["&B229&", "&H229&"] ","")&IF(L230="ĐỦ","• "&G230&" ["&B230&", "&H230&"] ","")&IF(L231="ĐỦ
    #   ","• "&G231&" ["&B231&", "&H231&"] ","")&IF(L232="ĐỦ","• "&G232&" ["&B232&", "&H232&"] ","")&IF(L233="ĐỦ","• "&G233&" ["&B233&", "&H233&"] ","")&IF(L2
    #   34="ĐỦ","• "&G234&" ["&B234&", "&H234&"] ","")&IF(L235="ĐỦ","• "&G235&" ["&B235&", "&H235&"] ","")&IF(L236="ĐỦ","• "&G236&" ["&B236&", "&H236&"] ","")
    #   &IF(L237="ĐỦ","• "&G237&" ["&B237&", "&H237&"] ","")&IF(L238="ĐỦ","• "&G238&" ["&B238&", "&H238&"] ","")&IF(L239="ĐỦ","• "&G239&" ["&B239&", "&H239&"]
    #    ","")&IF(L242="ĐỦ","• "&G242&" ["&B242&", "&H242&"] ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+205,c-22)),' ['),E.v(21,r+205,c-27)),', '),E.v(21,r+205,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+205,c-17),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+206,c-22)),' ['),E.v(21,r+206,c-27)),', '),E.v(21,r+206,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+206,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+207,c-22)),' ['),E.v(21,r+207,c-27)),', '),E.v(21,r+207,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+207,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+208,c-22)),' ['),E.v(21,r+208,c-27)),', '),E.v(21,r+208,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+208,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+209,c-22)),' ['),E.v(21,r+209,c-27)),', '),E.v(21,r+209,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+209,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+210,c-22)),' ['),E.v(21,r+210,c-27)),', '),E.v(21,r+210,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+210,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+211,c-22)),' ['),E.v(21,r+211,c-27)),', '),E.v(21,r+211,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+211,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+212,c-22)),' ['),E.v(21,r+212,c-27)),', '),E.v(21,r+212,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+212,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+213,c-22)),' ['),E.v(21,r+213,c-27)),', '),E.v(21,r+213,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+213,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+214,c-22)),' ['),E.v(21,r+214,c-27)),', '),E.v(21,r+214,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+214,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+215,c-22)),' ['),E.v(21,r+215,c-27)),', '),E.v(21,r+215,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+215,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+216,c-22)),' ['),E.v(21,r+216,c-27)),', '),E.v(21,r+216,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+216,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+217,c-22)),' ['),E.v(21,r+217,c-27)),', '),E.v(21,r+217,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+217,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+218,c-22)),' ['),E.v(21,r+218,c-27)),', '),E.v(21,r+218,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+218,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+219,c-22)),' ['),E.v(21,r+219,c-27)),', '),E.v(21,r+219,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+219,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+220,c-22)),' ['),E.v(21,r+220,c-27)),', '),E.v(21,r+220,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+220,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+221,c-22)),' ['),E.v(21,r+221,c-27)),', '),E.v(21,r+221,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+221,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+222,c-22)),' ['),E.v(21,r+222,c-27)),', '),E.v(21,r+222,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+222,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+223,c-22)),' ['),E.v(21,r+223,c-27)),', '),E.v(21,r+223,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+223,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+224,c-22)),' ['),E.v(21,r+224,c-27)),', '),E.v(21,r+224,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+224,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+225,c-22)),' ['),E.v(21,r+225,c-27)),', '),E.v(21,r+225,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+225,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+226,c-22)),' ['),E.v(21,r+226,c-27)),', '),E.v(21,r+226,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+226,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+227,c-22)),' ['),E.v(21,r+227,c-27)),', '),E.v(21,r+227,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+227,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+228,c-22)),' ['),E.v(21,r+228,c-27)),', '),E.v(21,r+228,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+228,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+229,c-22)),' ['),E.v(21,r+229,c-27)),', '),E.v(21,r+229,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+229,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+230,c-22)),' ['),E.v(21,r+230,c-27)),', '),E.v(21,r+230,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+230,c-17),'ĐỦ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+233,c-22)),' ['),E.v(21,r+233,c-27)),', '),E.v(21,r+233,c-21)),'] ') if X.truthy(X.eq(E.v(21,r+233,c-17),'ĐỦ')) else '')))

def f30(E, r, c):
    # Ô AD9
    # = TRIM(IF(L154="ĐỦ","• "&G154&" ["&B154&", "&H154&"] ","")&IF(L155="ĐỦ","• "&G155&" ["&B155&", "&H155&"] ",""))
    return X.TRIM(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+145,c-23)),' ['),E.v(21,r+145,c-28)),', '),E.v(21,r+145,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+145,c-18),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+146,c-23)),' ['),E.v(21,r+146,c-28)),', '),E.v(21,r+146,c-22)),'] ') if X.truthy(X.eq(E.v(21,r+146,c-18),'ĐỦ')) else '')))

def f31(E, r, c):
    # Ô AE9
    # = TRIM(IF(L240="ĐỦ","• "&G240&" ["&B240&", "&H240&"] ","")&IF(L241="ĐỦ","• "&G241&" ["&B241&", "&H241&"] ",""))
    return X.TRIM(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+231,c-24)),' ['),E.v(21,r+231,c-29)),', '),E.v(21,r+231,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+231,c-19),'ĐỦ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(21,r+232,c-24)),' ['),E.v(21,r+232,c-29)),', '),E.v(21,r+232,c-23)),'] ') if X.truthy(X.eq(E.v(21,r+232,c-19),'ĐỦ')) else '')))

def f32(E, r, c):
    # Ô C13
    # = "Vận "&1
    return X.cat('Vận ',1.0)

def f33(E, r, c):
    # Ô D13
    # = "Vận "&2
    return X.cat('Vận ',2.0)

def f34(E, r, c):
    # Ô E13
    # = "Vận "&3
    return X.cat('Vận ',3.0)

def f35(E, r, c):
    # Ô F13
    # = "Vận "&4
    return X.cat('Vận ',4.0)

def f36(E, r, c):
    # Ô G13
    # = "Vận "&5
    return X.cat('Vận ',5.0)

def f37(E, r, c):
    # Ô H13
    # = "Vận "&6
    return X.cat('Vận ',6.0)

def f38(E, r, c):
    # Ô I13
    # = "Vận "&7
    return X.cat('Vận ',7.0)

def f39(E, r, c):
    # Ô J13
    # = "Vận "&8
    return X.cat('Vận ',8.0)

def f40(E, r, c):
    # Ô C14, D14, E14, F14, G14, H14, I14, J14
    # = DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r+21,c),' ('),E.v(14,r+22,c)),'–'),X.add(E.v(14,r+22,c),9.0)),')')

def f41(E, r, c):
    # Ô C15, D15, E15, F15, G15, H15, I15, J15
    # = DAI_VAN!C39
    return E.v(14,r+24,c)

def f42(E, r, c):
    # Ô C16, D16, E16, F16, G16, H16, I16, J16
    # = INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,DAI_VAN!C34+1)+1)
    return X.INDEX(E.rg(24,122,16,131,25), X.add(E.v(23,5,5),1.0), X.add(X.INDEX(E.rg(24,18,4,29,4), X.add(E.v(14,r+18,c),1.0)),1.0))

def f43(E, r, c):
    # Ô C17, D17, E17, F17, G17, H17, I17, J17, C18, D18, E18, F18 … (16 ô)
    # = MOD(DAI_VAN!C42-TRUNG_GIAN!$E$7,5)
    return X.MOD(X.sub(E.v(14,r+25,c),E.v(23,7,5)), 5.0)

def f44(E, r, c):
    # Ô C19, D19, E19, F19, G19, H19, I19, J19
    # = "Can "&DAI_VAN!C39&": "&IF(C17=2,IF(OR(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1)),"PHÁT TÀI – vận tài, tài "&IF(MO
    #   D(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_
    #   HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND
    #   (SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<
    #   >IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (n
    #   hàn thần)"))))&" (""vốn có tài tinh hành đến tài vận thì đại phát""; ""trong trụ tài quan đều vượng, khi hành đến tài vận thì sẽ giàu"" – P2 tr.523-52
    #   5)",IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"HAO TÀI, MANG HỌA – thân nhược gặp vận tài (""thân nhược tài nhiều thì gặp tài là ma
    #   ng họa"" – P2 tr.526)","Vận tài, tài "&IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRU
    #   NG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$
    #   D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,
    #   DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng
    #    (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&" – bình thường")),IF(C17=0,IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0
    #   ),PP1_DINH_TINH!$F$98>=P_NHIEU),"PHÁT TÀI – thân nhược tài nhiều gặp vận tỉ kiếp (""thân nhược, tài nhiều khi hành đến vận tỉ kiếp thì mới phát tài"" 
    #   – P2 tr.521)",IF(OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+0,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_
    #   HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+0,5),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUN
    #   G_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"PHÁ TÀI – tỉ kiếp tranh tài (""thân vượng hành đến vận tỷ kiếp thì phá
    #    tài bại nghiệp""; ""thiên tài sợ hành vận tỷ kiếp"" – P2 tr.527)","Vận tỉ kiếp – không phải câu tài vận nào của sách")),IF(C17=3,IF(AND((IF(DOI_CHIEU
    #   !$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),NOT(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!
    #   $E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"PHÁ TÀI – thân nhược hành vận quan sát (""quan vượng, thân nhược, hành đến qua
    #   n vận thì hại vô cùng, phá tài là chắc chắn"" – P2 tr.527)",IF(PP1_DINH_TINH!$F$98>=1,"Thuận – vốn có tài tinh hành quan vận (""vốn có tài tinh thì nê
    #   n hành đến quan vận … hành quan vận sẽ phát tài"" – P2 tr.523)","Vận quan sát – không có tài trong mệnh")),IF(C17=1,IF(AND(PP1_DINH_TINH!$F$98>=1,OR((
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRU
    #   NG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"Thuận – thực thương sinh tài (""thương quan dụng tài, hành đến tài vận
    #    thì sẽ giàu""; ""thực thần gặp tài"" – P2 tr.521-523)","Vận thực thương – chưa đủ điều kiện sinh tài"),IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,D
    #   UNG_HY_KY!$C$7)=1),PP1_DINH_TINH!$F$100>=P_NHIEU),"Hao tán – thân vượng ấn vượng (""thân vượng, ấn vượng là tiền hao tán không tụ"" – P2 tr.527)","Vận
    #    ấn – sách không nêu trực tiếp về tài")))))&CHAR(10)&"Chi "&C16&": "&IF(C18=2,IF(OR(OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109
    #   ,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_
    #   KY!$C$7)=1)),"PHÁT TÀI – vận tài, tài "&IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TR
    #   UNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!
    #   $D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109
    #   ,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằn
    #   g (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&" (""vốn có tài tinh hành đến tài vận thì đại phát""; ""trong trụ tài quan đều vượng, khi 
    #   hành đến tài vận thì sẽ giàu"" – P2 tr.523-525)",IF((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"HAO TÀI, MANG HỌA – thân nhược gặp vận 
    #   tài (""thân nhược tài nhiều thì gặp tài là mang họa"" – P2 tr.526)","Vận tài, tài "&IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109
    #   ,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIA
    #   N!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+
    #   2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"
    #   là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&" – bình thường")),IF(C18=0,IF(AND((IF(DOI_CHI
    #   EU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),PP1_DINH_TINH!$F$98>=P_NHIEU),"PHÁT TÀI – thân nhược tài nhiều gặp vận tỉ kiếp (""thân nhược, tài nhiều
    #    khi hành đến vận tỉ kiếp thì mới phát tài"" – P2 tr.521)",IF(OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+0,5)
    #   =IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+0,5),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,
    #   DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"PHÁ TÀI – tỉ kiếp tranh t
    #   ài (""thân vượng hành đến vận tỷ kiếp thì phá tài bại nghiệp""; ""thiên tài sợ hành vận tỷ kiếp"" – P2 tr.527)","Vận tỉ kiếp – không phải câu tài vận 
    #   nào của sách")),IF(C18=3,IF(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),NOT(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"PHÁ TÀI – thân nhược hành vận qua
    #   n sát (""quan vượng, thân nhược, hành đến quan vận thì hại vô cùng, phá tài là chắc chắn"" – P2 tr.527)",IF(PP1_DINH_TINH!$F$98>=1,"Thuận – vốn có tài
    #    tinh hành quan vận (""vốn có tài tinh thì nên hành đến quan vận … hành quan vận sẽ phát tài"" – P2 tr.523)","Vận quan sát – không có tài trong mệnh")
    #   ),IF(C18=1,IF(AND(PP1_DINH_TINH!$F$98>=1,OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2
    #   ,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"Thuận – thực thương sinh 
    #   tài (""thương quan dụng tài, hành đến tài vận thì sẽ giàu""; ""thực thần gặp tài"" – P2 tr.521-523)","Vận thực thương – chưa đủ điều kiện sinh tài"),I
    #   F(AND((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),PP1_DINH_TINH!$F$100>=P_NHIEU),"Hao tán – thân vượng ấn vượng (""thân vượng, ấn vượng 
    #   là tiền hao tán không tụ"" – P2 tr.527)","Vận ấn – sách không nêu trực tiếp về tài")))))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can ',E.v(14,r+20,c)),': '),((X.cat(X.cat('PHÁT TÀI – vận tài, tài ',('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),' ("vốn có tài tinh hành đến tài vận thì đại phát"; "trong trụ tài quan đều vượng, khi hành đến tài vận thì sẽ giàu" – P2 tr.523-525)') if X.truthy(X.OR(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else ('HAO TÀI, MANG HỌA – thân nhược gặp vận tài ("thân nhược tài nhiều thì gặp tài là mang họa" – P2 tr.526)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else X.cat(X.cat('Vận tài, tài ',('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),' – bình thường'))) if X.truthy(X.eq(E.v(21,r-2,c),2.0)) else (('PHÁT TÀI – thân nhược tài nhiều gặp vận tỉ kiếp ("thân nhược, tài nhiều khi hành đến vận tỉ kiếp thì mới phát tài" – P2 tr.521)' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else ('PHÁ TÀI – tỉ kiếp tranh tài ("thân vượng hành đến vận tỷ kiếp thì phá tài bại nghiệp"; "thiên tài sợ hành vận tỷ kiếp" – P2 tr.527)' if X.truthy(X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 'Vận tỉ kiếp – không phải câu tài vận nào của sách')) if X.truthy(X.eq(E.v(21,r-2,c),0.0)) else (('PHÁ TÀI – thân nhược hành vận quan sát ("quan vượng, thân nhược, hành đến quan vận thì hại vô cùng, phá tài là chắc chắn" – P2 tr.527)' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.NOT(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else ('Thuận – vốn có tài tinh hành quan vận ("vốn có tài tinh thì nên hành đến quan vận … hành quan vận sẽ phát tài" – P2 tr.523)' if X.truthy(X.ge(E.v(9,98,6),1.0)) else 'Vận quan sát – không có tài trong mệnh')) if X.truthy(X.eq(E.v(21,r-2,c),3.0)) else (('Thuận – thực thương sinh tài ("thương quan dụng tài, hành đến tài vận thì sẽ giàu"; "thực thần gặp tài" – P2 tr.521-523)' if X.truthy(X.AND(X.ge(E.v(9,98,6),1.0), X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 'Vận thực thương – chưa đủ điều kiện sinh tài') if X.truthy(X.eq(E.v(21,r-2,c),1.0)) else ('Hao tán – thân vượng ấn vượng ("thân vượng, ấn vượng là tiền hao tán không tụ" – P2 tr.527)' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.ge(E.v(9,100,6),E.v(24,99,2)))) else 'Vận ấn – sách không nêu trực tiếp về tài')))))),X.CHAR(10.0)),'Chi '),E.v(21,r-3,c)),': '),((X.cat(X.cat('PHÁT TÀI – vận tài, tài ',('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),' ("vốn có tài tinh hành đến tài vận thì đại phát"; "trong trụ tài quan đều vượng, khi hành đến tài vận thì sẽ giàu" – P2 tr.523-525)') if X.truthy(X.OR(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else ('HAO TÀI, MANG HỌA – thân nhược gặp vận tài ("thân nhược tài nhiều thì gặp tài là mang họa" – P2 tr.526)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else X.cat(X.cat('Vận tài, tài ',('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),' – bình thường'))) if X.truthy(X.eq(E.v(21,r-1,c),2.0)) else (('PHÁT TÀI – thân nhược tài nhiều gặp vận tỉ kiếp ("thân nhược, tài nhiều khi hành đến vận tỉ kiếp thì mới phát tài" – P2 tr.521)' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else ('PHÁ TÀI – tỉ kiếp tranh tài ("thân vượng hành đến vận tỷ kiếp thì phá tài bại nghiệp"; "thiên tài sợ hành vận tỷ kiếp" – P2 tr.527)' if X.truthy(X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))), X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 'Vận tỉ kiếp – không phải câu tài vận nào của sách')) if X.truthy(X.eq(E.v(21,r-1,c),0.0)) else (('PHÁ TÀI – thân nhược hành vận quan sát ("quan vượng, thân nhược, hành đến quan vận thì hại vô cùng, phá tài là chắc chắn" – P2 tr.527)' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0), X.NOT(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else ('Thuận – vốn có tài tinh hành quan vận ("vốn có tài tinh thì nên hành đến quan vận … hành quan vận sẽ phát tài" – P2 tr.523)' if X.truthy(X.ge(E.v(9,98,6),1.0)) else 'Vận quan sát – không có tài trong mệnh')) if X.truthy(X.eq(E.v(21,r-1,c),3.0)) else (('Thuận – thực thương sinh tài ("thương quan dụng tài, hành đến tài vận thì sẽ giàu"; "thực thần gặp tài" – P2 tr.521-523)' if X.truthy(X.AND(X.ge(E.v(9,98,6),1.0), X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else 'Vận thực thương – chưa đủ điều kiện sinh tài') if X.truthy(X.eq(E.v(21,r-1,c),1.0)) else ('Hao tán – thân vượng ấn vượng ("thân vượng, ấn vượng là tiền hao tán không tụ" – P2 tr.527)' if X.truthy(X.AND(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.ge(E.v(9,100,6),E.v(24,99,2)))) else 'Vận ấn – sách không nêu trực tiếp về tài'))))))

def f45(E, r, c):
    # Ô C20, D20, E20, F20, G20, H20, I20, J20
    # = "Can "&DAI_VAN!C39&": "&IF(C17=3,IF(OR(OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1)),"THĂNG TIẾN – vận quan sát, quan
    #    sát "&IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$
    #   C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ 
    #   thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIA
    #   N!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/h
    #   ỷ/kỵ thần (nhàn thần)"))))&" (""quan, ấn tương sinh lại gặp tuế vận là quan trong triều"" – P2 tr.537)","BẤT LỢI – thân nhược hành vận quan sát (""qua
    #   n vượng, thân nhược, hành đến quan vận thì hại vô cùng"" – P2 tr.527; ""tứ trụ có quan tinh khi hành đến quan vận thì sẽ bị phá hại"" – P2 tr.542)"),I
    #   F(AND(C17=1,C15="Thương quan",(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0)),"MẤT CHỨC / BỊ BIẾM – vận thương quan khi mệnh có chính quan
    #    (""trụ có quan có chức nhưng hành vận thương quan thì bị bãi miễn hoặc bỏ chức về"" – P2 tr.541)",IF(AND(C17=1,(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SA
    #   O!$Y$5:$Y$20,0)>0)),"Thuận – thực thương chế sát (""thực thần chế ngự sát là anh hùng"" – P2 tr.535)",IF(AND(C17=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D
    #   $7,DUNG_HY_KY!$C$7)=1),(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0)),"MẤT CHỨC – thân vượng có sát gặp vận ấn (""thân vượng có sát, lại hà
    #   nh vận ấn thụ là mất chức"" – P2 tr.534)",IF(AND(C17=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),"Thuận – vận ấn sinh thân (""ấn thụ 
    #   được sinh vượng thì lợi cho quan vận"" – P2 tr.535)",IF(AND(C17=2,PP1_DINH_TINH!$F$99>=1,OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),
    #   OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,
    #   DUNG_HY_KY!$C$110)))),"Thuận – tài sinh quan (""vốn có quan tinh thì nên hành đến tài vận, hành tài vận sẽ sinh quan"" – P2 tr.523)",IF(AND(C17=0,(IF(
    #   DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),PP1_DINH_TINH!$F$99>=1),"Tranh chấp – thân vượng gặp tỉ kiếp (""quan chính ngôi, ngang vai mạnh 
    #   thì sự tranh chấp rất dữ dội"" – P2 tr.542)","Không có câu quan vận nào của sách ứng với vận này")))))))&CHAR(10)&"Chi "&C16&": "&IF(C18=3,IF(OR(OR(MO
    #   D(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_
    #   HY_KY!$C$110)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1)),"THĂNG TIẾN – vận quan sát, quan sát "&IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CH
    #   IEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),
    #   "là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$
    #   E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY
    #   !$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&" (""quan, ấn tươ
    #   ng sinh lại gặp tuế vận là quan trong triều"" – P2 tr.537)","BẤT LỢI – thân nhược hành vận quan sát (""quan vượng, thân nhược, hành đến quan vận thì h
    #   ại vô cùng"" – P2 tr.527; ""tứ trụ có quan tinh khi hành đến quan vận thì sẽ bị phá hại"" – P2 tr.542)"),IF(AND(C18=1,C16="Thương quan",(COUNTIFS(SAO!
    #   $F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0)),"MẤT CHỨC / BỊ BIẾM – vận thương quan khi mệnh có chính quan (""trụ có quan có chức nhưng hành vận thươn
    #   g quan thì bị bãi miễn hoặc bỏ chức về"" – P2 tr.541)",IF(AND(C18=1,(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0)),"Thuận – thực thương chế
    #    sát (""thực thần chế ngự sát là anh hùng"" – P2 tr.535)",IF(AND(C18=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),(COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0)>0)),"MẤT CHỨC – thân vượng có sát gặp vận ấn (""thân vượng có sát, lại hành vận ấn thụ là mất chức"" – P2 tr.534)",IF
    #   (AND(C18=4,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)),"Thuận – vận ấn sinh thân (""ấn thụ được sinh vượng thì lợi cho quan vận"" – P2 
    #   tr.535)",IF(AND(C18=2,PP1_DINH_TINH!$F$99>=1,OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$
    #   33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)))),"Thuận – tài sinh quan
    #    (""vốn có quan tinh thì nên hành đến tài vận, hành tài vận sẽ sinh quan"" – P2 tr.523)",IF(AND(C18=0,(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY
    #   !$C$7)=1),PP1_DINH_TINH!$F$99>=1),"Tranh chấp – thân vượng gặp tỉ kiếp (""quan chính ngôi, ngang vai mạnh thì sự tranh chấp rất dữ dội"" – P2 tr.542)"
    #   ,"Không có câu quan vận nào của sách ứng với vận này")))))))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can ',E.v(14,r+19,c)),': '),((X.cat(X.cat('THĂNG TIẾN – vận quan sát, quan sát ',('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),' ("quan, ấn tương sinh lại gặp tuế vận là quan trong triều" – P2 tr.537)') if X.truthy(X.OR(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else 'BẤT LỢI – thân nhược hành vận quan sát ("quan vượng, thân nhược, hành đến quan vận thì hại vô cùng" – P2 tr.527; "tứ trụ có quan tinh khi hành đến quan vận thì sẽ bị phá hại" – P2 tr.542)') if X.truthy(X.eq(E.v(21,r-3,c),3.0)) else ('MẤT CHỨC / BỊ BIẾM – vận thương quan khi mệnh có chính quan ("trụ có quan có chức nhưng hành vận thương quan thì bị bãi miễn hoặc bỏ chức về" – P2 tr.541)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),1.0), X.eq(E.v(21,r-5,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0))) else ('Thuận – thực thương chế sát ("thực thần chế ngự sát là anh hùng" – P2 tr.535)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0))) else ('MẤT CHỨC – thân vượng có sát gặp vận ấn ("thân vượng có sát, lại hành vận ấn thụ là mất chức" – P2 tr.534)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0))) else ('Thuận – vận ấn sinh thân ("ấn thụ được sinh vượng thì lợi cho quan vận" – P2 tr.535)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ('Thuận – tài sinh quan ("vốn có quan tinh thì nên hành đến tài vận, hành tài vận sẽ sinh quan" – P2 tr.523)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),2.0), X.ge(E.v(9,99,6),1.0), X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else ('Tranh chấp – thân vượng gặp tỉ kiếp ("quan chính ngôi, ngang vai mạnh thì sự tranh chấp rất dữ dội" – P2 tr.542)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),0.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.ge(E.v(9,99,6),1.0))) else 'Không có câu quan vận nào của sách ứng với vận này')))))))),X.CHAR(10.0)),'Chi '),E.v(21,r-4,c)),': '),((X.cat(X.cat('THĂNG TIẾN – vận quan sát, quan sát ',('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),' ("quan, ấn tương sinh lại gặp tuế vận là quan trong triều" – P2 tr.537)') if X.truthy(X.OR(X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0))) else 'BẤT LỢI – thân nhược hành vận quan sát ("quan vượng, thân nhược, hành đến quan vận thì hại vô cùng" – P2 tr.527; "tứ trụ có quan tinh khi hành đến quan vận thì sẽ bị phá hại" – P2 tr.542)') if X.truthy(X.eq(E.v(21,r-2,c),3.0)) else ('MẤT CHỨC / BỊ BIẾM – vận thương quan khi mệnh có chính quan ("trụ có quan có chức nhưng hành vận thương quan thì bị bãi miễn hoặc bỏ chức về" – P2 tr.541)' if X.truthy(X.AND(X.eq(E.v(21,r-2,c),1.0), X.eq(E.v(21,r-4,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0))) else ('Thuận – thực thương chế sát ("thực thần chế ngự sát là anh hùng" – P2 tr.535)' if X.truthy(X.AND(X.eq(E.v(21,r-2,c),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0))) else ('MẤT CHỨC – thân vượng có sát gặp vận ấn ("thân vượng có sát, lại hành vận ấn thụ là mất chức" – P2 tr.534)' if X.truthy(X.AND(X.eq(E.v(21,r-2,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0))) else ('Thuận – vận ấn sinh thân ("ấn thụ được sinh vượng thì lợi cho quan vận" – P2 tr.535)' if X.truthy(X.AND(X.eq(E.v(21,r-2,c),4.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ('Thuận – tài sinh quan ("vốn có quan tinh thì nên hành đến tài vận, hành tài vận sẽ sinh quan" – P2 tr.523)' if X.truthy(X.AND(X.eq(E.v(21,r-2,c),2.0), X.ge(E.v(9,99,6),1.0), X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))))) else ('Tranh chấp – thân vượng gặp tỉ kiếp ("quan chính ngôi, ngang vai mạnh thì sự tranh chấp rất dữ dội" – P2 tr.542)' if X.truthy(X.AND(X.eq(E.v(21,r-2,c),0.0), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.ge(E.v(9,99,6),1.0))) else 'Không có câu quan vận nào của sách ứng với vận này'))))))))

def f46(E, r, c):
    # Ô C21, D21, E21, F21, G21, H21, I21, J21
    # = "Can "&DAI_VAN!C39&": "&IF(AND(OR(C17=2,C17=3),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),PP1_DINH_TINH!$F$98<=1,PP1_DINH_TINH!$F$99<=1
    #   ),"NỔI DANH – trụ ngày vượng, tài quan yếu gặp vận tài quan (""… đến lúc hành vận tài quan thì sẽ nổi danh"" – P2 tr.552)",IF(C17=3,IF(OR((IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$
    #   7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Danh có khí – vận quan sát (""hành vận quan sát thì danh có khí"" – P2 tr.552)","D
    #   anh thái quá thì xấu – thân nhược gặp vận quan sát (""… nhưng thái quá (vượng quá) thì lại xấu"" – P2 tr.552)"),IF(AND(C17=4,OR(MOD(TRUNG_GIAN!$E$7+4,
    #   5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Lợi
    #    học tập, thi cử – vận ấn là dụng/hỷ (""ấn vượng, sát yếu thì thi đỗ"" – P2 tr.546)",IF(AND(C17=1,C15="Thương quan",(COUNTIFS(SAO!$F$5:$F$20,"Chính qu
    #   an",SAO!$Y$5:$Y$20,0)>0)),"Kị thi cử – vận thương quan gặp quan (""đi thi kị nhất thương quan gặp quan"" – P2 tr.548)",IF(AND(C17=2,(COUNTIFS(SAO!$F$5
    #   :$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0)),"Thành danh – thương quan có tài lại hành tài vận (""thương quan còn có tài, lại hành tài vận thì công thàn
    #   h danh toại"" – P2 tr.553)","Không có câu công danh nào của sách ứng với vận này")))))&CHAR(10)&"Chi "&C16&": "&IF(AND(OR(C18=2,C18=3),(IF(DOI_CHIEU!$
    #   C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),PP1_DINH_TINH!$F$98<=1,PP1_DINH_TINH!$F$99<=1),"NỔI DANH – trụ ngày vượng, tài quan yếu gặp vận tài quan ("
    #   "… đến lúc hành vận tài quan thì sẽ nổi danh"" – P2 tr.552)",IF(C18=3,IF(OR((IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=1),OR(MOD(TRUNG_GIA
    #   N!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$11
    #   0))),"Danh có khí – vận quan sát (""hành vận quan sát thì danh có khí"" – P2 tr.552)","Danh thái quá thì xấu – thân nhược gặp vận quan sát (""… nhưng 
    #   thái quá (vượng quá) thì lại xấu"" – P2 tr.552)"),IF(AND(C18=4,OR(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),M
    #   OD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))),"Lợi học tập, thi cử – vận ấn là dụng/hỷ (""ấn vượng, sát yếu thì 
    #   thi đỗ"" – P2 tr.546)",IF(AND(C18=1,C16="Thương quan",(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0)),"Kị thi cử – vận thương quan gặp qua
    #   n (""đi thi kị nhất thương quan gặp quan"" – P2 tr.548)",IF(AND(C18=2,(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0)),"Thành danh – thươn
    #   g quan có tài lại hành tài vận (""thương quan còn có tài, lại hành tài vận thì công thành danh toại"" – P2 tr.553)","Không có câu công danh nào của sá
    #   ch ứng với vận này")))))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Can ',E.v(14,r+18,c)),': '),('NỔI DANH – trụ ngày vượng, tài quan yếu gặp vận tài quan ("… đến lúc hành vận tài quan thì sẽ nổi danh" – P2 tr.552)' if X.truthy(X.AND(X.OR(X.eq(E.v(21,r-4,c),2.0), X.eq(E.v(21,r-4,c),3.0)), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.le(E.v(9,98,6),1.0), X.le(E.v(9,99,6),1.0))) else (('Danh có khí – vận quan sát ("hành vận quan sát thì danh có khí" – P2 tr.552)' if X.truthy(X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else 'Danh thái quá thì xấu – thân nhược gặp vận quan sát ("… nhưng thái quá (vượng quá) thì lại xấu" – P2 tr.552)') if X.truthy(X.eq(E.v(21,r-4,c),3.0)) else ('Lợi học tập, thi cử – vận ấn là dụng/hỷ ("ấn vượng, sát yếu thì thi đỗ" – P2 tr.546)' if X.truthy(X.AND(X.eq(E.v(21,r-4,c),4.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else ('Kị thi cử – vận thương quan gặp quan ("đi thi kị nhất thương quan gặp quan" – P2 tr.548)' if X.truthy(X.AND(X.eq(E.v(21,r-4,c),1.0), X.eq(E.v(21,r-6,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0))) else ('Thành danh – thương quan có tài lại hành tài vận ("thương quan còn có tài, lại hành tài vận thì công thành danh toại" – P2 tr.553)' if X.truthy(X.AND(X.eq(E.v(21,r-4,c),2.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0))) else 'Không có câu công danh nào của sách ứng với vận này')))))),X.CHAR(10.0)),'Chi '),E.v(21,r-5,c)),': '),('NỔI DANH – trụ ngày vượng, tài quan yếu gặp vận tài quan ("… đến lúc hành vận tài quan thì sẽ nổi danh" – P2 tr.552)' if X.truthy(X.AND(X.OR(X.eq(E.v(21,r-3,c),2.0), X.eq(E.v(21,r-3,c),3.0)), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.le(E.v(9,98,6),1.0), X.le(E.v(9,99,6),1.0))) else (('Danh có khí – vận quan sát ("hành vận quan sát thì danh có khí" – P2 tr.552)' if X.truthy(X.OR(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),1.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else 'Danh thái quá thì xấu – thân nhược gặp vận quan sát ("… nhưng thái quá (vượng quá) thì lại xấu" – P2 tr.552)') if X.truthy(X.eq(E.v(21,r-3,c),3.0)) else ('Lợi học tập, thi cử – vận ấn là dụng/hỷ ("ấn vượng, sát yếu thì thi đỗ" – P2 tr.546)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),4.0), X.OR(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))))) else ('Kị thi cử – vận thương quan gặp quan ("đi thi kị nhất thương quan gặp quan" – P2 tr.548)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),1.0), X.eq(E.v(21,r-5,c),'Thương quan'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0))) else ('Thành danh – thương quan có tài lại hành tài vận ("thương quan còn có tài, lại hành tài vận thì công thành danh toại" – P2 tr.553)' if X.truthy(X.AND(X.eq(E.v(21,r-3,c),2.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0))) else 'Không có câu công danh nào của sách ứng với vận này'))))))

def f47(E, r, c):
    # Ô C22
    # = "Thuận: "&IF(TRIM(IF(M22=1,C14&" ","")&IF(N22=1,D14&" ","")&IF(O22=1,E14&" ","")&IF(P22=1,F14&" ","")&IF(Q22=1,G14&" ","")&IF(R22=1,H14&" ","")&IF(S22
    #   =1,I14&" ","")&IF(T22=1,J14&" ",""))="","không có",TRIM(IF(M22=1,C14&" ","")&IF(N22=1,D14&" ","")&IF(O22=1,E14&" ","")&IF(P22=1,F14&" ","")&IF(Q22=1,G
    #   14&" ","")&IF(R22=1,H14&" ","")&IF(S22=1,I14&" ","")&IF(T22=1,J14&" ","")))&CHAR(10)&"Bất lợi / cẩn trọng: "&IF(TRIM(IF(M22=-1,C14&" ","")&IF(N22=-1,D
    #   14&" ","")&IF(O22=-1,E14&" ","")&IF(P22=-1,F14&" ","")&IF(Q22=-1,G14&" ","")&IF(R22=-1,H14&" ","")&IF(S22=-1,I14&" ","")&IF(T22=-1,J14&" ",""))="","kh
    #   ông có",TRIM(IF(M22=-1,C14&" ","")&IF(N22=-1,D14&" ","")&IF(O22=-1,E14&" ","")&IF(P22=-1,F14&" ","")&IF(Q22=-1,G14&" ","")&IF(R22=-1,H14&" ","")&IF(S2
    #   2=-1,I14&" ","")&IF(T22=-1,J14&" ","")))&CHAR(10)&"Lẫn lộn (can và chi của vận cho kết quả ngược nhau – vận tác dụng cả 10 năm, tốt xấu đan xen; xem M
    #   AU_THUAN M14): "&IF(TRIM(IF(M22=2,C14&" ","")&IF(N22=2,D14&" ","")&IF(O22=2,E14&" ","")&IF(P22=2,F14&" ","")&IF(Q22=2,G14&" ","")&IF(R22=2,H14&" ","")
    #   &IF(S22=2,I14&" ","")&IF(T22=2,J14&" ",""))="","không có",TRIM(IF(M22=2,C14&" ","")&IF(N22=2,D14&" ","")&IF(O22=2,E14&" ","")&IF(P22=2,F14&" ","")&IF(
    #   Q22=2,G14&" ","")&IF(R22=2,H14&" ","")&IF(S22=2,I14&" ","")&IF(T22=2,J14&" ","")))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thuận: ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-8,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),1.0)) else ''),(X.cat(E.v(21,r-8,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),1.0)) else '')),(X.cat(E.v(21,r-8,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),1.0)) else '')),(X.cat(E.v(21,r-8,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),1.0)) else '')),(X.cat(E.v(21,r-8,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),1.0)) else '')),(X.cat(E.v(21,r-8,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),1.0)) else '')),(X.cat(E.v(21,r-8,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),1.0)) else '')),(X.cat(E.v(21,r-8,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),1.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-8,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),1.0)) else ''),(X.cat(E.v(21,r-8,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),1.0)) else '')),(X.cat(E.v(21,r-8,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),1.0)) else '')),(X.cat(E.v(21,r-8,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),1.0)) else '')),(X.cat(E.v(21,r-8,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),1.0)) else '')),(X.cat(E.v(21,r-8,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),1.0)) else '')),(X.cat(E.v(21,r-8,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),1.0)) else '')),(X.cat(E.v(21,r-8,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),1.0)) else ''))))),X.CHAR(10.0)),'Bất lợi / cẩn trọng: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-8,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),X.neg(1.0))) else ''),(X.cat(E.v(21,r-8,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),X.neg(1.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-8,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),X.neg(1.0))) else ''),(X.cat(E.v(21,r-8,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),X.neg(1.0))) else '')),(X.cat(E.v(21,r-8,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),X.neg(1.0))) else ''))))),X.CHAR(10.0)),'Lẫn lộn (can và chi của vận cho kết quả ngược nhau – vận tác dụng cả 10 năm, tốt xấu đan xen; xem MAU_THUAN M14): '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-8,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),2.0)) else ''),(X.cat(E.v(21,r-8,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),2.0)) else '')),(X.cat(E.v(21,r-8,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),2.0)) else '')),(X.cat(E.v(21,r-8,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),2.0)) else '')),(X.cat(E.v(21,r-8,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),2.0)) else '')),(X.cat(E.v(21,r-8,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),2.0)) else '')),(X.cat(E.v(21,r-8,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),2.0)) else '')),(X.cat(E.v(21,r-8,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),2.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-8,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),2.0)) else ''),(X.cat(E.v(21,r-8,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),2.0)) else '')),(X.cat(E.v(21,r-8,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),2.0)) else '')),(X.cat(E.v(21,r-8,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),2.0)) else '')),(X.cat(E.v(21,r-8,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),2.0)) else '')),(X.cat(E.v(21,r-8,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),2.0)) else '')),(X.cat(E.v(21,r-8,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),2.0)) else '')),(X.cat(E.v(21,r-8,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),2.0)) else '')))))

def f48(E, r, c):
    # Ô M22, N22, O22, P22, Q22, R22, S22, T22
    # = IF(AND(OR(ISNUMBER(SEARCH("PHÁT TÀI",C19)),ISNUMBER(SEARCH("Thuận –",C19))),OR(ISNUMBER(SEARCH("PHÁ TÀI",C19)),ISNUMBER(SEARCH("HAO TÀI",C19)),ISNUMBE
    #   R(SEARCH("Hao tán",C19)))),2,IF(OR(ISNUMBER(SEARCH("PHÁT TÀI",C19)),ISNUMBER(SEARCH("Thuận –",C19))),1,IF(OR(ISNUMBER(SEARCH("PHÁ TÀI",C19)),ISNUMBER(
    #   SEARCH("HAO TÀI",C19)),ISNUMBER(SEARCH("Hao tán",C19))),-1,0)))
    return (2.0 if X.truthy(X.AND(X.OR(X.ISNUMBER(lambda: X.SEARCH('PHÁT TÀI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Thuận –', E.v(21,r-3,c-10)))), X.OR(X.ISNUMBER(lambda: X.SEARCH('PHÁ TÀI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('HAO TÀI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Hao tán', E.v(21,r-3,c-10)))))) else (1.0 if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('PHÁT TÀI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Thuận –', E.v(21,r-3,c-10))))) else (X.neg(1.0) if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('PHÁ TÀI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('HAO TÀI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Hao tán', E.v(21,r-3,c-10))))) else 0.0)))

def f49(E, r, c):
    # Ô C23
    # = "Thuận: "&IF(TRIM(IF(M23=1,C14&" ","")&IF(N23=1,D14&" ","")&IF(O23=1,E14&" ","")&IF(P23=1,F14&" ","")&IF(Q23=1,G14&" ","")&IF(R23=1,H14&" ","")&IF(S23
    #   =1,I14&" ","")&IF(T23=1,J14&" ",""))="","không có",TRIM(IF(M23=1,C14&" ","")&IF(N23=1,D14&" ","")&IF(O23=1,E14&" ","")&IF(P23=1,F14&" ","")&IF(Q23=1,G
    #   14&" ","")&IF(R23=1,H14&" ","")&IF(S23=1,I14&" ","")&IF(T23=1,J14&" ","")))&CHAR(10)&"Bất lợi / cẩn trọng: "&IF(TRIM(IF(M23=-1,C14&" ","")&IF(N23=-1,D
    #   14&" ","")&IF(O23=-1,E14&" ","")&IF(P23=-1,F14&" ","")&IF(Q23=-1,G14&" ","")&IF(R23=-1,H14&" ","")&IF(S23=-1,I14&" ","")&IF(T23=-1,J14&" ",""))="","kh
    #   ông có",TRIM(IF(M23=-1,C14&" ","")&IF(N23=-1,D14&" ","")&IF(O23=-1,E14&" ","")&IF(P23=-1,F14&" ","")&IF(Q23=-1,G14&" ","")&IF(R23=-1,H14&" ","")&IF(S2
    #   3=-1,I14&" ","")&IF(T23=-1,J14&" ","")))&CHAR(10)&"Lẫn lộn (can và chi của vận cho kết quả ngược nhau – vận tác dụng cả 10 năm, tốt xấu đan xen; xem M
    #   AU_THUAN M14): "&IF(TRIM(IF(M23=2,C14&" ","")&IF(N23=2,D14&" ","")&IF(O23=2,E14&" ","")&IF(P23=2,F14&" ","")&IF(Q23=2,G14&" ","")&IF(R23=2,H14&" ","")
    #   &IF(S23=2,I14&" ","")&IF(T23=2,J14&" ",""))="","không có",TRIM(IF(M23=2,C14&" ","")&IF(N23=2,D14&" ","")&IF(O23=2,E14&" ","")&IF(P23=2,F14&" ","")&IF(
    #   Q23=2,G14&" ","")&IF(R23=2,H14&" ","")&IF(S23=2,I14&" ","")&IF(T23=2,J14&" ","")))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thuận: ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-9,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),1.0)) else ''),(X.cat(E.v(21,r-9,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),1.0)) else '')),(X.cat(E.v(21,r-9,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),1.0)) else '')),(X.cat(E.v(21,r-9,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),1.0)) else '')),(X.cat(E.v(21,r-9,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),1.0)) else '')),(X.cat(E.v(21,r-9,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),1.0)) else '')),(X.cat(E.v(21,r-9,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),1.0)) else '')),(X.cat(E.v(21,r-9,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),1.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-9,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),1.0)) else ''),(X.cat(E.v(21,r-9,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),1.0)) else '')),(X.cat(E.v(21,r-9,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),1.0)) else '')),(X.cat(E.v(21,r-9,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),1.0)) else '')),(X.cat(E.v(21,r-9,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),1.0)) else '')),(X.cat(E.v(21,r-9,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),1.0)) else '')),(X.cat(E.v(21,r-9,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),1.0)) else '')),(X.cat(E.v(21,r-9,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),1.0)) else ''))))),X.CHAR(10.0)),'Bất lợi / cẩn trọng: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-9,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),X.neg(1.0))) else ''),(X.cat(E.v(21,r-9,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),X.neg(1.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-9,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),X.neg(1.0))) else ''),(X.cat(E.v(21,r-9,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),X.neg(1.0))) else '')),(X.cat(E.v(21,r-9,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),X.neg(1.0))) else ''))))),X.CHAR(10.0)),'Lẫn lộn (can và chi của vận cho kết quả ngược nhau – vận tác dụng cả 10 năm, tốt xấu đan xen; xem MAU_THUAN M14): '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-9,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),2.0)) else ''),(X.cat(E.v(21,r-9,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),2.0)) else '')),(X.cat(E.v(21,r-9,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),2.0)) else '')),(X.cat(E.v(21,r-9,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),2.0)) else '')),(X.cat(E.v(21,r-9,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),2.0)) else '')),(X.cat(E.v(21,r-9,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),2.0)) else '')),(X.cat(E.v(21,r-9,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),2.0)) else '')),(X.cat(E.v(21,r-9,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),2.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-9,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),2.0)) else ''),(X.cat(E.v(21,r-9,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),2.0)) else '')),(X.cat(E.v(21,r-9,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),2.0)) else '')),(X.cat(E.v(21,r-9,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),2.0)) else '')),(X.cat(E.v(21,r-9,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),2.0)) else '')),(X.cat(E.v(21,r-9,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),2.0)) else '')),(X.cat(E.v(21,r-9,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),2.0)) else '')),(X.cat(E.v(21,r-9,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),2.0)) else '')))))

def f50(E, r, c):
    # Ô M23, N23, O23, P23, Q23, R23, S23, T23
    # = IF(AND(OR(ISNUMBER(SEARCH("THĂNG TIẾN",C20)),ISNUMBER(SEARCH("Thuận –",C20))),OR(ISNUMBER(SEARCH("MẤT CHỨC",C20)),ISNUMBER(SEARCH("BẤT LỢI",C20)),ISNU
    #   MBER(SEARCH("Tranh chấp",C20)))),2,IF(OR(ISNUMBER(SEARCH("THĂNG TIẾN",C20)),ISNUMBER(SEARCH("Thuận –",C20))),1,IF(OR(ISNUMBER(SEARCH("MẤT CHỨC",C20)),
    #   ISNUMBER(SEARCH("BẤT LỢI",C20)),ISNUMBER(SEARCH("Tranh chấp",C20))),-1,0)))
    return (2.0 if X.truthy(X.AND(X.OR(X.ISNUMBER(lambda: X.SEARCH('THĂNG TIẾN', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Thuận –', E.v(21,r-3,c-10)))), X.OR(X.ISNUMBER(lambda: X.SEARCH('MẤT CHỨC', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Tranh chấp', E.v(21,r-3,c-10)))))) else (1.0 if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('THĂNG TIẾN', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Thuận –', E.v(21,r-3,c-10))))) else (X.neg(1.0) if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('MẤT CHỨC', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('BẤT LỢI', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Tranh chấp', E.v(21,r-3,c-10))))) else 0.0)))

def f51(E, r, c):
    # Ô C24
    # = "Thuận: "&IF(TRIM(IF(M24=1,C14&" ","")&IF(N24=1,D14&" ","")&IF(O24=1,E14&" ","")&IF(P24=1,F14&" ","")&IF(Q24=1,G14&" ","")&IF(R24=1,H14&" ","")&IF(S24
    #   =1,I14&" ","")&IF(T24=1,J14&" ",""))="","không có",TRIM(IF(M24=1,C14&" ","")&IF(N24=1,D14&" ","")&IF(O24=1,E14&" ","")&IF(P24=1,F14&" ","")&IF(Q24=1,G
    #   14&" ","")&IF(R24=1,H14&" ","")&IF(S24=1,I14&" ","")&IF(T24=1,J14&" ","")))&CHAR(10)&"Bất lợi / cẩn trọng: "&IF(TRIM(IF(M24=-1,C14&" ","")&IF(N24=-1,D
    #   14&" ","")&IF(O24=-1,E14&" ","")&IF(P24=-1,F14&" ","")&IF(Q24=-1,G14&" ","")&IF(R24=-1,H14&" ","")&IF(S24=-1,I14&" ","")&IF(T24=-1,J14&" ",""))="","kh
    #   ông có",TRIM(IF(M24=-1,C14&" ","")&IF(N24=-1,D14&" ","")&IF(O24=-1,E14&" ","")&IF(P24=-1,F14&" ","")&IF(Q24=-1,G14&" ","")&IF(R24=-1,H14&" ","")&IF(S2
    #   4=-1,I14&" ","")&IF(T24=-1,J14&" ","")))&CHAR(10)&"Lẫn lộn (can và chi của vận cho kết quả ngược nhau – vận tác dụng cả 10 năm, tốt xấu đan xen; xem M
    #   AU_THUAN M14): "&IF(TRIM(IF(M24=2,C14&" ","")&IF(N24=2,D14&" ","")&IF(O24=2,E14&" ","")&IF(P24=2,F14&" ","")&IF(Q24=2,G14&" ","")&IF(R24=2,H14&" ","")
    #   &IF(S24=2,I14&" ","")&IF(T24=2,J14&" ",""))="","không có",TRIM(IF(M24=2,C14&" ","")&IF(N24=2,D14&" ","")&IF(O24=2,E14&" ","")&IF(P24=2,F14&" ","")&IF(
    #   Q24=2,G14&" ","")&IF(R24=2,H14&" ","")&IF(S24=2,I14&" ","")&IF(T24=2,J14&" ","")))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thuận: ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-10,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),1.0)) else ''),(X.cat(E.v(21,r-10,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),1.0)) else '')),(X.cat(E.v(21,r-10,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),1.0)) else '')),(X.cat(E.v(21,r-10,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),1.0)) else '')),(X.cat(E.v(21,r-10,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),1.0)) else '')),(X.cat(E.v(21,r-10,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),1.0)) else '')),(X.cat(E.v(21,r-10,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),1.0)) else '')),(X.cat(E.v(21,r-10,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),1.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-10,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),1.0)) else ''),(X.cat(E.v(21,r-10,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),1.0)) else '')),(X.cat(E.v(21,r-10,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),1.0)) else '')),(X.cat(E.v(21,r-10,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),1.0)) else '')),(X.cat(E.v(21,r-10,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),1.0)) else '')),(X.cat(E.v(21,r-10,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),1.0)) else '')),(X.cat(E.v(21,r-10,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),1.0)) else '')),(X.cat(E.v(21,r-10,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),1.0)) else ''))))),X.CHAR(10.0)),'Bất lợi / cẩn trọng: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-10,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),X.neg(1.0))) else ''),(X.cat(E.v(21,r-10,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),X.neg(1.0))) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-10,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),X.neg(1.0))) else ''),(X.cat(E.v(21,r-10,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),X.neg(1.0))) else '')),(X.cat(E.v(21,r-10,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),X.neg(1.0))) else ''))))),X.CHAR(10.0)),'Lẫn lộn (can và chi của vận cho kết quả ngược nhau – vận tác dụng cả 10 năm, tốt xấu đan xen; xem MAU_THUAN M14): '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-10,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),2.0)) else ''),(X.cat(E.v(21,r-10,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),2.0)) else '')),(X.cat(E.v(21,r-10,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),2.0)) else '')),(X.cat(E.v(21,r-10,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),2.0)) else '')),(X.cat(E.v(21,r-10,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),2.0)) else '')),(X.cat(E.v(21,r-10,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),2.0)) else '')),(X.cat(E.v(21,r-10,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),2.0)) else '')),(X.cat(E.v(21,r-10,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),2.0)) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(21,r-10,c),' ') if X.truthy(X.eq(E.v(21,r,c+10),2.0)) else ''),(X.cat(E.v(21,r-10,c+1),' ') if X.truthy(X.eq(E.v(21,r,c+11),2.0)) else '')),(X.cat(E.v(21,r-10,c+2),' ') if X.truthy(X.eq(E.v(21,r,c+12),2.0)) else '')),(X.cat(E.v(21,r-10,c+3),' ') if X.truthy(X.eq(E.v(21,r,c+13),2.0)) else '')),(X.cat(E.v(21,r-10,c+4),' ') if X.truthy(X.eq(E.v(21,r,c+14),2.0)) else '')),(X.cat(E.v(21,r-10,c+5),' ') if X.truthy(X.eq(E.v(21,r,c+15),2.0)) else '')),(X.cat(E.v(21,r-10,c+6),' ') if X.truthy(X.eq(E.v(21,r,c+16),2.0)) else '')),(X.cat(E.v(21,r-10,c+7),' ') if X.truthy(X.eq(E.v(21,r,c+17),2.0)) else '')))))

def f52(E, r, c):
    # Ô M24, N24, O24, P24, Q24, R24, S24, T24
    # = IF(AND(OR(ISNUMBER(SEARCH("NỔI DANH",C21)),ISNUMBER(SEARCH("Danh có khí",C21)),ISNUMBER(SEARCH("Thành danh",C21)),ISNUMBER(SEARCH("Lợi học tập",C21)))
    #   ,OR(ISNUMBER(SEARCH("Kị thi cử",C21)),ISNUMBER(SEARCH("thái quá",C21)))),2,IF(OR(ISNUMBER(SEARCH("NỔI DANH",C21)),ISNUMBER(SEARCH("Danh có khí",C21)),
    #   ISNUMBER(SEARCH("Thành danh",C21)),ISNUMBER(SEARCH("Lợi học tập",C21))),1,IF(OR(ISNUMBER(SEARCH("Kị thi cử",C21)),ISNUMBER(SEARCH("thái quá",C21))),-1
    #   ,0)))
    return (2.0 if X.truthy(X.AND(X.OR(X.ISNUMBER(lambda: X.SEARCH('NỔI DANH', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Danh có khí', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Thành danh', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Lợi học tập', E.v(21,r-3,c-10)))), X.OR(X.ISNUMBER(lambda: X.SEARCH('Kị thi cử', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('thái quá', E.v(21,r-3,c-10)))))) else (1.0 if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('NỔI DANH', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Danh có khí', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Thành danh', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('Lợi học tập', E.v(21,r-3,c-10))))) else (X.neg(1.0) if X.truthy(X.OR(X.ISNUMBER(lambda: X.SEARCH('Kị thi cử', E.v(21,r-3,c-10))), X.ISNUMBER(lambda: X.SEARCH('thái quá', E.v(21,r-3,c-10))))) else 0.0)))

def f53(E, r, c):
    # Ô C27
    # = "• TÀI VẬN: "&LEFT(J6,FIND(CHAR(10),J6)-1)&" – tài "&IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng th
    #   ần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=
    #   2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG
    #   _HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ
    #    theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&"; đại vận – "&SUBSTITUTE(SUBSTITUTE(C22,"Thuận: ","thuận: "),CHAR(10),"; ")&
    #   "."&CHAR(10)&"• QUAN VẬN: "&LEFT(J7,FIND(CHAR(10),J7)-1)&" – quan sát "&IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!
    #   $C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+3,5)=
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+3,5),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI
    #   _CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiề
    #   u nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&"; đại vận – "&SUBSTITUTE(SUBSTITUTE(C23,"Thuận: ","thuận:
    #    "),CHAR(10),"; ")&"."&CHAR(10)&"• HỌC TẬP – THI CỬ: "&LEFT(J8,FIND(CHAR(10),J8)-1)&" – ấn "&IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_
    #   KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(
    #   TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+4,5),MOD(TRUNG_G
    #   IAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$
    #   C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&". DANH TIẾNG: "&LEFT(J9,FIND(CHAR(10
    #   ),J9)-1)&"; đại vận – "&SUBSTITUTE(SUBSTITUTE(C24,"Thuận: ","thuận: "),CHAR(10),"; ")&"."&CHAR(10)&"→ Đọc nhanh: THUẬN = sao chủ sự là dụng/hỷ thần và
    #    các câu của sách đủ điều kiện nghiêng về tốt; BẤT LỢI = sao chủ sự là kỵ thần; các vận nêu tên là lúc sách nói việc đó ứng nghiệm rõ nhất. Lưu ý: đây
    #    là xếp loại THEO TỪNG VIỆC (tiền bạc, c"&"hức vụ, danh tiếng) theo câu của sách về hành vận; xu thế CHUNG của vận theo dụng thần ở DAI_VAN có thể khá
    #   c ở cùng một vận vì hai bên trả lời hai câu hỏi khác nhau (ví dụ vận dụng thần là ấn tốt cho thân nhưng sách không nói vận ấn làm già"&"u)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• TÀI VẬN: ',X.LEFT(E.v(21,r-21,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(21,r-21,c+7)),1.0))),' – tài '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),'; đại vận – '),X.SUBSTITUTE(X.SUBSTITUTE(E.v(21,r-5,c), 'Thuận: ', 'thuận: '), X.CHAR(10.0), '; ')),'.'),X.CHAR(10.0)),'• QUAN VẬN: '),X.LEFT(E.v(21,r-20,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(21,r-20,c+7)),1.0))),' – quan sát '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),'; đại vận – '),X.SUBSTITUTE(X.SUBSTITUTE(E.v(21,r-4,c), 'Thuận: ', 'thuận: '), X.CHAR(10.0), '; ')),'.'),X.CHAR(10.0)),'• HỌC TẬP – THI CỬ: '),X.LEFT(E.v(21,r-19,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(21,r-19,c+7)),1.0))),' – ấn '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),'. DANH TIẾNG: '),X.LEFT(E.v(21,r-18,c+7), X.sub(X.FIND(X.CHAR(10.0), E.v(21,r-18,c+7)),1.0))),'; đại vận – '),X.SUBSTITUTE(X.SUBSTITUTE(E.v(21,r-3,c), 'Thuận: ', 'thuận: '), X.CHAR(10.0), '; ')),'.'),X.CHAR(10.0)),'→ Đọc nhanh: THUẬN = sao chủ sự là dụng/hỷ thần và các câu của sách đủ điều kiện nghiêng về tốt; BẤT LỢI = sao chủ sự là kỵ thần; các vận nêu tên là lúc sách nói việc đó ứng nghiệm rõ nhất. Lưu ý: đây là xếp loại THEO TỪNG VIỆC (tiền bạc, c'),'hức vụ, danh tiếng) theo câu của sách về hành vận; xu thế CHUNG của vận theo dụng thần ở DAI_VAN có thể khác ở cùng một vận vì hai bên trả lời hai câu hỏi khác nhau (ví dụ vận dụng thần là ấn tốt cho thân nhưng sách không nói vận ấn làm già'),'u).')

def f54(E, r, c):
    # Ô C28
    # = IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)>=0,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)-TRUNG_GIAN!$E$7,5)=3
    #   ,INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$5+1)=0),IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1))>0,"Dụng thần là quan sát và tứ 
    #   trụ có kình dương hỗ trợ: có uy, vinh hiển; ấn tinh có lực là cứu ứng của quan sát (P1 tr.117).","Dụng thần là quan sát mà tứ trụ không có kình dương:
    #    khó vinh hiển, uy quyền hạn chế; ấn tinh có lực là cứu ứng (P1 tr.117)."),"(không xét: dụng thần không phải quan sát, hoặc nhật can âm)")
    return (('Dụng thần là quan sát và tứ trụ có kình dương hỗ trợ: có uy, vinh hiển; ấn tinh có lực là cứu ứng của quan sát (P1 tr.117).' if X.truthy(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.INDEX(E.rg(24,5,8,14,8), X.add(E.v(23,5,5),1.0))),0.0)) else 'Dụng thần là quan sát mà tứ trụ không có kình dương: khó vinh hiển, uy quyền hạn chế; ấn tinh có lực là cứu ứng (P1 tr.117).') if X.truthy(X.AND(X.ge((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0), X.eq(X.MOD(X.sub((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),E.v(23,7,5)), 5.0),3.0), X.eq(X.INDEX(E.rg(24,5,4,14,4), X.add(E.v(23,5,5),1.0)),0.0))) else '(không xét: dụng thần không phải quan sát, hoặc nhật can âm)')

def f55(E, r, c):
    # Ô E62, F62
    # = IF((PP1_DINH_TINH!$F$98>=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,98,6),1.0)) else '—')

def f56(E, r, c):
    # Ô J62, J63, J64, J65, J66, J67, J68, J69, J70, J71, J72, J73 … (233 ô)
    # = IF(AND(E62="ĐỦ",F62="ĐỦ"),"cả hai",IF(E62="ĐỦ","PP1",IF(F62="ĐỦ","PP2","")))
    return ('cả hai' if X.truthy(X.AND(X.eq(E.v(21,r,c-5),'ĐỦ'), X.eq(E.v(21,r,c-4),'ĐỦ'))) else ('PP1' if X.truthy(X.eq(E.v(21,r,c-5),'ĐỦ')) else ('PP2' if X.truthy(X.eq(E.v(21,r,c-4),'ĐỦ')) else '')))

def f57(E, r, c):
    # Ô K62
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P2 tr.520): Có tài (rõ hoặc mờ) → "&IF(J62=
    #   "","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J62&") ⇒ "&G62&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P2 tr.520): Có tài (rõ hoặc mờ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f58(E, r, c):
    # Ô L62, L63, L64, L65, L66, L67, L68, L69, L70, L71, L72, L73 … (233 ô)
    # = IF(DOI_CHIEU!$C$33=2,F62,E62)
    return (E.v(21,r,c-6) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(21,r,c-7))

def f59(E, r, c):
    # Ô E63, F63
    # = IF(AND((COUNTIF(TRUNG_GIAN!$C$6:$F$6,CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4))>0),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)),0.0), X.ge(E.v(9,98,6),1.0))) else '—')

def f60(E, r, c):
    # Ô K63
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&"
    #    ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HAN
    #   H_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.520, 524): Có tài và có chi kho (mộ) của hành tài → "&IF(J63="",
    #   "chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J63&") ⇒ "&G63&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.520, 524): Có tài và có chi kho (mộ) của hành tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f61(E, r, c):
    # Ô E64
    # = IF(AND((DUNG_HY_KY!$C$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f62(E, r, c):
    # Ô F64
    # = IF(AND((DUNG_HY_KY!$D$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f63(E, r, c):
    # Ô K64
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.520): Thân vượng và tài có lực (tài lộ/bản khí vượng tại lệnh hoặc tài ≥ P_NHIEU chữ) → "&IF(J64="","c
    #   hưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J64&") ⇒ "&G64&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.520): Thân vượng và tài có lực (tài lộ/bản khí vượng tại lệnh hoặc tài ≥ P_NHIEU chữ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f64(E, r, c):
    # Ô E65, E103
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.ge(E.v(9,99,6),1.0))) else '—')

def f65(E, r, c):
    # Ô F65, F103
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2)))), X.ge(E.v(9,99,6),1.0))) else '—')

def f66(E, r, c):
    # Ô K65
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&
    #   "; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can
    #   ",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can
    #   ",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Th
    #   iên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ
    #   ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng
    #    tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.520): Thân
    #    vượng, tài có lực và có quan sát → "&IF(J65="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J65&") ⇒ "&G65&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.520): Thân vượng, tài có lực và có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f67(E, r, c):
    # Ô E66
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.ge(E.v(9,100,6),1.0))) else '—')

def f68(E, r, c):
    # Ô F66
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.ge(E.v(9,100,6),1.0))) else '—')

def f69(E, r, c):
    # Ô K66
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" →
    #    Quy tắc (P2 tr.520-521): Thân nhược, tài ≥ P_NHIEU chữ và có ấn → "&IF(J66="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J66&")
    #    ⇒ "&G66&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.520-521): Thân nhược, tài ≥ P_NHIEU chữ và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f70(E, r, c):
    # Ô E67, F67
    # = IF(AND((NHAP_LIEU!$C$22="Nam"),AND(COUNTIF(TRUNG_GIAN!$C$6:$F$6,4)>0,COUNTIF(TRUNG_GIAN!$C$6:$F$6,10)>0,COUNTIF(TRUNG_GIAN!$C$6:$F$6,1)>0,COUNTIF(TRUN
    #   G_GIAN!$C$6:$F$6,7)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(2,22,3),'Nam'), X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 4.0),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 10.0),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 1.0),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 7.0),0.0)))) else '—')

def f71(E, r, c):
    # Ô K67
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&" → Quy tắc (P2 tr.520): Nam, tứ trụ có đủ Thìn, Tuất, Sửu, Mùi → "&IF(J67="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J67&") ⇒ "&G67&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),' → Quy tắc (P2 tr.520): Nam, tứ trụ có đủ Thìn, Tuất, Sửu, Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f72(E, r, c):
    # Ô E68, F68
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5
    #   :$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0), X.NOT(X.gt(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),0.0)))) else '—')

def f73(E, r, c):
    # Ô K68
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,3)>0,"Thiên tài ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Thiên tài ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thiên tài ở t
    #   háng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thiên tài ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$
    #   5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P2 tr.520): Thiên tài ở trụ giờ và các trụ khác không có thiên tài → "&IF(J68="","chưa đủ điều kiện ở cả hai phư
    #   ơng pháp – không áp dụng.","ĐỦ ("&J68&") ⇒ "&G68&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thiên tài ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else ''),(X.cat(X.cat('Thiên tài ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Thiên tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.520): Thiên tài ở trụ giờ và các trụ khác không có thiên tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f74(E, r, c):
    # Ô E69, F69
    # = IF(OR(AND(TRUNG_GIAN!$E$5=9,TRUNG_GIAN!$E$6=5)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,5,5),9.0), X.eq(E.v(23,6,5),5.0)))) else '—')

def f75(E, r, c):
    # Ô K69
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.520): Trụ ngày là Quý Tị → "&IF(J69="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J69&") ⇒ "&G69&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.520): Trụ ngày là Quý Tị → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f76(E, r, c):
    # Ô E70, F70
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$K$5:$K$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),0.0)) else '—')

def f77(E, r, c):
    # Ô K70
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$
    #   20,0)>0,"Chính tài lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)&" chữ; ","")&IF(COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)>0,"Thiên tài lâm Trường sinh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.520, 522): Có chữ tài đóng ở đất Trường sinh (12 cung của can tài t
    #   ại chi trụ đó) → "&IF(J70="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J70&") ⇒ "&G70&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài lâm Trường sinh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.520, 522): Có chữ tài đóng ở đất Trường sinh (12 cung của can tài tại chi trụ đó) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f78(E, r, c):
    # Ô E71, F71
    # = IF(OR(AND((TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),(SAO!$M$6=1)),AND((TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,1
    #   0,1,4)),(SAO!$M$10=1)),AND((TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),(SAO!$M$14=1)),AND((TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!
    #   $E$7+2,5)+1,7,10,10,1,4)),(SAO!$M$18=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,6,13),1.0)), X.AND(X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,10,13),1.0)), X.AND(X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,14,13),1.0)), X.AND(X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,18,13),1.0)))) else '—')

def f79(E, r, c):
    # Ô K71
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Bị xung ở năm (chi "&TRUNG_GIAN!$C$11&")
    #   : "&IF(SAO!$M$6=1,"có","không")&"; "&"Bị xung ở tháng (chi "&TRUNG_GIAN!$D$11&"): "&IF(SAO!$M$10=1,"có","không")&"; "&"Bị xung ở ngày (chi "&TRUNG_GIA
    #   N!$E$11&"): "&IF(SAO!$M$14=1,"có","không")&"; "&"Bị xung ở giờ (chi "&TRUNG_GIAN!$F$11&"): "&IF(SAO!$M$18=1,"có","không")&" → Quy tắc (P2 tr.520-521):
    #    Chi kho của tài có trong tứ trụ và bị xung (kề) → "&IF(J71="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J71&") ⇒ "&G71&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Bị xung ở năm (chi '),E.v(23,11,3)),'): '),('có' if X.truthy(X.eq(E.v(13,6,13),1.0)) else 'không')),'; '),'Bị xung ở tháng (chi '),E.v(23,11,4)),'): '),('có' if X.truthy(X.eq(E.v(13,10,13),1.0)) else 'không')),'; '),'Bị xung ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,13),1.0)) else 'không')),'; '),'Bị xung ở giờ (chi '),E.v(23,11,6)),'): '),('có' if X.truthy(X.eq(E.v(13,18,13),1.0)) else 'không')),' → Quy tắc (P2 tr.520-521): Chi kho của tài có trong tứ trụ và bị xung (kề) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f80(E, r, c):
    # Ô E72
    # = IF(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0))) else '—')

def f81(E, r, c):
    # Ô F72
    # = IF(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0))) else '—')

def f82(E, r, c):
    # Ô K72
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&" → Quy tắc (P2 tr.520): Tài là dụng thần → "&IF(J72="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J72&") ⇒ "&G72&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.520): Tài là dụng thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f83(E, r, c):
    # Ô E73
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.le(E.v(9,98,6),1.0)), X.ge(E.v(9,97,6),1.0))) else '—')

def f84(E, r, c):
    # Ô F73
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.le(E.v(9,98,6),1.0)), X.ge(E.v(9,97,6),1.0))) else '—')

def f85(E, r, c):
    # Ô K73
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P2 tr.520, 522): Thân 
    #   vượng, tài ≤ 1 chữ, có thực thương → "&IF(J73="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J73&") ⇒ "&G73&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P2 tr.520, 522): Thân vượng, tài ≤ 1 chữ, có thực thương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f86(E, r, c):
    # Ô E74, F74, E111, F111, E145, F145
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,98,6),1.0))) else '—')

def f87(E, r, c):
    # Ô K74
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521-522): Có thương quan và có tài → "&IF(J74="","chưa đủ điều kiện ở cả hai phương pháp – không á
    #   p dụng.","ĐỦ ("&J74&") ⇒ "&G74&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.521-522): Có thương quan và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f88(E, r, c):
    # Ô E75, F75
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),0.0)) else '—')

def f89(E, r, c):
    # Ô K75
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"<>Can",SAO!$B$5:$B$20,2)>0,"Chính tài tàng can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>C
    #   an",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)>0
    #   ,"Thiên tài tàng can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)&" chữ; "
    #   ,""))&" → Quy tắc (P2 tr.521-522): Có tài ở chi ngày (tàng trong chi ngày) → "&IF(J75="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ
    #    ("&J75&") ⇒ "&G75&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.521-522): Có tài ở chi ngày (tàng trong chi ngày) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f90(E, r, c):
    # Ô E76, F76, E125, F125, E150, F150
    # = IF(AND(AND((PP1_DINH_TINH!$F$98>=1),(PP1_DINH_TINH!$F$99>=1)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,98,6),1.0), X.ge(E.v(9,99,6),1.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f91(E, r, c):
    # Ô K76
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.521): 
    #   Có đủ tài, quan sát, ấn → "&IF(J76="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J76&") ⇒ "&G76&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.521): Có đủ tài, quan sát, ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f92(E, r, c):
    # Ô E77, F77
    # = IF(AND((PP1_DINH_TINH!$F$98>=P_NHIEU),(PP1_DINH_TINH!$F$96>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),E.v(24,99,2)), X.ge(E.v(9,96,6),1.0))) else '—')

def f93(E, r, c):
    # Ô K77
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.522): Tài ≥ P_NHIEU chữ và có tỉ kiếp → "&IF(J77="","chưa 
    #   đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J77&") ⇒ "&G77&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.522): Tài ≥ P_NHIEU chữ và có tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f94(E, r, c):
    # Ô E78
    # = IF(AND(AND(DUNG_HY_KY!$C$109>=0,MOD(DUNG_HY_KY!$C$109-TRUNG_GIAN!$E$7,5)=2),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>
    #   0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,109,3),0.0), X.eq(X.MOD(X.sub(E.v(8,109,3),E.v(23,7,5)), 5.0),2.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f95(E, r, c):
    # Ô F78
    # = IF(AND(AND(DUNG_HY_KY!$D$109>=0,MOD(DUNG_HY_KY!$D$109-TRUNG_GIAN!$E$7,5)=2),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>
    #   0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(8,109,4),0.0), X.eq(X.MOD(X.sub(E.v(8,109,4),E.v(23,7,5)), 5.0),2.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f96(E, r, c):
    # Ô K78
    # = "Dữ kiện: "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$D$109+1),"—")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên tài: "
    #   &COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.522): Thiên tài là dụng thần (dụng thần là tà
    #   i và có thiên tài) → "&IF(J78="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J78&") ⇒ "&G78&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.522): Thiên tài là dụng thần (dụng thần là tài và có thiên tài) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f97(E, r, c):
    # Ô E79
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f98(E, r, c):
    # Ô F79
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f99(E, r, c):
    # Ô K79
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)
    #   >0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.522): Thân vượng, có thiên ấn v
    #   à có tài → "&IF(J79="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J79&") ⇒ "&G79&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.522): Thân vượng, có thiên ấn và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f100(E, r, c):
    # Ô E80, F80
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f101(E, r, c):
    # Ô K80
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.522): Có thực thần và có ấn → "&IF(J80="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J80&") ⇒ "&G80&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.522): Có thực thần và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f102(E, r, c):
    # Ô E81, F81
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),0.0)) else '—')

def f103(E, r, c):
    # Ô K81
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$B$5:$B$20,1)>0,"Chính ấn thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",S
    #   AO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Thiên 
    #   ấn thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → 
    #   Quy tắc (P2 tr.522): Can tháng là ấn → "&IF(J81="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J81&") ⇒ "&G81&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.522): Can tháng là ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f104(E, r, c):
    # Ô E82, F82
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),((COUNT
    #   IFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),0.0))) else '—')

def f105(E, r, c):
    # Ô K82
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Can",SAO!$B$5:$B$20,0)>0,"Tỉ kiên thiên can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B
    #   $5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Tỉ kiên thiê
    #   n can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)>0,"Kiếp tài thiên can ở năm: "&COUNTIFS(SAO!$F$5:$F$
    #   20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Kiếp tài thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$B$5:$B$20,3)>0,"Chính tài thiên can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",
    #   SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)>0,"Thiê
    #   n tài thiên can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)&" chữ; ",""))&" 
    #   → Quy tắc (P2 tr.523): Tỉ kiếp lộ ở can năm/tháng, tài lộ ở can giờ → "&IF(J82="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J82
    #   &") ⇒ "&G82&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Tỉ kiên thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Tỉ kiên thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Kiếp tài thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.523): Tỉ kiếp lộ ở can năm/tháng, tài lộ ở can giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f106(E, r, c):
    # Ô E83, F83
    # = IF((TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))) else '—')

def f107(E, r, c):
    # Ô K83
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.5
    #   23): Chi giờ là kho của tài → "&IF(J83="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J83&") ⇒ "&G83&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.523): Chi giờ là kho của tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f108(E, r, c):
    # Ô E84, E86
    # = IF(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f109(E, r, c):
    # Ô F84, F86
    # = IF(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f110(E, r, c):
    # Ô K84
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.521): Thân nhược, tài ≥ P_NHIEU chữ → "&IF(J84="","chưa đủ điều kiện ở cả
    #    hai phương pháp – không áp dụng.","ĐỦ ("&J84&") ⇒ "&G84&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.521): Thân nhược, tài ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f111(E, r, c):
    # Ô E85, F85
    # = IF((PP1_DINH_TINH!$D$98>=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,98,4),1.0)) else '—')

def f112(E, r, c):
    # Ô K85
    # = "Dữ kiện: "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&" → Quy tắc (P2 tr.520, 527): Tài lộ ở thiên can → "&IF(J85="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J85&") ⇒ "&G85&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài lộ rõ: '),E.v(9,98,4)),' → Quy tắc (P2 tr.520, 527): Tài lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f113(E, r, c):
    # Ô K86
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.526-527): Thân nhược, tài ≥ P_NHIEU chữ → "&IF(J86="","chưa đủ điều kiện 
    #   ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J86&") ⇒ "&G86&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.526-527): Thân nhược, tài ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f114(E, r, c):
    # Ô E87
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f115(E, r, c):
    # Ô F87
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f116(E, r, c):
    # Ô K87
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$
    #   J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521, 527): Thân nhược và tài có lực → "&IF(J87="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J87&") ⇒ "&G87&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.521, 527): Thân nhược và tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f117(E, r, c):
    # Ô E88, F88
    # = IF((PP1_DINH_TINH!$F$98>=3),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,98,6),3.0)) else '—')

def f118(E, r, c):
    # Ô K88
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P2 tr.521): Tài ≥ 3 chữ → "&IF(J88="","chưa
    #    đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J88&") ⇒ "&G88&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P2 tr.521): Tài ≥ 3 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f119(E, r, c):
    # Ô E89, F89
    # = IF(AND((PP1_DINH_TINH!$F$98>=P_NHIEU),(PP1_DINH_TINH!$D$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),E.v(24,99,2)), X.ge(E.v(9,98,4),1.0))) else '—')

def f120(E, r, c):
    # Ô K89
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"ngưỡ
    #   ng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.522): Tài ≥ P_NHIEU chữ và tài lộ ở thiên can → "&IF(J89="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J89&") ⇒ "&G89&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.522): Tài ≥ P_NHIEU chữ và tài lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f121(E, r, c):
    # Ô E90, F90
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '—')

def f122(E, r, c):
    # Ô K90
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$B$5:$B$20,3)>0,"Thiên tài thiên can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",
    #   SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Thiên tài lộ ở can giờ → "&IF(J90="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.
    #   ","ĐỦ ("&J90&") ⇒ "&G90&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên tài thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else ''))),' → Quy tắc (P2 tr.527): Thiên tài lộ ở can giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f123(E, r, c):
    # Ô E91
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$100>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,100,6),E.v(24,99,2)))) else '—')

def f124(E, r, c):
    # Ô F91
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$100>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,100,6),E.v(24,99,2)))) else '—')

def f125(E, r, c):
    # Ô K91
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," (
    #   "&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.527): Thân vượng, ấn ≥ P_NHIEU chữ → "&IF(J91="","chưa đủ điều kiện ở c
    #   ả hai phương pháp – không áp dụng.","ĐỦ ("&J91&") ⇒ "&G91&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.527): Thân vượng, ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f126(E, r, c):
    # Ô E92, F92
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f127(E, r, c):
    # Ô K92
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thầ
    #   n: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Có thực thần v
    #   à có thiên ấn → "&IF(J92="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J92&") ⇒ "&G92&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.527): Có thực thần và có thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f128(E, r, c):
    # Ô E93
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f129(E, r, c):
    # Ô F93
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,98,6),1.0))) else '—')

def f130(E, r, c):
    # Ô K93
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)
    #   >0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Thân nhược, có thiên ấn v
    #   à có tài → "&IF(J93="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J93&") ⇒ "&G93&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.527): Thân nhược, có thiên ấn và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f131(E, r, c):
    # Ô E94
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$96>=P_NHIEU)),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,96,6),E.v(24,99,2))), X.eq(E.v(9,98,6),0.0))) else '—')

def f132(E, r, c):
    # Ô F94
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$96>=P_NHIEU)),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,96,6),E.v(24,99,2))), X.eq(E.v(9,98,6),0.0))) else '—')

def f133(E, r, c):
    # Ô K94
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="","",
    #   " ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"
    #    → Quy tắc (P2 tr.527): Thân vượng, tỉ kiếp ≥ P_NHIEU, không có tài → "&IF(J94="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J94
    #   &") ⇒ "&G94&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.527): Thân vượng, tỉ kiếp ≥ P_NHIEU, không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f134(E, r, c):
    # Ô E95, F95
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,0)+COUNTIFS(SAO!$F$5
    #   :$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,0))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),0.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 0.0)),0.0)))) else '—')

def f135(E, r, c):
    # Ô K95
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$
    #   20,1)>0,"Chính tài tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)>0,"Thiên tài tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,0)>0,"Chí
    #   nh tài không tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$
    #   F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,0)>0,"Thiên tài không tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Mọi chữ tài đều rơi vào tuần không → "&IF(J95="","chưa đủ điều kiện ở c
    #   ả hai phương pháp – không áp dụng.","ĐỦ ("&J95&") ⇒ "&G95&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài không tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 0.0),0.0)) else '')),(X.cat(X.cat('Thiên tài không tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.527): Mọi chữ tài đều rơi vào tuần không → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f136(E, r, c):
    # Ô E96, F96
    # = IF(AND((PP1_DINH_TINH!$F$98>=1),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),1.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f137(E, r, c):
    # Ô K96
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Có tài và có kình dương (theo nhật 
    #   can) ở chi → "&IF(J96="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J96&") ⇒ "&G96&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.527): Có tài và có kình dương (theo nhật can) ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f138(E, r, c):
    # Ô E97, F97
    # = IF((PP1_DINH_TINH!$F$98=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(9,98,6),0.0)) else '—')

def f139(E, r, c):
    # Ô K97
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&" → Quy tắc (P2 tr.527): Không có tài → "&IF(J97="","chư
    #   a đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J97&") ⇒ "&G97&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P2 tr.527): Không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f140(E, r, c):
    # Ô E98
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$99>=P_NHIEU)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,99,6),E.v(24,99,2))), X.ge(E.v(9,97,6),1.0))) else '—')

def f141(E, r, c):
    # Ô F98
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$99>=P_NHIEU)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,99,6),E.v(24,99,2))), X.ge(E.v(9,97,6),1.0))) else '—')

def f142(E, r, c):
    # Ô K98
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU
    #   &" chữ"&" → Quy tắc (P2 tr.532): Thân vượng, quan sát ≥ P_NHIEU chữ và có thực thương chế → "&IF(J98="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J98&") ⇒ "&G98&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.532): Thân vượng, quan sát ≥ P_NHIEU chữ và có thực thương chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f143(E, r, c):
    # Ô E99
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f144(E, r, c):
    # Ô F99
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,100,6),1.0))) else '—')

def f145(E, r, c):
    # Ô K99
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," (
    #   "&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.532): Thân vượng, có thất sát
    #    và có ấn → "&IF(J99="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J99&") ⇒ "&G99&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.532): Thân vượng, có thất sát và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f146(E, r, c):
    # Ô E100, F100
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,97,6),1.0))) else '—')

def f147(E, r, c):
    # Ô K100
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt
    #   ): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535, 541): Có thất sát và có thực thương chế sát → "&IF(J100="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J100&") ⇒ "&G100&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.535, 541): Có thất sát và có thực thương chế sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f148(E, r, c):
    # Ô E101
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0), X.eq(E.v(8,7,3),1.0))) else '—')

def f149(E, r, c):
    # Ô F101
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0), X.eq(E.v(8,7,4),1.0))) else '—')

def f150(E, r, c):
    # Ô K101
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thất sát ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.533): Thất sát ở trụ tháng và thân vượng → "&IF(J101="","chưa đủ điều kiện ở
    #    cả hai phương pháp – không áp dụng.","ĐỦ ("&J101&") ⇒ "&G101&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.533): Thất sát ở trụ tháng và thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f151(E, r, c):
    # Ô E102, F102
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L
    #   $20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),0.0)) else '—')

def f152(E, r, c):
    # Ô K102
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L
    #   $20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)>0,"Chính quan không bị khắc/hợp mất không bị xung ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)>0,"Thất sát không bị khắc/hợp mất không bị xung ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất s
    #   át",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)>0,"Chính ấn không bị khắc/hợp mất không bị xung ở ngày: "&COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thi
    #   ên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)>0,"Thiên ấn không bị khắc/hợp mất không bị xung ở ngày: "
    #   &COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0,SAO!$B$5:$B$20,2)&" chữ; ",""))&" → Quy tắc (P
    #   2 tr.533): Trụ ngày có quan sát hoặc ấn, chữ đó không bị khắc/hợp/xung → "&IF(J102="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("
    #   &J102&") ⇒ "&G102&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan không bị khắc/hợp mất không bị xung ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Thất sát không bị khắc/hợp mất không bị xung ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Chính ấn không bị khắc/hợp mất không bị xung ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn không bị khắc/hợp mất không bị xung ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')))),' → Quy tắc (P2 tr.533): Trụ ngày có quan sát hoặc ấn, chữ đó không bị khắc/hợp/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f153(E, r, c):
    # Ô K103
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&
    #   "; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can
    #   ",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can
    #   ",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Th
    #   iên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ
    #   ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng
    #    tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS
    #   (SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.534-535): 
    #   Thân vượng, tài có lực và có quan sát → "&IF(J103="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J103&") ⇒ "&G103&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.534-535): Thân vượng, tài có lực và có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f154(E, r, c):
    # Ô E104, F104, E204, F204, E230, F230
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,S
    #   AO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20
    #   ,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0))) else '—')

def f155(E, r, c):
    # Ô K104
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,
    #   "Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" ch
    #   ữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng 
    #   tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Ấn (can l
    #   ộ hoặc bản khí) vượng tại lệnh → "&IF(J104="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J104&") ⇒ "&G104&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.535): Ấn (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f156(E, r, c):
    # Ô E105, F105, E126, F126
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),0.0)) else '—')

def f157(E, r, c):
    # Ô K105
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L
    #   $20,0,SAO!$M$5:$M$20,0)>0,"Chính quan không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,
    #   0)>0,"Thất sát không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$
    #   M$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.536): Có quan sát và quan sát còn nguyên (không bị khắc/hợp/xung) → "&IF(J105="","chưa đủ điều kiện ở cả hai 
    #   phương pháp – không áp dụng.","ĐỦ ("&J105&") ⇒ "&G105&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.536): Có quan sát và quan sát còn nguyên (không bị khắc/hợp/xung) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f158(E, r, c):
    # Ô E106, F106
    # = IF(AND((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)=1),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0))),"Đ
    #   Ủ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),1.0), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)))) else '—')

def f159(E, r, c):
    # Ô K106
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)>0,"Chính quan: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUN
    #   TIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.536): Chính quan chỉ 1 chữ, không có thất sát → "&I
    #   F(J106="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J106&") ⇒ "&G106&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.536): Chính quan chỉ 1 chữ, không có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f160(E, r, c):
    # Ô E107, F107
    # = IF(AND((PP1_DINH_TINH!$F$99>=1),OR((THAN_SAT!$H$85="Có"),(THAN_SAT!$H$86="Có"))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,99,6),1.0), X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có')))) else '—')

def f161(E, r, c):
    # Ô K107
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.536-537): Có quan sát và có Thiê
    #   n/Nguyệt đức → "&IF(J107="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J107&") ⇒ "&G107&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.536-537): Có quan sát và có Thiên/Nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f162(E, r, c):
    # Ô E108, F108
    # = IF(AND((PP1_DINH_TINH!$F$99>=1),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,99,6),1.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f163(E, r, c):
    # Ô K108
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC
    #   !$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.536-537): Có quan sát và có ấn → "&IF(J108="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J108&") ⇒ "&G108&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.536-537): Có quan sát và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f164(E, r, c):
    # Ô E109, F109
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),0.0), X.ge(E.v(9,99,6),1.0))) else '—')

def f165(E, r, c):
    # Ô K109
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): 
    #   "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Chính ấn thiên can ở tháng: "
    #   &COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"
    #   Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Thiên ấn thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Can tháng là ấn và có quan sát → "&IF
    #   (J109="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J109&") ⇒ "&G109&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính ấn thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.535): Can tháng là ấn và có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f166(E, r, c):
    # Ô E110, F110
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$B$5:$B$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),0.0)) else '—')

def f167(E, r, c):
    # Ô K110
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,0)>0,"Chính quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Thất sát ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.536): Có quan sát ở trụ năm → "&IF(J110="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J110&") ⇒ "&G110&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.536): Có quan sát ở trụ năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f168(E, r, c):
    # Ô K111
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Có thương quan và có tài → "&IF(J111="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J111&") ⇒ "&G111&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.535): Có thương quan và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f169(E, r, c):
    # Ô E112, F112, E296, F296
    # = IF((PP1_DINH_TINH!$F$99=0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(9,99,6),0.0)) else '—')

def f170(E, r, c):
    # Ô K112
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.541): Không có quan sát → "&IF(J
    #   112="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J112&") ⇒ "&G112&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.541): Không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f171(E, r, c):
    # Ô E113, F113
    # = IF(AND((PP1_DINH_TINH!$F$99=0),(PP1_DINH_TINH!$F$98=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(9,99,6),0.0), X.eq(E.v(9,98,6),0.0))) else '—')

def f172(E, r, c):
    # Ô K113
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.541): Không có quan sát và không có tài → "&IF(J113="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J113&") ⇒ "&G113&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.541): Không có quan sát và không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f173(E, r, c):
    # Ô E114, F114
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),0.0)) else '—')

def f174(E, r, c):
    # Ô K114
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can",SAO!$L$5:$L$20,1)>0,"Chính quan thiên can bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can",SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$2
    #   0,1)>0,"Thất sát thiên can bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20
    #   ,1)&" chữ; ",""))&" → Quy tắc (P2 tr.541): Quan sát lộ ở thiên can bị hợp mất hoặc bị khắc → "&IF(J114="","chưa đủ điều kiện ở cả hai phương pháp – kh
    #   ông áp dụng.","ĐỦ ("&J114&") ⇒ "&G114&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan thiên can bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.541): Quan sát lộ ở thiên can bị hợp mất hoặc bị khắc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f175(E, r, c):
    # Ô E115, F115
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$K$5:$K$20,8))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),0.0)) else '—')

def f176(E, r, c):
    # Ô K115
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K
    #   $20,8)>0,"Chính quan lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF(COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thất sát lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ",""))&" → Quy tắc (P2 tr.541): Có chữ quan sát đóng ở đất Mộ (12 cung của can quan tại chi trụ đó) → "&IF(J115
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J115&") ⇒ "&G115&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else ''),(X.cat(X.cat('Thất sát lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')))),' → Quy tắc (P2 tr.541): Có chữ quan sát đóng ở đất Mộ (12 cung của can quan tại chi trụ đó) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f177(E, r, c):
    # Ô E116, F116, E139, F139, E286, F286
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f178(E, r, c):
    # Ô K116
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.541-542)
    #   : Có thương quan và có chính quan → "&IF(J116="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J116&") ⇒ "&G116&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.541-542): Có thương quan và có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f179(E, r, c):
    # Ô E117, F117
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f180(E, r, c):
    # Ô K117
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y
    #   $5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.541): Thương quan ≥ P_NHIEU chữ
    #    → "&IF(J117="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J117&") ⇒ "&G117&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.541): Thương quan ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f181(E, r, c):
    # Ô E118, F118
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$97=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.eq(E.v(9,97,6),0.0))) else '—')

def f182(E, r, c):
    # Ô K118
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt
    #   ): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.541): Có thất sát, không có thực thương chế → "&IF(J118="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J118&") ⇒ "&G118&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.541): Có thất sát, không có thực thương chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f183(E, r, c):
    # Ô E119, F119
    # = IF(AND((PP1_DINH_TINH!$F$99>=1),(PP1_DINH_TINH!$F$100=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,99,6),1.0), X.eq(E.v(9,100,6),0.0))) else '—')

def f184(E, r, c):
    # Ô K119
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC
    #   !$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.541): Có quan sát nhưng không có ấn → "&IF(J119="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J119&") ⇒ "&G119&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.541): Có quan sát nhưng không có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f185(E, r, c):
    # Ô E120, F120
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0)+COUNTIFS(SAO!$F$
    #   5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0)),0.0)))) else '—')

def f186(E, r, c):
    # Ô K120
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M
    #   $20,1)>0,"Chính quan bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thất sát bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0)>0,"Chính qu
    #   an không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"
    #   Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,0)>0,"Thất sát không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$M$5:$M$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.541-542): Mọi chữ quan sát đều ở chi bị xung → "&IF(J120="","chưa đủ điều kiện ở cả hai p
    #   hương pháp – không áp dụng.","ĐỦ ("&J120&") ⇒ "&G120&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')),(X.cat(X.cat('Thất sát không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.541-542): Mọi chữ quan sát đều ở chi bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f187(E, r, c):
    # Ô E121, F121
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0))) else '—')

def f188(E, r, c):
    # Ô K121
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can")>0,"Chính quan thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COU
    #   NTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.541): Chính quan và thất sát cùng lộ ở thiên can (quan sát hỗn 
    #   tạp) → "&IF(J121="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J121&") ⇒ "&G121&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.541): Chính quan và thất sát cùng lộ ở thiên can (quan sát hỗn tạp) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f189(E, r, c):
    # Ô E122
    # = IF(AND((DUNG_HY_KY!$C$7=0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f190(E, r, c):
    # Ô F122
    # = IF(AND((DUNG_HY_KY!$D$7=0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f191(E, r, c):
    # Ô K122
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1)&" chữ; ",""))&" → Quy tắc (P2 tr.542): Thân nhược và có chính quan → "&IF(J122="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ (
    #   "&J122&") ⇒ "&G122&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.542): Thân nhược và có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f192(E, r, c):
    # Ô E123, F123
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f193(E, r, c):
    # Ô K123
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:
    #   $Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.542): Thất sát ≥ P_NHIEU chữ → "&IF(J123
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J123&") ⇒ "&G123&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.542): Thất sát ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f194(E, r, c):
    # Ô E124
    # = IF(AND((DUNG_HY_KY!$C$7=0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f195(E, r, c):
    # Ô F124
    # = IF(AND((DUNG_HY_KY!$D$7=0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$
    #   5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f196(E, r, c):
    # Ô K124
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUN
    #   TIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thấ
    #   t sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất 
    #   sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Thân nhược, quan sát có lực (can lộ hoặc bản khí vượng tại lệnh) → "&IF
    #   (J124="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J124&") ⇒ "&G124&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.527): Thân nhược, quan sát có lực (can lộ hoặc bản khí vượng tại lệnh) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f197(E, r, c):
    # Ô K125
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.546): 
    #   Có đủ tài, quan sát, ấn → "&IF(J125="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J125&") ⇒ "&G125&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.546): Có đủ tài, quan sát, ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f198(E, r, c):
    # Ô K126
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L
    #   $20,0,SAO!$M$5:$M$20,0)>0,"Chính quan không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$L$5:$L$20,0,SAO!$M$5:$M$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$M$20,
    #   0)>0,"Thất sát không bị khắc/hợp mất không bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,0,SAO!$M$5:$
    #   M$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.546): Có quan sát còn nguyên (không bị khắc/hợp/xung) → "&IF(J126="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J126&") ⇒ "&G126&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát không bị khắc/hợp mất không bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 0.0, E.rg(13,5,13,20,13), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.546): Có quan sát còn nguyên (không bị khắc/hợp/xung) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f199(E, r, c):
    # Ô E127
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,99,6),1.0))) else '—')

def f200(E, r, c):
    # Ô F127
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,99,6),1.0))) else '—')

def f201(E, r, c):
    # Ô K127
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""
    #   ," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.546, 552): Thân vượng và có quan sát → "&IF(J127="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J127&") ⇒ "&G127&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.546, 552): Thân vượng và có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f202(E, r, c):
    # Ô E128, F128, E195, F195, E242, F242
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f203(E, r, c):
    # Ô K128
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.546): Có thương quan và có ấn chế thương quan → "&IF(J128="","chưa đủ điều kiện ở cả hai phương 
    #   pháp – không áp dụng.","ĐỦ ("&J128&") ⇒ "&G128&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.546): Có thương quan và có ấn chế thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f204(E, r, c):
    # Ô E129, F129, E146, F146
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,100,6),1.0))) else '—')

def f205(E, r, c):
    # Ô K129
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.546-547): Có thất sát và có ấn (sát hóa ấn) → "&IF(J129="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J129&") ⇒ "&G129&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.546-547): Có thất sát và có ấn (sát hóa ấn) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f206(E, r, c):
    # Ô E130, F130
    # = IF(AND(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên
    #    ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$
    #   J$20,1))>0)),(PP1_DINH_TINH!$F$100>COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.gt(E.v(9,100,6),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)))) else '—')

def f207(E, r, c):
    # Ô K130
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lện
    #   h: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SA
    #   O!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"
    #   Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.546, 548): Ấn nhiều hơn thất sát và ấn có lực (vượng tại lệnh) → "&IF(J130="","chưa đủ đi
    #   ều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J130&") ⇒ "&G130&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.546, 548): Ấn nhiều hơn thất sát và ấn có lực (vượng tại lệnh) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f208(E, r, c):
    # Ô E131, F131
    # = IF((PP1_DINH_TINH!$F$100>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,100,6),E.v(24,99,2))) else '—')

def f209(E, r, c):
    # Ô K131
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 
    #   tr.547): Ấn ≥ P_NHIEU chữ → "&IF(J131="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J131&") ⇒ "&G131&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.547): Ấn ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f210(E, r, c):
    # Ô E132, F132
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"
    #   Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f211(E, r, c):
    # Ô K132
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thương quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!
    #   $J$5:$J$20,1)>0,"Thương quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí
    #   ",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.547): Thương quan (can lộ hoặc bản khí) vượng tại lệnh → "&IF(J132="","chưa đủ điều kiện ở cả hai
    #    phương pháp – không áp dụng.","ĐỦ ("&J132&") ⇒ "&G132&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.547): Thương quan (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f212(E, r, c):
    # Ô E133, F133
    # = IF((THAN_SAT!$H$89="Có"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(12,89,8),'Có')) else '—')

def f213(E, r, c):
    # Ô K133
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.547): Có hoa cái → "&IF(J133="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J133
    #   &") ⇒ "&G133&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.547): Có hoa cái → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f214(E, r, c):
    # Ô E134, F134
    # = IF((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2))) else '—')

def f215(E, r, c):
    # Ô K134
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.547): Hỏa ≥ P_NHIEU chữ → "&IF(J134="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J134&") ⇒ "&G134&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.547): Hỏa ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f216(E, r, c):
    # Ô E135, F135
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$L$5:$L$20,0))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else '—')

def f217(E, r, c):
    # Ô K135
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$L$5:$L$20,0)>0,"Thực thần thiên can không bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   C$5:$C$20,"Can",SAO!$L$5:$L$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.547): Thực thần lộ ở thiên can, không bị khắc/hợp mất → "&IF(J135="","chưa đủ điều 
    #   kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J135&") ⇒ "&G135&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần thiên can không bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,12,20,12), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.547): Thực thần lộ ở thiên can, không bị khắc/hợp mất → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f218(E, r, c):
    # Ô E136, F136
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$
    #   F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0)))) else '—')

def f219(E, r, c):
    # Ô K136
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$2
    #   0,1)>0,"Chính ấn ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$
    #   F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thiên ấn ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Chính tài ở thá
    #   ng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thiên tài ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5
    #   :$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.548): Trụ tháng có ấn, không có tài → "&IF(J136="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.
    #   ","ĐỦ ("&J136&") ⇒ "&G136&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Trụ tháng có ấn, không có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f220(E, r, c):
    # Ô E137, F137
    # = IF(AND((PP1_DINH_TINH!$F$100>=1),(THAN_SAT!$H$89="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,100,6),1.0), X.eq(E.v(12,89,8),'Có'))) else '—')

def f221(E, r, c):
    # Ô K137
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.548): Có ấn và có hoa cái → "&IF(J1
    #   37="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J137&") ⇒ "&G137&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.548): Có ấn và có hoa cái → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f222(E, r, c):
    # Ô E138
    # = IF(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f223(E, r, c):
    # Ô F138
    # = IF(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f224(E, r, c):
    # Ô K138
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7=""
    #   ,""," ("&MENH_CUC!$D$7&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.548): Thân nhược, thực thương ≥ P_NHIEU chữ → "&IF(J138="","chưa 
    #   đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J138&") ⇒ "&G138&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.548): Thân nhược, thực thương ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f225(E, r, c):
    # Ô K139
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.548): Có
    #    thương quan và có chính quan → "&IF(J139="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J139&") ⇒ "&G139&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Có thương quan và có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f226(E, r, c):
    # Ô E140, F140
    # = IF((PP1_DINH_TINH!$F$96>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(E.v(9,96,6),E.v(24,99,2))) else '—')

def f227(E, r, c):
    # Ô K140
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P
    #   2 tr.548): Tỉ kiếp ≥ P_NHIEU chữ → "&IF(J140="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J140&") ⇒ "&G140&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.548): Tỉ kiếp ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f228(E, r, c):
    # Ô E141
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$
    #   5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)))),(PP1_DINH_TINH!$F$100<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))), X.le(E.v(9,100,6),1.0))) else '—')

def f229(E, r, c):
    # Ô F141
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$
    #   5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)))),(PP1_DINH_TINH!$F$100<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))), X.le(E.v(9,100,6),1.0))) else '—')

def f230(E, r, c):
    # Ô K141
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," (
    #   "&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$
    #   J$5:$J$20,1)>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$
    #   5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn
    #    bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""
    #   )&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lện
    #   h: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.
    #   548): Thân nhược và ấn không có lực (không có ấn vượng tại lệnh, ấn ≤ 1 chữ) → "&IF(J141="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.",
    #   "ĐỦ ("&J141&") ⇒ "&G141&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Thân nhược và ấn không có lực (không có ấn vượng tại lệnh, ấn ≤ 1 chữ) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f231(E, r, c):
    # Ô E142
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$100<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.le(E.v(9,100,6),1.0))) else '—')

def f232(E, r, c):
    # Ô F142
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(PP1_DINH_TINH!$F$100<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.le(E.v(9,100,6),1.0))) else '—')

def f233(E, r, c):
    # Ô K142
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" →
    #    Quy tắc (P2 tr.548): Thân nhược, tài ≥ P_NHIEU, ấn ≤ 1 chữ → "&IF(J142="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J142&") ⇒ 
    #   "&G142&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.548): Thân nhược, tài ≥ P_NHIEU, ấn ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f234(E, r, c):
    # Ô E143, F143
    # = IF(AND((PP1_DINH_TINH!$F$98>=P_NHIEU),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính
    #    tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),E.v(24,99,2)), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f235(E, r, c):
    # Ô K143
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên
    #    quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can 
    #   vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIF
    #   S(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNT
    #   IFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.548): Tài ≥ P_NHIEU và tài c
    #   ó lực → "&IF(J143="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J143&") ⇒ "&G143&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Tài ≥ P_NHIEU và tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f236(E, r, c):
    # Ô E144, F144
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.eq(E.v(9,100,6),0.0))) else '—')

def f237(E, r, c):
    # Ô K144
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.548): Có thất sát nhưng không có ấn → "&IF(J144="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J144&") ⇒ "&G144&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.548): Có thất sát nhưng không có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f238(E, r, c):
    # Ô K145
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.551): Có thương quan và có tài → "&IF(J145="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J145&") ⇒ "&G145&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.551): Có thương quan và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f239(E, r, c):
    # Ô K146
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.551): Có thất sát và có ấn → "&IF(J146="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ (
    #   "&J146&") ⇒ "&G146&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.551): Có thất sát và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f240(E, r, c):
    # Ô E147, F147
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$
    #   20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f241(E, r, c):
    # Ô K147
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính q
    #   uan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20
    #   ,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)
    #   >0,"Chính ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)
    #   &" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí v
    #   ượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy
    #    tắc (P2 tr.553): Có chính quan và ấn vượng tại lệnh → "&IF(J147="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J147&") ⇒ "&G147&
    #   ".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.553): Có chính quan và ấn vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f242(E, r, c):
    # Ô E148
    # = IF(AND((DUNG_HY_KY!$C$7=1),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f243(E, r, c):
    # Ô F148
    # = IF(AND((DUNG_HY_KY!$D$7=1),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f244(E, r, c):
    # Ô K148
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS
    #   (SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.553): Thân vượng và ấn vượng tại lệnh → "&IF(J148="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J148&") ⇒ "&G148&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.553): Thân vượng và ấn vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f245(E, r, c):
    # Ô E149
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0))),(PP1_DINH_TINH!$D$98>=1)),"Đ
    #   Ủ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0))), X.ge(E.v(9,98,4),1.0))) else '—')

def f246(E, r, c):
    # Ô F149
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")+COUNTIFS(SAO!$F$5:$F
    #   $20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0))),(PP1_DINH_TINH!$D$98>=1)),"Đ
    #   Ủ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),0.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0))), X.ge(E.v(9,98,4),1.0))) else '—')

def f247(E, r, c):
    # Ô K149
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"Các chữ liên qu
    #   an (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Chính ấn tàng 
    #   can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thiên ấn tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Chính ấn thiên
    #    can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thiên ấn thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.552-553): Thân vượng, ấn chỉ tàng (không lộ), tài lộ → "&IF(J149="","chưa đủ điều kiện ở cả hai
    #    phương pháp – không áp dụng.","ĐỦ ("&J149&") ⇒ "&G149&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else ''),(X.cat(X.cat('Thiên ấn tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')),(X.cat(X.cat('Chính ấn thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')),(X.cat(X.cat('Thiên ấn thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.552-553): Thân vượng, ấn chỉ tàng (không lộ), tài lộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f248(E, r, c):
    # Ô K150
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.554): 
    #   Có đủ tài, quan sát, ấn → "&IF(J150="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J150&") ⇒ "&G150&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.554): Có đủ tài, quan sát, ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f249(E, r, c):
    # Ô E151, F151
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0)) else '—')

def f250(E, r, c):
    # Ô K151
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,2)>0,"Chính quan ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$
    #   F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Chính quan ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.554): Có chính quan ở trụ ngày hoặc giờ → "&IF(J151="","chưa đủ điều kiện ở cả h
    #   ai phương pháp – không áp dụng.","ĐỦ ("&J151&") ⇒ "&G151&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Chính quan ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.554): Có chính quan ở trụ ngày hoặc giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f251(E, r, c):
    # Ô E152, F152
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất 
    #   sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0)) else '—')

def f252(E, r, c):
    # Ô K152
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,0)>0,"Chính quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Chính quan ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Thất s
    #   át ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thất sát ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B
    #   $5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.547): Có quan sát ở trụ năm hoặc tháng → "&IF(J152="","chưa đủ điều kiện ở cả hai phương pháp – không áp 
    #   dụng.","ĐỦ ("&J152&") ⇒ "&G152&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Chính quan ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.547): Có quan sát ở trụ năm hoặc tháng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f253(E, r, c):
    # Ô E153, F153
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5
    #   :$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))))) else '—')

def f254(E, r, c):
    # Ô K153
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên
    #    quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F
    #   $20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)
    #   >0,"Thiên tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)
    #   &" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí
    #    vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&
    #   COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.553)
    #   : Có thực thần và tài có lực → "&IF(J153="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J153&") ⇒ "&G153&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.553): Có thực thần và tài có lực → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f255(E, r, c):
    # Ô E154, F154
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$98=0)),(PP1_DINH_TINH!$F$99=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.eq(E.v(9,98,6),0.0)), X.eq(E.v(9,99,6),0.0))) else '—')

def f256(E, r, c):
    # Ô K154
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.554): 
    #   Có thương quan nhưng không có tài và không có quan sát → "&IF(J154="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J154&") ⇒ "&G15
    #   4&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.554): Có thương quan nhưng không có tài và không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f257(E, r, c):
    # Ô E155, F155
    # = IF(AND(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f258(E, r, c):
    # Ô K155
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Kiếp tài:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiê
    #   n",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Tỉ kiên: "&COUNTIFS(SAO!$F$5:$F$20,"Tỉ kiên",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên ấn: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; "
    #   ,""))&" → Quy tắc (P2 tr.554): Có đủ kiếp tài, thương quan, tỉ kiên, thiên ấn → "&IF(J155="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J155&") ⇒ "&G155&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Tỉ kiên: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Tỉ kiên', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.554): Có đủ kiếp tài, thương quan, tỉ kiên, thiên ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f259(E, r, c):
    # Ô E156
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$99<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.le(E.v(9,98,6),1.0)), X.le(E.v(9,99,6),1.0))) else '—')

def f260(E, r, c):
    # Ô F156
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$98<=1)),(PP1_DINH_TINH!$F$99<=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.le(E.v(9,98,6),1.0)), X.le(E.v(9,99,6),1.0))) else '—')

def f261(E, r, c):
    # Ô K156
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (P2 tr.552): Thân vượng nh
    #   ưng tài ≤ 1 và quan sát ≤ 1 chữ → "&IF(J156="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J156&") ⇒ "&G156&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (P2 tr.552): Thân vượng nhưng tài ≤ 1 và quan sát ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f262(E, r, c):
    # Ô E160, F160
    # = IF(AND(COUNTIF(TRUNG_GIAN!$C$6:$F$6,4)>0,COUNTIF(TRUNG_GIAN!$C$6:$F$6,10)>0,COUNTIF(TRUNG_GIAN!$C$6:$F$6,1)>0,COUNTIF(TRUNG_GIAN!$C$6:$F$6,7)>0),"ĐỦ",
    #   "—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 4.0),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 10.0),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 1.0),0.0), X.gt(X.COUNTIF(E.rg(23,6,3,6,6), 7.0),0.0))) else '—')

def f263(E, r, c):
    # Ô K160
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&" → Quy
    #    tắc (P2 tr.532): Tứ trụ có đủ Thìn, Tuất, Sửu, Mùi → "&IF(J160="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J160&") ⇒ "&G160&"
    #   .")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),' → Quy tắc (P2 tr.532): Tứ trụ có đủ Thìn, Tuất, Sửu, Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f264(E, r, c):
    # Ô E161, F161
    # = IF(AND(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=0),AND(TRUE(),TRUNG_GIAN!$D$6=0),AND(TRUE(),TRUNG_GIAN!$E$6=0),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRU
    #   NG_GIAN!$F$6=0)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),T
    #   RUNG_GIAN!$F$6=6))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"
    #   ),TRUNG_GIAN!$F$6=3))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),TRUNG_GIAN!$F$6=9))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),0.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0))))) else '—')

def f265(E, r, c):
    # Ô K161
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.532-533): Tứ trụ có đủ Tí, 
    #   Ngọ, Mão, Dậu → "&IF(J161="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J161&") ⇒ "&G161&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.532-533): Tứ trụ có đủ Tí, Ngọ, Mão, Dậu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f266(E, r, c):
    # Ô E162, F162
    # = IF(OR(AND(PP2_DINH_LUONG!$C$43=1,TRUE(),OR(AND(TRUNG_GIAN!$C$5=2,TRUNG_GIAN!$D$5=7),AND(TRUNG_GIAN!$C$5=7,TRUNG_GIAN!$D$5=2))),AND(PP2_DINH_LUONG!$D$4
    #   3=1,TRUE(),OR(AND(TRUNG_GIAN!$D$5=2,TRUNG_GIAN!$E$5=7),AND(TRUNG_GIAN!$D$5=7,TRUNG_GIAN!$E$5=2))),AND(PP2_DINH_LUONG!$E$43=1,NOT(NHAP_LIEU!$C$9="Không
    #    rõ giờ"),OR(AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$F$5=7),AND(TRUNG_GIAN!$E$5=7,TRUNG_GIAN!$F$5=2)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(10,43,3),1.0), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,3),2.0), X.eq(E.v(23,5,4),7.0)), X.AND(X.eq(E.v(23,5,3),7.0), X.eq(E.v(23,5,4),2.0)))), X.AND(X.eq(E.v(10,43,4),1.0), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,4),2.0), X.eq(E.v(23,5,5),7.0)), X.AND(X.eq(E.v(23,5,4),7.0), X.eq(E.v(23,5,5),2.0)))), X.AND(X.eq(E.v(10,43,5),1.0), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,5,6),7.0)), X.AND(X.eq(E.v(23,5,5),7.0), X.eq(E.v(23,5,6),2.0)))))) else '—')

def f267(E, r, c):
    # Ô K162
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.532): Bính và Tân kề nhau t
    #   hành ngũ hợp → "&IF(J162="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J162&") ⇒ "&G162&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.532): Bính và Tân kề nhau thành ngũ hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f268(E, r, c):
    # Ô E163, F163
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f269(E, r, c):
    # Ô K163
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.532-536): Thất sát lộ ở thiên can và có kình dương ở ch
    #   i → "&IF(J163="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J163&") ⇒ "&G163&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.532-536): Thất sát lộ ở thiên can và có kình dương ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f270(E, r, c):
    # Ô E164
    # = IF(AND(AND(TRUNG_GIAN!$E$7=0,(DUNG_HY_KY!$C$7=1)),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,7,5),0.0), X.eq(E.v(8,7,3),1.0)), X.ge(E.v(9,99,6),1.0))) else '—')

def f271(E, r, c):
    # Ô F164
    # = IF(AND(AND(TRUNG_GIAN!$E$7=0,(DUNG_HY_KY!$D$7=1)),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,7,5),0.0), X.eq(E.v(8,7,4),1.0)), X.ge(E.v(9,99,6),1.0))) else '—')

def f272(E, r, c):
    # Ô K164
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""
    #   ," ("&MENH_CUC!$D$9&")")&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!
    #   $E$9+1)&")"&" → Quy tắc (P2 tr.532): Nhật chủ Mộc, thân vượng, có quan sát (Kim) → "&IF(J164="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J164&") ⇒ "&G164&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.532): Nhật chủ Mộc, thân vượng, có quan sát (Kim) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f273(E, r, c):
    # Ô E165, F165
    # = IF(AND(AND((COUNTIFS(SAO!$R$5:$R$20,1)>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),(PP1_DINH_TINH!$F$97>=1)),"ĐỦ",
    #   "—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.ge(E.v(9,97,6),1.0))) else '—')

def f274(E, r, c):
    # Ô K165
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt
    #   ): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.532, 535): C
    #   ó kình dương, có thất sát và có thực thương chế → "&IF(J165="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J165&") ⇒ "&G165&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.532, 535): Có kình dương, có thất sát và có thực thương chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f275(E, r, c):
    # Ô E166, F166
    # = IF(AND((THAN_SAT!$H$116="Có"),SUM(TRUNG_GIAN!$C$52:$F$52)>=3),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(12,116,8),'Có'), X.ge(X.SUM(E.rg(23,52,3,52,6)),3.0))) else '—')

def f276(E, r, c):
    # Ô K166
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.532): Có tam kỳ và có tam hợp cục → "&IF(J166="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J166&") ⇒ "&G166&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.532): Có tam kỳ và có tam hợp cục → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f277(E, r, c):
    # Ô E167, F167
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f278(E, r, c):
    # Ô K167
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính q
    #   uan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SA
    #   O!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.532-533): Có chính quan và có kình dương → "&IF(J167="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J167&") ⇒ "&G167&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.532-533): Có chính quan và có kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f279(E, r, c):
    # Ô E168, F168
    # = IF(AND((PP1_DINH_TINH!$F$100>=1),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,100,6),1.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f280(E, r, c):
    # Ô K168
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.532): Có ấn và có kình dương → "&IF(J16
    #   8="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J168&") ⇒ "&G168&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.532): Có ấn và có kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f281(E, r, c):
    # Ô E169, F169
    # = IF(AND((THAN_SAT!$H$111="Có"),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(12,111,8),'Có'), X.ge(E.v(9,100,6),1.0))) else '—')

def f282(E, r, c):
    # Ô K169
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.532): Có kim thần và có ấn → "&IF(J
    #   169="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J169&") ⇒ "&G169&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.532): Có kim thần và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f283(E, r, c):
    # Ô E170, F170
    # = IF(AND((PP1_DINH_TINH!$F$100>=P_NHIEU),(PP1_DINH_TINH!$D$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,100,6),E.v(24,99,2)), X.ge(E.v(9,98,4),1.0))) else '—')

def f284(E, r, c):
    # Ô K170
    # = "Dữ kiện: "&"Số chữ Tài lộ rõ: "&PP1_DINH_TINH!$D$98&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ng
    #   ưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.533): Ấn ≥ P_NHIEU chữ và tài lộ ở thiên can → "&IF(J170="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J170&") ⇒ "&G170&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài lộ rõ: '),E.v(9,98,4)),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.533): Ấn ≥ P_NHIEU chữ và tài lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f285(E, r, c):
    # Ô E171
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),(DUNG_HY_KY!$C$7=1)),
    #   OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất s
    #   át",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0), X.eq(E.v(8,7,3),1.0)), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f286(E, r, c):
    # Ô F171
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),(DUNG_HY_KY!$D$7=1)),
    #   OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất s
    #   át",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0), X.eq(E.v(8,7,4),1.0)), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f287(E, r, c):
    # Ô K171
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Thất sát bản khí ở tháng: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.533): Bản khí chi tháng là thất sát, thân vượng và sát vượng tại lện
    #   h → "&IF(J171="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J171&") ⇒ "&G171&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thất sát bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.533): Bản khí chi tháng là thất sát, thân vượng và sát vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f288(E, r, c):
    # Ô E172, F172
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1))>0)),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0))) else '—')

def f289(E, r, c):
    # Ô K172
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"k
    #   ình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.533): Có thương quan, kiếp tài và kình dương → "&IF(J172="","chưa đủ điều kiệ
    #   n ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J172&") ⇒ "&G172&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.533): Có thương quan, kiếp tài và kình dương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f290(E, r, c):
    # Ô E173, F173
    # = IF(AND(OR((THAN_SAT!$H$85="Có"),(THAN_SAT!$H$86="Có")),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có')), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f291(E, r, c):
    # Ô K173
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.533): Có thiên/nguyệt đức và có thư
    #   ơng quan → "&IF(J173="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J173&") ⇒ "&G173&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.533): Có thiên/nguyệt đức và có thương quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f292(E, r, c):
    # Ô E174, F174
    # = IF(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$5=6),AND(TRUE(),TRUNG_GIAN!$D$5=6),AND(TRUE(),TRUNG_GIAN!$E$5=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_G
    #   IAN!$F$5=6)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=7),AND(TRUE(),TRUNG_GIAN!$D$5=7),AND(TRUE(),TRUNG_GIAN!$E$5=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG
    #   _GIAN!$F$5=7))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=5)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),AND(NOT(NHAP_LIEU!$C$9="Không r
    #   õ giờ"),TRUNG_GIAN!$F$6=9))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Khôn
    #   g rõ giờ"),TRUNG_GIAN!$F$6=1)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0)))))) else '—')

def f293(E, r, c):
    # Ô K174
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.534): Có Canh và Tân ở can,
    #    đủ Tị Dậu Sửu ở chi → "&IF(J174="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J174&") ⇒ "&G174&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.534): Có Canh và Tân ở can, đủ Tị Dậu Sửu ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f294(E, r, c):
    # Ô E175, F175
    # = IF(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),AND(TRUE(),TRUNG_GIAN!$E$5=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_G
    #   IAN!$F$5=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3),AND(TRUE(),TRUNG_GIAN!$E$5=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG
    #   _GIAN!$F$5=3))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không r
    #   õ giờ"),TRUNG_GIAN!$F$6=6))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="K
    #   hông rõ giờ"),TRUNG_GIAN!$F$6=10)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))))) else '—')

def f295(E, r, c):
    # Ô K175
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.535): Có Bính và Đinh ở can
    #   , đủ Dần Ngọ Tuất ở chi → "&IF(J175="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J175&") ⇒ "&G175&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.535): Có Bính và Đinh ở can, đủ Dần Ngọ Tuất ở chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f296(E, r, c):
    # Ô E176, F176
    # = IF(AND(TRUNG_GIAN!$E$5=8,NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8,OR(AND(TRUE(),TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN
    #   !$D$5=8,TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=8,TRUNG_GIAN!$F$6=4)
    #   )),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),8.0), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),8.0), X.eq(E.v(23,6,6),4.0))))) else '—')

def f297(E, r, c):
    # Ô K176
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.534): Ngày Nhâm, giờ Nhâm v
    #   à có trụ Nhâm Thìn → "&IF(J176="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J176&") ⇒ "&G176&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.534): Ngày Nhâm, giờ Nhâm và có trụ Nhâm Thìn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f298(E, r, c):
    # Ô E177, F177
    # = IF(OR(AND(TRUE(),OR(TRUNG_GIAN!$C$12="Chính quan",TRUNG_GIAN!$C$12="Thất sát"),PP2_DINH_LUONG!$C$50=1),AND(TRUE(),OR(TRUNG_GIAN!$D$12="Chính quan",TRU
    #   NG_GIAN!$D$12="Thất sát"),PP2_DINH_LUONG!$D$50=1),AND(TRUE(),OR(TRUNG_GIAN!$E$12="Chính quan",TRUNG_GIAN!$E$12="Thất sát"),PP2_DINH_LUONG!$E$50=1),AND
    #   (NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$12="Chính quan",TRUNG_GIAN!$F$12="Thất sát"),PP2_DINH_LUONG!$F$50=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,12,3),'Chính quan'), X.eq(E.v(23,12,3),'Thất sát')), X.eq(E.v(10,50,3),1.0)), X.AND(X.TRUE(), X.OR(X.eq(E.v(23,12,4),'Chính quan'), X.eq(E.v(23,12,4),'Thất sát')), X.eq(E.v(10,50,4),1.0)), X.AND(X.TRUE(), X.OR(X.eq(E.v(23,12,5),'Chính quan'), X.eq(E.v(23,12,5),'Thất sát')), X.eq(E.v(10,50,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,12,6),'Chính quan'), X.eq(E.v(23,12,6),'Thất sát')), X.eq(E.v(10,50,6),1.0)))) else '—')

def f299(E, r, c):
    # Ô K177
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.534): Can quan sát nằm trong cặp ngũ hợp → "&IF(J177="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J177&") ⇒ "&G177&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.534): Can quan sát nằm trong cặp ngũ hợp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f300(E, r, c):
    # Ô E178, F178
    # = IF(OR((THAN_SAT!$H$114="Có"),(THAN_SAT!$H$113="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.eq(E.v(12,114,8),'Có'), X.eq(E.v(12,113,8),'Có'))) else '—')

def f301(E, r, c):
    # Ô K178
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.534, 536): Có giáp lộc hoặc củng lộc → "&IF(J178="","chưa đủ điều kiện ở cả hai phương pháp – không á
    #   p dụng.","ĐỦ ("&J178&") ⇒ "&G178&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.534, 536): Có giáp lộc hoặc củng lộc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f302(E, r, c):
    # Ô E179, F179
    # = IF(AND((PP1_DINH_TINH!$F$100>=1),(THAN_SAT!$H$96="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,100,6),1.0), X.eq(E.v(12,96,8),'Có'))) else '—')

def f303(E, r, c):
    # Ô K179
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.534, 536): Có ấn và có thiên ất quý
    #    nhân → "&IF(J179="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J179&") ⇒ "&G179&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.534, 536): Có ấn và có thiên ất quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f304(E, r, c):
    # Ô E180, F180
    # = IF(OR(AND((TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),(SAO!$M$6=1)),AND((TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,1
    #   0,1,4)),(SAO!$M$10=1)),AND((TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),(SAO!$M$14=1)),AND((TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!
    #   $E$7+3,5)+1,7,10,10,1,4)),(SAO!$M$18=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,6,13),1.0)), X.AND(X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,10,13),1.0)), X.AND(X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,14,13),1.0)), X.AND(X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)), X.eq(E.v(13,18,13),1.0)))) else '—')

def f305(E, r, c):
    # Ô K180
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Bị xung ở năm (chi "&TRUNG_GIAN!$C$11&")
    #   : "&IF(SAO!$M$6=1,"có","không")&"; "&"Bị xung ở tháng (chi "&TRUNG_GIAN!$D$11&"): "&IF(SAO!$M$10=1,"có","không")&"; "&"Bị xung ở ngày (chi "&TRUNG_GIA
    #   N!$E$11&"): "&IF(SAO!$M$14=1,"có","không")&"; "&"Bị xung ở giờ (chi "&TRUNG_GIAN!$F$11&"): "&IF(SAO!$M$18=1,"có","không")&" → Quy tắc (P2 tr.534): Chi
    #    mộ (kho) của hành quan có trong trụ và bị xung → "&IF(J180="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J180&") ⇒ "&G180&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Bị xung ở năm (chi '),E.v(23,11,3)),'): '),('có' if X.truthy(X.eq(E.v(13,6,13),1.0)) else 'không')),'; '),'Bị xung ở tháng (chi '),E.v(23,11,4)),'): '),('có' if X.truthy(X.eq(E.v(13,10,13),1.0)) else 'không')),'; '),'Bị xung ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,13),1.0)) else 'không')),'; '),'Bị xung ở giờ (chi '),E.v(23,11,6)),'): '),('có' if X.truthy(X.eq(E.v(13,18,13),1.0)) else 'không')),' → Quy tắc (P2 tr.534): Chi mộ (kho) của hành quan có trong trụ và bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f306(E, r, c):
    # Ô E181, F181
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),OR(AND(TRUE(),TRUNG_GIAN!$C$12="Thất sát",PP2_DINH_LUONG!$C$50=1)
    #   ,AND(TRUE(),TRUNG_GIAN!$D$12="Thất sát",PP2_DINH_LUONG!$D$50=1),AND(TRUE(),TRUNG_GIAN!$E$12="Thất sát",PP2_DINH_LUONG!$E$50=1),AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),TRUNG_GIAN!$F$12="Thất sát",PP2_DINH_LUONG!$F$50=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,12,3),'Thất sát'), X.eq(E.v(10,50,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,12,4),'Thất sát'), X.eq(E.v(10,50,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,12,5),'Thất sát'), X.eq(E.v(10,50,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,12,6),'Thất sát'), X.eq(E.v(10,50,6),1.0))))) else '—')

def f307(E, r, c):
    # Ô K181
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính q
    #   uan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.533-534, 538): Có chính quan, thất sát
    #    lộ và bị hợp (hợp sát lưu quan) → "&IF(J181="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J181&") ⇒ "&G181&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.533-534, 538): Có chính quan, thất sát lộ và bị hợp (hợp sát lưu quan) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f308(E, r, c):
    # Ô E182, F182
    # = IF(AND(OR(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AND(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),
    #   TRUNG_GIAN!$F$6=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=5),AND(TRUE(),TRUNG_GIAN!$D$6=5),AND(TRUE(),TRUNG_GIAN!$E$6=5),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"
    #   ),TRUNG_GIAN!$F$6=5))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8),AND(NOT(NHAP_LIEU!$C$9="Không rõ g
    #   iờ"),TRUNG_GIAN!$F$6=8))),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!
    #   $C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7)))),(THAN_SAT!$H$96="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),5.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),5.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),5.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))), X.eq(E.v(12,96,8),'Có'))) else '—')

def f309(E, r, c):
    # Ô K182
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.534, 538): Có tam hình (Dần
    #    Tị Thân hoặc Sửu Tuất Mùi) và thiên ất quý nhân → "&IF(J182="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J182&") ⇒ "&G182&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.534, 538): Có tam hình (Dần Tị Thân hoặc Sửu Tuất Mùi) và thiên ất quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f310(E, r, c):
    # Ô E183, F183
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(THAN_SAT!$H$91="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.eq(E.v(12,91,8),'Có'))) else '—')

def f311(E, r, c):
    # Ô K183
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính q
    #   uan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Có chính quan và có đào hoa → "&
    #   IF(J183="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J183&") ⇒ "&G183&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.535): Có chính quan và có đào hoa → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f312(E, r, c):
    # Ô E184, F184
    # = IF(AND((THAN_SAT!$H$104="Có"),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(12,104,8),'Có'), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f313(E, r, c):
    # Ô K184
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Có khôi canh và có thất sát → "&IF(J18
    #   4="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J184&") ⇒ "&G184&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.535): Có khôi canh và có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f314(E, r, c):
    # Ô E185, F185
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '—')

def f315(E, r, c):
    # Ô K185
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Bản khí",SAO!$B$5:$B$20,1)>0,"Chính ấn bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản 
    #   khí",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Bản khí chi tháng là chính ấn → "&IF(J185="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J185&") ⇒ "&G185&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.535): Bản khí chi tháng là chính ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f316(E, r, c):
    # Ô E186, F186
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Chính t
    #   ài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO
    #   !$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),(
    #   (COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.add(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),0.0))) else '—')

def f317(E, r, c):
    # Ô K186
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can",SAO!$B$5:$B$20,0)>0,"Chính tài thiên can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",
    #   SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Chín
    #   h tài thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&I
    #   F(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)>0,"Thiên tài thiên can ở năm: "&COUNTIF
    #   S(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên t
    #   ài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)>0,"Thiên tài thiên can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)>0,"Chính ấn thiên can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"Can",SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B
    #   $20,3)>0,"Thiên ấn thiên can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3)&" ch
    #   ữ; ",""))&" → Quy tắc (P2 tr.535): Tài lộ ở can năm/tháng, ấn lộ ở can giờ → "&IF(J186="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","Đ
    #   Ủ ("&J186&") ⇒ "&G186&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Chính tài thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Thiên tài thiên can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn thiên can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.535): Tài lộ ở can năm/tháng, ấn lộ ở can giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f318(E, r, c):
    # Ô E187
    # = IF(AND(AND((THAN_SAT!$H$89="Có"),(PP1_DINH_TINH!$F$100>=1)),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(12,89,8),'Có'), X.ge(E.v(9,100,6),1.0)), X.eq(E.v(8,7,3),1.0))) else '—')

def f319(E, r, c):
    # Ô F187
    # = IF(AND(AND((THAN_SAT!$H$89="Có"),(PP1_DINH_TINH!$F$100>=1)),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(12,89,8),'Có'), X.ge(E.v(9,100,6),1.0)), X.eq(E.v(8,7,4),1.0))) else '—')

def f320(E, r, c):
    # Ô K187
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," (
    #   "&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.535): Có hoa cái, có ấn và thân vượng → "&IF(J187="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.
    #   ","ĐỦ ("&J187&") ⇒ "&G187&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.535): Có hoa cái, có ấn và thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f321(E, r, c):
    # Ô E188, F188
    # = IF(AND(AND(OR(NOT(TRUE()),AND(MOD(TRUNG_GIAN!$C$5,2)=0,MOD(TRUNG_GIAN!$C$6,2)=0)),OR(NOT(TRUE()),AND(MOD(TRUNG_GIAN!$D$5,2)=0,MOD(TRUNG_GIAN!$D$6,2)=0
    #   )),OR(NOT(TRUE()),AND(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!$E$6,2)=0)),OR(NOT(NOT(NHAP_LIEU!$C$9="Không rõ giờ")),AND(MOD(TRUNG_GIAN!$F$5,2)=0,MOD(
    #   TRUNG_GIAN!$F$6,2)=0))),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIF
    #   S(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.NOT(X.TRUE()), X.AND(X.eq(X.MOD(E.v(23,5,3), 2.0),0.0), X.eq(X.MOD(E.v(23,6,3), 2.0),0.0))), X.OR(X.NOT(X.TRUE()), X.AND(X.eq(X.MOD(E.v(23,5,4), 2.0),0.0), X.eq(X.MOD(E.v(23,6,4), 2.0),0.0))), X.OR(X.NOT(X.TRUE()), X.AND(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0), X.eq(X.MOD(E.v(23,6,5), 2.0),0.0))), X.OR(X.NOT(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))), X.AND(X.eq(X.MOD(E.v(23,5,6), 2.0),0.0), X.eq(X.MOD(E.v(23,6,6), 2.0),0.0)))), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f322(E, r, c):
    # Ô K188
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có m
    #   ặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên ấn thiên can vượng
    #    tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!
    #   $F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(SAO!
    #   $F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.535): Tứ trụ thuần
    #    dương và thiên ấn vượng tại lệnh → "&IF(J188="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J188&") ⇒ "&G188&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.535): Tứ trụ thuần dương và thiên ấn vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f323(E, r, c):
    # Ô E189, F189
    # = IF(OR(OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0))>0),((COUNTIFS(SAO!$F$5:$F$
    #   20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,0))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"C
    #   an",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,0))>0),((COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,0))>0))),OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!
    #   $C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)
    #   )>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên
    #    ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20
    #   ,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0))),OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"
    #   ,SAO!$B$5:$B$20,2))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!
    #   $F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$B$5:$B$20,2))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,2))>0))),
    #   OR(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3))>0),((COUNTIFS(SAO!$F$5:$F$20,"Ch
    #   ính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$2
    #   0,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3))>0)),AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SA
    #   O!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$B$5:$B$20,3))>0),((COUNTIFS(SAO!$F$5:$F
    #   $20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,3))>0)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0)),0.0)), X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 0.0)),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 0.0),0.0))), X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),0.0)), X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 1.0)),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0))), X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0)),0.0)), X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 2.0)),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 2.0),0.0))), X.OR(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0)),0.0)), X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,2,20,2), 3.0)),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 3.0),0.0))))) else '—')

def f324(E, r, c):
    # Ô K189
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.535): Thất sát và ấn cùng một trụ (can – bản khí chi) → "&IF(J189="","chưa đủ điều kiện ở cả hai phươ
    #   ng pháp – không áp dụng.","ĐỦ ("&J189&") ⇒ "&G189&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.535): Thất sát và ấn cùng một trụ (can – bản khí chi) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f325(E, r, c):
    # Ô E190, F190
    # = IF(OR(AND(TRUE(),SAO!$R$9=1,TRUNG_GIAN!$D$12="Chính quan"),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$R$17=1,TRUNG_GIAN!$F$12="Chính quan")),"ĐỦ","—"
    #   )
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,9,18),1.0), X.eq(E.v(23,12,4),'Chính quan')), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,18),1.0), X.eq(E.v(23,12,6),'Chính quan')))) else '—')

def f326(E, r, c):
    # Ô K190
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.536): Trụ tháng hoặc giờ: chi là kình dương, can là chính quan → "&IF(J190="","chưa đủ điều kiện ở cả
    #    hai phương pháp – không áp dụng.","ĐỦ ("&J190&") ⇒ "&G190&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.536): Trụ tháng hoặc giờ: chi là kình dương, can là chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f327(E, r, c):
    # Ô E191, F191
    # = IF(OR(AND(TRUE(),SAO!$P$5=1,OR(TRUNG_GIAN!$C$12="Chính quan",TRUNG_GIAN!$C$12="Thất sát")),AND(TRUE(),SAO!$P$9=1,OR(TRUNG_GIAN!$D$12="Chính quan",TRUN
    #   G_GIAN!$D$12="Thất sát")),AND(TRUE(),SAO!$P$13=1,OR(TRUNG_GIAN!$E$12="Chính quan",TRUNG_GIAN!$E$12="Thất sát")),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ")
    #   ,SAO!$P$17=1,OR(TRUNG_GIAN!$F$12="Chính quan",TRUNG_GIAN!$F$12="Thất sát"))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0), X.OR(X.eq(E.v(23,12,3),'Chính quan'), X.eq(E.v(23,12,3),'Thất sát'))), X.AND(X.TRUE(), X.eq(E.v(13,9,16),1.0), X.OR(X.eq(E.v(23,12,4),'Chính quan'), X.eq(E.v(23,12,4),'Thất sát'))), X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0), X.OR(X.eq(E.v(23,12,5),'Chính quan'), X.eq(E.v(23,12,5),'Thất sát'))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0), X.OR(X.eq(E.v(23,12,6),'Chính quan'), X.eq(E.v(23,12,6),'Thất sát'))))) else '—')

def f328(E, r, c):
    # Ô K191
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.536): Có trụ mà chi là thiên ất quý nhân và can là quan sát → "&IF(J191="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J191&") ⇒ "&G191&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.536): Có trụ mà chi là thiên ất quý nhân và can là quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f329(E, r, c):
    # Ô E192, F192
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(THAN_SAT!$H$85="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.eq(E.v(12,85,8),'Có'))) else '—')

def f330(E, r, c):
    # Ô K192
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính q
    #   uan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.537): Có chính quan và có thiên đức → 
    #   "&IF(J192="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J192&") ⇒ "&G192&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.537): Có chính quan và có thiên đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f331(E, r, c):
    # Ô E193, F193
    # = IF(AND(AND((PP1_DINH_TINH!$F$99>=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),OR((THAN_SAT!$H$122="Có"),(THAN_SAT!$
    #   H$90="Có"))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,99,6),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.eq(E.v(12,122,8),'Có'), X.eq(E.v(12,90,8),'Có')))) else '—')

def f332(E, r, c):
    # Ô K193
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): 
    #   "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.537): Có quan sát, thất sát và (vong thần hoặc kiếp sát) → "&IF(J193="","chưa đủ điều kiện ở cả hai ph
    #   ương pháp – không áp dụng.","ĐỦ ("&J193&") ⇒ "&G193&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.537): Có quan sát, thất sát và (vong thần hoặc kiếp sát) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f333(E, r, c):
    # Ô E194, F194
    # = IF(AND((COUNTIFS(SAO!$R$5:$R$20,1)>0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)
    #   )>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f334(E, r, c):
    # Ô K194
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   )&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh:
    #    "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20
    #   ,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.537): Có kình dương và thất 
    #   sát vượng tại lệnh → "&IF(J194="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J194&") ⇒ "&G194&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.537): Có kình dương và thất sát vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f335(E, r, c):
    # Ô K195
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.537, 539): Có thương quan và có ấn chế → "&IF(J195="","chưa đủ điều kiện ở cả hai phương pháp – 
    #   không áp dụng.","ĐỦ ("&J195&") ⇒ "&G195&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.537, 539): Có thương quan và có ấn chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f336(E, r, c):
    # Ô E196, F196
    # = IF(AND(AND((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>=2),(COUNTIFS(SAO!$R$5:$R$20,1)>0)),(PP1_DINH_TINH!$F$97=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),2.0), X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)), X.eq(E.v(9,97,6),0.0))) else '—')

def f337(E, r, c):
    # Ô K196
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt
    #   ): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ","")&IF(
    #   COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.536): Thất sát ≥ 2 chữ, có kình dương, không c
    #   ó thực thương chế → "&IF(J196="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J196&") ⇒ "&G196&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else ''),(X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.536): Thất sát ≥ 2 chữ, có kình dương, không có thực thương chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f338(E, r, c):
    # Ô E197, F197
    # = IF(AND(MOD(TRUNG_GIAN!$C$9+2,5)=TRUNG_GIAN!$E$7,(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(X.MOD(X.add(E.v(23,9,3),2.0), 5.0),E.v(23,7,5)), X.ge(E.v(9,99,6),1.0))) else '—')

def f339(E, r, c):
    # Ô K197
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C
    #   $11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDE
    #   X(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.543): Hành chi năm khắc hành nhật can, có quan sát → "&IF(J
    #   197="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J197&") ⇒ "&G197&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.543): Hành chi năm khắc hành nhật can, có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f340(E, r, c):
    # Ô E201, F201
    # = IF(OR(AND(TRUE(),SAO!$P$13=1),AND(TRUE(),SAO!$Q$13=1),AND(TRUE(),LEN(THAN_SAT!$E$92)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,17),1.0)), X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0)))) else '—')

def f341(E, r, c):
    # Ô K201
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.546): Chi ngày có quý nhân hoặc lộc hoặc trạch mã → "&IF(J201="","chưa đủ điều kiện ở cả hai phương p
    #   háp – không áp dụng.","ĐỦ ("&J201&") ⇒ "&G201&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.546): Chi ngày có quý nhân hoặc lộc hoặc trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f342(E, r, c):
    # Ô E202
    # = IF(AND(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU)),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2))), X.eq(E.v(8,7,3),1.0))) else '—')

def f343(E, r, c):
    # Ô F202
    # = IF(AND(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU)),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2))), X.eq(E.v(8,7,4),1.0))) else '—')

def f344(E, r, c):
    # Ô K202
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.546): Hỏa 
    #   ≥ P_NHIEU, Mộc ≥ P_NHIEU và thân vượng → "&IF(J202="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J202&") ⇒ "&G202&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.546): Hỏa ≥ P_NHIEU, Mộc ≥ P_NHIEU và thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f345(E, r, c):
    # Ô E203, F203
    # = IF(AND((TRUNG_GIAN!$E$5=9),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),9.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0)))) else '—')

def f346(E, r, c):
    # Ô K203
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.5
    #   47): Nhật can Quý, chi giờ Dần → "&IF(J203="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J203&") ⇒ "&G203&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.547): Nhật can Quý, chi giờ Dần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f347(E, r, c):
    # Ô K204
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,
    #   "Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" ch
    #   ữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng 
    #   tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.547): Ấn (can l
    #   ộ hoặc bản khí) vượng tại lệnh – chờ vận ấn ở Đế vượng → "&IF(J204="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J204&") ⇒ "&G20
    #   4&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.547): Ấn (can lộ hoặc bản khí) vượng tại lệnh – chờ vận ấn ở Đế vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f348(E, r, c):
    # Ô E205, F205
    # = IF(AND(OR((TRUNG_GIAN!$E$5=4),(TRUNG_GIAN!$E$5=5)),AND(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E
    #   $6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIA
    #   N!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUN
    #   G_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),T
    #   RUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,5),5.0)), X.AND(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0)))))) else '—')

def f349(E, r, c):
    # Ô K205
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.547): Nhật can Mậu/Kỷ, đủ T
    #   hìn, Tuất, Sửu, Mùi → "&IF(J205="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J205&") ⇒ "&G205&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.547): Nhật can Mậu/Kỷ, đủ Thìn, Tuất, Sửu, Mùi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f350(E, r, c):
    # Ô E206, F206
    # = IF(AND(AND(AND((TRUNG_GIAN!$E$5=9),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=9)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=11),AND(TRUE(),TRUNG_GIAN!
    #   $D$6=11),AND(TRUE(),TRUNG_GIAN!$E$6=11),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=11))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG
    #   _GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(23,5,5),9.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),9.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),11.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),11.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),11.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))))) else '—')

def f351(E, r, c):
    # Ô K206
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.547): Nhật can Quý, can giờ
    #    Quý, có Hợi và Sửu → "&IF(J206="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J206&") ⇒ "&G206&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.547): Nhật can Quý, can giờ Quý, có Hợi và Sửu → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f352(E, r, c):
    # Ô E207, F207
    # = IF(OR(AND(AND(TRUE(),LEN(THAN_SAT!$C$92)>0),LEN(THAN_SAT!$C$106)>0),AND(AND(TRUE(),LEN(THAN_SAT!$D$92)>0),LEN(THAN_SAT!$D$106)>0),AND(AND(TRUE(),LEN(T
    #   HAN_SAT!$E$92)>0),LEN(THAN_SAT!$E$106)>0),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$92)>0),LEN(THAN_SAT!$F$106)>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,3)),0.0)), X.gt(X.LEN(E.v(12,106,3)),0.0)), X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,4)),0.0)), X.gt(X.LEN(E.v(12,106,4)),0.0)), X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0)), X.gt(X.LEN(E.v(12,106,5)),0.0)), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0)), X.gt(X.LEN(E.v(12,106,6)),0.0)))) else '—')

def f353(E, r, c):
    # Ô K207
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.547): Trạch mã và học đường cùng một trụ → "&IF(J207="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J207&") ⇒ "&G207&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.547): Trạch mã và học đường cùng một trụ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f354(E, r, c):
    # Ô E208, F208
    # = IF(AND(OR(AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$
    #   E$5=4,TRUNG_GIAN!$E$6=10)),OR(OR(AND(TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$C$6=4),AND(TRUNG_GIAN!$C$5=6,TRUNG_GIAN!$C$6=10),AND(TRUNG_GIAN!$C$5=6,TRUNG_GIAN!$
    #   C$6=4),AND(TRUNG_GIAN!$C$5=4,TRUNG_GIAN!$C$6=10)),OR(AND(TRUNG_GIAN!$D$5=8,TRUNG_GIAN!$D$6=4),AND(TRUNG_GIAN!$D$5=6,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN
    #   !$D$5=6,TRUNG_GIAN!$D$6=4),AND(TRUNG_GIAN!$D$5=4,TRUNG_GIAN!$D$6=10)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$F$5=8,TRUNG_GIAN!$F$6=
    #   4),AND(TRUNG_GIAN!$F$5=6,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$F$5=6,TRUNG_GIAN!$F$6=4),AND(TRUNG_GIAN!$F$5=4,TRUNG_GIAN!$F$6=10))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),10.0))), X.OR(X.OR(X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,6,3),4.0)), X.AND(X.eq(E.v(23,5,3),6.0), X.eq(E.v(23,6,3),10.0)), X.AND(X.eq(E.v(23,5,3),6.0), X.eq(E.v(23,6,3),4.0)), X.AND(X.eq(E.v(23,5,3),4.0), X.eq(E.v(23,6,3),10.0))), X.OR(X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,6,4),4.0)), X.AND(X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,6,4),4.0)), X.AND(X.eq(E.v(23,5,4),4.0), X.eq(E.v(23,6,4),10.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,6),8.0), X.eq(E.v(23,6,6),4.0)), X.AND(X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,6,6),4.0)), X.AND(X.eq(E.v(23,5,6),4.0), X.eq(E.v(23,6,6),10.0))))))) else '—')

def f355(E, r, c):
    # Ô K208
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.547): Trụ ngày là khôi canh
    #    và có thêm trụ khác cũng là khôi canh → "&IF(J208="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J208&") ⇒ "&G208&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.547): Trụ ngày là khôi canh và có thêm trụ khác cũng là khôi canh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f356(E, r, c):
    # Ô E209, F209
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0),0.0)) else '—')

def f357(E, r, c):
    # Ô K209
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$2
    #   0,1)>0,"Chính ấn lộc: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.548): 
    #   Chính ấn đóng ở chi lộc của nhật can → "&IF(J209="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J209&") ⇒ "&G209&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính ấn lộc: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.548): Chính ấn đóng ở chi lộc của nhật can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f358(E, r, c):
    # Ô E210, F210
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$97>=P_NHIEU)),(PP1_DINH_TINH!$F$97>COUNTIFS(S
    #   AO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,97,6),E.v(24,99,2))), X.gt(E.v(9,97,6),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)))) else '—')

def f359(E, r, c):
    # Ô K210
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các 
    #   chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)>0,"Thất sát: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.548): Có thất sát, thực thương ≥ P_NHIEU và nhiều hơn thất sát (chế quá) →
    #    "&IF(J210="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J210&") ⇒ "&G210&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Có thất sát, thực thương ≥ P_NHIEU và nhiều hơn thất sát (chế quá) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f360(E, r, c):
    # Ô E211
    # = IF(AND((DUNG_HY_KY!$C$7=1),TRUNG_GIAN!$D$28<=4),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.le(E.v(23,28,4),4.0))) else '—')

def f361(E, r, c):
    # Ô F211
    # = IF(AND((DUNG_HY_KY!$D$7=1),TRUNG_GIAN!$D$28<=4),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.le(E.v(23,28,4),4.0))) else '—')

def f362(E, r, c):
    # Ô K211
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&" → Quy tắc (P2 tr.548): Thân vượng và nhật can được lệnh (Trường s
    #   inh – Đế vượng tại chi tháng) → "&IF(J211="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J211&") ⇒ "&G211&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),' → Quy tắc (P2 tr.548): Thân vượng và nhật can được lệnh (Trường sinh – Đế vượng tại chi tháng) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f363(E, r, c):
    # Ô E212, F212
    # = IF(AND(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$
    #   20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$
    #   5:$L$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0)), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0))) else '—')

def f364(E, r, c):
    # Ô K212
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$2
    #   0,1)>0,"Chính ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SA
    #   O!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"
    #   Chính ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Th
    #   iên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Chính tài bị khắc/hợp m
    #   ất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên tài bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,
    #   SAO!$L$5:$L$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.548): Ấn bị khắc/xung và tài bị khắc → "&IF(J212="","chưa đủ điều kiện ở cả hai phương pháp – không
    #    áp dụng.","ĐỦ ("&J212&") ⇒ "&G212&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Ấn bị khắc/xung và tài bị khắc → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f365(E, r, c):
    # Ô E213, F213
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thấ
    #   t sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0))) else '—')

def f366(E, r, c):
    # Ô K213
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1
    #   )>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1
    #   )&" chữ; ",""))&" → Quy tắc (P2 tr.548): Thất sát (can lộ hoặc bản khí) vượng tại lệnh → "&IF(J213="","chưa đủ điều kiện ở cả hai phương pháp – không 
    #   áp dụng.","ĐỦ ("&J213&") ⇒ "&G213&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.548): Thất sát (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f367(E, r, c):
    # Ô E214, F214
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0))) else '—')

def f368(E, r, c):
    # Ô K214
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$
    #   20,0)>0,"Thực thần ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:
    #   $F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Thực thần ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)>0,"Thất sát ở n
    #   gày: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Thất sát ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$2
    #   0,3)&" chữ; ",""))&" → Quy tắc (P2 tr.551): Thực thần ở trụ năm/tháng, thất sát ở trụ ngày/giờ → "&IF(J214="","chưa đủ điều kiện ở cả hai phương pháp 
    #   – không áp dụng.","ĐỦ ("&J214&") ⇒ "&G214&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thực thần ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Thực thần ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thất sát ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.551): Thực thần ở trụ năm/tháng, thất sát ở trụ ngày/giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f369(E, r, c):
    # Ô E215, F215, E258, F258
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$P$5:$P$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$
    #   20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,16,20,16), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0))) else '—')

def f370(E, r, c):
    # Ô K215
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$P$5:$P$
    #   20,1)>0,"Thực thần quý nhân: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$P$5:$P$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)>0,"Thực thần thiên/nguyệt đức: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.551): Thực thần đóng cùng quý nhân hoặc thiên/nguyệt đức → "&IF(J215="","
    #   chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J215&") ⇒ "&G215&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần quý nhân: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,16,20,16), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,16,20,16), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.551): Thực thần đóng cùng quý nhân hoặc thiên/nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f371(E, r, c):
    # Ô E216, F216
    # = IF(AND(AND(AND((TRUNG_GIAN!$E$5=0),((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=2),AN
    #   D(TRUE(),TRUNG_GIAN!$D$6=2),AND(TRUE(),TRUNG_GIAN!$E$6=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=2))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6)
    #   ,AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=6))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(23,5,5),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0))))) else '—')

def f372(E, r, c):
    # Ô K216
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có m
    #   ặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.551): Nhật can Giáp, có thương quan, có Dần và Ngọ → "&IF(J216="","chưa đủ điều kiện ở c
    #   ả hai phương pháp – không áp dụng.","ĐỦ ("&J216&") ⇒ "&G216&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.551): Nhật can Giáp, có thương quan, có Dần và Ngọ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f373(E, r, c):
    # Ô E217, F217
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can"))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can"))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)))) else '—')

def f374(E, r, c):
    # Ô K217
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can")>0,"Chính quan thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COU
    #   NTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")>0,"Thất sát tàng can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5
    #   :$C$20,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Qu
    #   y tắc (P2 tr.552): Chính quan lộ ở can, thất sát chỉ tàng trong chi → "&IF(J217="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J2
    #   17&") ⇒ "&G217&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát tàng can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can'),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.552): Chính quan lộ ở can, thất sát chỉ tàng trong chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f375(E, r, c):
    # Ô E218, F218
    # = IF(AND((COUNTIFS(SAO!$R$5:$R$20,1)>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f376(E, r, c):
    # Ô K218
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$R$5:$R$20,1)>0,"kình dương: "&COUNTIFS(SAO!$R$5:$R$20,1)&" chữ; ",""
    #   )&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.552): Có kình dương và có thất sát → "&IF(J218="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng."
    #   ,"ĐỦ ("&J218&") ⇒ "&G218&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('kình dương: ',X.COUNTIFS(E.rg(13,5,18,20,18), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,18,20,18), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.552): Có kình dương và có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f377(E, r, c):
    # Ô E219, F219
    # = IF(AND(MOD(TRUNG_GIAN!$E$5,2)=0,(IF(AND(TRUE(),SAO!$R$5=1),1,0)+IF(AND(TRUE(),SAO!$R$9=1),1,0)+IF(AND(TRUE(),SAO!$R$13=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C
    #   $9="Không rõ giờ"),SAO!$R$17=1),1,0))>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0), X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(13,5,18),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(13,9,18),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.eq(E.v(13,13,18),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,18),1.0))) else 0.0)),2.0))) else '—')

def f378(E, r, c):
    # Ô K219
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.552): Nhật can dương, kình dương ≥ 2 chi → "&IF(J219="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J219&") ⇒ "&G219
    #   &".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.552): Nhật can dương, kình dương ≥ 2 chi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f379(E, r, c):
    # Ô E220, F220
    # = IF(AND((TRUNG_GIAN!$E$7=3),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),AND(TRUE(),TRUNG_GIAN!$E$5=2),AND(NOT(NHAP_LIEU!$C$9="Khô
    #   ng rõ giờ"),TRUNG_GIAN!$F$5=2)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3),AND(TRUE(),TRUNG_GIAN!$E$5=3),AND(NOT(NHAP_LIEU!$C$9="K
    #   hông rõ giờ"),TRUNG_GIAN!$F$5=3)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),3.0), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0)))))) else '—')

def f380(E, r, c):
    # Ô K220
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.552): Nhật chủ Kim, có Bính
    #   /Đinh lộ → "&IF(J220="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J220&") ⇒ "&G220&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.552): Nhật chủ Kim, có Bính/Đinh lộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f381(E, r, c):
    # Ô E221, F221
    # = IF(AND(OR(TRUNG_GIAN!$D$12="Chính quan",TRUNG_GIAN!$D$12="Thất sát"),MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_GIAN!$C$7),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,12,4),'Chính quan'), X.eq(E.v(23,12,4),'Thất sát')), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,3)))) else '—')

def f382(E, r, c):
    # Ô K221
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&" → Quy tắc (P2 tr.
    #   552): Can tháng là quan sát và hành can tháng sinh hành can năm → "&IF(J221="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J221&"
    #   ) ⇒ "&G221&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),' → Quy tắc (P2 tr.552): Can tháng là quan sát và hành can tháng sinh hành can năm → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f383(E, r, c):
    # Ô E222, F222
    # = IF(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=9),AND(TRUE(),TRUNG_GIAN!$D$6=9),AND(TRUE(),TRUNG_GIAN!$E$6=9),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_G
    #   IAN!$F$6=9)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=3),AND(TRUE(),TRUNG_GIAN!$D$6=3),AND(TRUE(),TRUNG_GIAN!$E$6=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG
    #   _GIAN!$F$6=3))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TR
    #   UNG_GIAN!$F$6=6))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),9.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),9.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),9.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),3.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0))))) else '—')

def f384(E, r, c):
    # Ô K222
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.552): Có đủ Dậu, Mão, Ngọ →
    #    "&IF(J222="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J222&") ⇒ "&G222&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.552): Có đủ Dậu, Mão, Ngọ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f385(E, r, c):
    # Ô E223, F223
    # = IF((IF(AND(TRUE(),OR(TRUNG_GIAN!$C$6=2,TRUNG_GIAN!$C$6=8,TRUNG_GIAN!$C$6=5,TRUNG_GIAN!$C$6=11)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$D$
    #   6=8,TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=11)),1,0)+IF(AND(TRUE(),OR(TRUNG_GIAN!$E$6=2,TRUNG_GIAN!$E$6=8,TRUNG_GIAN!$E$6=5,TRUNG_GIAN!$E$6=11)),1,0)+IF(AN
    #   D(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$F$6=2,TRUNG_GIAN!$F$6=8,TRUNG_GIAN!$F$6=5,TRUNG_GIAN!$F$6=11)),1,0))>=3,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,3),2.0), X.eq(E.v(23,6,3),8.0), X.eq(E.v(23,6,3),5.0), X.eq(E.v(23,6,3),11.0)))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),11.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.OR(X.eq(E.v(23,6,5),2.0), X.eq(E.v(23,6,5),8.0), X.eq(E.v(23,6,5),5.0), X.eq(E.v(23,6,5),11.0)))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,6,6),2.0), X.eq(E.v(23,6,6),8.0), X.eq(E.v(23,6,6),5.0), X.eq(E.v(23,6,6),11.0)))) else 0.0)),3.0)) else '—')

def f386(E, r, c):
    # Ô K223
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.552): Có ≥ 3 chi thuộc Dần,
    #    Thân, Tị, Hợi → "&IF(J223="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J223&") ⇒ "&G223&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.552): Có ≥ 3 chi thuộc Dần, Thân, Tị, Hợi → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f387(E, r, c):
    # Ô E224, F224
    # = IF(AND(AND(OR((TRUNG_GIAN!$E$5=4),(TRUNG_GIAN!$E$5=5)),AND(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIA
    #   N!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG
    #   _GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),
    #   TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE
    #   (),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7)))),OR(OR((THAN_SAT!$H$85="Có"),(THAN_SAT!$H$86="Có")),(THAN_SAT!$H$116
    #   ="Có"))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,5,5),5.0)), X.AND(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))), X.OR(X.OR(X.eq(E.v(12,85,8),'Có'), X.eq(E.v(12,86,8),'Có')), X.eq(E.v(12,116,8),'Có')))) else '—')

def f388(E, r, c):
    # Ô K224
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.552): Nhật can Mậu/Kỷ, đủ T
    #   hìn Tuất Sửu Mùi, có thiên/nguyệt đức hoặc tam kỳ → "&IF(J224="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J224&") ⇒ "&G224&"."
    #   )
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.552): Nhật can Mậu/Kỷ, đủ Thìn Tuất Sửu Mùi, có thiên/nguyệt đức hoặc tam kỳ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f389(E, r, c):
    # Ô E225, F225
    # = IF(OR(AND(TRUE(),SAO!$P$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0)))) else '—')

def f390(E, r, c):
    # Ô K225
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.552): Chi ngày hoặc chi giờ là thiên ất quý nhân → "&IF(J225="","chưa đủ điều kiện ở cả hai phương ph
    #   áp – không áp dụng.","ĐỦ ("&J225&") ⇒ "&G225&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.552): Chi ngày hoặc chi giờ là thiên ất quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f391(E, r, c):
    # Ô E226, F226
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0))>0),AND(TRUE(),SAO!$P$5=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),0.0), X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0)))) else '—')

def f392(E, r, c):
    # Ô K226
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,0)>0,"Chính quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Thất sát ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$
    #   X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.552, 556): Trụ năm có quan sát và chi năm là quý nhân → "&IF(J226="","chưa đủ điều kiện
    #    ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J226&") ⇒ "&G226&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Thất sát ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '')))),' → Quy tắc (P2 tr.552, 556): Trụ năm có quan sát và chi năm là quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f393(E, r, c):
    # Ô E227, F227
    # = IF(AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$5=1),AND(TRUE(),TRUNG_GIAN!$D$5=1),AND(TRUE(),TRUNG_GIAN!$E$5=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_G
    #   IAN!$F$5=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=2),AND(TRUE(),TRUNG_GIAN!$D$5=2),AND(TRUE(),TRUNG_GIAN!$E$5=2),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG
    #   _GIAN!$F$5=2))),OR(AND(TRUE(),TRUNG_GIAN!$C$5=3),AND(TRUE(),TRUNG_GIAN!$D$5=3),AND(TRUE(),TRUNG_GIAN!$E$5=3),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TR
    #   UNG_GIAN!$F$5=3))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),2.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),2.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),3.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),3.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),3.0))))) else '—')

def f394(E, r, c):
    # Ô K227
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.552): Có đủ Ất, Bính, Đinh 
    #   ở thiên can → "&IF(J227="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J227&") ⇒ "&G227&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.552): Có đủ Ất, Bính, Đinh ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f395(E, r, c):
    # Ô E228, F228
    # = IF(AND(OR(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO
    #   !$J$5:$J$20,1))>0)),(PP1_DINH_TINH!$F$98>=P_NHIEU)),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)), X.ge(E.v(9,98,6),E.v(24,99,2))), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f396(E, r, c):
    # Ô K228
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên
    #    quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )>0,"Chính tài thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thiên tài thiên can 
    #   vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIF
    #   S(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính tài bản khí vượng tại lệnh: "&COUNT
    #   IFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"T
    #   hiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên tài bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.5
    #   53): Tài có lực và có chính quan → "&IF(J228="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J228&") ⇒ "&G228&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.553): Tài có lực và có chính quan → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f397(E, r, c):
    # Ô E229
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thi
    #   ên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),0.0), X.eq(E.v(8,7,3),1.0))) else '—')

def f398(E, r, c):
    # Ô F229
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thi
    #   ên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1))>0),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),0.0), X.eq(E.v(8,7,4),1.0))) else '—')

def f399(E, r, c):
    # Ô K229
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Chính tài bản khí ở tháng: "&COUNTIFS(SA
    #   O!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên t
    #   ài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)>0,"Thiên tài bản khí ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.553): Bản khí chi tháng là tài và thân vư
    #   ợng → "&IF(J229="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J229&") ⇒ "&G229&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài bản khí ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.553): Bản khí chi tháng là tài và thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f400(E, r, c):
    # Ô K230
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,
    #   "Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" ch
    #   ữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vượng 
    #   tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTIFS(S
    #   AO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.553): Ấn (can l
    #   ộ hoặc bản khí) vượng tại lệnh → "&IF(J230="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J230&") ⇒ "&G230&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.553): Ấn (can lộ hoặc bản khí) vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f401(E, r, c):
    # Ô E231, F231
    # = IF(AND((INDEX(DUNG_HY_KY!$C$24:$C$28,2)>=P_NHIEU),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(X.INDEX(E.rg(8,24,3,28,3), 2.0),E.v(24,99,2)), X.ge(X.INDEX(E.rg(8,24,3,28,3), 1.0),E.v(24,99,2)))) else '—')

def f402(E, r, c):
    # Ô K231
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.553): Hỏa ≥ P_NHIEU và Mộc ≥ P_NHIEU → "&IF(J231="","chưa đủ điều kiện ở cả hai phương
    #    pháp – không áp dụng.","ĐỦ ("&J231&") ⇒ "&G231&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.553): Hỏa ≥ P_NHIEU và Mộc ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f403(E, r, c):
    # Ô E232, F232
    # = IF(IF(MOD(TRUNG_GIAN!$E$5,2)=0,COUNTIF(THAN_SAT!$H$85:$H$86,"Có")>0,OR(AND(TRUE(),SAO!$P$5=1),AND(TRUE(),SAO!$P$9=1),AND(TRUE(),SAO!$P$13=1),AND(NOT(N
    #   HAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy((X.gt(X.COUNTIF(E.rg(12,85,8,86,8), 'Có'),0.0) if X.truthy(X.eq(X.MOD(E.v(23,5,5), 2.0),0.0)) else X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,9,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0))))) else '—')

def f404(E, r, c):
    # Ô K232
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.553): Nhật can dương có thiên/nguyệt đức, hoặc nhật can âm có quý nhân → "&IF(J232="","chưa đủ điều kiện ở cả hai phương pháp – không áp 
    #   dụng.","ĐỦ ("&J232&") ⇒ "&G232&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.553): Nhật can dương có thiên/nguyệt đức, hoặc nhật can âm có quý nhân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f405(E, r, c):
    # Ô E233
    # = IF(AND(AND((TRUNG_GIAN!$E$7=0),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=6),AND(TRUE(),TRUNG_GIAN!$D$5=6),AND(TRUE(),TRUNG_GIAN!$E$5=6),AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),TRUNG_GIAN!$F$5=6)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=7),AND(TRUE(),TRUNG_GIAN!$D$5=7),AND(TRUE(),TRUNG_GIAN!$E$5=7),AND(NOT(NHAP_LIEU!$C$
    #   9="Không rõ giờ"),TRUNG_GIAN!$F$5=7)))),(DUNG_HY_KY!$C$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,7,5),0.0), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))))), X.eq(E.v(8,7,3),1.0))) else '—')

def f406(E, r, c):
    # Ô F233
    # = IF(AND(AND((TRUNG_GIAN!$E$7=0),OR(OR(AND(TRUE(),TRUNG_GIAN!$C$5=6),AND(TRUE(),TRUNG_GIAN!$D$5=6),AND(TRUE(),TRUNG_GIAN!$E$5=6),AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),TRUNG_GIAN!$F$5=6)),OR(AND(TRUE(),TRUNG_GIAN!$C$5=7),AND(TRUE(),TRUNG_GIAN!$D$5=7),AND(TRUE(),TRUNG_GIAN!$E$5=7),AND(NOT(NHAP_LIEU!$C$
    #   9="Không rõ giờ"),TRUNG_GIAN!$F$5=7)))),(DUNG_HY_KY!$D$7=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(23,7,5),0.0), X.OR(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0))))), X.eq(E.v(8,7,4),1.0))) else '—')

def f407(E, r, c):
    # Ô K233
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUN
    #   G_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1
    #   )&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HA
    #   NH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.553): Nhật chủ Mộc, Kim lộ ở can, thân vượng → "&IF(J233="","chưa đủ điều kiện ở cả hai phương pháp –
    #    không áp dụng.","ĐỦ ("&J233&") ⇒ "&G233&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.553): Nhật chủ Mộc, Kim lộ ở can, thân vượng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f408(E, r, c):
    # Ô E234, F234
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$20,1))>0),ISNUMBER(SEARCH("Có tác dụng",KET_LUAN!$C$40))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0)),0.0), X.ISNUMBER(lambda: X.SEARCH('Có tác dụng', E.v(3,40,3))))) else '—')

def f409(E, r, c):
    # Ô K234
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$
    #   20,1)>0,"Chính tài lộc: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F
    #   $20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$Q$5:$Q$20,1)>0,"Thiên tài lộc: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:
    #   $X$20,1,SAO!$Q$5:$Q$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.554): Có tài ở chi lộc và văn xương có tác dụng → "&IF(J234="","chưa đủ điều kiện ở cả hai 
    #   phương pháp – không áp dụng.","ĐỦ ("&J234&") ⇒ "&G234&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài lộc: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài lộc: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,17,20,17), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.554): Có tài ở chi lộc và văn xương có tác dụng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f410(E, r, c):
    # Ô E235, F235
    # = IF((IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$Q$17=1),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9=
    #   "Không rõ giờ"),LEN(THAN_SAT!$F$92)>0),1,0))>=2,"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add((1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0))) else 0.0),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,17),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0))) else 0.0)),2.0)) else '—')

def f411(E, r, c):
    # Ô K235
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.554): Chi giờ có ít nhất 2 trong 3: quý nhân, lộc, trạch mã → "&IF(J235="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J235&") ⇒ "&G235&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.554): Chi giờ có ít nhất 2 trong 3: quý nhân, lộc, trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f412(E, r, c):
    # Ô E236, F236
    # = IF(AND(OR(AND(PP2_DINH_LUONG!$C$43=1,TRUE(),OR(AND(TRUNG_GIAN!$C$5=3,TRUNG_GIAN!$D$5=8),AND(TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$D$5=3))),AND(PP2_DINH_LUONG!
    #   $D$43=1,TRUE(),OR(AND(TRUNG_GIAN!$D$5=3,TRUNG_GIAN!$E$5=8),AND(TRUNG_GIAN!$D$5=8,TRUNG_GIAN!$E$5=3))),AND(PP2_DINH_LUONG!$E$43=1,NOT(NHAP_LIEU!$C$9="K
    #   hông rõ giờ"),OR(AND(TRUNG_GIAN!$E$5=3,TRUNG_GIAN!$F$5=8),AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$F$5=3)))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIA
    #   N!$F$6=2)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.eq(E.v(10,43,3),1.0), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,3),3.0), X.eq(E.v(23,5,4),8.0)), X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,5,4),3.0)))), X.AND(X.eq(E.v(10,43,4),1.0), X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,4),3.0), X.eq(E.v(23,5,5),8.0)), X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,5,5),3.0)))), X.AND(X.eq(E.v(10,43,5),1.0), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,5),3.0), X.eq(E.v(23,5,6),8.0)), X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,5,6),3.0))))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),2.0)))) else '—')

def f413(E, r, c):
    # Ô K236
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.554): Đinh – Nhâm kề nhau n
    #   gũ hợp, chi giờ Dần → "&IF(J236="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J236&") ⇒ "&G236&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.554): Đinh – Nhâm kề nhau ngũ hợp, chi giờ Dần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f414(E, r, c):
    # Ô E237, F237
    # = IF(AND(AND(AND(AND((TRUNG_GIAN!$E$5=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=2,TRUNG_GIAN!$F$6=0)),NOT(OR(AND(TRUE(),TRUNG_GIAN!$C$6=
    #   6),AND(TRUE(),TRUNG_GIAN!$D$6=6),AND(TRUE(),TRUNG_GIAN!$E$6=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=6)))),NOT(OR(AND(TRUE(),TRUNG_GI
    #   AN!$C$5=6),AND(TRUE(),TRUNG_GIAN!$D$5=6),AND(TRUE(),TRUNG_GIAN!$E$5=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=6)))),NOT(OR(AND(TRUE(),
    #   TRUNG_GIAN!$C$5=7),AND(TRUE(),TRUNG_GIAN!$D$5=7),AND(TRUE(),TRUNG_GIAN!$E$5=7),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=7)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.AND(X.eq(E.v(23,5,5),1.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),2.0), X.eq(E.v(23,6,6),0.0))), X.NOT(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),6.0))))), X.NOT(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),6.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))))), X.NOT(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,5,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,5,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),7.0)))))) else '—')

def f415(E, r, c):
    # Ô K237
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.554): Nhật can Ất, trụ giờ 
    #   Bính Tí, không có Ngọ (xung Tí) và không có Canh/Tân → "&IF(J237="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J237&") ⇒ "&G237&
    #   ".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.554): Nhật can Ất, trụ giờ Bính Tí, không có Ngọ (xung Tí) và không có Canh/Tân → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f416(E, r, c):
    # Ô E238, F238
    # = IF(AND((TRUNG_GIAN!$E$5=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=8)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,5,5),6.0), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0)))) else '—')

def f417(E, r, c):
    # Ô K238
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"
    #   Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.5
    #   54): Nhật can Canh, chi giờ Thân (lộc) → "&IF(J238="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J238&") ⇒ "&G238&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.554): Nhật can Canh, chi giờ Thân (lộc) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f418(E, r, c):
    # Ô E239, F239
    # = IF((THAN_SAT!$H$92="Có"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(12,92,8),'Có')) else '—')

def f419(E, r, c):
    # Ô K239
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.554): Có trạch mã (có hiệu lực) → "&IF(J239="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụn
    #   g.","ĐỦ ("&J239&") ⇒ "&G239&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.554): Có trạch mã (có hiệu lực) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f420(E, r, c):
    # Ô E240, F240
    # = IF(AND(AND(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),(INDEX(DUNG_HY_KY!$C$24:$C$28,5)<=1)),(INDEX(DUNG_HY_KY!$C$24:$C$28,1)<=1)),"ĐỦ"
    #   ,"—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.le(X.INDEX(E.rg(8,24,3,28,3), 5.0),1.0)), X.le(X.INDEX(E.rg(8,24,3,28,3), 1.0),1.0))) else '—')

def f421(E, r, c):
    # Ô K240
    # = "Dữ kiện: "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&" → Q
    #   uy tắc (P2 tr.554): Sinh mùa đông (Hợi, Tí, Sửu), Thủy ≤ 1 và Mộc ≤ 1 chữ → "&IF(J240="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ
    #    ("&J240&") ⇒ "&G240&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),' → Quy tắc (P2 tr.554): Sinh mùa đông (Hợi, Tí, Sửu), Thủy ≤ 1 và Mộc ≤ 1 chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f422(E, r, c):
    # Ô E241, F241
    # = IF(OR(AND(TRUE(),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$C$6=3),AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$C$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$C$6=5),AND(TRU
    #   NG_GIAN!$D$6=5,TRUNG_GIAN!$C$6=2),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$C$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$C$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$
    #   C$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$C$6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$C$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$C$6=1),AND(TRUNG_GIAN!$D$
    #   6=10,TRUNG_GIAN!$C$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$C$6=10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$C$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$C$6=1),AN
    #   D(TRUNG_GIAN!$D$6=TRUNG_GIAN!$C$6,OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(TRUE(),OR(AND(TRUNG_GIAN!$D$6=0,
    #   TRUNG_GIAN!$E$6=3),AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$E$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$E$6=2),AND(TRUN
    #   G_GIAN!$D$6=5,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$E$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$E
    #   $6=2),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$
    #   6=7,TRUNG_GIAN!$E$6=10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$E$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$E$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$E$6,OR(TRUNG
    #   _GIAN!$D$6=4,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11)))),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$F$
    #   6=3),AND(TRUNG_GIAN!$D$6=3,TRUNG_GIAN!$F$6=0),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$F$6=2),AND(TRUNG_GIAN!$D$6=5,
    #   TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=5),AND(TRUNG_GIAN!$D$6=2,TRUNG_GIAN!$F$6=8),AND(TRUNG_GIAN!$D$6=8,TRUNG_GIAN!$F$6=2),AND(TRUN
    #   G_GIAN!$D$6=1,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN
    #   !$F$6=10),AND(TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$F$6=7),AND(TRUNG_GIAN!$D$6=7,TRUNG_GIAN!$F$6=1),AND(TRUNG_GIAN!$D$6=TRUNG_GIAN!$F$6,OR(TRUNG_GIAN!$D$6=4,T
    #   RUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=9,TRUNG_GIAN!$D$6=11))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,3),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,3),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,3),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,3),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,3),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,3),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,3),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,3),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,3),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,3),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,3),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,3),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,3),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,3),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,3)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,5),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,5),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,5),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,5),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,5),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,5)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,6),3.0)), X.AND(X.eq(E.v(23,6,4),3.0), X.eq(E.v(23,6,6),0.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),5.0)), X.AND(X.eq(E.v(23,6,4),2.0), X.eq(E.v(23,6,6),8.0)), X.AND(X.eq(E.v(23,6,4),8.0), X.eq(E.v(23,6,6),2.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,6),7.0)), X.AND(X.eq(E.v(23,6,4),7.0), X.eq(E.v(23,6,6),1.0)), X.AND(X.eq(E.v(23,6,4),E.v(23,6,6)), X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),9.0), X.eq(E.v(23,6,4),11.0))))))) else '—')

def f423(E, r, c):
    # Ô K241
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.554): Chi tháng (chỗ vượng)
    #    bị chi khác hình → "&IF(J241="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J241&") ⇒ "&G241&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.554): Chi tháng (chỗ vượng) bị chi khác hình → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f424(E, r, c):
    # Ô K242
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.556): Có thương quan và có ấn → "&IF(J242="","chưa đủ điều kiện ở cả hai phương pháp – không áp 
    #   dụng.","ĐỦ ("&J242&") ⇒ "&G242&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.556): Có thương quan và có ấn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f425(E, r, c):
    # Ô E246, F246
    # = IF(OR(AND(OR(TRUNG_GIAN!$C$6=4,TRUNG_GIAN!$C$6=10,TRUNG_GIAN!$C$6=1,TRUNG_GIAN!$C$6=7),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0))>0)),AND(OR(TRUNG_GIAN!$D$6=4,TRUNG_GIAN!$D$6=10,TRUNG_GIAN!$D$6=1,TRUNG_GIAN!$D$6=7),((COUNTIFS(SAO
    #   !$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1))>0)),AND(OR(TRUNG_GIAN!$E$6=4,TRUNG_GIAN!$E$6=10,TR
    #   UNG_GIAN!$E$6=1,TRUNG_GIAN!$E$6=7),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2))>0
    #   )),AND(OR(TRUNG_GIAN!$F$6=4,TRUNG_GIAN!$F$6=10,TRUNG_GIAN!$F$6=1,TRUNG_GIAN!$F$6=7),((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X
    #   $20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,6,3),4.0), X.eq(E.v(23,6,3),10.0), X.eq(E.v(23,6,3),1.0), X.eq(E.v(23,6,3),7.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),0.0)), X.AND(X.OR(X.eq(E.v(23,6,4),4.0), X.eq(E.v(23,6,4),10.0), X.eq(E.v(23,6,4),1.0), X.eq(E.v(23,6,4),7.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)), X.AND(X.OR(X.eq(E.v(23,6,5),4.0), X.eq(E.v(23,6,5),10.0), X.eq(E.v(23,6,5),1.0), X.eq(E.v(23,6,5),7.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)), X.AND(X.OR(X.eq(E.v(23,6,6),4.0), X.eq(E.v(23,6,6),10.0), X.eq(E.v(23,6,6),1.0), X.eq(E.v(23,6,6),7.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),0.0)))) else '—')

def f426(E, r, c):
    # Ô K246
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có m
    #   ặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)>0,"Thiên tài tàng can ở 
    #   năm: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$
    #   5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)>0,"Thiên tài tàng can ở tháng: "&COUNTIFS(SAO!$F$5:$F$2
    #   0,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)>0,"Thiên tài tàng can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0
    #   ,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$
    #   C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)>0,"Thiên tài tàng can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"<>Can",SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.520): Thiên tài tàng trong chi Thìn, Tuất, Sửu hoặc Mùi (tạp khí) → "&IF(J246="","chưa đủ
    #    điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J246&") ⇒ "&G246&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Thiên tài tàng can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thiên tài tàng can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.520): Thiên tài tàng trong chi Thìn, Tuất, Sửu hoặc Mùi (tạp khí) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f427(E, r, c):
    # Ô E247, F247
    # = IF(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10
    #   ,10,1,4)),AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD
    #   (TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE
    #   (MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(NOT(NHAP_LIEU!$C$9="Không rõ 
    #   giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)))))) else '—')

def f428(E, r, c):
    # Ô K247
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.520): Có chi kho (mộ) của h
    #   ành tài và chi kho của hành quan sát → "&IF(J247="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J247&") ⇒ "&G247&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.520): Có chi kho (mộ) của hành tài và chi kho của hành quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f429(E, r, c):
    # Ô E248, F248
    # = IF(OR(AND(OR(TRUNG_GIAN!$D$12="Chính tài",TRUNG_GIAN!$D$12="Thiên tài"),PP2_DINH_LUONG!$D$43=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(TRUNG_GIAN!$
    #   F$12="Chính tài",TRUNG_GIAN!$F$12="Thiên tài"),PP2_DINH_LUONG!$E$43=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.OR(X.eq(E.v(23,12,4),'Chính tài'), X.eq(E.v(23,12,4),'Thiên tài')), X.eq(E.v(10,43,4),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(E.v(23,12,6),'Chính tài'), X.eq(E.v(23,12,6),'Thiên tài')), X.eq(E.v(10,43,5),1.0)))) else '—')

def f430(E, r, c):
    # Ô K248
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.520): Can tháng hoặc can giờ là tài và ngũ hợp với nhật can → "&IF(J248="","chưa đủ điều kiện ở cả ha
    #   i phương pháp – không áp dụng.","ĐỦ ("&J248&") ⇒ "&G248&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.520): Can tháng hoặc can giờ là tài và ngũ hợp với nhật can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f431(E, r, c):
    # Ô E249
    # = IF(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$98>=4)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,3),1.0), X.ge(E.v(9,98,6),4.0))) else '—')

def f432(E, r, c):
    # Ô F249
    # = IF(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$98>=4)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(8,7,4),1.0), X.ge(E.v(9,98,6),4.0))) else '—')

def f433(E, r, c):
    # Ô K249
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("
    #   &MENH_CUC!$D$8&")")&" → Quy tắc (P2 tr.520): Thân vượng và tài ≥ 4 chữ (thuần tài) → "&IF(J249="","chưa đủ điều kiện ở cả hai phương pháp – không áp d
    #   ụng.","ĐỦ ("&J249&") ⇒ "&G249&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),' → Quy tắc (P2 tr.520): Thân vượng và tài ≥ 4 chữ (thuần tài) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f434(E, r, c):
    # Ô E250, F250
    # = IF((SAO!$Q$10=1),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(13,10,17),1.0)) else '—')

def f435(E, r, c):
    # Ô K250
    # = "Dữ kiện: "&"Lộc ở tháng (chi "&TRUNG_GIAN!$D$11&"): "&IF(SAO!$Q$10=1,"có","không")&" → Quy tắc (P2 tr.520): Chi tháng là lộc của nhật can (kiến lộc) 
    #   → "&IF(J250="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J250&") ⇒ "&G250&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Lộc ở tháng (chi '),E.v(23,11,4)),'): '),('có' if X.truthy(X.eq(E.v(13,10,17),1.0)) else 'không')),' → Quy tắc (P2 tr.520): Chi tháng là lộc của nhật can (kiến lộc) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f436(E, r, c):
    # Ô E251, F251
    # = IF(AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$Q$17=1),NOT(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SAT!$F$92)>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,17),1.0)), X.NOT(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0))))) else '—')

def f437(E, r, c):
    # Ô K251
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.520): Chi giờ là lộc, trụ giờ không có trạch mã → "&IF(J251="","chưa đủ điều kiện ở cả hai phương phá
    #   p – không áp dụng.","ĐỦ ("&J251&") ⇒ "&G251&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.520): Chi giờ là lộc, trụ giờ không có trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f438(E, r, c):
    # Ô E252, F252
    # = IF(AND(AND(AND((PP1_DINH_TINH!$F$100>=1),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)+COUNTIFS(SAO!$F$
    #   5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1))>0))),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$M$5:$M$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1))>0))),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.ge(E.v(9,100,6),1.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),0.0))), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),0.0))), X.ge(E.v(9,98,6),1.0))) else '—')

def f439(E, r, c):
    # Ô K252
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$1
    #   0="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Chính ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)>0,"Thiên ấn bị khắc/hợp mất: "&COUNTIFS(SAO!$F$5
    #   :$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$L$5:$L$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$M$5:$M$20,1)>0,"Chính ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ","")&IF(C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)>0,"Thiên ấn bị xung: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$M$5:$M$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521): Có ấn không bị khắc/xung và có tài → "&IF(J252="","chưa đủ điều k
    #   iện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J252&") ⇒ "&G252&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên ấn bị khắc/hợp mất: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,12,20,12), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bị xung: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,13,20,13), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.521): Có ấn không bị khắc/xung và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f440(E, r, c):
    # Ô E253, F253
    # = IF(OR(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7
    #   ,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(
    #   MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4))),NOT(OR(AND(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUN
    #   G_GIAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GI
    #   AN!$F$6)=6))),AND(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$C$6)=6),AND
    #   (TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6))),AND(AND(TRUE(),TRUNG_
    #   GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_G
    #   IAN!$D$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6))),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F
    #   $6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$D$
    #   6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$6)=6)))))),AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(TR
    #   UE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),AND(NO
    #   T(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4))),NOT(OR(AND(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRU
    #   NG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHA
    #   P_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6)=6))),AND(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),O
    #   R(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRU
    #   NG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6))),AND(AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TR
    #   UNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$D$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6))),
    #   AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIA
    #   N!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$6)=6))))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)))), X.NOT(X.OR(X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0))))))), X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)))), X.NOT(X.OR(X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0))))))))) else '—')

def f441(E, r, c):
    # Ô K253
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.521): Có chi kho của tài ho
    #   ặc quan nhưng không bị xung → "&IF(J253="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J253&") ⇒ "&G253&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.521): Có chi kho của tài hoặc quan nhưng không bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f442(E, r, c):
    # Ô E254, F254
    # = IF(OR(AND(TRUE(),TRUNG_GIAN!$C$52=1,PP2_DINH_LUONG!$C$24=MOD(TRUNG_GIAN!$E$7+2,5)),AND(TRUE(),TRUNG_GIAN!$D$52=1,PP2_DINH_LUONG!$D$24=MOD(TRUNG_GIAN!$
    #   E$7+2,5)),AND(TRUE(),TRUNG_GIAN!$E$52=1,PP2_DINH_LUONG!$E$24=MOD(TRUNG_GIAN!$E$7+2,5)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$52=1,PP2_D
    #   INH_LUONG!$F$24=MOD(TRUNG_GIAN!$E$7+2,5))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,52,3),1.0), X.eq(E.v(10,24,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.AND(X.TRUE(), X.eq(E.v(23,52,4),1.0), X.eq(E.v(10,24,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.AND(X.TRUE(), X.eq(E.v(23,52,5),1.0), X.eq(E.v(10,24,5),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,52,6),1.0), X.eq(E.v(10,24,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))))) else '—')

def f443(E, r, c):
    # Ô K254
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (P2 tr.521): Có tam hợp cục (đủ 3 chi) hóa thành hành tài → "&IF(J254="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J254&"
    #   ) ⇒ "&G254&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.521): Có tam hợp cục (đủ 3 chi) hóa thành hành tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f444(E, r, c):
    # Ô E255, F255
    # = IF(OR(OR(OR(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TR
    #   UNG_GIAN!$F$6=4)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ
    #   "),TRUNG_GIAN!$F$6=10))),AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!$C$9="Khôn
    #   g rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),TRUNG_GIAN!$E$6=7),AND(NOT(NHAP_LIEU!$C$9="Kh
    #   ông rõ giờ"),TRUNG_GIAN!$F$6=7)))),AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=1),AND(TRUE(),TRUNG_GIAN!$D$6=1),AND(TRUE(),TRUNG_GIAN!$E$6=1),AND(NOT(NHAP_LIEU!
    #   $C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=1)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),AND(NOT(NHAP_
    #   LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10)))),AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=10),AND(TRUE(),TRUNG_GIAN!$D$6=10),AND(TRUE(),TRUNG_GIAN!$E$6=10),A
    #   ND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=10)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=7),AND(TRUE(),TRUNG_GIAN!$D$6=7),AND(TRUE(),TRUNG_GIAN!$E$6=7
    #   ),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=7)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.OR(X.OR(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0)))), X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0))))), X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),1.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0))))), X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),10.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),10.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),10.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),7.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),7.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),7.0)))))) else '—')

def f445(E, r, c):
    # Ô K255
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.521): Có cặp Thìn–Tuất, Sửu
    #   –Mùi xung hoặc Sửu–Tuất, Tuất–Mùi hình → "&IF(J255="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J255&") ⇒ "&G255&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.521): Có cặp Thìn–Tuất, Sửu–Mùi xung hoặc Sửu–Tuất, Tuất–Mùi hình → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f446(E, r, c):
    # Ô E256, F256
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f447(E, r, c):
    # Ô K256
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521): Có thươn
    #   g quan và kiếp tài → "&IF(J256="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J256&") ⇒ "&G256&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.521): Có thương quan và kiếp tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f448(E, r, c):
    # Ô E257, F257
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,98,6),E.v(24,99,2)))) else '—')

def f449(E, r, c):
    # Ô K257
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên
    #    quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F
    #   $20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521): Có thực thần và tài ≥ P_NHIEU → "&IF(J257="","chưa đủ điều k
    #   iện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J257&") ⇒ "&G257&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.521): Có thực thần và tài ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f450(E, r, c):
    # Ô K258
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$P$5:$P$
    #   20,1)>0,"Thực thần quý nhân: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$P$5:$P$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)>0,"Thực thần thiên/nguyệt đức: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$O$5:$O$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521): Thực thần đóng cùng quý nhân hoặc thiên/nguyệt đức → "&IF(J258="","
    #   chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J258&") ⇒ "&G258&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần quý nhân: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,16,20,16), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,16,20,16), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần thiên/nguyệt đức: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,15,20,15), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.521): Thực thần đóng cùng quý nhân hoặc thiên/nguyệt đức → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f451(E, r, c):
    # Ô E259, F259
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20
    #   ,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính q
    #   uan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f452(E, r, c):
    # Ô K259
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$
    #   20,1)>0,"Chính tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(
    #   SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thiên tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!
    #   $Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,
    #   "Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$2
    #   0,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>
    #   0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" 
    #   chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí 
    #   vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&CO
    #   UNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.521): T
    #   ài và quan sát đều vượng tại lệnh → "&IF(J259="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J259&") ⇒ "&G259&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thiên tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.521): Tài và quan sát đều vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f453(E, r, c):
    # Ô E260, F260
    # = IF(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10
    #   ,10,1,4)),AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD
    #   (TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4))),NOT(OR(AND(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_G
    #   IAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!
    #   $F$6)=6))),AND(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$C$6)=6),AND(TR
    #   UE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6))),AND(AND(TRUE(),TRUNG_GIA
    #   N!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN
    #   !$D$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6))),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=
    #   CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$D$6)=
    #   6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$6)=6)))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0)))), X.NOT(X.OR(X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0)))))))) else '—')

def f454(E, r, c):
    # Ô K260
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.521): Có chi kho của tài và
    #    chi kho không bị xung → "&IF(J260="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J260&") ⇒ "&G260&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.521): Có chi kho của tài và chi kho không bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f455(E, r, c):
    # Ô E261, F261
    # = IF(OR(AND(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$
    #   X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:
    #   $B$20,0))>0)),AND(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",S
    #   AO!$B$5:$B$20,1))>0)),AND(AND(TRUE(),TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"
    #   <>Can",SAO!$B$5:$B$20,2))>0)),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),((COUNTIFS(SA
    #   O!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0)),0.0)), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),0.0)), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),0.0)), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0)),0.0)))) else '—')

def f456(E, r, c):
    # Ô K261
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&"; "&"Các chữ liên quan (chỉ liệt kê loại có m
    #   ặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)>0,"Chính tài tàng can ở 
    #   năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$
    #   5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)>0,"Thiên tài tàng can ở năm: "&COUNTIFS(SAO!$F$5:$F$20,
    #   "Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y
    #   $20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)>0,"Chính tài tàng can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C
    #   $5:$C$20,"<>Can",SAO!$B$5:$B$20,1)>0,"Thiên tài tàng can ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"<>Can",SAO!$B$5:$B$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$
    #   20,2)>0,"Chính tài tàng can ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)&"
    #    chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)>0,"Thiên tài tàng can ở 
    #   ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)>0,"Chính tài tàng can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20
    #   ,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$
    #   Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)>0,"Thiên tài tàng can ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,S
    #   AO!$X$5:$X$20,1,SAO!$C$5:$C$20,"<>Can",SAO!$B$5:$B$20,3)&" chữ; ",""))&" → Quy tắc (P2 tr.522): Tài tàng trong chi kho (mộ) của hành quan sát → "&IF(J
    #   261="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J261&") ⇒ "&G261&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài tàng can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Thiên tài tàng can ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Chính tài tàng can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên tài tàng can ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 1.0),0.0)) else '')),(X.cat(X.cat('Chính tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Thiên tài tàng can ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 2.0),0.0)) else '')),(X.cat(X.cat('Chính tài tàng can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Thiên tài tàng can ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), '<>Can', E.rg(13,5,2,20,2), 3.0),0.0)) else '')))),' → Quy tắc (P2 tr.522): Tài tàng trong chi kho (mộ) của hành quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f457(E, r, c):
    # Ô E262, F262
    # = IF(AND(AND((PP1_DINH_TINH!$F$99>=1),NOT(OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$2
    #   0,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính 
    #   quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,
    #   1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0)))),OR(((COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO
    #   !$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1))>0),((COUNTIFS(SAO!$F$5:$F$
    #   20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$
    #   5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,99,6),1.0), X.NOT(X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))), X.OR(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f458(E, r, c):
    # Ô K262
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): 
    #   "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính quan thiên can vượng 
    #   tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO
    #   !$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)>0,"Thất sát thiên can vượng tại lệnh: "&COUNTIFS(SAO!$
    #   F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO
    #   !$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính quan bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thất sát bản khí vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can",SAO!$J$5:$J$20,1)>0,"Chính ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C
    #   $20,"Can",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1
    #   )>0,"Thiên ấn thiên can vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can",SAO!$J$5:$J$20,1)&
    #   " chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Chính ấn bản khí vư
    #   ợng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)>0,"Thiên ấn bản khí vượng tại lệnh: "&COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên ấn",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Bản khí",SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.522): Quan 
    #   sát không vượng tại lệnh, ấn vượng tại lệnh → "&IF(J262="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J262&") ⇒ "&G262&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính quan bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thất sát bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn thiên can vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Chính ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')),(X.cat(X.cat('Thiên ấn bản khí vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên ấn', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Bản khí', E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.522): Quan sát không vượng tại lệnh, ấn vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f459(E, r, c):
    # Ô E263, F263
    # = IF(AND(AND(AND((SAO!$O$14=1),(SAO!$O$18=1)),(PP1_DINH_TINH!$F$98>=1)),(PP1_DINH_TINH!$F$99>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(13,14,15),1.0), X.eq(E.v(13,18,15),1.0)), X.ge(E.v(9,98,6),1.0)), X.ge(E.v(9,99,6),1.0))) else '—')

def f460(E, r, c):
    # Ô K263
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC
    #   !$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Thiên/nguyệt đức ở ngày (chi "&TRUNG_GIAN!$E$11&"): "&IF(SAO!$O$14=1,"có","không")&"; "&"Thiên/nguyệt đức ở 
    #   giờ (chi "&TRUNG_GIAN!$F$11&"): "&IF(SAO!$O$18=1,"có","không")&" → Quy tắc (P2 tr.522): Chi ngày và chi giờ đều đóng thiên/nguyệt đức, có tài và quan 
    #   sát → "&IF(J263="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J263&") ⇒ "&G263&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Thiên/nguyệt đức ở ngày (chi '),E.v(23,11,5)),'): '),('có' if X.truthy(X.eq(E.v(13,14,15),1.0)) else 'không')),'; '),'Thiên/nguyệt đức ở giờ (chi '),E.v(23,11,6)),'): '),('có' if X.truthy(X.eq(E.v(13,18,15),1.0)) else 'không')),' → Quy tắc (P2 tr.522): Chi ngày và chi giờ đều đóng thiên/nguyệt đức, có tài và quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f461(E, r, c):
    # Ô E264, F264
    # = IF(OR(AND(AND(TRUE(),LEN(THAN_SAT!$C$92)>0),TRUNG_GIAN!$C$9=MOD(TRUNG_GIAN!$E$7+2,5)),AND(AND(TRUE(),LEN(THAN_SAT!$D$92)>0),TRUNG_GIAN!$D$9=MOD(TRUNG_
    #   GIAN!$E$7+2,5)),AND(AND(TRUE(),LEN(THAN_SAT!$E$92)>0),TRUNG_GIAN!$E$9=MOD(TRUNG_GIAN!$E$7+2,5)),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),LEN(THAN_SA
    #   T!$F$92)>0),TRUNG_GIAN!$F$9=MOD(TRUNG_GIAN!$E$7+2,5))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,3)),0.0)), X.eq(E.v(23,9,3),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,4)),0.0)), X.eq(E.v(23,9,4),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.AND(X.AND(X.TRUE(), X.gt(X.LEN(E.v(12,92,5)),0.0)), X.eq(E.v(23,9,5),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.gt(X.LEN(E.v(12,92,6)),0.0)), X.eq(E.v(23,9,6),X.MOD(X.add(E.v(23,7,5),2.0), 5.0))))) else '—')

def f462(E, r, c):
    # Ô K264
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.522): Chi có trạch mã mang 
    #   hành tài (nhật can khắc) → "&IF(J264="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J264&") ⇒ "&G264&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.522): Chi có trạch mã mang hành tài (nhật can khắc) → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f463(E, r, c):
    # Ô E265
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=MOD(TRUNG_GIAN!$E$7+0,5))),"
    #   ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)))) else '—')

def f464(E, r, c):
    # Ô F265
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$98>=P_NHIEU)),(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)=MOD(TRUNG_GIAN!$E$7+0,5))),"
    #   ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,98,6),E.v(24,99,2))), X.eq((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),X.MOD(X.add(E.v(23,7,5),0.0), 5.0)))) else '—')

def f465(E, r, c):
    # Ô K265
    # = "Dữ kiện: "&"Trạng thái nhật chủ PP1: "&DUNG_HY_KY!$C$6&" / PP2: "&DUNG_HY_KY!$D$6&"; "&"Dụng PP1: "&IF(DUNG_HY_KY!$C$109>=0,INDEX(HANH_TEN,DUNG_HY_KY
    #   !$C$109+1),"—")&" / PP2: "&IF(DUNG_HY_KY!$D$109>=0,INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1),"—")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="
    #   ",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Quy tắc (P2 tr.522): Thân nhược, tài ≥ P_NHIEU, dụng thần là tỉ kiếp → "&IF(J265="","chưa đủ
    #    điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J265&") ⇒ "&G265&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trạng thái nhật chủ PP1: '),E.v(8,6,3)),' / PP2: '),E.v(8,6,4)),'; '),'Dụng PP1: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)) if X.truthy(X.ge(E.v(8,109,3),0.0)) else '—')),' / PP2: '),(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0)) if X.truthy(X.ge(E.v(8,109,4),0.0)) else '—')),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (P2 tr.522): Thân nhược, tài ≥ P_NHIEU, dụng thần là tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f466(E, r, c):
    # Ô E266, F266
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$2
    #   0,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")+COUNTIFS
    #   (SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0), X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),0.0))) else '—')

def f467(E, r, c):
    # Ô K266
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can")>0,"Chính tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thiên tài thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SA
    #   O!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:
    #   $C$20,"Can")>0,"Chính quan thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(C
    #   OUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",
    #   SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (P2 tr.522): Tài và quan sát đều lộ ở can → "&IF(J266="","chưa đủ đi
    #   ều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J266&") ⇒ "&G266&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thiên tài thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')),(X.cat(X.cat('Chính quan thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')),(X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (P2 tr.522): Tài và quan sát đều lộ ở can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f468(E, r, c):
    # Ô E267, F267
    # = IF(AND(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$99>=1)),((COUNTIFS(SAO!$F$5:$F$20,"Thực thần
    #   ",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,99,6),1.0)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f469(E, r, c):
    # Ô K267
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): 
    #   "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thiên tài: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,
    #   SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực
    #    thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.522): Có thiên tài, quan sát và thực thần → "&IF(J267="","chưa đủ điều kiện
    #    ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J267&") ⇒ "&G267&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thiên tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.522): Có thiên tài, quan sát và thực thần → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f470(E, r, c):
    # Ô E268, F268
    # = IF(AND(AND(OR(AND(TRUE(),SAO!$P$5=1),AND(TRUE(),SAO!$P$9=1),AND(TRUE(),SAO!$P$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$P$17=1)),OR(AND(TRUE()
    #   ,SAO!$Q$5=1),AND(TRUE(),SAO!$Q$9=1),AND(TRUE(),SAO!$Q$13=1),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),SAO!$Q$17=1))),(THAN_SAT!$H$92="Có")),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,9,16),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,16),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,16),1.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(13,5,17),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,9,17),1.0)), X.AND(X.TRUE(), X.eq(E.v(13,13,17),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(13,17,17),1.0)))), X.eq(E.v(12,92,8),'Có'))) else '—')

def f471(E, r, c):
    # Ô K268
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.522): Có quý nhân, có lộc và có trạch mã → "&IF(J268="","chưa đủ điều kiện ở cả hai phương pháp – khô
    #   ng áp dụng.","ĐỦ ("&J268&") ⇒ "&G268&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.522): Có quý nhân, có lộc và có trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f472(E, r, c):
    # Ô E269, F269
    # = IF(AND((TRUNG_GIAN!$E$7=4),AND(AND(OR(AND(TRUE(),TRUNG_GIAN!$C$6=8),AND(TRUE(),TRUNG_GIAN!$D$6=8),AND(TRUE(),TRUNG_GIAN!$E$6=8),AND(NOT(NHAP_LIEU!$C$9
    #   ="Không rõ giờ"),TRUNG_GIAN!$F$6=8)),OR(AND(TRUE(),TRUNG_GIAN!$C$6=0),AND(TRUE(),TRUNG_GIAN!$D$6=0),AND(TRUE(),TRUNG_GIAN!$E$6=0),AND(NOT(NHAP_LIEU!$C
    #   $9="Không rõ giờ"),TRUNG_GIAN!$F$6=0))),OR(AND(TRUE(),TRUNG_GIAN!$C$6=4),AND(TRUE(),TRUNG_GIAN!$D$6=4),AND(TRUE(),TRUNG_GIAN!$E$6=4),AND(NOT(NHAP_LIEU
    #   !$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=4)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.eq(E.v(23,7,5),4.0), X.AND(X.AND(X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),8.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),8.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),8.0))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),0.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),0.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),0.0)))), X.OR(X.AND(X.TRUE(), X.eq(E.v(23,6,3),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,4),4.0)), X.AND(X.TRUE(), X.eq(E.v(23,6,5),4.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),4.0)))))) else '—')

def f473(E, r, c):
    # Ô K269
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.522): Nhật can Nhâm/Quý, đủ
    #    Thân, Tí, Thìn → "&IF(J269="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J269&") ⇒ "&G269&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.522): Nhật can Nhâm/Quý, đủ Thân, Tí, Thìn → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f474(E, r, c):
    # Ô E270, F270
    # = IF(AND((PP1_DINH_TINH!$F$98>=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,98,6),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0))) else '—')

def f475(E, r, c):
    # Ô K270
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRI
    #   M(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.523): Có tài và có thất sát → "&IF(J270="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("
    #   &J270&") ⇒ "&G270&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.523): Có tài và có thất sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f476(E, r, c):
    # Ô E271, F271
    # = IF(AND(OR(TRUNG_GIAN!$C$7=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!$C$7+1,5)=TRUNG_GIAN!$E$7),OR(TRUNG_GIAN!$D$7=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!$D$7+1,5)=TRUNG_
    #   GIAN!$E$7),OR(NOT(NOT(NHAP_LIEU!$C$9="Không rõ giờ")),TRUNG_GIAN!$F$7=TRUNG_GIAN!$E$7,MOD(TRUNG_GIAN!$F$7+1,5)=TRUNG_GIAN!$E$7)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.OR(X.eq(E.v(23,7,3),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,3),1.0), 5.0),E.v(23,7,5))), X.OR(X.eq(E.v(23,7,4),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,4),1.0), 5.0),E.v(23,7,5))), X.OR(X.NOT(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))), X.eq(E.v(23,7,6),E.v(23,7,5)), X.eq(X.MOD(X.add(E.v(23,7,6),1.0), 5.0),E.v(23,7,5))))) else '—')

def f477(E, r, c):
    # Ô K271
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.523): Hành can năm, tháng, 
    #   giờ đều cùng hành hoặc sinh hành can ngày → "&IF(J271="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J271&") ⇒ "&G271&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.523): Hành can năm, tháng, giờ đều cùng hành hoặc sinh hành can ngày → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f478(E, r, c):
    # Ô E272, F272
    # = IF(AND(AND((PP1_DINH_TINH!$F$100>=1),(PP1_DINH_TINH!$F$98>=1)),(PP1_DINH_TINH!$F$96>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,100,6),1.0), X.ge(E.v(9,98,6),1.0)), X.ge(E.v(9,96,6),1.0))) else '—')

def f479(E, r, c):
    # Ô K272
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.523): C
    #   ó ấn, có tài và có tỉ kiếp → "&IF(J272="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J272&") ⇒ "&G272&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.523): Có ấn, có tài và có tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f480(E, r, c):
    # Ô E273, F273
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '—')

def f481(E, r, c):
    # Ô K273
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$2
    #   0,1)>0,"Kiếp tài vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P
    #   2 tr.527): Kiếp tài vượng tại lệnh → "&IF(J273="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J273&") ⇒ "&G273&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''))),' → Quy tắc (P2 tr.527): Kiếp tài vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f482(E, r, c):
    # Ô E274, F274
    # = IF(AND(AND((PP1_DINH_TINH!$F$100>=1),(PP1_DINH_TINH!$F$98>=1)),(PP1_DINH_TINH!$F$96=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,100,6),1.0), X.ge(E.v(9,98,6),1.0)), X.eq(E.v(9,96,6),0.0))) else '—')

def f483(E, r, c):
    # Ô K274
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&" → Quy tắc (P2 tr.527): C
    #   ó ấn và tài, không có tỉ kiếp → "&IF(J274="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J274&") ⇒ "&G274&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),' → Quy tắc (P2 tr.527): Có ấn và tài, không có tỉ kiếp → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f484(E, r, c):
    # Ô E275, F275
    # = IF((COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)>=P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2))) else '—')

def f485(E, r, c):
    # Ô K275
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:
    #   $Y$20,0)>0,"Kiếp tài: "&COUNTIFS(SAO!$F$5:$F$20,"Kiếp tài",SAO!$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Kiếp tài ≥ P_NHIEU chữ → "&IF(J275
    #   ="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J275&") ⇒ "&G275&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Kiếp tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Kiếp tài', E.rg(13,5,25,20,25), 0.0),0.0)) else ''))),' → Quy tắc (P2 tr.527): Kiếp tài ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f486(E, r, c):
    # Ô E276, F276
    # = IF(AND((PP1_DINH_TINH!$F$96>=P_NHIEU),(PP1_DINH_TINH!$F$98>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,96,6),E.v(24,99,2)), X.ge(E.v(9,98,6),1.0))) else '—')

def f487(E, r, c):
    # Ô K276
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!
    #   $D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.527): Tỉ kiếp ≥ P_NHIEU và có tài → "&IF(J276="","chưa đủ 
    #   điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J276&") ⇒ "&G276&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.527): Tỉ kiếp ≥ P_NHIEU và có tài → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f488(E, r, c):
    # Ô E277, F277
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)+COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1,SAO!$K$5:$K$20,8))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),0.0)) else '—')

def f489(E, r, c):
    # Ô K277
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$
    #   20,8)>0,"Chính tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5
    #   :$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$K$5:$K$20,8)>0,"Thiên tài lâm Mộ: "&COUNTIFS(SAO!$F$5:$F$20,"Thiên tài",SAO!$Y$5:$Y$20,0,SAO
    #   !$X$5:$X$20,1,SAO!$K$5:$K$20,8)&" chữ; ",""))&" → Quy tắc (P2 tr.527): Có chữ tài lâm Mộ → "&IF(J277="","chưa đủ điều kiện ở cả hai phương pháp – khôn
    #   g áp dụng.","ĐỦ ("&J277&") ⇒ "&G277&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else ''),(X.cat(X.cat('Thiên tài lâm Mộ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thiên tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,11,20,11), 8.0),0.0)) else '')))),' → Quy tắc (P2 tr.527): Có chữ tài lâm Mộ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f490(E, r, c):
    # Ô E278, F278
    # = IF((THAN_SAT!$H$124="Có"),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(12,124,8),'Có')) else '—')

def f491(E, r, c):
    # Ô K278
    # = "Dữ kiện: "&"(xem điều kiện)"&" → Quy tắc (P2 tr.528): Ngày thập ác đại bại → "&IF(J278="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","
    #   ĐỦ ("&J278&") ⇒ "&G278&".")
    return X.cat(X.cat(X.cat('Dữ kiện: ','(xem điều kiện)'),' → Quy tắc (P2 tr.528): Ngày thập ác đại bại → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f492(E, r, c):
    # Ô E279, F279
    # = IF(AND(NOT((NHAP_LIEU!$C$22="Nam")),OR(OR(AND(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!
    #   $C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6
    #   )=6))),AND(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$C$6)=6),AND(TRUE()
    #   ,ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6))),AND(AND(TRUE(),TRUNG_GIAN!$E
    #   $6=CHOOSE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$D$
    #   6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6))),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=CHOO
    #   SE(MOD(TRUNG_GIAN!$E$7+2,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$D$6)=6),A
    #   ND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$6)=6)))),OR(AND(AND(TRUE(),TRUNG_GIAN!$C$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),A
    #   BS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$C$6-T
    #   RUNG_GIAN!$F$6)=6))),AND(AND(TRUE(),TRUNG_GIAN!$D$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$C$6)
    #   =6),AND(TRUE(),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6)=6))),AND(AND(TRUE()
    #   ,TRUNG_GIAN!$E$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$E$6-
    #   TRUNG_GIAN!$D$6)=6),AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ABS(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6)=6))),AND(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_
    #   GIAN!$F$6=CHOOSE(MOD(TRUNG_GIAN!$E$7+3,5)+1,7,10,10,1,4)),OR(AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$C$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_G
    #   IAN!$D$6)=6),AND(TRUE(),ABS(TRUNG_GIAN!$F$6-TRUNG_GIAN!$E$6)=6)))))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.NOT(X.eq(E.v(2,22,3),'Nam')), X.OR(X.OR(X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0))))), X.OR(X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,3),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,3),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,4),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,5))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,4),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.TRUE(), X.eq(E.v(23,6,5),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,4))),6.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.ABS(X.sub(E.v(23,6,5),E.v(23,6,6))),6.0)))), X.AND(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.CHOOSE(X.add(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),1.0), lambda: 7.0, lambda: 10.0, lambda: 10.0, lambda: 1.0, lambda: 4.0))), X.OR(X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,3))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,4))),6.0)), X.AND(X.TRUE(), X.eq(X.ABS(X.sub(E.v(23,6,6),E.v(23,6,5))),6.0)))))))) else '—')

def f493(E, r, c):
    # Ô K279
    # = "Dữ kiện: "&"Giới tính: "&NHAP_LIEU!$C$22&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH
    #   _TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_G
    #   IAN!$D$9+1)&")"&"; "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&"
    #   )"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc
    #    (P2 tr.522): Nữ, có chi kho của tài/quan và chi kho đó bị xung → "&IF(J279="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J279&"
    #   ) ⇒ "&G279&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Giới tính: '),E.v(2,22,3)),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.522): Nữ, có chi kho của tài/quan và chi kho đó bị xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f494(E, r, c):
    # Ô E284, F284
    # = IF(AND(AND(TRUE(),OR(AND(TRUNG_GIAN!$E$5=0,TRUNG_GIAN!$E$6=2),AND(TRUNG_GIAN!$E$5=4,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=2,TRUNG_GIAN!$E$6=4),AND(TR
    #   UNG_GIAN!$E$5=6,TRUNG_GIAN!$E$6=4),AND(TRUNG_GIAN!$E$5=8,TRUNG_GIAN!$E$6=10))),OR(AND(TRUE(),OR(AND(TRUNG_GIAN!$C$5=8,TRUNG_GIAN!$C$6=4),AND(TRUNG_GIA
    #   N!$C$5=6,TRUNG_GIAN!$C$6=10),AND(TRUNG_GIAN!$C$5=6,TRUNG_GIAN!$C$6=4),AND(TRUNG_GIAN!$C$5=4,TRUNG_GIAN!$C$6=10))),AND(TRUE(),OR(AND(TRUNG_GIAN!$D$5=8,
    #   TRUNG_GIAN!$D$6=4),AND(TRUNG_GIAN!$D$5=6,TRUNG_GIAN!$D$6=10),AND(TRUNG_GIAN!$D$5=6,TRUNG_GIAN!$D$6=4),AND(TRUNG_GIAN!$D$5=4,TRUNG_GIAN!$D$6=10))),AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(AND(TRUNG_GIAN!$F$5=8,TRUNG_GIAN!$F$6=4),AND(TRUNG_GIAN!$F$5=6,TRUNG_GIAN!$F$6=10),AND(TRUNG_GIAN!$F$5=6,TRUNG_G
    #   IAN!$F$6=4),AND(TRUNG_GIAN!$F$5=4,TRUNG_GIAN!$F$6=10)))),OR(AND(TRUE(),TRUE(),OR(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1,INDEX(M_HINH,TRUNG_
    #   GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1)),AND(TRUE(),TRUE(),OR(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1,INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D
    #   $6+1)=1)),AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),OR(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)=1,INDEX(M_HINH,TRUNG_GIAN!$E$6+1,TRUNG_GIAN
    #   !$F$6+1)=1)))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,5),0.0), X.eq(E.v(23,6,5),2.0)), X.AND(X.eq(E.v(23,5,5),4.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),2.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),6.0), X.eq(E.v(23,6,5),4.0)), X.AND(X.eq(E.v(23,5,5),8.0), X.eq(E.v(23,6,5),10.0)))), X.OR(X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,3),8.0), X.eq(E.v(23,6,3),4.0)), X.AND(X.eq(E.v(23,5,3),6.0), X.eq(E.v(23,6,3),10.0)), X.AND(X.eq(E.v(23,5,3),6.0), X.eq(E.v(23,6,3),4.0)), X.AND(X.eq(E.v(23,5,3),4.0), X.eq(E.v(23,6,3),10.0)))), X.AND(X.TRUE(), X.OR(X.AND(X.eq(E.v(23,5,4),8.0), X.eq(E.v(23,6,4),4.0)), X.AND(X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,6,4),10.0)), X.AND(X.eq(E.v(23,5,4),6.0), X.eq(E.v(23,6,4),4.0)), X.AND(X.eq(E.v(23,5,4),4.0), X.eq(E.v(23,6,4),10.0)))), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.AND(X.eq(E.v(23,5,6),8.0), X.eq(E.v(23,6,6),4.0)), X.AND(X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,6,6),10.0)), X.AND(X.eq(E.v(23,5,6),6.0), X.eq(E.v(23,6,6),4.0)), X.AND(X.eq(E.v(23,5,6),4.0), X.eq(E.v(23,6,6),10.0))))), X.OR(X.AND(X.TRUE(), X.TRUE(), X.OR(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0))), X.AND(X.TRUE(), X.TRUE(), X.OR(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0))), X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.OR(X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0), X.eq(X.INDEX(E.rg(24,92,16,103,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)))))) else '—')

def f495(E, r, c):
    # Ô K284
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (P2 tr.548; UH tr.102-103, 194): N
    #   gày nhật đức + trụ khác là khôi cương (Nhâm Thìn, Canh Tuất, Canh Thìn, Mậu Tuất) + chi ngày bị hình/xung → "&IF(J284="","chưa đủ điều kiện ở cả hai p
    #   hương pháp – không áp dụng.","ĐỦ ("&J284&") ⇒ "&G284&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (P2 tr.548; UH tr.102-103, 194): Ngày nhật đức + trụ khác là khôi cương (Nhâm Thìn, Canh Tuất, Canh Thìn, Mậu Tuất) + chi ngày bị hình/xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f496(E, r, c):
    # Ô E285, F285
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!
    #   $X$5:$X$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)))) else '—')

def f497(E, r, c):
    # Ô K285
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.551, 536
    #   ; UH tr.96, 239): Có thương quan và tứ trụ không có một chữ chính quan nào (lộ hay tàng) – ""thương tận"" theo UH tr.239 → "&IF(J285="","chưa đủ điều 
    #   kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J285&") ⇒ "&G285&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.551, 536; UH tr.96, 239): Có thương quan và tứ trụ không có một chữ chính quan nào (lộ hay tàng) – "thương tận" theo UH tr.239 → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f498(E, r, c):
    # Ô K286
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thương
    #    quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.541; UH 
    #   tr.239): Có thương quan và có chính quan (chưa ""thương tận"") → "&IF(J286="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J286&")
    #    ⇒ "&G286&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.541; UH tr.239): Có thương quan và có chính quan (chưa "thương tận") → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f499(E, r, c):
    # Ô E287
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100>=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,100,6),1.0)))) else '—')

def f500(E, r, c):
    # Ô F287
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),OR(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",S
    #   AO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$100>=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.ge(E.v(9,100,6),1.0)))) else '—')

def f501(E, r, c):
    # Ô K287
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thực thần: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.552; UH tr.71-72, 76, 88): Thân vượng, có thất sát, có thực thần chế hoặc có ấn h
    #   óa (""thân cường sát nhược, giả sát vi quyền"") → "&IF(J287="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J287&") ⇒ "&G287&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thực thần: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.552; UH tr.71-72, 76, 88): Thân vượng, có thất sát, có thực thần chế hoặc có ấn hóa ("thân cường sát nhược, giả sát vi quyền") → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f502(E, r, c):
    # Ô E288
    # = IF(AND(AND(AND((DUNG_HY_KY!$C$7=0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thất sá
    #   t",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))), X.ge(E.v(9,100,6),1.0))) else '—')

def f503(E, r, c):
    # Ô F288
    # = IF(AND(AND(AND((DUNG_HY_KY!$D$7=0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Thất sá
    #   t",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),(PP1_DINH_TINH!$F$100>=1)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0))), X.ge(E.v(9,100,6),1.0))) else '—')

def f504(E, r, c):
    # Ô K288
    # = "Dữ kiện: "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&T
    #   RIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Thất sát: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X
    #   $5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thất sát vượng tại lệnh: "&COUNT
    #   IFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (P2 tr.533; UH tr.71-72): Thân nhược, có t
    #   hất sát nhưng không vượng tại lệnh, có ấn hóa sát → "&IF(J288="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J288&") ⇒ "&G288&"."
    #   )
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thất sát: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (P2 tr.533; UH tr.71-72): Thân nhược, có thất sát nhưng không vượng tại lệnh, có ấn hóa sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f505(E, r, c):
    # Ô E289
    # = IF(AND(AND((DUNG_HY_KY!$C$7=0),(PP1_DINH_TINH!$F$99>=P_NHIEU)),(PP1_DINH_TINH!$F$97=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),0.0), X.ge(E.v(9,99,6),E.v(24,99,2))), X.eq(E.v(9,97,6),0.0))) else '—')

def f506(E, r, c):
    # Ô F289
    # = IF(AND(AND((DUNG_HY_KY!$D$7=0),(PP1_DINH_TINH!$F$99>=P_NHIEU)),(PP1_DINH_TINH!$F$97=0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),0.0), X.ge(E.v(9,99,6),E.v(24,99,2))), X.eq(E.v(9,97,6),0.0))) else '—')

def f507(E, r, c):
    # Ô K289
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(
    #   MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (P2 tr.542; UH tr.62, 72-73, 86): Thân nhược, quan sát ≥
    #    P_NHIEU chữ và không có thực thương chế → "&IF(J289="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J289&") ⇒ "&G289&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (P2 tr.542; UH tr.62, 72-73, 86): Thân nhược, quan sát ≥ P_NHIEU chữ và không có thực thương chế → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f508(E, r, c):
    # Ô E294, F294
    # = IF(OR(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0),(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:
    #   $Y$20,0)>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0), X.ge(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),E.v(24,99,2)))) else '—')

def f509(E, r, c):
    # Ô K294
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thương quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1
    #   ,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0)>0,"Thương quan: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO
    #   !$Y$5:$Y$20,0)&" chữ; ",""))&" → Quy tắc (HG tr.155): Thương quan vượng tại lệnh hoặc thương quan ≥ P_NHIEU chữ → "&IF(J294="","chưa đủ điều kiện ở cả
    #    hai phương pháp – không áp dụng.","ĐỦ ("&J294&") ⇒ "&G294&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thương quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0),0.0)) else '')))),' → Quy tắc (HG tr.155): Thương quan vượng tại lệnh hoặc thương quan ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f510(E, r, c):
    # Ô E295, F295
    # = IF(AND((PP1_DINH_TINH!$F$96>=P_NHIEU),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)+COUNTIFS(SAO!$F$5
    #   :$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,96,6),E.v(24,99,2)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f511(E, r, c):
    # Ô K295
    # = "Dữ kiện: "&"Số chữ Tỉ kiếp: "&PP1_DINH_TINH!$F$96&IF(MENH_CUC!$D$6="",""," ("&MENH_CUC!$D$6&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ 
    #   liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính quan 
    #   vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Th
    #   ất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Thất sát vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5
    #   :$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Quy tắc (HG tr.155): Tỉ kiếp ≥ P_NHIEU, quan sát không có hoặc không vượng tại lệnh → "&IF(J295="","chưa 
    #   đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J295&") ⇒ "&G295&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tỉ kiếp: '),E.v(9,96,6)),('' if X.truthy(X.eq(E.v(18,6,4),'')) else X.cat(X.cat(' (',E.v(18,6,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (HG tr.155): Tỉ kiếp ≥ P_NHIEU, quan sát không có hoặc không vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f512(E, r, c):
    # Ô K296
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (HG tr.155): Tứ trụ không có quan sát →
    #    "&IF(J296="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J296&") ⇒ "&G296&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (HG tr.155): Tứ trụ không có quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f513(E, r, c):
    # Ô E297, F297
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0),(PP1_DINH_TINH!$F$99>PP1_DINH_TINH!$F$97)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0), X.gt(E.v(9,99,6),E.v(9,97,6)))) else '—')

def f514(E, r, c):
    # Ô K297
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(
    #   MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:
    #   $Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính tài: "&COUNTIFS(SAO!$F$5:$F$20,"Chính tài",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (HG tr.155)
    #   : Có chính tài, quan sát nhiều chữ hơn thực thương → "&IF(J297="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J297&") ⇒ "&G297&".
    #   ")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính tài: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính tài', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (HG tr.155): Có chính tài, quan sát nhiều chữ hơn thực thương → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f515(E, r, c):
    # Ô E298, F298
    # = IF((PP1_DINH_TINH!$F$97>PP1_DINH_TINH!$F$99),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(E.v(9,97,6),E.v(9,99,6))) else '—')

def f516(E, r, c):
    # Ô K298
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(
    #   MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&" → Quy tắc (HG tr.155): Thực thương nhiều chữ hơn quan sát → "&IF(J298="","chưa đủ điều kiện ở cả hai phư
    #   ơng pháp – không áp dụng.","ĐỦ ("&J298&") ⇒ "&G298&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),' → Quy tắc (HG tr.155): Thực thương nhiều chữ hơn quan sát → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f517(E, r, c):
    # Ô E299, F299
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5
    #   :$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0))) else '—')

def f518(E, r, c):
    # Ô K299
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$
    #   20,"Can")>0,"Thực thần thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thực thần",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(COUNTI
    #   FS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$
    #   Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (HG tr.156): Thực thần và thất sát cùng lộ ở thiên can → "&IF(J299="","ch
    #   ưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J299&") ⇒ "&G299&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thực thần thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thực thần', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (HG tr.156): Thực thần và thất sát cùng lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f519(E, r, c):
    # Ô E300, F300
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0),((COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y
    #   $5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can"))>0)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0))) else '—')

def f520(E, r, c):
    # Ô K300
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$
    #   C$20,"Can")>0,"Thương quan thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thương quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ","")&IF(
    #   COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")>0,"Thất sát thiên can: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát"
    #   ,SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$C$5:$C$20,"Can")&" chữ; ",""))&" → Quy tắc (HG tr.156): Thương quan và thất sát cùng lộ ở thiên can → "&IF(J30
    #   0="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J300&") ⇒ "&G300&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Thương quan thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thương quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else ''),(X.cat(X.cat('Thất sát thiên can: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can')),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,3,20,3), 'Can'),0.0)) else '')))),' → Quy tắc (HG tr.156): Thương quan và thất sát cùng lộ ở thiên can → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f521(E, r, c):
    # Ô E301, F301
    # = IF(AND((PP1_DINH_TINH!$F$100>=P_NHIEU),(PP1_DINH_TINH!$F$97>=P_NHIEU)),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.ge(E.v(9,100,6),E.v(24,99,2)), X.ge(E.v(9,97,6),E.v(24,99,2)))) else '—')

def f522(E, r, c):
    # Ô K301
    # = "Dữ kiện: "&"Số chữ Thực thương: "&PP1_DINH_TINH!$F$97&IF(MENH_CUC!$D$7="",""," ("&MENH_CUC!$D$7&")")&"; "&"Số chữ Ấn: "&PP1_DINH_TINH!$F$100&IF(MENH_
    #   CUC!$D$10="",""," ("&MENH_CUC!$D$10&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&" → Quy tắc (HG tr.156): Ấn ≥ P_NHIEU và thực thương ≥ P_NHIEU → "&IF(J
    #   301="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J301&") ⇒ "&G301&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Thực thương: '),E.v(9,97,6)),('' if X.truthy(X.eq(E.v(18,7,4),'')) else X.cat(X.cat(' (',E.v(18,7,4)),')'))),'; '),'Số chữ Ấn: '),E.v(9,100,6)),('' if X.truthy(X.eq(E.v(18,10,4),'')) else X.cat(X.cat(' (',E.v(18,10,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),' → Quy tắc (HG tr.156): Ấn ≥ P_NHIEU và thực thương ≥ P_NHIEU → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f523(E, r, c):
    # Ô E302, F302
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SA
    #   O!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0)) else '—')

def f524(E, r, c):
    # Ô K302
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,0)>0,"Chính quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F
    #   $5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Chính quan ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20
    #   ,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (HG tr.165): Chính quan ở trụ năm hoặc trụ tháng → "&IF(J302="","chưa đủ điều kiện ở c
    #   ả hai phương pháp – không áp dụng.","ĐỦ ("&J302&") ⇒ "&G302&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else ''),(X.cat(X.cat('Chính quan ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (HG tr.165): Chính quan ở trụ năm hoặc trụ tháng → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f525(E, r, c):
    # Ô E303, F303
    # = IF(AND(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)+COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3))>0),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)+COUNTIFS(SAO!$
    #   F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),0.0), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),0.0)))) else '—')

def f526(E, r, c):
    # Ô K303
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B
    #   $20,2)>0,"Chính quan ở ngày: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,2)&" chữ; ","")&IF(COUNTIFS(SAO!$
    #   F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)>0,"Chính quan ở giờ: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,
    #   0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,3)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)>0,"Chính
    #    quan ở năm: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,0)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính
    #    quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$B$5:$B$20,1)>0,"Chính quan ở tháng: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$B$5:$B$20,1)&" chữ; ",""))&" → Quy tắc (HG tr.165): Chính quan chỉ ở trụ ngày/giờ → "&IF(J303="","chưa đủ điều kiện ở cả hai phương pháp – k
    #   hông áp dụng.","ĐỦ ("&J303&") ⇒ "&G303&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('Chính quan ở ngày: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 2.0),0.0)) else ''),(X.cat(X.cat('Chính quan ở giờ: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 3.0),0.0)) else '')),(X.cat(X.cat('Chính quan ở năm: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 0.0),0.0)) else '')),(X.cat(X.cat('Chính quan ở tháng: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,2,20,2), 1.0),0.0)) else '')))),' → Quy tắc (HG tr.165): Chính quan chỉ ở trụ ngày/giờ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f527(E, r, c):
    # Ô E304, F304
    # = IF(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1))>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else '—')

def f528(E, r, c):
    # Ô K304
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N
    #   $20,1)>0,"Chính quan tuần không: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$N$5:$N$20,1)&" chữ; ",""))&" → Quy tắc 
    #   (HG tr.165): Chính quan rơi vào tuần không → "&IF(J304="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J304&") ⇒ "&G304&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('Chính quan tuần không: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,14,20,14), 1.0),0.0)) else ''))),' → Quy tắc (HG tr.165): Chính quan rơi vào tuần không → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f529(E, r, c):
    # Ô E305, F305
    # = IF(AND(AND((PP1_DINH_TINH!$F$98>=P_NHIEU),((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1))>0)),NOT(((COUNTIFS(SAO!$F$5:$F$20
    #   ,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.ge(E.v(9,98,6),E.v(24,99,2)), X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)), X.NOT(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)))) else '—')

def f530(E, r, c):
    # Ô K305
    # = "Dữ kiện: "&"Số chữ Tài: "&PP1_DINH_TINH!$F$98&IF(MENH_CUC!$D$8="",""," ("&MENH_CUC!$D$8&")")&"; "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Các chữ liên
    #    quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)>0,"Chính quan: "&COUNTIFS(SAO!$F$5:
    #   $F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:
    #   $J$20,1)>0,"Chính quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ",""))&" → Qu
    #   y tắc (HG tr.165): Tài ≥ P_NHIEU, có chính quan nhưng chính quan không vượng tại lệnh → "&IF(J305="","chưa đủ điều kiện ở cả hai phương pháp – không á
    #   p dụng.","ĐỦ ("&J305&") ⇒ "&G305&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Tài: '),E.v(9,98,6)),('' if X.truthy(X.eq(E.v(18,8,4),'')) else X.cat(X.cat(' (',E.v(18,8,4)),')'))),'; '),'ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''),(X.cat(X.cat('Chính quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (HG tr.165): Tài ≥ P_NHIEU, có chính quan nhưng chính quan không vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f531(E, r, c):
    # Ô E306
    # = IF(AND(AND((DUNG_HY_KY!$C$7=1),(PP1_DINH_TINH!$F$99=1)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)
    #   +COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,3),1.0), X.eq(E.v(9,99,6),1.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f532(E, r, c):
    # Ô F306
    # = IF(AND(AND((DUNG_HY_KY!$D$7=1),(PP1_DINH_TINH!$F$99=1)),NOT(((COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)
    #   +COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1))>0))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.AND(X.AND(X.eq(E.v(8,7,4),1.0), X.eq(E.v(9,99,6),1.0)), X.NOT(X.gt(X.add(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),0.0)))) else '—')

def f533(E, r, c):
    # Ô K306
    # = "Dữ kiện: "&"Số chữ Quan sát: "&PP1_DINH_TINH!$F$99&IF(MENH_CUC!$D$9="",""," ("&MENH_CUC!$D$9&")")&"; "&"Các chữ liên quan (chỉ liệt kê loại có mặt): 
    #   "&TRIM(IF(COUNTIFS(SAO!$F$5:$F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)>0,"Chính quan vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$
    #   F$20,"Chính quan",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","")&IF(COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$
    #   20,1,SAO!$J$5:$J$20,1)>0,"Thất sát vượng tại lệnh: "&COUNTIFS(SAO!$F$5:$F$20,"Thất sát",SAO!$Y$5:$Y$20,0,SAO!$X$5:$X$20,1,SAO!$J$5:$J$20,1)&" chữ; ","
    #   "))&" → Quy tắc (HG tr.165): Thân vượng, quan sát chỉ 1 chữ và không vượng tại lệnh → "&IF(J306="","chưa đủ điều kiện ở cả hai phương pháp – không áp 
    #   dụng.","ĐỦ ("&J306&") ⇒ "&G306&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Số chữ Quan sát: '),E.v(9,99,6)),('' if X.truthy(X.eq(E.v(18,9,4),'')) else X.cat(X.cat(' (',E.v(18,9,4)),')'))),'; '),'Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM(X.cat((X.cat(X.cat('Chính quan vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Chính quan', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else ''),(X.cat(X.cat('Thất sát vượng tại lệnh: ',X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,6,20,6), 'Thất sát', E.rg(13,5,25,20,25), 0.0, E.rg(13,5,24,20,24), 1.0, E.rg(13,5,10,20,10), 1.0),0.0)) else '')))),' → Quy tắc (HG tr.165): Thân vượng, quan sát chỉ 1 chữ và không vượng tại lệnh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f534(E, r, c):
    # Ô E307, F307
    # = IF(OR(AND(TRUE(),(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$C$6+1)=1)),AND(TRUE(),(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$D$6+1)=1)),AND(NOT(NHAP_LIE
    #   U!$C$9="Không rõ giờ"),(INDEX(M_XU,TRUNG_GIAN!$E$6+1,TRUNG_GIAN!$F$6+1)=1))),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.OR(X.AND(X.TRUE(), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,3),1.0)),1.0)), X.AND(X.TRUE(), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,4),1.0)),1.0)), X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.INDEX(E.rg(24,47,16,58,27), X.add(E.v(23,6,5),1.0), X.add(E.v(23,6,6),1.0)),1.0)))) else '—')

def f535(E, r, c):
    # Ô K307
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ tháng: "&TRUNG_GIAN!$D$10&" "&TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"&"; "&"Trụ ngày: "&T
    #   RUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$1
    #   0&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (HG tr.148): Chi ngày bị chi khác 
    #   xung → "&IF(J307="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J307&") ⇒ "&G307&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ tháng: '),E.v(23,10,4)),' '),E.v(23,11,4)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')'),'; '),'Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (HG tr.148): Chi ngày bị chi khác xung → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f536(E, r, c):
    # Ô E308, F308
    # = IF((COUNTIFS(SAO!$U$5:$U$20,1,SAO!$X$5:$X$20,1)>0),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,24,20,24), 1.0),0.0)) else '—')

def f537(E, r, c):
    # Ô K308
    # = "Dữ kiện: "&"Các chữ liên quan (chỉ liệt kê loại có mặt): "&TRIM(IF(COUNTIFS(SAO!$U$5:$U$20,1,SAO!$X$5:$X$20,1)>0,"trạch mã: "&COUNTIFS(SAO!$U$5:$U$20
    #   ,1,SAO!$X$5:$X$20,1)&" chữ; ",""))&" → Quy tắc (HG tr.148): Có trạch mã → "&IF(J308="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ (
    #   "&J308&") ⇒ "&G308&".")
    return X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Các chữ liên quan (chỉ liệt kê loại có mặt): '),X.TRIM((X.cat(X.cat('trạch mã: ',X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,24,20,24), 1.0)),' chữ; ') if X.truthy(X.gt(X.COUNTIFS(E.rg(13,5,21,20,21), 1.0, E.rg(13,5,24,20,24), 1.0),0.0)) else ''))),' → Quy tắc (HG tr.148): Có trạch mã → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f538(E, r, c):
    # Ô E309, F309
    # = IF((COUNTIF(TRUNG_GIAN!$C$5:$E$5,6)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$5=6),1,0)>=2),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.COUNTIF(E.rg(23,5,3,5,5), 6.0),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,5,6),6.0))) else 0.0)),2.0)) else '—')

def f539(E, r, c):
    # Ô K309
    # = "Dữ kiện: "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"T
    #   rụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"&" → Quy tắc (HG tr.14
    #   8): Thiên can có ≥ 2 chữ Canh → "&IF(J309="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J309&") ⇒ "&G309&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (HG tr.148): Thiên can có ≥ 2 chữ Canh → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f540(E, r, c):
    # Ô E310, F310
    # = IF((COUNTIF(TRUNG_GIAN!$C$7:$E$7,4)+COUNTIF(TRUNG_GIAN!$C$9:$E$9,4)+IF(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),(TRUNG_GIAN!$F$7=4)+(TRUNG_GIAN!$F$9=4),0)>=
    #   P_NHIEU),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.ge(X.add(X.add(X.COUNTIF(E.rg(23,7,3,7,5), 4.0),X.COUNTIF(E.rg(23,9,3,9,5), 4.0)),(X.add(X.eq(E.v(23,7,6),4.0),X.eq(E.v(23,9,6),4.0)) if X.truthy(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ'))) else 0.0)),E.v(24,99,2))) else '—')

def f541(E, r, c):
    # Ô K310
    # = "Dữ kiện: "&"ngưỡng nhiều = "&P_NHIEU&" chữ"&"; "&"Trụ năm: "&TRUNG_GIAN!$C$10&" "&TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&"/"&INDEX(H
    #   ANH_TEN,TRUNG_GIAN!$C$9+1)&")"&"; "&"Trụ giờ: "&TRUNG_GIAN!$F$10&" "&TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&"/"&INDEX(HANH_TEN,TRUNG_
    #   GIAN!$F$9+1)&")"&" → Quy tắc (HG tr.148): Can chi hành Thủy ≥ P_NHIEU chữ → "&IF(J310="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ
    #    ("&J310&") ⇒ "&G310&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','ngưỡng nhiều = '),E.v(24,99,2)),' chữ'),'; '),'Trụ năm: '),E.v(23,10,3)),' '),E.v(23,11,3)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')'),'; '),'Trụ giờ: '),E.v(23,10,6)),' '),E.v(23,11,6)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')'),' → Quy tắc (HG tr.148): Can chi hành Thủy ≥ P_NHIEU chữ → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f542(E, r, c):
    # Ô E311, F311
    # = IF((TRUNG_GIAN!$E$7=4),"ĐỦ","—")
    return ('ĐỦ' if X.truthy(X.eq(E.v(23,7,5),4.0)) else '—')

def f543(E, r, c):
    # Ô K311
    # = "Dữ kiện: "&"Trụ ngày: "&TRUNG_GIAN!$E$10&" "&TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"/"&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"&" → Qu
    #   y tắc (HG tr.148): Nhật chủ hành Thủy → "&IF(J311="","chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.","ĐỦ ("&J311&") ⇒ "&G311&".")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dữ kiện: ','Trụ ngày: '),E.v(23,10,5)),' '),E.v(23,11,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'/'),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')'),' → Quy tắc (HG tr.148): Nhật chủ hành Thủy → '),('chưa đủ điều kiện ở cả hai phương pháp – không áp dụng.' if X.truthy(X.eq(E.v(21,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat('ĐỦ (',E.v(21,r,c-1)),') ⇒ '),E.v(21,r,c-4)),'.')))

def f544(E, r, c):
    # Ô D313
    # = "Dụng thần – "&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)<0,"",CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+1,"M
    #   ộc: sơn lâm, gỗ, nhà cửa, lâm viên, vườn hoa, cây cảnh","Hỏa: điện lực, nhiên liệu, học thuật, dạy học, thông tin điện tử","Thổ: kiến trúc, công trình
    #   , nông sản, dệt may, hậu cần","Kim: vàng bạc, kim khí, máy móc, ngân hàng","Thủy: thủy lợi, đồ uống, y dược, vận tải, du lịch, hàng hải, giao thông, n
    #   hà nghỉ, giải trí"))&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)<0,"","; hỷ thần – "&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_H
    #   Y_KY!$C$110)<0,"",CHOOSE(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1,"Mộc: sơn lâm, gỗ, nhà cửa, lâm viên, vườn hoa, cây cảnh","Hỏa: đ
    #   iện lực, nhiên liệu, học thuật, dạy học, thông tin điện tử","Thổ: kiến trúc, công trình, nông sản, dệt may, hậu cần","Kim: vàng bạc, kim khí, máy móc,
    #    ngân hàng","Thủy: thủy lợi, đồ uống, y dược, vận tải, du lịch, hàng hải, giao thông, nhà nghỉ, giải trí")))&" (HG tr.156)."
    return X.cat(X.cat(X.cat('Dụng thần – ',('' if X.truthy(X.lt((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),0.0)) else X.CHOOSE(X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0), lambda: 'Mộc: sơn lâm, gỗ, nhà cửa, lâm viên, vườn hoa, cây cảnh', lambda: 'Hỏa: điện lực, nhiên liệu, học thuật, dạy học, thông tin điện tử', lambda: 'Thổ: kiến trúc, công trình, nông sản, dệt may, hậu cần', lambda: 'Kim: vàng bạc, kim khí, máy móc, ngân hàng', lambda: 'Thủy: thủy lợi, đồ uống, y dược, vận tải, du lịch, hàng hải, giao thông, nhà nghỉ, giải trí'))),('' if X.truthy(X.lt((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else X.cat('; hỷ thần – ',('' if X.truthy(X.lt((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),0.0)) else X.CHOOSE(X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0), lambda: 'Mộc: sơn lâm, gỗ, nhà cửa, lâm viên, vườn hoa, cây cảnh', lambda: 'Hỏa: điện lực, nhiên liệu, học thuật, dạy học, thông tin điện tử', lambda: 'Thổ: kiến trúc, công trình, nông sản, dệt may, hậu cần', lambda: 'Kim: vàng bạc, kim khí, máy móc, ngân hàng', lambda: 'Thủy: thủy lợi, đồ uống, y dược, vận tải, du lịch, hàng hải, giao thông, nhà nghỉ, giải trí'))))),' (HG tr.156).')

CELLS = {
    'Z2': (f0, 2, 26),
    'C6': (f1, 6, 3),
    'F6': (f2, 6, 6),
    'J6': (f3, 6, 10),
    'M6': (f4, 6, 13),
    'N6': (f5, 6, 14),
    'AB6': (f6, 6, 28),
    'AC6': (f7, 6, 29),
    'AD6': (f8, 6, 30),
    'AE6': (f9, 6, 31),
    'C7': (f1, 7, 3),
    'F7': (f2, 7, 6),
    'J7': (f10, 7, 10),
    'M7': (f11, 7, 13),
    'N7': (f12, 7, 14),
    'AA7': (f13, 7, 27),
    'AB7': (f14, 7, 28),
    'AC7': (f15, 7, 29),
    'AD7': (f16, 7, 30),
    'AE7': (f17, 7, 31),
    'C8': (f1, 8, 3),
    'F8': (f2, 8, 6),
    'J8': (f18, 8, 10),
    'M8': (f19, 8, 13),
    'N8': (f20, 8, 14),
    'AB8': (f21, 8, 28),
    'AC8': (f22, 8, 29),
    'AD8': (f23, 8, 30),
    'AE8': (f24, 8, 31),
    'C9': (f1, 9, 3),
    'F9': (f2, 9, 6),
    'J9': (f25, 9, 10),
    'M9': (f26, 9, 13),
    'N9': (f27, 9, 14),
    'AB9': (f28, 9, 28),
    'AC9': (f29, 9, 29),
    'AD9': (f30, 9, 30),
    'AE9': (f31, 9, 31),
    'C13': (f32, 13, 3),
    'D13': (f33, 13, 4),
    'E13': (f34, 13, 5),
    'F13': (f35, 13, 6),
    'G13': (f36, 13, 7),
    'H13': (f37, 13, 8),
    'I13': (f38, 13, 9),
    'J13': (f39, 13, 10),
    'C14': (f40, 14, 3),
    'D14': (f40, 14, 4),
    'E14': (f40, 14, 5),
    'F14': (f40, 14, 6),
    'G14': (f40, 14, 7),
    'H14': (f40, 14, 8),
    'I14': (f40, 14, 9),
    'J14': (f40, 14, 10),
    'C15': (f41, 15, 3),
    'D15': (f41, 15, 4),
    'E15': (f41, 15, 5),
    'F15': (f41, 15, 6),
    'G15': (f41, 15, 7),
    'H15': (f41, 15, 8),
    'I15': (f41, 15, 9),
    'J15': (f41, 15, 10),
    'C16': (f42, 16, 3),
    'D16': (f42, 16, 4),
    'E16': (f42, 16, 5),
    'F16': (f42, 16, 6),
    'G16': (f42, 16, 7),
    'H16': (f42, 16, 8),
    'I16': (f42, 16, 9),
    'J16': (f42, 16, 10),
    'C17': (f43, 17, 3),
    'D17': (f43, 17, 4),
    'E17': (f43, 17, 5),
    'F17': (f43, 17, 6),
    'G17': (f43, 17, 7),
    'H17': (f43, 17, 8),
    'I17': (f43, 17, 9),
    'J17': (f43, 17, 10),
    'C18': (f43, 18, 3),
    'D18': (f43, 18, 4),
    'E18': (f43, 18, 5),
    'F18': (f43, 18, 6),
    'G18': (f43, 18, 7),
    'H18': (f43, 18, 8),
    'I18': (f43, 18, 9),
    'J18': (f43, 18, 10),
    'C19': (f44, 19, 3),
    'D19': (f44, 19, 4),
    'E19': (f44, 19, 5),
    'F19': (f44, 19, 6),
    'G19': (f44, 19, 7),
    'H19': (f44, 19, 8),
    'I19': (f44, 19, 9),
    'J19': (f44, 19, 10),
    'C20': (f45, 20, 3),
    'D20': (f45, 20, 4),
    'E20': (f45, 20, 5),
    'F20': (f45, 20, 6),
    'G20': (f45, 20, 7),
    'H20': (f45, 20, 8),
    'I20': (f45, 20, 9),
    'J20': (f45, 20, 10),
    'C21': (f46, 21, 3),
    'D21': (f46, 21, 4),
    'E21': (f46, 21, 5),
    'F21': (f46, 21, 6),
    'G21': (f46, 21, 7),
    'H21': (f46, 21, 8),
    'I21': (f46, 21, 9),
    'J21': (f46, 21, 10),
    'C22': (f47, 22, 3),
    'M22': (f48, 22, 13),
    'N22': (f48, 22, 14),
    'O22': (f48, 22, 15),
    'P22': (f48, 22, 16),
    'Q22': (f48, 22, 17),
    'R22': (f48, 22, 18),
    'S22': (f48, 22, 19),
    'T22': (f48, 22, 20),
    'C23': (f49, 23, 3),
    'M23': (f50, 23, 13),
    'N23': (f50, 23, 14),
    'O23': (f50, 23, 15),
    'P23': (f50, 23, 16),
    'Q23': (f50, 23, 17),
    'R23': (f50, 23, 18),
    'S23': (f50, 23, 19),
    'T23': (f50, 23, 20),
    'C24': (f51, 24, 3),
    'M24': (f52, 24, 13),
    'N24': (f52, 24, 14),
    'O24': (f52, 24, 15),
    'P24': (f52, 24, 16),
    'Q24': (f52, 24, 17),
    'R24': (f52, 24, 18),
    'S24': (f52, 24, 19),
    'T24': (f52, 24, 20),
    'C27': (f53, 27, 3),
    'C28': (f54, 28, 3),
    'E62': (f55, 62, 5),
    'F62': (f55, 62, 6),
    'J62': (f56, 62, 10),
    'K62': (f57, 62, 11),
    'L62': (f58, 62, 12),
    'E63': (f59, 63, 5),
    'F63': (f59, 63, 6),
    'J63': (f56, 63, 10),
    'K63': (f60, 63, 11),
    'L63': (f58, 63, 12),
    'E64': (f61, 64, 5),
    'F64': (f62, 64, 6),
    'J64': (f56, 64, 10),
    'K64': (f63, 64, 11),
    'L64': (f58, 64, 12),
    'E65': (f64, 65, 5),
    'F65': (f65, 65, 6),
    'J65': (f56, 65, 10),
    'K65': (f66, 65, 11),
    'L65': (f58, 65, 12),
    'E66': (f67, 66, 5),
    'F66': (f68, 66, 6),
    'J66': (f56, 66, 10),
    'K66': (f69, 66, 11),
    'L66': (f58, 66, 12),
    'E67': (f70, 67, 5),
    'F67': (f70, 67, 6),
    'J67': (f56, 67, 10),
    'K67': (f71, 67, 11),
    'L67': (f58, 67, 12),
    'E68': (f72, 68, 5),
    'F68': (f72, 68, 6),
    'J68': (f56, 68, 10),
    'K68': (f73, 68, 11),
    'L68': (f58, 68, 12),
    'E69': (f74, 69, 5),
    'F69': (f74, 69, 6),
    'J69': (f56, 69, 10),
    'K69': (f75, 69, 11),
    'L69': (f58, 69, 12),
    'E70': (f76, 70, 5),
    'F70': (f76, 70, 6),
    'J70': (f56, 70, 10),
    'K70': (f77, 70, 11),
    'L70': (f58, 70, 12),
    'E71': (f78, 71, 5),
    'F71': (f78, 71, 6),
    'J71': (f56, 71, 10),
    'K71': (f79, 71, 11),
    'L71': (f58, 71, 12),
    'E72': (f80, 72, 5),
    'F72': (f81, 72, 6),
    'J72': (f56, 72, 10),
    'K72': (f82, 72, 11),
    'L72': (f58, 72, 12),
    'E73': (f83, 73, 5),
    'F73': (f84, 73, 6),
    'J73': (f56, 73, 10),
    'K73': (f85, 73, 11),
    'L73': (f58, 73, 12),
    'E74': (f86, 74, 5),
    'F74': (f86, 74, 6),
    'J74': (f56, 74, 10),
    'K74': (f87, 74, 11),
    'L74': (f58, 74, 12),
    'E75': (f88, 75, 5),
    'F75': (f88, 75, 6),
    'J75': (f56, 75, 10),
    'K75': (f89, 75, 11),
    'L75': (f58, 75, 12),
    'E76': (f90, 76, 5),
    'F76': (f90, 76, 6),
    'J76': (f56, 76, 10),
    'K76': (f91, 76, 11),
    'L76': (f58, 76, 12),
    'E77': (f92, 77, 5),
    'F77': (f92, 77, 6),
    'J77': (f56, 77, 10),
    'K77': (f93, 77, 11),
    'L77': (f58, 77, 12),
    'E78': (f94, 78, 5),
    'F78': (f95, 78, 6),
    'J78': (f56, 78, 10),
    'K78': (f96, 78, 11),
    'L78': (f58, 78, 12),
    'E79': (f97, 79, 5),
    'F79': (f98, 79, 6),
    'J79': (f56, 79, 10),
    'K79': (f99, 79, 11),
    'L79': (f58, 79, 12),
    'E80': (f100, 80, 5),
    'F80': (f100, 80, 6),
    'J80': (f56, 80, 10),
    'K80': (f101, 80, 11),
    'L80': (f58, 80, 12),
    'E81': (f102, 81, 5),
    'F81': (f102, 81, 6),
    'J81': (f56, 81, 10),
    'K81': (f103, 81, 11),
    'L81': (f58, 81, 12),
    'E82': (f104, 82, 5),
    'F82': (f104, 82, 6),
    'J82': (f56, 82, 10),
    'K82': (f105, 82, 11),
    'L82': (f58, 82, 12),
    'E83': (f106, 83, 5),
    'F83': (f106, 83, 6),
    'J83': (f56, 83, 10),
    'K83': (f107, 83, 11),
    'L83': (f58, 83, 12),
    'E84': (f108, 84, 5),
    'F84': (f109, 84, 6),
    'J84': (f56, 84, 10),
    'K84': (f110, 84, 11),
    'L84': (f58, 84, 12),
    'E85': (f111, 85, 5),
    'F85': (f111, 85, 6),
    'J85': (f56, 85, 10),
    'K85': (f112, 85, 11),
    'L85': (f58, 85, 12),
    'E86': (f108, 86, 5),
    'F86': (f109, 86, 6),
    'J86': (f56, 86, 10),
    'K86': (f113, 86, 11),
    'L86': (f58, 86, 12),
    'E87': (f114, 87, 5),
    'F87': (f115, 87, 6),
    'J87': (f56, 87, 10),
    'K87': (f116, 87, 11),
    'L87': (f58, 87, 12),
    'E88': (f117, 88, 5),
    'F88': (f117, 88, 6),
    'J88': (f56, 88, 10),
    'K88': (f118, 88, 11),
    'L88': (f58, 88, 12),
    'E89': (f119, 89, 5),
    'F89': (f119, 89, 6),
    'J89': (f56, 89, 10),
    'K89': (f120, 89, 11),
    'L89': (f58, 89, 12),
    'E90': (f121, 90, 5),
    'F90': (f121, 90, 6),
    'J90': (f56, 90, 10),
    'K90': (f122, 90, 11),
    'L90': (f58, 90, 12),
    'E91': (f123, 91, 5),
    'F91': (f124, 91, 6),
    'J91': (f56, 91, 10),
    'K91': (f125, 91, 11),
    'L91': (f58, 91, 12),
    'E92': (f126, 92, 5),
    'F92': (f126, 92, 6),
    'J92': (f56, 92, 10),
    'K92': (f127, 92, 11),
    'L92': (f58, 92, 12),
    'E93': (f128, 93, 5),
    'F93': (f129, 93, 6),
    'J93': (f56, 93, 10),
    'K93': (f130, 93, 11),
    'L93': (f58, 93, 12),
    'E94': (f131, 94, 5),
    'F94': (f132, 94, 6),
    'J94': (f56, 94, 10),
    'K94': (f133, 94, 11),
    'L94': (f58, 94, 12),
    'E95': (f134, 95, 5),
    'F95': (f134, 95, 6),
    'J95': (f56, 95, 10),
    'K95': (f135, 95, 11),
    'L95': (f58, 95, 12),
    'E96': (f136, 96, 5),
    'F96': (f136, 96, 6),
    'J96': (f56, 96, 10),
    'K96': (f137, 96, 11),
    'L96': (f58, 96, 12),
    'E97': (f138, 97, 5),
    'F97': (f138, 97, 6),
    'J97': (f56, 97, 10),
    'K97': (f139, 97, 11),
    'L97': (f58, 97, 12),
    'E98': (f140, 98, 5),
    'F98': (f141, 98, 6),
    'J98': (f56, 98, 10),
    'K98': (f142, 98, 11),
    'L98': (f58, 98, 12),
    'E99': (f143, 99, 5),
    'F99': (f144, 99, 6),
    'J99': (f56, 99, 10),
    'K99': (f145, 99, 11),
    'L99': (f58, 99, 12),
    'E100': (f146, 100, 5),
    'F100': (f146, 100, 6),
    'J100': (f56, 100, 10),
    'K100': (f147, 100, 11),
    'L100': (f58, 100, 12),
    'E101': (f148, 101, 5),
    'F101': (f149, 101, 6),
    'J101': (f56, 101, 10),
    'K101': (f150, 101, 11),
    'L101': (f58, 101, 12),
    'E102': (f151, 102, 5),
    'F102': (f151, 102, 6),
    'J102': (f56, 102, 10),
    'K102': (f152, 102, 11),
    'L102': (f58, 102, 12),
    'E103': (f64, 103, 5),
    'F103': (f65, 103, 6),
    'J103': (f56, 103, 10),
    'K103': (f153, 103, 11),
    'L103': (f58, 103, 12),
    'E104': (f154, 104, 5),
    'F104': (f154, 104, 6),
    'J104': (f56, 104, 10),
    'K104': (f155, 104, 11),
    'L104': (f58, 104, 12),
    'E105': (f156, 105, 5),
    'F105': (f156, 105, 6),
    'J105': (f56, 105, 10),
    'K105': (f157, 105, 11),
    'L105': (f58, 105, 12),
    'E106': (f158, 106, 5),
    'F106': (f158, 106, 6),
    'J106': (f56, 106, 10),
    'K106': (f159, 106, 11),
    'L106': (f58, 106, 12),
    'E107': (f160, 107, 5),
    'F107': (f160, 107, 6),
    'J107': (f56, 107, 10),
    'K107': (f161, 107, 11),
    'L107': (f58, 107, 12),
    'E108': (f162, 108, 5),
    'F108': (f162, 108, 6),
    'J108': (f56, 108, 10),
    'K108': (f163, 108, 11),
    'L108': (f58, 108, 12),
    'E109': (f164, 109, 5),
    'F109': (f164, 109, 6),
    'J109': (f56, 109, 10),
    'K109': (f165, 109, 11),
    'L109': (f58, 109, 12),
    'E110': (f166, 110, 5),
    'F110': (f166, 110, 6),
    'J110': (f56, 110, 10),
    'K110': (f167, 110, 11),
    'L110': (f58, 110, 12),
    'E111': (f86, 111, 5),
    'F111': (f86, 111, 6),
    'J111': (f56, 111, 10),
    'K111': (f168, 111, 11),
    'L111': (f58, 111, 12),
    'E112': (f169, 112, 5),
    'F112': (f169, 112, 6),
    'J112': (f56, 112, 10),
    'K112': (f170, 112, 11),
    'L112': (f58, 112, 12),
    'E113': (f171, 113, 5),
    'F113': (f171, 113, 6),
    'J113': (f56, 113, 10),
    'K113': (f172, 113, 11),
    'L113': (f58, 113, 12),
    'E114': (f173, 114, 5),
    'F114': (f173, 114, 6),
    'J114': (f56, 114, 10),
    'K114': (f174, 114, 11),
    'L114': (f58, 114, 12),
    'E115': (f175, 115, 5),
    'F115': (f175, 115, 6),
    'J115': (f56, 115, 10),
    'K115': (f176, 115, 11),
    'L115': (f58, 115, 12),
    'E116': (f177, 116, 5),
    'F116': (f177, 116, 6),
    'J116': (f56, 116, 10),
    'K116': (f178, 116, 11),
    'L116': (f58, 116, 12),
    'E117': (f179, 117, 5),
    'F117': (f179, 117, 6),
    'J117': (f56, 117, 10),
    'K117': (f180, 117, 11),
    'L117': (f58, 117, 12),
    'E118': (f181, 118, 5),
    'F118': (f181, 118, 6),
    'J118': (f56, 118, 10),
    'K118': (f182, 118, 11),
    'L118': (f58, 118, 12),
    'E119': (f183, 119, 5),
    'F119': (f183, 119, 6),
    'J119': (f56, 119, 10),
    'K119': (f184, 119, 11),
    'L119': (f58, 119, 12),
    'E120': (f185, 120, 5),
    'F120': (f185, 120, 6),
    'J120': (f56, 120, 10),
    'K120': (f186, 120, 11),
    'L120': (f58, 120, 12),
    'E121': (f187, 121, 5),
    'F121': (f187, 121, 6),
    'J121': (f56, 121, 10),
    'K121': (f188, 121, 11),
    'L121': (f58, 121, 12),
    'E122': (f189, 122, 5),
    'F122': (f190, 122, 6),
    'J122': (f56, 122, 10),
    'K122': (f191, 122, 11),
    'L122': (f58, 122, 12),
    'E123': (f192, 123, 5),
    'F123': (f192, 123, 6),
    'J123': (f56, 123, 10),
    'K123': (f193, 123, 11),
    'L123': (f58, 123, 12),
    'E124': (f194, 124, 5),
    'F124': (f195, 124, 6),
    'J124': (f56, 124, 10),
    'K124': (f196, 124, 11),
    'L124': (f58, 124, 12),
    'E125': (f90, 125, 5),
    'F125': (f90, 125, 6),
    'J125': (f56, 125, 10),
    'K125': (f197, 125, 11),
    'L125': (f58, 125, 12),
    'E126': (f156, 126, 5),
    'F126': (f156, 126, 6),
    'J126': (f56, 126, 10),
    'K126': (f198, 126, 11),
    'L126': (f58, 126, 12),
    'E127': (f199, 127, 5),
    'F127': (f200, 127, 6),
    'J127': (f56, 127, 10),
    'K127': (f201, 127, 11),
    'L127': (f58, 127, 12),
    'E128': (f202, 128, 5),
    'F128': (f202, 128, 6),
    'J128': (f56, 128, 10),
    'K128': (f203, 128, 11),
    'L128': (f58, 128, 12),
    'E129': (f204, 129, 5),
    'F129': (f204, 129, 6),
    'J129': (f56, 129, 10),
    'K129': (f205, 129, 11),
    'L129': (f58, 129, 12),
    'E130': (f206, 130, 5),
    'F130': (f206, 130, 6),
    'J130': (f56, 130, 10),
    'K130': (f207, 130, 11),
    'L130': (f58, 130, 12),
    'E131': (f208, 131, 5),
    'F131': (f208, 131, 6),
    'J131': (f56, 131, 10),
    'K131': (f209, 131, 11),
    'L131': (f58, 131, 12),
    'E132': (f210, 132, 5),
    'F132': (f210, 132, 6),
    'J132': (f56, 132, 10),
    'K132': (f211, 132, 11),
    'L132': (f58, 132, 12),
    'E133': (f212, 133, 5),
    'F133': (f212, 133, 6),
    'J133': (f56, 133, 10),
    'K133': (f213, 133, 11),
    'L133': (f58, 133, 12),
    'E134': (f214, 134, 5),
    'F134': (f214, 134, 6),
    'J134': (f56, 134, 10),
    'K134': (f215, 134, 11),
    'L134': (f58, 134, 12),
    'E135': (f216, 135, 5),
    'F135': (f216, 135, 6),
    'J135': (f56, 135, 10),
    'K135': (f217, 135, 11),
    'L135': (f58, 135, 12),
    'E136': (f218, 136, 5),
    'F136': (f218, 136, 6),
    'J136': (f56, 136, 10),
    'K136': (f219, 136, 11),
    'L136': (f58, 136, 12),
    'E137': (f220, 137, 5),
    'F137': (f220, 137, 6),
    'J137': (f56, 137, 10),
    'K137': (f221, 137, 11),
    'L137': (f58, 137, 12),
    'E138': (f222, 138, 5),
    'F138': (f223, 138, 6),
    'J138': (f56, 138, 10),
    'K138': (f224, 138, 11),
    'L138': (f58, 138, 12),
    'E139': (f177, 139, 5),
    'F139': (f177, 139, 6),
    'J139': (f56, 139, 10),
    'K139': (f225, 139, 11),
    'L139': (f58, 139, 12),
    'E140': (f226, 140, 5),
    'F140': (f226, 140, 6),
    'J140': (f56, 140, 10),
    'K140': (f227, 140, 11),
    'L140': (f58, 140, 12),
    'E141': (f228, 141, 5),
    'F141': (f229, 141, 6),
    'J141': (f56, 141, 10),
    'K141': (f230, 141, 11),
    'L141': (f58, 141, 12),
    'E142': (f231, 142, 5),
    'F142': (f232, 142, 6),
    'J142': (f56, 142, 10),
    'K142': (f233, 142, 11),
    'L142': (f58, 142, 12),
    'E143': (f234, 143, 5),
    'F143': (f234, 143, 6),
    'J143': (f56, 143, 10),
    'K143': (f235, 143, 11),
    'L143': (f58, 143, 12),
    'E144': (f236, 144, 5),
    'F144': (f236, 144, 6),
    'J144': (f56, 144, 10),
    'K144': (f237, 144, 11),
    'L144': (f58, 144, 12),
    'E145': (f86, 145, 5),
    'F145': (f86, 145, 6),
    'J145': (f56, 145, 10),
    'K145': (f238, 145, 11),
    'L145': (f58, 145, 12),
    'E146': (f204, 146, 5),
    'F146': (f204, 146, 6),
    'J146': (f56, 146, 10),
    'K146': (f239, 146, 11),
    'L146': (f58, 146, 12),
    'E147': (f240, 147, 5),
    'F147': (f240, 147, 6),
    'J147': (f56, 147, 10),
    'K147': (f241, 147, 11),
    'L147': (f58, 147, 12),
    'E148': (f242, 148, 5),
    'F148': (f243, 148, 6),
    'J148': (f56, 148, 10),
    'K148': (f244, 148, 11),
    'L148': (f58, 148, 12),
    'E149': (f245, 149, 5),
    'F149': (f246, 149, 6),
    'J149': (f56, 149, 10),
    'K149': (f247, 149, 11),
    'L149': (f58, 149, 12),
    'E150': (f90, 150, 5),
    'F150': (f90, 150, 6),
    'J150': (f56, 150, 10),
    'K150': (f248, 150, 11),
    'L150': (f58, 150, 12),
    'E151': (f249, 151, 5),
    'F151': (f249, 151, 6),
    'J151': (f56, 151, 10),
    'K151': (f250, 151, 11),
    'L151': (f58, 151, 12),
    'E152': (f251, 152, 5),
    'F152': (f251, 152, 6),
    'J152': (f56, 152, 10),
    'K152': (f252, 152, 11),
    'L152': (f58, 152, 12),
    'E153': (f253, 153, 5),
    'F153': (f253, 153, 6),
    'J153': (f56, 153, 10),
    'K153': (f254, 153, 11),
    'L153': (f58, 153, 12),
    'E154': (f255, 154, 5),
    'F154': (f255, 154, 6),
    'J154': (f56, 154, 10),
    'K154': (f256, 154, 11),
    'L154': (f58, 154, 12),
    'E155': (f257, 155, 5),
    'F155': (f257, 155, 6),
    'J155': (f56, 155, 10),
    'K155': (f258, 155, 11),
    'L155': (f58, 155, 12),
    'E156': (f259, 156, 5),
    'F156': (f260, 156, 6),
    'J156': (f56, 156, 10),
    'K156': (f261, 156, 11),
    'L156': (f58, 156, 12),
    'E160': (f262, 160, 5),
    'F160': (f262, 160, 6),
    'J160': (f56, 160, 10),
    'K160': (f263, 160, 11),
    'L160': (f58, 160, 12),
    'E161': (f264, 161, 5),
    'F161': (f264, 161, 6),
    'J161': (f56, 161, 10),
    'K161': (f265, 161, 11),
    'L161': (f58, 161, 12),
    'E162': (f266, 162, 5),
    'F162': (f266, 162, 6),
    'J162': (f56, 162, 10),
    'K162': (f267, 162, 11),
    'L162': (f58, 162, 12),
    'E163': (f268, 163, 5),
    'F163': (f268, 163, 6),
    'J163': (f56, 163, 10),
    'K163': (f269, 163, 11),
    'L163': (f58, 163, 12),
    'E164': (f270, 164, 5),
    'F164': (f271, 164, 6),
    'J164': (f56, 164, 10),
    'K164': (f272, 164, 11),
    'L164': (f58, 164, 12),
    'E165': (f273, 165, 5),
    'F165': (f273, 165, 6),
    'J165': (f56, 165, 10),
    'K165': (f274, 165, 11),
    'L165': (f58, 165, 12),
    'E166': (f275, 166, 5),
    'F166': (f275, 166, 6),
    'J166': (f56, 166, 10),
    'K166': (f276, 166, 11),
    'L166': (f58, 166, 12),
    'E167': (f277, 167, 5),
    'F167': (f277, 167, 6),
    'J167': (f56, 167, 10),
    'K167': (f278, 167, 11),
    'L167': (f58, 167, 12),
    'E168': (f279, 168, 5),
    'F168': (f279, 168, 6),
    'J168': (f56, 168, 10),
    'K168': (f280, 168, 11),
    'L168': (f58, 168, 12),
    'E169': (f281, 169, 5),
    'F169': (f281, 169, 6),
    'J169': (f56, 169, 10),
    'K169': (f282, 169, 11),
    'L169': (f58, 169, 12),
    'E170': (f283, 170, 5),
    'F170': (f283, 170, 6),
    'J170': (f56, 170, 10),
    'K170': (f284, 170, 11),
    'L170': (f58, 170, 12),
    'E171': (f285, 171, 5),
    'F171': (f286, 171, 6),
    'J171': (f56, 171, 10),
    'K171': (f287, 171, 11),
    'L171': (f58, 171, 12),
    'E172': (f288, 172, 5),
    'F172': (f288, 172, 6),
    'J172': (f56, 172, 10),
    'K172': (f289, 172, 11),
    'L172': (f58, 172, 12),
    'E173': (f290, 173, 5),
    'F173': (f290, 173, 6),
    'J173': (f56, 173, 10),
    'K173': (f291, 173, 11),
    'L173': (f58, 173, 12),
    'E174': (f292, 174, 5),
    'F174': (f292, 174, 6),
    'J174': (f56, 174, 10),
    'K174': (f293, 174, 11),
    'L174': (f58, 174, 12),
    'E175': (f294, 175, 5),
    'F175': (f294, 175, 6),
    'J175': (f56, 175, 10),
    'K175': (f295, 175, 11),
    'L175': (f58, 175, 12),
    'E176': (f296, 176, 5),
    'F176': (f296, 176, 6),
    'J176': (f56, 176, 10),
    'K176': (f297, 176, 11),
    'L176': (f58, 176, 12),
    'E177': (f298, 177, 5),
    'F177': (f298, 177, 6),
    'J177': (f56, 177, 10),
    'K177': (f299, 177, 11),
    'L177': (f58, 177, 12),
    'E178': (f300, 178, 5),
    'F178': (f300, 178, 6),
    'J178': (f56, 178, 10),
    'K178': (f301, 178, 11),
    'L178': (f58, 178, 12),
    'E179': (f302, 179, 5),
    'F179': (f302, 179, 6),
    'J179': (f56, 179, 10),
    'K179': (f303, 179, 11),
    'L179': (f58, 179, 12),
    'E180': (f304, 180, 5),
    'F180': (f304, 180, 6),
    'J180': (f56, 180, 10),
    'K180': (f305, 180, 11),
    'L180': (f58, 180, 12),
    'E181': (f306, 181, 5),
    'F181': (f306, 181, 6),
    'J181': (f56, 181, 10),
    'K181': (f307, 181, 11),
    'L181': (f58, 181, 12),
    'E182': (f308, 182, 5),
    'F182': (f308, 182, 6),
    'J182': (f56, 182, 10),
    'K182': (f309, 182, 11),
    'L182': (f58, 182, 12),
    'E183': (f310, 183, 5),
    'F183': (f310, 183, 6),
    'J183': (f56, 183, 10),
    'K183': (f311, 183, 11),
    'L183': (f58, 183, 12),
    'E184': (f312, 184, 5),
    'F184': (f312, 184, 6),
    'J184': (f56, 184, 10),
    'K184': (f313, 184, 11),
    'L184': (f58, 184, 12),
    'E185': (f314, 185, 5),
    'F185': (f314, 185, 6),
    'J185': (f56, 185, 10),
    'K185': (f315, 185, 11),
    'L185': (f58, 185, 12),
    'E186': (f316, 186, 5),
    'F186': (f316, 186, 6),
    'J186': (f56, 186, 10),
    'K186': (f317, 186, 11),
    'L186': (f58, 186, 12),
    'E187': (f318, 187, 5),
    'F187': (f319, 187, 6),
    'J187': (f56, 187, 10),
    'K187': (f320, 187, 11),
    'L187': (f58, 187, 12),
    'E188': (f321, 188, 5),
    'F188': (f321, 188, 6),
    'J188': (f56, 188, 10),
    'K188': (f322, 188, 11),
    'L188': (f58, 188, 12),
    'E189': (f323, 189, 5),
    'F189': (f323, 189, 6),
    'J189': (f56, 189, 10),
    'K189': (f324, 189, 11),
    'L189': (f58, 189, 12),
    'E190': (f325, 190, 5),
    'F190': (f325, 190, 6),
    'J190': (f56, 190, 10),
    'K190': (f326, 190, 11),
    'L190': (f58, 190, 12),
    'E191': (f327, 191, 5),
    'F191': (f327, 191, 6),
    'J191': (f56, 191, 10),
    'K191': (f328, 191, 11),
    'L191': (f58, 191, 12),
    'E192': (f329, 192, 5),
    'F192': (f329, 192, 6),
    'J192': (f56, 192, 10),
    'K192': (f330, 192, 11),
    'L192': (f58, 192, 12),
    'E193': (f331, 193, 5),
    'F193': (f331, 193, 6),
    'J193': (f56, 193, 10),
    'K193': (f332, 193, 11),
    'L193': (f58, 193, 12),
    'E194': (f333, 194, 5),
    'F194': (f333, 194, 6),
    'J194': (f56, 194, 10),
    'K194': (f334, 194, 11),
    'L194': (f58, 194, 12),
    'E195': (f202, 195, 5),
    'F195': (f202, 195, 6),
    'J195': (f56, 195, 10),
    'K195': (f335, 195, 11),
    'L195': (f58, 195, 12),
    'E196': (f336, 196, 5),
    'F196': (f336, 196, 6),
    'J196': (f56, 196, 10),
    'K196': (f337, 196, 11),
    'L196': (f58, 196, 12),
    'E197': (f338, 197, 5),
    'F197': (f338, 197, 6),
    'J197': (f56, 197, 10),
    'K197': (f339, 197, 11),
    'L197': (f58, 197, 12),
    'E201': (f340, 201, 5),
    'F201': (f340, 201, 6),
    'J201': (f56, 201, 10),
    'K201': (f341, 201, 11),
    'L201': (f58, 201, 12),
    'E202': (f342, 202, 5),
    'F202': (f343, 202, 6),
    'J202': (f56, 202, 10),
    'K202': (f344, 202, 11),
    'L202': (f58, 202, 12),
    'E203': (f345, 203, 5),
    'F203': (f345, 203, 6),
    'J203': (f56, 203, 10),
    'K203': (f346, 203, 11),
    'L203': (f58, 203, 12),
    'E204': (f154, 204, 5),
    'F204': (f154, 204, 6),
    'J204': (f56, 204, 10),
    'K204': (f347, 204, 11),
    'L204': (f58, 204, 12),
    'E205': (f348, 205, 5),
    'F205': (f348, 205, 6),
    'J205': (f56, 205, 10),
    'K205': (f349, 205, 11),
    'L205': (f58, 205, 12),
    'E206': (f350, 206, 5),
    'F206': (f350, 206, 6),
    'J206': (f56, 206, 10),
    'K206': (f351, 206, 11),
    'L206': (f58, 206, 12),
    'E207': (f352, 207, 5),
    'F207': (f352, 207, 6),
    'J207': (f56, 207, 10),
    'K207': (f353, 207, 11),
    'L207': (f58, 207, 12),
    'E208': (f354, 208, 5),
    'F208': (f354, 208, 6),
    'J208': (f56, 208, 10),
    'K208': (f355, 208, 11),
    'L208': (f58, 208, 12),
    'E209': (f356, 209, 5),
    'F209': (f356, 209, 6),
    'J209': (f56, 209, 10),
    'K209': (f357, 209, 11),
    'L209': (f58, 209, 12),
    'E210': (f358, 210, 5),
    'F210': (f358, 210, 6),
    'J210': (f56, 210, 10),
    'K210': (f359, 210, 11),
    'L210': (f58, 210, 12),
    'E211': (f360, 211, 5),
    'F211': (f361, 211, 6),
    'J211': (f56, 211, 10),
    'K211': (f362, 211, 11),
    'L211': (f58, 211, 12),
    'E212': (f363, 212, 5),
    'F212': (f363, 212, 6),
    'J212': (f56, 212, 10),
    'K212': (f364, 212, 11),
    'L212': (f58, 212, 12),
    'E213': (f365, 213, 5),
    'F213': (f365, 213, 6),
    'J213': (f56, 213, 10),
    'K213': (f366, 213, 11),
    'L213': (f58, 213, 12),
    'E214': (f367, 214, 5),
    'F214': (f367, 214, 6),
    'J214': (f56, 214, 10),
    'K214': (f368, 214, 11),
    'L214': (f58, 214, 12),
    'E215': (f369, 215, 5),
    'F215': (f369, 215, 6),
    'J215': (f56, 215, 10),
    'K215': (f370, 215, 11),
    'L215': (f58, 215, 12),
    'E216': (f371, 216, 5),
    'F216': (f371, 216, 6),
    'J216': (f56, 216, 10),
    'K216': (f372, 216, 11),
    'L216': (f58, 216, 12),
    'E217': (f373, 217, 5),
    'F217': (f373, 217, 6),
    'J217': (f56, 217, 10),
    'K217': (f374, 217, 11),
    'L217': (f58, 217, 12),
    'E218': (f375, 218, 5),
    'F218': (f375, 218, 6),
    'J218': (f56, 218, 10),
    'K218': (f376, 218, 11),
    'L218': (f58, 218, 12),
    'E219': (f377, 219, 5),
    'F219': (f377, 219, 6),
    'J219': (f56, 219, 10),
    'K219': (f378, 219, 11),
    'L219': (f58, 219, 12),
    'E220': (f379, 220, 5),
    'F220': (f379, 220, 6),
    'J220': (f56, 220, 10),
    'K220': (f380, 220, 11),
    'L220': (f58, 220, 12),
    'E221': (f381, 221, 5),
    'F221': (f381, 221, 6),
    'J221': (f56, 221, 10),
    'K221': (f382, 221, 11),
    'L221': (f58, 221, 12),
    'E222': (f383, 222, 5),
    'F222': (f383, 222, 6),
    'J222': (f56, 222, 10),
    'K222': (f384, 222, 11),
    'L222': (f58, 222, 12),
    'E223': (f385, 223, 5),
    'F223': (f385, 223, 6),
    'J223': (f56, 223, 10),
    'K223': (f386, 223, 11),
    'L223': (f58, 223, 12),
    'E224': (f387, 224, 5),
    'F224': (f387, 224, 6),
    'J224': (f56, 224, 10),
    'K224': (f388, 224, 11),
    'L224': (f58, 224, 12),
    'E225': (f389, 225, 5),
    'F225': (f389, 225, 6),
    'J225': (f56, 225, 10),
    'K225': (f390, 225, 11),
    'L225': (f58, 225, 12),
    'E226': (f391, 226, 5),
    'F226': (f391, 226, 6),
    'J226': (f56, 226, 10),
    'K226': (f392, 226, 11),
    'L226': (f58, 226, 12),
    'E227': (f393, 227, 5),
    'F227': (f393, 227, 6),
    'J227': (f56, 227, 10),
    'K227': (f394, 227, 11),
    'L227': (f58, 227, 12),
    'E228': (f395, 228, 5),
    'F228': (f395, 228, 6),
    'J228': (f56, 228, 10),
    'K228': (f396, 228, 11),
    'L228': (f58, 228, 12),
    'E229': (f397, 229, 5),
    'F229': (f398, 229, 6),
    'J229': (f56, 229, 10),
    'K229': (f399, 229, 11),
    'L229': (f58, 229, 12),
    'E230': (f154, 230, 5),
    'F230': (f154, 230, 6),
    'J230': (f56, 230, 10),
    'K230': (f400, 230, 11),
    'L230': (f58, 230, 12),
    'E231': (f401, 231, 5),
    'F231': (f401, 231, 6),
    'J231': (f56, 231, 10),
    'K231': (f402, 231, 11),
    'L231': (f58, 231, 12),
    'E232': (f403, 232, 5),
    'F232': (f403, 232, 6),
    'J232': (f56, 232, 10),
    'K232': (f404, 232, 11),
    'L232': (f58, 232, 12),
    'E233': (f405, 233, 5),
    'F233': (f406, 233, 6),
    'J233': (f56, 233, 10),
    'K233': (f407, 233, 11),
    'L233': (f58, 233, 12),
    'E234': (f408, 234, 5),
    'F234': (f408, 234, 6),
    'J234': (f56, 234, 10),
    'K234': (f409, 234, 11),
    'L234': (f58, 234, 12),
    'E235': (f410, 235, 5),
    'F235': (f410, 235, 6),
    'J235': (f56, 235, 10),
    'K235': (f411, 235, 11),
    'L235': (f58, 235, 12),
    'E236': (f412, 236, 5),
    'F236': (f412, 236, 6),
    'J236': (f56, 236, 10),
    'K236': (f413, 236, 11),
    'L236': (f58, 236, 12),
    'E237': (f414, 237, 5),
    'F237': (f414, 237, 6),
    'J237': (f56, 237, 10),
    'K237': (f415, 237, 11),
    'L237': (f58, 237, 12),
    'E238': (f416, 238, 5),
    'F238': (f416, 238, 6),
    'J238': (f56, 238, 10),
    'K238': (f417, 238, 11),
    'L238': (f58, 238, 12),
    'E239': (f418, 239, 5),
    'F239': (f418, 239, 6),
    'J239': (f56, 239, 10),
    'K239': (f419, 239, 11),
    'L239': (f58, 239, 12),
    'E240': (f420, 240, 5),
    'F240': (f420, 240, 6),
    'J240': (f56, 240, 10),
    'K240': (f421, 240, 11),
    'L240': (f58, 240, 12),
    'E241': (f422, 241, 5),
    'F241': (f422, 241, 6),
    'J241': (f56, 241, 10),
    'K241': (f423, 241, 11),
    'L241': (f58, 241, 12),
    'E242': (f202, 242, 5),
    'F242': (f202, 242, 6),
    'J242': (f56, 242, 10),
    'K242': (f424, 242, 11),
    'L242': (f58, 242, 12),
    'E246': (f425, 246, 5),
    'F246': (f425, 246, 6),
    'J246': (f56, 246, 10),
    'K246': (f426, 246, 11),
    'L246': (f58, 246, 12),
    'E247': (f427, 247, 5),
    'F247': (f427, 247, 6),
    'J247': (f56, 247, 10),
    'K247': (f428, 247, 11),
    'L247': (f58, 247, 12),
    'E248': (f429, 248, 5),
    'F248': (f429, 248, 6),
    'J248': (f56, 248, 10),
    'K248': (f430, 248, 11),
    'L248': (f58, 248, 12),
    'E249': (f431, 249, 5),
    'F249': (f432, 249, 6),
    'J249': (f56, 249, 10),
    'K249': (f433, 249, 11),
    'L249': (f58, 249, 12),
    'E250': (f434, 250, 5),
    'F250': (f434, 250, 6),
    'J250': (f56, 250, 10),
    'K250': (f435, 250, 11),
    'L250': (f58, 250, 12),
    'E251': (f436, 251, 5),
    'F251': (f436, 251, 6),
    'J251': (f56, 251, 10),
    'K251': (f437, 251, 11),
    'L251': (f58, 251, 12),
    'E252': (f438, 252, 5),
    'F252': (f438, 252, 6),
    'J252': (f56, 252, 10),
    'K252': (f439, 252, 11),
    'L252': (f58, 252, 12),
    'E253': (f440, 253, 5),
    'F253': (f440, 253, 6),
    'J253': (f56, 253, 10),
    'K253': (f441, 253, 11),
    'L253': (f58, 253, 12),
    'E254': (f442, 254, 5),
    'F254': (f442, 254, 6),
    'J254': (f56, 254, 10),
    'K254': (f443, 254, 11),
    'L254': (f58, 254, 12),
    'E255': (f444, 255, 5),
    'F255': (f444, 255, 6),
    'J255': (f56, 255, 10),
    'K255': (f445, 255, 11),
    'L255': (f58, 255, 12),
    'E256': (f446, 256, 5),
    'F256': (f446, 256, 6),
    'J256': (f56, 256, 10),
    'K256': (f447, 256, 11),
    'L256': (f58, 256, 12),
    'E257': (f448, 257, 5),
    'F257': (f448, 257, 6),
    'J257': (f56, 257, 10),
    'K257': (f449, 257, 11),
    'L257': (f58, 257, 12),
    'E258': (f369, 258, 5),
    'F258': (f369, 258, 6),
    'J258': (f56, 258, 10),
    'K258': (f450, 258, 11),
    'L258': (f58, 258, 12),
    'E259': (f451, 259, 5),
    'F259': (f451, 259, 6),
    'J259': (f56, 259, 10),
    'K259': (f452, 259, 11),
    'L259': (f58, 259, 12),
    'E260': (f453, 260, 5),
    'F260': (f453, 260, 6),
    'J260': (f56, 260, 10),
    'K260': (f454, 260, 11),
    'L260': (f58, 260, 12),
    'E261': (f455, 261, 5),
    'F261': (f455, 261, 6),
    'J261': (f56, 261, 10),
    'K261': (f456, 261, 11),
    'L261': (f58, 261, 12),
    'E262': (f457, 262, 5),
    'F262': (f457, 262, 6),
    'J262': (f56, 262, 10),
    'K262': (f458, 262, 11),
    'L262': (f58, 262, 12),
    'E263': (f459, 263, 5),
    'F263': (f459, 263, 6),
    'J263': (f56, 263, 10),
    'K263': (f460, 263, 11),
    'L263': (f58, 263, 12),
    'E264': (f461, 264, 5),
    'F264': (f461, 264, 6),
    'J264': (f56, 264, 10),
    'K264': (f462, 264, 11),
    'L264': (f58, 264, 12),
    'E265': (f463, 265, 5),
    'F265': (f464, 265, 6),
    'J265': (f56, 265, 10),
    'K265': (f465, 265, 11),
    'L265': (f58, 265, 12),
    'E266': (f466, 266, 5),
    'F266': (f466, 266, 6),
    'J266': (f56, 266, 10),
    'K266': (f467, 266, 11),
    'L266': (f58, 266, 12),
    'E267': (f468, 267, 5),
    'F267': (f468, 267, 6),
    'J267': (f56, 267, 10),
    'K267': (f469, 267, 11),
    'L267': (f58, 267, 12),
    'E268': (f470, 268, 5),
    'F268': (f470, 268, 6),
    'J268': (f56, 268, 10),
    'K268': (f471, 268, 11),
    'L268': (f58, 268, 12),
    'E269': (f472, 269, 5),
    'F269': (f472, 269, 6),
    'J269': (f56, 269, 10),
    'K269': (f473, 269, 11),
    'L269': (f58, 269, 12),
    'E270': (f474, 270, 5),
    'F270': (f474, 270, 6),
    'J270': (f56, 270, 10),
    'K270': (f475, 270, 11),
    'L270': (f58, 270, 12),
    'E271': (f476, 271, 5),
    'F271': (f476, 271, 6),
    'J271': (f56, 271, 10),
    'K271': (f477, 271, 11),
    'L271': (f58, 271, 12),
    'E272': (f478, 272, 5),
    'F272': (f478, 272, 6),
    'J272': (f56, 272, 10),
    'K272': (f479, 272, 11),
    'L272': (f58, 272, 12),
    'E273': (f480, 273, 5),
    'F273': (f480, 273, 6),
    'J273': (f56, 273, 10),
    'K273': (f481, 273, 11),
    'L273': (f58, 273, 12),
    'E274': (f482, 274, 5),
    'F274': (f482, 274, 6),
    'J274': (f56, 274, 10),
    'K274': (f483, 274, 11),
    'L274': (f58, 274, 12),
    'E275': (f484, 275, 5),
    'F275': (f484, 275, 6),
    'J275': (f56, 275, 10),
    'K275': (f485, 275, 11),
    'L275': (f58, 275, 12),
    'E276': (f486, 276, 5),
    'F276': (f486, 276, 6),
    'J276': (f56, 276, 10),
    'K276': (f487, 276, 11),
    'L276': (f58, 276, 12),
    'E277': (f488, 277, 5),
    'F277': (f488, 277, 6),
    'J277': (f56, 277, 10),
    'K277': (f489, 277, 11),
    'L277': (f58, 277, 12),
    'E278': (f490, 278, 5),
    'F278': (f490, 278, 6),
    'J278': (f56, 278, 10),
    'K278': (f491, 278, 11),
    'L278': (f58, 278, 12),
    'E279': (f492, 279, 5),
    'F279': (f492, 279, 6),
    'J279': (f56, 279, 10),
    'K279': (f493, 279, 11),
    'L279': (f58, 279, 12),
    'E284': (f494, 284, 5),
    'F284': (f494, 284, 6),
    'J284': (f56, 284, 10),
    'K284': (f495, 284, 11),
    'L284': (f58, 284, 12),
    'E285': (f496, 285, 5),
    'F285': (f496, 285, 6),
    'J285': (f56, 285, 10),
    'K285': (f497, 285, 11),
    'L285': (f58, 285, 12),
    'E286': (f177, 286, 5),
    'F286': (f177, 286, 6),
    'J286': (f56, 286, 10),
    'K286': (f498, 286, 11),
    'L286': (f58, 286, 12),
    'E287': (f499, 287, 5),
    'F287': (f500, 287, 6),
    'J287': (f56, 287, 10),
    'K287': (f501, 287, 11),
    'L287': (f58, 287, 12),
    'E288': (f502, 288, 5),
    'F288': (f503, 288, 6),
    'J288': (f56, 288, 10),
    'K288': (f504, 288, 11),
    'L288': (f58, 288, 12),
    'E289': (f505, 289, 5),
    'F289': (f506, 289, 6),
    'J289': (f56, 289, 10),
    'K289': (f507, 289, 11),
    'L289': (f58, 289, 12),
    'E294': (f508, 294, 5),
    'F294': (f508, 294, 6),
    'J294': (f56, 294, 10),
    'K294': (f509, 294, 11),
    'L294': (f58, 294, 12),
    'E295': (f510, 295, 5),
    'F295': (f510, 295, 6),
    'J295': (f56, 295, 10),
    'K295': (f511, 295, 11),
    'L295': (f58, 295, 12),
    'E296': (f169, 296, 5),
    'F296': (f169, 296, 6),
    'J296': (f56, 296, 10),
    'K296': (f512, 296, 11),
    'L296': (f58, 296, 12),
    'E297': (f513, 297, 5),
    'F297': (f513, 297, 6),
    'J297': (f56, 297, 10),
    'K297': (f514, 297, 11),
    'L297': (f58, 297, 12),
    'E298': (f515, 298, 5),
    'F298': (f515, 298, 6),
    'J298': (f56, 298, 10),
    'K298': (f516, 298, 11),
    'L298': (f58, 298, 12),
    'E299': (f517, 299, 5),
    'F299': (f517, 299, 6),
    'J299': (f56, 299, 10),
    'K299': (f518, 299, 11),
    'L299': (f58, 299, 12),
    'E300': (f519, 300, 5),
    'F300': (f519, 300, 6),
    'J300': (f56, 300, 10),
    'K300': (f520, 300, 11),
    'L300': (f58, 300, 12),
    'E301': (f521, 301, 5),
    'F301': (f521, 301, 6),
    'J301': (f56, 301, 10),
    'K301': (f522, 301, 11),
    'L301': (f58, 301, 12),
    'E302': (f523, 302, 5),
    'F302': (f523, 302, 6),
    'J302': (f56, 302, 10),
    'K302': (f524, 302, 11),
    'L302': (f58, 302, 12),
    'E303': (f525, 303, 5),
    'F303': (f525, 303, 6),
    'J303': (f56, 303, 10),
    'K303': (f526, 303, 11),
    'L303': (f58, 303, 12),
    'E304': (f527, 304, 5),
    'F304': (f527, 304, 6),
    'J304': (f56, 304, 10),
    'K304': (f528, 304, 11),
    'L304': (f58, 304, 12),
    'E305': (f529, 305, 5),
    'F305': (f529, 305, 6),
    'J305': (f56, 305, 10),
    'K305': (f530, 305, 11),
    'L305': (f58, 305, 12),
    'E306': (f531, 306, 5),
    'F306': (f532, 306, 6),
    'J306': (f56, 306, 10),
    'K306': (f533, 306, 11),
    'L306': (f58, 306, 12),
    'E307': (f534, 307, 5),
    'F307': (f534, 307, 6),
    'J307': (f56, 307, 10),
    'K307': (f535, 307, 11),
    'L307': (f58, 307, 12),
    'E308': (f536, 308, 5),
    'F308': (f536, 308, 6),
    'J308': (f56, 308, 10),
    'K308': (f537, 308, 11),
    'L308': (f58, 308, 12),
    'E309': (f538, 309, 5),
    'F309': (f538, 309, 6),
    'J309': (f56, 309, 10),
    'K309': (f539, 309, 11),
    'L309': (f58, 309, 12),
    'E310': (f540, 310, 5),
    'F310': (f540, 310, 6),
    'J310': (f56, 310, 10),
    'K310': (f541, 310, 11),
    'L310': (f58, 310, 12),
    'E311': (f542, 311, 5),
    'F311': (f542, 311, 6),
    'J311': (f56, 311, 10),
    'K311': (f543, 311, 11),
    'L311': (f58, 311, 12),
    'D313': (f544, 313, 4),
}
