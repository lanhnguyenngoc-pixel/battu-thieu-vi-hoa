# -*- coding: utf-8 -*-
"""Module sheet KET_LUAN – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: 6.10_G.xlsx (SHA-256 a34b9bd88b9052807174d240d6d3d83136557f1d078a7e3a3ba2d95e1eadc33f).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'KET_LUAN'
SI = 3

def f0(E, r, c):
    # Ô B2
    # = IF(PP2_DINH_LUONG!$D$7<>"",PP2_DINH_LUONG!$D$7&". ","")&IF(LEFT(PP1_DINH_TINH!$C$112,4)="NEED","⚠ PP1 CẦN LUẬN GIẢI (P1 tr.71) – xem PP1_DINH_TINH C11
    #   1:C112. ","")&IF(DOI_CHIEU!$C$31<>DOI_CHIEU!$C$32,"⚠ PP1 và PP2 cho kết quả khác nhau – lựa chọn ở DOI_CHIEU là quy ước kỹ thuật (G11). ","")&IF(LEFT(
    #   NHAP_LIEU!$G$7,1)="⚠",NHAP_LIEU!$G$7&" ","")&IF((NHAP_LIEU!$C$9="Không rõ giờ"),"⚠ KHÔNG RÕ GIỜ SINH: mọi kết quả ở sheet này tính với giờ thử "&KHONG
    #   _RO_GIO!$C$6&". Kết luận đáng tin (chỉ phần giống nhau ở cả 12 giờ) và bảng tìm giờ sinh ở sheet KHONG_RO_GIO mục F, E.","")
    return X.cat(X.cat(X.cat(X.cat((X.cat(E.v(10,7,4),'. ') if X.truthy(X.ne(E.v(10,7,4),'')) else ''),('⚠ PP1 CẦN LUẬN GIẢI (P1 tr.71) – xem PP1_DINH_TINH C111:C112. ' if X.truthy(X.eq(X.LEFT(E.v(9,112,3), 4.0),'NEED')) else '')),('⚠ PP1 và PP2 cho kết quả khác nhau – lựa chọn ở DOI_CHIEU là quy ước kỹ thuật (G11). ' if X.truthy(X.ne(E.v(7,31,3),E.v(7,32,3))) else '')),(X.cat(E.v(2,7,7),' ') if X.truthy(X.eq(X.LEFT(E.v(2,7,7), 1.0),'⚠')) else '')),(X.cat(X.cat('⚠ KHÔNG RÕ GIỜ SINH: mọi kết quả ở sheet này tính với giờ thử ',E.v(6,6,3)),'. Kết luận đáng tin (chỉ phần giống nhau ở cả 12 giờ) và bảng tìm giờ sinh ở sheet KHONG_RO_GIO mục F, E.') if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else ''))

def f1(E, r, c):
    # Ô C3, D3, E3, F3, G3, D4, E4, F4, G4, D5, E5, F5 … (43 ô)
    # = ""
    return ''

def f2(E, r, c):
    # Ô C4
    # = C29
    return E.v(3,r+25,c)

def f3(E, r, c):
    # Ô C5
    # = C68
    return E.v(3,r+63,c)

def f4(E, r, c):
    # Ô C6
    # = C79
    return E.v(3,r+73,c)

def f5(E, r, c):
    # Ô C7
    # = C93
    return E.v(3,r+86,c)

def f6(E, r, c):
    # Ô C8
    # = C108
    return E.v(3,r+100,c)

def f7(E, r, c):
    # Ô C9
    # = C135
    return E.v(3,r+126,c)

def f8(E, r, c):
    # Ô C10
    # = C160
    return E.v(3,r+150,c)

def f9(E, r, c):
    # Ô C11
    # = C179
    return E.v(3,r+168,c)

def f10(E, r, c):
    # Ô C12
    # = C193
    return E.v(3,r+181,c)

def f11(E, r, c):
    # Ô C15, D15, E15, F15
    # = NHAP_LIEU!C23
    return E.v(2,r+8,c)

def f12(E, r, c):
    # Ô G15
    # = IF(NHAP_LIEU!$C$13="Tự nhập","Lá số tự nhập","Lá số mẫu: "&NHAP_LIEU!$C$13)
    return ('Lá số tự nhập' if X.truthy(X.eq(E.v(2,13,3),'Tự nhập')) else X.cat('Lá số mẫu: ',E.v(2,13,3)))

def f13(E, r, c):
    # Ô C17
    # = TRUNG_GIAN!$C$10&" – "&TRUNG_GIAN!$C$12
    return X.cat(X.cat(E.v(23,10,3),' – '),E.v(23,12,3))

def f14(E, r, c):
    # Ô D17
    # = TRUNG_GIAN!$D$10&" – "&TRUNG_GIAN!$D$12
    return X.cat(X.cat(E.v(23,10,4),' – '),E.v(23,12,4))

def f15(E, r, c):
    # Ô E17
    # = TRUNG_GIAN!$E$10&" – "&"NHẬT CHỦ"
    return X.cat(X.cat(E.v(23,10,5),' – '),'NHẬT CHỦ')

def f16(E, r, c):
    # Ô F17
    # = TRUNG_GIAN!$F$10&" – "&TRUNG_GIAN!$F$12
    return X.cat(X.cat(E.v(23,10,6),' – '),E.v(23,12,6))

def f17(E, r, c):
    # Ô C18
    # = INDEX(HANH_TEN,TRUNG_GIAN!$C$7+1)&IF(TRUNG_GIAN!$C$8=0," dương"," âm")
    return X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,3),1.0)),(' dương' if X.truthy(X.eq(E.v(23,8,3),0.0)) else ' âm'))

def f18(E, r, c):
    # Ô D18
    # = INDEX(HANH_TEN,TRUNG_GIAN!$D$7+1)&IF(TRUNG_GIAN!$D$8=0," dương"," âm")
    return X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,4),1.0)),(' dương' if X.truthy(X.eq(E.v(23,8,4),0.0)) else ' âm'))

def f19(E, r, c):
    # Ô E18
    # = INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&IF(TRUNG_GIAN!$E$8=0," dương"," âm")
    return X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0)),(' dương' if X.truthy(X.eq(E.v(23,8,5),0.0)) else ' âm'))

def f20(E, r, c):
    # Ô F18
    # = INDEX(HANH_TEN,TRUNG_GIAN!$F$7+1)&IF(TRUNG_GIAN!$F$8=0," dương"," âm")
    return X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,6),1.0)),(' dương' if X.truthy(X.eq(E.v(23,8,6),0.0)) else ' âm'))

def f21(E, r, c):
    # Ô C19
    # = TRUNG_GIAN!$C$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$C$9+1)&")"
    return X.cat(X.cat(X.cat(E.v(23,11,3),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,3),1.0))),')')

def f22(E, r, c):
    # Ô D19
    # = TRUNG_GIAN!$D$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$D$9+1)&")"
    return X.cat(X.cat(X.cat(E.v(23,11,4),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,4),1.0))),')')

def f23(E, r, c):
    # Ô E19
    # = TRUNG_GIAN!$E$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$9+1)&")"
    return X.cat(X.cat(X.cat(E.v(23,11,5),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,5),1.0))),')')

def f24(E, r, c):
    # Ô F19
    # = TRUNG_GIAN!$F$11&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$F$9+1)&")"
    return X.cat(X.cat(X.cat(E.v(23,11,6),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,9,6),1.0))),')')

def f25(E, r, c):
    # Ô C20
    # = IF(TRUNG_GIAN!$C$24="","—",TRUNG_GIAN!$C$24)
    return ('—' if X.truthy(X.eq(E.v(23,24,3),'')) else E.v(23,24,3))

def f26(E, r, c):
    # Ô D20
    # = IF(TRUNG_GIAN!$D$24="","—",TRUNG_GIAN!$D$24)
    return ('—' if X.truthy(X.eq(E.v(23,24,4),'')) else E.v(23,24,4))

def f27(E, r, c):
    # Ô E20
    # = IF(TRUNG_GIAN!$E$24="","—",TRUNG_GIAN!$E$24)
    return ('—' if X.truthy(X.eq(E.v(23,24,5),'')) else E.v(23,24,5))

def f28(E, r, c):
    # Ô F20
    # = IF(TRUNG_GIAN!$F$24="","—",TRUNG_GIAN!$F$24)
    return ('—' if X.truthy(X.eq(E.v(23,24,6),'')) else E.v(23,24,6))

def f29(E, r, c):
    # Ô C21
    # = IF(TRUNG_GIAN!$C$25="","—",TRUNG_GIAN!$C$25)
    return ('—' if X.truthy(X.eq(E.v(23,25,3),'')) else E.v(23,25,3))

def f30(E, r, c):
    # Ô D21
    # = IF(TRUNG_GIAN!$D$25="","—",TRUNG_GIAN!$D$25)
    return ('—' if X.truthy(X.eq(E.v(23,25,4),'')) else E.v(23,25,4))

def f31(E, r, c):
    # Ô E21
    # = IF(TRUNG_GIAN!$E$25="","—",TRUNG_GIAN!$E$25)
    return ('—' if X.truthy(X.eq(E.v(23,25,5),'')) else E.v(23,25,5))

def f32(E, r, c):
    # Ô F21
    # = IF(TRUNG_GIAN!$F$25="","—",TRUNG_GIAN!$F$25)
    return ('—' if X.truthy(X.eq(E.v(23,25,6),'')) else E.v(23,25,6))

def f33(E, r, c):
    # Ô C22
    # = IF(TRUNG_GIAN!$C$26="","—",TRUNG_GIAN!$C$26)
    return ('—' if X.truthy(X.eq(E.v(23,26,3),'')) else E.v(23,26,3))

def f34(E, r, c):
    # Ô D22
    # = IF(TRUNG_GIAN!$D$26="","—",TRUNG_GIAN!$D$26)
    return ('—' if X.truthy(X.eq(E.v(23,26,4),'')) else E.v(23,26,4))

def f35(E, r, c):
    # Ô E22
    # = IF(TRUNG_GIAN!$E$26="","—",TRUNG_GIAN!$E$26)
    return ('—' if X.truthy(X.eq(E.v(23,26,5),'')) else E.v(23,26,5))

def f36(E, r, c):
    # Ô F22
    # = IF(TRUNG_GIAN!$F$26="","—",TRUNG_GIAN!$F$26)
    return ('—' if X.truthy(X.eq(E.v(23,26,6),'')) else E.v(23,26,6))

def f37(E, r, c):
    # Ô C23
    # = TRUNG_GIAN!$E$10&" tại "&TRUNG_GIAN!$C$11&": "&TRUNG_GIAN!$C$29
    return X.cat(X.cat(X.cat(X.cat(E.v(23,10,5),' tại '),E.v(23,11,3)),': '),E.v(23,29,3))

def f38(E, r, c):
    # Ô D23
    # = TRUNG_GIAN!$E$10&" tại "&TRUNG_GIAN!$D$11&": "&TRUNG_GIAN!$D$29
    return X.cat(X.cat(X.cat(X.cat(E.v(23,10,5),' tại '),E.v(23,11,4)),': '),E.v(23,29,4))

def f39(E, r, c):
    # Ô E23
    # = TRUNG_GIAN!$E$10&" tại "&TRUNG_GIAN!$E$11&": "&TRUNG_GIAN!$E$29
    return X.cat(X.cat(X.cat(X.cat(E.v(23,10,5),' tại '),E.v(23,11,5)),': '),E.v(23,29,5))

def f40(E, r, c):
    # Ô F23
    # = TRUNG_GIAN!$E$10&" tại "&TRUNG_GIAN!$F$11&": "&TRUNG_GIAN!$F$29
    return X.cat(X.cat(X.cat(X.cat(E.v(23,10,5),' tại '),E.v(23,11,6)),': '),E.v(23,29,6))

def f41(E, r, c):
    # Ô C24
    # = "Tỉ kiếp "&O24&" chữ"&IF(MOD(TRUNG_GIAN!$E$7+0,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (DỤNG)",IF(MOD(TRUNG_GIAN!$E$7+0,5)=IF(D
    #   OI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)," (HỶ)",IF(MOD(TRUNG_GIAN!$E$7+0,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," 
    #   (KỴ)","")))&IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$92,DUNG_HY_KY!$C$92)=0,MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_
    #   KY!$C$109),MOD(TRUNG_GIAN!$E$7+0,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))," (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)",
    #   "")&IF(O24<>PP1_DINH_TINH!$F$96," [PP1 sau hợp hóa đếm "&PP1_DINH_TINH!$F$96&"]","")&": "&IF(J24="","—",LEFT(J24,LEN(J24)-1))&CHAR(10)&"Thực thương "&
    #   P24&" chữ"&IF(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (DỤNG)",IF(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C
    #   $33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)," (HỶ)",IF(MOD(TRUNG_GIAN!$E$7+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," (KỴ)","")))
    #   &IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$92,DUNG_HY_KY!$C$92)=1,MOD(TRUNG_GIAN!$E$7+1,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),
    #   MOD(TRUNG_GIAN!$E$7+1,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))," (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)","")&IF(P24<
    #   >PP1_DINH_TINH!$F$97," [PP1 sau hợp hóa đếm "&PP1_DINH_TINH!$F$97&"]","")&": "&IF(K24="","—",LEFT(K24,LEN(K24)-1))&CHAR(10)&"Tài "&Q24&" chữ"&IF(MOD(T
    #   RUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (DỤNG)",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D
    #   $110,DUNG_HY_KY!$C$110)," (HỶ)",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," (KỴ)","")))&IF(AND(IF(DOI_CHIE
    #   U!$C$33=2,DUNG_HY_KY!$D$92,DUNG_HY_KY!$C$92)=2,MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7
    #   +2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))," (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)","")&IF(Q24<>PP1_DINH_TINH!$F$9
    #   8," [PP1 sau hợp hóa đếm "&PP1_DINH_TINH!$F$98&"]","")&": "&IF(L24="","—",LEFT(L24,LEN(L24)-1))&CHAR(10)&"Quan sát "&R24&" chữ"&IF(MOD(TRUNG_GIAN!$E$7
    #   +3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (DỤNG)",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_K
    #   Y!$C$110)," (HỶ)",IF(MOD(TRUNG_GIAN!$E$7+3,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," (KỴ)","")))&IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG
    #   _HY_KY!$D$92,DUNG_HY_KY!$C$92)=3,MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+3,5)<>IF(DOI_
    #   CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110))," (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)","")&IF(R24<>PP1_DINH_TINH!$F$99," [PP1 sau h
    #   ợp hóa đếm "&PP1_DINH_TINH!$F$99&"]","")&": "&IF(M24="","—",LEFT(M24,LEN(M24)-1))&CHAR(10)&"Ấn "&S24&" chữ"&IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$
    #   C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)," (DỤNG)",IF(MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)," (HỶ)",IF
    #   (MOD(TRUNG_GIAN!$E$7+4,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," (KỴ)","")))&IF(AND(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$92,DUNG_HY
    #   _KY!$C$92)=4,MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+4,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$110,DUNG_HY_KY!$C$110))," (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)","")&IF(S24<>PP1_DINH_TINH!$F$100," [PP1 sau hợp hóa đếm "&PP1_DI
    #   NH_TINH!$F$100&"]","")&": "&IF(N24="","—",LEFT(N24,LEN(N24)-1))&CHAR(10)&"Tổng "&(O24+P24+Q24+R24+S24)&" chữ = 3 can lộ + số tàng can. Vai trò theo ph
    #   ương pháp khuyến nghị (DOI_CHIEU C33)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tỉ kiếp ',E.v(3,r,c+12)),' chữ'),(' (DỤNG)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else (' (HỶ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else (' (KỴ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),(' (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)' if X.truthy(X.AND(X.eq((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),0.0), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),0.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat(X.cat(' [PP1 sau hợp hóa đếm ',E.v(9,96,6)),']') if X.truthy(X.ne(E.v(3,r,c+12),E.v(9,96,6))) else '')),': '),('—' if X.truthy(X.eq(E.v(3,r,c+7),'')) else X.LEFT(E.v(3,r,c+7), X.sub(X.LEN(E.v(3,r,c+7)),1.0)))),X.CHAR(10.0)),'Thực thương '),E.v(3,r,c+13)),' chữ'),(' (DỤNG)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else (' (HỶ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else (' (KỴ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),(' (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)' if X.truthy(X.AND(X.eq((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),1.0), X.ne(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),1.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat(X.cat(' [PP1 sau hợp hóa đếm ',E.v(9,97,6)),']') if X.truthy(X.ne(E.v(3,r,c+13),E.v(9,97,6))) else '')),': '),('—' if X.truthy(X.eq(E.v(3,r,c+8),'')) else X.LEFT(E.v(3,r,c+8), X.sub(X.LEN(E.v(3,r,c+8)),1.0)))),X.CHAR(10.0)),'Tài '),E.v(3,r,c+14)),' chữ'),(' (DỤNG)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else (' (HỶ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else (' (KỴ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),(' (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)' if X.truthy(X.AND(X.eq((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),2.0), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat(X.cat(' [PP1 sau hợp hóa đếm ',E.v(9,98,6)),']') if X.truthy(X.ne(E.v(3,r,c+14),E.v(9,98,6))) else '')),': '),('—' if X.truthy(X.eq(E.v(3,r,c+9),'')) else X.LEFT(E.v(3,r,c+9), X.sub(X.LEN(E.v(3,r,c+9)),1.0)))),X.CHAR(10.0)),'Quan sát '),E.v(3,r,c+15)),' chữ'),(' (DỤNG)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else (' (HỶ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else (' (KỴ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),(' (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)' if X.truthy(X.AND(X.eq((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),3.0), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),3.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat(X.cat(' [PP1 sau hợp hóa đếm ',E.v(9,99,6)),']') if X.truthy(X.ne(E.v(3,r,c+15),E.v(9,99,6))) else '')),': '),('—' if X.truthy(X.eq(E.v(3,r,c+10),'')) else X.LEFT(E.v(3,r,c+10), X.sub(X.LEN(E.v(3,r,c+10)),1.0)))),X.CHAR(10.0)),'Ấn '),E.v(3,r,c+16)),' chữ'),(' (DỤNG)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else (' (HỶ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else (' (KỴ)' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),(' (nhóm nhiều – cũng là kỵ theo cân bằng, P1 tr.180)' if X.truthy(X.AND(X.eq((E.v(8,92,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,92,3)),4.0), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),4.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else '')),(X.cat(X.cat(' [PP1 sau hợp hóa đếm ',E.v(9,100,6)),']') if X.truthy(X.ne(E.v(3,r,c+16),E.v(9,100,6))) else '')),': '),('—' if X.truthy(X.eq(E.v(3,r,c+11),'')) else X.LEFT(E.v(3,r,c+11), X.sub(X.LEN(E.v(3,r,c+11)),1.0)))),X.CHAR(10.0)),'Tổng '),X.add(X.add(X.add(X.add(E.v(3,r,c+12),E.v(3,r,c+13)),E.v(3,r,c+14)),E.v(3,r,c+15)),E.v(3,r,c+16))),' chữ = 3 can lộ + số tàng can. Vai trò theo phương pháp khuyến nghị (DOI_CHIEU C33).')

def f42(E, r, c):
    # Ô J24
    # = TRIM(IF(MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=0,TRUNG_GIAN!$C$10&" (can Năm), ","")&IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5
    #   )=0),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" (tàng bản khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$22>=0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)
    #   =0),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" (tàng trung khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5
    #   )=0),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" (tàng dư khí "&TRUNG_GIAN!$C$11&"), ","")&IF(MOD(TRUNG_GIAN!$D$7-TRUNG_GIAN!$E$7,5)=0,TRUNG_GIAN!$D$10&" (can
    #    Tháng), ","")&IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" (tàng bản khí "&TRUNG_GIAN!$D
    #   $11&"), ","")&IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-TRUNG_GIAN!$E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" (tàng trung khí "&TRUNG_GIAN!$
    #   D$11&"), ","")&IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" (tàng dư khí "&TRUNG_GIAN!$D$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-TRUNG_GIAN!$E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" (tàng bản khí "&TRUNG_GIAN!$E$1
    #   1&"), ","")&IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" (tàng trung khí "&TRUNG_GIAN!$E$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-TRUNG_GIAN!$E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" (tàng dư khí "&TRUNG_GIAN!$E$11
    #   &"), ","")&IF(MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=0,TRUNG_GIAN!$F$10&" (can Giờ), ","")&IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!
    #   $E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" (tàng bản khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$
    #   E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" (tàng trung khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!
    #   $E$7,5)=0),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" (tàng dư khí "&TRUNG_GIAN!$F$11&"), ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(23,10,3),' (can Năm), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),0.0)) else ''),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)),' (tàng bản khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)),' (tàng trung khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)),' (tàng dư khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(E.v(23,10,4),' (can Tháng), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),0.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)),' (tàng bản khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)),' (tàng trung khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)),' (tàng dư khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)),' (tàng bản khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)),' (tàng trung khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)),' (tàng dư khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(E.v(23,10,6),' (can Giờ), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),0.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)),' (tàng bản khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)),' (tàng trung khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),0.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)),' (tàng dư khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),0.0))) else '')))

def f43(E, r, c):
    # Ô K24
    # = TRIM(IF(MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=1,TRUNG_GIAN!$C$10&" (can Năm), ","")&IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5
    #   )=1),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" (tàng bản khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$22>=0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)
    #   =1),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" (tàng trung khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5
    #   )=1),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" (tàng dư khí "&TRUNG_GIAN!$C$11&"), ","")&IF(MOD(TRUNG_GIAN!$D$7-TRUNG_GIAN!$E$7,5)=1,TRUNG_GIAN!$D$10&" (can
    #    Tháng), ","")&IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" (tàng bản khí "&TRUNG_GIAN!$D
    #   $11&"), ","")&IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-TRUNG_GIAN!$E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" (tàng trung khí "&TRUNG_GIAN!$
    #   D$11&"), ","")&IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" (tàng dư khí "&TRUNG_GIAN!$D$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-TRUNG_GIAN!$E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" (tàng bản khí "&TRUNG_GIAN!$E$1
    #   1&"), ","")&IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" (tàng trung khí "&TRUNG_GIAN!$E$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-TRUNG_GIAN!$E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" (tàng dư khí "&TRUNG_GIAN!$E$11
    #   &"), ","")&IF(MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=1,TRUNG_GIAN!$F$10&" (can Giờ), ","")&IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!
    #   $E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" (tàng bản khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$
    #   E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" (tàng trung khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!
    #   $E$7,5)=1),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" (tàng dư khí "&TRUNG_GIAN!$F$11&"), ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(23,10,3),' (can Năm), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),1.0)) else ''),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)),' (tàng bản khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)),' (tàng trung khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)),' (tàng dư khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(E.v(23,10,4),' (can Tháng), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),1.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)),' (tàng bản khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)),' (tàng trung khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)),' (tàng dư khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)),' (tàng bản khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)),' (tàng trung khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)),' (tàng dư khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(E.v(23,10,6),' (can Giờ), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),1.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)),' (tàng bản khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)),' (tàng trung khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),1.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)),' (tàng dư khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),1.0))) else '')))

def f44(E, r, c):
    # Ô L24
    # = TRIM(IF(MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=2,TRUNG_GIAN!$C$10&" (can Năm), ","")&IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5
    #   )=2),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" (tàng bản khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$22>=0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)
    #   =2),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" (tàng trung khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5
    #   )=2),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" (tàng dư khí "&TRUNG_GIAN!$C$11&"), ","")&IF(MOD(TRUNG_GIAN!$D$7-TRUNG_GIAN!$E$7,5)=2,TRUNG_GIAN!$D$10&" (can
    #    Tháng), ","")&IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" (tàng bản khí "&TRUNG_GIAN!$D
    #   $11&"), ","")&IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-TRUNG_GIAN!$E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" (tàng trung khí "&TRUNG_GIAN!$
    #   D$11&"), ","")&IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" (tàng dư khí "&TRUNG_GIAN!$D$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-TRUNG_GIAN!$E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" (tàng bản khí "&TRUNG_GIAN!$E$1
    #   1&"), ","")&IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" (tàng trung khí "&TRUNG_GIAN!$E$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-TRUNG_GIAN!$E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" (tàng dư khí "&TRUNG_GIAN!$E$11
    #   &"), ","")&IF(MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=2,TRUNG_GIAN!$F$10&" (can Giờ), ","")&IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!
    #   $E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" (tàng bản khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$
    #   E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" (tàng trung khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!
    #   $E$7,5)=2),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" (tàng dư khí "&TRUNG_GIAN!$F$11&"), ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(23,10,3),' (can Năm), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),2.0)) else ''),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)),' (tàng bản khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)),' (tàng trung khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)),' (tàng dư khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(E.v(23,10,4),' (can Tháng), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),2.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)),' (tàng bản khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)),' (tàng trung khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)),' (tàng dư khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)),' (tàng bản khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)),' (tàng trung khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)),' (tàng dư khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(E.v(23,10,6),' (can Giờ), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),2.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)),' (tàng bản khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)),' (tàng trung khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),2.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)),' (tàng dư khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),2.0))) else '')))

def f45(E, r, c):
    # Ô M24
    # = TRIM(IF(MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=3,TRUNG_GIAN!$C$10&" (can Năm), ","")&IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5
    #   )=3),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" (tàng bản khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$22>=0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)
    #   =3),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" (tàng trung khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5
    #   )=3),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" (tàng dư khí "&TRUNG_GIAN!$C$11&"), ","")&IF(MOD(TRUNG_GIAN!$D$7-TRUNG_GIAN!$E$7,5)=3,TRUNG_GIAN!$D$10&" (can
    #    Tháng), ","")&IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" (tàng bản khí "&TRUNG_GIAN!$D
    #   $11&"), ","")&IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-TRUNG_GIAN!$E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" (tàng trung khí "&TRUNG_GIAN!$
    #   D$11&"), ","")&IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" (tàng dư khí "&TRUNG_GIAN!$D$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-TRUNG_GIAN!$E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" (tàng bản khí "&TRUNG_GIAN!$E$1
    #   1&"), ","")&IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" (tàng trung khí "&TRUNG_GIAN!$E$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-TRUNG_GIAN!$E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" (tàng dư khí "&TRUNG_GIAN!$E$11
    #   &"), ","")&IF(MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=3,TRUNG_GIAN!$F$10&" (can Giờ), ","")&IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!
    #   $E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" (tàng bản khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$
    #   E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" (tàng trung khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!
    #   $E$7,5)=3),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" (tàng dư khí "&TRUNG_GIAN!$F$11&"), ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(23,10,3),' (can Năm), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),3.0)) else ''),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)),' (tàng bản khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)),' (tàng trung khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)),' (tàng dư khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(E.v(23,10,4),' (can Tháng), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),3.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)),' (tàng bản khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)),' (tàng trung khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)),' (tàng dư khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)),' (tàng bản khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)),' (tàng trung khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)),' (tàng dư khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(E.v(23,10,6),' (can Giờ), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),3.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)),' (tàng bản khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)),' (tàng trung khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),3.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)),' (tàng dư khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),3.0))) else '')))

def f46(E, r, c):
    # Ô N24
    # = TRIM(IF(MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=4,TRUNG_GIAN!$C$10&" (can Năm), ","")&IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5
    #   )=4),INDEX(CAN_TEN,TRUNG_GIAN!$C$15+1)&" (tàng bản khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$22>=0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)
    #   =4),INDEX(CAN_TEN,TRUNG_GIAN!$C$17+1)&" (tàng trung khí "&TRUNG_GIAN!$C$11&"), ","")&IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5
    #   )=4),INDEX(CAN_TEN,TRUNG_GIAN!$C$19+1)&" (tàng dư khí "&TRUNG_GIAN!$C$11&"), ","")&IF(MOD(TRUNG_GIAN!$D$7-TRUNG_GIAN!$E$7,5)=4,TRUNG_GIAN!$D$10&" (can
    #    Tháng), ","")&IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$D$15+1)&" (tàng bản khí "&TRUNG_GIAN!$D
    #   $11&"), ","")&IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-TRUNG_GIAN!$E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$D$17+1)&" (tàng trung khí "&TRUNG_GIAN!$
    #   D$11&"), ","")&IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$D$19+1)&" (tàng dư khí "&TRUNG_GIAN!$D$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-TRUNG_GIAN!$E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$E$15+1)&" (tàng bản khí "&TRUNG_GIAN!$E$1
    #   1&"), ","")&IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$E$17+1)&" (tàng trung khí "&TRUNG_GIAN!$E$
    #   11&"), ","")&IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-TRUNG_GIAN!$E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$E$19+1)&" (tàng dư khí "&TRUNG_GIAN!$E$11
    #   &"), ","")&IF(MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=4,TRUNG_GIAN!$F$10&" (can Giờ), ","")&IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!
    #   $E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$F$15+1)&" (tàng bản khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$
    #   E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$F$17+1)&" (tàng trung khí "&TRUNG_GIAN!$F$11&"), ","")&IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!
    #   $E$7,5)=4),INDEX(CAN_TEN,TRUNG_GIAN!$F$19+1)&" (tàng dư khí "&TRUNG_GIAN!$F$11&"), ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(23,10,3),' (can Năm), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),4.0)) else ''),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,3),1.0)),' (tàng bản khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,3),1.0)),' (tàng trung khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,3),1.0)),' (tàng dư khí '),E.v(23,11,3)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(E.v(23,10,4),' (can Tháng), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),4.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,4),1.0)),' (tàng bản khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,4),1.0)),' (tàng trung khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,4),1.0)),' (tàng dư khí '),E.v(23,11,4)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,5),1.0)),' (tàng bản khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,5),1.0)),' (tàng trung khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,5),1.0)),' (tàng dư khí '),E.v(23,11,5)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(E.v(23,10,6),' (can Giờ), ') if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),4.0)) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,15,6),1.0)),' (tàng bản khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,17,6),1.0)),' (tàng trung khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),4.0))) else '')),(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,19,6),1.0)),' (tàng dư khí '),E.v(23,11,6)),'), ') if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),4.0))) else '')))

def f47(E, r, c):
    # Ô O24
    # = IF((MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$C$22>
    #   =0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5)=0),1,0)+IF((MOD(TRUNG_GIAN!$D$7-
    #   TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-
    #   TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-
    #   TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-
    #   TRUNG_GIAN!$E$7,5)=0),1,0)+IF((MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!$E$7,5)=0),1,
    #   0)+IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$E$7,5)=0),1,0)+IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!$E$7,5)=0),1,
    #   0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),0.0)) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),0.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),0.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),0.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),0.0))) else 0.0))

def f48(E, r, c):
    # Ô P24
    # = IF((MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$C$22>
    #   =0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5)=1),1,0)+IF((MOD(TRUNG_GIAN!$D$7-
    #   TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-
    #   TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-
    #   TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-
    #   TRUNG_GIAN!$E$7,5)=1),1,0)+IF((MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!$E$7,5)=1),1,
    #   0)+IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$E$7,5)=1),1,0)+IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!$E$7,5)=1),1,
    #   0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),1.0)) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),1.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),1.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),1.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),1.0))) else 0.0))

def f49(E, r, c):
    # Ô Q24
    # = IF((MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$C$22>
    #   =0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5)=2),1,0)+IF((MOD(TRUNG_GIAN!$D$7-
    #   TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-
    #   TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-
    #   TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-
    #   TRUNG_GIAN!$E$7,5)=2),1,0)+IF((MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!$E$7,5)=2),1,
    #   0)+IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$E$7,5)=2),1,0)+IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!$E$7,5)=2),1,
    #   0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),2.0)) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),2.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),2.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),2.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),2.0))) else 0.0))

def f50(E, r, c):
    # Ô R24
    # = IF((MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$C$22>
    #   =0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5)=3),1,0)+IF((MOD(TRUNG_GIAN!$D$7-
    #   TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-
    #   TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-
    #   TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-
    #   TRUNG_GIAN!$E$7,5)=3),1,0)+IF((MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!$E$7,5)=3),1,
    #   0)+IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$E$7,5)=3),1,0)+IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!$E$7,5)=3),1,
    #   0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),3.0)) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),3.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),3.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),3.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),3.0))) else 0.0))

def f51(E, r, c):
    # Ô S24
    # = IF((MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$C$21>=0,MOD(TRUNG_GIAN!$C$21-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$C$22>
    #   =0,MOD(TRUNG_GIAN!$C$22-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$C$23>=0,MOD(TRUNG_GIAN!$C$23-TRUNG_GIAN!$E$7,5)=4),1,0)+IF((MOD(TRUNG_GIAN!$D$7-
    #   TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$D$21>=0,MOD(TRUNG_GIAN!$D$21-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$D$22>=0,MOD(TRUNG_GIAN!$D$22-
    #   TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$D$23>=0,MOD(TRUNG_GIAN!$D$23-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$E$21>=0,MOD(TRUNG_GIAN!$E$21-
    #   TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$E$22>=0,MOD(TRUNG_GIAN!$E$22-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$E$23>=0,MOD(TRUNG_GIAN!$E$23-
    #   TRUNG_GIAN!$E$7,5)=4),1,0)+IF((MOD(TRUNG_GIAN!$F$7-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$F$21>=0,MOD(TRUNG_GIAN!$F$21-TRUNG_GIAN!$E$7,5)=4),1,
    #   0)+IF(AND(TRUNG_GIAN!$F$22>=0,MOD(TRUNG_GIAN!$F$22-TRUNG_GIAN!$E$7,5)=4),1,0)+IF(AND(TRUNG_GIAN!$F$23>=0,MOD(TRUNG_GIAN!$F$23-TRUNG_GIAN!$E$7,5)=4),1,
    #   0)
    return X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,3),E.v(23,7,5)), 5.0),4.0)) else 0.0),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,3),0.0), X.eq(X.MOD(X.sub(E.v(23,21,3),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,3),0.0), X.eq(X.MOD(X.sub(E.v(23,22,3),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,3),0.0), X.eq(X.MOD(X.sub(E.v(23,23,3),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,4),E.v(23,7,5)), 5.0),4.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,4),0.0), X.eq(X.MOD(X.sub(E.v(23,21,4),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,4),0.0), X.eq(X.MOD(X.sub(E.v(23,22,4),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,4),0.0), X.eq(X.MOD(X.sub(E.v(23,23,4),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,5),0.0), X.eq(X.MOD(X.sub(E.v(23,21,5),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,5),0.0), X.eq(X.MOD(X.sub(E.v(23,22,5),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,5),0.0), X.eq(X.MOD(X.sub(E.v(23,23,5),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(23,7,6),E.v(23,7,5)), 5.0),4.0)) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,21,6),0.0), X.eq(X.MOD(X.sub(E.v(23,21,6),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,22,6),0.0), X.eq(X.MOD(X.sub(E.v(23,22,6),E.v(23,7,5)), 5.0),4.0))) else 0.0)),(1.0 if X.truthy(X.AND(X.ge(E.v(23,23,6),0.0), X.eq(X.MOD(X.sub(E.v(23,23,6),E.v(23,7,5)), 5.0),4.0))) else 0.0))

def f52(E, r, c):
    # Ô C29
    # = ("NHẬT CHỦ "&INDEX(CAN_TEN,TRUNG_GIAN!$E$5+1)&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&") "&DOI_CHIEU!$C$34&" – độ tin cậy "&UPPER(DOI_CHIEU!$C$35)&". L
    #   ý do: PP1 "&DUNG_HY_KY!$C$6&" ("&PP1_DINH_TINH!$C$92&"); PP2 phe mình "&TEXT(PP2_DINH_LUONG!$C$101,"0.0%")&" (vượng khi ≥ 40%). "&DUNG_HY_KY!$C$179&DU
    #   NG_HY_KY!$D$190)&CHAR(10)&(IF(DOI_CHIEU!$C$33=2,"DỤNG "&DUNG_HY_KY!$D$106&" · HỶ "&DUNG_HY_KY!$D$107&" · KỴ "&DUNG_HY_KY!$D$108&" (theo PP2)","DỤNG "&
    #   DUNG_HY_KY!$C$106&" · HỶ "&DUNG_HY_KY!$C$107&" · KỴ "&DUNG_HY_KY!$C$108&IF(DOI_CHIEU!$C$33=0," (hai phương pháp thống nhất dụng thần)"," (theo PP1)"))
    #   &IF(DUNG_HY_KY!$C$109<>DUNG_HY_KY!$D$109," ⚠ Hai phương pháp khác dụng thần (PP1 "&DUNG_HY_KY!$C$106&", PP2 "&DUNG_HY_KY!$D$106&") – dùng theo khuyến 
    #   nghị, độ tin "&LOWER(DOI_CHIEU!$C$35)&".","")&" Cứu ứng: "&DUNG_HY_KY!$C$174)&CHAR(10)&"Cách cục: "&C35&". Mệnh cục (P1 tr.173): "&LEFT(C45,IFERROR(FI
    #   ND(CHAR(10),C45)-1,LEN(C45)))&CHAR(10)&"Dụng thần bổ trợ: "&DUNG_HY_KY!$C$184
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('NHẬT CHỦ ',X.INDEX(E.rg(24,5,1,14,1), X.add(E.v(23,5,5),1.0))),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),') '),E.v(7,34,3)),' – độ tin cậy '),X.UPPER(E.v(7,35,3))),'. Lý do: PP1 '),E.v(8,6,3)),' ('),E.v(9,92,3)),'); PP2 phe mình '),X.TEXT(E.v(10,101,3), '0.0%')),' (vượng khi ≥ 40%). '),E.v(8,179,3)),E.v(8,190,4)),X.CHAR(10.0)),X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('DỤNG ',E.v(8,106,4)),' · HỶ '),E.v(8,107,4)),' · KỴ '),E.v(8,108,4)),' (theo PP2)') if X.truthy(X.eq(E.v(7,33,3),2.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('DỤNG ',E.v(8,106,3)),' · HỶ '),E.v(8,107,3)),' · KỴ '),E.v(8,108,3)),(' (hai phương pháp thống nhất dụng thần)' if X.truthy(X.eq(E.v(7,33,3),0.0)) else ' (theo PP1)'))),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' ⚠ Hai phương pháp khác dụng thần (PP1 ',E.v(8,106,3)),', PP2 '),E.v(8,106,4)),') – dùng theo khuyến nghị, độ tin '),X.LOWER(E.v(7,35,3))),'.') if X.truthy(X.ne(E.v(8,109,3),E.v(8,109,4))) else '')),' Cứu ứng: '),E.v(8,174,3))),X.CHAR(10.0)),'Cách cục: '),E.v(3,r+6,c)),'. Mệnh cục (P1 tr.173): '),X.LEFT(E.v(3,r+16,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+16,c)),1.0), lambda: X.LEN(E.v(3,r+16,c))))),X.CHAR(10.0)),'Dụng thần bổ trợ: '),E.v(8,184,3))

def f53(E, r, c):
    # Ô C31
    # = DUNG_HY_KY!$C$6
    return E.v(8,6,3)

def f54(E, r, c):
    # Ô D31
    # = DUNG_HY_KY!$D$6
    return E.v(8,6,4)

def f55(E, r, c):
    # Ô E31
    # = DOI_CHIEU!$E$15
    return E.v(7,15,5)

def f56(E, r, c):
    # Ô F31
    # = "PP1: "&PP1_DINH_TINH!$C$92&" | PP2: phe mình "&TEXT(PP2_DINH_LUONG!$C$99,"0.0")&"/"&TEXT(PP2_DINH_LUONG!$C$100,"0.0")&" độ = "&TEXT(PP2_DINH_LUONG!$C
    #   $101,"0.0%")&" (vượng khi ≥40%, P2 tr.352)"&" → Diễn giải từng chữ, từng tiêu chí: dòng 17-19 ngay dưới."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(9,92,3)),' | PP2: phe mình '),X.TEXT(E.v(10,99,3), '0.0')),'/'),X.TEXT(E.v(10,100,3), '0.0')),' độ = '),X.TEXT(E.v(10,101,3), '0.0%')),' (vượng khi ≥40%, P2 tr.352)'),' → Diễn giải từng chữ, từng tiêu chí: dòng 17-19 ngay dưới.')

def f57(E, r, c):
    # Ô G31
    # = "Hành nhật chủ: "&INDEX(HANH_TEN,DUNG_HY_KY!$C$5+1)&IF(DUNG_HY_KY!$C$5<>DUNG_HY_KY!$D$5," (PP2 theo hóa thần: "&INDEX(HANH_TEN,DUNG_HY_KY!$D$5+1)&")",
    #   "")
    return X.cat(X.cat('Hành nhật chủ: ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,5,3),1.0))),(X.cat(X.cat(' (PP2 theo hóa thần: ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,5,4),1.0))),')') if X.truthy(X.ne(E.v(8,5,3),E.v(8,5,4))) else ''))

def f58(E, r, c):
    # Ô H31
    # = LUAN_GIAI!$J$144
    return E.v(4,144,10)

def f59(E, r, c):
    # Ô C32
    # = "① "&PP1_DINH_TINH!$C$104&CHAR(10)&"② "&PP1_DINH_TINH!$C$105&CHAR(10)&"   ↳ Diễn giải đắc địa (khác thông căn): "&PP1_DINH_TINH!$C$47&CHAR(10)&"③ "&PP
    #   1_DINH_TINH!$C$106&CHAR(10)&"④ "&PP1_DINH_TINH!$C$107&CHAR(10)&"⇒ "&PP1_DINH_TINH!$C$108&CHAR(10)&PP1_DINH_TINH!$C$109&IF(PP1_DINH_TINH!$C$111<>"",CHA
    #   R(10)&PP1_DINH_TINH!$C$111,"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('① ',E.v(9,104,3)),X.CHAR(10.0)),'② '),E.v(9,105,3)),X.CHAR(10.0)),'   ↳ Diễn giải đắc địa (khác thông căn): '),E.v(9,47,3)),X.CHAR(10.0)),'③ '),E.v(9,106,3)),X.CHAR(10.0)),'④ '),E.v(9,107,3)),X.CHAR(10.0)),'⇒ '),E.v(9,108,3)),X.CHAR(10.0)),E.v(9,109,3)),(X.cat(X.CHAR(10.0),E.v(9,111,3)) if X.truthy(X.ne(E.v(9,111,3),'')) else ''))

def f60(E, r, c):
    # Ô H32
    # = LUAN_GIAI!$J$145
    return E.v(4,145,10)

def f61(E, r, c):
    # Ô C33
    # = PP2_DINH_LUONG!$C$114&CHAR(10)&PP2_DINH_LUONG!$C$115&CHAR(10)&PP2_DINH_LUONG!$C$116&CHAR(10)&PP2_DINH_LUONG!$C$117
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(10,114,3),X.CHAR(10.0)),E.v(10,115,3)),X.CHAR(10.0)),E.v(10,116,3)),X.CHAR(10.0)),E.v(10,117,3))

def f62(E, r, c):
    # Ô H33
    # = LUAN_GIAI!$J$146
    return E.v(4,146,10)

def f63(E, r, c):
    # Ô C34
    # = "PP1: "&PP1_DINH_TINH!$C$110&CHAR(10)&"PP2: "&PP2_DINH_LUONG!$C$118&CHAR(10)&"⇒ "&DOI_CHIEU!$C$49&" (khi hai phương pháp khác nhau: phân tích nguyên n
    #   hân, khuyến nghị và cách tự kiểm chứng ở DOI_CHIEU dòng 41-49)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(9,110,3)),X.CHAR(10.0)),'PP2: '),E.v(10,118,3)),X.CHAR(10.0)),'⇒ '),E.v(7,49,3)),' (khi hai phương pháp khác nhau: phân tích nguyên nhân, khuyến nghị và cách tự kiểm chứng ở DOI_CHIEU dòng 41-49).')

def f64(E, r, c):
    # Ô H34
    # = LUAN_GIAI!$J$106
    return E.v(4,106,10)

def f65(E, r, c):
    # Ô C35
    # = IF(AND(DUNG_HY_KY!$D$72="",DUNG_HY_KY!$D$62=0,PP2_DINH_LUONG!$C$101<0.15,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0),"⚠ GẦN CÁCH TÒNG – "
    #   ,"")&DUNG_HY_KY!$C$102&IF(AND(DUNG_HY_KY!$D$72="",DUNG_HY_KY!$D$62=0,PP2_DINH_LUONG!$C$101<0.15,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=
    #   0),CHAR(10)&"⚠ GẦN CÁCH TÒNG: nhật can không có khí gốc, phe mình theo PP2 chỉ còn "&TEXT(PP2_DINH_LUONG!$C$101,"0%")&" nhưng chưa đủ điều kiện tòng c
    #   ủa sách (P2 tr.373-375 – vd tòng nhi đòi “không có quan, sát hoặc ấn khắc nhật chủ hoặc khắc thực thương”; xem DUNG_HY_KY dòng 62-72) nên công cụ giữ 
    #   chính cách và dụng thần phù ức (có thể khuyết trong tứ trụ)"&". Một số người xem coi đây là cách tòng theo nhóm mạnh nhất là "&IF(PP2_DINH_LUONG!$D$10
    #   7>=MAX(PP2_DINH_LUONG!$D$108,PP2_DINH_LUONG!$D$109),"thực thương (cách tòng nhi)",IF(PP2_DINH_LUONG!$D$108>=PP2_DINH_LUONG!$D$109,"tài (cách tòng tài)
    #   ","quan sát (cách tòng sát)"))&": khi đó dụng thần là nhóm này, còn ấn và tỉ kiếp thành kỵ – các vận, năm ấn, tỉ kiếp sẽ đọc NGƯỢC với kết quả hiện tạ
    #   i. Hãy kiểm chứng bằng các năm đã qua (giả định G25).","")
    return X.cat(X.cat(('⚠ GẦN CÁCH TÒNG – ' if X.truthy(X.AND(X.eq(E.v(8,72,4),''), X.eq(E.v(8,62,4),0.0), X.lt(E.v(10,101,3),0.15), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ''),E.v(8,102,3)),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ GẦN CÁCH TÒNG: nhật can không có khí gốc, phe mình theo PP2 chỉ còn '),X.TEXT(E.v(10,101,3), '0%')),' nhưng chưa đủ điều kiện tòng của sách (P2 tr.373-375 – vd tòng nhi đòi “không có quan, sát hoặc ấn khắc nhật chủ hoặc khắc thực thương”; xem DUNG_HY_KY dòng 62-72) nên công cụ giữ chính cách và dụng thần phù ức (có thể khuyết trong tứ trụ)'),'. Một số người xem coi đây là cách tòng theo nhóm mạnh nhất là '),('thực thương (cách tòng nhi)' if X.truthy(X.ge(E.v(10,107,4),X.MAX(E.rg(10,108,4,108,4), E.rg(10,109,4,109,4)))) else ('tài (cách tòng tài)' if X.truthy(X.ge(E.v(10,108,4),E.v(10,109,4))) else 'quan sát (cách tòng sát)'))),': khi đó dụng thần là nhóm này, còn ấn và tỉ kiếp thành kỵ – các vận, năm ấn, tỉ kiếp sẽ đọc NGƯỢC với kết quả hiện tại. Hãy kiểm chứng bằng các năm đã qua (giả định G25).') if X.truthy(X.AND(X.eq(E.v(8,72,4),''), X.eq(E.v(8,62,4),0.0), X.lt(E.v(10,101,3),0.15), X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0))) else ''))

def f66(E, r, c):
    # Ô D35
    # = DUNG_HY_KY!$D$102
    return E.v(8,102,4)

def f67(E, r, c):
    # Ô E35
    # = DOI_CHIEU!$E$22
    return E.v(7,22,5)

def f68(E, r, c):
    # Ô F35
    # = "Chính cách: chọn dụng theo phù ức (P1 tr.97-117). Cách đặc biệt: P2 tr.369-386."
    return 'Chính cách: chọn dụng theo phù ức (P1 tr.97-117). Cách đặc biệt: P2 tr.369-386.'

def f69(E, r, c):
    # Ô H35
    # = LUAN_GIAI!$J$108
    return E.v(4,108,10)

def f70(E, r, c):
    # Ô C36
    # = IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$106,DUNG_HY_KY!$C$106)
    return (E.v(8,106,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,106,3))

def f71(E, r, c):
    # Ô D36
    # = DUNG_HY_KY!$D$106
    return E.v(8,106,4)

def f72(E, r, c):
    # Ô E36
    # = DOI_CHIEU!$E$23
    return E.v(7,23,5)

def f73(E, r, c):
    # Ô F36
    # = "PP1: "&DUNG_HY_KY!$C$115&" Chữ đại diện: "&DUNG_HY_KY!$C$136&IF(DUNG_HY_KY!$C$114<>"Không"," – KHUYẾT","")&". || PP2: "&DUNG_HY_KY!$D$115&" Chữ đại d
    #   iện: "&DUNG_HY_KY!$C$137&IF(DUNG_HY_KY!$D$114<>"Không"," – KHUYẾT","")&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('PP1: ',E.v(8,115,3)),' Chữ đại diện: '),E.v(8,136,3)),(' – KHUYẾT' if X.truthy(X.ne(E.v(8,114,3),'Không')) else '')),'. || PP2: '),E.v(8,115,4)),' Chữ đại diện: '),E.v(8,137,3)),(' – KHUYẾT' if X.truthy(X.ne(E.v(8,114,4),'Không')) else '')),'.')

def f74(E, r, c):
    # Ô G36
    # = IF(DUNG_HY_KY!$C$109<0,"","PP1: "&INDEX(DUNG_HY_KY!$C$24:$C$28,DUNG_HY_KY!$C$109+1)&" chữ "&INDEX(HANH_TEN,DUNG_HY_KY!$C$109+1))&IF(DUNG_HY_KY!$D$109<
    #   0,""," | PP2: "&TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$109+1),"0.0")&" độ "&INDEX(HANH_TEN,DUNG_HY_KY!$D$109+1)&" ("&TEXT(INDEX(DUNG_HY_KY!$D
    #   $24:$D$28,DUNG_HY_KY!$D$109+1)/PP2_DINH_LUONG!$C$100,"0%")&")")
    return X.cat(('' if X.truthy(X.lt(E.v(8,109,3),0.0)) else X.cat(X.cat(X.cat('PP1: ',X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,109,3),1.0))),' chữ '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,3),1.0)))),('' if X.truthy(X.lt(E.v(8,109,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' | PP2: ',X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,109,4),1.0)), '0.0')),' độ '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,109,4),1.0))),' ('),X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,109,4),1.0)),E.v(10,100,3)), '0%')),')')))

def f75(E, r, c):
    # Ô H36
    # = LUAN_GIAI!$J$147
    return E.v(4,147,10)

def f76(E, r, c):
    # Ô C37
    # = IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$107,DUNG_HY_KY!$C$107)
    return (E.v(8,107,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,107,3))

def f77(E, r, c):
    # Ô D37
    # = DUNG_HY_KY!$D$107
    return E.v(8,107,4)

def f78(E, r, c):
    # Ô E37
    # = DOI_CHIEU!$E$24
    return E.v(7,24,5)

def f79(E, r, c):
    # Ô F37
    # = IF(OR(DUNG_HY_KY!$C$72<>"",DUNG_HY_KY!$D$72<>""),"Cách đặc biệt: hỷ là hành sinh dụng nếu thuộc hành vận tốt (P2 tr.382, 385-386). ","")&"Hỷ thần = hà
    #   nh sinh dụng thần; nếu hành đó thuộc phe gây mất cân bằng thì lấy hành khắc kỵ thần (P2 tr.382; P1 tr.112)."
    return X.cat(('Cách đặc biệt: hỷ là hành sinh dụng nếu thuộc hành vận tốt (P2 tr.382, 385-386). ' if X.truthy(X.OR(X.ne(E.v(8,72,3),''), X.ne(E.v(8,72,4),''))) else ''),'Hỷ thần = hành sinh dụng thần; nếu hành đó thuộc phe gây mất cân bằng thì lấy hành khắc kỵ thần (P2 tr.382; P1 tr.112).')

def f80(E, r, c):
    # Ô G37
    # = IF(DUNG_HY_KY!$C$110<0,"","PP1: "&INDEX(DUNG_HY_KY!$C$24:$C$28,DUNG_HY_KY!$C$110+1)&" chữ "&INDEX(HANH_TEN,DUNG_HY_KY!$C$110+1))&IF(DUNG_HY_KY!$D$110<
    #   0,""," | PP2: "&TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$110+1),"0.0")&" độ "&INDEX(HANH_TEN,DUNG_HY_KY!$D$110+1)&" ("&TEXT(INDEX(DUNG_HY_KY!$D
    #   $24:$D$28,DUNG_HY_KY!$D$110+1)/PP2_DINH_LUONG!$C$100,"0%")&")")
    return X.cat(('' if X.truthy(X.lt(E.v(8,110,3),0.0)) else X.cat(X.cat(X.cat('PP1: ',X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,110,3),1.0))),' chữ '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,3),1.0)))),('' if X.truthy(X.lt(E.v(8,110,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' | PP2: ',X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,110,4),1.0)), '0.0')),' độ '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,110,4),1.0))),' ('),X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,110,4),1.0)),E.v(10,100,3)), '0%')),')')))

def f81(E, r, c):
    # Ô H37
    # = LUAN_GIAI!$J$148
    return E.v(4,148,10)

def f82(E, r, c):
    # Ô C38
    # = IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$108,DUNG_HY_KY!$C$108)
    return (E.v(8,108,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,108,3))

def f83(E, r, c):
    # Ô D38
    # = DUNG_HY_KY!$D$108
    return E.v(8,108,4)

def f84(E, r, c):
    # Ô E38
    # = DOI_CHIEU!$E$25
    return E.v(7,25,5)

def f85(E, r, c):
    # Ô F38
    # = "Kỵ thần = hành khắc dụng thần (P2 tr.382). Theo cân bằng, nhóm nhiều cũng là kỵ (P1 tr.180): PP1 "&IF(OR(DUNG_HY_KY!$C$92=DUNG_HY_KY!$C$103,DUNG_HY_K
    #   Y!$C$92=DUNG_HY_KY!$C$104),DUNG_HY_KY!$C$112&" – nhưng nhóm này đang là "&IF(DUNG_HY_KY!$C$92=DUNG_HY_KY!$C$103,"dụng","hỷ")&" thần (thân vượng mà tài
    #   /quan sát nhiều – P2 tr.380-381) nên KHÔNG tính là kỵ",DUNG_HY_KY!$C$112)&"; PP2 "&IF(OR(DUNG_HY_KY!$D$92=DUNG_HY_KY!$D$103,DUNG_HY_KY!$D$92=DUNG_HY_K
    #   Y!$D$104),DUNG_HY_KY!$D$112&" – nhưng nhóm này đang là "&IF(DUNG_HY_KY!$D$92=DUNG_HY_KY!$D$103,"dụng","hỷ")&" thần (thân vượng mà tài/quan sát nhiều –
    #    P2 tr.380-381) nên KHÔNG tính là kỵ",DUNG_HY_KY!$D$112)&". Phe gây mất cân bằng – PP1: "&DUNG_HY_KY!$C$113&"; PP2: "&DUNG_HY_KY!$D$113
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Kỵ thần = hành khắc dụng thần (P2 tr.382). Theo cân bằng, nhóm nhiều cũng là kỵ (P1 tr.180): PP1 ',(X.cat(X.cat(X.cat(E.v(8,112,3),' – nhưng nhóm này đang là '),('dụng' if X.truthy(X.eq(E.v(8,92,3),E.v(8,103,3))) else 'hỷ')),' thần (thân vượng mà tài/quan sát nhiều – P2 tr.380-381) nên KHÔNG tính là kỵ') if X.truthy(X.OR(X.eq(E.v(8,92,3),E.v(8,103,3)), X.eq(E.v(8,92,3),E.v(8,104,3)))) else E.v(8,112,3))),'; PP2 '),(X.cat(X.cat(X.cat(E.v(8,112,4),' – nhưng nhóm này đang là '),('dụng' if X.truthy(X.eq(E.v(8,92,4),E.v(8,103,4))) else 'hỷ')),' thần (thân vượng mà tài/quan sát nhiều – P2 tr.380-381) nên KHÔNG tính là kỵ') if X.truthy(X.OR(X.eq(E.v(8,92,4),E.v(8,103,4)), X.eq(E.v(8,92,4),E.v(8,104,4)))) else E.v(8,112,4))),'. Phe gây mất cân bằng – PP1: '),E.v(8,113,3)),'; PP2: '),E.v(8,113,4))

def f86(E, r, c):
    # Ô G38
    # = IF(DUNG_HY_KY!$C$111<0,"","PP1: "&INDEX(DUNG_HY_KY!$C$24:$C$28,DUNG_HY_KY!$C$111+1)&" chữ "&INDEX(HANH_TEN,DUNG_HY_KY!$C$111+1))&IF(DUNG_HY_KY!$D$111<
    #   0,""," | PP2: "&TEXT(INDEX(DUNG_HY_KY!$D$24:$D$28,DUNG_HY_KY!$D$111+1),"0.0")&" độ "&INDEX(HANH_TEN,DUNG_HY_KY!$D$111+1)&" ("&TEXT(INDEX(DUNG_HY_KY!$D
    #   $24:$D$28,DUNG_HY_KY!$D$111+1)/PP2_DINH_LUONG!$C$100,"0%")&")")
    return X.cat(('' if X.truthy(X.lt(E.v(8,111,3),0.0)) else X.cat(X.cat(X.cat('PP1: ',X.INDEX(E.rg(8,24,3,28,3), X.add(E.v(8,111,3),1.0))),' chữ '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,3),1.0)))),('' if X.truthy(X.lt(E.v(8,111,4),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(' | PP2: ',X.TEXT(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,111,4),1.0)), '0.0')),' độ '),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(8,111,4),1.0))),' ('),X.TEXT(X.div(X.INDEX(E.rg(8,24,4,28,4), X.add(E.v(8,111,4),1.0)),E.v(10,100,3)), '0%')),')')))

def f87(E, r, c):
    # Ô H38
    # = LUAN_GIAI!$J$149
    return E.v(4,149,10)

def f88(E, r, c):
    # Ô C39
    # = DUNG_HY_KY!$C$147&CHAR(10)&DUNG_HY_KY!$C$148&CHAR(10)&"⇒ "&DUNG_HY_KY!$C$149&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$218,DUNG_HY_KY!$C$218)<>"",CHAR(10)
    #   &IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$218,DUNG_HY_KY!$C$218),"")&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$205,DUNG_HY_KY!$C$205)<>"",CHAR(10)&"Cách cục: "&
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$202,DUNG_HY_KY!$C$202)&" chưa thành cách – "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$205,DUNG_HY_KY!$C$205)&".","")&IF(D
    #   UNG_HY_KY!$C$226<>"",CHAR(10)&"Ứng viên dụng thần thay thế (tham khảo): "&DUNG_HY_KY!$C$226,"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(8,147,3),X.CHAR(10.0)),E.v(8,148,3)),X.CHAR(10.0)),'⇒ '),E.v(8,149,3)),(X.cat(X.CHAR(10.0),(E.v(8,218,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,218,3))) if X.truthy(X.ne((E.v(8,218,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,218,3)),'')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'Cách cục: '),(E.v(8,202,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,202,3))),' chưa thành cách – '),(E.v(8,205,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,205,3))),'.') if X.truthy(X.ne((E.v(8,205,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,205,3)),'')) else '')),(X.cat(X.cat(X.CHAR(10.0),'Ứng viên dụng thần thay thế (tham khảo): '),E.v(8,226,3)) if X.truthy(X.ne(E.v(8,226,3),'')) else ''))

def f89(E, r, c):
    # Ô H39
    # = LUAN_GIAI!$J$107
    return E.v(4,107,10)

def f90(E, r, c):
    # Ô C40
    # = THAN_SAT!$C$55
    return E.v(12,55,3)

def f91(E, r, c):
    # Ô F40
    # = "Thần sát không phụ thuộc phương pháp. Tra Văn Xương theo can ngày (và can năm) (P1 tr.140; P2 tr.294-295); gặp hình, xung, khắc, hại thì như không có
    #    (P1 tr.140). Trùng dụng/hỷ xét theo PP1."
    return 'Thần sát không phụ thuộc phương pháp. Tra Văn Xương theo can ngày (và can năm) (P1 tr.140; P2 tr.294-295); gặp hình, xung, khắc, hại thì như không có (P1 tr.140). Trùng dụng/hỷ xét theo PP1.'

def f92(E, r, c):
    # Ô G40
    # = "Số Văn Xương có mặt: "&THAN_SAT!$C$53
    return X.cat('Số Văn Xương có mặt: ',E.v(12,53,3))

def f93(E, r, c):
    # Ô H40
    # = LUAN_GIAI!$J$150
    return E.v(4,150,10)

def f94(E, r, c):
    # Ô C41
    # = THAN_SAT!$C$56&IF(THAN_SAT!$C$54=0," "&THAN_SAT!$C$57,"")
    return X.cat(E.v(12,56,3),(X.cat(' ',E.v(12,57,3)) if X.truthy(X.eq(E.v(12,54,3),0.0)) else ''))

def f95(E, r, c):
    # Ô F41
    # = "Tra theo can năm (quý nhân lớn) và can ngày (quý nhân nhỏ) (P1 tr.138; P2 tr.294); gặp hình xung khắc hại thì như không có; tứ trụ không có thì xem t
    #   hai nguyên, cung mệnh (P2 tr.296)."
    return 'Tra theo can năm (quý nhân lớn) và can ngày (quý nhân nhỏ) (P1 tr.138; P2 tr.294); gặp hình xung khắc hại thì như không có; tứ trụ không có thì xem thai nguyên, cung mệnh (P2 tr.296).'

def f96(E, r, c):
    # Ô G41
    # = "Số quý nhân có mặt: "&THAN_SAT!$C$54&"; tuần không: "&THAN_SAT!$D$17&", "&THAN_SAT!$D$18
    return X.cat(X.cat(X.cat(X.cat(X.cat('Số quý nhân có mặt: ',E.v(12,54,3)),'; tuần không: '),E.v(12,17,4)),', '),E.v(12,18,4))

def f97(E, r, c):
    # Ô H41
    # = LUAN_GIAI!$J$151
    return E.v(4,151,10)

def f98(E, r, c):
    # Ô C42
    # = DOI_CHIEU!C34&IF(PP1_DINH_TINH!$C$111<>"",CHAR(10)&PP1_DINH_TINH!$C$111,"")&IF(DOI_CHIEU!$C$52<>"",CHAR(10)&DOI_CHIEU!$C$52,"")
    return X.cat(X.cat(E.v(7,r-8,c),(X.cat(X.CHAR(10.0),E.v(9,111,3)) if X.truthy(X.ne(E.v(9,111,3),'')) else '')),(X.cat(X.CHAR(10.0),E.v(7,52,3)) if X.truthy(X.ne(E.v(7,52,3),'')) else ''))

def f99(E, r, c):
    # Ô E42
    # = "Độ tin: "&DOI_CHIEU!C35
    return X.cat('Độ tin: ',E.v(7,r-7,c-2))

def f100(E, r, c):
    # Ô F42
    # = DOI_CHIEU!C36
    return E.v(7,r-6,c-3)

def f101(E, r, c):
    # Ô G42
    # = DOI_CHIEU!C37
    return E.v(7,r-5,c-4)

def f102(E, r, c):
    # Ô H42
    # = LUAN_GIAI!$J$152
    return E.v(4,152,10)

def f103(E, r, c):
    # Ô C43
    # = IF(DUNG_HY_KY!$C$109=DUNG_HY_KY!$D$109,"Hai phương pháp THỐNG NHẤT dụng thần "&DUNG_HY_KY!$C$106&".","Hai phương pháp KHÁC NHAU về dụng thần (PP1: "&D
    #   UNG_HY_KY!$C$106&"; PP2: "&DUNG_HY_KY!$D$106&"). "&IF(DOI_CHIEU!$C$33=0,"Hai phương pháp cùng trạng thái nhật chủ; khác nhau do nhóm nhiều nhất (PP1 đ
    #   ếm chữ, PP2 tính độ) – xem DOI_CHIEU.","Khuyến nghị: "&DOI_CHIEU!$D$33&" → "&DOI_CHIEU!$C$37&" (độ tin "&LOWER(DOI_CHIEU!$C$35)&"; căn cứ ở DOI_CHIEU 
    #   dòng 28-37)."))
    return (X.cat(X.cat('Hai phương pháp THỐNG NHẤT dụng thần ',E.v(8,106,3)),'.') if X.truthy(X.eq(E.v(8,109,3),E.v(8,109,4))) else X.cat(X.cat(X.cat(X.cat(X.cat('Hai phương pháp KHÁC NHAU về dụng thần (PP1: ',E.v(8,106,3)),'; PP2: '),E.v(8,106,4)),'). '),('Hai phương pháp cùng trạng thái nhật chủ; khác nhau do nhóm nhiều nhất (PP1 đếm chữ, PP2 tính độ) – xem DOI_CHIEU.' if X.truthy(X.eq(E.v(7,33,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Khuyến nghị: ',E.v(7,33,4)),' → '),E.v(7,37,3)),' (độ tin '),X.LOWER(E.v(7,35,3))),'; căn cứ ở DOI_CHIEU dòng 28-37).'))))

def f104(E, r, c):
    # Ô H43
    # = LUAN_GIAI!$J$153
    return E.v(4,153,10)

def f105(E, r, c):
    # Ô C45
    # = "Giàu – nghèo: "&LEFT(C87,IFERROR(FIND(CHAR(10),C87)-1,LEN(C87)))&". Sang – hèn: "&LEFT(C99,IFERROR(FIND(CHAR(10),C99)-1,LEN(C99)))&". Cát – hung: "&L
    #   EFT(C173,IFERROR(FIND(CHAR(10),C173)-1,LEN(C173)))&". Thọ – yểu: "&LEFT(C143,IFERROR(FIND(CHAR(10),C143)-1,LEN(C143)))&" (Bảng P1 tr.173 – so lực nhóm
    #    sinh/khắc theo vượng suy.)"&CHAR(10)&"Cách đọc mức: “rõ rệt” = chiều này mạnh gấp đôi chiều kia; “nghiêng” = mạnh hơn nhưng chưa gấp đôi; “nghiêng nh
    #   ẹ” = một phương pháp cân bằng, phương pháp kia nghiêng (giả định của công cụ – MENH_CUC). Ý nghĩa từng chiều theo định nghĩa của "&"sách P1 tr.164-172
    #    (giả định G24)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Giàu – nghèo: ',X.LEFT(E.v(3,r+42,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+42,c)),1.0), lambda: X.LEN(E.v(3,r+42,c))))),'. Sang – hèn: '),X.LEFT(E.v(3,r+54,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+54,c)),1.0), lambda: X.LEN(E.v(3,r+54,c))))),'. Cát – hung: '),X.LEFT(E.v(3,r+128,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+128,c)),1.0), lambda: X.LEN(E.v(3,r+128,c))))),'. Thọ – yểu: '),X.LEFT(E.v(3,r+98,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+98,c)),1.0), lambda: X.LEN(E.v(3,r+98,c))))),' (Bảng P1 tr.173 – so lực nhóm sinh/khắc theo vượng suy.)'),X.CHAR(10.0)),'Cách đọc mức: “rõ rệt” = chiều này mạnh gấp đôi chiều kia; “nghiêng” = mạnh hơn nhưng chưa gấp đôi; “nghiêng nhẹ” = một phương pháp cân bằng, phương pháp kia nghiêng (giả định của công cụ – MENH_CUC). Ý nghĩa từng chiều theo định nghĩa của '),'sách P1 tr.164-172 (giả định G24).')

def f106(E, r, c):
    # Ô H45
    # = LUAN_GIAI!$J$181
    return E.v(4,181,10)

def f107(E, r, c):
    # Ô C47
    # = DUNG_HY_KY!$C$140&IF(AND(IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),4,IF(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),
    #   1,-1))>=0,IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),4,IF(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),1,-1))=IF(DOI_C
    #   HIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),CHAR(10)&"⚠ Hành điều hầu ("&INDEX(HANH_TEN,IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D
    #   $6=7),4,IF(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),1,-1))+1)&") TRÙNG kỵ thần theo phù ức: hai nhu cầu ngược nhau. Sách cho sinh ph
    #   ù (phù ức) làm chủ (P1 tr.113) nên công cụ vẫn chấm vận, năm theo kỵ thần; khi đọc vận, năm mang hành "&INDEX(HANH_TEN,IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_G
    #   IAN!$D$6=6,TRUNG_GIAN!$D$6=7),4,IF(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),1,-1))+1)&" cần biết thêm mặt điều hòa khí hậu (P1 tr.10
    #   0) – đây cũng là lý do một số người xem chọn "&INDEX(HANH_TEN,IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),4,IF(OR(TRUNG_GIAN!$D$6=11,
    #   TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),1,-1))+1)&" làm dụng thần.","")
    return X.cat(E.v(8,140,3),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Hành điều hầu ('),X.INDEX(E.rg(24,33,1,37,1), X.add((4.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else X.neg(1.0))),1.0))),') TRÙNG kỵ thần theo phù ức: hai nhu cầu ngược nhau. Sách cho sinh phù (phù ức) làm chủ (P1 tr.113) nên công cụ vẫn chấm vận, năm theo kỵ thần; khi đọc vận, năm mang hành '),X.INDEX(E.rg(24,33,1,37,1), X.add((4.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else X.neg(1.0))),1.0))),' cần biết thêm mặt điều hòa khí hậu (P1 tr.100) – đây cũng là lý do một số người xem chọn '),X.INDEX(E.rg(24,33,1,37,1), X.add((4.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else X.neg(1.0))),1.0))),' làm dụng thần.') if X.truthy(X.AND(X.ge((4.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else X.neg(1.0))),0.0), X.eq((4.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0))) else (1.0 if X.truthy(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0))) else X.neg(1.0))),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f108(E, r, c):
    # Ô H47
    # = LUAN_GIAI!$J$160
    return E.v(4,160,10)

def f109(E, r, c):
    # Ô C48
    # = DUNG_HY_KY!$C$141
    return E.v(8,141,3)

def f110(E, r, c):
    # Ô H48
    # = LUAN_GIAI!$J$161
    return E.v(4,161,10)

def f111(E, r, c):
    # Ô C49
    # = DUNG_HY_KY!$C$142
    return E.v(8,142,3)

def f112(E, r, c):
    # Ô H49
    # = LUAN_GIAI!$J$162
    return E.v(4,162,10)

def f113(E, r, c):
    # Ô C50
    # = DUNG_HY_KY!$C$143
    return E.v(8,143,3)

def f114(E, r, c):
    # Ô H50
    # = LUAN_GIAI!$J$163
    return E.v(4,163,10)

def f115(E, r, c):
    # Ô C51
    # = DUNG_HY_KY!$C$144
    return E.v(8,144,3)

def f116(E, r, c):
    # Ô H51
    # = LUAN_GIAI!$J$164
    return E.v(4,164,10)

def f117(E, r, c):
    # Ô C52
    # = DUNG_HY_KY!$C$145&IF(AND(DUNG_HY_KY!$C$145<>"",DUNG_HY_KY!$C$151<>""),CHAR(10),"")&DUNG_HY_KY!$C$151&IF(AND(DUNG_HY_KY!$C$145&DUNG_HY_KY!$C$151<>"",DU
    #   NG_HY_KY!$C$198<>""),CHAR(10),"")&DUNG_HY_KY!$C$198
    return X.cat(X.cat(X.cat(X.cat(E.v(8,145,3),(X.CHAR(10.0) if X.truthy(X.AND(X.ne(E.v(8,145,3),''), X.ne(E.v(8,151,3),''))) else '')),E.v(8,151,3)),(X.CHAR(10.0) if X.truthy(X.AND(X.ne(X.cat(E.v(8,145,3),E.v(8,151,3)),''), X.ne(E.v(8,198,3),''))) else '')),E.v(8,198,3))

def f118(E, r, c):
    # Ô C55
    # = TRUNG_GIAN!$C$10&IF(0=2," (nhật chủ)"," – "&TRUNG_GIAN!$C$12)
    return X.cat(E.v(23,10,3),(' (nhật chủ)' if X.truthy(X.eq(0.0,2.0)) else X.cat(' – ',E.v(23,12,3))))

def f119(E, r, c):
    # Ô D55
    # = TRUNG_GIAN!$D$10&IF(1=2," (nhật chủ)"," – "&TRUNG_GIAN!$D$12)
    return X.cat(E.v(23,10,4),(' (nhật chủ)' if X.truthy(X.eq(1.0,2.0)) else X.cat(' – ',E.v(23,12,4))))

def f120(E, r, c):
    # Ô E55
    # = TRUNG_GIAN!$E$10&IF(2=2," (nhật chủ)"," – "&TRUNG_GIAN!$E$12)
    return X.cat(E.v(23,10,5),(' (nhật chủ)' if X.truthy(X.eq(2.0,2.0)) else X.cat(' – ',E.v(23,12,5))))

def f121(E, r, c):
    # Ô F55
    # = TRUNG_GIAN!$F$10&IF(3=2," (nhật chủ)"," – "&TRUNG_GIAN!$F$12)
    return X.cat(E.v(23,10,6),(' (nhật chủ)' if X.truthy(X.eq(3.0,2.0)) else X.cat(' – ',E.v(23,12,6))))

def f122(E, r, c):
    # Ô C56
    # = THONG_CAN!$C$69
    return E.v(11,69,3)

def f123(E, r, c):
    # Ô D56
    # = THONG_CAN!$D$69
    return E.v(11,69,4)

def f124(E, r, c):
    # Ô E56
    # = THONG_CAN!$E$69
    return E.v(11,69,5)

def f125(E, r, c):
    # Ô F56
    # = THONG_CAN!$F$69
    return E.v(11,69,6)

def f126(E, r, c):
    # Ô H56
    # = LUAN_GIAI!$J$155
    return E.v(4,155,10)

def f127(E, r, c):
    # Ô C57
    # = THONG_CAN!$C$70&IF(THONG_CAN!$C$72="",""," – "&THONG_CAN!$C$72)
    return X.cat(E.v(11,70,3),('' if X.truthy(X.eq(E.v(11,72,3),'')) else X.cat(' – ',E.v(11,72,3))))

def f128(E, r, c):
    # Ô D57
    # = THONG_CAN!$D$70&IF(THONG_CAN!$D$72="",""," – "&THONG_CAN!$D$72)
    return X.cat(E.v(11,70,4),('' if X.truthy(X.eq(E.v(11,72,4),'')) else X.cat(' – ',E.v(11,72,4))))

def f129(E, r, c):
    # Ô E57
    # = THONG_CAN!$E$70&IF(THONG_CAN!$E$72="",""," – "&THONG_CAN!$E$72)
    return X.cat(E.v(11,70,5),('' if X.truthy(X.eq(E.v(11,72,5),'')) else X.cat(' – ',E.v(11,72,5))))

def f130(E, r, c):
    # Ô F57
    # = THONG_CAN!$F$70&IF(THONG_CAN!$F$72="",""," – "&THONG_CAN!$F$72)
    return X.cat(E.v(11,70,6),('' if X.truthy(X.eq(E.v(11,72,6),'')) else X.cat(' – ',E.v(11,72,6))))

def f131(E, r, c):
    # Ô H57
    # = LUAN_GIAI!$J$156
    return E.v(4,156,10)

def f132(E, r, c):
    # Ô C58
    # = IF(THONG_CAN!$F$76="Thấu và có gốc",THONG_CAN!$C$76&" ("&THONG_CAN!$D$76&", tàng ở "&THONG_CAN!$B$76&") ","")&IF(THONG_CAN!$F$77="Thấu và có gốc",THON
    #   G_CAN!$C$77&" ("&THONG_CAN!$D$77&", tàng ở "&THONG_CAN!$B$77&") ","")&IF(THONG_CAN!$F$78="Thấu và có gốc",THONG_CAN!$C$78&" ("&THONG_CAN!$D$78&", tàng
    #    ở "&THONG_CAN!$B$78&") ","")&IF(THONG_CAN!$F$79="Thấu và có gốc",THONG_CAN!$C$79&" ("&THONG_CAN!$D$79&", tàng ở "&THONG_CAN!$B$79&") ","")&IF(THONG_C
    #   AN!$F$80="Thấu và có gốc",THONG_CAN!$C$80&" ("&THONG_CAN!$D$80&", tàng ở "&THONG_CAN!$B$80&") ","")&IF(THONG_CAN!$F$81="Thấu và có gốc",THONG_CAN!$C$8
    #   1&" ("&THONG_CAN!$D$81&", tàng ở "&THONG_CAN!$B$81&") ","")&IF(THONG_CAN!$F$82="Thấu và có gốc",THONG_CAN!$C$82&" ("&THONG_CAN!$D$82&", tàng ở "&THONG
    #   _CAN!$B$82&") ","")&IF(THONG_CAN!$F$83="Thấu và có gốc",THONG_CAN!$C$83&" ("&THONG_CAN!$D$83&", tàng ở "&THONG_CAN!$B$83&") ","")&IF(THONG_CAN!$F$84="
    #   Thấu và có gốc",THONG_CAN!$C$84&" ("&THONG_CAN!$D$84&", tàng ở "&THONG_CAN!$B$84&") ","")&IF(THONG_CAN!$F$85="Thấu và có gốc",THONG_CAN!$C$85&" ("&THO
    #   NG_CAN!$D$85&", tàng ở "&THONG_CAN!$B$85&") ","")&IF(THONG_CAN!$F$86="Thấu và có gốc",THONG_CAN!$C$86&" ("&THONG_CAN!$D$86&", tàng ở "&THONG_CAN!$B$86
    #   &") ","")&IF(THONG_CAN!$F$87="Thấu và có gốc",THONG_CAN!$C$87&" ("&THONG_CAN!$D$87&", tàng ở "&THONG_CAN!$B$87&") ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,76,3),' ('),E.v(11,76,4)),', tàng ở '),E.v(11,76,2)),') ') if X.truthy(X.eq(E.v(11,76,6),'Thấu và có gốc')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,77,3),' ('),E.v(11,77,4)),', tàng ở '),E.v(11,77,2)),') ') if X.truthy(X.eq(E.v(11,77,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,78,3),' ('),E.v(11,78,4)),', tàng ở '),E.v(11,78,2)),') ') if X.truthy(X.eq(E.v(11,78,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,79,3),' ('),E.v(11,79,4)),', tàng ở '),E.v(11,79,2)),') ') if X.truthy(X.eq(E.v(11,79,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,80,3),' ('),E.v(11,80,4)),', tàng ở '),E.v(11,80,2)),') ') if X.truthy(X.eq(E.v(11,80,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,81,3),' ('),E.v(11,81,4)),', tàng ở '),E.v(11,81,2)),') ') if X.truthy(X.eq(E.v(11,81,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,82,3),' ('),E.v(11,82,4)),', tàng ở '),E.v(11,82,2)),') ') if X.truthy(X.eq(E.v(11,82,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,83,3),' ('),E.v(11,83,4)),', tàng ở '),E.v(11,83,2)),') ') if X.truthy(X.eq(E.v(11,83,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,84,3),' ('),E.v(11,84,4)),', tàng ở '),E.v(11,84,2)),') ') if X.truthy(X.eq(E.v(11,84,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,85,3),' ('),E.v(11,85,4)),', tàng ở '),E.v(11,85,2)),') ') if X.truthy(X.eq(E.v(11,85,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,86,3),' ('),E.v(11,86,4)),', tàng ở '),E.v(11,86,2)),') ') if X.truthy(X.eq(E.v(11,86,6),'Thấu và có gốc')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,87,3),' ('),E.v(11,87,4)),', tàng ở '),E.v(11,87,2)),') ') if X.truthy(X.eq(E.v(11,87,6),'Thấu và có gốc')) else ''))

def f133(E, r, c):
    # Ô H58
    # = LUAN_GIAI!$J$157
    return E.v(4,157,10)

def f134(E, r, c):
    # Ô C59
    # = IF(LEFT(THONG_CAN!$F$76,2)="Lộ",THONG_CAN!$C$76&" ("&THONG_CAN!$D$76&" – "&THONG_CAN!$F$76&") ","")&IF(LEFT(THONG_CAN!$F$77,2)="Lộ",THONG_CAN!$C$77&" 
    #   ("&THONG_CAN!$D$77&" – "&THONG_CAN!$F$77&") ","")&IF(LEFT(THONG_CAN!$F$78,2)="Lộ",THONG_CAN!$C$78&" ("&THONG_CAN!$D$78&" – "&THONG_CAN!$F$78&") ","")&
    #   IF(LEFT(THONG_CAN!$F$79,2)="Lộ",THONG_CAN!$C$79&" ("&THONG_CAN!$D$79&" – "&THONG_CAN!$F$79&") ","")&IF(LEFT(THONG_CAN!$F$80,2)="Lộ",THONG_CAN!$C$80&" 
    #   ("&THONG_CAN!$D$80&" – "&THONG_CAN!$F$80&") ","")&IF(LEFT(THONG_CAN!$F$81,2)="Lộ",THONG_CAN!$C$81&" ("&THONG_CAN!$D$81&" – "&THONG_CAN!$F$81&") ","")&
    #   IF(LEFT(THONG_CAN!$F$82,2)="Lộ",THONG_CAN!$C$82&" ("&THONG_CAN!$D$82&" – "&THONG_CAN!$F$82&") ","")&IF(LEFT(THONG_CAN!$F$83,2)="Lộ",THONG_CAN!$C$83&" 
    #   ("&THONG_CAN!$D$83&" – "&THONG_CAN!$F$83&") ","")&IF(LEFT(THONG_CAN!$F$84,2)="Lộ",THONG_CAN!$C$84&" ("&THONG_CAN!$D$84&" – "&THONG_CAN!$F$84&") ","")&
    #   IF(LEFT(THONG_CAN!$F$85,2)="Lộ",THONG_CAN!$C$85&" ("&THONG_CAN!$D$85&" – "&THONG_CAN!$F$85&") ","")&IF(LEFT(THONG_CAN!$F$86,2)="Lộ",THONG_CAN!$C$86&" 
    #   ("&THONG_CAN!$D$86&" – "&THONG_CAN!$F$86&") ","")&IF(LEFT(THONG_CAN!$F$87,2)="Lộ",THONG_CAN!$C$87&" ("&THONG_CAN!$D$87&" – "&THONG_CAN!$F$87&") ","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,76,3),' ('),E.v(11,76,4)),' – '),E.v(11,76,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,76,6), 2.0),'Lộ')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,77,3),' ('),E.v(11,77,4)),' – '),E.v(11,77,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,77,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,78,3),' ('),E.v(11,78,4)),' – '),E.v(11,78,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,78,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,79,3),' ('),E.v(11,79,4)),' – '),E.v(11,79,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,79,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,80,3),' ('),E.v(11,80,4)),' – '),E.v(11,80,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,80,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,81,3),' ('),E.v(11,81,4)),' – '),E.v(11,81,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,81,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,82,3),' ('),E.v(11,82,4)),' – '),E.v(11,82,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,82,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,83,3),' ('),E.v(11,83,4)),' – '),E.v(11,83,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,83,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,84,3),' ('),E.v(11,84,4)),' – '),E.v(11,84,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,84,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,85,3),' ('),E.v(11,85,4)),' – '),E.v(11,85,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,85,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,86,3),' ('),E.v(11,86,4)),' – '),E.v(11,86,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,86,6), 2.0),'Lộ')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(11,87,3),' ('),E.v(11,87,4)),' – '),E.v(11,87,6)),') ') if X.truthy(X.eq(X.LEFT(E.v(11,87,6), 2.0),'Lộ')) else ''))

def f135(E, r, c):
    # Ô H59
    # = LUAN_GIAI!$J$158
    return E.v(4,158,10)

def f136(E, r, c):
    # Ô C60
    # = "Nhật chủ "&E55&": "&LOWER(E56)&". Các can khác: năm "&C55&" – "&LOWER(C56)&"; tháng "&D55&" – "&LOWER(D56)&"; giờ "&F55&" – "&LOWER(F56)&". → Can có 
    #   gốc thì có lực, gặp vận/năm dễ phát huy; gốc dễ bị nhổ (chi bị xung) thì lực không bền; không có gốc (hư phù) thì yếu, dễ bị chế (P1 tr.32, 59-60; P2 
    #   tr.333-336)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Nhật chủ ',E.v(3,r-5,c+2)),': '),X.LOWER(E.v(3,r-4,c+2))),'. Các can khác: năm '),E.v(3,r-5,c)),' – '),X.LOWER(E.v(3,r-4,c))),'; tháng '),E.v(3,r-5,c+1)),' – '),X.LOWER(E.v(3,r-4,c+1))),'; giờ '),E.v(3,r-5,c+3)),' – '),X.LOWER(E.v(3,r-4,c+3))),'. → Can có gốc thì có lực, gặp vận/năm dễ phát huy; gốc dễ bị nhổ (chi bị xung) thì lực không bền; không có gốc (hư phù) thì yếu, dễ bị chế (P1 tr.32, 59-60; P2 tr.333-336).')

def f137(E, r, c):
    # Ô H60
    # = LUAN_GIAI!$J$159
    return E.v(4,159,10)

def f138(E, r, c):
    # Ô C62
    # = TRIM(IF(THAN_SAT!$H$85="Có","▸ THIÊN ĐỨC ("&TRIM(IF(THAN_SAT!$C$85="có","năm ","")&IF(THAN_SAT!$D$85="có","tháng ","")&IF(THAN_SAT!$E$85="có","ngày ",
    #   "")&IF(THAN_SAT!$F$85="có","giờ ",""))&"): "&THAN_SAT!$I$85&" [tra theo "&THAN_SAT!$G$85&"; "&THAN_SAT!$J$85&"]"&CHAR(10),"")&IF(THAN_SAT!$H$86="Có","
    #   ▸ NGUYỆT ĐỨC ("&TRIM(IF(THAN_SAT!$C$86="có","năm ","")&IF(THAN_SAT!$D$86="có","tháng ","")&IF(THAN_SAT!$E$86="có","ngày ","")&IF(THAN_SAT!$F$86="có","
    #   giờ ",""))&"): "&THAN_SAT!$I$86&" [tra theo "&THAN_SAT!$G$86&"; "&THAN_SAT!$J$86&"]"&CHAR(10),"")&IF(THAN_SAT!$H$87="Có","▸ LỘC ("&TRIM(IF(THAN_SAT!$C
    #   $87="có","năm ","")&IF(THAN_SAT!$D$87="có","tháng ","")&IF(THAN_SAT!$E$87="có","ngày ","")&IF(THAN_SAT!$F$87="có","giờ ",""))&"): "&THAN_SAT!$I$87&" [
    #   tra theo "&THAN_SAT!$G$87&"; "&THAN_SAT!$J$87&"]"&CHAR(10),"")&IF(THAN_SAT!$H$88="Có","▸ KÌNH DƯƠNG ("&TRIM(IF(THAN_SAT!$C$88="có","năm ","")&IF(THAN_
    #   SAT!$D$88="có","tháng ","")&IF(THAN_SAT!$E$88="có","ngày ","")&IF(THAN_SAT!$F$88="có","giờ ",""))&"): "&THAN_SAT!$I$88&" [tra theo "&THAN_SAT!$G$88&";
    #    "&THAN_SAT!$J$88&"]"&CHAR(10),"")&IF(THAN_SAT!$H$89="Có","▸ HOA CÁI ("&TRIM(IF(THAN_SAT!$C$89="có","năm ","")&IF(THAN_SAT!$D$89="có","tháng ","")&IF(
    #   THAN_SAT!$E$89="có","ngày ","")&IF(THAN_SAT!$F$89="có","giờ ",""))&"): "&THAN_SAT!$I$89&" [tra theo "&THAN_SAT!$G$89&"; "&THAN_SAT!$J$89&"]"&CHAR(10),
    #   "")&IF(THAN_SAT!$H$90="Có","▸ KIẾP SÁT ("&TRIM(IF(THAN_SAT!$C$90="có","năm ","")&IF(THAN_SAT!$D$90="có","tháng ","")&IF(THAN_SAT!$E$90="có","ngày ",""
    #   )&IF(THAN_SAT!$F$90="có","giờ ",""))&"): "&THAN_SAT!$I$90&" [tra theo "&THAN_SAT!$G$90&"; "&THAN_SAT!$J$90&"]"&CHAR(10),"")&IF(THAN_SAT!$H$91="Có","▸ 
    #   ĐÀO HOA ("&TRIM(IF(THAN_SAT!$C$91="có","năm ","")&IF(THAN_SAT!$D$91="có","tháng ","")&IF(THAN_SAT!$E$91="có","ngày ","")&IF(THAN_SAT!$F$91="có","giờ "
    #   ,""))&"): "&THAN_SAT!$I$91&" [tra theo "&THAN_SAT!$G$91&"; "&THAN_SAT!$J$91&"]"&CHAR(10),"")&IF(THAN_SAT!$H$92="Có","▸ TRẠCH MÃ ("&TRIM(IF(THAN_SAT!$C
    #   $92="có","năm ","")&IF(THAN_SAT!$D$92="có","tháng ","")&IF(THAN_SAT!$E$92="có","ngày ","")&IF(THAN_SAT!$F$92="có","giờ ",""))&"): "&THAN_SAT!$I$92&" [
    #   tra theo "&THAN_SAT!$G$92&"; "&THAN_SAT!$J$92&"]"&CHAR(10),"")&IF(THAN_SAT!$H$93="Có","▸ TƯỚNG TINH ("&TRIM(IF(THAN_SAT!$C$93="có","năm ","")&IF(THAN_
    #   SAT!$D$93="có","tháng ","")&IF(THAN_SAT!$E$93="có","ngày ","")&IF(THAN_SAT!$F$93="có","giờ ",""))&"): "&THAN_SAT!$I$93&" [tra theo "&THAN_SAT!$G$93&";
    #    "&THAN_SAT!$J$93&"]"&CHAR(10),"")&IF(THAN_SAT!$H$94="Có","▸ CÔ THẦN ("&TRIM(IF(THAN_SAT!$C$94="có","năm ","")&IF(THAN_SAT!$D$94="có","tháng ","")&IF(
    #   THAN_SAT!$E$94="có","ngày ","")&IF(THAN_SAT!$F$94="có","giờ ",""))&"): "&THAN_SAT!$I$94&" [tra theo "&THAN_SAT!$G$94&"; "&THAN_SAT!$J$94&"]"&CHAR(10),
    #   "")&IF(THAN_SAT!$H$95="Có","▸ QUẢ TÚ ("&TRIM(IF(THAN_SAT!$C$95="có","năm ","")&IF(THAN_SAT!$D$95="có","tháng ","")&IF(THAN_SAT!$E$95="có","ngày ","")&
    #   IF(THAN_SAT!$F$95="có","giờ ",""))&"): "&THAN_SAT!$I$95&" [tra theo "&THAN_SAT!$G$95&"; "&THAN_SAT!$J$95&"]"&CHAR(10),"")&IF(THAN_SAT!$H$96="Có","▸ TH
    #   IÊN ẤT QUÝ NHÂN ("&TRIM(IF(THAN_SAT!$C$96="có","năm ","")&IF(THAN_SAT!$D$96="có","tháng ","")&IF(THAN_SAT!$E$96="có","ngày ","")&IF(THAN_SAT!$F$96="có
    #   ","giờ ",""))&"): "&THAN_SAT!$I$96&" [tra theo "&THAN_SAT!$G$96&"; "&THAN_SAT!$J$96&"]"&CHAR(10),""))&IF(THAN_SAT!$C$134="","",CHAR(10)&THAN_SAT!$C$13
    #   4)
    return X.cat(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THIÊN ĐỨC (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,85,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,85,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,85,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,85,6),'có')) else '')))),'): '),E.v(12,85,9)),' [tra theo '),E.v(12,85,7)),'; '),E.v(12,85,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,85,8),'Có')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ NGUYỆT ĐỨC (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,86,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,86,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,86,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,86,6),'có')) else '')))),'): '),E.v(12,86,9)),' [tra theo '),E.v(12,86,7)),'; '),E.v(12,86,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,86,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ LỘC (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,87,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,87,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,87,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,87,6),'có')) else '')))),'): '),E.v(12,87,9)),' [tra theo '),E.v(12,87,7)),'; '),E.v(12,87,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,87,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ KÌNH DƯƠNG (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,88,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,88,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,88,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,88,6),'có')) else '')))),'): '),E.v(12,88,9)),' [tra theo '),E.v(12,88,7)),'; '),E.v(12,88,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,88,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ HOA CÁI (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,89,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,89,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,89,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,89,6),'có')) else '')))),'): '),E.v(12,89,9)),' [tra theo '),E.v(12,89,7)),'; '),E.v(12,89,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,89,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ KIẾP SÁT (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,90,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,90,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,90,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,90,6),'có')) else '')))),'): '),E.v(12,90,9)),' [tra theo '),E.v(12,90,7)),'; '),E.v(12,90,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,90,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ ĐÀO HOA (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,91,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,91,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,91,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,91,6),'có')) else '')))),'): '),E.v(12,91,9)),' [tra theo '),E.v(12,91,7)),'; '),E.v(12,91,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,91,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TRẠCH MÃ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,92,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,92,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,92,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,92,6),'có')) else '')))),'): '),E.v(12,92,9)),' [tra theo '),E.v(12,92,7)),'; '),E.v(12,92,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,92,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ TƯỚNG TINH (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,93,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,93,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,93,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,93,6),'có')) else '')))),'): '),E.v(12,93,9)),' [tra theo '),E.v(12,93,7)),'; '),E.v(12,93,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,93,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ CÔ THẦN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,94,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,94,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,94,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,94,6),'có')) else '')))),'): '),E.v(12,94,9)),' [tra theo '),E.v(12,94,7)),'; '),E.v(12,94,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,94,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ QUẢ TÚ (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,95,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,95,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,95,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,95,6),'có')) else '')))),'): '),E.v(12,95,9)),' [tra theo '),E.v(12,95,7)),'; '),E.v(12,95,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,95,8),'Có')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('▸ THIÊN ẤT QUÝ NHÂN (',X.TRIM(X.cat(X.cat(X.cat(('năm ' if X.truthy(X.eq(E.v(12,96,3),'có')) else ''),('tháng ' if X.truthy(X.eq(E.v(12,96,4),'có')) else '')),('ngày ' if X.truthy(X.eq(E.v(12,96,5),'có')) else '')),('giờ ' if X.truthy(X.eq(E.v(12,96,6),'có')) else '')))),'): '),E.v(12,96,9)),' [tra theo '),E.v(12,96,7)),'; '),E.v(12,96,10)),']'),X.CHAR(10.0)) if X.truthy(X.eq(E.v(12,96,8),'Có')) else ''))),('' if X.truthy(X.eq(E.v(12,134,3),'')) else X.cat(X.CHAR(10.0),E.v(12,134,3))))

def f139(E, r, c):
    # Ô H62
    # = LUAN_GIAI!$J$173
    return E.v(4,173,10)

def f140(E, r, c):
    # Ô C63
    # = TRIM(IF(THAN_SAT!$H$85="Như không có","Thiên đức; ","")&IF(THAN_SAT!$H$86="Như không có","Nguyệt đức; ","")&IF(THAN_SAT!$H$87="Như không có","Lộc; ","
    #   ")&IF(THAN_SAT!$H$88="Như không có","Kình dương; ","")&IF(THAN_SAT!$H$89="Như không có","Hoa cái; ","")&IF(THAN_SAT!$H$90="Như không có","Kiếp sát; ",
    #   "")&IF(THAN_SAT!$H$91="Như không có","Đào hoa; ","")&IF(THAN_SAT!$H$92="Như không có","Trạch mã; ","")&IF(THAN_SAT!$H$93="Như không có","Tướng tinh; "
    #   ,"")&IF(THAN_SAT!$H$94="Như không có","Cô thần; ","")&IF(THAN_SAT!$H$95="Như không có","Quả tú; ","")&IF(THAN_SAT!$H$96="Như không có","Thiên ất quý n
    #   hân; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('Thiên đức; ' if X.truthy(X.eq(E.v(12,85,8),'Như không có')) else ''),('Nguyệt đức; ' if X.truthy(X.eq(E.v(12,86,8),'Như không có')) else '')),('Lộc; ' if X.truthy(X.eq(E.v(12,87,8),'Như không có')) else '')),('Kình dương; ' if X.truthy(X.eq(E.v(12,88,8),'Như không có')) else '')),('Hoa cái; ' if X.truthy(X.eq(E.v(12,89,8),'Như không có')) else '')),('Kiếp sát; ' if X.truthy(X.eq(E.v(12,90,8),'Như không có')) else '')),('Đào hoa; ' if X.truthy(X.eq(E.v(12,91,8),'Như không có')) else '')),('Trạch mã; ' if X.truthy(X.eq(E.v(12,92,8),'Như không có')) else '')),('Tướng tinh; ' if X.truthy(X.eq(E.v(12,93,8),'Như không có')) else '')),('Cô thần; ' if X.truthy(X.eq(E.v(12,94,8),'Như không có')) else '')),('Quả tú; ' if X.truthy(X.eq(E.v(12,95,8),'Như không có')) else '')),('Thiên ất quý nhân; ' if X.truthy(X.eq(E.v(12,96,8),'Như không có')) else '')))

def f141(E, r, c):
    # Ô H63
    # = LUAN_GIAI!$J$174
    return E.v(4,174,10)

def f142(E, r, c):
    # Ô C64
    # = "Thần sát là yếu tố phụ, ""cũng từ trong âm dương, ngũ hành mà ra"" (P2 tr.660); sao gặp hình, xung, khắc, hại thì như không có (P1 tr.137). "&IF(C62=
    #   "","Lá số không có thần sát nào còn hiệu lực.","Các sao còn hiệu lực ở dòng trên chỉ làm tăng/giảm mức độ, không đổi kết luận theo dụng thần.")
    return X.cat('Thần sát là yếu tố phụ, "cũng từ trong âm dương, ngũ hành mà ra" (P2 tr.660); sao gặp hình, xung, khắc, hại thì như không có (P1 tr.137). ',('Lá số không có thần sát nào còn hiệu lực.' if X.truthy(X.eq(E.v(3,r-2,c),'')) else 'Các sao còn hiệu lực ở dòng trên chỉ làm tăng/giảm mức độ, không đổi kết luận theo dụng thần.'))

def f143(E, r, c):
    # Ô H64
    # = LUAN_GIAI!$J$109
    return E.v(4,109,10)

def f144(E, r, c):
    # Ô C65
    # = THAN_SAT!$C$144
    return E.v(12,144,3)

def f145(E, r, c):
    # Ô C68
    # = LEFT(C75,IFERROR(FIND(CHAR(10),C75)-1,LEN(C75)))&IF(LUU_NIEN!$B$5<=18," (Lưu ý tuổi: chưa thành niên – nghề chỉ là định hướng xa.)","")&CHAR(10)&"Nhân
    #    cách: "&(DUNG_HY_KY!$C$181)
    return X.cat(X.cat(X.cat(X.cat(X.LEFT(E.v(3,r+7,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+7,c)),1.0), lambda: X.LEN(E.v(3,r+7,c)))),(' (Lưu ý tuổi: chưa thành niên – nghề chỉ là định hướng xa.)' if X.truthy(X.le(E.v(15,5,2),18.0)) else '')),X.CHAR(10.0)),'Nhân cách: '),E.v(8,181,3))

def f146(E, r, c):
    # Ô C70
    # = TO_CHAT!C8
    return E.v(20,r-62,c)

def f147(E, r, c):
    # Ô H70
    # = LUAN_GIAI!$J$121
    return E.v(4,121,10)

def f148(E, r, c):
    # Ô C71
    # = "Điểm mạnh: "&IF(TO_CHAT!C12="","không có dòng nào của sách ứng với lá số",TO_CHAT!C12)&CHAR(10)&"Điểm yếu: "&IF(TO_CHAT!C13="","—",TO_CHAT!C13)
    return X.cat(X.cat(X.cat(X.cat('Điểm mạnh: ',('không có dòng nào của sách ứng với lá số' if X.truthy(X.eq(E.v(20,r-59,c),'')) else E.v(20,r-59,c))),X.CHAR(10.0)),'Điểm yếu: '),('—' if X.truthy(X.eq(E.v(20,r-58,c),'')) else E.v(20,r-58,c)))

def f149(E, r, c):
    # Ô H71
    # = LUAN_GIAI!$J$187
    return E.v(4,187,10)

def f150(E, r, c):
    # Ô C72
    # = IF(LUU_NIEN!$B$5<=18,"(Lưu ý tuổi: chưa thành niên – nghề là định hướng xa, đối chiếu thêm khi có sở thích và học vấn.)"&CHAR(10),"")&"Nghề hợp dụng t
    #   hần: "&TO_CHAT!C17&CHAR(10)&"Nghề hợp hỷ thần: "&TO_CHAT!C18&IF(TRIM(IF(AND(TO_CHAT!$L$55="ĐỦ",TO_CHAT!$I$55<>"Bất lợi"),"• "&TO_CHAT!$G$55&" ("&TO_CH
    #   AT!$H$55&") ","")&IF(AND(TO_CHAT!$L$56="ĐỦ",TO_CHAT!$I$56<>"Bất lợi"),"• "&TO_CHAT!$G$56&" ("&TO_CHAT!$H$56&") ","")&IF(AND(TO_CHAT!$L$57="ĐỦ",TO_CHAT
    #   !$I$57<>"Bất lợi"),"• "&TO_CHAT!$G$57&" ("&TO_CHAT!$H$57&") ","")&IF(AND(TO_CHAT!$L$58="ĐỦ",TO_CHAT!$I$58<>"Bất lợi"),"• "&TO_CHAT!$G$58&" ("&TO_CHAT!
    #   $H$58&") ","")&IF(AND(TO_CHAT!$L$59="ĐỦ",TO_CHAT!$I$59<>"Bất lợi"),"• "&TO_CHAT!$G$59&" ("&TO_CHAT!$H$59&") ","")&IF(AND(TO_CHAT!$L$60="ĐỦ",TO_CHAT!$I
    #   $60<>"Bất lợi"),"• "&TO_CHAT!$G$60&" ("&TO_CHAT!$H$60&") ","")&IF(AND(TO_CHAT!$L$61="ĐỦ",TO_CHAT!$I$61<>"Bất lợi"),"• "&TO_CHAT!$G$61&" ("&TO_CHAT!$H$
    #   61&") ","")&IF(AND(TO_CHAT!$L$62="ĐỦ",TO_CHAT!$I$62<>"Bất lợi"),"• "&TO_CHAT!$G$62&" ("&TO_CHAT!$H$62&") ","")&IF(AND(TO_CHAT!$L$63="ĐỦ",TO_CHAT!$I$63
    #   <>"Bất lợi"),"• "&TO_CHAT!$G$63&" ("&TO_CHAT!$H$63&") ",""))="","",CHAR(10)&"Theo thập thần (sách): "&TRIM(IF(AND(TO_CHAT!$L$55="ĐỦ",TO_CHAT!$I$55<>"B
    #   ất lợi"),"• "&TO_CHAT!$G$55&" ("&TO_CHAT!$H$55&") ","")&IF(AND(TO_CHAT!$L$56="ĐỦ",TO_CHAT!$I$56<>"Bất lợi"),"• "&TO_CHAT!$G$56&" ("&TO_CHAT!$H$56&") "
    #   ,"")&IF(AND(TO_CHAT!$L$57="ĐỦ",TO_CHAT!$I$57<>"Bất lợi"),"• "&TO_CHAT!$G$57&" ("&TO_CHAT!$H$57&") ","")&IF(AND(TO_CHAT!$L$58="ĐỦ",TO_CHAT!$I$58<>"Bất 
    #   lợi"),"• "&TO_CHAT!$G$58&" ("&TO_CHAT!$H$58&") ","")&IF(AND(TO_CHAT!$L$59="ĐỦ",TO_CHAT!$I$59<>"Bất lợi"),"• "&TO_CHAT!$G$59&" ("&TO_CHAT!$H$59&") ",""
    #   )&IF(AND(TO_CHAT!$L$60="ĐỦ",TO_CHAT!$I$60<>"Bất lợi"),"• "&TO_CHAT!$G$60&" ("&TO_CHAT!$H$60&") ","")&IF(AND(TO_CHAT!$L$61="ĐỦ",TO_CHAT!$I$61<>"Bất lợi
    #   "),"• "&TO_CHAT!$G$61&" ("&TO_CHAT!$H$61&") ","")&IF(AND(TO_CHAT!$L$62="ĐỦ",TO_CHAT!$I$62<>"Bất lợi"),"• "&TO_CHAT!$G$62&" ("&TO_CHAT!$H$62&") ","")&I
    #   F(AND(TO_CHAT!$L$63="ĐỦ",TO_CHAT!$I$63<>"Bất lợi"),"• "&TO_CHAT!$G$63&" ("&TO_CHAT!$H$63&") ","")))&IF(TO_CHAT!$M$16="","",CHAR(10)&"⚠ Lưu ý – sách nê
    #   u điểm bất lợi: "&TO_CHAT!$M$16)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat('(Lưu ý tuổi: chưa thành niên – nghề là định hướng xa, đối chiếu thêm khi có sở thích và học vấn.)',X.CHAR(10.0)) if X.truthy(X.le(E.v(15,5,2),18.0)) else ''),'Nghề hợp dụng thần: '),E.v(20,r-55,c)),X.CHAR(10.0)),'Nghề hợp hỷ thần: '),E.v(20,r-54,c)),('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ',E.v(20,55,7)),' ('),E.v(20,55,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,55,12),'ĐỦ'), X.ne(E.v(20,55,9),'Bất lợi'))) else ''),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,56,7)),' ('),E.v(20,56,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,56,12),'ĐỦ'), X.ne(E.v(20,56,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,57,7)),' ('),E.v(20,57,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,57,12),'ĐỦ'), X.ne(E.v(20,57,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,58,7)),' ('),E.v(20,58,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,58,12),'ĐỦ'), X.ne(E.v(20,58,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,59,7)),' ('),E.v(20,59,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,59,12),'ĐỦ'), X.ne(E.v(20,59,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,60,7)),' ('),E.v(20,60,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,60,12),'ĐỦ'), X.ne(E.v(20,60,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,61,7)),' ('),E.v(20,61,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,61,12),'ĐỦ'), X.ne(E.v(20,61,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,62,7)),' ('),E.v(20,62,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,62,12),'ĐỦ'), X.ne(E.v(20,62,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,63,7)),' ('),E.v(20,63,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,63,12),'ĐỦ'), X.ne(E.v(20,63,9),'Bất lợi'))) else ''))),'')) else X.cat(X.cat(X.CHAR(10.0),'Theo thập thần (sách): '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat('• ',E.v(20,55,7)),' ('),E.v(20,55,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,55,12),'ĐỦ'), X.ne(E.v(20,55,9),'Bất lợi'))) else ''),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,56,7)),' ('),E.v(20,56,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,56,12),'ĐỦ'), X.ne(E.v(20,56,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,57,7)),' ('),E.v(20,57,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,57,12),'ĐỦ'), X.ne(E.v(20,57,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,58,7)),' ('),E.v(20,58,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,58,12),'ĐỦ'), X.ne(E.v(20,58,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,59,7)),' ('),E.v(20,59,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,59,12),'ĐỦ'), X.ne(E.v(20,59,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,60,7)),' ('),E.v(20,60,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,60,12),'ĐỦ'), X.ne(E.v(20,60,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,61,7)),' ('),E.v(20,61,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,61,12),'ĐỦ'), X.ne(E.v(20,61,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,62,7)),' ('),E.v(20,62,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,62,12),'ĐỦ'), X.ne(E.v(20,62,9),'Bất lợi'))) else '')),(X.cat(X.cat(X.cat(X.cat('• ',E.v(20,63,7)),' ('),E.v(20,63,8)),') ') if X.truthy(X.AND(X.eq(E.v(20,63,12),'ĐỦ'), X.ne(E.v(20,63,9),'Bất lợi'))) else '')))))),('' if X.truthy(X.eq(E.v(20,16,13),'')) else X.cat(X.cat(X.CHAR(10.0),'⚠ Lưu ý – sách nêu điểm bất lợi: '),E.v(20,16,13))))

def f151(E, r, c):
    # Ô H72
    # = LUAN_GIAI!$J$188
    return E.v(4,188,10)

def f152(E, r, c):
    # Ô C73
    # = TRIM(TO_CHAT!C22&" "&TO_CHAT!C23)
    return X.TRIM(X.cat(X.cat(E.v(20,r-51,c),' '),E.v(20,r-50,c)))

def f153(E, r, c):
    # Ô H73
    # = LUAN_GIAI!$J$189
    return E.v(4,189,10)

def f154(E, r, c):
    # Ô C74
    # = SUBSTITUTE(TO_CHAT!$C$25,CHAR(10)," | ")&CHAR(10)&"Bản tính – hình dáng theo ngũ hành (P1 tr.25-26): "&TO_CHAT!$C$198&IF(TO_CHAT!$C$222="","",CHAR(10)
    #   &"Hình dáng – hoàn cảnh lúc sinh theo can ngày, giờ (P2 tr.503-504, sách: ""phải phối hợp với tướng tay và tướng mặt""): "&TO_CHAT!$C$222)
    return X.cat(X.cat(X.cat(X.cat(X.SUBSTITUTE(E.v(20,25,3), X.CHAR(10.0), ' | '),X.CHAR(10.0)),'Bản tính – hình dáng theo ngũ hành (P1 tr.25-26): '),E.v(20,198,3)),('' if X.truthy(X.eq(E.v(20,222,3),'')) else X.cat(X.cat(X.CHAR(10.0),'Hình dáng – hoàn cảnh lúc sinh theo can ngày, giờ (P2 tr.503-504, sách: "phải phối hợp với tướng tay và tướng mặt"): '),E.v(20,222,3))))

def f155(E, r, c):
    # Ô H74
    # = LUAN_GIAI!$J$122
    return E.v(4,122,10)

def f156(E, r, c):
    # Ô C75
    # = "Tố chất cốt lõi: "&C70&" Thiên hướng nghề hợp dụng thần: "&TO_CHAT!$C$17&"."
    return X.cat(X.cat(X.cat(X.cat('Tố chất cốt lõi: ',E.v(3,r-5,c)),' Thiên hướng nghề hợp dụng thần: '),E.v(20,17,3)),'.')

def f157(E, r, c):
    # Ô H75
    # = LUAN_GIAI!$J$190
    return E.v(4,190,10)

def f158(E, r, c):
    # Ô C76
    # = TRIM(IF(TRIM(IF(TO_CHAT!$L$227="ĐỦ","• "&TO_CHAT!$G$227&" (P2 tr.573; UH tr.102, 243-244) ","")&IF(TO_CHAT!$L$228="ĐỦ","• "&TO_CHAT!$G$228&" (P2 tr.57
    #   3; UH tr.168, 176) ",""))="","","Câu sách P2 (định nghĩa theo UH): "&TRIM(IF(TO_CHAT!$L$227="ĐỦ","• "&TO_CHAT!$G$227&" (P2 tr.573; UH tr.102, 243-244)
    #    ","")&IF(TO_CHAT!$L$228="ĐỦ","• "&TO_CHAT!$G$228&" (P2 tr.573; UH tr.168, 176) ",""))&" ")&IF(TRIM(IF(TO_CHAT!$L$229="ĐỦ","• "&TO_CHAT!$G$229&" (UH t
    #   r.103, 244-245) ","")&IF(TO_CHAT!$L$230="ĐỦ","• "&TO_CHAT!$G$230&" (UH tr.101-102, 242) ","")&IF(TO_CHAT!$L$231="ĐỦ","• "&TO_CHAT!$G$231&" (UH tr.102-
    #   103) ",""))="","",TRIM(IF(TO_CHAT!$L$229="ĐỦ","• "&TO_CHAT!$G$229&" (UH tr.103, 244-245) ","")&IF(TO_CHAT!$L$230="ĐỦ","• "&TO_CHAT!$G$230&" (UH tr.101
    #   -102, 242) ","")&IF(TO_CHAT!$L$231="ĐỦ","• "&TO_CHAT!$G$231&" (UH tr.102-103) ",""))))
    return X.TRIM(X.cat(('' if X.truthy(X.eq(X.TRIM(X.cat((X.cat(X.cat('• ',E.v(20,227,7)),' (P2 tr.573; UH tr.102, 243-244) ') if X.truthy(X.eq(E.v(20,227,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,228,7)),' (P2 tr.573; UH tr.168, 176) ') if X.truthy(X.eq(E.v(20,228,12),'ĐỦ')) else ''))),'')) else X.cat(X.cat('Câu sách P2 (định nghĩa theo UH): ',X.TRIM(X.cat((X.cat(X.cat('• ',E.v(20,227,7)),' (P2 tr.573; UH tr.102, 243-244) ') if X.truthy(X.eq(E.v(20,227,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,228,7)),' (P2 tr.573; UH tr.168, 176) ') if X.truthy(X.eq(E.v(20,228,12),'ĐỦ')) else '')))),' ')),('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,229,7)),' (UH tr.103, 244-245) ') if X.truthy(X.eq(E.v(20,229,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,230,7)),' (UH tr.101-102, 242) ') if X.truthy(X.eq(E.v(20,230,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,231,7)),' (UH tr.102-103) ') if X.truthy(X.eq(E.v(20,231,12),'ĐỦ')) else ''))),'')) else X.TRIM(X.cat(X.cat((X.cat(X.cat('• ',E.v(20,229,7)),' (UH tr.103, 244-245) ') if X.truthy(X.eq(E.v(20,229,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(20,230,7)),' (UH tr.101-102, 242) ') if X.truthy(X.eq(E.v(20,230,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(20,231,7)),' (UH tr.102-103) ') if X.truthy(X.eq(E.v(20,231,12),'ĐỦ')) else ''))))))

def f159(E, r, c):
    # Ô C79
    # = DUNG_HY_KY!$C$188&CHAR(10)&C86&DUNG_HY_KY!$D$192
    return X.cat(X.cat(X.cat(E.v(8,188,3),X.CHAR(10.0)),E.v(3,r+7,c)),E.v(8,192,4))

def f160(E, r, c):
    # Ô C81
    # = TAI_LOC!E17&". "&TAI_LOC!F16
    return X.cat(X.cat(E.v(19,r-64,c+2),'. '),E.v(19,r-65,c+3))

def f161(E, r, c):
    # Ô H81
    # = LUAN_GIAI!$J$182
    return E.v(4,182,10)

def f162(E, r, c):
    # Ô C82
    # = "Độ sao tài / tổng độ (PP2): "&TEXT(TAI_LOC!C6,"0.0")&" / "&TEXT(TAI_LOC!C7,"0.0")&" = "&TEXT(TAI_LOC!C8,"0.0%")&"; phe mình "&TEXT(TAI_LOC!C9,"0.0%")
    #   &"; tỉ số tài : phe mình "&IF(ISNUMBER(TAI_LOC!C10),TEXT(TAI_LOC!C10,"0.00"),"—")&"; tài đứng hạng "&TAI_LOC!C11&"/5. PP1: "&TAI_LOC!D6&"/"&TAI_LOC!D7
    #   &" chữ = "&TEXT(TAI_LOC!D8,"0%")&". Sách không cho ngưỡng quy số này ra mức giàu."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Độ sao tài / tổng độ (PP2): ',X.TEXT(E.v(19,r-76,c), '0.0')),' / '),X.TEXT(E.v(19,r-75,c), '0.0')),' = '),X.TEXT(E.v(19,r-74,c), '0.0%')),'; phe mình '),X.TEXT(E.v(19,r-73,c), '0.0%')),'; tỉ số tài : phe mình '),(X.TEXT(E.v(19,r-72,c), '0.00') if X.truthy(X.ISNUMBER(lambda: E.v(19,r-72,c))) else '—')),'; tài đứng hạng '),E.v(19,r-71,c)),'/5. PP1: '),E.v(19,r-76,c+1)),'/'),E.v(19,r-75,c+1)),' chữ = '),X.TEXT(E.v(19,r-74,c+1), '0%')),'. Sách không cho ngưỡng quy số này ra mức giàu.')

def f163(E, r, c):
    # Ô H82
    # = LUAN_GIAI!$J$183
    return E.v(4,183,10)

def f164(E, r, c):
    # Ô C83
    # = IF(TAI_LOC!C21="","Không có dòng nguồn tiền nào của sách ứng với lá số",TAI_LOC!C21)&IF(TAI_LOC!C22="",""," Bất lợi: "&TAI_LOC!C22)
    return X.cat(('Không có dòng nguồn tiền nào của sách ứng với lá số' if X.truthy(X.eq(E.v(19,r-62,c),'')) else E.v(19,r-62,c)),('' if X.truthy(X.eq(E.v(19,r-61,c),'')) else X.cat(' Bất lợi: ',E.v(19,r-61,c))))

def f165(E, r, c):
    # Ô H83
    # = LUAN_GIAI!$J$184
    return E.v(4,184,10)

def f166(E, r, c):
    # Ô C84
    # = "Thuận: "&IF(TAI_LOC!C25="","—",TAI_LOC!C25)&CHAR(10)&"Bất lợi: "&IF(TAI_LOC!C26="","—",TAI_LOC!C26)
    return X.cat(X.cat(X.cat(X.cat('Thuận: ',('—' if X.truthy(X.eq(E.v(19,r-59,c),'')) else E.v(19,r-59,c))),X.CHAR(10.0)),'Bất lợi: '),('—' if X.truthy(X.eq(E.v(19,r-58,c),'')) else E.v(19,r-58,c)))

def f167(E, r, c):
    # Ô H84
    # = LUAN_GIAI!$J$185
    return E.v(4,185,10)

def f168(E, r, c):
    # Ô C85
    # = "Dễ phát tài: "&IF(TAI_LOC!C37="","không có đại vận nào thỏa điều kiện",TAI_LOC!C37)&CHAR(10)&"Cẩn trọng: "&IF(TAI_LOC!C38="","—",TAI_LOC!C38)
    return X.cat(X.cat(X.cat(X.cat('Dễ phát tài: ',('không có đại vận nào thỏa điều kiện' if X.truthy(X.eq(E.v(19,r-48,c),'')) else E.v(19,r-48,c))),X.CHAR(10.0)),'Cẩn trọng: '),('—' if X.truthy(X.eq(E.v(19,r-47,c),'')) else E.v(19,r-47,c)))

def f169(E, r, c):
    # Ô H85
    # = LUAN_GIAI!$J$186
    return E.v(4,186,10)

def f170(E, r, c):
    # Ô C86
    # = "Bậc giàu: "&TAI_LOC!E17&". Tài vận (chương 18): "&LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LEN(SU_NGHIEP!$J$6)))&" – tài "&IF(MOD(
    #   TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY
    #   _KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"là kỵ thần",IF(AND(S
    #   U_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>I
    #   F(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)","không phải dụng/hỷ/kỵ thần (nhà
    #   n thần)"))))&". "&SUBSTITUTE(SU_NGHIEP!$C$22,CHAR(10),"; ")&"."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bậc giàu: ',E.v(19,r-69,c+2)),'. Tài vận (chương 18): '),X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10))))),' – tài '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),'. '),X.SUBSTITUTE(E.v(21,22,3), X.CHAR(10.0), '; ')),'.')

def f171(E, r, c):
    # Ô H86
    # = LUAN_GIAI!$J$120
    return E.v(4,120,10)

def f172(E, r, c):
    # Ô B87
    # = MENH_CUC!B18
    return E.v(18,r-69,c)

def f173(E, r, c):
    # Ô C87
    # = MENH_CUC!D18&CHAR(10)&MENH_CUC!G18&CHAR(10)&MENH_CUC!K18&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(MENH_CUC!$U$18="NGHÈO","kiếm tiền vất vả, làm ra khó giữ (tà
    #   i là kỵ, hoặc tài là dụng mà bị phá); gặp vận, năm tài vượng mà thân không gánh nổi thì dễ hao tán – “lúc gặp là tai họa” (P1 tr.166-167). Không phải 
    #   túng thiếu suốt đời: trong vận tốt đời sống vẫn khá lên",IF(MENH_CUC!$U$18="GIÀU","tiền của dồi dào, kiếm được và giữ được của; tài là dụng/hỷ có lực 
    #   nên gặp vận, năm tài thì phát (P1 tr.164-165)","cân bằng: hai chiều ngang nhau, tốt xấu tùy vận"))&". "&IF(AND($AB$87="",$AC$87=""),"Khi nào thể hiện 
    #   rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo
    #    phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu ("&IF($AD$87="","không có",$AD$87)&"); phần xấu dễ nổi hơn trong bại vận
    #    ("&IF($AE$87="","không có",$AE$87)&"). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.","Khi nào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT ("&IF($
    #   AB$87="","không có",$AB$87)&"); phần xấu nổi lên trong đại vận XẤU ("&IF($AC$87="","không có",$AC$87)&").")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(18,r-69,c+1),X.CHAR(10.0)),E.v(18,r-69,c+4)),X.CHAR(10.0)),E.v(18,r-69,c+8)),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),('kiếm tiền vất vả, làm ra khó giữ (tài là kỵ, hoặc tài là dụng mà bị phá); gặp vận, năm tài vượng mà thân không gánh nổi thì dễ hao tán – “lúc gặp là tai họa” (P1 tr.166-167). Không phải túng thiếu suốt đời: trong vận tốt đời sống vẫn khá lên' if X.truthy(X.eq(E.v(18,18,21),'NGHÈO')) else ('tiền của dồi dào, kiếm được và giữ được của; tài là dụng/hỷ có lực nên gặp vận, năm tài thì phát (P1 tr.164-165)' if X.truthy(X.eq(E.v(18,18,21),'GIÀU')) else 'cân bằng: hai chiều ngang nhau, tốt xấu tùy vận'))),'. '),(X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu (',('không có' if X.truthy(X.eq(E.v(3,87,30),'')) else E.v(3,87,30))),'); phần xấu dễ nổi hơn trong bại vận ('),('không có' if X.truthy(X.eq(E.v(3,87,31),'')) else E.v(3,87,31))),'). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.') if X.truthy(X.AND(X.eq(E.v(3,87,28),''), X.eq(E.v(3,87,29),''))) else X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT (',('không có' if X.truthy(X.eq(E.v(3,87,28),'')) else E.v(3,87,28))),'); phần xấu nổi lên trong đại vận XẤU ('),('không có' if X.truthy(X.eq(E.v(3,87,29),'')) else E.v(3,87,29))),').')))

def f174(E, r, c):
    # Ô H87
    # = LUAN_GIAI!$J$116
    return E.v(4,116,10)

def f175(E, r, c):
    # Ô AB87
    # = TRIM(IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C69,DAI_VAN!C54)="Tốt",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI
    #   _VAN!D69,DAI_VAN!D54)="Tốt",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E69,DAI_VAN!E54)="Tốt",DAI_
    #   VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F69,DAI_VAN!F54)="Tốt",DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DA
    #   I_VAN!F36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G69,DAI_VAN!G54)="Tốt",DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(IF
    #   (DOI_CHIEU!$C$33=2,DAI_VAN!H69,DAI_VAN!H54)="Tốt",DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I69,D
    #   AI_VAN!I54)="Tốt",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J69,DAI_VAN!J54)="Tốt",DAI_VAN!J35&" 
    #   ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-25),' ('),E.v(14,r-51,c-25)),'–'),X.add(E.v(14,r-51,c-25),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-25) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-25)),'Tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-24),' ('),E.v(14,r-51,c-24)),'–'),X.add(E.v(14,r-51,c-24),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-24) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-24)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-23),' ('),E.v(14,r-51,c-23)),'–'),X.add(E.v(14,r-51,c-23),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-23) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-23)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-22),' ('),E.v(14,r-51,c-22)),'–'),X.add(E.v(14,r-51,c-22),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-22) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-22)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-21),' ('),E.v(14,r-51,c-21)),'–'),X.add(E.v(14,r-51,c-21),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-21) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-21)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-20),' ('),E.v(14,r-51,c-20)),'–'),X.add(E.v(14,r-51,c-20),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-20) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-20)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-19),' ('),E.v(14,r-51,c-19)),'–'),X.add(E.v(14,r-51,c-19),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-19) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-19)),'Tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-18),' ('),E.v(14,r-51,c-18)),'–'),X.add(E.v(14,r-51,c-18),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-18) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-18)),'Tốt')) else '')))

def f176(E, r, c):
    # Ô AC87
    # = TRIM(IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C69,DAI_VAN!C54)="Xấu",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI
    #   _VAN!D69,DAI_VAN!D54)="Xấu",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!E69,DAI_VAN!E54)="Xấu",DAI_
    #   VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F69,DAI_VAN!F54)="Xấu",DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DA
    #   I_VAN!F36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G69,DAI_VAN!G54)="Xấu",DAI_VAN!G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(IF
    #   (DOI_CHIEU!$C$33=2,DAI_VAN!H69,DAI_VAN!H54)="Xấu",DAI_VAN!H35&" ("&DAI_VAN!H36&"–"&(DAI_VAN!H36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I69,D
    #   AI_VAN!I54)="Xấu",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)&" tuổi) ","")&IF(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J69,DAI_VAN!J54)="Xấu",DAI_VAN!J35&" 
    #   ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-26),' ('),E.v(14,r-51,c-26)),'–'),X.add(E.v(14,r-51,c-26),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-26) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-26)),'Xấu')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-25),' ('),E.v(14,r-51,c-25)),'–'),X.add(E.v(14,r-51,c-25),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-25) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-25)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-24),' ('),E.v(14,r-51,c-24)),'–'),X.add(E.v(14,r-51,c-24),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-24) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-24)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-23),' ('),E.v(14,r-51,c-23)),'–'),X.add(E.v(14,r-51,c-23),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-23) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-23)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-22),' ('),E.v(14,r-51,c-22)),'–'),X.add(E.v(14,r-51,c-22),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-22) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-22)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-21),' ('),E.v(14,r-51,c-21)),'–'),X.add(E.v(14,r-51,c-21),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-21) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-21)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-20),' ('),E.v(14,r-51,c-20)),'–'),X.add(E.v(14,r-51,c-20),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-20) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-20)),'Xấu')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-19),' ('),E.v(14,r-51,c-19)),'–'),X.add(E.v(14,r-51,c-19),9.0)),' tuổi) ') if X.truthy(X.eq((E.v(14,r-18,c-19) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-33,c-19)),'Xấu')) else '')))

def f177(E, r, c):
    # Ô AD87
    # = TRIM(IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C71,DAI_VAN!C56),7)="vận tốt",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_
    #   CHIEU!$C$33=2,DAI_VAN!D71,DAI_VAN!D56),7)="vận tốt",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VA
    #   N!E71,DAI_VAN!E56),7)="vận tốt",DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F71,DAI_VAN!F56),7
    #   )="vận tốt",DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G71,DAI_VAN!G56),7)="vận tốt",DAI_VAN!
    #   G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!H71,DAI_VAN!H56),7)="vận tốt",DAI_VAN!H35&" ("&DAI_VAN!H36
    #   &"–"&(DAI_VAN!H36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I71,DAI_VAN!I56),7)="vận tốt",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)
    #   &" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J71,DAI_VAN!J56),7)="vận tốt",DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-27),' ('),E.v(14,r-51,c-27)),'–'),X.add(E.v(14,r-51,c-27),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-27) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-27)), 7.0),'vận tốt')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-26),' ('),E.v(14,r-51,c-26)),'–'),X.add(E.v(14,r-51,c-26),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-26) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-26)), 7.0),'vận tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-25),' ('),E.v(14,r-51,c-25)),'–'),X.add(E.v(14,r-51,c-25),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-25) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-25)), 7.0),'vận tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-24),' ('),E.v(14,r-51,c-24)),'–'),X.add(E.v(14,r-51,c-24),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-24) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-24)), 7.0),'vận tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-23),' ('),E.v(14,r-51,c-23)),'–'),X.add(E.v(14,r-51,c-23),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-23) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-23)), 7.0),'vận tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-22),' ('),E.v(14,r-51,c-22)),'–'),X.add(E.v(14,r-51,c-22),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-22) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-22)), 7.0),'vận tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-21),' ('),E.v(14,r-51,c-21)),'–'),X.add(E.v(14,r-51,c-21),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-21) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-21)), 7.0),'vận tốt')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-20),' ('),E.v(14,r-51,c-20)),'–'),X.add(E.v(14,r-51,c-20),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-20) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-20)), 7.0),'vận tốt')) else '')))

def f178(E, r, c):
    # Ô AE87
    # = TRIM(IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!C71,DAI_VAN!C56),7)="bại vận",DAI_VAN!C35&" ("&DAI_VAN!C36&"–"&(DAI_VAN!C36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_
    #   CHIEU!$C$33=2,DAI_VAN!D71,DAI_VAN!D56),7)="bại vận",DAI_VAN!D35&" ("&DAI_VAN!D36&"–"&(DAI_VAN!D36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VA
    #   N!E71,DAI_VAN!E56),7)="bại vận",DAI_VAN!E35&" ("&DAI_VAN!E36&"–"&(DAI_VAN!E36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!F71,DAI_VAN!F56),7
    #   )="bại vận",DAI_VAN!F35&" ("&DAI_VAN!F36&"–"&(DAI_VAN!F36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!G71,DAI_VAN!G56),7)="bại vận",DAI_VAN!
    #   G35&" ("&DAI_VAN!G36&"–"&(DAI_VAN!G36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!H71,DAI_VAN!H56),7)="bại vận",DAI_VAN!H35&" ("&DAI_VAN!H36
    #   &"–"&(DAI_VAN!H36+9)&" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!I71,DAI_VAN!I56),7)="bại vận",DAI_VAN!I35&" ("&DAI_VAN!I36&"–"&(DAI_VAN!I36+9)
    #   &" tuổi) ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,DAI_VAN!J71,DAI_VAN!J56),7)="bại vận",DAI_VAN!J35&" ("&DAI_VAN!J36&"–"&(DAI_VAN!J36+9)&" tuổi) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-28),' ('),E.v(14,r-51,c-28)),'–'),X.add(E.v(14,r-51,c-28),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-28) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-28)), 7.0),'bại vận')) else ''),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-27),' ('),E.v(14,r-51,c-27)),'–'),X.add(E.v(14,r-51,c-27),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-27) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-27)), 7.0),'bại vận')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-26),' ('),E.v(14,r-51,c-26)),'–'),X.add(E.v(14,r-51,c-26),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-26) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-26)), 7.0),'bại vận')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-25),' ('),E.v(14,r-51,c-25)),'–'),X.add(E.v(14,r-51,c-25),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-25) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-25)), 7.0),'bại vận')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-24),' ('),E.v(14,r-51,c-24)),'–'),X.add(E.v(14,r-51,c-24),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-24) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-24)), 7.0),'bại vận')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-23),' ('),E.v(14,r-51,c-23)),'–'),X.add(E.v(14,r-51,c-23),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-23) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-23)), 7.0),'bại vận')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-22),' ('),E.v(14,r-51,c-22)),'–'),X.add(E.v(14,r-51,c-22),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-22) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-22)), 7.0),'bại vận')) else '')),(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(14,r-52,c-21),' ('),E.v(14,r-51,c-21)),'–'),X.add(E.v(14,r-51,c-21),9.0)),' tuổi) ') if X.truthy(X.eq(X.LEFT((E.v(14,r-16,c-21) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(14,r-31,c-21)), 7.0),'bại vận')) else '')))

def f179(E, r, c):
    # Ô C88
    # = SU_NGHIEP!$J$6&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(LEFT(LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6&CHAR(10))-1),7)="BẤT LỢI","tiền khó kiếm hoặc là
    #   m ra rồi hao, dễ vì tiền mà phiền lụy"&IF(ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6&CHAR(10))-1))),IF(AND(SU_NGHIEP
    #   !$M$6>=2*SU_NGHIEP!$N$6,SU_NGHIEP!$M$6>=4,TRUE())," – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có 
    #   chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hi
    #   ệu thuận ở trên), thể hiện trong vận tốt"),""),IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!$J$6&CHAR(10))-1))),"hơi b
    #   ất lợi – tiền khó kiếm hoặc làm ra rồi hao, dễ vì tiền mà phiền lụy, ở mức nhẹ",IF(ISNUMBER(SEARCH("TRUNG BÌNH",LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_N
    #   GHIEP!$J$6&CHAR(10))-1))),"bình thường – lúc thuận lúc không, tùy vận",IF(ISNUMBER(SEARCH("NGHIÊNG THUẬN",LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGHIEP!
    #   $J$6&CHAR(10))-1))),"khá – kiếm tiền thuận, có nguồn thu ổn định và có thể có thêm nguồn phụ, giữ được của, ở mức vừa, không phải thế mạnh nổi bật","k
    #   iếm tiền thuận, có nguồn thu ổn định và có thể có thêm nguồn phụ, giữ được của"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SU_NGHIEP!$J$6,FIND(CHAR(10),SU_NGH
    #   IEP!$J$6&CHAR(10))-1)))," – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)","")))))&"."&CHAR(10)&"Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: "&SU
    #   _NGHIEP!$C$6&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điều kiện: "&SU_NGHIEP!$F$6&CHAR(10)&"(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)"&CHAR(10)&SU_NGHIEP
    #   !$C$22
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(21,6,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('tiền khó kiếm hoặc làm ra rồi hao, dễ vì tiền mà phiền lụy',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(21,6,13),X.mul(2.0,E.v(21,6,14))), X.ge(E.v(21,6,13),4.0), X.TRUE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,6,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,6,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – tiền khó kiếm hoặc làm ra rồi hao, dễ vì tiền mà phiền lụy, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,6,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,6,10),X.CHAR(10.0))),1.0))))) else ('khá – kiếm tiền thuận, có nguồn thu ổn định và có thể có thêm nguồn phụ, giữ được của, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,6,10),X.CHAR(10.0))),1.0))))) else X.cat('kiếm tiền thuận, có nguồn thu ổn định và có thể có thêm nguồn phụ, giữ được của',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(21,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,6,10),X.CHAR(10.0))),1.0))))) else ''))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(21,6,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(21,6,6)),X.CHAR(10.0)),'(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)'),X.CHAR(10.0)),E.v(21,22,3))

def f180(E, r, c):
    # Ô H88
    # = LUAN_GIAI!$J$123
    return E.v(4,123,10)

def f181(E, r, c):
    # Ô C89
    # = "Chưa rút được quy luật riêng từ các ví dụ (xem THAM_KHAO bảng C). Dùng kết luận chính. Riêng tiền bạc từ phía vợ/chồng: xem dòng Hôn nhân (TK4)."
    return 'Chưa rút được quy luật riêng từ các ví dụ (xem THAM_KHAO bảng C). Dùng kết luận chính. Riêng tiền bạc từ phía vợ/chồng: xem dòng Hôn nhân (TK4).'

def f182(E, r, c):
    # Ô H89
    # = LUAN_GIAI!$J$207
    return E.v(4,207,10)

def f183(E, r, c):
    # Ô C90
    # = TRIM(IF(TRIM(IF(TAI_LOC!$L$84="ĐỦ","• "&TAI_LOC!$G$84&" (P2 tr.521; UH tr.239) ","")&IF(TAI_LOC!$L$85="ĐỦ","• "&TAI_LOC!$G$85&" (P2 tr.522; UH tr.173)
    #    ",""))="","","Câu sách P2 (định nghĩa theo UH): "&TRIM(IF(TAI_LOC!$L$84="ĐỦ","• "&TAI_LOC!$G$84&" (P2 tr.521; UH tr.239) ","")&IF(TAI_LOC!$L$85="ĐỦ",
    #   "• "&TAI_LOC!$G$85&" (P2 tr.522; UH tr.173) ",""))&" ")&IF(""="","",""))
    return X.TRIM(X.cat(('' if X.truthy(X.eq(X.TRIM(X.cat((X.cat(X.cat('• ',E.v(19,84,7)),' (P2 tr.521; UH tr.239) ') if X.truthy(X.eq(E.v(19,84,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(19,85,7)),' (P2 tr.522; UH tr.173) ') if X.truthy(X.eq(E.v(19,85,12),'ĐỦ')) else ''))),'')) else X.cat(X.cat('Câu sách P2 (định nghĩa theo UH): ',X.TRIM(X.cat((X.cat(X.cat('• ',E.v(19,84,7)),' (P2 tr.521; UH tr.239) ') if X.truthy(X.eq(E.v(19,84,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(19,85,7)),' (P2 tr.522; UH tr.173) ') if X.truthy(X.eq(E.v(19,85,12),'ĐỦ')) else '')))),' ')),('' if X.truthy(X.eq('','')) else '')))

def f184(E, r, c):
    # Ô C93
    # = "Sang – hèn (P1 tr.173): "&LEFT(C99,IFERROR(FIND(CHAR(10),C99)-1,LEN(C99)))&". "&C98
    return X.cat(X.cat(X.cat('Sang – hèn (P1 tr.173): ',X.LEFT(E.v(3,r+6,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+6,c)),1.0), lambda: X.LEN(E.v(3,r+6,c))))),'. '),E.v(3,r+5,c))

def f185(E, r, c):
    # Ô C95
    # = SU_NGHIEP!$J$7&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(ISNUMBER(SEARCH("SÁT CÓ CHẾ",SU_NGHIEP!$J$7)),"sao quan sát vốn là kỵ thần nhưng đã bị khắc, tiết, hợp
    #    chế ngự nên chuyển thành có thể dùng: hợp con đường chức vụ, công chức, quản lý – có thể giữ chức, được cất nhắc nhưng tiến từng bước, rõ nhất trong 
    #   các vận thuận nêu dưới; năm, vận là"&"m quan sát mạnh lên (vận Kim, vận Thổ sinh Kim) vẫn dễ bị áp lực, chèn ép",IF(LEFT(LEFT(SU_NGHIEP!$J$7,FIND(CHAR
    #   (10),SU_NGHIEP!$J$7&CHAR(10))-1),7)="BẤT LỢI","khó thăng tiến, dễ mất chức hoặc bị cấp trên chèn ép"&IF(ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SU_NGHIEP!$J$
    #   7,FIND(CHAR(10),SU_NGHIEP!$J$7&CHAR(10))-1))),IF(AND(SU_NGHIEP!$M$7>=2*SU_NGHIEP!$N$7,SU_NGHIEP!$M$7>=4,TRUE())," – nhưng điểm sáng lớn: dấu hiệu thuậ
    #   n gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đ
    #   úng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt"),""),IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SU_
    #   NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7&CHAR(10))-1))),"hơi bất lợi – khó thăng tiến, dễ mất chức hoặc bị cấp trên chèn ép, ở mức nhẹ",IF(ISNUMBER(SE
    #   ARCH("TRUNG BÌNH",LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7&CHAR(10))-1))),"bình thường – lúc thuận lúc không, tùy vận",IF(ISNUMBER(SEARCH("NGH
    #   IÊNG THUẬN",LEFT(SU_NGHIEP!$J$7,FIND(CHAR(10),SU_NGHIEP!$J$7&CHAR(10))-1))),"khá – có chức vụ, được cất nhắc, công việc có vị trí và quyền hạn, ở mức 
    #   vừa, không phải thế mạnh nổi bật","có chức vụ, được cất nhắc, công việc có vị trí và quyền hạn"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SU_NGHIEP!$J$7,FIND
    #   (CHAR(10),SU_NGHIEP!$J$7&CHAR(10))-1)))," – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)",""))))))&"."&CHAR(10)&"Vì sao – dấu hiệu THUẬN của sách đ
    #   ủ điều kiện: "&SU_NGHIEP!$C$7&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điều kiện: "&SU_NGHIEP!$F$7&CHAR(10)&"(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)"&C
    #   HAR(10)&SU_NGHIEP!$C$23
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(21,7,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('sao quan sát vốn là kỵ thần nhưng đã bị khắc, tiết, hợp chế ngự nên chuyển thành có thể dùng: hợp con đường chức vụ, công chức, quản lý – có thể giữ chức, được cất nhắc nhưng tiến từng bước, rõ nhất trong các vận thuận nêu dưới; năm, vận là','m quan sát mạnh lên (vận Kim, vận Thổ sinh Kim) vẫn dễ bị áp lực, chèn ép') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('SÁT CÓ CHẾ', E.v(21,7,10)))) else (X.cat('khó thăng tiến, dễ mất chức hoặc bị cấp trên chèn ép',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(21,7,13),X.mul(2.0,E.v(21,7,14))), X.ge(E.v(21,7,13),4.0), X.TRUE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,7,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,7,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – khó thăng tiến, dễ mất chức hoặc bị cấp trên chèn ép, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,7,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,7,10),X.CHAR(10.0))),1.0))))) else ('khá – có chức vụ, được cất nhắc, công việc có vị trí và quyền hạn, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,7,10),X.CHAR(10.0))),1.0))))) else X.cat('có chức vụ, được cất nhắc, công việc có vị trí và quyền hạn',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(21,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,7,10),X.CHAR(10.0))),1.0))))) else '')))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(21,7,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(21,7,6)),X.CHAR(10.0)),'(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)'),X.CHAR(10.0)),E.v(21,23,3))

def f186(E, r, c):
    # Ô H95
    # = LUAN_GIAI!$J$124
    return E.v(4,124,10)

def f187(E, r, c):
    # Ô C96
    # = SU_NGHIEP!$J$8&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(LEFT(LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8&CHAR(10))-1),7)="BẤT LỢI","học hành gián đoạn, t
    #   hi cử lận đận"&IF(ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8&CHAR(10))-1))),IF(AND(SU_NGHIEP!$M$8>=2*SU_NGHIEP!$N$8,
    #   SU_NGHIEP!$M$8>=4,TRUE())," – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 
    #   tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hi
    #   ện trong vận tốt"),""),IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8&CHAR(10))-1))),"hơi bất lợi – học hành gián đ
    #   oạn, thi cử lận đận, ở mức nhẹ",IF(ISNUMBER(SEARCH("TRUNG BÌNH",LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8&CHAR(10))-1))),"bình thường – lúc thu
    #   ận lúc không, tùy vận",IF(ISNUMBER(SEARCH("NGHIÊNG THUẬN",LEFT(SU_NGHIEP!$J$8,FIND(CHAR(10),SU_NGHIEP!$J$8&CHAR(10))-1))),"khá – học giỏi, thi cử đỗ đ
    #   ạt, có bằng cấp, ở mức vừa, không phải thế mạnh nổi bật","học giỏi, thi cử đỗ đạt, có bằng cấp"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SU_NGHIEP!$J$8,FIND
    #   (CHAR(10),SU_NGHIEP!$J$8&CHAR(10))-1)))," – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)","")))))&"."&CHAR(10)&"Vì sao – dấu hiệu THUẬN của sách đủ
    #    điều kiện: "&SU_NGHIEP!$C$8&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điều kiện: "&SU_NGHIEP!$F$8&CHAR(10)&"(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(21,8,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('học hành gián đoạn, thi cử lận đận',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(21,8,13),X.mul(2.0,E.v(21,8,14))), X.ge(E.v(21,8,13),4.0), X.TRUE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,8,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,8,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – học hành gián đoạn, thi cử lận đận, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,8,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,8,10),X.CHAR(10.0))),1.0))))) else ('khá – học giỏi, thi cử đỗ đạt, có bằng cấp, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,8,10),X.CHAR(10.0))),1.0))))) else X.cat('học giỏi, thi cử đỗ đạt, có bằng cấp',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(21,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,8,10),X.CHAR(10.0))),1.0))))) else ''))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(21,8,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(21,8,6)),X.CHAR(10.0)),'(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)')

def f188(E, r, c):
    # Ô H96
    # = LUAN_GIAI!$J$125
    return E.v(4,125,10)

def f189(E, r, c):
    # Ô C97
    # = SU_NGHIEP!$J$9&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(LEFT(LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9&CHAR(10))-1),7)="BẤT LỢI","khó nổi danh, dễ mang
    #    tiếng"&IF(ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9&CHAR(10))-1))),IF(AND(SU_NGHIEP!$M$9>=2*SU_NGHIEP!$N$9,SU_NGHI
    #   EP!$M$9>=4,TRUE())," – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535)
    #   , nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện tron
    #   g vận tốt"),""),IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9&CHAR(10))-1))),"hơi bất lợi – khó nổi danh, dễ mang 
    #   tiếng, ở mức nhẹ",IF(ISNUMBER(SEARCH("TRUNG BÌNH",LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9&CHAR(10))-1))),"bình thường – lúc thuận lúc không, 
    #   tùy vận",IF(ISNUMBER(SEARCH("NGHIÊNG THUẬN",LEFT(SU_NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9&CHAR(10))-1))),"khá – có tiếng tăm, được nhiều người biết
    #    đến và nể trọng, ở mức vừa, không phải thế mạnh nổi bật","có tiếng tăm, được nhiều người biết đến và nể trọng"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SU_
    #   NGHIEP!$J$9,FIND(CHAR(10),SU_NGHIEP!$J$9&CHAR(10))-1)))," – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)","")))))&"."&CHAR(10)&"Vì sao – dấu hiệu T
    #   HUẬN của sách đủ điều kiện: "&SU_NGHIEP!$C$9&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điều kiện: "&SU_NGHIEP!$F$9&CHAR(10)&"(Cách kết luận: SU_NGHIEP dòng 10, gi
    #   ả định G17.)"&CHAR(10)&SU_NGHIEP!$C$24
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(21,9,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('khó nổi danh, dễ mang tiếng',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(21,9,13),X.mul(2.0,E.v(21,9,14))), X.ge(E.v(21,9,13),4.0), X.TRUE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,9,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,9,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – khó nổi danh, dễ mang tiếng, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,9,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,9,10),X.CHAR(10.0))),1.0))))) else ('khá – có tiếng tăm, được nhiều người biết đến và nể trọng, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,9,10),X.CHAR(10.0))),1.0))))) else X.cat('có tiếng tăm, được nhiều người biết đến và nể trọng',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(21,9,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(21,9,10),X.CHAR(10.0))),1.0))))) else ''))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(21,9,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(21,9,6)),X.CHAR(10.0)),'(Cách kết luận: SU_NGHIEP dòng 10, giả định G17.)'),X.CHAR(10.0)),E.v(21,24,3))

def f190(E, r, c):
    # Ô H97
    # = LUAN_GIAI!$J$126
    return E.v(4,126,10)

def f191(E, r, c):
    # Ô C98
    # = SU_NGHIEP!$C$27&CHAR(10)&"Cách đọc mức (giả định G17): THUẬN = sao chủ sự là dụng/hỷ thần và dấu hiệu thuận của sách nhiều hơn; THUẬN NHƯNG CÓ TRỞ NGẠ
    #   I = sao chủ sự tốt nhưng dấu hiệu bất lợi nhiều hơn; NGHIÊNG THUẬN = sao chủ sự là nhàn thần, dấu hiệu thuận nhiều h"&"ơn; TRUNG BÌNH = ngang nhau; NG
    #   HIÊNG BẤT LỢI = nhàn thần, dấu hiệu bất lợi nhiều hơn; BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG = sao chủ sự là kỵ nhưng dấu hiệu thuận nhiều hơn; BẤT LỢI = sao chủ
    #    sự là kỵ (kể cả là nhóm nhiều nhất thuộc thực thương/tài/q"&"uan sát – P1 tr.180, giả định G26) và dấu hiệu bất lợi nhiều hơn."
    return X.cat(X.cat(X.cat(X.cat(E.v(21,27,3),X.CHAR(10.0)),'Cách đọc mức (giả định G17): THUẬN = sao chủ sự là dụng/hỷ thần và dấu hiệu thuận của sách nhiều hơn; THUẬN NHƯNG CÓ TRỞ NGẠI = sao chủ sự tốt nhưng dấu hiệu bất lợi nhiều hơn; NGHIÊNG THUẬN = sao chủ sự là nhàn thần, dấu hiệu thuận nhiều h'),'ơn; TRUNG BÌNH = ngang nhau; NGHIÊNG BẤT LỢI = nhàn thần, dấu hiệu bất lợi nhiều hơn; BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG = sao chủ sự là kỵ nhưng dấu hiệu thuận nhiều hơn; BẤT LỢI = sao chủ sự là kỵ (kể cả là nhóm nhiều nhất thuộc thực thương/tài/q'),'uan sát – P1 tr.180, giả định G26) và dấu hiệu bất lợi nhiều hơn.')

def f192(E, r, c):
    # Ô H98
    # = LUAN_GIAI!$J$191
    return E.v(4,191,10)

def f193(E, r, c):
    # Ô B99
    # = MENH_CUC!B19
    return E.v(18,r-80,c)

def f194(E, r, c):
    # Ô C99
    # = MENH_CUC!D19&CHAR(10)&MENH_CUC!G19&CHAR(10)&MENH_CUC!K19&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(MENH_CUC!$U$19="HÈN","địa vị thấp, ít quyền hành, khó thăng 
    #   tiến, thường phải làm việc dưới quyền người khác hoặc bị chèn ép; nguyên nhân là “quan sát không được trung hòa, nặng quá hoặc nhẹ quá” (P1 tr.169-170
    #   )",IF(MENH_CUC!$U$19="SANG","có địa vị, chức phận, danh tiếng, được người trên trọng dụng; “quan gặp tuế vận là có quan chức” – thăng tiến rõ trong vậ
    #   n tốt (P1 tr.167-169)","cân bằng: hai chiều ngang nhau, tốt xấu tùy vận"))&". "&IF(AND($AB$87="",$AC$87=""),"Khi nào thể hiện rõ: sau điều chỉnh (vận 
    #   tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125
    #   : phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu ("&IF($AD$87="","không có",$AD$87)&"); phần xấu dễ nổi hơn trong bại vận ("&IF($AE$87="","không 
    #   có",$AE$87)&"). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.","Khi nào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT ("&IF($AB$87="","không có",$AB$
    #   87)&"); phần xấu nổi lên trong đại vận XẤU ("&IF($AC$87="","không có",$AC$87)&").")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(18,r-80,c+1),X.CHAR(10.0)),E.v(18,r-80,c+4)),X.CHAR(10.0)),E.v(18,r-80,c+8)),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),('địa vị thấp, ít quyền hành, khó thăng tiến, thường phải làm việc dưới quyền người khác hoặc bị chèn ép; nguyên nhân là “quan sát không được trung hòa, nặng quá hoặc nhẹ quá” (P1 tr.169-170)' if X.truthy(X.eq(E.v(18,19,21),'HÈN')) else ('có địa vị, chức phận, danh tiếng, được người trên trọng dụng; “quan gặp tuế vận là có quan chức” – thăng tiến rõ trong vận tốt (P1 tr.167-169)' if X.truthy(X.eq(E.v(18,19,21),'SANG')) else 'cân bằng: hai chiều ngang nhau, tốt xấu tùy vận'))),'. '),(X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu (',('không có' if X.truthy(X.eq(E.v(3,87,30),'')) else E.v(3,87,30))),'); phần xấu dễ nổi hơn trong bại vận ('),('không có' if X.truthy(X.eq(E.v(3,87,31),'')) else E.v(3,87,31))),'). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.') if X.truthy(X.AND(X.eq(E.v(3,87,28),''), X.eq(E.v(3,87,29),''))) else X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT (',('không có' if X.truthy(X.eq(E.v(3,87,28),'')) else E.v(3,87,28))),'); phần xấu nổi lên trong đại vận XẤU ('),('không có' if X.truthy(X.eq(E.v(3,87,29),'')) else E.v(3,87,29))),').')))

def f195(E, r, c):
    # Ô H99
    # = LUAN_GIAI!$J$117
    return E.v(4,117,10)

def f196(E, r, c):
    # Ô C100
    # = "[Thấp] "&THAM_KHAO!$E$13
    return X.cat('[Thấp] ',E.v(28,13,5))

def f197(E, r, c):
    # Ô H100
    # = LUAN_GIAI!$J$205
    return E.v(4,205,10)

def f198(E, r, c):
    # Ô C101
    # = TRIM(IF(TRIM(IF(SU_NGHIEP!$L$284="ĐỦ","• "&SU_NGHIEP!$G$284&" (P2 tr.548; UH tr.102-103, 194) ","")&IF(SU_NGHIEP!$L$285="ĐỦ","• "&SU_NGHIEP!$G$285&" (
    #   P2 tr.551, 536; UH tr.96, 239) ","")&IF(SU_NGHIEP!$L$286="ĐỦ","• "&SU_NGHIEP!$G$286&" (P2 tr.541; UH tr.239) ","")&IF(SU_NGHIEP!$L$287="ĐỦ","• "&SU_NG
    #   HIEP!$G$287&" (P2 tr.552; UH tr.71-72, 76, 88) ","")&IF(SU_NGHIEP!$L$288="ĐỦ","• "&SU_NGHIEP!$G$288&" (P2 tr.533; UH tr.71-72) ","")&IF(SU_NGHIEP!$L$2
    #   89="ĐỦ","• "&SU_NGHIEP!$G$289&" (P2 tr.542; UH tr.62, 72-73, 86) ",""))="","","Câu sách P2 (định nghĩa theo UH): "&TRIM(IF(SU_NGHIEP!$L$284="ĐỦ","• "&
    #   SU_NGHIEP!$G$284&" (P2 tr.548; UH tr.102-103, 194) ","")&IF(SU_NGHIEP!$L$285="ĐỦ","• "&SU_NGHIEP!$G$285&" (P2 tr.551, 536; UH tr.96, 239) ","")&IF(SU_
    #   NGHIEP!$L$286="ĐỦ","• "&SU_NGHIEP!$G$286&" (P2 tr.541; UH tr.239) ","")&IF(SU_NGHIEP!$L$287="ĐỦ","• "&SU_NGHIEP!$G$287&" (P2 tr.552; UH tr.71-72, 76, 
    #   88) ","")&IF(SU_NGHIEP!$L$288="ĐỦ","• "&SU_NGHIEP!$G$288&" (P2 tr.533; UH tr.71-72) ","")&IF(SU_NGHIEP!$L$289="ĐỦ","• "&SU_NGHIEP!$G$289&" (P2 tr.542;
    #    UH tr.62, 72-73, 86) ",""))&" ")&IF(""="","",""))
    return X.TRIM(X.cat(('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(21,284,7)),' (P2 tr.548; UH tr.102-103, 194) ') if X.truthy(X.eq(E.v(21,284,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(21,285,7)),' (P2 tr.551, 536; UH tr.96, 239) ') if X.truthy(X.eq(E.v(21,285,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,286,7)),' (P2 tr.541; UH tr.239) ') if X.truthy(X.eq(E.v(21,286,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,287,7)),' (P2 tr.552; UH tr.71-72, 76, 88) ') if X.truthy(X.eq(E.v(21,287,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,288,7)),' (P2 tr.533; UH tr.71-72) ') if X.truthy(X.eq(E.v(21,288,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,289,7)),' (P2 tr.542; UH tr.62, 72-73, 86) ') if X.truthy(X.eq(E.v(21,289,12),'ĐỦ')) else ''))),'')) else X.cat(X.cat('Câu sách P2 (định nghĩa theo UH): ',X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(21,284,7)),' (P2 tr.548; UH tr.102-103, 194) ') if X.truthy(X.eq(E.v(21,284,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(21,285,7)),' (P2 tr.551, 536; UH tr.96, 239) ') if X.truthy(X.eq(E.v(21,285,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,286,7)),' (P2 tr.541; UH tr.239) ') if X.truthy(X.eq(E.v(21,286,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,287,7)),' (P2 tr.552; UH tr.71-72, 76, 88) ') if X.truthy(X.eq(E.v(21,287,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,288,7)),' (P2 tr.533; UH tr.71-72) ') if X.truthy(X.eq(E.v(21,288,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,289,7)),' (P2 tr.542; UH tr.62, 72-73, 86) ') if X.truthy(X.eq(E.v(21,289,12),'ĐỦ')) else '')))),' ')),('' if X.truthy(X.eq('','')) else '')))

def f199(E, r, c):
    # Ô C102
    # = TRIM(TRIM(IF(SU_NGHIEP!$L$294="ĐỦ","• "&SU_NGHIEP!$G$294&" (HG tr.155) ","")&IF(SU_NGHIEP!$L$295="ĐỦ","• "&SU_NGHIEP!$G$295&" (HG tr.155) ","")&IF(SU_
    #   NGHIEP!$L$296="ĐỦ","• "&SU_NGHIEP!$G$296&" (HG tr.155) ","")&IF(SU_NGHIEP!$L$297="ĐỦ","• "&SU_NGHIEP!$G$297&" (HG tr.155) ","")&IF(SU_NGHIEP!$L$298="Đ
    #   Ủ","• "&SU_NGHIEP!$G$298&" (HG tr.155) ","")&IF(SU_NGHIEP!$L$299="ĐỦ","• "&SU_NGHIEP!$G$299&" (HG tr.156) ","")&IF(SU_NGHIEP!$L$300="ĐỦ","• "&SU_NGHIE
    #   P!$G$300&" (HG tr.156) ","")&IF(SU_NGHIEP!$L$301="ĐỦ","• "&SU_NGHIEP!$G$301&" (HG tr.156) ",""))&" "&SU_NGHIEP!$D$313)
    return X.TRIM(X.cat(X.cat(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(21,294,7)),' (HG tr.155) ') if X.truthy(X.eq(E.v(21,294,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(21,295,7)),' (HG tr.155) ') if X.truthy(X.eq(E.v(21,295,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,296,7)),' (HG tr.155) ') if X.truthy(X.eq(E.v(21,296,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,297,7)),' (HG tr.155) ') if X.truthy(X.eq(E.v(21,297,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,298,7)),' (HG tr.155) ') if X.truthy(X.eq(E.v(21,298,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,299,7)),' (HG tr.156) ') if X.truthy(X.eq(E.v(21,299,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,300,7)),' (HG tr.156) ') if X.truthy(X.eq(E.v(21,300,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,301,7)),' (HG tr.156) ') if X.truthy(X.eq(E.v(21,301,12),'ĐỦ')) else ''))),' '),E.v(21,313,4)))

def f200(E, r, c):
    # Ô C103
    # = TRIM(IF(SU_NGHIEP!$L$302="ĐỦ","• "&SU_NGHIEP!$G$302&" (HG tr.165) ","")&IF(SU_NGHIEP!$L$303="ĐỦ","• "&SU_NGHIEP!$G$303&" (HG tr.165) ","")&IF(SU_NGHIE
    #   P!$L$304="ĐỦ","• "&SU_NGHIEP!$G$304&" (HG tr.165) ","")&IF(SU_NGHIEP!$L$305="ĐỦ","• "&SU_NGHIEP!$G$305&" (HG tr.165) ","")&IF(SU_NGHIEP!$L$306="ĐỦ","•
    #    "&SU_NGHIEP!$G$306&" (HG tr.165) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(21,302,7)),' (HG tr.165) ') if X.truthy(X.eq(E.v(21,302,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(21,303,7)),' (HG tr.165) ') if X.truthy(X.eq(E.v(21,303,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,304,7)),' (HG tr.165) ') if X.truthy(X.eq(E.v(21,304,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,305,7)),' (HG tr.165) ') if X.truthy(X.eq(E.v(21,305,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,306,7)),' (HG tr.165) ') if X.truthy(X.eq(E.v(21,306,12),'ĐỦ')) else '')))

def f201(E, r, c):
    # Ô C104
    # = TRIM(IF(SU_NGHIEP!$L$307="ĐỦ","• "&SU_NGHIEP!$G$307&" (HG tr.148) ","")&IF(SU_NGHIEP!$L$308="ĐỦ","• "&SU_NGHIEP!$G$308&" (HG tr.148) ","")&IF(SU_NGHIE
    #   P!$L$309="ĐỦ","• "&SU_NGHIEP!$G$309&" (HG tr.148) ","")&IF(SU_NGHIEP!$L$310="ĐỦ","• "&SU_NGHIEP!$G$310&" (HG tr.148) ","")&IF(SU_NGHIEP!$L$311="ĐỦ","•
    #    "&SU_NGHIEP!$G$311&" (HG tr.148) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(21,307,7)),' (HG tr.148) ') if X.truthy(X.eq(E.v(21,307,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(21,308,7)),' (HG tr.148) ') if X.truthy(X.eq(E.v(21,308,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,309,7)),' (HG tr.148) ') if X.truthy(X.eq(E.v(21,309,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,310,7)),' (HG tr.148) ') if X.truthy(X.eq(E.v(21,310,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(21,311,7)),' (HG tr.148) ') if X.truthy(X.eq(E.v(21,311,12),'ĐỦ')) else '')))

def f202(E, r, c):
    # Ô C105
    # = TRIM(IF(SU_NGHIEP!$L$302="ĐỦ","• "&SU_NGHIEP!$G$302&" (HG tr.165) ",""))
    return X.TRIM((X.cat(X.cat('• ',E.v(21,302,7)),' (HG tr.165) ') if X.truthy(X.eq(E.v(21,302,12),'ĐỦ')) else ''))

def f203(E, r, c):
    # Ô C108
    # = C121
    return E.v(3,r+13,c)

def f204(E, r, c):
    # Ô B110, B111, B112, B113, B114
    # = LUC_THAN!B30
    return E.v(16,r-80,c)

def f205(E, r, c):
    # Ô C110
    # = LUC_THAN!D30&CHAR(10)&"Vì sao (dữ kiện → điểm, P1 tr.159): "&LUC_THAN!G30&IF(OR(LUC_THAN!H30="",LEFT(LUC_THAN!H30,5)="Thời "),"",CHAR(10)&LUC_THAN!H30
    #   )&CHAR(10)&SUBSTITUTE(LUC_THAN!K30,"Chỉ cần đề phòng khi gặp","Ứng rõ nhất khi gặp")&IF(LUC_THAN!$AM$30="","",CHAR(10)&"Câu sách bổ sung (bản 5.0, LUC
    #   _THAN bảng E2 – không đổi mức chấm ở trên): "&LUC_THAN!$AM$30)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(16,r-80,c+1),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(16,r-80,c+4)),('' if X.truthy(X.OR(X.eq(E.v(16,r-80,c+5),''), X.eq(X.LEFT(E.v(16,r-80,c+5), 5.0),'Thời '))) else X.cat(X.CHAR(10.0),E.v(16,r-80,c+5)))),X.CHAR(10.0)),X.SUBSTITUTE(E.v(16,r-80,c+8), 'Chỉ cần đề phòng khi gặp', 'Ứng rõ nhất khi gặp')),('' if X.truthy(X.eq(E.v(16,30,39),'')) else X.cat(X.cat(X.CHAR(10.0),'Câu sách bổ sung (bản 5.0, LUC_THAN bảng E2 – không đổi mức chấm ở trên): '),E.v(16,30,39))))

def f206(E, r, c):
    # Ô H110
    # = LUAN_GIAI!$J$110
    return E.v(4,110,10)

def f207(E, r, c):
    # Ô C111
    # = LUC_THAN!D31&CHAR(10)&"Vì sao (dữ kiện → điểm, P1 tr.159): "&LUC_THAN!G31&IF(OR(LUC_THAN!H31="",LEFT(LUC_THAN!H31,5)="Thời "),"",CHAR(10)&LUC_THAN!H31
    #   )&CHAR(10)&SUBSTITUTE(LUC_THAN!K31,"Chỉ cần đề phòng khi gặp","Ứng rõ nhất khi gặp")&IF(LUC_THAN!$AM$31="","",CHAR(10)&"Câu sách bổ sung (bản 5.0, LUC
    #   _THAN bảng E2 – không đổi mức chấm ở trên): "&LUC_THAN!$AM$31)&IF(LUC_THAN!$AQ$31="","",CHAR(10)&"Thêm (bảng E3): "&LUC_THAN!$AQ$31)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(16,r-80,c+1),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(16,r-80,c+4)),('' if X.truthy(X.OR(X.eq(E.v(16,r-80,c+5),''), X.eq(X.LEFT(E.v(16,r-80,c+5), 5.0),'Thời '))) else X.cat(X.CHAR(10.0),E.v(16,r-80,c+5)))),X.CHAR(10.0)),X.SUBSTITUTE(E.v(16,r-80,c+8), 'Chỉ cần đề phòng khi gặp', 'Ứng rõ nhất khi gặp')),('' if X.truthy(X.eq(E.v(16,31,39),'')) else X.cat(X.cat(X.CHAR(10.0),'Câu sách bổ sung (bản 5.0, LUC_THAN bảng E2 – không đổi mức chấm ở trên): '),E.v(16,31,39)))),('' if X.truthy(X.eq(E.v(16,31,43),'')) else X.cat(X.cat(X.CHAR(10.0),'Thêm (bảng E3): '),E.v(16,31,43))))

def f208(E, r, c):
    # Ô H111
    # = LUAN_GIAI!$J$111
    return E.v(4,111,10)

def f209(E, r, c):
    # Ô C112
    # = LUC_THAN!D32&CHAR(10)&"Vì sao (dữ kiện → điểm, P1 tr.159): "&LUC_THAN!G32&IF(OR(LUC_THAN!H32="",LEFT(LUC_THAN!H32,5)="Thời "),"",CHAR(10)&LUC_THAN!H32
    #   )&CHAR(10)&SUBSTITUTE(LUC_THAN!K32,"Chỉ cần đề phòng khi gặp","Ứng rõ nhất khi gặp")&IF(LUC_THAN!$AM$32="","",CHAR(10)&"Câu sách bổ sung (bản 5.0, LUC
    #   _THAN bảng E2 – không đổi mức chấm ở trên): "&LUC_THAN!$AM$32)&IF(LUC_THAN!$AQ$32="","",CHAR(10)&"Thêm (bảng E3): "&LUC_THAN!$AQ$32)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(16,r-80,c+1),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(16,r-80,c+4)),('' if X.truthy(X.OR(X.eq(E.v(16,r-80,c+5),''), X.eq(X.LEFT(E.v(16,r-80,c+5), 5.0),'Thời '))) else X.cat(X.CHAR(10.0),E.v(16,r-80,c+5)))),X.CHAR(10.0)),X.SUBSTITUTE(E.v(16,r-80,c+8), 'Chỉ cần đề phòng khi gặp', 'Ứng rõ nhất khi gặp')),('' if X.truthy(X.eq(E.v(16,32,39),'')) else X.cat(X.cat(X.CHAR(10.0),'Câu sách bổ sung (bản 5.0, LUC_THAN bảng E2 – không đổi mức chấm ở trên): '),E.v(16,32,39)))),('' if X.truthy(X.eq(E.v(16,32,43),'')) else X.cat(X.cat(X.CHAR(10.0),'Thêm (bảng E3): '),E.v(16,32,43))))

def f210(E, r, c):
    # Ô H112
    # = LUAN_GIAI!$J$112
    return E.v(4,112,10)

def f211(E, r, c):
    # Ô C113
    # = LUC_THAN!D33&CHAR(10)&"Vì sao (dữ kiện → điểm, P1 tr.159): "&LUC_THAN!G33&IF(OR(LUC_THAN!H33="",LEFT(LUC_THAN!H33,5)="Thời "),"",CHAR(10)&LUC_THAN!H33
    #   )&CHAR(10)&SUBSTITUTE(LUC_THAN!K33,"Chỉ cần đề phòng khi gặp","Ứng rõ nhất khi gặp")&IF(LUC_THAN!$AM$33="","",CHAR(10)&"Câu sách bổ sung (bản 5.0, LUC
    #   _THAN bảng E3 – không đổi mức chấm ở trên): "&LUC_THAN!$AM$33)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(16,r-80,c+1),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(16,r-80,c+4)),('' if X.truthy(X.OR(X.eq(E.v(16,r-80,c+5),''), X.eq(X.LEFT(E.v(16,r-80,c+5), 5.0),'Thời '))) else X.cat(X.CHAR(10.0),E.v(16,r-80,c+5)))),X.CHAR(10.0)),X.SUBSTITUTE(E.v(16,r-80,c+8), 'Chỉ cần đề phòng khi gặp', 'Ứng rõ nhất khi gặp')),('' if X.truthy(X.eq(E.v(16,33,39),'')) else X.cat(X.cat(X.CHAR(10.0),'Câu sách bổ sung (bản 5.0, LUC_THAN bảng E3 – không đổi mức chấm ở trên): '),E.v(16,33,39))))

def f212(E, r, c):
    # Ô H113
    # = LUAN_GIAI!$J$113
    return E.v(4,113,10)

def f213(E, r, c):
    # Ô C114
    # = IF(LUU_NIEN!$B$5<=18,"(Lưu ý tuổi: chưa thành niên – mục con cái là xu hướng xa, chưa xét lúc này.)"&CHAR(10),"")&LUC_THAN!D34&CHAR(10)&"Vì sao (dữ ki
    #   ện → điểm, P1 tr.159): "&LUC_THAN!G34&IF(OR(LUC_THAN!H34="",LEFT(LUC_THAN!H34,5)="Thời "),"",CHAR(10)&LUC_THAN!H34)&CHAR(10)&SUBSTITUTE(LUC_THAN!K34,"
    #   Chỉ cần đề phòng khi gặp","Ứng rõ nhất khi gặp")&IF(LUC_THAN!$AM$34="","",CHAR(10)&"Câu sách bổ sung (bản 5.0, LUC_THAN bảng E3 – không đổi mức chấm ở
    #    trên): "&LUC_THAN!$AM$34)&IF(AND(OR(ISNUMBER(SEARCH("không có con",C126)),ISNUMBER(SEARCH("khó có con",C126)),ISNUMBER(SEARCH("ít con",C126)),ISNUMBE
    #   R(SEARCH("không có con",C128)),ISNUMBER(SEARCH("khó có con",C128)),ISNUMBER(SEARCH("ít con",C128))),LEFT(LUC_THAN!D34,3)<>"KÉM",LEFT(LUC_THAN!D34,3)<>
    #   "XẤU"),CHAR(10)&"Đối chiếu nguồn (quy tắc 6.6): Uyên Hải Tử Bình / Hóa giải toàn thư có câu xấu hơn về con cái (dòng 126, 128). Giữ kết luận Thiệu Vĩ 
    #   Hoa theo thứ tự tin cậy đã duyệt (Thiệu Vĩ Hoa > Uyên Hải Tử Bình > Cùng Thông Bảo Giám, Trích Thiên Tủy, Hóa giải toàn thư > tham khảo); nên xem là đ
    #   iểm cần lưu ý.","")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat('(Lưu ý tuổi: chưa thành niên – mục con cái là xu hướng xa, chưa xét lúc này.)',X.CHAR(10.0)) if X.truthy(X.le(E.v(15,5,2),18.0)) else ''),E.v(16,r-80,c+1)),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(16,r-80,c+4)),('' if X.truthy(X.OR(X.eq(E.v(16,r-80,c+5),''), X.eq(X.LEFT(E.v(16,r-80,c+5), 5.0),'Thời '))) else X.cat(X.CHAR(10.0),E.v(16,r-80,c+5)))),X.CHAR(10.0)),X.SUBSTITUTE(E.v(16,r-80,c+8), 'Chỉ cần đề phòng khi gặp', 'Ứng rõ nhất khi gặp')),('' if X.truthy(X.eq(E.v(16,34,39),'')) else X.cat(X.cat(X.CHAR(10.0),'Câu sách bổ sung (bản 5.0, LUC_THAN bảng E3 – không đổi mức chấm ở trên): '),E.v(16,34,39)))),(X.cat(X.CHAR(10.0),'Đối chiếu nguồn (quy tắc 6.6): Uyên Hải Tử Bình / Hóa giải toàn thư có câu xấu hơn về con cái (dòng 126, 128). Giữ kết luận Thiệu Vĩ Hoa theo thứ tự tin cậy đã duyệt (Thiệu Vĩ Hoa > Uyên Hải Tử Bình > Cùng Thông Bảo Giám, Trích Thiên Tủy, Hóa giải toàn thư > tham khảo); nên xem là điểm cần lưu ý.') if X.truthy(X.AND(X.OR(X.ISNUMBER(lambda: X.SEARCH('không có con', E.v(3,r+12,c))), X.ISNUMBER(lambda: X.SEARCH('khó có con', E.v(3,r+12,c))), X.ISNUMBER(lambda: X.SEARCH('ít con', E.v(3,r+12,c))), X.ISNUMBER(lambda: X.SEARCH('không có con', E.v(3,r+14,c))), X.ISNUMBER(lambda: X.SEARCH('khó có con', E.v(3,r+14,c))), X.ISNUMBER(lambda: X.SEARCH('ít con', E.v(3,r+14,c)))), X.ne(X.LEFT(E.v(16,r-80,c+1), 3.0),'KÉM'), X.ne(X.LEFT(E.v(16,r-80,c+1), 3.0),'XẤU'))) else ''))

def f214(E, r, c):
    # Ô H114
    # = LUAN_GIAI!$J$114
    return E.v(4,114,10)

def f215(E, r, c):
    # Ô C115
    # = LUC_THAN!D23&CHAR(10)&"Vì sao (dữ kiện → điểm, P1 tr.159): "&LUC_THAN!G23&" – "&LUC_THAN!E23
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(16,r-92,c+1),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(16,r-92,c+4)),' – '),E.v(16,r-92,c+2))

def f216(E, r, c):
    # Ô H115
    # = LUAN_GIAI!$J$175
    return E.v(4,175,10)

def f217(E, r, c):
    # Ô C117
    # = IF(LUU_NIEN!$B$5<=18,"(Lưu ý tuổi: chưa thành niên – mục hôn nhân là xu hướng xa, chưa xét lúc này.)"&CHAR(10),"")&HON_NHAN!D64&CHAR(10)&"Vì sao (dữ k
    #   iện → điểm, P1 tr.159): "&HON_NHAN!G64&IF(HON_NHAN!AG64>0,CHAR(10)&"⚠ Mức trên là đánh giá SAO PHỐI NGẪU (người "&IF(NHAP_LIEU!$C$22="Nam","vợ","chồng
    #   ")&" có lực, có năng lực hay không). Về sự êm ấm – bền vững của hôn nhân, sách có "&HON_NHAN!AG64&" dấu hiệu BẤT LỢI đủ điều kiện ngay trong lá số: "&
    #   HON_NHAN!AF64&"→ hôn nhân dễ trắc trở (va chạm, xa cách; nặng thì đổ vỡ); nặng nhẹ và thời điểm xem đại vận, năm bất lợi ở dòng 70 và HON_NHAN.","")&C
    #   HAR(10)&SUBSTITUTE(HON_NHAN!K64,"Chỉ cần đề phòng khi gặp","Ứng rõ nhất khi gặp")&IF(HON_NHAN!$AM$64="","",CHAR(10)&"Câu sách bổ sung (bản 5.0, HON_NH
    #   AN bảng F2 – không đổi mức chấm ở trên): "&HON_NHAN!$AM$64)
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat('(Lưu ý tuổi: chưa thành niên – mục hôn nhân là xu hướng xa, chưa xét lúc này.)',X.CHAR(10.0)) if X.truthy(X.le(E.v(15,5,2),18.0)) else ''),E.v(17,r-53,c+1)),X.CHAR(10.0)),'Vì sao (dữ kiện → điểm, P1 tr.159): '),E.v(17,r-53,c+4)),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Mức trên là đánh giá SAO PHỐI NGẪU (người '),('vợ' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else 'chồng')),' có lực, có năng lực hay không). Về sự êm ấm – bền vững của hôn nhân, sách có '),E.v(17,r-53,c+30)),' dấu hiệu BẤT LỢI đủ điều kiện ngay trong lá số: '),E.v(17,r-53,c+29)),'→ hôn nhân dễ trắc trở (va chạm, xa cách; nặng thì đổ vỡ); nặng nhẹ và thời điểm xem đại vận, năm bất lợi ở dòng 70 và HON_NHAN.') if X.truthy(X.gt(E.v(17,r-53,c+30),0.0)) else '')),X.CHAR(10.0)),X.SUBSTITUTE(E.v(17,r-53,c+8), 'Chỉ cần đề phòng khi gặp', 'Ứng rõ nhất khi gặp')),('' if X.truthy(X.eq(E.v(17,64,39),'')) else X.cat(X.cat(X.CHAR(10.0),'Câu sách bổ sung (bản 5.0, HON_NHAN bảng F2 – không đổi mức chấm ở trên): '),E.v(17,64,39))))

def f218(E, r, c):
    # Ô H117
    # = LUAN_GIAI!$J$115
    return E.v(4,115,10)

def f219(E, r, c):
    # Ô C118
    # = HON_NHAN!G64
    return E.v(17,r-54,c+4)

def f220(E, r, c):
    # Ô H118
    # = LUAN_GIAI!$J$176
    return E.v(4,176,10)

def f221(E, r, c):
    # Ô C119
    # = HON_NHAN!D66&CHAR(10)&HON_NHAN!H64
    return X.cat(X.cat(E.v(17,r-53,c+1),X.CHAR(10.0)),E.v(17,r-55,c+5))

def f222(E, r, c):
    # Ô H119
    # = LUAN_GIAI!$J$177
    return E.v(4,177,10)

def f223(E, r, c):
    # Ô C120
    # = HON_NHAN!D67
    return E.v(17,r-53,c+1)

def f224(E, r, c):
    # Ô H120
    # = LUAN_GIAI!$J$178
    return E.v(4,178,10)

def f225(E, r, c):
    # Ô C121
    # = C206
    return E.v(3,r+85,c)

def f226(E, r, c):
    # Ô H121
    # = LUAN_GIAI!$J$179
    return E.v(4,179,10)

def f227(E, r, c):
    # Ô C122
    # = "[Thấp] "&THAM_KHAO!$E$7&CHAR(10)&"[Thấp] "&THAM_KHAO!$E$8
    return X.cat(X.cat(X.cat(X.cat('[Thấp] ',E.v(28,7,5)),X.CHAR(10.0)),'[Thấp] '),E.v(28,8,5))

def f228(E, r, c):
    # Ô H122
    # = LUAN_GIAI!$J$202
    return E.v(4,202,10)

def f229(E, r, c):
    # Ô C123
    # = "[Thấp] "&THAM_KHAO!$E$10&CHAR(10)&"[Thấp] "&THAM_KHAO!$E$12
    return X.cat(X.cat(X.cat(X.cat('[Thấp] ',E.v(28,10,5)),X.CHAR(10.0)),'[Thấp] '),E.v(28,12,5))

def f230(E, r, c):
    # Ô H123
    # = LUAN_GIAI!$J$203
    return E.v(4,203,10)

def f231(E, r, c):
    # Ô C124
    # = "[Trung bình] "&THAM_KHAO!$E$9
    return X.cat('[Trung bình] ',E.v(28,9,5))

def f232(E, r, c):
    # Ô H124
    # = LUAN_GIAI!$J$204
    return E.v(4,204,10)

def f233(E, r, c):
    # Ô C125
    # = TRIM(IF(LUC_THAN!$L$328="ĐỦ","• "&LUC_THAN!$G$328&" (P2 tr.508; UH tr.62-63, 77, 156, 182) ","")&IF(LUC_THAN!$L$329="ĐỦ","• "&LUC_THAN!$G$329&" (P2 tr
    #   .440; UH tr.24) ",""))
    return X.TRIM(X.cat((X.cat(X.cat('• ',E.v(16,328,7)),' (P2 tr.508; UH tr.62-63, 77, 156, 182) ') if X.truthy(X.eq(E.v(16,328,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(16,329,7)),' (P2 tr.440; UH tr.24) ') if X.truthy(X.eq(E.v(16,329,12),'ĐỦ')) else '')))

def f234(E, r, c):
    # Ô C126
    # = TRIM(LUC_THAN!$D$345&" "&TRIM(IF(LUC_THAN!$L$330="ĐỦ","• "&LUC_THAN!$G$330&" (P2 tr.509; UH tr.170, 177) ","")&IF(LUC_THAN!$L$331="ĐỦ","• "&LUC_THAN!$
    #   G$331&" (UH tr.48-49, 52) ","")&IF(LUC_THAN!$L$332="ĐỦ","• "&LUC_THAN!$G$332&" (UH tr.48-49, 127) ","")&IF(LUC_THAN!$L$333="ĐỦ","• "&LUC_THAN!$G$333&"
    #    (UH tr.126-127) ","")&IF(LUC_THAN!$L$334="ĐỦ","• "&LUC_THAN!$G$334&" (UH tr.126-127) ","")&IF(LUC_THAN!$L$335="ĐỦ","• "&LUC_THAN!$G$335&" (UH tr.162)
    #    ","")&IF(LUC_THAN!$L$336="ĐỦ","• "&LUC_THAN!$G$336&" (UH tr.52, 178) ","")&IF(LUC_THAN!$L$337="ĐỦ","• "&LUC_THAN!$G$337&" (UH tr.159) ","")&IF(LUC_TH
    #   AN!$L$338="ĐỦ","• "&LUC_THAN!$G$338&" (UH tr.159) ","")&IF(LUC_THAN!$L$339="ĐỦ","• "&LUC_THAN!$G$339&" (UH tr.178) ","")&IF(LUC_THAN!$L$340="ĐỦ","• "&
    #   LUC_THAN!$G$340&" (UH tr.232, 238) ","")&IF(LUC_THAN!$L$341="ĐỦ","• "&LUC_THAN!$G$341&" (UH tr.239) ","")&IF(LUC_THAN!$L$342="ĐỦ","• "&LUC_THAN!$G$342
    #   &" (UH tr.239) ","")&IF(LUC_THAN!$L$343="ĐỦ","• "&LUC_THAN!$G$343&" (UH tr.127-128) ","")))
    return X.TRIM(X.cat(X.cat(E.v(16,345,4),' '),X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(16,330,7)),' (P2 tr.509; UH tr.170, 177) ') if X.truthy(X.eq(E.v(16,330,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(16,331,7)),' (UH tr.48-49, 52) ') if X.truthy(X.eq(E.v(16,331,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,332,7)),' (UH tr.48-49, 127) ') if X.truthy(X.eq(E.v(16,332,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,333,7)),' (UH tr.126-127) ') if X.truthy(X.eq(E.v(16,333,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,334,7)),' (UH tr.126-127) ') if X.truthy(X.eq(E.v(16,334,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,335,7)),' (UH tr.162) ') if X.truthy(X.eq(E.v(16,335,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,336,7)),' (UH tr.52, 178) ') if X.truthy(X.eq(E.v(16,336,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,337,7)),' (UH tr.159) ') if X.truthy(X.eq(E.v(16,337,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,338,7)),' (UH tr.159) ') if X.truthy(X.eq(E.v(16,338,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,339,7)),' (UH tr.178) ') if X.truthy(X.eq(E.v(16,339,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,340,7)),' (UH tr.232, 238) ') if X.truthy(X.eq(E.v(16,340,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,341,7)),' (UH tr.239) ') if X.truthy(X.eq(E.v(16,341,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,342,7)),' (UH tr.239) ') if X.truthy(X.eq(E.v(16,342,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,343,7)),' (UH tr.127-128) ') if X.truthy(X.eq(E.v(16,343,12),'ĐỦ')) else '')))))

def f235(E, r, c):
    # Ô C127
    # = TRIM(IF(TRIM(IF(HON_NHAN!$L$280="ĐỦ","• "&HON_NHAN!$G$280&" (P2 tr.470; UH tr.280) ","")&IF(HON_NHAN!$L$281="ĐỦ","• "&HON_NHAN!$G$281&" (P2 tr.465; UH
    #    tr.128) ","")&IF(HON_NHAN!$L$282="ĐỦ","• "&HON_NHAN!$G$282&" (P2 tr.465; UH tr.128) ",""))="","","Câu sách P2 (định nghĩa theo UH): "&TRIM(IF(HON_NHA
    #   N!$L$280="ĐỦ","• "&HON_NHAN!$G$280&" (P2 tr.470; UH tr.280) ","")&IF(HON_NHAN!$L$281="ĐỦ","• "&HON_NHAN!$G$281&" (P2 tr.465; UH tr.128) ","")&IF(HON_N
    #   HAN!$L$282="ĐỦ","• "&HON_NHAN!$G$282&" (P2 tr.465; UH tr.128) ",""))&" ")&IF(""="","",""))
    return X.TRIM(X.cat(('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat((X.cat(X.cat('• ',E.v(17,280,7)),' (P2 tr.470; UH tr.280) ') if X.truthy(X.eq(E.v(17,280,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(17,281,7)),' (P2 tr.465; UH tr.128) ') if X.truthy(X.eq(E.v(17,281,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(17,282,7)),' (P2 tr.465; UH tr.128) ') if X.truthy(X.eq(E.v(17,282,12),'ĐỦ')) else ''))),'')) else X.cat(X.cat('Câu sách P2 (định nghĩa theo UH): ',X.TRIM(X.cat(X.cat((X.cat(X.cat('• ',E.v(17,280,7)),' (P2 tr.470; UH tr.280) ') if X.truthy(X.eq(E.v(17,280,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(17,281,7)),' (P2 tr.465; UH tr.128) ') if X.truthy(X.eq(E.v(17,281,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(17,282,7)),' (P2 tr.465; UH tr.128) ') if X.truthy(X.eq(E.v(17,282,12),'ĐỦ')) else '')))),' ')),('' if X.truthy(X.eq('','')) else '')))

def f236(E, r, c):
    # Ô C128
    # = TRIM(IF(LUC_THAN!$L$350="ĐỦ","• "&LUC_THAN!$G$350&" (HG tr.193) ","")&IF(LUC_THAN!$L$351="ĐỦ","• "&LUC_THAN!$G$351&" (HG tr.193) ","")&IF(LUC_THAN!$L$
    #   352="ĐỦ","• "&LUC_THAN!$G$352&" (HG tr.193) ","")&IF(LUC_THAN!$L$353="ĐỦ","• "&LUC_THAN!$G$353&" (HG tr.193) ","")&IF(LUC_THAN!$L$354="ĐỦ","• "&LUC_TH
    #   AN!$G$354&" (HG tr.193) ","")&IF(LUC_THAN!$L$355="ĐỦ","• "&LUC_THAN!$G$355&" (HG tr.194) ","")&IF(LUC_THAN!$L$356="ĐỦ","• "&LUC_THAN!$G$356&" (HG tr.1
    #   94) ","")&IF(LUC_THAN!$L$357="ĐỦ","• "&LUC_THAN!$G$357&" (HG tr.194) ","")&IF(LUC_THAN!$L$358="ĐỦ","• "&LUC_THAN!$G$358&" (HG tr.194) ","")&IF(LUC_THA
    #   N!$L$359="ĐỦ","• "&LUC_THAN!$G$359&" (HG tr.195) ","")&IF(LUC_THAN!$L$360="ĐỦ","• "&LUC_THAN!$G$360&" (HG tr.195) ","")&IF(LUC_THAN!$L$361="ĐỦ","• "&L
    #   UC_THAN!$G$361&" (HG tr.195) ","")&IF(LUC_THAN!$L$362="ĐỦ","• "&LUC_THAN!$G$362&" (HG tr.195) ","")&IF(LUC_THAN!$L$363="ĐỦ","• "&LUC_THAN!$G$363&" (HG
    #    tr.196) ","")&IF(LUC_THAN!$L$364="ĐỦ","• "&LUC_THAN!$G$364&" (HG tr.196) ","")&IF(LUC_THAN!$L$365="ĐỦ","• "&LUC_THAN!$G$365&" (HG tr.196) ","")&IF(LU
    #   C_THAN!$L$366="ĐỦ","• "&LUC_THAN!$G$366&" (HG tr.197) ","")&IF(LUC_THAN!$L$367="ĐỦ","• "&LUC_THAN!$G$367&" (HG tr.198) ","")&IF(LUC_THAN!$L$368="ĐỦ","
    #   • "&LUC_THAN!$G$368&" (HG tr.199) ","")&IF(LUC_THAN!$L$369="ĐỦ","• "&LUC_THAN!$G$369&" (HG tr.200) ","")&IF(LUC_THAN!$L$370="ĐỦ","• "&LUC_THAN!$G$370&
    #   " (HG tr.200) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(16,350,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,350,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(16,351,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,351,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,352,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,352,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,353,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,353,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,354,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,354,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,355,7)),' (HG tr.194) ') if X.truthy(X.eq(E.v(16,355,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,356,7)),' (HG tr.194) ') if X.truthy(X.eq(E.v(16,356,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,357,7)),' (HG tr.194) ') if X.truthy(X.eq(E.v(16,357,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,358,7)),' (HG tr.194) ') if X.truthy(X.eq(E.v(16,358,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,359,7)),' (HG tr.195) ') if X.truthy(X.eq(E.v(16,359,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,360,7)),' (HG tr.195) ') if X.truthy(X.eq(E.v(16,360,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,361,7)),' (HG tr.195) ') if X.truthy(X.eq(E.v(16,361,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,362,7)),' (HG tr.195) ') if X.truthy(X.eq(E.v(16,362,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,363,7)),' (HG tr.196) ') if X.truthy(X.eq(E.v(16,363,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,364,7)),' (HG tr.196) ') if X.truthy(X.eq(E.v(16,364,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,365,7)),' (HG tr.196) ') if X.truthy(X.eq(E.v(16,365,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,366,7)),' (HG tr.197) ') if X.truthy(X.eq(E.v(16,366,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,367,7)),' (HG tr.198) ') if X.truthy(X.eq(E.v(16,367,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,368,7)),' (HG tr.199) ') if X.truthy(X.eq(E.v(16,368,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,369,7)),' (HG tr.200) ') if X.truthy(X.eq(E.v(16,369,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,370,7)),' (HG tr.200) ') if X.truthy(X.eq(E.v(16,370,12),'ĐỦ')) else '')))

def f237(E, r, c):
    # Ô C129
    # = LUC_THAN!$D$380
    return E.v(16,380,4)

def f238(E, r, c):
    # Ô C130
    # = TRIM(IF(LUC_THAN!$L$371="ĐỦ","• "&LUC_THAN!$G$371&" (HG tr.205) ","")&IF(LUC_THAN!$L$372="ĐỦ","• "&LUC_THAN!$G$372&" (HG tr.206) ","")&IF(LUC_THAN!$L$
    #   373="ĐỦ","• "&LUC_THAN!$G$373&" (HG tr.206) ","")&IF(LUC_THAN!$L$374="ĐỦ","• "&LUC_THAN!$G$374&" (HG tr.202) ","")&IF(LUC_THAN!$L$375="ĐỦ","• "&LUC_TH
    #   AN!$G$375&" (HG tr.203) ","")&IF(LUC_THAN!$L$376="ĐỦ","• "&LUC_THAN!$G$376&" (HG tr.202) ","")&IF(LUC_THAN!$L$377="ĐỦ","• "&LUC_THAN!$G$377&" (HG tr.2
    #   02) ","")&IF(LUC_THAN!$L$378="ĐỦ","• "&LUC_THAN!$G$378&" (HG tr.203) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(16,371,7)),' (HG tr.205) ') if X.truthy(X.eq(E.v(16,371,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(16,372,7)),' (HG tr.206) ') if X.truthy(X.eq(E.v(16,372,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,373,7)),' (HG tr.206) ') if X.truthy(X.eq(E.v(16,373,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,374,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(16,374,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,375,7)),' (HG tr.203) ') if X.truthy(X.eq(E.v(16,375,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,376,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(16,376,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,377,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(16,377,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,378,7)),' (HG tr.203) ') if X.truthy(X.eq(E.v(16,378,12),'ĐỦ')) else '')))

def f239(E, r, c):
    # Ô C131
    # = TRIM(IF(LUC_THAN!$L$350="ĐỦ","• "&LUC_THAN!$G$350&" (HG tr.193) ","")&IF(LUC_THAN!$L$351="ĐỦ","• "&LUC_THAN!$G$351&" (HG tr.193) ","")&IF(LUC_THAN!$L$
    #   352="ĐỦ","• "&LUC_THAN!$G$352&" (HG tr.193) ","")&IF(LUC_THAN!$L$353="ĐỦ","• "&LUC_THAN!$G$353&" (HG tr.193) ","")&IF(LUC_THAN!$L$354="ĐỦ","• "&LUC_TH
    #   AN!$G$354&" (HG tr.193) ","")&IF(LUC_THAN!$L$356="ĐỦ","• "&LUC_THAN!$G$356&" (HG tr.194) ","")&IF(LUC_THAN!$L$363="ĐỦ","• "&LUC_THAN!$G$363&" (HG tr.1
    #   96) ","")&IF(LUC_THAN!$L$364="ĐỦ","• "&LUC_THAN!$G$364&" (HG tr.196) ","")&IF(LUC_THAN!$L$365="ĐỦ","• "&LUC_THAN!$G$365&" (HG tr.196) ","")&IF(LUC_THA
    #   N!$L$368="ĐỦ","• "&LUC_THAN!$G$368&" (HG tr.199) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(16,350,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,350,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(16,351,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,351,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,352,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,352,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,353,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,353,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,354,7)),' (HG tr.193) ') if X.truthy(X.eq(E.v(16,354,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,356,7)),' (HG tr.194) ') if X.truthy(X.eq(E.v(16,356,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,363,7)),' (HG tr.196) ') if X.truthy(X.eq(E.v(16,363,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,364,7)),' (HG tr.196) ') if X.truthy(X.eq(E.v(16,364,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,365,7)),' (HG tr.196) ') if X.truthy(X.eq(E.v(16,365,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,368,7)),' (HG tr.199) ') if X.truthy(X.eq(E.v(16,368,12),'ĐỦ')) else '')))

def f240(E, r, c):
    # Ô C132
    # = TRIM(IF(LUC_THAN!$L$371="ĐỦ","• "&LUC_THAN!$G$371&" (HG tr.205) ","")&IF(LUC_THAN!$L$374="ĐỦ","• "&LUC_THAN!$G$374&" (HG tr.202) ","")&IF(LUC_THAN!$L$
    #   376="ĐỦ","• "&LUC_THAN!$G$376&" (HG tr.202) ","")&IF(LUC_THAN!$L$377="ĐỦ","• "&LUC_THAN!$G$377&" (HG tr.202) ",""))
    return X.TRIM(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(16,371,7)),' (HG tr.205) ') if X.truthy(X.eq(E.v(16,371,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(16,374,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(16,374,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,376,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(16,376,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(16,377,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(16,377,12),'ĐỦ')) else '')))

def f241(E, r, c):
    # Ô C135
    # = "Thọ – yểu (P1 tr.172–173): "&LEFT(C143,IFERROR(FIND(CHAR(10),C143)-1,LEN(C143)))&". "&C142
    return X.cat(X.cat(X.cat('Thọ – yểu (P1 tr.172–173): ',X.LEFT(E.v(3,r+8,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+8,c)),1.0), lambda: X.LEN(E.v(3,r+8,c))))),'. '),E.v(3,r+7,c))

def f242(E, r, c):
    # Ô C137
    # = SUC_KHOE!$J$6&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(LEFT(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6&CHAR(10))-1),7)="BẤT LỢI","sức khỏe yếu, dễ ốm"&IF(
    #   ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6&CHAR(10))-1))),IF(AND(SUC_KHOE!$M$6>=2*SUC_KHOE!$N$6,SUC_KHOE!$M$6>=4,FALSE
    #   ())," – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo 
    #   đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt"),"")
    #   ,IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6&CHAR(10))-1))),"hơi bất lợi – sức khỏe yếu, dễ ốm, ở mức nhẹ",IF(ISNU
    #   MBER(SEARCH("TRUNG BÌNH",LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6&CHAR(10))-1))),"bình thường – lúc thuận lúc không, tùy vận",IF(ISNUMBER(SEARCH
    #   ("NGHIÊNG THUẬN",LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6&CHAR(10))-1))),"khá – thể chất tốt, ít ốm đau, ở mức vừa, không phải thế mạnh nổi bật"
    #   ,"thể chất tốt, ít ốm đau"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6&CHAR(10))-1)))," – nhưng có trở ngại (các dấu 
    #   hiệu bất lợi ở trên)","")))))&"."&CHAR(10)&"Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: "&SUC_KHOE!$C$6&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điều kiện: "&
    #   SUC_KHOE!$F$6&CHAR(10)&"(Cách kết luận: SUC_KHOE dòng 9, giả định G17.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,6,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('sức khỏe yếu, dễ ốm',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(22,6,13),X.mul(2.0,E.v(22,6,14))), X.ge(E.v(22,6,13),4.0), X.FALSE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,6,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,6,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – sức khỏe yếu, dễ ốm, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,6,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,6,10),X.CHAR(10.0))),1.0))))) else ('khá – thể chất tốt, ít ốm đau, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,6,10),X.CHAR(10.0))),1.0))))) else X.cat('thể chất tốt, ít ốm đau',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(22,6,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,6,10),X.CHAR(10.0))),1.0))))) else ''))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(22,6,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(22,6,6)),X.CHAR(10.0)),'(Cách kết luận: SUC_KHOE dòng 9, giả định G17.)')

def f243(E, r, c):
    # Ô H137
    # = LUAN_GIAI!$J$127
    return E.v(4,127,10)

def f244(E, r, c):
    # Ô C138
    # = SUC_KHOE!$J$7&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(LEFT(LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7&CHAR(10))-1),7)="BẤT LỢI","dễ mắc bệnh ở các tạng p
    #   hủ nêu dưới"&IF(ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7&CHAR(10))-1))),IF(AND(SUC_KHOE!$M$7>=2*SUC_KHOE!$N$7,SUC_KH
    #   OE!$M$7>=4,FALSE())," – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535
    #   ), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện tro
    #   ng vận tốt"),""),IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7&CHAR(10))-1))),"hơi bất lợi – dễ mắc bệnh ở các tạng 
    #   phủ nêu dưới, ở mức nhẹ",IF(ISNUMBER(SEARCH("TRUNG BÌNH",LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7&CHAR(10))-1))),"bình thường – lúc thuận lúc kh
    #   ông, tùy vận",IF(ISNUMBER(SEARCH("NGHIÊNG THUẬN",LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7&CHAR(10))-1))),"khá – ít bệnh nặng, ở mức vừa, không p
    #   hải thế mạnh nổi bật","ít bệnh nặng"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SUC_KHOE!$J$7,FIND(CHAR(10),SUC_KHOE!$J$7&CHAR(10))-1)))," – nhưng có trở ngại
    #    (các dấu hiệu bất lợi ở trên)","")))))&"."&CHAR(10)&"Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: "&SUC_KHOE!$C$7&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điề
    #   u kiện: "&SUC_KHOE!$F$7&CHAR(10)&"(Cách kết luận: SUC_KHOE dòng 9, giả định G17.)"&CHAR(10)&SUC_KHOE!$C$18
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,7,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('dễ mắc bệnh ở các tạng phủ nêu dưới',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(22,7,13),X.mul(2.0,E.v(22,7,14))), X.ge(E.v(22,7,13),4.0), X.FALSE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,7,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,7,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – dễ mắc bệnh ở các tạng phủ nêu dưới, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,7,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,7,10),X.CHAR(10.0))),1.0))))) else ('khá – ít bệnh nặng, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,7,10),X.CHAR(10.0))),1.0))))) else X.cat('ít bệnh nặng',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(22,7,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,7,10),X.CHAR(10.0))),1.0))))) else ''))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(22,7,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(22,7,6)),X.CHAR(10.0)),'(Cách kết luận: SUC_KHOE dòng 9, giả định G17.)'),X.CHAR(10.0)),E.v(22,18,3))

def f245(E, r, c):
    # Ô H138
    # = LUAN_GIAI!$J$128
    return E.v(4,128,10)

def f246(E, r, c):
    # Ô C139
    # = SUC_KHOE!$J$8&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(LEFT(LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8&CHAR(10))-1),7)="BẤT LỢI","dễ gặp tai nạn, thương t
    #   ật, nhất là trong các vận, năm nêu dưới"&IF(ISNUMBER(SEARCH("ĐIỂM SÁNG",LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8&CHAR(10))-1))),IF(AND(SUC_KHOE!
    #   $M$8>=2*SUC_KHOE!$N$8,SUC_KHOE!$M$8>=4,FALSE())," – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có ch
    #   ế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới"," – nhưng vẫn có điểm sáng (các dấu hiệu
    #    thuận ở trên), thể hiện trong vận tốt"),""),IF(ISNUMBER(SEARCH("NGHIÊNG BẤT",LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8&CHAR(10))-1))),"hơi bất l
    #   ợi – dễ gặp tai nạn, thương tật, nhất là trong các vận, năm nêu dưới, ở mức nhẹ",IF(ISNUMBER(SEARCH("TRUNG BÌNH",LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_
    #   KHOE!$J$8&CHAR(10))-1))),"bình thường – lúc thuận lúc không, tùy vận",IF(ISNUMBER(SEARCH("NGHIÊNG THUẬN",LEFT(SUC_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$
    #   8&CHAR(10))-1))),"khá – ít tai nạn, thương tật, ở mức vừa, không phải thế mạnh nổi bật","ít tai nạn, thương tật"&IF(ISNUMBER(SEARCH("TRỞ NGẠI",LEFT(SU
    #   C_KHOE!$J$8,FIND(CHAR(10),SUC_KHOE!$J$8&CHAR(10))-1)))," – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)","")))))&"."&CHAR(10)&"Vì sao – dấu hiệu TH
    #   UẬN của sách đủ điều kiện: "&SUC_KHOE!$C$8&CHAR(10)&"Dấu hiệu BẤT LỢI đủ điều kiện: "&SUC_KHOE!$F$8&CHAR(10)&"(Cách kết luận: SUC_KHOE dòng 9, giả địn
    #   h G17.)"&CHAR(10)&SUC_KHOE!$C$28&CHAR(10)&SUC_KHOE!$C$157
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,8,10),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),(X.cat('dễ gặp tai nạn, thương tật, nhất là trong các vận, năm nêu dưới',((' – nhưng điểm sáng lớn: dấu hiệu thuận gấp đôi trở lên (sao kỵ bị chế, tiết, hợp – vd “thất sát có chế ngự là quân tử”, P2 tr.535), nên vẫn theo đuổi được, chỉ là tiến chậm và phải chọn đúng vận thuận nêu dưới' if X.truthy(X.AND(X.ge(E.v(22,8,13),X.mul(2.0,E.v(22,8,14))), X.ge(E.v(22,8,13),4.0), X.FALSE())) else ' – nhưng vẫn có điểm sáng (các dấu hiệu thuận ở trên), thể hiện trong vận tốt') if X.truthy(X.ISNUMBER(lambda: X.SEARCH('ĐIỂM SÁNG', X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,8,10),X.CHAR(10.0))),1.0))))) else '')) if X.truthy(X.eq(X.LEFT(X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,8,10),X.CHAR(10.0))),1.0)), 7.0),'BẤT LỢI')) else ('hơi bất lợi – dễ gặp tai nạn, thương tật, nhất là trong các vận, năm nêu dưới, ở mức nhẹ' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG BẤT', X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,8,10),X.CHAR(10.0))),1.0))))) else ('bình thường – lúc thuận lúc không, tùy vận' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRUNG BÌNH', X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,8,10),X.CHAR(10.0))),1.0))))) else ('khá – ít tai nạn, thương tật, ở mức vừa, không phải thế mạnh nổi bật' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('NGHIÊNG THUẬN', X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,8,10),X.CHAR(10.0))),1.0))))) else X.cat('ít tai nạn, thương tật',(' – nhưng có trở ngại (các dấu hiệu bất lợi ở trên)' if X.truthy(X.ISNUMBER(lambda: X.SEARCH('TRỞ NGẠI', X.LEFT(E.v(22,8,10), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(22,8,10),X.CHAR(10.0))),1.0))))) else ''))))))),'.'),X.CHAR(10.0)),'Vì sao – dấu hiệu THUẬN của sách đủ điều kiện: '),E.v(22,8,3)),X.CHAR(10.0)),'Dấu hiệu BẤT LỢI đủ điều kiện: '),E.v(22,8,6)),X.CHAR(10.0)),'(Cách kết luận: SUC_KHOE dòng 9, giả định G17.)'),X.CHAR(10.0)),E.v(22,28,3)),X.CHAR(10.0)),E.v(22,157,3))

def f247(E, r, c):
    # Ô H139
    # = LUAN_GIAI!$J$129
    return E.v(4,129,10)

def f248(E, r, c):
    # Ô C140
    # = SUC_KHOE!$C$38
    return E.v(22,38,3)

def f249(E, r, c):
    # Ô H140
    # = LUAN_GIAI!$J$192
    return E.v(4,192,10)

def f250(E, r, c):
    # Ô C141
    # = SUC_KHOE!$C$43&CHAR(10)&SUC_KHOE!$C$46
    return X.cat(X.cat(E.v(22,43,3),X.CHAR(10.0)),E.v(22,46,3))

def f251(E, r, c):
    # Ô H141
    # = LUAN_GIAI!$J$193
    return E.v(4,193,10)

def f252(E, r, c):
    # Ô C142
    # = SUC_KHOE!$C$56&CHAR(10)&"Cách đọc mức (giả định G17): THUẬN = sao chủ sự là dụng/hỷ thần và dấu hiệu thuận của sách nhiều hơn; THUẬN NHƯNG CÓ TRỞ NGẠI
    #    = sao chủ sự tốt nhưng dấu hiệu bất lợi nhiều hơn; NGHIÊNG THUẬN = sao chủ sự là nhàn thần, dấu hiệu thuận nhiều h"&"ơn; TRUNG BÌNH = ngang nhau; NGH
    #   IÊNG BẤT LỢI = nhàn thần, dấu hiệu bất lợi nhiều hơn; BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG = sao chủ sự là kỵ nhưng dấu hiệu thuận nhiều hơn; BẤT LỢI = sao chủ 
    #   sự là kỵ (kể cả là nhóm nhiều nhất thuộc thực thương/tài/q"&"uan sát – P1 tr.180, giả định G26) và dấu hiệu bất lợi nhiều hơn."&DUNG_HY_KY!$D$195
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(22,56,3),X.CHAR(10.0)),'Cách đọc mức (giả định G17): THUẬN = sao chủ sự là dụng/hỷ thần và dấu hiệu thuận của sách nhiều hơn; THUẬN NHƯNG CÓ TRỞ NGẠI = sao chủ sự tốt nhưng dấu hiệu bất lợi nhiều hơn; NGHIÊNG THUẬN = sao chủ sự là nhàn thần, dấu hiệu thuận nhiều h'),'ơn; TRUNG BÌNH = ngang nhau; NGHIÊNG BẤT LỢI = nhàn thần, dấu hiệu bất lợi nhiều hơn; BẤT LỢI NHƯNG CÓ ĐIỂM SÁNG = sao chủ sự là kỵ nhưng dấu hiệu thuận nhiều hơn; BẤT LỢI = sao chủ sự là kỵ (kể cả là nhóm nhiều nhất thuộc thực thương/tài/q'),'uan sát – P1 tr.180, giả định G26) và dấu hiệu bất lợi nhiều hơn.'),E.v(8,195,4))

def f253(E, r, c):
    # Ô H142
    # = LUAN_GIAI!$J$194
    return E.v(4,194,10)

def f254(E, r, c):
    # Ô B143
    # = MENH_CUC!B21
    return E.v(18,r-122,c)

def f255(E, r, c):
    # Ô C143
    # = MENH_CUC!$C$126&CHAR(10)&MENH_CUC!$C$127&CHAR(10)&"Bảng thực thương – ấn (P1 tr.173): "&MENH_CUC!D21&CHAR(10)&MENH_CUC!G21&CHAR(10)&MENH_CUC!K21&"► Ý 
    #   nghĩa cụ thể: "&IF(MENH_CUC!$C$114="YỂU","“người yểu là người nguyên khí mỏng, dụng thần được ít khí, tứ trụ thiên về khô” (P1 tr.171): sức khỏe dễ su
    #   y kiệt, các năm hạn nặng cần đề phòng; công cụ không kết luận năm mất (P1 tr.201)",IF(MENH_CUC!$C$114="THỌ","“người thọ tức người nguyên khí dày, dụng
    #    thần được sinh mà có lực” (P1 tr.171): sức bền tốt, qua được các năm hạn","cân bằng: hai chiều ngang nhau, tốt xấu tùy vận"))&". "&IF(AND($AB$87="",$
    #   AC$87=""),"Khi nào thể hiện rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận n
    #   ào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu ("&IF($AD$87="","không có",$AD$87)&"); phần 
    #   xấu dễ nổi hơn trong bại vận ("&IF($AE$87="","không có",$AE$87)&"). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.","Khi nào thể hiện rõ: phần tốt nổi l
    #   ên trong đại vận TỐT ("&IF($AB$87="","không có",$AB$87)&"); phần xấu nổi lên trong đại vận XẤU ("&IF($AC$87="","không có",$AC$87)&").")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(18,126,3),X.CHAR(10.0)),E.v(18,127,3)),X.CHAR(10.0)),'Bảng thực thương – ấn (P1 tr.173): '),E.v(18,r-122,c+1)),X.CHAR(10.0)),E.v(18,r-122,c+4)),X.CHAR(10.0)),E.v(18,r-122,c+8)),'► Ý nghĩa cụ thể: '),('“người yểu là người nguyên khí mỏng, dụng thần được ít khí, tứ trụ thiên về khô” (P1 tr.171): sức khỏe dễ suy kiệt, các năm hạn nặng cần đề phòng; công cụ không kết luận năm mất (P1 tr.201)' if X.truthy(X.eq(E.v(18,114,3),'YỂU')) else ('“người thọ tức người nguyên khí dày, dụng thần được sinh mà có lực” (P1 tr.171): sức bền tốt, qua được các năm hạn' if X.truthy(X.eq(E.v(18,114,3),'THỌ')) else 'cân bằng: hai chiều ngang nhau, tốt xấu tùy vận'))),'. '),(X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu (',('không có' if X.truthy(X.eq(E.v(3,87,30),'')) else E.v(3,87,30))),'); phần xấu dễ nổi hơn trong bại vận ('),('không có' if X.truthy(X.eq(E.v(3,87,31),'')) else E.v(3,87,31))),'). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.') if X.truthy(X.AND(X.eq(E.v(3,87,28),''), X.eq(E.v(3,87,29),''))) else X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT (',('không có' if X.truthy(X.eq(E.v(3,87,28),'')) else E.v(3,87,28))),'); phần xấu nổi lên trong đại vận XẤU ('),('không có' if X.truthy(X.eq(E.v(3,87,29),'')) else E.v(3,87,29))),').')))

def f256(E, r, c):
    # Ô H143
    # = LUAN_GIAI!$J$119
    return E.v(4,119,10)

def f257(E, r, c):
    # Ô C144
    # = TRIM(IF(MENH_CUC!C33="","",MENH_CUC!C33&" ("&MENH_CUC!E33&" t.mụ); ")&IF(MENH_CUC!C34="","",MENH_CUC!C34&" ("&MENH_CUC!E34&" t.mụ); ")&IF(MENH_CUC!C35
    #   ="","",MENH_CUC!C35&" ("&MENH_CUC!E35&" t.mụ); ")&IF(MENH_CUC!C36="","",MENH_CUC!C36&" ("&MENH_CUC!E36&" t.mụ); ")&IF(MENH_CUC!C37="","",MENH_CUC!C37&
    #   " ("&MENH_CUC!E37&" t.mụ); ")&IF(MENH_CUC!C38="","",MENH_CUC!C38&" ("&MENH_CUC!E38&" t.mụ); ")&IF(MENH_CUC!C39="","",MENH_CUC!C39&" ("&MENH_CUC!E39&" 
    #   t.mụ); ")&IF(MENH_CUC!C40="","",MENH_CUC!C40&" ("&MENH_CUC!E40&" t.mụ); ")&IF(MENH_CUC!C41="","",MENH_CUC!C41&" ("&MENH_CUC!E41&" t.mụ); ")&IF(MENH_CU
    #   C!C42="","",MENH_CUC!C42&" ("&MENH_CUC!E42&" t.mụ); ")&IF(MENH_CUC!C43="","",MENH_CUC!C43&" ("&MENH_CUC!E43&" t.mụ); ")&IF(MENH_CUC!C44="","",MENH_CUC
    #   !C44&" ("&MENH_CUC!E44&" t.mụ); ")&IF(MENH_CUC!C45="","",MENH_CUC!C45&" ("&MENH_CUC!E45&" t.mụ); ")&IF(MENH_CUC!C46="","",MENH_CUC!C46&" ("&MENH_CUC!E
    #   46&" t.mụ); ")&IF(MENH_CUC!C47="","",MENH_CUC!C47&" ("&MENH_CUC!E47&" t.mụ); "))&" – Chỉ là dấu hiệu theo sách; phải xét có cứu hay không (P1 tr.201).
    #   "&CHAR(10)&"Năm lo về người thân – sinh mệnh (tính cả năm đã qua, từ 1 dấu hiệu, tuổi mụ 1–90; P1 tr.201; P2 tr.288): "&LUU_NIEN!$DK$31&CHAR(10)&"Năm 
    #   có kình dương gặp/xung tuế quân hoặc một xung ba – tai vạ nặng (P2 tr.316, 566, 282-283): "&LUU_NIEN!$DL$31&CHAR(10)&"(Các danh sách này chỉ nêu năm c
    #   ó dấu hiệu theo sách; không nói được người cụ thể hay việc cụ thể. Người ứng xem LUU_NIEN cột DN.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(18,r-111,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-111,c),' ('),E.v(18,r-111,c+2)),' t.mụ); ')),('' if X.truthy(X.eq(E.v(18,r-110,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-110,c),' ('),E.v(18,r-110,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-109,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-109,c),' ('),E.v(18,r-109,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-108,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-108,c),' ('),E.v(18,r-108,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-107,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-107,c),' ('),E.v(18,r-107,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-106,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-106,c),' ('),E.v(18,r-106,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-105,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-105,c),' ('),E.v(18,r-105,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-104,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-104,c),' ('),E.v(18,r-104,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-103,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-103,c),' ('),E.v(18,r-103,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-102,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-102,c),' ('),E.v(18,r-102,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-101,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-101,c),' ('),E.v(18,r-101,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-100,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-100,c),' ('),E.v(18,r-100,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-99,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-99,c),' ('),E.v(18,r-99,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-98,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-98,c),' ('),E.v(18,r-98,c+2)),' t.mụ); '))),('' if X.truthy(X.eq(E.v(18,r-97,c),'')) else X.cat(X.cat(X.cat(E.v(18,r-97,c),' ('),E.v(18,r-97,c+2)),' t.mụ); ')))),' – Chỉ là dấu hiệu theo sách; phải xét có cứu hay không (P1 tr.201).'),X.CHAR(10.0)),'Năm lo về người thân – sinh mệnh (tính cả năm đã qua, từ 1 dấu hiệu, tuổi mụ 1–90; P1 tr.201; P2 tr.288): '),E.v(15,31,115)),X.CHAR(10.0)),'Năm có kình dương gặp/xung tuế quân hoặc một xung ba – tai vạ nặng (P2 tr.316, 566, 282-283): '),E.v(15,31,116)),X.CHAR(10.0)),'(Các danh sách này chỉ nêu năm có dấu hiệu theo sách; không nói được người cụ thể hay việc cụ thể. Người ứng xem LUU_NIEN cột DN.)')

def f258(E, r, c):
    # Ô H144
    # = LUAN_GIAI!$J$180
    return E.v(4,180,10)

def f259(E, r, c):
    # Ô C145
    # = "[Thấp] "&THAM_KHAO!$E$11
    return X.cat('[Thấp] ',E.v(28,11,5))

def f260(E, r, c):
    # Ô H145
    # = LUAN_GIAI!$J$206
    return E.v(4,206,10)

def f261(E, r, c):
    # Ô C147
    # = "Nguồn: tài liệu “Những cấu trúc bát tự dễ sinh ung tượng” do người dùng cung cấp (13/10/2026) – NGOÀI sách Thiệu Vĩ Hoa, xếp tầng [Tham khảo]. Công c
    #   ụ chỉ đọc số độ ngũ hành của PP2 (PP2_DINH_LUONG dòng 91-100) và thập thần của can lộ + bản"&" khí, đối chiếu 4 cấu trúc; mức RÕ / MỘT PHẦN / KHÔNG th
    #   eo ngưỡng ghi ở từng dòng (giả định G48). Đây là “nền đất” cần giữ gìn, KHÔNG phải chẩn đoán, KHÔNG phải án tử; không làm thay đổi kết luận chính (mục
    #    SỨC KHỎE ở trên)."
    return X.cat('Nguồn: tài liệu “Những cấu trúc bát tự dễ sinh ung tượng” do người dùng cung cấp (13/10/2026) – NGOÀI sách Thiệu Vĩ Hoa, xếp tầng [Tham khảo]. Công cụ chỉ đọc số độ ngũ hành của PP2 (PP2_DINH_LUONG dòng 91-100) và thập thần của can lộ + bản',' khí, đối chiếu 4 cấu trúc; mức RÕ / MỘT PHẦN / KHÔNG theo ngưỡng ghi ở từng dòng (giả định G48). Đây là “nền đất” cần giữ gìn, KHÔNG phải chẩn đoán, KHÔNG phải án tử; không làm thay đổi kết luận chính (mục SỨC KHỎE ở trên).')

def f262(E, r, c):
    # Ô C148
    # = "Mộc "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(0,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" · Hỏa "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$9
    #   5,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" · Thổ "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(2,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" · Ki
    #   m "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(3,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" · Thủy "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,
    #   MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" | Nhóm Ấn "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUO
    #   NG!$C$100,0),"0%")&", Tài "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+2,5),5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", Quan sá
    #   t "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&" | Số chữ (can lộ + bản khí): thi
    #   ên ấn "&(IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NO
    #   T(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$C$24))),1,0)
    #   +IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$F$24))),1,0))&", thất sát "&(IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$12
    #   ))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_
    #   GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$24))),1,0)
    #   +IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$
    #   24))),1,0))&", thực thần "&(IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D
    #   $12))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",T
    #   RUNG_GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$E$24)
    #   )),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$24))),1,0))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Mộc ',X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(0.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),' · Hỏa '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),' · Thổ '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(2.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),' · Kim '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(3.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),' · Thủy '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),' | Nhóm Ấn '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Tài '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),2.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Quan sát '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),' | Số chữ (can lộ + bản khí): thiên ấn '),X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,6))))) else 0.0))),', thất sát '),X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,6))))) else 0.0))),', thực thần '),X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,6))))) else 0.0)))

def f263(E, r, c):
    # Ô C149
    # = "["&AG149&"] "&"Thủy "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", Hỏa "&TEXT(IFERROR(INDEX(PP2_DINH_L
    #   UONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&". "&IF(AG149="KHÔNG","Không thuộc cấu trúc này.","Ý nghĩa: Thủy chủ hàn – trệ – tích, Hỏa
    #    chủ hóa – tán – sinh mệnh; Thủy nhiều mà Hỏa yếu hoặc bị khắc chết thì cơ thể mất năng lực “đốt – chuyển – hóa” (hóa khí bất thành – sinh cơ suy kiệt
    #   ): sợ lạnh, dễ mệt, dễ tích đàm, u, nang, khối; bện"&"h tiến chậm nhưng dai dẳng. Là nền đất sinh ung, không phải án tử.")&CHAR(10)&"Điều kiện: RÕ: Th
    #   ủy ≥ 40% và Hỏa ≤ 10%; MỘT PHẦN: Thủy ≥ 25% và Hỏa ≤ ½ Thủy."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('[',E.v(3,r,c+30)),'] '),'Thủy '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Hỏa '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),'. '),('Không thuộc cấu trúc này.' if X.truthy(X.eq(E.v(3,r,c+30),'KHÔNG')) else X.cat('Ý nghĩa: Thủy chủ hàn – trệ – tích, Hỏa chủ hóa – tán – sinh mệnh; Thủy nhiều mà Hỏa yếu hoặc bị khắc chết thì cơ thể mất năng lực “đốt – chuyển – hóa” (hóa khí bất thành – sinh cơ suy kiệt): sợ lạnh, dễ mệt, dễ tích đàm, u, nang, khối; bện','h tiến chậm nhưng dai dẳng. Là nền đất sinh ung, không phải án tử.'))),X.CHAR(10.0)),'Điều kiện: RÕ: Thủy ≥ 40% và Hỏa ≤ 10%; MỘT PHẦN: Thủy ≥ 25% và Hỏa ≤ ½ Thủy.')

def f264(E, r, c):
    # Ô AG149
    # = IF(AND(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DIN
    #   H_LUONG!$C$100,0)<=0.1),"RÕ",IF(AND(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.25,IFERROR(INDEX(PP2_DINH_LUONG!$
    #   F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0)<=IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0)/2),"MỘT PHẦN","KHÔNG"))
    return ('RÕ' if X.truthy(X.AND(X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('MỘT PHẦN' if X.truthy(X.AND(X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.25), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),X.div(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),2.0)))) else 'KHÔNG'))

def f265(E, r, c):
    # Ô C150
    # = "["&AG150&"] "&"Thiên ấn "&(IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$1
    #   2))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG
    #   _GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$E$24))),1,0
    #   )+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$F$24))),1,0))&" chữ, nhóm Ấn "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F
    #   $91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", Tài + Quan "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUN
    #   G_GIAN!$E$7+2,5),5)+1)/PP2_DINH_LUONG!$C$100,0)+IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0),"
    #   0%")&". "&IF(AG150="KHÔNG","Không thuộc cấu trúc này.","Ý nghĩa: Thiên ấn là tư duy nội tại, tự phòng vệ; quá mạnh mà không có Tài để tiết, Quan để đi
    #   ều thì người sống trong đầu, tích tụ cảm xúc, uất mà không xả, stress không thoát ra – kiểu “bệnh sinh từ tâm”, thân thể “tự đóng cửa – tự chống c"&"h
    #   ính mình”.")&CHAR(10)&"Điều kiện: RÕ: thiên ấn ≥ 2 chữ, Ấn ≥ 35%, Tài + Quan ≤ 15%; MỘT PHẦN: thiên ấn ≥ 2 chữ, Ấn ≥ 25%, Tài + Quan ≤ 25%."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('[',E.v(3,r,c+30)),'] '),'Thiên ấn '),X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,6))))) else 0.0))),' chữ, nhóm Ấn '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Tài + Quan '),X.TEXT(X.add(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),2.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0)), '0%')),'. '),('Không thuộc cấu trúc này.' if X.truthy(X.eq(E.v(3,r,c+30),'KHÔNG')) else X.cat('Ý nghĩa: Thiên ấn là tư duy nội tại, tự phòng vệ; quá mạnh mà không có Tài để tiết, Quan để điều thì người sống trong đầu, tích tụ cảm xúc, uất mà không xả, stress không thoát ra – kiểu “bệnh sinh từ tâm”, thân thể “tự đóng cửa – tự chống c','hính mình”.'))),X.CHAR(10.0)),'Điều kiện: RÕ: thiên ấn ≥ 2 chữ, Ấn ≥ 35%, Tài + Quan ≤ 15%; MỘT PHẦN: thiên ấn ≥ 2 chữ, Ấn ≥ 25%, Tài + Quan ≤ 25%.')

def f266(E, r, c):
    # Ô AG150
    # = IF(AND((IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT
    #   (NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$C$24))),1,0)+
    #   IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_LI
    #   EU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$F$24))),1,0))>=2,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),
    #   5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.35,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+2,5),5)+1)/PP2_DINH_LUONG!$C$100,0)+IFERROR(INDE
    #   X(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<=0.15),"RÕ",IF(AND((IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",
    #   TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("
    #   Thiên ấn",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GI
    #   AN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thiên ấn",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thiên ấn
    #   ",TRUNG_GIAN!$F$24))),1,0))>=2,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.25,IFERROR(INDE
    #   X(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+2,5),5)+1)/PP2_DINH_LUONG!$C$100,0)+IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$
    #   E$7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<=0.25),"MỘT PHẦN","KHÔNG"))
    return ('RÕ' if X.truthy(X.AND(X.ge(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,6))))) else 0.0)),2.0), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.35), X.le(X.add(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),2.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0)),0.15))) else ('MỘT PHẦN' if X.truthy(X.AND(X.ge(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thiên ấn', E.v(23,24,6))))) else 0.0)),2.0), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.25), X.le(X.add(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),2.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0)),0.25))) else 'KHÔNG'))

def f267(E, r, c):
    # Ô C151
    # = "["&AG151&"] "&"Thất sát "&(IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$1
    #   2))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG
    #   _GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$E$24))),1,0
    #   )+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$24))),1,0))&" chữ, Quan sát "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$
    #   F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", Ấn "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!
    #   $E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", thực thần "&(IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMB
    #   ER(SEARCH("Thực thần",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$12))),1,0)+IF(AND(T
    #   RUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(S
    #   EARCH("Thực thần",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$24))),1,0))&" chữ. "&IF
    #   (AG151="KHÔNG","Không thuộc cấu trúc này.","Ý nghĩa: Thất sát là áp lực, đối kháng, sống trong trạng thái “phải gồng”; sát mạnh mà không có Ấn hóa, kh
    #   ông có Thực thần tiết thì cơ thể sống lâu trong chế độ sinh tồn (sát khí nhập thân – khí huyết hao tổn; y học: cortisol cao, miễn dịch"&" rối loạn).")
    #   &CHAR(10)&"Điều kiện: RÕ: (thất sát ≥ 2 chữ, hoặc 1 chữ và Quan sát ≥ 30%), Ấn < 10%, không có thực thần; MỘT PHẦN: thất sát ≥ 1 chữ, Quan sát ≥ 20% v
    #   à thiếu Ấn hoặc thiếu thực thần. Dù gốc không có, năm / vận THẤT SÁT đến khi Ấn < 10% là thời điểm KÍ"&"CH HOẠT (dòng dưới)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('[',E.v(3,r,c+30)),'] '),'Thất sát '),X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,6))))) else 0.0))),' chữ, Quan sát '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Ấn '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', thực thần '),X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,6))))) else 0.0))),' chữ. '),('Không thuộc cấu trúc này.' if X.truthy(X.eq(E.v(3,r,c+30),'KHÔNG')) else X.cat('Ý nghĩa: Thất sát là áp lực, đối kháng, sống trong trạng thái “phải gồng”; sát mạnh mà không có Ấn hóa, không có Thực thần tiết thì cơ thể sống lâu trong chế độ sinh tồn (sát khí nhập thân – khí huyết hao tổn; y học: cortisol cao, miễn dịch',' rối loạn).'))),X.CHAR(10.0)),'Điều kiện: RÕ: (thất sát ≥ 2 chữ, hoặc 1 chữ và Quan sát ≥ 30%), Ấn < 10%, không có thực thần; MỘT PHẦN: thất sát ≥ 1 chữ, Quan sát ≥ 20% và thiếu Ấn hoặc thiếu thực thần. Dù gốc không có, năm / vận THẤT SÁT đến khi Ấn < 10% là thời điểm KÍ'),'CH HOẠT (dòng dưới).')

def f268(E, r, c):
    # Ô AG151
    # = IF(AND(OR((IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$12))),1,0)+IF(AND(
    #   NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$24))),1,
    #   0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP
    #   _LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$24))),1,0))>=2,AND((IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$12))),1
    #   ,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!
    #   $F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$24))),1,0)+IF(A
    #   ND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$24)))
    #   ,1,0))>=1,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.3)),IFERROR(INDEX(PP2_DINH_LUONG!$F$
    #   91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1,(IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TR
    #   UE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$12))),1,
    #   0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE()
    #   ,ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$24))),1,0))=
    #   0),"RÕ",IF(AND((IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$12))),1,0)+IF
    #   (AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$C$24)
    #   )),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT
    #   (NHAP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thất sát",TRUNG_GIAN!$F$24))),1,0))>=1,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E
    #   $7+3,5),5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.2,OR(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.
    #   1,(IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D$12))),1,0)+IF(AND(NOT(NH
    #   AP_LIEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$12))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$C$24))),1,0)+I
    #   F(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$D$24))),1,0)+IF(AND(TRUE(),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$E$24))),1,0)+IF(AND(NOT(NHAP_L
    #   IEU!$C$9="Không rõ giờ"),ISNUMBER(SEARCH("Thực thần",TRUNG_GIAN!$F$24))),1,0))=0)),"MỘT PHẦN","KHÔNG"))
    return ('RÕ' if X.truthy(X.AND(X.OR(X.ge(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,6))))) else 0.0)),2.0), X.AND(X.ge(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,6))))) else 0.0)),1.0), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.3))), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1), X.eq(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,6))))) else 0.0)),0.0))) else ('MỘT PHẦN' if X.truthy(X.AND(X.ge(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thất sát', E.v(23,24,6))))) else 0.0)),1.0), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),3.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.2), X.OR(X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1), X.eq(X.add(X.add(X.add(X.add(X.add(X.add((1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,3))))) else 0.0),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,12,6))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,3))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,4))))) else 0.0)),(1.0 if X.truthy(X.AND(X.TRUE(), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,5))))) else 0.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.ISNUMBER(lambda: X.SEARCH('Thực thần', E.v(23,24,6))))) else 0.0)),0.0)))) else 'KHÔNG'))

def f269(E, r, c):
    # Ô C152
    # = "["&AG152&"] "&"Thổ "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(2,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", Hỏa "&TEXT(IFERROR(INDEX(PP2_DINH_LU
    #   ONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", Mộc "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(0,5)+1)/PP2_DINH_LUONG!$C$100,0)
    #   ,"0%")&", Kim "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(3,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&". "&IF(AG152="KHÔNG","Không thuộc cấu trúc n
    #   ày.","Ý nghĩa: Thổ nhiều mà không sinh Kim, không được Mộc khơi, không được Hỏa hóa thì sinh đàm thấp hóa thực: u, nang, khối, tích tụ lâu ngày – “khô
    #   ng phải một ngày mà có”.")&CHAR(10)&"Điều kiện: RÕ: Thổ ≥ 35%, Hỏa, Mộc, Kim đều ≤ 10%; MỘT PHẦN: Thổ ≥ 30%, Hỏa ≤ 10% và Mộc hoặc Kim ≤ 10%."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('[',E.v(3,r,c+30)),'] '),'Thổ '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(2.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Hỏa '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Mộc '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(0.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Kim '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(3.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),'. '),('Không thuộc cấu trúc này.' if X.truthy(X.eq(E.v(3,r,c+30),'KHÔNG')) else 'Ý nghĩa: Thổ nhiều mà không sinh Kim, không được Mộc khơi, không được Hỏa hóa thì sinh đàm thấp hóa thực: u, nang, khối, tích tụ lâu ngày – “không phải một ngày mà có”.')),X.CHAR(10.0)),'Điều kiện: RÕ: Thổ ≥ 35%, Hỏa, Mộc, Kim đều ≤ 10%; MỘT PHẦN: Thổ ≥ 30%, Hỏa ≤ 10% và Mộc hoặc Kim ≤ 10%.')

def f270(E, r, c):
    # Ô AG152
    # = IF(AND(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(2,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.35,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DI
    #   NH_LUONG!$C$100,0)<=0.1,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(0,5)+1)/PP2_DINH_LUONG!$C$100,0)<=0.1,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MO
    #   D(3,5)+1)/PP2_DINH_LUONG!$C$100,0)<=0.1),"RÕ",IF(AND(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(2,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.3,IFERROR(INDEX(
    #   PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0)<=0.1,OR(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(0,5)+1)/PP2_DINH_LUONG!$C$100,0)<
    #   =0.1,IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(3,5)+1)/PP2_DINH_LUONG!$C$100,0)<=0.1)),"MỘT PHẦN","KHÔNG"))
    return ('RÕ' if X.truthy(X.AND(X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(2.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.35), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(0.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(3.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('MỘT PHẦN' if X.truthy(X.AND(X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(2.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.3), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1), X.OR(X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(0.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1), X.le(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(3.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1)))) else 'KHÔNG'))

def f271(E, r, c):
    # Ô C153
    # = IF(AH214&AH215&AH216&AH217&AH218&AH219&AH220&AH221="","Không có đại vận nào làm nặng thêm hoặc kích hoạt các cấu trúc trên.",AH214&AH215&AH216&AH217&A
    #   H218&AH219&AH220&AH221)&"(Nặng thêm = vận bồi đúng hành gây mất cân bằng; bù / điều tiết / hóa = vận có hành giải. Năm vận hành lấy từ DAI_VAN.)"
    return X.cat(('Không có đại vận nào làm nặng thêm hoặc kích hoạt các cấu trúc trên.' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+61,c+31),E.v(3,r+62,c+31)),E.v(3,r+63,c+31)),E.v(3,r+64,c+31)),E.v(3,r+65,c+31)),E.v(3,r+66,c+31)),E.v(3,r+67,c+31)),E.v(3,r+68,c+31)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+61,c+31),E.v(3,r+62,c+31)),E.v(3,r+63,c+31)),E.v(3,r+64,c+31)),E.v(3,r+65,c+31)),E.v(3,r+66,c+31)),E.v(3,r+67,c+31)),E.v(3,r+68,c+31))),'(Nặng thêm = vận bồi đúng hành gây mất cân bằng; bù / điều tiết / hóa = vận có hành giải. Năm vận hành lấy từ DAI_VAN.)')

def f272(E, r, c):
    # Ô C154
    # = IF(AG222&AG223&AG224&AG225&AG226&AG227&AG228&AG229&AG230&AG231&AG232&AG233&AG234&AG235&AG236&AG237&AG238&AG239&AG240&AG241&AG242&AG243&AG244&AG245&AG2
    #   46&AG247&AG248&AG249&AG250&AG251="","Không có năm nào kích hoạt trong 30 năm tới.",AG222&AG223&AG224&AG225&AG226&AG227&AG228&AG229&AG230&AG231&AG232&A
    #   G233&AG234&AG235&AG236&AG237&AG238&AG239&AG240&AG241&AG242&AG243&AG244&AG245&AG246&AG247&AG248&AG249&AG250&AG251)&CHAR(10)&"UT3 = năm thất sát đến khi
    #    gốc thiếu Ấn (< 10%); UT1 = năm can và chi đều Thủy khi lá số có UT1. Nên khám định kỳ, giữ sức trong các năm này."
    return X.cat(X.cat(('Không có năm nào kích hoạt trong 30 năm tới.' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+68,c+30),E.v(3,r+69,c+30)),E.v(3,r+70,c+30)),E.v(3,r+71,c+30)),E.v(3,r+72,c+30)),E.v(3,r+73,c+30)),E.v(3,r+74,c+30)),E.v(3,r+75,c+30)),E.v(3,r+76,c+30)),E.v(3,r+77,c+30)),E.v(3,r+78,c+30)),E.v(3,r+79,c+30)),E.v(3,r+80,c+30)),E.v(3,r+81,c+30)),E.v(3,r+82,c+30)),E.v(3,r+83,c+30)),E.v(3,r+84,c+30)),E.v(3,r+85,c+30)),E.v(3,r+86,c+30)),E.v(3,r+87,c+30)),E.v(3,r+88,c+30)),E.v(3,r+89,c+30)),E.v(3,r+90,c+30)),E.v(3,r+91,c+30)),E.v(3,r+92,c+30)),E.v(3,r+93,c+30)),E.v(3,r+94,c+30)),E.v(3,r+95,c+30)),E.v(3,r+96,c+30)),E.v(3,r+97,c+30)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+68,c+30),E.v(3,r+69,c+30)),E.v(3,r+70,c+30)),E.v(3,r+71,c+30)),E.v(3,r+72,c+30)),E.v(3,r+73,c+30)),E.v(3,r+74,c+30)),E.v(3,r+75,c+30)),E.v(3,r+76,c+30)),E.v(3,r+77,c+30)),E.v(3,r+78,c+30)),E.v(3,r+79,c+30)),E.v(3,r+80,c+30)),E.v(3,r+81,c+30)),E.v(3,r+82,c+30)),E.v(3,r+83,c+30)),E.v(3,r+84,c+30)),E.v(3,r+85,c+30)),E.v(3,r+86,c+30)),E.v(3,r+87,c+30)),E.v(3,r+88,c+30)),E.v(3,r+89,c+30)),E.v(3,r+90,c+30)),E.v(3,r+91,c+30)),E.v(3,r+92,c+30)),E.v(3,r+93,c+30)),E.v(3,r+94,c+30)),E.v(3,r+95,c+30)),E.v(3,r+96,c+30)),E.v(3,r+97,c+30))),X.CHAR(10.0)),'UT3 = năm thất sát đến khi gốc thiếu Ấn (< 10%); UT1 = năm can và chi đều Thủy khi lá số có UT1. Nên khám định kỳ, giữ sức trong các năm này.')

def f273(E, r, c):
    # Ô C155
    # = IF(OR($AG$149<>"KHÔNG",$AG$150<>"KHÔNG",$AG$151<>"KHÔNG",$AG$152<>"KHÔNG"),"[Tham khảo] Lá số có: "&IF($AG$149<>"KHÔNG","UT1 ("&$AG$149&") ","")&IF($A
    #   G$150<>"KHÔNG","UT2 ("&$AG$150&") ","")&IF($AG$151<>"KHÔNG","UT3 ("&$AG$151&") ","")&IF($AG$152<>"KHÔNG","UT4 ("&$AG$152&") ","")&"– là nền đất cần gi
    #   ữ gìn, không phải án tử. Gợi ý: "&IF($AG$149<>"KHÔNG","Bổ Hỏa: giữ ấm, vận động đều, ra nắng, tránh đồ lạnh; ","")&IF($AG$150<>"KHÔNG","Tìm lối xả: ch
    #   ia sẻ, vận động, việc hướng ngoại (tài – quan), không ôm lo nghĩ; ","")&IF($AG$151<>"KHÔNG","Giảm áp lực, ngủ đủ, có người đỡ đần (ấn), có lối tiết (t
    #   hực thần: sở thích, sáng tạo); ","")&IF($AG$152<>"KHÔNG","Ăn uống nhẹ, giảm đồ ngọt – béo – ẩm, vận động, giữ nơi ở khô thoáng; ","")&"khám sức khỏe đ
    #   ịnh kỳ, nhất là các đại vận / năm ở hai dòng trên.","[Tham khảo] Lá số không thuộc 4 cấu trúc dễ sinh ung tượng"&IF(LEN(AG222&AG223&AG224&AG225&AG226&
    #   AG227&AG228&AG229&AG230&AG231&AG232&AG233&AG234&AG235&AG236&AG237&AG238&AG239&AG240&AG241&AG242&AG243&AG244&AG245&AG246&AG247&AG248&AG249&AG250&AG251)
    #   >0,"; riêng các năm kích hoạt ở dòng trên nên giữ sức.","."))&" Không thay chẩn đoán y khoa."
    return X.cat((X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('[Tham khảo] Lá số có: ',(X.cat(X.cat('UT1 (',E.v(3,149,33)),') ') if X.truthy(X.ne(E.v(3,149,33),'KHÔNG')) else '')),(X.cat(X.cat('UT2 (',E.v(3,150,33)),') ') if X.truthy(X.ne(E.v(3,150,33),'KHÔNG')) else '')),(X.cat(X.cat('UT3 (',E.v(3,151,33)),') ') if X.truthy(X.ne(E.v(3,151,33),'KHÔNG')) else '')),(X.cat(X.cat('UT4 (',E.v(3,152,33)),') ') if X.truthy(X.ne(E.v(3,152,33),'KHÔNG')) else '')),'– là nền đất cần giữ gìn, không phải án tử. Gợi ý: '),('Bổ Hỏa: giữ ấm, vận động đều, ra nắng, tránh đồ lạnh; ' if X.truthy(X.ne(E.v(3,149,33),'KHÔNG')) else '')),('Tìm lối xả: chia sẻ, vận động, việc hướng ngoại (tài – quan), không ôm lo nghĩ; ' if X.truthy(X.ne(E.v(3,150,33),'KHÔNG')) else '')),('Giảm áp lực, ngủ đủ, có người đỡ đần (ấn), có lối tiết (thực thần: sở thích, sáng tạo); ' if X.truthy(X.ne(E.v(3,151,33),'KHÔNG')) else '')),('Ăn uống nhẹ, giảm đồ ngọt – béo – ẩm, vận động, giữ nơi ở khô thoáng; ' if X.truthy(X.ne(E.v(3,152,33),'KHÔNG')) else '')),'khám sức khỏe định kỳ, nhất là các đại vận / năm ở hai dòng trên.') if X.truthy(X.OR(X.ne(E.v(3,149,33),'KHÔNG'), X.ne(E.v(3,150,33),'KHÔNG'), X.ne(E.v(3,151,33),'KHÔNG'), X.ne(E.v(3,152,33),'KHÔNG'))) else X.cat('[Tham khảo] Lá số không thuộc 4 cấu trúc dễ sinh ung tượng',('; riêng các năm kích hoạt ở dòng trên nên giữ sức.' if X.truthy(X.gt(X.LEN(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+67,c+30),E.v(3,r+68,c+30)),E.v(3,r+69,c+30)),E.v(3,r+70,c+30)),E.v(3,r+71,c+30)),E.v(3,r+72,c+30)),E.v(3,r+73,c+30)),E.v(3,r+74,c+30)),E.v(3,r+75,c+30)),E.v(3,r+76,c+30)),E.v(3,r+77,c+30)),E.v(3,r+78,c+30)),E.v(3,r+79,c+30)),E.v(3,r+80,c+30)),E.v(3,r+81,c+30)),E.v(3,r+82,c+30)),E.v(3,r+83,c+30)),E.v(3,r+84,c+30)),E.v(3,r+85,c+30)),E.v(3,r+86,c+30)),E.v(3,r+87,c+30)),E.v(3,r+88,c+30)),E.v(3,r+89,c+30)),E.v(3,r+90,c+30)),E.v(3,r+91,c+30)),E.v(3,r+92,c+30)),E.v(3,r+93,c+30)),E.v(3,r+94,c+30)),E.v(3,r+95,c+30)),E.v(3,r+96,c+30))),0.0)) else '.'))),' Không thay chẩn đoán y khoa.')

def f274(E, r, c):
    # Ô C156, C157
    # = TRIM(IF(SUC_KHOE!$L$219="ĐỦ","• "&SUC_KHOE!$G$219&" (HG tr.201) ","")&IF(SUC_KHOE!$L$220="ĐỦ","• "&SUC_KHOE!$G$220&" (HG tr.201) ","")&IF(SUC_KHOE!$L$
    #   221="ĐỦ","• "&SUC_KHOE!$G$221&" (HG tr.201) ","")&IF(SUC_KHOE!$L$222="ĐỦ","• "&SUC_KHOE!$G$222&" (HG tr.201) ","")&IF(SUC_KHOE!$L$223="ĐỦ","• "&SUC_KH
    #   OE!$G$223&" (HG tr.202) ","")&IF(SUC_KHOE!$L$224="ĐỦ","• "&SUC_KHOE!$G$224&" (HG tr.202) ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('• ',E.v(22,219,7)),' (HG tr.201) ') if X.truthy(X.eq(E.v(22,219,12),'ĐỦ')) else ''),(X.cat(X.cat('• ',E.v(22,220,7)),' (HG tr.201) ') if X.truthy(X.eq(E.v(22,220,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(22,221,7)),' (HG tr.201) ') if X.truthy(X.eq(E.v(22,221,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(22,222,7)),' (HG tr.201) ') if X.truthy(X.eq(E.v(22,222,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(22,223,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(22,223,12),'ĐỦ')) else '')),(X.cat(X.cat('• ',E.v(22,224,7)),' (HG tr.202) ') if X.truthy(X.eq(E.v(22,224,12),'ĐỦ')) else '')))

def f275(E, r, c):
    # Ô C160
    # = ("Đang hành: "&C163&". "&DUNG_HY_KY!$C$189&" "&DUNG_HY_KY!$C$175&DUNG_HY_KY!$D$193)&CHAR(10)&(DUNG_HY_KY!$C$182&" "&C140&DUNG_HY_KY!$D$194)&CHAR(10)&"
    #   Cát – hung chung (P1 tr.173): "&LEFT(C173,IFERROR(FIND(CHAR(10),C173)-1,LEN(C173)))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Đang hành: ',E.v(3,r+3,c)),'. '),E.v(8,189,3)),' '),E.v(8,175,3)),E.v(8,193,4)),X.CHAR(10.0)),X.cat(X.cat(X.cat(E.v(8,182,3),' '),E.v(3,r-20,c)),E.v(8,194,4))),X.CHAR(10.0)),'Cát – hung chung (P1 tr.173): '),X.LEFT(E.v(3,r+13,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+13,c)),1.0), lambda: X.LEN(E.v(3,r+13,c)))))

def f276(E, r, c):
    # Ô C162
    # = IF(LEFT(DAI_VAN!$C$16,8)="CẢNH BÁO","⚠ "&DAI_VAN!$C$16&CHAR(10),"")&IF(AND(DAI_VAN!$C$17="",DAI_VAN!$C$18="",ISNUMBER(NHAP_LIEU!$C$53)),"[Số ngày đến 
    #   tiết suy từ tứ trụ + năm sinh: ngày sinh "&NHAP_LIEU!$D$52&"] ","")&"Tứ trụ lấy từ: "&NHAP_LIEU!$C$41&". "&IF(DAI_VAN!$C$8="LỖI","LỖI NGÀY GIỜ SINH – 
    #   kiểm tra NHAP_LIEU C26:G26. ","")&IF(DAI_VAN!$C$20="","CHƯA NHẬP ngày giờ sinh, số ngày đến tiết hoặc năm sinh khớp trụ năm (NHAP_LIEU C26 / C28 / C30
    #   ) → tạm lấy khởi vận 1 tuổi. ","")&"Tuổi khởi vận "&DAI_VAN!$C$21&" (chính xác "&IF(DAI_VAN!$C$20="","—",TEXT(DAI_VAN!$C$20,"0.00"))&" năm); chiều "&I
    #   F(DAI_VAN!$C$7=1,"thuận","nghịch")&"; năm sinh "&DAI_VAN!$C$22&". 8 vận: "&DAI_VAN!C35&" ("&DAI_VAN!C36&")"&", "&DAI_VAN!D35&" ("&DAI_VAN!D36&")"&", "
    #   &DAI_VAN!E35&" ("&DAI_VAN!E36&")"&", "&DAI_VAN!F35&" ("&DAI_VAN!F36&")"&", "&DAI_VAN!G35&" ("&DAI_VAN!G36&")"&", "&DAI_VAN!H35&" ("&DAI_VAN!H36&")"&",
    #    "&DAI_VAN!I35&" ("&DAI_VAN!I36&")"&", "&DAI_VAN!J35&" ("&DAI_VAN!J36&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('⚠ ',E.v(14,16,3)),X.CHAR(10.0)) if X.truthy(X.eq(X.LEFT(E.v(14,16,3), 8.0),'CẢNH BÁO')) else ''),(X.cat(X.cat('[Số ngày đến tiết suy từ tứ trụ + năm sinh: ngày sinh ',E.v(2,52,4)),'] ') if X.truthy(X.AND(X.eq(E.v(14,17,3),''), X.eq(E.v(14,18,3),''), X.ISNUMBER(lambda: E.v(2,53,3)))) else '')),'Tứ trụ lấy từ: '),E.v(2,41,3)),'. '),('LỖI NGÀY GIỜ SINH – kiểm tra NHAP_LIEU C26:G26. ' if X.truthy(X.eq(E.v(14,8,3),'LỖI')) else '')),('CHƯA NHẬP ngày giờ sinh, số ngày đến tiết hoặc năm sinh khớp trụ năm (NHAP_LIEU C26 / C28 / C30) → tạm lấy khởi vận 1 tuổi. ' if X.truthy(X.eq(E.v(14,20,3),'')) else '')),'Tuổi khởi vận '),E.v(14,21,3)),' (chính xác '),('—' if X.truthy(X.eq(E.v(14,20,3),'')) else X.TEXT(E.v(14,20,3), '0.00'))),' năm); chiều '),('thuận' if X.truthy(X.eq(E.v(14,7,3),1.0)) else 'nghịch')),'; năm sinh '),E.v(14,22,3)),'. 8 vận: '),E.v(14,r-127,c)),' ('),E.v(14,r-126,c)),')'),', '),E.v(14,r-127,c+1)),' ('),E.v(14,r-126,c+1)),')'),', '),E.v(14,r-127,c+2)),' ('),E.v(14,r-126,c+2)),')'),', '),E.v(14,r-127,c+3)),' ('),E.v(14,r-126,c+3)),')'),', '),E.v(14,r-127,c+4)),' ('),E.v(14,r-126,c+4)),')'),', '),E.v(14,r-127,c+5)),' ('),E.v(14,r-126,c+5)),')'),', '),E.v(14,r-127,c+6)),' ('),E.v(14,r-126,c+6)),')'),', '),E.v(14,r-127,c+7)),' ('),E.v(14,r-126,c+7)),')')

def f277(E, r, c):
    # Ô H162
    # = LUAN_GIAI!$J$165
    return E.v(4,165,10)

def f278(E, r, c):
    # Ô C163
    # = IF(DAI_VAN!$C$25>10,"Năm xem "&DAI_VAN!$C$23&" (tuổi "&DAI_VAN!$C$24&") vượt quá 10 đại vận đã lập (đến "&(DAI_VAN!$L$36+9)&" tuổi) – chọn năm xem nhỏ
    #    hơn ở NHAP_LIEU C31.",IF(DAI_VAN!$C$25=0,"Chưa vào đại vận (năm xem "&DAI_VAN!$C$23&")","Vận "&DAI_VAN!$C$25&": "&INDEX(DAI_VAN!$C$35:$L$35,DAI_VAN!$
    #   C$25)&" ("&INDEX(DAI_VAN!$C$36:$L$36,DAI_VAN!$C$25)&"–"&(INDEX(DAI_VAN!$C$36:$L$36,DAI_VAN!$C$25)+9)&" tuổi) – "&INDEX(DAI_VAN!$C$121:$L$121,DAI_VAN!$
    #   C$25)&": "&INDEX(DAI_VAN!$C$122:$L$122,DAI_VAN!$C$25)&". Can vận: "&INDEX(DAI_VAN!$C$123:$L$123,DAI_VAN!$C$25)&" Chi vận: "&INDEX(DAI_VAN!$C$124:$L$12
    #   4,DAI_VAN!$C$25)&" Dấu hiệu: "&INDEX(DAI_VAN!$C$92:$L$92,DAI_VAN!$C$25)))
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm xem ',E.v(14,23,3)),' (tuổi '),E.v(14,24,3)),') vượt quá 10 đại vận đã lập (đến '),X.add(E.v(14,36,12),9.0)),' tuổi) – chọn năm xem nhỏ hơn ở NHAP_LIEU C31.') if X.truthy(X.gt(E.v(14,25,3),10.0)) else (X.cat(X.cat('Chưa vào đại vận (năm xem ',E.v(14,23,3)),')') if X.truthy(X.eq(E.v(14,25,3),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận ',E.v(14,25,3)),': '),X.INDEX(E.rg(14,35,3,35,12), E.v(14,25,3))),' ('),X.INDEX(E.rg(14,36,3,36,12), E.v(14,25,3))),'–'),X.add(X.INDEX(E.rg(14,36,3,36,12), E.v(14,25,3)),9.0)),' tuổi) – '),X.INDEX(E.rg(14,121,3,121,12), E.v(14,25,3))),': '),X.INDEX(E.rg(14,122,3,122,12), E.v(14,25,3))),'. Can vận: '),X.INDEX(E.rg(14,123,3,123,12), E.v(14,25,3))),' Chi vận: '),X.INDEX(E.rg(14,124,3,124,12), E.v(14,25,3))),' Dấu hiệu: '),X.INDEX(E.rg(14,92,3,92,12), E.v(14,25,3)))))

def f279(E, r, c):
    # Ô H163
    # = LUAN_GIAI!$J$166
    return E.v(4,166,10)

def f280(E, r, c):
    # Ô B165, B166, B167, B168, B169, B170
    # = LUU_NIEN!A5&" ("&LUU_NIEN!B5&")"
    return X.cat(X.cat(X.cat(E.v(15,r-160,c-1),' ('),E.v(15,r-160,c)),')')

def f281(E, r, c):
    # Ô C165, C166, C167, C168, C169, C170
    # = LUU_NIEN!C5&" / TV "&LUU_NIEN!D5
    return X.cat(X.cat(E.v(15,r-160,c),' / TV '),E.v(15,r-160,c+1))

def f282(E, r, c):
    # Ô D165, D166, D167, D168, D169, D170
    # = LUU_NIEN!F5&" → "&LUU_NIEN!G5
    return X.cat(X.cat(E.v(15,r-160,c+2),' → '),E.v(15,r-160,c+3))

def f283(E, r, c):
    # Ô E165, E166, E167, E168, E169, E170
    # = LUU_NIEN!H5&" → "&LUU_NIEN!I5
    return X.cat(X.cat(E.v(15,r-160,c+3),' → '),E.v(15,r-160,c+4))

def f284(E, r, c):
    # Ô G165, G166, G167, G168, G169, G170
    # = LUU_NIEN!J5
    return E.v(15,r-160,c+3)

def f285(E, r, c):
    # Ô H165
    # = LUAN_GIAI!$J$167
    return E.v(4,167,10)

def f286(E, r, c):
    # Ô H166
    # = LUAN_GIAI!$J$168
    return E.v(4,168,10)

def f287(E, r, c):
    # Ô H167
    # = LUAN_GIAI!$J$169
    return E.v(4,169,10)

def f288(E, r, c):
    # Ô H168
    # = LUAN_GIAI!$J$170
    return E.v(4,170,10)

def f289(E, r, c):
    # Ô H169
    # = LUAN_GIAI!$J$171
    return E.v(4,171,10)

def f290(E, r, c):
    # Ô H170
    # = LUAN_GIAI!$J$172
    return E.v(4,172,10)

def f291(E, r, c):
    # Ô C171
    # = "Xu thế năm theo dụng thần (bảng P1 tr.133, phương pháp khuyến nghị) – năm TỐT: "&IF(TRIM(IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$5,LUU_NIEN!$G$5),3)
    #   ="Tốt",LUU_NIEN!$A$5&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$G$6),3)="Tốt",LUU_NIEN!$A$6&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU
    #   _NIEN!$I$7,LUU_NIEN!$G$7),3)="Tốt",LUU_NIEN!$A$7&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8),3)="Tốt",LUU_NIEN!$A$8&" ","")&IF(L
    #   EFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$9,LUU_NIEN!$G$9),3)="Tốt",LUU_NIEN!$A$9&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$10,LUU_NIEN!$G$10),3)="T
    #   ốt",LUU_NIEN!$A$10&" ",""))="","không có",TRIM(IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$5,LUU_NIEN!$G$5),3)="Tốt",LUU_NIEN!$A$5&" ","")&IF(LEFT(IF(DOI
    #   _CHIEU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$G$6),3)="Tốt",LUU_NIEN!$A$6&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$7,LUU_NIEN!$G$7),3)="Tốt",LUU_NIEN
    #   !$A$7&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8),3)="Tốt",LUU_NIEN!$A$8&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$9,LUU_
    #   NIEN!$G$9),3)="Tốt",LUU_NIEN!$A$9&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$10,LUU_NIEN!$G$10),3)="Tốt",LUU_NIEN!$A$10&" ","")))&"; năm XẤU: "&
    #   IF(TRIM(IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$5,LUU_NIEN!$G$5),3)="Xấu",LUU_NIEN!$A$5&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$
    #   G$6),3)="Xấu",LUU_NIEN!$A$6&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$7,LUU_NIEN!$G$7),3)="Xấu",LUU_NIEN!$A$7&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$3
    #   3=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8),3)="Xấu",LUU_NIEN!$A$8&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$9,LUU_NIEN!$G$9),3)="Xấu",LUU_NIEN!$A$9&" ","
    #   ")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$10,LUU_NIEN!$G$10),3)="Xấu",LUU_NIEN!$A$10&" ",""))="","không có",TRIM(IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NI
    #   EN!$I$5,LUU_NIEN!$G$5),3)="Xấu",LUU_NIEN!$A$5&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$6,LUU_NIEN!$G$6),3)="Xấu",LUU_NIEN!$A$6&" ","")&IF(LEFT
    #   (IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$7,LUU_NIEN!$G$7),3)="Xấu",LUU_NIEN!$A$7&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$8,LUU_NIEN!$G$8),3)="Xấu",L
    #   UU_NIEN!$A$8&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I$9,LUU_NIEN!$G$9),3)="Xấu",LUU_NIEN!$A$9&" ","")&IF(LEFT(IF(DOI_CHIEU!$C$33=2,LUU_NIEN!$I
    #   $10,LUU_NIEN!$G$10),3)="Xấu",LUU_NIEN!$A$10&" ","")))&"; các năm còn lại bình thường. Dấu hiệu bệnh tật – tai họa xét riêng theo 15 loại năm hạn (một 
    #   năm có thể xu thế bình thường mà vẫn có dấu hiệu hạn): "&SUBSTITUTE(SUC_KHOE!$C$38,CHAR(10),"; ")&" (cách hóa giải: SUC_KHOE mục D–E)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Xu thế năm theo dụng thần (bảng P1 tr.133, phương pháp khuyến nghị) – năm TỐT: ',('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(15,5,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7)), 3.0),'Tốt')) else ''),(X.cat(E.v(15,6,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,7,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,8,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,9,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,10,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7)), 3.0),'Tốt')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(15,5,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7)), 3.0),'Tốt')) else ''),(X.cat(E.v(15,6,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,7,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,8,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,9,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7)), 3.0),'Tốt')) else '')),(X.cat(E.v(15,10,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7)), 3.0),'Tốt')) else ''))))),'; năm XẤU: '),('không có' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(15,5,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7)), 3.0),'Xấu')) else ''),(X.cat(E.v(15,6,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,7,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,8,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,9,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,10,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7)), 3.0),'Xấu')) else ''))),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(E.v(15,5,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,5,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,5,7)), 3.0),'Xấu')) else ''),(X.cat(E.v(15,6,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,6,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,6,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,7,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,7,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,7,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,8,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,8,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,8,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,9,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,9,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,9,7)), 3.0),'Xấu')) else '')),(X.cat(E.v(15,10,1),' ') if X.truthy(X.eq(X.LEFT((E.v(15,10,9) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(15,10,7)), 3.0),'Xấu')) else ''))))),'; các năm còn lại bình thường. Dấu hiệu bệnh tật – tai họa xét riêng theo 15 loại năm hạn (một năm có thể xu thế bình thường mà vẫn có dấu hiệu hạn): '),X.SUBSTITUTE(E.v(22,38,3), X.CHAR(10.0), '; ')),' (cách hóa giải: SUC_KHOE mục D–E).')

def f292(E, r, c):
    # Ô H171
    # = LUAN_GIAI!$J$130
    return E.v(4,130,10)

def f293(E, r, c):
    # Ô C172
    # = "■ TRỤ NĂM (đến 15 tuổi) – cung tổ tiên, cha mẹ; thời niên thiếu (P1 tr.149): "&IF((J172+K172)>0,"THUẬN",IF((J172+K172)<0,"KÉM","LẪN LỘN"))&CHAR(10)&"
    #     Vì sao: "&T172&"; "&U172&"."&CHAR(10)&"  Cụ thể: "&IF((J172+K172)>0,"tuổi thơ và gia cảnh nhìn chung thuận lợi",IF((J172+K172)<0,"tuổi thơ và gia cả
    #   nh nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện","tuổi thơ và gia cảnh có mặt tốt, mặt xấu đan xen"))&"."&IF((J172+K172)<0," Sách: ""trụ 
    #   năm là kị thần thì ông cha nghèo đói"" (P2 tr.451).","")&IF(OR(AND(MOD(TRUNG_GIAN!$C$7-TRUNG_GIAN!$E$7+5,5)=2,TRUNG_GIAN!$C$7=IF(DOI_CHIEU!$C$33=2,DUN
    #   G_HY_KY!$D$111,DUNG_HY_KY!$C$111)),AND(MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=2,IF(PP2_DINH_LUONG!$
    #   C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)))," Sách: trụ năm có tài tinh mà là kị thần → 
    #   ""không được hưởng âm đức, tài sản để lại"" (P1 tr.151).","")&" Đối chiếu: LUC_THAN chấm ông bà – tổ nghiệp là “"&IFERROR(LEFT(LUC_THAN!$D$30,FIND(":"
    #   ,LUC_THAN!$D$30)-1),"")&"” vì cách chấm P1 tr.159 xét cả can năm ở 12 cung và cả câu sách thuận lẫn bất lợi (LUC_THAN dòng 30, cột G và các dòng TN); 
    #   dòng này chỉ xét hai chữ của trụ năm với nhật chủ, nên hai kết quả có thể khác mức."&" Xem thêm LUC_THAN dòng 30 (ông bà – tổ nghiệp), 31-32 (cha, mẹ)
    #   ."&CHAR(10)&"■ TRỤ THÁNG (16–31 tuổi) – cung anh chị em; giai đoạn học hành, lập nghiệp (P1 tr.149): "&IF((L172+M172)>0,"THUẬN",IF((L172+M172)<0,"KÉM"
    #   ,"LẪN LỘN"))&CHAR(10)&"  Vì sao: "&V172&"; "&W172&"."&CHAR(10)&"  Cụ thể: "&IF((L172+M172)>0,"giai đoạn 16–31 tuổi và quan hệ anh chị em nhìn chung th
    #   uận lợi",IF((L172+M172)<0,"giai đoạn 16–31 tuổi và quan hệ anh chị em nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện","giai đoạn 16–31 tuổi
    #    và quan hệ anh chị em có mặt tốt, mặt xấu đan xen"))&"."&" Xem thêm LUC_THAN dòng 33 (anh chị em)."&CHAR(10)&"■ TRỤ NGÀY (31–47 tuổi) – bản thân và c
    #   ung "&IF(NHAP_LIEU!$C$22="Nam","vợ","chồng")&" (chi ngày); thời trung niên (P1 tr.149): "&IF((N172+O172)>0,"THUẬN",IF((N172+O172)<0,"KÉM","LẪN LỘN"))&
    #   CHAR(10)&"  Vì sao: "&"can ngày là chính mình; chi "&Y172&"."&CHAR(10)&"  Cụ thể: "&IF((N172+O172)>0,"giai đoạn 31–47 tuổi và đời sống vợ chồng nhìn c
    #   hung thuận lợi",IF((N172+O172)<0,"giai đoạn 31–47 tuổi và đời sống vợ chồng nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện","giai đoạn 31–4
    #   7 tuổi và đời sống vợ chồng có mặt tốt, mặt xấu đan xen"))&"."&" Xem thêm HON_NHAN (vợ/chồng)."&CHAR(10)&"■ TRỤ GIỜ (47–65 tuổi) – cung con cái; tuổi 
    #   già (P1 tr.149): "&IF((P172+Q172)>0,"THUẬN",IF((P172+Q172)<0,"KÉM","LẪN LỘN"))&CHAR(10)&"  Vì sao: "&Z172&"; "&AA172&"."&CHAR(10)&"  Cụ thể: "&IF((P17
    #   2+Q172)>0,"giai đoạn 47–65 tuổi và đường con cái nhìn chung thuận lợi",IF((P172+Q172)<0,"giai đoạn 47–65 tuổi và đường con cái nhìn chung không thuận 
    #   – các mặt xấu nêu trên dễ thể hiện","giai đoạn 47–65 tuổi và đường con cái có mặt tốt, mặt xấu đan xen"))&"."&" Xem thêm LUC_THAN dòng 34 (con cái)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('■ TRỤ NĂM (đến 15 tuổi) – cung tổ tiên, cha mẹ; thời niên thiếu (P1 tr.149): ',('THUẬN' if X.truthy(X.gt(X.add(E.v(3,r,c+7),E.v(3,r,c+8)),0.0)) else ('KÉM' if X.truthy(X.lt(X.add(E.v(3,r,c+7),E.v(3,r,c+8)),0.0)) else 'LẪN LỘN'))),X.CHAR(10.0)),'  Vì sao: '),E.v(3,r,c+17)),'; '),E.v(3,r,c+18)),'.'),X.CHAR(10.0)),'  Cụ thể: '),('tuổi thơ và gia cảnh nhìn chung thuận lợi' if X.truthy(X.gt(X.add(E.v(3,r,c+7),E.v(3,r,c+8)),0.0)) else ('tuổi thơ và gia cảnh nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện' if X.truthy(X.lt(X.add(E.v(3,r,c+7),E.v(3,r,c+8)),0.0)) else 'tuổi thơ và gia cảnh có mặt tốt, mặt xấu đan xen'))),'.'),(' Sách: "trụ năm là kị thần thì ông cha nghèo đói" (P2 tr.451).' if X.truthy(X.lt(X.add(E.v(3,r,c+7),E.v(3,r,c+8)),0.0)) else '')),(' Sách: trụ năm có tài tinh mà là kị thần → "không được hưởng âm đức, tài sản để lại" (P1 tr.151).' if X.truthy(X.OR(X.AND(X.eq(X.MOD(X.add(X.sub(E.v(23,7,3),E.v(23,7,5)),5.0), 5.0),2.0), X.eq(E.v(23,7,3),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))), X.AND(X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),2.0), X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))))) else '')),' Đối chiếu: LUC_THAN chấm ông bà – tổ nghiệp là “'),X.IFERROR(lambda: X.LEFT(E.v(16,30,4), X.sub(X.FIND(':', E.v(16,30,4)),1.0)), lambda: '')),'” vì cách chấm P1 tr.159 xét cả can năm ở 12 cung và cả câu sách thuận lẫn bất lợi (LUC_THAN dòng 30, cột G và các dòng TN); dòng này chỉ xét hai chữ của trụ năm với nhật chủ, nên hai kết quả có thể khác mức.'),' Xem thêm LUC_THAN dòng 30 (ông bà – tổ nghiệp), 31-32 (cha, mẹ).'),X.CHAR(10.0)),'■ TRỤ THÁNG (16–31 tuổi) – cung anh chị em; giai đoạn học hành, lập nghiệp (P1 tr.149): '),('THUẬN' if X.truthy(X.gt(X.add(E.v(3,r,c+9),E.v(3,r,c+10)),0.0)) else ('KÉM' if X.truthy(X.lt(X.add(E.v(3,r,c+9),E.v(3,r,c+10)),0.0)) else 'LẪN LỘN'))),X.CHAR(10.0)),'  Vì sao: '),E.v(3,r,c+19)),'; '),E.v(3,r,c+20)),'.'),X.CHAR(10.0)),'  Cụ thể: '),('giai đoạn 16–31 tuổi và quan hệ anh chị em nhìn chung thuận lợi' if X.truthy(X.gt(X.add(E.v(3,r,c+9),E.v(3,r,c+10)),0.0)) else ('giai đoạn 16–31 tuổi và quan hệ anh chị em nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện' if X.truthy(X.lt(X.add(E.v(3,r,c+9),E.v(3,r,c+10)),0.0)) else 'giai đoạn 16–31 tuổi và quan hệ anh chị em có mặt tốt, mặt xấu đan xen'))),'.'),' Xem thêm LUC_THAN dòng 33 (anh chị em).'),X.CHAR(10.0)),'■ TRỤ NGÀY (31–47 tuổi) – bản thân và cung '),('vợ' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else 'chồng')),' (chi ngày); thời trung niên (P1 tr.149): '),('THUẬN' if X.truthy(X.gt(X.add(E.v(3,r,c+11),E.v(3,r,c+12)),0.0)) else ('KÉM' if X.truthy(X.lt(X.add(E.v(3,r,c+11),E.v(3,r,c+12)),0.0)) else 'LẪN LỘN'))),X.CHAR(10.0)),'  Vì sao: '),'can ngày là chính mình; chi '),E.v(3,r,c+22)),'.'),X.CHAR(10.0)),'  Cụ thể: '),('giai đoạn 31–47 tuổi và đời sống vợ chồng nhìn chung thuận lợi' if X.truthy(X.gt(X.add(E.v(3,r,c+11),E.v(3,r,c+12)),0.0)) else ('giai đoạn 31–47 tuổi và đời sống vợ chồng nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện' if X.truthy(X.lt(X.add(E.v(3,r,c+11),E.v(3,r,c+12)),0.0)) else 'giai đoạn 31–47 tuổi và đời sống vợ chồng có mặt tốt, mặt xấu đan xen'))),'.'),' Xem thêm HON_NHAN (vợ/chồng).'),X.CHAR(10.0)),'■ TRỤ GIỜ (47–65 tuổi) – cung con cái; tuổi già (P1 tr.149): '),('THUẬN' if X.truthy(X.gt(X.add(E.v(3,r,c+13),E.v(3,r,c+14)),0.0)) else ('KÉM' if X.truthy(X.lt(X.add(E.v(3,r,c+13),E.v(3,r,c+14)),0.0)) else 'LẪN LỘN'))),X.CHAR(10.0)),'  Vì sao: '),E.v(3,r,c+23)),'; '),E.v(3,r,c+24)),'.'),X.CHAR(10.0)),'  Cụ thể: '),('giai đoạn 47–65 tuổi và đường con cái nhìn chung thuận lợi' if X.truthy(X.gt(X.add(E.v(3,r,c+13),E.v(3,r,c+14)),0.0)) else ('giai đoạn 47–65 tuổi và đường con cái nhìn chung không thuận – các mặt xấu nêu trên dễ thể hiện' if X.truthy(X.lt(X.add(E.v(3,r,c+13),E.v(3,r,c+14)),0.0)) else 'giai đoạn 47–65 tuổi và đường con cái có mặt tốt, mặt xấu đan xen'))),'.'),' Xem thêm LUC_THAN dòng 34 (con cái).')

def f294(E, r, c):
    # Ô H172
    # = LUAN_GIAI!$J$154
    return E.v(4,154,10)

def f295(E, r, c):
    # Ô J172
    # = IF(OR(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$C$5
    #   1=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$
    #   52,TRUNG_GIAN!$C$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(MOD(
    #   IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$
    #   C$7)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LU
    #   ONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f296(E, r, c):
    # Ô K172
    # = IF(OR(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$C$
    #   24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!
    #   $C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(M
    #   OD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_G
    #   IAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_D
    #   INH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f297(E, r, c):
    # Ô L172
    # = IF(OR(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$D$5
    #   1=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$
    #   52,TRUNG_GIAN!$D$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(MOD(
    #   IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$
    #   D$7)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LU
    #   ONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f298(E, r, c):
    # Ô M172
    # = IF(OR(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$D$
    #   24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!
    #   $D$24,TRUNG_GIAN!$D$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(M
    #   OD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_G
    #   IAN!$D$9)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_D
    #   INH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f299(E, r, c):
    # Ô N172
    # = 0
    return 0.0

def f300(E, r, c):
    # Ô O172
    # = IF(OR(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$E$
    #   24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!
    #   $E$24,TRUNG_GIAN!$E$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(M
    #   OD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_G
    #   IAN!$E$9)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_D
    #   INH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f301(E, r, c):
    # Ô P172
    # = IF(OR(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$F$5
    #   1=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$
    #   52,TRUNG_GIAN!$F$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(MOD(
    #   IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$
    #   F$7)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LU
    #   ONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f302(E, r, c):
    # Ô Q172
    # = IF(OR(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$F$
    #   24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),1,IF(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!
    #   $F$24,TRUNG_GIAN!$F$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),-1,IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0,IF(OR(M
    #   OD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_G
    #   IAN!$F$9)-TRUNG_GIAN!$E$7+5,5)=4),1,-1),IF(OR(MOD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_D
    #   INH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)=4),-1,1))))
    return (1.0 if X.truthy(X.OR(X.eq((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.eq((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else (X.neg(1.0) if X.truthy(X.eq((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ((1.0 if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),4.0))) else X.neg(1.0)) if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else (X.neg(1.0) if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),4.0))) else 1.0))))

def f303(E, r, c):
    # Ô T172
    # = TRUNG_GIAN!$C$10&IF(PP2_DINH_LUONG!$C$51=1," (đã hợp hóa)","")&" = "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_
    #   GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)=IF(DOI_CHIEU!$
    #   C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)=IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUN
    #   G_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$C$51
    #   =1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh trợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(J172>0,"tốt
    #   : "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thô
    #   ng minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nếp, được người trên nâng đỡ, có danh phận","được học hành, được mẹ v
    #   à người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh
    #    giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó
    #   , dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)<>IF(DOI_CHI
    #   EU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$11
    #   0,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$C$51=1,PP2_DINH_LUONG!$C$52,TRUNG_GIAN!$C$7)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF
    #   (DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, khôn
    #   g cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(23,10,3),(' (đã hợp hóa)' if X.truthy(X.eq(E.v(10,51,3),1.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,52,3) if X.truthy(X.eq(E.v(10,51,3),1.0)) else E.v(23,7,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f304(E, r, c):
    # Ô U172
    # = INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&IF(PP2_DINH_LUONG!$C$24>=0," (đã hợp hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$C$24+1)&")","")&" = "&CHOOSE(MOD(IF(PP2_DIN
    #   H_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG
    #   !$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$C$24>=0,PP
    #   2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$
    #   C$24,TRUNG_GIAN!$C$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TR
    #   UNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh t
    #   rợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(K172>0,"tốt: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)-TRUN
    #   G_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thông minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nế
    #   p, được người trên nâng đỡ, có danh phận","được học hành, được mẹ và người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG
    #   !$C$24,TRUNG_GIAN!$C$9)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc th
    #   iếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó, dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_
    #   DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!$C$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$C$24>=0,PP2_DI
    #   NH_LUONG!$C$24,TRUNG_GIAN!$C$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$C$24>=0,PP2_DINH_LUONG!$C$24,TRUNG_GIAN!
    #   $C$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh tr
    #   ợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0)),(X.cat(X.cat(' (đã hợp hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,3),1.0))),')') if X.truthy(X.ge(E.v(10,24,3),0.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,24,3) if X.truthy(X.ge(E.v(10,24,3),0.0)) else E.v(23,9,3)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f305(E, r, c):
    # Ô V172
    # = TRUNG_GIAN!$D$10&IF(PP2_DINH_LUONG!$D$51=1," (đã hợp hóa)","")&" = "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_
    #   GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)=IF(DOI_CHIEU!$
    #   C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)=IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUN
    #   G_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$D$51
    #   =1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh trợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(L172>0,"tốt
    #   : "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thô
    #   ng minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nếp, được người trên nâng đỡ, có danh phận","được học hành, được mẹ v
    #   à người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh
    #    giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó
    #   , dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)<>IF(DOI_CHI
    #   EU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$11
    #   0,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$D$51=1,PP2_DINH_LUONG!$D$52,TRUNG_GIAN!$D$7)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF
    #   (DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, khôn
    #   g cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(23,10,4),(' (đã hợp hóa)' if X.truthy(X.eq(E.v(10,51,4),1.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,52,4) if X.truthy(X.eq(E.v(10,51,4),1.0)) else E.v(23,7,4)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f306(E, r, c):
    # Ô W172
    # = INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&IF(PP2_DINH_LUONG!$D$24>=0," (đã hợp hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$D$24+1)&")","")&" = "&CHOOSE(MOD(IF(PP2_DIN
    #   H_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG
    #   !$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$D$24>=0,PP
    #   2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$
    #   D$24,TRUNG_GIAN!$D$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TR
    #   UNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh t
    #   rợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(M172>0,"tốt: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)-TRUN
    #   G_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thông minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nế
    #   p, được người trên nâng đỡ, có danh phận","được học hành, được mẹ và người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG
    #   !$D$24,TRUNG_GIAN!$D$9)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc th
    #   iếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó, dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_
    #   DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!$D$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$D$24>=0,PP2_DI
    #   NH_LUONG!$D$24,TRUNG_GIAN!$D$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$D$24>=0,PP2_DINH_LUONG!$D$24,TRUNG_GIAN!
    #   $D$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh tr
    #   ợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0)),(X.cat(X.cat(' (đã hợp hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,4),1.0))),')') if X.truthy(X.ge(E.v(10,24,4),0.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,24,4) if X.truthy(X.ge(E.v(10,24,4),0.0)) else E.v(23,9,4)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f307(E, r, c):
    # Ô Y172
    # = INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&IF(PP2_DINH_LUONG!$E$24>=0," (đã hợp hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$E$24+1)&")","")&" = "&CHOOSE(MOD(IF(PP2_DIN
    #   H_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG
    #   !$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$E$24>=0,PP
    #   2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$
    #   E$24,TRUNG_GIAN!$E$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TR
    #   UNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh t
    #   rợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(O172>0,"tốt: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)-TRUN
    #   G_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thông minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nế
    #   p, được người trên nâng đỡ, có danh phận","được học hành, được mẹ và người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG
    #   !$E$24,TRUNG_GIAN!$E$9)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc th
    #   iếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó, dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_
    #   DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!$E$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$E$24>=0,PP2_DI
    #   NH_LUONG!$E$24,TRUNG_GIAN!$E$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$E$24>=0,PP2_DINH_LUONG!$E$24,TRUNG_GIAN!
    #   $E$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh tr
    #   ợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0)),(X.cat(X.cat(' (đã hợp hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,5),1.0))),')') if X.truthy(X.ge(E.v(10,24,5),0.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,24,5) if X.truthy(X.ge(E.v(10,24,5),0.0)) else E.v(23,9,5)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f308(E, r, c):
    # Ô Z172
    # = TRUNG_GIAN!$F$10&IF(PP2_DINH_LUONG!$F$51=1," (đã hợp hóa)","")&" = "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_
    #   GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)=IF(DOI_CHIEU!$
    #   C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)=IF(DOI_CHIEU!$C$33=2,DUNG_H
    #   Y_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUN
    #   G_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$F$51
    #   =1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh trợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(P172>0,"tốt
    #   : "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thô
    #   ng minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nếp, được người trên nâng đỡ, có danh phận","được học hành, được mẹ v
    #   à người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh
    #    giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó
    #   , dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)<>IF(DOI_CHI
    #   EU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$11
    #   0,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$F$51=1,PP2_DINH_LUONG!$F$52,TRUNG_GIAN!$F$7)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF
    #   (DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, khôn
    #   g cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(23,10,6),(' (đã hợp hóa)' if X.truthy(X.eq(E.v(10,51,6),1.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,52,6) if X.truthy(X.eq(E.v(10,51,6),1.0)) else E.v(23,7,6)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f309(E, r, c):
    # Ô AA172
    # = INDEX(CHI_TEN,TRUNG_GIAN!$F$6+1)&IF(PP2_DINH_LUONG!$F$24>=0," (đã hợp hóa "&INDEX(HANH_TEN,PP2_DINH_LUONG!$F$24+1)&")","")&" = "&CHOOSE(MOD(IF(PP2_DIN
    #   H_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)+1,"tỉ kiếp","thực thương","tài","quan sát","ấn")&", "&IF(IF(PP2_DINH_LUONG
    #   !$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"dụng thần",IF(IF(PP2_DINH_LUONG!$F$24>=0,PP
    #   2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"hỷ thần",IF(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$
    #   F$24,TRUNG_GIAN!$F$9)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),"kỵ thần",IF(OR(MOD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TR
    #   UNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)=0,MOD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)=4),"nhàn thần – sinh t
    #   rợ thân","nhàn thần – khắc tiết hao thân"))))&" → "&IF(Q172>0,"tốt: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)-TRUN
    #   G_GIAN!$E$7+5,5)+1,"được anh em, bạn bè trợ giúp","tài năng, sự thông minh được thể hiện và có kết quả","điều kiện tiền bạc, vật chất thuận","có nề nế
    #   p, được người trên nâng đỡ, có danh phận","được học hành, được mẹ và người trên che chở"),"xấu: "&CHOOSE(MOD(IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG
    #   !$F$24,TRUNG_GIAN!$F$9)-TRUNG_GIAN!$E$7+5,5)+1,"anh em, bạn bè tranh giành, dễ hao tiền","hao sức, lo nghĩ nhiều, tài năng khó thành quả","tiền bạc th
    #   iếu hụt hoặc thành gánh nặng, vất vả vì kinh tế","chịu áp lực, gò bó, dễ bệnh tật, tai nạn","ỷ lại, học hành dở dang, thiếu chủ động"))&IF(AND(IF(PP2_
    #   DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!$F$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),IF(PP2_DINH_LUONG!$F$24>=0,PP2_DI
    #   NH_LUONG!$F$24,TRUNG_GIAN!$F$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),IF(PP2_DINH_LUONG!$F$24>=0,PP2_DINH_LUONG!$F$24,TRUNG_GIAN!
    #   $F$9)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)),IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0," (thân nhược cần sinh tr
    #   ợ, sợ khắc tiết hao – P1 tr.125)"," (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)"),"")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0)),(X.cat(X.cat(' (đã hợp hóa ',X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(10,24,6),1.0))),')') if X.truthy(X.ge(E.v(10,24,6),0.0)) else '')),' = '),X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'tỉ kiếp', lambda: 'thực thương', lambda: 'tài', lambda: 'quan sát', lambda: 'ấn')),', '),('dụng thần' if X.truthy(X.eq((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('hỷ thần' if X.truthy(X.eq((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('kỵ thần' if X.truthy(X.eq((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('nhàn thần – sinh trợ thân' if X.truthy(X.OR(X.eq(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),0.0), X.eq(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),4.0))) else 'nhàn thần – khắc tiết hao thân'))))),' → '),(X.cat('tốt: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'được anh em, bạn bè trợ giúp', lambda: 'tài năng, sự thông minh được thể hiện và có kết quả', lambda: 'điều kiện tiền bạc, vật chất thuận', lambda: 'có nề nếp, được người trên nâng đỡ, có danh phận', lambda: 'được học hành, được mẹ và người trên che chở')) if X.truthy(X.gt(E.v(3,r,c-10),0.0)) else X.cat('xấu: ',X.CHOOSE(X.add(X.MOD(X.add(X.sub((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),E.v(23,7,5)),5.0), 5.0),1.0), lambda: 'anh em, bạn bè tranh giành, dễ hao tiền', lambda: 'hao sức, lo nghĩ nhiều, tài năng khó thành quả', lambda: 'tiền bạc thiếu hụt hoặc thành gánh nặng, vất vả vì kinh tế', lambda: 'chịu áp lực, gò bó, dễ bệnh tật, tai nạn', lambda: 'ỷ lại, học hành dở dang, thiếu chủ động')))),((' (thân nhược cần sinh trợ, sợ khắc tiết hao – P1 tr.125)' if X.truthy(X.eq((E.v(8,7,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,7,3)),0.0)) else ' (thân vượng cần khắc tiết hao, không cần thêm sinh trợ – P1 tr.125)') if X.truthy(X.AND(X.ne((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))), X.ne((E.v(10,24,6) if X.truthy(X.ge(E.v(10,24,6),0.0)) else E.v(23,9,6)),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3))))) else ''))

def f310(E, r, c):
    # Ô B173
    # = MENH_CUC!B20
    return E.v(18,r-153,c)

def f311(E, r, c):
    # Ô C173
    # = MENH_CUC!D20&CHAR(10)&MENH_CUC!G20&CHAR(10)&MENH_CUC!K20&CHAR(10)&"► Ý nghĩa cụ thể: "&IF(MENH_CUC!$U$20="HUNG","đời nhiều sóng gió, trắc trở; kỵ thần
    #    có lực khắc dụng thần nên dễ gặp thất bại, bệnh tật, tai họa khi vào vận xấu (P1 tr.170-171). “Nghiêng” nghĩa là vẫn có cứu ở vận tốt",IF(MENH_CUC!$U
    #   $20="CÁT","“cát tuy không nhất định sẽ giàu sang, nhưng cả đời ổn định, ít vất vả phong ba” (P1 tr.170)","cân bằng: hai chiều ngang nhau, tốt xấu tùy 
    #   vận"))&". "&IF(AND($AB$87="",$AC$87=""),"Khi nào thể hiện rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều
    #    ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu ("&IF($AD$87="
    #   ","không có",$AD$87)&"); phần xấu dễ nổi hơn trong bại vận ("&IF($AE$87="","không có",$AE$87)&"). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.","Khi n
    #   ào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT ("&IF($AB$87="","không có",$AB$87)&"); phần xấu nổi lên trong đại vận XẤU ("&IF($AC$87="","không có
    #   ",$AC$87)&").")
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(18,r-153,c+1),X.CHAR(10.0)),E.v(18,r-153,c+4)),X.CHAR(10.0)),E.v(18,r-153,c+8)),X.CHAR(10.0)),'► Ý nghĩa cụ thể: '),('đời nhiều sóng gió, trắc trở; kỵ thần có lực khắc dụng thần nên dễ gặp thất bại, bệnh tật, tai họa khi vào vận xấu (P1 tr.170-171). “Nghiêng” nghĩa là vẫn có cứu ở vận tốt' if X.truthy(X.eq(E.v(18,20,21),'HUNG')) else ('“cát tuy không nhất định sẽ giàu sang, nhưng cả đời ổn định, ít vất vả phong ba” (P1 tr.170)' if X.truthy(X.eq(E.v(18,20,21),'CÁT')) else 'cân bằng: hai chiều ngang nhau, tốt xấu tùy vận'))),'. '),(X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: sau điều chỉnh (vận tốt bị mệnh cục khắc/hợp thì giảm – DAI_VAN dòng 55) các đại vận đều ở mức Thường, nên không vận nào tốt hẳn hay xấu hẳn. Theo phân loại gốc P1 tr.125: phần tốt dễ nổi hơn trong vận sinh trợ đúng nhu cầu (',('không có' if X.truthy(X.eq(E.v(3,87,30),'')) else E.v(3,87,30))),'); phần xấu dễ nổi hơn trong bại vận ('),('không có' if X.truthy(X.eq(E.v(3,87,31),'')) else E.v(3,87,31))),'). Mức độ cụ thể xem từng năm ở sheet LUU_NIEN.') if X.truthy(X.AND(X.eq(E.v(3,87,28),''), X.eq(E.v(3,87,29),''))) else X.cat(X.cat(X.cat(X.cat('Khi nào thể hiện rõ: phần tốt nổi lên trong đại vận TỐT (',('không có' if X.truthy(X.eq(E.v(3,87,28),'')) else E.v(3,87,28))),'); phần xấu nổi lên trong đại vận XẤU ('),('không có' if X.truthy(X.eq(E.v(3,87,29),'')) else E.v(3,87,29))),').')))

def f312(E, r, c):
    # Ô H173
    # = LUAN_GIAI!$J$118
    return E.v(4,118,10)

def f313(E, r, c):
    # Ô C174
    # = TRIM(IF(DAI_VAN!$C$25<1,"",IF(INDEX(DAI_VAN!$C$93:$L$93,MIN(10,DAI_VAN!$C$25))="","","Đại vận đang hành: "&INDEX(DAI_VAN!$C$93:$L$93,MIN(10,DAI_VAN!$C
    #   $25))&" "))&IF(TRIM(IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+0,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+0)&": "&INDEX(LUU_NI
    #   EN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+0,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+1,LUU_NIEN!$A$
    #   30:$A$129,0))="","",(DAI_VAN!$C$23+1)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+1,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU
    #   _NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+2,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+2)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+2,
    #   LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+3,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+3)&":
    #    "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+3,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23
    #   +4,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+4)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+4,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERR
    #   OR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+5,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+5)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(D
    #   AI_VAN!$C$23+5,LUU_NIEN!$A$30:$A$129,0))&" "),""))="","","Năm xem và 5 năm tới: "&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+0,LUU_N
    #   IEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+0)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+0,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(IN
    #   DEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+1,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+1)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$
    #   C$23+1,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+2,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$2
    #   3+2)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+2,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VA
    #   N!$C$23+3,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+3)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+3,LUU_NIEN!$A$30:$A$129,0))&" "),""
    #   )&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+4,LUU_NIEN!$A$30:$A$129,0))="","",(DAI_VAN!$C$23+4)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,
    #   MATCH(DAI_VAN!$C$23+4,LUU_NIEN!$A$30:$A$129,0))&" "),"")&IFERROR(IF(INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+5,LUU_NIEN!$A$30:$A$129,0))="","
    #   ",(DAI_VAN!$C$23+5)&": "&INDEX(LUU_NIEN!$CU$30:$CU$129,MATCH(DAI_VAN!$C$23+5,LUU_NIEN!$A$30:$A$129,0))&" "),"")))
    return X.TRIM(X.cat(('' if X.truthy(X.lt(E.v(14,25,3),1.0)) else ('' if X.truthy(X.eq(X.INDEX(E.rg(14,93,3,93,12), X.MIN(10.0, E.rg(14,25,3,25,3))),'')) else X.cat(X.cat('Đại vận đang hành: ',X.INDEX(E.rg(14,93,3,93,12), X.MIN(10.0, E.rg(14,25,3,25,3)))),' '))),('' if X.truthy(X.eq(X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),0.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: ''),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),1.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),2.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),3.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),4.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),5.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: ''))),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Năm xem và 5 năm tới: ',X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),0.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),1.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),2.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),3.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),4.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')),X.IFERROR(lambda: ('' if X.truthy(X.eq(X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)),'')) else X.cat(X.cat(X.cat(X.add(E.v(14,23,3),5.0),': '),X.INDEX(E.rg(15,30,99,129,99), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0))),' ')), lambda: '')))))

def f314(E, r, c):
    # Ô C175
    # = TRIM(IF(DAI_VAN!$C$25<1,"","Vận đang hành: "&INDEX(DAI_VAN!$C$35:$L$35,MIN(10,DAI_VAN!$C$25))&" – "&INDEX(DAI_VAN!$C$102:$L$102,MIN(10,DAI_VAN!$C$25))
    #   &". "&INDEX(DAI_VAN!$C$138:$L$138,MIN(10,DAI_VAN!$C$25))&" ")&"Các vận: "&IF(DAI_VAN!$C$35="","",DAI_VAN!$C$35&" "&DAI_VAN!$C$102&"; ")&IF(DAI_VAN!$D$
    #   35="","",DAI_VAN!$D$35&" "&DAI_VAN!$D$102&"; ")&IF(DAI_VAN!$E$35="","",DAI_VAN!$E$35&" "&DAI_VAN!$E$102&"; ")&IF(DAI_VAN!$F$35="","",DAI_VAN!$F$35&" "
    #   &DAI_VAN!$F$102&"; ")&IF(DAI_VAN!$G$35="","",DAI_VAN!$G$35&" "&DAI_VAN!$G$102&"; ")&IF(DAI_VAN!$H$35="","",DAI_VAN!$H$35&" "&DAI_VAN!$H$102&"; ")&IF(D
    #   AI_VAN!$I$35="","",DAI_VAN!$I$35&" "&DAI_VAN!$I$102&"; ")&IF(DAI_VAN!$J$35="","",DAI_VAN!$J$35&" "&DAI_VAN!$J$102&"; ")&IF(DAI_VAN!$K$35="","",DAI_VAN
    #   !$K$35&" "&DAI_VAN!$K$102&"; ")&IF(DAI_VAN!$L$35="","",DAI_VAN!$L$35&" "&DAI_VAN!$L$102&"; "))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('' if X.truthy(X.lt(E.v(14,25,3),1.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Vận đang hành: ',X.INDEX(E.rg(14,35,3,35,12), X.MIN(10.0, E.rg(14,25,3,25,3)))),' – '),X.INDEX(E.rg(14,102,3,102,12), X.MIN(10.0, E.rg(14,25,3,25,3)))),'. '),X.INDEX(E.rg(14,138,3,138,12), X.MIN(10.0, E.rg(14,25,3,25,3)))),' ')),'Các vận: '),('' if X.truthy(X.eq(E.v(14,35,3),'')) else X.cat(X.cat(X.cat(E.v(14,35,3),' '),E.v(14,102,3)),'; '))),('' if X.truthy(X.eq(E.v(14,35,4),'')) else X.cat(X.cat(X.cat(E.v(14,35,4),' '),E.v(14,102,4)),'; '))),('' if X.truthy(X.eq(E.v(14,35,5),'')) else X.cat(X.cat(X.cat(E.v(14,35,5),' '),E.v(14,102,5)),'; '))),('' if X.truthy(X.eq(E.v(14,35,6),'')) else X.cat(X.cat(X.cat(E.v(14,35,6),' '),E.v(14,102,6)),'; '))),('' if X.truthy(X.eq(E.v(14,35,7),'')) else X.cat(X.cat(X.cat(E.v(14,35,7),' '),E.v(14,102,7)),'; '))),('' if X.truthy(X.eq(E.v(14,35,8),'')) else X.cat(X.cat(X.cat(E.v(14,35,8),' '),E.v(14,102,8)),'; '))),('' if X.truthy(X.eq(E.v(14,35,9),'')) else X.cat(X.cat(X.cat(E.v(14,35,9),' '),E.v(14,102,9)),'; '))),('' if X.truthy(X.eq(E.v(14,35,10),'')) else X.cat(X.cat(X.cat(E.v(14,35,10),' '),E.v(14,102,10)),'; '))),('' if X.truthy(X.eq(E.v(14,35,11),'')) else X.cat(X.cat(X.cat(E.v(14,35,11),' '),E.v(14,102,11)),'; '))),('' if X.truthy(X.eq(E.v(14,35,12),'')) else X.cat(X.cat(X.cat(E.v(14,35,12),' '),E.v(14,102,12)),'; '))))

def f315(E, r, c):
    # Ô C176
    # = TRIM(IFERROR((DAI_VAN!$C$23+0)&": "&MID(INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$C$23+0,LUU_NIEN!$A$30:$A$129,0)),12,FIND(" (",INDEX(LUU_NIEN!$CV$
    #   30:$CV$129,MATCH(DAI_VAN!$C$23+0,LUU_NIEN!$A$30:$A$129,0))&" (")-12)&"; ","")&IFERROR((DAI_VAN!$C$23+1)&": "&MID(INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(D
    #   AI_VAN!$C$23+1,LUU_NIEN!$A$30:$A$129,0)),12,FIND(" (",INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$C$23+1,LUU_NIEN!$A$30:$A$129,0))&" (")-12)&"; ","")
    #   &IFERROR((DAI_VAN!$C$23+2)&": "&MID(INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$C$23+2,LUU_NIEN!$A$30:$A$129,0)),12,FIND(" (",INDEX(LUU_NIEN!$CV$30:$
    #   CV$129,MATCH(DAI_VAN!$C$23+2,LUU_NIEN!$A$30:$A$129,0))&" (")-12)&"; ","")&IFERROR((DAI_VAN!$C$23+3)&": "&MID(INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_V
    #   AN!$C$23+3,LUU_NIEN!$A$30:$A$129,0)),12,FIND(" (",INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$C$23+3,LUU_NIEN!$A$30:$A$129,0))&" (")-12)&"; ","")&IFE
    #   RROR((DAI_VAN!$C$23+4)&": "&MID(INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$C$23+4,LUU_NIEN!$A$30:$A$129,0)),12,FIND(" (",INDEX(LUU_NIEN!$CV$30:$CV$1
    #   29,MATCH(DAI_VAN!$C$23+4,LUU_NIEN!$A$30:$A$129,0))&" (")-12)&"; ","")&IFERROR((DAI_VAN!$C$23+5)&": "&MID(INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$
    #   C$23+5,LUU_NIEN!$A$30:$A$129,0)),12,FIND(" (",INDEX(LUU_NIEN!$CV$30:$CV$129,MATCH(DAI_VAN!$C$23+5,LUU_NIEN!$A$30:$A$129,0))&" (")-12)&"; ",""))
    return X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.IFERROR(lambda: X.cat(X.cat(X.cat(X.add(E.v(14,23,3),0.0),': '),X.MID(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)), 12.0, X.sub(X.FIND(' (', X.cat(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),0.0), E.rg(15,30,1,129,1), 0.0)),' (')),12.0))),'; '), lambda: ''),X.IFERROR(lambda: X.cat(X.cat(X.cat(X.add(E.v(14,23,3),1.0),': '),X.MID(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)), 12.0, X.sub(X.FIND(' (', X.cat(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),1.0), E.rg(15,30,1,129,1), 0.0)),' (')),12.0))),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.cat(X.cat(X.add(E.v(14,23,3),2.0),': '),X.MID(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)), 12.0, X.sub(X.FIND(' (', X.cat(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),2.0), E.rg(15,30,1,129,1), 0.0)),' (')),12.0))),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.cat(X.cat(X.add(E.v(14,23,3),3.0),': '),X.MID(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)), 12.0, X.sub(X.FIND(' (', X.cat(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),3.0), E.rg(15,30,1,129,1), 0.0)),' (')),12.0))),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.cat(X.cat(X.add(E.v(14,23,3),4.0),': '),X.MID(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)), 12.0, X.sub(X.FIND(' (', X.cat(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),4.0), E.rg(15,30,1,129,1), 0.0)),' (')),12.0))),'; '), lambda: '')),X.IFERROR(lambda: X.cat(X.cat(X.cat(X.add(E.v(14,23,3),5.0),': '),X.MID(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)), 12.0, X.sub(X.FIND(' (', X.cat(X.INDEX(E.rg(15,30,100,129,100), X.MATCH(X.add(E.v(14,23,3),5.0), E.rg(15,30,1,129,1), 0.0)),' (')),12.0))),'; '), lambda: '')))

def f316(E, r, c):
    # Ô C179
    # = ("NÊN: dùng hành dụng thần "&CHOOSE(DUNG_HY_KY!$C$157+1,"Mộc – màu xanh lục, phương Đông","Hỏa – màu đỏ, phương Nam","Thổ – màu vàng, nâu, ở giữa","Ki
    #   m – màu trắng, phương Tây","Thủy – màu đen, xanh thẫm, phương Bắc")&" và hỷ thần "&IF(DUNG_HY_KY!$C$158<0,"—",CHOOSE(DUNG_HY_KY!$C$158+1,"Mộc – màu xa
    #   nh lục, phương Đông","Hỏa – màu đỏ, phương Nam","Thổ – màu vàng, nâu, ở giữa","Kim – màu trắng, phương Tây","Thủy – màu đen, xanh thẫm, phương Bắc"))&
    #   " (nghề, màu sắc, phương vị nơi ở – làm việc, tên gọi; P2 tr.377). TRÁNH: hành kỵ thần "&IF(DUNG_HY_KY!$C$159<0,"—",CHOOSE(DUNG_HY_KY!$C$159+1,"Mộc – 
    #   màu xanh lục, phương Đông","Hỏa – màu đỏ, phương Nam","Thổ – màu vàng, nâu, ở giữa","Kim – màu trắng, phương Tây","Thủy – màu đen, xanh thẫm, phương B
    #   ắc"))&"."&IF(AND(DUNG_HY_KY!$C$163="BẮT BUỘC",DUNG_HY_KY!$C$162<>DUNG_HY_KY!$C$159),IF(DUNG_HY_KY!$C$162=4," Điều hầu bắt buộc: lợi phương Bắc, màu đe
    #   n."," Điều hầu bắt buộc: lợi phương Nam, màu đỏ."),"")&" Nghề hợp dụng thần: "&TO_CHAT!$C$17&IF(TO_CHAT!$M$16="",""," (sách có lưu ý bất lợi về nghề)"
    #   )&" – đủ ở dòng Thiên hướng nghề.")&IF(AND(DOI_CHIEU!$C$33<>0,DUNG_HY_KY!$C$109<>DUNG_HY_KY!$D$109),CHAR(10)&"⚠ Nếu đọc theo "&IF(DOI_CHIEU!$C$33=2,"P
    #   P1","PP2")&": dụng thần "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$C$106,DUNG_HY_KY!$D$106)&", hỷ thần "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$C$107,DUNG_HY_KY!$D$
    #   107)&", kỵ thần "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$C$108,DUNG_HY_KY!$D$108)&" → lời khuyên NÊN/TRÁNH ở trên sẽ đảo lại. Hai phương pháp ngược nhau: hã
    #   y đối chiếu với các năm đã qua (năm xấu có rơi vào hành kỵ không, năm tốt có rơi vào hành dụng không) để chọn cách đọc đúng.","")&CHAR(10)&"Trọng tâm 
    #   hóa giải: "&LEFT(C181,IFERROR(FIND(CHAR(10),C181)-1,LEN(C181)))
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('NÊN: dùng hành dụng thần ',X.CHOOSE(X.add(E.v(8,157,3),1.0), lambda: 'Mộc – màu xanh lục, phương Đông', lambda: 'Hỏa – màu đỏ, phương Nam', lambda: 'Thổ – màu vàng, nâu, ở giữa', lambda: 'Kim – màu trắng, phương Tây', lambda: 'Thủy – màu đen, xanh thẫm, phương Bắc')),' và hỷ thần '),('—' if X.truthy(X.lt(E.v(8,158,3),0.0)) else X.CHOOSE(X.add(E.v(8,158,3),1.0), lambda: 'Mộc – màu xanh lục, phương Đông', lambda: 'Hỏa – màu đỏ, phương Nam', lambda: 'Thổ – màu vàng, nâu, ở giữa', lambda: 'Kim – màu trắng, phương Tây', lambda: 'Thủy – màu đen, xanh thẫm, phương Bắc'))),' (nghề, màu sắc, phương vị nơi ở – làm việc, tên gọi; P2 tr.377). TRÁNH: hành kỵ thần '),('—' if X.truthy(X.lt(E.v(8,159,3),0.0)) else X.CHOOSE(X.add(E.v(8,159,3),1.0), lambda: 'Mộc – màu xanh lục, phương Đông', lambda: 'Hỏa – màu đỏ, phương Nam', lambda: 'Thổ – màu vàng, nâu, ở giữa', lambda: 'Kim – màu trắng, phương Tây', lambda: 'Thủy – màu đen, xanh thẫm, phương Bắc'))),'.'),((' Điều hầu bắt buộc: lợi phương Bắc, màu đen.' if X.truthy(X.eq(E.v(8,162,3),4.0)) else ' Điều hầu bắt buộc: lợi phương Nam, màu đỏ.') if X.truthy(X.AND(X.eq(E.v(8,163,3),'BẮT BUỘC'), X.ne(E.v(8,162,3),E.v(8,159,3)))) else '')),' Nghề hợp dụng thần: '),E.v(20,17,3)),('' if X.truthy(X.eq(E.v(20,16,13),'')) else ' (sách có lưu ý bất lợi về nghề)')),' – đủ ở dòng Thiên hướng nghề.'),(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.CHAR(10.0),'⚠ Nếu đọc theo '),('PP1' if X.truthy(X.eq(E.v(7,33,3),2.0)) else 'PP2')),': dụng thần '),(E.v(8,106,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,106,4))),', hỷ thần '),(E.v(8,107,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,107,4))),', kỵ thần '),(E.v(8,108,3) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,108,4))),' → lời khuyên NÊN/TRÁNH ở trên sẽ đảo lại. Hai phương pháp ngược nhau: hãy đối chiếu với các năm đã qua (năm xấu có rơi vào hành kỵ không, năm tốt có rơi vào hành dụng không) để chọn cách đọc đúng.') if X.truthy(X.AND(X.ne(E.v(7,33,3),0.0), X.ne(E.v(8,109,3),E.v(8,109,4)))) else '')),X.CHAR(10.0)),'Trọng tâm hóa giải: '),X.LEFT(E.v(3,r+2,c), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(3,r+2,c)),1.0), lambda: X.LEN(E.v(3,r+2,c)))))

def f317(E, r, c):
    # Ô C181
    # = "[Sách] "&KET_LUAN!$C$140&CHAR(10)&KET_LUAN!$C$155
    return X.cat(X.cat(X.cat('[Sách] ',E.v(3,140,3)),X.CHAR(10.0)),E.v(3,155,3))

def f318(E, r, c):
    # Ô C182
    # = SUC_KHOE!$C$162
    return E.v(22,162,3)

def f319(E, r, c):
    # Ô C183
    # = SUC_KHOE!$C$163
    return E.v(22,163,3)

def f320(E, r, c):
    # Ô C184
    # = SUC_KHOE!$C$160
    return E.v(22,160,3)

def f321(E, r, c):
    # Ô C185
    # = SUC_KHOE!$C$164
    return E.v(22,164,3)

def f322(E, r, c):
    # Ô C186
    # = SUC_KHOE!$C$50&CHAR(10)&SUC_KHOE!$C$52&CHAR(10)&"Chọn ngày giải hạn: SUC_KHOE mục E."
    return X.cat(X.cat(X.cat(X.cat(E.v(22,50,3),X.CHAR(10.0)),E.v(22,52,3)),X.CHAR(10.0)),'Chọn ngày giải hạn: SUC_KHOE mục E.')

def f323(E, r, c):
    # Ô C187
    # = "Tháng sinh "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&"; Thủy "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&", H
    #   ỏa "&TEXT(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0),"0%")&". "&IF(AND(OR(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,T
    #   RUNG_GIAN!$D$6=1),IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4),NOT(OR(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TR
    #   UNG_GIAN!$D$6=7),IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4))),IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_
    #   HY_KY!$C$111)=1,"Lá số thiên hàn nhưng Hỏa là kỵ thần: giữ ấm VỪA ĐỦ, không bồi thêm Hỏa; phòng khô ráo, thoáng.","Lá số thiên HÀN: ưu tiên ẤM – KHÔ –
    #    SÁNG: giữ ấm cơ thể, đủ nắng, ăn chín ấm, nhà khô ráo. Cùng Thông Bảo Giám coi sưởi ấm là việc đầu tiên với lá số sinh mùa lạnh hoặc nhiều Thủy."),IF
    #   (AND(OR(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4),
    #   NOT(OR(OR(TRUNG_GIAN!$D$6=11,TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4))
    #   ),IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)=4,"Lá số thiên nhiệt nhưng Thủy là kỵ thần: giữ mát vừa phải, không bồi thêm Thủy.","Lá
    #    số thiên NÓNG – KHÔ: ưu tiên MÁT – NHUẬN: không gian thoáng mát, uống đủ nước, ngủ sớm, hạn chế cay nóng, rượu."),IF(AND(OR(OR(TRUNG_GIAN!$D$6=11,TRU
    #   NG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1),IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(4,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4),OR(OR(TRUNG_GIAN!$D$6=5,TRUNG_GI
    #   AN!$D$6=6,TRUNG_GIAN!$D$6=7),IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(1,5)+1)/PP2_DINH_LUONG!$C$100,0)>=0.4)),"Vừa có yếu tố hàn vừa có yếu tố nhi
    #   ệt: giữ nhịp sinh hoạt điều độ, tránh thái quá cả hai phía.","Lá số không thiên hàn hay nhiệt rõ: giữ nếp sinh hoạt điều độ.")))&" (Điều hầu – Cùng Th
    #   ông Bảo Giám; ngoài sách Thiệu Vĩ Hoa.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tháng sinh ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),'; Thủy '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),', Hỏa '),X.TEXT(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0), '0%')),'. '),(('Lá số thiên hàn nhưng Hỏa là kỵ thần: giữ ấm VỪA ĐỦ, không bồi thêm Hỏa; phòng khô ráo, thoáng.' if X.truthy(X.eq((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0)) else 'Lá số thiên HÀN: ưu tiên ẤM – KHÔ – SÁNG: giữ ấm cơ thể, đủ nắng, ăn chín ấm, nhà khô ráo. Cùng Thông Bảo Giám coi sưởi ấm là việc đầu tiên với lá số sinh mùa lạnh hoặc nhiều Thủy.') if X.truthy(X.AND(X.OR(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4)), X.NOT(X.OR(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0)), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4))))) else (('Lá số thiên nhiệt nhưng Thủy là kỵ thần: giữ mát vừa phải, không bồi thêm Thủy.' if X.truthy(X.eq((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),4.0)) else 'Lá số thiên NÓNG – KHÔ: ưu tiên MÁT – NHUẬN: không gian thoáng mát, uống đủ nước, ngủ sớm, hạn chế cay nóng, rượu.') if X.truthy(X.AND(X.OR(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0)), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4)), X.NOT(X.OR(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4))))) else ('Vừa có yếu tố hàn vừa có yếu tố nhiệt: giữ nhịp sinh hoạt điều độ, tránh thái quá cả hai phía.' if X.truthy(X.AND(X.OR(X.OR(X.eq(E.v(23,6,4),11.0), X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(4.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4)), X.OR(X.OR(X.eq(E.v(23,6,4),5.0), X.eq(E.v(23,6,4),6.0), X.eq(E.v(23,6,4),7.0)), X.ge(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(1.0, 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.4)))) else 'Lá số không thiên hàn hay nhiệt rõ: giữ nếp sinh hoạt điều độ.')))),' (Điều hầu – Cùng Thông Bảo Giám; ngoài sách Thiệu Vĩ Hoa.)')

def f324(E, r, c):
    # Ô C188
    # = IF(AK234&AK235&AK236&AK237&AK238&AK239="","Tứ trụ không có cặp chi xung – không cần thông quan.",AK234&AK235&AK236&AK237&AK238&AK239&"Nguyên tắc: hai 
    #   hành xung khắc thì dùng hành đứng giữa (hành bị bên khắc sinh ra và sinh cho bên bị khắc) để nối; sách nêu ở đề cương phối thang mục 7 (P2 tr.647). Ứn
    #   g dụng đời thường theo nhóm thập thần là kinh nghiệm ngoài sách.")
    return ('Tứ trụ không có cặp chi xung – không cần thông quan.' if X.truthy(X.eq(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+46,c+34),E.v(3,r+47,c+34)),E.v(3,r+48,c+34)),E.v(3,r+49,c+34)),E.v(3,r+50,c+34)),E.v(3,r+51,c+34)),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+46,c+34),E.v(3,r+47,c+34)),E.v(3,r+48,c+34)),E.v(3,r+49,c+34)),E.v(3,r+50,c+34)),E.v(3,r+51,c+34)),'Nguyên tắc: hai hành xung khắc thì dùng hành đứng giữa (hành bị bên khắc sinh ra và sinh cho bên bị khắc) để nối; sách nêu ở đề cương phối thang mục 7 (P2 tr.647). Ứng dụng đời thường theo nhóm thập thần là kinh nghiệm ngoài sách.'))

def f325(E, r, c):
    # Ô C189
    # = "Dụng thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),5)+1)&": "&INDEX({"giờ Dần – Mão (3–7h)","giờ Tị – Ngọ (9–13
    #   h)","giờ Thìn, Mùi, Tuất, Sửu (7–9h, 13–15h, 19–21h)","giờ Thân – Dậu (15–19h)","giờ Hợi – Tí (21–1h, nghỉ ngơi)"},IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$
    #   109,DUNG_HY_KY!$C$109)+1)&"; vật dụng "&INDEX({"cây xanh, đồ gỗ, màu xanh lá","đèn ánh sáng ấm, nắng, đồ màu đỏ – cam – hồng","gốm sứ, đá, đồ màu vàng
    #    – nâu đất","đồ kim loại trắng, màu trắng – bạc – ghi","nước, bể cá, màu đen – xanh dương"},IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)+
    #   1)&"; việc nên làm: "&INDEX({"đọc sách, làm vườn, đi bộ trong cây xanh","vận động ngoài nắng, giao lưu, hoạt động sôi nổi","nề nếp ổn định, ăn đúng bữ
    #   a, việc tay chân nhẹ","sắp xếp gọn gàng, thể thao có kỷ luật, hít thở sâu","bơi, thư giãn, ngủ đủ, uống đủ nước"},IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$1
    #   09,DUNG_HY_KY!$C$109)+1)&"."&CHAR(10)&"Hỷ thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),5)+1)&": "&INDEX({"giờ D
    #   ần – Mão (3–7h)","giờ Tị – Ngọ (9–13h)","giờ Thìn, Mùi, Tuất, Sửu (7–9h, 13–15h, 19–21h)","giờ Thân – Dậu (15–19h)","giờ Hợi – Tí (21–1h, nghỉ ngơi)"}
    #   ,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1)&"; vật dụng "&INDEX({"cây xanh, đồ gỗ, màu xanh lá","đèn ánh sáng ấm, nắng, đồ màu đỏ – 
    #   cam – hồng","gốm sứ, đá, đồ màu vàng – nâu đất","đồ kim loại trắng, màu trắng – bạc – ghi","nước, bể cá, màu đen – xanh dương"},IF(DOI_CHIEU!$C$33=2,D
    #   UNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1)&"; việc nên làm: "&INDEX({"đọc sách, làm vườn, đi bộ trong cây xanh","vận động ngoài nắng, giao lưu, hoạt động 
    #   sôi nổi","nề nếp ổn định, ăn đúng bữa, việc tay chân nhẹ","sắp xếp gọn gàng, thể thao có kỷ luật, hít thở sâu","bơi, thư giãn, ngủ đủ, uống đủ nước"},
    #   IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)+1)&"."&CHAR(10)&"Hạn chế: "&INDEX({"cây xanh, đồ gỗ, màu xanh lá","đèn ánh sáng ấm, nắng, đồ
    #    màu đỏ – cam – hồng","gốm sứ, đá, đồ màu vàng – nâu đất","đồ kim loại trắng, màu trắng – bạc – ghi","nước, bể cá, màu đen – xanh dương"},IF(DOI_CHIEU
    #   !$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)+1)&" (kỵ thần "&INDEX(HANH_TEN,MOD(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111),5)+1)&"). K
    #   hung giờ, vật dụng theo ngũ hành là kinh nghiệm phổ biến, ngoài sách; sách chỉ nêu nguyên tắc “đi theo phương hỉ, dụng thần” (P2 tr.631)."
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Dụng thần ',X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)), 5.0),1.0))),': '),X.INDEX(X.Arr([['giờ Dần – Mão (3–7h)', 'giờ Tị – Ngọ (9–13h)', 'giờ Thìn, Mùi, Tuất, Sửu (7–9h, 13–15h, 19–21h)', 'giờ Thân – Dậu (15–19h)', 'giờ Hợi – Tí (21–1h, nghỉ ngơi)']]), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),'; vật dụng '),X.INDEX(X.Arr([['cây xanh, đồ gỗ, màu xanh lá', 'đèn ánh sáng ấm, nắng, đồ màu đỏ – cam – hồng', 'gốm sứ, đá, đồ màu vàng – nâu đất', 'đồ kim loại trắng, màu trắng – bạc – ghi', 'nước, bể cá, màu đen – xanh dương']]), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),'; việc nên làm: '),X.INDEX(X.Arr([['đọc sách, làm vườn, đi bộ trong cây xanh', 'vận động ngoài nắng, giao lưu, hoạt động sôi nổi', 'nề nếp ổn định, ăn đúng bữa, việc tay chân nhẹ', 'sắp xếp gọn gàng, thể thao có kỷ luật, hít thở sâu', 'bơi, thư giãn, ngủ đủ, uống đủ nước']]), X.add((E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)),1.0))),'.'),X.CHAR(10.0)),'Hỷ thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)), 5.0),1.0))),': '),X.INDEX(X.Arr([['giờ Dần – Mão (3–7h)', 'giờ Tị – Ngọ (9–13h)', 'giờ Thìn, Mùi, Tuất, Sửu (7–9h, 13–15h, 19–21h)', 'giờ Thân – Dậu (15–19h)', 'giờ Hợi – Tí (21–1h, nghỉ ngơi)']]), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))),'; vật dụng '),X.INDEX(X.Arr([['cây xanh, đồ gỗ, màu xanh lá', 'đèn ánh sáng ấm, nắng, đồ màu đỏ – cam – hồng', 'gốm sứ, đá, đồ màu vàng – nâu đất', 'đồ kim loại trắng, màu trắng – bạc – ghi', 'nước, bể cá, màu đen – xanh dương']]), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))),'; việc nên làm: '),X.INDEX(X.Arr([['đọc sách, làm vườn, đi bộ trong cây xanh', 'vận động ngoài nắng, giao lưu, hoạt động sôi nổi', 'nề nếp ổn định, ăn đúng bữa, việc tay chân nhẹ', 'sắp xếp gọn gàng, thể thao có kỷ luật, hít thở sâu', 'bơi, thư giãn, ngủ đủ, uống đủ nước']]), X.add((E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)),1.0))),'.'),X.CHAR(10.0)),'Hạn chế: '),X.INDEX(X.Arr([['cây xanh, đồ gỗ, màu xanh lá', 'đèn ánh sáng ấm, nắng, đồ màu đỏ – cam – hồng', 'gốm sứ, đá, đồ màu vàng – nâu đất', 'đồ kim loại trắng, màu trắng – bạc – ghi', 'nước, bể cá, màu đen – xanh dương']]), X.add((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)),1.0))),' (kỵ thần '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)), 5.0),1.0))),'). Khung giờ, vật dụng theo ngũ hành là kinh nghiệm phổ biến, ngoài sách; sách chỉ nêu nguyên tắc “đi theo phương hỉ, dụng thần” (P2 tr.631).')

def f326(E, r, c):
    # Ô C190
    # = "[Sách] “Hạn nhỏ thường không cần hóa giải, chỉ cần dặn dò mệnh chủ chú ý là được”; “những hạn trước 60 tuổi thường dễ hóa giải” (P2 tr.662-663). Hóa 
    #   giải chỉ là bổ cứu hậu thiên (P2 tr.631) – không thay chẩn đoán, điều trị y khoa; không tự "&"dùng thuốc. Mỗi đầu năm đổi năm xem (NHAP_LIEU C31) để c
    #   ập nhật phương vị, cửu tinh, năm cần đề phòng."
    return X.cat('[Sách] “Hạn nhỏ thường không cần hóa giải, chỉ cần dặn dò mệnh chủ chú ý là được”; “những hạn trước 60 tuổi thường dễ hóa giải” (P2 tr.662-663). Hóa giải chỉ là bổ cứu hậu thiên (P2 tr.631) – không thay chẩn đoán, điều trị y khoa; không tự ','dùng thuốc. Mỗi đầu năm đổi năm xem (NHAP_LIEU C31) để cập nhật phương vị, cửu tinh, năm cần đề phòng.')

def f327(E, r, c):
    # Ô C193
    # = (DUNG_HY_KY!$C$180&IF(B2<>""," "&B2,"")&" Sách chỉ cho nguyên tắc; các ngưỡng số ghi “quy ước kỹ thuật” là của công cụ. Không thay thế lời khuyên y kh
    #   oa, pháp lý, tài chính.")&" "&("Kết luận tổng thể chỉ ghép các kết luận đã có căn cứ ở từng sheet, không thêm quy tắc. Độ tin vượng suy: "&DOI_CHIEU!$
    #   C$35&" – "&DOI_CHIEU!$F$35&" Sách nhấn mạnh phải xét tổng hợp, có cứu hay không, và kiểm nghiệm bằng sự việc đã qua (P1 tr.174, 201; P2 tr.294).")
    return X.cat(X.cat(X.cat(X.cat(E.v(8,180,3),(X.cat(' ',E.v(3,r-191,c-1)) if X.truthy(X.ne(E.v(3,r-191,c-1),'')) else '')),' Sách chỉ cho nguyên tắc; các ngưỡng số ghi “quy ước kỹ thuật” là của công cụ. Không thay thế lời khuyên y khoa, pháp lý, tài chính.'),' '),X.cat(X.cat(X.cat(X.cat('Kết luận tổng thể chỉ ghép các kết luận đã có căn cứ ở từng sheet, không thêm quy tắc. Độ tin vượng suy: ',E.v(7,35,3)),' – '),E.v(7,35,6)),' Sách nhấn mạnh phải xét tổng hợp, có cứu hay không, và kiểm nghiệm bằng sự việc đã qua (P1 tr.174, 201; P2 tr.294).'))

def f328(E, r, c):
    # Ô C195
    # = "Các dòng dưới là gợi ý theo kỹ thuật của tài liệu “Phân tích thí dụ thực tế”, đã viết thành điều kiện và thử trên 5 lá có sự việc thật. Không phải qu
    #   y tắc của sách, không làm thay đổi kết luận A–F; dùng để hỏi lại, đối chiếu. Mức tin ghi trong ngoặc."
    return 'Các dòng dưới là gợi ý theo kỹ thuật của tài liệu “Phân tích thí dụ thực tế”, đã viết thành điều kiện và thử trên 5 lá có sự việc thật. Không phải quy tắc của sách, không làm thay đổi kết luận A–F; dùng để hỏi lại, đối chiếu. Mức tin ghi trong ngoặc.'

def f329(E, r, c):
    # Ô H195
    # = LUAN_GIAI!$J$201
    return E.v(4,201,10)

def f330(E, r, c):
    # Ô C202
    # = "Tứ trụ "&C15&" "&D15&" "&E15&" "&F15&". Nhật chủ "&TRUNG_GIAN!$E$10&" ("&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&"), sinh tháng "&TRUNG_GIAN!$D$11&". Vượng
    #    suy: "&DOI_CHIEU!$C$34&", độ tin "&LOWER(DOI_CHIEU!$C$35)&" (PP1: "&C31&"; PP2: "&D31&"). Cách cục: "&IF(DOI_CHIEU!$C$33=2,D35,C35)&". "&DOI_CHIEU!$C
    #   $37&". [DOI_CHIEU, DUNG_HY_KY]"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Tứ trụ ',E.v(3,r-187,c)),' '),E.v(3,r-187,c+1)),' '),E.v(3,r-187,c+2)),' '),E.v(3,r-187,c+3)),'. Nhật chủ '),E.v(23,10,5)),' ('),X.INDEX(E.rg(24,33,1,37,1), X.add(E.v(23,7,5),1.0))),'), sinh tháng '),E.v(23,11,4)),'. Vượng suy: '),E.v(7,34,3)),', độ tin '),X.LOWER(E.v(7,35,3))),' (PP1: '),E.v(3,r-171,c)),'; PP2: '),E.v(3,r-171,c+1)),'). Cách cục: '),(E.v(3,r-167,c+1) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(3,r-167,c))),'. '),E.v(7,37,3)),'. [DOI_CHIEU, DUNG_HY_KY]')

def f331(E, r, c):
    # Ô H202
    # = LUAN_GIAI!$J$195
    return E.v(4,195,10)

def f332(E, r, c):
    # Ô C203
    # = TO_CHAT!C8&IF(TO_CHAT!C12="",""," Năng khiếu theo sách: "&TO_CHAT!C12)&" [TO_CHAT]"&" Tính cách (chương 22): "&LEFT(TO_CHAT!$C$25,MIN(LEN(TO_CHAT!$C$2
    #   5),400))&" [TO_CHAT]"
    return X.cat(X.cat(X.cat(X.cat(X.cat(E.v(20,r-195,c),('' if X.truthy(X.eq(E.v(20,r-191,c),'')) else X.cat(' Năng khiếu theo sách: ',E.v(20,r-191,c)))),' [TO_CHAT]'),' Tính cách (chương 22): '),X.LEFT(E.v(20,25,3), X.MIN(X.LEN(E.v(20,25,3)), 400.0))),' [TO_CHAT]')

def f333(E, r, c):
    # Ô H203
    # = LUAN_GIAI!$J$196
    return E.v(4,196,10)

def f334(E, r, c):
    # Ô C204
    # = "Sang – hèn: "&MENH_CUC!D19&". "&IF(TO_CHAT!C16="","","Thiên hướng: "&TO_CHAT!C16&" ")&"Nghề hợp dụng thần: "&TO_CHAT!C17&". [MENH_CUC, TO_CHAT]"&" Qu
    #   an vận (chương 19): "&LEFT(SU_NGHIEP!$J$7,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$7)-1,LEN(SU_NGHIEP!$J$7)))&"; học tập – thi cử: "&LEFT(SU_NGHIEP!$J$8,IFE
    #   RROR(FIND(CHAR(10),SU_NGHIEP!$J$8)-1,LEN(SU_NGHIEP!$J$8)))&"; danh tiếng: "&LEFT(SU_NGHIEP!$J$9,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$9)-1,LEN(SU_NGHIEP!
    #   $J$9)))&". [SU_NGHIEP]"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Sang – hèn: ',E.v(18,r-185,c+1)),'. '),('' if X.truthy(X.eq(E.v(20,r-188,c),'')) else X.cat(X.cat('Thiên hướng: ',E.v(20,r-188,c)),' '))),'Nghề hợp dụng thần: '),E.v(20,r-187,c)),'. [MENH_CUC, TO_CHAT]'),' Quan vận (chương 19): '),X.LEFT(E.v(21,7,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,7,10)),1.0), lambda: X.LEN(E.v(21,7,10))))),'; học tập – thi cử: '),X.LEFT(E.v(21,8,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,8,10)),1.0), lambda: X.LEN(E.v(21,8,10))))),'; danh tiếng: '),X.LEFT(E.v(21,9,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,9,10)),1.0), lambda: X.LEN(E.v(21,9,10))))),'. [SU_NGHIEP]')

def f335(E, r, c):
    # Ô H204
    # = LUAN_GIAI!$J$197
    return E.v(4,197,10)

def f336(E, r, c):
    # Ô C205
    # = "Bậc giàu: "&TAI_LOC!E17&". Giàu – nghèo (bảng P1 tr.173): "&TAI_LOC!C19&" – Nguồn tiền: "&IF(TAI_LOC!C21="","—",TAI_LOC!C21)&" Vận dễ phát tài: "&IF(
    #   TAI_LOC!C37="","không có",TAI_LOC!C37)&". [TAI_LOC, MENH_CUC]"&" Tài vận (chương 18): "&LEFT(SU_NGHIEP!$J$6,IFERROR(FIND(CHAR(10),SU_NGHIEP!$J$6)-1,LE
    #   N(SU_NGHIEP!$J$6)))&" – tài "&IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109),"là dụng thần",IF(MOD(TRUNG_GIAN!$
    #   E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110),"là hỷ thần",IF(MOD(TRUNG_GIAN!$E$7+2,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUN
    #   G_HY_KY!$C$111),"là kỵ thần",IF(AND(SU_NGHIEP!$Z$2=MOD(TRUNG_GIAN!$E$7+2,5),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_K
    #   Y!$C$109),MOD(TRUNG_GIAN!$E$7+2,5)<>IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)),"là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.1
    #   80)","không phải dụng/hỷ/kỵ thần (nhàn thần)"))))&". [SU_NGHIEP]"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Bậc giàu: ',E.v(19,r-188,c+2)),'. Giàu – nghèo (bảng P1 tr.173): '),E.v(19,r-186,c)),' – Nguồn tiền: '),('—' if X.truthy(X.eq(E.v(19,r-184,c),'')) else E.v(19,r-184,c))),' Vận dễ phát tài: '),('không có' if X.truthy(X.eq(E.v(19,r-168,c),'')) else E.v(19,r-168,c))),'. [TAI_LOC, MENH_CUC]'),' Tài vận (chương 18): '),X.LEFT(E.v(21,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(21,6,10)),1.0), lambda: X.LEN(E.v(21,6,10))))),' – tài '),('là dụng thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3)))) else ('là hỷ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3)))) else ('là kỵ thần' if X.truthy(X.eq(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else ('là nhóm nhiều nhất – cũng là kỵ theo cân bằng (P1 tr.180)' if X.truthy(X.AND(X.eq(E.v(21,2,26),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,109,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,109,3))), X.ne(X.MOD(X.add(E.v(23,7,5),2.0), 5.0),(E.v(8,110,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,110,3))))) else 'không phải dụng/hỷ/kỵ thần (nhàn thần)'))))),'. [SU_NGHIEP]')

def f337(E, r, c):
    # Ô H205
    # = LUAN_GIAI!$J$198
    return E.v(4,198,10)

def f338(E, r, c):
    # Ô C206
    # = "Ông bà – tổ nghiệp: "&LUC_THAN!$X$30&"; cha: "&LUC_THAN!$X$31&"; mẹ: "&LUC_THAN!$X$32&"; anh chị em: "&LUC_THAN!$X$33&"; con cái: "&LUC_THAN!$X$34&".
    #   "&CHAR(10)&IF(LUU_NIEN!$B$5<=18,"(Lưu ý tuổi: chưa thành niên – mục hôn nhân, con cái là xu hướng xa, chưa xét lúc này.) ","")&"Hôn nhân – mức "&UPPER
    #   (HON_NHAN!$X$64)&". Lý do: "&IF(NHAP_LIEU!$C$22="Nam","vợ (sao Tài)","chồng (sao Quan sát)")&": "&IF(HON_NHAN!$M$64=0,"không có sao này trong tứ trụ",
    #   HON_NHAN!$M$64&" chữ; "&IF(HON_NHAN!$O$64=1,"vượng tại lệnh","không vượng tại lệnh")&"; "&IF(HON_NHAN!$N$64=1,"ở trụ ngày","không ở trụ ngày")&"; vai 
    #   trò: "&SUBSTITUTE(IF(HON_NHAN!$R$64=HON_NHAN!$S$64,HON_NHAN!$R$64,"PP1 "&HON_NHAN!$R$64&", PP2 "&HON_NHAN!$S$64),"—","nhàn thần")&IF(HON_NHAN!$Q$64=0,
    #   "; không chữ nào còn nguyên vẹn (bị xung, khắc, hợp mất hoặc tuần không)",""))&IF(HON_NHAN!$T$64=1,"; chi ngày (cung phối ngẫu) bị xung/hình","")&". S
    #   ách có "&HON_NHAN!$AI$64&" dấu hiệu thuận, "&HON_NHAN!$AG$64&" dấu hiệu bất lợi về hôn nhân (chi tiết HON_NHAN dòng 64, 72-107). Thời điểm: "&HON_NHAN
    #   !$D$66&" [LUC_THAN, HON_NHAN]"&CHAR(10)&"(Lý giải từng người – vì sao, sách nêu gì: KET_LUAN dòng 110-119; chi tiết LUC_THAN dòng 30-34, HON_NHAN.)"&C
    #   HAR(10)&"(Giới hạn của công cụ: sách chưa có quy tắc riêng về ông bà ngoại – họ ngoại, và về năm sinh con cụ thể; công cụ chỉ có dấu hiệu chung về ngư
    #   ời thân và con cái theo năm – xem KET_LUAN dòng 144.)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Ông bà – tổ nghiệp: ',E.v(16,30,24)),'; cha: '),E.v(16,31,24)),'; mẹ: '),E.v(16,32,24)),'; anh chị em: '),E.v(16,33,24)),'; con cái: '),E.v(16,34,24)),'.'),X.CHAR(10.0)),('(Lưu ý tuổi: chưa thành niên – mục hôn nhân, con cái là xu hướng xa, chưa xét lúc này.) ' if X.truthy(X.le(E.v(15,5,2),18.0)) else '')),'Hôn nhân – mức '),X.UPPER(E.v(17,64,24))),'. Lý do: '),('vợ (sao Tài)' if X.truthy(X.eq(E.v(2,22,3),'Nam')) else 'chồng (sao Quan sát)')),': '),('không có sao này trong tứ trụ' if X.truthy(X.eq(E.v(17,64,13),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(17,64,13),' chữ; '),('vượng tại lệnh' if X.truthy(X.eq(E.v(17,64,15),1.0)) else 'không vượng tại lệnh')),'; '),('ở trụ ngày' if X.truthy(X.eq(E.v(17,64,14),1.0)) else 'không ở trụ ngày')),'; vai trò: '),X.SUBSTITUTE((E.v(17,64,18) if X.truthy(X.eq(E.v(17,64,18),E.v(17,64,19))) else X.cat(X.cat(X.cat('PP1 ',E.v(17,64,18)),', PP2 '),E.v(17,64,19))), '—', 'nhàn thần')),('; không chữ nào còn nguyên vẹn (bị xung, khắc, hợp mất hoặc tuần không)' if X.truthy(X.eq(E.v(17,64,17),0.0)) else '')))),('; chi ngày (cung phối ngẫu) bị xung/hình' if X.truthy(X.eq(E.v(17,64,20),1.0)) else '')),'. Sách có '),E.v(17,64,35)),' dấu hiệu thuận, '),E.v(17,64,33)),' dấu hiệu bất lợi về hôn nhân (chi tiết HON_NHAN dòng 64, 72-107). Thời điểm: '),E.v(17,66,4)),' [LUC_THAN, HON_NHAN]'),X.CHAR(10.0)),'(Lý giải từng người – vì sao, sách nêu gì: KET_LUAN dòng 110-119; chi tiết LUC_THAN dòng 30-34, HON_NHAN.)'),X.CHAR(10.0)),'(Giới hạn của công cụ: sách chưa có quy tắc riêng về ông bà ngoại – họ ngoại, và về năm sinh con cụ thể; công cụ chỉ có dấu hiệu chung về người thân và con cái theo năm – xem KET_LUAN dòng 144.)')

def f339(E, r, c):
    # Ô H206
    # = LUAN_GIAI!$J$199
    return E.v(4,199,10)

def f340(E, r, c):
    # Ô C207
    # = "Thọ – yểu: "&MENH_CUC!D21&" Năm có nhiều dấu hiệu nặng (cả đời): "&IF(MENH_CUC!C33="","không có năm nào vượt ngưỡng",TRIM(MENH_CUC!C33&" "&MENH_CUC!C
    #   34&" "&MENH_CUC!C35&" "&MENH_CUC!C36&" "&MENH_CUC!C37))&". Công cụ chỉ liệt kê dấu hiệu, không kết luận năm mất (P1 tr.201). [MENH_CUC]"&" Sức khỏe (c
    #   hương 21): "&LEFT(SUC_KHOE!$J$6,IFERROR(FIND(CHAR(10),SUC_KHOE!$J$6)-1,LEN(SUC_KHOE!$J$6)))&"; bệnh tật: "&LEFT(SUC_KHOE!$J$7,IFERROR(FIND(CHAR(10),SU
    #   C_KHOE!$J$7)-1,LEN(SUC_KHOE!$J$7)))&"; tai họa: "&LEFT(SUC_KHOE!$J$8,IFERROR(FIND(CHAR(10),SUC_KHOE!$J$8)-1,LEN(SUC_KHOE!$J$8)))&". "&SUBSTITUTE(SUC_K
    #   HOE!$C$38,CHAR(10),"; ")&" [SUC_KHOE]"&DUNG_HY_KY!$D$195
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Thọ – yểu: ',E.v(18,r-186,c+1)),' Năm có nhiều dấu hiệu nặng (cả đời): '),('không có năm nào vượt ngưỡng' if X.truthy(X.eq(E.v(18,r-174,c),'')) else X.TRIM(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(18,r-174,c),' '),E.v(18,r-173,c)),' '),E.v(18,r-172,c)),' '),E.v(18,r-171,c)),' '),E.v(18,r-170,c))))),'. Công cụ chỉ liệt kê dấu hiệu, không kết luận năm mất (P1 tr.201). [MENH_CUC]'),' Sức khỏe (chương 21): '),X.LEFT(E.v(22,6,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,6,10)),1.0), lambda: X.LEN(E.v(22,6,10))))),'; bệnh tật: '),X.LEFT(E.v(22,7,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,7,10)),1.0), lambda: X.LEN(E.v(22,7,10))))),'; tai họa: '),X.LEFT(E.v(22,8,10), X.IFERROR(lambda: X.sub(X.FIND(X.CHAR(10.0), E.v(22,8,10)),1.0), lambda: X.LEN(E.v(22,8,10))))),'. '),X.SUBSTITUTE(E.v(22,38,3), X.CHAR(10.0), '; ')),' [SUC_KHOE]'),E.v(8,195,4))

def f341(E, r, c):
    # Ô H207
    # = LUAN_GIAI!$J$200
    return E.v(4,200,10)

def f342(E, r, c):
    # Ô C208
    # = "Cát – hung chung: "&MENH_CUC!D20&". Đại vận tốt (dụng/hỷ): "&IF(LUC_THAN!$Y$31="","không có",LUC_THAN!$Y$31)&"; đại vận xấu (kỵ): "&IF(LUC_THAN!$Z$31
    #   ="","không có",LUC_THAN!$Z$31)&". Hiện tại: "&C163&" [DAI_VAN, LUU_NIEN]"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Cát – hung chung: ',E.v(18,r-188,c+1)),'. Đại vận tốt (dụng/hỷ): '),('không có' if X.truthy(X.eq(E.v(16,31,25),'')) else E.v(16,31,25))),'; đại vận xấu (kỵ): '),('không có' if X.truthy(X.eq(E.v(16,31,26),'')) else E.v(16,31,26))),'. Hiện tại: '),E.v(3,r-45,c)),' [DAI_VAN, LUU_NIEN]')

def f343(E, r, c):
    # Ô H208
    # = LUAN_GIAI!$A$137&": "&LUAN_GIAI!$C$137&CHAR(10)&CHAR(10)&LUAN_GIAI!$A$138&": "&LUAN_GIAI!$C$138&CHAR(10)&CHAR(10)&LUAN_GIAI!$A$139&": "&LUAN_GIAI!$C$1
    #   39&CHAR(10)&CHAR(10)&LUAN_GIAI!$A$140&": "&LUAN_GIAI!$C$140
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(4,137,1),': '),E.v(4,137,3)),X.CHAR(10.0)),X.CHAR(10.0)),E.v(4,138,1)),': '),E.v(4,138,3)),X.CHAR(10.0)),X.CHAR(10.0)),E.v(4,139,1)),': '),E.v(4,139,3)),X.CHAR(10.0)),X.CHAR(10.0)),E.v(4,140,1)),': '),E.v(4,140,3))

def f344(E, r, c):
    # Ô C209
    # = "Kết luận tổng thể chỉ ghép các kết luận đã có căn cứ ở từng sheet, không thêm quy tắc. Độ tin vượng suy: "&DOI_CHIEU!$C$35&" – "&DOI_CHIEU!$F$35&" Sá
    #   ch nhấn mạnh phải xét tổng hợp, có cứu hay không, và kiểm nghiệm bằng sự việc đã qua (P1 tr.174, 201; P2 tr.294)."
    return X.cat(X.cat(X.cat(X.cat('Kết luận tổng thể chỉ ghép các kết luận đã có căn cứ ở từng sheet, không thêm quy tắc. Độ tin vượng suy: ',E.v(7,35,3)),' – '),E.v(7,35,6)),' Sách nhấn mạnh phải xét tổng hợp, có cứu hay không, và kiểm nghiệm bằng sự việc đã qua (P1 tr.174, 201; P2 tr.294).')

def f345(E, r, c):
    # Ô H209
    # = LUAN_GIAI!$A$133&": "&LUAN_GIAI!$C$133&CHAR(10)&CHAR(10)&LUAN_GIAI!$A$134&": "&LUAN_GIAI!$C$134&CHAR(10)&CHAR(10)&LUAN_GIAI!$A$135&": "&LUAN_GIAI!$C$1
    #   35&CHAR(10)&CHAR(10)&LUAN_GIAI!$A$136&": "&LUAN_GIAI!$C$136
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(4,133,1),': '),E.v(4,133,3)),X.CHAR(10.0)),X.CHAR(10.0)),E.v(4,134,1)),': '),E.v(4,134,3)),X.CHAR(10.0)),X.CHAR(10.0)),E.v(4,135,1)),': '),E.v(4,135,3)),X.CHAR(10.0)),X.CHAR(10.0)),E.v(4,136,1)),': '),E.v(4,136,3))

def f346(E, r, c):
    # Ô AG210
    # = MOD(TRUNG_GIAN!$E$7+3,5)
    return X.MOD(X.add(E.v(23,7,5),3.0), 5.0)

def f347(E, r, c):
    # Ô AG211
    # = MOD(TRUNG_GIAN!$E$5,2)
    return X.MOD(E.v(23,5,5), 2.0)

def f348(E, r, c):
    # Ô AG212
    # = IF(IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1,1,0)
    return (1.0 if X.truthy(X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1)) else 0.0)

def f349(E, r, c):
    # Ô AG213
    # = IF($AG$149<>"KHÔNG",1,0)
    return (1.0 if X.truthy(X.ne(E.v(3,149,33),'KHÔNG')) else 0.0)

def f350(E, r, c):
    # Ô AG214
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$C$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)=4,AND(OR(DAI_VAN!$C$34=0,DAI_VAN!$C$34=4,DAI_VAN!$C$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$C$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$C$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$C$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$C$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$C$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$C$34=0,DAI_VAN!$C$34=4,DAI_VAN!$C$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$C$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$C$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$C$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$C$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$C$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$C$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$C$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$C$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$C$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$C$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$C$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   C$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$C$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$3
    #   4+1)=2),INT(DAI_VAN!$C$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)<>1,INT(DAI_VAN!$C$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$C$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)=1,INT(DAI_VAN!$C$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$C$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,3),0.0), X.eq(E.v(14,34,3),4.0), X.eq(E.v(14,34,3),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,3),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,3),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,3),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,3),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,3),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,3),0.0), X.eq(E.v(14,34,3),4.0), X.eq(E.v(14,34,3),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,3),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,3),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,3),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,3),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,3),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,3),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,3),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,3),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,3),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,3),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,3),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,3),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,3),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,3),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,3),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,3),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,3),1.0)),0.0))) else ''))))

def f351(E, r, c):
    # Ô AH214
    # = IF(AG214="","","• "&DAI_VAN!$C$35&" ("&DAI_VAN!$C$37&"–"&DAI_VAN!$C$38&"): "&LEFT(AG214,LEN(AG214)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,3)),' ('),E.v(14,37,3)),'–'),E.v(14,38,3)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f352(E, r, c):
    # Ô AG215
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$D$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)=4,AND(OR(DAI_VAN!$D$34=0,DAI_VAN!$D$34=4,DAI_VAN!$D$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$D$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$D$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$D$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$D$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$D$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$D$34=0,DAI_VAN!$D$34=4,DAI_VAN!$D$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$D$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$D$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$D$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$D$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$D$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$D$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$D$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$D$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$D$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$D$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$D$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   D$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$D$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$3
    #   4+1)=2),INT(DAI_VAN!$D$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)<>1,INT(DAI_VAN!$D$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$D$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)=1,INT(DAI_VAN!$D$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$D$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,4),0.0), X.eq(E.v(14,34,4),4.0), X.eq(E.v(14,34,4),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,4),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,4),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,4),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,4),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,4),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,4),0.0), X.eq(E.v(14,34,4),4.0), X.eq(E.v(14,34,4),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,4),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,4),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,4),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,4),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,4),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,4),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,4),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,4),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,4),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,4),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,4),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,4),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,4),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,4),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,4),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,4),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,4),1.0)),0.0))) else ''))))

def f353(E, r, c):
    # Ô AH215
    # = IF(AG215="","","• "&DAI_VAN!$D$35&" ("&DAI_VAN!$D$37&"–"&DAI_VAN!$D$38&"): "&LEFT(AG215,LEN(AG215)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,4)),' ('),E.v(14,37,4)),'–'),E.v(14,38,4)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f354(E, r, c):
    # Ô AG216
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$E$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)=4,AND(OR(DAI_VAN!$E$34=0,DAI_VAN!$E$34=4,DAI_VAN!$E$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$E$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$E$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$E$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$E$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$E$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$E$34=0,DAI_VAN!$E$34=4,DAI_VAN!$E$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$E$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$E$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$E$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$E$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$E$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$E$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$E$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$E$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$E$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$E$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$E$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   E$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$E$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$3
    #   4+1)=2),INT(DAI_VAN!$E$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)<>1,INT(DAI_VAN!$E$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$E$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)=1,INT(DAI_VAN!$E$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$E$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,5),0.0), X.eq(E.v(14,34,5),4.0), X.eq(E.v(14,34,5),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,5),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,5),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,5),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,5),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,5),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,5),0.0), X.eq(E.v(14,34,5),4.0), X.eq(E.v(14,34,5),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,5),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,5),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,5),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,5),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,5),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,5),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,5),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,5),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,5),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,5),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,5),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,5),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,5),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,5),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,5),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,5),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,5),1.0)),0.0))) else ''))))

def f355(E, r, c):
    # Ô AH216
    # = IF(AG216="","","• "&DAI_VAN!$E$35&" ("&DAI_VAN!$E$37&"–"&DAI_VAN!$E$38&"): "&LEFT(AG216,LEN(AG216)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,5)),' ('),E.v(14,37,5)),'–'),E.v(14,38,5)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f356(E, r, c):
    # Ô AG217
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$F$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)=4,AND(OR(DAI_VAN!$F$34=0,DAI_VAN!$F$34=4,DAI_VAN!$F$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$F$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$F$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$F$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$F$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$F$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$F$34=0,DAI_VAN!$F$34=4,DAI_VAN!$F$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$F$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$F$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$F$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$F$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$F$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$F$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$F$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$F$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$F$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$F$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$F$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   F$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$F$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$3
    #   4+1)=2),INT(DAI_VAN!$F$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)<>1,INT(DAI_VAN!$F$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$F$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)=1,INT(DAI_VAN!$F$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$F$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,6),0.0), X.eq(E.v(14,34,6),4.0), X.eq(E.v(14,34,6),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,6),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,6),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,6),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,6),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,6),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,6),0.0), X.eq(E.v(14,34,6),4.0), X.eq(E.v(14,34,6),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,6),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,6),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,6),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,6),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,6),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,6),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,6),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,6),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,6),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,6),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,6),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,6),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,6),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,6),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,6),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,6),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,6),1.0)),0.0))) else ''))))

def f357(E, r, c):
    # Ô AH217
    # = IF(AG217="","","• "&DAI_VAN!$F$35&" ("&DAI_VAN!$F$37&"–"&DAI_VAN!$F$38&"): "&LEFT(AG217,LEN(AG217)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,6)),' ('),E.v(14,37,6)),'–'),E.v(14,38,6)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f358(E, r, c):
    # Ô AG218
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$G$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)=4,AND(OR(DAI_VAN!$G$34=0,DAI_VAN!$G$34=4,DAI_VAN!$G$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$G$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$G$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$G$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$G$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$G$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$G$34=0,DAI_VAN!$G$34=4,DAI_VAN!$G$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$G$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$G$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$G$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$G$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$G$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$G$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$G$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$G$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$G$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$G$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$G$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   G$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$G$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$3
    #   4+1)=2),INT(DAI_VAN!$G$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)<>1,INT(DAI_VAN!$G$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$G$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)=1,INT(DAI_VAN!$G$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$G$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,7),0.0), X.eq(E.v(14,34,7),4.0), X.eq(E.v(14,34,7),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,7),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,7),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,7),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,7),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,7),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,7),0.0), X.eq(E.v(14,34,7),4.0), X.eq(E.v(14,34,7),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,7),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,7),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,7),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,7),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,7),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,7),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,7),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,7),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,7),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,7),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,7),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,7),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,7),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,7),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,7),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,7),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,7),1.0)),0.0))) else ''))))

def f359(E, r, c):
    # Ô AH218
    # = IF(AG218="","","• "&DAI_VAN!$G$35&" ("&DAI_VAN!$G$37&"–"&DAI_VAN!$G$38&"): "&LEFT(AG218,LEN(AG218)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,7)),' ('),E.v(14,37,7)),'–'),E.v(14,38,7)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f360(E, r, c):
    # Ô AG219
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$H$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)=4,AND(OR(DAI_VAN!$H$34=0,DAI_VAN!$H$34=4,DAI_VAN!$H$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$H$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$H$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$H$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$H$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$H$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$H$34=0,DAI_VAN!$H$34=4,DAI_VAN!$H$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$H$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$H$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$H$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$H$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$H$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$H$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$H$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$H$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$H$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$H$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$H$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   H$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$H$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$3
    #   4+1)=2),INT(DAI_VAN!$H$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)<>1,INT(DAI_VAN!$H$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$H$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)=1,INT(DAI_VAN!$H$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$H$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,8),0.0), X.eq(E.v(14,34,8),4.0), X.eq(E.v(14,34,8),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,8),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,8),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,8),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,8),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,8),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,8),0.0), X.eq(E.v(14,34,8),4.0), X.eq(E.v(14,34,8),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,8),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,8),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,8),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,8),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,8),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,8),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,8),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,8),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,8),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,8),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,8),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,8),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,8),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,8),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,8),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,8),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,8),1.0)),0.0))) else ''))))

def f361(E, r, c):
    # Ô AH219
    # = IF(AG219="","","• "&DAI_VAN!$H$35&" ("&DAI_VAN!$H$37&"–"&DAI_VAN!$H$38&"): "&LEFT(AG219,LEN(AG219)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,8)),' ('),E.v(14,37,8)),'–'),E.v(14,38,8)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f362(E, r, c):
    # Ô AG220
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$I$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)=4,AND(OR(DAI_VAN!$I$34=0,DAI_VAN!$I$34=4,DAI_VAN!$I$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$I$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$I$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$I$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$I$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$I$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$I$34=0,DAI_VAN!$I$34=4,DAI_VAN!$I$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$I$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$I$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$I$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$I$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$I$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$I$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$I$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$I$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$I$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$I$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$I$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   I$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$I$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$3
    #   4+1)=2),INT(DAI_VAN!$I$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)<>1,INT(DAI_VAN!$I$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$I$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)=1,INT(DAI_VAN!$I$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$I$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,9),0.0), X.eq(E.v(14,34,9),4.0), X.eq(E.v(14,34,9),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,9),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,9),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,9),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,9),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,9),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,9),0.0), X.eq(E.v(14,34,9),4.0), X.eq(E.v(14,34,9),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,9),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,9),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,9),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,9),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,9),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,9),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,9),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,9),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,9),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,9),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,9),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,9),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,9),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,9),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,9),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,9),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,9),1.0)),0.0))) else ''))))

def f363(E, r, c):
    # Ô AH220
    # = IF(AG220="","","• "&DAI_VAN!$I$35&" ("&DAI_VAN!$I$37&"–"&DAI_VAN!$I$38&"): "&LEFT(AG220,LEN(AG220)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,9)),' ('),E.v(14,37,9)),'–'),E.v(14,38,9)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f364(E, r, c):
    # Ô AG221
    # = IF($AG$149="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$J$33/2)=4,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)=4,AND(OR(DAI_VAN!$J$34=0,DAI_VAN!$J$34=4,DAI_VAN!$J$34=
    #   8),(COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$J$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$J$34+4,12)),1,0))>0,(
    #   COUNTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$J$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$J$34+8,12)),1,0))>0)),INT
    #   (DAI_VAN!$J$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)<>1),"nặng UT1 (thêm Thủy"&IF(AND(OR(DAI_VAN!$J$34=0,DAI_VAN!$J$34=4,DAI_VAN!$J$34=8),(COU
    #   NTIF(TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$J$34+4,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$J$34+4,12)),1,0))>0,(COUNTIF
    #   (TRUNG_GIAN!$C$6:$E$6,MOD(DAI_VAN!$J$34+8,12))+IF(AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),TRUNG_GIAN!$F$6=MOD(DAI_VAN!$J$34+8,12)),1,0))>0),", hợp Thủy
    #    cục","")&"); ",IF(OR(INT(DAI_VAN!$J$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)=1),"bù UT1 (có Hỏa); ","")))&IF($AG$150="KHÔNG","",IF(DAI_VAN!$J$
    #   39="Thiên ấn","nặng UT2 (thêm thiên ấn); ",IF(OR(INT(DAI_VAN!$J$33/2)=MOD(TRUNG_GIAN!$E$7+2,5),INT(DAI_VAN!$J$33/2)=MOD(TRUNG_GIAN!$E$7+3,5),INDEX(CSD
    #   L!$C$18:$C$29,DAI_VAN!$J$34+1)=MOD(TRUNG_GIAN!$E$7+2,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)=MOD(TRUNG_GIAN!$E$7+3,5)),"điều tiết UT2 (có tài/quan)
    #   ; ","")))&IF(AND(DAI_VAN!$J$39="Thất sát",IFERROR(INDEX(PP2_DINH_LUONG!$F$91:$F$95,MOD(MOD(TRUNG_GIAN!$E$7+4,5),5)+1)/PP2_DINH_LUONG!$C$100,0)<0.1),"K
    #   ÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ",IF(AND($AG$151<>"KHÔNG",OR(INT(DAI_VAN!$J$33/2)=MOD(TRUNG_GIAN!$E$7+4,5),INDEX(CSDL!$C$18:$C$29,DAI_VAN!$
    #   J$34+1)=MOD(TRUNG_GIAN!$E$7+4,5))),"hóa UT3 (có Ấn); ",""))&IF($AG$152="KHÔNG","",IF(AND(OR(INT(DAI_VAN!$J$33/2)=2,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$3
    #   4+1)=2),INT(DAI_VAN!$J$33/2)<>1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)<>1,INT(DAI_VAN!$J$33/2)<>0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)<>0),"nặng U
    #   T4 (thêm Thổ); ",IF(OR(INT(DAI_VAN!$J$33/2)=1,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)=1,INT(DAI_VAN!$J$33/2)=0,INDEX(CSDL!$C$18:$C$29,DAI_VAN!$J$34+1)
    #   =0),"bù UT4 (có Hỏa/Mộc); ","")))
    return X.cat(X.cat(X.cat(('' if X.truthy(X.eq(E.v(3,149,33),'KHÔNG')) else (X.cat(X.cat('nặng UT1 (thêm Thủy',(', hợp Thủy cục' if X.truthy(X.AND(X.OR(X.eq(E.v(14,34,10),0.0), X.eq(E.v(14,34,10),4.0), X.eq(E.v(14,34,10),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,10),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,10),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,10),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,10),8.0), 12.0)))) else 0.0)),0.0))) else '')),'); ') if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,10),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),4.0), X.AND(X.OR(X.eq(E.v(14,34,10),0.0), X.eq(E.v(14,34,10),4.0), X.eq(E.v(14,34,10),8.0)), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,10),4.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,10),4.0), 12.0)))) else 0.0)),0.0), X.gt(X.add(X.COUNTIF(E.rg(23,6,3,6,5), X.MOD(X.add(E.v(14,34,10),8.0), 12.0)),(1.0 if X.truthy(X.AND(X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(E.v(23,6,6),X.MOD(X.add(E.v(14,34,10),8.0), 12.0)))) else 0.0)),0.0))), X.ne(X.INT(X.div(E.v(14,33,10),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),1.0))) else ('bù UT1 (có Hỏa); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,10),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),1.0))) else ''))),('' if X.truthy(X.eq(E.v(3,150,33),'KHÔNG')) else ('nặng UT2 (thêm thiên ấn); ' if X.truthy(X.eq(E.v(14,39,10),'Thiên ấn')) else ('điều tiết UT2 (có tài/quan); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,10),2.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INT(X.div(E.v(14,33,10),2.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),X.MOD(X.add(E.v(23,7,5),2.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),X.MOD(X.add(E.v(23,7,5),3.0), 5.0)))) else '')))),('KÍCH HOẠT UT3 (vận thất sát, gốc thiếu Ấn); ' if X.truthy(X.AND(X.eq(E.v(14,39,10),'Thất sát'), X.lt(X.IFERROR(lambda: X.div(X.INDEX(E.rg(10,91,6,95,6), X.add(X.MOD(X.MOD(X.add(E.v(23,7,5),4.0), 5.0), 5.0),1.0)),E.v(10,100,3)), lambda: 0.0),0.1))) else ('hóa UT3 (có Ấn); ' if X.truthy(X.AND(X.ne(E.v(3,151,33),'KHÔNG'), X.OR(X.eq(X.INT(X.div(E.v(14,33,10),2.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0)), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),X.MOD(X.add(E.v(23,7,5),4.0), 5.0))))) else ''))),('' if X.truthy(X.eq(E.v(3,152,33),'KHÔNG')) else ('nặng UT4 (thêm Thổ); ' if X.truthy(X.AND(X.OR(X.eq(X.INT(X.div(E.v(14,33,10),2.0)),2.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),2.0)), X.ne(X.INT(X.div(E.v(14,33,10),2.0)),1.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),1.0), X.ne(X.INT(X.div(E.v(14,33,10),2.0)),0.0), X.ne(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),0.0))) else ('bù UT4 (có Hỏa/Mộc); ' if X.truthy(X.OR(X.eq(X.INT(X.div(E.v(14,33,10),2.0)),1.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),1.0), X.eq(X.INT(X.div(E.v(14,33,10),2.0)),0.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(14,34,10),1.0)),0.0))) else ''))))

def f365(E, r, c):
    # Ô AH221
    # = IF(AG221="","","• "&DAI_VAN!$J$35&" ("&DAI_VAN!$J$37&"–"&DAI_VAN!$J$38&"): "&LEFT(AG221,LEN(AG221)-2)&CHAR(10))
    return ('' if X.truthy(X.eq(E.v(3,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',E.v(14,35,10)),' ('),E.v(14,37,10)),'–'),E.v(14,38,10)),'): '),X.LEFT(E.v(3,r,c-1), X.sub(X.LEN(E.v(3,r,c-1)),2.0))),X.CHAR(10.0)))

def f366(E, r, c):
    # Ô AG222
    # = IF(AH222+AI222>0,(NHAP_LIEU!$C$31+0)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+0)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+0)-4,12)+1)&" ("&IF(AH2
    #   22=1,"UT3","")&IF(AH222+AI222=2,"+","")&IF(AI222=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),0.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),0.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),0.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f367(E, r, c):
    # Ô AH222
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+0)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+0)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),0.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),0.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f368(E, r, c):
    # Ô AI222
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+0)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+0)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),0.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),0.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f369(E, r, c):
    # Ô AG223
    # = IF(AH223+AI223>0,(NHAP_LIEU!$C$31+1)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+1)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+1)-4,12)+1)&" ("&IF(AH2
    #   23=1,"UT3","")&IF(AH223+AI223=2,"+","")&IF(AI223=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),1.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),1.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),1.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f370(E, r, c):
    # Ô AH223
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+1)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+1)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),1.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),1.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f371(E, r, c):
    # Ô AI223
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+1)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+1)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),1.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),1.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f372(E, r, c):
    # Ô AG224
    # = IF(AH224+AI224>0,(NHAP_LIEU!$C$31+2)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+2)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+2)-4,12)+1)&" ("&IF(AH2
    #   24=1,"UT3","")&IF(AH224+AI224=2,"+","")&IF(AI224=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),2.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),2.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),2.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f373(E, r, c):
    # Ô AH224
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+2)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+2)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),2.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),2.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f374(E, r, c):
    # Ô AI224
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+2)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+2)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),2.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),2.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f375(E, r, c):
    # Ô AG225
    # = IF(AH225+AI225>0,(NHAP_LIEU!$C$31+3)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+3)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+3)-4,12)+1)&" ("&IF(AH2
    #   25=1,"UT3","")&IF(AH225+AI225=2,"+","")&IF(AI225=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),3.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),3.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),3.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f376(E, r, c):
    # Ô AH225
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+3)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+3)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),3.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),3.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f377(E, r, c):
    # Ô AI225
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+3)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+3)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),3.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),3.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f378(E, r, c):
    # Ô AG226
    # = IF(AH226+AI226>0,(NHAP_LIEU!$C$31+4)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+4)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+4)-4,12)+1)&" ("&IF(AH2
    #   26=1,"UT3","")&IF(AH226+AI226=2,"+","")&IF(AI226=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),4.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),4.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),4.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f379(E, r, c):
    # Ô AH226
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+4)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+4)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),4.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),4.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f380(E, r, c):
    # Ô AI226
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+4)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+4)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),4.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),4.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f381(E, r, c):
    # Ô AG227
    # = IF(AH227+AI227>0,(NHAP_LIEU!$C$31+5)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+5)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+5)-4,12)+1)&" ("&IF(AH2
    #   27=1,"UT3","")&IF(AH227+AI227=2,"+","")&IF(AI227=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),5.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),5.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),5.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f382(E, r, c):
    # Ô AH227
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+5)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+5)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),5.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),5.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f383(E, r, c):
    # Ô AI227
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+5)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+5)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),5.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),5.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f384(E, r, c):
    # Ô AG228
    # = IF(AH228+AI228>0,(NHAP_LIEU!$C$31+6)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+6)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+6)-4,12)+1)&" ("&IF(AH2
    #   28=1,"UT3","")&IF(AH228+AI228=2,"+","")&IF(AI228=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),6.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),6.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),6.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f385(E, r, c):
    # Ô AH228
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+6)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+6)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),6.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),6.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f386(E, r, c):
    # Ô AI228
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+6)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+6)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),6.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),6.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f387(E, r, c):
    # Ô AG229
    # = IF(AH229+AI229>0,(NHAP_LIEU!$C$31+7)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+7)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+7)-4,12)+1)&" ("&IF(AH2
    #   29=1,"UT3","")&IF(AH229+AI229=2,"+","")&IF(AI229=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),7.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),7.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),7.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f388(E, r, c):
    # Ô AH229
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+7)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+7)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),7.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),7.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f389(E, r, c):
    # Ô AI229
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+7)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+7)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),7.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),7.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f390(E, r, c):
    # Ô AG230
    # = IF(AH230+AI230>0,(NHAP_LIEU!$C$31+8)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+8)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+8)-4,12)+1)&" ("&IF(AH2
    #   30=1,"UT3","")&IF(AH230+AI230=2,"+","")&IF(AI230=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),8.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),8.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),8.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f391(E, r, c):
    # Ô AH230
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+8)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+8)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),8.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),8.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f392(E, r, c):
    # Ô AI230
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+8)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+8)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),8.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),8.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f393(E, r, c):
    # Ô AG231
    # = IF(AH231+AI231>0,(NHAP_LIEU!$C$31+9)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+9)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+9)-4,12)+1)&" ("&IF(AH2
    #   31=1,"UT3","")&IF(AH231+AI231=2,"+","")&IF(AI231=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),9.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),9.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),9.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f394(E, r, c):
    # Ô AH231
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+9)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+9)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),9.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),9.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f395(E, r, c):
    # Ô AI231
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+9)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+9)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),9.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),9.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f396(E, r, c):
    # Ô AG232
    # = IF(AH232+AI232>0,(NHAP_LIEU!$C$31+10)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+10)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+10)-4,12)+1)&" ("&IF(
    #   AH232=1,"UT3","")&IF(AH232+AI232=2,"+","")&IF(AI232=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),10.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),10.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),10.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f397(E, r, c):
    # Ô AH232
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+10)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+10)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),10.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),10.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f398(E, r, c):
    # Ô AI232
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+10)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+10)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),10.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),10.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f399(E, r, c):
    # Ô AG233
    # = IF(AH233+AI233>0,(NHAP_LIEU!$C$31+11)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+11)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+11)-4,12)+1)&" ("&IF(
    #   AH233=1,"UT3","")&IF(AH233+AI233=2,"+","")&IF(AI233=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),11.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),11.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),11.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f400(E, r, c):
    # Ô AH233
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+11)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+11)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),11.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),11.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f401(E, r, c):
    # Ô AI233
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+11)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+11)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),11.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),11.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f402(E, r, c):
    # Ô AG234
    # = IF(AH234+AI234>0,(NHAP_LIEU!$C$31+12)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+12)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+12)-4,12)+1)&" ("&IF(
    #   AH234=1,"UT3","")&IF(AH234+AI234=2,"+","")&IF(AI234=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),12.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),12.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),12.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f403(E, r, c):
    # Ô AH234
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+12)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+12)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),12.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),12.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f404(E, r, c):
    # Ô AI234
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+12)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+12)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),12.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),12.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f405(E, r, c):
    # Ô AK234
    # = IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6-TRUNG_GIAN!$D$6+12,12)=6),"• "&INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) – "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&
    #   " (tháng)"&" xung: "&IF(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),"cùng Thổ – sách mục 7: dễ bệnh tì, vị, th
    #   ận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)",INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRU
    #   NG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)),5)+1)&" khắc "&INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CS
    #   DL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUN
    #   G_GIAN!$C$6+1)),5)+1)&" → thông quan bằng "&INDEX(HANH_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_
    #   GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1))+1,5),5)+1)&" (nhóm "&INDEX(NHOM_TEN,MOD(MOD(IF(MOD(I
    #   NDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$
    #   29,TRUNG_GIAN!$D$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&": "&INDEX({"bạn bè, cộng sự, hợp tác","nói ra, viết ra, sáng tạo, chia sẻ","việc thực tế, chi tiê
    #   u hợp lý","kỷ luật, nề nếp, trách nhiệm","học tập, nghỉ ngơi, người đỡ đầu"},MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!
    #   $C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&")"&I
    #   F(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(
    #   CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1))+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," – hành này là kỵ thần, chỉ dùng ở mức vừa phải",
    #   ""))&CHAR(10),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0))),' (năm) – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' (tháng)'),' xung: '),('cùng Thổ – sách mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)' if X.truthy(X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))), 5.0),1.0)),' khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0))), 5.0),1.0))),' → thông quan bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))),1.0), 5.0), 5.0),1.0))),' (nhóm '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),': '),X.INDEX(X.Arr([['bạn bè, cộng sự, hợp tác', 'nói ra, viết ra, sáng tạo, chia sẻ', 'việc thực tế, chi tiêu hợp lý', 'kỷ luật, nề nếp, trách nhiệm', 'học tập, nghỉ ngơi, người đỡ đầu']]), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),')'),(' – hành này là kỵ thần, chỉ dùng ở mức vừa phải' if X.truthy(X.eq(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),X.CHAR(10.0)) if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(X.sub(E.v(23,6,3),E.v(23,6,4)),12.0), 12.0),6.0))) else '')

def f406(E, r, c):
    # Ô AG235
    # = IF(AH235+AI235>0,(NHAP_LIEU!$C$31+13)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+13)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+13)-4,12)+1)&" ("&IF(
    #   AH235=1,"UT3","")&IF(AH235+AI235=2,"+","")&IF(AI235=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),13.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),13.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),13.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f407(E, r, c):
    # Ô AH235
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+13)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+13)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),13.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),13.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f408(E, r, c):
    # Ô AI235
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+13)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+13)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),13.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),13.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f409(E, r, c):
    # Ô AK235
    # = IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$C$6-TRUNG_GIAN!$E$6+12,12)=6),"• "&INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) – "&INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&
    #   " (ngày)"&" xung: "&IF(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),"cùng Thổ – sách mục 7: dễ bệnh tì, vị, thậ
    #   n; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)",INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUN
    #   G_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)),5)+1)&" khắc "&INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSD
    #   L!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG
    #   _GIAN!$C$6+1)),5)+1)&" → thông quan bằng "&INDEX(HANH_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_G
    #   IAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1))+1,5),5)+1)&" (nhóm "&INDEX(NHOM_TEN,MOD(MOD(IF(MOD(IN
    #   DEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$2
    #   9,TRUNG_GIAN!$E$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&": "&INDEX({"bạn bè, cộng sự, hợp tác","nói ra, viết ra, sáng tạo, chia sẻ","việc thực tế, chi tiêu
    #    hợp lý","kỷ luật, nề nếp, trách nhiệm","học tập, nghỉ ngơi, người đỡ đầu"},MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$
    #   C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&")"&IF
    #   (MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(C
    #   SDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1))+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," – hành này là kỵ thần, chỉ dùng ở mức vừa phải","
    #   "))&CHAR(10),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0))),' (năm) – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0))),' (ngày)'),' xung: '),('cùng Thổ – sách mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)' if X.truthy(X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))), 5.0),1.0)),' khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0))), 5.0),1.0))),' → thông quan bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0), 5.0),1.0))),' (nhóm '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),': '),X.INDEX(X.Arr([['bạn bè, cộng sự, hợp tác', 'nói ra, viết ra, sáng tạo, chia sẻ', 'việc thực tế, chi tiêu hợp lý', 'kỷ luật, nề nếp, trách nhiệm', 'học tập, nghỉ ngơi, người đỡ đầu']]), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),')'),(' – hành này là kỵ thần, chỉ dùng ở mức vừa phải' if X.truthy(X.eq(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),X.CHAR(10.0)) if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(X.sub(E.v(23,6,3),E.v(23,6,5)),12.0), 12.0),6.0))) else '')

def f410(E, r, c):
    # Ô AG236
    # = IF(AH236+AI236>0,(NHAP_LIEU!$C$31+14)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+14)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+14)-4,12)+1)&" ("&IF(
    #   AH236=1,"UT3","")&IF(AH236+AI236=2,"+","")&IF(AI236=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),14.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),14.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),14.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f411(E, r, c):
    # Ô AH236
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+14)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+14)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),14.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),14.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f412(E, r, c):
    # Ô AI236
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+14)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+14)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),14.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),14.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f413(E, r, c):
    # Ô AK236
    # = IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$C$6-TRUNG_GIAN!$F$6+12,12)=6),"• "&INDEX(CHI_TEN,TRUNG_GIAN!$C$6+1)&" (năm) – "&INDEX
    #   (CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ)"&" xung: "&IF(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),"cùng Thổ – sách
    #    mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)",INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=
    #   INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1)),5)+1)&" khắc "&INDEX(HA
    #   NH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),I
    #   NDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)),5)+1)&" → thông quan bằng "&INDEX(HANH_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=IND
    #   EX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5),5)+1)&" (nhóm "&INDEX
    #   (NHOM_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C
    #   $6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&": "&INDEX({"bạn bè, cộng sự, hợp tác","nói ra, viết ra, sáng tạo, chia 
    #   sẻ","việc thực tế, chi tiêu hợp lý","kỷ luật, nề nếp, trách nhiệm","học tập, nghỉ ngơi, người đỡ đầu"},MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIA
    #   N!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)-TRU
    #   NG_GIAN!$E$7+5,5)+1)&")"&IF(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$C$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29
    #   ,TRUNG_GIAN!$C$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," – hành này là kỵ thần,
    #    chỉ dùng ở mức vừa phải",""))&CHAR(10),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,3),1.0))),' (năm) – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0))),' (giờ)'),' xung: '),('cùng Thổ – sách mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)' if X.truthy(X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))), 5.0),1.0)),' khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0))), 5.0),1.0))),' → thông quan bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0), 5.0),1.0))),' (nhóm '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),': '),X.INDEX(X.Arr([['bạn bè, cộng sự, hợp tác', 'nói ra, viết ra, sáng tạo, chia sẻ', 'việc thực tế, chi tiêu hợp lý', 'kỷ luật, nề nếp, trách nhiệm', 'học tập, nghỉ ngơi, người đỡ đầu']]), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),')'),(' – hành này là kỵ thần, chỉ dùng ở mức vừa phải' if X.truthy(X.eq(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,3),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),X.CHAR(10.0)) if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(X.sub(E.v(23,6,3),E.v(23,6,6)),12.0), 12.0),6.0))) else '')

def f414(E, r, c):
    # Ô AG237
    # = IF(AH237+AI237>0,(NHAP_LIEU!$C$31+15)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+15)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+15)-4,12)+1)&" ("&IF(
    #   AH237=1,"UT3","")&IF(AH237+AI237=2,"+","")&IF(AI237=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),15.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),15.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),15.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f415(E, r, c):
    # Ô AH237
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+15)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+15)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),15.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),15.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f416(E, r, c):
    # Ô AI237
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+15)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+15)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),15.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),15.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f417(E, r, c):
    # Ô AK237
    # = IF(AND(TRUE(),TRUE(),MOD(TRUNG_GIAN!$D$6-TRUNG_GIAN!$E$6+12,12)=6),"• "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) – "&INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1
    #   )&" (ngày)"&" xung: "&IF(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),"cùng Thổ – sách mục 7: dễ bệnh tì, vị, t
    #   hận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)",INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TR
    #   UNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)),5)+1)&" khắc "&INDEX(HANH_TEN,MOD(IF(MOD(INDEX(C
    #   SDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRU
    #   NG_GIAN!$D$6+1)),5)+1)&" → thông quan bằng "&INDEX(HANH_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG
    #   _GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1))+1,5),5)+1)&" (nhóm "&INDEX(NHOM_TEN,MOD(MOD(IF(MOD(
    #   INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C
    #   $29,TRUNG_GIAN!$E$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&": "&INDEX({"bạn bè, cộng sự, hợp tác","nói ra, viết ra, sáng tạo, chia sẻ","việc thực tế, chi ti
    #   êu hợp lý","kỷ luật, nề nếp, trách nhiệm","học tập, nghỉ ngơi, người đỡ đầu"},MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL
    #   !$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&")"&
    #   IF(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX
    #   (CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1))+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," – hành này là kỵ thần, chỉ dùng ở mức vừa phải"
    #   ,""))&CHAR(10),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' (tháng) – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0))),' (ngày)'),' xung: '),('cùng Thổ – sách mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)' if X.truthy(X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))), 5.0),1.0)),' khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))), 5.0),1.0))),' → thông quan bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0), 5.0),1.0))),' (nhóm '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),': '),X.INDEX(X.Arr([['bạn bè, cộng sự, hợp tác', 'nói ra, viết ra, sáng tạo, chia sẻ', 'việc thực tế, chi tiêu hợp lý', 'kỷ luật, nề nếp, trách nhiệm', 'học tập, nghỉ ngơi, người đỡ đầu']]), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),')'),(' – hành này là kỵ thần, chỉ dùng ở mức vừa phải' if X.truthy(X.eq(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),X.CHAR(10.0)) if X.truthy(X.AND(X.TRUE(), X.TRUE(), X.eq(X.MOD(X.add(X.sub(E.v(23,6,4),E.v(23,6,5)),12.0), 12.0),6.0))) else '')

def f418(E, r, c):
    # Ô AG238
    # = IF(AH238+AI238>0,(NHAP_LIEU!$C$31+16)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+16)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+16)-4,12)+1)&" ("&IF(
    #   AH238=1,"UT3","")&IF(AH238+AI238=2,"+","")&IF(AI238=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),16.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),16.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),16.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f419(E, r, c):
    # Ô AH238
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+16)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+16)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),16.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),16.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f420(E, r, c):
    # Ô AI238
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+16)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+16)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),16.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),16.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f421(E, r, c):
    # Ô AK238
    # = IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$D$6-TRUNG_GIAN!$F$6+12,12)=6),"• "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" (tháng) – "&IND
    #   EX(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ)"&" xung: "&IF(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),"cùng Thổ – sá
    #   ch mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)",INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5
    #   )=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1)),5)+1)&" khắc "&INDEX(
    #   HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1)
    #   ,INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)),5)+1)&" → thông quan bằng "&INDEX(HANH_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=I
    #   NDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5),5)+1)&" (nhóm "&IND
    #   EX(NHOM_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!
    #   $D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&": "&INDEX({"bạn bè, cộng sự, hợp tác","nói ra, viết ra, sáng tạo, chi
    #   a sẻ","việc thực tế, chi tiêu hợp lý","kỷ luật, nề nếp, trách nhiệm","học tập, nghỉ ngơi, người đỡ đầu"},MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_G
    #   IAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)-T
    #   RUNG_GIAN!$E$7+5,5)+1)&")"&IF(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$D$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$
    #   29,TRUNG_GIAN!$D$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," – hành này là kỵ thầ
    #   n, chỉ dùng ở mức vừa phải",""))&CHAR(10),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,4),1.0))),' (tháng) – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0))),' (giờ)'),' xung: '),('cùng Thổ – sách mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)' if X.truthy(X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))), 5.0),1.0)),' khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0))), 5.0),1.0))),' → thông quan bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0), 5.0),1.0))),' (nhóm '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),': '),X.INDEX(X.Arr([['bạn bè, cộng sự, hợp tác', 'nói ra, viết ra, sáng tạo, chia sẻ', 'việc thực tế, chi tiêu hợp lý', 'kỷ luật, nề nếp, trách nhiệm', 'học tập, nghỉ ngơi, người đỡ đầu']]), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),')'),(' – hành này là kỵ thần, chỉ dùng ở mức vừa phải' if X.truthy(X.eq(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,4),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),X.CHAR(10.0)) if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(X.sub(E.v(23,6,4),E.v(23,6,6)),12.0), 12.0),6.0))) else '')

def f422(E, r, c):
    # Ô AG239
    # = IF(AH239+AI239>0,(NHAP_LIEU!$C$31+17)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+17)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+17)-4,12)+1)&" ("&IF(
    #   AH239=1,"UT3","")&IF(AH239+AI239=2,"+","")&IF(AI239=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),17.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),17.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),17.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f423(E, r, c):
    # Ô AH239
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+17)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+17)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),17.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),17.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f424(E, r, c):
    # Ô AI239
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+17)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+17)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),17.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),17.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f425(E, r, c):
    # Ô AK239
    # = IF(AND(TRUE(),NOT(NHAP_LIEU!$C$9="Không rõ giờ"),MOD(TRUNG_GIAN!$E$6-TRUNG_GIAN!$F$6+12,12)=6),"• "&INDEX(CHI_TEN,TRUNG_GIAN!$E$6+1)&" (ngày) – "&INDE
    #   X(CHI_TEN,TRUNG_GIAN!$F$6+1)&" (giờ)"&" xung: "&IF(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),"cùng Thổ – sác
    #   h mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)",INDEX(HANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)+2,5)
    #   =INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1)),5)+1)&" khắc "&INDEX(H
    #   ANH_TEN,MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),
    #   INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)),5)+1)&" → thông quan bằng "&INDEX(HANH_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)+2,5)=IN
    #   DEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5),5)+1)&" (nhóm "&INDE
    #   X(NHOM_TEN,MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$
    #   E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)-TRUNG_GIAN!$E$7+5,5)+1)&": "&INDEX({"bạn bè, cộng sự, hợp tác","nói ra, viết ra, sáng tạo, chia
    #    sẻ","việc thực tế, chi tiêu hợp lý","kỷ luật, nề nếp, trách nhiệm","học tập, nghỉ ngơi, người đỡ đầu"},MOD(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GI
    #   AN!$E$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)-TR
    #   UNG_GIAN!$E$7+5,5)+1)&")"&IF(MOD(IF(MOD(INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$E$6+1)+2,5)=INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1),INDEX(CSDL!$C$18:$C$2
    #   9,TRUNG_GIAN!$E$6+1),INDEX(CSDL!$C$18:$C$29,TRUNG_GIAN!$F$6+1))+1,5)=IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)," – hành này là kỵ thần
    #   , chỉ dùng ở mức vừa phải",""))&CHAR(10),"")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('• ',X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,5),1.0))),' (ngày) – '),X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(23,6,6),1.0))),' (giờ)'),' xung: '),('cùng Thổ – sách mục 7: dễ bệnh tì, vị, thận; dùng Thủy nhuận hoặc Mộc giảm nhẹ (P2 tr.647)' if X.truthy(X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))), 5.0),1.0)),' khắc '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0))), 5.0),1.0))),' → thông quan bằng '),X.INDEX(E.rg(24,33,1,37,1), X.add(X.MOD(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0), 5.0),1.0))),' (nhóm '),X.INDEX(E.rg(24,54,1,58,1), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),': '),X.INDEX(X.Arr([['bạn bè, cộng sự, hợp tác', 'nói ra, viết ra, sáng tạo, chia sẻ', 'việc thực tế, chi tiêu hợp lý', 'kỷ luật, nề nếp, trách nhiệm', 'học tập, nghỉ ngơi, người đỡ đầu']]), X.add(X.MOD(X.add(X.sub(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),E.v(23,7,5)),5.0), 5.0),1.0))),')'),(' – hành này là kỵ thần, chỉ dùng ở mức vừa phải' if X.truthy(X.eq(X.MOD(X.add((X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)) if X.truthy(X.eq(X.MOD(X.add(X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,5),1.0)),2.0), 5.0),X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0)))) else X.INDEX(E.rg(24,18,3,29,3), X.add(E.v(23,6,6),1.0))),1.0), 5.0),(E.v(8,111,4) if X.truthy(X.eq(E.v(7,33,3),2.0)) else E.v(8,111,3)))) else '')))),X.CHAR(10.0)) if X.truthy(X.AND(X.TRUE(), X.NOT(X.eq(E.v(2,9,3),'Không rõ giờ')), X.eq(X.MOD(X.add(X.sub(E.v(23,6,5),E.v(23,6,6)),12.0), 12.0),6.0))) else '')

def f426(E, r, c):
    # Ô AG240
    # = IF(AH240+AI240>0,(NHAP_LIEU!$C$31+18)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+18)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+18)-4,12)+1)&" ("&IF(
    #   AH240=1,"UT3","")&IF(AH240+AI240=2,"+","")&IF(AI240=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),18.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),18.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),18.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f427(E, r, c):
    # Ô AH240
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+18)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+18)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),18.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),18.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f428(E, r, c):
    # Ô AI240
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+18)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+18)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),18.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),18.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f429(E, r, c):
    # Ô AG241
    # = IF(AH241+AI241>0,(NHAP_LIEU!$C$31+19)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+19)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+19)-4,12)+1)&" ("&IF(
    #   AH241=1,"UT3","")&IF(AH241+AI241=2,"+","")&IF(AI241=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),19.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),19.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),19.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f430(E, r, c):
    # Ô AH241
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+19)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+19)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),19.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),19.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f431(E, r, c):
    # Ô AI241
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+19)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+19)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),19.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),19.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f432(E, r, c):
    # Ô AG242
    # = IF(AH242+AI242>0,(NHAP_LIEU!$C$31+20)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+20)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+20)-4,12)+1)&" ("&IF(
    #   AH242=1,"UT3","")&IF(AH242+AI242=2,"+","")&IF(AI242=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),20.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),20.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),20.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f433(E, r, c):
    # Ô AH242
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+20)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+20)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),20.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),20.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f434(E, r, c):
    # Ô AI242
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+20)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+20)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),20.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),20.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f435(E, r, c):
    # Ô AG243
    # = IF(AH243+AI243>0,(NHAP_LIEU!$C$31+21)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+21)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+21)-4,12)+1)&" ("&IF(
    #   AH243=1,"UT3","")&IF(AH243+AI243=2,"+","")&IF(AI243=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),21.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),21.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),21.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f436(E, r, c):
    # Ô AH243
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+21)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+21)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),21.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),21.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f437(E, r, c):
    # Ô AI243
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+21)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+21)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),21.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),21.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f438(E, r, c):
    # Ô AG244
    # = IF(AH244+AI244>0,(NHAP_LIEU!$C$31+22)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+22)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+22)-4,12)+1)&" ("&IF(
    #   AH244=1,"UT3","")&IF(AH244+AI244=2,"+","")&IF(AI244=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),22.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),22.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),22.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f439(E, r, c):
    # Ô AH244
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+22)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+22)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),22.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),22.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f440(E, r, c):
    # Ô AI244
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+22)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+22)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),22.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),22.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f441(E, r, c):
    # Ô AG245
    # = IF(AH245+AI245>0,(NHAP_LIEU!$C$31+23)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+23)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+23)-4,12)+1)&" ("&IF(
    #   AH245=1,"UT3","")&IF(AH245+AI245=2,"+","")&IF(AI245=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),23.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),23.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),23.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f442(E, r, c):
    # Ô AH245
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+23)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+23)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),23.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),23.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f443(E, r, c):
    # Ô AI245
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+23)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+23)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),23.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),23.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f444(E, r, c):
    # Ô AG246
    # = IF(AH246+AI246>0,(NHAP_LIEU!$C$31+24)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+24)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+24)-4,12)+1)&" ("&IF(
    #   AH246=1,"UT3","")&IF(AH246+AI246=2,"+","")&IF(AI246=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),24.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),24.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),24.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f445(E, r, c):
    # Ô AH246
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+24)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+24)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),24.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),24.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f446(E, r, c):
    # Ô AI246
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+24)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+24)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),24.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),24.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f447(E, r, c):
    # Ô AG247
    # = IF(AH247+AI247>0,(NHAP_LIEU!$C$31+25)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+25)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+25)-4,12)+1)&" ("&IF(
    #   AH247=1,"UT3","")&IF(AH247+AI247=2,"+","")&IF(AI247=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),25.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),25.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),25.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f448(E, r, c):
    # Ô AH247
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+25)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+25)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),25.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),25.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f449(E, r, c):
    # Ô AI247
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+25)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+25)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),25.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),25.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f450(E, r, c):
    # Ô AG248
    # = IF(AH248+AI248>0,(NHAP_LIEU!$C$31+26)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+26)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+26)-4,12)+1)&" ("&IF(
    #   AH248=1,"UT3","")&IF(AH248+AI248=2,"+","")&IF(AI248=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),26.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),26.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),26.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f451(E, r, c):
    # Ô AH248
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+26)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+26)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),26.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),26.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f452(E, r, c):
    # Ô AI248
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+26)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+26)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),26.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),26.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f453(E, r, c):
    # Ô AG249
    # = IF(AH249+AI249>0,(NHAP_LIEU!$C$31+27)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+27)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+27)-4,12)+1)&" ("&IF(
    #   AH249=1,"UT3","")&IF(AH249+AI249=2,"+","")&IF(AI249=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),27.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),27.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),27.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f454(E, r, c):
    # Ô AH249
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+27)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+27)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),27.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),27.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f455(E, r, c):
    # Ô AI249
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+27)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+27)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),27.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),27.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f456(E, r, c):
    # Ô AG250
    # = IF(AH250+AI250>0,(NHAP_LIEU!$C$31+28)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+28)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+28)-4,12)+1)&" ("&IF(
    #   AH250=1,"UT3","")&IF(AH250+AI250=2,"+","")&IF(AI250=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),28.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),28.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),28.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f457(E, r, c):
    # Ô AH250
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+28)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+28)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),28.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),28.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f458(E, r, c):
    # Ô AI250
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+28)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+28)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),28.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),28.0),4.0), 12.0),1.0)),4.0))) else 0.0)

def f459(E, r, c):
    # Ô AG251
    # = IF(AH251+AI251>0,(NHAP_LIEU!$C$31+29)&" "&INDEX(CAN_TEN,MOD((NHAP_LIEU!$C$31+29)-4,10)+1)&" "&INDEX(CHI_TEN,MOD((NHAP_LIEU!$C$31+29)-4,12)+1)&" ("&IF(
    #   AH251=1,"UT3","")&IF(AH251+AI251=2,"+","")&IF(AI251=1,"UT1","")&"); ","")
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.add(E.v(2,31,3),29.0),' '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),29.0),4.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),29.0),4.0), 12.0),1.0))),' ('),('UT3' if X.truthy(X.eq(E.v(3,r,c+1),1.0)) else '')),('+' if X.truthy(X.eq(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),2.0)) else '')),('UT1' if X.truthy(X.eq(E.v(3,r,c+2),1.0)) else '')),'); ') if X.truthy(X.gt(X.add(E.v(3,r,c+1),E.v(3,r,c+2)),0.0)) else '')

def f460(E, r, c):
    # Ô AH251
    # = IF(AND(INT(MOD((NHAP_LIEU!$C$31+29)-4,10)/2)=$AG$210,MOD(MOD((NHAP_LIEU!$C$31+29)-4,10),2)=$AG$211,$AG$212=1),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),29.0),4.0), 10.0),2.0)),E.v(3,210,33)), X.eq(X.MOD(X.MOD(X.sub(X.add(E.v(2,31,3),29.0),4.0), 10.0), 2.0),E.v(3,211,33)), X.eq(E.v(3,212,33),1.0))) else 0.0)

def f461(E, r, c):
    # Ô AI251
    # = IF(AND($AG$213=1,INT(MOD((NHAP_LIEU!$C$31+29)-4,10)/2)=4,INDEX(CSDL!$C$18:$C$29,MOD((NHAP_LIEU!$C$31+29)-4,12)+1)=4),1,0)
    return (1.0 if X.truthy(X.AND(X.eq(E.v(3,213,33),1.0), X.eq(X.INT(X.div(X.MOD(X.sub(X.add(E.v(2,31,3),29.0),4.0), 10.0),2.0)),4.0), X.eq(X.INDEX(E.rg(24,18,3,29,3), X.add(X.MOD(X.sub(X.add(E.v(2,31,3),29.0),4.0), 12.0),1.0)),4.0))) else 0.0)

CELLS = {
    'B2': (f0, 2, 2),
    'C3': (f1, 3, 3),
    'D3': (f1, 3, 4),
    'E3': (f1, 3, 5),
    'F3': (f1, 3, 6),
    'G3': (f1, 3, 7),
    'C4': (f2, 4, 3),
    'D4': (f1, 4, 4),
    'E4': (f1, 4, 5),
    'F4': (f1, 4, 6),
    'G4': (f1, 4, 7),
    'C5': (f3, 5, 3),
    'D5': (f1, 5, 4),
    'E5': (f1, 5, 5),
    'F5': (f1, 5, 6),
    'G5': (f1, 5, 7),
    'C6': (f4, 6, 3),
    'D6': (f1, 6, 4),
    'E6': (f1, 6, 5),
    'F6': (f1, 6, 6),
    'G6': (f1, 6, 7),
    'C7': (f5, 7, 3),
    'D7': (f1, 7, 4),
    'E7': (f1, 7, 5),
    'F7': (f1, 7, 6),
    'G7': (f1, 7, 7),
    'C8': (f6, 8, 3),
    'D8': (f1, 8, 4),
    'E8': (f1, 8, 5),
    'F8': (f1, 8, 6),
    'G8': (f1, 8, 7),
    'C9': (f7, 9, 3),
    'D9': (f1, 9, 4),
    'E9': (f1, 9, 5),
    'F9': (f1, 9, 6),
    'G9': (f1, 9, 7),
    'C10': (f8, 10, 3),
    'D10': (f1, 10, 4),
    'E10': (f1, 10, 5),
    'F10': (f1, 10, 6),
    'G10': (f1, 10, 7),
    'C11': (f9, 11, 3),
    'D11': (f1, 11, 4),
    'E11': (f1, 11, 5),
    'F11': (f1, 11, 6),
    'G11': (f1, 11, 7),
    'C12': (f10, 12, 3),
    'D12': (f1, 12, 4),
    'E12': (f1, 12, 5),
    'F12': (f1, 12, 6),
    'G12': (f1, 12, 7),
    'C15': (f11, 15, 3),
    'D15': (f11, 15, 4),
    'E15': (f11, 15, 5),
    'F15': (f11, 15, 6),
    'G15': (f12, 15, 7),
    'C17': (f13, 17, 3),
    'D17': (f14, 17, 4),
    'E17': (f15, 17, 5),
    'F17': (f16, 17, 6),
    'C18': (f17, 18, 3),
    'D18': (f18, 18, 4),
    'E18': (f19, 18, 5),
    'F18': (f20, 18, 6),
    'C19': (f21, 19, 3),
    'D19': (f22, 19, 4),
    'E19': (f23, 19, 5),
    'F19': (f24, 19, 6),
    'C20': (f25, 20, 3),
    'D20': (f26, 20, 4),
    'E20': (f27, 20, 5),
    'F20': (f28, 20, 6),
    'C21': (f29, 21, 3),
    'D21': (f30, 21, 4),
    'E21': (f31, 21, 5),
    'F21': (f32, 21, 6),
    'C22': (f33, 22, 3),
    'D22': (f34, 22, 4),
    'E22': (f35, 22, 5),
    'F22': (f36, 22, 6),
    'C23': (f37, 23, 3),
    'D23': (f38, 23, 4),
    'E23': (f39, 23, 5),
    'F23': (f40, 23, 6),
    'C24': (f41, 24, 3),
    'J24': (f42, 24, 10),
    'K24': (f43, 24, 11),
    'L24': (f44, 24, 12),
    'M24': (f45, 24, 13),
    'N24': (f46, 24, 14),
    'O24': (f47, 24, 15),
    'P24': (f48, 24, 16),
    'Q24': (f49, 24, 17),
    'R24': (f50, 24, 18),
    'S24': (f51, 24, 19),
    'C29': (f52, 29, 3),
    'C31': (f53, 31, 3),
    'D31': (f54, 31, 4),
    'E31': (f55, 31, 5),
    'F31': (f56, 31, 6),
    'G31': (f57, 31, 7),
    'H31': (f58, 31, 8),
    'C32': (f59, 32, 3),
    'H32': (f60, 32, 8),
    'C33': (f61, 33, 3),
    'H33': (f62, 33, 8),
    'C34': (f63, 34, 3),
    'H34': (f64, 34, 8),
    'C35': (f65, 35, 3),
    'D35': (f66, 35, 4),
    'E35': (f67, 35, 5),
    'F35': (f68, 35, 6),
    'G35': (f1, 35, 7),
    'H35': (f69, 35, 8),
    'C36': (f70, 36, 3),
    'D36': (f71, 36, 4),
    'E36': (f72, 36, 5),
    'F36': (f73, 36, 6),
    'G36': (f74, 36, 7),
    'H36': (f75, 36, 8),
    'C37': (f76, 37, 3),
    'D37': (f77, 37, 4),
    'E37': (f78, 37, 5),
    'F37': (f79, 37, 6),
    'G37': (f80, 37, 7),
    'H37': (f81, 37, 8),
    'C38': (f82, 38, 3),
    'D38': (f83, 38, 4),
    'E38': (f84, 38, 5),
    'F38': (f85, 38, 6),
    'G38': (f86, 38, 7),
    'H38': (f87, 38, 8),
    'C39': (f88, 39, 3),
    'H39': (f89, 39, 8),
    'C40': (f90, 40, 3),
    'F40': (f91, 40, 6),
    'G40': (f92, 40, 7),
    'H40': (f93, 40, 8),
    'C41': (f94, 41, 3),
    'F41': (f95, 41, 6),
    'G41': (f96, 41, 7),
    'H41': (f97, 41, 8),
    'C42': (f98, 42, 3),
    'E42': (f99, 42, 5),
    'F42': (f100, 42, 6),
    'G42': (f101, 42, 7),
    'H42': (f102, 42, 8),
    'C43': (f103, 43, 3),
    'H43': (f104, 43, 8),
    'C45': (f105, 45, 3),
    'H45': (f106, 45, 8),
    'C47': (f107, 47, 3),
    'H47': (f108, 47, 8),
    'C48': (f109, 48, 3),
    'H48': (f110, 48, 8),
    'C49': (f111, 49, 3),
    'H49': (f112, 49, 8),
    'C50': (f113, 50, 3),
    'H50': (f114, 50, 8),
    'C51': (f115, 51, 3),
    'H51': (f116, 51, 8),
    'C52': (f117, 52, 3),
    'C55': (f118, 55, 3),
    'D55': (f119, 55, 4),
    'E55': (f120, 55, 5),
    'F55': (f121, 55, 6),
    'C56': (f122, 56, 3),
    'D56': (f123, 56, 4),
    'E56': (f124, 56, 5),
    'F56': (f125, 56, 6),
    'H56': (f126, 56, 8),
    'C57': (f127, 57, 3),
    'D57': (f128, 57, 4),
    'E57': (f129, 57, 5),
    'F57': (f130, 57, 6),
    'H57': (f131, 57, 8),
    'C58': (f132, 58, 3),
    'H58': (f133, 58, 8),
    'C59': (f134, 59, 3),
    'H59': (f135, 59, 8),
    'C60': (f136, 60, 3),
    'H60': (f137, 60, 8),
    'C62': (f138, 62, 3),
    'H62': (f139, 62, 8),
    'C63': (f140, 63, 3),
    'H63': (f141, 63, 8),
    'C64': (f142, 64, 3),
    'H64': (f143, 64, 8),
    'C65': (f144, 65, 3),
    'C68': (f145, 68, 3),
    'C70': (f146, 70, 3),
    'H70': (f147, 70, 8),
    'C71': (f148, 71, 3),
    'H71': (f149, 71, 8),
    'C72': (f150, 72, 3),
    'H72': (f151, 72, 8),
    'C73': (f152, 73, 3),
    'H73': (f153, 73, 8),
    'C74': (f154, 74, 3),
    'H74': (f155, 74, 8),
    'C75': (f156, 75, 3),
    'H75': (f157, 75, 8),
    'C76': (f158, 76, 3),
    'C79': (f159, 79, 3),
    'C81': (f160, 81, 3),
    'H81': (f161, 81, 8),
    'C82': (f162, 82, 3),
    'H82': (f163, 82, 8),
    'C83': (f164, 83, 3),
    'H83': (f165, 83, 8),
    'C84': (f166, 84, 3),
    'H84': (f167, 84, 8),
    'C85': (f168, 85, 3),
    'H85': (f169, 85, 8),
    'C86': (f170, 86, 3),
    'H86': (f171, 86, 8),
    'B87': (f172, 87, 2),
    'C87': (f173, 87, 3),
    'H87': (f174, 87, 8),
    'AB87': (f175, 87, 28),
    'AC87': (f176, 87, 29),
    'AD87': (f177, 87, 30),
    'AE87': (f178, 87, 31),
    'C88': (f179, 88, 3),
    'H88': (f180, 88, 8),
    'C89': (f181, 89, 3),
    'H89': (f182, 89, 8),
    'C90': (f183, 90, 3),
    'C93': (f184, 93, 3),
    'C95': (f185, 95, 3),
    'H95': (f186, 95, 8),
    'C96': (f187, 96, 3),
    'H96': (f188, 96, 8),
    'C97': (f189, 97, 3),
    'H97': (f190, 97, 8),
    'C98': (f191, 98, 3),
    'H98': (f192, 98, 8),
    'B99': (f193, 99, 2),
    'C99': (f194, 99, 3),
    'H99': (f195, 99, 8),
    'C100': (f196, 100, 3),
    'H100': (f197, 100, 8),
    'C101': (f198, 101, 3),
    'C102': (f199, 102, 3),
    'C103': (f200, 103, 3),
    'C104': (f201, 104, 3),
    'C105': (f202, 105, 3),
    'C108': (f203, 108, 3),
    'B110': (f204, 110, 2),
    'C110': (f205, 110, 3),
    'H110': (f206, 110, 8),
    'B111': (f204, 111, 2),
    'C111': (f207, 111, 3),
    'H111': (f208, 111, 8),
    'B112': (f204, 112, 2),
    'C112': (f209, 112, 3),
    'H112': (f210, 112, 8),
    'B113': (f204, 113, 2),
    'C113': (f211, 113, 3),
    'H113': (f212, 113, 8),
    'B114': (f204, 114, 2),
    'C114': (f213, 114, 3),
    'H114': (f214, 114, 8),
    'C115': (f215, 115, 3),
    'H115': (f216, 115, 8),
    'C117': (f217, 117, 3),
    'H117': (f218, 117, 8),
    'C118': (f219, 118, 3),
    'H118': (f220, 118, 8),
    'C119': (f221, 119, 3),
    'H119': (f222, 119, 8),
    'C120': (f223, 120, 3),
    'H120': (f224, 120, 8),
    'C121': (f225, 121, 3),
    'H121': (f226, 121, 8),
    'C122': (f227, 122, 3),
    'H122': (f228, 122, 8),
    'C123': (f229, 123, 3),
    'H123': (f230, 123, 8),
    'C124': (f231, 124, 3),
    'H124': (f232, 124, 8),
    'C125': (f233, 125, 3),
    'C126': (f234, 126, 3),
    'C127': (f235, 127, 3),
    'C128': (f236, 128, 3),
    'C129': (f237, 129, 3),
    'C130': (f238, 130, 3),
    'C131': (f239, 131, 3),
    'C132': (f240, 132, 3),
    'C135': (f241, 135, 3),
    'C137': (f242, 137, 3),
    'H137': (f243, 137, 8),
    'C138': (f244, 138, 3),
    'H138': (f245, 138, 8),
    'C139': (f246, 139, 3),
    'H139': (f247, 139, 8),
    'C140': (f248, 140, 3),
    'H140': (f249, 140, 8),
    'C141': (f250, 141, 3),
    'H141': (f251, 141, 8),
    'C142': (f252, 142, 3),
    'H142': (f253, 142, 8),
    'B143': (f254, 143, 2),
    'C143': (f255, 143, 3),
    'H143': (f256, 143, 8),
    'C144': (f257, 144, 3),
    'H144': (f258, 144, 8),
    'C145': (f259, 145, 3),
    'H145': (f260, 145, 8),
    'C147': (f261, 147, 3),
    'C148': (f262, 148, 3),
    'C149': (f263, 149, 3),
    'AG149': (f264, 149, 33),
    'C150': (f265, 150, 3),
    'AG150': (f266, 150, 33),
    'C151': (f267, 151, 3),
    'AG151': (f268, 151, 33),
    'C152': (f269, 152, 3),
    'AG152': (f270, 152, 33),
    'C153': (f271, 153, 3),
    'C154': (f272, 154, 3),
    'C155': (f273, 155, 3),
    'C156': (f274, 156, 3),
    'C157': (f274, 157, 3),
    'C160': (f275, 160, 3),
    'C162': (f276, 162, 3),
    'H162': (f277, 162, 8),
    'C163': (f278, 163, 3),
    'H163': (f279, 163, 8),
    'B165': (f280, 165, 2),
    'C165': (f281, 165, 3),
    'D165': (f282, 165, 4),
    'E165': (f283, 165, 5),
    'G165': (f284, 165, 7),
    'H165': (f285, 165, 8),
    'B166': (f280, 166, 2),
    'C166': (f281, 166, 3),
    'D166': (f282, 166, 4),
    'E166': (f283, 166, 5),
    'G166': (f284, 166, 7),
    'H166': (f286, 166, 8),
    'B167': (f280, 167, 2),
    'C167': (f281, 167, 3),
    'D167': (f282, 167, 4),
    'E167': (f283, 167, 5),
    'G167': (f284, 167, 7),
    'H167': (f287, 167, 8),
    'B168': (f280, 168, 2),
    'C168': (f281, 168, 3),
    'D168': (f282, 168, 4),
    'E168': (f283, 168, 5),
    'G168': (f284, 168, 7),
    'H168': (f288, 168, 8),
    'B169': (f280, 169, 2),
    'C169': (f281, 169, 3),
    'D169': (f282, 169, 4),
    'E169': (f283, 169, 5),
    'G169': (f284, 169, 7),
    'H169': (f289, 169, 8),
    'B170': (f280, 170, 2),
    'C170': (f281, 170, 3),
    'D170': (f282, 170, 4),
    'E170': (f283, 170, 5),
    'G170': (f284, 170, 7),
    'H170': (f290, 170, 8),
    'C171': (f291, 171, 3),
    'H171': (f292, 171, 8),
    'C172': (f293, 172, 3),
    'H172': (f294, 172, 8),
    'J172': (f295, 172, 10),
    'K172': (f296, 172, 11),
    'L172': (f297, 172, 12),
    'M172': (f298, 172, 13),
    'N172': (f299, 172, 14),
    'O172': (f300, 172, 15),
    'P172': (f301, 172, 16),
    'Q172': (f302, 172, 17),
    'T172': (f303, 172, 20),
    'U172': (f304, 172, 21),
    'V172': (f305, 172, 22),
    'W172': (f306, 172, 23),
    'X172': (f1, 172, 24),
    'Y172': (f307, 172, 25),
    'Z172': (f308, 172, 26),
    'AA172': (f309, 172, 27),
    'B173': (f310, 173, 2),
    'C173': (f311, 173, 3),
    'H173': (f312, 173, 8),
    'C174': (f313, 174, 3),
    'C175': (f314, 175, 3),
    'C176': (f315, 176, 3),
    'C179': (f316, 179, 3),
    'C181': (f317, 181, 3),
    'C182': (f318, 182, 3),
    'C183': (f319, 183, 3),
    'C184': (f320, 184, 3),
    'C185': (f321, 185, 3),
    'C186': (f322, 186, 3),
    'C187': (f323, 187, 3),
    'C188': (f324, 188, 3),
    'C189': (f325, 189, 3),
    'C190': (f326, 190, 3),
    'C193': (f327, 193, 3),
    'C195': (f328, 195, 3),
    'H195': (f329, 195, 8),
    'C202': (f330, 202, 3),
    'H202': (f331, 202, 8),
    'C203': (f332, 203, 3),
    'H203': (f333, 203, 8),
    'C204': (f334, 204, 3),
    'H204': (f335, 204, 8),
    'C205': (f336, 205, 3),
    'H205': (f337, 205, 8),
    'C206': (f338, 206, 3),
    'H206': (f339, 206, 8),
    'C207': (f340, 207, 3),
    'H207': (f341, 207, 8),
    'C208': (f342, 208, 3),
    'H208': (f343, 208, 8),
    'C209': (f344, 209, 3),
    'H209': (f345, 209, 8),
    'AG210': (f346, 210, 33),
    'AG211': (f347, 211, 33),
    'AG212': (f348, 212, 33),
    'AG213': (f349, 213, 33),
    'AG214': (f350, 214, 33),
    'AH214': (f351, 214, 34),
    'AG215': (f352, 215, 33),
    'AH215': (f353, 215, 34),
    'AG216': (f354, 216, 33),
    'AH216': (f355, 216, 34),
    'AG217': (f356, 217, 33),
    'AH217': (f357, 217, 34),
    'AG218': (f358, 218, 33),
    'AH218': (f359, 218, 34),
    'AG219': (f360, 219, 33),
    'AH219': (f361, 219, 34),
    'AG220': (f362, 220, 33),
    'AH220': (f363, 220, 34),
    'AG221': (f364, 221, 33),
    'AH221': (f365, 221, 34),
    'AG222': (f366, 222, 33),
    'AH222': (f367, 222, 34),
    'AI222': (f368, 222, 35),
    'AG223': (f369, 223, 33),
    'AH223': (f370, 223, 34),
    'AI223': (f371, 223, 35),
    'AG224': (f372, 224, 33),
    'AH224': (f373, 224, 34),
    'AI224': (f374, 224, 35),
    'AG225': (f375, 225, 33),
    'AH225': (f376, 225, 34),
    'AI225': (f377, 225, 35),
    'AG226': (f378, 226, 33),
    'AH226': (f379, 226, 34),
    'AI226': (f380, 226, 35),
    'AG227': (f381, 227, 33),
    'AH227': (f382, 227, 34),
    'AI227': (f383, 227, 35),
    'AG228': (f384, 228, 33),
    'AH228': (f385, 228, 34),
    'AI228': (f386, 228, 35),
    'AG229': (f387, 229, 33),
    'AH229': (f388, 229, 34),
    'AI229': (f389, 229, 35),
    'AG230': (f390, 230, 33),
    'AH230': (f391, 230, 34),
    'AI230': (f392, 230, 35),
    'AG231': (f393, 231, 33),
    'AH231': (f394, 231, 34),
    'AI231': (f395, 231, 35),
    'AG232': (f396, 232, 33),
    'AH232': (f397, 232, 34),
    'AI232': (f398, 232, 35),
    'AG233': (f399, 233, 33),
    'AH233': (f400, 233, 34),
    'AI233': (f401, 233, 35),
    'AG234': (f402, 234, 33),
    'AH234': (f403, 234, 34),
    'AI234': (f404, 234, 35),
    'AK234': (f405, 234, 37),
    'AG235': (f406, 235, 33),
    'AH235': (f407, 235, 34),
    'AI235': (f408, 235, 35),
    'AK235': (f409, 235, 37),
    'AG236': (f410, 236, 33),
    'AH236': (f411, 236, 34),
    'AI236': (f412, 236, 35),
    'AK236': (f413, 236, 37),
    'AG237': (f414, 237, 33),
    'AH237': (f415, 237, 34),
    'AI237': (f416, 237, 35),
    'AK237': (f417, 237, 37),
    'AG238': (f418, 238, 33),
    'AH238': (f419, 238, 34),
    'AI238': (f420, 238, 35),
    'AK238': (f421, 238, 37),
    'AG239': (f422, 239, 33),
    'AH239': (f423, 239, 34),
    'AI239': (f424, 239, 35),
    'AK239': (f425, 239, 37),
    'AG240': (f426, 240, 33),
    'AH240': (f427, 240, 34),
    'AI240': (f428, 240, 35),
    'AG241': (f429, 241, 33),
    'AH241': (f430, 241, 34),
    'AI241': (f431, 241, 35),
    'AG242': (f432, 242, 33),
    'AH242': (f433, 242, 34),
    'AI242': (f434, 242, 35),
    'AG243': (f435, 243, 33),
    'AH243': (f436, 243, 34),
    'AI243': (f437, 243, 35),
    'AG244': (f438, 244, 33),
    'AH244': (f439, 244, 34),
    'AI244': (f440, 244, 35),
    'AG245': (f441, 245, 33),
    'AH245': (f442, 245, 34),
    'AI245': (f443, 245, 35),
    'AG246': (f444, 246, 33),
    'AH246': (f445, 246, 34),
    'AI246': (f446, 246, 35),
    'AG247': (f447, 247, 33),
    'AH247': (f448, 247, 34),
    'AI247': (f449, 247, 35),
    'AG248': (f450, 248, 33),
    'AH248': (f451, 248, 34),
    'AI248': (f452, 248, 35),
    'AG249': (f453, 249, 33),
    'AH249': (f454, 249, 34),
    'AI249': (f455, 249, 35),
    'AG250': (f456, 250, 33),
    'AH250': (f457, 250, 34),
    'AI250': (f458, 250, 35),
    'AG251': (f459, 251, 33),
    'AH251': (f460, 251, 34),
    'AI251': (f461, 251, 35),
}
