# -*- coding: utf-8 -*-
"""Module sheet NHAP_LIEU – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.4.xlsx (SHA-256 b554826cb416b851894b999cc981ef29ba64e78b7dcf541338e9f62a9bae0e6c).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'NHAP_LIEU'
SI = 2

def f0(E, r, c):
    # Ô G7
    # = IF($C$13<>"Tự nhập","⚠ Đang dùng LÁ SỐ MẪU "&$C$13&" ("&C23&", "&D23&", "&E23&", "&F23&"). Ô C7:F8 không được dùng – chọn C13 = Tự nhập.",IF(AND($C$12
    #   ="Tự động",$C$37<>""),IF(AND($C$9="Biết giờ",ISNUMBER($C$26),$C$26<=31,$F$26=""),"⚠ Chưa nhập GIỜ sinh (ô F26): đang tính như 0 giờ, trụ giờ "&F23&" c
    #   ó thể sai. Nhập giờ ở F26, hoặc chọn C9 = Không rõ giờ.","✔ Đang dùng tứ trụ tính TỰ ĐỘNG từ ngày giờ sinh: "&C23&", "&D23&", "&E23&", "&F23&". Số ngà
    #   y đến tiết và tuổi khởi vận cũng tự tính (xem D28)."&IF(OR(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C7)),"Quí","Quý"),"Kỉ","Kỷ")<>C$37,SUBSTITUTE(SUBSTITUTE(
    #   PROPER(TRIM(D7)),"Quí","Quý"),"Kỉ","Kỷ")<>D$37,SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(E7)),"Quí","Quý"),"Kỉ","Kỷ")<>E$37,SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(
    #   F7)),"Quí","Quý"),"Kỉ","Kỷ")<>F$37,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão")<>C$38,SUBSTITUTE(SUBSTITUTE(SU
    #   BSTITUTE(PROPER(TRIM(D8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão")<>D$38,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(E8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"
    #   )<>E$38,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(F8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão")<>F$38)," (Can chi ở C7:F8 khác và đang không được dùng; m
    #   uốn dùng thì xóa C26:G26 hoặc chọn C12 = Không.)","")),IF(LEN(IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN,0))
    #   ,"C7 “"&C7&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(D7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN,0)),"D7 “"&D7&"” ","")&IF(ISNA(MATCH(SUBSTITU
    #   TE(SUBSTITUTE(PROPER(TRIM(E7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN,0)),"E7 “"&E7&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(F7)),"Quí","Quý"
    #   ),"Kỉ","Kỷ"),CAN_TEN,0)),"F7 “"&F7&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN
    #   ,0)),"C8 “"&C8&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(D8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"D8 “"&D8&"” ",
    #   "")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(E8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"E8 “"&E8&"” ","")&IF(ISNA(MATCH(SU
    #   BSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(F8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"F8 “"&F8&"” ",""))>0,"⚠ Tên can/chi không hợp lệ ở: "&
    #   IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN,0)),"C7 “"&C7&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER
    #   (TRIM(D7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN,0)),"D7 “"&D7&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(E7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN
    #   ,0)),"E7 “"&E7&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(F7)),"Quí","Quý"),"Kỉ","Kỷ"),CAN_TEN,0)),"F7 “"&F7&"” ","")&IF(ISNA(MATCH(SUBS
    #   TITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"C8 “"&C8&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(SU
    #   BSTITUTE(PROPER(TRIM(D8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"D8 “"&D8&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM
    #   (E8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"E8 “"&E8&"” ","")&IF(ISNA(MATCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(F8)),"Tý","Tí"),"Tỵ
    #   ","Tị"),"Mẹo","Mão"),CHI_TEN,0)),"F8 “"&F8&"” ","")&"– chọn từ danh sách (Tí, Tị, Kỷ, Quý…).","✔ Đang dùng can chi tự nhập C7:F8: "&C23&", "&D23&", "&
    #   E23&", "&F23&IF(AND($C$26="",$C$28=""),IF(ISNUMBER(NHAP_LIEU!$C$53),". Khởi vận tự tính từ tứ trụ + năm sinh C30 (dòng 43-53).",". Chưa có ngày giờ si
    #   nh (C26), số ngày (C28) hoặc năm sinh khớp trụ năm (C30): đại vận tạm khởi 1 tuổi – xem dòng 43-53."),""))))
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('⚠ Đang dùng LÁ SỐ MẪU ',E.v(2,13,3)),' ('),E.v(2,r+16,c-4)),', '),E.v(2,r+16,c-3)),', '),E.v(2,r+16,c-2)),', '),E.v(2,r+16,c-1)),'). Ô C7:F8 không được dùng – chọn C13 = Tự nhập.') if X.truthy(X.ne(E.v(2,13,3),'Tự nhập')) else ((X.cat(X.cat('⚠ Chưa nhập GIỜ sinh (ô F26): đang tính như 0 giờ, trụ giờ ',E.v(2,r+16,c-1)),' có thể sai. Nhập giờ ở F26, hoặc chọn C9 = Không rõ giờ.') if X.truthy(X.AND(X.eq(E.v(2,9,3),'Biết giờ'), X.ISNUMBER(lambda: E.v(2,26,3)), X.le(E.v(2,26,3),31.0), X.eq(E.v(2,26,6),''))) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('✔ Đang dùng tứ trụ tính TỰ ĐỘNG từ ngày giờ sinh: ',E.v(2,r+16,c-4)),', '),E.v(2,r+16,c-3)),', '),E.v(2,r+16,c-2)),', '),E.v(2,r+16,c-1)),'. Số ngày đến tiết và tuổi khởi vận cũng tự tính (xem D28).'),(' (Can chi ở C7:F8 khác và đang không được dùng; muốn dùng thì xóa C26:G26 hoặc chọn C12 = Không.)' if X.truthy(X.OR(X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-4))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'),E.v(2,37,c-4)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-3))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'),E.v(2,37,c-3)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-2))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'),E.v(2,37,c-2)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-1))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'),E.v(2,37,c-1)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-4))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'),E.v(2,38,c-4)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-3))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'),E.v(2,38,c-3)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-2))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'),E.v(2,38,c-2)), X.ne(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-1))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'),E.v(2,38,c-1)))) else ''))) if X.truthy(X.AND(X.eq(E.v(2,12,3),'Tự động'), X.ne(E.v(2,37,3),''))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('⚠ Tên can/chi không hợp lệ ở: ',(X.cat(X.cat('C7 “',E.v(2,r,c-4)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-4))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('D7 “',E.v(2,r,c-3)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-3))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('E7 “',E.v(2,r,c-2)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-2))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('F7 “',E.v(2,r,c-1)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-1))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('C8 “',E.v(2,r+1,c-4)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-4))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),(X.cat(X.cat('D8 “',E.v(2,r+1,c-3)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-3))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),(X.cat(X.cat('E8 “',E.v(2,r+1,c-2)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-2))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),(X.cat(X.cat('F8 “',E.v(2,r+1,c-1)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-1))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),'– chọn từ danh sách (Tí, Tị, Kỷ, Quý…).') if X.truthy(X.gt(X.LEN(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat((X.cat(X.cat('C7 “',E.v(2,r,c-4)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-4))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else ''),(X.cat(X.cat('D7 “',E.v(2,r,c-3)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-3))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('E7 “',E.v(2,r,c-2)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-2))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('F7 “',E.v(2,r,c-1)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r,c-1))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ'), E.rg(24,5,1,14,1), 0.0))) else '')),(X.cat(X.cat('C8 “',E.v(2,r+1,c-4)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-4))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),(X.cat(X.cat('D8 “',E.v(2,r+1,c-3)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-3))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),(X.cat(X.cat('E8 “',E.v(2,r+1,c-2)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-2))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else '')),(X.cat(X.cat('F8 “',E.v(2,r+1,c-1)),'” ') if X.truthy(X.ISNA(lambda: X.MATCH(X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r+1,c-1))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão'), E.rg(24,18,1,29,1), 0.0))) else ''))),0.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('✔ Đang dùng can chi tự nhập C7:F8: ',E.v(2,r+16,c-4)),', '),E.v(2,r+16,c-3)),', '),E.v(2,r+16,c-2)),', '),E.v(2,r+16,c-1)),(('. Khởi vận tự tính từ tứ trụ + năm sinh C30 (dòng 43-53).' if X.truthy(X.ISNUMBER(lambda: E.v(2,53,3))) else '. Chưa có ngày giờ sinh (C26), số ngày (C28) hoặc năm sinh khớp trụ năm (C30): đại vận tạm khởi 1 tuổi – xem dòng 43-53.') if X.truthy(X.AND(X.eq(E.v(2,26,3),''), X.eq(E.v(2,28,3),''))) else '')))))

def f1(E, r, c):
    # Ô C16, D16, E16
    # = IF(AND($C$13="Tự nhập",$C$12="Tự động",C$37<>""),C$37,IF($C$13="Tự nhập",SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C7)),"Quí","Quý"),"Kỉ","Kỷ"),INDEX(KIEM_THU
    #   !C$6:C$90,MATCH($C$13,KIEM_THU!$B$6:$B$90,0))))
    return (E.v(2,37,c) if X.truthy(X.AND(X.eq(E.v(2,13,3),'Tự nhập'), X.eq(E.v(2,12,3),'Tự động'), X.ne(E.v(2,37,c),''))) else (X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r-9,c))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ') if X.truthy(X.eq(E.v(2,13,3),'Tự nhập')) else X.INDEX(E.rg(26,6,c,90,c), X.MATCH(E.v(2,13,3), E.rg(26,6,2,90,2), 0.0))))

def f2(E, r, c):
    # Ô F16
    # = IF((NHAP_LIEU!$C$9="Không rõ giờ"),INDEX(CAN_TEN,MOD(MOD(MATCH(E16,CAN_TEN,0)-1,5)*2+(MATCH(KHONG_RO_GIO!$C$6,CHI_TEN,0)-1),10)+1),IF(AND($C$13="Tự nh
    #   ập",$C$12="Tự động",F$37<>""),F$37,IF($C$13="Tự nhập",SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(F7)),"Quí","Quý"),"Kỉ","Kỷ"),INDEX(KIEM_THU!F$6:F$90,MATCH($C$
    #   13,KIEM_THU!$B$6:$B$90,0)))))
    return (X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.add(X.mul(X.MOD(X.sub(X.MATCH(E.v(2,r,c-1), E.rg(24,5,1,14,1), 0.0),1.0), 5.0),2.0),X.sub(X.MATCH(E.v(6,6,3), E.rg(24,18,1,29,1), 0.0),1.0)), 10.0),1.0)) if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else (E.v(2,37,c) if X.truthy(X.AND(X.eq(E.v(2,13,3),'Tự nhập'), X.eq(E.v(2,12,3),'Tự động'), X.ne(E.v(2,37,c),''))) else (X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r-9,c))), 'Quí', 'Quý'), 'Kỉ', 'Kỷ') if X.truthy(X.eq(E.v(2,13,3),'Tự nhập')) else X.INDEX(E.rg(26,6,c,90,c), X.MATCH(E.v(2,13,3), E.rg(26,6,2,90,2), 0.0)))))

def f3(E, r, c):
    # Ô C17, D17, E17
    # = IF(AND($C$13="Tự nhập",$C$12="Tự động",C$38<>""),C$38,IF($C$13="Tự nhập",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(TRIM(C8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo"
    #   ,"Mão"),INDEX(KIEM_THU!G$6:G$90,MATCH($C$13,KIEM_THU!$B$6:$B$90,0))))
    return (E.v(2,38,c) if X.truthy(X.AND(X.eq(E.v(2,13,3),'Tự nhập'), X.eq(E.v(2,12,3),'Tự động'), X.ne(E.v(2,38,c),''))) else (X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r-9,c))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão') if X.truthy(X.eq(E.v(2,13,3),'Tự nhập')) else X.INDEX(E.rg(26,6,c+4,90,c+4), X.MATCH(E.v(2,13,3), E.rg(26,6,2,90,2), 0.0))))

def f4(E, r, c):
    # Ô F17
    # = IF((NHAP_LIEU!$C$9="Không rõ giờ"),KHONG_RO_GIO!$C$6,IF(AND($C$13="Tự nhập",$C$12="Tự động",F$38<>""),F$38,IF($C$13="Tự nhập",SUBSTITUTE(SUBSTITUTE(SU
    #   BSTITUTE(PROPER(TRIM(F8)),"Tý","Tí"),"Tỵ","Tị"),"Mẹo","Mão"),INDEX(KIEM_THU!J$6:J$90,MATCH($C$13,KIEM_THU!$B$6:$B$90,0)))))
    return (E.v(6,6,3) if X.truthy(X.eq(E.v(2,9,3),'Không rõ giờ')) else (E.v(2,38,c) if X.truthy(X.AND(X.eq(E.v(2,13,3),'Tự nhập'), X.eq(E.v(2,12,3),'Tự động'), X.ne(E.v(2,38,c),''))) else (X.SUBSTITUTE(X.SUBSTITUTE(X.SUBSTITUTE(X.PROPER(X.TRIM(E.v(2,r-9,c))), 'Tý', 'Tí'), 'Tỵ', 'Tị'), 'Mẹo', 'Mão') if X.truthy(X.eq(E.v(2,13,3),'Tự nhập')) else X.INDEX(E.rg(26,6,c+4,90,c+4), X.MATCH(E.v(2,13,3), E.rg(26,6,2,90,2), 0.0)))))

def f5(E, r, c):
    # Ô C18
    # = IF($C$13="Tự nhập",IF(C11<>"",C11,IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$10)),INDEX(TIET_UTC,DAI_VAN!$C$11+1)-DAI_VAN!$C$10,"")),"")
    return ((E.v(2,r-7,c) if X.truthy(X.ne(E.v(2,r-7,c),'')) else (X.sub(X.INDEX(E.rg(25,4,1,2439,1), X.add(E.v(14,11,3),1.0)),E.v(14,10,3)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,10,3)))) else '')) if X.truthy(X.eq(E.v(2,13,3),'Tự nhập')) else '')

def f6(E, r, c):
    # Ô C20, D20, E20, F20
    # = MATCH(C16,CAN_TEN,0)-1
    return X.sub(X.MATCH(E.v(2,r-4,c), E.rg(24,5,1,14,1), 0.0),1.0)

def f7(E, r, c):
    # Ô C21, D21, E21, F21
    # = MATCH(C17,CHI_TEN,0)-1
    return X.sub(X.MATCH(E.v(2,r-4,c), E.rg(24,18,1,29,1), 0.0),1.0)

def f8(E, r, c):
    # Ô C23, D23, E23, F23
    # = C16&" "&C17
    return X.cat(X.cat(E.v(2,r-7,c),' '),E.v(2,r-6,c))

def f9(E, r, c):
    # Ô D28
    # = IF(AND(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),ISNUMBER(DAI_VAN!$C$17)),"✔ TỰ TÍNH từ ngày giờ sinh: "&TEXT(DAI_VAN!$C$17,"0.00")&" ngày (
    #   "&IF(DAI_VAN!$C$7=1,"thuận, đến "&DAI_VAN!$C$13,"nghịch, lùi về "&DAI_VAN!$C$12)&") → khởi vận "&DAI_VAN!$D$20&". Không cần nhập ô C28.",IF(AND(NHAP_L
    #   IEU!$C$28="",ISNUMBER(NHAP_LIEU!$C$53)),"✔ TỰ TÍNH từ tứ trụ + năm sinh C30: "&TEXT(NHAP_LIEU!$C$53,"0.00")&" ngày → khởi vận "&DAI_VAN!$D$20&" (ngày 
    #   sinh suy ra: "&NHAP_LIEU!$D$52&" – xem dòng 43-53). Không cần nhập ô C28.","Thuận: đếm tới tiết kế tiếp; nghịch: đếm lùi về tiết vừa qua (P1 tr.119-12
    #   0; P2 tr.275-278)"))
    return (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('✔ TỰ TÍNH từ ngày giờ sinh: ',X.TEXT(E.v(14,17,3), '0.00')),' ngày ('),(X.cat('thuận, đến ',E.v(14,13,3)) if X.truthy(X.eq(E.v(14,7,3),1.0)) else X.cat('nghịch, lùi về ',E.v(14,12,3)))),') → khởi vận '),E.v(14,20,4)),'. Không cần nhập ô C28.') if X.truthy(X.AND(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3))), X.ISNUMBER(lambda: E.v(14,17,3)))) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('✔ TỰ TÍNH từ tứ trụ + năm sinh C30: ',X.TEXT(E.v(2,53,3), '0.00')),' ngày → khởi vận '),E.v(14,20,4)),' (ngày sinh suy ra: '),E.v(2,52,4)),' – xem dòng 43-53). Không cần nhập ô C28.') if X.truthy(X.AND(X.eq(E.v(2,28,3),''), X.ISNUMBER(lambda: E.v(2,53,3)))) else 'Thuận: đếm tới tiết kế tiếp; nghịch: đếm lùi về tiết vừa qua (P1 tr.119-120; P2 tr.275-278)'))

def f10(E, r, c):
    # Ô D30
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),"✔ TỰ TÍNH từ ngày giờ sinh: "&DAI_VAN!$C$14&" (năm đổi tại Lập xuân – P2 tr.268). Không cần nh
    #   ập ô C30.","Năm tính theo Lập xuân (P2 tr.268): sinh từ 1/1 đến trước Lập xuân (~4/2) thì nhập năm trước, vd 13/1/1981 → 1980")
    return (X.cat(X.cat('✔ TỰ TÍNH từ ngày giờ sinh: ',E.v(14,14,3)),' (năm đổi tại Lập xuân – P2 tr.268). Không cần nhập ô C30.') if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else 'Năm tính theo Lập xuân (P2 tr.268): sinh từ 1/1 đến trước Lập xuân (~4/2) thì nhập năm trước, vd 13/1/1981 → 1980')

def f11(E, r, c):
    # Ô C37
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CAN_TEN,MOD(MOD(DAI_VAN!$C$14-1984,60),10)+1),"")
    return (X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.MOD(X.sub(E.v(14,14,3),1984.0), 60.0), 10.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f12(E, r, c):
    # Ô D37
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CAN_TEN,MOD(MOD(MOD(MOD(DAI_VAN!$C$14-1984,60),10),5)*2+2+MOD(DAI_VAN!$C$15-2,12),10)+1),
    #   "")
    return (X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.add(X.add(X.mul(X.MOD(X.MOD(X.MOD(X.sub(E.v(14,14,3),1984.0), 60.0), 10.0), 5.0),2.0),2.0),X.MOD(X.sub(E.v(14,15,3),2.0), 12.0)), 10.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f13(E, r, c):
    # Ô E37
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CAN_TEN,MOD(MOD(INT(DAI_VAN!$C$8+1/24)+8,60),10)+1),"")
    return (X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.MOD(X.add(X.INT(X.add(E.v(14,8,3),X.div(1.0,24.0))),8.0), 60.0), 10.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f14(E, r, c):
    # Ô F37
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CAN_TEN,MOD(MOD(MOD(MOD(INT(DAI_VAN!$C$8+1/24)+8,60),10),5)*2+MOD(INT((ROUND(MOD(DAI_VAN!
    #   $C$8,1)*1440,0)+60)/120),12),10)+1),"")
    return (X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.add(X.mul(X.MOD(X.MOD(X.MOD(X.add(X.INT(X.add(E.v(14,8,3),X.div(1.0,24.0))),8.0), 60.0), 10.0), 5.0),2.0),X.MOD(X.INT(X.div(X.add(X.ROUND(X.mul(X.MOD(E.v(14,8,3), 1.0),1440.0), 0.0),60.0),120.0)), 12.0)), 10.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f15(E, r, c):
    # Ô C38
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CHI_TEN,MOD(MOD(DAI_VAN!$C$14-1984,60),12)+1),"")
    return (X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.MOD(X.sub(E.v(14,14,3),1984.0), 60.0), 12.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f16(E, r, c):
    # Ô D38
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CHI_TEN,DAI_VAN!$C$15+1),"")
    return (X.INDEX(E.rg(24,18,1,29,1), X.add(E.v(14,15,3),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f17(E, r, c):
    # Ô E38
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CHI_TEN,MOD(MOD(INT(DAI_VAN!$C$8+1/24)+8,60),12)+1),"")
    return (X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.MOD(X.add(X.INT(X.add(E.v(14,8,3),X.div(1.0,24.0))),8.0), 60.0), 12.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f18(E, r, c):
    # Ô F38
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),INDEX(CHI_TEN,MOD(INT((ROUND(MOD(DAI_VAN!$C$8,1)*1440,0)+60)/120),12)+1),"")
    return (X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.INT(X.div(X.add(X.ROUND(X.mul(X.MOD(E.v(14,8,3), 1.0),1440.0), 0.0),60.0),120.0)), 12.0),1.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f19(E, r, c):
    # Ô C39
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),MOD(MOD(DAI_VAN!$C$14-1984,60),10)&"|"&MOD(MOD(DAI_VAN!$C$14-1984,60),12),"")
    return (X.cat(X.cat(X.MOD(X.MOD(X.sub(E.v(14,14,3),1984.0), 60.0), 10.0),'|'),X.MOD(X.MOD(X.sub(E.v(14,14,3),1984.0), 60.0), 12.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f20(E, r, c):
    # Ô D39
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),MOD(MOD(MOD(MOD(DAI_VAN!$C$14-1984,60),10),5)*2+2+MOD(DAI_VAN!$C$15-2,12),10)&"|"&DAI_VAN!$C$15
    #   ,"")
    return (X.cat(X.cat(X.MOD(X.add(X.add(X.mul(X.MOD(X.MOD(X.MOD(X.sub(E.v(14,14,3),1984.0), 60.0), 10.0), 5.0),2.0),2.0),X.MOD(X.sub(E.v(14,15,3),2.0), 12.0)), 10.0),'|'),E.v(14,15,3)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f21(E, r, c):
    # Ô E39
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),MOD(MOD(INT(DAI_VAN!$C$8+1/24)+8,60),10)&"|"&MOD(MOD(INT(DAI_VAN!$C$8+1/24)+8,60),12),"")
    return (X.cat(X.cat(X.MOD(X.MOD(X.add(X.INT(X.add(E.v(14,8,3),X.div(1.0,24.0))),8.0), 60.0), 10.0),'|'),X.MOD(X.MOD(X.add(X.INT(X.add(E.v(14,8,3),X.div(1.0,24.0))),8.0), 60.0), 12.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f22(E, r, c):
    # Ô F39
    # = IF(AND(ISNUMBER(DAI_VAN!$C$11),ISNUMBER(DAI_VAN!$C$8)),MOD(MOD(MOD(MOD(INT(DAI_VAN!$C$8+1/24)+8,60),10),5)*2+MOD(INT((ROUND(MOD(DAI_VAN!$C$8,1)*1440,0
    #   )+60)/120),12),10)&"|"&MOD(INT((ROUND(MOD(DAI_VAN!$C$8,1)*1440,0)+60)/120),12),"")
    return (X.cat(X.cat(X.MOD(X.add(X.mul(X.MOD(X.MOD(X.MOD(X.add(X.INT(X.add(E.v(14,8,3),X.div(1.0,24.0))),8.0), 60.0), 10.0), 5.0),2.0),X.MOD(X.INT(X.div(X.add(X.ROUND(X.mul(X.MOD(E.v(14,8,3), 1.0),1440.0), 0.0),60.0),120.0)), 12.0)), 10.0),'|'),X.MOD(X.INT(X.div(X.add(X.ROUND(X.mul(X.MOD(E.v(14,8,3), 1.0),1440.0), 0.0),60.0),120.0)), 12.0)) if X.truthy(X.AND(X.ISNUMBER(lambda: E.v(14,11,3)), X.ISNUMBER(lambda: E.v(14,8,3)))) else '')

def f23(E, r, c):
    # Ô C41
    # = IF($C$13<>"Tự nhập","Lá số mẫu "&$C$13,IF(AND($C$12="Tự động",C37<>""),"Ngày giờ sinh (dòng 37-38)","Ô tự nhập C7:F8"))
    return (X.cat('Lá số mẫu ',E.v(2,13,3)) if X.truthy(X.ne(E.v(2,13,3),'Tự nhập')) else ('Ngày giờ sinh (dòng 37-38)' if X.truthy(X.AND(X.eq(E.v(2,12,3),'Tự động'), X.ne(E.v(2,r-4,c),''))) else 'Ô tự nhập C7:F8'))

def f24(E, r, c):
    # Ô C45
    # = IF(ISNUMBER($C$30),IF(MOD($C$30-1984,60)=MOD(6*TRUNG_GIAN!$C$5-5*TRUNG_GIAN!$C$6,60),1,IF(AND(MOD($C$30-1-1984,60)=MOD(6*TRUNG_GIAN!$C$5-5*TRUNG_GIAN!
    #   $C$6,60),OR(TRUNG_GIAN!$D$6=0,TRUNG_GIAN!$D$6=1)),2,0)),"")
    return ((1.0 if X.truthy(X.eq(X.MOD(X.sub(E.v(2,30,3),1984.0), 60.0),X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0))) else (2.0 if X.truthy(X.AND(X.eq(X.MOD(X.sub(X.sub(E.v(2,30,3),1.0),1984.0), 60.0),X.MOD(X.sub(X.mul(6.0,E.v(23,5,3)),X.mul(5.0,E.v(23,6,3))), 60.0)), X.OR(X.eq(E.v(23,6,4),0.0), X.eq(E.v(23,6,4),1.0)))) else 0.0)) if X.truthy(X.ISNUMBER(lambda: E.v(2,30,3))) else '')

def f25(E, r, c):
    # Ô D45
    # = IF(ISNUMBER(DAI_VAN!$C$11),"Không cần – đã có ngày giờ sinh ở C26:G26 (khởi vận tính trực tiếp, xem D28).",IF($C$30="","Nhập năm sinh dương lịch của t
    #   rụ năm ở C30 (năm đổi tại Lập xuân).",IF(C45=1,"Khớp: năm "&$C$30&" là năm "&C16&" "&C17&".",IF(C45=2,"Khớp sau hiệu chỉnh: sinh tháng "&D17&" trước L
    #   ập xuân nên năm dương lịch "&$C$30&" thuộc năm can chi "&($C$30-1)&" ("&C16&" "&C17&") – công cụ tự dùng "&($C$30-1)&" (P2 tr.268).","KHÔNG KHỚP: năm 
    #   "&$C$30&" là năm "&INDEX(CAN_TEN,MOD($C$30-1984,10)+1)&" "&INDEX(CHI_TEN,MOD($C$30-1984,12)+1)&", trụ năm đang dùng là "&C16&" "&C17&" – sửa C30 (các 
    #   năm hợp lệ cách nhau 60 năm)."))))
    return ('Không cần – đã có ngày giờ sinh ở C26:G26 (khởi vận tính trực tiếp, xem D28).' if X.truthy(X.ISNUMBER(lambda: E.v(14,11,3))) else ('Nhập năm sinh dương lịch của trụ năm ở C30 (năm đổi tại Lập xuân).' if X.truthy(X.eq(E.v(2,30,3),'')) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Khớp: năm ',E.v(2,30,3)),' là năm '),E.v(2,r-29,c-1)),' '),E.v(2,r-28,c-1)),'.') if X.truthy(X.eq(E.v(2,r,c-1),1.0)) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Khớp sau hiệu chỉnh: sinh tháng ',E.v(2,r-28,c)),' trước Lập xuân nên năm dương lịch '),E.v(2,30,3)),' thuộc năm can chi '),X.sub(E.v(2,30,3),1.0)),' ('),E.v(2,r-29,c-1)),' '),E.v(2,r-28,c-1)),') – công cụ tự dùng '),X.sub(E.v(2,30,3),1.0)),' (P2 tr.268).') if X.truthy(X.eq(E.v(2,r,c-1),2.0)) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('KHÔNG KHỚP: năm ',E.v(2,30,3)),' là năm '),X.INDEX(E.rg(24,5,1,14,1), X.add(X.MOD(X.sub(E.v(2,30,3),1984.0), 10.0),1.0))),' '),X.INDEX(E.rg(24,18,1,29,1), X.add(X.MOD(X.sub(E.v(2,30,3),1984.0), 12.0),1.0))),', trụ năm đang dùng là '),E.v(2,r-29,c-1)),' '),E.v(2,r-28,c-1)),' – sửa C30 (các năm hợp lệ cách nhau 60 năm).')))))

def f26(E, r, c):
    # Ô C46
    # = IF(AND(C45<>1,C45<>2),"",IFERROR(IF(AND(INDEX(TIET_CHI,($C$54-1899)*12+1+MOD(TRUNG_GIAN!$D$6-2,12))=TRUNG_GIAN!$D$6,INDEX(TIET_NAM,($C$54-1899)*12+1+M
    #   OD(TRUNG_GIAN!$D$6-2,12))=$C$54),($C$54-1899)*12+1+MOD(TRUNG_GIAN!$D$6-2,12),""),""))
    return ('' if X.truthy(X.AND(X.ne(E.v(2,r-1,c),1.0), X.ne(E.v(2,r-1,c),2.0))) else X.IFERROR(lambda: (X.add(X.add(X.mul(X.sub(E.v(2,54,3),1899.0),12.0),1.0),X.MOD(X.sub(E.v(23,6,4),2.0), 12.0)) if X.truthy(X.AND(X.eq(X.INDEX(E.rg(25,4,4,2439,4), X.add(X.add(X.mul(X.sub(E.v(2,54,3),1899.0),12.0),1.0),X.MOD(X.sub(E.v(23,6,4),2.0), 12.0))),E.v(23,6,4)), X.eq(X.INDEX(E.rg(25,4,5,2439,5), X.add(X.add(X.mul(X.sub(E.v(2,54,3),1899.0),12.0),1.0),X.MOD(X.sub(E.v(23,6,4),2.0), 12.0))),E.v(2,54,3)))) else ''), lambda: ''))

def f27(E, r, c):
    # Ô D46
    # = IF(C46="","",INDEX(TIET_TEN,C46)&" → "&INDEX(TIET_TEN,C46+1)&" (tháng "&D17&" năm "&$C$54&")")
    return ('' if X.truthy(X.eq(E.v(2,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.INDEX(E.rg(25,4,3,2439,3), E.v(2,r,c-1)),' → '),X.INDEX(E.rg(25,4,3,2439,3), X.add(E.v(2,r,c-1),1.0))),' (tháng '),E.v(2,r-29,c)),' năm '),E.v(2,54,3)),')'))

def f28(E, r, c):
    # Ô C47
    # = IF(C46="","",INDEX(TIET_UTC,C46)+DAI_VAN!$C$9/24)
    return ('' if X.truthy(X.eq(E.v(2,r-1,c),'')) else X.add(X.INDEX(E.rg(25,4,1,2439,1), E.v(2,r-1,c)),X.div(E.v(14,9,3),24.0)))

def f29(E, r, c):
    # Ô C48
    # = IF(C46="","",INDEX(TIET_UTC,C46+1)+DAI_VAN!$C$9/24)
    return ('' if X.truthy(X.eq(E.v(2,r-2,c),'')) else X.add(X.INDEX(E.rg(25,4,1,2439,1), X.add(E.v(2,r-2,c),1.0)),X.div(E.v(14,9,3),24.0)))

def f30(E, r, c):
    # Ô C49
    # = IF(C46="","",INT(C47+1/24)+MOD(MOD(6*TRUNG_GIAN!$E$5-5*TRUNG_GIAN!$E$6,60)-MOD(INT(C47+1/24)+8,60),60))
    return ('' if X.truthy(X.eq(E.v(2,r-3,c),'')) else X.add(X.INT(X.add(E.v(2,r-2,c),X.div(1.0,24.0))),X.MOD(X.sub(X.MOD(X.sub(X.mul(6.0,E.v(23,5,5)),X.mul(5.0,E.v(23,6,5))), 60.0),X.MOD(X.add(X.INT(X.add(E.v(2,r-2,c),X.div(1.0,24.0))),8.0), 60.0)), 60.0)))

def f31(E, r, c):
    # Ô C50
    # = IF(C49="","",MAX(C49+(2*TRUNG_GIAN!$F$6-1)/24,C47))
    return ('' if X.truthy(X.eq(E.v(2,r-1,c),'')) else X.MAX(X.add(E.v(2,r-1,c),X.div(X.sub(X.mul(2.0,E.v(23,6,6)),1.0),24.0)), E.rg(2,r-3,c,r-3,c)))

def f32(E, r, c):
    # Ô C51
    # = IF(C49="","",MIN(C49+(2*TRUNG_GIAN!$F$6+1)/24,C48))
    return ('' if X.truthy(X.eq(E.v(2,r-2,c),'')) else X.MIN(X.add(E.v(2,r-2,c),X.div(X.add(X.mul(2.0,E.v(23,6,6)),1.0),24.0)), E.rg(2,r-3,c,r-3,c)))

def f33(E, r, c):
    # Ô C52
    # = IF(C49="","",IF(C51>C50,(C50+C51)/2,""))
    return ('' if X.truthy(X.eq(E.v(2,r-3,c),'')) else (X.div(X.add(E.v(2,r-2,c),E.v(2,r-1,c)),2.0) if X.truthy(X.gt(E.v(2,r-1,c),E.v(2,r-2,c))) else ''))

def f34(E, r, c):
    # Ô D52
    # = IF(C49="","",IF(C52="","KHÔNG CÓ ngày nào trong tháng "&D17&" năm "&$C$54&" khớp trụ ngày "&E16&" "&E17&" và giờ "&F17&" – kiểm tra lại tứ trụ hoặc nă
    #   m sinh C30.","Khoảng "&TEXT(C52,"dd/mm/yyyy hh:mm")&" (giờ "&F17&": "&TEXT(C50,"hh:mm")&"–"&TEXT(C51,"hh:mm")&")"))
    return ('' if X.truthy(X.eq(E.v(2,r-3,c-1),'')) else (X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('KHÔNG CÓ ngày nào trong tháng ',E.v(2,r-35,c)),' năm '),E.v(2,54,3)),' khớp trụ ngày '),E.v(2,r-36,c+1)),' '),E.v(2,r-35,c+1)),' và giờ '),E.v(2,r-35,c+2)),' – kiểm tra lại tứ trụ hoặc năm sinh C30.') if X.truthy(X.eq(E.v(2,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Khoảng ',X.TEXT(E.v(2,r,c-1), 'dd/mm/yyyy hh:mm')),' (giờ '),E.v(2,r-35,c+2)),': '),X.TEXT(E.v(2,r-2,c-1), 'hh:mm')),'–'),X.TEXT(E.v(2,r-1,c-1), 'hh:mm')),')')))

def f35(E, r, c):
    # Ô C53
    # = IF(C52="","",IF(DAI_VAN!$C$7=1,C48-C52,C52-C47))
    return ('' if X.truthy(X.eq(E.v(2,r-1,c),'')) else (X.sub(E.v(2,r-5,c),E.v(2,r-1,c)) if X.truthy(X.eq(E.v(14,7,3),1.0)) else X.sub(E.v(2,r-1,c),E.v(2,r-6,c))))

def f36(E, r, c):
    # Ô D53
    # = IF(C53="","",IF(DAI_VAN!$C$7=1,"Thuận: đến ","Nghịch: lùi về ")&IF(DAI_VAN!$C$7=1,INDEX(TIET_TEN,C46+1)&" "&TEXT(C48,"dd/mm/yyyy hh:mm"),INDEX(TIET_TE
    #   N,C46)&" "&TEXT(C47,"dd/mm/yyyy hh:mm"))&" → khởi vận "&INT(C53/3)&" tuổi "&ROUND((C53/3-INT(C53/3))*12,0)&" tháng (sai số ± "&TEXT((C51-C50)/2*120,"0
    #   ")&" ngày tuổi)")
    return ('' if X.truthy(X.eq(E.v(2,r,c-1),'')) else X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(('Thuận: đến ' if X.truthy(X.eq(E.v(14,7,3),1.0)) else 'Nghịch: lùi về '),(X.cat(X.cat(X.INDEX(E.rg(25,4,3,2439,3), X.add(E.v(2,r-7,c-1),1.0)),' '),X.TEXT(E.v(2,r-5,c-1), 'dd/mm/yyyy hh:mm')) if X.truthy(X.eq(E.v(14,7,3),1.0)) else X.cat(X.cat(X.INDEX(E.rg(25,4,3,2439,3), E.v(2,r-7,c-1)),' '),X.TEXT(E.v(2,r-6,c-1), 'dd/mm/yyyy hh:mm')))),' → khởi vận '),X.INT(X.div(E.v(2,r,c-1),3.0))),' tuổi '),X.ROUND(X.mul(X.sub(X.div(E.v(2,r,c-1),3.0),X.INT(X.div(E.v(2,r,c-1),3.0))),12.0), 0.0)),' tháng (sai số ± '),X.TEXT(X.mul(X.div(X.sub(E.v(2,r-2,c-1),E.v(2,r-3,c-1)),2.0),120.0), '0')),' ngày tuổi)'))

def f37(E, r, c):
    # Ô C54
    # = IF(C45=1,$C$30,IF(C45=2,$C$30-1,""))
    return (E.v(2,30,3) if X.truthy(X.eq(E.v(2,r-9,c),1.0)) else (X.sub(E.v(2,30,3),1.0) if X.truthy(X.eq(E.v(2,r-9,c),2.0)) else ''))

CELLS = {
    'G7': (f0, 7, 7),
    'C16': (f1, 16, 3),
    'D16': (f1, 16, 4),
    'E16': (f1, 16, 5),
    'F16': (f2, 16, 6),
    'C17': (f3, 17, 3),
    'D17': (f3, 17, 4),
    'E17': (f3, 17, 5),
    'F17': (f4, 17, 6),
    'C18': (f5, 18, 3),
    'C20': (f6, 20, 3),
    'D20': (f6, 20, 4),
    'E20': (f6, 20, 5),
    'F20': (f6, 20, 6),
    'C21': (f7, 21, 3),
    'D21': (f7, 21, 4),
    'E21': (f7, 21, 5),
    'F21': (f7, 21, 6),
    'C23': (f8, 23, 3),
    'D23': (f8, 23, 4),
    'E23': (f8, 23, 5),
    'F23': (f8, 23, 6),
    'D28': (f9, 28, 4),
    'D30': (f10, 30, 4),
    'C37': (f11, 37, 3),
    'D37': (f12, 37, 4),
    'E37': (f13, 37, 5),
    'F37': (f14, 37, 6),
    'C38': (f15, 38, 3),
    'D38': (f16, 38, 4),
    'E38': (f17, 38, 5),
    'F38': (f18, 38, 6),
    'C39': (f19, 39, 3),
    'D39': (f20, 39, 4),
    'E39': (f21, 39, 5),
    'F39': (f22, 39, 6),
    'C41': (f23, 41, 3),
    'C45': (f24, 45, 3),
    'D45': (f25, 45, 4),
    'C46': (f26, 46, 3),
    'D46': (f27, 46, 4),
    'C47': (f28, 47, 3),
    'C48': (f29, 48, 3),
    'C49': (f30, 49, 3),
    'C50': (f31, 50, 3),
    'C51': (f32, 51, 3),
    'C52': (f33, 52, 3),
    'D52': (f34, 52, 4),
    'C53': (f35, 53, 3),
    'D53': (f36, 53, 4),
    'C54': (f37, 54, 3),
}
