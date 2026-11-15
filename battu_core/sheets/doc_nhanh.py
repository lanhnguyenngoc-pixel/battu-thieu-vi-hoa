# -*- coding: utf-8 -*-
"""Module sheet DOC_NHANH – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.6.xlsx (SHA-256 c1e0254833a8f522615199e1e0140460e912227d0b9ffc88b9164ebc26d3e0b3).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'DOC_NHANH'
SI = 31

def f0(E, r, c):
    # Ô C4
    # = "Giới tính "&LOWER(NHAP_LIEU!$C$22)&" · tứ trụ (năm – tháng – ngày – giờ): "&KET_LUAN!C15&" – "&KET_LUAN!D15&" – "&KET_LUAN!E15&" – "&KET_LUAN!F15&" ·
    #    năm xem "&LUU_NIEN!A5&" (tuổi mụ "&LUU_NIEN!B5&")"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Giới tính ',X.LOWER(E.v(2,22,3))),' · tứ trụ (năm – tháng – ngày – giờ): '),E.v(3,r+11,c)),' – '),E.v(3,r+11,c+1)),' – '),E.v(3,r+11,c+2)),' – '),E.v(3,r+11,c+3)),' · năm xem '),E.v(15,r+1,c-2)),' (tuổi mụ '),E.v(15,r+1,c-1)),')')

def f1(E, r, c):
    # Ô C5
    # = IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C4,CHAR(10),REPT(" ",2000)),1,2000)),FIND(" Lý do:",TRIM(MID(SUBSTITUTE(KET_LUAN!C4,CHAR(10),REPT(" ",2000))
    #   ,1,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C4,CHAR(10),REPT(" ",2000)),1,2000)))&CHAR(10)&IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C4,CHAR(10),REPT("
    #    ",2000)),2001,2000)),FIND(" Cứu ứng",TRIM(MID(SUBSTITUTE(KET_LUAN!C4,CHAR(10),REPT(" ",2000)),2001,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C4,CHAR(10
    #   ),REPT(" ",2000)),2001,2000)))
    return X.cat(X.cat(X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-1,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0)), X.sub(X.FIND(' Lý do:', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-1,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-1,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),X.CHAR(10.0)),X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-1,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0)), X.sub(X.FIND(' Cứu ứng', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-1,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-1,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0))))

def f2(E, r, c):
    # Ô C6
    # = IFERROR(MID(KET_LUAN!C4,FIND("► Lời khuyên",KET_LUAN!C4),FIND("[Lời khuyên của công cụ",KET_LUAN!C4)-FIND("► Lời khuyên",KET_LUAN!C4)),"")&"[Lời khuyê
    #   n của công cụ – không phải câu sách]"
    return X.cat(X.IFERROR(lambda: X.MID(E.v(3,r-2,c), X.FIND('► Lời khuyên', E.v(3,r-2,c)), X.sub(X.FIND('[Lời khuyên của công cụ', E.v(3,r-2,c)),X.FIND('► Lời khuyên', E.v(3,r-2,c)))), lambda: ''),'[Lời khuyên của công cụ – không phải câu sách]')

def f3(E, r, c):
    # Ô C7
    # = IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C5,CHAR(10),REPT(" ",2000)),1,2000)),FIND("Thiên hướng nghề",TRIM(MID(SUBSTITUTE(KET_LUAN!C5,CHAR(10),REPT("
    #    ",2000)),1,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C5,CHAR(10),REPT(" ",2000)),1,2000)))&CHAR(10)&IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C5,CHAR(1
    #   0),REPT(" ",2000)),2001,2000)),FIND(" (Tiêu chí",TRIM(MID(SUBSTITUTE(KET_LUAN!C5,CHAR(10),REPT(" ",2000)),2001,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN
    #   !C5,CHAR(10),REPT(" ",2000)),2001,2000)))
    return X.cat(X.cat(X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-2,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0)), X.sub(X.FIND('Thiên hướng nghề', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-2,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-2,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),X.CHAR(10.0)),X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-2,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0)), X.sub(X.FIND(' (Tiêu chí', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-2,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-2,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0))))

def f4(E, r, c):
    # Ô C8
    # = IF(LEFT(KET_LUAN!C72,6)="(Lưu ý",TRIM(MID(SUBSTITUTE(KET_LUAN!C72,CHAR(10),REPT(" ",2000)),1,2000))&CHAR(10)&TRIM(MID(SUBSTITUTE(KET_LUAN!C72,CHAR(10)
    #   ,REPT(" ",2000)),2001,2000)),TRIM(MID(SUBSTITUTE(KET_LUAN!C72,CHAR(10),REPT(" ",2000)),1,2000)))&IFERROR(CHAR(10)&MID(KET_LUAN!C72,FIND("⚠",KET_LUAN!C
    #   72),400),"")
    return X.cat((X.cat(X.cat(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r+64,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0)),X.CHAR(10.0)),X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r+64,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0))) if X.truthy(X.eq(X.LEFT(E.v(3,r+64,c), 6.0),'(Lưu ý')) else X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r+64,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),X.IFERROR(lambda: X.cat(X.CHAR(10.0),X.MID(E.v(3,r+64,c), X.FIND('⚠', E.v(3,r+64,c)), 400.0)), lambda: ''))

def f5(E, r, c):
    # Ô C9
    # = IFERROR(LEFT(KET_LUAN!C6,FIND(" (điểm",KET_LUAN!C6)-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C6,CHAR(10),REPT(" ",2000)),1,2000)))&IFERROR(". "&MID(KET_LUAN!C6
    #   ,FIND("Các căn cứ",KET_LUAN!C6),FIND(CHAR(10),KET_LUAN!C6&CHAR(10),FIND("Các căn cứ",KET_LUAN!C6))-FIND("Các căn cứ",KET_LUAN!C6)),"")
    return X.cat(X.IFERROR(lambda: X.LEFT(E.v(3,r-3,c), X.sub(X.FIND(' (điểm', E.v(3,r-3,c)),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),X.IFERROR(lambda: X.cat('. ',X.MID(E.v(3,r-3,c), X.FIND('Các căn cứ', E.v(3,r-3,c)), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(3,r-3,c),X.CHAR(10.0)), X.FIND('Các căn cứ', E.v(3,r-3,c))),X.FIND('Các căn cứ', E.v(3,r-3,c))))), lambda: ''))

def f6(E, r, c):
    # Ô C10
    # = IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),1,2000)),FIND(" (theo",TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),
    #   1,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),1,2000)))&CHAR(10)&IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" 
    #   ",2000)),4001,2000)),FIND("; đại vận",TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),4001,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10
    #   ),REPT(" ",2000)),4001,2000)))&CHAR(10)&IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),6001,2000)),FIND("; đại vận",TRIM(MID(SU
    #   BSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),6001,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C7,CHAR(10),REPT(" ",2000)),6001,2000)))
    return X.cat(X.cat(X.cat(X.cat(X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0)), X.sub(X.FIND(' (theo', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),X.CHAR(10.0)),X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 4001.0, 2000.0)), X.sub(X.FIND('; đại vận', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 4001.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 4001.0, 2000.0)))),X.CHAR(10.0)),X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 6001.0, 2000.0)), X.sub(X.FIND('; đại vận', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 6001.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 6001.0, 2000.0))))

def f7(E, r, c):
    # Ô C11
    # = TRIM(MID(SUBSTITUTE(KET_LUAN!C8,CHAR(10),REPT(" ",2000)),1,2000))&CHAR(10)&IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C8,CHAR(10),REPT(" ",2000)),2001,
    #   2000)),FIND(". Lý do",TRIM(MID(SUBSTITUTE(KET_LUAN!C8,CHAR(10),REPT(" ",2000)),2001,2000)))),TRIM(MID(SUBSTITUTE(KET_LUAN!C8,CHAR(10),REPT(" ",2000)),
    #   2001,2000)))
    return X.cat(X.cat(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0)),X.CHAR(10.0)),X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0)), X.FIND('. Lý do', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0)))), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r-3,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 2001.0, 2000.0))))

def f8(E, r, c):
    # Ô C12
    # = "Chung: "&TRIM(MID(SUBSTITUTE(SUC_KHOE!J6,CHAR(10),REPT(" ",2000)),1,2000))&" · bệnh tật: "&TRIM(MID(SUBSTITUTE(SUC_KHOE!J7,CHAR(10),REPT(" ",2000)),1
    #   ,2000))&" · thương tật – tai họa: "&TRIM(MID(SUBSTITUTE(SUC_KHOE!J8,CHAR(10),REPT(" ",2000)),1,2000))&" (chi tiết KET_LUAN dòng 134-157)"
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat('Chung: ',X.TRIM(X.MID(X.SUBSTITUTE(E.v(22,r-6,c+7), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),' · bệnh tật: '),X.TRIM(X.MID(X.SUBSTITUTE(E.v(22,r-5,c+7), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),' · thương tật – tai họa: '),X.TRIM(X.MID(X.SUBSTITUTE(E.v(22,r-4,c+7), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),' (chi tiết KET_LUAN dòng 134-157)')

def f9(E, r, c):
    # Ô C13
    # = IFERROR(LEFT(KET_LUAN!C10,FIND(". ",KET_LUAN!C10,60)),LEFT(KET_LUAN!C10,220))&CHAR(10)&IFERROR(LEFT(DUNG_HY_KY!C189,FIND(" Các vận còn lại",DUNG_HY_KY
    #   !C189)-1),DUNG_HY_KY!C189)
    return X.cat(X.cat(X.IFERROR(lambda: X.LEFT(E.v(3,r-3,c), X.FIND('. ', E.v(3,r-3,c), 60.0)), lambda: X.LEFT(E.v(3,r-3,c), 220.0)),X.CHAR(10.0)),X.IFERROR(lambda: X.LEFT(E.v(8,r+176,c), X.sub(X.FIND(' Các vận còn lại', E.v(8,r+176,c)),1.0)), lambda: E.v(8,r+176,c)))

def f10(E, r, c):
    # Ô C14
    # = KET_LUAN!C140&CHAR(10)&"Năm lo về người thân – sinh mệnh (cả năm đã qua): "&LUU_NIEN!DK31&CHAR(10)&"Năm tai vạ nặng: "&LUU_NIEN!DL31
    return X.cat(X.cat(X.cat(X.cat(X.cat(X.cat(E.v(3,r+126,c),X.CHAR(10.0)),'Năm lo về người thân – sinh mệnh (cả năm đã qua): '),E.v(15,r+17,c+112)),X.CHAR(10.0)),'Năm tai vạ nặng: '),E.v(15,r+17,c+113))

def f11(E, r, c):
    # Ô C15
    # = IFERROR(LEFT(TRIM(MID(SUBSTITUTE(KET_LUAN!C179,CHAR(10),REPT(" ",2000)),1,2000)),FIND(" Nghề hợp dụng thần",TRIM(MID(SUBSTITUTE(KET_LUAN!C179,CHAR(10)
    #   ,REPT(" ",2000)),1,2000)))-1),TRIM(MID(SUBSTITUTE(KET_LUAN!C179,CHAR(10),REPT(" ",2000)),1,2000)))&IFERROR(CHAR(10)&MID(KET_LUAN!C179,FIND("⚠ Nếu đọc 
    #   theo",KET_LUAN!C179),FIND(CHAR(10),KET_LUAN!C179&CHAR(10),FIND("⚠ Nếu đọc theo",KET_LUAN!C179))-FIND("⚠ Nếu đọc theo",KET_LUAN!C179)),"")
    return X.cat(X.IFERROR(lambda: X.LEFT(X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r+164,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0)), X.sub(X.FIND(' Nghề hợp dụng thần', X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r+164,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),1.0)), lambda: X.TRIM(X.MID(X.SUBSTITUTE(E.v(3,r+164,c), X.CHAR(10.0), X.REPT(' ', 2000.0)), 1.0, 2000.0))),X.IFERROR(lambda: X.cat(X.CHAR(10.0),X.MID(E.v(3,r+164,c), X.FIND('⚠ Nếu đọc theo', E.v(3,r+164,c)), X.sub(X.FIND(X.CHAR(10.0), X.cat(E.v(3,r+164,c),X.CHAR(10.0)), X.FIND('⚠ Nếu đọc theo', E.v(3,r+164,c))),X.FIND('⚠ Nếu đọc theo', E.v(3,r+164,c))))), lambda: ''))

def f12(E, r, c):
    # Ô C16
    # = IFERROR(MID(KET_LUAN!C12,FIND("Độ tin vượng suy",KET_LUAN!C12),FIND("Sách nhấn mạnh",KET_LUAN!C12)-FIND("Độ tin vượng suy",KET_LUAN!C12)),"Xem KET_LUA
    #   N dòng 192")
    return X.IFERROR(lambda: X.MID(E.v(3,r-4,c), X.FIND('Độ tin vượng suy', E.v(3,r-4,c)), X.sub(X.FIND('Sách nhấn mạnh', E.v(3,r-4,c)),X.FIND('Độ tin vượng suy', E.v(3,r-4,c)))), lambda: 'Xem KET_LUAN dòng 192')

CELLS = {
    'C4': (f0, 4, 3),
    'C5': (f1, 5, 3),
    'C6': (f2, 6, 3),
    'C7': (f3, 7, 3),
    'C8': (f4, 8, 3),
    'C9': (f5, 9, 3),
    'C10': (f6, 10, 3),
    'C11': (f7, 11, 3),
    'C12': (f8, 12, 3),
    'C13': (f9, 13, 3),
    'C14': (f10, 14, 3),
    'C15': (f11, 15, 3),
    'C16': (f12, 16, 3),
}
