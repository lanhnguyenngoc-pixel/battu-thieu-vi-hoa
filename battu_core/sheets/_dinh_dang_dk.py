# -*- coding: utf-8 -*-
"""Định dạng có điều kiện – SINH TỰ ĐỘNG bởi tools/dich_dinh_dang_dk.py, KHÔNG SỬA TAY.
Nguồn: 6.10_G.xlsx (SHA-256 a34b9bd88b9052807174d240d6d3d83136557f1d078a7e3a3ba2d95e1eadc33f).
QUY_TAC: (sheet, [(c1, r1, c2, r2)…], ưu tiên, hàm điều kiện g(E, r, c), kiểu {dam, chu, nen})."""
from battu_core import xl as X

def g0(E, r, c):
    # KET_LUAN C36:D37: =1
    return 1.0

def g1(E, r, c):
    # LUAN_GIAI A71:C81: =$L71=1
    return X.eq(E.v(4,r,12),1.0)

def g2(E, r, c):
    # THEO_DOI A12:R111: =$A12=N(NHAP_LIEU!$C$31)
    return X.eq(E.v(5,r,1),X.N(E.v(2,31,3)))

def g3(E, r, c):
    # THEO_DOI G12:G111: =ISNUMBER(SEARCH("Tốt",$G12))
    return X.ISNUMBER(lambda: X.SEARCH('Tốt', E.v(5,r,7)))

def g4(E, r, c):
    # THEO_DOI G12:G111: =ISNUMBER(SEARCH("Xấu",$G12))
    return X.ISNUMBER(lambda: X.SEARCH('Xấu', E.v(5,r,7)))

def g5(E, r, c):
    # DAI_VAN C31:L44: =C$106=1
    return X.eq(E.v(14,106,c),1.0)

def g6(E, r, c):
    # DAI_VAN C54:L54: =C$106=1
    return X.eq(E.v(14,106,c),1.0)

def g7(E, r, c):
    # DAI_VAN C69:L69: =C$106=1
    return X.eq(E.v(14,106,c),1.0)

def g8(E, r, c):
    # DAI_VAN C102:L102: =AND(ISNUMBER(SEARCH("cát",C102)),NOT(ISNUMBER(SEARCH("hung",C102))))
    return X.AND(X.ISNUMBER(lambda: X.SEARCH('cát', E.v(14,r,c))), X.NOT(X.ISNUMBER(lambda: X.SEARCH('hung', E.v(14,r,c)))))

def g9(E, r, c):
    # LUU_NIEN A30:L129: =$CW30=1
    return X.eq(E.v(15,r,101),1.0)

def g10(E, r, c):
    # LUU_NIEN AA30:AA129: =$CW30=1
    return X.eq(E.v(15,r,101),1.0)

def g11(E, r, c):
    # LUU_NIEN AL30:AL129: =$CW30=1
    return X.eq(E.v(15,r,101),1.0)

def g12(E, r, c):
    # LUC_THAN X30:X34: =OR(X30="Tốt",X30="Khá")
    return X.OR(X.eq(E.v(16,r,c),'Tốt'), X.eq(E.v(16,r,c),'Khá'))

def g13(E, r, c):
    # HON_NHAN X64: =OR(X64="Tốt",X64="Khá")
    return X.OR(X.eq(E.v(17,r,c),'Tốt'), X.eq(E.v(17,r,c),'Khá'))

def g14(E, r, c):
    # MENH_CUC D18:D21: =OR(LEFT(D18,5)="THUẬN",LEFT(D18,13)="NGHIÊNG THUẬN",LEFT(D18,3)="TỐT",D18="Khá",ISNUMBER(SEARCH("VỀ GIÀU",D18)),ISNUMBER(SEARCH("VỀ SANG",D18)),ISNUMBER(SEARCH("VỀ CÁT",D18)),ISNUMBER(SEARCH("VỀ THỌ",D18)))
    return X.OR(X.eq(X.LEFT(E.v(18,r,c), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(18,r,c), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(18,r,c), 3.0),'TỐT'), X.eq(E.v(18,r,c),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(18,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(18,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(18,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(18,r,c))))

def g15(E, r, c):
    # SU_NGHIEP J6:J9: =OR(LEFT(J6,5)="THUẬN",LEFT(J6,13)="NGHIÊNG THUẬN",LEFT(J6,3)="TỐT",J6="Khá",ISNUMBER(SEARCH("VỀ GIÀU",J6)),ISNUMBER(SEARCH("VỀ SANG",J6)),ISNUMBER(SEARCH("VỀ CÁT",J6)),ISNUMBER(SEARCH("VỀ THỌ",J6)))
    return X.OR(X.eq(X.LEFT(E.v(21,r,c), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(21,r,c), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(21,r,c), 3.0),'TỐT'), X.eq(E.v(21,r,c),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(21,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(21,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(21,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(21,r,c))))

def g16(E, r, c):
    # SUC_KHOE J6:J8: =OR(LEFT(J6,5)="THUẬN",LEFT(J6,13)="NGHIÊNG THUẬN",LEFT(J6,3)="TỐT",J6="Khá",ISNUMBER(SEARCH("VỀ GIÀU",J6)),ISNUMBER(SEARCH("VỀ SANG",J6)),ISNUMBER(SEARCH("VỀ CÁT",J6)),ISNUMBER(SEARCH("VỀ THỌ",J6)))
    return X.OR(X.eq(X.LEFT(E.v(22,r,c), 5.0),'THUẬN'), X.eq(X.LEFT(E.v(22,r,c), 13.0),'NGHIÊNG THUẬN'), X.eq(X.LEFT(E.v(22,r,c), 3.0),'TỐT'), X.eq(E.v(22,r,c),'Khá'), X.ISNUMBER(lambda: X.SEARCH('VỀ GIÀU', E.v(22,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ SANG', E.v(22,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ CÁT', E.v(22,r,c))), X.ISNUMBER(lambda: X.SEARCH('VỀ THỌ', E.v(22,r,c))))

QUY_TAC = [
    ('KET_LUAN', [(3, 36, 4, 37)], 2, g0, {'dam': True, 'chu': '#c00000'}),
    ('LUAN_GIAI', [(1, 71, 3, 81)], 2, g1, {'dam': True, 'chu': '#c00000'}),
    ('THEO_DOI', [(1, 12, 18, 111)], 2, g2, {'nen': '#ffe699'}),
    ('THEO_DOI', [(7, 12, 7, 111)], 3, g3, {'nen': '#e2efda'}),
    ('THEO_DOI', [(7, 12, 7, 111)], 4, g4, {'nen': '#fbe4e1'}),
    ('DAI_VAN', [(3, 31, 12, 44)], 2, g5, {'dam': True, 'chu': '#c00000'}),
    ('DAI_VAN', [(3, 54, 12, 54)], 3, g6, {'dam': True, 'chu': '#c00000'}),
    ('DAI_VAN', [(3, 69, 12, 69)], 4, g7, {'dam': True, 'chu': '#c00000'}),
    ('DAI_VAN', [(3, 102, 12, 102)], 5, g8, {'dam': True, 'chu': '#c00000'}),
    ('LUU_NIEN', [(1, 30, 12, 129)], 2, g9, {'dam': True, 'chu': '#c00000'}),
    ('LUU_NIEN', [(27, 30, 27, 129)], 3, g10, {'dam': True, 'chu': '#c00000'}),
    ('LUU_NIEN', [(38, 30, 38, 129)], 4, g11, {'dam': True, 'chu': '#c00000'}),
    ('LUC_THAN', [(24, 30, 24, 34)], 2, g12, {'dam': True, 'chu': '#c00000'}),
    ('HON_NHAN', [(24, 64, 24, 64)], 2, g13, {'dam': True, 'chu': '#c00000'}),
    ('MENH_CUC', [(4, 18, 4, 21)], 2, g14, {'dam': True, 'chu': '#c00000'}),
    ('SU_NGHIEP', [(10, 6, 10, 9)], 2, g15, {'dam': True, 'chu': '#c00000'}),
    ('SUC_KHOE', [(10, 6, 10, 8)], 2, g16, {'dam': True, 'chu': '#c00000'}),
]
