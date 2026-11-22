# -*- coding: utf-8 -*-
"""Bản 6.9, mục 2.8 – bảng dấu hiệu con cái theo P2 tr.500–511 (nhóm B của bảng câu sách) + câu tổng hợp con cái.
Mỗi dấu hiệu là MỘT dòng ở LUC_THAN (mục E, từ dòng 40): C = điều kiện (1/0), D = nhóm, E = câu hiển thị, F = căn cứ trang.
Câu tổng hợp ở LUC_THAN!C37 (một câu) và danh sách dấu hiệu ở C38. KHÔNG dùng ngưỡng đếm dấu hiệu.
Không thêm dòng MAU_THUAN."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

S = 'LUC_THAN'
GIO_OK = 'NHAP_LIEU!$C$9<>"Không rõ giờ"'
NAM = 'NHAP_LIEU!$C$22="Nam"'
NU = 'NHAP_LIEU!$C$22<>"Nam"'
VUONG = 'DUNG_HY_KY!$C$155=1'
NHUOC = 'DUNG_HY_KY!$C$155=0'
NHT = 'DOI_CHIEU!$C$33=2'
DUNG = 'IF(%s,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)' % NHT
GS = 'INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!$F$6+1)'       # trường sinh của nhật can tại chi giờ
KD = 'INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1)'               # chi kình dương của nhật can
NAMC, THC, NGC, GIC = 'TRUNG_GIAN!$C$6', 'TRUNG_GIAN!$D$6', 'TRUNG_GIAN!$E$6', 'TRUNG_GIAN!$F$6'
CAN_GIO = 'TRUNG_GIAN!$F$12'
BASE = 'SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0'


def N(tts, pos=None, kind=None, extra=''):
    """Số chữ (hợp lệ, không phải nhật chủ) mang một trong các thập thần tts; lọc theo vị trí, loại (Can) và điều kiện thêm."""
    if isinstance(tts, str): tts = (tts,)
    out = []
    for t in tts:
        c = 'COUNTIFS(SAO!$F$5:$F$20,"%s",%s' % (t, BASE)
        if pos is not None: c += ',SAO!$B$5:$B$20,%d' % pos
        if kind: c += ',SAO!$C$5:$C$20,"%s"' % kind
        out.append(c + extra + ')')
    return '(' + '+'.join(out) + ')'


def K(lo, hi): return ',SAO!$K$5:$K$20,">=%d",SAO!$K$5:$K$20,"<=%d"' % (lo, hi)


QS = ('Chính quan', 'Thất sát'); TT = ('Thực thần', 'Thương quan'); AN = ('Chính ấn', 'Thiên ấn'); TAI = ('Chính tài', 'Thiên tài')
CHI = (NAMC, THC, NGC, GIC)
PAR = '(' + '+'.join('MOD(TRUNG_GIAN!$%s$%d,2)' % (c, r) for r in (5, 6) for c in 'CDEF') + ')'


def xung_voi(a, others): return 'OR(%s)' % ','.join('INDEX(M_XU,%s+1,%s+1)=1' % (a, o) for o in others)
def hinh_voi(a, others): return 'OR(%s)' % ','.join('INDEX(M_HINH,%s+1,%s+1)=1' % (a, o) for o in others)
def hai_voi(a, others): return 'OR(%s)' % ','.join('INDEX(M_HAI,%s+1,%s+1)=1' % (a, o) for o in others)


def chi_co(*idx):
    return 'AND(' + ','.join('COUNTIF(TRUNG_GIAN!$C$6:$F$6,%d)>0' % i for i in idx) + ')'


# Nhóm các dòng quy tắc CON CÁI đã có sẵn ở LUC_THAN (CC1–CC28 ở dòng 106–133; CC29–CC93 ở dòng 256–320; cột E = PP1, F = PP2, G = kết luận của sách).
G1 = (2, 5, 6, 23, 24, 25, 42, 43, 55, 59, 61, 64, 65, 70, 73, 76, 79, 81, 90, 92, 39)                        # ít con, khó có con
G2 = (1, 3, 4, 7, 19, 31, 36, 44, 48, 51)                                                                      # nhiều con
G3 = (10, 13, 14, 15, 18, 20, 21, 22, 26, 56, 60, 63, 66, 67, 69, 71, 72, 74, 75, 77, 84, 88, 89, 91, 93)       # khắc con, khó nuôi, con có trở ngại
G4 = (68, 80, 82, 83, 85, 87)                                                                                  # con nuôi, con khác họ
G5 = (3, 8, 9, 11, 12, 16, 27, 28, 29, 30, 32, 34, 35, 37, 38, 41, 45, 46, 47, 49, 50, 53, 54, 57, 58)         # con hiền, quý, thành đạt


def hang(n): return 105 + n if n <= 28 else 227 + n


def con_cai(p, why):
    from step20_bo_sung import LN_ROWS
    PP = 'DOI_CHIEU!$C$33=2'

    def co(n): return 'IF(%s,$F$%d,$E$%d)="ĐỦ"' % (PP, hang(n), hang(n))

    def nhom(ns): return 'OR(%s)' % ','.join(co(n) for n in ns)

    def ds(ns): return '&'.join('IF(%s,$G$%d&"; ","")' % (co(n), hang(n)) for n in ns)
    for r, t in ((36, 'Con cái – câu tổng hợp (P2 tr.500–511; P1 tr.157)'), (37, 'Con cái – các câu sách đang ứng với lá số'),
                 (38, 'Năm có dấu hiệu bất lợi cho CON (thương, thực gặp quan sát – P2 tr.507)')):
        p.set_text(S, 'B%d' % r, t, why=why, must_be_empty=True, style_from='B31')
    it, nhieu = nhom(G1), nhom(G2)
    so_con = ('IF(AND(%s,%s),IF(OR($X$34="Rất kém",$X$34="Kém"),"ít con hoặc khó có con",IF(OR($X$34="Khá",$X$34="Tốt"),"nhiều con","số con vừa phải (dấu hiệu trái chiều)")),IF(%s,"ít con hoặc khó có con",IF(%s,"nhiều con","số con vừa phải")))' % (it, nhieu, it, nhieu))
    def dem(ns): return '(' + '+'.join('IF(%s,1,0)' % co(n) for n in ns) + ')'
    nx, nt = dem(G3 + G4), dem(G5)       # đếm số câu sách ứng (đo 17–22/11/2026: gộp theo "có/không" thì 98% lá ra "có trở ngại" nên không còn phân biệt)
    nuoi = ('IF(%s-%s>=2,"có trở ngại, cần chú ý"&IF(%s>0," (nhưng cũng có dấu hiệu thuận bù lại)",""),IF(%s-%s>=2,"thuận, con hiền quý","bình thường"&IF(AND(%s>0,%s>0)," (có cả dấu hiệu thuận lẫn trở ngại)","")))' % (nx, nt, nt, nt, nx, nx, nt))
    traigai = '"con đầu lòng: "&$D$23&IF(%s," (nghiêng về – sách P1 tr.157 nói mệnh nữ thường không chuẩn)","")' % NU
    p.new_formula(S, 'C36', '"CON CÁI: "&%s&"; việc nuôi dạy "&%s&"; "&%s&"."' % (so_con, nuoi, traigai), text=True, style_from='C31', why=why)
    p.new_formula(S, 'C37', 'IF(LEN(%s)=0,"Không có câu sách nào ứng với lá số.","Các câu sách ứng với lá số: "&%s)' % (ds(G1 + G2 + G3 + G4 + G5), ds(G1 + G2 + G3 + G4 + G5)),
                  text=True, style_from='C31', why=why)
    # --- năm có dấu hiệu bất lợi cho CON: cột phụ DU, DV ở LUU_NIEN
    L = 'LUU_NIEN'
    p.set_text(L, 'DU29', 'Năm thương/thực gặp quan sát trong tứ trụ: tổn con (P2 tr.507)', why=why, must_be_empty=True, style_from='DN29')
    p.set_text(L, 'DV29', 'STT năm bất lợi cho con (phụ)', why=why, must_be_empty=True, style_from='DN29')
    for r in LN_ROWS:
        p.new_formula(L, 'DU%d' % r, 'IF(AND(%s>0,OR(O%d="Thương quan",O%d="Thực thần")),1,0)' % (N(QS), r, r), style_from='BT30', why=why)
        p.new_formula(L, 'DV%d' % r, 'IF(DU%d=1,COUNTIF($DU$30:DU%d,1),"")' % (r, r), style_from='BT30', why=why)
    parts = ['IFERROR(INDEX(LUU_NIEN!$A$30:$A$129,MATCH(%d,LUU_NIEN!$DV$30:$DV$129,0))&" ("&INDEX(LUU_NIEN!$C$30:$C$129,MATCH(%d,LUU_NIEN!$DV$30:$DV$129,0))&" tuổi mụ); ","")' % (k, k) for k in range(1, 11)]
    p.new_formula(S, 'C38', 'TRIM(' + '&'.join(parts) + ')', text=True, style_from='C31', why=why)


# ---------------------------------------------------------------------------------------------------------------
# Mục 2.11 điểm C (bản 6.9): điều 5 P2 tr.282 – can lưu niên ngũ hợp hóa với can tứ trụ thì xét vai trò theo hành hóa thần.
# Điều kiện hóa (P2 tr.336–337, thống nhất với PP2 và ngũ hợp tứ trụ): chỉ một can tứ trụ hợp với can lưu niên (không tranh hợp) và hóa thần
# cùng loại với hành nắm lệnh / hành chi tháng / nhân nguyên chi tháng. Hợp không hóa → giữ hành gốc (và vẫn tính "hợp chặt" như điều 3–4).
def hop_hoa_luu_nien(p, why):
    from step20_bo_sung import LN_ROWS
    L = 'LUU_NIEN'
    p.set_text(L, 'DW29', 'Điều 5 (P2 tr.282): hành hóa thần khi can lưu niên ngũ hợp hóa với can tứ trụ (−1 không hóa)', why=why, must_be_empty=True, style_from='DN29')
    p.set_text(L, 'DX29', 'Hành can lưu niên sau hợp hóa (dùng xét dụng/hỷ/kỵ)', why=why, must_be_empty=True, style_from='DN29')
    for r in LN_ROWS:
        m = ['INDEX(M_NH,TRUNG_GIAN!$%s$5+1,J%d+1)' % (c, r) for c in 'CDEF']
        hc = '+'.join('IF(%s>=0,1,0)' % x for x in m)
        hs = 'IF(%s>=0,%s,IF(%s>=0,%s,IF(%s>=0,%s,-1)))' % (m[0], m[0], m[1], m[1], m[3], m[3])
        cond = 'OR(%s=PP2_DINH_LUONG!$C$7,%s=TRUNG_GIAN!$D$9,%s=TRUNG_GIAN!$D$21,%s=TRUNG_GIAN!$D$22,%s=TRUNG_GIAN!$D$23)' % ((hs,) * 5)
        p.new_formula(L, 'DW%d' % r, 'IF(AND(%s>=0,%s=1,%s),%s,-1)' % (hs, '(' + hc + ')', cond, hs), style_from='BT30', why=why)
        p.new_formula(L, 'DX%d' % r, 'IF(DW%d>=0,DW%d,P%d)' % (r, r, r), style_from='BT30', why=why)
        for col in ('T', 'AE'):
            f = p.formula(L, '%s%d' % (col, r))
            assert 'P%d=' % r in f
            p.set_formula(L, '%s%d' % (col, r), f.replace('P%d=' % r, 'DX%d=' % r), why=why)
        for col in ('X', 'AI'):
            f = p.formula(L, '%s%d' % (col, r))
            for c, nm in (('C', 'Năm'), ('D', 'Tháng'), ('F', 'Giờ')):
                old = 'IF(INDEX(M_NH,TRUNG_GIAN!$%s$5+1,J%d+1)>=0,"%s hợp; ","")' % (c, r, nm)
                assert old in f, (col, r, nm)
                f = f.replace(old, 'IF(AND(DW%d<0,INDEX(M_NH,TRUNG_GIAN!$%s$5+1,J%d+1)>=0),"%s hợp; ","")' % (r, c, r, nm))
            p.set_formula(L, '%s%d' % (col, r), f, text=True, why=why)
        for col in ('BO', 'BP'):
            f = p.formula(L, '%s%d' % (col, r))
            p.set_formula(L, '%s%d' % (col, r), 'IF(DW%d>=0,"Điều 5: can năm "&INDEX(CAN_TEN,J%d+1)&" ngũ hợp hóa "&INDEX(HANH_TEN,DW%d+1)&" nên xét vai trò theo hành hóa thần (P2 tr.282 đ.5). ","")&' % (r, r, r) + f, text=True, why=why)
