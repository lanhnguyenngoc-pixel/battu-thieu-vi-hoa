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
    # con đầu lòng: hai quy tắc P2 tr.504 ngược nhau → nghiêng theo quy tắc can ngày – can giờ (tác giả kiểm nghiệm > 250 trường hợp; đúng 5/5 ví dụ tr.505–506), ghi chú quy tắc kia
    f23 = p.formula('LUC_THAN', 'D23'); cu = 'IF(OR(M23="",M23=L23),L23,"Chưa chắc")'
    assert cu in f23, f23
    p.set_formula('LUC_THAN', 'D23', f23.replace(cu, 'IF(OR(M23="",M23=L23),L23,L23&" (nghiêng về; quy tắc sao cha/mẹ chỉ ngược lại)")'), text=True, why=why)

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
    # đo trên 10 ví dụ P2 tr.505–512: kết luận thuận/trở ngại khớp 5/10 (không hơn đoán) → chỉ nêu số câu sách ứng, không kết luận
    nuoi = '"về việc nuôi dạy, sách có "&%s&" câu thuận và "&%s&" câu trở ngại ứng với lá số (chỉ tham khảo)"' % (nt, nx)
    traigai = '"con đầu lòng: "&$D$23&IF(%s," (nghiêng về – sách P1 tr.157 nói mệnh nữ thường không chuẩn)","")' % NU
    p.new_formula(S, 'C36', '"CON CÁI: "&%s&"; "&%s&"; "&%s&"."' % (so_con, nuoi, traigai), text=True, style_from='C31', why=why)
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


def hop_cuc_luu_nien(p, why):
    """Chi lưu niên hợp / hội cục với tứ trụ và đại vận (P2 tr.281-283 đ.5, 10, 12, 17; tr.254; P1 tr.62, 66-68).
    Chỉ GHI vào 'sự kiện của năm' (cột DY nối vào BN): không đổi xu thế tốt/xấu của năm."""
    from step20_bo_sung import LN_ROWS
    L = 'LUU_NIEN'
    p.set_text(L, 'DY29', 'Chi lưu niên hợp/hội cục với tứ trụ và đại vận (P2 tr.281-283 đ.5, 10, 12, 17; tr.254) – chỉ ghi, không đổi xu thế', why=why, must_be_empty=True, style_from='DN29')
    DU = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)'
    HY = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)'
    KY = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)'
    GIO = 'NHAP_LIEU!$C$9<>"Không rõ giờ"'
    CHI_TT = ['TRUNG_GIAN!$C$6', 'TRUNG_GIAN!$D$6', 'TRUNG_GIAN!$E$6', 'TRUNG_GIAN!$F$6']
    TRU = ['năm', 'tháng', 'ngày', 'giờ']
    HOA = [4, 3, 1, 0]   # tam hợp nhóm k=MOD(chi,4): Thủy, Kim, Hỏa, Mộc (0 Mộc,1 Hỏa,2 Thổ,3 Kim,4 Thủy)

    def chi(x):
        return 'INDEX(CHI_TEN,%s+1)' % x

    def hanh(x):
        return 'INDEX(HANH_TEN,%s+1)' % x

    def vai(h):
        return ('IF(%s<0,"",IF(OR(%s=%s,%s=%s)," – hành này là dụng/hỷ thần: tốt",IF(%s=%s," – hành này là kỵ thần: hung"," – hành này là nhàn thần")))'
                % (DU, h, DU, h, HY, h, KY))

    for r in LN_ROWS:
        K, G, P = 'K%d' % r, 'G%d' % r, 'P%d' % r

        def has_o(c):   # chi c có trong tứ trụ hoặc đại vận (không tính chi năm)
            return '(COUNTIF(TRUNG_GIAN!$C$6:$E$6,%s)+IF(AND(%s,TRUNG_GIAN!$F$6=%s),1,0)+IF(%s=%s,1,0)>0)' % (c, GIO, c, G, c)

        def has_s(c):   # có trong tứ trụ, đại vận hoặc là chi năm
            return '(%s*1+IF(%s=%s,1,0)>0)' % (has_o(c), K, c)

        parts = []
        # 1. tam hợp đủ ba chữ do chi năm hoàn thành
        c1, c2 = 'MOD(%s+4,12)' % K, 'MOD(%s+8,12)' % K
        h = 'CHOOSE(MOD(%s,4)+1,4,3,1,0)' % K
        parts.append('IF(AND(%s,%s),"• Chi năm "&%s&" cùng "&%s&", "&%s&" (tứ trụ/đại vận) hợp đủ ba chữ thành "&%s&" cục (tam hợp)"&%s&" (P2 tr.254; tr.281: lưu niên đưa hợp cục ra; P1 tr.67). ","")'
                     % (has_o(c1), has_o(c2), chi(K), chi(c1), chi(c2), hanh(h), vai(h)))
        # 2. tam hội đủ ba chữ
        u = 'INT(MOD(%s+1,12)/3)' % K
        pos = ['MOD(3*%s-1+%d,12)' % (u, j) for j in range(3)]
        full = '+'.join('IF(%s=%s,1,IF(%s,1,0))' % (x, K, has_o(x)) for x in pos)
        h2 = 'CHOOSE(%s+1,4,0,1,3)' % u
        parts.append('IF((%s)=3,"• Chi năm "&%s&" cùng hai chi còn lại hội đủ ba chữ thành "&%s&" cục (tam hội: lực lớn hơn tam hợp)"&%s&" (P2 tr.255; P1 tr.67). ","")'
                     % (full, chi(K), hanh(h2), vai(h2)))
        # 3. điều 12: hai tam hợp cục (sáu chi), có chi năm tham gia
        F = ['IF(AND(%s,%s,%s),1,0)' % (has_s(t), has_s('%d' % (t + 4)), has_s('%d' % (t + 8))) for t in range(4)]
        nf = '(%s)' % '+'.join(F)
        kh = 'OR(%s)' % ','.join('AND(%s=1,%s=1)' % (F[a], F[b]) for a, b in ((0, 2), (1, 2), (1, 3)))
        parts.append('IF(AND(%s>=2,CHOOSE(MOD(%s,4)+1,%s)=1),"• Tứ trụ, đại vận, lưu niên có sáu chi hợp thành hai tam hợp cục khác nhau"&IF(%s,", trong đó hai cục tương khắc nhau","")&": chủ về đại họa, cái chết khó tránh – chỉ là dấu hiệu cần đề phòng (P2 tr.282 đ.12). ","")'
                     % (nf, K, ','.join(F), kh))
        # 4. điều 17: đủ tí, ngọ, mão, dậu
        parts.append('IF(AND(OR(%s=0,%s=3,%s=6,%s=9),%s),"• Tứ trụ, đại vận, lưu niên có đủ bốn chữ tí, ngọ, mão, dậu: tốt (P2 tr.283 đ.17). ","")'
                     % (K, K, K, K, ','.join(has_s(str(c)) for c in (0, 3, 6, 9))))
        # 5. điều 10: bốn chi tứ trụ cùng lục hợp với chi năm
        n4 = '+'.join('IF(INDEX(M_LH,%s+1,%s+1)>=0,1,0)' % (CHI_TT[i], K) if i < 3 else 'IF(AND(%s,INDEX(M_LH,%s+1,%s+1)>=0),1,0)' % (GIO, CHI_TT[i], K) for i in range(4))
        parts.append('IF((%s)=4,"• Bốn chi trong tứ trụ cùng lục hợp với chi năm: tình cảm sóng gió, bệnh tật hoặc thất bại tài sản, kéo dài (P2 tr.282 đ.10). ","")' % n4)
        # 6. điều 5: chi năm lục hợp chi tứ trụ và hóa thần lộ ở can -> hợp hóa thành công
        for i in range(4):
            hh = 'INDEX(M_LH,%s+1,%s+1)' % (CHI_TT[i], K)
            lo = '(COUNTIF(TRUNG_GIAN!$C$7:$E$7,%s)+IF(AND(%s,TRUNG_GIAN!$F$7=%s),1,0)+IF(%s=%s,1,0)>0)' % (hh, GIO, hh, P, hh)
            cond = 'AND(%s>=0,%s)' % (hh, lo) if i < 3 else 'AND(%s,%s>=0,%s)' % (GIO, hh, lo)
            parts.append('IF(%s,"• Chi năm "&%s&" lục hợp chi "&%s&" (trụ %s) hóa "&%s&", hóa thần lộ ở can: nếu hợp hóa thành công thì lấy hóa thần xét cát hung"&%s&" (P2 tr.282 đ.5; P1 tr.62; sách không nêu lưu niên có cần kề sát hay không nên chỉ để tham khảo). ","")'
                         % (cond, chi(K), chi(CHI_TT[i]), TRU[i], hanh(hh), vai(hh)))
        # 7. tr.281 điểm 2: tứ trụ có hợp chưa hóa, can năm đưa hóa thần ra -> hóa cục
        for i in range(3):
            a, b = CHI_TT[i], CHI_TT[i + 1]
            hh = 'INDEX(M_LH,%s+1,%s+1)' % (a, b)
            ne = '(COUNTIF(TRUNG_GIAN!$C$7:$E$7,%s)+IF(AND(%s,TRUNG_GIAN!$F$7=%s),1,0)=0)' % (hh, GIO, hh)
            cond = 'AND(%s>=0,%s,%s=%s)' % (hh, ne, P, hh)
            if i == 2:
                cond = 'AND(%s,%s)' % (GIO, cond)
            parts.append('IF(%s,"• Tứ trụ có chi "&%s&" – "&%s&" (trụ %s, %s) lục hợp nhưng chưa hóa; can năm đưa hóa thần "&%s&" ra nên hợp cục thành hóa cục, khi tan cục sẽ có thay đổi lớn (P2 tr.281). ","")'
                         % (cond, chi(a), chi(b), TRU[i], TRU[i + 1], hanh(hh)))
        f = 'TRIM(%s)' % '&'.join(parts)
        p.new_formula(L, 'DY%d' % r, f, text=True, style_from='BN30', why=why)
        old = p.formula(L, 'BN%d' % r)
        tail = '&IF(LEN(CU%d)>0," "&CU%d,"")' % (r, r)
        assert old.endswith(tail), (r, old[-80:])
        p.set_formula(L, 'BN%d' % r, old + '&IF(LEN(DY%d)>0," "&DY%d,"")' % (r, r), text=True, why=why)
