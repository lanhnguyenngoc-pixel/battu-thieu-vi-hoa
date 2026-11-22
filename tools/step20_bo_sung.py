# -*- coding: utf-8 -*-
"""Phần bổ sung của bản 6.8 (gọi từ step20.py): mục 2.10 (Ngũ hành thiếu), 2.2a (sao con cái nam/nữ).
Mỗi hàm nhận Pkg p và đoạn why. Không thêm dòng MAU_THUAN."""

PA2 = 'LEFT(NGU_HANH_THIEU!$E$3,3)="PA2"'


def ngu_hanh_thieu(p, why):
    S = 'NGU_HANH_THIEU'
    # --- dòng 1: PA2 bỏ xét "giả"
    c7 = p.formula(S, 'C7'); i = c7.index('&IF(AND(' + PA2 + ',OR(TRUNG_GIAN!$D$5=4'); p.set_formula(S, 'C7', c7[:i], why=why)
    for ref, tail in (('D7', ',"Có","Không")'), ('E7', ',0.48,0)')):
        f = p.formula(S, ref); assert f.startswith('IF(AND(') and f.endswith(tail), ref
        j = f.index(',NOT(AND(' + PA2); x = f[len('IF(AND('):j]; assert x.endswith('<=4')
        p.set_formula(S, ref, 'IF(' + x + tail, why=why)
    # --- dòng 3: PA1 tính cả nhật chủ Mậu, Kỷ
    pre = 'IF(AND(TRUNG_GIAN!$E$7=2,NOT(' + PA2 + ')),"Nhật chủ Thổ – không xét (PA1, theo trang)",'
    c13 = p.formula(S, 'C13'); assert c13.startswith(pre) and c13.endswith(')'); p.set_formula(S, 'C13', c13[len(pre):-1], why=why)
    pre2 = 'IF(OR(AND(TRUNG_GIAN!$E$7=2,NOT(' + PA2 + ')),'
    for ref, tail, new in (('D13', '=""),"0","Có")', '="","0","Có")'), ('E13', '=""),0,0.2)', '="",0,0.2)')):
        f = p.formula(S, ref); assert f.startswith(pre2) and f.endswith(tail), ref
        p.set_formula(S, ref, 'IF(' + f[len(pre2):-len(tail)] + new, why=why)
    # --- câu cách tính
    p.set_text(S, 'F7', 'Cách tính: tra vòng trường sinh của CAN THÁNG tại CHI THÁNG (tháng theo tiết khí; Thổ theo Thủy: Mậu như Nhâm, Kỷ như Quý – bảng sách mục 2.2). '
               'Thuộc 5 trạng thái đầu (Trường sinh, Mộc dục, Quan đới, Lâm quan, Đế vượng) → được lệnh 48%; 7 trạng thái còn lại → 0. '
               'Cả hai chế độ (PA1, PA2) tính như nhau, không xét “giả” (bản 6.8, duyệt 18/11/2026). Không xét hợp.', why=why)
    p.set_text(S, 'F13', 'Cách tính: có ít nhất một CAN NĂM, CAN THÁNG hoặc CAN GIỜ CÙNG HÀNH với can ngày → 20% (tính một lần; trang tính ghi dòng này là “Xét được sinh”). '
               'Cả hai chế độ đều tính cả nhật chủ Mậu, Kỷ (bản 6.8). PA2: can hợp hóa tính theo hành đã hóa; can hợp mà không hóa vẫn giữ ngũ hành gốc nên vẫn tính '
               '(P2 tr.339, 351, 355 – thiên can theo P2; câu “ngũ hành đứng độc lập” của P1 tr.62 dùng cho địa chi). PA1 (theo trang): không xét hợp.', why=why)
    a2 = p.text_of(S, 'A2'); old = '(mục I khớp 48/49, mục II 47/48 lá số ảnh chụp trang tính)'; assert old in a2
    p.set_text(S, 'A2', a2.replace(old, '(mục II khớp 47/48 lá số ảnh chụp trang tính; mục I từ bản 6.8 tính cả nhật chủ Mậu, Kỷ nên ở các lá nhật chủ Thổ không còn tái tạo đúng trang)'), why=why)


def sao_con(p, why):
    S = 'LUC_THAN'
    ST = 'SAO!$F$5:$F$20,"%s",SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0'
    p.set_formula(S, 'G19', 'IF(NHAP_LIEU!$C$22="Nam","Thất sát, Thực thần (con trai); Chính quan, Thương quan (con gái)",'
                  '"Chính quan, Thương quan (con trai); Thất sát, Thực thần (con gái)")', why=why)
    p.set_formula(S, 'H19', '+'.join('COUNTIFS(%s)' % (ST % s) for s in ('Thất sát', 'Chính quan', 'Thực thần', 'Thương quan')), why=why)
    p.set_text(S, 'B20', 'Sao con: P1 tr.157 lấy thực thương là cách chung; P2 tr.504 chia trai/gái theo sao của cha (mệnh nam) và của mẹ (mệnh nữ). '
               '"Các chữ" ghi can gốc, chưa tính hợp hóa.', why=why)


COLS = 'CDEFGHIJKL'     # 10 đại vận


def dai_van(p, why):
    """Mục 2.4 (câu chữ 10 năm) + 2.5 (mức đại vận 6 bậc theo can – chi và chi vận với chi tứ trụ)."""
    S = 'DAI_VAN'
    DOI2 = 'DOI_CHIEU!$C$33=2'
    p.set_text(S, 'B108', 'F. KẾT LUẬN ĐẠI VẬN THEO THIỆU VĨ HOA – quy tắc chung (thuận/nghịch với thân, P1 tr.122–125): can vận làm chủ, chi vận bổ trợ tăng hay giảm; '
               'can chi tác dụng cả 10 năm (P1 tr.122; P2 tr.284) + quy tắc riêng của lá số', why=why)
    p.set_text(S, 'B103', 'Cách chia 5 năm đầu / 5 năm sau của Hóa giải toàn thư (HG tr.68) – chỉ tham khảo', why=why)
    p.set_text(S, 'B123', 'Vì sao – can vận (chủ, cả 10 năm)', why=why)
    p.set_text(S, 'B124', 'Vì sao – chi vận (bổ trợ, cả 10 năm)', why=why)
    labels = {130: 'Hóa thần của chi vận hợp với chi tứ trụ (−1 = không hóa; hóa thần phải lộ ở can – P1 tr.62)',
              131: 'Hành của chi vận sau hợp hóa', 132: 'Vai trò hành chi vận sau hợp hóa (1 dụng/hỷ, −1 kỵ, 0 nhàn)',
              133: 'Chi vận xung chữ tứ trụ (−1 xung mất dụng/hỷ, +1 xung mất kỵ, 0 không)', 134: 'Chiều của chi vận sau hợp hóa và xung',
              135: 'Yếu tố HẠ mức (chi khắc can, can khắc chi, can sinh chi, chi trái chiều can)', 136: 'Yếu tố TĂNG mức (chi sinh phù can, chi cùng chiều can)',
              137: 'Mức đại vận (0 vừa, 1 thường, 2 rất)', 138: 'Giải thích mức đại vận (P1 tr.122–124; P2 tr.282)', 139: 'Quan hệ can – chi sau hợp hóa (0 cùng hành, 1 can sinh chi, 2 chi sinh can, 3 can khắc chi, 4 chi khắc can)'}
    for r, t in labels.items():
        p.set_text(S, 'B%d' % r, t, style_from='B123', why=why, must_be_empty=True)
    for c in COLS:
        X = lambda r: '%s%d' % (c, r)
        # hóa thần: chi vận hợp lục hợp với từng chi tứ trụ (năm, tháng, ngày, giờ), hóa thần lộ ở can
        def hop(k):
            h = 'INDEX(M_LH,TRUNG_GIAN!$%s$6+1,%s$34+1)' % (k, c)
            lo = 'OR(TRUNG_GIAN!$C$7=%s,TRUNG_GIAN!$D$7=%s,TRUNG_GIAN!$E$7=%s,TRUNG_GIAN!$F$7=%s,%s$42=%s)' % ((h,) * 4 + (c, h))
            return h, lo
        parts = []
        for k in 'CDEF':
            h, lo = hop(k)
            cond = 'AND(%s>=0,%s)' % (h, lo) if k != 'F' else 'AND(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),%s>=0,%s)' % (h, lo)
            parts.append((cond, h))
        f = '-1'
        for cond, h in reversed(parts): f = 'IF(%s,%s,%s)' % (cond, h, f)
        p.new_formula(S, X(130), f, why=why, style_from=X(119))
        p.new_formula(S, X(131), 'IF(%s>=0,%s,%s43)' % (X(130), X(130), c), why=why, style_from=X(119))
        def role(col): return ('IF(OR(AND(DUNG_HY_KY!$%s$109>=0,%s=DUNG_HY_KY!$%s$109),AND(DUNG_HY_KY!$%s$110>=0,%s=DUNG_HY_KY!$%s$110)),1,'
                               'IF(AND(DUNG_HY_KY!$%s$111>=0,%s=DUNG_HY_KY!$%s$111),-1,0))' % (col, X(131), col, col, X(131), col, col, X(131), col))
        p.new_formula(S, X(132), 'IF(%s,%s,%s)' % (DOI2, role('D'), role('C')), why=why, style_from=X(119))
        p.new_formula(S, X(133), 'IF(LEN(IF(%s,%s,%s))>0,-1,IF(LEN(IF(%s,%s,%s))>0,1,0))' % (DOI2, X(67), X(52), DOI2, X(66), X(51)), why=why, style_from=X(119))
        p.new_formula(S, X(134), 'IF(%s<>0,%s,%s)' % (X(133), X(133), X(132)), why=why, style_from=X(119))
        p.new_formula(S, X(139), 'IF(%s42=%s,0,IF(MOD(%s42+1,5)=%s,1,IF(MOD(%s+1,5)=%s42,2,IF(MOD(%s42+2,5)=%s,3,4))))' % (c, X(131), c, X(131), X(131), c, c, X(131)), why=why, style_from=X(119))
        sg = 'SIGN(%s)' % X(119)
        p.new_formula(S, X(135), 'IF(%s=0,"",IF(OR(%s=1,%s=3,%s=4,AND(%s<>0,%s=-%s)),1,0))' % (X(119), X(139), X(139), X(139), X(134), X(134), sg), why=why, style_from=X(119))
        p.new_formula(S, X(136), 'IF(%s=0,"",IF(%s=1,0,IF(OR(%s=0,%s=2,AND(%s<>0,%s=%s)),1,0)))' % (X(119), X(135), X(139), X(139), X(134), X(134), sg), why=why, style_from=X(119))
        b = 'IF(ABS(%s)>=2,2,1)' % X(119)
        p.new_formula(S, X(137), 'IF(%s=0,"",IF(%s=1,MAX(0,%s-1),IF(%s=1,MIN(2,%s+1),%s)))' % (X(119), X(135), b, X(136), b, b), why=why, style_from=X(119))
        old = p.formula(S, X(121))
        new = 'IF(%s=0,%s,IF(%s>0,CHOOSE(%s+1,"TỐT VỪA","TỐT","RẤT TỐT"),CHOOSE(%s+1,"XẤU VỪA","XẤU","RẤT XẤU")))' % (X(119), old, X(119), X(137), X(137))
        p.set_formula(S, X(121), new, why=why)
        p.set_formula(S, X(122), '"cả 10 năm "&LOWER(%s)' % X(121), why=why)
        # giải thích: MỨC – can ...; chi ...; quan hệ → lý do
        # đơn giản hóa: câu giải thích tính bằng các ô trợ giúp, không lặp lại điều kiện
        quan_he = 'CHOOSE(%s+1,"can chi cùng hành","can sinh chi (can xì hơi)","chi sinh can (sinh phù)","can khắc chi (can hao)","chi khắc can (áp chế can)")' % X(139)
        sentence = ('IF(%s=0,%s,%s&" – can "&INDEX(CAN_TEN,%s$33+1)&" ("&%s$39&", hành "&INDEX(HANH_TEN,%s$42+1)&") làm chủ cả 10 năm; chi "&INDEX(CHI_TEN,%s$34+1)&'
                    'IF(%s>=0,", hợp hóa thành "&INDEX(HANH_TEN,%s+1),", hành "&INDEX(HANH_TEN,%s$43+1))&IF(%s>0," – cùng chiều can",IF(%s<0," – trái chiều can",""))&'
                    'IF(%s=-1,"; xung mất chữ mang dụng/hỷ",IF(%s=1,"; xung mất chữ mang kỵ",""))&"; "&%s&" → "&'
                    'IF(%s=1,"cái "&IF(%s>0,"tốt","xấu")&" giảm một mức",IF(%s=1,"cái "&IF(%s>0,"tốt","xấu")&" tăng thêm","giữ nguyên mức"))&" (P1 tr.122–124; P2 tr.282).")'
                    % (X(119), X(121) + '&" – can vận nhàn thần, xét theo chi vận (P1 tr.123)."', X(121), c, c, c, c, X(130), X(130), c, X(134), X(134), X(133), X(133), quan_he,
                       X(135), X(119), X(136), X(119)))
        # sửa: INDEX(CAN_TEN,%s$33+1) cần ký hiệu cột
        p.new_formula(S, X(138), sentence.replace('%s$', c + '$'), text=True, why=why, style_from=X(123))
    # câu "Căn cứ sách" và các chỗ đọc "5 năm đầu/sau"
    K = 'KET_LUAN'
    f = p.formula(K, 'C163'); old = '". 5 năm đầu: "&INDEX(DAI_VAN!$C$123:$L$123,DAI_VAN!$C$25)&" 5 năm sau: "&INDEX(DAI_VAN!$C$124:$L$124,DAI_VAN!$C$25)'
    assert old in f; p.set_formula(K, 'C163', f.replace(old, '". Can vận: "&INDEX(DAI_VAN!$C$123:$L$123,DAI_VAN!$C$25)&" Chi vận: "&INDEX(DAI_VAN!$C$124:$L$124,DAI_VAN!$C$25)'), why=why)
    f = p.formula(K, 'C175'); old = '&INDEX(DAI_VAN!$C$103:$L$103,MIN(10,DAI_VAN!$C$25))&" "'
    assert old in f; p.set_formula(K, 'C175', f.replace(old, '&INDEX(DAI_VAN!$C$138:$L$138,MIN(10,DAI_VAN!$C$25))&" "'), why=why)
    # dòng 94: gộp TỐT VỪA vào "tốt", XẤU VỪA vào "xấu"
    import re
    f = p.formula(S, 'C94')
    g = re.sub(r'([C-L])121="(TỐT|XẤU)"', r'LEFT(\g<1>121,3)="\2"', f)
    assert g != f; p.set_formula(S, 'C94', g, why=why)


def dieu_hau(p, why, moc_ha=True):
    """Mục 2.6: hai loại điều hầu – câu chữ; Mộc sinh mùa hạ cần Thủy (duyệt 17/11/2026, đo trước khi chốt)."""
    S = 'DUNG_HY_KY'
    if moc_ha:
        f = p.formula(S, 'C215'); old = 'OR(TRUNG_GIAN!$E$7=1,TRUNG_GIAN!$E$7=2,TRUNG_GIAN!$E$7=3)'; assert old in f
        p.set_formula(S, 'C215', f.replace(old, 'OR(TRUNG_GIAN!$E$7=0,TRUNG_GIAN!$E$7=1,TRUNG_GIAN!$E$7=2,TRUNG_GIAN!$E$7=3)'), why=why)
        p.set_text(S, 'B215', 'Hành điều hầu cần (P1 tr.76–79: kim, thủy, mộc sinh tháng Hợi–Tí–Sửu–Dần cần hỏa; mộc, hỏa, thổ, kim sinh tháng Tị–Ngọ–Mùi cần thủy; −1 = không)', why=why)
        c = p.formula(S, 'C219'); assert c.endswith('")')
        p.set_formula(S, 'C219', c[:-2] + '; P1 tr.76 (mộc mùa hạ rễ lá khô, cần nước tưới)")', why=why)
    f = p.formula(S, 'C171'); old = '&" Dụng thần sinh phù làm chủ, điều hầu là bổ trợ (P1 tr.100, 111-113).")'; assert f.endswith(old)
    new = ('&IF($C$215<0," Đây là dụng thần bổ trợ theo khí hậu mệnh cục, KHÔNG thay dụng thần và không đổi xu thế vận (P1 tr.100, 111–113)."'
           '," Có xét thay dụng thần và vận điều hầu chuyển tốt (P1 tr.76–79).")&" Dụng thần sinh phù làm chủ, điều hầu là bổ trợ (P1 tr.100, 111-113).")')
    p.set_formula(S, 'C171', f[:-len(old)] + new, why=why)


# ---------------------------------------------------------------------------------------------------------------
# Giai đoạn 2 (bản 6.9): LUU_NIEN – mục 2.3 (điều 6, 7, 8 P2 tr.288–289) và 2.11 (xu thế năm: hình, hợp mất kỵ, gốc Thường giữ Thường)
LN_ROWS = range(30, 130)
TRU = (('C', 'Năm'), ('D', 'Tháng'), ('E', 'Ngày'), ('F', 'Giờ'))


def _vt(c, h):
    """Vai trò của hành h theo cột dụng–hỷ–kỵ c (C: PP1, D: PP2) của DUNG_HY_KY."""
    return ('IF(AND(DUNG_HY_KY!$%s$109>=0,%s=DUNG_HY_KY!$%s$109),"dụng",IF(AND(DUNG_HY_KY!$%s$110>=0,%s=DUNG_HY_KY!$%s$110),"hỷ",'
            'IF(AND(DUNG_HY_KY!$%s$111>=0,%s=DUNG_HY_KY!$%s$111),"kỵ","không phải dụng/hỷ/kỵ")))' % (c, h, c, c, h, c, c, h, c))


def _xung_mat(r, c, mat):
    """Chữ trong tứ trụ bị chi lưu niên xung/hình + khắc hành (mat = 'ky' → chữ kỵ, 'dung' → chữ dụng/hỷ); với 'ky' thêm trường hợp chi lưu niên HỢP chữ kỵ (P2 tr.282 đ.1)."""
    kq = []
    for t, ten in TRU:
        chi = 'TRUNG_GIAN!$%s$6' % t; hc = 'TRUNG_GIAN!$%s$9' % t; vt = _vt(c, hc)
        xh = 'OR(INDEX(M_XU,K%d+1,%s+1)=1,INDEX(M_HINH,K%d+1,%s+1)=1)' % (r, chi, r, chi)
        kh = 'NOT(AND(Q%d=2,%s=2)),MOD(Q%d+2,5)=%s' % (r, hc, r, hc)
        if mat == 'ky':
            dk = 'OR(AND(%s,%s,%s="kỵ"),AND(INDEX(M_LH,%s+1,K%d+1)>=0,%s="kỵ"))' % (xh, kh, vt, chi, r, vt)
        else:
            dk = 'AND(%s,%s,OR(%s="dụng",%s="hỷ"))' % (xh, kh, vt, vt)
        kq.append('IF(%s,"%s ","")' % (dk, ten))
    return 'TRIM(' + '&'.join(kq) + ')'


def luu_nien(p, why):
    S = 'LUU_NIEN'; NHT = 'DOI_CHIEU!$C$33=2'
    p.set_text(S, 'DO29', 'Điều 6 (P2 tr.288): thân nhược gặp can năm hung sát khắc hại nhật nguyên/dụng thần', why=why, must_be_empty=True, style_from='DN29')
    p.set_text(S, 'DP29', 'Điều 7 (P2 tr.288): tuế vận tịnh lâm – 1 thất sát/kình dương (hung), 2 tài/quan/ấn (tốt), 3 còn lại (lớn nhưng không nhất định)', why=why, must_be_empty=True, style_from='DN29')
    p.set_text(S, 'DQ29', 'Điều 8 (P2 tr.289): chi mang dụng thần lục hợp chi lưu niên – dụng thần hợp thái tuế', why=why, must_be_empty=True, style_from='DN29')
    dung = 'IF(%s,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)' % NHT
    nhuoc = 'IF(%s,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0' % NHT
    for r in LN_ROWS:
        vtn = 'IF(%s,AE%d,T%d)' % (NHT, r, r)
        p.new_formula(S, 'DO%d' % r, 'IF(AND(%s,OR(O%d="Kiếp tài",O%d="Thất sát",O%d="Thương quan",O%d="Thiên ấn"),NOT(OR(%s="dụng",%s="hỷ")),'
                      'OR(MOD(P%d+2,5)=TRUNG_GIAN!$E$7,AND(%s>=0,MOD(P%d+2,5)=%s))),1,0)' % (nhuoc, r, r, r, r, vtn, vtn, r, dung, r, dung), style_from='DM30', why=why)
        p.new_formula(S, 'DP%d' % r, 'IF(AP%d=1,IF(OR(O%d="Thất sát",K%d=INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1)),1,'
                      'IF(OR(O%d="Chính tài",O%d="Thiên tài",O%d="Chính quan",O%d="Chính ấn",O%d="Thiên ấn"),2,3)),0)' % (r, r, r, r, r, r, r, r), style_from='DM30', why=why)
        hop = ','.join('AND(INDEX(M_LH,TRUNG_GIAN!$%s$6+1,K%d+1)>=0,%s>=0,TRUNG_GIAN!$%s$9=%s)' % (t, r, dung, t, dung) for t, _ in TRU)
        p.new_formula(S, 'DQ%d' % r, 'IF(OR(%s),1,0)' % hop, style_from='DM30', why=why)
        # --- 2.11: xung mất kỵ / mất dụng-hỷ có thêm hình (và hợp mất kỵ); gốc Thường giữ Thường (điều 1–2 chỉ đổi gốc Xấu → Tốt, Tốt → Xấu)
        for ck, cd, cy, cu, cx, cc, dc in (('V', 'W', 'Y', 'U', 'X', 'C', 'T'), ('AG', 'AH', 'AJ', 'AF', 'AI', 'D', 'AE')):
            p.set_formula(S, '%s%d' % (ck, r), _xung_mat(r, cc, 'ky'), why=why)
            p.set_formula(S, '%s%d' % (cd, r), _xung_mat(r, cc, 'dung'), why=why)
            V, W, U, X = '%s%d' % (ck, r), '%s%d' % (cd, r), '%s%d' % (cu, r), '%s%d' % (cx, r)
            p.set_formula(S, '%s%d' % (cy, r), 'IF(AND(LEN(%s)>0,LEN(%s)=0),IF(%s="Tốt","Xấu",%s),IF(AND(LEN(%s)>0,LEN(%s)=0),IF(%s="Xấu","Tốt",%s),'
                          'IF(AND(%s="Tốt",LEN(%s)>0),"Thường",IF(AND(%s="Xấu",LEN(%s)>0),"Thường",%s))))' % (W, V, U, U, V, W, U, U, U, X, U, X, U), why=why)
        # --- BM, DI, DM: tuế vận gặp nhau mà can năm là tài/quan/ấn là tốt, không tính dấu hiệu nặng
        p.replace_in_formula(S, 'BM%d' % r, 'AP%d+AQ%d' % (r, r), 'IF(OR(DP%d=1,DP%d=3),1,0)+AQ%d' % (r, r, r), why=why)
        p.replace_in_formula(S, 'DI%d' % r, 'AP%d=1' % r, 'AND(AP%d=1,DP%d<>2)' % (r, r), why=why)
        p.replace_in_formula(S, 'DM%d' % r, 'IF(AP%d=1,"người thân; ","")' % r, 'IF(AND(AP%d=1,DP%d<>2),"người thân; ","")' % (r, r), why=why)
        # --- câu sự kiện BN
        ap = 'IF(AP%d=1,"• Tuế vận tịnh lâm với đại vận: ""không chết mình cũng mất người thân"" (P2 tr.288, 566). ","")' % r
        p.replace_in_formula(S, 'BN%d' % r, ap,
                             'IF(DP%d=1,"• Tuế vận tịnh lâm với đại vận, can năm là thất sát hoặc chi năm là kình dương: ""không chết mình cũng mất người thân"" (P2 tr.288, 566). ","")'
                             '&IF(DP%d=2,"• Tuế vận tịnh lâm với đại vận, can năm là tài, quan, ấn: tuế vận gặp nhau là tốt (P2 tr.288). ","")'
                             '&IF(DP%d=3,"• Tuế vận tịnh lâm với đại vận: tai vạ lớn nhưng không nhất định (P2 tr.288). ","")' % (r, r, r), why=why)
        for ad, vt, pp, ben, tr in (('AD', 'T', 'PP1', ': đề phòng bệnh', '559, 562'), ('AO', 'AE', 'PP2', '', '559')):
            old = 'IF(%s%d=1,"• %s: thân nhược gặp năm thất sát%s (P2 tr.%s). ","")' % (ad, r, pp, ben, tr)
            new = ('IF(%s%d=1,"• %s: thân nhược gặp năm thất sát%s"&IF(OR(%s%d="dụng",%s%d="hỷ")," – thất sát là dụng thần – nhẹ","")&" (P2 tr.%s). ","")'
                   % (ad, r, pp, ben, vt, r, vt, r, tr))
            p.replace_in_formula(S, 'BN%d' % r, old, new, why=why)
        # thêm đúng trước câu "Lưu niên hợp trụ ngày"
        ba = 'IF(BA%d=1,"• Lưu niên hợp trụ ngày: tin mừng (P2 tr.289). ","")' % r
        extra = ('IF(DO%d=1,"• Hung sát khắc hại nhật nguyên/dụng thần: đề phòng tai vạ (P2 tr.288). ","")'
                 '&IF(DQ%d=1,"• Dụng thần hợp thái tuế: gặp may (P2 tr.289). ","")'
                 '&IF(AND(IF(%s,AF%d,U%d)="Thường",LEN(IF(%s,AH%d,W%d))>0,LEN(IF(%s,AG%d,V%d))=0),"• Chữ mang dụng/hỷ ("&IF(%s,AH%d,W%d)&") bị xung/hình khắc: gốc năm là Thường nên xu thế không đổi, nhưng có việc đáng đề phòng (P2 tr.282 đ.2). ","")'
                 '&IF(AND(IF(%s,AF%d,U%d)="Thường",LEN(IF(%s,AG%d,V%d))>0,LEN(IF(%s,AH%d,W%d))=0),"• Chữ mang kỵ ("&IF(%s,AG%d,V%d)&") bị xung/hình/hợp: gốc năm là Thường nên xu thế không đổi, có việc nhẹ bớt (P2 tr.282 đ.1). ","")'
                 % ((r, r) + (NHT, r, r, NHT, r, r, NHT, r, r, NHT, r, r) + (NHT, r, r, NHT, r, r, NHT, r, r, NHT, r, r)))
        p.replace_in_formula(S, 'BN%d' % r, ba, extra + '&' + ba, why=why)
        # --- diễn giải xu thế năm (BO: PP1, BP: PP2) ghi lý do chỉnh
        for bo, vt, u, y, v, w, x in (('BO', 'T', 'U', 'Y', 'V', 'W', 'X'), ('BP', 'AE', 'AF', 'AJ', 'AG', 'AH', 'AI')):
            U, Y, V, W, X = u + str(r), y + str(r), v + str(r), w + str(r), x + str(r)
            old = 'IF(%s<>%s," → điều chỉnh "&%s,"")' % (Y, U, Y)
            ly = ('IF(AND(LEN(%s)>0,LEN(%s)=0),"điều 2 – chữ mang dụng/hỷ bị xung/hình khắc ("&%s&"): tốt thành xấu",IF(AND(LEN(%s)>0,LEN(%s)=0),'
                  '"điều 1 – chữ mang kỵ bị xung/hình/hợp ("&%s&"): xấu thành tốt","điều 3–4 – can năm bị chế hoặc hợp ("&%s&"): thành bình thường"))' % (W, V, W, V, W, V, X))
            new = ('IF(%s<>%s," → điều chỉnh "&%s&" ("&%s&")","")&IF(AND(%s=%s,%s="Thường",OR(LEN(%s)>0,LEN(%s)>0))," – gốc Thường giữ Thường: chữ "&IF(LEN(%s)>0,"dụng/hỷ ("&%s&")","kỵ ("&%s&")")&" bị tác động, xem mục sự kiện (P2 tr.282 đ.1–2)","")'
                   % (Y, U, Y, ly, Y, U, U, V, W, W, W, V))
            p.replace_in_formula(S, '%s%d' % (bo, r), old, new, why=why)


# ---------------------------------------------------------------------------------------------------------------
# Mục 2.2 (bản 6.9): người ứng – cung làm chính + sao lục thân trong chữ bị chạm + dụng/kỵ + cứu + nền xu thế, MỘT câu kết luận
PTEN = ((0, 'Năm', 'ông bà, cha mẹ'), (1, 'Tháng', 'anh chị em'), (2, 'Ngày', None), (3, 'Giờ', 'con cái'))


def _cp(r, pos, ten):
    """Trụ pos bị chạm: có trong AR (thiên khắc địa xung) hoặc chi trụ bị chi lưu niên xung/hình."""
    t = 'CDEF'[pos]
    return ('OR(ISNUMBER(SEARCH("%s",AR%d)),INDEX(M_XU,K%d+1,TRUNG_GIAN!$%s$6+1)=1,INDEX(M_HINH,K%d+1,TRUNG_GIAN!$%s$6+1)=1)' % (ten, r, r, t, r, t))


def _sao(pos):
    """Chuỗi sao lục thân có trong trụ pos (can lộ + tàng can), nhãn theo giới (P1 tr.149, 157; P2 tr.504)."""
    def n(tt): return 'COUNTIFS(SAO!$B$5:$B$20,%d,SAO!$F$5:$F$20,"%s",SAO!$Y$5:$Y$20,0)>0' % (pos, tt)
    NAM = 'NHAP_LIEU!$C$22="Nam"'
    parts = ['IF(%s,"cha (thiên tài), ","")' % n('Thiên tài'), 'IF(%s,"mẹ (chính ấn), ","")' % n('Chính ấn'),
             'IF(%s,"dì hoặc mẹ kế (thiên ấn), ","")' % n('Thiên ấn'), 'IF(OR(%s,%s),"anh chị em (tỉ kiếp), ","")' % (n('Tỉ kiên'), n('Kiếp tài')),
             'IF(AND(%s,%s),"vợ (chính tài), ","")' % (NAM, n('Chính tài'))]
    for tt, trai_nam in (('Thất sát', 'con trai'), ('Thực thần', 'con trai'), ('Chính quan', 'con gái'), ('Thương quan', 'con gái')):
        parts.append('IF(AND(%s,%s),"%s (%s), ","")' % (NAM, n(tt), trai_nam, tt.lower()))
    for tt, nhan in (('Chính quan', 'con trai'), ('Thương quan', 'con trai'), ('Thất sát', 'con gái'), ('Thực thần', 'con gái')):
        if tt in ('Chính quan', 'Thất sát'):
            if pos == 3: lab = '%s (%s)' % (nhan, tt.lower())
            elif pos == 2: lab = 'chồng (%s)' % tt.lower()
            else: lab = 'chồng hoặc %s (%s)' % (nhan, tt.lower())
        else: lab = '%s (%s)' % (nhan, tt.lower())
        parts.append('IF(AND(NOT(%s),%s),"%s, ","")' % (NAM, n(tt), lab))
    return '&'.join(parts)


def nguoi_ung(p, why):
    S = 'LUU_NIEN'; NHT = 'DOI_CHIEU!$C$33=2'; NAM = 'NHAP_LIEU!$C$22="Nam"'
    for c, t in (('DR', 'Người ứng theo cung của trụ bị chạm (P1 tr.149; P2 tr.287)'), ('DS', 'Sao lục thân trong chữ bị chạm (P1 tr.149, 157; P2 tr.504)'),
                 ('DT', 'KẾT LUẬN người ứng với hạn của năm (một câu: cung + sao + dụng/kỵ + cứu + nền xu thế)')):
        p.set_text(S, c + '29', t, why=why, must_be_empty=True, style_from='DN29')
    dung = 'IF(%s,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)' % NHT; hy = 'IF(%s,DUNG_HY_KY!$D$110,DUNG_HY_KY!$C$110)' % NHT; ky = 'IF(%s,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)' % NHT

    def vt_in(h, ds):
        return 'OR(%s)' % ','.join('AND(%s>=0,%s=%s)' % (x, h, x) for x in ds)
    for r in LN_ROWS:
        cp = [_cp(r, pos, ten) for pos, ten, _ in PTEN]
        ng = []
        for pos, ten, nd in PTEN:
            lab = nd if nd else 'bản thân và "&IF(%s,"vợ","chồng")&"' % NAM
            ng.append('IF(%s,"%s; ","")' % (cp[pos], lab))
        p.new_formula(S, 'DR%d' % r, 'TRIM(' + '&'.join(ng) + ')', style_from='DM30', why=why)
        # DR có thể ghi cả chuỗi vợ/chồng nên dùng công thức riêng cho trụ ngày
        ds = ['IF(%s,%s,"")' % (cp[pos], _sao(pos)) for pos, _, _ in PTEN]
        p.new_formula(S, 'DS%d' % r, 'TRIM(' + '&'.join(ds) + ')', style_from='DM30', why=why)
        mang_dh = 'OR(%s)' % ','.join('AND(%s,OR(%s,%s))' % (cp[pos], vt_in('TRUNG_GIAN!$%s$7' % 'CDEF'[pos], (dung, hy)), vt_in('TRUNG_GIAN!$%s$9' % 'CDEF'[pos], (dung, hy))) for pos in range(4))
        mang_ky = 'OR(%s)' % ','.join('AND(%s,OR(%s,%s))' % (cp[pos], vt_in('TRUNG_GIAN!$%s$7' % 'CDEF'[pos], (ky,)), vt_in('TRUNG_GIAN!$%s$9' % 'CDEF'[pos], (ky,))) for pos in range(4))
        any_cp = 'OR(%s)' % ','.join(cp)
        cuu = ('IF(LEN(CB%d)>0,"Có cứu ("&CB%d&") nên nhẹ, có thể qua khỏi (P2 tr.287–288). ","Không thấy cứu nên nặng hơn; nếu bản thân không gặp thì chuyển lên người thân (P2 tr.287–288). ")' % (r, r))
        nen = ('"Nền xu thế: năm "&IF(%s,AJ%d,Y%d)&", đại vận "&IF(Z%d="","chưa vào",IF(%s,AK%d,Z%d))&" (P1 tr.133)."' % (NHT, r, r, r, NHT, r, r))
        cd = 'IF(OR(%s,%s)," Người ứng là vợ/chồng hoặc con cái: nếu đã có thì ứng vào họ, nếu chưa thì ứng vào bản thân (P1 tr.149, 156).","")' % (cp[2], cp[3])
        f = ('IF(%s,A%d&" (tuổi "&C%d&"): đề phòng việc không hay đến với "&LEFT(DR%d,LEN(DR%d)-1)&IF(LEN(DS%d)>0,", chữ bị tác động mang sao "&LEFT(DS%d,LEN(DS%d)-1),"")&". "'
             '&IF(%s,"Chữ bị tác động mang dụng/hỷ thần nên đáng lưu tâm (P2 tr.282 đ.2). ",IF(%s,"Chữ bị tác động là kỵ thần nên cái hại giảm, có khi thành phúc (P2 tr.282 đ.1). ",""))&%s&%s&%s,"")'
             % (any_cp, r, r, r, r, r, r, r, mang_dh, mang_ky, cuu, nen, cd))
        p.new_formula(S, 'DT%d' % r, f, text=True, style_from='DM30', why=why)


# ---------------------------------------------------------------------------------------------------------------
# Mục 2.7 (bản 6.9): điểm 1 (cách kình dương, P1 tr.117) và điểm 3 (thân vượng thiếu cả quan sát và thực thương, P1 tr.108).
# Điểm 2 (quan sát hỗn tạp cứu bằng hợp) và điểm 4 (tham hợp quên xung) chưa làm – xem báo cáo.
def bo_sung_27(p, why):
    NHT = 'DOI_CHIEU!$C$33=2'
    dung = 'IF(%s,DUNG_HY_KY!$D$109,DUNG_HY_KY!$C$109)' % NHT
    kd = 'INDEX(CSDL!$H$5:$H$14,TRUNG_GIAN!$E$5+1)'
    # điểm 1
    S = 'SU_NGHIEP'
    p.set_text(S, 'B28', 'Cách kình dương (P1 tr.117): dụng thần là quan sát và nhật can dương', why=why, must_be_empty=True, style_from='B27')
    f = ('IF(AND(%s>=0,MOD(%s-TRUNG_GIAN!$E$7,5)=3,INDEX(CSDL!$D$5:$D$14,TRUNG_GIAN!$E$5+1)=0),'
         'IF(COUNTIF(TRUNG_GIAN!$C$6:$F$6,%s)>0,"Dụng thần là quan sát và tứ trụ có kình dương hỗ trợ: có uy, vinh hiển; ấn tinh có lực là cứu ứng của quan sát (P1 tr.117).",'
         '"Dụng thần là quan sát mà tứ trụ không có kình dương: khó vinh hiển, uy quyền hạn chế; ấn tinh có lực là cứu ứng (P1 tr.117)."),'
         '"(không xét: dụng thần không phải quan sát, hoặc nhật can âm)")' % (dung, dung, kd))
    p.new_formula(S, 'C28', f, text=True, style_from='C27', why=why)
    # điểm 3
    S = 'MENH_CUC'
    p.set_text(S, 'B129', 'Yểu 10 – thân vượng mà tứ trụ thiếu cả quan sát và thực thương (P1 tr.108: nhất định khó nuôi hoặc chết yểu)', why=why, must_be_empty=True, style_from='B111')
    p.new_formula(S, 'C129', 'IF(AND(DUNG_HY_KY!$C$155=1,COUNTIFS(SAO!$G$5:$G$20,1,SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)+COUNTIFS(SAO!$G$5:$G$20,3,SAO!$X$5:$X$20,1,SAO!$Y$5:$Y$20,0)=0),1,0)',
                  style_from='C111', why=why)
    p.set_formula(S, 'C113', 'SUM(C103:C111)+C129', why=why)
    f = p.formula(S, 'C127'); old = '&IF(C111=1,", "&"bảng thực thương – ấn nghiêng yểu",""),3,999)'
    assert old in f
    p.set_formula(S, 'C127', f.replace(old, '&IF(C111=1,", "&"bảng thực thương – ấn nghiêng yểu","")&IF(C129=1,", "&"thân vượng thiếu cả quan sát và thực thương (P1 tr.108)",""),3,999)'), text=True, why=why)
