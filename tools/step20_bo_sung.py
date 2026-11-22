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
