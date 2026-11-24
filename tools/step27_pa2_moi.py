# -*- coding: utf-8 -*-
"""Bản 6.11 – bảng Ngũ hành thiếu, mục I, chế độ PA2 "cách A, Thổ theo Thủy, có xét giả ở E7 và E12" (người dùng chốt 24/11/2026 08:35; kế hoạch
KE_HOACH_PA2_MUC_I_NGU_HANH_THIEU.md). Chỉ đổi mục I của PA2; PA1 giữ nguyên; mục II, III, PP1, dụng thần, đại vận KHÔNG đổi.
  E7  = 48% khi CAN THÁNG ở 1 trong 5 trạng thái đầu tại CHI THÁNG và KHÔNG "giả" (PA2); PA1 không xét giả (như 6.10).
  E12 = 20% khi một trong CAN NĂM, CAN NGÀY, CAN GIỜ (giờ chỉ khi biết) tại CHI THÁNG ở Trường sinh, Lâm quan (Lộc), Đế vượng (Kình dương) hoặc Mộ (khố)
        và không "giả"; tính một lần. Không thêm điều kiện "can dương".
  E13, E14 = như PA1: hành gốc của can, không hợp hóa (PA2 không còn dùng hành sau hợp hóa của PP2 ở hai dòng này); nhật chủ Mậu/Kỷ vẫn tính.
"Giả" (SUY LUẬN, chờ người dùng xác nhận): hành chi tháng không cùng hành và không sinh hành của can đó."""

CH = 'CHOOSE(TRUNG_GIAN!$%s$5+1,11,6,2,9,8,3,5,0,8,3)'


def _st(c):   # trạng thái vòng trường sinh (Thổ theo Thủy) của can ở cột c của TRUNG_GIAN tại chi tháng
    ch = CH % c
    return 'IF(MOD(TRUNG_GIAN!$%s$5,2)=0,MOD(TRUNG_GIAN!$D$6-%s,12),MOD(%s-TRUNG_GIAN!$D$6,12))' % (c, ch, ch)


def _that(c):  # KHÔNG giả: chi tháng cùng hành hoặc sinh hành can
    return 'OR(TRUNG_GIAN!$D$9=TRUNG_GIAN!$%s$7,MOD(TRUNG_GIAN!$D$9+1,5)=TRUNG_GIAN!$%s$7)' % (c, c)


def pa2_moi(p, why):
    S = 'NGU_HANH_THIEU'
    PA2 = 'LEFT(NGU_HANH_THIEU!$E$3,3)="PA2"'
    # --- E7: được lệnh có xét giả (PA2)
    s = _st('D'); ok = 'OR(NOT(%s),%s)' % (PA2, _that('D'))
    p.set_formula(S, 'E7', 'IF(AND(%s<=4,%s),0.48,0)' % (s, ok), why=why)
    p.set_formula(S, 'D7', 'IF(%s<=4,IF(%s,"Có","Giả – không tính"),"Không")' % (s, ok), text=True, why=why)
    f = p.formula(S, 'C7')
    p.set_formula(S, 'C7', f + '&IF(AND(%s<=4,NOT(%s)),"  (giả)","")' % (s, ok), text=True, why=why)
    t = p.text_of(S, 'F7')
    p.set_text(S, 'F7', 'Cách tính: tra vòng trường sinh của CAN THÁNG tại CHI THÁNG (tháng theo tiết khí; Thổ theo Thủy: Mậu như Nhâm, Kỷ như Quý). Thuộc 5 trạng thái đầu '
               '(Trường sinh, Mộc dục, Quan đới, Lâm quan, Đế vượng) → được lệnh 48%; 7 trạng thái còn lại → 0. PA2 (bản 6.11, 24/11/2026) CÓ xét “giả”: hành chi tháng '
               'không cùng hành và không sinh hành can tháng → không tính. PA1 không xét giả (như 6.8). Không xét hợp.', why=why)
    # --- E12 (D9:D11): đắc địa của can năm, can ngày, can giờ tại chi tháng
    nhan = {9: ('C', 'năm'), 10: ('E', 'ngày'), 11: ('F', 'giờ')}
    for r, (c, ten) in nhan.items():
        st = _st(c)
        tt = 'OR(%s=0,%s=3,%s=4,%s=8)' % (st, st, st, st)
        gio = (c == 'F')
        cond = 'AND(%s,%s)' % (tt, _that(c))
        d = 'IF(NOT(%s),"PA1: 0",IF(%s,"Có","0"))' % (PA2, cond)
        if gio: d = 'IF(NHAP_LIEU!$C$9="Không rõ giờ","0",%s)' % d
        p.set_formula(S, 'D%d' % r, d, text=True, why=why)
        b = '"   Can %s: "&INDEX(CAN_TEN,TRUNG_GIAN!$%s$5+1)&" tại chi tháng "&TRUNG_GIAN!$D$11' % (ten, c)
        if gio: b = 'IF(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),%s,"   Can giờ: (không rõ giờ)")' % b
        p.set_formula(S, 'B%d' % r, b, text=True, why=why)
        cc = 'INDEX(STAGE_TEN,%s+1)' % st
        if gio: cc = 'IF(NOT(NHAP_LIEU!$C$9="Không rõ giờ"),%s,"")' % cc
        p.set_formula(S, 'C%d' % r, cc, text=True, why=why)
    p.set_formula(S, 'B8', '"Xét đắc địa – can năm, can ngày, can giờ tại chi tháng "&TRUNG_GIAN!$D$11&" (PA2 bản 6.11)"', text=True, why=why)
    p.set_text(S, 'F8', 'Cách tính. PA2 (bản 6.11, người dùng chốt 24/11/2026): xét CAN NĂM, CAN NGÀY, CAN GIỜ (giờ chỉ khi biết) tại CHI THÁNG theo vòng trường sinh Thổ theo Thủy '
               '(Mậu như Nhâm, Kỷ như Quý – chỉ dùng ở bảng này). Gặp Trường sinh, Mộ khố, Lộc (Lâm quan) hoặc Kình dương (Đế vượng) và KHÔNG giả (hành chi tháng cùng hành hoặc sinh hành can đó) '
               '→ có ít nhất một can đạt thì tính 20%, chỉ một lần. Không thêm điều kiện can dương. PA1 – Theo trang: luôn 0.', why=why)
    # --- E13, E14: hành gốc của can, không hợp hóa (PA2 giống PA1)
    for ref in ('C13', 'D13', 'E13', 'C14', 'D14', 'E14'):
        for c in 'CDF': p.replace_in_formula(S, ref, 'PP2_DINH_LUONG!$%s$52' % c, 'TRUNG_GIAN!$%s$7' % c, why=why)
    t = p.text_of(S, 'F13')
    p.set_text(S, 'F13', 'Cách tính: có ít nhất một CAN NĂM, CAN THÁNG hoặc CAN GIỜ CÙNG HÀNH với can ngày → 20% (tính một lần; trang tính ghi dòng này là “Xét được sinh”). '
               'Cả hai chế độ tính theo hành gốc của can, KHÔNG xét hợp hóa (bản 6.11), nhật chủ Mậu, Kỷ vẫn tính, không có ngoại lệ cho Thổ.', why=why)
    p.set_text(S, 'F14', 'Cách tính: có ít nhất một CAN NĂM, CAN THÁNG hoặc CAN GIỜ mà hành gốc của nó SINH cho hành can ngày (Mộc→Hỏa→Thổ→Kim→Thủy→Mộc) → 12% (tính một lần). '
               'Trang tính ghi dòng này là “Xét được trợ giúp”. Cả hai chế độ không xét hợp hóa (bản 6.11).', why=why)
