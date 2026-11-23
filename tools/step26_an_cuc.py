# -*- coding: utf-8 -*-
"""Bản 6.10 (B2, người dùng yêu cầu làm 23/11/2026): dấu hiệu "Ấn thành cục, gốc nhật can bị cuốn vào cục Ấn → thân hư phù" (sách "Một trăm cách đoán mệnh" VD2 tr.24–25:
"mẹ nhiều diệt con"). CHỈ GHI dấu hiệu ở PP1 dòng 130–138 và dòng KẾT LUẬN C42; KHÔNG đổi vượng/nhược, KHÔNG đổi PP2.
Mô tả: chi tháng mang hành Ấn của nhật chủ, can Ấn lộ ở tứ trụ (dẫn), và MỌI gốc của nhật can (chi tàng can cùng hành nhật chủ) đều nằm ở chi cùng phương hoặc cùng
tam hợp với chi tháng (nên bị cuốn vào cục Ấn). SUY LUẬN từ một lá (chưa có trong Thiệu Vĩ Hoa); đo trên danh mục lá chuẩn rồi mới cho đổi kết luận."""


def an_cuc(p, why):
    S = 'PP1_DINH_TINH'
    lab = {130: 'DẤU HIỆU ẤN THÀNH CỤC LÀM MẤT GỐC NHẬT CAN (SUY LUẬN từ sách một trăm cách VD2 – chỉ ghi, không đổi kết luận) – bản 6.10',
           131: 'Chi có tàng can cùng hành nhật chủ (gốc) (1/0)  [C năm, D tháng, E ngày, F giờ]',
           132: 'Chi cùng phương hoặc cùng tam hợp với chi tháng (1/0)',
           133: 'Gốc nhật can bị cuốn vào cục của chi tháng (1/0)',
           134: 'Gốc nhật can nằm ngoài cục của chi tháng (1/0)',
           135: 'Có can Ấn lộ ở tứ trụ (dẫn) và chi tháng là hành Ấn (1/0)',
           136: 'DẤU HIỆU: Ấn thành cục, mọi gốc nhật can bị cuốn vào cục (1/0)',
           137: 'Nhật chủ đang ở trạng thái (PP1) khi có dấu hiệu',
           138: 'Cảnh báo'}
    for r, t in lab.items(): p.set_text(S, 'B%d' % r, t, style_from='B72', why=why, must_be_empty=True)
    GIO = 'NHAP_LIEU!$C$9<>"Không rõ giờ"'
    for c in 'CDEF':
        T = lambda r: 'TRUNG_GIAN!%s$%d' % (c, r)
        g = 'IF(OR(%s=TRUNG_GIAN!$E$7,%s=TRUNG_GIAN!$E$7,%s=TRUNG_GIAN!$E$7),1,0)' % (T(21), T(22), T(23))
        if c == 'F': g = 'IF(%s,%s,0)' % (GIO, g)
        p.new_formula(S, '%s131' % c, g, style_from='C72', why=why)
        same = ('IF(OR(INT(MOD(%s+1,12)/3)=INT(MOD(TRUNG_GIAN!$D$6+1,12)/3),MOD(%s,4)=MOD(TRUNG_GIAN!$D$6,4)),1,0)' % (T(6), T(6)))
        p.new_formula(S, '%s132' % c, '1' if c == 'D' else same, style_from='C72', why=why)
        p.new_formula(S, '%s133' % c, '%s131*%s132' % (c, c), style_from='C72', why=why)
        p.new_formula(S, '%s134' % c, '%s131*(1-%s132)' % (c, c), style_from='C72', why=why)
    AN = 'MOD(TRUNG_GIAN!$E$7+4,5)'
    p.new_formula(S, 'C135', 'IF(AND(TRUNG_GIAN!$D$9=%s,COUNTIF(TRUNG_GIAN!$C$7:$F$7,%s)>0),1,0)' % (AN, AN), style_from='C72', why=why)
    p.new_formula(S, 'C136', 'IF(AND(C135=1,SUM(C133:F133)>0,SUM(C134:F134)=0),1,0)', style_from='C72', why=why)
    p.new_formula(S, 'C137', 'C91', text=True, style_from='C72', why=why)
    p.new_formula(S, 'C138', 'IF(C136=1,"⚠ ẤN THÀNH CỤC: chi tháng "&TRUNG_GIAN!$D$11&" mang hành Ấn ("&INDEX(HANH_TEN,%s+1)&") và có can Ấn lộ; mọi chữ làm gốc cho nhật can "&TRUNG_GIAN!$E$10&" đều nằm ở chi cùng phương hoặc cùng cục với chi tháng nên có thể bị hóa mất, thân hư phù – sách một trăm cách (VD2, tr.24–25) luận kiểu này là thân yếu (""mẹ nhiều diệt con""). Đây là DẤU HIỆU tham khảo (suy luận từ một lá, chưa thuộc bộ sách chính), công cụ vẫn giữ kết luận: "&C91&".","")' % AN, text=True, style_from='C72', why=why)
    f = p.formula('KET_LUAN', 'C42')
    p.set_formula('KET_LUAN', 'C42', f + '&IF(PP1_DINH_TINH!$C$138<>"",CHAR(10)&PP1_DINH_TINH!$C$138,"")', text=True, why=why)
