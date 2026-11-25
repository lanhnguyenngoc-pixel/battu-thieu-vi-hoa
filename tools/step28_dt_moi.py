# -*- coding: utf-8 -*-
"""Bản 6.12 – dụng–hỷ–kỵ theo P1 tr.73–124 (người dùng duyệt 25/11/2026).
  H1: Thổ sinh tháng Hợi/Tý/Sửu/Dần, thân nhược → cần Hỏa điều hầu (P1 tr.78); dưới 20% thì Hỏa làm dụng thần (quy ước công cụ).
  H2: xếp bậc vận khi dụng thần thứ nhất khuyết (P1 tr.102–103): vận của hành thứ nhất ±2, hành dụng thần được chọn ±1 (DAI_VAN dòng 114, 128).
  H8: kỵ thần thật (hành kỵ có can lộ trong tứ trụ – P1 tr.98): vận hành kỵ nặng thêm một mức (DAI_VAN dòng 129, 136).
  Không làm H3 (cặp Thực thương – Quan sát): P1 tr.110 nêu rõ cặp này là nhược đối với nhược, không thành mâu thuẫn nên không thông quan."""


def dt_moi(p, why):
    S = 'DUNG_HY_KY'
    E7 = 'TRUNG_GIAN!$E$7'
    # H1 – C215: thêm Thổ (2) cho nhóm tháng lạnh khi thân nhược theo ít nhất một phương pháp
    p.replace_in_formula(S, 'C215', 'OR(%s=0,%s=3,%s=4)),1,' % (E7, E7, E7),
                         'OR(%s=0,%s=3,%s=4,AND(%s=2,OR($C$7=0,$D$7=0)))),1,' % (E7, E7, E7, E7), count=1, why=why)
    # C217/D217: Thổ chỉ lấy Hỏa làm dụng thần khi chính phương pháp đó xác định thân nhược
    for c in 'CD':
        p.replace_in_formula(S, c + '217', 'AND(%s$91<>"TH",$C$215>=0,' % c,
                             'AND(%s$91<>"TH",$C$215>=0,IF(%s=2,%s$7=0,TRUE),' % (c, E7, c), count=1, why=why)
    # Thổ đông: "Hỏa thiếu" = không có Hỏa ở CAN LỘ (năm, tháng, giờ); tàng can không tính (Cùng Thông Bảo Giám); nhóm khác giữ ngưỡng 20%
    loan = ('IF(AND(%s=2,$C$215=1),NOT(OR(TRUNG_GIAN!$C$7=$C$215,TRUNG_GIAN!$D$7=$C$215,AND(NHAP_LIEU!$C$9<>"Không rõ giờ",TRUNG_GIAN!$F$7=$C$215))),N($C$216)<0.2)') % E7
    for c in 'CD':
        p.replace_in_formula(S, c + '217', 'N($C$216)<0.2', loan, count=1, why=why)
    p.replace_in_formula(S, 'C218', 'IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),"mùa nóng","mùa lạnh")',
                         'IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),"mùa nóng","mùa lạnh")', why=why)
    for c in 'CD':
        old = 'P1 tr.76–79: sinh mùa lạnh cần hỏa sưởi ấm'
        p.replace_in_formula(S, c + '219', old, 'P1 tr.76–79 (thổ mùa đông: tr.78): sinh mùa lạnh cần hỏa sưởi ấm', count=1, why=why)
    p.set_text(S, 'B215', 'Hành điều hầu cần (P1 tr.76–79: kim, thủy, mộc, và thổ khi thân nhược, sinh tháng Hợi–Tí–Sửu–Dần cần hỏa; mộc, hỏa, thổ, kim sinh tháng Tị–Ngọ–Mùi cần thủy; −1 = không)', why=why)
    # ---- H2 + H8 (DAI_VAN)
    V = 'DAI_VAN'; D = 'DUNG_HY_KY'
    def ung1(c):
        t = ('IF(AND(D!$X$72="",D!$X$91<>"TH",D!$X$93>=0,D!$X$103<>D!$X$93,INDEX(D!$X$209:$X$213,D!$X$93+1)=0),MOD(D!$X$5+D!$X$93,5),-1)')
        return t.replace('D!', D + '!').replace('$X$', '$' + c + '$').replace('$X$213', '$' + c + '$213')
    p.new_formula(V, 'C114', ung1('C'), why=why, style_from='C111')
    p.set_text(V, 'B114', 'PP1 – hành của dụng thần thứ nhất khi khuyết, được xếp bậc ±2 (−1 = không; P1 tr.102–103, bản 6.12)', style_from='B111', why=why)
    p.new_formula(V, 'C128', ung1('D'), why=why, style_from='C111')
    p.set_text(V, 'B128', 'PP2 – hành của dụng thần thứ nhất khi khuyết, được xếp bậc ±2 (−1 = không; P1 tr.102–103, bản 6.12)', style_from='B111', why=why)
    for col in 'CDEFGHIJKL':
        for r, c43, pp, cell in ((115, '42', 'C', '$C$114'), (116, '43', 'C', '$C$114'), (117, '42', 'D', '$C$128'), (118, '43', 'D', '$C$128')):
            a = '%s$%s=DUNG_HY_KY!$%s$109' % (col, c43, pp)
            b = 'MOD(%s$%s+2,5)=DUNG_HY_KY!$%s$109' % (col, c43, pp)
            rep = lambda x: x.replace('DUNG_HY_KY!$%s$109' % pp, 'IF(%s>=0,%s,DUNG_HY_KY!$%s$109)' % (cell, cell, pp))
            p.replace_in_formula(V, '%s%d' % (col, r), a, rep(a), count=1, why=why)
            p.replace_in_formula(V, '%s%d' % (col, r), b, rep(b), count=1, why=why)
    k = 'IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$111,DUNG_HY_KY!$C$111)'
    f = ('IF(AND(%s>=0,OR(TRUNG_GIAN!$C$7=%s,TRUNG_GIAN!$D$7=%s,AND(NHAP_LIEU!$C$9<>"Không rõ giờ",TRUNG_GIAN!$F$7=%s))),%s,-1)') % (k, k, k, k, k)
    p.new_formula(V, 'C129', f, why=why, style_from='C111')
    p.set_text(V, 'B129', 'Hành kỵ thần thật: hành kỵ thần có can lộ trong tứ trụ (−1 = không); vận hành này nặng thêm một mức (P1 tr.98, bản 6.12)', style_from='B111', why=why)
    for col in 'CDEFGHIJKL':
        p.replace_in_formula(V, col + '136', 'IF(%s135=1,0,IF(OR(' % col,
                             'IF(%s135=1,0,IF(OR(AND($C$129>=0,%s$42=$C$129,%s119<0),' % (col, col, col), count=1, why=why)
