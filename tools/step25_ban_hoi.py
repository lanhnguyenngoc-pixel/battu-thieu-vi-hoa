# -*- coding: utf-8 -*-
"""Bản 6.10 (mục 3 kế hoạch, người dùng yêu cầu 23/11/2026): BÁN HỘI chỉ ở PP1 bước 5b (P1 tr.71 b: "bán hợp hoặc bán hội").
Bốn cặp chặt kề nhau: Dần–Mão (Mộc), Tị–Ngọ (Hỏa), Thân–Dậu (Kim), Hợi–Tí (Thủy). SUY LUẬN ngoài sách (sách chỉ nhắc chữ "bán hội", không liệt kê cặp) – cần duyệt.
KHÔNG đưa vào PP2 (sách không cho trọng số), KHÔNG đổi mục I Ngũ hành thiếu, KHÔNG đổi các nơi khác đọc bán hợp (TRUNG_GIAN dòng 55–62)."""

PAIRS = ((2, 3, 0), (5, 6, 1), (8, 9, 3), (11, 0, 4))      # (chi a, chi b, hành)


def ban_hoi_pp1(p, why):
    S = 'PP1_DINH_TINH'
    lab = {121: 'BÁN HỘI cho bước 5b (SUY LUẬN ngoài sách, 4 cặp chặt kề nhau: Dần–Mão, Tị–Ngọ, Thân–Dậu, Hợi–Tí; chỉ dùng ở dòng 72, không vào PP2) – bản 6.10',
           122: 'Cặp kề là bán hội: hành (−1 = không)  [cột C = Năm–Tháng, D = Tháng–Ngày, E = Ngày–Giờ]',
           123: 'Bán hội giữ lại (hai chi chưa thuộc cục 3 chi / bán hợp) (1/0)',
           124: 'Diễn giải bán hội',
           125: 'Các cặp bán hội dùng ở 5b'}
    for r, t in lab.items(): p.set_text(S, 'B%d' % r, t, style_from='B72', why=why, must_be_empty=True)
    nxt = {'C': 'D', 'D': 'E', 'E': 'F'}
    for c, n in nxt.items():
        a, b = 'TRUNG_GIAN!%s$6' % c, 'TRUNG_GIAN!%s$6' % n
        f = '0'
        f = '-1'
        for (x, y, h) in reversed(PAIRS):
            f = 'IF(OR(AND(%s=%d,%s=%d),AND(%s=%d,%s=%d)),%d,%s)' % (a, x, b, y, a, y, b, x, h, f)
        if c == 'E': f = 'IF(NHAP_LIEU!$C$9="Không rõ giờ",-1,%s)' % f
        col = c
        p.new_formula(S, '%s122' % col, f, style_from='C72', why=why)
        p.new_formula(S, '%s123' % col, 'IF(AND(%s122>=0,TRUNG_GIAN!%s56=1,TRUNG_GIAN!%s57=0),1,0)' % (col, col, col), style_from='C72', why=why)
        p.new_formula(S, '%s124' % col, 'IF(%s123=1,INDEX(CHI_TEN,%s+1)&"–"&INDEX(CHI_TEN,%s+1)&" ("&INDEX(HANH_TEN,%s122+1)&")","")' % (col, a, b, col), text=True, style_from='C72', why=why)
    p.new_formula(S, 'C125', 'TRIM(C124&" "&D124&" "&E124)', text=True, style_from='C72', why=why)
    f = p.formula(S, 'C72'); assert f.startswith('SIGN(') and f.endswith(')'), f[:30]
    extra = ''.join('+IF(%s123=1,IF(OR(TRUNG_GIAN!%s$122=TRUNG_GIAN!$E$7,%s122=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)' % (c, c, c) for c in 'CDE')
    extra = ''.join('+IF(%s123=1,IF(OR(%s122=TRUNG_GIAN!$E$7,%s122=MOD(TRUNG_GIAN!$E$7+4,5)),1,-1),0)' % (c, c, c) for c in 'CDE')
    p.set_formula(S, 'C72', f[:-1] + extra + ')', why=why)
    p.set_text(S, 'B72', 'b) Bán hợp (cộng bán hội – 4 cặp chặt, SUY LUẬN ngoài sách, xem dòng 121–125) sinh–trợ (+) hay ngược lại (−)', why=why)
    o = 'IF(C72>0,"bán hợp sinh/trợ nhật chủ → +1",IF(C72<0,"bán hợp khắc/tiết/hao nhật chủ → −1","không có bán hợp liên quan → 0"))'
    f = p.formula(S, 'C109'); assert o in f
    p.set_formula(S, 'C109', f.replace(o, o + '&IF(LEN(C125)>0," (có tính bán hội "&C125&" – suy luận ngoài sách, chưa được sách xác nhận)","")', 1), text=True, why=why)
