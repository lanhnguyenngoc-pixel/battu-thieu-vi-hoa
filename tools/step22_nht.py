# -*- coding: utf-8 -*-
"""Bản 6.9, mục I bảng Ngũ hành thiếu (người dùng báo 22/11/2026 từ lá 15/8/2012 0h45 nữ: Nhâm Thìn, Mậu Thân, Mậu Thân, Nhâm Tí).
1) hien_thi_dac_dia: cột "Trạng thái" C9:C11 hiện trạng thái theo vòng trường sinh Thổ theo Thủy – cùng vòng với cách tính ở cột D và dòng 1
   (trước đó hiện vòng Thổ theo Hỏa của TRUNG_GIAN dòng 29 nên "Bệnh – Có" mâu thuẫn). Không đổi điểm.
2) cuc_chi_trung: chi trùng chữ với chi trong cục 3 chi (tam hợp, tam hội) cũng thuộc cục – hai chữ Thân của Thìn–Thân–Thân–Tí cùng hóa Thủy.
   Ảnh hưởng mọi nơi đọc TRUNG_GIAN dòng 52 (PP1, mục I…). Chỉ bật khi đo xong và người dùng duyệt."""


def _idx(c):
    ch = 'CHOOSE(TRUNG_GIAN!$E$5+1,11,6,2,9,8,3,5,0,8,3)'
    return 'IF(MOD(TRUNG_GIAN!$E$5,2)=0,MOD(TRUNG_GIAN!%s$6-%s,12),MOD(%s-TRUNG_GIAN!%s$6,12))' % (c, ch, ch, c)


def hien_thi_dac_dia(p, why):
    S = 'NGU_HANH_THIEU'
    for ref, c in (('C9', 'C'), ('C10', 'E'), ('C11', 'F')):
        f = p.formula(S, ref); old = 'TRUNG_GIAN!$%s$29' % c
        assert old in f, (ref, f)
        p.set_formula(S, ref, f.replace(old, 'INDEX(STAGE_TEN,%s+1)' % _idx(c)), text=True, why=why)


def ghi_chu_vong_thuy(p, why):
    """F8: nói rõ dòng đắc địa mục I dùng vòng Thủy cho Mậu/Kỷ (G140) – khác sách P1 tr.34 (Bính–Mậu, Đinh–Kỷ chung cột, vòng Hỏa)."""
    S = 'NGU_HANH_THIEU'
    t = p.text_of(S, 'F8')
    p.set_text(S, 'F8', t + ' LƯU Ý (6.9, 22/11/2026): riêng dòng đắc địa của mục I này tính Mậu như Nhâm, Kỷ như Quý (vòng trường sinh Thổ theo Thủy, theo yêu cầu G140) – '
               'KHÁC bảng sách P1 tr.34 (Mậu theo Bính, Kỷ theo Đinh – vòng Hỏa) mà PP1 dùng. Cột Trạng thái C9:C11 hiện đúng theo vòng Thủy này.', why=why)
