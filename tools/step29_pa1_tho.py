# -*- coding: utf-8 -*-
"""Bản 6.13 – khôi phục PA1 mục I dòng 3 "không xét nhật chủ Mậu, Kỷ" (người dùng yêu cầu 25/11/2026; kế hoạch
KE_HOACH_KHOI_PHUC_PA1_KHONG_XET_MAU_KY_25_11.md). Đảo đúng phép sửa của bản 6.8 (tools/step20_bo_sung.py): chỉ PA1; PA2 giữ nguyên.
Đo 34 lá trang: PA1 mục I khớp điểm 31 → 34, khớp nhóm 33 → 34."""
PA2 = 'LEFT(NGU_HANH_THIEU!$E$3,3)="PA2"'


def pa1_tho(p, why):
    S = 'NGU_HANH_THIEU'
    pre = 'IF(AND(TRUNG_GIAN!$E$7=2,NOT(' + PA2 + ')),"Nhật chủ Thổ – không xét (PA1, theo trang)",'
    c13 = p.formula(S, 'C13'); assert not c13.startswith(pre) and c13.endswith(')')
    p.set_formula(S, 'C13', pre + c13 + ')', why=why)
    pre2 = 'IF(OR(AND(TRUNG_GIAN!$E$7=2,NOT(' + PA2 + ')),'
    for ref, tail, new in (('D13', '="","0","Có")', '=""),"0","Có")'), ('E13', '="",0,0.2)', '=""),0,0.2)')):
        f = p.formula(S, ref); assert f.startswith('IF(IF(') and f.endswith(tail), ref
        p.set_formula(S, ref, pre2 + f[len('IF('):-len(tail)] + new, why=why)
    p.set_text(S, 'F13', 'Cách tính: có ít nhất một CAN NĂM, CAN THÁNG hoặc CAN GIỜ CÙNG HÀNH với can ngày → 20% (tính một lần; trang tính ghi dòng này là “Xét được sinh”). '
               'PA2 tính cả nhật chủ Mậu, Kỷ; PA1 (theo trang) KHÔNG xét khi nhật chủ là Mậu, Kỷ (khôi phục ở bản 6.13, 25/11/2026: trên 34 lá trang, các lá nhật chủ Thổ có can Thổ phụ đều ghi 0 cho dòng này). '
               'PA2: can hợp hóa tính theo hành đã hóa; can hợp mà không hóa vẫn giữ ngũ hành gốc nên vẫn tính (P2 tr.339, 351, 355 – thiên can theo P2; câu “ngũ hành đứng độc lập” của P1 tr.62 dùng cho địa chi). PA1 không xét hợp.', why=why)
    a2 = p.text_of(S, 'A2'); old = 'mục I từ bản 6.8 tính cả nhật chủ Mậu, Kỷ nên ở các lá nhật chủ Thổ không còn tái tạo đúng trang'; assert old in a2
    p.set_text(S, 'A2', a2.replace(old, 'mục I khớp điểm 34/34 lá trang có số liệu khi nhật chủ Mậu, Kỷ không xét ở dòng “cùng hành”, bản 6.13'), why=why)
    M = 'MAU_THUAN'; d = p.text_of(M, 'D71')
    o1 = 'mục I khớp 48/49 lá, mục II 47/48'; o2 = 'cùng hành tính cả nhật chủ Mậu/Kỷ;'
    assert o1 in d and o2 in d
    d = d.replace(o1, 'mục I khớp 34/34 lá trang có số liệu từ bản 6.13 (PA1 không xét nhật chủ Mậu, Kỷ), mục II 47/48').replace(
        o2, 'cùng hành: PA2 tính cả nhật chủ Mậu/Kỷ, PA1 không xét (bản 6.13);')
    p.set_text(M, 'D71', d, why=why)
