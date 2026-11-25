# -*- coding: utf-8 -*-
"""Lớp hiển thị dụng–hỷ–kỵ thần (bản 1.9.5, theo P1 tr.74–79, 96–98) – CHỈ chữ tham khảo, không đổi dụng thần.
Nội dung diễn đạt lại ngắn gọn từ sách (không chép nguyên văn). Chủ dự án duyệt 25/11/2026.
    the_ngu_hanh(hanh_nhat_chu, vuong)   → câu theo thế mạnh/yếu của nhật chủ (tr.74–75)
    theo_mua(hanh_nhat_chu, chi_thang)   → câu theo mùa sinh (tr.75–79)
    hy_ky_viec(nhom_dung)                → (khi thuận, khi bất lợi) của nhóm dụng thần (tr.96–98)"""

# Mỗi hành: [khi nhật chủ mạnh, khi nhật chủ yếu]
THE = {
    'Mộc': ['Mộc mạnh, thủy nhiều: dùng kim làm yếu mộc; kim ít thì dùng thổ.', 'Mộc yếu, kim mạnh: dùng hỏa chế kim; hỏa yếu thì dùng mộc.'],
    'Hỏa': ['Hỏa mạnh, mộc nhiều: dùng thủy; thủy yếu thì dùng kim.', 'Hỏa yếu, thủy vượng: dùng thổ chế thủy; thổ yếu thì dùng hỏa.'],
    'Thổ': ['Thổ dày, thủy ít: dùng mộc làm thưa; mộc yếu thì dùng thủy.', 'Thổ mỏng, mộc thịnh: dùng kim chế mộc; kim yếu thì dùng thổ.'],
    'Kim': ['Kim cứng, thổ dày: dùng hỏa luyện; hỏa yếu thì dùng mộc.', 'Kim mềm, mộc nặng: dùng thổ sinh kim; thổ yếu thì dùng kim.'],
    'Thủy': ['Thủy mạnh, kim nặng: dùng thổ ngăn; thổ yếu thì dùng hỏa.', 'Thủy yếu, thổ nhiều: dùng mộc chế thổ; mộc yếu thì dùng thủy.'],
}
MUA_TEN = {'Dần': 'xuân', 'Mão': 'xuân', 'Thìn': 'xuân', 'Tị': 'hạ', 'Ngọ': 'hạ', 'Mùi': 'hạ',
           'Thân': 'thu', 'Dậu': 'thu', 'Tuất': 'thu', 'Hợi': 'đông', 'Tý': 'đông', 'Tí': 'đông', 'Sửu': 'đông'}
MUA = {
    'Mộc': {'xuân': 'còn khí lạnh nên cần hỏa làm ấm; thủy vừa đủ, thổ không dày',
            'hạ': 'khô nóng, cần nước tưới; kỵ hỏa nhiều; thổ không dày, kim không mạnh',
            'thu': 'đầu thu cần thủy và thổ; giữa thu cần kim làm rụng quả; cuối thu cần hỏa; kỵ thủy thịnh',
            'đông': 'cần thổ nuôi, sợ thủy dìm, cần hỏa sưởi ấm, sợ rơi vào tử tuyệt'},
    'Hỏa': {'xuân': 'cần mộc nhưng không quá nhiều; thủy vừa phải; sợ thổ nhiều làm tắt',
            'hạ': 'cần thủy chế bớt, mộc phù trợ, kim thổ hợp; kỵ hỏa nhiều',
            'thu': 'cần mộc; sợ thủy dập, thổ nhiều che mờ, kim nhiều cướp thế',
            'đông': 'cần mộc; hỏa gặp hỏa thêm sức; sợ thủy, cần thổ chế thủy'},
    'Thổ': {'xuân': 'yếu, cần hỏa và thổ giúp; sợ mộc khắc, thủy cuốn; kim chế mộc thì có sức, kim vượng quá thì xì hơi',
            'hạ': 'khô, cần nhiều thủy; hỏa vượng càng khô; mộc hỏa không hợp',
            'thu': 'kim nhiều xì hơi; mộc thịnh khắc thổ; sợ thủy mạnh; hỏa không đáng ngại',
            'đông': 'ngoài lạnh trong ấm; thủy mạnh vẫn tốt; hỏa thịnh làm thổ thêm phồn vinh; thổ gặp thổ thêm khỏe'},
    'Kim': {'xuân': 'còn lạnh, gặp hỏa rất quý; cần thổ giúp; sợ thủy nhiều và mộc mạnh',
            'hạ': 'mềm yếu; thủy mạnh tốt nhất; hỏa nhiều xấu; thổ mỏng có lợi',
            'thu': 'khí mạnh; cần hỏa luyện, thổ nuôi, thủy làm tươi; kim quá cường thì gãy',
            'đông': 'lạnh; cần thổ chế thủy; hỏa sinh thổ tốt; sợ thủy vượng nhấn chìm'},
    'Thủy': {'xuân': 'dầm dề, cần thổ ngăn; kim vừa phải; hỏa không mạnh quá',
             'hạ': 'ngoài thực trong hư, đang khô nên cần thủy sánh vai; kim phù trợ tốt; sợ vượng hỏa',
             'thu': 'kim thủy cùng vượng; kim tương trợ thì trong; sợ thổ làm đục; hỏa mộc vừa phải',
             'đông': 'có khí thế; cần hỏa trừ hàn, thổ gom lại; mộc nhiều hữu tình; thủy mạnh cần thổ làm đê'},
}
VIEC = {
    'Quan sát': ('quý nhân nâng đỡ, thăng chức, thi đỗ, uy quyền tăng', 'hình khắc quan phù, dễ vướng thị phi pháp lý, danh dự tổn thất, bị khống chế hoặc liên lụy'),
    'Chính quan': ('quý nhân nâng đỡ, thăng chức, thi đỗ, uy quyền tăng', 'hình khắc quan phù, dễ vướng thị phi pháp lý, danh dự tổn thất, bị khống chế hoặc liên lụy'),
    'Chính ấn': ('công thành danh toại, lên chức, học thuật tiến bộ, thi cử đỗ', 'thân thể mệt mỏi, mất chức quyền, danh dự tổn thất, thi hỏng'),
    'Thiên ấn': ('công thành danh toại, lên chức, học thuật tiến bộ, thi cử đỗ', 'thân thể mệt mỏi, mất chức quyền, danh dự tổn thất, thi hỏng'),
    'Tỉ kiếp': ('được lợi, có người giúp sức, hôn nhân tốt, bệnh mau lành', 'tổn thất tiền của vợ, anh em bất hòa, bạn bè cản trở'),
    'Tỉ': ('được lợi, có người giúp sức, hôn nhân tốt, bệnh mau lành', 'tổn thất tiền của vợ, anh em bất hòa, bạn bè cản trở'),
    'Thực thương': ('tin mừng, thêm con cháu, thêm tuổi thọ, tài hoa, chức lộc đạt', 'con cháu gây liên lụy, thân nhược nhiều bệnh, giáng chức, học dở dang, thất nghiệp'),
    'Thực': ('tin mừng, thêm con cháu, thêm tuổi thọ, tài hoa, chức lộc đạt', 'con cháu gây liên lụy, thân nhược nhiều bệnh, giáng chức, học dở dang, thất nghiệp'),
    'Tài': ('hôn nhân thuận, tài lợi, hưởng âm đức, vợ con giúp, sự nghiệp mở mang', 'tài nhiều thân nhược, không được cha mẹ vợ con giúp, làm ăn quẫn bách'),
}
LUU_Y_TR98 = ('Chính quan, chính ấn, chính tài, thực thần là cát thần nhưng tổ hợp xấu hoặc gặp suy, bệnh, tử, tuyệt thì không tốt; '
              'thương quan, thiên ấn, thiên tài, thất sát, tỉ kiếp là thần phá hại nhưng tổ hợp tốt thì lại bổ cứu nên vẫn tốt (P1 tr.98).')


def the_ngu_hanh(hanh, vuong):
    c = THE.get(hanh)
    return ('Tham khảo thêm – can ngày %s (P1 tr.74–75): %s' % (hanh, c[0 if vuong else 1])) if c else ''


def theo_mua(hanh, chi_thang):
    mua = MUA_TEN.get(chi_thang)
    t = MUA.get(hanh, {}).get(mua) if mua else None
    return ('Tham khảo thêm – %s sinh mùa %s (P1 tr.75–79): %s.' % (hanh, mua, t)) if t else ''


def hy_ky_viec(nhom):
    for k, v in VIEC.items():
        if (nhom or '').startswith(k) or k in (nhom or ''):
            return v
    return None


_H = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']


def hy_cuu_ung(hanh_nhat_chu, hanh_dung, vuong, khuyet=False):
    """Hỷ thần cứu ứng (P1 tr.97–98): hai hướng – (a) hành sinh trợ dụng thần, (b) hành chế áp kỵ thần (hành dụng thần sinh ra);
    bỏ hành thuộc phe bất lợi cho thân. Khi dụng thần khuyết thì hỷ thần gánh vác phần việc của dụng thần (tr.97)."""
    if hanh_nhat_chu not in _H or hanh_dung not in _H: return ''
    n, d = _H.index(hanh_nhat_chu), _H.index(hanh_dung)
    xau = {(n + 0) % 5, (n + 4) % 5} if vuong else {(n + 1) % 5, (n + 2) % 5, (n + 3) % 5}
    a, b = (d - 1) % 5, (d + 1) % 5
    ds = []
    if a not in xau: ds.append('%s (sinh trợ dụng thần %s)' % (_H[a], hanh_dung))
    if b not in xau: ds.append('%s (chế áp kỵ thần %s)' % (_H[b], _H[(d - 2) % 5]))
    if not ds: return ''
    t = 'Hỷ thần cứu ứng (P1 tr.97–98): ' + '; '.join(ds) + '. Khi lực dụng thần không đủ, các hành này sinh trợ hoặc chế áp hung thần để cứu.'
    if khuyet: t += ' Dụng thần không có chữ trong tứ trụ nên hỷ thần gánh phần việc của dụng thần (tr.97).'
    return t
