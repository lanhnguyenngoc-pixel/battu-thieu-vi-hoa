"""Mẫu thử DV-1 (duyệt 15/11/2026): xu thế đại vận theo P1 tr.124–125, có lý giải từng bước.
Quy tắc P2 tr.281–282 (dụng/hỷ/kỵ/nhàn, đ.1–đ.4) chỉ ghi chú "tham khảo thêm", không đổi kết luận."""
import sys
sys.path.insert(0, '/home/claude/battu/tools')
from tt_co_ban import thap_than, TANG, NHOM, HANH_CAN, CAN
from vuong_suy_vs2 import trang_thai_cuoi

HT = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']
TEN = {'TK': 'tỉ kiếp', 'TT': 'thực thương', 'TA': 'tài', 'QS': 'quan sát', 'AN': 'ấn'}
CAU_VUONG = '"Thân vượng… quan vận, sát vận, tài vận, thương quan vận, thực thần vận đều được xem là vận tốt; ấn, tỉ kiếp là bại vận" (P1 tr.125)'
CAU_NHUOC = '"Thân nhược… chính ấn, thiên ấn sinh thân, ngang vai, kiếp tài giúp trợ thân đều là vận tốt; quan sát, thực thương, tài là bại vận" (P1 tr.125)'


def _nua(nc, nhuoc, chu, la_can, dung):
    tt = thap_than(nc, chu if la_can else TANG[chu][0]); g = NHOM[tt]
    h = HANH_CAN[CAN.index(chu if la_can else TANG[chu][0])]
    thuan = (g in ('TK', 'AN')) == nhuoc
    muc = 'Tốt' if thuan else 'Xấu'
    them = ''
    if dung is not None and thuan and h == dung:
        muc = 'Rất tốt'; them = ' Hành %s là dụng thần → "dụng thần lên ngôi" (P1 tr.124).' % HT[h]
    elif dung is not None and not thuan and (h + 2) % 5 == dung:
        muc = 'Rất xấu'; them = ' Hành %s khắc dụng thần %s → "vận khắc dụng thần là hung vận" (P1 tr.125).' % (HT[h], HT[dung])
    ly = '%s %s là %s (nhóm %s) → %s với thân %s.%s' % ('Can' if la_can else 'Chi', chu, tt, TEN[g],
                                                     'thuận' if thuan else 'nghịch', 'nhược' if nhuoc else 'vượng', them)
    return muc, ly


def luan_la(E, nhat_can, tu_tru=None):
    tt, vs = trang_thai_cuoi(E)
    nhuoc = 'NHƯỢC' in tt.upper()
    kn = E.gia_tri('DOI_CHIEU', 'C33'); off = 15 if kn == 2 else 0
    dung_txt = str(E.gia_tri('DUNG_HY_KY', 'D106' if kn == 2 else 'C106'))
    w = dung_txt.split()
    dung = HT.index(w[0]) if w and w[0] in HT and vs['ket'] == vs['so_bo'] else None
    uv = []
    if tu_tru:
        from dung_than_nhieu_ung_vien import xet as _xet_dt
        _dt = _xet_dt(E, tu_tru); uv = _dt['chi_tiet']
    out = []
    for col in 'CDEFGHIJ':
        g = lambda r: str(E.gia_tri('DAI_VAN', col + str(r)) or '')
        can, chi = g(35).split()
        tuoi = int(float(g(36) or 0))
        m1, l1 = _nua(nhat_can, nhuoc, can, True, dung)
        m2, l2 = _nua(nhat_can, nhuoc, chi, False, dung)
        # ghi chú tham khảo theo P2 tr.281–282 (không đổi kết luận)
        mat_ky, mat_dung, bi_khac = g(51 + off), g(52 + off), g(53 + off)
        gc = []
        if bi_khac and 'tốt' in m1.lower():
            gc.append('Tham khảo thêm: có thể TỐT BỊ SUY GIẢM – lý do: can vận %s bị chữ trong mệnh cục khắc/hợp (%s), "tốt nhưng không thực" (P2 tr.282 đ.3).' % (can, bi_khac.strip('; ')))
        if bi_khac and 'xấu' in m1.lower():
            gc.append('Tham khảo thêm: XẤU GIẢM – lý do: can vận %s bị chữ trong mệnh cục chế (%s), "xấu vừa" (P2 tr.282 đ.4).' % (can, bi_khac.strip('; ')))
        if mat_dung:
            gc.append('Tham khảo thêm: 5 năm sau có thể TỐT BỊ SUY GIẢM / XẤU THÊM – lý do: chi vận %s xung mất dụng/hỷ thần ở trụ %s (P2 tr.282 đ.2).' % (chi, mat_dung.strip('; ')))
        if mat_ky:
            gc.append('Tham khảo thêm: 5 năm sau XẤU GIẢM / TỐT THÊM – lý do: chi vận %s xung mất kỵ thần ở trụ %s (P2 tr.282 đ.1).' % (chi, mat_ky.strip('; ')))
        for hanh, nhom, ly in uv:
            for chu, nua in ((can, '5 năm đầu'), (TANG[chi][0], '5 năm sau')):
                if HT[HANH_CAN[CAN.index(chu)]] == hanh:
                    gc.append('Tham khảo thêm: %s có thể TỐT nếu dụng thần là %s (%s) – lý do: %s.' % (nua, hanh, nhom, ly))
        ca = m1 if ('tốt' in m1.lower()) == ('tốt' in m2.lower()) else '%s (đầu) – %s (sau)' % (m1, m2)
        out.append(dict(van='%s %s' % (can, chi), tuoi=tuoi, nua_dau=m1, nua_sau=m2, ca_van=ca, chinh=('T' if 'tốt' in m1.lower() else 'X'),
                        buoc=[('1. Thân', 'Thân %s (%s).' % ('nhược' if nhuoc else 'vượng', tt), CAU_NHUOC if nhuoc else CAU_VUONG),
                              ('2. 5 năm đầu (can chủ 70%)', l1 + ' → ' + m1.upper(), 'P1 tr.122; P2 tr.281'),
                              ('3. 5 năm sau (chi chủ 70%)', l2 + ' → ' + m2.upper(), 'P1 tr.122; P2 tr.281')],
                        ghi_chu=gc))
    return dict(than=tt, dung=HT[dung] if dung is not None else '(cần tính lại theo VS-2)', van=out)
