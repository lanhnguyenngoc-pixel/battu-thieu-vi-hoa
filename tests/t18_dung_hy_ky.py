# -*- coding: utf-8 -*-
"""T18 – Dụng–hỷ–kỵ theo P1 tr.73–124 (bản 6.12, 25/11/2026).
  • Thổ sinh tháng Hợi/Tý/Sửu/Dần: C215 = Hỏa khi thân nhược theo ít nhất một phương pháp, ngược lại −1; C217/D217 chỉ lấy Hỏa khi chính phương pháp đó nhược.
  • Các nhật chủ khác: C215 không đổi so với quy tắc cũ (kim, thủy, mộc cần hỏa tháng đông; mộc, hỏa, thổ, kim cần thủy tháng hạ).
  • Không còn ô thông quan Thực thương – Quan sát (P1 tr.110: cặp này không thành mâu thuẫn); DAI_VAN dòng 114/128/129 đúng định nghĩa H2, H8.
  • Bảng hiển thị đủ 5 hành × 2 thế, 5 × 4 mùa, nhóm dụng thần → việc.
Dùng: python -m tests.t18_dung_hy_ky"""
import sys, os, json, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.tong_hop import dung_hy_ky_hien as H
from tests.so_dap_an import oracle_nhap
GOC = os.path.dirname(os.path.abspath(__file__))


def chay(gioi_han=250):
    t0 = time.time(); loi = []; dem = {'tho_dong': 0, 'tho_dong_nhuoc': 0, 'khac': 0}
    for h in ('Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy'):
        if len(H.THE[h]) != 2: loi.append(('the', h))
        for m in ('xuân', 'hạ', 'thu', 'đông'):
            if not H.MUA[h].get(m): loi.append(('mua', h, m))
    for n in ('Quan sát', 'Chính ấn', 'Tỉ kiếp', 'Thực thương', 'Tài'):
        if not H.hy_ky_viec(n): loi.append(('viec', n))
    cs = json.load(open(os.path.join(GOC, 'cases501_a.json'), encoding='utf-8'))[:gioi_han]
    for c, nhap in oracle_nhap(cs):
        E = Engine({'NHAP_LIEU': nhap}, None); E.tinh_tat_ca()
        g = lambda s, r: E.gia_tri(s, r)
        hanh, chi = g('TRUNG_GIAN', 'E7'), g('TRUNG_GIAN', 'D6')
        c215 = g('DUNG_HY_KY', 'C215'); c7, d7 = g('DUNG_HY_KY', 'C7'), g('DUNG_HY_KY', 'D7')
        if chi in (11, 0, 1, 2):
            if hanh == 2:
                dem['tho_dong'] += 1
                mong = 1 if (c7 == 0 or d7 == 0) else -1
                dem['tho_dong_nhuoc'] += mong == 1
                if c215 != mong: loi.append((c['id'], 'C215', c215, mong))
                lo = any(g('TRUNG_GIAN', x) == 1 for x in ('C7', 'D7', 'F7'))
                if lo and c7 == 0 and g('DUNG_HY_KY', 'C217') != g('DUNG_HY_KY', 'C214') and g('DUNG_HY_KY', 'C215') == 1: loi.append((c['id'], 'C217 đổi dù Hỏa đã lộ ở can'))
                for cot, v7 in (('C', c7), ('D', d7)):
                    if v7 != 0 and g('DUNG_HY_KY', cot + '217') != g('DUNG_HY_KY', cot + '214'): loi.append((c['id'], cot + '217 đổi dù thân không nhược'))
            else:
                dem['khac'] += 1
                if (c215 == 1) != (hanh in (0, 3, 4)): loi.append((c['id'], 'C215 nhật chủ khác', c215, hanh))
        # H2: hành thứ nhất khuyết chỉ có khi ứng viên 1 không có chữ và dụng thần được chọn khác nó
        c114 = g('DAI_VAN', 'C114')
        if c114 != -1:
            dem['thu_nhat_khuyet'] = dem.get('thu_nhat_khuyet', 0) + 1
            if g('DUNG_HY_KY', 'C103') == g('DUNG_HY_KY', 'C93') or g('DUNG_HY_KY', 'C103') < 0: loi.append((c['id'], 'C114 sai điều kiện'))
        # H8: hành kỵ thật phải là hành kỵ thần
        c129 = g('DAI_VAN', 'C129')
        if c129 != -1:
            dem['ky_that'] = dem.get('ky_that', 0) + 1
            k = g('DUNG_HY_KY', 'D111' if int(g('DOI_CHIEU', 'C33') or 0) == 2 else 'C111')
            if c129 != k: loi.append((c['id'], 'C129 khác hành kỵ', c129, k))
    return loi, dem, time.time() - t0


if __name__ == '__main__':
    loi, dem, dt = chay()
    print('T18 dụng–hỷ–kỵ – %.1f giây: %s' % (dt, dem))
    for x in loi[:30]: print('  LỆCH', x)
    print('  ĐẠT' if not loi else '  CÓ %d LỆCH' % len(loi)); sys.exit(1 if loi else 0)
