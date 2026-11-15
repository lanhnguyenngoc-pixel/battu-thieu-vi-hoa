"""Mẫu thử A (duyệt 15/11/2026): XÉT DỤNG THẦN NHIỀU ỨNG VIÊN, có lý giải từng bước.
Dụng thần chính giữ theo quy tắc hiện tại (P1 tr.99 + CĐB-1/3/4). Các ứng viên thay thế chỉ để tham khảo."""
import sys
sys.path.insert(0, '/home/claude/battu/tools')
from tt_co_ban import HANH_CAN, CAN, TANG, NHOM, thap_than
from gd_b_dac_trung import vai_tro
from vuong_suy_vs2 import trang_thai_cuoi
from chon_dung_than import the_ngu_hanh
from chon_dung_than2 import ung_vien, HT, G, TK, TT, TA, QS, AN, DONG, HA

MA = ['TK', 'TT', 'TA', 'QS', 'AN']


def _co_mat(tt):
    """Số chữ mỗi nhóm: can lộ (trừ nhật can) + mọi tàng can."""
    p = [x.split() for x in tt.split(',')]; nc = p[2][0]; d = [0] * 5
    for i, (c, ch) in enumerate(p):
        if i != 2: d[MA.index(NHOM[thap_than(nc, c)])] += 1
        for x in TANG[ch]: d[MA.index(NHOM[thap_than(nc, x)])] += 1
    return d


def xet(E, tt):
    nh = HANH_CAN[CAN.index(tt.split(',')[2].split()[0])]
    tt_cuoi, _ = trang_thai_cuoi(E); nhuoc = 'NHƯỢC' in tt_cuoi.upper()
    kn = E.gia_tri('DOI_CHIEU', 'C33'); col = 'D' if kn == 2 else 'C'
    chinh = str(E.gia_tri('DUNG_HY_KY', col + '106')); hy = str(E.gia_tri('DUNG_HY_KY', col + '107'))
    cach = str(E.gia_tri('DUNG_HY_KY', col + '102'))
    L = vai_tro(E)['luong2']; a = [L.get(x) or 0 for x in G]; t = sum(a) or 1; p = [x / t for x in a]
    co = _co_mat(tt); thang = tt.split(',')[1].split()[1]
    B = [('1. Dụng thần chính', '%s (hỷ %s) – %s; thân %s.' % (chinh, hy, cach, 'nhược' if nhuoc else 'vượng'),
          'P1 tr.99 (sinh phù – áp chế); P2 tr.369–381 nếu là cách đặc biệt')]
    thay = []
    if 'đặc biệt' not in cach:
        Lu, ly, dom = ung_vien(nhuoc, a)
        g1 = Lu[0]
        if co[g1] == 0:
            thay.append((Lu[1], 'Ứng viên thứ nhất %s khuyết trong tứ trụ → "nếu không có … thì lấy …" (P1 tr.99)' % G[g1]))
        elif not nhuoc and g1 in (TT, TA, QS) and p[g1] >= p[TK] + p[AN]:
            for g2 in Lu[1:]:
                thay.append((g2, '%s đã mạnh (%.0f%%) hơn phe mình (%.0f%%) – "đã chế áp nên tương đối cân bằng" (VD36 P2 tr.427)' % (G[g1], 100 * p[g1], 100 * (p[TK] + p[AN]))))
        T, lyT = the_ngu_hanh(nh, a)
        thay.append((T[0], 'Theo thế ngũ hành của can ngày: %s → nên %s (P1 tr.74–75)' % (lyT, G[T[0]])))
        dh = 1 if thang in DONG | {'Dần'} else (4 if thang in HA else None)
        if dh is not None:
            gdh = (dh - nh) % 5
            mua = 'đầu xuân còn khí lạnh (P1 tr.76, 78)' if thang == 'Dần' else ('mùa đông' if dh == 1 else 'mùa hạ')
            if p[gdh] < 0.2:
                thay.append((gdh, 'Sinh %s, %s còn ít (%.0f%%) → điều hầu (P1 tr.76–79, 100)' % (mua, HT[dh], 100 * p[gdh])))
    chinh_h = chinh.split()[0]
    uv = []
    for g, ly in thay:
        hanh = HT[(nh + g) % 5]
        if hanh == chinh_h or any(u[0] == hanh for u in uv): continue
        uv.append((hanh, G[g], ly))
    B.append(('2. Ứng viên thay thế (tham khảo)', '; '.join('%s (%s): %s' % u for u in uv) or 'Không có – các nguyên tắc cùng chỉ một dụng thần.',
              'Tác giả cân nhắc nhiều nguyên tắc; sách không nêu thứ tự ưu tiên (P1 tr.99)'))
    B.append(('3. Mức chắc chắn', 'CAO – các nguyên tắc thống nhất.' if not uv else 'CẦN CÂN NHẮC – có %d ứng viên khác; luận vận có ghi chú theo các ứng viên này.' % len(uv), ''))
    return dict(chinh=chinh_h, ung_vien=[u[0] for u in uv], chi_tiet=uv, buoc=B)
