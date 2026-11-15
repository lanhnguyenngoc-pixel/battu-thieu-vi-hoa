"""Giai đoạn C – mẫu thử GỘP mọi quy tắc đã duyệt (15/11/2026), để chạy chéo trước khi vào workbook 6.5.
VS-2 + G11 + VS-3 (vượng suy, độ tin cậy) → CĐB-1/3/4 (cách đặc biệt) → DT-2a + DT-3 (dụng thần)
→ DV-3 + TL-1 (đại vận, từng lĩnh vực). B1/B2 (bệnh–mức, gần tòng) tính kèm."""
import sys
sys.path.insert(0, '/home/claude/battu/tools')
from tt_co_ban import thap_than, TANG, NHOM, HANH_CAN, CAN, truong_sinh
from gd_b_dac_trung import vai_tro
from vuong_suy_vs2 import trang_thai_cuoi
from chon_dung_than2 import ung_vien, G, HT, TK, TT, TA, QS, AN
from dung_than_nhieu_ung_vien import _co_mat
import dai_van_dv3 as DV

GOC_MANH = {'Trường sinh', 'Lâm quan', 'Đế vượng'}


def _goc_manh(tt):
    p = [x.split() for x in tt.split(',')]; nc = p[2][0]
    for c in (p[i][1] for i in range(4)):
        s = truong_sinh(nc, c)
        if s in GOC_MANH and not (s == 'Trường sinh' and CAN.index(nc) % 2 == 1): return True
    return False


def _so_chu_phe_minh(tt):
    p = [x.split() for x in tt.split(',')]; nc = p[2][0]; nh = HANH_CAN[CAN.index(nc)]; n = 0
    for c, ch in p:
        for x in (c, TANG[ch][0]):
            h = HANH_CAN[CAN.index(x)]
            if h == nh or (h + 1) % 5 == nh: n += 1
    return n


def luan(E, tt, gioi='Nam', thu_tu_pp2=False):
    out = {}
    # 1. VƯỢNG SUY
    tt_cuoi, vs = trang_thai_cuoi(E); nhuoc = 'NHƯỢC' in tt_cuoi.upper()
    p2 = 'nhược' not in str(E.gia_tri('DUNG_HY_KY', 'D6')).lower()
    pct = E.gia_tri('PP2_DINH_LUONG', 'C101') or 0
    if (not nhuoc) != p2: tin = 'Thấp'
    elif vs['can_ranh_gioi'] or 0.35 <= pct < 0.45: tin = 'Vừa'
    else: tin = 'Cao'
    out['vuong_suy'] = dict(ket='NHƯỢC' if nhuoc else 'VƯỢNG', tin_cay=tin, pp2='VƯỢNG' if p2 else 'NHƯỢC', buoc=vs['buoc'])
    # 2. CÁCH CỤC
    kn = E.gia_tri('DOI_CHIEU', 'C33'); col = 'D' if kn == 2 else 'C'
    cach = str(E.gia_tri('DUNG_HY_KY', col + '102')).replace('Cách đặc biệt: ', '')
    ly_cach = ''
    if cach == 'Tòng nhi' and _goc_manh(tt): cach, ly_cach = 'Chính cách', 'CĐB-1: nhật can có gốc mạnh nên không tòng nhi'
    if 'độc vượng' in cach.lower() or cach == 'Tòng vượng (độc vượng)': cach, ly_cach = 'Chính cách', 'CĐB-3: độc vượng PP2 không tự thành cách'
    if cach in ('Khúc trực', 'Viêm thượng', 'Gia tường', 'Tòng cách', 'Nhuận hạ') and _so_chu_phe_minh(tt) < 5:
        cach, ly_cach = 'Chính cách', 'CĐB-4: hành mình và hành sinh mình chưa tới 5/8 chữ'
    out['cach'] = dict(ten=cach, ly=ly_cach)
    # 3. DỤNG THẦN
    nh = HANH_CAN[CAN.index(tt.split(',')[2].split()[0])]
    if cach != 'Chính cách':
        dung = str(E.gia_tri('DUNG_HY_KY', col + '106')).split()[0]; hy = str(E.gia_tri('DUNG_HY_KY', col + '107')).split()[0]; lyd = 'theo cách %s (P2 tr.381)' % cach
    else:
        L = vai_tro(E)['luong2']; a = [L.get(x) or 0 for x in G]; t = sum(a) or 1
        # thứ tự ứng viên LẤY TỪ WORKBOOK (DUNG_HY_KY dòng 91, 93–95) – không tự tính lại
        Lu = [int(E.gia_tri('DUNG_HY_KY', col + str(r))) for r in (93, 94, 95)]
        Lu = [g for g in Lu if g >= 0] or [int(E.gia_tri('DUNG_HY_KY', col + '96'))]
        ly1 = 'Trường hợp %s, thứ tự ứng viên %s (P1 tr.99)' % (E.gia_tri('DUNG_HY_KY', col + '91'), ' → '.join(G[g] for g in Lu))
        co = _co_mat(tt)
        if thu_tu_pp2 or vs['ket'] != vs['so_bo']:   # nhóm 'nhiều' theo lượng PP2, hoặc VS-2 đổi trạng thái
            Lu, ly1, _ = ung_vien(nhuoc, a)
        hop = [g for g in Lu if co[g] > 0] or Lu
        d = hop[0]; lyd = '%s → %s' % (ly1, G[d]) + ('' if d == Lu[0] else ' (%s khuyết – DT-2a, P1 tr.99)' % G[Lu[0]])
        dh, ly_dh = DV.can_dieu_hau(tt)
        if dh is not None:
            gdh = (dh - nh) % 5
            if a[gdh] / t < 0.2 and gdh != d:
                d = gdh; lyd = ly_dh + ' → dụng %s (DT-3)' % HT[dh]
        con = [g for g in hop if g != d] + [g for g in Lu if g not in hop and g != d]
        dung = HT[(nh + d) % 5]; hy = HT[(nh + (con[0] if con else (d + 4) % 5)) % 5]
    out['dung_than'] = dict(dung=dung, hy=hy, ly=lyd)
    # 4. ĐẠI VẬN (DV-3, dùng dụng thần mới) + lĩnh vực (TL-1)
    kq = DV.luan_la(E, tt)
    # thay dụng thần mới cho mức rất tốt / rất xấu
    dmoi = HT.index(dung)
    p_ = [x.split() for x in tt.split(',')]; lo = [thap_than(p_[2][0], p_[i][0]) for i in (0, 1, 3)]
    dh, ly_dh = DV.can_dieu_hau(tt)
    for i, col_ in enumerate('CDEFGHIJ'):
        can, chi = str(E.gia_tri('DAI_VAN', col_ + '35')).split()
        d1, l1 = DV._nua(p_[2][0], nhuoc, can, True, dmoi, dh, ly_dh, lo)
        d2, l2 = DV._nua(p_[2][0], nhuoc, chi, False, dmoi, dh, ly_dh, lo)
        v = kq['van'][i]
        if (d1 > 0) == (d2 > 0): v['ket'] = DV.MUC[d1 if abs(d1) >= abs(d2) else d2]; v['tom'] = 'cả 10 năm %s' % v['ket'].lower()
        else: v['ket'] = DV.MUC[d1]; v['tom'] = '5 năm đầu %s, 5 năm sau %s' % (DV.MUC[d1].lower(), DV.MUC[d2].lower())
        v['chinh'] = 'T' if d1 > 0 else 'X'; v['vi_sao'] = ['5 năm đầu: ' + l1 + '.', '5 năm sau: ' + l2 + '.']
    kq['dung'] = dung
    DV.linh_vuc(E, tt, gioi, kq)
    out['dai_van'] = kq['van']
    # 5. B1 bệnh – mức, B2 gần tòng
    bm = E.gia_tri('DUNG_HY_KY', 'C176')
    if bm == 2 and ((pct >= 0.6 and nhuoc) or (pct < 0.3 and not nhuoc)): bm = 1
    out['benh_muc'] = bm
    out['gan_tong'] = (pct < 0.15 and nhuoc and cach == 'Chính cách')
    return out
