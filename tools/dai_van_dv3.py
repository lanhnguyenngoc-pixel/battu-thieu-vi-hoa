"""Mẫu thử DV-3: luận đại vận = quy tắc chung + quy tắc riêng khi lá số có điều kiện đặc biệt,
cho MỘT kết luận cuối mỗi vận, kèm chuỗi lý giải ngắn (không liệt kê ghi chú rời)."""
import sys
sys.path.insert(0, '/home/claude/battu/tools')
from tt_co_ban import thap_than, TANG, NHOM, HANH_CAN, CAN
from vuong_suy_vs2 import trang_thai_cuoi

HT = ['Mộc', 'Hỏa', 'Thổ', 'Kim', 'Thủy']
TEN = {'TK': 'tỉ kiếp', 'TT': 'thực thương', 'TA': 'tài', 'QS': 'quan sát', 'AN': 'ấn'}
HAN = {'Hợi', 'Tí', 'Sửu', 'Dần'}; NONG = {'Tị', 'Ngọ', 'Mùi'}
MUC = {2: 'RẤT TỐT', 1: 'TỐT', 0: 'BÌNH THƯỜNG', -1: 'XẤU', -2: 'RẤT XẤU'}


def can_dieu_hau(tt):
    """ĐB-1 (P1 tr.76–79): kim/thủy/mộc sinh mùa lạnh cần hỏa sưởi ấm; hỏa/thổ/kim sinh mùa nóng cần thủy."""
    nh = HANH_CAN[CAN.index(tt.split(',')[2].split()[0])]; th = tt.split(',')[1].split()[1]
    if th in HAN and nh in (0, 3, 4): return 1, '%s sinh tháng %s (mùa lạnh) cần hỏa sưởi ấm (P1 tr.76, 78, 79)' % (HT[nh], th)
    if th in NONG and nh in (1, 2, 3): return 4, '%s sinh tháng %s (mùa nóng) cần thủy điều hòa (P1 tr.77, 78)' % (HT[nh], th)
    return None, ''


def _nua(nc, nhuoc, chu, la_can, dung, dh, ly_dh, lo):
    goc = chu if la_can else TANG[chu][0]
    t = thap_than(nc, goc); g = NHOM[t]; h = HANH_CAN[CAN.index(goc)]
    thuan = (g in ('TK', 'AN')) == nhuoc
    d = 1 if thuan else -1
    ly = ['%s %s là %s → %s với thân %s (P1 tr.125)' % ('Can' if la_can else 'Chi', chu, t, 'thuận' if thuan else 'nghịch', 'nhược' if nhuoc else 'vượng')]
    # ĐB-1: điều hầu
    if dh is not None and h == dh and not thuan:
        d = 1; ly.append('nhưng %s đúng là hành điều hầu cần thiết (%s) → chuyển thành tốt' % (HT[h], ly_dh))
    # ĐB-2: thương quan gặp quan
    if t == 'Thương' and 'Quan' in lo and not any(NHOM[x] == 'AN' for x in lo):
        d = -1; ly.append('nhưng là thương quan gặp chính quan lộ trong trụ, không có ấn hóa → "kị nhất" (P2 tr.541–542) → xấu')
    # mức: dụng thần lên ngôi / khắc dụng
    if dung is not None and d > 0 and h == dung:
        d = 2; ly.append('hành %s là dụng thần lên ngôi (P1 tr.124) → rất tốt' % HT[h])
    elif dung is not None and d < 0 and (h + 2) % 5 == dung:
        d = -2; ly.append('hành %s khắc dụng thần %s (P1 tr.125) → rất xấu' % (HT[h], HT[dung]))
    return d, '; '.join(ly)


def luan_la(E, tt):
    tt_cuoi, vs = trang_thai_cuoi(E); nhuoc = 'NHƯỢC' in tt_cuoi.upper()
    p = [x.split() for x in tt.split(',')]; nc = p[2][0]
    lo = [thap_than(nc, p[i][0]) for i in (0, 1, 3)]
    kn = E.gia_tri('DOI_CHIEU', 'C33')
    w = str(E.gia_tri('DUNG_HY_KY', 'D106' if kn == 2 else 'C106')).split()
    dung = HT.index(w[0]) if w and w[0] in HT and vs['ket'] == vs['so_bo'] else None
    dh, ly_dh = can_dieu_hau(tt)
    out = []
    for col in 'CDEFGHIJ':
        can, chi = str(E.gia_tri('DAI_VAN', col + '35')).split()
        tuoi = int(float(E.gia_tri('DAI_VAN', col + '36') or 0))
        d1, l1 = _nua(nc, nhuoc, can, True, dung, dh, ly_dh, lo)
        d2, l2 = _nua(nc, nhuoc, chi, False, dung, dh, ly_dh, lo)
        if (d1 > 0) == (d2 > 0):
            ket = MUC[d1 if abs(d1) >= abs(d2) else d2]; tom = 'cả 10 năm %s' % ket.lower()
        else:
            ket = MUC[d1]; tom = '5 năm đầu %s, 5 năm sau %s' % (MUC[d1].lower(), MUC[d2].lower())
        out.append(dict(van='%s %s' % (can, chi), tuoi=tuoi, ket=ket, chinh='T' if d1 > 0 else 'X', tom=tom,
                        vi_sao=['5 năm đầu: ' + l1 + '.', '5 năm sau: ' + l2 + '.']))
    return dict(than=tt_cuoi, dung=HT[dung] if dung is not None else None, dieu_hau=ly_dh, van=out)


def in_ra(kq, so=8):
    print('Thân: %s. Dụng thần: %s.%s' % (kq['than'], kq['dung'], (' Điều hầu: ' + kq['dieu_hau'] + '.') if kq['dieu_hau'] else ''))
    for v in kq['van'][:so]:
        print('■ Vận %s (từ %d tuổi): %s – %s' % (v['van'], v['tuoi'], v['ket'], v['tom']))
        for s in v['vi_sao']: print('    Vì: ' + s)


KD = {'Giáp': 'Mão', 'Bính': 'Ngọ', 'Mậu': 'Ngọ', 'Canh': 'Dậu', 'Nhâm': 'Tí'}
XUNG = {frozenset(p) for p in [('Tí', 'Ngọ'), ('Sửu', 'Mùi'), ('Dần', 'Thân'), ('Mão', 'Dậu'), ('Thìn', 'Tuất'), ('Tị', 'Hợi')]}


def linh_vuc(E, tt, gioi, kq):
    """Kết luận từng lĩnh vực cho mỗi vận: lấy kết luận chung của vận, rồi áp quy tắc riêng của lĩnh vực
    khi lá số có điều kiện – ra MỘT kết luận, lý do ngắn; câu sách để trong 'nguon' (ẩn, bấm mới xem)."""
    p = [x.split() for x in tt.split(',')]; nc = p[2][0]; chi_ngay = p[2][1]
    nhuoc = 'NHƯỢC' in kq['than'].upper(); nu = gioi == 'Nữ'
    lo = [thap_than(nc, p[i][0]) for i in (0, 1, 3)]
    for v in kq['van']:
        can, chi = v['van'].split()
        t1 = thap_than(nc, can); t2 = thap_than(nc, TANG[chi][0]); g1, g2 = NHOM[t1], NHOM[t2]
        # TÀI LỘC – quy tắc chung của vận + TL-1
        tl, ly, ng = v['chinh'], 'theo xu thế chung của vận', ''
        if not nhuoc and g2 == 'TA':
            tl, ly, ng = 'T', 'thân vượng, chi vận là tài → 5 năm sau có tài', '"Trong trụ có tài, hành đến đất tài tất đại phát" (P2 tr.522)'
        elif nhuoc and g1 == 'TA' and g2 in ('TK', 'AN'):
            tl, ly, ng = 'T', 'vận tài nhưng chi vận giúp thân gánh được tài → có thể tụ tài', '"Tài vận phải gặp vận thân vượng mới tụ tài" (VD23 P2 tr.416); "thân nhược tài nhiều, hành vận tỉ kiếp mới phát tài" (P2 tr.521)'
        elif not nhuoc and g1 == 'TK':
            tl, ly, ng = 'X', 'thân vượng gặp vận tỉ kiếp → tranh tài, hao tài', '"Thân vượng hành đến vận tỷ kiếp thì phá tài bại nghiệp" (P2 tr.527)'
        elif nhuoc and g1 == 'TA':
            tl, ly, ng = 'X', 'thân nhược gặp vận tài → tiền đến nhưng khó giữ', '"Thân nhược không thắng tài thì vì của mà sinh nạn" (P2 tr.527)'
        # CÔNG DANH – quy tắc chung + ĐB-2 (đã có trong kết luận chung)
        sn, lys, ngs = v['chinh'], 'theo xu thế chung của vận', ''
        # (đã thử quy tắc riêng "thân vượng/nhược gặp vận quan sát" – kém hơn kết luận chung: 11 so với 13/18 → không dùng)
        if any('thương quan gặp chính quan' in x for x in v['vi_sao']):
            lys, ngs = 'thương quan gặp quan → dễ mất chức, thị phi', '"Thương quan khi thương hết thì kị nhất lại gặp quan tinh" (P2 tr.541–542)'
        # HÔN NHÂN – sự việc
        hn = []
        if t1 == 'Thương' or t2 == 'Thương': hn.append(('−', 'vận thương quan → %s' % ('khắc chồng' if nu else 'khắc vợ'), '(P2 tr.471)'))
        if frozenset((chi, chi_ngay)) in XUNG: hn.append(('−', 'chi vận xung cung phối ngẫu', '(P2 tr.462)'))
        if nc in KD and chi == KD[nc]: hn.append(('−', 'vận kình dương', '(P2 tr.468)'))
        if (g1 == 'QS' or g2 == 'QS') if nu else (g1 == 'TA' or g2 == 'TA'): hn.append(('+', 'gặp sao phối ngẫu → dễ có duyên', '(P1 tr.153–157)'))
        s = sum(1 if x[0] == '+' else -1 for x in hn)
        hk = 'THUẬN' if s > 0 else ('TRẮC TRỞ' if s < 0 else 'BÌNH THƯỜNG')
        v['linh_vuc'] = dict(
            tai_loc=dict(ket='TỐT' if tl == 'T' else 'XẤU', vi=ly, nguon=ng),
            cong_danh=dict(ket='TỐT' if sn == 'T' else 'XẤU', vi=lys, nguon=ngs),
            hon_nhan=dict(ket=hk, vi='; '.join(x[1] for x in hn) or 'không có dấu hiệu riêng', nguon=' '.join(x[2] for x in hn)))
    return kq
