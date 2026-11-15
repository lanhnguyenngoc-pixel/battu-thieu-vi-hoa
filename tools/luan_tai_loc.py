"""Nguyên mẫu: luận tài lộc theo QUY TRÌNH của tác giả, có lý giải từng bước (chưa vào workbook)."""
import sys
sys.path.insert(0, '/home/claude/battu/tools')
from gd_b_mohinh import co_ban, pct, MO
from tt_co_ban import *
MUC = ['Kém (khó giàu)', 'Vừa (cơm áo không thiếu)', 'Khá (tiểu phú)', 'Tốt (phú)']
def luan(r):
    a = co_ban(r); d = a['dem']; lo, bk, tg, ro, tat, chi, can = a['lo'], a['bk'], a['tang'], a['ro'], a['tat'], a['chi'], a['can']
    P = pct(r); B = []; nc = a['nc']
    B.append(('1. Thân và dụng thần', 'Nhật chủ %s %s; dụng thần %s, hỷ %s, kỵ %s.' % (nc, 'NHƯỢC' if a['nhuoc'] else ('VƯỢNG' if a['vuong'] else 'TRUNG HÒA'), r['role']['dung'], r['role']['hy'], r['role']['ky']),
              'Mọi lời luận của tác giả đều bắt đầu từ thân vượng/nhược và dụng thần (VD1–VD40).'))
    # 2. sức của sao tài
    so_lo = d('TA', lo)
    hanh_tai0 = (HANH_CAN[CAN.index(nc)] + 2) % 5
    CUC = [({'Thân','Tí','Thìn'}, 4, 'Tí'), ({'Hợi','Mão','Mùi'}, 0, 'Mão'), ({'Dần','Ngọ','Tuất'}, 1, 'Ngọ'), ({'Tị','Dậu','Sửu'}, 3, 'Dậu')]
    def chi_hong(c):   # chi nằm trong cục (tam hợp hoặc bán hợp có chi đế vượng) hóa thành hành khắc tài
        for nhom_chi, h, vuong in CUC:
            if c in nhom_chi and (h + 2) % 5 == hanh_tai0 and (vuong in chi and sum(x in nhom_chi for x in chi) >= 2): return True
        return False
    so_goc = sum(1 for c in chi if NHOM[thap_than(nc, TANG[c][0])] == 'TA')
    so_am = sum(1 for c in chi for x in TANG[c][1:] if NHOM[thap_than(nc, x)] == 'TA' and (HANH_CAN[CAN.index(TANG[c][0])] + 2) % 5 != hanh_tai0 and not chi_hong(c))
    so_goc += so_am; so_am = 0
    hanh_tai = (HANH_CAN[CAN.index(nc)] + 2) % 5
    kho = so_lo + so_goc + so_am > 0 and MO[hanh_tai] in chi
    lenh = NHOM[bk[1]] == 'TA'
    ts = [truong_sinh(c, chi[1]) for i, c in enumerate(can) if i != 2 and NHOM[thap_than(nc, c)] == 'TA']
    if so_lo + so_goc + so_am == 0: suc = 'khong'; t2 = 'Tứ trụ KHÔNG có tài (cả lộ lẫn tàng).'
    elif so_lo and not so_goc: suc = 'lo_khong_goc'; t2 = 'Tài lộ ở can nhưng KHÔNG có gốc ở bản khí chi' + (' (chỉ có %d chữ ẩn ở trung/dư khí – khí gốc yếu).' % so_am if so_am else '.')
    elif so_lo and so_goc: suc = 'thau_tang'; t2 = 'Tài "có thấu có tàng": lộ %d chữ ở can, có gốc %d chữ ở bản khí chi.' % (so_lo, so_goc)
    elif so_goc: suc = 'tang'; t2 = 'Tài chỉ tàng ở bản khí chi (%d chữ), không lộ.' % so_goc
    else: suc = 'am'; t2 = 'Tài chỉ ẩn ở trung/dư khí (%d chữ) – rất mờ.' % so_am
    them = []
    if lenh: them.append('tài được lệnh tháng')
    if kho: them.append('có kho tài (%s)' % MO[hanh_tai])
    if ts: them.append('can tài tại tháng ở %s' % '/'.join(ts))
    B.append(('2. Sức của sao tài', t2 + (' ' + '; '.join(them) + '.' if them else '') + ' Tỷ lệ tài trong lá số: %.0f%%.' % P['Tài'],
              '"Tứ trụ có tài thì cái ăn cái mặc không thiếu"; "tài nên tàng ẩn… tài có kho, phát tài thì giữ được" (P2 tr.520); "tài rơi vào không vong tất nghèo" (tr.527). Ví dụ: VD17 lộ không gốc → khó giữ (tr.410); VD27 tài tàng kho → tích lũy giàu (tr.420).'))
    # 3. thân có gánh được tài
    than_pct = P['Tỉ kiếp'] + P['Ấn']; ta_pct = P['Tài']
    if suc == 'khong': ganh = 'khong_tai'; t3 = 'Không có tài nên chưa xét gánh tài.'
    elif a['nhuoc'] and (ta_pct > than_pct or d('TA', ro) >= 3):
        # Trạng thái thân lấy từ kết luận cuối (PP1/VS-2 + G11); % PP2 chỉ dùng để so LƯỢNG tài với tỉ kiếp + ấn
        # (thử 15/11/2026 thay bằng đếm chữ: 35 → 34/65, vì đếm chữ không phân biệt chữ có lực / hư phù → giữ %)
        ganh = 'khong_noi'; t3 = 'Thân nhược (tỉ kiếp + ấn %.0f%%) mà tài nhiều (%.0f%%): KHÔNG gánh nổi tài.' % (than_pct, ta_pct)
    elif not a['nhuoc'] and (lenh or so_lo + so_goc >= 2 or (so_goc >= 1 and ta_pct >= 15)):
        ganh = 'ganh_duoc'; t3 = 'Thân không nhược (%.0f%%) và tài có lực (%.0f%%): thân GÁNH ĐƯỢC tài.' % (than_pct, ta_pct)
    elif not a['nhuoc']: ganh = 'than_vuong_tai_yeu'; t3 = 'Thân vượng (%.0f%%) nhưng tài yếu (%.0f%%).' % (than_pct, ta_pct)
    else: ganh = 'can_bang'; t3 = 'Thân nhược nhưng tài không nhiều (%.0f%% so với %.0f%%): tạm cân bằng.' % (ta_pct, than_pct)
    B.append(('3. Thân có gánh được tài?', t3, '"Thân vượng tài vượng là phúc"; "tài nhiều thân nhược cần ấn thụ hỗ trợ"; "thân suy không thắng tài là người vất vả" (P2 tr.520–521); "tài nhiều thân nhược là nhà cửa nhiều nhưng của cải ít" (tr.527, ứng P1-VD10).'))
    # 4. vai trò
    vt = a['vai']['TA']; tv = {'dung': 'DỤNG thần', 'hy': 'HỶ thần', 'ky': 'KỴ thần', 'nhan': 'nhàn thần'}[vt]
    B.append(('4. Tài là gì đối với lá số', 'Tài là %s.' % tv, '"Tài làm dụng thần là người giàu có" (P2 tr.520); tài là kỵ thì "rất khó giàu sang" (VD13, tr.407).'))
    # 5. nguồn sinh và kẻ phá
    sinh = d('TT', ro) >= 1 and suc != 'khong'; doat = d('TK', ro) >= 3 and d('TA', ro) <= 1; che = d('QS', ro) >= 1; an_ho = d('AN', ro) >= 1
    t5 = []
    if sinh: t5.append('có thực thương sinh tài')
    if doat: t5.append('tỉ kiếp nhiều (%d chữ rõ) tranh đoạt tài' % d('TK', ro) + (', nhưng có quan sát chế tỉ kiếp' if che else ', không có quan sát chế'))
    if ganh == 'khong_noi' and an_ho: t5.append('có ấn hỗ trợ thân')
    kieu_pha = d('AN', ro) >= 3 and suc != 'khong' and d('TK', ro) == 0
    thuc_qua = d('TT', ro) >= 3 and so_lo == 0 and a['nhuoc']
    to_hop = all(d(g, ro) >= 1 for g in ('TA', 'QS', 'AN'))
    if kieu_pha: t5.append('kiêu ấn nặng (%d chữ) gặp tài mà không có tỉ kiếp' % d('AN', ro))
    if thuc_qua: t5.append('thực thương quá nhiều (%d chữ) mà tài không hiện, thân nhược' % d('TT', ro))
    if to_hop: t5.append('tài – quan – ấn đều có')
    B.append(('5. Nguồn sinh tài và kẻ phá tài', ('; '.join(t5) or 'không có nguồn sinh rõ, không có kẻ phá rõ') + '.',
              '"Thương quan sinh tài thì tự nhiên giàu có"; "thân vượng tài nhược, được thực thương sinh tài mới giàu" (P2 tr.520–521); "tỉ kiếp tranh tài", "kiêu ấn gặp đất tài không có tỷ kiếp thì tai hoạ" (tr.527); "đủ tài quan ấn không sang thì giàu" (tr.521).'))
    # 6–7. mức gốc theo cách tác giả xử lý xung đột
    ly = []
    if ganh == 'khong_tai':
        m = 0 if (a['vuong'] and d('TK', ro) >= 3) else 1
        ly.append('Không có tài: mức tùy vận (VD3 tr.399: "gặp tài vận có thể phát"); ' + ('thân vượng tỉ kiếp nhiều không tài → kém (P2 tr.527)' if m == 0 else 'mức gốc VỪA'))
    elif suc == 'lo_khong_goc' and ganh != 'ganh_duoc' and not kho:
        m = 0; ly.append('Tài lộ không gốc, thân không gánh được, không có kho → khó giữ của (VD17 tr.410: gốc của tài thắng vai trò của tài)')
    elif ganh == 'khong_noi':
        m = 1 if an_ho else 0; ly.append('Thân không gánh nổi tài → ' + ('có ấn hộ thân nên VỪA' if an_ho else 'KÉM ("nhà nhiều của ít")'))
    elif ganh == 'ganh_duoc':
        m = 2; ly.append('Thân gánh được tài → KHÁ')
        if (vt in ('dung', 'hy') and (suc == 'thau_tang' or kho)) or (suc == 'thau_tang' and kho): m = 3; ly.append('tài lại là dụng/hỷ và có thấu có tàng hoặc có kho → TỐT ("thân vượng tài vượng là phúc"; "tài có kho phát tài thì giữ được")')
    elif ganh == 'than_vuong_tai_yeu':
        m = 2 if sinh else 1; ly.append('Thân vượng tài yếu → ' + ('có thực thương sinh tài nên KHÁ' if sinh else 'không có thực thương sinh tài nên VỪA'))
    else:
        m = 1; ly.append('Tạm cân bằng → VỪA')
    if suc == 'lo_khong_goc' and m > 1: m -= 1; ly.append('tài lộ không gốc → khó tích lũy, hạ 1 bậc (VD17)')
    if vt == 'ky' and m > 1: m -= 1; ly.append('tài là kỵ thần → hạ 1 bậc, nhưng không dưới VỪA: "phải gặp vận thân vượng mới tụ tài" (VD23 tr.416, VD30 tr.422)')
    if doat and not che and m > 0: m -= 1; ly.append('tỉ kiếp đoạt tài không có quan chế → hạ 1 bậc (VD4)')
    if kieu_pha and m > 0: m -= 1; ly.append('"kiêu ấn gặp đất tài lại không có tỷ kiếp thì tai hoạ" (P2 tr.527, ứng C18-14) → hạ 1 bậc')
    if thuc_qua and m > 0: m -= 1; ly.append('thực thương quá nhiều, tài không hiện (VD34 tr.426) → hạ 1 bậc')
    if to_hop and m < 3 and suc != 'lo_khong_goc': m += 1; ly.append('"trong trụ đủ cả tài, quan, ấn là người không sang thì giàu" (P2 tr.521) → nâng 1 bậc')
    if kho and ganh in ('khong_noi', 'can_bang') and m < 3: m += 1; ly.append('có kho tài giữ được của → nâng 1 bậc (VD21, VD27)')
    B.append(('6. Gộp các điều kiện (cách tác giả xử lý khi trái nhau)', '; '.join(ly) + '.', 'Thứ tự ưu tiên rút từ ví dụ: gốc của tài > cân bằng thân–tài > vai trò dụng/kỵ > nguồn sinh/kẻ phá.'))
    B.append(('7. KẾT LUẬN MỨC GỐC', MUC[m], 'Mức gốc là tiềm năng cả đời; thời điểm phát/hao xem bước 8 (đại vận).'))
    return m, B
if __name__ == '__main__':
    from gd_b_dac_trung import chay
    import collections
    X = chay(lv_list=('TL',))
    quy = lambda m: 'K' if m == 0 else ('V' if m == 1 else 'T')
    c = collections.Counter((r['sach'], quy(luan(r)[0])) for r in X); tong = collections.Counter(r['sach'] for r in X)
    print('KIỂM TRA (không chỉnh ngưỡng): khớp %d/%d' % (sum(c[(s, s)] for s in 'KVT'), len(X)), {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'})
    for ten in sys.argv[1:]:
        r = next(x for x in X if x['la'].startswith(ten))
        m, B = luan(r); print('\n=====', r['la'], r['tu_tru'], '| sách:', r['sach'], '–', r['cau'][:90])
        for t, n, cs in B: print('■ %s\n   %s\n   Căn cứ: %s' % (t, n, cs))
    from nguyen_nhan_lech import phan_loai
    lech = [r for r in X if quy(luan(r)[0]) != r['sach']]
    nn = collections.Counter(p.split(' (')[0] for r in lech for p in phan_loai(r, None))
    print('Nguyên nhân %d lá lệch:' % len(lech), dict(nn))
    for r in lech:
        if phan_loai(r, None) == ['quy trình lĩnh vực chưa đúng']:
            m, B = luan(r); print('  -', r['la'][:22], r['sach'], '→', quy(m), '|', B[1][1][:70], '|', B[5][1][:80], '|| SÁCH:', r['cau'][:80])
