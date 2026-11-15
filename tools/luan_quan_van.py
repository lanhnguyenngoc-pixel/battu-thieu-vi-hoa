"""Nguyên mẫu: luận QUAN VẬN theo quy trình của tác giả, có lý giải từng bước (chưa vào workbook)."""
import sys, collections
sys.path.insert(0, '/home/claude/battu/tools')
from gd_b_mohinh import co_ban, pct
from tt_co_ban import *
MUC = ['Kém (khó làm quan / chức không bền)', 'Vừa (chức nhỏ, chức phó)', 'Khá (cán bộ có chức)', 'Tốt (nắm quyền lớn)']
HOP = {frozenset(p) for p in [('Giáp','Kỷ'),('Ất','Canh'),('Bính','Tân'),('Đinh','Nhâm'),('Mậu','Quý')]}
KD = {'Giáp':'Mão','Bính':'Ngọ','Mậu':'Ngọ','Canh':'Dậu','Nhâm':'Tí'}
KHOI = {'Canh Thìn','Canh Tuất','Nhâm Thìn','Mậu Tuất'}
def luan(r):
    a = co_ban(r); d = a['dem']; lo, bk, tg, ro, tat, chi, can, nc = a['lo'], a['bk'], a['tang'], a['ro'], a['tat'], a['chi'], a['can'], a['nc']
    P = pct(r); B = []
    B.append(('1. Thân và dụng thần', 'Nhật chủ %s %s; dụng thần %s, hỷ %s, kỵ %s.' % (nc, 'NHƯỢC' if a['nhuoc'] else 'VƯỢNG', r['role']['dung'], r['role']['hy'], r['role']['ky']),
              'Lời luận quan vận của tác giả luôn đặt quan sát trong tương quan với thân (VD12, VD14, Ch.19 VD7).'))
    # 2. sức quan sát
    q_lo = [i for i, t in a['lo_pos'].items() if NHOM[t] == 'QS']; q_bk = d('QS', bk); q_am = d('QS', tg)
    ts = [truong_sinh(can[i], chi[1]) for i in q_lo]
    if not q_lo and not q_bk and not q_am: suc = 'khong'; t2 = 'Tứ trụ KHÔNG có quan sát.'
    elif not q_lo and not q_bk: suc = 'an'; t2 = 'Quan sát chỉ ẩn ở trung/dư khí (%d chữ) – tàng, không lộ.' % q_am
    elif not q_lo: suc = 'tang'; t2 = 'Quan sát tàng ở bản khí chi (%d chữ), không lộ ra can.' % q_bk
    else: suc = 'lo'; t2 = 'Quan sát lộ %d chữ ở can%s.' % (len(q_lo), ', có gốc ở chi' if q_bk else ', không có gốc bản khí')
    suy = bool(ts) and all(s in ('Tử','Mộ','Tuyệt') for s in ts)
    lenh = NHOM[bk[1]] == 'QS'
    if ts: t2 += ' Can quan sát tại tháng ở %s.' % '/'.join(ts)
    if lenh: t2 += ' Quan sát được lệnh tháng.'
    B.append(('2. Sức của quan sát', t2 + ' Tỷ lệ quan sát: %.0f%%.' % P['Quan sát'],
              '"Trong trụ không có quan tinh thì khó mà làm quan"; "quan ở mộ kho là không làm quan" (P2 tr.541); "quan ở sát địa khó giữ chức" (tr.542); "quan tàng không lộ → không có chức to" (VD16 tr.410, VD25 tr.418).'))
    # 3. thân gánh quan
    nhieu = d('QS', ro) >= 2 or P['Quan sát'] >= 30
    if suc == 'khong': g = 'khong'; t3 = 'Không có quan sát nên chưa xét.'
    elif a['nhuoc'] and nhieu and d('AN', ro) == 0: g = 'khong_noi'; t3 = 'Thân nhược mà quan sát nhiều, không có ấn hóa: thân KHÔNG gánh nổi quan.'
    elif a['nhuoc']: g = 'yeu'; t3 = 'Thân nhược' + (' nhưng có ấn hóa quan sát sinh thân.' if d('AN', ro) else '; quan sát không nhiều.')
    else: g = 'duoc'; t3 = 'Thân không nhược: gánh được quan sát.'
    B.append(('3. Thân có gánh được quan?', t3, '"Ngày nhược gặp quan tinh thì thành đạt không lâu" (P2 tr.542); "trong trụ sát vượng, thân vượng là bậc quan thanh liêm cao quý" (tr.533); thân nhược sát vượng bị đè nén (Ch.19 VD7, tr.540).'))
    vt = a['vai']['QS']
    B.append(('4. Quan sát là gì đối với lá số', 'Quan sát là %s.' % {'dung':'DỤNG thần','hy':'HỶ thần','ky':'KỴ thần','nhan':'nhàn thần'}[vt],
              'Quan là dụng/hỷ thì cầu quan thuận (VD2, VD29); là kỵ thì "suốt đời khó đạt quan quý" (VD30 tr.422).'))
    # 5. phá / cứu
    co_an = d('AN', ro) >= 1
    pos_tq = [i for i, t in a['lo_pos'].items() if t == 'Thương']
    pos_an = [i for i, t in a['lo_pos'].items() if NHOM[t] == 'AN'] + [i for i in range(4) if NHOM[bk[i]] == 'AN']
    an_ke_tq = any(abs(i - j) <= 1 for i in pos_tq for j in pos_an)   # ấn cùng trụ hoặc trụ kề thương quan mới chế được (VD1 so với P1-VD8)
    co_chinh_quan = 'Quan' in tat
    tq_gap_q = bool(pos_tq) and co_chinh_quan and not an_ke_tq
    tq_gap_sat = bool(pos_tq) and 'Sát' in ro and not co_chinh_quan
    hai_quan = sum(1 for i in q_lo if a['lo_pos'][i] == 'Quan') >= 2 or sum(1 for i in q_lo if a['lo_pos'][i] == 'Sát') >= 2
    q_hop = any(any(frozenset((can[i], can[j])) in HOP for j in (i - 1, i + 1) if 0 <= j < 4) for i in q_lo)
    hon = 'Quan' in ro and 'Sát' in ro
    hop_sat = hon and any(a['lo_pos'].get(i) == 'Sát' and any(frozenset((can[i], can[j])) in HOP for j in (i-1, i+1) if 0 <= j < 4) for i in (0, 1, 3))
    tq_nhieu = sum(x == 'Thương' for x in ro) >= 3
    sat_khong_che = 'Sát' in ro and d('TT', ro) == 0 and a['nhuoc'] and not co_an
    che_qua = 'Sát' in ro and d('TT', ro) >= 3
    pha, cuu = [], []
    if tq_gap_q: pha.append(('thương quan gặp quan, không có ấn chế', '"thương quan gặp quan là vừa không có quan…" (P2 tr.541); VD1 tr.398'))
    if q_hop: pha.append(('quan tinh bị hợp', '"quan tinh bị hợp là không làm quan hoặc làm nhưng thời gian ngắn" (tr.541); VD27, Ch.19 VD8'))
    if hon and not hop_sat: pha.append(('quan sát hỗn tạp', '"quan sát hỗn tạp thì không có quan, nếu có là giặc" (tr.541); VD13'))
    if tq_nhieu: pha.append(('thương quan nhiều', '"người mà thương quan nhiều thì khó mà làm quan, nếu có thì chức nhỏ" (tr.541)'))
    if sat_khong_che: pha.append(('thất sát không chế, thân nhược', '"thất sát không có chế ngự là tiểu nhân"; "kình dương nhẹ sát nặng không chế chỉ làm quan nhỏ" (tr.536, 541)'))
    if che_qua: pha.append(('thực thương chế sát thái quá', 'VD8 tr.403, VD21 tr.415'))
    if hai_quan: pha.append(('quan (sát) lộ trùng nhiều', '"quan vượng không nên gặp hai quan"; "quan tinh nhiều… sẽ có hại" (tr.541); "quan nhiều là sát… chức phó" (VD36 tr.427)'))
    if suy: pha.append(('quan sát ở tử/mộ/tuyệt', '"quan ở mộ kho là không làm quan" (tr.541); VD17, VD31'))
    an_ro_lo = len(pos_an) >= 1 and (d('AN', bk) >= 1)
    if q_lo and an_ro_lo and not tq_gap_q and not a['nhuoc']: cuu.append(('quan/sát lộ – ấn có gốc tương sinh, thân không nhược', '"thân mạnh gặp thất sát phục ấn là quan trấn thủ một vùng" (tr.532); "sát ấn đồng ngôi… cầm quyền" (tr.535); P1-VD15'))
    if tq_gap_sat: cuu.append(('thương quan gặp thất sát (không có chính quan)', '"thương quan, thất sát, nam gặp được là thăng quan" (tr.534); "thương quan hợp với sát thì danh tiếng lẫy lừng" (tr.555)'))
    if pos_tq and an_ke_tq: cuu.append(('thương quan phối ấn (thương quan có chế)', '"thương quan có chế ngự là uy trấn biên thuỳ" (tr.537); "thương quan có cả ấn, quý không nói hết" (P1-VD15, VD32)'))
    if nc in KD and KD[nc] in chi and any(a['lo_pos'][i] == 'Sát' for i in q_lo) and (not a['nhuoc'] or pos_tq): cuu.append(('kình dương gặp quan sát', '"kình dương gặp quan sát là người cầm quyền" (tr.533); Ch.19 VD1'))
    if hop_sat: cuu.append(('quan sát hỗn tạp nhưng sát bị hợp (hợp sát lưu quan)', '"quan tinh xen với thất sát, nếu hợp với thất sát là quý" (tr.533)'))
    if can[2] + ' ' + chi[2] in KHOI and not a['nhuoc']: cuu.append(('nhật trụ khôi canh, thân vượng', '"trụ ngày có khôi canh, gặp vận thân vượng là làm quan văn" (tr.537); VD37'))
    if all(d(x, ro) >= 1 for x in ('TA', 'QS', 'AN')) and not a['nhuoc'] and q_lo and not any('ấn' in c for c, _ in cuu): cuu.append(('đủ tài – quan – ấn', '"trong trụ có đủ tài, quan, ấn… anh hùng, có uy" (tr.536); "tài vượng sinh quan thì vinh hiển" (tr.534)'))
    t5 = ('Phá: ' + '; '.join(p for p, _ in pha) + '. ' if pha else 'Không có điều phá quan rõ. ') + ('Cứu/quý: ' + '; '.join(c for c, _ in cuu) + '.' if cuu else 'Không có tổ hợp quý.')
    B.append(('5. Điều phá quan và điều cứu/quý', t5, ' · '.join(cs for _, cs in pha + cuu) or '—'))
    # 6. gộp
    ly = []
    if suc == 'khong': m = 0; ly.append('không có quan sát → khó làm quan (tr.541)')
    elif suc == 'an': m = 1; ly.append('quan sát chỉ ẩn → nhiều nhất chức nhỏ')
    elif suc == 'tang': m = 1; ly.append('quan tàng không lộ → không có chức to (VD16, VD25)')
    elif q_bk or lenh: m = 2; ly.append('quan sát lộ ra, có gốc → có thể có chức')
    else: m = 1; ly.append('quan sát lộ nhưng không có gốc bản khí → chức không lớn (cùng cách hiểu "khí gốc" như tài, VD17)')
    if g == 'khong_noi': m -= 1; ly.append('thân không gánh nổi quan → hạ 1 bậc')
    if vt == 'ky' and m > 0: m -= 1; ly.append('quan sát là kỵ thần → hạ 1 bậc')
    so_pha = len(pha); so_cuu = len(cuu)
    if so_pha: m -= min(so_pha, 2); ly.append('%d điều phá → hạ %d bậc' % (so_pha, min(so_pha, 2)))
    if so_cuu and suc != 'khong':
        tang_bac = min(so_cuu, 2 if so_pha == 0 else 1); m += tang_bac; ly.append('%d điều cứu/quý → nâng %d bậc%s' % (so_cuu, tang_bac, ' (khi còn điều phá chỉ nâng tối đa 1)' if so_pha else ''))
    m = max(0, min(3, m))
    B.append(('6. Gộp các điều kiện', '; '.join(ly) + '.', 'Thứ tự rút từ ví dụ: có quan hay không, lộ hay tàng → thân gánh nổi không → điều phá (thương gặp quan, bị hợp, hỗn tạp) → điều cứu (ấn, kình dương, hợp sát) chỉ bù được một phần.'))
    B.append(('7. KẾT LUẬN MỨC GỐC', MUC[m], 'Thời điểm thăng/giáng xem đại vận (bước 8).'))
    return m, B
quy = lambda m: 'K' if m == 0 else ('V' if m == 1 else 'T')
if __name__ == '__main__':
    from gd_b_dac_trung import chay
    X = chay(lv_list=('SN',))
    c = collections.Counter((r['sach'], quy(luan(r)[0])) for r in X); tong = collections.Counter(r['sach'] for r in X)
    print('KIỂM TRA (không chỉnh ngưỡng): khớp %d/%d' % (sum(c[(s, s)] for s in 'KVT'), len(X)), {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'})
    print('Phần mềm hiện tại:', sum(r['cong_cu'] == r['sach'] for r in X))
    if len(sys.argv) > 1 and sys.argv[1] == 'lech':
        for r in X:
            m, B = luan(r)
            if quy(m) != r['sach']: print(r['la'][:14], r['sach'], '→', quy(m), '|', B[1][1][:50], '|', B[4][1][:90], '|', B[5][1][:70], '|| ', r['cau'][:70])
