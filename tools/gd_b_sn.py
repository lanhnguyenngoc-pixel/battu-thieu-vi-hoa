import sys, collections, pickle
sys.path.insert(0, '/home/claude/battu/tools'); sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from gd_b_dac_trung import chay
from tt_co_ban import *
X = chay(lv_list=('SN',))
HOP = {frozenset(p) for p in [('Giáp','Kỷ'),('Ất','Canh'),('Bính','Tân'),('Đinh','Nhâm'),('Mậu','Quý')]}
KD = {'Giáp':'Mão','Bính':'Ngọ','Mậu':'Ngọ','Canh':'Dậu','Nhâm':'Tí'}
def dh(r):
    a = phan_tich(r['tu_tru']); lo, bk, tg, can, chi = a['lo'], a['bk'], a['tang'], a['can'], a['chi']
    ro = lo + bk; tat = ro + tg; nhuoc = 'NHƯỢC' in r['than'].upper()
    qs_ro = sum(x in ('Quan','Sát') for x in ro); qs_tat = sum(x in ('Quan','Sát') for x in tat)
    an_ro = sum(x in ('Ấn','Kiêu') for x in ro); tt_ro = sum(x in ('Thực','Thương') for x in ro)
    q_hop = any(a['lo_pos'].get(i) in ('Quan','Sát') and any(frozenset((can[i], can[j])) in HOP for j in (i-1, i+1) if 0 <= j < 4 and j != i) for i in (0, 1, 3))
    q_ts = [truong_sinh(c, chi[1]) for i, c in enumerate(can) if i != 2 and thap_than(a['nc'], c) in ('Quan','Sát')]
    x = {}; t = {}
    x['Q_khong'] = qs_tat == 0
    x['Q_tang'] = qs_tat > 0 and qs_ro == 0                        # quan tàng không lộ
    x['TQ_gap_Q'] = 'Thương' in lo and qs_ro > 0 and an_ro == 0     # thương quan gặp quan, không ấn chế
    x['Hon_tap'] = 'Quan' in ro and 'Sát' in ro
    x['Q_hop'] = q_hop
    x['Q_suy'] = any(s in ('Tử','Mộ','Tuyệt','Thai','Bệnh') for s in q_ts)
    x['Than_nhuoc_QS_nhieu'] = nhuoc and qs_ro >= 2
    x['Sat_che_qua'] = 'Sát' in ro and tt_ro >= 3
    t['Sat_an'] = qs_ro > 0 and an_ro > 0 and not x['TQ_gap_Q']
    t['TQ_phoi_an'] = 'Thương' in lo and an_ro > 0
    t['KD_sat'] = a['nc'] in KD and KD[a['nc']] in chi and 'Sát' in ro
    t['Q_lenh'] = bk[1] in ('Quan','Sát') and not nhuoc
    t['Khoi_canh'] = can[2] + ' ' + chi[2] in ('Canh Thìn','Canh Tuất','Nhâm Thìn','Mậu Tuất')
    return x, t
def tom(r): r['x'], r['t'] = dh(r); r['nx'] = sum(r['x'].values()); r['nt'] = sum(r['t'].values())
for r in X: tom(r)
tong = collections.Counter(r['sach'] for r in X); print('Nhãn sách:', dict(tong))
for nm in ('x', 't'):
    st = collections.defaultdict(collections.Counter)
    for r in X:
        for k, v in r[nm].items():
            if v: st[k][r['sach']] += 1
    for k, c in st.items(): print('%-3s %-20s K %2d  V %2d  T %2d' % (nm, k, c['K'], c['V'], c['T']))
def dd(r, a=1):
    s = r['nt'] - r['nx']
    return 'T' if s >= a else ('K' if s <= -a else 'V')
for a in (1, 2):
    c = collections.Counter((r['sach'], dd(r, a)) for r in X)
    print('ngưỡng ±%d:' % a, {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'}, 'tổng', sum(c[(s, s)] for s in 'KVT'), '/', len(X))
Q = pickle.load(open('/home/claude/battu/tai_lieu/giai_doan_A/quan_the.pkl', 'rb'))
for q in Q: q2 = dict(q, qt=q['qt']['SN']); tom(q2); q['sn'] = dd(q2, 1)
print('Quần thể – công cụ:', collections.Counter(q['o']['SN'] for q in Q), '| đề xuất ±1:', collections.Counter(q['sn'] for q in Q))
print('Hiện tại trên ví dụ sách:', {s: '%d/%d' % (sum(r['cong_cu'] == s for r in X if r['sach'] == s), tong[s]) for s in 'KVT'})
