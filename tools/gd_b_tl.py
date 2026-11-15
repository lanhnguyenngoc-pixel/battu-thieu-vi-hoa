import sys, collections, pickle
sys.path.insert(0, '/home/claude/battu/tools'); sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from gd_b_dac_trung import chay
from tt_co_ban import *
X = chay(lv_list=('TL',))
MO = {0: 'Mùi', 1: 'Tuất', 2: 'Thìn', 3: 'Sửu', 4: 'Thìn'}   # kho của hành (Mộc Hỏa Thổ Kim Thủy)
def dh(r):
    a = phan_tich(r['tu_tru']); lo, bk, tg, can, chi = a['lo'], a['bk'], a['tang'], a['can'], a['chi']
    ro = lo + bk; tat = ro + tg; nhuoc = 'NHƯỢC' in r['than'].upper(); vuong = 'VƯỢNG' in r['than'].upper()
    ta = lambda L: sum(x in ('CTài','TTài') for x in L)
    tk = lambda L: sum(x in ('Tỉ','Kiếp') for x in L)
    hanh_tai = (HANH_CAN[CAN.index(a['nc'])] + 2) % 5
    ta_can = [c for i, c in enumerate(can) if i != 2 and thap_than(a['nc'], c) in ('CTài','TTài')]
    ta_ts = [truong_sinh(c, chi[1]) for c in ta_can]
    qt = r['qt']; tai_dung = qt.get('TV4', (False,))[0] if 'TV4' in qt else False
    x = {}; t = {}
    x['Ta_khong'] = ta(tat) == 0
    x['TK_doat'] = tk(ro) >= 3 and ta(ro) <= 1 and not any(v in ('Quan','Sát') for v in ro)
    x['Nhuoc_ta_nhieu'] = nhuoc and ta(ro) >= 3
    x['Ta_tuyet'] = bool(ta_ts) and all(s in ('Tử','Mộ','Tuyệt','Bệnh') for s in ta_ts)
    x['Ta_tang_khong_goc'] = ta(lo) >= 1 and ta(bk) == 0 and ta(tg) == 0
    t['Thuc_sinh_ta'] = any(v in ('Thực','Thương') for v in ro) and ta(ro) >= 1 and not nhuoc
    t['Ta_kho'] = ta(tat) >= 1 and MO[hanh_tai] in chi
    t['Vuong_ta_vuong'] = vuong and (ta(ro) >= 2 or bk[1] in ('CTài','TTài'))
    t['Ta_ngay'] = bk[2] in ('CTài','TTài') and not nhuoc
    t['Ta_lenh'] = bk[1] in ('CTài','TTài') and not nhuoc
    return x, t
def tom(r): r['x'], r['t'] = dh(r); r['nx'] = sum(r['x'].values()); r['nt'] = sum(r['t'].values())
for r in X: tom(r)
tong = collections.Counter(r['sach'] for r in X); print('Nhãn sách:', dict(tong))
for nm in ('x', 't'):
    st = collections.defaultdict(collections.Counter)
    for r in X:
        for k, v in r[nm].items():
            if v: st[k][r['sach']] += 1
    for k, c in st.items(): print('%-3s %-18s K %2d  V %2d  T %2d' % (nm, k, c['K'], c['V'], c['T']))
def dd(r, a=2, b=1):
    s = r['nt'] - r['nx']
    return 'T' if s >= a else ('K' if s <= -b else 'V')
for a, b in ((1, 1), (2, 1), (2, 2), (3, 1)):
    c = collections.Counter((r['sach'], dd(r, a, b)) for r in X)
    print('T nếu ≥%d, K nếu ≤-%d:' % (a, b), {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'}, 'tổng', sum(c[(s, s)] for s in 'KVT'), '/', len(X))
Q = pickle.load(open('/home/claude/battu/tai_lieu/giai_doan_A/quan_the.pkl', 'rb'))
for q in Q: q2 = dict(q, qt=q['qt']['TL']); tom(q2); q['tl'] = dd(q2, 2, 1)
print('Quần thể – công cụ:', collections.Counter(q['o']['TL'] for q in Q), '| đề xuất (2,1):', collections.Counter(q['tl'] for q in Q))
print('Hiện tại trên ví dụ sách:', {s: '%d/%d' % (sum(r['cong_cu'] == s for r in X if r['sach'] == s), tong[s]) for s in 'KVT'})
print('\n--- Kết hợp có trọng số đều ±1 theo câu sách ---')
CONG = ['TV24','TV11','TV3','TV2']; TRU = ['TV36','TV33','TV34']
def diem(r):
    f = lambda m: r['qt'].get(m, (False,))[0]
    return sum(f(m) for m in CONG) + r['t']['Thuc_sinh_ta'] - sum(f(m) for m in TRU) - r['x']['TK_doat'] - (1 if r['x']['Nhuoc_ta_nhieu'] and not f('TV56') else 0)
def dd3(r, a, b): s = diem(r); return 'T' if s >= a else ('K' if s <= b else 'V')
for a in (2, 3):
    for b in (-1, 0):
        c = collections.Counter((r['sach'], dd3(r, a, b)) for r in X)
        print('T≥%d K≤%d' % (a, b), {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'}, 'tổng %d/%d' % (sum(c[(s, s)] for s in 'KVT'), len(X)))
for q in Q: q2 = dict(q, qt=q['qt']['TL']); tom(q2); q['tl3'] = dd3(q2, 3, 0)
print('Quần thể – đề xuất (T≥3, K≤0):', collections.Counter(q['tl3'] for q in Q))
for q in Q: q2 = dict(q, qt=q['qt']['TL']); tom(q2); q['tl4'] = dd3(q2, 2, -1)
print('Quần thể – đề xuất (T≥2, K≤-1):', collections.Counter(q['tl4'] for q in Q))
print('\n=== LÁ LỆCH (T≥2,K≤-1) ===')
for r in X:
    p = dd3(r, 2, -1)
    if p != r['sach']:
        f = [m for m in CONG + TRU if r['qt'].get(m, (False,))[0]]
        print('%s|%s sách %s → %s (điểm %d) %s | %s | thân %s | %s' % (r['la'][:26], r['tu_tru'], r['sach'], p, diem(r), f, [k for k,v in {**r['x'], **r['t']}.items() if v], r['than'][:14], r['cau'][:110]))
