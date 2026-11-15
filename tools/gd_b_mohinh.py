"""Giai đoạn B – mô hình kết hợp điều kiện theo sách TVH (nghiên cứu, chưa vào workbook)."""
import sys, collections
sys.path.insert(0, '/home/claude/battu/tools'); sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from tt_co_ban import *
MO = {0: 'Mùi', 1: 'Tuất', 2: 'Thìn', 3: 'Sửu', 4: 'Thìn'}
NHOMTEN = {'TK': 'Tỉ kiếp', 'TT': 'Thực thương', 'TA': 'Tài', 'QS': 'Quan sát', 'AN': 'Ấn'}
def co_ban(r):
    a = phan_tich(r['tu_tru']); a['ro'] = a['lo'] + a['bk']; a['tat'] = a['ro'] + a['tang']
    a['nhuoc'] = 'NHƯỢC' in r['than'].upper(); a['vuong'] = 'VƯỢNG' in r['than'].upper() and not a['nhuoc']
    rl = r['role']; a['vai'] = {g: ('dung' if rl['dung'] == NHOMTEN[g] else 'hy' if rl['hy'] == NHOMTEN[g] else 'ky' if rl['ky'] == NHOMTEN[g] else 'nhan') for g in NHOMTEN}
    a['dem'] = lambda g, L: sum(NHOM[x] == g for x in L)
    return a
def tl_dau_hieu(r):
    a = co_ban(r); d = a['dem']; lo, bk, tg, ro, tat, chi = a['lo'], a['bk'], a['tang'], a['ro'], a['tat'], a['chi']
    hanh_tai = (HANH_CAN[CAN.index(a['nc'])] + 2) % 5
    goc = d('TA', bk) + d('TA', tg)
    f = {}
    f['thau_tang'] = d('TA', lo) >= 1 and goc >= 1                       # tài có thấu có tàng
    f['lo_khong_goc'] = d('TA', lo) >= 1 and goc == 0                     # tài lộ không gốc (VD17)
    f['kho'] = d('TA', tat) >= 1 and MO[hanh_tai] in chi                  # tài có kho
    f['lenh'] = NHOM[bk[1]] == 'TA'                                       # tài được lệnh tháng
    f['dung_hy'] = a['vai']['TA'] in ('dung', 'hy')
    f['ky'] = a['vai']['TA'] == 'ky'
    f['thuc_sinh'] = d('TT', ro) >= 1 and d('TA', tat) >= 1
    f['vuong_vuong'] = a['vuong'] and (d('TA', ro) >= 2 or f['lenh'])
    f['khong_tai'] = d('TA', tat) == 0
    f['tk_doat'] = d('TK', ro) >= 3 and d('TA', ro) <= 1 and d('QS', ro) == 0
    f['nhuoc_ta_ky'] = a['nhuoc'] and f['ky'] and d('TA', ro) >= 2
    f['thuc_qua_khong_tai'] = d('TT', ro) >= 3 and d('TA', lo) == 0
    f['ta_ngay'] = NHOM[bk[2]] == 'TA'
    return f
TL_W = {'thau_tang': 2, 'kho': 1, 'lenh': 1, 'dung_hy': 1, 'thuc_sinh': 1, 'vuong_vuong': 1, 'ta_ngay': 1,
        'lo_khong_goc': -2, 'ky': -1, 'khong_tai': -1, 'tk_doat': -1, 'nhuoc_ta_ky': -1, 'thuc_qua_khong_tai': -2}
def diem(f, W): return sum(W.get(k, 0) for k, v in f.items() if v)
def muc(s, a, b): return 'T' if s >= a else ('K' if s <= b else 'V')
def danh_gia(X, fn, W, luoi_a, luoi_b):
    for r in X: r['f'] = fn(r); r['s'] = diem(r['f'], W)
    tong = collections.Counter(r['sach'] for r in X)
    def kq(rs, a, b): return sum(muc(r['s'], a, b) == r['sach'] for r in rs)
    best = max(((kq(X, a, b), a, b) for a in luoi_a for b in luoi_b if b < a))
    # kiểm chéo bỏ-một: chọn ngưỡng trên phần còn lại
    loo = 0
    for i, r in enumerate(X):
        rest = X[:i] + X[i+1:]
        _, a, b = max(((kq(rest, a, b), a, b) for a in luoi_a for b in luoi_b if b < a))
        loo += muc(r['s'], a, b) == r['sach']
    c = collections.Counter((r['sach'], muc(r['s'], best[1], best[2])) for r in X)
    return dict(khop=best[0], n=len(X), a=best[1], b=best[2], loo=loo, theo_nhan={s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'})
if __name__ == '__main__':
    from gd_b_dac_trung import chay
    X = chay(lv_list=('TL',))
    print('TL:', danh_gia(X, tl_dau_hieu, TL_W, range(1, 6), range(-3, 2)))
    st = collections.defaultdict(collections.Counter); tong = collections.Counter(r['sach'] for r in X)
    for r in X:
        for k, v in r['f'].items():
            if v: st[k][r['sach']] += 1
    for k, c in st.items(): print('  %-20s K %2d/%d  V %2d/%d  T %2d/%d' % (k, c['K'], tong['K'], c['V'], tong['V'], c['T'], tong['T']))
def pct(r):
    L = {k: (v or 0) for k, v in r['role']['luong2'].items()}; t = sum(L.values()) or 1
    return {k: 100 * v / t for k, v in L.items()}
def tl_v3(r, p):
    """p = (ta_thap, ta_cao, than_min, bonus_thau_tang)"""
    P = pct(r); ta = P['Tài']; than = P['Tỉ kiếp'] + P['Ấn']; f = tl_dau_hieu(r)
    s = 0
    if ta >= p[1] and than >= p[2]: s += 2
    elif ta >= p[1]: s += 1 if than >= p[2] - 15 else -1      # tài nhiều mà thân yếu: không gánh nổi
    if ta <= p[0]: s -= 2
    if f['thau_tang'] or f['kho']: s += p[3]
    if f['lo_khong_goc']: s -= 2
    return s
def chon_tl3(X):
    import itertools
    luoi = list(itertools.product((2, 4, 6), (7, 9, 11, 13), (30, 40, 50), (0, 1)))
    def kq(rs, p, a, b): return sum(muc(tl_v3(r, p), a, b) == r['sach'] for r in rs)
    def tot_nhat(rs): return max(((kq(rs, p, a, b), p, a, b) for p in luoi for a in (1, 2, 3) for b in (-2, -1, 0) if b < a), key=lambda x: x[0])
    best = tot_nhat(X)
    loo = 0
    for i, r in enumerate(X):
        _, p, a, b = tot_nhat(X[:i] + X[i+1:]); loo += muc(tl_v3(r, p), a, b) == r['sach']
    tong = collections.Counter(r['sach'] for r in X)
    c = collections.Counter((r['sach'], muc(tl_v3(r, best[1]), best[2], best[3])) for r in X)
    return best, loo, {s: '%d/%d' % (c[(s, s)], tong[s]) for s in 'KVT'}
