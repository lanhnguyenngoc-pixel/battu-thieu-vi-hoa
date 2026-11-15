import sys, json, re, collections, itertools
sys.path.insert(0, '/home/claude/battu/tools'); sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from gd_b_dac_trung import chay
X = chay(lv_list=('HN',))
XAU = r'khắc|tái giá|hai lần|lần thứ hai|ly |ly dị|bất hòa|không thuận|mất vợ|mất chồng|tổn hại vợ|tổn thương vợ|không lợi cho vợ|xa lìa|cô đơn|tranh|tuyệt nghĩa|chú rể|cãi|khó giữ|không hoà|không hòa|vợ lẽ|vợ kế|ba vợ|nhiều chồng|tai ương|hình phạt'
TOT = r'hiền|đằm thắm|hòa thuận|hoà thuận|vinh|phú quý|đảm đang|nhờ vợ|nội trợ tốt|được quý|thọ|đoan trang|nghĩa khí|giúp|tháo vát|chồng vinh'
def cuc(g):
    x, t = bool(re.search(XAU, g.lower())), bool(re.search(TOT, g.lower()))
    return 'X' if x and not t else ('T' if t and not x else ('XT' if x and t else '-'))
# thống kê từng quy tắc
st = collections.defaultdict(lambda: collections.Counter())
for r in X:
    for ma, (du, dk, kl, row) in r['qt'].items():
        if du: st[ma][r['sach']] += 1
tong = collections.Counter(r['sach'] for r in X)
print('Số lá theo nhãn sách:', dict(tong))
qt0 = X[0]['qt']
print('%-6s %-3s %4s %4s %4s  %s' % ('Mã', 'cực', 'K', 'V', 'T', 'kết luận sách'))
for ma in sorted(st, key=lambda m: -sum(st[m].values())):
    c = st[ma]; kl = qt0[ma][2]
    print('%-6s %-3s %4d %4d %4d  %s' % (ma, cuc(kl), c['K'], c['V'], c['T'], kl[:80]))
print('\n===== THỬ CÁCH KẾT HỢP =====')
NANG = ['HB7','HB8','HB9','HB1','HB4','HB5','HB6','HB11','HB10','HB12','HB13','HB14','HB16','HB17','HB19','HN68','HN51','HN46','HN71','HN76','HB3','HB15','HN54','HN43','HV36','HV46','HN44']
TOTS = ['HV19','HN3','HN15','HN19','HB21','HB22','HB23','HN13','HV12','HV10','HN20','HN25']
def dem(r):
    b = sum(1 for m in NANG if r['qt'].get(m, (False,))[0]); g = sum(1 for m in TOTS if r['qt'].get(m, (False,))[0]); return b, g
import statistics
for r in X: r['b'], r['g'] = dem(r)
for s in 'KVT': print(s, 'xấu:', sorted(r['b'] for r in X if r['sach'] == s), ' tốt:', sorted(r['g'] for r in X if r['sach'] == s))
def du_doan(r, a, c):
    d = r['b'] - c * r['g']
    return 'K' if d >= a else ('V' if d >= a - 2 else 'T')
def chamdiem(rs, a, c):
    rec = {}
    for s in 'KVT':
        xs = [r for r in rs if r['sach'] == s]
        if xs: rec[s] = sum(du_doan(r, a, c) == s for r in xs) / len(xs)
    return sum(rec.values()) / len(rec), rec
best = max(((chamdiem(X, a, c)[0], a, c) for a in range(1, 9) for c in (0, 0.5, 1, 1.5, 2)))
print('Tốt nhất (cân bằng 3 nhãn):', best, chamdiem(X, best[1], best[2])[1])
# kiểm chéo bỏ-một
ok = 0
for i, r in enumerate(X):
    rest = X[:i] + X[i+1:]
    _, a, c = max(((chamdiem(rest, a, c)[0], a, c) for a in range(1, 9) for c in (0, 0.5, 1, 1.5, 2)))
    ok += du_doan(r, a, c) == r['sach']
print('Kiểm chéo bỏ-một: khớp %d/%d' % (ok, len(X)))
