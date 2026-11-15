import sys, collections
sys.path.insert(0, '/home/claude/battu/tools'); sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from gd_b_dac_trung import chay
from tt_co_ban import *
X = chay(lv_list=('HN',))
XUNG = {frozenset(p) for p in [('Tí','Ngọ'),('Sửu','Mùi'),('Dần','Thân'),('Mão','Dậu'),('Thìn','Tuất'),('Tị','Hợi')]}
HINH = {frozenset(p) for p in [('Dần','Tị'),('Tị','Thân'),('Dần','Thân'),('Sửu','Tuất'),('Tuất','Mùi'),('Sửu','Mùi'),('Tí','Mão')]}
AD_LECH = {'Bính Tí','Đinh Sửu','Mậu Dần','Tân Mão','Nhâm Thìn','Quý Tị','Bính Ngọ','Đinh Mùi','Mậu Thân','Tân Dậu','Nhâm Tuất','Quý Hợi'}
KD = {'Giáp':'Mão','Bính':'Ngọ','Mậu':'Ngọ','Canh':'Dậu','Nhâm':'Tí'}
def dh(r):
    a = phan_tich(r['tu_tru']); lo, bk, tg, chi = a['lo'], a['bk'], a['tang'], a['chi']; nu = r['gioi'] == 'Nữ'
    ro = lo + bk; tat = ro + tg; d = {}
    d['TQ_lo'] = lo.count('Thương') >= 1
    d['TQ_nhieu'] = tat.count('Thương') >= 3 or (lo.count('Thương') >= 1 and bk.count('Thương') >= 1)
    d['TK_trung'] = sum(x in ('Tỉ','Kiếp') for x in ro) >= 3
    d['KD'] = a['nc'] in KD and KD[a['nc']] in chi
    d['NT_tu'] = HANH_CAN[CAN.index(a['nc'])] == HANH_CAN[CAN.index(TANG[chi[2]][0])] or a['can'][2] + ' ' + chi[2] in AD_LECH or (a['nc'] == 'Giáp' and chi[2] in ('Thìn','Tuất'))
    d['CUNG_XH'] = any(frozenset((chi[2], chi[i])) in XUNG | HINH for i in (1, 3)) or chi[2] in (chi[1], chi[3]) and chi[2] in ('Thìn','Ngọ','Dậu','Hợi')
    if nu:
        d['QS_hon'] = 'Quan' in ro and 'Sát' in ro
        d['QS_khong'] = not any(x in ('Quan','Sát') for x in tat)
        d['KA_trung'] = sum(x in ('Kiêu','Ấn') for x in ro) >= 3
        d['TTSM'] = sum(c in ('Thìn','Tuất','Sửu','Mùi') for c in chi) >= 3
        tot = all(any(NHOM[x] == g for x in ro) for g in ('TA','QS','AN'))
    else:
        d['TA_2'] = 'CTài' in lo and 'TTài' in lo or (('CTài' in ro) and ('TTài' in ro) and lo.count('CTài') + lo.count('TTài') >= 1)
        d['TA_khong'] = not any(x in ('CTài','TTài') for x in tat)
        tot = bk[2] in ('CTài','TTài') and not d['CUNG_XH']
    tot = tot or r['qt'].get('HB22', (False,))[0] or r['qt'].get('HB23', (False,))[0]
    return d, tot
for r in X: r['d'], r['tot'] = dh(r); r['b'] = sum(r['d'].values())
st = collections.defaultdict(collections.Counter)
for r in X:
    for k, v in r['d'].items():
        if v: st[k][r['sach']] += 1
print('Dấu hiệu      K(44)  V(7)  T(5)'); [print('%-12s %4d %5d %5d' % (k, c['K'], c['V'], c['T'])) for k, c in sorted(st.items(), key=lambda x: -x[1]['K'])]
for s in 'KVT': print(s, 'số dấu hiệu:', sorted(r['b'] for r in X if r['sach'] == s), 'có dấu hiệu tốt:', sum(r['tot'] for r in X if r['sach'] == s))
def dd(r, a):
    b = r['b'] - (1 if r['tot'] else 0)
    return 'K' if b >= a else ('V' if b >= a - 1 else 'T')
for a in (1, 2, 3, 4):
    c = collections.Counter((r['sach'], dd(r, a)) for r in X)
    print('ngưỡng', a, 'khớp', sum(v for (s, p), v in c.items() if s == p), '/', len(X), dict(c))
import pickle
Q = pickle.load(open('/home/claude/battu/tai_lieu/giai_doan_A/quan_the.pkl', 'rb'))
for q in Q: q['qt'] = q['qt']['HN']; q['d'], q['tot'] = dh(q); q['b'] = sum(q['d'].values())
print('\nQuần thể %d lá ngẫu nhiên:' % len(Q))
print(' công cụ hiện tại:', collections.Counter(q['o']['HN'] for q in Q))
for a in (1, 2, 3): print(' đề xuất ngưỡng %d:' % a, collections.Counter(dd(q, a) for q in Q))
st2 = collections.Counter(k for q in Q for k, v in q['d'].items() if v)
print(' tần suất dấu hiệu trong quần thể:', {k: '%d%%' % (100 * v // len(Q)) for k, v in st2.most_common()})
W = {'NT_tu': .5, 'CUNG_XH': .5}
def dd2(r, a=2):
    b = sum(W.get(k, 1) for k, v in r['d'].items() if v) - (1 if r['tot'] else 0)
    return 'K' if b >= a else ('V' if b >= a - 1 else 'T')
print('\n--- Trọng số: dấu hiệu nhật trụ và cung xung/hình tính nửa ---')
for a in (1.5, 2):
    c = collections.Counter((r['sach'], dd2(r, a)) for r in X)
    k = {s: sum(v for (x, p), v in c.items() if x == s and p == s) for s in 'KVT'}
    print('ngưỡng', a, 'sách K đúng %d/44, V %d/7, T %d/5' % (k['K'], k['V'], k['T']), '| quần thể:', collections.Counter(dd2(q, a) for q in Q))
