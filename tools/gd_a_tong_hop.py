"""Giai đoạn A: gộp bộ đáp án 3 phần sách TVH thành bảng tổng hợp + JSON."""
import json, sys, re, collections
sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from dap_an_p1 import DA1
from dap_an_p2a import DA
from dap_an_p3 import DA3
LV = [('TC','Tính cách – tổng thể'),('SN','Sự nghiệp – quan vận'),('TL','Tài lộc'),('HN','Hôn nhân'),
      ('CC','Con cái'),('LT','Cha mẹ – anh em'),('SK','Sức khỏe (tham khảo)')]
# gộp P3 trùng P2
gop = {}
for k, v in list(DA1.items()) + list(DA.items()):
    gop[k] = dict(v, nguon=['P1' if k.startswith('P1') else 'P2'])
for k, v in DA3.items():
    p2 = v.get('trung_p2')
    if p2 and p2 in gop:
        g = gop[p2]
        g['nguon'].append('Ch.23 tr.' + v['trang'])
        for c, _ in LV + [('DV',''), ('HT','')]:
            if c in v:
                if c in g and c not in ('DV','HT') and g[c][0] != v[c][0]:
                    g.setdefault('khac_nhau', []).append((c, g[c][0], v[c][0]))
                if c in g and c not in ('DV','HT'):
                    g[c + '_P2'] = g[c]
                g[c] = v[c]
        if v.get('ghi_chu'): g['ghi_chu'] = v['ghi_chu']
        g['trang'] = g['trang'] + '; Ch.23 tr.' + v['trang']
    else:
        gop[k] = dict(v, nguon=['Ch.23'])
def khoa(k):
    m = re.match(r'P(\d)', k); p = int(m.group(1)) if m else 3
    n = re.search(r'VD(\d+)', k); tr = re.search(r'C23-(\d+)', k)
    return (p, int(n.group(1)) if n else int(tr.group(1)) if tr else 0)
ks = sorted(gop, key=khoa)
json.dump({k: gop[k] for k in ks}, open('/home/claude/battu/tai_lieu/giai_doan_A/dap_an_tong.json', 'w'), ensure_ascii=False, indent=1)
dem = collections.Counter(); nhan = collections.defaultdict(collections.Counter)
for k in ks:
    for c, _ in LV:
        if c in gop[k]: dem[c] += 1; nhan[c][gop[k][c][0]] += 1
    if 'DV' in gop[k]: dem['DV'] += 1
print('Số lá:', len(ks))
for c, t in LV: print(f'{t:28s} {dem[c]:3d}  T={nhan[c]["T"]} V={nhan[c]["V"]} K={nhan[c]["K"]}')
print('Giai đoạn đại vận', dem['DV'])
for k in ks:
    if gop[k].get('khac_nhau'): print('KHÁC P2/P3', k, gop[k]['khac_nhau'])
