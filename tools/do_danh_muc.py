# -*- coding: utf-8 -*-
"""Đo bộ công cụ trên danh mục lá số chuẩn gốc (tests/danh_muc_la_so_chuan.json, hạng A + B) và nhãn đại vận của sách.
Chạy:  python tools/do_danh_muc.py chay <ra.json>      # tính lại mọi lá bằng lõi hiện tại
       python tools/do_danh_muc.py so <truoc.json> <sau.json>   # so hai lần chạy
Chỉ so: (1) vượng/nhược cuối; (2) dụng thần (hành) với lá có nhãn; (3) đại vận tốt/xấu với lá có nhãn (nhan_dai_van_sach.py + nhan_dai_van_mtc.py)."""
import sys, os, json, multiprocessing as mp
GOC = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
sys.path.insert(0, GOC); sys.path.insert(0, os.path.join(GOC, 'tai_lieu', 'giai_doan_A'))
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


def _gt(E, sh, ref):
    try: return E.gia_tri(sh, ref)
    except Exception: return None


def _tinh(c):
    from battu_core.engine import Engine
    from battu_core.dich_vu import KetQua
    from battu_core.nhap_lieu import LaSo
    from tests.so_dap_an import oracle_nhap
    from battu_core.tong_hop.viet_ban import ban_luan_giai
    (cc, nhap), = list(oracle_nhap([{'id': c['id'], 'can': c['can'], 'chi': c['chi'], 'gioi_tinh': c['gioi']}]))
    try:
        E = Engine({'NHAP_LIEU': nhap}, None); E.tinh_tat_ca()
        kq = KetQua(LaSo(ho_ten=c['id'], cach_nhap='canchi'), E, 0)
        B, hs, pd = ban_luan_giai(kq); bm = hs['ban_menh']
        return c['id'], {'vs': str(bm.get('vuong_suy'))[:6], 'pp1': bm.get('pp1'), 'pct': bm.get('phe_minh_pt'), 'dung': bm.get('hanh_dung'),
                         'cach': bm.get('cach_cuc'), 'b2': _gt(E, 'PP1_DINH_TINH', 'C136'), 'bh': _gt(E, 'PP1_DINH_TINH', 'C125'), 'dv': [(d['ten'], d.get('ket_luan')) for d in hs['dai_van']],
                         'dvm': [(d['ten'], d.get('muc6') or d.get('muc')) for d in pd['dai_van']]}
    except Exception as e:
        return c['id'], {'loi': repr(e)[:200]}


def cac_la():
    L = []
    for r in json.load(open(os.path.join(GOC, 'tests/danh_muc_la_so_chuan.json'), encoding='utf-8')):
        if r['hang'] in ('A', 'B'):
            L.append({'id': r['ma'], 'can': r['can'], 'chi': r['chi'], 'gioi': 'Nam'})
    ds = {d['ten']: d for d in json.load(open(os.path.join(GOC, 'tai_lieu/giai_doan_A/dai_van_doi_chieu.json'), encoding='utf-8'))}
    from nhan_dai_van_sach import NHAN
    for ten in NHAN:
        d = ds.get(ten)
        if d:
            tt = [x.strip().split() for x in d['tt'].split(',')]
            L.append({'id': 'DV|' + ten, 'can': [x[0] for x in tt], 'chi': [x[1] for x in tt], 'gioi': d['g']})
    from nhan_dai_van_mtc import MTC
    for k, v in MTC.items():
        L.append({'id': 'DV|' + k, 'can': v['can'], 'chi': v['chi'], 'gioi': v['gioi']})
    return L


def chay(ra):
    L = cac_la()
    if os.environ.get('CHI_VAN'): L = [c for c in L if c['id'].startswith('DV|')]
    with mp.Pool(2) as P: R = dict(P.map(_tinh, L, chunksize=4))
    json.dump(R, open(ra, 'w'), ensure_ascii=False)
    print('xong', len(R), 'lỗi', sum('loi' in v for v in R.values()))


def _h(x):
    x = (x or '').strip().upper()
    return 'T' if x in ('TỐT', 'RẤT TỐT') else ('X' if x in ('XẤU', 'RẤT XẤU') else 'B')


def so(f):
    R = json.load(open(f, encoding='utf-8'))
    D = {r['ma']: r for r in json.load(open(os.path.join(GOC, 'tests/danh_muc_la_so_chuan.json'), encoding='utf-8'))}
    out = {}
    for h in 'AB':
        S = [(m, r) for m, r in D.items() if r['hang'] == h and m in R and 'loi' not in R[m]]
        out['vs' + h] = (sum(1 for m, r in S if R[m]['vs'][:3] == {'V': 'VƯỢ', 'N': 'NHƯ'}[r['nhan_vuong_suy']]), len(S))
        S2 = [(m, r) for m, r in S if r['nhan_dung_than']]
        out['dt' + h] = (sum(1 for m, r in S2 if R[m]['dung'] in r['nhan_dung_than']), len(S2))
    from nhan_dai_van_sach import NHAN
    from nhan_dai_van_mtc import MTC
    n = k = ng = 0; chi_tiet = []
    for ten, nh in NHAN.items():
        v = R.get('DV|' + ten)
        if not v or 'loi' in v: continue
        for vt, sach in nh.items():
            if vt > len(v['dv']): continue
            a = _h(v['dv'][vt - 1][1]); n += 1; k += a == sach; ng += {a, sach} == {'T', 'X'}
            chi_tiet.append((ten, vt, v['dv'][vt - 1][0], sach, a))
    out['dvTVH'] = (k, n, ng)
    n = k = ng = 0
    for ten, m in MTC.items():
        v = R.get('DV|' + ten)
        if not v or 'loi' in v: continue
        tenvan = dict(v['dv'])
        for tv, sach in m['van'].items():
            if tv not in tenvan: continue
            a = _h(tenvan[tv]); n += 1; k += a == sach; ng += {a, sach} == {'T', 'X'}
            chi_tiet.append((ten, 0, tv, sach, a))
    out['dvMTC'] = (k, n, ng)
    return out, chi_tiet


if __name__ == '__main__':
    if sys.argv[1] == 'chay': chay(sys.argv[2])
    else:
        a, ca = so(sys.argv[2])
        if len(sys.argv) > 3:
            b, cb = so(sys.argv[3])
            for k in a: print(k, a[k], '→', b[k])
            da = {(t, v): x for t, vt, v, s, x in ca}; db = {(t, v): (s, x) for t, vt, v, s, x in cb}
            for key, (s, x) in db.items():
                if key in da and da[key] != x: print('vận đổi', key, 'sách', s, ':', da[key], '→', x)
        else:
            for k in a: print(k, a[k])
