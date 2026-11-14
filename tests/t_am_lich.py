# -*- coding: utf-8 -*-
"""Kiểm đổi âm lịch: (1) s2l của battu_core trùng lunar.py gốc ở mọi ngày 1899–2101;
(2) khứ hồi l2s(s2l(d)) = d ở mọi ngày; (3) 39 lá data39.json (dl ↔ al).
Dùng: python -m tests.t_am_lich <09_TAI_LIEU_BO_SUNG/ngu_hanh_thieu>"""
import sys, os, json, datetime, importlib.util, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core import lich_am as L

def chay(thu_muc):
    spec = importlib.util.spec_from_file_location('lunar_goc', os.path.join(thu_muc, 'lunar.py'))
    G = importlib.util.module_from_spec(spec); spec.loader.exec_module(G)
    t0 = time.time(); d = datetime.date(1899, 1, 1); n = a = b = 0; loi = []
    while d <= datetime.date(2101, 12, 31):
        n += 1
        x = L.s2l(d.day, d.month, d.year)
        if x == G.s2l(d.day, d.month, d.year): a += 1
        else: loi.append(('s2l', d, x))
        if L.l2s(*x) == (d.day, d.month, d.year): b += 1
        else: loi.append(('l2s', d, x, L.l2s(*x)))
        d += datetime.timedelta(days=1)
    data = json.load(open(os.path.join(thu_muc, 'data39.json'), encoding='utf-8'))
    c = e = 0; ghi = []
    for r in data:
        dd, mm, yy = map(int, r['dl'].split('/'))
        nhuan = 1 if 'nhuận' in r['al'] else 0
        ad, am, ay = map(int, r['al'].replace('(nhuận)', '').replace(' ', '').split('/'))
        x = L.s2l(dd, mm, yy)
        if x == (ad, am, ay, nhuan): c += 1
        else: ghi.append(('dl→al', r['dl'], r['al'], x))
        if L.l2s(ad, am, ay, nhuan) == (dd, mm, yy): e += 1
        else: ghi.append(('al→dl', r['al'], r['dl'], L.l2s(ad, am, ay, nhuan)))
    return n, a, b, len(data), c, e, loi, ghi, time.time() - t0

if __name__ == '__main__':
    n, a, b, m, c, e, loi, ghi, dt = chay(sys.argv[1])
    print('Âm lịch: %d ngày 1899–2101 | s2l trùng lunar.py %d/%d | khứ hồi âm→dương %d/%d | data39: dương→âm %d/%d, âm→dương %d/%d (%.0f giây)' % (n, a, n, b, n, c, m, e, m, dt))
    for x in (loi + ghi)[:20]: print('  LỆCH', x)
