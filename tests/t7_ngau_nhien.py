# -*- coding: utf-8 -*-
"""T7 – 20 lá NGẪU NHIÊN MỚI (tests/t7_ngau_nhien.json, hạt giống 20261114: 14 lá dương lịch 1900–2100, 6 lá tự nhập
can chi), đáp án lấy bằng 05_KIEM_THU/scripts/oracle_dump.py nguyên bản (tests/dap_an_mo_rong/T7_NGAU_NHIEN).
So MỌI ô công thức của 33 sheet; KHÔNG mượn ô nào từ đáp án. Trạng thái nhập cộng dồn như oracle (so_dap_an.oracle_nhap).
Dùng: python -m tests.t7_ngau_nhien"""
import sys, os, json, time, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine, cell_modules
from battu_core.du_lieu import model
from tests.so_dap_an import oracle_nhap, khop, ten_tep

GOC = os.path.dirname(os.path.abspath(__file__))


def chay(cases_path=os.path.join(GOC, 't7_ngau_nhien.json'), da_dir=os.path.join(GOC, 'dap_an_mo_rong', 'T7_NGAU_NHIEN')):
    cases = json.load(open(cases_path, encoding='utf-8'))
    M = model(); F = cell_modules(); tong = collections.Counter(); dung = collections.Counter(); lech = []; t0 = time.time(); thieu = []
    for c, nhap in oracle_nhap(cases):
        p = os.path.join(da_dir, ten_tep(c['id']))
        if not os.path.exists(p): thieu.append(c['id']); continue
        da = json.load(open(p, encoding='utf-8'))
        E = Engine({'NHAP_LIEU': nhap}, None)
        for si in sorted(F):
            sh = M.sheets[si]
            for ref, gt in E.tinh_sheet(sh).items():
                if ref not in da.get(sh, {}): continue
                tong[sh] += 1
                if khop(gt, da[sh][ref]): dung[sh] += 1
                else: lech.append((c['id'], sh, ref, gt, da[sh][ref]))
        if E.chua_cai: raise SystemExit('Còn ô chưa cài: %d' % len(E.chua_cai))
    return len(cases) - len(thieu), tong, dung, lech, thieu, time.time() - t0


if __name__ == '__main__':
    n, tong, dung, lech, thieu, dt = chay()
    print('T7 – %d lá ngẫu nhiên, %.1f giây; thiếu đáp án: %s' % (n, dt, thieu))
    for sh in tong: print('  %-16s %6d / %6d  %s' % (sh, dung[sh], tong[sh], 'ĐẠT' if dung[sh] == tong[sh] else 'LỆCH'))
    print('  TỔNG             %6d / %6d' % (sum(dung.values()), sum(tong.values())))
    for x in lech[:40]: print('  LỆCH', x[0], '%s!%s' % (x[1], x[2]), repr(x[3])[:100], '| đáp án', repr(x[4])[:100])
    sys.exit(1 if (lech or thieu or n == 0) else 0)
