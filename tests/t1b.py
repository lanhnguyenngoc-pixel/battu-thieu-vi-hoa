# -*- coding: utf-8 -*-
"""T1b – đáp án từng ô cho các NHÁNH tùy chọn (MT-05) và T1-KRG (MT-04), lấy bằng tools/oracle_mo_rong.py.
Ô nhập = trình tự chuẩn của oracle (tests/so_dap_an.oracle_nhap) + phần đặt thêm của biến thể (tests/bien_the.json).
Chỉ so các sheet đã cài ở giai đoạn hiện tại; ô của giai đoạn sau được mượn từ đáp án của chính biến thể.
Dùng: python -m tests.t1b <thư_mục_05_KIEM_THU> <tests/dap_an_mo_rong> [biến_thể ...]"""
import sys, os, json, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine, cell_modules
from battu_core.du_lieu import model
from tests.so_dap_an import oracle_nhap, khop, ten_tep

BT = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bien_the.json'), encoding='utf-8'))


def chay_bien_the(kt, thu_muc, ten, cases_path=None):
    cases = json.load(open(cases_path or os.path.join(kt, 'cases_dap_an_o.json'), encoding='utf-8'))
    M = model(); F = cell_modules(); b = BT[ten]
    tong = collections.Counter(); dung = collections.Counter(); lech = []; n = 0
    for c, nl in oracle_nhap(cases):
        p = os.path.join(thu_muc, ten, ten_tep(c['id']))
        if not os.path.exists(p): continue
        n += 1
        da = json.load(open(p, encoding='utf-8'))
        nhap = {'NHAP_LIEU': dict(nl)}
        for part in (b.get('chung', {}), b.get('rieng', {}).get(c['id'], {})):
            for sh, d in part.items(): nhap.setdefault(sh, {}).update(d)
        E = Engine(nhap, da)
        for si in sorted(F):
            sh = M.sheets[si]
            for ref, gt in E.tinh_sheet(sh).items():
                if ref not in da.get(sh, {}): continue
                tong[sh] += 1
                if khop(gt, da[sh][ref]): dung[sh] += 1
                else: lech.append((c['id'], sh, ref, gt, da[sh][ref]))
    return n, tong, dung, lech



def so_krg0_voi_dap_an_o(kt, thu_muc):
    """Kiểm chính oracle mở rộng: KRG0 phải trùng dap_an_o ở mọi ô chung."""
    cases = json.load(open(os.path.join(kt, 'cases_dap_an_o.json'), encoding='utf-8'))
    n = d = 0
    for c in cases:
        a = json.load(open(os.path.join(kt, 'dap_an_o', ten_tep(c['id'])), encoding='utf-8'))
        p = os.path.join(thu_muc, 'KRG0', ten_tep(c['id']))
        if not os.path.exists(p): continue
        b = json.load(open(p, encoding='utf-8'))
        for sh in a:
            for k, v in a[sh].items():
                n += 1; d += khop(b[sh].get(k), v)
    return d, n


if __name__ == '__main__':
    kt, thu_muc = sys.argv[1], sys.argv[2]
    ten_list = sys.argv[3:] or list(BT)
    d, n = so_krg0_voi_dap_an_o(kt, thu_muc)
    print('Kiểm oracle mở rộng: KRG0 trùng dap_an_o %d/%d ô' % (d, n))
    for ten in ten_list:
        n, tong, dung, lech = chay_bien_the(kt, thu_muc, ten)
        print('%-16s %2d lá | %s | tổng %d/%d' % (ten, n, ' '.join('%s %d/%d' % (s, dung[s], tong[s]) for s in tong),
                                               sum(dung.values()), sum(tong.values())))
        for x in lech[:10]: print('   LỆCH', x[0], '%s!%s' % (x[1], x[2]), repr(x[3])[:100], '| đáp án', repr(x[4])[:100])
