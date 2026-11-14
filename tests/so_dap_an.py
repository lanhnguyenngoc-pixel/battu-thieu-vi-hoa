# -*- coding: utf-8 -*-
"""T1 – so đáp án TỪNG Ô với 05_KIEM_THU/dap_an_o/*.json.
Tiêu chí: số – sai lệch tương đối ≤ 1e-9; chữ – trùng tuyệt đối; khác kiểu = lệch.

Dữ liệu nhập được đặt ĐÚNG như oracle_dump.py đã làm: duyệt các lá theo thứ tự trong cases_dap_an_o.json,
trạng thái ô nhập CỘNG DỒN từ lá trước (oracle không xóa C7:F8 khi chuyển sang lá ngày giờ, v.v.).
Ô công thức thuộc giai đoạn sau (chưa cài) được "mượn" giá trị từ đáp án của chính lá đó.

Dùng: python -m tests.so_dap_an <thư_mục_05_KIEM_THU> [--ra bao_cao.json]
"""
import sys, os, json, re, argparse, collections, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine, cell_modules
from battu_core.du_lieu import model, ref_of


def oracle_nhap(cases):
    """Sinh lần lượt dict ô nhập NHAP_LIEU như oracle_dump.py (trạng thái cộng dồn)."""
    st = {}
    for c in cases:
        if 'can' in c:
            for i, col in enumerate('CDEF'):
                st[col + '7'] = c['can'][i]; st[col + '8'] = c['chi'][i]
            st['C12'] = 'Không'; st['C13'] = 'Tự nhập'
            for col in 'CDEFG': st[col + '26'] = None
        else:
            st['C12'] = 'Tự động'; st['C13'] = 'Tự nhập'; st['C9'] = 'Biết giờ'
            for col, k in zip('CDEFG', ('ngay', 'thang', 'nam', 'gio', 'phut')): st[col + '26'] = float(c[k])
            st['C27'] = float(c.get('mui_gio', 7))
        st['C10'] = c.get('gioi_tinh', 'Nam'); st['C11'] = None
        yield c, dict(st)


def khop(a, b):
    if isinstance(b, float) and isinstance(a, float):
        if a == b: return True
        return abs(a - b) <= 1e-9 * max(abs(a), abs(b))
    return type(a) is type(b) and a == b


def ten_tep(cid): return re.sub(r'[^\w\-]+', '_', cid) + '.json'


def chay(kt, chi_sheet=None):
    cases = json.load(open(os.path.join(kt, 'cases_dap_an_o.json'), encoding='utf-8'))
    M = model(); F = cell_modules()
    tong = collections.Counter(); dung = collections.Counter(); lech = []
    t0 = time.time()
    for c, nhap in oracle_nhap(cases):
        da = json.load(open(os.path.join(kt, 'dap_an_o', ten_tep(c['id'])), encoding='utf-8'))
        muon = {sh: {k: v for k, v in d.items()} for sh, d in da.items()}
        E = Engine({'NHAP_LIEU': nhap}, muon)
        for si in sorted(F):
            sh = M.sheets[si]
            if chi_sheet and sh not in chi_sheet: continue
            kq = E.tinh_sheet(sh)
            for ref, gt in kq.items():
                if ref not in da.get(sh, {}): continue  # ô không có đáp án (vùng Bảng dữ liệu)
                tong[sh] += 1
                if khop(gt, da[sh][ref]): dung[sh] += 1
                else: lech.append((c['id'], sh, ref, gt, da[sh][ref]))
    return tong, dung, lech, time.time() - t0, len(cases)


if __name__ == '__main__':
    ap = argparse.ArgumentParser(); ap.add_argument('kt'); ap.add_argument('--ra'); ap.add_argument('--sheet', nargs='*')
    a = ap.parse_args()
    tong, dung, lech, dt, n = chay(a.kt, a.sheet)
    print('T1 – %d lá, %.1f giây' % (n, dt))
    for sh in tong: print('  %-16s %6d / %6d  %s' % (sh, dung[sh], tong[sh], 'ĐẠT' if dung[sh] == tong[sh] else 'LỆCH'))
    print('  TỔNG             %6d / %6d' % (sum(dung.values()), sum(tong.values())))
    for x in lech[:40]: print('  LỆCH', x[0], '%s!%s' % (x[1], x[2]), 'pm=%r' % (str(x[3])[:120],), 'da=%r' % (str(x[4])[:120],))
    if a.ra:
        json.dump({'tong': tong, 'dung': dung, 'lech': [list(map(str, x)) for x in lech]}, open(a.ra, 'w'), ensure_ascii=False, indent=1)
