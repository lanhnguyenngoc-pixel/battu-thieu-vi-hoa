# -*- coding: utf-8 -*-
"""T2 – kết quả chính 501 lá (05_KIEM_THU/bo_la, đáp án ket_qua_6.3 do scripts/uno_run59.py lập).
(a) Các khóa của uno_run59.py thuộc sheet đã cài, đọc như uno_run59 (getValue nếu ô hằng số, ngược lại getString).
    Ô nhập như uno_run59: C7:F8, C13 = Tự nhập, C12 = Không, C10 = Nam, C11 trống (mỗi tệp lá chạy từ workbook mặc định).
    Khóa phụ thuộc ô của giai đoạn sau (d171, d175: DAI_VAN; d182: LUU_NIEN) dùng ô "mượn" từ đáp án từng ô T2_501.
(b) T1-501: so TỪNG Ô các sheet đã cài trên 501 lá với tests/dap_an_mo_rong/T2_501 (tools/oracle_mo_rong.py, biến thể
    T2_501 – đặt như uno_run59, xóa Bảng dữ liệu KHONG_RO_GIO).
Dùng: python -m tests.t2_501 <thư_mục_05_KIEM_THU> <tests/dap_an_mo_rong>"""
import sys, os, json, time, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine, cell_modules
from battu_core.du_lieu import model
from battu_core.hien_thi import doc_nhu_uno
from battu_core import sheets as S
from tests.so_dap_an import khop

READ = {'p1': ('PP1_DINH_TINH', 'C91'), 'p2': ('PP2_DINH_LUONG', 'C103'), 'dung1': ('DUNG_HY_KY', 'C106'),
        'dung2': ('DUNG_HY_KY', 'D106'), 'kn': ('DOI_CHIEU', 'C34'), 'tin': ('DOI_CHIEU', 'C35'), 'kn_d': ('DOI_CHIEU', 'D33'),
        'nht1': ('NGU_HANH_THIEU', 'E15'), 'nht_dl': ('NGU_HANH_THIEU', 'E7'), 'nht_c7': ('NGU_HANH_THIEU', 'C7')}
for r in (163, 170, 171, 174, 175, 178, 179, 180, 181, 182, 183, 184): READ['d%d' % r] = ('DUNG_HY_KY', 'C%d' % r)
for r in range(4, 13): READ['k%d' % r] = ('KET_LUAN', 'C%d' % r)
for a in ('J6', 'J7', 'J8'): READ['sk' + a] = ('SUC_KHOE', a)
READ.update(dk=('LUU_NIEN', 'DK31'), dl=('LUU_NIEN', 'DL31'), dn=('DOC_NHANH', 'C14'), m16=('TO_CHAT', 'M16'), c16=('TO_CHAT', 'C16'),
            ai64=('HON_NHAN', 'AI64'), ag64=('HON_NHAN', 'AG64'), x64=('HON_NHAN', 'X64'), r64=('HON_NHAN', 'R64'),
            s64=('HON_NHAN', 'S64'), c206=('KET_LUAN', 'C206'), c72=('KET_LUAN', 'C72'), c179=('KET_LUAN', 'C179'),
            f115=('LUAN_GIAI', 'F115'))
NGUON = [('G1', 'cases_G1.json'), ('VN', 'cases_books_VN.json'), ('TR', 'cases_trang49.json')]


def da_cai(sheet, ref):
    pv = S.PHAM_VI.get(sheet)
    if pv is None: return False
    if pv == 'all': return True
    from openpyxl.utils.cell import range_boundaries
    from battu_core.du_lieu import rc
    r, c = rc(ref)
    return any(c1 <= c <= c2 and r1 <= r <= r2 for c1, r1, c2, r2 in (range_boundaries(x) for x in pv))


def chay(kt, mr):
    global KQ
    KQ = next((k for k in ('ket_qua_6.10', 'ket_qua_6.9', 'ket_qua_6.8', 'ket_qua_6.7', 'ket_qua_6.6', 'ket_qua_6.5', 'ket_qua_6.4', 'ket_qua_6.3') if os.path.isdir(os.path.join(kt, k))), 'ket_qua_6.3')
    M = model(); F = cell_modules()
    khoa = {k: v for k, v in READ.items() if da_cai(*v)}
    tk = collections.Counter(); dk = collections.Counter(); lk = []
    to = collections.Counter(); do = collections.Counter(); lo = []; n = 0; t0 = time.time(); thieu = 0
    for tag, f in NGUON:
        cases = json.load(open(os.path.join(kt, 'bo_la', f), encoding='utf-8'))
        res = json.load(open(os.path.join(kt, KQ, 'res_' + f), encoding='utf-8'))
        for i, (c, r) in enumerate(zip(cases, res)):
            assert c['id'] == r['id']
            nhap = {'C13': 'Tự nhập', 'C12': 'Không', 'C10': 'Nam', 'C11': None}
            for j, col in enumerate('CDEF'): nhap[col + '7'] = c['can'][j]; nhap[col + '8'] = c['chi'][j]
            da = None
            for thu_muc in ('T2_501', 'T2_501_GD3', 'T2_501_GD4', 'T2_501_GD5'):     # các lượt oracle theo giai đoạn
                p = os.path.join(mr, thu_muc, '%s%03d.json' % (tag, i))
                if os.path.exists(p):
                    d_ = json.load(open(p, encoding='utf-8'))
                    if da is None: da = d_
                    else:
                        for sh_, v_ in d_.items(): da.setdefault(sh_, {}).update(v_)
            if da is None: thieu += 1
            E = Engine({'NHAP_LIEU': nhap}, da)
            n += 1
            for k, (sh, ref) in khoa.items():
                tk[k] += 1; v = doc_nhu_uno(E, sh, ref)
                if v == r[k]: dk[k] += 1
                else: lk.append((c['id'], k, v, r[k]))
            if da is not None:
                for si in sorted(F):
                    sh = M.sheets[si]
                    for ref, gt in E.tinh_sheet(sh).items():
                        if ref not in da.get(sh, {}): continue
                        to[sh] += 1
                        if khop(gt, da[sh][ref]): do[sh] += 1
                        else: lo.append((c['id'], sh, ref, gt, da[sh][ref]))
    return n, khoa, tk, dk, lk, to, do, lo, thieu, time.time() - t0


if __name__ == '__main__':
    n, khoa, tk, dk, lk, to, do, lo, thieu, dt = chay(sys.argv[1], sys.argv[2])
    print('T2 – %d lá (%.1f giây); thiếu đáp án từng ô: %d lá' % (n, dt, thieu))
    print('(a) khóa uno_run59 đã cài:', ' '.join('%s %d/%d' % (k, dk[k], tk[k]) for k in tk))
    for x in lk[:15]: print('   LỆCH', x[0], x[1], repr(x[2])[:90], '| đáp án', repr(x[3])[:90])
    print('(b) T1-501:', ' '.join('%s %d/%d' % (s, do[s], to[s]) for s in to), '| tổng %d/%d' % (sum(do.values()), sum(to.values())))
    for x in lo[:15]: print('   LỆCH', x[0], '%s!%s' % (x[1], x[2]), repr(x[3])[:80], '| đáp án', repr(x[4])[:80])
