# -*- coding: utf-8 -*-
"""T4 – hợp/xung: 510 lá (05_KIEM_THU/hop_xung/all.json gồm 501 lá sách + 9 ca biên edge.json), đáp án res.json
do scripts/run_hop.py lập. Tái hiện ĐÚNG run_hop.py:
  * ô nhập: C7:F8, C13 = Tự nhập, C12 = Không, C11 = trống (cộng dồn từ workbook mặc định, không đặt C10);
  * mỗi giá trị là CHỮ HIỂN THỊ của ô (getString) → dùng battu_core.hien_thi theo định dạng ô;
  * các tổng phụ (pp1tot, sumF, parts, canF) tính từ chữ hiển thị như run_hop.py (hàm N).
Gói hồ sơ không ghi văn bản "6 bất biến"; bộ này so TRÙNG TỪNG GIÁ TRỊ run_hop.py đã ghi (đáp án của workbook
đã đạt 0 vi phạm), nên mọi bất biến tính trên các giá trị đó cũng được giữ nguyên.
Dùng: python -m tests.t4_hop_xung <thư_mục_05_KIEM_THU>"""
import sys, os, json, time, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.hien_thi import doc_nhu_uno as hien_thi_o

RD = {'cuc3': ('TRUNG_GIAN', 52), 'banhop': ('TRUNG_GIAN', 59), 'trongcuc': ('TRUNG_GIAN', 60), 'luchop': ('TRUNG_GIAN', 69),
      'tranh': ('TRUNG_GIAN', 71), 'loai': ('TRUNG_GIAN', 62), 'hoa': ('PP2_DINH_LUONG', 24), 'xung': ('PP2_DINH_LUONG', 31),
      'khubi': ('PP2_DINH_LUONG', 33), 'd_hoa': ('PP2_DINH_LUONG', 76), 'd_ban': ('PP2_DINH_LUONG', 83),
      'd_t2': ('PP2_DINH_LUONG', 85), 'd_t3': ('PP2_DINH_LUONG', 87), 'tonthat': ('PP2_DINH_LUONG', 82),
      'can': ('PP2_DINH_LUONG', 72), 'tg1': ('TRUNG_GIAN', 15), 'tg2': ('TRUNG_GIAN', 17), 'tg3': ('TRUNG_GIAN', 19),
      'pp1': ('PP1_DINH_TINH', 96)}


def N(v):
    try: return float(v)
    except Exception: return 0.0


def chay(kt):
    cases = json.load(open(os.path.join(kt, 'hop_xung', 'all.json'), encoding='utf-8'))
    res = json.load(open(os.path.join(kt, 'hop_xung', 'res.json'), encoding='utf-8'))
    edge = {c['id'] for c in json.load(open(os.path.join(kt, 'hop_xung', 'edge.json'), encoding='utf-8'))}
    st = {}; t0 = time.time(); khop = collections.Counter(); tong = collections.Counter(); lech = []; la_dung = 0
    for c, r in zip(cases, res):
        assert c['id'] == r['id']
        for i, col in enumerate('CDEF'): st[col + '7'] = c['can'][i]; st[col + '8'] = c['chi'][i]
        st['C13'] = 'Tự nhập'; st['C12'] = 'Không'; st['C11'] = None
        E = Engine({'NHAP_LIEU': dict(st)})
        def row(sh, rr, cols='CDEF'): return [hien_thi_o(E, sh, '%s%d' % (x, rr)) for x in cols]
        o = {k: row(*v) for k, v in RD.items()}
        o['pp1tot'] = sum(N(row('PP1_DINH_TINH', x, 'F')[0]) for x in range(96, 101))
        o['tot'] = row('PP2_DINH_LUONG', 100, 'C')[0]
        o['sumF'] = sum(N(row('PP2_DINH_LUONG', x, 'F')[0]) for x in range(91, 96))
        o['parts'] = sum((N(row('PP2_DINH_LUONG', x, 'C')[0]) + N(row('PP2_DINH_LUONG', x, 'D')[0])) * N(row('PP2_DINH_LUONG', x, 'E')[0]) for x in range(91, 96))
        o['canF'] = sum(N(row('PP2_DINH_LUONG', x, 'C')[0]) for x in range(91, 96))
        ok_la = True
        for k, v in r.items():
            if k == 'id': continue
            tong[k] += 1
            a = o[k]
            same = (abs(a - v) <= 1e-9 * max(1, abs(v))) if isinstance(v, (int, float)) and isinstance(a, float) else a == v
            if same: khop[k] += 1
            else: ok_la = False; lech.append((c['id'], k, a, v))
        la_dung += ok_la
    return len(cases), la_dung, tong, khop, lech, len(edge), time.time() - t0


if __name__ == '__main__':
    n, ld, tong, khop, lech, ne, dt = chay(sys.argv[1])
    print('T4 hợp/xung: %d/%d lá khớp đủ mọi trường (%d ca biên) – %.1f giây' % (ld, n, ne, dt))
    for k in tong: print('  %-9s %4d/%4d' % (k, khop[k], tong[k]))
    for x in lech[:20]: print('  LỆCH', x)
