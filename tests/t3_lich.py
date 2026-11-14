# -*- coding: utf-8 -*-
"""T3 – lịch: 710 thời điểm (05_KIEM_THU/lich/dates.json), đáp án lich/res.json (do scripts/run_lich.py lập).
Đặt ô nhập như run_lich.py: workbook mặc định + C12 Tự động, C13 Tự nhập, C9 Biết giờ, C27 = 7, C26:G26 = ngày giờ.
So 8 ô tứ trụ NHAP_LIEU C37:F37 (can) và C37:F38 (chi). Cột thứ 9 của res.json là DUNG_HY_KY!C180
(cảnh báo sát giờ/sát tiết) – thuộc GĐ2, sẽ so khi cài DUNG_HY_KY.
Dùng: python -m tests.t3_lich <thư_mục_05_KIEM_THU>"""
import sys, os, json, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.hien_thi import doc_nhu_uno

O = [c + '37' for c in 'CDEF'] + [c + '38' for c in 'CDEF']
from battu_core import sheets as _S
CO_C180 = 'DUNG_HY_KY' in _S.PHAM_VI   # ô thứ 9: có từ GĐ2


def chay(kt):
    dates = json.load(open(os.path.join(kt, 'lich', 'dates.json')))
    res = json.load(open(os.path.join(kt, 'lich', 'res.json'), encoding='utf-8'))
    dung = 0; lech = []; t0 = time.time()
    for d, r in zip(dates, res):
        nhap = {'C12': 'Tự động', 'C13': 'Tự nhập', 'C9': 'Biết giờ', 'C27': 7.0}
        for col, v in zip('CDEFG', d): nhap[col + '26'] = float(v)
        E = Engine({'NHAP_LIEU': nhap})
        kq = [E.gia_tri('NHAP_LIEU', o) for o in O]
        if CO_C180: kq.append(doc_nhu_uno(E, 'DUNG_HY_KY', 'C180'))
        if kq == r[:len(kq)]: dung += 1
        else: lech.append((d, kq, r[:len(kq)]))
    return dung, len(dates), lech, time.time() - t0


if __name__ == '__main__':
    dung, n, lech, dt = chay(sys.argv[1])
    print('T3 lịch: %d/%d thời điểm khớp đủ %s (%.1f giây)' % (dung, n, '9 ô (tứ trụ + DUNG_HY_KY!C180)' if CO_C180 else '8 ô tứ trụ', dt))
    for x in lech[:20]: print('  LỆCH', x)
