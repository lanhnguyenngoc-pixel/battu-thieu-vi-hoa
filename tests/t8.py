# -*- coding: utf-8 -*-
"""T8 – Ngũ hành thiếu, biểu đồ: bộ lá CHUẨN tests/t8_chuan.json (55 lá ngày giờ dương lịch lấy từ data39, dv16;
tứ trụ do lịch tiết khí của bộ công cụ tính). So TỪNG Ô mọi sheet đã cài với đáp án workbook ở hai chế độ
NGU_HANH_THIEU E3: PA1 – Theo trang (T8C_PA1) và PA2 – Theo sách (T8C_PA2) – lấy bằng tools/oracle_mo_rong.py.
Dùng: python -m tests.t8 <thư_mục_05_KIEM_THU> <tests/dap_an_mo_rong>"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from tests.t1b import chay_bien_the

CASES = os.path.join(os.path.dirname(os.path.abspath(__file__)), 't8_chuan.json')


def chay(kt, mr):
    return {ten: chay_bien_the(kt, mr, ten, CASES) for ten in ('T8C_PA1', 'T8C_PA2')}


if __name__ == '__main__':
    for ten, (n, tong, dung, lech) in chay(sys.argv[1], sys.argv[2]).items():
        print('%s: %d lá | %s | tổng %d/%d' % (ten, n, ' '.join('%s %d/%d' % (s, dung[s], tong[s]) for s in tong if s in ('NGU_HANH_THIEU', 'DAI_VAN')),
                                          sum(dung.values()), sum(tong.values())))
        for x in lech[:10]: print('   LỆCH', x[0], '%s!%s' % (x[1], x[2]), repr(x[3])[:80], '| đáp án', repr(x[4])[:80])
