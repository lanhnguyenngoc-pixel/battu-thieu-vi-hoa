# -*- coding: utf-8 -*-
"""Lập dấu vân tay SHA-256 toàn bộ kết quả (JSON ket_qua) của 3 lá tự kiểm → dau_van_tay_chuan.py.
Chạy lại mỗi khi quy tắc/workbook đổi (sau khi kiểm thử tổng đạt 100%)."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_app import LA_TU_KIEM, dau_van_tay
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh
from battu_core import sheets as S
h = [dau_van_tay(tinh(LaSo(**d))) for d in LA_TU_KIEM]
open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'dau_van_tay_chuan.py'), 'w').write(
    '# -*- coding: utf-8 -*-\n"""SINH TỰ ĐỘNG bởi tools/dau_van_tay.py – workbook SHA-256 %s."""\nDAU_VAN_TAY = %r\n' % (S.NGUON_SHA256, h))
print(h)
