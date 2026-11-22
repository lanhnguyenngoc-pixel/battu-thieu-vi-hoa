# -*- coding: utf-8 -*-
"""Module sheet CSDL – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.9.xlsx (SHA-256 671245aea574f577f3cb312db9b3db0d1ce1973b107488ca7b828a9ceb4d60da).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'CSDL'
SI = 24

def f0(E, r, c):
    # Ô D103, D104, D105, D106, D107, D108, D109, D110, D111, D112, D113, D114 … (14 ô)
    # = IF(COUNTIF(F103:H103,MOD(C103-1,5))>0,MOD(C103-1,5),IF(F103<>C103,F103,IF(AND(G103<>"",G103<>C103),G103,H103)))
    return (X.MOD(X.sub(E.v(24,r,c-1),1.0), 5.0) if X.truthy(X.gt(X.COUNTIF(E.rg(24,r,c+2,r,c+4), X.MOD(X.sub(E.v(24,r,c-1),1.0), 5.0)),0.0)) else (E.v(24,r,c+2) if X.truthy(X.ne(E.v(24,r,c+2),E.v(24,r,c-1))) else (E.v(24,r,c+3) if X.truthy(X.AND(X.ne(E.v(24,r,c+3),''), X.ne(E.v(24,r,c+3),E.v(24,r,c-1)))) else E.v(24,r,c+4))))

def f1(E, r, c):
    # Ô E103, E104, E105, E106, E107, E108, E109, E110, E111, E112, E113, E114 … (14 ô)
    # = MOD(C103-2,5)
    return X.MOD(X.sub(E.v(24,r,c-2),2.0), 5.0)

CELLS = {
    'D103': (f0, 103, 4),
    'E103': (f1, 103, 5),
    'D104': (f0, 104, 4),
    'E104': (f1, 104, 5),
    'D105': (f0, 105, 4),
    'E105': (f1, 105, 5),
    'D106': (f0, 106, 4),
    'E106': (f1, 106, 5),
    'D107': (f0, 107, 4),
    'E107': (f1, 107, 5),
    'D108': (f0, 108, 4),
    'E108': (f1, 108, 5),
    'D109': (f0, 109, 4),
    'E109': (f1, 109, 5),
    'D110': (f0, 110, 4),
    'E110': (f1, 110, 5),
    'D111': (f0, 111, 4),
    'E111': (f1, 111, 5),
    'D112': (f0, 112, 4),
    'E112': (f1, 112, 5),
    'D113': (f0, 113, 4),
    'E113': (f1, 113, 5),
    'D114': (f0, 114, 4),
    'E114': (f1, 114, 5),
    'D115': (f0, 115, 4),
    'E115': (f1, 115, 5),
    'D116': (f0, 116, 4),
    'E116': (f1, 116, 5),
}
