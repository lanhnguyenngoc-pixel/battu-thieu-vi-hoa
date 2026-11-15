# -*- coding: utf-8 -*-
"""Module sheet BAO_PHU – SINH TỰ ĐỘNG bởi tools/dich_cong_thuc.py, KHÔNG SỬA TAY.
Nguồn: BAT_TU_THIEU_VI_HOA_6.6.xlsx (SHA-256 ef3858360c0984a0528d265328dcd5caaf8d951a2feed88333eec94b8b316638).
Phạm vi: toàn sheet. Mỗi hàm f<n>(E, r, c) = một mẫu công thức; r, c là dòng, cột của ô (đánh số từ 1)."""
from battu_core import xl as X

SHEET = 'BAO_PHU'
SI = 30

def f0(E, r, c):
    # Ô C705, C706, C707, C708, C709, C710, C711
    # = COUNTIF($B$5:$B$701,B705)
    return X.COUNTIF(E.rg(30,5,2,701,2), E.v(30,r,c-1))

def f1(E, r, c):
    # Ô D705, D706, D707, D708, D709, D710, D711
    # = COUNTIFS($B$5:$B$701,B705,$E$5:$E$701,"Đã có")+0.5*COUNTIFS($B$5:$B$701,B705,$E$5:$E$701,"Có một phần")
    return X.add(X.COUNTIFS(E.rg(30,5,2,701,2), E.v(30,r,c-2), E.rg(30,5,5,701,5), 'Đã có'),X.mul(0.5,X.COUNTIFS(E.rg(30,5,2,701,2), E.v(30,r,c-2), E.rg(30,5,5,701,5), 'Có một phần')))

def f2(E, r, c):
    # Ô E705, E706, E707, E708, E709, E710, E711
    # = COUNTIFS($B$5:$B$701,B705,$F$5:$F$701,"Đã có")+COUNTIFS($B$5:$B$701,B705,$F$5:$F$701,"Mới thêm*")+0.5*COUNTIFS($B$5:$B$701,B705,$F$5:$F$701,"Có một ph
    #   ần")
    return X.add(X.add(X.COUNTIFS(E.rg(30,5,2,701,2), E.v(30,r,c-3), E.rg(30,5,6,701,6), 'Đã có'),X.COUNTIFS(E.rg(30,5,2,701,2), E.v(30,r,c-3), E.rg(30,5,6,701,6), 'Mới thêm*')),X.mul(0.5,X.COUNTIFS(E.rg(30,5,2,701,2), E.v(30,r,c-3), E.rg(30,5,6,701,6), 'Có một phần')))

def f3(E, r, c):
    # Ô F705, F706, F707, F708, F709, F710, F711
    # = COUNTIFS($B$5:$B$701,B705,$F$5:$F$701,"Không lượng hóa được")
    return X.COUNTIFS(E.rg(30,5,2,701,2), E.v(30,r,c-4), E.rg(30,5,6,701,6), 'Không lượng hóa được')

def f4(E, r, c):
    # Ô G705, G706, G707, G708, G709, G710, G711
    # = IF(C705-F705=0,"",D705/(C705-F705))
    return ('' if X.truthy(X.eq(X.sub(E.v(30,r,c-4),E.v(30,r,c-1)),0.0)) else X.div(E.v(30,r,c-3),X.sub(E.v(30,r,c-4),E.v(30,r,c-1))))

def f5(E, r, c):
    # Ô H705, H706, H707, H708, H709, H710, H711
    # = IF(C705-F705=0,"",E705/(C705-F705))
    return ('' if X.truthy(X.eq(X.sub(E.v(30,r,c-5),E.v(30,r,c-2)),0.0)) else X.div(E.v(30,r,c-3),X.sub(E.v(30,r,c-5),E.v(30,r,c-2))))

CELLS = {
    'C705': (f0, 705, 3),
    'D705': (f1, 705, 4),
    'E705': (f2, 705, 5),
    'F705': (f3, 705, 6),
    'G705': (f4, 705, 7),
    'H705': (f5, 705, 8),
    'C706': (f0, 706, 3),
    'D706': (f1, 706, 4),
    'E706': (f2, 706, 5),
    'F706': (f3, 706, 6),
    'G706': (f4, 706, 7),
    'H706': (f5, 706, 8),
    'C707': (f0, 707, 3),
    'D707': (f1, 707, 4),
    'E707': (f2, 707, 5),
    'F707': (f3, 707, 6),
    'G707': (f4, 707, 7),
    'H707': (f5, 707, 8),
    'C708': (f0, 708, 3),
    'D708': (f1, 708, 4),
    'E708': (f2, 708, 5),
    'F708': (f3, 708, 6),
    'G708': (f4, 708, 7),
    'H708': (f5, 708, 8),
    'C709': (f0, 709, 3),
    'D709': (f1, 709, 4),
    'E709': (f2, 709, 5),
    'F709': (f3, 709, 6),
    'G709': (f4, 709, 7),
    'H709': (f5, 709, 8),
    'C710': (f0, 710, 3),
    'D710': (f1, 710, 4),
    'E710': (f2, 710, 5),
    'F710': (f3, 710, 6),
    'G710': (f4, 710, 7),
    'H710': (f5, 710, 8),
    'C711': (f0, 711, 3),
    'D711': (f1, 711, 4),
    'E711': (f2, 711, 5),
    'F711': (f3, 711, 6),
    'G711': (f4, 711, 7),
    'H711': (f5, 711, 8),
}
