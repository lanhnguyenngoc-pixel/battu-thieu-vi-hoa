"""Tạo quần thể lá ngẫu nhiên (ngày dương lịch thật) để đo tỷ lệ nền của mỗi mức kết luận."""
import sys, random, json, os, pickle
sys.path.insert(0, '/home/claude/battu'); sys.path.insert(0, '/home/claude/battu/tools')
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh
from do_luan_giai import lay
from gd_b_dac_trung import quy_tac, SHEET
F = '/home/claude/battu/tai_lieu/giai_doan_A/quan_the.pkl'
def tao(n=300, seed=7):
    Q = pickle.load(open(F, 'rb')) if os.path.exists(F) else []
    rnd = random.Random(seed); rnd.random()
    for _ in range(len(Q)): rnd.randint(1940, 2000); rnd.randint(1, 12); rnd.randint(1, 28); rnd.randint(0, 23); rnd.random()
    while len(Q) < n:
        y, m, d, h = rnd.randint(1940, 2000), rnd.randint(1, 12), rnd.randint(1, 28), rnd.randint(0, 23); g = 'Nam' if rnd.random() < .5 else 'Nữ'
        s = LaSo(cach_nhap='duong', ngay=d, thang=m, nam=y, gio=h, phut=30, gioi_tinh=g, nam_xem=2026)
        E = tinh(s).E
        tt = ','.join(str(E.gia_tri('KET_LUAN', c + '15')) for c in 'CDEF')
        Q.append(dict(tu_tru=tt, gioi=g, nam=y, o=lay(E), qt={lv: quy_tac(E, *SHEET[lv]) for lv in SHEET}, than=str(E.gia_tri('DOI_CHIEU', 'C34'))))
        if len(Q) % 25 == 0: pickle.dump(Q, open(F, 'wb')); print(len(Q), flush=True)
    pickle.dump(Q, open(F, 'wb')); return Q
if __name__ == '__main__':
    Q = tao(int(sys.argv[1]) if len(sys.argv) > 1 else 300); print(Q[0]['tu_tru'], Q[0]['gioi'])
