# -*- coding: utf-8 -*-
"""T11 – Lớp hiển thị / in (đề xuất hoàn thiện 15/11/2026, giai đoạn T1): C1 câu sách, C2 chữ bị cắt, C3 ghi tuổi.
Chạy trên 20 lá T7 (tests/t7_ngau_nhien.json). Không cần đáp án LibreOffice: giá trị ô đã do T7 kiểm.
  (a) Mọi ô của LUAN_GIAI, THEO_DOI, KET_LUAN: chữ hiển thị không có "…" thì lớp hiển thị không đổi gì ngoài ghi năm
      cho tuổi đại vận; chữ khôi phục không ngắn hơn chữ gốc.
  (b) Ô rút gọn kiểu IF(LEN(x)>N, LEFT(x,N)&"…", x) với x là một ô: chữ hiển thị không còn đoạn LEFT(x,N)&"…".
  (c) Bản in (_kl của xuat_pdf) không còn dòng "(xem căn cứ sách)".
  (d) Giá trị ô (Engine) không đổi sau khi hiển thị: so 3 sheet trước – sau.
Dùng: python -m tests.t11_hien_thi"""
import sys, os, re, json, time, inspect, importlib
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.hien_thi import hien_thi_o
from battu_core.du_lieu import model, ref_of
from battu_core import sheets as S
from battu_core import xl as X
from tests.so_dap_an import oracle_nhap

GOC = os.path.dirname(os.path.abspath(__file__))
SHEET = ('LUAN_GIAI', 'THEO_DOI', 'KET_LUAN')
_DON = re.compile(r"X\.cat\(X\.LEFT\(E\.v\((\d+),(\d+),(\d+)\),\s*([0-9.]+)\),'…'\)")


def _bo_nam(t):
    """Bỏ phần năm dương lịch lớp hiển thị ghi thêm cho tuổi đại vận (C3)."""
    return re.sub(r', \d{4}–\d{4}(?=[):])', '', t)


def _nguon_don(sheet):
    """{(r, c): [(si, r, c, N)]} cho ô có LEFT(một ô cố định, N)&"…"."""
    m = importlib.import_module('battu_core.sheets.' + S.MODULES[sheet]); kq = {}
    for ref, (fn, r, c) in m.CELLS.items():
        ds = [(int(a), int(b), int(d), float(n)) for a, b, d, n in _DON.findall(inspect.getsource(fn))]
        if ds: kq[(r, c)] = ds
    return kq


def chay(cases_path=os.path.join(GOC, 't7_ngau_nhien.json')):
    from export.noi_dung import chu_o
    from export.xuat_pdf import _kl
    from export.khoi_phuc_cat import _ek
    M = model(); t0 = time.time(); loi = []; dem = {'o': 0, 'khoi_phuc': 0, 'nguon_don': 0, 'kl': 0}
    nguon = {sh: _nguon_don(sh) for sh in SHEET}
    for c, nhap in oracle_nhap(json.load(open(cases_path, encoding='utf-8'))):
        E = Engine({'NHAP_LIEU': nhap}, None)
        truoc = {sh: dict(E.tinh_sheet(sh)) for sh in SHEET}
        for sh in SHEET:
            si = M.si[sh]
            for (r, cc) in sorted(M.formula_cells[si]):
                goc = hien_thi_o(E, sh, ref_of(r, cc)); moi = chu_o(E, sh, r, cc, sach=False); dem['o'] += 1
                if not isinstance(goc, str): continue
                if '…' not in goc and _bo_nam(moi) != goc:
                    loi.append((c['id'], sh, ref_of(r, cc), 'đổi chữ không bị cắt'))
                if len(moi) < len(goc): loi.append((c['id'], sh, ref_of(r, cc), 'ngắn hơn chữ gốc'))
                if moi != goc and '…' in goc: dem['khoi_phuc'] += 1
                for (s2, r2, c2, n) in nguon[sh].get((r, cc), ()):
                    v = _ek(E).raw(s2, r2, c2)        # chữ nguồn đã khôi phục (nguồn cũng có thể bị cắt)
                    if isinstance(v, str) and len(v) > n:
                        dem['nguon_don'] += 1
                        if (X.LEFT(v, n) + '…') in moi: loi.append((c['id'], sh, ref_of(r, cc), 'còn chữ nguồn bị cắt'))
            for (r, cc) in sorted(M.formula_cells[si]):
                h = _kl(chu_o(E, sh, r, cc)); dem['kl'] += 1
                if '(xem căn cứ sách)' in h: loi.append((c['id'], sh, ref_of(r, cc), 'còn "(xem căn cứ sách)"'))
        for sh in SHEET:
            if dict(E.tinh_sheet(sh)) != truoc[sh]: loi.append((c['id'], sh, '*', 'giá trị ô đổi sau khi hiển thị'))
    return dem, loi, time.time() - t0


if __name__ == '__main__':
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    dem, loi, dt = chay()
    print('T11 hiển thị – %.1f giây: %d ô, %d ô khôi phục chữ cắt, %d lần so chữ nguồn, %d ô bản in' % (
        dt, dem['o'], dem['khoi_phuc'], dem['nguon_don'], dem['kl']))
    print('  ĐẠT' if not loi else '  LỖI %d' % len(loi))
    for x in loi[:30]: print('  ', x)
    sys.exit(1 if loi else 0)
