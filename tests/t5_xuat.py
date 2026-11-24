# -*- coding: utf-8 -*-
"""T5 – Xuất PDF / XLSX / JSON (YÊU CẦU 5.4) trên lá T7 (đáp án oracle_dump.py):
(a) JSON: mọi ô {sheet: {ô: giá trị}} trùng đáp án oracle (đổi true/false → 1/0 như getDataArray);
(b) XLSX: mở lại bằng openpyxl, mọi ô công thức của 28 sheet hiển thị trùng giá trị JSON (chữ: sau làm sạch tham chiếu nội bộ); giữ ô gộp, cột/dòng ẩn;
(c) PDF: số trang > 0, mỗi trang có "Trang i/n", có khối KẾT LUẬN và phiên bản quy tắc (đọc lại bằng QtPdf).
Dùng: QT_QPA_PLATFORM=offscreen python -m tests.t5_xuat [số_lá]"""
import sys, os, json, tempfile, time, collections, datetime
from openpyxl.utils.datetime import to_excel
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from battu_core.engine import Engine
from battu_core.dich_vu import KetQua
from battu_core.nhap_lieu import LaSo
from tests.so_dap_an import oracle_nhap, khop, ten_tep
from export.xuat_json import du_lieu, xuat_json
from export.xuat_xlsx import xuat_xlsx, ten_trang
from export.noi_dung import SHEET_AN
from export.lam_sach import lam_sach

GOC = os.path.dirname(os.path.abspath(__file__))


def doc_pdf(p):
    from PySide6.QtPdf import QPdfDocument
    from PySide6.QtCore import QEventLoop
    d = QPdfDocument(); d.load(p)
    return d.pageCount(), [d.getAllText(i).text() for i in range(d.pageCount())]


from export.noi_dung import PHIEN_BAN_QUY_TAC


def chay(n=5):
    from ui.giao_dien import tao_app
    app = tao_app([])
    from export.xuat_pdf import xuat_pdf
    cases = json.load(open(os.path.join(GOC, 't7_ngau_nhien.json'), encoding='utf-8'))
    da_dir = os.path.join(GOC, 'dap_an_mo_rong', 'T7_NGAU_NHIEN'); tmp = tempfile.mkdtemp(); t0 = time.time()
    tj = collections.Counter(); dj = collections.Counter(); tx = collections.Counter(); dx = collections.Counter(); lech = []; pdf = []
    for i, (c, nhap) in enumerate(oracle_nhap(cases)):
        if i >= n: break
        da = json.load(open(os.path.join(da_dir, ten_tep(c['id'])), encoding='utf-8'))
        kq = KetQua(LaSo(ho_ten=c['id'], cach_nhap='canchi'), Engine({'NHAP_LIEU': nhap}, None), 0)
        d = du_lieu(kq)['ket_qua']
        for sh, o in da.items():
            for ref, v in o.items():
                x = d[sh][ref]; x = (1.0 if x else 0.0) if isinstance(x, bool) else x
                tj['JSON'] += 1
                if khop(x, v): dj['JSON'] += 1
                else: lech.append((c['id'], 'JSON', sh, ref, x, v))
        pj = xuat_json(kq, os.path.join(tmp, 'a.json')); assert json.load(open(pj, encoding='utf-8'))['ket_qua'] == json.loads(json.dumps(d))
        px = xuat_xlsx(kq, os.path.join(tmp, 'a.xlsx'))
        from openpyxl import load_workbook
        wb = load_workbook(px)
        for sh, o in d.items():
            if sh in SHEET_AN or sh in ('HUONG_DAN', 'TRANG_CHU', 'NHAP_LIEU'): continue
            ws = wb[ten_trang(sh)]
            for ref, v in o.items():
                x = ws[ref].value; tx['XLSX'] += 1
                if isinstance(x, datetime.datetime): x = to_excel(x)      # ô định dạng ngày giờ → số sê-ri
                if isinstance(v, str): v = lam_sach(v)          # XLSX hiển thị chữ đã bỏ tham chiếu nội bộ
                ok = (x is None and v == '') or x == v or (isinstance(v, float) and isinstance(x, (int, float)) and khop(float(x), v))
                if ok: dx['XLSX'] += 1
                else: lech.append((c['id'], 'XLSX', sh, ref, x, v))
        an_ok = wb[ten_trang('KET_LUAN')].row_dimensions[201].hidden and wb[ten_trang('LUC_THAN')].column_dimensions['M'].hidden and len(wb[ten_trang('KET_LUAN')].merged_cells.ranges) > 0
        pp = os.path.join(tmp, 'a.pdf'); so = xuat_pdf(kq, pp)
        n_, txt = doc_pdf(pp)
        pdf.append(n_ == so and n_ > 0 and all(('Trang %d/%d' % (k + 1, n_)) in t for k, t in enumerate(txt))
                   and 'Luận giải chung' in txt[0] and ('Quy tắc ' + PHIEN_BAN_QUY_TAC) in txt[0] and an_ok)
    return tj, dj, tx, dx, lech, pdf, time.time() - t0


if __name__ == '__main__':
    tj, dj, tx, dx, lech, pdf, dt = chay(int(sys.argv[1]) if len(sys.argv) > 1 else 5)
    print('T5 xuất – %.1f giây' % dt)
    print('  JSON  %d / %d' % (dj['JSON'], tj['JSON'])); print('  XLSX  %d / %d' % (dx['XLSX'], tx['XLSX']))
    print('  PDF   %d / %d lá đạt (số trang, "Trang i/n" mỗi trang, khối KẾT LUẬN, phiên bản quy tắc, ẩn/gộp XLSX)' % (sum(pdf), len(pdf)))
    for x in lech[:30]: print('  LỆCH', x)
