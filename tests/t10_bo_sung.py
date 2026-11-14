# -*- coding: utf-8 -*-
"""T10 – Ba loại ô nhập bổ sung (bản 1.4.0) trên 20 lá T7 (tests/bien_the.json → T10_BO_SUNG):
  * ngày sinh phối ngẫu NHAP_LIEU C34:G34 (đủ ngày giờ / chỉ năm / trống);
  * việc lớn để tìm giờ sinh KHONG_RO_GIO C11:E22 (3–12 việc; lá chẵn ở chế độ Không rõ giờ, Bảng 12 giờ tính bằng vòng);
  * đánh giá thực tế THEO_DOI N/O/P (15 năm mỗi lá).
(a) đáp án LibreOffice (tools/oracle_mo_rong.py) – so MỌI ô công thức của 33 sheet (t1b.chay_bien_the);
(b) đường nhập của phần mềm: LaSo.phoi_ngau / viec_tim_gio / theo_doi (theo NĂM) → dich_vu.tinh() phải đặt ĐÚNG
    các ô mà oracle đã nhập (năm của dòng THEO_DOI do lõi tính ra);
(c) lưu SQLite → mở lại giữ nguyên ba loại dữ liệu.
Dùng: python -m tests.t10_bo_sung"""
import sys, os, json, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from tests import t1b
from battu_core.nhap_lieu import LaSo
from battu_core.dich_vu import tinh

GOC = os.path.dirname(os.path.abspath(__file__))


def la_so_tu_bien_the(c, d):
    """Dựng LaSo như người dùng nhập trên giao diện từ phần 'riêng' của biến thể."""
    if 'can' in c:
        s = LaSo(cach_nhap='canchi', can=list(c['can']), chi=list(c['chi']))
    else:
        s = LaSo(ngay=c['ngay'], thang=c['thang'], nam=c['nam'], gio=c['gio'], phut=c['phut'], mui_gio=c.get('mui_gio', 7))
    s.gioi_tinh = c.get('gioi_tinh', 'Nam')
    nl = d.get('NHAP_LIEU', {})
    s.khong_ro_gio = nl.get('C9') == 'Không rõ giờ'
    s.phoi_ngau = [nl.get(col + '34') for col in 'CDEFG']
    k = d.get('KHONG_RO_GIO', {})
    s.viec_tim_gio = [[int(k['C%d' % r]), k.get('D%d' % r), k.get('E%d' % r)] for r in range(11, 23) if k.get('C%d' % r) is not None]
    return s


def chay():
    thu_muc = os.path.join(GOC, 'dap_an_mo_rong'); cases_p = os.path.join(GOC, 't7_ngau_nhien.json')
    n, tong, dung, lech = t1b.chay_bien_the(None, thu_muc, 'T10_BO_SUNG', cases_p)
    bt = json.load(open(os.path.join(GOC, 'bien_the.json'), encoding='utf-8'))['T10_BO_SUNG']['rieng']
    cases = {c['id']: c for c in json.load(open(cases_p, encoding='utf-8'))}
    dung_b = tong_b = 0; lech_b = []
    for cid, d in bt.items():
        da = json.load(open(os.path.join(thu_muc, 'T10_BO_SUNG', t1b.ten_tep(cid)), encoding='utf-8'))
        s = la_so_tu_bien_the(cases[cid], d)
        td = d.get('THEO_DOI', {})
        for r in range(12, 112):
            g = [td.get('%s%d' % (col, r)) for col in 'NOP']
            if any(g): s.theo_doi[int(da['THEO_DOI']['A%d' % r])] = g
        kq = tinh(s)
        for sh in ('NHAP_LIEU', 'KHONG_RO_GIO', 'THEO_DOI'):
            mong = d.get(sh, {})
            cot = {'NHAP_LIEU': ['%s34' % col for col in 'CDEFG'] + ['C9'],
                   'KHONG_RO_GIO': ['%s%d' % (col, r) for r in range(11, 23) for col in 'CDE'],
                   'THEO_DOI': ['%s%d' % (col, r) for r in range(12, 112) for col in 'NOP']}[sh]
            for ref in cot:
                tong_b += 1
                a = kq.E.gia_tri(sh, ref); b = mong.get(ref)
                if sh == 'NHAP_LIEU' and ref == 'C9': b = b or 'Biết giờ'
                if a == b: dung_b += 1
                else: lech_b.append((cid, sh, ref, a, b))
    from battu_core.kho_la_so import Kho
    kho = Kho(os.path.join(tempfile.mkdtemp(), 'bs.sqlite'))
    s = LaSo(ngay=26, thang=11, nam=1957, gio=6, phut=0, nam_xem=2026, phoi_ngau=[1959, None, None, None, None],
             viec_tim_gio=[[1983, 'Kết hôn', 'cưới'], [1990, 'Sinh con', None]], theo_doi={1983: ['Tốt', 'cưới', 'Đúng']})
    i = kho.luu(tinh(s), 'bs'); m = kho.la_so(i)
    ok_c = (m.phoi_ngau, m.viec_tim_gio, m.theo_doi) == (s.phoi_ngau, s.viec_tim_gio, s.theo_doi)
    return n, tong, dung, lech, (dung_b, tong_b, lech_b), ok_c


if __name__ == '__main__':
    n, tong, dung, lech, (db, tb, lb), okc = chay()
    print('T10 ô nhập bổ sung – %d lá: %d / %d ô' % (n, sum(dung.values()), sum(tong.values())),
          '| HON_NHAN %d / %d · KHONG_RO_GIO %d / %d · THEO_DOI %d / %d' % (
              dung['HON_NHAN'], tong['HON_NHAN'], dung['KHONG_RO_GIO'], tong['KHONG_RO_GIO'], dung['THEO_DOI'], tong['THEO_DOI']))
    print('  đường nhập giao diện → ô workbook: %d / %d ô · SQLite giữ dữ liệu: %s' % (db, tb, 'ĐẠT' if okc else 'LỖI'))
    for x in (lech + lb)[:20]: print('  LỆCH', x)
    sys.exit(0 if (not lech and not lb and okc and n == 20) else 1)
