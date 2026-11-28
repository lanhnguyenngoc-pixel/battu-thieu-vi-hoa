# -*- coding: utf-8 -*-
"""T14 – Bản luận giải tổng hợp không cần mạng (đề xuất hoàn thiện 15/11/2026, giai đoạn T4).
Trên lá mẫu Nguyễn Ngọc Lanh + 20 lá T7:
  • đủ tóm tắt + 10 mục theo đúng thứ tự; mỗi đại vận / mỗi năm xuất hiện đúng một lần trong bảng vận trình;
  • không mâu thuẫn trong bản: một năm không vừa ở danh sách "thuận" vừa ở danh sách "cần thận trọng"
    (gia đạo; tiền bạc); mức của năm trong bảng = mức của lớp phán đoán;
  • không còn "(xem căn cứ sách)", không còn chữ workbook bị cắt (đoạn kết thúc bằng chữ cái + "…");
  • PDF ≤ 6 trang A4 (chủ dự án duyệt 25/11/2026); .docx là gói hợp lệ, document.xml đọc được.
Dùng: python -m tests.t14_ban_luan_giai"""
import sys, os, json, re, time, tempfile, zipfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from xml.dom import minidom
from battu_core.engine import Engine
from battu_core.dich_vu import KetQua
from battu_core.nhap_lieu import LaSo
from battu_core.tong_hop.viet_ban import ban_luan_giai
from tests.so_dap_an import oracle_nhap

GOC = os.path.dirname(os.path.abspath(__file__))
MUC = ['Tóm tắt nhanh', '1. Lá số', '2. Bản mệnh', '3. Tính cách', '4. Học vấn', '5. Tài lộc', '6. Hôn nhân', '7. Cha mẹ',
       '8. Sức khỏe', '9. Đại vận', '10. Hóa giải']


def _chu(B):
    ra = []
    for k in B:
        if k[0] in ('h1', 'h2', 'p', 'nho'): ra.append(k[1])
        elif k[0] == 'ul': ra += k[1]
        elif k[0] == 'bang': ra += k[1] + [o for r in k[2] for o in r]
    return '\n'.join(ra)


def _ds(t, nhan):
    m = re.search(re.escape(nhan) + r'[^:]*:\s*([^.]*)', t)
    return set(re.findall(r'\b(\d{4})\b', m.group(1))) if m else set()


def kiem(ten, B, pd, tmp, pdf=True):
    loi = []; h2 = [k[1] for k in B if k[0] == 'h2']
    if len(h2) != len(MUC) or any(not h.startswith(m) for h, m in zip(h2, MUC)): loi.append((ten, 'mục', h2))
    bang = [k for k in B if k[0] == 'bang']
    dvb = next((b for b in bang if b[1][0] == 'Đại vận'), None); lnb = next((b for b in bang if 'tuổi mụ' in b[1][0]), None)
    if dvb:
        ten_dv = [r[0].split(' (')[0] for r in dvb[2]]
        if len(ten_dv) != len(set(ten_dv)): loi.append((ten, 'đại vận lặp', ten_dv))
    if lnb:
        nam = [r[0][:4] for r in lnb[2]]
        if len(nam) != len(set(nam)): loi.append((ten, 'năm lặp', nam))
        muc = {str(y['nam']): y['muc'] for y in pd['luu_nien']}
        for r in lnb[2]:
            if r[1].strip('*') != muc.get(r[0][:4]): loi.append((ten, 'mức năm khác phán đoán', r[0], r[1]))
    t = _chu(B)
    for a, b in (('Năm thuận gia đạo', 'Năm cần nhường nhịn'), ('Năm thuận kết hôn', 'Năm cần thận trọng'),
                 ('Năm thu nhập thuận', 'Năm cần giữ tiền')):
        cung = _ds(t, a) & _ds(t, b)
        if cung: loi.append((ten, 'năm ở cả "%s" và "%s"' % (a, b), sorted(cung)))
    if '(xem căn cứ sách)' in t: loi.append((ten, 'còn "(xem căn cứ sách)"'))
    cat = re.findall(r'[A-Za-zÀ-ỹ]…(?=\s|$|[;.,)])', t)
    if cat: loi.append((ten, 'chữ bị cắt', cat[:3]))
    from export.xuat_luan_giai import xuat_pdf, xuat_docx
    p = os.path.join(tmp, 'a.docx'); xuat_docx(B, p)
    try:
        with zipfile.ZipFile(p) as z: minidom.parseString(z.read('word/document.xml'))
    except Exception as e: loi.append((ten, 'docx hỏng', repr(e)))
    so = None
    if pdf:
        so = xuat_pdf(B, os.path.join(tmp, 'a.pdf'))
        if not (1 <= so <= 6): loi.append((ten, 'PDF %d trang' % so))
    return loi, so


def chay():
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    from ui.giao_dien import tao_app
    app = tao_app([])
    t0 = time.time(); tmp = tempfile.mkdtemp(); loi = []; trang = []
    from battu_core.dich_vu import tinh
    s = LaSo(ho_ten='Nguyễn Ngọc Lanh', cach_nhap='duong', ngay=28, thang=11, nam=1978, gio=8, phut=30, gioi_tinh='Nam', nam_xem=2026)
    B, hs, pd = ban_luan_giai(tinh(s)); l, so = kiem('mẫu', B, pd, tmp); loi += l; trang.append(so)
    t = _chu(B)
    for can in ('Thân vượng', 'Dụng thần Hỏa', '2032 Nhâm Tí', 'Rất xấu', 'Thọ – yểu: mệnh yểu', 'Khá giả', 'tự hình Ngọ–Ngọ',
                'Học tập – thi cử', 'Năm 2028 Mậu Thân', 'Bốn giai đoạn đời'):
        if can not in t: loi.append(('mẫu', 'thiếu "%s"' % can))
    for c, nhap in oracle_nhap(json.load(open(os.path.join(GOC, 't7_ngau_nhien.json'), encoding='utf-8'))):
        E = Engine({'NHAP_LIEU': nhap}, None); E.tinh_tat_ca()
        kq = KetQua(LaSo(ho_ten=c['id'], cach_nhap='canchi'), E, 0)
        try:
            B, hs, pd = ban_luan_giai(kq)
        except Exception as e:
            loi.append((c['id'], 'không soạn được', repr(e))); continue
        l, so = kiem(c['id'], B, pd, tmp); loi += l; trang.append(so)
    return loi, trang, time.time() - t0


if __name__ == '__main__':
    loi, trang, dt = chay()
    print('T14 bản luận giải – %.1f giây: %d lá, số trang PDF %s' % (dt, len(trang), sorted(set(x for x in trang if x))))
    print('  ĐẠT' if not loi else '  LỖI %d' % len(loi))
    for x in loi[:20]: print('   ', x)
    sys.exit(1 if loi else 0)
