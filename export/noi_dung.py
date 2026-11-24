# -*- coding: utf-8 -*-
"""Nội dung dùng chung cho xuất PDF / XLSX / in (không phụ thuộc Qt).
Lấy chữ hiển thị của từng ô đúng như giao diện (định dạng ô + định dạng có điều kiện của workbook),
bỏ cột/dòng ẩn của workbook (ô tính phụ) và vùng phụ trợ KET_LUAN dòng 200–251."""
import datetime
from openpyxl.utils.cell import range_boundaries
from battu_core.du_lieu import model, ref_of
from battu_core.hien_thi import hien_thi_o
from battu_core.dinh_dang_dk import kieu_sheet
from battu_core import sheets as S

PHIEN_BAN_QUY_TAC = '6.11'
PHIEN_BAN_PM = '1.9.2'
TEN_PHAN_MEM = 'Bát Tự Thiệu Vĩ Hoa'
SHEET_AN = ('MAU_THUAN', 'KIEM_THU', 'CSDL')          # chỉ tính ẩn (quyết định 14/11/2026)
DONG_PHU_TRO_KL = range(200, 252)

# Mục KET_LUAN theo lĩnh vực (YÊU CẦU 5.2) – (chữ cái, dòng tiêu đề, dòng KẾT LUẬN)
MUC_KL = [('B', 14, None), ('C', 28, 29), ('D', 67, 68), ('E', 78, 79), ('F', 92, 93), ('G', 107, 108),
          ('H', 134, 135), ('I', 159, 160), ('J', 178, 179), ('K', 192, 193)]


def dong_muc_kl(i):
    r1 = MUC_KL[i][1]
    r2 = MUC_KL[i + 1][1] - 1 if i + 1 < len(MUC_KL) else DONG_PHU_TRO_KL[0] - 1
    return list(range(r1, r2 + 1))


def _tieu_de_kl(chu):
    return {'B': 'B. Lá số – tứ trụ, thập thần, tàng can', 'C': 'C. Bản mệnh', 'D': 'D. Tính cách – năng khiếu',
            'E': 'E. Tài lộc – giàu nghèo', 'F': 'F. Sự nghiệp – công danh', 'G': 'G. Gia đạo', 'H': 'H. Sức khỏe (tham khảo)',
            'I': 'I. Vận trình', 'J': 'J. Hóa giải – nên/tránh', 'K': 'K. Độ tin cậy – ghi chú nguồn'}[chu]


# Mục có thể chọn khi xuất PDF: (khóa, tên, sheet, dòng hoặc None = cả sheet). Sheet quá rộng (LUU_NIEN 112 cột,
# THEO_DOI) xem ở bản XLSX.
MUC_XUAT = [('KL_%s' % chu, _tieu_de_kl(chu), 'KET_LUAN', dong_muc_kl(i)) for i, (chu, _, _) in enumerate(MUC_KL)] + [
    ('LUAN_GIAI', 'Luận giải (8 mục + tổng thể)', 'LUAN_GIAI', None), ('DOC_NHANH', 'Đọc nhanh', 'DOC_NHANH', None),
    ('NGU_HANH_THIEU', 'Ngũ hành thiếu (bảng)', 'NGU_HANH_THIEU', None), ('TRUNG_GIAN', 'Trung gian', 'TRUNG_GIAN', None),
    ('PP1_DINH_TINH', 'PP1 – Định tính', 'PP1_DINH_TINH', None), ('PP2_DINH_LUONG', 'PP2 – Định lượng', 'PP2_DINH_LUONG', None),
    ('THONG_CAN', 'Thông căn', 'THONG_CAN', None), ('DOI_CHIEU', 'Đối chiếu PP1 – PP2', 'DOI_CHIEU', None),
    ('DUNG_HY_KY', 'Dụng – hỷ – kỵ thần', 'DUNG_HY_KY', None), ('DAI_VAN', 'Đại vận', 'DAI_VAN', None),
    ('LUC_THAN', 'Lục thân', 'LUC_THAN', None), ('HON_NHAN', 'Hôn nhân', 'HON_NHAN', None), ('MENH_CUC', 'Mệnh cục', 'MENH_CUC', None),
    ('TAI_LOC', 'Tài lộc', 'TAI_LOC', None), ('TO_CHAT', 'Tố chất', 'TO_CHAT', None), ('SU_NGHIEP', 'Sự nghiệp', 'SU_NGHIEP', None),
    ('SUC_KHOE', 'Sức khỏe (tham khảo)', 'SUC_KHOE', None), ('THAN_SAT', 'Thần sát', 'THAN_SAT', None), ('SAO', 'Sao', 'SAO', None),
    ('BAO_PHU', 'Bảo phù', 'BAO_PHU', None), ('KHONG_RO_GIO', 'Không rõ giờ', 'KHONG_RO_GIO', None),
    ('SU_KIEN', 'Đối chiếu sự việc', 'SU_KIEN', None), ('THAM_KHAO', 'Tham khảo ngoài sách', 'THAM_KHAO', None)]
MUC_MAC_DINH = ['KL_%s' % chu for chu, _, _ in MUC_KL] + ['NGU_HANH_THIEU']


def chu_o(E, sheet, r, c, sach=True):
    """Chữ hiển thị của ô; sach=True: bỏ tham chiếu nội bộ của workbook (export.lam_sach).
    Chữ bị workbook rút gọn LEFT(…,N)&"…" được hiện đầy đủ (export.khoi_phuc_cat); khoảng tuổi đại vận được ghi kèm
    năm dương lịch (export.ghi_tuoi) – giá trị ô không đổi."""
    from export.lam_sach import lam_sach
    from export.khoi_phuc_cat import khoi_phuc
    from export.ghi_tuoi import them_nam
    M = model(); si = M.si[sheet]
    if (r, c) in M.formula_cells[si]: t = khoi_phuc(E, sheet, r, c, hien_thi_o(E, sheet, ref_of(r, c)))
    else:
        v = E.consts[si].get((r, c))
        if v is None: return ''
        t = hien_thi_o(E, sheet, ref_of(r, c)) if isinstance(v, (int, float)) else str(v)
    if (r, c) in M.formula_cells[si]: t = them_nam(E, t)
    return lam_sach(t) if sach else t


def la_tieu_de(t):
    return len(t) > 6 and t[:1].isupper() and (t.isupper() or t[:2] in tuple('%s.' % x for x in 'ABCDEFGHIJK'))


def khoi(E, sheet, dong=None, bo_dong_trong=True):
    """Một khối bảng đã bỏ cột/dòng ẩn → dict(cot, dong, o{(r,c): (chữ, kiểu)}, gop{(r,c): (số dòng, số cột)}, che set)."""
    M = model(); si = M.si[sheet]; bc = M.bo_cuc.get(sheet, {})
    keys = set(M.consts[si]) | set(M.formula_cells[si])
    nr = max(r for r, c in keys)
    dong = list(dong) if dong else list(range(1, nr + 1))
    an_d = set(bc.get('an_dong', [])) | (set(DONG_PHU_TRO_KL) if sheet == 'KET_LUAN' else set())
    an_c = set(bc.get('an_cot', []))
    dong = [r for r in dong if r not in an_d]
    dk = kieu_sheet(E, sheet); o = {}
    tap = set(dong)
    for (r, c) in keys:
        if r not in tap or c in an_c: continue
        t = chu_o(E, sheet, r, c)
        if t == '': continue
        k = dict(dk.get((r, c), {}))
        if (r, c) not in M.formula_cells[si] and la_tieu_de(t): k.setdefault('dam', True); k['tieu_de'] = True
        o[(r, c)] = (t, k)
    gop = {}; che = set()
    for g in bc.get('gop', []):
        c1, r1, c2, r2 = range_boundaries(g)
        if (r1, c1) not in o: continue
        rs = [r for r in range(r1, r2 + 1) if r in tap]; cs = [c for c in range(c1, c2 + 1) if c not in an_c]
        if len(rs) * len(cs) <= 1: continue
        gop[(r1, c1)] = (rs, cs)
        for r in rs:
            for c in cs:
                if (r, c) != (r1, c1): che.add((r, c))
    co_chu_dong = {r for r, c in o} | {r for (r, c), (rs, cs) in gop.items() for r in rs}
    if bo_dong_trong: dong = [r for r in dong if r in co_chu_dong]
    cot = sorted({c for r, c in o} | {c for (r, c), (rs, cs) in gop.items() for c in cs})
    # gộp chỉ còn trong các cột/dòng giữ lại
    gop2 = {}
    tap_d = set(dong); tap_c = set(cot)
    for k, (rs, cs) in gop.items():
        rs = [r for r in rs if r in tap_d]; cs = [c for c in cs if c in tap_c]
        gop2[k] = (len(rs), len(cs))
    return {'sheet': sheet, 'cot': cot, 'dong': dong, 'o': o, 'gop': gop2, 'che': che,
            'rong': {c: float(bc.get('rong', {}).get(str(c), bc.get('rong', {}).get(c, 10.0)) or 10.0) for c in cot}}


def thong_tin_la(kq):
    s = kq.la_so; E = kq.E
    tru = ' · '.join(chu_o(E, 'KET_LUAN', 15, c) for c in range(3, 7))
    if s.cach_nhap == 'duong':
        sinh = '%02d/%02d/%d %02d:%02d (UTC%+g)' % (s.ngay, s.thang, s.nam, s.gio or 0, s.phut or 0, s.mui_gio)
    elif s.cach_nhap == 'am':
        sinh = 'Âm lịch %d/%d%s/%d %02d:%02d' % (s.am_ngay, s.am_thang, ' (nhuận)' if s.am_nhuan else '', s.am_nam, s.gio or 0, s.phut or 0)
    else:
        sinh = 'Tự nhập can chi'
    if s.khong_ro_gio: sinh += ' · Không rõ giờ (đang xem giờ %s)' % s.gio_dang_xem
    return [('Họ tên', s.ho_ten or '(chưa đặt tên)'), ('Ngày giờ sinh', sinh), ('Giới tính', s.gioi_tinh),
            ('Tứ trụ (năm · tháng · ngày · giờ)', tru), ('Năm xem', chu_o(E, 'NHAP_LIEU', 31, 3))]


def chu_ky():
    return 'Phiên bản %s · Quy tắc %s' % (PHIEN_BAN_PM, PHIEN_BAN_QUY_TAC)


def ngay_in():
    return datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
