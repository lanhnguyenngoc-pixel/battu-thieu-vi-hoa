# -*- coding: utf-8 -*-
"""Điểm khởi chạy phần mềm (dùng cho PyInstaller).
  BatTuThieuViHoa                 → mở giao diện
  BatTuThieuViHoa --tu-kiem <thư_mục>  → tự kiểm bản đã cài (không giao diện, không cần mạng): tính 3 lá chuẩn, so dấu vân tay
                                    SHA-256 toàn bộ kết quả với giá trị lập trên bản gốc (tools/dau_van_tay.py), xuất PDF/XLSX/JSON
                                    vào thư mục; ghi tu_kiem.txt; mã thoát 0 = ĐẠT."""
import sys, os


def dau_van_tay(kq):
    import json, hashlib
    from export.xuat_json import du_lieu
    d = du_lieu(kq)['ket_qua']
    return hashlib.sha256(json.dumps(d, ensure_ascii=False, sort_keys=True).encode('utf-8')).hexdigest()


LA_TU_KIEM = [dict(ho_ten='Tự kiểm 1', ngay=26, thang=11, nam=1957, gio=6, phut=0, gioi_tinh='Nam', nam_xem=2026),
              dict(ho_ten='Tự kiểm 2', ngay=3, thang=7, nam=1971, gio=9, phut=30, gioi_tinh='Nữ', nam_xem=2026),
              dict(ho_ten='Tự kiểm 3', ngay=15, thang=3, nam=1985, gio=10, phut=30, gioi_tinh='Nam', nam_xem=2026, khong_ro_gio=True)]


def tu_kiem(thu_muc):
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    os.makedirs(thu_muc, exist_ok=True)
    from ui.giao_dien import tao_app
    app = tao_app([sys.argv[0]])
    from battu_core.nhap_lieu import LaSo
    from battu_core.dich_vu import tinh
    from export.xuat_pdf import xuat_pdf
    from export.xuat_xlsx import xuat_xlsx
    from export.xuat_json import xuat_json
    from dau_van_tay_chuan import DAU_VAN_TAY
    dong = []; dat = True
    for i, d in enumerate(LA_TU_KIEM):
        kq = tinh(LaSo(**d)); h = dau_van_tay(kq); ok = h == DAU_VAN_TAY[i]; dat &= ok
        dong.append('%s: %s (%s)' % (d['ho_ten'], 'ĐẠT' if ok else 'LỆCH', h[:16]))
    kq = tinh(LaSo(**LA_TU_KIEM[0]))
    n = xuat_pdf(kq, os.path.join(thu_muc, 'tu_kiem.pdf')); xuat_xlsx(kq, os.path.join(thu_muc, 'tu_kiem.xlsx'))
    xuat_json(kq, os.path.join(thu_muc, 'tu_kiem.json'))
    dong.append('Xuất PDF %d trang, XLSX, JSON: ĐẠT' % n)
    from battu_core.tong_hop.viet_ban import ban_luan_giai
    from export.xuat_luan_giai import xuat_pdf as pdf_th, xuat_docx
    from export.khoi_phuc_cat import _sinh_san
    B, _, _ = ban_luan_giai(kq); m = pdf_th(B, os.path.join(thu_muc, 'tu_kiem_tong_hop.pdf')); xuat_docx(B, os.path.join(thu_muc, 'tu_kiem_tong_hop.docx'))
    ok = 1 <= m <= 5 and _sinh_san() is not None; dat &= ok
    dong.append('Luận giải tổng hợp (không mạng) %d trang, Word; bảng hiện chữ đầy đủ có sẵn: %s' % (m, 'ĐẠT' if ok else 'LỖI'))
    dong.append('KẾT QUẢ: ' + ('ĐẠT' if dat else 'LỖI'))
    open(os.path.join(thu_muc, 'tu_kiem.txt'), 'w', encoding='utf-8').write('\n'.join(dong) + '\n')
    print('\n'.join(dong))
    return 0 if dat else 1


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == '--tu-kiem':
        sys.exit(tu_kiem(sys.argv[2]))
    from ui.giao_dien import main as chay
    chay()


if __name__ == '__main__':
    main()
