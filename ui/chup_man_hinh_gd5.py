# -*- coding: utf-8 -*-
"""Chụp ảnh màn hình GĐ5 (offscreen) + chạy thử luồng lưu / mở / tìm / so sánh lá số trên tệp SQLite tạm."""
import os, sys, tempfile
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from ui.giao_dien import tao_app, CuaSoChinh
from battu_core.kho_la_so import Kho
from PySide6.QtWidgets import QAbstractItemView as V

out = sys.argv[1] if len(sys.argv) > 1 else 'anh_gd5'
os.makedirs(out, exist_ok=True)
app = tao_app([])
db = os.path.join(tempfile.mkdtemp(), 'thu.sqlite'); kho = Kho(db)
w = CuaSoChinh(kho); w.show(); t = w.t_nhap


def chup(ten):
    app.processEvents(); w.grab().save(os.path.join(out, ten)); print(ten)


def nhap_dl(ten, d, m, y, h, p=0, gt=0):
    t.cach.dat(0); t.doi_cach(0); t.ho_ten.setText(ten); t.d_ngay.setValue(d); t.d_thang.setValue(m); t.d_nam.setValue(y)
    t.gio.setValue(h); t.phut.setValue(p); t.gt.dat(gt); t.krg.setChecked(False); app.processEvents(); t.tinh(); app.processEvents()


# 1. TÓM TẮT – màn hình đầu
nhap_dl('Lá mẫu 26/11/1957', 26, 11, 1957, 6); w.di(0); chup('01_tom_tat.png')
id1 = kho.luu(t.kq, 'Lá mẫu 26/11/1957', 'Sách TVH', 'lá trong dap_an_o')
# 2–5. Theo lĩnh vực: C bản mệnh, H sức khỏe, LUAN_GIAI (đỏ đậm theo cờ L), tìm kiếm
w.di(2); w.t_lv.mo('C'); chup('02_linh_vuc_C_ban_menh.png')
w.t_lv.mo('H'); chup('03_linh_vuc_H_suc_khoe.png')
w.t_lv.mo('LUAN_GIAI'); app.processEvents(); w.t_lv.muc[10][2].toi_o(71, 1); chup('04_luan_giai_muc7_do_dam.png')
w.t_lv.o_tim.setText('dụng thần'); w.t_lv.tim(); chup('05_tim_kiem.png')
w.t_lv.o_tim.setText(''); w.t_lv.tim()
# 6. Định dạng có điều kiện: Vận trình – Đại vận (đỏ đậm theo C106), Lưu niên (CW), Theo dõi (Tốt/Xấu)
w.di(5); w.t_vt.tab.setCurrentIndex(0); app.processEvents(); w.t_vt.bang['DAI_VAN'].scrollToItem(w.t_vt.bang['DAI_VAN'].item(28, 2), V.PositionAtTop)
chup('06_dai_van_dinh_dang_dk.png')
w.t_vt.tab.setCurrentIndex(2); app.processEvents(); w.t_vt.bang['THEO_DOI'].scrollToItem(w.t_vt.bang['THEO_DOI'].item(10, 0), V.PositionAtTop)
chup('07_theo_doi_tot_xau.png')
# 8. Tham khảo
w.di(9); w.t_tk.tab.setCurrentIndex(0); chup('08_tham_khao_trang_chu.png')
# 9. Lưu thêm 2 lá, danh sách + tìm
nhap_dl('TVH ví dụ nữ 1971', 3, 7, 1971, 9, 30, 1); kho.luu(t.kq, 'Ví dụ nữ 1971', 'Khách', 'đã đối chiếu sự việc')
nhap_dl('Lá 1985', 15, 3, 1985, 10, 30); id3 = kho.luu(t.kq, 'Lá 1985', 'Khách', '')
w.t_ds.nap_nhom(); w.t_ds.nap(); w.di(10); chup('09_danh_sach_la_so.png')
w.t_ds.o_tim.setText('khách'); chup('10_danh_sach_tim_nhom.png'); w.t_ds.o_tim.setText('')
# 11. Mở lại lá 1 → kết quả như lúc lưu
w.t_ds.mo_la.emit(kho.la_so(id1), id1); app.processEvents()
from battu_core.kho_la_so import anh_chup
import json
luc_luu = json.loads(kho.lay(id1)['anh_chup']); bay_gio = anh_chup(t.kq)
print('Mở lại lá 1 – khối TÓM TẮT trùng lúc lưu:', luc_luu == bay_gio)
chup('11_mo_lai_la_1.png')
# 12. So sánh hai lá
w.t_ds.so_sanh.emit(id1, id3); app.processEvents(); chup('12_so_sanh_hai_la.png')
# 13. Nhập liệu (giữ màn hình nhập)
w.di(1); chup('13_nhap_lieu.png')
# 14. Hộp thoại Xuất / In
from ui.hop_xuat import HopXuat
w.di(0); app.processEvents(); h = HopXuat(w, t.kq); h.show(); app.processEvents(); h.grab().save(os.path.join(out, '14_hop_xuat_in.png')); print('14_hop_xuat_in.png'); h.close()
# 15. Chuyên đề – Lục thân (đỏ đậm Tốt/Khá), 16. Không rõ giờ, 17. Ngũ hành thiếu
w.di(7); w.t_cd.tab.setCurrentIndex(0); app.processEvents(); w.t_cd.bang['LUC_THAN'].scrollToItem(w.t_cd.bang['LUC_THAN'].item(26, 0), V.PositionAtTop); chup('15_chuyen_de_luc_than.png')
w.di(6); chup('16_ngu_hanh_thieu.png')
w.di(4); chup('17_phan_tich.png')
