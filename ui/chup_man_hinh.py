# -*- coding: utf-8 -*-
"""Chụp ảnh màn hình giao diện (chạy không cần màn hình: QT_QPA_PLATFORM=offscreen)."""
import os, sys
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from ui.giao_dien import tao_app, CuaSoChinh

out = sys.argv[1] if len(sys.argv) > 1 else 'anh'
os.makedirs(out, exist_ok=True)
app = tao_app([])
import tempfile
from battu_core.kho_la_so import Kho
w = CuaSoChinh(Kho(os.path.join(tempfile.mkdtemp(), 'thu.sqlite'))); w.show(); w.di(1)
t = w.t_nhap
def chup(ten):
    app.processEvents(); t.tinh(); app.processEvents(); w.grab().save(os.path.join(out, ten)); print(ten)
# 1. dương lịch – lá DL_26-11-1957 06:00 nam (lá trong dap_an_o)
t.ho_ten.setText('Lá mẫu 26/11/1957'); t.d_ngay.setValue(26); t.d_thang.setValue(11); t.d_nam.setValue(1957)
t.gio.setValue(6); t.phut.setValue(0); chup('01_nhap_duong_lich.png')
# 2. âm lịch – 10/5 nhuận/1971 (data39)
t.cach.dat(1); t.doi_cach(1); t.ho_ten.setText('Nhập âm lịch, tháng nhuận')
t.a_ngay.setValue(10); t.a_thang.setValue(5); t.a_nam.setValue(1971); t.a_nhuan.setChecked(True); t.gio.setValue(8)
t.gt.dat(1); chup('02_nhap_am_lich_nhuan.png')
# 3. tự nhập can chi – TVH-001
t.cach.dat(2); t.doi_cach(2); t.ho_ten.setText('TVH-001 (tự nhập can chi)'); t.gt.dat(0)
for cb, x in zip(t.cb_can, ['Ất', 'Ất', 'Kỷ', 'Nhâm']): cb.setCurrentText(x)
for cb, x in zip(t.cb_chi, ['Tị', 'Dậu', 'Mùi', 'Thân']): cb.setCurrentText(x)
t.nut_nc.setChecked(True); chup('03_nhap_can_chi_nang_cao.png')
# 4. thẻ trung gian
t.nut_nc.setChecked(False); t.cach.dat(0); t.doi_cach(0); t.ho_ten.setText('Lá mẫu 26/11/1957'); t.gio.setValue(6); t.gt.dat(0)
app.processEvents(); t.tinh(); w.di(3); app.processEvents()
w.grab().save(os.path.join(out, '04_the_trung_gian.png')); print('04_the_trung_gian.png')
# 5–6. trang Phân tích – lá TVH-001 (tự nhập can chi)
t.cach.dat(2); t.doi_cach(2); t.ho_ten.setText('TVH-001')
for cb, x in zip(t.cb_can, ['Ất', 'Ất', 'Kỷ', 'Nhâm']): cb.setCurrentText(x)
for cb, x in zip(t.cb_chi, ['Tị', 'Dậu', 'Mùi', 'Thân']): cb.setCurrentText(x)
app.processEvents(); t.tinh(); w.di(4)
w.t_pt.tab.setCurrentIndex(3); app.processEvents()
w.grab().save(os.path.join(out, '05_phan_tich_doi_chieu.png')); print('05_phan_tich_doi_chieu.png')
w.t_pt.tab.setCurrentIndex(4); app.processEvents(); w.t_pt.bang['DUNG_HY_KY'].scrollToItem(w.t_pt.bang['DUNG_HY_KY'].item(152, 1)); app.processEvents()
w.grab().save(os.path.join(out, '06_phan_tich_dung_than_khoi_K.png')); print('06_phan_tich_dung_than_khoi_K.png')
# 7–9. GĐ3: Ngũ hành thiếu (lá rows34 #1: Bính Dần – Canh Dần – Giáp Thìn – Tân Mùi), PA1 và PA2; Vận trình – lưu niên
t.cach.dat(2); t.doi_cach(2); t.ho_ten.setText('rows34 #1 (1/3/1986)')
for cb, x in zip(t.cb_can, ['Bính', 'Canh', 'Giáp', 'Tân']): cb.setCurrentText(x)
for cb, x in zip(t.cb_chi, ['Dần', 'Dần', 'Thìn', 'Mùi']): cb.setCurrentText(x)
t.nht.setCurrentText('PA1 – Theo trang'); app.processEvents(); t.tinh()
w.di(6); app.processEvents()
w.grab().save(os.path.join(out, '07_ngu_hanh_thieu_PA1.png')); print('07_ngu_hanh_thieu_PA1.png')
t.nht.setCurrentText('PA2 – Theo sách'); app.processEvents(); t.tinh(); app.processEvents()
w.grab().save(os.path.join(out, '08_ngu_hanh_thieu_PA2.png')); print('08_ngu_hanh_thieu_PA2.png')
t.cach.dat(0); t.doi_cach(0); t.ho_ten.setText('Lá mẫu 26/11/1957'); t.d_ngay.setValue(26); t.d_thang.setValue(11); t.d_nam.setValue(1957); t.gio.setValue(6)
app.processEvents(); t.tinh(); w.di(5); w.t_vt.tab.setCurrentIndex(1); app.processEvents()
w.t_vt.bang['LUU_NIEN'].scrollToItem(w.t_vt.bang['LUU_NIEN'].item(29, 0)); app.processEvents()
w.grab().save(os.path.join(out, '09_van_trinh_luu_nien.png')); print('09_van_trinh_luu_nien.png')
w.t_vt.tab.setCurrentIndex(0); app.processEvents(); w.grab().save(os.path.join(out, '10_van_trinh_dai_van.png')); print('10_van_trinh_dai_van.png')
# 11–13. GĐ4: Chuyên đề (Hôn nhân, Sức khỏe) và Không rõ giờ – lá TVH-001
t.cach.dat(2); t.doi_cach(2); t.ho_ten.setText('TVH-001')
for cb, x in zip(t.cb_can, ['Ất', 'Ất', 'Kỷ', 'Nhâm']): cb.setCurrentText(x)
for cb, x in zip(t.cb_chi, ['Tị', 'Dậu', 'Mùi', 'Thân']): cb.setCurrentText(x)
app.processEvents(); t.tinh(); w.di(7); w.t_cd.tab.setCurrentIndex(1); app.processEvents()
w.grab().save(os.path.join(out, '11_chuyen_de_hon_nhan.png')); print('11_chuyen_de_hon_nhan.png')
w.t_cd.tab.setCurrentIndex(6); app.processEvents(); w.grab().save(os.path.join(out, '12_chuyen_de_suc_khoe.png')); print('12_chuyen_de_suc_khoe.png')
t.krg.setChecked(True); app.processEvents(); t.tinh(); w.di(8); app.processEvents()
from PySide6.QtWidgets import QAbstractItemView as _V; w.t_krg.bang.scrollToItem(w.t_krg.bang.item(24, 1), _V.PositionAtTop); app.processEvents()
w.grab().save(os.path.join(out, '13_khong_ro_gio_12_gio.png')); print('13_khong_ro_gio_12_gio.png')
