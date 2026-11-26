# -*- coding: utf-8 -*-
"""Mô hình dữ liệu nhập của MỘT lá số và cách đặt vào các ô nhập của workbook (NHAP_LIEU …).

Phần này không chứa quy tắc Bát tự: chỉ chuyển lựa chọn trên giao diện thành đúng các ô mà người dùng
Excel sẽ gõ. Mọi tính toán vẫn do các module sheet (dịch từ công thức workbook) thực hiện.

Ba cách nhập (YEU_CAU 5.1):
  * 'duong' – ngày giờ dương lịch  → C12 = "Tự động", C26:G26 = ngày | tháng | năm | giờ | phút
  * 'am'    – ngày giờ âm lịch     → đổi sang dương lịch (battu_core.lich_am, khớp lunar.py) rồi như 'duong';
                                     tháng nhuận phải chọn rõ
  * 'canchi'– tự nhập 4 trụ        → C12 = "Không", C7:F8 = can chi, C26:G26 để trống
"""
import datetime
from dataclasses import dataclass, field, asdict
from battu_core import lich_am

CAN = ['Giáp', 'Ất', 'Bính', 'Đinh', 'Mậu', 'Kỷ', 'Canh', 'Tân', 'Nhâm', 'Quý']
CHI = ['Tí', 'Sửu', 'Dần', 'Mão', 'Thìn', 'Tị', 'Ngọ', 'Mùi', 'Thân', 'Dậu', 'Tuất', 'Hợi']
VUONG_SUY = ['Theo khuyến nghị', 'Theo PP1', 'Theo PP2']          # NHAP_LIEU C14
CHE_DO_NHT = ['PA2 – Theo sách', 'PA1 – Theo trang']              # NGU_HANH_THIEU E3 (G141: không dùng tên PA3)


class LoiNhap(ValueError):
    pass


SO_DONG_SU_KIEN = 30
SO_VIEC_TIM_GIO = 12                 # KHONG_RO_GIO C11:E22
SO_NAM_THEO_DOI = 100                # THEO_DOI N/O/P dòng 12–111


@dataclass
class LaSo:
    ho_ten: str = ''
    ghi_chu: str = ''
    cach_nhap: str = 'duong'            # 'duong' | 'am' | 'canchi'
    # dương lịch
    ngay: int = None; thang: int = None; nam: int = None; gio: int = None; phut: int = None
    # âm lịch (khi cach_nhap = 'am'); giờ phút dùng chung gio/phut
    am_ngay: int = None; am_thang: int = None; am_nam: int = None; am_nhuan: bool = False
    # can chi tự nhập (năm, tháng, ngày, giờ)
    can: list = field(default_factory=lambda: ['', '', '', ''])
    chi: list = field(default_factory=lambda: ['', '', '', ''])
    gioi_tinh: str = 'Nam'               # C10
    khong_ro_gio: bool = False           # C9
    gio_dang_xem: str = 'Ngọ'            # KHONG_RO_GIO C6 (giờ thử khi không rõ giờ)
    mui_gio: float = 7.0                 # C27
    so_ngay_toi_tiet_lap: float = None   # C11 (tháng Thìn/Mùi/Tuất/Sửu – nhập tay được ưu tiên)
    so_ngay_den_tiet: float = None       # C28 (chỉ khi không có ngày giờ sinh)
    so_gio_le: float = None              # C29
    nam_sinh_tru_nam: int = None         # C30 (khi tự nhập can chi)
    nam_xem: int = None                  # C31 (mặc định: năm hiện tại của máy – MT-06)
    thang_am: int = None                 # C32 (tùy chọn – phương hướng kết hôn)
    ngay_am: int = None                  # C33 (tùy chọn)
    phoi_ngau: list = field(default_factory=lambda: [None, None, None, None, None])  # C34:G34
    vuong_suy: str = 'Theo khuyến nghị'  # C14
    la_mau: str = 'Tự nhập'              # C13 (lá số mẫu trong KIEM_THU)
    che_do_nht: str = 'PA2 – Theo sách'  # NGU_HANH_THIEU E3
    su_kien: list = field(default_factory=list)  # SU_KIEN B6:D35 – [[năm, loại sự việc, ghi chú], …] tối đa 30 dòng
    viec_tim_gio: list = field(default_factory=list)  # KHONG_RO_GIO C11:E22 – [[năm, loại việc, ghi chú], …] tối đa 12
    theo_doi: dict = field(default_factory=dict)      # THEO_DOI N/O/P – {năm: [thực tế, sự kiện, đánh giá]}
    # Lớp tổng hợp (đề xuất hoàn thiện 15/11/2026, B4) – KHÔNG phải ô nhập workbook, không đổi kết quả tính:
    tinh_trang_hon_nhan: str = ''        # '' | 'Độc thân' | 'Đã kết hôn' | 'Ly hôn' | 'Góa'
    so_con: int = None                   # None = không nhập
    # Giờ mặt trời thật (tùy chọn, mặc định tắt – duyệt 26/11/2026): quy đổi giờ đồng hồ → giờ mặt trời thật trước khi đặt vào ô nhập
    gio_mat_troi: bool = False
    noi_sinh: str = ''                   # tên tỉnh/thành trong battu_core.gio_mat_troi.KINH_DO, hoặc 'Khác (nhập kinh độ)'
    kinh_do: float = None                # kinh độ Đông (độ); nếu có thì ưu tiên hơn tên nơi sinh
    phoi_tu_tru: str = ''                # mục 2.9 (bản 6.9): tứ trụ vợ/chồng nhập theo can chi "Can Chi, Can Chi, Can Chi, Can Chi" (năm, tháng, ngày, giờ); chỉ dùng cho dòng tham khảo con cái

    def to_dict(self): return asdict(self)

    @classmethod
    def from_dict(cls, d):
        x = cls(); [setattr(x, k, v) for k, v in d.items() if hasattr(x, k)]
        x.theo_doi = {int(k): list(v) for k, v in (x.theo_doi or {}).items()}   # JSON đổi khóa năm thành chuỗi
        return x

    # ------------------------------------------------------------------
    def ngay_duong(self):
        """(ngày, tháng, năm) dương lịch dùng để tính (đã đổi nếu nhập âm lịch)."""
        if self.cach_nhap == 'am':
            if None in (self.am_ngay, self.am_thang, self.am_nam):
                raise LoiNhap('Chưa nhập đủ ngày/tháng/năm âm lịch.')
            if not 1899 <= self.am_nam <= 2101:
                raise LoiNhap('Năm âm lịch ngoài phạm vi 1899–2101.')
            nh = lich_am.thang_nhuan(self.am_nam)
            if self.am_nhuan and nh != self.am_thang:
                raise LoiNhap('Năm âm %d %s.' % (self.am_nam, 'không có tháng nhuận' if nh == 0 else
                                                  'chỉ nhuận tháng %d' % nh))
            kq = lich_am.l2s(self.am_ngay, self.am_thang, self.am_nam, 1 if self.am_nhuan else 0)
            if kq is None:
                raise LoiNhap('Ngày âm %d/%d%s/%d không có (tháng này chỉ có %d ngày).' % (
                    self.am_ngay, self.am_thang, ' nhuận' if self.am_nhuan else '', self.am_nam,
                    lich_am.so_ngay_thang_am(self.am_thang, self.am_nam, 1 if self.am_nhuan else 0)))
            return kq
        if self.cach_nhap == 'duong':
            if None in (self.ngay, self.thang, self.nam):
                raise LoiNhap('Chưa nhập đủ ngày/tháng/năm dương lịch.')
            return self.ngay, self.thang, self.nam
        return None

    def gio_quy_doi(self):
        """Khi bật giờ mặt trời thật (và biết ngày giờ): → (ngày, tháng, năm, giờ, phút, độ lệch phút); ngược lại None."""
        if not self.gio_mat_troi or self.khong_ro_gio or self.cach_nhap == 'canchi' or self.gio is None:
            return None
        from battu_core import gio_mat_troi as g
        kd = g.kinh_do_cua(self.noi_sinh, self.kinh_do)
        if kd is None:
            raise LoiNhap('Chọn nơi sinh hoặc nhập kinh độ để quy đổi giờ mặt trời thật.')
        if not -180 <= kd <= 180:
            raise LoiNhap('Kinh độ phải nằm trong khoảng −180 đến 180.')
        d, m, y = self.ngay_duong()
        try:
            t, lech = g.quy_doi(d, m, y, self.gio, self.phut or 0, kd, self.mui_gio if self.mui_gio is not None else 7.0)
        except ValueError:
            raise LoiNhap('Ngày dương lịch không hợp lệ.')
        return t.day, t.month, t.year, t.hour, t.minute, lech

    def o_nhap(self, nam_theo_doi=None):
        """→ {sheet: {ô: giá trị}} đúng như người dùng nhập vào workbook (None = để trống).
        nam_theo_doi: danh sách năm ở THEO_DOI A12:A111 (do lõi tính ra) để đặt đánh giá thực tế đúng dòng năm."""
        nl = {}
        nl['C9'] = 'Không rõ giờ' if self.khong_ro_gio else 'Biết giờ'
        nl['C10'] = self.gioi_tinh
        nl['C11'] = self.so_ngay_toi_tiet_lap
        nl['C13'] = self.la_mau or 'Tự nhập'
        nl['C14'] = self.vuong_suy
        nl['C27'] = float(self.mui_gio) if self.mui_gio is not None else None
        nl['C28'] = self.so_ngay_den_tiet
        nl['C29'] = self.so_gio_le
        nl['C30'] = float(self.nam_sinh_tru_nam) if self.nam_sinh_tru_nam else None
        nl['C31'] = float(self.nam_xem if self.nam_xem else datetime.date.today().year)
        # MT-12 (duyệt 14/11/2026): nhập âm lịch → tự điền C32/C33 bằng chính ngày âm đã nhập (nếu người dùng chưa
        # nhập riêng); nhập dương lịch → để trống trừ khi người dùng bấm "Điền theo âm lịch quy đổi".
        thang_am, ngay_am = self.thang_am, self.ngay_am
        if self.cach_nhap == 'am':
            thang_am = thang_am or self.am_thang; ngay_am = ngay_am or self.am_ngay
        nl['C32'] = float(thang_am) if thang_am else None
        nl['C33'] = float(ngay_am) if ngay_am else None
        for col, v in zip('CDEFG', self.phoi_ngau): nl[col + '34'] = float(v) if v not in (None, '') else None
        if self.cach_nhap in ('duong', 'am'):
            d, m, y = self.ngay_duong()
            gio, phut = self.gio, self.phut
            q = self.gio_quy_doi()
            if q: d, m, y, gio, phut = q[:5]
            nl['C12'] = 'Tự động'
            nl['C26'] = float(d); nl['D26'] = float(m); nl['E26'] = float(y)
            nl['F26'] = float(gio) if gio is not None else None
            nl['G26'] = float(phut) if phut is not None else None
            for i, col in enumerate('CDEF'):
                nl[col + '7'] = self.can[i] or None; nl[col + '8'] = self.chi[i] or None
        else:
            nl['C12'] = 'Không'
            for col in 'CDEFG': nl[col + '26'] = None
            for i, col in enumerate('CDEF'):
                nl[col + '7'] = self.can[i] or None; nl[col + '8'] = self.chi[i] or None
        sk = {}
        for i, dong in enumerate((self.su_kien or [])[:SO_DONG_SU_KIEN]):      # SU_KIEN B6:D35 (ô nhập của workbook)
            nam, loai, ghi = (list(dong) + [None, None, None])[:3]
            r = 6 + i
            sk['B%d' % r] = float(nam) if nam not in (None, '') else None
            sk['C%d' % r] = loai or None
            sk['D%d' % r] = ghi or None
        krg = {'C6': self.gio_dang_xem}
        for i, dong in enumerate((self.viec_tim_gio or [])[:SO_VIEC_TIM_GIO]):   # KHONG_RO_GIO C11:E22
            nam, loai, ghi = (list(dong) + [None, None, None])[:3]
            r = 11 + i
            krg['C%d' % r] = float(nam) if nam not in (None, '') else None
            krg['D%d' % r] = loai or None
            krg['E%d' % r] = ghi or None
        td = {}
        if self.theo_doi and nam_theo_doi:
            for i, nam in enumerate(list(nam_theo_doi)[:SO_NAM_THEO_DOI]):
                if not isinstance(nam, (int, float)) or isinstance(nam, bool): continue
                g = self.theo_doi.get(int(nam))
                if not g: continue
                tt, sk_, dg = (list(g) + [None, None, None])[:3]
                r = 12 + i
                td['N%d' % r] = tt or None; td['O%d' % r] = sk_ or None; td['P%d' % r] = dg or None
        ra = {'NHAP_LIEU': nl,
              'KHONG_RO_GIO': krg,
              'NGU_HANH_THIEU': {'E3': self.che_do_nht}}
        if sk: ra['SU_KIEN'] = sk
        if td: ra['THEO_DOI'] = td
        return ra
