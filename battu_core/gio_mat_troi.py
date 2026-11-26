# -*- coding: utf-8 -*-
"""Quy đổi GIỜ ĐỒNG HỒ nơi sinh → GIỜ MẶT TRỜI THẬT (tùy chọn, mặc định tắt – người dùng duyệt 26/11/2026).

Giờ mặt trời thật = giờ đồng hồ + 4 phút × (kinh độ nơi sinh − kinh tuyến chuẩn của múi giờ) + phương trình thời gian.
Kinh tuyến chuẩn = 15° × múi giờ (UTC+7 → 105°Đ). Phương trình thời gian (NOAA, sai số dưới 1 phút) lấy theo ngày trong năm.
Chỉ đổi ở lớp nhập liệu: lõi tính (dịch từ workbook) vẫn nhận ngày giờ như người dùng gõ vào ô; không đổi bộ đáp án cũ.
"""
import datetime, math

# Kinh độ gần đúng (độ Đông) của tỉnh lỵ / thành phố – chỉ dùng để chỉnh giờ, sai số < 0,1° ≈ 0,4 phút.
KINH_DO = {
    'Hà Nội': 105.85, 'Hải Phòng': 106.68, 'TP. Hồ Chí Minh': 106.70, 'Đà Nẵng': 108.22, 'Cần Thơ': 105.78,
    'Huế (Thừa Thiên Huế)': 107.59, 'An Giang (Long Xuyên)': 105.44, 'Bà Rịa – Vũng Tàu': 107.08, 'Bắc Giang': 106.19,
    'Bắc Kạn': 105.84, 'Bạc Liêu': 105.72, 'Bắc Ninh': 106.07, 'Bến Tre': 106.38, 'Bình Định (Quy Nhơn)': 109.22,
    'Bình Dương (Thủ Dầu Một)': 106.65, 'Bình Phước (Đồng Xoài)': 106.89, 'Bình Thuận (Phan Thiết)': 108.10, 'Cà Mau': 105.15,
    'Cao Bằng': 106.25, 'Đắk Lắk (Buôn Ma Thuột)': 108.04, 'Đắk Nông (Gia Nghĩa)': 107.69, 'Điện Biên': 103.02,
    'Đồng Nai (Biên Hòa)': 106.82, 'Đồng Tháp (Cao Lãnh)': 105.63, 'Gia Lai (Pleiku)': 107.99, 'Hà Giang': 104.98,
    'Hà Nam (Phủ Lý)': 105.92, 'Hà Tĩnh': 105.90, 'Hải Dương': 106.33, 'Hậu Giang (Vị Thanh)': 105.47, 'Hòa Bình': 105.34,
    'Hưng Yên': 106.05, 'Khánh Hòa (Nha Trang)': 109.19, 'Kiên Giang (Rạch Giá)': 105.08, 'Kon Tum': 108.00, 'Lai Châu': 103.46,
    'Lâm Đồng (Đà Lạt)': 108.44, 'Lạng Sơn': 106.76, 'Lào Cai': 103.97, 'Long An (Tân An)': 106.41, 'Nam Định': 106.18,
    'Nghệ An (Vinh)': 105.68, 'Ninh Bình': 105.97, 'Ninh Thuận (Phan Rang)': 108.99, 'Phú Thọ (Việt Trì)': 105.40,
    'Phú Yên (Tuy Hòa)': 109.30, 'Quảng Bình (Đồng Hới)': 106.62, 'Quảng Nam (Tam Kỳ)': 108.48, 'Quảng Ngãi': 108.80,
    'Quảng Ninh (Hạ Long)': 107.08, 'Quảng Trị (Đông Hà)': 107.10, 'Sóc Trăng': 105.97, 'Sơn La': 103.91, 'Tây Ninh': 106.10,
    'Thái Bình': 106.34, 'Thái Nguyên': 105.84, 'Thanh Hóa': 105.78, 'Tiền Giang (Mỹ Tho)': 106.36, 'Trà Vinh': 106.34,
    'Tuyên Quang': 105.22, 'Vĩnh Long': 105.97, 'Vĩnh Phúc (Vĩnh Yên)': 105.60, 'Yên Bái': 104.87,
}
NHAP_TAY = 'Khác (nhập kinh độ)'
DANH_SACH_NOI = sorted(KINH_DO) + [NHAP_TAY]


def phuong_trinh_thoi_gian(ngay: datetime.date, gio: float = 12.0) -> float:
    """Phương trình thời gian, đơn vị phút (dương: mặt trời thật đi trước đồng hồ trung bình)."""
    doy = ngay.timetuple().tm_yday
    nam = 366 if (ngay.year % 4 == 0 and (ngay.year % 100 or ngay.year % 400 == 0)) else 365
    g = 2 * math.pi / nam * (doy - 1 + (gio - 12) / 24)
    return 229.18 * (0.000075 + 0.001868 * math.cos(g) - 0.032077 * math.sin(g)
                     - 0.014615 * math.cos(2 * g) - 0.040849 * math.sin(2 * g))


def kinh_do_cua(noi_sinh: str, kinh_do=None):
    """Kinh độ dùng để tính: ô nhập tay (nếu có) ưu tiên, nếu không thì theo tên nơi sinh; không có → None."""
    if kinh_do not in (None, ''): return float(kinh_do)
    return KINH_DO.get(noi_sinh)


def quy_doi(ngay, thang, nam, gio, phut, kinh_do, mui_gio=7.0):
    """→ (datetime giờ mặt trời thật làm tròn đến phút, độ lệch phút so với đồng hồ).
    Độ lệch = 4 × (kinh độ − 15 × múi giờ) + phương trình thời gian (tính tại ngày giờ đồng hồ)."""
    t0 = datetime.datetime(nam, thang, ngay, int(gio or 0), int(phut or 0))
    lech = 4.0 * (float(kinh_do) - 15.0 * float(mui_gio)) + phuong_trinh_thoi_gian(t0.date(), t0.hour + t0.minute / 60)
    t1 = t0 + datetime.timedelta(minutes=round(lech))
    return t1, lech
