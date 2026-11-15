"""Mẫu thử VS-2 (duyệt 15/11/2026): xét vượng suy PP1 ĐỦ 5 BƯỚC, có lý giải từng bước.
Đọc kết quả các bước 1–4 từ PP1_DINH_TINH (công cụ hiện có), thêm bước 5 kiểm chứng.
Căn cứ: P1 tr.70–72. Hợp không hóa ở PP1 theo P1 tr.62 (đứng độc lập)."""

TC = ['được lệnh', 'đắc địa', 'được sinh', 'được trợ']


def xet(E):
    g = lambda a: E.gia_tri('PP1_DINH_TINH', a)
    co = [g('C%d' % r) == 1 for r in (62, 63, 64, 65)]
    manh = [g('D%d' % r) == 1 for r in (62, 63, 64, 65)]
    so_bo = 'V' if 'vượng' in str(g('C91')).lower() else 'N'
    qt = str(g('C92'))
    a, b, c, d, tong = (g('C%d' % r) or 0 for r in (71, 72, 84, 87, 88))
    minh, khac = g('C89') or 0, g('D89') or 0
    da_phan_xu = 'a–d' in qt or 'phân xử' in qt
    B = []
    # Bước 1
    ds = ', '.join(f"{t} ({'mạnh' if m else 'yếu'})" for t, x, m in zip(TC, co, manh) if x) or 'không có tiêu chí nào'
    B.append(('1. Bốn tiêu chí', f'Nhật can có: {ds}.', 'P1 tr.70: được lệnh, đắc địa, được sinh, được trợ'))
    # Bước 2
    B.append(('2. Kết luận sơ bộ', f"{'Thân vượng' if so_bo == 'V' else 'Thân nhược'} – {qt}.",
              'P1 tr.70: "can ngày được lệnh, đó là điều quan trọng nhất để phán đoán thân vượng"'))
    # Bước 3
    dau = lambda x: '+' if x > 0 else ('−' if x < 0 else '0')
    B.append(('3. Xét a–d', f'a) hợp hóa {dau(a)}; b) bán hợp {dau(b)}; c) vượng địa {dau(c)}; d) phe khắc bị xung/chế/hợp/xa {dau(d)}; '
              f"tổng {tong:+.0f} → {'nghiêng về vượng' if tong > 0 else ('nghiêng về nhược' if tong < 0 else 'cân bằng')}.",
              'P1 tr.71 a)–đ)'))
    # Bước 4
    B.append(('4. So sánh mức độ', f"Phe sinh–trợ {minh:.0f} chữ, phe khắc–tiết–hao {khac:.0f} chữ → "
              f"{'phe mình nhiều hơn' if minh > khac else ('phe khác nhiều hơn' if khac > minh else 'ngang nhau')}.",
              'P1 tr.71: "phải so sánh mức độ giữa được sinh và bị khắc"'))
    # Bước 5
    ket, co_canh_bao = so_bo, False
    if da_phan_xu:
        ly = 'Kết luận sơ bộ đã dựa trên phân xử a–d (nửa vượng nửa nhược) nên giữ nguyên.'
    elif not co[0]:
        if so_bo == 'V' and tong < 0 and khac > minh:
            ket = 'N'; ly = 'Không được lệnh; các điều trợ không "mạnh, giúp ích nhiều": cả bước 3 và bước 4 đều nghiêng về phe khắc → đổi thành THÂN NHƯỢC.'
        elif so_bo == 'N' and tong > 0 and minh > khac:
            ket = 'V'; ly = 'Không được lệnh nhưng cả bước 3 và bước 4 đều nghiêng về phe sinh–trợ → đổi thành THÂN VƯỢNG.'
        else:
            ly = 'Bước 3 và 4 không cùng phủ định kết luận sơ bộ → giữ nguyên.'
    else:
        if so_bo == 'V' and tong < 0 and khac > minh:
            co_canh_bao = True
            ly = ('Được lệnh nên giữ THÂN VƯỢNG (lệnh là quan trọng nhất), nhưng bước 3 và 4 đều nghiêng về phe khắc → '
                  'ghi "vượng nhưng bị khắc–tiết nặng (cận ranh giới)": gặp vận khắc–tiết dễ chuyển nhược.')
        else:
            ly = 'Được lệnh, bước 3 và 4 không cùng phủ định → giữ nguyên.'
    B.append(('5. Kiểm chứng', ly, 'P1 tr.70 quy tắc 3 ("mạnh, giúp ích nhiều"), tr.71–72 (phán đoán tổng hợp)'))
    return dict(so_bo=so_bo, ket=ket, can_ranh_gioi=co_canh_bao, buoc=B)


def in_ra(kq):
    for t, ly, cau in kq['buoc']:
        print(f'  {t}: {ly}  [{cau}]')
    print('  KẾT LUẬN:', 'THÂN VƯỢNG' if kq['ket'] == 'V' else 'THÂN NHƯỢC', '(cận ranh giới)' if kq['can_ranh_gioi'] else '')


def trang_thai_cuoi(E):
    """Kết luận vượng suy cuối (G11 + VS-2, duyệt 15/11/2026): nếu VS-2 đổi kết luận PP1 ở bước 5 thì
    coi như đã phân xử (G11 mục 2) → theo PP1; ngược lại giữ khuyến nghị DOI_CHIEU C34."""
    k = xet(E); cu = str(E.gia_tri('DOI_CHIEU', 'C34'))
    if k['ket'] != k['so_bo']:
        moi = 'NHƯỢC (Theo PP1 – VS-2)' if k['ket'] == 'N' else 'VƯỢNG (Theo PP1 – VS-2)'
    else:
        moi = cu
    return moi, k
