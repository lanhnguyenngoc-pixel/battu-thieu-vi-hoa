# -*- coding: utf-8 -*-
"""Chạy toàn bộ kiểm thử của giai đoạn hiện tại và ghi BAO_CAO_KIEM_THU.md.
Dùng: python -m tests.chay_tat_ca <thư_mục_HO_SO_PM> [tệp_báo_cáo.md]"""
import sys, os, json, time, datetime, collections
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from tests import so_dap_an, t3_lich, t_am_lich, t1b, t2_501, t4_hop_xung, t8
from battu_core import sheets as S

HS = sys.argv[1]; RA = sys.argv[2] if len(sys.argv) > 2 else 'BAO_CAO_KIEM_THU.md'
KT = os.path.join(HS, '05_KIEM_THU'); MR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dap_an_mo_rong')
GD = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'tools', 'pham_vi_hien_tai.json')))['giai_doan']
L = ['# BÁO CÁO KIỂM THỬ TỔNG – GIAI ĐOẠN %s' % GD, '',
     'Ngày chạy: %s · Nguồn: workbook SHA-256 `%s`' % (datetime.datetime.now().strftime('%d/%m/%Y %H:%M'), S.NGUON_SHA256), '',
     'Phạm vi đã cài: ' + ', '.join('%s (%s)' % (k, 'toàn sheet' if v == 'all' else ', '.join(v)) for k, v in S.PHAM_VI.items()), '']

tong, dung, lech, dt, n = so_dap_an.chay(KT)
L += ['## T1 – Đáp án từng ô (05_KIEM_THU/dap_an_o, %d lá, %.1f giây)' % (n, dt), '',
      '| Sheet | Ô khớp / tổng | Kết quả |', '|---|---|---|']
for sh in tong: L.append('| %s | %d / %d | %s |' % (sh, dung[sh], tong[sh], 'ĐẠT 100%' if dung[sh] == tong[sh] else 'LỆCH %d' % (tong[sh] - dung[sh])))
L.append('| **Tổng** | **%d / %d** | %s |' % (sum(dung.values()), sum(tong.values()), 'ĐẠT' if not lech else 'LỆCH'))
L += ['', 'Ô lệch: %s' % ('không có' if not lech else ''), '']
for x in lech[:200]: L.append('- %s %s!%s: phần mềm %r – đáp án %r' % x)
print('T1', sum(dung.values()), '/', sum(tong.values()))

d3, n3, l3, dt3 = t3_lich.chay(KT)
L += ['## T3 – Lịch (05_KIEM_THU/lich, %d thời điểm, %.1f giây)' % (n3, dt3), '',
      '| Hạng mục | Khớp / tổng |', '|---|---|',
      ('| Tứ trụ NHAP_LIEU C37:F38 + cảnh báo DUNG_HY_KY!C180 (9 ô) | %d / %d |' if t3_lich.CO_C180 else '| Tứ trụ NHAP_LIEU C37:F38 (8 ô) | %d / %d |') % (d3, n3), '']
for x in l3[:50]: L.append('- LỆCH %r' % (x,))
print('T3', d3, '/', n3)

na, a, b, m, c, e, loi, ghi, dta = t_am_lich.chay(os.path.join(HS, '09_TAI_LIEU_BO_SUNG', 'ngu_hanh_thieu'))
L += ['## Âm lịch (battu_core/lich_am.py)', '', '| Hạng mục | Khớp / tổng |', '|---|---|',
      '| Dương → âm trùng lunar.py gốc, mọi ngày 1899–2101 | %d / %d |' % (a, na),
      '| Khứ hồi âm → dương → âm, mọi ngày | %d / %d |' % (b, na),
      '| data39.json dương → âm | %d / %d |' % (c, m), '| data39.json âm → dương | %d / %d |' % (e, m), '']
for x in (loi + ghi)[:50]: L.append('- LỆCH %r' % (x,))
print('Âm lịch', a, b, c, e)

dk, nk = t1b.so_krg0_voi_dap_an_o(KT, MR)
L += ['## T1b – Nhánh tùy chọn (MT-05) và T1-KRG (MT-04) – đáp án lấy bằng tools/oracle_mo_rong.py', '',
      'Kiểm chính oracle mở rộng: biến thể KRG0 (mặc định) trùng dap_an_o **%d / %d** ô.' % (dk, nk), '',
      '| Biến thể | Mô tả | Số lá | Ô khớp / tổng (sheet đã cài) | Kết quả |', '|---|---|---|---|---|']
for ten, b_ in t1b.BT.items():
    if ten.startswith(('T2_', 'T8', 'T7_')) or not os.path.isdir(os.path.join(MR, ten)): continue
    n_, tg, dg, lg = t1b.chay_bien_the(KT, MR, ten)
    kq = 'ĐẠT 100%' if not lg else 'LỆCH %d' % len(lg)
    if ten == 'B1_KHONG_RO_GIO' and lg: kq += ' – đáp án hỏng do Bảng dữ liệu LibreOffice (MT-13), thay bằng B1b'
    L.append('| %s | %s | %d | %d / %d | %s |' % (ten, b_['mo_ta'], n_, sum(dg.values()), sum(tg.values()), kq))
    print(ten, sum(dg.values()), '/', sum(tg.values()))
L.append('')
if 'DUNG_HY_KY' in S.PHAM_VI:
    n2, khoa, tk, dk, lk, to, do, lo, thieu, dt2 = t2_501.chay(KT, MR)
    L += ['## T2 – Kết quả chính 501 lá (bo_la + ket_qua_6.3, %.1f giây)' % dt2, '',
          '(a) Khóa của scripts/uno_run59.py thuộc sheet đã cài (đọc như uno_run59: getString theo định dạng ô):', '',
          '| Khóa | Ô | Khớp / tổng |', '|---|---|---|']
    for k in tk: L.append('| %s | %s!%s | %d / %d |' % (k, khoa[k][0], khoa[k][1], dk[k], tk[k]))
    L += ['', 'Khóa thuộc giai đoạn sau (so khi cài sheet tương ứng): ' + ', '.join(k for k in t2_501.READ if k not in khoa), '',
          '(b) T1-501 – so TỪNG Ô các sheet đã cài trên 501 lá (đáp án tools/oracle_mo_rong.py, biến thể T2_501):', '',
          '| Sheet | Ô khớp / tổng |', '|---|---|']
    for sh in to: L.append('| %s | %d / %d |' % (sh, do[sh], to[sh]))
    L.append('| **Tổng** | **%d / %d** |' % (sum(do.values()), sum(to.values())))
    for x in (lk + lo)[:100]: L.append('- LỆCH %r' % (x,))
    L.append(''); print('T2', sum(dk.values()), '/', sum(tk.values()), '| T1-501', sum(do.values()), '/', sum(to.values()))
    n4, ld, t4t, t4k, l4, ne, dt4 = t4_hop_xung.chay(KT)
    L += ['## T4 – Hợp/xung (hop_xung, %d lá gồm %d ca biên, %.1f giây)' % (n4, ne, dt4), '',
          'So trùng từng giá trị scripts/run_hop.py đã ghi (đọc như run_hop: getString/getValue). Gói hồ sơ không ghi văn bản '
          '"6 bất biến"; các giá trị so ở đây là dữ liệu mà các bất biến được tính trên đó.', '',
          '| Trường | Ô | Khớp / tổng |', '|---|---|---|']
    for k in t4t:
        o = t4_hop_xung.RD.get(k); L.append('| %s | %s | %d / %d |' % (k, ('%s dòng %d' % o) if o else 'tổng phụ', t4k[k], t4t[k]))
    L += ['', 'Lá khớp đủ mọi trường: **%d / %d**' % (ld, n4), '']
    for x in l4[:50]: L.append('- LỆCH %r' % (x,))
    print('T4', ld, '/', n4)
if 'NGU_HANH_THIEU' in S.PHAM_VI:
    ra = t8.chay(KT, MR)
    L += ['## T8 – Ngũ hành thiếu, biểu đồ (bộ lá chuẩn tests/t8_chuan.json)', '',
          'Lá ngày giờ dương lịch (từ data39, dv16), tứ trụ theo lịch tiết khí. So từng ô mọi sheet đã cài với đáp án workbook '
          'ở hai chế độ E3.', '', '| Chế độ | Số lá | NGU_HANH_THIEU | Mọi sheet đã cài |', '|---|---|---|---|']
    for ten, (n_, tg, dg, lg) in ra.items():
        L.append('| %s | %d | %d / %d | %d / %d %s |' % (ten, n_, dg['NGU_HANH_THIEU'], tg['NGU_HANH_THIEU'], sum(dg.values()), sum(tg.values()),
                                                     'ĐẠT 100%' if not lg else 'LỆCH %d' % len(lg)))
        for x in lg[:30]: L.append('- LỆCH %r' % (x,))
        print(ten, sum(dg.values()), '/', sum(tg.values()))
    L.append('')
if 'KET_LUAN' in S.PHAM_VI:
    from tests import t_kho_la_so
    n6, d6, l6, k6, dt6 = t_kho_la_so.chay(12)
    L += ['## T6 – Kho lá số SQLite (YÊU CẦU 5.3)', '', 'Lưu 12 lá (t8_chuan.json, có lá Không rõ giờ) → mở lại → tính lại; '
          'so khối TÓM TẮT KET_LUAN B2:C12 + tứ trụ với ảnh chụp lúc lưu (%.1f giây).' % dt6, '', '| Hạng mục | Kết quả |', '|---|---|',
          '| Mở lại cho kết quả như lúc lưu | %d / %d |' % (d6, n6)] + ['| %s | %s |' % (k, 'ĐẠT' if v else 'LỖI') for k, v in k6.items()] + ['']
    print('T6', d6, '/', n6, k6)
if 'KET_LUAN' in S.PHAM_VI:
    from tests import t7_ngau_nhien, t5_xuat
    n7, t7, d7, l7, th7, dt7 = t7_ngau_nhien.chay()
    L += ['## T7 – 20 lá NGẪU NHIÊN MỚI (oracle_dump.py nguyên bản, không mượn ô nào)', '',
          'tests/t7_ngau_nhien.json (hạt giống 20261114): 14 lá dương lịch 1900–2100, 6 lá tự nhập can chi; %.1f giây.' % dt7, '',
          '| Sheet | Ô khớp / tổng | Kết quả |', '|---|---|---|']
    for sh in t7: L.append('| %s | %d / %d | %s |' % (sh, d7[sh], t7[sh], 'ĐẠT 100%' if d7[sh] == t7[sh] else 'LỆCH'))
    L += ['| **Tổng** | **%d / %d** | %s |' % (sum(d7.values()), sum(t7.values()), 'ĐẠT 100%' if not l7 else 'LỆCH %d' % len(l7)), '']
    for x in l7[:40]: L.append('- LỆCH %r' % (x,))
    print('T7', sum(d7.values()), '/', sum(t7.values()))
    tj, dj, tx, dx, l5, p5, dt5 = t5_xuat.chay(5)
    L += ['## T5 – Xuất PDF / XLSX / JSON (5 lá T7, %.1f giây)' % dt5, '', '| Định dạng | Kiểm | Kết quả |', '|---|---|---|',
          '| JSON | mọi ô công thức trùng đáp án oracle | %d / %d |' % (dj['JSON'], tj['JSON']),
          '| XLSX | mở lại, mọi ô công thức của 30 sheet hiển thị trùng JSON; giữ gộp ô, cột/dòng ẩn | %d / %d |' % (dx['XLSX'], tx['XLSX']),
          '| PDF | số trang, "Trang i/n" mỗi trang, khối KẾT LUẬN, phiên bản quy tắc | %d / %d lá |' % (sum(p5), len(p5)), '']
    for x in l5[:30]: L.append('- LỆCH %r' % (x,))
    print('T5', dj, dx, p5)
if 'SU_KIEN' in S.PHAM_VI:
    from tests import t9_su_kien
    n9, t9, d9, l9, (db9, tb9), ok9 = t9_su_kien.chay()
    L += ['## T9 – Nhập sự việc SU_KIEN (bản 1.1.0): 20 lá T7 × 6 sự việc, đáp án tools/oracle_mo_rong.py (T7_SU_KIEN)', '',
          '| Hạng mục | Kết quả |', '|---|---|',
          '| Mọi ô công thức 33 sheet | %d / %d |' % (sum(d9.values()), sum(t9.values())),
          '| Riêng SU_KIEN | %d / %d |' % (d9['SU_KIEN'], t9['SU_KIEN']),
          '| Đường nhập LaSo.su_kien → ô SU_KIEN B/C/D trùng ô oracle đã nhập | %d / %d lá |' % (db9, tb9),
          '| Lưu SQLite → mở lại giữ nguyên sự việc | %s |' % ('ĐẠT' if ok9 else 'LỖI'), '']
    for x in l9[:30]: L.append('- LỆCH %r' % (x,))
    print('T9', sum(d9.values()), '/', sum(t9.values()), db9, tb9, ok9)
open(RA, 'w').write('\n'.join(L) + '\n'); print('Ghi', RA)
