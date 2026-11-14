# BÁO CÁO KIỂM THỬ – GIAI ĐOẠN 2

Ngày chạy: 14/11/2026 00:17 · Nguồn: workbook SHA-256 `7ae759658d290172bbbf7b6642e24532239e3b9ea54ca71e2696802411bca4c5`

Phạm vi đã cài: CSDL (toàn sheet), NHAP_LIEU (toàn sheet), TRUNG_GIAN (toàn sheet), DAI_VAN (C5:C28), KIEM_THU (toàn sheet), PP1_DINH_TINH (toàn sheet), PP2_DINH_LUONG (toàn sheet), THONG_CAN (toàn sheet), DOI_CHIEU (toàn sheet), DUNG_HY_KY (toàn sheet)

## T1 – Đáp án từng ô (05_KIEM_THU/dap_an_o, 30 lá, 2.3 giây)

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| NHAP_LIEU | 1530 / 1530 | ĐẠT 100% |
| DOI_CHIEU | 3120 / 3120 | ĐẠT 100% |
| DUNG_HY_KY | 10950 / 10950 | ĐẠT 100% |
| PP1_DINH_TINH | 7380 / 7380 | ĐẠT 100% |
| PP2_DINH_LUONG | 8040 / 8040 | ĐẠT 100% |
| THONG_CAN | 8670 / 8670 | ĐẠT 100% |
| DAI_VAN | 720 / 720 | ĐẠT 100% |
| TRUNG_GIAN | 6750 / 6750 | ĐẠT 100% |
| CSDL | 840 / 840 | ĐẠT 100% |
| KIEM_THU | 420 / 420 | ĐẠT 100% |
| **Tổng** | **48420 / 48420** | ĐẠT |

Ô lệch: không có

## T3 – Lịch (05_KIEM_THU/lich, 710 thời điểm, 1.6 giây)

| Hạng mục | Khớp / tổng |
|---|---|
| Tứ trụ NHAP_LIEU C37:F38 + cảnh báo DUNG_HY_KY!C180 (9 ô) | 710 / 710 |

## Âm lịch (battu_core/lich_am.py)

| Hạng mục | Khớp / tổng |
|---|---|
| Dương → âm trùng lunar.py gốc, mọi ngày 1899–2101 | 74144 / 74144 |
| Khứ hồi âm → dương → âm, mọi ngày | 74144 / 74144 |
| data39.json dương → âm | 39 / 39 |
| data39.json âm → dương | 39 / 39 |

## T1b – Nhánh tùy chọn (MT-05) và T1-KRG (MT-04) – đáp án lấy bằng tools/oracle_mo_rong.py

Kiểm chính oracle mở rộng: biến thể KRG0 (mặc định) trùng dap_an_o **856410 / 856410** ô.

| Biến thể | Mô tả | Số lá | Ô khớp / tổng (sheet đã cài) | Kết quả |
|---|---|---|---|---|
| KRG0 | Mặc định (như dap_an_o) + vùng Bảng dữ liệu KHONG_RO_GIO C118:D285 (MT-04) | 30 | 48420 / 48420 | ĐẠT 100% |
| B1_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ | 30 | 47155 / 48420 | LỆCH 1265 – đáp án hỏng do Bảng dữ liệu LibreOffice (MT-13), thay bằng B1b |
| B2_THEO_PP1 | NHAP_LIEU C14 = Theo PP1 | 30 | 48420 / 48420 | ĐẠT 100% |
| B3_THEO_PP2 | NHAP_LIEU C14 = Theo PP2 | 30 | 48420 / 48420 | ĐẠT 100% |
| B4_PA1 | NGU_HANH_THIEU E3 = PA1 – Theo trang | 30 | 48420 / 48420 | ĐẠT 100% |
| B5_SU_KIEN | Nhập sự kiện mẫu: SU_KIEN B6:C10, KHONG_RO_GIO C11:D13 (năm theo DAI_VAN C22 của từng lá), THEO_DOI N/P dòng 12, 13, 20 | 30 | 48420 / 48420 | ĐẠT 100% |
| KRG_VONG | Mặc định + Bảng dữ liệu KHONG_RO_GIO tính bằng vòng 12 giờ (MT-13) | 30 | 48420 / 48420 | ĐẠT 100% |
| B1b_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ; Bảng dữ liệu tính bằng vòng 12 giờ (MT-13) – thay B1 | 30 | 48420 / 48420 | ĐẠT 100% |

## Các khối LỆNH 2 nêu riêng (đã nằm trong T1 và T1-501, liệt kê để dễ tra)

| Khối | Vùng | Ô công thức | T1 (30 lá) | T1-501 (501 lá) |
|---|---|---|---|---|
| Ô kiểm hội tụ A3 (G68) | PP2_DINH_LUONG!G68 | 1 | 30 / 30 | 501 / 501 |
| PP2 toàn sheet (A1–A4, hợp hóa 2 lượt, lục xung, kẹp xung, khử bì, A3 lặp) | PP2_DINH_LUONG!A1:G118 | 268 | 8040 / 8040 | 134268 / 134268 |
| Cách đặc biệt | DUNG_HY_KY!C37:D78 | 69 | 2070 / 2070 | 34569 / 34569 |
| Khối K (điều hầu, dự bị, thông quan, cứu ứng, bệnh – thuốc, hai phe, ranh giới, nhân cách, năm tốt/xấu) | DUNG_HY_KY!A153:P186 | 55 | 1650 / 1650 | 27555 / 27555 |
| Khuyến nghị theo NHAP_LIEU C14 | DOI_CHIEU!C28:D39 | 10 | 300 / 300 (+ T1b Theo PP1, Theo PP2: 600 / 600) | 5010 / 5010 |

Ví dụ G68 (TVH-001): "Kiểm hội tụ A3: ỔN ĐỊNH sau 3 vòng – mỗi lần khắc chỉ trừ một lần (P2 tr.339-340)".

## T2 – Kết quả chính 501 lá (bo_la + ket_qua_6.3, 13.2 giây)

(a) Khóa của scripts/uno_run59.py thuộc sheet đã cài (đọc như uno_run59: getString theo định dạng ô):

| Khóa | Ô | Khớp / tổng |
|---|---|---|
| p1 | PP1_DINH_TINH!C91 | 501 / 501 |
| p2 | PP2_DINH_LUONG!C103 | 501 / 501 |
| dung1 | DUNG_HY_KY!C106 | 501 / 501 |
| dung2 | DUNG_HY_KY!D106 | 501 / 501 |
| kn | DOI_CHIEU!C34 | 501 / 501 |
| tin | DOI_CHIEU!C35 | 501 / 501 |
| kn_d | DOI_CHIEU!D33 | 501 / 501 |
| d163 | DUNG_HY_KY!C163 | 501 / 501 |
| d170 | DUNG_HY_KY!C170 | 501 / 501 |
| d171 | DUNG_HY_KY!C171 | 501 / 501 |
| d174 | DUNG_HY_KY!C174 | 501 / 501 |
| d175 | DUNG_HY_KY!C175 | 501 / 501 |
| d178 | DUNG_HY_KY!C178 | 501 / 501 |
| d179 | DUNG_HY_KY!C179 | 501 / 501 |
| d180 | DUNG_HY_KY!C180 | 501 / 501 |
| d181 | DUNG_HY_KY!C181 | 501 / 501 |
| d182 | DUNG_HY_KY!C182 | 501 / 501 |
| d183 | DUNG_HY_KY!C183 | 501 / 501 |
| d184 | DUNG_HY_KY!C184 | 501 / 501 |

Khóa thuộc giai đoạn sau (so khi cài sheet tương ứng): nht1, nht_dl, nht_c7, k4, k5, k6, k7, k8, k9, k10, k11, k12, skJ6, skJ7, skJ8, dk, dl, dn, m16, c16, ai64, ag64, x64, r64, s64, c206, c72, c179, f115

(b) T1-501 – so TỪNG Ô các sheet đã cài trên 501 lá (đáp án tools/oracle_mo_rong.py, biến thể T2_501):

| Sheet | Ô khớp / tổng |
|---|---|
| NHAP_LIEU | 25551 / 25551 |
| DOI_CHIEU | 52104 / 52104 |
| DUNG_HY_KY | 182865 / 182865 |
| PP1_DINH_TINH | 123246 / 123246 |
| PP2_DINH_LUONG | 134268 / 134268 |
| THONG_CAN | 144789 / 144789 |
| DAI_VAN | 12024 / 12024 |
| TRUNG_GIAN | 112725 / 112725 |
| CSDL | 14028 / 14028 |
| KIEM_THU | 7014 / 7014 |
| **Tổng** | **808614 / 808614** |

## T4 – Hợp/xung (hop_xung, 510 lá gồm 9 ca biên, 2.2 giây)

So trùng từng giá trị scripts/run_hop.py đã ghi (đọc như run_hop: getString/getValue). Gói hồ sơ không ghi văn bản "6 bất biến"; các giá trị so ở đây là dữ liệu mà các bất biến được tính trên đó.

| Trường | Ô | Khớp / tổng |
|---|---|---|
| cuc3 | TRUNG_GIAN dòng 52 | 510 / 510 |
| banhop | TRUNG_GIAN dòng 59 | 510 / 510 |
| trongcuc | TRUNG_GIAN dòng 60 | 510 / 510 |
| luchop | TRUNG_GIAN dòng 69 | 510 / 510 |
| tranh | TRUNG_GIAN dòng 71 | 510 / 510 |
| loai | TRUNG_GIAN dòng 62 | 510 / 510 |
| hoa | PP2_DINH_LUONG dòng 24 | 510 / 510 |
| xung | PP2_DINH_LUONG dòng 31 | 510 / 510 |
| khubi | PP2_DINH_LUONG dòng 33 | 510 / 510 |
| d_hoa | PP2_DINH_LUONG dòng 76 | 510 / 510 |
| d_ban | PP2_DINH_LUONG dòng 83 | 510 / 510 |
| d_t2 | PP2_DINH_LUONG dòng 85 | 510 / 510 |
| d_t3 | PP2_DINH_LUONG dòng 87 | 510 / 510 |
| tonthat | PP2_DINH_LUONG dòng 82 | 510 / 510 |
| can | PP2_DINH_LUONG dòng 72 | 510 / 510 |
| tg1 | TRUNG_GIAN dòng 15 | 510 / 510 |
| tg2 | TRUNG_GIAN dòng 17 | 510 / 510 |
| tg3 | TRUNG_GIAN dòng 19 | 510 / 510 |
| pp1 | PP1_DINH_TINH dòng 96 | 510 / 510 |
| pp1tot | tổng phụ | 510 / 510 |
| tot | tổng phụ | 510 / 510 |
| sumF | tổng phụ | 510 / 510 |
| parts | tổng phụ | 510 / 510 |
| canF | tổng phụ | 510 / 510 |

Lá khớp đủ mọi trường: **510 / 510**

