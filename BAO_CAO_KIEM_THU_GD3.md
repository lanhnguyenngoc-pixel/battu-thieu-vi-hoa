# BÁO CÁO KIỂM THỬ – GIAI ĐOẠN 3

Ngày chạy: 14/11/2026 06:47 · Nguồn: workbook SHA-256 `7ae759658d290172bbbf7b6642e24532239e3b9ea54ca71e2696802411bca4c5`

Phạm vi đã cài: CSDL (toàn sheet), NHAP_LIEU (toàn sheet), TRUNG_GIAN (toàn sheet), DAI_VAN (toàn sheet), KIEM_THU (toàn sheet), PP1_DINH_TINH (toàn sheet), PP2_DINH_LUONG (toàn sheet), THONG_CAN (toàn sheet), DOI_CHIEU (toàn sheet), DUNG_HY_KY (toàn sheet), LUU_NIEN (toàn sheet), THEO_DOI (toàn sheet), NGU_HANH_THIEU (toàn sheet)

## T1 – Đáp án từng ô (05_KIEM_THU/dap_an_o, 30 lá, 12.3 giây)

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| NHAP_LIEU | 1530 / 1530 | ĐẠT 100% |
| THEO_DOI | 150180 / 150180 | ĐẠT 100% |
| DOI_CHIEU | 3120 / 3120 | ĐẠT 100% |
| DUNG_HY_KY | 10950 / 10950 | ĐẠT 100% |
| PP1_DINH_TINH | 7380 / 7380 | ĐẠT 100% |
| PP2_DINH_LUONG | 8040 / 8040 | ĐẠT 100% |
| THONG_CAN | 8670 / 8670 | ĐẠT 100% |
| DAI_VAN | 19740 / 19740 | ĐẠT 100% |
| LUU_NIEN | 334440 / 334440 | ĐẠT 100% |
| TRUNG_GIAN | 6750 / 6750 | ĐẠT 100% |
| CSDL | 840 / 840 | ĐẠT 100% |
| KIEM_THU | 420 / 420 | ĐẠT 100% |
| NGU_HANH_THIEU | 7260 / 7260 | ĐẠT 100% |
| **Tổng** | **559320 / 559320** | ĐẠT |

Ô lệch: không có

## T3 – Lịch (05_KIEM_THU/lich, 710 thời điểm, 1.4 giây)

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
| KRG0 | Mặc định (như dap_an_o) + vùng Bảng dữ liệu KHONG_RO_GIO C118:D285 (MT-04) | 30 | 559320 / 559320 | ĐẠT 100% |
| B1_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ | 30 | 556809 / 559320 | LỆCH 2511 – đáp án hỏng do Bảng dữ liệu LibreOffice (MT-13), thay bằng B1b |
| B2_THEO_PP1 | NHAP_LIEU C14 = Theo PP1 | 30 | 559320 / 559320 | ĐẠT 100% |
| B3_THEO_PP2 | NHAP_LIEU C14 = Theo PP2 | 30 | 559320 / 559320 | ĐẠT 100% |
| B4_PA1 | NGU_HANH_THIEU E3 = PA1 – Theo trang | 30 | 559320 / 559320 | ĐẠT 100% |
| B5_SU_KIEN | Nhập sự kiện mẫu: SU_KIEN B6:C10, KHONG_RO_GIO C11:D13 (năm theo DAI_VAN C22 của từng lá), THEO_DOI N/P dòng 12, 13, 20 | 30 | 559320 / 559320 | ĐẠT 100% |
| KRG_VONG | Mặc định + Bảng dữ liệu KHONG_RO_GIO tính bằng vòng 12 giờ (MT-13) | 30 | 559320 / 559320 | ĐẠT 100% |
| B1b_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ; Bảng dữ liệu tính bằng vòng 12 giờ (MT-13) – thay B1 | 30 | 559320 / 559320 | ĐẠT 100% |

## T2 – Kết quả chính 501 lá (bo_la + ket_qua_6.3, 202.0 giây)

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
| nht1 | NGU_HANH_THIEU!E15 | 501 / 501 |
| nht_dl | NGU_HANH_THIEU!E7 | 501 / 501 |
| nht_c7 | NGU_HANH_THIEU!C7 | 501 / 501 |
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
| dk | LUU_NIEN!DK31 | 501 / 501 |
| dl | LUU_NIEN!DL31 | 501 / 501 |

Khóa thuộc giai đoạn sau (so khi cài sheet tương ứng): k4, k5, k6, k7, k8, k9, k10, k11, k12, skJ6, skJ7, skJ8, dn, m16, c16, ai64, ag64, x64, r64, s64, c206, c72, c179, f115

(b) T1-501 – so TỪNG Ô các sheet đã cài trên 501 lá (đáp án tools/oracle_mo_rong.py, biến thể T2_501):

| Sheet | Ô khớp / tổng |
|---|---|
| NHAP_LIEU | 25551 / 25551 |
| THEO_DOI | 2508006 / 2508006 |
| DOI_CHIEU | 52104 / 52104 |
| DUNG_HY_KY | 182865 / 182865 |
| PP1_DINH_TINH | 123246 / 123246 |
| PP2_DINH_LUONG | 134268 / 134268 |
| THONG_CAN | 144789 / 144789 |
| DAI_VAN | 329658 / 329658 |
| LUU_NIEN | 5585148 / 5585148 |
| TRUNG_GIAN | 112725 / 112725 |
| CSDL | 14028 / 14028 |
| KIEM_THU | 7014 / 7014 |
| NGU_HANH_THIEU | 121242 / 121242 |
| **Tổng** | **9340644 / 9340644** |

## T4 – Hợp/xung (hop_xung, 510 lá gồm 9 ca biên, 2.0 giây)

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

## T8 – Ngũ hành thiếu, biểu đồ (rows34, data39, dv16 – 89 lá)

Tiêu chí YEU_CAU: khớp như workbook 6.3; chênh với trang thì ghi MAU_THUAN. (A) so từng ô với đáp án workbook của chính 89 lá (tools/oracle_mo_rong.py, ô nhập tests/cases_t8.json); (B) đối chiếu workbook với số trang ghi.

| (A) Biến thể | Số lá | Ô khớp / tổng (mọi sheet đã cài) |
|---|---|---|
| T8_PA1 | 89 | 1659316 / 1659316 ĐẠT 100% |
| T8_PA2 | 89 | 1659316 / 1659316 ĐẠT 100% |
| T8_D39_TRU | 34 | 633896 / 633896 ĐẠT 100% |

| (B) Đối chiếu workbook – trang | workbook = trang |
|---|---|
| rows34 – mục I (E15, PA1) | 34 / 34 |
| rows34 – mục II (E27, PA1) | 34 / 34 |
| data39 – 6 năng lực (E48:E53) | 31 / 39 |
| data39 – mục I (E15, PA1) | 24 / 34 |
| data39 – mục II (E27, PA1) | 25 / 34 |
| data39 (tứ trụ trang) – 6 năng lực | 34 / 34 |
| data39 (tứ trụ trang) – mục I (PA1) | 34 / 34 |
| data39 (tứ trụ trang) – mục II (PA1) | 34 / 34 |
| dv16 – 8 điểm đại vận (E70:E77) | 13 / 16 |
| dv16 – t = mục II (E27, PA1) | 8 / 16 |
| dv16 – X = Thực tế − Đại vận | 11 / 16 |

Chỗ workbook khác trang (không phải lỗi phần mềm – phần mềm khớp workbook):

- data39 – mục I (E15, PA1) · 1/3/1986 · workbook 20% · trang 0% · giờ dùng: 2h00; giờ: lịch Ất Sửu → trang Tân Mùi
- data39 – mục II (E27, PA1) · 1/3/1986 · workbook 2.5 · trang 1.5 · giờ dùng: 2h00; giờ: lịch Ất Sửu → trang Tân Mùi
- data39 – mục I (E15, PA1) · 02/4/1972 · workbook 68% · trang 80% · giờ dùng: 11h30; giờ: lịch Mậu Ngọ → trang Canh Thân
- data39 – mục II (E27, PA1) · 02/4/1972 · workbook -1.25 · trang 0.75 · giờ dùng: 11h30; giờ: lịch Mậu Ngọ → trang Canh Thân
- data39 – mục II (E27, PA1) · 11/5/1963 · workbook -0.25 · trang 0.75 · giờ dùng: 8h00; giờ: lịch Mậu Thìn → trang Quý Dậu
- data39 – 6 năng lực (E48:E53) · 27/10/1992 · workbook [10, 5, 0, 40, 10, 90] · trang [80, 0, 5, 15, 5, 50] · giờ dùng: không ghi
- data39 – mục I (E15, PA1) · 27/10/1992 · workbook 0% · trang 48% · giờ dùng: không ghi; tháng: lịch Canh Tuất → trang Tân Hợi; không ghi giờ – can giờ suy từ bảng trang
- data39 – mục I (E15, PA1) · 9/5/1997 · workbook 68% · trang 48% · giờ dùng: 6h00; giờ: lịch Tân Mão → trang Quý Tị
- data39 – mục II (E27, PA1) · 9/5/1997 · workbook -2.25 · trang -3.25 · giờ dùng: 6h00; giờ: lịch Tân Mão → trang Quý Tị
- data39 – 6 năng lực (E48:E53) · 7/1/1979 · workbook [30, 45, 20, 15, 10, 90] · trang [90, 30, 5, 20, 30, 5] · giờ dùng: 15h30
- data39 – mục I (E15, PA1) · 7/1/1979 · workbook 32% · trang 60% · giờ dùng: 15h30; năm: lịch Mậu Ngọ → trang Canh Ngọ; tháng: lịch Ất Sửu → trang Kỷ Sửu; ngày: lịch Giáp Tuất → trang Nhâm Tý; giờ: lịch Nhâm Thân → trang Giáp Thìn; chi năm, chi ngày suy ra (nhiều khả năng, cùng kết quả)
- data39 – mục II (E27, PA1) · 7/1/1979 · workbook -1 · trang -2 · giờ dùng: 15h30; năm: lịch Mậu Ngọ → trang Canh Ngọ; tháng: lịch Ất Sửu → trang Kỷ Sửu; ngày: lịch Giáp Tuất → trang Nhâm Tý; giờ: lịch Nhâm Thân → trang Giáp Thìn; chi năm, chi ngày suy ra (nhiều khả năng, cùng kết quả)
- data39 – 6 năng lực (E48:E53) · 4/12/2005 · workbook [40, 50, 30, 50, 10, 20] · trang [80, 100, 10, 15, 50, 5] · giờ dùng: 6h00
- data39 – mục I (E15, PA1) · 4/12/2005 · workbook 20% · trang 32% · giờ dùng: 6h00; ngày: lịch Nhâm Tuất → trang Đinh Mùi; giờ: lịch Quý Mão → trang Kỷ Dậu; cả ngày suy từ bảng trang
- data39 – mục II (E27, PA1) · 4/12/2005 · workbook 1.75 · trang -0.25 · giờ dùng: 6h00; ngày: lịch Nhâm Tuất → trang Đinh Mùi; giờ: lịch Quý Mão → trang Kỷ Dậu; cả ngày suy từ bảng trang
- data39 – mục I (E15, PA1) · 5/12/1991 · workbook 60% · trang 48% · giờ dùng: 6h00; giờ: lịch Đinh Mão → trang Giáp Tý
- data39 – mục II (E27, PA1) · 5/12/1991 · workbook -1.25 · trang -3.25 · giờ dùng: 6h00; giờ: lịch Đinh Mão → trang Giáp Tý
- data39 – 6 năng lực (E48:E53) · 24/3/1993 · workbook [10, 50, 30, 50, 0, 50] · trang [5, 30, 5, 20, 5, 30] · giờ dùng: không ghi
- data39 – mục I (E15, PA1) · 24/3/1993 · workbook 80% · trang 60% · giờ dùng: không ghi; tháng: lịch Ất Mão → trang Bính Thìn; không ghi giờ – can giờ suy từ bảng trang
- data39 – mục II (E27, PA1) · 24/3/1993 · workbook 3 · trang -2 · giờ dùng: không ghi; tháng: lịch Ất Mão → trang Bính Thìn; không ghi giờ – can giờ suy từ bảng trang
- data39 – 6 năng lực (E48:E53) · 4/7/1979 · workbook [10, 5, 10, 0, 80, 80] · trang [5, 10, 15, 0, 10, 45] · giờ dùng: không ghi
- data39 – mục I (E15, PA1) · 4/7/1979 · workbook 60% · trang 12% · giờ dùng: không ghi; tháng: lịch Canh Ngọ → trang Tân Mùi; không ghi giờ – can giờ suy từ bảng trang
- data39 – 6 năng lực (E48:E53) · 02/12/2011 · workbook [10, 40, 25, 30, 0, 100] · trang [50, 80, 50, 5, 40, 40] · giờ dùng: 10h20
- data39 – mục I (E15, PA1) · 02/12/2011 · workbook 80% · trang 68% · giờ dùng: 10h20; ngày: lịch Tân Mão → trang Giáp Tý; giờ: lịch Quý Tị → trang Ất Hợi; chi ngày có thể Tý/Ngọ/Thân/Tuất
- data39 – mục II (E27, PA1) · 02/12/2011 · workbook -2 · trang 1.75 · giờ dùng: 10h20; ngày: lịch Tân Mão → trang Giáp Tý; giờ: lịch Quý Tị → trang Ất Hợi; chi ngày có thể Tý/Ngọ/Thân/Tuất
- data39 – 6 năng lực (E48:E53) · 1/7/1979 · workbook [0, 80, 5, 0, 10, 20] · trang [5, 10, 0, 5, 5, 30] · giờ dùng: không ghi
- data39 – 6 năng lực (E48:E53) · 31/10/1984 · workbook [5, 30, 40, 5, 45, 10] · trang [0, 20, 15, 0, 80, 80] · giờ dùng: 1h00
- dv16 – t = mục II (E27, PA1) · 21/07/2018 · workbook -1.25 · trang -0.25 · Mậu Tuất – Kỷ Mùi – Giáp … – …
- dv16 – 8 điểm đại vận (E70:E77) · 02/03/2014 · workbook [-70, 30, -50, 0, 20, -10, 50, -5] · trang [100, -70, 30, -50, 0, 10, -20, 50] · Giáp Ngọ – Đinh Mão (tháng âm) – Nhâm … – …
- dv16 – 8 điểm đại vận (E70:E77) · 05/09/2013 · workbook [0, 50, -45, 0, -20, 5, 0, -100] · trang [-80, 5, 50, -45, 0, -10, 10, 5] · Quý Tị – Tân Dậu (tháng âm) – Giáp … – …
- dv16 – t = mục II (E27, PA1) · 05/09/2013 · workbook -3 · trang -1 · Quý Tị – Tân Dậu (tháng âm) – Giáp … – …
- dv16 – 8 điểm đại vận (E70:E77) · 05/09/2005 · workbook [-10, 5, 50, -50, 0, -80, 80, 50] · trang [0, 50, -50, 5, -80, 80, 50, -5] · Ất Dậu – Ất Dậu (tháng âm) – Nhâm … – …
- dv16 – t = mục II (E27, PA1) · 05/09/2005 · workbook 1.75 · trang 0 · Ất Dậu – Ất Dậu (tháng âm) – Nhâm … – …
- dv16 – X = Thực tế − Đại vận · 05/09/2005 · workbook -10 · trang -30 · Ất Dậu – Ất Dậu (tháng âm) – Nhâm … – …
- dv16 – t = mục II (E27, PA1) · 04/02/1977 (trước Lập xuân) · workbook 0 · trang -1 · Bính Thìn – Tân Sửu – Nhâm … – …
- dv16 – X = Thực tế − Đại vận · 04/02/1977 (trước Lập xuân) · workbook -30 · trang -50 · Bính Thìn – Tân Sửu – Nhâm … – …
- dv16 – t = mục II (E27, PA1) · 04/02/1977 (sau Lập xuân) · workbook -2 · trang 0 · Đinh Tị – Nhâm Dần – Nhâm … – …
- dv16 – X = Thực tế − Đại vận · 04/02/1977 (sau Lập xuân) · workbook -50 · trang -30 · Đinh Tị – Nhâm Dần – Nhâm … – …
- dv16 – t = mục II (E27, PA1) · 07/06/1972 · workbook -0.25 · trang 1.75 · Nhâm Tý – Bính Ngọ – Kỷ … – …
- dv16 – X = Thực tế − Đại vận · 07/06/1972 · workbook -50 · trang -10 · Nhâm Tý – Bính Ngọ – Kỷ … – …
- dv16 – t = mục II (E27, PA1) · 24/08/2011 · workbook 1 · trang 0 · Tân Mão – Bính Thân – Tân … – …
- dv16 – X = Thực tế − Đại vận · 24/08/2011 · workbook -20 · trang -30 · Tân Mão – Bính Thân – Tân … – …
- dv16 – t = mục II (E27, PA1) · 09/09/1987 · workbook 4.5 · trang 5.5 · Đinh Mão – Kỷ Dậu – Tân … – …


### Phân loại chỗ workbook khác trang (T8 – B)

| Nhóm | Số chỗ | Nguyên nhân | Bằng chứng |
|---|---|---|---|
| data39 nhập theo NGÀY GIỜ dương lịch – năng lực, mục I, mục II | 8 + 10 + 9 | Trang tự dùng trụ khác lịch (giờ, có khi cả tháng/ngày) – chính data39.json ghi ở trường `note` (vd "giờ: lịch Ất Sửu → trang Tân Mùi") | Nhập bằng TỨ TRỤ TRANG GHI (biến thể T8_D39_TRU): **34/34** năng lực, **34/34** mục I, **34/34** mục II khớp trang |
| dv16 – 8 điểm đại vận | 3 lá (02/03/2014, 05/09/2013, 05/09/2005) | Trang dùng **tháng âm lịch** làm chi tháng; workbook dùng tháng theo tiết khí | Cong_thuc_bieu_do_dai_van.md §4 đã ghi đúng 3 lá này |
| dv16 – t (mục II) và X | 8 lá / 5 lá | Trang không ghi đủ trụ ngày/giờ ("Giáp … – …"); mục II cần đủ 4 trụ nên không xác định được từ dữ liệu trang (phần mềm dùng giờ giả định ghi trong tests/cases_t8.json) | Cả 5 lá dv16 ghi đủ 4 trụ đều khớp t |
| rows34 – mục I, mục II (tứ trụ trang) | 0 | – | 34/34 và 34/34 |

Không có chỗ chênh nào do công thức. Đề xuất ghi các nhóm này vào MAU_THUAN (MT-15 – chờ duyệt).
