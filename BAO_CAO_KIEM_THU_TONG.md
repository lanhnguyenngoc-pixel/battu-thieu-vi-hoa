# BÁO CÁO KIỂM THỬ TỔNG – GIAI ĐOẠN 6

Ngày chạy: 14/11/2026 08:48 · Nguồn: workbook SHA-256 `7ae759658d290172bbbf7b6642e24532239e3b9ea54ca71e2696802411bca4c5`

Phạm vi đã cài: CSDL (toàn sheet), NHAP_LIEU (toàn sheet), TRUNG_GIAN (toàn sheet), DAI_VAN (toàn sheet), KIEM_THU (toàn sheet), PP1_DINH_TINH (toàn sheet), PP2_DINH_LUONG (toàn sheet), THONG_CAN (toàn sheet), DOI_CHIEU (toàn sheet), DUNG_HY_KY (toàn sheet), LUU_NIEN (toàn sheet), THEO_DOI (toàn sheet), NGU_HANH_THIEU (toàn sheet), LUC_THAN (toàn sheet), HON_NHAN (toàn sheet), MENH_CUC (toàn sheet), TAI_LOC (toàn sheet), TO_CHAT (toàn sheet), SU_NGHIEP (toàn sheet), SUC_KHOE (toàn sheet), THAN_SAT (toàn sheet), SAO (toàn sheet), BAO_PHU (toàn sheet), KHONG_RO_GIO (toàn sheet), KET_LUAN (toàn sheet), LUAN_GIAI (toàn sheet), DOC_NHANH (toàn sheet), SU_KIEN (toàn sheet), TRANG_CHU (toàn sheet), THAM_KHAO (toàn sheet)

## TỔNG HỢP

| Kiểm thử | Phạm vi | Khớp / tổng | Kết quả |
|---|---|---|---|
| T1 | 33 sheet × 30 lá (05_KIEM_THU/dap_an_o) | 856.410 / 856.410 | ĐẠT 100% |
| T1b | 7 biến thể tùy chọn × 30 lá (gồm Không rõ giờ – vòng 12 giờ) | 6.059.970 / 6.059.970 | ĐẠT 100% |
| T2 | 48 khóa kết quả chính × 501 lá | 24.048 / 24.048 | ĐẠT 100% |
| T1-501 | mọi ô công thức × 501 lá | 14.053.050 / 14.053.050 | ĐẠT 100% |
| T3 | lịch: 710 thời điểm × 9 ô | 710 / 710 | ĐẠT 100% |
| Âm lịch | lunar.py + data39 | 74.144 / 74.144 · 39 / 39 | ĐẠT 100% |
| T4 | hợp/xung 510 lá | 510 / 510 | ĐẠT 100% |
| T5 | xuất JSON / XLSX / PDF (5 lá) | 142.735 / 142.735 · 144.075 / 144.075 · 5 / 5 | ĐẠT 100% |
| T6 | kho lá số SQLite (mở lại như lúc lưu + 5 thao tác) | 12 / 12 · 5 / 5 | ĐẠT 100% |
| **T7** | **20 lá ngẫu nhiên MỚI**, oracle_dump.py, 33 sheet, không mượn ô | **570.940 / 570.940** | **ĐẠT 100%** |
| T8 | ngũ hành thiếu: 55 lá × PA1/PA2, mọi sheet | 3.174.270 / 3.174.270 | ĐẠT 100% |

Biến thể B1 (đáp án LibreOffice cũ khi Không rõ giờ, 860.359/865.710) là **đáp án hỏng do lỗi Bảng dữ liệu của LibreOffice (MT-13)**, đã thay bằng B1b (đạt 100%) – giữ lại trong bảng T1b để đối chứng.

## T1 – Đáp án từng ô (05_KIEM_THU/dap_an_o, 30 lá, 33.3 giây)

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| TRANG_CHU | 120 / 120 | ĐẠT 100% |
| NHAP_LIEU | 1530 / 1530 | ĐẠT 100% |
| KET_LUAN | 16110 / 16110 | ĐẠT 100% |
| LUAN_GIAI | 28500 / 28500 | ĐẠT 100% |
| THEO_DOI | 150180 / 150180 | ĐẠT 100% |
| KHONG_RO_GIO | 14910 / 14910 | ĐẠT 100% |
| DOI_CHIEU | 3120 / 3120 | ĐẠT 100% |
| DUNG_HY_KY | 10950 / 10950 | ĐẠT 100% |
| PP1_DINH_TINH | 7380 / 7380 | ĐẠT 100% |
| PP2_DINH_LUONG | 8040 / 8040 | ĐẠT 100% |
| THONG_CAN | 8670 / 8670 | ĐẠT 100% |
| THAN_SAT | 15330 / 15330 | ĐẠT 100% |
| SAO | 10560 / 10560 | ĐẠT 100% |
| DAI_VAN | 19740 / 19740 | ĐẠT 100% |
| LUU_NIEN | 334440 / 334440 | ĐẠT 100% |
| LUC_THAN | 50190 / 50190 | ĐẠT 100% |
| HON_NHAN | 35940 / 35940 | ĐẠT 100% |
| MENH_CUC | 9570 / 9570 | ĐẠT 100% |
| TAI_LOC | 8490 / 8490 | ĐẠT 100% |
| TO_CHAT | 27090 / 27090 | ĐẠT 100% |
| SU_NGHIEP | 39120 / 39120 | ĐẠT 100% |
| SUC_KHOE | 26520 / 26520 | ĐẠT 100% |
| TRUNG_GIAN | 6750 / 6750 | ĐẠT 100% |
| CSDL | 840 / 840 | ĐẠT 100% |
| KIEM_THU | 420 / 420 | ĐẠT 100% |
| THAM_KHAO | 360 / 360 | ĐẠT 100% |
| NGU_HANH_THIEU | 7260 / 7260 | ĐẠT 100% |
| BAO_PHU | 1260 / 1260 | ĐẠT 100% |
| DOC_NHANH | 390 / 390 | ĐẠT 100% |
| SU_KIEN | 12630 / 12630 | ĐẠT 100% |
| **Tổng** | **856410 / 856410** | ĐẠT |

Ô lệch: không có

## T3 – Lịch (05_KIEM_THU/lich, 710 thời điểm, 1.8 giây)

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
| KRG0 | Mặc định (như dap_an_o) + vùng Bảng dữ liệu KHONG_RO_GIO C118:D285 (MT-04) | 30 | 865710 / 865710 | ĐẠT 100% |
| B1_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ | 30 | 860359 / 865710 | LỆCH 5351 – đáp án hỏng do Bảng dữ liệu LibreOffice (MT-13), thay bằng B1b |
| B2_THEO_PP1 | NHAP_LIEU C14 = Theo PP1 | 30 | 865710 / 865710 | ĐẠT 100% |
| B3_THEO_PP2 | NHAP_LIEU C14 = Theo PP2 | 30 | 865710 / 865710 | ĐẠT 100% |
| B4_PA1 | NGU_HANH_THIEU E3 = PA1 – Theo trang | 30 | 865710 / 865710 | ĐẠT 100% |
| B5_SU_KIEN | Nhập sự kiện mẫu: SU_KIEN B6:C10, KHONG_RO_GIO C11:D13 (năm theo DAI_VAN C22 của từng lá), THEO_DOI N/P dòng 12, 13, 20 | 30 | 865710 / 865710 | ĐẠT 100% |
| KRG_VONG | Mặc định + Bảng dữ liệu KHONG_RO_GIO tính bằng vòng 12 giờ (MT-13) | 30 | 865710 / 865710 | ĐẠT 100% |
| B1b_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ; Bảng dữ liệu tính bằng vòng 12 giờ (MT-13) – thay B1 | 30 | 865710 / 865710 | ĐẠT 100% |

## T2 – Kết quả chính 501 lá (bo_la + ket_qua_6.3, 554.7 giây)

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
| k4 | KET_LUAN!C4 | 501 / 501 |
| k5 | KET_LUAN!C5 | 501 / 501 |
| k6 | KET_LUAN!C6 | 501 / 501 |
| k7 | KET_LUAN!C7 | 501 / 501 |
| k8 | KET_LUAN!C8 | 501 / 501 |
| k9 | KET_LUAN!C9 | 501 / 501 |
| k10 | KET_LUAN!C10 | 501 / 501 |
| k11 | KET_LUAN!C11 | 501 / 501 |
| k12 | KET_LUAN!C12 | 501 / 501 |
| skJ6 | SUC_KHOE!J6 | 501 / 501 |
| skJ7 | SUC_KHOE!J7 | 501 / 501 |
| skJ8 | SUC_KHOE!J8 | 501 / 501 |
| dk | LUU_NIEN!DK31 | 501 / 501 |
| dl | LUU_NIEN!DL31 | 501 / 501 |
| dn | DOC_NHANH!C14 | 501 / 501 |
| m16 | TO_CHAT!M16 | 501 / 501 |
| c16 | TO_CHAT!C16 | 501 / 501 |
| ai64 | HON_NHAN!AI64 | 501 / 501 |
| ag64 | HON_NHAN!AG64 | 501 / 501 |
| x64 | HON_NHAN!X64 | 501 / 501 |
| r64 | HON_NHAN!R64 | 501 / 501 |
| s64 | HON_NHAN!S64 | 501 / 501 |
| c206 | KET_LUAN!C206 | 501 / 501 |
| c72 | KET_LUAN!C72 | 501 / 501 |
| c179 | KET_LUAN!C179 | 501 / 501 |
| f115 | LUAN_GIAI!F115 | 501 / 501 |

Khóa thuộc giai đoạn sau (so khi cài sheet tương ứng): 

(b) T1-501 – so TỪNG Ô các sheet đã cài trên 501 lá (đáp án tools/oracle_mo_rong.py, biến thể T2_501):

| Sheet | Ô khớp / tổng |
|---|---|
| TRANG_CHU | 2004 / 2004 |
| NHAP_LIEU | 25551 / 25551 |
| KET_LUAN | 269037 / 269037 |
| LUAN_GIAI | 475950 / 475950 |
| THEO_DOI | 2508006 / 2508006 |
| DOI_CHIEU | 52104 / 52104 |
| DUNG_HY_KY | 182865 / 182865 |
| PP1_DINH_TINH | 123246 / 123246 |
| PP2_DINH_LUONG | 134268 / 134268 |
| THONG_CAN | 144789 / 144789 |
| THAN_SAT | 256011 / 256011 |
| SAO | 176352 / 176352 |
| DAI_VAN | 329658 / 329658 |
| LUU_NIEN | 5585148 / 5585148 |
| LUC_THAN | 838173 / 838173 |
| HON_NHAN | 600198 / 600198 |
| MENH_CUC | 159819 / 159819 |
| TAI_LOC | 141783 / 141783 |
| TO_CHAT | 452403 / 452403 |
| SU_NGHIEP | 653304 / 653304 |
| SUC_KHOE | 442884 / 442884 |
| TRUNG_GIAN | 112725 / 112725 |
| CSDL | 14028 / 14028 |
| KIEM_THU | 7014 / 7014 |
| THAM_KHAO | 6012 / 6012 |
| NGU_HANH_THIEU | 121242 / 121242 |
| BAO_PHU | 21042 / 21042 |
| DOC_NHANH | 6513 / 6513 |
| SU_KIEN | 210921 / 210921 |
| **Tổng** | **14053050 / 14053050** |

## T4 – Hợp/xung (hop_xung, 510 lá gồm 9 ca biên, 1.9 giây)

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

## T8 – Ngũ hành thiếu, biểu đồ (bộ lá chuẩn tests/t8_chuan.json)

Lá ngày giờ dương lịch (từ data39, dv16), tứ trụ theo lịch tiết khí. So từng ô mọi sheet đã cài với đáp án workbook ở hai chế độ E3.

| Chế độ | Số lá | NGU_HANH_THIEU | Mọi sheet đã cài |
|---|---|---|---|
| T8C_PA1 | 55 | 13310 / 13310 | 1587135 / 1587135 ĐẠT 100% |
| T8C_PA2 | 55 | 13310 / 13310 | 1587135 / 1587135 ĐẠT 100% |

## T6 – Kho lá số SQLite (YÊU CẦU 5.3)

Lưu 12 lá (t8_chuan.json, có lá Không rõ giờ) → mở lại → tính lại; so khối TÓM TẮT KET_LUAN B2:C12 + tứ trụ với ảnh chụp lúc lưu (17.3 giây).

| Hạng mục | Kết quả |
|---|---|
| Mở lại cho kết quả như lúc lưu | 12 / 12 |
| tìm theo tên | ĐẠT |
| lọc nhóm | ĐẠT |
| danh sách nhóm | ĐẠT |
| sửa thông tin | ĐẠT |
| xóa | ĐẠT |

## T7 – 20 lá NGẪU NHIÊN MỚI (oracle_dump.py nguyên bản, không mượn ô nào)

tests/t7_ngau_nhien.json (hạt giống 20261114): 14 lá dương lịch 1900–2100, 6 lá tự nhập can chi; 14.7 giây.

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| TRANG_CHU | 80 / 80 | ĐẠT 100% |
| NHAP_LIEU | 1020 / 1020 | ĐẠT 100% |
| KET_LUAN | 10740 / 10740 | ĐẠT 100% |
| LUAN_GIAI | 19000 / 19000 | ĐẠT 100% |
| THEO_DOI | 100120 / 100120 | ĐẠT 100% |
| KHONG_RO_GIO | 9940 / 9940 | ĐẠT 100% |
| DOI_CHIEU | 2080 / 2080 | ĐẠT 100% |
| DUNG_HY_KY | 7300 / 7300 | ĐẠT 100% |
| PP1_DINH_TINH | 4920 / 4920 | ĐẠT 100% |
| PP2_DINH_LUONG | 5360 / 5360 | ĐẠT 100% |
| THONG_CAN | 5780 / 5780 | ĐẠT 100% |
| THAN_SAT | 10220 / 10220 | ĐẠT 100% |
| SAO | 7040 / 7040 | ĐẠT 100% |
| DAI_VAN | 13160 / 13160 | ĐẠT 100% |
| LUU_NIEN | 222960 / 222960 | ĐẠT 100% |
| LUC_THAN | 33460 / 33460 | ĐẠT 100% |
| HON_NHAN | 23960 / 23960 | ĐẠT 100% |
| MENH_CUC | 6380 / 6380 | ĐẠT 100% |
| TAI_LOC | 5660 / 5660 | ĐẠT 100% |
| TO_CHAT | 18060 / 18060 | ĐẠT 100% |
| SU_NGHIEP | 26080 / 26080 | ĐẠT 100% |
| SUC_KHOE | 17680 / 17680 | ĐẠT 100% |
| TRUNG_GIAN | 4500 / 4500 | ĐẠT 100% |
| CSDL | 560 / 560 | ĐẠT 100% |
| KIEM_THU | 280 / 280 | ĐẠT 100% |
| THAM_KHAO | 240 / 240 | ĐẠT 100% |
| NGU_HANH_THIEU | 4840 / 4840 | ĐẠT 100% |
| BAO_PHU | 840 / 840 | ĐẠT 100% |
| DOC_NHANH | 260 / 260 | ĐẠT 100% |
| SU_KIEN | 8420 / 8420 | ĐẠT 100% |
| **Tổng** | **570940 / 570940** | ĐẠT 100% |

## T5 – Xuất PDF / XLSX / JSON (5 lá T7, 16.5 giây)

| Định dạng | Kiểm | Kết quả |
|---|---|---|
| JSON | mọi ô công thức trùng đáp án oracle | 142735 / 142735 |
| XLSX | mở lại, mọi ô công thức của 30 sheet hiển thị trùng JSON; giữ gộp ô, cột/dòng ẩn | 144075 / 144075 |
| PDF | số trang, "Trang i/n" mỗi trang, khối KẾT LUẬN, phiên bản quy tắc | 5 / 5 lá |


## Bộ cài – kiểm cài / gỡ / chạy không mạng (GitHub Actions, lượt 43547513998)

Mỗi máy: chạy T7 từ mã nguồn (570.940 ô) → PyInstaller → bộ cài → **cài → tự kiểm khi chặn mạng → gỡ → kiểm gỡ sạch**.
Tự kiểm = dấu vân tay SHA-256 toàn bộ kết quả (JSON {sheet: {ô: giá trị}}) của 3 lá chuẩn (có 1 lá Không rõ giờ) + xuất PDF/XLSX/JSON.

| Máy | T7 mã nguồn | Cài | Chặn mạng | Tự kiểm 3 lá (dấu vân tay) | Xuất | Gỡ |
|---|---|---|---|---|---|---|
| Windows Server (x64) – Inno Setup | 570.940/570.940 | ĐẠT | Tường lửa chặn chiều ra của BatTuThieuViHoa.exe | ĐẠT c9fa2c11 · cd14e7bd · 1ec29dbd | PDF 28 trang, XLSX, JSON | ĐẠT (mất tệp chạy + mục Settings > Apps) |
| macOS 15 – Apple M (arm64) – .dmg | 570.940/570.940 | ĐẠT | sandbox-exec deny network | ĐẠT – trùng | ĐẠT | ĐẠT |
| macOS 15 – Intel (x86_64) – .dmg | 570.940/570.940 | ĐẠT | sandbox-exec deny network | ĐẠT – trùng | ĐẠT | ĐẠT |
| Linux (bản gốc, đối chứng) | 570.940/570.940 | – | unshare -rn | ĐẠT – trùng | ĐẠT | – |

Dấu vân tay trùng tuyệt đối giữa 4 hệ điều hành → kết quả không phụ thuộc máy.

| Tệp | Dung lượng | SHA-256 |
|---|---|---|
| BatTuThieuViHoa_Setup_1.0.0_Windows.exe | 42,3 MB | 22a77b2e2fd4de36d39a3b09a2095bb5f56b9a1948c264183f14047faffecf0f |
| BatTuThieuViHoa_1.0.0_macOS_arm64.dmg | 51,7 MB | b281bc7aac9348ece864a1910dc6306aa7722cb3c6ca9ad268a3ced76d39f9ac |
| BatTuThieuViHoa_1.0.0_macOS_x86_64.dmg | 56,2 MB | 102971c5516e4f13d1a372a86fd9d4bb8cb257e25a0f4538055cba20fa2c6fa6 |

Chưa ký số: Windows SmartScreen / macOS Gatekeeper hỏi lần đầu (xem Hướng dẫn sử dụng mục 1).
