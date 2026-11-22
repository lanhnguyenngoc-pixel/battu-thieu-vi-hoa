# BÁO CÁO KIỂM THỬ TỔNG – GIAI ĐOẠN 6

Ngày chạy: 22/11/2026 18:55 · Nguồn: workbook SHA-256 `671245aea574f577f3cb312db9b3db0d1ce1973b107488ca7b828a9ceb4d60da`

Phạm vi đã cài: CSDL (toàn sheet), NHAP_LIEU (toàn sheet), TRUNG_GIAN (toàn sheet), DAI_VAN (toàn sheet), KIEM_THU (toàn sheet), PP1_DINH_TINH (toàn sheet), PP2_DINH_LUONG (toàn sheet), THONG_CAN (toàn sheet), DOI_CHIEU (toàn sheet), DUNG_HY_KY (toàn sheet), LUU_NIEN (toàn sheet), THEO_DOI (toàn sheet), NGU_HANH_THIEU (toàn sheet), LUC_THAN (toàn sheet), HON_NHAN (toàn sheet), MENH_CUC (toàn sheet), TAI_LOC (toàn sheet), TO_CHAT (toàn sheet), SU_NGHIEP (toàn sheet), SUC_KHOE (toàn sheet), THAN_SAT (toàn sheet), SAO (toàn sheet), BAO_PHU (toàn sheet), KHONG_RO_GIO (toàn sheet), KET_LUAN (toàn sheet), LUAN_GIAI (toàn sheet), DOC_NHANH (toàn sheet), SU_KIEN (toàn sheet), TRANG_CHU (toàn sheet), THAM_KHAO (toàn sheet)

## T1 – Đáp án từng ô (05_KIEM_THU/dap_an_o, 30 lá, 58.1 giây)

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| TRANG_CHU | 120 / 120 | ĐẠT 100% |
| NHAP_LIEU | 1560 / 1560 | ĐẠT 100% |
| KET_LUAN | 16110 / 16110 | ĐẠT 100% |
| LUAN_GIAI | 28500 / 28500 | ĐẠT 100% |
| THEO_DOI | 150180 / 150180 | ĐẠT 100% |
| KHONG_RO_GIO | 14910 / 14910 | ĐẠT 100% |
| DOI_CHIEU | 3150 / 3150 | ĐẠT 100% |
| DUNG_HY_KY | 12270 / 12270 | ĐẠT 100% |
| PP1_DINH_TINH | 7530 / 7530 | ĐẠT 100% |
| PP2_DINH_LUONG | 8040 / 8040 | ĐẠT 100% |
| THONG_CAN | 8670 / 8670 | ĐẠT 100% |
| THAN_SAT | 15330 / 15330 | ĐẠT 100% |
| SAO | 10560 / 10560 | ĐẠT 100% |
| DAI_VAN | 26730 / 26730 | ĐẠT 100% |
| LUU_NIEN | 367440 / 367440 | ĐẠT 100% |
| LUC_THAN | 50280 / 50280 | ĐẠT 100% |
| HON_NHAN | 35940 / 35940 | ĐẠT 100% |
| MENH_CUC | 10560 / 10560 | ĐẠT 100% |
| TAI_LOC | 9450 / 9450 | ĐẠT 100% |
| TO_CHAT | 27090 / 27090 | ĐẠT 100% |
| SU_NGHIEP | 39150 / 39150 | ĐẠT 100% |
| SUC_KHOE | 26520 / 26520 | ĐẠT 100% |
| TRUNG_GIAN | 6750 / 6750 | ĐẠT 100% |
| CSDL | 840 / 840 | ĐẠT 100% |
| KIEM_THU | 420 / 420 | ĐẠT 100% |
| THAM_KHAO | 360 / 360 | ĐẠT 100% |
| NGU_HANH_THIEU | 7260 / 7260 | ĐẠT 100% |
| BAO_PHU | 1260 / 1260 | ĐẠT 100% |
| DOC_NHANH | 390 / 390 | ĐẠT 100% |
| SU_KIEN | 17160 / 17160 | ĐẠT 100% |
| **Tổng** | **904530 / 904530** | ĐẠT |

Ô lệch: không có

## T3 – Lịch (05_KIEM_THU/lich, 710 thời điểm, 2.6 giây)

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

Kiểm chính oracle mở rộng: biến thể KRG0 (mặc định) trùng dap_an_o **904530 / 904530** ô.

| Biến thể | Mô tả | Số lá | Ô khớp / tổng (sheet đã cài) | Kết quả |
|---|---|---|---|---|
| KRG0 | Mặc định (như dap_an_o) + vùng Bảng dữ liệu KHONG_RO_GIO C118:D285 (MT-04) | 30 | 913830 / 913830 | ĐẠT 100% |
| B1_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ | 30 | 908143 / 913830 | LỆCH 5687 – đáp án hỏng do Bảng dữ liệu LibreOffice (MT-13), thay bằng B1b |
| B2_THEO_PP1 | NHAP_LIEU C14 = Theo PP1 | 30 | 913830 / 913830 | ĐẠT 100% |
| B3_THEO_PP2 | NHAP_LIEU C14 = Theo PP2 | 30 | 913830 / 913830 | ĐẠT 100% |
| B4_PA1 | NGU_HANH_THIEU E3 = PA1 – Theo trang | 30 | 913830 / 913830 | ĐẠT 100% |
| B5_SU_KIEN | Nhập sự kiện mẫu: SU_KIEN B6:C10, KHONG_RO_GIO C11:D13 (năm theo DAI_VAN C22 của từng lá), THEO_DOI N/P dòng 12, 13, 20 | 30 | 913830 / 913830 | ĐẠT 100% |
| KRG_VONG | Mặc định + Bảng dữ liệu KHONG_RO_GIO tính bằng vòng 12 giờ (MT-13) | 30 | 913830 / 913830 | ĐẠT 100% |
| B1b_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ; Bảng dữ liệu tính bằng vòng 12 giờ (MT-13) – thay B1 | 30 | 913830 / 913830 | ĐẠT 100% |

## T2 – Kết quả chính 501 lá (bo_la + ket_qua_6.3, 971.1 giây)

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
| NHAP_LIEU | 26052 / 26052 |
| KET_LUAN | 269037 / 269037 |
| LUAN_GIAI | 475950 / 475950 |
| THEO_DOI | 2508006 / 2508006 |
| DOI_CHIEU | 52605 / 52605 |
| DUNG_HY_KY | 204909 / 204909 |
| PP1_DINH_TINH | 125751 / 125751 |
| PP2_DINH_LUONG | 134268 / 134268 |
| THONG_CAN | 144789 / 144789 |
| THAN_SAT | 256011 / 256011 |
| SAO | 176352 / 176352 |
| DAI_VAN | 446391 / 446391 |
| LUU_NIEN | 6136248 / 6136248 |
| LUC_THAN | 839676 / 839676 |
| HON_NHAN | 600198 / 600198 |
| MENH_CUC | 176352 / 176352 |
| TAI_LOC | 157815 / 157815 |
| TO_CHAT | 452403 / 452403 |
| SU_NGHIEP | 653805 / 653805 |
| SUC_KHOE | 442884 / 442884 |
| TRUNG_GIAN | 112725 / 112725 |
| CSDL | 14028 / 14028 |
| KIEM_THU | 7014 / 7014 |
| THAM_KHAO | 6012 / 6012 |
| NGU_HANH_THIEU | 121242 / 121242 |
| BAO_PHU | 21042 / 21042 |
| DOC_NHANH | 6513 / 6513 |
| SU_KIEN | 286572 / 286572 |
| **Tổng** | **14856654 / 14856654** |

## T4 – Hợp/xung (hop_xung, 510 lá gồm 9 ca biên, 2.6 giây)

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
| T8C_PA1 | 55 | 13310 / 13310 | 1675355 / 1675355 ĐẠT 100% |
| T8C_PA2 | 55 | 13310 / 13310 | 1675355 / 1675355 ĐẠT 100% |

## T6 – Kho lá số SQLite (YÊU CẦU 5.3)

Lưu 12 lá (t8_chuan.json, có lá Không rõ giờ) → mở lại → tính lại; so khối TÓM TẮT KET_LUAN B2:C12 + tứ trụ với ảnh chụp lúc lưu (36.0 giây).

| Hạng mục | Kết quả |
|---|---|
| Mở lại cho kết quả như lúc lưu | 12 / 12 |
| tìm theo tên | ĐẠT |
| lọc nhóm | ĐẠT |
| danh sách nhóm | ĐẠT |
| sửa thông tin | ĐẠT |
| xóa | ĐẠT |
| sao lưu | ĐẠT |
| khôi phục | ĐẠT |
| thống kê đối chiếu | ĐẠT |
| tệp sai bị từ chối | ĐẠT |

## T7 – 20 lá NGẪU NHIÊN MỚI (oracle_dump.py nguyên bản, không mượn ô nào)

tests/t7_ngau_nhien.json (hạt giống 20261114): 14 lá dương lịch 1900–2100, 6 lá tự nhập can chi; 23.0 giây.

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| TRANG_CHU | 80 / 80 | ĐẠT 100% |
| NHAP_LIEU | 1040 / 1040 | ĐẠT 100% |
| KET_LUAN | 10740 / 10740 | ĐẠT 100% |
| LUAN_GIAI | 19000 / 19000 | ĐẠT 100% |
| THEO_DOI | 100120 / 100120 | ĐẠT 100% |
| KHONG_RO_GIO | 9940 / 9940 | ĐẠT 100% |
| DOI_CHIEU | 2100 / 2100 | ĐẠT 100% |
| DUNG_HY_KY | 8180 / 8180 | ĐẠT 100% |
| PP1_DINH_TINH | 5020 / 5020 | ĐẠT 100% |
| PP2_DINH_LUONG | 5360 / 5360 | ĐẠT 100% |
| THONG_CAN | 5780 / 5780 | ĐẠT 100% |
| THAN_SAT | 10220 / 10220 | ĐẠT 100% |
| SAO | 7040 / 7040 | ĐẠT 100% |
| DAI_VAN | 17820 / 17820 | ĐẠT 100% |
| LUU_NIEN | 244960 / 244960 | ĐẠT 100% |
| LUC_THAN | 33520 / 33520 | ĐẠT 100% |
| HON_NHAN | 23960 / 23960 | ĐẠT 100% |
| MENH_CUC | 7040 / 7040 | ĐẠT 100% |
| TAI_LOC | 6300 / 6300 | ĐẠT 100% |
| TO_CHAT | 18060 / 18060 | ĐẠT 100% |
| SU_NGHIEP | 26100 / 26100 | ĐẠT 100% |
| SUC_KHOE | 17680 / 17680 | ĐẠT 100% |
| TRUNG_GIAN | 4500 / 4500 | ĐẠT 100% |
| CSDL | 560 / 560 | ĐẠT 100% |
| KIEM_THU | 280 / 280 | ĐẠT 100% |
| THAM_KHAO | 240 / 240 | ĐẠT 100% |
| NGU_HANH_THIEU | 4840 / 4840 | ĐẠT 100% |
| BAO_PHU | 840 / 840 | ĐẠT 100% |
| DOC_NHANH | 260 / 260 | ĐẠT 100% |
| SU_KIEN | 11440 / 11440 | ĐẠT 100% |
| **Tổng** | **603020 / 603020** | ĐẠT 100% |

## T5 – Xuất PDF / XLSX / JSON (5 lá T7, 43.8 giây)

| Định dạng | Kiểm | Kết quả |
|---|---|---|
| JSON | mọi ô công thức trùng đáp án oracle | 150755 / 150755 |
| XLSX | mở lại, mọi ô công thức của 30 sheet hiển thị trùng JSON; giữ gộp ô, cột/dòng ẩn | 151815 / 151815 |
| PDF | số trang, "Trang i/n" mỗi trang, khối KẾT LUẬN, phiên bản quy tắc | 5 / 5 lá |

## T9 – Nhập sự việc SU_KIEN (bản 1.1.0): 20 lá T7 × 6 sự việc, đáp án tools/oracle_mo_rong.py (T7_SU_KIEN)

| Hạng mục | Kết quả |
|---|---|
| Mọi ô công thức 33 sheet | 609220 / 609220 |
| Riêng SU_KIEN | 11440 / 11440 |
| Đường nhập LaSo.su_kien → ô SU_KIEN B/C/D trùng ô oracle đã nhập | 20 / 20 lá |
| Lưu SQLite → mở lại giữ nguyên sự việc | ĐẠT |

## T10 – Ô nhập bổ sung (bản 1.4.0): phối ngẫu, việc lớn tìm giờ sinh, đánh giá thực tế – 20 lá T7, đáp án tools/oracle_mo_rong.py (T10_BO_SUNG)

| Hạng mục | Kết quả |
|---|---|
| Mọi ô công thức 33 sheet | 609220 / 609220 |
| Riêng Hôn nhân / Không rõ giờ / Theo dõi | 23960 / 23960 · 16140 / 16140 · 100120 / 100120 |
| Đường nhập giao diện → ô workbook (C34:G34, C11:E22, N/O/P theo năm) | 6840 / 6840 ô |
| Lưu SQLite → mở lại giữ nguyên | ĐẠT |


## Phụ lục – kiểm thử bổ sung bản 6.9 (22/11/2026)

| Bộ kiểm | Kết quả |
|---|---|
| T11 hiển thị | 129860 ô đạt |
| T12 quan hệ | 521 lá đạt |
| T13 phân đoạn (có kiểm LaSo.to_dict/from_dict giữ phoi_tu_tru) | ĐẠT |
| T14 bản luận giải, T15 AI, T16 sách | ĐẠT; T16: đại vận khớp sách 120/171 (70%) so với workbook 76 (44%); vượng–nhược 56/65 (86%) |
| Kho lá khó | ĐẠT |
| Ghi chú | Test PDF (T5) từng đòi nhãn "Quy tắc 6.7" nên báo 0/5; đã sửa test, chạy lại T5 đạt 5/5 |
