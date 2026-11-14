# BÁO CÁO KIỂM THỬ – GIAI ĐOẠN 1

Ngày chạy: 13/11/2026 23:55 · Nguồn: workbook SHA-256 `7ae759658d290172bbbf7b6642e24532239e3b9ea54ca71e2696802411bca4c5`

Phạm vi đã cài: CSDL (toàn sheet), NHAP_LIEU (toàn sheet), TRUNG_GIAN (toàn sheet), DAI_VAN (C5:C28), KIEM_THU (toàn sheet)

## T1 – Đáp án từng ô (05_KIEM_THU/dap_an_o, 30 lá, 1.8 giây)

| Sheet | Ô khớp / tổng | Kết quả |
|---|---|---|
| NHAP_LIEU | 1530 / 1530 | ĐẠT 100% |
| DAI_VAN | 720 / 720 | ĐẠT 100% |
| TRUNG_GIAN | 6750 / 6750 | ĐẠT 100% |
| CSDL | 840 / 840 | ĐẠT 100% |
| KIEM_THU | 420 / 420 | ĐẠT 100% |
| **Tổng** | **10260 / 10260** | ĐẠT |

Ô lệch: không có

## T3 – Lịch (05_KIEM_THU/lich, 710 thời điểm, 2.1 giây)

| Hạng mục | Khớp / tổng |
|---|---|
| Tứ trụ tự tính NHAP_LIEU C37:F38 (8 ô) | 710 / 710 |
| Cảnh báo sát giờ/sát tiết DUNG_HY_KY!C180 | chưa so – thuộc GĐ2 |

## Âm lịch (battu_core/lich_am.py)

| Hạng mục | Khớp / tổng |
|---|---|
| Dương → âm trùng lunar.py gốc, mọi ngày 1899–2101 | 74144 / 74144 |
| Khứ hồi âm → dương → âm, mọi ngày | 74144 / 74144 |
| data39.json dương → âm | 39 / 39 |
| data39.json âm → dương | 39 / 39 |

## T1b – Nhánh tùy chọn (MT-05) và T1-KRG (MT-04) – đáp án lấy bằng tools/oracle_mo_rong.py

Kiểm chính oracle mở rộng: biến thể KRG0 (mặc định) trùng dap_an_o **856410 / 856410** ô.

| Biến thể | Mô tả | Số lá | Ô khớp / tổng (sheet GĐ1) | Kết quả |
|---|---|---|---|---|
| KRG0 | Mặc định (như dap_an_o) + vùng Bảng dữ liệu KHONG_RO_GIO C118:D285 (MT-04) | 30 | 10260 / 10260 | ĐẠT 100% |
| B1_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ | 30 | 10133 / 10260 | LỆCH 127 – đáp án hỏng do Bảng dữ liệu LibreOffice (MT-13), thay bằng B1b |
| B2_THEO_PP1 | NHAP_LIEU C14 = Theo PP1 | 30 | 10260 / 10260 | ĐẠT 100% |
| B3_THEO_PP2 | NHAP_LIEU C14 = Theo PP2 | 30 | 10260 / 10260 | ĐẠT 100% |
| B4_PA1 | NGU_HANH_THIEU E3 = PA1 – Theo trang | 30 | 10260 / 10260 | ĐẠT 100% |
| B5_SU_KIEN | Nhập sự kiện mẫu: SU_KIEN B6:C10, KHONG_RO_GIO C11:D13 (năm theo DAI_VAN C22 của từng lá), THEO_DOI N/P dòng 12, 13, 20 | 30 | 10260 / 10260 | ĐẠT 100% |
| KRG_VONG | Mặc định + Bảng dữ liệu KHONG_RO_GIO tính bằng vòng 12 giờ (MT-13) | 30 | 10260 / 10260 | ĐẠT 100% |
| B1b_KHONG_RO_GIO | NHAP_LIEU C9 = Không rõ giờ; Bảng dữ liệu tính bằng vòng 12 giờ (MT-13) – thay B1 | 30 | 10260 / 10260 | ĐẠT 100% |

