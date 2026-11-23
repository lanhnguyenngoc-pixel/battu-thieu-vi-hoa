# Danh mục lá số chuẩn gốc của sách (23/11/2026)
Mục đích: một bộ lá có nguồn và nhãn rõ ràng để đo độ khớp của bộ công cụ 6.9 / phần mềm 1.9.2 khi sửa quy tắc. Chưa đổi quy tắc nào.
Tệp: `tests/danh_muc_la_so_chuan.json` (kho mã) và `DANH_MUC_LA_SO_CHUAN_GOC.xlsx` (bản đọc).

## 1. Cách lập
- Lấy 501 lá của bộ kiểm thử (`tests/cases501_a.json`, `cases501_b.json`). Bỏ 49 lá "Trang" (lá của người dùng, không nhãn) và 1 lá ngoài sách → còn 451 dòng từ sách.
- Gộp lá trùng tứ trụ giữa các nguồn → 433 tứ trụ khác nhau.
- Kiểm tứ trụ: đúng âm dương từng trụ; can tháng đúng ngũ hổ độn theo can năm; can giờ đúng ngũ thử độn theo can ngày. Can ngày không kiểm được vì thiếu ngày sinh.
- Chạy bộ công cụ trên từng lá (Nam, nhập can chi) lấy kết luận vượng/nhược cuối và dụng thần.

## 2. Phân hạng
| Hạng | Ý nghĩa | Số lá |
|---|---|---|
| A | Chuẩn gốc Thiệu Vĩ Hoa: tứ trụ qua kiểm tra, nhãn vượng/nhược rõ | 62 |
| B | Chuẩn gốc sách khác (Trích Thiên Tủy 261, Tử Bình Chân Thuyên 49, Uyên Hải 28, Hoàng Thăng Long 11 sau khi trừ lá trùng) | 344 |
| C | Cách đặc biệt (tòng, hóa, gia tường), không có nhãn vượng/nhược | 16 |
| M | Cùng tứ trụ nhưng hai sách ghi nhãn khác nhau | 1 |
| X | Tứ trụ không qua kiểm tra (nghi chép sai hoặc sách in sai) | 10 |

## 3. Độ khớp hiện tại
| Nhóm | Số lá | Vượng suy khớp | Lá có nhãn dụng thần | Dụng thần khớp |
|---|---|---|---|---|
| A – Thiệu Vĩ Hoa | 62 | 53 (85%) | 34 | 18 |
| B – sách khác | 344 | 225 (65%) | 258 | 90 |
| A + B | 406 | 278 (68%) | 292 | 108 |
Lưu ý: số này khác "64/73" ở báo cáo cũ vì cách đếm khác (bỏ cách đặc biệt, gộp lá trùng, bỏ lá nghi sai); khi so trước/sau phải dùng cùng một bộ này. Dụng thần: Trích Thiên Tủy có nhãn đủ; Thiệu Vĩ Hoa mới có 34 lá nhãn đã đối chiếu; ba sách còn lại gần như chưa có nhãn dụng thần. Tác giả khác chọn dụng thần theo phong cách khác nhau, nên hạng B chỉ là số tham khảo phụ.

## 4. Cách dùng khi sửa quy tắc
1. Đo trước và sau trên hạng A (chính), hạng B (phụ), báo cả hai.
2. Chỉ giữ thay đổi nếu hạng A không giảm và hạng B không giảm đáng kể.
3. Hạng X: soát lại ảnh sách rồi chuyển sang A/B. Hạng M: xem lại sách.
4. Lá mới người dùng gửi (như 9 lá "Một trăm cách đoán mệnh") thêm vào danh mục sau khi dựng tứ trụ và qua kiểm tra.

## 5. Hạng X – cần soát lại sách (10 lá)
- X-001 P1 tr.72 – Quý Mão,Mậu Ngọ,Tân Sửu,Đinh Mùi – P1 tr.72-73 (can giờ sai ngũ thử độn)
- X-002 Tòng sát VD1 – Nhâm Tí,Quý Sửu,Bính Tí,Canh Tí – P2 tr.374 (can giờ sai ngũ thử độn)
- X-003 TTT#203 – Mậu Thìn,Kỷ Dậu,Ất Hợi,Giáp Thân – Trích Thiên Tủy (Nhậm Thiết Tiều) (can tháng sai ngũ hổ độn)
- X-004 TTT#444 – Đinh Hợi,Tân Hợi,Tân Mùi,Mậu Thìn – Trích Thiên Tủy (Nhậm Thiết Tiều) (can giờ sai ngũ thử độn)
- X-005 TTT#455 – Nhâm Thìn,Kỷ Dậu,Quý Mão,Kỷ Mão – Trích Thiên Tủy (Nhậm Thiết Tiều) (can giờ sai ngũ thử độn)
- X-006 TB#181 – Giáp Tí,Bính Tí,Quý Sửu,Nhâm Thìn – Tử Bình Chân Thuyên (can giờ sai ngũ thử độn)
- X-007 UH63 – Ất Mão,Bính Tí,Bính Ngọ,Quý Tị – Uyên Hải Tử Bình (can tháng sai ngũ hổ độn)
- X-008 UH108 – Ất Dậu,Ất Mão,Tân Mão,Tân Mão – Uyên Hải Tử Bình (can tháng sai ngũ hổ độn)
- X-009 HTL X-V5 – Canh Tí,Đinh Sửu,Quý Dậu,Quý Tị – Hoàng Thăng Long (can tháng sai ngũ hổ độn; can giờ sai ngũ thử độn)
- X-010 HTL X-V12 – Ất Tị,Giáp Thân,Ất Mùi,Đinh Tị – Hoàng Thăng Long (can giờ sai ngũ thử độn)

## 6. Hạng A – lá còn lệch công cụ
- A-006 P2-VD7 (tính độ) (P2 tr.351-355) – Ất Dậu,Canh Thìn,Quý Dậu,Mậu Ngọ – sách vượng, công cụ nhược; dụng thần sách –, công cụ Kim (Ấn)
- A-009 P1-VD3 (nam) (P1 tr.184-186) – Giáp Thìn,Giáp Tuất,Quý Sửu,Nhâm Tí – sách nhược, công cụ vượng; dụng thần sách –, công cụ Thổ (Quan sát)
- A-023 P2-398 VD2 (P2 tr.398) – Ất Mão,Bính Tuất,Quý Tị,Nhâm Tí – sách nhược, công cụ vượng; dụng thần sách Kim, công cụ Thổ (Quan sát)
- A-024 P2-399 VD3 (P2 tr.399) – Canh Tuất,Bính Tuất,Tân Tị,Đinh Dậu – sách vượng, công cụ vượng; dụng thần sách Thủy, công cụ Hỏa (Quan sát)
- A-025 P2-400 VD4 (P2 tr.400 (trụ ngày điền Quý Mão theo lịch)) – Quý Sửu,Giáp Tí,Quý Mão,Quý Hợi – sách vượng, công cụ vượng; dụng thần sách Mộc, công cụ Thủy (Tỉ kiếp)
- A-027 P2-401 VD6 (P2 tr.401-402) – Đinh Tị,Canh Tuất,Đinh Tị,Mậu Thân – sách vượng, công cụ nhược; dụng thần sách Kim, công cụ Hỏa (Tỉ kiếp)
- A-030 P2-404 VD9 (P2 tr.404) – Ất Tị,Giáp Thân,Ất Mùi,Tân Tị – sách vượng, công cụ nhược; dụng thần sách Thổ, công cụ Thủy (Ấn)
- A-032 P2-405 VD11 (P2 tr.405) – Giáp Ngọ,Nhâm Thân,Tân Sửu,Canh Dần – sách vượng, công cụ vượng; dụng thần sách Hỏa, công cụ Mộc (Tài)
- A-035 P2-407 VD14 ; HTL X-E2 (P2 tr.407-408) – Tân Sửu,Quý Tị,Ất Sửu,Đinh Sửu – sách nhược, công cụ vượng; dụng thần sách Mộc, công cụ Hỏa (Thực thương)
- A-040 P2-413 VD20 (P2 tr.413) – Ất Mão,Giáp Thân,Nhâm Dần,Quý Mão – sách nhược, công cụ vượng; dụng thần sách Kim, công cụ Thổ (Quan sát)
- A-043 P2-416 VD23 (P2 tr.416) – Giáp Ngọ,Kỷ Tị,Tân Mão,Nhâm Thìn – sách nhược, công cụ nhược; dụng thần sách Thổ, công cụ Thủy (Thực thương)
- A-044 P2-417 VD25 (P2 tr.417-418) – Canh Thìn,Mậu Tí,Canh Tuất,Kỷ Mão – sách vượng, công cụ nhược; dụng thần sách Hỏa, công cụ Hỏa (Quan sát)
- A-045 P2-418 VD26 (P2 tr.418-419) – Nhâm Thìn,Đinh Mùi,Tân Mùi,Giáp Ngọ – sách vượng, công cụ nhược; dụng thần sách Mộc, công cụ Thủy (Thực thương)
- A-046 P2-419 VD27 (P2 tr.419-420) – Kỷ Sửu,Bính Dần,Đinh Mão,Nhâm Dần – sách vượng, công cụ vượng; dụng thần sách Kim, công cụ Mộc (Ấn)
- A-049 P2-422 VD31 ; P3 tr.605 VD6 (P2 tr.422-423 (trụ ngày điền Quý Tị theo lịch) ; P3 tr.605-608) – Canh Dần,Mậu Tí,Quý Tị,Nhâm Tí – sách vượng, công cụ vượng; dụng thần sách Thổ, công cụ Hỏa (Tài)
- A-052 P2-426 VD35 (P2 tr.426-427) – Giáp Dần,Nhâm Thân,Mậu Thân,Mậu Ngọ – sách nhược, công cụ nhược; dụng thần sách Hỏa, công cụ Thổ (Tỉ kiếp)
- A-053 P2-427 VD36 (P2 tr.427) – Bính Tí,Quý Tị,Bính Ngọ,Quý Tị – sách vượng, công cụ vượng; dụng thần sách Kim, công cụ Thủy (Quan sát)
- A-054 P2-428 VD37 ; P3 tr.614 VD8 (P2 tr.428 ; P3 tr.614-617) – Ất Mùi,Giáp Thân,Canh Tuất,Quý Mùi – sách vượng, công cụ vượng; dụng thần sách Hỏa, công cụ Mộc (Tài)
- A-056 P2-430 VD40 (P2 tr.430-431) – Bính Ngọ,Mậu Tuất,Nhâm Dần,Bính Ngọ – sách nhược, công cụ nhược; dụng thần sách Kim, công cụ Mộc (Thực thương)