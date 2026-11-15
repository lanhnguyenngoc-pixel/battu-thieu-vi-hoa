# Giai đoạn C – Chạy chéo các quy tắc đã duyệt + đối chiếu giả định (15/11/2026)

Mẫu gộp: `tools/luan_tong_hop.py` – VS-2 + G11 + VS-3 → CĐB-1/3/4 → DT-2a + DT-3 → DV-3 + TL-1; B1, B2.
Thứ tự ứng viên dụng thần **lấy từ workbook** (DUNG_HY_KY dòng 91, 93–95), chỉ tính lại khi VS-2 đổi trạng thái thân.

## 1. Kết quả gộp

| Phần | Bộ | Cũ (6.4) | Sau khi gộp |
|---|---|---|---|
| Vượng suy | 68 lá TVH | 59 | **60** |
| Độ tin cậy VS-3 | 68 lá TVH | – | Cao: 34/35 đúng · Vừa: 8/11 · Thấp: 18/22 |
| Dụng thần | 38 lá P2 tr.397–431 | 16 | **21** |
| Dụng thần | 20 lá P1 tr.181–200 + Ch.23 (bộ kiểm tra) | 11 | **12** |
| Dụng thần | 247 lá sách khác (chỉ để kiểm) | 71 | 69 |
| Đại vận chung | 196 vận | 52 | **141** |
| Đại vận – tài lộc | 35 vận | – | **29** |
| Đại vận – công danh | 18 vận | – | 13 |

Không có quy tắc nào làm hỏng phần khác: đại vận giữ 141/196 khi thay dụng thần mới; vượng suy không đổi khi đổi dụng thần.

## 2. Ảnh hưởng trên 300 lá ngẫu nhiên

| Hạng mục | Số lá / vận |
|---|---|
| Vượng / Nhược | 162 / 138 |
| Độ tin cậy Cao / Vừa / Thấp | 170 / 38 / 92 |
| Cách đặc biệt chuyển về chính cách (CĐB-1/3/4) | 12 lá (còn 11 lá cách đặc biệt) |
| Dụng thần đổi so với 6.4 | **68 lá**: DT-3 điều hầu 57, cách cục 5, VS-2 đổi thân 5, DT-2a 1 |
| Đại vận (2.400 vận): Rất tốt / Tốt / Xấu / Rất xấu | 650 / 634 / 669 / 447 – không còn "Bình thường" hàng loạt |
| Tài lộc theo vận: Tốt / Xấu | 1.485 / 915 |
| Hôn nhân theo vận: Thuận / Trắc trở / Bình thường | 666 / 545 / 1.189 |
| B1 bệnh nặng | 115 lá (trước 138) |
| B2 gần tòng | 10 lá |

## 3. Phát hiện khi chạy chéo

1. **DT-2a ít tác dụng khi dùng đúng thứ tự ứng viên của workbook** (1/300 lá đổi; TVH +1). Mức cải thiện lớn đo ở bước trước (+13 lá sách khác) là do mẫu thử tự tính lại "nhóm nhiều" theo lượng PP2. Đã thử lại: tính "nhóm nhiều" theo PP2 → sách khác 74/247 (+5) nhưng **đổi dụng thần 128/300 lá**, không hơn trên lá TVH → **không đề xuất**; giữ thứ tự của workbook.
2. **DT-3 (điều hầu) là thay đổi lớn nhất**: 57/300 lá đổi dụng thần. Trên lá TVH có điều kiện điều hầu: 6/12 đúng (trước 2/12); sách khác: giảm nhẹ (đúng trường phái TVH, không phải lỗi).
3. **Cách đặc biệt ở sách khác:** 29/247 lá công cụ xếp cách đặc biệt (12 khúc trực) mà sách khác luận chính cách – khác trường phái; các ví dụ cách đặc biệt của chính TVH công cụ nhận đúng hết → giữ.

## 4. Đối chiếu giả định bị các quy tắc mới tác động

| Giả định | Nội dung cũ (tóm tắt) | Xử lý ở 6.5 |
|---|---|---|
| **G11** chọn PP1/PP2 | Theo PP1 khi quy tắc khẳng định / nửa vượng nửa nhược… | **Sửa**: thêm "VS-2 đổi kết luận ở bước 5 → coi như đã phân xử → theo PP1" (đã duyệt) |
| **G21** cảnh báo quy tắc 3 | Chỉ cảnh báo khi không lệnh, ≥2 điều mạnh từ một chi | **Thay bằng** VS-2 bước 5 (kiểm chứng đủ bước, có thể đổi kết luận) |
| **G28** "vượng suy mong manh" | Liệt kê hai nhóm vận đọc ngược | **Thay bằng** VS-3 (độ tin cậy) + dòng "nếu đọc theo PP2…" khi độ tin cậy Thấp |
| **G50** "được lệnh giả" | Cảnh báo, không đổi kết luận | **Gộp** vào ghi chú "cận ranh giới" của VS-2 |
| **G54(a)** so lượng phe khắc khi ≥2 điều / ≤1 điều | Hướng nghiên cứu | **Đã giải quyết** bằng VS-2; (b), (c) còn mở |
| **G25** gần cách tòng | PP2 < 15% | **Sửa** theo B2: thêm điều kiện kết luận cuối là Nhược |
| **G88** bệnh – mức | ≥60% / <30% | **Sửa** theo B1 |
| **G61** P2 tr.282 đ.3–4 cho can đại vận | Đổi xu thế thành "bình thường" | **Thay bằng** DV-3: không đổi kết luận vận (đã thử: hạ về bình thường làm sai 45/55 vận); chỉ còn trong bảng chi tiết |
| **G76** đại vận can 7 – chi 3 | HG tr.68 | **Giữ** – khớp DV-3 (5 năm đầu theo can, 5 năm sau theo chi) |
| **G17, G26** kết luận chủ đề theo vai trò dụng/hỷ/kỵ | – | **Giữ** cách làm; vai trò tính lại theo dụng thần mới (DT-2a, DT-3) |
| **G12** bậc giàu: cách đặc biệt không xếp | – | **Giữ**; 12 lá chuyển về chính cách sẽ được xếp bậc |
| **G13** tố chất "quá vượng = PP2 độc vượng" | – | **Giữ** (dùng lượng; CĐB-3 chỉ bỏ "độc vượng" khỏi cách cục) |
| **G45** bảng điều hầu Cùng Thông Bảo Giám (tầng kinh nghiệm, ngoài sách) | – | **Giữ** ở tầng tham khảo; DT-3 dùng P1 tr.76–79 của TVH (trong sách) |
| **G56** gia tường "tứ trụ thuần thổ" (đối chiếu HTL) | – | **Giữ** – tương thích CĐB-4 (thuần thổ ≥ 5/8 chữ) |
| **G60, G96** dụng thần thái quá, bảng dự bị | – | **Giữ** (dùng lượng) |

Các giả định còn lại (lục thân, thần sát, tính cách, sức khỏe, NGU_HANH_THIEU, lưu niên…) không bị quy tắc mới tác động trực tiếp; chỉ thay đổi đầu vào khi dụng thần hoặc vượng suy đổi.

## 5. Danh sách đưa vào bản 6.5 (giai đoạn D)
VS-2 · G11 sửa · VS-3 · B1 · B2 · CĐB-1 · CĐB-3 · CĐB-4 · DT-2a · DT-3 · DV-3 (ĐB-1 điều hầu, ĐB-2 thương quan gặp quan, mức dụng lên ngôi/khắc dụng) · TL-1 · sửa mẫu KIEM_THU (P1-VD14 giờ Đinh Mùi; vá mẫu tài lộc 15 lá; sửa giới tính mẫu) · bổ sung MAU_THUAN (VD1 PP2 đủ 6 chỗ; VD6 +6 Mùi; các lá sách tự mâu thuẫn về vượng suy, cách đặc biệt) · sửa G21, G25, G28, G50, G54, G61, G88 · giao diện: kết luận + lý do ngắn, câu sách/trang ẩn sau liên kết.
