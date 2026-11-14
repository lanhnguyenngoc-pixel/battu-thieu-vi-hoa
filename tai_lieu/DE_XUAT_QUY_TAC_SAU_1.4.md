# Đề xuất chuyên môn sau bản 1.4.0 – chờ duyệt (14/11/2026)

Các đề xuất dưới đây đổi hoặc thêm quy tắc, nên chưa làm. Mục nào được duyệt sẽ sửa **đồng thời phần mềm và workbook**, rồi chạy lại toàn bộ kiểm thử.

## Đính chính số liệu đã báo

Lượt trước tôi báo "vượng suy khuyến nghị đúng 57/68 (≈84%)". Đọc lại bảng KIEM_THU của workbook 6.3 thì số đúng như sau:

| Phép so với lời giải sách | PP1 | PP2 | Khuyến nghị |
|---|---|---|---|
| Trạng thái vượng / nhược (67 lá sách có nêu) | 58 khớp | 49 khớp | **59 khớp (88%)** |
| Dụng thần, so theo hành (64 lá sách có nêu) | 39 khớp | 34 khớp | **30 khớp (47%)** |

Như vậy chỗ yếu lớn nhất là **chọn dụng thần**, không phải vượng suy.

## Đề xuất

| # | Đề xuất | Sách nói gì | Cách làm nếu duyệt | Ý kiến của tôi |
|---|---|---|---|---|
| 3.1 | **Phân tích 34 lá lệch dụng thần** | Workbook đã ghi nguyên nhân (KIEM_THU dòng 106–117): sách tự tính sai (P2-VD1); công cụ áp hợp hóa mà ví dụ của sách bỏ qua (P1-VD4, VD7); bảng "nhiều nhất" (P1 tr.108) khác với cách sách chọn riêng từng lá (VD8, VD11); công cụ nhận cách đặc biệt mà sách không nhận (P1-VD10, P2 VD4, VD17, VD27, VD34, VD40); sách lấy điều hậu làm dụng (VD8, VD25, VD29, P3 VD4) | Bước 1 (không đổi quy tắc): lập bảng 34 lá × nhóm nguyên nhân. Bước 2: thử trên bản sao 2 thay đổi ứng viên: (a) siết điều kiện thành cách đặc biệt; (b) điều hậu "BẮT BUỘC" được ưu tiên làm dụng. Đo số lá khớp trên 64 lá sách và số kết luận đổi trên 501 lá. Chỉ đề xuất sửa khi có số đo | **Nên làm trước tiên** |
| 3.2 | Rà soát **94 giả định / quy ước kỹ thuật** | Sách nêu tiêu chí, không cho số (ví dụ: mốc 35–45%, ±10 phút, tuổi kết hôn) | Tệp RA_SOAT_GIA_DINH_6.3.xlsx: chuyên gia chọn Giữ / Sửa / Bỏ cho từng mục | Nên làm; cần người có kinh nghiệm luận giải |
| 3.3 | Giờ mặt trời thực theo kinh độ | Quy tắc G10 (P1 tr.42–47): **dùng giờ đồng hồ nơi sinh** | Tùy chọn **mặc định tắt**, ghi rõ "ngoài sách"; thêm ô kinh độ; giờ hiệu chỉnh = giờ đồng hồ + 4 phút × (kinh độ − 15 × múi giờ) | Hạ mức ưu tiên so với lượt trước vì sách đã chọn giờ đồng hồ |
| 3.4 | Tùy chọn Tý muộn (23–24 giờ không đổi ngày) | P1 tr.45: sinh từ 23 giờ tính sang ngày hôm sau – workbook đang làm đúng sách | Tùy chọn mặc định tắt, ghi "ngoài sách" | Thấp; chỉ làm nếu anh cần so với trường phái khác |
| 3.5 | Cách cục theo nguyệt lệnh (8 chính cách) | Chưa xác minh sách có chương riêng (bản PDF là ảnh quét, không tra chữ được) | Anh chỉ trang sách → tôi đề xuất quy tắc cụ thể | Cần nguồn trước |
| 3.6 | Lưu nguyệt (luận theo tháng) | Chưa xác minh sách có nói | Như 3.5 | Cần nguồn trước |
| 3.7 | Hợp hôn hai lá | Bản 1.4.0 đã nhập được ngày sinh vợ/chồng: mục Hôn nhân xét nạp âm, can chi (P1 tr.176–177; P2 tr.487–499) | Nếu muốn so cả hai lá (dụng thần bổ nhau…) cần nguồn sách | Cần nguồn trước |
| 3.8 | Kho lá số có sự kiện thật | — | Dùng trang Đối chiếu thực tế (1.4.0) để ghi; thêm nút xuất thống kê tỉ lệ khớp PP1 / PP2 trên cả kho (không đổi quy tắc) | Làm được ngay nếu anh đồng ý |

## Mục 4 – sản phẩm

| Việc | Trạng thái |
|---|---|
| Sao lưu / khôi phục kho lá số | **Đã làm trong 1.4.0** (Danh sách lá số → Sao lưu kho… / Khôi phục…; khôi phục chỉ thêm, không xóa) |
| Kiểm tra bản cập nhật | Cần một nơi công khai để phần mềm hỏi phiên bản mới (kho mã đang riêng tư). Nếu làm: nút bấm tay, không tự gửi gì đi |
| Ký số bộ cài | Cần anh mua chứng thư: Windows (chứng thư ký mã) và Apple Developer Program (để Apple công chứng bản macOS); đều tính phí hằng năm |
