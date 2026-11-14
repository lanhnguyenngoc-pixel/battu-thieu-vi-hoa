# ĐỀ XUẤT BỐ CỤC GIAO DIỆN – HIỆN / ẨN (chờ duyệt)

Mục tiêu: một phần mềm phục vụ được ba nhóm người – **người đọc lá số** (chỉ cần hiểu kết quả), **người học** (cần hiểu vì sao),
**chuyên gia luận giải** (cần mọi căn cứ). Nguyên tắc: không đổi bất kỳ quy tắc/kết quả nào; chỉ đổi cách trình bày.
Mọi sheet vẫn được tính đầy đủ trong lõi và vẫn kiểm thử 100%.

## 1. Ba tầng hiển thị

| Tầng | Ai dùng | Hiện gì |
|---|---|---|
| **A. Đọc lá số** (mặc định) | Người đọc, người mới học | Tóm tắt · Lá số trực quan · Luận giải · Theo lĩnh vực · Vận trình (biểu đồ + năm) · Ngũ hành thiếu (biểu đồ) |
| **B. Chuyên gia** (bật bằng công tắc "Chế độ chuyên gia") | Chuyên gia, người nghiên cứu | Thêm: Phân tích (PP1, PP2, Thông căn, Đối chiếu, Dụng–hỷ–kỵ), Chuyên đề 10 sheet, Không rõ giờ, Trung gian, bảng Lưu niên 100 năm, Theo dõi; địa chỉ ô Excel khi rê chuột |
| **C. Tham khảo** | Mọi người | Hướng dẫn sử dụng, HUONG_DAN của workbook, Tham khảo ngoài sách, Tiết khí |

## 2. Bảng chi tiết từng phân hệ

| Phân hệ / sheet | Đề xuất | Lý do |
|---|---|---|
| KET_LUAN dòng 3–12 (TÓM TẮT) | **Hiện – màn hình đầu** | Câu trả lời ngắn nhất cho mọi người |
| KET_LUAN mục B–K | **Hiện** (tầng A) | Mỗi lĩnh vực: KẾT LUẬN trước, căn cứ sau |
| KET_LUAN dòng 200–251, cột ẩn | **Ẩn** (đã làm) | Vùng phụ trợ, ô tính phụ |
| LUAN_GIAI 8 mục + tổng thể | **Hiện** (tầng A), trình bày dạng thẻ đọc liền mạch thay vì lưới ô | Dành cho người đọc; cột phụ J:U ẩn (đã làm) |
| DOC_NHANH | **Hiện** – gộp vào Tóm tắt thành thẻ "Đọc nhanh" | Bản rút gọn cho chủ lá số |
| Lá số trực quan (từ KET_LUAN mục B + TRUNG_GIAN) | **Hiện** (tầng A): 4 trụ can chi, thập thần, tàng can, nạp âm, màu ngũ hành | Người học nhìn lá số như trong sách |
| TRUNG_GIAN (lưới ô) | **Chuyên gia** | Số liệu trung gian, không cần cho người đọc |
| PP1_DINH_TINH, PP2_DINH_LUONG, THONG_CAN, DOI_CHIEU, DUNG_HY_KY | **Chuyên gia**; kết quả chính đã có ở mục C | Căn cứ chi tiết |
| DAI_VAN, NGU_HANH_THIEU | **Hiện** biểu đồ + bảng tóm (tầng A); lưới đầy đủ ở chuyên gia | Biểu đồ dễ hiểu cho mọi người |
| LUU_NIEN (100 năm, 112 cột), THEO_DOI | **Chuyên gia**; tầng A chỉ hiện 10 năm từ năm xem (LUAN_GIAI mục 7) | Bảng quá rộng cho người đọc |
| LUC_THAN, HON_NHAN, MENH_CUC, TAI_LOC, TO_CHAT, SU_NGHIEP, SUC_KHOE, THAN_SAT, SAO, BAO_PHU | **Chuyên gia**; kết luận đã có ở mục D–H | Căn cứ chi tiết |
| KHONG_RO_GIO | **Hiện khi bật "Không rõ giờ"**; ẩn khi biết giờ | Chỉ có nghĩa khi không biết giờ |
| SU_KIEN, THEO_DOI (nhập sự việc) | **Hiện** dưới tên "Kiểm chứng sự việc" – cho phép nhập năm + loại sự việc | Hiện chỉ đọc; nên cho nhập như workbook |
| NHAP_LIEU (lưới ô) | **Ẩn** – đã thay bằng biểu mẫu nhập | Tránh trùng |
| TRANG_CHU | **Ẩn** – nội dung đã có ở Tóm tắt + Hướng dẫn | Trùng lặp |
| HUONG_DAN, THAM_KHAO, TIET_KHI | **Hiện** ở Tham khảo | Tra cứu |
| MAU_THUAN, KIEM_THU, CSDL | **Ẩn** (đã làm – quyết định 14/11) | Nội bộ |
| Bảng dữ liệu KHONG_RO_GIO (288 ô) | **Ẩn** | Vòng 12 giờ chạy ngầm |

## 3. Cải tiến trình bày đề xuất thêm (không đổi quy tắc)

| # | Đề xuất | Lợi ích |
|---|---|---|
| 1 | Công tắc **Chế độ chuyên gia** ở thanh bên (nhớ lựa chọn) | Người mới không bị ngợp; chuyên gia có đủ |
| 2 | Màn hình **Lá số trực quan** (4 trụ lớn, màu ngũ hành, thập thần, tàng can, đại vận dạng dải ngang) | Đẹp, giống lá số trong sách |
| 3 | **Luận giải dạng bài đọc** (thẻ có tiêu đề, chữ đỏ đậm theo cờ) thay vì lưới ô | Đọc liền mạch |
| 4 | **Chú giải thuật ngữ** khi rê chuột (thập thần, dụng thần, nhược/vượng…) lấy từ HUONG_DAN | Người học hiểu ngay |
| 5 | Thanh **"Đang xem: tên · tứ trụ · năm xem"** cố định trên mọi trang + nút đổi năm xem nhanh | Luôn biết đang xem lá nào |
| 6 | **Chế độ tối** và cỡ chữ lớn | Dễ đọc lâu |

Nếu anh duyệt, tôi làm theo đúng bảng trên (các mục cần duyệt riêng: #2, #3, #4 vì thêm cách trình bày mới; SU_KIEN cho nhập).
