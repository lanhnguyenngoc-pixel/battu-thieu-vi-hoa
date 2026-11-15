# Giai đoạn D – Workbook 6.5 và phần mềm 1.6.0 (15/11/2026)

Đưa toàn bộ quy tắc đã duyệt ở giai đoạn B, C vào workbook (bản 6.5) và phần mềm (bản 1.6.0); trình bày lại theo quy tắc mới của chủ dự án.

## 1. Quy tắc trình bày (duyệt 15/11/2026)

| Loại nội dung | Cách trình bày |
|---|---|
| Kết luận và nội dung do công cụ luận giải | **Chữ đậm** |
| Trích dẫn lý thuyết của sách (câu sách, số trang) | Chữ thường, **dòng riêng**, ẩn sau liên kết "📖 Căn cứ sách ▾"; bấm mới sổ ra |
| Bản in PDF (không bấm được) | Kết luận đậm; dòng "Căn cứ sách: …" chữ thường ngay bên dưới |

Workbook tách sẵn 3 ô cho mỗi kết luận mới: KẾT LUẬN · LÝ DO · CĂN CỨ SÁCH. Các câu cũ của workbook có số trang lẫn trong câu: phần mềm tự tách ở lớp hiển thị (không đổi giá trị gốc). Nếu cả câu chỉ là câu sách (ví dụ "Tính cách theo sách: • … • …") thì giữ nguyên văn, chữ thường.

## 2. Thay đổi trong workbook 6.5 (tools/step18.py)

| Quy tắc | Ô | Nội dung |
|---|---|---|
| VS-2 | PP1_DINH_TINH C91, C92, C110–C112, dòng 114–119 | Bước 5 kiểm chứng; C115 = kết luận sơ bộ, C117 = mã bước 5, C118 = lý giải, C119 = căn cứ sách |
| G11 + VS-2 | DOI_CHIEU C33 | Không phải sửa công thức: khi bước 5 đổi kết luận, công thức cũ đã cho "theo PP1" (đã kiểm) |
| VS-3 | DOI_CHIEU C35, F35, C46, C52 | Độ tin cậy Cao/Vừa/Thấp; C52 = cách đọc theo phương pháp còn lại khi Thấp |
| B1 | DUNG_HY_KY C176, C178 | Bệnh nặng chỉ khi % PP2 cùng chiều kết luận cuối |
| B2 | KET_LUAN C35 | Gần cách tòng: thêm điều kiện kết luận cuối là nhược |
| G28, G50 | KET_LUAN C42 | Bỏ hai cảnh báo cũ; thay bằng cảnh báo vượng cận ranh giới + cách đọc còn lại |
| CĐB-1/3/4 | DUNG_HY_KY C72/D72, dòng 201–206 | C202 = cách theo điều kiện cơ bản, C205 = lý do không thành cách, C206 = căn cứ |
| DT-2a | DUNG_HY_KY C96–C97, dòng 209–214 | "Có mặt" = can lộ (trừ nhật can) + mọi tàng can gốc |
| DT-3 | DUNG_HY_KY dòng 215–219 | Điều hầu < 20% tổng độ → dụng thần; C218 lý do, C219 căn cứ |
| DV-3 | DAI_VAN dòng 54, 69, 94, 108–127; KET_LUAN C163; DUNG_HY_KY C189 | Điểm từng nửa vận (PP1, PP2), KẾT LUẬN (dòng 121), tóm tắt, vì sao 5 năm đầu/sau, căn cứ, tài lộc, công danh |
| TL-1 | TAI_LOC C36:J38, dòng 87–91 | Kết luận tài lộc từng vận + lý do + căn cứ |
| Giới tính lá mẫu | NHAP_LIEU C22; KIEM_THU cột Y; 1.204 công thức đổi C10 → C22 | Chọn lá mẫu theo tên thì lấy giới tính của lá mẫu (trước đây mọi lá mẫu bị tính là nam) |
| Mẫu tài lộc | KIEM_THU dòng 143–149 | Vá 5 dòng theo vá_mau_tai_loc_6.5.json (đã duyệt 14/11/2026) |
| MAU_THUAN | dòng 6, 169–180 | VD1 PP2 đủ 6 chỗ; 10 dòng quy tắc mới/sửa giả định; 2 dòng lá sách tự mâu thuẫn |

So từng ô với 6.4: ngoài 1.204 công thức chỉ đổi tham chiếu giới tính, có 570 ô khác (trong đó 104 ô là giá trị bảng dữ liệu KHONG_RO_GIO do LibreOffice tính lại).

## 3. Phát hiện và đính chính trong lúc triển khai

| Việc | Kết quả |
|---|---|
| **P1-VD14 – không sửa mẫu.** Kế hoạch ghi "sửa trụ giờ thành Đinh Mùi". Đọc lại sách: P1 tr.196 in rõ giờ **Đinh Dậu**, tr.197 "chi giờ gặp lộc là đắc địa" (Dậu là lộc của Tân). Lá Đinh Mùi là ví dụ khác ở P1 tr.72 (mẫu "P1 tr.72" đã đúng). | Mẫu KIEM_THU đúng, không sửa. Dữ liệu nghiên cứu của tôi nhập nhầm giờ – đã sửa (vuong_suy_mo_rong.json, dai_van_doi_chieu.json). Với giờ đúng, bước 5 (VS-2) vẫn ra Nhược = sách, nên **vượng suy vẫn 60/68** |
| **DT-2a đã có sẵn.** Workbook 6.4 (DUNG_HY_KY C96) vốn đã "ứng viên không có mặt thì lấy ứng viên sau"; chỉ khác cách đếm "có mặt" (6.4 theo lượng, sau khử bì). | 6.5 đổi cách đếm như đã duyệt (can lộ + tàng can gốc). Ảnh hưởng nhỏ: 1/300 lá |
| **CĐB-4 đếm chi đã hóa theo hành hóa** (đúng câu đã duyệt). Mẫu thử giai đoạn C đếm theo bản khí. | Khác mẫu thử ở 4/300 lá ngẫu nhiên (vd Hợi–Mão–Mùi hóa mộc → đủ 5/8, vẫn là Khúc trực); không lá sách nào đổi |
| B1 trên lá chuyển về chính cách | 6.5 xét bệnh–mức (vì đã là chính cách); mẫu thử còn ghi "cách đặc biệt" – 6.5 nhất quán hơn |

## 4. Đối chiếu workbook 6.5 với mẫu thử đã duyệt (481 lá: 300 ngẫu nhiên + mẫu sách)

| Hạng mục | Khác nhau | Nguyên nhân |
|---|---|---|
| Vượng suy, độ tin cậy | 0 | – |
| Cách cục, dụng thần, đại vận | 4 lá | Cách đếm CĐB-4 (mục 3) |
| Bệnh – mức | 16 lá | Lá chuyển về chính cách (mục 3) |
| Gần cách tòng | 8 lá | 6.5 giữ điều kiện cũ "PP2 chính cách, không khí gốc" và thêm "kết luận cuối nhược" đúng như B2 đã duyệt; mẫu thử dùng cách cục của phương pháp khuyến nghị |

Trên các lá sách: kết quả 6.5 trùng mẫu thử → số liệu giai đoạn C giữ nguyên (vượng suy 60/68; dụng thần P2 21/38, P1 và Ch.23 12/20; đại vận 141/196; tài lộc 29/35).

## 5. Phần mềm 1.6.0

| Màn hình | Thay đổi |
|---|---|
| Luận giải chung, Theo lĩnh vực (ô kết luận), Luận giải chi tiết, các trang bài đọc | Kết luận đậm; căn cứ sách tách dòng, ẩn sau "📖 Căn cứ sách ▾" |
| Vận trình – Đại vận (chế độ thường) | Đầu trang là **Kết luận đại vận**: mỗi vận một mức (rất tốt … rất xấu), tóm tắt 10 năm, vì sao 5 năm đầu / 5 năm sau, tài lộc, công danh, căn cứ sách ẩn. Bảng chấm điểm chi tiết chỉ còn ở chế độ chuyên gia |
| Lá số – thẻ đại vận | Hiện mức kết luận và tóm tắt; rê chuột xem lý do |
| PDF | Kết luận đậm, "Căn cứ sách: …" chữ thường ở dòng dưới |
