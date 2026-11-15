# Quy trình luận giải theo cách của tác giả Thiệu Vĩ Hoa (bản nháp – 14/11/2026)
Nguyên tắc: học CÁCH tác giả suy luận từ ví dụ + lý thuyết; mỗi bước có câu sách và trang; KHÔNG chỉnh ngưỡng để khớp ví dụ. Đầu ra luôn kèm LÝ GIẢI TỪNG BƯỚC (điều thấy trong lá số → câu sách căn cứ → kết luận của bước).

## Mẫu đầu ra cho mỗi lĩnh vực
1. Điều kiện nền (thân, dụng thần) · 2. Sức sao chủ sự · 3. Thân có gánh được không · 4. Vai trò dụng/hỷ/kỵ · 5. Nguồn sinh – kẻ phá – tổ hợp · 6. Gộp khi trái nhau (ghi rõ điều nào thắng và vì sao) · 7. Kết luận mức gốc · 8. Giai đoạn theo đại vận · 9. Hình thức. Mỗi bước: "Lá số có… → Sách: '…' (trang) → nên…".

## TÀI LỘC (đã có nguyên mẫu: tools/luan_tai_loc.py)
Căn cứ chính: P2 tr.520–522 (dự đoán tài vận 1), tr.526–527 (tài vận 2), P1 tr.164–166; ví dụ VD1–VD40, Ch.18, Ch.23.
- B2 Sức tài: có/không; lộ; gốc (bản khí; trung/dư khí chỉ tính khi chi chủ không khắc và chi không nằm trong cục khắc tài – rút từ VD17 "không có khí gốc"); được lệnh; kho.
- B3 Thân gánh tài: "thân vượng tài vượng là phúc"; "tài nhiều thân nhược… nhà nhiều của ít".
- B4 Vai trò: "tài làm dụng thần là người giàu có"; tài kỵ → hạ 1 bậc nhưng không dưới Vừa ("phải gặp vận thân vượng mới tụ tài" VD23).
- B5 Sinh/phá/tổ hợp: thực thương sinh tài (+); tỉ kiếp đoạt (−); kiêu ấn gặp tài không tỉ kiếp (−, tr.527); thực thương quá nhiều tài không hiện (−, VD34); đủ tài quan ấn (+, tr.521); kho giữ được của (+).
- B6 Thứ tự khi trái nhau: gốc của tài > cân bằng thân–tài > vai trò > sinh/phá.
- Kiểm tra trên 65 câu sách (không chỉnh ngưỡng): khớp 35/65 (phần mềm hiện tại 20/65).

## QUAN VẬN (làm tiếp)
Căn cứ: P2 tr.531–544 (QV1–QV65 đã có trong công cụ), P1 tr.167–173 (sang hèn); ví dụ VD1, VD2, VD8, VD12–VD17, VD19–VD21, VD24–VD40, Ch.19, Ch.23.
- B2 Sức quan sát: có/không; lộ hay tàng ("quan tàng không lộ → không có chức to" VD16, VD25); ở lệnh; ở tử/mộ/tuyệt/thai ("quan ở đất thai không vượng" VD31; "sát ở tuyệt địa đóng mộ kho" VD17).
- B3 Thân gánh quan: "thân vượng có thể thắng tài quan" (VD12); thân nhược quan nhiều → bị đè nén (Ch.19 VD7).
- B4 Vai trò quan sát (dụng/kỵ).
- B5 Phá/cứu: thương quan gặp quan (VD1, Ch.19 VD8–9); quan bị hợp (VD27, Ch.19 VD8); quan sát hỗn tạp (VD13); thực chế sát thái quá (VD8, VD21); cứu: ấn chế thương ("thương quan có chế hóa là có quyền" Ch.19 VD3), sát ấn tương sinh (VD29, P1-VD15), kình dương + sát (Ch.19 VD1), khôi canh (VD37).
- B6 Thứ tự: bị phá (thương gặp quan, hợp, hỗn tạp) > sức quan > vai trò; có cứu thì giảm mức phá.
- Mức: không làm quan / chức nhỏ, chức phó / cán bộ vừa / có quyền lớn (từ ngữ tác giả dùng).

## HÔN NHÂN (làm tiếp)
Căn cứ: P2 tr.462–499 (HB1–HB23, HN14–HN71 đã có), P1 tr.153–159; ví dụ Ch.15 (35), VD4, VD9, VD22, VD25, Ch.23.
- B2 Sao phối ngẫu: có/không, lộ, sức (nam: tài; nữ: quan sát).
- B3 Cung phối ngẫu (chi ngày): bị xung/hình, tự khắc, âm dương lệch, khôi canh.
- B4 Dấu hiệu khắc phối ngẫu ở mức RÕ (như ví dụ): thương quan lộ; tỉ kiếp trùng trùng; kình dương; nữ: quan sát hỗn tạp, không quan, kiêu ấn trùng, thuần âm, đủ thìn tuất sửu mùi; nam: chính + thiên tài, không tài, tài mộ/tuyệt.
- B5 Dấu hiệu tốt / cứu: tài quan ấn đủ (nữ), nhật đóng tài/quý nhân, phối ngẫu là dụng; trụ năm hai người tương sinh (Ch.15 VD34).
- B6 Thứ tự: số dấu hiệu khắc rõ độc lập (≥2 → hôn nhân không thuận / kết hôn 2 lần) > cứu.
- Thời điểm: vận/năm thương quan, tỉ kiếp, kình dương, xung cung phối ngẫu (ví dụ Ch.15 đều nêu năm ly hôn).

## Phát hiện phụ
- Lá mẫu chọn theo tên không đặt giới tính → luôn tính nam (đã sửa trong đo; cần sửa trong workbook 6.5).
- Thay dụng thần của công cụ bằng dụng thần của sách KHÔNG làm tài lộc khớp hơn (18 so với 20/43) → khâu quyết định là quy trình lĩnh vực, không phải dụng thần.

## Cập nhật 15/11/2026 – nguyên mẫu 3 lĩnh vực (tools/luan_tai_loc.py, luan_quan_van.py, luan_hon_nhan.py)
Kiểm tra trên câu sách, KHÔNG chỉnh ngưỡng: tài lộc 35/65 (hiện 20), quan vận 27/54 (hiện 16), hôn nhân 35/56 (hiện 13).
Điều học thêm từ ví dụ:
- Quan vận: "thương quan gặp quan" chỉ xấu khi gặp CHÍNH QUAN mà ấn không đứng cùng/kề trụ thương quan (VD1 tr.398 so với P1-VD8 tr.191); thương quan gặp THẤT SÁT là quý (tr.534, 555); kình dương + sát lộ là quyền (tr.532–533); quan lộ không gốc → chức không lớn; quan lộ trùng → "quan nhiều là sát… chức phó" (VD36).
- Hôn nhân: hai mặt tách riêng – (a) độ bền hôn nhân theo tiêu chí khắc; (b) phẩm chất người phối ngẫu (giàu, đẹp, hiền) – điều tốt KHÔNG xóa được dấu hiệu khắc (C23-623). Kém khi ≥2 tiêu chí (tính cả cung phối ngẫu) hoặc 1 tiêu chí rất nặng; 1 tiêu chí → trắc trở, thành biến cố khi gặp vận/năm kích hoạt.
- Câu hỏi mở: theo cách đọc sát câu sách, 54% lá ngẫu nhiên có hôn nhân "không thuận" (phần mềm hiện tại 17%).
