# Giai đoạn B – Đối chiếu đại vận theo can chi (15/11/2026)

## 1. Cách làm
- 54 lá có lời sách về đại vận (P1 tr.181–200, P2 tr.397–431, Ch.23 tr.583–626). Nhãn từng vận (Tốt/Xấu/Bình thường) đọc theo **can chi** của vận: lời sách nêu tên vận, thập thần của vận, hoặc tuổi (đổi sang vận bằng tuổi khởi vận in trong bảng đại vận của sách). Nhãn: `nhan_dai_van_sach.py` (196 vận).
- Trạng thái thân: VS-2 + G11 (đã duyệt).

## 2. Kết quả công cụ hiện tại (DAI_VAN dòng 54/69)

| Sách \ Công cụ | Tốt | Xấu | Bình thường |
|---|---|---|---|
| Tốt (106) | 25 | 10 | **71** |
| Xấu (78) | 6 | 17 | **55** |
| Bình thường (12) | 0 | 2 | 10 |

Khớp **52/196 (27%)**. Lỗi chính không phải đoán ngược (16 vận) mà là **126 vận sách nói rõ tốt/xấu thì công cụ ghi "Bình thường"**:
- 67 vận: hành can vận không phải dụng/hỷ/kỵ → "bình thường" (quy tắc 3, P2 tr.281).
- ~55 vận: điều chỉnh đ.3/đ.4 (P2 tr.282: vận tốt bị mệnh cục khắc/hợp → "tốt không thực"; vận xấu bị chế → "xấu vừa") hạ về bình thường; sách vẫn giữ tốt/xấu ở 45/55 vận này.

Lưu ý nguồn: P2 tr.281 nói rõ các quy tắc này "ghi lại các điều được bàn trong các sách Bát tự tinh giải, Khán lưu niên đại vận… để độc giả tham khảo"; còn cách tác giả tự dùng khi luận ví dụ là **P1 tr.124–125**:
- "Thời gian tốt nhất là lúc dụng thần lên ngôi" (tr.124).
- Thân vượng: vận quan, sát, tài, thực, thương là vận tốt; vận ấn, tỉ kiếp là bại vận; vận khắc dụng thần là hung vận. Thân nhược: ngược lại (tr.125).

## 3. Thử cách của tác giả (P1 tr.125) – không chỉnh ngưỡng

| Cách | Khớp |
|---|---|
| Công cụ hiện tại | 52/196 (27%) |
| P1 tr.125 theo can vận (thân nhược: ấn, tỉ kiếp tốt; thân vượng: thực thương, tài, quan sát tốt) | **135/196 (69%)** |
| Như trên nhưng dùng vượng suy của sách | 144/196 (73%) |
| Thêm "dụng thần lên ngôi / vận khắc dụng" theo dụng thần công cụ | 130/196 |
| Thêm dụng thần ở can hoặc chi | 127/196 |

## 4. 49 vận còn lệch (cách P1 tr.125) – nhóm nguyên nhân
| Nhóm | Ví dụ | Hướng xử lý |
|---|---|---|
| Vượng suy khác sách (các lá MÂU THUẪN đã ghi) | VD2 tr.398, VD14 tr.408, VD20 tr.413, VD25 tr.417 | Giữ quy tắc đã thống nhất |
| Dụng thần điều hầu (sách lấy hỏa sưởi ấm) | VD8 tr.403: vận Bính Ngọ, Đinh Mùi tốt dù thân nhược | Nghiên cứu: khi dụng thần là điều hầu thì vận dụng lên ngôi tốt |
| Lời luận theo lĩnh vực | P1-VD14 "vận sát học giỏi"; VD23 "tài vận 41–55 tốt" | Đánh giá vận theo từng lĩnh vực (mẫu `luan_dai_van.py`) |
| Chi của vận (5 năm sau) mạnh | VD16 tr.410 Bính Tí, Ất Hợi (chi tỉ kiếp giúp thân nhược); VD3 tr.399 Canh Dần, Tân Mão (chi tài) | Hiện cả hai nửa: can 5 năm đầu, chi 5 năm sau (P1 tr.122; P2 tr.281) |
| Khác | VD37 tr.616 (6 vận), C23-595 | Đọc lại |

## 5. Đề xuất DV-1 (chờ duyệt)
1. Xu thế đại vận lấy theo **P1 tr.124–125**: 5 năm đầu theo can vận, 5 năm sau theo chi vận (bản khí); thân vượng → thực thương/tài/quan sát tốt, ấn/tỉ kiếp xấu; thân nhược ngược lại.
2. Mức: dụng thần lên ngôi → "Rất tốt"; vận khắc dụng thần → "Rất xấu" (P1 tr.124–125).
3. Quy tắc P2 tr.281–282 (dụng/hỷ/kỵ/nhàn; đ.1–đ.4 xung, khắc, hợp) **giữ làm ghi chú** ("tốt không thực", "xấu giảm") – không còn đổi kết luận thành "Bình thường".
4. Lý giải từng bước cho mỗi vận: thân → thập thần can/chi → thuận/nghịch (câu sách, trang) → mức → ghi chú P2 tr.282.

## 6. Quyết định (15/11/2026) và mẫu thử
- Duyệt DV-1, điều chỉnh mục 3: quy tắc P2 tr.281–282 hiện thành dòng **"Tham khảo thêm: có thể tốt bị suy giảm – lý do …"**, **"Tham khảo thêm: xấu giảm – lý do …"** (kèm đ.1–đ.4, trang), không đổi kết luận. Quy tắc này dùng cho cả mục C1 (dụng/hỷ/kỵ PP1) và C2 (PP2) – "P2" là Phần 2 của sách, không phải phương pháp 2.
- Mẫu thử `tools/dai_van_dv1.py` (`luan_la`): mỗi vận có 3 bước lý giải (thân → can 5 năm đầu → chi 5 năm sau) + mức Rất tốt/Tốt/Xấu/Rất xấu + ghi chú tham khảo. Khớp sách 135/196 (công cụ cũ 52/196). Khi VS-2 đổi trạng thái thân so với kết luận cũ, chưa chấm mức "dụng lên ngôi / khắc dụng" (dụng thần cần tính lại theo trạng thái mới ở 6.5).

## 7. Xét tiếp 61 vận còn lệch (15/11/2026)

### 7.1 Hai nửa vận
88 vận có nửa đầu (can) và nửa sau (chi) khác chiều: sách theo can 60 vận, theo chi 28 vận → giữ can là chính cho cả vận (P2 tr.281: 5 năm đầu can chủ 70%), nhưng hiện rõ hai nửa ("Tốt (đầu) – Xấu (sau)"). Các lời sách nêu tuổi rơi đúng nửa sau (VD7 tr.402 "44–49 bốc nhanh", C23-623 "51–55 huy hoàng") đều khớp nửa sau của DV-1.

### 7.2 Bốn hướng bổ sung đã thử – đều KHÔNG cải thiện → không đề xuất
| Thử | Kết quả (DV-1 = 135/196) | Lý do |
|---|---|---|
| Vận đúng hành điều hầu (DUNG_HY_KY C162) là tốt | 132–136 | Điều hầu sách dùng ở lá nào phụ thuộc cách chọn dụng thần (VD8 tháng Dần công cụ không xét điều hầu) |
| Thân nhược, ấn là dụng: vận quan sát (sinh ấn) là tốt | 134 | Sách không nhất quán: 7 vận tốt (P1-VD14 Đinh Tị, P1-VD15 Tân Tị, VD8 Bính Ngọ…) – 8 vận xấu (P1-VD2 Quý Hợi, Nhâm Tuất; VD13 Quý Sửu…) |
| Sự kiện xung/hình/hợp (P2 tr.282 đ.6–8) làm vận thành xấu | – | Vận tốt có thiên khắc địa xung: sách vẫn nói tốt 21, xấu 4; có hình: tốt 29, xấu 8 → chỉ là sự kiện trong vận, không đổi xu thế |
| Dụng thần lên ngôi đổi kết luận | 130 | Dụng thần công cụ khác sách ở một số lá |

### 7.3 21 vận lệch thật (hai nửa cùng chiều mà sách ngược)
| Nhóm | Lá (vận) | Hướng |
|---|---|---|
| Vượng suy khác sách (MÂU THUẪN đã ghi) | VD14 tr.408 (Kỷ Sửu), VD25 tr.417 (Đinh Hợi), VD37 tr.616 (Nhâm Ngọ, Canh Thìn – lời luận vận của sách theo kiểu thân nhược) | Giữ quy tắc |
| Sách chọn dụng thần đặc biệt | VD8 tr.403 (điều hầu hỏa – Bính Ngọ), VD26 tr.418 (thổ nhiều kim vùi lấp → Nhâm Tí), VD35 tr.426 (dụng Bính hỏa, hỷ mộc – Ất Hợi), C23-583 (Quý Mão "dụng lên ngôi") | Việc 2 đã duyệt: đối chiếu dụng thần theo can ngày + bốn mùa (P1 tr.74–80) |
| Quan sát sinh ấn (sách không nhất quán) | P1-VD14 Đinh Tị "vận sát học giỏi", P1-VD15 Tân Tị, VD32 Quý Dậu | Ghi MÂU THUẪN; luận theo lĩnh vực (công danh, học: "quan ấn tương sinh" P2 tr.537) |
| Sự kiện trong vận | P1-VD8 Mậu Dần "vận xung nguy", P1-VD9 Canh Tí "thương quan gặp quan nhiều tai họa", P1-VD11 Tân Mùi/Canh Ngọ "bị tù", VD27 tr.420 Nhâm Tuất "sóng gió" | Hiện ở dòng sự kiện của vận (DAI_VAN mục D) cạnh xu thế |
| Lời luận theo lĩnh vực | VD23 tr.416 Ất Hợi "tài vận tốt", VD38 tr.429 Đinh Tị, Mậu Ngọ "phát tài to", P1-VD5 Bính Dần (sức khỏe) | Luận đại vận riêng cho từng lĩnh vực |

## 8. Đại vận theo từng lĩnh vực (15/11/2026)

Nhãn theo lĩnh vực: chỉ các câu sách nói rõ về tài lộc / công danh của một vận (`NHAN_LV` trong `nhan_dai_van_sach.py`): 35 vận tài lộc, 18 vận công danh, 1 vận học.
So hai cách: (1) xu thế chung DV-1; (2) quy tắc riêng từng lĩnh vực (mẫu `luan_dai_van.py`: "vận tài thân gánh được → phát tài", "thân nhược gặp quan vận → bị đè nén", "thương quan gặp quan → mất chức", "vận tài phá ấn → mất chức"…).

| Lĩnh vực | DV-1 chung (theo can) | DV-1 chung (một nửa tốt là tốt) | Quy tắc riêng lĩnh vực |
|---|---|---|---|
| Tài lộc (35) | 24 | **26** | 22 |
| Công danh (18) | **11** | 8 | **11** |

Quy tắc riêng **không** hơn xu thế chung. Các vận còn lệch rơi vào nhóm đã biết: vượng suy khác sách (VD14, VD25), dụng thần điều hầu (VD8 Bính Ngọ, Đinh Mùi), "quan ấn tương sinh" (P1-VD14 "vận sát học giỏi", P1-VD15 Tân Tị), và tài lộc nhờ **chi của vận** (VD3 Canh Dần, Tân Mão: chi là tài; VD23 Giáp Tuất: chi Tuất là ấn giúp thân gánh tài).

**Đề xuất DV-2 (chờ duyệt):**
1. Kết luận từng vận cho mọi lĩnh vực **lấy theo DV-1** (một kết luận chung, hai nửa).
2. Mỗi lĩnh vực thêm **dòng ghi chú kèm câu sách** (tài lộc, công danh, hôn nhân – bộ câu trong `luan_dai_van.py`), ví dụ "Tài lộc: vận tài, thân gánh được → phát tài (P2 tr.522)"; không đổi kết luận chung.
3. Riêng tài lộc: nếu **chi của vận** (5 năm sau) là tài hoặc giúp thân gánh tài, ghi "5 năm sau thuận cho tài" (VD3 tr.399, VD23 tr.416).

## 9. Thay DV-2 bằng DV-3: quy tắc chung + quy tắc riêng → MỘT kết luận (15/11/2026)

Theo góp ý người dùng: không liệt kê ghi chú rời; khi lá số có điều kiện đặc biệt thì vận dụng quy tắc riêng **vào chính kết luận**. Mẫu `tools/dai_van_dv3.py`.

Thứ tự xét mỗi nửa vận (can 5 năm đầu, chi 5 năm sau):
1. Quy tắc chung P1 tr.125: thuận/nghịch với thân (VS-2 + G11).
2. Quy tắc riêng – chỉ áp khi lá số có điều kiện:
   - **ĐB-1 Điều hầu** (P1 tr.76–79): kim, thủy, mộc sinh tháng Hợi, Tí, Sửu, Dần cần hỏa; hỏa, thổ, kim sinh tháng Tị, Ngọ, Mùi cần thủy → vận đúng hành điều hầu chuyển thành tốt. Thử: 10 vận đúng, 3 vận sai → +5 vận.
   - **ĐB-2 Thương quan gặp quan** (P2 tr.541–542 "kị nhất"): vận thương quan, trụ có chính quan lộ, không có ấn → xấu. Thử: +1 vận.
   - Đã thử và **không** đưa vào vì không cải thiện: quan sát sinh ấn (7 đúng / 6 sai), xung/hình/hợp đổi xu thế, P2 tr.282 đ.3–đ.4.
3. Mức: dụng thần lên ngôi → rất tốt; khắc dụng thần → rất xấu (P1 tr.124–125).
4. Kết luận vận: "TỐT/XẤU – cả 10 năm …" hoặc "5 năm đầu …, 5 năm sau …", kèm 2 dòng "Vì: …".

Khớp sách: **141/196** (DV-1: 135; công cụ cũ: 52). Ví dụ VD8 tr.403: Bính Ngọ TỐT, Đinh Mùi RẤT TỐT, Mậu Thân RẤT TỐT – sách: "sau 34 (Bính Ngọ) khá… tốt nhất từ 54 (Mậu Thân)".

Các dòng "tham khảo thêm" của DV-1 (P2 tr.282) và ghi chú ứng viên dụng thần trong đại vận **bỏ khỏi phần kết luận vận** (chỉ giữ trong bảng chi tiết cho người muốn xem).

### 9.1 Quyết định (15/11/2026)
- Duyệt **DV-3** (thay DV-2) và **DT-3** (điều hầu ĐB-1 đưa vào chọn dụng thần).
- Yêu cầu giao diện (giai đoạn D): phần trích dẫn câu sách/trang **ẩn sau một liên kết**; người dùng bấm vào thì mở ra xem chi tiết. Phần hiển thị chính chỉ có kết luận và lý do ngắn.

## 10. Kết luận theo từng lĩnh vực trong mỗi vận (cùng cách DV-3) – 15/11/2026

Mẫu `dai_van_dv3.linh_vuc()`: mỗi vận cho **một** kết luận cho Tài lộc, Công danh, Hôn nhân + lý do một dòng; câu sách để ở trường `nguon` (ẩn, bấm mới xem).

| Lĩnh vực | Cách làm | Khớp sách |
|---|---|---|
| Tài lộc (35 vận) | Kết luận chung của vận, áp quy tắc riêng khi có điều kiện: **TL-1** thân vượng + chi vận là tài → có tài (P2 tr.522); thân nhược + can vận là tài nhưng chi vận giúp thân → có thể tụ tài (P2 tr.521, VD23 tr.416); thân vượng + vận tỉ kiếp → hao tài (P2 tr.527); thân nhược + vận tài → khó giữ (P2 tr.527) | 25 → **29** |
| Công danh (18 vận) | Kết luận chung (đã gồm ĐB-2 thương quan gặp quan). Quy tắc riêng "vận quan sát với thân vượng/nhược" thử ra kém hơn (11 so với 13) → không dùng | 13 |
| Hôn nhân | Chưa có nhãn theo vận để đo; dùng các dấu hiệu sự việc của sách: vận thương quan (P2 tr.471), chi vận xung cung phối ngẫu (P2 tr.462), vận kình dương (P2 tr.468), gặp sao phối ngẫu (P1 tr.153–157) → THUẬN / TRẮC TRỞ / BÌNH THƯỜNG | – |

Ví dụ VD23 tr.416 (thân nhược, dụng Thổ), vận Giáp Tuất từ 41 tuổi: kết luận chung RẤT XẤU (5 năm đầu rất xấu, 5 năm sau rất tốt); **Tài lộc: TỐT** – vận tài nhưng chi vận (Tuất – ấn) giúp thân gánh được tài → có thể tụ tài. Sách: "tài vận từ 41–55 tuổi tương đối tốt".
