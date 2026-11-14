# TIẾN ĐỘ – PHẦN MỀM BÁT TỰ THIỆU VĨ HOA (theo workbook 6.3)

Cập nhật: 14/11/2026 · Giai đoạn hiện tại: **GĐ6 – Xuất, đóng gói, kiểm thử tổng, hướng dẫn: XONG** · Phiên bản phần mềm **1.3.0** (quy tắc 6.3) · Nguồn: workbook SHA-256 `7ae75965…bca4c5`

## Quyết định đã duyệt
- 13/11/2026: MT-01 … MT-11 (KE_HOACH_GIAI_DOAN_0.md).
- 14/11/2026: MT-12 (tự điền C32/C33 khi nhập âm lịch), MT-13 (Không rõ giờ = vòng 12 giờ), MT-14 (điền tứ trụ tự tính vào C7:F8).
- 14/11/2026: **T8 chuẩn hóa** – bộ lá T8 là lá ngày giờ dương lịch, tứ trụ theo lịch tiết khí của bộ công cụ (tests/t8_chuan.json, 55 lá từ data39 + dv16); chỉ so phần mềm với workbook; bỏ phần đối chiếu số trang và ghi chú lệch trang. **MT-15 khép lại** – không sửa workbook. Tệp gốc 09_TAI_LIEU_BO_SUNG giữ nguyên.

## Đã làm
| GĐ | Việc | Kết quả |
|---|---|---|
| 1 | Khung, dữ liệu nền, thư viện hàm LibreOffice, bộ dịch công thức, lõi tính, nhập liệu, âm lịch, oracle mở rộng, giao diện Nhập liệu + Trung gian | Xong |
| 2 | PP1, PP2, THONG_CAN, DOI_CHIEU, DUNG_HY_KY; tính toán mảng; hiển thị theo định dạng ô; trang Phân tích | Xong |
| 3 | DAI_VAN, LUU_NIEN, THEO_DOI, NGU_HANH_THIEU; làm tròn đúng LibreOffice; 2 biểu đồ; trang Vận trình, Ngũ hành thiếu | Xong |
| 4 | Dịch 11 sheet: LUC_THAN 1.673, HON_NHAN 1.198, MENH_CUC 319, TAI_LOC 283, TO_CHAT 903, SU_NGHIEP 1.304, SUC_KHOE 884, THAN_SAT 511, SAO 352, BAO_PHU 42, KHONG_RO_GIO 519 = **7.988 ô** | Xong |
| 4 | **Bảng dữ liệu Excel → vòng 12 giờ** (`engine.py`): ô (r, c) của 12 bảng = công thức đầu cột c tính với KHONG_RO_GIO!C6 = giờ ở cột B dòng r; mỗi giờ một lõi tính phụ có nhớ đệm. **288 ô bảng × 30 lá trùng đáp án** ở cả chế độ Biết giờ và Không rõ giờ | Xong |
| 4 | Tối ưu (không đổi kết quả): `tools/phu_thuoc.py` tính trước tập ô phụ thuộc C6; lõi phụ chỉ tính ô đổi theo giờ, và nếu mọi ô tham chiếu trực tiếp C6 trùng lõi chính thì dùng lại toàn bộ (chế độ Biết giờ) | Xong |
| 4 | Giao diện: trang **Chuyên đề** (10 thẻ) và **Không rõ giờ** (chọn giờ đang xem, đồng bộ màn hình nhập; bảng 12 giờ, kết luận, tìm giờ) | Xong (ảnh `anh_gd4/`) |
| 4 | Oracle: T8 chuẩn (55 lá × PA1/PA2, ghi mọi sheet), 501 lá cho 10 sheet GĐ4 | Xong |
| 5 | Dịch 6 sheet cuối: KET_LUAN 537, LUAN_GIAI 950, DOC_NHANH 13, SU_KIEN 421, TRANG_CHU 4, THAM_KHAO 12 – **toàn bộ 33 sheet** | Xong |
| 5 | Định dạng có điều kiện: 17 quy tắc dịch nguyên văn (`tools/dich_dinh_dang_dk.py`) | Xong |
| 5 | Giao diện: **TÓM TẮT là màn hình đầu**; Theo lĩnh vực B–K (ẩn dòng 200–251) + LUAN_GIAI, DOC_NHANH, SU_KIEN + tìm kiếm; Tham khảo (TRANG_CHU, HUONG_DAN, THAM_KHAO, TIET_KHI); Danh sách lá số SQLite (lưu/sửa/xóa/tìm/nhóm/ghi chú/ghi đè); So sánh hai lá | Xong (ảnh `anh_gd5/`) |

## Mức khớp (phần mềm so đáp án)
| Sheet | Ô | T1 (30 lá) | T1-501 (501 lá) |
|---|---|---|---|
| 13 sheet GĐ1–3 | 18.568 | 559.320 / 559.320 | 9.340.644 / 9.340.644 |
| LUC_THAN | 1.673 | 50.190 / 50.190 | 838.173 / 838.173 |
| HON_NHAN | 1.198 | 35.940 / 35.940 | 600.198 / 600.198 |
| MENH_CUC | 319 | 9.570 / 9.570 | 159.819 / 159.819 |
| TAI_LOC | 283 | 8.490 / 8.490 | 141.783 / 141.783 |
| TO_CHAT | 903 | 27.090 / 27.090 | 452.403 / 452.403 |
| SU_NGHIEP | 1.304 | 39.120 / 39.120 | 653.304 / 653.304 |
| SUC_KHOE | 884 | 26.520 / 26.520 | 442.884 / 442.884 |
| THAN_SAT | 511 | 15.330 / 15.330 | 256.011 / 256.011 |
| SAO | 352 | 10.560 / 10.560 | 176.352 / 176.352 |
| BAO_PHU | 42 | 1.260 / 1.260 | 21.042 / 21.042 |
| KHONG_RO_GIO | 519 + 288 ô bảng | 14.910 / 14.910 (dap_an_o không có vùng bảng) | – (lượt 501 xóa bảng như uno_run59) |
| **Tổng** | **26.556 + 288** | **798.300 / 798.300** | **13.082.613 / 13.082.613** |
| 6 sheet còn lại (GĐ5) | – | chưa cài (1.937 ô có đáp án mỗi lá) | – |

Khác:
- **T1b** – 7 biến thể × 30 lá × 807.600 ô (gồm vùng bảng KHONG_RO_GIO 807 ô/lá): **5.653.200 / 5.653.200**; riêng KHONG_RO_GIO 24.210/24.210 mỗi biến thể, kể cả **B1b – Không rõ giờ** (vòng 12 giờ thực sự đổi kết quả).
- **T8 chuẩn** – 55 lá × 2 chế độ E3: **2.961.200 / 2.961.200** (mọi sheet đã cài, gồm vùng bảng).
- **T2** – 34 khóa × 501 = **17.034 / 17.034** (thêm SUC_KHOE J6:J8, TO_CHAT M16/C16, HON_NHAN AI64/AG64/X64/R64/S64). 14 khóa còn lại thuộc KET_LUAN, LUAN_GIAI, DOC_NHANH (GĐ5).
- **T3** 710/710 · **T4** 510/510 · Âm lịch 74.144/74.144, data39 39/39.
- Ô GĐ5 còn phải mượn khi kiểm: KET_LUAN C40, C60, C163.
- Tốc độ (máy rảnh): **672 ms/lá** (Biết giờ), **800 ms/lá** (Không rõ giờ, 12 giờ) – trong ngưỡng YEU_CAU (≤ 1 giây / ≤ 5 giây).

## Phát hiện kỹ thuật GĐ4
| Mã | Nội dung |
|---|---|
| KT-07 | 12 Bảng dữ liệu đều là bảng một biến theo cột (`dtr=0`, `r1=C6`); lõi tính báo lỗi rõ nếu gặp bảng 2 chiều (không có trong 6.3). |
| KT-08 | Chỉ 98 ô của chính KHONG_RO_GIO phụ thuộc vùng bảng; 10 sheet GĐ4 khác không phụ thuộc – dùng được đáp án 501 lá kiểu uno_run59 (xóa bảng). |

## Rủi ro
| Rủi ro | Mức | Xử lý |
|---|---|---|
| GĐ5 thêm 1.937 ô → thời gian một lá có thể chạm 1 giây | Vừa | Đã đo điểm nóng: so sánh `_cmp` và COUNTIFS (~40% thời gian). Hướng: nhớ đệm vector vùng hằng, rút ngắn đường so sánh số/chữ – không đổi kết quả, kiểm lại toàn bộ |

## MÂU THUẪN mới
Không có.

## Còn lại
| GĐ | Nội dung | Trạng thái |
|---|---|---|
| 5 | KET_LUAN, LUAN_GIAI, tham khảo, giao diện hoàn chỉnh, SQLite, so sánh | **Xong** – T1 856.410/856.410; T1b 7×865.710; T8 2×1.587.135; T2 24.048/24.048 (48 khóa); T1-501 14.053.050/14.053.050; T3 710/710; T4 510/510; T6 12/12; tốc độ ~0,7 s/lá |
| 6 | Xuất PDF/XLSX/JSON + in; bộ cài Windows/macOS; kiểm thử tổng T1–T8; hướng dẫn sử dụng; README build | **Xong** – xem bảng GĐ6 |

### Quyết định GĐ5 (14/11/2026)
- MAU_THUAN, KIEM_THU, CSDL: chỉ tính ẩn trong lõi (vẫn kiểm thử T1), KHÔNG hiển thị trên giao diện.
- Trang tham khảo hiển thị: TRANG_CHU, HUONG_DAN, THAM_KHAO.

## GĐ6 (14/11/2026)
| Việc | Kết quả |
|---|---|
| Xuất **PDF** (Qt, phông Inter): thông tin lá, khối KẾT LUẬN, 2 biểu đồ, mục chọn (mặc định KET_LUAN B–K + bảng Ngũ hành thiếu), "Trang x/y", ngày in, phiên bản quy tắc | Xong – 28 trang mặc định, 0,6 s |
| **In** (QPrintDialog, cùng bộ dựng trang với PDF) | Xong |
| Xuất **XLSX** – 30 sheet đúng địa chỉ ô, gộp ô, độ rộng, cột/dòng ẩn, định dạng số, màu định dạng có điều kiện; sheet THONG_TIN | Xong – 0,8 s |
| Xuất **JSON** – dữ liệu nhập + {sheet: {ô: giá trị}} mọi ô công thức 33 sheet | Xong – 0,05 s |
| Cột/dòng ẩn của workbook (ô tính phụ) áp vào giao diện, PDF, XLSX (`bo_cuc.an_cot/an_dong`) | Xong |
| Bộ cài: Windows .exe (Inno Setup tiếng Việt, có gỡ cài đặt) 42,3 MB; macOS .dmg arm64 51,7 MB, x86_64 56,2 MB – dựng trên kho riêng tư lanhnguyenngoc-pixel/battu-thieu-vi-hoa (GitHub Actions) | **Xong** – cả 3 máy: cài ĐẠT, tự kiểm khi chặn mạng ĐẠT (dấu vân tay trùng bản gốc), gỡ ĐẠT |
| Tự kiểm bản đã cài `--tu-kiem`: dấu vân tay SHA-256 toàn bộ kết quả 3 lá + xuất 3 định dạng | Xong |
| Kiểm thử tổng T1–T8 + T7 (20 lá ngẫu nhiên mới, oracle_dump.py) | **100%** – BAO_CAO_KIEM_THU_TONG.md |
| Hướng dẫn sử dụng PDF có hình (13 trang), README build | Xong |
| Đề xuất bố cục giao diện hiện/ẩn (tai_lieu/DE_XUAT_GIAO_DIEN.md) | Chờ duyệt |

Quyết định 14/11/2026: bộ cài chỉ Windows và macOS; Linux phát triển sau.

### MÂU THUẪN nhỏ (ghi nhận, không đổi kết quả)
| Mã | Nội dung | Xử lý |
|---|---|---|
| MT-16 | YÊU CẦU 5.2 ghi vùng phụ trợ KET_LUAN "dòng 200-250"; workbook ẩn dòng **201–251** (dòng 200 là tiêu đề "PHỤ TRỢ (ẨN)", dòng 251 chỉ có ô tính phụ AG:AI) | Phần mềm ẩn 200–251 (gồm cả hai cách hiểu). Anh xác nhận nếu muốn khác |
| MT-17 | YÊU CẦU 5.2 mục 8 liệt kê MAU_THUAN, KIEM_THU, CSDL ở trang Tham khảo | Theo quyết định của anh ngày 14/11: chỉ tính ẩn, không hiện |

### Quyết định 14/11/2026 (09:55)
- Chủ dự án **thống nhất 14 màn hình giao diện chính** của bản 1.0.0 (Tóm tắt, Nhập liệu, Theo lĩnh vực, Luận giải, Tìm kiếm, Phân tích, Đại vận, Theo dõi, Ngũ hành thiếu, Chuyên đề, Tham khảo, Danh sách lá số, So sánh hai lá, Xuất/In) – giữ nguyên như ảnh `anh_gd6/`.

## Bản 1.1.0 (14/11/2026) – 5 phần giao diện đã duyệt
| Việc | Kết quả |
|---|---|
| **Chế độ chuyên gia** (công tắc ở thanh bên, ghi nhớ): thường = Tóm tắt, Lá số, Nhập liệu, Theo lĩnh vực, Luận giải, Vận trình (Đại vận), Ngũ hành thiếu, Kiểm chứng sự việc, Tham khảo, Danh sách, So sánh; chuyên gia thêm Trung gian, Phân tích, Chuyên đề, Lưu niên 100 năm, Theo dõi, địa chỉ ô Excel; Không rõ giờ hiện khi bật Không rõ giờ | Xong |
| **Lá số trực quan**: 4 trụ (KET_LUAN dòng 15–23), màu theo hành đọc từ chữ trong ô, nạp âm năm (LUC_THAN dòng 5), thập thần 5 nhóm (KET_LUAN C24), dải 10 đại vận (DAI_VAN dòng 35–44, 102; tô theo cờ dòng 106) | Xong |
| **Luận giải dạng bài đọc** (LUAN_GIAI phần A, bảng vận trình mục 7 có tô đỏ theo cờ L, phần B 6 ý dạng thẻ) + tìm trong bài | Xong |
| **Chú giải thuật ngữ** khi rê chuột – NGUYÊN VĂN workbook: LUAN_GIAI cột E "Thuật ngữ bằng lời thường" + ô ghi chú dạng "X = …" (KET_LUAN G17…); 18 thuật ngữ, có số trang sách; bảng chú giải cuối trang Luận giải; áp vào mọi bảng | Xong |
| **Nhập sự việc SU_KIEN** B6:D35 (30 dòng; 10 loại sự việc trích tự động từ Data validation của workbook – `danh_sach_chon`); lưu cùng lá số | Xong |
| Kiểm thử tổng chạy lại + **T9** (20 lá × 6 sự việc, oracle mới) | **100%** – BAO_CAO_KIEM_THU_TONG.md |
| Bộ cài 1.1.0 Windows / macOS arm64 / macOS x86_64: cài – tự kiểm chặn mạng – gỡ | **ĐẠT cả 3 máy** |
| Hướng dẫn sử dụng 18 trang (thêm Chế độ chuyên gia, Lá số, Luận giải, Kiểm chứng sự việc) | Xong |

Quyết định 14/11/2026: MT-16 không cần thực hiện (giữ như hiện tại: ẩn KET_LUAN 200–251). Đề xuất giao diện (DE_XUAT_GIAO_DIEN.md) – 5 phần đã duyệt và làm xong ở bản 1.1.0.

## Bản 1.1.1 (14/11/2026)
- Quyết định: màn hình đầu tiên là **Nhập liệu**; thứ tự thanh bên: Nhập liệu → Lá số → **Luận giải chung** (tên cũ Tóm tắt) → **Luận giải chi tiết** (tên cũ Luận giải) → Vận trình → Theo lĩnh vực → Ngũ hành thiếu → Kiểm chứng sự việc → (chuyên gia: Không rõ giờ, Trung gian, Phân tích, Chuyên đề) → Tham khảo → Danh sách lá số → So sánh hai lá.
- Chỉ đổi giao diện, không đổi lõi tính; T7 + T9 + tự kiểm chạy lại trên Windows/macOS khi dựng bộ cài.

## Bản 1.1.2 (14/11/2026)
- Thẻ đại vận (màn hình Lá số): **XU THẾ chính theo Thiệu Vĩ Hoa** (DAI_VAN dòng 54/69 theo phương pháp khuyến nghị – như DAI_VAN!C106) in đậm; mức **Hóa giải toàn thư** (dòng 102, HG tr.83 – workbook ghi "chỉ hiển thị") in nhỏ bên dưới; trạng thái nhật can và tàng can chuyển vào chú giải khi rê chuột.
- Bỏ các ghi chú kỹ thuật thừa (theo đánh dấu của chủ dự án): tiêu đề dài KET_LUAN B3 ở Luận giải chung, "Kiểm tra và ghi chú (nguyên văn workbook)", "(dòng 37-38)", thời gian tính (ms), các nhãn số ô công thức / tên sheet / địa chỉ ô (C14, E3…), dòng chú thích cờ DAI_VAN dòng 106, đường dẫn tệp lá số; PDF bỏ tên sheet ở tiêu đề mục. Địa chỉ ô nguồn của chú giải chỉ hiện ở Chế độ chuyên gia.
- Chỉ đổi giao diện; lõi tính không đổi.

## Bản 1.2.0 (14/11/2026) – phần mềm độc lập với workbook
Quyết định của chủ dự án: phần mềm và bộ công cụ Excel khi hoàn thiện là **độc lập**; mọi chữ người đọc thấy không được trỏ về sheet/ô của workbook.

| Việc | Cách làm |
|---|---|
| **Lớp làm sạch khi hiển thị** (`export/lam_sach.py`) áp cho giao diện, PDF, XLSX, chú giải | Bỏ: tên sheet + địa chỉ ô/dòng/cột ("(DOI_CHIEU C33)", "PP1_DINH_TINH!C91", "chi tiết LUU_NIEN cột BN", "(dòng 37-38)", "(G11)"), danh sách sheet trong ngoặc ("[TAI_LOC, MENH_CUC]"), mã giả định G…, mã quy tắc nội bộ (TC66, HN31, CM14, UH3b…); "xem SHEET …" đổi thành tên mục của phần mềm ("xem Sự nghiệp"); trỏ tới phần chỉ chạy ẩn (MAU_THUAN, KIEM_THU, CSDL) thì bỏ. **Giữ nguyên** nội dung luận đoán và trích dẫn sách (P1 tr…, P2 tr…, HG tr…, UH tr…), ký hiệu PP1/PP2/PA1/PA2 |
| Lõi tính | **Không đổi** – giá trị gốc vẫn trùng workbook 100% (T1–T9 dùng giá trị gốc); JSON xuất giá trị gốc (dữ liệu đối soát) |
| Mục Tham khảo | Bỏ Trang chủ và Hướng dẫn của workbook (chỉ hướng dẫn thao tác Excel); còn Tham khảo ngoài sách, Tiết khí |
| XLSX | Tên trang tính tiếng Việt (Luận giải chung, Đại vận…); bỏ trang Nhập liệu/Trang chủ/Hướng dẫn của Excel |
| Bảng ở chế độ thường | Ẩn số dòng/chữ cột kiểu Excel; chế độ chuyên gia vẫn hiện |
| Chú thích kỹ thuật GĐ1 còn sót ("…DAI_VAN C17–C21. Bảng đại vận đầy đủ có ở Giai đoạn 3") | Đã bỏ |
| PDF | Đầu trang "Phiên bản 1.2.0 · Quy tắc 6.3" (không ghi mã workbook) |
| Hướng dẫn sử dụng | Viết lại, không nhắc workbook/sheet/ô |
| Kiểm thử | T5 (xuất, chữ XLSX so sau làm sạch) 100%; T7 570.940/570.940; T9 577.140/577.140; tự kiểm ĐẠT |

Ghi chú: đây là thay đổi TRÌNH BÀY, không phải đổi quy tắc tính; workbook 6.3 giữ nguyên (SHA-256 7ae75965…bca4c5).

## Bản 1.3.0 (14/11/2026) – trình bày dạng văn bản đọc
- **Theo lĩnh vực B–K**, **Vận trình – Đại vận**, **Ngũ hành thiếu** (và Đọc nhanh): chế độ thường trình bày dạng văn bản đọc (`ui/bai_doc.py`): tiêu đề, "Hạng mục: nội dung", các dòng liền nhau gộp thành một bảng gọn không số dòng/chữ cột, hạng mục dài thành tiêu đề + đoạn văn; giữ màu nhấn và chú giải thuật ngữ. Lưới ô gốc chỉ ở chế độ chuyên gia.
- Theo lĩnh vực: ẩn mục "Luận giải chi tiết (bảng ô)" và "Đối chiếu sự việc" ở chế độ thường (đã có trang riêng); kết quả tìm kiếm không hiện địa chỉ ô.
- Thẻ đại vận (Lá số): bỏ dòng Hóa giải toàn thư (theo yêu cầu) – chỉ còn Xu thế theo Thiệu Vĩ Hoa.
- Bộ làm sạch: bỏ thêm ghi chú kỹ thuật trong ngoặc có "Excel", "sheet", "ô E3", "chốt <ngày>".
- Chỉ đổi trình bày; T5, T7, tự kiểm đạt.
