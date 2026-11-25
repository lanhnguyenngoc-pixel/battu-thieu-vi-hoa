# TIẾN ĐỘ – PHẦN MỀM BÁT TỰ THIỆU VĨ HOA (theo workbook 6.3)

Cập nhật: 14/11/2026 · Giai đoạn hiện tại: **GĐ6 – Xuất, đóng gói, kiểm thử tổng, hướng dẫn: XONG** · Phiên bản phần mềm **1.5.0** (quy tắc 6.4) · Nguồn: workbook SHA-256 `7ae75965…bca4c5`

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
| T1 (đề xuất hoàn thiện 15/11/2026) | Lớp hiển thị/in, KHÔNG đổi giá trị ô: (C1) bản in giữ nguyên văn câu sách tại chỗ, màn hình hiện ý chính thay cho "(xem căn cứ sách)" – `ui/trich_dan.py`, `export/xuat_pdf.py`; (C2) chữ bị workbook rút gọn `LEFT(x,N)&"…"` (LUAN_GIAI 54 ô, THEO_DOI 100 ô) hiện đầy đủ – `export/khoi_phuc_cat.py`; (C3) tuổi đại vận ghi kèm năm dương lịch – `export/ghi_tuoi.py`. Kiểm thử mới T11 (`tests/t11_hien_thi.py`) | Xong – T5, T7, T9, T10, T11, kho lá số ĐẠT |
| T2 (đề xuất hoàn thiện) | Lớp tổng hợp `battu_core/tong_hop/`: (B1–B2) `quan_he.py` – ngũ hợp, can xung/khắc, lục hợp, tam hợp, bán hợp, củng hợp, tam hội, lục xung, hình (vô ân, trì thế, vô lễ, tự hình), hại, phá, ám hợp, kho, thiên khắc địa xung, phục ngâm, trùng chi – cho tứ trụ, đại vận, lưu niên, kèm chạm dụng thần/cung; `du_kien.py` – hồ sơ dữ kiện JSON (tứ trụ, bản mệnh, đại vận, 11 năm lưu niên, kết luận các mục). Bản in không cần mã nguồn sheet: `tools/sinh_khong_cat.py` → `export/_khong_cat_sinh.py` (gọi trong `dung_lai.sh`). Kiểm thử T12: khớp hợp/xung workbook trên 521 lá | Xong – T5, T7, T9, T10, T11, T12, kho lá số ĐẠT |
| T3 (đề xuất hoàn thiện) | Lớp phán đoán `battu_core/tong_hop/phan_doan.py` + quy tắc `battu_core/data/quy_tac_tong_hop.json` (chủ dự án duyệt, sửa không cần sửa mã): (A1) đại vận một thang 5 mức, tách 5 năm đầu/5 năm sau; (A2–A3) một điểm năm, trừ nặng xung chi mang dụng thần, cung phu thê, hạn nặng (trần "Tốt"); (A4) thọ – yểu gộp nhiều tiêu chí; (A5) tài lộc một thang; (A6) lục thân cân nguồn theo trọng số đã duyệt; (B3) cờ sát ngưỡng vượng suy + gợi ý kiểm chứng; (B4) lọc mục theo tuổi, trường mới `LaSo.tinh_trang_hon_nhan`, `LaSo.so_con` (không phải ô workbook). Kiểm thử T13 | Xong – T5, T7, T9–T13, kho lá số ĐẠT |
| T4 (đề xuất hoàn thiện) | Bản luận giải tổng hợp không cần mạng, khoảng 5 trang A4 (chủ dự án duyệt): `battu_core/tong_hop/viet_ban.py` – tóm tắt nhanh + 10 mục (lá số, bản mệnh, tính cách, học vấn – sự nghiệp, tài lộc, hôn nhân – con cái, cha mẹ – anh em, sức khỏe, đại vận – lưu niên, hóa giải – kết luận); `export/xuat_luan_giai.py` – HTML, PDF, Word .docx (thư viện chuẩn); giao diện: trang "✦ Luận giải tổng hợp" (`ui/trang_tong_hop.py`), ô nhập Hôn nhân / Số con. Sửa hồ sơ: cột LUU_NIEN "Đại vận" so trùng khớp (không lẫn "Đại vận số"). Kiểm thử T14 | Xong – T5, T7, T9–T14, kho lá số ĐẠT |
| T5 (đề xuất hoàn thiện) | "✨ Nâng cao bằng AI" – TÙY CHỌN, mặc định tắt: `battu_core/tong_hop/ai_claude.py` (Claude Messages API qua urllib; hồ sơ gọn KHÔNG họ tên, ngày giờ sinh; lời nhắc buộc giữ dữ kiện; bộ kiểm tra tứ trụ, dụng – hỷ – kỵ, năm đại vận, can chi và mức từng năm – sai thì giữ bản không mạng; lưu đệm `~/.battu_thieu_vi_hoa/ai/`; chứng chỉ gốc macOS cho bộ cài); `ui/khoa_api.py` (khóa của người dùng trong Credential Manager / Keychain, dự phòng tệp 0600; ANTHROPIC_API_KEY ưu tiên); trang Luận giải tổng hợp: nút AI, 🔑 khóa + mô hình (Sonnet 5.5 mặc định / Opus 5.5 / Haiku 5.5), hỏi trước khi gửi kèm ước lượng token, chạy nền, chuyển "Bản không mạng / Bản AI", xuất bản đang xem. Kiểm thử T15 (giả lập API) | Xong – T5, T7, T9–T15, kho lá số ĐẠT; chưa thử với khóa thật |

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

## Bản 1.4.0 (14/11/2026) – ô nhập còn thiếu so với workbook, sao lưu kho
| Việc | Kết quả |
|---|---|
| Ngày sinh vợ/chồng (NHAP_LIEU C34:G34) | Ô nhập ở màn hình Nhập liệu; chỉ nhập năm → C34 = năm (như hướng dẫn của workbook) |
| Việc lớn để tìm giờ sinh (KHONG_RO_GIO C11:E22) | Thẻ "Việc lớn đã xảy ra" trong mục Không rõ giờ (12 dòng, 8 loại việc lấy từ danh sách chọn của workbook) |
| Đánh giá thực tế (THEO_DOI N/O/P) | Trang mới "⚑ Đối chiếu thực tế"; nhập theo năm, phần mềm đặt đúng dòng năm do lõi tính |
| Sao lưu / khôi phục kho | Danh sách lá số → Sao lưu kho… / Khôi phục… (khôi phục chỉ thêm, bỏ qua lá trùng) |
| Kiểm thử | Toàn bộ T1–T10 đạt 100%; T10 mới 577.140 / 577.140 |
| Không đổi quy tắc | Không sửa workbook; mọi ô nhập mới là ô nhập sẵn có của workbook 6.3 |
| Thống kê đối chiếu trên cả kho (đề xuất 3.8 – duyệt 14/11/2026) | Danh sách lá số → ⚑ Thống kê đối chiếu: cộng ô Theo dõi (N, Q, R, P) của mọi lá; xuất Excel. Kiểm: tổng trùng cộng E4, E5×E4, E6×E4 từng lá |
| Phân tích 3.1 (duyệt 14/11/2026) | tai_lieu/PHAN_TICH_3_1_DUNG_THAN.md: dụng thần khớp sách 40/73 lá; 3 thay đổi thử không cải thiện rõ → đề xuất giữ quy tắc |
| Quyết định 14/11/2026 | 3.3 giữ giờ đồng hồ nơi sinh; 3.4 giữ đổi ngày lúc 23 giờ (P1 tr.45) – không thêm tùy chọn |
| Chờ duyệt | 3.2: 94 giả định (tai_lieu/GIA_DINH_6.3_CHI_TIET.md, RA_SOAT_GIA_DINH_6.3.xlsx) |

## Bản 1.5.0 – quy tắc 6.4 (14/11/2026)
| Việc | Kết quả |
|---|---|
| Quyết định | Phương án A (duyệt 14/11/2026): giữ dụng thần theo bảng P1 tr.108, P2 tr.380-384; thêm ghi chú "dụng thần theo cách bổ hành khuyết" |
| Workbook 6.4 | G142 – tools/step17.py (vá XML theo cách 06_MA_THAM_CHIEU/build_excel, LibreOffice tính lại). Chỉ thêm DUNG_HY_KY dòng 197–199 và nối vào KET_LUAN C52; MAU_THUAN dòng 168 |
| Phần mềm | Dịch lại từ 6.4; ghi chú hiện ở mục B – Bản mệnh; 13/85 lá ví dụ có ghi chú |
| Kiểm thử | Lấy lại toàn bộ đáp án trên 6.4; T1–T10 đạt 100% |

## Bản 1.6.0 – quy tắc 6.5 (15/11/2026) – giai đoạn D
| Việc | Kết quả |
|---|---|
| Workbook 6.5 (tools/step18.py, SHA-256 f19f3f74…c56f4) | VS-2, VS-3, B1, B2, CĐB-1/3/4, DT-2a, DT-3, DV-3, TL-1; giới tính lá mẫu; vá mẫu tài lộc; MAU_THUAN dòng 6, 169–180 |
| Quy tắc trình bày (duyệt 15/11/2026) | Kết luận, luận giải của công cụ: chữ đậm; trích dẫn sách: chữ thường, dòng riêng, ẩn sau "📖 Căn cứ sách" (PDF: in ngay dưới) |
| Giao diện | Kết luận đại vận ở đầu trang Vận trình (mức, vì sao 5 năm đầu/sau, tài lộc, công danh); thẻ đại vận ở Lá số hiện mức kết luận |
| Đính chính | P1-VD14 không phải sửa (sách P1 tr.196 là giờ Đinh Dậu – dữ liệu nghiên cứu nhập nhầm, đã sửa); DT-2a vốn đã có trong 6.4, 6.5 đổi cách đếm "có mặt" |
| Kiểm thử | Đáp án LibreOffice lấy lại toàn bộ trên 6.5; T1–T10 đạt 100% (tai_lieu/BAO_CAO_KIEM_THU_1.6.0_CHI_TIET.md); tự kiểm ĐẠT, dấu vân tay mới ae7e52ec… |
| Chi tiết | tai_lieu/giai_doan_A/GIAI_DOAN_D_TRIEN_KHAI.md |

## Bản 1.6.1 (15/11/2026) – tốc độ giao diện
| Việc | Kết quả |
|---|---|
| Nhập liệu | Sửa ô không tính ngay; bấm **▶ Luận giải** (hoặc Enter, F5). Báo "Thông tin đã thay đổi" khi kết quả đang hiện là của thông tin cũ |
| Tính ở nền | Giao diện không đứng trong lúc tính (khựng lâu nhất 0,12 giây khi đo) |
| Vẽ trang | Chỉ vẽ trang đang mở; sửa lỗi vẽ trang Vận trình khi đang ẩn |
| Nhớ kết quả | 6 lá tính gần nhất được nhớ – tính lại cùng lá 0,1 giây |
| Mở phần mềm | Lá mặc định tính ở nền |
| Kiểm thử | Lõi tính không đổi; T5, T6, T7, T9, T10, tự kiểm đạt |

## Bản 1.7.0 (15/11/2026) – lớp tổng hợp, luận giải tổng hợp ~5 trang, nâng cao bằng AI
Theo bản "Đề xuất hoàn thiện phần mềm Bát Tự Thiệu Vĩ Hoa" (chủ dự án duyệt 15/11/2026; phương án trung gian: làm lớp tổng hợp trước,
khi quy tắc gỡ mâu thuẫn chạy ổn trên nhiều lá mới chuyển vào workbook thành quy tắc 6.6).
| Việc | Kết quả |
|---|---|
| Lõi tính, workbook | KHÔNG đổi (quy tắc 6.5; dấu vân tay tự kiểm giữ nguyên ae7e52ec…, a33f9009…, eba9c43a…) |
| T1 hiển thị, in | Câu sách in đủ; chữ bị cắt "…" hiện đủ; tuổi đại vận kèm năm dương lịch |
| T2 dữ kiện | Quan hệ can chi đủ loại (khớp hợp/xung workbook trên 521 lá); hồ sơ dữ kiện |
| T3 phán đoán | Một kết luận cho mỗi đại vận/năm/lĩnh vực; quy tắc ở `battu_core/data/quy_tac_tong_hop.json` (T3-1) |
| T4 bản luận giải | Tóm tắt + 10 mục, khoảng 5 trang; PDF, Word; trang "✦ Luận giải tổng hợp"; ô Hôn nhân, Số con |
| T5 AI (tùy chọn) | Claude API bằng khóa của người dùng; không gửi họ tên; bộ kiểm tra dữ kiện; lưu đệm |
| Tự kiểm bộ cài | Thêm bước dựng bản luận giải tổng hợp (PDF ≤ 5 trang, Word) khi chặn mạng |
| Kiểm thử | T5, T7, T9, T10, T11–T15, kho lá số, tự kiểm: ĐẠT. T1–T4, T8 cần hồ sơ HO_SO_PM (không đổi lõi tính nên không ảnh hưởng) |
| Hướng dẫn | tai_lieu/HUONG_DAN_SU_DUNG.pdf thêm mục 6c (Luận giải tổng hợp, AI), câu hỏi thường gặp |

## 1.7.1 – đối chiếu với ví dụ sách (15/11/2026)
- Chạy lớp tổng hợp trên chính các ví dụ của Thiệu Vĩ Hoa phát hiện hồi quy: điểm đại vận riêng của lớp tổng hợp chỉ khớp sách 70/171 vận (41%), kém workbook 6.5 (123/171 = 72%).
- Sửa: hướng tốt/xấu của mỗi đại vận và mỗi nửa vận lấy theo KẾT LUẬN ĐẠI VẬN của workbook (mục F, P1 tr.124–125); điểm riêng chỉ dùng để nâng lên "Rất tốt/Rất xấu" khi cùng hướng và để ghi lý do. Kết quả: 123/171 (72%).
- Vượng – nhược trên 65 lá ví dụ P1–P3: lớp tổng hợp 57 (88%), workbook 56 (86%); 50 lá rút ngẫu nhiên: 44 (88%) so với 43 (86%).
- Thêm kiểm thử T16 (tests/t16_sach.py) vào quy trình đóng gói: lớp tổng hợp không được khớp sách kém hơn workbook.

## 1.8.0 – Quy tắc 6.6 (15/11/2026)
- Workbook 6.6 = 6.5 + một quy tắc vượng suy (DOI_CHIEU C30 = 36%, D30 = 52%, C33, C36): khi PP1 và PP2 khác nhau mà phe mình PP2 xa ngưỡng 40% (dưới 36% hoặc từ 52%) thì theo PP2. Bỏ điều kiện cũ (chỉ khi PP1 khẳng định vượng và PP2 < 30%) vì đã nằm trong quy tắc mới.
- Kiểm chứng trên 431 lá có lời giải vượng suy: 65 lá Thiệu Vĩ Hoa 57 → 59; 366 lá sách khác 242 → 289 (66% → 79%); kiểm chéo chia đôi 20 lần 68,7% → 77,2%. Đại vận khớp sách 123 → 126/171, ngược chiều 38 → 35. Dụng thần không đổi (21/38 – việc của đợt 2).
- Dựng lại bằng tools/dung_lai.sh; LibreOffice (oracle_mo_rong) trên 9 lá, gồm 6 lá có quy tắc tác động: 261 567 / 261 567 ô khớp. Sinh lại đáp án T7/T9/T10 từ 6.6 (chỉ 3 lá T7_01/02/06 đổi – đúng các lá quy tắc tác động); dấu vân tay tự kiểm sinh lại (lá Tự kiểm 3 đổi theo quy tắc).

## 1.8.1 – DT-3, DT-1, trang Ngũ hành thiếu (15/11/2026)
- DT-3 (workbook 6.6): KET_LUAN C36:C38 (Dụng/Hỷ/Kỵ thần) và 30 ô THAN_SAT theo phương pháp khuyến nghị (DOI_CHIEU C33; DUNG_HY_KY C155, C157:C159) – trước đây luôn lấy PP1 nên có thể ghi khác hành mà đại vận, lưu niên đang dùng.
- DT-1 (workbook 6.6): DUNG_HY_KY mục O (dòng 221–227) – ứng viên dụng thần thay thế (phù ức khi ưu tiên điều hầu; nếu luận chính cách; điều hầu theo mùa; dụng thần đã mạnh ngang phe mình → ứng viên sau theo P1 tr.99) và mức chắc chắn; KET_LUAN C39 ghi thêm dòng ứng viên; lớp tổng hợp ghi vào mục 2 bản luận giải. Không đổi dụng thần chính.
- Đo trên sách: 38 lá P2 tr.397–431 – dụng chính khớp 20 (53%), chính hoặc ứng viên 24 (63%); 247 lá TTT + HTL – 106 (43%) → 142 (57%).
- Trang Ngũ hành thiếu: 3 thẻ Kết quả lá số / Lý thuyết – cách tính / Bảng gốc (ui/trang_ngu_hanh_thieu.py).
- Kiểm: LibreOffice 14 lá 407 022 / 407 022 ô; đáp án T7/T9/T10 sinh lại (thêm 10 ô mục O mỗi lá); T5, T7, T9–T16, kho lá số, tự kiểm: ĐẠT.

## 1.8.2 – ba quy tắc lớp tổng hợp vào workbook 6.6 + gợi ý từ sự việc đã qua (15/11/2026)
- Thọ – yểu: MENH_CUC dòng 96–105 (6 dấu hiệu thọ: thân vượng, nguyệt đức, thiên đức, sức khỏe chung thuận, bệnh có cứu, bình sinh ít bệnh); bảng P1 tr.173 nghiêng yểu mà có ≥ 2 dấu hiệu → "CĂN CỨ NGƯỢC CHIỀU" ở KET_LUAN C143/C135. Lớp tổng hợp đọc thẳng kết quả này.
- Lưu niên: LUU_NIEN AA30:AA129, AL30:AL129 – năm có dấu hiệu hạn nặng (BM > 0) không xếp "Rất tốt" (hạ "Tốt vừa").
- Lục thân: KET_LUAN C114 – khi Uyên Hải / Hóa giải toàn thư có câu xấu hơn về con cái mà kết luận Thiệu Vĩ Hoa không kém, ghi đối chiếu nguồn theo thứ tự tin cậy đã duyệt.
- Sự việc đã qua: SU_KIEN cột phụ T:X, dòng 37–39 – so mỗi sự việc tốt/xấu với xếp hạng năm (P1 tr.133) theo PP1 và PP2, gợi ý giữ/đổi phương pháp (≥ 2 sự việc; đổi khi hơn ≥ 2); DOI_CHIEU C49 và trang Kiểm chứng sự việc hiện gợi ý; bản luận giải tổng hợp so thêm dụng thần chính với ứng viên thay thế (phan_doan.dung_than_su_kien – tham khảo).
- TRANG_CHU B3 ghi nội dung bản 6.6.
- Kiểm: LibreOffice 14 lá 409 262 / 409 262 ô; T7 578 460, T9/T10 584 660 ô (đáp án sinh lại); T5, T11–T16, kho, tự kiểm: ĐẠT.

## 1.9.0 – Quy tắc 6.7: bỏ quy tắc vượng suy 6.6, thọ – yểu xét lại (16/11/2026)
- Vượng suy: bỏ quy tắc 6.6 (theo PP2 khi phe mình < 36% hoặc ≥ 52%), trở lại cách 6.5 (PP1 đủ 5 bước, bước 5 có thể đổi kết luận). Đo trên 73 lá Thiệu Vĩ Hoa: 6.5 đúng 64, 6.6 đúng 63; bỏ 5 lá sách tự mâu thuẫn: 6.5 đúng 64/68, 6.6 đúng 61/68; 9 lá bị 6.6 chuyển chỉ đúng 4. DOI_CHIEU C30 = 0,3, D30 = 0,6, C33, C36 về công thức 6.5.
- Thọ – yểu (MENH_CUC dòng 96–128, thay khối 6.6): bước 1 theo Thiệu Vĩ Hoa – 6 dấu hiệu mệnh thọ, 9 dấu hiệu mệnh yểu (P1 tr.172 và bảng thực thương – ấn P1 tr.173), bên nhiều dấu hiệu hơn là kết luận; bước 2 luận thêm không dùng ngưỡng: nguyệt đức, thiên đức (HG tr.201), sức khỏe chung thuận, bệnh có cứu (P1 tr.97, 114–117), HG HS1, HS2, HS4, HS5, HS6 (tr.201–202) → nhẹ bớt / vững hơn; sức khỏe chung bất lợi, bệnh không có cứu → nặng hơn. KET_LUAN C135, C143, LUAN_GIAI F119, KHONG_RO_GIO W27 dùng chung kết luận cuối C126.
  - Trên 73 lá Thiệu Vĩ Hoa: bước 1 THỌ 32, CÂN BẰNG 20, YỂU 21; không cặp dấu hiệu nào trái nhau cùng bật. Sách không có đủ ví dụ ghi tuổi thọ của chính mệnh chủ để đo độ đúng.
- Lớp tổng hợp: mức đại vận, lưu niên, thọ – yểu lấy đúng kết luận của workbook theo phương pháp đang dùng (không còn điểm riêng, không còn quy tắc "xung 2 chi dụng thần" của lớp tổng hợp – quy tắc này chưa có trong workbook). Giới tính đọc NHAP_LIEU C22. Thông báo AI ghi đúng dữ liệu gửi đi.
- Đổi mã trùng: DT-3 (6.6) → DT-4, DT-1 (6.6) → DT-5. MAU_THUAN dòng 181–183.
- Trình bày trang đọc (yêu cầu 16/11/2026): ngắt đoạn ở lớp hiển thị (export/ngat_doan.py) – mỗi gạch đầu dòng "•", mục (a)…(f), nhãn "DỮ KIỆN / QUY TẮC / KẾT LUẬN", câu mở bằng nhãn ("Dấu hiệu mệnh yểu (3): …") xuống dòng riêng; đoạn dài gom thành đoạn ≤ ~230 ký tự; không ngắt trong ngoặc, câu trích, chữ đậm. Áp cho Luận giải chung, Theo lĩnh vực, Luận giải chi tiết, Luận giải tổng hợp, PDF, Word; giãn dòng 135–145% trên màn hình, bản in tổng hợp vẫn ≤ 5 trang. Không đổi giá trị ô.
- Kiểm: xem BAO_CAO_KIEM_THU_TONG.md mục 1.9.0 – mọi bộ ĐẠT.

## 1.9.1 – Quy tắc 6.8 giai đoạn 1 (22/11/2026)
- Đã làm (theo KE_HOACH_6.8.md): phương án D (ấn ở mọi can chi cả bốn trụ, P1 tr.70); NHT mục I (bỏ nhánh "giả" ở "được lệnh", bỏ nhánh Nhật chủ Thổ ở "cùng hành"); điều hầu Mộc mùa hạ cần Thủy (P1 tr.76, 100); sao con cái mệnh nữ (thất sát, thực thần ↔ chính quan, thương quan theo giới); đại vận: can chi cùng tác dụng cả 10 năm, mức 6 bậc (Rất tốt/Tốt/Tốt vừa/Xấu vừa/Xấu/Rất xấu) + câu giải thích, chi vận xét hợp hóa và xung; dòng con cái theo lá số vợ/chồng khi nhập đủ ngày giờ; lớp tổng hợp lấy đúng mức và giải thích của workbook; quy trình phát hành (tag v1.9.1).
- Đo: 194 vận, 139 → 136 vận tốt/xấu cùng chiều do phương án D (3 vận P1-VD3 Nhược → Vượng), không do logic mức mới; Mộc mùa hạ: 73 lá không đổi, 501 lá đổi 13/18/23 lá; NHT 34 lá: PA1 điểm 34 → 31, PA2 kết luận 26 → 29; ví dụ trong tài liệu Word khớp phần mềm.
- Kiểm: toàn bộ bộ kiểm thử ĐẠT (xem BAO_CAO_KIEM_THU_TONG.md); B1 lệch do lỗi LibreOffice MT-13, thay bằng B1b.
- Chưa làm (giai đoạn 2, cần thiết kế và đo riêng): xu thế năm can+chi (2.1/2.11), người ứng 6 bước (2.2), điều 6–8 (2.3), bốn bổ sung 2.7, bảng dấu hiệu con cái (2.8), ô nhập riêng và can chi vợ/chồng (2.9 phần mở rộng).
- Điểm cần quyết: đại vận can nhàn (điểm can = 0) giữ nhánh cũ "bình thường" vì kế hoạch 6 bậc chưa có mức này; hợp hóa chỉ xét lục hợp; ô vợ/chồng dùng chung phoi_ngau (cần đủ giờ); LUU_NIEN vẫn chia 5/5 khi ghép đại vận.

## 1.9.5 – Quy tắc 6.12: dụng – hỷ – kỵ thần theo P1 tr.73–124 (25/11/2026)
- Sổ tay 6.12 (tools/step28.py + step28_dt_moi.py): (1) Thổ sinh tháng Hợi/Tý/Sửu/Dần, thân nhược → cần Hỏa điều hầu (P1 tr.78), Hỏa không lộ ở can (tàng can không tính, theo Cùng Thông Bảo Giám; người dùng chọn 25/11) thì lấy làm dụng thần (DUNG_HY_KY C215, C217, D217, C219); nhóm điều hầu cũ giữ ngưỡng 20%; thân vượng thì không đổi. (2) Xếp bậc vận khi dụng thần thứ nhất khuyết (DAI_VAN dòng 114, 128; P1 tr.102–103): hành thứ nhất ±2, dụng thần được chọn ±1. (3) Kỵ thần thật (DAI_VAN dòng 129, 136; tr.98): vận hành kỵ nặng thêm một mức. Không thêm cặp thông quan Thực thương – Quan sát vì P1 tr.110 nêu cặp này không thành mâu thuẫn.
- Phần mềm: lớp hiển thị `tong_hop/dung_hy_ky_hien.py` – "tham khảo thêm" theo can ngày (tr.74–75), theo mùa sinh (tr.75–79), hỷ/kỵ của dụng thần → việc (tr.96–98), nằm trong mục 2 bản luận giải; bản luận giải ≤ 6 trang. Kiểm thử mới T18.
- Chủ dự án duyệt: tháng Dần tính cho Thổ đông; tr.74 hiển thị "thổ chế thủy" (sách in "mộc", ghi chú); Thổ đông dùng "Hỏa không lộ ở can" thay cho ngưỡng 20% (sách không nêu số).
- Mặc định đang áp dụng (chưa có chỉ thị riêng): ngưỡng "thái quá" chỉ hiển thị, khuyết điều hầu chỉ cảnh báo.
- Hỷ thần cứu ứng hai hướng và "gánh vác" khi dụng khuyết (tr.97–98) đã đưa vào bản luận giải. Chưa làm: bốn cặp cứu ứng chi tiết giữa các thập thần (phần còn lại của H14), vì chưa đối chiếu xong trang sách.
- Kiểm cuối (25/11/2026): đáp án LibreOffice sinh lại một lần cho 6.12; tests.chay_tat_ca đạt 100% (T1 949.320/949.320, T2 501 lá 15.604.647/15.604.647, T7–T10, T8C đều khớp; B1_KHONG_RO_GIO lệch 5.689 ô là lỗi cũ MT-13, B1b khớp 100%); T11–T18 ĐẠT. Đại vận 73 lá Thiệu Vĩ Hoa: khớp sách 76 → 77/171; vượng – nhược 56/65 không đổi.

## 1.9.4 – Bốn kiểu lá số dễ tai ương + Bảng sức khỏe 10 năm (24/11/2026)
- Chỉ ở phần mềm (không đổi ô nào của workbook, quy tắc vẫn 6.11): `battu_core/tong_hop/tai_uong.py`; trang "✚ Sức khỏe 10 năm" (xem, PDF, Word); tóm lược trong mục Sức khỏe của Luận giải tổng hợp (vẫn ≤ 5 trang).
- Bốn kiểu: Kiêu thần đoạt thực; Dương Nhẫn quá nặng (từ 2 chữ trở lên, tính mệnh + vận hiện tại); ngũ hành quá lệch (một hành ≥ 55% kèm hành gần như không có, hoặc thân quá yếu); đại vận vào Tử – Mộ – Tuyệt kèm thiên khắc địa xung trụ Ngày. Nhãn bắt buộc: "cách luận tham khảo thêm thuộc dễ gặp tai ương nhất." Không đổi mức tài lộc, sự nghiệp, hôn nhân.
- Bảng 10 năm liên tiếp từ năm đặt xem: mức năm lấy từ lưu niên, tạng phủ, dấu hiệu hạn nặng, cờ từng kiểu; cờ ở bảng đại vận.
- Đo: ngưỡng chọn trên 73 lá sách + 300 lá ngẫu nhiên (xem Project: KET_QUA_4_KIEU). Kiểm thử mới T17 (60 lá: nhãn, đủ 10 năm liên tiếp, cờ có lý do) ĐẠT; T14, T5, T11, T12, T15, T16 đạt; tự kiểm bộ cài đạt.

## 1.9.3 – Quy tắc 6.11: bảng Ngũ hành thiếu, mục I, PA2 mới (24/11/2026)
- Chế độ PA2 – Theo sách: được lệnh (48%) và đắc địa (20%) có xét giả; đắc địa xét can năm, can ngày, can giờ tại chi tháng (Trường sinh, Lộc, Kình dương, Mộ khố); cùng hành (20%) và được sinh (12%) theo hành gốc, không hợp hóa. PA1, mục II, PP1, dụng thần, đại vận không đổi.
- Đo 48 lá trang: khớp điểm 18 → 19, khớp nhóm 41 → 32. Kiểm thử tổng đạt (B1_KHONG_RO_GIO lệch là lỗi cũ MT-13, đã có B1b). Sổ tay 6.11; đáp án sinh lại.
- Đo các quy tắc rút từ ví dụ sách (R1, R2, R3a, R3b-1/2, R10): không quy tắc nào được bật (làm giảm chỉ số bộ bảo vệ hoặc không có lợi); chi tiết trong Project: KET_QUA_DO_R1_R10_24_11.md.

## 1.9.2 – Quy tắc 6.9: giai đoạn 2 của kế hoạch 6.8 (22/11/2026)
- Đã làm: (1) người ứng của hạn năm (6 bước) và xu thế năm can+chi; điều 6, 7, 8 (P2 tr.288–289); (2) sự nghiệp: cách kình dương; thọ – yểu: dấu hiệu "thân vượng thiếu cả quan sát và thực thương"; (3) con cái (P2 tr.500–514; P1 tr.157): câu tổng hợp (nhiều/vừa/ít con, nuôi dạy, con đầu lòng), danh sách các câu sách ứng với lá số, năm có dấu hiệu bất lợi cho con; hiển thị ở Theo lĩnh vực, Luận giải tổng hợp, PDF, Word; (4) lưu niên: can năm ngũ hợp hóa (P2 tr.282 điều 5) xét theo hành hóa thần; (5) nhập tứ trụ vợ/chồng theo can chi, dòng tham khảo con cái của vợ/chồng thêm vào Theo lĩnh vực (Gia đạo).
- Đo: "tham hợp quên xung" (PP1 dòng 24) đổi sang tính cả can hợp: 73 lá Thiệu Vĩ Hoa 63 → 62 khớp (P1-VD13 đúng → sai), 501 lá đổi 2 lá → KHÔNG áp dụng, chỉ ghi nhận. Câu tổng hợp con cái trên 574 lá: số con nhiều 209, vừa 274, ít 91; nuôi dạy: thuận 181, có cả hai mặt 194, có trở ngại 198 (bản đầu gộp theo "có/không" thì 98% lá ra "có trở ngại" nên đã đổi sang so số câu thuận với số câu trở ngại, chênh ≥ 2).
- 5 ví dụ con cái của sách (P2 tr.505–506), công cụ so với sách: số con khớp 3/5 (VD1 nhiều, VD3 vừa, VD4 nhiều), VD2 sách 1 trai 1 gái – công cụ "ít" (gần), VD5 (nữ) sách 3 con trai – công cụ "vừa" (lệch); việc nuôi dạy: sách đều nói con thành đạt, công cụ chỉ 1/5 ra "bình thường", 4/5 "có trở ngại (nhưng có dấu hiệu thuận)" → lệch, chưa chỉnh vì 5 ví dụ chưa đủ để đặt ngưỡng.
- Chưa làm / không làm được: 2.11 mục A (năm sự việc của sách – không có dữ liệu có nhãn để đo); câu "quan sát hỗn tạp cứu bằng hợp" (cần đối chiếu lại nguyên văn P1 tr.115–117); chưa đối chiếu đủ 15 ví dụ con cái (mới 5 ví dụ có kết quả rõ ở tr.505–506); danh sách nhóm C/D các câu sách chỉ liệt kê, không đổi kết luận.
- Lưu ý kỹ thuật: bộ ảnh trong hướng dẫn vẫn là ảnh cũ (tools chụp màn hình hỏng); dòng hôn phối chỉ lấy câu đã tính của lá thứ hai, workbook vẫn một lá số.

- Bổ sung cuối (22/11/2026): chi lưu niên hợp/hội cục với tứ trụ và đại vận (P2 tr.281–283, 254–255; điều 5, 10, 12, 17, hóa cục) → cột LUU_NIEN!DY, chỉ là câu sự việc tham khảo, KHÔNG đổi tốt/xấu của năm; chưa làm bán hợp cho lưu niên và chưa dùng hóa thần của chi để đổi vai trò năm (sách không nói rõ chi phải kề sát hay không).
- Nuôi dạy con: đổi sang đếm số câu thuận/trở ngại của sách, chỉ tham khảo (khớp sách 5/10 nên không dùng làm kết luận); D23 ghi "nghiêng về; quy tắc sao cha/mẹ chỉ ngược lại". 16 ví dụ sách đã đối chiếu (/tmp/vd16.py).
- Sửa lỗi: nhánh cache của giao diện làm mất phoi_tu_tru (đã sửa, có test T13); test PDF đòi nhãn 6.7 (đã sửa). Mục 2.11 A không đo được (không có dữ liệu năm–sự việc có nhãn). Đại vận "can nhàn" giữ "bình thường" (sách không có quy tắc riêng).
- Cập nhật 23/11/2026: (1) sửa lỗi Excel báo "Repaired" – tách 101 công thức > 8192 ký tự, có test tests/t_gioi_han_excel.py; (2) NHT mục I PA2: cột trạng thái đắc địa hiện theo vòng Thủy (khớp cách tính, G140), ghi chú F8 nêu khác sách P1 tr.34 – chỉ ở sheet ngũ hành thiếu; (3) oracle và toàn bộ kiểm thử chạy lại, ĐẠT. Còn lại: người dùng mở workbook bản mới bằng Excel thật và báo cảnh báo (nếu có).

## 1.9.6 – Quy tắc 6.13: PA1 mục I không xét nhật chủ Mậu, Kỷ (25/11/2026)
- Sổ tay 6.13 (tools/step29.py + step29_pa1_tho.py): chỉ ở chế độ PA1 – Theo trang, dòng "cùng hành" (mục I) không tính khi nhật chủ là Mậu/Kỷ (Thổ); PA2 – Theo sách giữ nguyên. Mục II, III, dụng thần, đại vận không đổi.
- Kiểm (25/11/2026): chỉ sinh lại đáp án B4_PA1 và T8C_PA1 (các bộ khác chạy PA2 mặc định, không đổi); tests.chay_tat_ca đạt 100% (B4_PA1 958.620/958.620, T8C_PA1 1.757.470/1.757.470; B1_KHONG_RO_GIO lệch đã biết MT-13).
