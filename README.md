# Phần mềm Bát Tự Thiệu Vĩ Hoa – mã nguồn (phiên bản 1.3.0, quy tắc 6.3)

Chuyển workbook **BAT_TU_THIEU_VI_HOA_6.3.xlsx** (SHA-256 `7ae75965…bca4c5`) thành phần mềm chạy không cần mạng.
Kết quả phải trùng workbook 100% ở mọi ô công thức (đáp án do LibreOffice tính – MT-01).

## Chạy thử
- **Windows:** cài Python 3.12 (python.org, chọn "Add to PATH"), bấm đúp `chay_thu.bat`. Lần đầu tự cài PySide6.
- **macOS / Linux:** `./chay_thu.sh`
- Người dùng cuối: dùng bộ cài (`BatTuThieuViHoa_Setup_1.3.0_Windows.exe`, `BatTuThieuViHoa_1.3.0_macOS_<arm64|x86_64>.dmg`) – xem `tai_lieu/HUONG_DAN_SU_DUNG.pdf`.

## Trạng thái
Hoàn thành (GĐ6): toàn bộ 33 sheet (28.581 ô công thức + 288 ô Bảng dữ liệu tính bằng vòng 12 giờ) khớp workbook 100% – xem TIEN_DO.md.

## Cấu trúc
| Thư mục | Nội dung |
|---|---|
| `battu_core/` | Lõi tính thuần (chỉ thư viện chuẩn Python), không phụ thuộc giao diện |
| `battu_core/data/nen/` | Bản chép nguyên văn 04_DU_LIEU_NEN (CSDL, TIET_KHI, Defined Names, dụng thần dự bị) |
| `battu_core/data/hang_so_sheet.json` | Ô hằng các sheet, trích **tự động** từ workbook (MT-02) |
| `battu_core/xl/` | Thư viện hàm Excel theo ngữ nghĩa LibreOffice (cộng/trừ/so sánh gần đúng, MOD/INT, ROUND, TEXT…) |
| `battu_core/sheets/` | **Mỗi sheet = một module**, SINH TỰ ĐỘNG từ công thức workbook – không sửa tay |
| `battu_core/engine.py` | Tính từng ô theo nhu cầu, có nhớ đệm; trả `{sheet: {ô: giá trị}}` đúng địa chỉ Excel |
| `battu_core/nhap_lieu.py` | Mô hình lá số → các ô nhập NHAP_LIEU (dương lịch / âm lịch / can chi) |
| `battu_core/lich_am.py` | Âm ↔ dương lịch (chép nguyên văn lunar.py + chiều ngược lại) |
| `ui/` | Giao diện PySide6 (font Inter, giấy phép OFL, kèm sẵn) |
| `tools/` | Bộ dịch công thức, trích hằng số, oracle mở rộng, thí nghiệm LibreOffice |
| `tests/` | Bộ kiểm thử T1, T1b, T3, âm lịch; `dap_an_mo_rong/` = đáp án T1b/KRG |
| `export/` | Xuất PDF (Qt, kèm 2 biểu đồ, số trang, phiên bản quy tắc) / XLSX (bố cục workbook) / JSON; in |
| `battu_app.py` | Điểm khởi chạy; `--tu-kiem <thư_mục>` = tự kiểm bản đã cài (dấu vân tay SHA-256 toàn bộ kết quả 3 lá) |
| `dong_goi/` | PyInstaller (`battu.spec`), Inno Setup (`battu_windows.iss`, tiếng Việt), biểu tượng |
| `.github/workflows/dong_goi.yml` | Dựng + kiểm cài/gỡ không mạng trên Windows và macOS (GitHub Actions) |
| `tai_lieu/` | Hướng dẫn sử dụng (PDF) |

## Khi workbook / quy tắc đổi (quy tắc 4)
1. Sửa workbook theo `06_MA_THAM_CHIEU/build_excel` (sau khi chủ dự án duyệt).
2. `tools/dung_lai.sh <workbook.xlsx> <04_DU_LIEU_NEN> tools/pham_vi_gdN.json` – trích lại hằng số và dịch lại công thức.
3. Lấy lại đáp án: `05_KIEM_THU/scripts/oracle_dump.py` và `tools/chay_oracle_mo_rong.sh`.
4. `python -m tests.chay_tat_ca <HO_SO_PM>` – phải đạt 100%.

## Kiểm thử
```
.venv/bin/python -m tests.chay_tat_ca /đường/dẫn/HO_SO_PM BAO_CAO_KIEM_THU.md
```
Không bao giờ `pkill soffice`: mọi script LibreOffice tự tắt đúng tiến trình theo PID.

## Dựng bộ cài (README build)
Chỉ Windows và macOS (Linux phát triển sau). Bộ cài dựng trên máy chính hệ điều hành đó – cách khuyến nghị là GitHub Actions:

1. Đẩy mã nguồn lên kho GitHub (riêng tư). Vào **Actions → dong-goi → Run workflow** (hoặc đẩy thẻ `v1.0.0`).
2. Mỗi máy (Windows, macOS arm64, macOS Intel) tự: chạy kiểm thử T7 từ mã nguồn (570.940 ô phải 100%) → PyInstaller → bộ cài →
   **cài → chạy `--tu-kiem` khi chặn mạng → gỡ → kiểm gỡ sạch**. Bước nào lỗi thì cả lượt đỏ.
3. Tải bộ cài ở mục *Artifacts* của lượt chạy (kèm `kiem_cai_go_*.txt`, `tu_kiem.*`).

Dựng tay (trên chính máy đích, Python 3.12):
```
python -m pip install -r requirements.txt pyinstaller==6.16.0
pyinstaller --noconfirm dong_goi/battu.spec            # → dist/BatTuThieuViHoa/ (Windows) hoặc dist/Bát Tự Thiệu Vĩ Hoa.app (macOS)
# Windows: cài Inno Setup 6.5+ rồi
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" dong_goi\battu_windows.iss   # → dist_cai\BatTuThieuViHoa_Setup_1.3.0_Windows.exe
# macOS:
hdiutil create -volname "Bát Tự Thiệu Vĩ Hoa" -srcfolder "dist/Bát Tự Thiệu Vĩ Hoa.app" -format UDZO BatTuThieuViHoa.dmg
dist/BatTuThieuViHoa/BatTuThieuViHoa --tu-kiem /tmp/tk   # tự kiểm; mã thoát 0 = ĐẠT
```
Khi quy tắc/workbook đổi: chạy `tools/dung_lai.sh`, kiểm thử tổng 100%, rồi `python tools/dau_van_tay.py` (lập lại dấu vân tay tự kiểm)
và `python tools/tao_huong_dan.py <ảnh> tai_lieu/HUONG_DAN_SU_DUNG.pdf` nếu giao diện đổi.

Ghi chú: bộ cài chưa ký số (Windows SmartScreen / macOS Gatekeeper sẽ hỏi lần đầu – xem hướng dẫn). Muốn bỏ cảnh báo cần chứng thư ký mã
(Windows) và tài khoản Apple Developer (macOS) – đề xuất cho bản sau.
