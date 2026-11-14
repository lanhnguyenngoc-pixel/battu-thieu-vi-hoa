#!/bin/bash
# Dựng lại dữ liệu + module sheet từ workbook (dùng khi workbook đổi – quy tắc 4: sửa workbook trước, rồi dịch lại).
# Dùng: tools/dung_lai.sh <workbook.xlsx> <thư_mục_04_DU_LIEU_NEN> [pham_vi.json]
set -e
cd "$(dirname "$0")/.."
PY=${PY:-.venv/bin/python}
$PY -I tools/trich_hang_so.py "$1" "$2" battu_core/data
$PY -I tools/dich_cong_thuc.py "$1" "${3:-tools/pham_vi_hien_tai.json}" battu_core/sheets
$PY -I tools/phu_thuoc.py "$1" battu_core/sheets
$PY -I tools/dich_dinh_dang_dk.py "$1" battu_core/sheets
