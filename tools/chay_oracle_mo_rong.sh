#!/bin/bash
# Chạy lần lượt các biến thể oracle mở rộng (mỗi biến thể một tiến trình soffice riêng, tự tắt theo PID)
WB=$1; KT=$2; OUT=$3; shift 3
PORT=2200
for B in "$@"; do
  PORT=$((PORT+1))
  python3 -I "$(dirname "$0")/oracle_mo_rong.py" "$WB" "$KT/cases_dap_an_o.json" "$(dirname "$0")/../tests/bien_the.json" "$B" "$OUT/$B" $PORT > "$OUT/$B.log" 2>&1
  echo "XONG $B $(ls $OUT/$B | wc -l) lá"
done
