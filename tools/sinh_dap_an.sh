#!/bin/bash
# Sinh lại toàn bộ đáp án LibreOffice cho một bản workbook. Dùng: bash tools/sinh_dap_an.sh 6.8 <workbook.xlsx>
# – 2 làn song song, mỗi làn một cổng LibreOffice riêng (script tự tắt theo PID)
BAN=${1:?ban, vd 6.8}; NGUON=${2:?workbook}
mkdir -p /tmp/wb$BAN; cp "$NGUON" /tmp/wb$BAN/wb_A.xlsx; cp "$NGUON" /tmp/wb$BAN/wb_B.xlsx
WB=/tmp/wb$BAN/wb_A.xlsx; WBB=/tmp/wb$BAN/wb_B.xlsx
HS=/home/claude/HS$BAN; KT=$HS/05_KIEM_THU
mkdir -p $KT/dap_an_o $KT/ket_qua_$BAN; ln -sfn /home/claude/HO_SO_PM/HO_SO_PM/09_TAI_LIEU_BO_SUNG $HS/09_TAI_LIEU_BO_SUNG
for x in bo_la cases_dap_an_o.json hop_xung lich scripts; do ln -sfn /home/claude/HO_SO_PM/HO_SO_PM/05_KIEM_THU/$x $KT/$x; done; S=/home/claude/HO_SO_PM/HO_SO_PM/05_KIEM_THU/scripts
T=/home/claude/battu/tests; MR=$T/dap_an_mo_rong; OM=/home/claude/battu/tools/oracle_mo_rong.py
cd /home/claude/battu
mr() { rm -rf $MR/$1; python3 $OM $WBB $2 $T/bien_the.json $1 $MR/$1 $3 > $MR/$1.log 2>&1; echo "$1 $(ls $MR/$1 2>/dev/null | wc -l) $(date +%H:%M)" >> /tmp/o_${BAN}_tien_do.txt; }
mra() { rm -rf $MR/$1; python3 $OM $WB $T/cases501_a.json $T/bien_the.json $1 $MR/$1 $2 > $MR/${1}_a.log 2>&1; python3 $OM $WB $T/cases501_b.json $T/bien_the.json $1 $MR/$1 $3 > $MR/${1}_b.log 2>&1; echo "$1 $(ls $MR/$1 | wc -l) $(date +%H:%M)" >> /tmp/o_${BAN}_tien_do.txt; }
lanA() {
  python3 $S/oracle_dump.py $WB $KT/cases_dap_an_o.json $KT/dap_an_o 2601 > /tmp/o_${BAN}_dapan.log 2>&1
  echo "dap_an_o $(ls $KT/dap_an_o | wc -l) $(date +%H:%M)" >> /tmp/o_${BAN}_tien_do.txt
  mra T2_501 2611 2612
  for g in GD3 GD4 GD5; do mra T2_501_$g 2613 2614; done
}
lanB() {
  sleep 20
  for f in cases_G1.json cases_books_VN.json cases_trang49.json; do python3 $S/uno_run59.py $WBB $KT/bo_la/$f $KT/ket_qua_$BAN/res_$f 2621 >> /tmp/o_${BAN}_run59.log 2>&1; done
  echo "run59 $(ls $KT/ket_qua_$BAN | wc -l) $(date +%H:%M)" >> /tmp/o_${BAN}_tien_do.txt
  rm -rf $MR/T7_NGAU_NHIEN; python3 $S/oracle_dump.py $WBB $T/t7_ngau_nhien.json $MR/T7_NGAU_NHIEN 2622 > $MR/T7_NGAU_NHIEN.log 2>&1
  echo "T7_NGAU_NHIEN $(ls $MR/T7_NGAU_NHIEN | wc -l) $(date +%H:%M)" >> /tmp/o_${BAN}_tien_do.txt
  mr T7_SU_KIEN $T/t7_ngau_nhien.json 2623; mr T10_BO_SUNG $T/t7_ngau_nhien.json 2624
  for v in KRG0 B1_KHONG_RO_GIO B1b_KHONG_RO_GIO B2_THEO_PP1 B3_THEO_PP2 B4_PA1 B5_SU_KIEN KRG_VONG; do mr $v $KT/cases_dap_an_o.json 2625; done
  mr T8C_PA1 $T/t8_chuan.json 2626; mr T8C_PA2 $T/t8_chuan.json 2627
}
: > /tmp/o_${BAN}_tien_do.txt; : > /tmp/o_${BAN}_run59.log
lanA & lanB & wait
echo TAT_CA_XONG $(date +%H:%M) >> /tmp/o_${BAN}_tien_do.txt
