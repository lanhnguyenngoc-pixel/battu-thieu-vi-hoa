# -*- coding: utf-8 -*-
"""Bản 6.5 = 6.4 + các quy tắc giai đoạn B/C đã duyệt (15/11/2026).
  Vượng suy : VS-2 (bước 5 kiểm chứng, P1 tr.70–72) · G11 nối VS-2 · VS-3 (độ tin cậy) · B1 (bệnh–mức) · B2 (gần cách tòng)
  Cách cục  : CĐB-1 (tòng nhi + gốc mạnh) · CĐB-3 (độc vượng PP2) · CĐB-4 (chuyên vượng ≥ 5/8 chữ)
  Dụng thần : DT-2a (ứng viên khuyết – xét có mặt theo can lộ + tàng can gốc) · DT-3 (điều hầu < 20%)
  Đại vận   : DV-3 (thuận/nghịch P1 tr.124–125 + ĐB-1 điều hầu, ĐB-2 thương quan gặp quan, mức lên ngôi/khắc dụng)
  Tài lộc   : TL-1 (kết luận tài lộc từng vận)
Mỗi kết luận có 3 ô riêng: KẾT LUẬN (phần mềm in đậm) · LÝ DO (luận giải của công cụ) · CĂN CỨ SÁCH (chữ thường, dòng riêng).
Dùng: python3 tools/step18.py <workbook_6.4.xlsx> <workbook_6.5.xlsx> <cổng>"""
import sys, os, re, time, shutil, subprocess, zipfile as _z
B = '/home/claude/HO_SO_PM/HO_SO_PM/06_MA_THAM_CHIEU/build_excel'
sys.path.insert(0, B)
import xpatch
from xpatch import Pkg
xpatch.SHEETS.update({'KET_LUAN': 4, 'TRANG_CHU': 1, 'MAU_THUAN': 28, 'DUNG_HY_KY': 9, 'DOI_CHIEU': 8, 'PP1_DINH_TINH': 10,
                      'DAI_VAN': 15, 'TAI_LOC': 20, 'KIEM_THU': 27})
import fix_dt, fix_links
SRC, DST, PORT = sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 2618
p = Pkg(SRC)
W = {}
W['VS2'] = 'VS-2 (duyệt 15/11/2026): xét vượng suy PP1 đủ 5 bước, bước 5 kiểm chứng (P1 tr.70–72)'
W['VS3'] = 'VS-3 (duyệt 15/11/2026): độ tin cậy theo sự đồng thuận của hai phương pháp'
W['B'] = 'B1, B2 (duyệt 15/11/2026): phần cần trạng thái thân lấy kết luận cuối'
W['CDB'] = 'CĐB-1, CĐB-3, CĐB-4 (duyệt 15/11/2026): điều kiện bổ sung của cách đặc biệt'
W['DT'] = 'DT-2a, DT-3 (duyệt 15/11/2026): ứng viên khuyết; điều hầu'
W['DV'] = 'DV-3 (duyệt 15/11/2026): luận đại vận = quy tắc chung + quy tắc riêng, một kết luận mỗi vận'
W['TL'] = 'TL-1 (duyệt 15/11/2026): kết luận tài lộc từng vận'
COLS10 = 'CDEFGHIJKL'
COLS8 = 'CDEFGHIJ'


def new(sheet, ref, f, text=False, sf=None, why=''):
    p.new_formula(sheet, ref, f, text=text, style_from=sf, why=why)


def lab(sheet, ref, t, sf, why):
    p.set_text(sheet, ref, t, style_from=sf, why=why, must_be_empty=True)


# =====================================================================================
# 1. PP1_DINH_TINH – VS-2: bước 5 kiểm chứng
# =====================================================================================
S = 'PP1_DINH_TINH'; why = W['VS2']
old91, old92 = p.formula(S, 'C91'), p.formula(S, 'C92')
assert old91.startswith('IF(AND(C62=1,C66=3),"Quá vượng"') and old92.startswith('IF(AND(C62=1,C66=3),"Được lệnh')
lab(S, 'B114', 'I. BƯỚC 5 – KIỂM CHỨNG KẾT LUẬN SƠ BỘ (P1 tr.70–72: điều kiện phải "mạnh, giúp ích nhiều"; phán đoán tổng hợp)', 'B90', why)
lab(S, 'B115', 'Kết luận sơ bộ (bước 1–4)', 'B91', why)
new(S, 'C115', old91, text=True, sf='C91', why=why)
lab(S, 'B116', 'Quy tắc sơ bộ', 'B92', why)
new(S, 'C116', old92, text=True, sf='C92', why=why)
lab(S, 'B117', 'Bước 5 (0 giữ nguyên · −1 đổi thành nhược · 1 đổi thành vượng · 2 giữ vượng nhưng cận ranh giới)', 'B92', why)
SOV = 'ISNUMBER(SEARCH("vượng",$C$115))'
PX = 'ISNUMBER(SEARCH("phân xử",$C$116))'
new(S, 'C117', f'IF({PX},0,IF($C$62=0,IF(AND({SOV},$C$88<0,$D$89>$C$89),-1,IF(AND(NOT({SOV}),$C$88>0,$C$89>$D$89),1,0)),'
                f'IF(AND({SOV},$C$88<0,$D$89>$C$89),2,0)))', sf='C88', why=why)
lab(S, 'B118', 'Bước 5 – lý giải', 'B92', why)
new(S, 'C118',
    'IF(' + PX + ',"Kết luận sơ bộ đã dựa trên phân xử a–d (nửa vượng nửa nhược) nên giữ nguyên.",'
    'CHOOSE($C$117+2,'
    '"Không được lệnh; các điều sinh trợ không mạnh, không giúp ích nhiều: phân xử a–d nghiêng về phe khắc (tổng "&$C$88&") và phe khắc–tiết–hao nhiều chữ hơn ("&$D$89&" so với "&$C$89&") → đổi thành THÂN NHƯỢC.",'
    'IF($C$62=1,"Được lệnh; phân xử a–d và đếm chữ không cùng phủ định kết luận sơ bộ → giữ nguyên.","Phân xử a–d và đếm chữ không cùng phủ định kết luận sơ bộ → giữ nguyên."),'
    '"Không được lệnh nhưng phân xử a–d nghiêng về phe sinh–trợ (tổng "&$C$88&") và phe sinh–trợ nhiều chữ hơn ("&$C$89&" so với "&$D$89&") → đổi thành THÂN VƯỢNG.",'
    '"Được lệnh nên giữ THÂN VƯỢNG (lệnh là quan trọng nhất), nhưng phân xử a–d (tổng "&$C$88&") và đếm chữ ("&$D$89&" chữ phe khắc so với "&$C$89&") đều nghiêng về phe khắc–tiết–hao → vượng cận ranh giới: gặp vận khắc–tiết dễ chuyển nhược."))',
    text=True, sf='C92', why=why)
lab(S, 'B119', 'Bước 5 – căn cứ sách', 'B92', why)
new(S, 'C119', '"P1 tr.70 (quy tắc 3: các điều phải mạnh, giúp ích nhiều); P1 tr.71–72 (so sánh mức độ được sinh – bị khắc; phán đoán tổng hợp)"',
    text=True, sf='C92', why=why)
p.set_formula(S, 'C91', 'IF($C$117=-1,"Thân nhược",IF($C$117=1,"Thân vượng",$C$115))', text=True, why=why)
p.set_formula(S, 'C92', '$C$116&IF(OR($C$117=-1,$C$117=1)," → bước 5 kiểm chứng: "&IF($C$117=-1,"đổi thành thân nhược","đổi thành thân vượng")&" (P1 tr.70–72)","")', text=True, why=why)
# C111: cảnh báo cũ (G21) thay bằng cảnh báo cận ranh giới của VS-2 (gộp G50)
p.set_formula(S, 'C111', 'IF($C$117=2,"⚠ VƯỢNG CẬN RANH GIỚI: "&$C$118,"")', text=True, why=why + ' – thay cảnh báo G21, gộp G50')
p.set_text(S, 'B111', 'Cảnh báo vượng cận ranh giới (bước 5)', why=why)
p.replace_in_formula(S, 'C112', '"mọi điều mạnh chỉ từ một chi"', '"vượng cận ranh giới (bước 5)"', count=1, why=why)
f110 = p.formula(S, 'C110')
p.set_formula(S, 'C110', f110 + '&IF($C$117<>0," Bước 5: "&$C$118,"")', text=True, why=why)

# =====================================================================================
# 2. DOI_CHIEU – VS-3 độ tin cậy; dòng tham khảo cách đọc còn lại
# =====================================================================================
S = 'DOI_CHIEU'; why = W['VS3']
FINAL_NH = '(IF(C33=2,C32,C31)="nhược")'
p.set_formula(S, 'C35', f'IF({FINAL_NH}<>(C32="nhược"),"Thấp",IF(OR(PP1_DINH_TINH!$C$117=2,AND(PP2_DINH_LUONG!$C$101>=0.35,PP2_DINH_LUONG!$C$101<0.45)),"Vừa","Cao"))',
              text=True, why=why)
p.set_formula(S, 'F35', 'CHOOSE(MATCH(C35,{"Cao","Vừa","Thấp"},0),'
              '"Cao: hai phương pháp cùng kết luận, không cận ranh giới. Trên 68 lá số của sách: đúng 34/35.",'
              '"Vừa: hai phương pháp cùng kết luận nhưng vượng cận ranh giới hoặc PP2 sát ngưỡng (35–45%) – nên đối chiếu thêm với sự việc đã xảy ra.",'
              '"Thấp: hai phương pháp khác kết luận – kết luận chính theo khuyến nghị, xem thêm cách đọc còn lại; nên kiểm nghiệm bằng sự việc đã qua.")',
              text=True, why=why)
p.replace_in_formula(S, 'C46', '"kết luận vượng theo quy tắc thứ ba nhưng mọi điều mạnh chỉ từ một chi (xem cảnh báo P1 tr.71 dưới đây) → KÉM chắc chắn"',
                     '"được lệnh nên vượng, nhưng phân xử a–d và đếm chữ đều nghiêng về phe khắc (vượng cận ranh giới) → KÉM chắc chắn"', count=1, why=why)
lab(S, 'B52', 'Tham khảo thêm – cách đọc theo phương pháp còn lại (khi độ tin cậy Thấp)', 'B49', why)
DV_LIST = lambda row: 'TRIM(' + '&'.join(f'IF(DAI_VAN!{c}${row}="Tốt",DAI_VAN!{c}$35&" ("&DAI_VAN!{c}$36&"–"&(DAI_VAN!{c}$36+9)&") ","")' for c in COLS8) + ')'
new(S, 'C52', 'IF(C35<>"Thấp","","Tham khảo thêm – nếu đọc theo "&IF(C33=2,"PP1 (xét lệnh trước)","PP2 (tính lực)")&": thân "&UPPER(IF(C33=2,C31,C32))'
              '&", dụng thần "&IF(C33=2,DUNG_HY_KY!$C$106,DUNG_HY_KY!$D$106)&", hỷ thần "&IF(C33=2,DUNG_HY_KY!$C$107,DUNG_HY_KY!$D$107)'
              '&"; các đại vận thuận theo cách đọc này: "&IF(IF(C33=2,' + DV_LIST(54) + ',' + DV_LIST(69) + ')="","không có",IF(C33=2,' + DV_LIST(54) + ',' + DV_LIST(69) + '))'
              '&". Kết luận chính vẫn theo "&D33&".")', text=True, sf='C49', why=why)
c38 = p.text_of(S, 'C38')
p.set_text(S, 'C38', 'Kiểm chứng bản 6.5 (15/11/2026) trên 68 lá có lời giải vượng suy trong sách: kết luận cuối (PP1 đủ 5 bước + khuyến nghị) đúng 60/68; '
           'PP2 đúng 46/68. Lá hai phương pháp cùng kết luận đúng 91%, khác kết luận đúng 82%. Trích Thiên Tủy (231 lá): PP1 149, PP2 190 – khác biệt trường phái. '
           '(Bản trước: ' + (c38 or '')[:400] + ')', why=why)

# =====================================================================================
# 3. DUNG_HY_KY – CĐB-1/3/4, DT-2a, DT-3, B1
# =====================================================================================
S = 'DUNG_HY_KY'
why = W['CDB']
old72c, old72d = p.formula(S, 'C72'), p.formula(S, 'D72')
lab(S, 'B201', 'M. ĐIỀU KIỆN BỔ SUNG CỦA CÁCH ĐẶC BIỆT (tòng nhi, độc vượng, chuyên vượng)', 'B150', why)
lab(S, 'B202', 'Cách đặc biệt theo điều kiện cơ bản (trước khi xét bổ sung)', 'B72', why)
new(S, 'C202', old72c, text=True, sf='C72', why=why)
new(S, 'D202', old72d, text=True, sf='D72', why=why)
lab(S, 'B203', 'Nhật can có gốc mạnh ở chi: Lâm quan, Đế vượng; Trường sinh với can dương (1/0)', 'B62', why)
ST = lambda col: f'INDEX(M_TS,TRUNG_GIAN!$E$5+1,TRUNG_GIAN!${col}$6+1)'
new(S, 'C203', 'IF(OR(' + ','.join(f'{ST(c)}=3,{ST(c)}=4,AND({ST(c)}=0,TRUNG_GIAN!$E$8=0)' for c in 'CDEF') + '),1,0)', sf='C62', why=why)
new(S, 'D203', '$C$203', sf='D62', why=why)
lab(S, 'B204', 'Số chữ hành nhật chủ + hành sinh nhật chủ (4 can, tính cả nhật can + bản khí 4 chi; chi đã hóa tính theo hành hóa)', 'B63', why)
HC = lambda col: f'IF(PP1_DINH_TINH!${col}$27>=0,PP1_DINH_TINH!${col}$27,TRUNG_GIAN!${col}$9)'
PM = lambda x: f'IF(OR({x}=TRUNG_GIAN!$E$7,{x}=MOD(TRUNG_GIAN!$E$7+4,5)),1,0)'
new(S, 'C204', '+'.join([PM(f'TRUNG_GIAN!${c}$7') for c in 'CDEF'] + [PM(HC(c)) for c in 'CDEF']), sf='C63', why=why)
new(S, 'D204', '$C$204', sf='D63', why=why)
lab(S, 'B205', 'Không thành cách đặc biệt vì (rỗng = vẫn thành cách)', 'B72', why)
for c in 'CD':
    new(S, f'{c}205', f'IF(AND({c}202="Tòng nhi",{c}203=1),"nhật can có gốc mạnh ở chi nên không tòng – giữ chính cách",'
                      f'IF({c}202="Tòng vượng (độc vượng)","độc vượng theo số độ không tự thành cách – chỉ thành cách khi đủ điều kiện cách chuyên vượng; giữ chính cách",'
                      f'IF(AND({c}61<>"",{c}202={c}61,{c}204<5),"hành nhật chủ và hành sinh nhật chủ mới có "&{c}204&"/8 chữ, chưa đủ điều kiện tứ trụ hành này nhiều – giữ chính cách","")))',
        text=True, sf=f'{c}72', why=why)
lab(S, 'B206', 'Điều kiện bổ sung – căn cứ sách', 'B72', why)
for c in 'CD':
    new(S, f'{c}206', f'IF({c}205="","",IF({c}202="Tòng nhi","P2 tr.373–375: cách tòng đòi nhật chủ không có khí gốc; các ví dụ tòng nhi của sách nhật can không có gốc mạnh",'
                      f'IF({c}202="Tòng vượng (độc vượng)","P2 tr.353 chỉ gọi tên kết quả tính độ; điều kiện thành cách chuyên vượng ở P2 tr.369–373",'
                      f'"P2 tr.369–373: cách chuyên vượng đòi tứ trụ hành nhật chủ (và hành sinh nó) nhiều – hiểu là quá nửa, ≥ 5/8 chữ")))',
        text=True, sf=f'{c}72', why=why)
p.set_formula(S, 'C72', 'IF($C$205<>"","",$C$202)', text=True, why=why)
p.set_formula(S, 'D72', 'IF($D$205<>"","",$D$202)', text=True, why=why)

why = W['DT']
# DT-2a: có mặt = can lộ (trừ nhật can) + mọi tàng can gốc
lab(S, 'B208', 'N. DỤNG THẦN – ỨNG VIÊN KHUYẾT VÀ ĐIỀU HẦU', 'B150', why)
NHOMS = ['Tỉ kiếp', 'Thực thương', 'Tài', 'Quan sát', 'Ấn']
for g in range(5):
    r = 209 + g
    lab(S, f'B{r}', f'Có mặt (can lộ trừ nhật can + tàng can gốc): {NHOMS[g]}', 'B18', why)
    for c in 'CD':
        X = f'MOD({c}$5+{g},5)'
        new(S, f'{c}{r}', f'IF(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7={X})*(TRUNG_GIAN!$C$4:$F$4<>2))+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23={X})*1)>0,1,0)',
            sf=f'{c}18', why=why)
lab(S, 'B214', 'Dụng thần theo thứ tự ứng viên P1 tr.99 (ứng viên khuyết thì lấy ứng viên sau)', 'B96', why)
for c in 'CD':
    o = p.formula(S, f'{c}96')
    assert o.count(f'INDEX({c}$18:{c}$22,') == 3, o
    new(S, f'{c}214', o.replace(f'INDEX({c}$18:{c}$22,', f'INDEX({c}$209:{c}$213,'), sf=f'{c}96', why=why)
# DT-3: điều hầu
lab(S, 'B215', 'Hành điều hầu cần (P1 tr.76–79: kim, thủy, mộc sinh tháng Hợi–Tí–Sửu–Dần cần hỏa; hỏa, thổ, kim sinh tháng Tị–Ngọ–Mùi cần thủy; −1 = không)', 'B162', why)
M, NH = 'TRUNG_GIAN!$D$6', 'TRUNG_GIAN!$E$7'
DH = (f'IF(AND(OR({M}=11,{M}=0,{M}=1,{M}=2),OR({NH}=0,{NH}=3,{NH}=4)),1,'
      f'IF(AND(OR({M}=5,{M}=6,{M}=7),OR({NH}=1,{NH}=2,{NH}=3)),4,-1))')
new(S, 'C215', DH, sf='C162', why=why)
new(S, 'D215', '$C$215', sf='C162', why=why)
lab(S, 'B216', 'Tỉ lệ hành điều hầu trong tổng độ (PP2)', 'B162', why)
new(S, 'C216', 'IF($C$215<0,"",IF(SUM($D$24:$D$28)=0,0,INDEX($D$24:$D$28,$C$215+1)/SUM($D$24:$D$28)))', sf='C162', why=why)
new(S, 'D216', '$C$216', sf='C162', why=why)
for c in 'CD':
    new(S, f'{c}217', f'IF(AND({c}$91<>"TH",$C$215>=0,N($C$216)<0.2,MOD($C$215-{c}$5,5)<>{c}214),MOD($C$215-{c}$5,5),{c}214)', sf=f'{c}96', why=why)
lab(S, 'B217', 'Dụng thần sau khi xét điều hầu (điều hầu còn ít dưới 20% thì lấy làm dụng thần)', 'B96', why)
lab(S, 'B218', 'Lý do chọn dụng thần theo điều hầu', 'B149', why)
MUA = 'IF(OR(TRUNG_GIAN!$D$6=5,TRUNG_GIAN!$D$6=6,TRUNG_GIAN!$D$6=7),"mùa nóng","mùa lạnh")'
for c in 'CD':
    new(S, f'{c}218', f'IF({c}217={c}214,"","Nhật chủ "&INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&" ("&{MUA}&") cần "'
                      f'&INDEX(HANH_TEN,$C$215+1)&" để điều hầu; "&INDEX(HANH_TEN,$C$215+1)&" chỉ chiếm "&TEXT($C$216,"0%")&" tổng độ nên lấy "'
                      f'&INDEX(HANH_TEN,$C$215+1)&" ("&INDEX(NHOM_TEN,{c}217+1)&") làm dụng thần thay cho "&INDEX(HANH_TEN,MOD({c}$5+{c}214,5)+1)&" ("&INDEX(NHOM_TEN,{c}214+1)&").")',
        text=True, sf='C149', why=why)
lab(S, 'B219', 'Điều hầu – căn cứ sách', 'B149', why)
for c in 'CD':
    new(S, f'{c}219', f'IF({c}218="","","P1 tr.76–79: sinh mùa lạnh cần hỏa sưởi ấm, sinh mùa nóng cần thủy điều hòa; P1 tr.100: điều hầu là một nguyên tắc chọn dụng thần")',
        text=True, sf='C149', why=why)
for c in 'CD':
    p.set_formula(S, f'{c}96', f'{c}217', why=why)
    p.set_formula(S, f'{c}97', f'IF({c}96<0,1,IF(INDEX({c}$209:{c}$213,{c}96+1)=1,0,1))', why=why)
# KET_LUAN C39 (vì sao chọn): nối lý do điều hầu
o = p.formula('KET_LUAN', 'C39'); assert o == 'DUNG_HY_KY!$C$147&CHAR(10)&DUNG_HY_KY!$C$148&CHAR(10)&"⇒ "&DUNG_HY_KY!$C$149', o
p.set_formula('KET_LUAN', 'C39', o + '&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$218,DUNG_HY_KY!$C$218)<>"",CHAR(10)&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$218,DUNG_HY_KY!$C$218),"")'
              '&IF(IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$205,DUNG_HY_KY!$C$205)<>"",CHAR(10)&"Cách cục: "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$202,DUNG_HY_KY!$C$202)&" chưa thành cách – "&IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$205,DUNG_HY_KY!$C$205)&".","")',
              text=True, why=why)

# B1 – bệnh–mức
why = W['B']
o = p.formula(S, 'C176')
oc = 'OR(PP2_DINH_LUONG!$C$101>=0.6,PP2_DINH_LUONG!$C$101<0.3,'
assert oc in o, o
NHF = 'IF($C$154=2,$D$7,$C$7)=0'
p.set_formula(S, 'C176', o.replace(oc, f'OR(AND(PP2_DINH_LUONG!$C$101>=0.6,NOT({NHF})),AND(PP2_DINH_LUONG!$C$101<0.3,{NHF}),'), why=why)
o = p.formula(S, 'C178'); k = '&" (phe mình PP2 "&TEXT(PP2_DINH_LUONG!$C$101,"0%")&")'
assert k in o, o
p.set_formula(S, 'C178', o.replace(k, k[:-2] + 'IF(AND($C$176=1,OR(PP2_DINH_LUONG!$C$101>=0.6,PP2_DINH_LUONG!$C$101<0.3)),'
                                     '" – số độ PP2 lệch ngược chiều kết luận vượng suy cuối nên chỉ xét bệnh nhẹ","")&")'), text=True, why=why)
# C189: không còn vận "Thường"
p.replace_in_formula(S, 'C189', '&" Các vận còn lại: Thường."', '', count=1, why=W['DV'])

# B2 – gần cách tòng (KET_LUAN C35)
o = p.formula('KET_LUAN', 'C35'); k = 'PP2_DINH_LUONG!$C$101<0.15)'
assert o.count(k) == 2, o
p.set_formula('KET_LUAN', 'C35', o.replace(k, 'PP2_DINH_LUONG!$C$101<0.15,IF(DOI_CHIEU!$C$33=2,DUNG_HY_KY!$D$7,DUNG_HY_KY!$C$7)=0)'), text=True, why=why)
# KET_LUAN C42: bỏ cảnh báo G28, G50 cũ → cận ranh giới (VS-2) + tham khảo cách đọc còn lại (VS-3)
o = p.formula('KET_LUAN', 'C42'); assert o.startswith('DOI_CHIEU!C34&IF(AND(ISNUMBER(SEARCH("phân xử",PP1_DINH_TINH!$C$92))') and o.endswith('(giả định G50).","")')
p.set_formula('KET_LUAN', 'C42', 'DOI_CHIEU!C34&IF(PP1_DINH_TINH!$C$111<>"",CHAR(10)&PP1_DINH_TINH!$C$111,"")&IF(DOI_CHIEU!$C$52<>"",CHAR(10)&DOI_CHIEU!$C$52,"")',
              text=True, why=W['VS3'] + ' – thay G28, G50')

# =====================================================================================
# 4. DAI_VAN – DV-3
# =====================================================================================
S = 'DAI_VAN'; why = W['DV']
lab(S, 'B108', 'F. KẾT LUẬN ĐẠI VẬN THEO THIỆU VĨ HOA – quy tắc chung (thuận/nghịch với thân, P1 tr.124–125) + quy tắc riêng của lá số; 5 năm đầu theo can, 5 năm sau theo chi', 'B46', why)
lab(S, 'B109', 'Thân nhược theo PP1 (1/0)', 'B47', why); new(S, 'C109', 'IF(DUNG_HY_KY!$C$7=0,1,0)', sf='C57', why=why)
lab(S, 'B110', 'Thân nhược theo PP2 (1/0)', 'B47', why); new(S, 'C110', 'IF(DUNG_HY_KY!$D$7=0,1,0)', sf='C57', why=why)
lab(S, 'B111', 'Hành điều hầu cần (−1 = không)', 'B47', why); new(S, 'C111', 'DUNG_HY_KY!$C$215', sf='C57', why=why)
lab(S, 'B112', 'Thương quan gặp chính quan lộ trong trụ, không có ấn lộ (1/0)', 'B47', why)
LO = ['TRUNG_GIAN!$C$12', 'TRUNG_GIAN!$D$12', 'TRUNG_GIAN!$F$12']
new(S, 'C112', 'IF(AND(OR(' + ','.join(f'{x}="Chính quan"' for x in LO) + '),NOT(OR(' + ','.join(f'{x}="Chính ấn",{x}="Thiên ấn"' for x in LO) + '))),1,0)', sf='C57', why=why)
lab(S, 'B113', 'Lý do cần điều hầu', 'B47', why)
new(S, 'C113', 'IF($C$111<0,"",INDEX(HANH_TEN,TRUNG_GIAN!$E$7+1)&" sinh tháng "&INDEX(CHI_TEN,TRUNG_GIAN!$D$6+1)&IF($C$111=1,", mùa lạnh, cần hỏa sưởi ấm",", mùa nóng, cần thủy điều hòa"))',
    text=True, sf='C58', why=why)


def diem(c, half, pp):
    """Điểm một nửa vận: ±1 thuận/nghịch; ĐB-1; ĐB-2; ±2 dụng lên ngôi / khắc dụng."""
    h = f'{c}$42' if half == 'can' else f'{c}$43'
    tt = f'{c}$39' if half == 'can' else f'INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,{c}$34+1)+1)'
    nh = '$C$109' if pp == 1 else '$C$110'
    dg = 'DUNG_HY_KY!$C$109' if pp == 1 else 'DUNG_HY_KY!$D$109'
    g = f'MOD({h}-TRUNG_GIAN!$E$7,5)'
    th = f'(OR({g}=0,{g}=4)=({nh}=1))'
    d0 = f'IF({th},1,-1)'
    dA = f'IF(AND($C$111>=0,{h}=$C$111,NOT{th}),1,{d0})'
    dB = f'IF(AND({tt}="Thương quan",$C$112=1),-1,{dA})'
    return f'IF(AND({dg}>=0,{dB}>0,{h}={dg}),2,IF(AND({dg}>=0,{dB}<0,MOD({h}+2,5)={dg}),-2,{dB}))'


ROWS = [(115, 'PP1 – điểm 5 năm đầu (can): ±1 thuận/nghịch, ±2 dụng lên ngôi / khắc dụng', 'can', 1),
        (116, 'PP1 – điểm 5 năm sau (chi)', 'chi', 1),
        (117, 'PP2 – điểm 5 năm đầu (can)', 'can', 2),
        (118, 'PP2 – điểm 5 năm sau (chi)', 'chi', 2)]
for r, t, half, pp in ROWS:
    lab(S, f'B{r}', t, 'B57', why)
    for c in COLS10:
        new(S, f'{c}{r}', diem(c, half, pp), sf=f'{c}57', why=why)
lab(S, 'B119', 'Điểm 5 năm đầu (phương pháp khuyến nghị)', 'B57', why)
lab(S, 'B120', 'Điểm 5 năm sau (phương pháp khuyến nghị)', 'B57', why)
MUC = lambda x: f'CHOOSE({x}+3,"RẤT XẤU","XẤU","BÌNH THƯỜNG","TỐT","RẤT TỐT")'
for c in COLS10:
    new(S, f'{c}119', f'IF(DOI_CHIEU!$C$33=2,{c}117,{c}115)', sf=f'{c}57', why=why)
    new(S, f'{c}120', f'IF(DOI_CHIEU!$C$33=2,{c}118,{c}116)', sf=f'{c}57', why=why)
lab(S, 'B121', 'KẾT LUẬN ĐẠI VẬN', 'B54', why)
lab(S, 'B122', 'Tóm tắt 10 năm', 'B54', why)
lab(S, 'B123', 'Vì sao – 5 năm đầu (can vận)', 'B58', why)
lab(S, 'B124', 'Vì sao – 5 năm sau (chi vận)', 'B58', why)
lab(S, 'B125', 'Căn cứ sách', 'B58', why)
NHC = 'IF(DOI_CHIEU!$C$33=2,$C$110,$C$109)'
DGC = 'DUNG_HY_KY!$C$157'


def vi_sao(c, half):
    h = f'{c}$42' if half == 'can' else f'{c}$43'
    bk = f'INDEX(CSDL!$D$18:$D$29,{c}$34+1)'
    tt = f'{c}$39' if half == 'can' else f'INDEX(M_TT,TRUNG_GIAN!$E$5+1,{bk}+1)'
    d = f'{c}119' if half == 'can' else f'{c}120'
    g = f'MOD({h}-TRUNG_GIAN!$E$7,5)'
    th = f'(OR({g}=0,{g}=4)=({NHC}=1))'
    head = (f'"Can "&INDEX(CAN_TEN,{c}$33+1)&" là "&{tt}&" ("&INDEX(HANH_TEN,{h}+1)&")"' if half == 'can' else
            f'"Chi "&INDEX(CHI_TEN,{c}$34+1)&" (bản khí "&INDEX(CAN_TEN,{bk}+1)&" là "&{tt}&", hành "&INDEX(HANH_TEN,{h}+1)&")"')
    return (head + f'&" – "&IF({th},"thuận","nghịch")&" với thân "&IF({NHC}=1,"nhược","vượng")'
            f'&IF(AND($C$111>=0,{h}=$C$111,NOT{th}),"; nhưng "&INDEX(HANH_TEN,{h}+1)&" là hành điều hầu cần thiết ("&$C$113&") nên chuyển thành tốt","")'
            f'&IF(AND({tt}="Thương quan",$C$112=1),"; nhưng là thương quan gặp chính quan lộ trong trụ, không có ấn hóa giải nên xấu","")'
            f'&IF({d}=2,"; hành "&INDEX(HANH_TEN,{h}+1)&" là dụng thần lên ngôi nên rất tốt",IF({d}=-2,"; hành "&INDEX(HANH_TEN,{h}+1)&" khắc dụng thần "&INDEX(HANH_TEN,{DGC}+1)&" nên rất xấu",""))&"."')


for c in COLS10:
    new(S, f'{c}121', f'IF(({c}119>0)=({c}120>0),{MUC(f"IF(ABS({c}119)>=ABS({c}120),{c}119,{c}120)")},{MUC(f"{c}119")})', text=True, sf=f'{c}54', why=why)
    new(S, f'{c}122', f'IF(({c}119>0)=({c}120>0),"cả 10 năm "&LOWER({c}121),"5 năm đầu "&LOWER({MUC(f"{c}119")})&", 5 năm sau "&LOWER({MUC(f"{c}120")}))', text=True, sf=f'{c}54', why=why)
    new(S, f'{c}123', vi_sao(c, 'can'), text=True, sf=f'{c}58', why=why)
    new(S, f'{c}124', vi_sao(c, 'chi'), text=True, sf=f'{c}58', why=why)
    dh_any = f'OR(AND($C$111>=0,{c}$42=$C$111,NOT(OR(MOD({c}$42-TRUNG_GIAN!$E$7,5)=0,MOD({c}$42-TRUNG_GIAN!$E$7,5)=4)=({NHC}=1))),AND($C$111>=0,{c}$43=$C$111,NOT(OR(MOD({c}$43-TRUNG_GIAN!$E$7,5)=0,MOD({c}$43-TRUNG_GIAN!$E$7,5)=4)=({NHC}=1))))'
    tq_any = f'AND($C$112=1,OR({c}$39="Thương quan",INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,{c}$34+1)+1)="Thương quan"))'
    new(S, f'{c}125', f'"P1 tr.124–125: thân nhược gặp vận sinh trợ, thân vượng gặp vận khắc–tiết–hao là thuận; dụng thần lên ngôi là tốt nhất, vận khắc dụng thần là xấu nhất"'
                      f'&IF({dh_any},"; P1 tr.76–79: sinh mùa lạnh cần hỏa sưởi ấm, sinh mùa nóng cần thủy điều hòa","")'
                      f'&IF({tq_any},"; P2 tr.541–542: thương quan gặp quan tinh là điều kị nhất","")'
                      f'&"; 5 năm đầu theo can, 5 năm sau theo chi (tham khảo Hóa giải toàn thư tr.68)"', text=True, sf=f'{c}58', why=why)
lab(S, 'B126', 'Tài lộc trong vận', 'B54', W['TL'])
lab(S, 'B127', 'Công danh trong vận', 'B54', why)
for c in COLS10:
    if c in COLS8:
        new(S, f'{c}126', f'TAI_LOC!{c}$89&" – "&TAI_LOC!{c}$90', text=True, sf=f'{c}58', why=W['TL'])
    tq_any = f'AND($C$112=1,OR({c}$39="Thương quan",INDEX(M_TT,TRUNG_GIAN!$E$5+1,INDEX(CSDL!$D$18:$D$29,{c}$34+1)+1)="Thương quan"))'
    new(S, f'{c}127', f'IF({c}119>0,"Tốt","Xấu")&" – "&IF({tq_any},"thương quan gặp quan nên dễ mất chức, thị phi","theo xu thế chung của vận")', text=True, sf=f'{c}58', why=why)
# Xu thế đại vận (dòng 54 PP1, 69 PP2) = hướng của DV-3 theo từng phương pháp
for c in COLS10:
    p.set_formula(S, f'{c}54', f'IF({c}115>0,"Tốt","Xấu")', text=True, why=why)
    p.set_formula(S, f'{c}69', f'IF({c}117>0,"Tốt","Xấu")', text=True, why=why)
p.set_text(S, 'B55', 'Quy tắc điều chỉnh P2 tr.281–282 (chỉ tham khảo, không đổi kết luận)', why=why)
p.set_text(S, 'B70', 'Quy tắc điều chỉnh P2 tr.281–282 (chỉ tham khảo, không đổi kết luận)', why=why)
p.set_text(S, 'B54', 'XU THẾ ĐẠI VẬN (hướng của kết luận mục F)', why=why)
p.set_text(S, 'B69', 'XU THẾ ĐẠI VẬN (hướng của kết luận mục F)', why=why)
# C94 – kết luận tổng hợp đại vận
L = lambda cond: 'TRIM(' + '&'.join(f'IF({cond.format(c=c)},{c}$35&" ("&{c}$36&"–"&({c}$36+9)&") ","")' for c in COLS8) + ')'
p.set_formula(S, 'C94', '"Đại vận rất tốt: "&IF(' + L('{c}121="RẤT TỐT"') + '="","không có",' + L('{c}121="RẤT TỐT"') + ')'
              '&"; tốt: "&IF(' + L('{c}121="TỐT"') + '="","không có",' + L('{c}121="TỐT"') + ')'
              '&"; xấu: "&IF(' + L('{c}121="XẤU"') + '="","không có",' + L('{c}121="XẤU"') + ')'
              '&"; rất xấu: "&IF(' + L('{c}121="RẤT XẤU"') + '="","không có",' + L('{c}121="RẤT XẤU"') + ')'
              '&". Đang hành: "&LEFT(KET_LUAN!$C$163,MIN(LEN(KET_LUAN!$C$163),300))&" …"', text=True, why=why)

# KET_LUAN C163 – đại vận đang hành theo kết luận mục F
o = p.formula('KET_LUAN', 'C163'); k = '" tuổi) – PP1: "&INDEX(DAI_VAN!$C$54:$L$54,DAI_VAN!$C$25)'
i = o.find(k); assert i > 0, o
p.set_formula('KET_LUAN', 'C163', o[:i] + '" tuổi) – "&INDEX(DAI_VAN!$C$121:$L$121,DAI_VAN!$C$25)&": "&INDEX(DAI_VAN!$C$122:$L$122,DAI_VAN!$C$25)'
              '&". 5 năm đầu: "&INDEX(DAI_VAN!$C$123:$L$123,DAI_VAN!$C$25)&" 5 năm sau: "&INDEX(DAI_VAN!$C$124:$L$124,DAI_VAN!$C$25)'
              '&" Dấu hiệu: "&INDEX(DAI_VAN!$C$92:$L$92,DAI_VAN!$C$25)))', text=True, why=why)

# =====================================================================================
# 5. TAI_LOC – TL-1
# =====================================================================================
S = 'TAI_LOC'; why = W['TL']
lab(S, 'B87', 'G. TÀI LỘC THEO TỪNG ĐẠI VẬN – KẾT LUẬN CUỐI (xu thế chung của vận + quy tắc riêng về tài)', 'B29', why)
lab(S, 'B88', 'Quy tắc riêng (0 theo xu thế chung · 1 thân vượng, chi vận là tài · 2 thân nhược, can tài, chi trợ thân · 3 thân vượng, can tỉ kiếp · 4 thân nhược, can tài)', 'B32', why)
lab(S, 'B89', 'KẾT LUẬN TÀI LỘC CỦA VẬN', 'B36', why)
lab(S, 'B90', 'Lý do', 'B37', why)
lab(S, 'B91', 'Căn cứ sách', 'B37', why)
NHT = 'IF(DOI_CHIEU!$C$33=2,DAI_VAN!$C$110,DAI_VAN!$C$109)=1'
for c in COLS8:
    g1 = f'MOD(DAI_VAN!{c}$42-TRUNG_GIAN!$E$7,5)'; g2 = f'MOD(DAI_VAN!{c}$43-TRUNG_GIAN!$E$7,5)'
    new(S, f'{c}88', f'IF(AND(NOT({NHT}),{g2}=2),1,IF(AND({NHT},{g1}=2,OR({g2}=0,{g2}=4)),2,IF(AND(NOT({NHT}),{g1}=0),3,IF(AND({NHT},{g1}=2),4,0))))', sf=f'{c}33', why=why)
    new(S, f'{c}89', f'IF(OR({c}88=1,{c}88=2),"Tốt",IF(OR({c}88=3,{c}88=4),"Xấu",IF(DAI_VAN!{c}$119>0,"Tốt","Xấu")))', text=True, sf=f'{c}36', why=why)
    new(S, f'{c}90', f'CHOOSE({c}88+1,"theo xu thế chung của vận","thân vượng, chi vận là tài nên 5 năm sau có tài","vận tài nhưng chi vận giúp thân gánh được tài nên có thể tụ tài",'
                     f'"thân vượng gặp vận tỉ kiếp nên tranh tài, hao tài","thân nhược gặp vận tài nên tiền đến nhưng khó giữ")', text=True, sf=f'{c}36', why=why)
    new(S, f'{c}91', f'CHOOSE({c}88+1,"P1 tr.124–125 (xu thế chung của vận)","P2 tr.522: trong trụ có tài, hành đến đất tài tất đại phát",'
                     f'"P2 tr.416 (VD23): tài vận phải gặp vận thân vượng mới tụ tài; P2 tr.521: thân nhược tài nhiều, hành vận tỉ kiếp mới phát tài",'
                     f'"P2 tr.527: thân vượng hành đến vận tỉ kiếp thì phá tài bại nghiệp","P2 tr.527: thân nhược không thắng tài thì vì của mà sinh nạn")', text=True, sf=f'{c}36', why=why)
    p.set_formula(S, f'{c}36', f'{c}89', text=True, why=why)
p.set_text(S, 'B36', 'KẾT LUẬN VẬN (tài lộc – mục G)', why=why)
p.set_text(S, 'B37', 'Vận tốt về tiền', why=why)
p.set_text(S, 'B38', 'Vận cần cẩn trọng về tiền', why=why)
p.set_formula(S, 'C37', 'TRIM(' + '&'.join(f'IF({c}36="Tốt",DAI_VAN!{c}35&" ("&DAI_VAN!{c}36&"–"&(DAI_VAN!{c}36+9)&" tuổi: "&{c}90&") ","")' for c in COLS8) + ')', text=True, why=why)
p.set_formula(S, 'C38', 'TRIM(' + '&'.join(f'IF({c}36="Xấu",DAI_VAN!{c}35&" ("&DAI_VAN!{c}36&"–"&(DAI_VAN!{c}36+9)&" tuổi: "&{c}90&") ","")' for c in COLS8) + ')', text=True, why=why)

# =====================================================================================
# 6. MAU_THUAN + TRANG_CHU
# =====================================================================================
t = p.text_of('TRANG_CHU', 'B3'); i = t.find(' · Căn cứ'); assert i > 0
p.set_text('TRANG_CHU', 'B3', 'Bản 6.5 – 15/11/2026 (6.4 + VS-2, VS-3, B1, B2, CĐB-1/3/4, DT-2a, DT-3, DV-3, TL-1; 6.4 = G142; 6.3 = G141; 6.2 = G140; 6.1 = G130, G132, G133; SU_KIEN; DOC_NHANH)' + t[i:], why='Bản 6.5')
MT = [
    ('Sửa', 'G11 + VS-2 (duyệt 15/11/2026). Vượng suy PP1 xét đủ 5 bước; bước 5 kiểm chứng: không được lệnh mà phân xử a–d và đếm chữ cùng nghiêng về phe khắc thì đổi thành nhược (và ngược lại); '
            'được lệnh mà cả hai nghiêng về phe khắc thì giữ vượng, ghi cận ranh giới. Khi bước 5 đổi kết luận thì coi như đã phân xử → theo PP1 (G11). Kiểm: 60/68 lá sách (trước 59).',
     'PP1_DINH_TINH dòng 91–92, 110–119; DOI_CHIEU C33–C35', 'Bỏ được cách xét thiếu bước so mức độ (P1 tr.71)', 'PP1_DINH_TINH'),
    ('Thay', 'G21 (cảnh báo quy tắc 3) thay bằng bước 5 của VS-2; G50 ("được lệnh giả") gộp vào cảnh báo vượng cận ranh giới; G54(a) đã giải quyết, (b)(c) còn mở.',
     'PP1_DINH_TINH C111; KET_LUAN C42', 'Một cảnh báo thống nhất thay hai cảnh báo rời', 'PP1_DINH_TINH'),
    ('Thay', 'G28 ("vượng suy mong manh") thay bằng VS-3: độ tin cậy Cao = hai phương pháp cùng kết luận, không cận ranh giới; Vừa = cùng kết luận nhưng cận ranh giới hoặc PP2 35–45%; '
             'Thấp = hai phương pháp khác kết luận → hiện thêm cách đọc theo phương pháp còn lại. Kiểm: lá ghi Cao đúng 34/35.',
     'DOI_CHIEU C35, F35, C52; KET_LUAN C42', 'Độ tin cậy phản ánh đúng tín hiệu đồng thuận', 'DOI_CHIEU'),
    ('Sửa', 'G88 → B1: bệnh "nặng" chỉ khi số độ PP2 cùng chiều kết luận cuối (vượng và ≥ 60%; nhược và < 30%) hoặc quá vượng/quá nhược/độc vượng; trái chiều → bệnh nhẹ, có ghi rõ.',
     'DUNG_HY_KY C176, C178', 'Không lấy % PP2 thay cho trạng thái thân', 'DUNG_HY_KY'),
    ('Sửa', 'G25 → B2: cảnh báo gần cách tòng thêm điều kiện kết luận vượng suy cuối là nhược.', 'KET_LUAN C35', 'Bỏ cảnh báo rơi vào lá kết luận vượng', 'KET_LUAN'),
    ('Thêm', 'CĐB-1: tòng nhi không thành khi nhật can có gốc mạnh ở chi (lâm quan, đế vượng; trường sinh của can dương) – P2 tr.373–375. '
             'CĐB-3: độc vượng PP2 không tự thành cách. CĐB-4: chuyên vượng đòi hành nhật chủ + hành sinh nó ≥ 5/8 chữ (giả định – sách không cho số; mọi ví dụ của sách có 6–8/8).',
     'DUNG_HY_KY dòng 72, 201–206', '12/300 lá ngẫu nhiên chuyển về chính cách', 'DUNG_HY_KY'),
    ('Sửa', 'DT-2a: khi chọn dụng thần chính cách, ứng viên "có mặt" xét theo can lộ (trừ nhật can) + mọi tàng can gốc (kể cả chi đã hợp hóa, khử bì); ứng viên khuyết thì lấy ứng viên sau (P1 tr.99).',
     'DUNG_HY_KY dòng 96–97, 209–214', 'Đúng câu "nếu không có … thì lấy …"', 'DUNG_HY_KY'),
    ('Thêm', 'DT-3: chính cách, sinh mùa lạnh (Hợi–Tí–Sửu–Dần) với kim/thủy/mộc cần hỏa; mùa nóng (Tị–Ngọ–Mùi) với hỏa/thổ/kim cần thủy; hành điều hầu chiếm dưới 20% tổng độ thì lấy làm dụng thần (P1 tr.76–79, 100). '
             'Kiểm: 6/12 lá sách có điều kiện chọn đúng (trước 2/12).',
     'DUNG_HY_KY dòng 96, 215–219', 'Thay đổi lớn nhất: 57/300 lá đổi dụng thần', 'DUNG_HY_KY'),
    ('Thay', 'G61 → DV-3: kết luận đại vận = thuận/nghịch với thân (P1 tr.124–125), 5 năm đầu theo can, 5 năm sau theo chi; quy tắc riêng: điều hầu (P1 tr.76–79), '
             'thương quan gặp quan không có ấn (P2 tr.541–542); mức rất tốt khi dụng thần lên ngôi, rất xấu khi khắc dụng thần. Quy tắc P2 tr.281–282 chỉ còn ở bảng chi tiết. Kiểm: 141/196 vận (trước 52).',
     'DAI_VAN dòng 54, 69, 94, 108–127', 'Không còn hàng loạt vận "Bình thường"', 'DAI_VAN'),
    ('Thêm', 'TL-1: tài lộc từng vận = xu thế chung, trừ khi: thân vượng + chi vận tài → tốt; thân nhược + can tài + chi trợ thân → tốt; thân vượng + can tỉ kiếp → xấu; thân nhược + can tài → xấu (P2 tr.416, 521, 522, 527). Kiểm: 29/35 vận.',
     'TAI_LOC dòng 36–38, 87–91', 'Kết luận tài lộc mỗi vận một câu', 'TAI_LOC'),
]
k0 = 169
assert p.text_of('MAU_THUAN', f'D{k0}') is None
for j, (loai, d, e, f, g) in enumerate(MT):
    k = k0 + j
    p.set_number('MAU_THUAN', f'B{k}', 166 + j, style_from='B168', why='Bản 6.5')
    p.set_text('MAU_THUAN', f'C{k}', loai, style_from='C168', why='Bản 6.5')
    for col, v in zip('DEFG', (d, e, f, g)):
        p.set_text('MAU_THUAN', f'{col}{k}', v, style_from=f'{col}168', why='Bản 6.5')


# =====================================================================================
# 7. KIEM_THU – mẫu kiểm thử: giới tính lá mẫu; vá mẫu tài lộc; MAU_THUAN bổ sung
# =====================================================================================
import json as _json
A = '/home/claude/battu/tai_lieu/giai_doan_A/'
why = 'Sửa mẫu kiểm thử (giai đoạn D, 15/11/2026)'
# 7a. Giới tính: chọn lá mẫu theo tên thì lấy giới tính của lá mẫu (KIEM_THU cột Y) – trước đây mọi lá mẫu bị tính là nam
GT = {}
for f in ('vuong_suy_mo_rong.json', 'dai_van_doi_chieu.json'):
    for x in _json.load(open(A + f)): GT.setdefault(x['ten'], x['g'])
lab('KIEM_THU', 'Y5', 'Giới tính (lá mẫu)', 'X5', why)
n_gt = 0
for r in range(6, 91):
    ten = p.text_of('KIEM_THU', f'B{r}')
    if not ten or ten == 'Tự nhập': continue
    g = 'Nữ' if '(nữ)' in ten else ('Nam' if '(nam)' in ten else GT.get(ten))
    if g:
        p.set_text('KIEM_THU', f'Y{r}', g, style_from=f'X{r}', why=why); n_gt += 1
lab('NHAP_LIEU', 'B22', 'Giới tính dùng để tính (chọn lá mẫu thì lấy giới tính của lá mẫu)', 'B10', why)
new('NHAP_LIEU', 'C22', 'IF($C$13="Tự nhập",$C$10,IFERROR(IF(INDEX(KIEM_THU!$Y$6:$Y$90,MATCH($C$13,KIEM_THU!$B$6:$B$90,0))<>"",INDEX(KIEM_THU!$Y$6:$Y$90,MATCH($C$13,KIEM_THU!$B$6:$B$90,0)),$C$10),$C$10))',
    text=True, sf='C23', why=why)
n_rep = 0
for k in list(p.data):
    if k.startswith('xl/worksheets/sheet') and k.endswith('.xml') and k != 'xl/worksheets/sheet3.xml':
        x = p.data[k].decode('utf-8'); n = x.count('NHAP_LIEU!$C$10') + x.count('NHAP_LIEU!C10')
        if n:
            x = x.replace('NHAP_LIEU!$C$10', 'NHAP_LIEU!$C$22').replace('NHAP_LIEU!C10', 'NHAP_LIEU!$C$22')
            p.data[k] = x.encode('utf-8'); n_rep += n
print('giới tính: lá mẫu', n_gt, '· tham chiếu đổi', n_rep)
# 7b. Vá mẫu tài lộc (KIEM_THU dòng 137–152; duyệt 14/11/2026)
V = _json.load(open(A + 'vá_mau_tai_loc_6.5.json'))
for s_ in V['sua']:
    r = s_['dong']
    for col in 'CDEFG':
        if col in s_:
            assert p.formula('KIEM_THU', f'{col}{r}') is None, (col, r)
            p.set_text('KIEM_THU', f'{col}{r}', s_[col], why='Vá mẫu tài lộc: ' + s_['ly_do'])
# 7c. MAU_THUAN: VD1 PP2 đủ 6 chỗ; các lá sách tự mâu thuẫn
assert p.text_of('MAU_THUAN', 'D6').startswith('VD1 (Ất Tị, Ất Dậu, Kỷ Mùi, Nhâm Thân)')
p.set_text('MAU_THUAN', 'D6', 'VD1 (Ất Tị, Ất Dậu, Kỷ Mùi, Nhâm Thân) – sách sai ở 6 chỗ so với chính quy tắc của sách: (1) Kỷ ngày ghi 36 "không bị khắc" – thực ra bị Ất tháng khắc gần −12, Ất năm khắc cách −6 (P2 tr.339–340) → 18; '
           '(2) Ất năm chi xì hơi −6 (P2 tr.341) → 30; (3) Ất tháng bị chi khắc −18 → 18 (bảng lại ghi tổng 66); (4) Nhâm giờ ghi "bị Dậu khắc −18" – Dậu không ở trụ giờ, kim lại sinh thủy → 24; '
           '(5) Tị bán hợp Tị–Dậu không hóa phải khử bì, chỉ còn bản khí Bính 18 (như sách làm ở VD5); (6) bỏ sót Đinh trong Mùi; tổng Kim không cộng ra được (đúng: 57,6). '
           'Tính đúng: phe mình 38,3% → nhược (sách: vượng). PP1 vượng (được lệnh Trường sinh), khớp lời sách.', why='Bản 6.5')
MT2 = [
    ('Mâu thuẫn', 'Vượng suy – lá sách tự mâu thuẫn, công cụ giữ quy tắc: VD3 tr.398 (sách "không gốc" – nhưng giờ Tí là lộc của Quý); VD6 tr.401 và VD25 tr.417 (sách tính chi đang hợp không hóa, trái P1 tr.62 "đứng độc lập"; '
                  'riêng phần tính độ P2 tr.350 lại tính chi hợp không hóa theo bản khí); VD9 tr.404 (theo quy tắc 5c của chính sách là nhược); VD14 tr.408 (lời sách vừa "mộc dục → thân vượng" vừa "nhật nguyên yếu").',
     'Giữ quy tắc đã thống nhất; kết luận khác sách ở 5 lá này', 'P1 tr.62, 70–72; P2 tr.398–417', 'PP1_DINH_TINH'),
    ('Mâu thuẫn', 'Cách đặc biệt – lá đủ điều kiện của sách nhưng tác giả luận theo chính cách: VD4 tr.400 (nhuận hạ, đủ 4 điều kiện tr.372–373); VD27 tr.419 (hóa mộc, gần giống hệt hóa mộc VD1 tr.375); '
                  'VD40 tr.430 và P1-VD10 (tòng tài, đủ điều kiện tr.373); C17-5 tr.517 (hóa mộc); C21-5 tr.562 (tòng sát – có thực thương ở chi, nhưng tòng sát VD1 của sách cũng có).',
     'Giữ điều kiện cách của sách (P2 tr.369–375)', 'P2 tr.369–375, 400, 419, 430, 517, 562', 'DUNG_HY_KY'),
]
k1 = k0 + len(MT)
for j, (loai, d, e, f, g) in enumerate(MT2):
    k = k1 + j
    p.set_number('MAU_THUAN', f'B{k}', 166 + len(MT) + j, style_from='B168', why='Bản 6.5')
    p.set_text('MAU_THUAN', f'C{k}', loai, style_from='C168', why='Bản 6.5')
    for col, v in zip('DEFG', (d, e, f, g)):
        p.set_text('MAU_THUAN', f'{col}{k}', v, style_from=f'{col}168', why='Bản 6.5')

tmp = '/tmp/s18_tmp.xlsx'; p.save(tmp)
print('vá xong', len(p.log), 'ô', flush=True)
if os.environ.get('CHI_VA'):
    shutil.copy(tmp, DST); sys.exit(0)

# --- tính lại bằng LibreOffice (chỉ tắt đúng PID)
import uno
from com.sun.star.beans import PropertyValue
pr = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/s18prof{PORT}',
                       f'--accept=socket,host=localhost,port={PORT};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
print('soffice PID', pr.pid, flush=True)
desk = None
try:
    local = uno.getComponentContext(); res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(90):
        try: ctx = res.resolve(f'uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext'); break
        except Exception: time.sleep(1)
    desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
    doc = None
    for _ in range(8):
        try: doc = desk.loadComponentFromURL('file://' + tmp, '_blank', 0, (pv('Hidden', True),))
        except Exception: doc = None
        if doc is not None: break
        time.sleep(5)
    doc.calculateAll(); doc.store(); doc.close(True)
finally:
    try: desk and desk.terminate()
    except Exception: pass
    time.sleep(2)
    if pr.poll() is None: pr.kill()
shutil.copy(tmp, DST)
print('dt', fix_dt.fix(DST, [117 + 14 * i for i in range(12)]), 'links', fix_links.fix(DST))
_zi = _z.ZipFile(DST); _items = _zi.infolist(); _data = {i.filename: _zi.read(i.filename) for i in _items}; _zi.close(); _n = 0
for _k in list(_data):
    if _k.startswith('xl/worksheets/sheet') and _k.endswith('.xml'):
        _x = _data[_k].decode('utf-8')
        _y = re.sub(r'(<row [^>]*ht=")([\d.]+)(")', lambda m: m.group(1) + ('409.5' if float(m.group(2)) > 409.5 else m.group(2)) + m.group(3), _x)
        if _y != _x: _n += 1; _data[_k] = _y.encode('utf-8')
with _z.ZipFile(DST, 'w', _z.ZIP_DEFLATED) as _zo:
    for _i in _items: _zo.writestr(_i, _data[_i.filename])
print('clamp', _n, '→', DST)
