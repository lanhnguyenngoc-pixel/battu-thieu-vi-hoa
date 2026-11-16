# -*- coding: utf-8 -*-
# Phần thọ – yểu của bản 6.7 (duyệt 16/11/2026) – được tools/step19.py nạp (exec) khi đặt biến THO_YEU=1.
# Bước 1 (Thiệu Vĩ Hoa): dấu hiệu mệnh thọ / mệnh yểu P1 tr.172 + bảng P1 tr.173 → bên nhiều dấu hiệu hơn là kết luận gốc.
# Bước 2 (luận thêm, không dùng ngưỡng số dấu hiệu): nguyệt đức, thiên đức, sức khỏe chung, bệnh có cứu (P1 tr.97, 114–117),
#          các quy tắc Hóa giải toàn thư HS1–HS6 (HG tr.201–202) → nhẹ bớt / vững hơn; sức khỏe chung bất lợi, bệnh không cứu → nặng hơn.
import re as _re
S = 'MENH_CUC'; why = 'Bản 6.7 (duyệt 16/11/2026): thọ – yểu theo P1 tr.172–173, luận thêm theo thần sát, sức khỏe, cứu ứng, Hóa giải toàn thư'
D, HY, KY, ST = 'DUNG_HY_KY!$C$157', 'DUNG_HY_KY!$C$158', 'DUNG_HY_KY!$C$159', 'DUNG_HY_KY!$C$155'
TTH = 'IF(DUNG_HY_KY!$C$154=2,DUNG_HY_KY!$D$6,DUNG_HY_KY!$C$6)'
NH = 'TRUNG_GIAN!$E$7'
L = lambda h: '(INDEX(DUNG_HY_KY!$D$24:$D$28,(%s)+1)/SUM(DUNG_HY_KY!$D$24:$D$28))' % h
P = lambda h: '(SUMPRODUCT((TRUNG_GIAN!$C$7:$F$7=(%s))*1)+SUMPRODUCT((TRUNG_GIAN!$C$21:$F$23=(%s))*1)>0)' % (h, h)
CAP = [('C', 'D'), ('C', 'E'), ('C', 'F'), ('D', 'E'), ('D', 'F'), ('E', 'F')]
XUNG = lambda a, b: 'ABS(TRUNG_GIAN!$%s$6-TRUNG_GIAN!$%s$6)=6' % (a, b)
HOI = lambda a: 'AND(PP1_DINH_TINH!$%s$27>=0,OR(PP1_DINH_TINH!$%s$27=%s,PP1_DINH_TINH!$%s$27=%s))' % (a, a, D, a, HY)
XUNG_HAI = 'OR(' + ','.join('AND(%s,NOT(OR(%s,%s)))' % (XUNG(a, b), HOI(a), HOI(b)) for a, b in CAP) + ')'
CHAM = lambda a: 'OR(TRUNG_GIAN!$%s$9=%s,TRUNG_GIAN!$%s$9=%s)' % (a, D, a, HY)
XUNG_DUNG = 'OR(' + ','.join('AND(%s,NOT(OR(%s,%s)),OR(%s,%s))' % (XUNG(a, b), HOI(a), HOI(b), CHAM(a), CHAM(b)) for a, b in CAP) + ')'
HANG = 'IF(DOI_CHIEU!$C$33=2,DAI_VAN!$C$69:$J$69,DAI_VAN!$C$54:$J$54)'
SK1 = 'IFERROR(LEFT(SUC_KHOE!$J$6,FIND(CHAR(10),SUC_KHOE!$J$6)-1),SUC_KHOE!$J$6)'
CO_BENH = 'AND(DUNG_HY_KY!$C$160="",DUNG_HY_KY!$C$156>=0,DUNG_HY_KY!$C$173>0)'
CUU = 'DUNG_HY_KY!$D$173+DUNG_HY_KY!$E$173+DUNG_HY_KY!$F$173'
q = lambda s: '"' + s + '"'
DONG = [
    # (dòng, nhãn, điều kiện, câu đưa vào lý do – chuỗi công thức)
    (97, 'Thọ 1 – thân vượng, dụng thần được sinh có lực (P1 tr.172)', 'AND(%s=1,%s>=0,%s>=0.2,%s>0)' % (ST, D, L(D), L('MOD(%s+4,5)' % D)), q('thân vượng, dụng thần được sinh có lực')),
    (98, 'Thọ 2 – thân nhược nhưng có ấn nắm quyền (P1 tr.172)', 'AND(%s=0,TRUNG_GIAN!$D$9=MOD(%s+4,5))' % (ST, NH), q('thân nhược nhưng ấn nắm lệnh tháng')),
    (99, 'Thọ 3 – gặp được các vận hỷ, dụng thần: đại vận tốt nhiều hơn xấu (P1 tr.172; P1 tr.124–125)', 'COUNTIF(%s,"Tốt")>COUNTIF(%s,"Xấu")' % (HANG, HANG), q('đại vận tốt nhiều hơn xấu')),
    (100, 'Thọ 4 – tứ trụ không bị xung, hoặc xung mà hội về dụng, hỷ (P1 tr.172)', 'NOT(%s)' % XUNG_HAI, q('tứ trụ không bị xung phá')),
    (101, 'Thọ 5 – ngũ hành đủ (P1 tr.172)', 'AND(' + ','.join(P(h) for h in range(5)) + ')', q('ngũ hành đủ')),
    (102, 'Thọ 6 – bảng thực thương – ấn nghiêng thọ (P1 tr.173)', '$U$21="THỌ"', q('bảng thực thương – ấn nghiêng thọ')),
    (103, 'Yểu 1 – dụng thần nhược lại bị khắc, kỵ thần có lực (P1 tr.172)', 'AND(%s>=0,%s<0.2,%s>%s,%s>=0.2)' % (D, L(D), L('MOD(%s+3,5)' % D), L(D), L(KY)), q('dụng thần nhược lại bị khắc, kỵ thần có lực')),
    (104, 'Yểu 2 – kỵ thần và dụng thần hỗn chiến (P1 tr.172)', 'AND(%s>=0,%s>=0,OR(MOD(%s+2,5)=%s,MOD(%s+2,5)=%s),%s>=0.2,%s>=0.2)' % (D, KY, KY, D, D, KY, L(D), L(KY)), q('kỵ thần và dụng thần hỗn chiến')),
    (105, 'Yểu 3 – thân quá vượng mà không bị khắc, tiết, hao (dụng thần dưới 20%) (P1 tr.172)', 'AND(OR(ISNUMBER(SEARCH("Quá vượng",%s)),ISNUMBER(SEARCH("Độc vượng",%s))),OR(%s<0,%s<0.2))' % (TTH, TTH, D, L('MAX(%s,0)' % D)), q('thân quá vượng mà không bị khắc tiết hao')),
    (106, 'Yểu 4 – thân quá nhược (P1 tr.172)', 'ISNUMBER(SEARCH("Quá nhược",%s))' % TTH, q('thân quá nhược')),
    (107, 'Yểu 5 – xung hợp bất lợi: chi mang dụng thần, hỷ thần bị xung (P1 tr.172)', 'AND(%s>=0,%s)' % (D, XUNG_DUNG), q('dụng thần hoặc hỷ thần bị xung')),
    (108, 'Yểu 6 – thân nhược gặp nhiều vận bất lợi (P1 tr.172)', 'AND(%s=0,COUNTIF(%s,"Xấu")>COUNTIF(%s,"Tốt"))' % (ST, HANG, HANG), q('thân nhược mà đại vận xấu nhiều hơn tốt')),
    (109, 'Yểu 7 – kim hàn thủy lạnh, hỏa nóng thổ táo: cần điều hầu mà thiếu (P1 tr.172)', 'AND(DUNG_HY_KY!$C$215>=0,N(DUNG_HY_KY!$C$216)<0.2)', q('mùa sinh cần điều hầu mà hành điều hầu thiếu')),
    (110, 'Yểu 8 – tứ trụ thiên về khô, thiếu thủy (P1 tr.171–172)', 'NOT(%s)' % P(4), q('tứ trụ khô, thiếu thủy')),
    (111, 'Yểu 9 – bảng thực thương – ấn nghiêng yểu (P1 tr.173)', '$U$21="YỂU"', q('bảng thực thương – ấn nghiêng yểu')),
    (115, 'Luận thêm (nhẹ bớt) – nguyệt đức có hiệu lực (HG tr.201)', 'THAN_SAT!$H$86="Có"', q('có nguyệt đức')),
    (116, 'Luận thêm (nhẹ bớt) – thiên đức có hiệu lực (HG tr.201)', 'THAN_SAT!$H$85="Có"', q('có thiên đức')),
    (117, 'Luận thêm (nhẹ bớt) – sức khỏe chung thuận (P2 chương 21)', 'OR(LEFT(SUC_KHOE!$J$6,5)="THUẬN",LEFT(SUC_KHOE!$J$6,13)="NGHIÊNG THUẬN")', '"sức khỏe chung "&LOWER(%s)' % SK1),
    (118, 'Luận thêm (nhẹ bớt) – bệnh có cứu: kỵ khắc dụng nhưng có thuốc (P1 tr.97, 114–117)', 'AND(%s,%s>0)' % (CO_BENH, CUU), q('bệnh có cứu (có thuốc)')),
    (119, 'Luận thêm (nhẹ bớt) – thân vượng gặp ấn, bình sinh ít bệnh (HG tr.201)', 'SUC_KHOE!$L$219="ĐỦ"', q('thân vượng gặp ấn, bình sinh ít bệnh')),
    (120, 'Luận thêm (nhẹ bớt) – tỉ kiếp, tài đều nhiều, khỏe mạnh (HG tr.201)', 'SUC_KHOE!$L$220="ĐỦ"', q('tỉ kiếp, tài đều nhiều')),
    (121, 'Luận thêm (nhẹ bớt) – tam hợp/tam hội cục sinh thân, giúp thân (HG tr.201)', 'SUC_KHOE!$L$222="ĐỦ"', q('có cục sinh thân, giúp thân')),
    (122, 'Luận thêm (nhẹ bớt) – thân yếu gặp ấn sinh thân (HG tr.202)', 'SUC_KHOE!$L$223="ĐỦ"', q('thân yếu gặp ấn sinh thân')),
    (123, 'Luận thêm (nhẹ bớt) – nguyệt ấn sinh thân (HG tr.202)', 'SUC_KHOE!$L$224="ĐỦ"', q('nguyệt ấn sinh thân')),
    (124, 'Luận thêm (nặng hơn) – sức khỏe chung bất lợi (P2 chương 21)', 'OR(LEFT(SUC_KHOE!$J$6,7)="BẤT LỢI",LEFT(SUC_KHOE!$J$6,15)="NGHIÊNG BẤT LỢI")', '"sức khỏe chung "&LOWER(%s)' % SK1),
    (125, 'Luận thêm (nặng hơn) – bệnh không có cứu: kỵ khắc dụng mà không có thuốc (P1 tr.97, 114–117)', 'AND(%s,%s=0)' % (CO_BENH, CUU), q('bệnh không có cứu (không thuốc)')),
]
DUNG_MOI = {r for r, *_ in DONG} | set(range(96, 129))


def F(ref, f, text=False):
    p.set_formula(S, ref, f, text=text, style_from='C21', why=why, must_exist=False)


def T(ref, t):
    p.set_text(S, ref, t, style_from='B21', why=why)


# xóa mọi ô cũ ở dòng 96–130 (khối 6.6) để khối mới không lẫn nội dung cũ
x = p.get(S)
for m in _re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>', x, _re.S):
    if 96 <= int(m.group(2)) <= 130 and m.group(3): p.set_text(S, m.group(1) + m.group(2), '', why=why)
T('B96', 'THỌ – YỂU (bản 6.7). Bước 1 – Thiệu Vĩ Hoa: dấu hiệu mệnh thọ, mệnh yểu (P1 tr.172) và bảng thực thương – ấn (P1 tr.173); bên nhiều dấu hiệu hơn là kết luận. '
          'Bước 2 – luận thêm: thần sát, sức khỏe chung, cứu ứng, Hóa giải toàn thư (HG tr.201–202) cho biết được hóa giải, nhẹ bớt hay nặng hơn – không dùng ngưỡng số dấu hiệu.')
CAU = {}
for r, nhan, f, cau in DONG:
    T('B%d' % r, nhan); F('C%d' % r, 'IF(%s,1,0)' % f); CAU[r] = cau
T('B112', 'Số dấu hiệu mệnh thọ'); F('C112', 'SUM(C97:C102)')
T('B113', 'Số dấu hiệu mệnh yểu'); F('C113', 'SUM(C103:C111)')
T('B114', 'BƯỚC 1 – KẾT LUẬN THEO THIỆU VĨ HOA (bên nhiều dấu hiệu hơn)'); F('C114', 'IF(C112>C113,"THỌ",IF(C113>C112,"YỂU","CÂN BẰNG"))', True)
ds = lambda rs: 'MID(' + '&'.join('IF(C%d=1,", "&%s,"")' % (r, CAU[r]) for r in rs) + ',3,999)'
G, N_ = 'SUM(C115:C123)', 'SUM(C124:C125)'
T('B126', 'THỌ – YỂU: KẾT LUẬN CUỐI (sau bước 2)')
F('C126', 'C114&IF(AND(%s>0,%s=0)," – "&IF(C114="YỂU","ĐƯỢC HÓA GIẢI MỘT PHẦN, NHẸ BỚT",IF(C114="THỌ","VỮNG HƠN","NGHIÊNG THUẬN")),'
          'IF(AND(%s>0,%s=0)," – "&IF(C114="THỌ","GIẢM BỚT",IF(C114="YỂU","NẶNG HƠN","NGHIÊNG BẤT LỢI")),'
          'IF(AND(%s>0,%s>0)," – CÓ CẢ YẾU TỐ HÓA GIẢI LẪN YẾU TỐ LÀM NẶNG, TỐT XẤU TÙY VẬN","")))' % (G, N_, N_, G, G, N_), True)
T('B127', 'Lý do')
F('C127', '"Dấu hiệu mệnh thọ ("&C112&"): "&IF(C112=0,"không có",%s)&". Dấu hiệu mệnh yểu ("&C113&"): "&IF(C113=0,"không có",%s)&"."'
          '&IF(%s>0," Yếu tố hóa giải, làm nhẹ: "&%s&".","")&IF(%s>0," Yếu tố làm nặng: "&%s&".","")'
          '&IF(%s+%s=0," Không có yếu tố hóa giải hay làm nặng – giữ nguyên kết luận bước 1.","")'
          % (ds(range(97, 103)), ds(range(103, 112)), G, ds(range(115, 124)), N_, ds((124, 125)), G, N_), True)
T('B128', 'Căn cứ sách')
F('C128', '"P1 tr.171: người thọ nguyên khí dày, dụng thần được sinh có lực; người yểu nguyên khí mỏng, dụng thần được ít khí, tứ trụ thiên về khô. '
          'P1 tr.172: cách xem mệnh thọ, mệnh yểu. P1 tr.173: thọ – yểu theo thực thương, ấn. P1 tr.97, 114–117: bệnh – thuốc (cứu ứng). '
          'Hóa giải toàn thư tr.201–202: thiên nguyệt nhị đức, thân vượng gặp ấn, tỉ kiếp tài đều nhiều, cục sinh thân, thân yếu gặp ấn, nguyệt ấn sinh thân → thân thể khỏe. '
          'Công cụ không kết luận năm mất (P1 tr.201)."', True)
# các trang khác dùng chung kết luận cuối
p.replace_in_formula('LUAN_GIAI', 'F119', 'MENH_CUC!$U$21', 'MENH_CUC!$C$126', why=why)
p.set_formula('KHONG_RO_GIO', 'W27', 'MENH_CUC!$C$126', text=True, why=why)
o = p.formula('KET_LUAN', 'C143'); i = o.find('&"► Ý nghĩa cụ thể: "'); assert i > 0
rest = o[i:].replace('MENH_CUC!$U$21="YỂU"', 'MENH_CUC!$C$114="YỂU"').replace('MENH_CUC!$U$21="THỌ"', 'MENH_CUC!$C$114="THỌ"')
p.set_formula('KET_LUAN', 'C143', 'MENH_CUC!$C$126&CHAR(10)&MENH_CUC!$C$127&CHAR(10)&"Bảng thực thương – ấn (P1 tr.173): "&MENH_CUC!D21&CHAR(10)&MENH_CUC!G21&CHAR(10)&MENH_CUC!K21' + rest, text=True, why=why)
p.replace_in_formula('KET_LUAN', 'C135', '"Thọ – yểu (P1 tr.173): "', '"Thọ – yểu (P1 tr.172–173): "', count=1, why=why)
k = k0 + 2
p.set_number('MAU_THUAN', 'B%d' % k, so0 + 3, style_from='B180', why='Bản 6.7')
p.set_text('MAU_THUAN', 'C%d' % k, 'Thay', style_from='C180', why='Bản 6.7')
for col, v in zip('DEFG', ('Thọ – yểu: bỏ quy tắc 6.6 ("căn cứ ngược chiều" khi ≥ 2 dấu hiệu). Bước 1 theo Thiệu Vĩ Hoa: dấu hiệu mệnh thọ, mệnh yểu (P1 tr.172) và bảng '
                           'thực thương – ấn (P1 tr.173); bên nhiều dấu hiệu hơn là kết luận. Bước 2 luận thêm, không dùng ngưỡng số dấu hiệu: nguyệt đức, thiên đức, '
                           'sức khỏe chung thuận, bệnh có cứu, Hóa giải toàn thư HS1–HS6 (tr.201–202) → nhẹ bớt / vững hơn; sức khỏe chung bất lợi, bệnh không cứu '
                           '→ nặng hơn. Mốc 20% (dụng có lực / dụng nhược, kỵ có lực, điều hầu thiếu) = tỉ lệ trung bình một hành (1/5). Hóa giải toàn thư trong công cụ chỉ có '
                           'phần sức khỏe tr.201–202; nếu sách có mục riêng về thọ – yểu, cứu thì cần bổ sung trang.',
                           'MENH_CUC dòng 96–128; KET_LUAN C135, C143; LUAN_GIAI F119; KHONG_RO_GIO W27', 'Mọi trang dùng chung một kết luận', 'MENH_CUC')):
    p.set_text('MAU_THUAN', '%s%d' % (col, k), v, style_from='%s180' % col, why='Bản 6.7')
t = p.text_of('TRANG_CHU', 'B3')
p.set_text('TRANG_CHU', 'B3', t.rstrip('. ') + '; thọ – yểu theo P1 tr.172–173 kèm luận thêm (thần sát, sức khỏe, cứu ứng, Hóa giải toàn thư).', why='Bản 6.7')
