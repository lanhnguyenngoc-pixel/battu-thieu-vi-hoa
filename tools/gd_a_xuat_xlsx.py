import json, collections
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
G = json.load(open('/home/claude/battu/tai_lieu/giai_doan_A/dap_an_tong.json'))
LV = [('TC','Tính cách – tổng thể'),('SN','Sự nghiệp – quan vận'),('TL','Tài lộc'),('HN','Hôn nhân'),
      ('CC','Con cái'),('LT','Cha mẹ – anh em'),('SK','Sức khỏe (tham khảo)')]
NH = {'T':'Tốt','V':'Vừa','K':'Kém'}
MAU = {'T':'C6EFCE','V':'FFEB9C','K':'FFC7CE'}
wb = Workbook(); thin = Side(style='thin', color='BBBBBB'); bd = Border(left=thin,right=thin,top=thin,bottom=thin)
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='305496')
def dau(ws, cot, rong):
    ws.append(cot)
    for i, c in enumerate(ws[1], 1):
        c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical='center'); c.border = bd
        ws.column_dimensions[c.column_letter].width = rong[i-1]
    ws.freeze_panes = 'B2'
# 1. Ma trận
ws = wb.active; ws.title = 'Ma_tran'
dau(ws, ['Lá số (tên như trong công cụ)','Giới','Nguồn','Trang']+[t for _,t in LV]+['Giai đoạn đại vận theo sách'], [34,6,16,18]+[12]*7+[60])
for k, v in G.items():
    ws.append([k, v.get('gioi'), ' + '.join(v['nguon']), v['trang']]+[NH.get(v[c][0],'') if c in v else '' for c,_ in LV]+[v['DV'][0] if 'DV' in v else ''])
    r = ws.max_row
    for j,(c,_) in enumerate(LV):
        cell = ws.cell(r, 5+j)
        if c in v: cell.fill = PatternFill('solid', fgColor=MAU[v[c][0]])
    for cell in ws[r]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd
# 2. Đáp án chi tiết
ws = wb.create_sheet('Dap_an_chi_tiet')
dau(ws, ['Lá số','Giới','Lĩnh vực','Nhãn đề xuất','Câu sách (OCR, rút gọn)','Trang','Điều kiện tác giả nêu (+ tăng / − giảm)','Anh duyệt','Nhãn sửa / ghi chú'], [34,6,20,9,70,10,60,10,30])
dv = DataValidation(type='list', formula1='"Giữ,Sửa,Bỏ"', allow_blank=True); ws.add_data_validation(dv)
for k, v in G.items():
    for c, t in LV + [('DV','Giai đoạn đại vận'), ('HT','Hình thức / nghề')]:
        if c not in v: continue
        x = v[c]
        if c in ('DV','HT'): row = [k, v.get('gioi'), t, '', x[0], x[1], '', '', '']
        else: row = [k, v.get('gioi'), t, NH[x[0]], x[1], x[2], x[3], '', '']
        ws.append(row); r = ws.max_row; dv.add(ws.cell(r, 8))
        if c not in ('DV','HT'): ws.cell(r,4).fill = PatternFill('solid', fgColor=MAU[x[0]])
        for cell in ws[r]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd
        if c + '_P2' in v:
            y = v[c+'_P2']
            ws.append([k, v.get('gioi'), t+' (lời P2, đã thay bằng P3)', NH[y[0]], y[1], y[2], y[3], '', ''])
            for cell in ws[ws.max_row]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd; cell.font = Font(italic=True, color='777777')
    if v.get('ghi_chu'):
        ws.append([k, v.get('gioi'), 'GHI CHÚ', '', v['ghi_chu'], '', '', '', '']); 
        for cell in ws[ws.max_row]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd; cell.font = Font(bold=True, color='C00000')
ws.auto_filter.ref = ws.dimensions
# 3. Tổng hợp
ws = wb.create_sheet('Tong_hop', 0)
dau(ws, ['Lĩnh vực','Số lá sách có nhận định','Tốt','Vừa','Kém','Nhận xét'], [26,14,8,8,8,80])
NX = {'HN':'Sách chủ yếu nói khi hôn nhân XẤU → khi đo phải tính cả lá sách im lặng, tránh thiên lệch.',
      'LT':'Tương tự: 17/20 nhận định là xấu (khắc cha mẹ, anh em bất hoà).',
      'CC':'Rất ít dữ liệu (8 lá) → chỉ đủ rút quy tắc định tính, không đủ đo tỷ lệ.',
      'SK':'Phần lớn là "chú ý bệnh theo hành yếu/vượng"; nội dung chỉ tham khảo, không phải chẩn đoán.',
      'TL':'Đủ dữ liệu nhất (49 lá) → lĩnh vực làm trước ở giai đoạn B.',
      'SN':'Đủ dữ liệu (45 lá) → làm cùng tài lộc.'}
for c, t in LV:
    n = collections.Counter(v[c][0] for v in G.values() if c in v)
    ws.append([t, sum(n.values()), n['T'], n['V'], n['K'], NX.get(c,'')])
ws.append(['Giai đoạn đại vận', sum(1 for v in G.values() if 'DV' in v), '', '', '', 'Tuổi/vận tốt – xấu theo lời sách, dùng để đo phần "giai đoạn".'])
ws.append([]); ws.append(['Tổng số lá có lời luận', len(G)])
for row in ws.iter_rows(min_row=2):
    for cell in row: cell.alignment = Alignment(wrap_text=True, vertical='top')
# 4. Mâu thuẫn
ws = wb.create_sheet('Mau_thuan')
dau(ws, ['#','Chỗ','Hiện đang ghi','Sách ghi (đọc lại)','Hệ quả','Đề xuất (chờ anh duyệt)'], [4,26,40,55,45,40])
MT = [
 (1,'Bộ mẫu tài lộc 15 lá (công cụ) – P2-420 VD28','"có thể tích luỹ giàu có" → Giàu','Câu này là của VD27 (tr.420). VD28 (nữ) sách không nhận định mức giàu.','Số đo tài lộc trước đây (8/15, 12/15) sai','ĐÃ DUYỆT: thay bằng VD27 (Giàu) – áp vào workbook 6.5'),
 (2,'Bộ mẫu tài lộc – P2-423 VD32','"cuối đời có thể giàu nhất một vùng" → Giàu','Câu này là của VD31 (tr.423). VD32 (nữ) sách không nhận định mức giàu.','như trên','Bỏ VD32; thêm VD31 (Tốt)'),
 (3,'Bộ mẫu tài lộc – P2-426 VD35','"suốt đời tiền ít, khó khăn" → Thấp','Câu này là của VD34 (tr.426). VD35 sách nói 33–41 tuổi "10 năm huy hoàng nhất, tài vận hanh thông" → Tốt.','như trên','Sửa VD35 = Tốt; thêm VD34 (Kém)'),
 (4,'Bộ mẫu tài lộc – P1-VD9','"phú chân chính" → Giàu','Câu này là của P1-VD8 (tr.191). P1-VD9 sách chỉ nói quan nhỏ, nhiều tai hoạ, không nhận định mức giàu.','như trên','Bỏ P1-VD9; thêm P1-VD8 (Tốt)'),
 (5,'Bộ mẫu tài lộc – P2-422 VD30','"tiểu phú quý" → Vừa','Câu này là của VD29 (tr.421). VD30 sách nói "cơm áo không thiếu… của cải không nhiều" → vẫn Vừa.','Nhãn đúng, câu trích sai','Sửa câu trích'),
 (6,'Cùng một lá P2-428 VD37 = P3 tr.614','P2 tr.428: dụng hoả, hỉ mộc','P3 tr.614: dụng mộc, hỉ hoả','Sách tự mâu thuẫn','Ghi nhận; không đổi quy tắc'),
]
for m in MT:
    ws.append(list(m))
    for cell in ws[ws.max_row]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd
wb.save('/home/claude/battu/tai_lieu/GIAI_DOAN_A_DAP_AN_TVH.xlsx'); print('ok')
# ---- bổ sung bước 2–4 ----
import sys
sys.path.insert(0, '/home/claude/battu/tai_lieu/giai_doan_A')
from openpyxl import load_workbook
from dap_an_chuong import VD
wb = load_workbook('/home/claude/battu/tai_lieu/GIAI_DOAN_A_DAP_AN_TVH.xlsx')
TENLV = {'TC':'Tổng thể','SN':'Quan vận','TL':'Tài lộc','HN':'Hôn nhân','CC':'Con cái','LT':'Cha mẹ – anh em – tổ nghiệp','SK':'Sức khỏe – tai nạn','CD':'Học hành – danh tiếng','TH':'Tù tội','TU':'Tu hành'}
NH2 = dict(NH, **{'-':'(trung tính)'})
ws = wb.create_sheet('Vi_du_chuong_12_21')
dau(ws, ['Mã','Giới','Trang','Tứ trụ (năm, tháng, ngày, giờ – ? = OCR mờ)','Lĩnh vực','Nhãn','Câu sách / điều kiện tác giả nêu','Sự việc thực tế','Anh duyệt','Ghi chú'], [10,6,7,34,22,9,70,40,10,25])
dv2 = DataValidation(type='list', formula1='"Giữ,Sửa,Bỏ"', allow_blank=True); ws.add_data_validation(dv2)
for (ma, gioi, tr, tt, lv, nh, cau, sv) in VD:
    ws.append([ma, gioi, tr, tt, TENLV[lv], NH2.get(nh, nh), cau, sv, '', ''])
    r = ws.max_row; dv2.add(ws.cell(r, 9))
    if nh in MAU: ws.cell(r, 6).fill = PatternFill('solid', fgColor=MAU[nh])
    for cell in ws[r]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd
ws.auto_filter.ref = ws.dimensions
# Đo hiện tại
R = json.load(open('/home/claude/battu/tai_lieu/giai_doan_A/do_hien_tai.json'))
ws = wb.create_sheet('Do_phan_mem_hien_tai', 1)
dau(ws, ['Lĩnh vực','Số câu so','Khớp đúng mức','Tỷ lệ','Lệch ≤ 1 bậc','Sách Tốt → công cụ','Sách Vừa → công cụ','Sách Kém → công cụ','Nhận xét'], [28,9,10,8,10,22,22,22,70])
NX2 = {'TL':'Công cụ dồn gần hết về mức giữa: người sách nói giàu bị xếp "trung bình/khá giả"; người sách nói nghèo cũng vậy.',
       'SN':'Công cụ lạc quan: hễ quan sát là dụng/hỷ thì "THUẬN", bỏ qua quan bị hợp, bị thương quan, sát không chế…',
       'HN':'Công cụ bỏ sót phần lớn hôn nhân xấu (thương quan, tỉ kiếp trùng, kình dương, nhật trụ tự hình/âm dương sai lệch): dấu hiệu có, nhưng cách cộng làm loãng.',
       'SK':'Công cụ gần như lá nào cũng "bất lợi" → không phân biệt được.',
       'TC':'Ô "Cát – hung" hầu như luôn "trung bình".',
       'LT':'Khá hơn: các câu khắc cha mẹ đã có; còn sót ví dụ tổ nghiệp hưng thịnh và trường hợp "có khắc nhưng được cứu".'}
for lv in ['TL','SN','HN','LT','CC','SK','TC','CD','TH']:
    rr = [r for r in R if r['lv'] == lv]
    if not rr: continue
    n = len(rr); a = sum(r['sach'] == r['cong_cu'] for r in rr)
    b = sum(r['cong_cu'] is not None and abs('KVT'.index(r['sach']) - 'KVT'.index(r['cong_cu'])) <= 1 for r in rr)
    def phan(s):
        c = collections.Counter(r['cong_cu'] for r in rr if r['sach'] == s); return ' '.join('%s %d' % (NH.get(k, k), v) for k, v in c.items())
    ws.append([TENLV[lv], n, a, '%d%%' % (100*a//n), b, phan('T'), phan('V'), phan('K'), NX2.get(lv, '')])
    for cell in ws[ws.max_row]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd
ws2 = wb.create_sheet('Do_chi_tiet')
dau(ws2, ['Lá','Nguồn','Lĩnh vực','Sách','Công cụ','Công cụ ghi','Câu sách','Trang'], [30,16,22,8,8,30,70,10])
for r in R:
    ws2.append([r['la'], r['nguon'], TENLV[r['lv']], NH.get(r['sach']), NH.get(r['cong_cu'], r['cong_cu']), r['cong_cu_raw'], r['cau'], r['trang']])
    if r['sach'] != r['cong_cu']: ws2.cell(ws2.max_row, 5).fill = PatternFill('solid', fgColor='FFC7CE')
ws2.auto_filter.ref = ws2.dimensions
# Câu sách ↔ ví dụ
C = json.load(open('/home/claude/battu/tai_lieu/giai_doan_A/cau_sach_vi_du.json'))
ws3 = wb.create_sheet('Cau_sach_vi_du')
dau(ws3, ['Câu sách (tác giả nói ví dụ "ứng câu")','Số ví dụ','Lĩnh vực','Nhãn các ví dụ','Ví dụ (trang)','Vị trí gần nhất trong công cụ (so chữ tự động)','Độ giống'], [60,8,14,16,40,30,8])
for o in C:
    ws3.append([o['cau'], o['so_vd'], ', '.join(TENLV.get(x, x) for x in o['linh_vuc']), ', '.join('%s %d' % (NH2.get(k,k), v) for k, v in o['nhan'].items()), '; '.join(o['vi_du']), o['trong_cong_cu'] or 'chưa thấy bằng so chữ – cần rà tay', o['do_giong']])
    for cell in ws3[ws3.max_row]: cell.alignment = Alignment(wrap_text=True, vertical='top'); cell.border = bd
# Tổng hợp: thêm dòng
ws = wb['Tong_hop']
ws.append([]); ws.append(['BỔ SUNG BƯỚC 2: ví dụ chương 12–21 (P2 tr.432–572)', len({v[0] for v in VD}), '', '', '', '%d dòng nhận định; %d có tứ trụ đầy đủ' % (len(VD), sum(1 for v in VD if '?' not in v[3]))])
c = collections.defaultdict(collections.Counter)
for v in VD: c[v[4]][v[5]] += 1
for lv, x in c.items(): ws.append(['  ' + TENLV[lv], sum(x.values()), x['T'], x['V'], x['K'], ('trung tính %d' % x['-']) if x['-'] else ''])
wb.save('/home/claude/battu/tai_lieu/GIAI_DOAN_A_DAP_AN_TVH.xlsx'); print('ok2')
