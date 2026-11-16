# -*- coding: utf-8 -*-
"""Bản 6.7 = 6.6 + các sửa đã duyệt ngày 16/11/2026:
  1. Bỏ quy tắc vượng suy 6.6 (theo PP2 khi phe mình < 36% hoặc ≥ 52%) – trở lại cách 6.5: PP1 đủ 5 bước (VS-2), G11 cũ.
     Đo trên 73 lá Thiệu Vĩ Hoa: 6.6 đúng 63, 6.5 đúng 64; bỏ 5 lá sách tự mâu thuẫn: 6.6 đúng 61/68, 6.5 đúng 64/68.
  2. Đổi mã trùng: "DT-3" của 6.6 (dụng thần hiển thị theo phương pháp khuyến nghị) → DT-4; "DT-1" của 6.6 (ứng viên thay thế) → DT-5.
Dùng: python3 tools/step19.py <workbook_6.6.xlsx> <workbook_6.7.xlsx> <cổng>"""
import sys, os, re, time, json, shutil, subprocess, zipfile as _z
B = '/home/claude/HO_SO_PM/HO_SO_PM/06_MA_THAM_CHIEU/build_excel'
sys.path.insert(0, B)
import xpatch
from xpatch import Pkg
xpatch.SHEETS.update({'KET_LUAN': 4, 'TRANG_CHU': 1, 'MAU_THUAN': 28, 'DUNG_HY_KY': 9, 'DOI_CHIEU': 8, 'PP1_DINH_TINH': 10,
                      'DAI_VAN': 15, 'TAI_LOC': 20, 'KIEM_THU': 27, 'MENH_CUC': 19, 'LUAN_GIAI': 5, 'KHONG_RO_GIO': 7})
import fix_dt, fix_links
SRC, DST, PORT = sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 2618
p = Pkg(SRC)
_wb = p.data['xl/workbook.xml'].decode(); _rel = p.data['xl/_rels/workbook.xml.rels'].decode()
for _nm, _rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', _wb):   # đủ tên sheet → số file
    _t = re.search(r'Id="%s"[^>]*Target="[^"]*sheet(\d+)\.xml"' % _rid, _rel)
    if _t: xpatch.SHEETS.setdefault(_nm, int(_t.group(1)))
CT = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'step19_cong_thuc_6.5.json'), encoding='utf-8'))
why = 'Bản 6.7 (duyệt 16/11/2026): bỏ quy tắc vượng suy 6.6, trở lại cách 6.5'

# 1. Vượng suy: trở lại 6.5
assert 'PP2_DINH_LUONG!$C$101>=$D$30' in p.formula('DOI_CHIEU', 'C33')
p.set_formula('DOI_CHIEU', 'C33', CT['C33'], why=why)
p.set_formula('DOI_CHIEU', 'C36', CT['C36'], text=True, why=why)
p.set_number('DOI_CHIEU', 'C30', 0.3, why=why)
p.set_number('DOI_CHIEU', 'D30', 0.6, why=why)

# 2. Trang chủ + bảng mâu thuẫn
t = p.text_of('TRANG_CHU', 'B3'); i = t.find(' · Căn cứ'); assert i > 0
p.set_text('TRANG_CHU', 'B3', 'Bản 6.7 – 16/11/2026 (6.6, bỏ quy tắc vượng suy theo ngưỡng PP2 – trở lại cách 6.5: PP1 đủ 5 bước; '
           'dụng thần hiển thị theo phương pháp khuyến nghị DT-4; ứng viên dụng thần thay thế DT-5; năm có hạn nặng không xếp “Rất tốt”; '
           'đối chiếu nguồn con cái; gợi ý phương pháp từ sự việc đã qua)' + t[i:], why='Bản 6.7')
MT = [
    ('Bỏ', 'Quy tắc vượng suy 6.6 (khi PP1 khác PP2 mà phe mình PP2 < 36% hoặc ≥ 52% thì theo PP2) bị bỏ, trở lại cách 6.5. Lý do: ngưỡng dò '
           'từ 431 lá, trong đó 366 lá sách khác (khác trường phái); chuyển cả lá được lệnh (trái P1 tr.70) và lật lại kết luận bước 5. '
           'Đo trên 73 lá Thiệu Vĩ Hoa: 6.6 đúng 63, 6.5 đúng 64; bỏ 5 lá sách tự mâu thuẫn: 6.6 đúng 61/68, 6.5 đúng 64/68; '
           '9 lá bị chuyển chỉ đúng 4; kiểm chéo bỏ-một-lá 62/73.', 'DOI_CHIEU C30, D30, C33, C36', 'Kết luận vượng suy theo PP1 đủ 5 bước', 'DOI_CHIEU'),
    ('Đổi mã', 'Mã của bản 6.6 trùng mã đã dùng ở 6.5: "DT-3" (ô dụng/hỷ/kỵ và thần sát theo phương pháp khuyến nghị) đổi thành DT-4; '
               '"DT-1" (ứng viên dụng thần thay thế) đổi thành DT-5. DT-3 vẫn là quy tắc điều hầu của 6.5.', 'KET_LUAN C36:C39; DUNG_HY_KY dòng 221–227',
     'Tránh nhầm mã', 'KET_LUAN'),
]
k0 = 181
assert p.text_of('MAU_THUAN', 'D%d' % k0) is None and p.text_of('MAU_THUAN', 'D180')
so0 = int(float(p.text_of('MAU_THUAN', 'B180') or 177))
for j, (loai, d, e, f, g) in enumerate(MT):
    k = k0 + j
    p.set_number('MAU_THUAN', 'B%d' % k, so0 + 1 + j, style_from='B180', why='Bản 6.7')
    p.set_text('MAU_THUAN', 'C%d' % k, loai, style_from='C180', why='Bản 6.7')
    for col, v in zip('DEFG', (d, e, f, g)):
        p.set_text('MAU_THUAN', '%s%d' % (col, k), v, style_from='%s180' % col, why='Bản 6.7')

if os.environ.get('THO_YEU'):
    exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'step19_tho_yeu.py'), encoding='utf-8').read())

tmp = '/tmp/s19_tmp.xlsx'; p.save(tmp)
print('vá xong', len(p.log), 'ô', flush=True)
if os.environ.get('CHI_VA'):
    shutil.copy(tmp, DST); sys.exit(0)

# --- tính lại bằng LibreOffice (chỉ tắt đúng PID)
import uno
from com.sun.star.beans import PropertyValue
pr = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/s19prof{PORT}',
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
