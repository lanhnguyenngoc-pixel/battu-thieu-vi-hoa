# -*- coding: utf-8 -*-
"""Bản 6.8 = 6.7 + PP1 "được sinh" theo cả bốn trụ (phương án D, người dùng duyệt 16/11/2026, chạy thứ 7 21/11/2026):
  ấn ở các can chi khác trong tứ trụ (P1 tr.70) – ấn lộ ở can năm, tháng, giờ; ấn tàng bản, trung, dư khí ở cả 4 chi
  (kể cả chi tháng khi được lệnh, kể cả chi đã làm gốc); chi hóa thành ấn ở cả 4 chi. Giữ: chi hợp mà không hóa không tính (P1 tr.62).
  Thay giả định G78. Theo yêu cầu người dùng: KHÔNG thêm dòng MAU_THUAN.
  Đo trước khi duyệt: 73 lá Thiệu Vĩ Hoa 64 → 63 (P1-VD3 Nhược → Vượng); 50 lá ngẫu nhiên: không đổi kết luận, 3 lá Rất vượng → Quá vượng.
Dùng: python3 tools/step20.py <workbook_6.7.xlsx> <workbook_6.8.xlsx> <cổng>"""
import sys, os, re, time, json, shutil, subprocess, zipfile as _z
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
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
why = 'Bản 6.8 (duyệt 16/11/2026): PP1 được sinh – ấn ở các can chi khác trong tứ trụ (P1 tr.70), cả bốn trụ'
S = 'PP1_DINH_TINH'; AN = 'MOD(TRUNG_GIAN!$E$7+4,5)'
assert 'G78' in p.text_of(S, 'B52') and p.formula(S, 'D53') == '0'
for c in 'CDEF':
    p.set_formula(S, '%s50' % c, 'IF(%s27=%s,1,0)' % (c, AN), why=why)
    for r, h in ((52, 21), (53, 22), (54, 23)):
        p.set_formula(S, '%s%d' % (c, r), 'IF(AND(%s27<0,%s28=0,TRUNG_GIAN!%s$%d=%s),1,0)' % (c, c, c, h, AN), why=why)
p.set_text(S, 'B50', '   Chi năm, tháng, ngày, giờ hóa thành ấn', why=why)
p.set_text(S, 'B52', '   Ấn tàng ở chi năm, tháng, ngày, giờ (bản khí) – ấn ở các can chi khác trong tứ trụ (P1 tr.70)', why=why)
p.set_text(S, 'B53', '   Ấn tàng ở chi năm, tháng, ngày, giờ (trung khí)', why=why)
p.set_text(S, 'B54', '   Ấn tàng ở chi năm, tháng, ngày, giờ (dư khí)', why=why)
o = p.formula(S, 'C106')
n = o.replace('cũng không tàng trong chi năm, ngày, giờ (chi tháng chỉ xét bản khí khi không được lệnh – G78; P1 tr.70-72).',
              'cũng không tàng trong chi năm, tháng, ngày, giờ (ấn ở các can chi khác trong tứ trụ – P1 tr.70-72).')
assert n != o; p.set_formula(S, 'C106', n, text=True, why=why)
t = p.text_of('TRANG_CHU', 'B3'); i = t.find(' · Căn cứ'); assert i > 0
p.set_text('TRANG_CHU', 'B3', 'Bản 6.9 – 22/11/2026 (6.7, PP1 “được sinh” xét ấn ở các can chi khác trong cả bốn trụ – P1 tr.70; '
           'vượng suy PP1 đủ 5 bước; thọ – yểu theo P1 tr.172–173 kèm luận thêm; đại vận 6 mức, can chi tác dụng cả 10 năm; Mộc mùa hạ cần Thủy điều hầu; '
           'Ngũ hành thiếu mục I; sao con cái mệnh nữ theo P2 tr.504; lưu niên: điều 6, 7, 8 P2 tr.288–289, xu thế năm có hình và hợp mất kỵ, gốc Thường giữ Thường)' + t[i:], why='Bản 6.8')

import step20_bo_sung as BS
BS.ngu_hanh_thieu(p, 'Bản 6.8 (duyệt 18/11/2026): Ngũ hành thiếu mục I')
BS.sao_con(p, 'Bản 6.8 (duyệt 17/11/2026): sao con cái nam/nữ theo P2 tr.504')
BS.dai_van(p, 'Bản 6.8 (duyệt 17/11/2026): đại vận 10 năm, mức 6 bậc (mục 2.4, 2.5)')
BS.dieu_hau(p, 'Bản 6.8 (duyệt 17/11/2026): điều hầu', moc_ha=os.environ.get('MOC_HA', '1') == '1')
BS.luu_nien(p, 'Bản 6.9 (giai đoạn 2): điều 6–8 P2 tr.288–289, xu thế năm có hình, hợp mất kỵ, gốc Thường giữ Thường')
BS.nguoi_ung(p, 'Bản 6.9 (giai đoạn 2, duyệt 16/11/2026): người ứng theo cung + sao lục thân (mục 2.2)')
BS.bo_sung_27(p, 'Bản 6.9 (giai đoạn 2, duyệt 17/11/2026): mục 2.7 điểm 1 (kình dương, P1 tr.117) và điểm 3 (P1 tr.108)')
import step21_con_cai as CC
CC.hop_hoa_luu_nien(p, 'Bản 6.9 (giai đoạn 2, yêu cầu 18/11/2026): mục 2.11 điểm C – điều 5 hợp hóa can lưu niên')
CC.con_cai(p, 'Bản 6.9 (giai đoạn 2, duyệt 17/11/2026): mục 2.8 – dấu hiệu con cái P2 tr.500–511 và câu tổng hợp')
CC.hop_cuc_luu_nien(p, 'Bản 6.9 (giai đoạn 2, 22/11/2026): mục 2.11 – chi lưu niên hợp/hội cục với tứ trụ, đại vận (P2 tr.281-283, 254)')
tmp = '/tmp/s20_tmp.xlsx'; p.save(tmp)
print('vá xong', len(p.log), 'ô', flush=True)
if os.environ.get('CHI_VA'):
    shutil.copy(tmp, DST); sys.exit(0)

# --- tính lại bằng LibreOffice (chỉ tắt đúng PID)
import uno
from com.sun.star.beans import PropertyValue
pr = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/s20prof{PORT}',
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
