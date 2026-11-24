# -*- coding: utf-8 -*-
"""Bản 6.11 = 6.10 + PA2 mục I mới (tools/step27_pa2_moi.py).
Dùng: python3 tools/step27.py <6.9.xlsx> <6.10.xlsx> <cổng>   (biến môi trường CAN_DAN=0 để không đòi can dẫn; DOI_MUC=0 chỉ ghi, không đổi mức)"""
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
for _nm, _rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', _wb):
    _t = re.search(r'Id="%s"[^>]*Target="[^"]*sheet(\d+)\.xml"' % _rid, _rel)
    if _t: xpatch.SHEETS.setdefault(_nm, int(_t.group(1)))
import step27_pa2_moi as P2
P2.pa2_moi(p, 'Bản 6.11 (24/11/2026): mục I PA2 – cách A, Thổ theo Thủy, có xét giả ở E7, E12; E13, E14 không hợp hóa')
tmp = '/tmp/s27_tmp%s.xlsx' % PORT; p.save(tmp)
print('vá xong', len(p.log), 'ô', flush=True)
if os.environ.get('CHI_VA'):
    shutil.copy(tmp, DST); sys.exit(0)

# --- tính lại bằng LibreOffice (chỉ tắt đúng PID)
import uno
from com.sun.star.beans import PropertyValue
pr = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/s27prof{PORT}',
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
