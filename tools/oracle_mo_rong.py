# -*- coding: utf-8 -*-
"""ORACLE MỞ RỘNG (MT-04, MT-05 đã duyệt 13/11/2026) – tệp mới, KHÔNG sửa 05_KIEM_THU/scripts/oracle_dump.py.
Giống oracle_dump.py (cùng cách đặt ô nhập, cùng trạng thái cộng dồn) nhưng:
  * ghi thêm vùng Bảng dữ liệu KHONG_RO_GIO C118:D285 (288 ô kết quả + ô công thức xen giữa) – T1-KRG;
  * mỗi "biến thể" đặt thêm một số ô (sau phần nhập chuẩn, ở mọi lá) – T1b.
Dùng: python3 oracle_mo_rong.py <workbook.xlsx> <cases.json> <bien_the.json> <tên_biến_thể> <thư_mục_ra> <port>
Không bao giờ pkill soffice: chỉ tắt đúng PID do script mở."""
import sys, json, time, subprocess, os, re, zipfile, uno
from com.sun.star.beans import PropertyValue
SRC, CASES, BT, TEN, OUT, PORT = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], int(sys.argv[6])
os.makedirs(OUT, exist_ok=True)
bien = json.load(open(BT, encoding='utf-8'))[TEN]
z = zipfile.ZipFile(SRC); wbx = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx); FC = {}
for name, rid in sheets:
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1)
    x = z.read('xl/' + tgt.replace('xl/', '')).decode()
    FC[name] = [m.group(1) for m in re.finditer(r'<c r="([A-Z]+\d+)"(?:\s[^>]*?)?(?<!/)>(?:(?!</c>).)*?<f', x, re.S)]
# vùng Bảng dữ liệu: mọi ô C118:D285
KRG = ['%s%d' % (col, r) for r in range(118, 286) for col in 'CD']
def rc(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref); c = 0
    for ch in m.group(1): c = c * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, c - 1
desk_ref = []
p = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/oraclemr{PORT}',
                      f'--accept=socket,host=localhost,port={PORT};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    local = uno.getComponentContext(); res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(90):
        try: ctx = res.resolve(f'uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext'); break
        except Exception: time.sleep(1)
    desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx); desk_ref.append(desk)
    def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
    doc = None
    for _ in range(8):
        try: doc = desk.loadComponentFromURL('file://' + os.path.abspath(SRC), '_blank', 0, (pv('Hidden', True),))
        except Exception: doc = None
        if doc is not None: break
        time.sleep(5)
    S = doc.Sheets; NL = S.getByName('NHAP_LIEU')
    def dat(sh, ref, v):
        cell = S.getByName(sh).getCellRangeByName(ref)
        if v is None: cell.setString('')
        elif isinstance(v, str): cell.setString(v)
        else: cell.setValue(v)
    for c in json.load(open(CASES, encoding='utf-8')):
        if 'can' in c:
            for i, col in enumerate('CDEF'):
                NL.getCellRangeByName(f'{col}7').setString(c['can'][i]); NL.getCellRangeByName(f'{col}8').setString(c['chi'][i])
            NL.getCellRangeByName('C12').setString('Không'); NL.getCellRangeByName('C13').setString('Tự nhập')
            for col in 'CDEFG': NL.getCellRangeByName(f'{col}26').setString('')
        else:
            NL.getCellRangeByName('C12').setString('Tự động'); NL.getCellRangeByName('C13').setString('Tự nhập'); NL.getCellRangeByName('C9').setString('Biết giờ')
            for col, k in zip('CDEFG', ('ngay', 'thang', 'nam', 'gio', 'phut')): NL.getCellRangeByName(f'{col}26').setValue(c[k])
            NL.getCellRangeByName('C27').setValue(c.get('mui_gio', 7))
        NL.getCellRangeByName('C10').setString(c.get('gioi_tinh', 'Nam')); NL.getCellRangeByName('C11').setString('')
        # phần đặt thêm của biến thể: chung + riêng từng lá
        for sh, d in bien.get('chung', {}).items():
            for ref, v in d.items(): dat(sh, ref, v)
        for sh, d in bien.get('rieng', {}).get(c['id'], {}).items():
            for ref, v in d.items(): dat(sh, ref, v)
        if bien.get('xoa_bang') and not getattr(dat, 'da_xoa', False):
            # như 05_KIEM_THU/scripts/uno_run59.py: xóa 12 Bảng dữ liệu KHONG_RO_GIO trước khi tính
            KR = S.getByName('KHONG_RO_GIO')
            for t in [118 + 14 * i for i in range(12)]: KR.getCellRangeByName(f'C{t}:D{t+11}').clearContents(1 | 2 | 4 | 16)
            dat.da_xoa = True
        if bien.get('bang_vong'):
            # MT-13: Bảng dữ liệu KHONG_RO_GIO tính bằng VÒNG 12 GIỜ (không dùng MULTIPLE.OPERATIONS của LibreOffice,
            # vì LibreOffice để sót giá trị của giờ cuối bảng ở các ô phụ thuộc – xem tools/lo_thu_krg.py).
            KR = S.getByName('KHONG_RO_GIO')
            tops = [117 + 14 * i for i in range(12)]
            for t in tops: KR.getCellRangeByName(f'C{t+1}:D{t+12}').clearContents(1 | 2 | 4 | 16)
            goc = KR.getCellRangeByName('C6').getString()
            gio = [KR.getCellRangeByName(f'B{118+i}').getString() for i in range(12)]
            for t in tops:
                assert [KR.getCellRangeByName(f'B{t+1+i}').getString() for i in range(12)] == gio, 'cột giờ của bảng khác nhau'
            vals = {}
            for i, h in enumerate(gio):
                KR.getCellRangeByName('C6').setString(h); doc.calculateAll()
                for t in tops:
                    row = KR.getCellRangeByName(f'C{t}:D{t}').getDataArray()[0]
                    vals[(t + 1 + i, 'C')] = row[0]; vals[(t + 1 + i, 'D')] = row[1]
            KR.getCellRangeByName('C6').setString(goc)
            for (r, col), v in vals.items():
                cell = KR.getCellRangeByName(f'{col}{r}')
                if v == '': cell.setFormula('=""')      # giữ CHUỖI RỖNG (setString('') làm ô trống hẳn)
                elif isinstance(v, str): cell.setString(v)
                else: cell.setValue(v)
        doc.calculateAll(); outc = {}
        chi = bien.get('chi_sheets'); them = bien.get('them_o', {})
        for name, refs in FC.items():
            if chi is not None and name not in chi and name not in them: continue
            if chi is not None and name not in chi: refs = them[name]
            sh = S.getByName(name); cur = sh.createCursor(); cur.gotoEndOfUsedArea(False)
            R, C = cur.RangeAddress.EndRow, cur.RangeAddress.EndColumn
            data = sh.getCellRangeByPosition(0, 0, C, R).getDataArray(); d = {}
            lst = refs if (name != 'KHONG_RO_GIO' or bien.get('xoa_bang')) else sorted(set(refs) | set(KRG), key=lambda x: rc(x))
            for ref in lst:
                r, k = rc(ref)
                if r <= R and k <= C: d[ref] = data[r][k]
            outc[name] = d
        json.dump(outc, open(os.path.join(OUT, re.sub(r'[^\w\-]+', '_', c['id']) + '.json'), 'w'), ensure_ascii=False)
        print(TEN, c['id'], flush=True)
    doc.close(True)
    try: desk.terminate()
    except Exception: pass
finally:
    for d_ in desk_ref:          # tắt soffice.bin qua UNO (tiến trình con không nằm dưới PID của p)
        try: d_.terminate()
        except Exception: pass
    time.sleep(2)
    if p.poll() is None: p.kill()
