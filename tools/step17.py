# -*- coding: utf-8 -*-
"""Bản 6.4 = 6.3 + G142 (phương án A, duyệt 14/11/2026): ghi chú "dụng thần theo cách bổ hành khuyết".
Chỉ thêm ô ghi chú – KHÔNG đổi dụng thần, vượng suy hay kết luận nào.
  * DUNG_HY_KY dòng 197–199 (mới): D198 = nhóm khuyết được gợi ý (−1 = không), C198 = câu ghi chú
  * KET_LUAN C52 (dòng "Tham khảo TTT · Dụng thần thái quá"): nối thêm câu C198
  * MAU_THUAN dòng 168, TRANG_CHU B3: ghi bản 6.4
Dùng: python3 tools/step17.py <workbook_6.3.xlsx> <workbook_6.4.xlsx> <cổng>
Theo cách của 06_MA_THAM_CHIEU/build_excel/step16.py: vá XML → LibreOffice tính lại, lưu (tắt đúng PID) → fix_dt, fix_links, kẹp chiều cao dòng."""
import sys, os, re, time, shutil, subprocess, zipfile as _z
B = '/home/claude/HO_SO_PM/HO_SO_PM/06_MA_THAM_CHIEU/build_excel'
sys.path.insert(0, B)
import xpatch
from xpatch import Pkg
xpatch.SHEETS.update({'KET_LUAN': 4, 'TRANG_CHU': 1, 'MAU_THUAN': 28, 'DUNG_HY_KY': 9})
import fix_dt, fix_links
SRC, DST, PORT = sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 2617
why = 'G142 (duyệt 14/11/2026 – phương án A): ghi chú dụng thần theo cách bổ hành khuyết; không đổi kết luận'
p = Pkg(SRC)

# --- DUNG_HY_KY: khối L (dòng 197–199)
p.set_text('DUNG_HY_KY', 'B197', 'L. DỤNG THẦN THEO CÁCH BỔ HÀNH KHUYẾT – chỉ ghi chú, không đổi dụng thần (G142)', style_from='B150', why=why, must_be_empty=True)
p.set_text('DUNG_HY_KY', 'B198', 'Nhóm khuyết trong phe dụng thần (D: độ lệch nhóm, −1 = không có) – ghi chú', style_from='B171', why=why, must_be_empty=True)
D198 = ('IF(OR($C$160<>"",$C$155=2,$C$156<0),-1,'
        'IF($C$155=1,IF(OR($C$156=1,$C$156=2,$C$156=3),'
        'IF(AND($C$156<>3,$C$21=0),3,IF(AND($C$156<>1,$C$19=0),1,IF(AND($C$156<>2,$C$20=0),2,-1))),-1),'
        'IF(OR($C$156=0,$C$156=4),IF(AND($C$156<>4,$C$22=0),4,IF(AND($C$156<>0,$C$18=0),0,-1)),-1)))')
p.new_formula('DUNG_HY_KY', 'D198', D198, style_from='C171', why=why)
C198 = ('IF($D$198<0,"","Dụng thần theo cách bổ hành khuyết: "&INDEX(HANH_TEN,MOD($C$5+$D$198,5)+1)&" ("&INDEX(NHOM_TEN,$D$198+1)'
        '&") – tứ trụ không có chữ nào, lộ hay tàng, thuộc nhóm này. Khi giải ví dụ, sách nhiều lần lấy nhóm đang khuyết làm dụng thần'
        ' (dụng thần khuyết thì chờ đại vận, lưu niên đưa tới – P1 tr.97). Chỉ là ghi chú để cân nhắc, không thay dụng thần "'
        '&INDEX(HANH_TEN,$C$157+1)&" ("&INDEX(NHOM_TEN,$C$156+1)&") đã chọn theo bảng P1 tr.108, P2 tr.380-384.")')
p.new_formula('DUNG_HY_KY', 'C198', C198, text=True, style_from='C171', why=why)
p.set_text('DUNG_HY_KY', 'B199', 'Quy tắc G142: chính cách, không trung hòa. Thân vượng: xét Quan sát → Thực thương → Tài; thân nhược: Ấn → Tỉ kiếp; '
           'lấy nhóm đầu tiên có 0 chữ (C18:C22) và khác nhóm dụng thần đã chọn. Kiểm trên 73 lá ví dụ có dụng thần: nếu lấy nhóm này '
           'làm dụng thì khớp sách 45/73 (hiện 40/73), nhưng làm kém đi trên Trích Thiên Tủy – nên chỉ ghi chú.', style_from='B171', why=why, must_be_empty=True)

# --- KET_LUAN C52: nối câu ghi chú
old = p.formula('KET_LUAN', 'C52')
assert old == 'DUNG_HY_KY!$C$145&IF(AND(DUNG_HY_KY!$C$145<>"",DUNG_HY_KY!$C$151<>""),CHAR(10),"")&DUNG_HY_KY!$C$151', old
p.set_formula('KET_LUAN', 'C52', old + '&IF(AND(DUNG_HY_KY!$C$145&DUNG_HY_KY!$C$151<>"",DUNG_HY_KY!$C$198<>""),CHAR(10),"")&DUNG_HY_KY!$C$198',
              text=True, why=why)
assert p.text_of('KET_LUAN', 'B52') == 'Tham khảo TTT · Dụng thần thái quá'
p.set_text('KET_LUAN', 'B52', 'Tham khảo TTT · Dụng thần thái quá · Bổ hành khuyết', why=why)

# --- TRANG_CHU B3, MAU_THUAN 168
t = p.text_of('TRANG_CHU', 'B3'); i = t.find(' · Căn cứ'); assert i > 0
p.set_text('TRANG_CHU', 'B3', 'Bản 6.4 – 14/11/2026 (6.3 + G142: ghi chú dụng thần theo cách bổ hành khuyết, không đổi kết luận; '
           '6.3 = G141 đổi tên chế độ PA2; 6.2 = G140; 6.1 = G130, G132, G133; SU_KIEN; DOC_NHANH)' + t[i:], why='Bản 6.4')
k = 168
assert p.text_of('MAU_THUAN', f'D{k}') is None
p.set_number('MAU_THUAN', f'B{k}', 165, style_from='B167', why=why)
p.set_text('MAU_THUAN', f'C{k}', 'Thêm', style_from='C167', why=why)
for c, v in zip('DEFG', ('G142 (duyệt 14/11/2026 – phương án A). Ghi chú "dụng thần theo cách bổ hành khuyết": ở chính cách, nếu trong phe dụng thần '
                          '(vượng: quan sát, thực thương, tài; nhược: ấn, tỉ kiếp) có nhóm không có chữ nào (lộ hay tàng) thì nêu nhóm đó. Căn cứ: '
                          'phân tích 73 lá ví dụ có dụng thần – sách chọn nhóm khuyết ở 9 lá, công cụ 2 lá; P1 tr.97. Không đổi dụng thần, hỷ, kỵ, '
                          'vượng suy hay bất kỳ kết luận nào.',
                          'DUNG_HY_KY dòng 197–199; KET_LUAN C52', 'Giúp người luận giải cân nhắc cách chọn của tác giả khi giải ví dụ', 'DUNG_HY_KY')):
    p.set_text('MAU_THUAN', f'{c}{k}', v, style_from=f'{c}167', why=why)

tmp = '/tmp/s17_tmp.xlsx'; p.save(tmp)

# --- tính lại bằng LibreOffice (chỉ tắt đúng PID)
import uno
from com.sun.star.beans import PropertyValue
pr = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/s17prof{PORT}',
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
