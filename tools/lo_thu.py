# -*- coding: utf-8 -*-
"""Đo ngữ nghĩa LibreOffice: tính danh sách công thức trên một bảng tính trống.
Dùng: python3 lo_thu.py <tệp_vào.json: [[ô_hằng_dict], [công_thức...]]> <tệp_ra.json> <port>
Chỉ tắt soffice theo PID của chính nó (không pkill)."""
import sys, json, time, subprocess, uno
from com.sun.star.beans import PropertyValue
IN, OUT, PORT = sys.argv[1], sys.argv[2], int(sys.argv[3])
consts, formulas = json.load(open(IN))
desk_ref = []
p = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/lothu{PORT}',
                      f'--accept=socket,host=localhost,port={PORT};urp;'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    local = uno.getComponentContext(); res = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(90):
        try: ctx = res.resolve(f'uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext'); break
        except Exception: time.sleep(1)
    desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx); desk_ref.append(desk)
    def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
    doc = desk.loadComponentFromURL('private:factory/scalc', '_blank', 0, (pv('Hidden', True),))
    sh = doc.Sheets.getByIndex(0)
    for ref, v in consts.items():
        c = sh.getCellRangeByName(ref)
        if isinstance(v, str): c.setString(v)
        else: c.setValue(v)
    for i, f in enumerate(formulas):
        sh.getCellByPosition(10, i).setFormula(f)
    doc.calculateAll()
    data = sh.getCellRangeByPosition(10, 0, 10, len(formulas) - 1).getDataArray()
    out = []
    for i, f in enumerate(formulas):
        cell = sh.getCellByPosition(10, i)
        out.append([f, data[i][0], cell.getString(), int(cell.getError())])
    json.dump(out, open(OUT, 'w'), ensure_ascii=False, indent=0)
    doc.close(True)
    try: desk.terminate()
    except Exception: pass
finally:
    for d_ in desk_ref:          # tắt soffice.bin qua UNO (tiến trình con không nằm dưới PID của p)
        try: d_.terminate()
        except Exception: pass
    time.sleep(1)
    if p.poll() is None: p.kill()
