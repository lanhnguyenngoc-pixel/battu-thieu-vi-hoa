# -*- coding: utf-8 -*-
"""Thí nghiệm: giá trị TRUNG_GIAN cột Giờ khi "Không rõ giờ" sau 1, 2 lần calculateAll và sau CalculateHard.
Dùng: python3 lo_thu_krg.py <workbook> <port>   (tắt soffice theo PID)"""
import sys, time, subprocess, os, uno
from com.sun.star.beans import PropertyValue
SRC, PORT = sys.argv[1], int(sys.argv[2])
desk_ref = []
p = subprocess.Popen(['soffice', '--headless', '--norestore', '--nologo', f'-env:UserInstallation=file:///tmp/lokrg{PORT}',
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
    S = doc.Sheets; NL = S.getByName('NHAP_LIEU'); TG = S.getByName('TRUNG_GIAN')
    for i, col in enumerate('CDEF'):
        NL.getCellRangeByName(f'{col}7').setString(['Ất', 'Ất', 'Kỷ', 'Nhâm'][i]); NL.getCellRangeByName(f'{col}8').setString(['Tị', 'Dậu', 'Mùi', 'Thân'][i])
    NL.getCellRangeByName('C12').setString('Không'); NL.getCellRangeByName('C13').setString('Tự nhập')
    NL.getCellRangeByName('C9').setString('Không rõ giờ')
    def doc_o(): return [TG.getCellRangeByName(r).getValue() for r in ('F6', 'F15', 'F17', 'F20', 'F28')]
    doc.calculateAll(); print('calculateAll 1:', doc_o())
    doc.calculateAll(); print('calculateAll 2:', doc_o())
    disp = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    disp.executeDispatch(doc.getCurrentController().getFrame(), '.uno:CalculateHard', '', 0, ())
    print('CalculateHard  :', doc_o())
    # xóa Bảng dữ liệu rồi tính
    KR = S.getByName('KHONG_RO_GIO')
    for t in [118 + 14 * i for i in range(12)]: KR.getCellRangeByName(f'C{t}:D{t+11}').clearContents(1 | 2 | 4 | 16)
    doc.calculateAll(); print('không bảng     :', doc_o())
    doc.close(True)
    try: desk.terminate()
    except Exception: pass
finally:
    for d_ in desk_ref:          # tắt soffice.bin qua UNO (tiến trình con không nằm dưới PID của p)
        try: d_.terminate()
        except Exception: pass
    time.sleep(2)
    if p.poll() is None: p.kill()
