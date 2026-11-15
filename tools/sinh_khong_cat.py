# -*- coding: utf-8 -*-
"""Sinh export/_khong_cat_sinh.py: bản sao "không cắt" của các hàm công thức có X.cat(X.LEFT(x, N), '…')
(đề xuất hoàn thiện 15/11/2026, mục C2). Bộ cài (PyInstaller) không kèm mã nguồn .py của battu_core/sheets, nên
export/khoi_phuc_cat.py dùng tệp sinh sẵn này thay cho đọc mã nguồn lúc chạy.
Chạy lại mỗi khi dựng lại sheet (tools/dung_lai.sh gọi sẵn):  python -I tools/sinh_khong_cat.py"""
import ast, importlib, inspect, os, sys, textwrap
GOC = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, GOC)
from battu_core import sheets as S
from export.khoi_phuc_cat import _BoCat


def main():
    ra = ['# -*- coding: utf-8 -*-', '"""SINH TỰ ĐỘNG bởi tools/sinh_khong_cat.py – KHÔNG SỬA TAY.',
          'Bản sao "không cắt" của hàm công thức: X.cat(X.LEFT(x, N), \'…\') → x. Nguồn sheet SHA-256 %s."""' % S.NGUON_SHA256,
          'from battu_core import xl as X', '', "NGUON_SHA256 = '%s'" % S.NGUON_SHA256, '']
    bang = {}; so = 0
    for sheet, ten in S.MODULES.items():
        m = importlib.import_module('battu_core.sheets.' + ten); xong = {}
        for ref, (fn, r, c) in m.CELLS.items():
            if fn.__name__ in xong: continue
            cay = ast.parse(textwrap.dedent(inspect.getsource(fn)))
            bc = _BoCat(); cay = ast.fix_missing_locations(bc.visit(cay))
            if not bc.so: xong[fn.__name__] = None; continue
            moi = '_%s__%s' % (ten, fn.__name__)
            cay.body[0].name = moi
            ra.append(ast.unparse(cay)); ra.append('')
            xong[fn.__name__] = moi; so += 1
        ds = {k: v for k, v in xong.items() if v}
        if ds: bang[sheet] = ds
    ra.append('BAN_SAO = {')
    for sheet, ds in bang.items():
        ra.append('    %r: {%s},' % (sheet, ', '.join('%r: %s' % (k, v) for k, v in ds.items())))
    ra.append('}')
    p = os.path.join(GOC, 'export', '_khong_cat_sinh.py')
    open(p, 'w', encoding='utf-8').write('\n'.join(ra) + '\n')
    print('Đã sinh %s: %d hàm, sheet %s' % (os.path.relpath(p, GOC), so, ', '.join(bang)))


if __name__ == '__main__':
    main()
