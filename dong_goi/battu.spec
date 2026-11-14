# -*- mode: python ; coding: utf-8 -*-
# PyInstaller – dựng thư mục chạy (onedir) cho Windows và .app cho macOS. Chạy từ thư mục gốc mã nguồn:
#   pyinstaller --noconfirm dong_goi/battu.spec
import sys, os
from PyInstaller.utils.hooks import collect_submodules
GOC = os.path.abspath(os.path.join(SPECPATH, '..'))
datas = [(os.path.join(GOC, 'battu_core', 'data'), os.path.join('battu_core', 'data')),
         (os.path.join(GOC, 'ui', 'fonts'), os.path.join('ui', 'fonts')),
         (os.path.join(GOC, 'ui', 'bieu_tuong.png'), 'ui')]
sys.path.insert(0, GOC)
import glob
# Module sheet được lõi nạp động theo tên (engine.cell_modules) → liệt kê tường minh mọi tệp .py
hidden = ['%s.%s' % (pk, os.path.splitext(os.path.basename(f))[0]) for pk in ('battu_core', 'battu_core.sheets', 'battu_core.xl', 'export', 'ui')
          for f in glob.glob(os.path.join(GOC, *pk.split('.'), '*.py')) if not f.endswith('__init__.py')]
hidden += ['battu_core.sheets', 'battu_core.xl', 'export', 'ui', 'dau_van_tay_chuan']
a = Analysis([os.path.join(GOC, 'battu_app.py')], pathex=[GOC], datas=datas, hiddenimports=hidden,
             excludes=['tkinter', 'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.Qt3DCore',
                       'PySide6.QtMultimedia', 'PySide6.QtQuick', 'PySide6.QtQml', 'PySide6.QtNetworkAuth', 'tests', 'tools'],
             noarchive=False)
pyz = PYZ(a.pure)
icon = os.path.join(SPECPATH, 'bieu_tuong.ico' if sys.platform == 'win32' else 'bieu_tuong.icns')
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='BatTuThieuViHoa', console=False, icon=icon,
          version=os.path.join(SPECPATH, 'phien_ban_win.txt') if sys.platform == 'win32' else None)
coll = COLLECT(exe, a.binaries, a.datas, name='BatTuThieuViHoa')
if sys.platform == 'darwin':
    app = BUNDLE(coll, name='Bát Tự Thiệu Vĩ Hoa.app', icon=icon, bundle_identifier='vn.battu.thieuvihoa',
                 info_plist={'CFBundleShortVersionString': '1.1.0', 'CFBundleVersion': '1.1.0',
                             'NSHighResolutionCapable': True, 'LSMinimumSystemVersion': '11.0',
                             'CFBundleDisplayName': 'Bát Tự Thiệu Vĩ Hoa'})
