# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['app.py'],
             pathex=['/Users/lordsonvimal/projects/excel'],
             binaries=[],
             datas=[],
             hiddenimports=['cmath'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='app',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='icon.ico')
app = BUNDLE(exe,
             name='app.app',
             icon='icon.ico',
             bundle_identifier=None)
