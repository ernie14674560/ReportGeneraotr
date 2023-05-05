# -*- mode: python ; coding: utf-8 -*-
import os

spec_root = os.path.abspath(SPECPATH)

block_cipher = None
added_files = [('config.yml', '.'),('icon.png', '.'),('icon2.png', '.'),('icon3.png', '.'), ('Defect code lists.xlsx', '.'),('code adjustment', 'code adjustment'),  ('reference', 'reference'), ('templates', 'templates'), (HOMEPATH + '\\PyQt5\\Qt\\bin\*', 'PyQt5\\Qt\\bin'), (HOMEPATH + '\\pandas\\io\\formats\\templates\\html.tpl','.')]

a = Analysis(['app.py'],
             pathex=[spec_root],
             binaries=[],
             datas=added_files,
             hiddenimports=["pandas", "jinja2"],
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
          [],
          exclude_binaries=True,
          name='ReportGenerator',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          icon='C:\\Users\ernie.chen\Desktop\Project\Project Week\icon.ico',
          console=False)
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='app')
