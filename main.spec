# -*- mode: python ; coding: utf-8 -*-
import os
from kivy_deps import sdl2, glew
spec_root = os.path.abspath(SPECPATH)

block_cipher = None

app_name="AllList-Converter"
win_icon= "C:\\Users\\David I Gonzalez\\Proyectos_python\\AllListConverter\\manzanaverde.ico"
a = Analysis(['./main.py'],
             pathex=[spec_root],
             binaries=[],
             datas=[("./All_List_Converter.kv","."),("./manzanaverde.ico",".")],
             hiddenimports=["win32timezone"],
             hookspath=[],
             hooksconfig={},
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
          name=app_name,
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
          console=False,
	  icon= win_icon,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
		*[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)], 
               strip=False,
               upx=True,
               upx_exclude=[],
               name=app_name)
