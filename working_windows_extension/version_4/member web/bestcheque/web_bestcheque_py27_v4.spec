# -*- mode: python -*-

block_cipher = None


a = Analysis(['web_bestcheque_py27_v4.py'],
             pathex=['C:\\Users\\qiman\\Desktop\\Github\\Joyce_stuff\\Joyce\\working_windows_extension\\version_4\\member web\\bestcheque'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='web_bestcheque_py27_v4',
          debug=False,
          strip=False,
          upx=True,
          console=True , icon='C:\\Users\\qiman\\Desktop\\Github\\Joyce_stuff\\Joyce\\working_windows_extension\\version_4\\member web\\bestcheque\\mb.ico')
