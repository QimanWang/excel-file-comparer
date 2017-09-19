# -*- mode: python -*-

block_cipher = None


a = Analysis(['compare_booking_py27_V3.py'],
             pathex=['C:\\Users\\qiman\\Desktop\\github\\joyce_stuff\\joyce_new\\Joyce\\working_windows_extension\\version_3\\booking'],
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
          name='compare_booking_py27_V3',
          debug=False,
          strip=False,
          upx=True,
          console=True , icon='C:\\Users\\qiman\\Desktop\\github\\joyce_stuff\\joyce_new\\Joyce\\working_windows_extension\\version_3\\booking\\icon.ico')
