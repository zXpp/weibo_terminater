# -*- mode: python -*-

block_cipher = None


a = Analysis(['Easy_PPT_Kits_v9_4.py'],
             pathex=['D:\\untar\\Panel\\Gudie'],
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
a.datas+=Tree('guidata\\images',prefix='guidata\\images',excludes=[''],typecode='DATA')
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Easy_PPT_Kits_v9_4',
          debug=False,
          strip=False,
          upx=True,
          console=False )
