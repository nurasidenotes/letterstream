# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['smith_downey_letters.py'],
             pathex=[],
             binaries=[],
             datas=[('C:\\Users\\tophl\\Documents\\EL_Programming\\Letterstream\\CompanyContacts.csv', '.'), ('C:\\Users\\tophl\\Documents\\EL_Programming\\Letterstream\\template_smith_downey.docx', '.'), ('C:\\Users\\tophl\\Documents\\EL_Programming\\Letterstream\\Batch_Template.csv', '.')],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=True)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [('v', None, 'OPTION')],
          name='Smith_Downey_letterstream',
          debug=True,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='C:\\Users\\tophl\\Desktop\\letterstream_logo.ico')
