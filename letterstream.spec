# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['letterstream.py'],
             pathex=[],
             binaries=[],
             datas=[('LetterForwardingTemplate.docx', '.'), ('CompanyContacts.csv', '.'), ('Batch_Template.csv', '.'), ('requirements.txt', '.')],
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
          [('v', None, 'OPTION')],
          exclude_binaries=True,
          name='letterstream',
          debug=True,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='C:\\Users\\tophl\\Desktop\\letterstream_logo.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas, 
               strip=False,
               upx=True,
               upx_exclude=[],
               name='letterstream')
