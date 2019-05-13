# -*- mode: python -*-

import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)
block_cipher = None



a = Analysis(['interface.py'],
             pathex=['C:\\Users\\Corentin\\ProjetAnnuel'],
             binaries=[],
             datas=[(path.join(site_packages,"docx","templates"), "docx/templates")],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
a.datas += [('imageGauche.png','C:\\Users\\Corentin\\ProjetAnnuel\\imageGauche.png', "DATA")]
a.datas += [('ansm.jpg','C:\\Users\\Corentin\\ProjetAnnuel\\ansm.jpg', "DATA")]
a.datas += [('EvIG_partie_9.jpg','C:\\Users\\Corentin\\ProjetAnnuel\\EvIG_partie_9.jpg', "DATA")]
a.datas += [('imageDroite.png','C:\\Users\\Corentin\\ProjetAnnuel\\imageDroite.png', "DATA")]
a.datas += [('imageGauche2.png','C:\\Users\\Corentin\\ProjetAnnuel\\imageGauche2.png', "DATA")]
a.datas += [('imageGauche3.png','C:\\Users\\Corentin\\ProjetAnnuel\\imageGauche3.png', "DATA")]
a.datas += [('num_patient.png','C:\\Users\\Corentin\\ProjetAnnuel\\num_patient.png', "DATA")]

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='interface',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='interface')
