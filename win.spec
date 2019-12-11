# -*- mode: python -*-

block_cipher = None


a = Analysis(['C:\\Users\\aliss\\Desktop\\Dropbox\\Faculdade\\PycharmProjects\\SemabioToolKit\\branchs\\alpha\\0.1.1\\launcher.py'],
             pathex=['C:\\Users\\aliss\\Dropbox\\Faculdade\\PycharmProjects\\SemabioToolKit\\branchs\\alpha\\0.1.1', 'C:\\Users\\aliss\\Desktop\\Dropbox\\Faculdade\\PycharmProjects\\SemabioToolKit\\branchs\\alpha\\0.1.1\\.pyupdater\\spec'],
             binaries=[],
             datas=[('C:\\Users\\aliss\\Desktop\\Dropbox\\Faculdade\\PycharmProjects\\SemabioToolKit\\branchs\\alpha\\0.1.1\\icons\\*', 'icons')],
             hiddenimports=['SocketServer','PIL._tkinter_finder'],
             hookspath=['c:\\users\\aliss\\appdata\\local\\programs\\python\\python36\\lib\\site-packages\\pyupdater\\hooks'],
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
          name='win',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False) 
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='win')
