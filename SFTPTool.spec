# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    'paramiko',
    'cryptography',
    'bcrypt',
    'nacl',
    'openpyxl',
    'xlsxwriter',
    'python_calamine',
    'sqlite3',
    'darkdetect',
    'customtkinter'
]

# 收集 python_calamine 的編譯二進位模組
tmp_cal = collect_all('python_calamine')
datas += tmp_cal[0]
binaries += tmp_cal[1]
hiddenimports += tmp_cal[2]

# 收集 customtkinter 的主題 JSON 與字型資產
tmp_ctk = collect_all('customtkinter')
datas += tmp_ctk[0]
binaries += tmp_ctk[1]
hiddenimports += tmp_ctk[2]

# 收集 paramiko 相關通訊依賴
tmp_prm = collect_all('paramiko')
datas += tmp_prm[0]
binaries += tmp_prm[1]
hiddenimports += tmp_prm[2]

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SFTP批次下載工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
