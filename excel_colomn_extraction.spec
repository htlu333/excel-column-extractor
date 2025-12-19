# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 配置文件
用于将 Excel 列提取工具打包为可执行文件
"""

block_cipher = None

a = Analysis(
    ['excel_colomn_extraction.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.cell',
        'openpyxl.cell._writer',
        'openpyxl.cell.text',
        'openpyxl.workbook',
        'openpyxl.workbook.workbook',
        'openpyxl.worksheet',
        'openpyxl.worksheet._write_only',
        'openpyxl.worksheet._writer',
        'openpyxl.worksheet.worksheet',
        'openpyxl.packaging',
        'openpyxl.packaging.manifest',
        'openpyxl.packaging.relationship',
        'openpyxl.xml',
        'openpyxl.xml.constants',
        'openpyxl.xml.functions',
        'et_xmlfile',
        'et_xmlfile.xmlfile',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'ctypes',
        'ctypes.windll',
        'threading',
    ],
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
    name='Excel列提取工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口（GUI应用）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 如果有图标文件，可以在这里指定路径，例如: 'icon.ico'
)

