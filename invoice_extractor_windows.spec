# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ─── 收集所有依赖 ─────────────────────────────
datas = []
hiddenimports = [
    "PIL", "PIL._imaging", "fitz", "fitz.fitz", "fitz.mupdf",
    "openpyxl", "openpyxl.xlsx.reader", "openpyxl.styles",
    "configparser", "tkinter", "tkinter.ttk", "tkinter.scrolledtext",
    "tkinter.filedialog", "tkinter.messagebox",
]
# 额外收集 PyMuPDF 的资源文件（字体等）
datas += collect_data_files("fitz")

a = Analysis(
    ['invoice_extractor.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "scipy"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='发票识别工具',
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
    icon=None,
    version=None,
    uac_admin=False,
    uac_uiAccess=False,
)
