# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = [('splitvision.ico', '.')]
datas += collect_data_files('pypdfium2')
datas += collect_data_files('pdf2docx')


a = Analysis(
    ['splitvision.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['winrt', 'winrt.windows.media.ocr', 'winrt.windows.globalization', 'winrt.windows.storage.streams', 'winrt.windows.graphics.imaging', 'winrt.windows.foundation', 'winrt.windows.foundation.collections', 'winocr', 'windnd', 'pdf2docx', 'docx', 'docx2pdf', 'win32com', 'conversor_pdf_word'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'tkinter.test'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    name='splitvision',
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
    icon=['splitvision.ico'],
)
