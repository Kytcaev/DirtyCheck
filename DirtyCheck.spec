# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('Summary of comments.xlsx', '.'), ('checks', 'checks')],
    hiddenimports=['checks.check_cyrillic_filename', 'checks.check_cyrillic_pdf', 'checks.check_footer', 'checks.check_pdf_comments', 'checks.check_duplicates', 'checks.check_package', 'checks.check_standards', 'checks.check_comments', 'PyPDF2', 'pdfplumber'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='DirtyCheck',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
