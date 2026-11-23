# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main1.py'],
    pathex=[],
    binaries=[],
    datas=[('Exemple_Import_Profil_En_Long.xlsx', '.'), ('Exemple_Import_Cas_Simples.xlsx', '.'), ('Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx', '.'), ('Tableau_Recapitulatif_Debits_Ventouses_MCP.docx', '.'), ('assets/MK_icon.ico', 'assets')],
    hiddenimports=[],
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
    name='MK_Prandtl-Colebrook_V01',
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
    icon=['assets/MK_icon.ico'],
)
