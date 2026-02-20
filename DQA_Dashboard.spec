# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['run_standalone.py'],
    pathex=[],
    binaries=[],
    datas=[('app', 'app'), ('data', 'data'), ('config', 'config'), ('assets', 'assets'), ('icon.ico', '.')],
    hiddenimports=['flask', 'werkzeug', 'jinja2', 'openpyxl', 'pandas', 'numpy', 'app.routes', 'app.analysis', 'app.storage', 'app.__init__', 'flask.cli', 'jinja2.ext', 'datetime', 'uuid', 'json', 'os', 'sys', 'logging', 'shutil', 'socket', 'webbrowser', 'threading'],
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
    name='DQA_Dashboard',
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
    icon=['icon.ico'],
)
