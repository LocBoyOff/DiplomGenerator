# Build with: python -m PyInstaller --noconfirm DiplomaGenerator.spec
from pathlib import Path

root = Path(SPECPATH)
a = Analysis(
    [str(root / 'diploma_generator.py')],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / 'assets' / 'app.ico'), 'assets')],
    hiddenimports=['comtypes.persist', 'comtypes.stream'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['comtypes.test', 'tkinter', 'unittest', 'pytest'],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='DiplomaGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    version=str(root / 'assets' / 'version_info.txt'),
    icon=str(root / 'assets' / 'app.ico'),
)
