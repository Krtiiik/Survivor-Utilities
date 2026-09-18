# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the Survivor desktop app.

Used both for local builds (`pyinstaller survivor.spec`) and by
.github/workflows/build.yml's Windows/Linux matrix. Bundles a single
self-contained executable (--onefile) since the target audience is a
non-technical successor who should be able to just download and double-click
one file, at the cost of a slower cold start than --onedir.
"""
import sys

from PyInstaller.utils.hooks import collect_all

datas = [("survivor_app/resources/example_config.json", "survivor_app/resources")]
binaries = []
hiddenimports = []

# ortools and PySide6 are both large, native-library-heavy packages with a real
# history of PyInstaller hook gaps -- collect_all is the safest default.
for package in ("ortools", "PySide6"):
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(package)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hiddenimports

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Survivor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    # Windows: no console window (successor just double-clicks the app).
    # Linux: keep the console attached, so anyone troubleshooting a launch
    # failure can see the traceback without extra setup.
    console=(sys.platform != "win32"),
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
