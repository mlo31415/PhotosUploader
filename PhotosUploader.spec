# -*- mode: python ; coding: utf-8 -*-
import os

a = Analysis(
    ['PhotosUploader.py'],
    # PiwigoHelpers lives one level up; tell PyInstaller to search it so that
    # DownloadAlbumStructure (imported via a runtime sys.path insert) is found
    # during analysis and bundled into the exe.
    pathex=[os.path.join(SPECPATH, '..', 'PiwigoHelpers')],
    binaries=[],
    # Bundle the icon so root.iconbitmap() can find it at runtime (via sys._MEIPASS)
    datas=[('PhotosUploader.ico', '.')],
    hiddenimports=['DownloadAlbumStructure'],
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
    name='PhotosUploader',
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
    icon=['PhotosUploader.ico'],
)
