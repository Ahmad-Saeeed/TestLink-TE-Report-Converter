# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec template. Copy into the new project as app.spec and:
#   - replace every TestLinkWEFormatter with the app's actual name
#   - list every bundled resource folder in datas (mirror resource_path() calls)
#   - add any extra hiddenimports the wrapped skill's dependencies need
#
# This is onedir (not onefile), UPX off, with a version resource wired in —
# the three defaults this skill always applies to avoid antivirus false
# positives. See references/pyinstaller-guide.md "Antivirus false positives"
# for why each of these three matters, and its onefile-alternative note if
# the user explicitly wants a single-file EXE instead and accepts the
# tradeoff.

from PyInstaller.utils.hooks import collect_data_files

# CustomTkinter ships theme JSON files and fonts inside its package folder;
# PyInstaller's static analysis doesn't see them, so collect them explicitly.
ctk_datas = collect_data_files('customtkinter')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    # the docx skeleton (header/footer/logos/styles) read via resource_path()
    datas=[('resources', 'resources')] + ctk_datas,
    # bs4 picks its parser backend dynamically; soupsieve is its selector engine
    hiddenimports=['bs4', 'soupsieve'],
    hookspath=[],
    runtime_hooks=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TestLinkWEFormatter',
    debug=False,
    strip=False,
    upx=False,
    console=True,
    version='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='TestLinkWEFormatter',
)
