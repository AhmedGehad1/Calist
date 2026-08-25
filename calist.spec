# PyInstaller build spec for Calist.
#
#   pyinstaller calist.spec --noconfirm
#
# Two things here are not optional:
#
#   hiddenimports=["ui"]  — calist.main() imports the UI lazily inside the
#       function body (so `import calist` stays GUI-free). PyInstaller's static
#       analysis cannot see that, and the frozen app dies at launch without it.
#
#   collect_data_files("customtkinter") — CTk loads its themes and Roboto fonts
#       from JSON/OTF files at runtime. Miss them and the window fails to draw.

from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files("customtkinter")

# Ship the reference template inside the executable, so a freshly downloaded
# copy has something to build against without hunting for one first.
datas += [("template/Device List.xlsx", "template")]

a = Analysis(
    ["calist.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=["ui"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Nothing in the app touches these; leaving them in doubles the size.
        "numpy", "matplotlib", "scipy", "pandas",
        "pytest", "_pytest", "pygments", "IPython",
        "PyQt5", "PySide2", "PySide6",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Calist",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,              # desktop app: no terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="docs/calist.ico",
)
