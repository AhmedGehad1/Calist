# PyInstaller build spec for Calist.
#
#   pyinstaller calist.spec --noconfirm                     -> dist/Calist.exe
#   CALIST_ONEDIR=1 pyinstaller calist.spec --noconfirm     -> dist/Calist/
#
# Both shapes are published. The single file is the convenient one; the folder
# is the one that gets past antivirus. See ANTIVIRUS below.
#
# Three things here are not optional:
#
#   hiddenimports=["ui", "access"] — calist.main() imports the UI lazily inside
#       the function body (so `import calist` stays GUI-free). PyInstaller's
#       static analysis cannot see that, and the frozen app dies at launch
#       without it.
#
#   collect_data_files("customtkinter") — CTk loads its themes and Roboto fonts
#       from JSON/OTF files at runtime. Miss them and the window fails to draw.
#
#   upx=False — UPX-packing an unsigned executable is one of the strongest
#       signals antivirus heuristics look for. It saves a few MB and costs the
#       download its reputation.
#
# ── ANTIVIRUS ────────────────────────────────────────────────────────────────
#
# Windows Defender flagged the v1.1.x download as Trojan:Win32/Sabsik.FL.A!ml.
# The !ml suffix means a machine-learning verdict rather than a signature match
# — a false positive, and a well-known one for PyInstaller. Three properties of
# the old build drove it, and all three are addressed here:
#
#   1. A one-file build unpacks itself into %TEMP%\_MEIxxxx at launch and runs
#      from there. Self-extract-then-execute is textbook malware behaviour and
#      is the single heaviest signal. The CALIST_ONEDIR build does not do it.
#   2. The executable carried no version resource at all — no company, no
#      product, no description. Legitimate software has these; packed malware
#      usually does not. `version=` below fills them in.
#   3. It is unsigned, and unsigned plus zero download history means no
#      reputation to weigh against the heuristic. Only a code-signing
#      certificate fixes that one.

import os
import re
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

ONEDIR = os.environ.get("CALIST_ONEDIR") == "1"

# ── version resource ─────────────────────────────────────────────────────────
# Read straight out of calist.py rather than importing it, so building never
# depends on the app's runtime imports resolving. CALIST_VERSION lets the
# release workflow stamp the git tag in, which keeps the tag and the file's
# Properties tab from ever disagreeing.
HERE = Path(SPECPATH)
_source = (HERE / "calist.py").read_text(encoding="utf-8")
VERSION = os.environ.get("CALIST_VERSION") or \
    re.search(r'^__version__ = "([^"]+)"', _source, re.M).group(1)

_parts = (VERSION.split(".") + ["0", "0", "0", "0"])[:4]
_vers = tuple(int(re.sub(r"\D", "", p) or 0) for p in _parts)

# ASCII only. This text is read back by Explorer, Task Manager and the
# SmartScreen prompt, and a stray non-ASCII byte renders as mojibake there.
VERSION_FILE = HERE / "build" / "version.txt"
VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
VERSION_FILE.write_text(f"""\
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={_vers},
    prodvers={_vers},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0),
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', 'Ahmed Gehad'),
          StringStruct('FileDescription',
                       'Calist - compile device inspection forms into one equipment register'),
          StringStruct('FileVersion', '{VERSION}'),
          StringStruct('InternalName', 'Calist'),
          StringStruct('LegalCopyright',
                       'Copyright (c) 2026 Ahmed Gehad. All rights reserved.'),
          StringStruct('OriginalFilename', 'Calist.exe'),
          StringStruct('ProductName', 'Calist'),
          StringStruct('ProductVersion', '{VERSION}'),
        ])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
""", encoding="ascii")

# ── analysis ─────────────────────────────────────────────────────────────────

datas = collect_data_files("customtkinter")

# Ship the reference template inside the executable, so a freshly downloaded
# copy has something to build against without hunting for one first.
datas += [("template/Device List.xlsx", "template")]

# The window icon. Keeping the docs/ prefix means ui.app_icon() resolves it the
# same way whether the app is frozen or running from a checkout.
datas += [("docs/calist.ico", "docs")]

a = Analysis(
    ["calist.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=["ui", "access"],
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

common = dict(
    name="Calist",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                  # see ANTIVIRUS above
    console=False,              # desktop app: no terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="docs/calist.ico",
    version=str(VERSION_FILE),
)

if ONEDIR:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, **common)
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False,
                   name="Calist")
else:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],
              runtime_tmpdir=None, **common)
