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

# ── run-log address ──────────────────────────────────────────────────────────
# Official builds report each Build press to the owner's sheet (runlog.py).
# The address and token arrive from the CI secrets as environment variables
# and are written into a module the app imports, under build/: git-ignored,
# and not on a checkout's path, so the source and the tests never have one.
# Anything but a web app's /exec address on script.google.com is ignored, and
# a stale module from an earlier build is deleted, so a mistyped secret gives
# a copy that sends nothing rather than one that sends anywhere. The values
# are never printed. The flag file tells the release workflow to look for the
# module in both executables without the workflow ever seeing the secret.
GEN = HERE / "build" / "_gen"
GEN.mkdir(parents=True, exist_ok=True)
ENDPOINT = GEN / "calist_endpoint.py"
INCLUDED = GEN / "endpoint-included"
_url = os.environ.get("CALIST_RUNLOG_URL", "").strip()
_token = os.environ.get("CALIST_RUNLOG_TOKEN", "").strip()
if re.fullmatch(r"https://script\.google\.com/(?:a/macros/[^/]+|macros)/s/[\w-]+/exec",
                _url):
    ENDPOINT.write_text(f"URL = {_url!r}\nTOKEN = {_token!r}\n", encoding="utf-8")
    INCLUDED.write_text("yes\n", encoding="ascii")
else:
    ENDPOINT.unlink(missing_ok=True)
    INCLUDED.unlink(missing_ok=True)
    # A step's `if:` cannot read secrets, so the guard lives here: a release
    # tag must not ship a copy that silently never reports. Local builds and
    # forks simply get one that sends nothing.
    if os.environ.get("GITHUB_REF", "").startswith("refs/tags/"):
        raise SystemExit("CALIST_RUNLOG_URL is missing or not a web app /exec "
                         "address, so this release would never report. Set the "
                         "repository secret, or build without a tag.")

# ── analysis ─────────────────────────────────────────────────────────────────

datas = collect_data_files("customtkinter")

# Ship the reference template inside the executable, so a freshly downloaded
# copy has something to build against without hunting for one first.
datas += [("template/Device List.xlsx", "template")]

# The window icon. Keeping the docs/ prefix means ui.app_icon() resolves it the
# same way whether the app is frozen or running from a checkout.
datas += [("docs/calist.ico", "docs")]
# The mark beside the wordmark and on the lock screen, one rendition per size
# and display scaling (ui.mark_image, drawn by docs/make_icon.py).
datas += [(str(path), "docs") for path in sorted(Path("docs").glob("calist-mark-*.png"))]

# Drag and drop: tkinterdnd2 ships no PyInstaller hook, and loads its tkdnd
# DLL from beside its own module (tkinterdnd2/tkdnd/<platform>), so that is
# where it goes. 64-bit Windows only — the build's own architecture.
import tkinterdnd2  # noqa: E402  (build time only)
datas += [(str(Path(tkinterdnd2.__file__).parent / "tkdnd" / "win-x64"),
           "tkinterdnd2/tkdnd/win-x64")]

a = Analysis(
    ["calist.py"],
    pathex=[str(GEN)],
    binaries=[],
    datas=datas,
    # runlog is imported by ui, and calist_endpoint by runlog inside a
    # function; listing both is free insurance for a frozen-only failure that
    # no test can see, the same reason ui and access are here.
    hiddenimports=["ui", "access", "runlog"]
    + (["calist_endpoint"] if ENDPOINT.exists() else []),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Nothing in the app touches these; leaving them in doubles the size.
        # Never add ssl, _ssl, http, email, urllib or socket: the run log's
        # sender needs them, and only the frozen exe would notice them gone.
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
