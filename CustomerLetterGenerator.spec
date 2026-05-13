# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import importlib.util
import os
project_dir = Path(SPECPATH)
def datas_from_dir(src_dir: Path, dest_root: str):
    out = []
    src_dir = Path(src_dir)
    for p in src_dir.rglob("*"):
        if p.is_file():
            rel_parent = p.parent.relative_to(src_dir)
            dest = str(Path(dest_root) / rel_parent).replace("\\", "/")
            out.append((str(p), dest))
    return out
pw_spec = importlib.util.find_spec("playwright")
if not pw_spec or not pw_spec.origin:
    raise RuntimeError(
        "Playwright is not installed in the Python environment running PyInstaller. "
        "Run `python -m pip install -r requirements.txt` first, or run "
        "`powershell -ExecutionPolicy Bypass -File ./build_windows.ps1` to install "
        "build dependencies, install Chromium, and package the app."
    )
playwright_pkg_dir = Path(pw_spec.origin).parent  # ...\site-packages\playwright
pw_driver_pkg_dir = playwright_pkg_dir / "driver" / "package"
print(f"Playwright package dir: {playwright_pkg_dir}")
print(f"Playwright driver package dir: {pw_driver_pkg_dir}")
def playwright_browser_cache_candidates():
    candidates = []
    env_path = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if env_path and env_path != "0":
        candidates.append(Path(env_path))
    candidates += [
        pw_driver_pkg_dir / ".local-browsers",
        Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright",
        Path.home() / "AppData" / "Local" / "ms-playwright",
    ]
    seen = set()
    for candidate in candidates:
        if not str(candidate):
            continue
        candidate = candidate.expanduser()
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists():
            yield candidate
datas_list = [
    (str(project_dir / "config.yaml"), "."),
    (str(project_dir / "customer_letter_template.docx"), "."),
]
# Bundle the entire driver/package folder. This includes Playwright's Node driver.
if pw_driver_pkg_dir.exists():
    print(f"✓ Found Playwright driver package at: {pw_driver_pkg_dir}")
    datas_list += datas_from_dir(pw_driver_pkg_dir, "playwright/driver/package")
else:
    print(f"✗ Playwright driver package NOT found at: {pw_driver_pkg_dir}")
# Bundle installed Chromium browsers into the package-local location expected by ui_app.py.
# This supports both `PLAYWRIGHT_BROWSERS_PATH=0` installs and normal `%LOCALAPPDATA%\ms-playwright` installs.
browser_cache_found = False
for browser_cache_dir in playwright_browser_cache_candidates():
    print(f"✓ Found Playwright browser cache at: {browser_cache_dir}")
    datas_list += datas_from_dir(browser_cache_dir, "playwright/driver/package/.local-browsers")
    browser_cache_found = True
if not browser_cache_found:
    print(
        "✗ Playwright Chromium browser cache not found. Run "
        "`python -m playwright install chromium` before packaging."
    )
a = Analysis(
    ["ui_app.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas_list,
    hiddenimports=["pythoncom", "pywintypes", "win32com", "win32com.client"],
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
    [],
    exclude_binaries=True,
    name="CustomerLetterGenerator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="CustomerLetterGenerator",
)