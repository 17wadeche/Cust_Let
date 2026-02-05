# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import importlib.util

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

# --- Find the installed playwright package location reliably ---
pw_spec = importlib.util.find_spec("playwright")
if not pw_spec or not pw_spec.origin:
    raise RuntimeError("Playwright is not installed in this environment.")

playwright_pkg_dir = Path(pw_spec.origin).parent  # ...\site-packages\playwright
pw_driver_pkg_dir = playwright_pkg_dir / "driver" / "package"

print(f"Playwright package dir: {playwright_pkg_dir}")
print(f"Playwright driver package dir: {pw_driver_pkg_dir}")

datas_list = [
    (str(project_dir / "config.yaml"), "."),
    (str(project_dir / "customer_letter_template.docx"), "."),
]

# Bundle the entire driver/package folder (this includes .local-browsers if installed there)
if pw_driver_pkg_dir.exists():
    print(f"✓ Found Playwright driver package at: {pw_driver_pkg_dir}")
    datas_list += datas_from_dir(pw_driver_pkg_dir, "playwright/driver/package")
else:
    print(f"✗ Playwright driver package NOT found at: {pw_driver_pkg_dir}")

a = Analysis(
    ["ui_app.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas_list,
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
    [],
    exclude_binaries=True,
    name="CustomerLetterGeneratorv3",
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
    name="CustomerLetterGeneratorv3",
)