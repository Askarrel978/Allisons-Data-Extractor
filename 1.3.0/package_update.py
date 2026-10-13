"""Create a USB-installable update ZIP from the latest PyInstaller build."""

from __future__ import annotations

import hashlib
import json
import os
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from version import APP_PRODUCT, APP_REVISION, APP_UPDATE_PRODUCT_ID


ROOT = Path(__file__).resolve().parent
DIST_DIR = ROOT / "dist" / "Allisons Data Extractor"
RELEASE_DIR = ROOT / "release"
ENTRY_EXE = "Allisons Data Extractor.exe"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_package() -> Path:
    if not DIST_DIR.is_dir():
        raise SystemExit(f"Build folder not found: {DIST_DIR}")
    if not (DIST_DIR / ENTRY_EXE).is_file():
        raise SystemExit(f"Main executable not found: {DIST_DIR / ENTRY_EXE}")

    RELEASE_DIR.mkdir(parents=True, exist_ok=True)
    output = RELEASE_DIR / f"Allisons_Data_Extractor_Update_v{APP_REVISION}.zip"

    checksums = {}
    files = [path for path in DIST_DIR.rglob("*") if path.is_file()]
    for path in files:
        relative = path.relative_to(DIST_DIR).as_posix()
        checksums[relative] = sha256(path)

    manifest = {
        "product": APP_PRODUCT,
        "product_id": APP_UPDATE_PRODUCT_ID,
        "version": APP_REVISION,
        "entry_exe": ENTRY_EXE,
        "created_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "sha256": checksums,
    }

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        archive.writestr("update_manifest.json", json.dumps(manifest, indent=2))
        for path in files:
            relative = path.relative_to(DIST_DIR).as_posix()
            archive.write(path, f"payload/{relative}")

    print(f"Created update package: {output}")
    return output


if __name__ == "__main__":
    build_package()
