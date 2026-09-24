"""USB/local update package validation and staged update launching."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from settings_manager import update_work_dir
from version import APP_UPDATE_PRODUCT_ID


class UpdateError(RuntimeError):
    pass


def _version_tuple(value: str) -> tuple[int, ...]:
    parts = str(value or "").strip().lstrip("vV").split(".")
    if not parts or any(not part.isdigit() for part in parts):
        raise UpdateError(f'Invalid revision number "{value}" in update package.')
    return tuple(int(part) for part in parts)


def is_newer_version(candidate: str, current: str) -> bool:
    a = list(_version_tuple(candidate))
    b = list(_version_tuple(current))
    length = max(len(a), len(b))
    a.extend([0] * (length - len(a)))
    b.extend([0] * (length - len(b)))
    return tuple(a) > tuple(b)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def inspect_update_package(zip_path: str, current_version: str) -> dict:
    package = Path(zip_path)
    if not package.is_file():
        raise UpdateError("The selected update package does not exist.")

    try:
        with zipfile.ZipFile(package, "r") as archive:
            names = set(archive.namelist())
            if "update_manifest.json" not in names:
                raise UpdateError("This ZIP is not an Allison's Data Extractor update package (manifest missing).")
            manifest = json.loads(archive.read("update_manifest.json").decode("utf-8"))
    except zipfile.BadZipFile as exc:
        raise UpdateError("The selected update package is not a valid ZIP file.") from exc
    except json.JSONDecodeError as exc:
        raise UpdateError("The update manifest is invalid JSON.") from exc

    if manifest.get("product_id") != APP_UPDATE_PRODUCT_ID:
        raise UpdateError("The selected update package is for a different application.")

    version = str(manifest.get("version", "")).strip()
    if not version:
        raise UpdateError("The update package does not specify a revision.")
    if not is_newer_version(version, current_version):
        raise UpdateError(
            f"Revision {version} is not newer than the currently installed revision {current_version}."
        )

    if not manifest.get("entry_exe"):
        raise UpdateError("The update manifest does not specify the application executable.")

    return manifest


def stage_update_package(zip_path: str, current_version: str) -> tuple[dict, str]:
    manifest = inspect_update_package(zip_path, current_version)
    work_root = update_work_dir()
    stage = Path(tempfile.mkdtemp(prefix=f"stage_{manifest['version']}_", dir=str(work_root)))

    try:
        with zipfile.ZipFile(zip_path, "r") as archive:
            # Defend against ZIP path traversal.
            stage_resolved = stage.resolve()
            for info in archive.infolist():
                destination = (stage / info.filename).resolve()
                if stage_resolved not in destination.parents and destination != stage_resolved:
                    raise UpdateError("Unsafe path found inside update package.")
            archive.extractall(stage)

        payload = stage / "payload"
        if not payload.is_dir():
            raise UpdateError("Update package payload is missing.")

        checksums = manifest.get("sha256", {})
        if not isinstance(checksums, dict) or not checksums:
            raise UpdateError("Update package does not include file-integrity hashes.")

        for relative_name, expected in checksums.items():
            candidate = payload / relative_name
            if not candidate.is_file():
                raise UpdateError(f"Update package is missing required file: {relative_name}")
            actual = _sha256(candidate)
            if actual.lower() != str(expected).lower():
                raise UpdateError(f"Update package integrity check failed for: {relative_name}")

        return manifest, str(payload)
    except Exception:
        shutil.rmtree(stage, ignore_errors=True)
        raise


def application_directory() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def updater_executable_path() -> str:
    base = Path(application_directory())
    candidates = [
        base / "Allisons Data Extractor Updater.exe",
        base / "_internal" / "Allisons Data Extractor Updater.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return ""


def launch_staged_update(manifest: dict, payload_dir: str) -> None:
    if not getattr(sys, "frozen", False):
        raise UpdateError("Updates can only be installed from the compiled portable application.")

    updater = updater_executable_path()
    if not updater:
        raise UpdateError(
            "The updater helper is missing from this installation. Reinstall the current portable build first."
        )

    target_dir = application_directory()
    entry_exe = str(manifest["entry_exe"])

    # Run a temporary copy so the updater itself can be replaced by the payload.
    temp_dir = Path(tempfile.mkdtemp(prefix="allisons_data_extractor_updater_", dir=str(update_work_dir())))
    temp_updater = temp_dir / "Allisons Data Extractor Updater.exe"
    shutil.copy2(updater, temp_updater)

    args = [
        str(temp_updater),
        "--wait-pid", str(os.getpid()),
        "--source", str(payload_dir),
        "--target", str(target_dir),
        "--entry-exe", entry_exe,
        "--version", str(manifest.get("version", "")),
    ]

    creationflags = 0
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(args, close_fds=True, creationflags=creationflags)


def consume_update_notice(current_version: str) -> dict | None:
    """Return and remove the last successful-update notice for this revision."""
    notice_path = update_work_dir().parent / "last_update.json"
    if not notice_path.is_file():
        return None
    try:
        data = json.loads(notice_path.read_text(encoding="utf-8"))
    except Exception:
        try:
            notice_path.unlink()
        except Exception:
            pass
        return None
    if str(data.get("version", "")) != str(current_version):
        return None
    try:
        notice_path.unlink()
    except Exception:
        pass
    return data
