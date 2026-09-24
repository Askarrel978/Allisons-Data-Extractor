"""Standalone updater helper for Allison's Data Extractor.

Built as its own one-file EXE. The main app copies this helper to a temporary
folder before launching it, allowing the installed application folder to be
replaced safely after the main process exits.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path


def wait_for_process(pid: int, timeout_seconds: int = 60) -> None:
    if os.name != "nt":
        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            try:
                os.kill(pid, 0)
            except OSError:
                return
            time.sleep(0.25)
        return

    SYNCHRONIZE = 0x00100000
    WAIT_OBJECT_0 = 0x00000000
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(SYNCHRONIZE, False, int(pid))
    if not handle:
        return
    try:
        result = kernel32.WaitForSingleObject(handle, int(timeout_seconds * 1000))
        if result != WAIT_OBJECT_0:
            time.sleep(1.0)
    finally:
        kernel32.CloseHandle(handle)


def backup_install(target: Path, version: str) -> Path | None:
    if not target.exists():
        return None
    root = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "AllisonsDataExtractor" / "update_backups"
    root.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    safe_version = version.replace("/", "_").replace("\\", "_") or "unknown"
    backup = root / f"before_{safe_version}_{stamp}"
    shutil.copytree(target, backup, dirs_exist_ok=True)
    return backup


def replace_install(source: Path, target: Path) -> None:
    if not source.is_dir():
        raise RuntimeError("Staged update payload is missing.")

    target_parent = target.parent
    target_parent.mkdir(parents=True, exist_ok=True)
    old = target_parent / f"{target.name}.update_old"

    if old.exists():
        shutil.rmtree(old, ignore_errors=True)

    # Renaming the old folder first avoids stale DLL/config files surviving an
    # update. The helper itself is running from a temp folder, so this is safe.
    if target.exists():
        os.replace(target, old)

    try:
        shutil.copytree(source, target)
    except Exception:
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)
        if old.exists():
            os.replace(old, target)
        raise

    shutil.rmtree(old, ignore_errors=True)


def write_update_notice(version: str) -> None:
    root = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "AllisonsDataExtractor"
    root.mkdir(parents=True, exist_ok=True)
    notice = {
        "version": version,
        "completed_at": datetime.now().isoformat(timespec="seconds"),
    }
    (root / "last_update.json").write_text(json.dumps(notice, indent=2), encoding="utf-8")


def relaunch(target: Path, entry_exe: str) -> None:
    exe = target / entry_exe
    if not exe.is_file():
        raise RuntimeError(f"Updated executable was not found: {exe}")
    os.startfile(str(exe))


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--wait-pid", type=int, required=True)
    parser.add_argument("--source", required=True)
    parser.add_argument("--target", required=True)
    parser.add_argument("--entry-exe", required=True)
    parser.add_argument("--version", default="")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source = Path(args.source).resolve()
    target = Path(args.target).resolve()

    wait_for_process(args.wait_pid)
    backup_install(target, args.version)
    replace_install(source, target)
    write_update_notice(args.version)
    relaunch(target, args.entry_exe)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        # The helper normally has no console, so retain a simple emergency log.
        root = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "AllisonsDataExtractor"
        root.mkdir(parents=True, exist_ok=True)
        with (root / "update_error.log").open("a", encoding="utf-8") as handle:
            handle.write(f"{datetime.now().isoformat(timespec='seconds')} - {exc}\n")
        if os.name == "nt":
            try:
                ctypes.windll.user32.MessageBoxW(
                    0,
                    f"The update could not be completed.\n\n{exc}\n\nDetails were saved to:\n{root / 'update_error.log'}",
                    "Allison's Data Extractor Update Error",
                    0x10,
                )
            except Exception:
                pass
        raise
