"""Persistent settings and administrator authentication.

Settings are stored in the current Windows user's LOCALAPPDATA folder so the
portable application can be updated/replaced without losing its configuration.
Only a salted PBKDF2 password hash is stored; the administrator passcode itself
is never written to disk.
"""

from __future__ import annotations

import copy
import hashlib
import hmac
import json
import os
import secrets
import shutil
from datetime import datetime
from pathlib import Path


SETTINGS_SCHEMA_VERSION = 1
PBKDF2_ITERATIONS = 310_000

DEFAULT_SETTINGS = {
    "schema_version": SETTINGS_SCHEMA_VERSION,
    "general": {
        "remember_output_file": True,
        "last_output_file": "",
    },
    "formatting": {
        "normalize_pack_style": True,
        "pack_style_separator": " / ",
        "pack_style_aliases": {
            "12x1ctclam": "12x1ct Clam",
            "12x1ct clam": "12x1ct Clam",
            "8x2ctclam": "8x2ct Clam",
            "8x2ct clam": "8x2ct Clam",
        },
    },
    "ocr": {
        "template_b_penalty": 10,
        "verified_threshold": 95,
        "review_threshold": 70,
        "tesseract_cmd_override": "",
    },
    "templates": {
        "Template A": {
            "markers": ["ATTENTION US"],
            "source_type": "pdf",
            "lot_required": True,
            "sales_order_pattern": r"^\d{5,9}$",
        },
        "Template B": {
            "markers": ["LOT#"],
            "source_type": "image",
            "lot_required": True,
            "sales_order_pattern": r"^\d{5,9}$",
            "record_start_terms": ["Plant Species"],
        },
        "Template C": {
            "markers": ["*Lettuce"],
            "source_type": "pdf",
            "lot_required": False,
            "sales_order_pattern": r"^\d{5,9}$",
        },
    },
    "extraction": {
        "table_labels": {
            "Date": ["shipping date", "date"],
            "Sales Order Number": ["sales order #", "sales order", "sales order number", "sales order no.", "sales order no"],
            "Customer and Location": ["customer and address", "customer and location", "customer/address", "customer"],
            "Pack Style": ["packaging style", "pack style", "packing"],
            "Number of Plants": ["number of plants", "plant count", "quantity of plants", "qty of plants"],
            "Lot Code": ["lot id", "lot#", "lot #", "lot code", "lot number"],
        },
        "generic_rules": {
            "Date": [
                r"Shipping\s*Date\s*[:\-]\s*([0-9]{1,2}[/\-][0-9]{1,2}[/\-][0-9]{2,4})",
                r"Shipping\s*Date\s*[:\-]\s*([A-Za-z]+\s+[0-9]{1,2},?\s+[0-9]{4})",
                r"Date\s*[:\-]\s*([0-9]{1,2}[/\-][0-9]{1,2}[/\-][0-9]{2,4})",
                r"Date\s*[:\-]\s*([A-Za-z]+\s+[0-9]{1,2},?\s+[0-9]{4})",
            ],
            "Sales Order Number": [
                r"Sales\s*Order\s*(?:Number|No\.?|#)?\s*[:\-]\s*(?:SO\s*#\s*)?([A-Za-z0-9\-]+)",
                r"\bS\.?O\.?\s*(?:Number|No\.?|#)?\s*[:\-]\s*([A-Za-z0-9\-]+)",
            ],
            "Customer and Location": [
                r"Customer\s*and\s*(?:Address|Location)\s*:\s*([\s\S]*?)(?=\s*(?:Lot\s*(?:ID|#|Code|Number)|Grown\s*at|Exported\s*By)\s*:?)",
            ],
            "Pack Style": [
                r"Packaging\s*Style\s*:\s*([\s\S]*?)(?=\s*Number\s*of\s*Plants\s*:)",
                r"Pack\s*Style\s*[:\-]\s*([^\n]+)",
            ],
            "Number of Plants": [
                r"Number\s*of\s*Plants\s*[:\-]\s*([^\n]+)",
                r"(?:Qty|Quantity)\s*(?:of\s*Plants)?\s*[:\-]\s*([^\n]+)",
            ],
            "Lot Code": [
                r"Lot\s*ID\s*[:\-]\s*([A-Za-z0-9\-]+)",
                r"Lot\s*#\s*[:\-]?\s*([A-Za-z0-9\-]+)",
                r"Lot\s*Code\s*[:\-]\s*([A-Za-z0-9\-]+)",
                r"Lot\s*Number\s*[:\-]\s*([A-Za-z0-9\-]+)",
            ],
        },
        "template_overrides": {
            "Template A": {"Lot Code": [r"Lot\s*ID\s*:\s*([A-Za-z0-9\-]+)"]},
            "Template B": {"Lot Code": [r"LOT\s*#\s*:?\s*([A-Za-z0-9\-]+)"]},
            "Template C": {"Lot Code": [r"Lot\s*(?:ID|#|Code|Number)\s*:?\s*([A-Za-z0-9\-]+)"]},
        },
        "template_b_label_terms": {
            "Date": ["shipping date"],
            "Sales Order Number": ["sales order"],
            "Customer and Location": ["customer and address", "customer and location"],
            "Pack Style": ["packaging style", "pack style"],
            "Number of Plants": ["number of plants"],
            "Lot Code": ["lot#", "lot #"],
        },
        "template_b_grid": {
            "bright_threshold": 100,
            "horizontal_coverage": 0.75,
            "vertical_coverage": 0.70,
        },
    },
    "admin": {
        "salt": "",
        "password_hash": "",
        "iterations": PBKDF2_ITERATIONS,
    },
}


def app_data_dir() -> Path:
    """Return the per-user data directory and migrate prior test/legacy folders once."""
    root = Path(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"))
    path = root / "AllisonsDataExtractor"
    legacy_candidates = [
        root / "DocuFlow",
        root / "AllisonsPDFExtractor",
    ]

    # Preserve settings/history from earlier development names when practical.
    if not path.exists():
        for legacy in legacy_candidates:
            if not legacy.exists():
                continue
            try:
                shutil.copytree(legacy, path, dirs_exist_ok=True)
                break
            except Exception:
                # Migration problems should never prevent the application from starting.
                pass

    path.mkdir(parents=True, exist_ok=True)
    return path


def settings_path() -> Path:
    return app_data_dir() / "settings.json"


def config_backup_dir() -> Path:
    path = app_data_dir() / "config_backups"
    path.mkdir(parents=True, exist_ok=True)
    return path


def update_work_dir() -> Path:
    path = app_data_dir() / "updates"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _deep_merge(defaults, loaded):
    if isinstance(defaults, dict) and isinstance(loaded, dict):
        result = copy.deepcopy(defaults)
        for key, value in loaded.items():
            if key in result:
                result[key] = _deep_merge(result[key], value)
            else:
                result[key] = copy.deepcopy(value)
        return result
    return copy.deepcopy(loaded)


def load_settings() -> dict:
    path = settings_path()
    if not path.exists():
        return copy.deepcopy(DEFAULT_SETTINGS)
    try:
        with path.open("r", encoding="utf-8") as handle:
            loaded = json.load(handle)
        if not isinstance(loaded, dict):
            raise ValueError("Settings root must be a JSON object")
        return _deep_merge(DEFAULT_SETTINGS, loaded)
    except Exception:
        # Preserve a damaged file for diagnosis and fall back safely.
        try:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(path, config_backup_dir() / f"settings_corrupt_{stamp}.json")
        except Exception:
            pass
        return copy.deepcopy(DEFAULT_SETTINGS)


def _write_settings(settings: dict) -> None:
    path = settings_path()
    temp = path.with_suffix(".json.tmp")
    with temp.open("w", encoding="utf-8") as handle:
        json.dump(settings, handle, indent=2, ensure_ascii=False)
    os.replace(temp, path)


def backup_settings() -> str:
    path = settings_path()
    if not path.exists():
        return ""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destination = config_backup_dir() / f"settings_{stamp}.json"
    shutil.copy2(path, destination)
    return str(destination)


def save_settings(settings: dict, create_backup: bool = True) -> str:
    current = load_settings()
    merged = _deep_merge(DEFAULT_SETTINGS, settings)
    # Never accidentally erase an existing admin hash because a partial config
    # was imported or edited.
    if current.get("admin", {}).get("password_hash") and not merged.get("admin", {}).get("password_hash"):
        merged["admin"] = copy.deepcopy(current["admin"])
    backup = backup_settings() if create_backup else ""
    _write_settings(merged)
    return backup


def get_setting(*path, default=None):
    value = load_settings()
    for part in path:
        if not isinstance(value, dict) or part not in value:
            return default
        value = value[part]
    return value


def set_setting(path: tuple[str, ...], value, create_backup: bool = False) -> None:
    settings = load_settings()
    target = settings
    for part in path[:-1]:
        target = target.setdefault(part, {})
    target[path[-1]] = value
    save_settings(settings, create_backup=create_backup)


def get_template_markers() -> dict:
    settings = load_settings()
    return {
        name: list(config.get("markers", []))
        for name, config in settings.get("templates", {}).items()
    }


def get_template_settings() -> dict:
    settings = load_settings()
    result = {}
    for name, config in settings.get("templates", {}).items():
        item = dict(config)
        item.pop("markers", None)
        result[name] = item
    return result


def admin_is_configured() -> bool:
    admin = load_settings().get("admin", {})
    return bool(admin.get("salt") and admin.get("password_hash"))


def _derive_password_hash(passcode: str, salt: bytes, iterations: int) -> bytes:
    return hashlib.pbkdf2_hmac(
        "sha256",
        passcode.encode("utf-8"),
        salt,
        int(iterations),
        dklen=32,
    )


def set_admin_passcode(passcode: str) -> None:
    if len(passcode) < 6:
        raise ValueError("Administrator passcode must be at least 6 characters.")
    settings = load_settings()
    salt = secrets.token_bytes(16)
    digest = _derive_password_hash(passcode, salt, PBKDF2_ITERATIONS)
    settings["admin"] = {
        "salt": salt.hex(),
        "password_hash": digest.hex(),
        "iterations": PBKDF2_ITERATIONS,
    }
    save_settings(settings, create_backup=True)


def verify_admin_passcode(passcode: str) -> bool:
    admin = load_settings().get("admin", {})
    if not admin.get("salt") or not admin.get("password_hash"):
        return False
    try:
        salt = bytes.fromhex(admin["salt"])
        expected = bytes.fromhex(admin["password_hash"])
        iterations = int(admin.get("iterations", PBKDF2_ITERATIONS))
    except Exception:
        return False
    actual = _derive_password_hash(passcode, salt, iterations)
    return hmac.compare_digest(actual, expected)


def export_configuration(path: str) -> None:
    settings = load_settings()
    # Admin credentials intentionally never leave the machine in config exports.
    settings["admin"] = {"salt": "", "password_hash": "", "iterations": PBKDF2_ITERATIONS}
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(settings, handle, indent=2, ensure_ascii=False)


def import_configuration(path: str) -> str:
    with open(path, "r", encoding="utf-8") as handle:
        incoming = json.load(handle)
    if not isinstance(incoming, dict):
        raise ValueError("Configuration file must contain a JSON object.")
    current = load_settings()
    imported = _deep_merge(DEFAULT_SETTINGS, incoming)
    imported["admin"] = copy.deepcopy(current.get("admin", DEFAULT_SETTINGS["admin"]))
    return save_settings(imported, create_backup=True)


def reset_to_defaults(preserve_admin: bool = True) -> str:
    current = load_settings()
    settings = copy.deepcopy(DEFAULT_SETTINGS)
    if preserve_admin:
        settings["admin"] = copy.deepcopy(current.get("admin", DEFAULT_SETTINGS["admin"]))
    return save_settings(settings, create_backup=True)
