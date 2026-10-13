import json
import os

import settings_manager
from settings_manager import load_settings, save_settings, set_setting, settings_path


def test_defaults_when_no_file():
    assert load_settings()["ocr"]["verified_threshold"] == 95


def test_callers_get_independent_copies():
    first = load_settings()
    first["ocr"]["verified_threshold"] = 1
    assert load_settings()["ocr"]["verified_threshold"] == 95


def test_save_is_visible_immediately():
    set_setting(("ocr", "review_threshold"), 55)
    assert load_settings()["ocr"]["review_threshold"] == 55


def test_external_file_change_is_picked_up():
    save_settings(load_settings(), create_backup=False)
    assert load_settings()["ocr"]["template_b_penalty"] == 10

    path = settings_path()
    data = json.loads(path.read_text(encoding="utf-8"))
    data["ocr"]["template_b_penalty"] = 33
    path.write_text(json.dumps(data), encoding="utf-8")
    stat = path.stat()
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000))

    assert load_settings()["ocr"]["template_b_penalty"] == 33


def test_file_is_not_reparsed_when_unchanged(monkeypatch):
    save_settings(load_settings(), create_backup=False)
    load_settings()
    calls = []
    original = settings_manager._read_settings_file
    monkeypatch.setattr(settings_manager, "_read_settings_file", lambda p: calls.append(p) or original(p))
    for _ in range(5):
        load_settings()
    assert calls == []


def test_corrupt_file_falls_back_to_defaults_and_is_preserved(isolated_app_data):
    path = settings_path()
    path.write_text("{ not json", encoding="utf-8")
    assert load_settings()["ocr"]["verified_threshold"] == 95
    load_settings()
    backups = list((isolated_app_data / "config_backups").glob("settings_corrupt_*.json"))
    assert len(backups) == 1
