"""Shared test setup.

The app keeps settings and history under %LOCALAPPDATA%. Point that at a
throwaway folder *before* the app modules are imported (pdf_extractor reads
settings at import time), then give every test its own fresh folder.
"""

import os
import tempfile

os.environ["LOCALAPPDATA"] = tempfile.mkdtemp(prefix="ade-tests-")

import pytest

import settings_manager


@pytest.fixture(autouse=True)
def isolated_app_data(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "localappdata"))
    settings_manager._reset_caches()
    yield tmp_path / "localappdata" / "AllisonsDataExtractor"
    settings_manager._reset_caches()


SAMPLE_RECORD = {
    "Date": "August 3, 2026",
    "Sales Order Number": "123456",
    "Customer and Location": "Acme Farms, 1 Main St, Salinas, CA",
    "Pack Style": "12x1ct Clam",
    "Number of Plants": "240 heads (12 heads x 20 cases)",
    "Lot Code": "L-77",
}


@pytest.fixture
def sample_record():
    return dict(SAMPLE_RECORD)
