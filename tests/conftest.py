"""
Shared pytest fixtures for the PO Translation System tests.

The environment is prepared before any app module is imported, so the tests
use the offline mock translator and never write to the real upload folders.
"""

import os
import tempfile
from pathlib import Path

_TEST_STORAGE = Path(tempfile.mkdtemp(prefix="po_translation_tests_"))
os.environ["TRANSLATION_PROVIDER"] = "mock"
os.environ["UPLOAD_DIR"] = str(_TEST_STORAGE / "uploads")
os.environ["OUTPUT_DIR"] = str(_TEST_STORAGE / "outputs")

import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.services.excel_service import ExcelService


@pytest.fixture
def excel_service() -> ExcelService:
    return ExcelService()


@pytest.fixture
def sample_po(tmp_path, excel_service) -> Path:
    """Path of a freshly generated sample purchase order workbook."""
    file_path = tmp_path / "Sample PO.xlsx"
    excel_service.create_sample_po(file_path)
    return file_path


@pytest.fixture
def client():
    with TestClient(app) as test_client:
        yield test_client
