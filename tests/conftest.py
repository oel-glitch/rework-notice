"""
Pytest конфигурация и фикстуры
"""
import pytest
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

@pytest.fixture
def test_data_dir():
    """Путь к тестовым данным"""
    return Path(__file__).parent / "fixtures"

@pytest.fixture
def temp_output_dir(tmp_path):
    """Временная директория для выходных файлов"""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return output_dir

@pytest.fixture
def temp_db_path(tmp_path):
    """Временная БД для тестов"""
    return tmp_path / "test.db"
