import pytest
import tempfile
import re
from pathlib import Path
import pandas as pd
import duckdb
import mcp_excel.server as server
from mcp_excel.loading.loader import ExcelLoader
from mcp_excel.utils.naming import TableRegistry


@pytest.fixture
def temp_dir():
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def temp_excel_dir():
    with tempfile.TemporaryDirectory(prefix="test_excel_") as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def sample_excel(temp_dir):
    file_path = temp_dir / "test.xlsx"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie"],
        "Age": [25, 30, 35],
        "City": ["NYC", "LA", "SF"]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)
    return file_path


@pytest.fixture
def conn():
    connection = duckdb.connect(":memory:")
    yield connection
    connection.close()


@pytest.fixture
def loader(conn):
    registry = TableRegistry()
    return ExcelLoader(conn, registry)


@pytest.fixture
def setup_server():
    server.conn = None
    server.registry = None
    server.loader = None
    server.catalog.clear()
    server.load_configs.clear()
    server.init_server()
    yield
    server.catalog.clear()
    server.load_configs.clear()


@pytest.fixture
def setup_server_http():
    server.conn = None
    server.registry = None
    server.loader = None
    server.catalog.clear()
    server.load_configs.clear()
    server._db_path = None
    server._use_http_mode = False
    server.init_server(use_http_mode=True)
    yield
    server.catalog.clear()
    server.load_configs.clear()
    if server._db_path and server._db_path != ":memory:":
        try:
            Path(server._db_path).unlink(missing_ok=True)
        except:
            pass


@pytest.fixture
def test_data_dir():
    with tempfile.TemporaryDirectory(prefix="test_") as tmpdir:
        tmpdir = Path(tmpdir)
        for i in range(3):
            df = pd.DataFrame({
                "Product": [f"Product{j}" for j in range(5)],
                "Quantity": [10 * (i + 1) + j for j in range(5)],
                "Price": [100.0 + i * 10 + j for j in range(5)]
            })
            file_path = tmpdir / f"sales_{i}.xlsx"
            df.to_excel(file_path, sheet_name="Summary", index=False)
        yield tmpdir


def get_sanitized_alias(path: Path) -> str:
    alias = path.name or "excel"
    alias = alias.lower()
    alias = alias.replace(' ', '_')
    alias = re.sub(r'[^a-z0-9_$]', '', alias)
    alias = re.sub(r'_+', '_', alias)
    alias = alias.strip('_')
    return alias if alias else "excel"


def pytest_configure(config):
    config.addinivalue_line("markers", "unit: Unit tests for individual components")
    config.addinivalue_line("markers", "integration: Integration tests for end-to-end workflows")
    config.addinivalue_line("markers", "regression: Regression tests for fixed bugs")
    config.addinivalue_line("markers", "concurrency: Concurrency and thread safety tests")
    config.addinivalue_line("markers", "stress: Stress and performance tests")
