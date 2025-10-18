import pytest
from pathlib import Path
import mcp_excel.server as server


@pytest.fixture(autouse=True)
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


def test_load_examples_directory():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not examples_dir.exists():
        pytest.skip("Examples directory not found")

    result = server.load_dir(path=str(examples_dir))

    assert result["files_count"] >= 7
    assert result["sheets_count"] >= 8
    assert result["tables_count"] >= 8


def test_load_clean_data_raw_mode():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not (examples_dir / "clean_data.xlsx").exists():
        pytest.skip("clean_data.xlsx not found")

    alias = examples_dir.name
    result = server.load_dir(path=str(examples_dir))

    tables = server.list_tables(alias=alias)
    clean_tables = [t for t in tables["tables"] if "clean_data" in t["table"]]

    assert len(clean_tables) >= 1
    assert any(t["mode"] == "RAW" for t in clean_tables)


def test_query_clean_data():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not (examples_dir / "clean_data.xlsx").exists():
        pytest.skip("clean_data.xlsx not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.clean_data.orders"')

    assert result["row_count"] == 1
    assert result["rows"][0][0] == 6


def test_load_with_overrides():
    examples_dir = Path(__file__).parent.parent / "examples"
    overrides_file = examples_dir / "examples_overrides.yaml"

    if not overrides_file.exists():
        pytest.skip("examples_overrides.yaml not found")

    import yaml
    with open(overrides_file) as f:
        overrides = yaml.safe_load(f)

    result = server.load_dir(path=str(examples_dir), overrides=overrides)

    assert result["files_count"] >= 6


def test_system_views_with_examples():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not examples_dir.exists():
        pytest.skip("Examples directory not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    files_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.__files"')
    assert files_result["row_count"] == 1
    assert files_result["rows"][0][0] >= 7

    tables_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.__tables"')
    assert tables_result["row_count"] == 1
    assert tables_result["rows"][0][0] >= 8
