import pytest
import tempfile
from pathlib import Path
import pandas as pd
import src.server as server


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


@pytest.fixture
def test_data_dir():
    with tempfile.TemporaryDirectory() as tmpdir:
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


def test_load_multiple_files(test_data_dir):
    result = server.load_dir(path=str(test_data_dir), alias="test")
    assert result["files_count"] == 3
    assert result["sheets_count"] == 3
    assert result["tables_count"] == 3


def test_query_across_tables(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    result = server.query("SELECT COUNT(*) as total FROM test__sales_0__summary")
    assert result["row_count"] == 1
    assert result["truncated"] is False
    assert result["execution_ms"] >= 0


def test_query_safety_reject_ddl(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    with pytest.raises(ValueError, match="Only SELECT queries allowed"):
        server.query("DROP TABLE test__sales_0__summary")


def test_query_safety_reject_dml(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    with pytest.raises(ValueError, match="Only SELECT queries allowed"):
        server.query("INSERT INTO test__sales_0__summary VALUES (1,2,3)")


def test_query_row_limit(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    result = server.query("SELECT * FROM test__sales_0__summary", max_rows=2)
    assert result["row_count"] == 2
    assert result["truncated"] is True


def test_list_tables_all(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    result = server.list_tables()
    assert len(result["tables"]) == 3


def test_list_tables_filtered(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test1")
    server.load_dir(path=str(test_data_dir), alias="test2")
    result = server.list_tables(alias="test1")
    assert all(t["table"].startswith("test1__") for t in result["tables"])


def test_get_schema(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]
    result = server.get_schema(table_name)
    assert "columns" in result
    assert len(result["columns"]) > 0


def test_refresh_incremental(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    result = server.refresh(alias="test", full=False)
    assert "changed" in result
    assert "total" in result


def test_refresh_full(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")
    result = server.refresh(alias="test", full=True)
    assert "dropped" in result
    assert "added" in result


def test_assisted_mode_with_overrides(test_data_dir):
    file_path = test_data_dir / "override_test.xlsx"
    df = pd.DataFrame({
        "Header": ["Skip", "Name", "Alice", "Bob", "Total:"],
        "Value": ["Skip", "Age", "25", "30", "55"]
    })
    df.to_excel(file_path, sheet_name="Data", index=False, header=False)

    overrides = {
        "override_test.xlsx": {
            "sheet_overrides": {
                "Data": {
                    "skip_rows": 1,
                    "skip_footer": 1,
                    "header_rows": 1
                }
            }
        }
    }

    result = server.load_dir(path=str(test_data_dir), alias="test", overrides=overrides)
    assert result["sheets_count"] > 0


def test_path_validation_nonexistent():
    with pytest.raises(ValueError, match="does not exist"):
        server.load_dir(path="/nonexistent/path", alias="test")


def test_path_validation_not_directory(test_data_dir):
    file_path = test_data_dir / "test.xlsx"
    pd.DataFrame({"A": [1]}).to_excel(file_path, index=False)

    with pytest.raises(ValueError, match="not a directory"):
        server.load_dir(path=str(file_path), alias="test")


def test_system_views(test_data_dir):
    server.load_dir(path=str(test_data_dir), alias="test")

    files_result = server.query("SELECT * FROM test____files")
    assert files_result["row_count"] == 3

    tables_result = server.query("SELECT * FROM test____tables")
    assert tables_result["row_count"] == 3

    tables_with_sheets = server.query("SELECT sheet_name FROM test____tables WHERE sheet_name = 'Summary'")
    assert tables_with_sheets["row_count"] == 3
