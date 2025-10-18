import pytest
import tempfile
from pathlib import Path
import pandas as pd
import re
import mcp_excel.server as server

pytestmark = [pytest.mark.integration, pytest.mark.usefixtures("setup_server")]


def get_sanitized_alias(path: Path) -> str:
    alias = path.name or "excel"
    alias = alias.lower()
    alias = alias.replace(' ', '_')
    alias = re.sub(r'[^a-z0-9_$]', '', alias)
    alias = re.sub(r'_+', '_', alias)
    alias = alias.strip('_')
    return alias if alias else "excel"


def test_load_multiple_files(test_data_dir):
    result = server.load_dir(path=str(test_data_dir))
    alias = get_sanitized_alias(Path(test_data_dir))
    assert result["files_count"] == 3
    assert result["sheets_count"] == 3
    assert result["tables_count"] == 3


def test_query_across_tables(test_data_dir):
    server.load_dir(path=str(test_data_dir))
    alias = get_sanitized_alias(Path(test_data_dir))
    result = server.query(f'SELECT COUNT(*) as total FROM "{alias}.sales_0.summary"')
    assert result["row_count"] == 1
    assert result["truncated"] is False
    assert result["execution_ms"] >= 0


def test_query_safety_reject_ddl(test_data_dir):
    server.load_dir(path=str(test_data_dir))
    alias = get_sanitized_alias(Path(test_data_dir))
    with pytest.raises(RuntimeError, match="Query failed"):
        server.query(f'DROP VIEW "{alias}.sales_0.summary"')


def test_query_safety_reject_dml(test_data_dir):
    server.load_dir(path=str(test_data_dir))
    alias = get_sanitized_alias(Path(test_data_dir))
    with pytest.raises(RuntimeError, match="Query failed"):
        server.query(f'INSERT INTO "{alias}.sales_0.summary" VALUES (1,2,3)')


def test_query_row_limit(test_data_dir):
    server.load_dir(path=str(test_data_dir))
    alias = get_sanitized_alias(Path(test_data_dir))
    result = server.query(f'SELECT * FROM "{alias}.sales_0.summary"', max_rows=2)
    assert result["row_count"] == 2
    assert result["truncated"] is True


def test_list_tables_all(test_data_dir):
    server.load_dir(path=str(test_data_dir))
    result = server.list_tables()
    assert len(result["tables"]) == 3


def test_list_tables_filtered(test_data_dir):
    alias1 = get_sanitized_alias(Path(test_data_dir))
    server.load_dir(path=str(test_data_dir))

    test_data_dir2 = test_data_dir.parent / "other_data"
    test_data_dir2.mkdir(exist_ok=True)
    df = pd.DataFrame({"A": [1, 2]})
    df.to_excel(test_data_dir2 / "test.xlsx", index=False)
    alias2 = get_sanitized_alias(Path(test_data_dir2))
    server.load_dir(path=str(test_data_dir2))

    result = server.list_tables(alias=alias1)
    assert all(t["table"].startswith(f"{alias1}.") for t in result["tables"])


def test_get_schema(test_data_dir):
    server.load_dir(path=str(test_data_dir))
    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]
    result = server.get_schema(table_name)
    assert "columns" in result
    assert len(result["columns"]) > 0


def test_refresh_incremental(test_data_dir):
    alias = get_sanitized_alias(Path(test_data_dir))
    server.load_dir(path=str(test_data_dir))
    result = server.refresh(alias=alias, full=False)
    assert "changed" in result
    assert "total" in result


def test_refresh_full(test_data_dir):
    alias = get_sanitized_alias(Path(test_data_dir))
    server.load_dir(path=str(test_data_dir))
    result = server.refresh(alias=alias, full=True)
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

    result = server.load_dir(path=str(test_data_dir), overrides=overrides)
    assert result["sheets_count"] > 0


def test_path_validation_nonexistent():
    with pytest.raises(ValueError, match="does not exist"):
        server.load_dir(path="/nonexistent/path")


def test_path_validation_not_directory(test_data_dir):
    file_path = test_data_dir / "test.xlsx"
    pd.DataFrame({"A": [1]}).to_excel(file_path, index=False)

    with pytest.raises(ValueError, match="not a directory"):
        server.load_dir(path=str(file_path))


def test_system_views(test_data_dir):
    alias = get_sanitized_alias(Path(test_data_dir))
    server.load_dir(path=str(test_data_dir))

    files_result = server.query(f'SELECT * FROM "{alias}.__files"')
    assert files_result["row_count"] == 3

    tables_result = server.query(f'SELECT * FROM "{alias}.__tables"')
    assert tables_result["row_count"] == 3

    tables_with_sheets = server.query(f'SELECT sheet_name FROM "{alias}.__tables" WHERE sheet_name = \'Summary\'')
    assert tables_with_sheets["row_count"] == 3


def test_query_timeout_enforcement(test_data_dir):
    server.load_dir(path=str(test_data_dir))

    slow_query = """
        WITH RECURSIVE slow AS (
            SELECT 1 as n
            UNION ALL
            SELECT n + 1 FROM slow WHERE n < 100000000
        )
        SELECT COUNT(*) FROM slow
    """

    with pytest.raises(TimeoutError, match="exceeded.*timeout"):
        server.query(slow_query, timeout_ms=100)


def test_refresh_handles_missing_file(test_data_dir):
    import os

    alias = get_sanitized_alias(Path(test_data_dir))
    server.load_dir(path=str(test_data_dir))
    initial_total = len(server.catalog)

    file_to_delete = test_data_dir / "sales_0.xlsx"
    os.remove(str(file_to_delete))

    result = server.refresh(alias=alias, full=False)

    assert "changed" in result
    assert "total" in result
    assert result["total"] == initial_total


def test_refresh_handles_moved_file(test_data_dir):
    import shutil

    alias = get_sanitized_alias(Path(test_data_dir))
    server.load_dir(path=str(test_data_dir))
    initial_total = len(server.catalog)

    moved_dir = test_data_dir.parent / "moved_files"
    moved_dir.mkdir(exist_ok=True)

    source_file = test_data_dir / "sales_1.xlsx"
    dest_file = moved_dir / "sales_1.xlsx"

    if source_file.exists():
        shutil.move(str(source_file), str(dest_file))

        result = server.refresh(alias=alias, full=False)

        assert "changed" in result
        assert "total" in result
        assert result["changed"] == 0

        shutil.rmtree(str(moved_dir))


def test_load_with_current_directory():
    import os
    original_cwd = os.getcwd()

    with tempfile.TemporaryDirectory(prefix="testdir_") as tmpdir:
        tmpdir = Path(tmpdir)

        df = pd.DataFrame({
            "Product": ["A", "B", "C"],
            "Price": [10, 20, 30]
        })
        df.to_excel(tmpdir / "test.xlsx", sheet_name="Data", index=False)

        os.chdir(tmpdir)

        try:
            alias = get_sanitized_alias(Path.cwd())
            result = server.load_dir(path=".")

            assert result["files_count"] == 1
            assert result["sheets_count"] == 1
            assert result["tables_count"] == 1

            tables = server.list_tables(alias=alias)
            assert len(tables["tables"]) == 1
        finally:
            os.chdir(original_cwd)
