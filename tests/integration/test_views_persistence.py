import pytest
from pathlib import Path
import pandas as pd

import mcp_excel.server as server

pytestmark = pytest.mark.integration


@pytest.fixture(autouse=True)
def reset_server():
    for config in server.load_configs.values():
        for view_file in config.root.glob(".view_*"):
            try:
                view_file.unlink()
            except:
                pass

    server.catalog.clear()
    server.views.clear()
    server.load_configs.clear()

    if server.conn:
        try:
            views_result = server.conn.execute("SELECT name FROM duckdb_views()").fetchall()
            for row in views_result:
                view_name = row[0]
                try:
                    server.conn.execute(f'DROP VIEW IF EXISTS "{view_name}"')
                except:
                    pass
        except:
            pass

    server.init_server()
    yield

    for config in server.load_configs.values():
        for view_file in config.root.glob(".view_*"):
            try:
                view_file.unlink()
            except:
                pass

    if server.conn:
        try:
            views_result = server.conn.execute("SELECT name FROM duckdb_views()").fetchall()
            for row in views_result:
                view_name = row[0]
                try:
                    server.conn.execute(f'DROP VIEW IF EXISTS "{view_name}"')
                except:
                    pass
        except:
            pass

    server.catalog.clear()
    server.views.clear()
    server.load_configs.clear()


def test_views_persist_across_restart(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1, 2, 3]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("persist1", 'SELECT 1 as val UNION SELECT 2')
    server.create_view("persist2", 'SELECT 3 as val')

    view_file1 = temp_excel_dir / ".view_persist1"
    view_file2 = temp_excel_dir / ".view_persist2"
    assert view_file1.exists()
    assert view_file2.exists()

    server.catalog.clear()
    server.views.clear()
    server.load_configs.clear()
    server.init_server()

    server.load_dir(str(temp_excel_dir))

    assert len(server.views) == 2
    assert "persist1" in server.views
    assert "persist2" in server.views

    result = server.query('SELECT * FROM "persist1"')
    assert result["row_count"] == 2

    result = server.query('SELECT * FROM "persist2"')
    assert result["row_count"] == 1


def test_views_loaded_on_startup(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"Name": ["Alice"]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))
    tables_result = server.list_tables()
    table_name = tables_result["tables"][0]["table"]

    view_file = temp_excel_dir / ".view_simple"
    view_file.write_text(f'SELECT * FROM "{table_name}"')

    server.catalog.clear()
    server.views.clear()
    server.load_configs.clear()
    server.init_server()

    server.load_dir(str(temp_excel_dir))

    assert "simple" in server.views

    result = server.query('SELECT * FROM "simple"')
    assert result["row_count"] >= 1


def test_views_survive_refresh(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"Value": [1]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("refreshview", 'SELECT 1 as x')

    server.refresh(full=True)

    assert "refreshview" in server.views

    result = server.query('SELECT * FROM "refreshview"')
    assert result["row_count"] == 1


def test_empty_view_file_skipped(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    view_file = temp_excel_dir / ".view_empty"
    view_file.write_text("")

    server.load_dir(str(temp_excel_dir))

    assert "empty" not in server.views


def test_invalid_view_file_skipped(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    view_file = temp_excel_dir / ".view_broken"
    view_file.write_text("SELECT * FROM nonexistent_table")

    server.load_dir(str(temp_excel_dir))

    assert "broken" not in server.views


def test_view_file_with_comments(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"Product": ["Widget"]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))
    tables_result = server.list_tables()
    table_name = tables_result["tables"][0]["table"]

    view_file = temp_excel_dir / ".view_commented"
    view_file.write_text(f"""-- This view shows all products
-- Created for analysis purposes
SELECT * FROM "{table_name}"
""")

    server.catalog.clear()
    server.views.clear()
    server.load_configs.clear()
    server.init_server()

    server.load_dir(str(temp_excel_dir))

    assert "commented" in server.views

    result = server.query('SELECT * FROM "commented"')
    assert result["row_count"] >= 1


def test_multiple_roots_separate_views(temp_dir):
    dir1 = temp_dir / "sales"
    dir1.mkdir()
    file1 = dir1 / "data.xlsx"
    df1 = pd.DataFrame({"A": [1]})
    df1.to_excel(file1, index=False)

    server.load_dir(str(dir1))

    server.create_view("rootview", 'SELECT 1 as x')

    view1_file = dir1 / ".view_rootview"
    assert view1_file.exists()
