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


def test_create_view_basic(temp_excel_dir):
    file_path = temp_excel_dir / "sales.xlsx"
    df = pd.DataFrame({
        "Product": ["Widget A", "Widget B", "Widget C"],
        "Amount": [100, 200, 300]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir))

    tables_result = server.list_tables()
    table_name = tables_result["tables"][0]["table"]

    schema = server.get_schema(table_name)
    col1 = schema["columns"][0]["name"]
    col2 = schema["columns"][1]["name"]

    result = server.create_view(
        view_name="high_value",
        sql=f'SELECT * FROM "{table_name}" WHERE {col1} != \'Product\' AND CAST({col2} AS INT) > 150'
    )

    assert result["view_name"] == "high_value"
    assert result["created"] is True
    assert result["est_rows"] == 2

    view_file = temp_excel_dir / ".view_high_value"
    assert view_file.exists()
    sql_content = view_file.read_text()
    assert "SELECT" in sql_content


def test_create_view_validates_name():
    with pytest.raises(ValueError, match="cannot contain dots"):
        server.create_view("my.view", "SELECT 1")

    with pytest.raises(ValueError, match="cannot start with underscore"):
        server.create_view("_myview", "SELECT 1")

    with pytest.raises(ValueError, match="letters, numbers, and underscores"):
        server.create_view("my-view", "SELECT 1")


def test_create_view_requires_select(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    with pytest.raises(ValueError, match="must be a SELECT query"):
        server.create_view("badview", "INSERT INTO x VALUES (1)")

    with pytest.raises(ValueError, match="must be a SELECT query"):
        server.create_view("badview", "DROP TABLE x")


def test_create_view_detects_conflicts(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("myview", "SELECT 1 as x")

    with pytest.raises(ValueError, match="already exists"):
        server.create_view("myview", "SELECT 2 as x")


def test_drop_view(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("dropview", "SELECT 1 as x")
    view_file = temp_excel_dir / ".view_dropview"
    assert view_file.exists()

    result = server.drop_view("dropview")

    assert result["view_name"] == "dropview"
    assert result["dropped"] is True
    assert not view_file.exists()
    assert "dropview" not in server.views


def test_drop_view_not_found():
    with pytest.raises(ValueError, match="does not exist"):
        server.drop_view("nonexistent")


def test_list_tables_includes_views(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({
        "Name": ["A", "B"],
        "Value": [1, 2]
    })
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("listview", 'SELECT 1 as TestCol')

    result = server.list_tables()

    assert "tables" in result
    assert "views" in result
    assert len(result["views"]) == 1

    view = result["views"][0]
    assert view["name"] == "listview"
    assert view["source"] == "view"
    assert "SELECT" in view["sql"]


def test_get_schema_works_on_views(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("schemaview", 'SELECT 1 as Name, 2 as Value')

    result = server.get_schema("schemaview")

    assert "columns" in result
    assert len(result["columns"]) == 2


def test_query_works_on_views(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1, 2, 3]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("numbers", 'SELECT 1 as num UNION SELECT 2 UNION SELECT 3')

    result = server.query('SELECT * FROM "numbers"')

    assert result["row_count"] == 3


def test_view_with_aggregation(temp_excel_dir):
    file_path = temp_excel_dir / "sales.xlsx"
    df = pd.DataFrame({
        "Product": ["Widget", "Widget", "Gadget"],
        "Amount": [100, 200, 150]
    })
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    tables_result = server.list_tables()
    table_name = tables_result["tables"][0]["table"]

    schema = server.get_schema(table_name)
    col1 = schema["columns"][0]["name"]
    col2 = schema["columns"][1]["name"]

    server.create_view(
        "product_totals",
        f'SELECT {col1} as Product, SUM(CAST({col2} AS INT)) as Total FROM "{table_name}" WHERE {col1} != \'Product\' GROUP BY {col1}'
    )

    result = server.query('SELECT * FROM "product_totals" ORDER BY Product')

    assert result["row_count"] == 2


def test_create_view_invalid_sql(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"A": [1]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    with pytest.raises(RuntimeError, match="Failed to create view"):
        server.create_view("badview", "SELECT * FROM nonexistent_table")


def test_multiple_views(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"Value": [1, 2, 3]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("view1", 'SELECT 1 as x')
    server.create_view("view2", 'SELECT 2 as x')
    server.create_view("view3", 'SELECT 3 as x')

    result = server.list_tables()

    assert len(result["views"]) == 3

    server.drop_view("view2")

    result = server.list_tables()
    assert len(result["views"]) == 2

    view_names = [v["name"] for v in result["views"]]
    assert "view1" in view_names
    assert "view3" in view_names
    assert "view2" not in view_names


def test_view_depends_on_another_view(temp_excel_dir):
    file_path = temp_excel_dir / "data.xlsx"
    df = pd.DataFrame({"Value": [1, 2, 3]})
    df.to_excel(file_path, index=False)

    server.load_dir(str(temp_excel_dir))

    server.create_view("nums", 'SELECT 1 as val UNION SELECT 2 UNION SELECT 3')
    server.create_view("doubled", 'SELECT val, val * 2 as doubled FROM "nums"')

    result = server.query('SELECT * FROM "doubled"')

    assert result["row_count"] == 3
