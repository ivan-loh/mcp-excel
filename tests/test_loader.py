import pytest
from pathlib import Path
import pandas as pd
from mcp_excel.types import SheetOverride

pytestmark = pytest.mark.unit


def test_load_raw_mode(loader, sample_excel):
    meta = loader.load_sheet(sample_excel, "test.xlsx", "Data", "excel", None)
    assert meta.mode == "RAW"
    assert meta.table_name
    assert meta.est_rows > 0


def test_load_assisted_mode_basic(loader, sample_excel):
    override = SheetOverride(header_rows=1)
    meta = loader.load_sheet(sample_excel, "test.xlsx", "Data", "excel", override)
    assert meta.mode == "ASSISTED"
    assert meta.est_rows > 0


def test_skip_rows(temp_dir, loader):
    file_path = temp_dir / "skip_test.xlsx"
    df = pd.DataFrame({
        "Header": ["Skip", "Skip", "Name", "Alice", "Bob"],
        "Col2": ["Skip", "Skip", "Age", "25", "30"]
    })
    df.to_excel(file_path, sheet_name="Data", index=False, header=False)

    override = SheetOverride(skip_rows=2, header_rows=1)
    meta = loader.load_sheet(file_path, "skip_test.xlsx", "Data", "excel", override)
    assert meta.mode == "ASSISTED"


def test_skip_footer(temp_dir, loader):
    file_path = temp_dir / "footer_test.xlsx"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob", "Total", "Notes"],
        "Age": ["25", "30", "55", "End"]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)

    override = SheetOverride(skip_footer=2, header_rows=1)
    meta = loader.load_sheet(file_path, "footer_test.xlsx", "Data", "excel", override)
    assert meta.est_rows == 2


def test_drop_regex(temp_dir, loader):
    file_path = temp_dir / "regex_test.xlsx"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob", "Total:", "Notes:"],
        "Value": [100, 200, 300, 0]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)

    override = SheetOverride(drop_regex="^(Total|Notes):", header_rows=1)
    meta = loader.load_sheet(file_path, "regex_test.xlsx", "Data", "excel", override)
    assert meta.est_rows == 2


def test_column_renames(temp_dir, loader):
    file_path = temp_dir / "rename_test.xlsx"
    df = pd.DataFrame({
        "OldName": ["A", "B"],
        "AnotherOld": [1, 2]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)

    override = SheetOverride(column_renames={"OldName": "NewName"}, header_rows=1)
    meta = loader.load_sheet(file_path, "rename_test.xlsx", "Data", "excel", override)

    result = loader.conn.execute(f'DESCRIBE "{meta.table_name}"').fetchall()
    column_names = [row[0] for row in result]
    assert "NewName" in column_names or "newname" in column_names


def test_get_sheet_names(loader, sample_excel):
    sheets = loader.get_sheet_names(sample_excel)
    assert "Data" in sheets
    assert len(sheets) >= 1


def test_multirow_headers(temp_dir, loader):
    file_path = temp_dir / "multirow_test.xlsx"
    data = [
        ["Region", "Q1", "Q1", "Q2", "Q2"],
        ["", "Sales", "Units", "Sales", "Units"],
        ["North", 1000, 50, 1200, 60],
        ["South", 800, 40, 900, 45]
    ]
    df = pd.DataFrame(data)
    df.to_excel(file_path, sheet_name="Data", index=False, header=False)

    override = SheetOverride(header_rows=2)
    meta = loader.load_sheet(file_path, "multirow_test.xlsx", "Data", "excel", override)

    result = loader.conn.execute(f'SELECT * FROM "{meta.table_name}" LIMIT 1').fetchall()
    assert len(result) > 0
    assert meta.est_rows == 2


def test_type_hints(temp_dir, loader):
    file_path = temp_dir / "types_test.xlsx"
    df = pd.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": ["25", "30"],
        "Salary": ["50000.50", "60000.75"],
        "Active": ["true", "false"]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)

    override = SheetOverride(
        header_rows=1,
        type_hints={
            "Age": "INT",
            "Salary": "DECIMAL",
            "Active": "BOOL"
        }
    )
    meta = loader.load_sheet(file_path, "types_test.xlsx", "Data", "excel", override)

    result = loader.conn.execute(f'SELECT * FROM "{meta.table_name}" LIMIT 1').fetchall()
    assert len(result) > 0


def test_unpivot(temp_dir, loader):
    file_path = temp_dir / "unpivot_test.xlsx"
    df = pd.DataFrame({
        "Region": ["North", "South"],
        "Jan": [100, 80],
        "Feb": [110, 85],
        "Mar": [120, 90]
    })
    df.to_excel(file_path, sheet_name="Data", index=False)

    override = SheetOverride(
        header_rows=1,
        unpivot={
            "id_vars": ["Region"],
            "value_vars": ["Jan", "Feb", "Mar"],
            "var_name": "Month",
            "value_name": "Sales"
        }
    )
    meta = loader.load_sheet(file_path, "unpivot_test.xlsx", "Data", "excel", override)

    result = loader.conn.execute(f'SELECT COUNT(*) FROM "{meta.table_name}"').fetchone()
    assert result[0] == 6
