import pytest
from pathlib import Path
import pandas as pd
import openpyxl
from mcp_excel.models import SheetOverride

pytestmark = pytest.mark.unit


def test_corrupted_excel_file(temp_dir, loader):
    file_path = temp_dir / "corrupted.xlsx"
    file_path.write_bytes(b"this is not a valid excel file")

    with pytest.raises(RuntimeError, match="Failed to load"):
        loader.load_sheet(file_path, "corrupted.xlsx", "Sheet1", "test", None)


def test_empty_excel_file(temp_dir, loader):
    file_path = temp_dir / "empty.xlsx"
    df = pd.DataFrame()
    df.to_excel(file_path, index=False)

    with pytest.raises(RuntimeError, match="No rows found in xlsx file"):
        loader.load_sheet(file_path, "empty.xlsx", "Sheet1", "test", None)


def test_zero_row_excel_file_raw_mode(temp_dir, loader):
    file_path = temp_dir / "zero_rows.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Name", "Age", "City"])
    wb.save(file_path)

    metas = loader.load_sheet(file_path, "zero_rows.xlsx", "Sheet", "test", None)
    meta = metas[0]

    result = loader.conn.execute(f'SELECT * FROM "{meta.table_name}"').fetchall()
    assert len(result) == 1
    assert result[0] == ("Name", "Age", "City")


def test_file_with_only_headers(temp_dir, loader):
    file_path = temp_dir / "only_headers.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Name", "Age", "City"])
    ws.append(["Alice", "25", "NYC"])
    wb.save(file_path)

    override = SheetOverride(header_rows=1)
    metas = loader.load_sheet(file_path, "only_headers.xlsx", "Sheet", "test", override)
    meta = metas[0]
    assert meta.mode == "ASSISTED"

    result = loader.conn.execute(f'SELECT * FROM "{meta.table_name}"').fetchall()
    assert len(result) == 1
    assert result[0][0] == "Alice"
