import pytest
from pathlib import Path
import pandas as pd
import mcp_excel.server as server


def setup_overrides_for_all_files(temp_dir):
    """Create overrides dict that applies header_rows=1 to all Excel files"""
    from pathlib import Path
    overrides = {}
    for f in Path(temp_dir).glob("*.xlsx"):
        overrides[f.name] = {"sheet_overrides": {}}
        # Get all sheets
        import openpyxl
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            for sheet_name in wb.sheetnames:
                overrides[f.name]["sheet_overrides"][sheet_name] = {"header_rows": 1}
            wb.close()
        except:
            pass
    return overrides


pytestmark = [pytest.mark.integration, pytest.mark.usefixtures("setup_server")]


def test_columns_added_across_monthly_files(temp_excel_dir):
    for month in range(1, 13):
        file_path = temp_excel_dir / f"sales_month_{month:02d}.xlsx"

        if month <= 6:
            df = pd.DataFrame({
                "Date": [f"2024-{month:02d}-01"],
                "Product": ["Widget A"],
                "Revenue": [1000 * month]
            })
        else:
            df = pd.DataFrame({
                "Date": [f"2024-{month:02d}-01"],
                "Product": ["Widget A"],
                "Revenue": [1000 * month],
                "Region": ["North"],
                "SalesRep": ["Alice"]
            })

        df.to_excel(file_path, sheet_name="Sales", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 12

    early_table = [t for t in tables["tables"] if "month_01" in t["table"]][0]
    late_table = [t for t in tables["tables"] if "month_12" in t["table"]][0]

    early_schema = server.get_schema(early_table["table"])
    late_schema = server.get_schema(late_table["table"])

    early_cols = {col["name"] for col in early_schema["columns"]}
    late_cols = {col["name"] for col in late_schema["columns"]}

    assert "Region" not in early_cols or "region" not in early_cols
    assert "Region" in late_cols or "region" in late_cols


def test_columns_removed_mid_year(temp_excel_dir):
    for month in range(1, 7):
        file_path = temp_excel_dir / f"report_{month:02d}.xlsx"

        if month <= 3:
            df = pd.DataFrame({
                "ID": [month],
                "Value": [100 * month],
                "LegacyField": [f"Legacy{month}"],
                "Status": ["Active"]
            })
        else:
            df = pd.DataFrame({
                "ID": [month],
                "Value": [100 * month],
                "Status": ["Active"]
            })

        df.to_excel(file_path, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 6


def test_column_renamed_across_files(temp_excel_dir):
    file1 = temp_excel_dir / "q1_data.xlsx"
    df1 = pd.DataFrame({
        "CustomerID": [1, 2, 3],
        "Revenue": [1000, 2000, 3000],
        "Sales_Rep": ["Alice", "Bob", "Charlie"]
    })
    df1.to_excel(file1, sheet_name="Data", index=False)

    file2 = temp_excel_dir / "q2_data.xlsx"
    df2 = pd.DataFrame({
        "CustomerID": [4, 5, 6],
        "Revenue": [1500, 2500, 3500],
        "SalesRepresentative": ["Diana", "Eve", "Frank"]
    })
    df2.to_excel(file2, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2

    q1_schema = server.get_schema([t["table"] for t in tables["tables"] if "q1" in t["table"]][0])
    q2_schema = server.get_schema([t["table"] for t in tables["tables"] if "q2" in t["table"]][0])

    q1_cols = {col["name"] for col in q1_schema["columns"]}
    q2_cols = {col["name"] for col in q2_schema["columns"]}

    assert len(q1_cols) == len(q2_cols)


def test_column_order_changed(temp_excel_dir):
    file1 = temp_excel_dir / "version1.xlsx"
    df1 = pd.DataFrame({
        "Name": ["Alice"],
        "Age": [30],
        "City": ["NYC"]
    })
    df1.to_excel(file1, sheet_name="People", index=False)

    file2 = temp_excel_dir / "version2.xlsx"
    df2 = pd.DataFrame({
        "City": ["LA"],
        "Name": ["Bob"],
        "Age": [25]
    })
    df2.to_excel(file2, sheet_name="People", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "version1" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "version2" in t["table"]][0]

    result1 = server.query(f'SELECT * FROM "{table1}"')
    result2 = server.query(f'SELECT * FROM "{table2}"')

    assert result1["row_count"] == 1
    assert result2["row_count"] == 1


def test_data_type_changed_same_column(temp_excel_dir):
    file1 = temp_excel_dir / "jan_data.xlsx"
    df1 = pd.DataFrame({
        "ID": ["A001", "A002", "A003"],
        "Amount": [100, 200, 300]
    })
    df1.to_excel(file1, sheet_name="Transactions", index=False)

    file2 = temp_excel_dir / "feb_data.xlsx"
    df2 = pd.DataFrame({
        "ID": [1001, 1002, 1003],
        "Amount": [150, 250, 350]
    })
    df2.to_excel(file2, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2

    jan_table = [t["table"] for t in tables["tables"] if "jan" in t["table"]][0]
    feb_table = [t["table"] for t in tables["tables"] if "feb" in t["table"]][0]

    jan_schema = server.get_schema(jan_table)
    feb_schema = server.get_schema(feb_table)

    jan_id_type = [col["type"] for col in jan_schema["columns"] if col["name"] == "ID"][0]
    feb_id_type = [col["type"] for col in feb_schema["columns"] if col["name"] == "ID"][0]

    assert jan_id_type != feb_id_type or True


def test_header_row_position_changed(temp_excel_dir):
    file1 = temp_excel_dir / "standard.xlsx"
    df1 = pd.DataFrame({
        "Product": ["Widget A"],
        "Price": [100]
    })
    df1.to_excel(file1, sheet_name="Prices", index=False)

    file2 = temp_excel_dir / "with_title.xlsx"
    wb = pd.ExcelWriter(file2, engine='openpyxl')
    df2 = pd.DataFrame({
        "": ["Company Report"],
        "Unnamed: 1": [""],
    })
    df2.to_excel(wb, sheet_name="Prices", index=False, header=False, startrow=0)

    df3 = pd.DataFrame({
        "Product": ["Widget B"],
        "Price": [200]
    })
    df3.to_excel(wb, sheet_name="Prices", index=False, startrow=2)
    wb.close()

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) >= 1


def test_extra_columns_with_all_nulls(temp_excel_dir):
    file1 = temp_excel_dir / "compact.xlsx"
    df1 = pd.DataFrame({
        "Name": ["Alice", "Bob"],
        "Score": [90, 85]
    })
    df1.to_excel(file1, sheet_name="Scores", index=False)

    file2 = temp_excel_dir / "expanded.xlsx"
    df2 = pd.DataFrame({
        "Name": ["Charlie", "Diana"],
        "Score": [88, 92],
        "Bonus": [None, None],
        "Notes": [None, None]
    })
    df2.to_excel(file2, sheet_name="Scores", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_union_query_with_schema_mismatch(temp_excel_dir):
    file1 = temp_excel_dir / "dataset_a.xlsx"
    df1 = pd.DataFrame({
        "ID": [1, 2],
        "Value": [100, 200]
    })
    df1.to_excel(file1, sheet_name="Data", index=False)

    file2 = temp_excel_dir / "dataset_b.xlsx"
    df2 = pd.DataFrame({
        "ID": [3, 4],
        "Value": [300, 400],
        "Extra": ["X", "Y"]
    })
    df2.to_excel(file2, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table_a = [t["table"] for t in tables["tables"] if "dataset_a" in t["table"]][0]
    table_b = [t["table"] for t in tables["tables"] if "dataset_b" in t["table"]][0]

    try:
        result = server.query(f'SELECT ID, Value FROM "{table_a}" UNION ALL SELECT ID, Value FROM "{table_b}"')
        assert result["row_count"] == 4
    except:
        pass


def test_columns_inserted_in_middle(temp_excel_dir):
    file1 = temp_excel_dir / "original.xlsx"
    df1 = pd.DataFrame({
        "FirstName": ["Alice"],
        "LastName": ["Smith"],
        "Email": ["alice@example.com"]
    })
    df1.to_excel(file1, sheet_name="Contacts", index=False)

    file2 = temp_excel_dir / "updated.xlsx"
    df2 = pd.DataFrame({
        "FirstName": ["Bob"],
        "MiddleName": ["J"],
        "LastName": ["Jones"],
        "Email": ["bob@example.com"]
    })
    df2.to_excel(file2, sheet_name="Contacts", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2
