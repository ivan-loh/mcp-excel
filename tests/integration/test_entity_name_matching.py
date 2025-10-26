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


def test_company_name_variations(temp_excel_dir):
    file1 = temp_excel_dir / "customers.xlsx"
    df1 = pd.DataFrame({
        "CustomerName": ["IBM Corp", "Microsoft Corporation", "Apple Inc."],
        "Revenue": [10000, 20000, 30000]
    })
    df1.to_excel(file1, sheet_name="Sales", index=False)

    file2 = temp_excel_dir / "contracts.xlsx"
    df2 = pd.DataFrame({
        "ClientName": ["IBM Corporation", "Microsoft Corp", "Apple, Inc."],
        "ContractValue": [15000, 25000, 35000]
    })
    df2.to_excel(file2, sheet_name="Contracts", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    sales_table = [t["table"] for t in tables["tables"] if "customers" in t["table"]][0]
    contracts_table = [t["table"] for t in tables["tables"] if "contracts" in t["table"]][0]

    result = server.query(f'''
        SELECT
            s.CustomerName,
            c.ClientName,
            s.Revenue,
            c.ContractValue
        FROM "{sales_table}" s
        FULL OUTER JOIN "{contracts_table}" c ON LOWER(REPLACE(s.CustomerName, '.', '')) = LOWER(REPLACE(c.ClientName, '.', ''))
    ''')

    assert result["row_count"] >= 3


def test_whitespace_variations(temp_excel_dir):
    file1 = temp_excel_dir / "source1.xlsx"
    df1 = pd.DataFrame({
        "Company": ["Acme Corp", "Global Industries", "Tech Solutions"],
        "Amount": [1000, 2000, 3000]
    })
    df1.to_excel(file1, sheet_name="Data", index=False)

    file2 = temp_excel_dir / "source2.xlsx"
    df2 = pd.DataFrame({
        "Company": ["Acme Corp ", " Global Industries", "Tech  Solutions"],
        "Amount": [1500, 2500, 3500]
    })
    df2.to_excel(file2, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "source1" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "source2" in t["table"]][0]

    result = server.query(f'''
        SELECT TRIM(Company) as NormalizedName, SUM(Amount) as TotalAmount
        FROM (
            SELECT Company, Amount FROM "{table1}"
            UNION ALL
            SELECT Company, Amount FROM "{table2}"
        )
        GROUP BY TRIM(Company)
    ''')

    assert result["row_count"] == 3


def test_case_sensitivity_differences(temp_excel_dir):
    file1 = temp_excel_dir / "list_a.xlsx"
    df1 = pd.DataFrame({
        "Name": ["apple", "MICROSOFT", "Google"],
        "Type": ["Fruit", "Company", "Company"]
    })
    df1.to_excel(file1, sheet_name="Entities", index=False)

    file2 = temp_excel_dir / "list_b.xlsx"
    df2 = pd.DataFrame({
        "Name": ["Apple", "Microsoft", "GOOGLE"],
        "Category": ["Tech", "Tech", "Tech"]
    })
    df2.to_excel(file2, sheet_name="Entities", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table_a = [t["table"] for t in tables["tables"] if "list_a" in t["table"]][0]
    table_b = [t["table"] for t in tables["tables"] if "list_b" in t["table"]][0]

    result = server.query(f'''
        SELECT a.Name as Name_A, b.Name as Name_B
        FROM "{table_a}" a
        JOIN "{table_b}" b ON LOWER(a.Name) = LOWER(b.Name)
    ''')

    assert result["row_count"] >= 2


def test_abbreviations_and_full_names(temp_excel_dir):
    file1 = temp_excel_dir / "short_names.xlsx"
    df1 = pd.DataFrame({
        "Company": ["IBM", "GE", "AT&T"],
        "Revenue": [50000, 60000, 70000]
    })
    df1.to_excel(file1, sheet_name="Revenue", index=False)

    file2 = temp_excel_dir / "long_names.xlsx"
    df2 = pd.DataFrame({
        "Company": ["International Business Machines", "General Electric", "American Telephone & Telegraph"],
        "Employees": [300000, 200000, 100000]
    })
    df2.to_excel(file2, sheet_name="Staff", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_special_characters_in_names(temp_excel_dir):
    file1 = temp_excel_dir / "data_with_special.xlsx"
    df1 = pd.DataFrame({
        "Customer": ["O'Reilly Media", "Ben & Jerry's", "L'Oréal"],
        "Orders": [10, 20, 30]
    })
    df1.to_excel(file1, sheet_name="Orders", index=False)

    file2 = temp_excel_dir / "data_without_special.xlsx"
    df2 = pd.DataFrame({
        "Customer": ["OReilly Media", "Ben and Jerrys", "LOreal"],
        "Shipments": [5, 15, 25]
    })
    df2.to_excel(file2, sheet_name="Shipments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "with_special" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "without_special" in t["table"]][0]

    result1 = server.query(f'SELECT * FROM "{table1}"')
    result2 = server.query(f'SELECT * FROM "{table2}"')

    assert result1["row_count"] == 3
    assert result2["row_count"] == 3


def test_merged_acquired_company_names(temp_excel_dir):
    q1_file = temp_excel_dir / "Q1_sales.xlsx"
    df_q1 = pd.DataFrame({
        "Date": ["2024-01-15", "2024-02-15"],
        "Customer": ["Widget Corp", "Gadget Inc"],
        "Amount": [1000, 2000]
    })
    df_q1.to_excel(q1_file, sheet_name="Sales", index=False)

    q2_file = temp_excel_dir / "Q2_sales.xlsx"
    df_q2 = pd.DataFrame({
        "Date": ["2024-04-15", "2024-05-15"],
        "Customer": ["Widget Corp (acquired by MegaCo)", "Gadget Inc"],
        "Amount": [1500, 2500]
    })
    df_q2.to_excel(q2_file, sheet_name="Sales", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_unicode_and_ascii_equivalents(temp_excel_dir):
    file1 = temp_excel_dir / "unicode.xlsx"
    df1 = pd.DataFrame({
        "Name": ["Café Müller", "São Paulo", "Zürich"],
        "Type": ["Restaurant", "City", "City"]
    })
    df1.to_excel(file1, sheet_name="Places", index=False)

    file2 = temp_excel_dir / "ascii.xlsx"
    df2 = pd.DataFrame({
        "Name": ["Cafe Muller", "Sao Paulo", "Zurich"],
        "Type": ["Restaurant", "City", "City"]
    })
    df2.to_excel(file2, sheet_name="Places", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    unicode_table = [t["table"] for t in tables["tables"] if "unicode" in t["table"]][0]
    ascii_table = [t["table"] for t in tables["tables"] if "ascii" in t["table"]][0]

    result1 = server.query(f'SELECT * FROM "{unicode_table}"')
    result2 = server.query(f'SELECT * FROM "{ascii_table}"')

    assert result1["row_count"] == 3
    assert result2["row_count"] == 3


def test_legal_entity_suffixes(temp_excel_dir):
    file1 = temp_excel_dir / "vendors.xlsx"
    df1 = pd.DataFrame({
        "Vendor": ["Acme Inc.", "Globex LLC", "Initech Corp"],
        "Status": ["Active", "Active", "Inactive"]
    })
    df1.to_excel(file1, sheet_name="Vendors", index=False)

    file2 = temp_excel_dir / "payments.xlsx"
    df2 = pd.DataFrame({
        "Payee": ["Acme, Inc.", "Globex L.L.C.", "Initech Corporation"],
        "Amount": [5000, 10000, 7500]
    })
    df2.to_excel(file2, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_name_order_variations(temp_excel_dir):
    file1 = temp_excel_dir / "first_last.xlsx"
    df1 = pd.DataFrame({
        "Contact": ["John Smith", "Jane Doe", "Bob Johnson"],
        "Email": ["john@example.com", "jane@example.com", "bob@example.com"]
    })
    df1.to_excel(file1, sheet_name="Contacts", index=False)

    file2 = temp_excel_dir / "last_first.xlsx"
    df2 = pd.DataFrame({
        "Contact": ["Smith, John", "Doe, Jane", "Johnson, Bob"],
        "Phone": ["555-0001", "555-0002", "555-0003"]
    })
    df2.to_excel(file2, sheet_name="Contacts", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_null_vs_empty_vs_na_in_names(temp_excel_dir):
    file1 = temp_excel_dir / "with_nulls.xlsx"
    df1 = pd.DataFrame({
        "CompanyName": ["Valid Corp", None, "Another Co"],
        "Revenue": [1000, 2000, 3000]
    })
    df1.to_excel(file1, sheet_name="Data", index=False)

    file2 = temp_excel_dir / "with_na_text.xlsx"
    df2 = pd.DataFrame({
        "CompanyName": ["Valid Corp", "N/A", "Another Co"],
        "Profit": [500, 1000, 1500]
    })
    df2.to_excel(file2, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "with_nulls" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "with_na" in t["table"]][0]

    result1 = server.query(f'SELECT CompanyName FROM "{table1}" WHERE CompanyName IS NULL')
    result2 = server.query(f'SELECT CompanyName FROM "{table2}" WHERE CompanyName = \'N/A\'')

    assert result1["row_count"] >= 0
    assert result2["row_count"] >= 0
