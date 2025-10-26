import pytest
from pathlib import Path
import mcp_excel.server as server
from mcp_excel.loader import ExcelLoader
from mcp_excel.types import SheetOverride, LocaleConfig

pytestmark = [pytest.mark.integration, pytest.mark.usefixtures("setup_server")]

examples_dir = Path(__file__).parent.parent / "examples"
finance_dir = examples_dir / "finance"


def test_load_examples_directory():
    if not examples_dir.exists():
        pytest.skip("Examples directory not found")

    result = server.load_dir(path=str(examples_dir))
    assert result["files_count"] >= 10
    assert result["sheets_count"] >= 12
    assert result["tables_count"] >= 12


def test_load_with_overrides():
    overrides_file = finance_dir / "finance_overrides.yaml"
    if not overrides_file.exists():
        pytest.skip("finance_overrides.yaml not found")

    import yaml
    with open(overrides_file) as f:
        overrides = yaml.safe_load(f)

    result = server.load_dir(path=str(finance_dir), overrides=overrides)
    assert result["files_count"] >= 10


def test_system_views_with_examples():
    if not examples_dir.exists():
        pytest.skip("Examples directory not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    files_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.__files"')
    assert files_result["row_count"] == 1
    assert files_result["rows"][0][0] >= 10

    tables_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.__tables"')
    assert tables_result["row_count"] == 1
    assert tables_result["rows"][0][0] >= 12


@pytest.mark.xfail(reason="European number format parsing not yet implemented for RAW mode CSV")
def test_european_number_format_detection():
    csv_file = examples_dir / "sales_european.csv"
    if not csv_file.exists():
        pytest.skip("sales_european.csv not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    result = server.query(f'SELECT "Revenue", "Cost", "Profit Margin" FROM "{alias}.sales_european.sheet1" LIMIT 1')
    assert result["row_count"] == 1

    revenue = result["rows"][0][0]
    assert isinstance(revenue, (int, float)), f"Revenue is {type(revenue)} not numeric"
    assert revenue > 1000, f"Revenue should be >1000, got {revenue}"


def test_latin1_encoding_detection():
    csv_file = examples_dir / "employees_latin1.csv"
    if not csv_file.exists():
        pytest.skip("employees_latin1.csv not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    result = server.query(f'SELECT "Name" FROM "{alias}.employees_latin1.sheet1"')
    assert result["row_count"] >= 10

    names = [row[0] for row in result["rows"]]
    assert any('José' in str(name) or 'jose' in str(name).lower() for name in names)
    assert any('François' in str(name) or 'francois' in str(name).lower() for name in names)


def test_utf8_bom_detection():
    csv_file = examples_dir / "products_utf8_bom.csv"
    if not csv_file.exists():
        pytest.skip("products_utf8_bom.csv not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    result = server.query(f'SELECT * FROM "{alias}.products_utf8_bom.sheet1" LIMIT 1')
    assert result["row_count"] == 1

    first_col_name = result["columns"][0]["name"]
    assert not first_col_name.startswith('\ufeff'), "BOM not properly removed from column name"
    assert "Product" in first_col_name or "product" in first_col_name.lower()


def test_quarterly_comparison_multi_table():
    xlsx_file = examples_dir / "quarterly_comparison.xlsx"
    if not xlsx_file.exists():
        pytest.skip("quarterly_comparison.xlsx not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    tables = server.list_tables(alias=alias)
    quarterly_table = next((t for t in tables["tables"] if "quarterly_comparison" in t["table"]), None)
    assert quarterly_table is not None, "quarterly_comparison table not found"

    result = server.query(f'SELECT COUNT(*) as count FROM "{quarterly_table["table"]}"')
    assert result["row_count"] == 1

    row_count = result["rows"][0][0]
    assert row_count >= 7, f"Expected >=7 rows from quarterly data, got {row_count}"


def test_general_ledger_performance():
    xlsx_file = finance_dir / "general_ledger.xlsx"
    if not xlsx_file.exists():
        pytest.skip("general_ledger.xlsx not found")

    import time
    alias = finance_dir.name
    start = time.time()
    server.load_dir(path=str(finance_dir))
    load_time = time.time() - start

    assert load_time < 10.0, f"Load took {load_time:.2f}s, expected <10s"

    result = server.query(f'SELECT COUNT(*) FROM "{alias}.general_ledger.entries"')
    assert result["rows"][0][0] >= 1000


def test_financial_statements_multi_sheet():
    xlsx_file = finance_dir / "financial_statements.xlsx"
    if not xlsx_file.exists():
        pytest.skip("financial_statements.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    income = server.query(f'SELECT COUNT(*) FROM "{alias}.financial_statements.income_statement"')
    balance = server.query(f'SELECT COUNT(*) FROM "{alias}.financial_statements.balance_sheet"')
    cashflow = server.query(f'SELECT COUNT(*) FROM "{alias}.financial_statements.cash_flow"')

    assert income["rows"][0][0] >= 1
    assert balance["rows"][0][0] >= 19
    assert cashflow["rows"][0][0] >= 16


def test_trial_balance_messy_headers():
    xlsx_file = finance_dir / "trial_balance.xlsx"
    if not xlsx_file.exists():
        pytest.skip("trial_balance.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    tables = server.list_tables(alias=alias)
    trial_table = next((t for t in tables["tables"] if "trial_balance" in t["table"]), None)
    assert trial_table is not None, "trial_balance table not found"

    result = server.query(f'SELECT * FROM "{trial_table["table"]}" LIMIT 5')
    assert result["row_count"] >= 1
    assert len(result["columns"]) >= 1


def test_revenue_segment_data_quality():
    xlsx_file = finance_dir / "revenue_by_segment.xlsx"
    if not xlsx_file.exists():
        pytest.skip("revenue_by_segment.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    result = server.query(f'SELECT COUNT(*) FROM "{alias}.revenue_by_segment.revenue"')
    assert result["rows"][0][0] >= 1000


def test_invoice_register_date_parsing():
    xlsx_file = finance_dir / "invoice_register.xlsx"
    if not xlsx_file.exists():
        pytest.skip("invoice_register.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    tables = server.list_tables(alias=alias)
    invoice_table = next((t for t in tables["tables"] if "invoice_register" in t["table"]), None)
    assert invoice_table is not None, "invoice_register table not found"

    schema = server.get_schema(invoice_table["table"])
    result = server.query(f'SELECT COUNT(*) FROM "{invoice_table["table"]}"')
    assert result["rows"][0][0] >= 400


def test_expense_reports_normalization():
    xlsx_file = finance_dir / "expense_reports.xlsx"
    if not xlsx_file.exists():
        pytest.skip("expense_reports.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    tables = server.list_tables(alias=alias)
    expense_table = next((t for t in tables["tables"] if "expense_reports" in t["table"]), None)
    assert expense_table is not None, "expense_reports table not found"

    result = server.query(f'SELECT COUNT(*) FROM "{expense_table["table"]}"')
    assert result["rows"][0][0] >= 200


def test_budget_vs_actuals_comparison():
    xlsx_file = finance_dir / "budget_vs_actuals.xlsx"
    if not xlsx_file.exists():
        pytest.skip("budget_vs_actuals.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    tables = server.list_tables(alias=alias)
    budget_table = next((t for t in tables["tables"] if "budget_vs_actuals" in t["table"]), None)
    assert budget_table is not None, "budget_vs_actuals table not found"

    result = server.query(f'SELECT COUNT(*) FROM "{budget_table["table"]}"')
    assert result["rows"][0][0] >= 60


def test_financial_ratios_kpi_data():
    xlsx_file = finance_dir / "financial_ratios.xlsx"
    if not xlsx_file.exists():
        pytest.skip("financial_ratios.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    tables = server.list_tables(alias=alias)
    ratios_table = next((t for t in tables["tables"] if "financial_ratios" in t["table"]), None)
    assert ratios_table is not None, "financial_ratios table not found"

    result = server.query(f'SELECT COUNT(*) FROM "{ratios_table["table"]}"')
    assert result["rows"][0][0] >= 13


def test_cash_flow_forecast_wide_format():
    xlsx_file = finance_dir / "cash_flow_forecast.xlsx"
    if not xlsx_file.exists():
        pytest.skip("cash_flow_forecast.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    tables = server.list_tables(alias=alias)
    forecast_table = next((t for t in tables["tables"] if "cash_flow_forecast" in t["table"]), None)
    assert forecast_table is not None, "cash_flow_forecast table not found"

    result = server.query(f'SELECT * FROM "{forecast_table["table"]}" LIMIT 1')
    assert len(result["columns"]) >= 5


def test_accounts_receivable_aging():
    xlsx_file = finance_dir / "accounts_receivable.xlsx"
    if not xlsx_file.exists():
        pytest.skip("accounts_receivable.xlsx not found")

    alias = finance_dir.name
    server.load_dir(path=str(finance_dir))

    result = server.query(f'SELECT COUNT(*) FROM "{alias}.accounts_receivable.ar_aging"')
    assert result["rows"][0][0] >= 300
