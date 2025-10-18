import pytest
from pathlib import Path
import mcp_excel.server as server

pytestmark = [pytest.mark.integration, pytest.mark.usefixtures("setup_server")]


def test_load_examples_directory():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not examples_dir.exists():
        pytest.skip("Examples directory not found")

    result = server.load_dir(path=str(examples_dir))

    assert result["files_count"] >= 10
    assert result["sheets_count"] >= 12
    assert result["tables_count"] >= 12


def test_load_general_ledger_raw_mode():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not (examples_dir / "general_ledger.xlsx").exists():
        pytest.skip("general_ledger.xlsx not found")

    alias = examples_dir.name
    result = server.load_dir(path=str(examples_dir))

    tables = server.list_tables(alias=alias)
    gl_tables = [t for t in tables["tables"] if "general_ledger" in t["table"]]

    assert len(gl_tables) >= 1
    assert any(t["mode"] == "RAW" for t in gl_tables)


def test_query_general_ledger():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not (examples_dir / "general_ledger.xlsx").exists():
        pytest.skip("general_ledger.xlsx not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.general_ledger.entries"')

    assert result["row_count"] == 1
    assert result["rows"][0][0] >= 1000


def test_load_with_overrides():
    examples_dir = Path(__file__).parent.parent / "examples"
    overrides_file = examples_dir / "finance_overrides.yaml"

    if not overrides_file.exists():
        pytest.skip("finance_overrides.yaml not found")

    import yaml
    with open(overrides_file) as f:
        overrides = yaml.safe_load(f)

    result = server.load_dir(path=str(examples_dir), overrides=overrides)

    assert result["files_count"] >= 10


def test_query_financial_statements():
    examples_dir = Path(__file__).parent.parent / "examples"
    if not (examples_dir / "financial_statements.xlsx").exists():
        pytest.skip("financial_statements.xlsx not found")

    alias = examples_dir.name
    server.load_dir(path=str(examples_dir))

    income_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.financial_statements.income_statement"')
    assert income_result["row_count"] == 1
    assert income_result["rows"][0][0] >= 1

    balance_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.financial_statements.balance_sheet"')
    assert balance_result["row_count"] == 1
    assert balance_result["rows"][0][0] >= 19

    cashflow_result = server.query(f'SELECT COUNT(*) as count FROM "{alias}.financial_statements.cash_flow"')
    assert cashflow_result["row_count"] == 1
    assert cashflow_result["rows"][0][0] >= 16


def test_system_views_with_examples():
    examples_dir = Path(__file__).parent.parent / "examples"
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
