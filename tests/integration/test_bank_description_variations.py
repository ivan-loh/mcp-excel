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


def test_ach_deposit_description_variations(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_inv = pd.DataFrame({
        "InvoiceNumber": ["INV-1001", "INV-1002"],
        "Customer": ["Acme Corp", "Global Industries"],
        "Amount": [5000.00, 7500.00]
    })
    df_inv.to_excel(invoices_file, sheet_name="Invoices", index=False)

    bank_file = temp_excel_dir / "bank_statements.xlsx"
    df_bank = pd.DataFrame({
        "TransactionID": [1, 2],
        "Description": [
            "ACH DEPOSIT ACME CORP INV1001",
            "GLOBAL INDUSTRIES ACH CREDIT INV1002"
        ],
        "Amount": [5000.00, 7500.00]
    })
    df_bank.to_excel(bank_file, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.InvoiceNumber,
            b.Description,
            i.Amount
        FROM "{inv_table}" i
        JOIN "{bank_table}" b ON i.Amount = b.Amount
        WHERE b.Description LIKE '%' || REPLACE(i.InvoiceNumber, '-', '') || '%'
    ''')

    assert result["row_count"] >= 1


def test_wire_transfer_vs_check_descriptions(temp_excel_dir):
    payments_file = temp_excel_dir / "expected_payments.xlsx"
    df_pay = pd.DataFrame({
        "PaymentRef": ["PAY001", "PAY002", "PAY003"],
        "PaymentMethod": ["Wire", "Check", "ACH"],
        "Amount": [10000.00, 5000.00, 2500.00]
    })
    df_pay.to_excel(payments_file, sheet_name="Payments", index=False)

    bank_file = temp_excel_dir / "bank_feeds.xlsx"
    df_bank = pd.DataFrame({
        "Date": ["2024-01-15", "2024-01-16", "2024-01-17"],
        "Description": [
            "WIRE TRANSFER FROM CUSTOMER PAY001",
            "CHECK #12345 DEPOSIT",
            "ACH CREDIT PAY003"
        ],
        "Amount": [10000.00, 5000.00, 2500.00]
    })
    df_bank.to_excel(bank_file, sheet_name="BankFeed", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    pay_table = [t["table"] for t in tables["tables"] if "expected" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            p.PaymentRef,
            p.PaymentMethod,
            b.Description
        FROM "{pay_table}" p
        JOIN "{bank_table}" b ON p.Amount = b.Amount
    ''')

    assert result["row_count"] == 3


def test_truncated_description_fields(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_inv = pd.DataFrame({
        "InvoiceNumber": ["INV-2024-0001"],
        "Customer": ["Very Long Corporation Name International Holdings LLC"],
        "Amount": [15000.00]
    })
    df_inv.to_excel(invoices_file, sheet_name="Invoices", index=False)

    bank_file = temp_excel_dir / "bank.xlsx"
    df_bank = pd.DataFrame({
        "TransactionID": [1],
        "Description": ["VERY LONG CORPORATION NAME I"],
        "Amount": [15000.00]
    })
    df_bank.to_excel(bank_file, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT *
        FROM "{bank_table}" b
        WHERE EXISTS (
            SELECT 1 FROM "{inv_table}" i
            WHERE i.Amount = b.Amount
            AND UPPER(i.Customer) LIKE UPPER(b.Description) || '%'
        )
    ''')

    assert result["row_count"] >= 1


def test_batch_deposit_single_line(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_inv = pd.DataFrame({
        "InvoiceNumber": ["INV-001", "INV-002", "INV-003", "INV-004", "INV-005"],
        "Amount": [100.00, 200.00, 300.00, 400.00, 1000.00]
    })
    df_inv.to_excel(invoices_file, sheet_name="Invoices", index=False)

    bank_file = temp_excel_dir / "bank_deposits.xlsx"
    df_bank = pd.DataFrame({
        "Date": ["2024-01-15"],
        "Description": ["BATCH DEPOSIT MULTIPLE INVOICES"],
        "Amount": [2000.00]
    })
    df_bank.to_excel(bank_file, sheet_name="Deposits", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            (SELECT SUM(Amount) FROM "{inv_table}" WHERE Amount < 500) as InvoiceTotal,
            b.Amount as BankDeposit
        FROM "{bank_table}" b
    ''')

    assert result["row_count"] == 1


def test_memo_field_variations(temp_excel_dir):
    transactions_file = temp_excel_dir / "transactions.xlsx"
    df_trans = pd.DataFrame({
        "TransactionID": ["T001", "T002", "T003"],
        "Memo": ["Payment for Invoice #123", "INV-456 Payment", "Ref: 789"],
        "Amount": [500.00, 750.00, 1000.00]
    })
    df_trans.to_excel(transactions_file, sheet_name="Transactions", index=False)

    bank_file = temp_excel_dir / "bank.xlsx"
    df_bank = pd.DataFrame({
        "ID": [1, 2, 3],
        "BankMemo": ["INV 123", "INVOICE 456", "REF 789"],
        "Amount": [500.00, 750.00, 1000.00]
    })
    df_bank.to_excel(bank_file, sheet_name="BankTransactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    trans_table = [t["table"] for t in tables["tables"] if "transactions" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            t.TransactionID,
            t.Memo,
            b.BankMemo
        FROM "{trans_table}" t
        JOIN "{bank_table}" b ON t.Amount = b.Amount
    ''')

    assert result["row_count"] == 3


def test_special_characters_in_descriptions(temp_excel_dir):
    customers_file = temp_excel_dir / "customers.xlsx"
    df_cust = pd.DataFrame({
        "CustomerName": ["O'Reilly Media", "Ben & Jerry's", "Toys \"R\" Us"],
        "CustomerID": [1, 2, 3]
    })
    df_cust.to_excel(customers_file, sheet_name="Customers", index=False)

    bank_file = temp_excel_dir / "bank.xlsx"
    df_bank = pd.DataFrame({
        "Description": ["OREILLY MEDIA", "BEN AND JERRYS", "TOYS R US"],
        "Amount": [100.00, 200.00, 300.00],
        "CustomerID": [1, 2, 3]
    })
    df_bank.to_excel(bank_file, sheet_name="Deposits", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    cust_table = [t["table"] for t in tables["tables"] if "customers" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            c.CustomerName,
            b.Description
        FROM "{cust_table}" c
        JOIN "{bank_table}" b ON c.CustomerID = b.CustomerID
    ''')

    assert result["row_count"] == 3


def test_duplicate_description_different_amounts(temp_excel_dir):
    bank_file = temp_excel_dir / "bank_statement.xlsx"
    df_bank = pd.DataFrame({
        "Date": ["2024-01-10", "2024-01-15", "2024-01-20"],
        "Description": ["ACME CORP PAYMENT", "ACME CORP PAYMENT", "ACME CORP PAYMENT"],
        "Amount": [1000.00, 1500.00, 1000.00]
    })
    df_bank.to_excel(bank_file, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            Description,
            Date,
            Amount
        FROM "{bank_table}"
        WHERE Description = 'ACME CORP PAYMENT'
        ORDER BY Date
    ''')

    assert result["row_count"] == 3


def test_payment_reversal_description(temp_excel_dir):
    bank_file = temp_excel_dir / "bank_activity.xlsx"
    df_bank = pd.DataFrame({
        "Date": ["2024-01-15", "2024-01-16"],
        "Description": [
            "CUSTOMER PAYMENT INV-123",
            "REVERSAL CUSTOMER PAYMENT INV-123"
        ],
        "Amount": [5000.00, -5000.00]
    })
    df_bank.to_excel(bank_file, sheet_name="Activity", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    bank_table = [t["table"] for t in tables["tables"] if "bank" in t["table"]][0]

    result = server.query(f'''
        SELECT
            SUM(Amount) as NetAmount
        FROM "{bank_table}"
        WHERE Description LIKE '%INV-123%'
    ''')

    assert result["row_count"] == 1


def test_foreign_currency_description(temp_excel_dir):
    bank_file = temp_excel_dir / "forex_transactions.xlsx"
    df_bank = pd.DataFrame({
        "Date": ["2024-01-15", "2024-01-16"],
        "Description": [
            "FX CONVERSION EUR TO USD",
            "WIRE TRANSFER IN EUR CONVERTED"
        ],
        "Amount_USD": [1180.00, 2360.00],
        "OriginalCurrency": ["EUR", "EUR"],
        "OriginalAmount": [1000.00, 2000.00]
    })
    df_bank.to_excel(bank_file, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    bank_table = [t["table"] for t in tables["tables"] if "forex" in t["table"]][0]

    result = server.query(f'''
        SELECT *
        FROM "{bank_table}"
        WHERE Description LIKE '%EUR%' OR OriginalCurrency = 'EUR'
    ''')

    assert result["row_count"] == 2


def test_stop_payment_vs_voided_check(temp_excel_dir):
    checks_file = temp_excel_dir / "checks_issued.xlsx"
    df_checks = pd.DataFrame({
        "CheckNumber": [1001, 1002, 1003],
        "Payee": ["Vendor A", "Vendor B", "Vendor C"],
        "Amount": [500.00, 750.00, 1000.00],
        "Status": ["Cleared", "Voided", "Stop Payment"]
    })
    df_checks.to_excel(checks_file, sheet_name="Checks", index=False)

    bank_file = temp_excel_dir / "bank_cleared.xlsx"
    df_bank = pd.DataFrame({
        "CheckNumber": [1001],
        "Description": ["CHECK #1001 VENDOR A"],
        "Amount": [500.00]
    })
    df_bank.to_excel(bank_file, sheet_name="ClearedChecks", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    checks_table = [t["table"] for t in tables["tables"] if "checks_issued" in t["table"]][0]
    bank_table = [t["table"] for t in tables["tables"] if "bank_cleared" in t["table"]][0]

    result = server.query(f'''
        SELECT
            c.CheckNumber,
            c.Status,
            b.Description
        FROM "{checks_table}" c
        LEFT JOIN "{bank_table}" b ON c.CheckNumber = b.CheckNumber
        WHERE c.Status IN ('Voided', 'Stop Payment') AND b.CheckNumber IS NOT NULL
    ''')

    assert result["row_count"] == 0
