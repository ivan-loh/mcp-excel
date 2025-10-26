import pytest
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
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


def test_single_invoice_multiple_payments(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV-001", "INV-002"],
        "Amount": [1000.00, 2000.00],
        "CustomerID": [1, 2]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1, 2, 3, 4],
        "InvoiceRef": ["INV-001", "INV-001", "INV-001", "INV-002"],
        "Amount": [300.00, 300.00, 400.00, 2000.00],
        "PaymentDate": ["2024-01-15", "2024-01-20", "2024-01-25", "2024-01-10"]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.InvoiceNumber,
            i.Amount as InvoiceAmount,
            SUM(p.Amount) as TotalPaid,
            i.Amount - COALESCE(SUM(p.Amount), 0) as Balance
        FROM "{inv_table}" i
        LEFT JOIN "{pay_table}" p ON i.InvoiceNumber = p.InvoiceRef
        GROUP BY i.InvoiceNumber, i.Amount
    ''')

    assert result["row_count"] == 2
    rows = {row[0]: row for row in result["rows"]}
    assert abs(rows["INV-001"][3]) < 0.01


def test_payment_without_invoice_reference(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceID": ["A001", "A002", "A003"],
        "Amount": [500.00, 750.00, 1000.00]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1, 2, 3, 4],
        "InvoiceRef": ["A001", None, "A003", "UNKNOWN"],
        "Amount": [500.00, 750.00, 1000.00, 200.00]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT * FROM "{pay_table}"
        WHERE InvoiceRef IS NULL OR InvoiceRef = 'UNKNOWN'
    ''')

    assert result["row_count"] >= 1


def test_overpayment_scenario(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV-100"],
        "Amount": [1000.00]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1, 2],
        "InvoiceRef": ["INV-100", "INV-100"],
        "Amount": [1000.00, 50.00]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.InvoiceNumber,
            i.Amount as InvoiceAmount,
            SUM(p.Amount) as TotalPaid,
            i.Amount - SUM(p.Amount) as Balance
        FROM "{inv_table}" i
        JOIN "{pay_table}" p ON i.InvoiceNumber = p.InvoiceRef
        GROUP BY i.InvoiceNumber, i.Amount
        HAVING SUM(p.Amount) > i.Amount
    ''')

    assert result["row_count"] == 1


def test_bulk_payment_covering_multiple_invoices(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV-201", "INV-202", "INV-203"],
        "CustomerID": [1, 1, 1],
        "Amount": [100.00, 200.00, 300.00]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1],
        "CustomerID": [1],
        "Amount": [600.00],
        "Note": ["Payment for multiple invoices"]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.CustomerID,
            SUM(i.Amount) as TotalInvoiced,
            p.Amount as TotalPaid
        FROM "{inv_table}" i
        JOIN "{pay_table}" p ON i.CustomerID = p.CustomerID
        GROUP BY i.CustomerID, p.Amount
    ''')

    assert result["row_count"] == 1


def test_voided_payment_scenario(temp_excel_dir):
    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1, 2, 3],
        "InvoiceRef": ["INV-001", "INV-001", "INV-002"],
        "Amount": [500.00, -500.00, 1000.00],
        "Status": ["Posted", "Voided", "Posted"]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            InvoiceRef,
            SUM(Amount) as NetPayment
        FROM "{pay_table}"
        GROUP BY InvoiceRef
    ''')

    assert result["row_count"] >= 1


def test_payment_applied_to_wrong_invoice(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV-301", "INV-302"],
        "Amount": [1000.00, 2000.00]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_original = temp_excel_dir / "payments_original.xlsx"
    df_pay_orig = pd.DataFrame({
        "PaymentID": [1],
        "InvoiceRef": ["INV-302"],
        "Amount": [1000.00]
    })
    df_pay_orig.to_excel(payments_original, sheet_name="Payments", index=False)

    payments_corrected = temp_excel_dir / "payments_corrected.xlsx"
    df_pay_corr = pd.DataFrame({
        "PaymentID": [1],
        "InvoiceRef": ["INV-301"],
        "Amount": [1000.00],
        "CorrectedFrom": ["INV-302"]
    })
    df_pay_corr.to_excel(payments_corrected, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) >= 2


def test_credit_memo_applied_to_invoice(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "DocumentNumber": ["INV-401", "CM-401"],
        "Type": ["Invoice", "Credit Memo"],
        "Amount": [1000.00, -100.00],
        "RelatedTo": [None, "INV-401"]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Documents", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    doc_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]

    result = server.query(f'''
        SELECT
            RelatedTo as InvoiceNumber,
            SUM(Amount) as NetAmount
        FROM "{doc_table}"
        WHERE RelatedTo IS NOT NULL OR DocumentNumber LIKE 'INV%'
        GROUP BY RelatedTo
    ''')

    assert result["row_count"] >= 1


def test_deposit_vs_payment(temp_excel_dir):
    transactions_file = temp_excel_dir / "transactions.xlsx"
    df_transactions = pd.DataFrame({
        "TransactionID": [1, 2, 3, 4],
        "Type": ["Deposit", "Invoice", "Payment", "Refund"],
        "Amount": [500.00, -1000.00, 500.00, 100.00],
        "InvoiceRef": [None, "INV-501", "INV-501", None]
    })
    df_transactions.to_excel(transactions_file, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    trans_table = [t["table"] for t in tables["tables"] if "transactions" in t["table"]][0]

    result = server.query(f'''
        SELECT
            Type,
            SUM(Amount) as Total
        FROM "{trans_table}"
        GROUP BY Type
    ''')

    assert result["row_count"] >= 3


def test_payment_made_before_invoice_issued(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV-601"],
        "InvoiceDate": [datetime(2024, 1, 20)],
        "Amount": [1000.00]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1],
        "InvoiceRef": ["INV-601"],
        "PaymentDate": [datetime(2024, 1, 15)],
        "Amount": [1000.00]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.InvoiceNumber,
            i.InvoiceDate,
            p.PaymentDate
        FROM "{inv_table}" i
        JOIN "{pay_table}" p ON i.InvoiceNumber = p.InvoiceRef
        WHERE p.PaymentDate < i.InvoiceDate
    ''')

    assert result["row_count"] == 1


def test_foreign_currency_payment_reconciliation(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV-701"],
        "Amount_USD": [1000.00],
        "Currency": ["USD"]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "PaymentID": [1],
        "InvoiceRef": ["INV-701"],
        "Amount_EUR": [850.00],
        "Currency": ["EUR"],
        "ExchangeRate": [1.18]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.InvoiceNumber,
            i.Amount_USD,
            p.Amount_EUR * p.ExchangeRate as Amount_USD_Converted
        FROM "{inv_table}" i
        JOIN "{pay_table}" p ON i.InvoiceNumber = p.InvoiceRef
    ''')

    assert result["row_count"] == 1
