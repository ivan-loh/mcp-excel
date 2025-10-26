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


def test_deal_closed_vs_revenue_booked_dates(temp_excel_dir):
    deals_file = temp_excel_dir / "deals.xlsx"
    df_deals = pd.DataFrame({
        "DealID": ["D001", "D002", "D003"],
        "CloseDate": [
            datetime(2024, 3, 31),
            datetime(2024, 4, 1),
            datetime(2024, 4, 15)
        ],
        "Amount": [10000, 20000, 15000]
    })
    df_deals.to_excel(deals_file, sheet_name="Deals", index=False)

    revenue_file = temp_excel_dir / "revenue.xlsx"
    df_revenue = pd.DataFrame({
        "DealID": ["D001", "D002", "D003"],
        "RevenueDate": [
            datetime(2024, 4, 1),
            datetime(2024, 4, 1),
            datetime(2024, 4, 15)
        ],
        "BookedAmount": [10000, 20000, 15000]
    })
    df_revenue.to_excel(revenue_file, sheet_name="Revenue", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    deals_table = [t["table"] for t in tables["tables"] if "deals" in t["table"]][0]
    revenue_table = [t["table"] for t in tables["tables"] if "revenue" in t["table"]][0]

    result = server.query(f'''
        SELECT
            d.DealID,
            d.CloseDate,
            r.RevenueDate,
            date_diff('day', d.CloseDate, r.RevenueDate) as DaysLag
        FROM "{deals_table}" d
        JOIN "{revenue_table}" r ON d.DealID = r.DealID
    ''')

    assert result["row_count"] == 3


def test_order_vs_shipment_lag(temp_excel_dir):
    orders_file = temp_excel_dir / "orders.xlsx"
    df_orders = pd.DataFrame({
        "OrderID": [1, 2, 3, 4],
        "OrderDate": [
            datetime(2024, 1, 10),
            datetime(2024, 1, 15),
            datetime(2024, 1, 20),
            datetime(2024, 1, 25)
        ],
        "CustomerID": [101, 102, 103, 104]
    })
    df_orders.to_excel(orders_file, sheet_name="Orders", index=False)

    shipments_file = temp_excel_dir / "shipments.xlsx"
    df_shipments = pd.DataFrame({
        "ShipmentID": [1, 2, 3, 4],
        "OrderID": [1, 2, 3, 4],
        "ShipDate": [
            datetime(2024, 1, 12),
            datetime(2024, 1, 20),
            datetime(2024, 1, 22),
            datetime(2024, 2, 1)
        ]
    })
    df_shipments.to_excel(shipments_file, sheet_name="Shipments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    orders_table = [t["table"] for t in tables["tables"] if "orders" in t["table"]][0]
    shipments_table = [t["table"] for t in tables["tables"] if "shipments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            o.OrderID,
            date_diff('day', o.OrderDate, s.ShipDate) as DaysToShip
        FROM "{orders_table}" o
        JOIN "{shipments_table}" s ON o.OrderID = s.OrderID
        WHERE date_diff('day', o.OrderDate, s.ShipDate) > 5
    ''')

    assert result["row_count"] >= 1


def test_invoice_date_vs_payment_date_lag(temp_excel_dir):
    invoices_file = temp_excel_dir / "invoices.xlsx"
    df_invoices = pd.DataFrame({
        "InvoiceNumber": ["INV001", "INV002", "INV003"],
        "InvoiceDate": [
            datetime(2024, 1, 1),
            datetime(2024, 1, 15),
            datetime(2024, 2, 1)
        ],
        "DueDate": [
            datetime(2024, 1, 31),
            datetime(2024, 2, 14),
            datetime(2024, 3, 1)
        ],
        "Amount": [1000, 2000, 1500]
    })
    df_invoices.to_excel(invoices_file, sheet_name="Invoices", index=False)

    payments_file = temp_excel_dir / "payments.xlsx"
    df_payments = pd.DataFrame({
        "InvoiceRef": ["INV001", "INV002", "INV003"],
        "PaymentDate": [
            datetime(2024, 1, 25),
            datetime(2024, 2, 20),
            datetime(2024, 3, 5)
        ],
        "Amount": [1000, 2000, 1500]
    })
    df_payments.to_excel(payments_file, sheet_name="Payments", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    inv_table = [t["table"] for t in tables["tables"] if "invoices" in t["table"]][0]
    pay_table = [t["table"] for t in tables["tables"] if "payments" in t["table"]][0]

    result = server.query(f'''
        SELECT
            i.InvoiceNumber,
            i.DueDate,
            p.PaymentDate,
            date_diff('day', i.DueDate, p.PaymentDate) as DaysOverdue
        FROM "{inv_table}" i
        JOIN "{pay_table}" p ON i.InvoiceNumber = p.InvoiceRef
        WHERE p.PaymentDate > i.DueDate
    ''')

    assert result["row_count"] >= 1


def test_purchase_order_vs_receipt_lag(temp_excel_dir):
    po_file = temp_excel_dir / "purchase_orders.xlsx"
    df_po = pd.DataFrame({
        "PONumber": ["PO-001", "PO-002", "PO-003"],
        "PODate": [
            datetime(2024, 1, 5),
            datetime(2024, 1, 10),
            datetime(2024, 1, 15)
        ],
        "ExpectedDate": [
            datetime(2024, 1, 15),
            datetime(2024, 1, 20),
            datetime(2024, 1, 25)
        ]
    })
    df_po.to_excel(po_file, sheet_name="POs", index=False)

    receipts_file = temp_excel_dir / "receipts.xlsx"
    df_receipts = pd.DataFrame({
        "ReceiptID": [1, 2, 3],
        "PONumber": ["PO-001", "PO-002", "PO-003"],
        "ReceiptDate": [
            datetime(2024, 1, 14),
            datetime(2024, 1, 25),
            datetime(2024, 2, 5)
        ]
    })
    df_receipts.to_excel(receipts_file, sheet_name="Receipts", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    po_table = [t["table"] for t in tables["tables"] if "purchase" in t["table"]][0]
    receipt_table = [t["table"] for t in tables["tables"] if "receipts" in t["table"]][0]

    result = server.query(f'''
        SELECT
            po.PONumber,
            po.ExpectedDate,
            r.ReceiptDate,
            CASE
                WHEN r.ReceiptDate > po.ExpectedDate THEN 'Late'
                ELSE 'On Time'
            END as Status
        FROM "{po_table}" po
        JOIN "{receipt_table}" r ON po.PONumber = r.PONumber
    ''')

    assert result["row_count"] == 3


def test_transaction_timestamp_vs_settlement_date(temp_excel_dir):
    transactions_file = temp_excel_dir / "transactions.xlsx"
    df_trans = pd.DataFrame({
        "TransactionID": [1, 2, 3],
        "TransactionTime": [
            datetime(2024, 1, 15, 23, 45, 0),
            datetime(2024, 1, 20, 0, 15, 0),
            datetime(2024, 1, 25, 12, 30, 0)
        ],
        "Amount": [100, 200, 300]
    })
    df_trans.to_excel(transactions_file, sheet_name="Transactions", index=False)

    settlements_file = temp_excel_dir / "settlements.xlsx"
    df_settle = pd.DataFrame({
        "TransactionID": [1, 2, 3],
        "SettlementDate": [
            datetime(2024, 1, 16),
            datetime(2024, 1, 20),
            datetime(2024, 1, 25)
        ]
    })
    df_settle.to_excel(settlements_file, sheet_name="Settlements", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    trans_table = [t["table"] for t in tables["tables"] if "transactions" in t["table"]][0]
    settle_table = [t["table"] for t in tables["tables"] if "settlements" in t["table"]][0]

    result = server.query(f'''
        SELECT
            t.TransactionID,
            t.TransactionTime,
            s.SettlementDate
        FROM "{trans_table}" t
        JOIN "{settle_table}" s ON t.TransactionID = s.TransactionID
    ''')

    assert result["row_count"] == 3


def test_backdated_entries(temp_excel_dir):
    ledger_file = temp_excel_dir / "ledger.xlsx"
    df_ledger = pd.DataFrame({
        "EntryID": [1, 2, 3],
        "TransactionDate": [
            datetime(2024, 1, 15),
            datetime(2024, 1, 10),
            datetime(2024, 1, 20)
        ],
        "EnteredDate": [
            datetime(2024, 1, 15),
            datetime(2024, 1, 25),
            datetime(2024, 1, 20)
        ],
        "Amount": [100, 200, 300]
    })
    df_ledger.to_excel(ledger_file, sheet_name="Ledger", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    ledger_table = [t["table"] for t in tables["tables"] if "ledger" in t["table"]][0]

    result = server.query(f'''
        SELECT *
        FROM "{ledger_table}"
        WHERE EnteredDate > TransactionDate
    ''')

    assert result["row_count"] >= 1


def test_timezone_differences_in_timestamps(temp_excel_dir):
    eastern_file = temp_excel_dir / "eastern_time.xlsx"
    df_eastern = pd.DataFrame({
        "EventID": [1],
        "EventTime": ["2024-01-15 17:00:00"],
        "Timezone": ["EST"]
    })
    df_eastern.to_excel(eastern_file, sheet_name="Events", index=False)

    pacific_file = temp_excel_dir / "pacific_time.xlsx"
    df_pacific = pd.DataFrame({
        "EventID": [1],
        "EventTime": ["2024-01-15 14:00:00"],
        "Timezone": ["PST"]
    })
    df_pacific.to_excel(pacific_file, sheet_name="Events", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_month_end_cutoff_differences(temp_excel_dir):
    january_file = temp_excel_dir / "january_cutoff.xlsx"
    df_jan = pd.DataFrame({
        "TransactionID": [1, 2],
        "TransactionDate": [
            datetime(2024, 1, 31, 23, 59, 0),
            datetime(2024, 2, 1, 0, 1, 0)
        ],
        "Amount": [1000, 2000],
        "Period": ["January", "February"]
    })
    df_jan.to_excel(january_file, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    trans_table = [t["table"] for t in tables["tables"] if "january" in t["table"]][0]

    result = server.query(f'''
        SELECT
            Period,
            COUNT(*) as TransactionCount
        FROM "{trans_table}"
        GROUP BY Period
    ''')

    assert result["row_count"] >= 1


def test_effective_date_vs_entry_date(temp_excel_dir):
    journal_file = temp_excel_dir / "journal_entries.xlsx"
    df_journal = pd.DataFrame({
        "EntryID": [1, 2, 3],
        "EffectiveDate": [
            datetime(2024, 1, 31),
            datetime(2024, 1, 31),
            datetime(2024, 2, 15)
        ],
        "EntryDate": [
            datetime(2024, 1, 31),
            datetime(2024, 2, 5),
            datetime(2024, 2, 15)
        ],
        "Amount": [5000, 3000, 2000]
    })
    df_journal.to_excel(journal_file, sheet_name="Journal", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    journal_table = [t["table"] for t in tables["tables"] if "journal" in t["table"]][0]

    result = server.query(f'''
        SELECT
            MONTH(EffectiveDate) as EffectiveMonth,
            MONTH(EntryDate) as EntryMonth,
            Amount
        FROM "{journal_table}"
        WHERE MONTH(EffectiveDate) != MONTH(EntryDate)
    ''')

    assert result["row_count"] >= 1


def test_accrual_vs_cash_basis_timing(temp_excel_dir):
    revenue_file = temp_excel_dir / "revenue.xlsx"
    df_revenue = pd.DataFrame({
        "InvoiceID": [1, 2, 3],
        "ServiceDate": [
            datetime(2024, 1, 15),
            datetime(2024, 1, 31),
            datetime(2024, 2, 5)
        ],
        "InvoiceDate": [
            datetime(2024, 1, 31),
            datetime(2024, 2, 5),
            datetime(2024, 2, 10)
        ],
        "PaymentDate": [
            datetime(2024, 2, 15),
            datetime(2024, 2, 20),
            datetime(2024, 3, 1)
        ],
        "Amount": [1000, 1500, 2000]
    })
    df_revenue.to_excel(revenue_file, sheet_name="Revenue", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    revenue_table = [t["table"] for t in tables["tables"] if "revenue" in t["table"]][0]

    result = server.query(f'''
        SELECT
            MONTH(ServiceDate) as AccrualMonth,
            MONTH(PaymentDate) as CashMonth,
            Amount
        FROM "{revenue_table}"
        WHERE MONTH(ServiceDate) != MONTH(PaymentDate)
    ''')

    assert result["row_count"] >= 2
