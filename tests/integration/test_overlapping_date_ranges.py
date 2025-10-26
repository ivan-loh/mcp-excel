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


def test_monthly_files_with_one_day_overlap(temp_excel_dir):
    file1 = temp_excel_dir / "january.xlsx"
    dates1 = pd.date_range(start="2024-01-01", end="2024-02-01", freq="D")
    df1 = pd.DataFrame({
        "Date": dates1,
        "Sales": range(len(dates1))
    })
    df1.to_excel(file1, sheet_name="Sales", index=False)

    file2 = temp_excel_dir / "february.xlsx"
    dates2 = pd.date_range(start="2024-02-01", end="2024-02-29", freq="D")
    df2 = pd.DataFrame({
        "Date": dates2,
        "Sales": range(100, 100 + len(dates2))
    })
    df2.to_excel(file2, sheet_name="Sales", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    jan_table = [t["table"] for t in tables["tables"] if "january" in t["table"]][0]
    feb_table = [t["table"] for t in tables["tables"] if "february" in t["table"]][0]

    result = server.query(f'''
        SELECT Date, Sales FROM "{jan_table}"
        UNION ALL
        SELECT Date, Sales FROM "{feb_table}"
    ''')

    total_rows = result["row_count"]
    expected_days_jan = 32
    expected_days_feb = 29
    expected_total_with_overlap = expected_days_jan + expected_days_feb

    assert total_rows == expected_total_with_overlap


def test_weekly_snapshots_with_full_overlap(temp_excel_dir):
    base_date = datetime(2024, 1, 1)

    for week in range(4):
        file_path = temp_excel_dir / f"week_{week + 1}.xlsx"
        start_date = base_date + timedelta(weeks=week)
        dates = pd.date_range(start=start_date, periods=7, freq="D")

        df = pd.DataFrame({
            "Date": dates,
            "Inventory": [1000 + (week * 100) + i for i in range(7)]
        })
        df.to_excel(file_path, sheet_name="Inventory", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 4

    all_tables = [t["table"] for t in tables["tables"]]
    union_query = " UNION ALL ".join([f'SELECT Date, Inventory FROM "{t}"' for t in all_tables])

    result = server.query(union_query)
    assert result["row_count"] == 28


def test_daily_and_monthly_aggregates_mixed(temp_excel_dir):
    daily_file = temp_excel_dir / "daily_transactions.xlsx"
    daily_dates = pd.date_range(start="2024-01-01", end="2024-01-31", freq="D")
    df_daily = pd.DataFrame({
        "Date": daily_dates,
        "Amount": [100 + i for i in range(len(daily_dates))],
        "Type": ["Daily"] * len(daily_dates)
    })
    df_daily.to_excel(daily_file, sheet_name="Data", index=False)

    monthly_file = temp_excel_dir / "monthly_summary.xlsx"
    df_monthly = pd.DataFrame({
        "Date": [datetime(2024, 1, 31)],
        "Amount": [sum(range(100, 100 + 31))],
        "Type": ["Monthly"]
    })
    df_monthly.to_excel(monthly_file, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    daily_table = [t["table"] for t in tables["tables"] if "daily" in t["table"]][0]
    monthly_table = [t["table"] for t in tables["tables"] if "monthly" in t["table"]][0]

    result = server.query(f'''
        SELECT Type, COUNT(*) as count, SUM(Amount) as total
        FROM (
            SELECT Date, Amount, Type FROM "{daily_table}"
            UNION ALL
            SELECT Date, Amount, Type FROM "{monthly_table}"
        )
        GROUP BY Type
    ''')

    assert result["row_count"] == 2


def test_overlapping_transaction_windows(temp_excel_dir):
    file1 = temp_excel_dir / "morning_batch.xlsx"
    df1 = pd.DataFrame({
        "TransactionID": [1, 2, 3],
        "Timestamp": ["2024-01-15 08:00:00", "2024-01-15 10:00:00", "2024-01-15 11:59:59"],
        "Amount": [100, 200, 300]
    })
    df1.to_excel(file1, sheet_name="Transactions", index=False)

    file2 = temp_excel_dir / "afternoon_batch.xlsx"
    df2 = pd.DataFrame({
        "TransactionID": [3, 4, 5],
        "Timestamp": ["2024-01-15 11:59:59", "2024-01-15 14:00:00", "2024-01-15 16:00:00"],
        "Amount": [300, 400, 500]
    })
    df2.to_excel(file2, sheet_name="Transactions", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    morning_table = [t["table"] for t in tables["tables"] if "morning" in t["table"]][0]
    afternoon_table = [t["table"] for t in tables["tables"] if "afternoon" in t["table"]][0]

    result = server.query(f'''
        SELECT TransactionID, COUNT(*) as occurrence_count
        FROM (
            SELECT TransactionID FROM "{morning_table}"
            UNION ALL
            SELECT TransactionID FROM "{afternoon_table}"
        )
        GROUP BY TransactionID
        HAVING COUNT(*) > 1
    ''')

    assert result["row_count"] >= 1


def test_quarterly_files_with_month_overlap(temp_excel_dir):
    q1_file = temp_excel_dir / "Q1_2024.xlsx"
    q1_dates = pd.date_range(start="2024-01-01", end="2024-03-31", freq="D")
    df_q1 = pd.DataFrame({
        "Date": q1_dates,
        "Revenue": [1000 + i for i in range(len(q1_dates))]
    })
    df_q1.to_excel(q1_file, sheet_name="Revenue", index=False)

    march_file = temp_excel_dir / "March_2024_Updated.xlsx"
    march_dates = pd.date_range(start="2024-03-01", end="2024-03-31", freq="D")
    df_march = pd.DataFrame({
        "Date": march_dates,
        "Revenue": [2000 + i for i in range(len(march_dates))]
    })
    df_march.to_excel(march_file, sheet_name="Revenue", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    q1_table = [t["table"] for t in tables["tables"] if "q1" in t["table"].lower()][0]
    march_table = [t["table"] for t in tables["tables"] if "march" in t["table"]][0]

    result = server.query(f'''
        SELECT Date, COUNT(*) as file_count
        FROM (
            SELECT Date FROM "{q1_table}"
            UNION ALL
            SELECT Date FROM "{march_table}"
        )
        GROUP BY Date
        HAVING COUNT(*) > 1
    ''')

    assert result["row_count"] == 31


def test_gap_in_date_ranges(temp_excel_dir):
    file1 = temp_excel_dir / "jan_feb.xlsx"
    dates1 = pd.date_range(start="2024-01-01", end="2024-02-15", freq="D")
    df1 = pd.DataFrame({
        "Date": dates1,
        "Value": range(len(dates1))
    })
    df1.to_excel(file1, sheet_name="Data", index=False)

    file2 = temp_excel_dir / "march.xlsx"
    dates2 = pd.date_range(start="2024-03-01", end="2024-03-31", freq="D")
    df2 = pd.DataFrame({
        "Date": dates2,
        "Value": range(100, 100 + len(dates2))
    })
    df2.to_excel(file2, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "jan_feb" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "march" in t["table"]][0]

    result = server.query(f'''
        SELECT MIN(Date) as min_date, MAX(Date) as max_date
        FROM (
            SELECT Date FROM "{table1}"
            UNION ALL
            SELECT Date FROM "{table2}"
        )
    ''')

    assert result["row_count"] == 1


def test_duplicate_entire_datasets(temp_excel_dir):
    df = pd.DataFrame({
        "ID": [1, 2, 3],
        "Name": ["Alice", "Bob", "Charlie"],
        "Amount": [100, 200, 300]
    })

    file1 = temp_excel_dir / "original.xlsx"
    df.to_excel(file1, sheet_name="Data", index=False)

    file2 = temp_excel_dir / "backup.xlsx"
    df.to_excel(file2, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "original" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "backup" in t["table"]][0]

    result = server.query(f'''
        SELECT ID, Name, Amount
        FROM "{table1}"
        UNION ALL
        SELECT ID, Name, Amount
        FROM "{table2}"
    ''')

    assert result["row_count"] == 6


def test_partial_overlap_with_different_granularity(temp_excel_dir):
    hourly_file = temp_excel_dir / "hourly_metrics.xlsx"
    hourly_times = pd.date_range(start="2024-01-15 00:00:00", end="2024-01-15 23:00:00", freq="H")
    df_hourly = pd.DataFrame({
        "Timestamp": hourly_times,
        "Metric": [100 + i for i in range(len(hourly_times))],
        "Granularity": ["Hourly"] * len(hourly_times)
    })
    df_hourly.to_excel(hourly_file, sheet_name="Metrics", index=False)

    daily_file = temp_excel_dir / "daily_summary.xlsx"
    df_daily = pd.DataFrame({
        "Timestamp": [datetime(2024, 1, 15)],
        "Metric": [sum(range(100, 100 + 24))],
        "Granularity": ["Daily"]
    })
    df_daily.to_excel(daily_file, sheet_name="Metrics", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_overlapping_fiscal_and_calendar_periods(temp_excel_dir):
    calendar_file = temp_excel_dir / "calendar_q1.xlsx"
    cal_dates = pd.date_range(start="2024-01-01", end="2024-03-31", freq="D")
    df_cal = pd.DataFrame({
        "Date": cal_dates,
        "Period": ["Calendar Q1"] * len(cal_dates),
        "Amount": range(len(cal_dates))
    })
    df_cal.to_excel(calendar_file, sheet_name="Data", index=False)

    fiscal_file = temp_excel_dir / "fiscal_q1.xlsx"
    fiscal_dates = pd.date_range(start="2024-02-01", end="2024-04-30", freq="D")
    df_fiscal = pd.DataFrame({
        "Date": fiscal_dates,
        "Period": ["Fiscal Q1"] * len(fiscal_dates),
        "Amount": range(1000, 1000 + len(fiscal_dates))
    })
    df_fiscal.to_excel(fiscal_file, sheet_name="Data", index=False)

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    cal_table = [t["table"] for t in tables["tables"] if "calendar" in t["table"]][0]
    fiscal_table = [t["table"] for t in tables["tables"] if "fiscal" in t["table"]][0]

    result = server.query(f'''
        SELECT Date
        FROM "{cal_table}"
        INTERSECT
        SELECT Date
        FROM "{fiscal_table}"
    ''')

    assert result["row_count"] > 0
