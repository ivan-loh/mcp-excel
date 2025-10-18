#!/usr/bin/env python3
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import random

output_dir = Path(__file__).parent
output_dir.mkdir(exist_ok=True)

print("Creating finance examples for Kopitiam Kita Sdn Bhd")
print("A Malaysian coffeehouse chain with outlets across Malaysia\n")

np.random.seed(42)
random.seed(42)

COMPANY_NAME = "Kopitiam Kita Sdn Bhd"
COMPANY_INFO = {
    "name": COMPANY_NAME,
    "address": "No. 123, Jalan Bukit Bintang, 55100 Kuala Lumpur",
    "registration": "202001234567 (1234567-X)",
    "fiscal_year": "2024"
}

def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))

start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 12, 31)

print("1. Creating general_ledger.xlsx (1000 transactions)...")
account_codes = {
    "1000": "Cash", "1100": "Accounts Receivable", "1200": "Inventory",
    "1500": "Fixed Assets", "1600": "Accumulated Depreciation",
    "2000": "Accounts Payable", "2100": "Accrued Expenses", "2500": "Long-term Debt",
    "3000": "Common Stock", "3100": "Retained Earnings",
    "4000": "Revenue - Products", "4100": "Revenue - Services",
    "5000": "Cost of Goods Sold", "6000": "Salaries Expense",
    "6100": "Rent Expense", "6200": "Utilities Expense", "6300": "Marketing Expense",
    "6400": "Depreciation Expense", "7000": "Interest Expense"
}

gl_entries = []
for i in range(1, 1001):
    date = random_date(start_date, end_date)
    account = random.choice(list(account_codes.keys()))
    debit = round(random.uniform(100, 50000), 2) if random.random() > 0.5 else 0
    credit = round(random.uniform(100, 50000), 2) if debit == 0 else 0

    descriptions = [
        "Outlet daily sales", "Coffee bean purchase", "Salary payment",
        "Rent payment - outlet", "Utilities payment", "Equipment purchase",
        "Marketing campaign", "Catering revenue", "Food supplies",
        "Professional fees", "Bank interest", "Tax payment"
    ]

    gl_entries.append({
        "entry_id": f"JE{i:06d}",
        "date": date.strftime("%Y-%m-%d"),
        "account_code": account,
        "account_name": account_codes[account],
        "description": random.choice(descriptions),
        "debit": debit if debit > 0 else None,
        "credit": credit if credit > 0 else None,
        "reference": f"REF{random.randint(1000, 9999)}",
        "posted_by": random.choice([
            "Fatimah binti Ibrahim", "Lee Wei Ming", "Ahmad bin Abdullah"
        ])
    })

gl_df = pd.DataFrame(gl_entries)
gl_df.to_excel(output_dir / "general_ledger.xlsx", sheet_name="Entries", index=False)

print("2. Creating financial_statements.xlsx (multi-sheet)...")
with pd.ExcelWriter(output_dir / "financial_statements.xlsx") as writer:
    income_statement_data = [
        [COMPANY_NAME],
        ["Income Statement - Consolidated"],
        ["For the Fiscal Year Ended 31 December 2024"],
        ["All amounts in Malaysian Ringgit (MYR)"],
        [""],
        ["Line Item", "Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024", "FY 2024"],
        ["Revenue - Food & Beverage", 2500000, 2750000, 2950000, 3200000, 11400000],
        ["Revenue - Catering Services", 850000, 920000, 1050000, 1180000, 4000000],
        ["Total Revenue", 3350000, 3670000, 4000000, 4380000, 15400000],
        ["Cost of Goods Sold", 1200000, 1350000, 1480000, 1600000, 5630000],
        ["Gross Profit", 2150000, 2320000, 2520000, 2780000, 9770000],
        ["Salaries & Wages", 450000, 480000, 520000, 560000, 2010000],
        ["Rent Expense - Outlets", 120000, 120000, 120000, 120000, 480000],
        ["Utilities", 35000, 38000, 42000, 45000, 160000],
        ["Marketing & Promotions", 180000, 210000, 245000, 280000, 915000],
        ["Depreciation", 95000, 95000, 95000, 95000, 380000],
        ["Total Operating Expenses", 880000, 943000, 1022000, 1100000, 3945000],
        ["Operating Income (EBIT)", 1270000, 1377000, 1498000, 1680000, 5825000],
        ["Interest Expense", 45000, 42000, 40000, 38000, 165000],
        ["Net Income", 1225000, 1335000, 1458000, 1642000, 5660000]
    ]
    income_df = pd.DataFrame(income_statement_data)
    income_df.to_excel(writer, sheet_name="Income Statement", index=False, header=False)

    balance_sheet = pd.DataFrame({
        "account": [
            "Cash and Cash Equivalents", "Accounts Receivable", "Inventory",
            "Prepaid Expenses", "Total Current Assets",
            "Property, Plant & Equipment", "Accumulated Depreciation",
            "Net PP&E", "Total Assets",
            "Accounts Payable", "Accrued Expenses", "Current Debt",
            "Total Current Liabilities", "Long-term Debt",
            "Total Liabilities", "Common Stock", "Retained Earnings",
            "Total Equity", "Total Liabilities & Equity"
        ],
        "amount": [
            3500000, 2800000, 1950000, 180000, 8430000,
            12500000, -3800000, 8700000, 17130000,
            1450000, 680000, 500000, 2630000, 4200000,
            6830000, 5000000, 5300000, 10300000, 17130000
        ],
        "category": [
            "Asset", "Asset", "Asset", "Asset", "Asset",
            "Asset", "Asset", "Asset", "Asset",
            "Liability", "Liability", "Liability", "Liability", "Liability",
            "Liability", "Equity", "Equity", "Equity", "Total"
        ]
    })
    balance_sheet.to_excel(writer, sheet_name="Balance Sheet", index=False)

    cash_flow = pd.DataFrame({
        "category": [
            "Net Income", "Depreciation", "Change in AR", "Change in Inventory",
            "Change in AP", "Cash from Operations",
            "Capital Expenditures", "Asset Disposals", "Cash from Investing",
            "Debt Issued", "Debt Repaid", "Dividends Paid", "Cash from Financing",
            "Net Change in Cash", "Beginning Cash", "Ending Cash"
        ],
        "amount": [
            5660000, 380000, -450000, -280000, 320000, 5630000,
            -2500000, 150000, -2350000,
            1000000, -800000, -1200000, -1000000,
            2280000, 1220000, 3500000
        ],
        "section": [
            "Operating", "Operating", "Operating", "Operating", "Operating", "Operating",
            "Investing", "Investing", "Investing",
            "Financing", "Financing", "Financing", "Financing",
            "Summary", "Summary", "Summary"
        ]
    })
    cash_flow.to_excel(writer, sheet_name="Cash Flow", index=False)

print("3. Creating accounts_receivable.xlsx (AR aging)...")
customers = [
    "Pavilion KL", "Suria KLCC", "Mid Valley Megamall", "The Gardens Mall",
    "1 Utama Shopping Centre", "Sunway Pyramid", "Gurney Plaza Penang",
    "Queensbay Mall", "Johor Bahru City Square", "Aeon Mall",
    "Tropicana City Mall", "IOI City Mall", "Paradigm Mall", "MyTOWN Shopping Centre",
    "IPC Shopping Centre", "Plaza Low Yat", "Berjaya Times Square",
    "Fahrenheit88", "Lot 10", "Starhill Gallery",
    "KL Sentral", "NU Sentral", "Bangsar Village", "Publika",
    "Empire Shopping Gallery", "The Curve", "Atria Shopping Gallery",
    "Sunway Velocity Mall", "Wangsa Walk Mall", "KLCC Convention Centre",
    "Putrajaya International Convention Centre", "KL Tower", "Menara Kuala Lumpur",
    "Petronas Twin Towers", "TRX Exchange 106", "Merdeka 118",
    "University of Malaya", "Universiti Teknologi Malaysia", "Sunway University",
    "Taylor's University", "HELP University", "Monash University Malaysia",
    "Maybank Tower", "CIMB Bank", "Public Bank", "RHB Bank",
    "Khazanah Nasional", "Axiata Group", "Telekom Malaysia", "TNB"
]
ar_data = []
for i in range(1, 301):
    invoice_date = random_date(datetime(2024, 1, 1), datetime(2024, 11, 30))
    amount = round(random.uniform(1000, 100000), 2)
    days_outstanding = (datetime(2024, 12, 31) - invoice_date).days

    ar_data.append({
        "invoice_number": f"INV-{i:06d}",
        "customer_name": random.choice(customers),
        "invoice_date": invoice_date.strftime("%Y-%m-%d"),
        "due_date": (invoice_date + timedelta(days=30)).strftime("%Y-%m-%d"),
        "amount": amount,
        "amount_paid": round(amount * random.uniform(0, 0.7), 2) if random.random() > 0.3 else 0,
        "days_outstanding": days_outstanding,
        "aging_bucket": "Current" if days_outstanding <= 30 else
                        "31-60 days" if days_outstanding <= 60 else
                        "61-90 days" if days_outstanding <= 90 else "90+ days"
    })

ar_df = pd.DataFrame(ar_data)
ar_df["balance"] = ar_df["amount"] - ar_df["amount_paid"]
ar_df.to_excel(output_dir / "accounts_receivable.xlsx", sheet_name="AR Aging", index=False)

print("4. Creating revenue_by_segment.xlsx (1000 rows)...")
products = [
    "Kopi Susu (Coffee with Milk)", "Teh Tarik (Pulled Tea)", "Kaya Toast",
    "Roti Bakar (Toasted Bread)", "Nasi Lemak", "Mee Goreng",
    "Catering - Corporate Events", "Catering - Weddings"
]
regions = [
    "Kuala Lumpur & Selangor", "Penang", "Johor Bahru",
    "Kota Kinabalu (Sabah)", "Kuching (Sarawak)"
]
segments = ["Walk-in Customers", "Corporate Clients", "Catering Services"]

revenue_data = []
for month in range(1, 13):
    for _ in range(85):
        revenue_data.append({
            "month": f"2024-{month:02d}",
            "product": random.choice(products),
            "region": random.choice(regions),
            "segment": random.choice(segments),
            "customer_count": random.randint(1, 20),
            "revenue": round(random.uniform(10000, 500000), 2),
            "cost": round(random.uniform(3000, 200000), 2)
        })

revenue_df = pd.DataFrame(revenue_data)
revenue_df["gross_profit"] = revenue_df["revenue"] - revenue_df["cost"]
revenue_df["margin_pct"] = (revenue_df["gross_profit"] / revenue_df["revenue"] * 100).round(2)
revenue_df.to_excel(output_dir / "revenue_by_segment.xlsx", sheet_name="Revenue", index=False)

print("5. Creating budget_vs_actuals.xlsx (variance analysis)...")
departments = [
    "Outlet Operations", "Kitchen & Food Prep", "Catering Services",
    "Marketing & Promotions", "Finance & Accounting", "Human Resources"
]
expense_categories = [
    "Salaries & Wages", "Employee Benefits", "Training & Development",
    "Rent - Outlets", "Utilities (Electric, Water, Gas)", "Raw Materials - Coffee & Tea",
    "Raw Materials - Food Ingredients", "Packaging & Supplies", "Equipment Maintenance",
    "Marketing & Advertising", "Delivery & Logistics", "Professional Fees"
]

budget_data = []
for dept in departments:
    for category in expense_categories:
        budget = round(random.uniform(10000, 200000), 2)
        actual = round(budget * random.uniform(0.7, 1.3), 2)
        budget_data.append({
            "department": dept,
            "expense_category": category,
            "budget": budget,
            "actual": actual,
            "variance": actual - budget,
            "variance_pct": round((actual - budget) / budget * 100, 2) if budget > 0 else 0
        })

budget_df = pd.DataFrame(budget_data)
budget_df.to_excel(output_dir / "budget_vs_actuals.xlsx", sheet_name="Analysis", index=False)

print("6. Creating invoice_register.xlsx (500 invoices)...")
invoice_data = []
for i in range(1, 501):
    invoice_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31))
    amount = round(random.uniform(500, 75000), 2)
    status = random.choice(["Paid", "Paid", "Paid", "Pending", "Overdue"])

    invoice_data.append({
        "invoice_id": f"INV{i:06d}",
        "customer_name": random.choice(customers),
        "invoice_date": invoice_date.strftime("%Y-%m-%d"),
        "due_date": (invoice_date + timedelta(days=30)).strftime("%Y-%m-%d"),
        "amount": amount,
        "tax_amount": round(amount * 0.08, 2),
        "total_amount": round(amount * 1.08, 2),
        "status": status,
        "payment_date": (invoice_date + timedelta(days=random.randint(1, 60))).strftime("%Y-%m-%d") if status == "Paid" else None,
        "payment_method": random.choice(["Wire Transfer", "ACH", "Check", "Credit Card"]) if status == "Paid" else None
    })

invoice_df = pd.DataFrame(invoice_data)
invoice_df.to_excel(output_dir / "invoice_register.xlsx", sheet_name="Invoices", index=False)

print("7. Creating trial_balance.xlsx (with messy headers)...")
trial_balance_data = [
    [COMPANY_NAME],
    ["Trial Balance"],
    ["As of 31 December 2024"],
    ["All amounts in Malaysian Ringgit (MYR)"],
    [""],
    ["Account Code", "Account Name", "Debit", "Credit"],
    ["1000", "Cash", "3500000.00", ""],
    ["1100", "Accounts Receivable", "2800000.00", ""],
    ["1200", "Inventory", "1950000.00", ""],
    ["1500", "Fixed Assets", "12500000.00", ""],
    ["1600", "Accumulated Depreciation", "", "3800000.00"],
    ["2000", "Accounts Payable", "", "1450000.00"],
    ["2100", "Accrued Expenses", "", "680000.00"],
    ["2500", "Long-term Debt", "", "4200000.00"],
    ["3000", "Common Stock", "", "5000000.00"],
    ["3100", "Retained Earnings", "", "300000.00"],
    ["4000", "Revenue - Products", "", "11400000.00"],
    ["4100", "Revenue - Services", "", "4000000.00"],
    ["5000", "Cost of Goods Sold", "5630000.00", ""],
    ["6000", "Salaries Expense", "2010000.00", ""],
    ["6100", "Rent Expense", "480000.00", ""],
    ["6300", "Marketing Expense", "915000.00", ""],
    ["6400", "Depreciation Expense", "380000.00", ""],
    ["7000", "Interest Expense", "165000.00", ""],
    ["", "TOTALS", "30330000.00", "30330000.00"],
    [""],
    ["Prepared by: Finance Department"],
    ["Date: 2025-01-05"]
]

trial_balance_df = pd.DataFrame(trial_balance_data)
trial_balance_df.to_excel(output_dir / "trial_balance.xlsx", sheet_name="Trial Balance",
                          index=False, header=False)

print("8. Creating cash_flow_forecast.xlsx (12-month projection, wide format)...")
forecast_categories = [
    "Cash Receipts - Collections", "Cash Receipts - New Sales",
    "Payroll", "Rent", "Utilities", "Vendor Payments",
    "Debt Service", "Capital Expenditures", "Tax Payments"
]

forecast_data = {"category": forecast_categories}
for month in range(1, 13):
    month_name = datetime(2025, month, 1).strftime("%b_%Y")
    forecast_data[month_name] = [
        round(random.uniform(800000, 1200000), 2),
        round(random.uniform(300000, 600000), 2),
        round(random.uniform(-500000, -400000), 2),
        -120000,
        round(random.uniform(-40000, -35000), 2),
        round(random.uniform(-300000, -500000), 2),
        -50000,
        round(random.uniform(-100000, -200000), 2) if month % 3 == 0 else 0,
        -150000 if month % 3 == 0 else 0
    ]

forecast_df = pd.DataFrame(forecast_data)
forecast_df.to_excel(output_dir / "cash_flow_forecast.xlsx", sheet_name="Monthly Forecast", index=False)

print("9. Creating expense_reports.xlsx (200 reports)...")
employees = [
    "Ahmad bin Abdullah", "Siti Nurhaliza binti Mohamed", "Lee Wei Ming", "Tan Ah Kow",
    "Raj Kumar a/l Subramaniam", "Fatimah binti Ibrahim", "Wong Mei Ling", "Kumar s/o Rajan",
    "Nurul Ain binti Hassan", "Lim Chee Keong", "Gopal a/l Krishnan", "Chen Li Ying",
    "Muhammad Faiz bin Ismail", "Sarah binti Ahmad", "Koh Boon Huat", "Rajeswari a/p Ganesh",
    "Azman bin Yusof", "Liew Sook Ching", "Ravi s/o Muthu", "Ng Pei San",
    "Zainab binti Mahmud", "Tan Kiat Seng", "Devi a/p Suresh", "Ong Kah Meng",
    "Nur Izzah binti Aziz", "Chan Siew Lan", "Ganesh a/l Raman", "Lim Pei Qi",
    "Hassan bin Omar", "Neo Wei Jie", "Siva s/o Kumar", "Teo Hui Min",
    "Halim bin Razak", "Chong Lai Fong", "Prem Kumar a/l Nathan", "Yap Su Lynn",
    "Ismail bin Ali", "Leong Mei Yee", "Selvam s/o Arumugam", "Chua Bee Lian",
    "Kamal bin Rashid", "Ang Li Na", "Murugan a/l Perumal", "Khoo Siew May",
    "Aziz bin Jaafar", "Tay Geok Lian", "Bala s/o Sundram", "Lau Pei Fang",
    "Rizal bin Hamzah", "Quek Siew Mei"
]
expense_types = [
    "Travel - Petrol/Toll", "Travel - Parking", "Client Meals & Entertainment",
    "Office Supplies", "Outlet Supplies", "Training & Courses",
    "Business Meals", "Transportation (Grab/Taxi)", "Equipment Purchases",
    "Staff Welfare Activities"
]

expense_reports = []
for i in range(1, 201):
    report_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31))
    expense_reports.append({
        "report_id": f"EXP{i:06d}",
        "employee_name": random.choice(employees),
        "department": random.choice(departments),
        "submission_date": report_date.strftime("%Y-%m-%d"),
        "expense_type": random.choice(expense_types),
        "amount": round(random.uniform(50, 5000), 2),
        "status": random.choice(["Approved", "Approved", "Pending", "Rejected"]),
        "approved_by": random.choice(["Manager A", "Manager B", "Manager C"]),
        "reimbursement_date": (report_date + timedelta(days=random.randint(7, 21))).strftime("%Y-%m-%d") if random.random() > 0.3 else None
    })

expense_df = pd.DataFrame(expense_reports)
expense_df.to_excel(output_dir / "expense_reports.xlsx", sheet_name="Reports", index=False)

print("10. Creating financial_ratios.xlsx (quarterly metrics)...")
quarters = ["Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"]
ratios_data = {
    "metric": [
        "Current Ratio", "Quick Ratio", "Debt-to-Equity",
        "Gross Profit Margin %", "Operating Margin %", "Net Profit Margin %",
        "Return on Assets %", "Return on Equity %",
        "Inventory Turnover", "AR Days Outstanding", "AP Days Outstanding",
        "Revenue Growth % YoY", "EBITDA ($M)"
    ],
    "Q1 2024": [3.2, 2.4, 0.66, 64.2, 37.9, 36.6, 7.8, 12.4, 5.2, 42, 38, 15.2, 1.365],
    "Q2 2024": [3.4, 2.6, 0.62, 63.2, 37.5, 36.4, 8.1, 13.1, 5.5, 40, 36, 18.5, 1.472],
    "Q3 2024": [3.5, 2.7, 0.58, 63.0, 37.5, 36.5, 8.9, 14.5, 5.8, 38, 35, 22.1, 1.593],
    "Q4 2024": [3.7, 2.9, 0.54, 63.5, 38.4, 37.5, 10.2, 16.8, 6.1, 35, 33, 25.8, 1.775]
}

ratios_df = pd.DataFrame(ratios_data)
ratios_df.to_excel(output_dir / "financial_ratios.xlsx", sheet_name="Quarterly Metrics", index=False)

print(f"\n✓ Created 10 finance example files for {COMPANY_NAME}")
print("\n" + "=" * 70)
print(f"Company: {COMPANY_NAME}")
print(f"Business: Malaysian coffeehouse chain with outlets across Malaysia")
print(f"Regions: Kuala Lumpur, Penang, Johor, Sabah, Sarawak")
print("=" * 70)
print("\nFiles Created:")
print("  1. general_ledger.xlsx          - 1,000 journal entries (MYR)")
print("  2. financial_statements.xlsx    - P&L, Balance Sheet, Cash Flow (3 sheets)")
print("  3. accounts_receivable.xlsx     - 300 invoices with AR aging")
print("  4. revenue_by_segment.xlsx      - 1,020 revenue records by product/region")
print("  5. budget_vs_actuals.xlsx       - 72 budget line items with variance")
print("  6. invoice_register.xlsx        - 500 invoices with payment tracking")
print("  7. trial_balance.xlsx           - Chart of accounts with messy headers")
print("  8. cash_flow_forecast.xlsx      - 12-month forecast 2025 (wide format)")
print("  9. expense_reports.xlsx         - 200 employee expense reports")
print(" 10. financial_ratios.xlsx        - Quarterly financial metrics/KPIs")
print(f"\nTotal: ~3,100+ financial records across all files (All in MYR)")
print("=" * 70)
