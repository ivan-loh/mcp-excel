#!/usr/bin/env python3
"""
Generate example Excel files demonstrating mcp-server-excel-sql capabilities
"""
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

output_dir = Path(__file__).parent
output_dir.mkdir(exist_ok=True)

print("Creating example Excel files...")

# 1. CLEAN DATA - No overrides needed (RAW mode)
print("1. Creating clean_data.xlsx...")
clean_df = pd.DataFrame({
    "order_id": [1001, 1002, 1003, 1004, 1005],
    "customer": ["Alice Corp", "Bob Inc", "Charlie Ltd", "Diana Co", "Eve LLC"],
    "amount": [1500.00, 2300.50, 890.25, 3200.00, 1750.75],
    "date": ["2024-01-15", "2024-01-16", "2024-01-17", "2024-01-18", "2024-01-19"],
    "status": ["Completed", "Completed", "Pending", "Completed", "Shipped"]
})
clean_df.to_excel(output_dir / "clean_data.xlsx", sheet_name="Orders", index=False)

# 2. MESSY REPORT - Demonstrates skip_rows, skip_footer, drop_regex
print("2. Creating messy_report.xlsx...")
messy_data = [
    ["Sales Report 2024"],
    ["Generated on: 2024-01-20"],
    [""],  # Empty row
    ["Region", "Product", "Units", "Revenue"],
    ["North", "Widget A", 150, 15000],
    ["North", "Widget B", 200, 20000],
    ["South", "Widget A", 180, 18000],
    ["South", "Widget B", 220, 22000],
    ["Total:", "", 750, 75000],
    [""],
    ["Notes: All values in USD"],
    ["Contact: sales@example.com"]
]
messy_df = pd.DataFrame(messy_data)
messy_df.to_excel(output_dir / "messy_report.xlsx", sheet_name="Summary",
                   index=False, header=False)

# 3. MULTI-ROW HEADER - Demonstrates header_rows=2
print("3. Creating multirow_header.xlsx...")
multirow_data = [
    ["", "Q1", "Q1", "Q2", "Q2"],
    ["Region", "Sales", "Units", "Sales", "Units"],
    ["North", 45000, 450, 52000, 520],
    ["South", 38000, 380, 41000, 410],
    ["East", 51000, 510, 58000, 580],
    ["West", 42000, 420, 47000, 470]
]
multirow_df = pd.DataFrame(multirow_data)
multirow_df.to_excel(output_dir / "multirow_header.xlsx", sheet_name="Quarterly",
                     index=False, header=False)

# 4. WIDE FORMAT - Demonstrates unpivot
print("4. Creating wide_format.xlsx...")
wide_df = pd.DataFrame({
    "Region": ["North", "South", "East", "West"],
    "Product": ["Widget A", "Widget A", "Widget A", "Widget A"],
    "Jan": [1200, 980, 1450, 1100],
    "Feb": [1350, 1020, 1520, 1180],
    "Mar": [1280, 1100, 1600, 1220],
    "Apr": [1400, 1150, 1680, 1290],
    "May": [1480, 1180, 1720, 1340],
    "Jun": [1520, 1220, 1800, 1380]
})
wide_df.to_excel(output_dir / "wide_format.xlsx", sheet_name="MonthlySales", index=False)

# 5. MIXED TYPES - Demonstrates type_hints
print("5. Creating mixed_types.xlsx...")
mixed_df = pd.DataFrame({
    "employee_id": ["E001", "E002", "E003", "E004", "E005"],
    "name": ["John Doe", "Jane Smith", "Bob Wilson", "Alice Brown", "Charlie Davis"],
    "salary": ["65000.00", "72000.50", "58000.00", "81000.75", "69000.25"],
    "hire_date": ["2020-03-15", "2019-07-22", "2021-01-10", "2018-11-05", "2020-09-18"],
    "department": ["Engineering", "Sales", "Marketing", "Engineering", "Sales"],
    "is_active": ["true", "true", "false", "true", "true"],
    "years_experience": ["8", "12", "5", "15", "9"]
})
mixed_df.to_excel(output_dir / "mixed_types.xlsx", sheet_name="Employees", index=False)

# 6. FINANCIAL REPORT - Comprehensive example with multiple features
print("6. Creating financial_report.xlsx...")
with pd.ExcelWriter(output_dir / "financial_report.xlsx") as writer:
    # Income Statement sheet with skip_rows and drop_regex
    income_data = [
        ["XYZ Corporation"],
        ["Income Statement"],
        ["For the Year Ended December 31, 2024"],
        [""],
        ["Account", "Amount"],
        ["Revenue", "1,250,000"],
        ["Cost of Goods Sold", "450,000"],
        ["Subtotal: Gross Profit", "800,000"],
        ["Operating Expenses", "320,000"],
        ["Depreciation", "85,000"],
        ["Subtotal: Operating Income", "395,000"],
        ["Interest Expense", "25,000"],
        ["Total: Net Income", "370,000"],
        [""],
        ["Notes: All amounts in USD"],
        ["Prepared by: Finance Department"]
    ]
    income_df = pd.DataFrame(income_data)
    income_df.to_excel(writer, sheet_name="Income", index=False, header=False)

    # Balance Sheet with proper structure
    balance_df = pd.DataFrame({
        "Account": ["Cash", "Accounts Receivable", "Inventory", "Property & Equipment",
                   "Accounts Payable", "Long-term Debt", "Equity"],
        "Amount": [185000, 245000, 320000, 890000, 165000, 450000, 1025000],
        "Category": ["Asset", "Asset", "Asset", "Asset", "Liability", "Liability", "Equity"]
    })
    balance_df.to_excel(writer, sheet_name="Balance", index=False)

# 7. INVENTORY REPORT - Shows complex filtering and renaming
print("7. Creating inventory_report.xlsx...")
inventory_data = [
    ["Inventory Status Report"],
    ["As of: 2024-01-20"],
    [""],
    ["SKU#", "Item Name", "Qty", "Unit $", "Total $", "Status"],
    ["WA-001", "Widget Alpha", 450, 25.50, 11475.00, "In Stock"],
    ["WB-002", "Widget Beta", 320, 32.00, 10240.00, "In Stock"],
    ["WC-003", "Widget Charlie", 0, 28.75, 0.00, "Out of Stock"],
    ["VOID-999", "Cancelled Item", 0, 0.00, 0.00, "VOID"],
    ["WD-004", "Widget Delta", 180, 45.00, 8100.00, "Low Stock"],
    ["WE-005", "Widget Echo", 520, 22.50, 11700.00, "In Stock"],
    ["CANCELLED-888", "Discontinued", 0, 0.00, 0.00, "CANCELLED"],
    [""],
    ["Summary Statistics:"],
    ["Total Items: 5"],
    ["Total Value: $41,515.00"]
]
inventory_df = pd.DataFrame(inventory_data)
inventory_df.to_excel(output_dir / "inventory_report.xlsx", sheet_name="Current",
                      index=False, header=False)

print(f"\n✓ Created 7 example Excel files in {output_dir}/")
print("\nExample files:")
print("  1. clean_data.xlsx       - Simple clean data (no overrides needed)")
print("  2. messy_report.xlsx     - Demonstrates skip_rows, skip_footer, drop_regex")
print("  3. multirow_header.xlsx  - Multi-row header with auto-merging")
print("  4. wide_format.xlsx      - Unpivot transformation")
print("  5. mixed_types.xlsx      - Type hints for proper casting")
print("  6. financial_report.xlsx - Comprehensive multi-sheet example")
print("  7. inventory_report.xlsx - Complex filtering and renaming")
print("\nNext: Create examples_overrides.yaml to demonstrate how to load these files")
