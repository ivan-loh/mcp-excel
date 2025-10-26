# Examples

Real-world Excel datasets demonstrating mcp-server-excel-sql capabilities across different industries and use cases.

## Available Datasets

### Finance Examples
**[examples/finance/](finance/)** - Financial analysis for Kopitiam Kita Sdn Bhd, a Malaysian coffeehouse chain

**Business Context:** Traditional kopitiam with outlets across 5 Malaysian cities
**Data:** 3,100+ financial records across 10 Excel files (MYR currency)
**Use Cases:**
- General ledger transaction analysis
- Accounts receivable aging and collections
- Revenue analysis by segment/region/product
- Budget variance tracking
- Invoice and payment tracking
- Financial statements and KPIs

**Files:**
- General Ledger (1,000 transactions)
- Financial Statements (P&L, Balance Sheet, Cash Flow)
- Accounts Receivable Aging (300 invoices)
- Revenue by Segment (1,020 records)
- Budget vs Actuals (60 department comparisons)
- Invoice Register (500 invoices)
- Trial Balance, Cash Flow Forecast, Expense Reports, Financial Ratios

**Quick Start:**
```bash
python examples/finance/create_finance_examples.py
uvx --from mcp-server-excel-sql mcp-excel --path examples/finance --overrides examples/finance/finance_overrides.yaml
```

**[Full Documentation →](finance/)**

---

### CNC Manufacturing Examples
**[examples/cnc/](cnc/)** - CNC machining operations for Tech Holdings Berhad, a Malaysian SME manufacturer

**Business Context:** CNC division with 52 employees, 18 machines across 3 plants in Penang
**Data:** 7,500+ operational records across 4 quarters (FY2025, RM currency)
**Use Cases:**
- Prove-out failure analysis and prevention
- Scrap and rework tracking
- Job costing and margin analysis
- Machine utilization and downtime
- Quality performance trending
- Customer churn analysis

**Files:**
- Job Orders (660 production orders)
- Program Validation (806 execution runs)
- Quality Inspections (3,607 measurements)
- Cost Analysis (job costing with variance)
- Scrap/Rework (64 events with root causes)
- Material Inventory, Tooling Management, Machine Downtime, Production Schedule, and more

**Quick Start:**
```bash
cd examples/cnc
python create_cnc_examples.py
uvx --from mcp-server-excel-sql mcp-excel --path examples/cnc --watch
```

**[Full Documentation →](cnc/)**

---

## Other Example Files

Located in the root examples directory:

**quarterly_comparison.xlsx**
- Multi-quarter comparison data
- Demonstrates cross-sheet analysis

**CSV Files (various encodings):**
- `employees_latin1.csv` - Latin-1 encoding handling
- `products_utf8_bom.csv` - UTF-8 with BOM
- `sales_european.csv` - European number formats (1.234,56)

These files demonstrate CSV format handling and encoding detection.

## General Usage Pattern

### 1. Generate Example Data
```bash
python examples/finance/create_finance_examples.py
cd examples/cnc && python create_cnc_examples.py
```

### 2. Start Server
```bash
# Finance data with type hints and transformations
uvx --from mcp-server-excel-sql mcp-excel --path examples/finance --overrides examples/finance/finance_overrides.yaml

# CNC data with auto-refresh
uvx --from mcp-server-excel-sql mcp-excel --path examples/cnc --watch

# All examples (finance + cnc + other files)
uvx --from mcp-server-excel-sql mcp-excel --path examples
```

### 3. Explore with Claude
Ask questions in plain English:
- "What's the total revenue by region?"
- "Show me customers with overdue invoices"
- "Which CNC jobs had the highest scrap rate?"

Claude writes SQL queries automatically using the available MCP tools.

## What Examples Demonstrate

### Data Transformations
**Finance examples show:**
- Multi-row headers (trial_balance.xlsx)
- Wide format pivoting (cash_flow_forecast.xlsx)
- Multi-sheet workbooks (financial_statements.xlsx)
- Type hints for dates and decimals
- Skip rows and footer handling

**CNC examples show:**
- Multi-table detection on single sheets
- Complex joins across operational data
- Time-series analysis (weekly, quarterly)
- Quality metrics with tolerances
- Business logic (cost variance, OEE calculations)

### Query Patterns
**Cross-table joins:**
- Revenue vs AR correlation
- Job profitability with quality metrics
- Customer performance trending

**Aggregations:**
- Revenue by segment/region/product
- Budget variance by department
- Scrap analysis by root cause

**Time-series:**
- Monthly revenue trends
- Quarterly performance comparison
- Weekly production scheduling

### Real-World Data Characteristics
- Incomplete records (missing payment dates, optional fields)
- Calculated columns (variances, margins, aging buckets)
- Referential integrity (customer names across tables)
- Mixed data quality (requiring transformation rules)
- Multiple encodings and formats

## Documentation Structure

```
examples/
├── README.md                           # This file - general overview
├── finance/
│   ├── README.md                       # Finance-specific documentation
│   ├── create_finance_examples.py      # Data generator
│   ├── finance_overrides.yaml          # Transformation rules
│   └── *.xlsx                          # 10 financial data files
├── cnc/
│   ├── README.md                       # CNC-specific documentation
│   ├── create_cnc_examples.py          # Data generator
│   └── *.xlsx                          # 18+ operational data files
└── *.xlsx, *.csv                       # Misc example files
```

## Next Steps

- **Finance analysis:** See [finance/README.md](finance/) for detailed query examples and prompt sequences
- **CNC operations:** See [cnc/README.md](cnc/) for manufacturing-specific analysis patterns
- **Development:** See [DEVELOPMENT.md](../DEVELOPMENT.md) for server architecture and deployment
- **Main documentation:** See [README.md](../README.md) for installation and setup
