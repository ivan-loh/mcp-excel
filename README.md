# mcp-server-excel-sql

[![PyPI version](https://badge.fury.io/py/mcp-server-excel-sql.svg)](https://pypi.org/project/mcp-server-excel-sql/)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![MCP Server](https://img.shields.io/badge/MCP-Server-blue.svg)](https://modelcontextprotocol.io)
[![Test](https://github.com/ivan-loh/mcp-excel/actions/workflows/test.yml/badge.svg)](https://github.com/ivan-loh/mcp-excel/actions)

Let Claude query your Excel and CSV files using SQL - no SQL knowledge required. Ask questions in plain English, Claude writes and executes the queries automatically.

## What It Does

**How it works:**
1. Point the server at your Excel/CSV files
2. Ask Claude questions in plain English
3. Claude writes SQL queries automatically
4. Get instant answers from your data

**Capabilities:**
- Each Excel sheet and CSV file becomes a queryable SQL table
- Join data across multiple files and formats (xlsx, xls, csv, tsv)
- Clean messy data with YAML transformation rules
- Deploy for teams with concurrent access
- Support for complex queries (aggregations, window functions, CTEs)

![Claude analyzing Excel budget data](examples/sample.png)

## Should You Use This?

**Great fit if you:**
- Work with Excel files under 100MB
- Want data insights without SQL knowledge
- Need to join multiple spreadsheets
- Use AI assistants (Claude writes the SQL for you)
- Prototype before building ETL pipelines

**Not the right tool if you:**
- Have files over 100MB (use database import instead)
- Need to modify Excel files (read-only)
- Need formulas/macros/VBA (values only)
- Building production data warehouse (prototyping only)

## Installation

**Install uv:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

That's it. No package installation needed - `uvx` runs the server on-demand.

## Try It Now

Test with example data:

```bash
# Clone examples
git clone https://github.com/ivan-loh/mcp-excel.git
cd mcp-excel

# Generate example financial data
python examples/finance/create_finance_examples.py

# Start server
uvx --from mcp-server-excel-sql mcp-excel --path examples/finance
```

Server runs with 10 financial Excel files loaded.

## Quick Start

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "finance-data": {
      "command": "uvx",
      "args": [
        "--from",
        "mcp-server-excel-sql",
        "mcp-excel",
        "--path",
        "/path/to/excel/files/"
      ]
    }
  }
}
```

**Customize:**
1. Replace `YOUR-USERNAME` with your username (run `whoami` in terminal)
2. Update the path to your Excel files location
3. Restart Claude Desktop

**Naming tip:** Use descriptive names like `finance-data`, `sales-reports`, or `excel` for easy reference in conversations.

### Command Line Testing

```bash
# Test with your files
uvx --from mcp-server-excel-sql mcp-excel --path /path/to/excel/files

# With auto-refresh
uvx --from mcp-server-excel-sql mcp-excel --path /path/to/files --watch
```

## Common Use Cases

**Financial Analysis**
- Join budget vs actuals across quarters
- AR aging analysis across regions
- Revenue trending from monthly reports

**Sales Reporting**
- Combine sales data from multiple territories
- Product performance across time periods
- Customer segmentation from CRM exports

**Operations**
- Inventory reconciliation from warehouse exports
- Vendor comparison from procurement files
- Project tracking from multiple PM spreadsheets

**Data Exploration**
- Quick SQL access to ad-hoc Excel exports
- Testing data quality before building pipelines
- Prototyping analytics without database setup

## How It Works

**You ask Claude in plain English:**
- "What's the total revenue by region?"
- "Show me customers with overdue invoices"
- "Compare Q1 budget vs actuals by department"

**Claude automatically:**
1. Writes the SQL query using the available tools
2. Executes the query against your Excel data
3. Returns formatted results

**No SQL knowledge required** - Claude handles query generation, table joins, and data formatting automatically.

## Available Tools

The server exposes 6 MCP tools for working with Excel data:

**tool_list_tables**
- Lists all available tables and views
- Shows file path, sheet name, row count for tables
- Shows SQL definition and row count for views
- Call this first to discover your data

**tool_get_schema**
- Shows column names and types for a table or view
- Use after listing tables to understand structure

**tool_query**
- Execute SQL queries on your Excel data
- Read-only (no modifications allowed)
- Supports joins, aggregations, filtering, CTEs

**tool_refresh**
- Reload data after Excel files have changed
- Automatic with `--watch` flag
- Supports incremental or full refresh

**tool_create_view**
- Create persistent SQL views for complex transformations
- Views automatically restored on server restart
- Useful for filtering, aggregation, and multi-table joins

**tool_drop_view**
- Delete a view and its persistent storage
- Cleanup for views no longer needed

## Examples

The repository includes example Excel files with financial data for a Malaysian coffeehouse chain.

**Generate examples:**
```bash
python examples/finance/create_finance_examples.py
uvx --from mcp-server-excel-sql mcp-excel --path examples/finance --overrides examples/finance/finance_overrides.yaml
```

**Example queries:**
```sql
-- Total debits in general ledger
SELECT SUM(COALESCE(debit, 0)) as total_debits
FROM "finance.general_ledger.entries";

-- Revenue by region
SELECT region, SUM(revenue) as total_revenue
FROM "finance.revenue_by_segment.revenue"
GROUP BY region
ORDER BY total_revenue DESC;

-- Budget variance analysis
SELECT
  department,
  budget_amount,
  actual_amount,
  (actual_amount - budget_amount) as variance,
  ROUND(((actual_amount - budget_amount) / budget_amount * 100), 2) as variance_pct
FROM "finance.budget_vs_actuals.data"
ORDER BY variance DESC;
```

**Included files:**
- General Ledger (MYR currency)
- Financial Statements
- Accounts Receivable Aging
- Revenue Analysis by Segment
- Budget vs Actuals
- Invoice Register
- Trial Balance
- Cash Flow Forecast

See `examples/README.md` for detailed query examples and usage patterns.

## Understanding Table Names

Excel files are converted to tables with this naming pattern:

**Format:** `<alias>.<filename>.<sheet>`

**Example:**
- File: `/data/sales/Q1-2024.xlsx`
- Sheet: `Summary`
- Table name: `sales.q12024.summary`

**Important:** Table names contain dots and must be quoted in SQL queries:

```sql
-- Correct
SELECT * FROM "sales.q12024.summary"

-- Wrong (will fail)
SELECT * FROM sales.q12024.summary
```

**Name cleaning:**
- Converted to lowercase
- Spaces become underscores
- Special characters removed
- Only `a-z`, `0-9`, `_`, `$` allowed

## System Views

Special tables for browsing loaded data:

**`<alias>.__files`** - File inventory
```sql
SELECT * FROM "sales.__files"
```
Shows: file paths, sheet count, total rows, modification time

**`<alias>.__tables`** - Table catalog
```sql
SELECT * FROM "sales.__tables"
```
Shows: table names, source file, sheet name, row count

## Persistent Views

Create reusable SQL views that persist across server restarts.

**What are views:**
- Named SQL queries stored on disk
- Automatically restored when server starts
- Useful for complex transformations, filtering, and joins
- Can reference other views

**Creating views:**
```sql
-- Filtering
CREATE VIEW high_value_sales AS
SELECT * FROM "sales.data.summary" WHERE amount > 1000

-- Aggregation
CREATE VIEW monthly_totals AS
SELECT date_trunc('month', date) as month, SUM(amount) as total
FROM "sales.data.summary"
GROUP BY 1

-- Multi-table joins
CREATE VIEW enriched_sales AS
SELECT s.*, p.category
FROM "sales.data.summary" s
JOIN "products.data.sheet" p ON s.product_id = p.id
```

**Managing views:**
- `tool_create_view(name, sql)` - Create or replace a view
- `tool_drop_view(name)` - Delete a view
- `tool_list_tables()` - Lists both tables and views
- Views stored as `.view_{name}` files in root directory

## Data Transformation

Excel files often have messy formatting. Use transformation rules to clean data.

**What you can fix:**
- Skip header/footer rows
- Combine multi-row headers (joined with "__" separator)
- Filter out total/summary rows with regex or conditions
- Rename columns
- Set data types (dates, decimals, etc.)
- Pivot wide tables to long format
- Specify cell ranges for Excel files
- Extract specific tables from multi-table sheets
- Filter rows based on multiple column conditions

**How it works:**
1. Create YAML configuration file
2. Specify transformations per file/sheet
3. Load with `--overrides config.yaml`

<details>
<summary>Show transformation example</summary>

Create `config.yaml`:

```yaml
sales.xlsx:
  sheet_overrides:
    Summary:
      skip_rows: 3                    # Skip header rows
      skip_footer: 2                  # Skip footer rows
      header_rows: 2                  # Combine multi-row headers (joined with "__")
      range: "A1:F100"                # Specify cell range (Excel only)
      drop_regex: "^Total:"           # Remove rows starting with "Total:"
      drop_conditions:                # Advanced multi-column filtering
        - column: "Status"
          regex: "^(DELETED|VOID)"
        - column: "Amount"
          is_null: false
      column_renames:
        "col_0": "region"             # Rename columns
      type_hints:
        amount: "DECIMAL(10,2)"       # Set column types
        date: "DATE"
      unpivot:                        # Pivot wide tables to long format
        id_vars: ["Region"]
        value_vars: ["Jan", "Feb", "Mar"]
        var_name: "Month"
        value_name: "Sales"
```

Run with overrides:
```bash
uvx --from mcp-server-excel-sql mcp-excel --path /data --overrides config.yaml
```

See `examples/finance/finance_overrides.yaml` for complete real-world examples.
</details>

**Modes:**
- **RAW** (default): Loads Excel as-is with all columns as text, no headers
- **ASSISTED**: Applies transformation rules from YAML configuration

## Auto-Detection Features

For complex Excel files, enable automatic structure detection to handle messy data without manual configuration.

**What auto-detection handles:**
- Merged cells (automatically unmerged and filled)
- Hidden rows and columns (filtered out)
- European number formats (1.234,56 → decimal values)
- Multiple tables on a single sheet
- Header row detection
- Metadata rows before data

**How to enable:**

```yaml
messy_report.xlsx:
  sheet_overrides:
    "Report":
      auto_detect: true              # Enable automatic detection
```

**Advanced auto-detection options:**

<details>
<summary>Show advanced auto-detection example</summary>

```yaml
financial_report.xlsx:
  sheet_overrides:
    "Monthly Report":
      # Enable auto-detection
      auto_detect: true

      # Merged cell handling (optional - auto-detected if present)
      merge_handling:
        strategy: fill                # fill | skip | error
        header_strategy: span         # span | fill | skip

      # Hidden rows (optional - auto-detected)
      include_hidden: false           # Exclude hidden rows/columns

      # Number locale (optional - auto-detected)
      locale:
        locale: de_DE                 # European format
        decimal_separator: ","
        thousands_separator: "."
        auto_detect: false            # Override auto-detection

      # Multi-table extraction (PHASE 3)
      extract_table: 0                # Extract first table only (if multiple detected)
      # OR
      # table_range: "A5:F50"         # Manually specify table range
      # Auto-detect creates separate views: sheet_table0, sheet_table1, etc.

      # Still works with manual overrides
      column_renames:
        "Betrag": "Amount"
      type_hints:
        Amount: "DECIMAL(10,2)"
```

</details>

**When to use auto_detect:**
- Files with merged cells in headers
- Reports with hidden calculation columns
- European-formatted financial data
- Sheets with multiple separate tables
- Complex layouts with title rows and metadata

**Limitation:** Auto-detection only works with .xlsx and .xlsm files (not .xls, .csv, .tsv)

## CLI Options

```bash
uvx --from mcp-server-excel-sql mcp-excel [OPTIONS]
```

**Options:**
- `--path` - Directory containing Excel files (default: current directory)
- `--overrides` - YAML configuration file for transformations
- `--watch` - Auto-refresh when files change
- `--transport` - Communication mode: `stdio`, `streamable-http`, `sse` (default: stdio)
- `--host` - Host for HTTP/SSE (default: 127.0.0.1)
- `--port` - Port for HTTP/SSE (default: 8000)
- `--require-auth` - Enable API key authentication (uses MCP_EXCEL_API_KEY env var)

## Additional Documentation

**Multi-user deployment, security, and development:**
See [DEVELOPMENT.md](DEVELOPMENT.md) for:
- Multi-user setup with authentication
- Security model and enforcement
- Architecture and design decisions
- Performance characteristics
- Testing and development workflow

**Finance examples:**
See [examples/README.md](examples/README.md) for detailed query examples and patterns.

## License

MIT
