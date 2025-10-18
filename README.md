# mcp-server-excel-sql

Let Claude query your Excel files using SQL - no SQL knowledge required. Ask questions in plain English, Claude writes and executes the queries automatically.

## What It Does

**How it works:**
1. Point the server at your Excel files
2. Ask Claude questions in plain English
3. Claude writes SQL queries automatically
4. Get instant answers from your data

**Capabilities:**
- Each Excel sheet becomes a queryable SQL table
- Join data across multiple spreadsheets
- Clean messy data with YAML transformation rules
- Deploy for teams with concurrent access
- Support for complex queries (aggregations, window functions, CTEs)

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

**Requirements:**
- Python 3.11 or higher
- pipx ([installation guide](https://pipx.pypa.io/stable/installation/))

```bash
pipx install mcp-server-excel-sql
```

## Try It Now

Get running in 30 seconds with example data:

```bash
# Install
pipx install mcp-server-excel-sql

# Download examples (or use your own Excel files)
git clone https://github.com/your-repo/mcp-excel.git
cd mcp-excel

# Generate example financial data
python examples/create_finance_examples.py

# Start server
mcp-excel --path examples
```

Now the server is running with 10 financial Excel files loaded. Use Claude Desktop or query directly.

## Quick Start

### Command Line

```bash
# Single user (local development)
mcp-excel --path /path/to/excel/files

# With auto-refresh when files change
mcp-excel --path /path/to/excel/files --watch

# Multiple concurrent users (team deployment)
# See DEVELOPMENT.md for authentication and security
mcp-excel --path /path/to/excel/files --transport sse --port 8000
```

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "mcp-excel",
      "args": ["--path", "/Users/your-username/data/excel"]
    }
  }
}
```

**With transformation rules:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "mcp-excel",
      "args": [
        "--path", "/path/to/excel/files",
        "--overrides", "/path/to/config.yaml"
      ]
    }
  }
}
```

Restart Claude Desktop after saving the configuration.

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

The server exposes 4 MCP tools for working with Excel data:

**tool_list_tables**
- Lists all available tables loaded from Excel files
- Shows file path, sheet name, row count
- Call this first to discover your data

**tool_get_schema**
- Shows column names and types for a specific table
- Use after listing tables to understand structure

**tool_query**
- Execute SQL queries on your Excel data
- Read-only (no modifications allowed)
- Supports joins, aggregations, filtering

**tool_refresh**
- Reload data after Excel files have changed
- Automatic with `--watch` flag

## Examples

The repository includes example Excel files with financial data for a Malaysian coffeehouse chain.

**Generate examples:**
```bash
python examples/create_finance_examples.py
mcp-excel --path examples --overrides examples/finance_overrides.yaml
```

**Example queries:**
```sql
-- Total debits in general ledger
SELECT SUM(COALESCE(debit, 0)) as total_debits
FROM "examples.general_ledger.entries";

-- Revenue by region
SELECT region, SUM(revenue) as total_revenue
FROM "examples.revenue_by_segment.revenue"
GROUP BY region
ORDER BY total_revenue DESC;

-- Budget variance analysis
SELECT
  department,
  budget_amount,
  actual_amount,
  (actual_amount - budget_amount) as variance,
  ROUND(((actual_amount - budget_amount) / budget_amount * 100), 2) as variance_pct
FROM "examples.budget_vs_actuals.data"
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

## Data Transformation

Excel files often have messy formatting. Use transformation rules to clean data.

**What you can fix:**
- Skip header/footer rows
- Combine multi-row headers
- Filter out total/summary rows
- Rename columns
- Set data types (dates, decimals, etc.)
- Pivot wide tables to long format

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
      header_rows: 2                  # Combine multi-row headers
      drop_regex: "^Total:"           # Remove rows starting with "Total:"
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
mcp-excel --path /data --overrides config.yaml
```

See `examples/finance_overrides.yaml` for complete real-world examples.
</details>

**Modes:**
- **RAW** (default): Loads Excel as-is with all columns as text, no headers
- **ASSISTED**: Applies transformation rules from YAML configuration

## CLI Options

```bash
mcp-excel [OPTIONS]
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
