# mcp-server-excel-sql

MCP server exposing Excel files as SQL-queryable DuckDB views with thread-safe concurrent access.

## Features

- SQL queries on Excel via DuckDB views
- RAW (all_varchar) or ASSISTED (sheet_overrides) modes
- Multi-row headers, type hints, unpivot, column renames
- Auto-refresh on file changes (--watch)
- Thread-safe concurrent access (HTTP/SSE transports)
- Isolated query timeouts (per-connection interrupt)

## Installation

```bash
pipx install mcp-server-excel-sql
```

## Usage

### CLI

```bash
# STDIO mode (single-threaded)
mcp-excel --path /data/excel --watch --overrides config.yaml

# HTTP mode (concurrent, ~10 users)
mcp-excel --path /data/excel --transport streamable-http --port 8000
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

With overrides:
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

## MCP Tools

**tool_load_dir** - Load Excel directory into views
**tool_query** - Execute SELECT (read-only, timeout/limit enforced)
**tool_list_tables** - List views with metadata
**tool_get_schema** - DESCRIBE table
**tool_refresh** - Rescan filesystem (incremental or full)

## Table Naming

**Format**: `<alias>.<filepath>.<sheet>` (dot-separated, lowercase, sanitized)

**Sanitization**:
- Spaces → `_`
- Special chars → removed
- Allowed: `[a-z0-9_$]`

**Alias**: Auto-generated from directory name

**Examples**:
```
/data/sales/Q1-2024.xlsx → "sales.q12024.summary"
/reports/P&L (Final).xlsx → "reports.plfinal.sheet1"
```

**IMPORTANT**: Dots require quoted identifiers in SQL:
```sql
SELECT * FROM "sales.q12024.summary"  -- correct
SELECT * FROM sales.q12024.summary    -- fails (Catalog Error)
```

## System Views

`<alias>.__files` - File metadata (path, sheet_count, total_rows, mtime)
`<alias>.__tables` - Table metadata (table_name, file, sheet, mode, est_rows)

Query: `SELECT * FROM "sales.__files"`

## Modes

**RAW**: `read_xlsx(..., all_varchar=true, header=false)`
**ASSISTED**: Apply per-sheet overrides

```yaml
sales.xlsx:
  sheet_overrides:
    Summary:
      skip_rows: 3
      skip_footer: 2
      header_rows: 2
      drop_regex: "^Total:"
      column_renames:
        "col_0": "region"
      type_hints:
        amount: "DECIMAL(10,2)"
        date: "DATE"
      unpivot:
        id_vars: ["Region"]
        value_vars: ["Jan", "Feb"]
        var_name: "Month"
        value_name: "Sales"
```

## Concurrency & Thread Safety

### Design

**STDIO mode**: Single shared connection (single-threaded)
**HTTP/SSE mode**: Isolated connections per request (concurrent)

Supports ~10-20 concurrent users with <1ms overhead per request.

### Thread Safety Features

1. **Isolated Connections** - Each HTTP request gets its own DuckDB connection
2. **Protected Catalog** - RLock guards all metadata dictionary access
3. **Per-Connection Timeouts** - Query interrupts don't affect other users
4. **Deadlock-Free** - Sequential lock acquisition (never nested)

### Critical Fixes

**Problem 1: Timeout Interference**
- Before: One user's timeout killed all concurrent queries
- After: Isolated connections ensure timeouts only affect their own query

**Problem 2: Catalog Race**
- Before: Concurrent refresh could corrupt table metadata
- After: RLock protects all catalog reads/writes

**Problem 3: Dictionary Mutation**
- Before: Python dict operations not atomic (GIL can be released)
- After: All dict mutations locked to prevent corruption

### Performance

- STDIO mode: 0% overhead (no changes to single-threaded path)
- HTTP mode: ~0.6ms per request (connection + lock overhead)
- Tested: 20 concurrent workers, 100+ operations

### Testing

**73/73 tests pass** including:
- 6 concurrency tests (parallel queries, timeout isolation, refresh safety)
- 4 stress tests (100+ concurrent operations, memory leak detection)

## Security

- Read-only: BEGIN TRANSACTION READ ONLY (DuckDB-enforced, blocks all write operations)
- Path-confined: Root path validation, no traversal
- Timeout: threading.Timer → conn.interrupt() (isolated per connection)
- Row limit: fetchmany(max_rows + 1)

## Examples

Finance examples for **Kopitiam Kita Sdn Bhd**, a Malaysian coffeehouse chain:

```bash
# Generate example files
python examples/create_finance_examples.py

# Load and query
mcp-excel --path examples --alias finance --overrides examples/finance_overrides.yaml

# Query examples
SELECT SUM(COALESCE(debit, 0)) as total_debits
FROM "finance.general_ledger.entries";

SELECT region, SUM(revenue) as total_revenue
FROM "finance.revenue_by_segment.revenue"
GROUP BY region ORDER BY total_revenue DESC;
```

**Includes**:
- 10 Excel files with 3,100+ financial records (MYR currency)
- General Ledger, Financial Statements, AR Aging, Revenue Analysis
- Budget Variance, Invoices, Trial Balance, Cash Flow Forecast
- Complete prompt chain sequences in `examples/README.md`

See `examples/README.md` for detailed usage and SQL query patterns.

## Development

```bash
pip install -e ".[dev]"
pytest --cov=mcp_excel tests/
python -m build
```

**Tests**: 73 passing (11 unit + 47 integration + 10 regression + 6 concurrency + 4 stress)
**Coverage**: Comprehensive test coverage including race condition scenarios

## Architecture Notes

### Transport Modes

**STDIO** (default):
- Single shared in-memory DuckDB connection
- Single-threaded (no concurrency concerns)
- Ideal for: Claude Desktop, local CLI usage

**HTTP/SSE** (--transport streamable-http):
- Persistent DuckDB file (temp, auto-cleanup)
- Isolated connection per request
- RLock-protected shared state (catalog, load_configs)
- Ideal for: Multiple concurrent users, web APIs

### State Management

**Global State** (thread-safe via locks):
- `catalog` - Table metadata dict (protected by `_catalog_lock`)
- `load_configs` - Directory configurations (protected by `_load_configs_lock`)
- `registry` - Name collision tracking (internal lock)

**Per-Request State** (isolated):
- DuckDB connection (HTTP mode only)
- Query timeout timer
- Transaction state

### Scale Limits

- Current design: Optimized for ~10 concurrent users
- Tested with: 20 concurrent workers
- Upgrade path: If >20 users needed, consider connection pooling

## License

MIT
