# mcp-server-excel-sql

MCP server exposing Excel files as SQL-queryable DuckDB views.

## Features

- SQL queries on Excel via DuckDB in-memory views
- RAW (all_varchar) or ASSISTED (sheet_overrides) modes
- Multi-row headers, type hints, unpivot, column renames
- Auto-refresh on file changes (--watch)
- Thread-safe timeout enforcement via conn.interrupt()

## Installation

```bash
pipx install mcp-server-excel-sql
```

## Usage

```bash
mcp-excel --path /data/excel --watch --overrides config.yaml
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

## Security

- Read-only: BEGIN TRANSACTION READ ONLY (DuckDB-enforced, blocks all write operations)
- Path-confined: Root path validation, no traversal
- Timeout: threading.Timer → conn.interrupt()
- Row limit: fetchmany(max_rows + 1)

## Development

```bash
pip install -e ".[dev]"
pytest --cov=mcp_excel tests/
python -m build
```

**Coverage**: 78% (567 statements, 126 missed)
**Tests**: 53 passing

## License

MIT
