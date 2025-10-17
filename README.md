# mcp-server-excel

MCP server exposing Excel files as SQL-queryable tables via DuckDB.

[![Tests](https://img.shields.io/badge/tests-44%2F44%20passing-brightgreen)]()
[![Coverage](https://img.shields.io/badge/coverage-77%25-yellowgreen)]()
[![Python](https://img.shields.io/badge/python-3.11%2B-blue)]()

## Features

- SQL queries on Excel via DuckDB
- RAW (as-is) or ASSISTED (cleanup rules) modes
- Multi-row headers, type hints, unpivot transforms
- System views (`__files`, `__tables`)
- File watching with auto-refresh

## Installation

```bash
pip install mcp-server-excel
```

## Usage

```bash
mcp-server-excel --path /data/excel --watch --overrides config.yaml
```

## MCP Tools

**load_dir** - Load Excel files from directory
**query** - Execute SQL (read-only)
**list_tables** - List loaded tables
**get_schema** - Get table schema
**refresh** - Reload data

## Modes

**RAW**: Load as VARCHAR, no processing
**ASSISTED**: Apply overrides (skip_rows, type_hints, unpivot, etc.)

Example override:
```yaml
sales.xlsx:
  sheet_overrides:
    Summary:
      skip_rows: 3
      type_hints:
        amount: "DECIMAL(10,2)"
      unpivot:
        id_vars: ["Region"]
        value_vars: ["Jan", "Feb"]
```

## Table Naming

`<alias>__<relpath>__<sheet>` (lowercased, alphanumeric only)

Example: `sales/Q1.xlsx` → `excel__sales_q1__summary`

## System Views

`<alias>____files` - File metadata
`<alias>____tables` - Table metadata

## Security

Read-only, path-confined, SELECT-only queries, timeout/row limits

## Development

```bash
pip install -e ".[dev]"
pytest --cov=src tests/
python -m build
```

## License

MIT
