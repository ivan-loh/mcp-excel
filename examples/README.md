# Examples

| File | Demonstrates |
|------|--------------|
| `clean_data.xlsx` | RAW mode |
| `messy_report.xlsx` | skip_rows, skip_footer, drop_regex |
| `multirow_header.xlsx` | Multi-row headers |
| `wide_format.xlsx` | Unpivot |
| `mixed_types.xlsx` | Type hints |
| `financial_report.xlsx` | Multi-sheet overrides |
| `inventory_report.xlsx` | Filtering, renaming |

## Usage

```bash
python examples/create_examples.py
mcp-excel --path examples --alias demo --overrides examples/examples_overrides.yaml --watch
```

## File Details

### clean_data.xlsx
RAW mode, no overrides needed.

### messy_report.xlsx
```yaml
skip_rows: 3
skip_footer: 3
drop_regex: "^Total:"
```

### multirow_header.xlsx
```yaml
header_rows: 2  # Merges to: Region, Q1__Sales, Q1__Units
```

### wide_format.xlsx
```yaml
unpivot:
  id_vars: ["Region", "Product"]
  value_vars: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
  var_name: "Month"
  value_name: "Sales"
```

### mixed_types.xlsx
```yaml
type_hints:
  salary: "DECIMAL(10,2)"
  hire_date: "DATE"
  is_active: "BOOL"
  years_experience: "INT"
```

### financial_report.xlsx
Multi-sheet with skip_rows, drop_regex, type_hints.

### inventory_report.xlsx
```yaml
skip_rows: 3
skip_footer: 3
drop_regex: "^(VOID|CANCELLED)"
column_renames:
  "SKU#": "sku"
type_hints:
  quantity: "INT"
```

## Queries

```sql
SELECT * FROM demo____tables;
SELECT * FROM demo__clean_data__orders WHERE status = 'Completed';
SELECT Region, Month, AVG(Sales) FROM demo__wide_format__monthlysales GROUP BY Region, Month;
```
