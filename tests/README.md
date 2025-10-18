# Test Suite

73 tests across 5 categories: unit, integration, regression, concurrency, stress.

## Test Files

| File | Category | Tests | Purpose |
|------|----------|-------|---------|
| `test_naming.py` | unit | 9 | Table name sanitization and collision handling |
| `test_loader.py` | unit | 10 | Excel loading, transformations (RAW/ASSISTED modes) |
| `test_watcher.py` | integration | 6 | File system watching and change detection |
| `test_integration.py` | integration | 18 | Server workflows, query safety, refresh logic |
| `test_examples.py` | integration | 5 | Real example files validation |
| `test_transport.py` | integration | 5 | HTTP transport and async operations |
| `test_issue_fixes.py` | regression | 10 | Bug fix verification (Issues 3, 4, 5) |
| `test_concurrency.py` | concurrency | 6 | Thread safety with HTTP mode |
| `test_stress_concurrency.py` | stress | 4 | Performance under high load |

## Fixtures (conftest.py)

| Fixture | Scope | Description |
|---------|-------|-------------|
| `temp_dir` | function | Temporary directory |
| `temp_excel_dir` | function | Temporary directory with cleanup |
| `test_data_dir` | function | Pre-populated test data (3 Excel files) |
| `sample_excel` | function | Simple Excel file (Name, Age, City) |
| `conn` | function | DuckDB in-memory connection |
| `loader` | function | ExcelLoader with registry |
| `setup_server` | function | Server state reset (autouse) |
| `setup_server_http` | function | Server in HTTP mode (autouse) |

## Markers

- `unit` - Fast, isolated component tests
- `integration` - End-to-end workflows
- `regression` - Bug fix verification
- `concurrency` - Thread safety
- `stress` - Performance under load
- `asyncio` - Async operations

## Running Tests

```bash
pytest tests/                    # all tests
pytest -m unit                   # unit tests only
pytest -m "not stress"           # exclude stress tests
pytest --cov=mcp_excel tests/    # with coverage
```
