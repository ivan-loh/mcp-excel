# Test Suite

240 tests across 5 categories: unit, integration, regression, concurrency, stress.

## Test Files

| File | Category | Tests | Purpose |
|------|----------|-------|---------|
| `test_naming.py` | unit | 38 | Table name sanitization, collision handling, hierarchical naming |
| `test_loader.py` | unit | 28 | File loading, RAW/ASSISTED modes, transformations |
| `test_multi_table.py` | unit | 23 | Multi-table detection, extraction, blank row separation |
| `test_multi_table_edge_cases.py` | unit | 17 | Merged cells, hidden rows, formulas, wide tables |
| `test_format_handling.py` | unit | 15 | Format detection, handlers, data normalization |
| `test_views.py` | unit | 13 | View creation, management, queries |
| `test_performance.py` | unit | 12 | LRU cache operations, structure analysis caching |
| `test_drop_conditions.py` | unit | 12 | Multi-column filtering with regex, equals, is_null |
| `test_validation.py` | unit | 7 | Configuration validation, conflicting options |
| `test_auth.py` | unit | 1 | API key authentication middleware |
| `test_integration.py` | integration | 19 | Server workflows, query safety, refresh logic, **golden path** |
| `test_examples_validation.py` | integration | 17 | Real-world example files, encoding detection |
| `test_views_persistence.py` | integration | 7 | View persistence across server restarts |
| `test_watcher.py` | integration | 6 | File system watching, change detection, debouncing |
| `test_transport.py` | integration | 5 | HTTP/SSE transport, async operations |
| `test_issue_fixes.py` | regression | 10 | Bug fix verification (Issues 3, 4, 5) |
| `test_concurrency.py` | concurrency | 6 | Thread safety with HTTP mode |
| `test_stress_concurrency.py` | stress | 4 | Performance under high load |

## Test Breakdown by Category

- **Unit tests**: 153 tests (naming, loader, multi-table, format handling, auth, views, drop conditions, performance, validation)
- **Integration tests**: 54 tests (server workflows, examples, watcher, transport, view persistence, **golden path**)
- **Regression tests**: 10 tests (bug fixes)
- **Concurrency tests**: 6 tests (thread safety)
- **Stress tests**: 4 tests (performance)
- **View management**: 20 tests (13 core + 7 persistence)

## Fixtures (conftest.py)

| Fixture | Scope | Description |
|---------|-------|-------------|
| `temp_dir` | function | Temporary directory for test files |
| `temp_excel_dir` | function | Temporary directory with explicit cleanup |
| `sample_excel` | function | Simple Excel file (Name, Age, City columns) |
| `test_data_dir` | function | Pre-populated directory with 3 Excel files |
| `conn` | function | DuckDB in-memory connection |
| `loader` | function | ExcelLoader instance with TableRegistry |
| `setup_server` | function | Reset server state for STDIO mode |
| `setup_server_http` | function | Reset server state for HTTP mode |

## Test Markers

- `unit` - Fast, isolated component tests
- `integration` - End-to-end workflows and server operations
- `regression` - Bug fix verification tests
- `concurrency` - Thread safety with concurrent operations
- `stress` - Performance and load testing
- `asyncio` - Async/await operations

## Running Tests

```bash
# All tests
pytest tests/

# By category
pytest -m unit                   # unit tests only (fast)
pytest -m integration            # integration tests
pytest -m "not stress"           # exclude stress tests

# By file
pytest tests/test_naming.py      # naming tests only
pytest tests/test_multi_table.py # multi-table tests

# With coverage
pytest --cov=mcp_excel tests/

# Parallel execution
pytest -n auto tests/            # requires pytest-xdist
```

## Recent Changes

**2025-10-26**: Phase 4 + View Management complete
- Added `test_drop_conditions.py` (12 tests) - Multi-column filtering
- Added `test_validation.py` (7 tests) - Configuration validation
- Added `test_performance.py` (12 tests) - LRU cache performance
- Added `test_views.py` (13 tests) - View creation and management
- Added `test_views_persistence.py` (7 tests) - View persistence across restarts
- Test suite reorganization from earlier (naming, examples, multi-table consolidation)
- Total: 240 tests (from 189 pre-Phase 4)
