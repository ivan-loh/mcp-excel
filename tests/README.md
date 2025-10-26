# Test Suite

240 tests across 5 categories: unit, integration, regression, concurrency, stress.

## Directory Structure

Tests are organized to mirror the source code structure:

```
tests/
├── conftest.py                          # Shared fixtures
├── fixtures/                            # Test data files
│   ├── multi_table_test.xlsx
│   └── create_multi_table_fixture.py
│
├── loading/                             # mirrors mcp_excel/loading/
│   ├── test_loader.py                   # ExcelLoader tests (28 tests)
│   ├── test_analyzer.py                 # Structure analyzer, LRU cache (12 tests)
│   ├── test_multi_table.py              # Multi-table detection (23 tests)
│   ├── test_multi_table_edge_cases.py   # Edge cases: merged cells, formulas (17 tests)
│   └── formats/                         # mirrors mcp_excel/loading/formats/
│       └── test_format_handling.py      # Format detection, handlers (15 tests)
│
├── utils/                               # mirrors mcp_excel/utils/
│   ├── test_auth.py                     # API key authentication (1 test)
│   ├── test_naming.py                   # Table naming, collision handling (38 tests)
│   └── test_watcher.py                  # File watching, debouncing (6 tests)
│
├── integration/                         # End-to-end and system tests
│   ├── test_integration.py              # Server workflows, golden path (19 tests)
│   ├── test_examples_validation.py      # Real-world examples (17 tests)
│   ├── test_views.py                    # View management (13 tests)
│   ├── test_views_persistence.py        # View persistence (7 tests)
│   ├── test_concurrency.py              # Thread safety (6 tests)
│   ├── test_stress_concurrency.py       # High-load scenarios (4 tests)
│   ├── test_transport.py                # HTTP/SSE transport (5 tests)
│   └── test_issue_fixes.py              # Regression tests (10 tests)
│
├── test_drop_conditions.py              # Multi-column filtering (12 tests)
└── test_validation.py                   # Configuration validation (7 tests)
```

## Test Files by Module

### Loading Tests (96 tests)
Tests for data loading, structure analysis, and format handling

| File | Tests | Purpose |
|------|-------|---------|
| `loading/test_loader.py` | 28 | File loading, RAW/ASSISTED modes, transformations |
| `loading/test_multi_table.py` | 23 | Multi-table detection, extraction, blank row separation |
| `loading/test_multi_table_edge_cases.py` | 17 | Merged cells, hidden rows, formulas, wide tables |
| `loading/formats/test_format_handling.py` | 15 | Format detection, handlers, data normalization |
| `loading/test_analyzer.py` | 12 | LRU cache, structure analysis performance |

### Utils Tests (45 tests)
Tests for utility modules

| File | Tests | Purpose |
|------|-------|---------|
| `utils/test_naming.py` | 38 | Table name sanitization, collision handling, hierarchical naming |
| `utils/test_watcher.py` | 6 | File system watching, change detection, debouncing |
| `utils/test_auth.py` | 1 | API key authentication middleware |

### Integration Tests (81 tests)
End-to-end workflows and system tests

| File | Tests | Purpose |
|------|-------|---------|
| `integration/test_integration.py` | 19 | Server workflows, query safety, refresh logic, golden path |
| `integration/test_examples_validation.py` | 17 | Real-world example files, encoding detection |
| `integration/test_views.py` | 13 | View creation, management, queries |
| `integration/test_issue_fixes.py` | 10 | Bug fix verification (Issues 3, 4, 5) |
| `integration/test_views_persistence.py` | 7 | View persistence across server restarts |
| `integration/test_concurrency.py` | 6 | Thread safety with HTTP mode |
| `integration/test_transport.py` | 5 | HTTP/SSE transport, async operations |
| `integration/test_stress_concurrency.py` | 4 | Performance under high load |

### Root Level Tests (18 tests)
Tests for models and configuration

| File | Tests | Purpose |
|------|-------|---------|
| `test_drop_conditions.py` | 12 | Multi-column filtering with regex, equals, is_null |
| `test_validation.py` | 7 | Configuration validation, conflicting options |

## Test Breakdown by Category

- **Unit tests**: 159 tests (loading, utils, drop conditions, validation)
- **Integration tests**: 81 tests (server workflows, examples, transport, views, concurrency, stress)
- **Regression tests**: 10 tests (bug fixes in integration/)
- **Concurrency tests**: 6 tests (thread safety in integration/)
- **Stress tests**: 4 tests (performance in integration/)
- **View management**: 20 tests (13 core + 7 persistence in integration/)

**Note:** Some categories overlap (e.g., regression, concurrency, stress are subsets of integration tests)

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

# By module
pytest tests/loading/            # all loading tests
pytest tests/utils/              # all utils tests
pytest tests/integration/        # all integration tests

# By file
pytest tests/loading/test_loader.py
pytest tests/utils/test_naming.py

# With coverage
pytest --cov=mcp_excel tests/

# Parallel execution
pytest -n auto tests/            # requires pytest-xdist
```

## Recent Changes

**2025-10-26**: Test structure reorganization
- Reorganized tests to mirror source code structure
- Created subdirectories: `loading/`, `utils/`, `integration/`
- Moved format tests to `loading/formats/`
- Renamed `test_performance.py` → `loading/test_analyzer.py` (clearer purpose)
- Moved integration/regression tests to `integration/` subdirectory
- Maintained 240 total tests with no changes to test code
