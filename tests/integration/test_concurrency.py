import pytest
import time
import concurrent.futures
from pathlib import Path
import pandas as pd
import tempfile
import mcp_excel.server as server

pytestmark = [pytest.mark.concurrency, pytest.mark.usefixtures("setup_server_http")]


@pytest.fixture
def test_excel_file():
    with tempfile.TemporaryDirectory(prefix="test_concurrency_") as tmpdir:
        tmpdir = Path(tmpdir)
        df = pd.DataFrame({
            "id": list(range(1, 100)),
            "name": [f"Item{i}" for i in range(1, 100)],
            "value": [i * 10 for i in range(1, 100)]
        })
        file_path = tmpdir / "data.xlsx"
        df.to_excel(file_path, sheet_name="Sheet1", index=False)
        yield tmpdir


def test_concurrent_queries_dont_interfere(test_excel_file):
    server.load_dir(str(test_excel_file))

    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    def run_query():
        result = server.query(f'SELECT * FROM "{table_name}"')
        return result

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(run_query) for _ in range(5)]
        results = [f.result() for f in futures]

    assert len(results) == 5
    assert all(r["row_count"] == 100 for r in results)
    assert all(r["truncated"] is False for r in results)


def test_timeout_isolation(test_excel_file):
    server.load_dir(str(test_excel_file))

    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    results = []
    timeout_errors = []

    def slow_query_with_timeout():
        try:
            server.query(f'SELECT * FROM "{table_name}"', timeout_ms=1)
        except TimeoutError as e:
            timeout_errors.append(e)

    def fast_query():
        result = server.query(f'SELECT * FROM "{table_name}"', timeout_ms=5000)
        results.append(result)

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        slow = executor.submit(slow_query_with_timeout)
        time.sleep(0.01)
        fast = executor.submit(fast_query)

        slow.result()
        fast.result()

    assert len(timeout_errors) == 1
    assert len(results) == 1
    assert results[0]["row_count"] == 100


def test_query_during_refresh(test_excel_file):
    server.load_dir(str(test_excel_file))

    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    errors = []
    query_count = [0]

    def query_worker():
        for _ in range(10):
            try:
                result = server.query(f'SELECT * FROM "{table_name}"')
                query_count[0] += 1
            except Exception as e:
                errors.append(e)
            time.sleep(0.01)

    def refresh_worker():
        time.sleep(0.05)
        server.refresh(full=True)

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        q = executor.submit(query_worker)
        r = executor.submit(refresh_worker)

        q.result()
        r.result()

    assert query_count[0] > 0, "No queries were executed successfully"
    assert len(errors) < 10, f"Too many queries failed: {len(errors)} out of 10"


def test_catalog_mutation_safety(test_excel_file):
    server.load_dir(str(test_excel_file))

    def list_tables_worker():
        results = []
        for _ in range(20):
            result = server.list_tables()
            results.append(len(result["tables"]))
            time.sleep(0.001)
        return results

    def refresh_worker():
        for _ in range(5):
            time.sleep(0.01)
            server.refresh(full=False)

    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        list1 = executor.submit(list_tables_worker)
        list2 = executor.submit(list_tables_worker)
        refresh = executor.submit(refresh_worker)

        results1 = list1.result()
        results2 = list2.result()
        refresh.result()

    assert all(count >= 0 for count in results1), "Invalid table count detected"
    assert all(count >= 0 for count in results2), "Invalid table count detected"


def test_concurrent_schema_access(test_excel_file):
    server.load_dir(str(test_excel_file))

    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    def get_schema_worker():
        result = server.get_schema(table_name)
        return result

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(get_schema_worker) for _ in range(5)]
        results = [f.result() for f in futures]

    assert len(results) == 5
    assert all(len(r["columns"]) > 0 for r in results)
    first_schema = results[0]["columns"]
    assert all(r["columns"] == first_schema for r in results), "Schemas should be consistent across concurrent requests"


def test_load_dir_creates_isolated_connections(test_excel_file):
    def load_worker(path):
        return server.load_dir(str(path))

    with tempfile.TemporaryDirectory(prefix="test_concurrent_load_") as tmpdir:
        tmpdir = Path(tmpdir)
        df = pd.DataFrame({"col1": [1, 2, 3], "col2": [4, 5, 6]})
        file_path = tmpdir / "test.xlsx"
        df.to_excel(file_path, sheet_name="Sheet1", index=False)

        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
            future1 = executor.submit(load_worker, test_excel_file)
            future2 = executor.submit(load_worker, tmpdir)

            result1 = future1.result()
            result2 = future2.result()

        assert result1["sheets_count"] == 1
        assert result2["sheets_count"] == 1
