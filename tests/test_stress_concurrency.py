import pytest
import time
import concurrent.futures
from pathlib import Path
import pandas as pd
import tempfile
import mcp_excel.server as server

pytestmark = [pytest.mark.stress, pytest.mark.usefixtures("setup_server_http")]


@pytest.fixture
def large_dataset():
    with tempfile.TemporaryDirectory(prefix="test_stress_") as tmpdir:
        tmpdir = Path(tmpdir)

        for i in range(5):
            df = pd.DataFrame({
                "id": list(range(1000)),
                "name": [f"Item{j}" for j in range(1000)],
                "value": [j * 10 for j in range(1000)],
                "category": [f"Cat{j % 10}" for j in range(1000)]
            })
            file_path = tmpdir / f"data_{i}.xlsx"
            df.to_excel(file_path, sheet_name="Sheet1", index=False)
        yield tmpdir


def test_stress_concurrent_operations(large_dataset):
    server.load_dir(str(large_dataset))

    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    errors = []
    successful_queries = [0]
    successful_schemas = [0]
    successful_lists = [0]

    def query_worker():
        for _ in range(20):
            try:
                result = server.query(f'SELECT * FROM "{table_name}" LIMIT 100')
                successful_queries[0] += 1
            except Exception as e:
                errors.append(("query", str(e)))
            time.sleep(0.001)

    def schema_worker():
        for _ in range(20):
            try:
                result = server.get_schema(table_name)
                successful_schemas[0] += 1
            except Exception as e:
                errors.append(("schema", str(e)))
            time.sleep(0.001)

    def list_worker():
        for _ in range(20):
            try:
                result = server.list_tables()
                successful_lists[0] += 1
            except Exception as e:
                errors.append(("list", str(e)))
            time.sleep(0.001)

    def refresh_worker():
        time.sleep(0.05)
        for _ in range(3):
            try:
                server.refresh(full=False)
            except Exception as e:
                errors.append(("refresh", str(e)))
            time.sleep(0.03)

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = []
        futures.extend([executor.submit(query_worker) for _ in range(3)])
        futures.extend([executor.submit(schema_worker) for _ in range(2)])
        futures.extend([executor.submit(list_worker) for _ in range(2)])
        futures.append(executor.submit(refresh_worker))

        for f in futures:
            f.result()

    print(f"Successful queries: {successful_queries[0]}")
    print(f"Successful schemas: {successful_schemas[0]}")
    print(f"Successful lists: {successful_lists[0]}")
    print(f"Errors: {len(errors)}")

    assert successful_queries[0] > 50, f"Too few successful queries: {successful_queries[0]}"
    assert successful_schemas[0] > 30, f"Too few successful schemas: {successful_schemas[0]}"
    assert successful_lists[0] > 30, f"Too few successful lists: {successful_lists[0]}"

    assert len(errors) < 20, f"Too many errors: {len(errors)}"


def test_stress_catalog_consistency(large_dataset):
    server.load_dir(str(large_dataset))

    initial_catalog_size = len(server.list_tables()["tables"])
    catalog_sizes = []

    def measure_catalog():
        for _ in range(50):
            with server._catalog_lock:
                size = len(server.catalog)
            catalog_sizes.append(size)
            time.sleep(0.001)

    def refresh_worker():
        for _ in range(10):
            server.refresh(full=False)
            time.sleep(0.01)

    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        m1 = executor.submit(measure_catalog)
        m2 = executor.submit(measure_catalog)
        r = executor.submit(refresh_worker)

        m1.result()
        m2.result()
        r.result()

    assert all(s >= 0 for s in catalog_sizes), "Negative catalog size detected!"
    assert all(s <= initial_catalog_size * 2 for s in catalog_sizes), "Catalog size exploded!"


def test_stress_no_memory_leaks(large_dataset):
    import tracemalloc
    tracemalloc.start()

    server.load_dir(str(large_dataset))
    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    snapshot1 = tracemalloc.take_snapshot()

    for i in range(100):
        try:
            server.query(f'SELECT * FROM "{table_name}" LIMIT 10')
        except:
            pass

    snapshot2 = tracemalloc.take_snapshot()

    top_stats = snapshot2.compare_to(snapshot1, 'lineno')
    total_diff = sum(stat.size_diff for stat in top_stats)

    tracemalloc.stop()

    assert total_diff < 10 * 1024 * 1024, f"Memory leaked: {total_diff / 1024 / 1024:.2f} MB"


def test_stress_connection_isolation(large_dataset):
    server.load_dir(str(large_dataset))
    tables = server.list_tables()
    table_name = tables["tables"][0]["table"]

    timeout_count = [0]
    success_count = [0]

    def timeout_query():
        try:
            server.query(f'SELECT * FROM "{table_name}"', timeout_ms=1)
        except TimeoutError:
            timeout_count[0] += 1
        except Exception:
            pass

    def normal_query():
        try:
            result = server.query(f'SELECT * FROM "{table_name}" LIMIT 10')
            success_count[0] += 1
        except Exception:
            pass

    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
        futures = []
        for _ in range(50):
            if _ % 5 == 0:
                futures.append(executor.submit(timeout_query))
            else:
                futures.append(executor.submit(normal_query))

        for f in futures:
            f.result()

    print(f"Timeouts: {timeout_count[0]}, Successes: {success_count[0]}")

    assert timeout_count[0] >= 1, f"Expected at least 1 timeout, got {timeout_count[0]}"
    assert success_count[0] >= 25, f"Timeouts interfered with normal queries: {success_count[0]}/40"
    assert success_count[0] + timeout_count[0] >= 35, f"Too many total failures: {success_count[0] + timeout_count[0]}/50"
