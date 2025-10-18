import re
import time
import threading
import tempfile
import contextlib
from pathlib import Path
from typing import Optional

import click
import duckdb
import yaml
from fastmcp import FastMCP

from .types import TableMeta, SheetOverride, LoadConfig
from .naming import TableRegistry
from .loader import ExcelLoader
from .watcher import FileWatcher
from . import logging as log


mcp = FastMCP("mcp-server-excel-sql")

catalog: dict[str, TableMeta] = {}
conn: Optional[duckdb.DuckDBPyConnection] = None
registry: Optional[TableRegistry] = None
loader: Optional[ExcelLoader] = None
load_configs: dict[str, LoadConfig] = {}
watcher: Optional[FileWatcher] = None

_catalog_lock = threading.RLock()
_load_configs_lock = threading.RLock()
_registry_lock = threading.RLock()
_db_path: Optional[str] = None
_use_http_mode = False


def init_server(use_http_mode: bool = False):
    global conn, registry, loader, _db_path, _use_http_mode

    _use_http_mode = use_http_mode
    registry = TableRegistry()

    if use_http_mode:
        import os
        import time
        temp_dir = tempfile.gettempdir()
        _db_path = os.path.join(temp_dir, f"mcp_excel_{os.getpid()}_{int(time.time() * 1000000)}.duckdb")
    else:
        _db_path = ":memory:"
        if not conn:
            conn = duckdb.connect(":memory:")

    if use_http_mode:
        loader = None
    else:
        loader = ExcelLoader(conn, registry) if conn else None


@contextlib.contextmanager
def get_connection():
    global conn
    if _use_http_mode:
        local_conn = duckdb.connect(_db_path)
        try:
            local_conn.execute("INSTALL excel")
            local_conn.execute("LOAD excel")
        except:
            pass
        try:
            yield local_conn
        finally:
            local_conn.close()
    else:
        if conn is None:
            conn = duckdb.connect(_db_path)
        yield conn


def validate_root_path(user_path: str) -> Path:
    path = Path(user_path).resolve()

    if not path.exists():
        raise ValueError(f"Path {path} does not exist")

    if not path.is_dir():
        raise ValueError(f"Path {path} is not a directory")

    return path


def _generate_alias_from_path(root: Path) -> str:
    alias = root.name or "excel"
    alias = alias.lower()
    alias = re.sub(r"[^a-z0-9_]+", "_", alias)
    alias = re.sub(r"_+", "_", alias)
    alias = alias.strip("_")

    return alias if alias else "excel"


def _parse_sheet_override(override_dict: dict) -> SheetOverride:
    return SheetOverride(
        skip_rows=override_dict.get("skip_rows", 0),
        header_rows=override_dict.get("header_rows", 1),
        skip_footer=override_dict.get("skip_footer", 0),
        range=override_dict.get("range", ""),
        drop_regex=override_dict.get("drop_regex", ""),
        column_renames=override_dict.get("column_renames", {}),
        type_hints=override_dict.get("type_hints", {}),
        unpivot=override_dict.get("unpivot", {}),
    )


def _should_exclude_file(file_path: Path, exclude_patterns: list[str]) -> bool:
    for pattern in exclude_patterns:
        if file_path.match(pattern):
            return True
    return False


def _prepare_system_view_data(catalog_dict: dict, alias: str) -> tuple[dict, list]:
    files_data = {}
    tables_data = []

    for table_name, meta in catalog_dict.items():
        if not table_name.startswith(f"{alias}."):
            continue

        file_key = meta.file
        if file_key not in files_data:
            files_data[file_key] = {
                "file_path": meta.file,
                "relpath": meta.relpath,
                "sheet_count": 0,
                "total_rows": 0,
            }

        files_data[file_key]["sheet_count"] += 1
        files_data[file_key]["total_rows"] += meta.est_rows

        tables_data.append({
            "table_name": table_name,
            "file_path": meta.file,
            "relpath": meta.relpath,
            "sheet_name": meta.sheet,
            "mode": meta.mode,
            "est_rows": meta.est_rows,
            "mtime": meta.mtime,
        })

    return files_data, tables_data


def _register_dataframe_view(conn, temp_table_name: str, view_name: str, dataframe):
    import pandas as pd

    try:
        conn.unregister(temp_table_name)
    except:
        pass

    conn.register(temp_table_name, dataframe)
    conn.execute(f'CREATE OR REPLACE VIEW "{view_name}" AS SELECT * FROM {temp_table_name}')


def _create_system_views(alias: str):
    import pandas as pd

    with _catalog_lock:
        files_data, tables_data = _prepare_system_view_data(catalog, alias)

    files_view_name = f"{alias}.__files"
    tables_view_name = f"{alias}.__tables"
    temp_files_table = f"temp_files_{alias}"
    temp_tables_table = f"temp_tables_{alias}"

    with get_connection() as conn:
        try:
            if files_data:
                files_df = pd.DataFrame(list(files_data.values()))
                _register_dataframe_view(conn, temp_files_table, files_view_name, files_df)

            if tables_data:
                tables_df = pd.DataFrame(tables_data)
                _register_dataframe_view(conn, temp_tables_table, tables_view_name, tables_df)

            log.info("system_views_created", alias=alias,
                    files_view=files_view_name, tables_view=tables_view_name)
        except Exception as e:
            log.warn("system_views_failed", alias=alias, error=str(e))


def load_dir(
    path: str,
    alias: str = None,
    include_glob: list[str] = None,
    exclude_glob: list[str] = None,
    overrides: dict = None,
) -> dict:
    include_glob = include_glob or ["**/*.xlsx"]
    exclude_glob = exclude_glob or []
    overrides = overrides or {}

    root = validate_root_path(path)

    if alias is None:
        alias = _generate_alias_from_path(root)

    log.info("load_start", path=str(root), alias=alias, patterns=include_glob)

    files_loaded = 0
    sheets_loaded = 0
    total_rows = 0
    failed_files = []

    load_config = LoadConfig(
        root=root,
        alias=alias,
        include_glob=include_glob,
        exclude_glob=exclude_glob,
        overrides=overrides,
    )
    with _load_configs_lock:
        load_configs[alias] = load_config

    with get_connection() as conn:
        loader = ExcelLoader(conn, registry)

        for pattern in include_glob:
            for file_path in root.glob(pattern):
                if not file_path.is_file():
                    continue

                relative_path = str(file_path.relative_to(root))

                if _should_exclude_file(file_path, exclude_glob):
                    continue

                try:
                    sheet_names = loader.get_sheet_names(file_path)
                    file_overrides = overrides.get(relative_path, {})
                    sheet_overrides_dict = file_overrides.get("sheet_overrides", {})

                    for sheet_name in sheet_names:
                        sheet_override_dict = sheet_overrides_dict.get(sheet_name)
                        sheet_override = None

                        if sheet_override_dict:
                            sheet_override = _parse_sheet_override(sheet_override_dict)

                        table_meta = loader.load_sheet(file_path, relative_path,
                                                       sheet_name, alias, sheet_override)

                        with _catalog_lock:
                            catalog[table_meta.table_name] = table_meta

                        sheets_loaded += 1
                        total_rows += table_meta.est_rows

                        log.info("table_created", table=table_meta.table_name,
                                file=relative_path, sheet=sheet_name,
                                rows=table_meta.est_rows, mode=table_meta.mode)

                    files_loaded += 1

                except Exception as e:
                    error_msg = str(e)
                    log.warn("load_failed", file=relative_path, error=error_msg)
                    failed_files.append({"file": relative_path, "error": error_msg})

    _create_system_views(alias)

    with _catalog_lock:
        tables_count = len([t for t in catalog if t.startswith(f"{alias}.")])

    result = {
        "alias": alias,
        "root": str(root),
        "files_count": files_loaded,
        "sheets_count": sheets_loaded,
        "tables_count": tables_count,
        "rows_estimate": total_rows,
        "cache_mode": "none",
        "materialized": False,
    }

    if failed_files:
        result["failed"] = failed_files

    log.info("load_complete", alias=alias, files=files_loaded,
            sheets=sheets_loaded, rows=total_rows, failed=len(failed_files))

    return result


def query(
    sql: str,
    max_rows: int = 10000,
    timeout_ms: int = 60000,
) -> dict:
    start_time = time.time()
    interrupted = [False]
    transaction_started = [False]

    with get_connection() as conn:
        def timeout_handler():
            interrupted[0] = True
            try:
                conn.interrupt()
            except Exception as e:
                log.warn("interrupt_failed", error=str(e))

        timeout_seconds = timeout_ms / 1000.0
        timer = threading.Timer(timeout_seconds, timeout_handler)
        timer.start()

        query_result = None
        columns = None

        try:
            conn.execute("BEGIN TRANSACTION READ ONLY")
            transaction_started[0] = True

            cursor = conn.execute(sql)
            query_result = cursor.fetchmany(max_rows + 1)
            columns = [
                {"name": desc[0], "type": str(desc[1])}
                for desc in cursor.description
            ]

            conn.execute("COMMIT")
            transaction_started[0] = False

        except Exception as e:
            if interrupted[0]:
                execution_ms = int((time.time() - start_time) * 1000)
                log.warn("query_timeout", execution_ms=execution_ms, timeout_ms=timeout_ms)
                raise TimeoutError(f"Query exceeded {timeout_ms}ms timeout")

            log.error("query_failed", error=str(e), sql=sql[:100])
            raise RuntimeError(f"Query failed: {e}")

        finally:
            timer.cancel()

            if transaction_started[0]:
                try:
                    conn.execute("ROLLBACK")
                    log.info("transaction_rolled_back", reason="cleanup")
                except Exception as rollback_error:
                    log.warn("rollback_failed", error=str(rollback_error))

    execution_ms = int((time.time() - start_time) * 1000)
    is_truncated = len(query_result) > max_rows
    rows = query_result[:max_rows]

    log.info("query_executed", rows=len(rows),
            execution_ms=execution_ms, truncated=is_truncated)

    return {
        "columns": columns,
        "rows": rows,
        "row_count": len(rows),
        "truncated": is_truncated,
        "execution_ms": execution_ms,
    }


def list_tables(alias: str = None) -> dict:
    tables = []

    with _catalog_lock:
        for table_name, table_meta in catalog.items():
            if alias and not table_name.startswith(f"{alias}."):
                continue

            tables.append({
                "table": table_name,
                "file": table_meta.file,
                "relpath": table_meta.relpath,
                "sheet": table_meta.sheet,
                "mode": table_meta.mode,
                "est_rows": table_meta.est_rows,
            })

    return {"tables": tables}


def get_schema(table_name: str) -> dict:
    with _catalog_lock:
        if table_name not in catalog:
            raise ValueError(f"Table {table_name} not found")

    with get_connection() as conn:
        try:
            schema_result = conn.execute(f'DESCRIBE "{table_name}"').fetchall()
            columns = [
                {"name": row[0], "type": row[1], "nullable": row[2] == "YES"}
                for row in schema_result
            ]
        except Exception as e:
            with _catalog_lock:
                if table_name not in catalog:
                    raise ValueError(f"Table {table_name} not found (removed during request)")
            raise RuntimeError(f"Failed to get schema: {e}")

    return {"columns": columns}


def _refresh_full(alias: str) -> dict:
    with _catalog_lock:
        tables_to_drop = [
            table_name for table_name in catalog
            if alias is None or table_name.startswith(f"{alias}.")
        ]

    dropped_count = 0

    with get_connection() as conn:
        for table_name in tables_to_drop:
            try:
                conn.execute(f'DROP VIEW IF EXISTS "{table_name}"')

                with _catalog_lock:
                    del catalog[table_name]

                dropped_count += 1
            except Exception:
                pass

    added_count = 0
    files_count = 0
    sheets_count = 0

    with _load_configs_lock:
        load_config = load_configs.get(alias) if alias else None

    if load_config:
        result = load_dir(
            path=str(load_config.root),
            alias=alias,
            include_glob=load_config.include_glob,
            exclude_glob=load_config.exclude_glob,
            overrides=load_config.overrides,
        )
        added_count = result["tables_count"]
        files_count = result.get("files_count", 0)
        sheets_count = result.get("sheets_count", 0)

    _create_system_views(alias)

    return {
        "files_count": files_count,
        "sheets_count": sheets_count,
        "changed": 0,
        "dropped": dropped_count,
        "added": added_count,
    }


def _refresh_incremental(alias: str) -> dict:
    changed_count = 0

    with _catalog_lock:
        catalog_snapshot = list(catalog.items())

    with get_connection() as conn:
        loader = ExcelLoader(conn, registry)

        for table_name, table_meta in catalog_snapshot:
            if alias and not table_name.startswith(f"{alias}."):
                continue

            try:
                file_path = Path(table_meta.file)

                if not file_path.exists():
                    log.warn("refresh_file_missing", table=table_name, file=table_meta.file)
                    continue

                current_mtime = file_path.stat().st_mtime

                if current_mtime <= table_meta.mtime:
                    continue

                with _load_configs_lock:
                    load_config = load_configs.get(table_meta.alias)

                if not load_config:
                    log.warn("refresh_no_config", table=table_name, alias=table_meta.alias)
                    continue

                try:
                    relative_path = str(file_path.relative_to(load_config.root))
                except ValueError:
                    log.warn("refresh_path_outside_root", table=table_name,
                            file=str(file_path), root=str(load_config.root))
                    continue

                sheet_override_dict = (
                    load_config.overrides
                    .get(relative_path, {})
                    .get("sheet_overrides", {})
                    .get(table_meta.sheet)
                )

                sheet_override = None
                if sheet_override_dict:
                    sheet_override = SheetOverride(**sheet_override_dict)

                conn.execute(f'DROP VIEW IF EXISTS "{table_name}"')
                new_meta = loader.load_sheet(file_path, relative_path,
                                            table_meta.sheet, load_config.alias,
                                            sheet_override)

                with _catalog_lock:
                    catalog[table_name] = new_meta

                changed_count += 1

            except Exception as e:
                log.warn("refresh_failed", table=table_name, error=str(e))
                continue

    with _catalog_lock:
        total = len(catalog)

    return {
        "changed": changed_count,
        "total": total,
    }


def refresh(alias: str = None, full: bool = False) -> dict:
    if full:
        return _refresh_full(alias)
    else:
        return _refresh_incremental(alias)


def _on_file_change():
    log.info("file_change_detected", message="Auto-refreshing tables")

    try:
        with _load_configs_lock:
            aliases = list(load_configs.keys())

        for alias in aliases:
            result = refresh(alias=alias, full=False)
            log.info("auto_refresh_complete", alias=alias,
                    changed=result.get("changed", 0))
    except Exception as e:
        log.error("auto_refresh_failed", error=str(e))


def start_watching(path: Path, debounce_seconds: float = 1.0):
    global watcher

    if watcher:
        log.warn("file_watcher_already_running", path=str(path))
        return

    watcher = FileWatcher(path, _on_file_change, debounce_seconds)
    watcher.start()


def stop_watching():
    global watcher

    if not watcher:
        return

    watcher.stop()
    watcher = None


@mcp.tool()
def tool_load_dir(
    path: str,
    alias: str = None,
    include_glob: list[str] = None,
    exclude_glob: list[str] = None,
    overrides: dict = None,
) -> dict:
    """
    Load Excel files into DuckDB views.

    Tables: <alias>.<filepath>.<sheet> (dot-separated, requires quotes in SQL)
    Alias: Auto-generated from directory name (sanitized: lowercase, [a-z0-9_$])
    System views: <alias>.__files, <alias>.__tables

    Modes:
    - RAW: all_varchar=true, header=false
    - ASSISTED: apply sheet_overrides (skip_rows, header_rows, skip_footer, drop_regex,
                column_renames, type_hints, unpivot)

    Example: /data/sales/Q1.xlsx → "sales.q1.summary"
    Query: SELECT * FROM "sales.q1.summary"
    """
    return load_dir(path, alias, include_glob, exclude_glob, overrides)


@mcp.tool()
def tool_query(sql: str, max_rows: int = 10000, timeout_ms: int = 60000) -> dict:
    """
    Execute SQL query in read-only transaction.

    Table names contain dots, require double quotes: SELECT * FROM "alias.file.sheet"
    Read-only enforced: BEGIN TRANSACTION READ ONLY (blocks DDL/DML at database level)
    Limits: timeout_ms via conn.interrupt(), max_rows via fetchmany()
    """
    return query(sql, max_rows, timeout_ms)


@mcp.tool()
def tool_list_tables(alias: str = None) -> dict:
    """
    List loaded tables with metadata (file, sheet, mode, est_rows).

    Filter: alias="sales" returns tables matching "sales.*"
    """
    return list_tables(alias)


@mcp.tool()
def tool_get_schema(table: str) -> dict:
    """
    Get table schema via DESCRIBE.

    Returns: columns[{name, type, nullable}]
    """
    return get_schema(table)


@mcp.tool()
def tool_refresh(alias: str = None, full: bool = False) -> dict:
    """
    Refresh tables from filesystem.

    full=False: incremental (mtime check)
    full=True: drop and reload
    """
    return refresh(alias, full)


@click.command()
@click.option("--path", default=".", help="Root directory with Excel files (default: current directory)")
@click.option("--overrides", type=click.Path(exists=True), help="YAML overrides file")
@click.option("--watch", is_flag=True, default=False, help="Watch for file changes and auto-refresh")
@click.option("--transport", default="stdio", type=click.Choice(["stdio", "streamable-http", "sse"]), help="MCP transport (default: stdio)")
@click.option("--host", default="127.0.0.1", help="Host for HTTP transports (default: 127.0.0.1)")
@click.option("--port", default=8000, type=int, help="Port for HTTP transports (default: 8000)")
def main(path: str, overrides: Optional[str], watch: bool, transport: str, host: str, port: int):
    use_http_mode = transport in ["streamable-http", "sse"]
    init_server(use_http_mode=use_http_mode)

    overrides_dict = {}
    if overrides:
        with open(overrides, "r") as f:
            overrides_dict = yaml.safe_load(f) or {}

    root_path = Path(path).resolve()
    load_dir(path=str(root_path), overrides=overrides_dict)

    if watch:
        start_watching(root_path)
        log.info("watch_mode_enabled", path=str(root_path))

    try:
        if transport in ["streamable-http", "sse"]:
            log.info("starting_http_server", transport=transport, host=host, port=port)
            mcp.run(transport=transport, host=host, port=port)
        else:
            mcp.run(transport=transport)
    finally:
        if watch:
            stop_watching()

        if use_http_mode and _db_path and _db_path != ":memory:":
            try:
                Path(_db_path).unlink(missing_ok=True)
                log.info("temp_db_cleaned", path=_db_path)
            except Exception as e:
                log.warn("temp_db_cleanup_failed", path=_db_path, error=str(e))


if __name__ == "__main__":
    main()
