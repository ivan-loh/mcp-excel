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
from starlette.middleware import Middleware

from . import __version__
from .models import TableMeta, SheetOverride, LoadConfig
from .utils.naming import TableRegistry
from .loading.loader import ExcelLoader
from .utils.watcher import FileWatcher
from .utils.auth import APIKeyMiddleware, get_api_key_from_env
from .utils import log
from .exceptions import MCPExcelError, ExtensionError


mcp = FastMCP("mcp-server-excel-sql")

catalog: dict[str, TableMeta]             = {}
conn: Optional[duckdb.DuckDBPyConnection] = None
registry: Optional[TableRegistry]         = None
loader: Optional[ExcelLoader]             = None
load_configs: dict[str, LoadConfig]       = {}
watcher: Optional[FileWatcher]            = None
views: dict[str, dict]                    = {}

_catalog_lock           = threading.RLock()
_load_configs_lock      = threading.RLock()
_registry_lock          = threading.RLock()
_views_lock             = threading.RLock()
_db_path: Optional[str] = None
_use_http_mode          = False


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
        except duckdb.ExtensionException:
            pass
        except duckdb.IOException as e:
            log.error("extension_install_failed", error=str(e), reason="io_error")
            raise ExtensionError(
                "Failed to install DuckDB excel extension. Check disk space and permissions.",
                extension_name="excel",
                operation="INSTALL",
                data={"error": str(e)}
            )
        except Exception as e:
            log.error("extension_install_unexpected", error=str(e), error_type=type(e).__name__)
            raise ExtensionError(
                f"Unexpected error installing excel extension: {e}",
                extension_name="excel",
                operation="INSTALL",
                data={"error_type": type(e).__name__}
            )
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


def _get_view_file_path(root: Path, view_name: str) -> Path:
    return root / f".view_{view_name}"


def _validate_view_name(view_name: str) -> None:
    if not view_name:
        raise ValueError("View name cannot be empty")

    if "." in view_name:
        raise ValueError("View names cannot contain dots (reserved for system tables)")

    if view_name.startswith("_"):
        raise ValueError("View names cannot start with underscore (reserved)")

    if not re.match(r"^[a-z0-9_]+$", view_name, re.IGNORECASE):
        raise ValueError("View names can only contain letters, numbers, and underscores")


def _load_views_from_disk(root: Path) -> int:
    loaded_count = 0

    for view_file in root.glob(".view_*"):
        view_name = view_file.name[6:]

        try:
            sql = view_file.read_text().strip()

            if not sql:
                log.warn("view_file_empty", view=view_name, file=str(view_file))
                continue

            with get_connection() as conn:
                conn.execute(f'CREATE OR REPLACE VIEW "{view_name}" AS {sql}')

            with _views_lock:
                views[view_name] = {
                    "sql": sql,
                    "file": str(view_file),
                    "created_at": view_file.stat().st_mtime
                }

            loaded_count += 1
            log.info("view_loaded", view=view_name, file=str(view_file))

        except Exception as e:
            log.warn("view_load_failed", view=view_name, file=str(view_file), error=str(e))

    return loaded_count


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
    except duckdb.CatalogException:
        log.debug("temp_table_not_found", table=temp_table_name)
    except Exception as e:
        log.warn("temp_table_cleanup_failed", table=temp_table_name, error=str(e))

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
    include_glob = include_glob or ["**/*.xlsx", "**/*.xlsm", "**/*.xls", "**/*.csv", "**/*.tsv"]
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

                        table_metas = loader.load_sheet(file_path, relative_path,
                                                       sheet_name, alias, sheet_override)

                        for table_meta in table_metas:
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

    views_loaded = _load_views_from_disk(root)

    with _catalog_lock:
        tables_count = len([t for t in catalog if t.startswith(f"{alias}.")])

    result = {
        "alias": alias,
        "root": str(root),
        "files_count": files_loaded,
        "sheets_count": sheets_loaded,
        "tables_count": tables_count,
        "views_count": views_loaded,
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
                except duckdb.TransactionException as rollback_error:
                    log.debug("rollback_already_done", error=str(rollback_error))
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
                "source": "file",
                "file": table_meta.file,
                "relpath": table_meta.relpath,
                "sheet": table_meta.sheet,
                "mode": table_meta.mode,
                "est_rows": table_meta.est_rows,
            })

    view_list = []
    with _views_lock:
        for view_name, view_info in views.items():
            with get_connection() as conn:
                try:
                    row_count_result = conn.execute(f'SELECT COUNT(*) FROM "{view_name}"').fetchone()
                    est_rows = row_count_result[0] if row_count_result else 0
                except duckdb.CatalogException as e:
                    log.debug("view_count_catalog_error", view=view_name, error=str(e))
                    est_rows = 0
                except duckdb.BinderException as e:
                    log.warn("view_count_binder_error", view=view_name, error=str(e))
                    est_rows = 0
                except Exception as e:
                    log.error("view_count_unexpected", view=view_name, error=str(e))
                    est_rows = 0

            sql_preview = view_info["sql"][:100]
            if len(view_info["sql"]) > 100:
                sql_preview += "..."

            view_list.append({
                "name": view_name,
                "source": "view",
                "sql": sql_preview,
                "est_rows": est_rows,
                "file": view_info["file"]
            })

    return {
        "tables": tables,
        "views": view_list
    }


def get_schema(table_name: str) -> dict:
    is_view = False
    with _catalog_lock:
        if table_name not in catalog:
            with _views_lock:
                if table_name not in views:
                    raise ValueError(f"Table or view '{table_name}' not found")
                is_view = True

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
                    with _views_lock:
                        if table_name not in views:
                            raise ValueError(f"Table or view '{table_name}' not found (removed during request)")
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
                new_metas = loader.load_sheet(file_path, relative_path,
                                            table_meta.sheet, load_config.alias,
                                            sheet_override)

                for new_meta in new_metas:
                    with _catalog_lock:
                        catalog[new_meta.table_name] = new_meta

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


def create_view(view_name: str, sql: str) -> dict:
    _validate_view_name(view_name)

    with _catalog_lock:
        if view_name in catalog:
            raise ValueError(f"Name '{view_name}' conflicts with existing table")

    with _views_lock:
        if view_name in views:
            raise ValueError(f"View '{view_name}' already exists")

    with _load_configs_lock:
        if not load_configs:
            raise RuntimeError("No data loaded. Call load_dir first.")

        root_path = list(load_configs.values())[0].root

    sql_clean = sql.strip()
    if not sql_clean.upper().startswith("SELECT"):
        raise ValueError("View SQL must be a SELECT query")

    with get_connection() as conn:
        try:
            conn.execute(f'CREATE OR REPLACE VIEW "{view_name}" AS {sql_clean}')

            row_count_result = conn.execute(f'SELECT COUNT(*) FROM "{view_name}"').fetchone()
            est_rows = row_count_result[0] if row_count_result else 0

        except Exception as e:
            raise RuntimeError(f"Failed to create view: {e}")

    view_file = _get_view_file_path(root_path, view_name)
    view_file.write_text(sql_clean)

    with _views_lock:
        views[view_name] = {
            "sql": sql_clean,
            "file": str(view_file),
            "created_at": view_file.stat().st_mtime
        }

    log.info("view_created", view=view_name, file=str(view_file), rows=est_rows)

    return {
        "view_name": view_name,
        "est_rows": est_rows,
        "file": str(view_file),
        "created": True
    }


def drop_view(view_name: str) -> dict:
    with _views_lock:
        if view_name not in views:
            raise ValueError(f"View '{view_name}' does not exist")

        view_info = views[view_name]
        view_file = Path(view_info["file"])

    with get_connection() as conn:
        try:
            conn.execute(f'DROP VIEW IF EXISTS "{view_name}"')
        except Exception as e:
            log.warn("view_drop_failed_in_db", view=view_name, error=str(e))

    if view_file.exists():
        view_file.unlink()

    with _views_lock:
        del views[view_name]

    log.info("view_dropped", view=view_name, file=str(view_file))

    return {
        "view_name": view_name,
        "dropped": True
    }


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


def version() -> dict:
    return {"version": __version__}


@mcp.tool()
def tool_query(sql: str, max_rows: int = 10000, timeout_ms: int = 60000) -> dict:
    """
    Execute read-only SQL query against tables/views. Call tool_list_tables first.

    Table names need quotes: SELECT * FROM "examples.sales.summary"
    View names don't: SELECT * FROM high_value_sales

    Parameters:
    - sql: SELECT query (DuckDB with CTEs, window functions)
    - max_rows: Return limit, default 10000
    - timeout_ms: Query timeout, default 60000

    Returns:
    - columns: [{name, type}] - Column definitions
    - rows: [[]] - Results (capped at max_rows)
    - row_count: Number of rows returned
    - truncated: true if more rows exist beyond limit
    - execution_ms: Query duration

    Read-only security enforced.
    """
    return query(sql, max_rows, timeout_ms)


@mcp.tool()
def tool_list_tables(alias: str = None) -> dict:
    """
    Discover loaded tables and views. CALL FIRST.

    Tables: <alias>.<filename>.<sheet> (lowercase)
    Views: <view_name> (no dots)

    Parameters:
    - alias: Optional namespace filter (e.g., "examples")

    Returns:
    - tables: [{table, source, file, relpath, sheet, mode, est_rows}]
      - mode: "RAW" (unprocessed) or "ASSISTED" (YAML transforms)
    - views: [{name, source, sql, est_rows, file}]
      - sql: View definition preview (truncated at 100 chars)

    Use exact names in tool_query/tool_get_schema.
    """
    return list_tables(alias)


@mcp.tool()
def tool_get_schema(table: str) -> dict:
    """
    Get column names and types for a table or view.

    Call after tool_list_tables to inspect structure.

    Parameters:
    - table: Exact name from tool_list_tables

    Returns: {columns: [{name, type, nullable}]}
    """
    return get_schema(table)


@mcp.tool()
def tool_refresh(alias: str = None, full: bool = False) -> dict:
    """
    Reload tables from modified Excel files.

    Parameters:
    - alias: Namespace to refresh or None for all
    - full: false (default) = incremental, true = full rebuild

    Returns when full=false: {changed, total}
    Returns when full=true: {files_count, sheets_count, dropped, added}

    Views persist independently and auto-reference updated tables.
    """
    return refresh(alias, full)


@mcp.tool()
def tool_create_view(view_name: str, sql: str) -> dict:
    """
    Create persistent view from SELECT query.

    Views stored as .view_{name} files, auto-restored on restart.
    Replaces existing view (CREATE OR REPLACE).

    Parameters:
    - view_name: Alphanumeric and underscores only (no dots, no leading underscore)
    - sql: SELECT query only

    Returns: {view_name, est_rows, file, created}

    Valid names: "high_value_sales", "monthly_totals"
    Invalid: "sales.data" (dots), "_private" (underscore prefix)
    """
    return create_view(view_name, sql)


@mcp.tool()
def tool_drop_view(view_name: str) -> dict:
    """
    Delete view permanently. Cannot be undone.

    Parameters:
    - view_name: Name of view to delete

    Returns: {view_name, dropped}
    """
    return drop_view(view_name)


@mcp.tool()
def tool_version() -> dict:
    """
    Get the version of the mcp-excel server.

    Returns: {version}
    """
    return version()


@click.command()
@click.option("--path", default=".", help="Root directory with Excel files (default: current directory)")
@click.option("--overrides", type=click.Path(exists=True), help="YAML overrides file")
@click.option("--watch", is_flag=True, default=False, help="Watch for file changes and auto-refresh")
@click.option("--transport", default="stdio", type=click.Choice(["stdio", "streamable-http", "sse"]), help="MCP transport (default: stdio)")
@click.option("--host", default="127.0.0.1", help="Host for HTTP transports (default: 127.0.0.1)")
@click.option("--port", default=8000, type=int, help="Port for HTTP transports (default: 8000)")
@click.option("--require-auth", is_flag=True, default=False, help="Require API key authentication (uses MCP_EXCEL_API_KEY env var)")
def main(path: str, overrides: Optional[str], watch: bool, transport: str, host: str, port: int, require_auth: bool):
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
            middleware = []

            if require_auth:
                api_key = get_api_key_from_env()
                if not api_key:
                    raise ValueError("--require-auth enabled but MCP_EXCEL_API_KEY environment variable not set")
                middleware.append(Middleware(APIKeyMiddleware, api_key=api_key))
                log.info("auth_enabled", key_length=len(api_key))

            log.info("starting_http_server", transport=transport, host=host, port=port, auth_enabled=require_auth)
            mcp.run(transport=transport, host=host, port=port, middleware=middleware)
        else:
            if require_auth:
                log.warn("auth_ignored", reason="stdio_transport_does_not_support_auth")
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
