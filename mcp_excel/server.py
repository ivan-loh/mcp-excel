import time
import threading
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


def init_server():
    global conn, registry, loader
    if not conn:
        conn = duckdb.connect(":memory:")
        registry = TableRegistry()
        loader = ExcelLoader(conn, registry)


def validate_root_path(user_path: str) -> Path:
    path = Path(user_path).resolve()

    if not path.exists():
        raise ValueError(f"Path {path} does not exist")

    if not path.is_dir():
        raise ValueError(f"Path {path} is not a directory")

    return path


def _create_system_views(alias: str):
    import pandas as pd

    files_data = {}
    tables_data = []

    for table_name, meta in catalog.items():
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

    files_view_name = f"{alias}.__files"
    tables_view_name = f"{alias}.__tables"

    try:
        if files_data:
            files_df = pd.DataFrame(list(files_data.values()))
            conn.register(f"temp_files_{alias}", files_df)
            conn.execute(f'CREATE OR REPLACE VIEW "{files_view_name}" AS SELECT * FROM temp_files_{alias}')

        if tables_data:
            tables_df = pd.DataFrame(tables_data)
            conn.register(f"temp_tables_{alias}", tables_df)
            conn.execute(f'CREATE OR REPLACE VIEW "{tables_view_name}" AS SELECT * FROM temp_tables_{alias}')

        log.info("system_views_created", alias=alias, files_view=files_view_name, tables_view=tables_view_name)
    except Exception as e:
        log.warn("system_views_failed", alias=alias, error=str(e))


def load_dir(
    path: str,
    alias: str = None,
    include_glob: list[str] = None,
    exclude_glob: list[str] = None,
    overrides: dict = None,
) -> dict:
    init_server()

    include_glob = include_glob or ["**/*.xlsx"]
    exclude_glob = exclude_glob or []
    overrides = overrides or {}

    root = validate_root_path(path)

    if alias is None:
        import re
        alias = root.name or "excel"
        alias = alias.lower()
        alias = re.sub(r"[^a-z0-9_]+", "_", alias)
        alias = re.sub(r"_+", "_", alias)
        alias = alias.strip("_")
        if not alias:
            alias = "excel"

    log.info("load_start", path=str(root), alias=alias, patterns=include_glob)

    files_loaded = 0
    sheets_loaded = 0
    total_rows = 0
    failed_files = []

    config = LoadConfig(
        root=root,
        alias=alias,
        include_glob=include_glob,
        exclude_glob=exclude_glob,
        overrides=overrides,
    )
    load_configs[alias] = config

    for pattern in include_glob:
        for file_path in root.glob(pattern):
            if not file_path.is_file():
                continue

            relpath = str(file_path.relative_to(root))

            should_exclude = False
            for exclude_pattern in exclude_glob:
                if file_path.match(exclude_pattern):
                    should_exclude = True
                    break
            if should_exclude:
                continue

            try:
                sheet_names = loader.get_sheet_names(file_path)

                file_overrides = overrides.get(relpath, {})
                sheet_overrides = file_overrides.get("sheet_overrides", {})

                for sheet_name in sheet_names:
                    override_dict = sheet_overrides.get(sheet_name)
                    override = None
                    if override_dict:
                        override = SheetOverride(
                            skip_rows=override_dict.get("skip_rows", 0),
                            header_rows=override_dict.get("header_rows", 1),
                            skip_footer=override_dict.get("skip_footer", 0),
                            range=override_dict.get("range", ""),
                            drop_regex=override_dict.get("drop_regex", ""),
                            column_renames=override_dict.get("column_renames", {}),
                            type_hints=override_dict.get("type_hints", {}),
                            unpivot=override_dict.get("unpivot", {}),
                        )

                    meta = loader.load_sheet(file_path, relpath, sheet_name, alias, override)
                    catalog[meta.table_name] = meta
                    sheets_loaded += 1
                    total_rows += meta.est_rows

                    log.info("table_created", table=meta.table_name, file=relpath,
                            sheet=sheet_name, rows=meta.est_rows, mode=meta.mode)

                files_loaded += 1
            except Exception as e:
                error_msg = str(e)
                log.warn("load_failed", file=relpath, error=error_msg)
                failed_files.append({"file": relpath, "error": error_msg})

    _create_system_views(alias)

    result = {
        "alias": alias,
        "root": str(root),
        "files_count": files_loaded,
        "sheets_count": sheets_loaded,
        "tables_count": len([t for t in catalog if t.startswith(f"{alias}.")]),
        "rows_estimate": total_rows,
        "cache_mode": "none",
        "materialized": False,
    }

    if failed_files:
        result["failed"] = failed_files

    log.info("load_complete", alias=alias, files=files_loaded, sheets=sheets_loaded,
            rows=total_rows, failed=len(failed_files))

    return result


def query(
    sql: str,
    max_rows: int = 10000,
    timeout_ms: int = 60000,
) -> dict:
    init_server()

    start = time.time()
    interrupted = [False]

    def timeout_handler():
        interrupted[0] = True
        try:
            conn.interrupt()
        except Exception as e:
            log.warn("interrupt_failed", error=str(e))

    timer = threading.Timer(timeout_ms / 1000.0, timeout_handler)
    timer.start()

    try:
        conn.execute("BEGIN TRANSACTION READ ONLY")
        try:
            cursor = conn.execute(sql)
            result = cursor.fetchmany(max_rows + 1)
            columns = [{"name": desc[0], "type": str(desc[1])} for desc in cursor.description]
            conn.execute("COMMIT")
        except Exception as e:
            conn.execute("ROLLBACK")
            raise e
    except Exception as e:
        timer.cancel()
        if interrupted[0]:
            execution_ms = int((time.time() - start) * 1000)
            log.warn("query_timeout", execution_ms=execution_ms, timeout_ms=timeout_ms)
            raise TimeoutError(f"Query exceeded {timeout_ms}ms timeout")
        log.error("query_failed", error=str(e), sql=sql[:100])
        raise RuntimeError(f"Query failed: {e}")
    finally:
        timer.cancel()

    execution_ms = int((time.time() - start) * 1000)
    truncated = len(result) > max_rows
    rows = result[:max_rows]

    log.info("query_executed", rows=len(rows), execution_ms=execution_ms, truncated=truncated)

    return {
        "columns": columns,
        "rows": rows,
        "row_count": len(rows),
        "truncated": truncated,
        "execution_ms": execution_ms,
    }


def list_tables(alias: str = None) -> dict:
    init_server()

    tables = []
    for table_name, meta in catalog.items():
        if alias and not table_name.startswith(f"{alias}."):
            continue
        tables.append({
            "table": table_name,
            "file": meta.file,
            "relpath": meta.relpath,
            "sheet": meta.sheet,
            "mode": meta.mode,
            "est_rows": meta.est_rows,
        })

    return {"tables": tables}


def get_schema(table: str) -> dict:
    init_server()

    if table not in catalog:
        raise ValueError(f"Table {table} not found")

    result = conn.execute(f'DESCRIBE "{table}"').fetchall()
    columns = [
        {"name": row[0], "type": row[1], "nullable": row[2] == "YES"}
        for row in result
    ]

    return {"columns": columns}


def refresh(alias: str = None, full: bool = False) -> dict:
    init_server()

    if full:
        changed = 0
        dropped = 0
        added = 0

        tables_to_drop = []
        for table_name in catalog:
            if alias is None or table_name.startswith(f"{alias}."):
                tables_to_drop.append(table_name)

        for table_name in tables_to_drop:
            try:
                conn.execute(f'DROP VIEW IF EXISTS "{table_name}"')
                del catalog[table_name]
                dropped += 1
            except Exception:
                pass

        if alias and alias in load_configs:
            config = load_configs[alias]
            result = load_dir(
                path=str(config.root),
                alias=alias,
                include_glob=config.include_glob,
                exclude_glob=config.exclude_glob,
                overrides=config.overrides,
            )
            added = result["tables_count"]

        _create_system_views(alias)

        return {"files_count": result.get("files_count", 0), "sheets_count": result.get("sheets_count", 0), "changed": changed, "dropped": dropped, "added": added}
    else:
        changed = 0
        for table_name, meta in list(catalog.items()):
            try:
                file_path = Path(meta.file)

                if not file_path.exists():
                    log.warn("refresh_file_missing", table=table_name, file=meta.file)
                    continue

                current_mtime = file_path.stat().st_mtime
                if current_mtime > meta.mtime:
                    config = load_configs.get(meta.table_name.split("__")[0])
                    if not config:
                        log.warn("refresh_no_config", table=table_name)
                        continue

                    try:
                        relpath = str(file_path.relative_to(config.root))
                    except ValueError:
                        log.warn("refresh_path_outside_root", table=table_name,
                                file=str(file_path), root=str(config.root))
                        continue

                    override_dict = config.overrides.get(relpath, {}).get("sheet_overrides", {}).get(meta.sheet)
                    override = None
                    if override_dict:
                        override = SheetOverride(**override_dict)

                    conn.execute(f'DROP VIEW IF EXISTS "{table_name}"')
                    new_meta = loader.load_sheet(file_path, relpath, meta.sheet, config.alias, override)
                    catalog[table_name] = new_meta
                    changed += 1
            except Exception as e:
                log.warn("refresh_failed", table=table_name, error=str(e))
                continue

        return {"changed": changed, "total": len(catalog)}


def _on_file_change():
    log.info("file_change_detected", message="Auto-refreshing tables")

    try:
        for alias in load_configs.keys():
            result = refresh(alias=alias, full=False)
            log.info("auto_refresh_complete", alias=alias, changed=result.get("changed", 0))
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
    init_server()

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


if __name__ == "__main__":
    main()
