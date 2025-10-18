import re
from pathlib import Path
from typing import Optional
import pandas as pd
import duckdb

from .types import SheetOverride, TableMeta
from .naming import TableRegistry


class ExcelLoader:
    def __init__(self, conn: duckdb.DuckDBPyConnection, registry: TableRegistry):
        self.conn = conn
        self.registry = registry
        self._ensure_excel_extension()

    def _ensure_excel_extension(self):
        try:
            self.conn.execute("INSTALL excel")
            self.conn.execute("LOAD excel")
        except Exception:
            pass

    def load_sheet(
        self,
        file: Path,
        relpath: str,
        sheet: str,
        alias: str,
        override: Optional[SheetOverride] = None,
    ) -> TableMeta:
        table_name = self.registry.register(alias, relpath, sheet)

        if override:
            return self._load_assisted(file, relpath, sheet, table_name, override)
        else:
            return self._load_raw(file, relpath, sheet, table_name)

    def _load_raw(self, file: Path, relpath: str, sheet: str, table_name: str) -> TableMeta:
        try:
            self.conn.execute(f"""
                CREATE OR REPLACE VIEW "{table_name}" AS
                SELECT * FROM read_xlsx(
                    '{file}',
                    sheet='{sheet}',
                    header=false,
                    all_varchar=true
                )
            """)

            count_result = self.conn.execute(f'SELECT COUNT(*) FROM "{table_name}"').fetchone()
            est_rows = count_result[0] if count_result else 0

            return TableMeta(
                table_name=table_name,
                file=str(file),
                relpath=relpath,
                sheet=sheet,
                mode="RAW",
                mtime=file.stat().st_mtime,
                est_rows=est_rows,
            )
        except Exception as e:
            error_msg = str(e)
            suggestion = self._get_error_suggestion(error_msg, "RAW")
            raise RuntimeError(f"Failed to load {file}:{sheet} in RAW mode: {error_msg}{suggestion}")

    def _load_assisted(
        self, file: Path, relpath: str, sheet: str, table_name: str, override: SheetOverride
    ) -> TableMeta:
        try:
            if override.header_rows > 1:
                df = self._load_multirow_header(file, sheet, override)
            else:
                has_header = override.header_rows > 0
                range_clause = f", range='{override.range}'" if override.range else ""

                df = self.conn.execute(f"""
                    SELECT * FROM read_xlsx(
                        '{file}',
                        sheet='{sheet}',
                        header={has_header},
                        all_varchar=true
                        {range_clause}
                    )
                """).df()

            if override.skip_rows > 0 and not override.range and override.header_rows <= 1:
                df = df.iloc[override.skip_rows:]

            if override.skip_footer > 0:
                df = df.iloc[:-override.skip_footer]

            if override.drop_regex:
                if len(df.columns) > 0:
                    first_col = df.columns[0]
                    pattern = override.drop_regex
                    df = df[~df[first_col].astype(str).str.match(pattern, na=False)]

            if override.column_renames:
                df = df.rename(columns=override.column_renames)

            if override.type_hints:
                df = self._apply_type_hints(df, override.type_hints)

            if override.unpivot:
                df = self._apply_unpivot(df, override.unpivot)

            import hashlib
            temp_view = f"temp_{hashlib.md5(table_name.encode()).hexdigest()[:8]}"
            self.conn.register(temp_view, df)
            self.conn.execute(f"""
                CREATE OR REPLACE VIEW "{table_name}" AS
                SELECT * FROM {temp_view}
            """)

            return TableMeta(
                table_name=table_name,
                file=str(file),
                relpath=relpath,
                sheet=sheet,
                mode="ASSISTED",
                mtime=file.stat().st_mtime,
                est_rows=len(df),
            )
        except Exception as e:
            error_msg = str(e)
            suggestion = self._get_error_suggestion(error_msg, "ASSISTED")
            raise RuntimeError(f"Failed to load {file}:{sheet} in ASSISTED mode: {error_msg}{suggestion}")

    def _load_multirow_header(self, file: Path, sheet: str, override: SheetOverride) -> pd.DataFrame:
        df_raw = self.conn.execute(f"""
            SELECT * FROM read_xlsx(
                '{file}',
                sheet='{sheet}',
                header=false,
                all_varchar=true
                {f", range='{override.range}'" if override.range else ""}
            )
        """).df()

        if override.skip_rows > 0:
            df_raw = df_raw.iloc[override.skip_rows:]

        header_rows = df_raw.iloc[:override.header_rows]
        data_rows = df_raw.iloc[override.header_rows:]

        new_columns = []
        for col_idx in range(len(header_rows.columns)):
            col_parts = []
            for row_idx in range(len(header_rows)):
                val = str(header_rows.iloc[row_idx, col_idx])
                if val and val != "nan":
                    col_parts.append(val)

            if col_parts:
                new_col_name = "__".join(col_parts)
            else:
                new_col_name = f"col_{col_idx}"

            new_columns.append(new_col_name)

        data_rows.columns = new_columns
        data_rows = data_rows.reset_index(drop=True)

        return data_rows

    def _apply_type_hints(self, df: pd.DataFrame, type_hints: dict[str, str]) -> pd.DataFrame:
        for col_name, type_hint in type_hints.items():
            if col_name not in df.columns:
                continue

            type_upper = type_hint.upper()
            integer_types = ("INT", "BIGINT", "SMALLINT")
            numeric_types = ("DECIMAL", "NUMERIC", "DOUBLE", "FLOAT")

            if any(t in type_upper for t in integer_types):
                df[col_name] = pd.to_numeric(df[col_name], errors="coerce").astype("Int64")
            elif any(t in type_upper for t in numeric_types):
                df[col_name] = pd.to_numeric(df[col_name], errors="coerce")
            elif "DATE" in type_upper:
                df[col_name] = pd.to_datetime(df[col_name], errors="coerce")
            elif "BOOL" in type_upper:
                df[col_name] = df[col_name].astype(str).str.lower().isin(["true", "1", "yes", "y"])

        return df

    def _apply_unpivot(self, df: pd.DataFrame, unpivot_config: dict) -> pd.DataFrame:
        id_vars = unpivot_config.get("id_vars", [])
        value_vars = unpivot_config.get("value_vars", [])
        var_name = unpivot_config.get("var_name", "variable")
        value_name = unpivot_config.get("value_name", "value")

        if not value_vars:
            value_vars = [col for col in df.columns if col not in id_vars]

        return df.melt(id_vars=id_vars, value_vars=value_vars, var_name=var_name, value_name=value_name)

    def _get_error_suggestion(self, error_msg: str, mode: str) -> str:
        error_lower = error_msg.lower()
        suggestions = []

        if "column" in error_lower and "mismatch" in error_lower:
            suggestions.append("Try adding 'skip_rows' to skip header rows")
            suggestions.append("Or use 'drop_regex' to filter problematic rows")

        if "header" in error_lower or "column name" in error_lower:
            suggestions.append("Consider using 'header_rows: 0' or 'header_rows: 2' to adjust header detection")
            suggestions.append("Use 'column_renames' to fix column names")

        if "row" in error_lower and ("empty" in error_lower or "null" in error_lower):
            suggestions.append("Use 'skip_footer' to remove trailing empty rows")
            suggestions.append("Or 'drop_regex' to filter specific rows")

        if "type" in error_lower or "convert" in error_lower or "cast" in error_lower:
            suggestions.append("Use 'type_hints' to specify column types explicitly")
            suggestions.append("Consider loading in RAW mode first to inspect the data")

        if "range" in error_lower:
            suggestions.append("Check that 'range' parameter uses valid Excel notation (e.g., 'A1:F100')")

        if mode == "RAW" and not suggestions:
            suggestions.append("Try ASSISTED mode with overrides to handle messy data")
            suggestions.append("Use 'skip_rows' and 'skip_footer' to exclude problematic rows")

        if suggestions:
            return "\n\nSuggestions:\n- " + "\n- ".join(suggestions)
        return ""

    def get_sheet_names(self, file: Path) -> list[str]:
        try:
            result = self.conn.execute(f"""
                SELECT sheet_name FROM st_read('{file}')
            """).fetchall()
            return [row[0] for row in result]
        except Exception:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
                sheets = wb.sheetnames
                wb.close()
                return sheets
            except Exception as e:
                raise RuntimeError(f"Failed to read sheet names from {file}: {e}")
