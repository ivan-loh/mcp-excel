from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional


@dataclass
class TableMeta:
    table_name: str
    file: str
    relpath: str
    sheet: str
    mode: str
    mtime: float
    alias: str
    est_rows: int = 0


@dataclass
class MergeHandlingConfig:
    strategy: str = 'fill'
    header_strategy: str = 'span'
    log_warnings: bool = True


@dataclass
class LocaleConfig:
    locale: Optional[str] = None
    decimal_separator: Optional[str] = None
    thousands_separator: Optional[str] = None
    currency_symbols: list[str] = field(default_factory=list)
    auto_detect: bool = True


@dataclass
class SheetOverride:
    skip_rows: int = 0
    header_rows: int = 1
    skip_footer: int = 0
    range: str = ""
    drop_regex: str = ""
    column_renames: dict[str, str] = field(default_factory=dict)
    type_hints: dict[str, str] = field(default_factory=dict)
    unpivot: dict[str, Any] = field(default_factory=dict)
    auto_detect: bool = False
    merge_handling: Optional[MergeHandlingConfig] = None
    include_hidden: bool = False
    locale: Optional[LocaleConfig] = None
    drop_conditions: list[dict] = field(default_factory=list)
    extract_table: Optional[int] = None
    table_range: Optional[str] = None


@dataclass
class StructureInfo:
    data_start_row: int
    data_end_row: int
    data_start_col: int
    data_end_col: int
    header_row: Optional[int]
    header_rows_count: int
    header_confidence: float
    metadata_rows: list[int]
    metadata_type: str
    merged_ranges: list[tuple[int, int, int, int]]
    merged_in_headers: bool
    merged_in_data: bool
    hidden_rows: list[int]
    hidden_columns: list[int]
    detected_locale: str
    decimal_separator: str
    thousands_separator: str
    num_tables: int
    table_ranges: list[dict]
    blank_rows: list[int]
    inconsistent_columns: bool
    has_formulas: bool
    suggested_skip_rows: int
    suggested_skip_footer: int
    suggested_overrides: dict


@dataclass
class LoadConfig:
    root: Path
    alias: str
    include_glob: list[str]
    exclude_glob: list[str]
    overrides: dict[str, dict] = field(default_factory=dict)
