from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class TableMeta:
    table_name: str
    file: str
    relpath: str
    sheet: str
    mode: str
    mtime: float
    est_rows: int = 0


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


@dataclass
class LoadConfig:
    root: Path
    alias: str
    include_glob: list[str]
    exclude_glob: list[str]
    overrides: dict[str, dict] = field(default_factory=dict)
