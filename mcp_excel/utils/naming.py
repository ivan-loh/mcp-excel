"""
Table name generation and collision handling.

TableRegistry: Production implementation (used by server.py and loader.py)
ImprovedTableRegistry: Experimental implementation with enhanced hierarchy preservation
"""

import re
import threading
from pathlib import Path
from collections import defaultdict


class TableRegistry:
    def __init__(self):
        self._names: dict[str, int] = {}
        self._collision_counts: defaultdict[str, int] = defaultdict(int)
        self._lock = threading.RLock()

    def register(self, alias: str, relpath: str, sheet: str, region_id: int = 0) -> str:
        with self._lock:
            sanitized = self._build_and_sanitize(alias, relpath, sheet, region_id)
            final_name = self._handle_collision(sanitized)
            self._names[final_name] = 1
            return final_name

    def _build_and_sanitize(self, alias: str, relpath: str, sheet: str, region_id: int) -> str:
        relpath_no_ext = relpath.rsplit(".", 1)[0] if "." in relpath else relpath

        relpath_components = relpath_no_ext.replace("\\", "/").split("/")

        parts = [alias] + relpath_components + [sheet]
        if region_id > 0:
            parts.append(f"r{region_id}")

        sanitized_parts = [self._sanitize_component(p) for p in parts if self._sanitize_component(p)]

        if not sanitized_parts:
            return "table"

        if len(sanitized_parts) == 1:
            sanitized_parts.append("table")

        name = ".".join(sanitized_parts)

        if name and name[0].isdigit():
            name = f"t_{name}"
        elif any(part and part[0].isdigit() for part in sanitized_parts):
            name = f"t_{name}"

        if len(name) > 63:
            name = name[:63]

        return name

    def _sanitize_component(self, component: str) -> str:
        component = component.lower()
        component = component.replace(' ', '_')
        component = re.sub(r'[^a-z0-9_$]', '', component)
        component = re.sub(r'_+', '_', component)
        component = component.strip('_')
        return component

    def _handle_collision(self, name: str) -> str:
        if name not in self._names:
            return name

        self._collision_counts[name] += 1
        collision_num = self._collision_counts[name] + 1

        while f"{name}_{collision_num}" in self._names:
            collision_num += 1

        return f"{name}_{collision_num}"

    def clear(self):
        with self._lock:
            self._names.clear()
            self._collision_counts.clear()


class ImprovedTableRegistry:
    """
    Experimental table naming that preserves folder hierarchy using dots.

    Examples:
        cnc/job_orders.xlsx → excel.cnc.job_orders.orders
        reports/2024/Q1/sales.xlsx → excel.reports.2024.q1.sales.summary

    Not currently used in production. For testing and comparison purposes.
    """

    def __init__(self):
        self._names: dict[str, int] = {}
        self._collision_counts: defaultdict[str, int] = defaultdict(int)
        self._lock = threading.RLock()

    def register(self, alias: str, relpath: str, sheet: str, region_id: int = 0) -> str:
        with self._lock:
            sanitized = self._build_hierarchical_name(alias, relpath, sheet, region_id)
            final_name = self._handle_collision(sanitized)
            self._names[final_name] = 1
            return final_name

    def _build_hierarchical_name(self, alias: str, relpath: str, sheet: str, region_id: int) -> str:
        path = Path(relpath)
        file_stem = path.stem
        folder_parts = list(path.parent.parts) if path.parent != Path('.') else []

        parts = [alias]

        if folder_parts:
            parts.extend(folder_parts)

        parts.append(file_stem)
        parts.append(sheet)

        if region_id > 0:
            parts.append(f"r{region_id}")

        sanitized_parts = [self._sanitize_component(p) for p in parts if p]
        sanitized_parts = [p for p in sanitized_parts if p]

        if not sanitized_parts:
            return "table"

        name = ".".join(sanitized_parts)

        if name and name[0].isdigit():
            name = f"t_{name}"

        if len(name) > 63:
            name = self._smart_truncate(name, 63)

        return name

    def _sanitize_component(self, component: str) -> str:
        component = component.lower()
        component = component.replace(' ', '_')
        component = component.replace('-', '_')
        component = re.sub(r'[^a-z0-9_]', '', component)
        component = re.sub(r'_+', '_', component)
        component = component.strip('_')
        return component

    def _smart_truncate(self, name: str, max_length: int) -> str:
        if len(name) <= max_length:
            return name

        parts = name.split('.')

        if len(parts) <= 2:
            return name[:max_length]

        first = parts[0]
        last = parts[-1]
        available = max_length - len(first) - len(last) - 2

        if available <= 0:
            return name[:max_length]

        middle_parts = parts[1:-1]
        middle_str = '.'.join(middle_parts)

        if len(middle_str) <= available:
            return name

        if len(middle_parts) == 1:
            return f"{first}.{middle_parts[0][:available]}.{last}"

        per_part = max(2, available // (len(middle_parts) * 2))
        shortened = []
        remaining = available

        for i, part in enumerate(middle_parts):
            if i == len(middle_parts) - 1:
                part_len = min(len(part), remaining)
            else:
                part_len = min(len(part), per_part)
                remaining -= part_len + 1

            if part_len > 0:
                shortened.append(part[:part_len])

        result = '.'.join([first] + shortened + [last])

        if len(result) > max_length:
            return result[:max_length]

        return result

    def _handle_collision(self, name: str) -> str:
        if name not in self._names:
            return name

        self._collision_counts[name] += 1
        collision_num = self._collision_counts[name] + 1

        while f"{name}_{collision_num}" in self._names:
            collision_num += 1

        return f"{name}_{collision_num}"

    def clear(self):
        with self._lock:
            self._names.clear()
            self._collision_counts.clear()
