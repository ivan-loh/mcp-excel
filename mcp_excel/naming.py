import re
import threading
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

        parts = [alias, relpath_no_ext, sheet]
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
