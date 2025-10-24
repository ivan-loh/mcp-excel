import re
import threading
from pathlib import Path
from collections import defaultdict


class ImprovedTableRegistry:
    """
    Improved table naming that preserves folder hierarchy using dots.

    Examples:
        cnc/job_orders.xlsx → excel.cnc.job_orders.orders
        reports/2024/Q1/sales.xlsx → excel.reports.2024.q1.sales.summary
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
        # Convert path to Path object for easier handling
        path = Path(relpath)

        # Remove file extension
        file_stem = path.stem

        # Get folder parts (if any)
        folder_parts = list(path.parent.parts) if path.parent != Path('.') else []

        # Build the hierarchical name
        parts = [alias]

        # Add folder hierarchy
        if folder_parts:
            parts.extend(folder_parts)

        # Add file name
        parts.append(file_stem)

        # Add sheet name
        parts.append(sheet)

        # Add region if specified
        if region_id > 0:
            parts.append(f"r{region_id}")

        # Sanitize each part individually
        sanitized_parts = [self._sanitize_component(p) for p in parts if p]

        # Filter out empty parts
        sanitized_parts = [p for p in sanitized_parts if p]

        if not sanitized_parts:
            return "table"

        # Join with dots
        name = ".".join(sanitized_parts)

        # Handle leading digit
        if name and name[0].isdigit():
            name = f"t_{name}"

        # Truncate if too long (PostgreSQL identifier limit)
        if len(name) > 63:
            name = self._smart_truncate(name, 63)

        return name

    def _sanitize_component(self, component: str) -> str:
        """Sanitize a single path component."""
        # Convert to lowercase
        component = component.lower()

        # Replace spaces with underscores
        component = component.replace(' ', '_')

        # Replace hyphens with underscores (more readable than removing)
        component = component.replace('-', '_')

        # Remove non-alphanumeric characters except underscore
        component = re.sub(r'[^a-z0-9_]', '', component)

        # Collapse multiple underscores
        component = re.sub(r'_+', '_', component)

        # Remove leading/trailing underscores
        component = component.strip('_')

        return component

    def _smart_truncate(self, name: str, max_length: int) -> str:
        """
        Truncate name while trying to preserve structure.
        Keep alias and sheet, truncate middle parts.
        """
        if len(name) <= max_length:
            return name

        parts = name.split('.')

        if len(parts) <= 2:
            # Simple truncation
            return name[:max_length]

        # Keep first (alias) and last (sheet) parts
        first = parts[0]
        last = parts[-1]

        # Calculate available space for middle parts
        available = max_length - len(first) - len(last) - 2  # -2 for dots

        if available <= 0:
            # Can't preserve structure, just truncate
            return name[:max_length]

        # Truncate middle parts
        middle_parts = parts[1:-1]
        middle_str = '.'.join(middle_parts)

        if len(middle_str) <= available:
            return name  # Shouldn't happen, but just in case

        # Shorten middle parts to fit
        if len(middle_parts) == 1:
            # Single middle part - truncate it
            return f"{first}.{middle_parts[0][:available]}.{last}"

        # Multiple middle parts - abbreviate each
        per_part = max(2, available // (len(middle_parts) * 2))  # Account for dots

        shortened = []
        remaining = available

        for i, part in enumerate(middle_parts):
            if i == len(middle_parts) - 1:
                # Last middle part gets remaining space
                part_len = min(len(part), remaining)
            else:
                part_len = min(len(part), per_part)
                remaining -= part_len + 1  # +1 for dot

            if part_len > 0:
                shortened.append(part[:part_len])

        result = '.'.join([first] + shortened + [last])

        # Final safety check
        if len(result) > max_length:
            return result[:max_length]

        return result

    def _handle_collision(self, name: str) -> str:
        """Handle naming collisions by appending numbers."""
        if name not in self._names:
            return name

        self._collision_counts[name] += 1
        collision_num = self._collision_counts[name] + 1

        while f"{name}_{collision_num}" in self._names:
            collision_num += 1

        return f"{name}_{collision_num}"

    def clear(self):
        """Clear the registry."""
        with self._lock:
            self._names.clear()
            self._collision_counts.clear()


# Comparison function for testing
def compare_naming_approaches():
    """Compare current vs improved naming."""
    from mcp_excel.naming import TableRegistry

    current = TableRegistry()
    improved = ImprovedTableRegistry()

    test_cases = [
        ('excel', 'data.xlsx', 'Sheet1'),
        ('excel', 'folder/data.xlsx', 'Sheet1'),
        ('excel', 'cnc/job_orders.xlsx', 'Orders'),
        ('excel', 'cnc/cost_analysis.xlsx', 'Job_Costing'),
        ('excel', 'reports/2024/Q1/sales.xlsx', 'Summary'),
        ('excel', 'reports/2024/Q2/sales.xlsx', 'Summary'),
        ('excel', 'dept-01/region/data.xlsx', 'Sheet1'),
        ('excel', 'a/b/c/d/e/f/g/very_long_filename.xlsx', 'VeryLongSheetName'),
    ]

    print("Comparison of Naming Approaches\n")
    print("=" * 100)
    print(f"{'File Path':<45} {'Current':<35} {'Improved':<35}")
    print("-" * 100)

    for alias, relpath, sheet in test_cases:
        current_name = current.register(alias, relpath, sheet)
        improved_name = improved.register(alias, relpath, sheet)

        print(f"{relpath:<45} {current_name:<35} {improved_name:<35}")

    # Test collision handling
    print("\n\nCollision Test:")
    print("-" * 100)

    # Reset registries
    current.clear()
    improved.clear()

    collision_tests = [
        ('excel', 'cnc/reports.xlsx', 'Sheet1'),
        ('excel', 'cncreports.xlsx', 'Sheet1'),
        ('excel', 'cnc_reports.xlsx', 'Sheet1'),
    ]

    print(f"{'File Path':<30} {'Current':<35} {'Improved':<35}")
    print("-" * 100)

    for alias, relpath, sheet in collision_tests:
        current_name = current.register(alias, relpath, sheet)
        improved_name = improved.register(alias, relpath, sheet)

        print(f"{relpath:<30} {current_name:<35} {improved_name:<35}")


if __name__ == "__main__":
    compare_naming_approaches()