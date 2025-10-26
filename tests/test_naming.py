import pytest
import threading
from pathlib import Path
from mcp_excel.naming import TableRegistry
from mcp_excel.naming_improved import ImprovedTableRegistry

pytestmark = pytest.mark.unit


class TestTableRegistry:
    def setup_method(self):
        self.registry = TableRegistry()

    def test_basic_sanitization(self):
        name = self.registry.register("excel", "sales/report.xlsx", "Summary")
        assert name == "excel.sales.report.summary"

    def test_special_chars(self):
        name = self.registry.register("excel", "data/Q1-2024 (Final).xlsx", "Sheet1")
        assert name == "excel.data.q12024_final.sheet1"

    def test_unicode(self):
        name = self.registry.register("excel", "données/rapport.xlsx", "Feuille")
        assert name == "excel.donnes.rapport.feuille"

    def test_leading_numbers(self):
        name = self.registry.register("excel", "2024/report.xlsx", "1stQuarter")
        assert name.startswith("t_")

    def test_collision_handling(self):
        name1 = self.registry.register("excel", "sales/report.xlsx", "Summary")
        name2 = self.registry.register("excel", "sales/report.xlsx", "Summary")
        assert name1 != name2
        assert name2.endswith("_2")

    def test_multiple_collisions(self):
        name1 = self.registry.register("excel", "data.xlsx", "Sheet")
        name2 = self.registry.register("excel", "data.xlsx", "Sheet")
        name3 = self.registry.register("excel", "data.xlsx", "Sheet")
        assert name1 == "excel.data.sheet"
        assert name2 == "excel.data.sheet_2"
        assert name3 == "excel.data.sheet_3"

    def test_long_names(self):
        long_relpath = "a" * 100
        name = self.registry.register("excel", f"{long_relpath}.xlsx", "Sheet")
        assert len(name) <= 63

    def test_empty_components(self):
        name = self.registry.register("excel", ".xlsx", "")
        assert name
        assert "." in name

    def test_clear(self):
        name1 = self.registry.register("excel", "test.xlsx", "Sheet")
        self.registry.clear()
        name2 = self.registry.register("excel", "test.xlsx", "Sheet")
        assert name1 == name2

    def test_hierarchical_structure(self):
        name = self.registry.register("excel", "folder/data.xlsx", "Sheet1")
        assert name == "excel.folder.data.sheet1"

    def test_subfolder_hierarchy(self):
        name = self.registry.register("excel", "cnc/job_orders.xlsx", "Orders")
        assert name == "excel.cnc.job_orders.orders"

    def test_deep_path_hierarchy(self):
        name = self.registry.register("excel", "reports/2024/Q1/sales.xlsx", "Summary")
        assert name == "t_excel.reports.2024.q1.sales.summary"

    def test_no_collision_with_hierarchy(self):
        name1 = self.registry.register("excel", "cnc/reports.xlsx", "Sheet1")
        name2 = self.registry.register("excel", "cncreports.xlsx", "Sheet1")
        assert name1 == "excel.cnc.reports.sheet1"
        assert name2 == "excel.cncreports.sheet1"

    def test_space_handling(self):
        name = self.registry.register("excel", "my folder/data.xlsx", "Sheet1")
        assert name == "excel.my_folder.data.sheet1"

    def test_special_chars_removed(self):
        name = self.registry.register("excel", "dept-01/données.xlsx", "Feuille1")
        assert name == "excel.dept01.donnes.feuille1"

    def test_windows_path_separators(self):
        windows_path = "folder\\subfolder\\data.xlsx"
        name = self.registry.register("excel", windows_path, "Sheet1")
        assert ".folder.subfolder." in name


class TestImprovedTableRegistry:
    def setup_method(self):
        self.registry = ImprovedTableRegistry()

    def test_basic_sanitization(self):
        name = self.registry.register("excel", "sales/report.xlsx", "Summary")
        assert name == "excel.sales.report.summary"

    def test_hierarchical_structure(self):
        name = self.registry.register("excel", "folder/data.xlsx", "Sheet1")
        assert name == "excel.folder.data.sheet1"

    def test_subfolder_hierarchy(self):
        name = self.registry.register("excel", "cnc/job_orders.xlsx", "Orders")
        assert name == "excel.cnc.job_orders.orders"

    def test_deep_path_hierarchy(self):
        name = self.registry.register("excel", "reports/2024/Q1/sales.xlsx", "Summary")
        assert name == "excel.reports.2024.q1.sales.summary"

    def test_no_collision_with_hierarchy(self):
        name1 = self.registry.register("excel", "cnc/reports.xlsx", "Sheet1")
        name2 = self.registry.register("excel", "cncreports.xlsx", "Sheet1")
        assert name1 == "excel.cnc.reports.sheet1"
        assert name2 == "excel.cncreports.sheet1"

    def test_space_handling(self):
        name = self.registry.register("excel", "my folder/data.xlsx", "Sheet1")
        assert name == "excel.my_folder.data.sheet1"

    def test_special_chars_handling(self):
        name = self.registry.register("excel", "dept-01/région/données.xlsx", "Feuille1")
        assert name == "excel.dept_01.rgion.donnes.feuille1"

    def test_hyphen_to_underscore(self):
        name = self.registry.register("excel", "my-folder/my-file.xlsx", "Sheet1")
        assert name == "excel.my_folder.my_file.sheet1"

    def test_truncation_preserves_structure(self):
        long_path = "a/b/c/d/e/f/g/very_long_filename_that_exceeds_limits.xlsx"
        name = self.registry.register("excel", long_path, "VeryLongSheetName")
        assert len(name) <= 63
        assert name.startswith("excel.a.b.c.")


class TestNamingComparison:
    def setup_method(self):
        self.current = TableRegistry()
        self.improved = ImprovedTableRegistry()

    def test_query_pattern_precision(self):
        files = [
            "cnc/job_orders.xlsx",
            "cnc/cost_analysis.xlsx",
            "cnc/machine_status.xlsx",
            "finance/reports.xlsx",
            "cncreports.xlsx",
        ]

        current_names = []
        improved_names = []

        for file in files:
            current_names.append(self.current.register("excel", file, "Sheet1"))
            improved_names.append(self.improved.register("excel", file, "Sheet1"))

        cnc_current = [n for n in current_names if ".cnc." in n]
        cnc_improved = [n for n in improved_names if ".cnc." in n]

        assert len(cnc_current) == 3
        assert len(cnc_improved) == 3

    def test_collision_handling_difference(self):
        collision_tests = [
            ("cnc/reports.xlsx", "excel.cnc.reports.sheet1", "excel.cnc.reports.sheet1"),
            ("cncreports.xlsx", "excel.cncreports.sheet1", "excel.cncreports.sheet1"),
            ("cnc_reports.xlsx", "excel.cnc_reports.sheet1", "excel.cnc_reports.sheet1"),
        ]

        for file, expected_current, expected_improved in collision_tests:
            current_name = self.current.register("excel", file, "Sheet1")
            improved_name = self.improved.register("excel", file, "Sheet1")
            assert current_name == expected_current
            assert improved_name == expected_improved

    def test_readability_comparison(self):
        test_file = "departments/finance/reports/2024/quarterly/summary.xlsx"
        current_name = self.current.register("excel", test_file, "Q1")
        improved_name = self.improved.register("excel", test_file, "Q1")

        for part in [".departments.", ".finance.", ".reports.", ".2024."]:
            assert part in current_name
            assert part in improved_name


class TestEdgeCases:
    def test_empty_components(self):
        current = TableRegistry()
        improved = ImprovedTableRegistry()
        assert current.register("excel", "//data.xlsx", "Sheet1")
        assert improved.register("excel", "//data.xlsx", "Sheet1")

    def test_unicode_handling(self):
        current = TableRegistry()
        improved = ImprovedTableRegistry()
        unicode_file = "日本/データ.xlsx"
        current_name = current.register("excel", unicode_file, "シート1")
        improved_name = improved.register("excel", unicode_file, "シート1")
        assert all(ord(c) < 128 for c in current_name)
        assert all(ord(c) < 128 for c in improved_name)

    def test_very_long_paths(self):
        current = TableRegistry()
        improved = ImprovedTableRegistry()
        long_path = "/".join(["folder"] * 20) + "/file.xlsx"
        current_name = current.register("excel", long_path, "Sheet1")
        improved_name = improved.register("excel", long_path, "Sheet1")
        assert len(current_name) <= 63
        assert len(improved_name) <= 63

    def test_numeric_prefixes(self):
        current = TableRegistry()
        improved = ImprovedTableRegistry()
        name1_current = current.register("excel", "2024/reports.xlsx", "Sheet1")
        name1_improved = improved.register("excel", "2024/reports.xlsx", "Sheet1")
        assert not name1_current[0].isdigit()
        assert not name1_improved[0].isdigit()


class TestSQLCompatibility:
    def test_quoting_requirement(self):
        current = TableRegistry()
        improved = ImprovedTableRegistry()
        test_cases = [
            "data.xlsx",
            "folder/data.xlsx",
            "my folder/my-file.xlsx",
            "2024/reports.xlsx",
            "dept-01/région/données.xlsx",
        ]

        valid_chars = set("abcdefghijklmnopqrstuvwxyz0123456789_.")
        for file in test_cases:
            current_name = current.register("excel", file, "Sheet1")
            improved_name = improved.register("excel", file, "Sheet1")
            assert all(c in valid_chars for c in current_name)
            assert all(c in valid_chars for c in improved_name)

    def test_reserved_words_handling(self):
        current = TableRegistry()
        improved = ImprovedTableRegistry()
        reserved_words = ["select", "from", "where", "table", "order"]

        for word in reserved_words:
            file = f"{word}/{word}.xlsx"
            current_name = current.register("excel", file, word.title())
            improved_name = improved.register("excel", file, word.title())
            assert len(current_name) > 0
            assert len(improved_name) > 0
            assert word in current_name.lower()
            assert word in improved_name.lower()


@pytest.mark.parametrize("registry_class", [TableRegistry, ImprovedTableRegistry])
class TestRegistryMethods:
    def test_clear_method(self, registry_class):
        registry = registry_class()
        registry.register("excel", "file1.xlsx", "Sheet1")
        registry.register("excel", "file2.xlsx", "Sheet1")
        registry.clear()
        name1 = registry.register("excel", "file1.xlsx", "Sheet1")
        name2 = registry.register("excel", "file1.xlsx", "Sheet1")
        assert name1 != name2
        assert name2.endswith("_2")

    def test_thread_safety(self, registry_class):
        registry = registry_class()
        names = []

        def register_names():
            for i in range(10):
                name = registry.register("excel", f"file{i}.xlsx", "Sheet1")
                names.append(name)

        threads = [threading.Thread(target=register_names) for _ in range(5)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        assert len(names) == 50
        assert len(set(names)) == len(names)
