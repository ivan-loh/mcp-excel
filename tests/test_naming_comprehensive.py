"""
Comprehensive tests for table naming conventions.
Tests both current implementation and proposed improvements.
"""

import pytest
from pathlib import Path

from mcp_excel.naming import TableRegistry
from mcp_excel.naming_improved import ImprovedTableRegistry


class TestCurrentNaming:
    """Test current naming implementation to document its behavior."""

    def setup_method(self):
        self.registry = TableRegistry()

    def test_flat_structure(self):
        """Current implementation flattens folder structure."""
        name = self.registry.register('excel', 'folder/data.xlsx', 'Sheet1')
        # Current flattens: folder/data → folderdata
        assert name == 'excel.folderdata.sheet1'

    def test_subfolder_concatenation(self):
        """Subfolders are concatenated, not hierarchical."""
        name = self.registry.register('excel', 'cnc/job_orders.xlsx', 'Orders')
        assert name == 'excel.cncjob_orders.orders'
        # Note: 'cnc' and 'job_orders' are concatenated

    def test_deep_path_flattening(self):
        """Deep paths become long concatenated strings."""
        name = self.registry.register('excel', 'reports/2024/Q1/sales.xlsx', 'Summary')
        assert name == 'excel.reports2024q1sales.summary'
        # All path separators removed

    def test_collision_with_flattening(self):
        """Flattening can cause collisions."""
        # These different files could collide
        name1 = self.registry.register('excel', 'cnc/reports.xlsx', 'Sheet1')
        name2 = self.registry.register('excel', 'cncreports.xlsx', 'Sheet1')

        # Current implementation handles collision with suffix
        assert name1 == 'excel.cncreports.sheet1'
        assert name2 == 'excel.cncreports.sheet1_2'  # Collision!

    def test_space_handling(self):
        """Spaces become underscores."""
        name = self.registry.register('excel', 'my folder/data.xlsx', 'Sheet1')
        assert name == 'excel.my_folderdata.sheet1'

    def test_special_chars_removed(self):
        """Special characters are removed."""
        name = self.registry.register('excel', 'dept-01/données.xlsx', 'Feuille1')
        assert name == 'excel.dept01donnes.feuille1'


class TestImprovedNaming:
    """Test improved hierarchical naming implementation."""

    def setup_method(self):
        self.registry = ImprovedTableRegistry()

    def test_hierarchical_structure(self):
        """Improved implementation preserves hierarchy."""
        name = self.registry.register('excel', 'folder/data.xlsx', 'Sheet1')
        # Improved preserves: folder.data
        assert name == 'excel.folder.data.sheet1'

    def test_subfolder_hierarchy(self):
        """Subfolders maintain hierarchy with dots."""
        name = self.registry.register('excel', 'cnc/job_orders.xlsx', 'Orders')
        assert name == 'excel.cnc.job_orders.orders'
        # Note: 'cnc' and 'job_orders' are separated by dot

    def test_deep_path_hierarchy(self):
        """Deep paths maintain clear hierarchy."""
        name = self.registry.register('excel', 'reports/2024/Q1/sales.xlsx', 'Summary')
        assert name == 'excel.reports.2024.q1.sales.summary'
        # Clear hierarchy: reports → 2024 → Q1 → sales

    def test_no_collision_with_hierarchy(self):
        """Hierarchical naming reduces collisions."""
        name1 = self.registry.register('excel', 'cnc/reports.xlsx', 'Sheet1')
        name2 = self.registry.register('excel', 'cncreports.xlsx', 'Sheet1')

        # Different files get different names
        assert name1 == 'excel.cnc.reports.sheet1'
        assert name2 == 'excel.cncreports.sheet1'  # No collision!

    def test_space_handling(self):
        """Spaces become underscores, hierarchy preserved."""
        name = self.registry.register('excel', 'my folder/data.xlsx', 'Sheet1')
        assert name == 'excel.my_folder.data.sheet1'

    def test_special_chars_handling(self):
        """Special chars removed but hierarchy preserved."""
        name = self.registry.register('excel', 'dept-01/région/données.xlsx', 'Feuille1')
        assert name == 'excel.dept_01.rgion.donnes.feuille1'

    def test_hyphen_to_underscore(self):
        """Hyphens become underscores for readability."""
        name = self.registry.register('excel', 'my-folder/my-file.xlsx', 'Sheet1')
        assert name == 'excel.my_folder.my_file.sheet1'

    def test_truncation_preserves_structure(self):
        """Long names are truncated while preserving structure."""
        # Create a very long path
        long_path = 'a/b/c/d/e/f/g/very_long_filename_that_exceeds_limits.xlsx'
        name = self.registry.register('excel', long_path, 'VeryLongSheetName')

        # Should be truncated but maintain dots
        assert len(name) <= 63
        assert name.startswith('excel.a.b.c.')
        assert '.verylongsheetname' in name or name.endswith('.ver')  # Sheet name preserved or truncated


class TestNamingComparison:
    """Direct comparison tests between current and improved."""

    def setup_method(self):
        self.current = TableRegistry()
        self.improved = ImprovedTableRegistry()

    def test_query_pattern_difference(self):
        """Test how different naming affects query patterns."""
        # Register same files with both systems
        files = [
            'cnc/job_orders.xlsx',
            'cnc/cost_analysis.xlsx',
            'cnc/machine_status.xlsx',
            'finance/reports.xlsx',
            'cncreports.xlsx',  # Potential collision
        ]

        current_names = []
        improved_names = []

        for file in files:
            current_names.append(self.current.register('excel', file, 'Sheet1'))
            improved_names.append(self.improved.register('excel', file, 'Sheet1'))

        # Current: Finding CNC tables is ambiguous
        cnc_current = [n for n in current_names if 'cnc' in n]
        assert len(cnc_current) == 4  # Includes 'cncreports.xlsx'!

        # Improved: Finding CNC tables is precise
        cnc_improved = [n for n in improved_names if '.cnc.' in n]
        assert len(cnc_improved) == 3  # Only actual cnc/ folder files

    def test_collision_handling_difference(self):
        """Test how each system handles potential collisions."""
        collision_tests = [
            ('cnc/reports.xlsx', 'excel.cncreports.sheet1', 'excel.cnc.reports.sheet1'),
            ('cncreports.xlsx', 'excel.cncreports.sheet1_2', 'excel.cncreports.sheet1'),  # Collision in current
            ('cnc_reports.xlsx', 'excel.cnc_reports.sheet1', 'excel.cnc_reports.sheet1'),
        ]

        for file, expected_current, expected_improved in collision_tests:
            current_name = self.current.register('excel', file, 'Sheet1')
            improved_name = self.improved.register('excel', file, 'Sheet1')

            assert current_name == expected_current, f"Current naming failed for {file}"
            assert improved_name == expected_improved, f"Improved naming failed for {file}"

    def test_readability_comparison(self):
        """Test readability of generated names."""
        test_file = 'departments/finance/reports/2024/quarterly/summary.xlsx'

        current_name = self.current.register('excel', test_file, 'Q1')
        improved_name = self.improved.register('excel', test_file, 'Q1')

        # Current creates unreadable concatenation
        assert 'departmentsfinancereports2024quarterlysummary' in current_name

        # Improved maintains readable hierarchy
        assert '.departments.' in improved_name
        assert '.finance.' in improved_name
        assert '.reports.' in improved_name
        assert '.2024.' in improved_name


class TestEdgeCases:
    """Test edge cases and unusual inputs."""

    def test_empty_components(self):
        """Test handling of empty path components."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        # Test with empty folder name (shouldn't happen but test anyway)
        assert current.register('excel', '//data.xlsx', 'Sheet1')
        assert improved.register('excel', '//data.xlsx', 'Sheet1')

    def test_unicode_handling(self):
        """Test Unicode characters in paths."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        # Unicode in path
        unicode_file = '日本/データ.xlsx'
        current_name = current.register('excel', unicode_file, 'シート1')
        improved_name = improved.register('excel', unicode_file, 'シート1')

        # Both should handle by removing non-ASCII
        assert all(ord(c) < 128 for c in current_name)
        assert all(ord(c) < 128 for c in improved_name)

    def test_very_long_paths(self):
        """Test handling of very long paths."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        # Create extremely long path
        long_path = '/'.join(['folder'] * 20) + '/file.xlsx'

        current_name = current.register('excel', long_path, 'Sheet1')
        improved_name = improved.register('excel', long_path, 'Sheet1')

        # Both should truncate to max length
        assert len(current_name) <= 63
        assert len(improved_name) <= 63

    def test_numeric_prefixes(self):
        """Test handling of numeric prefixes."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        # File starting with number
        name1_current = current.register('excel', '2024/reports.xlsx', 'Sheet1')
        name1_improved = improved.register('excel', '2024/reports.xlsx', 'Sheet1')

        # Should add prefix to avoid starting with digit
        assert not name1_current[0].isdigit()
        assert not name1_improved[0].isdigit()

    def test_windows_path_separators(self):
        """Test handling of Windows-style paths."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        # Windows path with backslashes
        windows_path = 'folder\\subfolder\\data.xlsx'
        # Convert to forward slashes for consistency
        normalized_path = windows_path.replace('\\', '/')

        current_name = current.register('excel', normalized_path, 'Sheet1')
        improved_name = improved.register('excel', normalized_path, 'Sheet1')

        assert 'foldersubfolderdata' in current_name or 'folder_subfolder_data' in current_name
        assert '.folder.subfolder.' in improved_name


class TestSQLCompatibility:
    """Test that generated names are SQL-compatible."""

    def test_quoting_requirement(self):
        """All generated names should be quotable in SQL."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        test_cases = [
            'data.xlsx',
            'folder/data.xlsx',
            'my folder/my-file.xlsx',
            '2024/reports.xlsx',
            'dept-01/région/données.xlsx',
        ]

        for file in test_cases:
            current_name = current.register('excel', file, 'Sheet1')
            improved_name = improved.register('excel', file, 'Sheet1')

            # Names should be valid when quoted
            # (contain only alphanumeric, underscore, dot)
            valid_chars = set('abcdefghijklmnopqrstuvwxyz0123456789_.')
            assert all(c in valid_chars for c in current_name)
            assert all(c in valid_chars for c in improved_name)

    def test_reserved_words_handling(self):
        """Test handling of SQL reserved words."""
        current = TableRegistry()
        improved = ImprovedTableRegistry()

        # File names that are SQL reserved words
        reserved_words = ['select', 'from', 'where', 'table', 'order']

        for word in reserved_words:
            file = f'{word}/{word}.xlsx'
            current_name = current.register('excel', file, word.title())
            improved_name = improved.register('excel', file, word.title())

            # Should generate valid names despite reserved words
            assert len(current_name) > 0
            assert len(improved_name) > 0
            # Names are lowercase, which helps avoid keyword conflicts
            assert word in current_name.lower()
            assert word in improved_name.lower()


@pytest.mark.parametrize("registry_class", [TableRegistry, ImprovedTableRegistry])
class TestRegistryMethods:
    """Test common registry methods for both implementations."""

    def test_clear_method(self, registry_class):
        """Test that clear() removes all registered names."""
        registry = registry_class()

        # Register some names
        registry.register('excel', 'file1.xlsx', 'Sheet1')
        registry.register('excel', 'file2.xlsx', 'Sheet1')

        # Clear should reset everything
        registry.clear()

        # Same name should be available again
        name1 = registry.register('excel', 'file1.xlsx', 'Sheet1')
        name2 = registry.register('excel', 'file1.xlsx', 'Sheet1')

        # Should get collision since we registered twice
        assert name1 != name2
        assert name2.endswith('_2')

    def test_thread_safety(self, registry_class):
        """Test thread-safe registration."""
        import threading

        registry = registry_class()
        names = []

        def register_names():
            for i in range(10):
                name = registry.register('excel', f'file{i}.xlsx', 'Sheet1')
                names.append(name)

        # Run multiple threads
        threads = [threading.Thread(target=register_names) for _ in range(5)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # Should have registered 50 names
        assert len(names) == 50
        # All names should be unique
        assert len(set(names)) == len(names)