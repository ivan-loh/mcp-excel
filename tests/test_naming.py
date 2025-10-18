import pytest
from mcp_excel.naming import TableRegistry

pytestmark = pytest.mark.unit


def test_basic_sanitization():
    registry = TableRegistry()
    name = registry.register("excel", "sales/report.xlsx", "Summary")
    assert name == "excel.salesreport.summary"


def test_special_chars():
    registry = TableRegistry()
    name = registry.register("excel", "data/Q1-2024 (Final).xlsx", "Sheet1")
    assert name == "excel.dataq12024_final.sheet1"


def test_unicode():
    registry = TableRegistry()
    name = registry.register("excel", "données/rapport.xlsx", "Feuille")
    assert name == "excel.donnesrapport.feuille"


def test_leading_numbers():
    registry = TableRegistry()
    name = registry.register("excel", "2024/report.xlsx", "1stQuarter")
    assert name.startswith("t_")


def test_collision_handling():
    registry = TableRegistry()
    name1 = registry.register("excel", "sales/report.xlsx", "Summary")
    name2 = registry.register("excel", "sales/report.xlsx", "Summary")
    assert name1 != name2
    assert name2.endswith("_2")


def test_multiple_collisions():
    registry = TableRegistry()
    name1 = registry.register("excel", "data.xlsx", "Sheet")
    name2 = registry.register("excel", "data.xlsx", "Sheet")
    name3 = registry.register("excel", "data.xlsx", "Sheet")
    assert name1 == "excel.data.sheet"
    assert name2 == "excel.data.sheet_2"
    assert name3 == "excel.data.sheet_3"


def test_long_names():
    registry = TableRegistry()
    long_relpath = "a" * 100
    name = registry.register("excel", f"{long_relpath}.xlsx", "Sheet")
    assert len(name) <= 64


def test_empty_components():
    registry = TableRegistry()
    name = registry.register("excel", ".xlsx", "")
    assert name
    assert "." in name


def test_clear():
    registry = TableRegistry()
    name1 = registry.register("excel", "test.xlsx", "Sheet")
    registry.clear()
    name2 = registry.register("excel", "test.xlsx", "Sheet")
    assert name1 == name2
