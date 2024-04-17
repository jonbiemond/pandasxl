"""Tests for the Excel utility functions."""

from pathlib import Path

import pytest
from openpyxl import Workbook

import pandasxl


def test_open_workbook():
    """Test the open_workbook function."""
    wb = pandasxl.excel.open_workbook(Path("tests/test_files/ranges_and_tables.xlsx"))
    assert isinstance(wb, Workbook)


class TestRanges:
    """Tests for working with named ranges, references and tables in Excel."""

    @pytest.fixture()
    def wb(self):
        """Fixture to open the workbook."""
        return pandasxl.excel.open_workbook(
            Path("tests/test_files/ranges_and_tables.xlsx")
        )

    def test_named_ranges(self, wb: Workbook):
        """Test the named_ranges function."""
        ranges = pandasxl.excel.named_ranges(wb)
        assert ranges == {
            "named_range1": ("Sheet1", "$A$2:$A$4"),
            "named_range2": ("Sheet1", "$C$2:$D$4"),
            "named_range3": ("Sheet1", "$F$2:$H$2"),
            "named_range4": ("Sheet1", "$J$2:$K$5"),
            "named_range5": ("Sheet1", "$M$2"),
            "named_range6": ("Sheet1", "$O$2:$O$5"),
        }

    def test_table_ranges(self, wb: Workbook):
        """Test the table_ranges function."""
        ranges = pandasxl.excel.table_ranges(wb)
        assert ranges == {
            "table1": ("Sheet1", "Q2:R5"),
            "table2": ("Sheet1", "T2:T5"),
        }

    def test_range_to_array(self, wb: Workbook):
        """Test the range_to_array function."""
        excel_range = pandasxl.excel.named_ranges(wb)["named_range1"]
        region = pandasxl.excel.range_to_array(excel_range, wb)
        values = [cell.value for row in region for cell in row]
        assert values == ["a", "b", "c"]

    def test_scalar_range_to_array(self, wb: Workbook):
        """Test the range_to_array function with a one cell range."""
        excel_range = pandasxl.excel.named_ranges(wb)["named_range5"]
        region = pandasxl.excel.range_to_array(excel_range, wb)
        values = [cell.value for row in region for cell in row]
        assert values == ["a"]

    def test_reference_to_array(self, wb: Workbook):
        """Test the reference_to_array function."""
        region = pandasxl.excel.reference_to_array("Sheet1!A2:A4", wb)
        values = [cell.value for row in region for cell in row]
        assert values == ["a", "b", "c"]
