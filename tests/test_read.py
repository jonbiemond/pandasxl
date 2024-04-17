"""Tests for reading data from Excel to a DataFrame."""

from pathlib import Path

import numpy as np
import pandas as pd
import pytest

import pandasxl


class TestRanges:
    """Tests for reading data from Excel ranges."""

    @pytest.mark.parametrize(
        "name, rtype, expected, header",
        [
            ("named_range1", pd.Series, pd.Series(["a", "b", "c"]), None),
            (
                "named_range2",
                pd.DataFrame,
                pd.DataFrame(np.array([["a", 1], ["b", 2], ["c", 3]])),
                None,
            ),
            ("named_range3", pd.Series, pd.Series(["a", "b", "c"]), None),
            (
                "named_range4",
                pd.DataFrame,
                pd.DataFrame({"A": ["a", "b", "c"], "B": ["1", "2", "3"]}),
                True,
            ),
            ("named_range5", str, "a", None),
            ("named_range6", pd.Series, pd.Series(["1", "2", "3"], name="A"), True),
            (
                "table1",
                pd.DataFrame,
                pd.DataFrame({"A": ["a", "b", "c"], "B": ["1", "2", "3"]}),
                None,
            ),
            ("table2", pd.Series, pd.Series(["a", "b", "c"], name="A"), None),
        ],
    )
    def test_from_name(self, name, rtype, expected, header):
        """Test the from_name function."""
        wb = pandasxl.excel.open_workbook(
            Path("tests/test_files/ranges_and_tables.xlsx")
        )
        data = pandasxl.read.from_name(wb, name, header)
        assert isinstance(data, rtype)
        if rtype == str:
            assert data == expected
        else:
            assert data.equals(expected)

    def test_from_reference(self):
        """Test the from_reference function."""
        wb = pandasxl.excel.open_workbook(
            Path("tests/test_files/ranges_and_tables.xlsx")
        )
        data = pandasxl.read.from_reference(wb, "Sheet1!$A$2:$A$4")
        expected = pd.Series(["a", "b", "c"])
        assert data.equals(expected)
