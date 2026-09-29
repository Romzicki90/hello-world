import datetime as dt

import pytest

from auditwb.ingest import normalise_id, parse_date, parse_number, read_rows
from auditwb.loader import load_library


@pytest.mark.parametrize("text, notation, expected", [
    ("1,234,567.89", "1,234,567.89", 1234567.89),
    ("12,34,567.89", "1,234,567.89", 1234567.89),       # Indian grouping
    ("2,36,000.00-", "1,234,567.89", -236000.0),        # SAP trailing minus
    ("(1,000.50)", "1,234,567.89", -1000.5),
    ("1.234.567,89", "1.234.567,89", 1234567.89),
    ("1 234 567,89-", "1 234 567,89", -1234567.89),
    ("Rs. 5,000", "1,234,567.89", 5000.0),
    ("", "1,234,567.89", None),
    (42, "1,234,567.89", 42.0),
])
def test_parse_number(text, notation, expected):
    assert parse_number(text, notation) == expected


def test_parse_number_rejects_text():
    with pytest.raises(ValueError):
        parse_number("abc", "1,234,567.89")


@pytest.mark.parametrize("value, fmt, expected", [
    ("31.03.2026", "DD.MM.YYYY", dt.date(2026, 3, 31)),
    ("03/31/2026", "MM/DD/YYYY", dt.date(2026, 3, 31)),
    ("20260331", "DD.MM.YYYY", dt.date(2026, 3, 31)),   # SE16N internal format
    (20260331, "DD.MM.YYYY", dt.date(2026, 3, 31)),
    ("00.00.0000", "DD.MM.YYYY", None),
    (dt.datetime(2026, 3, 31, 10, 0), "DD.MM.YYYY", dt.date(2026, 3, 31)),
    (46112, "DD.MM.YYYY", dt.date(2026, 3, 31)),         # Excel serial date
])
def test_parse_date(value, fmt, expected):
    assert parse_date(value, fmt) == expected


def test_ambiguous_date_is_not_guessed():
    with pytest.raises(ValueError):
        parse_date("31/03/2026", "MM/DD/YYYY")


def test_normalise_id():
    assert normalise_id("0000100005") == "100005"
    assert normalise_id(4500000001.0) == "4500000001"
    assert normalise_id("0001", keep_zeros=True) == "0001"
    assert normalise_id("WBS P-1001", ["WBS "]) == "P-1001"
    assert normalise_id("  ") is None


def test_unconverted_list_file(tmp_path):
    p = tmp_path / "list.txt"
    p.write_text("Title\n-----\n|Material Document|Item|\n-----\n|4900000001|0001|\n-----\n", encoding="utf-8")
    rows = [r for r in read_rows(p) if r]
    assert rows[1] == ["Material Document", "Item"]
    assert rows[2] == ["4900000001", "0001"]


def test_library_is_consistent():
    lib = load_library()
    assert len(lib.topics) == 166
    assert all(t.topics for t in lib.tests.values())
