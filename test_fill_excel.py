
import logging
import pytest
import pandas as pd

from fill_excel import parse_start_row, find_weekday_row, validate_data_input



# parse_start_row
# ---------------------------------------------------------------------------


def test_parse_start_row_first_data_row():
    assert parse_start_row("C6") == 5

def test_parse_start_row_three_digit():
    assert parse_start_row("C106") == 105

def test_parse_start_row_two_digit():
    assert parse_start_row("C40") == 39



# find_weekday_row
# ---------------------------------------------------------------------------

@pytest.fixture
def weekday_block() -> pd.DataFrame:
    """
    Mock of a Print2025 weekly block with Monday–Friday data.
    Column 0 = date label, column 1 = German weekday name
    """
    return pd.DataFrame({
        0: [
            "Mo., 02. Oktober 2023",
            "Di., 03. Oktober 2023",
            "Mi., 04. Oktober 2023",
            "Do., 05. Oktober 2023",
            "Fr., 06. Oktober 2023",
        ],
        1: ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"],
    })


def test_find_weekday_row_first_row(weekday_block):
    assert find_weekday_row(weekday_block, 0, "Montag") == 0

def test_find_weekday_row_with_offset(weekday_block):
    assert find_weekday_row(weekday_block, 0, "Donnerstag") == 3

def test_find_weekday_row_last_row(weekday_block):
    assert find_weekday_row(weekday_block, 0, "Freitag") == 4

def test_find_weekday_row_day_not_in_block(weekday_block):
    # Samstag  not in the fixture — should return None
    assert find_weekday_row(weekday_block, 0, "Samstag") is None

def test_find_weekday_row_start_beyond_dataframe(weekday_block):
    # start_row past the end of the DataFrame → immediate None
    assert find_weekday_row(weekday_block, 99, "Montag") is None

def test_find_weekday_row_partial_block():
    """Publication that starts on Tuesday (no Monday edition)."""
    df = pd.DataFrame({
        0: ["Di., 03. Oktober 2023", "Mi., 04. Oktober 2023", "Do., 05. Oktober 2023"],
        1: ["Dienstag", "Mittwoch", "Donnerstag"],
    })
    # Monday is not available → None
    assert find_weekday_row(df, 0, "Montag") is None
    # Thursday is available at offset 2
    assert find_weekday_row(df, 0, "Donnerstag") == 2



# validate_data_input
# ---------------------------------------------------------------------------

FILTER_MAP = {"Zeitung Arena": "C6", "Zeitung A": "C40"}


def _make_df(titles, dates):
    return pd.DataFrame({
        "Titel": titles,
        "ET":    pd.to_datetime(dates, errors="coerce"),
    })


def test_validate_passes_for_clean_data():
    df = _make_df(["Zeitung Arena", "Zeitung A"], ["2025-06-06", "2025-06-09"])
    assert validate_data_input(df, FILTER_MAP) == {}


def test_validate_skips_missing_et():
    df = _make_df(["Zeitung Arena"], [None])
    assert set(validate_data_input(df, FILTER_MAP)) == {0}


def test_validate_skips_missing_titel():
    df = _make_df([None], ["2025-06-06"])
    assert set(validate_data_input(df, FILTER_MAP)) == {0}


def test_validate_skips_empty_titel():
    df = _make_df(["   "], ["2025-06-06"])
    assert set(validate_data_input(df, FILTER_MAP)) == {0}


def test_validate_skips_unknown_titel():
    df = _make_df(["Unbekannte Zeitung"], ["2025-06-06"])
    assert set(validate_data_input(df, FILTER_MAP)) == {0}


def test_validate_returns_all_bad_rows():
    """Row 0 has missing ET, row 1 has unknown title — both must be in the returned dict."""
    df = _make_df(
        ["Zeitung Arena", "Unbekannte Zeitung"],
        [None,            "2025-06-06"],
    )
    assert set(validate_data_input(df, FILTER_MAP)) == {0, 1}


def test_validate_excel_row_number_in_log(caplog):
    """Row index 1 (0-based) → Excel row 3; the warning must name it."""
    df = _make_df(["Zeitung Arena", "Zeitung Arena"], ["2025-06-06", None])
    with caplog.at_level(logging.WARNING):
        validate_data_input(df, FILTER_MAP)
    assert "Row 3" in caplog.text
