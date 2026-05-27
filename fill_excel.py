import logging
import re
import sys

import openpyxl
import pandas as pd
from pathlib import Path

from config import EXCEL_PATH, OUTPUT_PATH

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# German weekdays
WEEKDAYS: dict[int, str] = {
    0: "Montag",
    1: "Dienstag",
    2: "Mittwoch",
    3: "Donnerstag",
    4: "Freitag",
    5: "Samstag",
    6: "Sonntag",
}

VALID_WEEKDAY_NAMES: set[str] = set(WEEKDAYS.values())


# Loading helpers
# ---------------------------------------------------------------------------

def load_filter_map(xf: pd.ExcelFile) -> dict[str, str]:
    """Build a title -> start-cell dict from the data_filter sheet."""
    df = xf.parse(sheet_name="data_filter", header=None)
    return {str(row[0]).strip(): str(row[1]).strip() for _, row in df.iterrows()}


def load_print2025(xf: pd.ExcelFile) -> pd.DataFrame:
    """Load Print2025 as a raw DataFrame for row-level access."""
    return xf.parse(sheet_name="Print2025", header=None)


def load_data_input(xf: pd.ExcelFile) -> pd.DataFrame:
    """
    Load and clean data_input.

    Strips whitespace from column names (the ET column has a leading space in
    the source file) and parses the ET column as datetime.
    Invalid or missing dates become NaT instead of raising.
    """
    df = xf.parse(sheet_name="data_input")
    df.columns = df.columns.str.strip()
    df["ET"] = pd.to_datetime(df["ET"], errors="coerce")
    return df


# Validation
# ---------------------------------------------------------------------------

def validate_data_input(df: pd.DataFrame, filter_map: dict[str, str]) -> dict[int, str]:
    """
    Return a {row_index: reason} dict for rows that cannot be processed.

    Logs a warning for each skipped row. D:G will be left blank for these rows.
    Checks ET first; only checks title when ET is valid, which prevents a NaN
    title from being double-reported as both 'missing' and 'not found'.
    """
    skip: dict[int, str] = {}

    for idx, row in df.iterrows():
        excel_row = idx + 2  # +1 for 1-based indexing, +1 for header row

        if pd.isna(row["ET"]):
            reason = "missing or invalid date"
            logging.warning(f"Row {excel_row}: {reason} - D:G left blank")
            skip[idx] = reason
            continue

        title = "" if pd.isna(row["Titel"]) else str(row["Titel"]).strip()

        if title == "":
            reason = "missing title"
            logging.warning(f"Row {excel_row}: {reason} - D:G left blank")
            skip[idx] = reason
        elif title not in filter_map:
            reason = f"title '{title}' not found in data_filter"
            logging.warning(f"Row {excel_row}: {reason} - D:G left blank")
            skip[idx] = reason

    return skip


# Lookup logic
# ---------------------------------------------------------------------------

def parse_start_row(cell_ref: str) -> int:
    """
    Convert an Excel cell reference like 'C83' to a 0-based pandas row index.

    The column letter is ignored.
    """
    match = re.search(r"\d+$", cell_ref)
    if match is None:
        raise ValueError(f"No row number found in cell reference: {cell_ref!r}")
    return int(match.group()) - 1  # Excel rows are 1-indexed; pandas rows are 0-indexed


def find_weekday_row(
    print_df: pd.DataFrame, start_row: int, german_day: str
) -> int | None:
    """
    Search Print2025 downward from start_row for a row whose weekday column matches german_day.

    Stops as soon as a non-weekday value is encountered or after 7 rows (Mon-Sun).
    Returns the 0-based row index, or None if the publication does not run on that day.
    """
    for offset in range(7): 
        row_idx = start_row + offset
        if row_idx >= len(print_df):
            break

        weekday_cell = print_df.iloc[row_idx, 1]  # column B

        # Stop if it left the data block
        if pd.isna(weekday_cell) or str(weekday_cell).strip() not in VALID_WEEKDAY_NAMES:
            break

        if str(weekday_cell).strip() == german_day:
            return row_idx

    return None


def extract_print_values(
    print_df: pd.DataFrame, row_idx: int
) -> dict[str, float | None]:
    """
    Extract the four required values from Print2025 at the given row index.

    Print2025 col (0-based) -> result key -> written to data_input col:
        3 (D) -> verbreitete_auflage -> col 4 (D)
        8 (I) -> nrw_pct             -> col 5 (E)
        5 (F) -> brutto_hhf          -> col 6 (F)
        4 (E) -> brutto_a14          -> col 7 (G)
    """
    row = print_df.iloc[row_idx]
    return {
        "verbreitete_auflage": row[3],
        "nrw_pct":             row[8],
        "brutto_hhf":          row[5],
        "brutto_a14":          row[4],
    }


# Main pipeline
# ---------------------------------------------------------------------------

def fill_data_input(
    excel_path: Path, output_path: Path | None = None
) -> pd.DataFrame:
    """
    Fill columns D-G in data_input by looking up each row's title and date in
    Print2025 via the data_filter mapping, then save the result.

    Rows with missing/invalid dates or titles absent from data_filter are
    skipped: their D:G cells are left blank, a warning is logged, and they
    are written to a separate 'Skipped' sheet in the output file.

    Parameters
    ----------
    excel_path  : path to the source .xlsx file
    output_path : where to write the result (defaults to output.xlsx next to source)

    Returns
    -------
    The updated data_input DataFrame.
    """
    if output_path is None:
        output_path = excel_path.parent / "output.xlsx"

    try:
        xf = pd.ExcelFile(excel_path)
    except FileNotFoundError:
        logging.error(f"Source file not found: {excel_path}")
        raise

    with xf:
        filter_map = load_filter_map(xf)
        print_df   = load_print2025(xf)
        df_input   = load_data_input(xf)

    skip_rows = validate_data_input(df_input, filter_map)

    # Pre-fill output columns with None; valid rows overwrite below
    df_input["Verbreitete Auflage"] = None
    df_input["NRW in %"] = None
    df_input["Bruttokontakte HHF18-59"] = None
    df_input["Bruttokontakte A14+"] = None

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["data_input"]

    for df_idx, row in df_input.iterrows():
        excel_row = df_idx + 2  # row 1 is the header; data starts at row 2

        if df_idx in skip_rows:
            for col in (4, 5, 6, 7):  # columns D-G
                ws.cell(row=excel_row, column=col).value = None
            continue

        title = str(row["Titel"]).strip()
        date = row["ET"]
        start_cell = filter_map[title]
        start_row = parse_start_row(start_cell)
        german_day = WEEKDAYS[date.weekday()]
        matched_row = find_weekday_row(print_df, start_row, german_day)

        if matched_row is None:
            reason = f"'{title}' has no {german_day} entry in Print2025 - check the booking date"
            logging.warning(f"Row {excel_row}: {reason} - D:G left blank")
            skip_rows[df_idx] = reason
            for col in (4, 5, 6, 7):  # columns D-G
                ws.cell(row=excel_row, column=col).value = None
            continue

        values = extract_print_values(print_df, matched_row)

        df_input.at[df_idx, "Verbreitete Auflage"] = values["verbreitete_auflage"]
        df_input.at[df_idx, "NRW in %"] = values["nrw_pct"]
        df_input.at[df_idx, "Bruttokontakte HHF18-59"] = values["brutto_hhf"]
        df_input.at[df_idx, "Bruttokontakte A14+"] = values["brutto_a14"]

        ws.cell(row=excel_row, column=4).value = values["verbreitete_auflage"]  # D
        ws.cell(row=excel_row, column=5).value = values["nrw_pct"] # E
        ws.cell(row=excel_row, column=6).value = values["brutto_hhf"] # F
        ws.cell(row=excel_row, column=7).value = values["brutto_a14"] # G

    # Write skipped rows to a separate sheet so they can be reviewed
    if "Skipped" in wb.sheetnames:
        del wb["Skipped"]  # clear from previous run

    if skip_rows:
        ws_skip = wb.create_sheet("Skipped")
        ws_skip.append(["Excel Row", "Titel", "ET", "Reason"])
        for df_idx, reason in skip_rows.items():
            source = df_input.iloc[df_idx]
            titel  = "" if pd.isna(source["Titel"]) else source["Titel"]
            et     = "" if pd.isna(source["ET"])    else source["ET"].strftime("%d.%m.%Y")
            ws_skip.append([df_idx + 2, titel, et, reason])
        logging.info(f"{len(skip_rows)} skipped row(s) written to 'Skipped' sheet.")

    try:
        wb.save(output_path)
    except PermissionError:
        logging.error(f"Cannot write to '{output_path}' - is the file open in Excel?")
        raise

    logging.info(f"Output written to: {output_path}")
    return df_input


if __name__ == "__main__":
    try:
        df_result = fill_data_input(EXCEL_PATH, OUTPUT_PATH)
    except (FileNotFoundError, PermissionError, ValueError):
        sys.exit(1)

    cols = [
        "Titel", "ET",
        "Verbreitete Auflage", "NRW in %",
        "Bruttokontakte HHF18-59", "Bruttokontakte A14+",
    ]
    logging.info("\nFilled data_input columns:\n" + df_result[cols].to_string(index=False))
