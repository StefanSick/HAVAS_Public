# Print Media Automation

Automatically fills columns  the `data_input` sheet of the Premium Print Bibel Excel file by looking up circulation and audience data from `Print2025` via a title-to-row mapping defined in `data_filter`.

## How it works

1. **`data_filter`** maps each publication title to a start cell in `Print2025`, which marks the Monday row of that publication's weekly data block.
2. For each row in **`data_input`**, the script reads the title and date, converts the date to a German weekday name, and searches forward from the start cell until it finds the matching weekday row.
3. Four values are copied from that row into `data_input`:

| Field |
|---|
| Verbreitete Auflage |
| NRW in % |
| Bruttokontakte HHF 18-59 |
| Bruttokontakte A14+ |

If a title is not found in `data_filter`, or the publication does not run on the given weekday, the target cells are left empty. There will be a sheet `Skipped` added to the excel file stating problems that occurred.

## Project structure

```
HAVAS/
├── config.py               # File paths and configuration constants
├── fill_excel.py           # Main script
├── test_fill_excel.py      # Unit tests
├── requirements.txt        # Python dependencies
└── data/
    ├── Premium_Print_Bibel_2025-BewerberAufgabe.xlsx   # Source file
    └── output.xlsx                                     # Generated output
```

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python fill_excel.py
```

Output is written to `data/output.xlsx`.

## Running the tests

```bash
pytest test_fill_excel.py -v
```

## Configuration

All tuneable values in `config.py`:

| Constant | Default | Description |
|---|---|---|
| `EXCEL_PATH` | `data/Premium_Print_Bibel_…xlsx` | Source file |
| `OUTPUT_PATH` | `data/output.xlsx` | Output file |
