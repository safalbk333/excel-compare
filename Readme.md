# Excel Compare

Load two Excel files into SQLite, inspect column data, then match rows between the two sheets and export a comparison report (pass/fail per column).

## Requirements

- **Python 3.12+** recommended (pandas/openpyxl; avoid very new 3.14 if packages fail to install)
- Dependencies: `pandas`, `openpyxl` (see `requirements.txt`)

## Setup

From the project folder:

```powershell
cd c:\Users\Work\Desktop\bk\python\excel-compare

# Create virtual environment (example: Python 3.12)
py -3.12 -m venv env312

# Activate (PowerShell)
.\env312\Scripts\Activate.ps1

# Install packages
python -m pip install -r requirements.txt
```

Place your input files in the project directory (defaults):

| File | Default path |
|------|----------------|
| Excel 1 | `input1.xlsx` |
| Excel 2 | `input2.xlsx` |

## Workflow (3 steps)

```
input1.xlsx + input2.xlsx  →  main.py  →  excel_data.db
                                              ↓
                                         stage2.py (optional analysis)
                                              ↓
                                         stage3_match.py  →  match_results.xlsx
```

---

### Step 1 — Import Excel into SQLite (`main.py`)

```powershell
python main.py
```

**Common options**

| Option | Description |
|--------|-------------|
| `--excel1 path.xlsx` | First file (default: `input1.xlsx`) |
| `--excel2 path.xlsx` | Second file (default: `input2.xlsx`) |
| `--db excel_data.db` | SQLite output path |
| `--list-sheets` | List sheet names in both files |
| `--sheet NAME_OR_INDEX` | Same sheet for both files |
| `--sheet1` / `--sheet2` | Sheet per file |
| `--interactive` | Prompt to pick sheets |
| `--all-sheets` | Import every sheet (one table per sheet) |

**Examples**

```powershell
# List sheets
python main.py --list-sheets

# Pick sheets by name
python main.py --sheet1 Input1 --sheet2 Output

# Or by index (0 = first sheet)
python main.py --sheet1 0 --sheet2 0
```

After import, note the **table names** printed in the console, e.g. `excel1_input1` and `excel2_output`. Names are `excel1_<sheet>` / `excel2_<sheet>` (sanitized). Use these in step 3 if they differ from the defaults.

---

### Step 2 — Column analysis (optional, `stage2.py`)

Distinct values per column and which column has the most distinct values:

```powershell
python stage2.py --table excel1_input1
python stage2.py --table excel2_input2 --max-print 10
python stage2.py --json distinct_report.json
```

---

### Step 3 — Match rows and export (`stage3_match.py`)

Matches each row in Excel 1 to at most one row in Excel 2, then compares every shared column.

```powershell
python stage3_match.py
```

**Defaults**

- Database: `excel_data.db`
- Tables: `excel1_input1`, `excel2_input2` (change with `--table1` / `--table2` if your import used other sheet names)
- Output: `match_results.xlsx`, `match_summary.csv`

**Useful options**

| Option | Description |
|--------|-------------|
| `--limit N` | Process only first N rows from table 1 (0 = all) |
| `--row ROWID` | Process one SQLite `rowid` from table 1 |
| `--workers N` | Parallel threads (default: ~2× CPU cores; use `1` to disable) |
| `--excel path.xlsx` | Custom Excel report path |
| `--no-excel` | Skip Excel; keep CSV only |
| `--json path.json` | Full JSON report |

**Example — full run with custom tables**

```powershell
python stage3_match.py --table1 excel1_input1 --table2 excel2_input2 --workers 8
```

---

## Output files

### `match_results.xlsx`

| Sheet | Contents |
|-------|----------|
| **statistics** | Total / matched / unmatched rows and percentages |
| **column_statistics** | Per-column pass/fail counts and match % |
| **comparison** | One row per matched pair: `Id-excel1`, `Id-excel2`, then for each column: value, value, `pass`/`fail` |
| **all_rows** | Every processed row and match status |

Row IDs are SQLite **`rowid`** values from the imported tables (not necessarily the original Excel row number).

### `match_summary.csv`

Per-row summary: ids, status, column match counts.

---

## Matching logic (short)

1. Columns are paired **by name** (must exist in both tables).
2. Columns are tried in order of **highest distinct count** in table 1 to narrow candidates in table 2.
3. Empty values are not used as filters.
4. When exactly one row remains in table 2, that pair is **matched**; all shared columns are compared (`pass` / `fail`).

---

## Project layout

| File | Role |
|------|------|
| `main.py` | Excel → SQLite |
| `stage2.py` | Distinct values / column stats |
| `stage3_match.py` | Row matching + reports |
| `requirements.txt` | Python dependencies |

Generated files (`excel_data.db`, `match_results.xlsx`, etc.) are listed in `.gitignore`.

## Troubleshooting

- **`ModuleNotFoundError: pandas`** — activate the venv and run `pip install -r requirements.txt`.
- **Wrong table names** — re-run `main.py` and copy the table names from the “Done. Tables: …” line, then pass `--table1` / `--table2` to `stage3_match.py`.
- **Few matched rows** — many rows may be **ambiguous** (multiple candidates in Excel 2); check the `all_rows` sheet and statistics.
