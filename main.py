"""Read two Excel files and store each sheet's data in SQLite."""

import argparse
import re
import sqlite3
from pathlib import Path

import pandas as pd


def list_sheet_names(excel_path: Path) -> list[str]:
    with pd.ExcelFile(excel_path, engine="openpyxl") as book:
        return book.sheet_names


def parse_sheet_choice(value: str) -> str | int:
    """Sheet name, or 0-based index if value is a non-negative integer."""
    value = value.strip()
    if value.isdigit():
        return int(value)
    return value


def resolve_sheets_for_file(
    excel_path: Path,
    *,
    all_sheets: bool,
    sheet: str | None,
    file_sheet: str | None,
    interactive: bool,
) -> str | int | list[str | int] | None:
    if all_sheets:
        if sheet or file_sheet:
            raise SystemExit("Use only one of --all-sheets and --sheet / --sheet1 / --sheet2.")
        return None

    if file_sheet is not None:
        return parse_sheet_choice(file_sheet)
    if sheet is not None:
        return parse_sheet_choice(sheet)

    if interactive:
        names = list_sheet_names(excel_path)
        print(f"\nSheets in {excel_path.name}:")
        for i, name in enumerate(names):
            print(f"  [{i}] {name}")
        choice = input("Select sheet (index or name): ").strip()
        if not choice:
            raise SystemExit("No sheet selected.")
        picked = parse_sheet_choice(choice)
        if isinstance(picked, int) and not 0 <= picked < len(names):
            raise SystemExit(f"Sheet index out of range (0–{len(names) - 1}).")
        if isinstance(picked, str) and picked not in names:
            raise SystemExit(f"Unknown sheet '{picked}'. Available: {names}")
        return picked

    return 0


def print_sheet_catalog(excel1: Path, excel2: Path) -> None:
    for label, path in ("excel1", excel1), ("excel2", excel2):
        path = Path(path)
        if not path.is_file():
            print(f"{label} ({path}): file not found")
            continue
        names = list_sheet_names(path)
        print(f"{label} ({path.name}):")
        for i, name in enumerate(names):
            print(f"  [{i}] {name}")


def sanitize_table_name(name: str) -> str:
    """Make a valid SQLite table name from a file or sheet label."""
    name = re.sub(r"[^\w]", "_", name.strip())
    name = re.sub(r"_+", "_", name).strip("_")
    if not name:
        name = "data"
    if name[0].isdigit():
        name = f"t_{name}"
    return name.lower()


def excel_to_sqlite(
    excel_path: Path,
    conn: sqlite3.Connection,
    table_prefix: str,
    sheet_name: str | int | None = 0,
) -> list[str]:
    """
    Load an Excel file into SQLite. One table per sheet.
    Returns list of table names created.
    """
    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    sheets = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    created: list[str] = []

    if isinstance(sheets, pd.DataFrame):
        sheets = {excel_path.stem: sheets}

    for sheet_label, df in sheets.items():
        table = sanitize_table_name(f"{table_prefix}_{sheet_label}")
        df.columns = [sanitize_table_name(str(c)) for c in df.columns]
        df.to_sql(table, conn, if_exists="replace", index=False)
        created.append(table)
        print(f"  {excel_path.name} [{sheet_label}] -> table '{table}' ({len(df)} rows, {len(df.columns)} columns)")
        print(f"    columns: {list(df.columns)}")

    return created


def main() -> None:
    parser = argparse.ArgumentParser(description="Load Excel files into SQLite.")
    parser.add_argument(
        "--excel1",
        type=Path,
        default=Path("input1.xlsx"),
        help="First Excel file (default: excel1.xlsx)",
    )
    parser.add_argument(
        "--excel2",
        type=Path,
        default=Path("input2.xlsx"),
        help="Second Excel file (default: excel2.xlsx)",
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=Path("excel_data.db"),
        help="SQLite database path (default: excel_data.db)",
    )
    parser.add_argument(
        "--all-sheets",
        action="store_true",
        help="Import every sheet",
    )
    parser.add_argument(
        "--sheet",
        metavar="NAME_OR_INDEX",
        help="Sheet for both files (name or 0-based index). Default: first sheet",
    )
    parser.add_argument(
        "--sheet1",
        metavar="NAME_OR_INDEX",
        help="Sheet for excel1 only (overrides --sheet for file 1)",
    )
    parser.add_argument(
        "--sheet2",
        metavar="NAME_OR_INDEX",
        help="Sheet for excel2 only (overrides --sheet for file 2)",
    )
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="List sheet names in both Excel files and exit",
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Prompt to pick a sheet for each file (ignored if --sheet/--sheet1/--sheet2 set)",
    )
    args = parser.parse_args()

    if args.list_sheets:
        print_sheet_catalog(args.excel1, args.excel2)
        return

    sheet1 = resolve_sheets_for_file(
        args.excel1,
        all_sheets=args.all_sheets,
        sheet=args.sheet,
        file_sheet=args.sheet1,
        interactive=args.interactive,
    )
    sheet2 = resolve_sheets_for_file(
        args.excel2,
        all_sheets=args.all_sheets,
        sheet=args.sheet,
        file_sheet=args.sheet2,
        interactive=args.interactive,
    )

    args.db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(args.db)

    try:
        print(f"Database: {args.db.resolve()}\n")
        tables1 = excel_to_sqlite(args.excel1, conn, "excel1", sheet1)
        tables2 = excel_to_sqlite(args.excel2, conn, "excel2", sheet2)
        conn.commit()
        print(f"\nDone. Tables: {tables1 + tables2}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
