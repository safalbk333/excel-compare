"""Report distinct values for each column in a SQLite table."""

import argparse
import json
import sqlite3
from pathlib import Path


def quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def get_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    rows = conn.execute(f"PRAGMA table_info({quote_ident(table)})").fetchall()
    if not rows:
        tables = [
            r[0]
            for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            )
        ]
        raise SystemExit(
            f"Table '{table}' not found. Tables in database: {tables or '(none)'}"
        )
    return [row[1] for row in rows]


def distinct_values_for_column(
    conn: sqlite3.Connection, table: str, column: str
) -> list[str | None]:
    sql = (
        f"SELECT DISTINCT {quote_ident(column)} AS v "
        f"FROM {quote_ident(table)} "
        f"ORDER BY v IS NULL, v"
    )
    return [row[0] for row in conn.execute(sql)]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="List distinct values for every column in a SQLite table."
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=Path("excel_data.db"),
        help="SQLite database path (default: excel_data.db)",
    )
    parser.add_argument(
        "--table",
        default="excel1_input1",
        help="Table name (default: excel1_input1)",
    )
    parser.add_argument(
        "--json",
        type=Path,
        metavar="FILE",
        help="Write results to a JSON file",
    )
    parser.add_argument(
        "--max-print",
        type=int,
        default=3,
        help="Max distinct values to print per column (default: 3; 0 = all)",
    )
    args = parser.parse_args()

    if not args.db.is_file():
        raise SystemExit(f"Database not found: {args.db.resolve()}")

    conn = sqlite3.connect(args.db)
    try:
        columns = get_columns(conn, args.table)
        report: dict[str, dict[str, object]] = {}

        print(f"Table: {args.table} ({args.db.resolve()})\n")

        for col in columns:
            values = distinct_values_for_column(conn, args.table, col)
            non_null = [v for v in values if v is not None and str(v).strip() != ""]
            report[col] = {
                "distinct_count": len(values),
                "non_empty_count": len(non_null),
                # "values": values,
            }

            print(f"=== {col} ({len(values)} distinct) ===")
            to_show = values if args.max_print == 0 else values[: args.max_print]
            for v in to_show:
                display = "<NULL>" if v is None else str(v)
                print(f"  {display}")
            if args.max_print and len(values) > args.max_print:
                print(f"  ... and {len(values) - args.max_print} more")
            print()

        counts = {col: int(report[col]["distinct_count"]) for col in columns}
        max_count = max(counts.values())
        max_columns = sorted(col for col, n in counts.items() if n == max_count)

        print("--- Distinct count by column (highest first) ---")
        for col, n in sorted(counts.items(), key=lambda item: (-item[1], item[0])):
            tag = "  [MAX]" if n == max_count else ""
            print(f"  {col}: {n}{tag}")
        print(
            f"\nLargest distinct count: {max_count} "
            f"in column(s): {', '.join(max_columns)}"
        )

        summary = {
            "distinct_count_by_column": counts,
            "max_distinct_count": max_count,
            "max_distinct_columns": max_columns,
        }

        if args.json:
            payload = {"columns": report, "summary": summary}
            args.json.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
            )
            print(f"\nWrote JSON: {args.json.resolve()}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
