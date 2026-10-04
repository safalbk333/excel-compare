"""
Match rows from excel1 to excel2 using distinct-count column priority,
then compare every paired column cell-by-cell.
"""

import argparse
import csv
import json
import os
import sqlite3
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import pandas as pd

from stage2 import get_columns, quote_ident


def default_worker_count() -> int:
    cpus = os.cpu_count() or 4
    return max(1, min(32, cpus * 2))


def open_readonly_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(
        f"file:{db_path.resolve().as_posix()}?mode=ro",
        uri=True,
        check_same_thread=False,
    )
    conn.row_factory = sqlite3.Row
    return conn


def distinct_count_one(db_path: Path, table: str, column: str) -> tuple[str, int]:
    conn = open_readonly_db(db_path)
    try:
        sql = (
            f"SELECT COUNT(DISTINCT {quote_ident(column)}) "
            f"FROM {quote_ident(table)}"
        )
        return column, int(conn.execute(sql).fetchone()[0])
    finally:
        conn.close()


def distinct_counts(
    db_path: Path,
    table: str,
    columns: list[str],
    workers: int,
) -> dict[str, int]:
    if workers <= 1 or len(columns) < 4:
        conn = open_readonly_db(db_path)
        try:
            counts: dict[str, int] = {}
            for col in columns:
                sql = (
                    f"SELECT COUNT(DISTINCT {quote_ident(col)}) "
                    f"FROM {quote_ident(table)}"
                )
                counts[col] = int(conn.execute(sql).fetchone()[0])
            return counts
        finally:
            conn.close()

    counts: dict[str, int] = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(distinct_count_one, db_path, table, col) for col in columns]
        for fut in as_completed(futures):
            col, n = fut.result()
            counts[col] = n
    return counts


def norm_cell(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    return text


def cells_match(left: Any, right: Any) -> bool:
    return norm_cell(left) == norm_cell(right)


def fetch_row_dict(
    conn: sqlite3.Connection,
    table: str,
    rowid: int,
    colnames: list[str] | None = None,
) -> dict[str, Any]:
    row = conn.execute(
        f"SELECT * FROM {quote_ident(table)} WHERE rowid = ?", (rowid,)
    ).fetchone()
    if row is None:
        raise ValueError(f"rowid {rowid} not found in {table}")
    if colnames is None:
        colnames = [
            d[0]
            for d in conn.execute(f"SELECT * FROM {quote_ident(table)} LIMIT 0").description
        ]
    return dict(zip(colnames, row))


def query_table2_candidates(
    conn: sqlite3.Connection,
    table2: str,
    filters: list[tuple[str, Any]],
) -> list[int]:
    if not filters:
        sql = f"SELECT rowid FROM {quote_ident(table2)}"
        return [r[0] for r in conn.execute(sql)]
    clauses = []
    params: list[Any] = []
    for col, val in filters:
        if norm_cell(val) is None:
            clauses.append(f"{quote_ident(col)} IS NULL")
        else:
            clauses.append(f"{quote_ident(col)} = ?")
            params.append(val)
    where = " AND ".join(clauses)
    sql = f"SELECT rowid FROM {quote_ident(table2)} WHERE {where}"
    return [r[0] for r in conn.execute(sql, params)]


def find_match_rowid(
    conn: sqlite3.Connection,
    table2: str,
    row1: dict[str, Any],
    ranked_columns: list[str],
    cols2: set[str],
) -> tuple[int | None, list[tuple[str, Any]], str]:
    """
    Narrow table2 using row1 values, highest-distinct columns first.
    Returns (excel2_rowid or None, filters used, status message).
    """
    filters: list[tuple[str, Any]] = []

    for col in ranked_columns:
        if col not in cols2:
            continue

        val = row1.get(col)
        if norm_cell(val) is None:
            continue

        trial = filters + [(col, val)]
        candidates = query_table2_candidates(conn, table2, trial)

        if len(candidates) == 0:
            continue

        filters = trial

        if len(candidates) == 1:
            return candidates[0], filters, "matched"

    if not filters:
        return None, filters, "no_filter_matched"

    final = query_table2_candidates(conn, table2, filters)
    if len(final) == 1:
        return final[0], filters, "matched"
    if len(final) == 0:
        return None, filters, "no_candidate"
    return None, filters, f"ambiguous_{len(final)}_candidates"


def compare_row_pair(
    row1: dict[str, Any],
    row2: dict[str, Any],
    paired_columns: list[str],
) -> list[dict[str, Any]]:
    details: list[dict[str, Any]] = []
    for col in paired_columns:
        v1 = row1.get(col)
        v2 = row2.get(col)
        details.append(
            {
                "column": col,
                "excel1_value": v1,
                "excel2_value": v2,
                "match": cells_match(v1, v2),
            }
        )
    return details


def process_excel1_row(
    db_path: Path,
    table1: str,
    table2: str,
    rid1: int,
    ranked_columns: list[str],
    cols2: set[str],
    paired_columns: list[str],
    cols1_names: list[str],
    cols2_names: list[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    conn = open_readonly_db(db_path)
    try:
        row1 = fetch_row_dict(conn, table1, rid1, cols1_names)
        rid2, filters_used, status = find_match_rowid(
            conn, table2, row1, ranked_columns, cols2
        )

        entry: dict[str, Any] = {
            "excel1_rowid": rid1,
            "excel2_rowid": rid2,
            "status": status,
            "filters_used": [{"column": c, "value": v} for c, v in filters_used],
            "column_comparisons": [],
            "matching_columns": 0,
            "mismatching_columns": 0,
        }

        if rid2 is not None:
            row2 = fetch_row_dict(conn, table2, rid2, cols2_names)
            comparisons = compare_row_pair(row1, row2, paired_columns)
            entry["column_comparisons"] = comparisons
            entry["matching_columns"] = sum(1 for x in comparisons if x["match"])
            entry["mismatching_columns"] = sum(1 for x in comparisons if not x["match"])

        pair_text = (
            format_row_pair(rid1, rid2)
            if rid2 is not None and status == "matched"
            else ""
        )
        summary = {
            "id:id": pair_text,
            "excel1_rowid": rid1,
            "excel2_rowid": rid2 if rid2 is not None else "",
            "status": status,
            "filters_applied": len(filters_used),
            "columns_match": entry["matching_columns"],
            "columns_mismatch": entry["mismatching_columns"],
        }
        return entry, summary
    finally:
        conn.close()


def run_matching(
    db_path: Path,
    table1: str,
    table2: str,
    rowids: list[int],
    ranked: list[str],
    cols2: set[str],
    paired: list[str],
    cols1_names: list[str],
    cols2_names: list[str],
    workers: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    if workers <= 1 or len(rowids) <= 1:
        results: list[dict[str, Any]] = []
        summary_rows: list[dict[str, Any]] = []
        for rid1 in rowids:
            entry, summary = process_excel1_row(
                db_path,
                table1,
                table2,
                rid1,
                ranked,
                cols2,
                paired,
                cols1_names,
                cols2_names,
            )
            results.append(entry)
            summary_rows.append(summary)
        return results, summary_rows

    by_rid: dict[int, tuple[dict[str, Any], dict[str, Any]]] = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {
            pool.submit(
                process_excel1_row,
                db_path,
                table1,
                table2,
                rid1,
                ranked,
                cols2,
                paired,
                cols1_names,
                cols2_names,
            ): rid1
            for rid1 in rowids
        }
        for fut in as_completed(futures):
            rid1 = futures[fut]
            by_rid[rid1] = fut.result()

    results = [by_rid[rid][0] for rid in rowids]
    summary_rows = [by_rid[rid][1] for rid in rowids]
    return results, summary_rows


def format_row_pair(excel1_rowid: int, excel2_rowid: int) -> str:
    return f"{excel1_rowid}:{excel2_rowid}"


def build_wide_comparison(
    matched: list[dict[str, Any]],
    column_order: list[str],
) -> pd.DataFrame:
    """
    One row per matched pair.
    Id-excel1 | Id-excel2 | col | col | pass/fail | col2 | col2 | pass/fail | ...
    """
    headers: list[str] = ["Id-excel1", "Id-excel2"]
    for col in column_order:
        headers.extend([col, col, "pass/fail"])

    rows: list[list[Any]] = []
    for r in matched:
        by_col = {c["column"]: c for c in r["column_comparisons"]}
        row: list[Any] = [r["excel1_rowid"], r["excel2_rowid"]]
        for col in column_order:
            comp = by_col.get(col)
            if comp is None:
                row.extend(["", "", ""])
            else:
                row.extend(
                    [
                        comp["excel1_value"],
                        comp["excel2_value"],
                        "pass" if comp["match"] else "fail",
                    ]
                )
        rows.append(row)

    return pd.DataFrame(rows, columns=headers)


def pct(part: int, whole: int) -> float:
    if whole == 0:
        return 0.0
    return round(100.0 * part / whole, 2)


def build_statistics(
    results: list[dict[str, Any]],
    column_order: list[str],
) -> dict[str, Any]:
    total = len(results)
    matched_rows = [r for r in results if r["status"] == "matched"]
    matched_count = len(matched_rows)
    unmatched_count = total - matched_count
    ambiguous_count = sum(1 for r in results if str(r["status"]).startswith("ambiguous"))
    no_match_count = unmatched_count - ambiguous_count

    overall = {
        "total_rows_processed": total,
        "matched_rows": matched_count,
        "unmatched_rows": unmatched_count,
        "ambiguous_rows": ambiguous_count,
        "no_match_rows": no_match_count,
        "matched_percentage": pct(matched_count, total),
        "unmatched_percentage": pct(unmatched_count, total),
    }

    col_pass: dict[str, int] = {c: 0 for c in column_order}
    col_fail: dict[str, int] = {c: 0 for c in column_order}

    for r in matched_rows:
        by_name = {c["column"]: c for c in r["column_comparisons"]}
        for col in column_order:
            comp = by_name.get(col)
            if comp is None:
                continue
            if comp["match"]:
                col_pass[col] += 1
            else:
                col_fail[col] += 1

    column_stats: list[dict[str, Any]] = []
    total_cells_pass = 0
    total_cells_compared = 0
    for col in column_order:
        passed = col_pass[col]
        failed = col_fail[col]
        compared = passed + failed
        total_cells_pass += passed
        total_cells_compared += compared
        column_stats.append(
            {
                "column": col,
                "cells_compared": compared,
                "pass": passed,
                "fail": failed,
                "match_percentage": pct(passed, compared),
            }
        )

    overall["total_cells_compared"] = total_cells_compared
    overall["total_cells_pass"] = total_cells_pass
    overall["total_cells_fail"] = total_cells_compared - total_cells_pass
    overall["overall_cell_match_percentage"] = pct(total_cells_pass, total_cells_compared)

    return {"overall": overall, "by_column": column_stats}


def print_statistics(stats: dict[str, Any]) -> None:
    o = stats["overall"]
    print("--- Match statistics ---")
    print(f"Total rows processed: {o['total_rows_processed']}")
    print(f"Matched rows: {o['matched_rows']} ({o['matched_percentage']}%)")
    print(f"Unmatched rows: {o['unmatched_rows']} ({o['unmatched_percentage']}%)")
    print(f"  Ambiguous: {o['ambiguous_rows']}")
    print(f"  No match: {o['no_match_rows']}")
    print(
        f"Cell-level (matched pairs only): {o['total_cells_pass']}/{o['total_cells_compared']} "
        f"pass ({o['overall_cell_match_percentage']}%)"
    )
    print("\nColumn-wise match % (among matched row pairs):")
    for row in stats["by_column"]:
        print(
            f"  {row['column']}: {row['match_percentage']}% "
            f"({row['pass']} pass / {row['fail']} fail)"
        )


def export_to_excel(
    path: Path,
    results: list[dict[str, Any]],
    summary_rows: list[dict[str, Any]],
    column_order: list[str],
    stats: dict[str, Any],
) -> None:
    matched = [r for r in results if r["status"] == "matched" and r["excel2_rowid"] is not None]

    wide_df = build_wide_comparison(matched, column_order)
    if wide_df.empty and column_order:
        wide_df = build_wide_comparison([], column_order)

    overall_df = pd.DataFrame(
        [
            {"metric": "total_rows_processed", "value": stats["overall"]["total_rows_processed"]},
            {"metric": "matched_rows", "value": stats["overall"]["matched_rows"]},
            {"metric": "unmatched_rows", "value": stats["overall"]["unmatched_rows"]},
            {"metric": "ambiguous_rows", "value": stats["overall"]["ambiguous_rows"]},
            {"metric": "no_match_rows", "value": stats["overall"]["no_match_rows"]},
            {"metric": "matched_percentage", "value": stats["overall"]["matched_percentage"]},
            {"metric": "unmatched_percentage", "value": stats["overall"]["unmatched_percentage"]},
            {
                "metric": "overall_cell_match_percentage",
                "value": stats["overall"]["overall_cell_match_percentage"],
            },
            {"metric": "total_cells_compared", "value": stats["overall"]["total_cells_compared"]},
            {"metric": "total_cells_pass", "value": stats["overall"]["total_cells_pass"]},
            {"metric": "total_cells_fail", "value": stats["overall"]["total_cells_fail"]},
        ]
    )
    column_stats_df = pd.DataFrame(stats["by_column"])

    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        overall_df.to_excel(writer, sheet_name="statistics", index=False)
        column_stats_df.to_excel(writer, sheet_name="column_statistics", index=False)
        wide_df.to_excel(writer, sheet_name="comparison", index=False)
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="all_rows", index=False)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Match excel1 rows to excel2 and compare column values."
    )
    parser.add_argument("--db", type=Path, default=Path("excel_data.db"))
    parser.add_argument("--table1", default="excel1_input1")
    parser.add_argument("--table2", default="excel2_input2")
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="Process only first N rows from table1 (0 = all)",
    )
    parser.add_argument(
        "--row",
        type=int,
        metavar="ROWID",
        help="Process a single excel1 rowid only",
    )
    parser.add_argument("--json", type=Path, metavar="FILE", help="Write full JSON report")
    parser.add_argument(
        "--csv",
        type=Path,
        metavar="FILE",
        default=Path("match_summary.csv"),
        help="Summary CSV (default: match_summary.csv)",
    )
    parser.add_argument(
        "--excel",
        type=Path,
        metavar="FILE",
        default=Path("match_results.xlsx"),
        help="Excel export (default: match_results.xlsx)",
    )
    parser.add_argument(
        "--no-excel",
        action="store_true",
        help="Skip Excel export",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=default_worker_count(),
        help=f"Parallel threads for matching (default: {default_worker_count()}; use 1 to disable)",
    )
    args = parser.parse_args()
    if args.workers < 1:
        raise SystemExit("--workers must be >= 1")

    if not args.db.is_file():
        raise SystemExit(f"Database not found: {args.db.resolve()}")

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row
    try:
        cols1 = get_columns(conn, args.table1)
        cols2_list = get_columns(conn, args.table2)
        cols2 = set(cols2_list)
        paired = [c for c in cols1 if c in cols2]
        if not paired:
            raise SystemExit("No common column names between the two tables.")

        counts = distinct_counts(args.db, args.table1, paired, args.workers)
        ranked = sorted(paired, key=lambda c: (-counts[c], c))

        print(f"Table1: {args.table1}  |  Table2: {args.table2}")
        print(f"Workers: {args.workers}")
        print(f"Paired columns: {len(paired)}")
        print("Column priority (highest distinct count first):")
        for col in ranked[:10]:
            print(f"  {col}: {counts[col]} distinct")
        if len(ranked) > 10:
            print(f"  ... and {len(ranked) - 10} more")
        print()

        if args.row is not None:
            rowids = [args.row]
        else:
            sql = f"SELECT rowid FROM {quote_ident(args.table1)} ORDER BY rowid"
            if args.limit:
                sql += f" LIMIT {int(args.limit)}"
            rowids = [r[0] for r in conn.execute(sql)]

        results, summary_rows = run_matching(
            args.db,
            args.table1,
            args.table2,
            rowids,
            ranked,
            cols2,
            paired,
            cols1,
            cols2_list,
            args.workers,
        )

        stats = build_statistics(results, paired)
        print_statistics(stats)

        if results:
            sample = results[0]
            print(f"\nExample excel1 rowid={sample['excel1_rowid']} -> {sample['status']}")
            if sample["filters_used"]:
                print("  Filters used:")
                for f in sample["filters_used"]:
                    print(f"    {f['column']} = {f['value']!r}")
            if sample["column_comparisons"]:
                mism = [c for c in sample["column_comparisons"] if not c["match"]]
                print(f"  Column compare: {sample['matching_columns']} match, {sample['mismatching_columns']} mismatch")
                for c in mism[:8]:
                    print(f"    MISMATCH {c['column']}: {c['excel1_value']!r} vs {c['excel2_value']!r}")
                if len(mism) > 8:
                    print(f"    ... and {len(mism) - 8} more mismatches")

        if args.csv and summary_rows:
            args.csv.parent.mkdir(parents=True, exist_ok=True)
            with args.csv.open("w", newline="", encoding="utf-8") as fh:
                writer = csv.DictWriter(fh, fieldnames=list(summary_rows[0].keys()))
                writer.writeheader()
                writer.writerows(summary_rows)
            print(f"\nWrote summary CSV: {args.csv.resolve()}")

        if not args.no_excel and summary_rows:
            export_to_excel(args.excel, results, summary_rows, paired, stats)
            matched_pairs = sum(1 for r in results if r["status"] == "matched")
            print(
                f"Wrote Excel: {args.excel.resolve()} "
                f"({matched_pairs} rows on sheet 'comparison')"
            )

        if args.json:
            payload = {
                "table1": args.table1,
                "table2": args.table2,
                "column_priority": [{c: counts[c]} for c in ranked],
                "statistics": stats,
                "results": results,
            }
            args.json.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False, default=str),
                encoding="utf-8",
            )
            print(f"Wrote JSON: {args.json.resolve()}")

    finally:
        conn.close()


if __name__ == "__main__":
    main()
