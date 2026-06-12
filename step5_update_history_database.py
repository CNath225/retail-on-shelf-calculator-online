import argparse
from datetime import datetime, timezone
import os
from pathlib import Path

import duckdb
import pandas as pd


TOOL_DIR = Path(__file__).parent
DEFAULT_OUTPUT_ROOT = Path(os.environ.get("ONSHELF_OUTPUT_ROOT", TOOL_DIR / "step1_outputs"))
DEFAULT_HISTORY_DB = Path(os.environ.get("ONSHELF_HISTORY_DB", TOOL_DIR / "retail_on_shelf_history.duckdb"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update the persistent on-shelf history database.")
    parser.add_argument("--month", required=True, help="Report month, for example 2026-06.")
    parser.add_argument("--raw-file", required=True, help="Path to the monthly Resply raw export.")
    parser.add_argument("--range-file", required=True, help="Path to the range workbook used.")
    parser.add_argument("--range-sheet", required=True, help="Range sheet used.")
    parser.add_argument("--template-file", required=True, help="Path to the report template used.")
    parser.add_argument("--template-sheet", required=True, help="Template sheet used.")
    parser.add_argument("--month-label", required=True, help="Report month column used.")
    parser.add_argument(
        "--history-db",
        default=str(DEFAULT_HISTORY_DB),
        help="Persistent DuckDB history database path.",
    )
    return parser.parse_args()


def read_required_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise SystemExit(f"Required file is missing: {path}")
    return pd.read_csv(path)


def add_run_columns(df: pd.DataFrame, args: argparse.Namespace, table_kind: str) -> pd.DataFrame:
    result = df.copy()
    metadata = {
        "month": args.month,
        "table_kind": table_kind,
        "source_raw_file": str(Path(args.raw_file)),
        "source_range_file": str(Path(args.range_file)),
        "source_range_sheet": args.range_sheet,
        "source_template_file": str(Path(args.template_file)),
        "source_template_sheet": args.template_sheet,
        "report_month_label": args.month_label,
        "updated_at_utc": datetime.now(timezone.utc).isoformat(),
    }
    for position, (column, value) in enumerate(metadata.items()):
        if column in result.columns:
            result[column] = value
        else:
            result.insert(min(position, len(result.columns)), column, value)
    return result


def ensure_same_schema(con: duckdb.DuckDBPyConnection, table_name: str, df: pd.DataFrame) -> None:
    exists = con.execute(
        """
        SELECT COUNT(*)
        FROM information_schema.tables
        WHERE table_name = ?
        """,
        [table_name],
    ).fetchone()[0]

    if not exists:
        return

    existing_columns = [
        row[1]
        for row in con.execute(f"PRAGMA table_info('{table_name}')").fetchall()
    ]
    incoming_columns = list(df.columns)
    if existing_columns != incoming_columns:
        raise SystemExit(
            f"History table schema changed for {table_name}.\n"
            f"Existing columns: {existing_columns}\n"
            f"Incoming columns: {incoming_columns}\n"
            "Stop here instead of mixing different data shapes."
        )


def replace_month(con: duckdb.DuckDBPyConnection, table_name: str, df: pd.DataFrame, month: str) -> None:
    ensure_same_schema(con, table_name, df)
    view_name = f"new_{table_name}"
    con.register(view_name, df)

    exists = con.execute(
        """
        SELECT COUNT(*)
        FROM information_schema.tables
        WHERE table_name = ?
        """,
        [table_name],
    ).fetchone()[0]

    if not exists:
        con.execute(f"CREATE TABLE {table_name} AS SELECT * FROM {view_name} WHERE 1 = 0")

    con.execute(f"DELETE FROM {table_name} WHERE month = ?", [month])
    con.execute(f"INSERT INTO {table_name} SELECT * FROM {view_name}")
    con.unregister(view_name)


def main() -> None:
    args = parse_args()
    month_dir = DEFAULT_OUTPUT_ROOT / args.month

    display_long = read_required_csv(month_dir / f"display_long_{args.month}.csv")
    display_counts = read_required_csv(month_dir / f"display_count_summary_{args.month}.csv")
    rate_detail = read_required_csv(month_dir / f"key_sku_display_rate_{args.month}.csv")

    data_quality_file = month_dir / f"data_quality_{args.month}.csv"
    data_quality_issue_count = 0
    data_quality = pd.DataFrame(columns=["issue_type", "category", "sku", "detail"])
    if data_quality_file.exists():
        data_quality = pd.read_csv(data_quality_file)
        data_quality_issue_count = len(data_quality)

    duplicate_rate_keys = rate_detail.duplicated(
        subset=["country", "category", "account", "sku"], keep=False
    )
    if duplicate_rate_keys.any():
        duplicates = rate_detail.loc[
            duplicate_rate_keys, ["country", "category", "account", "sku"]
        ].head(20)
        raise SystemExit(
            "Duplicate rate keys found. Stop here instead of mixing rows:\n"
            + duplicates.to_string(index=False)
        )

    display_long = add_run_columns(display_long, args, "display_long")
    display_counts = add_run_columns(display_counts, args, "display_count_summary")
    rate_detail = add_run_columns(rate_detail, args, "key_sku_display_rate")
    data_quality = add_run_columns(data_quality, args, "data_quality")

    history_db = Path(args.history_db)
    history_db.parent.mkdir(parents=True, exist_ok=True)
    with duckdb.connect(str(history_db)) as con:
        replace_month(con, "display_long_history", display_long, args.month)
        replace_month(con, "display_count_history", display_counts, args.month)
        replace_month(con, "rate_history", rate_detail, args.month)
        replace_month(con, "data_quality_history", data_quality, args.month)

        run_row = pd.DataFrame(
            [
                {
                    "month": args.month,
                    "report_month_label": args.month_label,
                    "source_raw_file": str(Path(args.raw_file)),
                    "source_range_file": str(Path(args.range_file)),
                    "source_range_sheet": args.range_sheet,
                    "source_template_file": str(Path(args.template_file)),
                    "source_template_sheet": args.template_sheet,
                    "display_long_rows": len(display_long),
                    "display_count_rows": len(display_counts),
                    "rate_rows": len(rate_detail),
                    "data_quality_issues": data_quality_issue_count,
                    "updated_at_utc": datetime.now(timezone.utc).isoformat(),
                }
            ]
        )
        replace_month(con, "run_history", run_row, args.month)

    print("Step 5: Update history database")
    print(f"History database: {history_db}")
    print(f"Month replaced: {args.month}")
    print(f"display_long rows: {len(display_long)}")
    print(f"display_count rows: {len(display_counts)}")
    print(f"rate rows: {len(rate_detail)}")
    print(f"data quality issues saved as warnings: {data_quality_issue_count}")
    print()
    print("Success. Step 5 is complete.")


if __name__ == "__main__":
    main()
