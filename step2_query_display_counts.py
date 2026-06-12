import argparse
import os
from pathlib import Path

import duckdb


DEFAULT_OUTPUT_ROOT = Path(os.environ.get("ONSHELF_OUTPUT_ROOT", Path(__file__).parent / "step1_outputs"))


SQL = """
SELECT
    account,
    account_name,
    category,
    sku,
    COUNT(*) AS visited_rows,
    COUNT(DISTINCT place_id) AS distinct_stores,
    SUM(CASE WHEN is_on_display_fixture THEN 1 ELSE 0 END) AS on_display_fixture
FROM display_long
GROUP BY account, account_name, category, sku
ORDER BY account, category, sku;
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Query display counts from DuckDB.")
    parser.add_argument("--month", default="2026-06", help="Report month, for example 2026-06.")
    parser.add_argument(
        "--db-file",
        default="",
        help="Optional DuckDB file path. If blank, uses the Step 1 output for --month.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    month_dir = DEFAULT_OUTPUT_ROOT / args.month
    db_file = Path(args.db_file) if args.db_file else month_dir / f"retail_display_{args.month}.duckdb"
    output_csv = month_dir / f"display_count_summary_{args.month}.csv"

    with duckdb.connect(str(db_file), read_only=True) as con:
        summary = con.execute(SQL).fetchdf()

    summary.to_csv(output_csv, index=False)

    print("Step 2: Query display counts with SQL")
    print(f"Database: {db_file}")
    print()
    print("SQL used:")
    print(SQL.strip())
    print()
    print("Summary rows:", len(summary))
    print()
    print("First 25 rows:")
    print(summary.head(25).to_string(index=False))
    print()
    print(f"Saved summary CSV: {output_csv}")
    print()
    print("Success. Step 2 is complete.")


if __name__ == "__main__":
    main()
