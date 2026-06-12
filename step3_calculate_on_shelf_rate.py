import argparse
import os
from pathlib import Path

import duckdb
import pandas as pd

from master_data import DEFAULT_MASTER_DB, range_master_for_step3


TOOL_DIR = Path(__file__).parent
DEFAULT_OUTPUT_ROOT = Path(os.environ.get("ONSHELF_OUTPUT_ROOT", TOOL_DIR / "step1_outputs"))
DEFAULT_RANGE_FILE = TOOL_DIR / "range_template.xlsx"
DEFAULT_RANGE_SHEET = "Master data"


DISPLAY_SQL = """
SELECT
    country,
    account,
    category,
    sku,
    COUNT(*) AS visited_rows,
    COUNT(DISTINCT place_id) AS distinct_stores,
    SUM(CASE WHEN is_on_display_fixture THEN 1 ELSE 0 END) AS display_observations,
    COUNT(DISTINCT CASE WHEN is_on_display_fixture THEN place_id ELSE NULL END) AS display_stores
FROM display_long
GROUP BY country, account, category, sku;
"""


def normalize_key(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split()).upper()


def safe_rate(numerator: float, denominator: float, range_percent: float):
    if pd.isna(denominator) or denominator == 0:
        return None
    if pd.isna(range_percent) or range_percent == 0:
        return None
    return min(1.0, numerator / denominator / range_percent)


def rate_status(denominator: float, range_percent: float, range_count: object) -> str:
    if pd.isna(denominator) or denominator == 0:
        return "not visited"
    if pd.isna(range_percent) or range_percent == 0:
        return "range missing"
    if pd.isna(pd.to_numeric(pd.Series([range_count]), errors="coerce").iloc[0]):
        return "range missing"
    return "ok"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Calculate on-shelf rate using display counts and range table.")
    parser.add_argument("--month", default="2026-06", help="Report month, for example 2026-06.")
    parser.add_argument(
        "--db-file",
        default="",
        help="Optional DuckDB file path. If blank, uses the Step 1 output for --month.",
    )
    parser.add_argument(
        "--range-file",
        default=str(DEFAULT_RANGE_FILE),
        help="Workbook containing the Master data range table.",
    )
    parser.add_argument(
        "--range-sheet",
        default=DEFAULT_RANGE_SHEET,
        help="Sheet name containing Country/Category/SKU/Account/TTL Store#/Range#.",
    )
    parser.add_argument(
        "--range-source",
        choices=["excel", "master_db"],
        default="excel",
        help="Use an Excel range table or the editable master database.",
    )
    parser.add_argument(
        "--master-db",
        default=str(DEFAULT_MASTER_DB),
        help="DuckDB file containing editable range_master data.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    month_dir = DEFAULT_OUTPUT_ROOT / args.month
    db_file = Path(args.db_file) if args.db_file else month_dir / f"retail_display_{args.month}.duckdb"
    range_file = Path(args.range_file)
    output_csv = month_dir / f"key_sku_display_rate_{args.month}.csv"

    if args.range_source == "master_db":
        range_df = range_master_for_step3(Path(args.master_db))
    else:
        range_df = pd.read_excel(range_file, sheet_name=args.range_sheet)
        range_df = range_df.rename(
            columns={
                "Country": "country",
                "Category": "category",
                "SKU": "sku",
                "Account": "account",
                "TTL Store#": "ttl_store_count",
                "Range#": "range_store_count",
                "Range%": "range_percent",
            }
        )
        range_df = range_df[
            [
                "country",
                "category",
                "sku",
                "account",
                "ttl_store_count",
                "range_store_count",
                "range_percent",
            ]
        ].copy()

    range_df["ttl_store_count_numeric"] = pd.to_numeric(
        range_df["ttl_store_count"], errors="coerce"
    )
    range_df["range_store_count_numeric"] = pd.to_numeric(
        range_df["range_store_count"], errors="coerce"
    )
    range_df["range_percent_from_file"] = pd.to_numeric(
        range_df["range_percent"], errors="coerce"
    )
    range_df["range_percent_calculated"] = range_df.apply(
        lambda row: (
            row["range_store_count_numeric"] / row["ttl_store_count_numeric"]
            if pd.notna(row["ttl_store_count_numeric"])
            and row["ttl_store_count_numeric"] != 0
            and pd.notna(row["range_store_count_numeric"])
            else pd.NA
        ),
        axis=1,
    )
    range_df["range_percent"] = range_df["range_percent_calculated"].combine_first(
        range_df["range_percent_from_file"]
    )
    range_df["range_percent_source"] = range_df.apply(
        lambda row: "calculated_from_range_and_ttl"
        if pd.notna(row["range_percent_calculated"])
        else "from_file_or_missing",
        axis=1,
    )
    for column in ["country", "category", "sku", "account"]:
        range_df[f"{column}_key"] = range_df[column].apply(normalize_key)

    with duckdb.connect(str(db_file), read_only=True) as con:
        display_df = con.execute(DISPLAY_SQL).fetchdf()

    for column in ["country", "category", "sku", "account"]:
        display_df[f"{column}_key"] = display_df[column].apply(normalize_key)

    merged = range_df.merge(
        display_df,
        on=["country_key", "category_key", "sku_key", "account_key"],
        how="left",
        suffixes=("", "_display"),
    )

    for column in [
        "visited_rows",
        "distinct_stores",
        "display_observations",
        "display_stores",
    ]:
        merged[column] = merged[column].fillna(0)

    merged["visit_based_rate"] = merged.apply(
        lambda row: safe_rate(
            row["display_observations"], row["visited_rows"], row["range_percent"]
        ),
        axis=1,
    )
    merged["store_based_rate"] = merged.apply(
        lambda row: safe_rate(row["display_stores"], row["distinct_stores"], row["range_percent"]),
        axis=1,
    )
    merged["higher_rate"] = merged[["visit_based_rate", "store_based_rate"]].max(axis=1)
    # Final business metric: each field visit is a real shelf snapshot, so repeated
    # visits to the same store should count as repeated observations.
    merged["final_on_shelf_rate"] = merged["visit_based_rate"]
    merged["final_rate_basis"] = "visit_based"
    merged["rate_status"] = merged.apply(
        lambda row: rate_status(row["visited_rows"], row["range_percent"], row["range_store_count"]),
        axis=1,
    )

    result = merged[
        [
            "country",
            "category",
            "sku",
            "account",
            "ttl_store_count",
            "range_store_count",
            "range_percent",
            "range_percent_source",
            "visited_rows",
            "distinct_stores",
            "display_observations",
            "display_stores",
            "visit_based_rate",
            "store_based_rate",
            "higher_rate",
            "final_on_shelf_rate",
            "final_rate_basis",
            "rate_status",
        ]
    ].copy()

    result.to_csv(output_csv, index=False)

    print("Step 3: Calculate on-shelf rate")
    print(f"Database: {db_file}")
    if args.range_source == "master_db":
        print(f"Range table: {args.master_db} / range_master")
    else:
        print(f"Range table: {range_file} / {args.range_sheet}")
    print()
    print("Rate formulas:")
    print("visit_based_rate = display_observations / visited_rows / range_percent")
    print("store_based_rate = display_stores / distinct_stores / range_percent")
    print("higher_rate = MAX(visit_based_rate, store_based_rate)")
    print("final_on_shelf_rate = visit_based_rate")
    print()
    print("Rows:", len(result))
    print("Status counts:")
    print(result["rate_status"].value_counts(dropna=False).to_string())
    print()
    print("Sample rows where June data exists:")
    sample = result[result["visited_rows"] > 0].head(25)
    print(sample.to_string(index=False))
    print()
    print(f"Saved result CSV: {output_csv}")
    print()
    print("Success. Step 3 is complete.")


if __name__ == "__main__":
    main()
