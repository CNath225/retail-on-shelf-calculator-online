import argparse
import os
from pathlib import Path
import re
from typing import Optional

import duckdb
import pandas as pd

from master_data import (
    DEFAULT_MASTER_DB,
    active_account_name_map,
    active_sku_specs_from_master,
)


RAW_FILE = Path("raw_export.xlsx")
OUTPUT_DIR = Path(os.environ.get("ONSHELF_OUTPUT_ROOT", Path(__file__).parent / "step1_outputs"))


SKU_SPECS = [
    {"category": "Robot", "sku": "X60 Ultra", "raw_columns": ["X60 Ultra", "X60 Ultra "]},
    {
        "category": "Robot",
        "sku": "L50S Pro Ultra",
        "raw_columns": ["L50S Pro Ultra", "L50S Pro Ultra (JB)"],
    },
    {"category": "Robot", "sku": "L40 Ultra VE", "raw_columns": ["L40 Ultra VE", "L40 Ultra VE "]},
    {"category": "Robot", "sku": "L40 Plus", "raw_columns": ["L40 Plus", "L40 Plus "]},
    {"category": "Robot", "sku": "Matrix10 Ultra", "raw_columns": ["Matrix10 Ultra"]},
    {
        "category": "Robot",
        "sku": "Aqua10 Ultra Track S",
        "raw_columns": ["Aqua10 Ultra Track S", "Aqua10 Ultra Track S "],
    },
    {"category": "Robot", "sku": "Aqua10 Ultra Roller", "raw_columns": ["Aqua10 Ultra Roller"]},
    {"category": "Robot", "sku": "Aqua10 Roller AE", "raw_columns": ["Aqua10 Roller AE (JB)"]},
    {"category": "Robot", "sku": "Aqua10 Roller", "raw_columns": ["Aqua10 Roller"]},
    {"category": "Robot", "sku": "L50 Ultra", "raw_columns": ["L50 Ultra"]},
    {"category": "Stick", "sku": "Z20 Station", "raw_columns": ["Z20 Station"]},
    {"category": "Stick", "sku": "Z30 Station", "raw_columns": ["Z30 Station"]},
    {"category": "Stick", "sku": "Z50 Station", "raw_columns": ["Z50 Station", "Z50 Station (HN)"]},
    {"category": "Stick", "sku": "X1", "raw_columns": ["X1", "X1 (TGG)"]},
    {"category": "Stick", "sku": "X2", "raw_columns": ["X2", "X2 (JB)"]},
    {"category": "Stick", "sku": "X3 Station", "raw_columns": ["X3 Station", "X3 Station (TGG, JB)"]},
    {"category": "W&D", "sku": "H16 Pro Steam", "raw_columns": ["H16 Pro Steam"]},
    {"category": "W&D", "sku": "H15 Pro Heat", "raw_columns": ["H15 Pro Heat"]},
    {"category": "W&D", "sku": "T16 AE", "raw_columns": ["T16 AE", "T16 AE "]},
    {"category": "Beauty", "sku": "Dazzle Pro", "raw_columns": ["Dazzle Pro", "Dazzle Pro (JB)"]},
    {"category": "Beauty", "sku": "Aero Straight Pro", "raw_columns": ["Aero Straight Pro", "Aero Straight"]},
    {"category": "Beauty", "sku": "Gusto", "raw_columns": ["Gusto"]},
    {"category": "Beauty", "sku": "Airstyle Era", "raw_columns": ["Airstyle Era"]},
    {"category": "Beauty", "sku": "Pilot", "raw_columns": ["Pilot", "Pilot (HN)"]},
]


ACCOUNT_MAPPING = {
    "jb": "JB",
    "JbVIC": "JB",
    "AUDJ": "DJS",
}


ACCOUNT_NAME_TO_CODE = {
    "JB Hi-Fi": "JB",
    "Harvey Norman": "HN",
    "The Good Guys": "TGG",
    "Bing Lee": "BL",
    "David Jones": "DJS",
    "Betta Home Living Top 40": "BHLT",
}


def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip()) if not pd.isna(value) else ""


def get_account_from_place_id(place_id: object) -> str:
    if pd.isna(place_id):
        return ""

    match = re.match(r"^[A-Za-z]+", str(place_id).strip())
    if not match:
        return ""

    account = match.group(0)
    return ACCOUNT_MAPPING.get(account, account)


def get_account_code(row: pd.Series, account_name_to_code: dict[str, str]) -> str:
    account_name = normalize_text(row.get("Account Name", ""))
    if account_name in account_name_to_code:
        return account_name_to_code[account_name]

    account = normalize_text(row.get("Account", ""))
    if account:
        return ACCOUNT_MAPPING.get(account, account)

    return get_account_from_place_id(row.get("Place ID", ""))


def find_raw_column(raw_df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    normalized_columns = {normalize_text(column): column for column in raw_df.columns}
    for candidate in candidates:
        normalized_candidate = normalize_text(candidate)
        if normalized_candidate in normalized_columns:
            return normalized_columns[normalized_candidate]
    return None


def read_raw_submissions(
    raw_file: Path,
    month: str,
    account_name_to_code: dict[str, str],
) -> pd.DataFrame:
    raw_df = pd.read_excel(raw_file, sheet_name="Submissions")

    # Keep only real submissions. Excel sometimes has formatted empty rows.
    raw_df = raw_df[raw_df["ID"].notna()].copy()

    raw_df["Month Clean"] = month
    raw_df["Country Clean"] = raw_df["Country"] if "Country" in raw_df.columns else "AU"
    raw_df["Country Clean"] = raw_df["Country Clean"].fillna("AU")
    raw_df["Account Name Clean"] = (
        raw_df["Account Name"] if "Account Name" in raw_df.columns else ""
    )
    raw_df["Account Raw"] = raw_df["Account"] if "Account" in raw_df.columns else ""
    raw_df["Account Clean"] = raw_df.apply(
        lambda row: get_account_code(row, account_name_to_code), axis=1
    )

    return raw_df


def make_display_long(
    raw_df: pd.DataFrame,
    sku_specs: list[dict[str, object]],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    clean_parts = []
    quality_rows = []

    base_columns = [
        "Month Clean",
        "ID",
        "Place ID",
        "Place",
        "Address",
        "Representative",
        "Check-in type",
        "Country Clean",
        "Account Name Clean",
        "Account Raw",
        "Account Clean",
    ]

    for column in base_columns:
        if column not in raw_df.columns:
            raw_df[column] = ""

    for spec in sku_specs:
        raw_column = find_raw_column(raw_df, spec["raw_columns"])

        if raw_column is None:
            quality_rows.append(
                {
                    "issue_type": "missing_sku_column",
                    "category": spec["category"],
                    "sku": spec["sku"],
                    "detail": f"Could not find any of: {', '.join(spec['raw_columns'])}",
                }
            )
            continue

        category_long = raw_df[base_columns + [raw_column]].copy()
        category_long = category_long.rename(columns={raw_column: "Display Status"})
        category_long["Category"] = spec["category"]
        category_long["SKU"] = spec["sku"]
        category_long["Raw SKU Column"] = raw_column
        clean_parts.append(category_long)

    if not clean_parts:
        raise ValueError("No SKU columns were found. Please check the raw file format.")

    display_long = pd.concat(clean_parts, ignore_index=True)

    unknown_accounts = raw_df[raw_df["Account Clean"].eq("")]
    for _, row in unknown_accounts.iterrows():
        quality_rows.append(
            {
                "issue_type": "unknown_account",
                "category": "",
                "sku": "",
                "detail": f"Place ID={row.get('Place ID', '')}, Account Name={row.get('Account Name Clean', '')}",
            }
        )

    data_quality = pd.DataFrame(
        quality_rows, columns=["issue_type", "category", "sku", "detail"]
    )

    display_long["Is On Display Fixture"] = (
        display_long["Display Status"].astype(str).str.strip() == "On display - fixture"
    )

    display_long = display_long[
        [
            "Month Clean",
            "ID",
            "Country Clean",
            "Account Name Clean",
            "Account Raw",
            "Account Clean",
            "Category",
            "SKU",
            "Raw SKU Column",
            "Display Status",
            "Is On Display Fixture",
            "Place ID",
            "Place",
            "Address",
            "Representative",
            "Check-in type",
        ]
    ].rename(
        columns={
            "Month Clean": "month",
            "Country Clean": "country",
            "Account Name Clean": "account_name",
            "Account Raw": "account_raw",
            "Account Clean": "account",
            "Category": "category",
            "SKU": "sku",
            "Raw SKU Column": "raw_sku_column",
            "Display Status": "display_status",
            "Is On Display Fixture": "is_on_display_fixture",
            "Place ID": "place_id",
            "Place": "place",
            "Address": "address",
            "Representative": "representative",
            "Check-in type": "check_in_type",
        }
    )

    return display_long, data_quality


def save_outputs(
    raw_df: pd.DataFrame,
    display_long: pd.DataFrame,
    data_quality: pd.DataFrame,
    output_dir: Path,
) -> tuple[Path, Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    month = display_long["month"].iloc[0]
    csv_output = output_dir / f"display_long_{month}.csv"
    quality_output = output_dir / f"data_quality_{month}.csv"
    db_output = output_dir / f"retail_display_{month}.duckdb"

    display_long.to_csv(csv_output, index=False)
    data_quality.to_csv(quality_output, index=False)

    with duckdb.connect(str(db_output)) as con:
        con.register("raw_submissions_input", raw_df)
        con.register("display_long_input", display_long)
        con.register("data_quality_input", data_quality)
        con.execute("CREATE OR REPLACE TABLE raw_submissions AS SELECT * FROM raw_submissions_input")
        con.execute("CREATE OR REPLACE TABLE display_long AS SELECT * FROM display_long_input")
        con.execute("CREATE OR REPLACE TABLE data_quality AS SELECT * FROM data_quality_input")
        con.unregister("raw_submissions_input")
        con.unregister("display_long_input")
        con.unregister("data_quality_input")

    return csv_output, quality_output, db_output


def print_summary(
    raw_file: Path,
    raw_df: pd.DataFrame,
    display_long: pd.DataFrame,
    data_quality: pd.DataFrame,
    csv_output: Path,
    quality_output: Path,
    db_output: Path,
) -> None:
    summary = (
        display_long.groupby(["account", "category", "sku"], dropna=False)
        .agg(
            visited_rows=("ID", "count"),
            on_display_fixture=("is_on_display_fixture", "sum"),
        )
        .reset_index()
        .sort_values(["account", "category", "sku"])
    )

    print("Step 1: Prepare raw Resply data")
    print(f"Input file: {raw_file}")
    print(f"Month: {display_long['month'].iloc[0]}")
    print(f"Raw submission rows: {len(raw_df)}")
    print(f"Clean display rows: {len(display_long)}")
    print()
    print("Accounts found:")
    for account, count in raw_df["Account Clean"].value_counts(dropna=False).items():
        print(f"- {account}: {count} submissions")

    print()
    print("Account name to report account sample:")
    account_map_sample = (
        raw_df[["Account Name Clean", "Account Clean", "Place ID"]]
        .drop_duplicates()
        .sort_values(["Account Clean", "Account Name Clean"])
        .head(12)
    )
    print(account_map_sample.to_string(index=False))

    print()
    print("Sample summary, first 15 rows:")
    print(summary.head(15).to_string(index=False))

    print()
    print("Display status values:")
    print(display_long["display_status"].fillna("(blank)").value_counts().head(12).to_string())

    print()
    print(f"Data quality issues: {len(data_quality)}")
    if len(data_quality) > 0:
        print(data_quality.head(20).to_string(index=False))

    print()
    print(f"Saved CSV: {csv_output}")
    print(f"Saved data quality CSV: {quality_output}")
    print(f"Saved DuckDB database: {db_output}")
    print()
    print("Success. Step 1 is complete.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare Resply raw data for on-shelf reporting.")
    parser.add_argument("--month", default="2026-04", help="Report month, for example 2026-06.")
    parser.add_argument(
        "--raw-file",
        default=str(RAW_FILE),
        help="Path to Resply raw export xlsx file.",
    )
    parser.add_argument(
        "--output-dir",
        default=str(OUTPUT_DIR),
        help="Folder for generated CSV and DuckDB outputs.",
    )
    parser.add_argument(
        "--master-db",
        default=str(DEFAULT_MASTER_DB),
        help="DuckDB file containing editable SKU and Account master data.",
    )
    parser.add_argument(
        "--disable-master-data",
        action="store_true",
        help="Use the built-in SKU and Account mappings instead of editable master data.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    raw_file = Path(args.raw_file)
    output_dir = Path(args.output_dir) / args.month
    master_db = Path(args.master_db)

    if args.disable_master_data:
        sku_specs = SKU_SPECS
        account_name_to_code = ACCOUNT_NAME_TO_CODE
    else:
        try:
            sku_specs = active_sku_specs_from_master(master_db)
            account_name_to_code = active_account_name_map(master_db)
        except Exception as error:
            print(f"Warning: could not load master data, using built-in mappings. {error}")
            sku_specs = SKU_SPECS
            account_name_to_code = ACCOUNT_NAME_TO_CODE

    raw_df = read_raw_submissions(raw_file, args.month, account_name_to_code)
    display_long, data_quality = make_display_long(raw_df, sku_specs)
    csv_output, quality_output, db_output = save_outputs(
        raw_df, display_long, data_quality, output_dir
    )
    print_summary(
        raw_file,
        raw_df,
        display_long,
        data_quality,
        csv_output,
        quality_output,
        db_output,
    )


if __name__ == "__main__":
    main()
