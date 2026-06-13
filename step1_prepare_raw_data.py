import argparse
import os
from pathlib import Path
import re
from typing import Optional

import duckdb
import pandas as pd

from identifier_matching import (
    account_decision_map,
    apply_sku_decisions_to_specs,
    country_decision_map,
    load_alias_decisions,
    normalize_identifier,
)
from alias_workbook import load_alias_decisions_from_workbook
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
    {"category": "Robot", "sku": "Aqua10 Roller AE", "raw_columns": ["Aqua10 Roller AE", "Aqua10 Roller AE (JB)"]},
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


def normalize_match_key(value: object) -> str:
    return normalize_text(value).upper()


def strip_trailing_alias_suffix(value: object) -> str:
    text = normalize_text(value)
    suffix_pattern = re.compile(r"\s*(\([^()]*\)|\[[^\[\]]*\]|\{[^{}]*\})\s*$")
    while suffix_pattern.search(text):
        text = suffix_pattern.sub("", text).strip()
    return text


def raw_column_matches_candidate(column: object, candidate: object) -> bool:
    column_text = normalize_text(column)
    candidate_text = normalize_text(candidate)
    if not column_text or not candidate_text:
        return False

    column_key = normalize_match_key(column_text)
    candidate_key = normalize_match_key(candidate_text)
    if column_key == candidate_key:
        return True

    column_base_key = normalize_match_key(strip_trailing_alias_suffix(column_text))
    candidate_base_key = normalize_match_key(strip_trailing_alias_suffix(candidate_text))
    return column_base_key == candidate_key or column_base_key == candidate_base_key


def get_account_from_place_id(place_id: object) -> str:
    if pd.isna(place_id):
        return ""

    match = re.match(r"^[A-Za-z]+", str(place_id).strip())
    if not match:
        return ""

    account = match.group(0)
    return ACCOUNT_MAPPING.get(account, account)


def normalized_lookup(mapping: dict[str, str]) -> dict[str, str]:
    return {
        normalize_identifier(raw_value).key: canonical
        for raw_value, canonical in mapping.items()
        if normalize_identifier(raw_value).key
    }


def get_account_code(
    row: pd.Series,
    account_name_to_code: dict[str, str],
    smart_account_name_to_code: Optional[dict[str, str]] = None,
) -> str:
    account_name = normalize_text(row.get("Account Name", ""))
    if account_name in account_name_to_code:
        return account_name_to_code[account_name]
    if smart_account_name_to_code:
        account_name_key = normalize_identifier(account_name).key
        if account_name_key in smart_account_name_to_code:
            return smart_account_name_to_code[account_name_key]

    account = normalize_text(row.get("Account", ""))
    if account:
        if smart_account_name_to_code:
            account_key = normalize_identifier(account).key
            if account_key in smart_account_name_to_code:
                return smart_account_name_to_code[account_key]
        return ACCOUNT_MAPPING.get(account, account)

    place_account = get_account_from_place_id(row.get("Place ID", ""))
    if smart_account_name_to_code:
        place_account_key = normalize_identifier(place_account).key
        if place_account_key in smart_account_name_to_code:
            return smart_account_name_to_code[place_account_key]
    return place_account


def find_raw_column(raw_df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    normalized_columns = {normalize_match_key(column): column for column in raw_df.columns}
    for candidate in candidates:
        normalized_candidate = normalize_match_key(candidate)
        if normalized_candidate in normalized_columns:
            return normalized_columns[normalized_candidate]

    for candidate in candidates:
        for column in raw_df.columns:
            if raw_column_matches_candidate(column, candidate):
                return column
    return None


def find_raw_column_smart(raw_df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    candidate_keys = {
        normalize_identifier(candidate).key
        for candidate in candidates
        if normalize_identifier(candidate).key
    }
    matches = []
    for column in raw_df.columns:
        if normalize_identifier(column).key in candidate_keys:
            matches.append(column)
    if len(matches) > 1:
        raise ValueError(
            "Smart matching collision: multiple raw columns match the same canonical candidate: "
            + ", ".join(str(match) for match in matches)
        )
    return matches[0] if matches else None


def normalize_header(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def find_raw_submission_sheet(raw_file: Path) -> str:
    workbook = pd.ExcelFile(raw_file)
    sheet_names = workbook.sheet_names
    ordered_sheets = []
    if "Submissions" in sheet_names:
        ordered_sheets.append("Submissions")
    if "Submission" in sheet_names:
        ordered_sheets.append("Submission")
    ordered_sheets.extend(sheet for sheet in sheet_names if sheet not in ordered_sheets)

    for sheet_name in ordered_sheets:
        preview = pd.read_excel(raw_file, sheet_name=sheet_name, nrows=10)
        normalized_columns = {normalize_header(column) for column in preview.columns}
        has_id = "ID" in normalized_columns
        has_submission_date = "DATEANDTIME" in normalized_columns or "DATE" in normalized_columns
        has_place = "PLACEID" in normalized_columns or "PLACE" in normalized_columns
        normalized_sheet = normalize_header(sheet_name)
        likely_submission_sheet = normalized_sheet in {"SUBMISSION", "SUBMISSIONS"}
        if has_place and (has_id or has_submission_date or likely_submission_sheet):
            return sheet_name

    raise SystemExit(
        "Could not find a Repsly raw submissions sheet. "
        "Expected a sheet named Submissions/Submission, or any sheet with Place ID plus ID or Date and time columns. "
        f"Found sheets: {', '.join(sheet_names)}"
    )


def normalize_raw_submission_columns(raw_df: pd.DataFrame) -> pd.DataFrame:
    raw_df = raw_df.copy()

    column_aliases = {
        "ID": ["ID", "Submission ID"],
        "Place ID": ["Place ID", "PlaceID"],
        "Date and time": ["Date and time", "Date", "Submitted at"],
    }
    for target, candidates in column_aliases.items():
        if target in raw_df.columns:
            continue
        source = find_raw_column(raw_df, candidates)
        if source is not None:
            raw_df = raw_df.rename(columns={source: target})

    if "Place ID" not in raw_df.columns:
        raise ValueError("Raw file is missing Place ID, so submissions cannot be matched to stores.")

    if "ID" not in raw_df.columns:
        date_values = (
            raw_df["Date and time"].fillna("").tolist()
            if "Date and time" in raw_df.columns
            else [""] * len(raw_df)
        )
        raw_df["ID"] = [
            f"{place_id}|{date_value}|{row_number}"
            for row_number, (place_id, date_value) in enumerate(
                zip(raw_df["Place ID"].fillna(""), date_values),
                start=1,
            )
        ]

    return raw_df


def raw_month_counts(raw_df: pd.DataFrame) -> dict[str, int]:
    if "Date and time" not in raw_df.columns:
        return {}

    parsed_dates = pd.to_datetime(raw_df["Date and time"], errors="coerce", format="mixed")
    months = parsed_dates.dropna().dt.strftime("%Y-%m")
    if months.empty:
        return {}
    return {str(month): int(count) for month, count in months.value_counts().sort_index().items()}


def validate_raw_month(
    raw_df: pd.DataFrame,
    expected_month: str,
    allow_month_mismatch: bool,
) -> None:
    counts = raw_month_counts(raw_df)
    raw_df.attrs["raw_month_counts"] = counts
    if allow_month_mismatch or not counts:
        return

    dominant_month = max(counts, key=counts.get)
    if dominant_month != expected_month:
        counts_text = ", ".join(f"{month}: {count}" for month, count in counts.items())
        raise SystemExit(
            "Raw file date check failed. "
            f"Selected report month is {expected_month}, but raw submission dates are mostly {dominant_month}. "
            f"Month counts: {counts_text}. "
            "Choose the matching month or use the correct raw export."
        )


def read_raw_submissions(
    raw_file: Path,
    month: str,
    account_name_to_code: dict[str, str],
    allow_month_mismatch: bool = False,
    account_aliases: Optional[dict[str, str]] = None,
    country_aliases: Optional[dict[str, str]] = None,
) -> pd.DataFrame:
    raw_sheet = find_raw_submission_sheet(raw_file)
    raw_df = pd.read_excel(raw_file, sheet_name=raw_sheet)
    raw_df = normalize_raw_submission_columns(raw_df)
    raw_df.attrs["source_sheet"] = raw_sheet

    # Keep only real submissions. Excel sometimes has formatted empty rows.
    raw_df = raw_df[raw_df["Place ID"].notna()].copy()
    validate_raw_month(raw_df, month, allow_month_mismatch)

    raw_df["Month Clean"] = month
    raw_df["Country Clean"] = raw_df["Country"] if "Country" in raw_df.columns else "AU"
    raw_df["Country Clean"] = raw_df["Country Clean"].fillna("AU")
    if country_aliases:
        country_lookup = normalized_lookup(country_aliases)
        raw_df["Country Clean"] = raw_df["Country Clean"].apply(
            lambda value: country_lookup.get(normalize_identifier(value).key, value)
        )
    raw_df["Account Name Clean"] = (
        raw_df["Account Name"] if "Account Name" in raw_df.columns else ""
    )
    raw_df["Account Raw"] = raw_df["Account"] if "Account" in raw_df.columns else ""
    smart_account_lookup = normalized_lookup(account_aliases or {})
    raw_df["Account Clean"] = raw_df.apply(
        lambda row: get_account_code(row, account_name_to_code, smart_account_lookup), axis=1
    )

    return raw_df


def make_display_long(
    raw_df: pd.DataFrame,
    sku_specs: list[dict[str, object]],
    smart_matching: bool = False,
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
        raw_column_candidates = list(spec["raw_columns"])
        sku = spec.get("sku")
        if sku:
            raw_column_candidates.append(sku)
        raw_column = (
            find_raw_column_smart(raw_df, raw_column_candidates)
            if smart_matching
            else find_raw_column(raw_df, raw_column_candidates)
        )

        if raw_column is None:
            quality_rows.append(
                {
                    "issue_type": "missing_sku_column",
                    "category": spec["category"],
                    "sku": spec["sku"],
                    "detail": f"Could not find any of: {', '.join(raw_column_candidates)}",
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
    print(f"Input sheet: {raw_df.attrs.get('source_sheet', 'Submissions')}")
    print(f"Month: {display_long['month'].iloc[0]}")
    raw_months = raw_df.attrs.get("raw_month_counts", {})
    if raw_months:
        print("Raw date month counts:")
        for raw_month, count in raw_months.items():
            print(f"- {raw_month}: {count} submissions")
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
    parser.add_argument(
        "--allow-month-mismatch",
        action="store_true",
        help="Allow raw submission dates to differ from --month.",
    )
    parser.add_argument(
        "--enable-smart-matching",
        action="store_true",
        help="Enable beta identifier matching. Default off keeps the v1.0 path.",
    )
    parser.add_argument(
        "--identifier-alias-map",
        default="",
        help="Deprecated. Per-session JSON alias decisions for beta smart identifier matching.",
    )
    parser.add_argument(
        "--identifier-alias-workbook",
        default="",
        help="Workbook containing an embedded hidden-sheet alias map.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    raw_file = Path(args.raw_file)
    output_dir = Path(args.output_dir) / args.month
    master_db = Path(args.master_db)

    alias_decisions: list[dict[str, object]] = []
    if args.enable_smart_matching and args.identifier_alias_workbook:
        alias_decisions = load_alias_decisions_from_workbook(Path(args.identifier_alias_workbook))
    elif args.enable_smart_matching and args.identifier_alias_map:
        alias_decisions = load_alias_decisions(Path(args.identifier_alias_map))

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

    if args.enable_smart_matching:
        sku_specs = apply_sku_decisions_to_specs(sku_specs, alias_decisions)
        account_aliases = account_decision_map(alias_decisions)
        country_aliases = country_decision_map(alias_decisions)
    else:
        account_aliases = {}
        country_aliases = {}

    raw_df = read_raw_submissions(
        raw_file,
        args.month,
        account_name_to_code,
        allow_month_mismatch=args.allow_month_mismatch,
        account_aliases=account_aliases,
        country_aliases=country_aliases,
    )
    display_long, data_quality = make_display_long(
        raw_df,
        sku_specs,
        smart_matching=args.enable_smart_matching,
    )
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
