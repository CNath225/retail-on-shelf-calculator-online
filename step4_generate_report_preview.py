import argparse
import os
from pathlib import Path
from typing import Any
import re

import pandas as pd
from xlsxwriter.utility import xl_col_to_name


TOOL_DIR = Path(__file__).parent
DEFAULT_OUTPUT_ROOT = Path(os.environ.get("ONSHELF_OUTPUT_ROOT", TOOL_DIR / "step1_outputs"))
DEFAULT_TEMPLATE_FILE = TOOL_DIR / "report_template.xlsx"
DEFAULT_TEMPLATE_SHEET = "ANZ On-Shelf Retailer"
MONTH_LABELS = {
    1: "JAN",
    2: "FEB",
    3: "MAR",
    4: "APR",
    5: "MAY",
    6: "JUN",
    7: "JUL",
    8: "AUG",
    9: "SEP",
    10: "OCT",
    11: "NOV",
    12: "DEC",
}


def normalize_key(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split()).upper()


def normalize_column_label(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def base_month_label(column: object) -> str:
    normalized = normalize_column_label(column)
    for label in MONTH_LABELS.values():
        if normalized.startswith(label):
            return label
    return ""


def is_month_column(column: object) -> bool:
    return bool(base_month_label(column))


def canonical_month_label(value: object) -> str:
    return base_month_label(value) or str(value).strip().upper()


def align_month_column(frame: pd.DataFrame, requested_label: str) -> tuple[pd.DataFrame, str]:
    label = canonical_month_label(requested_label)
    if label in frame.columns:
        return frame, label

    wanted = normalize_column_label(label)
    for column in frame.columns:
        if normalize_column_label(column) == wanted:
            return frame.rename(columns={column: label}), label

    return frame, label


def is_number(value: Any) -> bool:
    return pd.notna(pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0])


def as_number(value: Any):
    return pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]


def rate_to_report_value(row: pd.Series):
    if row["rate_status"] == "ok":
        return row["final_on_shelf_rate"]
    if row["rate_status"] == "range missing":
        return 0.0
    if row["rate_status"] == "not visited":
        return "not visited"
    return ""


def trend_value(current: Any, previous: Any) -> str:
    if not is_number(current) or not is_number(previous):
        return ""

    current_number = as_number(current)
    previous_number = as_number(previous)

    if current_number > previous_number:
        return "▲"
    if current_number < previous_number:
        return "▼"
    return "-"


def mostly_same_as_sku(frame: pd.DataFrame, column: str) -> bool:
    if "SKU" not in frame.columns:
        return False

    comparison = pd.DataFrame(
        {
            "left": frame[column].fillna("").astype(str).str.strip(),
            "sku": frame["SKU"].fillna("").astype(str).str.strip(),
        }
    )
    non_blank = comparison["left"].ne("")
    if not non_blank.any():
        return False
    return comparison.loc[non_blank, "left"].eq(comparison.loc[non_blank, "sku"]).mean() >= 0.95


def clean_report_columns(
    output_df: pd.DataFrame,
    month_label: str,
    previous_month_label: str,
    trend_column: str,
    keep_history_columns: bool,
) -> pd.DataFrame:
    cleaned = output_df.copy()

    rename_map = {}
    drop_columns = []
    for column in cleaned.columns:
        column_text = str(column)
        if not column_text.startswith("Unnamed:"):
            continue

        values = cleaned[column]
        non_blank_values = values.dropna().astype(str).str.strip()
        if len(non_blank_values) == 0:
            drop_columns.append(column)
        elif non_blank_values.eq("New").all():
            rename_map[column] = "New"
        elif mostly_same_as_sku(cleaned, column):
            drop_columns.append(column)
        else:
            rename_map[column] = "Note"

    cleaned = cleaned.drop(columns=drop_columns)
    cleaned = cleaned.rename(columns=rename_map)

    if keep_history_columns:
        return cleaned

    protected_columns = {
        "Country",
        "Category",
        "Channel",
        "SKU",
        previous_month_label,
        month_label,
        trend_column,
        "New",
        "Note",
    }
    selected_columns = []
    for column in cleaned.columns:
        if column in protected_columns:
            selected_columns.append(column)
        elif is_month_column(column):
            continue
        elif column.startswith("Unnamed:"):
            continue

    ordered = [
        column
        for column in [
            "Country",
            "Category",
            "Channel",
            "SKU",
            previous_month_label,
            month_label,
            trend_column,
            "New",
            "Note",
        ]
        if column in selected_columns
    ]
    extras = [column for column in selected_columns if column not in ordered]
    return cleaned[ordered + extras]


def is_blank(value: Any) -> bool:
    return pd.isna(value) or normalize_key(value) == ""


def excel_average_references(column_index: int, row_positions: list[int]) -> str:
    column_letter = xl_col_to_name(column_index)
    sorted_positions = sorted(row_positions)
    ranges = []
    range_start = sorted_positions[0]
    previous = sorted_positions[0]

    for position in sorted_positions[1:]:
        if position == previous + 1:
            previous = position
            continue

        if range_start == previous:
            ranges.append(f"{column_letter}{range_start + 2}")
        else:
            ranges.append(f"{column_letter}{range_start + 2}:{column_letter}{previous + 2}")
        range_start = position
        previous = position

    if range_start == previous:
        ranges.append(f"{column_letter}{range_start + 2}")
    else:
        ranges.append(f"{column_letter}{range_start + 2}:{column_letter}{previous + 2}")

    return ",".join(ranges)


def ttl_child_positions(frame: pd.DataFrame, ttl_position: int) -> list[int]:
    ttl_row = frame.iloc[ttl_position]
    ttl_country = normalize_key(ttl_row.get("Country", ""))
    ttl_category = normalize_key(ttl_row.get("Category", ""))
    ttl_channel = normalize_key(ttl_row.get("Channel", ""))

    positions = []
    for position, row in frame.iterrows():
        if position == ttl_position:
            continue

        same_country = normalize_key(row.get("Country", "")) == ttl_country
        same_category = normalize_key(row.get("Category", "")) == ttl_category
        if not same_country or not same_category:
            continue

        row_sku_is_ttl = normalize_key(row.get("SKU", "")) == "TTL"
        row_channel = normalize_key(row.get("Channel", ""))

        if ttl_channel:
            if row_channel == ttl_channel and not row_sku_is_ttl:
                positions.append(position)
        else:
            if row_channel and row_sku_is_ttl:
                positions.append(position)

    return positions


def write_ttl_average_formulas(
    worksheet,
    frame: pd.DataFrame,
    value_columns: list[str],
    percent_format,
) -> int:
    formula_count = 0
    ttl_positions = [
        position
        for position, row in frame.iterrows()
        if normalize_key(row.get("SKU", "")) == "TTL"
    ]

    for ttl_position in ttl_positions:
        child_positions = ttl_child_positions(frame, ttl_position)
        if not child_positions:
            continue

        for value_column in value_columns:
            if value_column not in frame.columns:
                continue

            column_index = frame.columns.get_loc(value_column)
            references = excel_average_references(column_index, child_positions)
            formula = f'=IFERROR(AVERAGE({references}),"")'
            cached_value = frame.at[ttl_position, value_column]
            if is_number(cached_value):
                cached_value = float(as_number(cached_value))
            else:
                cached_value = ""
            worksheet.write_formula(
                ttl_position + 1,
                column_index,
                formula,
                percent_format,
                cached_value,
            )
            formula_count += 1

    return formula_count


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate final on-shelf report preview.")
    parser.add_argument("--month", default="2026-06", help="Report month, for example 2026-06.")
    parser.add_argument(
        "--month-label",
        default="JUN",
        help="Column label to use in the final report, for example JUN.",
    )
    parser.add_argument(
        "--previous-month-label",
        default="May",
        help="Existing template column to compare trend against.",
    )
    parser.add_argument(
        "--template-file",
        default=str(DEFAULT_TEMPLATE_FILE),
        help="Workbook containing the final report layout.",
    )
    parser.add_argument(
        "--template-sheet",
        default=DEFAULT_TEMPLATE_SHEET,
        help="Sheet containing the final report layout.",
    )
    parser.add_argument(
        "--keep-history-columns",
        action="store_true",
        help="Keep all month/history columns from the template in Report Preview.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    month_dir = DEFAULT_OUTPUT_ROOT / args.month
    step3_file = month_dir / f"key_sku_display_rate_{args.month}.csv"
    output_xlsx = month_dir / f"on_shelf_report_preview_{args.month}.xlsx"

    report_df = pd.read_excel(args.template_file, sheet_name=args.template_sheet)
    rates_df = pd.read_csv(step3_file)

    rates_df["report_value"] = rates_df.apply(rate_to_report_value, axis=1)
    for column in ["country", "category", "sku", "account"]:
        rates_df[f"{column}_key"] = rates_df[column].apply(normalize_key)

    rate_lookup = {
        (
            row["country_key"],
            row["category_key"],
            row["account_key"],
            row["sku_key"],
        ): row["report_value"]
        for _, row in rates_df.iterrows()
    }

    output_df = report_df.copy()
    output_df, month_label = align_month_column(output_df, args.month_label)
    output_df, previous_month_label = align_month_column(output_df, args.previous_month_label)

    if month_label in output_df.columns:
        output_df[month_label] = ""
    else:
        insert_at = (
            output_df.columns.get_loc("Trend")
            if "Trend" in output_df.columns
            else len(output_df.columns)
        )
        output_df.insert(insert_at, month_label, "")

    for index, row in output_df.iterrows():
        sku = row.get("SKU", "")
        if normalize_key(sku) == "TTL":
            continue

        lookup_key = (
            normalize_key(row.get("Country", "")),
            normalize_key(row.get("Category", "")),
            normalize_key(row.get("Channel", "")),
            normalize_key(sku),
        )
        output_df.at[index, month_label] = rate_lookup.get(lookup_key, "")

    # Calculate TTL rows as the average of numeric rows in the same Country + Category + Channel block.
    group_columns = ["Country", "Category", "Channel"]
    for group_key, group in output_df.groupby(group_columns, dropna=False):
        ttl_index = group[normalize_key_series(group["SKU"]).eq("TTL")].index
        if len(ttl_index) == 0:
            continue

        detail_values = output_df.loc[
            group.index.difference(ttl_index), month_label
        ].apply(lambda value: as_number(value) if is_number(value) else pd.NA)
        numeric_values = detail_values.dropna()

        if len(numeric_values) > 0:
            output_df.loc[ttl_index, month_label] = float(numeric_values.mean())
        else:
            output_df.loc[ttl_index, month_label] = ""

    # Calculate category-level TTL rows where Channel is blank.
    category_ttl_mask = (
        output_df["SKU"].apply(normalize_key).eq("TTL")
        & output_df["Channel"].isna()
    )
    for ttl_index, ttl_row in output_df[category_ttl_mask].iterrows():
        child_ttl_mask = (
            output_df["Country"].apply(normalize_key).eq(normalize_key(ttl_row["Country"]))
            & output_df["Category"].apply(normalize_key).eq(normalize_key(ttl_row["Category"]))
            & output_df["SKU"].apply(normalize_key).eq("TTL")
            & output_df["Channel"].notna()
        )
        child_values = output_df.loc[child_ttl_mask, month_label].apply(
            lambda value: as_number(value) if is_number(value) else pd.NA
        )
        numeric_child_values = child_values.dropna()
        if len(numeric_child_values) > 0:
            output_df.at[ttl_index, month_label] = float(numeric_child_values.mean())

    trend_column = "Trend" if "Trend" in output_df.columns else f"Trend vs {previous_month_label}"
    output_df[trend_column] = output_df.apply(
        lambda row: trend_value(row[month_label], row.get(previous_month_label, "")),
        axis=1,
    )
    output_df = clean_report_columns(
        output_df=output_df,
        month_label=month_label,
        previous_month_label=previous_month_label,
        trend_column=trend_column,
        keep_history_columns=args.keep_history_columns,
    ).reset_index(drop=True)

    key_sku_columns = [
        "country",
        "category",
        "sku",
        "account",
        "ttl_store_count",
        "range_store_count",
        "range_percent",
        "range_percent_source",
        "visited_rows",
        "display_observations",
        "final_on_shelf_rate",
        "rate_status",
    ]
    key_sku_df = rates_df[key_sku_columns].copy()

    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        output_df.to_excel(writer, sheet_name="Report Preview", index=False)
        key_sku_df.to_excel(writer, sheet_name="Key SKU Display", index=False)

        workbook = writer.book
        percent_format = workbook.add_format({"num_format": "0%"})
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7"})
        status_format = workbook.add_format({"bg_color": "#FCE4D6"})

        for sheet_name, df in {
            "Report Preview": output_df,
            "Key SKU Display": key_sku_df,
        }.items():
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            for col_idx, column_name in enumerate(df.columns):
                width = min(max(len(str(column_name)) + 2, 12), 28)
                worksheet.set_column(col_idx, col_idx, width)
                worksheet.write(0, col_idx, column_name, header_format)

        report_sheet = writer.sheets["Report Preview"]
        month_col = output_df.columns.get_loc(month_label)
        report_sheet.set_column(month_col, month_col, 12, percent_format)

        if previous_month_label in output_df.columns:
            previous_col = output_df.columns.get_loc(previous_month_label)
            report_sheet.set_column(previous_col, previous_col, 12, percent_format)

        formula_columns = [
            column
            for column in [previous_month_label, month_label]
            if column in output_df.columns
        ]
        ttl_formula_count = write_ttl_average_formulas(
            worksheet=report_sheet,
            frame=output_df,
            value_columns=formula_columns,
            percent_format=percent_format,
        )

        key_sheet = writer.sheets["Key SKU Display"]
        for col_name in ["range_percent", "final_on_shelf_rate"]:
            col_idx = key_sku_df.columns.get_loc(col_name)
            key_sheet.set_column(col_idx, col_idx, 14, percent_format)

        status_col = key_sku_df.columns.get_loc("rate_status")
        key_sheet.conditional_format(
            1,
            status_col,
            len(key_sku_df),
            status_col,
            {
                "type": "text",
                "criteria": "containing",
                "value": "range missing",
                "format": status_format,
            },
        )

    matched = output_df[month_label].ne("").sum()
    ttl_rows = output_df["SKU"].apply(normalize_key).eq("TTL").sum()
    template_countries = sorted(
        output_df["Country"].dropna().astype(str).str.strip().unique().tolist()
    )
    rate_countries = sorted(
        rates_df["country"].dropna().astype(str).str.strip().unique().tolist()
    )
    missing_template_countries = [
        country for country in rate_countries if country not in template_countries
    ]

    print("Step 4: Generate report preview")
    print(f"Template: {args.template_file} / {args.template_sheet}")
    print(f"Step 3 result: {step3_file}")
    print(f"Report rows: {len(output_df)}")
    print(f"Rows filled for {month_label}: {matched}")
    print(f"TTL rows recalculated: {ttl_rows}")
    print(f"TTL Excel formulas written: {ttl_formula_count}")
    print(f"Template countries: {template_countries}")
    print(f"Rate table countries: {rate_countries}")
    print(f"History columns kept: {args.keep_history_columns}")
    if missing_template_countries:
        print(f"Warning: countries in rate table but not in template: {missing_template_countries}")
    print()
    print("First 20 report rows:")
    preview_columns = [
        column
        for column in ["Country", "Category", "Channel", "SKU", previous_month_label, month_label, trend_column]
        if column in output_df.columns
    ]
    print(output_df[preview_columns].head(20).to_string(index=False))
    print()
    print(f"Saved report preview: {output_xlsx}")
    print()
    print("Success. Step 4 is complete.")


def normalize_key_series(series: pd.Series) -> pd.Series:
    return series.apply(normalize_key)


if __name__ == "__main__":
    main()
