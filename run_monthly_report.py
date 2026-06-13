import argparse
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Optional


TOOL_DIR = Path(__file__).parent
DEFAULT_RANGE_FILE = TOOL_DIR / "range_template.xlsx"
DEFAULT_TEMPLATE_FILE = TOOL_DIR / "report_template.xlsx"
DEFAULT_MASTER_DB = TOOL_DIR / "retail_on_shelf_history.duckdb"
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
MONTH_NAME_TO_NUMBER = {
    "JAN": 1,
    "JANUARY": 1,
    "FEB": 2,
    "FEBRUARY": 2,
    "MAR": 3,
    "MARCH": 3,
    "APR": 4,
    "APRIL": 4,
    "MAY": 5,
    "JUN": 6,
    "JUNE": 6,
    "JUL": 7,
    "JULY": 7,
    "AUG": 8,
    "AUGUST": 8,
    "SEP": 9,
    "SEPT": 9,
    "SEPTEMBER": 9,
    "OCT": 10,
    "OCTOBER": 10,
    "NOV": 11,
    "NOVEMBER": 11,
    "DEC": 12,
    "DECEMBER": 12,
}


def run_step(command: list[str]) -> None:
    print()
    print("=" * 80)
    print("Running:")
    print(" ".join(command))
    print("=" * 80, flush=True)
    subprocess.run(command, check=True)


def infer_month_from_filename(filename: str) -> tuple[Optional[str], Optional[str]]:
    numeric_match = re.search(r"(20\d{2})[-_. /\\]?(0[1-9]|1[0-2])", filename)
    if numeric_match:
        year = numeric_match.group(1)
        month_number = int(numeric_match.group(2))
        return f"{year}-{month_number:02d}", MONTH_LABELS[month_number]

    month_names = "|".join(sorted(MONTH_NAME_TO_NUMBER, key=len, reverse=True))
    month_year_match = re.search(
        rf"\b({month_names})\b[-_. ]*(20\d{{2}})", filename, flags=re.IGNORECASE
    )
    year_month_match = re.search(
        rf"(20\d{{2}})[-_. ]*\b({month_names})\b", filename, flags=re.IGNORECASE
    )
    if month_year_match:
        month_name = month_year_match.group(1).upper()
        year = month_year_match.group(2)
    elif year_month_match:
        year = year_month_match.group(1)
        month_name = year_month_match.group(2).upper()
    else:
        return None, None

    month_number = MONTH_NAME_TO_NUMBER[month_name]
    return f"{year}-{month_number:02d}", MONTH_LABELS[month_number]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the full monthly on-shelf report flow.")
    parser.add_argument("--month", required=True, help="Report month, for example 2026-06.")
    parser.add_argument("--month-label", required=True, help="Report column label, for example JUN.")
    parser.add_argument(
        "--previous-month-label",
        required=True,
        help="Previous month column label in the template, for example May.",
    )
    parser.add_argument(
        "--raw-file",
        required=True,
        help="Path to the monthly Resply raw export xlsx file.",
    )
    parser.add_argument(
        "--range-file",
        default=str(DEFAULT_RANGE_FILE),
        help="Workbook containing TTL Store#/Range#/Range%%.",
    )
    parser.add_argument(
        "--range-sheet",
        default="Master data",
        help="Range table sheet name.",
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
        help="DuckDB file containing editable master data.",
    )
    parser.add_argument(
        "--disable-master-data",
        action="store_true",
        help="Use built-in SKU/Account mappings instead of editable master data in Step 1.",
    )
    parser.add_argument(
        "--template-file",
        default=str(DEFAULT_TEMPLATE_FILE),
        help="Workbook containing the final report layout.",
    )
    parser.add_argument(
        "--template-sheet",
        default="ANZ On-Shelf Retailer",
        help="Final report layout sheet name.",
    )
    parser.add_argument(
        "--allow-month-mismatch",
        action="store_true",
        help="Allow the raw file name month to differ from --month.",
    )
    parser.add_argument(
        "--keep-history-columns",
        action="store_true",
        help="Keep all month/history columns from the template in Report Preview.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    python = sys.executable
    inferred_month, inferred_label = infer_month_from_filename(Path(args.raw_file).name)
    if inferred_month and inferred_month != args.month and not args.allow_month_mismatch:
        raise SystemExit(
            f"Raw file looks like {inferred_month}, but --month is {args.month}. "
            "Fix the month or pass --allow-month-mismatch if this is intentional."
        )
    if inferred_label and not args.month_label.upper().startswith(inferred_label):
        print(
            f"Warning: raw file looks like {inferred_label}, "
            f"but report column is {args.month_label}."
        )

    step1_command = [
        python,
        str(TOOL_DIR / "step1_prepare_raw_data.py"),
        "--month",
        args.month,
        "--raw-file",
        args.raw_file,
        "--master-db",
        args.master_db,
    ]
    if args.disable_master_data:
        step1_command.append("--disable-master-data")
    if args.allow_month_mismatch:
        step1_command.append("--allow-month-mismatch")
    run_step(step1_command)

    run_step(
        [
            python,
            str(TOOL_DIR / "step2_query_display_counts.py"),
            "--month",
            args.month,
        ]
    )

    run_step(
        [
            python,
            str(TOOL_DIR / "step3_calculate_on_shelf_rate.py"),
            "--month",
            args.month,
            "--range-file",
            args.range_file,
            "--range-sheet",
            args.range_sheet,
            "--range-source",
            args.range_source,
            "--master-db",
            args.master_db,
        ]
    )

    step4_command = [
        python,
        str(TOOL_DIR / "step4_generate_report_preview.py"),
        "--month",
        args.month,
        "--month-label",
        args.month_label,
        "--previous-month-label",
        args.previous_month_label,
        "--template-file",
        args.template_file,
        "--template-sheet",
        args.template_sheet,
    ]
    if args.keep_history_columns:
        step4_command.append("--keep-history-columns")
    run_step(step4_command)

    run_step(
        [
            python,
            str(TOOL_DIR / "step5_update_history_database.py"),
            "--month",
            args.month,
            "--raw-file",
            args.raw_file,
            "--range-file",
            args.master_db if args.range_source == "master_db" else args.range_file,
            "--range-sheet",
            "range_master" if args.range_source == "master_db" else args.range_sheet,
            "--template-file",
            args.template_file,
            "--template-sheet",
            args.template_sheet,
            "--month-label",
            args.month_label,
            "--history-db",
            args.master_db,
        ]
    )

    output_root = Path(os.environ.get("ONSHELF_OUTPUT_ROOT", TOOL_DIR / "step1_outputs"))
    final_report = output_root / args.month / f"on_shelf_report_preview_{args.month}.xlsx"

    print()
    print("=" * 80)
    print("Monthly report flow complete.")
    print(f"Final report preview: {final_report}")
    print("=" * 80)


if __name__ == "__main__":
    main()
