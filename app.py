from datetime import datetime, timezone
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Optional
import uuid

import pandas as pd
import streamlit as st


TOOL_DIR = Path(__file__).parent
RUNTIME_DIR = TOOL_DIR / ".runtime"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_ROOT = RUNTIME_DIR / "outputs"
APP_TITLE = "Retail On-shelf Rate Calculator Online by CodeNATHAN"

RANGE_SHEET_PREFERENCES = ["Master data", "Key SKU Display% (2)", "Key SKU Display%"]
TEMPLATE_SHEET_PREFERENCES = ["ANZ On-Shelf Retailer", "ANZ On-Shelf Retailer (2)", "数据不含公式版"]
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
ALL_MONTH_LABELS = list(MONTH_LABELS.values())
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


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip()
    return cleaned or "upload.xlsx"


def save_upload(uploaded_file, target_dir: Path) -> Optional[Path]:
    if uploaded_file is None:
        return None
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = target_dir / safe_filename(uploaded_file.name)
    target_path.write_bytes(uploaded_file.getbuffer())
    return target_path


def xlsx_sheets(path: Optional[Path]) -> list[str]:
    if not path or not path.exists():
        return []
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception:
        return []


def preferred_range_sheet(sheets: list[str]) -> str:
    for sheet in RANGE_SHEET_PREFERENCES:
        if sheet in sheets:
            return sheet
    return sheets[0] if sheets else "Master data"


def normalize_headers(columns: list[object]) -> list[str]:
    return [normalize_column_label(column) for column in columns]


def sheet_columns(path: Optional[Path], sheet: str) -> list[object]:
    if not path or not path.exists() or not sheet:
        return []
    try:
        return list(pd.read_excel(path, sheet_name=sheet, nrows=0).columns)
    except Exception:
        return []


def first_sheet_matching(path: Optional[Path], sheets: list[str], kind: str) -> str:
    if not path or not path.exists():
        return sheets[0] if sheets else ""

    if kind == "range":
        for preferred in RANGE_SHEET_PREFERENCES:
            if preferred in sheets:
                return preferred
    if kind == "template":
        for preferred in TEMPLATE_SHEET_PREFERENCES:
            if preferred in sheets:
                return preferred

    best_sheet = sheets[0] if sheets else ""
    best_score = -1
    for sheet in sheets:
        columns = sheet_columns(path, sheet)
        score = sheet_score(columns, kind)
        if score > best_score:
            best_score = score
            best_sheet = sheet
    return best_sheet


def sheet_score(columns: list[object], kind: str) -> int:
    normalized = normalize_headers(columns)
    normalized_set = set(normalized)

    if kind == "raw":
        score = 0
        score += 4 if "ID" in normalized_set else 0
        score += 3 if "PLACEID" in normalized_set else 0
        score += 2 if "PLACE" in normalized_set else 0
        score += 2 if "ACCOUNTNAME" in normalized_set else 0
        return score

    if kind == "range":
        score = 0
        for required in ["COUNTRY", "CATEGORY", "SKU", "ACCOUNT"]:
            score += 2 if required in normalized_set else 0
        score += 3 if "TTLSTORE" in normalized_set else 0
        score += 2 * normalized.count("RANGE")
        return score

    if kind == "template":
        score = 0
        for required in ["COUNTRY", "CATEGORY", "CHANNEL", "SKU"]:
            score += 3 if required in normalized_set else 0
        score += sum(1 for column in columns if is_month_column(column))
        return score

    return 0


def workbook_looks_like(path: Optional[Path], kind: str, sheet: Optional[str] = None) -> bool:
    sheets = xlsx_sheets(path)
    if not sheets:
        return False
    sheets_to_check = [sheet] if sheet else sheets
    threshold = {"raw": 7, "range": 11, "template": 12}[kind]
    return any(sheet_score(sheet_columns(path, one_sheet), kind) >= threshold for one_sheet in sheets_to_check)


def month_label_from_month(month: str) -> str:
    match = re.match(r"^20\d{2}-(0[1-9]|1[0-2])$", month)
    if not match:
        return st.session_state.get("month_label", "JUN")
    return MONTH_LABELS[int(match.group(1))]


def month_options() -> list[str]:
    current_year = datetime.now().year
    return [
        f"{year}-{month_number:02d}"
        for year in range(current_year - 1, current_year + 3)
        for month_number in range(1, 13)
    ]


def base_month_label(column: object) -> str:
    normalized = normalize_column_label(column)
    for label in MONTH_LABELS.values():
        if normalized.startswith(label):
            return label
    return ""


def is_month_column(column: object) -> bool:
    return bool(base_month_label(column))


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
        month_only_match = re.search(rf"\b({month_names})\b", filename, flags=re.IGNORECASE)
        if not month_only_match:
            return None, None
        month_name = month_only_match.group(1).upper()
        year = str(datetime.now().year)

    month_number = MONTH_NAME_TO_NUMBER[month_name]
    return f"{year}-{month_number:02d}", MONTH_LABELS[month_number]


def previous_label(month: str) -> str:
    match = re.match(r"^(20\d{2})-(0[1-9]|1[0-2])$", month)
    if not match:
        return st.session_state.get("previous_month_label", "MAY")
    month_number = int(match.group(2))
    previous_number = 12 if month_number == 1 else month_number - 1
    return MONTH_LABELS[previous_number]


def normalize_column_label(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def choose_report_column(month_label: Optional[str], template_columns: list[str]) -> Optional[str]:
    if not month_label:
        return None
    wanted = normalize_column_label(month_label)
    for column in template_columns:
        if normalize_column_label(column) == wanted:
            return base_month_label(column) or str(column).upper()
    return month_label


def month_select_index(value: str, fallback: str) -> int:
    normalized = base_month_label(value) or base_month_label(fallback) or "JUN"
    return ALL_MONTH_LABELS.index(normalized) if normalized in ALL_MONTH_LABELS else 5


def render_app_title() -> None:
    st.markdown(
        """
        <style>
          .app-title {
            margin: 0.25rem 0 1.5rem;
            font-size: clamp(2.1rem, 4vw, 3.5rem);
            font-weight: 800;
            line-height: 1.08;
            color: #f4f5f7;
          }
          .code-nathan {
            display: inline-block;
            color: transparent;
            background: linear-gradient(100deg, #a98235 0%, #7c682f 35%, #b7bbc2 65%, #24272d 100%);
            -webkit-background-clip: text;
            background-clip: text;
            text-shadow: 0 0 12px rgba(169, 130, 53, 0.35), 0 0 2px rgba(255, 255, 255, 0.25);
            animation: codeJump 1.25s steps(2, end) infinite;
          }
          @keyframes codeJump {
            0%, 100% { transform: translateY(0); filter: brightness(1); }
            28% { transform: translateY(-1px); filter: brightness(1.25); }
            32% { transform: translateY(1px); filter: brightness(0.9); }
            48% { transform: translateY(0); filter: brightness(1.1); }
            52% { transform: translateY(-2px); filter: brightness(1.35); }
          }
          @media (prefers-reduced-motion: reduce) {
            .code-nathan { animation: none; }
          }
        </style>
        <div class="app-title">Retail On-shelf Rate Calculator Online by <span class="code-nathan">CodeNATHAN</span></div>
        """,
        unsafe_allow_html=True,
    )


def run_command(command: list[str], output_root: Path, history_db: Path) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["ONSHELF_OUTPUT_ROOT"] = str(output_root)
    env["ONSHELF_HISTORY_DB"] = str(history_db)
    env["PYTHONPYCACHEPREFIX"] = str(output_root / "_pycache")
    return subprocess.run(command, capture_output=True, text=True, env=env)


def summarize_raw_upload(path: Optional[Path]) -> dict[str, object]:
    if not path or not path.exists():
        return {}

    try:
        from step1_prepare_raw_data import (
            ACCOUNT_NAME_TO_CODE,
            find_raw_submission_sheet,
            get_account_code,
            normalize_raw_submission_columns,
            raw_month_counts,
        )

        raw_sheet = find_raw_submission_sheet(path)
        raw_df = pd.read_excel(path, sheet_name=raw_sheet)
        raw_df = normalize_raw_submission_columns(raw_df)
        raw_df = raw_df[raw_df["Place ID"].notna()].copy()
        account_counts = (
            raw_df.apply(lambda row: get_account_code(row, ACCOUNT_NAME_TO_CODE), axis=1)
            .replace("", "(unknown)")
            .value_counts(dropna=False)
            .sort_index()
        )
        month_counts = raw_month_counts(raw_df)
        dominant_month = max(month_counts, key=month_counts.get) if month_counts else ""
        return {
            "sheet": raw_sheet,
            "rows": int(len(raw_df)),
            "month_counts": month_counts,
            "dominant_month": dominant_month,
            "account_counts": {str(key): int(value) for key, value in account_counts.items()},
        }
    except SystemExit as error:
        return {"error": str(error)}
    except Exception as error:
        return {"error": f"Could not inspect raw file: {error}"}


def file_signature(path: Optional[Path]) -> Optional[dict[str, object]]:
    if not path or not path.exists():
        return None
    stat = path.stat()
    return {
        "name": path.name,
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def input_signature(
    month: str,
    month_label: str,
    previous_month_label: str,
    raw_path: Optional[Path],
    range_path: Optional[Path],
    template_path: Optional[Path],
    range_sheet: str,
    template_sheet: str,
    keep_history_columns: bool,
) -> dict[str, object]:
    return {
        "month": month,
        "month_label": month_label,
        "previous_month_label": previous_month_label,
        "raw": file_signature(raw_path),
        "range": file_signature(range_path),
        "template": file_signature(template_path),
        "range_sheet": range_sheet,
        "template_sheet": template_sheet,
        "keep_history_columns": keep_history_columns,
    }


def render_report(final_report: Path, step3_csv: Path, count_csv: Path) -> None:
    st.download_button(
        "Download Report",
        data=final_report.read_bytes(),
        file_name=final_report.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click="ignore",
    )
    preview_tab, detail_tab, count_tab = st.tabs(["Preview", "Rate Detail", "Display Counts"])
    with preview_tab:
        st.dataframe(pd.read_excel(final_report, sheet_name="Report Preview"), use_container_width=True, height=520)
    with detail_tab:
        if step3_csv.exists():
            st.dataframe(pd.read_csv(step3_csv), use_container_width=True, height=520)
    with count_tab:
        if count_csv.exists():
            st.dataframe(pd.read_csv(count_csv), use_container_width=True, height=520)


def init_state() -> None:
    st.session_state.setdefault("session_id", uuid.uuid4().hex)
    st.session_state.setdefault("month", "2026-06")
    st.session_state.setdefault("month_label", "JUN")
    st.session_state.setdefault("previous_month_label", "MAY")
    st.session_state.setdefault("_last_raw_filename", "")
    st.session_state.setdefault("last_report_state", None)
    st.session_state.setdefault("last_run_log", "")


def clear_session_files() -> None:
    session_id = st.session_state["session_id"]
    for path in [UPLOAD_DIR / session_id, OUTPUT_ROOT / session_id]:
        if path.exists():
            shutil.rmtree(path)
    st.session_state["last_report_state"] = None
    st.session_state["last_run_log"] = ""


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    render_app_title()

    with st.sidebar:
        raw_file = st.file_uploader("Resply Raw Export", type=["xlsx"])
        range_file = st.file_uploader("Range Table", type=["xlsx"])
        template_file = st.file_uploader("Report Template", type=["xlsx"])

        if raw_file is not None:
            inferred_month, inferred_label = infer_month_from_filename(raw_file.name)
            if (
                inferred_month
                and inferred_label
                and st.session_state["_last_raw_filename"] != raw_file.name
            ):
                st.session_state["month"] = inferred_month
                st.session_state["month_label"] = inferred_label
                st.session_state["previous_month_label"] = previous_label(inferred_month)
                st.session_state["_last_raw_filename"] = raw_file.name

    session_upload_dir = UPLOAD_DIR / st.session_state["session_id"]
    raw_path = save_upload(raw_file, session_upload_dir) if raw_file else None
    range_path = save_upload(range_file, session_upload_dir) if range_file else None
    template_path = save_upload(template_file, session_upload_dir) if template_file else None

    range_sheets = xlsx_sheets(range_path)
    template_sheets = xlsx_sheets(template_path)
    template_columns = []
    range_default_sheet = first_sheet_matching(range_path, range_sheets, "range")
    template_default_sheet = first_sheet_matching(template_path, template_sheets, "template")

    with st.sidebar:
        range_sheet = st.selectbox(
            "Range Sheet",
            range_sheets or ["Master data"],
            index=(
                range_sheets.index(range_default_sheet)
                if range_default_sheet in range_sheets
                else 0
            ),
        )
        template_sheet = st.selectbox(
            "Template Sheet",
            template_sheets or ["ANZ On-Shelf Retailer"],
            index=(
                template_sheets.index(template_default_sheet)
                if template_default_sheet in template_sheets
                else 0
            ),
        )

        if template_path and template_path.exists():
            try:
                template_columns = list(pd.read_excel(template_path, sheet_name=template_sheet, nrows=0).columns)
            except Exception:
                template_columns = []

        month_values = month_options()
        if st.session_state["month"] not in month_values:
            month_values.insert(0, st.session_state["month"])
        selected_month = st.selectbox(
            "Month",
            month_values,
            index=month_values.index(st.session_state["month"]),
        )
        if selected_month != st.session_state["month"]:
            st.session_state["month"] = selected_month
            st.session_state["month_label"] = month_label_from_month(selected_month)
            st.session_state["previous_month_label"] = previous_label(selected_month)

        month = selected_month
        month_label = st.selectbox(
            "Report Column",
            ALL_MONTH_LABELS,
            index=month_select_index(st.session_state["month_label"], month_label_from_month(month)),
        )
        st.session_state["month_label"] = month_label

        default_compare = previous_label(month)
        compare_options = ALL_MONTH_LABELS.copy()
        previous_state_label = base_month_label(st.session_state["previous_month_label"]) or default_compare
        if previous_state_label not in compare_options:
            st.session_state["previous_month_label"] = default_compare
        previous_month_label = st.selectbox(
            "Compare To",
            compare_options,
            index=month_select_index(previous_state_label, default_compare),
        )
        st.session_state["previous_month_label"] = previous_month_label
        keep_history_columns = st.checkbox("Keep history columns", value=False)
        st.button("Clear Session Files", on_click=clear_session_files)

    validation_errors = []
    raw_summary = summarize_raw_upload(raw_path) if raw_path else {}
    if raw_path and not workbook_looks_like(raw_path, "raw"):
        validation_errors.append(
            "Raw Export file looks wrong. Upload the Repsly raw export here; it should contain Place ID plus ID or Date and time columns."
        )
    if raw_summary.get("error"):
        validation_errors.append(str(raw_summary["error"]))
    elif raw_summary.get("dominant_month") and raw_summary["dominant_month"] != month:
        validation_errors.append(
            "Raw month check failed. "
            f"The selected report month is {month}, but the raw file dates are mostly {raw_summary['dominant_month']}. "
            f"Raw month counts: {raw_summary.get('month_counts', {})}. "
            "Please select the matching month or upload the correct raw export."
        )
    if range_path and not workbook_looks_like(range_path, "range", range_sheet):
        validation_errors.append(
            "Range Table file/sheet looks wrong. Upload the workbook sheet with Country, Category, SKU, Account, TTL Store#, and Range#."
        )
    if template_path and not workbook_looks_like(template_path, "template", template_sheet):
        validation_errors.append(
            "Report Template file/sheet looks wrong. Upload the final report template sheet with Country, Category, Channel, and SKU."
        )

    st.subheader("Inputs")
    st.write(
        {
            "month": month,
            "month_label": month_label,
            "previous_month_label": previous_month_label,
            "raw_file": raw_file.name if raw_file else "",
            "range_file": range_file.name if range_file else "",
            "template_file": template_file.name if template_file else "",
            "range_sheet": range_sheet,
            "template_sheet": template_sheet,
            "raw_sheet": raw_summary.get("sheet", ""),
            "raw_rows": raw_summary.get("rows", ""),
            "raw_month_counts": raw_summary.get("month_counts", {}),
            "raw_account_counts": raw_summary.get("account_counts", {}),
        }
    )

    if validation_errors:
        for error in validation_errors:
            st.error(error)

    ready = all([month, month_label, previous_month_label, raw_path, range_path, template_path]) and not validation_errors
    if not ready:
        st.info("Upload Raw Export, Range Table, and Report Template workbooks.")
        return

    current_input_signature = input_signature(
        month=month,
        month_label=month_label,
        previous_month_label=previous_month_label,
        raw_path=raw_path,
        range_path=range_path,
        template_path=template_path,
        range_sheet=range_sheet,
        template_sheet=template_sheet,
        keep_history_columns=keep_history_columns,
    )
    if (
        st.session_state.get("last_report_state")
        and st.session_state["last_report_state"].get("input_signature") != current_input_signature
    ):
        st.session_state["last_report_state"] = None
        st.info("Inputs changed. Generate a new report so an old result is not shown with new selections.")

    if st.button("Generate Report", type="primary"):
        run_id = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S") + "_" + uuid.uuid4().hex[:8]
        run_output_root = OUTPUT_ROOT / st.session_state["session_id"] / run_id
        history_db = run_output_root / "retail_on_shelf_history.duckdb"
        command = [
            sys.executable,
            str(TOOL_DIR / "run_monthly_report.py"),
            "--month",
            month,
            "--month-label",
            month_label,
            "--previous-month-label",
            previous_month_label,
            "--raw-file",
            str(raw_path),
            "--range-file",
            str(range_path),
            "--range-sheet",
            range_sheet,
            "--range-source",
            "excel",
            "--disable-master-data",
            "--master-db",
            str(history_db),
            "--template-file",
            str(template_path),
            "--template-sheet",
            template_sheet,
        ]
        if keep_history_columns:
            command.append("--keep-history-columns")

        with st.spinner("Generating report..."):
            result = run_command(command, run_output_root, history_db)

        st.session_state["last_run_log"] = result.stdout + "\n" + result.stderr
        if result.returncode != 0:
            st.error("Report generation failed.")
            st.code(st.session_state["last_run_log"])
            return

        st.session_state["last_report_state"] = {
            "final_report": str(run_output_root / month / f"on_shelf_report_preview_{month}.xlsx"),
            "step3_csv": str(run_output_root / month / f"key_sku_display_rate_{month}.csv"),
            "count_csv": str(run_output_root / month / f"display_count_summary_{month}.csv"),
            "input_signature": current_input_signature,
        }
        st.success("Report generated.")

    if st.session_state.get("last_run_log"):
        with st.expander("Last Run Log", expanded=False):
            st.code(st.session_state["last_run_log"])

    if st.session_state.get("last_report_state"):
        state = st.session_state["last_report_state"]
        final_report = Path(state["final_report"])
        if final_report.exists():
            render_report(final_report, Path(state["step3_csv"]), Path(state["count_csv"]))


if __name__ == "__main__":
    main()
