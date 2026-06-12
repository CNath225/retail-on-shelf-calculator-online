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


def previous_label(month: str) -> str:
    match = re.match(r"^(20\d{2})-(0[1-9]|1[0-2])$", month)
    if not match:
        return st.session_state.get("previous_month_label", "May")
    month_number = int(match.group(2))
    previous_number = 12 if month_number == 1 else month_number - 1
    return MONTH_LABELS[previous_number].title()


def normalize_column_label(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def choose_report_column(month_label: Optional[str], template_columns: list[str]) -> Optional[str]:
    if not month_label:
        return None
    wanted = normalize_column_label(month_label)
    for column in template_columns:
        if normalize_column_label(column) == wanted:
            return str(column)
    return month_label


def run_command(command: list[str], output_root: Path, history_db: Path) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["ONSHELF_OUTPUT_ROOT"] = str(output_root)
    env["ONSHELF_HISTORY_DB"] = str(history_db)
    env["PYTHONPYCACHEPREFIX"] = str(output_root / "_pycache")
    return subprocess.run(command, capture_output=True, text=True, env=env)


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
    st.session_state.setdefault("previous_month_label", "May")
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

    st.title(APP_TITLE)

    with st.sidebar:
        raw_file = st.file_uploader("Resply Raw Export", type=["xlsx"])
        range_file = st.file_uploader("Range Table", type=["xlsx"])
        use_range_as_template = st.checkbox("Use range file as template", value=True)
        template_file = None
        if not use_range_as_template:
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
    template_path = range_path if use_range_as_template else save_upload(template_file, session_upload_dir)

    range_sheets = xlsx_sheets(range_path)
    template_sheets = xlsx_sheets(template_path)
    template_columns = []

    with st.sidebar:
        range_sheet = st.selectbox(
            "Range Sheet",
            range_sheets or ["Master data"],
            index=(
                range_sheets.index(preferred_range_sheet(range_sheets))
                if preferred_range_sheet(range_sheets) in range_sheets
                else 0
            ),
        )
        template_sheet = st.selectbox(
            "Template Sheet",
            template_sheets or ["ANZ On-Shelf Retailer"],
            index=(
                template_sheets.index("ANZ On-Shelf Retailer")
                if "ANZ On-Shelf Retailer" in template_sheets
                else 0
            ),
        )

        if template_path and template_path.exists():
            try:
                template_columns = list(pd.read_excel(template_path, sheet_name=template_sheet, nrows=0).columns)
            except Exception:
                template_columns = []
        preferred_column = choose_report_column(st.session_state["month_label"], template_columns)
        if preferred_column:
            st.session_state["month_label"] = preferred_column

        month = st.text_input("Month", key="month")
        month_label = st.text_input("Report Column", key="month_label")
        previous_month_label = st.text_input("Compare To", key="previous_month_label")
        keep_history_columns = st.checkbox("Keep history columns", value=False)
        st.button("Clear Session Files", on_click=clear_session_files)

    st.subheader("Inputs")
    st.write(
        {
            "month": month,
            "month_label": month_label,
            "previous_month_label": previous_month_label,
            "raw_file": raw_file.name if raw_file else "",
            "range_file": range_file.name if range_file else "",
            "template_file": (range_file.name if use_range_as_template and range_file else template_file.name if template_file else ""),
            "range_sheet": range_sheet,
            "template_sheet": template_sheet,
        }
    )

    ready = all([month, month_label, previous_month_label, raw_path, range_path, template_path])
    if not ready:
        st.info("Upload the raw export and range/template workbook.")
        return

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
