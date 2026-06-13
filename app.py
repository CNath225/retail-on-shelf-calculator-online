from datetime import datetime, timezone
import hashlib
import json
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

from identifier_matching import (
    CanonicalEntry,
    build_resolution_items,
    clean_text,
    display_like_columns,
    items_to_frame,
    normalize_identifier,
    summarize_items,
)
from alias_workbook import create_alias_decisions_workbook, load_alias_decisions_from_workbook

TOOL_DIR = Path(__file__).parent
RUNTIME_DIR = TOOL_DIR / ".runtime"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_ROOT = RUNTIME_DIR / "outputs"
APP_TITLE = "Retail On-shelf Rate Calculator Online by CodeNATHAN"
APP_VERSION = "v1.2-beta"

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


def app_title_html() -> str:
    return """
        <style>
          .app-title {
            margin: 0.25rem 0 0.7rem;
            font-size: clamp(2.1rem, 4vw, 3.5rem);
            font-weight: 800;
            line-height: 1.08;
            color: inherit;
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
          .app-version {
            display: inline-block;
            margin: 0;
            padding: 0.16rem 0.52rem;
            border: 1px solid rgba(185, 148, 78, 0.5);
            border-radius: 6px;
            color: #c8b27d;
            background: rgba(32, 34, 38, 0.72);
            font-size: 0.82rem;
            letter-spacing: 0;
          }
          .version-row {
            display: flex;
            align-items: flex-start;
            gap: 0.42rem;
            margin: 0 0 1.25rem;
          }
          .patch-notes {
            position: relative;
            display: inline-block;
            font-size: 0.82rem;
            line-height: 1.35;
            color: #c8a760;
          }
          .patch-notes summary {
            list-style: none;
            cursor: pointer;
            user-select: none;
            padding: 0.16rem 0.58rem;
            border: 1px solid rgba(185, 148, 78, 0.78);
            border-radius: 6px;
            color: #c8a760;
            background: linear-gradient(180deg, rgba(40, 34, 22, 0.96), rgba(21, 18, 14, 0.96));
            box-shadow: inset 0 0 0 1px rgba(255, 221, 137, 0.12), 0 0 12px rgba(169, 130, 53, 0.16);
          }
          .patch-notes summary::-webkit-details-marker {
            display: none;
          }
          .patch-notes summary::after {
            content: "▾";
            margin-left: 0.36rem;
            color: #f0cf80;
          }
          .patch-notes[open] summary {
            border-bottom-color: rgba(240, 207, 128, 0.95);
            box-shadow: inset 0 0 0 1px rgba(255, 221, 137, 0.2), 0 0 18px rgba(169, 130, 53, 0.28);
          }
          .patch-notes[open] summary::after {
            content: "▴";
          }
          .patch-panel {
            position: absolute;
            z-index: 20;
            top: calc(100% + 0.38rem);
            left: 0;
            width: min(30rem, calc(100vw - 2rem));
            padding: 0.85rem 0.95rem;
            border: 1px solid rgba(200, 167, 96, 0.7);
            border-radius: 8px;
            color: #efe5cc;
            background: linear-gradient(180deg, rgba(25, 25, 28, 0.98), rgba(12, 12, 14, 0.98));
            box-shadow: 0 18px 40px rgba(0, 0, 0, 0.38), inset 0 0 0 1px rgba(255, 255, 255, 0.04);
          }
          .patch-version {
            margin: 0 0 0.72rem;
            padding-left: 0.65rem;
            border-left: 2px solid rgba(200, 167, 96, 0.78);
          }
          .patch-version:last-child {
            margin-bottom: 0;
          }
          .patch-version strong {
            display: block;
            margin-bottom: 0.22rem;
            color: #f0cf80;
          }
          .patch-version ul {
            margin: 0;
            padding-left: 1rem;
          }
          .patch-version li {
            margin: 0.12rem 0;
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
        <div class="version-row">
          <div class="app-version">APP_VERSION_PLACEHOLDER</div>
          <details class="patch-notes">
            <summary>Patch notes</summary>
            <div class="patch-panel">
              <div class="patch-version">
                <strong>v1-beta</strong>
                <ul>
                  <li>Online upload flow for raw export, range table, and report template.</li>
                  <li>Raw sheet detection, missing ID handling, month checks, and stale-result hiding.</li>
                </ul>
              </div>
              <div class="patch-version">
                <strong>v1.1-beta</strong>
                <ul>
                  <li>Smart identifier matching beta, default off.</li>
                  <li>Category-TTL fix includes single-SKU channels; W&amp;D matches 53%.</li>
                  <li>Report colour formatting and theme-adaptive title polish.</li>
                </ul>
              </div>
              <div class="patch-version">
                <strong>v1.2-beta</strong>
                <ul>
                  <li>Alias map is embedded in a hidden report sheet.</li>
                  <li>Generated reports now work as rolling templates with previous + current month columns.</li>
                  <li>Added this changelog dropdown.</li>
                </ul>
              </div>
              <div class="patch-version">
                <strong>future</strong>
                <ul>
                  <li>Reserved for upcoming approved changes.</li>
                </ul>
              </div>
            </div>
          </details>
        </div>
        """.replace("APP_VERSION_PLACEHOLDER", APP_VERSION)


def render_app_title() -> None:
    html = app_title_html()
    st.markdown(
        html,
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
    smart_matching_enabled: bool,
    alias_decision_signature: str,
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
        "smart_matching_enabled": smart_matching_enabled,
        "alias_decision_signature": alias_decision_signature if smart_matching_enabled else "",
    }


def render_report(
    final_report: Path,
    step3_csv: Path,
    count_csv: Path,
    identifier_quality_csv: Optional[Path] = None,
) -> None:
    st.download_button(
        "Download Report",
        data=final_report.read_bytes(),
        file_name=final_report.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click="ignore",
    )
    has_presentation_sheet = "For Presentation" in xlsx_sheets(final_report)
    tabs = (
        ["Preview", "For Presentation", "Rate Detail", "Display Counts"]
        if has_presentation_sheet
        else ["Preview", "Rate Detail", "Display Counts"]
    )
    has_identifier_quality = bool(identifier_quality_csv and identifier_quality_csv.exists())
    if has_identifier_quality:
        tabs.append("Identifier Quality")
    rendered_tabs = st.tabs(tabs)
    preview_tab = rendered_tabs[0]
    next_tab_index = 1
    presentation_tab = rendered_tabs[next_tab_index] if has_presentation_sheet else None
    if has_presentation_sheet:
        next_tab_index += 1
    detail_tab = rendered_tabs[next_tab_index]
    count_tab = rendered_tabs[next_tab_index + 1]
    identifier_tab = rendered_tabs[next_tab_index + 2] if has_identifier_quality else None
    with preview_tab:
        from step4_generate_report_preview import style_report_preview

        preview_df = pd.read_excel(final_report, sheet_name="Report Preview")
        month_label = st.session_state.get("month_label", "JUN")
        try:
            st.dataframe(
                style_report_preview(preview_df, month_label),
                use_container_width=True,
                height=520,
            )
        except Exception:
            st.dataframe(preview_df, use_container_width=True, height=520)
    if presentation_tab is not None:
        with presentation_tab:
            from step4_generate_report_preview import style_presentation_frame

            presentation_df = pd.read_excel(final_report, sheet_name="For Presentation")
            month_label = st.session_state.get("month_label", "JUN")
            try:
                st.dataframe(
                    style_presentation_frame(presentation_df, month_label),
                    use_container_width=True,
                    height=520,
                )
            except Exception:
                st.dataframe(presentation_df, use_container_width=True, height=520)
    with detail_tab:
        if step3_csv.exists():
            st.dataframe(pd.read_csv(step3_csv), use_container_width=True, height=520)
    with count_tab:
        if count_csv.exists():
            st.dataframe(pd.read_csv(count_csv), use_container_width=True, height=520)
    if identifier_tab is not None:
        with identifier_tab:
            st.dataframe(pd.read_csv(identifier_quality_csv), use_container_width=True, height=520)


def init_state() -> None:
    st.session_state.setdefault("session_id", uuid.uuid4().hex)
    st.session_state.setdefault("month", "2026-06")
    st.session_state.setdefault("month_label", "JUN")
    st.session_state.setdefault("previous_month_label", "MAY")
    st.session_state.setdefault("_last_raw_filename", "")
    st.session_state.setdefault("last_report_state", None)
    st.session_state.setdefault("last_run_log", "")
    st.session_state.setdefault("identifier_alias_decisions", [])


def clear_session_files() -> None:
    session_id = st.session_state["session_id"]
    for path in [UPLOAD_DIR / session_id, OUTPUT_ROOT / session_id]:
        if path.exists():
            shutil.rmtree(path)
    st.session_state["last_report_state"] = None
    st.session_state["last_run_log"] = ""


def decision_key(domain: str, raw_value: object) -> str:
    digest = hashlib.sha1(f"{domain}|{clean_text(raw_value)}".encode("utf-8")).hexdigest()[:12]
    return f"{domain}_{digest}"


def alias_decision_signature(decisions: list[dict[str, object]]) -> str:
    payload = json.dumps(decisions, ensure_ascii=False, sort_keys=True)
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def template_alias_signature(path: Optional[Path]) -> str:
    signature = file_signature(path)
    return json.dumps(signature, ensure_ascii=False, sort_keys=True) if signature else ""


def load_template_alias_decisions(path: Optional[Path]) -> None:
    signature = template_alias_signature(path)
    if not signature or st.session_state.get("_last_template_alias_source") == signature:
        return
    try:
        decisions = load_alias_decisions_from_workbook(path)
    except Exception as error:
        st.warning(f"Could not load embedded alias map from template: {error}")
        decisions = []
    st.session_state["identifier_alias_decisions"] = decisions
    st.session_state["_last_template_alias_source"] = signature
    st.session_state["_template_alias_decision_count"] = len(decisions)


def read_sheet_safe(path: Optional[Path], sheet: str) -> pd.DataFrame:
    if not path or not path.exists() or not sheet:
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def unique_column_values(frame: pd.DataFrame, columns: list[str]) -> list[str]:
    values: set[str] = set()
    normalized_columns = {normalize_column_label(column): column for column in frame.columns}
    for wanted in columns:
        source = normalized_columns.get(normalize_column_label(wanted))
        if source is None:
            continue
        values.update(clean_text(value) for value in frame[source].dropna().tolist() if clean_text(value))
    return sorted(values)


def canonical_entries_from_sources(
    range_path: Optional[Path],
    range_sheet: str,
    template_path: Optional[Path],
    template_sheet: str,
) -> dict[str, list[CanonicalEntry]]:
    from step1_prepare_raw_data import ACCOUNT_NAME_TO_CODE, SKU_SPECS

    range_frame = read_sheet_safe(range_path, range_sheet)
    template_frame = read_sheet_safe(template_path, template_sheet)

    sku_entries: dict[str, CanonicalEntry] = {}
    for spec in SKU_SPECS:
        sku = clean_text(spec.get("sku", ""))
        if not sku:
            continue
        sku_entries[normalize_identifier(sku).key] = CanonicalEntry(
            domain="sku",
            canonical=sku,
            aliases=tuple(str(value) for value in spec.get("raw_columns", [])),
            metadata={"category": clean_text(spec.get("category", ""))},
        )
    for _, row in range_frame.iterrows():
        sku = clean_text(row.get("SKU", row.get("sku", "")))
        category = clean_text(row.get("Category", row.get("category", "")))
        if sku and normalize_identifier(sku).key not in sku_entries:
            sku_entries[normalize_identifier(sku).key] = CanonicalEntry(
                domain="sku",
                canonical=sku,
                aliases=(sku,),
                metadata={"category": category},
            )

    account_entries = {
        normalize_identifier(account).key: CanonicalEntry(
            domain="account",
            canonical=account,
            aliases=tuple(
                account_name
                for account_name, mapped_account in ACCOUNT_NAME_TO_CODE.items()
                if mapped_account == account
            ),
        )
        for account in sorted(set(ACCOUNT_NAME_TO_CODE.values()) | set(unique_column_values(range_frame, ["Account", "Channel"])))
        if account
    }

    category_values = set(unique_column_values(range_frame, ["Category"]))
    category_values.update(unique_column_values(template_frame, ["Category"]))
    category_values.update(entry.metadata.get("category", "") for entry in sku_entries.values())
    category_entries = [
        CanonicalEntry(domain="category", canonical=value)
        for value in sorted(value for value in category_values if value)
    ]

    country_values = set(unique_column_values(range_frame, ["Country"]))
    country_values.update(unique_column_values(template_frame, ["Country"]))
    country_values.update(["AU"])
    country_entries = [
        CanonicalEntry(domain="country", canonical=value)
        for value in sorted(value for value in country_values if value)
    ]

    return {
        "sku": list(sku_entries.values()),
        "account": list(account_entries.values()),
        "category": category_entries,
        "country": country_entries,
    }


def raw_identifier_values(
    raw_path: Optional[Path],
    range_path: Optional[Path],
    range_sheet: str,
    template_path: Optional[Path],
    template_sheet: str,
) -> dict[str, list[str]]:
    if not raw_path or not raw_path.exists():
        return {"sku": [], "account": [], "category": [], "country": []}
    from step1_prepare_raw_data import (
        find_raw_submission_sheet,
        get_account_from_place_id,
        normalize_raw_submission_columns,
    )

    raw_sheet = find_raw_submission_sheet(raw_path)
    raw_frame = pd.read_excel(raw_path, sheet_name=raw_sheet)
    raw_frame = normalize_raw_submission_columns(raw_frame)
    raw_frame = raw_frame[raw_frame["Place ID"].notna()].copy()

    range_frame = read_sheet_safe(range_path, range_sheet)
    template_frame = read_sheet_safe(template_path, template_sheet)

    account_values = set()
    for column in ["Account Name", "Account"]:
        if column in raw_frame.columns:
            account_values.update(clean_text(value) for value in raw_frame[column].dropna().tolist() if clean_text(value))
    account_values.update(
        clean_text(get_account_from_place_id(value))
        for value in raw_frame["Place ID"].dropna().tolist()
        if clean_text(get_account_from_place_id(value))
    )

    country_values = (
        set(clean_text(value) for value in raw_frame["Country"].dropna().tolist() if clean_text(value))
        if "Country" in raw_frame.columns
        else {"AU"}
    )

    category_values = set(unique_column_values(range_frame, ["Category"]))
    category_values.update(unique_column_values(template_frame, ["Category"]))

    return {
        "sku": display_like_columns(raw_frame),
        "account": sorted(account_values),
        "category": sorted(category_values),
        "country": sorted(country_values),
    }


def build_identifier_resolution(
    raw_path: Optional[Path],
    range_path: Optional[Path],
    range_sheet: str,
    template_path: Optional[Path],
    template_sheet: str,
    decisions: list[dict[str, object]],
) -> tuple[list[dict[str, object]], dict[str, list[CanonicalEntry]]]:
    canonicals = canonical_entries_from_sources(range_path, range_sheet, template_path, template_sheet)
    raw_values = raw_identifier_values(raw_path, range_path, range_sheet, template_path, template_sheet)
    items = []
    for domain in ["sku", "account", "category", "country"]:
        items.extend(
            item.as_dict()
            for item in build_resolution_items(
                domain,
                raw_values[domain],
                canonicals[domain],
                decisions=decisions,
            )
        )
    return items, canonicals


def auto_decisions_from_quality(quality_frame: pd.DataFrame) -> list[dict[str, object]]:
    if quality_frame.empty:
        return []
    rows = []
    for _, row in quality_frame[quality_frame["status"].eq("auto_normalized")].iterrows():
        rows.append(
            {
                "domain": row["domain"],
                "raw_value": row["raw_value"],
                "canonical": row["canonical"],
                "action": "map",
                "channel": "A",
                "reason": row["reason"],
            }
        )
    return rows


def merge_alias_decisions(
    auto_decisions: list[dict[str, object]],
    manual_decisions: list[dict[str, object]],
) -> list[dict[str, object]]:
    merged = {
        decision_key(decision.get("domain", ""), decision.get("raw_value", "")): decision
        for decision in auto_decisions
    }
    for decision in manual_decisions:
        merged[decision_key(decision.get("domain", ""), decision.get("raw_value", ""))] = decision
    return list(merged.values())


def valid_manual_decision(decision: dict[str, object]) -> bool:
    action = decision.get("action", "")
    if action in {"skip_this_run", "skip_always", "map"}:
        return True
    if action == "add_new" and decision.get("domain") in {"sku", "account"}:
        required = ["canonical", "category", "country", "ttl_store_count", "range_store_count"]
        return all(clean_text(decision.get(column, "")) for column in required)
    if action == "add_new":
        return bool(clean_text(decision.get("canonical", "")))
    return False


def render_identifier_resolution(
    raw_path: Optional[Path],
    range_path: Optional[Path],
    range_sheet: str,
    template_path: Optional[Path],
    template_sheet: str,
) -> tuple[bool, list[dict[str, object]], pd.DataFrame]:
    decisions = st.session_state.get("identifier_alias_decisions", [])
    items, canonicals = build_identifier_resolution(
        raw_path,
        range_path,
        range_sheet,
        template_path,
        template_sheet,
        decisions,
    )
    item_objects = [
        build_resolution_items(
            item["domain"],
            [item["raw_value"]],
            canonicals[item["domain"]],
            decisions=decisions,
        )[0]
        for item in items
    ]
    summary_by_domain = []
    for domain in ["sku", "account", "category", "country"]:
        domain_items = [item for item in item_objects if item.domain == domain]
        summary = summarize_items(domain_items)
        summary_by_domain.append({"domain": domain, **summary})

    quality_frame = pd.DataFrame([item.as_dict() for item in item_objects])
    st.subheader("Beta Identifier Resolution")
    st.caption("Beta ON: report generation is blocked until review/ambiguous/unmatched values are resolved or explicitly skipped.")
    st.dataframe(pd.DataFrame(summary_by_domain), use_container_width=True, height=180)

    auto_frame = quality_frame[quality_frame["status"].eq("auto_normalized")] if not quality_frame.empty else pd.DataFrame()
    pending_frame = quality_frame[quality_frame["status"].isin(["confusable", "ambiguous", "unmatched"])] if not quality_frame.empty else pd.DataFrame()

    with st.expander("Auto-normalized spelling-only matches", expanded=False):
        st.dataframe(auto_frame, use_container_width=True, height=260)

    saved_decisions = {decision_key(item["domain"], item["raw_value"]): item for item in decisions}
    draft_decisions: list[dict[str, object]] = []
    if pending_frame.empty:
        st.success("No pending identifier issues. Beta smart matching can generate the report.")
        return True, decisions, quality_frame

    st.warning("Some identifiers need confirmation before beta report generation.")
    for _, row in pending_frame.head(80).iterrows():
        key = decision_key(row["domain"], row["raw_value"])
        existing_decision = saved_decisions.get(key, {})
        st.markdown(f"**{row['domain'].upper()}**: `{row['raw_value']}`")
        if row["status"] == "confusable":
            st.error(
                "Confusable pair: likely DIFFERENT entity. "
                f"Diff: {row.get('differing_token', '')}; near: {row.get('confusable_with', [])}"
            )
        elif row["status"] == "ambiguous":
            st.error(f"Ambiguous: multiple candidates {row.get('candidates', [])}")

        options = ["", "Map to existing", "Add as new", "Skip this run", "Skip always"]
        default_option = "Add as new" if row["status"] == "confusable" else ""
        if existing_decision:
            existing_action = existing_decision.get("action", "")
            default_option = {
                "map": "Map to existing",
                "add_new": "Add as new",
                "skip_this_run": "Skip this run",
                "skip_always": "Skip always",
            }.get(existing_action, default_option)
        action = st.selectbox(
            "Action",
            options,
            index=options.index(default_option) if default_option in options else 0,
            key=f"{key}_action",
        )
        canonical_options = [entry.canonical for entry in canonicals[row["domain"]]]
        canonical = existing_decision.get("canonical", row["raw_value"])
        if action == "Map to existing":
            canonical = st.selectbox(
                "Existing master entry",
                canonical_options,
                index=canonical_options.index(canonical) if canonical in canonical_options else 0,
                key=f"{key}_canonical",
            )
        elif action == "Add as new":
            canonical = st.text_input(
                "New canonical value",
                value=str(existing_decision.get("canonical", row["raw_value"])),
                key=f"{key}_new_canonical",
            )
            if row["domain"] in {"sku", "account"}:
                category = st.text_input(
                    "Category required for new SKU/account",
                    value=str(existing_decision.get("category", "")),
                    key=f"{key}_category",
                )
                country = st.text_input(
                    "Country required for new SKU/account",
                    value=str(existing_decision.get("country", "AU")),
                    key=f"{key}_country",
                )
                ttl_store = st.text_input(
                    "TTL Store# required for new SKU/account",
                    value=str(existing_decision.get("ttl_store_count", "")),
                    key=f"{key}_ttl",
                )
                range_store = st.text_input(
                    "Range# required for new SKU/account",
                    value=str(existing_decision.get("range_store_count", "")),
                    key=f"{key}_range",
                )
            else:
                category = ""
                country = ""
                ttl_store = ""
                range_store = ""
        else:
            category = ""
            country = ""
            ttl_store = ""
            range_store = ""

        if action:
            internal_action = {
                "Map to existing": "map",
                "Add as new": "add_new",
                "Skip this run": "skip_this_run",
                "Skip always": "skip_always",
            }[action]
            draft = {
                "domain": row["domain"],
                "raw_value": row["raw_value"],
                "canonical": canonical if internal_action in {"map", "add_new"} else "",
                "action": internal_action,
                "channel": "B" if row["status"] in {"confusable", "ambiguous"} else "manual",
                "status_at_review": row["status"],
            }
            if internal_action == "add_new" and row["domain"] in {"sku", "account"}:
                draft.update(
                    {
                        "category": category,
                        "country": country,
                        "ttl_store_count": ttl_store,
                        "range_store_count": range_store,
                    }
                )
            draft_decisions.append(draft)

    if st.button("Save Beta Resolution Decisions"):
        invalid = [decision for decision in draft_decisions if not valid_manual_decision(decision)]
        if invalid:
            st.error("Some beta decisions are incomplete. New SKU/account needs category, country, TTL Store#, and Range#.")
            st.stop()
        merged = [decision for decision in decisions if decision_key(decision.get("domain", ""), decision.get("raw_value", "")) not in {decision_key(item["domain"], item["raw_value"]) for item in draft_decisions}]
        merged.extend(draft_decisions)
        st.session_state["identifier_alias_decisions"] = merged
        st.rerun()

    resolved_keys = {
        decision_key(decision.get("domain", ""), decision.get("raw_value", ""))
        for decision in st.session_state.get("identifier_alias_decisions", [])
        if valid_manual_decision(decision)
    }
    pending_keys = {
        decision_key(row["domain"], row["raw_value"])
        for _, row in pending_frame.iterrows()
    }
    ready = pending_keys.issubset(resolved_keys)
    if not ready:
        st.info("Resolve or skip every pending identifier before generating the beta report.")
    return ready, st.session_state.get("identifier_alias_decisions", []), quality_frame


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    render_app_title()

    with st.sidebar:
        raw_file = st.file_uploader("Resply Raw Export", type=["xlsx"])
        range_file = st.file_uploader("Range Table", type=["xlsx"])
        template_file = st.file_uploader("Report Template", type=["xlsx"])
        smart_matching_enabled = st.checkbox(
            "Enable beta: smart identifier matching",
            value=False,
            help=(
                "Beta. Auto-matches raw SKU / account / category / country names to your master "
                "entries when they differ only by channel suffix (e.g. '(JB)') or spelling, and flags "
                "genuine look-alikes for you to confirm. Default OFF keeps the v1.0 matching path unchanged."
            ),
        )
        st.caption(
            "Beta matching stores confirmed alias decisions inside the generated report workbook. "
            "Upload last month's report as this month's template to reuse them."
        )

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
    load_template_alias_decisions(template_path)

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
        keep_history_columns = False
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
            "smart_identifier_matching": "ON" if smart_matching_enabled else "OFF",
            "app_version": APP_VERSION,
        }
    )

    if validation_errors:
        for error in validation_errors:
            st.error(error)

    ready = all([month, month_label, previous_month_label, raw_path, range_path, template_path]) and not validation_errors
    if not ready:
        st.info("Upload Raw Export, Range Table, and Report Template workbooks.")
        return

    identifier_quality_frame = pd.DataFrame()
    beta_alias_decisions = []
    if smart_matching_enabled:
        beta_ready, beta_alias_decisions, identifier_quality_frame = render_identifier_resolution(
            raw_path,
            range_path,
            range_sheet,
            template_path,
            template_sheet,
        )
        ready = ready and beta_ready
        if not ready:
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
        smart_matching_enabled=smart_matching_enabled,
        alias_decision_signature=alias_decision_signature(beta_alias_decisions),
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
        identifier_quality_csv = run_output_root / month / f"identifier_quality_{month}.csv"
        alias_workbook_for_run = template_path
        alias_decisions_for_run: list[dict[str, object]] = []
        if smart_matching_enabled:
            run_output_root.mkdir(parents=True, exist_ok=True)
            (run_output_root / month).mkdir(parents=True, exist_ok=True)
            identifier_quality_frame.to_csv(identifier_quality_csv, index=False)
            alias_decisions_for_run = merge_alias_decisions(
                auto_decisions_from_quality(identifier_quality_frame),
                beta_alias_decisions,
            )
            alias_workbook_for_run = create_alias_decisions_workbook(
                target=run_output_root / "identifier_alias_workbook.xlsx",
                decisions=alias_decisions_for_run,
            )
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
        if smart_matching_enabled:
            command.extend(
                [
                    "--enable-smart-matching",
                    "--identifier-alias-workbook",
                    str(alias_workbook_for_run),
                ]
            )
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
            "identifier_quality_csv": str(identifier_quality_csv) if smart_matching_enabled else "",
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
            identifier_quality_path = (
                Path(state["identifier_quality_csv"])
                if state.get("identifier_quality_csv")
                else None
            )
            render_report(
                final_report,
                Path(state["step3_csv"]),
                Path(state["count_csv"]),
                identifier_quality_path,
            )


if __name__ == "__main__":
    main()
