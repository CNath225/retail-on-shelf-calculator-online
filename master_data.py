from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import re
from typing import Iterable, Optional

import duckdb
import pandas as pd


TOOL_DIR = Path(__file__).parent
DEFAULT_MASTER_DB = TOOL_DIR / "retail_on_shelf_history.duckdb"
DEFAULT_RANGE_FILE = TOOL_DIR / "range_template.xlsx"
DEFAULT_RANGE_SHEET = "Master data"

RANGE_COLUMNS = [
    "country",
    "category",
    "sku",
    "account",
    "ttl_store_count",
    "range_store_count",
    "range_percent",
    "active",
    "notes",
    "source",
    "updated_at_utc",
]

SKU_COLUMNS = [
    "category",
    "sku",
    "raw_column_aliases",
    "active",
    "notes",
    "updated_at_utc",
]

ACCOUNT_COLUMNS = [
    "account",
    "account_name",
    "active",
    "notes",
    "updated_at_utc",
]

COUNTRY_COLUMNS = [
    "country",
    "active",
    "notes",
    "updated_at_utc",
]

CATEGORY_COLUMNS = [
    "category",
    "active",
    "notes",
    "updated_at_utc",
]

DEFAULT_SKU_SPECS = [
    {"category": "Robot", "sku": "X60 Ultra", "raw_columns": ["X60 Ultra", "X60 Ultra "]},
    {"category": "Robot", "sku": "L50S Pro Ultra", "raw_columns": ["L50S Pro Ultra", "L50S Pro Ultra (JB)"]},
    {"category": "Robot", "sku": "L40 Ultra VE", "raw_columns": ["L40 Ultra VE", "L40 Ultra VE "]},
    {"category": "Robot", "sku": "L40 Plus", "raw_columns": ["L40 Plus", "L40 Plus "]},
    {"category": "Robot", "sku": "Matrix10 Ultra", "raw_columns": ["Matrix10 Ultra"]},
    {"category": "Robot", "sku": "Aqua10 Ultra Track S", "raw_columns": ["Aqua10 Ultra Track S", "Aqua10 Ultra Track S "]},
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

DEFAULT_ACCOUNT_NAME_TO_CODE = {
    "JB Hi-Fi": "JB",
    "Harvey Norman": "HN",
    "The Good Guys": "TGG",
    "Bing Lee": "BL",
    "David Jones": "DJS",
    "Betta Home Living Top 40": "BHLT",
}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_key(value: object) -> str:
    return clean_text(value).upper()


def normalize_column_name(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def split_aliases(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    return [part.strip() for part in re.split(r"[|;]", text) if part.strip()]


def join_aliases(values: Iterable[object]) -> str:
    cleaned = []
    seen = set()
    for value in values:
        text = clean_text(value)
        if not text:
            continue
        key = normalize_key(text)
        if key in seen:
            continue
        cleaned.append(text)
        seen.add(key)
    return " | ".join(cleaned)


def numeric_or_na(value: object):
    if pd.isna(value):
        return pd.NA
    if isinstance(value, str) and not value.strip():
        return pd.NA
    return pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]


def clean_count_text(value: object) -> str:
    number = numeric_or_na(value)
    if pd.notna(number):
        number = float(number)
        if number.is_integer():
            return str(int(number))
        return str(number)
    return clean_text(value)


def truthy(value: object) -> bool:
    if pd.isna(value):
        return True
    text = clean_text(value).lower()
    if text in {"", "true", "yes", "y", "1", "active"}:
        return True
    if text in {"false", "no", "n", "0", "inactive", "archived"}:
        return False
    return bool(value)


def table_exists(con: duckdb.DuckDBPyConnection, table_name: str) -> bool:
    return (
        con.execute(
            """
            SELECT COUNT(*)
            FROM information_schema.tables
            WHERE table_name = ?
            """,
            [table_name],
        ).fetchone()[0]
        > 0
    )


def ensure_master_schema(master_db: Path = DEFAULT_MASTER_DB) -> None:
    master_db.parent.mkdir(parents=True, exist_ok=True)
    with duckdb.connect(str(master_db)) as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS range_master (
                country VARCHAR,
                category VARCHAR,
                sku VARCHAR,
                account VARCHAR,
                ttl_store_count DOUBLE,
                range_store_count VARCHAR,
                range_percent DOUBLE,
                active BOOLEAN,
                notes VARCHAR,
                source VARCHAR,
                updated_at_utc VARCHAR
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS sku_master (
                category VARCHAR,
                sku VARCHAR,
                raw_column_aliases VARCHAR,
                active BOOLEAN,
                notes VARCHAR,
                updated_at_utc VARCHAR
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS account_master (
                account VARCHAR,
                account_name VARCHAR,
                active BOOLEAN,
                notes VARCHAR,
                updated_at_utc VARCHAR
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS country_master (
                country VARCHAR,
                active BOOLEAN,
                notes VARCHAR,
                updated_at_utc VARCHAR
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS category_master (
                category VARCHAR,
                active BOOLEAN,
                notes VARCHAR,
                updated_at_utc VARCHAR
            )
            """
        )


def standardize_range_columns(frame: pd.DataFrame) -> pd.DataFrame:
    alias_map = {
        "COUNTRY": "country",
        "CATEGORY": "category",
        "SKU": "sku",
        "ACCOUNT": "account",
        "CHANNEL": "account",
        "TTLSTORE": "ttl_store_count",
        "TTLSTORECOUNT": "ttl_store_count",
        "TOTALSTORE": "ttl_store_count",
        "TOTALSTORES": "ttl_store_count",
        "VISITEDSTORE": "visited_store_count",
        "VISITEDSTORECOUNT": "visited_store_count",
        "RANGE": "range_store_count",
        "RANGECOUNT": "range_store_count",
        "RANGESTORECOUNT": "range_store_count",
        "RANGEPERCENT": "range_percent",
        "ACTIVE": "active",
        "NOTES": "notes",
        "NOTE": "notes",
    }
    renamed = {}
    for column in frame.columns:
        normalized = normalize_column_name(column)
        if normalized in alias_map and alias_map[normalized] not in renamed.values():
            renamed[column] = alias_map[normalized]
    return frame.rename(columns=renamed)


def normalize_range_frame(frame: pd.DataFrame, source: str) -> pd.DataFrame:
    frame = standardize_range_columns(frame).copy()
    for column in ["country", "category", "sku", "account"]:
        if column not in frame.columns:
            raise ValueError(f"Missing required range column: {column}")

    result = pd.DataFrame()
    for column in ["country", "category", "sku", "account"]:
        result[column] = frame[column].apply(clean_text)

    ttl_source = (
        frame["ttl_store_count"]
        if "ttl_store_count" in frame.columns
        else pd.Series([pd.NA] * len(frame), index=frame.index)
    )
    result["ttl_store_count"] = pd.to_numeric(ttl_source, errors="coerce")
    result["range_store_count"] = (
        frame["range_store_count"].apply(clean_count_text)
        if "range_store_count" in frame.columns
        else ""
    )

    range_count_numeric = pd.to_numeric(result["range_store_count"], errors="coerce")
    calculated_range_percent = range_count_numeric / result["ttl_store_count"]
    calculated_range_percent = calculated_range_percent.where(
        result["ttl_store_count"].notna()
        & result["ttl_store_count"].ne(0)
        & range_count_numeric.notna()
    )
    if "range_store_count" in frame.columns:
        result["range_percent"] = calculated_range_percent
    else:
        result["range_percent"] = pd.to_numeric(frame.get("range_percent", pd.NA), errors="coerce")
    result["active"] = (
        frame["active"].apply(truthy) if "active" in frame.columns else True
    )
    result["notes"] = frame["notes"].apply(clean_text) if "notes" in frame.columns else ""
    result["source"] = frame["source"].apply(clean_text) if "source" in frame.columns else source
    result["updated_at_utc"] = (
        frame["updated_at_utc"].apply(clean_text)
        if "updated_at_utc" in frame.columns
        else utc_now()
    )
    result.loc[result["source"].eq(""), "source"] = source
    result.loc[result["updated_at_utc"].eq(""), "updated_at_utc"] = utc_now()
    return result[RANGE_COLUMNS]


def normalize_sku_frame(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    for column in SKU_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    result = frame[SKU_COLUMNS].copy()
    result["category"] = result["category"].apply(clean_text)
    result["sku"] = result["sku"].apply(clean_text)
    result["raw_column_aliases"] = result.apply(
        lambda row: join_aliases(split_aliases(row["raw_column_aliases"]) + [row["sku"]]),
        axis=1,
    )
    result["active"] = result["active"].apply(truthy)
    result["notes"] = result["notes"].apply(clean_text)
    result["updated_at_utc"] = result["updated_at_utc"].apply(clean_text)
    result.loc[result["updated_at_utc"].eq(""), "updated_at_utc"] = utc_now()
    return result[SKU_COLUMNS]


def normalize_account_frame(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    for column in ACCOUNT_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    result = frame[ACCOUNT_COLUMNS].copy()
    result["account"] = result["account"].apply(clean_text)
    result["account_name"] = result["account_name"].apply(clean_text)
    result["active"] = result["active"].apply(truthy)
    result["notes"] = result["notes"].apply(clean_text)
    result["updated_at_utc"] = result["updated_at_utc"].apply(clean_text)
    result.loc[result["updated_at_utc"].eq(""), "updated_at_utc"] = utc_now()
    return result[ACCOUNT_COLUMNS]


def normalize_simple_master_frame(frame: pd.DataFrame, columns: list[str], key_column: str) -> pd.DataFrame:
    frame = frame.copy()
    for column in columns:
        if column not in frame.columns:
            frame[column] = ""
    result = frame[columns].copy()
    result[key_column] = result[key_column].apply(clean_text)
    result["active"] = result["active"].apply(truthy)
    result["notes"] = result["notes"].apply(clean_text)
    result["updated_at_utc"] = result["updated_at_utc"].apply(clean_text)
    result.loc[result["updated_at_utc"].eq(""), "updated_at_utc"] = utc_now()
    return result[columns]


def validate_range_master(frame: pd.DataFrame) -> list[str]:
    errors = []
    active_frame = frame[frame["active"].apply(truthy)].copy()
    for column in ["country", "category", "sku", "account"]:
        blank = active_frame[column].apply(clean_text).eq("")
        if blank.any():
            errors.append(f"{column} has blank values in active rows.")

    ttl = pd.to_numeric(active_frame["ttl_store_count"], errors="coerce")
    bad_ttl = ttl.isna() | ttl.le(0)
    if bad_ttl.any():
        sample = active_frame.loc[bad_ttl, ["country", "category", "sku", "account"]].head(10)
        errors.append("Active rows need TTL Store# greater than 0:\n" + sample.to_string(index=False))

    range_numeric = pd.to_numeric(active_frame["range_store_count"], errors="coerce")
    negative_range = range_numeric.notna() & range_numeric.lt(0)
    if negative_range.any():
        sample = active_frame.loc[negative_range, ["country", "category", "sku", "account"]].head(10)
        errors.append("Range# cannot be negative:\n" + sample.to_string(index=False))

    too_large_range = range_numeric.notna() & ttl.notna() & range_numeric.gt(ttl)
    if too_large_range.any():
        sample = active_frame.loc[too_large_range, ["country", "category", "sku", "account", "ttl_store_count", "range_store_count"]].head(10)
        errors.append("Range# cannot be larger than TTL Store#:\n" + sample.to_string(index=False))

    duplicate_keys = active_frame.assign(
        key=active_frame[["country", "category", "sku", "account"]]
        .apply(lambda row: "|".join(normalize_key(value) for value in row), axis=1)
    )
    duplicate_mask = duplicate_keys["key"].duplicated(keep=False)
    if duplicate_mask.any():
        sample = active_frame.loc[duplicate_mask, ["country", "category", "sku", "account"]].head(20)
        errors.append("Duplicate active range keys found:\n" + sample.to_string(index=False))

    return errors


def validate_unique_master(frame: pd.DataFrame, key_columns: list[str], label: str) -> list[str]:
    active_frame = frame[frame["active"].apply(truthy)].copy()
    errors = []
    for column in key_columns:
        if active_frame[column].apply(clean_text).eq("").any():
            errors.append(f"{label}: {column} has blank values in active rows.")

    keys = active_frame[key_columns].apply(
        lambda row: "|".join(normalize_key(value) for value in row), axis=1
    )
    duplicate_mask = keys.duplicated(keep=False)
    if duplicate_mask.any():
        errors.append(
            f"{label}: duplicate active keys found:\n"
            + active_frame.loc[duplicate_mask, key_columns].head(20).to_string(index=False)
        )
    return errors


def replace_table(master_db: Path, table_name: str, frame: pd.DataFrame, columns: list[str]) -> None:
    ensure_master_schema(master_db)
    clean_frame = frame[columns].copy()
    with duckdb.connect(str(master_db)) as con:
        con.register("incoming_master_frame", clean_frame)
        con.execute(f"CREATE OR REPLACE TABLE {table_name} AS SELECT * FROM incoming_master_frame")
        con.unregister("incoming_master_frame")


def read_table(master_db: Path, table_name: str, columns: list[str]) -> pd.DataFrame:
    ensure_master_schema(master_db)
    with duckdb.connect(str(master_db)) as con:
        if not table_exists(con, table_name):
            return pd.DataFrame(columns=columns)
        frame = con.execute(f"SELECT * FROM {table_name}").fetchdf()
    for column in columns:
        if column not in frame.columns:
            frame[column] = ""
    return frame[columns]


def save_range_master(frame: pd.DataFrame, master_db: Path = DEFAULT_MASTER_DB) -> None:
    normalized = normalize_range_frame(frame, source="manual_edit")
    normalized["source"] = normalized["source"].replace("", "manual_edit")
    normalized["updated_at_utc"] = utc_now()
    errors = validate_range_master(normalized)
    if errors:
        raise ValueError("\n\n".join(errors))
    replace_table(master_db, "range_master", normalized, RANGE_COLUMNS)


def save_sku_master(frame: pd.DataFrame, master_db: Path = DEFAULT_MASTER_DB) -> None:
    normalized = normalize_sku_frame(frame)
    errors = validate_unique_master(normalized, ["category", "sku"], "SKU master")
    if errors:
        raise ValueError("\n\n".join(errors))
    replace_table(master_db, "sku_master", normalized, SKU_COLUMNS)


def save_account_master(frame: pd.DataFrame, master_db: Path = DEFAULT_MASTER_DB) -> None:
    normalized = normalize_account_frame(frame)
    errors = validate_unique_master(normalized, ["account"], "Account master")
    if errors:
        raise ValueError("\n\n".join(errors))
    replace_table(master_db, "account_master", normalized, ACCOUNT_COLUMNS)


def save_country_master(frame: pd.DataFrame, master_db: Path = DEFAULT_MASTER_DB) -> None:
    normalized = normalize_simple_master_frame(frame, COUNTRY_COLUMNS, "country")
    errors = validate_unique_master(normalized, ["country"], "Country master")
    if errors:
        raise ValueError("\n\n".join(errors))
    replace_table(master_db, "country_master", normalized, COUNTRY_COLUMNS)


def save_category_master(frame: pd.DataFrame, master_db: Path = DEFAULT_MASTER_DB) -> None:
    normalized = normalize_simple_master_frame(frame, CATEGORY_COLUMNS, "category")
    errors = validate_unique_master(normalized, ["category"], "Category master")
    if errors:
        raise ValueError("\n\n".join(errors))
    replace_table(master_db, "category_master", normalized, CATEGORY_COLUMNS)


def read_range_master(master_db: Path = DEFAULT_MASTER_DB, active_only: bool = False) -> pd.DataFrame:
    frame = read_table(master_db, "range_master", RANGE_COLUMNS)
    if active_only:
        frame = frame[frame["active"].apply(truthy)].copy()
    return normalize_range_frame(frame, source="master_db") if not frame.empty else frame


def read_sku_master(master_db: Path = DEFAULT_MASTER_DB, active_only: bool = False) -> pd.DataFrame:
    frame = read_table(master_db, "sku_master", SKU_COLUMNS)
    frame = normalize_sku_frame(frame) if not frame.empty else frame
    if active_only:
        frame = frame[frame["active"].apply(truthy)].copy()
    return frame


def read_account_master(master_db: Path = DEFAULT_MASTER_DB, active_only: bool = False) -> pd.DataFrame:
    frame = read_table(master_db, "account_master", ACCOUNT_COLUMNS)
    frame = normalize_account_frame(frame) if not frame.empty else frame
    if active_only:
        frame = frame[frame["active"].apply(truthy)].copy()
    return frame


def read_country_master(master_db: Path = DEFAULT_MASTER_DB, active_only: bool = False) -> pd.DataFrame:
    frame = read_table(master_db, "country_master", COUNTRY_COLUMNS)
    frame = normalize_simple_master_frame(frame, COUNTRY_COLUMNS, "country") if not frame.empty else frame
    if active_only:
        frame = frame[frame["active"].apply(truthy)].copy()
    return frame


def read_category_master(master_db: Path = DEFAULT_MASTER_DB, active_only: bool = False) -> pd.DataFrame:
    frame = read_table(master_db, "category_master", CATEGORY_COLUMNS)
    frame = normalize_simple_master_frame(frame, CATEGORY_COLUMNS, "category") if not frame.empty else frame
    if active_only:
        frame = frame[frame["active"].apply(truthy)].copy()
    return frame


def default_sku_master_frame() -> pd.DataFrame:
    rows = []
    now = utc_now()
    for spec in DEFAULT_SKU_SPECS:
        rows.append(
            {
                "category": spec["category"],
                "sku": spec["sku"],
                "raw_column_aliases": join_aliases(spec["raw_columns"]),
                "active": True,
                "notes": "",
                "updated_at_utc": now,
            }
        )
    return normalize_sku_frame(pd.DataFrame(rows))


def default_account_master_frame(range_frame: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    rows = []
    now = utc_now()
    for account_name, account in DEFAULT_ACCOUNT_NAME_TO_CODE.items():
        rows.append(
            {
                "account": account,
                "account_name": account_name,
                "active": True,
                "notes": "",
                "updated_at_utc": now,
            }
        )
    if range_frame is not None and not range_frame.empty:
        existing = {normalize_key(row["account"]) for row in rows}
        for account in sorted(range_frame["account"].dropna().astype(str).unique()):
            if normalize_key(account) not in existing:
                rows.append(
                    {
                        "account": clean_text(account),
                        "account_name": "",
                        "active": True,
                        "notes": "Created from range table.",
                        "updated_at_utc": now,
                    }
                )
    return normalize_account_frame(pd.DataFrame(rows))


def default_simple_master_frame(values: Iterable[object], key_column: str, note: str = "") -> pd.DataFrame:
    now = utc_now()
    rows = []
    seen = set()
    for value in values:
        text = clean_text(value)
        key = normalize_key(text)
        if not text or key in seen:
            continue
        rows.append(
            {
                key_column: text,
                "active": True,
                "notes": note,
                "updated_at_utc": now,
            }
        )
        seen.add(key)
    columns = COUNTRY_COLUMNS if key_column == "country" else CATEGORY_COLUMNS
    return normalize_simple_master_frame(pd.DataFrame(rows), columns, key_column)


def import_range_file(
    range_file: Path = DEFAULT_RANGE_FILE,
    sheet_name: str = DEFAULT_RANGE_SHEET,
    master_db: Path = DEFAULT_MASTER_DB,
    merge: bool = True,
) -> pd.DataFrame:
    uploaded = pd.read_excel(range_file, sheet_name=sheet_name)
    normalized_upload = normalize_range_frame(uploaded, source=f"import:{range_file.name}/{sheet_name}")
    if merge:
        existing = read_range_master(master_db)
        combined = merge_range_frames(existing, normalized_upload)
    else:
        combined = normalized_upload
    errors = validate_range_master(combined)
    if errors:
        raise ValueError("\n\n".join(errors))
    replace_table(master_db, "range_master", combined, RANGE_COLUMNS)
    return combined


def merge_range_frames(existing: pd.DataFrame, incoming: pd.DataFrame) -> pd.DataFrame:
    existing = normalize_range_frame(existing, source="existing") if not existing.empty else pd.DataFrame(columns=RANGE_COLUMNS)
    incoming = normalize_range_frame(incoming, source="incoming") if not incoming.empty else pd.DataFrame(columns=RANGE_COLUMNS)
    if existing.empty:
        return incoming[RANGE_COLUMNS].copy()
    if incoming.empty:
        return existing[RANGE_COLUMNS].copy()

    key_columns = ["country", "category", "sku", "account"]
    incoming_keys = set(
        incoming[key_columns].apply(lambda row: "|".join(normalize_key(value) for value in row), axis=1)
    )
    existing_keys = existing[key_columns].apply(lambda row: "|".join(normalize_key(value) for value in row), axis=1)
    kept_existing = existing[~existing_keys.isin(incoming_keys)]
    return pd.concat([kept_existing, incoming], ignore_index=True)[RANGE_COLUMNS]


def initialize_master_data(
    master_db: Path = DEFAULT_MASTER_DB,
    range_file: Path = DEFAULT_RANGE_FILE,
    range_sheet: str = DEFAULT_RANGE_SHEET,
) -> None:
    ensure_master_schema(master_db)
    range_frame = read_range_master(master_db)
    if range_frame.empty and range_file.exists():
        range_frame = import_range_file(range_file, range_sheet, master_db, merge=False)

    sku_frame = read_sku_master(master_db)
    if sku_frame.empty:
        default_skus = default_sku_master_frame()
        if not range_frame.empty:
            existing_keys = set(
                default_skus[["category", "sku"]].apply(
                    lambda row: "|".join(normalize_key(value) for value in row), axis=1
                )
            )
            extra_rows = []
            now = utc_now()
            for _, row in range_frame[["category", "sku"]].drop_duplicates().iterrows():
                key = "|".join(normalize_key(value) for value in row)
                if key not in existing_keys:
                    extra_rows.append(
                        {
                            "category": row["category"],
                            "sku": row["sku"],
                            "raw_column_aliases": row["sku"],
                            "active": False,
                            "notes": "Created from range table; activate after confirming raw column alias.",
                            "updated_at_utc": now,
                        }
                    )
            if extra_rows:
                default_skus = pd.concat([default_skus, pd.DataFrame(extra_rows)], ignore_index=True)
        save_sku_master(default_skus, master_db)

    account_frame = read_account_master(master_db)
    if account_frame.empty:
        save_account_master(default_account_master_frame(range_frame), master_db)

    country_frame = read_country_master(master_db)
    if country_frame.empty:
        countries = []
        if not range_frame.empty:
            countries.extend(range_frame["country"].tolist())
        save_country_master(default_simple_master_frame(countries or ["AU"], "country"), master_db)

    category_frame = read_category_master(master_db)
    if category_frame.empty:
        categories = []
        if not range_frame.empty:
            categories.extend(range_frame["category"].tolist())
        if not sku_frame.empty:
            categories.extend(sku_frame["category"].tolist())
        save_category_master(default_simple_master_frame(categories, "category"), master_db)


def active_sku_specs_from_master(master_db: Path = DEFAULT_MASTER_DB) -> list[dict[str, object]]:
    initialize_master_data(master_db)
    sku_frame = read_sku_master(master_db, active_only=True)
    specs = []
    for _, row in sku_frame.iterrows():
        aliases = split_aliases(row["raw_column_aliases"])
        if not aliases:
            aliases = [row["sku"]]
        specs.append(
            {
                "category": row["category"],
                "sku": row["sku"],
                "raw_columns": aliases,
            }
        )
    return specs


def active_account_name_map(master_db: Path = DEFAULT_MASTER_DB) -> dict[str, str]:
    initialize_master_data(master_db)
    account_frame = read_account_master(master_db, active_only=True)
    result = dict(DEFAULT_ACCOUNT_NAME_TO_CODE)
    for _, row in account_frame.iterrows():
        account_name = clean_text(row["account_name"])
        account = clean_text(row["account"])
        if account_name and account:
            result[account_name] = account
    return result


def range_master_for_step3(master_db: Path = DEFAULT_MASTER_DB) -> pd.DataFrame:
    initialize_master_data(master_db)
    frame = read_range_master(master_db, active_only=True)
    return frame[
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


def scan_unknown_display_columns(raw_file: Path, master_db: Path = DEFAULT_MASTER_DB) -> pd.DataFrame:
    initialize_master_data(master_db)
    raw_df = pd.read_excel(raw_file, sheet_name="Submissions")
    known_aliases = set()
    for _, row in read_sku_master(master_db, active_only=True).iterrows():
        known_aliases.update(normalize_key(alias) for alias in split_aliases(row["raw_column_aliases"]))

    status_words = re.compile(r"display|ranged|stock|fixture|space", re.IGNORECASE)
    rows = []
    for column in raw_df.columns:
        if normalize_key(column) in known_aliases:
            continue
        values = raw_df[column].dropna().astype(str)
        sample_values = values[values.str.contains(status_words, na=False)].head(5).tolist()
        if sample_values:
            rows.append(
                {
                    "raw_column": column,
                    "non_blank_rows": int(values.ne("").sum()),
                    "sample_values": " | ".join(sample_values),
                    "suggestion": "Add this as a SKU raw alias if it is a display SKU.",
                }
            )
    return pd.DataFrame(rows)
