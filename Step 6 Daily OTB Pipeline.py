"""
Step 6 Daily OTB Pipeline
=========================
Combines Steps 1-5 of the Daily OTB pipeline into a single script:
  1. Reference data loading (per-property workbooks)
  2. File discovery & readers (raw CSV -> standardized columns)
  3. Cleaning (totals extraction, dedup, date normalization, rate code
     validation, null checks, historical consistency)
  4. Business logic (pickup, pace/STLY variance, monthly roll-up with
     RevPAR/Occupancy %, market segment roll-up)
  5. Validation (total-row reconciliation against source, industry-
     standard +/-3% tolerance)

Run this script directly (`python "Step 6 Daily OTB Pipeline.py"`) to process
both properties end to end. Designed to run once per day each morning.

NOT included here: the hand-verified example spot-checks from Step 5
(specific hardcoded StayDates/RateCodes like 'exp'/Aug 20 or 'bar'/Aug 11).
Those were one-time proof points during development, not a production
check that should run against every day's new data -- they live in
Step 5's original notebook for reference.
"""

import calendar
import re
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd

pd.set_option("display.max_columns", 50)


# ============================================================
# CONFIG
# ============================================================

REFERENCE_DATA_PATHS = {
    63638: Path(r"D:\OTB Reports\Rate Code OTB This Year - The Villages\Reference Data OE Villages.xlsx"),
    23929: Path(r"D:\OTB Reports\Rate Code OTB This Year - Mansion\Reference Data OE Mansion.xlsx"),
}

RAW_DATA_FOLDERS = {
    63638: Path(r"D:\OTB Reports\Rate Code OTB This Year - The Villages"),  # The Villages
    23929: Path(r"D:\OTB Reports\Rate Code OTB This Year - Mansion"),        # Mansion Hotel
}

PROPERTY_NAMES = {
    63638: "The Villages",
    23929: "Mansion Hotel",
}

REFERENCE_SHEETS = {
    "dim_property":       ("Property Info",        3),
    "dim_room_type":      ("Room Types",            2),
    "dim_rate_code":      ("Rate Codes",             2),
    "dim_market_segment": ("Market Segments",        2),
    "dim_comp_set":       ("Comp Set Properties",    3),
    "dim_date":           ("Date",                   3),
    "event_calendar":     ("EventCalendar",          3),
    "semantic_naming":    ("Semantic Naming",        3),
}

# Industry-standard tolerance for hospitality revenue reconciliation is +/-3%,
# per Brian -- source system rounding/governance practices routinely produce
# small discrepancies this size without indicating a real data problem.
RECONCILIATION_TOLERANCE_PCT = 0.03


# ============================================================
# STEP 1: REFERENCE DATA
# ============================================================

def load_reference_data(path: Path) -> dict:
    """
    Loads all dimension tables and the semantic naming alias table from a
    property's reference workbook into a dict of DataFrames.

    Drops non-data rows that live inside the Excel Table range but aren't
    real dimension rows -- summary/total rows (e.g. Room Types' "Total" row),
    identified by a blank value in the sheet's first column, which is always
    populated for a genuine data row.
    """
    ref = {}
    for key, (sheet_name, header_row) in REFERENCE_SHEETS.items():
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
            df = df.dropna(how="all")
            df = df[df[df.columns[0]].notna()]
            ref[key] = df.reset_index(drop=True)
        except Exception as e:
            print(f"WARNING: could not load sheet '{sheet_name}' from {path.name} -> {e}")
            ref[key] = pd.DataFrame()
    return ref


def build_alias_map(semantic_naming_df: pd.DataFrame) -> dict:
    """
    Builds a {source_alias: semantic_model_name} lookup from the Semantic
    Naming tab, used to rename raw source columns into governed names.

    Source Alias is unique per row; the same Semantic Model Name can
    legitimately repeat (that's the intended design) -- so this maps
    alias -> semantic name, not the reverse.
    """
    alias_map = {}
    for _, row in semantic_naming_df.iterrows():
        semantic_name = row.get("Semantic Model Name")
        source_alias = row.get("Source Alias")
        if pd.notna(semantic_name) and pd.notna(source_alias):
            alias_map[str(source_alias).strip()] = str(semantic_name).strip()
    return alias_map


# ============================================================
# STEP 2: FILE DISCOVERY & READERS
# ============================================================

def discover_files(folder: Path) -> list:
    """Lists every standardized OTB file in a folder, sorted oldest to newest."""
    if not folder.exists():
        raise FileNotFoundError(f"Folder does not exist: {folder}")
    pattern = re.compile(r"^(\d{4}-\d{2}-\d{2})_Rate_Code_OTB_.+\.csv$")
    results = []
    for f in sorted(folder.glob("*.csv")):
        m = pattern.match(f.name)
        if m:
            snapshot_date = datetime.strptime(m.group(1), "%Y-%m-%d").date()
            results.append((f, snapshot_date))
        else:
            print(f"  SKIPPED (unrecognized filename): {f.name}")
    return sorted(results, key=lambda x: x[1])


def read_otb_file(filepath: Path, snapshot_date, property_id: int) -> pd.DataFrame:
    """
    Reads one raw Rate Code OTB CSV, skipping the 2 non-tabular rows up top.
    Returns None (rather than raising) for an empty/unreadable file -- a
    single bad file should not crash the whole daily run.
    """
    try:
        df = pd.read_csv(filepath, skiprows=2, header=0)
    except pd.errors.EmptyDataError:
        print(f"  SKIPPED (empty file): {filepath.name}")
        return None
    except Exception as e:
        print(f"  SKIPPED (unreadable: {e}): {filepath.name}")
        return None
    if df.empty:
        print(f"  SKIPPED (no data rows): {filepath.name}")
        return None
    df["PropertyID"] = property_id
    df["SnapshotDate"] = snapshot_date
    return df


def standardize_columns(df: pd.DataFrame, alias_map: dict) -> pd.DataFrame:
    """Renames raw source columns to governed Semantic Model Names."""
    rename_dict = {}
    unmapped = []
    for col in df.columns:
        if col in ("PropertyID", "SnapshotDate"):
            continue
        if col in alias_map:
            rename_dict[col] = alias_map[col]
        else:
            unmapped.append(col)
    if unmapped:
        print(f"  NOTE: {len(unmapped)} column(s) not found in the alias map: {unmapped}")
    return df.rename(columns=rename_dict)


def ingest_all_files(folder: Path, property_id: int, alias_map: dict) -> pd.DataFrame:
    """Discovers, reads, and standardizes every file in a folder for one property."""
    files = discover_files(folder)
    if not files:
        print(f"  No standardized files found in {folder}")
        return pd.DataFrame()
    frames = []
    skipped_count = 0
    for filepath, snapshot_date in files:
        raw = read_otb_file(filepath, snapshot_date, property_id)
        if raw is None:
            skipped_count += 1
            continue
        frames.append(standardize_columns(raw, alias_map))
    if not frames:
        print(f"  No readable files found in {folder}")
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    print(f"  Ingested {len(frames)} file(s), skipped {skipped_count}, {len(combined)} total rows, from {folder.name}")
    return combined


# ============================================================
# STEP 3: CLEANING
# ============================================================

def init_log() -> list:
    return []


def log_event(log: list, rule: str, action: str, detail: str, row_id=None):
    log.append({"Timestamp": datetime.now(), "Rule": rule, "Action": action, "Detail": detail, "RowIdentifier": row_id})


def get_log_df(log: list) -> pd.DataFrame:
    return pd.DataFrame(log)


def extract_and_remove_totals(df: pd.DataFrame, log: list) -> tuple:
    """
    Pulls out the source's own "Total" row(s) -- and the legacy "'-'" total
    marker -- into a separate dataframe, kept for reconciliation rather than
    discarded. Removed from the main dataframe since it's a different grain
    than every other row.
    """
    is_total = (
        (df["Stay Date"] == "Total")
        | (df["Rate Code"] == "Total")
        | (df["Rate Code"] == "'-'")
    )
    totals_df = df[is_total].copy()
    cleaned = df[~is_total].copy()
    if not totals_df.empty:
        totals_df["_marker"] = totals_df["Rate Code"].apply(lambda x: "Total" if x == "Total" else "'-' (legacy total marker)")
        summary = totals_df.groupby(["SnapshotDate", "PropertyID", "_marker"]).size()
        for (snap, prop, marker), count in summary.items():
            log_event(log, "extract_totals", "removed_from_main",
                       f"{marker}: {count} row(s) for SnapshotDate={snap}, PropertyID={prop} set aside for later reconciliation")
        totals_df = totals_df.drop(columns=["_marker"])
    return cleaned, totals_df


def normalize_dates(df: pd.DataFrame, log: list) -> pd.DataFrame:
    """Converts 'Stay Date' text (e.g. "Thu Jan 1, 2026") into a real date."""
    df = df.copy()

    def parse_date(val):
        try:
            return datetime.strptime(val, "%a %b %d, %Y").date()
        except (ValueError, TypeError):
            return pd.NaT

    df["StayDate"] = df["Stay Date"].apply(parse_date)
    failed = df["StayDate"].isna().sum()
    if failed:
        log_event(log, "normalize_dates", "flagged", f"{failed} row(s) failed to parse Stay Date")
    return df.drop(columns=["Stay Date"])


def deduplicate(df: pd.DataFrame, log: list) -> pd.DataFrame:
    """Drops rows sharing the same (PropertyID, SnapshotDate, StayDate, RateCode)."""
    key_cols = ["PropertyID", "SnapshotDate", "StayDate", "Rate Code"]
    before = len(df)
    df = df.drop_duplicates(subset=key_cols, keep="first")
    removed = before - len(df)
    if removed:
        log_event(log, "deduplicate", "removed", f"{removed} duplicate row(s) removed")
    else:
        log_event(log, "deduplicate", "none_found", "No duplicates found")
    return df


def validate_rate_codes(df: pd.DataFrame, dim_rate_code: pd.DataFrame, log: list) -> pd.DataFrame:
    """
    Flags any RateCode not present in the reference mapping table. There is
    no such thing as a truly "invalid" rate code here -- if one shows up, it
    means the master Rate Codes mapping table needs a new entry. Kept in the
    data, not dropped.
    """
    df = df.copy()
    known_codes = set(dim_rate_code["Rate Code"].astype(str))
    df["RateCode_KnownInMapping"] = df["Rate Code"].astype(str).isin(known_codes)
    unknown = df.loc[~df["RateCode_KnownInMapping"], "Rate Code"].unique()
    for code in unknown:
        count = (df["Rate Code"] == code).sum()
        log_event(log, "validate_rate_codes", "needs_mapping_table_update",
                   f"Rate Code '{code}' not found in master mapping table ({count} row(s)) -- add to Reference Data, not a data error")
    return df


def check_nulls(df: pd.DataFrame, critical_cols: list, log: list) -> pd.DataFrame:
    """Flags missing values in critical columns. Does NOT fill them."""
    for col in critical_cols:
        if col not in df.columns:
            continue
        n_null = df[col].isna().sum()
        if n_null:
            log_event(log, "check_nulls", "flagged", f"{n_null} null value(s) found in '{col}'")
    return df


def check_historical_consistency(df: pd.DataFrame, as_of_date: date, log: list) -> pd.DataFrame:
    """
    Past StayDates should never change once they've happened. Flags any
    StayDate < as_of_date where values differ across snapshots.
    """
    past = df[df["StayDate"] < as_of_date]
    if past.empty:
        log_event(log, "historical_consistency", "skipped", "No past StayDates in this data yet")
        return df
    check_cols = [c for c in ["Rooms Sold", "ADR", "Revenue"] if c in df.columns]
    grouped = past.groupby(["PropertyID", "StayDate", "Rate Code"])[check_cols].nunique()
    inconsistent = grouped[(grouped > 1).any(axis=1)]
    if inconsistent.empty:
        log_event(log, "historical_consistency", "passed",
                   f"All past StayDates consistent across snapshots ({past['SnapshotDate'].nunique()} snapshot(s) checked)")
    else:
        for idx in inconsistent.index:
            log_event(log, "historical_consistency", "MISMATCH",
                       f"PropertyID/StayDate/RateCode {idx} has different values across snapshots -- investigate")
    return df


def clean_otb_data(raw_df: pd.DataFrame, dim_rate_code: pd.DataFrame, as_of_date: date):
    """Runs the full cleaning sequence. Returns (cleaned_df, totals_df, log_df)."""
    log = init_log()
    df, totals_df = extract_and_remove_totals(raw_df, log)
    df = normalize_dates(df, log)
    df = deduplicate(df, log)
    df = validate_rate_codes(df, dim_rate_code, log)
    df = check_nulls(df, critical_cols=["Rooms Sold", "Rate Code", "StayDate"], log=log)
    df = check_historical_consistency(df, as_of_date, log)
    return df, totals_df, get_log_df(log)


# ============================================================
# STEP 4: BUSINESS LOGIC
# ============================================================

def calculate_pickup(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adds a Pickup column: RoomsSold minus RoomsSold from the immediately
    prior snapshot for the same StayDate/RateCode. NaN (not 0) when there's
    no prior snapshot to compare to -- "no data yet" and "no change" are
    different things.
    """
    df = df.sort_values(["PropertyID", "StayDate", "Rate Code", "SnapshotDate"]).copy()
    df["Pickup"] = df.groupby(["PropertyID", "StayDate", "Rate Code"])["Rooms Sold"].diff()
    return df


def safe_pct_change(current: pd.Series, base: pd.Series) -> pd.Series:
    """
    (current - base) / base, with explicit zero-base handling:
      base == 0 and current == 0  -> 0.0 (no change, not undefined)
      base == 0 and current != 0  -> NaN (percent change not meaningful from zero)
    """
    diff = current - base
    with np.errstate(divide="ignore", invalid="ignore"):
        pct = np.where(base == 0, np.where(current == 0, 0.0, np.nan), diff / base)
    return pd.Series(pct, index=current.index)


def calculate_pace_variance(df: pd.DataFrame) -> pd.DataFrame:
    """
    Adds Diff and % Change columns for both STLY alignments (DOW and Date),
    for Rooms Sold, ADR, and Revenue -- recomputed from the STLY value
    columns already present, not copied from the source's own Diff/%Change.
    """
    df = df.copy()
    metrics = ["Rooms Sold", "ADR", "Revenue"]
    alignments = {"DOW-Aligned": "STLY (DOW-Aligned)", "Date-Aligned": "STLY (Date-Aligned)"}
    for metric in metrics:
        for _, stly_suffix in alignments.items():
            stly_col = f"{metric} {stly_suffix}"
            if stly_col not in df.columns:
                continue
            df[f"{metric} Variance vs {stly_suffix}"] = df[metric] - df[stly_col]
            df[f"{metric} % Variance vs {stly_suffix}"] = safe_pct_change(df[metric], df[stly_col])
    return df


def monthly_rollup(df: pd.DataFrame, snapshot_date, dim_property: pd.DataFrame) -> pd.DataFrame:
    """
    Rolls up to PropertyID/Year/Month for one snapshot's worth of data.
    Adds Occupancy % and RevPAR, matching the formulas confirmed from the
    original 365-Day Pickup report:
      Occupancy % = Room Nights / (Total Rooms * Days in Period)
      RevPAR      = Revenue / (Total Rooms * Days in Period)
    Total Rooms comes from dim_property, not hardcoded.
    """
    snap = df[df["SnapshotDate"] == snapshot_date].copy()
    snap["Year"] = pd.to_datetime(snap["StayDate"]).dt.year
    snap["Month"] = pd.to_datetime(snap["StayDate"]).dt.month

    rollup = snap.groupby(["PropertyID", "Year", "Month"]).agg(
        RoomNights=("Rooms Sold", "sum"),
        Revenue=("Revenue", "sum"),
    ).reset_index()
    rollup["ADR"] = np.where(rollup["RoomNights"] == 0, 0.0, rollup["Revenue"] / rollup["RoomNights"])

    room_count_lookup = dim_property.set_index("PropertyID")["TotalRoomCount"].to_dict()
    rollup["TotalRooms"] = rollup["PropertyID"].map(room_count_lookup)
    rollup["DaysInMonth"] = rollup.apply(lambda r: calendar.monthrange(int(r["Year"]), int(r["Month"]))[1], axis=1)
    rollup["AvailableRoomNights"] = rollup["TotalRooms"] * rollup["DaysInMonth"]
    rollup["Occupancy %"] = np.where(rollup["AvailableRoomNights"] == 0, 0.0, rollup["RoomNights"] / rollup["AvailableRoomNights"])
    rollup["RevPAR"] = np.where(rollup["AvailableRoomNights"] == 0, 0.0, rollup["Revenue"] / rollup["AvailableRoomNights"])
    return rollup


def market_segment_rollup(df: pd.DataFrame, dim_rate_code: pd.DataFrame, snapshot_date) -> pd.DataFrame:
    """Rolls up to PropertyID/StayDate/MarketSegment, via Rate Code's segment attribute."""
    snap = df[df["SnapshotDate"] == snapshot_date].copy()
    rate_to_segment = dim_rate_code.set_index("Rate Code")["Market Segment"].to_dict()
    snap["Market Segment"] = snap["Rate Code"].map(rate_to_segment)
    unmapped = snap["Market Segment"].isna().sum()
    if unmapped:
        print(f"  NOTE: {unmapped} row(s) have a Rate Code not found in dim_rate_code -- Market Segment is blank for these")
    return snap.groupby(["PropertyID", "StayDate", "Market Segment"]).agg(
        RoomsSold=("Rooms Sold", "sum"),
        Revenue=("Revenue", "sum"),
    ).reset_index()


# ============================================================
# STEP 5: VALIDATION
# ============================================================

def reconcile_totals(clean_df: pd.DataFrame, totals_df: pd.DataFrame) -> pd.DataFrame:
    """
    For each (PropertyID, SnapshotDate), compares our own summed Rooms
    Sold/Revenue against the source file's own "Total" row. Flags a genuine
    mismatch only if the percent difference exceeds the industry-standard
    +/-3% tolerance.
    """
    our_sums = clean_df.groupby(["PropertyID", "SnapshotDate"]).agg(
        OurRoomsSold=("Rooms Sold", "sum"),
        OurRevenue=("Revenue", "sum"),
    ).reset_index()

    source_totals = totals_df[totals_df["Rate Code"] == "Total"][
        ["PropertyID", "SnapshotDate", "Rooms Sold", "Revenue"]
    ].rename(columns={"Rooms Sold": "SourceRoomsSold", "Revenue": "SourceRevenue"})

    comparison = our_sums.merge(source_totals, on=["PropertyID", "SnapshotDate"], how="outer")
    comparison["RoomsSold_Diff"] = comparison["OurRoomsSold"] - comparison["SourceRoomsSold"]
    comparison["Revenue_Diff"] = comparison["OurRevenue"] - comparison["SourceRevenue"]
    comparison["RoomsSold_Diff_Pct"] = safe_pct_change(comparison["OurRoomsSold"], comparison["SourceRoomsSold"]).abs()
    comparison["Revenue_Diff_Pct"] = safe_pct_change(comparison["OurRevenue"], comparison["SourceRevenue"]).abs()
    comparison["WithinTolerance"] = (
        (comparison["RoomsSold_Diff_Pct"] <= RECONCILIATION_TOLERANCE_PCT) &
        (comparison["Revenue_Diff_Pct"] <= RECONCILIATION_TOLERANCE_PCT)
    )
    return comparison


# ============================================================
# ORCHESTRATION -- runs the full pipeline for one property
# ============================================================

def run_pipeline_for_property(property_id: int) -> dict:
    """
    Runs Steps 1-5 end to end for a single property. Returns a dict of every
    intermediate and final dataframe, keyed by name, so the caller can
    inspect any stage (e.g. result["clean"], result["monthly"]).
    """
    name = PROPERTY_NAMES[property_id]
    print(f"\n{'='*60}\n{name} (PropertyID {property_id})\n{'='*60}")

    reference = load_reference_data(REFERENCE_DATA_PATHS[property_id])
    alias_map = build_alias_map(reference["semantic_naming"])
    print(f"Reference data + alias map loaded ({len(alias_map)} alias entries)")

    raw = ingest_all_files(RAW_DATA_FOLDERS[property_id], property_id, alias_map)
    clean, totals, log = clean_otb_data(raw, reference["dim_rate_code"], as_of_date=date.today())

    pickup = calculate_pickup(clean)
    variance = calculate_pace_variance(pickup)

    latest_snapshot = variance["SnapshotDate"].max()
    monthly = monthly_rollup(variance, snapshot_date=latest_snapshot, dim_property=reference["dim_property"])
    segment = market_segment_rollup(variance, reference["dim_rate_code"], snapshot_date=latest_snapshot)

    reconciliation = reconcile_totals(clean, totals)
    reconciliation_mismatches = reconciliation[~reconciliation["WithinTolerance"]]

    print(f"Cleaned: {len(clean)} rows across {clean['SnapshotDate'].nunique()} snapshot(s)")
    print(f"Reconciliation: {len(reconciliation)} snapshot(s) checked, "
          f"{len(reconciliation_mismatches)} outside +/-{RECONCILIATION_TOLERANCE_PCT:.0%} tolerance")

    return {
        "reference": reference,
        "raw": raw,
        "clean": clean,
        "totals": totals,
        "log": log,
        "pickup": pickup,
        "variance": variance,
        "monthly": monthly,
        "segment": segment,
        "reconciliation": reconciliation,
    }


def main():
    results = {}
    for property_id in RAW_DATA_FOLDERS:
        results[property_id] = run_pipeline_for_property(property_id)

    print(f"\n{'='*60}\nDone. Results available for: {[PROPERTY_NAMES[p] for p in results]}\n{'='*60}")
    return results


if __name__ == "__main__":
    all_results = main()
