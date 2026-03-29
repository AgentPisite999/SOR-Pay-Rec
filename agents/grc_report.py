import os
import re
import json
import base64
import argparse
import sys
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# ===================== FIX: FORCE UTF-8 PRINTING =====================
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ===================== DEFAULTS =====================
DEFAULT_OUTPUT_NAME = "GRC_OUTPUT.xlsx"
PROCESSED_SHEET_NAME = "GRC_Data"
PIVOT_SHEET_NAME = "GRC_Pivot"

# ===================== VENDOR NAME 2 MAPPING =====================
VENDOR_KEYWORDS = [
    ("myntra",                  "Myntra jabong"),
    ("jabong",                  "Myntra jabong"),
    ("ajio sor",                "Reliance Retail Ajio SOR"),
    ("ajio",                    "Reliance Retail Ajio SOR"),
    ("centro",                  "Relience Centro"),
    ("shoppers stop",           "SHOPPERS STOP"),
    ("shoppersstop",            "SHOPPERS STOP"),
    ("v retail",                "V - Retail"),
    ("leayan",                  "ZUUP"),
    ("zuup",                    "ZUUP"),
    ("lifestyle international", "LS"),
    ("kora",                    "Kora"),
    ("sabharwal nirankar",      "NIRANKAR"),
    ("nirankar",                "NIRANKAR"),
    ("flipkart",                "Flipkart"),
]

MONTH_MAP = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR",
    5: "MAY", 6: "JUN", 7: "JUL", 8: "AUG",
    9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC",
}

# ===================== CLI =====================
def parse_args():
    ap = argparse.ArgumentParser(description="GRC Purchase Register processor")
    ap.add_argument("--input",       required=True, help="Path to input Excel file")
    ap.add_argument("--output_dir",  required=True, help="Folder to write output file")
    ap.add_argument("--output_name", default=DEFAULT_OUTPUT_NAME, help="Output Excel file name")
    return ap.parse_args()


# ===================== HELPERS =====================
def canon(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def build_canon_map(columns):
    m = {}
    for col in columns:
        m[canon(col)] = col
    return m


def to_num(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="float64")
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    cleaned = (
        s.astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    cleaned = cleaned.replace({"": None, "nan": None, "None": None, "NaN": None})
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def normalize_ean(x) -> str:
    if x is None:
        return ""
    s = str(x).replace("\u00A0", " ").strip()
    s = re.sub(r"\.0$", "", s)
    if re.match(r"^\d+(\.\d+)?e\+\d+$", s.lower()):
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    return s


def env_clean(v: str, default: str = "") -> str:
    if v is None:
        return default
    s = str(v).strip().strip('"').strip("'").strip()
    if s.endswith(";"):
        s = s[:-1].strip()
    return s or default


def map_vendor_name2(vendor: str) -> str:
    v = str(vendor).lower().replace("_", " ").replace("-", " ")
    v = " ".join(v.split())
    for keyword, mapped in VENDOR_KEYWORDS:
        if keyword in v:
            return mapped
    return "Others"


# ===================== MONTH/YEAR EXTRACTION =====================
def extract_month_year_from_file(filepath: Path):
    """
    Reads first few rows to find date range line.
    E.g. 'Date Range: 01-02-2026 - 28-02-2026' -> ('FEB', '2026')
    """
    try:
        preview = pd.read_excel(filepath, header=None, nrows=5)
        for _, row in preview.iterrows():
            for cell in row:
                cell_str = str(cell)
                m = re.search(r"(\d{2})-(\d{2})-(\d{4})", cell_str)
                if m:
                    month_num = int(m.group(2))
                    year = m.group(3)
                    month = MONTH_MAP.get(month_num, "")
                    if month and year:
                        return month, year
    except Exception as e:
        print(f"Warning: could not extract month/year from file content: {e}")

    # Fallback: try filename
    fname = filepath.stem
    m = re.search(r"(\d{2})[_\-](\d{4})", fname)
    if m:
        month_num = int(m.group(1))
        year = m.group(2)
        month = MONTH_MAP.get(month_num, "")
        if month:
            return month, year

    return "", ""


# ===================== READ INPUT FILE =====================
def read_grc_file(filepath: Path) -> pd.DataFrame:
    """
    GRC Excel structure:
      Row 1: Title        e.g. 'Purchase Register'
      Row 2: Date range   e.g. 'Date Range: 01-02-2026 - 28-02-2026'
      Row 3: Headers
      Row 4+: Data
    header=2 (0-indexed) skips rows 0 and 1.
    """
    df = pd.read_excel(filepath, header=2)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


# ===================== GOOGLE SHEETS =====================
def get_gspread_client(write: bool = False):
    load_dotenv(override=True)
    b64 = os.getenv("GOOGLE_SA_JSON_B64")
    if not b64:
        raise RuntimeError("Missing GOOGLE_SA_JSON_B64 in .env")
    sa_json = json.loads(base64.b64decode(b64).decode("utf-8"))
    scopes = (
        ["https://www.googleapis.com/auth/spreadsheets"]
        if write
        else ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    creds = Credentials.from_service_account_info(sa_json, scopes=scopes)
    return gspread.authorize(creds)


def worksheet_to_df(ws) -> pd.DataFrame:
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    headers = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ===================== LOAD COGS MASTER =====================
def load_cogs_master():
    """
    Loads COGS Master from same Google Sheet used in rec_TD.
    Columns: EANCODE + Rate
    Used for COGS Rate lookup by BARCODE.
    """
    load_dotenv(override=True)

    sheet_id = env_clean(os.getenv("MARGIN_SPREADSHEET_ID2", ""))
    cogs_tab = env_clean(os.getenv("COGS_SHEET_NAME2", "COGS Master"))

    if not sheet_id:
        raise RuntimeError("Missing MARGIN_SPREADSHEET_ID2 in .env")

    gc = get_gspread_client(write=False)
    sh = gc.open_by_key(sheet_id)

    ws_c = sh.worksheet(cogs_tab)
    df_c = worksheet_to_df(ws_c)
    if df_c.empty:
        raise RuntimeError(f"No data found in sheet: {cogs_tab}")

    c_map = build_canon_map(df_c.columns)
    if "eancode" not in c_map or "rate" not in c_map:
        raise RuntimeError("COGS Master must have columns: EANCODE, Rate")

    e_col = c_map["eancode"]
    r_col = c_map["rate"]

    cogs_map = {}
    for _, row in df_c.iterrows():
        ean = normalize_ean(row.get(e_col, ""))
        if not ean:
            continue
        rate = pd.to_numeric(
            str(row.get(r_col, "")).replace(",", ""), errors="coerce"
        )
        cogs_map[ean] = float(rate) if pd.notna(rate) else 0.0

    print(f"Loaded COGS Master: {len(cogs_map)} records from '{cogs_tab}'")
    return cogs_map


# ===================== PROCESS DATA =====================
def process_grc(df: pd.DataFrame, cogs_map: dict, month: str, year: str) -> pd.DataFrame:
    df = df.copy()
    col_map = build_canon_map(df.columns)

    # ---- Vendor Name 2 ----
    # GRC file has two cols both named "Vendor Name"
    # pandas renames second to "Vendor Name.1" automatically
    # We use the first one (full raw vendor name) for mapping
    vendor_col = col_map.get("vendor name")
    if vendor_col:
        df["Vendor Name 2"] = df[vendor_col].fillna("").apply(map_vendor_name2)
    else:
        df["Vendor Name 2"] = "Others"

    # ---- Qty ----
    qty_col = col_map.get("purchase qty & purchase return qty")
    qty = to_num(df[qty_col]) if qty_col else pd.Series(0.0, index=df.index)

    # ---- Gross Amount ----
    gross_col = col_map.get("gross amount")
    gross_amount = to_num(df[gross_col]) if gross_col else pd.Series(0.0, index=df.index)

    # ==========================================
    # MRP Rate  = Gross Amount / Qty
    # MRP Value = MRP Rate x Qty  (= Gross Amount)
    # ==========================================
    mrp_rate = pd.Series(0.0, index=df.index)
    nz = qty != 0
    mrp_rate[nz] = (gross_amount[nz] / qty[nz]).round(2)
    mrp_value = (mrp_rate * qty).round(2)

    df["MRP Rate"]  = mrp_rate
    df["MRP Value"] = mrp_value

    # ==========================================
    # COGS Rate  = lookup BARCODE in COGS Master
    # COGS Value = COGS Rate x Qty
    # ==========================================
    barcode_col = col_map.get("barcode")
    if barcode_col:
        barcode_series = df[barcode_col].map(normalize_ean)
    else:
        barcode_series = pd.Series("", index=df.index)

    cogs_rate_raw = barcode_series.map(lambda e: cogs_map.get(e, None))

    # Display: rate value or "not found in master"
    cogs_rate_display = cogs_rate_raw.where(
        cogs_rate_raw.notna(), other="not found in master"
    )

    cogs_rate_num = pd.to_numeric(
        pd.Series(cogs_rate_raw).astype(str).str.replace(",", "", regex=False),
        errors="coerce",
    ).fillna(0)

    cogs_value = (cogs_rate_num * qty).round(2)

    df["COGS Rate"]  = cogs_rate_display
    df["COGS Value"] = cogs_value

    # ---- Insert Month & Year at front ----
    df.insert(0, "Year",  year)
    df.insert(1, "Month", month)

    return df


# ===================== BUILD PIVOT =====================
def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pivot grouped by: Year, Month, Vendor Name 2
    Sums: Purchase Qty, NET_AMOUNT, Total GST, Gross Amount, COGS Value, MRP Value
    """
    col_map = build_canon_map(df.columns)

    qty_col   = col_map.get("purchase qty & purchase return qty")
    net_col   = col_map.get("net_amount")
    gst_col   = col_map.get("total gst")
    gross_col = col_map.get("gross amount")

    group_cols = ["Year", "Month", "Vendor Name 2"]

    pivot_df = df.copy()

    # Convert to numeric for aggregation
    for c in [qty_col, net_col, gst_col, gross_col]:
        if c:
            pivot_df[c] = pd.to_numeric(pivot_df[c], errors="coerce").fillna(0)

    pivot_df["COGS Value"] = pd.to_numeric(pivot_df["COGS Value"], errors="coerce").fillna(0)
    pivot_df["MRP Value"]  = pd.to_numeric(pivot_df["MRP Value"],  errors="coerce").fillna(0)

    agg = {}
    if qty_col:   agg[qty_col]   = "sum"
    if net_col:   agg[net_col]   = "sum"
    if gst_col:   agg[gst_col]   = "sum"
    if gross_col: agg[gross_col] = "sum"
    agg["COGS Value"] = "sum"
    agg["MRP Value"]  = "sum"

    pivot = pivot_df.groupby(group_cols, as_index=False).agg(agg)

    # Rename to expected pivot headers
    rename_map = {"Vendor Name 2": "Row Labels"}
    if qty_col:   rename_map[qty_col]   = "Sum of Purchase Qty & Purchase Return Qty"
    if net_col:   rename_map[net_col]   = "Sum of NET_AMOUNT"
    if gst_col:   rename_map[gst_col]   = "Sum of Total GST"
    if gross_col: rename_map[gross_col] = "Sum of Gross Amount"
    rename_map["COGS Value"] = "Sum of COGS Value"
    rename_map["MRP Value"]  = "Sum of MRP Value"

    pivot = pivot.rename(columns=rename_map)

    # Round floats
    for c in pivot.columns:
        if pivot[c].dtype == float:
            pivot[c] = pivot[c].round(2)

    return pivot


# ===================== PUSH PIVOT TO GOOGLE SHEETS =====================
def _normalize_for_compare(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Year"       in df.columns: df["Year"]       = df["Year"].astype(str).str.strip()
    if "Month"      in df.columns: df["Month"]      = df["Month"].astype(str).str.strip().str.upper()
    if "Row Labels" in df.columns: df["Row Labels"] = df["Row Labels"].astype(str).str.strip()
    return df


def _align_columns(existing_df: pd.DataFrame, df_new: pd.DataFrame):
    for c in existing_df.columns:
        if c not in df_new.columns:
            df_new[c] = ""
    for c in df_new.columns:
        if c not in existing_df.columns:
            existing_df[c] = ""
    existing_cols = list(existing_df.columns)
    extra_cols = [c for c in df_new.columns if c not in existing_cols]
    final_cols = existing_cols + extra_cols
    return existing_df[final_cols], df_new[final_cols]


def push_pivot_to_sheets(pivot_df: pd.DataFrame, month: str, year: str):
    """
    Upsert pivot to Google Sheet.
    Key: Year + Month + Row Labels
    Removes matching rows, appends new, keeps other months untouched.
    """
    load_dotenv(override=True)
    sheet_id = env_clean(os.getenv("GRC_SHEET_ID", ""))
    tab_name = env_clean(os.getenv("GRC_TAB_NAME", "GRC"))

    if not sheet_id:
        print("GRC_SHEET_ID missing in .env — skipping Google Sheets push.")
        return

    gc = get_gspread_client(write=True)
    sh = gc.open_by_key(sheet_id)

    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=tab_name,
            rows=1000,
            cols=max(10, pivot_df.shape[1] + 2),
        )

    existing_vals = ws.get_all_values()
    if not existing_vals:
        existing_df = pd.DataFrame()
    else:
        headers = existing_vals[0]
        rows    = existing_vals[1:]
        existing_df = pd.DataFrame(rows, columns=headers)
        existing_df.columns = [str(c).strip() for c in existing_df.columns]

    df_new = pivot_df.copy().replace([pd.NA, float("inf"), float("-inf")], "")
    df_new = df_new.where(pd.notnull(df_new), "")

    if existing_df.empty:
        out_df = df_new
    else:
        existing_df, df_new = _align_columns(existing_df, df_new)
        existing_df = _normalize_for_compare(existing_df)
        df_new      = _normalize_for_compare(df_new)

        required = {"Year", "Month", "Row Labels"}
        if required.issubset(existing_df.columns) and required.issubset(df_new.columns):
            new_keys = set(zip(
                df_new["Year"].astype(str).str.strip(),
                df_new["Month"].astype(str).str.upper(),
                df_new["Row Labels"].astype(str).str.strip(),
            ))
            existing_keys = list(zip(
                existing_df["Year"].astype(str).str.strip(),
                existing_df["Month"].astype(str).str.upper(),
                existing_df["Row Labels"].astype(str).str.strip(),
            ))
            mask_same = [k in new_keys for k in existing_keys]
            existing_df = existing_df.loc[[not m for m in mask_same]].copy()

        out_df = pd.concat([existing_df, df_new], ignore_index=True)

    out_df = out_df.where(pd.notnull(out_df), "")
    values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

    ws.clear()
    ws.update(values, value_input_option="USER_ENTERED")
    print(f"Pivot pushed to Google Sheet tab '{tab_name}' ({len(df_new)} rows upserted).")


# ===================== MAIN =====================
def main():
    args = parse_args()

    input_path  = Path(args.input)
    output_dir  = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / (args.output_name or DEFAULT_OUTPUT_NAME)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # ---- Extract Month/Year from file ----
    month, year = extract_month_year_from_file(input_path)
    if month and year:
        print(f"PERIOD: {month}-{year}")
    else:
        print("Warning: Could not detect month/year from file.")

    # ---- Load COGS Master ----
    cogs_map = load_cogs_master()

    # ---- Read input ----
    print(f"Reading: {input_path.name}")
    df = read_grc_file(input_path)
    print(f"Input rows: {len(df)}")

    # ---- Process ----
    df_processed = process_grc(df, cogs_map, month, year)

    # ---- Build pivot ----
    pivot_df = build_pivot(df_processed)

    # ---- Write output Excel ----
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_processed.to_excel(writer, sheet_name=PROCESSED_SHEET_NAME, index=False)
        pivot_df.to_excel(writer, sheet_name=PIVOT_SHEET_NAME, index=False)

    print("")
    print("Done. Output Excel created at:")
    print(str(output_file))
    print(f"Processed rows : {len(df_processed)}")
    print(f"Pivot rows     : {len(pivot_df)}")

    # ---- Push pivot to Google Sheets ----
    push_pivot_to_sheets(pivot_df, month, year)


if __name__ == "__main__":
    main()
