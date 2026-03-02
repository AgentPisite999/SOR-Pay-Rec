# import os
# import json
# import base64
# import re
# import argparse
# from pathlib import Path

# import pandas as pd
# import numpy as np
# from dotenv import load_dotenv
# import gspread
# from google.oauth2.service_account import Credentials


# # ===================== OUTPUT SHEET NAMES =====================
# BASE_SHEET_NAME_OUT = "Pay_Stk_Base_Data"
# PIVOT_SHEET_NAME_OUT = "Pay_Stk_Final2"

# # How many top rows to scan for header row
# HEADER_SCAN_ROWS = 60

# # How many rows/cols to scan for the "Period" text
# PERIOD_SCAN_ROWS = 20
# PERIOD_SCAN_COLS = 10


# # ===================== COLUMN ALIASES (FIX Qty., etc.) =====================
# COLUMN_ALIASES = {
#     "Qty.": "Qty",
#     "QTY.": "Qty",
#     "QTY": "Qty",
#     "Quantity": "Qty",
#     "quantity": "Qty",
#     "Party name": "Party Name",
#     "Party name1": "Party Name1",
# }


# # ===================== CLI ARGS =====================
# def parse_args():
#     ap = argparse.ArgumentParser(description="Payable Stock processing")
#     ap.add_argument("--input", required=True, help="Path to input Excel file (.xlsx/.xls)")
#     ap.add_argument("--output_dir", required=True, help="Directory to write outputs")
#     ap.add_argument(
#         "--output_name",
#         default="Data_Processed_STK.xlsx",
#         help="Output Excel file name (default: Data_Processed_STK.xlsx)",
#     )
#     return ap.parse_args()


# def apply_column_aliases(df: pd.DataFrame) -> pd.DataFrame:
#     """Rename known variant headers to canonical names (e.g., Qty. -> Qty)."""
#     df = df.copy()
#     df.columns = [str(c).strip() for c in df.columns]
#     df.rename(columns=lambda c: COLUMN_ALIASES.get(c, c), inplace=True)
#     return df


# def ensure_party_name_fallback(df: pd.DataFrame) -> pd.DataFrame:
#     """
#     Ensure 'Party Name' exists and is filled:
#       - If Party Name missing but Party Name1 exists -> create Party Name from Party Name1
#       - If both exist -> fill blank Party Name using Party Name1
#     """
#     df = df.copy()

#     if "Party Name" not in df.columns and "Party Name1" in df.columns:
#         df["Party Name"] = df["Party Name1"]

#     if "Party Name" in df.columns and "Party Name1" in df.columns:
#         party = df["Party Name"].astype(str).replace("nan", "").str.strip()
#         party1 = df["Party Name1"].astype(str).replace("nan", "").str.strip()
#         df["Party Name"] = np.where(party.eq("") | party.isna(), party1, df["Party Name"])

#     return df


# # ===================== HELPERS =====================
# def _get_env_str(key: str) -> str:
#     val = os.getenv(key)
#     if val is None or str(val).strip() == "":
#         raise ValueError(f"Missing env var: {key}")
#     return val.strip()


# def normalize_colname(x: str) -> str:
#     """Normalize headers: lowercase, strip, collapse spaces, apply alias mapping."""
#     if x is None:
#         return ""
#     s = str(x).strip()
#     s = COLUMN_ALIASES.get(s, s)  # apply alias before normalizing
#     s = s.lower()
#     s = " ".join(s.split())
#     return s


# def find_header_row_in_sheet(excel_path: Path, sheet_name: str | int, required_cols: list[str]) -> int:
#     """
#     Reads the first N rows with header=None and finds which row looks like the header
#     by checking if it contains the required column names (normalized).
#     Returns the header row index (0-based).
#     """
#     preview = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=HEADER_SCAN_ROWS)

#     required_norm = {normalize_colname(c) for c in required_cols}

#     best_row = None
#     best_score = -1

#     for r in range(len(preview)):
#         row_vals = preview.iloc[r].tolist()
#         row_norm = {normalize_colname(v) for v in row_vals if normalize_colname(v)}

#         score = len(required_norm.intersection(row_norm))
#         if score > best_score:
#             best_score = score
#             best_row = r

#         if score == len(required_norm):
#             return r

#     min_needed = max(3, int(0.7 * len(required_norm)))  # 70% match
#     if best_score >= min_needed and best_row is not None:
#         return best_row

#     raise ValueError(
#         f"Could not detect header row in sheet '{sheet_name}'. "
#         f"Best match row index={best_row} with score={best_score}/{len(required_norm)}. "
#         f"Try increasing HEADER_SCAN_ROWS or check column names."
#     )


# def read_sheet_with_auto_header(excel_path: Path, required_cols: list[str]):
#     """
#     Picks the first sheet, finds header row automatically, reads full data from there.
#     Returns (df, sheet_name_used).
#     """
#     xls = pd.ExcelFile(excel_path)
#     sheet_to_read = xls.sheet_names[0]  # default: first sheet
#     print("Sheets found in input file:", xls.sheet_names)
#     print("Using sheet:", sheet_to_read)

#     header_row = find_header_row_in_sheet(excel_path, sheet_to_read, required_cols)
#     print(f"Detected header row (0-based): {header_row}  -> Excel row: {header_row + 1}")

#     df = pd.read_excel(excel_path, sheet_name=sheet_to_read, header=header_row)
#     df = df.dropna(how="all").copy()
#     df.columns = [str(c).strip() for c in df.columns]

#     # ✅ Fix column variants like Qty. and Party Name fallback
#     df = apply_column_aliases(df)
#     df = ensure_party_name_fallback(df)

#     return df, sheet_to_read


# # ===================== PERIOD (Month/Year) DETECTION =====================
# def extract_month_year_from_period_text(text: str):
#     """
#     Extract month/year from strings like:
#       "Period: 002. DEC-25 To 002. DEC-25"
#       "Period: DEC-25"
#       "Period: DEC 2025"
#     Returns (month_str, year_int) or (None, None)
#     """
#     if not text:
#         return None, None

#     s = str(text).upper()

#     m = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b\s*[-/ ]\s*(\d{2,4})", s)
#     if not m:
#         m = re.search(r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\s*[-/ ]\s*(\d{2,4})", s)

#     if not m:
#         return None, None

#     mon = m.group(1)
#     yy = m.group(2)

#     if mon == "SEPT":
#         mon = "SEP"

#     year = int(yy)
#     if len(yy) == 2:
#         year = 2000 + year  # assumes 20xx

#     return mon, year


# def find_month_year_in_sheet(excel_path: Path, sheet_name: str | int):
#     """
#     Scans the top-left area of the sheet to find a cell containing 'PERIOD'
#     (or any cell with DEC-25 like pattern) and returns (month, year).
#     """
#     preview = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=PERIOD_SCAN_ROWS)
#     preview = preview.iloc[:, :PERIOD_SCAN_COLS]

#     # 1) Prefer cells containing 'PERIOD'
#     for r in range(preview.shape[0]):
#         for c in range(preview.shape[1]):
#             val = preview.iat[r, c]
#             if pd.isna(val):
#                 continue
#             txt = str(val)
#             if "PERIOD" in txt.upper():
#                 mon, year = extract_month_year_from_period_text(txt)
#                 if mon and year:
#                     return mon, year

#     # 2) Fallback: any cell with month-year pattern
#     for r in range(preview.shape[0]):
#         for c in range(preview.shape[1]):
#             val = preview.iat[r, c]
#             if pd.isna(val):
#                 continue
#             mon, year = extract_month_year_from_period_text(str(val))
#             if mon and year:
#                 return mon, year

#     return None, None


# # ===================== GOOGLE SHEETS MAPPING =====================
# def load_mapping_from_gsheet() -> pd.DataFrame:
#     """
#     Reads the mapping table from Google Sheet using a Service Account JSON stored as base64 in .env.
#     Expects columns: Party Name, Brand, Party Name1
#     """
#     load_dotenv()

#     sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")
#     sheet_id = _get_env_str("GOOGLE_SHEET_ID")
#     tab_name = _get_env_str("GOOGLE_SHEET_TAB_PAYSTK").strip().strip('"').strip("'")

#     try:
#         sa_json_text = base64.b64decode(sa_b64).decode("utf-8")
#         sa_info = json.loads(sa_json_text)
#     except Exception as e:
#         raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e

#     scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
#     creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
#     gc = gspread.authorize(creds)

#     sh = gc.open_by_key(sheet_id)
#     ws = sh.worksheet(tab_name)

#     values = ws.get_all_values()
#     if not values or len(values) < 2:
#         raise ValueError(f"Google sheet tab '{tab_name}' seems empty or has no data rows.")

#     header = [h.strip() for h in values[0]]
#     rows = values[1:]
#     df_map = pd.DataFrame(rows, columns=header)

#     for c in ["Party Name", "Brand", "Party Name1"]:
#         if c in df_map.columns:
#             df_map[c] = df_map[c].astype(str).str.strip()

#     return df_map


# # ===================== YOUR ORIGINAL LOGIC =====================
# def add_prj_prefix_for_puma(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     for c in ["Branch Name", "Party Name", "Party Name1"]:
#         if c not in df.columns:
#             df[c] = np.nan

#     branch = df["Branch Name"].fillna("").astype(str).str.strip()
#     party = df["Party Name"].fillna("").astype(str)
#     party1 = df["Party Name1"].fillna("").astype(str)

#     mask = (
#         branch.str.upper().str.startswith("PRJ")
#         & (
#             party.str.upper().str.contains("PUMA", na=False)
#             | party1.str.upper().str.contains("PUMA", na=False)
#         )
#     )

#     def prefix_prj(series: pd.Series) -> pd.Series:
#         s = series.fillna("").astype(str)
#         already = s.str.upper().str.startswith("PRJ-")
#         return np.where(already, s, "PRJ-" + s)

#     df.loc[mask, "Party Name"] = prefix_prj(df.loc[mask, "Party Name"])
#     df.loc[mask, "Party Name1"] = prefix_prj(df.loc[mask, "Party Name1"])
#     return df


# def add_calculated_columns(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     df["Cost Rate"] = pd.to_numeric(df["Cost Rate"], errors="coerce")
#     df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")

#     df["Total Cost"] = df["Cost Rate"] * df["Qty"]
#     df["Gst%"] = np.where(df["Cost Rate"] <= 2500, 0.05, 0.18)
#     df["Gst"] = df["Total Cost"] * df["Gst%"]
#     df["Actual Cost"] = df["Total Cost"] + df["Gst"]
#     return df


# def create_pay_stk_pivot(df: pd.DataFrame, month: str | None, year: int | None) -> pd.DataFrame:
#     df = df.copy()
#     for col in ["Qty", "Total Cost", "Gst", "Actual Cost"]:
#         df[col] = pd.to_numeric(df[col], errors="coerce")

#     pivot = pd.pivot_table(
#         df,
#         index=["SOR/ Outright", "Brand/Inhouse", "Party Name1"],
#         values=["Qty", "Total Cost", "Gst", "Actual Cost"],
#         aggfunc="sum",
#         fill_value=0,
#         margins=False,
#     ).reset_index()

#     pivot.insert(0, "Year", year)
#     pivot.insert(0, "Month", month)

#     return pivot


# def main():
#     args = parse_args()

#     input_path = Path(args.input)
#     if not input_path.exists():
#         raise FileNotFoundError(f"Input file not found: {input_path}")

#     out_dir = Path(args.output_dir)
#     out_dir.mkdir(parents=True, exist_ok=True)
#     output_file = out_dir / args.output_name

#     # ✅ Party Name can be missing; we'll fallback from Party Name1 after reading
#     required_base_cols = [
#         "Brand",
#         "Cost Rate",
#         "Qty",  # Qty. will be normalized to Qty
#         "SOR/ Outright",
#         "Brand/Inhouse",
#         "Branch Name",
#     ]

#     df_base, sheet_used = read_sheet_with_auto_header(input_path, required_base_cols)

#     # ✅ Must have at least one
#     if "Party Name" not in df_base.columns and "Party Name1" not in df_base.columns:
#         raise ValueError("Missing required column: need either 'Party Name' or 'Party Name1' in input.")

#     # ✅ Ensure Party Name exists for mapping merge
#     df_base = ensure_party_name_fallback(df_base)

#     missing_base = [c for c in required_base_cols if c not in df_base.columns]
#     if missing_base:
#         raise ValueError(
#             f"Missing required columns in input after header detection/aliasing: {missing_base}\n"
#             f"Found columns: {list(df_base.columns)}"
#         )

#     month, year = find_month_year_in_sheet(input_path, sheet_used)
#     print("Detected Month/Year from Period:", month, year)

#     df_map = load_mapping_from_gsheet()

#     required_map_cols = ["Party Name", "Brand", "Party Name1"]
#     missing_map = [c for c in required_map_cols if c not in df_map.columns]
#     if missing_map:
#         raise ValueError(
#             f"Missing required columns in Google mapping tab: {missing_map}\n"
#             f"Found: {list(df_map.columns)}"
#         )

#     df_map_clean = (
#         df_map[required_map_cols]
#         .dropna(subset=["Party Name", "Brand"])
#         .drop_duplicates(subset=["Party Name", "Brand"], keep="first")
#         .copy()
#         .rename(columns={"Party Name1": "Party Name1_map"})
#     )

#     df_base = df_base.merge(
#         df_map_clean,
#         on=["Party Name", "Brand"],
#         how="left",
#         validate="m:1",
#     )

#     if "Party Name1" not in df_base.columns:
#         df_base["Party Name1"] = np.nan

#     df_base["Party Name1"] = (
#         df_base["Party Name1_map"]
#         .fillna(df_base["Party Name1"])
#         .fillna(df_base["Party Name"])
#     )
#     df_base.drop(columns=["Party Name1_map"], inplace=True)

#     df_base = add_prj_prefix_for_puma(df_base)
#     df_base = add_calculated_columns(df_base)

#     df_pivot = create_pay_stk_pivot(df_base, month, year)

#     with pd.ExcelWriter(str(output_file), engine="openpyxl") as writer:
#         df_base.to_excel(writer, sheet_name=BASE_SHEET_NAME_OUT, index=False)
#         df_pivot.to_excel(writer, sheet_name=PIVOT_SHEET_NAME_OUT, index=False)

#     print(f"Done. Output written to: {output_file}")


# if __name__ == "__main__":
#     main()

import os
import json
import base64
import re
import argparse
from pathlib import Path

import pandas as pd
import numpy as np
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials


# ===================== OUTPUT SHEET NAMES =====================
BASE_SHEET_NAME_OUT = "Pay_Stk_Base_Data"
PIVOT_SHEET_NAME_OUT = "Pay_Stk_Final2"

# How many top rows to scan for header row
HEADER_SCAN_ROWS = 60

# How many rows/cols to scan for the "Period" text
PERIOD_SCAN_ROWS = 20
PERIOD_SCAN_COLS = 10

# ===================== COLUMN ALIASES (FIX Qty., etc.) =====================
COLUMN_ALIASES = {
    "Qty.": "Qty",
    "QTY.": "Qty",
    "QTY": "Qty",
    "Quantity": "Qty",
    "quantity": "Qty",
    "Party name": "Party Name",
    "Party name1": "Party Name1",
}


# ===================== CLI ARGS =====================
def parse_args():
    ap = argparse.ArgumentParser(description="Payable Stock processing")
    ap.add_argument("--input", required=True, help="Path to input Excel file (.xlsx/.xls)")
    ap.add_argument("--output_dir", required=True, help="Directory to write outputs")
    ap.add_argument(
        "--output_name",
        default="Data_Processed_STK.xlsx",
        help="Output Excel file name (default: Data_Processed_STK.xlsx)",
    )
    return ap.parse_args()


def apply_column_aliases(df: pd.DataFrame) -> pd.DataFrame:
    """Rename known variant headers to canonical names (e.g., Qty. -> Qty)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns=lambda c: COLUMN_ALIASES.get(c, c), inplace=True)
    return df


def ensure_party_name_fallback(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure 'Party Name' exists and is filled:
      - If Party Name missing but Party Name1 exists -> create Party Name from Party Name1
      - If both exist -> fill blank Party Name using Party Name1
    """
    df = df.copy()

    if "Party Name" not in df.columns and "Party Name1" in df.columns:
        df["Party Name"] = df["Party Name1"]

    if "Party Name" in df.columns and "Party Name1" in df.columns:
        party = df["Party Name"].astype(str).replace("nan", "").str.strip()
        party1 = df["Party Name1"].astype(str).replace("nan", "").str.strip()
        df["Party Name"] = np.where(party.eq("") | party.isna(), party1, df["Party Name"])

    return df


# ===================== HELPERS =====================
def _get_env_str(key: str) -> str:
    val = os.getenv(key)
    if val is None or str(val).strip() == "":
        raise ValueError(f"Missing env var: {key}")
    return val.strip()


def normalize_colname(x: str) -> str:
    """Normalize headers: lowercase, strip, collapse spaces, apply alias mapping."""
    if x is None:
        return ""
    s = str(x).strip()
    s = COLUMN_ALIASES.get(s, s)  # apply alias before normalizing
    s = s.lower()
    s = " ".join(s.split())
    return s


def find_header_row_in_sheet(excel_path: Path, sheet_name: str | int, required_cols: list[str]) -> int:
    """
    Reads the first N rows with header=None and finds which row looks like the header
    by checking if it contains the required column names (normalized).
    Returns the header row index (0-based).
    """
    preview = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=HEADER_SCAN_ROWS)

    required_norm = {normalize_colname(c) for c in required_cols}

    best_row = None
    best_score = -1

    for r in range(len(preview)):
        row_vals = preview.iloc[r].tolist()
        row_norm = {normalize_colname(v) for v in row_vals if normalize_colname(v)}

        score = len(required_norm.intersection(row_norm))
        if score > best_score:
            best_score = score
            best_row = r

        if score == len(required_norm):
            return r

    min_needed = max(3, int(0.7 * len(required_norm)))  # 70% match
    if best_score >= min_needed and best_row is not None:
        return best_row

    raise ValueError(
        f"Could not detect header row in sheet '{sheet_name}'. "
        f"Best match row index={best_row} with score={best_score}/{len(required_norm)}. "
        f"Try increasing HEADER_SCAN_ROWS or check column names."
    )


def read_sheet_with_auto_header(excel_path: Path, required_cols: list[str]):
    """
    Picks the first sheet, finds header row automatically, reads full data from there.
    Returns (df, sheet_name_used).
    """
    xls = pd.ExcelFile(excel_path)
    sheet_to_read = xls.sheet_names[0]  # default: first sheet
    print("Sheets found in input file:", xls.sheet_names)
    print("Using sheet:", sheet_to_read)

    header_row = find_header_row_in_sheet(excel_path, sheet_to_read, required_cols)
    print(f"Detected header row (0-based): {header_row}  -> Excel row: {header_row + 1}")

    df = pd.read_excel(excel_path, sheet_name=sheet_to_read, header=header_row)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    # ✅ Fix column variants like Qty. and Party Name fallback
    df = apply_column_aliases(df)
    df = ensure_party_name_fallback(df)

    return df, sheet_to_read


# ===================== PERIOD (Month/Year) DETECTION =====================
def extract_month_year_from_period_text(text: str):
    """
    Extract month/year from strings like:
      "Period: 002. DEC-25 To 002. DEC-25"
      "Period: DEC-25"
      "Period: DEC 2025"
    Returns (month_str, year_int) or (None, None)
    """
    if not text:
        return None, None

    s = str(text).upper()

    m = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b\s*[-/ ]\s*(\d{2,4})", s)
    if not m:
        m = re.search(r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\s*[-/ ]\s*(\d{2,4})", s)

    if not m:
        return None, None

    mon = m.group(1)
    yy = m.group(2)

    if mon == "SEPT":
        mon = "SEP"

    year = int(yy)
    if len(yy) == 2:
        year = 2000 + year  # assumes 20xx

    return mon, year


def find_month_year_in_sheet(excel_path: Path, sheet_name: str | int):
    """
    Scans the top-left area of the sheet to find a cell containing 'PERIOD'
    (or any cell with DEC-25 like pattern) and returns (month, year).
    """
    preview = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=PERIOD_SCAN_ROWS)
    preview = preview.iloc[:, :PERIOD_SCAN_COLS]

    # 1) Prefer cells containing 'PERIOD'
    for r in range(preview.shape[0]):
        for c in range(preview.shape[1]):
            val = preview.iat[r, c]
            if pd.isna(val):
                continue
            txt = str(val)
            if "PERIOD" in txt.upper():
                mon, year = extract_month_year_from_period_text(txt)
                if mon and year:
                    return mon, year

    # 2) Fallback: any cell with month-year pattern
    for r in range(preview.shape[0]):
        for c in range(preview.shape[1]):
            val = preview.iat[r, c]
            if pd.isna(val):
                continue
            mon, year = extract_month_year_from_period_text(str(val))
            if mon and year:
                return mon, year

    return None, None


# ===================== GOOGLE SHEETS MAPPING =====================
def load_mapping_from_gsheet() -> pd.DataFrame:
    """
    Reads the mapping table from Google Sheet using a Service Account JSON stored as base64 in .env.
    Expects columns: Party Name, Brand, Party Name1
    """
    load_dotenv()

    sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")
    sheet_id = _get_env_str("GOOGLE_SHEET_ID")
    tab_name = _get_env_str("GOOGLE_SHEET_TAB_PAYSTK").strip().strip('"').strip("'")

    try:
        sa_json_text = base64.b64decode(sa_b64).decode("utf-8")
        sa_info = json.loads(sa_json_text)
    except Exception as e:
        raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(creds)

    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(tab_name)

    values = ws.get_all_values()
    if not values or len(values) < 2:
        raise ValueError(f"Google sheet tab '{tab_name}' seems empty or has no data rows.")

    header = [h.strip() for h in values[0]]
    rows = values[1:]
    df_map = pd.DataFrame(rows, columns=header)

    for c in ["Party Name", "Brand", "Party Name1"]:
        if c in df_map.columns:
            df_map[c] = df_map[c].astype(str).str.strip()

    return df_map


# ===================== GOOGLE SHEETS OUTPUT (PAY_STK) =====================
def get_gspread_client():
    """
    Creates a gspread client using GOOGLE_SA_JSON_B64 from .env.
    Needs edit access because we will update PAY_STK tab.
    """
    load_dotenv()
    sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")

    try:
        sa_json_text = base64.b64decode(sa_b64).decode("utf-8")
        sa_info = json.loads(sa_json_text)
    except Exception as e:
        raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def _coerce_sheet_df_to_expected_types(df: pd.DataFrame) -> pd.DataFrame:
    """
    Existing PAY_STK data comes as strings. Normalize Month/Year to compare reliably.
    """
    df = df.copy()
    if "Month" in df.columns:
        df["Month"] = df["Month"].astype(str).str.strip().str.upper()
    if "Year" in df.columns:
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
    return df


def upsert_pivot_to_pay_stk_tab(df_pivot: pd.DataFrame, month: str | None, year: int | None):
    """
    Writes ONLY the pivot (Pay_Stk_Final2) into Google Sheet tab PAY_STK.

    Rules:
      - If same Month+Year already exists in PAY_STK -> remove those rows, then append new rows.
      - Else -> append new rows at bottom.
      - Keeps header in row 1.
    """
    if month is None or year is None:
        raise ValueError("Month/Year could not be detected, cannot upsert into PAY_STK tab.")

    load_dotenv()
    out_sheet_id = _get_env_str("PAY_STK_SHEET_ID")
    out_tab_name = _get_env_str("PAY_STK_TAB_NAME").strip().strip('"').strip("'")

    # Normalize outgoing df
    df_out = df_pivot.copy()
    df_out["Month"] = df_out["Month"].astype(str).str.strip().str.upper()
    df_out["Year"] = pd.to_numeric(df_out["Year"], errors="coerce").astype("Int64")

    gc = get_gspread_client()
    sh = gc.open_by_key(out_sheet_id)
    ws = sh.worksheet(out_tab_name)

    values = ws.get_all_values()

    # If empty sheet: write header + data
    if not values:
        payload = [df_out.columns.tolist()] + df_out.replace({np.nan: ""}).values.tolist()
        ws.update(payload)
        print(f"[PAY_STK] Sheet was empty. Written {len(df_out)} rows.")
        return

    # Build existing df
    header = [h.strip() for h in values[0]]
    data_rows = values[1:]

    if not header or "Month" not in header or "Year" not in header:
        # If header is missing or wrong, reset sheet with correct header
        payload = [df_out.columns.tolist()] + df_out.replace({np.nan: ""}).values.tolist()
        ws.clear()
        ws.update(payload)
        print(f"[PAY_STK] Header missing/invalid. Rebuilt sheet with {len(df_out)} rows.")
        return

    df_existing = pd.DataFrame(data_rows, columns=header)
    df_existing = _coerce_sheet_df_to_expected_types(df_existing)

    # Remove existing rows for same Month+Year
    df_filtered = df_existing[
        ~((df_existing["Month"] == str(month).upper()) & (df_existing["Year"] == int(year)))
    ].copy()

    # Ensure output columns match pivot columns (if sheet has extra cols, keep them; else align)
    # We'll rebuild from the pivot columns only (clean and predictable).
    final_df = pd.concat([df_filtered[df_out.columns] if set(df_out.columns).issubset(df_filtered.columns) else df_filtered.reindex(columns=df_out.columns),
                          df_out],
                         ignore_index=True)

    # Write back all (clear + update)
    payload = [df_out.columns.tolist()] + final_df.replace({np.nan: ""}).values.tolist()
    ws.clear()
    ws.update(payload)

    print(f"[PAY_STK] Upsert complete for {month}-{year}. Now total rows: {len(final_df)}")


# ===================== YOUR ORIGINAL LOGIC =====================
def add_prj_prefix_for_puma(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ["Branch Name", "Party Name", "Party Name1"]:
        if c not in df.columns:
            df[c] = np.nan

    branch = df["Branch Name"].fillna("").astype(str).str.strip()
    party = df["Party Name"].fillna("").astype(str)
    party1 = df["Party Name1"].fillna("").astype(str)

    mask = (
        branch.str.upper().str.startswith("PRJ")
        & (
            party.str.upper().str.contains("PUMA", na=False)
            | party1.str.upper().str.contains("PUMA", na=False)
        )
    )

    def prefix_prj(series: pd.Series) -> pd.Series:
        s = series.fillna("").astype(str)
        already = s.str.upper().str.startswith("PRJ-")
        return np.where(already, s, "PRJ-" + s)

    df.loc[mask, "Party Name"] = prefix_prj(df.loc[mask, "Party Name"])
    df.loc[mask, "Party Name1"] = prefix_prj(df.loc[mask, "Party Name1"])
    return df


def add_calculated_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Cost Rate"] = pd.to_numeric(df["Cost Rate"], errors="coerce")
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")

    df["Total Cost"] = df["Cost Rate"] * df["Qty"]
    df["Gst%"] = np.where(df["Cost Rate"] <= 2500, 0.05, 0.18)
    df["Gst"] = df["Total Cost"] * df["Gst%"]
    df["Actual Cost"] = df["Total Cost"] + df["Gst"]
    return df


def create_pay_stk_pivot(df: pd.DataFrame, month: str | None, year: int | None) -> pd.DataFrame:
    df = df.copy()
    for col in ["Qty", "Total Cost", "Gst", "Actual Cost"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    pivot = pd.pivot_table(
        df,
        index=["SOR/ Outright", "Brand/Inhouse", "Party Name1"],
        values=["Qty", "Total Cost", "Gst", "Actual Cost"],
        aggfunc="sum",
        fill_value=0,
        margins=False,
    ).reset_index()

    pivot.insert(0, "Year", year)
    pivot.insert(0, "Month", month)

    return pivot


def main():
    args = parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    output_file = out_dir / args.output_name

    # ✅ Party Name can be missing; we'll fallback from Party Name1 after reading
    required_base_cols = [
        "Brand",
        "Cost Rate",
        "Qty",  # Qty. will be normalized to Qty
        "SOR/ Outright",
        "Brand/Inhouse",
        "Branch Name",
    ]

    df_base, sheet_used = read_sheet_with_auto_header(input_path, required_base_cols)

    # ✅ Must have at least one
    if "Party Name" not in df_base.columns and "Party Name1" not in df_base.columns:
        raise ValueError("Missing required column: need either 'Party Name' or 'Party Name1' in input.")

    # ✅ Ensure Party Name exists for mapping merge
    df_base = ensure_party_name_fallback(df_base)

    missing_base = [c for c in required_base_cols if c not in df_base.columns]
    if missing_base:
        raise ValueError(
            f"Missing required columns in input after header detection/aliasing: {missing_base}\n"
            f"Found columns: {list(df_base.columns)}"
        )

    month, year = find_month_year_in_sheet(input_path, sheet_used)
    print("Detected Month/Year from Period:", month, year)

    df_map = load_mapping_from_gsheet()

    required_map_cols = ["Party Name", "Brand", "Party Name1"]
    missing_map = [c for c in required_map_cols if c not in df_map.columns]
    if missing_map:
        raise ValueError(
            f"Missing required columns in Google mapping tab: {missing_map}\n"
            f"Found: {list(df_map.columns)}"
        )

    df_map_clean = (
        df_map[required_map_cols]
        .dropna(subset=["Party Name", "Brand"])
        .drop_duplicates(subset=["Party Name", "Brand"], keep="first")
        .copy()
        .rename(columns={"Party Name1": "Party Name1_map"})
    )

    df_base = df_base.merge(
        df_map_clean,
        on=["Party Name", "Brand"],
        how="left",
        validate="m:1",
    )

    if "Party Name1" not in df_base.columns:
        df_base["Party Name1"] = np.nan

    df_base["Party Name1"] = (
        df_base["Party Name1_map"]
        .fillna(df_base["Party Name1"])
        .fillna(df_base["Party Name"])
    )
    df_base.drop(columns=["Party Name1_map"], inplace=True)

    df_base = add_prj_prefix_for_puma(df_base)
    df_base = add_calculated_columns(df_base)

    df_pivot = create_pay_stk_pivot(df_base, month, year)

    # ✅ Upsert ONLY pivot output to PAY_STK tab (append/replace by Month+Year)
    upsert_pivot_to_pay_stk_tab(df_pivot, month, year)

    # ✅ Write local Excel output as before
    with pd.ExcelWriter(str(output_file), engine="openpyxl") as writer:
        df_base.to_excel(writer, sheet_name=BASE_SHEET_NAME_OUT, index=False)
        df_pivot.to_excel(writer, sheet_name=PIVOT_SHEET_NAME_OUT, index=False)

    print(f"Done. Output written to: {output_file}")


if __name__ == "__main__":
    main()
