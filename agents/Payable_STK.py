
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


# # ===================== ENV LOADING =====================
# def _load_env_once():
#     env_path_override = os.getenv("ENV_PATH", "").strip()
#     if env_path_override:
#         p = Path(env_path_override).expanduser().resolve()
#         load_dotenv(dotenv_path=p, override=True)
#         return

#     here = Path(__file__).resolve().parent
#     server_env = here / "server" / ".env"
#     if server_env.exists():
#         load_dotenv(dotenv_path=server_env, override=True)
#         return

#     load_dotenv(override=True)


# _load_env_once()


# # ===================== OUTPUT SHEET NAMES =====================
# BASE_SHEET_NAME_OUT  = "Pay_Stk_Base_Data"
# PIVOT_SHEET_NAME_OUT = "Pay_Stk_Final2"

# # How many top rows to scan for header row
# HEADER_SCAN_ROWS = 60

# # How many rows/cols to scan for the "Period" text
# PERIOD_SCAN_ROWS = 20
# PERIOD_SCAN_COLS = 10


# # ===================== COLUMN ALIASES =====================
# COLUMN_ALIASES = {
#     "Qty.":           "Qty",
#     "QTY.":           "Qty",
#     "QTY":            "Qty",
#     "Quantity":       "Qty",
#     "quantity":       "Qty",
#     "Party name":     "Party Name",
#     "Party name1":    "Party Name1",
#     "Hsn":            "HSN",
#     "Hsn Code":       "HSN",
#     "HSN Code":       "HSN",
#     "Brand/Inhouse ": "Brand/Inhouse",   # trailing space variant
#     "SOR/ Outright ": "SOR/ Outright",   # trailing space variant
# }


# # ===================== CLI ARGS =====================
# def parse_args():
#     ap = argparse.ArgumentParser(description="Payable Stock processing")
#     ap.add_argument("--input",      required=True, help="Path to input Excel file (.xlsx/.xls)")
#     ap.add_argument("--output_dir", required=True, help="Directory to write outputs")
#     ap.add_argument(
#         "--output_name",
#         default="Payable_STK.xlsx",
#         help="Output Excel file name (default: Payable_STK.xlsx)",
#     )
#     return ap.parse_args()


# # ===================== COLUMN HELPERS =====================
# def apply_column_aliases(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     df.columns = [str(c).strip() for c in df.columns]
#     df.rename(columns=lambda c: COLUMN_ALIASES.get(c, c), inplace=True)
#     return df


# def ensure_party_name_fallback(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()

#     if "Party Name" not in df.columns and "Party Name1" in df.columns:
#         df["Party Name"] = df["Party Name1"]

#     if "Party Name" in df.columns and "Party Name1" in df.columns:
#         party  = df["Party Name"].astype(str).replace("nan", "").str.strip()
#         party1 = df["Party Name1"].astype(str).replace("nan", "").str.strip()
#         df["Party Name"] = np.where(party.eq("") | party.isna(), party1, df["Party Name"])

#     return df


# # ===================== HELPERS =====================
# def _get_env_str(key: str) -> str:
#     val = os.getenv(key)
#     if val is None or str(val).strip() == "":
#         raise ValueError(f"Missing env var: {key}")
#     return val.strip().strip('"').strip("'")


# def normalize_colname(x: str) -> str:
#     if x is None:
#         return ""
#     s = str(x).strip()
#     s = COLUMN_ALIASES.get(s, s)   # apply alias before normalizing
#     s = s.lower()
#     s = " ".join(s.split())
#     return s


# def normalize_hsn(val) -> str:
#     if pd.isna(val):
#         return ""
#     s = str(val).strip()
#     if s.endswith(".0"):
#         s = s[:-2]
#     return s


# def parse_percent_to_decimal(val) -> float:
#     if pd.isna(val):
#         return np.nan
#     s = str(val).strip()
#     if s == "":
#         return np.nan
#     try:
#         if s.endswith("%"):
#             return float(s[:-1].strip()) / 100.0
#         num = float(s)
#         return num / 100.0 if num > 1 else num
#     except Exception:
#         return np.nan


# # ===================== HEADER DETECTION =====================
# def find_header_row_in_sheet(excel_path: Path, sheet_name, required_cols: list) -> int:
#     preview = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=HEADER_SCAN_ROWS)

#     required_norm = {normalize_colname(c) for c in required_cols}

#     best_row  = None
#     best_score = -1

#     for r in range(len(preview)):
#         row_vals = preview.iloc[r].tolist()
#         row_norm = {normalize_colname(v) for v in row_vals if normalize_colname(v)}

#         score = len(required_norm.intersection(row_norm))
#         if score > best_score:
#             best_score = score
#             best_row   = r

#         if score == len(required_norm):
#             return r

#     min_needed = max(3, int(0.7 * len(required_norm)))
#     if best_score >= min_needed and best_row is not None:
#         return best_row

#     raise ValueError(
#         f"Could not detect header row in sheet '{sheet_name}'. "
#         f"Best match row index={best_row} with score={best_score}/{len(required_norm)}. "
#         f"Try increasing HEADER_SCAN_ROWS or check column names."
#     )


# def read_sheet_with_auto_header(excel_path: Path, required_cols: list):
#     xls           = pd.ExcelFile(excel_path)
#     sheet_to_read = xls.sheet_names[0]
#     print("Sheets found in input file:", xls.sheet_names)
#     print("Using sheet:", sheet_to_read)

#     header_row = find_header_row_in_sheet(excel_path, sheet_to_read, required_cols)
#     print(f"Detected header row (0-based): {header_row}  -> Excel row: {header_row + 1}")

#     df = pd.read_excel(excel_path, sheet_name=sheet_to_read, header=header_row)
#     df = df.dropna(how="all").copy()
#     df.columns = [str(c).strip() for c in df.columns]

#     df = apply_column_aliases(df)
#     df = ensure_party_name_fallback(df)

#     return df, sheet_to_read


# # ===================== PERIOD (Month/Year) DETECTION =====================
# def extract_month_year_from_period_text(text: str):
#     if not text:
#         return None, None

#     s = str(text).upper()

#     m = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b\s*[-/ ]\s*(\d{2,4})", s)
#     if not m:
#         m = re.search(r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\s*[-/ ]\s*(\d{2,4})", s)

#     if not m:
#         return None, None

#     mon = m.group(1)
#     yy  = m.group(2)

#     if mon == "SEPT":
#         mon = "SEP"

#     year = int(yy)
#     if len(yy) == 2:
#         year = 2000 + year

#     return mon, year


# def find_month_year_in_sheet(excel_path: Path, sheet_name):
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


# # ===================== GOOGLE SHEETS — MAPPING =====================
# def get_gspread_client_readonly():
#     sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")
#     try:
#         sa_info = json.loads(base64.b64decode(sa_b64).decode("utf-8"))
#     except Exception as e:
#         raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e
#     scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
#     creds  = Credentials.from_service_account_info(sa_info, scopes=scopes)
#     return gspread.authorize(creds)


# def get_gspread_client_write():
#     sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")
#     try:
#         sa_info = json.loads(base64.b64decode(sa_b64).decode("utf-8"))
#     except Exception as e:
#         raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e
#     scopes = ["https://www.googleapis.com/auth/spreadsheets"]
#     creds  = Credentials.from_service_account_info(sa_info, scopes=scopes)
#     return gspread.authorize(creds)


# def load_mapping_from_gsheet() -> pd.DataFrame:
#     """
#     Reads mapping tab from Google Sheet.
#     Expects columns: Party Name, Brand, Party Name1
#     Uses GOOGLE_SHEET_ID_PAY_STK + GOOGLE_SHEET_TAB_PAYSTK env vars.
#     """
#     gc = get_gspread_client_readonly()

#     sheet_id = _get_env_str("GOOGLE_SHEET_ID_PAY_STK")
#     tab_name = _get_env_str("GOOGLE_SHEET_TAB_PAYSTK")

#     sh = gc.open_by_key(sheet_id)
#     ws = sh.worksheet(tab_name)

#     values = ws.get_all_values()
#     if not values or len(values) < 2:
#         raise ValueError(f"Google sheet tab '{tab_name}' seems empty or has no data rows.")

#     header    = [h.strip() for h in values[0]]
#     rows      = values[1:]
#     df_map    = pd.DataFrame(rows, columns=header)

#     for c in ["Party Name", "Brand", "Party Name1"]:
#         if c in df_map.columns:
#             df_map[c] = df_map[c].astype(str).str.strip()

#     return df_map


# def load_hsn_from_gsheet() -> pd.DataFrame:
#     """
#     Reads HSN tab from Google Sheet.
#     Expected columns: Hsn Code (or HSN), Gst %
#     Uses GOOGLE_SHEET_ID_PAY_STK + GOOGLE_SHEET_TAB_HSN env vars.
#     """
#     gc = get_gspread_client_readonly()

#     sheet_id = _get_env_str("GOOGLE_SHEET_ID_PAY_STK")
#     tab_name = _get_env_str("GOOGLE_SHEET_TAB_HSN")

#     sh = gc.open_by_key(sheet_id)
#     ws = sh.worksheet(tab_name)

#     values = ws.get_all_values()
#     if not values or len(values) < 2:
#         raise ValueError(f"Google sheet tab '{tab_name}' seems empty or has no data rows.")

#     header    = [str(h).strip() for h in values[0]]
#     rows      = values[1:]
#     df_hsn    = pd.DataFrame(rows, columns=header)
#     df_hsn.columns = [str(c).strip() for c in df_hsn.columns]

#     hsn_col = None
#     gst_col = None

#     for c in df_hsn.columns:
#         n = normalize_colname(c)
#         if n in {"hsn code", "hsn", "hsn sap code"} and hsn_col is None:
#             hsn_col = c
#         if n in {"gst %", "gst%", "gst"} and gst_col is None:
#             gst_col = c

#     if hsn_col is None or gst_col is None:
#         raise ValueError(
#             f"HSN tab '{tab_name}' must contain Hsn Code and Gst % columns. "
#             f"Found columns: {list(df_hsn.columns)}"
#         )

#     out = df_hsn[[hsn_col, gst_col]].copy()
#     out.columns = ["HSN", "GST_FROM_HSN"]
#     out["HSN"]          = out["HSN"].map(normalize_hsn)
#     out["GST_FROM_HSN"] = out["GST_FROM_HSN"].map(parse_percent_to_decimal)

#     out = (
#         out[(out["HSN"] != "") & (~out["GST_FROM_HSN"].isna())]
#         .drop_duplicates(subset=["HSN"], keep="first")
#     )
#     return out


# # ===================== GOOGLE SHEETS — OUTPUT (PAY_STK UPSERT) =====================
# def _coerce_sheet_df_to_expected_types(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     if "Month" in df.columns:
#         df["Month"] = df["Month"].astype(str).str.strip().str.upper()
#     if "Year" in df.columns:
#         df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
#     return df


# def upsert_pivot_to_pay_stk_tab(df_pivot: pd.DataFrame, month, year):
#     """
#     Writes ONLY the pivot (Pay_Stk_Final2) into Google Sheet tab PAY_STK.

#     Rules:
#       - If same Month+Year already exists -> remove those rows, then append new rows.
#       - Else -> append new rows at bottom.
#       - Keeps header in row 1.
#     """
#     if month is None or year is None:
#         raise ValueError("Month/Year could not be detected, cannot upsert into PAY_STK tab.")

#     out_sheet_id = _get_env_str("PAY_STK_SHEET_ID")
#     out_tab_name = _get_env_str("PAY_STK_TAB_NAME")

#     # Normalize outgoing df
#     df_out          = df_pivot.copy()
#     df_out["Month"] = df_out["Month"].astype(str).str.strip().str.upper()
#     df_out["Year"]  = pd.to_numeric(df_out["Year"], errors="coerce").astype("Int64")

#     gc = get_gspread_client_write()
#     sh = gc.open_by_key(out_sheet_id)
#     ws = sh.worksheet(out_tab_name)

#     values = ws.get_all_values()

#     # Empty sheet: write header + data
#     if not values:
#         payload = [df_out.columns.tolist()] + df_out.replace({np.nan: ""}).values.tolist()
#         ws.update(payload)
#         print(f"[PAY_STK] Sheet was empty. Written {len(df_out)} rows.")
#         return

#     header    = [h.strip() for h in values[0]]
#     data_rows = values[1:]

#     if not header or "Month" not in header or "Year" not in header:
#         ws.clear()
#         payload = [df_out.columns.tolist()] + df_out.replace({np.nan: ""}).values.tolist()
#         ws.update(payload)
#         print(f"[PAY_STK] Header missing/invalid. Rebuilt sheet with {len(df_out)} rows.")
#         return

#     df_existing = pd.DataFrame(data_rows, columns=header)
#     df_existing = _coerce_sheet_df_to_expected_types(df_existing)

#     # Remove existing rows for same Month+Year
#     df_filtered = df_existing[
#         ~((df_existing["Month"] == str(month).upper()) & (df_existing["Year"] == int(year)))
#     ].copy()

#     # Align columns
#     if set(df_out.columns).issubset(df_filtered.columns):
#         df_filtered = df_filtered[df_out.columns]
#     else:
#         df_filtered = df_filtered.reindex(columns=df_out.columns)

#     final_df = pd.concat([df_filtered, df_out], ignore_index=True)

#     payload = [df_out.columns.tolist()] + final_df.replace({np.nan: ""}).values.tolist()
#     ws.clear()
#     ws.update(payload)

#     print(f"[PAY_STK] Upsert complete for {month}-{year}. Now total rows: {len(final_df)}")


# # ===================== BUSINESS LOGIC =====================
# def add_prj_prefix_for_puma(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     for c in ["Branch Name", "Party Name", "Party Name1"]:
#         if c not in df.columns:
#             df[c] = np.nan

#     branch = df["Branch Name"].fillna("").astype(str).str.strip()
#     party  = df["Party Name"].fillna("").astype(str)
#     party1 = df["Party Name1"].fillna("").astype(str)

#     mask = (
#         branch.str.upper().str.startswith("PRJ")
#         & (
#             party.str.upper().str.contains("PUMA", na=False)
#             | party1.str.upper().str.contains("PUMA", na=False)
#         )
#     )

#     def prefix_prj(series: pd.Series) -> pd.Series:
#         s       = series.fillna("").astype(str)
#         already = s.str.upper().str.startswith("PRJ-")
#         return np.where(already, s, "PRJ-" + s)

#     df.loc[mask, "Party Name"]  = prefix_prj(df.loc[mask, "Party Name"])
#     df.loc[mask, "Party Name1"] = prefix_prj(df.loc[mask, "Party Name1"])
#     return df


# def add_calculated_columns(df: pd.DataFrame, df_hsn: pd.DataFrame) -> pd.DataFrame:
#     """
#     Calculates Total Cost, Gst%, Gst, Actual Cost.

#     GST logic:
#       - For FOOTWEAR ACCESSORIES division: look up GST% from HSN master sheet
#       - For all others: use Cost Rate threshold (<=2500 -> 5%, >2500 -> 18%)
#     """
#     df = df.copy()

#     df["Cost Rate"] = pd.to_numeric(df["Cost Rate"], errors="coerce")
#     df["Qty"]       = pd.to_numeric(df["Qty"], errors="coerce")

#     if "Division" not in df.columns:
#         raise ValueError("Missing required column in input: Division")
#     if "HSN" not in df.columns:
#         raise ValueError("Missing required column in input: HSN")

#     df["Division"] = df["Division"].astype(str).str.strip()
#     df["HSN"]      = df["HSN"].map(normalize_hsn)

#     df["Total Cost"] = df["Cost Rate"] * df["Qty"]

#     # Merge HSN GST rates
#     df = df.merge(df_hsn, on="HSN", how="left")

#     # Default GST by cost threshold
#     default_gst = np.where(df["Cost Rate"] <= 2500, 0.05, 0.18)

#     # Footwear Accessories: use HSN lookup if available, else fall back to default
#     footwear_acc_mask = df["Division"].str.upper().eq("FOOTWEAR ACCESSORIES")

#     df["Gst%"] = np.where(
#         footwear_acc_mask & df["GST_FROM_HSN"].notna(),
#         df["GST_FROM_HSN"],
#         default_gst,
#     )

#     df["Gst"]         = df["Total Cost"] * df["Gst%"]
#     df["Actual Cost"] = df["Total Cost"] + df["Gst"]

#     return df


# def create_pay_stk_pivot(df: pd.DataFrame, month, year) -> pd.DataFrame:
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

#     pivot.insert(0, "Year",  year)
#     pivot.insert(0, "Month", month)

#     return pivot


# # ===================== MAIN =====================
# def main():
#     args = parse_args()

#     input_path = Path(args.input)
#     if not input_path.exists():
#         raise FileNotFoundError(f"Input file not found: {input_path}")

#     out_dir = Path(args.output_dir)
#     out_dir.mkdir(parents=True, exist_ok=True)
#     output_file = out_dir / args.output_name

#     required_base_cols = [
#         "Brand",
#         "Cost Rate",
#         "Qty",           # Qty. will be normalized to Qty via aliases
#         "SOR/ Outright",
#         "Brand/Inhouse",
#         "Branch Name",
#         "Division",
#         "HSN",
#     ]

#     df_base, sheet_used = read_sheet_with_auto_header(input_path, required_base_cols)

#     if "Party Name" not in df_base.columns and "Party Name1" not in df_base.columns:
#         raise ValueError("Missing required column: need either 'Party Name' or 'Party Name1' in input.")

#     df_base = ensure_party_name_fallback(df_base)

#     missing_base = [c for c in required_base_cols if c not in df_base.columns]
#     if missing_base:
#         raise ValueError(
#             f"Missing required columns in input after header detection/aliasing: {missing_base}\n"
#             f"Found columns: {list(df_base.columns)}"
#         )

#     month, year = find_month_year_in_sheet(input_path, sheet_used)
#     print("Detected Month/Year from Period:", month, year)

#     # Load masters from Google Sheets
#     df_map = load_mapping_from_gsheet()
#     df_hsn = load_hsn_from_gsheet()

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

#     # Business logic
#     df_base  = add_prj_prefix_for_puma(df_base)
#     df_base  = add_calculated_columns(df_base, df_hsn)
#     df_pivot = create_pay_stk_pivot(df_base, month, year)

#     # Upsert pivot to PAY_STK Google Sheet tab
#     upsert_pivot_to_pay_stk_tab(df_pivot, month, year)

#     # Write local Excel output
#     with pd.ExcelWriter(str(output_file), engine="openpyxl") as writer:
#         df_base.to_excel(writer,  sheet_name=BASE_SHEET_NAME_OUT,  index=False)
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


# ===================== ENV LOADING =====================
def _load_env_once():
    env_path_override = os.getenv("ENV_PATH", "").strip()
    if env_path_override:
        p = Path(env_path_override).expanduser().resolve()
        load_dotenv(dotenv_path=p, override=True)
        return

    here = Path(__file__).resolve().parent
    server_env = here / "server" / ".env"
    if server_env.exists():
        load_dotenv(dotenv_path=server_env, override=True)
        return

    load_dotenv(override=True)


_load_env_once()


# ===================== OUTPUT SHEET NAMES =====================
BASE_SHEET_NAME_OUT  = "Pay_Stk_Base_Data"
PIVOT_SHEET_NAME_OUT = "Pay_Stk_Final2"

# How many top rows to scan for header row
HEADER_SCAN_ROWS = 60

# How many rows/cols to scan for the "Period" text
PERIOD_SCAN_ROWS = 20
PERIOD_SCAN_COLS = 10


# ===================== COLUMN ALIASES =====================
COLUMN_ALIASES = {
    "Qty.":           "Qty",
    "QTY.":           "Qty",
    "QTY":            "Qty",
    "Quantity":       "Qty",
    "quantity":       "Qty",
    "Party name":     "Party Name",
    "Party name1":    "Party Name1",
    "Hsn":            "HSN",
    "Hsn Code":       "HSN",
    "HSN Code":       "HSN",
    "Brand/Inhouse ": "Brand/Inhouse",   # trailing space variant
    "SOR/ Outright ": "SOR/ Outright",   # trailing space variant
}


# ===================== CLI ARGS =====================
def parse_args():
    ap = argparse.ArgumentParser(description="Payable Stock processing")
    ap.add_argument("--input",      required=True, help="Path to input Excel file (.xlsx/.xls)")
    ap.add_argument("--output_dir", required=True, help="Directory to write outputs")
    ap.add_argument(
        "--output_name",
        default="Payable_STK.xlsx",
        help="Output Excel file name (default: Payable_STK.xlsx)",
    )
    return ap.parse_args()


# ===================== COLUMN HELPERS =====================
def apply_column_aliases(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns=lambda c: COLUMN_ALIASES.get(c, c), inplace=True)
    return df


def ensure_party_name_fallback(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "Party Name" not in df.columns and "Party Name1" in df.columns:
        df["Party Name"] = df["Party Name1"]

    if "Party Name" in df.columns and "Party Name1" in df.columns:
        party  = df["Party Name"].astype(str).replace("nan", "").str.strip()
        party1 = df["Party Name1"].astype(str).replace("nan", "").str.strip()
        df["Party Name"] = np.where(party.eq("") | party.isna(), party1, df["Party Name"])

    return df


# ===================== HELPERS =====================
def _get_env_str(key: str) -> str:
    val = os.getenv(key)
    if val is None or str(val).strip() == "":
        raise ValueError(f"Missing env var: {key}")
    return val.strip().strip('"').strip("'")


def normalize_colname(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    s = COLUMN_ALIASES.get(s, s)   # apply alias before normalizing
    s = s.lower()
    s = " ".join(s.split())
    return s


def normalize_hsn(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def parse_percent_to_decimal(val) -> float:
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if s == "":
        return np.nan
    try:
        if s.endswith("%"):
            return float(s[:-1].strip()) / 100.0
        num = float(s)
        return num / 100.0 if num > 1 else num
    except Exception:
        return np.nan


# ===================== HEADER DETECTION =====================
def find_header_row_in_sheet(excel_path: Path, sheet_name, required_cols: list) -> int:
    preview = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=HEADER_SCAN_ROWS)

    required_norm = {normalize_colname(c) for c in required_cols}

    best_row  = None
    best_score = -1

    for r in range(len(preview)):
        row_vals = preview.iloc[r].tolist()
        row_norm = {normalize_colname(v) for v in row_vals if normalize_colname(v)}

        score = len(required_norm.intersection(row_norm))
        if score > best_score:
            best_score = score
            best_row   = r

        if score == len(required_norm):
            return r

    min_needed = max(3, int(0.7 * len(required_norm)))
    if best_score >= min_needed and best_row is not None:
        return best_row

    raise ValueError(
        f"Could not detect header row in sheet '{sheet_name}'. "
        f"Best match row index={best_row} with score={best_score}/{len(required_norm)}. "
        f"Try increasing HEADER_SCAN_ROWS or check column names."
    )


def read_sheet_with_auto_header(excel_path: Path, required_cols: list):
    xls           = pd.ExcelFile(excel_path)
    sheet_to_read = xls.sheet_names[0]
    print("Sheets found in input file:", xls.sheet_names)
    print("Using sheet:", sheet_to_read)

    header_row = find_header_row_in_sheet(excel_path, sheet_to_read, required_cols)
    print(f"Detected header row (0-based): {header_row}  -> Excel row: {header_row + 1}")

    df = pd.read_excel(excel_path, sheet_name=sheet_to_read, header=header_row)
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]

    df = apply_column_aliases(df)
    df = ensure_party_name_fallback(df)

    return df, sheet_to_read


# ===================== PERIOD (Month/Year) DETECTION =====================
def extract_month_year_from_period_text(text: str):
    if not text:
        return None, None

    s = str(text).upper()

    m = re.search(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\b\s*[-/ ]\s*(\d{2,4})", s)
    if not m:
        m = re.search(r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)\s*[-/ ]\s*(\d{2,4})", s)

    if not m:
        return None, None

    mon = m.group(1)
    yy  = m.group(2)

    if mon == "SEPT":
        mon = "SEP"

    year = int(yy)
    if len(yy) == 2:
        year = 2000 + year

    return mon, year


def find_month_year_in_sheet(excel_path: Path, sheet_name):
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


# ===================== GOOGLE SHEETS — MAPPING =====================
def get_gspread_client_readonly():
    sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")
    try:
        sa_info = json.loads(base64.b64decode(sa_b64).decode("utf-8"))
    except Exception as e:
        raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds  = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def get_gspread_client_write():
    sa_b64 = _get_env_str("GOOGLE_SA_JSON_B64")
    try:
        sa_info = json.loads(base64.b64decode(sa_b64).decode("utf-8"))
    except Exception as e:
        raise ValueError("GOOGLE_SA_JSON_B64 is not valid base64-encoded service account JSON") from e
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def load_mapping_from_gsheet() -> pd.DataFrame:
    gc = get_gspread_client_readonly()

    sheet_id = _get_env_str("GOOGLE_SHEET_ID_PAY_STK")
    tab_name = _get_env_str("GOOGLE_SHEET_TAB_PAYSTK")

    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(tab_name)

    values = ws.get_all_values()
    if not values or len(values) < 2:
        raise ValueError(f"Google sheet tab '{tab_name}' seems empty or has no data rows.")

    header    = [h.strip() for h in values[0]]
    rows      = values[1:]
    df_map    = pd.DataFrame(rows, columns=header)

    for c in ["Party Name", "Brand", "Party Name1"]:
        if c in df_map.columns:
            df_map[c] = df_map[c].astype(str).str.strip()

    return df_map


def load_hsn_from_gsheet() -> pd.DataFrame:
    gc = get_gspread_client_readonly()

    sheet_id = _get_env_str("GOOGLE_SHEET_ID_PAY_STK")
    tab_name = _get_env_str("GOOGLE_SHEET_TAB_HSN")

    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(tab_name)

    values = ws.get_all_values()
    if not values or len(values) < 2:
        raise ValueError(f"Google sheet tab '{tab_name}' seems empty or has no data rows.")

    header    = [str(h).strip() for h in values[0]]
    rows      = values[1:]
    df_hsn    = pd.DataFrame(rows, columns=header)
    df_hsn.columns = [str(c).strip() for c in df_hsn.columns]

    hsn_col = None
    gst_col = None

    for c in df_hsn.columns:
        n = normalize_colname(c)
        if n in {"hsn code", "hsn", "hsn sap code"} and hsn_col is None:
            hsn_col = c
        if n in {"gst %", "gst%", "gst"} and gst_col is None:
            gst_col = c

    if hsn_col is None or gst_col is None:
        raise ValueError(
            f"HSN tab '{tab_name}' must contain Hsn Code and Gst % columns. "
            f"Found columns: {list(df_hsn.columns)}"
        )

    out = df_hsn[[hsn_col, gst_col]].copy()
    out.columns = ["HSN", "GST_FROM_HSN"]
    out["HSN"]          = out["HSN"].map(normalize_hsn)
    out["GST_FROM_HSN"] = out["GST_FROM_HSN"].map(parse_percent_to_decimal)

    out = (
        out[(out["HSN"] != "") & (~out["GST_FROM_HSN"].isna())]
        .drop_duplicates(subset=["HSN"], keep="first")
    )
    return out


# ===================== GOOGLE SHEETS — OUTPUT (PAY_STK UPSERT) =====================
def _coerce_sheet_df_to_expected_types(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Month" in df.columns:
        df["Month"] = df["Month"].astype(str).str.strip().str.upper()
    if "Year" in df.columns:
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")
    return df


def upsert_pivot_to_pay_stk_tab(df_pivot: pd.DataFrame, month, year):
    if month is None or year is None:
        raise ValueError("Month/Year could not be detected, cannot upsert into PAY_STK tab.")

    out_sheet_id = _get_env_str("PAY_STK_SHEET_ID")
    out_tab_name = _get_env_str("PAY_STK_TAB_NAME")

    df_out          = df_pivot.copy()
    df_out["Month"] = df_out["Month"].astype(str).str.strip().str.upper()
    df_out["Year"]  = pd.to_numeric(df_out["Year"], errors="coerce").astype("Int64")

    gc = get_gspread_client_write()
    sh = gc.open_by_key(out_sheet_id)
    ws = sh.worksheet(out_tab_name)

    values = ws.get_all_values()

    if not values:
        payload = [df_out.columns.tolist()] + df_out.replace({np.nan: ""}).values.tolist()
        ws.update(payload)
        print(f"[PAY_STK] Sheet was empty. Written {len(df_out)} rows.")
        return

    header    = [h.strip() for h in values[0]]
    data_rows = values[1:]

    if not header or "Month" not in header or "Year" not in header:
        ws.clear()
        payload = [df_out.columns.tolist()] + df_out.replace({np.nan: ""}).values.tolist()
        ws.update(payload)
        print(f"[PAY_STK] Header missing/invalid. Rebuilt sheet with {len(df_out)} rows.")
        return

    df_existing = pd.DataFrame(data_rows, columns=header)
    df_existing = _coerce_sheet_df_to_expected_types(df_existing)

    df_filtered = df_existing[
        ~((df_existing["Month"] == str(month).upper()) & (df_existing["Year"] == int(year)))
    ].copy()

    if set(df_out.columns).issubset(df_filtered.columns):
        df_filtered = df_filtered[df_out.columns]
    else:
        df_filtered = df_filtered.reindex(columns=df_out.columns)

    final_df = pd.concat([df_filtered, df_out], ignore_index=True)

    payload = [df_out.columns.tolist()] + final_df.replace({np.nan: ""}).values.tolist()
    ws.clear()
    ws.update(payload)

    print(f"[PAY_STK] Upsert complete for {month}-{year}. Now total rows: {len(final_df)}")


# ===================== BUSINESS LOGIC =====================
def add_prj_prefix_for_puma(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in ["Branch Name", "Party Name", "Party Name1"]:
        if c not in df.columns:
            df[c] = np.nan

    branch = df["Branch Name"].fillna("").astype(str).str.strip()
    party  = df["Party Name"].fillna("").astype(str)
    party1 = df["Party Name1"].fillna("").astype(str)

    mask = (
        branch.str.upper().str.startswith("PRJ")
        & (
            party.str.upper().str.contains("PUMA", na=False)
            | party1.str.upper().str.contains("PUMA", na=False)
        )
    )

    def prefix_prj(series: pd.Series) -> pd.Series:
        s       = series.fillna("").astype(str)
        already = s.str.upper().str.startswith("PRJ-")
        return np.where(already, s, "PRJ-" + s)

    df.loc[mask, "Party Name"]  = prefix_prj(df.loc[mask, "Party Name"])
    df.loc[mask, "Party Name1"] = prefix_prj(df.loc[mask, "Party Name1"])
    return df


def add_calculated_columns(df: pd.DataFrame, df_hsn: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df["Cost Rate"] = pd.to_numeric(df["Cost Rate"], errors="coerce")
    df["Qty"]       = pd.to_numeric(df["Qty"], errors="coerce")

    if "Division" not in df.columns:
        raise ValueError("Missing required column in input: Division")
    if "HSN" not in df.columns:
        raise ValueError("Missing required column in input: HSN")

    df["Division"] = df["Division"].astype(str).str.strip()
    df["HSN"]      = df["HSN"].map(normalize_hsn)

    df["Total Cost"] = df["Cost Rate"] * df["Qty"]

    df = df.merge(df_hsn, on="HSN", how="left")

    default_gst = np.where(df["Cost Rate"] <= 2500, 0.05, 0.18)

    footwear_acc_mask = df["Division"].str.upper().eq("FOOTWEAR ACCESSORIES")

    df["Gst%"] = np.where(
        footwear_acc_mask & df["GST_FROM_HSN"].notna(),
        df["GST_FROM_HSN"],
        default_gst,
    )

    df["Gst"]         = df["Total Cost"] * df["Gst%"]
    df["Actual Cost"] = df["Total Cost"] + df["Gst"]

    return df


def create_pay_stk_pivot(df: pd.DataFrame, month, year) -> pd.DataFrame:
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

    pivot.insert(0, "Year",  year)
    pivot.insert(0, "Month", month)

    return pivot


# ===================== MAIN =====================
def main():
    args = parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    output_file = out_dir / args.output_name

    required_base_cols = [
        "Brand",
        "Cost Rate",
        "Qty",
        "SOR/ Outright",
        "Brand/Inhouse",
        "Branch Name",
        "Division",
        "HSN",
    ]

    df_base, sheet_used = read_sheet_with_auto_header(input_path, required_base_cols)

    if "Party Name" not in df_base.columns and "Party Name1" not in df_base.columns:
        raise ValueError("Missing required column: need either 'Party Name' or 'Party Name1' in input.")

    df_base = ensure_party_name_fallback(df_base)

    missing_base = [c for c in required_base_cols if c not in df_base.columns]
    if missing_base:
        raise ValueError(
            f"Missing required columns in input after header detection/aliasing: {missing_base}\n"
            f"Found columns: {list(df_base.columns)}"
        )

    month, year = find_month_year_in_sheet(input_path, sheet_used)
    print("Detected Month/Year from Period:", month, year)
    # ✅ Signal data period to server.js for Drive folder naming
    if month and year:
        print(f"PERIOD: {month}-{year}")

    # Load masters from Google Sheets
    df_map = load_mapping_from_gsheet()
    df_hsn = load_hsn_from_gsheet()

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

    # Business logic
    df_base  = add_prj_prefix_for_puma(df_base)
    df_base  = add_calculated_columns(df_base, df_hsn)
    df_pivot = create_pay_stk_pivot(df_base, month, year)

    # Upsert pivot to PAY_STK Google Sheet tab
    upsert_pivot_to_pay_stk_tab(df_pivot, month, year)

    # Write local Excel output
    with pd.ExcelWriter(str(output_file), engine="openpyxl") as writer:
        df_base.to_excel(writer,  sheet_name=BASE_SHEET_NAME_OUT,  index=False)
        df_pivot.to_excel(writer, sheet_name=PIVOT_SHEET_NAME_OUT, index=False)

    print(f"Done. Output written to: {output_file}")


if __name__ == "__main__":
    main()